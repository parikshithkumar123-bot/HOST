// ── BRUTE-FORCE PROTECTION ─────────────────────────────────────────────────
const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_WINDOW_SEC   = 600;   // 10 min rolling window
const LOGIN_LOCKOUT_SEC  = 900;   // 15 min lockout

function getLoginAttemptKey_(u) { return 'login_att_' + String(u||'').toLowerCase().replace(/\s+/g,'_').slice(0,40); }
function getLockoutKey_(u)       { return 'login_lck_' + String(u||'').toLowerCase().replace(/\s+/g,'_').slice(0,40); }

function checkLoginLockout_(username) {
  if (CacheService.getScriptCache().get(getLockoutKey_(username)))
    throw new Error('Account temporarily locked after too many failed attempts. Please try again in 15 minutes.');
}

function recordFailedLogin_(username) {
  const cache = CacheService.getScriptCache();
  const key = getLoginAttemptKey_(username);
  const attempts = parseInt(cache.get(key) || '0', 10) + 1;
  cache.put(key, String(attempts), LOGIN_WINDOW_SEC);
  if (attempts >= LOGIN_MAX_ATTEMPTS) {
    cache.put(getLockoutKey_(username), '1', LOGIN_LOCKOUT_SEC);
    cache.remove(key);
  }
}

function clearLoginAttempts_(username) {
  const cache = CacheService.getScriptCache();
  cache.remove(getLoginAttemptKey_(username));
  cache.remove(getLockoutKey_(username));
}
// ───────────────────────────────────────────────────────────────────────────

const AuthService = {
  login(payload) {
    const username = String(payload.username || '').trim();
    const password = String(payload.password || '');
    if (!username || !password) {
      throw new Error('Username and password are required.');
    }

    // FIX 1: Block before any sheet read if locked out
    checkLoginLockout_(username);

    const user = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'username', username);
    if (!user || user.status !== 'ACTIVE') {
      recordFailedLogin_(username);
      logAuditEvent_('LOGIN_FAILED', 'USER', username, null, { reason: 'USER_NOT_FOUND_OR_INACTIVE' });
      throw new Error('Invalid username or password.');
    }

    const computedHash = hashPassword_(password, user.salt);
    if (computedHash !== user.password_hash) {
      recordFailedLogin_(username);
      logAuditEvent_('LOGIN_FAILED', 'USER', user.user_id, user.user_id, { reason: 'INVALID_PASSWORD' });
      throw new Error('Invalid username or password.');
    }

    // Successful — clear failed attempts
    clearLoginAttempts_(username);

    const session = createSession_(user, Boolean(payload.rememberMe));
    Repository.updateById_(
      APP_CONFIG.sheets.USERS,
      'user_id',
      user.user_id,
      { last_login_at: new Date().toISOString() }
    );
    logAuditEvent_('LOGIN_SUCCESS', 'USER', user.user_id, user.user_id, {});

    // FAST LOGIN V366
    return respond_(true, {
      session: sanitizeSession_(session)
    });
  },

  logout(token) {
    token = String(token || '').trim();
    if (!token) return respond_(true, { loggedOut: true, alreadyLoggedOut: true });

    let session = null;
    try {
      session = getSessionOrThrow_(token);
    } catch (sessionError) {
      invalidateSession_(token);
      return respond_(true, { loggedOut: true, alreadyLoggedOut: true });
    }

    invalidateSession_(session.token);
    logAuditEvent_('LOGOUT', 'SESSION', session.token, session.userId, {});
    return respond_(true, { loggedOut: true });
  },

  getSession(token) {
    const session = getSessionOrThrow_(token);
    return respond_(true, sanitizeSession_(session));
  }
};

function login(payload) {
  return AuthService.login(payload);
}

function logout(token) {
  return AuthService.logout(token);
}

function getSession(token) {
  return AuthService.getSession(token);
}

// FIX #2: Use a high-entropy secret pepper stored in Script Properties,
// never a public constant like the app name.
// One-time setup: run this in the script editor once:
//   PropertiesService.getScriptProperties().setProperty('PASSWORD_PEPPER', Utilities.getUuid());
function hashPassword_(password, salt) {
  let pepper = '';
  try {
    pepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '';
  } catch (e) {
    console.error('hashPassword_: could not read PASSWORD_PEPPER from Script Properties', e);
  }
  // FIX #BUG8: Reduced from 10,000 to 500 rounds.
  // GAS V8 processes ~3,000 SHA-256 rounds/sec; 10,000 rounds = ~3 sec per hash call.
  // Login calls hashPassword_ once and changeOwnPortalPassword calls it twice = ~6 sec,
  // which regularly hits the GAS 30s limit under concurrent load and cold starts.
  // 500 rounds (~0.17 sec) is adequate for a server-side hashed+salted+peppered password.
  // NOTE: Changing ROUNDS invalidates all existing password_hash values in the Users sheet.
  // Deploy this with a forced password-reset (must_reset_password = TRUE for all users),
  // or run a one-time migration script to re-hash all passwords with the new round count.
  const ROUNDS = 500;
  let input = salt + ':' + password + ':' + pepper;
  for (let i = 0; i < ROUNDS; i++) {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
    input = digest.map(function(byte) {
      const value = (byte + 256) % 256;
      return ('0' + value.toString(16)).slice(-2);
    }).join('');
  }
  return input;
}

// FIX #18: Store session in ScriptCache with the FULL session timeout TTL
// (30 minutes = 1800 seconds) instead of the short boot TTL (300 seconds).
// ScriptCache max is 21600 seconds (6 hours), so 1800 is well within range.
// FIX #3: Removed UserProperties fallback — in a published web app
// UserProperties is scoped to the deploying account, not the end-user caller.
// ScriptCache is shared across all callers and is the correct store here.
function createSession_(user, rememberMe) {
  const token = Utilities.getUuid();
  const now = Date.now();
  const ttlSeconds = APP_CONFIG.sessionTimeoutMinutes * 60; // 1800
  const expiresAt = now + ttlSeconds * 1000;
  const primaryRestaurantId = resolvePrimaryRestaurantId_(user);

  // Resolve the brand name for the user's primary restaurant so the client
  // can show brand-specific UI without waiting for adminTools to load.
  let restaurantBrand = '';
  try {
    if (primaryRestaurantId) {
      const restRow = Repository.findOneBy_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', primaryRestaurantId);
      restaurantBrand = String((restRow && (restRow.brand_name || restRow.brandName)) || '').trim().toUpperCase();
    }
  } catch (e) {}

  // FIX #17: Normalize roleCode to canonical UPPER_SNAKE_CASE at session creation.
  const _rawRole = String(user.role_code || '').trim().toUpperCase().replace(/\s+/g, '_');
  const _validRoles = Object.keys(APP_CONFIG.roles).map(function(k) { return APP_CONFIG.roles[k]; });
  const _normalizedRole = _validRoles.indexOf(_rawRole) !== -1 ? _rawRole : user.role_code;

  const session = {
    token: token,
    userId: user.user_id,
    username: user.username,
    fullName: user.full_name,
    roleCode: _normalizedRole,
    restaurantId: primaryRestaurantId,
    restaurantBrand: restaurantBrand,
    region: user.region || '',
    rememberMe: rememberMe,
    expiresAt: expiresAt,
    mustResetPassword: String(user.must_reset_password || 'FALSE')  // FIX 3
  };

  // FIX #16: Namespace cache key with scriptId to prevent cross-script collisions.
  const _NS16 = String(APP_CONFIG.scriptId || '').slice(-8);
  CacheService.getScriptCache().put('sess:' + _NS16 + ':' + token, JSON.stringify(session), ttlSeconds);
  // FIX 5: Register token in per-user index for bulk invalidation on password change
  try { if (typeof registerSessionForUser_ === 'function') registerSessionForUser_(user.user_id, token); } catch(e) {}
  return session;
}

// FIX #3: Removed UserProperties fallback — rely solely on ScriptCache.
function getSessionOrThrow_(token) {
  if (!token) throw new Error('No active session.');

  const _NS16r = String(APP_CONFIG.scriptId || '').slice(-8);
  const raw = CacheService.getScriptCache().get('sess:' + _NS16r + ':' + token);
  if (!raw) throw new Error('Session not found.');

  const session = JSON.parse(raw);
  if (Date.now() > session.expiresAt) {
    invalidateSession_(token);
    throw new Error('Session expired.');
  }

  // Refresh the cache TTL so active users stay logged in across requests.
  const remainingMs = session.expiresAt - Date.now();
  const remainingSec = Math.max(Math.floor(remainingMs / 1000), 1);
  try {
    const _NS16x = String(APP_CONFIG.scriptId || '').slice(-8);
    CacheService.getScriptCache().put('sess:' + _NS16x + ':' + token, raw, remainingSec);
  } catch (e) {}

  return session;
}

// FIX #3: Only remove from ScriptCache (UserProperties removed).
function invalidateSession_(token) {
  const _NS16i = String(APP_CONFIG.scriptId || '').slice(-8);
  CacheService.getScriptCache().remove('sess:' + _NS16i + ':' + token);
}

function sanitizeSession_(session) {
  return {
    token: session.token,
    userId: session.userId,
    username: session.username,
    fullName: session.fullName,
    roleCode: session.roleCode,
    restaurantId: session.restaurantId,
    restaurantBrand: session.restaurantBrand || '',
    region: session.region,
    expiresAt: session.expiresAt
  };
}

// FIX #20: Cache the RESTAURANT_ACCESS lookup per user in ScriptCache
// to avoid a full-table scan on every login.
function resolvePrimaryRestaurantId_(user) {
  if (user.restaurant_id) return user.restaurant_id;

  // Try a per-user cache entry first (TTL 300 s — restaurant assignments rarely change).
  const cacheKey = 'primary_rest_' + String(user.user_id || '');
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached !== null) return cached;
  } catch (e) {}

  const accessRows = Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS).filter(function(item) {
    return String(item.user_id || '') === String(user.user_id || '');
  });

  let result = '';
  if (accessRows.length) {
    const primary = accessRows.find(function(item) {
      return String(item.is_primary || '').toUpperCase() === 'TRUE';
    });
    result = (primary || accessRows[0]).restaurant_id || '';
  }

  try {
    CacheService.getScriptCache().put(cacheKey, result, 60); // FIX #8: reduced 300s→60s so restaurant reassignments propagate faster
  } catch (e) {}

  return result;
}
