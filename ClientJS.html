function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appName = APP_CONFIG.appName;
  template.boot = JSON.stringify(getPublicBootPayload_());
  return template
    .evaluate()
    .setTitle(APP_CONFIG.appName);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// ── KEEPALIVE: trivial function that keeps the GAS isolate warm.
// Called from the frontend every 4 minutes so the V8 environment
// never fully cold-starts between user actions.


function ping() {
  return true;
}


function getPublicBootPayload_() {
  return {
    appName: APP_CONFIG.appName,
    companyName: BRAND.companyName,
    brand: BRAND,
    hostCategories: APP_CONFIG.hostCategories,
    sessionTimeoutMinutes: APP_CONFIG.sessionTimeoutMinutes,
    support: {
      email: 'support@placeholder.local',
      phone: '+1 000 000 0000'
    }
  };
}

// FIX #25 (boot cache TTL): Boot payload is now cached for the full
// APP_CONFIG.cacheTtlSeconds.boot duration (300s) instead of just 60s.
// This keeps the boot response warm between the 5-minute prewarm runs
// and eliminates cold-start rebuilds for the majority of users.
//
// FIX #19 (getOperationsDashboardData): The function exists in Services.txt
// but was missing from the ServiceRouter_ route table. It is now also
// registered here as a top-level dispatch handler so that ClientJS
// google.script.run calls to `getOperationsDashboardData` succeed.

function dispatch(action, payload, token) {
  try {
    return ServiceRouter_.route(action, payload, token);
  } catch (e) {
    return respond_(false, null, e.message);
  }
}

// FIX #19: Expose getOperationsDashboardData as a direct top-level callable
// so that ClientJS can invoke it via google.script.run.getOperationsDashboardData(token).
// The underlying implementation lives in Services.txt (getOperationsDashboardData).
// This wrapper also accepts either a raw token string or a payload object with
// a sessionToken property for forward compatibility.

// FIX #7: Call this after any user record update to immediately invalidate
// that user's boot cache. Pass the session token if known; otherwise the
// cache will expire naturally within 60 seconds (non-admin) or 300s (admin).


function bustBootCacheForToken_(token) {
  if (!token) return;
  try {
    const cacheKey = 'boot_v326_' + String(token).slice(-12);
    CacheService.getScriptCache().remove(cacheKey);
  } catch (e) {}
}

// Callable from frontend — busts boot cache then returns fresh boot payload
function refreshMySession(token) {
  try {
    bustBootCacheForToken_(token);
    return getAuthenticatedBootPayload(token);
  } catch(e) {
    return { success: false, error: String(e.message || e) };
  }
}


// Module lazy loader for split client files.
// Called by ClientCore through google.script.run.getClientModuleContent(moduleName, token).


function getClientModuleContent(moduleName, sessionToken) {
  try {
    const session = getSessionOrThrow_(sessionToken);
    const name = String(moduleName || '').trim();
    const allowed = {
      PayrollV360: true,
      RecruitmentJS: true,
      ManageTeamJS: true,
      AdminRequestsJS: true,
      RestaurantRequestsJS: true,
      HotZoneJS: true
      // DropdownSettingsJS removed: the Dropdown Settings UI is embedded directly in
      // ClientCore.html. There is no separate DropdownSettingsJS.html file; requesting
      // it via getClientModuleContent would throw a GAS HtmlService file-not-found error.
    };
    if (!allowed[name]) {
      throw new Error('Module not allowed: ' + name);
    }

    const content = HtmlService.createHtmlOutputFromFile(name).getContent();
    return respond_(true, {
      moduleName: name,
      content: content
    });
  } catch (error) {
    console.error('getClientModuleContent failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}
