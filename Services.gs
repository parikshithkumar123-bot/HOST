var DashboardService = {
  getDashboardBoot(session) {
    const boot = {
      hostLinks: HostLinkService.getLinksForSession(session),
      newsletters: NewsletterService.getActivePosts(session),
      notifications: NotificationService.getUnread(session),
      metrics: this.getTopMetrics(session),
      roleActions: RoleActionService.getActionsForSession(session),
      canManageWorkspace: canManageWorkspace_(session),
      isAdminToolsDeferred: canManageWorkspace_(session)
    };

    return boot;
  },

  // FIX #14: Stub removed — real implementation is the DashboardService.getTopMetrics
  // override defined at the bottom of this file (metrics_v329 with ScriptCache + live counts).
  // This placeholder ensures the method exists on the object at parse time.
  getTopMetrics(session) {
    // Overridden below — see "Real getTopMetrics" section at end of file.
    return { openRequests: 0 };
  }
};

var HostLinkService = {
  getLinksForSession(session) {
    const links = Repository.getAll_(APP_CONFIG.sheets.HOST_LINKS);
    const now = Date.now();
    return links.filter(function(link) {
      const roleMatch = !link.target_role_codes || String(link.target_role_codes).includes(session.roleCode);
      const restaurantMatch = !link.target_restaurant_ids ||
        String(link.target_restaurant_ids).trim() === '' ||
        String(link.target_restaurant_ids).includes(session.restaurantId);
      const activeStart = !link.start_at || new Date(link.start_at).getTime() <= now;
      const activeEnd = !link.end_at || new Date(link.end_at).getTime() >= now;
      return String(link.status).toUpperCase() === 'ACTIVE' && roleMatch && restaurantMatch && activeStart && activeEnd;
    });
  }
};

var NewsletterService = {
  getActivePosts(session) {
    const posts = Repository.getAll_(APP_CONFIG.sheets.NEWSLETTERS);
    const now = Date.now();
    return posts.filter(function(post) {
      const roleMatch = !post.audience_roles || String(post.audience_roles).includes(session.roleCode);
      const restaurantMatch = !post.audience_restaurants ||
        String(post.audience_restaurants).trim() === '' ||
        String(post.audience_restaurants).includes(session.restaurantId);
      const activeStart = !post.start_at || new Date(post.start_at).getTime() <= now;
      const activeEnd = !post.end_at || new Date(post.end_at).getTime() >= now;
      return String(post.status).toUpperCase() === 'ACTIVE' && roleMatch && restaurantMatch && activeStart && activeEnd;
    });
  }
};

var NotificationService = {
  getUnread(session) {
    // FIX #7: Cache unread notifications per role+restaurant for 60 seconds.
    // This avoids 5+ full sheet scans on every login/page load.
    const cacheKey = 'notif_unread_v1_' + String(session.roleCode || '') + '_' + String(session.restaurantId || 'all').slice(-8);
    const cache = CacheService.getScriptCache();
    try {
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {}

    const result = this._computeUnread(session);

    try { cache.put(cacheKey, JSON.stringify(result), 60); } catch (e) {}
    return result;
  },

  _computeUnread(session) {
    const items = [];
    const now = Date.now();
    const role = String(session.roleCode || '');
    const restaurantId = String(session.restaurantId || '');
    const isSuperAdmin = role === APP_CONFIG.roles.SUPER_ADMIN;
    const isAdminTeam = [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ].indexOf(role) !== -1;

    function clean(value) {
      return String(value == null ? '' : value).trim();
    }

    function upper(value) {
      return clean(value).toUpperCase();
    }

    function isOpenStatus(status) {
      const value = upper(status || 'NEW');
      return ['COMPLETED','RESOLVED','CLOSED','REJECTED','DELETED','INACTIVE','DONE'].indexOf(value) === -1;
    }

    // Row-limited sheet reader — reads at most `limit` rows from the bottom
    // (most recent) to avoid slow full-sheet scans on every 60-second cache miss.
    function safeRows(sheetName, limit) {
      try {
        const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return [];
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];
        const maxRows = limit && limit > 0 ? limit : 200;
        const startRow = Math.max(2, lastRow - maxRows + 1);
        const numRows = lastRow - startRow + 1;
        const lastCol = sheet.getLastColumn();
        if (lastCol < 1) return [];
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
          return String(h || '').trim().toLowerCase().replace(/\s+/g, '_');
        });
        const data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
        return data.map(function(row) {
          const obj = {};
          headers.forEach(function(h, i) { obj[h] = row[i]; });
          return obj;
        }).filter(function(row) {
          return headers.some(function(h) { return String(row[h] || '').trim() !== ''; });
        });
      } catch (error) {
        console.error('NotificationService safeRows failed for ' + sheetName, error);
        return [];
      }
    }

    function push(item) {
      if (!item || !item.title) return;
      items.push({
        notification_id: item.id || Utilities.getUuid(),
        title: item.title,
        message: item.message || '',
        priority: item.priority || 'Info',
        created_at: item.createdAt || new Date().toISOString(),
        type: item.type || 'Info',
        actionKey: item.actionKey || '',
        action_key: item.actionKey || '',
        source_id: item.sourceId || ''
      });
    }

    if (isAdminTeam) {
      safeRows(APP_CONFIG.sheets.CONTACT_REQUESTS, 200).forEach(function(row) {
        if (!isOpenStatus(row.status)) return;
        push({
          id: 'contact-' + clean(row.contact_request_id || row.request_id || row.requested_at),
          title: 'New Contact Us Request',
          message: clean(row.name || row.username || 'User') + (clean(row.message) ? ' - ' + clean(row.message) : ''),
          priority: upper(row.status || 'NEW') === 'NEW' ? 'New' : clean(row.status || 'Open'),
          createdAt: clean(row.requested_at || row.created_at || row.updated_at),
          type: 'Contact Request',
          actionKey: 'CONTACT_US_REQUESTS',
          sourceId: clean(row.contact_request_id || row.request_id)
        });
      });
    }

    if (isSuperAdmin || role === APP_CONFIG.roles.HR_ADMIN) {
      safeRows(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS, 100).forEach(function(row) {
        if (!isOpenStatus(row.status)) return;
        push({
          id: 'password-' + clean(row.password_request_id || row.requested_at),
          title: 'Password Change Request',
          message: clean(row.username || row.name || 'User') + ' needs password support',
          priority: clean(row.status || 'Pending'),
          createdAt: clean(row.requested_at || row.created_at || row.updated_at),
          type: 'Password Request',
          actionKey: 'PASSWORD_CHANGE_REQUESTS',
          sourceId: clean(row.password_request_id)
        });
      });
    }

    if (isAdminTeam) {
      safeRows(APP_CONFIG.sheets.REQUESTS_HR, 200).forEach(function(row) {
        if (!isOpenStatus(row.status)) return;
        if (restaurantId && role === APP_CONFIG.roles.RESTAURANT_USER && clean(row.restaurant_id) !== restaurantId) return;
        var payload = {};
        try { payload = JSON.parse(row.request_payload_json || '{}'); } catch (e) {}
        push({
          id: 'hiring-' + clean(row.request_id || row.submitted_at),
          title: 'Hiring Request',
          message: 'Restaurant ' + clean(row.restaurant_id || payload.restaurantId || '') + ' requested ' + clean(row.requested_for_role || payload.position || 'staff'),
          priority: clean(row.urgency || payload.priority || row.status || 'Pending'),
          createdAt: clean(row.submitted_at || row.created_at || row.updated_at),
          type: 'Hiring Request',
          actionKey: 'RECRUITMENT_REQUESTS',
          sourceId: clean(row.request_id)
        });
      });
    }

    safeRows(APP_CONFIG.sheets.NEWSLETTERS, 50).forEach(function(post) {
      const roleMatch = !post.audience_roles || String(post.audience_roles).trim() === '' || String(post.audience_roles).indexOf(role) !== -1;
      const restaurantMatch = !post.audience_restaurants || String(post.audience_restaurants).trim() === '' || String(post.audience_restaurants).indexOf(restaurantId) !== -1;
      const activeStart = !post.start_at || new Date(post.start_at).getTime() <= now;
      const activeEnd = !post.end_at || new Date(post.end_at).getTime() >= now;
      if (upper(post.status || 'ACTIVE') !== 'ACTIVE' || !roleMatch || !restaurantMatch || !activeStart || !activeEnd) return;
      push({
        id: 'newsletter-' + clean(post.post_id || post.title),
        title: 'Active Newsletter',
        message: clean(post.title || 'Newsletter available'),
        priority: clean(post.priority || 'Info'),
        createdAt: clean(post.created_at || post.start_at || post.updated_at),
        type: 'Newsletter',
        actionKey: 'NEWSLETTER_BOARD',
        sourceId: clean(post.post_id)
      });
    });

    items.sort(function(a, b) {
      return new Date(b.created_at || 0).getTime() - new Date(a.created_at || 0).getTime();
    });

    return items.slice(0, 50);
  }
};

var DropdownService = {
  getOptionsMap() {
    const rows = Repository.getAll_(APP_CONFIG.sheets.DROPDOWN_OPTIONS).filter(function(row) {
      return String(row.status || 'ACTIVE').toUpperCase() === 'ACTIVE';
    });
    const grouped = {};
    rows.forEach(function(row) {
      const key = String(row.dropdown_key || '').trim();
      if (!key) return;
      grouped[key] = grouped[key] || [];
      grouped[key].push({
        optionId: row.option_id,
        value: row.option_value || '',
        sortOrder: Number(row.sort_order || 0)
      });
    });
    Object.keys(grouped).forEach(function(key) {
      grouped[key].sort(function(a, b) { return a.sortOrder - b.sortOrder || String(a.value).localeCompare(String(b.value)); });
      grouped[key] = grouped[key].map(function(item) { return item.value; });
    });
    return grouped;
  }
};

var RoleActionService = {
  // Legacy compatibility for roles that had their button rows saved before
  // PAYROLL_ADMIN existed. We only backfill when the saved set exactly matches
  // the older default snapshot for that role.
  normalizeActionKeysV367_: function(items) {
    return uniqueValues_((Array.isArray(items) ? items : [])
      .map(function(item) {
        return String(item || '').trim();
      })
      .filter(function(item) {
        return Boolean(item) &&
          item !== ROLE_ACTION_NO_ACCESS_MARKER &&
          ALL_ROLE_ACTION_BUTTONS.indexOf(item) !== -1;
      }));
  },

  applyLegacyPayrollAdminBackfillV367_: function(roleCode, actionKeys) {
    const normalized = this.normalizeActionKeysV367_(actionKeys);
    const defaults = this.normalizeActionKeysV367_(DEFAULT_ROLE_ACTION_PERMISSIONS[roleCode] || []);

    if (defaults.indexOf('PAYROLL_ADMIN') === -1) return normalized;
    if (normalized.indexOf('PAYROLL_ADMIN') !== -1) return normalized;

    const legacyDefaults = defaults.filter(function(key) {
      return key !== 'PAYROLL_ADMIN';
    });
    const matchesLegacyDefaults =
      normalized.length === legacyDefaults.length &&
      legacyDefaults.every(function(key) {
        return normalized.indexOf(key) !== -1;
      });

    if (!matchesLegacyDefaults) return normalized;
    return uniqueValues_(normalized.concat(['PAYROLL_ADMIN']));
  },

  getActionsForSession(session) {
    if (String(session.roleCode) === String(APP_CONFIG.roles.SUPER_ADMIN)) {
      return ALL_ROLE_ACTION_BUTTONS.map(function(key) {
        return {
          key: key,
          label: formatActionLabel_(key)
        };
      });
    }

    const permissions = Repository.getAll_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS);

    const roleRows = permissions.filter(function(item) {
      return String(item.role_code) === String(session.roleCode);
    });

    const configured = roleRows.filter(function(item) {
      return String(item.is_allowed).toUpperCase() === 'TRUE';
    }).map(function(item) {
      return String(item.button_key || '').trim();
    }).filter(function(buttonKey) {
      return Boolean(buttonKey) && buttonKey !== ROLE_ACTION_NO_ACCESS_MARKER;
    });

    const defaults = DEFAULT_ROLE_ACTION_PERMISSIONS[session.roleCode] || [];
    const actionKeys = roleRows.length
      ? this.applyLegacyPayrollAdminBackfillV367_(session.roleCode, configured)
      : this.normalizeActionKeysV367_(defaults);

    return actionKeys.map(function(key) {
      return {
        key: key,
        label: formatActionLabel_(key)
      };
    });
  },

  savePermissions(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);

    const roleCode = String(payload.roleCode || '').trim();
    if (!roleCode) throw new Error('Role is required.');

    if (roleCode === APP_CONFIG.roles.SUPER_ADMIN) {
      throw new Error('Super Admin always has all buttons. You do not need to edit Super Admin buttons.');
    }

    const validRoles = Object.keys(APP_CONFIG.roles).map(function(key) {
      return APP_CONFIG.roles[key];
    });

    if (validRoles.indexOf(roleCode) === -1) {
      throw new Error('Invalid role selected.');
    }

    const allowedButtons = uniqueValues_(
      (Array.isArray(payload.allowedButtons) ? payload.allowedButtons : [])
        .map(function(item) {
          return String(item || '').trim();
        })
        .filter(function(item) {
          return ALL_ROLE_ACTION_BUTTONS.indexOf(item) !== -1;
        })
    );

    const now = new Date().toISOString();

    Repository.replaceRowsByColumn_(
      APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS,
      'role_code',
      roleCode,
      (allowedButtons.length ? allowedButtons : [ROLE_ACTION_NO_ACCESS_MARKER]).map(function(buttonKey) {
        return {
          permission_id: Utilities.getUuid(),
          role_code: roleCode,
          button_key: buttonKey,
          is_allowed: buttonKey === ROLE_ACTION_NO_ACCESS_MARKER ? 'FALSE' : 'TRUE',
          updated_at: now,
          updated_by: session.userId
        };
      })
    );

    logAuditEvent_('ROLE_ACTION_PERMISSIONS_UPDATED', 'ROLE', roleCode, session.userId, {
      allowedButtons: allowedButtons
    });

    // FIX #12: Invalidate BOTH the V328 derived-permissions cache AND the
    // underlying V258 getAll_ raw-data cache, so neither stale layer survives.
    try { CacheService.getScriptCache().remove('role_btn_perms_v329'); } catch (e) {}
    try { Repository.clearCacheV258_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS); } catch (e) {}

    return respond_(true, {
      roleCode: roleCode,
      allowedButtons: allowedButtons
    });
  }
};

function getRecentHiringRequests(token) {
  const session = getSessionOrThrow_(token);
  // FIX #5: Enforce role-based access — unauthenticated or unauthorized callers
  // must not be able to enumerate hiring requests.
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const visibleIds = getRecruitmentRestaurantScopeIds_(session);

  const rows = Repository.getAll_(APP_CONFIG.sheets.REQUESTS_HR)
    .filter(function(row) {
      if (visibleIds !== null && visibleIds.indexOf(String(row.restaurant_id || '')) === -1) return false;
      return row.status !== 'DELETED';
    })
    .sort(function(a, b) {
      return new Date(b.submitted_at || 0) - new Date(a.submitted_at || 0);
    })
    .slice(0, 10)
    .map(function(row) {
      let payload = {};
      try {
        payload = JSON.parse(row.request_payload_json || '{}');
      } catch (e) { console.warn('Ignored service error', e); }

      return {
        requestId: row.request_id || '',
        position: row.requested_for_role || payload.position || '',
        employmentType: payload.employmentType || '',
        submitterName: payload.submitterName || row.requested_by || '',
        submittedAt: row.submitted_at || '',
        status: row.status === 'COMPLETED' ? 'Hiring Done' : 'Viewed'
      };
    });

  return respond_(true, { requests: rows });
}

var ContactRequestService = {
  create(payload) {
    // FIX #4: Require a valid authenticated session before accepting contact requests.
    // Previously any unauthenticated POST could write to the Contact Requests sheet
    // with no rate limiting, allowing spam/abuse.
    const session = getSessionOrThrow_(payload.sessionToken);
    const now = new Date().toISOString();
    const requestId = Utilities.getUuid();
    Repository.append_(APP_CONFIG.sheets.CONTACT_REQUESTS, {
      contact_request_id: requestId,
      name: String(payload.name || (session && session.fullName) || '').trim(),
      username: String(payload.username || (session && session.username) || '').trim(),
      email: String(payload.email || '').trim(),
      phone: String(payload.phone || '').trim(),
      message: String(payload.message || '').trim(),
      source: String(payload.source || 'PORTAL').trim(),
      requested_at: now,
      status: 'NEW',
      action_taken_notes: '',
      updated_at: now,
      updated_by: session ? session.userId : ''
    });
    return respond_(true, { requestId: requestId });
  },

  update(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
    Repository.updateById_(APP_CONFIG.sheets.CONTACT_REQUESTS, 'contact_request_id', payload.contactRequestId, {
      status: String(payload.status || '').trim() || 'IN_PROGRESS',
      action_taken_notes: String(payload.actionTakenNotes || '').trim(),
      updated_at: new Date().toISOString(),
      updated_by: session.userId
    });
    return respond_(true, { contactRequestId: payload.contactRequestId });
  }
};

var PasswordChangeService = {
  create(payload) {
    const now = new Date().toISOString();
    // FIX #3: Require valid session — prevents unauthenticated writes to Password Change Requests sheet.
    const session = getSessionOrThrow_(payload.sessionToken);
    const requestId = Utilities.getUuid();
    Repository.append_(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS, {
      password_request_id: requestId,
      name: String(payload.name || (session && session.fullName) || '').trim(),
      username: String(payload.username || (session && session.username) || '').trim(),
      email: String(payload.email || '').trim(),
      phone: String(payload.phone || '').trim(),
      message: String(payload.message || 'Password reset requested').trim(),
      requested_at: now,
      status: 'NEW',
      temporary_password: '',
      action_taken_notes: '',
      updated_at: now,
      updated_by: session ? session.userId : ''
    });
    return respond_(true, { requestId: requestId });
  },

  fulfill(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
    const request = Repository.findOneBy_(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS, 'password_request_id', payload.passwordRequestId);
    if (!request) throw new Error('Password request not found.');
    const user = request.username ? Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'username', request.username) : null;
    const tempPassword = String(payload.temporaryPassword || '').trim();
    if (tempPassword.length < 8) throw new Error('Temporary password must be at least 8 characters.');

    if (user) {
      const salt = Utilities.getUuid().replace(/-/g, '');
      Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', user.user_id, {
        password_hash: hashPassword_(tempPassword, salt),
        salt: salt,
        must_reset_password: 'TRUE',
        updated_at: new Date().toISOString(),
        updated_by: session.userId
      });
    }

    Repository.updateById_(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS, 'password_request_id', payload.passwordRequestId, {
      status: 'COMPLETED',
      temporary_password: '',  // not stored in plain text
      action_taken_notes: String(payload.actionTakenNotes || '').trim(),
      updated_at: new Date().toISOString(),
      updated_by: session.userId
    });
    // FIX #1: Do NOT return the plaintext temporaryPassword in the API response.
    // The admin already knows the password they typed; sending it back over the wire
    // (and potentially logging it) is unnecessary and a security risk.
    return respond_(true, { passwordRequestId: payload.passwordRequestId });
  }
};

var BrandService = {
  getAll() {
    return Repository.getAll_(APP_CONFIG.sheets.BRANDS).map(function(row) {
      return {
        brandId: row.brand_id,
        brandName: row.brand_name || '',
        status: row.status || 'ACTIVE',
        notes: row.notes || ''
      };
    });
  },

  save(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN]);
    const brandId = String(payload.brandId || '').trim() || Utilities.getUuid();
    const now = new Date().toISOString();
    const record = {
      brand_id: brandId,
      brand_name: String(payload.brandName || '').trim(),
      status: String(payload.status || 'ACTIVE').trim() || 'ACTIVE',
      notes: String(payload.notes || '').trim(),
      updated_at: now
    };
    if (!record.brand_name) throw new Error('Brand name is required.');
    const existing = Repository.findOneBy_(APP_CONFIG.sheets.BRANDS, 'brand_id', brandId);
    if (existing) {
      Repository.updateById_(APP_CONFIG.sheets.BRANDS, 'brand_id', brandId, record);
    } else {
      record.created_at = now;
      Repository.append_(APP_CONFIG.sheets.BRANDS, record);
    }
    return respond_(true, { brandId: brandId });
  }
};

var RequestService = {
  createRequest(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SUPER_ADMIN
    ]);
    validateRequestAccess_(session, payload);

    const now = new Date().toISOString();
    const requestId = Utilities.getUuid();
    const sheetName = resolveRequestSheet_(payload.requestGroup);
    const record = {
      request_id: requestId,
      request_type: payload.requestType,
      restaurant_id: payload.restaurantId || session.restaurantId,
      requested_by: session.userId,
      status: 'PENDING',
      submitted_at: now,
      request_payload_json: JSON.stringify(payload || {}),
      timeline_json: JSON.stringify([
        {
          at: now,
          by: session.userId,
          action: 'CREATED',
          note: payload.businessReason || payload.description || ''
        }
      ])
    };

    Object.keys(payload || {}).forEach(function(key) {
      if (!(key in record)) record[key] = payload[key];
    });

    Repository.append_(sheetName, record);
    logAuditEvent_('REQUEST_CREATED', 'REQUEST', requestId, session.userId, { sheetName: sheetName });
    return respond_(true, { requestId: requestId });
  }
};

function createRequest(payload) {
  return RequestService.createRequest(payload);
}

function getNewHeroEmployees(token) {
  const session = getSessionOrThrow_(token)

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.RESTAURANT_USER
  ])

  ensureTeamMembersRecruitmentColumns_()

  const restaurants = AdminService.getRestaurants_()
  const visibleRestaurantIds = getRecruitmentRestaurantScopeIds_(session)
  const restaurantLookup = {}

  restaurants.forEach(function(r) {
    restaurantLookup[String(r.restaurantId)] = r
  })

  const visibleRestaurants = visibleRestaurantIds === null
    ? restaurants
    : restaurants.filter(function(r) {
        return visibleRestaurantIds.indexOf(String(r.restaurantId)) !== -1
      })

  const teamMembers = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS)
    .filter(function(row) {
      return teamMemberInRecruitmentScope_(row, visibleRestaurantIds)
    })
    .sort(function(a, b) {
      return new Date(b.created_at || b.updated_at || 0) - new Date(a.created_at || a.updated_at || 0)
    })
    .map(function(row) {
      const restaurantIds = String(row.restaurant_id || '')
        .split(',')
        .map(function(item) {
          return item.trim()
        })
        .filter(Boolean)

      const restaurantNames = restaurantIds.map(function(id) {
        const r = restaurantLookup[id]
        if (!r) return id
        return (r.restaurantCode ? r.restaurantCode + ' ' : '') + r.restaurantName
      })

      const restaurantCodes = restaurantIds.map(function(id) {
        const r = restaurantLookup[id]
        if (!r) return ''
        return String(r.restaurantCode || '').trim()
      }).filter(Boolean)

      return {
        teamMemberId: row.team_member_id || '',
        name: row.full_name || '',
        mobile: row.phone || '',
        hiredDate: row.hire_date || '',
        restaurantIds: restaurantIds,
        restaurantCodes: restaurantCodes,
        restaurantCode: restaurantCodes.join(', '),
        restaurantName: restaurantNames.join(', '),
        status: row.status || 'Ready to Start',
        outcomeStatus: row.recruitment_outcome_status || '',
        completedAt: row.recruitment_completed_at || '',
        quitReason: row.recruitment_reason || '',
        createdAt: row.created_at || '',
        updatedAt: row.updated_at || ''
      }
    })

  return respond_(true, {
    restaurants: visibleRestaurants,
    teamMembers: teamMembers,
    employees: teamMembers
  })
}

function ensureTeamMembersRecruitmentColumns_() {
  // FIX #13: Use a ScriptCache flag so the column-existence check only runs
  // once (or after the cache expires in 6 hours) instead of on every request.
  const flagKey = 'team_recruit_cols_ok_v1';
  try {
    if (CacheService.getScriptCache().get(flagKey) === 'true') return;
  } catch (e) {}

  const sheetName = APP_CONFIG.sheets.TEAM_MEMBERS;
  const sheet = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId).getSheetByName(sheetName);

  if (!sheet) throw new Error('Team Members sheet not found.');

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });

  const neededHeaders = [
    'recruitment_reason',
    'recruitment_outcome_status',
    'recruitment_completed_at'
  ];

  let changed = false;
  neededHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(header);
      headers.push(header);
      changed = true;
    }
  });

  // Only set the flag once all columns are confirmed present.
  try { CacheService.getScriptCache().put(flagKey, 'true', 21600); } catch (e) {}
}

function getRecruitmentRestaurantScopeIds_(session) {
  // FIX #5: SENIOR_MANAGEMENT gets full cross-restaurant view (same as SUPER_ADMIN/HR_ADMIN)
  if (
    session.roleCode === APP_CONFIG.roles.SUPER_ADMIN ||
    session.roleCode === APP_CONFIG.roles.HR_ADMIN ||
    session.roleCode === APP_CONFIG.roles.SENIOR_MANAGEMENT
  ) {
    return null
  }

  const ids = []

  if (session.restaurantId) ids.push(String(session.restaurantId))

  Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS).forEach(function(row) {
    if (String(row.user_id || '') === String(session.userId || '') && row.restaurant_id) {
      ids.push(String(row.restaurant_id))
    }
  })

  return uniqueValues_(ids)
}

function teamMemberInRecruitmentScope_(row, visibleRestaurantIds) {
  if (visibleRestaurantIds === null) return true;
  if (!visibleRestaurantIds.length) return false;

  const restaurantIds = String(row.restaurant_id || '')
    .split(',')
    .map(function(item) {
      return item.trim();
    })
    .filter(Boolean);

  return restaurantIds.some(function(id) {
    return visibleRestaurantIds.indexOf(String(id)) !== -1;
  });
}

function saveNewHeroEmployee(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const name = String(payload.name || '').trim();
  const mobile = String(payload.mobile || '').trim();
  const hiredDate = String(payload.hiredDate || '').trim();
  const restaurantIds = Array.isArray(payload.restaurantIds)
    ? payload.restaurantIds
    : Array.isArray(payload.assignedRestaurants)
      ? payload.assignedRestaurants
      : [];

  const cleanRestaurantIds = restaurantIds.map(function(id) {
    return String(id || '').trim();
  }).filter(Boolean);

  if (!name) throw new Error('Name is required.');
  if (!mobile || !/^[0-9+\-\s()]{7,20}$/.test(mobile)) {
    throw new Error('Enter a valid mobile number.');
  }
  if (!hiredDate) throw new Error('Hired date is required.');
  if (!cleanRestaurantIds.length) throw new Error('Select at least one restaurant.');

  const now = new Date().toISOString();
  const teamMemberId = Utilities.getUuid();

  Repository.append_(APP_CONFIG.sheets.TEAM_MEMBERS, {
    team_member_id: teamMemberId,
    employee_id: teamMemberId,
    full_name: name,
    restaurant_id: cleanRestaurantIds.join(','),
    job_title: 'Team Member',
    department: '',
    employment_type: '',
    hire_date: hiredDate,
    status: 'Ready to Start',
    manager_name: '',
    email: '',
    phone: mobile,
    created_at: now,
    updated_at: now
  });

  logAuditEvent_('TEAM_MEMBER_CREATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'NEW_HEROES',
    name: name,
    restaurantIds: cleanRestaurantIds
  });

  return respond_(true, {
    teamMemberId: teamMemberId
  });
}

function updateNewHeroEmployee(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const teamMemberId = String(payload.teamMemberId || '').trim();
  const name = String(payload.name || '').trim();
  const mobile = String(payload.mobile || '').trim();
  const hiredDate = String(payload.hiredDate || '').trim();
  const restaurantIds = Array.isArray(payload.restaurantIds)
    ? payload.restaurantIds
    : Array.isArray(payload.assignedRestaurants)
      ? payload.assignedRestaurants
      : [];

  const cleanRestaurantIds = restaurantIds.map(function(id) {
    return String(id || '').trim();
  }).filter(Boolean);

  if (!teamMemberId) throw new Error('Employee ID is required.');
  if (!name) throw new Error('Name is required.');
  if (!mobile || !/^[0-9+\-\s()]{7,20}$/.test(mobile)) {
    throw new Error('Enter a valid mobile number.');
  }
  if (!hiredDate) throw new Error('Hired date is required.');
  if (!cleanRestaurantIds.length) throw new Error('Select at least one restaurant.');

  const existing = Repository.findOneBy_(
    APP_CONFIG.sheets.TEAM_MEMBERS,
    'team_member_id',
    teamMemberId
  );
  if (!existing) throw new Error('Employee not found.');

  Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, {
    full_name: name,
    phone: mobile,
    hire_date: hiredDate,
    restaurant_id: cleanRestaurantIds.join(','),
    status: existing.status || 'Ready to Start',
    updated_at: new Date().toISOString()
  });

  logAuditEvent_('TEAM_MEMBER_UPDATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'NEW_HEROES',
    name: name,
    restaurantIds: cleanRestaurantIds
  });

  return respond_(true, {
    teamMemberId: teamMemberId
  });
}

function deleteNewHeroEmployee(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const teamMemberId = String(payload.teamMemberId || '').trim();
  if (!teamMemberId) throw new Error('Employee ID is required.');

  Repository.deleteById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId);

  logAuditEvent_('TEAM_MEMBER_DELETED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'NEW_HEROES'
  });

  return respond_(true, {
    teamMemberId: teamMemberId
  });
}



/*******************************************************
 V310 USER MANAGEMENT PERMISSION AND ERROR VISIBILITY
 Allows roles that were granted USER_MANAGEMENT button permission
 to load the Users page data. Super Admin remains full access.
 HR or other granted roles cannot create or edit SUPER_ADMIN users.
*******************************************************/
function hasRoleActionPermissionV310_(session, buttonKey) {
  if (!session || !buttonKey) return false;
  if (String(session.roleCode) === String(APP_CONFIG.roles.SUPER_ADMIN)) return true;
  try {
    const actions = RoleActionService.getActionsForSession(session) || [];
    return actions.some(function(action) {
      return String(action.key || '') === String(buttonKey);
    });
  } catch (error) {
    console.error('Role action permission check failed', error);
    return false;
  }
}

function assertUserManagementAccessV310_(session) {
  if (!session) throw new Error('No active session.');
  if (String(session.roleCode) === String(APP_CONFIG.roles.SUPER_ADMIN)) return true;
  if (hasRoleActionPermissionV310_(session, 'USER_MANAGEMENT')) return true;
  throw new Error('You are not authorized to manage users.');
}

function assertSafeUserRoleChangeV310_(session, targetRoleCode) {
  const targetRole = String(targetRoleCode || '').trim();
  if (String(session.roleCode) === String(APP_CONFIG.roles.SUPER_ADMIN)) return true;
  if (targetRole === String(APP_CONFIG.roles.SUPER_ADMIN)) {
    throw new Error('Only Super Admin can create or edit Super Admin users.');
  }
  return true;
}

function logServiceErrorV310_(source, error, metadata) {
  try {
    console.error(source, error);
    if (typeof logAuditEvent_ === 'function') {
      logAuditEvent_('SERVICE_ERROR', 'SYSTEM', source || 'UNKNOWN', '', {
        error: String(error && error.message ? error.message : error),
        metadata: metadata || {}
      });
    }
  } catch (ignored) {}
}



/*******************************************************
 V318 RESTAURANT SCOPE SAFETY
 Keeps Super Admin, HR Admin, and Senior Management global.
 Restricts MRO, Restaurant User, and other assigned users to
 their assigned restaurants from Restaurant Access and session.
*******************************************************/
function cleanV318_(value) {
  return String(value == null ? '' : value).trim();
}

function splitCsvV318_(value) {
  return cleanV318_(value)
    .split(',')
    .map(function(item) { return cleanV318_(item); })
    .filter(Boolean);
}

function normalizeRoleCodeV320_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function isRoleMatchV320_(sessionRole, configRole, fallbackNames) {
  const role = normalizeRoleCodeV320_(sessionRole);
  const target = normalizeRoleCodeV320_(configRole);
  if (role && target && role === target) return true;
  return (fallbackNames || []).some(function(name) {
    return role === normalizeRoleCodeV320_(name);
  });
}

function isGlobalRestaurantScopeV318_(session) {
  if (!session) return false;
  const role = session.roleCode || session.role_code || session.role || session.userRole || '';
  return isRoleMatchV320_(role, APP_CONFIG.roles.SUPER_ADMIN, ['SUPER_ADMIN', 'SUPER ADMIN', 'ADMIN']) ||
    isRoleMatchV320_(role, APP_CONFIG.roles.HR_ADMIN, ['HR_ADMIN', 'HR ADMIN', 'HR']) ||
    isRoleMatchV320_(role, APP_CONFIG.roles.SENIOR_MANAGEMENT, ['SENIOR_MANAGEMENT', 'SENIOR MANAGEMENT', 'SENIOR']);
}

function getSessionRestaurantScopeV318_(session) {
  if (!session) return [];
  if (isGlobalRestaurantScopeV318_(session)) return null;

  const ids = [];

  if (session.restaurantId) ids.push(cleanV318_(session.restaurantId));
  if (session.restaurant_id) ids.push(cleanV318_(session.restaurant_id));
  if (session.restaurantCode) ids.push(cleanV318_(session.restaurantCode));
  if (session.restaurant_code) ids.push(cleanV318_(session.restaurant_code));

  try {
    Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS).forEach(function(row) {
      if (String(row.user_id || '') !== String(session.userId || '')) return;
      if (row.restaurant_id) ids.push(cleanV318_(row.restaurant_id));
      if (row.restaurant_code) ids.push(cleanV318_(row.restaurant_code));
      if (row.restaurant_number) ids.push(cleanV318_(row.restaurant_number));
    });
  } catch (error) {
    console.error('Restaurant scope lookup failed V318', error);
  }

  return uniqueValues_(ids.filter(Boolean));
}

function restaurantMatchesScopeV318_(restaurant, scopeIds) {
  if (scopeIds === null) return true;
  if (!scopeIds || !scopeIds.length) return false;

  const values = [
    restaurant.restaurantId,
    restaurant.restaurant_id,
    restaurant.restaurantCode,
    restaurant.restaurant_code,
    restaurant.restaurantNumber,
    restaurant.restaurant_number,
    restaurant.id
  ].map(cleanV318_).filter(Boolean);

  return values.some(function(value) {
    return scopeIds.indexOf(value) !== -1;
  });
}

function rowMatchesRestaurantScopeV318_(row, scopeIds) {
  if (scopeIds === null) return true;
  if (!scopeIds || !scopeIds.length) return false;

  const values = [];
  [
    'restaurant_id',
    'restaurantId',
    'restaurant_code',
    'restaurantCode',
    'restaurant_number',
    'restaurantNumber',
    'assigned_restaurant_ids',
    'assignedRestaurantIds',
    'target_restaurant_ids',
    'audience_restaurants'
  ].forEach(function(field) {
    splitCsvV318_(row[field]).forEach(function(value) {
      values.push(value);
    });
  });

  return values.some(function(value) {
    return scopeIds.indexOf(value) !== -1;
  });
}

function filterRestaurantsForSessionV318_(session, restaurants) {
  const scopeIds = getSessionRestaurantScopeV318_(session);
  return (restaurants || []).filter(function(restaurant) {
    return restaurantMatchesScopeV318_(restaurant, scopeIds);
  });
}

function filterRowsForSessionByRestaurantV318_(session, rows) {
  const scopeIds = getSessionRestaurantScopeV318_(session);
  return (rows || []).filter(function(row) {
    return rowMatchesRestaurantScopeV318_(row, scopeIds);
  });
}

function filterUsersForSessionByRestaurantV318_(session, users) {
  const scopeIds = getSessionRestaurantScopeV318_(session);
  if (scopeIds === null) return users || [];
  if (!scopeIds || !scopeIds.length) return [];

  const accessRows = Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS);
  const allowedUsers = {};

  accessRows.forEach(function(row) {
    const userId = cleanV318_(row.user_id);
    if (!userId) return;
    const rowIds = splitCsvV318_(row.restaurant_id)
      .concat(splitCsvV318_(row.restaurant_code))
      .concat(splitCsvV318_(row.restaurant_number));
    if (rowIds.some(function(id) { return scopeIds.indexOf(id) !== -1; })) {
      allowedUsers[userId] = true;
    }
  });

  return (users || []).filter(function(user) {
    const userId = cleanV318_(user.userId || user.user_id);
    if (userId && allowedUsers[userId]) return true;
    return rowMatchesRestaurantScopeV318_(user, scopeIds);
  });
}
var AdminService = {
  getUserManagementBoot_(session) {
    assertUserManagementAccessV310_(session);

    const restaurants = this.getRestaurantsForManagementV320_(session);
    const users = this.getUsersForSession_(session);
    const hostLinks = this.getHostLinks_();
    const teamMembers = this.getEmployeesForSession_(session);
    const newHires = this.getNewHiresForSession_(session);
    const restaurantMappings = this.getRestaurantMappingsForSession_(users, restaurants, session);
    const brands = BrandService.getAll();
    const newsletters = this.getNewsletters_();
    // FIX #8: Do NOT load full contactRequests/passwordRequests arrays into the boot payload.
    // They can be thousands of rows and push the payload well past 90 KB, busting the cache.
    // Counts only are included in the boot; the full lists are loaded lazily on demand
    // when the admin opens the relevant section (listContactRequests / listPasswordChangeRequests).
    const contactRequestCount = Repository.getAll_(APP_CONFIG.sheets.CONTACT_REQUESTS).length;
    const passwordRequestCount = Repository.getAll_(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS).length;
    const buttonPermissions = this.getButtonPermissionMatrix_();
    const dropdownOptions = DropdownService.getOptionsMap();
    const activeUsers = users.filter(function(user) { return String(user.status).toUpperCase() === 'ACTIVE'; }).length;
    const inactiveUsers = users.length - activeUsers;
    const websiteControlRows = buildWebsiteControlRows_({
      totalRestaurants: restaurants.length,
      totalUsers: users.length,
      activeUsers: activeUsers,
      inactiveUsers: inactiveUsers,
      hostLinkCount: hostLinks.length,
      newsletterCount: newsletters.length,
      contactRequestCount: contactRequestCount,
      passwordRequestCount: passwordRequestCount
    });
    const reportRows = buildMetricRows_(DashboardService.getTopMetrics(session), 'All-company reporting overview');

    return {
      roles: Object.keys(APP_CONFIG.roles).map(function(key) {
        return APP_CONFIG.roles[key];
      }),
      restaurants: restaurants,
      hostLinks: hostLinks,
      users: users,
      employees: teamMembers,
      newHires: newHires,
      restaurantMappings: restaurantMappings,
      brands: brands,
      newsletters: newsletters,
      websiteControlRows: websiteControlRows,
      reportRows: reportRows,
      // FIX #8: contactRequests excluded from boot — loaded lazily via listContactRequests action.
      // FIX #8: passwordRequests excluded from boot — loaded lazily via listPasswordChangeRequests action.
      buttonPermissions: buttonPermissions,
      dropdownOptions: dropdownOptions,
      allowedSections: [
        'WEBSITE_CONTROL',
        'HOST_LINKS',
        'ROLE_BUTTON_CONTROL',
        'RESTAURANTS',
        'BRANDS',
        'USERS',
        'EMPLOYEES',
        'NEW_HIRES',
        'RESTAURANT_MAPPINGS',
        'NEWSLETTERS',
        'CONTACT_REQUESTS',
        'PASSWORD_CHANGE_REQUESTS',
        'DROPDOWN_OPTIONS'
      ],
      stats: {
        totalRestaurants: restaurants.length,
        totalUsers: users.length,
        activeUsers: activeUsers,
        inactiveUsers: inactiveUsers,
        hostLinkCount: hostLinks.length,
        brandCount: brands.length,
        contactRequestCount: contactRequestCount,
        passwordRequestCount: passwordRequestCount,
        employeeCount: teamMembers.length,
        newHireCount: newHires.length,
        mappingCount: restaurantMappings.length
      }
    };
  },

  getRestaurants_() {
    return Repository.getAll_(APP_CONFIG.sheets.RESTAURANTS).map(function(restaurant) {
      const restaurantNumber = String(
        restaurant.restaurant_number ||
        restaurant.restaurant_code ||
        restaurant.restaurantNumber ||
        restaurant.restaurantCode ||
        ''
      ).trim();

      const restaurantId = String(
        restaurant.restaurant_id ||
        restaurant.restaurantId ||
        (restaurantNumber ? 'REST_' + restaurantNumber : '')
      ).trim();

      const restaurantName = String(restaurant.restaurant_name || restaurant.restaurantName || restaurant.name || '').trim();
      const brandName = String(restaurant.brand || restaurant.brand_name || restaurant.brandName || '').trim();

      return {
        restaurantId: restaurantId,
        restaurant_id: restaurantId,
        restaurantCode: restaurantNumber,
        restaurant_code: restaurantNumber,
        restaurantNumber: restaurantNumber,
        restaurant_number: restaurantNumber,
        restaurantName: restaurantName,
        restaurant_name: restaurantName,
        brandName: brandName,
        brand_name: brandName,
        brand: brandName,
        region: restaurant.region || '',
        city: restaurant.city || '',
        provinceState: restaurant.province_state || '',
        province_state: restaurant.province_state || '',
        country: restaurant.country || '',
        status: restaurant.status || 'ACTIVE',
        openingDate: restaurant.opening_date || '',
        opening_date: restaurant.opening_date || '',
        gmName: restaurant.gm_name || '',
        gm_name: restaurant.gm_name || '',
        gmEmail: restaurant.gm_email || '',
        gm_email: restaurant.gm_email || '',
        mroOwnerUserId: restaurant.mro_owner_user_id || '',
        mro_owner_user_id: restaurant.mro_owner_user_id || '',
        notes: restaurant.notes || '',
        createdAt: restaurant.created_at || '',
        created_at: restaurant.created_at || '',
        updatedAt: restaurant.updated_at || '',
        updated_at: restaurant.updated_at || ''
      };
    }).filter(function(restaurant) {
      return restaurant.restaurantNumber || restaurant.restaurantName;
    });
  },

  getRestaurantsForSession_(session) {
    return filterRestaurantsForSessionV318_(session, this.getRestaurants_());
  },

  getRestaurantsForManagementV320_(session) {
    const allRestaurants = this.getRestaurants_();
    if (isGlobalRestaurantScopeV318_(session)) return allRestaurants;
    return filterRestaurantsForSessionV318_(session, allRestaurants);
  },

  getEmployeesForSession_(session) {
    return filterRowsForSessionByRestaurantV318_(session, Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS)).map(function(item) {
      return {
        teamMemberId: item.team_member_id,
        full_name: item.full_name || '',
        fullName: item.full_name || '',
        job_title: item.job_title || '',
        restaurant_id: item.restaurant_id || '',
        restaurantId: item.restaurant_id || '',
        status: item.status || 'ACTIVE',
        email: item.email || '',
        phone: item.phone || '',
        employee_id: item.employee_id || '',
        notes: item.notes || ''
      };
    });
  },

  getNewHiresForSession_(session) {
    return filterRowsForSessionByRestaurantV318_(session, Repository.getAll_(APP_CONFIG.sheets.NEW_HIRES)).map(function(item) {
      return {
        hireId: item.hire_id,
        hire_id: item.hire_id,
        full_name: item.candidate_name || '',
        candidate_name: item.candidate_name || '',
        job_title: item.applied_for_role || '',
        applied_for_role: item.applied_for_role || '',
        restaurant_id: item.restaurant_id || '',
        current_status: item.current_status || '',
        email: item.email || '',
        phone: item.phone || '',
        notes: item.notes || ''
      };
    });
  },

  getUsersForSession_(session) {
    return filterUsersForSessionByRestaurantV318_(session, this.getUsers_());
  },

  getRestaurantMappingsForSession_(users, restaurants, session) {
    const scopeIds = getSessionRestaurantScopeV318_(session);
    const userLookup = {};
    const restaurantLookup = {};
    (users || []).forEach(function(user) { userLookup[user.userId] = user; });
    (restaurants || []).forEach(function(restaurant) { restaurantLookup[restaurant.restaurantId] = restaurant; });

    return Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS)
      .filter(function(mapping) {
        if (scopeIds === null) return true;
        return scopeIds.indexOf(cleanV318_(mapping.restaurant_id)) !== -1;
      })
      .map(function(mapping) {
        const user = userLookup[mapping.user_id] || {};
        const restaurant = restaurantLookup[mapping.restaurant_id] || {};
        return {
          accessId: mapping.access_id,
          userId: mapping.user_id || '',
          username: user.username || '',
          fullName: user.fullName || '',
          restaurantId: mapping.restaurant_id || '',
          restaurantCode: restaurant.restaurantCode || '',
          restaurantName: restaurant.restaurantName || '',
          accessLevel: mapping.access_level || '',
          isPrimary: String(mapping.is_primary || 'FALSE'),
          createdAt: mapping.created_at || ''
        };
      });
  },

  getHostLinks_() {
    return Repository.getAll_(APP_CONFIG.sheets.HOST_LINKS).map(function(link) {
      return {
        link_id: link.link_id,
        linkId: link.link_id,
        category_code: link.category_code || '',
        title: link.title || '',
        url: link.url || '',
        description: link.description || '',
        target_role_codes: link.target_role_codes || '',
        target_restaurant_ids: link.target_restaurant_ids || '',
        status: link.status || 'ACTIVE',
        start_at: link.start_at || '',
        end_at: link.end_at || '',
        sort_order: link.sort_order || '',
        created_by: link.created_by || '',
        updated_at: link.updated_at || ''
      };
    });
  },

  getEmployees_() {
    return Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS).map(function(item) {
      return {
        teamMemberId: item.team_member_id,
        full_name: item.full_name || '',
        fullName: item.full_name || '',
        job_title: item.job_title || '',
        restaurant_id: item.restaurant_id || '',
        restaurantId: item.restaurant_id || '',
        status: item.status || 'ACTIVE',
        email: item.email || '',
        phone: item.phone || '',
        employee_id: item.employee_id || '',
        notes: item.notes || ''
      };
    });
  },

  getNewHires_() {
    return Repository.getAll_(APP_CONFIG.sheets.NEW_HIRES).map(function(item) {
      return {
        hireId: item.hire_id,
        hire_id: item.hire_id,
        full_name: item.candidate_name || '',
        candidate_name: item.candidate_name || '',
        job_title: item.applied_for_role || '',
        applied_for_role: item.applied_for_role || '',
        restaurant_id: item.restaurant_id || '',
        current_status: item.current_status || '',
        email: item.email || '',
        phone: item.phone || '',
        notes: item.notes || ''
      };
    });
  },

  getNewsletters_() {
    return Repository.getAll_(APP_CONFIG.sheets.NEWSLETTERS).map(function(item) {
      return {
        postId: item.post_id,
        title: item.title || '',
        message: item.body || '',
        audienceRoles: item.audience_roles || '',
        audienceRestaurants: item.audience_restaurants || '',
        startAt: item.start_at || '',
        endAt: item.end_at || '',
        status: item.status || 'ACTIVE'
      };
    });
  },

  getButtonPermissionMatrix_() {
    const records = Repository.getAll_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS);
    const matrix = {};
    Object.keys(APP_CONFIG.roles).forEach(function(key) {
      const roleCode = APP_CONFIG.roles[key];
      const roleRecords = records.filter(function(record) {
        return String(record.role_code || '') === String(roleCode);
      });

      // Super Admin always has all role action buttons.
      // It should not depend on saved rows in Role Button Permissions.
      if (String(roleCode) === String(APP_CONFIG.roles.SUPER_ADMIN)) {
        matrix[roleCode] = RoleActionService.normalizeActionKeysV367_(ALL_ROLE_ACTION_BUTTONS || []);
        return;
      }

      if (!roleRecords.length) {
        matrix[roleCode] = RoleActionService.normalizeActionKeysV367_(
          DEFAULT_ROLE_ACTION_PERMISSIONS[roleCode] || []
        );
        return;
      }

      matrix[roleCode] = RoleActionService.applyLegacyPayrollAdminBackfillV367_(roleCode, roleRecords
        .filter(function(record) {
          return String(record.is_allowed).toUpperCase() === 'TRUE' &&
            String(record.button_key || '').trim() !== ROLE_ACTION_NO_ACCESS_MARKER;
        })
        .map(function(record) {
          return String(record.button_key || '').trim();
        }));
    });
    return matrix;
  },

  getUsers_() {
    const users = Repository.getAll_(APP_CONFIG.sheets.USERS);
    const accessRows = Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS);
    const accessCountByUser = {};

    accessRows.forEach(function(accessRow) {
      const userId = String(accessRow.user_id || '');
      if (!userId) return;
      accessCountByUser[userId] = (accessCountByUser[userId] || 0) + 1;
    });

    return users.map(function(user) {
      return {
        userId: user.user_id,
        username: user.username,
        fullName: user.full_name || '',
        email: user.email || '',
        phone: user.phone || '',
        roleCode: user.role_code || '',
        employeeId: user.employee_id || '',
        restaurantId: user.restaurant_id || '',
        region: user.region || '',
        status: user.status || 'ACTIVE',
        lastLoginAt: user.last_login_at || '',
        mustResetPassword: String(user.must_reset_password || 'FALSE'),
        assignedRestaurantCount: accessCountByUser[user.user_id] || 0,
        welcomeEmailPending: true
      };
    });
  },

  getRestaurantMappings_(users, restaurants) {
    const userLookup = {};
    const restaurantLookup = {};
    (users || []).forEach(function(user) { userLookup[user.userId] = user; });
    (restaurants || []).forEach(function(restaurant) { restaurantLookup[restaurant.restaurantId] = restaurant; });

    return Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS).map(function(mapping) {
      const user = userLookup[mapping.user_id] || {};
      const restaurant = restaurantLookup[mapping.restaurant_id] || {};
      return {
        accessId: mapping.access_id,
        userId: mapping.user_id || '',
        username: user.username || '',
        fullName: user.fullName || '',
        restaurantId: mapping.restaurant_id || '',
        restaurantCode: restaurant.restaurantCode || '',
        restaurantName: restaurant.restaurantName || '',
        accessLevel: mapping.access_level || '',
        isPrimary: String(mapping.is_primary || 'FALSE'),
        createdAt: mapping.created_at || ''
      };
    });
  },

  saveRestaurant(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);

    const restaurantNumber = String(payload.restaurantNumber || payload.restaurantCode || payload.restaurant_number || payload.restaurant_code || '').trim();
    const restaurantId = String(payload.restaurantId || payload.restaurant_id || '').trim() || (restaurantNumber ? 'REST_' + restaurantNumber : Utilities.getUuid());
    const restaurantName = String(payload.restaurantName || payload.restaurant_name || '').trim();
    const brandName = String(payload.brandName || payload.brand || payload.brand_name || '').trim();
    const now = new Date().toISOString();

    const record = {
      restaurant_id: restaurantId,
      restaurant_number: restaurantNumber,
      restaurant_code: restaurantNumber,
      restaurant_name: restaurantName,
      brand: brandName,
      brand_name: brandName,
      region: String(payload.region || '').trim(),
      city: String(payload.city || '').trim(),
      province_state: String(payload.provinceState || '').trim(),
      country: String(payload.country || '').trim(),
      status: String(payload.status || 'ACTIVE').trim() || 'ACTIVE',
      opening_date: String(payload.openingDate || '').trim(),
      gm_name: String(payload.gmName || '').trim(),
      gm_email: String(payload.gmEmail || '').trim(),
      mro_owner_user_id: String(payload.mroOwnerUserId || '').trim(),
      notes: String(payload.notes || '').trim(),
      updated_at: now
    };

    if (!record.restaurant_number || !record.restaurant_name || !brandName || !record.region) {
      throw new Error('Restaurant number, restaurant name, brand, and region are required.');
    }

    const existing = Repository.findOneBy_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', restaurantId);
    if (!existing) {
      record.created_at = now;
      Repository.append_(APP_CONFIG.sheets.RESTAURANTS, record);
      logAuditEvent_('RESTAURANT_CREATED', 'RESTAURANT', restaurantId, session.userId, { restaurantNumber: record.restaurant_number });
    } else {
      Repository.updateById_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', restaurantId, record);
      logAuditEvent_('RESTAURANT_UPDATED', 'RESTAURANT', restaurantId, session.userId, { restaurantNumber: record.restaurant_number });
    }

    return respond_(true, {
      restaurantId: restaurantId,
      restaurantNumber: restaurantNumber,
      restaurantCode: restaurantNumber
    });
  },

  deleteRestaurant(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
    const restaurantId = String(payload.restaurantId || '').trim();
    if (!restaurantId) throw new Error('Restaurant ID is required.');

    Repository.updateById_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', restaurantId, {
      status: 'INACTIVE',
      updated_at: new Date().toISOString()
    });
    logAuditEvent_('RESTAURANT_DEACTIVATED', 'RESTAURANT', restaurantId, session.userId, {});
    return respond_(true, { restaurantId: restaurantId });
  },

  createUser(payload) {
    return this.saveUser(payload);
  },

  saveUser(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertUserManagementAccessV310_(session);

    const userId = String(payload.userId || '').trim() || Utilities.getUuid();
    const username = String(payload.username || '').trim();
    const password = String(payload.password || '');
    const fullName = String(payload.fullName || '').trim();
    const email = String(payload.email || '').trim();
    const roleCode = String(payload.roleCode || '').trim();
    const region = String(payload.region || '').trim();
    const employeeId = String(payload.employeeId || '').trim();
    const phone = String(payload.phone || '').trim();
    const status = String(payload.status || 'ACTIVE').trim() || 'ACTIVE';
    assertSafeUserRoleChangeV310_(session, roleCode);
    const assignedRestaurants = (Array.isArray(payload.assignedRestaurantIds) ? payload.assignedRestaurantIds : String(payload.assignedRestaurantIds || '').split(','))
      .map(function(value) { return String(value || '').trim(); })
      .filter(Boolean);
    const primaryRestaurantId = assignedRestaurants[0] || '';

    if (!username || !fullName || !roleCode) {
      throw new Error('Full name, username, and role are required.');
    }

    if (Object.keys(APP_CONFIG.roles).map(function(key) { return APP_CONFIG.roles[key]; }).indexOf(roleCode) === -1) {
      throw new Error('Invalid role selected.');
    }

    const existing = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'user_id', userId);
    const duplicateUsername = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'username', username);
    if (duplicateUsername && String(duplicateUsername.user_id) !== userId) {
      throw new Error('Username already exists.');
    }

    if (email) {
      const duplicateEmail = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'email', email);
      if (duplicateEmail && String(duplicateEmail.user_id) !== userId) {
        throw new Error('Email already exists.');
      }
    }

    if ((roleCode === APP_CONFIG.roles.RESTAURANT_USER || roleCode === APP_CONFIG.roles.MRO) && !assignedRestaurants.length) {
      throw new Error('At least one assigned restaurant is required for this role.');
    }

    // FIX #14: Validate each assigned restaurant exists before writing.
    for (var _ri = 0; _ri < assignedRestaurants.length; _ri++) {
      var _rid = assignedRestaurants[_ri];
      if (!Repository.findOneBy_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', _rid)) {
        throw new Error('Restaurant not found: ' + _rid);
      }
    }

    const now = new Date().toISOString();
    const record = {
      user_id: userId,
      username: username,
      full_name: fullName,
      email: email,
      phone: phone,
      role_code: roleCode,
      employee_id: employeeId,
      restaurant_id: primaryRestaurantId,
      region: region,
      status: status,
      updated_at: now,
      updated_by: session.userId
    };

    if (!existing) {
      if (password.length < 8) throw new Error('Password must be at least 8 characters.');
      const salt = Utilities.getUuid().replace(/-/g, '');
      record.password_hash = hashPassword_(password, salt);
      record.salt = salt;
      record.last_login_at = '';
      record.must_reset_password = 'TRUE';
      record.remember_token = '';
      record.created_at = now;
      record.created_by = session.userId;
      Repository.append_(APP_CONFIG.sheets.USERS, record);
      logAuditEvent_('USER_CREATED', 'USER', userId, session.userId, { username: username, roleCode: roleCode });
    } else {
      Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', userId, record);
      logAuditEvent_('USER_UPDATED', 'USER', userId, session.userId, { username: username, roleCode: roleCode });
    }

    const restaurantIds = uniqueValues_(assignedRestaurants);
    Repository.replaceRowsByColumn_(APP_CONFIG.sheets.RESTAURANT_ACCESS, 'user_id', userId, restaurantIds.map(function(restaurantId, index) {
      return {
        access_id: Utilities.getUuid(),
        user_id: userId,
        restaurant_id: restaurantId,
        access_level: index === 0 ? 'PRIMARY' : 'ASSIGNED',
        is_primary: index === 0 ? 'TRUE' : 'FALSE',
        created_at: now,
        created_by: session.userId
      };
    }));

    return respond_(true, {
      userId: userId,
      username: username,
      roleCode: roleCode
    });
  },

  deleteUser(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertUserManagementAccessV310_(session);
    const userId = String(payload.userId || '').trim();
    if (!userId) throw new Error('User ID is required.');
    const targetUser = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'user_id', userId);
    if (targetUser) assertSafeUserRoleChangeV310_(session, targetUser.role_code);

    Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', userId, {
      status: 'INACTIVE',
      updated_at: new Date().toISOString(),
      updated_by: session.userId
    });
    logAuditEvent_('USER_DEACTIVATED', 'USER', userId, session.userId, {});
    return respond_(true, { userId: userId });
  },

  resetUserPassword(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertUserManagementAccessV310_(session);
    const userId = String(payload.userId || '').trim();
    const password = String(payload.password || '');
    if (!userId || password.length < 8) throw new Error('User ID and a minimum 8 character password are required.');
    const targetUser = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'user_id', userId);
    if (targetUser) assertSafeUserRoleChangeV310_(session, targetUser.role_code);

    const salt = Utilities.getUuid().replace(/-/g, '');
    Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', userId, {
      password_hash: hashPassword_(password, salt),
      salt: salt,
      must_reset_password: 'TRUE',
      updated_at: new Date().toISOString(),
      updated_by: session.userId
    });
    logAuditEvent_('PASSWORD_RESET', 'USER', userId, session.userId, {});
    invalidateAllUserSessions_(userId, null);
    return respond_(true, { userId: userId });
  },

  saveSectionRecord(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);

    const config = getAdminSectionConfig_(payload.section);
    assertAuthorized_(session, getAllowedRolesForSection_(payload.section));
    const record = payload.record || {};
    const idValue = String(record[config.idField] || '').trim() || Utilities.getUuid();
    const now = new Date().toISOString();
    const normalized = {};

    config.fields.forEach(function(field) {
      normalized[field] = record[field] !== undefined ? record[field] : '';
    });
    normalized[config.idField] = idValue;
    if (config.fields.indexOf('updated_at') !== -1) normalized.updated_at = now;
    if (config.fields.indexOf('created_at') !== -1 && !record.created_at) normalized.created_at = now;

    const existing = Repository.findOneBy_(config.sheetName, config.idField, idValue);
    if (existing) {
      Repository.updateById_(config.sheetName, config.idField, idValue, normalized);
      logAuditEvent_(config.auditUpdateEvent, config.entityType, idValue, session.userId, {});
    } else {
      Repository.append_(config.sheetName, normalized);
      logAuditEvent_(config.auditCreateEvent, config.entityType, idValue, session.userId, {});
    }

    if (payload.section === 'NEW_HIRES') {
      const currentStatus = String(normalized.current_status || '').toUpperCase();
      if (currentStatus === 'STARTED WORKING' || currentStatus === 'STARTED_WORK' || currentStatus === 'STARTED WORK') {
        const employeeRecord = {
          team_member_id: Utilities.getUuid(),
          employee_id: normalized.employee_id || '',
          full_name: normalized.candidate_name || normalized.full_name || '',
          restaurant_id: normalized.restaurant_id || '',
          job_title: normalized.applied_for_role || normalized.job_title || '',
          department: '',
          employment_type: normalized.employment_type || '',
          hire_date: normalized.start_date || new Date().toISOString().slice(0, 10),
          status: 'ACTIVE',
          manager_name: '',
          email: normalized.email || '',
          phone: normalized.phone || '',
          created_at: now,
          updated_at: now
        };
        Repository.append_(APP_CONFIG.sheets.TEAM_MEMBERS, employeeRecord);
      }
    }

    return respond_(true, { id: idValue });
  },

  deleteSectionRecord(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    const config = getAdminSectionConfig_(payload.section);
    assertAuthorized_(session, getAllowedRolesForSection_(payload.section));
    const idValue = String(payload.id || '').trim();
    if (!idValue) throw new Error('Record ID is required.');

    if (config.softDeleteField) {
      const patch = {};
      patch[config.softDeleteField] = 'INACTIVE';
      if (config.fields.indexOf('updated_at') !== -1) patch.updated_at = new Date().toISOString();
      Repository.updateById_(config.sheetName, config.idField, idValue, patch);
    } else {
      Repository.deleteById_(config.sheetName, config.idField, idValue);
    }

    logAuditEvent_(config.auditDeleteEvent, config.entityType, idValue, session.userId, {});
    return respond_(true, { id: idValue });
  },

  importSectionRows(payload) {
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, getAllowedRolesForSection_(payload.section));
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const section = String(payload.section || '');
    if (section === 'RESTAURANTS') {
      rows.forEach(function(record) {
        AdminService.saveRestaurant({
          sessionToken: payload.sessionToken,
          restaurantId: record.restaurant_id || record.restaurantId || '',
          restaurantCode: record.restaurant_code || record.restaurantCode || '',
          restaurantName: record.restaurant_name || record.restaurantName || '',
          brandName: record.brand_name || record.brandName || '',
          region: record.region || '',
          city: record.city || '',
          provinceState: record.province_state || record.provinceState || '',
          country: record.country || '',
          status: record.status || 'ACTIVE',
          openingDate: record.opening_date || record.openingDate || '',
          gmName: record.gm_name || record.gmName || '',
          gmEmail: record.gm_email || record.gmEmail || '',
          mroOwnerUserId: record.mro_owner_user_id || record.mroOwnerUserId || '',
          notes: record.notes || ''
        });
      });
      return respond_(true, { rowCount: rows.length });
    }

    if (section === 'USERS' || section === 'HR_ADMINS') {
      rows.forEach(function(record) {
        AdminService.saveUser({
          sessionToken: payload.sessionToken,
          userId: record.user_id || record.userId || '',
          username: record.username || '',
          fullName: record.full_name || record.fullName || '',
          email: record.email || '',
          phone: record.phone || '',
          roleCode: record.role_code || record.roleCode || (section === 'HR_ADMINS' ? APP_CONFIG.roles.HR_ADMIN : ''),
          employeeId: record.employee_id || record.employeeId || '',
          restaurantId: record.restaurant_id || record.restaurantId || '',
          region: record.region || '',
          status: record.status || 'ACTIVE',
          assignedRestaurantIds: String(record.assigned_restaurant_ids || record.assignedRestaurantIds || ''),
          // FIX #22: Never fall back to a hardcoded default password.
          // If no password is supplied in the import row, throw so the admin
          // is forced to provide one explicitly.
          password: (function() {
            const pw = String(record.password || record.temporary_password || record.temporaryPassword || '').trim();
            if (!pw) throw new Error('Import row for "' + (record.username || 'unknown') + '" is missing a password. Please include a temporary_password column in your import data.');
            return pw;
          })()
        });
      });
      return respond_(true, { rowCount: rows.length });
    }

    rows.forEach(function(record) {
      AdminService.saveSectionRecord({
        sessionToken: payload.sessionToken,
        section: section,
        record: record
      });
    });
    logAuditEvent_('SECTION_IMPORT', 'SECTION', section, session.userId, { rowCount: rows.length });
    return respond_(true, { rowCount: rows.length });
  }
};

function createPortalUser(payload) {
  return AdminService.createUser(payload);
}

function savePortalUser(payload) {
  return AdminService.saveUser(payload);
}

function deletePortalUser(payload) {
  return AdminService.deleteUser(payload);
}

function resetPortalUserPassword(payload) {
  return AdminService.resetUserPassword(payload);
}

function saveRestaurantRecord(payload) {
  return AdminService.saveRestaurant(payload);
}

function deleteRestaurantRecord(payload) {
  return AdminService.deleteRestaurant(payload);
}

function getAdminWorkspacePayload(token) {
  const session = getSessionOrThrow_(token);
  return respond_(true, { adminTools: getManagementWorkspaceBoot_(session) });
}

function saveAdminSectionRecord(payload) {
  return AdminService.saveSectionRecord(payload);
}

function deleteAdminSectionRecord(payload) {
  return AdminService.deleteSectionRecord(payload);
}

function importAdminSectionRows(payload) {
  return AdminService.importSectionRows(payload);
}

function submitContactRequest(payload) {
  return ContactRequestService.create(payload);
}

function updateContactRequest(payload) {
  return ContactRequestService.update(payload);
}

function submitPasswordChangeRequest(payload) {
  return PasswordChangeService.create(payload);
}

function fulfillPasswordChangeRequest(payload) {
  return PasswordChangeService.fulfill(payload);
}

function saveRoleActionPermissions(payload) {
  return RoleActionService.savePermissions(payload);
}

function getAdminSectionConfig_(section) {
  const map = {
    HOST_LINKS: {
      sheetName: APP_CONFIG.sheets.HOST_LINKS,
      idField: 'link_id',
      fields: getSheetSchemas_()['HOST Links'],
      entityType: 'HOST_LINK',
      auditCreateEvent: 'HOST_LINK_CREATED',
      auditUpdateEvent: 'HOST_LINK_UPDATED',
      auditDeleteEvent: 'HOST_LINK_DELETED',
      softDeleteField: 'status'
    },
    BRANDS: {
      sheetName: APP_CONFIG.sheets.BRANDS,
      idField: 'brand_id',
      fields: getSheetSchemas_()['Brands'],
      entityType: 'BRAND',
      auditCreateEvent: 'BRAND_CREATED',
      auditUpdateEvent: 'BRAND_UPDATED',
      auditDeleteEvent: 'BRAND_DELETED',
      softDeleteField: 'status'
    },
    TEAM_MEMBERS: {
      sheetName: APP_CONFIG.sheets.TEAM_MEMBERS,
      idField: 'team_member_id',
      fields: getSheetSchemas_()['Team Members'],
      entityType: 'TEAM_MEMBER',
      auditCreateEvent: 'TEAM_MEMBER_CREATED',
      auditUpdateEvent: 'TEAM_MEMBER_UPDATED',
      auditDeleteEvent: 'TEAM_MEMBER_DELETED',
      softDeleteField: 'status'
    },
    NEW_HIRES: {
      sheetName: APP_CONFIG.sheets.NEW_HIRES,
      idField: 'hire_id',
      fields: getSheetSchemas_()['New Hires'],
      entityType: 'NEW_HIRE',
      auditCreateEvent: 'NEW_HIRE_CREATED',
      auditUpdateEvent: 'NEW_HIRE_UPDATED',
      auditDeleteEvent: 'NEW_HIRE_DELETED',
      softDeleteField: ''
    },
    NEWSLETTERS: {
      sheetName: APP_CONFIG.sheets.NEWSLETTERS,
      idField: 'post_id',
      fields: getSheetSchemas_()['Newsletters'],
      entityType: 'NEWSLETTER',
      auditCreateEvent: 'NEWSLETTER_CREATED',
      auditUpdateEvent: 'NEWSLETTER_UPDATED',
      auditDeleteEvent: 'NEWSLETTER_DELETED',
      softDeleteField: 'status'
    },
    CONTACT_REQUESTS: {
      sheetName: APP_CONFIG.sheets.CONTACT_REQUESTS,
      idField: 'contact_request_id',
      fields: getSheetSchemas_()['Contact Requests'],
      entityType: 'CONTACT_REQUEST',
      auditCreateEvent: 'CONTACT_REQUEST_CREATED',
      auditUpdateEvent: 'CONTACT_REQUEST_UPDATED',
      auditDeleteEvent: 'CONTACT_REQUEST_DELETED',
      softDeleteField: 'status'
    },
    PASSWORD_CHANGE_REQUESTS: {
      sheetName: APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS,
      idField: 'password_request_id',
      fields: getSheetSchemas_()['Password Change Requests'],
      entityType: 'PASSWORD_REQUEST',
      auditCreateEvent: 'PASSWORD_REQUEST_CREATED',
      auditUpdateEvent: 'PASSWORD_REQUEST_UPDATED',
      auditDeleteEvent: 'PASSWORD_REQUEST_DELETED',
      softDeleteField: 'status'
    },
    DROPDOWN_OPTIONS: {
      sheetName: APP_CONFIG.sheets.DROPDOWN_OPTIONS,
      idField: 'option_id',
      fields: getSheetSchemas_()['Dropdown Options'],
      entityType: 'DROPDOWN_OPTION',
      auditCreateEvent: 'DROPDOWN_OPTION_CREATED',
      auditUpdateEvent: 'DROPDOWN_OPTION_UPDATED',
      auditDeleteEvent: 'DROPDOWN_OPTION_DELETED',
      softDeleteField: 'status'
    },
    RESTAURANT_ACCESS: {
      sheetName: APP_CONFIG.sheets.RESTAURANT_ACCESS,
      idField: 'access_id',
      fields: getSheetSchemas_()['Restaurant Access'],
      entityType: 'RESTAURANT_ACCESS',
      auditCreateEvent: 'RESTAURANT_ACCESS_CREATED',
      auditUpdateEvent: 'RESTAURANT_ACCESS_UPDATED',
      auditDeleteEvent: 'RESTAURANT_ACCESS_DELETED',
      softDeleteField: ''
    }
  };

  if (!map[section]) throw new Error('Unknown admin section.');
  return map[section];
}

function getAllowedRolesForSection_(section) {
  const map = {
    WEBSITE_CONTROL: [APP_CONFIG.roles.SUPER_ADMIN],
    HOST_LINKS: [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT],
    NEWSLETTERS: [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO],
    NEW_HIRES: [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN],
    REPORTS: [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.SENIOR_MANAGEMENT],
    ASSIGNED_REPORTS: [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.MRO]
  };
  return map[section] || [APP_CONFIG.roles.SUPER_ADMIN];
}

function canManageWorkspace_(session) {
  return getAllowedRolesForSection_('HOST_LINKS').indexOf(session.roleCode) !== -1;
}

function getManagementWorkspaceBoot_(session) {
  // FIX v500 #3: Cache per-user workspace boot for 120s.
  // Busted by bustBootCacheForToken_ on user/restaurant changes.
  const _v500wsKey = 'v500_ws_boot_' + String(session.roleCode || '') + '_' + String(session.userId || '').slice(-8);
  try {
    const _v500c = CacheService.getScriptCache().get(_v500wsKey);
    if (_v500c) return JSON.parse(_v500c);
  } catch (_v500e) {}

  if (session.roleCode === APP_CONFIG.roles.SUPER_ADMIN) {
    const _v500r = AdminService.getUserManagementBoot_(session);
    try {
      const _v500j = JSON.stringify(_v500r);
      if (_v500j.length < 200000) CacheService.getScriptCache().put(_v500wsKey, _v500j, 120);
    } catch (_v500e) {}
    return _v500r;
  }

  const canManageHostLinks = getAllowedRolesForSection_('HOST_LINKS').indexOf(session.roleCode) !== -1;
  const canManageNewsletters = getAllowedRolesForSection_('NEWSLETTERS').indexOf(session.roleCode) !== -1;
  const canManageNewHires = getAllowedRolesForSection_('NEW_HIRES').indexOf(session.roleCode) !== -1;
  const canViewReports = getAllowedRolesForSection_('REPORTS').indexOf(session.roleCode) !== -1;
  const canViewAssignedReports = getAllowedRolesForSection_('ASSIGNED_REPORTS').indexOf(session.roleCode) !== -1;
  const canManageUsers = hasRoleActionPermissionV310_(session, 'USER_MANAGEMENT');

  if (!canManageHostLinks && !canManageNewsletters && !canManageNewHires && !canViewReports && !canViewAssignedReports && !canManageUsers) {
    throw new Error('You are not authorized to open this workspace.');
  }

  const hostLinks = canManageHostLinks ? AdminService.getHostLinks_() : [];
  const restaurants = AdminService.getRestaurantsForSession_(session);
  const newsletters = canManageNewsletters ? AdminService.getNewsletters_() : [];
  const newHires = canManageNewHires ? AdminService.getNewHiresForSession_(session) : [];
  const reportRows = canViewReports ? buildMetricRows_(DashboardService.getTopMetrics(session), 'Enterprise reporting overview') : [];
  const assignedReportRows = canViewAssignedReports ? buildMetricRows_(DashboardService.getTopMetrics(session), 'Assigned restaurant reporting overview') : [];
  const users = canManageUsers ? AdminService.getUsersForSession_(session) : [];
  const restaurantMappings = canManageUsers ? AdminService.getRestaurantMappingsForSession_(users, restaurants, session) : [];

  const allowedSections = [];
  if (canManageHostLinks) allowedSections.push('HOST_LINKS');
  if (canManageNewsletters) allowedSections.push('NEWSLETTERS');
  if (canManageNewHires) allowedSections.push('NEW_HIRES');
  if (canViewReports) allowedSections.push('REPORTS');
  if (canViewAssignedReports) allowedSections.push('ASSIGNED_REPORTS');
  if (canManageUsers) allowedSections.push('USERS');

  const activeUsers = users.filter(function(user) { return String(user.status || '').toUpperCase() === 'ACTIVE'; }).length;
  const inactiveUsers = users.filter(function(user) { return String(user.status || '').toUpperCase() === 'INACTIVE'; }).length;

  const stats = {
    hostLinkCount: hostLinks.length,
    activeHostLinks: hostLinks.filter(function(item) { return String(item.status || '').toUpperCase() === 'ACTIVE'; }).length,
    hospitalityLinks: hostLinks.filter(function(item) { return String(item.category_code || '').toUpperCase() === 'H'; }).length,
    systemsLinks: hostLinks.filter(function(item) { return String(item.category_code || '').toUpperCase() === 'S'; }).length,
    totalUsers: users.length,
    activeUsers: activeUsers,
    inactiveUsers: inactiveUsers
  };

  const _v500result = {
    roles: Object.keys(APP_CONFIG.roles).map(function(key) {
      return APP_CONFIG.roles[key];
    }),
    restaurants: restaurants,
    hostLinks: hostLinks,
    newsletters: newsletters,
    newHires: newHires,
    users: users,
    restaurantMappings: restaurantMappings,
    reportRows: reportRows,
    assignedReportRows: assignedReportRows,
    allowedSections: allowedSections,
    stats: stats
  };
  // FIX v500 #3: store workspace boot in ScriptCache
  try {
    const _v500j = JSON.stringify(_v500result);
    if (_v500j.length < 200000) CacheService.getScriptCache().put(_v500wsKey, _v500j, 120);
  } catch (_v500e) {}
  return _v500result;
}

function buildMetricRows_(metrics, details) {
  return Object.keys(metrics || {}).map(function(key) {
    return {
      item: formatLabel_(key),
      value: metrics[key],
      details: details || 'Live summary metric',
      status: 'ACTIVE'
    };
  });
}

function buildWebsiteControlRows_(stats) {
  return [
    { item: 'Portal Status', value: 'ONLINE', details: 'Core portal services are available.', status: 'ACTIVE' },
    { item: 'Restaurants', value: stats.totalRestaurants || 0, details: 'Stores mapped into the portal database.', status: 'ACTIVE' },
    { item: 'Users', value: stats.totalUsers || 0, details: 'Total portal accounts across all roles.', status: 'ACTIVE' },
    { item: 'Active Users', value: stats.activeUsers || 0, details: 'Users currently marked active.', status: 'ACTIVE' },
    { item: 'Inactive Users', value: stats.inactiveUsers || 0, details: 'Users currently marked inactive.', status: 'INACTIVE' },
    { item: 'HOST Links', value: stats.hostLinkCount || 0, details: 'Quick-access links configured across H/O/S/T.', status: 'ACTIVE' },
    { item: 'Newsletters', value: stats.newsletterCount || 0, details: 'Newsletter posts available in the system.', status: 'ACTIVE' },
    { item: 'Contact Requests', value: stats.contactRequestCount || 0, details: 'Support requests submitted by users.', status: 'ACTIVE' },
    { item: 'Password Requests', value: stats.passwordRequestCount || 0, details: 'Pending or completed password help requests.', status: 'ACTIVE' }
  ];
}

var HireService = {
  updateStatus(hireId, newStatus, note, sessionToken) {
    const session = getSessionOrThrow_(sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.SUPER_ADMIN]);
    if (APP_CONFIG.hireStatuses.indexOf(newStatus) === -1) {
      throw new Error('Invalid hire status.');
    }

    Repository.updateById_(APP_CONFIG.sheets.NEW_HIRES, 'hire_id', hireId, {
      current_status: newStatus,
      notes: note || '',
      updated_at: new Date().toISOString()
    });
    logAuditEvent_('HIRE_STATUS_UPDATED', 'HIRE', hireId, session.userId, { newStatus: newStatus });
    return respond_(true, { hireId: hireId, status: newStatus });
  }
};

function resolveRequestSheet_(requestGroup) {
  const map = {
    HR: APP_CONFIG.sheets.REQUESTS_HR,
    ADMIN: APP_CONFIG.sheets.REQUESTS_ADMIN,
    SUPPLIES: APP_CONFIG.sheets.REQUESTS_SUPPLIES,
    MAINTENANCE: APP_CONFIG.sheets.REQUESTS_MAINTENANCE
  };
  if (!map[requestGroup]) throw new Error('Unknown request group.');
  return map[requestGroup];
}
var ServiceRouter_ = {
  route: function(action, payload, token) {
    // FIX #1: Pre-dispatch authentication guard.
    // Routes listed in PAYLOAD_AUTH_ROUTES embed the session token inside payload
    // and validate it themselves. All other routes MUST supply a valid token at
    // the router level — this provides defense in depth so any future route that
    // forgets its own auth check is blocked here rather than silently open.
    const PAYLOAD_AUTH_ROUTES = {
      submitContactRequest: true,
      submitPasswordChangeRequest: true,
      createRequest: true,
      getOperationsDashboardData: true,
      submitPayrollOnBehalf: true,
      recallPayrollSubmission: true,
      exportPayrollSummary: true,
      getMROPayrollDashboardData: true,
      queryPayrollAdminWorkspaceData: true,
      exportPayrollAdminCsv: true,
      returnPayrollSubmissionToRestaurant: true
    };
    if (!PAYLOAD_AUTH_ROUTES[action]) {
      if (!token) throw new Error('Authentication required.');
      getSessionOrThrow_(token); // throws immediately on invalid/expired session
    }
    const routes = {
      getRecentHiringRequests: function() { return getRecentHiringRequests(token); },
      createRequest: function() { return createRequest(payload); },
      getManageTeamData: function() { return getManageTeamData(token); },
      saveManageTeamEmployee: function() { return saveManageTeamEmployee(payload); },
      updateManageTeamStatus: function() { return updateManageTeamStatus(payload); },
      saveTeamMemberFromManageTeam: function() { return saveTeamMemberFromManageTeam(payload); },
      updateTeamMemberStatusFromManageTeam: function() { return updateTeamMemberStatusFromManageTeam(payload); },
      updateManageTeamEmployeeDetails: function() { return updateManageTeamEmployeeDetails(payload); },
      deactivateManageTeamEmployee: function() { return deactivateManageTeamEmployee(payload); },
      getNewHeroEmployees: function() { return getNewHeroEmployees(token); },
      updateNewHeroEmployeeStatus: function() { return updateNewHeroEmployeeStatus(payload); },
      saveNewHeroEmployee: function() { return saveNewHeroEmployee(payload); },
      updateNewHeroEmployee: function() { return updateNewHeroEmployee(payload); },
      deleteNewHeroEmployee: function() { return deleteNewHeroEmployee(payload); },
      createPortalUser: function() { return createPortalUser(payload); },
      savePortalUser: function() { return savePortalUser(payload); },
      deletePortalUser: function() { return deletePortalUser(payload); },
      resetPortalUserPassword: function() { return resetPortalUserPassword(payload); },
      saveRestaurantRecord: function() { return saveRestaurantRecord(payload); },
      deleteRestaurantRecord: function() { return deleteRestaurantRecord(payload); },
      getAdminWorkspacePayload: function() { return getAdminWorkspacePayload(token); },
      saveAdminSectionRecord: function() { return saveAdminSectionRecord(payload); },
      deleteAdminSectionRecord: function() { return deleteAdminSectionRecord(payload); },
      importAdminSectionRows: function() { return importAdminSectionRows(payload); },
      submitContactRequest: function() { return submitContactRequest(payload); },
      updateContactRequest: function() { return updateContactRequest(payload); },
      listContactRequests: function() { return listContactRequests(payload); },
      submitPasswordChangeRequest: function() { return submitPasswordChangeRequest(payload); },
      listPasswordChangeRequests: function() { return listPasswordChangeRequests(payload); },
      updatePasswordChangeRequest: function() { return updatePasswordChangeRequest(payload); },
      fulfillPasswordChangeRequest: function() { return fulfillPasswordChangeRequest(payload); },
      saveRoleActionPermissions: function() { return saveRoleActionPermissions(payload); },
      getNewHeroesData: function() { return getNewHeroesData(token); },
      updateNewHeroRecruitmentStatus: function() { return updateNewHeroRecruitmentStatus(payload); },
      updateNewHero: function() { return updateNewHero(token, payload); },
      updateFollowUpStatus: function() { return updateFollowUpStatus(token, payload); },
      getHiringRequestWorkspaceData: function() { return getHiringRequestWorkspaceData(token); },
      createHiringRequest: function() { return createHiringRequest(payload); },
      // FIX #19: Register getOperationsDashboardData so ClientJS dispatch() calls succeed.
      getOperationsDashboardData: function() { return getOperationsDashboardData(payload || token); },
      // PAYROLL V360 — new routes
      submitPayrollOnBehalf:      function() { return submitPayrollOnBehalf(payload); },
      recallPayrollSubmission:    function() { return recallPayrollSubmission(payload); },
      exportPayrollSummary:       function() { return exportPayrollSummary(payload); },
      getMROPayrollDashboardData: function() { return getMROPayrollDashboardData(payload); },
      queryPayrollAdminWorkspaceData: function() { return queryPayrollAdminWorkspaceData(payload); },
      exportPayrollAdminCsv:      function() { return exportPayrollAdminCsv(payload); },
      returnPayrollSubmissionToRestaurant: function() { return returnPayrollSubmissionToRestaurant(payload); },

      // ── PAYROLL routes missing from original table (FIX) ──────────────
      getPayrollSubmissionWorkspaceData: function() { return getPayrollSubmissionWorkspaceData(token); },
      getPayrollReviewWorkspaceData:     function() { return getPayrollReviewWorkspaceData(token); },
      savePayrollSubmissionDraft:        function() { return savePayrollSubmissionDraft(payload); },
      updatePayrollSubmissionStatus:     function() { return updatePayrollSubmissionStatus(payload); },
      getPayrollRoeApplicationData:      function() { return getPayrollRoeApplicationData(payload); },
      submitPayrollRoeApplication:       function() { return submitPayrollRoeApplication(payload); },
      updatePayrollRoeStatus:            function() { return updatePayrollRoeStatus(payload); },
      deactivateEmployeeFromRoe:         function() { return deactivateEmployeeFromRoe(payload); },
      updatePayrollAdminWebsiteStatus:   function() { return updatePayrollAdminWebsiteStatus(payload); },

      // ── Role permissions ──────────────────────────────────────────────
      getRoleButtonPermissions:          function() { return getRoleButtonPermissions(token); },

      // ── HOST Link management ─────────────────────────────────────────
      saveHostLinkRecord:                function() { return saveHostLinkRecord(payload); },
      listHostLinkRecords:               function() { return listHostLinkRecords(payload); },
      updateHostLinkStatus:              function() { return updateHostLinkStatus(payload); },

      // ── Newsletter management ────────────────────────────────────────
      saveNewsletterRecord:              function() { return saveNewsletterRecord(payload); },
      listNewsletterRecords:             function() { return listNewsletterRecords(payload); },
      updateNewsletterStatus:            function() { return updateNewsletterStatus(payload); },

      // ── Dropdown Settings ────────────────────────────────────────────
      saveDropdownOptionRecord:          function() { return saveDropdownOptionRecord(payload); },
      listDropdownOptionRecords:         function() { return listDropdownOptionRecords(payload); },
      updateDropdownOptionStatus:        function() { return updateDropdownOptionStatus(payload); },

      // ── Payroll Submission (direct submit, not via MRO) ─────────────
      submitPayrollSubmission:           function() { return submitPayrollSubmission(payload); },

      // ── Change own password ─────────────────────────────────────────
      changeOwnPortalPassword:           function() { return changeOwnPortalPassword(payload); },
      initializePortalSheetsForCurrentUser: function() { return initializePortalSheetsForCurrentUser(token); },
      setNewPasswordOnFirstLogin: function() { return setNewPasswordOnFirstLogin(payload); }
    };
    if (!routes[action]) throw new Error('Unknown action: ' + action);
    try {
      return routes[action]();
    } catch (error) {
      logServiceErrorV310_('ServiceRouter.' + action, error, { action: action });
      throw error;
    }
  }
};

function assertAuthorized_(session, allowedRoles) {
  // FIX #17: Normalize roleCode before checking so non-canonical variants
  // (e.g. "SUPER ADMIN" stored in the sheet) don't silently lock users out.
  var role = String(session.roleCode || '').trim().toUpperCase().replace(/\s+/g, '_');
  if (allowedRoles.indexOf(role) === -1) {
    throw new Error('You are not authorized to perform this action.');
  }
}

function respond_(success, data, error) {
  return {
    success: success,
    data: data === undefined ? null : toSerializable_(data),
    error: error || null,
    serverTime: new Date().toISOString()
  };
}

// sanitizeClientError_: strip internal stack traces and file paths before
// sending error messages to the browser. Prevents leaking implementation details.
// Was referenced throughout the codebase but never defined — fixes ReferenceError.
function registerSessionForUser_(userId, token) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'user_sess_idx_' + String(userId || '').slice(-20);
    var raw = cache.get(key);
    var tokens = raw ? JSON.parse(raw) : [];
    tokens.push(token);
    if (tokens.length > 20) tokens = tokens.slice(-20);
    cache.put(key, JSON.stringify(tokens), 1800);
  } catch(e) {}
}

function invalidateAllUserSessions_(userId, keepToken) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'user_sess_idx_' + String(userId || '').slice(-20);
    var raw = cache.get(key);
    if (!raw) return;
    var tokens = JSON.parse(raw);
    var _NS = String(APP_CONFIG.scriptId || '').slice(-8);
    tokens.forEach(function(t) {
      if (t !== keepToken) cache.remove('sess:' + _NS + ':' + t);
    });
    cache.remove(key);
  } catch(e) {}
}

function sanitizeClientError_(error) {
  if (!error) return 'An unexpected error occurred.';
  const message = String(error.message || error || '');
  // Strip Apps Script internal paths like "at Object.route (Services:2041:5)"
  const cleaned = message.replace(/\s+at\s+.+/g, '').trim();
  // Return the cleaned message, or a generic fallback if it's empty or looks internal
  if (!cleaned || cleaned.length > 400) return 'An unexpected error occurred.';
  return cleaned;
}

function toSerializable_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(function(item) {
      return toSerializable_(item);
    });
  }
  if (value && typeof value === 'object') {
    const normalized = {};
    Object.keys(value).forEach(function(key) {
      normalized[key] = toSerializable_(value[key]);
    });
    return normalized;
  }
  return value;
}

function uniqueValues_(values) {
  const seen = {};
  return values.filter(function(value) {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function safeGetSession_(token) {
  try {
    return token ? getSessionOrThrow_(token) : null;
  } catch (error) {
    return null;
  }
}

function validateRequestAccess_(session, payload) {
  const requestMode = String((payload && payload.requestMode) || '').trim().toUpperCase();
  if (!requestMode) return;

  // Security check: the superadmin recruitment flow must stay inaccessible
  // even if a non-superadmin crafts a direct browser request manually.
  if (requestMode === 'SUPERADMIN_RECRUITMENT') {
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
  }
}

function formatLabel_(key) {
  return String(key || '')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replaceAll('_', ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\b\w/g, function(character) {
      return character.toUpperCase();
    });
}

var RecruitmentService = {

  getNewHeroesData: function(token) {
  const session = getSessionOrThrow_(token);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.RESTAURANT_USER
  ]);

  const allHires = Repository.getAll_(APP_CONFIG.sheets.NEW_HIRES);
  const allTeam = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS);

  // FIX #4: Scope RESTAURANT_USER to their assigned restaurant(s)
  const isRestaurantUser = session.roleCode === APP_CONFIG.roles.RESTAURANT_USER;
  const scopeRestaurantIds = isRestaurantUser
    ? getRecruitmentRestaurantScopeIds_(session)
    : null; // null = all restaurants

  function inRestaurantScope(restaurantId) {
    if (!scopeRestaurantIds) return true;
    return scopeRestaurantIds.indexOf(String(restaurantId || '')) !== -1;
  }

  const now = new Date();
  const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  const active = [];
  const followUp = [];
  const history = [];

  allHires.filter(function(hire) {
    return inRestaurantScope(hire.restaurant_id || hire.restaurantId || '');
  }).forEach(function(hire) {
    const status = String(hire.current_status || '').trim();
    const updatedAt = hire.updated_at ? new Date(hire.updated_at) : null;

    if (status === 'STARTED_WORK' || status === 'Work Started' || status === 'Already Started') {
      history.push(hire);
      return;
    }

    if (status === 'Quit' || status === 'Never Show Up') return;

    const badge =
      String(hire.notes || '').toLowerCase().indexOf('ready to start') !== -1 ||
      String(hire.current_status || '').toLowerCase().indexOf('ready to start') !== -1 ||
      String(hire.restaurant_confirmation_at || '') !== '';

    const stuckOnWelcomeKit =
      status === 'Welcome Kit Done' &&
      updatedAt &&
      updatedAt < thirtyDaysAgo;

    if (badge && stuckOnWelcomeKit) {
      followUp.push(hire);
      return;
    }

    active.push(hire);
  });

  allTeam.filter(function(member) {
    return inRestaurantScope(member.restaurant_id || member.restaurantId || '');
  }).forEach(function(member) {
    const status = String(member.status || '').trim();
    const updatedAt = member.updated_at ? new Date(member.updated_at) : null;

    const item = {
      hire_id: member.team_member_id || '',
      candidate_name: member.full_name || '',
      phone: member.phone || '',
      applied_for_role: member.job_title || 'Team Member',
      restaurant_id: member.restaurant_id || '',
      current_status: status,
      notes: status,
      start_date: member.hire_date || '',
      updated_at: member.updated_at || '',
      created_at: member.created_at || ''
    };

    if (status === 'ACTIVE' || status === 'Work Started' || status === 'Already Started') {
      history.push(item);
      return;
    }

    if ((status === 'Ready to Work' || status === 'Ready to Start') && updatedAt && updatedAt < thirtyDaysAgo) {
      followUp.push(item);
      return;
    }

    if (status === 'Newly Hired' || status === 'Orientation Done' || status === 'Ready to Work' || status === 'Ready to Start' || status === 'Welcome Kit Done') {
      active.push(item);
    }
  });

  history.sort(function(a, b) {
    return new Date(b.updated_at || b.created_at || 0) - new Date(a.updated_at || a.created_at || 0);
  });

  return respond_(true, {
    active: active,
    followUp: followUp,
    history: history
  });
},

  updateNewHero(token, payload) {
    const session = getSessionOrThrow_(token);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.RESTAURANT_USER
    ]);

    const hireId = String(payload.hireId || '').trim();
    if (!hireId) throw new Error('Hire ID is required.');

    const now = new Date().toISOString();
    const patch = {
      updated_at: now
    };

    if (payload.mobileNumber !== undefined)
      patch.phone = String(payload.mobileNumber || '').trim();
    if (payload.availability !== undefined)
      patch.notes = String(payload.availability || '').trim();
    if (payload.position !== undefined)
      patch.applied_for_role = String(payload.position || '').trim();
    if (payload.currentStatus !== undefined)
      patch.current_status = String(payload.currentStatus || '').trim();
    if (payload.firstDayOfWork !== undefined)
      patch.start_date = String(payload.firstDayOfWork || '').trim();

    // If status is Work Started, make employee ACTIVE or create new one
if (patch.current_status === 'Work Started' || patch.current_status === 'Already Started') {

  const existingHire = Repository.findOneBy_(
    APP_CONFIG.sheets.NEW_HIRES,
    'hire_id',
    hireId
  );

  if (existingHire) {

    const existingTeam = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS)
      .find(function(row) {
        return String(row.full_name || '').trim().toLowerCase() ===
               String(existingHire.candidate_name || '').trim().toLowerCase();
      });

    if (existingTeam) {

      Repository.updateById_(
        APP_CONFIG.sheets.TEAM_MEMBERS,
        'team_member_id',
        existingTeam.team_member_id,
        {
          status: 'ACTIVE',
hire_date: existingTeam.hire_date || patch.start_date || '',
restaurant_id: existingHire.restaurant_id || existingTeam.restaurant_id,
phone: payload.mobileNumber || existingTeam.phone || '',
updated_at: now
        }
      );

    } else {

      const _fixedTmId = Utilities.getUuid();
      const teamRecord = {
        team_member_id: _fixedTmId,
        employee_id: _fixedTmId, // FIX #3b
        full_name: existingHire.candidate_name || '',
        restaurant_id: existingHire.restaurant_id || '',
        job_title: payload.position || existingHire.applied_for_role || '',
        department: '',
        employment_type: payload.availability || '',
        hire_date: payload.firstDayOfWork || existingHire.start_date || now.slice(0,10),
        status: 'ACTIVE',
        manager_name: '',
        email: '',
        phone: payload.mobileNumber || existingHire.phone || '',
        created_at: now,
        updated_at: now
      };

      Repository.append_(APP_CONFIG.sheets.TEAM_MEMBERS, teamRecord);
    }
  }
}

    Repository.updateById_(APP_CONFIG.sheets.NEW_HIRES, 'hire_id', hireId, patch);
    logAuditEvent_('NEW_HERO_UPDATED', 'NEW_HIRE', hireId, session.userId, { patch: patch });
    return respond_(true, { hireId: hireId });
  },

  updateFollowUpStatus(token, payload) {
    const session = getSessionOrThrow_(token);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.RESTAURANT_USER
    ]);

    const hireId = String(payload.hireId || '').trim();
    const newStatus = String(payload.status || '').trim();
    const allowed = ['Already Started', 'Quit', 'Never Show Up'];

    if (!hireId) throw new Error('Hire ID is required.');
    if (allowed.indexOf(newStatus) === -1) throw new Error('Invalid status.');

    Repository.updateById_(APP_CONFIG.sheets.NEW_HIRES, 'hire_id', hireId, {
      current_status: newStatus,
      updated_at: new Date().toISOString()
    });

    logAuditEvent_('FOLLOWUP_STATUS_UPDATED', 'NEW_HIRE', hireId, session.userId, { newStatus: newStatus });
    return respond_(true, { hireId: hireId, status: newStatus });
  }
};

// Expose to client
function getNewHeroesData(token) {
  return RecruitmentService.getNewHeroesData(token);
}

function updateNewHeroRecruitmentStatus(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.RESTAURANT_USER
  ]);

  const id = String(payload.id || '').trim();
  const status = String(payload.status || '').trim();
  const quitReason = String(payload.quitReason || '').trim();

  const allowed = ['Welcome Kit Done', 'Work Started', "Didn't Show Up", 'Quit'];

  if (!id) throw new Error('Employee ID is required.');
  if (allowed.indexOf(status) === -1) throw new Error('Invalid status.');
  if (status === 'Quit' && !quitReason) throw new Error('Quit reason is required.');

  const now = new Date().toISOString();

  const teamMember = Repository.findOneBy_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', id);

  if (teamMember) {
    Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', id, {
      status: status === 'Work Started' ? 'ACTIVE' : status,
      notes: status === 'Quit' ? quitReason : teamMember.notes || '',
      updated_at: now
    });

    logAuditEvent_('NEW_HERO_STATUS_UPDATED', 'TEAM_MEMBER', id, session.userId, {
      status: status,
      quitReason: quitReason
    });

    return respond_(true, { id: id, status: status });
  }

  const hire = Repository.findOneBy_(APP_CONFIG.sheets.NEW_HIRES, 'hire_id', id);
  if (!hire) throw new Error('New Hero not found.');

  Repository.updateById_(APP_CONFIG.sheets.NEW_HIRES, 'hire_id', id, {
    current_status: status,
    cancel_reason: status === 'Quit' ? quitReason : hire.cancel_reason || '',
    updated_at: now
  });

  logAuditEvent_('NEW_HERO_STATUS_UPDATED', 'NEW_HIRE', id, session.userId, {
    status: status,
    quitReason: quitReason
  });

  return respond_(true, { id: id, status: status });
}

function updateNewHero(token, payload) {
  return RecruitmentService.updateNewHero(token, payload);
}

function updateFollowUpStatus(token, payload) {
  return RecruitmentService.updateFollowUpStatus(token, payload);
}
function getHiringRequestWorkspaceData(token) {
  const session = getSessionOrThrow_(token)

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.RESTAURANT_USER
  ])

  const restaurants = AdminService.getRestaurantsForSession_(session)
    .filter(function(r) {
      return String(r.status || 'ACTIVE').toUpperCase() === 'ACTIVE'
    })
    .map(function(r) {
      return {
        restaurantId: String(r.restaurantId || '').trim(),
        restaurantCode: String(r.restaurantCode || '').trim(),
        restaurantName: String(r.restaurantName || '').trim()
      }
    })
    .filter(function(r) {
      return r.restaurantId || r.restaurantCode || r.restaurantName
    })
    .sort(function(a, b) {
      return String(a.restaurantCode || '').localeCompare(String(b.restaurantCode || ''))
    })

  return respond_(true, {
    restaurants: restaurants,
    requests: getHiringRequestRows_(session)
  })
}

function createHiringRequest(payload) {
  const session = getSessionOrThrow_(payload.sessionToken)

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.RESTAURANT_USER
  ])

  ensureHiringRequestsSheet_()

  const requesterName = String(payload.requesterName || payload.name || '').trim()
  const restaurantId = String(payload.restaurantId || '').trim()
  const position = String(payload.position || payload.requested_for_role || '').trim()
  const employmentType = String(payload.employmentType || payload.fullPartTime || '').trim()
  const availability = String(payload.availability || payload.shift || '').trim()
  const numberNeeded = String(payload.numberNeeded || payload.headcount || '1').trim()
  const priority = String(payload.priority || payload.urgency || 'Normal').trim()
  const neededBy = String(payload.neededBy || payload.preferredStartDate || '').trim()
  const reason = String(payload.reason || payload.businessReason || payload.business_reason || '').trim()
  const notes = String(payload.notes || payload.otherNotes || '').trim()

  if (!requesterName) throw new Error('Name is required.')
  if (!restaurantId) throw new Error('Restaurant is required.')
  if (!position) throw new Error('Position is required.')
  if (!employmentType) throw new Error('Full time or part time is required.')
  if (!availability) throw new Error('Availability is required.')
  if (!numberNeeded || Number(numberNeeded) < 1) throw new Error('Number needed is required.')
  if (!reason) throw new Error('Reason is required.')

  const visibleRestaurantIds = getRecruitmentRestaurantScopeIds_(session)

  if (visibleRestaurantIds !== null && visibleRestaurantIds.indexOf(String(restaurantId)) === -1) {
    throw new Error('You do not have access to this restaurant.')
  }

  const requestId = 'HR-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss')
  const now = new Date().toISOString()

  const payloadJson = {
    requesterName: requesterName,
    position: position,
    employmentType: employmentType,
    availability: availability,
    numberNeeded: numberNeeded,
    shift: availability,
    priority: priority,
    neededBy: neededBy,
    reason: reason,
    notes: notes
  }

  Repository.append_(APP_CONFIG.sheets.REQUESTS_HR, {
    request_id: requestId,
    request_type: 'HIRING_REQUEST',
    restaurant_id: restaurantId,
    requested_by: session.userId || '',
    requested_for_role: position,
    headcount: numberNeeded,
    urgency: priority,
    business_reason: reason,
    status: 'PENDING',
    submitted_at: now,
    assigned_hr_user_id: '',
    reviewed_at: '',
    completed_at: '',
    sla_due_at: '',
    timeline_json: JSON.stringify([
      {
        at: now,
        by: session.userId || '',
        action: 'CREATED',
        note: reason
      }
    ]),
    attachments: '',
    last_comment: notes,
    request_payload_json: JSON.stringify(payloadJson)
  })

  logAuditEvent_('HIRING_REQUEST_CREATED', 'HIRING_REQUEST', requestId, session.userId, {
    requesterName: requesterName,
    restaurantId: restaurantId,
    position: position,
    employmentType: employmentType,
    availability: availability,
    numberNeeded: numberNeeded,
    priority: priority
  })

  return respond_(true, {
    requestId: requestId
  })
}


function ensureHiringRequestsSheet_() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId)
  let sheet = ss.getSheetByName(APP_CONFIG.sheets.REQUESTS_HR)

  const schemaMap = getSheetSchemas_()
  const headers = schemaMap[APP_CONFIG.sheets.REQUESTS_HR] || schemaMap['Requests HR'] || [
    'request_id',
    'request_type',
    'restaurant_id',
    'requested_by',
    'requested_for_role',
    'headcount',
    'urgency',
    'business_reason',
    'status',
    'assigned_hr_user_id',
    'submitted_at',
    'reviewed_at',
    'completed_at',
    'sla_due_at',
    'timeline_json',
    'attachments',
    'last_comment',
    'request_payload_json'
  ]

  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.sheets.REQUESTS_HR)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
    return
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
    return
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1)
  const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim()
  })

  headers.forEach(function(header) {
    if (existingHeaders.indexOf(header) === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn())
      sheet.getRange(1, sheet.getLastColumn()).setValue(header)
      existingHeaders.push(header)
    }
  })

  sheet.setFrozenRows(1)
}


function getHiringRequestRows_(session) {
  ensureHiringRequestsSheet_()

  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId)
  const sheet = ss.getSheetByName(APP_CONFIG.sheets.REQUESTS_HR)

  if (!sheet || sheet.getLastRow() < 2) return []

  const values = sheet.getDataRange().getValues()
  const headers = values[0].map(function(header) {
    return String(header || '').trim()
  })

  const restaurantLookup = {}
  try {
    AdminService.getRestaurantsForSession_(session).forEach(function(r) {
      restaurantLookup[String(r.restaurantId || '')] = r
      restaurantLookup[String(r.restaurantCode || '')] = r
    })
  } catch (error) { console.warn('Ignored service error', error); }

  const rows = []

  for (let i = 1; i < values.length; i++) {
    const row = values[i]
    if (!row.some(function(cell) { return cell !== '' })) continue

    const obj = {}

    headers.forEach(function(header, index) {
      obj[header] = row[index]
    })

    let payload = {}

    try {
      payload = JSON.parse(obj.request_payload_json || '{}')
    } catch (error) {
      payload = {}
    }

    const restaurantId = String(obj.restaurant_id || '').trim()
    if (session && !rowMatchesRestaurantScopeV318_(obj, getSessionRestaurantScopeV318_(session))) continue
    const restaurant = restaurantLookup[restaurantId] || {}

    rows.push({
      requestId: obj.request_id || '',
      createdAt: obj.submitted_at || '',
      submittedAt: obj.submitted_at || '',
      restaurantId: restaurantId,
      restaurantCode: restaurant.restaurantCode || '',
      restaurantName: restaurant.restaurantName || '',
      requesterName: payload.requesterName || '',
      position: obj.requested_for_role || payload.position || '',
      employmentType: payload.employmentType || '',
      availability: payload.availability || payload.shift || '',
      numberNeeded: obj.headcount || payload.numberNeeded || '',
      shift: payload.shift || payload.availability || '',
      priority: obj.urgency || payload.priority || '',
      neededBy: payload.neededBy || '',
      reason: obj.business_reason || payload.reason || '',
      notes: obj.last_comment || payload.notes || '',
      status: obj.status || 'PENDING',
      createdByUserId: obj.requested_by || ''
    })
  }

  const _v500sorted = rows.sort(function(a, b) {
    return new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
  })
  // FIX v500 #1b: store in ScriptCache before returning
  try {
    const _v500json = JSON.stringify(_v500sorted)
    if (_v500json.length < 200000) CacheService.getScriptCache().put(_v500Key, _v500json, 90)
  } catch (_v500e) {}
  return _v500sorted
}


/*******************************************************
 MANAGE TEAM CLEARVIEW PATCH V195
 Safe replacements for Manage Team data, save, and status update.
*******************************************************/

function saveManageTeamEmployee(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const name = String(payload.name || '').trim();
  const mobile = String(payload.mobile || '').trim();
  const hiredDate = String(payload.hiredDate || '').trim();

  const restaurantIds = Array.isArray(payload.restaurantIds)
    ? payload.restaurantIds
    : Array.isArray(payload.assignedRestaurants)
      ? payload.assignedRestaurants
      : [];

  const cleanRestaurantIds = restaurantIds.map(function(id) {
    return String(id || '').trim();
  }).filter(Boolean);

  const status = payload.newlyHired ? 'Ready to Start' : 'ACTIVE';

  if (!name) throw new Error('Name is required.');
  if (!mobile || !/^[0-9+\-\s()]{7,20}$/.test(mobile)) {
    throw new Error('Enter a valid mobile number.');
  }
  if (!hiredDate) throw new Error('Hired date is required.');
  if (!cleanRestaurantIds.length) throw new Error('Select at least one restaurant.');

  const now = new Date().toISOString();
  const teamMemberId = Utilities.getUuid();

  Repository.append_(APP_CONFIG.sheets.TEAM_MEMBERS, {
    team_member_id: teamMemberId,
    employee_id: teamMemberId,
    full_name: name,
    restaurant_id: cleanRestaurantIds.join(','),
    job_title: 'Team Member',
    department: '',
    employment_type: '',
    hire_date: hiredDate,
    status: status,
    manager_name: '',
    email: '',
    phone: mobile,
    created_at: now,
    updated_at: now
  });

  let hireId = '';
  if (payload.newlyHired) {
    hireId = Utilities.getUuid();
    try {
      Repository.append_(APP_CONFIG.sheets.NEW_HIRES, {
        hire_id: hireId,
        candidate_name: name,
        applied_for_role: 'Team Member',
        restaurant_id: cleanRestaurantIds.join(','),
        assigned_hr_user_id: session.userId,
        application_date: hiredDate,
        current_status: 'Ready to Start',
        interview_date: '',
        orientation_date: '',
        sent_to_restaurant_at: '',
        restaurant_confirmation_at: '',
        start_date: '',
        notes: 'Mobile: ' + mobile + ' | Created from Manage Team',
        hold_reason: '',
        cancel_reason: '',
        updated_at: now
      });
    } catch (error) {
      console.error('New Hire mirror record failed', error);
    }
  }

  logAuditEvent_('TEAM_MEMBER_CREATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'MANAGE_TEAM_CLEARVIEW',
    name: name,
    restaurantIds: cleanRestaurantIds,
    newlyHired: Boolean(payload.newlyHired),
    hireId: hireId
  });

  return respond_(true, {
    teamMemberId: teamMemberId,
    hireId: hireId
  });
}

function updateManageTeamStatus(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const teamMemberId = String(payload.teamMemberId || '').trim();
  let status = String(payload.status || '').trim();

  const allowed = [
    'ACTIVE',
    'Active',
    'Newly Hired',
    'Orientation Done',
    'Ready to Work',
    'Ready to Start',
    'Follow Up Needed',
    'Work Started',
    'Quit',
    'Didnt Show Up',
    'Did Not Show Up',
    'INACTIVE',
    'Inactive'
  ];

  if (!teamMemberId) throw new Error('Employee ID is required.');
  if (allowed.indexOf(status) === -1) throw new Error('Invalid status.');

  if (status === 'Ready to Work') status = 'Ready to Start';
  if (status === 'Active' || status === 'Work Started') status = 'ACTIVE';
  if (status === 'Inactive') status = 'INACTIVE';
  if (status === 'Did Not Show Up' || status === "Didn't Show Up") status = 'Didnt Show Up';

  Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, {
    status: status,
    updated_at: new Date().toISOString()
  });

  logAuditEvent_('TEAM_MEMBER_STATUS_UPDATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'MANAGE_TEAM_CLEARVIEW',
    status: status
  });

  return respond_(true, {
    teamMemberId: teamMemberId,
    status: status
  });
}


function updateManageTeamEmployeeDetails(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const teamMemberId = String(payload.teamMemberId || '').trim();
  const name = String(payload.name || '').trim();
  const mobile = String(payload.mobile || '').trim();
  const hiredDate = String(payload.hiredDate || '').trim();

  const restaurantIds = Array.isArray(payload.restaurantIds)
    ? payload.restaurantIds
    : Array.isArray(payload.assignedRestaurants)
      ? payload.assignedRestaurants
      : [];

  const cleanRestaurantIds = restaurantIds.map(function(id) {
    return String(id || '').trim();
  }).filter(Boolean);

  if (!teamMemberId) throw new Error('Employee ID is required.');
  if (!name) throw new Error('Name is required.');
  if (!mobile || !/^[0-9+\-\s()]{7,20}$/.test(mobile)) {
    throw new Error('Enter a valid mobile number.');
  }
  if (!hiredDate) throw new Error('Hired date is required.');
  if (!cleanRestaurantIds.length) throw new Error('Select at least one restaurant.');

  const existing = Repository.findOneBy_(
    APP_CONFIG.sheets.TEAM_MEMBERS,
    'team_member_id',
    teamMemberId
  );

  if (!existing) throw new Error('Employee not found.');

  Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, {
    full_name: name,
    phone: mobile,
    hire_date: hiredDate,
    restaurant_id: cleanRestaurantIds.join(','),
    updated_at: new Date().toISOString()
  });

  logAuditEvent_('TEAM_MEMBER_UPDATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'MANAGE_TEAM_CLEARVIEW',
    name: name,
    restaurantIds: cleanRestaurantIds
  });

  return respond_(true, {
    teamMemberId: teamMemberId
  });
}

function deactivateManageTeamEmployee(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const teamMemberId = String(payload.teamMemberId || '').trim();
  if (!teamMemberId) throw new Error('Employee ID is required.');

  const existing = Repository.findOneBy_(
    APP_CONFIG.sheets.TEAM_MEMBERS,
    'team_member_id',
    teamMemberId
  );

  if (!existing) throw new Error('Employee not found.');

  Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, {
    status: 'INACTIVE',
    updated_at: new Date().toISOString()
  });

  logAuditEvent_('TEAM_MEMBER_DEACTIVATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'MANAGE_TEAM_CLEARVIEW'
  });

  return respond_(true, {
    teamMemberId: teamMemberId,
    status: 'INACTIVE'
  });
}

function saveTeamMemberFromManageTeam(payload) {
  return saveManageTeamEmployee(payload);
}

function updateTeamMemberStatusFromManageTeam(payload) {
  return updateManageTeamStatus(payload);
}

function updateManageTeamEmployeeDetailsFromManageTeam(payload) {
  return updateManageTeamEmployeeDetails(payload);
}

function deactivateManageTeamEmployeeFromManageTeam(payload) {
  return deactivateManageTeamEmployee(payload);
}


/*******************************************************
 MANAGE TEAM TRANSFER PATCH V199
 Adds Employee Transfer History and transfer workflow.
*******************************************************/

function ensureEmployeeTransferHistorySheet_() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const sheetName = 'Employee Transfer History';
  const headers = [
    'transfer_id',
    'team_member_id',
    'employee_name',
    'from_restaurant_id',
    'from_restaurant_name',
    'to_restaurant_id',
    'to_restaurant_name',
    'transfer_date',
    'notes',
    'transferred_by',
    'transferred_by_name',
    'employee_snapshot_json',
    'created_at'
  ];

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getEmployeeTransferHistoryRows_() {
  const sheet = ensureEmployeeTransferHistorySheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  return values.slice(1).filter(function(row) {
    return row.some(function(cell) { return cell !== ''; });
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    return {
      transferId: obj.transfer_id || '',
      teamMemberId: obj.team_member_id || '',
      employeeName: obj.employee_name || '',
      fromRestaurantId: obj.from_restaurant_id || '',
      fromRestaurantName: obj.from_restaurant_name || '',
      toRestaurantId: obj.to_restaurant_id || '',
      toRestaurantName: obj.to_restaurant_name || '',
      transferDate: obj.transfer_date || '',
      notes: obj.notes || '',
      transferredBy: obj.transferred_by || '',
      transferredByName: obj.transferred_by_name || '',
      createdAt: obj.created_at || ''
    };
  });
}

function appendEmployeeTransferHistory_(record) {
  const sheet = ensureEmployeeTransferHistorySheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });

  const row = headers.map(function(header) {
    return header in record ? record[header] : '';
  });

  sheet.appendRow(row);
}

function getRestaurantDisplayNameById_(restaurantLookup, restaurantId) {
  const restaurant = restaurantLookup[String(restaurantId || '')] || {};
  if (!restaurant) return String(restaurantId || '');
  return ((restaurant.restaurantCode ? restaurant.restaurantCode + ' ' : '') + (restaurant.restaurantName || '')).trim() || String(restaurantId || '');
}

function getManageTeamData(token) {
  const session = getSessionOrThrow_(token);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const restaurants = AdminService.getRestaurantsForSession_(session);
  const restaurantLookup = {};

  restaurants.forEach(function(r) {
    restaurantLookup[String(r.restaurantId)] = r;
  });

  const transferHistory = filterRowsForSessionByRestaurantV318_(session, getEmployeeTransferHistoryRows_());
  const transferHistoryByMember = {};

  transferHistory.forEach(function(row) {
    const key = String(row.teamMemberId || '');
    if (!key) return;
    transferHistoryByMember[key] = transferHistoryByMember[key] || [];
    transferHistoryByMember[key].push(row);
  });

  const teamMembers = filterRowsForSessionByRestaurantV318_(session, Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS))
    .sort(function(a, b) {
      return new Date(b.created_at || b.updated_at || 0) - new Date(a.created_at || a.updated_at || 0);
    })
    .map(function(row) {
      const restaurantIds = String(row.restaurant_id || '')
        .split(',')
        .map(function(item) { return item.trim(); })
        .filter(Boolean);

      const restaurantNames = restaurantIds.map(function(id) {
        return getRestaurantDisplayNameById_(restaurantLookup, id);
      });

      const restaurantCodes = restaurantIds.map(function(id) {
        const r = restaurantLookup[id];
        if (!r) return '';
        return String(r.restaurantCode || '').trim();
      }).filter(Boolean);

      const teamMemberId = row.team_member_id || '';

      return {
        teamMemberId: teamMemberId,
        employeeId: row.employee_id || '',
        name: row.full_name || '',
        mobile: row.phone || '',
        email: row.email || '',
        hiredDate: row.hire_date || '',
        restaurantIds: restaurantIds,
        restaurantName: restaurantNames.join(', '),
        restaurantCodes: restaurantCodes,
        restaurantCode: restaurantCodes.join(', '),
        jobTitle: row.job_title || '',
        employmentType: row.employment_type || '',
        status: row.status || 'Ready to Start',
        createdAt: row.created_at || '',
        updatedAt: row.updated_at || '',
        transferHistory: transferHistoryByMember[String(teamMemberId)] || []
      };
    });

  return respond_(true, {
    restaurants: restaurants,
    teamMembers: teamMembers,
    employees: teamMembers,
    transferHistory: transferHistory
  });
}

function transferManageTeamEmployee(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);

  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO
  ]);

  const teamMemberId = String(payload.teamMemberId || '').trim();
  const toRestaurantId = String(payload.toRestaurantId || '').trim();
  const transferDate = String(payload.transferDate || '').trim();
  const notes = String(payload.notes || '').trim();

  if (!teamMemberId) throw new Error('Employee is required.');
  if (!toRestaurantId) throw new Error('Transfer restaurant is required.');
  if (!transferDate) throw new Error('Transfer date is required.');
  if (!notes) throw new Error('Transfer notes are required.');

  const existing = Repository.findOneBy_(
    APP_CONFIG.sheets.TEAM_MEMBERS,
    'team_member_id',
    teamMemberId
  );

  if (!existing) throw new Error('Employee not found.');

  const restaurants = AdminService.getRestaurantsForSession_(session);
  const restaurantLookup = {};
  restaurants.forEach(function(r) {
    restaurantLookup[String(r.restaurantId)] = r;
  });

  if (!restaurantLookup[toRestaurantId]) {
    throw new Error('Transfer restaurant not found.');
  }

  const fromRestaurantIds = String(existing.restaurant_id || '')
    .split(',')
    .map(function(item) { return item.trim(); })
    .filter(Boolean);

  const fromRestaurantId = fromRestaurantIds[0] || '';

  if (String(fromRestaurantId) === String(toRestaurantId)) {
    throw new Error('Employee is already assigned to this restaurant.');
  }

  const now = new Date().toISOString();
  const transferId = Utilities.getUuid();
  const employeeName = existing.full_name || '';
  const fromRestaurantName = getRestaurantDisplayNameById_(restaurantLookup, fromRestaurantId);
  const toRestaurantName = getRestaurantDisplayNameById_(restaurantLookup, toRestaurantId);

  Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, {
    restaurant_id: toRestaurantId,
    updated_at: now
  });

  appendEmployeeTransferHistory_({
    transfer_id: transferId,
    team_member_id: teamMemberId,
    employee_name: employeeName,
    from_restaurant_id: fromRestaurantId,
    from_restaurant_name: fromRestaurantName,
    to_restaurant_id: toRestaurantId,
    to_restaurant_name: toRestaurantName,
    transfer_date: transferDate,
    notes: notes,
    transferred_by: session.userId,
    transferred_by_name: session.fullName || session.username || '',
    employee_snapshot_json: JSON.stringify(existing || {}),
    created_at: now
  });

  logAuditEvent_('TEAM_MEMBER_TRANSFERRED', 'TEAM_MEMBER', teamMemberId, session.userId, {
    source: 'MANAGE_TEAM_CLEARVIEW',
    transferId: transferId,
    fromRestaurantId: fromRestaurantId,
    toRestaurantId: toRestaurantId,
    transferDate: transferDate,
    notes: notes
  });

  return respond_(true, {
    transferId: transferId,
    teamMemberId: teamMemberId,
    fromRestaurantId: fromRestaurantId,
    toRestaurantId: toRestaurantId,
    transferDate: transferDate
  });
}

/*******************************************************
 ADMIN REQUEST SAFE PATCH V213
 Paste this at the bottom of the current working Services.gs.
 Do not replace existing login or dashboard functions.
*******************************************************/

function createAdminJobLetterRequest(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  if (!payload.restaurantId) throw new Error('Select restaurant.');
  if (!payload.letterType) throw new Error('Select type of letter.');
  if (!payload.employeeId) throw new Error('Select employee name.');
  if (!payload.gender) throw new Error('Select team member gender.');
  if (!payload.position) throw new Error('Select team member position.');
  if (!payload.employmentType) throw new Error('Select full time or part time.');
  if (!payload.startedDate) throw new Error('Select start date.');

  return createAdminRequestRecordV213_(session, 'JOB_LETTER', payload);
}

function createAdminOfficeSupplyRequest(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  if (!payload.restaurantId) throw new Error('Select restaurant.');
  if (!payload.supplyType) throw new Error('Select request type.');
  if (!payload.supplyDetails) throw new Error('Please specify your office supplies need.');

  return createAdminRequestRecordV213_(session, 'OFFICE_SUPPLIES', payload);
}

function createAdminRequestRecordV213_(session, requestType, payload) {
  const sheet = ensureAdminRequestsSheetV213_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });

  const now = new Date();
  const nowIso = now.toISOString();
  const requestId = Utilities.getUuid();
  const restaurantNumber = getRestaurantNumberForAdminRequestV213_(payload.restaurantId);
  const statusHistory = [
    {
      at: nowIso,
      by: session.fullName || session.username || session.userId || '',
      status: 'New Request',
      note: 'Request created'
    }
  ];

  const record = {
    request_id: requestId,
    request_type: requestType,
    restaurant_id: payload.restaurantId || '',
    restaurant_number: restaurantNumber,
    requested_by: session.userId || '',
    requested_by_name: session.fullName || session.username || '',
    letter_type: payload.letterType || '',
    employee_id: payload.employeeId || '',
    employee_name: payload.employeeName || '',
    employee_status: payload.employeeStatus || '',
    gender: payload.gender || '',
    position: payload.position || '',
    manual_position: payload.manualPosition || '',
    employment_type: payload.employmentType || '',
    started_date: payload.startedDate || '',
    supply_type: payload.supplyType || '',
    supply_details: payload.supplyDetails || '',
    issuer_notes: payload.issuerNotes || '',
    admin_notes: '',
    status: 'New Request',
    status_updated_by: '',
    status_updated_at: '',
    status_history_json: JSON.stringify(statusHistory),
    submitted_at: now,
    updated_at: now,
    request_payload_json: JSON.stringify(payload || {})
  };

  const row = headers.map(function(header) {
    return header in record ? record[header] : '';
  });

  sheet.appendRow(row);
  try { formatAdminRequestDateColumnsV418_(sheet); } catch (error) {}
  logAuditEvent_('ADMIN_REQUEST_CREATED', 'ADMIN_REQUEST', requestId, session.userId, {
    requestType: requestType,
    restaurantId: payload.restaurantId || ''
  });

    try { repairAdminRequestRowsV420(); } catch (error) { Logger.log(error.message); }

  return respond_(true, {
    requestId: requestId,
    request: {
      requestId: requestId,
      requestType: requestType,
      restaurantId: payload.restaurantId || '',
      restaurantNumber: restaurantNumber,
      restaurantCode: restaurantNumber,
      letterType: payload.letterType || '',
      employeeId: payload.employeeId || '',
      employeeName: payload.employeeName || '',
      employeeStatus: payload.employeeStatus || '',
      gender: payload.gender || '',
      position: payload.position || '',
      manualPosition: payload.manualPosition || '',
      employmentType: payload.employmentType || '',
      startedDate: payload.startedDate || '',
      supplyType: payload.supplyType || '',
      supplyDetails: payload.supplyDetails || '',
      issuerNotes: payload.issuerNotes || '',
      adminNotes: '',
      status: 'New Request',
      statusHistory: statusHistory,
      createdAt: nowIso,
      submittedAt: nowIso,
      updatedAt: nowIso
    }
  });
}


function parseAdminRequestPayloadFromRowV416_(obj) {
  const keys = Object.keys(obj || {});
  const candidates = [
    obj.request_payload_json,
    obj.payload_json,
    obj.requestPayloadJson,
    obj.payloadJson
  ];

  for (let i = 0; i < keys.length; i++) {
    candidates.push(obj[keys[i]]);
  }

  for (let i = 0; i < candidates.length; i++) {
    const text = String(candidates[i] || '').trim();
    if (!text || (text[0] !== '{' && text[0] !== '[')) continue;

    try {
      const parsed = JSON.parse(text);
      if (
        parsed &&
        typeof parsed === 'object' &&
        !Array.isArray(parsed) &&
        (parsed.supplyType || parsed.supplyDetails || parsed.letterType || parsed.employeeName || parsed.restaurantId || parsed.requestType)
      ) {
        return parsed;
      }
    } catch (error) {}
  }

  return {};
}

function cleanAdminRequestCellV416_(value, fallback) {
  const text = String(value || '').trim();
  if (!text) return fallback || '';

  if (text[0] === '{' || text[0] === '[') {
    try {
      JSON.parse(text);
      return fallback || '';
    } catch (error) {}
  }

  return text;
}

function getAdminRequestRowsV213_() {
  // FIX v500 #1a: Cache results 90s — called 3-5x per request with no prior caching.
  const _v500Key = 'v500_admin_req_rows';
  try {
    const _v500c = CacheService.getScriptCache().get(_v500Key);
    if (_v500c) return JSON.parse(_v500c);
  } catch (_v500e) {}

  const sheet = ensureAdminRequestsSheetV213_();
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row.some(function(cell) { return cell !== ''; })) continue;

    const obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    let statusHistory = [];
    try {
      statusHistory = JSON.parse(obj.status_history_json || '[]');
    } catch (error) {
      statusHistory = [];
    }

    const payload = parseAdminRequestPayloadFromRowV416_(obj);
    const requestType = cleanAdminRequestCellV416_(obj.request_type, payload.requestType || payload.request_type || '');
    const inferredType = requestType ||
      (payload.supplyType || payload.supplyDetails || obj.supply_type || obj.supply_details ? 'OFFICE_SUPPLIES' : '') ||
      (payload.letterType || payload.employeeName || obj.letter_type || obj.employee_name ? 'JOB_LETTER' : '');

    rows.push({
      requestId: cleanAdminRequestCellV416_(obj.request_id, payload.requestId || payload.request_id || ''),
      requestType: inferredType,
      restaurantId: cleanAdminRequestCellV416_(obj.restaurant_id, payload.restaurantId || payload.restaurant_id || ''),
      restaurantNumber: cleanAdminRequestCellV416_(obj.restaurant_number, payload.restaurantNumber || payload.restaurant_number || ''),
      restaurantCode: cleanAdminRequestCellV416_(obj.restaurant_number, payload.restaurantNumber || payload.restaurant_number || ''),
      requestedBy: cleanAdminRequestCellV416_(obj.requested_by, ''),
      requestedByName: cleanAdminRequestCellV416_(obj.requested_by_name, ''),
      letterType: cleanAdminRequestCellV416_(obj.letter_type, payload.letterType || payload.letter_type || ''),
      employeeId: cleanAdminRequestCellV416_(obj.employee_id, payload.employeeId || payload.employee_id || ''),
      employeeName: cleanAdminRequestCellV416_(obj.employee_name, payload.employeeName || payload.employee_name || ''),
      employeeStatus: cleanAdminRequestCellV416_(obj.employee_status, payload.employeeStatus || payload.employee_status || ''),
      gender: cleanAdminRequestCellV416_(obj.gender, payload.gender || ''),
      position: cleanAdminRequestCellV416_(obj.position, payload.position || ''),
      manualPosition: cleanAdminRequestCellV416_(obj.manual_position, payload.manualPosition || payload.manual_position || ''),
      employmentType: cleanAdminRequestCellV416_(obj.employment_type, payload.employmentType || payload.employment_type || ''),
      startedDate: cleanAdminRequestCellV416_(obj.started_date, payload.startedDate || payload.started_date || ''),
      supplyType: cleanAdminRequestCellV416_(obj.supply_type, payload.supplyType || payload.supply_type || ''),
      supplyDetails: cleanAdminRequestCellV416_(obj.supply_details, payload.supplyDetails || payload.supply_details || ''),
      issuerNotes: cleanAdminRequestCellV416_(obj.issuer_notes, payload.issuerNotes || payload.issuer_notes || payload.notes || payload.supplyDetails || ''),
      adminNotes: cleanAdminRequestCellV416_(obj.admin_notes, ''),
      restaurantResponseNotes: cleanAdminRequestCellV416_(obj.restaurant_response_notes, payload.restaurantResponseNotes || ''),
      restaurantRespondedBy: cleanAdminRequestCellV416_(obj.restaurant_responded_by, ''),
      restaurantRespondedAt: cleanAdminRequestCellV416_(obj.restaurant_responded_at, ''),
      status: cleanAdminRequestCellV416_(obj.status, payload.status || 'New Request') || 'New Request',
      statusUpdatedBy: cleanAdminRequestCellV416_(obj.status_updated_by, ''),
      statusUpdatedAt: cleanAdminRequestCellV416_(obj.status_updated_at, ''),
      statusHistory: statusHistory,
      createdAt: cleanAdminRequestCellV416_(obj.submitted_at, obj.created_at || payload.createdAt || ''),
      submittedAt: cleanAdminRequestCellV416_(obj.submitted_at, obj.created_at || payload.createdAt || ''),
      updatedAt: cleanAdminRequestCellV416_(obj.updated_at, '')
    });
  }

  const _v500sorted = rows.sort(function(a, b) {
    return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
  });
  // FIX v500 #1a: store result in ScriptCache
  try {
    const _v500json = JSON.stringify(_v500sorted);
    if (_v500json.length < 200000) CacheService.getScriptCache().put(_v500Key, _v500json, 90);
  } catch (_v500e) {}
  return _v500sorted;
}

// NOTE: updateAdminRequestStatus is defined once below (canonical V422 version).
// The earlier duplicate definitions at lines ~9625 and ~9998 are removed.
// ── CANONICAL updateAdminRequestStatus (replaces all prior duplicates) ──
function updateAdminRequestStatus(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const requestId = String(payload.requestId || '').trim();
  const newStatus = String(payload.status || '').trim() || 'New Request';
  const adminNotes = String(payload.adminNotes || '').trim();
  if (!requestId) throw new Error('Missing request id.');

  const sheet = ensureAdminRequestsSheetV213_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  const idIndex = headers.indexOf('request_id');
  if (idIndex === -1) throw new Error('Request id column not found.');

  const statusIndex = headers.indexOf('status');
  const adminNotesIndex = headers.indexOf('admin_notes');
  const updatedByIndex = headers.indexOf('status_updated_by');
  const updatedAtIndex = headers.indexOf('status_updated_at');
  const updatedIndex = headers.indexOf('updated_at');
  const historyIndex = headers.indexOf('status_history_json');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

      const now = new Date();
      const nowIso = now.toISOString();
      let history = [];
      try {
        history = JSON.parse(values[r][historyIndex] || '[]');
      } catch (error) {
        history = [];
      }

      history.push({
        at: nowIso,
        by: session.fullName || session.username || session.userId || '',
        status: newStatus,
        note: adminNotes
      });

      if (statusIndex !== -1) values[r][statusIndex] = newStatus;
      if (adminNotesIndex !== -1) values[r][adminNotesIndex] = adminNotes;
      if (updatedByIndex !== -1) values[r][updatedByIndex] = session.fullName || session.username || session.userId || '';
      if (updatedAtIndex !== -1) values[r][updatedAtIndex] = now;
      if (updatedIndex !== -1) values[r][updatedIndex] = now;
      if (historyIndex !== -1) values[r][historyIndex] = JSON.stringify(history);

      sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);
      try { formatAdminRequestDateColumnsV418_(sheet); } catch (error) {}
      logAuditEvent_('ADMIN_REQUEST_STATUS_UPDATED', 'ADMIN_REQUEST', requestId, session.userId, {
        status: newStatus
      });

      return respond_(true, {
        requestId: requestId,
        status: newStatus
      });
  }

  throw new Error('Request not found.');
}

function getRestaurantNumberForAdminRequestV213_(restaurantId) {
  const id = String(restaurantId || '').trim();
  try {
    const restaurant = AdminService.getRestaurants_().find(function(r) {
      return String(r.restaurantId || '') === id;
    });
    if (restaurant) return String(restaurant.restaurantCode || restaurant.restaurantNumber || restaurant.restaurant_name || id || '').trim();
  } catch (error) { console.warn('Ignored service error', error); }
  return id;
}



/**
 * V418 date formatting helper.
 * Only formats Admin Request date columns.
 * Does not touch login, session, or Users logic.
 */
function formatAdminRequestDateColumnsV418_(sheet) {
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  ['submitted_at', 'updated_at', 'status_updated_at'].forEach(function(header) {
    const index = headers.indexOf(header);
    if (index !== -1) {
      sheet.getRange(2, index + 1, Math.max(sheet.getMaxRows() - 1, 1))
        .setNumberFormat('yyyy-mm-dd hh:mm AM/PM');
    }
  });
}
/*******************************************************
 RESTAURANT REQUEST RECEIVER PATCH V219
 Paste this at the bottom of the current working Services.gs.
 Do not replace existing login or dashboard functions.
*******************************************************/

function updateRestaurantRequestReceiverStatus(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const source = String(payload.source || '').trim().toUpperCase();
  const requestId = String(payload.requestId || '').trim();
  const status = String(payload.status || payload.displayStatus || '').trim();
  const notes = String(payload.notes || '').trim();
  const quantity = String(payload.quantity || '').trim();

  if (!requestId) throw new Error('Missing request id.');
  if (!status) throw new Error('Missing status.');

  if (source === 'HIRING') {
    return updateHiringRequestReceiverStatusV219_(session, requestId, status, notes);
  }

  return updateAdminRequestReceiverStatusV219_(session, requestId, status, notes, quantity);
}

function updateHiringRequestReceiverStatusV219_(session, requestId, status, notes) {
  ensureHiringRequestsSheet_();
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(APP_CONFIG.sheets.REQUESTS_HR);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Hiring request sheet is empty.');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idIndex = headers.indexOf('request_id');
  if (idIndex === -1) throw new Error('request_id column not found.');

  const statusIndex = headers.indexOf('status');
  const lastCommentIndex = headers.indexOf('last_comment');
  const timelineIndex = headers.indexOf('timeline_json');
  const reviewedAtIndex = headers.indexOf('reviewed_at');
  const completedAtIndex = headers.indexOf('completed_at');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

    const now = new Date().toISOString();
    let timeline = [];
    try {
      timeline = JSON.parse(values[r][timelineIndex] || '[]');
    } catch (error) {
      timeline = [];
    }

    timeline.push({
      at: now,
      by: session.fullName || session.username || session.userId || '',
      action: 'STATUS_UPDATED',
      status: status,
      note: notes
    });

    if (statusIndex !== -1) values[r][statusIndex] = status;
    if (lastCommentIndex !== -1) values[r][lastCommentIndex] = notes;
    if (timelineIndex !== -1) values[r][timelineIndex] = JSON.stringify(timeline);
    if (reviewedAtIndex !== -1 && !values[r][reviewedAtIndex]) values[r][reviewedAtIndex] = now;
    if (completedAtIndex !== -1 && (status === 'Hiring Complete' || status === 'COMPLETED' || status === 'Completed')) values[r][completedAtIndex] = now;

    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);

    logAuditEvent_('HIRING_REQUEST_STATUS_UPDATED', 'HIRING_REQUEST', requestId, session.userId, {
      status: status
    });

    return respond_(true, {
      requestId: requestId,
      status: status
    });
  }

  throw new Error('Hiring request not found.');
}

function uploadRestaurantRequestLetter(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const requestId = String(payload.requestId || '').trim();
  const fileName = String(payload.fileName || 'letter.pdf').trim();
  const mimeType = String(payload.mimeType || 'application/octet-stream').trim();
  const base64 = String(payload.base64 || '').trim();

  if (!requestId) throw new Error('Missing request id.');
  if (!base64) throw new Error('Missing file data.');

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const folder = getRestaurantRequestLetterFolderV219_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const attachment = {
    name: fileName,
    fileId: file.getId(),
    url: file.getUrl(),
    uploadedBy: session.fullName || session.username || session.userId || '',
    uploadedAt: new Date().toISOString()
  };

  appendAdminRequestAttachmentV219_(session, requestId, attachment);

  return respond_(true, {
    attachment: attachment
  });
}

function getRestaurantRequestLetterFolderV219_() {
  const folderName = 'Restaurant Request Letters';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function appendAdminRequestAttachmentV219_(session, requestId, attachment) {
  const sheet = ensureAdminRequestsSheetV219_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idIndex = headers.indexOf('request_id');
  const attachmentsIndex = headers.indexOf('attachments_json');
  const historyIndex = headers.indexOf('status_history_json');
  const updatedIndex = headers.indexOf('updated_at');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

    let attachments = [];
    try {
      attachments = JSON.parse(values[r][attachmentsIndex] || '[]');
    } catch (error) {
      attachments = [];
    }
    attachments.push(attachment);

    let history = [];
    try {
      history = JSON.parse(values[r][historyIndex] || '[]');
    } catch (error2) {
      history = [];
    }
    history.push({
      at: attachment.uploadedAt,
      by: session.fullName || session.username || session.userId || '',
      status: values[r][headers.indexOf('status')] || 'New Request',
      note: 'Letter uploaded: ' + attachment.name
    });

    if (attachmentsIndex !== -1) values[r][attachmentsIndex] = JSON.stringify(attachments);
    if (historyIndex !== -1) values[r][historyIndex] = JSON.stringify(history);
    if (updatedIndex !== -1) values[r][updatedIndex] = attachment.uploadedAt;

    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);
    return true;
  }

  throw new Error('Admin request not found.');
}

function ensureAdminRequestsSheetV219_() {
  const sheet = ensureAdminRequestsSheetV213_();
  const required = ['quantity', 'attachments_json'];
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });

  required.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue(header);
      headers.push(header);
    }
  });

  return sheet;
}

function getRestaurantRequestAdminRowsV219_() {
  const sheet = ensureAdminRequestsSheetV219_();
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row.some(function(cell) { return cell !== ''; })) continue;

    const obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    let statusHistory = [];
    try {
      statusHistory = JSON.parse(obj.status_history_json || '[]');
    } catch (error) { console.warn('Ignored service error', error); }

    let attachments = [];
    try {
      attachments = JSON.parse(obj.attachments_json || '[]');
    } catch (error2) {}

    rows.push({
      requestId: obj.request_id || '',
      requestType: obj.request_type || '',
      restaurantId: obj.restaurant_id || '',
      restaurantNumber: obj.restaurant_number || '',
      restaurantCode: obj.restaurant_number || '',
      requestedBy: obj.requested_by || '',
      requestedByName: obj.requested_by_name || '',
      letterType: obj.letter_type || '',
      employeeId: obj.employee_id || '',
      employeeName: obj.employee_name || '',
      employeeStatus: obj.employee_status || '',
      gender: obj.gender || '',
      position: obj.position || '',
      manualPosition: obj.manual_position || '',
      employmentType: obj.employment_type || '',
      startedDate: obj.started_date || '',
      supplyType: obj.supply_type || '',
      supplyDetails: obj.supply_details || '',
      issuerNotes: obj.issuer_notes || '',
      adminNotes: obj.admin_notes || '',
      quantity: obj.quantity || '',
      attachments: attachments,
      status: obj.status || 'New Request',
      statusUpdatedBy: obj.status_updated_by || '',
      statusUpdatedAt: obj.status_updated_at || '',
      statusHistory: statusHistory,
      createdAt: obj.submitted_at || '',
      submittedAt: obj.submitted_at || '',
      updatedAt: obj.updated_at || ''
    });
  }

  return rows.sort(function(a, b) {
    return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
  });
}

function getHiringRequestTimelineV219_(requestId) {
  try {
    ensureHiringRequestsSheet_();
    const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
    const sheet = ss.getSheetByName(APP_CONFIG.sheets.REQUESTS_HR);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(function(h) { return String(h || '').trim(); });
    const idIndex = headers.indexOf('request_id');
    const timelineIndex = headers.indexOf('timeline_json');
    if (idIndex === -1 || timelineIndex === -1) return [];

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIndex] || '') === String(requestId || '')) {
        try {
          return JSON.parse(values[r][timelineIndex] || '[]');
        } catch (error) {
          return [];
        }
      }
    }
  } catch (error2) {}
  return [];
}

/*******************************************************
 ROLE ACTION LABEL OVERRIDE V220
 Keeps button keys same but shows correct names.
*******************************************************/
function markRestaurantRequestReceiverViewed(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const source = String(payload.source || '').trim().toUpperCase();
  const requestId = String(payload.requestId || '').trim();
  if (!requestId) throw new Error('Missing request id.');

  if (source === 'HIRING') {
    return markHiringRequestViewedV222_(session, requestId);
  }
  return markAdminRequestViewedV222_(session, requestId);
}

function markHiringRequestViewedV222_(session, requestId) {
  ensureHiringRequestsSheet_();
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(APP_CONFIG.sheets.REQUESTS_HR);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Hiring request sheet is empty.');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idIndex = headers.indexOf('request_id');
  const statusIndex = headers.indexOf('status');
  const lastCommentIndex = headers.indexOf('last_comment');
  const timelineIndex = headers.indexOf('timeline_json');
  const reviewedAtIndex = headers.indexOf('reviewed_at');

  if (idIndex === -1) throw new Error('request_id column not found.');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

    const currentStatus = statusIndex !== -1 ? String(values[r][statusIndex] || '') : '';
    if (currentStatus === 'Hiring Complete' || currentStatus === 'COMPLETED' || currentStatus === 'Completed') {
      return respond_(true, { requestId: requestId, status: currentStatus });
    }

    const now = new Date().toISOString();
    let timeline = [];
    try {
      timeline = timelineIndex !== -1 ? JSON.parse(values[r][timelineIndex] || '[]') : [];
    } catch (error) {
      timeline = [];
    }

    timeline.push({
      at: now,
      by: session.fullName || session.username || session.userId || '',
      action: 'REQUEST_VIEWED',
      status: 'Viewed',
      note: 'Request viewed'
    });

    if (statusIndex !== -1) values[r][statusIndex] = 'Viewed';
    if (lastCommentIndex !== -1 && !values[r][lastCommentIndex]) values[r][lastCommentIndex] = 'Request viewed';
    if (timelineIndex !== -1) values[r][timelineIndex] = JSON.stringify(timeline);
    if (reviewedAtIndex !== -1 && !values[r][reviewedAtIndex]) values[r][reviewedAtIndex] = now;

    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);

    logAuditEvent_('HIRING_REQUEST_VIEWED', 'HIRING_REQUEST', requestId, session.userId, {});
    return respond_(true, { requestId: requestId, status: 'Viewed' });
  }

  throw new Error('Hiring request not found.');
}

function markAdminRequestViewedV222_(session, requestId) {
  const sheet = ensureAdminRequestsSheetV219_();
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Admin request sheet is empty.');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idIndex = headers.indexOf('request_id');
  const statusIndex = headers.indexOf('status');
  const updatedByIndex = headers.indexOf('status_updated_by');
  const updatedAtIndex = headers.indexOf('status_updated_at');
  const updatedIndex = headers.indexOf('updated_at');
  const historyIndex = headers.indexOf('status_history_json');

  if (idIndex === -1) throw new Error('request_id column not found.');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

    const currentStatus = statusIndex !== -1 ? String(values[r][statusIndex] || '') : '';
    if (currentStatus === 'Completed') {
      return respond_(true, { requestId: requestId, status: currentStatus });
    }

    const now = new Date().toISOString();
    let history = [];
    try {
      history = historyIndex !== -1 ? JSON.parse(values[r][historyIndex] || '[]') : [];
    } catch (error) {
      history = [];
    }

    history.push({
      at: now,
      by: session.fullName || session.username || session.userId || '',
      status: 'Viewed',
      note: 'Request viewed'
    });

    if (statusIndex !== -1) values[r][statusIndex] = 'Viewed';
    if (updatedByIndex !== -1) values[r][updatedByIndex] = session.fullName || session.username || session.userId || '';
    if (updatedAtIndex !== -1) values[r][updatedAtIndex] = now;
    if (updatedIndex !== -1) values[r][updatedIndex] = now;
    if (historyIndex !== -1) values[r][historyIndex] = JSON.stringify(history);

    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);

    logAuditEvent_('ADMIN_REQUEST_VIEWED', 'ADMIN_REQUEST', requestId, session.userId, {});
    return respond_(true, { requestId: requestId, status: 'Viewed' });
  }

  throw new Error('Admin request not found.');
}

/*******************************************************
 RESTAURANT REQUEST RECEIVER PATCH V223
 Paste this at the very bottom of current working Services.gs.
 This overrides only data mapping so the page loads cleaner
 and hiring request restaurant/details show properly.
*******************************************************/

function mapRestaurantRequestHiringRowV223_(row) {
  row = row || {};
  const requestId = rrPickV223_(row, [
    'requestId', 'request_id', 'id', 'Request ID'
  ]);
  const restaurantNumber = rrPickV223_(row, [
    'restaurantNumber', 'restaurant_number', 'restaurantCode', 'restaurant_code',
    'restaurant', 'Restaurant', 'Restaurant Number', 'Restaurant #',
    'restaurantId', 'restaurant_id', 'location', 'assignedRestaurant', 'Assigned Restaurant'
  ]);
  const status = rrPickV223_(row, [
    'status', 'Status'
  ]) || 'Pending';
  const createdAt = rrPickV223_(row, [
    'createdAt', 'created_at', 'submittedAt', 'submitted_at', 'timestamp', 'Timestamp', 'Request Date'
  ]);

  let timeline = [];
  try {
    if (typeof getHiringRequestTimelineV219_ === 'function' && requestId) {
      timeline = getHiringRequestTimelineV219_(requestId);
    } else {
      timeline = JSON.parse(rrPickV223_(row, ['timeline_json', 'timelineJson', 'statusHistory']) || '[]');
    }
  } catch (error) {
    timeline = [];
  }

  return {
    requestId: requestId,
    source: 'HIRING',
    requestType: 'HIRING',
    restaurantNumber: restaurantNumber,
    restaurantCode: restaurantNumber,
    restaurantId: rrPickV223_(row, ['restaurantId', 'restaurant_id']) || restaurantNumber,
    requesterName: rrPickV223_(row, ['requesterName', 'requestedByName', 'requested_by_name', 'requested_by', 'requestedBy', 'Your Name', 'name']),
    requestedByName: rrPickV223_(row, ['requestedByName', 'requested_by_name', 'requesterName', 'Your Name']),
    position: rrPickV223_(row, ['position', 'Position', 'jobTitle', 'Job Title']),
    employmentType: rrPickV223_(row, ['employmentType', 'employment_type', 'type', 'Full-time/Part-time', 'Full time or Part time']),
    availability: rrPickV223_(row, ['availability', 'Availability', 'availabilityRequirement', 'Availability Requirement']),
    numberNeeded: rrPickV223_(row, ['numberNeeded', 'number_needed', 'qty', 'quantity', 'Number Needed']),
    priority: rrPickV223_(row, ['priority', 'Priority']),
    neededBy: rrPickV223_(row, ['neededBy', 'needed_by', 'Needed By']),
    reason: rrPickV223_(row, ['reason', 'Reason', 'additionalDetails', 'Additional Details', 'notes', 'Notes']),
    restaurantNotes: rrPickV223_(row, ['restaurantNotes', 'restaurant_notes', 'additionalDetails', 'Additional Details', 'notes', 'Notes']),
    hrNotes: rrPickV223_(row, ['hrNotes', 'hr_notes', 'last_comment', 'adminNotes', 'admin_notes']),
    adminNotes: rrPickV223_(row, ['adminNotes', 'admin_notes', 'hrNotes', 'hr_notes', 'last_comment']),
    status: status,
    timeline: timeline,
    statusHistory: timeline,
    createdAt: createdAt,
    submittedAt: createdAt,
    rawJson: JSON.stringify(row)
  };
}

function rrPickV223_(row, keys) {
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
      return row[key];
    }
  }
  return '';
}

/*******************************************************
 V300 NEWSLETTER FULL FIX
 Fixes Newsletter Management list, save, status update,
 and dashboard active newsletter visibility.
*******************************************************/
(function newsletterFullFixV300_() {
  function clean_(value) {
    return String(value == null ? '' : value).trim();
  }

  function upper_(value) {
    return clean_(value).toUpperCase();
  }

  function isNewsletterManager_(session) {
    return [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ].indexOf(session.roleCode) !== -1;
  }

  function normalizeAudienceRolesForSave_(value) {
    var text = clean_(value);
    if (!text) return '';
    var upper = upper_(text).replace(/\s+/g, '_');
    if (upper === 'ALL' || upper === 'ALL_USERS' || upper === 'ALL_USER') return '';
    return upper;
  }

  function normalizeAudienceRestaurantsForSave_(value) {
    var text = clean_(value);
    if (!text) return '';
    var upper = upper_(text).replace(/\s+/g, '_');
    if (upper === 'ALL' || upper === 'ALL_RESTAURANTS' || upper === 'ALL_RESTAURANT') return '';
    return text;
  }

  function normalizeStatusForSave_(value) {
    var upper = upper_(value || 'ACTIVE').replace(/[\s-]+/g, '_');
    if (upper === 'INACTIVE' || upper === 'DEACTIVATED' || upper === 'DISABLED') return 'INACTIVE';
    if (upper === 'DRAFT') return 'DRAFT';
    return 'ACTIVE';
  }

  function parseDateValue_(value, endOfDay) {
    var text = clean_(value);
    if (!text) return '';
    var date = new Date(text);
    if (isNaN(date.getTime())) return text;
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
      if (endOfDay) {
        date.setHours(23, 59, 59, 999);
      } else {
        date.setHours(0, 0, 0, 0);
      }
    }
    return date.toISOString();
  }

  function safeDateString_(value) {
    if (!value) return '';
    if (Object.prototype.toString.call(value) === '[object Date]') {
      return isNaN(value.getTime()) ? '' : value.toISOString();
    }
    return clean_(value);
  }

  function normalizeNewsletterRowV300_(row) {
    row = row || {};
    var postId = clean_(row.post_id || row.newsletter_id || row.newsletterId || row.postId || row.id);
    return {
      post_id: postId,
      newsletter_id: postId,
      newsletterId: postId,
      postId: postId,
      id: postId,
      title: clean_(row.title || row.newsletter_title || row.subject),
      body: clean_(row.body || row.message || row.description),
      message: clean_(row.message || row.body || row.description),
      audience_roles: clean_(row.audience_roles || row.audienceRoles || row.audience),
      audienceRoles: clean_(row.audienceRoles || row.audience_roles || row.audience),
      audience: clean_(row.audience || row.audience_roles || row.audienceRoles || 'All Users') || 'All Users',
      audience_restaurants: clean_(row.audience_restaurants || row.audienceRestaurants || row.restaurant_number || row.restaurantNumber),
      audienceRestaurants: clean_(row.audienceRestaurants || row.audience_restaurants || row.restaurant_number || row.restaurantNumber),
      restaurant_number: clean_(row.restaurant_number || row.restaurantNumber || row.audience_restaurants || row.audienceRestaurants),
      restaurantNumber: clean_(row.restaurantNumber || row.restaurant_number || row.audience_restaurants || row.audienceRestaurants),
      priority: clean_(row.priority || 'Normal') || 'Normal',
      start_at: safeDateString_(row.start_at || row.startAt || row.start_date || row.startDate),
      startAt: safeDateString_(row.startAt || row.start_at || row.start_date || row.startDate),
      start_date: safeDateString_(row.start_date || row.start_at || row.startAt || row.startDate),
      startDate: safeDateString_(row.startDate || row.start_at || row.startAt || row.start_date),
      end_at: safeDateString_(row.end_at || row.endAt || row.end_date || row.endDate),
      endAt: safeDateString_(row.endAt || row.end_at || row.end_date || row.endDate),
      end_date: safeDateString_(row.end_date || row.end_at || row.endAt || row.endDate),
      endDate: safeDateString_(row.endDate || row.end_at || row.endAt || row.end_date),
      status: clean_(row.status || 'ACTIVE') || 'ACTIVE',
      cta_label: clean_(row.cta_label || row.ctaLabel || row.attachment_name || row.attachmentName),
      cta_url: clean_(row.cta_url || row.ctaUrl || row.attachment_url || row.attachmentUrl || row.file_url),
      attachment_type: clean_(row.attachment_type || row.attachmentType),
      attachmentType: clean_(row.attachmentType || row.attachment_type),
      attachment_name: clean_(row.attachment_name || row.attachmentName || row.cta_label || row.ctaLabel),
      attachmentName: clean_(row.attachmentName || row.attachment_name || row.cta_label || row.ctaLabel),
      attachment_url: clean_(row.attachment_url || row.attachmentUrl || row.cta_url || row.ctaUrl || row.file_url),
      attachmentUrl: clean_(row.attachmentUrl || row.attachment_url || row.cta_url || row.ctaUrl || row.file_url),
      attachment_file_id: clean_(row.attachment_file_id || row.attachmentFileId),
      attachmentFileId: clean_(row.attachmentFileId || row.attachment_file_id),
      created_by: clean_(row.created_by || row.createdBy),
      createdBy: clean_(row.createdBy || row.created_by),
      created_at: safeDateString_(row.created_at || row.createdAt),
      createdAt: safeDateString_(row.createdAt || row.created_at),
      updated_at: safeDateString_(row.updated_at || row.updatedAt),
      updatedAt: safeDateString_(row.updatedAt || row.updated_at)
    };
  }

  function newsletterMatchesAudienceV300_(post, session) {
    var role = clean_(session && session.roleCode);
    var restaurantId = clean_(session && session.restaurantId);
    var rolesRaw = clean_(post.audience_roles || post.audience || post.audienceRoles);
    var restaurantsRaw = clean_(post.audience_restaurants || post.restaurant_number || post.restaurantNumber || post.audienceRestaurants);
    var rolesUpper = rolesRaw.toUpperCase().replace(/\s+/g, '_');
    var restaurantsUpper = restaurantsRaw.toUpperCase().replace(/\s+/g, '_');

    var roleMatch = !rolesRaw || rolesUpper === 'ALL' || rolesUpper === 'ALL_USERS' || rolesUpper.indexOf(role) !== -1;
    var restaurantMatch = !restaurantsRaw || restaurantsUpper === 'ALL' || restaurantsUpper === 'ALL_RESTAURANTS' || restaurantsRaw.indexOf(restaurantId) !== -1;
    return roleMatch && restaurantMatch;
  }

  function newsletterIsActiveByDateV300_(post) {
    var now = Date.now();
    var startRaw = safeDateString_(post.start_at || post.startAt || post.start_date || post.startDate);
    var endRaw = safeDateString_(post.end_at || post.endAt || post.end_date || post.endDate);
    var startOk = true;
    var endOk = true;
    if (startRaw) {
      var startTime = new Date(startRaw).getTime();
      if (!isNaN(startTime)) startOk = startTime <= now;
    }
    if (endRaw) {
      var endTime = new Date(endRaw).getTime();
      if (!isNaN(endTime)) endOk = endTime >= now;
    }
    return startOk && endOk;
  }

  NewsletterService.getActivePosts = function(session) {
    return Repository.getAll_(APP_CONFIG.sheets.NEWSLETTERS)
      .map(normalizeNewsletterRowV300_)
      .filter(function(post) {
        return normalizeStatusForSave_(post.status) === 'ACTIVE' &&
          newsletterMatchesAudienceV300_(post, session) &&
          newsletterIsActiveByDateV300_(post);
      })
      .sort(function(a, b) {
        return new Date(b.updated_at || b.created_at || b.start_at || 0).getTime() - new Date(a.updated_at || a.created_at || a.start_at || 0).getTime();
      });
  };

  this.listNewsletterRecords = function(payload) {
    var session = getSessionOrThrow_(payload && payload.sessionToken);
    if (!isNewsletterManager_(session)) throw new Error('You are not authorized to manage newsletters.');

    var rows = Repository.getAll_(APP_CONFIG.sheets.NEWSLETTERS)
      .map(normalizeNewsletterRowV300_)
      .sort(function(a, b) {
        return new Date(b.updated_at || b.created_at || b.start_at || 0).getTime() - new Date(a.updated_at || a.created_at || a.start_at || 0).getTime();
      });

    return {
      success: true,
      newsletters: rows,
      data: { newsletters: rows }
    };
  };

  this.saveNewsletterRecord = function(payload) {
    var session = getSessionOrThrow_(payload && payload.sessionToken);
    if (!isNewsletterManager_(session)) throw new Error('You are not authorized to manage newsletters.');

    var input = (payload && payload.record) || {};
    var now = new Date().toISOString();
    var newsletterId = clean_(input.newsletterId || input.newsletter_id || input.post_id || input.postId || input.id) || Utilities.getUuid();
    var title = clean_(input.title);
    var message = clean_(input.message || input.body);
    if (!title) throw new Error('Newsletter title is required.');
    if (!message) throw new Error('Newsletter message is required.');

    var attachmentUrl = clean_(input.attachmentUrl || input.attachment_url || input.cta_url || input.file_url);
    var attachmentName = clean_(input.attachmentName || input.attachment_name || input.cta_label);
    var record = {
      post_id: newsletterId,
      title: title,
      body: message,
      audience_roles: normalizeAudienceRolesForSave_(input.audience || input.audienceRoles || input.audience_roles),
      audience_restaurants: normalizeAudienceRestaurantsForSave_(input.restaurantNumber || input.restaurant_number || input.audienceRestaurants || input.audience_restaurants),
      priority: clean_(input.priority || 'Normal') || 'Normal',
      start_at: parseDateValue_(input.startDate || input.start_date || input.startAt || input.start_at, false),
      end_at: parseDateValue_(input.endDate || input.end_date || input.endAt || input.end_at, true),
      status: normalizeStatusForSave_(input.status),
      cta_label: attachmentName,
      cta_url: attachmentUrl,
      attachment_type: clean_(input.attachmentType || input.attachment_type || (attachmentUrl ? 'Drive Link' : 'None')) || 'None',
      attachment_name: attachmentName,
      attachment_url: attachmentUrl,
      attachment_file_id: clean_(input.attachmentFileId || input.attachment_file_id),
      updated_at: now
    };

    var existing = Repository.findOneBy_(APP_CONFIG.sheets.NEWSLETTERS, 'post_id', newsletterId);
    if (existing) {
      Repository.updateById_(APP_CONFIG.sheets.NEWSLETTERS, 'post_id', newsletterId, record);
    } else {
      record.created_by = session.userId;
      record.created_at = now;
      Repository.append_(APP_CONFIG.sheets.NEWSLETTERS, record);
    }

    var saved = normalizeNewsletterRowV300_(record);
    logAuditEvent_('NEWSLETTER_SAVED', 'NEWSLETTER', newsletterId, session.userId, { title: title, status: record.status });

    return {
      success: true,
      newsletterId: newsletterId,
      record: saved,
      newsletter: saved,
      data: { newsletterId: newsletterId, record: saved, newsletter: saved }
    };
  };

  this.updateNewsletterStatus = function(payload) {
    var session = getSessionOrThrow_(payload && payload.sessionToken);
    if (!isNewsletterManager_(session)) throw new Error('You are not authorized to manage newsletters.');

    var newsletterId = clean_(payload && (payload.newsletterId || payload.postId || payload.id));
    if (!newsletterId) throw new Error('Newsletter ID is required.');
    var status = normalizeStatusForSave_(payload && payload.status);

    Repository.updateById_(APP_CONFIG.sheets.NEWSLETTERS, 'post_id', newsletterId, {
      status: status,
      updated_at: new Date().toISOString()
    });

    logAuditEvent_('NEWSLETTER_STATUS_UPDATED', 'NEWSLETTER', newsletterId, session.userId, { status: status });

    return {
      success: true,
      newsletterId: newsletterId,
      status: status,
      data: { newsletterId: newsletterId, status: status }
    };
  };
})();


/*******************************************************
 V305 NEWSLETTER GLOBAL BACKEND FIX
 Keeps Newsletter Management stable when google.script.run
 needs direct global functions.
*******************************************************/
function listNewsletterRecords(payload) {
  const session = getSessionOrThrow_(payload && payload.sessionToken);
  if (!isNewsletterManagerV305_(session)) throw new Error('You are not authorized to manage newsletters.');

  const rows = Repository.getAll_(APP_CONFIG.sheets.NEWSLETTERS)
    .map(normalizeNewsletterRowV305_)
    .sort(function(a, b) {
      return new Date(b.updated_at || b.created_at || b.start_at || 0).getTime() - new Date(a.updated_at || a.created_at || a.start_at || 0).getTime();
    });

  return {
    success: true,
    newsletters: rows,
    data: { newsletters: rows }
  };
}

function saveNewsletterRecord(payload) {
  const session = getSessionOrThrow_(payload && payload.sessionToken);
  if (!isNewsletterManagerV305_(session)) throw new Error('You are not authorized to manage newsletters.');

  const input = (payload && payload.record) || {};
  const now = new Date().toISOString();
  const newsletterId = cleanNewsletterV305_(input.newsletterId || input.newsletter_id || input.post_id || input.postId || input.id) || Utilities.getUuid();
  const title = cleanNewsletterV305_(input.title);
  const message = cleanNewsletterV305_(input.message || input.body);
  if (!title) throw new Error('Newsletter title is required.');
  if (!message) throw new Error('Newsletter message is required.');

  const attachmentUrl = cleanNewsletterV305_(input.attachmentUrl || input.attachment_url || input.cta_url || input.file_url);
  const attachmentName = cleanNewsletterV305_(input.attachmentName || input.attachment_name || input.cta_label);

  const record = {
    post_id: newsletterId,
    title: title,
    body: message,
    audience_roles: normalizeNewsletterAudienceRolesV305_(input.audience || input.audienceRoles || input.audience_roles),
    audience_restaurants: normalizeNewsletterRestaurantsV305_(input.restaurantNumber || input.restaurant_number || input.audienceRestaurants || input.audience_restaurants),
    priority: cleanNewsletterV305_(input.priority || 'Normal') || 'Normal',
    start_at: parseNewsletterDateV305_(input.startDate || input.start_date || input.startAt || input.start_at, false),
    end_at: parseNewsletterDateV305_(input.endDate || input.end_date || input.endAt || input.end_at, true),
    status: normalizeNewsletterStatusV305_(input.status),
    cta_label: attachmentName,
    cta_url: attachmentUrl,
    attachment_type: cleanNewsletterV305_(input.attachmentType || input.attachment_type || (attachmentUrl ? 'Drive Link' : 'None')) || 'None',
    attachment_name: attachmentName,
    attachment_url: attachmentUrl,
    attachment_file_id: cleanNewsletterV305_(input.attachmentFileId || input.attachment_file_id),
    updated_at: now
  };

  const existing = Repository.findOneBy_(APP_CONFIG.sheets.NEWSLETTERS, 'post_id', newsletterId);
  if (existing) {
    Repository.updateById_(APP_CONFIG.sheets.NEWSLETTERS, 'post_id', newsletterId, record);
  } else {
    record.created_by = session.userId;
    record.created_at = now;
    Repository.append_(APP_CONFIG.sheets.NEWSLETTERS, record);
  }

  const saved = normalizeNewsletterRowV305_(record);
  try { logAuditEvent_('NEWSLETTER_SAVED', 'NEWSLETTER', newsletterId, session.userId, { title: title, status: record.status }); } catch (error) { console.warn('Ignored service error', error); }

  return {
    success: true,
    newsletterId: newsletterId,
    record: saved,
    newsletter: saved,
    data: { newsletterId: newsletterId, record: saved, newsletter: saved }
  };
}

function updateNewsletterStatus(payload) {
  const session = getSessionOrThrow_(payload && payload.sessionToken);
  if (!isNewsletterManagerV305_(session)) throw new Error('You are not authorized to manage newsletters.');

  const newsletterId = cleanNewsletterV305_(payload && (payload.newsletterId || payload.postId || payload.id));
  if (!newsletterId) throw new Error('Newsletter ID is required.');
  const status = normalizeNewsletterStatusV305_(payload && payload.status);

  Repository.updateById_(APP_CONFIG.sheets.NEWSLETTERS, 'post_id', newsletterId, {
    status: status,
    updated_at: new Date().toISOString()
  });

  try { logAuditEvent_('NEWSLETTER_STATUS_UPDATED', 'NEWSLETTER', newsletterId, session.userId, { status: status }); } catch (error) { console.warn('Ignored service error', error); }

  return {
    success: true,
    newsletterId: newsletterId,
    status: status,
    data: { newsletterId: newsletterId, status: status }
  };
}

function isNewsletterManagerV305_(session) {
  const role = String(session && session.roleCode || '').trim();
  return [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ].indexOf(role) !== -1;
}

function cleanNewsletterV305_(value) {
  return String(value == null ? '' : value).trim();
}

function normalizeNewsletterStatusV305_(status) {
  const value = cleanNewsletterV305_(status || 'Active').toUpperCase().replace(/[\s-]+/g, '_');
  if (value === 'INACTIVE' || value === 'DEACTIVATED' || value === 'DISABLED') return 'INACTIVE';
  if (value === 'DRAFT') return 'DRAFT';
  if (value === 'EXPIRED') return 'EXPIRED';
  return 'ACTIVE';
}

function normalizeNewsletterAudienceRolesV305_(value) {
  const raw = cleanNewsletterV305_(value);
  if (!raw || raw.toUpperCase() === 'ALL_USERS' || raw.toUpperCase() === 'ALL USERS' || raw.toUpperCase() === 'ALL') return '';
  return raw.toUpperCase().replace(/\s+/g, '_');
}

function normalizeNewsletterRestaurantsV305_(value) {
  const raw = cleanNewsletterV305_(value);
  if (!raw || raw.toUpperCase() === 'ALL_RESTAURANTS' || raw.toUpperCase() === 'ALL RESTAURANTS' || raw.toUpperCase() === 'ALL') return '';
  return raw;
}

function parseNewsletterDateV305_(value, endOfDay) {
  value = cleanNewsletterV305_(value);
  if (!value) return '';
  const date = new Date(value);
  if (isNaN(date.getTime())) return value;
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    if (endOfDay) date.setHours(23, 59, 59, 999);
    else date.setHours(0, 0, 0, 0);
  }
  return date.toISOString();
}

function normalizeNewsletterRowV305_(row) {
  row = row || {};
  const id = cleanNewsletterV305_(row.post_id || row.newsletter_id || row.id);
  const title = cleanNewsletterV305_(row.title || row.newsletter_title || row.subject);
  const body = cleanNewsletterV305_(row.body || row.message || row.description || row.details);
  const roles = cleanNewsletterV305_(row.audience_roles || row.audience || row.target_role_codes || row.target_roles);
  const restaurants = cleanNewsletterV305_(row.audience_restaurants || row.restaurant_number || row.restaurantNumber || row.target_restaurant_ids);
  const attachmentUrl = cleanNewsletterV305_(row.attachment_url || row.cta_url || row.file_url);
  const attachmentName = cleanNewsletterV305_(row.attachment_name || row.cta_label || (attachmentUrl ? 'Attachment' : ''));

  return {
    id: id,
    post_id: id,
    newsletterId: id,
    title: title,
    body: body,
    message: body,
    audience: roles ? roles.replace(/_/g, ' ') : 'All Users',
    audience_roles: roles,
    audienceRoles: roles,
    restaurantNumber: restaurants,
    restaurant_number: restaurants,
    audience_restaurants: restaurants,
    audienceRestaurants: restaurants,
    priority: cleanNewsletterV305_(row.priority || 'Normal') || 'Normal',
    startDate: cleanNewsletterV305_(row.start_at || row.startDate || row.start_date),
    start_at: cleanNewsletterV305_(row.start_at || row.startDate || row.start_date),
    endDate: cleanNewsletterV305_(row.end_at || row.endDate || row.end_date),
    end_at: cleanNewsletterV305_(row.end_at || row.endDate || row.end_date),
    status: normalizeNewsletterStatusV305_(row.status),
    cta_label: cleanNewsletterV305_(row.cta_label || attachmentName),
    cta_url: cleanNewsletterV305_(row.cta_url || attachmentUrl),
    attachmentType: cleanNewsletterV305_(row.attachment_type || (attachmentUrl ? 'Drive Link' : 'None')),
    attachment_type: cleanNewsletterV305_(row.attachment_type || (attachmentUrl ? 'Drive Link' : 'None')),
    attachmentName: attachmentName,
    attachment_name: attachmentName,
    attachmentUrl: attachmentUrl,
    attachment_url: attachmentUrl,
    attachment_file_id: cleanNewsletterV305_(row.attachment_file_id),
    created_by: cleanNewsletterV305_(row.created_by),
    createdAt: cleanNewsletterV305_(row.created_at),
    created_at: cleanNewsletterV305_(row.created_at),
    updatedAt: cleanNewsletterV305_(row.updated_at || row.created_at),
    updated_at: cleanNewsletterV305_(row.updated_at || row.created_at)
  };
}


/*******************************************************
 V306 HOST LINK MANAGEMENT SAVE FIX
 Adds direct backend functions used by ClientJS HOST Link Management.
 Keeps Newsletter v305 logic untouched.
*******************************************************/

function ensureHostLinkHeadersV306_() {
  try {
    var ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
    var sheet = ss.getSheetByName(APP_CONFIG.sheets.HOST_LINKS) || ss.insertSheet(APP_CONFIG.sheets.HOST_LINKS);
    var needed = [
      'link_id',
      'category_code',
      'title',
      'url',
      'description',
      'target_role_codes',
      'target_restaurant_ids',
      'start_at',
      'end_at',
      'status',
      'sort_order',
      'created_by',
      'created_at',
      'updated_by',
      'updated_at'
    ];

    if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
      sheet.getRange(1, 1, 1, needed.length).setValues([needed]);
      sheet.setFrozenRows(1);
      return;
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
      return String(header || '').trim();
    });

    needed.forEach(function(header) {
      if (headers.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        headers.push(header);
      }
    });
    sheet.setFrozenRows(1);
  } catch (error) {
    console.error('ensureHostLinkHeadersV306_ failed', error);
  }
}

function normalizeHostLinkRecordV306_(row) {
  row = row || {};
  var linkId = String(row.link_id || row.linkId || '').trim();
  var category = String(row.category_code || row.category || '').trim().toUpperCase();
  var audience = String(row.target_role_codes || row.targetRoleCodes || row.audience || '').trim();
  var restaurant = String(row.target_restaurant_ids || row.restaurantNumber || row.restaurant_number || row.restaurant || '').trim();
  var sortOrder = String(row.sort_order || row.sortOrder || '100').trim() || '100';
  var startAt = String(row.start_at || row.startAt || '').trim();
  var endAt = String(row.end_at || row.endAt || '').trim();

  return {
    id: linkId,
    linkId: linkId,
    link_id: linkId,
    category: category,
    category_code: category,
    title: String(row.title || '').trim(),
    url: String(row.url || '').trim(),
    description: String(row.description || '').trim(),
    audience: audience || 'All Users',
    targetRoleCodes: audience,
    target_role_codes: audience,
    restaurant: restaurant,
    restaurantNumber: restaurant,
    target_restaurant_ids: restaurant,
    startAt: startAt,
    start_at: startAt,
    endAt: endAt,
    end_at: endAt,
    status: String(row.status || 'ACTIVE').trim().toUpperCase() || 'ACTIVE',
    sortOrder: sortOrder,
    sort_order: sortOrder,
    created_by: String(row.created_by || '').trim(),
    created_at: row.created_at || '',
    updated_by: String(row.updated_by || '').trim(),
    updated_at: row.updated_at || ''
  };
}

function listHostLinkRecords(payload) {
  payload = payload || {};
  var session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
  ensureHostLinkHeadersV306_();

  var rows = Repository.getAll_(APP_CONFIG.sheets.HOST_LINKS)
    .map(normalizeHostLinkRecordV306_)
    .filter(function(row) {
      return row.title || row.url || row.category;
    })
    .sort(function(a, b) {
      return String(a.category).localeCompare(String(b.category)) ||
        Number(a.sortOrder || 100) - Number(b.sortOrder || 100) ||
        String(a.title).localeCompare(String(b.title));
    });

  return respond_(true, { hostLinks: rows });
}

function saveHostLinkRecord(payload) {
  payload = payload || {};
  var session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
  ensureHostLinkHeadersV306_();

  var input = payload.record || {};
  var linkId = String(input.linkId || input.link_id || input.id || '').trim();
  var category = String(input.category || input.category_code || '').trim().toUpperCase();
  var title = String(input.title || '').trim();
  var url = String(input.url || '').trim();
  var description = String(input.description || '').trim();
  var targetRoleCodes = String(input.targetRoleCodes || input.target_role_codes || input.audience || '').trim();
  if (targetRoleCodes === 'All Users') targetRoleCodes = '';
  var restaurantNumber = String(input.restaurantNumber || input.restaurant_number || input.restaurant || input.target_restaurant_ids || '').trim();
  var sortOrder = String(input.sortOrder || input.sort_order || '100').trim() || '100';
  var startAt = String(input.startAt || input.start_at || '').trim();
  var endAt = String(input.endAt || input.end_at || '').trim();
  var status = String(input.status || 'ACTIVE').trim().toUpperCase() || 'ACTIVE';

  if (!category) throw new Error('HOST category is required.');
  if (!title) throw new Error('Title is required.');
  if (!url) throw new Error('Link URL is required.');

  var now = new Date().toISOString();
  var record = {
    category_code: category,
    title: title,
    url: url,
    description: description,
    target_role_codes: targetRoleCodes,
    target_restaurant_ids: restaurantNumber,
    start_at: startAt,
    end_at: endAt,
    status: status,
    sort_order: sortOrder,
    updated_by: session.userId,
    updated_at: now
  };

  if (linkId) {
    Repository.updateById_(APP_CONFIG.sheets.HOST_LINKS, 'link_id', linkId, record);
    record.link_id = linkId;
    logAuditEvent_('HOST_LINK_UPDATED', 'HOST_LINK', linkId, session.userId, { title: title, category: category, status: status });
    return respond_(true, { linkId: linkId, record: normalizeHostLinkRecordV306_(record) });
  }

  linkId = Utilities.getUuid();
  record.link_id = linkId;
  record.created_by = session.userId;
  record.created_at = now;
  Repository.append_(APP_CONFIG.sheets.HOST_LINKS, record);
  logAuditEvent_('HOST_LINK_CREATED', 'HOST_LINK', linkId, session.userId, { title: title, category: category, status: status });
  return respond_(true, { linkId: linkId, record: normalizeHostLinkRecordV306_(record) });
}

function updateHostLinkStatus(payload) {
  payload = payload || {};
  var session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
  ensureHostLinkHeadersV306_();

  var linkId = String(payload.linkId || payload.link_id || payload.id || '').trim();
  if (!linkId) throw new Error('HOST link id is required.');
  var status = String(payload.status || 'ACTIVE').trim().toUpperCase() || 'ACTIVE';
  var now = new Date().toISOString();

  Repository.updateById_(APP_CONFIG.sheets.HOST_LINKS, 'link_id', linkId, {
    status: status,
    updated_by: session.userId,
    updated_at: now
  });

  logAuditEvent_('HOST_LINK_STATUS_UPDATED', 'HOST_LINK', linkId, session.userId, { status: status });
  return respond_(true, { linkId: linkId, status: status, updated_at: now });
}


/*******************************************************
 V308 CONTACT AND PASSWORD REQUEST STABILITY
 Adds direct backend functions used by the Clearview pages.
 Keeps existing ContactRequestService and PasswordChangeService logic.
*******************************************************/
function normalizeRequestStatusV308_(status) {
  var value = String(status || 'NEW').trim().toUpperCase().replace(/[\s-]+/g, '_');
  if (!value || value === 'BLANK') return 'NEW';
  if (value === 'OPEN' || value === 'NEW_REQUEST' || value === 'PENDING' || value === 'SUBMITTED') return 'NEW';
  if (value === 'VIEWED' || value === 'INPROGRESS' || value === 'WORKING' || value === 'WORK_STARTED') return 'IN_PROGRESS';
  if (value === 'DONE' || value === 'COMPLETE' || value === 'CLOSED' || value === 'RESOLVED') return 'COMPLETED';
  return value;
}

function listContactRequests(payload) {
  try {
    var session = getSessionOrThrow_(payload && payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    var rows = Repository.getAll_(APP_CONFIG.sheets.CONTACT_REQUESTS).map(function(row) {
      return {
        contact_request_id: row.contact_request_id || '',
        source: row.source || '',
        name: row.name || '',
        username: row.username || '',
        email: row.email || '',
        phone: row.phone || '',
        restaurant_number: row.restaurant_number || row.restaurant_id || '',
        message: row.message || '',
        status: normalizeRequestStatusV308_(row.status),
        action_taken_notes: row.action_taken_notes || '',
        requested_at: row.requested_at || row.created_at || '',
        updated_at: row.updated_at || '',
        updated_by: row.updated_by || ''
      };
    }).sort(function(a, b) {
      return new Date(b.requested_at || 0).getTime() - new Date(a.requested_at || 0).getTime();
    });

    return respond_(true, { requests: rows });
  } catch (error) {
    console.error('listContactRequests failed', error);
    return respond_(false, { requests: [] }, 'Could not load contact requests.');
  }
}

function listPasswordChangeRequests(payload) {
  try {
    var session = getSessionOrThrow_(payload && payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    var rows = Repository.getAll_(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS).map(function(row) {
      return {
        password_request_id: row.password_request_id || '',
        name: row.name || '',
        username: row.username || '',
        email: row.email || '',
        phone: row.phone || '',
        message: row.message || '',
        requested_at: row.requested_at || row.created_at || '',
        status: normalizeRequestStatusV308_(row.status),
        temporary_password: '',
        action_taken_notes: row.action_taken_notes || '',
        updated_at: row.updated_at || '',
        updated_by: row.updated_by || ''
      };
    }).sort(function(a, b) {
      return new Date(b.requested_at || 0).getTime() - new Date(a.requested_at || 0).getTime();
    });

    return respond_(true, { requests: rows });
  } catch (error) {
    console.error('listPasswordChangeRequests failed', error);
    return respond_(false, { requests: [] }, 'Could not load password change requests.');
  }
}

function updatePasswordChangeRequest(payload) {
  try {
    var session = getSessionOrThrow_(payload && payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN]);
    var requestId = String(payload.passwordRequestId || payload.password_request_id || '').trim();
    if (!requestId) throw new Error('Password request ID is required.');

    Repository.updateById_(APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS, 'password_request_id', requestId, {
      status: normalizeRequestStatusV308_(payload.status || 'IN_PROGRESS'),
      action_taken_notes: String(payload.actionTakenNotes || payload.action_taken_notes || '').trim(),
      updated_at: new Date().toISOString(),
      updated_by: session.userId
    });
    return respond_(true, { passwordRequestId: requestId });
  } catch (error) {
    console.error('updatePasswordChangeRequest failed', error);
    return respond_(false, null, 'Could not update password change request.');
  }
}

/*******************************************************
 V308 NEW HERO REASON FIELD CLEANUP
 Keeps reason visible for Quit and Didnt Show Up statuses.
*******************************************************/
function updateNewHeroEmployeeStatusV308_(payload) {
  var session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO]);

  var teamMemberId = String(payload.teamMemberId || payload.team_member_id || '').trim();
  if (!teamMemberId) throw new Error('Team member ID is required.');

  var status = String(payload.status || '').trim();
  var reason = String(payload.quitReason || payload.reason || payload.recruitment_reason || '').trim();
  var now = new Date().toISOString();

  var patch = {
    status: status,
    current_status: status,
    quit_reason: reason,
    no_show_reason: reason,
    recruitment_reason: reason,
    notes: reason || undefined,
    updated_at: now,
    updated_by: session.userId
  };

  Object.keys(patch).forEach(function(key) {
    if (patch[key] === undefined) delete patch[key];
  });

  Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, patch);
  return respond_(true, { teamMemberId: teamMemberId });
}

/*******************************************************
 V308 OVERRIDE NEW HERO STATUS SAVE
 Stores reason for both Quit and Didnt Show Up.
*******************************************************/
function updateNewHeroEmployeeStatus(payload) {
  try {
    const session = getSessionOrThrow_(payload.sessionToken);

    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.RESTAURANT_USER
    ]);

    ensureTeamMembersRecruitmentColumns_();

    const teamMemberId = String(payload.teamMemberId || '').trim();
    const status = String(payload.status || '').trim();
    const reason = String(payload.quitReason || payload.reason || '').trim();

    const allowedStatuses = [
      'Ready to Work',
      'Ready to Start',
      'Welcome Kit Done',
      'Work Started',
      'Didnt Show Up',
      'Quit'
    ];

    if (!teamMemberId) throw new Error('Employee ID is required.');
    if (allowedStatuses.indexOf(status) === -1) throw new Error('Invalid status selected.');
    if ((status === 'Quit' || status === 'Didnt Show Up') && !reason) {
      throw new Error('Reason is required.');
    }

    const existing = Repository.findOneBy_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId);
    if (!existing) throw new Error('Employee not found.');

    const visibleRestaurantIds = getRecruitmentRestaurantScopeIds_(session);
    if (!teamMemberInRecruitmentScope_(existing, visibleRestaurantIds)) {
      throw new Error('You do not have access to update this employee.');
    }

    let savedStatus = status;
    let outcomeStatus = '';
    let completedAt = '';
    let savedReason = '';

    if (status === 'Work Started') {
      savedStatus = 'ACTIVE';
      outcomeStatus = 'Work Started';
      completedAt = new Date().toISOString();
    }

    if (status === 'Didnt Show Up') {
      savedStatus = 'Didnt Show Up';
      outcomeStatus = 'Didnt Show Up';
      completedAt = new Date().toISOString();
      savedReason = reason;
    }

    if (status === 'Quit') {
      savedStatus = 'Quit';
      outcomeStatus = 'Quit';
      completedAt = new Date().toISOString();
      savedReason = reason;
    }

    Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', teamMemberId, {
      status: savedStatus,
      recruitment_reason: savedReason,
      recruitment_outcome_status: outcomeStatus,
      recruitment_completed_at: completedAt,
      updated_at: new Date().toISOString()
    });

    logAuditEvent_('NEW_HERO_STATUS_UPDATED', 'TEAM_MEMBER', teamMemberId, session.userId, {
      status: status,
      savedStatus: savedStatus,
      outcomeStatus: outcomeStatus,
      reason: savedReason
    });

    return respond_(true, {
      teamMemberId: teamMemberId,
      status: savedStatus,
      outcomeStatus: outcomeStatus,
      completedAt: completedAt,
      quitReason: savedReason
    });
  } catch (error) {
    console.error('updateNewHeroEmployeeStatus failed', error);
    return respond_(false, null, error && error.message ? error.message : 'Could not update employee status.');
  }
}

/*******************************************************
 V321 RESTAURANT MANAGEMENT DATA FIX
 Keeps restricted users scoped, but lets global roles load
 full Restaurant Management data reliably.
*******************************************************/
function normalizeRoleV321_(value) {
  return String(value == null ? '' : value).trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function isGlobalRestaurantManagementRoleV321_(session) {
  if (!session) return false;
  const role = normalizeRoleV321_(session.roleCode || session.role_code || session.role || session.userRole || '');
  return [
    'SUPER_ADMIN',
    'SUPERADMIN',
    'ADMIN',
    'HR_ADMIN',
    'HRADMIN',
    'HR',
    'SENIOR_MANAGEMENT',
    'SENIORMANAGEMENT',
    'SENIOR'
  ].indexOf(role) !== -1;
}

function getRestaurantsForManagementV321_(session) {
  const rows = AdminService.getRestaurants_();
  if (isGlobalRestaurantManagementRoleV321_(session)) return rows;
  return filterRestaurantsForSessionV318_(session, rows);
}

function getRestaurantManagementData(token) {
  const session = getSessionOrThrow_(token);

  if (!isGlobalRestaurantManagementRoleV321_(session)) {
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.SENIOR_MANAGEMENT,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.RESTAURANT_USER
    ]);
  }

  const restaurants = getRestaurantsForManagementV321_(session);
  const brands = BrandService.getAll();

  return respond_(true, {
    restaurants: restaurants,
    brands: brands
  });
}

try {
  if (typeof AdminService !== 'undefined') {
    AdminService.getRestaurantsForManagementV320_ = function(session) {
      return getRestaurantsForManagementV321_(session);
    };
  }
} catch (error) {
  console.error('V321 restaurant management override failed', error);
}

/*******************************************************
 V326 — PERFORMANCE + ADMIN DASHBOARD
 1. Cached boot payload
 2. Session TTL extended to 30 min
 3. Cached notifications result
 4. ensureSheet one-time flag
 5. withSheetLock timeout reduced
 6. getOperationsDashboardData — full dashboard backend
 7. OPERATIONS_DASHBOARD added to label map + role buttons
 Nothing above this line is modified.
*******************************************************/

// ── 1. Cached boot payload ────────────────────────────
function getAuthenticatedBootPayload(token) {
  try {
    const session = getSessionOrThrow_(token);
    // FIX: Use full token hash for cache key to prevent cross-user collisions
    var rawKey = String(token || '');
    var cacheHash = '';
    try {
      var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawKey);
      cacheHash = digest.slice(0, 16).map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
    } catch (e) { cacheHash = rawKey.slice(-24); }
    const cacheKey = 'boot_v326_' + cacheHash;
    const cache = CacheService.getScriptCache();
    try {
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {}
    const result = respond_(true, {
      session: sanitizeSession_(session),
      dashboard: DashboardService.getDashboardBoot(session)
    });
    try {
      const json = JSON.stringify(result);
      if (json.length < 90000) cache.put(cacheKey, json, 60);
    } catch (e) {}
    return result;
  } catch (error) {
    console.error('Authenticated boot failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

// One-time flag so ensureAdminRequestsSheetV213_ skips the sheet scan after first successful run.
var _sethAdminSheetReadyV326_ = false;

// withSheetLock_ is defined once in Repository.gs.
// Removed duplicate global copy from Services.gs to avoid Apps Script override risk.

// ── 6. Operations Dashboard backend ──────────────────
function getOperationsDashboardData(payload) {
  const token = typeof payload === 'string' ? payload : String((payload && payload.sessionToken) || '');
  const session = getSessionOrThrow_(token);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const cacheKey = 'ops_dash_v326_' + String(session.userId || '').slice(-8);
  const cache = CacheService.getScriptCache();
  if (!(payload && payload.forceRefresh)) {
    try {
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {}
  }

  // Scoping — MRO sees only assigned restaurants
  const scopeIds = getRecruitmentRestaurantScopeIds_(session);
  function inScope(restaurantId) {
    if (!scopeIds) return true;
    return scopeIds.indexOf(String(restaurantId || '')) !== -1;
  }

  const now = Date.now();
  const TWO_DAYS_MS = 2 * 24 * 60 * 60 * 1000;

  // FIX #6: Pre-fetch ALL sheets before any processing to avoid N+1 sheet reads.
  // Repository.getAll_ has a ScriptCache layer so repeated calls within TTL are free,
  // but on a cold start each is a separate spreadsheet read. Fetching upfront also
  // lets us pass raw arrays down rather than re-reading inside helper functions.
  let adminRows = [];
  let hiringRowsRaw = [];
  let teamRowsRaw = [];
  try { adminRows = typeof getRestaurantRequestAdminRowsV219_ === 'function'
    ? getRestaurantRequestAdminRowsV219_()
    : []; } catch (e) { adminRows = []; }
  try { hiringRowsRaw = Repository.getAll_(APP_CONFIG.sheets.HIRING_REQUESTS); } catch (e) { hiringRowsRaw = []; }
  try { teamRowsRaw = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS); } catch (e) { teamRowsRaw = []; }

  // Filter by scope
  adminRows = adminRows.filter(function(r) { return inScope(r.restaurantId); });

  const ATTENTION_THRESHOLD = TWO_DAYS_MS;
  let newRequests = 0, followUpNeeded = 0;
  let jobLetterTotal = 0, jobLetterPending = 0;
  let officeSupplyTotal = 0, officeSupplyAwaitingReceipt = 0;
  let completedToday = 0;
  const todayStart = new Date(); todayStart.setHours(0,0,0,0);

  // Needs attention list
  const needsAttention = [];
  // Busiest restaurants map
  const restaurantOpenCounts = {};
  // Recent activity feed
  const activityFeed = [];

  adminRows.forEach(function(r) {
    const status = String(r.status || 'New Request');
    const type = String(r.requestType || '').toUpperCase();
    const isJobLetter = type.indexOf('JOB') !== -1 || type.indexOf('LETTER') !== -1;
    const isSupply = type.indexOf('SUPPLY') !== -1 || type.indexOf('SUPPLIES') !== -1;
    const rn = String(r.restaurantNumber || r.restaurantCode || r.restaurantId || '-');
    const submittedAt = new Date(r.submittedAt || r.createdAt || 0).getTime();
    const updatedAt = new Date(r.updatedAt || r.submittedAt || 0).getTime();
    const isTerminal = (status === 'Completed' || status === 'Received');

    // Counts
    if (status === 'New Request') newRequests++;
    if (status === 'Viewed' && (now - updatedAt) > ATTENTION_THRESHOLD) followUpNeeded++;

    // Job letter
    if (isJobLetter) {
      jobLetterTotal++;
      if (!isTerminal) jobLetterPending++;
    }
    // Office supply
    if (isSupply) {
      officeSupplyTotal++;
      if (status === 'Completed') officeSupplyAwaitingReceipt++;
    }
    // Completed today
    if (isTerminal && updatedAt >= todayStart.getTime()) completedToday++;

    // Open restaurant count
    if (!isTerminal) {
      restaurantOpenCounts[rn] = (restaurantOpenCounts[rn] || 0) + 1;
    }

    // Needs attention — open requests older than 2 days
    if (!isTerminal && (now - submittedAt) > ATTENTION_THRESHOLD) {
      const ageDays = Math.floor((now - submittedAt) / 86400000);
      needsAttention.push({
        requestId: r.requestId || '',
        label: (isJobLetter ? 'Job letter' : isSupply ? 'Supplies' : 'Request') + ' #' + String(r.requestId || '').slice(-6).toUpperCase(),
        restaurant: rn,
        type: isJobLetter ? 'JOB_LETTER' : 'OFFICE_SUPPLIES',
        status: status,
        ageDays: ageDays,
        submittedAt: r.submittedAt || r.createdAt || ''
      });
    }

    // Activity feed — from status_history_json
    const hist = Array.isArray(r.statusHistory) ? r.statusHistory : [];
    hist.forEach(function(h) {
      if (!h || !h.at) return;
      activityFeed.push({
        at: h.at,
        by: String(h.by || ''),
        status: String(h.status || ''),
        note: String(h.note || ''),
        restaurant: rn,
        requestId: r.requestId || '',
        type: isJobLetter ? 'Job Letter' : isSupply ? 'Office Supplies' : 'Admin Request',
        requestType: isJobLetter ? 'JOB_LETTER' : 'OFFICE_SUPPLIES'
      });
    });
  });

  // Sort needs attention by age descending
  needsAttention.sort(function(a, b) { return b.ageDays - a.ageDays; });

  // Busiest restaurants — top 5 open
  const busiestRestaurants = Object.keys(restaurantOpenCounts)
    .map(function(rn) { return { restaurant: rn, openCount: restaurantOpenCounts[rn] }; })
    .sort(function(a, b) { return b.openCount - a.openCount; })
    .slice(0, 5);

  // ── Hiring requests ─────────────────────────────── (uses pre-fetched hiringRowsRaw)
  let hiringRows = [];
  try { hiringRows = hiringRowsRaw.filter(function(r) { return inScope(r.restaurantId); }); } catch (e) {}

  let hiringActive = 0, hiringDone = 0;
  hiringRows.forEach(function(r) {
    const s = String(r.status || '').toUpperCase();
    if (s === 'HIRING_COMPLETE' || s === 'COMPLETED' || s === 'HIRING COMPLETE') hiringDone++;
    else hiringActive++;
    // Add to activity
    const hist = Array.isArray(r.statusHistory || r.timeline) ? (r.statusHistory || r.timeline) : [];
    hist.forEach(function(h) {
      if (!h || !h.at) return;
      activityFeed.push({
        at: h.at,
        by: String(h.by || ''),
        status: String(h.status || h.action || ''),
        note: String(h.note || ''),
        restaurant: String(r.restaurantNumber || r.restaurantCode || '-'),
        requestId: r.requestId || '',
        type: 'Hiring Request',
        requestType: 'HIRING'
      });
    });
  });

  // ── Team members / New Heroes ─────────────────────  (uses pre-fetched teamRowsRaw)
  let teamRows = [];
  try { teamRows = teamRowsRaw.filter(function(r) { return inScope(r.restaurant_id); }); } catch (e) {}

  const onboardingStatuses = ['Ready to Work','Ready to Start','Welcome Kit Done'];
  let newHeroesOnboarding = 0, workStartedThisMonth = 0, noShowQuit = 0;
  const thisMonthStart = new Date(); thisMonthStart.setDate(1); thisMonthStart.setHours(0,0,0,0);

  teamRows.forEach(function(r) {
    const status = String(r.status || '');
    const outcomeStatus = String(r.recruitment_outcome_status || '');
    if (onboardingStatuses.indexOf(status) !== -1) newHeroesOnboarding++;
    const completedAt = new Date(r.recruitment_completed_at || 0).getTime();
    if ((status === 'ACTIVE' || outcomeStatus === 'Work Started') && completedAt >= thisMonthStart.getTime()) workStartedThisMonth++;
    if (outcomeStatus === 'Didnt Show Up' || outcomeStatus === 'Quit' || status === 'Didnt Show Up' || status === 'Quit') {
      const quitAt = new Date(r.recruitment_completed_at || r.updated_at || 0).getTime();
      if (now - quitAt < 30 * 24 * 60 * 60 * 1000) noShowQuit++;
    }
  });

  // Sort activity feed — most recent first, cap at 20
  activityFeed.sort(function(a, b) { return new Date(b.at) - new Date(a.at); });
  const recentActivity = activityFeed.slice(0, 20);

  const result = respond_(true, {
    adminRequests: {
      newRequests: newRequests,
      followUpNeeded: followUpNeeded,
      jobLetterTotal: jobLetterTotal,
      jobLetterPending: jobLetterPending,
      officeSupplyTotal: officeSupplyTotal,
      officeSupplyAwaitingReceipt: officeSupplyAwaitingReceipt,
      completedToday: completedToday
    },
    hiring: {
      active: hiringActive,
      done: hiringDone,
      total: hiringActive + hiringDone
    },
    team: {
      newHeroesOnboarding: newHeroesOnboarding,
      workStartedThisMonth: workStartedThisMonth,
      noShowQuitLast30Days: noShowQuit
    },
    busiestRestaurants: busiestRestaurants,
    needsAttention: needsAttention.slice(0, 10),
    recentActivity: recentActivity,
    generatedAt: new Date().toISOString(),
    scopeLabel: scopeIds ? 'Your assigned restaurants' : 'All restaurants'
  });

  try {
    const json = JSON.stringify(result);
    if (json.length < 90000) cache.put(cacheKey, json, 60);
  } catch (e) {}

  return result;
}

// ── 7. Label map + role button additions ──────────────
(function patchRoleActionsV326_() {
  const origFormat = formatActionLabel_;
  formatActionLabel_ = function(key) {
    if (key === 'OPERATIONS_DASHBOARD') return 'Operations Dashboard';
    return origFormat(key);
  };
})();

/*******************************************************
 V327 — THREE FIXES
 1. getAdminRequestWorkspaceData — scopes requests to
    restaurant for RESTAURANT_USER (data privacy fix)
 2. getTopMetrics — real counts instead of hardcoded
 3. Email notifications on request status completion
 Nothing above this line is modified.
*******************************************************/

// ── Helper guards — defined here in case V325/V326 additions
// are not present in this Services file. If already defined
// above they are overwritten harmlessly (GAS uses last definition).
function sethCacheKey_(name) {
  return 'seth_v325_' + name;
}

function sethInvalidateRequestCaches_() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll([
      sethCacheKey_('rr_receiver_data'),
      sethCacheKey_('admin_workspace_data'),
      sethCacheKey_('last_updated'),
      // FIX v500 #1: Also bust new getAdminRequestRows + getHiringRequestRows caches
      'v500_admin_req_rows',
      'v500_hiring_rows_all'
    ]);
    cache.put(sethCacheKey_('last_updated'), String(Date.now()), 600);
  } catch (e) {
    console.warn('sethInvalidateRequestCaches_ failed', e);
  }
}

// ── 1. Scoped getAdminRequestWorkspaceData ────────────
function getAdminRequestWorkspaceData(token) {
  const forceRefresh = (typeof token === 'object') && token && token.forceRefresh;
  const actualToken = (typeof token === 'object') ? String((token && token.sessionToken) || '') : String(token || '');
  const session = getSessionOrThrow_(actualToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const cacheKey = sethCacheKey_('admin_workspace_data_' + String(session.userId || '').slice(-8));
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    try {
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {}
  }

  const restaurants = AdminService.getRestaurantsForSession_(session);
  const restaurantLookup = {};
  restaurants.forEach(function(r) {
    restaurantLookup[String(r.restaurantId || '')] = r;
  });

  const teamMembers = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS).map(function(row) {
    const restaurantIds = String(row.restaurant_id || '')
      .split(',').map(function(i) { return i.trim(); }).filter(Boolean);
    const restaurantCodes = restaurantIds.map(function(id) {
      const r = restaurantLookup[id] || {};
      return String(r.restaurantCode || r.restaurant_number || id || '').trim();
    }).filter(Boolean);
    return {
      teamMemberId: row.team_member_id || row.employee_id || '',
      employeeId: row.employee_id || '',
      name: row.full_name || '',
      mobile: row.phone || '',
      status: row.status || '',
      jobTitle: row.job_title || '',
      employmentType: row.employment_type || '',
      hiredDate: row.hire_date || '',
      restaurantIds: restaurantIds,
      restaurantCodes: restaurantCodes,
      restaurantCode: restaurantCodes.join(', ')
    };
  });

  // Scope requests to restaurant for RESTAURANT_USER
  let allRequests = getAdminRequestRowsV213_();
  const isRestaurantUser = session.roleCode === APP_CONFIG.roles.RESTAURANT_USER;
  if (isRestaurantUser) {
    const scopeIds = getSessionRestaurantScopeV318_(session);
    if (scopeIds && scopeIds.length) {
      allRequests = allRequests.filter(function(r) {
        return rowMatchesRestaurantScopeV318_(r, scopeIds);
      });
    }
  }

  const result = respond_(true, {
    restaurants: restaurants,
    employees: teamMembers,
    requests: allRequests
  });

  try {
    const json = JSON.stringify(result);
    if (json.length < 90000) cache.put(cacheKey, json, 90);
  } catch (e) {}

  return result;
}


// V500 SPEED FIX: Dead first getTopMetrics override (V327) removed.
// The active V329 override below is the only definition now.

// ── 3. Email notifications on status completion ────────
function sendAdminRequestStatusEmail_(requestRow, newStatus, session) {
  try {
    if (!requestRow || !newStatus) return;
    const isCompletion = newStatus === 'Completed' || newStatus === 'Received';
    if (!isCompletion) return;

    // Find the email of the user who submitted the request
    const requestedById = String(requestRow.requestedBy || requestRow.requested_by || '');
    if (!requestedById) return;

    const user = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'user_id', requestedById);
    const email = user && String(user.email || '').trim();
    if (!email || email.indexOf('@') === -1) return;

    const type = String(requestRow.requestType || requestRow.request_type || '').toUpperCase();
    const isJobLetter = type.indexOf('JOB') !== -1 || type.indexOf('LETTER') !== -1;
    const isSupply = type.indexOf('SUPPLY') !== -1 || type.indexOf('SUPPLIES') !== -1;
    const restaurantCode = String(requestRow.restaurantNumber || requestRow.restaurant_number || requestRow.restaurantCode || '');
    const requestId = String(requestRow.requestId || requestRow.request_id || '').slice(-8).toUpperCase();

    let subject = '';
    let body = '';

    if (isJobLetter && newStatus === 'Completed') {
      subject = '[Saisethsons] Your job letter request is ready — ' + restaurantCode;
      body = 'Hello,\n\nYour job letter request (#' + requestId + ') for restaurant ' + restaurantCode + ' has been completed.\n\nPlease log in to the portal and go to Admin Request → Manage Requests to download your letter.\n\nSaisethsons Hospitality Group';
    } else if (isSupply && newStatus === 'Completed') {
      subject = '[Saisethsons] Office supplies request completed — ' + restaurantCode;
      body = 'Hello,\n\nYour office supplies request (#' + requestId + ') for restaurant ' + restaurantCode + ' has been marked as completed.\n\nPlease log in to the portal and go to Admin Request → Manage Requests to confirm receipt.\n\nSaisethsons Hospitality Group';
    } else if (newStatus === 'Received') {
      return; // No email needed for Received — it is triggered by the restaurant themselves
    } else {
      return;
    }

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });

    console.log('Notification email sent to ' + email + ' for request ' + requestId + ' status ' + newStatus);
  } catch (e) {
    // Email failures must never block the main save
    console.error('sendAdminRequestStatusEmail_ failed', e);
  }
}

// Patch updateAdminRequestReceiverStatusV219_ to send email on completion
function updateAdminRequestReceiverStatusV219_(session, requestId, status, notes, quantity) {
  const sheet = ensureAdminRequestsSheetV219_();
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Admin request sheet is empty.');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idIndex = headers.indexOf('request_id');
  const statusIndex = headers.indexOf('status');
  const notesIndex = headers.indexOf('admin_notes');
  const qtyIndex = headers.indexOf('quantity');
  const updatedByIndex = headers.indexOf('status_updated_by');
  const updatedAtIndex = headers.indexOf('status_updated_at');
  const updatedIndex = headers.indexOf('updated_at');
  const historyIndex = headers.indexOf('status_history_json');
  if (idIndex === -1) throw new Error('request_id column not found.');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

    const now = new Date().toISOString();
    let history = [];
    try { history = historyIndex !== -1 ? JSON.parse(values[r][historyIndex] || '[]') : []; } catch (e) {}
    history.push({
      at: now,
      by: session.fullName || session.username || session.userId || '',
      status: status,
      note: notes || ''
    });

    if (statusIndex !== -1) values[r][statusIndex] = status;
    if (notesIndex !== -1 && notes) values[r][notesIndex] = notes;
    if (qtyIndex !== -1 && quantity !== undefined && quantity !== '') values[r][qtyIndex] = quantity;
    if (updatedByIndex !== -1) values[r][updatedByIndex] = session.fullName || session.username || session.userId || '';
    if (updatedAtIndex !== -1) values[r][updatedAtIndex] = now;
    if (updatedIndex !== -1) values[r][updatedIndex] = now;
    if (historyIndex !== -1) values[r][historyIndex] = JSON.stringify(history);

    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);

    // Build row object for email
    const rowObj = {};
    headers.forEach(function(h, i) { rowObj[h] = values[r][i]; });
    rowObj.requestId = rowObj.request_id || '';
    rowObj.requestType = rowObj.request_type || '';
    rowObj.requestedBy = rowObj.requested_by || '';
    rowObj.restaurantNumber = rowObj.restaurant_number || '';

    logAuditEvent_('ADMIN_REQUEST_STATUS_UPDATED', 'ADMIN_REQUEST', requestId, session.userId, { status: status });
    sethInvalidateRequestCaches_();

    // Send email notification — fire and forget
    sendAdminRequestStatusEmail_(rowObj, status, session);

    return respond_(true, { requestId: requestId, status: status });
  }
  throw new Error('Admin request not found: ' + requestId);
}

/*******************************************************
 V328 — ROLE BUTTON PERMISSIONS FIX
 Adds a lightweight endpoint that returns ONLY the
 permission matrix. Called by the Role Action Button
 Control page on open so it always loads fresh data
 from the sheet instead of relying on state.adminTools
 which may not be populated.
*******************************************************/
function getRoleButtonPermissions(token) {
  const session = getSessionOrThrow_(token);
  assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);

  const cacheKey = 'role_btn_perms_v13_super_admin_all';
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) return respond_(true, { buttonPermissions: JSON.parse(cached) });
  } catch (e) {}

  const records = Repository.getAll_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS);
  const matrix = {};

  Object.keys(APP_CONFIG.roles).forEach(function(key) {
    const roleCode = APP_CONFIG.roles[key];
    const roleRecords = records.filter(function(record) {
      return String(record.role_code || '') === String(roleCode);
    });

    // Super Admin always has all role action buttons.
    // It should show all checkboxes selected in Role Action Button Control.
    if (String(roleCode) === String(APP_CONFIG.roles.SUPER_ADMIN)) {
      matrix[roleCode] = RoleActionService.normalizeActionKeysV367_(ALL_ROLE_ACTION_BUTTONS || []);
      return;
    }

    if (!roleRecords.length) {
      matrix[roleCode] = RoleActionService.normalizeActionKeysV367_(
        DEFAULT_ROLE_ACTION_PERMISSIONS[roleCode] || []
      );
      return;
    }

    matrix[roleCode] = RoleActionService.applyLegacyPayrollAdminBackfillV367_(roleCode, roleRecords
      .filter(function(record) {
        return String(record.is_allowed || '').toUpperCase() === 'TRUE' &&
          String(record.button_key || '').trim() !== ROLE_ACTION_NO_ACCESS_MARKER;
      })
      .map(function(record) {
        return String(record.button_key || '').trim();
      }));
  });

  try { cache.put(cacheKey, JSON.stringify(matrix), 30); } catch (e) {}
  return respond_(true, { buttonPermissions: matrix });
}

/*******************************************************
 V329 — CACHE FIXES
 1. getRestaurantRequestReceiverData — proper cached version
    (definitions at 3883/4303/4854 are uncached.
    GAS uses last definition so this wins.)
 2. getTopMetrics — 60s per-role cache
 Both invalidated by sethInvalidateRequestCaches_.
*******************************************************/

function getRestaurantRequestReceiverData(payload) {
  const token = typeof payload === 'string' ? payload : String((payload && payload.sessionToken) || '');
  const forceRefresh = payload && payload.forceRefresh;
  const session = getSessionOrThrow_(token);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const cacheKey = sethCacheKey_('rr_receiver_data_v2');
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    try {
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {}
  }

  let hiringRequests = [];
  let adminRequests = [];

  try {
    hiringRequests = getHiringRequestRows_().map(function(row) {
      return mapRestaurantRequestHiringRowV223_(row);
    });
  } catch (e) {
    console.error('Hiring request load failed', e);
  }

  try {
    adminRequests = typeof getRestaurantRequestAdminRowsV219_ === 'function'
      ? getRestaurantRequestAdminRowsV219_()
      : getAdminRequestRowsV213_();
  } catch (e) {
    console.error('Admin request load failed', e);
  }

  const result = respond_(true, {
    hiringRequests: hiringRequests,
    adminRequests: adminRequests
  });

  try {
    const json = JSON.stringify(result);
    if (json.length < 90000) cache.put(cacheKey, json, 90);
  } catch (e) {}

  return result;
}

DashboardService.getTopMetrics = function(session) {
  const role = session.roleCode;
  const cacheKey = 'metrics_v329_' + String(role || '') + '_' + String(session.restaurantId || 'all').slice(-6);
  const cache = CacheService.getScriptCache();
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  let result;
  try {
    const scopeIds = getSessionRestaurantScopeV318_(session);
    function inScope(row) { return rowMatchesRestaurantScopeV318_(row, scopeIds); }

    if (role === APP_CONFIG.roles.RESTAURANT_USER) {
      const adminRows = getAdminRequestRowsV213_().filter(inScope);
      const openRequests = adminRows.filter(function(r) {
        const s = String(r.status || '').toUpperCase();
        return s !== 'COMPLETED' && s !== 'RECEIVED' && s !== 'CANCELLED';
      }).length;
      const lettersReady = adminRows.filter(function(r) {
        return String(r.status || '') === 'Completed' &&
          String(r.requestType || '').toUpperCase().indexOf('JOB') !== -1;
      }).length;
      let hiresPending = 0;
      try {
        hiresPending = getHiringRequestRows_().filter(function(r) {
          const s = String(r.status || '').toUpperCase();
          return s !== 'HIRING_COMPLETE' && s !== 'COMPLETED' && s !== 'CANCELLED';
        }).length;
      } catch (e) {}
      result = { openRequests: openRequests, lettersReady: lettersReady, hiresPending: hiresPending };

    } else if (
      role === APP_CONFIG.roles.HR_ADMIN ||
      role === APP_CONFIG.roles.MRO ||
      role === APP_CONFIG.roles.SENIOR_MANAGEMENT
    ) {
      const adminRows = getAdminRequestRowsV213_();
      let openHiring = 0;
      try {
        openHiring = getHiringRequestRows_().filter(function(r) {
          const s = String(r.status || '').toUpperCase();
          return s !== 'HIRING_COMPLETE' && s !== 'COMPLETED' && s !== 'CANCELLED';
        }).length;
      } catch (e) {}
      const newAdminRequests = adminRows.filter(function(r) {
        return String(r.status || '') === 'New Request';
      }).length;
      const followUpNeeded = adminRows.filter(function(r) {
        return String(r.status || '') === 'Viewed' &&
          (Date.now() - new Date(r.updatedAt || r.createdAt || 0).getTime()) > 2 * 24 * 60 * 60 * 1000;
      }).length;
      result = { openHiringRequests: openHiring, newAdminRequests: newAdminRequests, followUpNeeded: followUpNeeded };

    } else {
      let openHiring = 0;
      try {
        openHiring = getHiringRequestRows_().filter(function(r) {
          const s = String(r.status || '').toUpperCase();
          return s !== 'HIRING_COMPLETE' && s !== 'COMPLETED' && s !== 'CANCELLED';
        }).length;
      } catch (e) {}
      const adminRows = getAdminRequestRowsV213_();
      const restaurants = AdminService.getRestaurants_();
      result = {
        totalRestaurants: restaurants.filter(function(r) {
          return String(r.status || '').toUpperCase() === 'ACTIVE';
        }).length,
        openAdminRequests: adminRows.filter(function(r) {
          const s = String(r.status || '').toUpperCase();
          return s !== 'COMPLETED' && s !== 'RECEIVED' && s !== 'CANCELLED';
        }).length,
        openHiringRequests: openHiring,
        newRequestsToday: adminRows.filter(function(r) {
          const d = new Date(r.createdAt || 0);
          const today = new Date(); today.setHours(0,0,0,0);
          return d >= today;
        }).length
      };
    }
  } catch (e) {
    console.error('getTopMetrics failed', e);
    result = { openRequests: 0, hiresPending: 0 };
  }

  try { cache.put(cacheKey, JSON.stringify(result), 60); } catch (e) {}
  return result;
};

/*******************************************************
 V333 — RECEIVED CONFIRMATION FOR ADMIN REQUESTS
 Restaurant confirms receipt after admin marks a request completed.
 This keeps Completed and Received as two different steps.
*******************************************************/
function confirmOfficeSupplyReceived(payload) {
  payload = payload || {};
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const requestId = String(payload.requestId || '').trim();
  if (!requestId) throw new Error('Missing request id.');

  return withSheetLock_(function() {
    const sheet = ensureAdminRequestsSheetV213_();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) throw new Error('Admin request sheet is empty.');

    const headers = values[0].map(function(header) {
      return String(header || '').trim();
    });

    const idIndex = headers.indexOf('request_id');
    const statusIndex = headers.indexOf('status');
    const updatedByIndex = headers.indexOf('status_updated_by');
    const updatedAtIndex = headers.indexOf('status_updated_at');
    const updatedIndex = headers.indexOf('updated_at');
    const historyIndex = headers.indexOf('status_history_json');

    if (idIndex === -1) throw new Error('request_id column not found.');
    if (statusIndex === -1) throw new Error('status column not found.');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIndex] || '') !== requestId) continue;

      const currentStatus = String(values[r][statusIndex] || '').trim();
      if (currentStatus === 'Received') {
        return respond_(true, { requestId: requestId, status: 'Received', alreadyReceived: true });
      }

      const now = new Date().toISOString();
      let history = [];
      try {
        history = historyIndex !== -1 ? JSON.parse(values[r][historyIndex] || '[]') : [];
      } catch (error) {
        history = [];
      }
      if (!Array.isArray(history)) history = [];

      history.push({
        at: now,
        by: session.fullName || session.username || session.userId || '',
        status: 'Received',
        note: 'Restaurant confirmed received'
      });

      values[r][statusIndex] = 'Received';
      if (updatedByIndex !== -1) values[r][updatedByIndex] = session.fullName || session.username || session.userId || '';
      if (updatedAtIndex !== -1) values[r][updatedAtIndex] = now;
      if (updatedIndex !== -1) values[r][updatedIndex] = now;
      if (historyIndex !== -1) values[r][historyIndex] = JSON.stringify(history);

      sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);

      try { logAuditEvent_('ADMIN_REQUEST_RECEIVED_CONFIRMED', 'ADMIN_REQUEST', requestId, session.userId, { status: 'Received' }, session); } catch (e) {}
      try { sethInvalidateRequestCaches_(); } catch (e2) {}

      return respond_(true, { requestId: requestId, status: 'Received' });
    }

    throw new Error('Request not found.');
  });
}


// ═══════════════════════════════════════════════════════════════════════
// PAYROLL MODULE V340
// Restaurant side: Payroll Submission
// Admin side: Payroll Review
// ═══════════════════════════════════════════════════════════════════════


function payrollEmployeeRestaurantKeysV366_(row) {
  row = row || {};
  const values = [
    row.restaurant_number,
    row.restaurantNumber,
    row.restaurant_code,
    row.restaurantCode,
    row.assigned_restaurant,
    row.assignedRestaurant,
    row.assigned_restaurant_number,
    row.assignedRestaurantNumber,
    row.restaurant_id,
    row.restaurantId
  ];
  const seen = {};
  return values.map(function(v) { return cleanPayrollTextV340_(v); })
    .filter(function(v) { if (!v || seen[v]) return false; seen[v] = true; return true; });
}

function payrollNormalizeRestaurantKeyV366_(value) {
  return cleanPayrollTextV340_(value).replace(/[^0-9A-Za-z]/g, '').toUpperCase();
}

function payrollEmployeeMatchesRestaurantV366_(row, restaurantNumber) {
  const target = payrollNormalizeRestaurantKeyV366_(restaurantNumber);
  if (!target) return true;
  return payrollEmployeeRestaurantKeysV366_(row).some(function(v) {
    return payrollNormalizeRestaurantKeyV366_(v) === target;
  });
}

function payrollEmployeeIsActiveV366_(row) {
  const status = String((row && (row.status || row.employee_status)) || 'ACTIVE').toUpperCase().trim();
  return ['INACTIVE', 'ARCHIVED', 'QUIT', 'DID NOT SHOW UP', 'DIDNT SHOW UP', 'TERMINATED', 'DEACTIVATED'].indexOf(status) === -1;
}

function payrollMapEmployeeV366_(row) {
  const fullName = cleanPayrollTextV340_(row.full_name || row.employee_name || row.name || row.team_member_name);
  const parts = fullName.split(' ');
  const keys = payrollEmployeeRestaurantKeysV366_(row);
  return {
    employee_id: cleanPayrollTextV340_(row.employee_id || row.team_member_id || row.id || fullName),
    full_name: fullName,
    first_name: cleanPayrollTextV340_(row.first_name || row.firstName || parts[0] || ''),
    last_name: cleanPayrollTextV340_(row.last_name || row.lastName || parts.slice(1).join(' ')),
    restaurant_number: cleanPayrollTextV340_(row.restaurant_number || row.restaurantNumber || row.assigned_restaurant || row.assignedRestaurant || keys[0] || ''),
    restaurant_id: cleanPayrollTextV340_(row.restaurant_id || row.restaurantId || ''),
    restaurant_code: cleanPayrollTextV340_(row.restaurant_code || row.restaurantCode || ''),
    assigned_restaurant: cleanPayrollTextV340_(row.assigned_restaurant || row.assignedRestaurant || ''),
    restaurant_keys: keys.join('|'),
    job_title: cleanPayrollTextV340_(row.job_title || row.position),
    status: cleanPayrollTextV340_(row.status || row.employee_status || 'ACTIVE')
  };
}

function legacyGetPayrollSubmissionWorkspaceDataV340_(sessionToken) {
  try {
    const session = getSessionOrThrow_(sessionToken);
    const isMRORole = [
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ].indexOf(session.roleCode) !== -1;

    // Build restaurant list — MRO/HR sees all active restaurants they are scoped to,
    // restaurant users see only their own.
    const allRestaurants = Repository.getAll_(APP_CONFIG.sheets.RESTAURANTS)
      .filter(function(r) { return String(r.status || '').toUpperCase() !== 'INACTIVE'; });
    const scopeIds = isMRORole ? getRecruitmentRestaurantScopeIds_(session) : null;
    const restaurants = allRestaurants
      .filter(function(r) {
        if (!scopeIds) return true; // SUPER_ADMIN sees all
        return scopeIds.indexOf(String(r.restaurant_id || '')) !== -1;
      })
      .map(function(r) {
        return {
          restaurant_id: cleanPayrollTextV340_(r.restaurant_id),
          restaurant_number: cleanPayrollTextV340_(r.restaurant_number || r.restaurant_code),
          restaurant_name: cleanPayrollTextV340_(r.restaurant_name)
        };
      });

    const restaurantNumber = getPrimaryRestaurantNumberForSessionV340_(session);

    // Employees: MRO/HR gets all active employees across their restaurants;
    // restaurant users get only their restaurant's employees.
    const allTeam = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS);
    const scopedRns = restaurants.map(function(r) { return String(r.restaurant_number || ''); });
    const employees = allTeam.filter(function(row) {
      if (!payrollEmployeeIsActiveV366_(row)) return false;
      if (isMRORole) {
        if (!scopedRns.length) return true;
        return scopedRns.some(function(rn) { return payrollEmployeeMatchesRestaurantV366_(row, rn); });
      }
      return payrollEmployeeMatchesRestaurantV366_(row, restaurantNumber);
    }).map(payrollMapEmployeeV366_).sort(function(a, b) { return a.full_name.localeCompare(b.full_name); });

    const submissions = getPayrollSubmissionsForSessionV340_(session, restaurantNumber);

    return respond_(true, {
      session: sanitizeSession_(session),
      restaurantNumber: restaurantNumber,
      restaurants: restaurants,
      employees: employees,
      submissions: submissions,
      isMRORole: isMRORole
    });
  } catch (error) {
    console.error('getPayrollSubmissionWorkspaceData failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function legacyGetPayrollReviewWorkspaceDataV340_(sessionToken) {
  try {
    const session = getSessionOrThrow_(sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const submissions = getPayrollSubmissionsForAdminV340_();
    const restaurants = Repository.getAll_(APP_CONFIG.sheets.RESTAURANTS)
      .filter(function(row) { return String(row.status || '').toUpperCase() !== 'INACTIVE'; })
      .map(function(row) {
        return {
          restaurant_id: cleanPayrollTextV340_(row.restaurant_id),
          restaurant_number: cleanPayrollTextV340_(row.restaurant_number || row.restaurant_code),
          restaurant_name: cleanPayrollTextV340_(row.restaurant_name)
        };
      });

    // Include employees so MRO edit mode has the dropdown populated
    const allTeam = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS);
    const scopeIds2 = getRecruitmentRestaurantScopeIds_(session);
    const rnSet = restaurants.map(function(r) { return String(r.restaurant_number || ''); });
    const employees = allTeam.filter(function(row) {
      if (!payrollEmployeeIsActiveV366_(row)) return false;
      if (!scopeIds2) return true; // SUPER_ADMIN sees all
      return rnSet.some(function(rn) { return payrollEmployeeMatchesRestaurantV366_(row, rn); });
    }).map(payrollMapEmployeeV366_).sort(function(a, b) { return a.full_name.localeCompare(b.full_name); });

    return respond_(true, {
      session: sanitizeSession_(session),
      restaurants: restaurants,
      employees: employees,
      submissions: submissions
    });
  } catch (error) {
    console.error('getPayrollReviewWorkspaceData failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function legacySavePayrollSubmissionDraftV340_(payload) {
  return upsertPayrollSubmissionV340_(payload, 'DRAFT');
}

function legacySubmitPayrollSubmissionV340_(payload) {
  return upsertPayrollSubmissionV340_(payload, 'SUBMITTED');
}

function legacyUpdatePayrollSubmissionStatusV340_(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const submissionId = cleanPayrollTextV340_(payload.submissionId);
    const status = normalizePayrollStatusV340_(payload.status || 'IN_REVIEW');
    if (!submissionId) throw new Error('Submission ID is required.');

    const now = new Date().toISOString();
    const rows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions');
    const current = rows.find(function(row) { return String(row.submission_id) === String(submissionId); });
    if (!current) throw new Error('Payroll submission not found.');

    // FIX 4: Validate status transition
    const currentStatus = String(current.status || '').toUpperCase();
    const ALLOWED_TRANSITIONS = {
      'SUBMITTED':         ['IN_REVIEW','REJECTED'],
      'IN_REVIEW':         ['APPROVED','CORRECTION_NEEDED','SUBMITTED','REJECTED'],
      'RESUBMITTED':       ['IN_REVIEW','APPROVED','CORRECTION_NEEDED','REJECTED'],
      'CORRECTION_NEEDED': ['IN_REVIEW'],
      'APPROVED':          [],
      'REJECTED':          [],
    };
    const allowed = ALLOWED_TRANSITIONS[currentStatus] || [];
    if (allowed.length > 0 && allowed.indexOf(status) === -1) {
      throw new Error('Invalid status transition: ' + currentStatus + ' → ' + status + '.');
    }

    const history = parsePayrollJsonV340_(current.status_history_json, []);
    history.push({
      status: status,
      note: cleanPayrollTextV340_(payload.adminNotes || payload.note || ''),
      updated_by: session.fullName || session.username || session.userId,
      updated_at: now
    });

    const patch = {
      status: status,
      admin_notes: cleanPayrollTextV340_(payload.adminNotes || current.admin_notes || ''),
      correction_notes: cleanPayrollTextV340_(payload.correctionNotes || current.correction_notes || ''),
      reviewed_by: session.fullName || session.username || session.userId,
      reviewed_at: current.reviewed_at || now,
      approved_at: status === 'APPROVED' ? now : current.approved_at || '',
      updated_at: now,
      updated_by: session.fullName || session.username || session.userId,
      status_history_json: JSON.stringify(history)
    };

    Repository.updateById_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions', 'submission_id', submissionId, patch);
    logAuditEvent_('PAYROLL_STATUS_UPDATED', 'PAYROLL', submissionId, session.userId, { status: status }, session);
    invalidatePayrollCaches_(session.userId);
    return respond_(true, { submissionId: submissionId, status: status });
  } catch (error) {
    console.error('updatePayrollSubmissionStatus failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function upsertPayrollSubmissionV340_(payload, requestedStatus) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO
    ]);

    const now = new Date().toISOString();
    const restaurantNumber = cleanPayrollTextV340_(payload.restaurantNumber || getPrimaryRestaurantNumberForSessionV340_(session));
    if (!restaurantNumber) throw new Error('Restaurant number is required.');

    const payPeriodEnding = cleanPayrollTextV340_(payload.payPeriodEnding);
    if (requestedStatus === 'SUBMITTED' && !payPeriodEnding) throw new Error('Pay period ending is required.');

    const submissionId = cleanPayrollTextV340_(payload.submissionId) || Utilities.getUuid();
    const status = requestedStatus === 'DRAFT' ? 'DRAFT' : 'SUBMITTED';
    const existingRows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions');
    const existing = existingRows.find(function(row) { return String(row.submission_id) === String(submissionId); });

    if (!existing && status === 'SUBMITTED') {
      const duplicate = existingRows.find(function(row) {
        const sameRestaurant = String(row.restaurant_number) === String(restaurantNumber);
        const samePeriod = String(row.pay_period_ending) === String(payPeriodEnding);
        const openStatus = ['DRAFT', 'REJECTED'].indexOf(String(row.status || '').toUpperCase()) === -1;
        return sameRestaurant && samePeriod && openStatus;
      });
      if (duplicate) throw new Error('Payroll already submitted for this restaurant and pay period.');
    }

    const payloadJson = JSON.stringify(buildPayrollPayloadForStorageV340_(payload));
    const history = existing ? parsePayrollJsonV340_(existing.status_history_json, []) : [];
    history.push({
      status: status,
      note: status === 'DRAFT' ? 'Draft saved' : 'Payroll submitted',
      updated_by: session.fullName || session.username || session.userId,
      updated_at: now
    });

    const record = {
      submission_id: submissionId,
      restaurant_id: session.restaurantId || '',
      restaurant_number: restaurantNumber,
      pay_period_ending: payPeriodEnding,
      submitted_by: session.fullName || session.username || '',
      submitted_by_user_id: session.userId || '',
      submitted_at: existing && existing.submitted_at ? existing.submitted_at : (status === 'SUBMITTED' ? now : ''),
      status: status,
      time_card_updated: cleanPayrollTextV340_(payload.timeCardUpdated),
      stat_hours_1_0: cleanPayrollTextV340_(payload.statHours),
      payroll_approval: cleanPayrollTextV340_(payload.payrollApproval),
      special_notes: cleanPayrollTextV340_(payload.specialNotes),
      admin_notes: existing ? cleanPayrollTextV340_(existing.admin_notes) : '',
      correction_notes: existing ? cleanPayrollTextV340_(existing.correction_notes) : '',
      reviewed_by: existing ? cleanPayrollTextV340_(existing.reviewed_by) : '',
      reviewed_at: existing ? cleanPayrollTextV340_(existing.reviewed_at) : '',
      approved_at: existing ? cleanPayrollTextV340_(existing.approved_at) : '',
      updated_at: now,
      updated_by: session.fullName || session.username || session.userId,
      payload_json: payloadJson,
      status_history_json: JSON.stringify(history)
    };

    if (existing) {
      Repository.updateById_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions', 'submission_id', submissionId, record);
      replacePayrollRowsForSubmissionV340_(submissionId, restaurantNumber, payload);
    } else {
      Repository.append_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions', record);
      replacePayrollRowsForSubmissionV340_(submissionId, restaurantNumber, payload);
    }

    logAuditEvent_(status === 'DRAFT' ? 'PAYROLL_DRAFT_SAVED' : 'PAYROLL_SUBMITTED', 'PAYROLL', submissionId, session.userId, { restaurantNumber: restaurantNumber, payPeriodEnding: payPeriodEnding }, session);
    return respond_(true, { submissionId: submissionId, status: status });
  } catch (error) {
    console.error('upsertPayrollSubmission failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function replacePayrollRowsForSubmissionV340_(submissionId, restaurantNumber, payload) {
  const sheetName = APP_CONFIG.sheets.PAYROLL_ROWS || 'Payroll Rows';
  try {
    Repository.replaceRowsByColumn_(sheetName, 'submission_id', submissionId, []);
  } catch (error) {
    // Some older Repository versions may not allow deleting with empty rows.
    const sheet = Repository.getSheet_(sheetName);
    const values = sheet.getDataRange().getValues();
    if (values.length > 1) {
      const headers = values[0].map(function(h) { return String(h || '').trim(); });
      const idIndex = headers.indexOf('submission_id');
      for (var i = values.length - 1; i >= 1; i--) {
        if (String(values[i][idIndex]) === String(submissionId)) sheet.deleteRow(i + 1);
      }
    }
  }

  const now = new Date().toISOString();
  const rows = [];
  function addRows(sectionKey, list) {
    (Array.isArray(list) ? list : []).forEach(function(item) {
      item = item || {};
      const employeeName = cleanPayrollTextV340_(item.employeeName || item.name || '');
      if (!employeeName && !cleanPayrollTextV340_(item.employeeId)) return;
      rows.push({
        row_id: Utilities.getUuid(),
        submission_id: submissionId,
        restaurant_number: restaurantNumber,
        section_key: sectionKey,
        employee_id: cleanPayrollTextV340_(item.employeeId),
        employee_name: employeeName,
        first_name: cleanPayrollTextV340_(item.firstName),
        last_name: cleanPayrollTextV340_(item.lastName),
        attendance_percentage: cleanPayrollTextV340_(item.attendancePercentage),
        start_date: cleanPayrollTextV340_(item.startDate),
        termination_date: cleanPayrollTextV340_(item.terminationDate),
        leave_of_absence_date: cleanPayrollTextV340_(item.leaveOfAbsenceDate),
        reason_for_separation: cleanPayrollTextV340_(item.reasonForSeparation),
        uniform_deduction: cleanPayrollTextV340_(item.uniformDeduction),
        missing_hours: cleanPayrollTextV340_(item.missingHours),
        training_hours: cleanPayrollTextV340_(item.trainingHours),
        reason: cleanPayrollTextV340_(item.reason),
        item: cleanPayrollTextV340_(item.item),
        amount: cleanPayrollTextV340_(item.amount),
        notes: cleanPayrollTextV340_(item.notes),
        created_at: now
      });
    });
  }

  addRows('ATTENDANCE_BONUS', payload.attendanceBonusRows);
  addRows('NEW_HIRE_RETURNED', payload.newHireRows);
  addRows('ROE_TERMINATION_LOA', payload.roeRows);
  addRows('MISSING_TRAINING_HOURS', payload.missingTrainingRows);
  addRows('UNIFORM_DEDUCTION', payload.uniformDeductionRows);

  rows.forEach(function(row) { Repository.append_(sheetName, row); });
}

function getPayrollSubmissionsForSessionV340_(session, restaurantNumber) {
  const rows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions');
  const isMRORole = [
    APP_CONFIG.roles.MRO, APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.SENIOR_MANAGEMENT
  ].indexOf(session.roleCode) !== -1;

  // MRO/HR: return all submissions (scoping handled at dashboard level)
  if (isMRORole) {
    return rows.map(enrichPayrollSubmissionV340_).sort(sortPayrollSubmissionsV340_);
  }

  // Restaurant user: return all submissions for restaurants this user has access to
  const accessRows = Repository.getAllCached_(APP_CONFIG.sheets.RESTAURANT_ACCESS).filter(function(a) {
    return String(a.user_id || '') === String(session.userId || '');
  });
  const assignedRns = [];
  if (restaurantNumber) assignedRns.push(String(restaurantNumber));
  accessRows.forEach(function(a) {
    const rn = String(a.restaurant_number || a.restaurant_id || '');
    if (rn && assignedRns.indexOf(rn) === -1) assignedRns.push(rn);
  });

  return rows.filter(function(row) {
    // Match by restaurant number OR by the user who submitted it
    const rnMatch = assignedRns.length
      ? assignedRns.indexOf(String(row.restaurant_number || '')) !== -1
      : String(row.restaurant_number || '') === String(restaurantNumber || '');
    const userMatch = String(row.submitted_by_user_id || '') === String(session.userId || '');
    return rnMatch || userMatch;
  }).map(enrichPayrollSubmissionV340_).sort(sortPayrollSubmissionsV340_);
}

function getPayrollSubmissionsForAdminV340_() {
  const rows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions');
  return rows.map(enrichPayrollSubmissionV340_).sort(sortPayrollSubmissionsV340_);
}

function legacyEnrichPayrollSubmissionV340_(row) {
  row = row || {};
  const payload = parsePayrollJsonV340_(row.payload_json, {});
  const history = parsePayrollJsonV340_(row.status_history_json, []);
  return {
    submission_id: cleanPayrollTextV340_(row.submission_id),
    restaurant_number: cleanPayrollTextV340_(row.restaurant_number),
    pay_period_ending: cleanPayrollTextV340_(row.pay_period_ending),
    submitted_by: cleanPayrollTextV340_(row.submitted_by),
    submitted_at: cleanPayrollTextV340_(row.submitted_at),
    status: cleanPayrollTextV340_(row.status || 'DRAFT'),
    time_card_updated: cleanPayrollTextV340_(row.time_card_updated),
    stat_hours_1_0: cleanPayrollTextV340_(row.stat_hours_1_0),
    payroll_approval: cleanPayrollTextV340_(row.payroll_approval),
    special_notes: cleanPayrollTextV340_(row.special_notes),
    admin_notes: cleanPayrollTextV340_(row.admin_notes),
    correction_notes: cleanPayrollTextV340_(row.correction_notes),
    approved_by: cleanPayrollTextV340_(row.approved_by),
    approved_by_name: cleanPayrollTextV340_(row.approved_by_name || row.approved_by || row.reviewed_by),
    reviewed_by: cleanPayrollTextV340_(row.reviewed_by || row.approved_by_name || row.approved_by),
    reviewed_at: cleanPayrollTextV340_(row.reviewed_at),
    approved_at: cleanPayrollTextV340_(row.approved_at),
    updated_at: cleanPayrollTextV340_(row.updated_at),
    updated_by: cleanPayrollTextV340_(row.updated_by),
    payload: payload,
    status_history: history,
    action_needed_by: getPayrollActionNeededByV340_(row.status)
  };
}

function sortPayrollSubmissionsV340_(a, b) {
  return String(b.updated_at || b.submitted_at || '').localeCompare(String(a.updated_at || a.submitted_at || ''));
}

function legacyGetPayrollActionNeededByV340_(status) {
  const value = normalizePayrollStatusV340_(status);
  if (value === 'CORRECTION_NEEDED') return 'Restaurant';
  if (value === 'SUBMITTED' || value === 'VIEWED' || value === 'IN_REVIEW' || value === 'RESUBMITTED') return 'Admin';
  if (value === 'APPROVED') return 'None';
  return 'Restaurant';
}

function buildPayrollPayloadForStorageV340_(payload) {
  return {
    attendanceBonusRows: Array.isArray(payload.attendanceBonusRows) ? payload.attendanceBonusRows : [],
    newHireRows: Array.isArray(payload.newHireRows) ? payload.newHireRows : [],
    roeRows: Array.isArray(payload.roeRows) ? payload.roeRows : [],
    missingTrainingRows: Array.isArray(payload.missingTrainingRows) ? payload.missingTrainingRows : [],
    uniformDeductionRows: Array.isArray(payload.uniformDeductionRows) ? payload.uniformDeductionRows : []
  };
}

function getPayrollEmployeesForRestaurantV340_(restaurantNumber) {
  const teamSheetName = APP_CONFIG.sheets.TEAM_MEMBERS || 'Team Members';
  return Repository.getAll_(teamSheetName).filter(function(row) {
    return payrollEmployeeIsActiveV366_(row) && payrollEmployeeMatchesRestaurantV366_(row, restaurantNumber);
  }).map(payrollMapEmployeeV366_).sort(function(a, b) { return a.full_name.localeCompare(b.full_name); });
}

function getPrimaryRestaurantNumberForSessionV340_(session) {
  if (!session) return '';
  const restaurants = Repository.getAllCached_(APP_CONFIG.sheets.RESTAURANTS);
  const direct = restaurants.find(function(row) {
    return String(row.restaurant_id || '') === String(session.restaurantId || '');
  });
  if (direct) return cleanPayrollTextV340_(direct.restaurant_number || direct.restaurant_code || direct.restaurant_id);
  return cleanPayrollTextV340_(session.restaurantNumber || session.restaurant_number || session.restaurantId || '');
}

function cleanPayrollTextV340_(value) {
  return String(value == null ? '' : value).trim();
}

function parsePayrollJsonV340_(value, fallback) {
  try {
    if (!value) return fallback;
    return JSON.parse(String(value));
  } catch (error) {
    return fallback;
  }
}

// Last definition wins in Apps Script. This keeps the new payroll labels visible.
function formatActionLabel_(key) {
  const labelMap = {
    RECRUITMENT: 'Recruitment',
    RESTAURANT_REQUEST: 'Admin Request',
    REPAIR_MAINTENANCE: 'Repair and Maintenance',
    SNAPSHOT_5W: '5W Tracker',
    NEWSLETTER_MANAGEMENT: 'Newsletter Management',
    NEW_HIRE_MANAGEMENT: 'New Hire Management',
    RECRUITMENT_REQUESTS: 'Restaurant Request',
    MAINTENANCE_REQUESTS: 'Maintenance Requests',
    ASSIGNED_REPORTS: 'Assigned Restaurant Reports',
    REPORTS: 'Reports',
    WEBSITE_CONTROL: 'Website Control',
    ROLE_ACTION_BUTTON_CONTROL: 'Role Action Button Control',
    RESTAURANT_MANAGEMENT: 'Restaurant Management',
    BRAND_MANAGEMENT: 'Brand Management',
    EMPLOYEE_MANAGEMENT: 'Manage Team (add employees)',
    USER_MANAGEMENT: 'User Management',
    HOST_LINK_MANAGEMENT: 'HOST Link Management',
    CONTACT_US_REQUESTS: 'Contact Us Requests',
    PASSWORD_CHANGE_REQUESTS: 'Password Change Requests',
    DROPDOWN_SETTINGS: 'Dropdown Settings',
    OPERATIONS_DASHBOARD: 'Operations Dashboard',
    PAYROLL_SUBMISSION: 'Payroll Submission',
    PAYROLL_REVIEW: 'Payroll Review',
    PAYROLL_ADMIN: 'Payroll Admin',
    YEAR_END_TRANSITION: 'Year End Transition'
  };
  return labelMap[key] || String(key || '').replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, function(char) { return char.toUpperCase(); });
}

// ═══════════════════════════════════════════════════════════════════════
// PAYROLL MODULE V360 — Five targeted improvements
// 1. HR_ADMIN / MRO can submit on behalf of a restaurant
// 2. Employee ID validation against Team Members sheet
// 3. Pay period normalized to YYYY-MM-DD before save/duplicate-check
// 4. Export: summary spreadsheet per pay period
// 5. Recall: restaurant can withdraw DRAFT or SUBMITTED (before MRO views)
// ═══════════════════════════════════════════════════════════════════════

// ── IMPROVEMENT 1 + 2 + 3: Enhanced upsert with HR submit + validation + normalization ──
function legacySubmitPayrollOnBehalfV340_(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    // FIX 1: HR_ADMIN and MRO can now submit on behalf of a restaurant.
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO
    ]);

    const now = new Date().toISOString();
    // FIX 3: Normalize pay period to YYYY-MM-DD regardless of what the client sends.
    const rawPeriod = cleanPayrollTextV340_(payload.payPeriodEnding);
    const payPeriodEnding = normalizePayPeriod_(rawPeriod);
    if (!payPeriodEnding) throw new Error('A valid week ending date is required (YYYY-MM-DD).');

    const restaurantNumber = cleanPayrollTextV340_(
      payload.restaurantNumber ||
      (payload.onBehalfOf && payload.onBehalfOf.restaurantNumber) ||
      getPrimaryRestaurantNumberForSessionV340_(session)
    );
    if (!restaurantNumber) throw new Error('Restaurant number is required.');

    const submissionId = cleanPayrollTextV340_(payload.submissionId) || Utilities.getUuid();
    const existingRows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS);
    const existing = existingRows.find(function(r) { return String(r.submission_id) === submissionId; });

    // Duplicate guard using normalized period
    if (!existing) {
      const duplicate = existingRows.find(function(r) {
        return String(r.restaurant_number) === String(restaurantNumber)
          && normalizePayPeriod_(String(r.pay_period_ending)) === payPeriodEnding
          && ['DRAFT', 'RECALLED'].indexOf(String(r.status || '').toUpperCase()) === -1;
      });
      if (duplicate) throw new Error('Payroll already submitted for restaurant ' + restaurantNumber + ' and period ' + payPeriodEnding + '.');
    }

    // FIX 2: Validate employee IDs against Team Members sheet.
    const teamMembers = Repository.getAllCached_(APP_CONFIG.sheets.TEAM_MEMBERS);
    const validIds = {};
    teamMembers.forEach(function(m) {
      const id = cleanPayrollTextV340_(m.employee_id || m.team_member_id || '');
      if (id) validIds[id] = true;
    });
    const sections = ['attendanceBonusRows','newHireRows','roeRows','missingTrainingRows','uniformDeductionRows'];
    const invalidEmployees = [];
    sections.forEach(function(section) {
      (payload[section] || []).forEach(function(row) {
        const empId = cleanPayrollTextV340_(row.employeeId || '');
        if (empId && !validIds[empId]) {
          invalidEmployees.push(empId + ' (' + cleanPayrollTextV340_(row.employeeName || '') + ')');
        }
      });
    });
    if (invalidEmployees.length) {
      throw new Error('Unknown employee IDs found — not in Team Members sheet: ' + invalidEmployees.join(', ') + '. Please check your roster.');
    }

    const history = existing ? parsePayrollJsonV340_(existing.status_history_json, []) : [];
    const submittedBy = existing ? (existing.submitted_by || session.fullName) : session.fullName;
    const submittedByUserId = existing ? (existing.submitted_by_user_id || session.userId) : session.userId;
    // Note if HR is submitting on behalf
    const onBehalf = (session.roleCode === APP_CONFIG.roles.HR_ADMIN || session.roleCode === APP_CONFIG.roles.MRO)
      ? ' (on behalf of restaurant, submitted by ' + session.fullName + ')'
      : '';
    history.push({ status: 'SUBMITTED', note: 'Payroll submitted' + onBehalf, updated_by: session.fullName, updated_at: now });

    const record = {
      submission_id: submissionId,
      restaurant_id: session.restaurantId || payload.restaurantId || '',
      restaurant_number: restaurantNumber,
      pay_period_ending: payPeriodEnding,
      submitted_by: submittedBy,
      submitted_by_user_id: submittedByUserId,
      submitted_at: existing && existing.submitted_at ? existing.submitted_at : now,
      status: 'SUBMITTED',
      time_card_updated: cleanPayrollTextV340_(payload.timeCardUpdated),
      stat_hours_1_0: cleanPayrollTextV340_(payload.statHours),
      payroll_approval: cleanPayrollTextV340_(payload.payrollApproval),
      special_notes: cleanPayrollTextV340_(payload.specialNotes),
      admin_notes: existing ? cleanPayrollTextV340_(existing.admin_notes) : '',
      correction_notes: existing ? cleanPayrollTextV340_(existing.correction_notes) : '',
      reviewed_by: existing ? cleanPayrollTextV340_(existing.reviewed_by) : '',
      reviewed_at: existing ? cleanPayrollTextV340_(existing.reviewed_at) : '',
      approved_at: existing ? cleanPayrollTextV340_(existing.approved_at) : '',
      updated_at: now,
      updated_by: session.fullName,
      payload_json: JSON.stringify(buildPayrollPayloadForStorageV340_(payload)),
      status_history_json: JSON.stringify(history)
    };

    if (existing) {
      Repository.updateById_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS, 'submission_id', submissionId, record);
    } else {
      Repository.append_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS, record);
    }
    replacePayrollRowsForSubmissionV340_(submissionId, restaurantNumber, payload);
    logAuditEvent_('PAYROLL_SUBMITTED', 'PAYROLL', submissionId, session.userId,
      { restaurantNumber: restaurantNumber, payPeriodEnding: payPeriodEnding, onBehalf: Boolean(onBehalf) }, session);
    invalidatePayrollCaches_(session.userId);
    return respond_(true, { submissionId: submissionId, status: 'SUBMITTED', payPeriodEnding: payPeriodEnding });
  } catch (error) {
    console.error('submitPayrollOnBehalf failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

// ── IMPROVEMENT 5: Recall — restaurant withdraws before MRO views it ──
function recallPayrollSubmission(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO
    ]);

    const submissionId = cleanPayrollTextV340_(payload.submissionId);
    if (!submissionId) throw new Error('Submission ID required.');

    const rows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS);
    const current = rows.find(function(r) { return String(r.submission_id) === submissionId; });
    if (!current) throw new Error('Payroll submission not found.');

    // Only allow recall before MRO has started reviewing
    const recallableStatuses = ['DRAFT', 'SUBMITTED'];
    if (recallableStatuses.indexOf(String(current.status || '').toUpperCase()) === -1) {
      throw new Error('This payroll can no longer be recalled — MRO has already started reviewing it. Contact your MRO directly.');
    }

    // Restaurant users can only recall their own restaurant's submission
    if (session.roleCode === APP_CONFIG.roles.RESTAURANT_USER) {
      if (String(current.submitted_by_user_id) !== String(session.userId) &&
          String(current.restaurant_number) !== String(session.restaurantNumber || session.restaurant_number || '')) {
        throw new Error('You can only recall payroll for your own restaurant.');
      }
    }

    const now = new Date().toISOString();
    const history = parsePayrollJsonV340_(current.status_history_json, []);
    history.push({ status: 'RECALLED', note: cleanPayrollTextV340_(payload.reason || 'Recalled by submitter'), updated_by: session.fullName, updated_at: now });

    Repository.updateById_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS, 'submission_id', submissionId, {
      status: 'RECALLED',
      updated_at: now,
      updated_by: session.fullName,
      status_history_json: JSON.stringify(history)
    });
    logAuditEvent_('PAYROLL_RECALLED', 'PAYROLL', submissionId, session.userId, { reason: payload.reason || '' }, session);
    invalidateMROPayrollDashboardCacheV367_(session.userId);
    return respond_(true, { submissionId: submissionId, status: 'RECALLED' });
  } catch (error) {
    console.error('recallPayrollSubmission failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function payrollCsvEscapeV367_(value) {
  return '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
}

function createPayrollExportFileV367_(fileName, contents) {
  let file = null;
  try {
    const spreadsheetFile = DriveApp.getFileById(APP_CONFIG.spreadsheetId);
    const parents = spreadsheetFile.getParents();
    if (parents.hasNext()) {
      file = parents.next().createFile(fileName, contents, MimeType.CSV);
    }
  } catch (error) {
    console.warn('createPayrollExportFileV367_ parent folder fallback', error);
  }
  return file || DriveApp.createFile(fileName, contents, MimeType.CSV);
}

function buildPayrollSummaryCsvV367_(rows) {
  const headers = [
    'Submission ID', 'Restaurant #', 'Pay Period Ending', 'Submitted By',
    'Submitted At', 'Status', 'Time Card Updated', 'Stat Hours',
    'Reviewed By', 'Reviewed At', 'Approved At',
    'Attendance Bonus Count', 'New Hire Count', 'ROE Count',
    'Missing/Training Hours Count', 'Uniform Deduction Count',
    'Special Notes', 'Admin Notes', 'Correction Notes'
  ];
  const lines = [headers.map(payrollCsvEscapeV367_).join(',')];
  rows.forEach(function(r) {
    const p = r.payload || {};
    lines.push([
      r.submission_id || '',
      r.restaurant_number || '',
      r.pay_period_ending || '',
      r.submitted_by || '',
      r.submitted_at || '',
      r.status || '',
      r.time_card_updated || '',
      r.stat_hours_1_0 || '',
      r.reviewed_by || '',
      r.reviewed_at || '',
      r.approved_at || '',
      (p.attendanceBonusRows || []).length,
      (p.newHireRows || []).length,
      (p.roeRows || []).length,
      (p.missingTrainingRows || []).length,
      (p.uniformDeductionRows || []).length,
      r.special_notes || '',
      r.admin_notes || '',
      r.correction_notes || ''
    ].map(payrollCsvEscapeV367_).join(','));
  });
  return lines.join('\n');
}

// ── IMPROVEMENT 4: Export payroll summary to CSV in Drive ──
function exportPayrollSummary(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const filterStatus = cleanPayrollTextV340_(payload.filterStatus || '');
    const filterPeriod = normalizePayPeriod_(cleanPayrollTextV340_(payload.filterPeriod || ''));

    let rows = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS)
      .map(enrichPayrollSubmissionV361_);

    if (filterStatus) rows = rows.filter(function(r) { return String(r.status || '').toUpperCase() === filterStatus.toUpperCase(); });
    if (filterPeriod) rows = rows.filter(function(r) { return normalizePayPeriod_(String(r.pay_period_ending || '')) === filterPeriod; });

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    const fileName = 'payroll-summary-' + stamp + '.csv';
    const csv = buildPayrollSummaryCsvV367_(rows);
    const exportFile = createPayrollExportFileV367_(fileName, csv);

    logAuditEvent_('PAYROLL_EXPORTED', 'PAYROLL', exportFile.getId(), session.userId,
      { count: rows.length, filterStatus: filterStatus, filterPeriod: filterPeriod, fileName: exportFile.getName() }, session);
    return respond_(true, {
      fileId: exportFile.getId(),
      fileName: exportFile.getName(),
      fileUrl: exportFile.getUrl(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + exportFile.getId(),
      rowCount: rows.length,
      message: 'Exported ' + rows.length + ' submission(s) to CSV file "' + exportFile.getName() + '".'
    });
  } catch (error) {
    console.error('exportPayrollSummary failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

// ── MRO PAYROLL DASHBOARD DATA ──────────────────────────────────────────
// ── Payroll cache invalidation ────────────────────────────────────────────
function invalidatePayrollCaches_(userId) {
  try {
    const cache = CacheService.getScriptCache();
    const s = String(userId || '').slice(-8);
    ['payroll_sub_ws_', 'payroll_rev_ws_', 'payroll_admin_ws_'].forEach(function(p) {
      try { cache.remove(p + s); } catch(e) {}
    });
    // FIX: Also clear the Repository.getAll_ cache for PAYROLL_SUBMISSIONS
    // so the next openReview_() call gets fresh data, not 45-second stale cache
    try {
      if (typeof Repository !== 'undefined' && typeof Repository.clearCacheV258_ === 'function') {
        Repository.clearCacheV258_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions');
      }
    } catch(e) {}
  } catch(e) {}
}

function getMROPayrollDashboardData(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    // ── Date mode cutoff ─────────────────────────────────────────────────
    const dateMode = String(payload.dateMode || '7d');
    const now = new Date();
    function dateCutoff(mode) {
      const d = new Date(now);
      d.setHours(0, 0, 0, 0);
      if (mode === '7d')    { d.setDate(d.getDate() - 7); return d; }
      if (mode === 'month') { return new Date(d.getFullYear(), d.getMonth(), 1); }
      if (mode === 'ytd')   { return new Date(d.getFullYear(), 0, 1); }
      return null; // 'all' = no cutoff
    }
    const cutoff = dateCutoff(dateMode);

    // Cache key includes dateMode so different periods are cached separately
    const cacheKey = 'mro_dash_v365_' + dateMode + '_' + String(session.userId || '').slice(-8);
    const cache = CacheService.getScriptCache();
    if (!payload.forceRefresh) {
      try { const c = cache.get(cacheKey); if (c) return JSON.parse(c); } catch (e) {}
    }

    const allSubmissions = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS)
      .map(enrichPayrollSubmissionV361_);

    // Scope MRO to their assigned restaurants
    const scopeIds = getRecruitmentRestaurantScopeIds_(session);
    const allScoped = scopeIds
      ? allSubmissions.filter(function(r) { return scopeIds.indexOf(String(r.restaurant_id || r.restaurant_number || '')) !== -1; })
      : allSubmissions;

    // Date-filtered submissions (for counts, progress, attention)
    const submissions = cutoff
      ? allScoped.filter(function(r) {
          const s = String(r.status || '').toUpperCase();
          // For APPROVED, use approved_at; for others, use submitted_at
          const t = s === 'APPROVED'
            ? new Date(r.approved_at || r.updated_at || r.submitted_at || 0)
            : new Date(r.submitted_at || r.updated_at || r.created_at || 0);
          return t >= cutoff;
        })
      : allScoped;

    // Status counts — from date-filtered set
    const counts = { DRAFT: 0, SUBMITTED: 0, IN_REVIEW: 0, CORRECTION_NEEDED: 0, APPROVED: 0, RECALLED: 0 };
    submissions.forEach(function(r) {
      const s = String(r.status || 'DRAFT').toUpperCase();
      if (counts[s] !== undefined) counts[s]++;
    });

    // Restaurants scope
    const allRestaurants = Repository.getAllCached_(APP_CONFIG.sheets.RESTAURANTS);
    const scopedRestaurants = scopeIds
      ? allRestaurants.filter(function(r) { return scopeIds.indexOf(String(r.restaurant_id || '')) !== -1; })
      : allRestaurants;

    // Submission progress — latest pay period whose ending date falls within the date window
    // If dateMode=7d, show the latest period that ended in the last 7 days (so approved ones count)
    // Fall back to the absolute latest period if none fall in window
    function periodInWindow(periodStr) {
      if (!cutoff || !periodStr) return true;
      const pd = new Date(periodStr);
      return !isNaN(pd.getTime()) && pd >= cutoff;
    }
    const allPeriods = allScoped
      .map(function(r) { return normalizePayPeriod_(String(r.pay_period_ending || '')); })
      .filter(Boolean).sort().reverse();
    // Try to find the latest period within the date window
    const periodsInWindow = allPeriods.filter(periodInWindow);
    const latestPeriod = periodsInWindow.length ? periodsInWindow[0] : (allPeriods[0] || '');
    const submittedRestaurants = {};
    if (latestPeriod) {
      allScoped.filter(function(r) {
        return normalizePayPeriod_(String(r.pay_period_ending || '')) === latestPeriod
          && String(r.status || '').toUpperCase() !== 'RECALLED';
      }).forEach(function(r) { submittedRestaurants[String(r.restaurant_number || '')] = true; });
    }
    const missingRestaurants = scopedRestaurants.filter(function(r) {
      const rn = String(r.restaurant_number || r.restaurant_code || '');
      return rn && !submittedRestaurants[rn];
    }).map(function(r) {
      return {
        restaurantNumber: String(r.restaurant_number || r.restaurant_code || ''),
        restaurantName: String(r.restaurant_name || r.name || '')
      };
    });

    // Needs attention — SUBMITTED or CORRECTION_NEEDED, within date window
    const TWO_DAYS = 2 * 24 * 60 * 60 * 1000;
    const nowMs = now.getTime();
    const needsAttention = submissions.filter(function(r) {
      const s = String(r.status || '').toUpperCase();
      if (s !== 'SUBMITTED' && s !== 'CORRECTION_NEEDED') return false;
      return (nowMs - new Date(r.submitted_at || r.updated_at || 0).getTime()) > TWO_DAYS;
    }).slice(0, 10);

    // Recent approvals — filtered by dateMode
    const recentApprovals = allScoped.filter(function(r) {
      if (String(r.status || '').toUpperCase() !== 'APPROVED') return false;
      if (!cutoff) return true;
      return new Date(r.approved_at || 0) >= cutoff;
    }).sort(function(a, b) { return new Date(b.approved_at) - new Date(a.approved_at); }).slice(0, 20);

    // Pending final approval — IN_REVIEW (not date-filtered, always show all)
    const pendingFinalApproval = allScoped.filter(function(r) {
      return String(r.status || '').toUpperCase() === 'IN_REVIEW';
    });

    const result = respond_(true, {
      counts: counts,
      dateMode: dateMode,
      latestPeriod: latestPeriod,
      progressCount: Object.keys(submittedRestaurants).length,  // restaurants that submitted for latestPeriod
      missingRestaurants: missingRestaurants,
      needsAttention: needsAttention,
      recentApprovals: recentApprovals,
      pendingFinalApproval: pendingFinalApproval,
      totalRestaurants: scopedRestaurants.length,
      generatedAt: new Date().toISOString()
    });

    try { const j = JSON.stringify(result); if (j.length < 90000) cache.put(cacheKey, j, 90); } catch (e) {}
    return result;
  } catch (error) {
    console.error('getMROPayrollDashboardData failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

// ── Helper: normalize any date string to YYYY-MM-DD ──────────────────────
function normalizePayPeriod_(value) {
  if (!value) return '';
  var s = String(value).trim();
  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // Try JS Date parse for other formats (MM/DD/YYYY, "March 15 2025", etc.)
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
  }
  return s; // return as-is if unparseable — will fail downstream validation
}


// ═══════════════════════════════════════════════════════════════════════
// PAYROLL CORRECTION TRACKING V361
// Last definitions win in Apps Script. This block safely overrides only
// payroll submission/status functions and keeps the same public function names.
// ═══════════════════════════════════════════════════════════════════════

function normalizePayrollStatusV340_(status) {
  const value = String(status || '').trim().toUpperCase().replace(/\s+/g, '_');
  const allowed = [
    'DRAFT',
    'SUBMITTED',
    'VIEWED',
    'IN_REVIEW',
    'CORRECTION_NEEDED',
    'RESUBMITTED',
    'APPROVED',
    'REJECTED',
    'RECALLED'
  ];
  return allowed.indexOf(value) !== -1 ? value : 'SUBMITTED';
}

function payrollActorNameV361_(session) {
  return cleanPayrollTextV340_((session && (session.fullName || session.username || session.userId)) || 'System');
}

function payrollHistoryListV361_(row) {
  return parsePayrollJsonV340_((row && (row.payroll_history_json || row.status_history_json)) || '[]', []);
}

function payrollPushHistoryV361_(history, action, note, session, extra) {
  history = Array.isArray(history) ? history : [];
  extra = extra || {};
  history.push({
    action: cleanPayrollTextV340_(action),
    status: cleanPayrollTextV340_(extra.status || action),
    note: cleanPayrollTextV340_(note || ''),
    updated_by: payrollActorNameV361_(session),
    updated_role: cleanPayrollTextV340_((session && session.roleCode) || ''),
    updated_user_id: cleanPayrollTextV340_((session && session.userId) || ''),
    updated_at: cleanPayrollTextV340_(extra.timestamp || new Date().toISOString())
  });
  return history;
}

function payrollIsAdminSessionV361_(session) {
  return [
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ].indexOf(session.roleCode) !== -1;
}

function payrollCanRestaurantEditStatusV361_(status, lockedByMro) {
  const s = normalizePayrollStatusV340_(status || 'DRAFT');
  const locked = String(lockedByMro || '').toUpperCase() === 'YES';
  if (locked) return false;
  return ['DRAFT', 'SUBMITTED', 'CORRECTION_NEEDED', 'RECALLED'].indexOf(s) !== -1;
}

function legacySavePayrollSubmissionDraftV361_(payload) {
  return upsertPayrollSubmissionV361_(payload, 'DRAFT');
}

function legacySubmitPayrollSubmissionV361_(payload) {
  return upsertPayrollSubmissionV361_(payload, 'SUBMITTED');
}

function legacySubmitPayrollOnBehalfV361_(payload) {
  return upsertPayrollSubmissionV361_(payload, 'SUBMITTED');
}

function upsertPayrollSubmissionV361_(payload, requestedStatus) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const now = new Date().toISOString();
    const isAdmin = payrollIsAdminSessionV361_(session);
    const rawPeriod = cleanPayrollTextV340_(payload.payPeriodEnding);
    const payPeriodEnding = normalizePayPeriod_(rawPeriod);
    const restaurantNumber = cleanPayrollTextV340_(
      payload.restaurantNumber ||
      (payload.onBehalfOf && payload.onBehalfOf.restaurantNumber) ||
      getPrimaryRestaurantNumberForSessionV340_(session)
    );

    if (!restaurantNumber) throw new Error('Restaurant number is required.');
    if (requestedStatus !== 'DRAFT' && !payPeriodEnding) throw new Error('A valid week ending date is required.');

    const submissionId = cleanPayrollTextV340_(payload.submissionId) || Utilities.getUuid();
    const sheetName = APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions';
    const existingRows = Repository.getAll_(sheetName);
    const existing = existingRows.find(function(r) { return String(r.submission_id) === String(submissionId); });
    const currentStatus = existing ? normalizePayrollStatusV340_(existing.status || 'DRAFT') : 'DRAFT';

    if (existing && !isAdmin && !payrollCanRestaurantEditStatusV361_(currentStatus, existing.locked_by_mro)) {
      throw new Error('This payroll is locked because MRO has already opened it.');
    }

    if (!existing && requestedStatus !== 'DRAFT') {
      const duplicate = existingRows.find(function(r) {
        const sameRestaurant = String(r.restaurant_number || '') === String(restaurantNumber);
        const samePeriod = normalizePayPeriod_(String(r.pay_period_ending || '')) === payPeriodEnding;
        const activeStatus = ['RECALLED', 'REJECTED'].indexOf(String(r.status || '').toUpperCase()) === -1;
        return sameRestaurant && samePeriod && activeStatus;
      });
      if (duplicate) throw new Error('Payroll already exists for restaurant ' + restaurantNumber + ' and week ending ' + payPeriodEnding + '.');
    }

    validatePayrollEmployeesV361_(payload);

    let finalStatus = 'DRAFT';
    let historyAction = 'DRAFT_SAVED';
    let historyNote = 'Draft saved';

    if (requestedStatus !== 'DRAFT') {
      if (existing && currentStatus === 'CORRECTION_NEEDED') {
        finalStatus = 'RESUBMITTED';
        historyAction = 'RESUBMITTED';
        historyNote = 'Payroll corrected and resubmitted';
      } else {
        finalStatus = 'SUBMITTED';
        historyAction = 'SUBMITTED';
        historyNote = existing ? 'Payroll updated and submitted' : 'Payroll submitted';
      }
    }

    let history = payrollHistoryListV361_(existing || {});
    payrollPushHistoryV361_(history, historyAction, historyNote, session, { status: finalStatus, timestamp: now });

    const submittedBy = existing ? (existing.submitted_by || payrollActorNameV361_(session)) : (payrollActorNameV361_(session) || cleanPayrollTextV340_(payload.submittedBy || payload.submitted_by || ''));
    const submittedByUserId = existing ? (existing.submitted_by_user_id || session.userId || '') : (session.userId || '');

    const record = {
      submission_id: submissionId,
      restaurant_id: (existing && existing.restaurant_id) || session.restaurantId || payload.restaurantId || '',
      restaurant_number: restaurantNumber,
      pay_period_ending: payPeriodEnding,
      submitted_by: submittedBy,
      submitted_by_user_id: submittedByUserId,
      submitted_at: existing && existing.submitted_at ? existing.submitted_at : (finalStatus === 'DRAFT' ? '' : now),
      status: finalStatus,
      time_card_updated: cleanPayrollTextV340_(payload.timeCardUpdated),
      stat_hours_1_0: cleanPayrollTextV340_(payload.statHours),
      payroll_approval: cleanPayrollTextV340_(payload.payrollApproval),
      special_notes: cleanPayrollTextV340_(payload.specialNotes),
      admin_notes: finalStatus === 'RESUBMITTED' ? '' : (existing ? cleanPayrollTextV340_(existing.admin_notes) : ''),
      correction_notes: finalStatus === 'RESUBMITTED' ? '' : (existing ? cleanPayrollTextV340_(existing.correction_notes) : ''),
      reviewed_by: existing ? cleanPayrollTextV340_(existing.reviewed_by) : '',
      reviewed_at: existing ? cleanPayrollTextV340_(existing.reviewed_at) : '',
      approved_at: existing ? cleanPayrollTextV340_(existing.approved_at) : '',
      locked_by_mro: finalStatus === 'DRAFT' || finalStatus === 'SUBMITTED' || finalStatus === 'RESUBMITTED' ? 'NO' : (existing ? cleanPayrollTextV340_(existing.locked_by_mro) : 'NO'),
      locked_at: finalStatus === 'DRAFT' || finalStatus === 'SUBMITTED' || finalStatus === 'RESUBMITTED' ? '' : (existing ? cleanPayrollTextV340_(existing.locked_at) : ''),
      correction_count: existing ? cleanPayrollTextV340_(existing.correction_count || '0') : '0',
      last_resubmitted_at: finalStatus === 'RESUBMITTED' ? now : (existing ? cleanPayrollTextV340_(existing.last_resubmitted_at) : ''),
      last_reviewed_at: existing ? cleanPayrollTextV340_(existing.last_reviewed_at) : '',
      last_approved_at: existing ? cleanPayrollTextV340_(existing.last_approved_at) : '',
      correction_requested_by: finalStatus === 'RESUBMITTED' ? '' : (existing ? cleanPayrollTextV340_(existing.correction_requested_by) : ''),
      correction_requested_at: finalStatus === 'RESUBMITTED' ? '' : (existing ? cleanPayrollTextV340_(existing.correction_requested_at) : ''),
      updated_at: now,
      updated_by: payrollActorNameV361_(session),
      payload_json: JSON.stringify(buildPayrollPayloadForStorageV340_(payload)),
      status_history_json: JSON.stringify(history),
      payroll_history_json: JSON.stringify(history)
    };

    if (existing) {
      Repository.updateById_(sheetName, 'submission_id', submissionId, record);
    } else {
      Repository.append_(sheetName, record);
    }
    replacePayrollRowsForSubmissionV340_(submissionId, restaurantNumber, payload);

    logAuditEvent_(historyAction, 'PAYROLL', submissionId, session.userId, {
      restaurantNumber: restaurantNumber,
      payPeriodEnding: payPeriodEnding,
      status: finalStatus
    }, session);
    invalidateMROPayrollDashboardCacheV367_(session.userId);

    return respond_(true, { submissionId: submissionId, status: finalStatus, payPeriodEnding: payPeriodEnding });
  } catch (error) {
    console.error('upsertPayrollSubmissionV361_ failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function validatePayrollEmployeesV361_(payload) {
  const teamMembers = Repository.getAllCached_(APP_CONFIG.sheets.TEAM_MEMBERS || 'Team Members');
  const validIds = {};
  teamMembers.forEach(function(m) {
    const id = cleanPayrollTextV340_(m.employee_id || m.team_member_id || '');
    if (id) validIds[id] = true;
  });
  const sections = ['attendanceBonusRows', 'newHireRows', 'roeRows', 'missingTrainingRows', 'uniformDeductionRows'];
  const invalidEmployees = [];
  sections.forEach(function(section) {
    (payload[section] || []).forEach(function(row) {
      const empId = cleanPayrollTextV340_(row.employeeId || '');
      if (empId && !validIds[empId]) invalidEmployees.push(empId + ' (' + cleanPayrollTextV340_(row.employeeName || '') + ')');
    });
  });
  if (invalidEmployees.length) {
    throw new Error('Unknown employee IDs found in payroll: ' + invalidEmployees.join(', '));
  }
}

function buildPayrollStatusPatchV367_(current, requestedStatus, note, session, now) {
  let finalStatus = requestedStatus;
  let action = requestedStatus;
  let actionNote = note;
  const patch = {
    status: finalStatus,
    updated_at: now,
    updated_by: payrollActorNameV361_(session),
    reviewed_by: current.reviewed_by || payrollActorNameV361_(session),
    reviewed_at: current.reviewed_at || now,
    last_reviewed_at: now
  };

  if (requestedStatus === 'VIEWED' || requestedStatus === 'IN_REVIEW') {
    finalStatus = 'IN_REVIEW';
    action = 'VIEWED_BY_MRO';
    actionNote = note || 'Payroll opened by MRO';
    patch.status = finalStatus;
    patch.locked_by_mro = 'YES';
    patch.locked_at = current.locked_at || now;
    patch.last_reviewed_at = now;
  } else if (requestedStatus === 'CORRECTION_NEEDED') {
    action = 'CORRECTION_REQUESTED';
    patch.status = 'CORRECTION_NEEDED';
    patch.admin_notes = note;
    patch.correction_notes = note;
    patch.locked_by_mro = 'NO';
    patch.locked_at = '';
    patch.correction_count = String((parseInt(current.correction_count || '0', 10) || 0) + 1);
    patch.correction_requested_by = payrollActorNameV361_(session);
    patch.correction_requested_at = now;
  } else if (requestedStatus === 'APPROVED') {
    action = 'APPROVED';
    actionNote = note || 'Payroll approved';
    patch.status = 'APPROVED';
    patch.admin_notes = note || 'Payroll approved';
    patch.correction_notes = '';
    patch.correction_requested_by = '';
    patch.correction_requested_at = '';
    patch.locked_by_mro = 'YES';
    patch.locked_at = current.locked_at || now;
    patch.approved_at = now;
    patch.last_approved_at = now;
    patch.approved_by = payrollActorNameV361_(session);
    patch.approved_by_name = payrollActorNameV361_(session);
    patch.reviewed_by = payrollActorNameV361_(session);
    patch.reviewed_at = current.reviewed_at || now;
  } else {
    patch.status = finalStatus;
    if (note) patch.admin_notes = note;
  }

  let history = payrollHistoryListV361_(current);
  payrollPushHistoryV361_(history, action, actionNote, session, { status: patch.status, timestamp: now });
  patch.status_history_json = JSON.stringify(history);
  patch.payroll_history_json = JSON.stringify(history);

  return {
    patch: patch,
    action: action
  };
}

function invalidateMROPayrollDashboardCacheV367_(userId) {
  const suffix = String(userId || '').slice(-8);
  if (!suffix) return;
  const cache = CacheService.getScriptCache();
  ['7d', 'month', 'ytd', 'all'].forEach(function(dateMode) {
    try {
      cache.remove('mro_dash_v365_' + dateMode + '_' + suffix);
    } catch (error) {}
  });
}

function legacyUpdatePayrollSubmissionStatusV361_(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    // Support both single submissionId and batch submissionIds array
    const submissionIds = Array.isArray(payload.submissionIds)
      ? payload.submissionIds.map(cleanPayrollTextV340_).filter(Boolean)
      : [cleanPayrollTextV340_(payload.submissionId)].filter(Boolean);
    if (!submissionIds.length) throw new Error('Submission ID is required.');
    const submissionId = submissionIds[0]; // kept for single-call compatibility below

    const requestedStatus = normalizePayrollStatusV340_(payload.status || 'IN_REVIEW');
    const now = new Date().toISOString();
    const sheetName = APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions';
    const note = cleanPayrollTextV340_(payload.adminNotes || payload.correctionNotes || payload.note || '');
    if (requestedStatus === 'CORRECTION_NEEDED' && !note) {
      throw new Error('Correction note is required.');
    }

    const updateResult = withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const values = sheet.getDataRange().getValues();
      if (values.length < 2) throw new Error('Payroll submission not found.');

      const headers = values[0].map(function(h) { return String(h || '').trim(); });
      const idIndex = headers.indexOf('submission_id');
      if (idIndex === -1) throw new Error('submission_id column not found.');

      const headerIndexMap = {};
      headers.forEach(function(header, index) {
        headerIndexMap[header] = index;
      });

      const rowIndexById = {};
      for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
        const rowId = String(values[rowIndex][idIndex] || '');
        if (rowId && rowIndexById[rowId] == null) {
          rowIndexById[rowId] = rowIndex;
        }
      }

      const updated = [];
      const updatedRowIndexes = [];
      const missing = [];
      const audits = [];
      let singleStatus = '';

      submissionIds.forEach(function(id) {
        const rowIndex = rowIndexById[id];
        if (rowIndex == null) {
          missing.push(id);
          return;
        }

        const current = rowToObject_(headers, values[rowIndex]);
        const result = buildPayrollStatusPatchV367_(current, requestedStatus, note, session, now);

        Object.keys(result.patch).forEach(function(key) {
          if (!Object.prototype.hasOwnProperty.call(headerIndexMap, key)) return;
          values[rowIndex][headerIndexMap[key]] = result.patch[key];
        });

        updated.push(id);
        updatedRowIndexes.push(rowIndex);
        if (!singleStatus) singleStatus = result.patch.status;
        audits.push({
          submissionId: id,
          action: result.action,
          status: result.patch.status
        });
      });

      if (!updated.length) {
        if (submissionIds.length === 1 && missing.length === 1) {
          throw new Error('Payroll submission not found.');
        }
        throw new Error('No payroll submissions were updated.');
      }

      updatedRowIndexes.sort(function(a, b) { return a - b; }).forEach(function(rowIndex) {
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values[rowIndex]]);
      });

      SpreadsheetApp.flush();
      return {
        updated: updated,
        missing: missing,
        audits: audits,
        singleStatus: singleStatus
      };
    });

    updateResult.audits.forEach(function(audit) {
      logAuditEvent_('PAYROLL_' + audit.action, 'PAYROLL', audit.submissionId, session.userId, { status: audit.status }, session);
    });

    // Clear every payroll cache after status change.
    // This fixes approved payrolls returning to Inbox after Refresh.
    invalidatePayrollCaches_(session.userId);
    invalidateMROPayrollDashboardCacheV367_(session.userId);

    return respond_(true, {
      submissionId: updateResult.updated[0] || submissionId,
      submissionIds: updateResult.updated,
      status: updateResult.updated.length === 1 ? updateResult.singleStatus : '',
      count: updateResult.updated.length,
      missing: updateResult.missing
    });
  } catch (error) {
    console.error('updatePayrollSubmissionStatus V361 failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function enrichPayrollSubmissionV361_(row) {
  const payload = parsePayrollJsonV340_(row.payload_json, {});
  const history = payrollHistoryListV361_(row);
  return {
    submission_id: cleanPayrollTextV340_(row.submission_id),
    restaurant_number: cleanPayrollTextV340_(row.restaurant_number),
    pay_period_ending: cleanPayrollTextV340_(row.pay_period_ending),
    submitted_by: cleanPayrollTextV340_(row.submitted_by),
    submitted_at: cleanPayrollTextV340_(row.submitted_at),
    status: cleanPayrollTextV340_(row.status || 'DRAFT'),
    time_card_updated: cleanPayrollTextV340_(row.time_card_updated),
    stat_hours_1_0: cleanPayrollTextV340_(row.stat_hours_1_0),
    payroll_approval: cleanPayrollTextV340_(row.payroll_approval),
    special_notes: cleanPayrollTextV340_(row.special_notes),
    admin_notes: cleanPayrollTextV340_(row.admin_notes),
    correction_notes: cleanPayrollTextV340_(row.correction_notes),
    correction_count: cleanPayrollTextV340_(row.correction_count || '0'),
    correction_requested_by: cleanPayrollTextV340_(row.correction_requested_by),
    correction_requested_at: cleanPayrollTextV340_(row.correction_requested_at),
    locked_by_mro: cleanPayrollTextV340_(row.locked_by_mro || 'NO'),
    locked_at: cleanPayrollTextV340_(row.locked_at),
    approved_by: cleanPayrollTextV340_(row.approved_by),
    approved_by_name: cleanPayrollTextV340_(row.approved_by_name || row.approved_by || row.reviewed_by),
    reviewed_by: cleanPayrollTextV340_(row.reviewed_by || row.approved_by_name || row.approved_by),
    reviewed_at: cleanPayrollTextV340_(row.reviewed_at),
    approved_at: cleanPayrollTextV340_(row.approved_at || row.last_approved_at),
    last_reviewed_at: cleanPayrollTextV340_(row.last_reviewed_at),
    last_resubmitted_at: cleanPayrollTextV340_(row.last_resubmitted_at),
    last_approved_at: cleanPayrollTextV340_(row.last_approved_at),
    updated_at: cleanPayrollTextV340_(row.updated_at),
    updated_by: cleanPayrollTextV340_(row.updated_by),
    admin_viewed_status: cleanPayrollTextV340_(row.admin_viewed_status || (row.admin_viewed_at ? 'VIEWED' : 'NOT_VIEWED')),
    admin_viewed_by: cleanPayrollTextV340_(row.admin_viewed_by),
    admin_viewed_at: cleanPayrollTextV340_(row.admin_viewed_at),
    website_update_status: cleanPayrollTextV340_(row.website_update_status || (row.website_updated_at ? 'WEBSITE_UPDATED' : 'NOT_UPDATED')),
    website_updated_by: cleanPayrollTextV340_(row.website_updated_by),
    website_updated_at: cleanPayrollTextV340_(row.website_updated_at),
    website_update_batch_id: cleanPayrollTextV340_(row.website_update_batch_id),
    brand: cleanPayrollTextV340_(row.brand),
    payroll_batch_key: cleanPayrollTextV340_(row.payroll_batch_key),
    payload_json: cleanPayrollTextV340_(row.payload_json),
    status_history_json: JSON.stringify(history),
    payroll_history_json: JSON.stringify(history),
    payload: payload,
    action_needed_by: getPayrollActionNeededByV361_(row.status)
  };
}

// Compatibility shims for older payroll helpers that are still referenced by
// shared loaders. These intentionally forward to the active V361 behavior.
function enrichPayrollSubmissionV340_(row) {
  return enrichPayrollSubmissionV361_(row);
}

function getPayrollActionNeededByV361_(status) {
  const value = normalizePayrollStatusV340_(status);
  if (value === 'DRAFT') return 'Restaurant';
  if (value === 'SUBMITTED' || value === 'RESUBMITTED' || value === 'IN_REVIEW' || value === 'VIEWED') return 'MRO';
  if (value === 'CORRECTION_NEEDED') return 'Restaurant';
  if (value === 'APPROVED') return 'None';
  return 'MRO';
}

function getPayrollActionNeededByV340_(status) {
  return getPayrollActionNeededByV361_(status);
}


// ═══════════════════════════════════════════════════════════════════════
// PAYROLL ROE APPLICATIONS + PAYROLL ADMIN V365
// Adds ROE pool, Payroll Admin monitor, and employee deactivation.
// Keeps payroll submissions as the same flow.
// ═══════════════════════════════════════════════════════════════════════

function payrollRoeSheetNameV365_() {
  return (APP_CONFIG.sheets && APP_CONFIG.sheets.PAYROLL_ROE_APPLICATIONS) || 'Payroll ROE Applications';
}

function payrollRoeNowV365_() {
  return new Date().toISOString();
}

function payrollRoeCleanV365_(value) {
  return cleanPayrollTextV340_(value);
}

function payrollRoeHistoryV365_(row) {
  return parsePayrollJsonV340_((row && row.history_json) || '[]', []);
}

function payrollRoePushHistoryV365_(history, action, note, session, extra) {
  history = Array.isArray(history) ? history : [];
  extra = extra || {};
  history.push({
    action: payrollRoeCleanV365_(action),
    note: payrollRoeCleanV365_(note || ''),
    updated_by: payrollActorNameV361_(session),
    updated_role: payrollRoeCleanV365_(session && session.roleCode),
    updated_user_id: payrollRoeCleanV365_(session && session.userId),
    updated_at: payrollRoeCleanV365_(extra.timestamp || payrollRoeNowV365_())
  });
  return history;
}

function normalizePayrollRoeStatusV365_(status) {
  const value = String(status || '').trim().toUpperCase().replace(/\s+/g, '_');
  const allowed = [
    'ROE_SUBMITTED',
    'ATTACHED_TO_PAYROLL',
    'REVIEWED_BY_MRO',
    'ROE_APPROVED',
    'COMPLETED',
    'CANCELLED',
    'CORRECTION_NEEDED'
  ];
  return allowed.indexOf(value) !== -1 ? value : 'ROE_SUBMITTED';
}

function generatePayrollRoeIdV365_(restaurantNumber) {
  const rn = payrollRoeCleanV365_(restaurantNumber || '000');
  const ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const rows = Repository.getAll_(payrollRoeSheetNameV365_());
  const prefix = 'ROE-' + rn + '-' + ymd + '-';
  const count = rows.filter(function(r) { return String(r.roe_id || '').indexOf(prefix) === 0; }).length + 1;
  return prefix + String(count).padStart(3, '0');
}

function getPayrollRoeApplicationData(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken || payload);
    const rn = payrollRoeCleanV365_(payload.restaurantNumber || getPrimaryRestaurantNumberForSessionV340_(session));
    const all = Repository.getAll_(payrollRoeSheetNameV365_()).map(enrichPayrollRoeApplicationV365_);
    const isAdmin = payrollIsAdminSessionV361_(session);
    const roes = all.filter(function(row) {
      if (isAdmin) return true;
      return String(row.restaurant_number || '') === String(rn || '');
    });
    const pool = roes.filter(function(row) {
      const status = normalizePayrollRoeStatusV365_(row.status);
      return !row.payroll_submission_id && ['ROE_SUBMITTED', 'ROE_APPROVED', 'CORRECTION_NEEDED'].indexOf(status) !== -1;
    });
    return respond_(true, { restaurantNumber: rn, roes: roes, pool: pool });
  } catch (error) {
    console.error('getPayrollRoeApplicationData failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function submitPayrollRoeApplication(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const now = payrollRoeNowV365_();
    const restaurantNumber = payrollRoeCleanV365_(payload.restaurantNumber || getPrimaryRestaurantNumberForSessionV340_(session));
    if (!restaurantNumber) throw new Error('Restaurant number is required.');

    const employeeId = payrollRoeCleanV365_(payload.employeeId);
    let employeeName = payrollRoeCleanV365_(payload.employeeName);
    let firstName = payrollRoeCleanV365_(payload.firstName);
    let lastName = payrollRoeCleanV365_(payload.lastName);

    if (employeeId && (!employeeName || !firstName)) {
      const team = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS);
      const emp = team.find(function(m) {
        return String(m.employee_id || m.team_member_id || '') === String(employeeId);
      });
      if (emp) {
        employeeName = employeeName || payrollRoeCleanV365_(emp.full_name || emp.employee_name || emp.name);
        const parts = employeeName.split(' ');
        firstName = firstName || payrollRoeCleanV365_(parts[0] || '');
        lastName = lastName || payrollRoeCleanV365_(parts.slice(1).join(' '));
      }
    }

    if (!employeeName && (!firstName || !lastName)) throw new Error('Employee name is required.');
    if (!employeeName) employeeName = (firstName + ' ' + lastName).trim();
    if (!payload.lastDayOfWork) throw new Error('Last day of work is required.');
    if (!payload.reasonForSeparation) throw new Error('Reason for separation is required.');
    if (!payload.legalStatus) throw new Error('Legal status is required.');


    // ── Duplicate check: one active ROE per employee per restaurant ──────
    const closedStatuses = ['COMPLETED','DEACTIVATED','REJECTED','CANCELLED'];
    const allExistingRoes = Repository.getAll_(payrollRoeSheetNameV365_());
    const activeRoeForEmployee = allExistingRoes.find(function(r) {
      const sameRestaurant = payrollRoeCleanV365_(r.restaurant_number) === restaurantNumber;
      const sameEmployee = (employeeId && payrollRoeCleanV365_(r.employee_id || r.employeeId || '') === employeeId)
                        || (employeeName && payrollRoeCleanV365_(r.employee_name || r.employeeName || '').toLowerCase() === employeeName.toLowerCase());
      const isActive = closedStatuses.indexOf(String(r.status || '').toUpperCase()) === -1;
      return sameRestaurant && sameEmployee && isActive;
    });
    if (activeRoeForEmployee) {
      throw new Error('An active ROE already exists for ' + employeeName + ' (ROE ID: ' + (activeRoeForEmployee.roe_id || '') + '). Please resolve the existing ROE before submitting a new one.');
    }

    const roeId = payrollRoeCleanV365_(payload.roeId) || generatePayrollRoeIdV365_(restaurantNumber);
    const existingRows = Repository.getAll_(payrollRoeSheetNameV365_());
    const existing = existingRows.find(function(r) { return String(r.roe_id || '') === String(roeId); });

    let history = existing ? payrollRoeHistoryV365_(existing) : [];
    payrollRoePushHistoryV365_(history, existing ? 'ROE_UPDATED' : 'ROE_SUBMITTED', 'ROE application submitted by restaurant', session, { timestamp: now });

    const record = {
      roe_id: roeId,
      submission_id: roeId,
      restaurant_id: payrollRoeCleanV365_(payload.restaurantId || session.restaurantId || ''),
      restaurant_number: restaurantNumber,
      employee_id: employeeId,
      employee_name: employeeName,
      first_name: firstName,
      last_name: lastName,
      last_day_of_work: payrollRoeCleanV365_(payload.lastDayOfWork),
      reason_for_separation: payrollRoeCleanV365_(payload.reasonForSeparation),
      legal_status: payrollRoeCleanV365_(payload.legalStatus),
      health_benefits: payrollRoeCleanV365_(payload.healthBenefits || 'No'),
      station_champion: payrollRoeCleanV365_(payload.stationChampion || 'No'),
      probation_termination_reason: payrollRoeCleanV365_(payload.probationTerminationReason || ''),
      status: existing ? normalizePayrollRoeStatusV365_(existing.status) : 'ROE_SUBMITTED',
      payroll_submission_id: existing ? payrollRoeCleanV365_(existing.payroll_submission_id) : '',
      pay_period_ending: existing ? payrollRoeCleanV365_(existing.pay_period_ending) : '',
      submitted_by: existing ? payrollRoeCleanV365_(existing.submitted_by) : payrollActorNameV361_(session),
      submitted_by_user_id: existing ? payrollRoeCleanV365_(existing.submitted_by_user_id) : payrollRoeCleanV365_(session.userId),
      submitted_at: existing ? payrollRoeCleanV365_(existing.submitted_at) : now,
      reviewed_by: existing ? payrollRoeCleanV365_(existing.reviewed_by) : '',
      reviewed_at: existing ? payrollRoeCleanV365_(existing.reviewed_at) : '',
      approved_by: existing ? payrollRoeCleanV365_(existing.approved_by) : '',
      approved_at: existing ? payrollRoeCleanV365_(existing.approved_at) : '',
      completed_at: existing ? payrollRoeCleanV365_(existing.completed_at) : '',
      deactivated_employee: existing ? payrollRoeCleanV365_(existing.deactivated_employee) : 'NO',
      deactivated_at: existing ? payrollRoeCleanV365_(existing.deactivated_at) : '',
      admin_notes: existing ? payrollRoeCleanV365_(existing.admin_notes) : '',
      history_json: JSON.stringify(history),
      updated_at: now,
      updated_by: payrollActorNameV361_(session)
    };

    if (existing) Repository.updateById_(payrollRoeSheetNameV365_(), 'roe_id', roeId, record);
    else Repository.append_(payrollRoeSheetNameV365_(), record);

    logAuditEvent_('PAYROLL_ROE_SUBMITTED', 'PAYROLL_ROE', roeId, session.userId, { restaurantNumber: restaurantNumber, employeeName: employeeName }, session);
    return respond_(true, { roeId: roeId, status: record.status });
  } catch (error) {
    console.error('submitPayrollRoeApplication failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function enrichPayrollRoeApplicationV365_(row) {
  return {
    roe_id: payrollRoeCleanV365_(row.roe_id || row.submission_id),
    submission_id: payrollRoeCleanV365_(row.submission_id || row.roe_id),
    restaurant_id: payrollRoeCleanV365_(row.restaurant_id),
    restaurant_number: payrollRoeCleanV365_(row.restaurant_number),
    employee_id: payrollRoeCleanV365_(row.employee_id),
    employee_name: payrollRoeCleanV365_(row.employee_name || ((row.first_name || '') + ' ' + (row.last_name || '')).trim()),
    first_name: payrollRoeCleanV365_(row.first_name),
    last_name: payrollRoeCleanV365_(row.last_name),
    last_day_of_work: payrollRoeCleanV365_(row.last_day_of_work),
    reason_for_separation: payrollRoeCleanV365_(row.reason_for_separation),
    legal_status: payrollRoeCleanV365_(row.legal_status),
    health_benefits: payrollRoeCleanV365_(row.health_benefits),
    station_champion: payrollRoeCleanV365_(row.station_champion),
    probation_termination_reason: payrollRoeCleanV365_(row.probation_termination_reason),
    status: payrollRoeCleanV365_(row.status || 'ROE_SUBMITTED'),
    payroll_submission_id: payrollRoeCleanV365_(row.payroll_submission_id),
    pay_period_ending: payrollRoeCleanV365_(row.pay_period_ending),
    submitted_by: payrollRoeCleanV365_(row.submitted_by),
    submitted_at: payrollRoeCleanV365_(row.submitted_at),
    reviewed_by: payrollRoeCleanV365_(row.reviewed_by),
    reviewed_at: payrollRoeCleanV365_(row.reviewed_at),
    approved_by: payrollRoeCleanV365_(row.approved_by),
    approved_at: payrollRoeCleanV365_(row.approved_at),
    completed_at: payrollRoeCleanV365_(row.completed_at),
    deactivated_employee: payrollRoeCleanV365_(row.deactivated_employee || 'NO'),
    deactivated_at: payrollRoeCleanV365_(row.deactivated_at),
    admin_notes: payrollRoeCleanV365_(row.admin_notes),
    history_json: payrollRoeCleanV365_(row.history_json || '[]'),
    updated_at: payrollRoeCleanV365_(row.updated_at),
    updated_by: payrollRoeCleanV365_(row.updated_by)
  };
}

function getPayrollAdminWorkspaceData(payload) {
  try {
    payload = payload || {};
    ensurePayrollAdminStatusColumns_();
    const session = getSessionOrThrow_(payload.sessionToken || payload);
    const _aKey = 'payroll_admin_ws_' + String(session.userId||'').slice(-8);
    const _aCache = CacheService.getScriptCache();
    if (!payload.forceRefresh) { try { const _c = _aCache.get(_aKey); if (_c) return JSON.parse(_c); } catch(e) {} }
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT]);

    const restaurantRows = Repository.getAll_(APP_CONFIG.sheets.RESTAURANTS)
      .filter(function(r) { return String(r.status || '').toUpperCase() !== 'INACTIVE'; });

    const restaurantBrandMap = {};
    const restaurants = restaurantRows.map(function(r) {
      const number = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber || r.restaurant_code || r.restaurantCode || r.number || r.code);
      const brand = payrollRoeCleanV365_(r.brand || r.brand_name || r.restaurant_brand || r.Brand || r.banner || r.concept || 'Unassigned Brand') || 'Unassigned Brand';
      if (number) restaurantBrandMap[number] = brand;
      return {
        restaurant_number: number,
        restaurant_name: payrollRoeCleanV365_(r.restaurant_name || r.restaurantName || r.name),
        brand: brand
      };
    }).filter(function(r) { return r.restaurant_number; });

    const payrolls = Repository.getAll_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS)
      .map(enrichPayrollSubmissionV361_)
      .filter(function(r) { return String(r.status || '').toUpperCase() === 'APPROVED'; })
      .map(function(r) {
        const rn = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber);
        const week = normalizePayPeriod_(r.pay_period_ending || r.payPeriodEnding || '');
        const brand = payrollRoeCleanV365_(r.brand || restaurantBrandMap[rn] || 'Unassigned Brand') || 'Unassigned Brand';
        r.restaurant_number = rn;
        r.brand = brand;
        r.payroll_batch_key = brand + '|' + week;
        r.admin_viewed_status = payrollRoeCleanV365_(r.admin_viewed_status || (r.admin_viewed_at ? 'VIEWED' : '') || 'NOT_VIEWED') || 'NOT_VIEWED';
        r.website_update_status = payrollRoeCleanV365_(r.website_update_status || (r.website_updated_at ? 'WEBSITE_UPDATED' : '') || 'NOT_UPDATED') || 'NOT_UPDATED';
        r.website_update_batch_id = payrollRoeCleanV365_(r.website_update_batch_id || r.payroll_batch_key || '');
        return r;
      });

    const roes = payrollAdminGetRowsFastV375_(payrollRoeSheetNameV365_()).map(enrichPayrollRoeApplicationV365_)
      .map(function(r) {
        const rn = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber);
        r.restaurant_number = rn;
        r.brand = payrollRoeCleanV365_(r.brand || restaurantBrandMap[rn] || 'Unassigned Brand') || 'Unassigned Brand';
        return r;
      });

    const mroMap = {};
    payrolls.forEach(function(r) { const m = payrollRoeCleanV365_(r.reviewed_by || r.updated_by || 'Unknown'); mroMap[m] = (mroMap[m] || 0) + 1; });

    return respond_(true, {
      session: sanitizeSession_(session),
      approvedPayrolls: payrolls,
      roeApplications: roes,
      restaurants: restaurants,
      summary: {
        approvedPayrolls: payrolls.length,
        roeApplications: roes.length,
        roePending: roes.filter(function(r) { return ['ROE_SUBMITTED', 'ATTACHED_TO_PAYROLL', 'REVIEWED_BY_MRO'].indexOf(normalizePayrollRoeStatusV365_(r.status)) !== -1; }).length,
        employeesDeactivated: roes.filter(function(r) { return String(r.deactivated_employee || '').toUpperCase() === 'YES'; }).length,
        payrollWebsiteUpdated: payrolls.filter(function(r) { return String(r.website_update_status || '').toUpperCase() === 'WEBSITE_UPDATED'; }).length,
        payrollAdminViewed: payrolls.filter(function(r) { return String(r.admin_viewed_status || '').toUpperCase() === 'VIEWED'; }).length,
        mroApprovedCounts: mroMap
      }
    });
  } catch (error) {
    console.error('getPayrollAdminWorkspaceData failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function payrollAdminNormalizeFiltersV367_(payload) {
  payload = payload || {};
  const tab = cleanPayrollTextV340_(payload.tab || 'BATCHES').toUpperCase();
  const dateMode = cleanPayrollTextV340_(payload.dateMode || '7d').toLowerCase();
  return {
    tab: ['BATCHES', 'APPROVED', 'ROE_PENDING', 'ROE_DONE', 'AUDIT'].indexOf(tab) !== -1 ? tab : 'BATCHES',
    page: Math.max(0, parseInt(payload.page, 10) || 0),
    pageSize: Math.max(1, Math.min(50, parseInt(payload.pageSize, 10) || 10)),
    search: cleanPayrollTextV340_(payload.search || '').toLowerCase(),
    restaurant: cleanPayrollTextV340_(payload.restaurant || 'ALL') || 'ALL',
    brand: cleanPayrollTextV340_(payload.brand || 'ALL') || 'ALL',
    dateMode: ['7d', 'month', 'ytd', 'all'].indexOf(dateMode) !== -1 ? dateMode : '7d',
    batchKey: cleanPayrollTextV340_(payload.batchKey || '')
  };
}

function payrollAdminDateCutoffV367_(dateMode) {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  if (dateMode === '7d') {
    const d = new Date(now);
    d.setDate(d.getDate() - 7);
    return d;
  }
  if (dateMode === 'month') return new Date(now.getFullYear(), now.getMonth(), 1);
  if (dateMode === 'ytd') return new Date(now.getFullYear(), 0, 1);
  return null;
}

function payrollAdminPayloadCountV367_(row, key) {
  const payload = row && row.payload ? row.payload : parsePayrollJsonV340_(row && row.payload_json, {});
  return ((payload && payload[key]) || []).length;
}


function inferPayrollSubmittedByV374_(row) {
  row = row || {};
  var direct = cleanPayrollTextV340_(row.submitted_by || row.submitted_by_name || row.created_by_name || row.created_by || '');
  if (direct) return direct;
  var history = parsePayrollJsonV340_(row.status_history_json || row.payroll_history_json || '[]', []);
  for (var i = 0; i < history.length; i++) {
    var item = history[i] || {};
    var action = String(item.action || item.status || '').toUpperCase();
    if (action === 'SUBMITTED' || action === 'RESUBMITTED' || action.indexOf('SUBMITTED') !== -1) {
      var actor = cleanPayrollTextV340_(item.updated_by || item.by || item.user || item.actor || '');
      if (actor) return actor;
    }
  }
  return '';
}


function payrollAdminGetRowsFastV375_(sheetName) {
  try {
    if (Repository && typeof Repository.getAllCached_ === 'function') {
      return Repository.getAllCached_(sheetName) || [];
    }
  } catch (e) {}
  try {
    return Repository.getAll_(sheetName) || [];
  } catch (e2) {
    console.error('Payroll admin fast row load failed for ' + sheetName, e2);
    return [];
  }
}

function payrollAdminRespondCachedV375_(cache, key, response) {
  const wrapped = respond_(true, response);
  try {
    const json = JSON.stringify(wrapped);
    if (json.length < 100000) cache.put(key, json, 90);
  } catch (e) {}
  return wrapped;
}

function payrollAdminLoadDatasetV367_() {
  ensurePayrollAdminStatusColumns_();
  const restaurantRows = Repository.getAllCached_(APP_CONFIG.sheets.RESTAURANTS)
    .filter(function(r) { return String(r.status || '').toUpperCase() !== 'INACTIVE'; });

  const restaurantBrandMap = {};
  const restaurants = restaurantRows.map(function(r) {
    const number = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber || r.restaurant_code || r.restaurantCode || r.number || r.code);
    const brand = payrollRoeCleanV365_(r.brand || r.brand_name || r.restaurant_brand || r.Brand || r.banner || r.concept || 'Unassigned Brand') || 'Unassigned Brand';
    if (number) restaurantBrandMap[number] = brand;
    return {
      restaurant_number: number,
      restaurant_name: payrollRoeCleanV365_(r.restaurant_name || r.restaurantName || r.name),
      brand: brand
    };
  }).filter(function(r) { return r.restaurant_number; }).sort(function(a, b) {
    return String(a.restaurant_number || '').localeCompare(String(b.restaurant_number || ''), undefined, { numeric: true });
  });

  const brandMap = {};
  restaurants.forEach(function(r) { brandMap[String(r.brand || 'Unassigned Brand')] = true; });

  const payrolls = payrollAdminGetRowsFastV375_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS)
    .map(enrichPayrollSubmissionV361_)
    .filter(function(r) { return String(r.status || '').toUpperCase() === 'APPROVED'; })
    .map(function(r) {
      const rn = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber);
      const week = normalizePayPeriod_(r.pay_period_ending || r.payPeriodEnding || '');
      const brand = payrollRoeCleanV365_(r.brand || restaurantBrandMap[rn] || 'Unassigned Brand') || 'Unassigned Brand';
      brandMap[brand] = true;
      r.restaurant_number = rn;
      r.brand = brand;
      r.submitted_by = inferPayrollSubmittedByV374_(r) || r.submitted_by || '';
      r.payroll_batch_key = brand + '|' + week;
      r.admin_viewed_status = payrollRoeCleanV365_(r.admin_viewed_status || (r.admin_viewed_at ? 'VIEWED' : '') || 'NOT_VIEWED') || 'NOT_VIEWED';
      r.website_update_status = payrollRoeCleanV365_(r.website_update_status || (r.website_updated_at ? 'WEBSITE_UPDATED' : '') || 'NOT_UPDATED') || 'NOT_UPDATED';
      r.website_update_batch_id = payrollRoeCleanV365_(r.website_update_batch_id || r.payroll_batch_key || '');
      return r;
    });

  const roes = payrollAdminGetRowsFastV375_(payrollRoeSheetNameV365_()).map(enrichPayrollRoeApplicationV365_)
    .map(function(r) {
      const rn = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber);
      r.restaurant_number = rn;
      r.brand = payrollRoeCleanV365_(r.brand || restaurantBrandMap[rn] || 'Unassigned Brand') || 'Unassigned Brand';
      brandMap[r.brand] = true;
      return r;
    });

  return {
    restaurants: restaurants,
    brands: Object.keys(brandMap).sort(),
    payrolls: payrolls,
    roes: roes
  };
}

function payrollAdminFilterApprovedPayrollsV367_(rows, filters) {
  const cutoff = payrollAdminDateCutoffV367_(filters.dateMode);
  return rows.filter(function(r) {
    if (filters.restaurant !== 'ALL' && String(r.restaurant_number || '') !== String(filters.restaurant)) return false;
    if (filters.brand !== 'ALL' && String(r.brand || 'Unassigned Brand') !== String(filters.brand)) return false;
    if (cutoff && filters.dateMode !== 'all') {
      const value = normalizePayPeriod_(r.pay_period_ending || r.approved_at || r.submitted_at || '');
      if (value) {
        const parts = value.split('-');
        const dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        if (dt < cutoff) return false;
      }
    }
    if (filters.search) {
      const haystack = [
        r.submission_id,
        r.restaurant_number,
        r.brand,
        r.submitted_by,
        r.reviewed_by,
        r.pay_period_ending,
        r.website_update_status,
        r.admin_viewed_status
      ].join(' ').toLowerCase();
      if (haystack.indexOf(filters.search) === -1) return false;
    }
    return true;
  });
}

function payrollAdminFilterRoesV367_(rows, filters) {
  return rows.filter(function(r) {
    if (filters.restaurant !== 'ALL' && String(r.restaurant_number || '') !== String(filters.restaurant)) return false;
    if (filters.brand !== 'ALL' && String(r.brand || 'Unassigned Brand') !== String(filters.brand)) return false;
    if (filters.search) {
      const haystack = [
        r.roe_id,
        r.restaurant_number,
        r.employee_name,
        r.reason_for_separation,
        r.status
      ].join(' ').toLowerCase();
      if (haystack.indexOf(filters.search) === -1) return false;
    }
    return true;
  });
}

function payrollAdminBuildBatchesV367_(pays, restaurants) {
  const restByBrand = {};
  restaurants.forEach(function(r) {
    const brand = cleanPayrollTextV340_(r.brand || 'Unassigned Brand') || 'Unassigned Brand';
    const rn = cleanPayrollTextV340_(r.restaurant_number || '');
    if (!rn) return;
    restByBrand[brand] = restByBrand[brand] || {};
    restByBrand[brand][rn] = true;
  });
  const map = {};
  pays.forEach(function(r) {
    const week = normalizePayPeriod_(r.pay_period_ending || '') || cleanPayrollTextV340_(r.pay_period_ending || 'No Week');
    const brand = cleanPayrollTextV340_(r.brand || 'Unassigned Brand') || 'Unassigned Brand';
    const key = brand + '|' + week;
    if (!map[key]) map[key] = { key: key, brand: brand, week: week, rows: [], restaurantSeen: {} };
    map[key].rows.push(r);
    map[key].restaurantSeen[cleanPayrollTextV340_(r.restaurant_number || '')] = true;
  });
  return Object.keys(map).map(function(key) {
    const batch = map[key];
    const brandRestaurants = restByBrand[batch.brand] ? Object.keys(restByBrand[batch.brand]).length : 0;
    const received = Object.keys(batch.restaurantSeen).length;
    batch.totalRestaurants = Math.max(brandRestaurants, received);
    batch.received = received;
    batch.pending = Math.max(0, batch.totalRestaurants - received);
    batch.updated = batch.rows.filter(function(r) { return String(r.website_update_status || '').toUpperCase() === 'WEBSITE_UPDATED'; }).length;
    batch.viewed = batch.rows.filter(function(r) { return String(r.admin_viewed_status || '').toUpperCase() === 'VIEWED'; }).length;
    batch.progress = batch.totalRestaurants ? Math.round((batch.updated / batch.totalRestaurants) * 100) : 0;
    batch.status = (batch.received > 0 && batch.updated >= batch.received && batch.pending === 0)
      ? 'COMPLETED'
      : (batch.updated > 0 ? 'PARTIAL' : 'PENDING');
    batch.rows.sort(function(a, b) {
      return String(a.restaurant_number || '').localeCompare(String(b.restaurant_number || ''), undefined, { numeric: true });
    });
    return batch;
  }).sort(function(a, b) {
    return String(b.week || '').localeCompare(String(a.week || '')) || String(a.brand || '').localeCompare(String(b.brand || ''));
  });
}

function payrollAdminBuildAuditEventsV367_(pays, roes) {
  const events = [];
  pays.forEach(function(r) {
    const history = parsePayrollJsonV340_(r.status_history_json || r.payroll_history_json, []);
    if (history && history.length) {
      history.forEach(function(x) {
        events.push({
          at: x.at || x.updated_at || x.timestamp || '',
          by: x.updated_by || x.by || '',
          action: x.action || x.status || '',
          note: x.note || '',
          ref: r.submission_id || '',
          restaurant: r.restaurant_number || ''
        });
      });
    } else {
      if (r.submitted_at) events.push({ at: r.submitted_at, by: r.submitted_by || '', action: 'SUBMITTED', note: 'Payroll submitted', ref: r.submission_id || '', restaurant: r.restaurant_number || '' });
      if (r.reviewed_at || r.admin_viewed_at) events.push({ at: r.reviewed_at || r.admin_viewed_at, by: r.reviewed_by || r.admin_viewed_by || '', action: 'VIEWED_BY_MRO', note: 'Payroll viewed', ref: r.submission_id || '', restaurant: r.restaurant_number || '' });
      if (r.approved_at || r.last_approved_at) events.push({ at: r.approved_at || r.last_approved_at, by: r.approved_by_name || r.approved_by || r.reviewed_by || '', action: 'APPROVED', note: 'Payroll approved', ref: r.submission_id || '', restaurant: r.restaurant_number || '' });
      if (r.website_updated_at) events.push({ at: r.website_updated_at, by: r.website_updated_by || '', action: 'WEBSITE_UPDATED', note: 'Payroll marked done', ref: r.submission_id || '', restaurant: r.restaurant_number || '' });
    }
  });
  roes.forEach(function(r) {
    const history = parsePayrollJsonV340_(r.history_json, []);
    if (history && history.length) {
      history.forEach(function(x) {
        events.push({
          at: x.at || x.timestamp || x.updated_at || '',
          by: x.by || x.updated_by || '',
          action: x.action || x.status || '',
          note: x.note || '',
          ref: r.roe_id || '',
          restaurant: r.restaurant_number || ''
        });
      });
    } else {
      if (r.created_at || r.submitted_at) events.push({ at: r.created_at || r.submitted_at, by: r.created_by || r.submitted_by || '', action: 'ROE_SUBMITTED', note: 'ROE submitted', ref: r.roe_id || '', restaurant: r.restaurant_number || '' });
      if (r.reviewed_at) events.push({ at: r.reviewed_at, by: r.reviewed_by || '', action: r.status || 'ROE_REVIEWED', note: r.admin_notes || '', ref: r.roe_id || '', restaurant: r.restaurant_number || '' });
      if (r.deactivated_at || r.completed_at) events.push({ at: r.deactivated_at || r.completed_at, by: r.updated_by || '', action: 'EMPLOYEE_DEACTIVATED', note: 'Employee deactivated from ROE', ref: r.roe_id || '', restaurant: r.restaurant_number || '' });
    }
  });
  return events.sort(function(a, b) {
    return new Date(b.at || 0).getTime() - new Date(a.at || 0).getTime();
  });
}

function buildPayrollAdminCsvV367_(rows) {
  const headers = ['Restaurant', 'Brand', 'Week Ending', 'Submission ID', 'Approved By', 'Approved Date', 'Viewed', 'Done Status', 'ROEs'];
  const lines = [headers.map(payrollCsvEscapeV367_).join(',')];
  rows.forEach(function(r) {
    lines.push([
      r.restaurant_number || '',
      r.brand || '',
      normalizePayPeriod_(r.pay_period_ending || ''),
      r.submission_id || '',
      r.reviewed_by || r.updated_by || '',
      r.approved_at || r.last_approved_at || '',
      String(r.admin_viewed_status || '').toUpperCase() === 'VIEWED' ? 'Viewed' : 'Not Viewed',
      String(r.website_update_status || '').toUpperCase() === 'WEBSITE_UPDATED' ? 'Done' : 'Pending',
      payrollAdminPayloadCountV367_(r, 'roeRows')
    ].map(payrollCsvEscapeV367_).join(','));
  });
  return lines.join('\n');
}



/*******************************************************
 PAYROLL ADMIN ROE SPEED FIX V376
 Loads only restaurant + ROE data for Pending ROE and Completed ROE tabs.
 This avoids loading and grouping all approved payroll submissions when the admin only opens ROE tabs.
*******************************************************/
function payrollAdminLoadRoeDatasetOnlyV376_() {
  const restaurantRows = Repository.getAllCached_(APP_CONFIG.sheets.RESTAURANTS)
    .filter(function(r) { return String(r.status || '').toUpperCase() !== 'INACTIVE'; });

  const restaurantBrandMap = {};
  const restaurants = restaurantRows.map(function(r) {
    const number = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber || r.restaurant_code || r.restaurantCode || r.number || r.code);
    const brand = payrollRoeCleanV365_(r.brand || r.brand_name || r.restaurant_brand || r.Brand || r.banner || r.concept || 'Unassigned Brand') || 'Unassigned Brand';
    if (number) restaurantBrandMap[number] = brand;
    return {
      restaurant_number: number,
      restaurant_name: payrollRoeCleanV365_(r.restaurant_name || r.restaurantName || r.name),
      brand: brand
    };
  }).filter(function(r) { return r.restaurant_number; }).sort(function(a, b) {
    return String(a.restaurant_number || '').localeCompare(String(b.restaurant_number || ''), undefined, { numeric: true });
  });

  const brandMap = {};
  restaurants.forEach(function(r) { brandMap[String(r.brand || 'Unassigned Brand')] = true; });

  const roes = payrollAdminGetRowsFastV375_(payrollRoeSheetNameV365_()).map(enrichPayrollRoeApplicationV365_)
    .map(function(r) {
      const rn = payrollRoeCleanV365_(r.restaurant_number || r.restaurantNumber);
      r.restaurant_number = rn;
      r.brand = payrollRoeCleanV365_(r.brand || restaurantBrandMap[rn] || 'Unassigned Brand') || 'Unassigned Brand';
      brandMap[r.brand] = true;
      return r;
    });

  return {
    restaurants: restaurants,
    brands: Object.keys(brandMap).sort(),
    roes: roes
  };
}

function payrollAdminActiveRestaurantCountV376_(restaurants, filters) {
  const seen = {};
  (restaurants || []).forEach(function(r) {
    const rn = cleanPayrollTextV340_(r.restaurant_number || '');
    const brand = cleanPayrollTextV340_(r.brand || 'Unassigned Brand') || 'Unassigned Brand';
    if (!rn) return;
    if (filters.restaurant !== 'ALL' && String(rn) !== String(filters.restaurant)) return;
    if (filters.brand !== 'ALL' && String(brand) !== String(filters.brand)) return;
    seen[rn] = true;
  });
  return Object.keys(seen).length;
}

function payrollAdminBuildRoeTabResponseV376_(payload, session, filters, cache, cacheKey) {
  const dataset = payrollAdminLoadRoeDatasetOnlyV376_();
  const filteredRoes = payrollAdminFilterRoesV367_(dataset.roes, filters);

  const pendingRoes = filteredRoes.filter(function(r) {
    return ['ROE_SUBMITTED', 'ATTACHED_TO_PAYROLL', 'REVIEWED_BY_MRO', 'CORRECTION_NEEDED'].indexOf(normalizePayrollRoeStatusV365_(r.status)) !== -1;
  });

  const completedRoes = filteredRoes.filter(function(r) {
    return ['ROE_APPROVED', 'COMPLETED', 'DEACTIVATED'].indexOf(normalizePayrollRoeStatusV365_(r.status)) !== -1
      || String(r.deactivated_employee || '').toUpperCase() === 'YES';
  });

  const source = filters.tab === 'ROE_PENDING' ? pendingRoes : completedRoes;
  const start = filters.page * filters.pageSize;

  const response = {
    session: sanitizeSession_(session),
    restaurants: dataset.restaurants,
    brands: dataset.brands,
    filters: filters,
    summary: {
      totalRestaurants: payrollAdminActiveRestaurantCountV376_(dataset.restaurants, filters),
      payrollsReceived: 0,
      pendingPayrolls: 0,
      payrollWebsiteUpdated: 0,
      approvedPayrolls: 0,
      pendingRoes: pendingRoes.length,
      completedRoes: completedRoes.length
    },
    tab: filters.tab,
    batchKey: filters.batchKey,
    counts: {
      approvedPayrolls: 0,
      pendingRoes: pendingRoes.length,
      completedRoes: completedRoes.length
    },
    viewMode: filters.tab === 'ROE_PENDING' ? 'roe_pending' : 'roe_done',
    page: filters.page,
    pageSize: filters.pageSize,
    total: source.length,
    rows: source.slice(start, start + filters.pageSize)
  };

  return payrollAdminRespondCachedV375_(cache, cacheKey, response);
}


function payrollAdminCacheVersionV377_() {
  try {
    return PropertiesService.getScriptProperties().getProperty('PAYROLL_ADMIN_CACHE_VERSION_V377') || '1';
  } catch (e) {
    return '1';
  }
}

function payrollAdminBumpCacheVersionV377_() {
  try {
    PropertiesService.getScriptProperties().setProperty('PAYROLL_ADMIN_CACHE_VERSION_V377', String(new Date().getTime()));
  } catch (e) {}
  try { Repository.clearCacheV258_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions'); } catch (e1) {}
  try { Repository.clearCacheV258_(payrollRoeSheetNameV365_()); } catch (e2) {}
  try { Repository.clearCacheV258_(APP_CONFIG.sheets.TEAM_MEMBERS || 'Team Members'); } catch (e3) {}
  try { Repository.clearCacheV258_(APP_CONFIG.sheets.EMPLOYEE_HISTORY || 'Employee History'); } catch (e4) {}
}

function queryPayrollAdminWorkspaceData(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken || payload);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT]);

    const filters = payrollAdminNormalizeFiltersV367_(payload);

    // Server-side cache keyed by user, tab, page, filters.
    const _qKey = [
      'q_admin_v377',
      payrollAdminCacheVersionV377_(),
      String(session.userId || '').slice(-8),
      filters.tab || 'BATCHES',
      filters.batchKey || '',
      filters.page || 0,
      filters.pageSize || 10,
      filters.dateMode || '7d',
      filters.brand || 'ALL',
      filters.restaurant || 'ALL',
      filters.search || ''
    ].join('_');
    const _qCache = CacheService.getScriptCache();
    if (!payload.forceRefresh) {
      try {
        const _qc = _qCache.get(_qKey);
        if (_qc) return JSON.parse(_qc);
      } catch (e) {}
    }

    // ROE tabs use a lightweight path. Do not load all payroll submissions here.
    if (filters.tab === 'ROE_PENDING' || filters.tab === 'ROE_DONE') {
      return payrollAdminBuildRoeTabResponseV376_(payload, session, filters, _qCache, _qKey);
    }

    const dataset = payrollAdminLoadDatasetV367_();
    const filteredPays = payrollAdminFilterApprovedPayrollsV367_(dataset.payrolls, filters);
    const filteredRoes = payrollAdminFilterRoesV367_(dataset.roes, filters);
    const pendingRoes = filteredRoes.filter(function(r) {
      return ['ROE_SUBMITTED', 'ATTACHED_TO_PAYROLL', 'REVIEWED_BY_MRO', 'CORRECTION_NEEDED'].indexOf(normalizePayrollRoeStatusV365_(r.status)) !== -1;
    });
    const completedRoes = filteredRoes.filter(function(r) {
      return ['ROE_APPROVED', 'COMPLETED', 'DEACTIVATED'].indexOf(normalizePayrollRoeStatusV365_(r.status)) !== -1
        || String(r.deactivated_employee || '').toUpperCase() === 'YES';
    });
    const batches = payrollAdminBuildBatchesV367_(filteredPays, dataset.restaurants);
    const batchSummaries = batches.map(function(batch) {
      return {
        key: batch.key,
        brand: batch.brand,
        week: batch.week,
        totalRestaurants: batch.totalRestaurants,
        received: batch.received,
        pending: batch.pending,
        updated: batch.updated,
        viewed: batch.viewed,
        progress: batch.progress,
        status: batch.status
      };
    });

    const activeRestaurantSeenV373 = {};
    dataset.restaurants.forEach(function(r) {
      const rn = cleanPayrollTextV340_(r.restaurant_number || '');
      const brand = cleanPayrollTextV340_(r.brand || 'Unassigned Brand') || 'Unassigned Brand';
      if (!rn) return;
      if (filters.restaurant !== 'ALL' && String(rn) !== String(filters.restaurant)) return;
      if (filters.brand !== 'ALL' && String(brand) !== String(filters.brand)) return;
      activeRestaurantSeenV373[rn] = true;
    });

    const latestWeekForPendingV373 = batchSummaries.length ? batchSummaries[0].week : '';
    const latestBatchesForPendingV373 = latestWeekForPendingV373
      ? batchSummaries.filter(function(item) { return item.week === latestWeekForPendingV373; })
      : [];

    const summary = {
      totalRestaurants: Object.keys(activeRestaurantSeenV373).length,
      payrollsReceived: filteredPays.length,
      pendingPayrolls: latestBatchesForPendingV373.reduce(function(total, batch) { return total + batch.pending; }, 0),
      payrollWebsiteUpdated: filteredPays.filter(function(r) { return String(r.website_update_status || '').toUpperCase() === 'WEBSITE_UPDATED'; }).length,
      approvedPayrolls: filteredPays.length,
      pendingRoes: pendingRoes.length,
      completedRoes: completedRoes.length
    };

    const response = {
      session: sanitizeSession_(session),
      restaurants: dataset.restaurants,
      brands: dataset.brands,
      filters: filters,
      summary: summary,
      tab: filters.tab,
      batchKey: filters.batchKey,
      counts: {
        approvedPayrolls: filteredPays.length,
        pendingRoes: pendingRoes.length,
        completedRoes: completedRoes.length
      }
    };

    if (filters.tab === 'BATCHES') {
      if (filters.batchKey) {
        const batch = batches.find(function(item) { return item.key === filters.batchKey; }) || null;
        response.viewMode = 'batch_detail';
        response.batch = batch;
        response.rows = batch ? batch.rows : [];
        response.total = response.rows.length;
      } else {
        const latestWeek = batchSummaries.length ? batchSummaries[0].week : '';
        response.viewMode = 'batch_list';
        response.batches = batchSummaries;
        response.latestWeek = latestWeek;
        response.latestBatches = latestWeek ? batchSummaries.filter(function(item) { return item.week === latestWeek; }) : [];
        response.otherBatches = latestWeek ? batchSummaries.filter(function(item) { return item.week !== latestWeek; }) : [];
      }
      return payrollAdminRespondCachedV375_(_qCache, _qKey, response);
    }

    if (filters.tab === 'APPROVED') {
      response.viewMode = 'approved';
      response.batches = batchSummaries;
      return payrollAdminRespondCachedV375_(_qCache, _qKey, response);
    }

    const events = payrollAdminBuildAuditEventsV367_(filteredPays, filteredRoes);
    const start = filters.page * filters.pageSize;
    response.viewMode = 'audit';
    response.page = filters.page;
    response.pageSize = filters.pageSize;
    response.total = events.length;
    response.events = events.slice(start, start + filters.pageSize);
    return payrollAdminRespondCachedV375_(_qCache, _qKey, response);

  } catch (error) {
    console.error('queryPayrollAdminWorkspaceData failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function exportPayrollAdminCsv(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT]);

    const filters = payrollAdminNormalizeFiltersV367_(payload);
    const dataset = payrollAdminLoadDatasetV367_();
    const rows = payrollAdminFilterApprovedPayrollsV367_(dataset.payrolls, filters);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    const fileName = 'payroll-admin-' + stamp + '.csv';
    const csv = buildPayrollAdminCsvV367_(rows);
    const exportFile = createPayrollExportFileV367_(fileName, csv);

    logAuditEvent_('PAYROLL_ADMIN_EXPORTED', 'PAYROLL', exportFile.getId(), session.userId, {
      fileName: exportFile.getName(),
      rowCount: rows.length,
      filters: filters
    }, session);
    return respond_(true, {
      fileId: exportFile.getId(),
      fileName: exportFile.getName(),
      fileUrl: exportFile.getUrl(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + exportFile.getId(),
      rowCount: rows.length,
      message: 'Exported ' + rows.length + ' payroll admin row(s) to CSV file "' + exportFile.getName() + '".'
    });
  } catch (error) {
    console.error('exportPayrollAdminCsv failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function returnPayrollSubmissionToRestaurant(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT]);

    const submissionIds = Array.isArray(payload.submissionIds)
      ? payload.submissionIds.map(cleanPayrollTextV340_).filter(Boolean)
      : [cleanPayrollTextV340_(payload.submissionId)].filter(Boolean);
    if (!submissionIds.length) throw new Error('Submission ID is required.');
    const submissionId = submissionIds[0];
    const note = cleanPayrollTextV340_(payload.note || payload.reason || 'Returned to restaurant for further changes');
    const now = new Date().toISOString();
    const sheetName = APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions';

    const result = withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const values = sheet.getDataRange().getValues();
      if (values.length < 2) throw new Error('Payroll submission not found.');

      const headers = values[0].map(function(h) { return String(h || '').trim(); });
      const idIndex = headers.indexOf('submission_id');
      if (idIndex === -1) throw new Error('submission_id column not found.');

      const rowIndex = values.findIndex(function(row, index) {
        return index > 0 && String(row[idIndex] || '') === submissionId;
      });
      if (rowIndex === -1) throw new Error('Payroll submission not found.');

      const current = rowToObject_(headers, values[rowIndex]);
      const currentStatus = normalizePayrollStatusV340_(current.status || 'DRAFT');
      if (currentStatus !== 'IN_REVIEW' && String(current.locked_by_mro || '').toUpperCase() !== 'YES') {
        throw new Error('Only MRO-opened payrolls can be returned to the restaurant.');
      }

      const returnStatus = cleanPayrollTextV340_(current.last_resubmitted_at) ? 'RESUBMITTED' : 'SUBMITTED';
      const history = payrollHistoryListV361_(current);
      payrollPushHistoryV361_(history, 'RETURNED_TO_RESTAURANT', note, session, { status: returnStatus, timestamp: now });

      const patch = {
        status: returnStatus,
        locked_by_mro: 'NO',
        locked_at: '',
        updated_at: now,
        updated_by: payrollActorNameV361_(session),
        status_history_json: JSON.stringify(history),
        payroll_history_json: JSON.stringify(history)
      };

      headers.forEach(function(header, index) {
        if (Object.prototype.hasOwnProperty.call(patch, header)) {
          values[rowIndex][index] = patch[header];
        }
      });
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values[rowIndex]]);
      SpreadsheetApp.flush();
      return { submissionId: submissionId, status: returnStatus };
    });

    logAuditEvent_('PAYROLL_RETURNED_TO_RESTAURANT', 'PAYROLL', submissionId, session.userId, { note: note, status: result.status }, session);
    invalidateMROPayrollDashboardCacheV367_(session.userId);
    return respond_(true, result);
  } catch (error) {
    console.error('returnPayrollSubmissionToRestaurant failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function payrollSmokeCheckSheetV367_(sheetName, requiredHeaders) {
  const sheet = Repository.getSheet_(sheetName);
  if (!sheet) {
    return {
      present: false,
      sheetName: sheetName,
      rowCount: 0,
      missingHeaders: requiredHeaders.slice()
    };
  }
  const headerCount = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, headerCount).getValues()[0].map(function(h) { return String(h || '').trim(); });
  return {
    present: true,
    sheetName: sheetName,
    rowCount: Math.max(sheet.getLastRow() - 1, 0),
    headers: headers,
    missingHeaders: requiredHeaders.filter(function(header) { return headers.indexOf(header) === -1; })
  };
}

function runPayrollSmokeCheck() {
  ensurePayrollAdminStatusColumns_();
  const report = {
    checkedAt: new Date().toISOString(),
    timeZone: Session.getScriptTimeZone(),
    ok: true,
    sheets: {},
    functions: {},
    warnings: []
  };

  report.sheets.payrollSubmissions = payrollSmokeCheckSheetV367_(APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions', [
    'submission_id', 'restaurant_number', 'pay_period_ending', 'status', 'payload_json', 'status_history_json',
    'locked_by_mro', 'admin_viewed_status', 'website_update_status'
  ]);
  report.sheets.payrollRows = payrollSmokeCheckSheetV367_(APP_CONFIG.sheets.PAYROLL_ROWS || 'Payroll Rows', [
    'submission_id', 'section_key', 'employee_id'
  ]);
  report.sheets.restaurants = payrollSmokeCheckSheetV367_(APP_CONFIG.sheets.RESTAURANTS || 'Restaurants', [
    'restaurant_number', 'brand'
  ]);
  report.sheets.teamMembers = payrollSmokeCheckSheetV367_(APP_CONFIG.sheets.TEAM_MEMBERS || 'Team Members', []);
  report.sheets.payrollRoeApplications = payrollSmokeCheckSheetV367_(payrollRoeSheetNameV365_(), [
    'roe_id', 'restaurant_number', 'status', 'employee_name'
  ]);

  if (report.sheets.teamMembers.present) {
    const teamHeaders = report.sheets.teamMembers.headers || [];
    const hasIdentityHeader = teamHeaders.indexOf('employee_id') !== -1 || teamHeaders.indexOf('team_member_id') !== -1;
    if (!hasIdentityHeader) {
      report.ok = false;
      report.warnings.push('Team Members sheet is missing both employee_id and team_member_id.');
    }
  }

  ['payrollSubmissions', 'payrollRows', 'restaurants', 'teamMembers', 'payrollRoeApplications'].forEach(function(key) {
    const check = report.sheets[key];
    if (!check.present || (check.missingHeaders && check.missingHeaders.length)) {
      report.ok = false;
    }
  });

  report.functions = {
    submitPayrollSubmission: typeof submitPayrollSubmission === 'function',
    submitPayrollOnBehalf: typeof submitPayrollOnBehalf === 'function',
    updatePayrollSubmissionStatus: typeof updatePayrollSubmissionStatus === 'function',
    recallPayrollSubmission: typeof recallPayrollSubmission === 'function',
    exportPayrollSummary: typeof exportPayrollSummary === 'function',
    queryPayrollAdminWorkspaceData: typeof queryPayrollAdminWorkspaceData === 'function',
    updatePayrollAdminWebsiteStatus: typeof updatePayrollAdminWebsiteStatus === 'function',
    returnPayrollSubmissionToRestaurant: typeof returnPayrollSubmissionToRestaurant === 'function',
    exportPayrollAdminCsv: typeof exportPayrollAdminCsv === 'function',
    updatePayrollRoeStatus: typeof updatePayrollRoeStatus === 'function',
    deactivateEmployeeFromRoe: typeof deactivateEmployeeFromRoe === 'function'
  };
  Object.keys(report.functions).forEach(function(key) {
    if (!report.functions[key]) report.ok = false;
  });

  return report;
}


function ensurePayrollAdminStatusColumns_() {
  const sheetName = APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions';
  const sheet = Repository.getSheet_(sheetName);
  if (!sheet) return;
  const needed = ['admin_viewed_status','admin_viewed_by','admin_viewed_at','website_update_status','website_updated_by','website_updated_at','website_update_batch_id','brand','payroll_batch_key','approved_by_name','approved_by','status_history_json','payroll_history_json'];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h || '').trim(); });
  let nextCol = headers.length + 1;
  let changed = false;
  needed.forEach(function(h){
    if (headers.indexOf(h) !== -1) return;
    if (sheet.getMaxColumns() < nextCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
    sheet.getRange(1, nextCol).setValue(h);
    headers.push(h);
    nextCol += 1;
    changed = true;
  });
  if (changed) {
    Repository.__headerCacheV324_ = Repository.__headerCacheV324_ || {};
    Repository.__headerCacheV324_[String(sheetName)] = headers.slice();
  }
}

function updatePayrollAdminWebsiteStatus(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const ids = Array.isArray(payload.submissionIds)
      ? payload.submissionIds.map(cleanPayrollTextV340_).filter(Boolean)
      : [cleanPayrollTextV340_(payload.submissionId)].filter(Boolean);
    if (!ids.length) throw new Error('Select at least one payroll.');

    const action = cleanPayrollTextV340_(payload.action || 'VIEWED').toUpperCase();
    if (['VIEWED', 'WEBSITE_UPDATED', 'DONE'].indexOf(action) === -1) throw new Error('Invalid payroll admin action.');

    const sheetName = APP_CONFIG.sheets.PAYROLL_SUBMISSIONS || 'Payroll Submissions';
    const now = new Date().toISOString();
    const actor = session.fullName || session.username || session.userId || '';
    const updateResult = withSheetLock_(function() {
      ensurePayrollAdminStatusColumns_();
      const sheet = Repository.getSheet_(sheetName);
      if (!sheet || sheet.getLastRow() < 2) throw new Error('Payroll Submissions sheet is empty.');

      const values = sheet.getDataRange().getValues();
      const headers = values[0].map(function(h) { return String(h || '').trim(); });
      function idx(name) { return headers.indexOf(name); }

      const idIndex = idx('submission_id');
      const statusIndex = idx('status');
      const historyIndex = idx('status_history_json') !== -1 ? idx('status_history_json') : idx('payroll_history_json');
      const payrollHistoryIndex = idx('payroll_history_json');
      const updatedAtIndex = idx('updated_at');
      const updatedByIndex = idx('updated_by');
      const adminViewedStatusIndex = idx('admin_viewed_status');
      const adminViewedByIndex = idx('admin_viewed_by');
      const adminViewedAtIndex = idx('admin_viewed_at');
      const websiteStatusIndex = idx('website_update_status');
      const websiteByIndex = idx('website_updated_by');
      const websiteAtIndex = idx('website_updated_at');
      const websiteBatchIndex = idx('website_update_batch_id');

      if (idIndex === -1) throw new Error('submission_id column not found.');

      const rowIndexById = {};
      for (let r = 1; r < values.length; r++) {
        const rowId = String(values[r][idIndex] || '');
        if (rowId && rowIndexById[rowId] == null) {
          rowIndexById[rowId] = r;
        }
      }

      const finalAction = action === 'DONE' ? 'WEBSITE_UPDATED' : action;
      const updated = [];
      const updatedRowIndexes = [];
      const missing = [];
      const audits = [];

      ids.forEach(function(id) {
        const rowIndex = rowIndexById[id];
        if (rowIndex == null) {
          missing.push(id);
          return;
        }

        const row = values[rowIndex];
        if (statusIndex !== -1 && String(row[statusIndex] || '').toUpperCase() !== 'APPROVED') {
          return;
        }

        let history = [];
        if (historyIndex !== -1) {
          try { history = JSON.parse(row[historyIndex] || '[]'); } catch (error) { history = []; }
          if (!Array.isArray(history)) history = [];
        }

        const note = finalAction === 'VIEWED' ? 'Viewed by Payroll Admin' : 'Marked done by Payroll Admin';
        history.push({
          at: now,
          by: actor,
          role: session.roleCode || '',
          status: finalAction,
          action: finalAction,
          note: note
        });

        if (updatedAtIndex !== -1) row[updatedAtIndex] = now;
        if (updatedByIndex !== -1) row[updatedByIndex] = actor;
        if (historyIndex !== -1) row[historyIndex] = JSON.stringify(history);
        if (payrollHistoryIndex !== -1) row[payrollHistoryIndex] = JSON.stringify(history);

        if (finalAction === 'VIEWED') {
          if (adminViewedStatusIndex !== -1) row[adminViewedStatusIndex] = 'VIEWED';
          if (adminViewedByIndex !== -1) row[adminViewedByIndex] = actor;
          if (adminViewedAtIndex !== -1) row[adminViewedAtIndex] = now;
        }

        if (finalAction === 'WEBSITE_UPDATED') {
          if (adminViewedStatusIndex !== -1) row[adminViewedStatusIndex] = 'VIEWED';
          if (adminViewedByIndex !== -1 && !row[adminViewedByIndex]) row[adminViewedByIndex] = actor;
          if (adminViewedAtIndex !== -1 && !row[adminViewedAtIndex]) row[adminViewedAtIndex] = now;
          if (websiteStatusIndex !== -1) row[websiteStatusIndex] = 'WEBSITE_UPDATED';
          if (websiteByIndex !== -1) row[websiteByIndex] = actor;
          if (websiteAtIndex !== -1) row[websiteAtIndex] = now;
          if (websiteBatchIndex !== -1) row[websiteBatchIndex] = cleanPayrollTextV340_(payload.batchKey || row[websiteBatchIndex] || '');
        }

        updated.push(id);
        updatedRowIndexes.push(rowIndex);
        audits.push({
          submissionId: id,
          action: finalAction
        });
      });

      if (!updated.length) {
        if (ids.length === 1 && missing.length === 1) {
          throw new Error('Payroll submission not found.');
        }
        throw new Error('No approved payrolls were updated.');
      }

      updatedRowIndexes.sort(function(a, b) { return a - b; }).forEach(function(rowIndex) {
        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values[rowIndex]]);
      });

      SpreadsheetApp.flush();
      return {
        updated: updated,
        missing: missing,
        audits: audits,
        finalAction: finalAction
      };
    });

    updateResult.audits.forEach(function(audit) {
      try {
        logAuditEvent_('PAYROLL_ADMIN_' + audit.action, 'PAYROLL', audit.submissionId, session.userId, { action: audit.action, batchKey: payload.batchKey || '' }, session);
      } catch (auditError) {
        console.error('Payroll admin audit failed', auditError);
      }
    });

    payrollAdminBumpCacheVersionV377_();

    return respond_(true, {
      updated: updateResult.updated,
      missing: updateResult.missing,
      count: updateResult.updated.length,
      action: updateResult.finalAction
    });
  } catch (error) {
    console.error('updatePayrollAdminWebsiteStatus failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function updatePayrollRoeStatus(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT]);
    const roeId = payrollRoeCleanV365_(payload.roeId);
    if (!roeId) throw new Error('ROE ID is required.');
    const status = normalizePayrollRoeStatusV365_(payload.status || 'REVIEWED_BY_MRO');
    const rows = Repository.getAll_(payrollRoeSheetNameV365_());
    const current = rows.find(function(r) { return String(r.roe_id || r.submission_id || '') === String(roeId); });
    if (!current) throw new Error('ROE application not found.');
    const now = payrollRoeNowV365_();
    let history = payrollRoeHistoryV365_(current);
    payrollRoePushHistoryV365_(history, status, payload.adminNotes || '', session, { timestamp: now });
    const patch = { status: status, admin_notes: payrollRoeCleanV365_(payload.adminNotes || current.admin_notes || ''), reviewed_by: payrollActorNameV361_(session), reviewed_at: current.reviewed_at || now, updated_at: now, updated_by: payrollActorNameV361_(session), history_json: JSON.stringify(history) };
    if (status === 'ROE_APPROVED') { patch.approved_by = payrollActorNameV361_(session); patch.approved_at = now; }
    if (status === 'COMPLETED') { patch.completed_at = now; }
    Repository.updateById_(payrollRoeSheetNameV365_(), 'roe_id', roeId, patch);
    logAuditEvent_('PAYROLL_ROE_STATUS_UPDATED', 'PAYROLL_ROE', roeId, session.userId, { status: status }, session);
    payrollAdminBumpCacheVersionV377_();
    return respond_(true, { roeId: roeId, status: status });
  } catch (error) {
    console.error('updatePayrollRoeStatus failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function deactivateEmployeeFromRoe(payload) {
  try {
    payload = payload || {};
    const session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT]);
    const roeId = payrollRoeCleanV365_(payload.roeId);
    if (!roeId) throw new Error('ROE ID is required.');
    const roes = Repository.getAll_(payrollRoeSheetNameV365_());
    const roe = roes.find(function(r) { return String(r.roe_id || r.submission_id || '') === String(roeId); });
    if (!roe) throw new Error('ROE application not found.');
    if (String(roe.status || '').toUpperCase() !== 'ROE_APPROVED' && String(roe.status || '').toUpperCase() !== 'COMPLETED') {
      throw new Error('Approve the ROE before deactivating the employee.');
    }

    const team = Repository.getAll_(APP_CONFIG.sheets.TEAM_MEMBERS);
    const empId = payrollRoeCleanV365_(roe.employee_id);
    const empName = payrollRoeCleanV365_(roe.employee_name);
    const rn = payrollRoeCleanV365_(roe.restaurant_number);
    const emp = team.find(function(m) {
      if (empId && String(m.employee_id || m.team_member_id || '') === String(empId)) return true;
      return String(m.restaurant_number || '') === String(rn) && String(m.full_name || '').toUpperCase() === String(empName).toUpperCase();
    });
    if (!emp) throw new Error('Employee not found in Team Members.');

    const teamMemberId = payrollRoeCleanV365_(emp.team_member_id);
    const employeeId = payrollRoeCleanV365_(emp.employee_id || emp.team_member_id);
    const targetColumn = teamMemberId ? 'team_member_id' : 'employee_id';
    const targetId = teamMemberId || employeeId;
    if (!targetId) throw new Error('Employee record is missing both team_member_id and employee_id.');
    const now = payrollRoeNowV365_();
    Repository.updateById_(APP_CONFIG.sheets.TEAM_MEMBERS, targetColumn, targetId, {
      status: 'INACTIVE',
      active: 'NO',
      end_date: payrollRoeCleanV365_(roe.last_day_of_work),
      separation_reason: payrollRoeCleanV365_(roe.reason_for_separation),
      roe_id: roeId,
      deactivated_at: now,
      deactivated_by: payrollActorNameV361_(session),
      updated_at: now,
      notes: (payrollRoeCleanV365_(emp.notes) ? payrollRoeCleanV365_(emp.notes) + ' | ' : '') + 'Deactivated from ROE ' + roeId
    });

    try {
      Repository.append_(APP_CONFIG.sheets.EMPLOYEE_HISTORY, {
        history_id: Utilities.getUuid(),
        employee_id: payrollRoeCleanV365_(emp.employee_id || emp.team_member_id),
        employee_name: payrollRoeCleanV365_(emp.full_name || empName),
        restaurant_number: rn,
        change_type: 'ROE_DEACTIVATION',
        old_value: payrollRoeCleanV365_(emp.status || 'ACTIVE'),
        new_value: 'INACTIVE',
        notes: 'Deactivated from ROE ' + roeId + '. Reason: ' + payrollRoeCleanV365_(roe.reason_for_separation),
        changed_by: payrollActorNameV361_(session),
        changed_at: now
      });
    } catch (historyErr) {
      console.warn('Employee history write failed for ROE deactivation', historyErr);
    }

    const history = payrollRoeHistoryV365_(roe);
    payrollRoePushHistoryV365_(history, 'EMPLOYEE_DEACTIVATED', 'Employee deactivated from active lists', session, { timestamp: now });
    Repository.updateById_(payrollRoeSheetNameV365_(), 'roe_id', roeId, {
      status: 'COMPLETED',
      completed_at: now,
      deactivated_employee: 'YES',
      deactivated_at: now,
      updated_at: now,
      updated_by: payrollActorNameV361_(session),
      history_json: JSON.stringify(history)
    });

    logAuditEvent_('EMPLOYEE_DEACTIVATED_FROM_ROE', 'TEAM_MEMBER', targetId, session.userId, { roeId: roeId, restaurantNumber: rn }, session);
    try { logAuditEvent_('PAYROLL_ROE_COMPLETED', 'PAYROLL_ROE', roeId, session.userId, { employeeId: employeeId, employeeName: empName, restaurantNumber: rn }, session); } catch (auditErr) { console.error('ROE completed audit failed', auditErr); }
    payrollAdminBumpCacheVersionV377_();
    return respond_(true, { roeId: roeId, employeeId: employeeId, status: 'INACTIVE' });
  } catch (error) {
    console.error('deactivateEmployeeFromRoe failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function markPayrollRoesAttachedV365_(submissionId, payload, session) {
  try {
    const roeRows = (payload && payload.roeRows) || [];
    if (!roeRows.length) return;
    const sheetName = payrollRoeSheetNameV365_();
    const rows = Repository.getAll_(sheetName);
    const payPeriod = payrollRoeCleanV365_(payload.payPeriodEnding);
    const now = payrollRoeNowV365_();
    roeRows.forEach(function(item) {
      const roeId = payrollRoeCleanV365_(item.roeApplicationId || item.roeId || item.submissionId);
      if (!roeId) return;
      const current = rows.find(function(r) { return String(r.roe_id || r.submission_id || '') === String(roeId); });
      if (!current) return;
      let history = payrollRoeHistoryV365_(current);
      payrollRoePushHistoryV365_(history, 'ATTACHED_TO_PAYROLL', 'Attached to payroll ' + submissionId, session, { timestamp: now });
      Repository.updateById_(sheetName, 'roe_id', roeId, {
        status: 'ATTACHED_TO_PAYROLL',
        payroll_submission_id: submissionId,
        pay_period_ending: payPeriod,
        updated_at: now,
        updated_by: payrollActorNameV361_(session),
        history_json: JSON.stringify(history)
      });
    });
  } catch (error) {
    console.error('markPayrollRoesAttachedV365_ failed', error);
  }
}

// Final submit wrappers. Last definitions win in Apps Script.
function submitPayrollSubmission(payload) {
  const result = upsertPayrollSubmissionV361_(payload, 'SUBMITTED');
  try {
    if (result && result.success && result.data && result.data.submissionId) {
      const session = getSessionOrThrow_((payload || {}).sessionToken);
      markPayrollRoesAttachedV365_(result.data.submissionId, payload || {}, session);
    }
  } catch (e) { console.error('submitPayrollSubmission ROE attach failed', e); }
  return result;
}

// ── Payroll public API wrappers ───────────────────────────────────────────
function updatePayrollSubmissionStatus(payload) {
  return legacyUpdatePayrollSubmissionStatusV361_(payload);
}

function savePayrollSubmissionDraft(payload) {
  return legacySavePayrollSubmissionDraftV361_(payload);
}

function submitPayrollOnBehalf(payload) {
  const result = upsertPayrollSubmissionV361_(payload, 'SUBMITTED');
  try {
    if (result && result.success && result.data && result.data.submissionId) {
      const session = getSessionOrThrow_((payload || {}).sessionToken);
      markPayrollRoesAttachedV365_(result.data.submissionId, payload || {}, session);
    }
  } catch (e) { console.error('submitPayrollOnBehalf ROE attach failed', e); }
  return result;
}

/************************************************************
 PAYROLL V366 RESTAURANT EMPLOYEE MAPPING FIX
 Last definition wins in Apps Script.
 Fixes employee dropdowns by mapping restaurant_id and restaurant_number.
************************************************************/
function payrollRestaurantMapsV366_() {
  const rows = Repository.getAllCached_(APP_CONFIG.sheets.RESTAURANTS || 'Restaurants');
  const byKey = {};
  const restaurants = rows.filter(function(r) {
    return String(r.status || '').toUpperCase() !== 'INACTIVE';
  }).map(function(r) {
    const id = cleanPayrollTextV340_(r.restaurant_id || r.id || '');
    const number = cleanPayrollTextV340_(r.restaurant_number || r.restaurant_code || r.number || id);
    const name = cleanPayrollTextV340_(r.restaurant_name || r.name || '');
    if (id) byKey[String(id).toUpperCase()] = number;
    if (number) byKey[String(number).toUpperCase()] = number;
    const code = cleanPayrollTextV340_(r.restaurant_code || '');
    if (code) byKey[String(code).toUpperCase()] = number;
    return { restaurant_id: id, restaurant_number: number, restaurant_name: name };
  });
  return { restaurants: restaurants, byKey: byKey };
}
function payrollMapRestaurantNumberV366_(value, maps) {
  value = cleanPayrollTextV340_(value);
  if (!value) return '';
  maps = maps || payrollRestaurantMapsV366_();
  return maps.byKey[String(value).toUpperCase()] || value;
}
function payrollEmployeeRestaurantNumbersV366_(row, maps) {
  maps = maps || payrollRestaurantMapsV366_();
  const raw = [
    row.restaurant_number,
    row.restaurantNumber,
    row.restaurant_code,
    row.restaurantCode,
    row.assigned_restaurant,
    row.assignedRestaurant,
    row.assigned_restaurant_number,
    row.assignedRestaurantNumber,
    row.restaurant_id,
    row.restaurantId
  ];
  const out = [];
  raw.forEach(function(v) {
    String(v || '').split(',').forEach(function(part) {
      const mapped = payrollMapRestaurantNumberV366_(part, maps);
      if (mapped && out.indexOf(mapped) === -1) out.push(mapped);
    });
  });
  return out;
}
function payrollIsActiveEmployeeV366_(row) {
  const st = String(row.status || row.employee_status || 'ACTIVE').toUpperCase();
  return ['INACTIVE', 'ARCHIVED', 'QUIT', 'DID NOT SHOW UP', 'DIDNT SHOW UP', 'TERMINATED', 'DEACTIVATED'].indexOf(st) === -1;
}
function payrollBuildEmployeeV366_(row, maps) {
  const fullName = cleanPayrollTextV340_(row.full_name || row.employee_name || row.name || row.team_member_name || ((row.first_name || '') + ' ' + (row.last_name || '')).trim());
  const parts = fullName.split(' ');
  const numbers = payrollEmployeeRestaurantNumbersV366_(row, maps);
  return {
    employee_id: cleanPayrollTextV340_(row.employee_id || row.team_member_id || row.id || fullName),
    full_name: fullName,
    first_name: cleanPayrollTextV340_(row.first_name || parts[0] || ''),
    last_name: cleanPayrollTextV340_(row.last_name || parts.slice(1).join(' ')),
    restaurant_number: cleanPayrollTextV340_(numbers[0] || ''),
    restaurant_numbers: numbers.join(','),
    restaurant_id: cleanPayrollTextV340_(row.restaurant_id || row.restaurantId || ''),
    job_title: cleanPayrollTextV340_(row.job_title || ''),
    status: cleanPayrollTextV340_(row.status || row.employee_status || 'ACTIVE')
  };
}
function payrollEmployeesForNumbersV366_(restaurantNumbers) {
  const maps = payrollRestaurantMapsV366_();
  const wanted = (Array.isArray(restaurantNumbers) ? restaurantNumbers : [restaurantNumbers])
    .map(function(v) { return payrollMapRestaurantNumberV366_(v, maps); })
    .filter(Boolean)
    .map(function(v) { return String(v).toUpperCase(); });
  const seen = {};
  return Repository.getAllCached_(APP_CONFIG.sheets.TEAM_MEMBERS || 'Team Members')
    .filter(payrollIsActiveEmployeeV366_)
    .map(function(row) { return payrollBuildEmployeeV366_(row, maps); })
    .filter(function(emp) {
      if (!emp.employee_id || seen[emp.employee_id]) return false;
      seen[emp.employee_id] = true;
      if (!wanted.length) return true;
      const empNumbers = String(emp.restaurant_numbers || emp.restaurant_number || '').split(',').map(function(v) { return String(v).trim().toUpperCase(); }).filter(Boolean);
      return empNumbers.some(function(v) { return wanted.indexOf(v) !== -1; });
    })
    .sort(function(a, b) { return String(a.full_name || '').localeCompare(String(b.full_name || '')); });
}

function getPayrollSubmissionWorkspaceData(sessionToken) {
  try {
    const session = getSessionOrThrow_(sessionToken);
    // V22: Do not cache submission workspace on server.
    // My Submissions must show live status after MRO approves, returns, or moves payroll to In Review.
    assertAuthorized_(session, [
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    const isAdmin = payrollIsAdminSessionV361_ ? payrollIsAdminSessionV361_(session) : [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.MRO, APP_CONFIG.roles.SENIOR_MANAGEMENT].indexOf(session.roleCode) !== -1;
    const maps = payrollRestaurantMapsV366_();
    const restaurantNumber = payrollMapRestaurantNumberV366_(getPrimaryRestaurantNumberForSessionV340_(session), maps);
    const restaurants = isAdmin ? maps.restaurants : maps.restaurants.filter(function(r) { return String(r.restaurant_number) === String(restaurantNumber); });
    const scopeNumbers = restaurants.map(function(r) { return r.restaurant_number; });
    const employees = isAdmin ? payrollEmployeesForNumbersV366_(scopeNumbers) : payrollEmployeesForNumbersV366_(restaurantNumber);
    const submissions = getPayrollSubmissionsForSessionV340_(session, restaurantNumber);

    return respond_(true, {
      session: sanitizeSession_(session),
      restaurantNumber: restaurantNumber,
      restaurants: restaurants,
      employees: employees,
      submissions: submissions,
      isMRORole: isAdmin
    });

  } catch (error) {
    console.error('getPayrollSubmissionWorkspaceData V366 failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function getPayrollReviewWorkspaceData(sessionToken) {
  try {
    const session = getSessionOrThrow_(sessionToken);

    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    // Important fix. Do not cache Payroll Review workspace.
    // Review Inbox must always read live sheet data so approved payrolls
    // do not come back to Inbox after Refresh.
    const maps = payrollRestaurantMapsV366_();
    const submissions = getPayrollSubmissionsForAdminV340_();
    const employees = payrollEmployeesForNumbersV366_(maps.restaurants.map(function(r) {
      return r.restaurant_number;
    }));

    return respond_(true, {
      session: sanitizeSession_(session),
      restaurants: maps.restaurants,
      employees: employees,
      submissions: submissions
    });
  } catch (error) {
    console.error('getPayrollReviewWorkspaceData V366 failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}


/**
 * V419 Admin Request row repair.
 * Fixes old rows where payload JSON was saved inside submitted_at or other visible columns.
 * Also keeps submitted_at as a real date.
 */
function repairAdminRequestRowsV419() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Admin Requests', 'Requests Admin'].forEach(function(tabName) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    const headers = values[0].map(function(h) { return String(h || '').trim(); });
    const idx = {};
    headers.forEach(function(h, i) {
      idx[normalizeHeaderV419_(h)] = i;
    });

    function col(name) {
      return idx[normalizeHeaderV419_(name)];
    }

    const submittedAtCol = col('submitted_at');
    const requestTypeCol = col('request_type');
    const restaurantIdCol = col('restaurant_id');
    const restaurantNumberCol = col('restaurant_number');
    const supplyTypeCol = col('supply_type');
    const supplyDetailsCol = col('supply_details');
    const issuerNotesCol = col('issuer_notes');
    const createdAtCol = col('created_at');
    const updatedAtCol = col('updated_at');

    let changed = false;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];

      let payload = null;

      // Look for JSON accidentally saved in any cell.
      for (let c = 0; c < row.length; c++) {
        const parsed = parseJsonPayloadV419_(row[c]);
        if (parsed) {
          payload = parsed;
          break;
        }
      }

      if (!payload) continue;

      if (requestTypeCol !== undefined && !String(row[requestTypeCol] || '').trim()) {
        row[requestTypeCol] = payload.supplyType || payload.supplyDetails ? 'OFFICE_SUPPLIES' : (payload.letterType ? 'JOB_LETTER' : '');
        changed = true;
      }

      if (restaurantIdCol !== undefined && !String(row[restaurantIdCol] || '').trim() && payload.restaurantId) {
        row[restaurantIdCol] = payload.restaurantId;
        changed = true;
      }

      if (restaurantNumberCol !== undefined && !String(row[restaurantNumberCol] || '').trim() && payload.restaurantId) {
        row[restaurantNumberCol] = String(payload.restaurantId).replace(/^REST[_-]?/i, '');
        changed = true;
      }

      if (supplyTypeCol !== undefined && !String(row[supplyTypeCol] || '').trim() && payload.supplyType) {
        row[supplyTypeCol] = payload.supplyType;
        changed = true;
      }

      if (supplyDetailsCol !== undefined && !String(row[supplyDetailsCol] || '').trim() && payload.supplyDetails) {
        row[supplyDetailsCol] = payload.supplyDetails;
        changed = true;
      }

      if (issuerNotesCol !== undefined && !String(row[issuerNotesCol] || '').trim() && payload.issuerNotes) {
        row[issuerNotesCol] = payload.issuerNotes;
        changed = true;
      }

      // If submitted_at has JSON, replace with created_at or current date.
      if (submittedAtCol !== undefined && parseJsonPayloadV419_(row[submittedAtCol])) {
        row[submittedAtCol] = createdAtCol !== undefined && row[createdAtCol] ? row[createdAtCol] : new Date();
        changed = true;
      }

      // If created_at has JSON, replace with submitted_at or current date.
      if (createdAtCol !== undefined && parseJsonPayloadV419_(row[createdAtCol])) {
        row[createdAtCol] = submittedAtCol !== undefined && row[submittedAtCol] && !parseJsonPayloadV419_(row[submittedAtCol]) ? row[submittedAtCol] : new Date();
        changed = true;
      }

      if (updatedAtCol !== undefined && parseJsonPayloadV419_(row[updatedAtCol])) {
        row[updatedAtCol] = new Date();
        changed = true;
      }
    }

    if (changed) {
      removeFilterSafeV419_(sheet);
      sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      formatAdminRequestDateColumnsV419_(sheet);
      Logger.log(tabName + ' repaired.');
    } else {
      Logger.log(tabName + ' no repair needed.');
    }
  });
}


function parseJsonPayloadV419_(value) {
  const text = String(value || '').trim();
  if (!text || text[0] !== '{') return null;

  try {
    const obj = JSON.parse(text);
    if (
      obj &&
      typeof obj === 'object' &&
      !Array.isArray(obj) &&
      (obj.supplyType || obj.supplyDetails || obj.restaurantId || obj.letterType || obj.employeeName)
    ) {
      return obj;
    }
  } catch (error) {}

  return null;
}


function normalizeHeaderV419_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
}


function removeFilterSafeV419_(sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (error) {}
}


function formatAdminRequestDateColumnsV419_(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  headers.forEach(function(header, i) {
    const h = normalizeHeaderV419_(header);
    if (['created_at', 'updated_at', 'submitted_at', 'status_updated_at', 'received_at'].indexOf(h) !== -1) {
      sheet.getRange(2, i + 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm AM/PM');
    }
  });
}



/**
 * V420 Admin Request hard repair.
 * This overrides ensureAdminRequestsSheetV213_ so future saves use the correct columns.
 * It also repairs old rows where submitted_at has payload JSON.
 */
function ensureAdminRequestsSheetV213_() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const sheetName = APP_CONFIG.sheets.REQUESTS_ADMIN;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, getAdminRequestHeadersV420_().length).setValues([getAdminRequestHeadersV420_()]);
    sheet.setFrozenRows(1);
    formatAdminRequestDateColumnsV420_(sheet);
    return sheet;
  }

  repairAdminRequestSheetStructureV420_(sheet);
  return sheet;
}


function getAdminRequestHeadersV420_() {
  return [
    'request_id',
    'request_type',
    'restaurant_id',
    'restaurant_number',
    'requested_by',
    'requested_by_name',
    'letter_type',
    'employee_id',
    'employee_name',
    'employee_status',
    'gender',
    'position',
    'manual_position',
    'employment_type',
    'started_date',
    'supply_type',
    'supply_details',
    'issuer_notes',
    'admin_notes',
    'status',
    'status_updated_by',
    'status_updated_at',
    'status_history_json',
    'submitted_at',
    'updated_at',
    'request_payload_json'
  ];
}


function repairAdminRequestSheetStructureV420_(sheet) {
  const requiredHeaders = getAdminRequestHeadersV420_();
  const values = sheet.getDataRange().getValues();

  if (!values.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    formatAdminRequestDateColumnsV420_(sheet);
    return;
  }

  const oldHeaders = values[0].map(function(h) {
    return String(h || '').trim();
  });

  const oldRows = values.length > 1 ? values.slice(1) : [];
  const newRows = [];

  oldRows.forEach(function(row) {
    if (!row.some(function(cell) { return String(cell || '').trim() !== ''; })) return;

    const obj = {};
    oldHeaders.forEach(function(header, index) {
      if (!header) return;
      if (obj[header] === undefined || obj[header] === '') {
        obj[header] = row[index];
      }
    });

    const payload = findAdminRequestPayloadV420_(row, oldHeaders);

    const requestId = cleanAdminCellV420_(
      getAnyAdminValueV420_(obj, ['request_id', 'requestId', 'id']),
      payload.requestId || payload.request_id || Utilities.getUuid()
    );

    const requestType = cleanAdminCellV420_(
      getAnyAdminValueV420_(obj, ['request_type', 'requestType']),
      inferAdminRequestTypeV420_(obj, payload)
    ) || inferAdminRequestTypeV420_(obj, payload);

    const restaurantId = cleanAdminCellV420_(
      getAnyAdminValueV420_(obj, ['restaurant_id', 'restaurantId']),
      payload.restaurantId || payload.restaurant_id || ''
    );

    const restaurantNumber = cleanAdminCellV420_(
      getAnyAdminValueV420_(obj, ['restaurant_number', 'restaurantNumber', 'restaurant_code', 'restaurantCode']),
      payload.restaurantNumber || payload.restaurant_number || String(restaurantId || '').replace(/^REST[_-]?/i, '')
    );

    const submittedAt = cleanDateValueV420_(
      getAnyAdminValueV420_(obj, ['submitted_at', 'submittedAt']),
      getAnyAdminValueV420_(obj, ['created_at', 'createdAt']) ||
      getAnyAdminValueV420_(obj, ['updated_at', 'updatedAt']) ||
      new Date()
    );

    const updatedAt = cleanDateValueV420_(
      getAnyAdminValueV420_(obj, ['updated_at', 'updatedAt']),
      submittedAt || new Date()
    );

    let statusHistory = cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['status_history_json', 'statusHistoryJson']), '');
    if (!statusHistory) {
      statusHistory = JSON.stringify([{
        at: submittedAt instanceof Date ? submittedAt.toISOString() : String(submittedAt || new Date()),
        by: cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['requested_by_name', 'requestedByName']), ''),
        status: cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['status']), 'New Request') || 'New Request',
        note: 'Request created'
      }]);
    }

    const payloadJson = Object.keys(payload).length
      ? JSON.stringify(payload)
      : cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['request_payload_json', 'payload_json', 'payloadJson']), '');

    newRows.push([
      requestId,
      requestType,
      restaurantId,
      restaurantNumber,
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['requested_by', 'requestedBy']), ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['requested_by_name', 'requestedByName']), ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['letter_type', 'letterType']), payload.letterType || payload.letter_type || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['employee_id', 'employeeId']), payload.employeeId || payload.employee_id || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['employee_name', 'employeeName']), payload.employeeName || payload.employee_name || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['employee_status', 'employeeStatus']), payload.employeeStatus || payload.employee_status || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['gender']), payload.gender || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['position']), payload.position || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['manual_position', 'manualPosition']), payload.manualPosition || payload.manual_position || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['employment_type', 'employmentType']), payload.employmentType || payload.employment_type || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['started_date', 'startedDate']), payload.startedDate || payload.started_date || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['supply_type', 'supplyType']), payload.supplyType || payload.supply_type || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['supply_details', 'supplyDetails']), payload.supplyDetails || payload.supply_details || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['issuer_notes', 'issuerNotes', 'notes']), payload.issuerNotes || payload.issuer_notes || ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['admin_notes', 'adminNotes']), ''),
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['status']), payload.status || 'New Request') || 'New Request',
      cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['status_updated_by', 'statusUpdatedBy']), ''),
      cleanDateValueV420_(getAnyAdminValueV420_(obj, ['status_updated_at', 'statusUpdatedAt']), ''),
      statusHistory,
      submittedAt,
      updatedAt,
      payloadJson
    ]);
  });

  removeFilterSafeV420_(sheet);
  sheet.clear();

  sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);

  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, requiredHeaders.length).setValues(newRows);
  }

  sheet.setFrozenRows(1);
  formatAdminRequestDateColumnsV420_(sheet);
}


function repairAdminRequestRowsV420() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const sheetName = APP_CONFIG.sheets.REQUESTS_ADMIN;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Requests Admin sheet not found.');
  repairAdminRequestSheetStructureV420_(sheet);
  Logger.log('Requests Admin repaired.');
}


function findAdminRequestPayloadV420_(row, headers) {
  for (let i = 0; i < row.length; i++) {
    const text = String(row[i] || '').trim();
    if (!text || text.charAt(0) !== '{') continue;

    try {
      const parsed = JSON.parse(text);
      if (
        parsed &&
        typeof parsed === 'object' &&
        !Array.isArray(parsed) &&
        (parsed.supplyType || parsed.supplyDetails || parsed.restaurantId || parsed.letterType || parsed.employeeName)
      ) {
        return parsed;
      }
    } catch (error) {}
  }
  return {};
}


function inferAdminRequestTypeV420_(obj, payload) {
  const raw = cleanAdminCellV420_(getAnyAdminValueV420_(obj, ['request_type', 'requestType']), payload.requestType || payload.request_type || '');
  if (raw) return raw;

  if (
    payload.supplyType ||
    payload.supplyDetails ||
    getAnyAdminValueV420_(obj, ['supply_type', 'supplyType']) ||
    getAnyAdminValueV420_(obj, ['supply_details', 'supplyDetails'])
  ) {
    return 'OFFICE_SUPPLIES';
  }

  if (
    payload.letterType ||
    payload.employeeName ||
    getAnyAdminValueV420_(obj, ['letter_type', 'letterType']) ||
    getAnyAdminValueV420_(obj, ['employee_name', 'employeeName'])
  ) {
    return 'JOB_LETTER';
  }

  return '';
}


function getAnyAdminValueV420_(obj, keys) {
  obj = obj || {};
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (obj[key] !== undefined && obj[key] !== null && String(obj[key]).trim() !== '') {
      return obj[key];
    }
  }

  const normalized = {};
  Object.keys(obj).forEach(function(key) {
    normalized[normalizeHeaderV420_(key)] = obj[key];
  });

  for (let j = 0; j < keys.length; j++) {
    const n = normalizeHeaderV420_(keys[j]);
    if (normalized[n] !== undefined && normalized[n] !== null && String(normalized[n]).trim() !== '') {
      return normalized[n];
    }
  }

  return '';
}


function cleanAdminCellV420_(value, fallback) {
  const text = String(value || '').trim();
  if (!text) return fallback || '';

  if (text.charAt(0) === '{' || text.charAt(0) === '[') {
    try {
      JSON.parse(text);
      return fallback || '';
    } catch (error) {}
  }

  return value;
}


function cleanDateValueV420_(value, fallback) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const text = String(value || '').trim();
  if (!text) return fallback || '';

  if (text.charAt(0) === '{' || text.charAt(0) === '[') {
    try {
      JSON.parse(text);
      return fallback || new Date();
    } catch (error) {}
  }

  const d = new Date(text);
  if (!isNaN(d.getTime())) return d;

  return fallback || '';
}


function normalizeHeaderV420_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
}


function removeFilterSafeV420_(sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (error) {}
}


function formatAdminRequestDateColumnsV420_(sheet) {
  const headers = getAdminRequestHeadersV420_();
  const dateHeaders = ['submitted_at', 'updated_at', 'status_updated_at'];
  const lastRow = Math.max(sheet.getLastRow(), 2);

  dateHeaders.forEach(function(header) {
    const index = headers.indexOf(header);
    if (index !== -1) {
      sheet.getRange(2, index + 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm AM/PM');
    }
  });

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#163a63')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setWrap(true);

  headers.forEach(function(header, i) {
    const width = header.indexOf('json') !== -1 || header.indexOf('notes') !== -1 || header.indexOf('details') !== -1 ? 320 :
      header.indexOf('date') !== -1 || header.indexOf('_at') !== -1 ? 180 :
      header.indexOf('restaurant') !== -1 ? 170 :
      header.indexOf('name') !== -1 ? 220 : 150;
    sheet.setColumnWidth(i + 1, width);
  });

  removeFilterSafeV420_(sheet);
  sheet.getRange(1, 1, lastRow, headers.length).createFilter();
}



/*******************************************************
 V421 ADMIN REQUEST RESTAURANT RESPONSE WORKFLOW
 Safe patch only for Admin Request / Restaurant Request.
 Payroll files are not changed.
 Adds Restaurant Responded flow for Waiting for Restaurant status.
*******************************************************/



function normalizeAdminRequestStatusV421_(status) {
  const raw = String(status || '').trim();
  const key = raw.toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  const map = {
    NEW_REQUEST: 'New Request',
    VIEWED: 'Viewed',
    IN_REVIEW: 'In Review',
    IN_PROGRESS: 'In Progress',
    WAITING_FOR_RESTAURANT: 'Waiting for Restaurant',
    WAITING_RESTAURANT: 'Waiting for Restaurant',
    RESTAURANT_RESPONDED: 'Restaurant Responded',
    COMPLETED: 'Completed',
    RECEIVED: 'Received'
  };
  return map[key] || raw || 'New Request';
}


function submitAdminRequestRestaurantResponse(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  assertAuthorized_(session, [
    APP_CONFIG.roles.RESTAURANT_USER,
    APP_CONFIG.roles.SUPER_ADMIN,
    APP_CONFIG.roles.HR_ADMIN,
    APP_CONFIG.roles.MRO,
    APP_CONFIG.roles.SENIOR_MANAGEMENT
  ]);

  const requestId = String(payload.requestId || '').trim();
  const responseNotes = String(payload.responseNotes || payload.restaurantResponseNotes || payload.notes || '').trim();
  if (!requestId) throw new Error('Missing request id.');
  if (!responseNotes) throw new Error('Please enter your response before submitting.');

  const sheet = ensureAdminRequestsSheetV213_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(header) { return String(header || '').trim(); });

  const idIndex = headers.indexOf('request_id');
  if (idIndex === -1) throw new Error('request_id column not found.');

  const statusIndex = headers.indexOf('status');
  const responseNotesIndex = headers.indexOf('restaurant_response_notes');
  const respondedByIndex = headers.indexOf('restaurant_responded_by');
  const respondedAtIndex = headers.indexOf('restaurant_responded_at');
  const updatedByIndex = headers.indexOf('status_updated_by');
  const updatedAtIndex = headers.indexOf('status_updated_at');
  const updatedIndex = headers.indexOf('updated_at');
  const historyIndex = headers.indexOf('status_history_json');
  const restaurantIdIndex = headers.indexOf('restaurant_id');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idIndex] || '') !== requestId) continue;

    const rowRestaurantId = restaurantIdIndex !== -1 ? String(values[r][restaurantIdIndex] || '').trim() : '';
    const scopeIds = getSessionRestaurantScopeV318_(session);
    if (scopeIds !== null && rowRestaurantId && scopeIds.indexOf(rowRestaurantId) === -1) {
      throw new Error('You do not have access to this request.');
    }

    const currentStatus = normalizeAdminRequestStatusV421_(statusIndex !== -1 ? values[r][statusIndex] : '');
    if (currentStatus !== 'Waiting for Restaurant') {
      throw new Error('This request is not waiting for a restaurant response.');
    }

    const now = new Date();
    const nowIso = now.toISOString();
    const userName = session.fullName || session.username || session.userId || '';
    let history = [];
    try { history = historyIndex !== -1 ? JSON.parse(values[r][historyIndex] || '[]') : []; } catch (error) { history = []; }

    history.push({
      at: nowIso,
      by: userName,
      status: 'Restaurant Responded',
      note: responseNotes
    });

    if (statusIndex !== -1) values[r][statusIndex] = 'Restaurant Responded';
    if (responseNotesIndex !== -1) values[r][responseNotesIndex] = responseNotes;
    if (respondedByIndex !== -1) values[r][respondedByIndex] = userName;
    if (respondedAtIndex !== -1) values[r][respondedAtIndex] = now;
    if (updatedByIndex !== -1) values[r][updatedByIndex] = userName;
    if (updatedAtIndex !== -1) values[r][updatedAtIndex] = now;
    if (updatedIndex !== -1) values[r][updatedIndex] = now;
    if (historyIndex !== -1) values[r][historyIndex] = JSON.stringify(history);

    sheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);
    try { formatAdminRequestDateColumnsV418_(sheet); } catch (error2) {}
    try { sethInvalidateRequestCaches_(); } catch (error3) {}
    logAuditEvent_('ADMIN_REQUEST_RESTAURANT_RESPONDED', 'ADMIN_REQUEST', requestId, session.userId, { note: responseNotes });

    return respond_(true, {
      requestId: requestId,
      status: 'Restaurant Responded',
      responseNotes: responseNotes,
      respondedBy: userName,
      respondedAt: nowIso
    });
  }

  throw new Error('Admin request not found: ' + requestId);
}

try {
  if (typeof ServiceRouter_ !== 'undefined' && ServiceRouter_ && typeof ServiceRouter_.route === 'function' && !ServiceRouter_._restaurantResponsePatchV421) {
    ServiceRouter_._restaurantResponsePatchV421 = true;
    ServiceRouter_._routeBeforeRestaurantResponseV421 = ServiceRouter_.route;
    ServiceRouter_.route = function(action, payload, token) {
      if (action === 'submitAdminRequestRestaurantResponse') {
        return submitAdminRequestRestaurantResponse(payload);
      }
      return ServiceRouter_._routeBeforeRestaurantResponseV421.call(ServiceRouter_, action, payload, token);
    };
  }
} catch (error) {
  console.error('Restaurant response route patch failed', error);
}


/*******************************************************
 V422 ADMIN REQUEST RESTAURANT RESUBMIT WORKFLOW
 Safe patch only for Admin Request / Restaurant Request.
 Keeps the same request id when restaurant responds.
 Allows editing/resubmitting only when status is Waiting for Restaurant.
*******************************************************/


function normalizeAdminRequestStatusV422_(status) {
  var raw = String(status || '').trim();
  var key = raw.toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  var map = {
    NEW_REQUEST: 'New Request',
    VIEWED: 'Viewed',
    IN_REVIEW: 'In Review',
    IN_PROGRESS: 'In Progress',
    WAITING_FOR_RESTAURANT: 'Waiting for Restaurant',
    WAITING_RESTAURANT: 'Waiting for Restaurant',
    ACTION_REQUIRED: 'Waiting for Restaurant',
    RESTAURANT_RESPONDED: 'Restaurant Responded',
    RESUBMITTED: 'Restaurant Responded',
    COMPLETED: 'Completed',
    RECEIVED: 'Received'
  };
  return map[key] || raw || 'New Request';
}

function parseAdminRequestJsonSafeV422_(value) {
  try {
    var obj = JSON.parse(String(value || '{}'));
    return obj && typeof obj === 'object' ? obj : {};
  } catch (error) {
    return {};
  }
}

function setIfHeaderExistsV422_(row, headers, headerName, value) {
  var index = headers.indexOf(headerName);
  if (index !== -1) row[index] = value;
}

function getHeaderValueV422_(row, headers, headerName) {
  var index = headers.indexOf(headerName);
  return index === -1 ? '' : row[index];
}

function mergeAdminRequestPayloadV422_(existingPayload, updates) {
  var payload = existingPayload || {};
  Object.keys(updates || {}).forEach(function(key) {
    if (updates[key] !== undefined) payload[key] = updates[key];
  });
  return payload;
}



try {
  if (typeof ServiceRouter_ !== 'undefined' && ServiceRouter_ && typeof ServiceRouter_.route === 'function' && !ServiceRouter_._restaurantResubmitPatchV422) {
    ServiceRouter_._restaurantResubmitPatchV422 = true;
    ServiceRouter_._routeBeforeRestaurantResubmitV422 = ServiceRouter_.route;
    ServiceRouter_.route = function(action, payload, token) {
      if (action === 'submitAdminRequestRestaurantResponse') {
        return submitAdminRequestRestaurantResponse(payload || {});
      }
      if (action === 'updateAdminRequestStatus') {
        return updateAdminRequestStatus(payload || {});
      }
      return ServiceRouter_._routeBeforeRestaurantResubmitV422.call(ServiceRouter_, action, payload, token);
    };
  }
} catch (error) {
  console.error('Restaurant resubmit route patch failed', error);
}

// ── MISSING FUNCTION IMPLEMENTATIONS (added by fix pass) ─────────────────────

function changeOwnPortalPassword(payload) {
  try {
    var token = payload && payload.sessionToken;
    var session = getSessionOrThrow_(token);
    var currentPassword = String(payload.currentPassword || '').trim();
    var newPassword = String(payload.newPassword || '').trim();
    if (!currentPassword || !newPassword) {
      return respond_(false, null, 'Current password and new password are required.');
    }
    if (newPassword.length < 8) {
      return respond_(false, null, 'New password must be at least 8 characters.');
    }
    // Verify current password matches stored hash
    var users = Repository.getAll_(APP_CONFIG.sheets.USERS);
    var user = users.find(function(u) {
      return String(u.userId || u.user_id || '') === String(session.userId || '');
    });
    if (!user) return respond_(false, null, 'User not found.');
    // FIX #BUG2+3: field is password_hash (not password); use updateById_ (not the non-existent updateRow_)
    var storedPassword = String(user.password_hash || '');
    var storedSalt = String(user.salt || user.password_salt || '');
    var currentHashed = (typeof hashPassword_ === 'function')
      ? hashPassword_(currentPassword, storedSalt)
      : currentPassword;
    if (storedPassword && currentHashed !== storedPassword) {
      return respond_(false, null, 'Current password is incorrect.');
    }
    // Update password
    var hashed = (typeof hashPassword_ === 'function') ? hashPassword_(newPassword, storedSalt) : newPassword;
    Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', String(user.user_id || user.userId), { password_hash: hashed, must_reset_password: 'FALSE', updated_at: new Date().toISOString() });
    invalidateAllUserSessions_(session.userId, token);
    bustBootCacheForToken_(token);
    return respond_(true, { message: 'Password updated successfully.' });
  } catch (error) {
    console.error('changeOwnPortalPassword failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function setNewPasswordOnFirstLogin(payload) {
  try {
    var token = payload && payload.sessionToken;
    var session = getSessionOrThrow_(token);
    var newPassword = String(payload.newPassword || '').trim();
    if (newPassword.length < 8) return respond_(false, null, 'Password must be at least 8 characters.');
    var user = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'user_id', session.userId);
    if (!user) return respond_(false, null, 'User not found.');
    if (String(user.must_reset_password || '').toUpperCase() !== 'TRUE') {
      return respond_(false, null, 'Password reset not required.');
    }
    var salt = String(user.salt || user.password_salt || '');
    var hashed = hashPassword_(newPassword, salt);
    Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', session.userId, {
      password_hash: hashed, must_reset_password: 'FALSE',
      updated_at: new Date().toISOString(), updated_by: session.userId
    });
    invalidateAllUserSessions_(session.userId, token);
    bustBootCacheForToken_(token);
    logAuditEvent_('PASSWORD_CHANGED_FIRST_LOGIN', 'USER', session.userId, session.userId, {});
    return respond_(true, { message: 'Password updated.' });
  } catch (error) {
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function saveDropdownOptionRecord(payload) {
  try {
    var token = payload && payload.sessionToken;
    var session = getSessionOrThrow_(token);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
    var record = payload && payload.record;
    if (!record) return respond_(false, null, 'Record is required.');
    var sheet = APP_CONFIG.sheets.DROPDOWN_OPTIONS || 'DropdownOptions';
    var optionId = String(record.optionId || record.option_id || '').trim();
    if (optionId) {
      // Update existing — FIX: use updateById_ (updateRow_ does not exist)
      Repository.updateById_(sheet, 'option_id', optionId, {
        dropdown_key: String(record.dropdownKey || record.dropdown_key || ''),
        option_value: String(record.optionValue || record.option_value || ''),
        sort_order: String(record.sortOrder || record.sort_order || '100'),
        status: String(record.status || 'ACTIVE'),
        updated_at: new Date().toISOString()
      });
    } else {
      // Insert new — FIX: use append_ (insertRow_ does not exist)
      optionId = 'DD_' + Date.now();
      Repository.append_(sheet, {
        option_id: optionId,
        dropdown_key: String(record.dropdownKey || record.dropdown_key || ''),
        option_value: String(record.optionValue || record.option_value || ''),
        sort_order: String(record.sortOrder || record.sort_order || '100'),
        status: String(record.status || 'ACTIVE'),
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      });
    }
    return respond_(true, { record: Object.assign({}, record, { optionId: optionId }) });
  } catch (error) {
    console.error('saveDropdownOptionRecord failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function listDropdownOptionRecords(payload) {
  try {
    var token = payload && (payload.sessionToken || payload);
    var session = getSessionOrThrow_(token);
    var sheet = APP_CONFIG.sheets.DROPDOWN_OPTIONS || 'DropdownOptions';
    var rows = [];
    try { rows = Repository.getAll_(sheet); } catch(e) { rows = []; }
    return respond_(true, { dropdownOptions: rows, ok: true });
  } catch (error) {
    console.error('listDropdownOptionRecords failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

function updateDropdownOptionStatus(payload) {
  try {
    var token = payload && payload.sessionToken;
    var session = getSessionOrThrow_(token);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
    var optionId = String(payload.optionId || '').trim();
    var status = String(payload.status || 'ACTIVE').trim();
    if (!optionId) return respond_(false, null, 'Option ID is required.');
    var sheet = APP_CONFIG.sheets.DROPDOWN_OPTIONS || 'DropdownOptions';
    Repository.updateById_(sheet, 'option_id', optionId, { status: status, updated_at: new Date().toISOString() });
    return respond_(true, { ok: true });
  } catch (error) {
    console.error('updateDropdownOptionStatus failed', error);
    return respond_(false, null, sanitizeClientError_(error));
  }
}

// initializePortalSheetsForCurrentUser is defined once in Seed.gs.
// Removed duplicate simplified global copy from Services.gs to keep the full setup logic active.

function forcePasswordResetAllUsers() {
      const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
      const sheet = ss.getSheetByName('Users');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idx = headers.indexOf('must_reset_password');
      if (idx === -1) { Logger.log('Column not found'); return; }
      for (let i = 1; i < data.length; i++) {
        sheet.getRange(i + 1, idx + 1).setValue('TRUE');
      }
      Logger.log('Done — ' + (data.length - 1) + ' users marked for password reset.');
    }

// ─── Welcome Email ───────────────────────────────────────────────────────────

// ─── Welcome Email ───────────────────────────────────────────────────────────
function sendWelcomeEmail(payload) {
  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);

    var toEmail      = String(payload.toEmail        || '').trim();
    var webappUrl    = String(payload.webappUrl      || '').trim();
    var userId       = String(payload.userId         || '').trim();
    if (!toEmail)   return respond_(false, null, 'Email address is required.');
    if (!webappUrl) return respond_(false, null, 'Webapp URL is required.');
    if (!userId)    return respond_(false, null, 'User ID is required.');

    var ss       = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
    var sheet    = ss.getSheetByName('Users');
    var data     = sheet.getDataRange().getValues();
    var headers  = data[0].map(function(h){ return String(h||'').trim(); });
    var idIdx    = headers.indexOf('user_id');
    var nameIdx  = headers.indexOf('full_name');
    var unameIdx = headers.indexOf('username');
    var targetUser = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx] || '') === userId) {
        targetUser = {
          fullName:      String(data[i][nameIdx]  || ''),
          username:      String(data[i][unameIdx] || ''),
        };
        break;
      }
    }
    if (!targetUser) return respond_(false, null, 'User not found.');

    var logoUrl   = 'https://drive.google.com/thumbnail?id=1KGIgikEbb7NWKt1vXhT4V10yXLw-ukmA&sz=w200-h200';
    var firstName = (targetUser.fullName || '').split(' ')[0] || targetUser.username;

    // Gmail-safe table-based layout (no flex/grid/CSS classes)
    var htmlBody =
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
      '<body style="margin:0;padding:0;background:#eef2f7;font-family:Arial,sans-serif;">' +
      '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#eef2f7;padding:32px 0;">' +
        '<tr><td align="center">' +
          '<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">' +

            // Header
            '<tr><td style="background:linear-gradient(135deg,#002b5c 0%,#0076b6 100%);padding:32px 40px 24px;text-align:center;">' +
              '<img src="' + logoUrl + '" width="68" height="68" alt="Saisethsons" style="border-radius:50%;border:3px solid rgba(255,255,255,0.30);display:block;margin:0 auto 14px auto;">' +
              '<div style="color:#ffffff;font-size:22px;font-weight:700;letter-spacing:0.5px;margin-bottom:6px;">Saisethsons Hospitality Group</div>' +
              '<div style="color:rgba(255,255,255,0.75);font-size:12px;letter-spacing:2px;text-transform:uppercase;">Valuing&nbsp;People &nbsp;&middot;&nbsp; Product &nbsp;&middot;&nbsp; Performance</div>' +
            '</td></tr>' +

            // Body
            '<tr><td style="padding:36px 40px 28px;">' +
              '<p style="font-size:20px;font-weight:700;color:#002b5c;margin:0 0 12px 0;">Welcome, ' + firstName + '! &#128075;</p>' +
              '<p style="font-size:15px;color:#4a5568;line-height:1.7;margin:0 0 24px 0;">Your Sethpitality portal account has been set up and is ready to use. Use your username below to sign in.</p>' +

              // Credentials table
              '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f4f8ff;border:1px solid #cde0f5;border-radius:10px;margin-bottom:26px;">' +

                '<tr>' +
                  '<td style="padding:13px 20px;border-bottom:1px solid #daeaf7;vertical-align:middle;">' +
                    '<span style="font-size:12px;color:#6b7a94;font-weight:600;text-transform:uppercase;letter-spacing:0.8px;">Username</span>' +
                  '</td>' +
                  '<td style="padding:13px 20px;vertical-align:middle;">' +
                    '<span style="font-size:14px;color:#002b5c;font-weight:700;">' + targetUser.username + '</span>' +
                  '</td>' +
                '</tr>' +

              '</table>' +

              // CTA button
              '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">' +
                '<tr><td align="center" style="padding:8px 0;">' +
                  '<a href="' + webappUrl + '" style="display:inline-block;background:#0076b6;color:#ffffff;text-decoration:none;padding:14px 42px;border-radius:8px;font-size:16px;font-weight:700;font-family:Arial,sans-serif;mso-padding-alt:0;border:1px solid #0076b6;">Open Sethpitality Portal &rarr;</a>' +
                '</td></tr>' +
                '<tr><td align="center" style="padding:6px 0 0 0;">' +
                  '<span style="font-size:11px;color:#9aa3b0;">If the button does not open, copy this link: </span>' +
                  '<a href="' + webappUrl + '" style="font-size:11px;color:#0076b6;word-break:break-all;">' + webappUrl + '</a>' +
                '</td></tr>' +
              '</table>' +

              // Security note
              '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
                '<tr><td style="background:#fffbeb;border:1px solid #f7d87c;border-radius:8px;padding:13px 16px;">' +
                  '<span style="font-size:13px;color:#7a5f00;line-height:1.6;">&#128274; For your security, please change your password immediately after your first login. If you have trouble signing in, contact your system administrator.</span>' +
                '</td></tr>' +
              '</table>' +
            '</td></tr>' +

            // Footer
            '<tr><td style="background:#f4f8ff;border-top:1px solid #daeaf7;padding:18px 40px;text-align:center;">' +
              '<p style="font-size:12px;color:#9aa3b0;margin:0;line-height:1.6;"><strong style="color:#0076b6;">Saisethsons Hospitality Group</strong><br>This is an automated message &mdash; please do not reply directly to this email.</p>' +
            '</td></tr>' +

          '</table>' +
        '</td></tr>' +
      '</table>' +
      '</body></html>';

    MailApp.sendEmail({
      to:       toEmail,
      subject:  'Welcome to Saisethsons Hospitality Group HOST Portal \u2014 Your Account is Ready, ' + firstName + '!',
      htmlBody: htmlBody,
      body:     'Welcome ' + firstName + '! Username: ' + targetUser.username + ' | Portal: ' + webappUrl
    });

    return respond_(true, { message: 'Welcome email sent to ' + toEmail });
  } catch (e) {
    console.error('sendWelcomeEmail failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 5W TRACKER
// ═══════════════════════════════════════════════════════════════════════════
function get5WTrackerData(payload) {
  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    var role    = String(session.roleCode || '');
    var isAdmin = role === APP_CONFIG.roles.SUPER_ADMIN;
    var isMRO   = role === APP_CONFIG.roles.MRO;

    // Only Restaurant Users and Super Admin (MRO excluded per spec)
    if (!isAdmin && role !== APP_CONFIG.roles.RESTAURANT_USER) {
      return respond_(false, null, 'Access denied.');
    }

    var FIVE_W_SHEET_ID = (function() {
      try {
        var fromProps = PropertiesService.getScriptProperties().getProperty('FIVE_W_SHEET_ID');
        if (fromProps) return fromProps;
      } catch (e) {}
      return '1GQwlH8oDaufRv0X7FYV0HOy986ZGs9CNUiDXY-QBAKM'; // fallback default
    })();
    var RESPONSES_TAB     = 'Sethpitality 5W Evaluation';
    var PASS_OVERALL      = 18;
    var PASS_SIG          = 10;

    var ss    = SpreadsheetApp.openById(FIVE_W_SHEET_ID);
    var sheet = ss.getSheetByName(RESPONSES_TAB);
    if (!sheet || sheet.getLastRow() < 2) {
      return respond_(true, { rows: [], restaurants: [] });
    }

    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 36);
    var data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Determine which restaurant numbers this session can see
    var allowedRestaurants = null; // null = all (super admin)
    if (!isAdmin) {
      // Restaurant User: get their assigned restaurant numbers
      var accessRows = Repository.getAll_(APP_CONFIG.sheets.RESTAURANT_ACCESS)
        .filter(function(r) { return String(r.user_id || '') === String(session.userId || ''); });
      var restaurantIds = accessRows.map(function(r) { return String(r.restaurant_id || ''); });
      // Map restaurant IDs to restaurant numbers
      var restaurants = Repository.getAll_(APP_CONFIG.sheets.RESTAURANTS);
      allowedRestaurants = [];
      restaurantIds.forEach(function(id) {
        var rest = restaurants.find(function(r) { return String(r.restaurant_id || '') === id; });
        if (rest) {
          var num = String(rest.restaurant_number || rest.restaurant_code || '').replace(/\.0$/, '').trim();
          if (num) allowedRestaurants.push(num);
        }
      });
      // Fallback: also try primary restaurant from session
      if (!allowedRestaurants.length && session.restaurantId) {
        var primary = restaurants.find(function(r) { return String(r.restaurant_id || '') === String(session.restaurantId); });
        if (primary) {
          var pNum = String(primary.restaurant_number || primary.restaurant_code || '').replace(/\.0$/, '').trim();
          if (pNum) allowedRestaurants.push(pNum);
        }
      }
    }

    function isYes(v) {
      var s = String(v || '').toLowerCase();
      return s.indexOf('yes') > -1 || s.indexOf('✅') > -1;
    }

    function parseScore(str, max) {
      if (!str) return 0;
      var s = String(str).trim();
      var m = s.match(/(\d+)\s*\/\s*(\d+)/);
      if (m) return Math.min(parseInt(m[1], 10), max);
      var n = parseInt(s, 10);
      return isNaN(n) ? 0 : Math.min(n, max);
    }

    var rows = [];
    var restaurantSet = {};

    data.forEach(function(row) {
      var rawTs = row[0];
      if (!rawTs) return;
      var ts = rawTs instanceof Date ? rawTs : new Date(rawTs);
      if (isNaN(ts.getTime())) return;

      var rn = String(row[4] || '').replace(/\.0$/, '').trim();
      if (!rn) return;

      // Filter by allowed restaurants for non-admin
      if (allowedRestaurants && allowedRestaurants.indexOf(rn) === -1) return;

      var sig = parseScore(row[34], 10);
      var ov  = parseScore(row[35], 25);

      var w1 = Math.min((isYes(row[7])?2:0)+(isYes(row[8])?1:0)+(isYes(row[9])?1:0)+(isYes(row[10])?1:0), 5);
      var w2 = Math.min((isYes(row[12])?2:0)+(isYes(row[13])?1:0)+(isYes(row[14])?1:0)+(isYes(row[15])?1:0), 5);
      var w3 = Math.min((isYes(row[17])?2:0)+(isYes(row[18])?1:0)+(isYes(row[19])?1:0)+(isYes(row[20])?1:0), 5);
      var w4 = Math.min((isYes(row[22])?2:0)+(isYes(row[23])?1:0)+(isYes(row[24])?1:0)+(isYes(row[25])?1:0), 5);
      var w5 = Math.min((isYes(row[27])?2:0)+(isYes(row[28])?1:0)+(isYes(row[29])?1:0)+(isYes(row[30])?1:0), 5);

      restaurantSet[rn] = true;
      rows.push({
        ts:  ts.getTime(),
        dt:  Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MMM d, yyyy'),
        ev:  String(row[2] || '').trim(),
        mb:  String(row[3] || '').trim(),
        rn:  rn,
        sh:  String(row[5] || '').trim(),
        sg:  sig,
        ov:  ov,
        w1:  w1, w2: w2, w3: w3, w4: w4, w5: w5,
        ok:  ov  >= PASS_OVERALL,
        pf:  sig >= PASS_SIG
      });
    });

    rows.sort(function(a, b) { return b.ts - a.ts; });

    return respond_(true, {
      rows:        rows,
      restaurants: Object.keys(restaurantSet).sort(),
      isAdmin:     isAdmin,
      passOverall: PASS_OVERALL,
      passSig:     PASS_SIG
    });

  } catch (e) {
    console.error('get5WTrackerData failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// HOT ZONE NEWSPAPER — Backend v2
// Separate spreadsheet via HOTZONE_SPREADSHEET_ID Script Property.
// Falls back to main spreadsheet if not set.
// Sections: WOW_OF_WEEK, HOST_GOALS, OPERATIONS, TRAINING, SPOTLIGHT,
//           PROMOTIONS, ACTIVE_MENU, DISCONTINUED_MENU, ANNOUNCEMENTS, EVENTS
// ═══════════════════════════════════════════════════════════════════════════

var HZ_SECTIONS_ = [
  { key: 'WOW_OF_WEEK',       label: 'Wow of the Week' },
  { key: 'HOST_GOALS',        label: 'HOST Goals & Operations' },
  { key: 'OPERATIONS',        label: 'Operations Update' },
  { key: 'TRAINING',          label: 'Training & Development' },
  { key: 'SPOTLIGHT',         label: 'Staff Spotlight' },
  { key: 'PROMOTIONS',        label: 'Upcoming Promotions' },
  { key: 'ACTIVE_MENU',       label: 'Active Menu Items' },
  { key: 'DISCONTINUED_MENU', label: 'Discontinued Items' },
  { key: 'ANNOUNCEMENTS',     label: 'Announcements & Reminders' },
  { key: 'EVENTS',            label: 'Upcoming Events' }
];

// ── Sheet helpers ─────────────────────────────────────────────────────────
function hzSS_() {
  try {
    var id = PropertiesService.getScriptProperties().getProperty('HOTZONE_SPREADSHEET_ID');
    if (id) return SpreadsheetApp.openById(id);
  } catch(e) {}
  return SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
}

function hzSheet_(name) {
  var ss = hzSS_();
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  var headers = {
    'HotZone Editions':      ['edition_id','title','week_ending','status','webapp_url','qr_url','created_by','created_by_id','created_at','updated_at'],
    'HotZone Contributions': ['contribution_id','edition_id','section','title','body','link_url','link_title','image_url','image_file_id','submitted_by','submitted_by_id','submitted_at','status'],
    'HotZone Templates':     ['template_id','section','title','body','link_url','link_title','created_by','created_by_id','created_at'],
    'CalendarEvents':        ['event_id','event_date','title','message','link_url','link_title','created_by','created_by_id','created_at','updated_at','status'],
    'SavedCategories':       ['category_id','section','title','body','link_url','link_title','image_url','image_file_id','created_by','created_by_id','created_at','status']
  };
  if (headers[name]) sh.getRange(1,1,1,headers[name].length).setValues([headers[name]]);
  sh.setFrozenRows(1);
  return sh;
}

function hzAll_(name) {
  var sh = hzSheet_(name); var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = data[0].map(function(v){return String(v||'').trim();});
  return data.slice(1).map(function(row){ var o={}; h.forEach(function(k,i){o[k]=row[i]==null?'':row[i];}); return o; });
}

function hzAppend_(name, record) {
  var sh = hzSheet_(name);
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(v){return String(v||'').trim();});
  sh.appendRow(h.map(function(k){return record[k]!==undefined?record[k]:'';}));
}

function hzUpdate_(name, idCol, idVal, patch) {
  var sh = hzSheet_(name); var data = sh.getDataRange().getValues();
  var h = data[0].map(function(v){return String(v||'').trim();});
  var ci = h.indexOf(idCol); if (ci===-1) return;
  for (var r=1;r<data.length;r++) {
    if (String(data[r][ci])===String(idVal)) {
      Object.keys(patch).forEach(function(k){ var ki=h.indexOf(k); if(ki!==-1) sh.getRange(r+1,ki+1).setValue(patch[k]); });
      return;
    }
  }
}

function hzFind_(name, col, val) {
  return hzAll_(name).find(function(r){return String(r[col])===String(val);})||null;
}

function isHzManager_(session) {
  return [APP_CONFIG.roles.SUPER_ADMIN, APP_CONFIG.roles.HR_ADMIN, APP_CONFIG.roles.SENIOR_MANAGEMENT].indexOf(session.roleCode) !== -1;
}

function hzBuildHtml_(edition, grouped) {
  function esc(v){return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
  var cols=['','',''];
  HZ_SECTIONS_.forEach(function(sec,i){
    var items=(grouped[sec.key]||[]).filter(function(c){return String(c.status)!=='DELETED';});
    if(!items.length) return;
    var html='<div style="margin-bottom:14px"><div style="background:#1a1a1a;color:#fff;font-size:8px;letter-spacing:.1em;text-transform:uppercase;padding:2px 6px;font-family:Arial,sans-serif;font-weight:700;margin-bottom:7px">'+esc(sec.label)+'</div>';
    items.forEach(function(c){
      html+='<div style="margin-bottom:11px;padding-bottom:11px;border-bottom:1px dotted #ddd">'
        +(c.title?'<div style="font-family:Playfair Display,Georgia,serif;font-size:13px;font-weight:700;margin:0 0 4px;color:#1a1a1a">'+esc(c.title)+'</div>':'')
        +(c.image_url?'<img src="'+esc(c.image_url)+'" style="width:100%;max-height:90px;object-fit:cover;border-radius:3px;margin:5px 0">':'')
        +(c.body?'<div style="font-size:10px;line-height:1.6;color:#333">'+esc(c.body).replace(/\n/g,'<br>')+'</div>':'')
        +'<div style="font-size:8px;color:#999;text-transform:uppercase;letter-spacing:.06em;margin-top:3px;font-family:Arial,sans-serif">— '+esc(c.submitted_by)+'</div>'
        +'</div>';
    });
    html+='</div>';
    cols[i%3]+=html;
  });
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    +'<link href="https://fonts.googleapis.com/css2?family=UnifrakturMaguntia&family=Playfair+Display:wght@700;900&display=swap" rel="stylesheet">'
    +'<style>body{margin:0;padding:20px;background:#fff;font-family:Georgia,serif;color:#1a1a1a}@media print{body{padding:0}}</style>'
    +'</head><body><div style="max-width:760px;margin:0 auto">'
    +'<div style="text-align:center;border-bottom:4px double #1a1a1a;padding-bottom:10px;margin-bottom:0">'
    +'<p style="font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:#666;margin:0 0 4px;font-family:Arial,sans-serif">Serve People · Grow Opportunities · Build Champions</p>'
    +'<div style="font-family:UnifrakturMaguntia,cursive;font-size:52px;line-height:1.1;color:#1a1a1a">The Saisethsons Hot Zone</div>'
    +'<div style="display:flex;justify-content:space-between;border-top:1px solid #1a1a1a;border-bottom:1px solid #1a1a1a;margin:6px 0;padding:3px 0;font-size:9px;letter-spacing:.08em;text-transform:uppercase;font-family:Arial,sans-serif;color:#555">'
    +'<span>'+esc(edition.title||'')+'</span>'
    +'<span style="background:#1a1a1a;color:#fff;padding:2px 10px;font-weight:700">Weekly Edition</span>'
    +'<span>Week ending '+esc(edition.week_ending||'')+'</span>'
    +'</div></div>'
    +'<div style="display:grid;grid-template-columns:1fr 1px 1fr 1px 1fr;gap:0;padding-top:12px">'
    +'<div style="padding:0 12px 0 0">'+(cols[0]||'<p style="color:#ccc;font-style:italic;font-size:11px">No content</p>')+'</div>'
    +'<div style="background:#d0cfc8"></div>'
    +'<div style="padding:0 12px">'+(cols[1]||'')+'</div>'
    +'<div style="background:#d0cfc8"></div>'
    +'<div style="padding:0 0 0 12px">'+(cols[2]||'')+'</div>'
    +'</div>'
    +'<div style="border-top:4px double #1a1a1a;margin-top:14px;padding-top:7px;display:flex;justify-content:space-between;font-size:8px;letter-spacing:.07em;text-transform:uppercase;color:#888;font-family:Arial,sans-serif">'
    +'<span>Saisethsons Hospitality Group · Internal Use Only</span>'
    +'<span>Generated via Sethpitality Portal</span>'
    +'<span>Week ending '+esc(edition.week_ending||'')+'</span>'
    +'</div></div></body></html>';
}

// ── 1. Get all editions list (60s cache per user role) ────────────────────
function getHotzoneEditions(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const cacheKey = 'hz_editions_' + String(session.userId || '').slice(-12);
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      return respond_(true, Object.assign(parsed, { canManage: isHzManager_(session) }));
    }
  } catch(e) {}
  const editions = hzAll_('HotZone Editions').reverse().map(function(ed) {
    var contribs = hzAll_('HotZone Contributions').filter(function(c){
      return String(c.edition_id)===String(ed.edition_id) && String(c.status)!=='DELETED';
    });
    return Object.assign({}, ed, { contrib_count: contribs.length });
  });
  const templates = hzAll_('HotZone Templates');
  const data = { editions: editions, templates: templates };
  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(data), 60); } catch(e) {}
  return respond_(true, Object.assign({}, data, { canManage: isHzManager_(session) }));
}

// ── 2. Get single edition detail with contributions ────────────────────────
function getHotzoneEditionDetail(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const edition = hzFind_('HotZone Editions', 'edition_id', payload.edition_id);
  if (!edition) return respond_(false, null, 'Edition not found.');
  const contributions = hzAll_('HotZone Contributions').filter(function(c){
    return String(c.edition_id)===String(payload.edition_id);
  });
  const grouped = {};
  HZ_SECTIONS_.forEach(function(s){grouped[s.key]=[];});
  contributions.forEach(function(c){
    var k = String(c.section||'').toUpperCase();
    if (!grouped[k]) grouped[k]=[];
    grouped[k].push(c);
  });
  const templates = hzAll_('HotZone Templates');
  return respond_(true, { edition: edition, contributions: grouped, templates: templates, canManage: isHzManager_(session) });
}

// ── 3. Create new edition ─────────────────────────────────────────────────
function createHotzoneEdition(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Only Senior Management, HR Admin, and Super Admin can create editions.');
  const title = String(payload.title||'').trim();
  const weekEnd = String(payload.week_ending||'').trim();
  if (!title) throw new Error('Title is required.');
  if (!weekEnd) throw new Error('Week ending date is required.');
  const now = new Date().toISOString();
  const edition = {
    edition_id: Utilities.getUuid(),
    title: title,
    week_ending: weekEnd,
    status: 'DRAFT',
    created_by: session.fullName || session.username || '',
    created_by_id: session.userId || '',
    created_at: now,
    updated_at: now
  };
  hzAppend_('HotZone Editions', edition);
  // Clear editions cache so list refreshes
  try { CacheService.getScriptCache().remove('hz_editions_' + String(session.userId || '').slice(-12)); } catch(e) {}
  return respond_(true, { edition_id: edition.edition_id });
}

// ── 4. Save / update a contribution ───────────────────────────────────────
function saveHotzoneContribution(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const edition = hzFind_('HotZone Editions', 'edition_id', payload.edition_id);
  if (!edition) throw new Error('Edition not found.');
  const status = String(edition.status||'').toUpperCase();
  if (status === 'PUBLISHED' && !isHzManager_(session)) throw new Error('This edition is published. Only managers can edit it.');
  const section = String(payload.section||'').toUpperCase().trim();
  if (!section) throw new Error('Section is required.');
  const title = String(payload.title||'').trim();
  const body  = String(payload.body||'').trim();
  if (!title && !body) throw new Error('Title or body is required.');
  const now = new Date().toISOString();
  const existingId = String(payload.contribution_id||'').trim();

  if (existingId) {
    const existing = hzFind_('HotZone Contributions', 'contribution_id', existingId);
    if (existing && String(existing.submitted_by_id) !== String(session.userId) && !isHzManager_(session)) {
      throw new Error('You can only edit your own contributions.');
    }
    hzUpdate_('HotZone Contributions', 'contribution_id', existingId, {
      title: title, body: body,
      link_url: String(payload.link_url||'').trim(),
      link_title: String(payload.link_title||'').trim(),
      image_url: String(payload.image_url||''),
      image_file_id: String(payload.image_file_id||''), status: 'ACTIVE'
    });
  } else {
    hzAppend_('HotZone Contributions', {
      contribution_id: Utilities.getUuid(),
      edition_id: payload.edition_id,
      section: section,
      title: title, body: body,
      link_url: String(payload.link_url||'').trim(),
      link_title: String(payload.link_title||'').trim(),
      image_url: String(payload.image_url||''),
      image_file_id: String(payload.image_file_id||''),
      submitted_by: session.fullName || session.username || '',
      submitted_by_id: session.userId || '',
      submitted_at: now,
      status: 'ACTIVE'
    });
  }

  // Save as template if requested
  if (String(payload.save_as_template||'') === '1' && (title||body)) {
    hzAppend_('HotZone Templates', {
      template_id: Utilities.getUuid(),
      section: section,
      title: title, body: body,
      link_url: String(payload.link_url||'').trim(),
      link_title: String(payload.link_title||'').trim(),
      created_by: session.fullName || session.username || '',
      created_by_id: session.userId || '',
      created_at: now
    });
  }
  return respond_(true, {});
}

// ── 5. Delete a contribution ───────────────────────────────────────────────
function deleteHotzoneContribution(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const cid = String(payload.contribution_id||'').trim();
  const existing = hzFind_('HotZone Contributions', 'contribution_id', cid);
  if (existing && String(existing.submitted_by_id) !== String(session.userId) && !isHzManager_(session)) {
    throw new Error('You can only delete your own contributions.');
  }
  hzUpdate_('HotZone Contributions', 'contribution_id', cid, { status: 'DELETED' });
  return respond_(true, {});
}

// ── 6. Upload image ────────────────────────────────────────────────────────
function uploadHotzoneImage(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const base64 = String(payload.base64||'').trim();
  if (!base64) throw new Error('No image data.');
  const folderName = 'HotZone Images';
  var folder;
  var folders = DriveApp.getFoldersByName(folderName);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const bytes = Utilities.base64Decode(base64);
  const blob  = Utilities.newBlob(bytes, String(payload.mimeType||'image/jpeg'), String(payload.fileName||'hotzone.jpg'));
  const file  = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return respond_(true, { file_id: file.getId(), image_url: 'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w600' });
}

// ── 7. Update edition status ───────────────────────────────────────────────
function updateHotzoneEditionStatus(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const eid = String(payload.edition_id||'').trim();
  const status = String(payload.status||'').toUpperCase().trim();
  if (!['DRAFT','SAVED','PUBLISHED'].includes(status)) throw new Error('Invalid status.');
  var patch = { status: status, updated_at: new Date().toISOString() };
  // Store webapp_url on publish so QR code can encode a direct link
  if (status === 'PUBLISHED' && payload.webapp_url) {
    var url = decodeURIComponent(String(payload.webapp_url));
    patch.webapp_url = url;
    patch.qr_url = url; // same URL — decoded for QR encoding
  }
  hzUpdate_('HotZone Editions', 'edition_id', eid, patch);
  return respond_(true, {});
}

// ── 8. Email edition to all users ─────────────────────────────────────────
function mailHotzoneEdition(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const edition = hzFind_('HotZone Editions', 'edition_id', payload.edition_id);
  if (!edition) throw new Error('Edition not found.');
  const contribs = hzAll_('HotZone Contributions').filter(function(c){
    return String(c.edition_id)===String(payload.edition_id) && String(c.status)!=='DELETED';
  });
  const grouped={};
  HZ_SECTIONS_.forEach(function(s){grouped[s.key]=[];});
  contribs.forEach(function(c){var k=String(c.section||'').toUpperCase();if(grouped[k])grouped[k].push(c);});
  const html = hzBuildHtml_(edition, grouped);
  const subject = '\uD83D\uDCF0 ' + (edition.title||'The Saisethsons Hot Zone') + ' \u2014 Week ending ' + (edition.week_ending||'');
  const users = Repository.getAll_(APP_CONFIG.sheets.USERS).filter(function(u){
    return String(u.status||'').toUpperCase()==='ACTIVE' && u.email;
  });
  var sent = 0;
  users.forEach(function(u){
    try { MailApp.sendEmail({ to: u.email, subject: subject, htmlBody: html, body: 'The Saisethsons Hot Zone is ready. Week ending '+edition.week_ending }); sent++; }
    catch(e){ console.warn('Email failed: '+u.email, e); }
  });
  return respond_(true, { sent: sent });
}

// ── 9. Get pre-fill data for newsletter upload (user completes form in NM page)
function publishHotzoneToNewsletter(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const edition = hzFind_('HotZone Editions', 'edition_id', payload.edition_id);
  if (!edition) throw new Error('Edition not found.');
  const contribs = hzAll_('HotZone Contributions').filter(function(c){
    return String(c.edition_id)===String(payload.edition_id) && String(c.status)!=='DELETED';
  });
  const sectionSummary = HZ_SECTIONS_.map(function(s){
    var items = contribs.filter(function(c){return String(c.section||'').toUpperCase()===s.key;});
    return items.length ? s.label+': '+items.length+' update'+(items.length>1?'s':'') : null;
  }).filter(Boolean).join(' · ');
  const today = new Date().toISOString().slice(0,10);
  const nextWeek = new Date(Date.now()+7*86400000).toISOString().slice(0,10);
  // Return prefill data — frontend opens NM page with fields populated
  return respond_(true, {
    prefill: {
      title: edition.title || 'Hot Zone — Week ending '+edition.week_ending,
      message: sectionSummary || 'Weekly Hot Zone edition.',
      priority: 'Normal',
      start: today,
      end: nextWeek,
      audience: 'All Users'
    }
  });
}

// ── Calendar Events ────────────────────────────────────────────────────────

function getHotzoneCalendarEvents(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const events = hzAll_('CalendarEvents').filter(function(e) {
    return String(e.status || '') !== 'DELETED';
  }).sort(function(a,b) { return String(a.event_date).localeCompare(String(b.event_date)); });
  return respond_(true, { events: events });
}

function saveHotzoneCalendarEvent(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const eventDate = String(payload.event_date || '').trim();
  const title     = String(payload.title      || '').trim();
  if (!eventDate) throw new Error('Date is required.');
  if (!title)     throw new Error('Title is required.');
  const now = new Date().toISOString();
  const existingId = String(payload.event_id || '').trim();
  // Auto-create CalendarEvents sheet with headers if it doesn't exist
  hzSheet_('CalendarEvents'); // ensures sheet exists
  if (existingId) {
    hzUpdate_('CalendarEvents', 'event_id', existingId, {
      event_date: eventDate, title: title,
      message: String(payload.message || '').trim(),
      link_url: String(payload.link_url || '').trim(),
      link_title: String(payload.link_title || '').trim(),
      updated_at: now
    });
  } else {
    hzAppend_('CalendarEvents', {
      event_id: Utilities.getUuid(),
      event_date: eventDate, title: title,
      message: String(payload.message || '').trim(),
      link_url: String(payload.link_url || '').trim(),
      link_title: String(payload.link_title || '').trim(),
      created_by: session.fullName || session.username || '',
      created_by_id: session.userId || '',
      created_at: now, updated_at: now, status: 'ACTIVE'
    });
  }
  return respond_(true, {});
}

function deleteHotzoneCalendarEvent(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  hzUpdate_('CalendarEvents', 'event_id', String(payload.event_id || ''), { status: 'DELETED' });
  return respond_(true, {});
}

// ── Get latest published edition (for Host Hot Zone viewer button) ─────────
function getHotzoneLatestPublished(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  const editions = hzAll_('HotZone Editions')
    .filter(function(ed) { return String(ed.status || '') === 'PUBLISHED'; })
    .sort(function(a,b) { return String(b.week_ending||'').localeCompare(String(a.week_ending||'')); });
  if (!editions.length) return respond_(false, null, 'No published editions found.');
  const edition = editions[0];
  const contributions = hzAll_('HotZone Contributions').filter(function(c) {
    return String(c.edition_id) === String(edition.edition_id);
  });
  const grouped = {};
  HZ_SECTIONS_.forEach(function(s) { grouped[s.key] = []; });
  contributions.forEach(function(c) {
    var k = String(c.section || '').toUpperCase();
    if (grouped[k]) grouped[k].push(c);
  });
  return respond_(true, { edition: edition, contributions: grouped });
}

// Update CalendarEvents sheet header creation in hzSheet_
// (Add to the headers map inside hzSheet_)
// NOTE: hzSheet_ already handles auto-creation. The CalendarEvents headers
// are added via the hzAppend_ call which reads existing headers on first write.
// If the sheet is empty, run saveHotzoneCalendarEvent once to initialise it,
// or manually create with headers:
// event_id | event_date | title | message | link_url | link_title | created_by | created_by_id | created_at | updated_at | status

// ── Email targeting functions ─────────────────────────────────────────────

function getHotzoneEmailRecipientData(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const users = Repository.getAll_(APP_CONFIG.sheets.USERS).filter(function(u){
    return String(u.status||'').toUpperCase()==='ACTIVE';
  });
  const restSet = {};
  users.forEach(function(u){
    var r = String(u.restaurant_number||u.restaurantNumber||'').trim();
    if(r) restSet[r] = true;
  });
  try {
    Repository.getAll_(APP_CONFIG.sheets.RESTAURANTS).forEach(function(r){
      var rn = String(r.restaurant_number||r.restaurantNumber||'').trim();
      if(rn) restSet[rn] = true;
    });
  } catch(e) {}
  var brands = [];
  try {
    brands = Repository.getAll_(APP_CONFIG.sheets.BRANDS).map(function(b){
      return String(b.brand_name||b.name||b.brand||'').trim();
    }).filter(Boolean);
  } catch(e) {}
  return respond_(true, { restaurants: Object.keys(restSet).sort(), brands: brands });
}

function previewHotzoneEmailRecipients(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  return respond_(true, { count: resolveHotzoneRecipients_(payload).length });
}

function sendHotzoneEmail(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const edition = hzFind_('HotZone Editions', 'edition_id', payload.edition_id);
  if (!edition) throw new Error('Edition not found.');
  const contribs = hzAll_('HotZone Contributions').filter(function(c){
    return String(c.edition_id)===String(payload.edition_id) && String(c.status)!=='DELETED';
  });
  const grouped={};
  HZ_SECTIONS_.forEach(function(s){grouped[s.key]=[];});
  contribs.forEach(function(c){var k=String(c.section||'').toUpperCase();if(grouped[k])grouped[k].push(c);});
  const html = hzBuildHtml_(edition, grouped);
  const subject = 'Hot Zone: ' + (edition.title||'The Saisethsons Hot Zone') + ' - Week ending ' + (edition.week_ending||'');
  const recipients = resolveHotzoneRecipients_(payload);
  var sent = 0;
  recipients.forEach(function(email){
    try { MailApp.sendEmail({ to: email, subject: subject, htmlBody: html, body: 'The Saisethsons Hot Zone is ready.' }); sent++; }
    catch(e){ console.warn('Email failed: '+email, e); }
  });
  return respond_(true, { sent: sent });
}

function resolveHotzoneRecipients_(payload) {
  const mode = String(payload.mode||'all').toLowerCase();
  const allUsers = Repository.getAll_(APP_CONFIG.sheets.USERS).filter(function(u){
    return String(u.status||'').toUpperCase()==='ACTIVE' && u.email;
  });
  var filtered;
  if (mode === 'role') {
    var roles = Array.isArray(payload.roles) ? payload.roles : [];
    filtered = allUsers.filter(function(u){ return roles.indexOf(String(u.role_code||u.roleCode||'')) !== -1; });
  } else if (mode === 'restaurant') {
    var rests = Array.isArray(payload.restaurants) ? payload.restaurants.map(String) : [];
    filtered = allUsers.filter(function(u){ return rests.indexOf(String(u.restaurant_number||u.restaurantNumber||'')) !== -1; });
  } else if (mode === 'brand') {
    var brands = Array.isArray(payload.brands) ? payload.brands.map(String) : [];
    filtered = allUsers.filter(function(u){ return brands.indexOf(String(u.brand||u.brand_name||'')) !== -1; });
  } else {
    filtered = allUsers;
  }
  var seen = {}, result = [];
  filtered.forEach(function(u){
    var e = String(u.email||'').toLowerCase().trim();
    if(e && !seen[e]){ seen[e]=true; result.push(u.email); }
  });
  return result;
}

// ── Row deletion helper ────────────────────────────────────────────────────
// Deletes ALL rows in a sheet where idCol === idVal (handles multiple matches).
function hzDeleteRows_(sheetName, idCol, idVal) {
  var sh = hzSheet_(sheetName);
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(v){return String(v||'').trim();});
  var ci = headers.indexOf(idCol);
  if (ci === -1) return;
  // Walk backwards so row deletions don't shift indices
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][ci]) === String(idVal)) {
      sh.deleteRow(r + 1);
    }
  }
}

// ── Delete Edition (+ cascade-delete contributions) ────────────────────────
function deleteHotzoneEdition(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const eid = String(payload.edition_id || '').trim();
  if (!eid) throw new Error('edition_id required.');
  // Cascade-delete all contributions for this edition
  hzDeleteRows_('HotZone Contributions', 'edition_id', eid);
  // Delete the edition row itself
  hzDeleteRows_('HotZone Editions', 'edition_id', eid);
  // Clear editions cache
  try { CacheService.getScriptCache().remove('hz_editions_' + String(session.userId || '').slice(-12)); } catch(e) {}
  return respond_(true, {});
}

// ── Saved Categories CRUD ─────────────────────────────────────────────────

function getHotzoneCategories(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const categories = hzAll_('SavedCategories').filter(function(c) {
    return String(c.status || '') !== 'DELETED';
  }).sort(function(a, b) {
    return String(a.section).localeCompare(String(b.section)) || String(a.title).localeCompare(String(b.title));
  });
  return respond_(true, { categories: categories });
}

function saveHotzoneCategory(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const section  = String(payload.section   || '').trim().toUpperCase();
  const title    = String(payload.title     || '').trim();
  const body     = String(payload.body      || '').trim();
  if (!section) throw new Error('Section is required.');
  if (!title && !body) throw new Error('Title or body is required.');
  const now = new Date().toISOString();
  const existingId = String(payload.category_id || '').trim();
  if (existingId) {
    // Update existing
    hzUpdate_('SavedCategories', 'category_id', existingId, {
      section:    section,
      title:      title,
      body:       body,
      link_url:   String(payload.link_url   || '').trim(),
      link_title: String(payload.link_title || '').trim(),
      image_url:  String(payload.image_url  || '').trim(),
      image_file_id: String(payload.image_file_id || '').trim()
    });
    return respond_(true, { category_id: existingId });
  }
  const cId = Utilities.getUuid();
  hzAppend_('SavedCategories', {
    category_id:   cId,
    section:       section,
    title:         title,
    body:          body,
    link_url:      String(payload.link_url   || '').trim(),
    link_title:    String(payload.link_title || '').trim(),
    image_url:     String(payload.image_url  || '').trim(),
    image_file_id: String(payload.image_file_id || '').trim(),
    created_by:    session.fullName || session.username || '',
    created_by_id: session.userId || '',
    created_at:    now,
    status:        'ACTIVE'
  });
  return respond_(true, { category_id: cId });
}

function deleteHotzoneCategory(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const cid = String(payload.category_id || '').trim();
  if (!cid) throw new Error('category_id required.');
  hzUpdate_('SavedCategories', 'category_id', cid, { status: 'DELETED' });
  return respond_(true, {});
}

// ── Batch delete multiple editions + their contributions ──────────────────
function batchDeleteHotzoneEditions(payload) {
  const session = getSessionOrThrow_(payload.sessionToken);
  if (!isHzManager_(session)) throw new Error('Not authorized.');
  const ids = Array.isArray(payload.edition_ids) ? payload.edition_ids : [];
  if (!ids.length) throw new Error('No edition IDs provided.');
  ids.forEach(function(eid) {
    var id = String(eid || '').trim();
    if (!id) return;
    hzDeleteRows_('HotZone Contributions', 'edition_id', id);
    hzDeleteRows_('HotZone Editions', 'edition_id', id);
  });
  try { CacheService.getScriptCache().remove('hz_editions_' + String(session.userId || '').slice(-12)); } catch(e) {}
  return respond_(true, { deleted: ids.length });
}

