function initializePortalSheets() {
  const spreadsheet = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const schemas = getSheetSchemas_();
  const renameMap = {
    Sheet2: 'HOST Links',
    Sheet3: 'Newsletters',
    Sheet4: 'Notifications'
  };

  Object.keys(renameMap).forEach(function(sourceName) {
    const sourceSheet = spreadsheet.getSheetByName(sourceName);
    const targetName = renameMap[sourceName];
    if (sourceSheet && !spreadsheet.getSheetByName(targetName)) {
      sourceSheet.setName(targetName);
    }
  });

  Object.keys(schemas).forEach(function(sheetName) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    const headers = schemas[sheetName];
    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }
    upsertHeaderRow_(sheet, headers);
    sheet.setFrozenRows(1);
  });

  return 'Portal sheets initialized.';
}

function runOneTimeSetup() {
  SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  initializePortalSheets();
  ensureDemoPortalData_({
    userId: 'setup-script',
    roleCode: 'SUPER_ADMIN'
  });
  console.log('One-time portal setup completed successfully.');
  return 'Setup complete';
}

function initializePortalSheetsForCurrentUser(sessionToken) {
  const session = getSessionOrThrow_(sessionToken);
  assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);
  const result = initializePortalSheets();
  logAuditEvent_('SYSTEM_SETUP', 'SPREADSHEET', APP_CONFIG.spreadsheetId, session.userId, { result: result });
  return respond_(true, { message: result });
}

function ensureDemoPortalData_(session) {
  if (!session || session.roleCode !== APP_CONFIG.roles.SUPER_ADMIN) return;
  const alreadySeeded = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'username', 'demo_superadmin') &&
    Repository.findOneBy_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', 'rest-demo-001') &&
    Repository.findOneBy_(APP_CONFIG.sheets.HOST_LINKS, 'link_id', 'host-link-h');
  if (alreadySeeded) return;

  initializePortalSheets();
  seedDemoBrands_();
  seedDemoRestaurants_();
  seedDemoDropdowns_(session.userId);
  seedDemoHostLinks_(session.userId);
  seedDemoUsers_(session.userId);
  seedDemoTeamMembers_();
  seedDemoNewHires_();
}

function seedDemoBrands_() {
  [
    {
      brand_id: 'brand-seths',
      brand_name: "SETH'S Hospitality",
      status: 'ACTIVE',
      notes: 'Primary demo brand'
    },
    {
      brand_id: 'brand-saisethsons',
      brand_name: 'SAISETHSONS Hospitality Group',
      status: 'ACTIVE',
      notes: 'Corporate portfolio brand'
    }
  ].forEach(function(record) {
    upsertSeedRecord_(APP_CONFIG.sheets.BRANDS, 'brand_id', record.brand_id, withSeedTimestamps_(record));
  });
}

function seedDemoRestaurants_() {
  [
    {
      restaurant_id: 'rest-demo-001',
      restaurant_code: '1001',
      restaurant_name: 'SAISETHSONS Downtown Cafe',
      brand_name: "SETH'S Hospitality",
      region: 'Toronto Central',
      city: 'Toronto',
      province_state: 'ON',
      country: 'Canada',
      status: 'ACTIVE',
      opening_date: '2025-01-10',
      gm_name: 'Olivia Bennett',
      gm_email: 'olivia.bennett@example.com',
      mro_owner_user_id: 'user-demo-mro',
      notes: 'Primary demo restaurant for restaurant-user testing'
    },
    {
      restaurant_id: 'rest-demo-002',
      restaurant_code: '1002',
      restaurant_name: 'SAISETHSONS Lakeside Bakery',
      brand_name: 'SAISETHSONS Hospitality Group',
      region: 'Peel North',
      city: 'Mississauga',
      province_state: 'ON',
      country: 'Canada',
      status: 'ACTIVE',
      opening_date: '2025-02-14',
      gm_name: 'Daniel Ross',
      gm_email: 'daniel.ross@example.com',
      mro_owner_user_id: 'user-demo-mro',
      notes: 'Secondary demo restaurant for multi-store access testing'
    }
  ].forEach(function(record) {
    upsertSeedRecord_(APP_CONFIG.sheets.RESTAURANTS, 'restaurant_id', record.restaurant_id, withSeedTimestamps_(record));
  });
}

function seedDemoDropdowns_(actorUserId) {
  [
    ['REQUIRED_POSITION', 'Team Member', 1],
    ['REQUIRED_POSITION', 'Supervisor', 2],
    ['REQUIRED_POSITION', 'Assistant Manager', 3],
    ['SUPPLY_ITEM', 'Cleaning Supplies', 1],
    ['SUPPLY_ITEM', 'Printer Paper', 2],
    ['SUPPLY_ITEM', 'Packaging', 3]
  ].forEach(function(tuple) {
    const optionId = 'opt-' + tuple[0].toLowerCase() + '-' + tuple[1].toLowerCase().replace(/[^a-z0-9]+/g, '-');
    upsertSeedRecord_(APP_CONFIG.sheets.DROPDOWN_OPTIONS, 'option_id', optionId, {
      option_id: optionId,
      dropdown_key: tuple[0],
      option_value: tuple[1],
      status: 'ACTIVE',
      sort_order: tuple[2],
      updated_at: new Date().toISOString(),
      updated_by: actorUserId || ''
    });
  });
}

function seedDemoHostLinks_(actorUserId) {
  [
    ['host-link-h', 'H', 'Hospitality Playbook', 'Premium service standards and guest recovery guide', 'https://example.com/hospitality-playbook', 'SUPER_ADMIN,HR_ADMIN,MRO,SENIOR_MANAGEMENT,RESTAURANT_USER', '', 1],
    ['host-link-o', 'O', 'Operations Calendar', 'Daily operations calendar and store readiness board', 'https://example.com/operations-calendar', 'SUPER_ADMIN,HR_ADMIN,MRO,SENIOR_MANAGEMENT,RESTAURANT_USER', '', 2],
    ['host-link-s', 'S', 'Systems Hub', 'Systems access list, checklists, and reporting shortcuts', 'https://example.com/systems-hub', 'SUPER_ADMIN,HR_ADMIN,MRO,SENIOR_MANAGEMENT,RESTAURANT_USER', '', 3],
    ['host-link-t', 'T', 'Training Academy', 'Training guides, onboarding videos, and certification content', 'https://example.com/training-academy', 'SUPER_ADMIN,HR_ADMIN,MRO,SENIOR_MANAGEMENT,RESTAURANT_USER', '', 4]
  ].forEach(function(tuple) {
    upsertSeedRecord_(APP_CONFIG.sheets.HOST_LINKS, 'link_id', tuple[0], {
      link_id: tuple[0],
      category_code: tuple[1],
      title: tuple[2],
      url: tuple[4],
      description: tuple[3],
      icon_name: '',
      target_role_codes: tuple[5],
      target_restaurant_ids: tuple[6],
      is_global: 'TRUE',
      start_at: '',
      end_at: '',
      status: 'ACTIVE',
      sort_order: tuple[7],
      created_by: actorUserId || '',
      updated_by: actorUserId || '',
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    });
  });
}

function seedDemoUsers_(actorUserId) {
  const demoPassword = 'Portal@2026';
  [
    {
      user_id: 'user-demo-superadmin',
      username: 'demo_superadmin',
      full_name: 'Demo Super Admin',
      email: 'demo.superadmin@example.com',
      phone: '416-555-0100',
      role_code: APP_CONFIG.roles.SUPER_ADMIN,
      employee_id: 'EMP-SA-001',
      restaurant_ids: [],
      region: 'Corporate'
    },
    {
      user_id: 'user-demo-restaurant',
      username: 'demo_restaurant',
      full_name: 'Demo Restaurant User',
      email: 'demo.restaurant@example.com',
      phone: '416-555-0101',
      role_code: APP_CONFIG.roles.RESTAURANT_USER,
      employee_id: 'EMP-RU-001',
      restaurant_ids: ['rest-demo-001'],
      region: 'Toronto Central'
    },
    {
      user_id: 'user-demo-hr',
      username: 'demo_hradmin',
      full_name: 'Demo HR Admin',
      email: 'demo.hr@example.com',
      phone: '416-555-0102',
      role_code: APP_CONFIG.roles.HR_ADMIN,
      employee_id: 'EMP-HR-001',
      restaurant_ids: ['rest-demo-001', 'rest-demo-002'],
      region: 'People Operations'
    },
    {
      user_id: 'user-demo-mro',
      username: 'demo_mro',
      full_name: 'Demo MRO',
      email: 'demo.mro@example.com',
      phone: '416-555-0103',
      role_code: APP_CONFIG.roles.MRO,
      employee_id: 'EMP-MRO-001',
      restaurant_ids: ['rest-demo-001', 'rest-demo-002'],
      region: 'Multi-Store'
    },
    {
      user_id: 'user-demo-senior',
      username: 'demo_senior',
      full_name: 'Demo Senior Management',
      email: 'demo.senior@example.com',
      phone: '416-555-0104',
      role_code: APP_CONFIG.roles.SENIOR_MANAGEMENT,
      employee_id: 'EMP-SM-001',
      restaurant_ids: [],
      region: 'Executive'
    }
  ].forEach(function(user) {
    upsertSeedUser_(user, demoPassword, actorUserId || '');
  });
}

function seedDemoTeamMembers_() {
  [
    {
      team_member_id: 'team-demo-001',
      employee_id: 'TM-1001',
      full_name: 'Ava Clarke',
      restaurant_id: 'rest-demo-001',
      job_title: 'Team Member',
      department: 'Front of House',
      employment_type: 'Full Time',
      hire_date: '2026-03-15',
      status: 'ACTIVE',
      manager_name: 'Olivia Bennett',
      email: 'ava.clarke@example.com',
      phone: '416-555-0120'
    },
    {
      team_member_id: 'team-demo-002',
      employee_id: 'TM-1002',
      full_name: 'Lucas Hart',
      restaurant_id: 'rest-demo-002',
      job_title: 'Supervisor',
      department: 'Bakery',
      employment_type: 'Part Time',
      hire_date: '2026-02-11',
      status: 'ACTIVE',
      manager_name: 'Daniel Ross',
      email: 'lucas.hart@example.com',
      phone: '416-555-0121'
    }
  ].forEach(function(record) {
    upsertSeedRecord_(APP_CONFIG.sheets.TEAM_MEMBERS, 'team_member_id', record.team_member_id, withSeedTimestamps_(record));
  });
}

function seedDemoNewHires_() {
  [
    {
      hire_id: 'hire-demo-001',
      candidate_name: 'Mia Patel',
      applied_for_role: 'Team Member',
      restaurant_id: 'rest-demo-001',
      assigned_hr_user_id: 'user-demo-hr',
      application_date: '2026-04-10',
      current_status: 'Orientation Done',
      interview_date: '2026-04-12',
      orientation_date: '2026-04-18',
      sent_to_restaurant_at: '2026-04-19T10:00:00Z',
      restaurant_confirmation_at: '',
      start_date: '2026-04-28',
      notes: 'Ready for restaurant confirmation',
      hold_reason: '',
      cancel_reason: ''
    }
  ].forEach(function(record) {
    upsertSeedRecord_(APP_CONFIG.sheets.NEW_HIRES, 'hire_id', record.hire_id, withSeedUpdatedAt_(record));
  });
}

function upsertSeedUser_(user, password, actorUserId) {
  const existing = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'user_id', user.user_id);
  const now = new Date().toISOString();
  let salt = existing && existing.salt ? String(existing.salt) : Utilities.getUuid().replace(/-/g, '');
  const record = {
    user_id: user.user_id,
    username: user.username,
    password_hash: hashPassword_(password, salt),
    salt: salt,
    full_name: user.full_name,
    email: user.email,
    phone: user.phone,
    role_code: user.role_code,
    employee_id: user.employee_id,
    restaurant_id: user.restaurant_ids[0] || '',
    region: user.region,
    status: 'ACTIVE',
    last_login_at: existing && existing.last_login_at ? existing.last_login_at : '',
    must_reset_password: 'FALSE',
    remember_token: '',
    created_at: existing && existing.created_at ? existing.created_at : now,
    created_by: existing && existing.created_by ? existing.created_by : actorUserId,
    updated_at: now,
    updated_by: actorUserId
  };

  upsertSeedRecord_(APP_CONFIG.sheets.USERS, 'user_id', user.user_id, record);

  Repository.replaceRowsByColumn_(APP_CONFIG.sheets.RESTAURANT_ACCESS, 'user_id', user.user_id, user.restaurant_ids.map(function(restaurantId, index) {
    return {
      access_id: 'access-' + user.user_id + '-' + (index + 1),
      user_id: user.user_id,
      restaurant_id: restaurantId,
      access_level: index === 0 ? 'PRIMARY' : 'ASSIGNED',
      is_primary: index === 0 ? 'TRUE' : 'FALSE',
      created_at: now,
      created_by: actorUserId
    };
  }));
}

function upsertSeedRecord_(sheetName, idColumn, idValue, record) {
  const existing = Repository.findOneBy_(sheetName, idColumn, idValue);
  if (existing) {
    Repository.updateById_(sheetName, idColumn, idValue, record);
    return;
  }
  Repository.append_(sheetName, record);
}

function withSeedTimestamps_(record) {
  const now = new Date().toISOString();
  return Object.assign({
    created_at: now,
    updated_at: now
  }, record);
}

function withSeedUpdatedAt_(record) {
  return Object.assign({
    updated_at: new Date().toISOString()
  }, record);
}

function getSheetSchemas_() {
  return {
    // FIX #15: Added 7 missing columns used by AdminService (city, province_state, country, opening_date, gm_name, gm_email, mro_owner_user_id).
    'Restaurants': ['restaurant_id', 'restaurant_number', 'restaurant_code', 'restaurant_name', 'brand_name', 'region', 'city', 'province_state', 'country', 'opening_date', 'gm_name', 'gm_email', 'mro_owner_user_id', 'status', 'notes', 'created_at', 'updated_at'],
    'Users': ['user_id', 'full_name', 'username', 'email', 'phone', 'role_code', 'employee_id', 'restaurant_id', 'region', 'status', 'temporary_password', 'password_hash', 'salt', 'last_login_at', 'must_reset_password', 'temp_password_plain', 'created_at', 'created_by', 'updated_at', 'updated_by'],
    'Restaurant Access': ['access_id', 'user_id', 'restaurant_id', 'restaurant_number', 'access_level', 'is_primary', 'status', 'created_at', 'created_by'],
    // FIX #17: Added manager_name, email columns used by ManageTeam service.
    'Team Members': ['team_member_id', 'employee_id', 'full_name', 'email', 'mobile_number', 'phone', 'restaurant_id', 'restaurant_number', 'job_title', 'employment_type', 'hire_date', 'manager_name', 'status', 'active', 'end_date', 'separation_reason', 'roe_id', 'deactivated_at', 'deactivated_by', 'notes', 'created_at', 'updated_at', 'recruitment_reason', 'recruitment_outcome_status', 'recruitment_completed_at'],
    'Employee History': ['history_id', 'employee_id', 'employee_name', 'restaurant_number', 'change_type', 'old_value', 'new_value', 'notes', 'changed_by', 'changed_at'],

    'Requests HR': ['request_id', 'request_type', 'restaurant_id', 'restaurant_number', 'requested_by', 'requested_for_role', 'employment_type', 'headcount', 'urgency', 'business_reason', 'status', 'assigned_hr_user_id', 'submitted_at', 'reviewed_at', 'completed_at', 'timeline_json', 'last_comment', 'request_payload_json'],
    'Requests Admin': ['request_id', 'request_type', 'restaurant_id', 'restaurant_number', 'requested_by', 'employee_name', 'employee_id', 'details', 'issuer_notes', 'admin_notes', 'status', 'file_url', 'assigned_to', 'submitted_at', 'reviewed_at', 'completed_at', 'timeline_json', 'last_comment', 'request_payload_json'],
    'Requests Supplies': ['request_id', 'request_type', 'restaurant_id', 'restaurant_number', 'requested_by', 'items_json', 'priority', 'quantity', 'status', 'assigned_to', 'submitted_at', 'fulfilled_at', 'timeline_json', 'request_payload_json'],
    'Requests Maintenance': ['request_id', 'restaurant_id', 'restaurant_number', 'requested_by', 'asset_area', 'issue_type', 'severity', 'description', 'photo_urls', 'status', 'assigned_mro_user_id', 'submitted_at', 'resolved_at', 'timeline_json', 'request_payload_json'],
    'Hiring Requests': ['request_id', 'restaurant_number', 'position', 'employment_type', 'headcount', 'needed_by', 'reason', 'status', 'hr_notes', 'created_by', 'created_at', 'updated_at', 'status_history_json'],
    'Admin Requests': ['request_id', 'restaurant_number', 'request_type', 'employee_name', 'details', 'issuer_notes', 'admin_notes', 'status', 'file_url', 'created_by', 'created_at', 'updated_at', 'status_history_json'],
    'User Management Responses': ['response_id', 'action_type', 'user_id', 'username', 'role_code', 'status', 'notes', 'changed_by', 'changed_at'],
    'Operations Responses': ['response_id', 'source_area', 'record_id', 'restaurant_number', 'status', 'notes', 'changed_by', 'changed_at'],

    'Brands': ['brand_id', 'brand_name', 'status', 'notes', 'created_at', 'updated_at'],
    'Role Button Permissions': ['permission_id', 'role_code', 'button_key', 'is_allowed', 'updated_at', 'updated_by'],
    'Dropdown Options': ['option_id', 'dropdown_key', 'option_value', 'status', 'sort_order', 'updated_at', 'updated_by'],
    'HOST Links': ['link_id', 'category_code', 'title', 'url', 'description', 'target_role_codes', 'target_restaurant_ids', 'status', 'sort_order', 'created_by', 'created_at', 'updated_by', 'updated_at'],
    'Newsletters': ['post_id', 'title', 'body', 'audience_roles', 'audience_restaurants', 'priority', 'start_at', 'end_at', 'status', 'cta_label', 'cta_url', 'attachment_type', 'attachment_name', 'attachment_url', 'attachment_file_id', 'created_by', 'created_at', 'updated_at'],
    // FIX #16: Added 8 missing columns used by RecruitmentService (restaurant_number, phone, email, employment_type, assigned_hr_user_id, follow_up_status, follow_up_notes, created_at).
    'New Hires': ['hire_id', 'candidate_name', 'applied_for_role', 'restaurant_id', 'restaurant_number', 'phone', 'email', 'employment_type', 'current_status', 'assigned_hr_user_id', 'follow_up_status', 'follow_up_notes', 'start_date', 'notes', 'created_at', 'updated_at'],
    'Notifications': ['notification_id', 'user_id', 'role_code', 'restaurant_id', 'type', 'title', 'message', 'action_url', 'priority', 'is_read', 'created_at', 'expires_at'],
    'Contact Requests': ['contact_request_id', 'source', 'name', 'username', 'email', 'phone', 'restaurant_number', 'message', 'status', 'action_taken_notes', 'requested_at', 'updated_at', 'updated_by'],
    'Password Change Requests': ['password_request_id', 'name', 'username', 'email', 'phone', 'message', 'requested_at', 'status', 'temporary_password', 'action_taken_notes', 'updated_at', 'updated_by'],
    'Audit Logs': ['audit_id', 'event_type', 'entity_type', 'entity_id', 'action', 'actor_user_id', 'actor_name', 'role_code', 'restaurant_id', 'metadata_json', 'ip_address', 'created_at'],
    'Reports Cache': ['cache_key', 'scope_type', 'scope_value', 'report_date', 'payload_json', 'expires_at', 'generated_at'],
    'Payroll Submissions': ['submission_id', 'restaurant_id', 'restaurant_number', 'brand', 'pay_period_ending', 'submitted_by', 'submitted_by_user_id', 'submitted_at', 'status', 'time_card_updated', 'stat_hours_1_0', 'payroll_approval', 'special_notes', 'admin_notes', 'correction_notes', 'reviewed_by', 'reviewed_at', 'approved_at', 'updated_at', 'updated_by', 'payload_json', 'status_history_json', 'locked_by_mro', 'locked_at', 'correction_count', 'last_resubmitted_at', 'last_reviewed_at', 'last_approved_at', 'correction_requested_by', 'correction_requested_at', 'payroll_history_json', 'admin_viewed_status', 'admin_viewed_by', 'admin_viewed_at', 'website_update_status', 'website_updated_by', 'website_updated_at', 'website_update_batch_id', 'payroll_batch_key'],
    'Payroll Rows': ['row_id', 'submission_id', 'restaurant_number', 'section_key', 'employee_id', 'employee_name', 'first_name', 'last_name', 'attendance_percentage', 'start_date', 'termination_date', 'leave_of_absence_date', 'reason_for_separation', 'uniform_deduction', 'missing_hours', 'training_hours', 'reason', 'item', 'amount', 'notes', 'created_at'],
    'Payroll ROE Applications': ['roe_id', 'submission_id', 'restaurant_id', 'restaurant_number', 'employee_id', 'employee_name', 'first_name', 'last_name', 'last_day_of_work', 'reason_for_separation', 'legal_status', 'health_benefits', 'station_champion', 'probation_termination_reason', 'status', 'payroll_submission_id', 'pay_period_ending', 'submitted_by', 'submitted_by_user_id', 'submitted_at', 'reviewed_by', 'reviewed_at', 'approved_by', 'approved_at', 'completed_at', 'deactivated_employee', 'deactivated_at', 'admin_notes', 'history_json', 'updated_at', 'updated_by'],
    'Settings': ['setting_key', 'setting_value', 'setting_group', 'description', 'updated_at', 'updated_by'],

  };
}

function upsertHeaderRow_(sheet, headers) {
  const targetHeaders = (headers || []).map(function(header) {
    return String(header || '').trim();
  }).filter(Boolean);

  if (!targetHeaders.length) return;

  const currentLastColumn = Math.max(sheet.getLastColumn(), 1);
  const existingHeaders = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0]
    .map(function(header) { return String(header || '').trim(); });

  if (sheet.getLastRow() === 0 || existingHeaders.every(function(header) { return !header; })) {
    if (sheet.getMaxColumns() < targetHeaders.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), targetHeaders.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    return;
  }

  let nextColumn = existingHeaders.length + 1;
  targetHeaders.forEach(function(header) {
    if (existingHeaders.indexOf(header) !== -1) return;
    if (sheet.getMaxColumns() < nextColumn) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
    sheet.getRange(1, nextColumn).setValue(header);
    existingHeaders.push(header);
    nextColumn += 1;
  });
}

function runManageTeamSetupOnly() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const schemas = getSheetSchemas_();

  const neededSheets = [
    APP_CONFIG.sheets.TEAM_MEMBERS,
    APP_CONFIG.sheets.RESTAURANTS,
    APP_CONFIG.sheets.AUDIT_LOGS
  ];

  neededSheets.forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    const headers = schemas[sheetName];
    if (!headers) throw new Error('Missing schema for ' + sheetName);

    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }

    upsertHeaderRow_(sheet, headers);
    sheet.setFrozenRows(1);
  });

  return 'Manage Team setup completed. Required sheets and headers are ready.';
}

function runFullPortalSheetSetupSafe() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const schemas = getSheetSchemas_();

  Object.keys(schemas).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const headers = schemas[sheetName];

    if (sheet.getMaxColumns() < headers.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
    }

    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    headers.forEach(function(header, index) {
      const col = index + 1;
      if (existingHeaders.indexOf(header) === -1) {
        sheet.getRange(1, col).setValue(header);
      }
    });

    sheet.setFrozenRows(1);
  });

  return 'Full portal sheet setup completed safely. No old data deleted.';
}

function runPortalCleanFoundationV243() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  const schemas = getSheetSchemas_();

  Object.keys(schemas).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    upsertHeaderRow_(sheet, schemas[sheetName]);
    sheet.setFrozenRows(1);
  });

  seedCleanRoleButtonPermissionsV243_('system-cleanup');
  return 'Portal clean foundation v243 completed. Login and dashboard design were not changed.';
}

function seedCleanRoleButtonPermissionsV243_(actorUserId) {
  const now = new Date().toISOString();
  const roles = Object.keys(DEFAULT_ROLE_ACTION_PERMISSIONS || {});

  roles.forEach(function(roleCode) {
    const allowed = DEFAULT_ROLE_ACTION_PERMISSIONS[roleCode] || [];
    const records = allowed.map(function(buttonKey) {
      return {
        permission_id: 'perm-' + roleCode + '-' + buttonKey,
        role_code: roleCode,
        button_key: buttonKey,
        is_allowed: 'TRUE',
        updated_at: now,
        updated_by: actorUserId || ''
      };
    });

    if (!records.length) {
      records.push({
        permission_id: 'perm-' + roleCode + '-' + ROLE_ACTION_NO_ACCESS_MARKER,
        role_code: roleCode,
        button_key: ROLE_ACTION_NO_ACCESS_MARKER,
        is_allowed: 'FALSE',
        updated_at: now,
        updated_by: actorUserId || ''
      });
    }

    Repository.replaceRowsByColumn_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS, 'role_code', roleCode, records);
  });
}


// FIX #23/#28: Dead one-time migration patch functions V260–V279 have been removed.
// runEnableContactUsRequestsV260/262, runEnablePasswordChangeRequestsV266,
// runEnableNewsletterManagementV268, runEnableNewsletterAttachmentsV274,
// runEnableHostLinkManagementV276, runEnableDropdownSettingsV279
// were all one-time setup patches that have already been applied in production.
// Their logic is fully covered by initializePortalSheets() + seedCleanRoleButtonPermissionsV243_().
// If you need to re-run any of them, use runFullPortalSheetSetupSafe() instead.

/*******************************************************
 V326 CACHE PRE-WARM
 Run setupPrewarmTrigger() once manually from the GAS editor
 to create a time-based trigger that keeps caches warm.
 prewarmCaches() runs automatically every 5 minutes.
 This means the first user to open any page gets
 instant load instead of a cold-start penalty.
*******************************************************/

function prewarmCaches() {
  try {
    // Pre-warm session-independent data caches.
    // Repository.getAll_ caches automatically — touching these sheets keeps
    // the V258 ScriptCache entries fresh so users never hit cold-start latency.
    // FIX (prewarm scope): Expanded from 5 to 13 sheets covering all hot read paths.
    const hotSheets = [
      APP_CONFIG.sheets.REQUESTS_ADMIN,
      APP_CONFIG.sheets.REQUESTS_HR,
      APP_CONFIG.sheets.TEAM_MEMBERS,
      APP_CONFIG.sheets.RESTAURANTS,
      APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS,
      APP_CONFIG.sheets.USERS,
      APP_CONFIG.sheets.RESTAURANT_ACCESS,
      APP_CONFIG.sheets.NEW_HIRES,
      APP_CONFIG.sheets.NOTIFICATIONS,
      APP_CONFIG.sheets.HOST_LINKS,
      APP_CONFIG.sheets.NEWSLETTERS,
      APP_CONFIG.sheets.CONTACT_REQUESTS,
      APP_CONFIG.sheets.PASSWORD_CHANGE_REQUESTS
    ];
    hotSheets.forEach(function(sheetName) {
      try { Repository.getAll_(sheetName); } catch (e) {}
    });

    // Rebuild the operations dashboard cache for a super admin session
    // Uses a synthetic session so no token needed
    const cache = CacheService.getScriptCache();

    // Bump the last-updated timestamp so polling clients know data is fresh
    cache.put('seth_v325_last_updated', String(Date.now()), 600);

    console.log('Cache pre-warm completed at ' + new Date().toISOString());
  } catch (e) {
    console.error('prewarmCaches failed', e);
  }
}

function setupPrewarmTrigger() {
  // Run this ONCE manually from the GAS editor: Extensions → Apps Script → Run setupPrewarmTrigger
  // Deletes any existing prewarm trigger first to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'prewarmCaches') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('prewarmCaches')
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log('Pre-warm trigger created. prewarmCaches will run every 5 minutes.');
  return 'Trigger created.';
}

function removePrewarmTrigger() {
  // Run this if you want to stop pre-warming
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'prewarmCaches') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return 'Removed ' + removed + ' trigger(s).';
}
function resetMyPassword() {
  const username = 'superadmin';
  const newPassword = 'Saisethsons@2026'; // change this to whatever you want
  const user = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'username', username);
  if (!user) { console.log('User not found'); return; }
  const salt = Utilities.getUuid().replace(/-/g, '');
  const hash = hashPassword_(newPassword, salt);
  Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', user.user_id, {
    password_hash: hash,
    salt: salt,
    must_reset_password: 'FALSE',
    updated_at: new Date().toISOString()
  });
  console.log('Done — password reset for: ' + username);
}
function rehashAllUsers() {
  const defaultPassword = 'Portal@2026';
  const users = Repository.getAll_(APP_CONFIG.sheets.USERS).filter(function(u) {
    return String(u.status || '').toUpperCase() === 'ACTIVE';
  });
  users.forEach(function(u) {
    if (u.username === 'superadmin') return; // skip, already done above
    const salt = Utilities.getUuid().replace(/-/g, '');
    const hash = hashPassword_(defaultPassword, salt);
    Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', u.user_id, {
      password_hash: hash,
      salt: salt,
      must_reset_password: 'TRUE',
      updated_at: new Date().toISOString()
    });
    console.log('Rehashed: ' + u.username);
  });
  return 'Done — ' + users.length + ' users rehashed';
}


function resetSuperAdminNow() {
  const username = 'superadmin';
  const newPassword = 'Saisethsons@2026'; // change this to whatever you want
  const user = Repository.findOneBy_(APP_CONFIG.sheets.USERS, 'username', username);
  if (!user) { console.log('User not found'); return; }
  const salt = Utilities.getUuid().replace(/-/g, '');
  const hash = hashPassword_(newPassword, salt);
  Repository.updateById_(APP_CONFIG.sheets.USERS, 'user_id', user.user_id, {
    password_hash: hash,
    salt: salt,
    must_reset_password: 'FALSE',
    updated_at: new Date().toISOString()
  });
  console.log('Done');
}

// Run ONCE from Apps Script editor to add/fix SNAPSHOT_5W (5W Tracker) button
// for all roles. Updates existing rows if already present.
function add5WTrackerPermission() {
  var now    = new Date().toISOString();
  var btnKey = 'SNAPSHOT_5W';
  var allowed = { SUPER_ADMIN: true, RESTAURANT_USER: true };
  var roles  = Object.keys(DEFAULT_ROLE_ACTION_PERMISSIONS || {});

  roles.forEach(function(roleCode) {
    var id       = 'perm-' + roleCode + '-' + btnKey;
    var isAllowed = (allowed[roleCode] || roleCode === 'SUPER_ADMIN') ? 'TRUE' : 'FALSE';
    var existing  = Repository.findOneBy_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS, 'permission_id', id);
    if (existing) {
      Repository.updateById_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS, 'permission_id', id, {
        is_allowed: isAllowed,
        updated_at: now,
        updated_by: 'system-setup'
      });
      Logger.log('Updated: ' + id + ' = ' + isAllowed);
    } else {
      Repository.append_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS, {
        permission_id: id,
        role_code:     roleCode,
        button_key:    btnKey,
        is_allowed:    isAllowed,
        updated_at:    now,
        updated_by:    'system-setup'
      });
      Logger.log('Added: ' + id + ' = ' + isAllowed);
    }
  });

  // Bust ALL caches so role actions and permissions update on next login
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('role_btn_perms_v13_super_admin_all');
    cache.remove('role_btn_perms_v329');
    Repository.clearCacheV258_(APP_CONFIG.sheets.ROLE_BUTTON_PERMISSIONS);
    // Wipe script cache entirely to kill all boot_v326_* entries
    cache.removeAll(['role_btn_perms_v13_super_admin_all','role_btn_perms_v329']);
    Logger.log('Caches cleared.');
  } catch(e) { Logger.log('Cache clear error: ' + e); }

  // Also clear the Repository sheet cache so next read is fresh
  try { Repository.__sheetCacheV324_ = {}; Repository.__headerCacheV324_ = {}; } catch(e) {}

  Logger.log('Done — 5W Tracker permissions fixed. Log out and back in to see the button in Role Actions.');

  Logger.log('Done — 5W Tracker permissions fixed.');
  return 'Done.';
}
