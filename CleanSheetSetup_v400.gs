/**
 * CleanSheetSetup_v406
 *
 * Safer clean rebuild for SAISETHSONS HOST portal.
 *
 * What this version does:
 * 1. Keeps all tabs used by current webapp buttons and backend flows.
 * 2. Keeps legacy tab names that your webapp may still call.
 * 3. Creates a backup before changing anything.
 * 4. Removes filters safely before recreating filters.
 * 5. Rebuilds columns in clean order.
 * 6. Preserves old row data by matching existing header names.
 *
 * Run first:
 * previewPortalSheetCleanupV406()
 *
 * Then run:
 * rebuildPortalSheetsCleanV406()
 */


/**
 * Tabs that must stay because the webapp still uses them.
 * Do not remove a tab from this list unless the related button and backend code are removed.
 */
const PORTAL_KEEP_TABS_V406 = [
  'APP_CONFIG',

  // Main access and setup
  'Users',
  'Restaurants',
  'Restaurant Access',
  'Role Button Permissions',
  'Dropdown Options',
  'Brands',

  // HOST and newsletter
  'HOST Links',
  'Newsletters',

  // Team and recruitment
  'Team Members',
  'New Hires',
  'Hiring Requests',
  'Requests HR',

  // Admin request flows
  'Admin Requests',
  'Requests Admin',

  // Contact and password buttons
  'Contact Requests',
  'Password Change Requests',

  // Payroll
  'Payroll Submissions',
  'Payroll Reviews',
  'Payroll Batches',
  'Payroll ROE Applications',
  'Payroll Audit',

  // Notifications and audit
  'Notifications',
  'Audit Logs',
  'Audit Log',
  'Security Audit'
];


/**
 * Clean column structure for each needed tab.
 * Legacy tabs are kept with safe columns so old functions do not break.
 */
const PORTAL_SCHEMAS_V406 = {
  'APP_CONFIG': [
    'key',
    'value',
    'notes',
    'updated_at',
    'updated_by'
  ],

  'Users': [
    'user_id',
    'username',
    'full_name',
    'email',
    'phone',
    'role_code',
    'employee_id',
    'restaurant_id',
    'assigned_restaurant_ids',
    'region',
    'status',
    'password_hash',
    'temporary_password',
    'created_at',
    'updated_at',
    'last_login_at',
    'notes'
  ],

  'Restaurants': [
    'restaurant_id',
    'restaurant_number',
    'restaurant_name',
    'brand',
    'region',
    'status',
    'notes',
    'created_at',
    'updated_at'
  ],

  'Restaurant Access': [
    'access_id',
    'user_id',
    'username',
    'restaurant_id',
    'restaurant_number',
    'access_level',
    'is_primary',
    'status',
    'created_at',
    'updated_at'
  ],

  'Role Button Permissions': [
    'role_code',
    'allowed_buttons',
    'status',
    'updated_at',
    'updated_by'
  ],

  'Dropdown Options': [
    'option_id',
    'dropdown_key',
    'option_value',
    'status',
    'sort_order',
    'created_at',
    'updated_at'
  ],

  'Brands': [
    'brand_id',
    'brand_name',
    'status',
    'notes',
    'created_at',
    'updated_at'
  ],

  'HOST Links': [
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
    'created_at',
    'updated_at'
  ],

  'Newsletters': [
    'post_id',
    'title',
    'body',
    'audience_roles',
    'audience_restaurants',
    'start_at',
    'end_at',
    'status',
    'created_at',
    'updated_at',
    'created_by'
  ],

  'Team Members': [
    'team_member_id',
    'employee_id',
    'full_name',
    'name',
    'mobile',
    'email',
    'restaurant_id',
    'restaurant_number',
    'job_title',
    'department',
    'employment_type',
    'hire_date',
    'hired_date',
    'status',
    'recruitment_outcome_status',
    'quit_reason',
    'manager_name',
    'created_at',
    'updated_at',
    'completed_at',
    'notes'
  ],

  'New Hires': [
    'hire_id',
    'candidate_name',
    'full_name',
    'mobile',
    'email',
    'applied_for_role',
    'restaurant_id',
    'restaurant_number',
    'assigned_hr_user_id',
    'application_date',
    'current_status',
    'interview_date',
    'orientation_date',
    'start_date',
    'sent_to_restaurant_at',
    'restaurant_confirmation_at',
    'notes',
    'hold_reason',
    'cancel_reason',
    'created_at',
    'updated_at'
  ],

  'Hiring Requests': [
    'request_id',
    'restaurant_id',
    'restaurant_number',
    'required_position',
    'employment_type',
    'department',
    'availability',
    'required_from',
    'headcount',
    'business_reason',
    'issuer_notes',
    'admin_notes',
    'status',
    'status_history',
    'created_by',
    'created_at',
    'updated_at'
  ],

  'Requests HR': [
    'request_id',
    'request_type',
    'restaurant_id',
    'restaurant_number',
    'position',
    'employment_type',
    'department',
    'availability',
    'required_from',
    'headcount',
    'reason',
    'issuer_notes',
    'admin_notes',
    'status',
    'status_history',
    'created_by',
    'created_at',
    'updated_at'
  ],

  'Admin Requests': [
    'request_id',
    'request_type',
    'restaurant_id',
    'restaurant_number',
    'employee_id',
    'employee_name',
    'letter_type',
    'gender',
    'position',
    'manual_position',
    'employment_type',
    'started_date',
    'supply_type',
    'supply_details',
    'issuer_notes',
    'admin_notes',
    'attachments',
    'status',
    'status_history',
    'created_by',
    'created_at',
    'updated_at',
    'received_at'
  ],

  'Requests Admin': [
    'request_id',
    'request_type',
    'restaurant_id',
    'restaurant_number',
    'employee_id',
    'employee_name',
    'letter_type',
    'gender',
    'position',
    'manual_position',
    'employment_type',
    'started_date',
    'supply_type',
    'supply_details',
    'reason',
    'issuer_notes',
    'admin_notes',
    'attachments',
    'status',
    'status_history',
    'created_by',
    'created_at',
    'updated_at',
    'received_at'
  ],

  'Contact Requests': [
    'contact_request_id',
    'name',
    'username',
    'email',
    'phone',
    'message',
    'source',
    'requested_at',
    'status',
    'action_taken_notes',
    'created_at',
    'updated_at'
  ],

  'Password Change Requests': [
    'password_request_id',
    'name',
    'username',
    'email',
    'phone',
    'message',
    'requested_at',
    'status',
    'temporary_password',
    'action_taken_notes',
    'created_at',
    'updated_at'
  ],

  'Payroll Submissions': [
    'submission_id',
    'restaurant_number',
    'pay_period_ending',
    'time_card_updated',
    'stat_hours_1_0',
    'payroll_approval',
    'special_notes',
    'payload_json',
    'status',
    'locked_by_mro',
    'viewed_by',
    'viewed_at',
    'approved_by',
    'approved_at',
    'correction_notes',
    'correction_count',
    'created_by',
    'created_at',
    'updated_at'
  ],

  'Payroll Reviews': [
    'review_id',
    'submission_id',
    'restaurant_number',
    'pay_period_ending',
    'review_status',
    'mro_notes',
    'reviewed_by',
    'reviewed_at',
    'created_at',
    'updated_at'
  ],

  'Payroll Batches': [
    'batch_id',
    'batch_key',
    'brand',
    'restaurant_number',
    'pay_period_ending',
    'submission_ids',
    'status',
    'created_by',
    'created_at',
    'updated_at'
  ],

  'Payroll ROE Applications': [
    'roe_id',
    'restaurant_number',
    'employee_id',
    'employee_name',
    'last_day_of_work',
    'reason_for_separation',
    'legal_status',
    'health_benefits',
    'station_champion',
    'probation_termination_reason',
    'payroll_submission_id',
    'status',
    'created_by',
    'created_at',
    'updated_at',
    'completed_at'
  ],

  'Payroll Audit': [
    'event_id',
    'event_type',
    'entity_type',
    'entity_id',
    'actor_email',
    'details_json',
    'created_at'
  ],

  'Notifications': [
    'notification_id',
    'title',
    'message',
    'target_role_codes',
    'target_restaurant_ids',
    'status',
    'created_at',
    'expires_at',
    'created_by'
  ],

  'Audit Logs': [
    'audit_id',
    'event',
    'entity_type',
    'entity_id',
    'actor_email',
    'metadata_json',
    'created_at'
  ],

  'Audit Log': [
    'audit_id',
    'event',
    'entity_type',
    'entity_id',
    'actor_email',
    'metadata_json',
    'created_at'
  ],

  'Security Audit': [
    'audit_id',
    'event',
    'username',
    'user_id',
    'ip_address',
    'user_agent',
    'status',
    'details',
    'created_at'
  ]
};


/**
 * Preview only.
 * No UI popup.
 * Check View > Logs after running.
 */
function previewPortalSheetCleanupV406() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets().map(sheet => sheet.getName());

  const keep = PORTAL_KEEP_TABS_V406.slice();
  const deleteList = allSheets.filter(name => !PORTAL_KEEP_TABS_V406.includes(name));

  Logger.log('KEEP TABS');
  keep.forEach(name => Logger.log('KEEP: ' + name));

  Logger.log('DELETE TABS');
  deleteList.forEach(name => Logger.log('DELETE: ' + name));

  Logger.log('IMPORTANT');
  Logger.log('Tabs like Requests HR, Requests Admin, Contact Requests, Password Change Requests, Brands, and Audit Log are now kept because your webapp still has buttons or backend references for them.');
  Logger.log('__TEMP_CLEAN_REBUILD__ can be deleted if it exists. It is only a temporary cleanup tab.');

  return {
    keep: keep,
    delete: deleteList
  };
}


/**
 * Main clean rebuild.
 */
function rebuildPortalSheetsCleanV406() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createPortalBackupV406_(ss);

  const keepSet = new Set(PORTAL_KEEP_TABS_V406);

  // Create all required tabs first.
  PORTAL_KEEP_TABS_V406.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // Rebuild headers, keep data where possible.
  PORTAL_KEEP_TABS_V406.forEach(name => {
    const headers = PORTAL_SCHEMAS_V406[name];
    if (!headers || !headers.length) return;

    const sheet = ss.getSheetByName(name);
    rebuildOnePortalSheetV406_(sheet, headers);
  });

  // Delete only tabs that are truly outside keep list.
  // Keep at least one sheet at all times.
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const name = sheet.getName();

    if (keepSet.has(name)) return;

    // Safe to remove temp tab.
    if (name === '__TEMP_CLEAN_REBUILD__') {
      if (ss.getSheets().length > 1) ss.deleteSheet(sheet);
      return;
    }

    if (ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
    }
  });

  reorderPortalTabsV406_(ss);
  SpreadsheetApp.flush();

  Logger.log('Clean rebuild V406 completed.');
}


/**
 * Backup all sheets into a new backup spreadsheet before rebuild.
 */
function createPortalBackupV406_(ss) {
  const tz = Session.getScriptTimeZone() || 'America/Toronto';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const backup = SpreadsheetApp.create('HOST Portal Backup Before Clean Rebuild ' + stamp);

  const defaultSheet = backup.getSheets()[0];

  ss.getSheets().forEach(sourceSheet => {
    const copied = sourceSheet.copyTo(backup);
    copied.setName(safeSheetNameV406_(sourceSheet.getName()));
  });

  if (backup.getSheets().length > 1) {
    backup.deleteSheet(defaultSheet);
  }

  Logger.log('Backup created: ' + backup.getUrl());
}


/**
 * Rebuild one sheet in clean column order.
 * Preserves existing data by matching current headers to new headers.
 */
function rebuildOnePortalSheetV406_(sheet, headers) {
  if (!sheet) return;

  removeFilterSafelyV406_(sheet);

  const oldValues = sheet.getDataRange().getValues();
  const oldHeaders = oldValues.length ? oldValues[0].map(v => String(v || '').trim()) : [];
  const oldRows = oldValues.length > 1 ? oldValues.slice(1) : [];

  const newRows = oldRows
    .filter(row => row.some(cell => String(cell || '').trim() !== ''))
    .map(row => {
      return headers.map(header => {
        const oldIndex = findHeaderIndexV406_(oldHeaders, header);
        return oldIndex >= 0 ? row[oldIndex] : '';
      });
    });

  sheet.clear();

  const rowCount = Math.max(newRows.length + 1, 2);
  const colCount = headers.length;

  sheet.getRange(1, 1, 1, colCount).setValues([headers]);

  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, colCount).setValues(newRows);
  }

  stylePortalSheetV406_(sheet, headers, rowCount, colCount);
}


/**
 * Header matching with common aliases.
 */
function findHeaderIndexV406_(oldHeaders, newHeader) {
  const target = normalizeHeaderV406_(newHeader);

  for (let i = 0; i < oldHeaders.length; i++) {
    if (normalizeHeaderV406_(oldHeaders[i]) === target) return i;
  }

  const aliases = getHeaderAliasesV406_();
  const possible = aliases[target] || [];

  for (let i = 0; i < oldHeaders.length; i++) {
    const normalizedOld = normalizeHeaderV406_(oldHeaders[i]);
    if (possible.includes(normalizedOld)) return i;
  }

  return -1;
}


function normalizeHeaderV406_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '');
}


function getHeaderAliasesV406_() {
  return {
    restaurant_number: ['restaurant_code', 'restaurantcode', 'restaurant', 'number', 'code'],
    restaurant_name: ['restaurantname', 'name'],
    brand: ['brand_name', 'brandname'],
    full_name: ['fullname', 'name', 'employee_name', 'team_member_name', 'candidate_name'],
    name: ['full_name', 'fullname', 'employee_name', 'team_member_name', 'candidate_name'],
    mobile: ['phone', 'mobile_number', 'mobilenumber', 'phone_number'],
    email: ['email_address', 'emailaddress'],
    hire_date: ['hired_date', 'hiredate', 'hired_date'],
    hired_date: ['hire_date', 'hiredate', 'hireddate'],
    status: ['current_status', 'request_status', 'employee_status'],
    created_at: ['timestamp', 'createdat', 'submitted_at', 'submittedat', 'requested_at'],
    updated_at: ['updatedat', 'last_updated'],
    request_id: ['requestid', 'id'],
    team_member_id: ['teammemberid', 'employee_id', 'employeeid'],
    employee_id: ['employeeid', 'team_member_id', 'teammemberid'],
    issuer_notes: ['notes', 'other_notes', 'business_reason', 'reason'],
    admin_notes: ['hr_notes', 'management_notes', 'action_taken_notes'],
    body: ['message', 'newsletter_message'],
    allowed_buttons: ['allowedbuttons']
  };
}


/**
 * Format sheet.
 */
function stylePortalSheetV406_(sheet, headers, rowCount, colCount) {
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange
    .setFontWeight('bold')
    .setBackground('#163a63')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(1, 1, rowCount, colCount)
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Alternating color by simple row background.
  if (rowCount > 1) {
    sheet.getRange(2, 1, rowCount - 1, colCount).setBackground('#ffffff');
  }

  headers.forEach((header, index) => {
    const col = index + 1;
    sheet.setColumnWidth(col, getWidthForHeaderV406_(header));
  });

  removeFilterSafelyV406_(sheet);

  // Create filter only if there is a valid range.
  if (rowCount >= 1 && colCount >= 1) {
    sheet.getRange(1, 1, rowCount, colCount).createFilter();
  }
}


function getWidthForHeaderV406_(header) {
  const h = String(header || '').toLowerCase();

  if (h.includes('id')) return 170;
  if (h.includes('email')) return 240;
  if (h.includes('name')) return 220;
  if (h.includes('message') || h.includes('notes') || h.includes('json') || h.includes('history') || h.includes('details')) return 320;
  if (h.includes('url')) return 320;
  if (h.includes('date') || h.includes('_at')) return 170;
  if (h.includes('status')) return 150;
  if (h.includes('restaurant')) return 190;
  if (h.includes('role')) return 210;
  if (h.includes('phone') || h.includes('mobile')) return 160;

  return 150;
}


function removeFilterSafelyV406_(sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (error) {
    Logger.log('Could not remove filter on ' + sheet.getName() + ': ' + error.message);
  }
}


function reorderPortalTabsV406_(ss) {
  PORTAL_KEEP_TABS_V406.forEach((name, index) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });
}


function safeSheetNameV406_(name) {
  let safe = String(name || 'Sheet').replace(/[\\\/\?\*\[\]\:]/g, ' ').trim();
  if (!safe) safe = 'Sheet';
  if (safe.length > 90) safe = safe.slice(0, 90);
  return safe;
}


/**
 * Optional helper.
 * Use this if you only want to format the current tabs and not delete anything.
 */
function formatPortalSheetsOnlyV406() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  PORTAL_KEEP_TABS_V406.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const headers = PORTAL_SCHEMAS_V406[name];
    if (!sheet || !headers || !headers.length) return;

    const lastRow = Math.max(sheet.getLastRow(), 2);
    const lastCol = Math.max(sheet.getLastColumn(), headers.length);

    removeFilterSafelyV406_(sheet);
    sheet.setFrozenRows(1);

    sheet.getRange(1, 1, 1, Math.min(lastCol, headers.length))
      .setFontWeight('bold')
      .setBackground('#163a63')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setWrap(true);

    headers.forEach((header, index) => {
      sheet.setColumnWidth(index + 1, getWidthForHeaderV406_(header));
    });

    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  });

  Logger.log('Formatting completed.');
}
