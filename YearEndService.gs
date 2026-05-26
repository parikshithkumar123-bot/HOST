// ═══════════════════════════════════════════════════════════════════════════
// YEAR END SERVICE
// Handles year-end archiving, transition, and multi-year data routing.
//
// Script Properties used:
//   ACTIVE_YEAR           → e.g. "2025"  (current live year)
//   ARCHIVE_2024          → Google Spreadsheet ID of 2024 archive
//   ARCHIVE_2023          → Google Spreadsheet ID of 2023 archive
//   FIVE_W_SHEET_ID       → Current active 5W Tracker sheet ID
//   FIVE_W_SHEET_ID_2024  → Archived 5W Tracker sheet ID for 2024
//   YEAR_END_LOCK         → "true" while transition is running (prevents double-run)
// ═══════════════════════════════════════════════════════════════════════════

// ── Tabs to CLEAR (data rows only — headers kept) ────────────────────────
var YEAR_END_CLEAR_TABS_ = [
  'Payroll Submissions',
  'Payroll Rows',
  'Payroll ROE Applications',
  'Requests Admin',
  'Requests HR',
  'Contact Requests',
  'Password Change Requests',
  'New Hires',
  'Audit Logs',
  'Employee Transfer History',
  'HotZone Editions',
  'HotZone Contributions'
];

// ── Tabs to KEEP exactly as-is (carry forward) ───────────────────────────
var YEAR_END_KEEP_TABS_ = [
  'Users',
  'Restaurants',
  'Restaurant Access',
  'HOST Links',
  'Newsletters',
  'Brands',
  'Role Button Permissions',
  'Dropdown Options',
  'Team Members',
  'HotZone Templates',
  'SavedCategories',
  'CalendarEvents'
];

// ── Pages that support year switching (read-only for archive years) ───────
var YEAR_AWARE_PAGES_ = [
  'PAYROLL_ADMIN',
  'PAYROLL_REVIEW',
  'PAYROLL_SUBMISSION',
  'RESTAURANT_REQUEST',
  'RECRUITMENT_REQUESTS',
  'OPERATIONS_DASHBOARD',
  'RECRUITMENT',
  'HOST_HOTZONE_VIEW',
  'HOTZONE_MANAGEMENT',
  'SNAPSHOT_5W'
];


// ═══════════════════════════════════════════════════════════════════════════
// PUBLIC API
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Returns available years for the year dropdown.
 * Active year is always first, then archives descending.
 */
function getAvailableYears(token) {
  try {
    var session = getSessionOrThrow_(token);
    var props   = PropertiesService.getScriptProperties().getProperties();
    var activeYear = yearEndGetActiveYear_();
    var years = [{ year: activeYear, label: String(activeYear) + ' (Current)', isActive: true, isArchive: false }];

    // Find all ARCHIVE_YYYY keys
    Object.keys(props).forEach(function(key) {
      var m = key.match(/^ARCHIVE_(\d{4})$/);
      if (m) {
        var y = parseInt(m[1], 10);
        years.push({ year: y, label: String(y) + ' (Archive)', isActive: false, isArchive: true });
      }
    });

    years.sort(function(a, b) { return b.year - a.year; });

    return respond_(true, {
      years: years,
      activeYear: activeYear,
      yearAwarePages: YEAR_AWARE_PAGES_
    });
  } catch (e) {
    console.error('getAvailableYears failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

/**
 * Main year-end transition function.
 * Called by the Year End Transition button in the Super Admin workspace.
 *
 * payload = {
 *   sessionToken:   string,
 *   fiveWSheetUrl:  string  (optional — new 5W sheet URL for next year)
 *   dryRun:         boolean (optional — simulate without making changes)
 * }
 */
function runYearEndTransition(payload) {
  payload = payload || {};
  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [APP_CONFIG.roles.SUPER_ADMIN]);

    var props      = PropertiesService.getScriptProperties();
    var activeYear = yearEndGetActiveYear_();
    var nextYear   = activeYear + 1;

    // ── Guard: already archived this year? ───────────────────────────────
    if (props.getProperty('ARCHIVE_' + activeYear)) {
      return respond_(false, null,
        'Year ' + activeYear + ' has already been archived. ' +
        'The active year is already ' + String(nextYear) + '.');
    }

    // ── Guard: prevent double-run ─────────────────────────────────────────
    if (props.getProperty('YEAR_END_LOCK') === 'true') {
      return respond_(false, null,
        'A year-end transition is already in progress. Please wait a moment and try again.');
    }

    // ── Dry run — validate only ───────────────────────────────────────────
    if (payload.dryRun) {
      return respond_(true, {
        dryRun: true,
        activeYear: activeYear,
        nextYear: nextYear,
        tabsToClear: YEAR_END_CLEAR_TABS_,
        tabsToKeep: YEAR_END_KEEP_TABS_,
        fiveWSheetUrl: payload.fiveWSheetUrl || null,
        message: 'Dry run complete. No changes made. Ready to run the real transition.'
      });
    }

    // ── Set lock ──────────────────────────────────────────────────────────
    props.setProperty('YEAR_END_LOCK', 'true');

    var log = [];
    var archiveId = null;

    try {
      // ── STEP 1: Capture HotZone published editions to carry forward ───
      log.push('Reading HotZone published editions...');
      var hotzoneCarryForward = yearEndCaptureHotzoneEditions_();
      log.push('Found ' + hotzoneCarryForward.length + ' published HotZone edition(s) to carry forward.');

      // ── STEP 2: Copy spreadsheet to archive ───────────────────────────
      log.push('Copying spreadsheet to archive "Saisethsons Portal Data ' + activeYear + '"...');
      var archiveFile = yearEndCopySpreadsheet_(activeYear);
      archiveId = archiveFile.getId();
      log.push('Archive created. ID: ' + archiveId);

      // ── STEP 3: Make archive view-only ────────────────────────────────
      log.push('Setting archive to view-only...');
      yearEndLockArchive_(archiveFile);
      log.push('Archive locked.');

      // ── STEP 4: Register archive in Script Properties ─────────────────
      log.push('Registering archive in Script Properties...');
      props.setProperty('ARCHIVE_' + activeYear, archiveId);

      // ── STEP 5: Save current 5W sheet ID to archive registry ─────────
      var currentFiveW = props.getProperty('FIVE_W_SHEET_ID') || '';
      if (currentFiveW) {
        props.setProperty('FIVE_W_SHEET_ID_' + activeYear, currentFiveW);
        log.push('5W Tracker ID archived for ' + activeYear + '.');
      }

      // ── STEP 6: Update 5W sheet if new URL provided ───────────────────
      if (payload.fiveWSheetUrl && String(payload.fiveWSheetUrl).trim()) {
        var newFiveWId = yearEndExtractSheetId_(payload.fiveWSheetUrl);
        if (!newFiveWId) throw new Error('Could not extract Sheet ID from the 5W Tracker URL. Please check the URL and try again.');
        props.setProperty('FIVE_W_SHEET_ID', newFiveWId);
        log.push('5W Tracker updated to new sheet for ' + nextYear + '. ID: ' + newFiveWId);
      } else {
        log.push('No new 5W Tracker URL provided — existing sheet ID carried forward.');
      }

      // ── STEP 7: Clear transactional tabs on active sheet ─────────────
      log.push('Clearing transactional tabs...');
      var clearResults = yearEndClearTabs_();
      clearResults.forEach(function(r) { log.push('  ' + r); });

      // ── STEP 8: Restore HotZone carry-forwards (templates already kept)
      log.push('Restoring HotZone published editions as drafts for ' + nextYear + '...');
      var restored = yearEndRestoreHotzoneEditions_(hotzoneCarryForward, nextYear, session);
      log.push('Restored ' + restored + ' HotZone edition(s) as drafts.');

      // ── STEP 9: Advance active year ───────────────────────────────────
      props.setProperty('ACTIVE_YEAR', String(nextYear));
      log.push('Active year advanced to ' + nextYear + '.');

      // ── STEP 10: Bust all caches ──────────────────────────────────────
      log.push('Clearing caches...');
      yearEndBustCaches_();
      log.push('Caches cleared.');

      // ── STEP 11: Audit log ────────────────────────────────────────────
      logAuditEvent_('YEAR_END_TRANSITION', 'SYSTEM', String(activeYear), session.userId, {
        archivedYear: activeYear,
        newActiveYear: nextYear,
        archiveId: archiveId,
        fiveWUpdated: !!(payload.fiveWSheetUrl && String(payload.fiveWSheetUrl).trim()),
        hotzoneEditionsCarriedForward: restored
      });

      log.push('✅ Year-end transition complete. ' + activeYear + ' archived. Portal is now running on ' + nextYear + '.');

    } finally {
      // Always release the lock
      props.deleteProperty('YEAR_END_LOCK');
    }

    return respond_(true, {
      success: true,
      archivedYear: activeYear,
      newActiveYear: nextYear,
      archiveId: archiveId,
      log: log
    });

  } catch (e) {
    // Release lock on error
    try { PropertiesService.getScriptProperties().deleteProperty('YEAR_END_LOCK'); } catch(le) {}
    console.error('runYearEndTransition failed', e);
    return respond_(false, null, 'Year-end transition failed: ' + sanitizeClientError_(e));
  }
}

/**
 * Returns the spreadsheet ID for a given year.
 * Active year → APP_CONFIG.spreadsheetId
 * Archive year → Script Property ARCHIVE_YYYY
 * Used by all year-aware data functions.
 */
function getSpreadsheetIdForYear_(year) {
  var activeYear = yearEndGetActiveYear_();
  if (!year || parseInt(year, 10) === activeYear) {
    return APP_CONFIG.spreadsheetId;
  }
  var archiveId = PropertiesService.getScriptProperties()
    .getProperty('ARCHIVE_' + String(year));
  if (!archiveId) throw new Error('No archive found for year ' + year + '. Available years can be checked in Script Properties.');
  return archiveId;
}

/**
 * Opens and returns the spreadsheet for a given year.
 */
function openSpreadsheetForYear_(year) {
  return SpreadsheetApp.openById(getSpreadsheetIdForYear_(year));
}

/**
 * Returns true if the given year is an archive (read-only) year.
 */
function isArchiveYear_(year) {
  if (!year) return false;
  return parseInt(year, 10) !== yearEndGetActiveYear_();
}

/**
 * Asserts that a request is not trying to write to an archive year.
 * Call this at the top of any write function that accepts a year param.
 */
function assertNotArchiveYear_(year) {
  if (isArchiveYear_(year)) {
    throw new Error('Archive data is read-only. You cannot modify ' + year + ' data.');
  }
}

/**
 * Returns the 5W Tracker sheet ID for a given year.
 */
function getFiveWSheetIdForYear_(year) {
  var activeYear = yearEndGetActiveYear_();
  if (!year || parseInt(year, 10) === activeYear) {
    return PropertiesService.getScriptProperties().getProperty('FIVE_W_SHEET_ID') || '';
  }
  return PropertiesService.getScriptProperties().getProperty('FIVE_W_SHEET_ID_' + String(year)) || '';
}


// ═══════════════════════════════════════════════════════════════════════════
// PRIVATE HELPERS
// ═══════════════════════════════════════════════════════════════════════════

function yearEndGetActiveYear_() {
  var stored = PropertiesService.getScriptProperties().getProperty('ACTIVE_YEAR');
  if (stored && /^\d{4}$/.test(stored)) return parseInt(stored, 10);
  return new Date().getFullYear();
}

function yearEndCopySpreadsheet_(activeYear) {
  var file = DriveApp.getFileById(APP_CONFIG.spreadsheetId);
  var copyName = 'Saisethsons Portal Data ' + activeYear;
  var copy = file.makeCopy(copyName);
  // Try to place it in the same parent folder as the original
  try {
    var parents = file.getParents();
    if (parents.hasNext()) {
      var folder = parents.next();
      folder.addFile(copy);
      DriveApp.getRootFolder().removeFile(copy);
    }
  } catch(e) {
    console.warn('Could not move archive to same folder — it will be in root Drive.', e);
  }
  return copy;
}

function yearEndLockArchive_(file) {
  try {
    // Remove edit access for all non-owners — make it view-only
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    console.warn('yearEndLockArchive_ sharing update failed (may need manual lock)', e);
  }
}

function yearEndClearTabs_() {
  var ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
  var results = [];

  YEAR_END_CLEAR_TABS_.forEach(function(tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      results.push('SKIP ' + tabName + ' — tab not found');
      return;
    }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      results.push('SKIP ' + tabName + ' — already empty');
      return;
    }
    // Delete all rows except header (row 1)
    sheet.deleteRows(2, lastRow - 1);
    results.push('CLEARED ' + tabName + ' (' + (lastRow - 1) + ' rows removed)');
  });

  return results;
}

function yearEndCaptureHotzoneEditions_() {
  // Read published editions + their contributions from the CURRENT active HZ sheet
  // Returns array of {edition, contributions[]} for all PUBLISHED editions
  try {
    var editions  = hzAll_('HotZone Editions').filter(function(ed) {
      return String(ed.status || '').toUpperCase() === 'PUBLISHED';
    });

    if (!editions.length) return [];

    var allContribs = hzAll_('HotZone Contributions');

    return editions.map(function(ed) {
      var contribs = allContribs.filter(function(c) {
        return String(c.edition_id) === String(ed.edition_id) &&
               String(c.status || '') !== 'DELETED';
      });
      return { edition: ed, contributions: contribs };
    });
  } catch(e) {
    console.warn('yearEndCaptureHotzoneEditions_ failed — HotZone editions not carried forward', e);
    return [];
  }
}

function yearEndRestoreHotzoneEditions_(captured, nextYear, session) {
  if (!captured || !captured.length) return 0;
  var now = new Date().toISOString();
  var restored = 0;

  captured.forEach(function(item) {
    try {
      var oldEd    = item.edition;
      var contribs = item.contributions;

      // Create new edition as DRAFT with new ID
      var newEditionId = Utilities.getUuid();
      hzAppend_('HotZone Editions', {
        edition_id:   newEditionId,
        title:        String(oldEd.title || '') + ' (Carried from ' + (nextYear - 1) + ')',
        week_ending:  '',   // blank — admin sets correct date for new year
        status:       'DRAFT',
        webapp_url:   '',
        qr_url:       '',
        created_by:   session.fullName || session.username || 'Year-End Transition',
        created_by_id: session.userId || '',
        created_at:   now,
        updated_at:   now
      });

      // Copy all contributions under new edition ID
      contribs.forEach(function(c) {
        hzAppend_('HotZone Contributions', {
          contribution_id: Utilities.getUuid(),
          edition_id:      newEditionId,
          section:         c.section || '',
          title:           c.title || '',
          body:            c.body || '',
          link_url:        c.link_url || '',
          link_title:      c.link_title || '',
          image_url:       c.image_url || '',
          image_file_id:   c.image_file_id || '',
          submitted_by:    c.submitted_by || '',
          submitted_by_id: c.submitted_by_id || '',
          submitted_at:    now,
          status:          'ACTIVE'
        });
      });

      restored++;
    } catch(e) {
      console.warn('Failed to restore HotZone edition: ' + (item.edition && item.edition.edition_id), e);
    }
  });

  return restored;
}

function yearEndExtractSheetId_(url) {
  if (!url) return null;
  url = String(url).trim();
  // Handle full URL: https://docs.google.com/spreadsheets/d/SHEET_ID/edit...
  var m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  // Handle bare ID (no slashes, reasonable length)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url)) return url;
  return null;
}

function yearEndBustCaches_() {
  try {
    var cache = CacheService.getScriptCache();
    // Remove known cache key prefixes — GAS doesn't support wildcard removes
    // so we remove the ones we know about
    var knownKeys = [
      'v500_admin_req_rows',
      'v500_hiring_rows_all',
      'seth_v325_rr_receiver_data',
      'seth_v325_admin_workspace_data',
      'seth_v325_last_updated',
      'team_recruit_cols_ok_v1',
      'role_btn_perms_v13_super_admin_all'
    ];
    cache.removeAll(knownKeys);
    // Also bump payroll admin cache version
    try { payrollAdminBumpCacheVersionV377_(); } catch(e) {}
  } catch(e) {
    console.warn('yearEndBustCaches_ partial failure', e);
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// YEAR-AWARE DATA FETCHERS
// Drop-in replacements for existing data functions.
// Each accepts an optional `year` param in the payload.
// If year === activeYear or absent → reads from APP_CONFIG.spreadsheetId
// If year is an archive year → reads from the archive spreadsheet
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Year-aware version of getAdminRequestWorkspaceData.
 * Pass payload.year to read from an archive year.
 */
function getAdminRequestWorkspaceDataForYear(payload) {
  payload = payload || {};
  var year = payload.year;

  if (!year || !isArchiveYear_(year)) {
    // Active year — use existing function
    return getAdminRequestWorkspaceData(payload);
  }

  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.RESTAURANT_USER,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    var archiveSsId = getSpreadsheetIdForYear_(year);
    var requests    = yearAwareReadSheet_(archiveSsId, 'Requests Admin');
    var restaurants = yearAwareReadSheet_(archiveSsId, 'Restaurants');

    return respond_(true, {
      year: year,
      isArchive: true,
      readOnly: true,
      restaurants: restaurants,
      employees: [],   // not needed for archive view
      requests: requests
    });
  } catch(e) {
    console.error('getAdminRequestWorkspaceDataForYear failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

/**
 * Year-aware payroll review workspace.
 */
function getPayrollReviewWorkspaceDataForYear(payload) {
  payload = payload || {};
  var year = payload.year;

  if (!year || !isArchiveYear_(year)) {
    return getPayrollReviewWorkspaceData(payload.sessionToken || payload);
  }

  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.SENIOR_MANAGEMENT
    ]);

    var archiveSsId   = getSpreadsheetIdForYear_(year);
    var submissions   = yearAwareReadSheet_(archiveSsId, 'Payroll Submissions');
    var restaurants   = yearAwareReadSheet_(archiveSsId, 'Restaurants');

    return respond_(true, {
      year: year,
      isArchive: true,
      readOnly: true,
      restaurants: restaurants,
      employees: [],
      submissions: submissions.map(enrichPayrollSubmissionV361_)
    });
  } catch(e) {
    console.error('getPayrollReviewWorkspaceDataForYear failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

/**
 * Year-aware hiring requests.
 */
function getHiringRequestWorkspaceDataForYear(payload) {
  payload = payload || {};
  var year = payload.year;

  if (!year || !isArchiveYear_(year)) {
    return getHiringRequestWorkspaceData(payload.sessionToken || payload);
  }

  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO,
      APP_CONFIG.roles.RESTAURANT_USER
    ]);

    var archiveSsId = getSpreadsheetIdForYear_(year);
    var requests    = yearAwareReadSheet_(archiveSsId, 'Requests HR');
    var restaurants = yearAwareReadSheet_(archiveSsId, 'Restaurants');

    return respond_(true, {
      year: year,
      isArchive: true,
      readOnly: true,
      restaurants: restaurants,
      requests: requests
    });
  } catch(e) {
    console.error('getHiringRequestWorkspaceDataForYear failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

/**
 * Year-aware New Heroes / Manage Team.
 */
function getManageTeamDataForYear(payload) {
  payload = payload || {};
  var year = payload.year;

  if (!year || !isArchiveYear_(year)) {
    return getManageTeamData(payload.sessionToken || payload);
  }

  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    assertAuthorized_(session, [
      APP_CONFIG.roles.SUPER_ADMIN,
      APP_CONFIG.roles.HR_ADMIN,
      APP_CONFIG.roles.MRO
    ]);

    var archiveSsId = getSpreadsheetIdForYear_(year);
    var teamMembers = yearAwareReadSheet_(archiveSsId, 'Team Members');
    var restaurants = yearAwareReadSheet_(archiveSsId, 'Restaurants');

    return respond_(true, {
      year: year,
      isArchive: true,
      readOnly: true,
      restaurants: restaurants,
      teamMembers: teamMembers,
      employees: teamMembers,
      transferHistory: []
    });
  } catch(e) {
    console.error('getManageTeamDataForYear failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

/**
 * Year-aware HotZone editions list.
 */
function getHotzoneEditionsForYear(payload) {
  payload = payload || {};
  var year = payload.year;

  if (!year || !isArchiveYear_(year)) {
    return getHotzoneEditions(payload);
  }

  try {
    var session = getSessionOrThrow_(payload.sessionToken);
    // Archive HZ data lives in the archived main spreadsheet
    var archiveSsId = getSpreadsheetIdForYear_(year);
    var editions    = yearAwareReadSheet_(archiveSsId, 'HotZone Editions');
    var contribRaw  = yearAwareReadSheet_(archiveSsId, 'HotZone Contributions');

    var enriched = editions.map(function(ed) {
      var count = contribRaw.filter(function(c) {
        return String(c.edition_id) === String(ed.edition_id) &&
               String(c.status || '') !== 'DELETED';
      }).length;
      return Object.assign({}, ed, { contrib_count: count });
    });

    return respond_(true, {
      year: year,
      isArchive: true,
      readOnly: true,
      editions: enriched,
      templates: [],
      canManage: false
    });
  } catch(e) {
    console.error('getHotzoneEditionsForYear failed', e);
    return respond_(false, null, sanitizeClientError_(e));
  }
}

/**
 * Year-aware 5W Tracker data.
 * Uses FIVE_W_SHEET_ID_YYYY for archive years.
 */
function get5WTrackerDataForYear(payload) {
  payload = payload || {};
  var year = payload.year;

  if (!year || !isArchiveYear_(year)) {
    return get5WTrackerData(payload);
  }

  // Temporarily swap the sheet ID so get5WTrackerData reads the archive
  var archiveFiveWId = getFiveWSheetIdForYear_(year);
  if (!archiveFiveWId) {
    return respond_(false, null, 'No 5W Tracker archive found for year ' + year + '.');
  }

  // We override the property temporarily for this execution
  var props = PropertiesService.getScriptProperties();
  var originalId = props.getProperty('FIVE_W_SHEET_ID');
  try {
    props.setProperty('FIVE_W_SHEET_ID', archiveFiveWId);
    var result = get5WTrackerData(payload);
    // Tag the result as archive
    if (result && result.data) {
      result.data.year = year;
      result.data.isArchive = true;
      result.data.readOnly  = true;
    }
    return result;
  } finally {
    // Always restore original
    if (originalId) props.setProperty('FIVE_W_SHEET_ID', originalId);
    else props.deleteProperty('FIVE_W_SHEET_ID');
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// UTILITY
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Generic sheet reader for an archive spreadsheet.
 * Returns array of row objects (same format as Repository.getAll_).
 */
function yearAwareReadSheet_(spreadsheetId, sheetName) {
  try {
    var ss    = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var values  = sheet.getDataRange().getValues();
    var headers = values[0].map(function(h) { return String(h || '').trim(); });

    return values.slice(1).filter(function(row) {
      return row.some(function(cell) { return cell !== ''; });
    }).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i] == null ? '' : row[i]; });
      return obj;
    });
  } catch(e) {
    console.error('yearAwareReadSheet_ failed for ' + sheetName, e);
    return [];
  }
}

/**
 * Register year-end routes in ServiceRouter_.
 * Call this once after ServiceRouter_ is defined.
 */
(function registerYearEndRoutes_() {
  try {
    if (typeof ServiceRouter_ === 'undefined' || !ServiceRouter_ || !ServiceRouter_.route) return;
    if (ServiceRouter_._yearEndRoutesRegistered) return;
    ServiceRouter_._yearEndRoutesRegistered = true;

    var original = ServiceRouter_.route;
    ServiceRouter_.route = function(action, payload, token) {
      switch (action) {
        case 'getAvailableYears':
          return getAvailableYears(token);
        case 'runYearEndTransition':
          return runYearEndTransition(payload);
        case 'getAdminRequestWorkspaceDataForYear':
          return getAdminRequestWorkspaceDataForYear(payload);
        case 'getPayrollReviewWorkspaceDataForYear':
          return getPayrollReviewWorkspaceDataForYear(payload);
        case 'getHiringRequestWorkspaceDataForYear':
          return getHiringRequestWorkspaceDataForYear(payload);
        case 'getManageTeamDataForYear':
          return getManageTeamDataForYear(payload);
        case 'getHotzoneEditionsForYear':
          return getHotzoneEditionsForYear(payload);
        case 'get5WTrackerDataForYear':
          return get5WTrackerDataForYear(payload);
        default:
          return original.call(ServiceRouter_, action, payload, token);
      }
    };
  } catch(e) {
    console.error('registerYearEndRoutes_ failed', e);
  }
})();
