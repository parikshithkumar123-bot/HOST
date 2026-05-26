/**
 * NotificationService_Patch_v2
 *
 * FIX #9: _computeUnread previously scanned all rows of 4-6 sheets on every
 * cache-miss (every 60 seconds). For large sheets this is slow and quota-expensive.
 *
 * Drop this file into your Apps Script project. It replaces the
 * NotificationService._computeUnread method defined in Services.gs by
 * re-assigning it after Services.gs has loaded.
 *
 * Changes vs original:
 * - safeRows() now accepts an optional rowLimit and stops reading after that
 *   many rows (newest-first by reading from the bottom of the sheet).
 * - CONTACT_REQUESTS and REQUESTS_HR are capped at 200 rows.
 * - PASSWORD_CHANGE_REQUESTS is capped at 100 rows.
 * - NEWSLETTERS is capped at 50 rows (newsletters are usually few).
 * - Items cap raised from 50 to 50 (unchanged) but the source reads are bounded.
 *
 * All filtering logic is identical to the original.
 */

// Re-assign _computeUnread after Services.gs defines NotificationService.
// This works because Apps Script loads all .gs files before executing any of them.
NotificationService._computeUnread = function _computeUnreadV2(session) {
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
    return ['COMPLETED', 'RESOLVED', 'CLOSED', 'REJECTED', 'DELETED', 'INACTIVE', 'DONE'].indexOf(value) === -1;
  }

  // FIX #9: Row-limited sheet reader. Reads at most `limit` rows so large
  // historical sheets don't cause slow scans on every 60-second cache miss.
  // Repository.getAll_ is replaced here with a bounded alternative.
  function safeRows(sheetName, limit) {
    try {
      const ss = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];

      const maxRows = limit && limit > 0 ? limit : 200;
      // Read from the bottom (most recent rows) up to maxRows.
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
      let payload = {};
      try { payload = JSON.parse(row.request_payload_json || '{}'); } catch (error) { console.warn('Ignored service error', error); }
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
};
