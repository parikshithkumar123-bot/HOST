// FIX #11: logAuditEvent_ is now NON-BLOCKING.
// Instead of writing synchronously to the sheet on every call
// (which acquires the script lock and adds 500ms–2s latency),
// we buffer the record into ScriptCache and let flushAuditBuffer()
// drain it via a 1-minute time-based trigger.
//
// HOW TO SET UP THE FLUSH TRIGGER — run once in the script editor:
//   ScriptApp.newTrigger('flushAuditBuffer').timeBased().everyMinutes(1).create();
//
// FIX (audit completeness): actor_name and role_code are now populated
// from a session object when one is passed as an optional 6th argument.
// Removed logAuditEvent_: removed duplicate, kept last version.

// FIX #11: Flush buffered audit records to the Audit Logs sheet.
// This function is called by a 1-minute time-based trigger.
// It reads all buffered keys from ScriptCache and batch-appends them.
//
// NOTE: ScriptCache does not support prefix-scanning natively.
// We maintain a small "index" list of pending keys in a separate cache entry.
// Because of this limitation, logAuditEvent_ also pushes its key into the index.
//
// SIMPLIFIED APPROACH: since ScriptCache doesn't support listing keys,
// we use a size-bounded queue stored in a single JSON array under a fixed key.
// The queue holds up to 200 serialized records. The flush trigger drains it.

function logAuditEvent_(eventType, entityType, entityId, actorUserId, metadata, session) {
  try {
    const record = {
      audit_id: Utilities.getUuid(),
      event_type: String(eventType || ''),
      entity_type: String(entityType || ''),
      entity_id: String(entityId || ''),
      action: String(eventType || ''),
      actor_user_id: String(actorUserId || ''),
      actor_name: (session && session.fullName) ? String(session.fullName) : '',
      role_code: (session && session.roleCode) ? String(session.roleCode) : '',
      restaurant_id: (session && session.restaurantId) ? String(session.restaurantId) : '',
      metadata_json: JSON.stringify(metadata || {}),
      ip_address: '',
      created_at: new Date().toISOString()
    };

    const cache = CacheService.getScriptCache();
    const queueKey = 'audit_queue_v1';
    const lock = LockService.getScriptLock();

    try {
      lock.waitLock(3000);
    } catch (lockErr) {
      // If we can't get the lock quickly, just do a direct write to avoid losing the event.
      Repository.append_(APP_CONFIG.sheets.AUDIT_LOGS, record);
      return;
    }

    try {
      let queue = [];
      const existing = cache.get(queueKey);
      if (existing) {
        try { queue = JSON.parse(existing); } catch (e) { queue = []; }
      }
      queue.push(record);
      // Cap the queue at 200 to avoid cache size limits (~90KB max).
      // If over the cap, flush inline to avoid losing records.
      if (queue.length > 200) {
        // FIX #13: Never call flushAuditBuffer() while holding any lock — causes deadlock
        // because flushAuditBuffer() also tries to acquire the same script lock.
        // Instead, signal urgency via a separate cache key and let the trigger handle it.
        lock.releaseLock();
        try { CacheService.getScriptCache().put('audit_flush_urgent', '1', 120); } catch (e) {}
        // Still persist the over-limit record directly to avoid loss
        Repository.append_(APP_CONFIG.sheets.AUDIT_LOGS, record);
        return;
      }
      const serialized = JSON.stringify(queue);
      if (serialized.length > 85000) {
        // Approaching cache size limit — signal urgent flush without calling it inline.
        lock.releaseLock();
        try { CacheService.getScriptCache().put('audit_flush_urgent', '1', 120); } catch (e) {}
        Repository.append_(APP_CONFIG.sheets.AUDIT_LOGS, record);
        return;
      }
      cache.put(queueKey, serialized, 600);
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }
  } catch (error) {
    // Ultimate fallback: direct synchronous write.
    console.error('Audit queue write failed — direct write fallback', error);
    try {
      Repository.append_(APP_CONFIG.sheets.AUDIT_LOGS, {
        audit_id: Utilities.getUuid(),
        event_type: String(eventType || ''),
        entity_type: String(entityType || ''),
        entity_id: String(entityId || ''),
        action: String(eventType || ''),
        actor_user_id: String(actorUserId || ''),
        actor_name: '',
        role_code: '',
        restaurant_id: '',
        metadata_json: JSON.stringify(metadata || {}),
        ip_address: '',
        created_at: new Date().toISOString()
      });
    } catch (fallbackErr) {
      console.error('Audit direct write also failed', fallbackErr);
    }
  }
}

// Called by a 1-minute time-based trigger (set up once via setupAuditFlushTrigger).
function flushAuditBuffer() {
  const cache = CacheService.getScriptCache();
  const queueKey = 'audit_queue_v1';
  const lock = LockService.getScriptLock();

  // FIX #13: Clear the urgent flush flag so logAuditEvent_ knows we're flushing.
  try { CacheService.getScriptCache().remove('audit_flush_urgent'); } catch (e) {}

  try {
    lock.waitLock(15000);
  } catch (e) {
    console.warn('flushAuditBuffer: could not acquire lock', e);
    return;
  }

  try {
    const raw = cache.get(queueKey);
    if (!raw) return; // Nothing to flush.

    let queue = [];
    try { queue = JSON.parse(raw); } catch (e) { return; }
    if (!queue.length) return;

    // Clear the queue immediately under lock so concurrent executions don't double-write.
    cache.remove(queueKey);
    lock.releaseLock();

    // Batch-append all queued records.
    const sheet = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId)
      .getSheetByName(APP_CONFIG.sheets.AUDIT_LOGS);
    if (!sheet) {
      console.error('flushAuditBuffer: Audit Logs sheet not found');
      return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });

    const rows = queue.map(function(record) {
      return headers.map(function(header) {
        return header in record ? record[header] : '';
      });
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }

    console.log('flushAuditBuffer: flushed ' + rows.length + ' audit record(s)');
  } catch (e) {
    console.error('flushAuditBuffer failed', e);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Run this ONCE from the script editor to create the 1-minute flush trigger.
function setupAuditFlushTrigger() {
  // Remove any existing audit flush triggers first.
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'flushAuditBuffer') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('flushAuditBuffer')
    .timeBased()
    .everyMinutes(1)
    .create();
  return 'Audit flush trigger created. flushAuditBuffer will run every minute.';
}
