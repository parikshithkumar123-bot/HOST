const Repository = {
  getSheet_(sheetName) {
    const cacheKey = String(sheetName || '');
    Repository.__sheetCacheV324_ = Repository.__sheetCacheV324_ || {};
    if (Repository.__sheetCacheV324_[cacheKey]) {
      return Repository.__sheetCacheV324_[cacheKey];
    }

    // Cache the spreadsheet object — openById is expensive, only call once per execution
    if (!Repository.__spreadsheetV326_) {
      Repository.__spreadsheetV326_ = SpreadsheetApp.openById(APP_CONFIG.spreadsheetId);
    }
    const spreadsheet = Repository.__spreadsheetV326_;

    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      // Sheet doesn't exist — create it with headers
      const schemas = getSheetSchemas_();
      const headers = schemas[sheetName];
      if (!headers) throw new Error(`Missing sheet: ${sheetName}`);
      sheet = spreadsheet.insertSheet(sheetName);
      upsertHeaderRow_(sheet, headers);
      sheet.setFrozenRows(1);
    }
    // Sheet already exists — trust it, no header validation on every read
    Repository.__sheetCacheV324_[cacheKey] = sheet;
    return sheet;
  },

  getHeaders_(sheetName) {
    const cacheKey = String(sheetName || '');
    Repository.__headerCacheV324_ = Repository.__headerCacheV324_ || {};
    if (Repository.__headerCacheV324_[cacheKey]) {
      return Repository.__headerCacheV324_[cacheKey].slice();
    }

    const sheet = this.getSheet_(sheetName);
    if (sheet.getLastColumn() < 1) return [];

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
      return String(header || '').trim();
    });

    Repository.__headerCacheV324_[cacheKey] = headers;
    return headers.slice();
  },

  getAll_(sheetName) {
    const sheet = this.getSheet_(sheetName);
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];
    const headers = values[0];
    return values.slice(1).filter(function(row) {
      return row.some(function(cell) { return cell !== ''; });
    }).map(function(row) {
      return rowToObject_(headers, row);
    });
  },

  findOneBy_(sheetName, columnName, expectedValue) {
    const records = this.getAll_(sheetName);
    return records.find(function(record) {
      return String(record[columnName]) === String(expectedValue);
    }) || null;
  },

  append_(sheetName, record) {
    return withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const row = headers.map(function(header) {
        return header in record ? record[header] : '';
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
      return record;
    });
  },

  updateById_(sheetName, idColumn, idValue, patch) {
    return withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const values = sheet.getDataRange().getValues();
      const headers = values[0].map(function(h) { return String(h || '').trim(); });
      const idIndex = headers.indexOf(idColumn);
      if (idIndex === -1) throw new Error(`Column ${idColumn} not found in ${sheetName}`);
      for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
        if (String(values[rowIndex][idIndex]) === String(idValue)) {
          headers.forEach(function(header, headerIndex) {
            if (header in patch) {
              values[rowIndex][headerIndex] = patch[header];
            }
          });
          sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values[rowIndex]]);
          return true;
        }
      }
      throw new Error(`Record not found in ${sheetName}: ${idValue}`);
    });
  },

  deleteById_(sheetName, idColumn, idValue) {
    return withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const idIndex = headers.indexOf(idColumn);
      if (idIndex === -1) throw new Error(`Missing ID column ${idColumn} in ${sheetName}`);

      for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
        if (String(values[rowIndex][idIndex]) === String(idValue)) {
          sheet.deleteRow(rowIndex + 1);
          return true;
        }
      }
      throw new Error(`Record not found in ${sheetName}: ${idValue}`);
    });
  },

  replaceRowsByColumn_(sheetName, columnName, expectedValue, records) {
    return withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const columnIndex = headers.indexOf(columnName);
      if (columnIndex === -1) throw new Error('Missing column ' + columnName + ' in ' + sheetName);

      // FIX #5: Collect all row indices to delete first, then delete bottom-up in one pass.
      // Deleting bottom-up avoids index-shift bugs. We no longer delete inside the scan loop
      // which could skip rows when indices shift mid-iteration.
      const rowsToDelete = [];
      for (let i = values.length - 1; i >= 1; i--) {
        if (String(values[i][columnIndex]) === String(expectedValue)) {
          rowsToDelete.push(i + 1); // already in descending order
        }
      }
      for (let di = 0; di < rowsToDelete.length; di++) {
        sheet.deleteRow(rowsToDelete[di]);
      }

      if (records && records.length) {
        const rows = records.map(function(record) {
          return headers.map(function(header) {
            return header in record ? record[header] : '';
          });
        });
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
      }

      return true;
    });
  }
};

// FIX #8: Request-scoped in-memory cache.
// Lives only for the duration of one Apps Script execution (one HTTP request).
// Use Repository.getAllCached_(sheetName) inside service functions that call
// multiple sub-functions — each sub-function gets the same in-memory array
// without re-reading the sheet or ScriptCache.
// The cache is NOT cleared on writes (it's per-request), so only use it for
// read-only aggregation within a single function call chain.
Repository.__requestCache_ = Repository.__requestCache_ || {};

Repository.getAllCached_ = function(sheetName) {
  const key = String(sheetName || '');
  if (!Repository.__requestCache_[key]) {
    Repository.__requestCache_[key] = Repository.getAll_(sheetName);
  }
  return Repository.__requestCache_[key];
};

function rowToObject_(headers, row) {
  const object = {};
  headers.forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key) object[key] = row[index];
  });
  return object;
}

function withSheetLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error('The system is busy. Please try again in a moment.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}


/*******************************************************
 V257 BACKEND SAVE CLEANUP
 Extra Repository helpers for safer and clearer saves.
 These helpers do not delete old data and do not change login.
*******************************************************/

(function addRepositoryHelpersV257_() {
  if (Repository.updateByIdWithResult_) return;

  Repository.updateByIdWithResult_ = function(sheetName, idColumn, idValue, patch) {
    return withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const values = sheet.getDataRange().getValues();
      if (values.length < 1) throw new Error('Sheet has no header row: ' + sheetName);

      const headers = values[0].map(function(h) { return String(h || '').trim(); });
      const idIndex = headers.indexOf(idColumn);
      if (idIndex === -1) throw new Error('Column ' + idColumn + ' not found in ' + sheetName);

      for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
        if (String(values[rowIndex][idIndex]) !== String(idValue)) continue;

        const oldRecord = rowToObject_(headers, values[rowIndex]);
        headers.forEach(function(header, headerIndex) {
          if (header in patch) values[rowIndex][headerIndex] = patch[header];
        });

        sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([values[rowIndex]]);
        const newRecord = rowToObject_(headers, values[rowIndex]);

        return {
          rowNumber: rowIndex + 1,
          oldRecord: oldRecord,
          newRecord: newRecord
        };
      }

      throw new Error('Record not found in ' + sheetName + ': ' + idValue);
    });
  };

  Repository.appendWithResult_ = function(sheetName, record) {
    return withSheetLock_(function() {
      const sheet = Repository.getSheet_(sheetName);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      const row = headers.map(function(header) {
        return header in record ? record[header] : '';
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
      return {
        rowNumber: sheet.getLastRow(),
        record: record
      };
    });
  };

  Repository.upsertByIdWithResult_ = function(sheetName, idColumn, idValue, record) {
    const existing = Repository.findOneBy_(sheetName, idColumn, idValue);
    if (existing) {
      const updateResult = Repository.updateByIdWithResult_(sheetName, idColumn, idValue, record);
      return {
        action: 'UPDATED',
        oldRecord: updateResult.oldRecord,
        newRecord: updateResult.newRecord,
        rowNumber: updateResult.rowNumber
      };
    }

    const appendResult = Repository.appendWithResult_(sheetName, record);
    return {
      action: 'CREATED',
      oldRecord: null,
      newRecord: record,
      rowNumber: appendResult.rowNumber
    };
  };
})();


/*******************************************************
 V258 SPEED AND STABILITY CLEANUP
 Small cache for repeated sheet reads.
 Cache is cleared automatically after writes.
*******************************************************/
(function addRepositorySpeedHelpersV258_() {
  if (Repository.__v258SpeedReady) return;
  Repository.__v258SpeedReady = true;

  Repository.cacheKeyV258_ = function(sheetName) {
    return 'repo_v258_' + String(APP_CONFIG.spreadsheetId || '').slice(-8) + '_' + String(sheetName || '').replace(/[^A-Za-z0-9_]/g, '_');
  };

  Repository.clearCacheV258_ = function(sheetName) {
    try {
      const cache = CacheService.getScriptCache();
      if (sheetName) cache.remove(Repository.cacheKeyV258_(sheetName));
    } catch (error) {}
  };

  Repository.clearKnownCachesV258_ = function() {
    try {
      const cache = CacheService.getScriptCache();
      const sheets = Object.keys(APP_CONFIG.sheets || {}).map(function(key) { return APP_CONFIG.sheets[key]; });
      const keys = sheets.map(function(name) { return Repository.cacheKeyV258_(name); });
      if (keys.length) cache.removeAll(keys);
    } catch (error) {}
  };

  const originalGetAll = Repository.getAll_;
  Repository.getAll_ = function(sheetName) {
    const key = Repository.cacheKeyV258_(sheetName);
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(key);
      if (cached) return JSON.parse(cached);
      const records = originalGetAll.call(Repository, sheetName);
      const json = JSON.stringify(records);
      if (json.length < 90000) cache.put(key, json, 45);
      return records;
    } catch (error) {
      return originalGetAll.call(Repository, sheetName);
    }
  };

  function clearAfterWrite(sheetName) {
    Repository.clearCacheV258_(sheetName);
    try {
      if (Repository.__sheetCacheV324_) delete Repository.__sheetCacheV324_[String(sheetName || '')];
      if (Repository.__headerCacheV324_) delete Repository.__headerCacheV324_[String(sheetName || '')];
    } catch (error) {}
    // FIX v500 #2: Also clear __requestCache_ so a read→write→read sequence
    // within the same request sees fresh data, not the stale pre-write snapshot.
    try {
      if (Repository.__requestCache_ && sheetName) {
        delete Repository.__requestCache_[String(sheetName)];
      }
    } catch (error) {}
  }

  const originalAppend = Repository.append_;
  Repository.append_ = function(sheetName, record) {
    const result = originalAppend.call(Repository, sheetName, record);
    clearAfterWrite(sheetName);
    return result;
  };

  const originalUpdateById = Repository.updateById_;
  Repository.updateById_ = function(sheetName, idColumn, idValue, patch) {
    const result = originalUpdateById.call(Repository, sheetName, idColumn, idValue, patch);
    clearAfterWrite(sheetName);
    return result;
  };

  const originalDeleteById = Repository.deleteById_;
  Repository.deleteById_ = function(sheetName, idColumn, idValue) {
    const result = originalDeleteById.call(Repository, sheetName, idColumn, idValue);
    clearAfterWrite(sheetName);
    return result;
  };

  const originalReplaceRowsByColumn = Repository.replaceRowsByColumn_;
  Repository.replaceRowsByColumn_ = function(sheetName, columnName, expectedValue, records) {
    const result = originalReplaceRowsByColumn.call(Repository, sheetName, columnName, expectedValue, records);
    clearAfterWrite(sheetName);
    return result;
  };

  if (Repository.updateByIdWithResult_) {
    const originalUpdateByIdWithResult = Repository.updateByIdWithResult_;
    Repository.updateByIdWithResult_ = function(sheetName, idColumn, idValue, patch) {
      const result = originalUpdateByIdWithResult.call(Repository, sheetName, idColumn, idValue, patch);
      clearAfterWrite(sheetName);
      return result;
    };
  }

  if (Repository.appendWithResult_) {
    const originalAppendWithResult = Repository.appendWithResult_;
    Repository.appendWithResult_ = function(sheetName, record) {
      const result = originalAppendWithResult.call(Repository, sheetName, record);
      clearAfterWrite(sheetName);
      return result;
    };
  }
})();
