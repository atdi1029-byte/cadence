// ============================================================
// Cadence — Cloud Sync Google Apps Script
// Stores/retrieves task data in a Google Sheet cell.
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Saves use chunked GET (POST is broken due to GAS redirect)
// ============================================================

var SHEET_NAME = 'CadenceData';
var DATA_CELL = 'A1';        // stores full JSON blob
var TIMESTAMP_CELL = 'B1';   // last sync timestamp
var BACKUP_CELL = 'D1';      // backup of last good save
var BACKUP_TS_CELL = 'E1';   // backup timestamp

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  var callback = e.parameter.callback || '';

  var result;
  if (action === 'load') {
    result = handleLoad();
  } else if (action === 'save') {
    result = handleSave(e.parameter.data || '');
  } else if (action === 'save_chunk') {
    result = handleSaveChunk(e);
  } else if (action === 'save_done') {
    result = handleSaveDone(e);
  } else if (action === 'restore_backup') {
    result = handleRestoreBackup();
  } else {
    result = jsonOut({ok: false, error: 'Unknown action'});
  }

  // JSONP support — wrap in callback for CORS-free loading
  if (callback) {
    var json = result.getContent();
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return result;
}

function doPost(e) {
  try {
    var body = e.postData ? e.postData.contents : '';
    var dataStr = '';
    // Handle text/plain JSON body: { action: 'save', data: '...' }
    try {
      var parsed = JSON.parse(body);
      if (parsed.action === 'save' && parsed.data) {
        dataStr = parsed.data;
      } else {
        dataStr = body;
      }
    } catch(pe) {
      // Fallback: raw body or form-encoded
      dataStr = e.parameter && e.parameter.data ? e.parameter.data : body;
    }
    return handleSave(dataStr);
  } catch (err) {
    return jsonOut({ok: false, error: err.message});
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Validate data is parseable JSON with a tasks array
function validateData(dataStr) {
  if (!dataStr || typeof dataStr !== 'string') return false;
  try {
    var obj = JSON.parse(dataStr);
    return obj && Array.isArray(obj.tasks);
  } catch(e) {
    return false;
  }
}

// Backup current A1 to D1 before overwriting
function backupCurrent(sheet) {
  var current = sheet.getRange(DATA_CELL).getValue();
  if (current && validateData(current)) {
    sheet.getRange(BACKUP_CELL).setValue(current);
    sheet.getRange(BACKUP_TS_CELL).setValue(
      sheet.getRange(TIMESTAMP_CELL).getValue() || new Date().toISOString()
    );
  }
}

function handleSave(dataStr) {
  try {
    if (!dataStr) return jsonOut({ok: false, error: 'No data'});
    if (!validateData(dataStr)) {
      return jsonOut({ok: false, error: 'Invalid JSON — save rejected'});
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    backupCurrent(sheet);
    sheet.getRange(DATA_CELL).setValue(dataStr);
    sheet.getRange(TIMESTAMP_CELL).setValue(new Date().toISOString());
    return jsonOut({ok: true, saved: true, ts: new Date().toISOString()});
  } catch (err) {
    return jsonOut({ok: false, error: err.message});
  }
}

// Store one chunk of data (used for large payloads that don't fit in a single GET URL)
function handleSaveChunk(e) {
  try {
    var idx = parseInt(e.parameter.i || '0');
    var chunk = e.parameter.cd || '';
    if (!chunk) return jsonOut({ok: false, error: 'No chunk data'});

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    // Store chunk in column C (C1=chunk0, C2=chunk1, etc.)
    sheet.getRange('C' + (idx + 1)).setValue(chunk);
    return jsonOut({ok: true, chunk: idx});
  } catch (err) {
    return jsonOut({ok: false, error: err.message});
  }
}

// Assemble all chunks into the final data blob
function handleSaveDone(e) {
  try {
    var total = parseInt(e.parameter.n || '1');
    var expectedLen = parseInt(e.parameter.len || '0');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Concatenate all chunks
    var fullData = '';
    for (var i = 0; i < total; i++) {
      fullData += (sheet.getRange('C' + (i + 1)).getValue() || '');
    }

    // Integrity check: verify length matches what client sent
    if (expectedLen > 0 && fullData.length !== expectedLen) {
      // Clean up chunks but don't save corrupt data
      for (var i = 0; i < total; i++) {
        sheet.getRange('C' + (i + 1)).clearContent();
      }
      return jsonOut({
        ok: false,
        error: 'Length mismatch: expected ' + expectedLen + ', got ' + fullData.length
      });
    }

    // Validate assembled data is valid JSON with tasks
    if (!validateData(fullData)) {
      // Clean up chunks but don't save corrupt data
      for (var i = 0; i < total; i++) {
        sheet.getRange('C' + (i + 1)).clearContent();
      }
      return jsonOut({ok: false, error: 'Assembled data is not valid JSON — save rejected'});
    }

    // Backup current good data before overwriting
    backupCurrent(sheet);

    // Save assembled data
    sheet.getRange(DATA_CELL).setValue(fullData);
    sheet.getRange(TIMESTAMP_CELL).setValue(new Date().toISOString());

    // Clean up chunk cells
    for (var i = 0; i < total; i++) {
      sheet.getRange('C' + (i + 1)).clearContent();
    }

    return jsonOut({ok: true, saved: true, size: fullData.length, ts: new Date().toISOString()});
  } catch (err) {
    return jsonOut({ok: false, error: err.message});
  }
}

// Emergency restore: copy backup (D1) back to A1
function handleRestoreBackup() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonOut({ok: false, error: 'No sheet'});

    var backup = sheet.getRange(BACKUP_CELL).getValue();
    if (!backup) return jsonOut({ok: false, error: 'No backup exists'});
    if (!validateData(backup)) return jsonOut({ok: false, error: 'Backup is also corrupt'});

    sheet.getRange(DATA_CELL).setValue(backup);
    sheet.getRange(TIMESTAMP_CELL).setValue(new Date().toISOString());
    return jsonOut({ok: true, restored: true, ts: new Date().toISOString()});
  } catch (err) {
    return jsonOut({ok: false, error: err.message});
  }
}

function handleLoad() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonOut({ok: true, data: null, ts: null});

    var dataStr = sheet.getRange(DATA_CELL).getValue();
    var ts = sheet.getRange(TIMESTAMP_CELL).getValue();
    return jsonOut({ok: true, data: dataStr || null, ts: ts || null});
  } catch (err) {
    return jsonOut({ok: false, error: err.message});
  }
}
