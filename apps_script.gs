// ============================================================
// Cadence — Cloud Sync Google Apps Script
// Stores/retrieves task data in a Google Sheet cell.
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Supports JSONP (callback param) for cross-origin loads
// ============================================================

var SHEET_NAME = 'CadenceData';
var DATA_CELL = 'A1';        // stores full JSON blob
var TIMESTAMP_CELL = 'B1';   // last sync timestamp

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  var callback = e.parameter.callback || '';

  if (action === 'save') {
    return handleSave(e, callback);
  }
  if (action === 'load') {
    return handleLoad(e, callback);
  }

  return wrapResponse({ok: false, error: 'Unknown action'}, callback);
}

function wrapResponse(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSave(e, callback) {
  try {
    var dataStr = e.parameter.data || '';
    if (!dataStr) {
      return wrapResponse({ok: false, error: 'No data'}, callback);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    sheet.getRange(DATA_CELL).setValue(dataStr);
    sheet.getRange(TIMESTAMP_CELL).setValue(new Date().toISOString());

    return wrapResponse({ok: true, saved: true, ts: new Date().toISOString()}, callback);
  } catch (err) {
    return wrapResponse({ok: false, error: err.message}, callback);
  }
}

function handleLoad(e, callback) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return wrapResponse({ok: true, data: null, ts: null}, callback);
    }

    var dataStr = sheet.getRange(DATA_CELL).getValue();
    var ts = sheet.getRange(TIMESTAMP_CELL).getValue();

    return wrapResponse({ok: true, data: dataStr || null, ts: ts || null}, callback);
  } catch (err) {
    return wrapResponse({ok: false, error: err.message}, callback);
  }
}
