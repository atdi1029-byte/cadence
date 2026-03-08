// ============================================================
// Cadence — Cloud Sync Google Apps Script
// Stores/retrieves task data in a Google Sheet cell.
// Deploy as Web App: Execute as "Me", Access "Anyone"
// ============================================================

var SHEET_NAME = 'CadenceData';
var DATA_CELL = 'A1';        // stores full JSON blob
var TIMESTAMP_CELL = 'B1';   // last sync timestamp

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();

  if (action === 'save') {
    return handleSave(e);
  }
  if (action === 'load') {
    return handleLoad(e);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ok: false, error: 'Unknown action'})
  ).setMimeType(ContentService.MimeType.JSON);
}

function handleSave(e) {
  try {
    var dataStr = e.parameter.data || '';
    if (!dataStr) {
      return ContentService.createTextOutput(
        JSON.stringify({ok: false, error: 'No data'})
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    sheet.getRange(DATA_CELL).setValue(dataStr);
    sheet.getRange(TIMESTAMP_CELL).setValue(new Date().toISOString());

    return ContentService.createTextOutput(
      JSON.stringify({ok: true, saved: true, ts: new Date().toISOString()})
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ok: false, error: err.message})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleLoad(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ok: true, data: null, ts: null})
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var dataStr = sheet.getRange(DATA_CELL).getValue();
    var ts = sheet.getRange(TIMESTAMP_CELL).getValue();

    return ContentService.createTextOutput(
      JSON.stringify({ok: true, data: dataStr || null, ts: ts || null})
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ok: false, error: err.message})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
