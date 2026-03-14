// ============================================================
// Cadence — Cloud Sync Google Apps Script
// Stores/retrieves task data in a Google Sheet cell.
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Saves use chunked GET (POST is broken due to GAS redirect)
// ============================================================

var SHEET_NAME = 'CadenceData';
var DATA_CELL = 'A1';        // stores full JSON blob
var TIMESTAMP_CELL = 'B1';   // last sync timestamp

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

function handleSave(dataStr) {
  try {
    if (!dataStr) return jsonOut({ok: false, error: 'No data'});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
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
    var chunk = e.parameter.c || '';
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Concatenate all chunks
    var fullData = '';
    for (var i = 0; i < total; i++) {
      fullData += (sheet.getRange('C' + (i + 1)).getValue() || '');
    }

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
