// ═══════════════════════════════════════════════════════════════════════════
//  CWC Exhibition — Apps Script Proxy
//  Paste this entire file into your Google Sheet's Apps Script editor
//  (Extensions → Apps Script), then deploy as a Web App.
//
//  Deploy settings:
//    Execute as:      Me  (your Google account — gives it sheet access)
//    Who has access:  Anyone  (the HTML app calls it without login)
// ═══════════════════════════════════════════════════════════════════════════

// Column name → score key mapping
const SCORE_KEYS = [
  'Selector1Score',
  'Selector2Score',
  'Selector3Score',
  'Selector4Score',
];

// ── Entry point for GET requests (ping + readSheet) ──────────────────────
function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'ping') {
    return jsonResponse({ status: 'ok', message: 'CWC proxy is running' });
  }

  if (action === 'readSheet') {
    const sheetId = e.parameter.sheetId || '';
    if (!sheetId) return jsonResponse({ status: 'error', error: 'Missing sheetId' });
    try {
      return jsonResponse(readSheetData(sheetId));
    } catch(err) {
      return jsonResponse({ status: 'error', error: err.message });
    }
  }

  return jsonResponse({ status: 'error', error: 'Unknown action: ' + action });
}

// ── Entry point for POST requests (writeVote) ────────────────────────────
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch(err) {
    return jsonResponse({ status: 'error', error: 'Invalid JSON body' });
  }

  const action = body.action || '';

  if (action === 'writeVote') {
    const { sheetId, rowIndex, selectorKey, value } = body;

    // Validate
    if (!sheetId)     return jsonResponse({ status:'error', error:'Missing sheetId' });
    if (!rowIndex)    return jsonResponse({ status:'error', error:'Missing rowIndex' });
    if (!selectorKey) return jsonResponse({ status:'error', error:'Missing selectorKey' });
    if (!SCORE_KEYS.includes(selectorKey))
      return jsonResponse({ status:'error', error:'Invalid selectorKey: ' + selectorKey });
    if (value !== 0 && value !== 1 && value !== 2)
      return jsonResponse({ status:'error', error:'value must be 0, 1 or 2' });

    try {
      writeVote(sheetId, Number(rowIndex), selectorKey, Number(value));
      return jsonResponse({ status: 'ok' });
    } catch(err) {
      return jsonResponse({ status: 'error', error: err.message });
    }
  }

  return jsonResponse({ status: 'error', error: 'Unknown action: ' + action });
}

// ── Read the Form Responses + Selectors tabs ─────────────────────────────
function readSheetData(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);

  // ── Form Responses ──
  const formSheet = ss.getSheetByName('Form Responses');
  if (!formSheet) throw new Error('Sheet "Form Responses" not found');
  const formData = formSheet.getDataRange().getValues();
  if (formData.length < 2) return { status:'ok', headers:[], rows:[], selectors:[] };

  const headers = formData[0].map(String);
  const rows    = formData.slice(1).map(row => row.map(v => (v === null || v === undefined) ? '' : String(v)));

  // ── Selectors ──
  let selectors = [];
  const selSheet = ss.getSheetByName('Selectors');
  if (selSheet) {
    const selData = selSheet.getDataRange().getValues();
    if (selData.length > 1) {
      const sh  = selData[0].map(String);
      const idx = k => sh.indexOf(k);
      selectors = selData.slice(1).map(row => {
        const num = parseInt(row[idx('SelectorNumber')]) || 0;
        return {
          number:   num,
          name:     String(row[idx('Name')]     || ''),
          email:    String(row[idx('Email')]    || '').toLowerCase().trim(),
          password: String(row[idx('Password')] || ''),
          key:      'Selector' + num + 'Score',
        };
      }).filter(s => s.number > 0);
    }
  }

  return { status:'ok', headers, rows, selectors };
}

// ── Write a single vote into the correct cell ────────────────────────────
function writeVote(sheetId, rowIndex, selectorKey, value) {
  const ss        = SpreadsheetApp.openById(sheetId);
  const sheet     = ss.getSheetByName('Form Responses');
  if (!sheet) throw new Error('Sheet "Form Responses" not found');

  // Find column index of selectorKey in header row
  const headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex  = headers.indexOf(selectorKey); // 0-based
  if (colIndex < 0) throw new Error('Column "' + selectorKey + '" not found in header row');

  // Write value (rowIndex is 1-based sheet row; col is 1-based)
  sheet.getRange(rowIndex, colIndex + 1).setValue(value);

  // Flush to ensure it's written immediately (important for concurrent selectors)
  SpreadsheetApp.flush();
}

// ── Helper: return JSON with CORS headers ────────────────────────────────
function jsonResponse(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
