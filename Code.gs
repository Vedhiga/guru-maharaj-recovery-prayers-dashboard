/**
 * Google Apps Script — Narasimha Kavach Dashboard Backend
 *
 * Deployment:
 *   1. Create a new Google Sheet (or use existing) with headers:
 *      Date | Devotees | Chanting | Narasimha Kavach | Tulasi Parikrama | Tulasi offered
 *   2. Extensions → Apps Script → paste this file → save
 *   3. Deploy → New deployment → Web app → Execute as "Me", Access "Anyone"
 *   4. Copy the web app URL → paste into the dashboard config bar
 */

// ─── CONFIG ────────────────────────────────────────────────
const SHEET_NAME = 'Sheet1';           // tab name
const HEADER_ROW = 1;                  // row containing column names
const REQUIRED   = ['date', 'devotees'];
const ACTIVITIES  = ['chanting', 'narasimha kavach', 'tulasi parikrama', 'tulasi offered'];
// ───────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();

    if (values.length < 2) {
      return jsonResponse({ success: true, data: [], count: 0, message: 'Sheet is empty' });
    }

    const headers = values[HEADER_ROW - 1].map(h => String(h).toLowerCase().trim());
    const map = buildColumnMap(headers);

    // Validate required columns
    for (const col of REQUIRED) {
      if (map[col] === undefined) {
        return errorResponse(`Missing required column "${col}". Found: [${headers.join(', ')}]`);
      }
    }

    const rows = values.slice(HEADER_ROW).filter(r => r[map.date] && String(r[map.date]).trim());

    const data = rows.map(r => ({
      date:     normalizeDate(String(r[map.date] || '').trim()),
      devotee:  String(r[map.devotees] || '').trim(),
      chanting:  parseNum(r[map['chanting']]),
      kavach:    parseNum(r[map['narasimha kavach']]),
      parikrama: parseNum(r[map['tulasi parikrama']]),
      tulasi:    parseNum(r[map['tulasi offered']]),
    })).filter(r => r.devotee && r.date);

    return jsonResponse({ success: true, data, count: data.length });
  } catch (err) {
    return errorResponse(err.message);
  }
}

// ─── HELPERS ───────────────────────────────────────────────

function buildColumnMap(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (/date/.test(h))                 map['date'] = i;
    if (/devotee|devotees|name/.test(h)) map['devotees'] = i;
    if (/chanting/.test(h))             map['chanting'] = i;
    if (/narasimha.*kavach|kavach/.test(h)) map['narasimha kavach'] = i;
    if (/tulasi.*parikrama|parikrama/.test(h)) map['tulasi parikrama'] = i;
    if (/tulasi.*offer|tulasi/.test(h)) map['tulasi offered'] = i;
  }
  return map;
}

function normalizeDate(d) {
  const parsed = new Date(d);
  return isNaN(parsed) ? d : parsed.toISOString().slice(0, 10);
}

function parseNum(v) {
  const n = Number(String(v).replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? 0 : Math.round(n);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(msg) {
  return jsonResponse({ success: false, error: msg });
}
