// ════════════════════════════════════════════════════════════
//  PBI Label Generator — Google Apps Script Backend
//  Sheet: "PBI Labels"  |  Tab: "LOG"
//
//  SETUP:
//  1. Open your Google Sheet → Extensions → Apps Script
//  2. Paste this entire file into Code.gs
//  3. Deploy → New Deployment → Web App
//       Execute as: Me
//       Who has access: Anyone
//  4. Copy the Web App URL → paste into index.html GAS_URL
// ════════════════════════════════════════════════════════════

const SHEET_NAME = 'LOG';

// ── Column headers (written once on first run) ──────────────
const HEADERS = [
  'Timestamp',
  'Customer',
  'Material',
  'Size',
  'Total Qty',
  'Date on Label',
  'Pages Generated',
  'Labels Per Page',
];

// ── GET — health check / CORS preflight ────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'PBI Label API is live.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST — main entry point ────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'saveLabel') {
      return handleSaveLabel(payload.data);
    }

    return respond({ status: 'error', message: 'Unknown action: ' + payload.action });

  } catch (err) {
    return respond({ status: 'error', message: err.message });
  }
}

// ── SAVE LABEL RECORD ──────────────────────────────────────
function handleSaveLabel(data) {
  // Validate required fields
  const required = ['customer', 'material', 'size', 'qty', 'date'];
  for (const key of required) {
    if (!data[key]) {
      return respond({ status: 'error', message: 'Missing field: ' + key });
    }
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAME);

  // Write headers if the sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    // Bold + freeze header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0d1b4b');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Auto-resize columns
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  const qty   = parseInt(data.qty) || 0;
  const pages = Math.ceil(qty / 12);

  const row = [
    new Date(),            // Timestamp
    data.customer,         // Customer
    data.material,         // Material
    data.size,             // Size
    qty,                   // Total Qty
    data.date,             // Date on Label
    pages,                 // Pages Generated
    12,                    // Labels Per Page
  ];

  sheet.appendRow(row);

  // Format the timestamp cell in the new row
  const newRowNum = sheet.getLastRow();
  sheet.getRange(newRowNum, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Alternating row color for readability
  const rowRange = sheet.getRange(newRowNum, 1, 1, HEADERS.length);
  if (newRowNum % 2 === 0) {
    rowRange.setBackground('#f0f4ff');
  }

  return respond({
    status: 'ok',
    message: 'Label record saved.',
    row: newRowNum,
    data: { customer: data.customer, qty, pages }
  });
}

// ── HELPERS ───────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
