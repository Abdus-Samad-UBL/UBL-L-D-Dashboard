/**
 * ═══════════════════════════════════════════════════════════════
 * UBL L&D EXECUTIVE DASHBOARD — Google Apps Script Backend
 * ═══════════════════════════════════════════════════════════════
 *
 * SETUP INSTRUCTIONS:
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Paste this entire file into the editor
 * 3. Click "Deploy" → "New deployment" → Type: "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone (or Anyone within your organisation)
 * 4. Copy the Web App URL
 * 5. Paste the URL into your dashboard.html as:
 *    const APPS_SCRIPT_URL = 'YOUR_URL_HERE';
 *
 * SHEET STRUCTURE (auto-created on first run):
 *   Sheet 1: "Data Repository"  — all uploaded training records
 *   Sheet 2: "Upload Log"       — metadata for each monthly upload
 * ═══════════════════════════════════════════════════════════════
 */

const SPREADSHEET_ID = ''; // ← PASTE YOUR GOOGLE SHEET ID HERE
                             // (from the URL: docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit)

const DATA_SHEET    = 'Data Repository';
const LOG_SHEET     = 'Upload Log';

// ── COLUMN HEADERS ───────────────────────────────────────
const DATA_HEADERS = [
  'Employee No.', 'Employee Name', 'Functional Title', 'Grade', 'Group',
  'POP Code', 'Cluster', 'Attendance Status', 'Modules', 'Course Title',
  'Delivery Mode', 'Training Type', 'Training Status', 'From', 'To',
  'Duration', 'Working Hours', 'Category',
  '_month', '_quarter', '_uploadMonth', '_uploadedAt'
];

const LOG_HEADERS = [
  'Month Label', 'Record Count', 'Courses', 'Clusters',
  'Uploaded At', 'Uploaded By'
];

// ════════════════════════════════════════════════════════
// WEB APP ENTRY POINT
// ════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    if (action === 'upload')   return respond(handleUpload(payload));
    if (action === 'getData')  return respond(handleGetData(payload));
    if (action === 'getMeta')  return respond(handleGetMeta());
    if (action === 'delete')   return respond(handleDelete(payload));
    if (action === 'clearAll') return respond(handleClearAll());

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doGet(e) {
  // Allow GET for simple connectivity checks
  return respond({ success: true, message: 'UBL L&D Apps Script is live.' });
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════
// HANDLERS
// ════════════════════════════════════════════════════════

/**
 * Upload a month's data rows to the sheet
 * payload: { action, monthLabel, rows: [...] }
 */
function handleUpload(payload) {
  const { monthLabel, rows, uploadedBy } = payload;
  if (!rows || rows.length === 0) throw new Error('No rows provided.');

  const ss      = getSpreadsheet();
  const dataSheet = getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS);
  const logSheet  = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);

  // Delete existing rows for this month
  deleteRowsForMonth(dataSheet, monthLabel);

  // Append new rows
  const matrix = rows.map(r => DATA_HEADERS.map(h => r[h] || ''));
  if (matrix.length > 0) {
    dataSheet.getRange(
      dataSheet.getLastRow() + 1, 1,
      matrix.length, DATA_HEADERS.length
    ).setValues(matrix);
  }

  // Update log
  const courses  = countUnique(rows, 'Course Title');
  const clusters = [...new Set(rows.map(r => r['Cluster'] || '').filter(Boolean))].join(', ');
  updateLog(logSheet, monthLabel, rows.length, courses, clusters, uploadedBy || 'Admin');

  return {
    success: true,
    message: `${rows.length} records saved for ${monthLabel}.`,
    totalRows: dataSheet.getLastRow() - 1
  };
}

/**
 * Retrieve all data, optionally filtered
 * payload: { action, month?, quarter?, cluster?, category?, delivery? }
 */
function handleGetData(payload) {
  const ss        = getSpreadsheet();
  const dataSheet = getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS);
  const lastRow   = dataSheet.getLastRow();

  if (lastRow < 2) return { success: true, rows: [] };

  const values = dataSheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
  let rows = values.map(row => {
    const obj = {};
    DATA_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  // Apply filters
  if (payload.month   && payload.month   !== 'all') rows = rows.filter(r => r._month   === payload.month);
  if (payload.quarter && payload.quarter !== 'all') rows = rows.filter(r => r._quarter === payload.quarter);
  if (payload.cluster && payload.cluster !== 'all') rows = rows.filter(r => norm(r.Cluster) === payload.cluster);
  if (payload.category && payload.category !== 'all') rows = rows.filter(r => norm(r.Category) === payload.category);
  if (payload.delivery && payload.delivery !== 'all') rows = rows.filter(r => norm(r['Delivery Mode']) === payload.delivery);

  return { success: true, rows, total: rows.length };
}

/**
 * Retrieve upload log metadata
 */
function handleGetMeta() {
  const ss       = getSpreadsheet();
  const logSheet = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);
  const lastRow  = logSheet.getLastRow();

  if (lastRow < 2) return { success: true, meta: [] };

  const values = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  const meta   = values.map(row => {
    const obj = {};
    LOG_HEADERS.forEach((h, i) => { obj[h.replace(/ /g,'_').toLowerCase()] = row[i]; });
    return obj;
  });

  return { success: true, meta };
}

/**
 * Delete all rows for a given month
 * payload: { action, monthLabel }
 */
function handleDelete(payload) {
  const { monthLabel } = payload;
  const ss        = getSpreadsheet();
  const dataSheet = getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS);
  const logSheet  = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);

  const deleted = deleteRowsForMonth(dataSheet, monthLabel);
  deleteLogEntry(logSheet, monthLabel);

  return { success: true, message: `Deleted ${deleted} rows for ${monthLabel}.` };
}

/**
 * Clear all data from both sheets (keep headers)
 */
function handleClearAll() {
  const ss        = getSpreadsheet();
  const dataSheet = getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS);
  const logSheet  = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);

  clearSheetData(dataSheet);
  clearSheetData(logSheet);

  return { success: true, message: 'All data cleared.' };
}

// ════════════════════════════════════════════════════════
// SHEET HELPERS
// ════════════════════════════════════════════════════════

function getSpreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#0175C0')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function deleteRowsForMonth(sheet, monthLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // Find _uploadMonth column index (22nd column, 0-indexed)
  const uploadMonthIdx = DATA_HEADERS.indexOf('_uploadMonth') + 1;
  const values = sheet.getRange(2, uploadMonthIdx, lastRow - 1, 1).getValues();

  // Delete from bottom to top to avoid index shift
  let count = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === monthLabel) {
      sheet.deleteRow(i + 2);
      count++;
    }
  }
  return count;
}

function deleteLogEntry(sheet, monthLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() === monthLabel) sheet.deleteRow(i + 2);
  }
}

function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
}

function updateLog(logSheet, monthLabel, count, courses, clusters, uploadedBy) {
  deleteLogEntry(logSheet, monthLabel);
  logSheet.appendRow([
    monthLabel, count, courses, clusters,
    new Date().toLocaleString('en-PK'), uploadedBy
  ]);
}

function countUnique(rows, field) {
  return new Set(rows.map(r => r[field] || '').filter(Boolean)).size;
}

function norm(v) {
  return v == null ? '' : String(v).trim();
}

// ════════════════════════════════════════════════════════
// OPTIONAL: AUTO-FORMAT SHEET ON UPLOAD
// ════════════════════════════════════════════════════════
function formatDataSheet() {
  const ss   = getSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  // Alternate row colours
  for (let i = 2; i <= lastRow; i++) {
    const color = i % 2 === 0 ? '#e8f4fd' : '#ffffff';
    sheet.getRange(i, 1, 1, lastCol).setBackground(color);
  }

  // Auto-resize columns
  for (let c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
  }
}
