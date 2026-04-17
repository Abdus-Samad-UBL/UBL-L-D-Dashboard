const SPREADSHEET_ID = '1FdGv-vtcn-q-c9JTRYnw3wkCEeQoS2TTTjYYcA0ds0o';

const DATA_SHEET   = 'Data Repository';
const LOG_SHEET    = 'Upload Log';
const HC_SHEET     = 'Headcount';
const HC_LOG_SHEET = 'Headcount Log';
const PROG_SHEET   = 'Program Config';

const DATA_HEADERS = [
  'Region','Employee No.','Employee Name','Functional Title','Grade','Group',
  'POP Code','Cluster','Attendance Status','Course Title',
  'Training Type','Test Score','Trainer Name','Trainer Feedback',
  'From','To','Duration','Training Venue','Category',
  '_month','_quarter','_uploadMonth','_uploadedAt'
];
const LOG_HEADERS = ['Month Label','Record Count','Courses','Clusters','Uploaded At','Uploaded By'];

function doPost(e) {
  try {
    // Handle both Content-Type: application/json and text/plain (no Content-Type)
    // Browsers omitting Content-Type still send body as e.postData.contents
    const body = e.postData ? e.postData.contents : '{}';
    const p = JSON.parse(body);
    const a = p.action;
    if (a === 'upload')          return respond(handleUpload(p));
    if (a === 'delete')          return respond(handleDelete(p));
    if (a === 'clearAll')        return respond(handleClearAll());
    if (a === 'uploadHeadcount')  return respond(handleUploadHeadcount(p));
    if (a === 'deleteHeadcount')  return respond(handleDeleteHeadcount(p));
    if (a === 'clearHeadcount')   return respond(handleClearHeadcount());
    if (a === 'savePrograms')     return respond(handleSavePrograms(p));
    return respond({ success:false, error:'Unknown action: '+a });
  } catch(err) { return respond({ success:false, error:err.message }); }
}

function doGet(e) {
  try {
    const a = (e.parameter && e.parameter.action) || 'ping';
    if (a === 'ping')            return respond({ success:true, message:'UBL L&D backend live.' });
    if (a === 'getTrainingData') return respond(handleGetTrainingData(e.parameter));
    if (a === 'getHeadcount')    return respond(handleGetHeadcount());
    if (a === 'getMeta')         return respond(handleGetMeta());
    if (a === 'getPrograms')     return respond(handleGetPrograms());
    if (a === 'getResignedIds')  return respond(handleGetResignedIds());
    return respond({ success:false, error:'Unknown GET action: '+a });
  } catch(err) { return respond({ success:false, error:err.message }); }
}

function respond(data) {
  // Return JSON — Apps Script handles CORS automatically for deployed web apps
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle OPTIONS preflight (CORS)
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function handleGetTrainingData(params) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { success:true, rows:[] };

  const lastRow = sheet.getLastRow();
  const values  = sheet.getRange(2, 1, lastRow-1, DATA_HEADERS.length).getValues();

  let rows = values.map(row => {
    const obj = {};
    DATA_HEADERS.forEach((h,i) => {
      const v = row[i];
      if (v instanceof Date)
        obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      else
        obj[h] = (v === null || v === undefined) ? '' : String(v);
    });
    return obj;
  }).filter(r => r['Course Title'] || r['Employee No.']);

  if (params && params.month && params.month !== 'all')
    rows = rows.filter(r => r._uploadMonth === params.month || r._month === params.month);

  return { success:true, rows, total:rows.length };
}

function handleGetHeadcount() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(HC_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { success:true, rows:[], month:'' };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const values  = sheet.getRange(2,1,lastRow-1,lastCol).getValues();

  const rows = values.map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i] === null ? '' : String(row[i]); });
    return obj;
  }).filter(r => r['Employee No.']);

  let month = '';
  try {
    const log = ss.getSheetByName(HC_LOG_SHEET);
    if (log && log.getLastRow() >= 2) {
      const v = log.getRange(2,1,log.getLastRow()-1,1).getValues();
      month = String(v[v.length-1][0]);
    }
  } catch(e) {}

  return { success:true, rows, total:rows.length, month };
}

function handleGetMeta() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { success:true, meta:[] };
  const lastRow = sheet.getLastRow();
  const values  = sheet.getRange(2,1,lastRow-1,LOG_HEADERS.length).getValues();
  const meta = values.filter(r => r[0]).map(row => ({
    month: String(row[0]), count: row[1], courses: row[2],
    clusters: row[3], uploadedAt: String(row[4])
  }));
  return { success:true, meta };
}

function handleGetResignedIds() {
  try {
    const ids = JSON.parse(PropertiesService.getScriptProperties()
      .getProperty('resigned_ids') || '[]');
    return { success:true, resignedIds:ids };
  } catch(e) { return { success:true, resignedIds:[] }; }
}

function handleGetPrograms() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(PROG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { success:true, programs:[] };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(String);
  const values  = sheet.getRange(2,1,lastRow-1,lastCol).getValues();

  const programs = values.filter(r => r[0] || r[1]).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    // Parse courses JSON
    try { obj.courses = JSON.parse(obj.courses_json || '[]'); } catch(e) { obj.courses = []; }
    // Parse comma arrays
    ['groups','titles','phases'].forEach(f => {
      obj[f] = typeof obj[f] === 'string' && obj[f]
        ? obj[f].split(',').map(s => s.trim()).filter(Boolean) : [];
    });
    return obj;
  });

  return { success:true, programs };
}

function handleUpload(payload) {
  const { monthLabel, rows, isFirstChunk, isFinalChunk, totalRows, uploadedBy } = payload;
  if (!rows || !rows.length) throw new Error('No rows provided.');
  const ss        = getSpreadsheet();
  const dataSheet = getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS);
  const logSheet  = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);
  if (isFirstChunk !== false) deleteRowsForMonth(dataSheet, monthLabel);
  const matrix = rows.map(r => DATA_HEADERS.map(h => r[h] !== undefined ? r[h] : ''));
  if (matrix.length)
    dataSheet.getRange(dataSheet.getLastRow()+1,1,matrix.length,DATA_HEADERS.length).setValues(matrix);
  if (isFinalChunk !== false) {
    const finalTotal = totalRows || (dataSheet.getLastRow()-1);
    const courses    = countUnique(rows,'Course Title');
    const clusters   = [...new Set(rows.map(r=>r['Cluster']||'').filter(Boolean))].join(', ');
    updateLog(logSheet, monthLabel, finalTotal, courses, clusters, uploadedBy||'Admin');
  }
  return { success:true, message:`Saved ${rows.length} rows for ${monthLabel}.` };
}

function handleDelete(payload) {
  const ss   = getSpreadsheet();
  const data = getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS);
  const log  = getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS);
  const n    = deleteRowsForMonth(data, payload.monthLabel);
  deleteLogEntry(log, payload.monthLabel);
  return { success:true, message:`Deleted ${n} rows for ${payload.monthLabel}.` };
}

function handleClearAll() {
  const ss = getSpreadsheet();
  clearSheetData(getOrCreateSheet(ss, DATA_SHEET, DATA_HEADERS));
  clearSheetData(getOrCreateSheet(ss, LOG_SHEET, LOG_HEADERS));
  return { success:true };
}

function handleUploadHeadcount(payload) {
  const { monthLabel, rows, resignedIds, isFirstChunk, isFinalChunk, totalRows } = payload;
  if (!rows || !rows.length) throw new Error('No headcount rows provided.');
  const ss    = getSpreadsheet();
  const sheet = getOrCreateHcSheet(ss);
  const hcLog = getOrCreateSheet(ss, HC_LOG_SHEET, ['Month','Total','Uploaded At','Resigned Count']);
  if (isFirstChunk !== false) clearSheetData(sheet);
  const headers = Object.keys(rows[0]).filter(k => !k.startsWith('_'));
  ensureHcHeaders(sheet, headers);
  const matrix = rows.map(r => [...headers.map(h=>r[h]||''), monthLabel]);
  if (matrix.length)
    sheet.getRange(sheet.getLastRow()+1,1,matrix.length,headers.length+1).setValues(matrix);
  if (resignedIds && resignedIds.length)
    PropertiesService.getScriptProperties().setProperty('resigned_ids', JSON.stringify(resignedIds));
  if (isFinalChunk !== false) {
    const logLastRow = hcLog.getLastRow();
    if (logLastRow >= 2) {
      const lv = hcLog.getRange(2,1,logLastRow-1,1).getValues();
      for (let i=lv.length-1; i>=0; i--)
        if (String(lv[i][0]).trim()===monthLabel) hcLog.deleteRow(i+2);
    }
    hcLog.appendRow([monthLabel, totalRows||sheet.getLastRow()-1,
      new Date().toLocaleString(), (resignedIds||[]).length]);
  }
  return { success:true, message:`HC chunk saved: ${rows.length} rows.` };
}

function handleSavePrograms(payload) {
  const { programs } = payload;
  const ss   = getSpreadsheet();
  const cols = ['id','name','category','groups','notes','createdAt','courses_json'];
  const sheet = getOrCreateSheet(ss, PROG_SHEET, cols);
  clearSheetData(sheet);
  if (programs && programs.length) {
    const rowsData = programs.map(p => [
      p.id||'', p.name||'', p.category||'',
      (p.groups||[]).join(', '), p.notes||'', p.createdAt||'',
      JSON.stringify(p.courses||[])
    ]);
    sheet.getRange(2,1,rowsData.length,cols.length).setValues(rowsData);
  }
  return { success:true, message:`${(programs||[]).length} programs saved.` };
}

function handleDeleteHeadcount(payload) {
  // Headcount uses a replace model (one active HC), so deleting = clearing the sheet
  // We still accept monthLabel for logging purposes
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(HC_SHEET);
  if (!sheet) return { success:true, message:'No headcount sheet found.' };
  clearSheetData(sheet);
  // Remove log entry
  const logSheet = ss.getSheetByName(HC_LOG_SHEET);
  if (logSheet && payload.monthLabel) deleteLogEntry(logSheet, payload.monthLabel);
  // Clear resigned IDs
  PropertiesService.getScriptProperties().deleteProperty('resigned_ids');
  return { success:true, message:'Headcount deleted from Sheets.' };
}

function handleClearHeadcount() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(HC_SHEET);
  if (sheet) clearSheetData(sheet);
  const logSheet = ss.getSheetByName(HC_LOG_SHEET);
  if (logSheet) clearSheetData(logSheet);
  PropertiesService.getScriptProperties().deleteProperty('resigned_ids');
  return { success:true, message:'All headcount data cleared from Sheets.' };
}

function getSpreadsheet() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID)
                        : SpreadsheetApp.getActiveSpreadsheet();
}
function getOrCreateSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1,1,1,headers.length).setValues([headers]);
    s.getRange(1,1,1,headers.length).setBackground('#0175C0').setFontColor('#fff').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function getOrCreateHcSheet(ss) {
  let s = ss.getSheetByName(HC_SHEET);
  if (!s) { s = ss.insertSheet(HC_SHEET); s.setFrozenRows(1); }
  return s;
}
function ensureHcHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    const h = [...headers,'_uploadMonth'];
    sheet.getRange(1,1,1,h.length).setValues([h]);
    sheet.getRange(1,1,1,h.length).setBackground('#0175C0').setFontColor('#fff').setFontWeight('bold');
  }
}
function deleteRowsForMonth(sheet, monthLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const col  = DATA_HEADERS.indexOf('_uploadMonth') + 1;
  const vals = sheet.getRange(2, col, lastRow-1, 1).getValues();
  let n = 0;
  for (let i=vals.length-1; i>=0; i--)
    if (String(vals[i][0]).trim()===monthLabel) { sheet.deleteRow(i+2); n++; }
  return n;
}
function deleteLogEntry(sheet, monthLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const vals = sheet.getRange(2,1,lastRow-1,1).getValues();
  for (let i=vals.length-1; i>=0; i--)
    if (String(vals[i][0]).trim()===monthLabel) sheet.deleteRow(i+2);
}
function clearSheetData(sheet) {
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow()-1);
}
function updateLog(logSheet, monthLabel, count, courses, clusters, uploadedBy) {
  deleteLogEntry(logSheet, monthLabel);
  logSheet.appendRow([monthLabel, count, courses, clusters,
    new Date().toLocaleString('en-PK'), uploadedBy]);
}
function countUnique(rows, field) {
  return new Set(rows.map(r => r[field]||'').filter(Boolean)).size;
}
function norm(v) { return v==null?'':String(v).trim(); }
