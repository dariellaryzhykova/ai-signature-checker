function probeDrive() {
  var id = '1jg0faj8FH6sgF86izaGj1L0RC7dn-Zx-';
  try {
    var f = DriveApp.getFileById(id);
    Logger.log('Name: ' + f.getName());
    Logger.log('Mime: ' + f.getMimeType()); // should be application/pdf
    Logger.log('Size: ' + f.getSize());
  } catch (e) {
    Logger.log('Drive error: ' + e + "test");
  }
}

/***********************
 * CONFIG
 ***********************/
var ROOT_FOLDER_ID = 'PUT_2025_INSPECTION_FOLDER_ID_HERE'; // e.g., from URL .../folders/<ID>
var RECIPIENT_EMAIL = 'dariella.ryzhykova@gmail.com';
//var RECIPIENT_EMAIL = 'kendra.abrokwa@endotronix.com';
var SLACK_WEBHOOK_URL = 'YOUR_SLACK_WEBHOOK_URL';

var DOC_AI_PROJECT_ID = 'ai-qa-signature-checker';
var DOC_AI_LOCATION   = 'us';
var DOC_AI_PROCESSOR  = '5deea843de258b6a';

var DEBUG = true;        // verbose logs to console
var SEND_EMAIL = true;  // set true when ready to email results
var SEND_SLACK = false;  // set true when ready to post to Slack

// Only files whose name ends with " RTR.pdf" are processed (e.g., E250102-01 RTR.pdf)
var NAME_SUFFIX = ' RTR.pdf';

// Avoid reprocessing: store processed file IDs
var PROP_KEY_PROCESSED = 'PROCESSED_IDS';

/***********************
 * QUICK ONE-FILE TEST (by ID)
 ***********************/
function testSingleFile() {
  var fileId = '1EnEagMV6EWC0tz8diRrULEpNwv6aTpmD';
  Logger.log('Getting file by ID: ' + fileId);
  var file = DriveApp.getFileById(fileId);
  Logger.log('Name: ' + file.getName() + ' | Mime: ' + file.getMimeType() + ' | Size: ' + file.getSize());
  if (file.getMimeType() !== MimeType.PDF) {
    throw new Error('File is not a PDF. MimeType = ' + file.getMimeType());
  }
  processFile(file); // use the universal pipeline
}

/***********************
 * MAIN SCAN (entire 2025 Inspection tree)
 ***********************/
function checkNewUploads() {
  var processed = loadProcessedSet();
  var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  var candidates = [];

  Logger.log('Scanning root: ' + root.getName() + ' (' + ROOT_FOLDER_ID + ')');
  gatherCandidatesRecursively(root, candidates);
  Logger.log('Found ' + candidates.length + ' candidate files before de-dup.');

  // Process newest first (optional)
  candidates.sort(function(a,b){ return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });

  var newCount = 0;
  for (var i = 0; i < candidates.length; i++) {
    var file = candidates[i];
    var id = file.getId();
    if (processed[id]) {
      if (DEBUG) Logger.log('Skip already processed: ' + file.getName() + ' | ' + id);
      continue;
    }
    Logger.log('Processing: ' + file.getName() + ' | ' + id);
    try {
      processFile(file);
      processed[id] = true;
      newCount++;
    } catch (e) {
      Logger.log('Error processing ' + file.getName() + ': ' + e);
    }
  }

  saveProcessedSet(processed);
  Logger.log('Processing complete. New files processed this run: ' + newCount);
}

/***********************
 * DISCOVERY
 ***********************/
function gatherCandidatesRecursively(folder, outArr) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (shouldConsiderFile(f)) outArr.push(f);
  }
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    var sub = subs.next();
    gatherCandidatesRecursively(sub, outArr);
  }
}

function shouldConsiderFile(file) {
  try {
    var name = file.getName();
    var mime = file.getMimeType();
    if (mime !== MimeType.PDF) return false;
    if (!endsWith(name, NAME_SUFFIX)) return false;
    return true;
  } catch (e) {
    Logger.log('shouldConsiderFile error: ' + e);
    return false;
  }
}

function endsWith(s, suffix) {
  s = String(s || '');
  suffix = String(suffix || '');
  if (suffix.length === 0) return true;
  return s.substring(s.length - suffix.length) === suffix;
}

/***********************
 * PER FILE PIPELINE
 ***********************/
function processFile(file) {
  // 1) Document AI
  var docJson = callDocumentAI(file);

  // 2) Extract Quality rows from tables on all pages
  var rows = summarizeQualityFromTables(docJson, file.getName());

  // 3) Log to console
  logPlainTable('Quality Signature Summary: ' + file.getName(), rows);

  // 4) Output
  if (SEND_EMAIL && rows.length > 0) sendSummaryEmail(file.getName(), rows);
  if (SEND_SLACK && rows.length > 0) sendSlackSummary(file.getName(), rows);
}

/***********************
 * DOC AI CALL
 ***********************/
function callDocumentAI(file) {
  var url = 'https://' + DOC_AI_LOCATION + '-documentai.googleapis.com/v1/projects/' +
            DOC_AI_PROJECT_ID + '/locations/' + DOC_AI_LOCATION + '/processors/' + DOC_AI_PROCESSOR + ':process';

  var payload = {
    rawDocument: { content: Utilities.base64Encode(file.getBlob().getBytes()), mimeType: 'application/pdf' }
  };

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var status = resp.getResponseCode();
  var body = resp.getContentText();
  Logger.log('DocAI HTTP ' + status + ' for ' + file.getName());
  if (DEBUG) Logger.log('DocAI body (first 12000 chars): ' + body.substring(0, 12000));
  if (status < 200 || status >= 300) throw new Error('Document AI error ' + status + ': ' + body);
  return JSON.parse(body);
}

/***********************
 * UNIVERSAL TABLE PARSER FOR QUALITY ROWS
 ***********************/
function summarizeQualityFromTables(docJson, fileName) {
  var doc = docJson && docJson.document;
  if (!doc) { Logger.log('No document payload'); return []; }

  var fullText = doc.text || '';
  var pages = doc.pages || [];
  var out = [];

  for (var pi = 0; pi < pages.length; pi++) {
    var page = pages[pi];
    var pageNum = page.pageNumber || (pi + 1);
    var tables = page.tables || [];
    var pageTextLower = getPageText(page, fullText).toLowerCase();

    if (DEBUG) Logger.log('Page ' + pageNum + ': tables=' + tables.length);

    for (var ti = 0; ti < tables.length; ti++) {
      var t = tables[ti];

      // Determine header
      var headerArr = [];
      if (t.headerRows && t.headerRows.length > 0) {
        headerArr = extractRowTextArray(t.headerRows[0], fullText);
      } else if (t.bodyRows && t.bodyRows.length > 0) {
        headerArr = extractRowTextArray(t.bodyRows[0], fullText);
      }
      if (DEBUG) Logger.log('  Table ' + (ti+1) + ' header guess: ' + JSON.stringify(headerArr));

      var idxFunction  = findColumn(headerArr, ['function', 'role', 'department']);
      var idxName      = findColumn(headerArr, ['print name', 'printed name', 'name', 'printed by']);
      var idxSignature = findColumn(headerArr, ['signature', 'sign']);
      var idxDate      = findColumn(headerArr, ['date', 'signed date', 'signature date']);

      if (idxFunction === -1 || (idxName === -1 && idxSignature === -1 && idxDate === -1)) {
        if (DEBUG) Logger.log('    Skip table (no recognizable columns)');
        continue;
      }

      var bodyStart = (t.headerRows && t.headerRows.length > 0) ? 0 : 1;
      var bodyRows = t.bodyRows || [];

      for (var ri = bodyStart; ri < bodyRows.length; ri++) {
        var cells = extractRowTextArray(bodyRows[ri], fullText);
        var funcText = safeCell(cells, idxFunction);
        if (!funcText) continue;

        var f = funcText.toLowerCase();
        var isQuality = (f.indexOf('quality') >= 0) || (f === 'qa') || (f.indexOf('quality assurance') >= 0);
        if (!isQuality) continue;

        var nameText = safeCell(cells, idxName);
        var sigText  = safeCell(cells, idxSignature);
        var dateText = safeCell(cells, idxDate);

        // Heuristics if columns are imperfect
        if (!nameText) nameText = firstNonMatchingCell(cells, [idxFunction, idxSignature, idxDate]);
        if (!dateText) dateText = firstCellThatLooksLikeDate(cells, idxFunction);
        if (!sigText && idxSignature === -1) sigText = guessSignatureCell(cells, [idxFunction, idxName, idxDate]);

        var sectionLabel = inferSectionLabel(headerArr, pageTextLower);
        out.push({
          section: sectionLabel || 'Completion Sign-Off',
          status: normalizePresence(sigText) ? 'Yes' : 'Missing',
          name: nameText || '',
          date: normalizeDate(dateText),
          signature: sigText || '',
          page: pageNum
        });
      }
    }
  }

  return out;
}

/***********************
 * PAGE TEXT + SECTION INFERENCE
 ***********************/
function getPageText(page, fullText) {
  try {
    if (page && page.layout && page.layout.textAnchor &&
        page.layout.textAnchor.textSegments &&
        page.layout.textAnchor.textSegments.length > 0) {
      var seg = page.layout.textAnchor.textSegments[0];
      var s = seg.startIndex ? parseInt(seg.startIndex, 10) : 0;
      var e = seg.endIndex ? parseInt(seg.endIndex, 10) : 0;
      if (e > s && e <= fullText.length) return fullText.substring(s, e);
    }
  } catch (e) {}
  return '';
}

function inferSectionLabel(headerArr, pageTextLower) {
  if (pageTextLower) {
    if (pageTextLower.indexOf('router (task list) completion sign-off') >= 0 ||
        pageTextLower.indexOf('router task list') >= 0) return 'Router Task List Completion';
    if (pageTextLower.indexOf('materials list completion') >= 0) return 'Materials List Completion';
    if (pageTextLower.indexOf('dhr log completion') >= 0) return 'DHR Log Completion';
    if (pageTextLower.indexOf('configuration forms') >= 0) return 'FM3811 Configuration Forms';
    if (pageTextLower.indexOf('label printing form') >= 0 && pageTextLower.indexOf('mpi026-fm01') >= 0) return 'MPI026-FM01 Label Printing Form';
    if (pageTextLower.indexOf('label printing form') >= 0 && pageTextLower.indexOf('mpi026-fm02') >= 0) return 'MPI026-FM02 Label Printing Form';
    if (pageTextLower.indexOf('chfs kit assembly form') >= 0 || pageTextLower.indexOf('mpi022-fm02') >= 0) return 'MPI022-FM02 CHFS Kit Assembly Form';
    if (pageTextLower.indexOf('completion sign-off') >= 0) return 'Completion Sign-Off';
  }

  var generic = {
    'function':1,'role':1,'department':1,'print name':1,'printed name':1,'name':1,'printed by':1,
    'signature':1,'sign':1,'date':1,'signed date':1,'signature date':1
  };
  if (headerArr && headerArr.length) {
    for (var i = 0; i < headerArr.length; i++) {
      var val = (headerArr[i] || '').trim();
      if (!val) continue;
      if (!generic[val.toLowerCase ? val.toLowerCase() : val]) return val;
    }
  }
  return 'Completion Sign-Off';
}

/***********************
 * TABLE/TEXT HELPERS
 ***********************/
function extractRowTextArray(rowObj, fullText) {
  if (!rowObj || !rowObj.cells) return [];
  var out = [];
  for (var i = 0; i < rowObj.cells.length; i++) {
    out.push(textFromAnchor(rowObj.cells[i].layout && rowObj.cells[i].layout.textAnchor, fullText));
  }
  return out;
}

function textFromAnchor(anchor, fullText) {
  if (!anchor || !anchor.textSegments || !fullText) return '';
  var s = '';
  for (var i = 0; i < anchor.textSegments.length; i++) {
    var seg = anchor.textSegments[i];
    var start = seg.startIndex ? parseInt(seg.startIndex, 10) : 0;
    var end   = seg.endIndex ? parseInt(seg.endIndex, 10) : 0;
    if (end > start && end <= fullText.length) s += fullText.substring(start, end);
  }
  return s.replace(/\s+/g, ' ').trim();
}

function findColumn(headerArray, keys) {
  if (!headerArray || !headerArray.length) return -1;
  for (var i = 0; i < headerArray.length; i++) {
    var h = headerArray[i] ? headerArray[i].toLowerCase() : '';
    for (var k = 0; k < keys.length; k++) if (h === keys[k].toLowerCase()) return i;
  }
  for (var j = 0; j < headerArray.length; j++) {
    var h2 = headerArray[j] ? headerArray[j].toLowerCase() : '';
    for (var k2 = 0; k2 < keys.length; k2++) if (h2.indexOf(keys[k2].toLowerCase()) >= 0) return j;
  }
  for (var j2 = 0; j2 < headerArray.length; j2++) {
    var h3 = headerArray[j2] ? headerArray[j2].toLowerCase().replace(/[:\-]/g, ' ').trim() : '';
    for (var k3 = 0; k3 < keys.length; k3++) if (h3.indexOf(keys[k3].toLowerCase()) >= 0) return j2;
  }
  return -1;
}

function safeCell(arr, idx) {
  if (idx < 0 || idx >= arr.length) return '';
  return (arr[idx] || '').trim();
}

function firstNonMatchingCell(cells, skipIdxs) {
  var skip = {};
  for (var i = 0; i < skipIdxs.length; i++) skip[skipIdxs[i]] = true;
  for (var j = 0; j < cells.length; j++) {
    if (skip[j]) continue;
    var v = (cells[j] || '').trim();
    if (v) return v;
  }
  return '';
}

function firstCellThatLooksLikeDate(cells, funcIdx) {
  for (var i = 0; i < cells.length; i++) {
    if (i === funcIdx) continue;
    var v = (cells[i] || '').trim();
    if (looksLikeDate(v)) return v;
  }
  return '';
}

function looksLikeDate(s) {
  if (!s) return false;
  var v = s.trim();
  if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(v)) return true;  // 2025-02-06
  if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/.test(v)) return true;  // 02/06/2025
  if (/^[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}$/.test(v)) return true; // Feb 6, 2025
  return false;
}

function normalizeDate(s) {
  if (!s) return '';
  var v = s.trim();
  var m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(v);
  if (m) return zero(m[1],4) + '-' + zero(m[2],2) + '-' + zero(m[3],2);
  m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(v);
  if (m) return zero(m[3],4) + '-' + zero(m[1],2) + '-' + zero(m[2],2);
  return v;
}

function zero(n, width) { n = String(n); while (n.length < width) n = '0' + n; return n; }

function guessSignatureCell(cells, skipIdxs) {
  var skip = {};
  for (var i = 0; i < skipIdxs.length; i++) skip[skipIdxs[i]] = true;
  for (var j = 0; j < cells.length; j++) {
    if (skip[j]) continue;
    var v = (cells[j] || '').trim();
    if (!v) continue;
    if (!looksLikeDate(v) && v.toLowerCase().indexOf('quality') === -1) return v;
  }
  return '';
}

function normalizePresence(sigCell) {
  if (!sigCell) return false;
  var v = sigCell.trim().toLowerCase();
  if (!v) return false;
  if (v === 'n/a' || v === 'na' || v === 'none') return false;
  return true;
}

/***********************
 * OUTPUTS
 ***********************/
function sendSummaryEmail(fileName, rows) {
  var html = '';
  html += '<h2>Quality Signature Summary for <em>' + escapeHtml(fileName) + '</em></h2>';
  html += '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">';
  html += '<tr><th>Section / Form</th><th>Quality Signature Present?</th><th>Name</th><th>Date</th><th>Page</th></tr>';
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    html += '<tr>';
    html += '<td>' + escapeHtml(r.section) + '</td>';
    html += '<td>' + escapeHtml(r.status) + '</td>';
    html += '<td>' + escapeHtml(r.name) + '</td>';
    html += '<td>' + escapeHtml(r.date) + '</td>';
    html += '<td>' + escapeHtml(String(r.page)) + '</td>';
    html += '</tr>';
  }
  html += '</table>';
  GmailApp.sendEmail(RECIPIENT_EMAIL, 'Quality Signature Summary: ' + fileName, 'See HTML body', { htmlBody: html });
}

function sendSlackSummary(fileName, rows) {
  var blocks = [
    { "type": "header", "text": { "type": "plain_text", "text": "QA Signature Summary: " + fileName } },
    { "type": "section", "text": { "type": "mrkdwn", "text": "*Summary of Quality Signatures:*" } }
  ];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    blocks.push({
      "type": "section",
      "text": { "type": "mrkdwn",
        "text": "*" + r.section + "* (Page " + r.page + ")\n" +
                 "Status: " + r.status + "\n" +
                 "Name: " + r.name + "\n" +
                 "Date: " + r.date }
    });
    blocks.push({ "type": "divider" });
  }
  var payload = JSON.stringify({ blocks: blocks });
  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: payload });
}

/***********************
 * LOG TABLE
 ***********************/
function logPlainTable(title, rows) {
  Logger.log('---- ' + title + ' ----');
  if (!rows || !rows.length) { Logger.log('(no rows)'); return; }
  var headers = ['Section / Form','Present?','Name','Date','Page'];
  var widths  = [36, 8, 26, 14, 6];
  var line = pad(headers[0], widths[0]) + ' | ' +
             pad(headers[1], widths[1]) + ' | ' +
             pad(headers[2], widths[2]) + ' | ' +
             pad(headers[3], widths[3]) + ' | ' +
             pad(headers[4], widths[4]);
  Logger.log(line);
  Logger.log(Array(line.length + 1).join('-'));
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var ln = pad(r.section, widths[0]) + ' | ' +
             pad(r.status,  widths[1]) + ' | ' +
             pad(r.name,    widths[2]) + ' | ' +
             pad(r.date,    widths[3]) + ' | ' +
             pad(String(r.page), widths[4]);
    Logger.log(ln);
  }
}

function pad(s, len) { s = s == null ? '' : String(s); if (s.length > len) return s.substring(0, len-1) + '…'; while (s.length < len) s += ' '; return s; }
function escapeHtml(s) { if (s == null) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }

/***********************
 * STATE
 ***********************/
function loadProcessedSet() {
  var raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY_PROCESSED) || '{}';
  try { return JSON.parse(raw) || {}; } catch(e) { return {}; }
}
function saveProcessedSet(obj) {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY_PROCESSED, JSON.stringify(obj || {}));
}

/***********************
 * OPTIONAL: TIME TRIGGER
 ***********************/
function installTrigger_every5min() {
  ScriptApp.newTrigger('checkNewUploads').timeBased().everyMinutes(5).create();
}

/***********************
 * OPTIONAL: ONE-FILE BY URL OR REST FALLBACK
 ***********************/
function processSingleFileByUrl(url) {
  var id = extractDriveId(url);
  processSingleFileById(id);
}
function processSingleFileById(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    if (file.getMimeType() !== MimeType.PDF) throw new Error('Not a PDF: ' + file.getMimeType());
    processFile(file);
  } catch (e) {
    Logger.log('DriveApp failed, trying Drive REST: ' + e);
    processSingleFile_viaDriveApi(fileId);
  }
}
function processSingleFile_viaDriveApi(fileId) {
  var token = ScriptApp.getOAuthToken();
  var meta = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=id,name,mimeType,size&supportsAllDrives=true', {
    method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (meta.getResponseCode() !== 200) throw new Error('Drive metadata failed: ' + meta.getContentText());
  var m = JSON.parse(meta.getContentText());
  if (m.mimeType !== 'application/pdf') throw new Error('Not a PDF: ' + m.mimeType);

  var dl = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media&supportsAllDrives=true', {
    method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (dl.getResponseCode() !== 200) throw new Error('Drive download failed: ' + dl.getContentText());

  var docJson = callDocumentAI_bytes(dl.getContent());
  var rows = summarizeQualityFromTables(docJson, m.name || 'document.pdf');
  logPlainTable('Quality Signature Summary (single file via REST): ' + (m.name || fileId), rows);
  if (SEND_EMAIL && rows.length > 0) sendSummaryEmail(m.name || 'document.pdf', rows);
  if (SEND_SLACK && rows.length > 0) sendSlackSummary(m.name || 'document.pdf', rows);
}
function callDocumentAI_bytes(pdfBytes) {
  var url = 'https://' + DOC_AI_LOCATION + '-documentai.googleapis.com/v1/projects/' +
            DOC_AI_PROJECT_ID + '/locations/' + DOC_AI_LOCATION + '/processors/' + DOC_AI_PROCESSOR + ':process';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), 'Content-Type': 'application/json' },
    payload: JSON.stringify({ rawDocument: { content: Utilities.base64Encode(pdfBytes), mimeType: 'application/pdf' } }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    throw new Error('DocAI error ' + resp.getResponseCode() + ': ' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}
function extractDriveId(urlOrId) {
  if (urlOrId.indexOf('http') !== 0 && urlOrId.indexOf('/') < 0) return urlOrId;
  var m = /\/d\/([A-Za-z0-9_-]+)/.exec(urlOrId) || /[?&]id=([A-Za-z0-9_-]+)/.exec(urlOrId);
  return m ? m[1] : urlOrId;
}





