
  //SPREADSHEET_ID: '1oqs__qWVG1DWC5FoBI3LugVBG_lLgUAoXUNV38CfIbM', // Leave blank to use the bound spreadsheet (recommended)
  //RESUMES_FOLDER_ID: '1ksNJzhqBofXbAAh3RsXOkGf9ke4O-eTd',


/**
 * Resort Recruitment System — Backend
 * Google Apps Script Web App + Google Sheets.
 *
 * SETUP (once):
 *  1. Create a Google Sheet, open Extensions > Apps Script.
 *  2. Paste Code.gs, create an HTML file named "Index", paste Index.html.
 *  3. Project Settings > show appsscript.json > paste appsscript.json.
 *     (This also enables the Advanced Drive Service used for resume parsing.)
 *  4. Run initialSetup() once and grant access.
 *  5. Deploy > New deployment > Web app (see README for execute-as guidance).
 */

var CONFIG = {
  SPREADSHEET_ID: '1oqs__qWVG1DWC5FoBI3LugVBG_lLgUAoXUNV38CfIbM', // blank = use the bound spreadsheet
  SHEETS: {
    CANDIDATES: 'Candidates',
    INTERVIEWS: 'Interviews',
    USERS: 'Users',
    AUDIT_LOG: 'AuditLog'
  },
  STATUSES: ['New', 'Not Screened', 'Shortlisted', 'Interviewed', 'Selected', 'Rejected', 'On Hold'],
  ROLES: ['Admin', 'HR', 'Interviewer', 'Viewer'],
  RESUMES_FOLDER_ID: '1Y2Q8j3N8DR_GyxNrSW-XLPyz6cuoBZYf',
  MAX_UPLOAD_BYTES: 20 * 1024 * 1024
};

// ============================================================
// WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('SM Recruitment')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body style="font-family:system-ui;padding:30px;color:#0f172a">' +
      '<h2 style="color:#dc2626">Server error</h2><p>' +
      escapeHtmlServer_(err && err.message ? err.message : 'Unknown error') +
      '</p><p style="color:#64748b">Most common cause: the HTML file is not named exactly <b>Index</b>, ' +
      'or initialSetup() was never run. See README.</p></body></html>'
    );
  }
}

// ============================================================
// AUTHORIZATION
// ============================================================
function authorize() {
  var out = {};
  try { out.email = getActiveEmail_(); } catch (e) { out.emailErr = e.message; }
  try { out.spreadsheet = getSpreadsheet_().getName(); } catch (e) { out.sheetErr = e.message; }
  try { out.driveFolder = getResumesFolder_().getName(); } catch (e) { out.driveErr = e.message; }
  Logger.log('authorize() -> %s', JSON.stringify(out));
  return out;
}

function ping() {
  return { ok: true, time: nowIso_(), user: getActiveEmail_() };
}

// ============================================================
// HELPERS
// ============================================================
function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No bound spreadsheet. Bind the script to a Sheet or set CONFIG.SPREADSHEET_ID.');
  return ss;
}
function getSheet_(name) {
  var sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run initialSetup().');
  return sh;
}
function getResumesFolder_() {
  try { return DriveApp.getFolderById(CONFIG.RESUMES_FOLDER_ID); }
  catch (e) { throw new Error('Cannot access resumes folder ' + CONFIG.RESUMES_FOLDER_ID + ': ' + e.message); }
}
function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0], rows = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null) continue;
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var v = values[i][j];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = isNaN(v.getTime()) ? '' : v.toISOString();
      }
      row[headers[j]] = v;
    }
    rows.push(row);
  }
  return rows;
}
function findRowByKey_(sheet, keyCol, keyVal) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(keyVal)) return i + 1;
  }
  return -1;
}
function uuid_() { return Utilities.getUuid(); }
function nowIso_() { return new Date().toISOString(); }
function getActiveEmail_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) { try { email = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  return email || '';
}
function escapeHtmlServer_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function safeFileName_(s) { return String(s || 'resume').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80); }

// Short content hash (first 12 hex of MD5) for duplicate-file detection.
function hashBytes_(bytes) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes);
  var hex = digest.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  return hex.slice(0, 12);
}
// Look for an existing file in the folder whose name carries the same hash tag.
function findFileByHash_(folder, hash) {
  var it = folder.searchFiles('title contains "h-' + hash + '"');
  while (it.hasNext()) {
    var f = it.next();
    if (!f.isTrashed() && f.getName().indexOf('h-' + hash) !== -1) return f;
  }
  return null;
}

// ---- Field validation ----
function validEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(s || '').trim()); }
function validPhone_(s) {
  var d = String(s || '').replace(/[^\d]/g, '');
  return d.length >= 7 && d.length <= 15;
}

// ============================================================
// SETUP
// ============================================================
function initialSetup() {
  var ss = getSpreadsheet_();
  var defs = {
    'Candidates': ['CandidateID', 'Name', 'Phone', 'Email', 'RoleApplied', 'ResumeLink', 'Source', 'Status', 'CreatedAt'],
    'Interviews': ['InterviewID', 'CandidateID', 'Round', 'Interviewer', 'DateTime', 'Status', 'Feedback', 'Score'],
    'Users': ['UserID', 'Name', 'Email', 'Role'],
    'AuditLog': ['LogID', 'Email', 'Action', 'Details', 'Timestamp']
  };
  Object.keys(defs).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var first = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    if (first.join('') === '') {
      sh.getRange(1, 1, 1, defs[name].length).setValues([defs[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold')
        .setBackground('#6366f1').setFontColor('#ffffff');
    }
  });
  var users = ss.getSheetByName('Users');
  var adminEmail = getActiveEmail_() || 'admin@example.com';
  if (users.getLastRow() < 2) users.appendRow([uuid_(), 'Admin User', adminEmail, 'Admin']);

  var cand = ss.getSheetByName('Candidates');
  if (cand.getLastRow() < 2) {
    var now = new Date();
    function daysAgo(n) { return new Date(now.getTime() - n * 864e5).toISOString(); }
    var ids = {};
    ['arjun', 'meera', 'sandeep', 'kavya', 'rohit', 'neha', 'farhan', 'divya'].forEach(function (k) { ids[k] = uuid_(); });
    cand.getRange(2, 1, 8, 9).setValues([
      [ids.arjun, 'Arjun Nair', '+919812340001', 'arjun.nair@example.com', 'Front Desk Executive', '', 'Naukri', 'Selected', daysAgo(40)],
      [ids.meera, 'Meera Iyer', '+919812340002', 'meera.iyer@example.com', 'Housekeeping Supervisor', '', 'Referral', 'Interviewed', daysAgo(25)],
      [ids.sandeep, 'Sandeep Kumar', '+919812340003', 'sandeep.k@example.com', 'Chef de Partie', '', 'Walk-in', 'Shortlisted', daysAgo(15)],
      [ids.kavya, 'Kavya Reddy', '+919812340004', 'kavya.reddy@example.com', 'Spa Therapist', '', 'LinkedIn', 'Not Screened', daysAgo(8)],
      [ids.rohit, 'Rohit Sharma', '+919812340005', 'rohit.sharma@example.com', 'F&B Manager', '', 'Referral', 'Rejected', daysAgo(35)],
      [ids.neha, 'Neha Pillai', '+919812340006', 'neha.pillai@example.com', 'Front Desk Executive', '', 'Naukri', 'Not Screened', daysAgo(3)],
      [ids.farhan, 'Farhan Ahmed', '+919812340007', 'farhan.a@example.com', 'Sous Chef', '', 'Indeed', 'New', daysAgo(2)],
      [ids.divya, 'Divya Menon', '+919812340008', 'divya.menon@example.com', 'Guest Relations', '', 'Walk-in', 'On Hold', daysAgo(20)]
    ]);
    var iv = ss.getSheetByName('Interviews');
    iv.getRange(2, 1, 4, 8).setValues([
      [uuid_(), ids.arjun, 'HR', adminEmail, daysAgo(35), 'Completed', 'Strong hospitality fit.', 9],
      [uuid_(), ids.meera, 'HR', adminEmail, daysAgo(20), 'Completed', 'Solid experience.', 8],
      [uuid_(), ids.sandeep, 'Technical', adminEmail, daysAgo(10), 'Completed', 'Good knife skills.', 7],
      [uuid_(), ids.farhan, 'Technical', adminEmail, new Date(now.getTime() + 2 * 864e5).toISOString(), 'Scheduled', '', '']
    ]);
  }
  return 'Setup complete.';
}

// ============================================================
// AUTH / RBAC
// ============================================================
function getCurrentUser() {
  var email = getActiveEmail_();
  if (!email) return { authorized: false, message: 'Unable to detect your Google account.' };
  var users = sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS));
  var me = null;
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].Email).toLowerCase() === email.toLowerCase()) { me = users[i]; break; }
  }
  if (!me && users.length === 0) me = { Name: email, Email: email, Role: 'Admin' };
  if (!me) return { authorized: false, email: email, message: 'Access denied. ' + email + ' is not a registered user.' };
  return { authorized: true, email: email, name: me.Name || email, role: me.Role, permissions: rolePermissions_(me.Role) };
}
function rolePermissions_(role) {
  switch (role) {
    case 'Admin': return { manageUsers: true, addCandidate: true, editCandidate: true, scheduleInterview: true, addFeedback: true, changeStatus: true, exportCsv: true };
    case 'HR': return { manageUsers: false, addCandidate: true, editCandidate: true, scheduleInterview: true, addFeedback: true, changeStatus: true, exportCsv: false };
    case 'Interviewer': return { manageUsers: false, addCandidate: false, editCandidate: false, scheduleInterview: false, addFeedback: true, changeStatus: false, exportCsv: false };
    case 'Viewer': return { manageUsers: false, addCandidate: false, editCandidate: false, scheduleInterview: false, addFeedback: false, changeStatus: false, exportCsv: false };
    default: return {};
  }
}
function authorizeUser_(requiredRoles) {
  var info = getCurrentUser();
  if (!info.authorized) throw new Error(info.message || 'Access denied.');
  if (requiredRoles && requiredRoles.length && requiredRoles.indexOf(info.role) === -1) {
    throw new Error('Access denied: requires ' + requiredRoles.join(' or ') + '.');
  }
  return info;
}
function logAudit_(email, action, details) {
  try { getSheet_(CONFIG.SHEETS.AUDIT_LOG).appendRow([uuid_(), email, action, JSON.stringify(details || {}), nowIso_()]); }
  catch (e) {}
}

// ============================================================
// BOOTSTRAP
// ============================================================
function bootstrapApp() {
  var user = getCurrentUser();
  if (!user.authorized) return { user: user };
  var candidates = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  var interviews = sheetToObjects_(getSheet_(CONFIG.SHEETS.INTERVIEWS));
  if (user.role === 'Interviewer') {
    var mine = {};
    interviews = interviews.filter(function (i) {
      var keep = String(i.Interviewer).toLowerCase() === user.email.toLowerCase();
      if (keep) mine[i.CandidateID] = true; return keep;
    });
    candidates = candidates.filter(function (c) { return mine[c.CandidateID]; });
  }
  var counts = {}; CONFIG.STATUSES.forEach(function (s) { counts[s] = 0; });
  var roleMap = {};
  candidates.forEach(function (c) {
    if (counts[c.Status] !== undefined) counts[c.Status]++;
    var r = c.RoleApplied || 'Unspecified';
    if (!roleMap[r]) roleMap[r] = { role: r, total: 0, selected: 0 };
    roleMap[r].total++;
    if (c.Status === 'Selected') roleMap[r].selected++;
  });
  var total = candidates.length || 1;
  var screened = counts['Shortlisted'] + counts['Interviewed'] + counts['Selected'] + counts['Rejected'];
  var interviewedPlus = counts['Interviewed'] + counts['Selected'];
  return {
    user: user, candidates: candidates, interviews: interviews,
    kpis: {
      total: candidates.length, new: counts['New'], notScreened: counts['Not Screened'],
      shortlisted: counts['Shortlisted'], interviewed: counts['Interviewed'],
      selected: counts['Selected'], rejected: counts['Rejected'], onHold: counts['On Hold']
    },
    metrics: {
      screeningRate: Math.round((screened / total) * 100),
      interviewConversionRate: interviewedPlus ? Math.round((counts['Selected'] / interviewedPlus) * 100) : 0,
      rejectionRate: Math.round((counts['Rejected'] / total) * 100),
      funnel: { new: candidates.length, screened: screened, interviewed: interviewedPlus, selected: counts['Selected'] }
    },
    roleStats: Object.keys(roleMap).map(function (k) { return roleMap[k]; }),
    serverTime: nowIso_()
  };
}

// ============================================================
// RESUME UPLOAD + AUTO-EXTRACT  (A + B: regex first, Gemini fallback)
//
// 1. Save the uploaded file to the configured Drive folder.
// 2. Convert it to text (temp Google Doc, OCR for images) and run a fast
//    offline regex pass — free, instant, good for clean text resumes.
// 3. If any of name / phone / email is still missing AND a Gemini API key is
//    configured, send the original file straight to Gemini Flash (it reads
//    PDFs / images natively, so it also handles scanned/image-only PDFs) and
//    fill in the gaps. Gemini also suggests the role.
// ============================================================
function setGeminiKey(key) {
  if (!key) throw new Error('Pass your Gemini API key, e.g. setGeminiKey("AIza...").');
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', String(key).trim());
  return 'Gemini API key saved.';
}
function getGeminiKey_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}
function setGroqKey(key) {
  if (!key) throw new Error('Pass your Groq API key, e.g. setGroqKey("gsk_...").');
  PropertiesService.getScriptProperties().setProperty('GROQ_API_KEY', String(key).trim());
  return 'Groq API key saved.';
}
function getGroqKey_() {
  return PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY') || '';
}

function scanResume(payload) {
  authorizeUser_(['Admin', 'HR']);
  if (!payload || !payload.base64 || !payload.fileName) throw new Error('Missing upload data.');

  var b64 = String(payload.base64).replace(/^data:[^;]+;base64,/, '');
  var bytes = Utilities.base64Decode(b64);
  if (bytes.length > CONFIG.MAX_UPLOAD_BYTES) {
    throw new Error('File too large. Limit is ' + Math.round(CONFIG.MAX_UPLOAD_BYTES / 1024 / 1024) + ' MB.');
  }

  var mime = payload.mimeType || 'application/octet-stream';
  var ext = (payload.fileName.match(/\.[^.]+$/) || [''])[0];
  var displayName = safeFileName_(payload.fileName).replace(/\.[^.]+$/, '');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');

  // Content hash to detect duplicate uploads of the SAME file. We embed a short
  // hash in the saved file name (h-XXXX) and look for it before creating a copy.
  var hash = hashBytes_(bytes);
  var folder = getResumesFolder_();
  var dup = findFileByHash_(folder, hash);
  if (dup) {
    return {
      ok: true, duplicate: true,
      url: dup.getUrl(), fileName: dup.getName(),
      suggested: { name: '', phone: '', email: '', role: '' },
      method: {}, usedAi: false, aiMode: '', aiConfigured: !!(getGeminiKey_() || getGroqKey_()),
      aiError: '', extractedChars: 0, textOk: false, debugText: ''
    };
  }

  var savedName = displayName + '_' + stamp + '_h-' + hash + ext;
  var savedBlob = Utilities.newBlob(bytes, mime, savedName);
  var savedFile = folder.createFile(savedBlob);
  try { savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { savedFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {} }

  // --- Pass 1: offline parsing over Google-OCR'd text (free, no quota) ---
  var text = '';
  try { text = extractText_(savedBlob, mime, folder.getId()); }
  catch (e) { logAudit_(getActiveEmail_(), 'extractTextFailed', { file: savedName, err: e.message }); }
  // Some PDFs have a broken/encoded text layer that converts to garbage. If the
  // text doesn't look like real words, treat it as empty so text-only parsing
  // (regex / Groq) is skipped and AI is asked to READ THE FILE instead.
  var goodText = textLooksReal_(text) ? text : '';
  var out = parseResume_(goodText, payload.fileName);
  out.role = out.role || '';
  var method = { name: out.name ? 'regex' : '', phone: out.phone ? 'regex' : '', email: out.email ? 'regex' : '' };

  // --- Pass 2: AI fallback for whatever regex missed. Gemini first (it can read
  // the file itself — handles scanned / broken-text-layer PDFs), then Groq
  // (text-only, only useful when we have real text). Best-effort; never blocks. ---
  var usedAi = false, aiError = '', aiMode = '';
  function mergeAi_(ai) {
    usedAi = true; aiMode = ai.__mode;
    ['name', 'phone', 'email', 'role'].forEach(function (f) {
      if ((!out[f] || out[f] === '') && ai[f]) { out[f] = ai[f]; if (f !== 'role') method[f] = 'ai'; }
    });
  }
  function stillMissing_() { return !out.name || !out.phone || !out.email; }

  if (stillMissing_() && getGeminiKey_()) {
    try { mergeAi_(geminiExtract_(goodText, b64, mime)); }
    catch (e) { aiError = 'Gemini: ' + e.message; logAudit_(getActiveEmail_(), 'geminiFailed', { file: savedName, err: e.message }); }
  }
  if (stillMissing_() && getGroqKey_() && goodText && goodText.length > 60) {
    try { mergeAi_(groqExtract_(goodText)); aiError = ''; }
    catch (e) { aiError = (aiError ? aiError + ' | ' : '') + 'Groq: ' + e.message; logAudit_(getActiveEmail_(), 'groqFailed', { file: savedName, err: e.message }); }
  }

  return {
    ok: true,
    url: savedFile.getUrl(),
    fileName: savedName,
    suggested: { name: out.name, phone: out.phone, email: out.email, role: out.role || '' },
    method: method,
    usedAi: usedAi,
    aiMode: aiMode,
    aiConfigured: !!(getGeminiKey_() || getGroqKey_()),
    aiError: aiError,
    extractedChars: text.length,
    textOk: !!goodText,
    debugText: (text || '').replace(/\s+/g, ' ').trim().slice(0, 600)
  };
}

// Heuristic: does the OCR'd text contain enough real words to parse? PDFs with
// broken font encodings convert to mostly blank/symbol text.
function textLooksReal_(text) {
  if (!text) return false;
  var words = text.match(/[A-Za-z]{2,}/g) || [];
  var alpha = (text.match(/[A-Za-z]/g) || []).length;
  var nonSpace = text.replace(/\s/g, '').length || 1;
  return words.length >= 12 && (alpha / nonSpace) >= 0.45;
}

function extractText_(blob, mime, parentFolderId) {
  if (mime === 'text/plain') return blob.getDataAsString();
  // Convert PDF / DOCX / image to a temporary Google Doc, read text, then trash it.
  var tempDoc = Drive.Files.create({
    name: '__rrtmp_' + uuid_(),
    mimeType: 'application/vnd.google-apps.document',
    parents: [parentFolderId]
  }, blob, { ocrLanguage: 'en' });
  var docId = tempDoc.id;
  var text = '';
  try { text = DocumentApp.openById(docId).getBody().getText() || ''; }
  finally { try { Drive.Files.update({ trashed: true }, docId); } catch (e) {} }
  return text;
}

// Call Gemini for structured JSON, with automatic failover across several
// models so a single overload (503) or daily-quota (429) doesn't kill it.
//   TEXT mode — Drive-OCR text is sent (cheap; generous text quota).
//   FILE mode — only when OCR returned nothing (image-only PDF). Sends bytes.
// Model list: Script Property GEMINI_MODEL (comma-separated) overrides the
// default chain. Each model is tried once; on a transient/quota error we move
// to the next model after a short pause.
function geminiExtract_(extractedText, base64, mime) {
  var key = getGeminiKey_();
  if (!key) throw new Error('No Gemini API key configured.');
  var override = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || '';
  var models = override
    ? override.split(',').map(function (s) { return s.trim(); }).filter(Boolean)
    : ['gemini-2.5-flash-lite', 'gemini-2.0-flash-lite', 'gemini-2.0-flash', 'gemini-1.5-flash'];

  var prompt = 'Extract fields from this job candidate resume. Return ONLY JSON ' +
    'with keys: name, phone, email, role. name = candidate full name. ' +
    'phone = primary phone, digits only with optional + country code. ' +
    'email = primary email. role = job title they are applying for or current designation. ' +
    'Use an empty string for any field not present.';

  var parts, mode;
  if (extractedText && extractedText.length > 60) {
    parts = [{ text: prompt + '\n\nRESUME TEXT:\n' + extractedText.slice(0, 8000) }];
    mode = 'TEXT';
  } else {
    var sendMime = mime;
    if (['application/pdf', 'image/png', 'image/jpeg', 'image/webp', 'text/plain'].indexOf(mime) === -1) sendMime = 'application/pdf';
    parts = [{ text: prompt }, { inline_data: { mime_type: sendMime, data: base64 } }];
    mode = 'FILE';
  }
  var body = JSON.stringify({ contents: [{ parts: parts }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } });

  var lastErr = '';
  for (var i = 0; i < models.length; i++) {
    var model = models[i];
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(key);
    var resp = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: body, muteHttpExceptions: true });
    var code = resp.getResponseCode();
    var raw = resp.getContentText();
    if (code === 200) {
      var json = JSON.parse(raw);
      var txt = json && json.candidates && json.candidates[0] && json.candidates[0].content &&
                json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
                json.candidates[0].content.parts[0].text;
      if (!txt) { lastErr = model + ': empty response'; continue; }
      var parsed = JSON.parse(txt);
      return {
        name: String(parsed.name || '').trim(),
        phone: String(parsed.phone || '').replace(/[^\d+]/g, ''),
        email: String(parsed.email || '').trim(),
        role: String(parsed.role || '').trim(),
        __mode: mode + ':' + model
      };
    }
    // Transient (503/500/429) or model-missing (404): record and try next model.
    lastErr = model + ' HTTP ' + code;
    if (code === 503 || code === 500 || code === 429) { Utilities.sleep(700); continue; }
    if (code === 404) { continue; }
    // Other errors (bad key, etc.) are not retryable.
    throw new Error('Gemini HTTP ' + code + ': ' + raw.slice(0, 160));
  }
  throw new Error('All Gemini models busy/over quota (' + lastErr + '). Google OCR still saved & parsed what it could — fill any gaps manually.');
}

// Groq (OpenAI-compatible) — text-only fallback using the OCR'd resume text.
// Free tier is fast and high-limit. Model chain via Script Property GROQ_MODEL
// (comma-separated) or the default below.
function groqExtract_(extractedText) {
  var key = getGroqKey_();
  if (!key) throw new Error('No Groq API key configured.');
  if (!extractedText || extractedText.length < 60) throw new Error('No resume text for Groq to read.');
  var override = PropertiesService.getScriptProperties().getProperty('GROQ_MODEL') || '';
  var models = override
    ? override.split(',').map(function (s) { return s.trim(); }).filter(Boolean)
    : ['llama-3.1-8b-instant', 'llama-3.3-70b-versatile'];

  var sys = 'You extract fields from a job candidate resume and reply with ONLY a JSON object ' +
    'with keys name, phone, email, role. name=full name; phone=primary phone digits with optional + country code; ' +
    'email=primary email; role=job title applied for or current designation. Empty string if absent.';
  var body = JSON.stringify({
    model: '__MODEL__',
    messages: [{ role: 'system', content: sys }, { role: 'user', content: extractedText.slice(0, 8000) }],
    temperature: 0,
    response_format: { type: 'json_object' }
  });

  var lastErr = '';
  for (var i = 0; i < models.length; i++) {
    var payload = body.replace('__MODEL__', models[i]);
    var resp = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + key },
      payload: payload, muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var raw = resp.getContentText();
    if (code === 200) {
      var json = JSON.parse(raw);
      var content = json && json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content;
      if (!content) { lastErr = models[i] + ': empty'; continue; }
      var parsed = JSON.parse(content);
      return {
        name: String(parsed.name || '').trim(),
        phone: String(parsed.phone || '').replace(/[^\d+]/g, ''),
        email: String(parsed.email || '').trim(),
        role: String(parsed.role || '').trim(),
        __mode: 'GROQ:' + models[i]
      };
    }
    lastErr = models[i] + ' HTTP ' + code;
    if (code === 503 || code === 500 || code === 429) { Utilities.sleep(700); continue; }
    if (code === 404) { continue; }
    throw new Error('Groq HTTP ' + code + ': ' + raw.slice(0, 160));
  }
  throw new Error('All Groq models busy/over quota (' + lastErr + ').');
}

function parseResume_(text, fileNameHint) {
  var lines = (text || '').split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
  var lower = (text || '').toLowerCase();

  // ---- EMAIL: prefer a "Email:" labelled value, else first match ----
  var email = pick_(text, /e-?mail\s*[:\-]?\s*([\w.+-]+@[\w-]+\.[\w.-]+)/i) ||
              ((text || '').match(/[\w.+-]+@[\w-]+\.[\w.-]+/) || [''])[0];
  email = email.replace(/[.,;]+$/, '');

  // ---- PHONE: prefer a labelled value, else Indian, else any long run ----
  var phone = pick_(text, /(?:mobile|phone|contact|mob|cell|tel|ph)\s*(?:no\.?|number)?\s*[:\-]?\s*(\+?[\d][\d\s().-]{7,}\d)/i);
  if (!phone) { var ind = (text || '').match(/(?:\+?91[\s-]?)?[6-9]\d{9}/); if (ind) phone = ind[0]; }
  if (!phone) { var gen = (text || '').match(/\+?\d[\d\s().-]{8,}\d/); if (gen) phone = gen[0]; }
  phone = phone ? phone.replace(/[^\d+]/g, '') : '';

  // ---- ROLE: only from an explicit label followed by a colon (avoids
  // grabbing stray words like "Skills"/"Applied" from body text) ----
  var role = pick_(text, /(?:position|designation|job\s*title|role\s*applied|applied\s*for|applying\s*for)\b[^:\n]{0,15}:\s*([A-Za-z][A-Za-z &/.\-]{2,40})/i);
  role = role ? role.replace(/\s+/g, ' ').trim() : '';

  // ---- NAME: labelled "Name:" > header line heuristic > filename ----
  var name = pick_(text, /\bname\s*[:\-]\s*([A-Z][A-Za-z'.\-]+(?:\s+[A-Z][A-Za-z'.\-]+){1,4})/);
  if (name && isNameLike_(name)) name = titleCase_(name); else name = '';
  if (!name) {
    var skip = /^(resume|curriculum vitae|cv|profile|objective|summary|contact|personal details|career|address|phone|email|mobile|name|nationality|date of birth)\b/i;
    for (var i = 0; i < Math.min(18, lines.length); i++) {
      var l = lines[i].replace(/^(name|naam)\s*[:\-]\s*/i, '').trim();
      if (l.length < 3 || l.length > 50) continue;
      if (skip.test(l)) continue;
      if (/[@\d]|https?:|www\./i.test(l)) continue;
      if (isNameLike_(l)) { name = titleCase_(l); break; }
    }
  }
  if (!name && fileNameHint) name = nameFromFilename_(fileNameHint);

  return { name: name, phone: phone, email: email, role: role };
}

// Helpers for parsing.
function pick_(text, re) { var m = (text || '').match(re); return m && m[1] ? m[1].trim() : ''; }
function isNameLike_(s) {
  var words = s.trim().split(/\s+/);
  if (words.length < 2 || words.length > 5) return false;
  var compact = s.replace(/\s/g, '');
  if ((compact.match(/[A-Za-z]/g) || []).length / compact.length < 0.85) return false;
  return words.every(function (w) { return /^[A-Z][a-z]+(['.-][A-Za-z]+)*$|^[A-Z][A-Z'.-]+$/.test(w); });
}
function titleCase_(s) {
  return s.trim().split(/\s+/).map(function (w) {
    return w === w.toUpperCase() ? (w.charAt(0) + w.slice(1).toLowerCase()) : w;
  }).join(' ');
}
// Resumes are very often named "Arjun Nair Resume.pdf" / "arjun_nair_cv.pdf".
function nameFromFilename_(fname) {
  var base = String(fname).replace(/\.[^.]+$/, '');
  base = base.replace(/[_\-.]+/g, ' ')
             .replace(/\b(resume|cv|curriculum vitae|final|updated?|copy|new|profile|biodata|naukri|indeed)\b/ig, ' ')
             .replace(/\d+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!base) return '';
  var words = base.split(/\s+/).filter(function (w) { return /^[A-Za-z][A-Za-z'.-]*$/.test(w); });
  if (words.length < 2 || words.length > 4) return '';
  // Force proper capitalisation (filenames are often all-lowercase).
  return words.map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); }).join(' ');
}

// ============================================================
// CANDIDATES
// ============================================================
function addCandidate(data) {
  var me = authorizeUser_(['Admin', 'HR']);
  if (!data || !data.Name || (!data.Phone && !data.Email)) throw new Error('Name and (Phone or Email) are required.');

  // Validate formats.
  if (data.Email && !validEmail_(data.Email)) throw new Error('Please enter a valid email address.');
  if (data.Phone && !validPhone_(data.Phone)) throw new Error('Please enter a valid phone number (7–15 digits).');

  var sheet = getSheet_(CONFIG.SHEETS.CANDIDATES);
  var existing = sheetToObjects_(sheet);
  for (var i = 0; i < existing.length; i++) {
    if (data.Email && existing[i].Email && String(existing[i].Email).toLowerCase() === String(data.Email).toLowerCase()) {
      throw new Error('A candidate with that email already exists.');
    }
    if (data.Phone && existing[i].Phone &&
        String(existing[i].Phone).replace(/[^\d]/g, '') === String(data.Phone).replace(/[^\d]/g, '')) {
      throw new Error('A candidate with that phone already exists.');
    }
    // Same resume file already attached to another candidate.
    if (data.ResumeLink && existing[i].ResumeLink && String(existing[i].ResumeLink) === String(data.ResumeLink)) {
      throw new Error('That resume is already attached to another candidate.');
    }
  }
  var id = uuid_();
  var status = (data.Status && CONFIG.STATUSES.indexOf(data.Status) !== -1) ? data.Status : 'New';
  sheet.appendRow([
    id, data.Name, data.Phone || '', data.Email || '',
    data.RoleApplied || '', data.ResumeLink || '', data.Source || '', status, nowIso_()
  ]);
  logAudit_(me.email, 'addCandidate', { id: id, name: data.Name, resume: !!data.ResumeLink });
  return { ok: true, CandidateID: id, ResumeLink: data.ResumeLink || '' };
}
function updateStatus(candidateId, newStatus) {
  var me = authorizeUser_(['Admin', 'HR']);
  if (CONFIG.STATUSES.indexOf(newStatus) === -1) throw new Error('Invalid status.');
  var sheet = getSheet_(CONFIG.SHEETS.CANDIDATES);
  var row = findRowByKey_(sheet, 0, candidateId);
  if (row === -1) throw new Error('Candidate not found.');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.getRange(row, headers.indexOf('Status') + 1).setValue(newStatus);
  logAudit_(me.email, 'updateStatus', { id: candidateId, to: newStatus });
  return { ok: true };
}

// ============================================================
// INTERVIEWS
// ============================================================
function addInterview(data) {
  var me = authorizeUser_(['Admin', 'HR']);
  if (!data || !data.CandidateID || !data.Round || !data.DateTime) throw new Error('Candidate, Round, and Date/Time are required.');
  var id = uuid_();
  getSheet_(CONFIG.SHEETS.INTERVIEWS).appendRow([
    id, data.CandidateID, data.Round, data.Interviewer || '',
    data.DateTime, data.Status || 'Scheduled', data.Feedback || '', data.Score || ''
  ]);
  logAudit_(me.email, 'addInterview', { id: id, candidate: data.CandidateID });
  return { ok: true, InterviewID: id };
}
function updateInterviewFeedback(data) {
  var me = authorizeUser_(['Admin', 'HR', 'Interviewer']);
  if (!data || !data.InterviewID) throw new Error('InterviewID required.');
  var sheet = getSheet_(CONFIG.SHEETS.INTERVIEWS);
  var row = findRowByKey_(sheet, 0, data.InterviewID);
  if (row === -1) throw new Error('Interview not found.');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (me.role === 'Interviewer') {
    var owner = sheet.getRange(row, headers.indexOf('Interviewer') + 1).getValue();
    if (String(owner).toLowerCase() !== me.email.toLowerCase()) throw new Error('You can only update your own interviews.');
  }
  ['Feedback', 'Score', 'Status'].forEach(function (f) {
    if (data[f] !== undefined) {
      var col = headers.indexOf(f);
      if (col !== -1) sheet.getRange(row, col + 1).setValue(data[f]);
    }
  });
  logAudit_(me.email, 'updateFeedback', { id: data.InterviewID });
  return { ok: true };
}

// ============================================================
// USERS (Admin)
// ============================================================
function getUsers() { authorizeUser_(['Admin']); return sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS)); }
function addUser(data) {
  var me = authorizeUser_(['Admin']);
  if (!data || !data.Email || !data.Role) throw new Error('Email and Role required.');
  if (CONFIG.ROLES.indexOf(data.Role) === -1) throw new Error('Invalid role.');
  var users = sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS));
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].Email).toLowerCase() === String(data.Email).toLowerCase()) throw new Error('User already exists.');
  }
  getSheet_(CONFIG.SHEETS.USERS).appendRow([uuid_(), data.Name || '', data.Email, data.Role]);
  logAudit_(me.email, 'addUser', { email: data.Email, role: data.Role });
  return { ok: true };
}
function deleteUser(userId) {
  var me = authorizeUser_(['Admin']);
  var sheet = getSheet_(CONFIG.SHEETS.USERS);
  var row = findRowByKey_(sheet, 0, userId);
  if (row === -1) throw new Error('User not found.');
  var email = sheet.getRange(row, 3).getValue();
  if (String(email).toLowerCase() === me.email.toLowerCase()) throw new Error('You cannot delete your own account.');
  sheet.deleteRow(row);
  logAudit_(me.email, 'deleteUser', { userId: userId });
  return { ok: true };
}

// ============================================================
// EXPORT
// ============================================================
function exportCandidatesCsv() {
  authorizeUser_(['Admin']);
  var rows = getSheet_(CONFIG.SHEETS.CANDIDATES).getDataRange().getValues();
  return rows.map(function (r) {
    return r.map(function (v) {
      var s = (v === null || v === undefined) ? '' : String(v);
      if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    }).join(',');
  }).join('\n');
}
