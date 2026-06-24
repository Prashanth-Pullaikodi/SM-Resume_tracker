/**
 * Resort Recruitment System - Backend
 * Google Apps Script Web App
 *
 * Deployment: Deploy as Web App, "Execute as: User accessing the web app",
 * "Who has access: Anyone within <your domain>" (or as appropriate).
 */

// ============================================================
// CONFIGURATION
// ============================================================
var CONFIG = {
  SPREADSHEET_ID: '', // Leave blank to use the bound spreadsheet (recommended)
  SHEETS: {
    CANDIDATES: 'Candidates',
    INTERVIEWS: 'Interviews',
    STATUS_HISTORY: 'StatusHistory',
    USERS: 'Users',
    AUDIT_LOG: 'AuditLog'
  },
  STATUSES: ['New', 'Not Screened', 'Shortlisted', 'Interviewed', 'Selected', 'Rejected', 'On Hold'],
  ROLES: ['Admin', 'HR', 'Interviewer', 'Viewer']
};

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';

  if (page === 'manifest') {
    return ContentService
      .createTextOutput(getManifestJson())
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (page === 'sw') {
    return ContentService
      .createTextOutput(getServiceWorkerJs())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  var t = HtmlService.createTemplateFromFile('Index');
  t.webAppUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle('Resort Recruitment')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .addMetaTag('theme-color', '#0d6efd')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// SHEET HELPERS
// ============================================================
function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Set CONFIG.SPREADSHEET_ID or bind script to a sheet.');
  return ss;
}

function getSheet_(name) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run setupSheets().');
  return sh;
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) row[headers[j]] = values[i][j];
    if (row[headers[0]] === '' || row[headers[0]] === null) continue;
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

function uuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

// ============================================================
// SETUP (run once)
// ============================================================
function setupSheets() {
  var ss = getSpreadsheet_();
  var defs = {
    'Candidates': ['CandidateID','Name','Phone','Email','RoleApplied','ResumeLink','Source','Status','CreatedAt'],
    'Interviews': ['InterviewID','CandidateID','Round','Interviewer','DateTime','Status','Feedback','Score'],
    'StatusHistory': ['LogID','CandidateID','OldStatus','NewStatus','ChangedBy','Timestamp'],
    'Users': ['UserID','Name','Email','Role'],
    'AuditLog': ['LogID','Email','Action','Details','Timestamp']
  };
  Object.keys(defs).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    var existing = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    if (existing.join('') === '') {
      sh.getRange(1, 1, 1, defs[name].length).setValues([defs[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
    }
  });

  // Seed first admin user with the current account if Users sheet empty
  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet.getLastRow() < 2) {
    var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    if (email) {
      usersSheet.appendRow([uuid_(), 'Admin User', email, 'Admin']);
    }
  }
  return 'Setup complete.';
}

// ============================================================
// AUTH / RBAC
// ============================================================
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  if (!email) email = Session.getEffectiveUser().getEmail();
  if (!email) return { authorized: false, message: 'Unable to detect Google account. Please sign in.' };

  var role = getUserRole(email);
  if (!role) {
    return { authorized: false, email: email, message: 'Access Denied. Your email is not registered. Contact admin.' };
  }
  var users = sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS));
  var me = users.filter(function(u){ return String(u.Email).toLowerCase() === email.toLowerCase(); })[0];
  return {
    authorized: true,
    email: email,
    name: me ? me.Name : email,
    role: role,
    permissions: rolePermissions_(role)
  };
}

function getUserRole(email) {
  if (!email) return null;
  var users = sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS));
  var match = users.filter(function(u){ return String(u.Email).toLowerCase() === String(email).toLowerCase(); })[0];
  return match ? String(match.Role) : null;
}

function authorizeUser(requiredRoles) {
  var info = getCurrentUser();
  if (!info.authorized) throw new Error(info.message || 'Access Denied');
  if (requiredRoles && requiredRoles.length && requiredRoles.indexOf(info.role) === -1) {
    throw new Error('Access Denied: requires role ' + requiredRoles.join(' or '));
  }
  return info;
}

function rolePermissions_(role) {
  switch (role) {
    case 'Admin':
      return { manageUsers: true, addCandidate: true, editCandidate: true, scheduleInterview: true,
               addFeedback: true, changeStatus: true, exportCsv: true, viewAll: true };
    case 'HR':
      return { manageUsers: false, addCandidate: true, editCandidate: true, scheduleInterview: true,
               addFeedback: true, changeStatus: true, exportCsv: false, viewAll: true };
    case 'Interviewer':
      return { manageUsers: false, addCandidate: false, editCandidate: false, scheduleInterview: false,
               addFeedback: true, changeStatus: false, exportCsv: false, viewAll: false };
    case 'Viewer':
      return { manageUsers: false, addCandidate: false, editCandidate: false, scheduleInterview: false,
               addFeedback: false, changeStatus: false, exportCsv: false, viewAll: true };
    default:
      return {};
  }
}

function logAudit_(email, action, details) {
  try {
    getSheet_(CONFIG.SHEETS.AUDIT_LOG)
      .appendRow([uuid_(), email, action, JSON.stringify(details || {}), nowIso_()]);
  } catch (e) { /* swallow */ }
}

// ============================================================
// CANDIDATES
// ============================================================
function getCandidates() {
  var me = authorizeUser();
  var all = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  if (me.role === 'Interviewer') {
    // Only return candidates with interviews assigned to this user
    var interviews = sheetToObjects_(getSheet_(CONFIG.SHEETS.INTERVIEWS));
    var assignedIds = {};
    interviews.forEach(function(i){
      if (String(i.Interviewer).toLowerCase() === me.email.toLowerCase()) assignedIds[i.CandidateID] = true;
    });
    all = all.filter(function(c){ return assignedIds[c.CandidateID]; });
  }
  return all;
}

function getCandidate(candidateId) {
  authorizeUser();
  var all = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  return all.filter(function(c){ return c.CandidateID === candidateId; })[0] || null;
}

function addCandidate(data) {
  var me = authorizeUser(['Admin','HR']);

  if (!data || !data.Name || (!data.Phone && !data.Email)) {
    throw new Error('Name and Phone or Email are required.');
  }

  var sheet = getSheet_(CONFIG.SHEETS.CANDIDATES);
  var existing = sheetToObjects_(sheet);

  // Duplicate detection by email or phone
  for (var i = 0; i < existing.length; i++) {
    if (data.Email && existing[i].Email && String(existing[i].Email).toLowerCase() === String(data.Email).toLowerCase()) {
      throw new Error('Duplicate candidate: email already exists.');
    }
    if (data.Phone && existing[i].Phone && String(existing[i].Phone) === String(data.Phone)) {
      throw new Error('Duplicate candidate: phone already exists.');
    }
  }

  var id = uuid_();
  var status = data.Status && CONFIG.STATUSES.indexOf(data.Status) !== -1 ? data.Status : 'New';
  sheet.appendRow([
    id, data.Name, data.Phone || '', data.Email || '',
    data.RoleApplied || '', data.ResumeLink || '', data.Source || '',
    status, nowIso_()
  ]);
  logAudit_(me.email, 'addCandidate', { id: id, name: data.Name });
  return { ok: true, CandidateID: id };
}

function updateCandidate(data) {
  var me = authorizeUser(['Admin','HR']);
  if (!data || !data.CandidateID) throw new Error('CandidateID required.');
  var sheet = getSheet_(CONFIG.SHEETS.CANDIDATES);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var row = findRowByKey_(sheet, 0, data.CandidateID);
  if (row === -1) throw new Error('Candidate not found.');

  ['Name','Phone','Email','RoleApplied','ResumeLink','Source'].forEach(function(field){
    if (data[field] !== undefined) {
      var col = headers.indexOf(field);
      if (col !== -1) sheet.getRange(row, col + 1).setValue(data[field]);
    }
  });
  logAudit_(me.email, 'updateCandidate', { id: data.CandidateID });
  return { ok: true };
}

function updateStatus(candidateId, newStatus) {
  var me = authorizeUser(['Admin','HR']);
  if (CONFIG.STATUSES.indexOf(newStatus) === -1) throw new Error('Invalid status.');

  var sheet = getSheet_(CONFIG.SHEETS.CANDIDATES);
  var row = findRowByKey_(sheet, 0, candidateId);
  if (row === -1) throw new Error('Candidate not found.');
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var statusCol = headers.indexOf('Status') + 1;
  var oldStatus = sheet.getRange(row, statusCol).getValue();
  sheet.getRange(row, statusCol).setValue(newStatus);

  getSheet_(CONFIG.SHEETS.STATUS_HISTORY)
    .appendRow([uuid_(), candidateId, oldStatus, newStatus, me.email, nowIso_()]);
  logAudit_(me.email, 'updateStatus', { id: candidateId, from: oldStatus, to: newStatus });
  return { ok: true };
}

function getNotScreenedCandidates() {
  authorizeUser();
  return sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES))
    .filter(function(c){ return c.Status === 'Not Screened'; });
}

// ============================================================
// INTERVIEWS
// ============================================================
function addInterview(data) {
  var me = authorizeUser(['Admin','HR']);
  if (!data || !data.CandidateID || !data.Round || !data.DateTime) {
    throw new Error('CandidateID, Round, and DateTime are required.');
  }
  var id = uuid_();
  getSheet_(CONFIG.SHEETS.INTERVIEWS).appendRow([
    id, data.CandidateID, data.Round, data.Interviewer || '',
    data.DateTime, data.Status || 'Scheduled', data.Feedback || '', data.Score || ''
  ]);
  logAudit_(me.email, 'addInterview', { id: id, candidate: data.CandidateID });
  return { ok: true, InterviewID: id };
}

function getInterviews() {
  var me = authorizeUser();
  var all = sheetToObjects_(getSheet_(CONFIG.SHEETS.INTERVIEWS));
  if (me.role === 'Interviewer') {
    all = all.filter(function(i){
      return String(i.Interviewer).toLowerCase() === me.email.toLowerCase();
    });
  }
  return all;
}

function getInterviewsByCandidate(candidateId) {
  authorizeUser();
  return sheetToObjects_(getSheet_(CONFIG.SHEETS.INTERVIEWS))
    .filter(function(i){ return i.CandidateID === candidateId; });
}

function updateInterviewFeedback(data) {
  var me = authorizeUser(['Admin','HR','Interviewer']);
  if (!data || !data.InterviewID) throw new Error('InterviewID required.');
  var sheet = getSheet_(CONFIG.SHEETS.INTERVIEWS);
  var row = findRowByKey_(sheet, 0, data.InterviewID);
  if (row === -1) throw new Error('Interview not found.');
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  if (me.role === 'Interviewer') {
    var interviewerCol = headers.indexOf('Interviewer');
    var existing = sheet.getRange(row, interviewerCol + 1).getValue();
    if (String(existing).toLowerCase() !== me.email.toLowerCase()) {
      throw new Error('Interviewers can only update their own interviews.');
    }
  }

  ['Feedback','Score','Status'].forEach(function(field){
    if (data[field] !== undefined) {
      var col = headers.indexOf(field);
      if (col !== -1) sheet.getRange(row, col + 1).setValue(data[field]);
    }
  });
  logAudit_(me.email, 'updateInterviewFeedback', { id: data.InterviewID });
  return { ok: true };
}

// ============================================================
// USERS (Admin only)
// ============================================================
function getUsers() {
  authorizeUser(['Admin']);
  return sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS));
}

function validateUser(email) {
  return !!getUserRole(email);
}

function addUser(data) {
  var me = authorizeUser(['Admin']);
  if (!data || !data.Email || !data.Role) throw new Error('Email and Role required.');
  if (CONFIG.ROLES.indexOf(data.Role) === -1) throw new Error('Invalid role.');
  if (getUserRole(data.Email)) throw new Error('User already exists.');
  getSheet_(CONFIG.SHEETS.USERS)
    .appendRow([uuid_(), data.Name || '', data.Email, data.Role]);
  logAudit_(me.email, 'addUser', { email: data.Email, role: data.Role });
  return { ok: true };
}

function deleteUser(userId) {
  var me = authorizeUser(['Admin']);
  var sheet = getSheet_(CONFIG.SHEETS.USERS);
  var row = findRowByKey_(sheet, 0, userId);
  if (row === -1) throw new Error('User not found.');
  var email = sheet.getRange(row, 3).getValue();
  if (String(email).toLowerCase() === me.email.toLowerCase()) {
    throw new Error('You cannot delete your own account.');
  }
  sheet.deleteRow(row);
  logAudit_(me.email, 'deleteUser', { userId: userId, email: email });
  return { ok: true };
}

// ============================================================
// ANALYTICS
// ============================================================
function getDashboardKpis() {
  authorizeUser();
  var candidates = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  var by = {};
  CONFIG.STATUSES.forEach(function(s){ by[s] = 0; });
  candidates.forEach(function(c){ if (by[c.Status] !== undefined) by[c.Status]++; });
  return {
    total: candidates.length,
    new: by['New'],
    notScreened: by['Not Screened'],
    shortlisted: by['Shortlisted'],
    interviewed: by['Interviewed'],
    selected: by['Selected'],
    rejected: by['Rejected'],
    onHold: by['On Hold']
  };
}

function getConversionMetrics() {
  authorizeUser();
  var candidates = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  var total = candidates.length || 1;
  var counts = { new: 0, notScreened: 0, screened: 0, interviewed: 0, selected: 0, rejected: 0 };
  candidates.forEach(function(c){
    switch (c.Status) {
      case 'New': counts.new++; break;
      case 'Not Screened': counts.notScreened++; break;
      case 'Shortlisted': counts.screened++; break;
      case 'Interviewed': counts.interviewed++; break;
      case 'Selected': counts.selected++; break;
      case 'Rejected': counts.rejected++; break;
    }
  });
  var screenedOrBeyond = counts.screened + counts.interviewed + counts.selected + counts.rejected;
  var interviewedOrBeyond = counts.interviewed + counts.selected;
  return {
    totalCandidates: candidates.length,
    screeningRate: Math.round((screenedOrBeyond / total) * 100),
    interviewConversionRate: interviewedOrBeyond ? Math.round((counts.selected / interviewedOrBeyond) * 100) : 0,
    rejectionRate: Math.round((counts.rejected / total) * 100),
    funnel: {
      new: candidates.length,
      screened: screenedOrBeyond,
      interviewed: interviewedOrBeyond,
      selected: counts.selected
    }
  };
}

function getRoleWiseStats() {
  authorizeUser();
  var candidates = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  var byRole = {};
  candidates.forEach(function(c){
    var r = c.RoleApplied || 'Unspecified';
    if (!byRole[r]) byRole[r] = { role: r, total: 0, selected: 0, rejected: 0 };
    byRole[r].total++;
    if (c.Status === 'Selected') byRole[r].selected++;
    if (c.Status === 'Rejected') byRole[r].rejected++;
  });
  return Object.keys(byRole).map(function(k){ return byRole[k]; });
}

// ============================================================
// EXPORT
// ============================================================
function exportCandidatesCsv() {
  authorizeUser(['Admin']);
  var rows = getSheet_(CONFIG.SHEETS.CANDIDATES).getDataRange().getValues();
  var csv = rows.map(function(r){
    return r.map(function(v){
      var s = (v === null || v === undefined) ? '' : String(v);
      if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
        s = '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\n');
  return csv;
}

// ============================================================
// MANIFEST AND SERVICE WORKER (served via doGet)
// ============================================================
function getManifestJson() {
  var url = ScriptApp.getService().getUrl();
  return JSON.stringify({
    name: 'Resort Recruitment',
    short_name: 'Recruit',
    description: 'Resort Recruitment Management System',
    start_url: url,
    scope: url,
    display: 'standalone',
    orientation: 'portrait',
    background_color: '#ffffff',
    theme_color: '#0d6efd',
    icons: [
      { src: 'https://www.gstatic.com/images/branding/product/2x/forms_2020q4_48dp.png', sizes: '192x192', type: 'image/png' },
      { src: 'https://www.gstatic.com/images/branding/product/2x/forms_2020q4_48dp.png', sizes: '512x512', type: 'image/png' }
    ]
  });
}

function getServiceWorkerJs() {
  return [
    "const CACHE = 'recruit-shell-v1';",
    "const SHELL = [self.registration.scope];",
    "self.addEventListener('install', e => {",
    "  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(()=>{}));",
    "  self.skipWaiting();",
    "});",
    "self.addEventListener('activate', e => {",
    "  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));",
    "  self.clients.claim();",
    "});",
    "self.addEventListener('fetch', e => {",
    "  if (e.request.method !== 'GET') return;",
    "  e.respondWith(",
    "    fetch(e.request).then(res => {",
    "      const copy = res.clone();",
    "      caches.open(CACHE).then(c => c.put(e.request, copy)).catch(()=>{});",
    "      return res;",
    "    }).catch(() => caches.match(e.request).then(r => r || caches.match(self.registration.scope)))",
    "  );",
    "});"
  ].join('\n');
}
