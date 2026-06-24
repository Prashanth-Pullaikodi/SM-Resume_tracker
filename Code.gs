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
  ROLES: ['Admin', 'HR', 'Interviewer', 'Viewer'],

  // ---- Drive upload settings --------------------------------------------
  // Paste the ID of the Drive folder where all resumes should be saved.
  // Get it from the folder URL: https://drive.google.com/drive/folders/<THIS_PART>
  // If left blank, a folder named RESUMES_FOLDER_NAME is auto-created in
  // the script owner's "My Drive" root on first upload.
  RESUMES_FOLDER_ID: '',
  RESUMES_FOLDER_NAME: 'Resort Recruitment Resumes',
  MAX_UPLOAD_BYTES: 25 * 1024 * 1024, // 25 MB hard cap per resume

  // Cache TTL for the bootstrap payload (seconds). Writes invalidate the
  // cache, so this is just an upper bound for stale-reads from other tabs.
  CACHE_TTL_SECONDS: 300,
  CACHE_KEY_BOOT: 'boot_v2',
  CACHE_KEY_USER_ROLE: 'role_v1_'
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
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
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
    for (var j = 0; j < headers.length; j++) {
      var v = values[i][j];
      // Coerce Date cells to ISO strings — google.script.run sometimes
      // serializes Date objects unreliably across the postMessage bridge,
      // which can result in an empty/undefined response on the client.
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = isNaN(v.getTime()) ? '' : v.toISOString();
      }
      row[headers[j]] = v;
    }
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
// INITIAL SETUP — run once from the Apps Script editor.
// Creates sheets + headers, seeds you as Admin, then populates
// sample users, candidates, interviews, and status history so the
// dashboard and analytics screens have something to show.
// Safe to re-run: it skips seeding if data already exists.
// ============================================================
function initialSetup() {
  setupSheets();
  var ss = getSpreadsheet_();

  var usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  var candSheet  = ss.getSheetByName(CONFIG.SHEETS.CANDIDATES);
  var ivSheet    = ss.getSheetByName(CONFIG.SHEETS.INTERVIEWS);
  var histSheet  = ss.getSheetByName(CONFIG.SHEETS.STATUS_HISTORY);
  var auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);

  var adminEmail = Session.getActiveUser().getEmail() ||
                   Session.getEffectiveUser().getEmail() ||
                   'admin@example.com';

  // -------- Sample Users (skip if already populated beyond the admin row) --------
  if (usersSheet.getLastRow() < 3) {
    var sampleUsers = [
      [uuid_(), 'Priya HR',        'priya.hr@example.com',        'HR'],
      [uuid_(), 'Rahul Manager',   'rahul.manager@example.com',   'Interviewer'],
      [uuid_(), 'Anita Chef',      'anita.chef@example.com',      'Interviewer'],
      [uuid_(), 'Vikram Viewer',   'vikram.viewer@example.com',   'Viewer']
    ];
    usersSheet.getRange(usersSheet.getLastRow() + 1, 1, sampleUsers.length, 4)
              .setValues(sampleUsers);
  }

  // -------- Sample Candidates --------
  if (candSheet.getLastRow() < 2) {
    var now = new Date();
    function daysAgo(n) {
      var d = new Date(now.getTime() - n * 24 * 60 * 60 * 1000);
      return d.toISOString();
    }
    var candIds = {
      arjun:   uuid_(),
      meera:   uuid_(),
      sandeep: uuid_(),
      kavya:   uuid_(),
      rohit:   uuid_(),
      neha:    uuid_(),
      farhan:  uuid_(),
      divya:   uuid_(),
      kiran:   uuid_(),
      sneha:   uuid_()
    };
    var sampleCandidates = [
      [candIds.arjun,   'Arjun Nair',       '+919812340001', 'arjun.nair@example.com',    'Front Desk Executive', 'https://drive.google.com/file/d/sample-arjun',  'Naukri',   'Selected',     daysAgo(40)],
      [candIds.meera,   'Meera Iyer',       '+919812340002', 'meera.iyer@example.com',    'Housekeeping Supervisor','https://drive.google.com/file/d/sample-meera', 'Referral', 'Interviewed',  daysAgo(25)],
      [candIds.sandeep, 'Sandeep Kumar',    '+919812340003', 'sandeep.k@example.com',     'Chef de Partie',       'https://drive.google.com/file/d/sample-sandeep','Walk-in',  'Shortlisted',  daysAgo(15)],
      [candIds.kavya,   'Kavya Reddy',      '+919812340004', 'kavya.reddy@example.com',   'Spa Therapist',        'https://drive.google.com/file/d/sample-kavya',  'LinkedIn', 'Not Screened', daysAgo(8)],
      [candIds.rohit,   'Rohit Sharma',     '+919812340005', 'rohit.sharma@example.com',  'F&B Manager',          'https://drive.google.com/file/d/sample-rohit',  'Referral', 'Rejected',     daysAgo(35)],
      [candIds.neha,    'Neha Pillai',      '+919812340006', 'neha.pillai@example.com',   'Front Desk Executive', 'https://drive.google.com/file/d/sample-neha',   'Naukri',   'Not Screened', daysAgo(3)],
      [candIds.farhan,  'Farhan Ahmed',     '+919812340007', 'farhan.a@example.com',      'Sous Chef',            'https://drive.google.com/file/d/sample-farhan', 'Indeed',   'New',          daysAgo(2)],
      [candIds.divya,   'Divya Menon',      '+919812340008', 'divya.menon@example.com',   'Guest Relations',      'https://drive.google.com/file/d/sample-divya',  'Walk-in',  'On Hold',      daysAgo(20)],
      [candIds.kiran,   'Kiran Bose',       '+919812340009', 'kiran.bose@example.com',    'Housekeeping Supervisor','https://drive.google.com/file/d/sample-kiran', 'Referral', 'Interviewed',  daysAgo(18)],
      [candIds.sneha,   'Sneha Krishnan',   '+919812340010', 'sneha.k@example.com',       'Spa Therapist',        'https://drive.google.com/file/d/sample-sneha',  'LinkedIn', 'New',          daysAgo(1)]
    ];
    candSheet.getRange(candSheet.getLastRow() + 1, 1, sampleCandidates.length, 9)
             .setValues(sampleCandidates);

    // -------- Sample Interviews (linked to the candidates above) --------
    function inDays(n) {
      var d = new Date(now.getTime() + n * 24 * 60 * 60 * 1000);
      return d.toISOString();
    }
    var sampleInterviews = [
      [uuid_(), candIds.arjun,   'HR',        'priya.hr@example.com',      daysAgo(35), 'Completed', 'Excellent communication, strong hospitality fit.', 9],
      [uuid_(), candIds.arjun,   'Manager',   'rahul.manager@example.com', daysAgo(30), 'Completed', 'Recommended for selection.',                       9],
      [uuid_(), candIds.meera,   'HR',        'priya.hr@example.com',      daysAgo(20), 'Completed', 'Solid housekeeping experience.',                   8],
      [uuid_(), candIds.meera,   'Manager',   'rahul.manager@example.com', daysAgo(15), 'Completed', 'Awaiting reference check.',                        7],
      [uuid_(), candIds.sandeep, 'Technical', 'anita.chef@example.com',    daysAgo(10), 'Completed', 'Good knife skills, needs training in plating.',     7],
      [uuid_(), candIds.rohit,   'HR',        'priya.hr@example.com',      daysAgo(32), 'Completed', 'Did not meet salary expectations.',                 5],
      [uuid_(), candIds.kiran,   'HR',        'priya.hr@example.com',      daysAgo(14), 'Completed', 'Polite, organized.',                                8],
      [uuid_(), candIds.kiran,   'Manager',   'rahul.manager@example.com', daysAgo(12), 'Completed', 'Shortlisted pending manager approval.',             8],
      [uuid_(), candIds.farhan,  'Technical', 'anita.chef@example.com',    inDays(2),   'Scheduled', '',                                                  ''],
      [uuid_(), candIds.sneha,   'HR',        'priya.hr@example.com',      inDays(3),   'Scheduled', '',                                                  '']
    ];
    ivSheet.getRange(ivSheet.getLastRow() + 1, 1, sampleInterviews.length, 8)
           .setValues(sampleInterviews);

    // -------- Sample Status History --------
    var sampleHistory = [
      [uuid_(), candIds.arjun,   'New',          'Shortlisted',  'priya.hr@example.com', daysAgo(38)],
      [uuid_(), candIds.arjun,   'Shortlisted',  'Interviewed',  'priya.hr@example.com', daysAgo(30)],
      [uuid_(), candIds.arjun,   'Interviewed',  'Selected',     adminEmail,             daysAgo(28)],
      [uuid_(), candIds.meera,   'New',          'Shortlisted',  'priya.hr@example.com', daysAgo(22)],
      [uuid_(), candIds.meera,   'Shortlisted',  'Interviewed',  'priya.hr@example.com', daysAgo(15)],
      [uuid_(), candIds.rohit,   'Interviewed',  'Rejected',     'priya.hr@example.com', daysAgo(31)],
      [uuid_(), candIds.kiran,   'New',          'Shortlisted',  'priya.hr@example.com', daysAgo(16)],
      [uuid_(), candIds.kiran,   'Shortlisted',  'Interviewed',  'priya.hr@example.com', daysAgo(12)]
    ];
    histSheet.getRange(histSheet.getLastRow() + 1, 1, sampleHistory.length, 6)
             .setValues(sampleHistory);

    // -------- Audit Log entry --------
    auditSheet.appendRow([uuid_(), adminEmail, 'initialSetup',
                          JSON.stringify({ seeded: { users: 4, candidates: 10, interviews: 10 } }),
                          nowIso_()]);

    return 'Initial setup complete. Sample data seeded.';
  }

  return 'Initial setup complete. Sheets ready (sample data already present, skipped).';
}

// ============================================================
// AUTH / RBAC
// ============================================================
/**
 * One-shot bootstrap that returns the user + all dashboards' data in a
 * single google.script.run round trip. Cuts navigation latency by ~3x
 * because the frontend no longer needs a separate call per view.
 */
// ---- Cache helpers ----------------------------------------------------
function getCache_() { return CacheService.getScriptCache(); }

function invalidateBootCache_() {
  try { getCache_().remove(CONFIG.CACHE_KEY_BOOT); } catch (e) {}
}

function bootstrapApp() {
  try {
    // Serve from cache when present — turns ~1–3s sheet reads into <100ms.
    var cache = getCache_();
    var cached = cache.get(CONFIG.CACHE_KEY_BOOT);
    if (cached) {
      try {
        var payload = JSON.parse(cached);
        // Still re-check the user; permissions could have changed.
        payload.user = getCurrentUser();
        if (!payload.user.authorized) return { user: payload.user };
        payload.cached = true;
        return payload;
      } catch (e) { /* fall through to fresh read */ }
    }

    var fresh = bootstrapAppImpl_();
    if (fresh && fresh.user && fresh.user.authorized) {
      try {
        var json = JSON.stringify(fresh);
        // CacheService caps each value at 100 KB. Skip silently above that.
        if (json.length < 95000) cache.put(CONFIG.CACHE_KEY_BOOT, json, CONFIG.CACHE_TTL_SECONDS);
      } catch (e) {}
    }
    return fresh;
  } catch (e) {
    // Surface server-side errors as an authorized=false payload so the
    // frontend can render a useful message instead of going blank.
    return {
      user: {
        authorized: false,
        message: 'Server error: ' + (e && e.message ? e.message : e) +
                 '. Run initialSetup() from the Apps Script editor.'
      }
    };
  }
}

function bootstrapAppImpl_() {
  var user = getCurrentUser();
  if (!user.authorized) return { user: user };

  var candidates = sheetToObjects_(getSheet_(CONFIG.SHEETS.CANDIDATES));
  var interviews = sheetToObjects_(getSheet_(CONFIG.SHEETS.INTERVIEWS));

  // RBAC: Interviewers only see their assigned candidates / interviews.
  if (user.role === 'Interviewer') {
    var emailLc = String(user.email).toLowerCase();
    var assignedIds = {};
    interviews = interviews.filter(function(i){
      var keep = String(i.Interviewer).toLowerCase() === emailLc;
      if (keep) assignedIds[i.CandidateID] = true;
      return keep;
    });
    candidates = candidates.filter(function(c){ return assignedIds[c.CandidateID]; });
  }

  // Build KPIs and analytics from already-loaded data — no extra sheet reads.
  var statusCounts = {};
  CONFIG.STATUSES.forEach(function(s){ statusCounts[s] = 0; });
  var roleMap = {};
  candidates.forEach(function(c){
    if (statusCounts[c.Status] !== undefined) statusCounts[c.Status]++;
    var r = c.RoleApplied || 'Unspecified';
    if (!roleMap[r]) roleMap[r] = { role: r, total: 0, selected: 0, rejected: 0 };
    roleMap[r].total++;
    if (c.Status === 'Selected') roleMap[r].selected++;
    if (c.Status === 'Rejected') roleMap[r].rejected++;
  });

  var total = candidates.length || 1;
  var screenedOrBeyond = statusCounts['Shortlisted'] + statusCounts['Interviewed'] +
                         statusCounts['Selected'] + statusCounts['Rejected'];
  var interviewedOrBeyond = statusCounts['Interviewed'] + statusCounts['Selected'];

  return {
    user: user,
    candidates: candidates,
    interviews: interviews,
    kpis: {
      total: candidates.length,
      new: statusCounts['New'],
      notScreened: statusCounts['Not Screened'],
      shortlisted: statusCounts['Shortlisted'],
      interviewed: statusCounts['Interviewed'],
      selected: statusCounts['Selected'],
      rejected: statusCounts['Rejected'],
      onHold: statusCounts['On Hold']
    },
    metrics: {
      totalCandidates: candidates.length,
      screeningRate: Math.round((screenedOrBeyond / total) * 100),
      interviewConversionRate: interviewedOrBeyond
        ? Math.round((statusCounts['Selected'] / interviewedOrBeyond) * 100) : 0,
      rejectionRate: Math.round((statusCounts['Rejected'] / total) * 100),
      funnel: {
        new: candidates.length,
        screened: screenedOrBeyond,
        interviewed: interviewedOrBeyond,
        selected: statusCounts['Selected']
      }
    },
    roleStats: Object.keys(roleMap).map(function(k){ return roleMap[k]; }),
    serverTime: nowIso_()
  };
}

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
  var cache = getCache_();
  var key = CONFIG.CACHE_KEY_USER_ROLE + String(email).toLowerCase();
  var cached = cache.get(key);
  if (cached !== null) return cached === '__none__' ? null : cached;

  var users = sheetToObjects_(getSheet_(CONFIG.SHEETS.USERS));
  var match = users.filter(function(u){ return String(u.Email).toLowerCase() === String(email).toLowerCase(); })[0];
  var role = match ? String(match.Role) : null;
  try { cache.put(key, role || '__none__', 600); } catch (e) {}
  return role;
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
  // Every successful write goes through here, so it's the natural
  // place to invalidate the bootstrap cache.
  invalidateBootCache_();
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

  // If a resume payload was attached, save it to Drive and use the URL.
  var resumeLink = data.ResumeLink || '';
  if (data.ResumeUpload && data.ResumeUpload.base64 && data.ResumeUpload.fileName) {
    var saved = saveResumeToDrive_(id, data.Name, data.ResumeUpload);
    resumeLink = saved.url;
  }

  sheet.appendRow([
    id, data.Name, data.Phone || '', data.Email || '',
    data.RoleApplied || '', resumeLink, data.Source || '',
    status, nowIso_()
  ]);
  logAudit_(me.email, 'addCandidate', { id: id, name: data.Name, resume: !!resumeLink });
  return { ok: true, CandidateID: id, ResumeLink: resumeLink };
}

// ============================================================
// RESUME UPLOAD → DRIVE
// ============================================================
function getOrCreateResumesFolder_() {
  // 1. Explicit folder ID wins (recommended for shared drives / shared folders).
  if (CONFIG.RESUMES_FOLDER_ID) {
    try {
      return DriveApp.getFolderById(CONFIG.RESUMES_FOLDER_ID);
    } catch (e) {
      throw new Error('Configured RESUMES_FOLDER_ID is invalid or inaccessible: ' +
                      CONFIG.RESUMES_FOLDER_ID);
    }
  }
  // 2. Otherwise reuse a folder of the configured name from My Drive.
  var folders = DriveApp.getFoldersByName(CONFIG.RESUMES_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  // 3. Last resort — create it.
  return DriveApp.createFolder(CONFIG.RESUMES_FOLDER_NAME);
}

function safeFileName_(name) {
  return String(name || 'resume').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80);
}

function saveResumeToDrive_(candidateId, candidateName, upload) {
  if (!upload || !upload.base64 || !upload.fileName) {
    throw new Error('Missing resume upload payload.');
  }
  // Strip any data: URL prefix the frontend may have included.
  var b64 = String(upload.base64).replace(/^data:[^;]+;base64,/, '');
  var bytes;
  try { bytes = Utilities.base64Decode(b64); }
  catch (e) { throw new Error('Invalid file data.'); }

  if (bytes.length > CONFIG.MAX_UPLOAD_BYTES) {
    throw new Error('File too large. Limit is ' +
      Math.round(CONFIG.MAX_UPLOAD_BYTES / 1024 / 1024) + ' MB.');
  }

  var mime = upload.mimeType || 'application/octet-stream';
  var ext = (upload.fileName.match(/\.[^.]+$/) || [''])[0];
  var clean = safeFileName_(candidateName) + '_' +
              String(candidateId).slice(0, 8) + '_' +
              Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') +
              ext;

  var blob = Utilities.newBlob(bytes, mime, clean);
  var folder = getOrCreateResumesFolder_();
  var file = folder.createFile(blob);

  // Make the file accessible to anyone in the org / with the link.
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Workspace policy may block ANYONE_WITH_LINK; fall back to domain.
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) { /* keep default sharing */ }
  }

  return { url: file.getUrl(), id: file.getId(), name: clean };
}

function uploadResumeForCandidate(candidateId, upload) {
  var me = authorizeUser(['Admin','HR']);
  if (!candidateId) throw new Error('CandidateID required.');

  var sheet = getSheet_(CONFIG.SHEETS.CANDIDATES);
  var row = findRowByKey_(sheet, 0, candidateId);
  if (row === -1) throw new Error('Candidate not found.');

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameCol = headers.indexOf('Name') + 1;
  var linkCol = headers.indexOf('ResumeLink') + 1;
  var candidateName = sheet.getRange(row, nameCol).getValue();

  var saved = saveResumeToDrive_(candidateId, candidateName, upload);
  sheet.getRange(row, linkCol).setValue(saved.url);

  logAudit_(me.email, 'uploadResume', { id: candidateId, file: saved.name });
  return { ok: true, ResumeLink: saved.url };
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
  try { getCache_().remove(CONFIG.CACHE_KEY_USER_ROLE + String(data.Email).toLowerCase()); } catch (e) {}
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
  try { getCache_().remove(CONFIG.CACHE_KEY_USER_ROLE + String(email).toLowerCase()); } catch (e) {}
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
// Both live in HTML files because Apps Script only allows .gs and .html.
// ============================================================
function getManifestJson() {
  var raw = HtmlService.createHtmlOutputFromFile('Manifest').getContent();
  // Strip the leading <!-- ... --> comment block.
  raw = raw.replace(/<!--[\s\S]*?-->/, '').trim();
  var url = ScriptApp.getService().getUrl();
  return raw.replace(/__START_URL__/g, url);
}

function getServiceWorkerJs() {
  var raw = HtmlService.createHtmlOutputFromFile('ServiceWorker').getContent();
  // Extract the JS between the //<SW> ... //</SW> markers.
  var m = raw.match(/\/\/<SW>([\s\S]*?)\/\/<\/SW>/);
  return m ? m[1].trim() : '';
}
