/**
 * TUSU Budget Pitching — Feedback. Apps Script web app; runs as the account that deployed it (needs edit access to
 * our Sheet and Drive folder, view access to the Submissions Sheet).
 *
 * Reads  the pitching site's Sheet "TUSU Budget Pitching — Submissions" (tabs Organization, Project, Submission_Log),
 *        written by github.com/pppainkiller/budget-pitching-form-submission. Read only: never written here.
 * Writes our own Sheet (FEEDBACK_SHEET_ID):
 *        Users         Email | Role (Admin / Chief / Staff) | Org Refs (persons in charge) | Updated At | Updated By | Name
 * Reads  the pitching scheduling site's Sheet (script property SCHEDULE_SHEET_ID, optional), tab "Schedule":
 *        Project Ref | Date | Start | End | Room. One row per project; shown on the Pitching tab. Read only.
 *        PrePitch      one row per project: approval + feedback (latest)
 *        PrePitch_Log  append-only: one row per changed field on every Pre-pitch save (Admins only)
 *        Pitching      one row per project: summary, final approved budget, Noted documents
 *        Resubmission  one row per project: Approved / Not Approved for the resubmitted project
 *        Settings      Key | Value: site-wide values set in the admin console (Allocated Budget)
 *   and uploaded files into DRIVE_FOLDER_ID / <Org Ref> / "<Project Ref> <purpose> — <file name>".
 *   Files inherit the folder's sharing; they are never deleted (old log rows still point at them).
 *   In the Sheet a file cell shows the file name linked to the file in Drive (one per line for Noted Documents).
 *
 * Every request carries a Firebase ID token (Google sign-in on the website). The email's role comes from the Users tab;
 * the script owner is always Admin, so the first Admin can sign in before the tab has any rows.
 *
 * Speed: each Sheet is opened once per request (memo_), and answers are kept in the script cache (cached_) under a
 * version stamp that every write bumps, so a write is seen at once and outside changes within CACHE_SECONDS.
 *
 * Setup: README.md. After changing this file: clasp push, then update the existing deployment (same /exec URL).
 */

// Script properties with the same names override these.
const DEFAULTS = {
  FIREBASE_API_KEY: 'AIzaSyCSpl4GTgkYvlSWFIqnkUABioPS-GZZOHE', // public web key (same as firebaseConfig.apiKey)
  SUBMISSIONS_SHEET_ID: '18R4b2MfUqUNe8cl9TtxknH1SJlTY6Ut4s_b6m2jo6XE', // the pitching site's "TUSU Budget Pitching — Submissions" Sheet
};
const FEEDBACK_SHEET_ID = '1tDak_LDOEvLEuMMa3Wri9Ej3tWjwLcEDS4tvFb-0N0k';
const DRIVE_FOLDER_ID = '191_2g6tBQwxIJrmUmWMuwECPiRd1Oy2o';
// Apps Script request bodies top out around 50 MB and base64 adds a third.
const MAX_BYTES = 25 * 1024 * 1024;
const CACHE_SECONDS = 60;
const KINDS = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pdf: 'application/pdf',
};
// Staff: edit organization data. Chief: + admin console (users, persons in charge, removed organizations), no edit log.
// Admin: everything. A role not listed here (e.g. the retired Viewer) gives no access.
const RANK = { Staff: 1, Chief: 2, Admin: 3 };
const RENAMED = { Reviewer: 'Staff' }; // roles from earlier versions
const ORG_REF = /^[A-Z]{2}-[A-Z]{3}-\d{2,}$/;
const PROJECT_REF = /^[A-Z]{2}-[A-Z]{3}-\d{2,}-P\d{2}$/;

const TABS = {
  // Name was added later, so it sits last: older Users tabs get the header added by tab_().
  Users: ['Email', 'Role', 'Org Refs', 'Updated At', 'Updated By', 'Name'],
  PrePitch: ['Project Ref', 'Org Ref', 'Project Name', 'Approved Budget', 'Docx Approved', 'Docx Feedback File',
    'PDF Approved', 'PDF Feedback File', 'Feedback Note', 'Updated At', 'Updated By'],
  PrePitch_Log: ['Logged At', 'By', 'Project Ref', 'Org Ref', 'Field', 'Old', 'New'],
  Pitching: ['Project Ref', 'Org Ref', 'Project Name', 'Summary', 'Final Approved Budget', 'Noted Documents',
    'Updated At', 'Updated By'],
  Resubmission: ['Project Ref', 'Org Ref', 'Project Name', 'Decision', 'Updated At', 'Updated By'],
  Settings: ['Key', 'Value', 'Updated At', 'Updated By'],
};
// Cells that hold files: written as file names linked to Drive, read back from those links (filesIn_).
const FILE_COLS = ['Docx Feedback File', 'PDF Feedback File', 'Noted Documents'];
// Pre-pitch fields that are logged.
const PRE_FIELDS = ['Approved Budget', 'Docx Approved', 'Docx Feedback File', 'PDF Approved', 'PDF Feedback File', 'Feedback Note'];

/* ---------- PILOT DEMO ----------
 * Sample submission for RS-ACA-01 (ชุมนุม TEDxThammasatU) so the pilot can be tested before real submissions exist.
 * Used only while the Submissions Sheet has no row for RS-ACA-01; saved Pre-pitch / Pitching rows replace the samples.
 * To remove: set DEMO = null.
 */
const DEMO = {
  org: {
    'Org Ref': 'RS-ACA-01', 'Version': 1, 'Submitted At': '2026-10-01T03:00:00.000Z', 'Submitted By': 'ตัวอย่าง ผู้ส่ง (demo)',
    'Submitter Position': 'ประธานชุมนุม', 'Telegram Phone': '0812345678', 'Emergency Contact': 'ตัวอย่าง ผู้ติดต่อ (demo)',
    'Emergency Position': 'รองประธานชุมนุม', 'Emergency Phone': '0898765432', 'Total Proposed Budget': 185000, 'Project Count': 3,
  },
  projects: [
    { 'Project Ref': 'RS-ACA-01-P01', 'Project Name': 'TEDxThammasatU 2027: Main Event', 'Org Ref': 'RS-ACA-01', 'Project No': 1,
      'Budget': 120000, 'Project Dates': '2027-07-17, 2027-07-18', 'Proposal .docx': '', 'Proposal .pdf': '' },
    { 'Project Ref': 'RS-ACA-01-P02', 'Project Name': 'Speaker Coaching Workshop', 'Org Ref': 'RS-ACA-01', 'Project No': 2,
      'Budget': 45000, 'Project Dates': '2027-06-14, 2027-06-15, 2027-06-16, 2027-06-21', 'Proposal .docx': '', 'Proposal .pdf': '' },
    { 'Project Ref': 'RS-ACA-01-P03', 'Project Name': 'TEDx Salon Night Series', 'Org Ref': 'RS-ACA-01', 'Project No': 3,
      'Budget': 20000, 'Project Dates': '2027-08-06, 2027-08-20, 2027-09-03', 'Proposal .docx': '', 'Proposal .pdf': '' },
  ],
  PrePitch: {
    'RS-ACA-01-P01': { 'Approved Budget': 110000, 'Docx Approved': 'Yes', 'PDF Approved': 'Yes',
      'Feedback Note': 'Clear plan. Venue quotation attached in the budget sheet.', 'Updated By': 'demo' },
    'RS-ACA-01-P02': { 'Approved Budget': 40000, 'Docx Approved': 'No', 'PDF Approved': 'Yes',
      'Feedback Note': 'Please add the trainer fee breakdown to the .docx.', 'Updated By': 'demo' },
  },
  Pitching: {
    'RS-ACA-01-P01': { 'Summary': 'Strong pitch. Committee asked to cut printing costs and confirm the venue date.',
      'Final Approved Budget': 105000, 'Updated By': 'demo' },
    'RS-ACA-01-P02': { 'Summary': 'Approved with a smaller trainer budget.', 'Final Approved Budget': 38000, 'Updated By': 'demo' },
  },
  Schedule: {
    'RS-ACA-01-P01': { date: '2026-10-20', start: '13:00', end: '13:20', room: 'Demo room SC 1001' },
    'RS-ACA-01-P02': { date: '2026-10-20', start: '13:20', end: '13:40', room: 'Demo room SC 1001' },
    'RS-ACA-01-P03': { date: '2026-10-21', start: '09:00', end: '09:20', room: 'Microsoft Teams' },
  },
};

const ACTIONS = {
  boot: boot_, org: org_, upload: upload_, savePrePitch: savePrePitch_, savePitching: savePitching_,
  admin: admin_, saveUsers: saveUsers_, setInCharge: setInCharge_, saveResubmit: saveResubmit_,
  dashboard: dashboard_, saveAllocated: saveAllocated_,
};

function doPost(e) {
  MEMO_ = {};
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('Unknown action');
    const user = verifyUser_(req.idToken);
    user.role = roleOf_(user.email);
    if (!user.role) throw new Error('NO_ACCESS');
    return json_(Object.assign({ ok: true }, fn(req, user)));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() { return json_({ ok: true, service: 'TUSU Budget Pitching — Feedback' }); }

/* ---------- reads ---------- */

// List page: per-organization summary, unfinished approvals, persons in charge, removed organizations.
// Admins also get the user list (for the person-in-charge dropdowns).
function boot_(req, user) {
  const core = cached_('boot', function () {
    const s = subs_();
    const pre = byKey_(feedbackRows_('PrePitch'), 'Project Ref');
    const orgs = {};
    const org = function (ref) { return orgs[ref] || (orgs[ref] = { count: 0, pending: 0, at: '', people: [] }); };
    s.orgs.forEach(function (o) {
      const x = org(o['Org Ref']);
      x.count = Number(o['Project Count']) || 0;
      x.at = iso_(o['Submitted At']);
    });
    s.projects.forEach(function (p) { if (!finished_(pre[p['Project Ref']])) org(p['Org Ref']).pending++; });
    users_().forEach(function (u) { u.orgs.forEach(function (ref) { org(ref).people.push(u.email); }); });
    return { orgs: orgs, removed: removed_(s.log) };
  });
  // names: display name per email (set in the admin console); the site shows the email where none is set.
  const names = {};
  users_().forEach(function (u) { if (u.name) names[u.email] = u.name; });
  const out = { me: { email: user.email, role: user.role }, orgs: core.orgs, removed: core.removed, names: names };
  if (RANK[user.role] >= RANK.Chief) out.users = users_().filter(function (u) { return u.role; }).map(function (u) { return u.email; });
  return out;
}

// Detail page: everything the three tabs need, in one call.
function org_(req, user) {
  const ref = String(req.ref || '');
  if (!ORG_REF.test(ref)) throw new Error('Bad organization code');
  return cached_('org:' + ref, function () {
    const s = subs_();
    const o = s.orgs.filter(function (x) { return x['Org Ref'] === ref; })[0];
    const projects = s.projects.filter(function (p) { return p['Org Ref'] === ref; })
      .sort(function (a, b) { return Number(a['Project No']) - Number(b['Project No']); })
      .map(function (p) {
        return {
          ref: p['Project Ref'], name: p['Project Name'], budget: num_(p['Budget']),
          dates: String(p['Project Dates'] || '').split(/[,\s]+/).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); }),
          docx: p['Proposal .docx'] || '', pdf: p['Proposal .pdf'] || '',
        };
      });
    const refs = projects.map(function (p) { return p.ref; });
    const pre = {};
    feedbackRows_('PrePitch').forEach(function (r) { if (refs.indexOf(r['Project Ref']) >= 0) pre[r['Project Ref']] = preOut_(r); });
    const pitch = {};
    feedbackRows_('Pitching').forEach(function (r) { if (refs.indexOf(r['Project Ref']) >= 0) pitch[r['Project Ref']] = pitchOut_(r); });
    const resub = {};
    rows_(tab_('Resubmission')).forEach(function (r) { if (refs.indexOf(r['Project Ref']) >= 0) resub[r['Project Ref']] = resubOut_(r); });
    const slots = {};
    schedule_().forEach(function (r) { if (refs.indexOf(r.ref) >= 0) slots[r.ref] = r; });
    return {
      org: o ? {
        at: iso_(o['Submitted At']), version: o['Version'], total: num_(o['Total Proposed Budget']),
        name: o['Submitted By'], position: o['Submitter Position'], phone: String(o['Telegram Phone'] || ''),
        emName: o['Emergency Contact'], emPosition: o['Emergency Position'], emPhone: String(o['Emergency Phone'] || ''),
      } : null,
      projects: projects, pre: pre, pitch: pitch, resub: resub, slots: slots,
      people: users_().filter(function (u) { return u.orgs.indexOf(ref) >= 0; }).map(function (u) { return u.email; }),
    };
  });
}

/* ---------- writes ---------- */

// One file into DRIVE_FOLDER_ID/<Org Ref>/. Returns { file: { id, name } } for a following save.
function upload_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  const name = String(req.name || '').replace(/[\\/:*?"<>|]/g, '_').slice(0, 180);
  const kind = ((name.match(/\.([a-z0-9]+)$/i) || [])[1] || '').toLowerCase();
  if (!KINDS[kind] || req.mime !== KINDS[kind]) throw new Error('Only .docx or .pdf files');
  if (req.only && req.only !== kind) throw new Error('Only a .' + req.only + ' file can go here');
  const purpose = { feedback: 'feedback', noted: 'noted' }[req.purpose];
  if (!purpose) throw new Error('Bad purpose');
  const bytes = Utilities.base64Decode(req.data || '');
  if (!bytes.length) throw new Error('Empty file');
  if (bytes.length > MAX_BYTES) throw new Error('File larger than 25 MB');
  const folder = child_(DriveApp.getFolderById(DRIVE_FOLDER_ID), p.org);
  const f = folder.createFile(Utilities.newBlob(bytes, KINDS[kind], p.ref + ' ' + purpose + ' — ' + name));
  return { file: { id: f.getId(), name: name } };
}

function savePrePitch_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  const d = req.data || {};
  const docxOk = yesNo_(d.docxOk);
  const pdfOk = yesNo_(d.pdfOk);
  return locked_(function () {
    const sh = tab_('PrePitch');
    const all = rows_(sh);
    const old = all.filter(function (r) { return r['Project Ref'] === p.ref; })[0] || demoRow_('PrePitch', p.ref) || {};
    const now = new Date();
    const row = {
      'Project Ref': p.ref, 'Org Ref': p.org, 'Project Name': p.name,
      'Approved Budget': money_(d.approvedBudget),
      'Docx Approved': docxOk, 'Docx Feedback File': docxOk === 'No' ? one_(file_(d.docxFile, p.org, 'docx')) : [],
      'PDF Approved': pdfOk, 'PDF Feedback File': pdfOk === 'No' ? one_(file_(d.pdfFile, p.org, 'pdf')) : [],
      'Feedback Note': text_(d.note, 5000), 'Updated At': now, 'Updated By': user.email,
    };
    const log = PRE_FIELDS.filter(function (k) { return logText_(old[k]) !== logText_(row[k]); })
      .map(function (k) { return [now, user.email, p.ref, p.org, k, cell_(logText_(old[k])), cell_(logText_(row[k]))]; });
    if (log.length) {
      upsert_(sh, all, TABS.PrePitch, row);
      const lg = tab_('PrePitch_Log');
      lg.getRange(lg.getLastRow() + 1, 1, log.length, log[0].length).setValues(log);
      bump_();
    }
    return { pre: preOut_(log.length ? row : old), changed: log.length };
  });
}

function savePitching_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  const d = req.data || {};
  const noted = (Array.isArray(d.noted) ? d.noted : []).slice(0, 30).map(function (f) { return file_(f, p.org); }).filter(Boolean);
  const row = {
    'Project Ref': p.ref, 'Org Ref': p.org, 'Project Name': p.name, 'Summary': text_(d.summary, 20000),
    'Final Approved Budget': money_(d.finalBudget), 'Noted Documents': noted,
    'Updated At': new Date(), 'Updated By': user.email,
  };
  locked_(function () {
    const sh = tab_('Pitching');
    upsert_(sh, rows_(sh), TABS.Pitching, row);
    bump_();
  });
  return { pitch: pitchOut_(row) };
}

// Re-submission decision for one project: 'Approved', 'Not Approved' or '' (cleared).
function saveResubmit_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  const decision = String(req.decision || '');
  if (decision && decision !== 'Approved' && decision !== 'Not Approved') throw new Error('Decision must be Approved or Not Approved');
  const row = { 'Project Ref': p.ref, 'Org Ref': p.org, 'Project Name': p.name, 'Decision': decision, 'Updated At': new Date(), 'Updated By': user.email };
  locked_(function () {
    const sh = tab_('Resubmission');
    upsert_(sh, rows_(sh), TABS.Resubmission, row);
    bump_();
  });
  return { resub: resubOut_(row) };
}

function resubOut_(r) { return { decision: r['Decision'] || '', at: iso_(r['Updated At']), by: r['Updated By'] || '' }; }

/* ---------- dashboard ---------- */

// Every project of the organizations on the list (removed ones left out) with its budget at each stage:
// budget (proposed), approved (Pre-pitch approved budget), pitched (Pitching final approved budget), decision
// (Re-submission). null = not set yet. The page adds them up and lists them. Plus the allocated budget.
function dashboard_(req, user) {
  return cached_('dash', function () {
    const s = subs_();
    const removed = removed_(s.log);
    const money = function (name, col) {
      const out = {};
      feedbackRows_(name).forEach(function (r) { const v = num_(r[col]); if (v != null) out[r['Project Ref']] = v; });
      return out;
    };
    const approved = money('PrePitch', 'Approved Budget');
    const pitched = money('Pitching', 'Final Approved Budget');
    const decision = {};
    rows_(tab_('Resubmission')).forEach(function (r) { if (r['Decision']) decision[r['Project Ref']] = r['Decision']; });
    const or = function (o, k) { return k in o ? o[k] : null; };
    return {
      allocated: allocated_(),
      projects: s.projects.filter(function (p) { return removed.indexOf(p['Org Ref']) < 0; }).map(function (p) {
        const ref = p['Project Ref'];
        return { ref: ref, org: p['Org Ref'], name: p['Project Name'], budget: num_(p['Budget']) || 0,
          approved: or(approved, ref), pitched: or(pitched, ref), decision: decision[ref] || '' };
      }),
    };
  });
}

function allocated_() {
  const r = rows_(tab_('Settings')).filter(function (x) { return x['Key'] === 'Allocated Budget'; })[0];
  return r ? num_(r['Value']) : null;
}

// Allocated budget (กรอบงบประมาณ): Admins and Chiefs. Empty = not set.
function saveAllocated_(req, user) {
  need_(user, 'Chief');
  const v = money_(req.allocated);
  locked_(function () {
    const sh = tab_('Settings');
    const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (x) { return x[0]; }) : [];
    const i = keys.indexOf('Allocated Budget');
    sh.getRange(i >= 0 ? i + 2 : sh.getLastRow() + 1, 1, 1, 4).setValues([['Allocated Budget', v, new Date(), user.email]]);
    bump_();
  });
  return { allocated: v === '' ? null : v };
}

/* ---------- admin ---------- */

// Admin console. The edit log goes to Admins only. Each part is cached until the next write.
function admin_(req, user) {
  need_(user, 'Chief');
  const out = cached_('admin', function () { return { removed: removed_(submissionLog_()), allocated: allocated_() }; });
  out.users = users_();
  out.owner = owner_();
  if (user.role === 'Admin') out.log = cached_('log', logTail_);
  return out;
}

// Removed organizations need only Submission_Log: skip the Organization and Project tabs unless already read.
function submissionLog_() {
  if (MEMO_.subs) return MEMO_.subs.log;
  return rows_(SpreadsheetApp.openById(prop_('SUBMISSIONS_SHEET_ID')).getSheetByName('Submission_Log'));
}

// The latest 500 PrePitch_Log rows, newest first. Reads only those rows, however long the log grows.
function logTail_() {
  const sh = tab_('PrePitch_Log');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(500, last - 1);
  return sh.getRange(last - n + 1, 1, n, TABS.PrePitch_Log.length).getValues().reverse()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { at: iso_(r[0]), by: r[1], ref: r[2], org: r[3], field: r[4], old: String(r[5]), now: String(r[6]) }; });
}

// Replaces the whole Users tab. A Chief can't add, change or remove Admins (or they could make themselves one).
function saveUsers_(req, user) {
  need_(user, 'Chief');
  const before = {};
  users_().forEach(function (u) { before[u.email] = u.role; });
  const list = Array.isArray(req.users) ? req.users : [];
  if (list.length > 500) throw new Error('Too many users');
  const seen = {};
  const now = new Date();
  const rows = list.map(function (u) {
    const email = String(u.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Bad email: ' + email);
    if (seen[email]) throw new Error('Listed twice: ' + email);
    seen[email] = true;
    const role = email === owner_() ? 'Admin' : u.role;
    if (!RANK[role]) throw new Error('Bad role for ' + email);
    if (user.role !== 'Admin' && (role === 'Admin') !== (before[email] === 'Admin')) throw new Error('Only Admins can add or change Admins');
    const orgs = (Array.isArray(u.orgs) ? u.orgs : []).map(function (r) { return String(r).trim().toUpperCase(); }).filter(String);
    orgs.forEach(function (r) { if (!ORG_REF.test(r)) throw new Error('Bad organization code ' + r + ' for ' + email); });
    return [email, role, orgs.join(', '), now, user.email, cell_(text_(u.name, 80).trim())];
  });
  if (user.role !== 'Admin') {
    Object.keys(before).forEach(function (e) { if (before[e] === 'Admin' && e !== owner_() && !seen[e]) throw new Error('Only Admins can remove Admins'); });
  }
  locked_(function () {
    const sh = tab_('Users');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TABS.Users.length).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, TABS.Users.length).setValues(rows);
    bump_();
  });
  return { users: users_() };
}

// Person in charge of one organization (list-page dropdown): the chosen user gets the RefCode, everyone else loses it.
// email '' = nobody.
function setInCharge_(req, user) {
  need_(user, 'Chief');
  const ref = String(req.ref || '');
  if (!ORG_REF.test(ref)) throw new Error('Bad organization code');
  const email = String(req.email || '').trim().toLowerCase();
  return locked_(function () {
    const sh = tab_('Users');
    const n = sh.getLastRow() - 1;
    if (email && n < 1) throw new Error(email + ' is not on the user list');
    const vals = n > 0 ? sh.getRange(2, 1, n, 3).getValues() : [];
    if (email && !vals.some(function (r) { return String(r[0]).trim().toLowerCase() === email; })) throw new Error(email + ' is not on the user list');
    const now = new Date();
    vals.forEach(function (r, i) {
      const has = String(r[2] || '').split(/[,\s]+/).filter(function (x) { return x; });
      const keep = has.filter(function (x) { return x !== ref; });
      if (String(r[0]).trim().toLowerCase() === email) keep.push(ref);
      if (keep.join(', ') !== has.join(', ')) sh.getRange(i + 2, 3, 1, 3).setValues([[keep.join(', '), now, user.email]]);
    });
    bump_();
    return { people: email ? [email] : [] };
  });
}

// Pitching slots from the scheduling site: [{ ref, date: 'yyyy-mm-dd', start: 'HH:mm', end: 'HH:mm', room }].
// Empty until SCHEDULE_SHEET_ID is set (plus the demo slots while the pilot runs).
function schedule_() {
  return memo_('schedule', function () {
    const id = PropertiesService.getScriptProperties().getProperty('SCHEDULE_SHEET_ID');
    const fmt = function (v, f) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), f) : String(v || '').trim(); };
    const day = function (v) { return fmt(v, 'yyyy-MM-dd'); };
    const time = function (v) { return fmt(v, 'HH:mm'); };
    const rows = id ? rows_(SpreadsheetApp.openById(id).getSheetByName('Schedule')) : [];
    const out = rows.filter(function (r) { return PROJECT_REF.test(String(r['Project Ref'] || '')); }).map(function (r) {
      return { ref: r['Project Ref'], date: day(r['Date']), start: time(r['Start']), end: time(r['End']), room: String(r['Room'] || '') };
    });
    if (DEMO) Object.keys(DEMO.Schedule).forEach(function (ref) {
      if (!out.some(function (r) { return r.ref === ref; })) out.push(Object.assign({ ref: ref }, DEMO.Schedule[ref]));
    });
    return out;
  });
}

/* ---------- rules ---------- */

// Finished = Approved budget filled, and .docx and PDF each Yes, or No with a feedback file in that field.
function finished_(r) {
  if (!r || r['Approved Budget'] === '' || r['Approved Budget'] == null) return false;
  return ['Docx', 'PDF'].every(function (k) {
    const a = r[k + ' Approved'];
    return a === 'Yes' || (a === 'No' && files_(r[k + ' Feedback File']).length > 0);
  });
}

// Organizations whose latest remove / restore in Submission_Log is a remove (the pitching site hides them).
function removed_(log) {
  const last = {};
  log.forEach(function (r) { if (r['Action'] === 'remove' || r['Action'] === 'restore') last[r['Org Ref']] = r['Action']; });
  return Object.keys(last).filter(function (k) { return last[k] === 'remove'; }).sort();
}

function need_(user, role) { if (RANK[user.role] < RANK[role]) throw new Error('Only ' + role + 's can do this'); }

/* ---------- data helpers ---------- */

// The pitching site's Sheet, read once per request (+ the pilot demo while it has no RS-ACA-01 row).
function subs_() {
  return memo_('subs', function () {
    const ss = SpreadsheetApp.openById(prop_('SUBMISSIONS_SHEET_ID'));
    const s = {
      orgs: rows_(ss.getSheetByName('Organization')),
      projects: rows_(ss.getSheetByName('Project')),
      log: rows_(ss.getSheetByName('Submission_Log')),
    };
    if (DEMO && !s.orgs.some(function (o) { return o['Org Ref'] === DEMO.org['Org Ref']; })) {
      s.orgs.push(DEMO.org);
      s.projects = s.projects.concat(DEMO.projects);
    }
    return s;
  });
}

// Our PrePitch / Pitching rows, plus pilot samples for demo projects that have no saved row yet.
function feedbackRows_(name) {
  const rows = rows_(tab_(name));
  if (!DEMO || !subs_().projects.some(function (p) { return p['Org Ref'] === DEMO.org['Org Ref'] && DEMO.projects.indexOf(p) >= 0; })) return rows;
  const have = rows.map(function (r) { return r['Project Ref']; });
  return rows.concat(Object.keys(DEMO[name]).filter(function (ref) { return have.indexOf(ref) < 0; }).map(function (ref) { return demoRow_(name, ref); }));
}
function demoRow_(name, ref) {
  return DEMO && DEMO[name][ref] ? Object.assign({ 'Project Ref': ref, 'Updated At': DEMO.org['Submitted At'] }, DEMO[name][ref]) : null;
}

// The project as it stands in the latest submission; the organization always comes from here, never from the client.
function project_(ref) {
  ref = String(ref || '');
  if (!PROJECT_REF.test(ref)) throw new Error('Bad project code');
  const p = subs_().projects.filter(function (x) { return x['Project Ref'] === ref; })[0];
  if (!p) throw new Error('Project ' + ref + ' is not in the latest submissions');
  return { ref: ref, org: p['Org Ref'], name: p['Project Name'] };
}

// A file picked earlier through upload_: must sit in DRIVE_FOLDER_ID/<org>/. Returns { id, name } or null.
function file_(f, org, kind) {
  if (!f || !f.id) return null;
  const id = String(f.id);
  if (!/^[\w-]{10,}$/.test(id)) throw new Error('Bad file id');
  const it = DriveApp.getFileById(id);
  const parent = it.getParents().hasNext() ? it.getParents().next() : null;
  const grand = parent && parent.getParents().hasNext() ? parent.getParents().next() : null;
  if (!parent || parent.getName() !== org || !grand || grand.getId() !== DRIVE_FOLDER_ID) throw new Error('File is not in this organization\'s folder');
  if (kind && it.getMimeType() !== KINDS[kind]) throw new Error('Feedback file must be .' + kind);
  return { id: id, name: String(f.name || it.getName()).slice(0, 180) };
}
function one_(f) { return f ? [f] : []; }
function files_(v) { return Array.isArray(v) ? v : []; }
function driveUrl_(id) { return 'https://drive.google.com/file/d/' + id + '/view'; }
// For the log and change checks: files as "name (link)", anything else as text.
function logText_(v) {
  return Array.isArray(v) ? v.map(function (f) { return f.name + ' (' + driveUrl_(f.id) + ')'; }).join(', ') : String(v == null ? '' : v);
}

// Users tab, cached: every request needs it for the role check. The owner may be listed (for a name or
// organizations in charge); the owner's role is always Admin whatever the tab says.
function users_() {
  return memo_('users', function () {
    return cached_('users', function () {
      const owner = owner_();
      return rows_(tab_('Users')).map(function (r) {
        const email = String(r['Email']).trim().toLowerCase();
        return {
          email: email, role: email === owner ? 'Admin' : RANK[RENAMED[r['Role']] || r['Role']] ? RENAMED[r['Role']] || r['Role'] : '', name: String(r['Name'] || '').trim(),
          orgs: String(r['Org Refs'] || '').split(/[,\s]+/).filter(function (x) { return ORG_REF.test(x); }),
        };
      }).filter(function (u) { return u.email; });
    });
  });
}

function roleOf_(email) {
  if (email === owner_()) return 'Admin';
  const u = users_().filter(function (x) { return x.email === email; })[0];
  return u ? u.role : '';
}

function owner_() { return memo_('owner', function () { return String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); }); }

function preOut_(r) {
  return {
    approvedBudget: num_(r['Approved Budget']), docxOk: r['Docx Approved'] || '', docxFile: files_(r['Docx Feedback File'])[0] || null,
    pdfOk: r['PDF Approved'] || '', pdfFile: files_(r['PDF Feedback File'])[0] || null, note: String(r['Feedback Note'] || ''),
    at: iso_(r['Updated At']), by: r['Updated By'] || '', finished: finished_(r),
  };
}

function pitchOut_(r) {
  return {
    summary: String(r['Summary'] || ''), finalBudget: num_(r['Final Approved Budget']),
    noted: files_(r['Noted Documents']), at: iso_(r['Updated At']), by: r['Updated By'] || '',
  };
}

/* ---------- speed: per-request memo + script cache ---------- */

let MEMO_ = {};
function memo_(k, fn) { return Object.prototype.hasOwnProperty.call(MEMO_, k) ? MEMO_[k] : (MEMO_[k] = fn()); }

// Answers kept CACHE_SECONDS under the current version stamp; bump_() after a write makes every older answer unused.
function cached_(key, fn) {
  const c = CacheService.getScriptCache();
  const k = key + '@' + memo_('ver', function () { return c.get('ver') || '0'; });
  const hit = c.get(k);
  if (hit) return JSON.parse(hit);
  const v = fn();
  try { c.put(k, JSON.stringify(v), CACHE_SECONDS); } catch (e) { /* over 100 KB: not cached */ }
  return v;
}
function bump_() {
  const v = Date.now().toString(36) + Math.random().toString(36).slice(2, 8); // unique even for two writes in one ms
  CacheService.getScriptCache().put('ver', v, 21600);
  // Keep the open spreadsheets and the Submissions rows (not changed by our writes); drop everything read from our tabs.
  const keep = { ver: v };
  ['owner', 'ours', 'subs'].forEach(function (k) { if (MEMO_[k] !== undefined) keep[k] = MEMO_[k]; });
  MEMO_ = keep;
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- sheet helpers ---------- */

// Our tab, created with a frozen header row when missing. The spreadsheet is opened once per request.
function tab_(name) {
  const ss = memo_('ours', function () { return SpreadsheetApp.openById(FEEDBACK_SHEET_ID); });
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < TABS[name].length) {
    // A column added in a later version: complete the header row (existing rows keep their cells).
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
  }
  return sh;
}

// Rows as objects keyed by the header row. File columns come back as [{ id, name }] (filesIn_).
function rows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  const fileCols = head.map(function (h, i) { return FILE_COLS.indexOf(h) >= 0 ? i : -1; }).filter(function (i) { return i >= 0; });
  const rich = {};
  fileCols.forEach(function (i) { rich[i] = sh.getRange(2, i + 1, v.length, 1).getRichTextValues(); });
  return v.map(function (r, n) {
    const o = {};
    head.forEach(function (h, i) { o[h] = i in rich ? filesIn_(rich[i][n][0], r[i]) : r[i]; });
    return o;
  }).filter(function (o) { return o[head[0]] !== ''; });
}

// A file cell: every linked run is one file (text = name, link = Drive). Cells written by older versions hold
// JSON ({ id, name } or a list of them) and are still read.
function filesIn_(rt, raw) {
  const out = [];
  (rt ? rt.getRuns() : []).forEach(function (run) {
    const m = String(run.getLinkUrl() || '').match(/\/d\/([\w-]+)/);
    if (m && run.getText().trim()) out.push({ id: m[1], name: run.getText().trim() });
  });
  if (out.length) return out;
  const old = parse_(raw);
  return old ? (Array.isArray(old) ? old : [old]).filter(function (f) { return f && f.id; }) : [];
}

// A file cell: the names, one per line, each linked to its file in Drive.
function filesOut_(list) {
  const text = list.map(function (f) { return f.name; }).join('\n');
  const b = SpreadsheetApp.newRichTextValue().setText(text);
  let at = 0;
  list.forEach(function (f) { if (f.name) b.setLinkUrl(at, at + f.name.length, driveUrl_(f.id)); at += f.name.length + 1; });
  return b.build();
}

function byKey_(rows, key) {
  const o = {};
  rows.forEach(function (r) { o[r[key]] = r; });
  return o;
}

// Write `row` (object) over the row with the same Project Ref, or append it.
function upsert_(sh, all, cols, row) {
  const i = all.map(function (r) { return r['Project Ref']; }).indexOf(row['Project Ref']);
  const at = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, cols.length).setValues([cols.map(function (c) { return Array.isArray(row[c]) ? '' : cell_(row[c]); })]);
  cols.forEach(function (c, i) { if (Array.isArray(row[c]) && row[c].length) sh.getRange(at, i + 1).setRichTextValue(filesOut_(row[c])); });
}

function child_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* ---------- value helpers ---------- */

// Text into a cell: a leading = + - @ would make Sheets read it as a formula.
function cell_(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : (v == null ? '' : v); }
function text_(v, max) { return String(v == null ? '' : v).slice(0, max); }
function yesNo_(v) {
  if (v === '' || v == null) return '';
  if (v !== 'Yes' && v !== 'No') throw new Error('Answer must be Yes or No');
  return v;
}
function money_(v) {
  if (v === '' || v == null) return '';
  const n = Number(v);
  if (!isFinite(n) || n < 0 || n > 1e12) throw new Error('Bad amount');
  return Math.round(n * 100) / 100;
}
function num_(v) { return v === '' || v == null || isNaN(Number(v)) ? null : Math.round(Number(v) * 100) / 100; }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function parse_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function prop_(k) {
  const v = PropertiesService.getScriptProperties().getProperty(k) || DEFAULTS[k];
  if (!v) throw new Error('Script property ' + k + ' not set');
  return v;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- auth ---------- */

// Firebase ID token → verified email (same check as the pitching site's Code.gs). A verified token is remembered
// for 10 minutes (tokens live 1 hour), so most requests skip the round trip to Google.
function verifyUser_(idToken) {
  if (!idToken) throw new Error('Not signed in');
  const c = CacheService.getScriptCache();
  const key = 'tok:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  const hit = c.get(key);
  if (hit) return { email: hit };
  const res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + prop_('FIREBASE_API_KEY'), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ idToken: idToken }),
  });
  if (res.getResponseCode() !== 200) throw new Error('Session expired — sign in again');
  const u = JSON.parse(res.getContentText()).users[0];
  if (!u.email || !u.emailVerified) throw new Error('Email not verified');
  const email = u.email.toLowerCase();
  c.put(key, email, 600);
  return { email: email };
}

/** Run once from the editor: authorizes the scopes and creates our tabs. */
function setup() {
  Object.keys(TABS).forEach(tab_);
  DriveApp.getFolderById(DRIVE_FOLDER_ID).getName();
  SpreadsheetApp.openById(prop_('SUBMISSIONS_SHEET_ID')).getName();
  Logger.log('Ready. Owner (always Admin): ' + owner_());
}
