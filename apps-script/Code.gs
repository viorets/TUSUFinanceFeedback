/**
 * TUSU Budget Pitching — Feedback. Apps Script web app; runs as the account that deployed it (needs edit access to
 * our Sheet and Drive folder, view access to the Submissions Sheet).
 *
 * Reads  the pitching site's Sheet "TUSU Budget Pitching — Submissions" (tabs Organization, Project, Submission_Log),
 *        written by github.com/pppainkiller/budget-pitching-form-submission. Read only: never written here.
 * Writes our own Sheet (FEEDBACK_SHEET_ID):
 *        Users         Email | Role (Admin / Chief / Staff) | Org Refs (persons in charge) | Updated At | Updated By | Name
 * Reads  the pitching scheduling site's Sheet (script property SCHEDULE_SHEET_ID, optional), tab "Schedule":
 *        Project Ref | Date | Start | End | Room | Meeting Link (optional). One row per project; shown on the Pitching
 *        tab and My projects. Read only.
 *        PrePitch      one row per project: approval + feedback (latest)
 *        PrePitch_Log  append-only: one row per changed field on every Pre-pitch save (Admins only)
 *        Pitching      one row per project: summary, final approved budget, Noted documents
 *        Pitching_Log  append-only: one row per changed field on every Pitching save (v2.7); the organization's
 *                      Correct Excel file is logged with an empty Project Ref (v2.9)
 *        OrgPitching   one row per organization: its Correct Excel file (v2.9; one per organization, no longer per
 *                      project: purgeProjectExcel() trashed the per-project files)
 *        OrgPrePitch   one row per organization: note + feedback file on its Budget proposal form (v2.7; logged in
 *                      PrePitch_Log with an empty Project Ref)
 *        Export_Log    append-only: who exported data, when, which parts and format (v2.7; exports are Admin-only)
 *        Resubmission  one row per project: Approved / Not Approved for the resubmitted project
 *        Resubmission_Log  append-only: every decision change and case close / open (Admins and Data team staff)
 *        Settings      Key | Value: site-wide values (Allocated Budget; Required = the required questions, JSON; Phases;
 *                      External Link = { label, url } of the extra top tab; Restored Orgs = organizations removed on
 *                      the pitching site that Admins / Chiefs put back on our list, JSON list)
 *        OrgStatus     Org Ref | Status (Closed / Unclosed) | ...: a closed organization's data cannot be changed
 *        Texts         Key (built-in English line) | EN | TH | ...: the Owner's replacements for site text
 *   and uploaded files into DRIVE_FOLDER_ID / <Org Ref> / "<Project Ref> <purpose> — <file name>".
 *   Files inherit the folder's sharing; they are never deleted (old log rows still point at them).
 *   In the Sheet a file cell shows the file name linked to the file in Drive (one per line for Noted Documents).
 *
 * Every request carries a Firebase ID token (Google sign-in on the website). The email's role comes from the Users tab;
 * the script owner is always Admin, so the first Admin can sign in before the tab has any rows.
 *
 * Speed: each Sheet is opened once per request (memo_), and answers are kept in the script cache (cached_) under a
 * version stamp that every write bumps, so a write is seen at once and outside changes within CACHE_SECONDS.
 * Cached answers are gzipped and split into pieces, so large ones (dashboard, boot) stay cached.
 * File cells have a hidden "<column> JSON" twin holding the same files as text, so reads skip the slow rich-text
 * read. A file cell edited by hand in the Sheet leaves its twin stale: change files through the site.
 * Each Sheet access costs ~0.3-0.5 s whatever its size (load test, v2.6.2), so reads avoid the Sheets entirely when
 * they can: each of our tabs is cached on its own (tabRows_) under a per-tab stamp that only writes to that tab
 * change, and the Submissions rows are cached apart (our writes never change them). A save then rebuilds pages from
 * cache plus one read of the tab it wrote.
 * Saves (v2.6.3): updating a project's existing row is one write and needs no lock; only adding a new row takes the
 * script lock (saveRow_), so saves no longer queue behind each other. Log rows are added with appendRow (atomic).
 * Save conflicts (v2.8): a save carries the Updated At its page loaded (base); when the row changed since, the save is
 * refused (fresh_) instead of silently overwriting. Save users does the same with a Users version (usersVer_). Right
 * before writing, the case is checked again in the Sheet itself (stillOpen_).
 * Uploads (v2.9): 10 MB. The page asks uploadStart for a Drive resumable-upload session in the organization folder
 * and sends the file straight to Google (no base64, no Apps Script in between), then uploadDone checks where it
 * landed. When that fails the page falls back to upload_ (base64 through this script).
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
const MAX_BYTES = 10 * 1024 * 1024; // v2.9: 10 MB (was 25)
const CACHE_SECONDS = 60;
const KINDS = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pdf: 'application/pdf',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls: 'application/vnd.ms-excel',
};
const EXCEL = ['xlsx', 'xls'];
// Staff: edit organization data. Data team staff: Staff + Control (with the Owner) + the edit log.
// Chief: + admin console (users, persons in charge, removed organizations), no edit log.
// Admin: everything. A role not listed here (e.g. the retired Viewer) gives no access.
const DATA_TEAM = 'Data team staff';
const RANK = { Staff: 1, 'Data team staff': 1, Chief: 2, Admin: 3 };
const RENAMED = { Reviewer: 'Staff' }; // roles from earlier versions
const ORG_REF = /^[A-Z]{2}-[A-Z]{3}-\d{2,}$/;
const PROJECT_REF = /^[A-Z]{2}-[A-Z]{3}-\d{2,}-P\d{2}$/;

const TABS = {
  // Name was added later, so it sits last: older Users tabs get the header added by tab_().
  Users: ['Email', 'Role', 'Org Refs', 'Updated At', 'Updated By', 'Name'],
  PrePitch: ['Project Ref', 'Org Ref', 'Project Name', 'Approved Budget', 'Docx Approved', 'Docx Feedback File',
    'PDF Approved', 'PDF Feedback File', 'Feedback Note', 'Updated At', 'Updated By',
    'Docx Feedback File JSON', 'PDF Feedback File JSON'],
  PrePitch_Log: ['Logged At', 'By', 'Project Ref', 'Org Ref', 'Field', 'Old', 'New'],
  Pitching: ['Project Ref', 'Org Ref', 'Project Name', 'Summary', 'Final Approved Budget', 'Noted Documents',
    'Updated At', 'Updated By', 'Correct Excel File', 'Noted Documents JSON', 'Correct Excel File JSON'],
  Pitching_Log: ['Logged At', 'By', 'Project Ref', 'Org Ref', 'Field', 'Old', 'New'],
  OrgPrePitch: ['Org Ref', 'Feedback Note', 'Budget Form Feedback File', 'Updated At', 'Updated By', 'Budget Form Feedback File JSON'],
  OrgPitching: ['Org Ref', 'Correct Excel File', 'Updated At', 'Updated By', 'Correct Excel File JSON'],
  Export_Log: ['Logged At', 'By', 'Parts', 'Format'],
  Resubmission: ['Project Ref', 'Org Ref', 'Project Name', 'Decision', 'Updated At', 'Updated By', 'Feedback'], // Feedback (v2.10): required for Not Approved
  Resubmission_Log: ['Logged At', 'By', 'Project Ref', 'Org Ref', 'Field', 'Old', 'New'],
  Settings: ['Key', 'Value', 'Updated At', 'Updated By'],
  OrgStatus: ['Org Ref', 'Status', 'Updated At', 'Updated By'],
  Texts: ['Key', 'EN', 'TH', 'Updated At', 'Updated By'],
};
// Feedback questions the Owner can mark Required / Not required, with their defaults (= the rules before v2).
// Keep in step with QUESTIONS in public/js/project.js.
// Pre-pitch is for self-noting since v2.2: nothing there is required or tracked.
const QUESTIONS = {
  'pitch.summary': false, 'pitch.finalBudget': true, 'pitch.noted': false, 'resub.decision': true,
};
// Cells that hold files: written as file names linked to Drive, read back from those links (filesIn_).
const FILE_COLS = ['Docx Feedback File', 'PDF Feedback File', 'Noted Documents', 'Correct Excel File', 'Budget Form Feedback File'];
const TWIN = ' JSON'; // hidden text twin of a file column (see Speed above)
// Pre-pitch fields that are logged.
const PRE_FIELDS = ['Approved Budget', 'Docx Approved', 'Docx Feedback File', 'PDF Approved', 'PDF Feedback File', 'Feedback Note'];
// Pitching fields that are logged (v2.7).
const PITCH_FIELDS = ['Summary', 'Final Approved Budget', 'Noted Documents']; // Correct Excel file: per organization (v2.9)
// Budget-form feedback (Pre-pitch, one per organization): sheet column -> log field name.
const FORM_FIELDS = { 'Feedback Note': 'Budget form note', 'Budget Form Feedback File': 'Budget form file' };

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
  setOrgStatus: setOrgStatus_, saveTexts: saveTexts_, saveRequired: saveRequired_, savePhases: savePhases_,
  exportData: exportData_, printData: printData_, saveExternalLink: saveExternalLink_,
  saveOrgPrePitch: saveOrgPrePitch_, myProjects: myProjects_, setRestored: setRestored_,
  saveOrgPitching: saveOrgPitching_, uploadStart: uploadStart_, uploadDone: uploadDone_, mailData: mailData_,
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
    const pitch = byKey_(feedbackRows_('Pitching'), 'Project Ref');
    const resub = byKey_(tabRows_('Resubmission'), 'Project Ref');
    const ph = phases_();
    const orgs = {};
    // pending = projects unfinished on Pitching or Re-submission; proposed / approved = budget totals (approved:
    // the Pitching tab's final approved budgets, null while none is set).
    // v2.10 adds: est = Estimated Approved budget (Pre-pitch approved budget, the proposed budget where empty),
    // pp / rp = projects unfinished on Pitching / Re-submission, ok = projects Approved at Re-submission, fin = their
    // pitched budgets (the final budget), ver / vers = submission versions, rs = re-submissions sent,
    // slots = the pitching slots of its projects.
    const pre = byKey_(feedbackRows_('PrePitch'), 'Project Ref');
    const org = function (ref) {
      return orgs[ref] || (orgs[ref] = { count: 0, pending: 0, at: '', people: [], proposed: 0, approved: null, est: 0, pp: 0, rp: 0, ok: 0, fin: null,
        ver: null, vers: [], rs: 0, slots: [] });
    };
    s.orgs.forEach(function (o) {
      const x = org(o['Org Ref']);
      x.count = Number(o['Project Count']) || 0;
      x.at = iso_(o['Submitted At']);
      x.ver = Number(o['Version']) || null;
    });
    s.log.forEach(function (r) {
      const a = r['Action']; const ref = r['Org Ref'];
      if ((a !== 'submit' && a !== 'resubmit') || !ORG_REF.test(String(ref))) return;
      const x = org(ref);
      if (a === 'resubmit') x.rs++;
      x.vers.push({ v: Number(r['Version']) || null, at: iso_(r['Submitted At'] || r['Logged At'] || r['Timestamp'] || r['At']), kind: a,
        total: num_(r['Total Proposed Budget'] != null ? r['Total Proposed Budget'] : r['Total Budget']), n: Number(r['Project Count']) || null });
    });
    s.projects.forEach(function (p) {
      const x = org(p['Org Ref']); const ref = p['Project Ref'];
      const pendP = ph.pitching && !pitchDone_(pitch[ref]);
      const pendR = ph.resubmit && !resubDone_(resub[ref]);
      if (pendP || pendR) x.pending++;
      if (pendP) x.pp++;
      if (pendR) x.rp++;
      const prop = num_(p['Budget']) || 0;
      x.proposed += prop;
      const est = pre[ref] ? num_(pre[ref]['Approved Budget']) : null;
      x.est += est != null ? est : prop;
      const fin = pitch[ref] ? num_(pitch[ref]['Final Approved Budget']) : null;
      if (fin != null) x.approved = (x.approved || 0) + fin;
      if (resub[ref] && resub[ref]['Decision'] === 'Approved') { x.ok++; if (fin != null) x.fin = (x.fin || 0) + fin; }
    });
    schedule_().forEach(function (r) {
      const ref = String(r.ref).replace(/-P\d+$/, '');
      if (orgs[ref] && r.date) orgs[ref].slots.push([r.date, r.start, r.end]);
    });
    users_().forEach(function (u) { u.orgs.forEach(function (ref) { org(ref).people.push(u.email); }); });
    return { orgs: orgs, removed: hidden_(s.log), closed: closed_(), texts: texts_(), required: required_(), phases: phases_(), link: externalLink_() };
  });
  // names: display name per email (set in the admin console); the site shows the email where none is set.
  const names = {};
  users_().forEach(function (u) { if (u.name) names[u.email] = u.name; });
  const out = { me: { email: user.email, role: user.role, owner: user.email === owner_(), control: canControl_(user), log: canLog_(user) },
    orgs: core.orgs, removed: core.removed, closed: core.closed, texts: core.texts, required: core.required, phases: core.phases,
    link: core.link, names: names };
  // The person-in-charge picker (v2.10: Staff too).
  out.users = users_().filter(function (u) { return u.role; }).map(function (u) { return u.email; });
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
    resubRows_().forEach(function (r) { if (refs.indexOf(r['Project Ref']) >= 0) resub[r['Project Ref']] = resubOut_(r); });
    const slots = {};
    schedule_().forEach(function (r) { if (refs.indexOf(r.ref) >= 0) slots[r.ref] = r; });
    return {
      org: o ? {
        at: iso_(o['Submitted At']), version: o['Version'], total: num_(o['Total Proposed Budget']),
        name: o['Submitted By'], position: o['Submitter Position'], phone: String(o['Telegram Phone'] || ''),
        emName: o['Emergency Contact'], emPosition: o['Emergency Position'], emPhone: String(o['Emergency Phone'] || ''),
        budgetForm: String(o['Budget Proposal Form (.xlsx)'] || ''), // Drive link, one per organization
        emails: emails_(o),
      } : null,
      projects: projects, pre: pre, pitch: pitch, resub: resub, slots: slots, orgPre: orgPreOut_(orgPreRow_(ref)),
      orgPitch: orgPitchOut_(orgPitchRow_(ref)),
      closed: closed_()[ref] || null, required: required_(), phases: phases_(),
      people: users_().filter(function (u) { return u.orgs.indexOf(ref) >= 0; }).map(function (u) { return u.email; }),
    };
  });
}

/* ---------- writes ---------- */

// One file into DRIVE_FOLDER_ID/<Org Ref>/ through this script (base64). Returns { file: { id, name } } for a
// following save. Since v2.9 the page uses it only when the direct upload (uploadStart_) fails.
function upload_(req, user) {
  const u = uploadCheck_(req, user);
  const bytes = Utilities.base64Decode(req.data || '');
  if (!bytes.length) throw new Error('Empty file');
  if (bytes.length > MAX_BYTES) throw new Error('File larger than 10 MB');
  const f = u.folder.createFile(Utilities.newBlob(bytes, KINDS[u.kind], u.driveName));
  return { file: { id: f.getId(), name: u.name } };
}

// Who may upload what, where: role, organization / project, case, phase, type. Returns the target folder, the type
// and the names. purpose 'form' (budget-form feedback) and 'excel' (Correct Excel file, v2.9) belong to an
// organization, 'feedback' and 'noted' to a project.
function uploadCheck_(req, user) {
  need_(user, 'Staff');
  const purpose = { feedback: 'feedback', noted: 'noted', excel: 'excel', form: 'form' }[req.purpose];
  if (!purpose) throw new Error('Bad purpose');
  const p = purpose === 'form' || purpose === 'excel' ? org__(req.org) : project_(req.ref);
  open_(p.org);
  const name = String(req.name || '').replace(/[\\/:*?"<>|]/g, '_').slice(0, 180);
  const kind = ((name.match(/\.([a-z0-9]+)$/i) || [])[1] || '').toLowerCase();
  if (!KINDS[kind] || req.mime !== KINDS[kind]) throw new Error('Only .docx, .pdf, .xlsx or .xls files');
  if (purpose !== 'form' && (purpose === 'excel' ? EXCEL.indexOf(kind) < 0 : EXCEL.indexOf(kind) >= 0)) throw new Error(purpose === 'excel' ? 'Only an Excel file (.xlsx or .xls) can go here' : 'Only .docx or .pdf files');
  if (req.only && req.only !== 'excel' && req.only !== kind) throw new Error('Only a .' + req.only + ' file can go here');
  phaseOn_(purpose === 'feedback' || purpose === 'form' ? 'prepitch' : 'pitching');
  return { org: p.org, kind: kind, name: name, folder: child_(DriveApp.getFolderById(DRIVE_FOLDER_ID), p.org),
    driveName: (p.ref || p.org) + ' ' + purpose + ' — ' + name };
}

// Direct upload, step 1 (v2.9): a Drive resumable-upload session for one file, opened with this script's own token in
// the organization's folder. The page then PUTs the raw file to the returned URL (Drive allows it from the page's
// origin) and calls uploadDone. The session is for this one file and size only.
function uploadStart_(req, user) {
  const u = uploadCheck_(req, user);
  const size = Number(req.size);
  if (!(size > 0)) throw new Error('Empty file');
  if (size > MAX_BYTES) throw new Error('File larger than 10 MB');
  const origin = String(req.origin || '');
  if (!/^https:\/\/[\w.-]+$/.test(origin)) throw new Error('Bad origin');
  // Any failure here reads 'Drive upload could not start': the page then sends the file the old way (upload_).
  let res;
  try {
    res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name', {
      method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': KINDS[u.kind],
        'X-Upload-Content-Length': String(size), Origin: origin },
      payload: JSON.stringify({ name: u.driveName, parents: [u.folder.getId()], mimeType: KINDS[u.kind] }),
    });
  } catch (e) {
    throw new Error('Drive upload could not start (' + (e && e.message || e) + ')');
  }
  const h = res.getAllHeaders();
  const url = h.Location || h.location;
  if (res.getResponseCode() !== 200 || !url) throw new Error('Drive upload could not start (' + res.getResponseCode() + ')');
  return { url: url, name: u.name };
}

// Direct upload, step 2: the file the page sent. It must sit in the organization's folder (as every save checks again).
function uploadDone_(req, user) {
  need_(user, 'Staff');
  const p = req.purpose === 'form' || req.purpose === 'excel' ? org__(req.org) : project_(req.ref);
  const f = file_({ id: req.id, name: req.name }, p.org);
  if (DriveApp.getFileById(f.id).getSize() > MAX_BYTES) { DriveApp.getFileById(f.id).setTrashed(true); throw new Error('File larger than 10 MB'); }
  return { file: f };
}

function savePrePitch_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  open_(p.org);
  phaseOn_('prepitch');
  const d = req.data || {};
  const docxOk = yesNo_(d.docxOk);
  const pdfOk = yesNo_(d.pdfOk);
  const sh = tab_('PrePitch');
  const at = rowOf_(sh, p.ref);
  const old = (at && rowAt_(sh, at)) || demoRow_('PrePitch', p.ref) || {};
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
    fresh_(at && old, req.base);
    stillOpen_(p.org);
    saveRow_(sh, TABS.PrePitch, row, at);
    const lg = tab_('PrePitch_Log');
    log.forEach(function (r) { lg.appendRow(r); });
    bump_('PrePitch', 'PrePitch_Log');
  }
  return { pre: preOut_(log.length ? row : old), changed: log.length };
}

// Budget-form feedback (Pre-pitch, v2.7): one note + one file per organization. Each changed field is logged in
// PrePitch_Log with an empty Project Ref.
function saveOrgPrePitch_(req, user) {
  need_(user, 'Staff');
  const o = org__(req.org);
  open_(o.org);
  phaseOn_('prepitch');
  const d = req.data || {};
  const sh = tab_('OrgPrePitch');
  const at = rowOf_(sh, o.org);
  const old = (at && rowAt_(sh, at)) || {};
  const now = new Date();
  const row = {
    'Org Ref': o.org, 'Feedback Note': text_(d.note, 5000), 'Budget Form Feedback File': one_(file_(d.file, o.org)),
    'Updated At': now, 'Updated By': user.email,
  };
  const log = Object.keys(FORM_FIELDS).filter(function (k) { return logText_(old[k]) !== logText_(row[k]); })
    .map(function (k) { return [now, user.email, '', o.org, FORM_FIELDS[k], cell_(logText_(old[k])), cell_(logText_(row[k]))]; });
  if (log.length) {
    fresh_(at && old, req.base);
    stillOpen_(o.org);
    saveRow_(sh, TABS.OrgPrePitch, row, at);
    const lg = tab_('PrePitch_Log');
    log.forEach(function (r) { lg.appendRow(r); });
    bump_('OrgPrePitch', 'PrePitch_Log');
  }
  return { orgPre: orgPreOut_(log.length ? row : old), changed: log.length };
}

function savePitching_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  open_(p.org);
  phaseOn_('pitching');
  const d = req.data || {};
  const noted = (Array.isArray(d.noted) ? d.noted : []).slice(0, 30).map(function (f) { return file_(f, p.org); }).filter(Boolean);
  // Correct Excel File stays empty on project rows since v2.9 (one per organization: saveOrgPitching_).
  const row = {
    'Project Ref': p.ref, 'Org Ref': p.org, 'Project Name': p.name, 'Summary': text_(d.summary, 20000),
    'Final Approved Budget': money_(d.finalBudget), 'Noted Documents': noted,
    'Updated At': new Date(), 'Updated By': user.email, 'Correct Excel File': [],
  };
  const sh = tab_('Pitching');
  const at = rowOf_(sh, p.ref);
  const old = (at && rowAt_(sh, at)) || demoRow_('Pitching', p.ref) || {};
  const log = PITCH_FIELDS.filter(function (k) { return logText_(old[k]) !== logText_(row[k]); })
    .map(function (k) { return [row['Updated At'], user.email, p.ref, p.org, k, cell_(logText_(old[k])), cell_(logText_(row[k]))]; });
  if (!log.length) return { pitch: pitchOut_(old), changed: 0 };
  fresh_(at && old, req.base);
  stillOpen_(p.org);
  saveRow_(sh, TABS.Pitching, row, at);
  const lg = tab_('Pitching_Log');
  log.forEach(function (r) { lg.appendRow(r); });
  bump_('Pitching', 'Pitching_Log');
  return { pitch: pitchOut_(row), changed: log.length };
}

// Correct Excel file (Pitching, v2.9): one per organization. A change is logged in Pitching_Log with an empty
// Project Ref, like the budget-form feedback in PrePitch_Log.
function saveOrgPitching_(req, user) {
  need_(user, 'Staff');
  const o = org__(req.org);
  open_(o.org);
  phaseOn_('pitching');
  const d = req.data || {};
  const excel = file_(d.excel, o.org);
  if (excel && [KINDS.xlsx, KINDS.xls].indexOf(DriveApp.getFileById(excel.id).getMimeType()) < 0) throw new Error('Correct Excel file must be .xlsx or .xls');
  const sh = tab_('OrgPitching');
  const at = rowOf_(sh, o.org);
  const old = (at && rowAt_(sh, at)) || {};
  const now = new Date();
  const row = { 'Org Ref': o.org, 'Correct Excel File': one_(excel), 'Updated At': now, 'Updated By': user.email };
  if (logText_(old['Correct Excel File']) === logText_(row['Correct Excel File'])) return { orgPitch: orgPitchOut_(at ? old : null), changed: 0 };
  fresh_(at && old, req.base);
  stillOpen_(o.org);
  saveRow_(sh, TABS.OrgPitching, row, at);
  tab_('Pitching_Log').appendRow([now, user.email, '', o.org, 'Correct Excel file', cell_(logText_(old['Correct Excel File'])), cell_(logText_(row['Correct Excel File']))]);
  bump_('OrgPitching', 'Pitching_Log');
  return { orgPitch: orgPitchOut_(row), changed: 1 };
}

// Re-submission decision for one project: 'Approved', 'Not Approved' or '' (cleared).
function saveResubmit_(req, user) {
  need_(user, 'Staff');
  const p = project_(req.ref);
  open_(p.org);
  phaseOn_('resubmit');
  const decision = String(req.decision || '');
  if (decision && decision !== 'Approved' && decision !== 'Not Approved') throw new Error('Decision must be Approved or Not Approved');
  const now = new Date();
  // v2.10: Not Approved needs written feedback; any other decision clears it.
  const feedback = decision === 'Not Approved' ? text_(req.feedback, 2000).trim() : '';
  if (decision === 'Not Approved' && !feedback) throw new Error('Write feedback for a Not Approved project');
  const row = { 'Project Ref': p.ref, 'Org Ref': p.org, 'Project Name': p.name, 'Decision': decision, 'Updated At': now, 'Updated By': user.email, 'Feedback': feedback };
  const sh = tab_('Resubmission');
  const at = rowOf_(sh, p.ref);
  const was = at ? rowAt_(sh, at) : null;
  const old = was ? String(was['Decision'] || '') : '';
  const oldFb = was ? String(was['Feedback'] || '') : '';
  if (old !== decision || oldFb !== feedback) fresh_(was, req.base);
  stillOpen_(p.org);
  saveRow_(sh, TABS.Resubmission, row, at);
  const changes = [];
  if (old !== decision) changes.push([now, user.email, p.ref, p.org, 'Decision', old, decision]);
  if (oldFb !== feedback) changes.push([now, user.email, p.ref, p.org, 'Feedback', cell_(oldFb), cell_(feedback)]);
  if (changes.length) resubLog_(changes);
  bump_('Resubmission', 'Resubmission_Log');
  return { resub: resubOut_(row) };
}

// Resubmission rows, read once per request.
function resubRows_() { return tabRows_('Resubmission'); }
// Settings rows (allocated budget, required questions, phases, extra tab).
function settings_() { return tabRows_('Settings'); }

function resubOut_(r) { return { decision: r['Decision'] || '', feedback: String(r['Feedback'] || ''), at: iso_(r['Updated At']), by: r['Updated By'] || '' }; }

/* ---------- dashboard ---------- */

// Every project of the organizations on the list (removed ones left out) with its budget at each stage:
// budget (proposed), approved (Pre-pitch approved budget), pitched (Pitching final approved budget), decision
// (Re-submission). null = not set yet. The page adds them up and lists them. Plus the allocated budget.
function dashboard_(req, user) {
  return cached_('dash', function () {
    const s = subs_();
    const removed = hidden_(s.log);
    const money = function (name, col) {
      const out = {};
      feedbackRows_(name).forEach(function (r) { const v = num_(r[col]); if (v != null) out[r['Project Ref']] = v; });
      return out;
    };
    const approved = money('PrePitch', 'Approved Budget');
    const pitched = money('Pitching', 'Final Approved Budget');
    const decision = {};
    tabRows_('Resubmission').forEach(function (r) { if (r['Decision']) decision[r['Project Ref']] = r['Decision']; });
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
  const r = settings_().filter(function (x) { return x['Key'] === 'Allocated Budget'; })[0];
  return r ? num_(r['Value']) : null;
}

// Allocated budget (กรอบงบประมาณ): Admins and Chiefs. Empty = not set.
function saveAllocated_(req, user) {
  need_(user, 'Chief');
  const v = money_(req.allocated);
  setSetting_('Allocated Budget', v, user);
  return { allocated: v === '' ? null : v };
}

// One Settings row (Key | Value | Updated At | Updated By), written over the same key or appended.
function setSetting_(key, value, user) {
  locked_(function () {
    const sh = tab_('Settings');
    const keys = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (x) { return x[0]; }) : [];
    const i = keys.indexOf(key);
    sh.getRange(i >= 0 ? i + 2 : sh.getLastRow() + 1, 1, 1, 4).setValues([[key, value, new Date(), user.email]]);
    bump_('Settings');
  });
}

/* ---------- closed organizations ---------- */

// Closed organizations: { 'RS-ART-03': { at, by } }. No row, or Status other than Closed = Unclosed.
function closed_() {
  return memo_('closed', function () {
    const out = {};
    tabRows_('OrgStatus').forEach(function (r) { if (r['Status'] === 'Closed') out[r['Org Ref']] = { at: iso_(r['Updated At']), by: r['Updated By'] || '' }; });
    return out;
  });
}
// Writes to a closed organization's projects are refused, whoever asks.
function open_(orgRef) { if (closed_()[orgRef]) throw new Error('This organization is closed. Its data cannot be changed.'); }
// L8 (v2.8): the same check read from the Sheet itself right before a write (closed_ can be a few seconds old).
function stillOpen_(orgRef) {
  const sh = tab_('OrgStatus');
  const at = rowOf_(sh, orgRef);
  if (at && sh.getRange(at, TABS.OrgStatus.indexOf('Status') + 1).getValue() === 'Closed') throw new Error('This organization is closed. Its data cannot be changed.');
}

// L1 (v2.8): refuses a save when the row (`now`, read from the Sheet; null = none yet) changed after the page loaded
// it. base = the Updated At the page had ('' = no row). Requests without a base (pages from before v2.8) skip it.
// A second of slack: the Sheet may round the stored time.
function fresh_(now, base) {
  if (base === undefined || !now) return;
  const t = Date.parse(iso_(now['Updated At']));
  const b = Date.parse(String(base || ''));
  if (isNaN(t) || (!isNaN(b) && Math.abs(t - b) < 1000)) return;
  throw new Error('CONFLICT:' + JSON.stringify({ by: String(now['Updated By'] || ''), at: iso_(now['Updated At']) }));
}

// Closed / Unclosed switch on the organization list: Admins and Chiefs.
function setOrgStatus_(req, user) {
  need_(user, 'Chief');
  const ref = String(req.ref || '');
  if (!ORG_REF.test(ref)) throw new Error('Bad organization code');
  const closing = req.closed === true;
  // L9 (v2.8): opening a closed case again needs a reason; it goes into Resubmission_Log.
  const reason = text_(req.reason, 500).trim();
  // A case closes only when every project in the latest submission is Approved on Re-submission.
  if (closing) {
    const decided = byKey_(tabRows_('Resubmission'), 'Project Ref');
    const left = subs_().projects.filter(function (p) { return p['Org Ref'] === ref && (decided[p['Project Ref']] || {})['Decision'] !== 'Approved'; });
    if (left.length) throw new Error('Every project must be Approved before the case can be closed (' + left.length + ' not Approved).');
  }
  const now = new Date();
  locked_(function () {
    const sh = tab_('OrgStatus');
    const refs = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (x) { return x[0]; }) : [];
    const i = refs.indexOf(ref);
    const was = i >= 0 && sh.getRange(i + 2, 2).getValue() === 'Closed' ? 'Closed' : 'Open';
    if (was === 'Closed' && !closing && !reason) throw new Error('Give a reason for opening the case again.');
    sh.getRange(i >= 0 ? i + 2 : sh.getLastRow() + 1, 1, 1, 4).setValues([[ref, closing ? 'Closed' : 'Unclosed', now, user.email]]);
    if (was !== (closing ? 'Closed' : 'Open')) resubLog_([[now, user.email, '', ref, 'Case', was, closing ? 'Closed' : 'Open — ' + reason]]);
    bump_('OrgStatus', 'Resubmission_Log');
  });
  return { closed: closing ? { at: now.toISOString(), by: user.email } : null };
}

/* ---------- Owner: site text and required questions ---------- */

// Control tab: the Owner and Data team staff. Edit log: Admins and Data team staff.
function canControl_(user) { return user.email === owner_() || user.role === DATA_TEAM; }
function canLog_(user) { return user.role === 'Admin' || user.role === DATA_TEAM; }
function needControl_(user) { if (!canControl_(user)) throw new Error('Only the Owner and Data team staff can do this'); }

// The Owner's site text: { 'Built-in English line': { en, th } } (empty = keep the built-in line).
function texts_() {
  const out = {};
  tabRows_('Texts').forEach(function (r) {
    const en = String(r['EN'] || ''); const th = String(r['TH'] || '');
    if (en || th) out[String(r['Key'])] = { en: en, th: th };
  });
  return out;
}

// Replaces the whole Texts tab with the lines sent. No < or > (the site puts these lines into its pages).
function saveTexts_(req, user) {
  needControl_(user);
  const list = Array.isArray(req.texts) ? req.texts : [];
  if (list.length > 3000) throw new Error('Too many lines');
  const now = new Date();
  const seen = {};
  const rows = list.map(function (t) {
    const key = text_(t.key, 1000); const en = text_(t.en, 2000).trim(); const th = text_(t.th, 2000).trim();
    if (!key || seen[key]) throw new Error('Bad or repeated line: ' + key);
    seen[key] = true;
    if (/[<>]/.test(en + th)) throw new Error('Text cannot contain < or >: ' + key);
    return [cell_(key), cell_(en), cell_(th), now, user.email];
  }).filter(function (r) { return r[1] !== '' || r[2] !== ''; });
  locked_(function () {
    const sh = tab_('Texts');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TABS.Texts.length).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, TABS.Texts.length).setValues(rows);
    bump_('Texts');
  });
  return { texts: texts_() };
}

// Required / Not required per question: the defaults overlaid with the Owner's choices (Settings, key "Required").
function required_() {
  return memo_('required', function () {
    const r = settings_().filter(function (x) { return x['Key'] === 'Required'; })[0];
    const saved = r ? parse_(String(r['Value'])) || {} : {};
    const out = {};
    Object.keys(QUESTIONS).forEach(function (k) { out[k] = k in saved ? saved[k] === true : QUESTIONS[k]; });
    return out;
  });
}
function isReq_(id) { return required_()[id]; }

function saveRequired_(req, user) {
  needControl_(user);
  const v = req.required || {};
  const out = {};
  Object.keys(QUESTIONS).forEach(function (k) { out[k] = k in v ? v[k] === true : QUESTIONS[k]; });
  setSetting_('Required', JSON.stringify(out), user);
  return { required: out };
}

/* ---------- pitching phases ---------- */

// Pre-pitch, Pitching and Re-submission each On or Off (Settings, key "Phases"). Default: all On.
// Off: the tab shows "Pending Phase" and its writes are refused.
const PHASES = ['prepitch', 'pitching', 'resubmit'];
function phases_() {
  return memo_('phases', function () {
    const r = settings_().filter(function (x) { return x['Key'] === 'Phases'; })[0];
    const saved = r ? parse_(String(r['Value'])) || {} : {};
    const out = {};
    PHASES.forEach(function (k) { out[k] = saved[k] !== false; });
    return out;
  });
}
function phaseOn_(k) { if (!phases_()[k]) throw new Error('This phase is not open yet.'); }

// Phase switches in the admin console: Admins and Chiefs.
function savePhases_(req, user) {
  need_(user, 'Chief');
  const v = req.phases || {};
  const out = {};
  PHASES.forEach(function (k) { out[k] = k in v ? v[k] !== false : phases_()[k]; });
  setSetting_('Phases', JSON.stringify(out), user);
  return { phases: out };
}

/* ---------- admin ---------- */

// Admin console. The edit log goes to Admins only. Each part is cached until the next write.
function admin_(req, user) {
  // Data team staff see only the edit log.
  const logs = function (o) {
    o.log = cached_('log', function () { return logTail_('PrePitch_Log'); });
    o.pitchLog = cached_('pitchLog', function () { return logTail_('Pitching_Log'); });
    o.resubLog = cached_('resubLog', function () { return logTail_('Resubmission_Log'); });
    return o;
  };
  if (user.role === DATA_TEAM) return logs({ logOnly: true });
  need_(user, 'Chief');
  // removed: every organization removed on the pitching site; restored: those put back on our list.
  const out = cached_('admin', function () { return { removed: removed_(submissionLog_()), restored: restored_(), allocated: allocated_(), phases: phases_() }; });
  out.users = users_();
  out.usersVer = usersVer_();
  out.owner = owner_();
  out.sheetUrl = 'https://docs.google.com/spreadsheets/d/' + FEEDBACK_SHEET_ID + '/edit';
  return canLog_(user) ? logs(out) : out;
}

// Print campus (organization list): the org data of every organization of one campus that has a submission
// (removed ones left out), in one call. Reads are shared across organizations (memo_).
function printData_(req, user) {
  need_(user, 'Staff');
  const campus = String(req.campus || '');
  if (!/^[A-Z]{2}$/.test(campus)) throw new Error('Bad campus code');
  const s = subs_();
  const gone = hidden_(s.log);
  const refs = s.orgs.map(function (o) { return o['Org Ref']; })
    .filter(function (r) { return r.split('-')[0] === campus && gone.indexOf(r) < 0; }).sort();
  return { orgs: refs.map(function (ref) { return { ref: ref, data: org_({ ref: ref }, user) }; }) };
}

// Extra top tab (Control): { label, url } linking to another website, or null (tab hidden).
function externalLink_() {
  const r = settings_().filter(function (x) { return x['Key'] === 'External Link'; })[0];
  const v = r ? parse_(String(r['Value'] || '')) : null;
  return v && /^https:\/\/\S+$/.test(v.url || '') ? { label: String(v.label || ''), url: v.url } : null;
}
function saveExternalLink_(req, user) {
  needControl_(user);
  const url = String(req.url || '').trim();
  const label = text_(req.label, 40).trim();
  if (url && !/^https:\/\/\S+$/.test(url)) throw new Error('The link must start with https://');
  if (url.length > 500) throw new Error('The link is too long');
  if (/[<>]/.test(label)) throw new Error('The label cannot contain < or >');
  setSetting_('External Link', url ? JSON.stringify({ label: label, url: url }) : '', user);
  return { link: externalLink_() };
}

// Export (admin console): organizations always, plus the phases asked for. Removed organizations are left out.
// The page turns the rows into an Excel or CSV file. Money as numbers, files as Drive links.
// Admin-only since v2.7 (personal data: phones, emails); every export is logged in Export_Log.
function exportData_(req, user) {
  need_(user, 'Admin');
  const want = req.phases || {};
  const s = subs_();
  const gone = hidden_(s.log);
  const parts = ['organizations', 'projects'].concat(['prepitch', 'pitching', 'resubmit', 'emails'].filter(function (k) { return want[k]; }));
  const fmt = req.format === 'csv' ? 'CSV' : 'Excel';
  tab_('Export_Log').appendRow([new Date(), user.email, parts.join(', '), fmt]);
  const keep = function (ref) { return gone.indexOf(ref) < 0; };
  const shut = closed_();
  const links = function (list) { return files_(list).map(function (f) { return driveUrl_(f.id); }).join('\n'); };
  const orgs = s.orgs.filter(function (o) { return keep(o['Org Ref']); }).map(function (o) {
    const ref = o['Org Ref'];
    return {
      'Org Ref': ref, 'Organization': o['Organization'] || '', 'Org Type': o['Org Type'] || '', 'Campus': o['Campus'] || o['Campus Code'] || '',
      'Category': o['Category'] || o['Category Code'] || '', 'Version': o['Version'], 'Submitted At': iso_(o['Submitted At']),
      'Submitted By': o['Submitted By'] || '', 'Submitter Email': o['Submitter Email'] || '', 'Submitter Position': o['Submitter Position'] || '',
      'Telegram Phone': String(o['Telegram Phone'] || ''), 'Emergency Contact': o['Emergency Contact'] || '',
      'Emergency Position': o['Emergency Position'] || '', 'Emergency Phone': String(o['Emergency Phone'] || ''),
      'Total Proposed Budget': num_(o['Total Proposed Budget']), 'Project Count': Number(o['Project Count']) || 0,
      'Budget Proposal Form (.xlsx)': String(o['Budget Proposal Form (.xlsx)'] || ''),
      'Notification Emails': emails_(o).join(', '),
      'Persons in charge': users_().filter(function (u) { return u.orgs.indexOf(ref) >= 0; }).map(function (u) { return u.email; }).join(', '),
      'Case': shut[ref] ? 'Closed' : 'Open',
    };
  });
  const projects = s.projects.filter(function (p) { return keep(p['Org Ref']); }).map(function (p) {
    return { 'Project Ref': p['Project Ref'], 'Org Ref': p['Org Ref'], 'Project Name': p['Project Name'] || '', 'Project No': Number(p['Project No']) || '',
      'Budget': num_(p['Budget']), 'Project Dates': String(p['Project Dates'] || ''), 'Proposal .docx': p['Proposal .docx'] || '', 'Proposal .pdf': p['Proposal .pdf'] || '' };
  });
  const out = { orgs: orgs, projects: projects };
  // One row per notification email.
  if (want.emails) {
    out.emails = [];
    orgs.forEach(function (o) {
      emails_(s.orgs.filter(function (x) { return x['Org Ref'] === o['Org Ref']; })[0]).forEach(function (e) {
        out.emails.push({ 'Org Ref': o['Org Ref'], 'Organization': o['Organization'], 'Campus': o['Campus'], 'Email': e });
      });
    });
  }
  if (want.prepitch) {
    const pre = byKey_(feedbackRows_('PrePitch'), 'Project Ref');
    out.prepitch = projects.map(function (p) {
      const r = pre[p['Project Ref']] || {};
      return { 'Project Ref': p['Project Ref'], 'Estimated Approved Budget': num_(r['Approved Budget']), 'Docx Approved': r['Docx Approved'] || '',
        'Docx Feedback File': links(r['Docx Feedback File']), 'PDF Approved': r['PDF Approved'] || '', 'PDF Feedback File': links(r['PDF Feedback File']),
        'Note': String(r['Feedback Note'] || ''), 'Updated At': iso_(r['Updated At']), 'Updated By': r['Updated By'] || '' };
    });
  }
  if (want.pitching) {
    const pit = byKey_(feedbackRows_('Pitching'), 'Project Ref');
    // Correct Excel file: one per organization (v2.9), on the Organizations sheet; v2.10 repeats it on each Pitching row.
    const xl = byKey_(tabRows_('OrgPitching'), 'Org Ref');
    orgs.forEach(function (o) { o['Correct Excel File'] = links((xl[o['Org Ref']] || {})['Correct Excel File']); });
    // v2.10: same layout as the Pre-pitch sheet (Docx / PDF / Note become Summary / Noted Documents), so the two compare.
    out.pitching = projects.map(function (p) {
      const r = pit[p['Project Ref']] || {};
      return { 'Project Ref': p['Project Ref'], 'Pitched Final Approved Budget': num_(r['Final Approved Budget']), 'Summary': String(r['Summary'] || ''),
        'Noted Documents': links(r['Noted Documents']), 'Correct Excel File': links((xl[p['Org Ref']] || {})['Correct Excel File']),
        'Updated At': iso_(r['Updated At']), 'Updated By': r['Updated By'] || '' };
    });
  }
  if (want.resubmit) {
    const rs = byKey_(tabRows_('Resubmission'), 'Project Ref');
    out.resubmit = projects.map(function (p) {
      const r = rs[p['Project Ref']] || {};
      return { 'Project Ref': p['Project Ref'], 'Decision': r['Decision'] || '', 'Feedback': String(r['Feedback'] || ''), 'Updated At': iso_(r['Updated At']), 'Updated By': r['Updated By'] || '' };
    });
  }
  return out;
}

// Export data for mailing (v2.9, Admin only, logged like exportData_): the fields of the "Official Budget Allocation
// Update" email. Organizations: emails, budget requested / approved and the change, the category and campus
// benchmarks (active organizations = those with at least one final approved budget; average = total approved ÷
// active), the four shares (%, 2 decimals) and the links. Projects: contacts, dates (dd/mm/yyyy, Buddhist year),
// budgets, Committee feedback (= the Pitching summary) and file links. Approved = the Pitching final approved budget.
function mailData_(req, user) {
  need_(user, 'Admin');
  tab_('Export_Log').appendRow([new Date(), user.email, 'mailing', 'Mailing']);
  const s = subs_();
  const gone = hidden_(s.log);
  const pit = byKey_(feedbackRows_('Pitching'), 'Project Ref');
  const xl = byKey_(tabRows_('OrgPitching'), 'Org Ref');
  const links = function (list) { return files_(list).map(function (f) { return driveUrl_(f.id); }).join('\n'); };
  const pct = function (a, b) { return b ? Math.round(a / b * 10000) / 100 : null; };
  const orgs = s.orgs.filter(function (o) { return gone.indexOf(o['Org Ref']) < 0; });
  const projects = s.projects.filter(function (p) { return gone.indexOf(p['Org Ref']) < 0; });
  const sum = {};
  orgs.forEach(function (o) { sum[o['Org Ref']] = { req: 0, appr: null, n: 0 }; });
  projects.forEach(function (p) {
    const x = sum[p['Org Ref']]; if (!x) return;
    x.n++; x.req += num_(p['Budget']) || 0;
    const fin = pit[p['Project Ref']] ? num_(pit[p['Project Ref']]['Final Approved Budget']) : null;
    if (fin != null) x.appr = (x.appr || 0) + fin;
  });
  const group = function (key) {
    const g = {};
    orgs.forEach(function (o) {
      const k = key(o['Org Ref']); const x = sum[o['Org Ref']];
      g[k] = g[k] || { total: 0, active: 0 };
      if (x.appr != null) { g[k].total += x.appr; g[k].active++; }
    });
    return g;
  };
  const catOf = function (ref) { return ref.split('-')[1]; };
  const campusOf = function (ref) { return ref.split('-')[0]; };
  const cat = group(catOf); const camp = group(campusOf);
  const grand = Object.keys(sum).reduce(function (t, k) { return t + (sum[k].appr || 0); }, 0);
  const round2 = function (v) { return v == null ? null : Math.round(v * 100) / 100; };
  const outOrgs = orgs.map(function (o) {
    const ref = o['Org Ref']; const x = sum[ref]; const c = cat[catOf(ref)]; const m = camp[campusOf(ref)];
    const e = emails_(o); const appr = x.appr || 0;
    return {
      'Email 1': e[0] || '', 'Email 2': e[1] || '', 'Email 3': e[2] || '', 'Org Ref': ref, 'Organization': o['Organization'] || '',
      'Campus': o['Campus'] || o['Campus Code'] || campusOf(ref), 'Category': o['Category'] || o['Category Code'] || catOf(ref),
      'Total projects': x.n, 'Requested budget': round2(x.req), 'Approved budget': x.appr == null ? null : round2(appr),
      'Change (THB)': x.appr == null ? null : round2(appr - x.req), 'Change (%)': x.appr == null ? null : pct(appr - x.req, x.req),
      'Category active organizations': c.active, 'Category average allocation': c.active ? round2(c.total / c.active) : null,
      'Campus active organizations': m.active, 'Campus average allocation': m.active ? round2(m.total / m.active) : null,
      'Share of category (%)': x.appr == null ? null : pct(appr, c.total), 'Share of campus (%)': x.appr == null ? null : pct(appr, m.total),
      'Category share of all (%)': pct(c.total, grand), 'Campus share of all (%)': pct(m.total, grand),
      'Budget form link': String(o['Budget Proposal Form (.xlsx)'] || ''),
      'Correct Excel file': links((xl[ref] || {})['Correct Excel File']),
    };
  });
  const be = function (iso) { const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + '/' + m[2] + '/' + (Number(m[1]) + 543) : ''; };
  const byOrg = byKey_(orgs, 'Org Ref');
  const outProjects = projects.map(function (p) {
    const o = byOrg[p['Org Ref']] || {}; const r = pit[p['Project Ref']] || {};
    const dates = String(p['Project Dates'] || '').split(/[,\s]+/).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); });
    return {
      'Org Ref': p['Org Ref'], 'Project Ref': p['Project Ref'], 'Project name': p['Project Name'] || '', 'Event dates': dates.map(be).join(', '),
      'Primary contact': o['Submitted By'] || '', 'Primary position': o['Submitter Position'] || '', 'Primary phone': String(o['Telegram Phone'] || ''),
      'Emergency contact': o['Emergency Contact'] || '', 'Emergency position': o['Emergency Position'] || '', 'Emergency phone': String(o['Emergency Phone'] || ''),
      'Requested budget': num_(p['Budget']), 'Approved budget': num_(r['Final Approved Budget']),
      'Committee feedback': String(r['Summary'] || ''), 'Proposal PDF': p['Proposal .pdf'] || '', 'Proposal DOCX': p['Proposal .docx'] || '',
      'Feedback files': links(r['Noted Documents']),
    };
  });
  return { orgs: outOrgs, projects: outProjects };
}

// One-time, from the Apps Script editor after deploying v2.9 (Run -> purgeProjectExcel): the Correct Excel file is
// one per organization now, so every per-project Correct Excel file is moved to the Drive trash (Google empties it
// after 30 days) and its cells are cleared. Each one is logged in Pitching_Log. Safe to run twice.
function purgeProjectExcel() {
  MEMO_ = {};
  const sh = tab_('Pitching');
  const rows = rows_(sh);
  const col = TABS.Pitching.indexOf('Correct Excel File') + 1;
  const twin = TABS.Pitching.indexOf('Correct Excel File' + TWIN) + 1;
  const lg = tab_('Pitching_Log');
  let files = 0; let projects = 0;
  rows.forEach(function (r) {
    const list = files_(r['Correct Excel File']);
    if (!list.length) return;
    const i = rowOf_(sh, r['Project Ref']) - 2; // rows_ leaves out blank rows: find the sheet row by key
    list.forEach(function (f) { try { DriveApp.getFileById(f.id).setTrashed(true); files++; } catch (e) { Logger.log('Not found: ' + f.id); } });
    sh.getRange(i + 2, col).clearContent();
    if (twin) sh.getRange(i + 2, twin).setValue('[]');
    lg.appendRow([new Date(), Session.getEffectiveUser().getEmail(), r['Project Ref'], r['Org Ref'], 'Correct Excel File',
      cell_(logText_(list)), '(removed in v2.9: one Correct Excel file per organization)']);
    projects++;
  });
  bump_('Pitching', 'Pitching_Log');
  Logger.log('Trashed ' + files + ' files from ' + projects + ' projects');
}

// Removed organizations need only Submission_Log: skip the Organization and Project tabs unless already read.
function submissionLog_() {
  if (MEMO_.subs) return MEMO_.subs.log;
  return rows_(SpreadsheetApp.openById(prop_('SUBMISSIONS_SHEET_ID')).getSheetByName('Submission_Log'));
}

// Re-submission log rows.
function resubLog_(rows) {
  const lg = tab_('Resubmission_Log');
  rows.forEach(function (r) { lg.appendRow(r); }); // atomic: safe without the lock
}

// The latest 500 rows of a log tab (PrePitch_Log, Pitching_Log or Resubmission_Log), newest first. Reads only those rows.
function logTail_(name) {
  const sh = tab_(name || 'PrePitch_Log');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.min(500, last - 1);
  return sh.getRange(last - n + 1, 1, n, TABS[name || 'PrePitch_Log'].length).getValues().reverse()
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
  const ver = locked_(function () {
    // L2 (v2.8): refused when the tab changed after the console loaded it (another Chief's save, or a person in
    // charge set on the list). Consoles from before v2.8 send no version and skip the check.
    if (req.version !== undefined && req.version !== usersVer_()) throw new Error('CONFLICT_USERS');
    const sh = tab_('Users');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TABS.Users.length).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, TABS.Users.length).setValues(rows);
    bump_('Users');
    return usersVer_();
  });
  return { users: users_(), usersVer: ver };
}

// The Users tab's version: a digest of its cells, read from the Sheet itself (never the cache).
function usersVer_() {
  const sh = tab_('Users');
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, TABS.Users.length).getValues() : [];
  // Times to the second, so the Sheet's rounding can't change the version.
  v.forEach(function (r) { r.forEach(function (c, i) { if (c instanceof Date) r[i] = Math.floor(c.getTime() / 1000); }); });
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(v)));
}

// Person in charge of one organization (list-page dropdown): the chosen user gets the RefCode, everyone else loses it.
// email '' = nobody.
// v2.10: Staff can use it, and several people can be in charge of one organization. With `on` (true / false) the
// person is added to / removed from the organization; without `on` (older pages) the person replaces everyone.
function setInCharge_(req, user) {
  need_(user, 'Staff');
  const ref = String(req.ref || '');
  if (!ORG_REF.test(ref)) throw new Error('Bad organization code');
  const email = String(req.email || '').trim().toLowerCase();
  const toggle = req.on === true || req.on === false;
  if (toggle && !email) throw new Error('Choose a person');
  return locked_(function () {
    const sh = tab_('Users');
    const n = sh.getLastRow() - 1;
    if (email && n < 1) throw new Error(email + ' is not on the user list');
    const vals = n > 0 ? sh.getRange(2, 1, n, 3).getValues() : [];
    if (email && !vals.some(function (r) { return String(r[0]).trim().toLowerCase() === email; })) throw new Error(email + ' is not on the user list');
    const now = new Date();
    vals.forEach(function (r, i) {
      const has = String(r[2] || '').split(/[,\s]+/).filter(function (x) { return x; });
      const mine = String(r[0]).trim().toLowerCase() === email;
      // Toggle: only the named person's row changes. Replace: everyone loses the organization, the named person gains it.
      const keep = has.filter(function (x) { return x !== ref || (toggle && !mine); });
      if (mine && (!toggle || req.on)) keep.push(ref);
      r[2] = keep.join(', ');
      if (keep.join(', ') !== has.join(', ')) sh.getRange(i + 2, 3, 1, 3).setValues([[keep.join(', '), now, user.email]]);
    });
    bump_('Users');
    return { people: vals.filter(function (r) { return String(r[2] || '').split(/[,\s]+/).indexOf(ref) >= 0; }).map(function (r) { return String(r[0]).trim().toLowerCase(); }) };
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
      const link = String(r['Meeting Link'] || '').trim();
      return { ref: r['Project Ref'], date: day(r['Date']), start: time(r['Start']), end: time(r['End']), room: String(r['Room'] || ''),
        link: /^https:\/\/\S+$/.test(link) ? link : '' };
    });
    if (DEMO) Object.keys(DEMO.Schedule).forEach(function (ref) {
      if (!out.some(function (r) { return r.ref === ref; })) out.push(Object.assign({ ref: ref }, DEMO.Schedule[ref]));
    });
    return out;
  });
}

/* ---------- rules ---------- */

// Pre-pitch "finished" (kept for the row data; not tracked since v2.2): Approved budget filled, and .docx and PDF
// each Yes, or No with a feedback file in that field.
function finished_(r) {
  if (!r || r['Approved Budget'] === '' || r['Approved Budget'] == null) return false;
  return ['Docx', 'PDF'].every(function (k) {
    const a = r[k + ' Approved'];
    return a === 'Yes' || (a === 'No' && files_(r[k + ' Feedback File']).length > 0);
  });
}
// Pitching / Re-submission done = every required question answered (Owner's choices; keep in step with
// UNFINISHED in public/js/project.js).
function pitchDone_(r) {
  r = r || {};
  if (isReq_('pitch.summary') && !String(r['Summary'] || '').trim()) return false;
  if (isReq_('pitch.finalBudget') && num_(r['Final Approved Budget']) == null) return false;
  if (isReq_('pitch.noted') && !files_(r['Noted Documents']).length) return false;
  return true;
}
function resubDone_(r) { return !isReq_('resub.decision') || !!(r && r['Decision']); }

// Organizations whose latest remove / restore in Submission_Log is a remove (the pitching site hides them).
function removed_(log) {
  const last = {};
  log.forEach(function (r) { if (r['Action'] === 'remove' || r['Action'] === 'restore') last[r['Org Ref']] = r['Action']; });
  return Object.keys(last).filter(function (k) { return last[k] === 'remove'; }).sort();
}

// Organizations removed on the pitching site that Admins / Chiefs put back on our list (v2.7).
function restored_() {
  const r = settings_().filter(function (x) { return x['Key'] === 'Restored Orgs'; })[0];
  const v = r ? parse_(String(r['Value'] || '')) : null;
  return Array.isArray(v) ? v.filter(function (x) { return ORG_REF.test(String(x)); }) : [];
}
// Organizations hidden from our list, Dashboard, export and print: removed on the pitching site and not put back.
function hidden_(log) {
  const back = restored_();
  return removed_(log).filter(function (r) { return back.indexOf(r) < 0; });
}
// Put a removed organization back on our list (on = true) or take it off again: Admins and Chiefs. Logged in
// Resubmission_Log (Field = List).
function setRestored_(req, user) {
  need_(user, 'Chief');
  const ref = String(req.ref || '');
  if (!ORG_REF.test(ref)) throw new Error('Bad organization code');
  const on = req.on === true;
  return locked_(function () {
    const list = restored_().filter(function (r) { return r !== ref; });
    const was = list.length !== restored_().length;
    if (on) list.push(ref);
    list.sort();
    setSetting_('Restored Orgs', JSON.stringify(list), user);
    if (was !== on) resubLog_([[new Date(), user.email, '', ref, 'List', was ? 'Added back' : 'Removed', on ? 'Added back' : 'Removed']]);
    bump_('Settings', 'Resubmission_Log');
    return { restored: list };
  });
}

function need_(user, role) { if (RANK[user.role] < RANK[role]) throw new Error('Only ' + role + 's can do this'); }

/* ---------- data helpers ---------- */

// The pitching site's Sheet, read once per request (+ the pilot demo while it has no RS-ACA-01 row).
function subs_() {
  return memo_('subs', function () {
    const c = CacheService.getScriptCache();
    let s = cacheGet_(c, 'subs');
    if (s === undefined) {
      const ss = SpreadsheetApp.openById(prop_('SUBMISSIONS_SHEET_ID'));
      s = {
        orgs: rows_(ss.getSheetByName('Organization')),
        projects: rows_(ss.getSheetByName('Project')),
        log: rows_(ss.getSheetByName('Submission_Log')),
      };
      cachePut_(c, 'subs', s);
    }
    if (DEMO && !s.orgs.some(function (o) { return o['Org Ref'] === DEMO.org['Org Ref']; })) {
      s.orgs.push(DEMO.org);
      s.projects = s.projects.concat(DEMO.projects);
    }
    return s;
  });
}

// Our PrePitch / Pitching rows, plus pilot samples for demo projects that have no saved row yet.
function feedbackRows_(name) { return memo_('fb:' + name, function () { return feedbackRows0_(name); }); }
function feedbackRows0_(name) {
  const rows = tabRows_(name);
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

// An organization with a submission; the same shape as project_ (org = its RefCode).
function org__(ref) {
  ref = String(ref || '');
  if (!ORG_REF.test(ref)) throw new Error('Bad organization code');
  if (!subs_().orgs.some(function (o) { return o['Org Ref'] === ref; })) throw new Error('Organization ' + ref + ' is not in the latest submissions');
  return { org: ref };
}

// Notification emails from the submission (Organization tab, "Notification Emails"): up to 3, checked, lower case.
function emails_(o) {
  const seen = [];
  String((o && o['Notification Emails']) || '').split(/[\s,;]+/).forEach(function (e) {
    e = e.trim().toLowerCase();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) && seen.indexOf(e) < 0 && seen.length < 3) seen.push(e);
  });
  return seen;
}

// Correct Excel file row of one organization (or null), and what the page gets: { file, at, by } or null.
function orgPitchRow_(ref) {
  return tabRows_('OrgPitching').filter(function (r) { return r['Org Ref'] === ref; })[0] || null;
}
function orgPitchOut_(r) {
  if (!r) return null;
  return { file: files_(r['Correct Excel File'])[0] || null, at: iso_(r['Updated At']), by: r['Updated By'] || '' };
}

// Budget-form feedback row of one organization (or null).
function orgPreRow_(ref) {
  return tabRows_('OrgPrePitch').filter(function (r) { return r['Org Ref'] === ref; })[0] || null;
}
function orgPreOut_(r) {
  r = r || {};
  return { note: String(r['Feedback Note'] || ''), file: files_(r['Budget Form Feedback File'])[0] || null, at: iso_(r['Updated At']), by: r['Updated By'] || '' };
}

// My projects (v2.7): the org data of every organization I'm in charge of that has a submission.
function myProjects_(req, user) {
  need_(user, 'Staff');
  const me = users_().filter(function (u) { return u.email === user.email; })[0];
  const have = subs_().orgs.map(function (o) { return o['Org Ref']; });
  const refs = (me ? me.orgs : []).filter(function (r) { return have.indexOf(r) >= 0; }).sort();
  return { orgs: refs.map(function (ref) { return { ref: ref, data: org_({ ref: ref }, user) }; }) };
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
    return (function () {
      const owner = owner_();
      return tabRows_('Users').map(function (r) {
        const email = String(r['Email']).trim().toLowerCase();
        return {
          email: email, role: email === owner ? 'Admin' : RANK[RENAMED[r['Role']] || r['Role']] ? RENAMED[r['Role']] || r['Role'] : '', name: String(r['Name'] || '').trim(),
          orgs: String(r['Org Refs'] || '').split(/[,\s]+/).filter(function (x) { return ORG_REF.test(x); }),
        };
      }).filter(function (u) { return u.email; });
    })();
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
    noted: files_(r['Noted Documents']),
    at: iso_(r['Updated At']), by: r['Updated By'] || '',
  };
}

/* ---------- speed: per-request memo + script cache ---------- */

let MEMO_ = {};
function memo_(k, fn) { return Object.prototype.hasOwnProperty.call(MEMO_, k) ? MEMO_[k] : (MEMO_[k] = fn()); }

// Answers kept CACHE_SECONDS under the current version stamp; bump_() after a write makes every older answer unused.
// Each answer is gzipped, base64-encoded and split into pieces under the 100 KB per-value limit: key = piece count,
// key#0.. = the pieces. A missing piece counts as a miss.
const PIECE = 90000;
function cached_(key, fn) {
  const c = CacheService.getScriptCache();
  const k = key + '@' + memo_('ver', function () { return c.get('ver') || '0'; });
  const hit = cacheGet_(c, k);
  if (hit !== undefined) return hit;
  const v = fn();
  cachePut_(c, k, v);
  return v;
}
function cachePut_(c, k, v) {
  try {
    const z = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(v), 'application/json')).getBytes());
    const n = Math.ceil(z.length / PIECE);
    if (n > 50) return; // too big to be worth it
    const o = {};
    o[k] = String(n);
    for (let i = 0; i < n; i++) o[k + '#' + i] = z.slice(i * PIECE, (i + 1) * PIECE);
    c.putAll(o, CACHE_SECONDS);
  } catch (e) { /* not cached */ }
}
function cacheGet_(c, k) {
  const n = Number(c.get(k));
  if (!(n > 0)) return undefined;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(k + '#' + i);
  const got = c.getAll(keys);
  let z = '';
  for (let i = 0; i < n; i++) { if (got[keys[i]] == null) return undefined; z += got[keys[i]]; }
  try { return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(z), 'application/x-gzip')).getDataAsString()); } catch (e) { return undefined; }
}
// After a write: a new version stamp (every composed answer is rebuilt) and a new stamp for each tab written (only
// those tabs are read again). No names = every tab.
function bump_() {
  const v = Date.now().toString(36) + Math.random().toString(36).slice(2, 8); // unique even for two writes in one ms
  const names = arguments.length ? Array.prototype.slice.call(arguments) : Object.keys(TABS);
  const o = { ver: v };
  names.forEach(function (n) { o['tv:' + n] = v; });
  CacheService.getScriptCache().putAll(o, 21600);
  // Keep the open spreadsheets and the Submissions rows (not changed by our writes); drop everything read from our tabs.
  const keep = { ver: v };
  ['owner', 'ours', 'subs'].forEach(function (k) { if (MEMO_[k] !== undefined) keep[k] = MEMO_[k]; }); // closed / required re-read
  MEMO_ = keep;
}

// One of our tabs as rows_, cached under its own stamp (changed only by writes to that tab). Writes read the Sheet
// directly, never this.
function tabRows_(name) {
  return memo_('rows:' + name, function () {
    const c = CacheService.getScriptCache();
    const stamps = memo_('tv', function () { return c.getAll(Object.keys(TABS).map(function (n) { return 'tv:' + n; })); });
    const k = 'rows:' + name + '@' + (stamps['tv:' + name] || '0');
    const hit = cacheGet_(c, k);
    if (hit !== undefined) return hit;
    const v = rows_(tab_(name));
    cachePut_(c, k, v);
    return v;
  });
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
  } else return sh;
  // File twins are for the site only.
  TABS[name].forEach(function (h, i) { if (h.slice(-TWIN.length) === TWIN) sh.hideColumns(i + 1); });
  return sh;
}

// Rows as objects keyed by the header row. File columns come back as [{ id, name }] (filesIn_).
function rows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return objects_(sh, head, v, 2);
}

// Values of consecutive sheet rows starting at row `first` -> objects, as rows_ (blank key rows left out).
function objects_(sh, head, v, first) {
  const fileCols = head.map(function (h, i) { return FILE_COLS.indexOf(h) >= 0 ? i : -1; }).filter(function (i) { return i >= 0; });
  // A file column is read from its text twin; the slow rich-text read runs only when some filled cell has no twin yet
  // (rows saved before v2.5 and not backfilled).
  const twin = {}; const rich = {};
  fileCols.forEach(function (i) {
    const t = head.indexOf(head[i] + TWIN);
    if (t >= 0) twin[i] = t;
    if (t < 0 || v.some(function (r) { return r[i] !== '' && r[t] === ''; })) rich[i] = sh.getRange(first, i + 1, v.length, 1).getRichTextValues();
  });
  return v.map(function (r, n) {
    const o = {};
    head.forEach(function (h, i) {
      if (i in twin && r[twin[i]] !== '') o[h] = files_(parse_(String(r[twin[i]])));
      else o[h] = i in rich ? filesIn_(rich[i][n][0], r[i]) : r[i];
    });
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
  writeRow_(sh, i >= 0 ? i + 2 : sh.getLastRow() + 1, cols, row);
}

// Sheet row number of a key (Project Ref / Org Ref, from the first column), 0 = none.
function rowOf_(sh, ref) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const i = sh.getRange(2, 1, n, 1).getValues().map(function (r) { return r[0]; }).indexOf(ref);
  return i < 0 ? 0 : i + 2;
}

// One sheet row as an object, as rows_ gives it.
function rowAt_(sh, at) {
  const w = sh.getLastColumn();
  return objects_(sh, sh.getRange(1, 1, 1, w).getValues()[0], sh.getRange(at, 1, 1, w).getValues(), at)[0] || null;
}

// Writes `row` over its project's row (at = its row number if already known). An existing row is overwritten in one
// write, without the lock (two saves of one project at once: the last wins). A new row is added under the lock, after
// looking again, so two first saves of one project can't both append.
function saveRow_(sh, cols, row, at) {
  const key = row[cols[0]]; // Project Ref, or Org Ref on OrgPrePitch
  at = at || rowOf_(sh, key);
  if (at) { writeRow_(sh, at, cols, row); return; }
  locked_(function () { writeRow_(sh, rowOf_(sh, key) || sh.getLastRow() + 1, cols, row); });
}

function writeRow_(sh, at, cols, row) {
  // A file twin holds its column's files as JSON text ('[]' = none).
  const twinOf = function (c) { const base = c.slice(0, -TWIN.length); return c.slice(-TWIN.length) === TWIN && Array.isArray(row[base]) ? JSON.stringify(row[base]) : null; };
  sh.getRange(at, 1, 1, cols.length).setValues([cols.map(function (c) { const t = twinOf(c); return t != null ? t : Array.isArray(row[c]) ? '' : cell_(row[c]); })]);
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

/* ---------- daily backup (v2.6.1) ---------- */
// A dated copy of our Sheet in "TUSU Pitch 2 Backups" (the script owner's My Drive, private), keeping the latest 14.
// Runs from a daily time-driven trigger added once in the editor (Triggers -> dailyBackup -> Day timer). Google emails
// the owner if a run fails. Restoring is manual: open a backup and copy the rows back.
const BACKUP_FOLDER = 'TUSU Pitch 2 Backups';
const BACKUP_PREFIX = 'TUSU Pitch 2 Feedback - backup ';
const BACKUP_KEEP = 14;
function dailyBackup() {
  const root = DriveApp.getRootFolder();
  const found = root.getFoldersByName(BACKUP_FOLDER);
  const folder = found.hasNext() ? found.next() : root.createFolder(BACKUP_FOLDER);
  const name = BACKUP_PREFIX + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!folder.getFilesByName(name).hasNext()) DriveApp.getFileById(FEEDBACK_SHEET_ID).makeCopy(name, folder);
  const all = [];
  const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf(BACKUP_PREFIX) === 0) all.push(f); }
  all.sort(function (a, b) { return a.getName() < b.getName() ? 1 : -1; }); // newest first (dates sort as text)
  all.slice(BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });
  Logger.log('Backup ' + name + ' in ' + folder.getUrl() + ' (' + Math.min(all.length, BACKUP_KEEP) + ' kept)');
}

/** Run once from the editor after deploying v2.5: fills the hidden file twins of rows saved before them. */
function backfillFileJson() {
  ['PrePitch', 'Pitching'].forEach(function (name) {
    const sh = tab_(name);
    if (sh.getLastRow() < 2) return;
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    FILE_COLS.forEach(function (c) {
      const i = head.indexOf(c); const t = head.indexOf(c + TWIN);
      if (i < 0 || t < 0) return;
      const rich = sh.getRange(2, i + 1, v.length, 1).getRichTextValues();
      const raw = sh.getRange(2, i + 1, v.length, 1).getValues();
      sh.getRange(2, t + 1, v.length, 1).setValues(v.map(function (r, n) { return [r[0] === '' ? '' : JSON.stringify(filesIn_(rich[n][0], raw[n][0]))]; }));
    });
  });
  bump_('PrePitch', 'Pitching');
  Logger.log('File twins filled.');
}

/** Run once from the editor: authorizes the scopes and creates our tabs. */
function setup() {
  Object.keys(TABS).forEach(tab_);
  DriveApp.getFolderById(DRIVE_FOLDER_ID).getName();
  SpreadsheetApp.openById(prop_('SUBMISSIONS_SHEET_ID')).getName();
  Logger.log('Ready. Owner (always Admin): ' + owner_());
}
