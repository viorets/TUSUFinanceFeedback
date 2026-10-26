# TUSU Pitch Feedback

Reviewers record feedback, approvals and pitching results for every project submitted on the
[budget pitching site](https://tusu-budget-pitching.web.app) ([repo](https://github.com/pppainkiller/budget-pitching-form-submission)).
English by default with a Thai toggle.

## Architecture

| Part | What it does |
| --- | --- |
| `public/` | Vanilla JS ES modules, no framework, no build step, hash router. Each detail tab's code loads only when opened; one Apps Script call per page view (reads are reused for 60 s). Same Apple-HIG tokens as the pitching site. |
| Firebase (our own project) | Hosting and Google sign-in only. No Firestore. |
| `apps-script/Code.gs` | Web app running as the owner. Verifies the sign-in token, checks the role, reads the pitching site's Submissions Sheet, writes our Sheet and Drive folder. |
| `public/js/masterdata.js` | Copied from the pitching repo (`public/js/masterdata.js`): campuses, categories, all organizations. Copy it again when that file changes. |

**Reads** the pitching site's Sheet "TUSU Budget Pitching — Submissions" (tabs `Organization`, `Project`, `Submission_Log`; never written).
**Writes** [our Sheet](https://docs.google.com/spreadsheets/d/1tDak_LDOEvLEuMMa3Wri9Ej3tWjwLcEDS4tvFb-0N0k/edit) and
[our Drive folder](https://drive.google.com/drive/folders/191_2g6tBQwxIJrmUmWMuwECPiRd1Oy2o) (`<Org Ref>/<Project Ref> feedback|noted — <file>`).

| Tab (ours) | One row per | Columns |
| --- | --- | --- |
| `Users` | person | Email, Role (Admin / Chief / Staff), Org Refs (in charge of), Updated At, Updated By, Name |
| `PrePitch` | project | Project Ref, Org Ref, Project Name, Approved Budget, Docx Approved, Docx Feedback File, PDF Approved, PDF Feedback File, Feedback Note, Updated At, Updated By, Docx Feedback File JSON, PDF Feedback File JSON (hidden) |
| `PrePitch_Log` | changed field (append-only) | Logged At, By, Project Ref, Org Ref, Field, Old, New |
| `Pitching` | project | Project Ref, Org Ref, Project Name, Summary, Final Approved Budget, Noted Documents, Updated At, Updated By, Correct Excel File (empty since v2.9), Noted Documents JSON, Correct Excel File JSON (hidden) |
| `OrgPitching` | organization | Org Ref, Correct Excel File, Updated At, Updated By, Correct Excel File JSON (hidden) (v2.9) |
| `Resubmission` | project | Project Ref, Org Ref, Project Name, Decision, Updated At, Updated By, Feedback (v2.10) |
| `Settings` | setting | Key, Value, Updated At, Updated By (`Allocated Budget`; `Required` = JSON of the Owner's required questions) |
| `OrgStatus` | organization | Org Ref, Status (Closed / Unclosed), Updated At, Updated By (v2) |
| `Texts` | changed line of site text | Key (built-in English line), EN, TH, Updated At, Updated By (v2) |

Files are stored as JSON `{ id, name }`. Tabs are created by `setup()` / on first use. Rows are keyed by Project Ref
(organization code + `-P<nn>`, position in the latest submission; see the open question in the plan).

**Rules (server-side, `Code.gs`)**
- Roles: Admin = everything (admin console + edit log). Chief = everything except the edit log, and cannot add, change or remove Admins. Staff = edit Pre-pitch and Pitching on any project, decide Re-submission, pick persons in charge (v2.10), upload; no admin console. Data team staff (v2.5) = Staff + the Control tab + the edit log (admin console shows only the log). Emails not in `Users` (or with an old Viewer role) get "No access". Old Reviewer rows count as Staff. The script owner is always Admin. Second Gmail (v2.11): `Users` column `Email 2` lets one person sign in with up to 2 Gmail addresses; the second is read as the main `Email` (`mainEmail_`), so role, name, persons in charge and logs follow it. It can't be the owner's email or listed twice, and a Chief can't change an Admin's or the owner's.
- Approval finished (`finished_`): every required Pre-pitch question answered. Defaults: Approved budget filled, and .docx and PDF each Yes, or No with a feedback file uploaded in that field. Unfinished projects are counted on the list and the detail page.
- Closed cases (v2, moved in v2.2): Admins, Chiefs and Data team staff (v2.11) close or open an organization's case at the top of its Re-submission tab; the list shows Open / Closed. Every write to a closed organization (Pre-pitch, Pitching, Re-submission, uploads) is refused; the person in charge can still change.
- Owner (v2) = the script owner. The Owner and Data team staff see the Control tab (site text split into tabs by page): site text in English and Thai (`Texts`; the built-in English line is the key, `{n}` marks stand for values) and Required / Not required per feedback question (`QUESTIONS` in `Code.gs` and `public/js/project.js`).
- Pre-pitch save (v2): a window confirms the organization, project number, project name and approved budget (OK and Cancel both close it); a failed save shows an error window and keeps the edits. Leaving a page with unsaved edits opens a window naming each unsaved project (Save and leave / Leave without saving / Stay). Closing the tab or refreshing is not guarded.
- Print (v2): the organization page's Print button prints the master data and every project on all three tabs, without progress tracking, "Saved by" lines or the person in charge.
- Organizations whose last remove/restore in `Submission_Log` is a remove are hidden from the list and shown in the admin console.
- Uploads: .docx, PDF, or Excel where the slot takes it, 12 MB max (v2.12; 10 MB in v2.9), saved to the organization's subfolder. The browser sends the file straight to Drive through a resumable session the script opens (`uploadStart` / `uploadDone`); if that fails it goes base64 through the script (`upload`). A save only accepts files inside that subfolder.
- Re-submission tab is a prototype: the resubmission website is not built yet, so its fields show "No data submitted yet". The decision saves.

- Pitching phases (v2.1): Admins and Chiefs switch Pre-pitch, Pitching and Re-submission On / Off in the admin console (`Settings` key `Phases`, default all On). An Off phase shows "Pending Phase" on every organization page and in print, and its writes and uploads are refused.

## Design

v2.1 follows Apple's newest look: a floating translucent ("glass") top bar, rounder layered cards, filled fields,
iOS switches and short spring animations, in the TUSU teal. The v2.1 block at the end of `public/css/app.css`
overrides the earlier styles. Glass turns solid for reduced transparency; motion stops for reduced motion.

- v2.2: Pre-pitch is for self-noting ("Note"; nothing required, no tracking). Unfinished counts and filters cover Pitching and Re-submission only. Pitching has a "Correct Excel file" upload (.xlsx or .xls, column `Correct Excel File`). The list and each organization page's header show Total proposed project budget and Total approved budget (Pitching final approved budgets); the list's long headings wrap so it fits without sideways scrolling.

## Versions

- `v2.15` (26 Oct): easier filters, one Noted .docx and one Noted PDF, clickable links in exports.
  - Organizations: search on its own line, then one row of filter buttons (Category, Campus, Person in charge, Budget, Round code). A button in use turns teal with its count; the budget range opens inside its button; Choose all / Clear all in every picker; one picker open at a time (Esc or a click outside closes it). A chip per choice under the bar (× removes it) and Clear filters.
  - Person in charge filter: No person in charge yet, Me (replaces "Mine only"; a device that had it on starts with Me ticked) and everyone in charge of an organization, each with a count.
  - Pitching: Noted documents = one Word file (.docx or .doc) and one PDF per project, each Upload / Replace / Remove, either may be empty. Edit in the site saves into that slot (`noted:<ref>:docx|pdf`). The server keeps the newest of each kind (`notedPair_`); rows saved before keep their other files in Drive. Required = at least one of the two. A .doc gets no Edit button.
  - Data for mailing: Feedback files became Feedback DOCX and Feedback PDF. Export data: Noted Documents became Noted DOCX and Noted PDF. Every cell holding one web address is a hyperlink in both Excel files.

- `v2.12` (20 Oct): edit files in the site, timeline, Pre-pitch documents.
  - Pitching shows the submitted Budget proposal form and each project's PDF / .docx with Mark up (PDF) or Edit in Docs / Sheets. The editor opens in a new tab (`#/markup`, `views/markup.js`): PDF marks are drawn in the browser and kept as a server draft ("markup draft <file id>.json" in the organization folder) plus a copy on the device; .docx / .xlsx open as a Google copy in the organization folder, shared only with the signed-in Gmail, showing Google's last edit. Save to site uploads the result (an edited budget form becomes the Correct Excel file with its locks put back; PDFs and .docx are added to Noted documents) and hands it to the open Pitching tab, where Save keeps it. Pre-pitch has no edit buttons.
  - "Show Pre-pitch data" on Pitching: each project's Pre-pitch notes and documents, read only.
  - Pre-pitch: Documents per project (any number of .docx / PDF), column "Documents" in PrePitch.
  - Timeline: a Gantt chart at the top of the Organizations tab (phases, month / week scale, one colour per step, today line, progress on Pitching and Re-submission). Dates in the admin console (Settings "Phase Dates"; defaults from the TUSU calendar: System test 20–31 Oct, Submission 1–20 Nov, Pitching 17–19 Dec, Mailing 20 Dec, Re-submission campus 21–25 Dec and Centre 27–31 Dec, Council 4–21 Jan). Dates don't switch phases yet.
  - Admin console: default text box size of the PDF editor (Settings "Markup Text Size", 11 pt); the editor also has a Text size box.
  - Uploads: 12 MB. Cached answers carry `CODE_VERSION`, so a deployment never reads answers another code version cached.
  - Local test runs: on localhost, `localStorage.scriptUrl` points the page at a test deployment.
- `v2.12` (in progress, earlier on 20 Oct): markup building blocks, deployed but not linked from any page yet. `public/js/markup/pdf-markup.js` (PDF editor: text box LINE Seed Sans TH 11 pt #990000, highlight, draw, line, arrow, rectangle; 12 MB limit), `public/js/markup/xlsx-relock.js` (puts back sheet protection, the workbook structure lock, cell locks, column widths and merges from the original after a Google Sheets edit; tested on the v3 budget forms), `public/css/markup.css`, `public/fonts/LINESeedSansTH_Rg.ttf`. Local test page `public/markup-test.html` (not deployed). Stress test with all users editing at once is still to do (see the plan doc).
- `v2.11.1` (20 Oct): Persons in charge: Staff and Data team staff pick only themselves (`setInCharge` refuses anyone else); ticking yourself where someone else is in charge opens a caution window (Cancel / Add me too / Replace = take over). Admins and Chiefs still pick anyone. Light background: white gradients tinted with the CI teal #006371 (a little cyan #2babc2); the title stays teal.
- `v2.11` (20 Oct, live): Organizations list phase tabs in a sticky frosted-glass bar (80% see-through) with a tracking line on Pitching and Re-submission (projects finished / decided, organizations unfinished); light background one step darker; Dashboard waterfall removed; Data team staff can close and reopen cases; masterdata 95ae1a4a9d (122 organizations, adds CT-OTH-01..05 and 12 more OTH). Speed: caches are keyed by the Submissions and Schedule Sheets' last-updated times (`stamp_`, asked from Drive every 30 s), with 5 min (answers) and 10 min (our tabs) backstops, and the Schedule Sheet is now cached; a cold `boot` went from 13.9 s to about 2 s.
- `v2.11` (`main`): Admin console "Import CSV" under Users and roles: columns Username, Gmail, Gmail#2, comma or tab separated (found by the header row when there is one, else in that order; Gmail#2 may be empty). Invalid emails are ignored; an invalid Gmail falls back to Gmail#2. A row whose email is already listed (on the site or earlier in the file) updates that user's name and second Gmail; other rows join as Staff; rows with no valid email are listed in a window. New emails join the list as Staff; emails already listed or repeated in the file are skipped; Save users writes them (same `saveUsers` checks). Second Gmail per user (`Users` column `Email 2`, set in Users and roles; CSV import skips emails already listed as either Gmail). `Code.gs` changes: push it and deploy a new version on the same `/exec` link.
- `v2.10` (tag): phases compared everywhere, Staff roles widened.
  - Organizations list: tabs All / Pre-pitch / Pitching / Re-submission, each with its own budget column and a "% of total" column; filters (category, campus, budget range with a type picker on All) kept on the device; a mini dashboard; Pitching shows the next pitching time and progress; Re-submission shows Submitted / Not submitted with a version dropdown. `boot` adds per organization `est`, `pp`, `rp`, `ok`, `fin`, `ver`, `vers`, `rs`, `slots` (`Code.gs` `boot_`).
  - Dashboard: tiles on one grid; "% of total" in the project table and Top 5 lists; a waterfall (proposed, pitching change, pitched, re-submission change, final) and a dumbbell by campus or category.
  - Organization page: the submission version beside the RefCode; a notice when someone else is in charge.
  - Staff: pick persons in charge (`setInCharge` with `on`; several people per organization); Not Approved needs written feedback (`Resubmission` tab, column `Feedback`, logged in `Resubmission_Log`).
  - Export: the Pitching sheet follows the Pre-pitch layout (Correct Excel file on each row); both end with a Total row in Excel.
  - Masterdata: not updated in this release (see the plan).
- `v2.9` (tag, `main`): the 13 Oct v2.8 plan, released as v2.9, plus one Correct Excel file per organization.
  - Correct Excel file: one per organization at the top of the Pitching tab (`OrgPitching` tab, `saveOrgPitching`, logged in `Pitching_Log` with an empty Project Ref), no longer per project and no longer a Required question. Print and export follow. After deploying, run `purgeProjectExcel()` once from the editor: it moves every per-project Correct Excel file to the Drive trash, clears the cells and logs each one.
  - Uploads: 10 MB, a "Max 10 MB · <types>" note under each button, straight from the browser to Drive with a progress bar (fallback: the old base64 path).
  - Every save and upload runs in a progress window (`busy` in `ui.js`); the page can't be clicked behind it, and leaving the tab asks first. A save that can't reach the server is kept on this device (`save` / `flushOutbox` in `api.js`), shows "Pending", and is sent when the network comes back, every minute, or on the next visit.
  - Admin console: disabled main buttons stay cyan; "In charge of" folds into a chip; role names in full; a "Preparing export…" window; "Export data for mailing" (Admin only, logged as Mailing): Organizations (emails 1–3, requested / approved budget and change, category and campus benchmarks, four shares, budget form and Correct Excel links) and Projects (contacts, dates in dd/mm/yyyy Buddhist year, budgets, Committee feedback = Pitching summary, proposal and feedback links).
  - My projects: a "Time left" column (days, or hours, or minutes; Now; Pitched), refreshed every minute. Pitching is online in Microsoft Teams: the room is no longer shown and links read "Join on Teams".
- `v2.8` (tag): loophole fixes L1, L2, L8, L9. A Pre-pitch, budget-form, Pitching or Re-submission save is refused ("X saved this at …, reload") when the row changed after the page loaded it (`fresh_`; the page sends the Updated At it had as `base`). Save users is refused when the Users tab changed after the console loaded it (`usersVer_`, a digest of the tab). Right before writing, a save checks the case again in the Sheet itself (`stillOpen_`). Opening a closed case again asks for a reason, logged in `Resubmission_Log` as "Open — <reason>". L10 needed no change: every request reads the role from the Users tab, and Save users refreshes that cache at once.
- `v2.7` (tag): Pre-pitch feedback box on the Budget proposal form (note + one file; `OrgPrePitch` tab, logged in `PrePitch_Log` with an empty Project Ref). "My projects" tab (`#/mine`): every project of my organizations, with Pitching / Re-submission progress, pitching time and meeting link (optional "Meeting Link" column on the scheduling site's Schedule tab). Organization print on one vertical page with a paper choice (Auto = A4 wide and as tall as the content; A4, A3, Letter); campus print is a summary table with project sub-rows. Admin console: Open Google Sheet button, Export first, section icons, "Projects in charge" column, add back organizations removed on the pitching site (Settings `Restored Orgs`, hidden everywhere else), Pitching edit log (`Pitching_Log`). Notification emails (Submissions `Organization` column "Notification Emails", up to 3) on the organization page and in the export (one row per email). Loophole L6: export is Admin-only, asks for confirmation and is logged in `Export_Log`.
- `v2.6.3` (tag): saves no longer queue: updating an existing row is one write without the lock (`saveRow_`); only a new row takes the lock (checked again inside it). Log rows use `appendRow`. The site retries a busy-server error twice (1.5 s, 3 s; `call` in `public/js/api.js`). Load test: 0 errors at 10, 20 and 30 people saving, saves ~6 s typical (max 11 s), no duplicate rows.
- `v2.6.2` (tag): speed after the load test. Each of our tabs is cached on its own (`tabRows_`) and a write clears only the tabs it changed; the Submissions rows are cached apart. Load test (105 organizations, 525 projects, every field filled): page loads 2-3 s typical with 0 errors at 30 people at once (before: 6-9 s, 17% errors). Saves still queue on one lock: 20+ saves in the same moment can time out.
- `v2.6.1` (tag): `dailyBackup()` copies our Sheet each day to "TUSU Pitch 2 Backups" in the script owner's My Drive (private) and keeps the latest 14. Add once in the Apps Script editor: Triggers -> Add trigger -> dailyBackup -> Time-driven -> Day timer.
- `v2.6` (tag): Print redesigned in the site style: a cover (master data, totals, project index), then one page per phase (Pre-pitch, Pitching, Re-submission) listing every project; every document is a link to its Drive file; a window picks the phases (`public/js/print.js`). Print campus on the organization list (one `printData` call). "Mine only" filter on the list. `Resubmission_Log` tab (decision changes, case close / open), shown with the Pre-pitch log in the admin console (Admins, Data team staff). Extra top tab linking to another website, set in Control (Settings key `External Link`).
- `v2.5` (tag): v2.4 + v2.5 in one release. Pre-pitch shows the organization's Budget proposal form (.xlsx) (Submissions `Organization` column "Budget Proposal Form (.xlsx)") and calls its money field "Estimated Approved budget" (column stays `Approved Budget`). Admin console: Export (Excel via SheetJS from cdnjs, or CSV; Include toggles per phase; downloads on Save) and a Save button for pitching phases. A case closes only when every project is Approved. Data team staff role. Unsaved-changes window offers Save and leave on every page with a Save. Dashboard in the Mosaic Blue palette (#00758F), Total row first in All project budgets. Speed: gzipped, chunked script cache; hidden JSON twins of file columns so reads skip rich text — after deploying, run `backfillFileJson()` once from the editor; edit file cells only through the site.
- `v2.3.3` (tag): the account menu opens above the tabs bar (its name and email were hidden).
- `v2.3.2` (tag): blue-and-white page background (soft blue glows on white; navy at night), whiter top bar.
- `v2.3.1` (tag): the organization list page is wider (up to 1600 px).
- `v2.3` (tag): a closed case counts as Done on the list (filter, counts, Unfinished column) and its tabs show no unfinished badges; each campus is its own card on the list, with space between.
- `v2.2.1` (tag): budget totals on the organization page; the list fits the page width.
- `v2.2` (tag): v2.0 look, v2.1 features, and the v2.2 changes above.
- `v2.1` (tag): pitching phases and the glass design refresh (design later reverted).
- `v2` (branch and tag): save window, closed organizations, Owner control tab, print. Kept unchanged.
- `v1-prototype` (branch and tag): the first prototype, kept unchanged.

## Setup

1. **Firebase**: create a new Firebase project → add a Web app → copy its config into `firebaseConfig` in `public/js/config.js`, and its project id into `.firebaserc`. Authentication → Sign-in method → enable **Google**.
2. **Apps Script**, as the account that owns our Sheet and Drive folder and can open the Submissions Sheet:
   1. https://script.google.com → New project. Paste `apps-script/Code.gs`; Project Settings → "Show appsscript.json" → paste `apps-script/appsscript.json`.
   2. Project Settings → Script properties: `FIREBASE_API_KEY` = our `firebaseConfig.apiKey`; `SUBMISSIONS_SHEET_ID` = the id in the Submissions Sheet's URL.
   3. Run `setup()` once and authorize (creates the tabs, checks access to both sheets and the folder).
   4. Deploy → New deployment → Web app → Execute as **Me**, Who has access **Anyone** → copy the `/exec` URL into `SCRIPT_URL` in `public/js/config.js`.
      Later changes: Deploy → Manage deployments → Edit → Version: **New version** (keeps the same URL).
3. **Drive sharing**: uploaded files inherit the Drive folder's sharing. Reviewers need at least view access to the folder for the download buttons; the pitching site's proposal files are already "anyone with the link".
4. **Deploy**: `firebase deploy --only hosting`. Sign in as the script owner and add people in Admin console.

## Local preview

Any static server, e.g. `npx serve public`. Google sign-in works on `localhost` (an authorized domain by default).

## Licence

This project is released under the [GNU General Public License v3.0](LICENSE) (GPL-3.0), as required by [HyperFormula](https://hyperformula.handsontable.com/), which calculates the Excel formulas in the in-page Excel editor and is used under GPL-3.0. The other libraries and fonts keep their own licences, all of which can be used in a GPL-3.0 project: pdf.js and the Firebase SDK (Apache-2.0), pdf-lib, ExcelJS and JSZip (MIT), and LINE Seed Sans TH and Noto Sans Thai (SIL Open Font License). Editor fonts in `public/fonts` (v2.13, `css/fonts.css`): TH SarabunPSK and TH Sarabun New (SIPA, free to share), Sarabun and Noto Sans Thai (SIL OFL), and Carlito, Arimo and Tinos (SIL OFL), which stand in for Calibri, Arial / Tahoma and Times New Roman on devices without them. Microsoft fonts are never shipped: their names use the installed font where there is one.
