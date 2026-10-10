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
| `PrePitch` | project | Project Ref, Org Ref, Project Name, Approved Budget, Docx Approved, Docx Feedback File, PDF Approved, PDF Feedback File, Feedback Note, Updated At, Updated By |
| `PrePitch_Log` | changed field (append-only) | Logged At, By, Project Ref, Org Ref, Field, Old, New |
| `Pitching` | project | Project Ref, Org Ref, Project Name, Summary, Final Approved Budget, Noted Documents, Updated At, Updated By |
| `Resubmission` | project | Project Ref, Org Ref, Project Name, Decision, Updated At, Updated By |
| `Settings` | setting | Key, Value, Updated At, Updated By (`Allocated Budget`; `Required` = JSON of the Owner's required questions) |
| `OrgStatus` | organization | Org Ref, Status (Closed / Unclosed), Updated At, Updated By (v2) |
| `Texts` | changed line of site text | Key (built-in English line), EN, TH, Updated At, Updated By (v2) |

Files are stored as JSON `{ id, name }`. Tabs are created by `setup()` / on first use. Rows are keyed by Project Ref
(organization code + `-P<nn>`, position in the latest submission; see the open question in the plan).

**Rules (server-side, `Code.gs`)**
- Roles: Admin = everything (admin console + edit log). Chief = everything except the edit log, and cannot add, change or remove Admins. Staff = edit Pre-pitch and Pitching, upload; no admin console. Emails not in `Users` (or with an old Viewer role) get "No access". Old Reviewer rows count as Staff. The script owner is always Admin.
- Approval finished (`finished_`): every required Pre-pitch question answered. Defaults: Approved budget filled, and .docx and PDF each Yes, or No with a feedback file uploaded in that field. Unfinished projects are counted on the list and the detail page.
- Closed cases (v2, moved in v2.2): Admins and Chiefs close or open an organization's case at the top of its Re-submission tab; the list shows Open / Closed. Every write to a closed organization (Pre-pitch, Pitching, Re-submission, uploads) is refused; the person in charge can still change.
- Owner (v2) = the script owner. Only the Owner sees the Control tab: site text in English and Thai (`Texts`; the built-in English line is the key, `{n}` marks stand for values) and Required / Not required per feedback question (`QUESTIONS` in `Code.gs` and `public/js/project.js`).
- Pre-pitch save (v2): a window confirms the organization, project number, project name and approved budget (OK and Cancel both close it); a failed save shows an error window and keeps the edits. Leaving a page with unsaved edits opens a window naming each unsaved project (Save and leave / Leave without saving / Stay). Closing the tab or refreshing is not guarded.
- Print (v2): the organization page's Print button prints the master data and every project on all three tabs, without progress tracking, "Saved by" lines or the person in charge.
- Organizations whose last remove/restore in `Submission_Log` is a remove are hidden from the list and shown in the admin console.
- Uploads: .docx or PDF only, 25 MB max (Apps Script limit), saved to the organization's subfolder. A save only accepts files inside that subfolder.
- Re-submission tab is a prototype: the resubmission website is not built yet, so its fields show "No data submitted yet". The decision saves.

- Pitching phases (v2.1): Admins and Chiefs switch Pre-pitch, Pitching and Re-submission On / Off in the admin console (`Settings` key `Phases`, default all On). An Off phase shows "Pending Phase" on every organization page and in print, and its writes and uploads are refused.

## Design

v2.1 follows Apple's newest look: a floating translucent ("glass") top bar, rounder layered cards, filled fields,
iOS switches and short spring animations, in the TUSU teal. The v2.1 block at the end of `public/css/app.css`
overrides the earlier styles. Glass turns solid for reduced transparency; motion stops for reduced motion.

- v2.2: Pre-pitch is for self-noting ("Note"; nothing required, no tracking). Unfinished counts and filters cover Pitching and Re-submission only. Pitching has a "Correct Excel file" upload (.xlsx or .xls, column `Correct Excel File`). The list shows Total proposed project budget and Total approved budget (Pitching final approved budgets).

## Versions

- `v2.2` (tag, `main`): v2.0 look, v2.1 features, and the v2.2 changes above.
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
