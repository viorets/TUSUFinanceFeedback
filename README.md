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

Files are stored as JSON `{ id, name }`. Tabs are created by `setup()` / on first use. Rows are keyed by Project Ref
(organization code + `-P<nn>`, position in the latest submission; see the open question in the plan).

**Rules (server-side, `Code.gs`)**
- Roles: Admin = everything (admin console + edit log). Chief = everything except the edit log, and cannot add, change or remove Admins. Staff = edit Pre-pitch and Pitching, upload; no admin console. Emails not in `Users` (or with an old Viewer role) get "No access". Old Reviewer rows count as Staff. The script owner is always Admin.
- Approval finished (`finished_`): Approved budget filled, and .docx and PDF each Yes, or No with a feedback file uploaded in that field. Unfinished projects are counted on the list and the detail page.
- Organizations whose last remove/restore in `Submission_Log` is a remove are hidden from the list and shown in the admin console.
- Uploads: .docx or PDF only, 25 MB max (Apps Script limit), saved to the organization's subfolder. A save only accepts files inside that subfolder.
- Re-submission tab is a display-only prototype: the resubmission website is not built yet.

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
