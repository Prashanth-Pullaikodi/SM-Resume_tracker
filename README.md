# Resort Recruitment System

Mobile-first Progressive Web App for managing resort recruitment, built on
Google Apps Script + Google Sheets with role-based access control.

## Architecture

```
┌─────────────────────────────┐
│  Index.html (SPA + PWA)     │  ← Bootstrap 5, mobile-first
│  - Login gate               │
│  - Dashboard / KPIs         │
│  - Candidates / Interviews  │
│  - Analytics + Funnel       │
└──────────────┬──────────────┘
               │ google.script.run
┌──────────────▼──────────────┐
│  Code.gs (backend)          │
│  - RBAC (Admin/HR/INT/View) │
│  - CRUD + analytics         │
│  - Audit log                │
└──────────────┬──────────────┘
               │
┌──────────────▼──────────────┐
│  Google Sheets DB           │
│  Candidates / Interviews    │
│  StatusHistory / Users      │
│  AuditLog                   │
└─────────────────────────────┘
```

## File List

Google Apps Script projects only allow `.gs` and `.html` files, so every
artifact in this repo maps directly to one of those two types.

| Repo file | Apps Script file | Type | Purpose |
|-----------|------------------|------|---------|
| `Code.gs` | `Code.gs` | Script | RBAC backend, CRUD, analytics, doGet router |
| `Index.html` | `Index.html` | HTML | Mobile-first SPA frontend |
| `Manifest.html` | `Manifest.html` | HTML | PWA manifest JSON (served as JSON via `?page=manifest`) |
| `ServiceWorker.html` | `ServiceWorker.html` | HTML | Service worker JS (served as JS via `?page=sw`) |
| `README.md` | — | — | This file (not added to the Apps Script project) |

> `Manifest.html` and `ServiceWorker.html` wrap the raw JSON / JS in a
> commented HTML container. `Code.gs` strips the wrapper and serves the
> inner content with the correct MIME type via `doGet`.

---

## 1. Google Sheets Setup

You can either let `setupSheets()` create all sheets automatically, or
create them by hand.

### Option A — Automatic (recommended)

1. Open Google Drive → New → **Google Sheets**. Name it `Resort Recruitment DB`.
2. Extensions → **Apps Script** → name the project `Resort Recruitment`.
3. Paste contents of `Code.gs` into the default `Code.gs`.
4. Create three HTML files (**File → New → HTML file**) with these *exact* names
   (case-sensitive — Apps Script appends `.html` for you):
   - `Index` → paste `Index.html`
   - `Manifest` → paste `Manifest.html`
   - `ServiceWorker` → paste `ServiceWorker.html`
5. Save the project, then in the editor run **`initialSetup`** once.
   - Grant the requested OAuth scopes.
   - This creates all 5 sheets with headers, seeds you as the first **Admin**,
     and populates sample users, candidates, interviews, and status history
     so the dashboard and analytics screens have data to show.
   - Safe to re-run — if candidates already exist, sample data is skipped.
   - If you want headers only (no sample data), run `setupSheets` instead.

### Option B — Manual

Create these 5 sheets with these exact headers in row 1:

**Candidates**
```
CandidateID | Name | Phone | Email | RoleApplied | ResumeLink | Source | Status | CreatedAt
```

**Interviews**
```
InterviewID | CandidateID | Round | Interviewer | DateTime | Status | Feedback | Score
```

**StatusHistory**
```
LogID | CandidateID | OldStatus | NewStatus | ChangedBy | Timestamp
```

**Users**
```
UserID | Name | Email | Role
```

Add at least one row to **Users** with your own email and `Role = Admin`.

**AuditLog**
```
LogID | Email | Action | Details | Timestamp
```

### Allowed values

- **Status**: `New`, `Not Screened`, `Shortlisted`, `Interviewed`, `Selected`, `Rejected`, `On Hold`
- **Role**: `Admin`, `HR`, `Interviewer`, `Viewer`
- **Round**: `HR`, `Technical`, `Manager`

---

## 2. Deploy as a Web App

1. In the Apps Script editor click **Deploy → New deployment**.
2. Select type **Web app**.
3. Settings:
   - **Description**: `Resort Recruitment v1`
   - **Execute as**: **User accessing the web app** (so `Session.getActiveUser()` returns the visitor's email — required for RBAC).
   - **Who has access**: `Anyone within <your Workspace domain>` (recommended) or `Anyone`.
4. Click **Deploy**, authorize, then copy the **Web App URL**.
5. Open the URL on a phone → Chrome → menu → **Install app / Add to Home Screen**.

> ⚠️ If you choose **Execute as: Me**, every user will look like the script
> owner and RBAC will not work. Always pick **User accessing the web app**.

> ℹ️ Google Apps Script Web Apps serve content from a `userContent.google.com`
> sandbox iframe. Browser support for service workers in that sandbox is
> limited; the app falls back gracefully to online-only when SW isn't allowed.
> The PWA manifest, "Add to Home Screen", responsive layout, and bottom
> navigation still work.

---

## 3. RBAC Matrix

| Capability | Admin | HR | Interviewer | Viewer |
|-----------|:-----:|:--:|:-----------:|:------:|
| View dashboard | ✓ | ✓ | ✓ (assigned only) | ✓ |
| Add / edit candidates | ✓ | ✓ | — | — |
| Schedule interviews | ✓ | ✓ | — | — |
| Update candidate status | ✓ | ✓ | — | — |
| Add interview feedback | ✓ | ✓ | ✓ (own only) | — |
| Manage users | ✓ | — | — | — |
| Export CSV | ✓ | — | — | — |

- A user not present in the **Users** sheet is blocked with `Access Denied`.
- All backend mutations call `authorizeUser([...roles])` before running.
- Interviewers can only see candidates with interviews assigned to their email,
  and can only edit feedback on their own interviews.

---

## 4. PWA / Mobile

- Manifest is served from `?page=manifest` and linked in `Index.html`.
- Service worker is served from `?page=sw` and registered on boot.
- Bottom navigation: Dashboard · Candidates · Interviews · Analytics · Profile.
- All buttons are touch-friendly; modals are used for Add/Detail/Feedback.
- Theme color `#0d6efd` (Bootstrap primary).

### Icons

Apps Script can't easily host binary assets, so the manifest references a
Google-hosted PNG as a fallback. To use custom icons:

1. Upload `icon-192.png` and `icon-512.png` to a public URL (Google Drive
   public link, GitHub Pages, or your CDN).
2. Edit `getManifestJson()` in `Code.gs` and replace the `icons[].src` URLs.

---

## 5. Backend API (`google.script.run` callable)

```
getCurrentUser()                  → { authorized, email, name, role, permissions }
getUserRole(email)                → "Admin" | "HR" | "Interviewer" | "Viewer" | null

getCandidates()
getCandidate(id)
addCandidate(data)                // Admin, HR
updateCandidate(data)             // Admin, HR
updateStatus(id, newStatus)       // Admin, HR
getNotScreenedCandidates()

addInterview(data)                // Admin, HR
getInterviews()
getInterviewsByCandidate(id)
updateInterviewFeedback(data)     // Admin, HR, Interviewer (own only)

getUsers()                        // Admin
addUser({ Name, Email, Role })    // Admin
deleteUser(userId)                // Admin

getDashboardKpis()
getConversionMetrics()
getRoleWiseStats()
exportCandidatesCsv()             // Admin
```

---

## 6. Performance

- `bootstrapApp()` returns user + candidates + interviews + KPIs + analytics
  in **one round trip** so the frontend only hits the network on load and on
  writes — every navigation, candidate click, and tab switch renders from
  the in-memory cache instantly.
- The bootstrap payload is **cached in `CacheService`** (script-wide, 5 min
  TTL). The first user pays the sheet-read cost; everyone else gets <100 ms
  responses. Every write call (via `logAudit_`) invalidates the cache so
  changes are visible immediately.
- The `getUserRole(email)` lookup is **cached per email** (10 min TTL) so
  the Users sheet isn't re-read on every backend call. Add/delete user
  invalidates that key.
- Date cells are coerced to ISO strings in `sheetToObjects_` so
  `google.script.run` can serialize the payload reliably.

If the cache is hot, total round-trip is typically **300–700 ms**.

## 7. Bonus Features (implemented)

- ✅ Duplicate detection on candidate email / phone
- ✅ Auto UUID generation (`Utilities.getUuid()`) for all IDs
- ✅ Audit log (`AuditLog` sheet) for every write action
- ✅ CSV export (Admin only) via the Profile screen
- ✅ Status change history in `StatusHistory`
- ✅ **Resume upload to Google Drive** — accepts PDF, DOC/DOCX, RTF, TXT,
  ODT, and images (any MIME type really; the picker is a hint). Files land
  in a Drive folder named **Resort Recruitment Resumes** (auto-created),
  are shared with the link, and the URL is written back to the candidate's
  `ResumeLink` field. Available on the Add Candidate form and on the
  candidate detail modal. 25 MB cap per file.

### Pinning the upload folder

By default, every upload lands in a Drive folder named **Resort Recruitment
Resumes** that's auto-created in the script owner's *My Drive* on first use.
To force uploads into a specific folder (recommended — works with Shared
Drives and folders you've already shared with HR):

1. Open the target folder in Drive.
2. Copy the ID from the URL: `https://drive.google.com/drive/folders/`**`<this part>`**
3. In `Code.gs`, set:
   ```js
   RESUMES_FOLDER_ID: 'PASTE_THE_ID_HERE',
   ```
4. Save and redeploy a new version. All future uploads go into that folder.

If `RESUMES_FOLDER_ID` is blank, the script falls back to the named folder
behaviour above.

### Drive permissions

Uploading triggers an extra OAuth scope (`drive`). The first time you click
upload after deploying a new version, Apps Script will prompt you to grant
access. Re-deploy a new version of the Web App so users see the consent
screen the next time they open the app.

---

## 8. Updating the App

After editing `Code.gs` or `Index.html`:

1. **Deploy → Manage deployments**.
2. Click ✏️ on the active deployment.
3. Version: **New version**, then **Deploy**.

The Web App URL stays the same.

---

## 9. Troubleshooting

| Problem | Fix |
|--------|-----|
| "Access Denied" for every user | You deployed as **Execute as: Me**. Redeploy as **User accessing the web app**. |
| "Sheet not found" | Run `setupSheets()` from the Apps Script editor. |
| Empty dashboard | The Sheets are empty — add a candidate via the **+** button. |
| Service worker not registering | Apps Script sandbox doesn't allow SW. PWA install + manifest still work. |
| Locked out of Users sheet | Open the Sheet directly and edit the `Users` row to restore your `Admin` role. |
