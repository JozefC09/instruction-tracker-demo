# Library Instruction Tracker (web)

A web app that runs on the university website and works directly on the instruction workbooks in SharePoint
(`2026-2027 Instruction Classes.xlsx` and the others in the Instruction folder). Staff sign in with their DeSales Microsoft 365
account, and every change is written straight into the SharePoint workbook through the Microsoft Graph Excel API.
The workbook stays the system of record, and anyone can still open it in Excel.

This replaces the desktop `.exe` described in [spec.md](../spec.md). Data model, concurrency and dashboard follow that spec.
Only the delivery is different: a static web page instead of a Windows program.

| Page | Who | What |
|---|---|---|
| `index.html` | Librarians, coordinator, leadership | Dashboard, sessions by term with stage views, requests inbox, tours, settings |
| `request-form.html` | Faculty (Phase 2) | Request a session. Posts to a Power Automate flow, never to the workbook directly |

## What's in it

- **Dashboard:** summary tiles, a month/week calendar colored by librarian, and bar charts for sessions by librarian
  (sessions or students), bookings by program, and bookings by college/school. Filters cover student type, librarian, college/school,
  format and session type. Clicking a bar filters by it. The period can be this term, any term, the academic year, a custom
  range, or a past year from `Archive/`. Every chart has a "View as table" toggle and **Copy chart**, which copies it as a picture (PNG, white background, with the period and filters under the title) to paste into slides or email. Browsers that can't copy pictures download the PNG instead. **Export report** writes an `.xlsx`.
- **Sessions:** one tab per term sheet and stage views (Request / Librarian / Planning / After Event). Each stage has a
  "needs attention" filter. There's a "needs review" filter for times and dates that couldn't be read, plus search,
  **Export view** (`.xlsx`) and **Open in Excel**.
- **Editor:** a side panel. It gives real date and start/end time pickers, dropdowns from the lists, and College/Student Type
  suggested from the course number. It warns when a date is outside the term. It also has Duplicate, Move to term, Delete,
  and **Add to my calendar** (`.ics`, with a stable UID so re-exporting updates the same Outlook event).
- **Several staff at once:** the page re-reads the workbook every 30 s, and changed rows flash. Saves use optimistic
  concurrency on `_RowVersion`: if someone else changed *other* fields, both changes are kept. If they changed the *same*
  field, a dialog shows Theirs and Mine side by side. Edits made directly in Excel are detected by a hash of the row.
  Rows are always found by `_RecordId`, never by row number.
- **Requests inbox:** reads `Instruction Requests.xlsx`. **Accept** opens a prefilled new session in the right term and links
  the request to it. **Decline** records a note. **Copy email** copies the faculty address.
- **Prepare workbook** (Settings): a one-time, re-runnable migration. It turns term sheets into Excel tables and adds the new
  columns plus the hidden system columns. It also gives every row an id, fills blank College/Student Type cells with
  suggestions, creates the hidden `Lists` sheet (term dates, Closed flags, lock) and `Instruction Requests.xlsx`.
  On each term sheet it also makes the two column changes from Kathleen's doc (spec §7 steps 4–5):
  - **Entered into LibInsight** is renamed `_LegacyLibInsight` and hidden. Its values stay in the workbook, out of view.
  - The **unlabeled column** (today's column O): its values are appended to the row's Notes, then the column is emptied and
    becomes **Actual Student Attendance**, in the same place.

  Finally it writes the dropdown lists to the Lists sheet and points Excel's dropdowns at them (see below).
  Re-running it is safe. Apart from moving the stray column-O text into Notes, it never overwrites existing text.
- **Excel dropdowns:** the dropdowns people see when they edit the workbook directly in Excel (Term Year, College/School,
  Student Type, Librarian, Session Type, Format, Location, Room Status, Assessment) come from the same lists as the app.
  After editing a list in `config.js`, press **Settings → Update Excel dropdowns**. Excel warns about a value that isn't in the
  list (with "continue anyway"); Location only suggests. Starting a new year updates the Term Year list by itself.
- **Close term / Start new academic year:** see below.
- **Demo mode:** with no `clientId` in `config.js` (or `?demo=1`), the app runs on built-in sample data. Use it to try
  everything. Settings has buttons that simulate a colleague editing a session (to see the conflict dialog) and a colleague
  starting the new year from another PC (to see this page pause and then switch). Reload the page to reset the demo.

## Files

```
index.html            app page
request-form.html     faculty request form (Phase 2)
config.js             ← the only file to edit: Entra IDs, SharePoint folder link, flow URL, dropdown lists, term dates
css/app.css
js/app.js             shell, sign-in, routing, 30-second sync
js/auth.js            Microsoft 365 sign-in (MSAL.js, auth code + PKCE, no secrets)
js/graph.js           Microsoft Graph Excel API calls (with retry on 423/409/429/5xx)
js/store.js           save/merge/conflict logic, add/delete/move, requests, Prepare workbook
js/model.js           parsing (times, dates, prep), suggestions, filters, dashboard maths, .ics
js/dashboard.js, sessions.js, requests.js, settings.js, ui.js
js/demo.js            in-memory sample workbooks for demo mode
js/newyear.js         Start new academic year wizard
js/chartpng.js        Copy chart (draws a dashboard chart as a PNG)
tests.html            open in a browser: 89 tests (parsers, merge, aggregation, .ics, concurrency A2–A6, close term, lock, year rollover, legacy columns, dropdowns)
serve.ps1             local test server: powershell -ExecutionPolicy Bypass -File serve.ps1
```

There is no build step. Copy the folder to the web server as-is.

## Setup

### 1. IT: register the app in Entra ID (one time)

1. Entra admin center → **App registrations → New registration**
   - Name: `Library Instruction Tracker`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: platform **Single-page application (SPA)**, with the exact URL the page will be served from, e.g.
     `https://www.desales.edu/library/instruction-tracker/index.html`. Add `http://localhost:8080/` as well for testing.
2. **API permissions** → Microsoft Graph → **Delegated**: `User.Read`, `Files.ReadWrite.All` → **Grant admin consent**.
   (If IT prefers, `Sites.Selected` scoped to the Library Staff – Instruction Services site also works. The app only uses
   delegated access, so each person can only reach files they could already open in SharePoint.)
3. No client secret is needed. Copy the **Application (client) ID** and **Directory (tenant) ID**.

### 2. Coordinator: fill in `config.js`

- `clientId`, `tenantId` from step 1.
- `instructionFolderUrl`: in SharePoint, open the Instruction folder → **Share** / **Copy link** and paste the link.
  The app finds the newest `YYYY-YYYY Instruction Classes.xlsx` there, the `Archive/` subfolder, and `Instruction Requests.xlsx`.
- Check `lists` (librarians and their chart colors, colleges, course-prefix → college map, locations) and the `terms` date
  ranges. The college list and prefix map are placeholders until the official list is confirmed (spec Q10).

### 3. Web team: publish on the university website

- Upload the folder to any static location on the site (no server-side code, database or PHP needed).
- The page loads two libraries from CDNs: `cdn.jsdelivr.net` (MSAL) and `cdnjs.cloudflare.com` (SheetJS for Excel export).
  If the site's Content-Security-Policy blocks them, download those two files, put them next to `index.html`, and change
  the two `<script src>` lines. The CSP must also allow `connect-src https://graph.microsoft.com https://login.microsoftonline.com`.
- **Embedding in LibGuides or a CMS page in an iframe** works. Inside an iframe, sign-in automatically uses a popup instead of
  a redirect. The iframe's page URL must be the registered redirect URI, and pop-ups must be allowed for the site.
- The page is staff-only in practice: anyone can load it, but without a DeSales sign-in that has access to the
  SharePoint folder it shows nothing. `noindex` keeps it out of search engines.

### 4. First run

1. **Download a backup copy of the workbook from SharePoint.**
2. Open the page, sign in, go to **Settings → Prepare workbook**. Read the report. New columns are added at the right-hand end
   of each table, and you can drag them next to the faculty columns in Excel if you like (the app works by header name).
3. Check rows the report flags. Use **Sessions → Needs review**, and the **Request** stage view's "Needs attention" for
   suggested colleges.

### 5. Phase 2: faculty request form

1. Power Automate → new **Instant cloud flow** → trigger **When an HTTP request is received** (method POST). Use a request
   body JSON schema with the field names from `request-form.html` (`Faculty Name`, `Faculty Email`, `Course #`, …).
2. Action **Excel Online (Business) → Add a row into a table**: the Instruction folder, `Instruction Requests.xlsx`,
   table `tbl_Requests`, mapping each column from the trigger body. Optionally add **Send an email** to the instruction team.
3. Save, copy the HTTP POST URL into `requestFlowUrl` in `config.js`, and link or embed `request-form.html` on the
   Library Instruction LibGuide.

Microsoft Forms + the same "Add a row" action is a no-code alternative (spec §9). The inbox doesn't care which one writes the rows.

## Testing locally

```powershell
powershell -ExecutionPolicy Bypass -File serve.ps1
```

Then open http://localhost:8080/ (demo mode) and http://localhost:8080/tests.html.

## Closing terms and starting a new year

Term dates and each term's Closed flag live in the workbook's hidden **`Lists`** sheet (table `tbl_Terms`), so every
computer sees the same thing. Prepare workbook creates it, using the dates in `config.js` as the starting values.
The same sheet holds `tbl_AppLock`, a one-row advisory lock (spec §5.6).

- **Close term** (Sessions → the term's tab → *Close term…*): lists open items (pending rooms, no librarian, missing
  attendance) and lets you close anyway. A closed term shows 🔒, is read-only on every computer, stays in the workbook,
  and still counts on the dashboard. *Reopen term* undoes it.
- **Start new academic year** (Settings → *Start new academic year…*): one page with four steps.
  1. Shows which terms are still open and closes them as part of the run.
  2. Proposes next year's sheets (`FALL 26` → `FALL 27`, …) with this year's dates moved forward one year. Edit them to the calendar.
  3. Lists sessions already booked in the new year's date range and moves the checked ones across, keeping their record id.
  4. Runs it:
     - **copies** this workbook in SharePoint (so formatting, dropdowns and the Lists sheet carry over), empties and renames the sheets, and writes the new term dates;
     - moves the chosen sessions;
     - marks the old workbook *Archived* and moves it into `Archive/` (created if missing).

  `Instruction Requests.xlsx` isn't touched.
- **Other computers** re-check for the lock before every save. While a transition runs they show "Editing is paused"
  and refuse writes. When the old workbook is archived they switch to the new year by themselves. A lock older than
  10 minutes is treated as abandoned.
- If the run stops part-way, the log says what happened. Before archiving, the old workbook is unlocked and stays active.
  A half-made new workbook must be deleted or renamed before trying again (the wizard refuses to overwrite it).
- Past years stay on the dashboard's year selector, read-only.

## Known gaps

Left as it is on purpose:
- The LibInsight column on the `oriet tours` sheet is left visible. Spec Q11 only covers term sheets.
