# E-Co-Op Program Tracker

Internal ops portal for the Entrepreneurship Cohort Program (E-Co-Op), C-SHINE, KLE Technological University. Replaces the scattered Excel/PPT files with one dashboard: team directory, task tracker, mentor directory, and program resources.

**Data model:** a Google Sheet is the single source of truth (editable, with dropdowns, color coding, locked headers, and Google's built-in version history as auto-recover). This site is a static, read-only frontend that reads that sheet live. No backend, no build step — just HTML/CSS/JS, deployable to GitHub Pages.

## Folder structure

```
portal/
  index.html       Overview / dashboard
  teams.html        Team directory (search + filter)
  tasks.html        Task tracker (search + filter)
  mentors.html      Faculty + industry mentor directory
  resources.html    Training decks / program PDF (downloadable)
  css/style.css
  js/config.js      <- put your Google Sheet ID here
  js/data.js        Fetches the sheet, falls back to data/seed.json
  js/app.js         Shared UI helpers
  data/seed.json    Snapshot of the original Excel data (fallback only)
  assets/resources/ Copies of the original PPT/PDF files
  apps-script/Setup.gs   One-time script to build the Google Sheet
```

## 1. Set up the Google Sheet

1. Go to [sheets.new](https://sheets.new) to create a blank spreadsheet.
2. **Extensions > Apps Script**. Delete the placeholder code.
3. Paste in the entire contents of [`apps-script/Setup.gs`](apps-script/Setup.gs).
4. Click **Run > setupWorkbook** (top toolbar). The first run will ask you to authorize — this is your own script running in your own sheet, so it's safe to approve.
5. Switch back to the spreadsheet tab. You'll see 7 tabs, fully formatted and seeded with the real team/task/mentor data: **Overview, Teams, Tasks, Evaluation, Surveys, Faculty Mentors, Industry Mentors**.
   - Header rows are locked.
   - Status/dropdown columns have data validation (no free typing garbage values).
   - Status values are color-coded automatically (conditional formatting).
   - **Teams** has 6 structured student slots (Name / SRN / Phone each), grouped by field type so formulas can reference simple ranges, plus a **Faculty Mentor Email** and **Mentor Notes** column (used by the admin app's mentor self-service view and email reminders — see `admin-app/README.md`). A live **Team Size** formula counts filled slots, and a live **Duplicate SRN?** formula flags any SRN repeated within the same team *or* appearing on a different team — highlighted in red. See "A real duplicate this caught" below.
   - **Tasks** due dates are stored as real dates (DD-MM-YYYY) wherever the original value was unambiguous, with live **Days Until Due** and **Overdue?** formula columns (overdue = past due and not marked Complete).
   - **Evaluation** and **Surveys** start empty (no data existed for these in the original files) — populate them through the admin app. Evaluation tracks the real Stage 1 (5-criterion, 70/100 threshold) and Stage 2 (C-SHINE panel) scores per team; Surveys tracks Customer Insight Survey interview progress and faculty review status per team.
   - **Overview** rolls up intelligence: total teams, mentors, tasks complete, overdue task count, teams missing a faculty mentor, duplicate SRNs found, average team size, teams that passed Stage 1, teams admitted at Stage 2, and surveys reviewed — all formulas, always current.
   - `Faculty Mentors` is derived from `Teams` and is warning-protected (it's meant to stay auto-generated — edit mentor info via `Teams` instead).
   - Formula columns (Team Size, Duplicate SRN?, Days Until Due, Overdue?, Stage 1 Total/Result, Survey Progress) are applied per-row — to the rows with real seeded data now, and automatically to any new row added through the admin app. They're deliberately *not* bulk pre-filled across empty rows, since that would make Sheets think those rows are "in use" and break Add for everything after them.
   - Google Sheets' built-in version history (File > Version history) is your auto-recover / undo safety net — nothing extra to configure.
   - This portal (the public read-only site) currently displays Teams/Tasks/Faculty Mentors/Industry Mentors only — Evaluation and Surveys are visible in the admin app, not yet on this public dashboard. Ask if you want that added.
6. **File > Share > General access > Anyone with the link > Viewer.** This is required so the static site can read the data (it only ever reads, never writes). Add real editors (yourself, co-coordinators) by email with Editor access separately — that's unaffected by the link-sharing setting.

⚠️ Re-running `setupWorkbook()` **rebuilds every tab from scratch** — only do this before you've started editing real data by hand, not after.

### A real duplicate this caught

The Duplicate SRN? formula (and the portal's own duplicate check) flagged that **Aditya Naik** (team "AI-Enabled eCommerce for Retailers") and **Sushant Maheshwari** (team "Integrated academic performance monitoring system") share the exact same SRN, `01FE23BCS211`, in the original `Final list of E-Co-Op teams.xlsx`. That's almost certainly a typo on one of the two rows — worth fixing in the sheet once it's set up.

### Two dates worth double-checking

The original `Tasks E-Co-Op.xlsx` had two due dates stored as ambiguous Excel date values (`2026-01-07` / `2026-04-08`, day/month swapped depending on locale). Based on the surrounding timeline they were corrected to **01-07-2026** and **04-08-2026**, and both rows are flagged in the sheet's Remarks column and highlighted in the Due Date cell. Please confirm the real intended dates once you're in the sheet.

## 2. Point the site at your sheet

Open `js/config.js` and paste your sheet's ID (the long string in its URL between `/d/` and `/edit`):

```js
const CONFIG = {
  SHEET_ID: "your-sheet-id-here",
};
```

That's it — every page fetches live from `Teams`, `Tasks`, `Faculty Mentors`, and `Industry Mentors` via Google's public `gviz` JSON endpoint (no API key needed, since the sheet is link-shared as Viewer). If `SHEET_ID` is left blank, or the fetch fails for any reason, the site automatically falls back to the bundled `data/seed.json` snapshot and shows a small "showing bundled snapshot" indicator so it's obvious which one you're looking at.

## 3. Deploy to GitHub Pages

From inside this `portal/` folder:

```bash
git init
git add .
git commit -m "Initial E-Co-Op portal"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

Then on GitHub: **Settings > Pages > Source: Deploy from branch > main / (root)**. Your site will be live at `https://<your-username>.github.io/<your-repo>/` within a minute or two.

## Notes

- The **Student Orientation** deck (`assets/resources/Student Orinetation.pptx`) is currently just a title slide in the source file — replace it once the real deck is ready.
- The portal is read-only by design. All real edits (adding teams, updating task status, adding mentors) happen in the Google Sheet, and show up on the site on next page load/refresh.
