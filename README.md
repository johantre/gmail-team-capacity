# gmail-team-capacity

Visualize team capacity for the upcoming 4 sprints directly in Google Sheets, based on **Out of Office** events in Google Calendar.

No external tools, no servers — just Google Apps Script connected to the Google Calendar API.

## What it does

- Reads **Out of Office** events (`eventType: outOfOffice`) from each team member's Google Calendar
- Supports full-day and partial absences (half day, quarter day, ...)
- Shows available days per team member per week (e.g. `3.5 / 5 days`)
- **🦸 Hero checkbox** per sprint per member — for the rotating incident-duty role. Checking it instantly zeroes that member's capacity for that sprint's weeks (`0 / 4`, `0 / 5`, `0 / 1`) and recalculates Projected SP, no refresh needed
- Displays total availability (MD) and percentage per sprint
- Calculates **Projected SP** per sprint based on an editable Average Velocity per team
- Color coding: 🟢 fully available · 🟡 partially available · 🔴 mostly absent
- Teams and members are configured via a dedicated **Teams** tab — no code changes needed

## Architecture

```
gmail-team-capacity/
├── Code.js                    # Google Apps Script — main logic
├── Loading.html               # Pulsing Acme logo dialog shown during refresh
├── appsscript.json            # Apps Script manifest (scopes, timezone)
├── .clasp.json                # Link to the Apps Script project
├── res/
│   ├── calendar-event.svg     # Tabler icon (reference — used inline in code)
│   └── chart-bar.svg          # Tabler icon (reference — used inline in code)
└── .gitignore
```

## Teams tab

The **Teams** sheet drives everything. It has two distinct sections:

### Section A–C: Team members

| Team    | Name    | Email                  |
|---------|---------|------------------------|
| Team A  | Pier    | pier@acme.be         |
| Team A  | Pol     | pol@acme.be          |
| Team B  | Marie   | marie@acme.be        |

- **Team**: must match exactly with the team names used in the Capacity sheet
- **Name**: display name shown in the Capacity sheet
- **Email**: used to query Google Calendar — must be the member's Google Workspace email

> **Email formula**: emails follow the pattern `firstname@acme.be`. If your organisation uses a different pattern (e.g. `firstname.lastname@acme.be`), you can use a Sheets formula in the Email column:
> ```
> =LOWER(A2)&"@acme.be"
> ```
> The script detects any cell containing `@` as a valid email address.

Add or remove rows freely — the script picks them up automatically on the next refresh. New teams are assigned a color automatically.

### Section E–F: Team IDs (Jira velocity links)

| Team    | id   |
|---------|------|
| Team A  | 42   |
| Team B  | 57   |

- **Team**: exact team name (must match column A)
- **id**: numeric Jira board ID

When a team has an ID, a clickable 📊 icon appears next to the team name in the Capacity sheet, linking directly to:
```
https://eforge.atlassian.net/jira/software/c/projects/ACME/boards/{id}/reports/velocity
```
Teams without an ID simply show no icon.

### Section H: Sprint configuration

At least **2 sprint rows** are required. The script derives the sprint duration from those 2 examples and auto-generates all future sprints.

| Column H                        |
|---------------------------------|
| Sprint 14 (16 jun – 30 jun)     |
| Sprint 15 (30 jun – 14 jul)     |

**Format**: `Sprint {nr} ({day} {month} – {day} {month})`

**Convention**: the first day of a sprint counts, the last (overlapping) day does **not** count — `[start, end)`.

Supported month abbreviations: `jan feb mar apr may jun jul aug sep oct nov dec` (also `mrt mei okt` for Dutch).

The script shows the **current sprint + 3 upcoming sprints** (4 total), each spanning 3 calendar weeks.

## Average Velocity

Each team header row in the Capacity sheet has an editable **Average Velocity** field (in Story Points). Fill this in manually per team. The script uses it to calculate **Projected SP** per sprint:

```
Projected SP = Velocity × (Current MD / Full capacity MD)
```

This value updates live as you change the velocity — no refresh needed. The velocity is preserved across refreshes.

## Hero checkbox

Each sprint block has a **🦸 Hero** checkbox column, one row per team member, right before that sprint's week columns. Check it when that person is the rotating incident-duty ("Hero") role for that sprint — this:

- Sets their `x / y` cells for that sprint's weeks to `0 / y` (they contribute 0 capacity that sprint)
- Recalculates **Current MD** and **→ Projected SP** for that sprint immediately, since both are formulas reading the member cells
- Is per (member, sprint) — a Hero-checked sprint doesn't affect that person's other sprints

Checked state is preserved across refreshes, matched by member name + sprint number (not row/column position, which shifts as sprints roll forward).

The member row's background color (🟢/🟡/🔴) is conditional formatting driven by the cell's own "x / y" text, so it also updates live when Hero is toggled — no refresh needed.

**Known limitation**: the "MD / Sprint (%)" summary line is still set once at refresh time, not live — it's a merged label, not a per-cell value, so it can't be driven by the same per-cell conditional formatting/formula approach without a bigger rework. Checking Hero after a refresh updates the member cells, colors and → Projected SP instantly, but that one summary line only catches up on the next **🔄 Refresh capacity**.

## Requirements

- Google Workspace account with access to your colleagues' Google Calendars
- Google Apps Script project linked to a Google Sheet
- **Advanced Google Services** enabled in Apps Script: `Google Calendar API`

## Column positions in the Teams tab

| Section         | Columns       | Flexible? |
|-----------------|---------------|-----------|
| Team members    | A, B, C       | ❌ Fixed  |
| Team IDs        | any two adjacent columns with `id` as header | ✅ Flexible |
| Sprint config   | H             | ❌ Fixed  |

The team ID section is detected dynamically: the script scans the first 3 rows for a cell containing exactly `id` and treats the column to its left as the team name. It can be placed anywhere as long as the header row has `id` in it.

## Setup

### 1. Create a new Google Sheet

Create a Google Sheet with two tabs:
- **Teams** — structured as described above (columns A–C for members, H for sprints, optional ID section)
- **Capacity** — leave this empty; the script creates and fills it automatically

### 2. Link Apps Script to the Sheet

1. Open your Sheet → **Extensions → Apps Script**
2. Delete any existing code and paste the contents of `Code.js`
3. Enable the Calendar API: **Services → Google Calendar API → Add**
4. Save and close

### 3. Get the Script ID

In Apps Script: **Project settings (⚙️) → Script ID** — copy this value.

### 4. Update `.clasp.json`

Update the `scriptId` in `.clasp.json` with the value from step 3:

```json
{
  "scriptId": "your-new-script-id-here",
  "rootDir": ""
}
```

Push to Bitbucket — the pipeline deploys the code to your new Apps Script project automatically.

> If you're working with multiple sheets, each sheet has its own Apps Script project and its own `scriptId`. The code itself is sheet-agnostic — it always reads from the sheet it is attached to.

### 5. Auto-refresh on sheet open

Run **📅 Capacity → ⚙️ Install auto-refresh trigger** once. This installs a per-user installable trigger that shows the refresh dialog automatically when the sheet is opened.

### 5. Set up CI/CD

#### Bitbucket (primary)

Add the base64-encoded contents of `~/.clasprc.json` as a repository variable in Bitbucket:

**Repository settings → Repository variables**

| Variable name                  | Value                                        | Secured |
|-------------------------------|----------------------------------------------|---------|
| `CLASPRC_JSON_TEAM_CAPACITY`  | base64-encoded contents of `~/.clasprc.json` | ✅      |

Generate the base64 value:
```bash
# Linux/Mac
base64 -i ~/.clasprc.json

# Windows PowerShell
[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-Content "$env:USERPROFILE\.clasprc.json" -Raw)))
```

Every push to `main` that touches `gmail-team-capacity/**` automatically deploys to Apps Script.

#### GitHub Actions (secondary)

Add the raw contents of `~/.clasprc.json` as a GitHub secret:

**Settings → Secrets and variables → Actions → New repository secret**

| Secret name    | Value                               |
|----------------|-------------------------------------|
| `CLASPRC_JSON` | Contents of `~/.clasprc.json`       |

## Usage

Open your Google Sheet and use the **📅 Capacity** menu:
- **🔄 Refresh capacity** — manually trigger a refresh
- **⚙️ Install auto-refresh trigger** — install the on-open trigger (once per user)

### First-time authorization (once per colleague)

Every team member needs their own Google authorization for this script, since it reads their colleagues' calendars. The **first** time someone clicks **🔄 Refresh capacity**, Google shows a standard authorization prompt (choose account → *Advanced* → *Go to [project] (unsafe)*, since the script isn't published/verified — this is expected for an internal script). Approve it once; every refresh after that works normally.

If a colleague instead sees a generic "permission needed" error inside the loading dialog, it means they skipped that first authorization (e.g. the sheet was shared with the auto-refresh trigger already installed, so the dialog opened before they ever got the prompt). Have them run **🔄 Refresh capacity** manually once to trigger the proper authorization screen.

## CI/CD flow

```
git push origin main
       │
       ▼
Bitbucket Pipelines (on changes in gmail-team-capacity/**)
       │  clasp push --force
       ▼
Google Apps Script
       │
       ▼
Google Sheet
```

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0). This means you can:

- Share — copy and redistribute the material in any medium or format
- Adapt — remix, transform, and build upon the material

Under the following terms:

- **Attribution** — you must give appropriate credit, provide a link to the license, and indicate if changes were made
- **NonCommercial** — you may not use the material for commercial purposes

See the [LICENSE](LICENSE) file for details.

---

**Made with ❤️ for efficiency and automation**
