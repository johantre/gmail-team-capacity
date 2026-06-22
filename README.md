# gmail-team-capacity

[![Deploy to Apps Script](https://github.com/johantre/gmail-team-capacity/actions/workflows/deploy.yml/badge.svg?branch=main)](https://github.com/johantre/gmail-team-capacity/actions/workflows/deploy.yml)

Visualize team capacity for the upcoming 4 weeks directly in Google Sheets, based on **Out of Office** events in Google Calendar.

No external tools, no servers — just Google Apps Script connected to the Google Calendar API.

## What it does

- Reads **Out of Office** events (`eventType: outOfOffice`) from each team member's Google Calendar
- Supports full-day and partial absences (half day, quarter day, ...)
- Shows available days per team member per week (e.g. `3.5 / 5 days`)
- Displays a total availability percentage per team per week
- Color coding: 🟢 fully available · 🟡 partially available · 🔴 mostly absent
- Teams and members are configured via a dedicated **Teams** tab in the Sheet — no code changes needed

## Architecture

```
gmail-team-capacity/
├── .github/
│   └── workflows/
│       └── deploy.yml         # GitHub Actions: automatic deploy via clasp
├── Code.js                    # Google Apps Script — main logic
├── appsscript.json            # Apps Script manifest
├── .clasp.json                # Link to the Apps Script project
└── .gitignore
```

## Teams tab

Configure teams and members in the **"Teams"** tab of your Google Sheet:

| Team   | Name      | Email                  |
|--------|-----------|------------------------|
| Team A | Member 1  | member1@yourcompany.be |
| Team A | Member 2  | member2@yourcompany.be |
| Team B | Member 3  | member3@yourcompany.be |
| Team B | Member 4  | member4@yourcompany.be |

Simply add rows for Team C, Team D, etc. — the script picks them up automatically.

## Requirements

- Google Workspace account with access to your colleagues' Google Calendars
- Google Apps Script project linked to a Google Sheet
- **Advanced Google Services** enabled in Apps Script: `Google Calendar API`

## Setup

### 1. Prepare your Google Sheet

Create a Google Sheet with two tabs:
- **Teams** — with columns `Team | Name | Email`
- **Capaciteit** — created and filled automatically by the script

### 2. Link Apps Script

Go to **Extensions → Apps Script** in your Sheet, paste the contents of `Code.js` and enable the **Google Calendar API** via **Services → Google Calendar API**.

### 3. Link locally via clasp

```bash
npm install -g @google/clasp
clasp login
clasp clone <script-id>
```

Find your Script ID via **Apps Script → Project settings**.

### 4. Set up GitHub Actions

Add your `.clasprc.json` as a GitHub Secret:

**Settings → Secrets and variables → Actions → New repository secret**

| Secret name    | Value                               |
|----------------|-------------------------------------|
| `CLASPRC_JSON` | Contents of `~/.clasprc.json`       |

Every push to `main` automatically deploys to Apps Script via clasp.

## Usage

Click the **🔄 Refresh capaciteit** button in your Google Sheet (or via the **📅 Capaciteit** menu). The overview updates with the latest calendar data.

## CI/CD

```
git push origin main
       │
       ▼
GitHub Actions
       │  clasp push --force
       ▼
Google Apps Script
       │
       ▼
Google Sheet
```

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0). This means you can:

- **Share** — copy and redistribute the material in any medium or format
- **Adapt** — remix, transform, and build upon the material

Under the following terms:

- **Attribution** — you must give appropriate credit, provide a link to the license, and indicate if changes were made
- **NonCommercial** — you may not use the material for commercial purposes
- **No additional restrictions** — you may not apply legal terms or technological measures that legally restrict others from doing anything the license permits

See the [LICENSE](LICENSE) file for details.

---

**Made with ❤️ for efficiency and automation**