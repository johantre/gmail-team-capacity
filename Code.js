// ============================================================
// TEAM CAPACITY - Google Apps Script v3
// ============================================================
// Configuration via the "Teams" tab in your Sheet:
//
//   | Team   | Name      | Email                |   ...   | Sprint naming signature    |
//   |--------|-----------|----------------------|---------|---------------------------|
//   | Team A | Pier      | pier@acme.be       |         | Sprint 14 (16 jun – 30 jun)|
//   | Team A | Pol       | pol@acme.be        |         | Sprint 15 (30 jun – 14 jul)|
//
// Sprint naming signature is in column H (index 7).
// Format: Sprint {nr} ({day} {month} – {day} {month})
// Convention: first day of sprint counts, last (overlapping) day does NOT count.
// Sprints span 3 calendar weeks (e.g. 4d / 5d / 1d). No half-day adjustments.
//
// Each sprint has a "🦸 Hero" checkbox column (rotating incident-duty role,
// 0 capacity that sprint) right before its own week columns. It drives the
// member's "x / y" cells via a live formula, so checking it instantly zeroes
// that member's capacity for the sprint and recalculates → Projected SP —
// see leesOpgeslagenHeroes() (Heroes.js) for how checked state survives a refresh.
//
// This file holds the entry points and the main doRefresh() orchestration:
// it gathers the data and then hands the row-by-row sheet drawing off to
// SheetRenderer.js. Everything it calls out to lives in its own file:
//   Config.js          — constants (sheet names, colors, month map, ...)
//   Teams.js            — parsing the "Teams" tab into teams/emails/ids
//   SprintConfig.js     — sprint parsing, sprint groups, week/absence date math
//   Heroes.js           — Hero checkbox persistence across a refresh
//   CalendarService.js  — Calendar API access
//   SheetRenderer.js     — drawing the "Capacity" sheet itself
//   Utils.js            — pure helpers (column letters, date formatting)
// ============================================================

// ============================================================
// MAIN FUNCTION
// refreshCapaciteit() shows the loading dialog; the dialog calls doRefresh()
// ============================================================

// Each team member can install their own "on open" trigger for
// refreshCapaciteit (see installeerTrigger()) — deliberately per-user, so the
// auto-refresh keeps working under any remaining member's own credentials
// even if others leave the org. The downside: if several people currently
// have one installed, opening the sheet fires all of them at once. Since
// nobody but the trigger's own owner can remove it (Google restriction), we
// guard here instead with a cross-user cooldown, so whichever trigger fires
// first "wins" and the rest no-op within the cooldown window.
//
// LockService (not just CacheService) is required here: get-then-put on the
// cache is two separate steps, not atomic, so when several triggers fire
// within milliseconds of each other they can all read "not refreshing yet"
// before the first one writes its flag — LockService.getScriptLock() gives
// real mutual exclusion across concurrent executions of this script so only
// one of them makes that check-and-set at a time.
const AUTO_REFRESH_COOLDOWN_SEC = 60;

function refreshCapaciteit() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let moetVerversen;
  try {
    const cache = CacheService.getScriptCache();
    moetVerversen = !cache.get('refreshLopend');
    if (moetVerversen) cache.put('refreshLopend', 'true', AUTO_REFRESH_COOLDOWN_SEC);
  } finally {
    lock.releaseLock();
  }

  if (!moetVerversen) {
    SpreadsheetApp.getActiveSpreadsheet().toast('⏳ Net al ververst, momentje...', '🗓️ Team Capacity', 4);
    return;
  }

  // Touch the Calendar service here, directly from the menu click, so that on
  // first use Apps Script can show the real Google authorization prompt.
  // google.script.run calls made from inside the dialog below cannot trigger
  // that prompt — they just fail with a generic "permission needed" error.
  Calendar.CalendarList.list({ maxResults: 1 });

  const html = HtmlService.createHtmlOutputFromFile('Loading')
    .setWidth(480).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function doRefresh(vanSidebar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const teamsSheet = ss.getSheetByName(SHEET_TEAMS);
  if (!teamsSheet) { SpreadsheetApp.getUi().alert('❌ Tab "Teams" not found.'); return; }

  const { teams, teamIds, teamNamen } = verzamelTeams(teamsSheet);
  if (teamNamen.length === 0) { SpreadsheetApp.getUi().alert('❌ No teams found.'); return; }

  const sprints = leesSprintConfig(teamsSheet);
  const vandaag = new Date();
  const { sprintGroepen, weekKolommen, aantalKolommen } = bouwSprintGroepen(sprints, vandaag);

  // ONE API call per person for the full period
  const alleEmails = [...new Set(teamNamen.flatMap(t => teams[t].map(l => l.email)))];
  const periodeStart = weekKolommen[0].week.start;
  const periodeEinde = new Date(weekKolommen[weekKolommen.length - 1].week.einde);
  periodeEinde.setHours(23, 59, 59, 999);

  const afwezigCache = {};
  const kalenderFouten = {};
  alleEmails.forEach(email => {
    const { dagData, fout } = haalAfwezighedenOp(email, periodeStart, periodeEinde);
    afwezigCache[email] = dagData;
    if (fout) kalenderFouten[email] = fout;
  });

  // Read velocities and Hero checkboxes before clearing so user-entered/
  // -checked state survives the refresh.
  const { savedVelocities, savedHeroes } = leesBestaandeSheetState(ss, teamNamen);
  const sheet = maakSheetLeeg(ss, aantalKolommen, sprintGroepen, vandaag);
  const refreshData = { teamIds, savedVelocities, savedHeroes, afwezigCache, kalenderFouten };

  // Day-cell ranges across all teams — colored via conditional formatting
  // (applied once, after the loop below) instead of a static background, so
  // the color stays correct live when a Hero checkbox is toggled.
  const dagCelRanges = [];

  let rij = 2;
  teamNamen.forEach(teamNaam => {
    const leden = teams[teamNaam];
    const teamHeaderRij = rij;

    tekenTeamHeader(sheet, rij, teamNaam, aantalKolommen, refreshData);
    rij++;

    rij = tekenSprintEnWeekHeaders(sheet, rij, sprintGroepen, aantalKolommen);

    const eersteLedenRij = teamHeaderRij + 3;
    const laatsLedenRij = eersteLedenRij + leden.length - 1;
    tekenLedenRijen(sheet, eersteLedenRij, leden, teamNaam, sprintGroepen, aantalKolommen, refreshData, dagCelRanges);
    rij = laatsLedenRij + 1;

    rij = tekenSamenvattingRijen(sheet, rij, teamHeaderRij, eersteLedenRij, laatsLedenRij, leden, teamNaam, sprintGroepen, aantalKolommen, refreshData);

    // Empty row
    sheet.setRowHeight(rij, 16);
    rij++;
  });

  pasConditionalFormattingToe(sheet, dagCelRanges);
  tekenFoutenblok(ss, sheet, rij, kalenderFouten, aantalKolommen, vanSidebar);
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📅 Capacity")
    .addItem("🔄 Refresh capacity", "refreshCapaciteit")
    .addItem("⚙️ Install auto-refresh trigger", "installeerTrigger")
    .addToUi();
}

// Run once via the menu to install the installable on-open trigger.
// An installable trigger CAN call the Calendar API; simple onOpen() cannot.
function installeerTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshCapaciteit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshCapaciteit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert('✅ Auto-refresh installed — capacity will load automatically on open.');
}
