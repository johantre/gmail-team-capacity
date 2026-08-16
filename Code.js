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
// ============================================================

const WERKUREN_PER_DAG = 8;

// Inline SVG icons (white, Tabler style) — no internet needed, sourced from res/
const SVG_CALENDAR = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 5m0 2a2 2 0 0 1 2 -2h12a2 2 0 0 1 2 2v12a2 2 0 0 1 -2 2h-12a2 2 0 0 1 -2 -2z"/><path d="M16 3l0 4"/><path d="M8 3l0 4"/><path d="M4 11l16 0"/><path d="M8 15h2v2h-2z"/></svg>';
const SVG_CHART    = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 16v-5"/><path d="M11 16v-8"/><path d="M15 16v-3"/></svg>';
const SHEET_TEAMS = "Teams";
const SHEET_CAPACITEIT = "Capacity";

// Acme brand colors
const BLAUW      = "#3E7CBF";
const TURQUOISE  = "#27B4AF";
const GRIJS      = "#3C3C3B";
const BLAUW_LICHT     = "#d6e8f5";
const TURQUOISE_LICHT = "#d0edec";

const TEAM_KLEUREN = {
  "Team A": BLAUW,
  "Team B": TURQUOISE,
  "Team C": "#2e6da4", // darker blue
  "Team D": "#1e9490", // darker turquoise
};

const MAAND_MAP = {
  jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
  jul:6, aug:7, sep:8, oct:9, nov:10, dec:11,
  mrt:2, mei:4, okt:9,
};

// ============================================================
// SPRINT CONFIG — read from column H, auto-generate future sprints
// ============================================================
function leesSprintConfig(teamsSheet) {
  const data = teamsSheet.getDataRange().getValues();
  const sprints = [];
  for (let i = 1; i < data.length; i++) {
    const val = (data[i][7] || "").toString().trim();
    if (!val) continue;
    const match = val.match(/Sprint\s+(\d+)\s*\((\d+)\s+(\w+)\s*[–\-]\s*(\d+)\s+(\w+)\)/i);
    if (!match) continue;
    const [, nr, startDag, startMaand, eindeDag, eindeMaand] = match;
    const jaar = new Date().getFullYear();
    const sm = MAAND_MAP[startMaand.toLowerCase()];
    const em = MAAND_MAP[eindeMaand.toLowerCase()];
    if (sm === undefined || em === undefined) continue;
    const start = new Date(jaar, sm, parseInt(startDag)); start.setHours(0,0,0,0);
    const eindeJaar = em < sm ? jaar + 1 : jaar;
    const einde = new Date(eindeJaar, em, parseInt(eindeDag)); einde.setHours(0,0,0,0);
    sprints.push({ nr: parseInt(nr), naam: `Sprint ${nr}`, start, einde });
  }

  // Derive duration from the 2 examples and auto-generate future sprints
  if (sprints.length >= 2) {
    sprints.sort((a, b) => a.nr - b.nr);
    const duurMs = sprints[1].start - sprints[0].start;
    const vandaag = new Date();
    let laatste = sprints[sprints.length - 1];
    // Generate until we have at least 4 sprints beyond today
    while (laatste.einde <= vandaag || sprints.filter(s => s.start > vandaag).length < 4) {
      const start = new Date(laatste.einde);
      const einde = new Date(start.getTime() + duurMs);
      const nr = laatste.nr + 1;
      laatste = { nr, naam: `Sprint ${nr}`, start, einde };
      sprints.push(laatste);
      if (sprints.length > 50) break;
    }
  }

  return sprints;
}

// Build the 3 calendar weeks a sprint spans, with actual working days per week.
// Convention: [sprint.start, sprint.einde) — first day in, last day out.
function buildSprintWeken(sprint) {
  const weken = [];
  const cursor = getMaandagVanWeek(sprint.start);
  while (cursor < sprint.einde) {
    const weekStart = new Date(cursor);
    const weekEinde = new Date(cursor); weekEinde.setDate(cursor.getDate() + 4);
    // Count working days (Mon–Fri) within [sprint.start, sprint.einde)
    let werkdagen = 0;
    const dc = new Date(weekStart);
    while (dc <= weekEinde) {
      const dag = dc.getDay();
      if (dag !== 0 && dag !== 6 && dc >= sprint.start && dc < sprint.einde) werkdagen++;
      dc.setDate(dc.getDate() + 1);
    }
    if (werkdagen > 0) {
      weken.push({ week: { start: weekStart, einde: weekEinde }, sprint, werkdagen });
    }
    cursor.setDate(cursor.getDate() + 7);
  }
  return weken;
}

// Count absent days from cache, only for days within [sprintStart, sprintEinde)
function berekenAfwezigInWeek(dagCache, weekStart, weekEinde, sprintStart, sprintEinde) {
  let totaalUren = 0;
  const cursor = new Date(weekStart);
  while (cursor <= weekEinde) {
    if (cursor >= sprintStart && cursor < sprintEinde) {
      totaalUren += dagCache[cursor.toDateString()] || 0;
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return Math.round((totaalUren / WERKUREN_PER_DAG) * 4) / 4;
}

// ============================================================
// MAIN FUNCTION
// refreshCapaciteit() shows the loading dialog; the dialog calls doRefresh()
// ============================================================
function refreshCapaciteit() {
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

  const data = teamsSheet.getDataRange().getValues();
  const teams = {};
  const teamIds = {};
  // Find "id" column dynamically (it's in a separate E-F section, not col D)
  let idKolIdx = -1, idTeamKolIdx = -1;
  for (let r = 0; r < Math.min(3, data.length); r++) {
    for (let c = 0; c < data[r].length; c++) {
      if ((data[r][c] || '').toString().trim().toLowerCase() === 'id') {
        idKolIdx = c; idTeamKolIdx = c - 1; break;
      }
    }
    if (idKolIdx >= 0) break;
  }
  data.forEach(rij => {
    const [team, naam, email] = rij.map(v => v.toString().trim());
    if (email && email.includes("@")) {
      if (!teams[team]) teams[team] = [];
      teams[team].push({ naam, email });
    }
    if (idKolIdx >= 0) {
      const idTeam = (rij[idTeamKolIdx] || '').toString().trim();
      const id = rij[idKolIdx];
      if (idTeam && typeof id === 'number' && id > 0) teamIds[idTeam] = id.toString();
    }
  });

  const teamNamen = Object.keys(teams);
  if (teamNamen.length === 0) { SpreadsheetApp.getUi().alert('❌ No teams found.'); return; }

  const sprints = leesSprintConfig(teamsSheet);
  const vandaag = new Date();

  // Current sprint = today within [start, einde), fallback to next upcoming
  const huidigeSprint = sprints.find(s => vandaag >= s.start && vandaag < s.einde)
    || sprints.slice().sort((a, b) => a.start - b.start).find(s => s.start > vandaag);

  // Build sprint groups: 4 sprints starting from current
  const sprintGroepen = [];
  if (huidigeSprint) {
    for (let i = 0; i < 4; i++) {
      const sprint = sprints.find(s => s.nr === huidigeSprint.nr + i);
      if (sprint) sprintGroepen.push({ sprint, weken: buildSprintWeken(sprint) });
    }
  }

  // Flatten to ordered week columns and set startKolIdx per group
  const weekKolommen = sprintGroepen.flatMap(g => g.weken);
  let kolIdx = 0;
  sprintGroepen.forEach(groep => { groep.startKolIdx = kolIdx; kolIdx += groep.weken.length; });

  const aantalKolommen = 2 + weekKolommen.length;

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

  // Set up sheet — read velocities before clearing so user-entered values survive refresh
  let sheet = ss.getSheetByName(SHEET_CAPACITEIT);
  const savedVelocities = {};
  if (sheet) {
    const existingVals = sheet.getDataRange().getValues();
    existingVals.forEach(row => {
      const naam = (row[0] || '').toString().trim();
      if (teamNamen.includes(naam)) {
        for (let i = row.length - 1; i > 0; i--) {
          if (typeof row[i] === 'number' && row[i] > 0) { savedVelocities[naam] = row[i]; break; }
        }
      }
    });
  }
  if (!sheet) sheet = ss.insertSheet(SHEET_CAPACITEIT);
  sheet.getImages().forEach(img => img.remove());
  sheet.clearContents();
  sheet.clearFormats();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 36);
  for (let k = 3; k <= aantalKolommen; k++) sheet.setColumnWidth(k, 120);

  // Main title
  sheet.getRange(1, 1, 1, aantalKolommen).merge();
  sheet.getRange(1, 1)
    .setValue("📅  Team Capacity — updated on " + formatDatum(vandaag))
    .setFontFamily("Montserrat").setFontSize(13).setFontWeight("bold")
    .setBackground(GRIJS).setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  const pendingImages = [];

  let rij = 2;

  teamNamen.forEach(teamNaam => {
    const leden = teams[teamNaam];
    const teamKleur = TEAM_KLEUREN[teamNaam] || "#5f6368";

    // Team title: [name col1] [📊 col2] [Max Capacity Velocity: col3..N-1] [velocity col N]
    const teamHeaderRij = rij;
    const velocity = savedVelocities[teamNaam] || 0;
    const teamId = teamIds[teamNaam] || '';
    const teamStyle = { font: "Montserrat", bg: teamKleur, fg: "#ffffff" };

    sheet.getRange(rij, 1).setValue(teamNaam)
      .setFontFamily(teamStyle.font).setFontSize(11).setFontWeight("bold")
      .setBackground(teamStyle.bg).setFontColor(teamStyle.fg)
      .setHorizontalAlignment("left");

    const teamLetter = teamNaam.replace(/^Team\s*/i, '').charAt(0).toUpperCase();
    const cel2 = sheet.getRange(rij, 2);
    cel2.setBackground(teamStyle.bg).setFontColor(teamStyle.fg)
      .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
    if (teamId) {
      const jiraUrl = `https://acme.atlassian.net/jira/software/c/projects/ACME/boards/${teamId}/reports/velocity`;
      cel2.setFormula(`=HYPERLINK("${jiraUrl}";"📊")`);
    } else {
      cel2.setValue("");
    }
    sheet.getRange(rij, 3).setBackground(teamStyle.bg);

    if (aantalKolommen > 4) {
      sheet.getRange(rij, 4, 1, aantalKolommen - 4).merge()
        .setValue("Max Capacity Velocity:")
        .setFontFamily(teamStyle.font).setFontSize(10).setFontWeight("normal")
        .setBackground(teamStyle.bg).setFontColor(teamStyle.fg)
        .setHorizontalAlignment("right");
    }
    sheet.getRange(rij, aantalKolommen)
      .setValue(velocity || null)
      .setNumberFormat('0" SP"')
      .setFontFamily(teamStyle.font).setFontSize(10).setFontWeight("bold")
      .setBackground(teamStyle.bg).setFontColor(teamStyle.fg)
      .setHorizontalAlignment("right");
    sheet.setRowHeight(rij, 28);
    rij++;

    // Sprint header row — merged over each sprint's columns
    sheet.getRange(rij, 1).setValue("").setBackground(BLAUW_LICHT);
    sheet.getRange(rij, 2).setBackground(BLAUW_LICHT);
    sprintGroepen.forEach(groep => {
      const kolStart = groep.startKolIdx + 3;
      const kolBreedte = groep.weken.length;
      const cel = sheet.getRange(rij, kolStart, 1, kolBreedte);
      if (kolBreedte > 1) cel.merge();
      const sprintLabel = `${groep.sprint.naam} (${formatDatumKort(groep.sprint.start)} – ${formatDatumKort(groep.sprint.einde)})`;
      cel.setValue(sprintLabel)
        .setFontFamily("Montserrat").setFontWeight("bold").setFontSize(10)
        .setBackground(BLAUW_LICHT).setFontColor(GRIJS).setHorizontalAlignment("center");
    });
    sheet.setRowHeight(rij, 22);
    rij++;

    // Week headers — date range + working days count
    const headers = ["Team member", ""];
    weekKolommen.forEach(wk => {
      headers.push(`${formatDatumKort(wk.week.start)} – ${formatDatumKort(wk.week.einde)}\n(${wk.werkdagen})`);
    });
    sheet.getRange(rij, 1, 1, aantalKolommen).setValues([headers])
      .setFontFamily("Montserrat").setFontWeight("normal")
      .setBackground(BLAUW_LICHT).setFontColor(GRIJS)
      .setHorizontalAlignment("center").setWrap(true);
    sheet.getRange(rij, 1).setFontWeight("bold");
    sheet.setRowHeight(rij, 44);
    rij++;

    // Team members
    leden.forEach(lid => {
      const dagCache = afwezigCache[lid.email] || {};
      const heeftFout = !!kalenderFouten[lid.email];
      const rowData = [lid.naam, ""];
      const beschikbaarData = [];

      weekKolommen.forEach(wk => {
        const afwezig = berekenAfwezigInWeek(dagCache, wk.week.start, wk.week.einde, wk.sprint.start, wk.sprint.einde);
        const beschikbaar = Math.max(0, Math.round((wk.werkdagen - afwezig) * 4) / 4);
        beschikbaarData.push({ beschikbaar, werkdagen: wk.werkdagen });
        if (heeftFout) {
          rowData.push("⚠️ no access");
        } else {
          const label = Number.isInteger(beschikbaar) ? beschikbaar : beschikbaar.toFixed(2).replace('.', ',');
          rowData.push(`${label} / ${wk.werkdagen}`);
        }
      });

      sheet.getRange(rij, 1, 1, aantalKolommen).setValues([rowData]).setFontFamily("Montserrat");
      sheet.getRange(rij, 3, 1, weekKolommen.length).setHorizontalAlignment("center");

      if (heeftFout) {
        sheet.getRange(rij, 1, 1, aantalKolommen).setBackground("#fff3e0").setFontColor("#e65100");
      } else {
        beschikbaarData.forEach(({ beschikbaar, werkdagen }, idx) => {
          const cel = sheet.getRange(rij, 3 + idx);
          const ratio = werkdagen > 0 ? beschikbaar / werkdagen : 1;
          if (ratio >= 1) cel.setBackground("#e6f4ea");
          else if (ratio >= 0.5) cel.setBackground("#fef9c3");
          else cel.setBackground("#fce8e6");
        });
      }
      rij++;
    });

    // Member row range for SUMPRODUCT formulas (team header + sprint header + week header = +3)
    const eersteLedenRij = teamHeaderRij + 3;
    const laatsLedenRij = eersteLedenRij + leden.length - 1;

    // MD / Sprint (%) row — merged per sprint
    const geldigeLeden = leden.filter(l => !kalenderFouten[l.email]).length;
    sheet.getRange(rij, 1).setValue("MD / Sprint (%)")
      .setFontFamily("Montserrat").setFontWeight("bold")
      .setBackground(TURQUOISE_LICHT).setFontColor(GRIJS);
    sheet.getRange(rij, 2).setBackground(TURQUOISE_LICHT);
    const sprintRanges = [];
    sprintGroepen.forEach(groep => {
      const kolStart = groep.startKolIdx + 3;
      const kolBreedte = groep.weken.length;
      const kolEinde = kolStart + kolBreedte - 1;
      const maxMandagen = geldigeLeden * groep.weken.reduce((s, wk) => s + wk.werkdagen, 0);
      let sprintMandagen = 0;
      groep.weken.forEach(wk => {
        leden.forEach(lid => {
          if (kalenderFouten[lid.email]) return;
          const dagCache = afwezigCache[lid.email] || {};
          const afwezig = berekenAfwezigInWeek(dagCache, wk.week.start, wk.week.einde, wk.sprint.start, wk.sprint.einde);
          sprintMandagen += Math.max(0, wk.werkdagen - afwezig);
        });
      });
      const pct = maxMandagen > 0 ? Math.round((sprintMandagen / maxMandagen) * 100) : 0;
      const mdLabel = Number.isInteger(sprintMandagen) ? sprintMandagen : sprintMandagen.toFixed(1).replace('.', ',');
      const cel = sheet.getRange(rij, kolStart, 1, kolBreedte);
      if (kolBreedte > 1) cel.merge();
      cel.setValue(`${mdLabel} MD (${pct}%)`)
        .setFontFamily("Montserrat").setFontWeight("bold")
        .setBackground(TURQUOISE_LICHT).setFontColor(GRIJS)
        .setHorizontalAlignment("center");
      sprintRanges.push({ kolStart, kolBreedte, kolEinde });
    });
    rij++;

    // Hidden "Current MD" row — SUMPRODUCT of x values from "x / y" member cells
    const rijCurrentMD = rij;
    sprintRanges.forEach(d => {
      const bereik = `${colLetter(d.kolStart)}${eersteLedenRij}:${colLetter(d.kolEinde)}${laatsLedenRij}`;
      const formula = `=SUMPRODUCT(IFERROR(VALUE(TRIM(LEFT(${bereik};FIND(" / ";${bereik})-1)));0))`;
      const cel = sheet.getRange(rij, d.kolStart, 1, d.kolBreedte);
      if (d.kolBreedte > 1) cel.merge();
      cel.setFormula(formula);
    });
    sheet.hideRows(rijCurrentMD, 1);
    rij++;

    // Hidden "Full Cap. MD" row — SUMPRODUCT of y values from "x / y" member cells
    const rijFullMD = rij;
    sprintRanges.forEach(d => {
      const bereik = `${colLetter(d.kolStart)}${eersteLedenRij}:${colLetter(d.kolEinde)}${laatsLedenRij}`;
      const formula = `=SUMPRODUCT(IFERROR(VALUE(TRIM(MID(${bereik};FIND(" / ";${bereik})+3;100)));0))`;
      const cel = sheet.getRange(rij, d.kolStart, 1, d.kolBreedte);
      if (d.kolBreedte > 1) cel.merge();
      cel.setFormula(formula);
    });
    sheet.hideRows(rijFullMD, 1);
    rij++;

    // Projected SP row — formula referencing velocity cell + hidden MD rows (semicolons for Belgian locale)
    const velocityRef = colLetter(aantalKolommen) + teamHeaderRij;
    sheet.getRange(rij, 1).setValue("→ Projected SP")
      .setFontFamily("Montserrat").setFontWeight("bold")
      .setBackground(TURQUOISE_LICHT).setFontColor(GRIJS);
    sheet.getRange(rij, 2).setBackground(TURQUOISE_LICHT);
    sprintRanges.forEach(d => {
      const spRef = colLetter(d.kolStart) + rijCurrentMD;
      const maxRef = colLetter(d.kolStart) + rijFullMD;
      const formula = `=IF(${velocityRef}="";"";IFERROR(ROUND(${velocityRef}*${spRef}/${maxRef};0);""))`;
      const cel = sheet.getRange(rij, d.kolStart, 1, d.kolBreedte);
      if (d.kolBreedte > 1) cel.merge();
      cel.setFormula(formula)
        .setNumberFormat('"→ "0" projected SP"')
        .setFontFamily("Montserrat").setFontWeight("bold")
        .setBackground(TURQUOISE_LICHT).setFontColor(GRIJS)
        .setHorizontalAlignment("center");
    });
    rij++;

    // Vertical sprint dividers — right border on last col of each sprint except the last
    const rijTeamEinde = rij - 1;
    sprintRanges.slice(0, -1).forEach(d => {
      sheet.getRange(teamHeaderRij, d.kolEinde, rijTeamEinde - teamHeaderRij + 1, 1)
        .setBorder(null, null, null, true, null, null, "#888888", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    });

    // Empty row
    sheet.setRowHeight(rij, 16);
    rij++;
  });


  // Errors at the bottom
  const foutenLijst = Object.entries(kalenderFouten);
  if (foutenLijst.length > 0) {
    rij++;
    sheet.getRange(rij, 1, 1, aantalKolommen).merge();
    sheet.getRange(rij, 1)
      .setValue("⚠️ Calendar access errors — check permissions for the accounts below")
      .setFontWeight("bold").setBackground("#fff3e0").setFontColor("#e65100");
    rij++;
    foutenLijst.forEach(([email, fout]) => {
      sheet.getRange(rij, 1, 1, aantalKolommen).merge();
      sheet.getRange(rij, 1).setValue(`${email}: ${fout}`)
        .setBackground("#fff8f0").setFontColor("#bf360c");
      rij++;
    });
    if (!vanSidebar) ss.toast('✅ Capacity updated!', '🗓️ Team Capacity', 5);
    if (foutenLijst.length > 0) {
      Logger.log(`Calendar access errors: ${foutenLijst.map(([e]) => e).join(', ')}`);
    }
  }
}

// ============================================================
// ONE API call per person — returns data + any error
// ============================================================
// TODO: all-day OOO events are overcounted by 1 day here. event.start.date /
// event.end.date are bare "YYYY-MM-DD" strings, which `new Date(...)` always
// parses as UTC midnight — in Brussels (always ahead of UTC) that lands just
// after local midnight, so the day-walking loop below picks up one extra
// calendar day per event. Fixed in gmail-team-days-worked's copy of this
// function by parsing all-day dates as local dates directly instead of via
// `new Date(dateString)`; not yet applied here since this only feeds
// percentages here (not exact day counts used to check invoices).
function haalAfwezighedenOp(email, periodeStart, periodeEinde) {
  const dagData = {};
  try {
    const response = Calendar.Events.list(email, {
      timeMin: periodeStart.toISOString(),
      timeMax: periodeEinde.toISOString(),
      eventTypes: ["outOfOffice"],
      singleEvents: true,
    });
    (response.items || []).forEach(event => {
      const evStart = new Date(event.start.dateTime || event.start.date);
      const evEinde = new Date(event.end.dateTime || event.end.date);
      const isAllDay = !!event.start.date && !event.start.dateTime;
      const cursor = new Date(evStart); cursor.setHours(0,0,0,0);
      while (cursor < evEinde) {
        const dag = cursor.getDay();
        if (dag !== 0 && dag !== 6) {
          const key = cursor.toDateString();
          if (isAllDay) {
            dagData[key] = WERKUREN_PER_DAG;
          } else {
            const dagStart = new Date(cursor); dagStart.setHours(9,0,0,0);
            const dagEinde = new Date(cursor); dagEinde.setHours(17,0,0,0);
            const overlapMs = Math.min(evEinde, dagEinde) - Math.max(evStart, dagStart);
            if (overlapMs > 0) {
              dagData[key] = Math.min(WERKUREN_PER_DAG, (dagData[key] || 0) + overlapMs / (1000 * 60 * 60));
            }
          }
        }
        cursor.setDate(cursor.getDate() + 1);
      }
    });
    return { dagData, fout: null };
  } catch (e) {
    Logger.log("Error for " + email + ": " + e.message);
    return { dagData, fout: e.message };
  }
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
    .filter(t => t.getHandlerFunction() === 'doRefresh')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshCapaciteit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert('✅ Auto-refresh installed — capacity will load automatically on open.');
}

// ============================================================
// HELPERS
// ============================================================
function getMaandagVanWeek(datum) {
  const d = new Date(datum);
  const diff = d.getDay() === 0 ? -6 : 1 - d.getDay();
  d.setDate(d.getDate() + diff);
  d.setHours(0,0,0,0);
  return d;
}

function colLetter(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function formatDatum(d) {
  return d.toLocaleDateString("en-GB", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
}

function formatDatumKort(d) {
  return d.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
}
