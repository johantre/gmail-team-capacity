// ============================================================
// TEAM CAPACITY - Google Apps Script v3
// ============================================================
// Configuration via the "Teams" tab in your Sheet:
//
//   | Team   | Name      | Email                |   ...   | Sprint naming signature    |
//   |--------|-----------|----------------------|---------|---------------------------|
//   | Team A | Pier      | pier@robaws.be       |         | Sprint 14 (16 jun – 30 jun)|
//   | Team A | Pol       | pol@robaws.be        |         | Sprint 15 (30 jun – 14 jul)|
//
// Sprint naming signature is in column H (index 7).
// Format: Sprint {nr} ({day} {month} – {day} {month})
// Convention: first day of sprint counts, last (overlapping) day does NOT count.
// Sprints span 3 calendar weeks (e.g. 4d / 5d / 1d). No half-day adjustments.
// ============================================================

const WERKUREN_PER_DAG = 8;
const SHEET_TEAMS = "Teams";
const SHEET_CAPACITEIT = "Capacity";

// Robaws brand colors
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
    // Generate until we have at least 2 sprints beyond today
    while (laatste.einde <= vandaag || sprints.filter(s => s.start > vandaag).length < 2) {
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
// ============================================================
function refreshCapaciteit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const teamsSheet = ss.getSheetByName(SHEET_TEAMS);
  if (!teamsSheet) { SpreadsheetApp.getUi().alert('❌ Tab "Teams" not found.'); return; }

  const data = teamsSheet.getDataRange().getValues();
  const teams = {};
  data.forEach(rij => {
    const [team, naam, email] = rij.map(v => v.toString().trim());
    if (!email || !email.includes("@")) return;
    if (!teams[team]) teams[team] = [];
    teams[team].push({ naam, email });
  });

  const teamNamen = Object.keys(teams);
  if (teamNamen.length === 0) { SpreadsheetApp.getUi().alert('❌ No teams found.'); return; }

  const sprints = leesSprintConfig(teamsSheet);
  const vandaag = new Date();

  // Current sprint = today within [start, einde), fallback to next upcoming
  const huidigeSprint = sprints.find(s => vandaag >= s.start && vandaag < s.einde)
    || sprints.slice().sort((a, b) => a.start - b.start).find(s => s.start > vandaag);
  const volgendeSprint = huidigeSprint
    ? sprints.find(s => s.nr === huidigeSprint.nr + 1)
    : null;

  // Build sprint groups: 3 weeks per sprint
  const sprintGroepen = [];
  [huidigeSprint, volgendeSprint].filter(Boolean).forEach(sprint => {
    const weken = buildSprintWeken(sprint);
    sprintGroepen.push({ sprint, weken });
  });

  // Flatten to ordered week columns and set startKolIdx per group
  const weekKolommen = sprintGroepen.flatMap(g => g.weken);
  let kolIdx = 0;
  sprintGroepen.forEach(groep => { groep.startKolIdx = kolIdx; kolIdx += groep.weken.length; });

  const aantalKolommen = 1 + weekKolommen.length;

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

  // Set up sheet
  let sheet = ss.getSheetByName(SHEET_CAPACITEIT);
  if (!sheet) sheet = ss.insertSheet(SHEET_CAPACITEIT);
  sheet.clearContents();
  sheet.clearFormats();

  sheet.setColumnWidth(1, 200);
  for (let k = 2; k <= aantalKolommen; k++) sheet.setColumnWidth(k, 120);

  // Main title
  sheet.getRange(1, 1, 1, aantalKolommen).merge();
  sheet.getRange(1, 1)
    .setValue("🗓️ Team Capacity — updated on " + formatDatum(vandaag))
    .setFontFamily("Montserrat").setFontSize(13).setFontWeight("bold")
    .setBackground(GRIJS).setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  let rij = 2;

  teamNamen.forEach(teamNaam => {
    const leden = teams[teamNaam];
    const teamKleur = TEAM_KLEUREN[teamNaam] || "#5f6368";

    // Team title
    sheet.getRange(rij, 1, 1, aantalKolommen).merge();
    sheet.getRange(rij, 1).setValue(teamNaam)
      .setFontFamily("Montserrat").setFontSize(11).setFontWeight("bold")
      .setBackground(teamKleur).setFontColor("#ffffff")
      .setHorizontalAlignment("left");
    sheet.setRowHeight(rij, 28);
    rij++;

    // Sprint header row — merged over each sprint's columns
    sheet.getRange(rij, 1).setValue("").setBackground(BLAUW_LICHT);
    sprintGroepen.forEach(groep => {
      const kolStart = groep.startKolIdx + 2;
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
    const headers = ["Team member"];
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
      const rowData = [lid.naam];
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
      sheet.getRange(rij, 2, 1, weekKolommen.length).setHorizontalAlignment("center");

      if (heeftFout) {
        sheet.getRange(rij, 1, 1, aantalKolommen).setBackground("#fff3e0").setFontColor("#e65100");
      } else {
        beschikbaarData.forEach(({ beschikbaar, werkdagen }, idx) => {
          const cel = sheet.getRange(rij, 2 + idx);
          const ratio = werkdagen > 0 ? beschikbaar / werkdagen : 1;
          if (ratio >= 1) cel.setBackground("#e6f4ea");
          else if (ratio >= 0.5) cel.setBackground("#fef9c3");
          else cel.setBackground("#fce8e6");
        });
      }
      rij++;
    });

    // MD / Sprint (%) row — merged per sprint
    const geldigeLeden = leden.filter(l => !kalenderFouten[l.email]).length;
    sheet.getRange(rij, 1).setValue("MD / Sprint (%)")
      .setFontFamily("Montserrat").setFontWeight("bold")
      .setBackground(TURQUOISE_LICHT).setFontColor(GRIJS);
    sprintGroepen.forEach(groep => {
      const kolStart = groep.startKolIdx + 2;
      const kolBreedte = groep.weken.length;
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
    });
    rij++;

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
    SpreadsheetApp.getUi().alert(
      `⚠️ Capacity updated with ${foutenLijst.length} error(s).\n\n` +
      `No access to calendar of:\n` +
      foutenLijst.map(([email]) => `• ${email}`).join("\n") +
      `\n\nCheck if the calendars are shared with the account running this script.`
    );
  } else {
    SpreadsheetApp.getUi().alert("✅ Capacity updated!");
  }
}

// ============================================================
// ONE API call per person — returns data + any error
// ============================================================
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
    .addToUi();
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

function formatDatum(d) {
  return d.toLocaleDateString("en-GB", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
}

function formatDatumKort(d) {
  return d.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
}
