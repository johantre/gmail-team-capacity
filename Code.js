// ============================================================
// TEAM CAPACITEIT - Google Apps Script v3
// ============================================================
// Configuratie via het tabblad "Teams" in je Sheet:
//
//   | Team   | Naam      | Email                |
//   |--------|-----------|----------------------|
//   | Team A | Pier      | pier@robaws.be       |
//   | Team A | Pol       | pol@robaws.be        |
//   | Team B | Jef       | jef@robaws.be        |
//   | ...    | ...       | ...                  |
//
// Voeg gewoon rijen toe voor Team C, Team D, etc.
// ============================================================

const WEKEN_VOORUIT = 4;
const WERKUREN_PER_DAG = 8;
const SHEET_TEAMS = "Teams";
const SHEET_CAPACITEIT = "Capaciteit";

const TEAM_KLEUREN = {
  "Team A": "#1a73e8",
  "Team B": "#188038",
  "Team C": "#e37400",
  "Team D": "#a142f4",
};

// ============================================================
// HOOFDFUNCTIE
// ============================================================
function refreshCapaciteit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Lees teamconfiguratie ----
  const teamsSheet = ss.getSheetByName(SHEET_TEAMS);
  if (!teamsSheet) {
    SpreadsheetApp.getUi().alert('❌ Tabblad "Teams" niet gevonden.');
    return;
  }

  const data = teamsSheet.getDataRange().getValues();
  const teams = {};

  data.forEach(rij => {
    const [team, naam, email] = rij.map(v => v.toString().trim());
    if (!email || !email.includes("@")) return;
    if (!teams[team]) teams[team] = [];
    teams[team].push({ naam, email });
  });

  const teamNamen = Object.keys(teams);
  if (teamNamen.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Geen teams gevonden.');
    return;
  }

  // ---- Bereken weken ----
  const vandaag = new Date();
  const maandag = getMaandagVanWeek(vandaag);
  const weken = [];
  for (let i = 0; i < WEKEN_VOORUIT; i++) {
    const start = new Date(maandag);
    start.setDate(maandag.getDate() + i * 7);
    const einde = new Date(start);
    einde.setDate(start.getDate() + 4);
    weken.push({ start, einde });
  }

  // ---- Haal alle unieke emails op ----
  const alleEmails = [...new Set(teamNamen.flatMap(t => teams[t].map(l => l.email)))];

  // ---- ÉÉN API call per persoon voor de volledige periode ----
  // Resultaat: { "email": { "Mon Jun 23 2026": urenAfwezig, ... }, ... }
  const periodeStart = weken[0].start;
  const periodeEinde = new Date(weken[weken.length - 1].einde);
  periodeEinde.setHours(23, 59, 59, 999);

  const afwezigCache = {};
  alleEmails.forEach(email => {
    afwezigCache[email] = haalAfwezighedenOp(email, periodeStart, periodeEinde);
  });

  // ---- Capaciteitssheet klaarzetten ----
  let sheet = ss.getSheetByName(SHEET_CAPACITEIT);
  if (!sheet) sheet = ss.insertSheet(SHEET_CAPACITEIT);
  sheet.clearContents();
  sheet.clearFormats();

  const aantalKolommen = 1 + WEKEN_VOORUIT;

  // ---- Hoofdtitel ----
  sheet.getRange(1, 1, 1, aantalKolommen).merge();
  sheet.getRange(1, 1)
    .setValue("🗓️ Team Capaciteit — bijgewerkt op " + formatDatum(vandaag))
    .setFontSize(13).setFontWeight("bold")
    .setBackground("#202124").setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  sheet.setColumnWidth(1, 220);
  for (let k = 2; k <= aantalKolommen; k++) sheet.setColumnWidth(k, 150);

  let rij = 2;

  // ---- Per team een blok ----
  teamNamen.forEach(teamNaam => {
    const leden = teams[teamNaam];
    const teamKleur = TEAM_KLEUREN[teamNaam] || "#5f6368";

    // Teamtitel
    sheet.getRange(rij, 1, 1, aantalKolommen).merge();
    sheet.getRange(rij, 1)
      .setValue(teamNaam)
      .setFontSize(11).setFontWeight("bold")
      .setBackground(teamKleur).setFontColor("#ffffff")
      .setHorizontalAlignment("left");
    sheet.setRowHeight(rij, 28);
    rij++;

    // Weekheaders
    const headers = ["Teamlid"];
    weken.forEach(w => {
      headers.push(`Week ${getWeeknummer(w.start)}\n${formatDatumKort(w.start)} – ${formatDatumKort(w.einde)}`);
    });
    sheet.getRange(rij, 1, 1, aantalKolommen).setValues([headers])
      .setFontWeight("bold").setBackground("#e8f0fe")
      .setHorizontalAlignment("center").setWrap(true);
    sheet.setRowHeight(rij, 44);
    rij++;

    // Teamleden — berekening vanuit cache, geen extra API calls
    leden.forEach(lid => {
      const dagCache = afwezigCache[lid.email] || {};
      const rowData = [lid.naam];
      const afwezigPerWeek = [];

      weken.forEach(week => {
        const afwezig = berekenAfwezigVanuitCache(dagCache, week.start, week.einde);
        afwezigPerWeek.push(afwezig);
        const beschikbaar = Math.round((5 - afwezig) * 4) / 4;
        const label = Number.isInteger(beschikbaar)
          ? beschikbaar
          : beschikbaar.toFixed(2).replace('.', ',');
        rowData.push(label + " / 5 dagen");
      });

      sheet.getRange(rij, 1, 1, aantalKolommen).setValues([rowData]);
      sheet.getRange(rij, 2, 1, WEKEN_VOORUIT).setHorizontalAlignment("center");

      afwezigPerWeek.forEach((afwezig, idx) => {
        const beschikbaar = Math.round((5 - afwezig) * 4) / 4;
        const cel = sheet.getRange(rij, 2 + idx);
        if (beschikbaar === 5) cel.setBackground("#e6f4ea");
        else if (beschikbaar >= 3) cel.setBackground("#fef9c3");
        else cel.setBackground("#fce8e6");
      });

      rij++;
    });

    // Totaalrij — ook vanuit cache
    const totaalRij = [`${teamNaam} totaal (%)`];
    weken.forEach(week => {
      let totaalBeschikbaar = 0;
      leden.forEach(lid => {
        const dagCache = afwezigCache[lid.email] || {};
        const afwezig = berekenAfwezigVanuitCache(dagCache, week.start, week.einde);
        totaalBeschikbaar += (5 - afwezig);
      });
      const pct = Math.round((totaalBeschikbaar / (5 * leden.length)) * 100);
      totaalRij.push(pct + "%");
    });

    sheet.getRange(rij, 1, 1, aantalKolommen).setValues([totaalRij])
      .setFontWeight("bold").setBackground("#f1f3f4");
    sheet.getRange(rij, 2, 1, WEKEN_VOORUIT).setHorizontalAlignment("center");
    rij++;

    // Lege rij
    sheet.setRowHeight(rij, 16);
    rij++;
  });

  SpreadsheetApp.getUi().alert("✅ Capaciteit bijgewerkt!");
}

// ============================================================
// ÉÉN API call per persoon — geeft map van datum → afwezige uren
// ============================================================
function haalAfwezighedenOp(email, periodeStart, periodeEinde) {
  const afwezigUrenPerDag = {};

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

      const cursor = new Date(evStart);
      cursor.setHours(0, 0, 0, 0);

      while (cursor < evEinde) {
        const dag = cursor.getDay();
        if (dag !== 0 && dag !== 6) {
          const key = cursor.toDateString();
          if (isAllDay) {
            afwezigUrenPerDag[key] = WERKUREN_PER_DAG;
          } else {
            const dagStart = new Date(cursor); dagStart.setHours(0, 0, 0, 0);
            const dagEinde = new Date(cursor); dagEinde.setHours(23, 59, 59, 999);
            const overlapMs = Math.min(evEinde, dagEinde) - Math.max(evStart, dagStart);
            if (overlapMs > 0) {
              const uren = overlapMs / (1000 * 60 * 60);
              afwezigUrenPerDag[key] = Math.min(WERKUREN_PER_DAG, (afwezigUrenPerDag[key] || 0) + uren);
            }
          }
        }
        cursor.setDate(cursor.getDate() + 1);
      }
    });

  } catch (e) {
    Logger.log("Fout voor " + email + ": " + e.message);
  }

  return afwezigUrenPerDag;
}

// Berekening vanuit cache — geen API call
function berekenAfwezigVanuitCache(dagCache, weekStart, weekEinde) {
  let totaalUren = 0;
  const cursor = new Date(weekStart);
  while (cursor <= weekEinde) {
    const key = cursor.toDateString();
    totaalUren += dagCache[key] || 0;
    cursor.setDate(cursor.getDate() + 1);
  }
  return Math.min(Math.round((totaalUren / WERKUREN_PER_DAG) * 4) / 4, 5);
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📅 Capaciteit")
    .addItem("🔄 Refresh capaciteit", "refreshCapaciteit")
    .addToUi();
}

// ============================================================
// HULPFUNCTIES
// ============================================================
function getMaandagVanWeek(datum) {
  const d = new Date(datum);
  const diff = d.getDay() === 0 ? -6 : 1 - d.getDay();
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getWeeknummer(datum) {
  const d = new Date(Date.UTC(datum.getFullYear(), datum.getMonth(), datum.getDate()));
  const dagNr = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dagNr + 3);
  const eersteJan = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((d - eersteJan) / 86400000 - 3 + (eersteJan.getUTCDay() + 6) % 7) / 7);
}

function formatDatum(d) {
  return d.toLocaleDateString("nl-BE", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
}

function formatDatumKort(d) {
  return d.toLocaleDateString("nl-BE", { day: "numeric", month: "short" });
}