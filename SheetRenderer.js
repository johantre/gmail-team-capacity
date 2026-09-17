// ============================================================
// SHEET RENDERER — everything that reads/writes the "Capacity" sheet itself.
// Called by doRefresh() (Code.js), which gathers the data these functions
// draw and threads the row cursor (`rij`) and `dagCelRanges` accumulator
// through them in sequence, one team at a time.
// ============================================================

// Reads velocities and Hero checkboxes from the sheet as it existed before
// this refresh, so a refresh doesn't wipe user-entered/-checked state. Must
// run before maakSheetLeeg() clears the sheet.
function leesBestaandeSheetState(ss, teamNamen) {
  const sheet = ss.getSheetByName(SHEET_CAPACITEIT);
  const savedVelocities = {};
  let savedHeroes = {};
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
    savedHeroes = leesOpgeslagenHeroes(existingVals, teamNamen);
  }
  return { savedVelocities, savedHeroes };
}

// Gets or creates the "Capacity" sheet, clears it, and lays out its fixed
// frame: column widths and the title row. Team content is drawn row-by-row
// by tekenTeamHeader/tekenSprintEnWeekHeaders/tekenLedenRijen/
// tekenSamenvattingRijen below, called once per team by doRefresh.
function maakSheetLeeg(ss, aantalKolommen, sprintGroepen, vandaag) {
  let sheet = ss.getSheetByName(SHEET_CAPACITEIT);
  if (!sheet) sheet = ss.insertSheet(SHEET_CAPACITEIT);
  sheet.getImages().forEach(img => img.remove());
  sheet.clearContents();
  sheet.clearFormats();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 36);
  for (let k = 3; k <= aantalKolommen; k++) sheet.setColumnWidth(k, 120);
  sprintGroepen.forEach(groep => sheet.setColumnWidth(groep.heroKolIdx + 3, 56));

  // Main title
  sheet.getRange(1, 1, 1, aantalKolommen).merge();
  sheet.getRange(1, 1)
    .setValue("📅  Team Capacity — updated on " + formatDatum(vandaag))
    .setFontFamily("Montserrat").setFontSize(13).setFontWeight("bold")
    .setBackground(GRIJS).setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 36);

  return sheet;
}

// Team title row: [name col1] [📊 col2] [Max Capacity Velocity: col3..N-1]
// [velocity col N].
function tekenTeamHeader(sheet, rij, teamNaam, aantalKolommen, refreshData) {
  const velocity = refreshData.savedVelocities[teamNaam] || 0;
  const teamId = refreshData.teamIds[teamNaam] || '';
  const teamKleur = TEAM_KLEUREN[teamNaam] || "#5f6368";
  const teamStyle = { font: "Montserrat", bg: teamKleur, fg: "#ffffff" };

  sheet.getRange(rij, 1).setValue(teamNaam)
    .setFontFamily(teamStyle.font).setFontSize(11).setFontWeight("bold")
    .setBackground(teamStyle.bg).setFontColor(teamStyle.fg)
    .setHorizontalAlignment("left");

  const cel2 = sheet.getRange(rij, 2);
  cel2.setBackground(teamStyle.bg).setFontColor(teamStyle.fg)
    .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  if (teamId) {
    const jiraUrl = `https://eforge.atlassian.net/jira/software/c/projects/ACME/boards/${teamId}/reports/velocity`;
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
}

// Sprint header row (merged per sprint, including its Hero column) + week
// header row (Hero label + date range/working days per week). Returns the
// row after these two.
function tekenSprintEnWeekHeaders(sheet, rij, sprintGroepen, aantalKolommen) {
  sheet.getRange(rij, 1).setValue("").setBackground(BLAUW_LICHT);
  sheet.getRange(rij, 2).setBackground(BLAUW_LICHT);
  sprintGroepen.forEach(groep => {
    const heroKol = groep.heroKolIdx + 3;
    const kolBreedte = 1 + groep.weken.length;
    const cel = sheet.getRange(rij, heroKol, 1, kolBreedte);
    if (kolBreedte > 1) cel.merge();
    const sprintLabel = `${groep.sprint.naam} (${formatDatumKort(groep.sprint.start)} – ${formatDatumKort(groep.sprint.einde)})`;
    cel.setValue(sprintLabel)
      .setFontFamily("Montserrat").setFontWeight("bold").setFontSize(10)
      .setBackground(BLAUW_LICHT).setFontColor(GRIJS).setHorizontalAlignment("center");
  });
  sheet.setRowHeight(rij, 22);
  rij++;

  const headers = ["Team member", ""];
  sprintGroepen.forEach(groep => {
    headers.push("🦸 Hero");
    groep.weken.forEach(wk => {
      headers.push(`${formatDatumKort(wk.week.start)} – ${formatDatumKort(wk.week.einde)}\n(${wk.werkdagen})`);
    });
  });
  sheet.getRange(rij, 1, 1, aantalKolommen).setValues([headers])
    .setFontFamily("Montserrat").setFontWeight("normal")
    .setBackground(BLAUW_LICHT).setFontColor(GRIJS)
    .setHorizontalAlignment("center").setWrap(true);
  sheet.getRange(rij, 1).setFontWeight("bold");
  sheet.setRowHeight(rij, 44);
  rij++;

  return rij;
}

// Member rows: names, per-sprint Hero checkboxes, and "x / y" availability
// formulas. Formulas (not plain values) so toggling a Hero checkbox
// recalculates that member's line — and, since the hidden Current/Full MD
// rows (tekenSamenvattingRijen) read these cells via SUMPRODUCT, → Projected
// SP recalculates too. Color is conditional formatting, applied once for the
// whole sheet by pasConditionalFormattingToe over the ranges pushed onto
// dagCelRanges here, not a static background, so it also updates live.
//
// Written as bulk setValues/setFormulas calls per sprint group (one call for
// all members at once) rather than per cell — with a Hero column per sprint
// that's 4x as many cells as before, per-cell calls made a refresh
// noticeably slower.
function tekenLedenRijen(sheet, eersteLedenRij, leden, teamNaam, sprintGroepen, aantalKolommen, refreshData, dagCelRanges) {
  const { savedHeroes, afwezigCache, kalenderFouten } = refreshData;

  sheet.getRange(eersteLedenRij, 1, leden.length, 1).setValues(leden.map(lid => [lid.naam]));
  sheet.getRange(eersteLedenRij, 1, leden.length, aantalKolommen).setFontFamily("Montserrat");
  sheet.getRange(eersteLedenRij, 3, leden.length, aantalKolommen - 2).setHorizontalAlignment("center");

  sprintGroepen.forEach(groep => {
    const heroKolAbs = groep.heroKolIdx + 3;
    const heroWaarden = leden.map(lid => [!!savedHeroes[`${teamNaam}||${lid.naam}||${groep.sprint.nr}`]]);
    const heroRange = sheet.getRange(eersteLedenRij, heroKolAbs, leden.length, 1);
    heroRange.insertCheckboxes();
    heroRange.setValues(heroWaarden);

    const startKolAbs = groep.startKolIdx + 3;
    const formules = leden.map((lid, i) => {
      if (kalenderFouten[lid.email]) return groep.weken.map(() => '=""');
      const dagCache = afwezigCache[lid.email] || {};
      const heroRef = `${colLetter(heroKolAbs)}${eersteLedenRij + i}`;
      return groep.weken.map(wk => {
        const afwezig = berekenAfwezigInWeek(dagCache, wk.week.start, wk.week.einde, wk.sprint.start, wk.sprint.einde);
        const beschikbaar = Math.max(0, Math.round((wk.werkdagen - afwezig) * 4) / 4);
        const label = Number.isInteger(beschikbaar) ? beschikbaar : beschikbaar.toFixed(2).replace('.', ',');
        return `=IF(${heroRef};"0 / ${wk.werkdagen}";"${label} / ${wk.werkdagen}")`;
      });
    });
    const weekRange = sheet.getRange(eersteLedenRij, startKolAbs, leden.length, groep.weken.length);
    weekRange.setFormulas(formules);
    dagCelRanges.push(weekRange);

    // Calendar-access errors override with plain "⚠️ no access" text — rare, so a per-row fixup is fine.
    leden.forEach((lid, i) => {
      if (!kalenderFouten[lid.email]) return;
      sheet.getRange(eersteLedenRij + i, startKolAbs, 1, groep.weken.length).setValue("⚠️ no access");
    });
  });

  leden.forEach((lid, i) => {
    if (!kalenderFouten[lid.email]) return;
    sheet.getRange(eersteLedenRij + i, 1, 1, aantalKolommen).setBackground("#fff3e0").setFontColor("#e65100");
  });
}

// Summary rows below the member list: MD/Sprint (%) (visible), hidden
// Current MD and Full Cap. MD (SUMPRODUCT over the member "x / y" cells,
// feeding the next row), → Projected SP (velocity * current/full ratio),
// and the vertical dividers between sprints. Returns the row after these.
function tekenSamenvattingRijen(sheet, rij, teamHeaderRij, eersteLedenRij, laatsLedenRij, leden, teamNaam, sprintGroepen, aantalKolommen, refreshData) {
  const { savedHeroes, afwezigCache, kalenderFouten } = refreshData;

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
    // Color the Hero column at this row too, so the bar reads as one
    // continuous line instead of breaking at every Hero column's width.
    sheet.getRange(rij, groep.heroKolIdx + 3).setBackground(TURQUOISE_LICHT);
    const maxMandagen = geldigeLeden * groep.weken.reduce((s, wk) => s + wk.werkdagen, 0);
    let sprintMandagen = 0;
    groep.weken.forEach(wk => {
      leden.forEach(lid => {
        if (kalenderFouten[lid.email]) return;
        if (savedHeroes[`${teamNaam}||${lid.naam}||${groep.sprint.nr}`]) return; // Hero this sprint: 0 MD
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
  sprintGroepen.forEach(groep => {
    sheet.getRange(rij, groep.heroKolIdx + 3).setBackground(TURQUOISE_LICHT);
  });
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

  return rij;
}

// Colors the day cells via conditional formatting instead of a static
// background, so it stays correct live when a Hero checkbox is toggled —
// same "x / y" parsing as the hidden SUMPRODUCT rows in
// tekenSamenvattingRijen, wrapped in IFERROR so cells that don't match (e.g.
// "⚠️ no access") get no color. The formula is anchored to the first range's
// top-left cell and applies relatively across every other range in the
// rule, same as a normal conditional format rule dragged across multiple
// areas.
function pasConditionalFormattingToe(sheet, dagCelRanges) {
  if (dagCelRanges.length === 0) {
    sheet.setConditionalFormatRules([]);
    return;
  }
  const anker = `${colLetter(dagCelRanges[0].getColumn())}${dagCelRanges[0].getRow()}`;
  const x = `VALUE(TRIM(LEFT(${anker};FIND(" / ";${anker})-1)))`;
  const y = `VALUE(TRIM(MID(${anker};FIND(" / ";${anker})+3;100)))`;
  // Compared as 2*x vs y instead of x vs 0.5*y — a literal decimal point in
  // a formula isn't valid in a comma-decimal locale (this sheet's), which
  // silently broke both of these rules (caught by IFERROR, cell stayed
  // white) while the green rule — no decimal literal — worked fine.
  const regels = [
    { kleur: "#e6f4ea", formule: `=IFERROR(${x}>=${y};FALSE)` },
    { kleur: "#fef9c3", formule: `=IFERROR(AND(2*${x}>=${y};${x}<${y});FALSE)` },
    { kleur: "#fce8e6", formule: `=IFERROR(2*${x}<${y};FALSE)` },
  ].map(({ kleur, formule }) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formule)
      .setBackground(kleur)
      .setRanges(dagCelRanges)
      .build()
  );
  sheet.setConditionalFormatRules(regels);
}

// Error footer listing Calendar accounts the refresh couldn't read.
function tekenFoutenblok(ss, sheet, rij, kalenderFouten, aantalKolommen, vanSidebar) {
  const foutenLijst = Object.entries(kalenderFouten);
  if (foutenLijst.length === 0) return;

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
  Logger.log(`Calendar access errors: ${foutenLijst.map(([e]) => e).join(', ')}`);
}
