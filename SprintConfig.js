// ============================================================
// SPRINT CONFIG — read from column H, auto-generate future sprints,
// and derive week/absence math from a sprint's date range
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

// Picks the 4 sprints starting from "today"'s sprint (or the next upcoming
// one) and flattens them into ordered week columns. Each group also gets a
// Hero checkbox column right before its own weeks — heroKolIdx is that
// column, startKolIdx is the group's first week column, shifted to make room.
function bouwSprintGroepen(sprints, vandaag) {
  const huidigeSprint = sprints.find(s => vandaag >= s.start && vandaag < s.einde)
    || sprints.slice().sort((a, b) => a.start - b.start).find(s => s.start > vandaag);

  const sprintGroepen = [];
  if (huidigeSprint) {
    for (let i = 0; i < 4; i++) {
      const sprint = sprints.find(s => s.nr === huidigeSprint.nr + i);
      if (sprint) sprintGroepen.push({ sprint, weken: buildSprintWeken(sprint) });
    }
  }

  const weekKolommen = sprintGroepen.flatMap(g => g.weken);
  let kolIdx = 0;
  sprintGroepen.forEach(groep => {
    groep.heroKolIdx = kolIdx;
    groep.startKolIdx = kolIdx + 1;
    kolIdx += 1 + groep.weken.length;
  });

  return { sprintGroepen, weekKolommen, aantalKolommen: 2 + kolIdx };
}
