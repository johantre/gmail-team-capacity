// ============================================================
// HEROES — persistence of the "🦸 Hero" checkbox state across a refresh
// ============================================================

// Reads which (team, member, sprint nr) Hero checkboxes were checked on the
// sheet as it existed before this refresh, so a refresh doesn't silently
// un-check everyone. Sprint header cells sit at the leftmost column of their
// merge — which is exactly the Hero column for that sprint — so scanning for
// "Sprint N" text tells us which column to read per team block, regardless
// of how many sprints/columns are currently shown.
function leesOpgeslagenHeroes(values, teamNamen) {
  const heroes = {};
  let huidigTeam = null;
  let sprintKolommen = [];
  values.forEach(row => {
    const col1 = (row[0] || '').toString().trim();
    if (teamNamen.includes(col1)) {
      huidigTeam = col1;
      sprintKolommen = [];
      return;
    }
    if (!huidigTeam) return;
    const gevondenSprints = [];
    for (let c = 2; c < row.length; c++) {
      const match = (row[c] || '').toString().match(/Sprint\s+(\d+)/i);
      if (match) gevondenSprints.push({ kol: c, sprintNr: parseInt(match[1]) });
    }
    if (gevondenSprints.length > 0) { sprintKolommen = gevondenSprints; return; }
    if (col1 === 'Team member' || col1 === '') { if (!col1) huidigTeam = null; return; }
    if (col1 === 'MD / Sprint (%)' || col1 === '→ Projected SP') { huidigTeam = null; return; }
    sprintKolommen.forEach(({ kol, sprintNr }) => {
      if (row[kol] === true) heroes[`${huidigTeam}||${col1}||${sprintNr}`] = true;
    });
  });
  return heroes;
}
