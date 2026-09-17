// ============================================================
// TEAMS — parsing the "Teams" tab's team/member/email/id columns
// ============================================================

// Parses the "Teams" tab into { teamNaam: [{naam,email}, ...] } plus the
// per-team Jira board id. The id is read from a separate "id" column that
// isn't at a fixed index — it's found dynamically because the E-F section
// it lives in shifts depending on how many team columns come before it.
function verzamelTeams(teamsSheet) {
  const data = teamsSheet.getDataRange().getValues();
  const teams = {};
  const teamIds = {};
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
  return { teams, teamIds, teamNamen: Object.keys(teams) };
}
