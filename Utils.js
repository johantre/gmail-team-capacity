// ============================================================
// UTILS — pure helper functions
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
