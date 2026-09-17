// ============================================================
// CALENDAR SERVICE — one Calendar API call per person, returns data + any error
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
