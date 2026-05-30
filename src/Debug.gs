// Manual introspection helpers for inspecting Google Health datapoints.
// Run from the Apps Script editor; output goes to the Executions log.
// Read-only.
//
// Purpose: learn the JSON shape of manually-logged Strength Training
// workouts (the "Workout summary" with per-exercise sets/reps/weight shown
// in the Google Health app) so we can mirror that structure when creating
// our own exercise datapoints from the sheet's exercise columns.
//
// Usage from the editor (all callable with no args):
//   dumpLatestMatchedSession()  -> newest row's Matched Health Session
//   dumpLatestCreatedId()       -> newest row's first Created Health ID
//   dumpTodaysForeignStrength() -> today's foreign Strength sessions

function dumpLatestMatchedSession() {
  const { rows } = readRows();
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (r.matchedHealthSession) {
      console.log('Row ' + r.rowNum + ' (' + ymd(r.date) + ') matched session: ' + r.matchedHealthSession);
      return dumpDataPoint_(r.matchedHealthSession);
    }
  }
  console.log('No row has a Matched Health Session.');
  return null;
}

function dumpLatestCreatedId() {
  const { rows } = readRows();
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (r.healthIds && r.healthIds.length > 0) {
      const name = r.healthIds[0];
      console.log('Row ' + r.rowNum + ' (' + ymd(r.date) + ') created ID: ' + name);
      return dumpDataPoint_(name);
    }
  }
  console.log('No row has a Created Health ID.');
  return null;
}

function dumpTodaysForeignStrength() {
  const date = new Date();
  const candidates = listForeignStrengthOnDate(date);
  console.log('Foreign Strength Training on ' + ymd(date) + ': ' + candidates.length + ' datapoint(s)');
  candidates.forEach((c, i) => {
    console.log('=== [' + (i + 1) + '/' + candidates.length + '] ' + c.name);
    dumpDataPoint_(c.name);
  });
  return candidates;
}

function dumpDataPoint_(name) {
  const point = getDataPointByName(name);
  console.log(JSON.stringify(point, null, 2));
  return point;
}
