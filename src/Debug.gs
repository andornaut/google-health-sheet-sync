// Manual introspection helpers. Run from the Apps Script editor; output goes
// to the Executions log (console.log / Logger). All read-only.
//
// All functions accept zero arguments and use sensible defaults so they can be
// run directly from the editor's function picker:
//   showExercisesOnDate()  -> today
//   showWeightsOnDate()    -> today
//   showRecentExercises()  -> last 7 days
//   showRowSyncState()     -> last data row in the sheet
//   showDataPointByName()  -> requires a resource name (no default)

function parseYmd_(yyyymmdd) {
  if (!yyyymmdd) return new Date();
  const s = String(yyyymmdd).trim();
  const m = /^(\d{4})-?(\d{2})-?(\d{2})$/.exec(s);
  if (!m) throw new Error('Bad date "' + s + '"; expected yyyy-mm-dd or yyyymmdd.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function showExercisesOnDate(yyyymmdd) {
  const date = parseYmd_(yyyymmdd);
  const points = listExercisesOnDate(date);
  console.log('Exercises on ' + ymd(date) + ': ' + points.length + ' datapoint(s)');
  points.forEach((p, i) => {
    const ex = p.exercise || {};
    const iv = ex.interval || {};
    console.log('--- [' + (i + 1) + '] ' + (p.name || '<no name>'));
    console.log('  start: ' + iv.startTime + '  end: ' + iv.endTime);
    console.log('  type: ' + ex.exerciseType + '  displayName: ' + ex.displayName);
    console.log('  recordingMethod: ' + ((p.dataSource && p.dataSource.recordingMethod) || '<none>'));
    console.log('  notes:\n' + (ex.notes || '<none>'));
  });
  return points;
}

function showWeightsOnDate(yyyymmdd) {
  const date = parseYmd_(yyyymmdd);
  const points = listWeightsOnDate(date);
  console.log('Weights on ' + ymd(date) + ': ' + points.length + ' datapoint(s)');
  points.forEach((p, i) => {
    const w = p.weight || {};
    const grams = w.weightGrams || 0;
    const lbs = grams / GRAMS_PER_LB;
    const t = w.sampleTime || {};
    console.log('--- [' + (i + 1) + '] ' + (p.name || '<no name>'));
    console.log('  sampleTime: ' + t.physicalTime);
    console.log('  weight: ' + grams + 'g (' + lbs.toFixed(2) + ' lb)');
    console.log('  recordingMethod: ' + ((p.dataSource && p.dataSource.recordingMethod) || '<none>'));
  });
  return points;
}

function showRecentExercises(days) {
  const n = days && days > 0 ? Math.floor(days) : 7;
  const today = new Date();
  const totals = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - i);
    const points = showExercisesOnDate(ymd(d));
    totals.push({ date: ymd(d), count: points.length });
  }
  console.log('=== Summary (last ' + n + ' day(s)) ===');
  totals.forEach(t => console.log('  ' + t.date + ': ' + t.count));
  return totals;
}

function showRowSyncState(rowNum) {
  const { rows } = readRows();
  if (rows.length === 0) {
    console.log('Sheet has no data rows.');
    return null;
  }
  const target = rowNum
    ? rows.find(r => r.rowNum === Number(rowNum))
    : rows[rows.length - 1];
  if (!target) {
    console.log('Row ' + rowNum + ' not found or has no date.');
    return null;
  }
  console.log('Row ' + target.rowNum + '  date: ' + ymd(target.date));
  console.log('  bodyweight: ' + (target.bodyweight === null ? '<none>' : target.bodyweight + ' lb'));
  console.log('  syncedAt: ' + (target.syncedAt || '<dirty>'));
  console.log('  matchedHealthSession: ' + (target.matchedHealthSession || '<none>'));
  console.log('  healthIds (' + target.healthIds.length + '):');
  target.healthIds.forEach(n => console.log('    ' + n));
  console.log('  exercises (' + target.exercises.length + '):');
  target.exercises.forEach(ex => {
    console.log('    ' + ex.name + ': ' + ex.entries.map(formatEntry).join(', '));
  });
  if (target.exercises.length > 0) {
    console.log('  notes that would be sent:');
    console.log(buildNotes(target.exercises));
    console.log('  displayName that would be sent: ' + buildDisplayName(target.exercises));
  }
  return target;
}

function showDataPointByName(name) {
  if (!name) throw new Error('showDataPointByName: pass a resource name (e.g. one from the Health IDs column).');
  const dp = getDataPointByName(name);
  console.log(JSON.stringify(dp, null, 2));
  return dp;
}
