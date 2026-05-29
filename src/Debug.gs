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
  console.log('  exerciseSyncedAt: ' + (target.exerciseSyncedAt || '<dirty>'));
  console.log('  weightSyncedAt: ' + (target.weightSyncedAt || '<dirty>'));
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

// Build (but do not POST) the exact exercise payload that syncOneRow_ would
// send for the given row. Useful for inspecting what we're sending when the
// API returns 5xx so we can spot malformed fields.
function debugBuildExercisePayloadForRow(rowNum) {
  const target = findRow_(rowNum);
  if (!target) return null;
  if (target.exercises.length === 0) {
    console.log('Row ' + target.rowNum + ' has no exercise content; nothing to build.');
    return null;
  }
  const { rows } = readRows();
  const ordinal = buildOrdinalMap_(rows)[target.rowNum];
  const timing = resolveRowTiming_(target, ordinal);
  const ex = timing.exercise;
  const durationSec = Math.max(0, Math.round((ex.endUtcMs - ex.startUtcMs) / 1000));
  const payload = {
    dataSource: { recordingMethod: 'MANUAL' },
    exercise: {
      interval: buildIntervalFromUtc_(ex.startUtcMs, ex.startOffsetSeconds, ex.endUtcMs, ex.endOffsetSeconds),
      exerciseType: 'STRENGTH_TRAINING',
      displayName: buildDisplayName(target.exercises) || 'Strength Training',
      notes: buildNotes(target.exercises),
      activeDuration: durationSec + 's',
      metricsSummary: { caloriesKcal: 0 }
    }
  };
  console.log('Row ' + target.rowNum + '  date: ' + ymd(target.date) + '  timing source: ' + timing.source);
  console.log('POST ' + HEALTH_API_BASE + '/users/me/dataTypes/exercise/dataPoints');
  console.log(JSON.stringify(payload, null, 2));
  return payload;
}

// Actually POST the exercise payload for the given row and log the response
// (or error). Side-effect: creates a new datapoint in Google Health on
// success. Does NOT touch the sheet's Health IDs column, so subsequent
// normal sync runs won't know about the datapoint created here — clean up
// manually via showDataPointByName / Health app if needed.
function debugPostExerciseForRow(rowNum) {
  const payload = debugBuildExercisePayloadForRow(rowNum);
  if (!payload) return null;
  const url = HEALTH_API_BASE + '/users/me/dataTypes/exercise/dataPoints';
  try {
    const resp = httpJson_('POST', url, payload);
    console.log('Response:');
    console.log(JSON.stringify(resp, null, 2));
    const name = extractDataPointName_(resp);
    console.log('Created datapoint name: ' + (name || '<none extracted>'));
    return resp;
  } catch (err) {
    console.error('POST failed: ' + err);
    throw err;
  }
}

// Same idea for weight.
function debugBuildWeightPayloadForRow(rowNum) {
  const target = findRow_(rowNum);
  if (!target) return null;
  if (target.bodyweight === null) {
    console.log('Row ' + target.rowNum + ' has no bodyweight; nothing to build.');
    return null;
  }
  const { rows } = readRows();
  const ordinal = buildOrdinalMap_(rows)[target.rowNum];
  const timing = resolveRowTiming_(target, ordinal);
  const wt = timing.weight;
  const grams = Math.round(target.bodyweight * GRAMS_PER_LB);
  const payload = {
    dataSource: { recordingMethod: 'MANUAL' },
    weight: {
      weightGrams: grams,
      sampleTime: buildSampleTimeFromUtc_(wt.utcMs, wt.offsetSeconds)
    }
  };
  console.log('Row ' + target.rowNum + '  date: ' + ymd(target.date) + '  timing source: ' + timing.source);
  console.log('POST ' + HEALTH_API_BASE + '/users/me/dataTypes/weight/dataPoints');
  console.log(JSON.stringify(payload, null, 2));
  return payload;
}

function debugPostWeightForRow(rowNum) {
  const payload = debugBuildWeightPayloadForRow(rowNum);
  if (!payload) return null;
  const url = HEALTH_API_BASE + '/users/me/dataTypes/weight/dataPoints';
  try {
    const resp = httpJson_('POST', url, payload);
    console.log('Response:');
    console.log(JSON.stringify(resp, null, 2));
    const name = extractDataPointName_(resp);
    console.log('Created datapoint name: ' + (name || '<none extracted>'));
    return resp;
  } catch (err) {
    console.error('POST failed: ' + err);
    throw err;
  }
}

// Resolve the row to operate on:
//   1. explicit arg (e.g. findRow_(45)), OR
//   2. script property DEBUG_ROW_NUM (set in Project Settings -> Script
//      properties), OR
//   3. last data row in the sheet.
// Lets you click "Run" in the editor without editing source: set the script
// property once and re-run as needed.
function findRow_(rowNum) {
  const { rows } = readRows();
  if (rows.length === 0) {
    console.log('Sheet has no data rows.');
    return null;
  }
  let effective = rowNum;
  if (!effective) {
    const fromProp = PropertiesService.getScriptProperties().getProperty('DEBUG_ROW_NUM');
    if (fromProp) effective = Number(fromProp);
  }
  if (!effective) {
    const last = rows[rows.length - 1];
    console.log('No rowNum/DEBUG_ROW_NUM provided; using last data row ' + last.rowNum);
    return last;
  }
  const target = rows.find(r => r.rowNum === Number(effective));
  if (!target) {
    console.log('Row ' + effective + ' not found or has no date.');
    return null;
  }
  return target;
}
