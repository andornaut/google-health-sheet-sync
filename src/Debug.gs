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
//   dumpLatestMatchedSession()    -> newest row's Matched Health Session
//   dumpLatestCreatedId()         -> newest row's first Created Health ID
//   dumpTodaysForeignStrength()   -> today's foreign Strength sessions
//   findExerciseWithMetadata()    -> first exercise in last 90 days whose
//                                    exerciseMetadata is non-empty (so we
//                                    can copy the structured shape)
//   patchLatestExerciseNotes()    -> PATCH notes on the newest row's
//                                    sync-created exercise datapoint with
//                                    freshly-built notes, then re-fetch to
//                                    confirm the update took

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

// Scan the last 90 days for any exercise datapoint with a non-empty
// exerciseMetadata field, so we can copy the structured shape Google Fit /
// Health Connect apps use for per-segment workout content. Logs the first
// hit in full and returns it; logs nothing-found if all metadata is empty.
function findExerciseWithMetadata() {
  const today = new Date();
  for (let i = 0; i < 90; i++) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - i);
    const points = listExercisesOnDate(d);
    for (const p of points) {
      const meta = p.exercise && p.exercise.exerciseMetadata;
      if (meta && Object.keys(meta).length > 0) {
        console.log('Found populated exerciseMetadata on ' + ymd(d) + ': ' + p.name);
        console.log(JSON.stringify(p, null, 2));
        return p;
      }
    }
  }
  console.log('No exercise datapoint in the last 90 days has a populated exerciseMetadata.');
  return null;
}

// Single-row test of patchExerciseNotes: finds the newest row that has both
// exercise content and a sync-created exercise datapoint, rebuilds notes
// from the current sheet content, PATCHes them onto the existing datapoint,
// and re-fetches to confirm the server stored the new text. Does NOT touch
// any sheet column.
function patchLatestExerciseNotes() {
  const { rows } = readRows();
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (r.exercises.length === 0) continue;
    const exerciseIds = splitHealthIdsByType_(r.healthIds).exercise;
    if (exerciseIds.length === 0) continue;
    const name = exerciseIds[0];
    const notes = buildNotes(r.exercises);
    console.log('Row ' + r.rowNum + ' (' + ymd(r.date) + ') patching ' + name);
    console.log('--- New notes ---');
    console.log(notes);
    console.log('--- PATCH response ---');
    const resp = patchExerciseNotes(name, notes);
    console.log(JSON.stringify(resp, null, 2));
    // Re-fetch repeatedly to test the eventual-consistency theory: the
    // server may return stale notes immediately after PATCH and only reflect
    // the new text seconds later.
    const delaysSec = [0, 3, 10, 30];
    for (const sec of delaysSec) {
      if (sec > 0) Utilities.sleep(sec * 1000);
      console.log('--- Re-fetch after +' + sec + 's ---');
      const point = dumpDataPoint_(name);
      const stored = point && point.exercise && point.exercise.notes;
      console.log('notes match sent? ' + (stored === notes));
      if (stored === notes) return point;
    }
    console.log('Notes still stale after ' + delaysSec[delaysSec.length - 1] + 's.');
    return null;
  }
  console.log('No row has both exercise content and a sync-created exercise ID.');
  return null;
}

// Diagnostic: pick the newest sync-created weight datapoint, PATCH its
// weightGrams to a known-different value, re-fetch to see if the change
// stuck, then restore the original value. Mirrors the Go client library's
// dataPoints.patch test (which only covers weight). Tells us whether the
// PATCH endpoint is fundamentally functional from Apps Script, isolating
// whether the exercise-notes failure is exercise-specific or universal.
function patchLatestWeightTest() {
  const { rows } = readRows();
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    const weightIds = splitHealthIdsByType_(r.healthIds).weight;
    if (weightIds.length === 0) continue;
    const name = weightIds[0];
    const meName = name.replace(/^users\/[^/]+\//, 'users/me/');
    const url = HEALTH_API_BASE + '/' + meName;

    const before = getDataPointByName(name);
    const beforeGrams = before.weight && before.weight.weightGrams;
    console.log('Row ' + r.rowNum + ' (' + ymd(r.date) + ') patching weight ' + name);
    console.log('Before weightGrams: ' + beforeGrams);

    const newGrams = Number(beforeGrams) + 100;
    const newWeight = Object.assign({}, before.weight, { weightGrams: newGrams });
    delete newWeight.createTime;
    delete newWeight.updateTime;
    const patchBody = { name: meName, weight: newWeight };
    console.log('--- PATCH body ---');
    console.log(JSON.stringify(patchBody, null, 2));
    const resp = httpJson_('PATCH', url, patchBody);
    console.log('--- PATCH response ---');
    console.log(JSON.stringify(resp, null, 2));

    const after = getDataPointByName(name);
    const afterGrams = after.weight && after.weight.weightGrams;
    console.log('After weightGrams: ' + afterGrams);
    console.log('Changed? ' + (Number(afterGrams) !== Number(beforeGrams)));

    // Best-effort restore so we don't leave the datapoint with the +100 nudge.
    try {
      const restoreWeight = Object.assign({}, before.weight, { weightGrams: beforeGrams });
      delete restoreWeight.createTime;
      delete restoreWeight.updateTime;
      httpJson_('PATCH', url, { name: meName, weight: restoreWeight });
      const restored = getDataPointByName(name);
      console.log('Restored weightGrams: ' + (restored.weight && restored.weight.weightGrams));
    } catch (err) {
      console.warn('Restore PATCH failed (datapoint may be left at ' + newGrams + 'g): ' + err);
    }
    return { name: name, beforeGrams: beforeGrams, afterGrams: afterGrams };
  }
  console.log('No row has a sync-created weight datapoint.');
  return null;
}

// Probe whether any exercise field is mutable via PATCH. We've confirmed
// weight PATCH works (patchLatestWeightTest) but exercise notes PATCH
// doesn't. This tries several body shapes against displayName,
// activeDuration, and notes to narrow down whether the failure is
// notes-specific, exercise-wide, or body-shape-sensitive. Bypasses
// httpJson_'s retry loop since 500s here are consistent (not transient).
// Restores original notes/displayName/activeDuration at the end.
function probeExercisePatch() {
  const { rows } = readRows();
  let target = null;
  for (let i = rows.length - 1; i >= 0; i--) {
    const ids = splitHealthIdsByType_(rows[i].healthIds).exercise;
    if (ids.length > 0) {
      target = { row: rows[i], name: ids[0] };
      break;
    }
  }
  if (!target) {
    console.log('No sync-created exercise datapoint found.');
    return null;
  }

  const meName = target.name.replace(/^users\/[^/]+\//, 'users/me/');
  const url = HEALTH_API_BASE + '/' + meName;
  const original = getDataPointByName(target.name);
  const orig = original.exercise || {};
  const origNotes = orig.notes;
  const origDisplay = orig.displayName;
  const origDuration = orig.activeDuration;
  console.log('Probing ' + target.name);
  console.log('Original notes:         ' + origNotes);
  console.log('Original displayName:   ' + origDisplay);
  console.log('Original activeDuration:' + origDuration);

  function fullExercise_(overrides) {
    const ex = Object.assign({}, orig, overrides);
    delete ex.createTime;
    delete ex.updateTime;
    return ex;
  }

  const variants = [
    { label: 'V1 minimal: displayName only',
      body: { name: meName, exercise: { displayName: 'PROBE_DISPLAY_V1' } } },
    { label: 'V2 minimal: activeDuration only',
      body: { name: meName, exercise: { activeDuration: '1234s' } } },
    { label: 'V3 minimal: notes only',
      body: { name: meName, exercise: { notes: 'PROBE_NOTES_V3' } } },
    { label: 'V4 full: notes swapped, no dataSource',
      body: { name: meName, exercise: fullExercise_({ notes: 'PROBE_NOTES_V4' }) } },
    { label: 'V5 full: notes swapped, with dataSource',
      body: { name: meName, dataSource: original.dataSource, exercise: fullExercise_({ notes: 'PROBE_NOTES_V5' }) } },
    { label: 'V6 full: displayName swapped, with dataSource',
      body: { name: meName, dataSource: original.dataSource, exercise: fullExercise_({ displayName: 'PROBE_DISPLAY_V6' }) } }
  ];

  for (const v of variants) {
    console.log('\n=== ' + v.label + ' ===');
    const resp = UrlFetchApp.fetch(url, {
      method: 'PATCH',
      contentType: 'application/json',
      headers: authHeaders_(),
      payload: JSON.stringify(v.body),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    console.log('HTTP ' + code);
    if (code >= 300) {
      console.log('Error body: ' + text);
      continue;
    }
    const after = getDataPointByName(target.name);
    const aex = after.exercise || {};
    console.log('  after.notes:          ' + aex.notes);
    console.log('  after.displayName:    ' + aex.displayName);
    console.log('  after.activeDuration: ' + aex.activeDuration);
    console.log('  notes changed?           ' + (aex.notes !== origNotes));
    console.log('  displayName changed?     ' + (aex.displayName !== origDisplay));
    console.log('  activeDuration changed?  ' + (aex.activeDuration !== origDuration));
  }

  console.log('\n=== Restoring original values ===');
  try {
    const restore = { name: meName, dataSource: original.dataSource,
      exercise: fullExercise_({ notes: origNotes, displayName: origDisplay, activeDuration: origDuration }) };
    httpJson_('PATCH', url, restore);
    const restored = getDataPointByName(target.name);
    const rex = restored.exercise || {};
    console.log('  notes:          ' + rex.notes);
    console.log('  displayName:    ' + rex.displayName);
    console.log('  activeDuration: ' + rex.activeDuration);
  } catch (err) {
    console.warn('Restore PATCH failed; datapoint may be left in probe state: ' + err);
  }
  return null;
}

function dumpDataPoint_(name) {
  const point = getDataPointByName(name);
  console.log(JSON.stringify(point, null, 2));
  return point;
}
