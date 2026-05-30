// Probe whether POST or PUT directed at an existing datapoint can update it
// in place, as an alternative to the delete+POST cycle our sync uses for
// own datapoints (and as the only theoretical update path for foreign
// datapoints, since DELETE on foreign is blocked at the API).
//
// Targets the most-recent sync-created exercise (safe; we own it) and the
// most-recent foreign exercise (typically a Fitbit watch session). For
// each, runs four method/body combinations and re-GETs after every attempt
// to see whether `notes` actually changed. Attempts to restore notes via
// the same mechanism if it appeared to work, so a successful mutation
// doesn't permanently rewrite a Fitbit-sourced note.
//
// WARNING: this *attempts* to mutate foreign data. If any variant succeeds
// and the restore step fails, that Fitbit session's notes will be modified.
// Original notes are logged before any attempt so manual recovery is
// possible from the Executions log.
//
// No args. Run from the Apps Script editor.
function probeOverwrite() {
  const { rows } = readRows();

  let syncCreatedName = null;
  for (let i = rows.length - 1; i >= 0; i--) {
    const ids = splitHealthIdsByType_(rows[i].healthIds).exercise;
    if (ids.length > 0) { syncCreatedName = ids[0]; break; }
  }

  const ourIds = {};
  rows.forEach(r => {
    splitHealthIdsByType_(r.healthIds).exercise.forEach(n => { ourIds[n] = true; });
  });

  let foreignPoint = null;
  const today = new Date();
  for (let i = 0; i < 30 && !foreignPoint; i++) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - i);
    let points;
    try {
      points = listExercisesOnDate(d);
    } catch (err) {
      console.warn('probeOverwrite: listExercisesOnDate failed for ' + ymd(d) + ': ' + err);
      continue;
    }
    points.sort((a, b) => {
      const aS = new Date((a.exercise && a.exercise.interval && a.exercise.interval.startTime) || 0).getTime();
      const bS = new Date((b.exercise && b.exercise.interval && b.exercise.interval.startTime) || 0).getTime();
      return bS - aS;
    });
    for (const p of points) {
      if (!p.name || ourIds[p.name]) continue;
      foreignPoint = p;
      break;
    }
  }

  console.log('=== A. OWN datapoint ===');
  if (syncCreatedName) {
    probeOverwriteOne_(syncCreatedName);
  } else {
    console.log('No sync-created exercise datapoint found.');
  }

  console.log('\n=== B. FOREIGN datapoint ===');
  if (foreignPoint) {
    const platform = (foreignPoint.dataSource && foreignPoint.dataSource.platform) || '<unknown>';
    console.log('Selected ' + foreignPoint.name + ' (platform=' + platform + ')');
    probeOverwriteOne_(foreignPoint.name);
  } else {
    console.log('No foreign exercise datapoint found in last 30 days.');
  }
}

function probeOverwriteOne_(name) {
  const meName = name.replace(/^users\/[^/]+\//, 'users/me/');
  const url = HEALTH_API_BASE + '/' + meName;

  console.log('--- GET original ---');
  const before = getRaw_(name);
  console.log('HTTP ' + before.code);
  if (before.code < 200 || before.code >= 300) {
    console.log(before.body);
    return;
  }
  const orig = JSON.parse(before.body);
  const origNotes = (orig.exercise && orig.exercise.notes) || '';
  console.log('Original notes: ' + JSON.stringify(origNotes));

  const probeNotes = (origNotes || '<no notes>') + ' [PROBE-OVERWRITE]';
  const variants = [
    { label: 'V1: POST to resource URL, minimal body',
      method: 'POST',
      body: { name: meName, exercise: { notes: probeNotes } } },
    { label: 'V2: POST to resource URL, full body',
      method: 'POST',
      body: fullBodyWithNotes_(orig, probeNotes, meName) },
    { label: 'V3: PUT to resource URL, minimal body',
      method: 'PUT',
      body: { name: meName, exercise: { notes: probeNotes } } },
    { label: 'V4: PUT to resource URL, full body',
      method: 'PUT',
      body: fullBodyWithNotes_(orig, probeNotes, meName) }
  ];

  for (const v of variants) {
    console.log('\n--- ' + v.label + ' ---');
    const resp = UrlFetchApp.fetch(url, {
      method: v.method,
      contentType: 'application/json',
      headers: authHeaders_(),
      payload: JSON.stringify(v.body),
      muteHttpExceptions: true
    });
    console.log('HTTP ' + resp.getResponseCode());
    console.log(resp.getContentText());

    const after = getRaw_(name);
    if (after.code < 200 || after.code >= 300) {
      console.log('Re-fetch failed: HTTP ' + after.code + ' ' + after.body);
      continue;
    }
    const post = JSON.parse(after.body);
    const postNotes = (post.exercise && post.exercise.notes) || '';
    const changed = postNotes !== origNotes;
    console.log('Notes changed? ' + changed);

    if (changed) {
      console.log('  After notes: ' + JSON.stringify(postNotes));
      console.log('  Attempting restore via same method...');
      const restoreBody = (v.body.exercise && 'notes' in v.body.exercise)
        ? Object.assign({}, v.body, {
            exercise: Object.assign({}, v.body.exercise, { notes: origNotes })
          })
        : null;
      if (!restoreBody) {
        console.error('  Could not build restore body; restore manually using the original-notes value above.');
        return;
      }
      const restoreResp = UrlFetchApp.fetch(url, {
        method: v.method,
        contentType: 'application/json',
        headers: authHeaders_(),
        payload: JSON.stringify(restoreBody),
        muteHttpExceptions: true
      });
      console.log('  Restore HTTP ' + restoreResp.getResponseCode());
      const verify = getRaw_(name);
      const verifyNotes = (JSON.parse(verify.body).exercise || {}).notes || '';
      console.log('  Verified notes after restore: ' + JSON.stringify(verifyNotes));
      if (verifyNotes !== origNotes) {
        console.error('  RESTORE FAILED. Notes are: ' + JSON.stringify(verifyNotes));
        console.error('  Original was: ' + JSON.stringify(origNotes));
        return;
      }
    }
  }
}

// Find sync-created exercise datapoints (platform=GOOGLE_WEB_API plus
// matching googleWebClientId) that aren't referenced by any sheet row's
// Created Health IDs, and delete them. Catches the recreated
// datapoint that the round-trip probe leaves behind, plus any other
// orphans (rows deleted from the sheet without Force Resync first).
//
// Scans the last 30 days. Run when no sync is in flight; does not take
// the script lock.
function cleanupProbeRemnant() {
  const ourClientId = PropertiesService.getScriptProperties().getProperty(HEALTH_OAUTH_CLIENT_ID_KEY);
  if (!ourClientId) {
    console.error('cleanupProbeRemnant: ' + HEALTH_OAUTH_CLIENT_ID_KEY
      + ' script property not set; cannot identify which datapoints are ours.');
    return;
  }

  const { rows } = readRows();
  const trackedIds = {};
  rows.forEach(r => {
    splitHealthIdsByType_(r.healthIds).exercise.forEach(n => { trackedIds[n] = true; });
  });

  const today = new Date();
  const LOOKBACK_DAYS = 30;
  const orphans = [];
  for (let i = 0; i < LOOKBACK_DAYS; i++) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() - i);
    let points;
    try {
      points = listExercisesOnDate(d);
    } catch (err) {
      console.warn('cleanupProbeRemnant: listExercisesOnDate failed for ' + ymd(d) + ': ' + err);
      continue;
    }
    points.forEach(p => {
      if (!p.name || trackedIds[p.name]) return;
      const ds = p.dataSource || {};
      const app = ds.application || {};
      if (ds.platform !== 'GOOGLE_WEB_API') return;
      if (app.googleWebClientId !== ourClientId) return;
      orphans.push({
        name: p.name,
        date: ymd(d),
        startTime: (p.exercise && p.exercise.interval && p.exercise.interval.startTime) || '<unknown>',
        notes: (p.exercise && p.exercise.notes) || ''
      });
    });
  }

  if (orphans.length === 0) {
    console.log('No orphaned sync-created exercise datapoints found in last '
      + LOOKBACK_DAYS + ' days.');
    return;
  }

  console.log('Found ' + orphans.length + ' orphan(s):');
  orphans.forEach(o => {
    console.log('  ' + o.date + ' ' + o.startTime + ' ' + o.name);
    if (o.notes) console.log('    notes: ' + o.notes.split('\n').join(' | '));
  });

  console.log('Deleting...');
  try {
    deleteDataPointsByName(orphans.map(o => o.name));
    console.log('Deleted ' + orphans.length + ' datapoint(s).');
  } catch (err) {
    console.error('Delete failed: ' + err);
  }
}

function fullBodyWithNotes_(orig, newNotes, meName) {
  const ex = Object.assign({}, orig.exercise || {});
  delete ex.createTime;
  delete ex.updateTime;
  delete ex.metricsSummary;
  ex.notes = newNotes;
  return {
    name: meName,
    dataSource: orig.dataSource,
    exercise: ex
  };
}

function getRaw_(name) {
  const meName = name.replace(/^users\/[^/]+\//, 'users/me/');
  const resp = UrlFetchApp.fetch(HEALTH_API_BASE + '/' + meName, {
    method: 'GET',
    headers: authHeaders_(),
    muteHttpExceptions: true
  });
  return { code: resp.getResponseCode(), body: resp.getContentText() };
}
