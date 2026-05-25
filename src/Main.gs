function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sync')
    .addItem('Run now', 'runSyncNow')
    .addItem('Force resync current row', 'forceResyncCurrentRow')
    .addSeparator()
    .addItem('Authorize Health API', 'authorizeHealthApi')
    .addItem('Revoke Health API authorization', 'revokeHealthApi')
    .addSeparator()
    .addItem('Run tests', 'runParserTests')
    .addItem('Re-install triggers', 'installTriggers')
    .addToUi();
}

function authorizeHealthApi() {
  const ui = SpreadsheetApp.getUi();
  let service;
  try {
    service = getHealthService();
  } catch (err) {
    ui.alert('Setup needed', String(err.message || err), ui.ButtonSet.OK);
    return;
  }
  if (service.hasAccess()) {
    ui.alert('Already authorized', 'Google Health API access is active. Use "Revoke" first if you want to re-authorize.', ui.ButtonSet.OK);
    return;
  }
  const url = service.getAuthorizationUrl();
  const html = '<p>Click the link below, sign in with the Google account that owns your Google Health data, and grant the requested scopes.</p>'
    + '<p><a href="' + url + '" target="_blank" rel="noopener">Authorize Google Health API</a></p>'
    + '<p>After the success page appears, you can close that tab and return here.</p>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(220), 'Authorize Google Health');
}

function revokeHealthApi() {
  resetHealthAuth();
  SpreadsheetApp.getUi().alert('Google Health authorization cleared. Use "Authorize Health API" to grant again.');
}

function setup() {
  ensureManagedColumns();
  installTriggers();
}

function installTriggers() {
  const handlers = new Set(['onEditTrigger', 'debounceFlush', 'backstop']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('debounceFlush').timeBased().everyMinutes(DEBOUNCE_CHECK_INTERVAL_MIN).create();
  ScriptApp.newTrigger('backstop').timeBased().everyHours(BACKSTOP_INTERVAL_HOURS).create();
}

function onEditTrigger(e) {
  try {
    onEditMarkDirty(e);
  } catch (err) {
    console.error('onEditTrigger error: ' + err);
  }
}

function debounceFlush() {
  syncDirtyRows();
}

function backstop() {
  syncDirtyRows();
}

function runSyncNow() {
  const result = syncDirtyRows();
  const ui = SpreadsheetApp.getUi();
  if (!result) {
    ui.alert('Sync skipped (another run holds the lock). Try again shortly.');
    return;
  }
  let msg = 'Synced ' + result.ok + ' row(s)';
  if (result.errors > 0) msg += ', ' + result.errors + ' error(s)';
  msg += '.';
  if (result.errors > 0) msg += '\n\nSee Executions for details.';
  ui.alert(msg);
}

function forceResyncCurrentRow() {
  const sheet = getSheet_();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return;
  const { map } = getHeaderMap_(sheet);
  const col = map[SYNCED_AT_COLUMN_HEADER];
  if (!col) {
    SpreadsheetApp.getUi().alert('Synced At column missing. Run setup.');
    return;
  }
  clearRowSynced(row, col);
  syncDirtyRows();
}

function syncDirtyRows() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    console.warn('syncDirtyRows: another run holds the lock; skipping.');
    return null;
  }
  let ok = 0;
  let errors = 0;
  try {
    const { rows, syncedAtCol, healthIdsCol } = readRows();
    if (!syncedAtCol || !healthIdsCol) {
      console.error('syncDirtyRows: managed columns missing; run setup().');
      return { ok: 0, errors: 1 };
    }
    const dirty = rows.filter(r => !r.syncedAt);
    if (dirty.length === 0) return { ok: 0, errors: 0 };

    const now = Date.now();
    const ready = [];
    const waiting = [];
    dirty.forEach(r => {
      if (!r.lastEditedAt) {
        ready.push(r);
        return;
      }
      const sinceMs = now - r.lastEditedAt.getTime();
      if (sinceMs >= LAST_EDIT_QUIESCE_MS) ready.push(r);
      else waiting.push({ row: r, sinceMs: sinceMs });
    });

    if (waiting.length > 0) {
      console.info('syncDirtyRows: ' + waiting.length + ' row(s) still in quiesce window (need '
        + LAST_EDIT_QUIESCE_MS + 'ms since lastEditedAt); will retry next pass.');
    }
    if (ready.length === 0) {
      console.info('syncDirtyRows: no rows ready to sync.');
      return { ok: 0, errors: 0 };
    }

    console.info('syncDirtyRows: ' + ready.length + ' ready row(s)');

    const byDate = {};
    ready.forEach(r => {
      const key = ymd(r.date);
      (byDate[key] = byDate[key] || []).push(r);
    });

    let done = 0;
    Object.keys(byDate).forEach(dateKey => {
      const dateRows = byDate[dateKey];
      dateRows.sort((a, b) => a.rowNum - b.rowNum);
      const allRowsForDate = rows.filter(row => ymd(row.date) === dateKey);
      allRowsForDate.sort((a, b) => a.rowNum - b.rowNum);
      for (let i = 0; i < dateRows.length; i++) {
        done++;
        const ordinal = allRowsForDate.indexOf(dateRows[i]);
        if (syncOneRow_(dateRows[i], ordinal, syncedAtCol, healthIdsCol, done, ready.length)) ok++;
        else errors++;
      }
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: ok, errors: errors };
}

// Resolve the exercise interval (and weight sampleTime) for a row.
// Prefers edit-derived timing when First/Last Edited At are both present.
// Falls back to synthetic noon-ordinal otherwise.
function resolveRowTiming_(row, ordinal) {
  if (row.firstEditedAt && row.lastEditedAt) {
    const startMs = row.firstEditedAt.getTime();
    const endMs = row.lastEditedAt.getTime();
    const tz = Session.getScriptTimeZone();
    const startOffset = getTzOffsetSeconds_(tz, row.firstEditedAt);
    const endOffset = getTzOffsetSeconds_(tz, row.lastEditedAt);
    return {
      source: 'edit',
      exercise: {
        startUtcMs: startMs,
        startOffsetSeconds: startOffset,
        endUtcMs: endMs,
        endOffsetSeconds: endOffset
      },
      weight: { utcMs: startMs, offsetSeconds: startOffset }
    };
  }
  const ex = syntheticExerciseInterval_(row.date, ordinal);
  const wt = syntheticWeightSample_(row.date);
  return { source: 'synthetic', exercise: ex, weight: wt };
}

function syncOneRow_(row, ordinal, syncedAtCol, healthIdsCol, doneIdx, total) {
  const dateKey = ymd(row.date);
  const tag = '[' + doneIdx + '/' + total + '] ' + dateKey + ' row ' + row.rowNum;
  console.info(tag + ': starting (exercises=' + row.exercises.length
    + ', bodyweight=' + (row.bodyweight === null ? 'none' : row.bodyweight)
    + ', oldIds=' + row.healthIds.length + ')');

  if (row.healthIds.length > 0) {
    console.info(tag + ': deleting ' + row.healthIds.length + ' previous datapoint(s)');
    deleteDataPointsByName(row.healthIds);
  }

  let timing;
  try {
    timing = resolveRowTiming_(row, ordinal);
  } catch (err) {
    console.error(tag + ': resolveRowTiming_ failed: ' + err);
    return false;
  }
  console.info(tag + ': timing source=' + timing.source);

  const newIds = [];
  let failed = false;

  if (SYNC_EXERCISES && row.exercises.length > 0) {
    let ex = timing.exercise;
    if (timing.source === 'edit') {
      try {
        const matches = findForeignOverlappingExercises(ex.startUtcMs, ex.endUtcMs);
        if (matches.length > 0) {
          const m = matches[0];
          console.info(tag + ': adopting foreign exercise ' + m.name
            + ' (overlap=' + m.overlapMs + 'ms); deleting and recreating with our content');
          deleteDataPointsByName([m.name]);
          ex = {
            startUtcMs: m.startUtcMs,
            startOffsetSeconds: m.startUtcOffsetSeconds,
            endUtcMs: m.endUtcMs,
            endOffsetSeconds: m.endUtcOffsetSeconds
          };
        }
      } catch (err) {
        console.warn(tag + ': foreign-match lookup failed; using edit-derived times. ' + err);
      }
    }
    try {
      const notes = buildNotes(row.exercises);
      const displayName = buildDisplayName(row.exercises);
      const name = createExerciseAt(ex.startUtcMs, ex.startOffsetSeconds,
        ex.endUtcMs, ex.endOffsetSeconds, notes, displayName);
      if (name) newIds.push(name);
      console.info(tag + ': createExerciseAt -> ' + (name || '<no name>'));
    } catch (err) {
      console.error(tag + ': createExerciseAt failed: ' + err);
      failed = true;
    }
  }

  if (SYNC_WEIGHT && row.bodyweight !== null) {
    try {
      const wt = timing.weight;
      const name = createWeightAt(wt.utcMs, wt.offsetSeconds, row.bodyweight);
      if (name) newIds.push(name);
      console.info(tag + ': createWeightAt(' + row.bodyweight + ' lb) -> ' + (name || '<no name>'));
    } catch (err) {
      console.error(tag + ': createWeightAt failed: ' + err);
      failed = true;
    }
  }

  writeHealthIds(row.rowNum, healthIdsCol, newIds);
  if (failed) {
    console.warn(tag + ': FAILED (partial); will retry on next sync.');
    return false;
  }
  markRowSynced(row.rowNum, syncedAtCol, new Date().toISOString());
  console.info(tag + ': done');
  return true;
}
