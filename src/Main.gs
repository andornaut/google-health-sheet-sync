function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sync')
    .addItem('Run now', 'runSyncNow')
    .addItem('Force resync current row', 'forceResyncCurrentRow')
    .addItem('Force resync ALL rows', 'forceResyncAllRows')
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
  const handlers = new Set(['onEditTrigger', 'flushIfPending', 'backstop']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('flushIfPending').timeBased().everyMinutes(POLL_INTERVAL_MIN).create();
  ScriptApp.newTrigger('backstop').timeBased().everyHours(BACKSTOP_INTERVAL_HOURS).create();
}

function onEditTrigger(e) {
  try {
    onEditMarkDirty(e);
  } catch (err) {
    console.error('onEditTrigger error: ' + err);
  }
}

function flushIfPending() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PENDING_DIRTY_KEY) !== '1') {
    return;
  }
  syncDirtyRows(false, 0);
}

function backstop() {
  syncDirtyRows(false, 0);
}

// "Run now" is an explicit manual action: bypasses the quiesce window so the
// user sees results immediately. If they keep editing afterward, the row goes
// dirty again and the next sync replaces the Health datapoint.
function runSyncNow() {
  const result = syncDirtyRows(true, LOCK_WAIT_MS);
  SpreadsheetApp.getUi().alert(formatSyncResult_(result, 'Synced'));
}

function forceResyncCurrentRow() {
  const sheet = getSheet_();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return;
  const ui = SpreadsheetApp.getUi();
  const { map } = getHeaderMap_(sheet);
  const col = map[SYNCED_AT_COLUMN_HEADER];
  if (!col) {
    ui.alert('Synced At column missing. Run setup.');
    return;
  }
  clearRowSynced(row, col);
  SpreadsheetApp.flush();
  PropertiesService.getScriptProperties().setProperty(PENDING_DIRTY_KEY, '1');
  const result = syncDirtyRows(true, LOCK_WAIT_MS);
  ui.alert(formatSyncResult_(result, 'Resynced'));
}

function forceResyncAllRows() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  const col = map[SYNCED_AT_COLUMN_HEADER];
  if (!col) {
    ui.alert('Synced At column missing. Run setup.');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data rows.');
    return;
  }
  const dataRowCount = lastRow - 1;
  const confirm = ui.alert(
    'Force resync ALL rows?',
    'This will clear "Synced At" for ' + dataRowCount + ' row(s) and re-upload every row to Google Health '
      + '(delete + recreate). The quiesce window is bypassed, so rows edited recently will resync too. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const blanks = [];
  for (let i = 0; i < dataRowCount; i++) blanks.push(['']);
  sheet.getRange(2, col, dataRowCount, 1).setValues(blanks);
  SpreadsheetApp.flush();
  PropertiesService.getScriptProperties().setProperty(PENDING_DIRTY_KEY, '1');

  const result = syncDirtyRows(true, LOCK_WAIT_MS);
  ui.alert(formatSyncResult_(result, 'Resynced'));
}

function formatSyncResult_(result, verb) {
  if (!result) return 'Sync skipped (another run holds the lock). Try again shortly.';
  let msg = verb + ' ' + result.ok + ' row(s)';
  if (result.errors > 0) msg += ', ' + result.errors + ' error(s)';
  msg += '.';
  if (result.errors > 0) msg += '\n\nSee Executions for details.';
  return msg;
}

function syncDirtyRows(bypassQuiesce, lockWaitMs) {
  const lock = LockService.getScriptLock();
  const waitMs = (lockWaitMs === undefined || lockWaitMs === null) ? LOCK_WAIT_MS : lockWaitMs;
  if (!lock.tryLock(waitMs)) {
    if (waitMs > 0) {
      console.warn('syncDirtyRows: another run holds the lock; skipping.');
    } else {
      console.info('syncDirtyRows: another run holds the lock; skipping this tick.');
    }
    return null;
  }
  const props = PropertiesService.getScriptProperties();
  let ok = 0;
  let errors = 0;
  let waitingCount = 0;
  try {
    // Clear early so any onEditMarkDirty that fires during this pass can
    // re-set the flag and be picked up by the next poll without ambiguity.
    props.deleteProperty(PENDING_DIRTY_KEY);

    const { rows, syncedAtCol, healthIdsCol, lastEditedAtCol } = readRows();
    if (!syncedAtCol || !healthIdsCol) {
      console.error('syncDirtyRows: managed columns missing; run setup().');
      errors = 1;
      return { ok: 0, errors: errors };
    }
    const dirty = rows.filter(r => !r.syncedAt);
    if (dirty.length === 0) {
      return { ok: 0, errors: 0 };
    }

    const ordinalByRowNum = buildOrdinalMap_(rows);

    const now = Date.now();
    const ready = [];
    dirty.forEach(r => {
      if (bypassQuiesce || !r.lastEditedAt) {
        ready.push(r);
        return;
      }
      const sinceMs = now - r.lastEditedAt.getTime();
      if (sinceMs >= LAST_EDIT_QUIESCE_MS) ready.push(r);
      else waitingCount++;
    });

    if (waitingCount > 0) {
      console.info('syncDirtyRows: ' + waitingCount + ' row(s) still in quiesce window (need '
        + LAST_EDIT_QUIESCE_MS + 'ms since lastEditedAt); will retry next pass.');
    }
    if (ready.length === 0) {
      console.info('syncDirtyRows: no rows ready to sync.');
      return { ok: 0, errors: 0 };
    }
    if (bypassQuiesce) console.info('syncDirtyRows: bypassQuiesce=true');

    console.info('syncDirtyRows: ' + ready.length + ' ready row(s)');

    ready.sort((a, b) => a.rowNum - b.rowNum);
    for (let i = 0; i < ready.length; i++) {
      const r = ready[i];
      const ordinal = ordinalByRowNum[r.rowNum];
      if (syncOneRow_(r, ordinal, syncedAtCol, healthIdsCol, lastEditedAtCol, i + 1, ready.length)) ok++;
      else errors++;
    }
  } finally {
    // Re-set the flag if work remains: quiescing rows need a future poll to
    // pick them up, and failed rows should be retried promptly. (syncOneRow_
    // also sets the flag itself when it defers a row due to a concurrent
    // edit, so even rows counted as ok can leave the flag set.)
    if (waitingCount > 0 || errors > 0) {
      props.setProperty(PENDING_DIRTY_KEY, '1');
    }
    lock.releaseLock();
  }
  return { ok: ok, errors: errors };
}

// Group all rows by their civil date and assign each row an ordinal equal to
// its rank (by rowNum) within that date. Used by the synthetic-timing fallback
// to give each row on the same date a distinct startHour offset.
function buildOrdinalMap_(rows) {
  const byDate = {};
  rows.forEach(r => {
    const key = ymd(r.date);
    (byDate[key] = byDate[key] || []).push(r);
  });
  const ordinalByRowNum = {};
  Object.keys(byDate).forEach(dateKey => {
    const dateRows = byDate[dateKey];
    dateRows.sort((a, b) => a.rowNum - b.rowNum);
    dateRows.forEach((r, i) => { ordinalByRowNum[r.rowNum] = i; });
  });
  return ordinalByRowNum;
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

function syncOneRow_(row, ordinal, syncedAtCol, healthIdsCol, lastEditedAtCol, doneIdx, total) {
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

  // Concurrent-edit guard: if the user edited this row while we were
  // processing it, Last Edited At in the sheet is newer than what we
  // captured at the start of the pass. Skip stamping Synced At so the row
  // stays dirty; the next sync replaces our just-created datapoint with one
  // that reflects the new content.
  //
  // Two transitions to detect:
  //  - non-null -> different value (the row already had edit timestamps)
  //  - null     -> non-null        (a legacy/backfill row got its first edit
  //                                 while sync was running)
  if (lastEditedAtCol) {
    const currentLastEdit = toDate_(getSheet_().getRange(row.rowNum, lastEditedAtCol).getValue());
    const previousMs = row.lastEditedAt ? row.lastEditedAt.getTime() : null;
    const currentMs = currentLastEdit ? currentLastEdit.getTime() : null;
    if (currentMs !== previousMs) {
      const prevLabel = row.lastEditedAt ? row.lastEditedAt.toISOString() : '<none>';
      const currLabel = currentLastEdit ? currentLastEdit.toISOString() : '<cleared>';
      console.info(tag + ': concurrent edit detected (Last Edited At '
        + prevLabel + ' -> ' + currLabel + '); deferring Synced At stamp, will retry next sync.');
      PropertiesService.getScriptProperties().setProperty(PENDING_DIRTY_KEY, '1');
      return true;
    }
  }

  markRowSynced(row.rowNum, syncedAtCol, new Date().toISOString());
  console.info(tag + ': done');
  return true;
}
