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
    PropertiesService.getScriptProperties().setProperty(LAST_EDIT_MS_KEY, String(Date.now()));
  } catch (err) {
    console.error('onEditTrigger error: ' + err);
  }
}

function debounceFlush() {
  const props = PropertiesService.getScriptProperties();
  const lastEditMs = Number(props.getProperty(LAST_EDIT_MS_KEY) || 0);
  if (!lastEditMs) {
    console.info('debounceFlush: no pending edits; skipping.');
    return;
  }
  const sinceMs = Date.now() - lastEditMs;
  if (sinceMs < DEBOUNCE_MS) {
    console.info('debounceFlush: last edit ' + sinceMs + 'ms ago (< ' + DEBOUNCE_MS + 'ms); waiting.');
    return;
  }
  console.info('debounceFlush: last edit ' + sinceMs + 'ms ago; flushing.');
  props.deleteProperty(LAST_EDIT_MS_KEY);
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

    console.info('syncDirtyRows: ' + dirty.length + ' dirty row(s)');

    const byDate = {};
    dirty.forEach(r => {
      const key = ymd(r.date);
      (byDate[key] = byDate[key] || []).push(r);
    });

    let done = 0;
    Object.keys(byDate).forEach(dateKey => {
      const dateRows = byDate[dateKey];
      dateRows.sort((a, b) => a.rowNum - b.rowNum);
      for (let i = 0; i < dateRows.length; i++) {
        done++;
        if (syncOneRow_(dateRows[i], i, syncedAtCol, healthIdsCol, done, dirty.length)) ok++;
        else errors++;
      }
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: ok, errors: errors };
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

  const newIds = [];
  let failed = false;

  if (SYNC_EXERCISES && row.exercises.length > 0) {
    try {
      const notes = buildNotes(row.exercises);
      const displayName = buildDisplayName(row.exercises);
      const name = createExercise(row.date, ordinal, notes, displayName);
      if (name) newIds.push(name);
      console.info(tag + ': createExercise -> ' + (name || '<no name>'));
    } catch (err) {
      console.error(tag + ': createExercise failed: ' + err);
      failed = true;
    }
  }

  if (SYNC_WEIGHT && row.bodyweight !== null) {
    try {
      const name = createWeight(row.date, row.bodyweight);
      if (name) newIds.push(name);
      console.info(tag + ': createWeight(' + row.bodyweight + ' lb) -> ' + (name || '<no name>'));
    } catch (err) {
      console.error(tag + ': createWeight failed: ' + err);
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
