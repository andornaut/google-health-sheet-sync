function setup() {
  ensureManagedColumns();
  installTriggers();
}

function onOpen() {
  console.info('onOpen: installing Sync menu');
  SpreadsheetApp.getUi()
    .createMenu('Sync')
    .addItem('Run now', 'runSyncNow')
    .addItem('Resync selected rows', 'resyncSelectedRows')
    .addItem('Resync all rows', 'resyncAllRows')
    .addSeparator()
    .addItem('Run setup', 'setup')
    .addItem('Authorize Health API', 'authorizeHealthApi')
    .addItem('Revoke Health API', 'revokeHealthApi')
    .addSeparator()
    // Only the parser / pure-helper suite is wired here. The orchestration
    // suite (runSyncTests) depends on the in-memory sheet/properties fakes that
    // test/run.js injects as SYNC_TEST_HARNESS_, which does not exist in the
    // Apps Script runtime — running it here throws ReferenceError. Run it
    // locally with `npm test`.
    .addItem('Run tests', 'runParserTests')
    .addToUi();
}

function authorizeHealthApi() {
  const ui = SpreadsheetApp.getUi();
  let service;
  try {
    service = getHealthService();
  } catch (err) {
    toast_('Setup needed: ' + String(err.message || err), 30);
    return;
  }
  if (service.hasAccess()) {
    toast_('Already authorized. Use "Revoke" first if you want to re-authorize.', 10);
    return;
  }
  const url = service.getAuthorizationUrl();
  console.log('Authorize Google Health API: ' + url);
  const html = '<p>Click the link below, sign in with the Google account that owns your Google Health data, and grant the requested scopes.</p>'
    + '<p><a href="' + url + '" target="_blank" rel="noopener">Authorize Google Health API</a></p>'
    + '<p>After the success page appears, you can close that tab and return here.</p>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(220), 'Authorize Google Health');
}

function revokeHealthApi() {
  resetHealthAuth();
  toast_('Google Health authorization cleared. Use "Authorize Health API" to grant again.', 10);
}

function installTriggers() {
  const handlers = new Set(['syncOnEdit', 'flushPending', 'backstop']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('syncOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('flushPending').timeBased().everyMinutes(POLL_INTERVAL_MIN).create();
  ScriptApp.newTrigger('backstop').timeBased().everyHours(BACKSTOP_INTERVAL_HOURS).create();
}

function syncOnEdit(e) {
  try {
    // onEditMarkDirty does the fast, lock-free work (clear stamps, advance edit
    // timestamps, bump the dirty generation) and reports whether anything was
    // marked dirty. If so, attempt an immediate sync of the dirty row(s) under
    // a non-blocking lock (lockWaitMs=0): if another sync holds the lock this
    // tick skips and the row stays dirty for the next flushPending poll. An
    // unrecoverable throw is logged here (vs. re-thrown for an owner email as in
    // flushPending) since onEdit fires constantly; the poll handles the email.
    if (onEditMarkDirty(e)) {
      syncDirtyRows(0);
    }
  } catch (err) {
    console.error('syncOnEdit error: ' + err);
  }
}

function flushPending() {
  // Sync only when there is pending work. No foreign-match re-review here: the
  // backstop re-dirties recent exercise rows (matched AND unmatched) for
  // foreign alignment, so the 5-min poll issues no Health API calls unless an
  // onEdit (or manual sync) left the dirty flag set. This keeps steady-state
  // read traffic off the API — a row with no pending edits is not re-queried
  // every 5 minutes.
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PENDING_DIRTY_KEY)) {
    console.info('flushPending: no pending edits, skipping');
    return;
  }
  console.info('flushPending: pending edits detected, syncing');
  syncDirtyRows(0);
}

// Select recent exercise rows to re-review, split by foreign-match state:
//   - wantMatched=false: rows NOT yet aligned to a foreign session (empty
//     Matched Health Session). Re-dirtying these lets a Fitbit session that
//     synced late align the row on the next pass.
//   - wantMatched=true: rows already aligned. Re-dirtying these re-borrows the
//     foreign session's CURRENT interval, catching a foreign session whose end
//     was extended after the last sheet edit.
// The backstop re-dirties both; selectBackstopRows_ stays split so each branch
// is independently testable.
// In both cases a row qualifies only when its Date falls within the last
// `lookbackDays` civil days and it has sendable exercise content. Re-dirtying is
// cheap because the exercise re-sync is idempotent (syncOneRow_ skips the
// recreate when nothing changed), so an unmatched row with no new session or a
// matched row with an unchanged interval resolves to a GET + no-op. Pure (date
// keys via ymd, which honors the script time zone).
function selectBackstopRows_(rows, nowMs, lookbackDays, wantMatched) {
  const recentKeys = new Set();
  for (let i = 0; i < lookbackDays; i++) {
    recentKeys.add(ymd(new Date(nowMs - i * 24 * 60 * 60 * 1000)));
  }
  // Order by cost: the free matched-state test, then the ymd date-key lookup
  // (a timezone op), then the nested-loop hasSendableExercises_ last so it runs
  // only on rows that already qualify.
  return rows.filter(r =>
    (wantMatched ? !!r.matchedHealthSession : !r.matchedHealthSession)
    && recentKeys.has(ymd(r.date))
    && hasSendableExercises_(r.exercises));
}

// Backstop trigger (every BACKSTOP_INTERVAL_HOURS): re-dirty recent exercise rows
// (BOTH matched and unmatched) for foreign-match re-review. Clearing Exercise
// Synced At + advancing the dirty generation makes the next flushPending re-run
// the normal sync (including resolveForeignMatches_):
//   - unmatched rows: a Fitbit session that synced after the row was pushed (no
//     further edit to fire onEdit) gets aligned;
//   - matched rows: a foreign session whose interval was extended/changed AFTER
//     the row was last synced gets re-borrowed.
// This runs off the 5-min poll so steady-state foreign re-review doesn't query
// the Health API every 5 minutes; the tradeoff is that a late foreign session
// aligns within BACKSTOP_INTERVAL_HOURS rather than minutes. The exercise re-sync
// is idempotent, so a row whose foreign match/interval did not change resolves to
// a GET + no-op rather than churning the datapoint.
// Also reconciles orphans. No-arg so it's editor-runnable.
function backstop() {
  // Take the script lock for the whole run. Orphan reconciliation deletes
  // exercise datapoints no row references; without the lock it could race an
  // in-flight syncOneRow_ that has POSTed a datapoint but not yet persisted its
  // ID, and delete a legitimately-fresh datapoint. The lock also guarantees the
  // readRows snapshot reflects every persisted ID, so the "known" set is
  // complete. If a sync holds the lock, skip — the next backstop run retries.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    console.warn('backstop: another run holds the lock; skipping this run.');
    return;
  }
  try {
    const now = Date.now();
    const { rows, exerciseSyncedAtCol } = readRows();

    // Reconcile orphans first, while the rows snapshot is unmodified. (reDirty
    // below doesn't touch Created Health IDs, so order is not load-bearing, but
    // reconciling against the clean snapshot keeps the intent obvious.)
    try {
      reconcileExerciseOrphans_(rows, now, ORPHAN_RECONCILE_LOOKBACK_DAYS);
    } catch (err) {
      console.warn('backstop: exercise orphan reconciliation failed (continuing): ' + err);
    }
    try {
      reconcileWeightOrphans_(rows, now, ORPHAN_RECONCILE_LOOKBACK_DAYS);
    } catch (err) {
      console.warn('backstop: weight orphan reconciliation failed (continuing): ' + err);
    }

    if (!exerciseSyncedAtCol) {
      console.warn('backstop: Exercise Synced At column missing; run "Run setup".');
      return;
    }
    // Re-dirty both matched and unmatched recent rows (a row is exactly one of
    // the two, so the union needs no dedup).
    const unmatched = selectBackstopRows_(rows, now, BACKSTOP_LOOKBACK_DAYS, false);
    const matched = selectBackstopRows_(rows, now, BACKSTOP_LOOKBACK_DAYS, true);
    const targets = unmatched.concat(matched);
    if (targets.length === 0) {
      console.info('backstop: no recent exercise rows to re-review.');
      return;
    }
    reDirtyRows_(targets.map(r => r.rowNum), { exerciseCol: exerciseSyncedAtCol });
    console.info('backstop: re-dirtied ' + targets.length
      + ' recent exercise row(s) (' + unmatched.length + ' unmatched, '
      + matched.length + ' matched) for foreign-match re-review.');
  } finally {
    lock.releaseLock();
  }
}

// Given datapoint candidates listed across the reconciliation window (each an
// object carrying at least `name` and `googleWebClientId`) and the set of
// datapoint names already tracked in some row's Created Health IDs (an object
// used as a name -> true set), return the candidate names safe to delete as
// orphans: created by the SAME web client that owns our tracked datapoints, but
// referenced by no row.
//
// "Our client" is derived from the candidates themselves — the
// googleWebClientId of any candidate that IS tracked — rather than from a
// configured value, so we only ever delete datapoints from the exact client
// that created the ones we still track. Foreign device / first-party / in-app-
// assistant sessions (null googleWebClientId) and any other web app are never
// selected. If no tracked candidate is present we can't attribute ownership,
// so nothing is returned. Type-agnostic and pure (no API/sheet access) — used
// for both exercise and weight reconciliation.
function selectOrphanDataPointNames_(candidates, knownNames) {
  const ourClientIds = {};
  candidates.forEach(c => {
    if (c.googleWebClientId && knownNames[c.name]) {
      ourClientIds[c.googleWebClientId] = true;
    }
  });
  const orphans = [];
  candidates.forEach(c => {
    if (knownNames[c.name]) return;
    if (c.googleWebClientId && ourClientIds[c.googleWebClientId]) {
      orphans.push(c.name);
    }
  });
  return orphans;
}

// Delete sync-created datapoints of one data type that no row's Created Health
// IDs references — orphans leaked by the two accepted create windows (a create
// POST that succeeded server-side but timed out client-side and was retried; a
// 6-minute hard kill after the POST returned but before the ID was persisted).
// `typeKey` selects the splitHealthIdsByType_ bucket ('exercise' or 'weight')
// and `listOnDate(date)` lists same-type candidates for a civil date (each with
// `name` + `googleWebClientId`). Scans the last `lookbackDays` civil days,
// derives ownership from the listed datapoints (see selectOrphanDataPointNames_),
// and deletes the orphans. Each delete is independent: a failure (including a
// 403 if a name turns out to be foreign) is logged and skipped, never aborting
// the rest. Must run under the script lock (see backstop) so an in-flight
// sync's not-yet-persisted create can't be mistaken for an orphan.
function reconcileDataPointOrphans_(rows, nowMs, lookbackDays, typeKey, listOnDate) {
  const tag = 'reconcileDataPointOrphans_(' + typeKey + ')';
  const known = {};
  rows.forEach(r => {
    splitHealthIdsByType_(r.healthIds)[typeKey].forEach(n => { known[n] = true; });
  });

  const candidates = [];
  for (let i = 0; i < lookbackDays; i++) {
    const date = new Date(nowMs - i * 24 * 60 * 60 * 1000);
    try {
      listOnDate(date).forEach(c => candidates.push(c));
    } catch (err) {
      console.warn(tag + ': list failed for ' + ymd(date) + ': ' + err);
    }
  }

  const orphans = selectOrphanDataPointNames_(candidates, known);
  if (orphans.length === 0) {
    console.info(tag + ': no orphans found (' + candidates.length + ' datapoint(s) scanned).');
    return;
  }
  let deleted = 0;
  orphans.forEach(name => {
    try {
      deleteDataPointsByName([name]);
      deleted++;
      console.info(tag + ': deleted orphan ' + name);
    } catch (err) {
      console.warn(tag + ': delete failed for ' + name + ': ' + err);
    }
  });
  console.info(tag + ': removed ' + deleted + ' of ' + orphans.length + ' orphan datapoint(s).');
}

// Reconcile orphaned sync-created exercise datapoints (STRENGTH_TRAINING).
function reconcileExerciseOrphans_(rows, nowMs, lookbackDays) {
  reconcileDataPointOrphans_(rows, nowMs, lookbackDays, 'exercise', listStrengthOnDate);
}

// Reconcile orphaned sync-created weight datapoints. The exercise path's
// create-orphan windows (timed-out-but-succeeded POST retry; hard kill after
// POST before the ID is persisted) apply to the weight POST in syncOneRow_ too,
// so a duplicate untracked weight datapoint can leak the same way. Same
// ownership logic — only datapoints from our web client, referenced by no row,
// are deleted.
function reconcileWeightOrphans_(rows, nowMs, lookbackDays) {
  reconcileDataPointOrphans_(rows, nowMs, lookbackDays, 'weight', listWeightOnDate);
}

// Write a fresh generation marker into PENDING_DIRTY_KEY. The value matters
// (syncDirtyRows compares start vs end to detect concurrent edits), so always
// advance it — never just re-write the same string.
function markPendingDirty_() {
  PropertiesService.getScriptProperties()
    .setProperty(PENDING_DIRTY_KEY, String(Date.now()));
}

// Re-dirty a set of rows: clear the requested synced stamp(s), persist the
// writes, and advance the dirty generation so the next sync pass re-processes
// them. Shared by the manual selective resync and the backstop; pass
// cols.exerciseCol and/or cols.weightCol for the phase(s) to re-dirty. (The
// "resync all rows" path clears its columns in a single batched setValues
// instead, so it doesn't route through here.)
function reDirtyRows_(rowNums, cols) {
  rowNums.forEach(rowNum => {
    if (cols.exerciseCol) clearRowExerciseSynced(rowNum, cols.exerciseCol);
    if (cols.weightCol) clearRowWeightSynced(rowNum, cols.weightCol);
  });
  SpreadsheetApp.flush();
  markPendingDirty_();
}

// "Run now" is an explicit manual action: syncs all dirty rows immediately,
// waiting for the script lock (unlike the automatic triggers, which skip when
// the lock is held). If the user keeps editing afterward, the row goes dirty
// again and the next sync reconciles the Health datapoint(s).
function runSyncNow() {
  runSyncAndToast_('Synced');
}

// Manual sync entry points share this wrapper so an unexpected throw from
// syncDirtyRows (e.g. readRows can't find the Date column) surfaces as a
// toast instead of Apps Script's modal error dialog. The error is re-thrown
// so it still lands in Executions for diagnosis.
function runSyncAndToast_(verb) {
  let result;
  try {
    result = syncDirtyRows(LOCK_WAIT_MS);
  } catch (err) {
    toast_('Sync failed: ' + String(err.message || err), 30);
    throw err;
  }
  toastSyncResult_(result, verb);
}

function resyncSelectedRows() {
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  const exerciseCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER];
  if (!exerciseCol) {
    toast_('Exercise Synced At column missing. Run setup.', 30);
    return;
  }
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER];
  if (!weightSyncedAtCol) {
    toast_('Weight Synced At column missing. Run setup.', 30);
    return;
  }
  const rows = {};
  const ranges = sheet.getActiveRangeList().getRanges();
  for (const range of ranges) {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let row = start; row <= end; row++) {
      if (row >= 2) rows[row] = true;
    }
  }
  const rowNums = Object.keys(rows);
  if (rowNums.length === 0) {
    toast_('No data rows selected.', 10);
    return;
  }
  reDirtyRows_(rowNums.map(Number), { exerciseCol: exerciseCol, weightCol: weightSyncedAtCol });
  runSyncAndToast_('Resynced');
}

function resyncAllRows() {
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  const exerciseCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER];
  if (!exerciseCol) {
    toast_('Exercise Synced At column missing. Run setup.', 30);
    return;
  }
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER];
  if (!weightSyncedAtCol) {
    toast_('Weight Synced At column missing. Run setup.', 30);
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    toast_('No data rows.', 10);
    return;
  }
  const dataRowCount = lastRow - 1;

  const blanks = [];
  for (let i = 0; i < dataRowCount; i++) blanks.push(['']);
  sheet.getRange(2, exerciseCol, dataRowCount, 1).setValues(blanks);
  sheet.getRange(2, weightSyncedAtCol, dataRowCount, 1).setValues(blanks);
  SpreadsheetApp.flush();
  markPendingDirty_();

  runSyncAndToast_('Resynced');
}

function formatSyncResult_(result, verb) {
  if (!result) return 'Sync skipped (another run holds the lock). Try again shortly.';
  let msg = verb + ' ' + result.ok + ' row(s)';
  if (result.errors > 0) msg += ', ' + result.errors + ' error(s)';
  if (result.deferred > 0) msg += ', ' + result.deferred + ' deferred';
  msg += '.';
  if (result.errors > 0) msg += '\n\nSee Executions for details.';
  return msg;
}

// Non-blocking status notification. Apps Script's ui.alert() is modal and
// counts against the 6-minute execution budget, which caused timeouts on
// large resyncs. Toasts auto-dismiss and don't block the script.
function toast_(msg, seconds) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Sync', seconds);
}

function toastSyncResult_(result, verb) {
  const seconds = (result && result.errors > 0) ? 30 : 10;
  toast_(formatSyncResult_(result, verb), seconds);
}

function humanizeMs_(ms) {
  if (ms < 0) ms = 0;
  if (ms < 1000) return ms + 'ms';
  const totalSec = Math.round(ms / 1000);
  if (totalSec < 60) return totalSec + 's';
  const min = Math.floor(totalSec / 60);
  const sec = totalSec % 60;
  return sec === 0 ? min + 'm' : min + 'm ' + sec + 's';
}

function humanizeDate_(date) {
  if (!date) return '<none>';
  return Utilities.formatDate(date, getTz_(), 'yyyy-MM-dd HH:mm:ss');
}

function syncDirtyRows(lockWaitMs) {
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
  // Capture the dirty-marker generation at start. onEditMarkDirty advances it
  // on every edit (Date.now() string), so a concurrent edit during the pass
  // shows up as a mismatch at end-of-pass. The flag is NOT cleared here:
  // a hard kill (6-min Apps Script timeout, uncaught throw) before the
  // finally block would otherwise drop the signal and orphan the dirty rows
  // until the next manual sync or new edit.
  const genAtStart = props.getProperty(PENDING_DIRTY_KEY);
  let ok = 0;
  let errors = 0;
  let deferredCount = 0;
  // Set when the pass hits a misconfiguration that retrying can't fix (missing
  // columns, duplicate headers, or any other throw out of the body). The catch
  // emails the owner; the finally clears the dirty flag so we don't re-run and
  // re-email every poll. A later edit or manual sync re-triggers.
  let unrecoverable = false;
  try {
    const { rows, exerciseSyncedAtCol, weightSyncedAtCol, healthIdsCol, exercisesLastEditedAtCol, weightEditedAtCol, matchedHealthSessionCol } = readRows();
    if (!exerciseSyncedAtCol || !weightSyncedAtCol || !healthIdsCol) {
      throw new Error('Managed columns missing; run "Run setup" from the Sync menu before syncing.');
    }
    const dirty = rows.filter(r => !r.exerciseSyncedAt || !r.weightSyncedAt);
    if (dirty.length === 0) {
      return { ok: 0, errors: 0 };
    }

    const ordinalByRowNum = buildOrdinalMap_(rows);

    // Per-row phase readiness (no debounce — onEdit syncs immediately and the
    // poll/backstop retry): a row's weight phase is ready when weight-dirty, its
    // exercise phase when exercise-dirty. There is no "still typing" wait; a
    // poll firing mid-burst just re-pushes current state, and the idempotent
    // exercise re-sync (syncOneRow_) skips the recreate when nothing changed.
    const ready = [];
    dirty.forEach(r => {
      const weightReady = !r.weightSyncedAt;
      const exerciseReady = !r.exerciseSyncedAt;
      if (weightReady || exerciseReady) {
        ready.push({ row: r, weightReady: weightReady, exerciseReady: exerciseReady });
      }
    });

    if (ready.length === 0) {
      console.info('syncDirtyRows: no rows ready to sync.');
      return { ok: 0, errors: 0 };
    }

    // Newest-first so recent edits land in Health quickly when the cap defers
    // some of the backlog. Tie-break by rowNum descending for stable ordering
    // within a single date.
    ready.sort((a, b) => {
      const dateDiff = b.row.date.getTime() - a.row.date.getTime();
      return dateDiff !== 0 ? dateDiff : b.row.rowNum - a.row.rowNum;
    });
    if (ready.length > MAX_ROWS_PER_SYNC) {
      deferredCount = ready.length - MAX_ROWS_PER_SYNC;
      console.info('syncDirtyRows: ' + ready.length + ' ready row(s); capping at '
        + MAX_ROWS_PER_SYNC + ', deferring ' + deferredCount + ' to next pass');
      ready.length = MAX_ROWS_PER_SYNC;
    } else {
      console.info('syncDirtyRows: ' + ready.length + ' ready row(s)');
    }

    // Every exercise-dirty row creates its own datapoint; the plan only
    // supplies a foreign session's interval to align timing when one overlaps
    // the row's edit window. See FOREIGN_MATCH_BUFFER_MS in Config.gs.
    const exerciseReadyRows = ready.filter(r => r.exerciseReady).map(r => r.row);
    const alignmentPlan = resolveForeignMatches_(rows, exerciseReadyRows);
    const cols = {
      exerciseSyncedAtCol: exerciseSyncedAtCol,
      weightSyncedAtCol: weightSyncedAtCol,
      healthIdsCol: healthIdsCol,
      exercisesLastEditedAtCol: exercisesLastEditedAtCol,
      weightEditedAtCol: weightEditedAtCol,
      matchedHealthSessionCol: matchedHealthSessionCol
    };
    for (let i = 0; i < ready.length; i++) {
      const entry = ready[i];
      const ordinal = ordinalByRowNum[entry.row.rowNum];
      const foreignMatch = alignmentPlan[entry.row.rowNum] || null;
      if (syncOneRow_(entry.row, ordinal, foreignMatch, entry.weightReady, entry.exerciseReady, cols, i + 1, ready.length)) ok++;
      else errors++;
    }
  } catch (err) {
    // Unrecoverable: a throw out of the sync body (missing required columns,
    // duplicate exercise headers, etc.). Per-row API failures never reach here
    // — syncOneRow_ catches them and they retry. Re-throw so the failure
    // propagates uncaught: Apps Script then emails the script owner about the
    // failed trigger execution (no MailApp needed), and manual entry points
    // still toast and land the error in Executions. The finally LEAVES the
    // dirty flag set (see below) so the next poll retries automatically once
    // the misconfig is fixed — clearing it would orphan already-dirty rows
    // until some future edit. Repeated failure emails are throttled via the
    // trigger's notification cadence (Apps Script ▸ Triggers ▸ notifications),
    // not by suppressing the retry.
    unrecoverable = true;
    console.error('syncDirtyRows: unrecoverable error: ' + err);
    throw err;
  } finally {
    // NB: no `return` in this finally — a finally-return would swallow the
    // re-thrown unrecoverable error and defeat the manual-path toast.
    if (unrecoverable) {
      // Leave the dirty flag untouched (still set) so the backlog syncs on the
      // next poll after the misconfig is fixed. Just release the lock.
    } else {
      // End-of-pass flag resolution:
      //   - If work remains (errors, deferred): ensure the flag is set so a
      //     future poll picks it up. If a concurrent edit advanced the
      //     generation already, its value is fine; otherwise write a fresh one.
      //     (syncOneRow_ also calls markPendingDirty_ for partial-progress rows,
      //     so ok-counted rows can leave the flag set too.)
      //   - If no work remains AND the generation hasn't moved: the pass fully
      //     drained the queue, safe to clear.
      //   - If no work remains BUT the generation moved: an edit landed during
      //     the pass that this pass didn't see (readRows snapshotted before it).
      //     Leave the new generation in place so the next poll runs.
      const genAtEnd = props.getProperty(PENDING_DIRTY_KEY);
      const concurrentEdit = genAtEnd !== genAtStart;
      const workRemaining = errors > 0 || deferredCount > 0;
      if (workRemaining) {
        if (!concurrentEdit) markPendingDirty_();
      } else if (!concurrentEdit) {
        props.deleteProperty(PENDING_DIRTY_KEY);
      }
      // Bookend the pass: outcome counts plus whether the dirty flag survives
      // (so a reader knows if another poll will follow without inspecting it).
      const willRetry = workRemaining || concurrentEdit;
      console.info('syncDirtyRows: pass complete — ' + ok + ' synced, ' + errors + ' error(s)'
        + (deferredCount ? ', ' + deferredCount + ' deferred' : '')
        + (concurrentEdit ? ', concurrent edit landed' : '')
        + (willRetry ? '; pending flag kept, next poll will retry.' : '; queue drained.'));
    }
    lock.releaseLock();
  }
  return { ok: ok, errors: errors, deferred: deferredCount };
}

// Returns rowNum -> foreign Strength Training session whose interval should be
// copied onto the row's created exercise datapoint (timing alignment). The
// sync never skips its own create; this only supplies a more accurate
// start/end when a manually-logged foreign session overlaps the row's edit
// window. Time-overlap only — rows without same-date exercise edit timestamps
// get no match and fall through to synthetic/prior timing (there is no ordinal
// fallback).
//
// Candidate exclusions (global, keyed by resource name — names are globally
// unique, so no date keying is needed, and global keying is what lets a
// neighbor-day candidate be excluded correctly):
//   - sync-created: name appears in some row's Created Health IDs (our own
//     datapoint — never align to ourselves, including a row's prior datapoint
//     on re-sync).
//   - aligned-elsewhere: name is a non-ready row's Matched Health Session, so
//     two rows don't borrow the same foreign session's times across runs.
//
// Cross-date: candidates are gathered for every civil date any row's window
// touches (a window near local midnight pulls the neighbor day too), deduped
// by name. Overlap is computed in absolute UTC, so midnight-crossing workouts
// match regardless of which civil date the foreign session was logged under.
//
// Ordering note: this runs once near the top of syncDirtyRows, before the
// per-row syncOneRow_ loop. The `allRows` snapshot comes from readRows() at
// the start of the pass, and `listStrengthOnDate` calls the API before any
// row in this pass has had its create issued. So same-pass freshly-created
// datapoints are not yet visible to the API and not yet in any row's
// Created Health IDs — both gaps cancel out and there's no self-match risk.
function resolveForeignMatches_(allRows, readyRows) {
  const plan = {};
  const readyRowNums = {};
  readyRows.forEach(r => { readyRowNums[r.rowNum] = true; });

  const excluded = {};
  allRows.forEach(r => {
    if (!readyRowNums[r.rowNum] && r.matchedHealthSession) {
      excluded[r.matchedHealthSession] = true;
    }
    splitHealthIdsByType_(r.healthIds).exercise.forEach(name => { excluded[name] = true; });
  });

  // Only rows with on-row-date edit timestamps anchor a trustworthy window.
  // capExerciseDurationToMax_ caps the window the same way resolveRowTiming_
  // caps the recorded interval, so a row whose exercisesLastEditedAt drifted
  // far past exerciseFirstEditedAt (late corrections that keep sticky
  // first-edit + advance last-edit) doesn't produce a multi-day window biased
  // toward the longest unrelated candidate.
  const windows = readyRows
    .filter(r => hasSendableExercises_(r.exercises) && exerciseEditIsOnRowDate_(r))
    .map(r => {
      const startMs = r.exerciseFirstEditedAt.getTime();
      const clampedEndMs = startMs + capExerciseDurationToMax_(r.exercisesLastEditedAt.getTime() - startMs);
      return {
        rowNum: r.rowNum,
        windowStart: startMs - FOREIGN_MATCH_BUFFER_MS,
        windowEnd: clampedEndMs + FOREIGN_MATCH_BUFFER_MS
      };
    });
  if (windows.length === 0) return plan;

  // Gather candidates across every civil date a window touches (start and end
  // edges — adjacent when a window straddles midnight), deduped by name and
  // with our own / aligned-elsewhere names dropped.
  const probeDates = {};
  windows.forEach(w => {
    const start = new Date(w.windowStart);
    const end = new Date(w.windowEnd);
    probeDates[ymd(start)] = start;
    probeDates[ymd(end)] = end;
  });
  const byName = {};
  Object.keys(probeDates).forEach(key => {
    let list;
    try {
      list = listStrengthOnDate(probeDates[key]);
    } catch (err) {
      console.warn('resolveForeignMatches_: list failed for ' + key + ': ' + err);
      return;
    }
    list.forEach(c => {
      if (excluded[c.name]) return;
      byName[c.name] = c;
    });
  });
  const candidates = Object.values(byName);
  if (candidates.length === 0) return plan;

  // Assign in rowNum order; each row claims its best-overlap remaining
  // candidate so two rows can't align to the same foreign session.
  windows.sort((a, b) => a.rowNum - b.rowNum);
  windows.forEach(w => {
    let bestIdx = -1;
    let bestOverlap = 0;
    candidates.forEach((c, i) => {
      const overlap = Math.min(c.endUtcMs, w.windowEnd) - Math.max(c.startUtcMs, w.windowStart);
      if (overlap > bestOverlap) {
        bestIdx = i;
        bestOverlap = overlap;
      }
    });
    if (bestIdx >= 0) {
      plan[w.rowNum] = candidates[bestIdx];
      console.info('resolveForeignMatches_: row ' + w.rowNum + ' aligns to '
        + candidates[bestIdx].name + ' (overlap=' + humanizeMs_(bestOverlap) + ')');
      candidates.splice(bestIdx, 1);
    }
  });
  return plan;
}

function groupRowsByDate_(rows) {
  const byDate = {};
  rows.forEach(r => {
    const key = ymd(r.date);
    (byDate[key] = byDate[key] || []).push(r);
  });
  return byDate;
}

// Each row's rank (by rowNum) within its civil date. Used by the synthetic-
// timing fallback to give same-date rows distinct startHour offsets.
function buildOrdinalMap_(rows) {
  const byDate = groupRowsByDate_(rows);
  const ordinalByRowNum = {};
  Object.keys(byDate).forEach(dateKey => {
    const dateRows = byDate[dateKey];
    dateRows.sort((a, b) => a.rowNum - b.rowNum);
    dateRows.forEach((r, i) => { ordinalByRowNum[r.rowNum] = i; });
  });
  return ordinalByRowNum;
}

// Cap an edit-derived exercise duration at MAX_EXERCISE_DURATION_MS (MAX only,
// no MIN floor). Used by the foreign-match window in resolveForeignMatches_ to
// keep its upper bound consistent with editDerivedDurationMs_ — which the
// recorded 'edit' interval uses, and which applies the same MAX cap PLUS a MIN
// floor / start-only default that this helper deliberately omits.
function capExerciseDurationToMax_(rawDurationMs) {
  return Math.min(rawDurationMs, MAX_EXERCISE_DURATION_MS);
}

// True when the row's exercise edit timestamps form a trustworthy same-day
// window: both first/last edit are set AND the first edit's civil date matches
// the row's Date. This gates the live-workout 'edit' timing path (endTime can
// advance during a workout) and the foreign-match window — an off-date edit
// (correcting an old row today) must not anchor timing to today.
function exerciseEditIsOnRowDate_(row) {
  return !!(row.exerciseFirstEditedAt && row.exercisesLastEditedAt
    && ymd(row.exerciseFirstEditedAt) === ymd(row.date));
}

// Map a raw edit-derived duration (last edit - first edit) to the recorded
// exercise duration by clamping to [MIN_EXERCISE_DURATION_MS,
// MAX_EXERCISE_DURATION_MS]. A single-edit row has raw <= 0 (start == last
// edit, no observed end), which the MIN floor turns into the start-only default
// — so the start-only case needs no special handling.
function editDerivedDurationMs_(rawDurationMs) {
  return Math.min(Math.max(rawDurationMs, MIN_EXERCISE_DURATION_MS), MAX_EXERCISE_DURATION_MS);
}

// Resolve the exercise interval and weight sample time independently, per
// phase, with these rules:
//
//   Exercise (first matching rule wins):
//     - 'foreign'   if a foreignInterval is provided (an overlapping foreign
//                   session whose manual start/stop is more accurate than our
//                   edit-derived window). Its interval is used verbatim.
//     - 'edit'      if exerciseFirstEditedAt's civil date == row.date AND
//                   exercisesLastEditedAt is set. This lets endTime advance
//                   during a live workout as more sets are typed in.
//     - 'prior'     if a previous datapoint is provided. Its interval is
//                   reused verbatim so an off-date edit (e.g. correcting
//                   an old row today) doesn't shift startTime to today.
//     - 'synthetic' otherwise: noon+ordinal on row.date.
//
//   Weight (only consumed on the POST path; PATCH preserves sampleTime
//   server-side by echoing back the prior GET, so the weight resolver
//   isn't called with a prior):
//     - 'edit'      if weightEditedAt's civil date == row.date.
//     - 'synthetic' otherwise: noon on row.date.
//
// priorExercise is the GET response for the row's existing exercise
// datapoint (or null if first-sync, or null if the GET failed — in which
// case we fall through to the edit/synthetic path rather than erroring).
// foreignInterval is an overlapping foreign session (from
// resolveForeignMatches_) whose start/end should be borrowed, or null.
function resolveRowTiming_(row, ordinal, priorExercise, foreignInterval) {
  const tz = getTz_();
  const rowDateKey = ymd(row.date);

  let exercise = null;
  let exerciseSource = null;
  const exerciseEditOnRowDate = exerciseEditIsOnRowDate_(row);
  if (foreignInterval) {
    exercise = {
      startUtcMs: foreignInterval.startUtcMs,
      startOffsetSeconds: foreignInterval.startUtcOffsetSeconds,
      endUtcMs: foreignInterval.endUtcMs,
      endOffsetSeconds: foreignInterval.endUtcOffsetSeconds
    };
    exerciseSource = 'foreign';
  } else if (exerciseEditOnRowDate) {
    const startMs = row.exerciseFirstEditedAt.getTime();
    const rawDuration = row.exercisesLastEditedAt.getTime() - startMs;
    // A single edit (start == last, rawDuration <= 0) has no observed end;
    // editDerivedDurationMs_'s MIN floor gives it the start-only default. A
    // second edit produces a real span (clamped to [MIN, MAX]).
    const endMs = startMs + editDerivedDurationMs_(rawDuration);
    exercise = {
      startUtcMs: startMs,
      startOffsetSeconds: getTzOffsetSeconds_(tz, row.exerciseFirstEditedAt),
      endUtcMs: endMs,
      endOffsetSeconds: getTzOffsetSeconds_(tz, new Date(endMs))
    };
    exerciseSource = 'edit';
  } else if (priorExercise) {
    const i = priorExercise.exercise && priorExercise.exercise.interval;
    if (i && i.startTime && i.endTime) {
      exercise = {
        startUtcMs: new Date(i.startTime).getTime(),
        startOffsetSeconds: parseOffsetSeconds_(i.startUtcOffset),
        endUtcMs: new Date(i.endTime).getTime(),
        endOffsetSeconds: parseOffsetSeconds_(i.endUtcOffset)
      };
      exerciseSource = 'prior';
    }
  }
  if (!exercise) {
    exercise = syntheticExerciseInterval_(row.date, ordinal);
    exerciseSource = 'synthetic';
  }

  let weight;
  let weightSource;
  if (row.weightEditedAt && ymd(row.weightEditedAt) === rowDateKey) {
    weight = {
      utcMs: row.weightEditedAt.getTime(),
      offsetSeconds: getTzOffsetSeconds_(tz, row.weightEditedAt)
    };
    weightSource = 'edit';
  } else {
    weight = syntheticWeightSample_(row.date);
    weightSource = 'synthetic';
  }
  return {
    exercise: exercise,
    weight: weight,
    exerciseSource: exerciseSource,
    weightSource: weightSource
  };
}

// True when an existing exercise datapoint (the GET response) already carries
// the target interval and notes, so a re-sync can skip the delete+recreate and
// keep its resource name (no churn). Compares interval start/end in absolute ms
// (the offsets are derived from the same instants, so ms equality suffices) and
// the notes string exactly. A missing interval/endpoint counts as changed so
// the row recreates. Pure (no API/sheet access).
function exerciseUnchanged_(prior, targetStartUtcMs, targetEndUtcMs, targetNotes) {
  const ex = prior && prior.exercise;
  const i = ex && ex.interval;
  if (!i || !i.startTime || !i.endTime) return false;
  if (new Date(i.startTime).getTime() !== targetStartUtcMs) return false;
  if (new Date(i.endTime).getTime() !== targetEndUtcMs) return false;
  return (ex.notes || '') === (targetNotes || '');
}

// Sync a single row in two independent phases (weight, exercise). Either or
// both phases may run on a given pass:
//   - Weight phase runs whenever the row's Weight Synced At is cleared. It
//     reconciles weight IDs with the sheet's bodyweight (write/delete) and
//     stamps Weight Synced At on success.
//   - Exercise phase runs whenever the caller passed exerciseReady=true (the
//     row's Exercise Synced At is cleared). It reconciles exercise IDs with the
//     sheet's exercises (delete + recreate, optionally aligning the interval to
//     an overlapping foreign session) and stamps Exercise Synced At on success.
//     The recreate is skipped when the existing datapoint already matches the
//     freshly-computed interval + notes (see exerciseUnchanged_), so a re-sync
//     of an unchanged row keeps its resource name and costs only a GET.
// Returns true if the pass made forward progress on the row without errors
// (including the case where the row stays dirty because the other phase is
// still pending). Returns false if any attempted phase failed.
function syncOneRow_(row, ordinal, foreignMatch, weightReady, exerciseReady, cols, doneIdx, total) {
  const dateKey = ymd(row.date);
  const tag = '[' + doneIdx + '/' + total + '] ' + dateKey + ' row ' + row.rowNum;
  const phases = [];
  if (weightReady) phases.push('weight');
  if (exerciseReady) phases.push('exercise');
  console.info(tag + ': starting phases=[' + phases.join(',') + '] (exercises=' + row.exercises.length
    + ', bodyweight=' + (row.bodyweight === null ? 'none' : row.bodyweight)
    + ', oldIds=' + row.healthIds.length
    + (foreignMatch ? ', align=' + foreignMatch.name : '') + ')');

  const split = splitHealthIdsByType_(row.healthIds);

  // Phases that will actually issue a create (and thus need a resolved
  // interval/sampleTime). Delete-only phases don't need timing. The
  // timing log line shows only the phases listed here, so weight-only rows
  // don't surface a misleading "edit/synthetic" label for an interval that
  // will never be sent.
  const exerciseWillCreate = exerciseReady && hasSendableExercises_(row.exercises);

  // Fetch prior datapoints. Exercise: whenever the row has a prior exercise id
  // and the phase will create — the GET serves two purposes now. (1) The 'prior'
  // timing source reuses its interval verbatim when neither foreign-match nor
  // same-date editing applies (the resolver ignores it for foreign/edit, which
  // win). (2) The idempotency check compares the prior interval + notes to the
  // freshly-computed ones to skip an unchanged recreate. Weight: when we'll
  // PATCH (i.e. prior weight ID exists AND bodyweight is set) — the PATCH body
  // requires sampleTime, read from this GET. Exercise GET failure is non-fatal
  // (timing falls through to edit/synthetic and the recreate proceeds); weight
  // GET failure forces the PATCH to fail and the row to retry next pass.
  let priorExercise = null;
  let priorWeight = null;
  let priorWeightFetchFailed = false;
  const exerciseEditOnRowDate = exerciseEditIsOnRowDate_(row);
  if (exerciseWillCreate && split.exercise.length > 0) {
    try {
      priorExercise = getDataPoint(split.exercise[0]);
    } catch (err) {
      console.warn(tag + ': GET prior exercise failed; will recompute timing and recreate: ' + err);
    }
  }
  const weightWillPatch = weightReady && row.bodyweight !== null && split.weight.length > 0;
  if (weightWillPatch) {
    try {
      priorWeight = getDataPoint(split.weight[0]);
    } catch (err) {
      if (isNotFoundError_(err)) {
        // The prior weight datapoint is gone server-side (e.g. deleted in the
        // Health app). Drop the stale ID so the dispatch below falls through to
        // POST a fresh one instead of PATCHing/GETting a name that 404s forever.
        console.warn(tag + ': prior weight datapoint not found (404); dropping stale ID and recreating.');
        split.weight = [];
      } else {
        console.warn(tag + ': GET prior weight failed; PATCH will fail and the row will retry: ' + err);
        priorWeightFetchFailed = true;
      }
    }
  }
  // Only POST creates need timing resolution. The PATCH path (prior weight ID
  // present + bodyweight set) preserves sampleTime server-side, so no prior
  // GET and no timing label. Computed after the GET block so a 404-dropped
  // stale ID is reflected here (the row now POSTs rather than PATCHes).
  const weightWillCreate = weightReady && row.bodyweight !== null && split.weight.length === 0;

  let timing;
  try {
    // `foreignMatch`, when set, is an overlapping foreign session whose
    // interval the resolver borrows verbatim ('foreign' wins over
    // edit/prior/synthetic).
    timing = resolveRowTiming_(row, ordinal, priorExercise, foreignMatch);
  } catch (err) {
    console.error(tag + ': resolveRowTiming_ failed: ' + err);
    return false;
  }
  const labelParts = [];
  if (exerciseWillCreate) labelParts.push('exercise=' + timing.exerciseSource);
  if (weightWillCreate) labelParts.push('weight=' + timing.weightSource);
  if (labelParts.length > 0) console.info(tag + ': timing ' + labelParts.join(' '));
  let newWeightIds = split.weight;
  let newExerciseIds = split.exercise;
  let weightFailed = false;
  let exerciseFailed = false;

  if (weightReady) {
    const hasBodyweight = row.bodyweight !== null;
    if (split.weight.length > 0 && hasBodyweight) {
      // PATCH in place. Preserves sampleTime (echoed back from the prior
      // GET — the API rejects PATCH bodies without sampleTime), createTime,
      // dataSource. Resource name stays the same so Created Health IDs
      // doesn't churn.
      const sampleTime = priorWeight && priorWeight.weight && priorWeight.weight.sampleTime;
      if (!sampleTime) {
        const reason = priorWeightFetchFailed
          ? 'prior weight GET failed'
          : 'prior datapoint missing sampleTime';
        console.error(tag + ': patchWeight skipped (' + reason + '); will retry next sync.');
        weightFailed = true;
      } else {
        try {
          patchWeight(split.weight[0], sampleTime, row.bodyweight);
          console.info(tag + ': patchWeight(' + row.bodyweight + ' lb) -> ' + split.weight[0]);
        } catch (err) {
          console.error(tag + ': patchWeight failed: ' + err);
          weightFailed = true;
        }
      }
    } else if (split.weight.length > 0 && !hasBodyweight) {
      // Bodyweight cleared on a row that previously had one: delete.
      console.info(tag + ': deleting ' + split.weight.length + ' previous weight datapoint(s)');
      try {
        deleteDataPointsByName(split.weight);
        newWeightIds = [];
      } catch (err) {
        console.error(tag + ': delete previous weight datapoint(s) failed: ' + err);
        weightFailed = true;
        // Keep newWeightIds = split.weight so the next sync retries delete.
      }
    } else if (split.weight.length === 0 && hasBodyweight) {
      // First weight for this row: POST.
      newWeightIds = [];
      try {
        const wt = timing.weight;
        // createWeightAt throws if the create returns no resource name, so a
        // returned name is always usable here.
        const name = createWeightAt(wt.utcMs, wt.offsetSeconds, row.bodyweight);
        newWeightIds.push(name);
        // Persist immediately so a 6-minute kill before the end-of-row
        // write can't orphan a freshly-created datapoint we no longer
        // have a sheet reference for.
        writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
        console.info(tag + ': createWeightAt(' + row.bodyweight + ' lb) -> ' + name);
      } catch (err) {
        console.error(tag + ': createWeightAt failed: ' + err);
        weightFailed = true;
      }
    } else {
      // No prior, no current. Nothing to do.
      newWeightIds = [];
    }
  }

  if (exerciseReady) {
    const ex = timing.exercise;
    const wantCreate = hasSendableExercises_(row.exercises);
    const notes = wantCreate ? buildNotes(ex.endUtcMs - ex.startUtcMs, row.exercises) : null;

    // Idempotency: if the row's single existing exercise datapoint already
    // carries the target interval + notes, skip the delete+recreate entirely
    // and keep its resource name. This is what makes the per-poll / per-day
    // re-dirty cheap — an unchanged row costs just the prior GET, no write and
    // no resource-name churn. Only applies when there's exactly one prior id
    // (multiple priors are consolidated by recreating).
    const unchanged = wantCreate && split.exercise.length === 1 && priorExercise
      && exerciseUnchanged_(priorExercise, ex.startUtcMs, ex.endUtcMs, notes);

    if (unchanged) {
      console.info(tag + ': exercise unchanged; skip recreate -> ' + split.exercise[0]);
      newExerciseIds = split.exercise;
    } else {
      if (split.exercise.length > 0) {
        console.info(tag + ': deleting ' + split.exercise.length + ' previous exercise datapoint(s)');
        try {
          deleteDataPointsByName(split.exercise);
          newExerciseIds = [];
        } catch (err) {
          if (isNotFoundError_(err)) {
            // Already gone server-side (e.g. deleted in the Health app). Treat as
            // deleted so the row recreates instead of retrying a delete that
            // 404s forever.
            console.warn(tag + ': previous exercise datapoint(s) not found (404); treating as deleted.');
            newExerciseIds = [];
          } else {
            console.error(tag + ': delete previous exercise datapoint(s) failed: ' + err);
            exerciseFailed = true;
            // Keep newExerciseIds = split.exercise so the next sync retries delete.
          }
        }
      } else {
        newExerciseIds = [];
      }
      if (!exerciseFailed && wantCreate) {
        try {
          // createExerciseAt throws if the create returns no resource name, so a
          // returned name is always usable here.
          const name = createExerciseAt(ex.startUtcMs, ex.startOffsetSeconds,
            ex.endUtcMs, ex.endOffsetSeconds, notes);
          newExerciseIds.push(name);
          // Same rationale as the weight write above: persist before any
          // later step can fail and leave the datapoint untracked.
          writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
          console.info(tag + ': createExerciseAt' + (foreignMatch ? ' (foreign-aligned)' : '')
            + ' -> ' + name);
        } catch (err) {
          console.error(tag + ': createExerciseAt failed: ' + err);
          exerciseFailed = true;
        }
      }
    }
  }

  writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
  if (exerciseReady) {
    writeMatchedHealthSession(row.rowNum, cols.matchedHealthSessionCol, foreignMatch ? foreignMatch.name : '');
  }

  // Concurrent-edit guards, phase-isolated. Each phase compares its own
  // edit-time column (Exercises Last Edited At / Weight Edited At) against
  // the value captured at the start of the pass. onEditMarkDirty advances
  // them only on the matching column class, so this catches edits to the
  // right phase regardless of whether the cell's value actually changed
  // (covers the "edit value, edit back" case that a value-comparison guard
  // would miss). The two phases are independent: a concurrent weight edit
  // doesn't defer the exercise stamp and vice versa.
  let exerciseConcurrentEdit = false;
  if (exerciseReady && cols.exercisesLastEditedAtCol) {
    const currentEdit = toDate_(getSheet_().getRange(row.rowNum, cols.exercisesLastEditedAtCol).getValue());
    const previousMs = row.exercisesLastEditedAt ? row.exercisesLastEditedAt.getTime() : null;
    const currentMs = currentEdit ? currentEdit.getTime() : null;
    if (currentMs !== previousMs) {
      console.info(tag + ': concurrent exercise edit detected (Exercises Last Edited At '
        + humanizeDate_(row.exercisesLastEditedAt) + ' -> '
        + (currentEdit ? humanizeDate_(currentEdit) : '<cleared>')
        + '); deferring Exercise Synced At stamp.');
      exerciseConcurrentEdit = true;
    }
  }
  let weightConcurrentEdit = false;
  if (weightReady && cols.weightEditedAtCol) {
    const currentEdit = toDate_(getSheet_().getRange(row.rowNum, cols.weightEditedAtCol).getValue());
    const previousMs = row.weightEditedAt ? row.weightEditedAt.getTime() : null;
    const currentMs = currentEdit ? currentEdit.getTime() : null;
    if (currentMs !== previousMs) {
      console.info(tag + ': concurrent weight edit detected (Weight Edited At '
        + humanizeDate_(row.weightEditedAt) + ' -> '
        + (currentEdit ? humanizeDate_(currentEdit) : '<cleared>')
        + '); deferring Weight Synced At stamp.');
      weightConcurrentEdit = true;
    }
  }

  const stampIso = new Date().toISOString();
  if (weightReady && !weightFailed && !weightConcurrentEdit) {
    markRowWeightSynced(row.rowNum, cols.weightSyncedAtCol, stampIso);
  }
  if (exerciseReady && !exerciseFailed && !exerciseConcurrentEdit) {
    markRowExerciseSynced(row.rowNum, cols.exerciseSyncedAtCol, stampIso);
  }

  if (weightFailed || exerciseFailed) {
    console.warn(tag + ': FAILED (partial); will retry on next sync.');
    return false;
  }

  // If the row still has unstamped phases (either because we skipped a phase
  // this pass or a concurrent edit blocked the stamp), advance the dirty
  // generation so syncDirtyRows' end-of-pass check leaves the flag set
  // (and a future poll picks the row up).
  const weightStampMissing = !row.weightSyncedAt && !(weightReady && !weightConcurrentEdit);
  const exerciseStampMissing = !row.exerciseSyncedAt && !(exerciseReady && !exerciseConcurrentEdit);
  if (weightStampMissing || exerciseStampMissing) {
    markPendingDirty_();
    console.info(tag + ': partial progress; row stays dirty (weightStamped='
      + (!weightStampMissing) + ', exerciseStamped=' + (!exerciseStampMissing) + ')');
  } else {
    console.info(tag + ': done');
  }
  return true;
}
