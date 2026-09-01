function setup() {
  ensureManagedColumns();
  installTriggers();
}

function onOpen() {
  console.info("onOpen: installing Sync menu");
  SpreadsheetApp.getUi()
    .createMenu("Sync")
    .addItem("Run now", "runSyncNow")
    .addItem("Resync selected rows", "resyncSelectedRows")
    .addItem("Resync all rows", "resyncAllRows")
    .addSeparator()
    .addItem("Run setup", "setup")
    .addItem("Authorize Health API", "authorizeHealthApi")
    .addItem("Revoke Health API", "revokeHealthApi")
    .addSeparator()
    // Both suites, via runAllTests in test/Harness.gs. The orchestration one
    // runs against the fakes there, which withSyncTestHarness_ swaps in for the
    // real services and takes back out again. Each suite reports in its own
    // alert.
    .addItem("Run tests", "runAllTests")
    .addToUi();
}

function authorizeHealthApi() {
  const ui = SpreadsheetApp.getUi();
  let service;
  try {
    service = getHealthService();
  } catch (err) {
    toast_(`Setup needed: ${String(err.message || err)}`, 30);
    return;
  }
  if (service.hasAccess()) {
    toast_(
      'Already authorized. Use "Revoke" first if you want to re-authorize.',
      10,
    );
    return;
  }
  const url = service.getAuthorizationUrl();
  console.log(`Authorize Google Health API: ${url}`);
  const html =
    `<p>Click the link below, sign in with the Google account that owns your Google Health ` +
    `data, and grant the requested scopes.</p>` +
    `<p><a href="${url}" target="_blank" rel="noopener">Authorize Google Health API</a></p>` +
    `<p>After the success page appears, you can close that tab and return here.</p>`;
  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(220),
    "Authorize Google Health",
  );
}

function revokeHealthApi() {
  resetHealthAuth();
  toast_(
    'Google Health authorization cleared. Use "Authorize Health API" to grant again.',
    10,
  );
}

function installTriggers() {
  const handlers = new Set(["syncOnEdit", "flushPending", "backstop"]);
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (handlers.has(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("syncOnEdit").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("flushPending")
    .timeBased()
    .everyMinutes(POLL_INTERVAL_MIN)
    .create();
  ScriptApp.newTrigger("backstop")
    .timeBased()
    .everyHours(BACKSTOP_INTERVAL_HOURS)
    .create();
}

// Return the first date violation among `datedRows` ([{ rowNum, date }] in
// sheet order) as a human-readable string, or null when valid. Comparison is
// by civil-date key (ymd, script time zone), so two rows on the same civil
// day count as duplicates even if their Date cells carry different times.
// Rows without a parseable Date are skipped by the caller, matching readRows,
// so they never sync and cannot place a datapoint at a bogus time. Pure.
function findRowDateViolation_(datedRows) {
  let prev = null;
  for (let i = 0; i < datedRows.length; i++) {
    const r = datedRows[i];
    const key = ymd(r.date);
    const year = Number(key.slice(0, 4));
    if (year < MIN_ROW_DATE_YEAR || year > MAX_ROW_DATE_YEAR) {
      return `row ${r.rowNum}: date ${key} is outside the allowed years ${
        MIN_ROW_DATE_YEAR
      }-${MAX_ROW_DATE_YEAR}`;
    }
    if (prev) {
      if (key === prev.key) {
        return `rows ${prev.rowNum} and ${r.rowNum} share the date ${key}`;
      }
      if (key < prev.key) {
        return `row ${r.rowNum} (${key}) is dated before row ${
          prev.rowNum
        } (${prev.key}); rows must be in increasing date order`;
      }
    }
    prev = { key, rowNum: r.rowNum };
  }
  return null;
}

// Read the Date column and return the first row-date violation (a
// human-readable string) or null. Every trigger runs this before doing any
// other work; only the one column is read, so the per-trigger cost is a
// single range read and no Health API calls. Deliberately does NOT route
// through readRows: its structural throws (missing Weight column, duplicate
// column headers) belong to syncDirtyRows' unrecoverable handling, and a
// missing Date column is likewise left for that path rather than reported as
// a validation failure here.
function validateRowDates_() {
  const sheet = getSheet_();
  const dateCol = getHeaderMap_(sheet).map[DATE_COLUMN_HEADER];
  if (!dateCol) {
    return null;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
  const datedRows = [];
  values.forEach((v, idx) => {
    if (!v[0]) {
      return;
    }
    const date = toDate_(v[0]);
    if (date) {
      datedRows.push({ date, rowNum: idx + 2 });
    }
  });
  return findRowDateViolation_(datedRows);
}

// Shared date-validation guard for the time-based triggers (vs. syncOnEdit's
// uncaught throw): they fire on a schedule, so throwing here would email the
// owner every cycle until the sheet is fixed: log-and-skip instead. Returns
// true when the trigger should skip its run. The dirty flag is left
// untouched, so any backlog syncs on the first run after the fix.
function dateValidationBlocksTrigger_(triggerName) {
  const violation = validateRowDates_();
  if (!violation) {
    return false;
  }
  console.error(
    `${triggerName}: date validation failed; skipping: ${violation}`,
  );
  return true;
}

// Date-validation guard for the manual Sync-menu entry points: the user is
// present, so surface the violation as a toast (not an email or a buried
// log) and abort before any stamps are cleared or datapoints written.
// Returns true when the action should abort.
function dateValidationBlocksManual_() {
  const violation = validateRowDates_();
  if (!violation) {
    return false;
  }
  console.error(`date validation failed: ${violation}`);
  toast_(`Date validation failed: ${violation}`, 30);
  return true;
}

function syncOnEdit(e) {
  // Date validation runs first and OUTSIDE the catch-all below: a violation
  // is thrown uncaught so Apps Script emails the owner about the failed
  // trigger execution: the edit that broke the rule is the moment to alarm.
  // The time-based triggers (flushPending, backstop) log-and-skip instead so
  // a standing violation doesn't email every cycle. Note the edit is NOT
  // dirty-marked on this path (the throw precedes onEditMarkDirty), so
  // content typed while the sheet is invalid does not sync by itself once
  // the dates are fixed. Recovery: exercise rows within
  // BACKSTOP_LOOKBACK_DAYS self-heal: the next backstop re-dirties them and
  // the following poll syncs them. Older exercise rows and weight edits need
  // a re-edit of the cell or "Resync selected rows".
  const violation = validateRowDates_();
  if (violation) {
    throw new Error(`syncOnEdit: date validation failed: ${violation}`);
  }
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
    console.error(`syncOnEdit error: ${err}`);
  }
}

function flushPending() {
  // Sync only when there is pending work. No foreign-match re-review here: the
  // backstop re-dirties recent exercise rows (matched AND unmatched) for
  // foreign alignment, so the 5-min poll issues no Health API calls unless an
  // onEdit (or manual sync) left the dirty flag set. This keeps steady-state
  // read traffic off the API: a row with no pending edits is not re-queried
  // every 5 minutes.
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(PENDING_DIRTY_KEY)) {
    console.info("flushPending: no pending edits, skipping");
    return;
  }
  // Validation sits after the fast-path return above so an idle poll stays a
  // single PropertiesService read (no sheet I/O); nothing can sync without
  // pending work, so the late check gates every sync all the same.
  if (dateValidationBlocksTrigger_("flushPending")) {
    return;
  }
  console.info("flushPending: pending edits detected, syncing");
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
  recentCivilDates_(nowMs, lookbackDays).forEach((d) => {
    recentKeys.add(ymd(d));
  });
  // Order by cost: the free matched-state test, then the ymd date-key lookup
  // (a timezone op), then the nested-loop hasSendableExercises_ last so it runs
  // only on rows that already qualify.
  return rows.filter(
    (r) =>
      (wantMatched
        ? r.matchedHealthSessions.length > 0
        : r.matchedHealthSessions.length === 0) &&
      recentKeys.has(ymd(r.date)) &&
      hasSendableExercises_(r.exercises),
  );
}

// Rows whose recorded Health datapoints contradict their current cell content:
// an exercise datapoint tracked for a row with no sendable exercise content, or
// a weight datapoint tracked for a row with no bodyweight. Returns
// { exerciseRowNums, weightRowNums }.
//
// This is the reconciliation path for CLEARED content, and it is deliberately
// state-based rather than event-based. onEditMarkDirty can only guess whether a
// multi-cell edit cleared something (Apps Script supplies `oldValue` for single
// cells only), and every guess has boundary cases that silently skip the delete.
// The state here has no ambiguity: the sheet says one thing, Created Health IDs
// says another, and exactly one of them is right. It therefore catches a clear
// however the user made it: single cell, multi-cell, paste, or a mixed range
// that blanks one cell while writing another.
//
// Only rows currently STAMPED synced are returned; an already-dirty row is
// picked up by the next pass regardless, so re-dirtying it would just be an
// extra write. Self-limiting: the sync drops the id, so the row stops matching.
//
// This DELETES data, so emptiness is read from the raw cells (hasExerciseText /
// hasWeightText) and not from the parse result. The two are different claims:
// an unparseable cell parses to nothing while plainly still holding the user's
// data, so treating "the parser produced nothing" as "the user cleared it"
// turns any parser or schema change into a mass deletion across all history.
// The caller bounds the blast radius further, via STALE_RECONCILE_MAX_ROWS.
//
// Not covered: a row whose Date is blank, which readRows drops entirely. Such a
// row keeps its datapoints by design (see readRows), so its content is not
// authoritative here either. Pure (no API/sheet access).
function selectStaleDataPointRows_(rows) {
  const exerciseRowNums = [];
  const weightRowNums = [];
  rows.forEach((r) => {
    const split = splitHealthIdsByType_(r.healthIds);
    if (
      r.exerciseSyncedAt &&
      split.exercise.length > 0 &&
      !hasSendableExercises_(r.exercises) &&
      !r.hasExerciseText
    ) {
      exerciseRowNums.push(r.rowNum);
    }
    if (
      r.weightSyncedAt &&
      split.weight.length > 0 &&
      r.bodyweight === null &&
      !r.hasWeightText
    ) {
      weightRowNums.push(r.rowNum);
    }
  });
  return { exerciseRowNums, weightRowNums };
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
  if (dateValidationBlocksTrigger_("backstop")) {
    return;
  }
  // Take the script lock for the whole run. Orphan reconciliation deletes
  // exercise datapoints no row references; without the lock it could race an
  // in-flight syncOneRow_ that has POSTed a datapoint but not yet persisted its
  // ID, and delete a legitimately-fresh datapoint. The lock also guarantees the
  // readRows snapshot reflects every persisted ID, so the "known" set is
  // complete. If a sync holds the lock, skip: the next backstop run retries.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    console.warn("backstop: another run holds the lock; skipping this run.");
    return;
  }
  try {
    const now = Date.now();
    const { allHealthIds, exerciseSyncedAtCol, rows, weightSyncedAtCol } =
      readRows();

    // Reconcile orphans first, while the rows snapshot is unmodified. (reDirty
    // below doesn't touch Created Health IDs, so order is not load-bearing, but
    // reconciling against the clean snapshot keeps the intent obvious.)
    // Both take allHealthIds, the full-sheet id list, which includes rows
    // readRows drops for a blank/unparseable Date.
    try {
      reconcileExerciseOrphans_(
        allHealthIds,
        now,
        ORPHAN_RECONCILE_LOOKBACK_DAYS,
      );
    } catch (err) {
      console.warn(
        `backstop: exercise orphan reconciliation failed (continuing): ${err}`,
      );
    }
    try {
      reconcileWeightOrphans_(
        allHealthIds,
        now,
        ORPHAN_RECONCILE_LOOKBACK_DAYS,
      );
    } catch (err) {
      console.warn(
        `backstop: weight orphan reconciliation failed (continuing): ${err}`,
      );
    }

    if (!exerciseSyncedAtCol) {
      console.warn(
        'backstop: Exercise Synced At column missing; run "Run setup".',
      );
      return;
    }
    // Re-dirty both matched and unmatched recent rows (a row is exactly one of
    // the two, so the union needs no dedup).
    const unmatched = selectBackstopRows_(
      rows,
      now,
      BACKSTOP_LOOKBACK_DAYS,
      false,
    );
    const matched = selectBackstopRows_(
      rows,
      now,
      BACKSTOP_LOOKBACK_DAYS,
      true,
    );
    // Plus rows whose recorded datapoints contradict their content, i.e. the
    // user cleared cells. Deliberately NOT bounded by BACKSTOP_LOOKBACK_DAYS:
    // the scan is pure sheet state with no API calls, a clear on an old row is
    // exactly the case nothing else recovers, and the set is normally empty.
    // Disjoint from the foreign-match selections above by construction (those
    // require sendable exercise content, the stale exercise set requires none),
    // so the concat needs no dedup.
    let stale = selectStaleDataPointRows_(rows);
    // Bound the destructive branch. Clearing cells is human-scale; a stale set
    // this large is evidence of a systemic change (a column deleted, a bulk
    // reformat) where reconciling would destroy history rather than repair it.
    // Log and reconcile NOTHING rather than part of it, so the sheet is left
    // exactly as found for a human to inspect. Foreign re-review still runs.
    const staleCount =
      stale.exerciseRowNums.length + stale.weightRowNums.length;
    if (staleCount > STALE_RECONCILE_MAX_ROWS) {
      console.error(
        `backstop: ${staleCount} rows look cleared (limit ${
          STALE_RECONCILE_MAX_ROWS
        }). That is more than a person clears by hand, so this is` +
          ` probably a column or format change rather than a clear. Reconciling nothing;` +
          ` use "Resync selected rows" if the change really was intended.`,
      );
      stale = { exerciseRowNums: [], weightRowNums: [] };
    }
    const exerciseTargets = unmatched
      .concat(matched)
      .map((r) => r.rowNum)
      .concat(stale.exerciseRowNums);
    const weightTargets = stale.weightRowNums;
    if (exerciseTargets.length === 0 && weightTargets.length === 0) {
      console.info(
        "backstop: no exercise rows to re-review and no stale datapoints.",
      );
      return;
    }
    if (exerciseTargets.length > 0) {
      reDirtyRows_(exerciseTargets, { exerciseCol: exerciseSyncedAtCol });
    }
    if (weightTargets.length > 0) {
      reDirtyRows_(weightTargets, { weightCol: weightSyncedAtCol });
    }
    console.info(
      `backstop: re-dirtied ${exerciseTargets.length} exercise row(s) (${
        unmatched.length
      } unmatched, ${matched.length} matched for foreign re-review, ${
        stale.exerciseRowNums.length
      } stale) and ${weightTargets.length} stale weight row(s).`,
    );
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
// "Our client" is derived from the candidates themselves (the
// googleWebClientId of any candidate that IS tracked) rather than from a
// configured value, so we only ever delete datapoints from the exact client
// that created the ones we still track. Foreign device / first-party / in-app-
// assistant sessions (null googleWebClientId) and any other web app are never
// selected. If no tracked candidate is present we can't attribute ownership,
// so nothing is returned. Type-agnostic and pure (no API/sheet access): used
// for both exercise and weight reconciliation.
function selectOrphanDataPointNames_(candidates, knownNames) {
  const ourClientIds = {};
  candidates.forEach((c) => {
    if (c.googleWebClientId && knownNames[c.name]) {
      ourClientIds[c.googleWebClientId] = true;
    }
  });
  const orphans = [];
  candidates.forEach((c) => {
    if (knownNames[c.name]) {
      return;
    }
    if (c.googleWebClientId && ourClientIds[c.googleWebClientId]) {
      orphans.push(c.name);
    }
  });
  return orphans;
}

// Delete sync-created datapoints of one data type that no row's Created Health
// IDs references: orphans leaked by the two accepted create windows (a create
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
//
// `allHealthIds` must be readRows' full-sheet id list, NOT the ids reachable
// from its `rows`: a row whose Date is blank or unparseable is absent from
// `rows` while its datapoints are still live, and reconciling against the
// narrower set would delete them.
function reconcileDataPointOrphans_(
  allHealthIds,
  nowMs,
  lookbackDays,
  typeKey,
  listOnDate,
) {
  const tag = `reconcileDataPointOrphans_(${typeKey})`;
  const known = {};
  splitHealthIdsByType_(allHealthIds)[typeKey].forEach((n) => {
    known[n] = true;
  });

  const candidates = [];
  recentCivilDates_(nowMs, lookbackDays).forEach((date) => {
    try {
      listOnDate(date).forEach((c) => candidates.push(c));
    } catch (err) {
      console.warn(`${tag}: list failed for ${ymd(date)}: ${err}`);
    }
  });

  const orphans = selectOrphanDataPointNames_(candidates, known);
  if (orphans.length === 0) {
    console.info(
      `${tag}: no orphans found (${candidates.length} datapoint(s) scanned).`,
    );
    return;
  }
  let deleted = 0;
  orphans.forEach((name) => {
    try {
      deleteDataPointsByName([name]);
      deleted++;
      console.info(`${tag}: deleted orphan ${name}`);
    } catch (err) {
      console.warn(`${tag}: delete failed for ${name}: ${err}`);
    }
  });
  console.info(
    `${tag}: removed ${deleted} of ${orphans.length} orphan datapoint(s).`,
  );
}

// Reconcile orphaned sync-created exercise datapoints (STRENGTH_TRAINING).
function reconcileExerciseOrphans_(allHealthIds, nowMs, lookbackDays) {
  reconcileDataPointOrphans_(
    allHealthIds,
    nowMs,
    lookbackDays,
    "exercise",
    listStrengthOnDate,
  );
}

// Reconcile orphaned sync-created weight datapoints. The exercise path's
// create-orphan windows (timed-out-but-succeeded POST retry; hard kill after
// POST before the ID is persisted) apply to the weight POST in syncOneRow_ too,
// so a duplicate untracked weight datapoint can leak the same way. Same
// ownership logic: only datapoints from our web client, referenced by no row,
// are deleted.
function reconcileWeightOrphans_(allHealthIds, nowMs, lookbackDays) {
  reconcileDataPointOrphans_(
    allHealthIds,
    nowMs,
    lookbackDays,
    "weight",
    listWeightOnDate,
  );
}

// Write a fresh generation marker into PENDING_DIRTY_KEY. The value matters
// (syncDirtyRows compares start vs end to detect concurrent edits), so always
// advance it: never just re-write the same string.
function markPendingDirty_() {
  PropertiesService.getScriptProperties().setProperty(
    PENDING_DIRTY_KEY,
    String(Date.now()),
  );
}

// Re-dirty a set of rows: clear the requested synced stamp(s), persist the
// writes, and advance the dirty generation so the next sync pass re-processes
// them. Shared by the manual selective resync and the backstop; pass
// cols.exerciseCol and/or cols.weightCol for the phase(s) to re-dirty. (The
// "resync all rows" path clears its columns in a single batched setValues
// instead, so it doesn't route through here.)
function reDirtyRows_(rowNums, cols) {
  rowNums.forEach((rowNum) => {
    if (cols.exerciseCol) {
      clearRowExerciseSynced(rowNum, cols.exerciseCol);
    }
    if (cols.weightCol) {
      clearRowWeightSynced(rowNum, cols.weightCol);
    }
  });
  SpreadsheetApp.flush();
  markPendingDirty_();
}

// "Run now" is an explicit manual action: syncs all dirty rows immediately,
// waiting for the script lock (unlike the automatic triggers, which skip when
// the lock is held). If the user keeps editing afterward, the row goes dirty
// again and the next sync reconciles the Health datapoint(s).
function runSyncNow() {
  if (dateValidationBlocksManual_()) {
    return;
  }
  runSyncAndToast_("Synced");
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
    toast_(`Sync failed: ${String(err.message || err)}`, 30);
    throw err;
  }
  toastSyncResult_(result, verb);
}

function resyncSelectedRows() {
  if (dateValidationBlocksManual_()) {
    return;
  }
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  const exerciseCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER];
  if (!exerciseCol) {
    toast_("Exercise Synced At column missing. Run setup.", 30);
    return;
  }
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER];
  if (!weightSyncedAtCol) {
    toast_("Weight Synced At column missing. Run setup.", 30);
    return;
  }
  // The selection is spreadsheet-global: it belongs to whichever tab is
  // focused, not necessarily the synced one (always the first tab, per
  // getSheet_). Re-dirtying by row number from another tab's selection would
  // hit unrelated rows here, so require the synced tab to be active.
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (activeSheet.getSheetId() !== sheet.getSheetId()) {
    toast_(`Select rows on the "${sheet.getName()}" tab first.`, 10);
    return;
  }
  const rows = {};
  // Null when nothing is selected at all.
  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    toast_("No data rows selected.", 10);
    return;
  }
  const ranges = rangeList.getRanges();
  // Clamp to the data range. A whole-column selection (clicking the column
  // header) reports the full sheet height, which would re-dirty every empty row
  // below the data and spend the pass on thousands of stamp writes.
  const lastDataRow = sheet.getLastRow();
  for (const range of ranges) {
    const start = range.getRow();
    const end = Math.min(start + range.getNumRows() - 1, lastDataRow);
    for (let row = start; row <= end; row++) {
      if (row >= 2) {
        rows[row] = true;
      }
    }
  }
  const rowNums = Object.keys(rows);
  if (rowNums.length === 0) {
    toast_("No data rows selected.", 10);
    return;
  }
  reDirtyRows_(rowNums.map(Number), {
    exerciseCol,
    weightCol: weightSyncedAtCol,
  });
  runSyncAndToast_("Resynced");
}

function resyncAllRows() {
  if (dateValidationBlocksManual_()) {
    return;
  }
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  const exerciseCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER];
  if (!exerciseCol) {
    toast_("Exercise Synced At column missing. Run setup.", 30);
    return;
  }
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER];
  if (!weightSyncedAtCol) {
    toast_("Weight Synced At column missing. Run setup.", 30);
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    toast_("No data rows.", 10);
    return;
  }
  const dataRowCount = lastRow - 1;

  const blanks = [];
  for (let i = 0; i < dataRowCount; i++) {
    blanks.push([""]);
  }
  sheet.getRange(2, exerciseCol, dataRowCount, 1).setValues(blanks);
  sheet.getRange(2, weightSyncedAtCol, dataRowCount, 1).setValues(blanks);
  SpreadsheetApp.flush();
  markPendingDirty_();

  runSyncAndToast_("Resynced");
}

function formatSyncResult_(result, verb) {
  if (!result) {
    return "Sync skipped (another run holds the lock). Try again shortly.";
  }
  let msg = `${verb} ${result.ok} row(s)`;
  if (result.errors > 0) {
    msg += `, ${result.errors} error(s)`;
  }
  if (result.deferred > 0) {
    msg += `, ${result.deferred} deferred`;
  }
  msg += ".";
  if (result.errors > 0) {
    msg += "\n\nSee Executions for details.";
  }
  return msg;
}

// Non-blocking status notification. Apps Script's ui.alert() is modal and
// counts against the 6-minute execution budget, which caused timeouts on
// large resyncs. Toasts auto-dismiss and don't block the script.
function toast_(msg, seconds) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Sync", seconds);
}

function toastSyncResult_(result, verb) {
  const seconds = result && result.errors > 0 ? 30 : 10;
  toast_(formatSyncResult_(result, verb), seconds);
}

function humanizeMs_(ms) {
  const elapsed = ms < 0 ? 0 : ms;
  if (elapsed < 1000) {
    return `${elapsed}ms`;
  }
  const totalSec = Math.round(elapsed / 1000);
  if (totalSec < 60) {
    return `${totalSec}s`;
  }
  const min = Math.floor(totalSec / 60);
  const sec = totalSec % 60;
  return sec === 0 ? `${min}m` : `${min}m ${sec}s`;
}

function humanizeDate_(date) {
  if (!date) {
    return "<none>";
  }
  return Utilities.formatDate(date, getTz_(), "yyyy-MM-dd HH:mm:ss");
}

function syncDirtyRows(lockWaitMs) {
  const lock = LockService.getScriptLock();
  const waitMs =
    lockWaitMs === undefined || lockWaitMs === null ? LOCK_WAIT_MS : lockWaitMs;
  if (!lock.tryLock(waitMs)) {
    if (waitMs > 0) {
      console.warn("syncDirtyRows: another run holds the lock; skipping.");
    } else {
      console.info(
        "syncDirtyRows: another run holds the lock; skipping this tick.",
      );
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
  // Set when the pass hits a misconfiguration retrying can't fix (missing
  // columns, duplicate headers), or when the row loop's summary throw reports
  // rows that failed unexpectedly. The catch emails the owner; the finally
  // LEAVES the dirty flag set (it skips flag resolution entirely on this path)
  // so the backlog syncs itself on the next poll once the cause clears.
  // Clearing it would orphan already-dirty rows until some future edit. The
  // summary throw depends on that, which is why it sets the flag first.
  let unrecoverable = false;
  try {
    const {
      allHealthIds,
      allMatchedSessions,
      dateCol,
      exerciseSyncedAtCol,
      exercisesLastEditedAtCol,
      healthIdsCol,
      matchedHealthSessionCol,
      rows,
      weightEditedAtCol,
      weightSyncedAtCol,
    } = readRows();
    if (!exerciseSyncedAtCol || !weightSyncedAtCol || !healthIdsCol) {
      throw new Error(
        'Managed columns missing; run "Run setup" from the Sync menu before syncing.',
      );
    }
    const dirty = rows.filter((r) => !r.exerciseSyncedAt || !r.weightSyncedAt);
    if (dirty.length === 0) {
      return { errors: 0, ok: 0 };
    }

    // Per-row phase readiness (no debounce: onEdit syncs immediately and the
    // poll/backstop retry): a row's weight phase is ready when weight-dirty, its
    // exercise phase when exercise-dirty. There is no "still typing" wait; a
    // poll firing mid-burst just re-pushes current state, and the idempotent
    // exercise re-sync (syncOneRow_) skips the recreate when nothing changed.
    const ready = [];
    dirty.forEach((r) => {
      const weightReady = !r.weightSyncedAt;
      const exerciseReady = !r.exerciseSyncedAt;
      if (weightReady || exerciseReady) {
        ready.push({ exerciseReady, row: r, weightReady });
      }
    });

    if (ready.length === 0) {
      console.info("syncDirtyRows: no rows ready to sync.");
      return { errors: 0, ok: 0 };
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
      console.info(
        `syncDirtyRows: ${ready.length} ready row(s); capping at ${
          MAX_ROWS_PER_SYNC
        }, deferring ${deferredCount} to next pass`,
      );
      ready.length = MAX_ROWS_PER_SYNC;
    } else {
      console.info(`syncDirtyRows: ${ready.length} ready row(s)`);
    }

    // Every exercise-dirty row creates its own datapoint; the plan only
    // supplies a foreign session's interval to align timing when one overlaps
    // the row's edit window. See FOREIGN_MATCH_BUFFER_MS in Config.gs.
    const exerciseReadyRows = ready
      .filter((r) => r.exerciseReady)
      .map((r) => r.row);
    const alignmentPlan = resolveForeignMatches_(
      allHealthIds,
      allMatchedSessions,
      exerciseReadyRows,
    );
    const cols = {
      dateCol,
      exerciseSyncedAtCol,
      exercisesLastEditedAtCol,
      healthIdsCol,
      matchedHealthSessionCol,
      weightEditedAtCol,
      weightSyncedAtCol,
    };
    // A throw out of syncOneRow_ is unexpected: it catches its own Health API
    // failures, so what reaches here is sheet I/O failing on a transient
    // Spreadsheets service error: the final writeHealthIds /
    // writeMatchedHealthSession, the concurrent-edit guard reads, and the stamp
    // writes. The two persist-immediately writeHealthIds calls sit inside the
    // create try/catch blocks, so a failure there is recorded as
    // weightFailed/exerciseFailed, but that does NOT contain it: the same
    // transient condition then fails the unconditional write at the end of the
    // row, which does reach this catch, with the just-created datapoint
    // untracked (an orphan for reconciliation to reclaim).
    //
    // Isolate it per row rather than letting it abort the loop. Rows are
    // processed newest-first, so aborting would leave every older row behind
    // the failure unsynced, and a row that throws deterministically would do
    // so on every subsequent pass, wedging the backlog indefinitely. The pass
    // therefore always runs to the end: no failure count stops it. (Stopping
    // early to limit orphans was tried and removed: every threshold either
    // failed to bound them, since the next pass re-creates what this one
    // couldn't record, or re-created the wedge.)
    //
    // Reconciliation is the backstop, but only within
    // ORPHAN_RECONCILE_LOOKBACK_DAYS: it lists candidates by the DATAPOINT's
    // own civil date, which is the row's Date, not its creation time. So
    // orphans leaked for rows dated further back than that lookback (the
    // resyncAllRows-over-history case) are never listed and never reclaimed.
    // That is the accepted residual, not something the loop covers.
    // The failures are still reported: the summary throw below routes them
    // through the unrecoverable path so the owner is emailed and the dirty flag
    // is kept, but only after every other ready row has synced and stamped.
    const unexpected = [];
    for (let i = 0; i < ready.length; i++) {
      const entry = ready[i];
      const foreignMatches = alignmentPlan[entry.row.rowNum] || [];
      try {
        if (
          syncOneRow_(
            entry.row,
            foreignMatches,
            entry.weightReady,
            entry.exerciseReady,
            cols,
            i + 1,
            ready.length,
          )
        ) {
          ok++;
        } else {
          errors++;
        }
      } catch (err) {
        errors++;
        unexpected.push(`row ${entry.row.rowNum}: ${err}`);
        console.error(
          `syncDirtyRows: unexpected error on row ${
            entry.row.rowNum
          }; skipping it and continuing the pass: ${err}`,
        );
      }
    }
    if (unexpected.length > 0) {
      // Set the flag before throwing. The unrecoverable path skips end-of-pass
      // flag resolution entirely (it only leaves whatever is already there), so
      // a pass entered with no flag set (runSyncNow is the one entry point that
      // doesn't set one) would strand these rows dirty with nothing to retry
      // them: flushPending short-circuits on its fast path, and the backstop
      // only re-dirties exercise rows with sendable content.
      markPendingDirty_();
      // Carry the pass outcome in the message. This throw replaces the normal
      // return, so the ok/errors/deferred counts would otherwise be lost, and
      // on the manual entry points runSyncAndToast_ shows this text, where
      // "79 synced" is the difference between "the resync worked apart from one
      // row" and an unqualified "Sync failed".
      //
      // Counts go FIRST, and only the first few rows are named: the toast
      // truncates, so with a broad failure (up to MAX_ROWS_PER_SYNC rows) a
      // trailing summary would be cut off, exactly the outcome the summary
      // exists to prevent. Nothing is lost by trimming the list: every failure
      // was logged individually above and lands in Executions.
      const MAX_NAMED_FAILURES = 5;
      const named = unexpected.slice(0, MAX_NAMED_FAILURES);
      const unnamed = unexpected.length - named.length;
      throw new Error(
        `${ok} synced, ${errors} error(s)${
          deferredCount > 0 ? `, ${deferredCount} deferred by the row cap` : ""
        }. Unexpected per-row failure(s): ${named.join("; ")}${
          unnamed > 0 ? `; +${unnamed} more (see Executions)` : ""
        }.`,
      );
    }
  } catch (err) {
    // Unrecoverable: a throw out of the sync body (missing required columns,
    // duplicate column headers, etc.), or the summary throw for rows whose
    // sheet I/O failed unexpectedly. Per-row API failures never reach here,
    // since syncOneRow_ catches them and they retry. Re-throw so the failure
    // propagates uncaught: Apps Script then emails the script owner about the
    // failed trigger execution (no MailApp needed), and manual entry points
    // still toast and land the error in Executions. The finally LEAVES the
    // dirty flag set (see below) so the next poll retries automatically once
    // the misconfig is fixed: clearing it would orphan already-dirty rows
    // until some future edit. Repeated failure emails are throttled via the
    // trigger's notification cadence (Apps Script ▸ Triggers ▸ notifications),
    // not by suppressing the retry.
    unrecoverable = true;
    console.error(`syncDirtyRows: unrecoverable error: ${err}`);
    throw err;
  } finally {
    // NB: no `return` in this finally. A finally-return would swallow the
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
        if (!concurrentEdit) {
          markPendingDirty_();
        }
      } else if (!concurrentEdit) {
        props.deleteProperty(PENDING_DIRTY_KEY);
      }
      // Bookend the pass: outcome counts plus whether the dirty flag survives
      // (so a reader knows if another poll will follow without inspecting it).
      const willRetry = workRemaining || concurrentEdit;
      console.info(
        `syncDirtyRows: pass complete, ${ok} synced, ${errors} error(s)${
          deferredCount ? `, ${deferredCount} deferred` : ""
        }${
          concurrentEdit ? ", concurrent edit landed" : ""
        }${willRetry ? "; pending flag kept, next poll will retry." : "; queue drained."}`,
      );
    }
    lock.releaseLock();
  }
  return { deferred: deferredCount, errors, ok };
}

// Returns rowNum -> foreign Strength Training session whose interval should be
// copied onto the row's created exercise datapoint (timing alignment). The
// sync never skips its own create; this only supplies a more accurate
// start/end when a manually-logged foreign session overlaps the row's edit
// window. Time-overlap only: rows without same-date exercise edit timestamps
// get no match and fall through to synthetic/prior timing (there is no
// positional fallback).
//
// Candidate exclusions (global, keyed by resource name: names are globally
// unique, so no date keying is needed, and global keying is what lets a
// neighbor-day candidate be excluded correctly):
//   - sync-created: name appears in `allHealthIds` (our own datapoint, never
//     align to ourselves, including a row's prior datapoint on re-sync). This
//     is readRows' full-sheet id list rather than the ids reachable from the
//     pass's `rows` snapshot, so a row dropped for a blank/unparseable Date
//     still shields its datapoints from being borrowed as "foreign".
//   - aligned-elsewhere: name is a non-ready row's Matched Health Session, so
//     two rows don't borrow the same foreign session's times across runs. Also
//     full-sheet (`allMatchedSessions`) rather than row-derived: a row dropped
//     for a blank Date can never be ready, so the session it borrowed has to
//     stay excluded.
//
// Cross-date: candidates are gathered for every civil date any row's window
// touches (a window near local midnight pulls the neighbor day too), deduped
// by name. Overlap is computed in absolute UTC, so midnight-crossing workouts
// match regardless of which civil date the foreign session was logged under.
//
// Ordering note: this runs once near the top of syncDirtyRows, before the
// per-row syncOneRow_ loop. Both exclusion lists come from readRows() at the
// start of the pass, and `listStrengthOnDate` calls the API before any row in
// this pass has had its create issued. So same-pass freshly-created datapoints
// are not yet visible to the API and not yet in Created Health IDs, so both gaps
// cancel out and there's no self-match risk.
function resolveForeignMatches_(allHealthIds, allMatchedSessions, readyRows) {
  const plan = {};
  const readyRowNums = {};
  readyRows.forEach((r) => {
    readyRowNums[r.rowNum] = true;
  });

  const tracked = {};
  splitHealthIdsByType_(allHealthIds).exercise.forEach((name) => {
    tracked[name] = true;
  });
  const excluded = {};
  Object.keys(tracked).forEach((name) => {
    excluded[name] = true;
  });
  allMatchedSessions.forEach((m) => {
    if (!readyRowNums[m.rowNum]) {
      excluded[m.name] = true;
    }
  });

  // Only rows with on-row-date edit timestamps anchor a trustworthy window.
  // capExerciseDurationToMax_ caps the window the same way resolveRowTiming_
  // caps the recorded interval, so a row whose exercisesLastEditedAt drifted
  // far past exerciseFirstEditedAt (late corrections that keep sticky
  // first-edit + advance last-edit) doesn't produce a multi-day window biased
  // toward the longest unrelated candidate.
  const windows = readyRows
    .filter(
      (r) => hasSendableExercises_(r.exercises) && exerciseEditIsOnRowDate_(r),
    )
    .map((r) => {
      const startMs = r.exerciseFirstEditedAt.getTime();
      const clampedEndMs =
        startMs +
        capExerciseDurationToMax_(r.exercisesLastEditedAt.getTime() - startMs);
      return {
        rowNum: r.rowNum,
        windowEnd: clampedEndMs + FOREIGN_MATCH_BUFFER_MS,
        windowStart: startMs - FOREIGN_MATCH_BUFFER_MS,
      };
    });
  if (windows.length === 0) {
    return plan;
  }

  // Gather candidates across every civil date a window touches (start and end
  // edges, which are adjacent when a window straddles midnight), deduped by name and
  // with our own / aligned-elsewhere names dropped.
  const probeDates = {};
  windows.forEach((w) => {
    const start = new Date(w.windowStart);
    const end = new Date(w.windowEnd);
    probeDates[ymd(start)] = start;
    probeDates[ymd(end)] = end;
  });
  // Our own client's datapoints are never candidates, tracked or not. The
  // `excluded` names cover what the sheet tracks, but an orphan leaked by the
  // accepted create windows is untracked, was created FROM a row's edit
  // window, and so often overlaps that window better than the real device
  // session: the row would borrow its own leak's interval until orphan
  // reconciliation removes it. Ownership is derived the way
  // selectOrphanDataPointNames_ derives it: the googleWebClientId carried by
  // any listed candidate the sheet tracks is ours (foreign device/first-party
  // sessions carry null and are unaffected).
  const byName = {};
  const ourClientIds = {};
  Object.keys(probeDates).forEach((key) => {
    let list;
    try {
      list = listStrengthOnDate(probeDates[key]);
    } catch (err) {
      console.warn(`resolveForeignMatches_: list failed for ${key}: ${err}`);
      return;
    }
    list.forEach((c) => {
      if (tracked[c.name] && c.googleWebClientId) {
        ourClientIds[c.googleWebClientId] = true;
      }
      if (excluded[c.name]) {
        return;
      }
      byName[c.name] = c;
    });
  });
  const candidates = Object.values(byName).filter(
    (c) => !(c.googleWebClientId && ourClientIds[c.googleWebClientId]),
  );
  if (candidates.length === 0) {
    return plan;
  }

  // Assign in rowNum order; a row claims EVERY remaining candidate its window
  // overlaps, and claimed candidates are removed so two rows can't align to the
  // same foreign session. All of them rather than the best one because a single
  // day can hold several app-recorded workouts and the sheet has one row per
  // date (findRowDateViolation_ forbids two rows sharing one), so the row's
  // exercises are split across the sessions they were logged during rather than
  // all landing on whichever overlapped most. Sorted by start time so the split
  // and the resulting datapoints run in workout order.
  windows.sort((a, b) => a.rowNum - b.rowNum);
  windows.forEach((w) => {
    const claimed = [];
    for (let i = candidates.length - 1; i >= 0; i--) {
      const c = candidates[i];
      const overlap =
        Math.min(c.endUtcMs, w.windowEnd) -
        Math.max(c.startUtcMs, w.windowStart);
      if (overlap > 0) {
        claimed.push(c);
        candidates.splice(i, 1);
      }
    }
    if (claimed.length > 0) {
      claimed.sort((a, b) => a.startUtcMs - b.startUtcMs);
      plan[w.rowNum] = claimed;
      console.info(
        `resolveForeignMatches_: row ${w.rowNum} aligns to ${
          claimed.length
        } session(s): ${claimed.map((c) => c.name).join(", ")}`,
      );
    }
  });
  return plan;
}

// Split a row's exercises across the app-recorded sessions its window claimed,
// by each exercise's OWN first-edit timestamp (Exercise Edit Times), so a day
// holding two workouts writes two datapoints with the right exercises on each
// rather than one carrying both.
//
// An exercise is attributed to the session whose interval contains its first
// edit, or, when none does, to the nearest session within
// FOREIGN_MATCH_BUFFER_MS of it: the same slack the window itself uses, since
// the first set of a workout is often typed a moment before the session is
// started in the app. `first` rather than `last` because it is sticky: a typo
// corrected during the next workout must not move the exercise into it.
//
// An exercise with no recorded first-edit time falls back to
// `fallbackFirstEdit`, the row-level Exercise First Edited At. That is what
// keeps a row written before Exercise Edit Times existed behaving exactly as it
// did: one session over the row's window still catches every exercise and the
// interval is still borrowed. Where the day held two workouts such a row puts
// everything in the first, which is the pre-split result and the most the
// row-level timestamp can support.
//
// Exercises outside every session's slack, with no fallback either, go into one
// trailing group with a null session, which the caller times from that group's
// own edit window (groupEditWindow_: min first-edit / max last-edit of its
// members, row-level fallback) through the usual resolver rules. Groups with
// nothing sendable are dropped: an app session that overlapped the window but
// caught no exercises produces no datapoint.
//
// With no sessions, or with one that catches everything, the result is a single
// group and the behavior is identical to the pre-split sync. Pure.
function partitionExercisesBySession_(
  sessions,
  exerciseEditTimes,
  fallbackFirstEdit,
  exercises,
) {
  const bySession = sessions.map((session) => ({ exercises: [], session }));
  const unattributed = [];
  exercises.forEach((ex) => {
    const times = exerciseEditTimes && exerciseEditTimes[ex.name];
    const first = (times && times.first) || fallbackFirstEdit;
    let bestIdx = -1;
    if (first) {
      const ms = first.getTime();
      let bestDistance = Infinity;
      sessions.forEach((sn, i) => {
        let distance = 0;
        if (ms < sn.startUtcMs) {
          distance = sn.startUtcMs - ms;
        } else if (ms > sn.endUtcMs) {
          distance = ms - sn.endUtcMs;
        }
        if (distance <= FOREIGN_MATCH_BUFFER_MS && distance < bestDistance) {
          bestDistance = distance;
          bestIdx = i;
        }
      });
    }
    if (bestIdx >= 0) {
      bySession[bestIdx].exercises.push(ex);
    } else {
      unattributed.push(ex);
    }
  });
  const groups = bySession.concat(
    unattributed.length > 0 ? [{ exercises: unattributed, session: null }] : [],
  );
  return groups.filter((g) => hasSendableExercises_(g.exercises));
}

// The edit window that describes ONE group of exercises rather than the whole
// row: the min first-edit / max last-edit of the group's own entries in
// Exercise Edit Times, with the row-level timestamps as the fallback for any
// exercise that has no entry (rows written before the column existed). Because
// the row-level first edit is never later than any per-exercise first (it was
// seeded by the earliest exercise edit), a group containing one legacy
// exercise degrades to the row-level window rather than inventing a narrower
// one. Used to time the null-session group after the attributed exercises are
// split away, so an exercise typed between two workouts is recorded at the
// time it was typed, not across both. Pure.
function groupEditWindow_(row, exerciseEditTimes, exercises) {
  let first = null;
  let last = null;
  const consider = (candidate, current, pick) => {
    if (!candidate) {
      return current;
    }
    if (!current) {
      return candidate;
    }
    return pick(candidate.getTime(), current.getTime()) ? candidate : current;
  };
  exercises.forEach((ex) => {
    const times = exerciseEditTimes && exerciseEditTimes[ex.name];
    const exFirst = (times && times.first) || row.exerciseFirstEditedAt;
    const exLast = (times && times.last) || row.exercisesLastEditedAt;
    first = consider(exFirst, first, (a, b) => a < b);
    last = consider(exLast, last, (a, b) => a > b);
  });
  return { first, last };
}

// Cap an edit-derived exercise duration at MAX_EXERCISE_DURATION_MS (MAX only,
// no MIN floor). Used by the foreign-match window in resolveForeignMatches_ to
// keep its upper bound consistent with editDerivedDurationMs_, which the
// recorded 'edit' interval uses, and which applies the same MAX cap PLUS a MIN
// floor / start-only default that this helper deliberately omits.
function capExerciseDurationToMax_(rawDurationMs) {
  return Math.min(rawDurationMs, MAX_EXERCISE_DURATION_MS);
}

// True when the row's exercise edit timestamps form a trustworthy same-day
// window: both first/last edit are set AND the first edit's civil date matches
// the row's Date. This gates the foreign-match window: an off-date edit
// (correcting an old row today) must not anchor timing to today. The window
// itself is clamped to MAX_EXERCISE_DURATION_MS from the first edit, so a late
// last-edit can't distort it.
function exerciseEditIsOnRowDate_(row) {
  return Boolean(
    row.exerciseFirstEditedAt &&
    row.exercisesLastEditedAt &&
    ymd(row.exerciseFirstEditedAt) === ymd(row.date),
  );
}

// True when those timestamps describe the workout itself rather than a later
// correction, which is the stricter thing the 'edit' timing path needs: the
// last edit counts as evidence of when the session ended only if it is close
// enough to the first to plausibly be part of it.
//
// Without the span test, ANY later edit rebuilds the interval as
// firstEdit .. firstEdit + MAX_EXERCISE_DURATION_MS, because editDerivedDurationMs_
// clamps the (huge) raw span to the cap. Fixing a typo at 7pm on the workout's
// own day, or touching the row months later, would both turn a recorded 30
// minute session into a fabricated 2 hour one. Beyond the cap the last edit
// tells us nothing about the workout, so resolveRowTiming_ falls through to
// 'prior' and reuses the recorded interval verbatim.
//
// Only meaningful when there is a recorded interval to protect; see
// exerciseEditIsUsable_, which is what resolveRowTiming_ actually gates on.
//
// Consequence worth knowing: a workout logged sparsely (first set typed at
// 9:00, the rest filled in at 11:30) freezes at whatever the 9:00 sync
// recorded, which is the 10 minute start-only default. Timestamps cannot
// distinguish "still logging this workout" from "correcting it later", and
// MAX_EXERCISE_DURATION_MS is the stated belief about how long a workout can
// run, so an edit further out than that is treated as a correction. Logging
// sets as you go keeps the interval accurate.
function exerciseEditSpansWorkout_(row) {
  if (!exerciseEditIsOnRowDate_(row)) {
    return false;
  }
  const spanMs =
    row.exercisesLastEditedAt.getTime() - row.exerciseFirstEditedAt.getTime();
  return spanMs <= MAX_EXERCISE_DURATION_MS;
}

// Whether resolveRowTiming_ should build the interval from the edit timestamps.
// The span test guards an interval we already recorded, so it only applies when
// there is one: with no prior datapoint there is nothing to protect, and the
// row's observed on-date start is far better evidence than synthetic noon. In
// that case the timestamps are used with editDerivedDurationMs_'s clamp, which
// is the only path where its MAX cap still does work.
function exerciseEditIsUsable_(row, priorExercise) {
  if (!exerciseEditIsOnRowDate_(row)) {
    return false;
  }
  return exerciseEditSpansWorkout_(row) || !priorExercise;
}

// Map a raw edit-derived duration (last edit - first edit) to the recorded
// exercise duration by clamping to [MIN_EXERCISE_DURATION_MS,
// MAX_EXERCISE_DURATION_MS]. A single-edit row has raw <= 0 (start == last
// edit, no observed end), which the MIN floor turns into the start-only
// default, so the start-only case needs no special handling.
//
// The MAX cap binds only on the no-prior path: exerciseEditIsUsable_ otherwise
// refuses the 'edit' path once the raw span passes the cap, so a row with a
// recorded interval keeps it instead of being stretched to the cap.
function editDerivedDurationMs_(rawDurationMs) {
  return Math.min(
    Math.max(rawDurationMs, MIN_EXERCISE_DURATION_MS),
    MAX_EXERCISE_DURATION_MS,
  );
}

// Resolve the exercise interval and weight sample time independently, per
// phase, with these rules:
//
//   Exercise (first matching rule wins):
//     - 'foreign'   if a foreignInterval is provided (an overlapping foreign
//                   session whose manual start/stop is more accurate than our
//                   edit-derived window). Its interval is used verbatim.
//     - 'edit'      if exerciseEditIsUsable_(row, priorExercise): the first
//                   edit is on row.date, and either the last edit is within
//                   MAX_EXERCISE_DURATION_MS of it or there is no prior
//                   datapoint to protect. This lets endTime advance during a
//                   live workout as more sets are typed in, without letting a
//                   later correction rewrite an interval already recorded.
//     - 'prior'     if a previous datapoint is provided. Its interval is
//                   reused verbatim, so neither an off-date edit (correcting
//                   an old row today) nor a late same-day one (fixing a typo
//                   in the evening) shifts the recorded times.
//     - 'synthetic' otherwise: noon on row.date.
//
//   Weight (only consumed on the POST path; PATCH preserves sampleTime
//   server-side by echoing back the prior GET, so the weight resolver
//   isn't called with a prior):
//     - 'edit'      if weightEditedAt's civil date == row.date.
//     - 'synthetic' otherwise: noon on row.date.
//
// priorExercise is the GET response for the row's existing exercise
// datapoint (or null if first-sync, or null if the GET failed, in which
// case we fall through to the edit/synthetic path rather than erroring).
// foreignInterval is an overlapping foreign session (from
// resolveForeignMatches_) whose start/end should be borrowed, or null.
function resolveRowTiming_(row, priorExercise, foreignInterval) {
  const tz = getTz_();
  const rowDateKey = ymd(row.date);

  let exercise = null;
  let exerciseSource = null;
  const exerciseEditIsUsable = exerciseEditIsUsable_(row, priorExercise);
  if (foreignInterval) {
    exercise = {
      endOffsetSeconds: foreignInterval.endUtcOffsetSeconds,
      endUtcMs: foreignInterval.endUtcMs,
      startOffsetSeconds: foreignInterval.startUtcOffsetSeconds,
      startUtcMs: foreignInterval.startUtcMs,
    };
    exerciseSource = "foreign";
  } else if (exerciseEditIsUsable) {
    const startMs = row.exerciseFirstEditedAt.getTime();
    const rawDuration = row.exercisesLastEditedAt.getTime() - startMs;
    // A single edit (start == last, rawDuration <= 0) has no observed end;
    // editDerivedDurationMs_'s MIN floor gives it the start-only default. A
    // second edit produces a real span (clamped to [MIN, MAX]).
    const endMs = startMs + editDerivedDurationMs_(rawDuration);
    exercise = {
      endOffsetSeconds: getTzOffsetSeconds_(tz, new Date(endMs)),
      endUtcMs: endMs,
      startOffsetSeconds: getTzOffsetSeconds_(tz, row.exerciseFirstEditedAt),
      startUtcMs: startMs,
    };
    exerciseSource = "edit";
  } else if (priorExercise) {
    const i = priorExercise.exercise && priorExercise.exercise.interval;
    if (i && i.startTime && i.endTime) {
      exercise = {
        endOffsetSeconds: parseOffsetSeconds_(i.endUtcOffset),
        endUtcMs: new Date(i.endTime).getTime(),
        startOffsetSeconds: parseOffsetSeconds_(i.startUtcOffset),
        startUtcMs: new Date(i.startTime).getTime(),
      };
      exerciseSource = "prior";
    }
  }
  if (!exercise) {
    exercise = syntheticExerciseInterval_(row.date);
    exerciseSource = "synthetic";
  }
  // Whole seconds, matching what createExerciseAt serializes (its RFC3339
  // format drops milliseconds). Edit timestamps carry milliseconds, and a
  // foreign interval can, so without the floor a re-sync target never equals
  // the GET-back of the datapoint the previous pass created from the same
  // inputs, exerciseUnchanged_ never matches, and every backstop re-review
  // delete+recreates an identical datapoint: resource-name churn, and the
  // rapid same-interval recreate pattern the Health card layer punishes.
  exercise.startUtcMs = Math.floor(exercise.startUtcMs / 1000) * 1000;
  exercise.endUtcMs = Math.floor(exercise.endUtcMs / 1000) * 1000;

  let weight;
  let weightSource;
  if (row.weightEditedAt && ymd(row.weightEditedAt) === rowDateKey) {
    weight = {
      offsetSeconds: getTzOffsetSeconds_(tz, row.weightEditedAt),
      utcMs: row.weightEditedAt.getTime(),
    };
    weightSource = "edit";
  } else {
    weight = syntheticWeightSample_(row.date);
    weightSource = "synthetic";
  }
  return {
    exercise,
    exerciseSource,
    weight,
    weightSource,
  };
}

// True when an existing exercise datapoint (the GET response) already carries
// the target interval and notes, so a re-sync can skip the delete+recreate and
// keep its resource name (no churn). Compares interval start/end in absolute ms
// (the offsets are derived from the same instants, so ms equality suffices) and
// the notes string exactly. A missing interval/endpoint counts as changed so
// the row recreates. Pure (no API/sheet access).
function exerciseUnchanged_(
  prior,
  targetStartUtcMs,
  targetEndUtcMs,
  targetNotes,
) {
  const ex = prior && prior.exercise;
  const i = ex && ex.interval;
  if (!i || !i.startTime || !i.endTime) {
    return false;
  }
  if (new Date(i.startTime).getTime() !== targetStartUtcMs) {
    return false;
  }
  if (new Date(i.endTime).getTime() !== targetEndUtcMs) {
    return false;
  }
  return (ex.notes || "") === (targetNotes || "");
}

// Delete a row's prior datapoints of one type, ONE NAME PER CALL, and return
// the names that must stay in Created Health IDs (their delete genuinely
// failed, so the next sync retries them). A 404 counts as already deleted
// (the datapoint was removed in the Health app) and its name is dropped;
// retrying a delete that can never succeed would hold the row dirty and
// re-issue the same call every poll.
//
// Per-name rather than one batched deleteDataPointsByName call because
// :batchDelete fails as a unit: one missing name would 404 the whole batch, and
// treating that as "all deleted" would drop still-live siblings from the sheet,
// leaving them untracked and reclaimable only by orphan reconciliation (never,
// for a row older than ORPHAN_RECONCILE_LOOKBACK_DAYS). Rows normally carry a
// single id per type, so this is the same one call in the common case.
function deletePriorDataPoints_(tag, label, names) {
  const remaining = [];
  names.forEach((name) => {
    try {
      deleteDataPointsByName([name]);
    } catch (err) {
      if (isNotFoundError_(err)) {
        console.warn(
          `${tag}: previous ${label} datapoint not found (404); treating as deleted: ${name}`,
        );
        return;
      }
      console.error(
        `${tag}: delete previous ${label} datapoint failed (${name}): ${err}`,
      );
      remaining.push(name);
    }
  });
  return remaining;
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
function syncOneRow_(
  row,
  foreignMatches,
  weightReady,
  exerciseReady,
  cols,
  doneIdx,
  total,
) {
  const dateKey = ymd(row.date);
  const tag = `[${doneIdx}/${total}] ${dateKey} row ${row.rowNum}`;
  const phases = [];
  if (weightReady) {
    phases.push("weight");
  }
  if (exerciseReady) {
    phases.push("exercise");
  }
  console.info(
    `${tag}: starting phases=[${phases.join(",")}] (exercises=${
      row.exercises.length
    }, bodyweight=${
      row.bodyweight === null ? "none" : row.bodyweight
    }, oldIds=${row.healthIds.length}${
      foreignMatches.length > 0 ? `, align=${foreignMatches.length}` : ""
    })`,
  );

  // Row numbers are positional: a row inserted or deleted above this one after
  // the pass's snapshot shifts every row below it, and a write to the captured
  // rowNum would land on a neighbor, recording this row's datapoint ids and
  // stamps against another row's content. Structural changes fire onChange,
  // not onEdit, so neither the generation marker nor the phase guards can see
  // them; re-reading the Date cell at the row's turn catches the shift
  // instead (dates are strictly increasing, so a shifted row never shows the
  // snapshot's date). On a mismatch the row is skipped before any API call or
  // write; it stays dirty and the next pass snapshots fresh row numbers. A
  // shift landing within this one row's processing remains unguarded, the
  // same best-effort window the stamp-time edit guards accept.
  if (cols.dateCol) {
    const liveDate = toDate_(
      getSheet_().getRange(row.rowNum, cols.dateCol).getValue(),
    );
    if (!liveDate || liveDate.getTime() !== row.date.getTime()) {
      console.warn(
        `${tag}: Date cell changed since the snapshot (row inserted/deleted ` +
          `above?); deferring this row to the next pass.`,
      );
      return false;
    }
  }

  const split = splitHealthIdsByType_(row.healthIds);

  // Phases that will actually issue a create (and thus need a resolved
  // interval/sampleTime). Delete-only phases don't need timing. The
  // timing log line shows only the phases listed here, so weight-only rows
  // don't surface a misleading "edit/synthetic" label for an interval that
  // will never be sent.
  const exerciseWillCreate =
    exerciseReady && hasSendableExercises_(row.exercises);

  // Fetch prior datapoints. Exercise: whenever the row has a prior exercise id
  // and the phase will create: the GET serves two purposes now. (1) The 'prior'
  // timing source reuses its interval verbatim when neither foreign-match nor
  // same-date editing applies (the resolver ignores it for foreign/edit, which
  // win). (2) The idempotency check compares the prior interval + notes to the
  // freshly-computed ones to skip an unchanged recreate. Weight: when we'll
  // PATCH (i.e. prior weight ID exists AND bodyweight is set): the PATCH body
  // requires sampleTime, read from this GET. Exercise GET failure is non-fatal
  // (timing falls through to edit/synthetic and the recreate proceeds); weight
  // GET failure forces the PATCH to fail and the row to retry next pass.
  // Every prior, not just the first: a row can now hold one datapoint per
  // app-recorded workout, and the idempotency check below matches each target
  // against the priors so an unchanged workout keeps its resource name even
  // when a sibling on the same row changed.
  const priorExercises = [];
  let priorWeight = null;
  let priorWeightFetchFailed = false;
  if (exerciseWillCreate) {
    split.exercise.forEach((name) => {
      try {
        priorExercises.push({ dp: getDataPoint(name), name });
      } catch (err) {
        console.warn(
          `${tag}: GET prior exercise ${name} failed; will recreate: ${err}`,
        );
      }
    });
  }
  const priorExercise = priorExercises.length > 0 ? priorExercises[0].dp : null;
  const weightWillPatch =
    weightReady && row.bodyweight !== null && split.weight.length > 0;
  if (weightWillPatch) {
    try {
      priorWeight = getDataPoint(split.weight[0]);
    } catch (err) {
      if (isNotFoundError_(err)) {
        // The prior weight datapoint is gone server-side (e.g. deleted in the
        // Health app). Drop the stale ID so the dispatch below falls through to
        // POST a fresh one instead of PATCHing/GETting a name that 404s forever.
        console.warn(
          `${tag}: prior weight datapoint not found (404); dropping stale ID and recreating.`,
        );
        split.weight = [];
      } else {
        console.warn(
          `${tag}: GET prior weight failed; PATCH will fail and the row will retry: ${err}`,
        );
        priorWeightFetchFailed = true;
      }
    }
  }
  // Only POST creates need timing resolution. The PATCH path (prior weight ID
  // present + bodyweight set) preserves sampleTime server-side, so no prior
  // GET and no timing label. Computed after the GET block so a 404-dropped
  // stale ID is reflected here (the row now POSTs rather than PATCHes).
  const weightWillCreate =
    weightReady && row.bodyweight !== null && split.weight.length === 0;

  // Weight timing, and the exercise timing for any group with no app session to
  // borrow from (edit/prior/synthetic). A group that DID claim a session
  // resolves separately below, passing that session so 'foreign' wins.
  let timing;
  try {
    timing = resolveRowTiming_(row, priorExercise, null);
  } catch (err) {
    console.error(`${tag}: resolveRowTiming_ failed: ${err}`);
    return false;
  }
  if (weightWillCreate) {
    console.info(`${tag}: timing weight=${timing.weightSource}`);
  }
  let newWeightIds = split.weight;
  let newExerciseIds = split.exercise;
  let weightFailed = false;
  let exerciseFailed = false;

  if (weightReady) {
    const hasBodyweight = row.bodyweight !== null;
    if (split.weight.length > 0 && hasBodyweight) {
      // PATCH in place. Preserves sampleTime (echoed back from the prior
      // GET: the API rejects PATCH bodies without sampleTime), createTime,
      // dataSource. Resource name stays the same so Created Health IDs
      // doesn't churn.
      const sampleTime =
        priorWeight && priorWeight.weight && priorWeight.weight.sampleTime;
      if (!sampleTime) {
        const reason = priorWeightFetchFailed
          ? "prior weight GET failed"
          : "prior datapoint missing sampleTime";
        console.error(
          `${tag}: patchWeight skipped (${reason}); will retry next sync.`,
        );
        weightFailed = true;
      } else {
        try {
          patchWeight(split.weight[0], sampleTime, row.bodyweight);
          console.info(
            `${tag}: patchWeight(${row.bodyweight} lb) -> ${split.weight[0]}`,
          );
        } catch (err) {
          console.error(`${tag}: patchWeight failed: ${err}`);
          weightFailed = true;
        }
      }
    } else if (split.weight.length > 0 && !hasBodyweight) {
      // Bodyweight cleared on a row that previously had one: delete.
      console.info(
        `${tag}: deleting ${split.weight.length} previous weight datapoint(s)`,
      );
      // Names left in newWeightIds are the ones whose delete failed; they stay
      // in Created Health IDs so the next sync retries them.
      newWeightIds = deletePriorDataPoints_(tag, "weight", split.weight);
      if (newWeightIds.length > 0) {
        weightFailed = true;
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
        writeHealthIds(
          row.rowNum,
          cols.healthIdsCol,
          newWeightIds.concat(newExerciseIds).concat(split.other),
        );
        console.info(`${tag}: createWeightAt(${row.bodyweight} lb) -> ${name}`);
      } catch (err) {
        console.error(`${tag}: createWeightAt failed: ${err}`);
        weightFailed = true;
      }
    } else {
      // No prior, no current. Nothing to do.
      newWeightIds = [];
    }
  }

  const usedSessionNames = [];
  if (exerciseReady) {
    // One datapoint per app-recorded workout the row's exercises were logged
    // during, plus one for whatever was not attributable to any of them. With
    // no sessions (or one that caught everything) this is a single group and
    // behaves exactly as the pre-split sync did.
    const groups = exerciseWillCreate
      ? partitionExercisesBySession_(
          foreignMatches,
          row.exerciseEditTimes,
          row.exerciseFirstEditedAt,
          row.exercises,
        )
      : [];

    // Resolve each group's interval and notes. Two groups landing on the SAME
    // interval are collapsed to the first: POSTing a second datapoint with an
    // identical (client, recordingMethod, exerciseType, interval) returns the
    // existing one and silently discards the new body, so the second group's
    // notes would be lost rather than written. See the write-time upsert note
    // in the README.
    const targets = [];
    const seenIntervals = {};
    groups.forEach((g) => {
      // A session group borrows its session's interval. The null-session group
      // is timed from ITS OWN exercises' edit window (groupEditWindow_): after
      // the attributed exercises are split away, the row-level first..last span
      // covers the app workouts too, and would stretch this group's datapoint
      // across them. The same resolver rules (on-date gate, span-vs-MAX,
      // prior fallback, clamps) apply to the narrowed window.
      let groupTiming;
      if (g.session) {
        groupTiming = resolveRowTiming_(row, priorExercise, g.session);
      } else {
        const w = groupEditWindow_(row, row.exerciseEditTimes, g.exercises);
        groupTiming = resolveRowTiming_(
          Object.assign({}, row, {
            exerciseFirstEditedAt: w.first,
            exercisesLastEditedAt: w.last,
          }),
          priorExercise,
          null,
        );
      }
      const gex = groupTiming.exercise;
      const key = `${gex.startUtcMs}-${gex.endUtcMs}`;
      if (seenIntervals[key]) {
        console.warn(
          `${tag}: two exercise groups resolved to the same interval; ` +
            `merging is not possible, dropping the later one.`,
        );
        return;
      }
      seenIntervals[key] = true;
      targets.push({
        endOffsetSeconds: gex.endOffsetSeconds,
        endUtcMs: gex.endUtcMs,
        notes: buildNotes(gex.endUtcMs - gex.startUtcMs, g.exercises),
        sessionName: g.session ? g.session.name : null,
        source: groupTiming.exerciseSource,
        startOffsetSeconds: gex.startOffsetSeconds,
        startUtcMs: gex.startUtcMs,
      });
    });
    if (targets.length > 0) {
      console.info(
        `${tag}: timing exercise=[${targets.map((t) => t.source).join(",")}]`,
      );
    }
    targets.forEach((t) => {
      if (t.sessionName) {
        usedSessionNames.push(t.sessionName);
      }
    });

    // Why delete+recreate rather than PATCH: the Health API's exercise PATCH
    // does not merge `notes`, which is the field a row edit changes. Measured
    // against the live API on 2026-08-19: a full body returns 200 done:true
    // and leaves the notes as they were (only `interval` merges), so a PATCH
    // here would stamp the row synced while Health kept the old text. See the
    // Google Cloud caveats in the README.
    //
    // Idempotency: a target whose interval + notes already match one of the
    // row's existing datapoints keeps that datapoint untouched, so an unchanged
    // re-sync costs only the prior GETs, with no write and no resource-name
    // churn. Priors are matched by CONTENT, not by position: a row can hold
    // several, and a recreate gives a new resource name, so there is no stable
    // ordering to match on. Each prior is claimed at most once.
    const claimedPriors = {};
    const keptNames = [];
    const toCreate = [];
    targets.forEach((t) => {
      const match = priorExercises.filter(
        (p) =>
          !claimedPriors[p.name] &&
          exerciseUnchanged_(p.dp, t.startUtcMs, t.endUtcMs, t.notes),
      )[0];
      if (match) {
        claimedPriors[match.name] = true;
        keptNames.push(match.name);
      } else {
        toCreate.push(t);
      }
    });
    const staleNames = split.exercise.filter((n) => !claimedPriors[n]);
    newExerciseIds = keptNames.slice();

    if (keptNames.length > 0) {
      console.info(
        `${tag}: ${keptNames.length} exercise datapoint(s) unchanged; skip recreate`,
      );
    }
    if (staleNames.length > 0) {
      console.info(
        `${tag}: deleting ${staleNames.length} previous exercise datapoint(s)`,
      );
      // Survivors are the failures: they stay in Created Health IDs so the next
      // sync retries them, and any survivor blocks the creates below so the row
      // does not end up holding both the stale datapoint and its replacement.
      const failedDeletes = deletePriorDataPoints_(tag, "exercise", staleNames);
      if (failedDeletes.length > 0) {
        exerciseFailed = true;
        newExerciseIds = newExerciseIds.concat(failedDeletes);
      }
    }
    if (!exerciseFailed) {
      for (let i = 0; i < toCreate.length; i++) {
        const t = toCreate[i];
        try {
          // createExerciseAt throws if the create returns no resource name, so a
          // returned name is always usable here.
          const name = createExerciseAt(
            t.startUtcMs,
            t.startOffsetSeconds,
            t.endUtcMs,
            t.endOffsetSeconds,
            t.notes,
          );
          newExerciseIds.push(name);
          // Same rationale as the weight write above: persist before any
          // later step can fail and leave the datapoint untracked.
          writeHealthIds(
            row.rowNum,
            cols.healthIdsCol,
            newWeightIds.concat(newExerciseIds).concat(split.other),
          );
          console.info(
            `${tag}: createExerciseAt${
              t.sessionName ? " (foreign-aligned)" : ""
            } -> ${name}`,
          );
        } catch (err) {
          console.error(`${tag}: createExerciseAt failed: ${err}`);
          exerciseFailed = true;
          break;
        }
      }
    }
  }

  writeHealthIds(
    row.rowNum,
    cols.healthIdsCol,
    newWeightIds.concat(newExerciseIds).concat(split.other),
  );
  if (exerciseReady) {
    writeMatchedHealthSessions(
      row.rowNum,
      cols.matchedHealthSessionCol,
      usedSessionNames,
    );
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
    const currentEdit = toDate_(
      getSheet_()
        .getRange(row.rowNum, cols.exercisesLastEditedAtCol)
        .getValue(),
    );
    const previousMs = row.exercisesLastEditedAt
      ? row.exercisesLastEditedAt.getTime()
      : null;
    const currentMs = currentEdit ? currentEdit.getTime() : null;
    if (currentMs !== previousMs) {
      console.info(
        `${tag}: concurrent exercise edit detected (Exercises Last Edited At ${humanizeDate_(
          row.exercisesLastEditedAt,
        )} -> ${
          currentEdit ? humanizeDate_(currentEdit) : "<cleared>"
        }); deferring Exercise Synced At stamp.`,
      );
      exerciseConcurrentEdit = true;
    }
  }
  let weightConcurrentEdit = false;
  if (weightReady && cols.weightEditedAtCol) {
    const currentEdit = toDate_(
      getSheet_().getRange(row.rowNum, cols.weightEditedAtCol).getValue(),
    );
    const previousMs = row.weightEditedAt ? row.weightEditedAt.getTime() : null;
    const currentMs = currentEdit ? currentEdit.getTime() : null;
    if (currentMs !== previousMs) {
      console.info(
        `${tag}: concurrent weight edit detected (Weight Edited At ${humanizeDate_(
          row.weightEditedAt,
        )} -> ${
          currentEdit ? humanizeDate_(currentEdit) : "<cleared>"
        }); deferring Weight Synced At stamp.`,
      );
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
    console.warn(`${tag}: FAILED (partial); will retry on next sync.`);
    return false;
  }

  // If the row still has unstamped phases (either because we skipped a phase
  // this pass or a concurrent edit blocked the stamp), advance the dirty
  // generation so syncDirtyRows' end-of-pass check leaves the flag set
  // (and a future poll picks the row up).
  const weightStampMissing =
    !row.weightSyncedAt && !(weightReady && !weightConcurrentEdit);
  const exerciseStampMissing =
    !row.exerciseSyncedAt && !(exerciseReady && !exerciseConcurrentEdit);
  if (weightStampMissing || exerciseStampMissing) {
    markPendingDirty_();
    console.info(
      `${tag}: partial progress; row stays dirty ` +
        `(weightStamped=${!weightStampMissing}, exerciseStamped=${!exerciseStampMissing})`,
    );
  } else {
    console.info(`${tag}: done`);
  }
  return true;
}
