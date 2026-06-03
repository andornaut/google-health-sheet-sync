function setup() {
  ensureManagedColumns();
  installTriggers();
}

function onOpen() {
  console.info('onOpen: installing Sync menu');
  SpreadsheetApp.getUi()
    .createMenu('Sync')
    .addItem('Run now', 'runSyncNow')
    .addItem('Force resync current row', 'forceResyncCurrentRow')
    .addItem('Force resync all rows', 'forceResyncAllRows')
    .addSeparator()
    .addItem('Run setup', 'setup')
    .addItem('Authorize Health API', 'authorizeHealthApi')
    .addItem('Revoke Health API', 'revokeHealthApi')
    .addSeparator()
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
  const handlers = new Set(['onEditTrigger', 'flushIfPending']);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('flushIfPending').timeBased().everyMinutes(POLL_INTERVAL_MIN).create();
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
  if (!props.getProperty(PENDING_DIRTY_KEY)) {
    console.info('flushIfPending: no pending edits, skipping');
    return;
  }
  console.info('flushIfPending: pending edits detected, syncing');
  syncDirtyRows(false, 0);
}

// Write a fresh generation marker into PENDING_DIRTY_KEY. The value matters
// (syncDirtyRows compares start vs end to detect concurrent edits), so always
// advance it — never just re-write the same string.
function markPendingDirty_() {
  PropertiesService.getScriptProperties()
    .setProperty(PENDING_DIRTY_KEY, String(Date.now()));
}

// "Run now" is an explicit manual action: bypasses the exercise edit-burst
// debounce so the user sees results immediately. Weight phase has no
// debounce, so this only changes behavior for exercise content. If they
// keep editing afterward, the row goes dirty again and the next sync
// replaces the Health datapoint(s).
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
    result = syncDirtyRows(true, LOCK_WAIT_MS);
  } catch (err) {
    toast_('Sync failed: ' + String(err.message || err), 30);
    throw err;
  }
  toastSyncResult_(result, verb);
}

function forceResyncCurrentRow() {
  const sheet = getSheet_();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return;
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
  clearRowExerciseSynced(row, exerciseCol);
  clearRowWeightSynced(row, weightSyncedAtCol);
  SpreadsheetApp.flush();
  markPendingDirty_();
  runSyncAndToast_('Resynced');
}

function forceResyncAllRows() {
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
  // Capture the dirty-marker generation at start. onEditMarkDirty advances it
  // on every edit (Date.now() string), so a concurrent edit during the pass
  // shows up as a mismatch at end-of-pass. The flag is NOT cleared here:
  // a hard kill (6-min Apps Script timeout, uncaught throw) before the
  // finally block would otherwise drop the signal and orphan the dirty rows
  // until the next manual sync or new edit.
  const genAtStart = props.getProperty(PENDING_DIRTY_KEY);
  let ok = 0;
  let errors = 0;
  let waitingCount = 0;
  let deferredCount = 0;
  try {
    const { rows, exerciseSyncedAtCol, weightSyncedAtCol, weightCol, healthIdsCol, exercisesLastEditedAtCol, weightEditedAtCol, matchedHealthSessionCol } = readRows();
    if (!exerciseSyncedAtCol || !weightSyncedAtCol || !healthIdsCol) {
      console.error('syncDirtyRows: managed columns missing; run setup().');
      errors = 1;
      return { ok: 0, errors: errors };
    }
    const dirty = rows.filter(r => !r.exerciseSyncedAt || !r.weightSyncedAt);
    if (dirty.length === 0) {
      return { ok: 0, errors: 0 };
    }

    const ordinalByRowNum = buildOrdinalMap_(rows);

    // Per-row phase readiness:
    //   - Weight phase: always ready when the row is weight-dirty (no debounce).
    //   - Exercise phase: ready iff bypassed, or the row has no Exercises
    //     Edited At timestamp, or the debounce window has elapsed since the
    //     last exercise edit. Rows with no exercise content also pass
    //     instantly since there's nothing to time.
    const now = Date.now();
    const ready = [];
    let maxRemainingMs = 0;
    dirty.forEach(r => {
      const weightReady = !r.weightSyncedAt;
      let exerciseReady = false;
      let remainingMs = 0;
      if (!r.exerciseSyncedAt) {
        if (bypassQuiesce || !r.exercisesLastEditedAt || r.exercises.length === 0) {
          exerciseReady = true;
        } else {
          const sinceMs = now - r.exercisesLastEditedAt.getTime();
          exerciseReady = sinceMs >= LAST_EDIT_QUIESCE_MS;
          remainingMs = LAST_EDIT_QUIESCE_MS - sinceMs;
        }
      }
      if (weightReady || exerciseReady) {
        ready.push({ row: r, weightReady: weightReady, exerciseReady: exerciseReady });
      } else {
        waitingCount++;
        if (remainingMs > maxRemainingMs) maxRemainingMs = remainingMs;
      }
    });

    if (waitingCount > 0) {
      console.info('syncDirtyRows: ' + waitingCount + ' row(s) still in edit-debounce window ('
        + humanizeMs_(maxRemainingMs) + ' remaining of ' + humanizeMs_(LAST_EDIT_QUIESCE_MS)
        + '); will retry next pass.');
    }
    if (ready.length === 0) {
      console.info('syncDirtyRows: no rows ready to sync.');
      return { ok: 0, errors: 0 };
    }
    if (bypassQuiesce) console.info('syncDirtyRows: bypassQuiesce=true');

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
    const matchPlan = resolveForeignMatches_(rows, exerciseReadyRows);
    const cols = {
      exerciseSyncedAtCol: exerciseSyncedAtCol,
      weightSyncedAtCol: weightSyncedAtCol,
      weightCol: weightCol,
      healthIdsCol: healthIdsCol,
      exercisesLastEditedAtCol: exercisesLastEditedAtCol,
      weightEditedAtCol: weightEditedAtCol,
      matchedHealthSessionCol: matchedHealthSessionCol
    };
    for (let i = 0; i < ready.length; i++) {
      const entry = ready[i];
      const ordinal = ordinalByRowNum[entry.row.rowNum];
      const match = matchPlan[entry.row.rowNum] || null;
      if (syncOneRow_(entry.row, ordinal, match, entry.weightReady, entry.exerciseReady, cols, i + 1, ready.length)) ok++;
      else errors++;
    }
  } finally {
    // End-of-pass flag resolution:
    //   - If work remains (quiescing, errors, deferred): ensure the flag is
    //     set so a future poll picks it up. If a concurrent edit advanced
    //     the generation already, its value is fine; otherwise write a fresh
    //     one. (syncOneRow_ also calls markPendingDirty_ for partial-progress
    //     rows, so ok-counted rows can leave the flag set too.)
    //   - If no work remains AND the generation hasn't moved: the pass fully
    //     drained the queue, safe to clear.
    //   - If no work remains BUT the generation moved: an edit landed during
    //     the pass that this pass didn't see (readRows snapshotted before it).
    //     Leave the new generation in place so the next poll runs.
    const genAtEnd = props.getProperty(PENDING_DIRTY_KEY);
    const concurrentEdit = genAtEnd !== genAtStart;
    const workRemaining = waitingCount > 0 || errors > 0 || deferredCount > 0;
    if (workRemaining) {
      if (!concurrentEdit) markPendingDirty_();
    } else if (!concurrentEdit) {
      props.deleteProperty(PENDING_DIRTY_KEY);
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
  // Clamp to MAX_EXERCISE_DURATION_MS the same way resolveRowTiming_ does so a
  // row whose exercisesLastEditedAt drifted far past exerciseFirstEditedAt
  // (late corrections that keep sticky first-edit + advance last-edit) doesn't
  // produce a multi-day window biased toward the longest unrelated candidate.
  const windows = readyRows
    .filter(r => r.exercises.length > 0
      && r.exerciseFirstEditedAt && r.exercisesLastEditedAt
      && ymd(r.exerciseFirstEditedAt) === ymd(r.date))
    .map(r => {
      const startMs = r.exerciseFirstEditedAt.getTime();
      const rawDuration = r.exercisesLastEditedAt.getTime() - startMs;
      const clampedEndMs = startMs + Math.min(rawDuration, MAX_EXERCISE_DURATION_MS);
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
  const exerciseEditOnRowDate = row.exerciseFirstEditedAt && row.exercisesLastEditedAt
    && ymd(row.exerciseFirstEditedAt) === rowDateKey;
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
    const clampedDuration = Math.min(
      Math.max(rawDuration, MIN_EXERCISE_DURATION_MS),
      MAX_EXERCISE_DURATION_MS
    );
    const endMs = startMs + clampedDuration;
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

// Sync a single row in two independent phases (weight, exercise). Either or
// both phases may run on a given pass:
//   - Weight phase runs whenever the row's Weight Synced At is cleared. It
//     reconciles weight IDs with the sheet's bodyweight (write/delete) and
//     stamps Weight Synced At on success.
//   - Exercise phase runs only when the caller passed exerciseReady=true
//     (i.e. quiesce passed or no exercise content). It reconciles exercise
//     IDs with the sheet's exercises (delete + recreate, optionally aligning
//     the interval to an overlapping foreign session) and stamps Exercise
//     Synced At on success.
// Returns true if the pass made forward progress on the row without errors
// (including the case where the row stays dirty because the other phase is
// still pending). Returns false if any attempted phase failed.
function syncOneRow_(row, ordinal, match, weightReady, exerciseReady, cols, doneIdx, total) {
  const dateKey = ymd(row.date);
  const tag = '[' + doneIdx + '/' + total + '] ' + dateKey + ' row ' + row.rowNum;
  const phases = [];
  if (weightReady) phases.push('weight');
  if (exerciseReady) phases.push('exercise');
  console.info(tag + ': starting phases=[' + phases.join(',') + '] (exercises=' + row.exercises.length
    + ', bodyweight=' + (row.bodyweight === null ? 'none' : row.bodyweight)
    + ', oldIds=' + row.healthIds.length
    + (match ? ', align=' + match.name : '') + ')');

  const split = splitHealthIdsByType_(row.healthIds);

  // Phases that will actually issue a create (and thus need a resolved
  // interval/sampleTime). Delete-only phases don't need timing. The
  // timing log line shows only the phases listed here, so weight-only rows
  // don't surface a misleading "edit/synthetic" label for an interval that
  // will never be sent.
  const exerciseWillCreate = exerciseReady && SYNC_EXERCISES && row.exercises.length > 0;
  // Only POST creates need timing resolution. The PATCH path (prior weight ID
  // present + bodyweight set) preserves sampleTime server-side, so no prior
  // GET and no timing label.
  const weightWillCreate = weightReady && SYNC_WEIGHT && row.bodyweight !== null && split.weight.length === 0;

  // Fetch prior datapoints. Exercise: only when the edit isn't on row.date
  // (otherwise the live-workout endTime-advancement path takes over) — the
  // timing resolver reuses the prior interval verbatim. Weight: when we'll
  // PATCH (i.e. prior weight ID exists AND bodyweight is set) — the PATCH
  // body requires sampleTime, which is read from this GET. Exercise GET
  // failure is non-fatal (timing falls through to edit/synthetic); weight
  // GET failure forces the PATCH to fail and the row to retry next pass.
  let priorExercise = null;
  let priorWeight = null;
  let priorWeightFetchFailed = false;
  const exerciseEditOnRowDate = row.exerciseFirstEditedAt && row.exercisesLastEditedAt
    && ymd(row.exerciseFirstEditedAt) === ymd(row.date);
  if (exerciseWillCreate && !match && !exerciseEditOnRowDate && split.exercise.length > 0) {
    try {
      priorExercise = getDataPoint(split.exercise[0]);
    } catch (err) {
      console.warn(tag + ': GET prior exercise failed; will recompute timing: ' + err);
    }
  }
  const weightWillPatch = weightReady && SYNC_WEIGHT && row.bodyweight !== null && split.weight.length > 0;
  if (weightWillPatch) {
    try {
      priorWeight = getDataPoint(split.weight[0]);
    } catch (err) {
      console.warn(tag + ': GET prior weight failed; PATCH will fail and the row will retry: ' + err);
      priorWeightFetchFailed = true;
    }
  }

  let timing;
  try {
    // `match`, when set, is an overlapping foreign session whose interval the
    // resolver borrows verbatim ('foreign' wins over edit/prior/synthetic).
    timing = resolveRowTiming_(row, ordinal, priorExercise, match);
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
    const hasBodyweight = SYNC_WEIGHT && row.bodyweight !== null;
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
        const name = createWeightAt(wt.utcMs, wt.offsetSeconds, row.bodyweight);
        if (name) {
          newWeightIds.push(name);
          // Persist immediately so a 6-minute kill before the end-of-row
          // write can't orphan a freshly-created datapoint we no longer
          // have a sheet reference for.
          writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
        }
        console.info(tag + ': createWeightAt(' + row.bodyweight + ' lb) -> ' + (name || '<no name>'));
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
    if (split.exercise.length > 0) {
      console.info(tag + ': deleting ' + split.exercise.length + ' previous exercise datapoint(s)');
      try {
        deleteDataPointsByName(split.exercise);
        newExerciseIds = [];
      } catch (err) {
        console.error(tag + ': delete previous exercise datapoint(s) failed: ' + err);
        exerciseFailed = true;
        // Keep newExerciseIds = split.exercise so the next sync retries delete.
      }
    } else {
      newExerciseIds = [];
    }
    if (!exerciseFailed && SYNC_EXERCISES && row.exercises.length > 0) {
      try {
        const ex = timing.exercise;
        const notes = buildNotes(row.exercises);
        const name = createExerciseAt(ex.startUtcMs, ex.startOffsetSeconds,
          ex.endUtcMs, ex.endOffsetSeconds, notes);
        if (name) {
          newExerciseIds.push(name);
          // Same rationale as the weight write above: persist before any
          // later step can fail and leave the datapoint untracked.
          writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
        }
        console.info(tag + ': createExerciseAt -> ' + (name || '<no name>'));
      } catch (err) {
        console.error(tag + ': createExerciseAt failed: ' + err);
        exerciseFailed = true;
      }
    }
  }

  writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
  if (exerciseReady) {
    writeMatchedHealthSession(row.rowNum, cols.matchedHealthSessionCol, match ? match.name : '');
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
