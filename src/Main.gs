function onOpen() {
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
  const html = '<p>Click the link below, sign in with the Google account that owns your Google Health data, and grant the requested scopes.</p>'
    + '<p><a href="' + url + '" target="_blank" rel="noopener">Authorize Google Health API</a></p>'
    + '<p>After the success page appears, you can close that tab and return here.</p>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(220), 'Authorize Google Health');
}

function revokeHealthApi() {
  resetHealthAuth();
  toast_('Google Health authorization cleared. Use "Authorize Health API" to grant again.', 10);
}

function setup() {
  ensureManagedColumns();
  installTriggers();
}

function installTriggers() {
  // 'backstop' is no longer installed but kept in the cleanup set so the
  // hourly trigger an earlier setup() installed gets removed on next run.
  const handlers = new Set(['onEditTrigger', 'flushIfPending', 'backstop']);
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
  if (props.getProperty(PENDING_DIRTY_KEY) !== '1') {
    console.info('flushIfPending: no pending edits, skipping');
    return;
  }
  console.info('flushIfPending: pending edits detected, syncing');
  syncDirtyRows(false, 0);
}

// "Run now" is an explicit manual action: bypasses the exercise quiesce
// window so the user sees results immediately. Weight already syncs without
// quiesce, so this only changes behavior for exercise content. If they keep
// editing afterward, the row goes dirty again and the next sync replaces the
// Health datapoint(s).
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
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER] || null;
  clearRowExerciseSynced(row, exerciseCol);
  clearRowWeightSynced(row, weightSyncedAtCol);
  SpreadsheetApp.flush();
  PropertiesService.getScriptProperties().setProperty(PENDING_DIRTY_KEY, '1');
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
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER] || null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    toast_('No data rows.', 10);
    return;
  }
  const dataRowCount = lastRow - 1;

  const blanks = [];
  for (let i = 0; i < dataRowCount; i++) blanks.push(['']);
  sheet.getRange(2, exerciseCol, dataRowCount, 1).setValues(blanks);
  if (weightSyncedAtCol) sheet.getRange(2, weightSyncedAtCol, dataRowCount, 1).setValues(blanks);
  SpreadsheetApp.flush();
  PropertiesService.getScriptProperties().setProperty(PENDING_DIRTY_KEY, '1');

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
  const totalSec = Math.round(ms / 1000);
  if (totalSec < 60) return totalSec + 's';
  const min = Math.floor(totalSec / 60);
  const sec = totalSec % 60;
  return sec === 0 ? min + 'm' : min + 'm ' + sec + 's';
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
  let deferredCount = 0;
  try {
    // Clear early so any onEditMarkDirty that fires during this pass can
    // re-set the flag and be picked up by the next poll without ambiguity.
    props.deleteProperty(PENDING_DIRTY_KEY);

    const { rows, exerciseSyncedAtCol, weightSyncedAtCol, weightCol, healthIdsCol, lastEditedAtCol, matchedHealthSessionCol } = readRows();
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
    //   - Weight phase: always ready when the row is weight-dirty (no quiesce).
    //   - Exercise phase: ready iff bypassed, or the row has no edit timestamp,
    //     or quiesce time has elapsed since the last edit. Rows with no
    //     exercise content also pass instantly since there's nothing to time.
    const now = Date.now();
    const ready = [];
    let maxRemainingMs = 0;
    dirty.forEach(r => {
      const weightReady = !r.weightSyncedAt;
      let exerciseReady = false;
      let remainingMs = 0;
      if (!r.exerciseSyncedAt) {
        if (bypassQuiesce || !r.lastEditedAt || r.exercises.length === 0) {
          exerciseReady = true;
        } else {
          const sinceMs = now - r.lastEditedAt.getTime();
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

    const exerciseReadyRows = ready.filter(r => r.exerciseReady).map(r => r.row);
    const matchPlan = resolveForeignMatches_(rows, exerciseReadyRows);
    const cols = {
      exerciseSyncedAtCol: exerciseSyncedAtCol,
      weightSyncedAtCol: weightSyncedAtCol,
      weightCol: weightCol,
      healthIdsCol: healthIdsCol,
      lastEditedAtCol: lastEditedAtCol,
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
    // Re-set the flag if work remains: quiescing rows need a future poll to
    // pick them up, and failed rows should be retried promptly. (syncOneRow_
    // also sets the flag itself when it defers a row due to a concurrent
    // edit, so even rows counted as ok can leave the flag set.)
    if (waitingCount > 0 || errors > 0 || deferredCount > 0) {
      props.setProperty(PENDING_DIRTY_KEY, '1');
    }
    lock.releaseLock();
  }
  return { ok: ok, errors: errors, deferred: deferredCount };
}

// Returns rowNum -> foreign Strength Training session for ready rows whose
// lifting content is already covered by a non-sync-created Health datapoint.
// Two-phase per date: time-range overlap (rows with edit timestamps) then
// 1:1 ordinal pairing (rows without). Candidates already claimed by other
// rows (foreign-matched or script-created) are excluded first so a session
// this script wrote can't be re-matched and incremental syncs can't
// double-claim.
//
// Ordering note: this runs once near the top of syncDirtyRows, before the
// per-row syncOneRow_ loop. The `allRows` snapshot comes from readRows() at
// the start of the pass, and `listForeignStrengthOnDate` calls the API
// before any row in this pass has had its create issued. So same-pass
// freshly-created datapoints are not yet visible to the API and not yet in
// any row's Created Health IDs — both gaps cancel out and there's no
// self-match risk.
function resolveForeignMatches_(allRows, readyRows) {
  const plan = {};
  const readyRowNums = {};
  readyRows.forEach(r => { readyRowNums[r.rowNum] = true; });
  const claimedByOthers = {};
  const claim_ = (date, name) => {
    const key = ymd(date);
    (claimedByOthers[key] = claimedByOthers[key] || {})[name] = true;
  };
  allRows.forEach(r => {
    if (!readyRowNums[r.rowNum] && r.matchedHealthSession) claim_(r.date, r.matchedHealthSession);
    splitHealthIdsByType_(r.healthIds).exercise.forEach(name => claim_(r.date, name));
  });
  const byDate = groupRowsByDate_(readyRows.filter(r => r.exercises.length > 0));
  Object.keys(byDate).forEach(dateKey => {
    const dayRows = byDate[dateKey];
    let candidates;
    try {
      candidates = listForeignStrengthOnDate(dayRows[0].date);
    } catch (err) {
      console.warn('resolveForeignMatches_: list failed for ' + dateKey + ': ' + err);
      return;
    }
    const claimed = claimedByOthers[dateKey];
    if (claimed) {
      const before = candidates.length;
      candidates = candidates.filter(c => !claimed[c.name]);
      const removed = before - candidates.length;
      if (removed > 0) {
        console.info('resolveForeignMatches_: ' + dateKey + ' excluded ' + removed
          + ' candidate(s) already claimed by other sheet row(s)');
      }
    }
    if (candidates.length === 0) return;

    const timeRangeRows = dayRows.filter(r => r.firstEditedAt && r.lastEditedAt);
    const ordinalRows = dayRows.filter(r => !(r.firstEditedAt && r.lastEditedAt));

    timeRangeRows.forEach(r => {
      const windowStart = r.firstEditedAt.getTime() - FOREIGN_MATCH_BUFFER_MS;
      const windowEnd = r.lastEditedAt.getTime() + FOREIGN_MATCH_BUFFER_MS;
      let bestIdx = -1;
      let bestOverlap = 0;
      candidates.forEach((c, i) => {
        const overlap = Math.min(c.endUtcMs, windowEnd) - Math.max(c.startUtcMs, windowStart);
        if (overlap > bestOverlap) {
          bestIdx = i;
          bestOverlap = overlap;
        }
      });
      if (bestIdx >= 0) {
        plan[r.rowNum] = candidates[bestIdx];
        console.info('resolveForeignMatches_: ' + dateKey + ' row ' + r.rowNum
          + ' time-range matches ' + candidates[bestIdx].name + ' (overlap=' + bestOverlap + 'ms)');
        candidates.splice(bestIdx, 1);
      }
    });

    if (ordinalRows.length === 0 || candidates.length === 0) return;
    if (ordinalRows.length !== candidates.length) {
      console.info('resolveForeignMatches_: ' + dateKey + ' has ' + ordinalRows.length
        + ' no-edit-time row(s) and ' + candidates.length
        + ' remaining foreign session(s); counts disagree, skipping ordinal pairing.');
      return;
    }
    ordinalRows.sort((a, b) => a.rowNum - b.rowNum);
    candidates.sort((a, b) => a.startUtcMs - b.startUtcMs);
    ordinalRows.forEach((r, i) => {
      plan[r.rowNum] = candidates[i];
      console.info('resolveForeignMatches_: ' + dateKey + ' row ' + r.rowNum
        + ' ordinal[' + i + '] matches ' + candidates[i].name);
    });
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

// Resolve the exercise interval (and weight sampleTime) for a row.
// Prefers edit-derived timing when First/Last Edited At are both present.
// Falls back to synthetic noon-ordinal otherwise.
function resolveRowTiming_(row, ordinal) {
  if (row.firstEditedAt && row.lastEditedAt) {
    const startMs = row.firstEditedAt.getTime();
    const rawDuration = row.lastEditedAt.getTime() - startMs;
    const clampedDuration = Math.min(
      Math.max(rawDuration, MIN_EXERCISE_DURATION_MS),
      MAX_EXERCISE_DURATION_MS
    );
    const endMs = startMs + clampedDuration;
    const tz = getTz_();
    const startOffset = getTzOffsetSeconds_(tz, row.firstEditedAt);
    const endOffset = getTzOffsetSeconds_(tz, new Date(endMs));
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

// Sync a single row in two independent phases (weight, exercise). Either or
// both phases may run on a given pass:
//   - Weight phase runs whenever the row's Weight Synced At is cleared. It
//     reconciles weight IDs with the sheet's bodyweight (write/delete) and
//     stamps Weight Synced At on success.
//   - Exercise phase runs only when the caller passed exerciseReady=true
//     (i.e. quiesce passed or no exercise content). It reconciles exercise
//     IDs with the sheet's exercises (or matches a foreign session) and
//     stamps Exercise Synced At on success.
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
    + (match ? ', match=' + match.name : '') + ')');

  let timing;
  try {
    timing = resolveRowTiming_(row, ordinal);
  } catch (err) {
    console.error(tag + ': resolveRowTiming_ failed: ' + err);
    return false;
  }
  console.info(tag + ': timing source=' + timing.source);

  const split = splitHealthIdsByType_(row.healthIds);
  let newWeightIds = split.weight;
  let newExerciseIds = split.exercise;
  let weightFailed = false;
  let exerciseFailed = false;
  let weightAttempted = false;
  let exerciseAttempted = false;

  if (weightReady) {
    weightAttempted = true;
    if (split.weight.length > 0) {
      console.info(tag + ': deleting ' + split.weight.length + ' previous weight datapoint(s)');
      try {
        deleteDataPointsByName(split.weight);
        newWeightIds = [];
      } catch (err) {
        console.error(tag + ': delete previous weight datapoint(s) failed: ' + err);
        weightFailed = true;
        // Keep newWeightIds = split.weight so the next sync retries delete.
      }
    } else {
      newWeightIds = [];
    }
    if (!weightFailed && SYNC_WEIGHT && row.bodyweight !== null) {
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
    }
  }

  if (exerciseReady) {
    exerciseAttempted = true;
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
      if (match) {
        console.info(tag + ': skipping exercise create; matched foreign ' + match.name);
      } else {
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
  }

  writeHealthIds(row.rowNum, cols.healthIdsCol, newWeightIds.concat(newExerciseIds).concat(split.other));
  if (exerciseAttempted) {
    writeMatchedHealthSession(row.rowNum, cols.matchedHealthSessionCol, match ? match.name : '');
  }

  // Concurrent-edit guards, phase-isolated. onEditMarkDirty only advances
  // Last Edited At on exercise-relevant edits, so the exercise phase uses
  // it as its "row changed during sync" signal. The weight phase compares
  // the bodyweight value we synced against the cell's current value — a
  // weight edit during sync doesn't touch Last Edited At, so we'd miss it
  // otherwise. The two phases are independent: a concurrent weight edit
  // doesn't defer the exercise stamp and vice versa.
  let exerciseConcurrentEdit = false;
  if (cols.lastEditedAtCol) {
    const currentLastEdit = toDate_(getSheet_().getRange(row.rowNum, cols.lastEditedAtCol).getValue());
    const previousMs = row.lastEditedAt ? row.lastEditedAt.getTime() : null;
    const currentMs = currentLastEdit ? currentLastEdit.getTime() : null;
    if (currentMs !== previousMs) {
      const prevLabel = row.lastEditedAt ? row.lastEditedAt.toISOString() : '<none>';
      const currLabel = currentLastEdit ? currentLastEdit.toISOString() : '<cleared>';
      console.info(tag + ': concurrent exercise edit detected (Last Edited At '
        + prevLabel + ' -> ' + currLabel + '); deferring Exercise Synced At stamp.');
      exerciseConcurrentEdit = true;
    }
  }
  let weightConcurrentEdit = false;
  if (weightAttempted && cols.weightCol) {
    const currentBodyweight = parseBodyweight(getSheet_().getRange(row.rowNum, cols.weightCol).getValue());
    if (currentBodyweight !== row.bodyweight) {
      console.info(tag + ': concurrent weight edit detected (bodyweight '
        + (row.bodyweight === null ? '<none>' : row.bodyweight) + ' -> '
        + (currentBodyweight === null ? '<cleared>' : currentBodyweight)
        + '); deferring Weight Synced At stamp.');
      weightConcurrentEdit = true;
    }
  }

  const stampIso = new Date().toISOString();
  if (weightAttempted && !weightFailed && !weightConcurrentEdit) {
    markRowWeightSynced(row.rowNum, cols.weightSyncedAtCol, stampIso);
  }
  if (exerciseAttempted && !exerciseFailed && !exerciseConcurrentEdit) {
    markRowExerciseSynced(row.rowNum, cols.exerciseSyncedAtCol, stampIso);
  }

  if (weightFailed || exerciseFailed) {
    console.warn(tag + ': FAILED (partial); will retry on next sync.');
    return false;
  }

  // If the row still has unstamped phases (either because we skipped a phase
  // this pass or a concurrent edit blocked the stamp), keep PENDING_DIRTY_KEY
  // set so a future poll picks it up.
  const weightStampMissing = !row.weightSyncedAt && !(weightAttempted && !weightConcurrentEdit);
  const exerciseStampMissing = !row.exerciseSyncedAt && !(exerciseAttempted && !exerciseConcurrentEdit);
  if (weightStampMissing || exerciseStampMissing) {
    PropertiesService.getScriptProperties().setProperty(PENDING_DIRTY_KEY, '1');
    console.info(tag + ': partial progress; row stays dirty (weightStamped='
      + (!weightStampMissing) + ', exerciseStamped=' + (!exerciseStampMissing) + ')');
  } else {
    console.info(tag + ': done');
  }
  return true;
}
