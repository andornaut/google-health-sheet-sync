// Orchestration tests for the stateful glue that the pure-helper tests in
// Parser.test.gs can't reach: onEditMarkDirty's column-aware dirty marking,
// syncDirtyRows' dirty-flag lifecycle (drain / error-retained / concurrent-edit),
// syncOneRow_'s phase dispatch + idempotency skip, and weight orphan
// reconciliation. Runs against the in-memory Apps Script fakes wired up in
// run.js (SYNC_TEST_HARNESS_); the Health API functions are stubbed per-test via
// globalThis, the same pattern resolveForeignMatches_'s tests use for
// listStrengthOnDate.
function runSyncTests() {
  const results = [];
  const t = (name, fn) => {
    try { fn(); results.push('PASS ' + name); }
    catch (err) { results.push('FAIL ' + name + ': ' + (err && err.stack || err)); }
  };
  const eq = (a, b, msg) => {
    const sa = JSON.stringify(a), sb = JSON.stringify(b);
    if (sa !== sb) throw new Error((msg || 'mismatch') + ' expected ' + sb + ' got ' + sa);
  };
  const ok = (cond, msg) => { if (!cond) throw new Error(msg || 'expected truthy'); };

  const SHEET = SYNC_TEST_HARNESS_.sheet;
  const PROPS = SYNC_TEST_HARNESS_.scriptProps;
  const LOCK = SYNC_TEST_HARNESS_.lockState;
  const ACTIVE = SYNC_TEST_HARNESS_.activeSheetRef;

  // Shared column layout for the fake sheet. 1-based column numbers in COL so
  // they can be passed straight to getRange.
  const HEADERS = ['Date', 'Weight', 'Bench', 'Exercise Synced At', 'Weight Synced At',
    'Created Health IDs', 'Exercise First Edited At', 'Exercises Last Edited At',
    'Weight Edited At', 'Matched Health Session'];
  const COL = {};
  HEADERS.forEach((h, i) => { COL[h] = i + 1; });

  const TOASTS = SYNC_TEST_HARNESS_.toasts;

  const reset = dataRows => {
    SHEET._setGrid([HEADERS.slice()].concat(dataRows || []));
    SHEET._setSelection(null);
    ACTIVE.sheet = SHEET;
    PROPS._clear();
    TOASTS.length = 0;
    LOCK.held = false;
  };
  const cell = header => SHEET.getRange(2, COL[header]).getValue();

  // Swap a set of globals (Health API stubs), run fn, restore. Returns fn's value.
  const withStubs = (stubs, fn) => {
    const saved = {};
    Object.keys(stubs).forEach(k => { saved[k] = globalThis[k]; globalThis[k] = stubs[k]; });
    try { return fn(); } finally { Object.keys(saved).forEach(k => { globalThis[k] = saved[k]; }); }
  };
  const NO_FOREIGN = { listStrengthOnDate: () => [] };

  // ---- onEditMarkDirty: column-aware dirty marking -----------------------

  t('onEditMarkDirty exercise edit clears exercise stamp, seeds edit timestamps, leaves weight alone', () => {
    reset([['2026-01-15', '', '135x5', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL['Bench']) });
    ok(marked === true, 'returns true');
    eq(cell('Exercise Synced At'), '', 'exercise synced cleared');
    eq(cell('Weight Synced At'), 'PREV-WT', 'weight synced untouched');
    ok(cell('Exercise First Edited At') !== '', 'first edited seeded');
    ok(cell('Exercises Last Edited At') !== '', 'last edited seeded');
    eq(cell('Weight Edited At'), '', 'weight edited untouched');
    ok(PROPS.getProperty('pendingDirty') !== null, 'pending flag set');
  });

  t('onEditMarkDirty weight edit clears weight stamp, sets weight edited, leaves exercise timestamps alone', () => {
    reset([['2026-01-15', '185', '', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL['Weight']) });
    ok(marked === true, 'returns true');
    eq(cell('Weight Synced At'), '', 'weight synced cleared');
    eq(cell('Exercise Synced At'), 'PREV-EX', 'exercise synced untouched');
    ok(cell('Weight Edited At') !== '', 'weight edited set');
    eq(cell('Exercise First Edited At'), '', 'exercise first edited NOT seeded by weight edit');
    eq(cell('Exercises Last Edited At'), '', 'exercise last edited NOT advanced by weight edit');
  });

  // Clearing real content is an edit: it must reach the delete paths in
  // syncOneRow_ (bodyweight cleared -> DELETE, exercises cleared ->
  // delete-only). oldValue is only supplied for single-cell edits, so a
  // multi-cell clear stays a no-op.
  t('onEditMarkDirty single-cell clear of an exercise value marks the row dirty', () => {
    reset([['2026-01-15', '', '', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL['Bench']), oldValue: '135x5x3' });
    ok(marked === true, 'returns true');
    eq(cell('Exercise Synced At'), '', 'exercise synced cleared');
    eq(cell('Weight Synced At'), 'PREV-WT', 'weight synced untouched');
    ok(cell('Exercises Last Edited At') !== '', 'last edited advanced');
  });

  t('onEditMarkDirty single-cell clear of the bodyweight marks the row dirty', () => {
    reset([['2026-01-15', '', '', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL['Weight']), oldValue: '185' });
    ok(marked === true, 'returns true');
    eq(cell('Weight Synced At'), '', 'weight synced cleared');
    eq(cell('Exercise Synced At'), 'PREV-EX', 'exercise synced untouched');
    ok(cell('Weight Edited At') !== '', 'weight edited set');
  });

  t('onEditMarkDirty clearing an already-blank cell stays a no-op', () => {
    reset([['2026-01-15', '', '', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    ok(onEditMarkDirty({ range: SHEET.getRange(2, COL['Bench']), oldValue: '' }) === false, 'no oldValue content');
    ok(onEditMarkDirty({ range: SHEET.getRange(2, COL['Bench']) }) === false, 'no oldValue at all');
    eq(cell('Exercise Synced At'), 'PREV-EX', 'stamp untouched');
  });

  t('onEditMarkDirty multi-cell clear stays a no-op (no oldValue for ranges)', () => {
    reset([['2026-01-15', '', '', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    const range = SHEET.getRange(2, COL['Weight'], 1, 2);   // Weight + Bench
    ok(onEditMarkDirty({ range: range }) === false, 'returns false');
    eq(cell('Exercise Synced At'), 'PREV-EX', 'stamps untouched');
  });

  // readRows ignores blank-header columns, so a scratch column parked to the
  // right of Weight has no exercise content to sync. Marking the row dirty for
  // it would stretch the 'edit' interval and recreate the datapoint for nothing.
  t('onEditMarkDirty ignores a blank-header scratch column', () => {
    const scratchCol = HEADERS.length + 1;
    SHEET._setGrid([
      HEADERS.concat(['']),
      ['2026-01-15', '', '135x5', 'PREV-EX', 'PREV-WT', '', '', '', '', '', 'note to self']
    ]);
    SHEET._setSelection(null);
    ACTIVE.sheet = SHEET;
    PROPS._clear();
    LOCK.held = false;
    ok(onEditMarkDirty({ range: SHEET.getRange(2, scratchCol) }) === false, 'typing there is a no-op');
    ok(onEditMarkDirty({ range: SHEET.getRange(2, scratchCol), oldValue: 'note to self' }) === false,
      'clearing it is a no-op too');
    eq(cell('Exercise Synced At'), 'PREV-EX', 'exercise stamp untouched');
    eq(cell('Exercises Last Edited At'), '', 'edit timestamp not advanced');
  });

  // Blankness must be decided the same way readRows decides it, or a header
  // cell holding a falsy-but-real value becomes a column whose content is read
  // by readRows while its edits never mark the row dirty.
  t('onEditMarkDirty treats a numeric-zero header as a real exercise column', () => {
    const headers = HEADERS.slice();
    headers[2] = 0;   // the exercise column's header is the number 0
    SHEET._setGrid([headers, ['2026-01-15', '', '135x5', 'PREV-EX', 'PREV-WT', '', '', '', '', '']]);
    SHEET._setSelection(null);
    ACTIVE.sheet = SHEET;
    PROPS._clear();
    LOCK.held = false;
    ok(onEditMarkDirty({ range: SHEET.getRange(2, 3) }) === true, 'marks the row dirty');
    eq(SHEET.getRange(2, COL['Exercise Synced At']).getValue(), '', 'exercise stamp cleared');
  });

  t('onEditMarkDirty date-only edit is a no-op', () => {
    reset([['2026-01-15', '', '', '', '', '', '', '', '', '']]);
    const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL['Date']) });
    ok(marked === false, 'returns false');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag');
  });

  t('onEditMarkDirty ignores edits on a different sheet', () => {
    reset([['2026-01-15', '185', '', '', '', '', '', '', '', '']]);
    const otherSheetRange = { getSheet: () => ({ getSheetId: () => 99999 }) };
    ok(onEditMarkDirty({ range: otherSheetRange }) === false, 'returns false for foreign sheet');
  });

  t('onEditMarkDirty keeps Exercise First Edited At sticky but advances Exercises Last Edited At', () => {
    reset([['2026-01-15', '', '135x5', '', '', '', 'STICKY-FIRST', 'OLD-LAST', '', '']]);
    onEditMarkDirty({ range: SHEET.getRange(2, COL['Bench']) });
    eq(cell('Exercise First Edited At'), 'STICKY-FIRST', 'first edited stays sticky');
    ok(cell('Exercises Last Edited At') !== 'OLD-LAST', 'last edited advanced');
  });

  // ---- readRows: the full-sheet id/session lists -------------------------

  // The contract every ownership decision depends on. A row with no parseable
  // Date is not syncable and is absent from `rows`, but it still OWNS its
  // datapoints and still holds its foreign session, so both must appear in the
  // full-sheet lists. Deriving either from `rows` instead makes orphan
  // reconciliation delete live datapoints and lets foreign matching hand the
  // same session to another row.
  t('readRows reports ids and matched sessions of rows it drops for a blank Date', () => {
    const undatedWeight = 'users/me/dataTypes/weight/dataPoints/W-undated';
    const datedExercise = 'users/me/dataTypes/exercise/dataPoints/E-dated';
    reset([
      ['', '185', '', 'SYNC', 'SYNC', JSON.stringify([undatedWeight]), '', '', '', 'foreign/undated'],
      ['2026-01-16', '', '135x5', 'SYNC', 'SYNC', JSON.stringify([datedExercise]), '', '', '', 'foreign/dated']
    ]);
    const r = readRows();
    eq(r.rows.map(row => row.rowNum), [3], 'only the dated row is syncable');
    eq(r.allHealthIds, [undatedWeight, datedExercise], 'ids from BOTH rows, dropped one included');
    eq(r.allMatchedSessions,
      [{ rowNum: 2, name: 'foreign/undated' }, { rowNum: 3, name: 'foreign/dated' }],
      'matched sessions from BOTH rows');
  });

  // ---- syncDirtyRows: lifecycle ------------------------------------------

  t('syncDirtyRows returns null when the lock is held', () => {
    reset([['2026-01-15', '185', '', 'SYNC', '', '', '', '', '', '']]);
    LOCK.held = true;
    const r = withStubs(NO_FOREIGN, () => syncDirtyRows(0));
    eq(r, null, 'skips when lock held');
  });

  t('syncDirtyRows with no dirty rows returns zero counts', () => {
    reset([['2026-01-15', '185', '', 'SYNC', 'SYNC', '[]', '', '', '', '']]);
    const r = withStubs(NO_FOREIGN, () => syncDirtyRows(0));
    eq(r, { ok: 0, errors: 0 }, 'nothing to do');
  });

  t('syncDirtyRows weight first-sync POSTs, stamps, persists ID, drains the flag', () => {
    reset([['2026-01-15', '185', '', 'SYNC', '', '', '', '', '', '']]);
    PROPS.setProperty('pendingDirty', 'GEN1');
    const calls = [];
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: (utc, off, lbs) => { calls.push(['createWeightAt', lbs]); return 'users/me/dataTypes/weight/dataPoints/W1'; }
    }), () => syncDirtyRows(0));
    eq(r.ok, 1, 'one row synced');
    eq(calls, [['createWeightAt', 185]], 'POSTed bodyweight 185');
    eq(cell('Created Health IDs'), JSON.stringify(['users/me/dataTypes/weight/dataPoints/W1']), 'ID persisted');
    ok(cell('Weight Synced At') !== '', 'weight synced stamped');
    ok(PROPS.getProperty('pendingDirty') === null, 'flag cleared after a clean drain');
  });

  t('syncDirtyRows keeps the dirty flag set when a row errors', () => {
    reset([['2026-01-15', '185', '', 'SYNC', '', '', '', '', '', '']]);
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => { throw new Error('boom'); }
    }), () => syncDirtyRows(0));
    eq(r.ok, 0, 'no rows synced');
    eq(r.errors, 1, 'one error');
    ok(PROPS.getProperty('pendingDirty') !== null, 'flag retained so the next poll retries');
    eq(cell('Weight Synced At'), '', 'weight stamp not written on failure');
  });

  t('syncDirtyRows preserves a concurrent-edit generation rather than clearing it', () => {
    reset([['2026-01-15', '185', '', 'SYNC', '', '', '', '', '', '']]);
    PROPS.setProperty('pendingDirty', 'GEN1');
    // A stub that advances the generation mid-pass simulates an edit landing
    // after readRows snapshotted. End-of-pass must NOT clear the flag.
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => { PROPS.setProperty('pendingDirty', 'GEN2'); return 'users/me/dataTypes/weight/dataPoints/W1'; }
    }), () => syncDirtyRows(0));
    eq(r.ok, 1, 'row synced');
    eq(PROPS.getProperty('pendingDirty'), 'GEN2', 'concurrent-edit generation kept');
  });

  // An unexpected throw out of syncOneRow_ (sheet I/O, not a Health API call)
  // must not abort the pass: rows are processed newest-first, so aborting would
  // strand every older row behind the failure on every subsequent pass. The row
  // is isolated, the rest of the pass completes, and a summary throw still
  // routes the failure through the unrecoverable path (owner email, flag kept).
  t('syncDirtyRows isolates an unexpected per-row failure and still reports it', () => {
    const older = new Date(Date.now() - 24 * 60 * 60 * 1000);
    reset([
      [older, '185', '', 'SYNC', '', '', '', '', '', ''],      // row 2, older
      [new Date(), '186', '', 'SYNC', '', '', '', '', '', '']  // row 3, newest -> processed first
    ]);
    PROPS.setProperty('pendingDirty', 'GEN1');
    const synced = [];
    let thrown = null;
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: (utc, off, lbs) => 'users/me/dataTypes/weight/dataPoints/W-' + lbs,
      writeHealthIds: (rowNum, col, names) => {
        if (rowNum === 3) throw new Error('simulated Spreadsheets service failure');
        if (synced.indexOf(rowNum) === -1) synced.push(rowNum);   // called twice per row
      }
    }), () => {
      try { syncDirtyRows(0); } catch (err) { thrown = err; }
    });
    ok(thrown !== null, 'throws so Apps Script emails the owner');
    ok(String(thrown).indexOf('row 3') !== -1, 'message names the failing row: ' + thrown);
    eq(synced, [2], 'the older row still synced instead of being stranded');
    ok(SHEET.getRange(2, COL['Weight Synced At']).getValue() !== '', 'older row stamped');
    eq(SHEET.getRange(3, COL['Weight Synced At']).getValue(), '', 'failing row left dirty');
    ok(PROPS.getProperty('pendingDirty') !== null, 'dirty flag kept so the next poll retries');
  });

  // The unrecoverable path skips end-of-pass flag resolution, so it only
  // preserves a flag that is already there. runSyncNow is the one entry point
  // that can start a pass with no flag set; without an explicit
  // markPendingDirty_() before the summary throw, the failed row is left dirty
  // with nothing scheduled to retry it (flushPending short-circuits, and the
  // backstop only re-dirties exercise rows with sendable content).
  t('syncDirtyRows SETS the dirty flag when an unexpected failure hits a pass that had none', () => {
    reset([['2026-01-15', '185', '', 'SYNC', '', '', '', '', '', '']]);
    ok(PROPS.getProperty('pendingDirty') === null, 'precondition: no flag at pass start');
    let thrown = null;
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => 'users/me/dataTypes/weight/dataPoints/W1',
      writeHealthIds: () => { throw new Error('simulated Spreadsheets service failure'); }
    }), () => {
      try { syncDirtyRows(0); } catch (err) { thrown = err; }
    });
    ok(thrown !== null, 'failure reported');
    ok(PROPS.getProperty('pendingDirty') !== null, 'flag created, so a later poll retries the row');
  });

  // No failure count stops the pass. A row that throws deterministically would
  // do so on every pass, so any early stop strands the rows behind it. Rows
  // are processed newest-first, and the same rows lead every time. Orphan
  // reconciliation, not a stopping rule, is what handles the datapoints created
  // by rows whose id write failed.
  const outageGrid = count => {
    const day = 24 * 60 * 60 * 1000;
    const now = Date.now();
    const grid = [];
    for (let i = count; i >= 1; i--) {   // ascending dates down the sheet
      grid.push([new Date(now - i * day), '1' + (80 + i), '', 'SYNC', '', '', '', '', '', '']);
    }
    return grid;
  };

  // The toast truncates, so a broad failure must lead with the counts rather
  // than a wall of row errors that pushes the summary off the end.
  t('syncDirtyRows summary leads with the counts and trims the row list', () => {
    reset(outageGrid(8));
    let thrown = null;
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => 'users/me/dataTypes/weight/dataPoints/W',
      writeHealthIds: () => { throw new Error('simulated Spreadsheets outage'); }
    }), () => {
      try { syncDirtyRows(0); } catch (err) { thrown = err; }
    });
    const msg = String(thrown.message);
    ok(msg.indexOf('0 synced, 8 error(s)') === 0, 'counts come first: ' + msg);
    ok(msg.indexOf('+3 more (see Executions)') !== -1, 'row list trimmed with the remainder named: ' + msg);
  });

  t('syncDirtyRows attempts every ready row even when all of them fail', () => {
    reset(outageGrid(6));
    let created = 0;
    let thrown = null;
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => { created++; return 'users/me/dataTypes/weight/dataPoints/W' + created; },
      writeHealthIds: () => { throw new Error('simulated Spreadsheets outage'); }
    }), () => {
      try { syncDirtyRows(0); } catch (err) { thrown = err; }
    });
    eq(created, 6, 'no row was skipped because earlier ones failed');
    ok(String(thrown).indexOf('0 synced, 6 error(s)') !== -1,
      'all six reported in one summary: ' + thrown);
    ok(PROPS.getProperty('pendingDirty') !== null, 'flag kept so the next pass retries');
  });

  // The summary throw replaces the normal return, so anything not interpolated
  // into its message is lost. deferredCount is the easiest to forget: it only
  // appears when the row cap and an unexpected failure coincide.
  t('syncDirtyRows summary reports the deferred backlog alongside the failure', () => {
    const day = 24 * 60 * 60 * 1000;
    const now = Date.now();
    const grid = [];
    for (let i = MAX_ROWS_PER_SYNC + 1; i >= 1; i--) {
      grid.push([new Date(now - i * day), '185', '', 'SYNC', '', '', '', '', '', '']);
    }
    reset(grid);
    const newestRow = grid.length + 1;   // ascending dates, so the last row is newest
    let thrown = null;
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => 'users/me/dataTypes/weight/dataPoints/W',
      writeHealthIds: rowNum => {
        if (rowNum === newestRow) throw new Error('simulated Spreadsheets service failure');
      }
    }), () => {
      try { syncDirtyRows(0); } catch (err) { thrown = err; }
    });
    ok(String(thrown).indexOf('1 deferred by the row cap') !== -1,
      'deferred backlog surfaced, not swallowed by the throw: ' + thrown);
  });

  t('syncDirtyRows keeps syncing healthy rows past several failing ones', () => {
    // 8 rows, every other one failing: the healthy rows must still sync rather
    // than being stranded behind the failures ahead of them.
    reset(outageGrid(8));
    let thrown = null;
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => 'users/me/dataTypes/weight/dataPoints/W',
      writeHealthIds: rowNum => {
        if (rowNum % 2 === 0) throw new Error('simulated row-specific write failure');
      }
    }), () => {
      try { syncDirtyRows(0); } catch (err) { thrown = err; }
    });
    ok(thrown !== null, 'failures still reported');
    ok(String(thrown).indexOf('4 synced, 4 error(s)') !== -1,
      'every row attempted, counts carried in the message: ' + thrown);
    ok(PROPS.getProperty('pendingDirty') !== null, 'flag kept so the next pass retries');
  });

  // ---- syncOneRow_ via syncDirtyRows: exercise idempotency ---------------

  t('syncDirtyRows skips the delete+recreate when the prior exercise datapoint is unchanged', () => {
    const startMs = Date.UTC(2026, 0, 15, 17, 0, 0);
    const endMs = Date.UTC(2026, 0, 15, 17, 30, 0);
    const exercises = [{ name: 'Bench', entries: [{ weight: 135, reps: 5, sets: 3, assisted: false }] }];
    const priorNotes = buildNotes(endMs - startMs, exercises);
    const priorName = 'users/me/dataTypes/exercise/dataPoints/E1';
    reset([['2026-01-15', '', '135x5x3', '', 'SYNC', JSON.stringify([priorName]), '', '', '', '']]);
    const calls = [];
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      getDataPoint: () => ({ exercise: { interval: {
        startTime: new Date(startMs).toISOString(), endTime: new Date(endMs).toISOString()
      }, notes: priorNotes } }),
      deleteDataPointsByName: names => { calls.push(['delete', names]); },
      createExerciseAt: () => { calls.push(['create']); return 'users/me/dataTypes/exercise/dataPoints/E2'; }
    }), () => syncDirtyRows(0));
    eq(r.ok, 1, 'row counted ok');
    eq(calls, [], 'no delete or create issued for an unchanged exercise');
    eq(cell('Created Health IDs'), JSON.stringify([priorName]), 'resource name preserved');
    ok(cell('Exercise Synced At') !== '', 'exercise synced stamped');
  });

  t('syncDirtyRows recreates the exercise datapoint when notes changed', () => {
    const startMs = Date.UTC(2026, 0, 15, 17, 0, 0);
    const endMs = Date.UTC(2026, 0, 15, 17, 30, 0);
    const priorName = 'users/me/dataTypes/exercise/dataPoints/E1';
    // Prior notes describe a different set/rep scheme than the current cell.
    reset([['2026-01-15', '', '135x5x3', '', 'SYNC', JSON.stringify([priorName]), '', '', '', '']]);
    const calls = [];
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      getDataPoint: () => ({ exercise: { interval: {
        startTime: new Date(startMs).toISOString(), endTime: new Date(endMs).toISOString()
      }, notes: 'Bench, 999 lbs, 1 set of 1' } }),
      deleteDataPointsByName: names => { calls.push(['delete', names.slice()]); },
      createExerciseAt: () => { calls.push(['create']); return 'users/me/dataTypes/exercise/dataPoints/E2'; }
    }), () => syncDirtyRows(0));
    eq(r.ok, 1, 'row counted ok');
    eq(calls, [['delete', [priorName]], ['create']], 'deletes old then creates new');
    eq(cell('Created Health IDs'), JSON.stringify(['users/me/dataTypes/exercise/dataPoints/E2']), 'new resource name recorded');
  });

  // Clearing a bodyweight whose datapoint was already deleted in the Health app
  // must not wedge the row: retrying a delete that 404s forever would keep the
  // row dirty and the pending flag set, re-issuing the same call every poll.
  t('syncDirtyRows treats a 404 on the weight delete as already deleted', () => {
    const priorName = 'users/me/dataTypes/weight/dataPoints/W-gone';
    reset([['2026-01-15', '', '', 'SYNC', '', JSON.stringify([priorName]), '', '', '', '']]);
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      deleteDataPointsByName: () => {
        const err = new Error('Health API POST ... -> 404: not found');
        err.statusCode = 404;
        throw err;
      }
    }), () => syncDirtyRows(0));
    eq(r.ok, 1, 'row counted ok');
    eq(cell('Created Health IDs'), JSON.stringify([]), 'stale weight id dropped');
    ok(cell('Weight Synced At') !== '', 'weight stamped instead of retrying forever');
    ok(PROPS.getProperty('pendingDirty') === null, 'queue drained');
  });

  // Deletes go one name per call so a 404 is attributed to the name that is
  // actually gone. A batched :batchDelete fails as a unit, so treating that
  // failure as "all deleted" would drop still-live siblings from the sheet.
  t('syncDirtyRows keeps a live sibling when only one prior weight id is gone', () => {
    const gone = 'users/me/dataTypes/weight/dataPoints/W-gone';
    const live = 'users/me/dataTypes/weight/dataPoints/W-live';
    reset([['2026-01-15', '', '', 'SYNC', '', JSON.stringify([gone, live]), '', '', '', '']]);
    const deleted = [];
    const r = withStubs(Object.assign({}, NO_FOREIGN, {
      deleteDataPointsByName: names => {
        if (names[0] === gone) {
          const err = new Error('Health API POST ... -> 404: not found');
          err.statusCode = 404;
          throw err;
        }
        deleted.push(names[0]);
      }
    }), () => syncDirtyRows(0));
    eq(r.ok, 1, 'row counted ok');
    eq(deleted, [live], 'the live sibling was actually deleted, not assumed gone');
    eq(cell('Created Health IDs'), JSON.stringify([]), 'both ids resolved');
  });

  // The accepted tradeoff of running every pass to completion, stated
  // end-to-end so it reads as a known design point rather than a latent bug.
  // A row whose id write fails has already created its datapoint, so the id is
  // lost and the NEXT pass creates another: the leak no failure threshold can
  // bound (stopping early only trades it for stranding the rows behind it).
  // Orphan reconciliation is what closes the loop, but only for rows dated
  // within ORPHAN_RECONCILE_LOOKBACK_DAYS, which is what this covers (the rows
  // are dated today/yesterday). Candidates are listed by the datapoint's own
  // civil date, i.e. the row's Date, so a leak on an older row is never listed
  // and never reclaimed; that residual is documented, not tested, because no
  // code path closes it.
  t('an untracked create is re-made next pass and later reclaimed by reconciliation', () => {
    const tracked = 'users/me/dataTypes/weight/dataPoints/W-tracked';
    reset([
      [new Date(Date.now() - 24 * 60 * 60 * 1000), '185', '', 'SYNC', '', '', '', '', '', ''],
      [new Date(), '', '', 'SYNC', 'SYNC', JSON.stringify([tracked]), '', '', '', '']
    ]);
    const created = [];
    const brokenWrite = Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => {
        const name = 'users/me/dataTypes/weight/dataPoints/W-leaked' + (created.length + 1);
        created.push(name);
        return name;
      },
      writeHealthIds: () => { throw new Error('simulated Spreadsheets service failure'); }
    });
    withStubs(brokenWrite, () => { try { syncDirtyRows(0); } catch (err) { /* reported */ } });
    withStubs(brokenWrite, () => { try { syncDirtyRows(0); } catch (err) { /* reported */ } });
    eq(created.length, 2, 'each pass created a datapoint it could not record');
    eq(cell('Created Health IDs'), '', 'neither id reached the sheet');

    // The backstop reclaims both: they carry our client id and no row
    // references them. The tracked datapoint on the other row establishes
    // ownership and must survive.
    const deleted = [];
    let listed = false;
    withStubs({
      listStrengthOnDate: () => [],
      listWeightOnDate: () => {
        if (listed) return [];   // one civil day's worth, not once per lookback day
        listed = true;
        return [tracked, created[0], created[1]].map(name => ({ name: name, googleWebClientId: 'ours' }));
      },
      deleteDataPointsByName: names => { deleted.push.apply(deleted, names); }
    }, () => backstop());
    eq(deleted, created, 'both leaked datapoints reclaimed, tracked one spared');
  });

  // ---- weight orphan reconciliation --------------------------------------

  t('reconcileWeightOrphans_ deletes our untracked weight datapoint, spares tracked/foreign', () => {
    const tracked = 'users/me/dataTypes/weight/dataPoints/W-tracked';
    const orphan = 'users/me/dataTypes/weight/dataPoints/W-orphan';
    const device = 'users/me/dataTypes/weight/dataPoints/W-device';
    const deleted = [];
    withStubs({
      listWeightOnDate: () => ([
        { name: tracked, googleWebClientId: 'ours' },
        { name: orphan, googleWebClientId: 'ours' },
        { name: device, googleWebClientId: null }
      ]),
      deleteDataPointsByName: names => { deleted.push.apply(deleted, names); }
    }, () => {
      reconcileWeightOrphans_([tracked], Date.UTC(2026, 0, 15, 12, 0, 0), 1);
    });
    eq(deleted, [orphan], 'only our untracked weight datapoint is deleted');
  });

  // A row whose Date cell is blank is dropped by readRows, but its datapoints
  // are still live and still referenced by the sheet. Reconciling against the
  // full-sheet id list (readRows' allHealthIds) rather than the surviving rows
  // is what keeps the backstop from deleting them.
  t('backstop spares datapoints of a row whose Date is blank', () => {
    const liveEx = 'users/me/dataTypes/exercise/dataPoints/E-undated';
    const liveWt = 'users/me/dataTypes/weight/dataPoints/W-undated';
    const trackedEx = 'users/me/dataTypes/exercise/dataPoints/E-tracked';
    const trackedWt = 'users/me/dataTypes/weight/dataPoints/W-tracked';
    reset([
      ['', '185', '135x5x3', 'SYNC', 'SYNC', JSON.stringify([liveEx, liveWt]), '', '', '', ''],
      [new Date(), '', '225x5x3', 'SYNC', 'SYNC', JSON.stringify([trackedEx, trackedWt]), '', '', '', '']
    ]);
    const deleted = [];
    withStubs({
      listStrengthOnDate: () => ([
        { name: liveEx, googleWebClientId: 'ours' },
        { name: trackedEx, googleWebClientId: 'ours' }
      ]),
      listWeightOnDate: () => ([
        { name: liveWt, googleWebClientId: 'ours' },
        { name: trackedWt, googleWebClientId: 'ours' }
      ]),
      deleteDataPointsByName: names => { deleted.push.apply(deleted, names); }
    }, () => backstop());
    eq(deleted, [], 'nothing deleted: the undated row still owns both datapoints');
  });

  // ---- backstop: re-dirties BOTH matched and unmatched recent rows --------

  // backstop() uses the real Date.now() (not injectable) for its lookback
  // window, so rows must be dated "now". A Date object in the Date cell is read
  // back verbatim by toDate_, sidestepping the UTC-midnight civil-date boundary
  // that a 'yyyy-MM-dd' string would hit in a non-UTC script time zone.
  t('backstop re-dirties BOTH matched and unmatched recent rows and sets the pending flag', () => {
    // Distinct ascending dates (yesterday, today) so the date validation pass
    // doesn't reject the sheet; both fall inside BACKSTOP_LOOKBACK_DAYS = 2.
    const today = new Date();
    const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
    reset([
      [yesterday, '', '135x5x3', 'SYNC', '', '', '', '', '', ''],        // row 2: unmatched
      [today, '', '225x5x3', 'SYNC', '', '', '', '', '', 'foreign/F1']   // row 3: matched
    ]);
    withStubs({ listStrengthOnDate: () => [], listWeightOnDate: () => [] }, () => backstop());
    eq(SHEET.getRange(2, COL['Exercise Synced At']).getValue(), '', 'unmatched row re-dirtied');
    eq(SHEET.getRange(3, COL['Exercise Synced At']).getValue(), '', 'matched row re-dirtied');
    ok(PROPS.getProperty('pendingDirty') !== null, 'pending flag set for the next poll');
  });

  // ---- resyncSelectedRows: selection guards --------------------------------

  // The selection is spreadsheet-global, so row numbers from another tab would
  // re-dirty unrelated rows on the synced tab.
  t('resyncSelectedRows refuses a selection on another tab', () => {
    reset([['2026-01-15', '', '135x5', 'SYNC', 'SYNC', '', '', '', '', '']]);
    SHEET._setSelection([[2, 1]]);
    ACTIVE.sheet = SYNC_TEST_HARNESS_.otherSheet;
    try { withStubs(NO_FOREIGN, () => resyncSelectedRows()); }
    finally { ACTIVE.sheet = SHEET; }
    eq(cell('Exercise Synced At'), 'SYNC', 'stamps not cleared');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag');
    ok(TOASTS.length === 1 && TOASTS[0].indexOf('tab first') !== -1,
      'told the user why nothing happened: ' + JSON.stringify(TOASTS));
  });

  // Clicking a column header selects the full sheet height. Without clamping,
  // every empty row below the data would be re-dirtied and stamped.
  // Asserts on which rows were targeted, not on the resulting cell values:
  // clearing a stamp writes '', which is indistinguishable from an untouched
  // empty cell below the data, so a value-based assertion passes either way.
  t('resyncSelectedRows clamps a whole-column selection to the data range', () => {
    reset([
      ['2026-01-15', '', '135x5', 'SYNC', 'SYNC', '[]', '', '', '', ''],
      ['2026-01-16', '', '225x5', 'SYNC', 'SYNC', '[]', '', '', '', '']
    ]);
    SHEET._setSelection([[1, 1000]]);   // whole column, header row included
    const cleared = [];
    withStubs(Object.assign({}, NO_FOREIGN, {
      createExerciseAt: () => 'users/me/dataTypes/exercise/dataPoints/E',
      clearRowExerciseSynced: rowNum => { cleared.push(rowNum); },
      clearRowWeightSynced: () => {}
    }), () => resyncSelectedRows());
    eq(cleared, [2, 3], 'only the two data rows were re-dirtied, not the full sheet height');
  });

  t('resyncSelectedRows toasts instead of throwing when nothing is selected', () => {
    reset([['2026-01-15', '', '135x5', 'SYNC', 'SYNC', '', '', '', '', '']]);
    SHEET._setSelection(null);
    withStubs(NO_FOREIGN, () => resyncSelectedRows());
    eq(cell('Exercise Synced At'), 'SYNC', 'stamps not cleared');
    ok(TOASTS.length === 1 && TOASTS[0].indexOf('No data rows selected') !== -1,
      'told the user why nothing happened: ' + JSON.stringify(TOASTS));
  });

  // The re-dirty is followed by an immediate sync, so the observable effect is
  // the recreated datapoint rather than a cleared stamp.
  t('resyncSelectedRows re-dirties and re-pushes the selected rows on the synced tab', () => {
    reset([['2026-01-15', '', '135x5', 'SYNC', 'SYNC', '[]', '', '', '', '']]);
    SHEET._setSelection([[2, 1]]);
    const newName = 'users/me/dataTypes/exercise/dataPoints/E9';
    withStubs(Object.assign({}, NO_FOREIGN, {
      createExerciseAt: () => newName
    }), () => resyncSelectedRows());
    eq(cell('Created Health IDs'), JSON.stringify([newName]), 'exercise datapoint recreated');
    ok(cell('Exercise Synced At') !== 'SYNC', 'exercise stamp refreshed');
  });

  // ---- trigger-entry date validation --------------------------------------

  t('syncOnEdit throws on a date validation violation (uncaught -> owner email)', () => {
    reset([
      ['2026-01-15', '', '135x5', '', '', '', '', '', '', ''],
      ['2026-01-15', '', '225x5', '', '', '', '', '', '', '']   // duplicate date
    ]);
    let thrown = null;
    try { syncOnEdit({ range: SHEET.getRange(2, COL['Bench']) }); }
    catch (err) { thrown = err; }
    ok(thrown !== null, 'throws out of syncOnEdit');
    ok(String(thrown).indexOf('date validation failed') !== -1, 'message names the validation: ' + thrown);
    // Thrown before dirty marking, so the edit is not recorded.
    eq(cell('Exercises Last Edited At'), '', 'edit not dirty-marked');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag');
  });

  t('flushPending logs and skips (no throw) on a date validation violation', () => {
    reset([
      ['2026-01-16', '185', '', 'SYNC', '', '', '', '', '', ''],
      ['2026-01-15', '186', '', 'SYNC', '', '', '', '', '', '']   // out of order
    ]);
    PROPS.setProperty('pendingDirty', 'GEN1');
    const calls = [];
    withStubs(Object.assign({}, NO_FOREIGN, {
      createWeightAt: () => { calls.push('create'); return 'users/me/dataTypes/weight/dataPoints/W1'; }
    }), () => flushPending());
    eq(calls, [], 'no sync work attempted');
    eq(PROPS.getProperty('pendingDirty'), 'GEN1', 'dirty flag left so the backlog syncs after the fix');
  });

  t('backstop skips on a date validation violation', () => {
    reset([['2024-06-01', '', '135x5x3', 'SYNC', '', '', '', '', '', '']]);   // year below MIN
    withStubs({ listStrengthOnDate: () => [], listWeightOnDate: () => [] }, () => backstop());
    eq(cell('Exercise Synced At'), 'SYNC', 'row not re-dirtied');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag');
  });

  t('manual resyncAllRows toasts and aborts on a date validation violation', () => {
    reset([
      ['2026-01-15', '', '135x5', 'SYNC', 'SYNC', '', '', '', '', ''],
      ['2026-01-15', '', '225x5', 'SYNC', 'SYNC', '', '', '', '', '']   // duplicate date
    ]);
    withStubs(NO_FOREIGN, () => resyncAllRows());
    eq(cell('Exercise Synced At'), 'SYNC', 'stamps not cleared');
    eq(cell('Weight Synced At'), 'SYNC', 'weight stamps not cleared');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag');
    ok(TOASTS.length === 1 && TOASTS[0].indexOf('Date validation failed') !== -1,
      'surfaced the violation to the user: ' + JSON.stringify(TOASTS));
  });

  t('backstop skips entirely when the lock is held', () => {
    reset([[new Date(), '', '135x5x3', 'SYNC', '', '', '', '', '', '']]);
    LOCK.held = true;
    withStubs({ listStrengthOnDate: () => [], listWeightOnDate: () => [] }, () => backstop());
    eq(SHEET.getRange(2, COL['Exercise Synced At']).getValue(), 'SYNC', 'row untouched when lock held');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag when skipped');
  });

  const msg = results.join('\n');
  const passed = results.filter(r => r.startsWith('PASS ')).length;
  const summary = results.length + ' tests: ' + passed + ' passed, ' + (results.length - passed) + ' failed';
  console.log(msg + '\n\n' + summary);
}
