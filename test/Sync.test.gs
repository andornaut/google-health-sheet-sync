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

  // Shared column layout for the fake sheet. 1-based column numbers in COL so
  // they can be passed straight to getRange.
  const HEADERS = ['Date', 'Weight', 'Bench', 'Exercise Synced At', 'Weight Synced At',
    'Created Health IDs', 'Exercise First Edited At', 'Exercises Last Edited At',
    'Weight Edited At', 'Matched Health Session'];
  const COL = {};
  HEADERS.forEach((h, i) => { COL[h] = i + 1; });

  const reset = dataRows => {
    SHEET._setGrid([HEADERS.slice()].concat(dataRows || []));
    PROPS._clear();
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

  // ---- weight orphan reconciliation --------------------------------------

  t('reconcileWeightOrphans_ deletes our untracked weight datapoint, spares tracked/foreign', () => {
    const tracked = 'users/me/dataTypes/weight/dataPoints/W-tracked';
    const orphan = 'users/me/dataTypes/weight/dataPoints/W-orphan';
    const device = 'users/me/dataTypes/weight/dataPoints/W-device';
    const rows = [{ healthIds: [tracked] }];
    const deleted = [];
    withStubs({
      listWeightOnDate: () => ([
        { name: tracked, googleWebClientId: 'ours' },
        { name: orphan, googleWebClientId: 'ours' },
        { name: device, googleWebClientId: null }
      ]),
      deleteDataPointsByName: names => { deleted.push.apply(deleted, names); }
    }, () => {
      reconcileWeightOrphans_(rows, Date.UTC(2026, 0, 15, 12, 0, 0), 1);
    });
    eq(deleted, [orphan], 'only our untracked weight datapoint is deleted');
  });

  // ---- backstop: re-dirties BOTH matched and unmatched recent rows --------

  // backstop() uses the real Date.now() (not injectable) for its lookback
  // window, so rows must be dated "now". A Date object in the Date cell is read
  // back verbatim by toDate_, sidestepping the UTC-midnight civil-date boundary
  // that a 'yyyy-MM-dd' string would hit in a non-UTC script time zone.
  t('backstop re-dirties BOTH matched and unmatched recent rows and sets the pending flag', () => {
    const today = new Date();
    reset([
      [today, '', '135x5x3', 'SYNC', '', '', '', '', '', ''],            // row 2: unmatched
      [today, '', '225x5x3', 'SYNC', '', '', '', '', '', 'foreign/F1']   // row 3: matched
    ]);
    withStubs({ listStrengthOnDate: () => [], listWeightOnDate: () => [] }, () => backstop());
    eq(SHEET.getRange(2, COL['Exercise Synced At']).getValue(), '', 'unmatched row re-dirtied');
    eq(SHEET.getRange(3, COL['Exercise Synced At']).getValue(), '', 'matched row re-dirtied');
    ok(PROPS.getProperty('pendingDirty') !== null, 'pending flag set for the next poll');
  });

  t('backstop skips entirely when the lock is held', () => {
    reset([[new Date(), '', '135x5x3', 'SYNC', '', '', '', '', '', '']]);
    LOCK.held = true;
    withStubs({ listStrengthOnDate: () => [], listWeightOnDate: () => [] }, () => backstop());
    eq(SHEET.getRange(2, COL['Exercise Synced At']).getValue(), 'SYNC', 'row untouched when lock held');
    ok(PROPS.getProperty('pendingDirty') === null, 'no pending flag when skipped');
  });

  const msg = results.join('\n');
  console.log(msg);
}
