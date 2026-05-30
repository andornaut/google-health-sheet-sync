function runParserTests() {
  const results = [];
  const t = (name, fn) => {
    try { fn(); results.push('PASS ' + name); }
    catch (err) { results.push('FAIL ' + name + ': ' + err); }
  };
  const eq = (a, b, msg) => {
    const sa = JSON.stringify(a), sb = JSON.stringify(b);
    if (sa !== sb) throw new Error((msg || 'mismatch') + ' expected ' + sb + ' got ' + sa);
  };

  t('empty cell -> []', () => eq(parseExerciseCell(''), []));
  t('null cell -> []', () => eq(parseExerciseCell(null), []));
  t('single weight "135" (reps/sets unknown)', () => eq(parseExerciseCell('135'),
    [{ weight: 135, reps: null, sets: null, assisted: false }]));
  t('weight x reps "135x5" (sets unknown)', () => eq(parseExerciseCell('135x5'),
    [{ weight: 135, reps: 5, sets: null, assisted: false }]));
  t('weight x reps x sets "135x5x3"', () => eq(parseExerciseCell('135x5x3'),
    [{ weight: 135, reps: 5, sets: 3, assisted: false }]));
  t('assisted prefix "*135x5x3"', () => eq(parseExerciseCell('*135x5x3'),
    [{ weight: 135, reps: 5, sets: 3, assisted: true }]));
  t('multiline cell', () => eq(parseExerciseCell('135x5x3\n145x3x2'), [
    { weight: 135, reps: 5, sets: 3, assisted: false },
    { weight: 145, reps: 3, sets: 2, assisted: false }
  ]));
  t('comma-separated cell "95x5x2, 85x5x5"', () => eq(parseExerciseCell('95x5x2, 85x5x5'), [
    { weight: 95, reps: 5, sets: 2, assisted: false },
    { weight: 85, reps: 5, sets: 5, assisted: false }
  ]));
  t('mixed comma + newline', () => eq(parseExerciseCell('135x5x3, 145x3x2\n*155x1'), [
    { weight: 135, reps: 5, sets: 3, assisted: false },
    { weight: 145, reps: 3, sets: 2, assisted: false },
    { weight: 155, reps: 1, sets: null, assisted: true }
  ]));
  t('whitespace tolerance "  135 x 5 x 3  "', () => eq(parseExerciseCell('  135 x 5 x 3  '),
    [{ weight: 135, reps: 5, sets: 3, assisted: false }]));
  t('uppercase X "135X5X3"', () => eq(parseExerciseCell('135X5X3'),
    [{ weight: 135, reps: 5, sets: 3, assisted: false }]));
  t('junk line skipped', () => eq(parseExerciseCell('garbage\n135x5'),
    [{ weight: 135, reps: 5, sets: null, assisted: false }]));
  t('zero reps rejected', () => eq(parseExerciseCell('135x0'), []));
  t('negative weight rejected', () => eq(parseExerciseCell('-5'), []));
  t('decimal weight allowed "22.5" (reps/sets unknown)', () => eq(parseExerciseCell('22.5'),
    [{ weight: 22.5, reps: null, sets: null, assisted: false }]));
  t('decimal reps rejected "135x5.5"', () => eq(parseExerciseCell('135x5.5'), []));
  t('too many x segments rejected "135x5x3x2"', () => eq(parseExerciseCell('135x5x3x2'), []));
  t('semicolon-separated cell "95x5x2;85x5x5"', () => eq(parseExerciseCell('95x5x2;85x5x5'), [
    { weight: 95, reps: 5, sets: 2, assisted: false },
    { weight: 85, reps: 5, sets: 5, assisted: false }
  ]));

  t('bodyweight parse "185"', () => eq(parseBodyweight('185'), 185));
  t('bodyweight parse "185.4"', () => eq(parseBodyweight('185.4'), 185.4));
  t('bodyweight empty -> null', () => eq(parseBodyweight(''), null));
  t('bodyweight junk -> null', () => eq(parseBodyweight('heavy'), null));
  t('bodyweight zero -> null', () => eq(parseBodyweight('0'), null));

  t('formatEntryNote_ multiple sets', () => eq(
    formatEntryNote_('Bench press', { weight: 190, reps: 5, sets: 5, assisted: false }),
    'Bench press, 190 lbs, 5 sets of 5'
  ));
  t('formatEntryNote_ single set', () => eq(
    formatEntryNote_('Bench press', { weight: 135, reps: 5, sets: 1, assisted: false }),
    'Bench press, 135 lbs, 1 set of 5'
  ));
  t('formatEntryNote_ single rep', () => eq(
    formatEntryNote_('Bench press', { weight: 225, reps: 1, sets: 1, assisted: false }),
    'Bench press, 225 lbs, 1 set of 1'
  ));
  t('formatEntryNote_ assisted suffix', () => eq(
    formatEntryNote_('Pull up', { weight: 25, reps: 5, sets: 3, assisted: true }),
    'Pull up, 25 lbs, 3 sets of 5 (assisted)'
  ));
  t('formatEntryNote_ decimal weight', () => eq(
    formatEntryNote_('Lateral raise', { weight: 22.5, reps: 10, sets: 3, assisted: false }),
    'Lateral raise, 22.5 lbs, 3 sets of 10'
  ));
  t('formatEntryNote_ weight only (reps/sets unknown)', () => eq(
    formatEntryNote_('Bench press', { weight: 135, reps: null, sets: null, assisted: false }),
    'Bench press, 135 lbs'
  ));
  t('formatEntryNote_ weight + reps (sets unknown)', () => eq(
    formatEntryNote_('Bench press', { weight: 135, reps: 5, sets: null, assisted: false }),
    'Bench press, 135 lbs, 5 reps'
  ));
  t('formatEntryNote_ weight + 1 rep (sets unknown)', () => eq(
    formatEntryNote_('Bench press', { weight: 225, reps: 1, sets: null, assisted: false }),
    'Bench press, 225 lbs, 1 rep'
  ));
  t('formatEntryNote_ weight only + assisted', () => eq(
    formatEntryNote_('Pull up', { weight: 25, reps: null, sets: null, assisted: true }),
    'Pull up, 25 lbs (assisted)'
  ));

  t('buildNotes one line per entry', () => {
    const notes = buildNotes([
      { name: 'Bench press', entries: [
        { weight: 135, reps: 5, sets: 3, assisted: false },
        { weight: 145, reps: 3, sets: 2, assisted: false }
      ] },
      { name: 'Squat', entries: [{ weight: 225, reps: 5, sets: 3, assisted: false }] }
    ]);
    eq(notes,
      'Bench press, 135 lbs, 3 sets of 5\n'
      + 'Bench press, 145 lbs, 2 sets of 3\n'
      + 'Squat, 225 lbs, 3 sets of 5');
  });

  t('parseHealthIds_ empty/null', () => {
    eq(parseHealthIds_(''), []);
    eq(parseHealthIds_(null), []);
    eq(parseHealthIds_(undefined), []);
    eq(parseHealthIds_('   '), []);
  });
  t('parseHealthIds_ valid JSON array of strings', () => eq(
    parseHealthIds_('["users/me/dataTypes/exercise/dataPoints/abc","users/me/dataTypes/weight/dataPoints/def"]'),
    ['users/me/dataTypes/exercise/dataPoints/abc', 'users/me/dataTypes/weight/dataPoints/def']
  ));
  t('parseHealthIds_ filters non-string elements', () => eq(
    parseHealthIds_('["a", 1, null, "b", {}]'),
    ['a', 'b']
  ));
  t('parseHealthIds_ non-array JSON -> []', () => eq(parseHealthIds_('{"x":1}'), []));
  t('parseHealthIds_ malformed JSON -> []', () => eq(parseHealthIds_('not json'), []));

  t('extractDataPointName_ null/undefined -> null', () => {
    eq(extractDataPointName_(null), null);
    eq(extractDataPointName_(undefined), null);
    eq(extractDataPointName_({}), null);
  });
  t('extractDataPointName_ LRO response shape', () => eq(
    extractDataPointName_({ done: true, response: { name: 'users/me/dataTypes/exercise/dataPoints/xyz' } }),
    'users/me/dataTypes/exercise/dataPoints/xyz'
  ));
  t('extractDataPointName_ direct name with dataPoints segment', () => eq(
    extractDataPointName_({ name: 'users/me/dataTypes/weight/dataPoints/abc' }),
    'users/me/dataTypes/weight/dataPoints/abc'
  ));
  t('extractDataPointName_ ignores non-datapoint name', () => eq(
    extractDataPointName_({ name: 'operations/123' }),
    null
  ));

  t('parseOffsetSeconds_ empty -> 0', () => {
    eq(parseOffsetSeconds_(''), 0);
    eq(parseOffsetSeconds_(null), 0);
    eq(parseOffsetSeconds_(undefined), 0);
  });
  t('parseOffsetSeconds_ "3600s" -> 3600', () => eq(parseOffsetSeconds_('3600s'), 3600));
  t('parseOffsetSeconds_ "-18000s" -> -18000', () => eq(parseOffsetSeconds_('-18000s'), -18000));
  t('parseOffsetSeconds_ bare number string', () => eq(parseOffsetSeconds_('7200'), 7200));
  t('parseOffsetSeconds_ garbage -> 0', () => eq(parseOffsetSeconds_('abc'), 0));

  t('formatSyncResult_ null -> lock-held message', () => eq(
    formatSyncResult_(null, 'Synced'),
    'Sync skipped (another run holds the lock). Try again shortly.'
  ));
  t('formatSyncResult_ ok only', () => eq(
    formatSyncResult_({ ok: 3, errors: 0 }, 'Synced'),
    'Synced 3 row(s).'
  ));
  t('formatSyncResult_ ok + errors', () => eq(
    formatSyncResult_({ ok: 2, errors: 1 }, 'Resynced'),
    'Resynced 2 row(s), 1 error(s).\n\nSee Executions for details.'
  ));
  t('formatSyncResult_ zero ok with errors', () => eq(
    formatSyncResult_({ ok: 0, errors: 4 }, 'Synced'),
    'Synced 0 row(s), 4 error(s).\n\nSee Executions for details.'
  ));
  t('formatSyncResult_ ok + deferred', () => eq(
    formatSyncResult_({ ok: 75, errors: 0, deferred: 25 }, 'Synced'),
    'Synced 75 row(s), 25 deferred.'
  ));
  t('formatSyncResult_ ok + errors + deferred', () => eq(
    formatSyncResult_({ ok: 70, errors: 5, deferred: 25 }, 'Synced'),
    'Synced 70 row(s), 5 error(s), 25 deferred.\n\nSee Executions for details.'
  ));

  t('toDate_ passes Date through', () => {
    const d = new Date('2026-01-15T12:00:00Z');
    if (toDate_(d) !== d) throw new Error('expected same Date instance');
  });
  t('toDate_ parses ISO string', () => {
    const d = toDate_('2026-01-15T12:00:00Z');
    if (!(d instanceof Date) || d.getTime() !== Date.UTC(2026, 0, 15, 12, 0, 0)) {
      throw new Error('expected parsed Date, got ' + d);
    }
  });
  t('toDate_ invalid -> null', () => {
    eq(toDate_('not a date'), null);
    eq(toDate_(''), null);
  });

  t('splitHealthIdsByType_ empty/null', () => {
    eq(splitHealthIdsByType_(null), { weight: [], exercise: [], other: [] });
    eq(splitHealthIdsByType_([]), { weight: [], exercise: [], other: [] });
  });
  t('splitHealthIdsByType_ buckets weight vs exercise', () => eq(
    splitHealthIdsByType_([
      'users/me/dataTypes/weight/dataPoints/w1',
      'users/me/dataTypes/exercise/dataPoints/e1',
      'users/me/dataTypes/weight/dataPoints/w2'
    ]),
    {
      weight: ['users/me/dataTypes/weight/dataPoints/w1', 'users/me/dataTypes/weight/dataPoints/w2'],
      exercise: ['users/me/dataTypes/exercise/dataPoints/e1'],
      other: []
    }
  ));
  t('splitHealthIdsByType_ unknown type -> other', () => eq(
    splitHealthIdsByType_(['users/me/dataTypes/sleep/dataPoints/s1']),
    { weight: [], exercise: [], other: ['users/me/dataTypes/sleep/dataPoints/s1'] }
  ));
  t('splitHealthIdsByType_ malformed name -> other', () => eq(
    splitHealthIdsByType_(['not-a-resource-name', 'users/me/dataTypes/weight/dataPoints/ok']),
    {
      weight: ['users/me/dataTypes/weight/dataPoints/ok'],
      exercise: [],
      other: ['not-a-resource-name']
    }
  ));

  // DST behavior for America/Toronto in 2026:
  //   Spring forward: Mar 8, 02:00 EST -> 03:00 EDT  (clocks jump forward)
  //   Fall back:      Nov 1, 02:00 EDT -> 01:00 EST  (clocks jump back)
  const TORONTO = 'America/Toronto';
  const EST = -5 * 3600;
  const EDT = -4 * 3600;

  t('localCivilToUtcMs_ winter EST (Jan)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 1, 15, 12, 0);
    eq(r.offsetSeconds, EST);
    eq(r.utcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });
  t('localCivilToUtcMs_ summer EDT (Jul)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 7, 15, 12, 0);
    eq(r.offsetSeconds, EDT);
    eq(r.utcMs, Date.UTC(2026, 6, 15, 16, 0, 0));
  });
  t('localCivilToUtcMs_ day after spring-forward (Mar 9)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 3, 9, 12, 0);
    eq(r.offsetSeconds, EDT);
    eq(r.utcMs, Date.UTC(2026, 2, 9, 16, 0, 0));
  });
  t('localCivilToUtcMs_ day after fall-back (Nov 2)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 11, 2, 12, 0);
    eq(r.offsetSeconds, EST);
    eq(r.utcMs, Date.UTC(2026, 10, 2, 17, 0, 0));
  });
  t('localCivilToUtcMs_ spring-forward day pre-cutover (Mar 8 01:00 = EST)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 3, 8, 1, 0);
    eq(r.offsetSeconds, EST);
    eq(r.utcMs, Date.UTC(2026, 2, 8, 6, 0, 0));
  });
  t('localCivilToUtcMs_ spring-forward day post-cutover (Mar 8 03:00 = EDT)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 3, 8, 3, 0);
    eq(r.offsetSeconds, EDT);
    eq(r.utcMs, Date.UTC(2026, 2, 8, 7, 0, 0));
  });
  t('localCivilToUtcMs_ fall-back day pre-cutover (Nov 1 00:00 = EDT)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 11, 1, 0, 0);
    eq(r.offsetSeconds, EDT);
    eq(r.utcMs, Date.UTC(2026, 10, 1, 4, 0, 0));
  });
  t('localCivilToUtcMs_ fall-back day post-cutover (Nov 1 03:00 = EST)', () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 11, 1, 3, 0);
    eq(r.offsetSeconds, EST);
    eq(r.utcMs, Date.UTC(2026, 10, 1, 8, 0, 0));
  });

  t('getTzOffsetSeconds_ winter EST', () => eq(
    getTzOffsetSeconds_(TORONTO, new Date(Date.UTC(2026, 0, 15, 17, 0, 0))), EST
  ));
  t('getTzOffsetSeconds_ summer EDT', () => eq(
    getTzOffsetSeconds_(TORONTO, new Date(Date.UTC(2026, 6, 15, 16, 0, 0))), EDT
  ));
  t('getTzOffsetSeconds_ GMT zero', () => eq(
    getTzOffsetSeconds_('GMT', new Date(Date.UTC(2026, 6, 15, 16, 0, 0))), 0
  ));

  t('buildIntervalFromUtc_ formats interval', () => eq(
    buildIntervalFromUtc_(Date.UTC(2026, 0, 15, 17, 0, 0), EST, Date.UTC(2026, 0, 15, 18, 0, 0), EST),
    {
      startTime: '2026-01-15T17:00:00Z',
      startUtcOffset: '-18000s',
      endTime: '2026-01-15T18:00:00Z',
      endUtcOffset: '-18000s'
    }
  ));

  // syntheticExerciseInterval_ and resolveRowTiming_ use getTz_() which is
  // cached on first call and resolves to the test runner's default
  // (America/Toronto). Test dates use noon UTC so civilDateParts_ returns
  // the intended calendar day in EST (UTC-5).
  const JAN_15_NOON_UTC = new Date(Date.UTC(2026, 0, 15, 12, 0, 0));

  t('syntheticExerciseInterval_ ordinal 0 -> noon-1pm EST', () => {
    const r = syntheticExerciseInterval_(JAN_15_NOON_UTC, 0);
    eq(r.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.endUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
    eq(r.startOffsetSeconds, EST);
    eq(r.endOffsetSeconds, EST);
  });
  t('syntheticExerciseInterval_ ordinal 1 -> 1pm-2pm EST', () => {
    const r = syntheticExerciseInterval_(JAN_15_NOON_UTC, 1);
    eq(r.startUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
    eq(r.endUtcMs, Date.UTC(2026, 0, 15, 19, 0, 0));
  });
  t('syntheticExerciseInterval_ throws when end spills past midnight', () => {
    let threw = false;
    try { syntheticExerciseInterval_(JAN_15_NOON_UTC, 12); } catch (e) { threw = true; }
    if (!threw) throw new Error('expected throw for ordinal 12 (startHour=24, endHour=25)');
  });

  t('resolveRowTiming_ edit source preserves interval within bounds', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(Date.UTC(2026, 0, 15, 18, 0, 0));
    const r = resolveRowTiming_({ firstEditedAt: first, exercisesEditedAt: last, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.exerciseSource, 'edit');
    eq(r.weightSource, 'edit');
    eq(r.exercise.startUtcMs, first.getTime());
    eq(r.exercise.endUtcMs, last.getTime());
    eq(r.exercise.startOffsetSeconds, EST);
    eq(r.exercise.endOffsetSeconds, EST);
    eq(r.weight, { utcMs: first.getTime(), offsetSeconds: EST });
  });
  t('resolveRowTiming_ edit source clamps too-short duration to MIN (5 min)', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + 60 * 1000);
    const r = resolveRowTiming_({ firstEditedAt: first, exercisesEditedAt: last, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 5 * 60 * 1000);
  });
  t('resolveRowTiming_ edit source clamps too-long duration to MAX (120 min)', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + 5 * 60 * 60 * 1000);
    const r = resolveRowTiming_({ firstEditedAt: first, exercisesEditedAt: last, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 120 * 60 * 1000);
  });
  t('resolveRowTiming_ synthetic source when no edit timestamps', () => {
    const r = resolveRowTiming_({ firstEditedAt: null, exercisesEditedAt: null, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.weightSource, 'synthetic');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.exercise.endUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
  });
  t('resolveRowTiming_ weight uses firstEditedAt as legacy fallback when weightEditedAt is missing; exercise falls back to synthetic', () => {
    // 3pm EST = 20:00 UTC; distinct from synthetic-noon EST = 17:00 UTC so we
    // can verify which path each phase took.
    const first = new Date(Date.UTC(2026, 0, 15, 20, 0, 0));
    const r = resolveRowTiming_({ firstEditedAt: first, exercisesEditedAt: null, weightEditedAt: null, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.weightSource, 'edit');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.exercise.endUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
    eq(r.weight, { utcMs: first.getTime(), offsetSeconds: EST });
  });
  t('resolveRowTiming_ weight uses weightEditedAt when set, takes priority over firstEditedAt', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 20, 0, 0));   // 3pm EST
    const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));   // 5pm EST
    const r = resolveRowTiming_({ firstEditedAt: first, exercisesEditedAt: null, weightEditedAt: wEdit, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.weight, { utcMs: wEdit.getTime(), offsetSeconds: EST });
  });
  t('resolveRowTiming_ weight uses weightEditedAt on weight-only row with no firstEditedAt', () => {
    const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));
    const r = resolveRowTiming_({ firstEditedAt: null, exercisesEditedAt: null, weightEditedAt: wEdit, date: JAN_15_NOON_UTC }, 0, null, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.weightSource, 'edit');
    eq(r.weight, { utcMs: wEdit.getTime(), offsetSeconds: EST });
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));   // exercise synthetic
  });

  // Off-date edits: row.date is JAN-15 but the edit timestamps are on
  // JAN-20 (5 days later). The trust rule kicks in: edit-derived timing
  // is rejected because firstEditedAt's civil date != row.date.
  const JAN_20_3PM_EST = new Date(Date.UTC(2026, 0, 20, 20, 0, 0));
  const JAN_20_4PM_EST = new Date(Date.UTC(2026, 0, 20, 21, 0, 0));

  t('resolveRowTiming_ off-date edit with no prior -> synthetic', () => {
    const r = resolveRowTiming_({
      firstEditedAt: JAN_20_3PM_EST,
      exercisesEditedAt: JAN_20_4PM_EST,
      weightEditedAt: JAN_20_3PM_EST,
      date: JAN_15_NOON_UTC
    }, 0, null, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.weightSource, 'synthetic');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));   // noon EST on row.date
    eq(r.weight.utcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });

  t('resolveRowTiming_ off-date edit with prior exercise -> reuses prior interval', () => {
    const priorStart = Date.UTC(2026, 0, 15, 18, 30, 0);
    const priorEnd = Date.UTC(2026, 0, 15, 19, 15, 0);
    const priorExercise = {
      exercise: {
        interval: {
          startTime: '2026-01-15T18:30:00Z',
          startUtcOffset: '-18000s',
          endTime: '2026-01-15T19:15:00Z',
          endUtcOffset: '-18000s'
        }
      }
    };
    const r = resolveRowTiming_({
      firstEditedAt: JAN_20_3PM_EST,
      exercisesEditedAt: JAN_20_4PM_EST,
      date: JAN_15_NOON_UTC
    }, 0, priorExercise, null);
    eq(r.exerciseSource, 'prior');
    eq(r.exercise.startUtcMs, priorStart);
    eq(r.exercise.endUtcMs, priorEnd);
    eq(r.exercise.startOffsetSeconds, EST);
    eq(r.exercise.endOffsetSeconds, EST);
  });

  t('resolveRowTiming_ same-date edit beats prior exercise (live-workout endTime advance)', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(Date.UTC(2026, 0, 15, 18, 0, 0));
    const priorExercise = {
      exercise: {
        interval: {
          startTime: '2026-01-15T16:00:00Z',
          startUtcOffset: '-18000s',
          endTime: '2026-01-15T16:30:00Z',
          endUtcOffset: '-18000s'
        }
      }
    };
    const r = resolveRowTiming_({
      firstEditedAt: first,
      exercisesEditedAt: last,
      date: JAN_15_NOON_UTC
    }, 0, priorExercise, null);
    eq(r.exerciseSource, 'edit');
    eq(r.exercise.startUtcMs, first.getTime());
    eq(r.exercise.endUtcMs, last.getTime());
  });

  t('resolveRowTiming_ prior weight always wins on re-sync', () => {
    const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));   // 5pm EST, on row.date
    const priorWeight = {
      weight: {
        sampleTime: { physicalTime: '2026-01-15T15:00:00Z', utcOffset: '-18000s' }
      }
    };
    const r = resolveRowTiming_({
      firstEditedAt: null,
      exercisesEditedAt: null,
      weightEditedAt: wEdit,
      date: JAN_15_NOON_UTC
    }, 0, null, priorWeight);
    eq(r.weightSource, 'prior');
    eq(r.weight, { utcMs: Date.UTC(2026, 0, 15, 15, 0, 0), offsetSeconds: EST });
  });

  t('resolveRowTiming_ malformed prior exercise falls through to synthetic', () => {
    const r = resolveRowTiming_({
      firstEditedAt: JAN_20_3PM_EST,
      exercisesEditedAt: JAN_20_4PM_EST,
      date: JAN_15_NOON_UTC
    }, 0, { exercise: {} }, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });

  // resolveForeignMatches_ tests. listStrengthOnDate is stubbed per-test so
  // we control the foreign candidate list without hitting the API.
  const withStubbedList = (stub, fn) => {
    // Apps Script declares top-level functions on the global scope, but the
    // VM sandbox these tests run in treats the function name like a const-
    // binding at the outer scope, so direct reassignment throws. Stash and
    // restore via globalThis instead.
    const orig = globalThis.listStrengthOnDate;
    globalThis.listStrengthOnDate = stub;
    try { fn(); } finally { globalThis.listStrengthOnDate = orig; }
  };

  // 2026-01-15 12:00 UTC = 2026-01-15 07:00 EST, civil date 2026-01-15.
  const FOREIGN_DATE = new Date(Date.UTC(2026, 0, 15, 12, 0, 0));
  const fRow_ = (overrides) => Object.assign({
    rowNum: 10,
    date: FOREIGN_DATE,
    exercises: [{ name: 'Bench', entries: [{ weight: 135, reps: 5, sets: 3, assisted: false }] }],
    healthIds: [],
    matchedHealthSession: '',
    firstEditedAt: null,
    exercisesEditedAt: null
  }, overrides);
  const fCand_ = (name, startUtcMs, endUtcMs) => ({
    name: name, startUtcMs: startUtcMs, endUtcMs: endUtcMs,
    startUtcOffsetSeconds: EST, endUtcOffsetSeconds: EST
  });

  t('resolveForeignMatches_ time-range matches on-date row to overlapping candidate', () => {
    const row = fRow_({
      firstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),   // 5pm EST
      exercisesEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)) // 6pm EST
    });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([row], [row]);
      eq(plan[10] && plan[10].name, 'foreign/A');
    });
  });

  t('resolveForeignMatches_ off-date row skips time-range, falls into ordinal pairing', () => {
    // Row dated 2026-01-15 but edited 2026-01-20. Under the old logic this
    // landed in timeRangeRows and silently failed (zero overlap). New logic
    // sends it to ordinalRows so 1-row-1-candidate cases pair correctly.
    const row = fRow_({
      firstEditedAt: new Date(Date.UTC(2026, 0, 20, 22, 0, 0)),
      exercisesEditedAt: new Date(Date.UTC(2026, 0, 20, 23, 0, 0))
    });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([row], [row]);
      eq(plan[10] && plan[10].name, 'foreign/A');
    });
  });

  t('resolveForeignMatches_ no-timestamp row pairs via ordinal when counts match', () => {
    const row = fRow_({ rowNum: 10 });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([row], [row]);
      eq(plan[10] && plan[10].name, 'foreign/A');
    });
  });

  t('resolveForeignMatches_ ordinal pairs N rows to N candidates sorted asc', () => {
    const rowA = fRow_({ rowNum: 5 });
    const rowB = fRow_({ rowNum: 10 });
    const candEarly = fCand_('foreign/early',
      Date.UTC(2026, 0, 15, 15, 0, 0), Date.UTC(2026, 0, 15, 16, 0, 0));
    const candLate = fCand_('foreign/late',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [candLate, candEarly], () => {   // unsorted input
      const plan = resolveForeignMatches_([rowA, rowB], [rowA, rowB]);
      eq(plan[5] && plan[5].name, 'foreign/early');         // lowest rowNum -> earliest cand
      eq(plan[10] && plan[10].name, 'foreign/late');
    });
  });

  t('resolveForeignMatches_ ordinal counts disagree -> no pairing', () => {
    const rowA = fRow_({ rowNum: 5 });
    const rowB = fRow_({ rowNum: 10 });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([rowA, rowB], [rowA, rowB]);
      eq(plan[5], undefined);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ excludes sync-created candidates (self-owned name)', () => {
    const ownName = 'users/me/dataTypes/exercise/dataPoints/123';
    const row = fRow_({ rowNum: 10, healthIds: [ownName] });
    const cand = fCand_(ownName,
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([row], [row]);
      eq(plan[10], undefined);   // own datapoint not re-matched to self
    });
  });

  t('resolveForeignMatches_ excludes candidates already matched-elsewhere by a non-ready row', () => {
    const readyRow = fRow_({ rowNum: 10 });
    const nonReadyRow = fRow_({ rowNum: 5, matchedHealthSession: 'foreign/A' });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([readyRow, nonReadyRow], [readyRow]);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ time-range picks the best-overlap candidate when several exist', () => {
    // Row window 4:30pm-6:30pm EST (5pm-6pm edit + 30min buffer each side).
    // candA: 7am-8am EST, no overlap. candB: 5pm-6pm EST, full overlap.
    const row = fRow_({
      firstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
      exercisesEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0))
    });
    const candA = fCand_('foreign/early',
      Date.UTC(2026, 0, 15, 12, 0, 0), Date.UTC(2026, 0, 15, 13, 0, 0));
    const candB = fCand_('foreign/match',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [candA, candB], () => {
      const plan = resolveForeignMatches_([row], [row]);
      eq(plan[10] && plan[10].name, 'foreign/match');
    });
  });

  const msg = results.join('\n');
  console.log(msg);
  try { SpreadsheetApp.getUi().alert('Parser tests\n\n' + msg); } catch (e) {}
}
