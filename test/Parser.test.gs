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
  t('zero reps skipped silently "135x0" (not performed)', () => eq(parseExerciseCell('135x0'), []));
  t('zero sets retained "90x6x0" (start-only marker, suppressed in notes)', () => eq(parseExerciseCell('90x6x0'),
    [{ weight: 90, reps: 6, sets: 0, assisted: false }]));
  t('zero-reps dropped, zero-sets retained, valid kept', () => eq(parseExerciseCell('135x5x3, 90x6x0, 95x0'),
    [{ weight: 135, reps: 5, sets: 3, assisted: false }, { weight: 90, reps: 6, sets: 0, assisted: false }]));
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
  t('bodyweight at cap "499" -> 499', () => eq(parseBodyweight('499'), 499));
  t('bodyweight above cap "500" -> null', () => eq(parseBodyweight('500'), null));
  t('bodyweight typo "1850" -> null (implausible)', () => eq(parseBodyweight('1850'), null));
  t('bodyweight at floor "50" -> 50', () => eq(parseBodyweight('50'), 50));
  t('bodyweight below floor "49" -> null', () => eq(parseBodyweight('49'), null));
  t('bodyweight rep-count typo "5" -> null (below floor)', () => eq(parseBodyweight('5'), null));

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

  t('buildNotes one line per entry, session suffix on last line', () => {
    const notes = buildNotes(45 * 60 * 1000, [
      { name: 'Bench press', entries: [
        { weight: 135, reps: 5, sets: 3, assisted: false },
        { weight: 145, reps: 3, sets: 2, assisted: false }
      ] },
      { name: 'Squat', entries: [{ weight: 225, reps: 5, sets: 3, assisted: false }] }
    ]);
    eq(notes,
      'Bench press, 135 lbs, 3 sets of 5\n'
      + 'Bench press, 145 lbs, 2 sets of 3\n'
      + 'Squat, 225 lbs, 3 sets of 5, 45 minute session');
  });

  t('buildNotes suppresses zero-set entries (start-only markers)', () => {
    const notes = buildNotes(10 * 60 * 1000, [
      { name: 'Bench press', entries: [
        { weight: 200, reps: 5, sets: 0, assisted: false },
        { weight: 200, reps: 5, sets: 2, assisted: false }
      ] }
    ]);
    eq(notes, 'Bench press, 200 lbs, 2 sets of 5, 10 minute session');
  });

  t('buildNotes with only zero-set entries -> just the session suffix', () => {
    const notes = buildNotes(10 * 60 * 1000, [
      { name: 'Bench press', entries: [{ weight: 200, reps: 5, sets: 0, assisted: false }] }
    ]);
    eq(notes, '10 minute session');
  });

  t('buildNotes matches foreign single-entry example', () => {
    const notes = buildNotes(2700 * 1000, [
      { name: 'Bench press', entries: [{ weight: 190, reps: 5, sets: 5, assisted: false }] }
    ]);
    eq(notes, 'Bench press, 190 lbs, 5 sets of 5, 45 minute session');
  });

  t('buildNotes rounds duration to nearest minute', () => {
    const notes = buildNotes(29 * 1000, [
      { name: 'Bench press', entries: [{ weight: 190, reps: 5, sets: 5, assisted: false }] }
    ]);
    eq(notes, 'Bench press, 190 lbs, 5 sets of 5');
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

  t('toMeName_ rewrites numeric user id to me', () => eq(
    toMeName_('users/1234567890/dataTypes/weight/dataPoints/abc'),
    'users/me/dataTypes/weight/dataPoints/abc'
  ));
  t('toMeName_ leaves an already-me name unchanged', () => eq(
    toMeName_('users/me/dataTypes/exercise/dataPoints/xyz'),
    'users/me/dataTypes/exercise/dataPoints/xyz'
  ));

  t('isNotFoundError_ true for 404', () => eq(isNotFoundError_({ statusCode: 404 }), true));
  t('isNotFoundError_ false for 500', () => eq(isNotFoundError_({ statusCode: 500 }), false));
  t('isNotFoundError_ false for null', () => eq(isNotFoundError_(null), false));

  t('capExerciseDurationToMax_ leaves sub-max durations untouched', () => eq(
    capExerciseDurationToMax_(30 * 60 * 1000), 30 * 60 * 1000
  ));
  t('capExerciseDurationToMax_ caps at MAX_EXERCISE_DURATION_MS', () => eq(
    capExerciseDurationToMax_(10 * 60 * 60 * 1000), MAX_EXERCISE_DURATION_MS
  ));

  t('humanizeMs_ sub-second -> ms', () => eq(humanizeMs_(500), '500ms'));
  t('humanizeMs_ seconds rounds', () => eq(humanizeMs_(1500), '2s'));
  t('humanizeMs_ exact minute', () => eq(humanizeMs_(60 * 1000), '1m'));
  t('humanizeMs_ minutes + seconds', () => eq(humanizeMs_(90 * 1000), '1m 30s'));
  t('humanizeMs_ negative clamps to 0ms', () => eq(humanizeMs_(-5), '0ms'));

  t('buildSampleTimeFromUtc_ formats physical + civil time', () => eq(
    buildSampleTimeFromUtc_(Date.UTC(2026, 0, 15, 17, 0, 0), -18000),
    {
      physicalTime: '2026-01-15T17:00:00Z',
      utcOffset: '-18000s',
      civilTime: {
        date: { year: 2026, month: 1, day: 15 },
        time: { hours: 12, minutes: 0, seconds: 0 }
      }
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
  t('syntheticExerciseInterval_ clamps to final slot when end would spill past midnight', () => {
    // ordinal 12 would yield startHour=24/endHour=25; instead of throwing, it
    // clamps into the last 1h slot of the day (23:00-24:00 local = 4-5am UTC
    // next day in EST).
    const r = syntheticExerciseInterval_(JAN_15_NOON_UTC, 12);
    eq(r.startUtcMs, Date.UTC(2026, 0, 16, 4, 0, 0));
    eq(r.endUtcMs, Date.UTC(2026, 0, 16, 5, 0, 0));
    eq(r.startOffsetSeconds, EST);
    eq(r.endOffsetSeconds, EST);
  });

  t('resolveRowTiming_ edit source preserves interval within bounds', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(Date.UTC(2026, 0, 15, 18, 0, 0));
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: last, weightEditedAt: first, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'edit');
    eq(r.weightSource, 'edit');
    eq(r.exercise.startUtcMs, first.getTime());
    eq(r.exercise.endUtcMs, last.getTime());
    eq(r.exercise.startOffsetSeconds, EST);
    eq(r.exercise.endOffsetSeconds, EST);
    eq(r.weight, { utcMs: first.getTime(), offsetSeconds: EST });
  });
  t('resolveRowTiming_ edit source clamps too-short duration to MIN (10 min)', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + 60 * 1000);
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: last, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 10 * 60 * 1000);
  });
  t('resolveRowTiming_ edit source accepts a span right at MAX (120 min)', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + MAX_EXERCISE_DURATION_MS);
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: last, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'edit');
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 120 * 60 * 1000);
  });
  // Past the cap the last edit is a later correction, not the end of the
  // workout. Using it would rebuild a short recorded session as a fabricated
  // 2 h one, so the row falls through to 'prior' (or 'synthetic' with no prior).
  t('resolveRowTiming_ span past MAX falls through to prior, keeping the recorded interval', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + 5 * 60 * 60 * 1000);   // corrected hours later
    const prior = { exercise: { interval: {
      startTime: '2026-01-15T17:00:00Z', endTime: '2026-01-15T17:30:00Z',
      startUtcOffset: EST + 's', endUtcOffset: EST + 's'
    } } };
    const r = resolveRowTiming_(
      { exerciseFirstEditedAt: first, exercisesLastEditedAt: last, date: JAN_15_NOON_UTC }, 0, prior);
    eq(r.exerciseSource, 'prior');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 30 * 60 * 1000, 'recorded 30 min kept, not stretched');
  });
  // The span test guards an interval we already recorded, so with no prior it
  // does not apply: the observed on-date start beats synthetic noon, and the
  // MAX clamp bounds the duration. This is the one path where that clamp works.
  t('resolveRowTiming_ span past MAX with no prior still uses the observed start', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + 5 * 60 * 60 * 1000);
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: last, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'edit');
    eq(r.exercise.startUtcMs, first.getTime(), 'keeps the 9am start rather than synthetic noon');
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, MAX_EXERCISE_DURATION_MS, 'clamped to the cap');
  });
  // Known consequence of protecting the recorded interval: timestamps cannot
  // tell "still logging this workout" from "correcting it later", so a sparsely
  // logged session stays at whatever the first sync recorded.
  t('resolveRowTiming_ sparse logging past MAX keeps the start-only default', () => {
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const last = new Date(first.getTime() + 2.5 * 60 * 60 * 1000);
    const startOnly = { exercise: { interval: {
      startTime: '2026-01-15T17:00:00Z', endTime: '2026-01-15T17:10:00Z',
      startUtcOffset: EST + 's', endUtcOffset: EST + 's'
    } } };
    const r = resolveRowTiming_(
      { exerciseFirstEditedAt: first, exercisesLastEditedAt: last, date: JAN_15_NOON_UTC }, 0, startOnly);
    eq(r.exerciseSource, 'prior');
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, MIN_EXERCISE_DURATION_MS);
  });
  // A midnight-crossing workout keeps the 'edit' path: the span stays small
  // even though the last edit lands on the following civil date.
  t('resolveRowTiming_ edit source survives a midnight-crossing workout', () => {
    const first = new Date(Date.UTC(2026, 0, 16, 4, 45, 0));   // 11:45pm EST Jan 15
    const last = new Date(Date.UTC(2026, 0, 16, 5, 15, 0));    // 12:15am EST Jan 16
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: last, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'edit');
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 30 * 60 * 1000);
  });
  t('resolveRowTiming_ synthetic source when no edit timestamps', () => {
    const r = resolveRowTiming_({ exerciseFirstEditedAt: null, exercisesLastEditedAt: null, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.weightSource, 'synthetic');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.exercise.endUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
  });
  t('resolveRowTiming_ weight falls back to synthetic when weightEditedAt is missing', () => {
    // Exercise-only row (e.g. set bodyweight separately or not at all).
    // exerciseFirstEditedAt no longer feeds weight, so without weightEditedAt
    // the weight phase falls through to synthetic noon on row.date.
    const first = new Date(Date.UTC(2026, 0, 15, 20, 0, 0));
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: null, weightEditedAt: null, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.weightSource, 'synthetic');
    eq(r.weight.utcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });
  t('resolveRowTiming_ weight uses weightEditedAt when set', () => {
    const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));   // 5pm EST
    const r = resolveRowTiming_({ exerciseFirstEditedAt: null, exercisesLastEditedAt: null, weightEditedAt: wEdit, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.weight, { utcMs: wEdit.getTime(), offsetSeconds: EST });
  });
  t('resolveRowTiming_ weight uses weightEditedAt on weight-only row with no exerciseFirstEditedAt', () => {
    const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));
    const r = resolveRowTiming_({ exerciseFirstEditedAt: null, exercisesLastEditedAt: null, weightEditedAt: wEdit, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'synthetic');
    eq(r.weightSource, 'edit');
    eq(r.weight, { utcMs: wEdit.getTime(), offsetSeconds: EST });
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));   // exercise synthetic
  });

  // Off-date edits: row.date is JAN-15 but the edit timestamps are on
  // JAN-20 (5 days later). The trust rule kicks in: edit-derived timing
  // is rejected because exerciseFirstEditedAt's civil date != row.date.
  const JAN_20_3PM_EST = new Date(Date.UTC(2026, 0, 20, 20, 0, 0));
  const JAN_20_4PM_EST = new Date(Date.UTC(2026, 0, 20, 21, 0, 0));

  t('resolveRowTiming_ off-date edit with no prior -> synthetic', () => {
    const r = resolveRowTiming_({
      exerciseFirstEditedAt: JAN_20_3PM_EST,
      exercisesLastEditedAt: JAN_20_4PM_EST,
      weightEditedAt: JAN_20_3PM_EST,
      date: JAN_15_NOON_UTC
    }, 0, null);
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
      exerciseFirstEditedAt: JAN_20_3PM_EST,
      exercisesLastEditedAt: JAN_20_4PM_EST,
      date: JAN_15_NOON_UTC
    }, 0, priorExercise);
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
      exerciseFirstEditedAt: first,
      exercisesLastEditedAt: last,
      date: JAN_15_NOON_UTC
    }, 0, priorExercise);
    eq(r.exerciseSource, 'edit');
    eq(r.exercise.startUtcMs, first.getTime());
    eq(r.exercise.endUtcMs, last.getTime());
  });

  t('resolveRowTiming_ malformed prior exercise falls through to synthetic', () => {
    const r = resolveRowTiming_({
      exerciseFirstEditedAt: JAN_20_3PM_EST,
      exercisesLastEditedAt: JAN_20_4PM_EST,
      date: JAN_15_NOON_UTC
    }, 0, { exercise: {} });
    eq(r.exerciseSource, 'synthetic');
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });

  t('resolveRowTiming_ single edit (start-only) -> 10 min default duration', () => {
    // Only one exercise edit: first == last, no observed end. The MIN floor
    // (10 min) doubles as the start-only default.
    const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
    const r = resolveRowTiming_({ exerciseFirstEditedAt: first, exercisesLastEditedAt: first, date: JAN_15_NOON_UTC }, 0, null);
    eq(r.exerciseSource, 'edit');
    eq(r.exercise.startUtcMs, first.getTime());
    eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 10 * 60 * 1000);
  });

  // editDerivedDurationMs_: maps raw (last - first) to recorded duration. The
  // MIN floor (10 min) is also the start-only default (raw <= 0).
  t('editDerivedDurationMs_ zero (single edit / start-only) -> MIN (10 min)', () =>
    eq(editDerivedDurationMs_(0), 10 * 60 * 1000));
  t('editDerivedDurationMs_ negative -> MIN (10 min)', () =>
    eq(editDerivedDurationMs_(-1000), 10 * 60 * 1000));
  t('editDerivedDurationMs_ short span clamps up to MIN (10 min)', () =>
    eq(editDerivedDurationMs_(60 * 1000), 10 * 60 * 1000));
  t('editDerivedDurationMs_ mid span passes through', () =>
    eq(editDerivedDurationMs_(30 * 60 * 1000), 30 * 60 * 1000));
  t('editDerivedDurationMs_ long span clamps to MAX (120 min)', () =>
    eq(editDerivedDurationMs_(5 * 60 * 60 * 1000), 120 * 60 * 1000));

  // hasSendableExercises_: a zero-set-only row has nothing to send.
  t('hasSendableExercises_ false for empty', () => eq(hasSendableExercises_([]), false));
  t('hasSendableExercises_ false when all entries are zero-set', () => eq(
    hasSendableExercises_([{ name: 'Bench', entries: [{ weight: 200, reps: 5, sets: 0, assisted: false }] }]), false));
  t('hasSendableExercises_ true for a real set', () => eq(
    hasSendableExercises_([{ name: 'Bench', entries: [{ weight: 200, reps: 5, sets: 1, assisted: false }] }]), true));
  t('hasSendableExercises_ true for unknown-sets entry (sets null)', () => eq(
    hasSendableExercises_([{ name: 'Bench', entries: [{ weight: 200, reps: 5, sets: null, assisted: false }] }]), true));
  t('hasSendableExercises_ true when mixed zero-set and real', () => eq(
    hasSendableExercises_([{ name: 'Bench', entries: [
      { weight: 200, reps: 5, sets: 0, assisted: false },
      { weight: 200, reps: 5, sets: 2, assisted: false }
    ] }]), true));

  // selectBackstopRows_: recent + sendable + not-yet-matched rows only.
  const bsNow = Date.UTC(2026, 0, 15, 17, 0, 0);   // noon EST Jan 15
  const bsDate = ymd => new Date(Date.UTC(2026, 0, ymd, 17, 0, 0));
  const sendable = [{ name: 'Bench', entries: [{ weight: 200, reps: 5, sets: 2, assisted: false }] }];
  const zeroOnly = [{ name: 'Bench', entries: [{ weight: 200, reps: 5, sets: 0, assisted: false }] }];
  t('selectBackstopRows_ picks recent unmatched sendable rows, drops matched/old/empty', () => {
    const rows = [
      { rowNum: 2, date: bsDate(15), exercises: sendable, matchedHealthSession: '' },        // today, unmatched -> pick
      { rowNum: 3, date: bsDate(14), exercises: sendable, matchedHealthSession: '' },        // yesterday, unmatched -> pick
      { rowNum: 4, date: bsDate(15), exercises: sendable, matchedHealthSession: 'foreign/x' }, // matched -> skip
      { rowNum: 5, date: bsDate(13), exercises: sendable, matchedHealthSession: '' },        // 2 days back (outside lookback=2) -> skip
      { rowNum: 6, date: bsDate(15), exercises: zeroOnly, matchedHealthSession: '' }          // no sendable content -> skip
    ];
    eq(selectBackstopRows_(rows, bsNow, 2).map(r => r.rowNum), [2, 3]);
  });

  t('selectBackstopRows_ wantMatched=true picks recent matched sendable rows, drops unmatched/old/empty', () => {
    const rows = [
      { rowNum: 2, date: bsDate(15), exercises: sendable, matchedHealthSession: 'foreign/a' },  // today, matched -> pick
      { rowNum: 3, date: bsDate(14), exercises: sendable, matchedHealthSession: 'foreign/b' },  // yesterday, matched -> pick
      { rowNum: 4, date: bsDate(15), exercises: sendable, matchedHealthSession: '' },           // unmatched -> skip
      { rowNum: 5, date: bsDate(13), exercises: sendable, matchedHealthSession: 'foreign/c' },  // outside lookback -> skip
      { rowNum: 6, date: bsDate(15), exercises: zeroOnly, matchedHealthSession: 'foreign/d' }   // no sendable content -> skip
    ];
    eq(selectBackstopRows_(rows, bsNow, 2, true).map(r => r.rowNum), [2, 3]);
  });

  // exerciseUnchanged_: skip the recreate only when interval + notes all match.
  const priorEx = (startIso, endIso, notes) => ({
    exercise: { interval: { startTime: startIso, endTime: endIso }, notes: notes }
  });
  const exStart = Date.UTC(2026, 0, 15, 17, 0, 0);
  const exEnd = Date.UTC(2026, 0, 15, 17, 30, 0);
  t('exerciseUnchanged_ true when interval and notes all match', () => eq(
    exerciseUnchanged_(priorEx('2026-01-15T17:00:00Z', '2026-01-15T17:30:00Z', 'Bench: 200x5x2'),
      exStart, exEnd, 'Bench: 200x5x2'), true));
  t('exerciseUnchanged_ false when endTime differs', () => eq(
    exerciseUnchanged_(priorEx('2026-01-15T17:00:00Z', '2026-01-15T17:30:00Z', 'Bench: 200x5x2'),
      exStart, exEnd + 60000, 'Bench: 200x5x2'), false));
  t('exerciseUnchanged_ false when notes differ', () => eq(
    exerciseUnchanged_(priorEx('2026-01-15T17:00:00Z', '2026-01-15T17:30:00Z', 'Bench: 200x5x2'),
      exStart, exEnd, 'Bench: 200x5x3'), false));
  t('exerciseUnchanged_ false when interval missing', () => eq(
    exerciseUnchanged_({ exercise: { notes: 'x' } }, exStart, exEnd, 'x'), false));

  // selectOrphanDataPointNames_: delete untracked datapoints from our own web
  // client, leave tracked / foreign / other-client / unattributable ones alone.
  const oCand = (name, clientId) => ({ name: name, googleWebClientId: clientId || null });
  t('selectOrphanDataPointNames_ deletes our untracked datapoint (client derived from a tracked one)', () => {
    const candidates = [
      oCand('ex/tracked', 'ours'),   // tracked -> establishes "ours"
      oCand('ex/orphan', 'ours')     // untracked, same client -> orphan
    ];
    eq(selectOrphanDataPointNames_(candidates, { 'ex/tracked': true }), ['ex/orphan']);
  });
  t('selectOrphanDataPointNames_ keeps tracked datapoints', () => {
    const candidates = [oCand('ex/tracked', 'ours')];
    eq(selectOrphanDataPointNames_(candidates, { 'ex/tracked': true }), []);
  });
  t('selectOrphanDataPointNames_ keeps foreign datapoints (null client id)', () => {
    const candidates = [
      oCand('ex/tracked', 'ours'),
      oCand('foreign/device', null)   // device/first-party -> never an orphan
    ];
    eq(selectOrphanDataPointNames_(candidates, { 'ex/tracked': true }), []);
  });
  t('selectOrphanDataPointNames_ keeps untracked datapoints from a different web client', () => {
    const candidates = [
      oCand('ex/tracked', 'ours'),
      oCand('other/app', 'theirs')    // another web app, not ours -> keep
    ];
    eq(selectOrphanDataPointNames_(candidates, { 'ex/tracked': true }), []);
  });
  t('selectOrphanDataPointNames_ deletes nothing when ownership cannot be attributed', () => {
    const candidates = [oCand('ex/orphan', 'ours')];   // no tracked candidate to derive "ours"
    eq(selectOrphanDataPointNames_(candidates, {}), []);
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
    exerciseFirstEditedAt: null,
    exercisesLastEditedAt: null
  }, overrides);
  const fCand_ = (name, startUtcMs, endUtcMs) => ({
    name: name, startUtcMs: startUtcMs, endUtcMs: endUtcMs,
    startUtcOffsetSeconds: EST, endUtcOffsetSeconds: EST
  });

  t('resolveForeignMatches_ time-range matches on-date row to overlapping candidate', () => {
    const row = fRow_({
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),   // 5pm EST
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0))    // 6pm EST
    });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([], [], [row]);
      eq(plan[10] && plan[10].name, 'foreign/A');
    });
  });

  t('resolveForeignMatches_ matches across a civil-date boundary (midnight-crossing workout)', () => {
    // Edits 11:45pm EST Jan 15 -> 12:15am EST Jan 16 (exerciseFirstEditedAt is
    // still on row.date Jan 15). The window straddles midnight, so candidates
    // are probed for both Jan 15 and Jan 16. The foreign session was logged
    // just after midnight (12:00-12:30am EST Jan 16) — a different civil date
    // than the row — and must still match on absolute-UTC overlap.
    const row = fRow_({
      rowNum: 10,
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 16, 4, 45, 0)),   // 11:45pm EST Jan 15
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 16, 5, 15, 0))    // 12:15am EST Jan 16
    });
    const cand = fCand_('foreign/after-midnight',
      Date.UTC(2026, 0, 16, 5, 0, 0), Date.UTC(2026, 0, 16, 5, 30, 0));   // 12:00-12:30am EST Jan 16
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([], [], [row]);
      eq(plan[10] && plan[10].name, 'foreign/after-midnight');
    });
  });

  t('resolveForeignMatches_ off-date row gets no match (no ordinal fallback)', () => {
    // Row dated Jan 15 but edited Jan 20: off-date timestamps anchor no
    // trustworthy window. With the ordinal fallback removed the row gets no
    // alignment and falls through to its own synthetic/prior timing.
    const row = fRow_({
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 20, 22, 0, 0)),
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 20, 23, 0, 0))
    });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 20, 22, 0, 0), Date.UTC(2026, 0, 20, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([], [], [row]);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ no-timestamp row gets no match (no ordinal fallback)', () => {
    const row = fRow_({ rowNum: 10 });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([], [], [row]);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ excludes sync-created candidates (own datapoint not realigned)', () => {
    // On-date row whose window overlaps the candidate — but the candidate IS
    // the row's own prior datapoint, so it must be excluded rather than
    // aligned to itself on re-sync.
    const ownName = 'users/me/dataTypes/exercise/dataPoints/123';
    const row = fRow_({
      rowNum: 10,
      healthIds: [ownName],
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0))
    });
    const cand = fCand_(ownName,
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([ownName], [], [row]);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ excludes ids from rows readRows dropped (blank Date)', () => {
    // The candidate belongs to a row whose Date cell is blank, so readRows
    // dropped it and it is absent from allRows, but its id is still in
    // allHealthIds. Without that, the ready row would borrow our own
    // datapoint's interval as if it were a foreign session.
    const undatedRowsName = 'users/me/dataTypes/exercise/dataPoints/456';
    const row = fRow_({
      rowNum: 10,
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0))
    });
    const cand = fCand_(undatedRowsName,
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([undatedRowsName], [], [row]);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ excludes candidates already aligned-elsewhere by a non-ready row', () => {
    // The ready row's window overlaps the candidate (so it would align absent
    // the exclusion), but row 5 already aligned to it. The exclusion is keyed
    // off the full-sheet allMatchedSessions list, so it holds whether row 5 is
    // merely not-ready this pass or was dropped by readRows for a blank Date.
    const readyRow = fRow_({
      rowNum: 10,
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0))
    });
    const cand = fCand_('foreign/A',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([], [{ rowNum: 5, name: 'foreign/A' }], [readyRow]);
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ time-range window is clamped to MAX_EXERCISE_DURATION_MS', () => {
    // Row's exerciseFirstEditedAt is 9am on row.date; exercisesLastEditedAt
    // drifted 5 days forward to 9am the next workout week. Without clamping,
    // the window would balloon to 5 days and incorrectly catch a candidate
    // at 5pm on row.date. With clamping (2h + 30min buffer = 12:30pm
    // cutoff), that candidate is excluded.
    const row = fRow_({
      rowNum: 10,
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 14, 0, 0)),  // 9am EST Jan 15
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 20, 14, 0, 0))   // 9am EST Jan 20
    });
    // Candidate at 5pm-6pm EST Jan 15 — outside the clamped window but
    // inside the unclamped one.
    const cand = fCand_('foreign/late',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [cand], () => {
      const plan = resolveForeignMatches_([], [], [row]);
      // Clamped window (9am + 2h + 30min buffer = 12:30pm cutoff) doesn't
      // reach the 5pm candidate, so no overlap and no alignment. (If clamping
      // were absent, the 5-day window would have caught 'foreign/late'.)
      eq(plan[10], undefined);
    });
  });

  t('resolveForeignMatches_ time-range picks the best-overlap candidate when several exist', () => {
    // Row window 4:30pm-6:30pm EST (5pm-6pm edit + 30min buffer each side).
    // candA: 7am-8am EST, no overlap. candB: 5pm-6pm EST, full overlap.
    const row = fRow_({
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0))
    });
    const candA = fCand_('foreign/early',
      Date.UTC(2026, 0, 15, 12, 0, 0), Date.UTC(2026, 0, 15, 13, 0, 0));
    const candB = fCand_('foreign/match',
      Date.UTC(2026, 0, 15, 22, 0, 0), Date.UTC(2026, 0, 15, 23, 0, 0));
    withStubbedList(() => [candA, candB], () => {
      const plan = resolveForeignMatches_([], [], [row]);
      eq(plan[10] && plan[10].name, 'foreign/match');
    });
  });

  // findRowDateViolation_: trigger-entry date validation (increasing order,
  // no duplicate dates, year within [MIN_ROW_DATE_YEAR, MAX_ROW_DATE_YEAR]).
  // UTC noon keeps the civil date stable in the test time zone.
  const vDate = (y, m, d) => new Date(Date.UTC(y, m - 1, d, 12, 0, 0));
  const vRow = (rowNum, date) => ({ rowNum: rowNum, date: date });

  t('findRowDateViolation_ empty rows -> null', () => eq(findRowDateViolation_([]), null));
  t('findRowDateViolation_ increasing in-range dates -> null', () =>
    eq(findRowDateViolation_([
      vRow(2, vDate(2025, 1, 1)), vRow(3, vDate(2026, 6, 1)), vRow(4, vDate(2049, 12, 31))
    ]), null));
  t('findRowDateViolation_ duplicate date flagged', () => {
    const v = findRowDateViolation_([vRow(2, vDate(2026, 1, 15)), vRow(3, vDate(2026, 1, 15))]);
    eq(/rows 2 and 3 share the date 2026-01-15/.test(v), true, v);
  });
  t('findRowDateViolation_ same civil day, different times is a duplicate', () => {
    const v = findRowDateViolation_([
      vRow(2, new Date(Date.UTC(2026, 0, 15, 13, 0, 0))),
      vRow(3, new Date(Date.UTC(2026, 0, 15, 20, 0, 0)))
    ]);
    eq(/share the date 2026-01-15/.test(v), true, v);
  });
  t('findRowDateViolation_ decreasing date flagged', () => {
    const v = findRowDateViolation_([vRow(2, vDate(2026, 1, 16)), vRow(3, vDate(2026, 1, 15))]);
    eq(/row 3 \(2026-01-15\) is dated before row 2 \(2026-01-16\)/.test(v), true, v);
  });
  t('findRowDateViolation_ year below MIN flagged', () => {
    const v = findRowDateViolation_([vRow(2, vDate(2024, 12, 31))]);
    eq(/row 2: date 2024-12-31 is outside the allowed years 2025-2049/.test(v), true, v);
  });
  t('findRowDateViolation_ year above MAX flagged', () => {
    const v = findRowDateViolation_([vRow(2, vDate(2050, 1, 1))]);
    eq(/outside the allowed years/.test(v), true, v);
  });

  const msg = results.join('\n');
  const passed = results.filter(r => r.startsWith('PASS ')).length;
  const summary = results.length + ' tests: ' + passed + ' passed, ' + (results.length - passed) + ' failed';
  console.log(msg + '\n\n' + summary);
  try { SpreadsheetApp.getUi().alert('Parser / pure-helper tests\n\n' + msg + '\n\n' + summary); } catch (e) {}
}
