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
  t('single weight "135"', () => eq(parseExerciseCell('135'),
    [{ weight: 135, reps: 1, sets: 1, assisted: false }]));
  t('weight x reps "135x5"', () => eq(parseExerciseCell('135x5'),
    [{ weight: 135, reps: 5, sets: 1, assisted: false }]));
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
    { weight: 155, reps: 1, sets: 1, assisted: true }
  ]));
  t('whitespace tolerance "  135 x 5 x 3  "', () => eq(parseExerciseCell('  135 x 5 x 3  '),
    [{ weight: 135, reps: 5, sets: 3, assisted: false }]));
  t('uppercase X "135X5X3"', () => eq(parseExerciseCell('135X5X3'),
    [{ weight: 135, reps: 5, sets: 3, assisted: false }]));
  t('junk line skipped', () => eq(parseExerciseCell('garbage\n135x5'),
    [{ weight: 135, reps: 5, sets: 1, assisted: false }]));
  t('zero reps rejected', () => eq(parseExerciseCell('135x0'), []));
  t('negative weight rejected', () => eq(parseExerciseCell('-5'), []));
  t('decimal weight allowed "22.5"', () => eq(parseExerciseCell('22.5'),
    [{ weight: 22.5, reps: 1, sets: 1, assisted: false }]));

  t('bodyweight parse "185"', () => eq(parseBodyweight('185'), 185));
  t('bodyweight parse "185.4"', () => eq(parseBodyweight('185.4'), 185.4));
  t('bodyweight empty -> null', () => eq(parseBodyweight(''), null));
  t('bodyweight junk -> null', () => eq(parseBodyweight('heavy'), null));

  t('formatEntry single weight', () => eq(formatEntry({ weight: 135, reps: 1, sets: 1, assisted: false }), '135'));
  t('formatEntry weight x reps', () => eq(formatEntry({ weight: 135, reps: 5, sets: 1, assisted: false }), '135x5'));
  t('formatEntry weight x reps x sets', () => eq(formatEntry({ weight: 135, reps: 5, sets: 3, assisted: false }), '135x5x3'));
  t('formatEntry assisted preserved', () => eq(formatEntry({ weight: 135, reps: 5, sets: 3, assisted: true }), '*135x5x3'));

  t('buildNotes structure', () => {
    const notes = buildNotes([
      { name: 'Bench press', entries: [{ weight: 135, reps: 5, sets: 3, assisted: false }] },
      { name: 'Squat', entries: [{ weight: 225, reps: 5, sets: 3, assisted: false }] }
    ]);
    eq(notes, SYNC_MARKER + '\nBench press: 135x5x3\nSquat: 225x5x3');
  });

  t('buildDisplayName uses abbreviations when known', () => eq(
    buildDisplayName([
      { name: 'Bench press', entries: [] },
      { name: 'Squat', entries: [] }
    ]),
    'Strength: BP, SQ'
  ));
  t('buildDisplayName falls back to full name when unknown', () => eq(
    buildDisplayName([{ name: 'Custom Exercise', entries: [] }]),
    'Strength: Custom Exercise'
  ));

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

  const msg = results.join('\n');
  console.log(msg);
  try { SpreadsheetApp.getUi().alert('Parser tests\n\n' + msg); } catch (e) {}
}
