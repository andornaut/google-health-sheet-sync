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

  const msg = results.join('\n');
  console.log(msg);
  try { SpreadsheetApp.getUi().alert('Parser tests\n\n' + msg); } catch (e) {}
}
