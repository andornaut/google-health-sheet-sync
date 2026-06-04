// Debug.gs — disposable. Investigation of the Google Health app's session-
// shadowing behavior is complete (see AGENTS.md "Google Health app: card display
// / session-shadowing behavior" and bug-report.md). Only this cleanup remains.
//
// probeCleanupTestData: delete every STRENGTH_TRAINING datapoint created by OUR
// client on the test days below. Scoped to our client id so it never touches
// device/foreign datapoints. NOTE: this also deletes the real datapoint for any
// sheet row on these days — run forceResyncCurrentRow / forceResyncAllRows
// afterward to recreate them per the normal (coincident) sync policy.
//
// Takes no arguments (Apps Script editor "Run"). NOT wired into the Sync menu.
// Delete this file (and `clasp push`) when no probe is needed.

const DEBUG_OUR_CLIENT_ID = '712252384998-1ckpo521s8h473tq573dd4l95cg26ka1.apps.googleusercontent.com';
const DEBUG_TEST_DATES = ['2026-06-02', '2026-06-03', '2026-06-04'];

function probeCleanupTestData() {
  let total = 0;
  DEBUG_TEST_DATES.forEach(s => {
    const a = String(s).split('-').map(Number);
    const date = new Date(a[0], a[1] - 1, a[2], 12, 0, 0);
    const ours = listExercisesOnDate(date)
      .filter(p => p.exercise && p.exercise.exerciseType === 'STRENGTH_TRAINING')
      .filter(p => p.dataSource && p.dataSource.application
        && p.dataSource.application.googleWebClientId === DEBUG_OUR_CLIENT_ID)
      .map(p => p.name);
    if (ours.length) {
      deleteDataPointsByName(ours);
      console.log('Deleted ' + ours.length + ' of our STRENGTH datapoint(s) on ' + s + '.');
    }
    total += ours.length;
  });
  console.log('Done. Deleted ' + total + ' total across ' + DEBUG_TEST_DATES.join(', ')
    + '. Run forceResyncAllRows (or forceResyncCurrentRow on affected rows) to recreate real records.');
}
