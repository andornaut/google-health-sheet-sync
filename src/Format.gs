// Natural-language phrasing mirroring Fitbit-originated Strength Training
// notes ("Bench press, 190 lbs, 5 sets of 5"), which the Google Health UI
// evidently parses to render the in-app "Workout summary" card. The API
// itself has no structured slot for per-set strength data
// (Exercise.exerciseMetadata only carries hasGps / poolLengthMillimeters),
// so the notes field is the only lever we have.
//
// Each entry is terminated with a period and the duration is its own sentence
// on its own line. Commas are already the FIELD delimiter within an entry
// ("Bench press, 190 lbs, 5 sets of 5"), so newlines alone carrying the item
// boundary makes the whole blob ambiguous anywhere the pipeline collapses
// whitespace: entry boundaries become indistinguishable from field boundaries.
// A period survives that collapse and outranks the commas. It also keeps the
// duration from reading as a fourth attribute of the last exercise, which is
// how a glued ", 30 minute session" suffix parses.
function buildNotes(durationMs, parsedExercises) {
  const lines = [];
  for (const ex of parsedExercises) {
    for (const entry of ex.entries) {
      // Non-sendable (zero-set "not yet performed") entries produce no note.
      if (!exerciseEntryIsSendable_(entry)) continue;
      lines.push(formatEntryNote_(ex.name, entry) + '.');
    }
  }
  const minutes = Math.round(durationMs / 60000);
  if (minutes > 0) {
    lines.push(minutes + ' minute session.');
  }
  return lines.join('\n');
}

// A parsed entry is "sendable" when it represents a performed set rather than a
// zero-set "not yet performed" start marker (e.g. "200x5x0"). This is the one
// place the zero-set rule is defined; buildNotes and hasSendableExercises_ both
// consult it.
function exerciseEntryIsSendable_(entry) {
  return entry.sets !== 0;
}

// True when the parsed exercises contain at least one sendable entry. A row
// whose only entries are zero-set has timing but nothing to send to Health, so
// callers gate datapoint creation and foreign-match windows on this rather than
// on a bare entry count.
function hasSendableExercises_(parsedExercises) {
  for (const ex of parsedExercises) {
    for (const entry of ex.entries) {
      if (exerciseEntryIsSendable_(entry)) return true;
    }
  }
  return false;
}

function formatEntryNote_(exerciseName, entry) {
  let line = exerciseName + ', ' + entry.weight + ' lbs';
  if (entry.sets !== null && entry.reps !== null) {
    const setLabel = entry.sets === 1 ? '1 set' : entry.sets + ' sets';
    line += ', ' + setLabel + ' of ' + entry.reps;
  } else if (entry.reps !== null) {
    line += ', ' + entry.reps + (entry.reps === 1 ? ' rep' : ' reps');
  }
  return entry.assisted ? line + ' (assisted)' : line;
}
