// Natural-language phrasing mirroring Fitbit-originated Strength Training
// notes ("Bench press, 190 lbs, 5 sets of 5"), which the Google Health UI
// evidently parses to render the in-app "Workout summary" card. The API
// itself has no structured slot for per-set strength data
// (Exercise.exerciseMetadata only carries hasGps / poolLengthMillimeters),
// so the notes field is the only lever we have.
function buildNotes(durationMs, parsedExercises) {
  const lines = [];
  for (const ex of parsedExercises) {
    for (const entry of ex.entries) {
      // Non-sendable (zero-set "not yet performed") entries produce no note.
      if (!exerciseEntryIsSendable_(entry)) continue;
      lines.push(formatEntryNote_(ex.name, entry));
    }
  }
  let notes = lines.join('\n');
  const minutes = Math.round(durationMs / 60000);
  if (minutes > 0) {
    notes += (notes ? ', ' : '') + minutes + ' minute session';
  }
  return notes;
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
