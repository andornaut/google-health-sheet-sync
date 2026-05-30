// Natural-language phrasing mirroring Fitbit-originated Strength Training
// notes ("Bench press, 190 lbs, 5 sets of 5"), which the Google Health UI
// evidently parses to render the in-app "Workout summary" card. The API
// itself has no structured slot for per-set strength data
// (Exercise.exerciseMetadata only carries hasGps / poolLengthMillimeters),
// so the notes field is the only lever we have.
function buildNotes(parsedExercises) {
  const lines = [];
  for (const ex of parsedExercises) {
    for (const entry of ex.entries) {
      lines.push(formatEntryNote_(ex.name, entry));
    }
  }
  return lines.join('\n');
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
