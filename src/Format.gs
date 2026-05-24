function formatEntry(entry) {
  const prefix = entry.assisted ? '*' : '';
  if (entry.reps === 1 && entry.sets === 1) {
    return prefix + entry.weight;
  }
  if (entry.sets === 1) {
    return prefix + entry.weight + 'x' + entry.reps;
  }
  return prefix + entry.weight + 'x' + entry.reps + 'x' + entry.sets;
}

function formatExerciseLine(exerciseName, entries) {
  return exerciseName + ': ' + entries.map(formatEntry).join(', ');
}

function buildNotes(parsedExercises) {
  const lines = [SYNC_MARKER];
  for (const ex of parsedExercises) {
    lines.push(formatExerciseLine(ex.name, ex.entries));
  }
  return lines.join('\n');
}

// Server normalizes displayName for STRENGTH_TRAINING to 'Strength training'
// regardless of what we send, so this is mostly cosmetic. Kept short.
function buildDisplayName(parsedExercises) {
  const abbrs = parsedExercises.map(ex => EXERCISE_ABBREVIATIONS[ex.name] || ex.name);
  return 'Strength: ' + abbrs.join(', ');
}
