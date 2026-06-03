function parseExerciseCell(raw) {
  if (raw === null || raw === undefined) return [];
  const text = String(raw).trim();
  if (text === '') return [];

  const entries = [];
  const parts = text.split(/[\r\n,;]+/);
  for (const part of parts) {
    const trimmed = part.trim();
    if (trimmed === '') continue;
    const entry = parseLine_(trimmed);
    if (entry === null) {
      console.warn('Parser: could not parse "' + trimmed + '"');
      continue;
    }
    // 0 sets or 0 reps means the exercise was not performed: skip the entry
    // (no exercise log), distinct from an unknown/undefined count (null), which
    // is logged.
    if (entry.sets === 0 || entry.reps === 0) {
      console.info('Parser: skipping not-performed entry "' + trimmed + '" (0 sets/reps)');
      continue;
    }
    entries.push(entry);
  }
  return entries;
}

function parseLine_(line) {
  let assisted = false;
  let body = line;
  if (body.startsWith('*')) {
    assisted = true;
    body = body.slice(1).trim();
  }
  const parts = body.split(/\s*x\s*/i);
  if (parts.length < 1 || parts.length > 3) return null;
  const nums = parts.map(p => Number(p));
  if (nums.some(n => !Number.isFinite(n) || n < 0)) return null;

  const weight = nums[0];
  const reps = parts.length >= 2 ? nums[1] : null;
  const sets = parts.length === 3 ? nums[2] : null;
  if (reps !== null && (!Number.isInteger(reps) || reps < 0)) return null;
  if (sets !== null && (!Number.isInteger(sets) || sets < 0)) return null;

  return { weight: weight, reps: reps, sets: sets, assisted: assisted };
}

function parseBodyweight(raw) {
  if (raw === null || raw === undefined) return null;
  const text = String(raw).trim();
  if (text === '') return null;
  const n = Number(text);
  if (!Number.isFinite(n) || n <= 0) return null;
  return n;
}
