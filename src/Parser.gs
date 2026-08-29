function parseExerciseCell(raw) {
  if (raw === null || raw === undefined) {
    return [];
  }
  const text = String(raw).trim();
  if (text === "") {
    return [];
  }

  const entries = [];
  const parts = text.split(/[\r\n,;]+/);
  for (const part of parts) {
    const trimmed = part.trim();
    if (trimmed === "") {
      continue;
    }
    const entry = parseLine_(trimmed);
    if (entry === null) {
      console.warn(`Parser: could not parse "${trimmed}"`);
      continue;
    }
    // 0 reps means the exercise was not performed: skip the entry (no exercise
    // log), distinct from an unknown/undefined count (null), which is logged.
    // 0 sets (e.g. "200x5x0") is RETAINED: it marks the start of an exercise
    // that hasn't been completed yet, so the row counts as a real exercise edit
    // (anchoring the start time / foreign-match window). The notes layer
    // suppresses zero-set entries from the Health notes (see hasSendableExercises_
    // and buildNotes), so a row whose only entries are zero-set produces no
    // datapoint but still records its timing.
    if (entry.reps === 0) {
      console.info(
        `Parser: skipping not-performed entry "${trimmed}" (0 reps)`,
      );
      continue;
    }
    entries.push(entry);
  }
  return entries;
}

function parseLine_(line) {
  let assisted = false;
  let body = line;
  if (body.startsWith("*")) {
    assisted = true;
    body = body.slice(1).trim();
  }
  const parts = body.split(/\s*x\s*/i);
  if (parts.length > 3) {
    return null;
  }
  const nums = parts.map((p) => Number(p));
  if (nums.some((n) => !Number.isFinite(n) || n < 0)) {
    return null;
  }

  const weight = nums[0];
  const reps = parts.length >= 2 ? nums[1] : null;
  const sets = parts.length === 3 ? nums[2] : null;
  if (reps !== null && !Number.isInteger(reps)) {
    return null;
  }
  if (sets !== null && !Number.isInteger(sets)) {
    return null;
  }

  return { assisted, reps, sets, weight };
}

function parseBodyweight(raw) {
  if (raw === null || raw === undefined) {
    return null;
  }
  const text = String(raw).trim();
  if (text === "") {
    return null;
  }
  const n = Number(text);
  if (!Number.isFinite(n)) {
    return null;
  }
  // Reject implausible values (likely typos, e.g. 1850 for 185.0). Returning
  // null treats the cell as no-bodyweight rather than syncing a bad value.
  if (n > MAX_BODYWEIGHT_LB) {
    console.warn(
      `Parser: ignoring implausible bodyweight "${text}" (> ${MAX_BODYWEIGHT_LB} lb)`,
    );
    return null;
  }
  // Reject sub-floor values (likely a rep/set count typed into the Weight
  // column, e.g. "5"). Treat as no-bodyweight rather than syncing it.
  if (n < MIN_BODYWEIGHT_LB) {
    console.warn(
      `Parser: ignoring implausible bodyweight "${text}" (< ${MIN_BODYWEIGHT_LB} lb)`,
    );
    return null;
  }
  return n;
}
