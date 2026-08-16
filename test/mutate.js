// Mutation check for the test suite itself: `npm run mutate`.
//
// Every entry below undoes one decision the code makes on purpose. The suite is
// expected to FAIL for each. A mutation that survives means the decision is no
// longer pinned by any test, so a future change could silently undo it and the
// green suite would say nothing. That is the failure mode this file exists to
// catch: `npm test` proves the code works today, this proves the tests would
// notice if it stopped.
//
// Adding to the catalog: when a review or a bug turns up a decision worth
// keeping, add the mutation that undoes it here, in the same commit as the test
// that catches it. `name` is the decision stated as its undoing.
//
// `find` must match the source EXACTLY and exactly once. A pattern that no
// longer matches is a hard error, not a skip: the code moved and the entry
// needs re-anchoring, which is the moment to check the decision still holds.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const root = path.resolve(__dirname, "..");

const MUTATIONS = [
  // ---- Parser / notes -----------------------------------------------------
  {
    file: "Parser.gs",
    name: "zero-rep entries are kept instead of dropped as not-performed",
    find: "    if (entry.reps === 0) {",
    replace: "    if (false) {",
  },
  {
    file: "Parser.gs",
    name: "bodyweight plausibility floor removed",
    find: "  if (n < MIN_BODYWEIGHT_LB) {",
    replace: "  if (false) {",
  },
  {
    file: "Parser.gs",
    name: "bodyweight plausibility cap removed",
    find: "  if (n > MAX_BODYWEIGHT_LB) {",
    replace: "  if (false) {",
  },
  {
    file: "Format.gs",
    name: "zero-set entries are no longer suppressed from the notes",
    find: "      if (!exerciseEntryIsSendable_(entry)) continue;",
    replace: "      void entry;",
  },
  {
    file: "Format.gs",
    name: "entry lines are no longer period-terminated",
    find: "      lines.push(formatEntryNote_(ex.name, entry) + '.');",
    replace: "      lines.push(formatEntryNote_(ex.name, entry));",
  },
  {
    file: "Format.gs",
    name: "the duration is glued onto the last entry line again",
    find: "    lines.push(minutes + ' minute session.');",
    replace:
      "    if (lines.length) lines[lines.length - 1] += ', ' + minutes + ' minute session';\n" +
      "    else lines.push(minutes + ' minute session');",
  },

  // ---- Health API shaping -------------------------------------------------
  {
    file: "HealthApi.gs",
    name: "a create returning no resource name no longer throws (exercise)",
    find: "    throw new Error('createExerciseAt: create returned no datapoint name: ' + JSON.stringify(resp));",
    replace: "    return null;",
  },
  {
    file: "HealthApi.gs",
    name: "a create returning no resource name no longer throws (weight)",
    find: "    throw new Error('createWeightAt: create returned no datapoint name: ' + JSON.stringify(resp));",
    replace: "    return null;",
  },
  {
    file: "HealthApi.gs",
    name: "listStrengthOnDate no longer filters to STRENGTH_TRAINING",
    find: "    if (exType !== 'STRENGTH_TRAINING') continue;",
    replace: "    void exType;",
  },
  {
    file: "HealthApi.gs",
    name: "patchWeight no longer rewrites the user id to the literal me",
    find: "  const url = HEALTH_API_BASE + '/' + toMeName_(name);\n  const grams = Math.round(lbs * GRAMS_PER_LB);",
    replace:
      "  const url = HEALTH_API_BASE + '/' + name;\n  const grams = Math.round(lbs * GRAMS_PER_LB);",
  },

  // ---- Timing resolution --------------------------------------------------
  {
    file: "Main.gs",
    name: "a foreign match no longer wins over edit/prior timing",
    find: "  if (foreignInterval) {",
    replace: "  if (false && foreignInterval) {",
  },
  {
    file: "Main.gs",
    name: "exerciseUnchanged_ stops comparing startTime",
    find: "  if (new Date(i.startTime).getTime() !== targetStartUtcMs) return false;\n",
    replace: "",
  },
  {
    file: "Main.gs",
    name: "exerciseUnchanged_ stops comparing notes",
    find: "  return (ex.notes || '') === (targetNotes || '');",
    replace: "  return true;",
  },
  {
    file: "Main.gs",
    name: "the edit-span test no longer protects a recorded interval",
    find: "  return exerciseEditSpansWorkout_(row) || !priorExercise;",
    replace: "  return true;",
  },
  {
    file: "Main.gs",
    name: "edit-derived duration loses its MIN floor / start-only default",
    find: "  return Math.min(Math.max(rawDurationMs, MIN_EXERCISE_DURATION_MS), MAX_EXERCISE_DURATION_MS);",
    replace: "  return Math.min(rawDurationMs, MAX_EXERCISE_DURATION_MS);",
  },

  // ---- Foreign matching ---------------------------------------------------
  {
    file: "Main.gs",
    name: "two rows may borrow the same foreign session",
    find: "      candidates.splice(bestIdx, 1);",
    replace: "      void 0;",
  },
  {
    file: "Main.gs",
    name: "sync-created datapoints are no longer excluded as foreign candidates",
    find: "  splitHealthIdsByType_(allHealthIds).exercise.forEach(name => { excluded[name] = true; });",
    replace: "  void allHealthIds;",
  },
  {
    file: "Main.gs",
    name: "a session aligned to a non-ready row is no longer excluded",
    find: "    if (!readyRowNums[m.rowNum]) excluded[m.name] = true;",
    replace: "    void m;",
  },
  {
    file: "Main.gs",
    name: "a zero-set-only row may anchor a foreign-match window",
    find: "    .filter(r => hasSendableExercises_(r.exercises) && exerciseEditIsOnRowDate_(r))",
    replace: "    .filter(r => exerciseEditIsOnRowDate_(r))",
  },
  {
    file: "Main.gs",
    name: "the foreign-match window is no longer clamped to MAX_EXERCISE_DURATION_MS",
    find: "      const clampedEndMs = startMs + capExerciseDurationToMax_(r.exercisesLastEditedAt.getTime() - startMs);",
    replace: "      const clampedEndMs = r.exercisesLastEditedAt.getTime();",
  },

  // ---- Sync pass lifecycle ------------------------------------------------
  {
    file: "Main.gs",
    name: "the row cap no longer bounds the pass",
    find: "      ready.length = MAX_ROWS_PER_SYNC;",
    replace: "      void 0;",
  },
  {
    file: "Main.gs",
    name: "rows are no longer processed newest-first",
    find: "      const dateDiff = b.row.date.getTime() - a.row.date.getTime();",
    replace:
      "      const dateDiff = a.row.date.getTime() - b.row.date.getTime();",
  },
  {
    file: "Main.gs",
    name: "a per-row failure aborts the pass instead of being isolated",
    find: "      } catch (err) {\n        errors++;\n        unexpected.push('row ' + entry.row.rowNum + ': ' + err);",
    replace:
      "      } catch (err) {\n        throw err;\n        // eslint-disable-next-line no-unreachable\n        unexpected.push('row ' + entry.row.rowNum + ': ' + err);",
  },
  {
    file: "Main.gs",
    name: "the summary throw no longer sets the dirty flag",
    find: "      markPendingDirty_();\n      // Carry the pass outcome in the message.",
    replace: "      // Carry the pass outcome in the message.",
  },
  {
    file: "Main.gs",
    name: "the summary throw drops the deferred count",
    find: "        + (deferredCount > 0 ? ', ' + deferredCount + ' deferred by the row cap' : '')",
    replace: "        + ''",
  },
  {
    file: "Main.gs",
    name: "a concurrent edit during a pass no longer preserves the dirty flag",
    find: "      } else if (!concurrentEdit) {\n        props.deleteProperty(PENDING_DIRTY_KEY);",
    replace: "      } else {\n        props.deleteProperty(PENDING_DIRTY_KEY);",
  },

  // ---- Per-row sync -------------------------------------------------------
  {
    file: "Main.gs",
    name: "the exercise concurrent-edit guard is gone",
    find: "  if (exerciseReady && cols.exercisesLastEditedAtCol) {",
    replace: "  if (false && exerciseReady && cols.exercisesLastEditedAtCol) {",
  },
  {
    file: "Main.gs",
    name: "the weight concurrent-edit guard is gone",
    find: "  if (weightReady && cols.weightEditedAtCol) {",
    replace: "  if (false && weightReady && cols.weightEditedAtCol) {",
  },
  {
    file: "Main.gs",
    name: "a partially-synced row no longer advances the dirty generation",
    find: "    markPendingDirty_();\n    console.info(tag + ': partial progress",
    replace: "    console.info(tag + ': partial progress",
  },
  {
    file: "Main.gs",
    name: "a 404 on the prior weight GET no longer drops the stale id",
    find: "        split.weight = [];",
    replace: "        void 0;",
  },
  {
    file: "Main.gs",
    name: "the idempotency skip no longer requires exactly one prior id",
    find: "wantCreate && split.exercise.length === 1 && priorExercise",
    replace: "wantCreate && split.exercise.length >= 1 && priorExercise",
  },
  {
    file: "Main.gs",
    name: "a 404 on a prior datapoint delete is retried forever instead of dropped",
    find: "      if (isNotFoundError_(err)) {\n        console.warn(tag + ': previous ' + label",
    replace:
      "      if (false) {\n        console.warn(tag + ': previous ' + label",
  },
  {
    file: "Main.gs",
    name: "prior datapoints are deleted in one batch instead of one name per call",
    find: "  const remaining = [];\n  names.forEach(name => {\n    try {\n      deleteDataPointsByName([name]);",
    replace:
      "  const remaining = [];\n  names.forEach(name => {\n    try {\n      deleteDataPointsByName(names);",
  },
  {
    file: "Main.gs",
    name: "the matched foreign session is no longer recorded on the row",
    find: "    writeMatchedHealthSession(row.rowNum, cols.matchedHealthSessionCol, foreignMatch ? foreignMatch.name : '');",
    replace: "    void foreignMatch;",
  },

  // ---- Cleared-content reconciliation (the backstop path) -----------------
  {
    file: "Main.gs",
    name: "cleared rows are no longer reconciled at all",
    find: "    let stale = selectStaleDataPointRows_(rows);",
    replace: "    let stale = { exerciseRowNums: [], weightRowNums: [] };",
  },
  {
    file: "Main.gs",
    name: "stale weight rows are selected but never re-dirtied",
    find: "      reDirtyRows_(weightTargets, { weightCol: weightSyncedAtCol });",
    replace: "      void weightTargets;",
  },
  {
    file: "Main.gs",
    name: "an unparseable exercise cell counts as cleared (mass-delete on a header change)",
    find: "      && !hasSendableExercises_(r.exercises) && !r.hasExerciseText) {",
    replace: "      && !hasSendableExercises_(r.exercises)) {",
  },
  {
    file: "Main.gs",
    name: "an unparseable bodyweight counts as cleared (mass-delete on a reformat)",
    find: "      && r.bodyweight === null && !r.hasWeightText) {",
    replace: "      && r.bodyweight === null) {",
  },
  {
    file: "Main.gs",
    name: "reconciliation no longer requires a tracked datapoint to exist",
    find: "    if (r.exerciseSyncedAt && split.exercise.length > 0\n      && !hasSendableExercises_(r.exercises) && !r.hasExerciseText) {",
    replace:
      "    if (r.exerciseSyncedAt\n      && !hasSendableExercises_(r.exercises) && !r.hasExerciseText) {",
  },
  {
    file: "Main.gs",
    name: "already-dirty rows are re-dirtied again",
    find: "    if (r.exerciseSyncedAt && split.exercise.length > 0\n      && !hasSendableExercises_(r.exercises) && !r.hasExerciseText) {",
    replace:
      "    if (split.exercise.length > 0\n      && !hasSendableExercises_(r.exercises) && !r.hasExerciseText) {",
  },
  {
    file: "Main.gs",
    name: "the mass-deletion bound is removed",
    find: "    if (staleCount > STALE_RECONCILE_MAX_ROWS) {",
    replace: "    if (false) {",
  },
  {
    file: "Main.gs",
    name: "the mass-deletion bound is off by one",
    find: "    if (staleCount > STALE_RECONCILE_MAX_ROWS) {",
    replace: "    if (staleCount >= STALE_RECONCILE_MAX_ROWS) {",
  },
  {
    file: "Sheet.gs",
    name: "hasExerciseText ignores blank-header columns (header change looks like a clear)",
    find: "    if (managedColNums.indexOf(c) !== -1) continue;\n    textCols.push(c);",
    replace:
      "    if (managedColNums.indexOf(c) !== -1) continue;\n    if (!String(headers[c - 1] || '').trim()) continue;\n    textCols.push(c);",
  },
  {
    file: "Sheet.gs",
    name: "hasWeightText always reports empty",
    find: "      hasWeightText: hasText(row[weightCol - 1])",
    replace: "      hasWeightText: false",
  },

  // ---- Backstop selection -------------------------------------------------
  {
    file: "Main.gs",
    name: "the backstop re-reviews rows with no sendable content",
    find: "    && hasSendableExercises_(r.exercises));",
    replace: "  );",
  },

  // ---- readRows contracts -------------------------------------------------
  {
    file: "Sheet.gs",
    name: "allHealthIds covers only rows with a parseable Date",
    find: "    healthIds.forEach(n => allHealthIds.push(n));",
    replace: "    void healthIds;",
  },
  {
    file: "Sheet.gs",
    name: "allMatchedSessions covers only rows with a parseable Date",
    find: "      if (matched) allMatchedSessions.push({ name: matched, rowNum: rowNum });",
    replace: "      void matched;",
  },
  {
    file: "Sheet.gs",
    name: "duplicate exercise column headers are accepted silently",
    find: "  if (duplicates.length > 0) {",
    replace: "  if (false) {",
  },

  // ---- onEdit dirty marking -----------------------------------------------
  {
    file: "Sheet.gs",
    name: "a single-cell clear of real content is no longer an edit",
    find: "  const clearedContent = singleCell\n    && isEmptyValue(newValues[0] && newValues[0][0])\n    && !isEmptyValue(e.oldValue);",
    replace: "  const clearedContent = false;",
  },
  {
    file: "Sheet.gs",
    name: "clearing an already-blank cell now counts as an edit",
    find: "  const clearedContent = singleCell\n    && isEmptyValue(newValues[0] && newValues[0][0])\n    && !isEmptyValue(e.oldValue);",
    replace: "  const clearedContent = singleCell;",
  },
  {
    file: "Sheet.gs",
    name: "a weight edit also advances/seeds the exercise timestamps",
    find: "  if (marks.exerciseRows.size > 0) {",
    replace:
      "  marks.weightRows.forEach(r => marks.exerciseRows.add(r));\n  if (marks.exerciseRows.size > 0) {",
  },
  {
    file: "Sheet.gs",
    name: "an exercise edit also advances the weight timestamps",
    find: "  if (marks.weightRows.size > 0) {",
    replace:
      "  marks.exerciseRows.forEach(r => marks.weightRows.add(r));\n  if (marks.weightRows.size > 0) {",
  },
  {
    file: "Sheet.gs",
    name: "an exercise edit marks every row in the range, not just the ones with content",
    find: "      if (headerName) exerciseRows.add(rowNum);",
    replace:
      "      if (headerName) { for (let k = 0; k < numRows; k++) exerciseRows.add(firstRow + k); }",
  },
  {
    file: "Sheet.gs",
    name: "a weight edit marks every row in the range, not just the ones with content",
    find: "      if (c === weightCol) { weightRows.add(rowNum); continue; }",
    replace:
      "      if (c === weightCol) { for (let k = 0; k < numRows; k++) weightRows.add(firstRow + k); continue; }",
  },
  {
    file: "Sheet.gs",
    name: "stampRows_ overwrites rows outside the marked set",
    find: "  rows.forEach(r => {\n    const i = r - block.first;\n    if (block.values[i][0] === value) return;",
    replace:
      "  block.values.forEach((_, i) => {\n    if (block.values[i][0] === value) return;",
  },
  {
    file: "Sheet.gs",
    name: "seedRows_ overwrites a non-blank Exercise First Edited At",
    find: "    if (current !== '' && current !== null && current !== undefined) return;",
    replace: "    void current;",
  },
  {
    file: "Sheet.gs",
    name: "a blank-header scratch column counts as an exercise column",
    find: "      if (headerName) exerciseRows.add(rowNum);",
    replace: "      exerciseRows.add(rowNum);",
  },
  {
    file: "Sheet.gs",
    name: "edits on another sheet are no longer ignored",
    find: "  if (sheet.getSheetId() !== getSheet_().getSheetId()) return false;",
    replace: "  void sheet;",
  },

  // ---- Date validation ----------------------------------------------------
  {
    file: "Main.gs",
    name: "duplicate row dates are accepted",
    find: "      if (key === prev.key) {",
    replace: "      if (false) {",
  },
  {
    file: "Main.gs",
    name: "out-of-order row dates are accepted",
    find: "      if (key < prev.key) {",
    replace: "      if (false) {",
  },
  {
    file: "Main.gs",
    name: "implausible row years are accepted",
    find: "    if (year < MIN_ROW_DATE_YEAR || year > MAX_ROW_DATE_YEAR) {",
    replace: "    if (false) {",
  },
  {
    file: "Main.gs",
    name: "syncOnEdit swallows a date violation instead of alarming",
    find: "  if (violation) throw new Error('syncOnEdit: date validation failed: ' + violation);",
    replace: "  if (violation) return;",
  },
  {
    file: "Main.gs",
    name: "a time-based trigger throws on a date violation instead of skipping",
    find: "  console.error(triggerName + ': date validation failed; skipping: ' + violation);\n  return true;",
    replace: "  throw new Error(triggerName + ': ' + violation);",
  },

  // ---- Orphan reconciliation ----------------------------------------------
  {
    file: "Main.gs",
    name: "orphan deletion no longer requires attributable ownership",
    find: "    if (c.googleWebClientId && ourClientIds[c.googleWebClientId]) {",
    replace: "    if (true) {",
  },
  {
    file: "Main.gs",
    name: "orphan reconciliation deletes tracked datapoints too",
    find: "    if (knownNames[c.name]) return;",
    replace: "    void knownNames;",
  },

  // ---- Manual entry points ------------------------------------------------
  {
    file: "Main.gs",
    name: "a selection on another tab is accepted by resyncSelectedRows",
    find: "  if (activeSheet.getSheetId() !== sheet.getSheetId()) {",
    replace: "  if (false) {",
  },
  {
    file: "Main.gs",
    name: "a whole-column selection is no longer clamped to the data range",
    find: "    const end = Math.min(start + range.getNumRows() - 1, lastDataRow);",
    replace: "    const end = start + range.getNumRows() - 1;",
  },
  {
    file: "Main.gs",
    name: "the backstop runs without holding the script lock",
    find: "  if (!lock.tryLock(LOCK_WAIT_MS)) {\n    console.warn('backstop: another run holds the lock; skipping this run.');\n    return;\n  }",
    replace: "  lock.tryLock(LOCK_WAIT_MS);",
  },
];

// Run the suite once; true when it reported at least one failure.
function suiteFails() {
  try {
    const out = execFileSync("node", [path.join(root, "test", "run.js")], {
      cwd: root,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
    });
    return /^FAIL /m.test(out);
  } catch {
    // Non-zero exit means the suite failed.
    return true;
  }
}

function run() {
  // Without this the whole report is a lie: an already-red suite "catches"
  // every mutation trivially and reports 100% while pinning nothing.
  if (suiteFails()) {
    console.error(
      "Baseline suite is already failing. Fix `npm test` first; until it is green" +
        " every mutation appears caught and this check proves nothing.",
    );
    process.exit(1);
  }

  const originals = {};
  const read = (f) => {
    if (!(f in originals)) {
      originals[f] = fs.readFileSync(path.join(root, "src", f), "utf8");
    }
    return originals[f];
  };
  const write = (f, text) => fs.writeFileSync(path.join(root, "src", f), text);

  // Anchor every pattern BEFORE mutating anything, so a stale catalog reports
  // all of its bad entries at once instead of one per run.
  const unanchored = [];
  MUTATIONS.forEach((m) => {
    const src = read(m.file);
    const count = src.split(m.find).length - 1;
    if (count !== 1) {
      unanchored.push(`${m.file}: ${count} match(es) for "${m.name}"`);
    }
  });
  if (unanchored.length > 0) {
    console.error(
      `Catalog is stale, these patterns no longer anchor to exactly one place:\n  ${unanchored.join(
        "\n  ",
      )}\n\nRe-anchor each entry, and while doing so confirm the decision it pins still holds.`,
    );
    process.exit(1);
  }

  const survivors = [];
  try {
    MUTATIONS.forEach((m, i) => {
      write(m.file, read(m.file).replace(m.find, m.replace));
      const caught = suiteFails();
      write(m.file, read(m.file));
      const label = `[${i + 1}/${MUTATIONS.length}] ${m.file}: ${m.name}`;
      console.log(`${caught ? "caught   " : "SURVIVED "} ${label}`);
      if (!caught) {
        survivors.push(label);
      }
    });
  } finally {
    Object.keys(originals).forEach((f) => write(f, originals[f]));
  }

  console.log(
    `\n${MUTATIONS.length} mutations: ${MUTATIONS.length - survivors.length} caught, ` +
      `${survivors.length} survived`,
  );
  if (survivors.length > 0) {
    console.error(
      `\nUnpinned decisions (no test fails when they are undone):\n  ${survivors.join("\n  ")}`,
    );
    process.exit(1);
  }
}

run();
