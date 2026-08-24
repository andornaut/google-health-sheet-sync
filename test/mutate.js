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
// Reformatting the sources breaks every anchor at once, so CI runs this check
// alongside the suite; without that the catalog rots silently and proves
// nothing while still reporting success on the entries that happen to match.
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const root = path.resolve(__dirname, "..");

const MUTATIONS = [
  // ---- Parser / notes -----------------------------------------------------
  {
    // ANY invalid value rejects the line, not just all of them. A single-number
    // cell cannot tell the two apart, so only a multi-segment entry catches it.
    file: "Parser.gs",
    name: "a line is rejected only when EVERY value is invalid",
    find: "  if (nums.some((n) => !Number.isFinite(n) || n < 0)) {",
    replace: "  if (nums.every((n) => !Number.isFinite(n) || n < 0)) {",
  },
  {
    file: "Format.gs",
    name: "a one-minute session loses its duration line",
    find: "  if (minutes > 0) {",
    replace: "  if (minutes > 1) {",
  },
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
    file: "Parser.gs",
    name: "a fractional set count is accepted",
    find: "  if (sets !== null && !Number.isInteger(sets)) {",
    replace: "  if (sets !== null && false) {",
  },
  {
    file: "Parser.gs",
    name: "a fractional rep count is accepted",
    find: "  if (reps !== null && !Number.isInteger(reps)) {",
    replace: "  if (reps !== null && false) {",
  },
  {
    file: "Parser.gs",
    name: "negative numbers are no longer rejected",
    find: "  if (nums.some((n) => !Number.isFinite(n) || n < 0)) {",
    replace: "  if (nums.some((n) => !Number.isFinite(n))) {",
  },
  {
    // Drops NaN, not just non-positive values: Number("heavy") is NaN, and a
    // leaked NaN reaches the API as weightGrams: null.
    file: "Parser.gs",
    name: "an unparseable bodyweight is no longer rejected as non-finite",
    find: "  if (!Number.isFinite(n) || n <= 0) {",
    replace: "  if (n <= 0) {",
  },
  {
    file: "Format.gs",
    name: "zero-set entries are no longer suppressed from the notes",
    find: "      if (!exerciseEntryIsSendable_(entry)) {\n        continue;\n      }",
    replace: "      void entry;",
  },
  {
    file: "Format.gs",
    name: "entry lines are no longer period-terminated",
    find: "      lines.push(`${formatEntryNote_(ex.name, entry)}.`);",
    replace: "      lines.push(formatEntryNote_(ex.name, entry));",
  },
  {
    file: "Format.gs",
    name: "the duration is glued onto the last entry line again",
    find: "    lines.push(`${minutes} minute session.`);",
    replace:
      "    if (lines.length) {\n" +
      "      lines[lines.length - 1] += `, ${minutes} minute session`;\n" +
      "    } else {\n" +
      "      lines.push(`${minutes} minute session`);\n" +
      "    }",
  },
  {
    file: "Format.gs",
    name: "the session duration is truncated instead of rounded",
    find: "  const minutes = Math.round(durationMs / 60000);",
    replace: "  const minutes = Math.floor(durationMs / 60000);",
  },

  // ---- HTTP transport (httpJson_) -----------------------------------------
  {
    file: "HealthApi.gs",
    name: "the 2xx success range no longer stops below 300",
    find: "    if (code >= 200 && code < 300) {",
    replace: "    if (code >= 200 && code <= 300) {",
  },
  {
    file: "HealthApi.gs",
    name: "a 429 is treated as permanent instead of retried",
    find: "      transient = code === 429 || (code >= 500 && code < 600);",
    replace: "      transient = code >= 500 && code < 600;",
  },
  {
    file: "HealthApi.gs",
    name: "5xx responses are treated as permanent instead of retried",
    find: "      transient = code === 429 || (code >= 500 && code < 600);",
    replace: "      transient = code === 429;",
  },
  {
    file: "HealthApi.gs",
    name: "a network fault is treated as permanent instead of retried",
    find: "      lastErr = err;\n      transient = true;",
    replace: "      lastErr = err;\n      transient = false;",
  },
  {
    file: "HealthApi.gs",
    name: "the retry budget is cut from 4 attempts",
    find: "  const maxAttempts = 4;",
    replace: "  const maxAttempts = 3;",
  },
  {
    file: "HealthApi.gs",
    name: "retry backoff is no longer exponential",
    find: "    const backoffMs = 500 * Math.pow(2, attempt);",
    replace: "    const backoffMs = 500;",
  },
  {
    // isNotFoundError_ reads this, and every 404-recovery path keys off it.
    file: "HealthApi.gs",
    name: "the HTTP status is no longer attached to the thrown error",
    find: "    lastErr.statusCode = code;",
    replace: "    lastErr.statusCode = null;",
  },
  {
    file: "HealthApi.gs",
    name: "HTTP error codes throw instead of being read from the response",
    find: "    muteHttpExceptions: true,",
    replace: "    muteHttpExceptions: false,",
  },
  {
    file: "HealthApi.gs",
    name: "a request body is sent even when there is no payload",
    find: "  if (payload !== undefined) {",
    replace: "  if (true) {",
  },
  {
    // The cache-hit path: without this, every request after the first in an
    // execution goes out unauthenticated.
    file: "HealthApi.gs",
    name: "the cached auth headers come back empty",
    find: "  if (cachedAuthHeaders_) {\n    return cachedAuthHeaders_;\n  }",
    replace: "  if (cachedAuthHeaders_) {\n    return null;\n  }",
  },
  {
    file: "HealthApi.gs",
    name: "requests no longer ask for a JSON response",
    find: '    Accept: "application/json",',
    replace: '    Accept: "text/plain",',
  },
  {
    file: "HealthApi.gs",
    name: "request bodies are no longer sent as JSON",
    find: '    contentType: "application/json",',
    replace: '    contentType: "text/plain",',
  },
  {
    file: "HealthApi.gs",
    name: "requests go out without the OAuth authorization header",
    find: "    Authorization: `Bearer ${getHealthAccessToken_()}`,",
    replace: '    Authorization: "",',
  },
  // ---- Health API shaping -------------------------------------------------
  {
    // Feeds weight orphan attribution: selectOrphanDataPointNames_ only spares a
    // candidate whose googleWebClientId is null, so claiming ours for a foreign
    // datapoint makes the backstop delete data we do not own.
    file: "HealthApi.gs",
    name: "listWeightOnDate claims our client id for every datapoint",
    find:
      "      googleWebClientId: (app && app.googleWebClientId) || null,\n" +
      "      name: p.name,\n" +
      "    });\n" +
      "  }\n" +
      "  return out;\n" +
      "}\n" +
      "\n" +
      "function getTzOffsetSeconds_",
    replace:
      '      googleWebClientId: "ours",\n' +
      "      name: p.name,\n" +
      "    });\n" +
      "  }\n" +
      "  return out;\n" +
      "}\n" +
      "\n" +
      "function getTzOffsetSeconds_",
  },
  {
    // Verified against the live API; a wrong member is a 400, not an empty list.
    file: "HealthApi.gs",
    name: "the weight list filter member is wrong",
    find: '    "weight.sample_time.civil_time",',
    replace: '    "weight.sample_timeZZ.civil_time",',
  },
  {
    // Under-listing is silent: orphan reconciliation misses datapoints and
    // foreign matching misses sessions, with no error either way.
    file: "HealthApi.gs",
    name: "paging stops after the first page",
    find: "    pageToken = json.nextPageToken || null;",
    replace: "    pageToken = null;",
  },
  {
    file: "HealthApi.gs",
    name: "deletes are grouped by the wrong name segment",
    find: "    const dataType = m[1];",
    replace: "    const dataType = m[2];",
  },
  {
    file: "HealthApi.gs",
    name: "an unparseable datapoint name is no longer skipped on delete",
    find:
      "    if (!m) {\n" +
      "      console.warn(\n" +
      '        `deleteDataPointsByName: unparseable name "${n}"; skipping.`,\n' +
      "      );\n" +
      "      return;\n" +
      "    }",
    replace: "",
  },
  {
    file: "HealthApi.gs",
    name: "an empty delete list still issues a request",
    find: "  if (!names || names.length === 0) {",
    replace: "  if (false) {",
  },
  {
    // Every tested zone is a whole hour, so the minutes term needs a half-hour
    // zone (India, Newfoundland) to be exercised at all.
    file: "HealthApi.gs",
    name: "the timezone offset drops its minutes term",
    find: "  return sign * (hours * 3600 + mins * 60);",
    replace: "  return sign * (hours * 3600);",
  },
  {
    file: "HealthApi.gs",
    name: "patchWeight truncates the gram conversion instead of rounding",
    find:
      "  const grams = Math.round(lbs * GRAMS_PER_LB);\n" +
      "  const payload = {\n" +
      "    weight: { sampleTime, weightGrams: grams },",
    replace:
      "  const grams = Math.floor(lbs * GRAMS_PER_LB);\n" +
      "  const payload = {\n" +
      "    weight: { sampleTime, weightGrams: grams },",
  },
  {
    file: "HealthApi.gs",
    name: "a create returning no resource name no longer throws (exercise)",
    find:
      "    throw new Error(\n" +
      "      `createExerciseAt: create returned no datapoint name: ${JSON.stringify(resp)}`,\n" +
      "    );",
    replace: "    return null;",
  },
  {
    file: "HealthApi.gs",
    name: "a create returning no resource name no longer throws (weight)",
    find:
      "    throw new Error(\n" +
      "      `createWeightAt: create returned no datapoint name: ${JSON.stringify(resp)}`,\n" +
      "    );",
    replace: "    return null;",
  },
  {
    file: "Config.gs",
    name: "requests are addressed to the wrong Health API base",
    find: 'const HEALTH_API_BASE = "https://health.googleapis.com/v4";',
    replace: 'const HEALTH_API_BASE = "https://health.googleapis.com/v3";',
  },
  {
    file: "Config.gs",
    name: "the pound-to-gram conversion factor is wrong",
    find: "const GRAMS_PER_LB = 453.59237;",
    replace: "const GRAMS_PER_LB = 453.6;",
  },
  {
    file: "HealthApi.gs",
    name: "listStrengthOnDate no longer filters to STRENGTH_TRAINING",
    find: '    if (exType !== "STRENGTH_TRAINING") {\n      continue;\n    }',
    replace: "    void exType;",
  },
  {
    file: "HealthApi.gs",
    name: "patchWeight no longer rewrites the user id to the literal me",
    find:
      "  const url = `${HEALTH_API_BASE}/${toMeName_(name)}`;\n" +
      "  const grams = Math.round(lbs * GRAMS_PER_LB);",
    replace:
      "  const url = `${HEALTH_API_BASE}/${name}`;\n" +
      "  const grams = Math.round(lbs * GRAMS_PER_LB);",
  },
  {
    file: "HealthApi.gs",
    name: "createExerciseAt sends its own activeDuration again",
    find: '      exerciseType: "STRENGTH_TRAINING",\n      interval: buildIntervalFromUtc_(\n        startUtcMs,',
    replace:
      '      activeDuration: "1800s",\n' +
      '      exerciseType: "STRENGTH_TRAINING",\n' +
      "      interval: buildIntervalFromUtc_(\n        startUtcMs,",
  },
  {
    file: "HealthApi.gs",
    name: "createExerciseAt sends its own displayName again",
    find: '      exerciseType: "STRENGTH_TRAINING",\n      interval: buildIntervalFromUtc_(',
    replace:
      '      displayName: "Strength Training",\n' +
      '      exerciseType: "STRENGTH_TRAINING",\n' +
      "      interval: buildIntervalFromUtc_(",
  },
  // ---- Timing resolution --------------------------------------------------
  {
    file: "Main.gs",
    name: "an edit span exactly at MAX no longer counts as spanning the workout",
    find: "  return spanMs <= MAX_EXERCISE_DURATION_MS;",
    replace: "  return spanMs < MAX_EXERCISE_DURATION_MS;",
  },
  {
    file: "Main.gs",
    name: "a foreign match no longer wins over edit/prior timing",
    find: "  if (foreignInterval) {",
    replace: "  if (false && foreignInterval) {",
  },
  {
    file: "Main.gs",
    name: "exerciseUnchanged_ stops comparing startTime",
    find: "  if (new Date(i.startTime).getTime() !== targetStartUtcMs) {\n    return false;\n  }\n",
    replace: "",
  },
  {
    file: "Main.gs",
    name: "exerciseUnchanged_ stops comparing notes",
    find: '  return (ex.notes || "") === (targetNotes || "");',
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
    find:
      "  return Math.min(\n" +
      "    Math.max(rawDurationMs, MIN_EXERCISE_DURATION_MS),\n" +
      "    MAX_EXERCISE_DURATION_MS,\n" +
      "  );",
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
    find:
      "  splitHealthIdsByType_(allHealthIds).exercise.forEach((name) => {\n" +
      "    excluded[name] = true;\n" +
      "  });",
    replace: "  void allHealthIds;",
  },
  {
    file: "Main.gs",
    name: "a session aligned to a non-ready row is no longer excluded",
    find: "    if (!readyRowNums[m.rowNum]) {\n      excluded[m.name] = true;\n    }",
    replace: "    void m;",
  },
  {
    file: "Main.gs",
    name: "a zero-set-only row may anchor a foreign-match window",
    find: "      (r) => hasSendableExercises_(r.exercises) && exerciseEditIsOnRowDate_(r),",
    replace: "      (r) => exerciseEditIsOnRowDate_(r),",
  },
  {
    file: "Main.gs",
    name: "the foreign-match window is no longer clamped to MAX_EXERCISE_DURATION_MS",
    find:
      "      const clampedEndMs =\n" +
      "        startMs +\n" +
      "        capExerciseDurationToMax_(r.exercisesLastEditedAt.getTime() - startMs);",
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
    find:
      "      } catch (err) {\n" +
      "        errors++;\n" +
      "        unexpected.push(`row ${entry.row.rowNum}: ${err}`);",
    replace:
      "      } catch (err) {\n" +
      "        throw err;\n" +
      "        // eslint-disable-next-line no-unreachable\n" +
      "        unexpected.push(`row ${entry.row.rowNum}: ${err}`);",
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
    find: '          deferredCount > 0 ? `, ${deferredCount} deferred by the row cap` : ""',
    replace: '          ""',
  },
  {
    file: "Main.gs",
    name: "a concurrent edit during a pass no longer preserves the dirty flag",
    find: "      } else if (!concurrentEdit) {\n        props.deleteProperty(PENDING_DIRTY_KEY);",
    replace: "      } else {\n        props.deleteProperty(PENDING_DIRTY_KEY);",
  },

  // ---- Per-row sync -------------------------------------------------------
  {
    // These four gate both the Synced At stamp and the retry. Not setting one
    // stamps the row synced and reports success, so nothing ever retries while
    // Health is missing the datapoint.
    file: "Main.gs",
    name: "a failed weight PATCH still stamps the phase synced",
    find:
      "          console.error(`${tag}: patchWeight failed: ${err}`);\n" +
      "          weightFailed = true;",
    replace: "          console.error(`${tag}: patchWeight failed: ${err}`);",
  },
  {
    file: "Main.gs",
    name: "a partially-failed weight delete still stamps the phase synced",
    find:
      '      newWeightIds = deletePriorDataPoints_(tag, "weight", split.weight);\n' +
      "      if (newWeightIds.length > 0) {\n" +
      "        weightFailed = true;\n" +
      "      }",
    replace:
      '      newWeightIds = deletePriorDataPoints_(tag, "weight", split.weight);',
  },
  {
    file: "Main.gs",
    name: "a partially-failed exercise delete still stamps the phase synced",
    find:
      "        if (newExerciseIds.length > 0) {\n" +
      "          exerciseFailed = true;\n" +
      "        }",
    replace: "        void newExerciseIds;",
  },
  {
    file: "Main.gs",
    name: "a failed exercise create still stamps the phase synced",
    find:
      "          console.error(`${tag}: createExerciseAt failed: ${err}`);\n" +
      "          exerciseFailed = true;",
    replace:
      "          console.error(`${tag}: createExerciseAt failed: ${err}`);",
  },
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
    find: "    markPendingDirty_();\n    console.info(\n      `${tag}: partial progress",
    replace: "    console.info(\n      `${tag}: partial progress",
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
    find: "      split.exercise.length === 1 &&",
    replace: "      split.exercise.length >= 1 &&",
  },
  {
    file: "Main.gs",
    name: "a 404 on a prior datapoint delete is retried forever instead of dropped",
    find: "      if (isNotFoundError_(err)) {\n        console.warn(\n          `${tag}: previous ${label} datapoint",
    replace:
      "      if (false) {\n        console.warn(\n          `${tag}: previous ${label} datapoint",
  },
  {
    file: "Main.gs",
    name: "prior datapoints are deleted in one batch instead of one name per call",
    find:
      "  const remaining = [];\n" +
      "  names.forEach((name) => {\n" +
      "    try {\n" +
      "      deleteDataPointsByName([name]);",
    replace:
      "  const remaining = [];\n" +
      "  names.forEach((name) => {\n" +
      "    try {\n" +
      "      deleteDataPointsByName(names);",
  },
  {
    file: "Main.gs",
    name: "the matched foreign session is no longer recorded on the row",
    find:
      "    writeMatchedHealthSession(\n" +
      "      row.rowNum,\n" +
      "      cols.matchedHealthSessionCol,\n" +
      '      foreignMatch ? foreignMatch.name : "",\n' +
      "    );",
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
    find:
      "      !hasSendableExercises_(r.exercises) &&\n" +
      "      !r.hasExerciseText\n" +
      "    ) {",
    replace: "      !hasSendableExercises_(r.exercises)\n    ) {",
  },
  {
    file: "Main.gs",
    name: "an unparseable bodyweight counts as cleared (mass-delete on a reformat)",
    find:
      "      r.bodyweight === null &&\n" +
      "      !r.hasWeightText\n" +
      "    ) {",
    replace: "      r.bodyweight === null\n    ) {",
  },
  {
    file: "Main.gs",
    name: "reconciliation no longer requires a tracked datapoint to exist",
    find:
      "      r.exerciseSyncedAt &&\n" +
      "      split.exercise.length > 0 &&\n" +
      "      !hasSendableExercises_(r.exercises) &&",
    replace:
      "      r.exerciseSyncedAt &&\n" +
      "      !hasSendableExercises_(r.exercises) &&",
  },
  {
    file: "Main.gs",
    name: "already-dirty rows are re-dirtied again",
    find:
      "      r.exerciseSyncedAt &&\n" +
      "      split.exercise.length > 0 &&\n" +
      "      !hasSendableExercises_(r.exercises) &&",
    replace:
      "      split.exercise.length > 0 &&\n" +
      "      !hasSendableExercises_(r.exercises) &&",
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
    find:
      "    if (managedColNums.indexOf(c) !== -1) {\n" +
      "      continue;\n" +
      "    }\n" +
      "    textCols.push(c);",
    replace:
      "    if (managedColNums.indexOf(c) !== -1) {\n" +
      "      continue;\n" +
      "    }\n" +
      '    if (!String(headers[c - 1] || "").trim()) {\n' +
      "      continue;\n" +
      "    }\n" +
      "    textCols.push(c);",
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
    find:
      "      recentKeys.has(ymd(r.date)) &&\n" +
      "      hasSendableExercises_(r.exercises),",
    replace: "      recentKeys.has(ymd(r.date)),",
  },

  // ---- readRows contracts -------------------------------------------------
  {
    // A header typed with a trailing space would stop resolving, and the sync
    // aborts with "column missing" on a sheet that looks correct.
    file: "Sheet.gs",
    name: "header lookup no longer tolerates surrounding whitespace",
    find: "    map[String(h).trim()] = i + 1;",
    replace: "    map[String(h)] = i + 1;",
  },
  {
    // Decides whether the backstop sees a row as emptied. Without the trim a
    // cell cleared by typing a space keeps its datapoint.
    file: "Sheet.gs",
    name: "a whitespace-only cell no longer counts as empty",
    find: '    v !== null && v !== undefined && String(v).trim() !== "";',
    replace: '    v !== null && v !== undefined && String(v) !== "";',
  },
  {
    // An AND across the exercise columns makes a row with one blank column look
    // emptied, and the backstop then deletes its datapoint.
    file: "Sheet.gs",
    name: "hasExerciseText requires EVERY exercise column to hold text",
    find: "      hasExerciseText: textCols.some((c) => hasText(row[c - 1])),",
    replace:
      "      hasExerciseText: textCols.every((c) => hasText(row[c - 1])),",
  },
  {
    file: "Sheet.gs",
    name: "allHealthIds covers only rows with a parseable Date",
    find: "    healthIds.forEach((n) => allHealthIds.push(n));",
    replace: "    void healthIds;",
  },
  {
    file: "Sheet.gs",
    name: "allMatchedSessions covers only rows with a parseable Date",
    find:
      "      if (matched) {\n" +
      "        allMatchedSessions.push({ name: matched, rowNum });\n" +
      "      }",
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
    // Losing the conjunction lets a single-row range be read as one cell, so a
    // paste that blanks a cell would reach the delete path as a clear.
    file: "Sheet.gs",
    name: "a multi-cell range can be read as a single-cell clear",
    find: "  const singleCell = firstRow === rangeLastRow && firstCol === lastCol;",
    replace:
      "  const singleCell = firstRow === rangeLastRow || firstCol === lastCol;",
  },
  {
    file: "Sheet.gs",
    name: "a single-cell clear of real content is no longer an edit",
    find:
      "  const clearedContent =\n" +
      "    singleCell &&\n" +
      "    isEmptyValue(newValues[0] && newValues[0][0]) &&\n" +
      "    !isEmptyValue(e.oldValue);",
    replace: "  const clearedContent = false;",
  },
  {
    file: "Sheet.gs",
    name: "clearing an already-blank cell now counts as an edit",
    find:
      "  const clearedContent =\n" +
      "    singleCell &&\n" +
      "    isEmptyValue(newValues[0] && newValues[0][0]) &&\n" +
      "    !isEmptyValue(e.oldValue);",
    replace: "  const clearedContent = singleCell;",
  },
  {
    file: "Sheet.gs",
    name: "a weight edit also advances/seeds the exercise timestamps",
    find: "  if (marks.exerciseRows.size > 0) {",
    replace:
      "  marks.weightRows.forEach((r) => marks.exerciseRows.add(r));\n" +
      "  if (marks.exerciseRows.size > 0) {",
  },
  {
    file: "Sheet.gs",
    name: "an exercise edit also advances the weight timestamps",
    find: "  if (marks.weightRows.size > 0) {",
    replace:
      "  marks.exerciseRows.forEach((r) => marks.weightRows.add(r));\n" +
      "  if (marks.weightRows.size > 0) {",
  },
  {
    file: "Sheet.gs",
    name: "an exercise edit marks every row in the range, not just the ones with content",
    find: "      if (headerName) {\n        exerciseRows.add(rowNum);\n      }",
    replace:
      "      if (headerName) {\n" +
      "        for (let k = 0; k < numRows; k++) {\n" +
      "          exerciseRows.add(firstRow + k);\n" +
      "        }\n" +
      "      }",
  },
  {
    file: "Sheet.gs",
    name: "a weight edit marks every row in the range, not just the ones with content",
    find:
      "      if (c === weightCol) {\n" +
      "        weightRows.add(rowNum);\n" +
      "        continue;\n" +
      "      }",
    replace:
      "      if (c === weightCol) {\n" +
      "        for (let k = 0; k < numRows; k++) {\n" +
      "          weightRows.add(firstRow + k);\n" +
      "        }\n" +
      "        continue;\n" +
      "      }",
  },
  {
    file: "Sheet.gs",
    name: "stampRows_ overwrites rows outside the marked set",
    find:
      "  rows.forEach((r) => {\n" +
      "    const i = r - block.first;\n" +
      "    if (block.values[i][0] === value) {\n" +
      "      return;\n" +
      "    }",
    replace:
      "  block.values.forEach((_, i) => {\n" +
      "    if (block.values[i][0] === value) {\n" +
      "      return;\n" +
      "    }",
  },
  {
    file: "Sheet.gs",
    name: "seedRows_ overwrites a non-blank Exercise First Edited At",
    find:
      '    if (current !== "" && current !== null && current !== undefined) {\n' +
      "      return;\n" +
      "    }",
    replace: "    void current;",
  },
  {
    file: "Sheet.gs",
    name: "a blank-header scratch column counts as an exercise column",
    find: "      if (headerName) {\n        exerciseRows.add(rowNum);\n      }",
    replace: "      void headerName;\n      exerciseRows.add(rowNum);",
  },
  {
    file: "Sheet.gs",
    name: "edits on another sheet are no longer ignored",
    find: "  if (sheet.getSheetId() !== getSheet_().getSheetId()) {\n    return false;\n  }",
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
    find:
      "  if (violation) {\n" +
      "    throw new Error(`syncOnEdit: date validation failed: ${violation}`);\n" +
      "  }",
    replace: "  if (violation) {\n    return;\n  }",
  },
  {
    file: "Main.gs",
    name: "a time-based trigger throws on a date violation instead of skipping",
    find:
      "  console.error(\n" +
      "    `${triggerName}: date validation failed; skipping: ${violation}`,\n" +
      "  );\n" +
      "  return true;",
    replace: "  throw new Error(`${triggerName}: ${violation}`);",
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
    find:
      "    if (knownNames[c.name]) {\n" +
      "      return;\n" +
      "    }\n" +
      "    if (c.googleWebClientId && ourClientIds[c.googleWebClientId]) {",
    replace:
      "    if (c.googleWebClientId && ourClientIds[c.googleWebClientId]) {",
  },

  // ---- Manual entry points ------------------------------------------------
  {
    // Only its date-validation abort was covered before, so the clear could
    // have been a no-op, or bounded to the first row, unnoticed.
    file: "Main.gs",
    name: "resyncAllRows no longer clears the exercise stamps",
    find: "  sheet.getRange(2, exerciseCol, dataRowCount, 1).setValues(blanks);",
    replace: "",
  },
  {
    file: "Main.gs",
    name: "resyncAllRows no longer clears the weight stamps",
    find: "  sheet.getRange(2, weightSyncedAtCol, dataRowCount, 1).setValues(blanks);",
    replace: "",
  },
  {
    file: "Main.gs",
    name: "resyncAllRows clears only the first data row",
    find: "  const dataRowCount = lastRow - 1;",
    replace: "  const dataRowCount = 1;",
  },
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
    find:
      "  if (!lock.tryLock(LOCK_WAIT_MS)) {\n" +
      '    console.warn("backstop: another run holds the lock; skipping this run.");\n' +
      "    return;\n" +
      "  }",
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
