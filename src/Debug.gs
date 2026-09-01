// Read-only probes for the "two workouts on one day" investigation.
// Run `probeRunAllDiagnostics` from the Apps Script editor and copy the whole
// Executions log. Delete this file (and `npm run deploy`) when done.
//
// Every line is redacted on the way out (probeRedact_): the log is written to
// leave this spreadsheet, so numeric user ids and OAuth client ids are
// stripped and datapoint ids keep only their last four digits, enough to
// cross-reference sessions within one log without identifying the account.

// Strip identifiers but keep datapoints distinguishable: two ids are
// vanishingly unlikely to share their last four digits within one day's log.
function probeRedact_(text) {
  return String(text)
    .replace(/users\/[^/\s"]+\//g, "users/USER/")
    .replace(/dataPoints\/[^\s",\]}]*([^\s",\]}]{4})/g, "dataPoints/…$1")
    .replace(/[a-z0-9-]+\.apps\.googleusercontent\.com/g, "CLIENT");
}

function probeLog_(text) {
  console.log(probeRedact_(text));
}

// NaN/blank-safe timestamp formatter shared by every probe, so a datapoint
// with a missing interval prints "(none)" instead of throwing out of the
// forEach and losing the rest of the report.
function probeFmt_(tz, value) {
  const ms = value instanceof Date ? value.getTime() : value;
  if (ms === null || ms === undefined || isNaN(ms)) {
    return "(none)";
  }
  return Utilities.formatDate(new Date(ms), tz, "yyyy-MM-dd HH:mm:ss z");
}

// One sheet read per execution: module-level bindings reset per invocation, so
// an aggregator run reads the sheet once and every probe reports on the same
// snapshot rather than each taking its own.
let probeReadRowsCache_ = null;
function probeReadRows_() {
  if (!probeReadRowsCache_) {
    probeReadRowsCache_ = readRows();
  }
  return probeReadRowsCache_;
}

// The last two rows with exercise text, in sheet (ascending date) order, so a
// bodyweight-only row logged mid-investigation does not shift the probes off
// the days under study. No-arg discovery per the editor-runnable convention.
function probeTargetRows_() {
  return probeReadRows_()
    .rows.filter((r) => r.hasExerciseText)
    .slice(-2);
}

function probeSheetRows() {
  const tz = getTz_();
  probeTargetRows_().forEach((r) => {
    try {
      probeLog_(`--- row ${r.rowNum} date=${ymd(r.date)} ---`);
      probeLog_(`  exerciseSyncedAt: ${r.exerciseSyncedAt || "(blank)"}`);
      probeLog_(`  weightSyncedAt:   ${r.weightSyncedAt || "(blank)"}`);
      probeLog_(
        `  exerciseFirstEditedAt: ${probeFmt_(tz, r.exerciseFirstEditedAt)}`,
      );
      probeLog_(
        `  exercisesLastEditedAt: ${probeFmt_(tz, r.exercisesLastEditedAt)}`,
      );
      probeLog_(`  weightEditedAt:        ${probeFmt_(tz, r.weightEditedAt)}`);
      if (r.exerciseFirstEditedAt && r.exercisesLastEditedAt) {
        const span =
          r.exercisesLastEditedAt.getTime() - r.exerciseFirstEditedAt.getTime();
        probeLog_(
          `  edit span: ${humanizeMs_(span)} (MAX=${humanizeMs_(
            MAX_EXERCISE_DURATION_MS,
          )}, spansWorkout=${exerciseEditSpansWorkout_(r)})`,
        );
      }
      probeLog_(
        `  matchedHealthSessions: ${JSON.stringify(r.matchedHealthSessions)}`,
      );
      probeLog_(`  exerciseEditTimes: ${JSON.stringify(r.exerciseEditTimes)}`);
      probeLog_(`  healthIds: ${JSON.stringify(r.healthIds)}`);
      probeLog_(`  bodyweight: ${r.bodyweight}`);
      probeLog_(`  exercises (parsed): ${JSON.stringify(r.exercises)}`);
    } catch (err) {
      probeLog_(`  row ${r.rowNum} FAILED: ${err}`);
    }
  });
}

function probeExercisesOnTargetDates() {
  const tz = getTz_();
  const ours = {};
  probeReadRows_().allHealthIds.forEach((n) => {
    ours[toMeName_(n)] = true;
  });
  probeTargetRows_().forEach((r) => {
    probeLog_(`=== exercise dataPoints on ${ymd(r.date)} ===`);
    // Per-date, so a transient listing failure on one day still reports the
    // other; the sync's own foreign-match listing degrades the same way.
    let points;
    try {
      points = listExercisesOnDate(r.date);
    } catch (err) {
      probeLog_(`  list FAILED: ${err}`);
      return;
    }
    probeLog_(`  count: ${points.length}`);
    points.forEach((p) => {
      const ex = p.exercise || {};
      const i = ex.interval || {};
      const app = (p.dataSource && p.dataSource.application) || null;
      const startMs = i.startTime ? new Date(i.startTime).getTime() : NaN;
      const endMs = i.endTime ? new Date(i.endTime).getTime() : NaN;
      probeLog_(`  - name: ${p.name}`);
      probeLog_(`    tracked by sheet: ${Boolean(ours[toMeName_(p.name)])}`);
      probeLog_(`    exerciseType: ${ex.exerciseType}`);
      probeLog_(
        `    recordingMethod: ${p.dataSource && p.dataSource.recordingMethod}`,
      );
      probeLog_(
        `    googleWebClientId: ${(app && app.googleWebClientId) || "(null: foreign)"}`,
      );
      probeLog_(
        `    interval: ${probeFmt_(tz, startMs)} .. ${probeFmt_(tz, endMs)}` +
          `${isNaN(endMs - startMs) ? "" : ` (${humanizeMs_(endMs - startMs)})`}`,
      );
      probeLog_(`    activeDuration: ${ex.activeDuration}`);
      probeLog_(`    metricsSummary: ${JSON.stringify(ex.metricsSummary)}`);
      probeLog_(`    notes: ${JSON.stringify(ex.notes || null)}`);
    });
  });
}

function probeWeightOnTargetDates() {
  probeTargetRows_().forEach((r) => {
    probeLog_(`=== weight dataPoints on ${ymd(r.date)} ===`);
    let points;
    try {
      points = listDataPointsByCivilDate_(
        "weight",
        "weight.sample_time.civil_time",
        r.date,
      );
    } catch (err) {
      probeLog_(`  list FAILED: ${err}`);
      return;
    }
    points.forEach((p) => {
      const app = (p.dataSource && p.dataSource.application) || null;
      probeLog_(
        `  - ${p.name} sampleTime=${JSON.stringify(
          p.weight && p.weight.sampleTime,
        )} grams=${p.weight && p.weight.weightGrams} client=${
          (app && app.googleWebClientId) || "(null: foreign)"
        }`,
      );
    });
  });
}

// What an exercise re-sync of the target rows would do right now, without
// writing: claimed sessions, the per-session split, each group's resolved
// interval and notes, and the keep / create / delete plan against the row's
// existing datapoints. Mirrors syncOneRow_'s exercise phase step for step
// (prior GETs included, so 'prior' timing and the idempotency skip show up as
// they would in a real pass); the ready set is the currently exercise-dirty
// rows plus the targets, matching a "Resync selected rows" over the targets.
function probeForeignMatchPlan() {
  const { allHealthIds, allMatchedSessions, rows } = probeReadRows_();
  const targets = probeTargetRows_();
  const ready = {};
  const readyRows = [];
  rows.forEach((r) => {
    if (!r.exerciseSyncedAt) {
      ready[r.rowNum] = true;
      readyRows.push(r);
    }
  });
  targets.forEach((r) => {
    if (!ready[r.rowNum]) {
      ready[r.rowNum] = true;
      readyRows.push(r);
    }
  });
  const plan = resolveForeignMatches_(
    allHealthIds,
    allMatchedSessions,
    readyRows,
  );
  const tz = getTz_();
  targets.forEach((r) => {
    const matches = plan[r.rowNum] || [];
    probeLog_(`--- row ${r.rowNum} (${ymd(r.date)}) ---`);
    probeLog_(`  claimed sessions: ${matches.length}`);
    matches.forEach((m) => {
      probeLog_(
        `    ${m.name} ${probeFmt_(tz, m.startUtcMs)} .. ${probeFmt_(tz, m.endUtcMs)}`,
      );
    });

    const priorIds = splitHealthIdsByType_(r.healthIds).exercise;
    const priors = [];
    priorIds.forEach((name) => {
      try {
        priors.push({ dp: getDataPoint(name), name });
      } catch (err) {
        probeLog_(`  GET prior ${name} FAILED (sync would recreate): ${err}`);
      }
    });
    const priorExercise = priors.length > 0 ? priors[0].dp : null;

    const groups = partitionExercisesBySession_(
      matches,
      r.exerciseEditTimes,
      r.exerciseFirstEditedAt,
      r.exercises,
    );
    const seenIntervals = {};
    const claimed = {};
    probeLog_(`  plan (${groups.length} group(s)):`);
    groups.forEach((g) => {
      const timing = resolveRowTiming_(r, priorExercise, g.session);
      const ex = timing.exercise;
      const key = `${ex.startUtcMs}-${ex.endUtcMs}`;
      if (seenIntervals[key]) {
        probeLog_(
          `    DROPPED (interval already taken; a same-key POST would be ` +
            `silently discarded): session=${g.session ? g.session.name : "(none)"}`,
        );
        return;
      }
      seenIntervals[key] = true;
      const notes = buildNotes(ex.endUtcMs - ex.startUtcMs, g.exercises);
      const match = priors.filter(
        (p) =>
          !claimed[p.name] &&
          exerciseUnchanged_(p.dp, ex.startUtcMs, ex.endUtcMs, notes),
      )[0];
      if (match) {
        claimed[match.name] = true;
      }
      probeLog_(
        `    [${timing.exerciseSource}] ${probeFmt_(tz, ex.startUtcMs)} .. ` +
          `${probeFmt_(tz, ex.endUtcMs)} session=${
            g.session ? g.session.name : "(none)"
          } -> ${match ? `KEEP ${match.name}` : "CREATE"}`,
      );
      probeLog_(`      notes: ${JSON.stringify(notes)}`);
    });
    const stale = priorIds.filter((n) => !claimed[n]);
    if (stale.length > 0) {
      probeLog_(`  would DELETE: ${JSON.stringify(stale)}`);
    }
  });
}

// One-shot repair for 2026-08-30 after a paste overwrote the row with date
// values at 2026-08-31 23:41 EDT: the exercise and Weight cells, Exercise
// First Edited At, Created Health IDs, and Matched Health Session were all
// replaced, and the onEdit sync then stamped the emptied row, leaving the
// row's two live datapoints (the combined exercise session and the weight
// sample) untracked, where the next backstop would collect them as orphans.
//
// Restores the cell content and edit timestamps from the values recorded in
// this investigation, re-tracks the datapoints by rediscovering their
// resource names from the API (ours = non-null googleWebClientId on that
// date; nothing identifying is hardcoded here), clears the pasted date
// values from the other exercise columns, then hands off to
// backfillTwoWorkoutSplit20260830 for the split, sync, and report. MUTATING;
// safe to re-run.
function restoreAndSplitRow20260830() {
  const targetDateKey = "2026-08-30";
  const violation = validateRowDates_();
  if (violation) {
    probeLog_(`ABORT: date validation failed: ${violation}`);
    return;
  }
  const sheet = getSheet_();
  const { headers, map } = getHeaderMap_(sheet);
  const need = [
    DATE_COLUMN_HEADER,
    WEIGHT_COLUMN_HEADER,
    "Bench press",
    "Dumbbell shoulder press",
    EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER,
    EXERCISES_LAST_EDITED_AT_COLUMN_HEADER,
    WEIGHT_EDITED_AT_COLUMN_HEADER,
    HEALTH_IDS_COLUMN_HEADER,
    MATCHED_HEALTH_SESSION_COLUMN_HEADER,
    EXERCISE_SYNCED_AT_COLUMN_HEADER,
  ];
  const missingCols = need.filter((h) => !map[h]);
  if (missingCols.length > 0) {
    probeLog_(`ABORT: missing column(s): ${missingCols.join(", ")}`);
    return;
  }
  const { rows } = readRows();
  const row = rows.filter((r) => ymd(r.date) === targetDateKey)[0];
  if (!row) {
    probeLog_(`ABORT: no row dated ${targetDateKey}.`);
    return;
  }

  // Re-track the row's live datapoints: ours are the ones on that date whose
  // dataSource carries a web client id (device/first-party sessions carry
  // null). The exercise list is every strength session; the weight list is
  // already projected to { name, googleWebClientId }.
  const date = row.date;
  const ourIds = [];
  listWeightOnDate(date).forEach((c) => {
    if (c.googleWebClientId) {
      ourIds.push(c.name);
    }
  });
  listStrengthOnDate(date).forEach((c) => {
    if (c.googleWebClientId) {
      ourIds.push(c.name);
    }
  });
  if (ourIds.length === 0) {
    probeLog_(
      "ABORT: no sync-created datapoints found on the date; nothing to re-track.",
    );
    return;
  }
  probeLog_(`re-tracking ${ourIds.length} datapoint(s).`);

  // Clear the pasted date values from every non-Date, non-managed column in
  // the row: a real exercise or Weight cell never legitimately holds a Date.
  const managedColNums = MANAGED_COLUMN_HEADERS.map((h) => map[h]).filter(
    (c) => c,
  );
  for (let c = 1; c <= headers.length; c++) {
    if (c === map[DATE_COLUMN_HEADER] || managedColNums.indexOf(c) !== -1) {
      continue;
    }
    const cell = sheet.getRange(row.rowNum, c);
    if (cell.getValue() instanceof Date) {
      cell.setValue("");
    }
  }

  // The row's content and timestamps as recorded before the paste.
  const set = (header, value) =>
    sheet.getRange(row.rowNum, map[header]).setValue(value);
  set("Bench press", "175x6x6");
  set("Dumbbell shoulder press", "35x6x5");
  // The paste left the Weight cell date-formatted, so a numeric write renders
  // as a 1900-era date and reads back as a Date, which parseBodyweight
  // rejects. Reset the format before writing the value.
  sheet
    .getRange(row.rowNum, map[WEIGHT_COLUMN_HEADER])
    .setNumberFormat("0.###");
  set(WEIGHT_COLUMN_HEADER, 256);
  set(EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER, "2026-08-30T17:09:18.000Z");
  set(EXERCISES_LAST_EDITED_AT_COLUMN_HEADER, "2026-08-30T18:09:38.000Z");
  set(WEIGHT_EDITED_AT_COLUMN_HEADER, "2026-08-30T17:04:43.000Z");
  set(HEALTH_IDS_COLUMN_HEADER, JSON.stringify(ourIds));
  set(MATCHED_HEALTH_SESSION_COLUMN_HEADER, "");
  SpreadsheetApp.flush();
  probeLog_(`row ${row.rowNum} restored.`);

  backfillTwoWorkoutSplit20260830();
}

// One-shot backfill for 2026-08-30, the day two workouts were logged into one
// row before Exercise Edit Times existed. Writes each exercise's first-edit
// time (from the row's own recorded timestamps: 13:09:18 EDT falls inside
// workout 1, 14:09:38 EDT inside workout 2), re-dirties the exercise phase,
// syncs, then runs the read-only diagnostics. MUTATING, so deliberately not
// part of probeRunAllDiagnostics; safe to re-run (existing edit-time entries
// are kept, and an unchanged row re-syncs to a no-op). The sync's own log
// lines are unredacted; paste from the REDACTED REPORT banner down.
function backfillTwoWorkoutSplit20260830() {
  const targetDateKey = "2026-08-30";
  const backfill = {
    "Bench press": { first: "2026-08-30T17:09:18.000Z" },
    "Dumbbell shoulder press": { first: "2026-08-30T18:09:38.000Z" },
  };

  const violation = validateRowDates_();
  if (violation) {
    probeLog_(`ABORT: date validation failed: ${violation}`);
    return;
  }
  const { exerciseEditTimesCol, exerciseSyncedAtCol, rows } = readRows();
  if (!exerciseEditTimesCol) {
    probeLog_('ABORT: Exercise Edit Times column missing; run "Run setup".');
    return;
  }
  const row = rows.filter((r) => ymd(r.date) === targetDateKey)[0];
  if (!row) {
    probeLog_(`ABORT: no row dated ${targetDateKey}.`);
    return;
  }
  const names = row.exercises.map((e) => e.name);
  const missing = Object.keys(backfill).filter((n) => names.indexOf(n) === -1);
  if (missing.length > 0) {
    probeLog_(
      `ABORT: row ${row.rowNum} lacks expected exercise(s): ${missing.join(
        ", ",
      )} (has: ${names.join(", ")}).`,
    );
    return;
  }

  const sheet = getSheet_();
  const cell = sheet.getRange(row.rowNum, exerciseEditTimesCol);
  const current = parseExerciseEditTimes_(cell.getValue());
  let added = 0;
  Object.keys(backfill).forEach((name) => {
    if (!current[name]) {
      current[name] = backfill[name];
      added++;
    }
  });
  if (added > 0) {
    cell.setValue(JSON.stringify(current));
    probeLog_(`row ${row.rowNum}: wrote ${added} edit-time entr(y/ies).`);
  } else {
    probeLog_(`row ${row.rowNum}: edit times already present; not rewritten.`);
  }

  reDirtyRows_([row.rowNum], { exerciseCol: exerciseSyncedAtCol });
  const result = syncDirtyRows(LOCK_WAIT_MS);
  probeLog_(
    result
      ? `sync: ok=${result.ok} errors=${result.errors} deferred=${result.deferred}`
      : "sync: skipped (lock held); run Sync > Run now shortly.",
  );

  probeReadRowsCache_ = null;
  probeLog_("\n======== REDACTED REPORT BELOW: paste from here ========");
  probeRunAllDiagnostics();
}

function probeRunAllDiagnostics() {
  const probes = [
    ["probeSheetRows", probeSheetRows],
    ["probeExercisesOnTargetDates", probeExercisesOnTargetDates],
    ["probeWeightOnTargetDates", probeWeightOnTargetDates],
    ["probeForeignMatchPlan", probeForeignMatchPlan],
  ];
  probeLog_(`timezone: ${getTz_()}   now: ${new Date().toISOString()}`);
  probes.forEach(([label, fn]) => {
    probeLog_(`\n################ ${label} ################`);
    try {
      fn();
    } catch (err) {
      probeLog_(`  FAILED: ${err && err.stack ? err.stack : err}`);
    }
  });
}
