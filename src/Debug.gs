// Manual live-API checks. Nothing here runs on the sync path and nothing calls
// it: these are the functions you pick by hand in the Apps Script editor when
// the API's behavior needs measuring rather than reading about. Two entry
// points:
//
//   debugRunAll()  - runs every check and logs ONE report block, meant to be
//                    selected and pasted somewhere else whole.
//   debugCleanup() - deletes any debug datapoint an interrupted run left
//                    behind. debugRunAll() already deletes its own, so this is
//                    only needed after a 6-minute kill or a thrown delete.
//
// The report is written to leave this spreadsheet, so it carries no
// identifiers: resource names, user ids and OAuth client ids are stripped on
// the way out (debugRedact_).
//
// A debug run writes to real Health data. Its datapoint sits at 03:00 (see
// below), lives for the seconds the run takes, and is sync-created and
// referenced by no row, so even a leaked one is inside the backstop's orphan
// reconciliation window and gets collected on the next pass.

// 03:00 local, ten minutes long: clear of the synthetic noon slot a backfill
// row would occupy and of any plausible edit window, so a datapoint left behind
// for a few minutes cannot be picked up as a row's foreign match.
const DEBUG_START_HOUR_ = 3;
const DEBUG_DURATION_MIN_ = 10;
// Every debug datapoint's notes start with this, which is what debugCleanup()
// matches on. The notes check below appends to it rather than replacing it, so
// a datapoint stays findable after being patched.
const DEBUG_NOTES_PREFIX_ = "debugRunAll: throwaway, safe to delete";

// One entry per field, each swapped into a full copy of the datapoint's current
// `exercise` object. `read` pulls the value worth comparing out of the body we
// sent and out of the GET that follows. `expect` is what the Health API team's
// 2026-08 reply implies, so the report can flag a surprise instead of leaving
// the reader to work out which lines matter:
//   - activeDuration: reported fixed -> applied.
//   - displayName: server-derived from exerciseType for every type but OTHER,
//     so a custom name on a STRENGTH_TRAINING datapoint -> ignored.
//   - notes / interval: not mentioned in the reply. Expected applied because
//     that is what the REST reference documents; "ignored" here is the finding,
//     and is what decides whether the exercise sync can PATCH instead of
//     delete+recreate.
const DEBUG_PATCH_CHECKS_ = [
  {
    expect: "applied",
    field: "activeDuration",
    // Deliberately not the create-time value (the interval's length, a round
    // number of minutes), so an echo of what we created cannot be mistaken for
    // an applied update.
    patch: () => ({ activeDuration: "137s" }),
    read: (ex) => ex.activeDuration || null,
  },
  {
    expect: "applied",
    field: "notes",
    patch: () => ({ notes: `${DEBUG_NOTES_PREFIX_} (patched)` }),
    read: (ex) => ex.notes || null,
  },
  {
    expect: "ignored",
    field: "displayName",
    patch: () => ({ displayName: "Debug Custom Name" }),
    read: (ex) => ex.displayName || null,
  },
  {
    expect: "applied",
    field: "interval.startTime",
    patch: (ex) => ({
      interval: debugShiftedInterval_(ex.interval, 5 * 60 * 1000),
    }),
    read: (ex) => (ex.interval && ex.interval.startTime) || null,
  },
];

function debugExerciseOf_(dataPoint) {
  return (dataPoint && dataPoint.exercise) || {};
}

function debugIsoUtc_(ms) {
  return Utilities.formatDate(new Date(ms), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

// Everything that identifies this account or its data, removed from the report
// before it is logged. Resource ids are what the API hands back, so they are
// replaced rather than dropped: the report still shows that a name was there.
function debugRedact_(text) {
  return String(text)
    .replace(/users\/[^/\s"]+/g, "users/USER")
    .replace(/dataPoints\/[^/\s",}]+/g, "dataPoints/DATAPOINT")
    .replace(/("googleWebClientId":)"[^"]*"/g, '$1"CLIENT"');
}

// Shift an interval's physical times by `deltaMs`, keeping the offsets. Any
// civil* member the GET returned is dropped rather than shifted: echoing a
// civil time that no longer agrees with the physical one would leave the server
// to pick a winner, and the check would not know which field it had tested.
function debugShiftedInterval_(interval, deltaMs) {
  const out = {};
  Object.keys(interval || {}).forEach((k) => {
    if (k.indexOf("civil") !== 0) {
      out[k] = interval[k];
    }
  });
  if (interval && interval.startTime) {
    out.startTime = debugIsoUtc_(
      new Date(interval.startTime).getTime() + deltaMs,
    );
  }
  if (interval && interval.endTime) {
    out.endTime = debugIsoUtc_(new Date(interval.endTime).getTime() + deltaMs);
  }
  return out;
}

// The debug datapoint's 03:00 slot: today's once that hour has passed, and
// yesterday's otherwise, so a run at 01:00 doesn't place a session in the
// future while measuring something else.
function debugStartSlot_(tz, now) {
  const base =
    civilDateParts_(tz, now).hours > DEBUG_START_HOUR_
      ? now
      : new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const p = civilDateParts_(tz, base);
  return localCivilToUtcMs_(tz, p.year, p.month, p.day, DEBUG_START_HOUR_, 0);
}

function debugVerdict_(result) {
  if (result.error) {
    return "ERROR";
  }
  return result.applied ? "applied" : "ignored";
}

// One report line per check, padded so a column of them reads as a table, with
// the checks whose outcome contradicts the reply marked. An ERROR is always
// marked: a 500 is a finding whichever field it lands on.
function debugFormatResult_(result) {
  const verdict = debugVerdict_(result);
  const flag =
    verdict === result.expect ? "" : `   <-- expected ${result.expect}`;
  const pad = (s, n) => {
    let out = String(s);
    while (out.length < n) {
      out += " ";
    }
    return out;
  };
  return (
    `  ${pad(result.field, 30)}${pad(verdict, 9)}` +
    `sent ${JSON.stringify(result.want)}, read back ${JSON.stringify(result.got)}${flag}`
  );
}

// One field: GET the current state, PATCH a full body with that field swapped,
// GET again to see whether it stuck. The full body is the shape the server
// accepted (200 + done:true) back when it applied nothing, so a partial-body
// 500 cannot be mistaken for a field that refuses to merge.
function runDebugPatchCheck_(name, check) {
  const before = debugExerciseOf_(getDataPoint(name));
  const body = Object.assign({}, before, check.patch(before));
  const result = {
    applied: false,
    error: null,
    expect: check.expect,
    field: check.field,
    got: null,
    want: check.read(body),
  };
  try {
    patchExercise(name, body);
  } catch (err) {
    result.error = String(err);
  }
  result.got = check.read(debugExerciseOf_(getDataPoint(name)));
  result.applied = JSON.stringify(result.got) === JSON.stringify(result.want);
  return result;
}

// The partial-body case from the bug report: `{ exercise: { activeDuration } }`
// and nothing else used to 500 INTERNAL. Run last, because that 500 costs the
// full httpJson_ retry backoff (~7s) before it gives up.
function runDebugPartialBodyCheck_(name) {
  const result = {
    applied: false,
    error: null,
    expect: "applied",
    field: "activeDuration (partial body)",
    got: null,
    // Distinct from the full-body check's value for the same reason.
    want: "251s",
  };
  try {
    patchExercise(name, { activeDuration: result.want });
  } catch (err) {
    result.error = String(err);
  }
  result.got = debugExerciseOf_(getDataPoint(name)).activeDuration || null;
  result.applied = result.got === result.want;
  return result;
}

// Runs every live-API check and logs one report. Copy the whole block between
// the BEGIN and END lines.
//
// What it covers:
//   1. create: the datapoint is created WITHOUT a displayName, so the GET that
//      follows says whether the server really does derive one from
//      exerciseType (the reason the sync stopped sending its own).
//   2. patch: each field of DEBUG_PATCH_CHECKS_ in turn, full body.
//   3. partial body: the 500 from the original report.
//   4. delete: always, even when a check throws.
function debugRunAll() {
  const tz = getTz_();
  const start = debugStartSlot_(tz, new Date());
  // One offset for both ends: DST transitions land at 02:00 local, so a ten
  // minute session starting at 03:00 never straddles one.
  const endUtcMs = start.utcMs + DEBUG_DURATION_MIN_ * 60 * 1000;

  const lines = [];
  lines.push("----- BEGIN HEALTH API DEBUG REPORT -----");
  lines.push("(everything between the BEGIN and END lines is the report)");
  lines.push(`run at:       ${new Date().toISOString()} (script tz ${tz})`);
  lines.push(`api base:     ${HEALTH_API_BASE}`);

  let name = null;
  try {
    name = createExerciseAt(
      start.utcMs,
      start.offsetSeconds,
      endUtcMs,
      start.offsetSeconds,
      DEBUG_NOTES_PREFIX_,
    );
  } catch (err) {
    lines.push(`create:       FAILED ${err}`);
    lines.push("----- END HEALTH API DEBUG REPORT -----");
    const failed = debugRedact_(lines.join("\n"));
    console.log(failed);
    return failed;
  }

  const results = [];
  try {
    const created = debugExerciseOf_(getDataPoint(name));
    lines.push("create:       OK (no displayName sent)");
    lines.push(
      `  interval read back:       ${JSON.stringify(created.interval || null)}`,
    );
    lines.push(
      `  exerciseType read back:   ${JSON.stringify(created.exerciseType || null)}`,
    );
    lines.push(
      `  displayName read back:    ${JSON.stringify(created.displayName || null)}`,
    );
    lines.push(
      `  activeDuration read back: ${JSON.stringify(created.activeDuration || null)}`,
    );
    lines.push("");
    lines.push("patch checks (full body, one field swapped):");
    DEBUG_PATCH_CHECKS_.forEach((check) => {
      results.push(runDebugPatchCheck_(name, check));
    });
    results.push(runDebugPartialBodyCheck_(name));
    results.forEach((r) => {
      lines.push(debugFormatResult_(r));
    });
    const errors = results.filter((r) => r.error);
    if (errors.length > 0) {
      lines.push("");
      lines.push("errors in full:");
      errors.forEach((r) => {
        lines.push(`  ${r.field}: ${r.error}`);
      });
    }
    lines.push("");
    lines.push(
      `final exercise: ${JSON.stringify(debugExerciseOf_(getDataPoint(name)))}`,
    );
  } catch (err) {
    lines.push(`ABORTED: ${err}`);
  } finally {
    try {
      deleteDataPointsByName([name]);
      lines.push("cleanup:      deleted the debug datapoint");
    } catch (err) {
      lines.push(`cleanup:      FAILED (${err}); run debugCleanup()`);
    }
  }
  lines.push("----- END HEALTH API DEBUG REPORT -----");

  const report = debugRedact_(lines.join("\n"));
  console.log(report);
  return report;
}

// Delete any debug datapoint left behind by an interrupted debugRunAll(), on
// the two civil dates its 03:00 slot can fall on. Matching is on the notes
// marker, which nothing but debugRunAll() writes, so a real session can never
// be caught by it. Safe to run when there is nothing to clean up.
function debugCleanup() {
  const now = new Date();
  const dates = [new Date(now.getTime() - 24 * 60 * 60 * 1000), now];
  const lines = ["----- BEGIN HEALTH API DEBUG CLEANUP -----"];
  const names = [];
  const seen = {};
  dates.forEach((date) => {
    let points;
    try {
      points = listExercisesOnDate(date);
    } catch (err) {
      lines.push(`  ${ymd(date)}: list failed (${err})`);
      return;
    }
    points.forEach((p) => {
      const notes = (p.exercise && p.exercise.notes) || "";
      // Deduped because the two dates are listed independently: a name that
      // came back on both would otherwise be deleted twice, and the second
      // delete's 404 would be reported as a failure that isn't one.
      if (p.name && notes.indexOf(DEBUG_NOTES_PREFIX_) === 0 && !seen[p.name]) {
        seen[p.name] = true;
        names.push(p.name);
      }
    });
  });
  if (names.length === 0) {
    lines.push("  nothing to clean up");
  }
  names.forEach((n) => {
    try {
      deleteDataPointsByName([n]);
      lines.push(`  deleted ${n}`);
    } catch (err) {
      lines.push(`  FAILED to delete ${n}: ${err}`);
    }
  });
  lines.push("----- END HEALTH API DEBUG CLEANUP -----");
  const report = debugRedact_(lines.join("\n"));
  console.log(report);
  return report;
}
