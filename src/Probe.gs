// Live-API probes. Nothing here runs on the sync path: these are functions you
// run by hand from the Apps Script editor's function picker (or the console)
// when the API's behavior needs re-checking, and whose findings then get
// written up as comments next to the code that depends on them. The pattern
// predates this file: the filter-spelling note in HealthApi.gs came from a
// throwaway Debug.gs of the same shape.
//
// A probe writes to the user's real Health data, so every one of them cleans up
// after itself and says loudly what it left behind if it cannot. The safety net
// if a probe dies mid-run: its datapoint is sync-created, dated within the last
// day, and referenced by no row, so it is well inside the backstop's orphan
// reconciliation window and gets collected on the next pass.

// 03:00 local, ten minutes long: outside the synthetic noon slot a backfill row
// would occupy, and far from any plausible edit window, so a probe datapoint
// left behind for a few minutes cannot be picked up as a row's foreign match.
const PROBE_START_HOUR_ = 3;
const PROBE_DURATION_MIN_ = 10;
const PROBE_NOTES_ = "probeExercisePatch: safe to delete";

// How each field is exercised. `patch` returns the fields to swap into a full
// copy of the datapoint's current `exercise` object; `read` pulls the value
// worth comparing back out (of the body we sent and of the GET that follows).
// `expect` is what the API team's 2026-08 reply says should happen, so the log
// can call out a surprise rather than just printing values:
//   - activeDuration: "the issue has been resolved" -> applied.
//   - displayName: server-derived from exerciseType for every type but OTHER,
//     so a custom name on a STRENGTH_TRAINING datapoint -> ignored.
//   - notes / interval: never mentioned in the reply. Expected applied because
//     that is what the REST reference documents; an "ignored" here is the
//     finding, and is what decides whether the sync's delete+recreate path can
//     become a PATCH.
const EXERCISE_PATCH_PROBES_ = [
  {
    expect: "applied",
    field: "activeDuration",
    // Deliberately shorter than the interval: an echo of the create-time value
    // (interval length) is then distinguishable from a real update.
    patch: () => ({ activeDuration: "600s" }),
    read: (ex) => ex.activeDuration || null,
  },
  {
    expect: "applied",
    field: "notes",
    patch: () => ({ notes: `${PROBE_NOTES_} (patched)` }),
    read: (ex) => ex.notes || null,
  },
  {
    expect: "ignored",
    field: "displayName",
    patch: () => ({ displayName: "Probe Custom Name" }),
    read: (ex) => ex.displayName || null,
  },
  {
    expect: "applied",
    field: "interval",
    patch: (ex) => ({ interval: shiftedInterval_(ex.interval, 5 * 60 * 1000) }),
    read: (ex) => (ex.interval && ex.interval.startTime) || null,
  },
];

function exerciseOf_(dataPoint) {
  return (dataPoint && dataPoint.exercise) || {};
}

function isoUtc_(ms) {
  return Utilities.formatDate(new Date(ms), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

// Shift an interval's physical times by `deltaMs`, keeping the offsets. Any
// civil* member the GET returned is dropped rather than shifted: echoing a
// civil time that no longer agrees with the physical one would leave the
// server to pick a winner, and the probe would not know which field it had
// actually tested.
function shiftedInterval_(interval, deltaMs) {
  const out = {};
  Object.keys(interval || {}).forEach((k) => {
    if (k.indexOf("civil") !== 0) {
      out[k] = interval[k];
    }
  });
  if (interval && interval.startTime) {
    out.startTime = isoUtc_(new Date(interval.startTime).getTime() + deltaMs);
  }
  if (interval && interval.endTime) {
    out.endTime = isoUtc_(new Date(interval.endTime).getTime() + deltaMs);
  }
  return out;
}

function formatProbeResult_(result) {
  let verdict;
  if (result.error) {
    verdict = `ERROR (${result.error})`;
  } else if (result.applied) {
    verdict = "applied";
  } else {
    verdict = "ignored (no-op)";
  }
  const surprise =
    !result.error && verdict.indexOf(result.expect) === 0
      ? ""
      : `  <-- expected ${result.expect}`;
  return (
    `probeExercisePatch: ${result.field}: ${verdict}; ` +
    `sent ${JSON.stringify(result.want)}, read back ${JSON.stringify(result.got)}${surprise}`
  );
}

// One field, one PATCH, one GET. The body is a full copy of the datapoint's
// current `exercise` with the probe's fields swapped in, which is the shape the
// server accepted (200 + done:true) back when it applied nothing.
function runExercisePatchProbe_(name, probe) {
  const before = exerciseOf_(getDataPoint(name));
  const body = Object.assign({}, before, probe.patch(before));
  const result = {
    applied: false,
    error: null,
    expect: probe.expect,
    field: probe.field,
    got: null,
    want: probe.read(body),
  };
  try {
    patchExercise(name, body);
  } catch (err) {
    result.error = String(err);
  }
  result.got = probe.read(exerciseOf_(getDataPoint(name)));
  result.applied = JSON.stringify(result.got) === JSON.stringify(result.want);
  console.info(formatProbeResult_(result));
  return result;
}

// The partial-body case from the bug report: `{ exercise: { activeDuration } }`
// with nothing else used to 500. Runs last because a 500 costs ~7s of
// httpJson_ backoff before it gives up, and because it is the one probe whose
// failure is the expected outcome.
function runMinimalBodyProbe_(name) {
  const result = {
    applied: false,
    error: null,
    expect: "applied",
    field: "activeDuration (partial body)",
    got: null,
    want: "900s",
  };
  try {
    patchExercise(name, { activeDuration: result.want });
  } catch (err) {
    result.error = String(err);
  }
  result.got = exerciseOf_(getDataPoint(name)).activeDuration || null;
  result.applied = result.got === result.want;
  console.info(formatProbeResult_(result));
  return result;
}

// The probe datapoint's 03:00 slot, on today's date once that hour has passed
// and on yesterday's otherwise: a run at 01:00 would otherwise place a session
// in the future, which is not a shape worth asking the API to accept while
// measuring something else.
function probeStart_(tz, now) {
  const base =
    civilDateParts_(tz, now).hours > PROBE_START_HOUR_
      ? now
      : new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const p = civilDateParts_(tz, base);
  return localCivilToUtcMs_(tz, p.year, p.month, p.day, PROBE_START_HOUR_, 0);
}

// Create a throwaway STRENGTH_TRAINING datapoint, PATCH one field at a time,
// GET after each to see whether the change stuck, then delete it. Returns the
// per-field results and logs them; read the Executions log (or the return
// value) for the verdict.
function probeExercisePatch() {
  const tz = getTz_();
  const start = probeStart_(tz, new Date());
  // One offset for both ends: DST transitions land at 02:00 local, so a ten
  // minute session starting at 03:00 never straddles one.
  const endUtcMs = start.utcMs + PROBE_DURATION_MIN_ * 60 * 1000;
  const name = createExerciseAt(
    start.utcMs,
    start.offsetSeconds,
    endUtcMs,
    start.offsetSeconds,
    PROBE_NOTES_,
  );
  console.info(`probeExercisePatch: created ${name}`);

  const results = [];
  try {
    EXERCISE_PATCH_PROBES_.forEach((probe) => {
      results.push(runExercisePatchProbe_(name, probe));
    });
    results.push(runMinimalBodyProbe_(name));
  } finally {
    try {
      deleteDataPointsByName([name]);
      console.info(`probeExercisePatch: deleted ${name}`);
    } catch (err) {
      console.error(
        `probeExercisePatch: could not delete ${name}; delete it in the ` +
          `Health app (orphan reconciliation will also collect it): ${err}`,
      );
    }
  }
  return results;
}
