// Cached for the duration of one Apps Script execution. The OAuth2 token is
// valid for ~1 hour and Apps Script caps executions at 6 minutes, so the
// cached header cannot outlive the token. Reset between executions because
// each invocation gets a fresh V8 context.
let cachedAuthHeaders_ = null;

function authHeaders_() {
  if (cachedAuthHeaders_) {
    return cachedAuthHeaders_;
  }
  cachedAuthHeaders_ = {
    Accept: "application/json",
    Authorization: `Bearer ${getHealthAccessToken_()}`,
  };
  return cachedAuthHeaders_;
}

function httpJson_(method, url, payload) {
  const options = {
    contentType: "application/json",
    headers: authHeaders_(),
    method,
    muteHttpExceptions: true,
  };
  if (payload !== undefined) {
    options.payload = JSON.stringify(payload);
  }

  // Retry caveat: a create POST that succeeded server-side but timed out
  // client-side (UrlFetchApp.fetch throws -> transient -> retry) will be
  // re-issued, creating a second datapoint whose name we record while the
  // first is orphaned in Health. GET/batchDelete/PATCH are idempotent so
  // they retry safely. The Health API exposes no idempotency key, so this
  // is an accepted (rare) tradeoff rather than something we can guard.
  const maxAttempts = 4;
  let lastErr;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    let resp = null;
    let transient;
    try {
      resp = UrlFetchApp.fetch(url, options);
    } catch (err) {
      lastErr = err;
      transient = true;
    }
    if (resp) {
      const code = resp.getResponseCode();
      const body = resp.getContentText();
      if (code >= 200 && code < 300) {
        return body ? JSON.parse(body) : {};
      }
      lastErr = new Error(`Health API ${method} ${url} -> ${code}: ${body}`);
      lastErr.statusCode = code;
      transient = code === 429 || (code >= 500 && code < 600);
    }
    const isLastAttempt = attempt === maxAttempts - 1;
    const prefix = `Health API ${method} attempt ${attempt + 1}/${maxAttempts} failed`;
    if (!transient || isLastAttempt) {
      if (transient) {
        console.warn(`${prefix}; giving up. Error: ${lastErr}`);
      }
      throw lastErr;
    }
    const backoffMs = 500 * Math.pow(2, attempt);
    console.warn(
      `${prefix}; retrying in ${humanizeMs_(backoffMs)}. Error: ${lastErr}`,
    );
    Utilities.sleep(backoffMs);
  }
  // Unreachable: the last attempt always returns or throws inside the loop,
  // since isLastAttempt short-circuits the transient check. Kept because
  // eslint's consistent-return requires the function to end in a value or a
  // throw. Do not delete it as dead code without also satisfying that rule.
  throw lastErr;
}

// True when a thrown httpJson_ error carries an HTTP 404 (the datapoint is
// gone server-side: e.g. deleted in the Health app). Callers use this to
// recover (recreate weight, treat an exercise delete as already done) instead
// of retrying a GET/DELETE that will never succeed and wedging the row.
function isNotFoundError_(err) {
  return Boolean(err && err.statusCode === 404);
}

// Canonical datapoint resource name, e.g.
//   users/{user}/dataTypes/{type}/dataPoints/{id}
// Capture group 1 is the data type. No /g flag, so .exec is stateless and the
// shared instance is safe to reuse.
const DATAPOINT_NAME_RE_ =
  /^users\/[^/]+\/dataTypes\/([^/]+)\/dataPoints\/[^/]+$/;

// Rewrite a stored resource name's numeric user id to the literal `me` the
// API requires on GET/PATCH/batchDelete URLs.
function toMeName_(name) {
  return String(name).replace(/^users\/[^/]+\//, "users/me/");
}

// POST/create response is a Long-Running Operation wrapper:
//   { done: true, response: { name: "users/.../dataPoints/<id>", ... } }
// Returns the created datapoint's resource name (or null if absent).
function extractDataPointName_(createResponse) {
  if (!createResponse) {
    return null;
  }
  if (createResponse.response && createResponse.response.name) {
    return createResponse.response.name;
  }
  if (createResponse.name && /\/dataPoints\//.test(createResponse.name)) {
    return createResponse.name;
  }
  return null;
}

// GET a single datapoint by its full resource name. Used to preserve the
// prior interval/sampleTime when re-syncing an existing row, so an
// off-date edit (e.g. correcting an old row's notes today) doesn't shift
// the Health datapoint's start/sample time to today. Stored names use the
// numeric user id; we rewrite to `me` for consistency with the rest of
// the client (matches the literal-`me` requirement already documented for
// batchDelete).
function getDataPoint(name) {
  const url = `${HEALTH_API_BASE}/${toMeName_(name)}`;
  return httpJson_("GET", url);
}

// Lists raw dataPoints of `dataType` whose civil time (per the AIP-160 filter
// leaf `filterMember`) falls on `date` in the script's time zone, walking every
// page. Callers apply their own projection/sort. `filterMember` is snake_case
// (e.g. 'exercise.interval.civil_start_time', 'weight.sample_time.civil_time')
// while `dataType` is the path segment in the URL, which is kebab-case: the
// heart-rate type is `heart-rate` in the path but `heart_rate` in a filter.
// Wrong spelling in the path returns 400 INVALID_PARENT_DATA_TYPE_COLLECTION,
// and wrong spelling in the filter returns 400 INVALID_DATA_POINT_FILTER.
function listDataPointsByCivilDate_(dataType, filterMember, date) {
  const startDay = ymd(date);
  const nextDay = ymd(
    new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1),
  );

  const filter = `${filterMember} >= "${startDay}" AND ${filterMember} < "${nextDay}"`;
  const url =
    `${HEALTH_API_BASE}/users/me/dataTypes/${dataType}/dataPoints` +
    `?filter=${encodeURIComponent(filter)}&pageSize=100`;

  const points = [];
  let pageToken = null;
  do {
    const pagedUrl = pageToken
      ? `${url}&pageToken=${encodeURIComponent(pageToken)}`
      : url;
    const json = httpJson_("GET", pagedUrl);
    const page = json.dataPoints || [];
    for (const p of page) {
      points.push(p);
    }
    pageToken = json.nextPageToken || null;
  } while (pageToken);
  return points;
}

// Lists exercise datapoints whose civil start time falls on `date` in the
// script's time zone, sorted ascending by start time. Used by listStrengthOnDate
// to discover non-sync-created activities to match against.
function listExercisesOnDate(date) {
  const sessions = listDataPointsByCivilDate_(
    "exercise",
    "exercise.interval.civil_start_time",
    date,
  );
  sessions.sort((a, b) => {
    const aStart =
      a.exercise && a.exercise.interval && a.exercise.interval.startTime;
    const bStart =
      b.exercise && b.exercise.interval && b.exercise.interval.startTime;
    return new Date(aStart || 0) - new Date(bStart || 0);
  });
  return sessions;
}

// Lists weight datapoints whose civil sample time falls on `date` in the
// script's time zone, returning the minimal shape orphan reconciliation needs:
// `{ name, googleWebClientId }` (null googleWebClientId for device / first-party
// / foreign-app sources, mirroring listStrengthOnDate). Used by
// reconcileWeightOrphans_ to find untracked weight datapoints created by our
// own web client.
//
// Filter member: `weight.sample_time.civil_time` (snake_case in AIP-160 filter
// expressions), matching the documented heart-rate sample filter shape and
// `listExercisesOnDate`'s civil-time convention. Verified against the live API
// 2026-06-09 via a Debug.gs probe: this spelling returns 200 (as does the
// `weight.sample_time.physical_time` RFC3339 form); the bare `weight.sample_time`
// 400s with INVALID_DATA_POINT_FILTER_DATA_TYPE_MEMBER, matching the heart-rate
// "member not supported" behavior. The fail-safe still holds if the server ever
// changes: a 400 makes httpJson_ throw, the caller's per-date try/catch logs a
// warning and reconciles nothing for that day: a safe no-op (never a wrong
// deletion), degrading to "weight orphans not collected" rather than data loss.
function listWeightOnDate(date) {
  const out = [];
  const points = listDataPointsByCivilDate_(
    "weight",
    "weight.sample_time.civil_time",
    date,
  );
  for (const p of points) {
    if (!p || !p.name) {
      continue;
    }
    const app = p.dataSource && p.dataSource.application;
    out.push({
      googleWebClientId: (app && app.googleWebClientId) || null,
      name: p.name,
    });
  }
  return out;
}

function getTzOffsetSeconds_(tz, date) {
  const offsetStr = Utilities.formatDate(date, tz, "Z");
  const sign = offsetStr.startsWith("-") ? -1 : 1;
  const hours = Number(offsetStr.slice(1, 3));
  const mins = Number(offsetStr.slice(3, 5));
  return sign * (hours * 3600 + mins * 60);
}

// One formatDate call instead of six. Caller picks the fields it needs.
function civilDateParts_(tz, date) {
  const parts = Utilities.formatDate(date, tz, "yyyy MM dd HH mm ss").split(
    " ",
  );
  return {
    day: Number(parts[2]),
    hours: Number(parts[3]),
    minutes: Number(parts[4]),
    month: Number(parts[1]),
    seconds: Number(parts[5]),
    year: Number(parts[0]),
  };
}

// Resolve a local civil time (in script tz) to UTC ms and the offset that
// actually applies at that instant. The naive "subtract midnight offset"
// approach is wrong on DST-transition days because the offset at, say, noon
// differs from the offset at midnight. One re-evaluation pass is sufficient
// since tz transitions move the offset by at most a couple of hours.
function localCivilToUtcMs_(tz, year, month, day, hour, minute) {
  const baseMs = Date.UTC(year, month - 1, day, hour, minute || 0, 0);
  let offset = getTzOffsetSeconds_(tz, new Date(baseMs));
  let utcMs = baseMs - offset * 1000;
  const offset2 = getTzOffsetSeconds_(tz, new Date(utcMs));
  if (offset2 !== offset) {
    utcMs = baseMs - offset2 * 1000;
    offset = offset2;
  }
  return { offsetSeconds: offset, utcMs };
}

function buildIntervalFromUtc_(
  startUtcMs,
  startOffsetSeconds,
  endUtcMs,
  endOffsetSeconds,
) {
  return {
    endTime: Utilities.formatDate(
      new Date(endUtcMs),
      "GMT",
      "yyyy-MM-dd'T'HH:mm:ss'Z'",
    ),
    endUtcOffset: `${endOffsetSeconds}s`,
    startTime: Utilities.formatDate(
      new Date(startUtcMs),
      "GMT",
      "yyyy-MM-dd'T'HH:mm:ss'Z'",
    ),
    startUtcOffset: `${startOffsetSeconds}s`,
  };
}

function buildSampleTimeFromUtc_(utcMs, offsetSeconds) {
  const civil = new Date(utcMs);
  const p = civilDateParts_(getTz_(), civil);
  return {
    civilTime: {
      date: { day: p.day, month: p.month, year: p.year },
      time: { hours: p.hours, minutes: p.minutes, seconds: p.seconds },
    },
    physicalTime: Utilities.formatDate(
      civil,
      "GMT",
      "yyyy-MM-dd'T'HH:mm:ss'Z'",
    ),
    utcOffset: `${offsetSeconds}s`,
  };
}

// Compute synthetic [startUtcMs, endUtcMs] for a row given its date and
// ordinal among rows on that date. Used as the fallback when a row has no
// edit-derived timing (legacy rows, or rows imported in bulk).
function syntheticExerciseInterval_(date, ordinal) {
  let startHour = SYNTHETIC_START_HOUR + ordinal;
  let endHour = startHour + SYNTHETIC_DURATION_HOURS;
  if (endHour > 24) {
    // More same-date rows than there are distinct hours left in the day. Rather
    // than error the row (which would re-fail every pass), clamp it into the
    // final slot. Synthetic timing is the bulk-import fallback, so overlapping
    // a few rows at end-of-day is acceptable: the goal is a valid interval.
    console.warn(
      `syntheticExerciseInterval_: ordinal ${ordinal} would spill past ` +
        `midnight; clamping to the final ${SYNTHETIC_DURATION_HOURS}h slot of the day.`,
    );
    endHour = 24;
    startHour = endHour - SYNTHETIC_DURATION_HOURS;
  }
  const tz = getTz_();
  const p = civilDateParts_(tz, date);
  const start = localCivilToUtcMs_(tz, p.year, p.month, p.day, startHour, 0);
  const end = localCivilToUtcMs_(tz, p.year, p.month, p.day, endHour, 0);
  return {
    endOffsetSeconds: end.offsetSeconds,
    endUtcMs: end.utcMs,
    startOffsetSeconds: start.offsetSeconds,
    startUtcMs: start.utcMs,
  };
}

function syntheticWeightSample_(date) {
  const tz = getTz_();
  const p = civilDateParts_(tz, date);
  const sample = localCivilToUtcMs_(tz, p.year, p.month, p.day, 12, 0);
  return { offsetSeconds: sample.offsetSeconds, utcMs: sample.utcMs };
}

// List all Strength Training datapoints whose civil start time falls on
// `date`. Sync-created sessions are NOT filtered out here; the caller
// (resolveForeignMatches_) is responsible for excluding any name it
// already accounts for (sync-created or matched to another row).
// Returned candidates are sorted ascending by startUtcMs.
//
// `googleWebClientId` is the OAuth web-client identifier attributed to the
// datapoint (dataSource.application.googleWebClientId), or null for
// device/first-party/in-app-assistant sessions (whose `application` is null).
// reconcileExerciseOrphans_ uses it to tell our own sync-created datapoints
// apart from foreign ones without trusting a configured client id.
function listStrengthOnDate(date) {
  const points = listExercisesOnDate(date);
  const out = [];
  for (const p of points) {
    if (!p || !p.name) {
      continue;
    }
    const exType = p.exercise && p.exercise.exerciseType;
    if (exType !== "STRENGTH_TRAINING") {
      continue;
    }
    const interval = p.exercise && p.exercise.interval;
    if (!interval || !interval.startTime || !interval.endTime) {
      continue;
    }
    const pStartMs = new Date(interval.startTime).getTime();
    const pEndMs = new Date(interval.endTime).getTime();
    if (isNaN(pStartMs) || isNaN(pEndMs)) {
      continue;
    }
    const app = p.dataSource && p.dataSource.application;
    out.push({
      endUtcMs: pEndMs,
      endUtcOffsetSeconds: parseOffsetSeconds_(interval.endUtcOffset),
      googleWebClientId: (app && app.googleWebClientId) || null,
      name: p.name,
      startUtcMs: pStartMs,
      startUtcOffsetSeconds: parseOffsetSeconds_(interval.startUtcOffset),
    });
  }
  // Already sorted by startUtcMs because listExercisesOnDate sorts by the
  // same key (interval.startTime).
  return out;
}

// Sort Health datapoint resource names into per-data-type buckets so the
// weight and exercise sync phases can manage their IDs independently.
// Any name that doesn't match a known type goes into `other` and is left
// untouched by both phases.
function splitHealthIdsByType_(names) {
  const out = { exercise: [], other: [], weight: [] };
  (names || []).forEach((n) => {
    const m = DATAPOINT_NAME_RE_.exec(n);
    if (!m) {
      out.other.push(n);
      return;
    }
    if (m[1] === "weight") {
      out.weight.push(n);
    } else if (m[1] === "exercise") {
      out.exercise.push(n);
    } else {
      out.other.push(n);
    }
  });
  return out;
}

function parseOffsetSeconds_(raw) {
  const m = /^(-?\d+)s$/.exec(String(raw));
  if (m) {
    return Number(m[1]);
  }
  const n = Number(raw);
  return isNaN(n) ? 0 : n;
}

// Returns the created datapoint's resource name. Accepts an explicit interval
// in UTC ms plus the offsets that applied at each endpoint.
function createExerciseAt(
  startUtcMs,
  startOffsetSeconds,
  endUtcMs,
  endOffsetSeconds,
  notes,
) {
  const url = `${HEALTH_API_BASE}/users/me/dataTypes/exercise/dataPoints`;
  const payload = {
    dataSource: { recordingMethod: "MANUAL" },
    exercise: {
      // No activeDuration: the server computes it from the interval and
      // ignores the client either way. Measured against the live API on
      // 2026-08-19: a create sending 137s for a 600s interval reads back 600s,
      // a create omitting the field reads back 600s just the same, and a PATCH
      // that lengthens the interval moves it on its own. The value we used to
      // send was the interval's length, so this changes nothing on the card.
      //
      // No displayName: for every exerciseType except OTHER the server ignores
      // whatever the client sends and derives the card's title from
      // exerciseType. Confirmed by the Health API team in 2026-08 and measured
      // on 2026-08-19: a datapoint created without one reads back as
      // "Strength training", which is not even the "Strength Training" this
      // client used to send. exerciseType is the single lever on the title.
      exerciseType: "STRENGTH_TRAINING",
      interval: buildIntervalFromUtc_(
        startUtcMs,
        startOffsetSeconds,
        endUtcMs,
        endOffsetSeconds,
      ),
      notes,
    },
  };
  const resp = httpJson_("POST", url, payload);
  const name = extractDataPointName_(resp);
  if (!name) {
    // POST returned 2xx but no parseable resource name. We can't track an
    // untracked datapoint (no ID to delete/re-sync later), so treat it as a
    // failed create: throw so the caller retries instead of silently stamping
    // the row synced and orphaning whatever may have been created server-side.
    throw new Error(
      `createExerciseAt: create returned no datapoint name: ${JSON.stringify(resp)}`,
    );
  }
  return name;
}

// Returns the created datapoint's resource name.
function createWeightAt(sampleUtcMs, sampleOffsetSeconds, lbs) {
  const grams = Math.round(lbs * GRAMS_PER_LB);
  const url = `${HEALTH_API_BASE}/users/me/dataTypes/weight/dataPoints`;
  const payload = {
    dataSource: { recordingMethod: "MANUAL" },
    weight: {
      sampleTime: buildSampleTimeFromUtc_(sampleUtcMs, sampleOffsetSeconds),
      weightGrams: grams,
    },
  };
  const resp = httpJson_("POST", url, payload);
  const name = extractDataPointName_(resp);
  if (!name) {
    // See createExerciseAt: a create with no parseable resource name is a
    // failed create, not a no-op success. Throw so the row retries.
    throw new Error(
      `createWeightAt: create returned no datapoint name: ${JSON.stringify(resp)}`,
    );
  }
  return name;
}

// Update an existing weight datapoint's weightGrams in place. The body MUST
// include sampleTime: empirically, any PATCH body without it returns 500
// INTERNAL (matching the documented "minimal body 500s" pattern observed
// for exercise PATCH). `name` is omitted from the body since the URL
// already identifies the resource (AIP-134); the empirical probe confirms
// the server accepts a body without `name`. The caller passes the prior
// datapoint's sampleTime verbatim from a GET so the sample timestamp is
// echoed back unchanged. createTime, dataSource, and the resource name are
// preserved server-side.
function patchWeight(name, sampleTime, lbs) {
  const url = `${HEALTH_API_BASE}/${toMeName_(name)}`;
  const grams = Math.round(lbs * GRAMS_PER_LB);
  const payload = {
    weight: { sampleTime, weightGrams: grams },
  };
  httpJson_("PATCH", url, payload);
}

// Delete previously-created datapoints. Groups by data type and calls
// batchDelete per type. Throws on API failure so the caller can keep the
// IDs in the sheet and retry next sync (otherwise we orphan datapoints in
// Health that the script no longer tracks).
//
// The Health API is picky about the parent/name combination on batchDelete:
//   - URL parent MUST be `users/me/dataTypes/{type}` (literal "me", not the
//     numeric user id).
//   - body names MUST be the canonical numeric form returned by the API.
// Mixing these yields opaque 400 / 500 errors.
function deleteDataPointsByName(names) {
  if (!names || names.length === 0) {
    return;
  }
  const byType = {};
  names.forEach((n) => {
    const m = DATAPOINT_NAME_RE_.exec(n);
    if (!m) {
      console.warn(
        `deleteDataPointsByName: unparseable name "${n}"; skipping.`,
      );
      return;
    }
    const dataType = m[1];
    (byType[dataType] = byType[dataType] || []).push(n);
  });
  Object.keys(byType).forEach((dataType) => {
    const url = `${HEALTH_API_BASE}/users/me/dataTypes/${dataType}/dataPoints:batchDelete`;
    httpJson_("POST", url, { names: byType[dataType] });
  });
}
