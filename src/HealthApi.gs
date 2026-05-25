function authHeaders_() {
  return {
    Authorization: 'Bearer ' + getHealthAccessToken_(),
    Accept: 'application/json'
  };
}

function httpJson_(method, url, payload) {
  const options = {
    method: method,
    contentType: 'application/json',
    headers: authHeaders_(),
    muteHttpExceptions: true
  };
  if (payload !== undefined) options.payload = JSON.stringify(payload);

  const maxAttempts = 4;
  let lastErr;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      const body = resp.getContentText();
      if (code >= 200 && code < 300) {
        return body ? JSON.parse(body) : {};
      }
      lastErr = new Error('Health API ' + method + ' ' + url + ' -> ' + code + ': ' + body);
      const transient = code === 429 || (code >= 500 && code < 600);
      if (!transient || attempt === maxAttempts - 1) throw lastErr;
    } catch (err) {
      lastErr = err;
      if (attempt === maxAttempts - 1) throw lastErr;
    }
    const backoffMs = 500 * Math.pow(2, attempt);
    console.warn('Health API ' + method + ' -> failed; retry ' + (attempt + 1) + '/' + (maxAttempts - 1) + ' in ' + backoffMs + 'ms. Error: ' + lastErr);
    Utilities.sleep(backoffMs);
  }
  throw lastErr;
}

// POST/create response is a Long-Running Operation wrapper:
//   { done: true, response: { name: "users/.../dataPoints/<id>", ... } }
// Returns the created datapoint's resource name (or null if absent).
function extractDataPointName_(createResponse) {
  if (!createResponse) return null;
  if (createResponse.response && createResponse.response.name) return createResponse.response.name;
  if (createResponse.name && /\/dataPoints\//.test(createResponse.name)) return createResponse.name;
  return null;
}

// Lists exercise datapoints whose civil start time falls on `date` in the
// script's time zone. Used by findForeignOverlappingExercises to discover
// non-sync-created activities to adopt, and by Debug.gs introspection.
function listExercisesOnDate(date) {
  const startDay = ymd(date);
  const nextDay = ymd(new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1));

  const filter = 'exercise.interval.civil_start_time >= "' + startDay + '"'
    + ' AND exercise.interval.civil_start_time < "' + nextDay + '"';
  const url = HEALTH_API_BASE
    + '/users/me/dataTypes/exercise/dataPoints'
    + '?filter=' + encodeURIComponent(filter)
    + '&pageSize=100';

  const sessions = [];
  let pageToken = null;
  do {
    const pagedUrl = pageToken ? url + '&pageToken=' + encodeURIComponent(pageToken) : url;
    const json = httpJson_('GET', pagedUrl);
    const points = json.dataPoints || [];
    for (const p of points) sessions.push(p);
    pageToken = json.nextPageToken || null;
  } while (pageToken);

  sessions.sort((a, b) => {
    const aStart = a.exercise && a.exercise.interval && a.exercise.interval.startTime;
    const bStart = b.exercise && b.exercise.interval && b.exercise.interval.startTime;
    return new Date(aStart || 0) - new Date(bStart || 0);
  });
  return sessions;
}

function getTzOffsetSeconds_(tz, date) {
  const offsetStr = Utilities.formatDate(date, tz, 'Z');
  const sign = offsetStr.startsWith('-') ? -1 : 1;
  const hours = Number(offsetStr.slice(1, 3));
  const mins = Number(offsetStr.slice(3, 5));
  return sign * (hours * 3600 + mins * 60);
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
  return { utcMs: utcMs, offsetSeconds: offset };
}

function buildIntervalFromUtc_(startUtcMs, startOffsetSeconds, endUtcMs, endOffsetSeconds) {
  return {
    startTime: Utilities.formatDate(new Date(startUtcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    startUtcOffset: startOffsetSeconds + 's',
    endTime: Utilities.formatDate(new Date(endUtcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    endUtcOffset: endOffsetSeconds + 's'
  };
}

function buildSampleTimeFromUtc_(utcMs, offsetSeconds) {
  const tz = Session.getScriptTimeZone();
  const civil = new Date(utcMs);
  const year = Number(Utilities.formatDate(civil, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(civil, tz, 'MM'));
  const day = Number(Utilities.formatDate(civil, tz, 'dd'));
  const hours = Number(Utilities.formatDate(civil, tz, 'HH'));
  const minutes = Number(Utilities.formatDate(civil, tz, 'mm'));
  const seconds = Number(Utilities.formatDate(civil, tz, 'ss'));
  return {
    physicalTime: Utilities.formatDate(civil, 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    utcOffset: offsetSeconds + 's',
    civilTime: {
      date: { year: year, month: month, day: day },
      time: { hours: hours, minutes: minutes, seconds: seconds }
    }
  };
}

// Compute synthetic [startUtcMs, endUtcMs] for a row given its date and
// ordinal among rows on that date. Used as the fallback when a row has no
// edit-derived timing (legacy rows, or rows imported in bulk).
function syntheticExerciseInterval_(date, ordinal) {
  const startHour = SYNTHETIC_START_HOUR + ordinal;
  const endHour = startHour + SYNTHETIC_DURATION_HOURS;
  if (endHour > 24) {
    throw new Error('syntheticExerciseInterval_: ordinal ' + ordinal + ' yields endHour '
      + endHour + ' which spills past midnight (SYNTHETIC_START_HOUR='
      + SYNTHETIC_START_HOUR + ', SYNTHETIC_DURATION_HOURS='
      + SYNTHETIC_DURATION_HOURS + ').');
  }
  const tz = Session.getScriptTimeZone();
  const year = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(date, tz, 'MM'));
  const day = Number(Utilities.formatDate(date, tz, 'dd'));
  const start = localCivilToUtcMs_(tz, year, month, day, startHour, 0);
  const end = localCivilToUtcMs_(tz, year, month, day, endHour, 0);
  return {
    startUtcMs: start.utcMs,
    startOffsetSeconds: start.offsetSeconds,
    endUtcMs: end.utcMs,
    endOffsetSeconds: end.offsetSeconds
  };
}

function syntheticWeightSample_(date) {
  const tz = Session.getScriptTimeZone();
  const year = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(date, tz, 'MM'));
  const day = Number(Utilities.formatDate(date, tz, 'dd'));
  const sample = localCivilToUtcMs_(tz, year, month, day, 12, 0);
  return { utcMs: sample.utcMs, offsetSeconds: sample.offsetSeconds };
}

// Find non-sync-created exercise datapoints whose interval overlaps
// [startMs - FOREIGN_MATCH_BUFFER_MS, endMs + FOREIGN_MATCH_BUFFER_MS] on
// the civil date(s) spanned by [startMs, endMs]. Excludes anything whose
// notes carry SYNC_MARKER (those are ours and are deleted via healthIds).
function findForeignOverlappingExercises(startMs, endMs) {
  const bufferMs = FOREIGN_MATCH_BUFFER_MS;
  const windowStart = startMs - bufferMs;
  const windowEnd = endMs + bufferMs;
  const dates = [new Date(windowStart)];
  const startDay = ymd(new Date(windowStart));
  const endDay = ymd(new Date(windowEnd));
  if (endDay !== startDay) dates.push(new Date(windowEnd));

  const seen = {};
  const candidates = [];
  for (const d of dates) {
    const points = listExercisesOnDate(d);
    for (const p of points) {
      if (!p || !p.name || seen[p.name]) continue;
      seen[p.name] = true;
      const interval = p.exercise && p.exercise.interval;
      if (!interval || !interval.startTime || !interval.endTime) continue;
      const pStartMs = new Date(interval.startTime).getTime();
      const pEndMs = new Date(interval.endTime).getTime();
      if (isNaN(pStartMs) || isNaN(pEndMs)) continue;
      const overlapStart = Math.max(pStartMs, windowStart);
      const overlapEnd = Math.min(pEndMs, windowEnd);
      if (overlapEnd <= overlapStart) continue;
      const notes = (p.exercise && p.exercise.notes) || '';
      if (notes.indexOf(SYNC_MARKER) !== -1) continue;
      candidates.push({
        name: p.name,
        startUtcMs: pStartMs,
        endUtcMs: pEndMs,
        startUtcOffsetSeconds: parseOffsetSeconds_(interval.startUtcOffset),
        endUtcOffsetSeconds: parseOffsetSeconds_(interval.endUtcOffset),
        overlapMs: overlapEnd - overlapStart
      });
    }
  }
  candidates.sort((a, b) => b.overlapMs - a.overlapMs);
  return candidates;
}

function parseOffsetSeconds_(raw) {
  if (!raw) return 0;
  const m = /^(-?\d+)s$/.exec(String(raw));
  if (m) return Number(m[1]);
  const n = Number(raw);
  return isNaN(n) ? 0 : n;
}

// Returns the created datapoint's resource name. Accepts an explicit interval
// in UTC ms plus the offsets that applied at each endpoint.
function createExerciseAt(startUtcMs, startOffsetSeconds, endUtcMs, endOffsetSeconds, notes, displayName) {
  const url = HEALTH_API_BASE + '/users/me/dataTypes/exercise/dataPoints';
  const durationSec = Math.max(0, Math.round((endUtcMs - startUtcMs) / 1000));
  const payload = {
    dataSource: { recordingMethod: 'MANUAL' },
    exercise: {
      interval: buildIntervalFromUtc_(startUtcMs, startOffsetSeconds, endUtcMs, endOffsetSeconds),
      exerciseType: 'STRENGTH_TRAINING',
      displayName: displayName || 'Strength Training',
      notes: notes,
      activeDuration: durationSec + 's',
      metricsSummary: { caloriesKcal: 0 }
    }
  };
  const resp = httpJson_('POST', url, payload);
  return extractDataPointName_(resp);
}

// Returns the created datapoint's resource name.
function createWeightAt(sampleUtcMs, sampleOffsetSeconds, lbs) {
  const grams = Math.round(lbs * GRAMS_PER_LB);
  const url = HEALTH_API_BASE + '/users/me/dataTypes/weight/dataPoints';
  const payload = {
    dataSource: { recordingMethod: 'MANUAL' },
    weight: {
      weightGrams: grams,
      sampleTime: buildSampleTimeFromUtc_(sampleUtcMs, sampleOffsetSeconds)
    }
  };
  const resp = httpJson_('POST', url, payload);
  return extractDataPointName_(resp);
}

// Health API only supports `physical_time` for filtering sample-based data
// points (civil_time.date is rejected). Convert the script-tz day boundaries
// to UTC instants.
function listWeightsOnDate(date) {
  const tz = Session.getScriptTimeZone();
  const year = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(date, tz, 'MM'));
  const day = Number(Utilities.formatDate(date, tz, 'dd'));
  const start = localCivilToUtcMs_(tz, year, month, day, 0, 0);
  const end = localCivilToUtcMs_(tz, year, month, day + 1, 0, 0);
  const startIso = Utilities.formatDate(new Date(start.utcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const endIso = Utilities.formatDate(new Date(end.utcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'");

  const filter = 'weight.sample_time.physical_time >= "' + startIso + '"'
    + ' AND weight.sample_time.physical_time < "' + endIso + '"';
  const url = HEALTH_API_BASE
    + '/users/me/dataTypes/weight/dataPoints'
    + '?filter=' + encodeURIComponent(filter)
    + '&pageSize=100';

  const points = [];
  let pageToken = null;
  do {
    const pagedUrl = pageToken ? url + '&pageToken=' + encodeURIComponent(pageToken) : url;
    const json = httpJson_('GET', pagedUrl);
    (json.dataPoints || []).forEach(p => points.push(p));
    pageToken = json.nextPageToken || null;
  } while (pageToken);
  return points;
}

function getDataPointByName(name) {
  return httpJson_('GET', HEALTH_API_BASE + '/' + name);
}

// Best-effort delete of previously-created datapoints. Groups by data type
// and calls batchDelete per type. Failures (e.g. user already deleted the
// point in the app) are logged but do not throw.
//
// The Health API is picky about the parent/name combination on batchDelete:
//   - URL parent MUST be `users/me/dataTypes/{type}` (literal "me", not the
//     numeric user id).
//   - body names MUST be the canonical numeric form returned by the API.
// Mixing these yields opaque 400 / 500 errors.
function deleteDataPointsByName(names) {
  if (!names || names.length === 0) return;
  const byType = {};
  names.forEach(n => {
    const m = /^users\/[^/]+\/dataTypes\/([^/]+)\/dataPoints\/[^/]+$/.exec(n);
    if (!m) {
      console.warn('deleteDataPointsByName: unparseable name "' + n + '"; skipping.');
      return;
    }
    const dataType = m[1];
    (byType[dataType] = byType[dataType] || []).push(n);
  });
  Object.keys(byType).forEach(dataType => {
    const url = HEALTH_API_BASE + '/users/me/dataTypes/' + dataType + '/dataPoints:batchDelete';
    try {
      httpJson_('POST', url, { names: byType[dataType] });
    } catch (err) {
      console.warn('batchDelete failed for ' + dataType + ': ' + err);
    }
  });
}
