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

// Listing kept for diagnostic / future-correlation use, even though the
// sync path no longer needs it.
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

function buildSampleTime_(date, hour) {
  const tz = Session.getScriptTimeZone();
  const year = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(date, tz, 'MM'));
  const day = Number(Utilities.formatDate(date, tz, 'dd'));
  const offsetSeconds = getTzOffsetSeconds_(tz, date);
  const sampleUtcMs = Date.UTC(year, month - 1, day, hour, 0, 0) - (offsetSeconds * 1000);
  const physicalTime = Utilities.formatDate(new Date(sampleUtcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  return {
    physicalTime: physicalTime,
    utcOffset: offsetSeconds + 's',
    civilTime: {
      date: {
        year: year,
        month: month,
        day: day
      },
      time: { hours: hour }
    }
  };
}

function buildInterval_(date, startHour, endHour) {
  const tz = Session.getScriptTimeZone();
  const year = Number(Utilities.formatDate(date, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(date, tz, 'MM'));
  const day = Number(Utilities.formatDate(date, tz, 'dd'));
  const offsetSeconds = getTzOffsetSeconds_(tz, date);
  const startUtcMs = Date.UTC(year, month - 1, day, startHour, 0, 0) - (offsetSeconds * 1000);
  const endUtcMs = Date.UTC(year, month - 1, day, endHour, 0, 0) - (offsetSeconds * 1000);
  return {
    startTime: Utilities.formatDate(new Date(startUtcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    startUtcOffset: offsetSeconds + 's',
    endTime: Utilities.formatDate(new Date(endUtcMs), 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    endUtcOffset: offsetSeconds + 's'
  };
}

// Returns the created datapoint's resource name.
function createExercise(date, ordinal, notes, displayName) {
  const url = HEALTH_API_BASE + '/users/me/dataTypes/exercise/dataPoints';
  const startHour = SYNTHETIC_START_HOUR + ordinal;
  const endHour = startHour + SYNTHETIC_DURATION_HOURS;
  const durationSec = SYNTHETIC_DURATION_HOURS * 3600;
  const payload = {
    dataSource: { recordingMethod: 'MANUAL' },
    exercise: {
      interval: buildInterval_(date, startHour, endHour),
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
function createWeight(date, lbs) {
  const grams = Math.round(lbs * GRAMS_PER_LB);
  const url = HEALTH_API_BASE + '/users/me/dataTypes/weight/dataPoints';
  const payload = {
    dataSource: { recordingMethod: 'MANUAL' },
    weight: {
      weightGrams: grams,
      sampleTime: buildSampleTime_(date, 12)
    }
  };
  const resp = httpJson_('POST', url, payload);
  return extractDataPointName_(resp);
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
