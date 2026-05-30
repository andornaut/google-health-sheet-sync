const DATE_COLUMN_HEADER = 'Date';
const WEIGHT_COLUMN_HEADER = 'Weight';
// Weight syncs immediately on edit (no quiesce), exercises wait for the
// quiesce window. Each content type has its own stamp so a row with both can
// record weight progress without losing track that exercise is still pending.
// Both cleared on edit.
const EXERCISE_SYNCED_AT_COLUMN_HEADER = 'Exercise Synced At';
const WEIGHT_SYNCED_AT_COLUMN_HEADER = 'Weight Synced At';
const HEALTH_IDS_COLUMN_HEADER = 'Created Health IDs';
const FIRST_EDITED_AT_COLUMN_HEADER = 'First Edited At';
const LAST_EDITED_AT_COLUMN_HEADER = 'Last Edited At';
// Resource name of the foreign STRENGTH_TRAINING datapoint a row was matched
// to. Recomputed every sync. Also consulted by resolveForeignMatches_ to
// avoid two sheet rows claiming the same foreign session across incremental
// runs (the dirty row excludes candidates already held by synced rows).
const MATCHED_HEALTH_SESSION_COLUMN_HEADER = 'Matched Health Session';
const MANAGED_COLUMN_HEADERS = [
  EXERCISE_SYNCED_AT_COLUMN_HEADER,
  WEIGHT_SYNCED_AT_COLUMN_HEADER,
  HEALTH_IDS_COLUMN_HEADER,
  FIRST_EDITED_AT_COLUMN_HEADER,
  LAST_EDITED_AT_COLUMN_HEADER,
  MATCHED_HEALTH_SESSION_COLUMN_HEADER
];

// How often the polling trigger fires to check for ready-to-sync rows.
// Apps Script has no setTimeout; deferred work happens via periodic triggers.
const POLL_INTERVAL_MIN = 5;
const BACKSTOP_INTERVAL_HOURS = 1;

// How long manual sync entry points (Run now, force-resync) wait to acquire
// the script lock before giving up. Automatic triggers pass 0.
const LOCK_WAIT_MS = 30 * 1000;

// Cap rows processed per sync pass to stay under Apps Script's 6-minute
// execution limit. At ~3.3s/row observed, 75 rows leaves a comfortable margin.
// Remaining dirty rows are deferred to the next pass via PENDING_DIRTY_KEY.
const MAX_ROWS_PER_SYNC = 75;

// Script-properties key. Set to '1' whenever a non-managed edit happens or a
// force-resync path clears the synced-at stamps; cleared by syncDirtyRows
// when it finds no dirty rows. flushIfPending short-circuits when this is
// unset so most poll ticks are just a property read.
const PENDING_DIRTY_KEY = 'pendingDirty';

// Synthetic timing is the fallback when a row has no First/Last Edited At
// timestamps (e.g. rows that pre-date this feature, or rows imported in bulk).
// Each row on a given date gets startHour = SYNTHETIC_START_HOUR + ordinal,
// endHour = startHour + SYNTHETIC_DURATION_HOURS, so the second strength row
// on the same date starts an hour after the first, and so on.
const SYNTHETIC_START_HOUR = 12;
const SYNTHETIC_DURATION_HOURS = 1;

// Bounds for edit-derived exercise interval duration. A row edited just once
// produces a zero-span interval that the Health API 500s on; a row edited
// across days (or weeks) produces a multi-hour "workout" that's wrong on the
// other end. Clamp to [min, max] so the duration stays plausible.
const MIN_EXERCISE_DURATION_MS = 20 * 60 * 1000;
const MAX_EXERCISE_DURATION_MS = 120 * 60 * 1000;

// Per-row quiet period. A dirty row is held back from syncing until this many
// ms have passed since its Last Edited At. The goal: stamp the activity's
// end time with roughly when the workout actually finished, not when the
// first sync trigger happened to fire. If you re-edit a synced row, this
// re-applies (the row goes dirty again and waits another quiesce window).
// Rows with no Last Edited At (legacy/backfill) bypass the wait and sync
// immediately.
const LAST_EDIT_QUIESCE_MS = 45 * 60 * 1000;

// When matching foreign Google Health activities (ones this script didn't
// create) to a row's edit window, treat STRENGTH_TRAINING datapoints whose
// interval overlaps [firstEdit - buffer, lastEdit + buffer] as candidates.
// Picks up watch- or app-logged workouts whose recorded start/end is slightly
// offset from the spreadsheet edit window so we can recognize them as the
// same workout and skip writing our own duplicate exercise datapoint.
const FOREIGN_MATCH_BUFFER_MS = 30 * 60 * 1000;

// Toggle which data types this script writes to the Google Health API.
// Both default to true; flip off to debug or to disable a category temporarily.
const SYNC_EXERCISES = true;
const SYNC_WEIGHT = true;

const HEALTH_OAUTH_CLIENT_ID_KEY = 'HEALTH_OAUTH_CLIENT_ID';
const HEALTH_OAUTH_CLIENT_SECRET_KEY = 'HEALTH_OAUTH_CLIENT_SECRET';
const HEALTH_SERVICE_NAME = 'googlehealth';
const HEALTH_OAUTH_SCOPES = [
  'https://www.googleapis.com/auth/googlehealth.activity_and_fitness.readonly',
  'https://www.googleapis.com/auth/googlehealth.activity_and_fitness.writeonly',
  'https://www.googleapis.com/auth/googlehealth.health_metrics_and_measurements.readonly',
  'https://www.googleapis.com/auth/googlehealth.health_metrics_and_measurements.writeonly'
].join(' ');

const HEALTH_API_BASE = 'https://health.googleapis.com/v4';
const SYNC_MARKER = '[gs-sync]';

const EXERCISE_ABBREVIATIONS = {
  'Barbell curl': 'BC',
  'Barbell triceps extension': 'BTE',
  'Bench press': 'BP',
  'Deadlift': 'DL',
  'Dumbell curl': 'DC',
  'Dumbell rows': 'DR',
  'Landmine rows': 'LR',
  'Lateral raises': 'LAT',
  'Shoulder press': 'SP',
  'Squat': 'SQ'
};

const GRAMS_PER_LB = 453.59237;
