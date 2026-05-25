const SHEET_NAME = 'Sheet1';

const DATE_COLUMN_HEADER = 'Date';
const WEIGHT_COLUMN_HEADER = 'Weight';
const SYNCED_AT_COLUMN_HEADER = 'Synced At';
const HEALTH_IDS_COLUMN_HEADER = 'Health IDs';
const FIRST_EDITED_AT_COLUMN_HEADER = 'First Edited At';
const LAST_EDITED_AT_COLUMN_HEADER = 'Last Edited At';
const MANAGED_COLUMN_HEADERS = [
  SYNCED_AT_COLUMN_HEADER,
  HEALTH_IDS_COLUMN_HEADER,
  FIRST_EDITED_AT_COLUMN_HEADER,
  LAST_EDITED_AT_COLUMN_HEADER
];

const DEBOUNCE_CHECK_INTERVAL_MIN = 1;
const BACKSTOP_INTERVAL_HOURS = 1;

// Synthetic timing is the fallback when a row has no First/Last Edited At
// timestamps (e.g. rows that pre-date this feature, or rows imported in bulk).
// Each row on a given date gets startHour = SYNTHETIC_START_HOUR + ordinal,
// endHour = startHour + SYNTHETIC_DURATION_HOURS, so the second strength row
// on the same date starts an hour after the first, and so on.
const SYNTHETIC_START_HOUR = 12;
const SYNTHETIC_DURATION_HOURS = 1;

// Per-row quiet period. A dirty row is held back from syncing until this many
// ms have passed since its Last Edited At. The goal: stamp the activity's
// end time with roughly when the workout actually finished, not when the
// first sync trigger happened to fire. If you re-edit a synced row, this
// re-applies (the row goes dirty again and waits another quiesce window).
// Rows with no Last Edited At (legacy/backfill) bypass the wait and sync
// immediately.
const LAST_EDIT_QUIESCE_MS = 60 * 60 * 1000;

// When matching foreign Google Health activities (ones this script didn't
// create) to a row's edit window, treat datapoints whose interval overlaps
// [firstEdit - buffer, lastEdit + buffer] as candidates. Picks up watch- or
// app-logged workouts that are slightly offset from the spreadsheet edit
// window so we can adopt their real start/end times instead of inventing
// our own.
const FOREIGN_MATCH_BUFFER_MS = 30 * 60 * 1000;

// Toggle which data types this script writes to the Google Health API.
// Both default to true; flip off to debug or to disable a category temporarily.
const SYNC_EXERCISES = true;
const SYNC_WEIGHT = true;

const HEALTH_OAUTH_CLIENT_ID_KEY = 'HEALTH_OAUTH_CLIENT_ID';
const HEALTH_OAUTH_CLIENT_SECRET_KEY = 'HEALTH_OAUTH_CLIENT_SECRET';
const HEALTH_SERVICE_NAME = 'googlehealth';
const HEALTH_OAUTH_SCOPES = [
  'https://www.googleapis.com/auth/googlehealth.activity_and_fitness',
  'https://www.googleapis.com/auth/googlehealth.health_metrics_and_measurements'
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
