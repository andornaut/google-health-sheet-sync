const SHEET_NAME = 'Sheet1';

const DATE_COLUMN_HEADER = 'Date';
const WEIGHT_COLUMN_HEADER = 'Weight';
const SYNCED_AT_COLUMN_HEADER = 'Synced At';
const HEALTH_IDS_COLUMN_HEADER = 'Health IDs';
const MANAGED_COLUMN_HEADERS = [SYNCED_AT_COLUMN_HEADER, HEALTH_IDS_COLUMN_HEADER];

const DEBOUNCE_MS = 60 * 1000;
const DEBOUNCE_CHECK_INTERVAL_MIN = 1;
const BACKSTOP_INTERVAL_HOURS = 1;

const SYNTHETIC_START_HOUR = 12;
const SYNTHETIC_DURATION_HOURS = 1;

// Toggle which data types this script writes to the Google Health API.
// Both default to true; flip off to debug or to disable a category temporarily.
const SYNC_EXERCISES = true;
const SYNC_WEIGHT = true;

const LAST_EDIT_MS_KEY = 'lastEditMs';

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
