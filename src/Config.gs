const DATE_COLUMN_HEADER = 'Date';
const WEIGHT_COLUMN_HEADER = 'Weight';
// Weight syncs immediately on edit (no debounce); exercises wait for the
// edit-burst debounce window. Each content type has its own stamp so a row
// with both can record weight progress without losing track that exercise
// is still pending. Cleared independently per phase by onEditMarkDirty.
const EXERCISE_SYNCED_AT_COLUMN_HEADER = 'Exercise Synced At';
const WEIGHT_SYNCED_AT_COLUMN_HEADER = 'Weight Synced At';
const HEALTH_IDS_COLUMN_HEADER = 'Created Health IDs';
// First exercise-relevant edit (sticky). Drives the exercise interval's
// startTime. Only set by exercise-column edits — a weight or Date edit must
// not seed it, otherwise the exercise interval would start before any
// exercise content was typed.
const EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER = 'Exercise First Edited At';
// Last edit time for exercise-relevant columns. Drives the exercise
// interval's endTime and the exercise phase's concurrent-edit guard.
// Weight-only edits do NOT advance it — otherwise a bodyweight change
// would drag the exercise endTime forward.
const EXERCISES_LAST_EDITED_AT_COLUMN_HEADER = 'Exercises Last Edited At';
// Last edit time for the Weight column. Drives the weight sample time and
// the weight phase's concurrent-edit guard. Advances on every weight edit
// so weight re-edits update the sample time accordingly.
const WEIGHT_EDITED_AT_COLUMN_HEADER = 'Weight Edited At';
// Resource name of the foreign STRENGTH_TRAINING datapoint a row's timing was
// aligned to. Recomputed every sync. Also consulted by resolveForeignMatches_
// so a foreign session already aligned to a non-ready row is excluded and
// can't be aligned to a different row in a later incremental run.
const MATCHED_HEALTH_SESSION_COLUMN_HEADER = 'Matched Health Session';
const MANAGED_COLUMN_HEADERS = [
  EXERCISE_SYNCED_AT_COLUMN_HEADER,
  WEIGHT_SYNCED_AT_COLUMN_HEADER,
  HEALTH_IDS_COLUMN_HEADER,
  EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER,
  EXERCISES_LAST_EDITED_AT_COLUMN_HEADER,
  WEIGHT_EDITED_AT_COLUMN_HEADER,
  MATCHED_HEALTH_SESSION_COLUMN_HEADER
];

// How often the polling trigger fires to check for ready-to-sync rows.
// Apps Script has no setTimeout; deferred work happens via periodic triggers.
const POLL_INTERVAL_MIN = 5;

// How long manual sync entry points (Run now, force-resync) wait to acquire
// the script lock before giving up. Automatic triggers pass 0.
const LOCK_WAIT_MS = 30 * 1000;

// Cap rows processed per sync pass to stay under Apps Script's 6-minute
// execution limit. At ~3.5s/row observed, 90 rows (~315s) leaves margin under
// the 360s kill. Remaining dirty rows are deferred to the next pass via
// PENDING_DIRTY_KEY.
const MAX_ROWS_PER_SYNC = 90;

// Script-properties key. Stores a generation marker (Date.now() string)
// rather than a boolean: every dirty-marking call writes a fresh value, so
// syncDirtyRows can detect concurrent edits during a pass by comparing the
// value at end-of-pass with the value at start-of-pass. The flag is cleared
// only at end-of-pass when (a) no work remains AND (b) the generation has
// not advanced. This avoids the early-clear / hard-kill orphan window:
// an Apps Script 6-minute timeout (or any uncaught throw) before the finally
// block runs leaves the flag set, so the next flushIfPending retries the
// remaining dirty rows instead of stalling until a new edit.
const PENDING_DIRTY_KEY = 'pendingDirty';

// Synthetic timing is the fallback when a row has no Exercise First Edited
// At / Exercises Last Edited At / Weight Edited At timestamps (e.g. rows
// imported in bulk). Each row on a given date gets
// startHour = SYNTHETIC_START_HOUR + ordinal, endHour = startHour +
// SYNTHETIC_DURATION_HOURS, so the second strength row on the same date
// starts an hour after the first, and so on.
const SYNTHETIC_START_HOUR = 12;
const SYNTHETIC_DURATION_HOURS = 1;

// Bounds for edit-derived exercise interval duration. The MIN floor doubles as
// the start-only default: a single-edit row (start == last edit, no observed
// end) has a zero/negative raw span that floors to MIN, and two near-instant
// edits (which would otherwise produce a span the Health API 500s on) floor to
// the same value, so every edit-derived exercise is at least MIN. The MAX cap
// keeps a row edited across days (late corrections) from recording a multi-hour
// "workout". See editDerivedDurationMs_.
const MIN_EXERCISE_DURATION_MS = 10 * 60 * 1000;
const MAX_EXERCISE_DURATION_MS = 120 * 60 * 1000;

// Daily backstop (dailyBackstop): re-reviews recent exercise rows so a foreign
// session that synced AFTER the row was already pushed can still re-align the
// row's interval. LOOKBACK covers today + yesterday so a late-night workout
// whose foreign session lands the next morning is caught; HOUR is when the
// daily time-based trigger fires.
const BACKSTOP_LOOKBACK_DAYS = 2;
const BACKSTOP_HOUR = 4;

// Per-row "still typing" guard. A dirty row whose Exercises Last Edited At
// is within this window is skipped on the current poll and picked up by the
// next one. Without it, a poll firing mid-edit would push out a half-typed
// row. Since every sync delete+recreates with the row's current state,
// syncing more often than necessary just churns the Health datapoint
// without changing the eventual state — the guard avoids that churn during
// an edit burst. Rows with no Exercises Last Edited At bypass the wait and
// sync immediately. Weight phase has no debounce.
const LAST_EDIT_QUIESCE_MS = 60 * 1000;

// Foreign-session timing alignment (resolveForeignMatches_).
//
// The sync NEVER skips creating its own exercise datapoint: every
// exercise-dirty row writes its own session so the sheet's sets/reps notes
// always reach Health. This is safe because the Google Health app merges
// overlapping same-type sessions into one summary card server-side (confirmed
// 2026-06-02 via the in-app AI assistant: one card, no visible duplicate), so
// a device-logged session and our sync-created one coexist without a visible
// duplicate.
//
// When a row's edit window overlaps a pre-existing foreign STRENGTH_TRAINING
// session (e.g. one started/stopped manually on a watch/Fitbit), we copy that
// foreign session's start/end onto our created datapoint — the manual
// start/stop is more accurate than our edit-derived window. A foreign session
// is a candidate when its interval overlaps [firstEdit - buffer, lastEdit +
// buffer]; the largest-overlap candidate wins. Overlap is computed in absolute
// UTC, so a workout that crosses local midnight still matches a candidate
// logged on the adjacent civil date. Rows without same-date exercise edit
// timestamps get no alignment and fall through to synthetic/prior timing.
//
// Caveat: two overlapping STRENGTH_TRAINING datapoints (ours + the device's)
// may both feed daily aggregates (:dailyRollUp, active minutes) even though
// they display as one card.
const FOREIGN_MATCH_BUFFER_MS = 30 * 60 * 1000;

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

const GRAMS_PER_LB = 453.59237;

// Reject bodyweight values above this as fat-finger typos (e.g. 1850 for
// 185.0). Without the cap the Health API accepts and syncs the bad value.
const MAX_BODYWEIGHT_LB = 999;
