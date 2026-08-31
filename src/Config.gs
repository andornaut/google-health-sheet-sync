const DATE_COLUMN_HEADER = "Date";
const WEIGHT_COLUMN_HEADER = "Weight";
// Each content type has its own synced-at stamp so a row with both can record
// weight progress without losing track that exercise is still pending. Cleared
// independently per phase by onEditMarkDirty, which also triggers an immediate
// sync of the edited row(s).
const EXERCISE_SYNCED_AT_COLUMN_HEADER = "Exercise Synced At";
const WEIGHT_SYNCED_AT_COLUMN_HEADER = "Weight Synced At";
const HEALTH_IDS_COLUMN_HEADER = "Created Health IDs";
// First exercise-relevant edit (sticky). Drives the exercise interval's
// startTime. Only set by exercise-column edits: a weight or Date edit must
// not seed it, otherwise the exercise interval would start before any
// exercise content was typed.
const EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER = "Exercise First Edited At";
// Last edit time for exercise-relevant columns. Drives the exercise
// interval's endTime and the exercise phase's concurrent-edit guard.
// Weight-only edits do NOT advance it: otherwise a bodyweight change
// would drag the exercise endTime forward.
const EXERCISES_LAST_EDITED_AT_COLUMN_HEADER = "Exercises Last Edited At";
// Per-exercise-column edit timestamps for the row, as a JSON object keyed by
// exercise column header: {"Bench press":{"first":"<ISO>","last":"<ISO>"}}.
// The row-level Exercise First/Last Edited At columns say when the row was
// touched; this says when each individual exercise was, which is what lets a
// row's exercises be attributed to the separate app-recorded workout sessions
// they were logged during. `first` is sticky per exercise (a later correction
// does not move that exercise to a different session), `last` advances on every
// edit to that column. Keys are column headers, so renaming an exercise column
// strands its entry; the exercise then falls through to the unattributed group
// rather than being misattributed.
const EXERCISE_EDIT_TIMES_COLUMN_HEADER = "Exercise Edit Times";
// Last edit time for the Weight column. Drives the weight sample time and
// the weight phase's concurrent-edit guard. Advances on every weight edit
// so weight re-edits update the sample time accordingly.
const WEIGHT_EDITED_AT_COLUMN_HEADER = "Weight Edited At";
// Resource name of the foreign STRENGTH_TRAINING datapoint a row's timing was
// aligned to. Recomputed every sync. Also consulted by resolveForeignMatches_
// so a foreign session already aligned to a non-ready row is excluded and
// can't be aligned to a different row in a later incremental run.
const MATCHED_HEALTH_SESSION_COLUMN_HEADER = "Matched Health Session";
const MANAGED_COLUMN_HEADERS = [
  EXERCISE_SYNCED_AT_COLUMN_HEADER,
  WEIGHT_SYNCED_AT_COLUMN_HEADER,
  HEALTH_IDS_COLUMN_HEADER,
  EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER,
  EXERCISES_LAST_EDITED_AT_COLUMN_HEADER,
  WEIGHT_EDITED_AT_COLUMN_HEADER,
  MATCHED_HEALTH_SESSION_COLUMN_HEADER,
  EXERCISE_EDIT_TIMES_COLUMN_HEADER,
];

// How often the polling trigger (flushPending) fires. onEdit syncs the edited
// row(s) immediately; this poll is the retry net for edits that hit lock
// contention or a transient failure, and the pass that re-aligns recent
// unmatched rows once a late foreign session arrives. Apps Script has no
// setTimeout; deferred work happens via periodic triggers.
const POLL_INTERVAL_MIN = 5;

// How long manual sync entry points (Run now, force-resync) wait to acquire
// the script lock before giving up. Automatic triggers pass 0.
const LOCK_WAIT_MS = 30 * 1000;

// Cap rows processed per sync pass to stay under Apps Script's 6-minute
// execution limit. At ~2.83s/row observed (80-row resyncAllRows in 226s),
// 100 rows (~285s) leaves margin under the 360s kill. Remaining dirty rows
// are deferred to the next pass via PENDING_DIRTY_KEY.
const MAX_ROWS_PER_SYNC = 100;

// Script-properties key. Stores a generation marker (Date.now() string)
// rather than a boolean: every dirty-marking call writes a fresh value, so
// syncDirtyRows can detect concurrent edits during a pass by comparing the
// value at end-of-pass with the value at start-of-pass. The flag is cleared
// only at end-of-pass when (a) no work remains AND (b) the generation has
// not advanced. This avoids the early-clear / hard-kill orphan window:
// an Apps Script 6-minute timeout (or any uncaught throw) before the finally
// block runs leaves the flag set, so the next flushPending retries the
// remaining dirty rows instead of stalling until a new edit.
const PENDING_DIRTY_KEY = "pendingDirty";

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
// the same value, so every edit-derived exercise is at least MIN.
//
// MAX does double duty. It bounds the duration on the one path that still
// clamps (a row with no prior datapoint), and more importantly it is the
// stated belief about how long a workout can run: exerciseEditSpansWorkout_
// treats an edit further than MAX from the first one as a later correction
// rather than part of the session, so a row that already has a recorded
// interval keeps it instead of being stretched to the cap. See
// editDerivedDurationMs_ and exerciseEditIsUsable_.
const MIN_EXERCISE_DURATION_MS = 10 * 60 * 1000;
const MAX_EXERCISE_DURATION_MS = 120 * 60 * 1000;

// Backstop (backstop): re-reviews recent exercise rows so a foreign session that
// synced AFTER the row was already pushed can still re-align the row's interval.
// LOOKBACK covers today + yesterday so a late-night workout whose foreign session
// lands the next morning is caught; INTERVAL_HOURS is the time-based trigger
// cadence (it runs off the 5-min poll so foreign re-review doesn't query the
// Health API every 5 minutes; a late foreign session aligns within this interval).
const BACKSTOP_LOOKBACK_DAYS = 2;
const BACKSTOP_INTERVAL_HOURS = 4;

// Orphan reconciliation (reconcileExerciseOrphans_, run from backstop):
// how many civil days back to scan for sync-created exercise datapoints that no
// row's Created Health IDs references. These are leaked by the two accepted
// create-orphan windows (a create POST that succeeds server-side but times out
// client-side and is retried; a 6-minute hard kill landing after the POST
// returns but before the ID is persisted). A wider lookback than the
// foreign-match lookback since an orphan can sit indefinitely and is only ever
// removed here; bounded so each scan stays cheap (one list call per day scanned).
const ORPHAN_RECONCILE_LOOKBACK_DAYS = 7;

// Safety bound on the backstop's cleared-content reconciliation
// (selectStaleDataPointRows_), which DELETES Health datapoints. Clearing cells
// is a per-row, human-scale action: a handful of rows at a time. A stale set
// larger than this is evidence of a systemic change instead (an exercise column
// deleted, a bulk reformat), where deleting would destroy history rather than
// reconcile it. Past the bound the backstop logs an error and reconciles
// nothing, leaving the sheet untouched for a human to look at; "Resync selected
// rows" still works if the large set really was intended.
const STALE_RECONCILE_MAX_ROWS = 10;

// Foreign-session timing alignment (resolveForeignMatches_).
//
// The sync NEVER skips creating its own exercise datapoint: every
// exercise-dirty row writes its own session so the sheet's sets/reps notes
// always reach Health. On a cross-source overlap day the Google Health app
// SHADOWS our session in the card UI: the device's card wins and ours (the
// one carrying the reps/sets) is hidden, and the two are NOT content-merged.
// Writing our datapoint is still worthwhile: the in-app AI assistant reads the
// notes of BOTH overlapping sessions and folds our sets/reps into its summary,
// so our data still reaches the user even when our card itself is shadowed.
// We accept the hidden card rather than offsetting our session to sit beside
// the foreign one, which would unshadow the card at the cost of a start time
// that no longer matches the real workout.
//
// When a row's edit window overlaps a pre-existing foreign STRENGTH_TRAINING
// session (e.g. one started/stopped manually on a watch/Fitbit), we copy that
// foreign session's start/end onto our created datapoint: the manual
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
const FOREIGN_MATCH_BUFFER_MS = 10 * 60 * 1000;

const HEALTH_OAUTH_CLIENT_ID_KEY = "HEALTH_OAUTH_CLIENT_ID";
const HEALTH_OAUTH_CLIENT_SECRET_KEY = "HEALTH_OAUTH_CLIENT_SECRET";
const HEALTH_SERVICE_NAME = "googlehealth";
const HEALTH_OAUTH_SCOPES = [
  "https://www.googleapis.com/auth/googlehealth.activity_and_fitness.readonly",
  "https://www.googleapis.com/auth/googlehealth.activity_and_fitness.writeonly",
  "https://www.googleapis.com/auth/googlehealth.health_metrics_and_measurements.readonly",
  "https://www.googleapis.com/auth/googlehealth.health_metrics_and_measurements.writeonly",
].join(" ");

const HEALTH_API_BASE = "https://health.googleapis.com/v4";

const GRAMS_PER_LB = 453.59237;

// Reject bodyweight values above this as fat-finger typos (e.g. 1850 for
// 185.0). Without the cap the Health API accepts and syncs the bad value.
const MAX_BODYWEIGHT_LB = 499;

// Reject bodyweight values below this as mis-entries: typically a rep count or
// set count accidentally typed into the Weight column (e.g. "5"). A real adult
// bodyweight never lands here, so treating sub-floor values as no-bodyweight
// avoids syncing a 5 lb "weight" datapoint. Symmetric counterpart to
// MAX_BODYWEIGHT_LB.
const MIN_BODYWEIGHT_LB = 50;

// Row-date validation (validateRowDates_, run at the start of every trigger):
// data rows must be in strictly increasing date order down the sheet (which
// also forbids two rows on the same date), and every date's year must fall
// within [MIN_ROW_DATE_YEAR, MAX_ROW_DATE_YEAR]: the year bounds catch
// fat-fingered dates (e.g. 0226-06-12 for 2026-06-12) before a sync writes a
// datapoint at the bogus time. On violation, syncOnEdit throws uncaught so
// Apps Script emails the owner about the failed trigger execution (the edit
// that broke the rule is the moment to alarm); flushPending and backstop log
// an error and skip instead, so a standing violation doesn't email every
// 5-minute poll.
const MIN_ROW_DATE_YEAR = 2025;
const MAX_ROW_DATE_YEAR = 2049;
