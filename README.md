# Google Health Sheet Sync

Google Apps Script to sync strength-training and bodyweight data from Google Sheets to the **Google Health app** via the **Google Health API v4**.

## Features

- **Strength Exercises**: Parses lifts (e.g., `135x5x3`, `*135x5x3` for assisted) and logs them as `STRENGTH_TRAINING` sessions.
- **Bodyweight**: Logs weight data points.
- **Delete+recreate sync (idempotent)**: Each sync pass reconciles the row's Health datapoint(s) with current cell content. Editing a synced row clears the `Synced At` stamps and the row is reprocessed, refreshing `endTime`, `activeDuration`, and notes from the current state. The exercise datapoint is delete+recreated **only when the interval or notes actually changed** — an unchanged re-sync keeps the existing datapoint (and its resource name) and costs just a read, so the periodic re-checks don't churn. (The Google Health API has no functional update mechanism for exercise datapoints — `PATCH` is a server-side no-op and `PUT`/POST-to-resource don't exist; delete+POST is the only update path, just skipped when nothing changed.)
- **Edit-derived timing, date-aware**: The activity's start/end times come from when you first/last edited the row, so the Google Health session reflects when you actually did the workout. A single edit (start only — e.g. you've typed the first set but not finished the workout) records a default-length session; a second edit records the real end time, so `endTime` grows as you log more sets during a live workout (clamped to `[MIN, MAX]`, default 10 min – 2 h). Edit timestamps are only trusted when their civil date matches the row's `Date` column — editing an old row today won't shift its Health datapoint to today. On re-sync, the prior datapoint's interval/sample time is fetched and preserved (the content updates but the timing stays put). Rows without edit timestamps (e.g. backfill) fall back to a synthetic noon-ordinal slot on the row's `Date`.
- **Foreign-activity matching**: If a Strength Training session you logged on a watch or in another app overlaps the same workout (the row's edit timestamps fall on its `Date` and overlap the foreign session's interval, with 30 minutes of slack), the script borrows that session's start/end times for the datapoint it writes — the manual start/stop is more accurate than the edit-derived window — and records the matched session in the row's `Matched Health Session` column. The script **always writes its own exercise datapoint** (carrying the reps/sets notes); the foreign datapoint is left untouched, so its calories, heart rate, and recording method are preserved. (The API enforces this anyway — DELETE/PATCH on a non-script-owned datapoint returns `403 DATA_POINT_NOT_OWNED_BY_CLIENT`.) The 5-minute poll re-reviews recent **unmatched** rows so a foreign session that synced *after* the row was already pushed can re-align it within minutes, and a daily backstop re-reviews recent **matched** rows so a foreign session whose interval was extended after your last edit gets re-borrowed. Bodyweight always syncs as normal.
- **Two-phase sync**: Weight and exercise sync independently with their own `Synced At` stamps and column-aware dirty tracking. Editing only the Weight column re-pushes just the weight datapoint and leaves the exercise interval untouched; editing only an exercise column re-pushes just the exercise datapoint. Editing the Date column alone on an otherwise-empty row is a no-op (nothing to sync until exercise or weight content lands).
- **Sync on edit**: Editing a row syncs it **immediately** (the `onEdit` trigger marks the row dirty and pushes it under a non-blocking script lock). During a fast edit burst the lock batches the work — an edit whose trigger can't grab the lock is caught by the next poll — so you get a few syncs per burst, not one per cell.
- **Automated retry net**: A 5-minute poll re-syncs anything an `onEdit` left dirty (lock contention or a transient failure) and re-reviews recent unmatched rows for late foreign sessions. A daily backstop re-reviews recent matched rows (for foreign-interval changes) and deletes any sync-created datapoints that an interrupted run orphaned (no row references them).
- **Per-pass cap**: Each sync processes up to 100 rows (newest dates first) to stay under Apps Script's 6-minute execution limit. Remaining rows are picked up on the next poll.

> [!NOTE]
> The Google Health API enforces strict ownership. Datapoints logged by this script live side-by-side with sessions recorded by your watch (e.g., Pixel Watch, Fitbit).

---

## Spreadsheet Layout

The script always operates on the first tab of the spreadsheet (leftmost; name doesn't matter). It is selected purely by position, so reordering tabs to move a different sheet into the leftmost slot silently redirects the sync to that sheet — keep your data tab leftmost. Columns are auto-detected by header name and position doesn't matter — any header that isn't `Date`, `Weight`, or a managed column is treated as an exercise. Recommended layout:

- **`Date`** (leftmost)
- **Exercise Columns** (middle; headers are used as exercise names)
- **`Weight`** (rightmost; bodyweight in lb)

### Grammar for Exercise Cells

One entry per line (newline, comma, or semicolon separated):

| Cell | Meaning |
| ------ | --------- |
| `135` | 135 lb (reps and sets unspecified) |
| `135x5` | 5 reps at 135 lb (sets unspecified) |
| `135x5x3` | 5 reps × 3 sets at 135 lb |
| `135x5x0` | Start marker — 0 sets done yet; retained to anchor the workout's start time but not shown in notes |
| `*135x5x3` | Assisted reps |
| `135x5x3, 145x3x2` | Multiple distinct sets/exercises in one cell |

Unspecified fields are omitted from the Google Health notes. For example, a `Bench press` cell containing `135` renders as `Bench press, 135 lbs`; `135x5` renders as `Bench press, 135 lbs, 5 reps`; `135x5x3` renders as `Bench press, 135 lbs, 3 sets of 5`.

A zero-set entry (`x0`, e.g. `135x5x0`) is **retained but suppressed** from the notes: it marks the start of an exercise you haven't completed yet, so the row anchors its start time and can match a foreign session, but no datapoint is written until you log a real set (and a row whose only entries are zero-set never produces a datapoint). A zero-rep entry (`x0` in the reps position, e.g. `135x0`) is dropped entirely as "not performed".

### Example

Trimmed to a few exercise columns and the most relevant managed columns. Health resource names abbreviated as `ex/NNN` / `wt/NNN`; the remaining managed columns `Exercise First Edited At`, `Exercises Last Edited At`, and `Weight Edited At` are omitted here. `Exercise Synced At` and `Weight Synced At` are stamped independently per phase — a row with only bodyweight leaves `Exercise Synced At` blank, and vice versa:

| Date         | Bench press | Deadlift         | Weight | Exercise Synced At   | Weight Synced At     | Created Health IDs | Matched Health Session |
| ------------ | ----------- | ---------------- | ------ | -------------------- | -------------------- | ------------------ | ---------------------- |
| Jan 2, 2026  | 210         |                  | 190.0  | 2026-01-02T18:30:00Z | 2026-01-02T17:45:00Z | [ex/001, wt/002]   |                        |
| Jan 15, 2026 | 215x4       |                  |        | 2026-01-15T20:00:00Z |                      | [ex/004]           |                        |
| Jan 18, 2026 |             | 295x4x6          | 187.5  | 2026-01-18T17:45:00Z | 2026-01-18T17:00:00Z | [ex/005, wt/006]   |                        |
| Jan 24, 2026 | *225        |                  |        | 2026-01-24T18:00:00Z |                      | [ex/007]           |                        |
| Apr 25, 2026 |             | 325x5x3, 335x5x2 |        | 2026-04-25T19:30:00Z |                      | [ex/008]           |                        |
| May 24, 2026 |             | 335x5x5          | 182.4  | 2026-05-24T18:00:00Z | 2026-05-24T17:15:00Z | [ex/010, wt/009]   | ex/777                 |

The last row's strength session was also logged on a watch. The script still wrote its own exercise datapoint (`ex/010`, carrying the reps/sets notes) but borrowed the watch session's start/end times for it, and recorded the watch session's resource name (`ex/777`) in `Matched Health Session`. The watch datapoint itself is left untouched.

---

## Setup

### Prerequisites

- A Google account signed into the Google Health app (with workout history).
- Edit access to the spreadsheet and a GCP project.

### 1. Google Cloud Project (GCP)

1. Create a project at [GCP Console](https://console.cloud.google.com).
2. Enable the **Google Health API**.
3. Configure the **OAuth Consent Screen** (User type: **External**). Two publishing options:
   - **In production** (recommended): set **Branding** home/privacy/terms URLs to a domain you've verified in Search Console, then set publishing status to *In production*. A single-user personal sync stays under the 100-user cap, so it runs **without** submitting for Google verification, and — unlike *Testing* — restricted-scope refresh tokens do **not** expire every 7 days. A "verification required" banner appears; ignore it while under the cap. (An improperly configured consent screen, e.g. missing verified-domain branding URLs, can make the consent page return `Error 500` — see caveats.)
   - **Testing** (simpler, but weekly re-auth): leave the status on *Testing* and add your Google email under **Test Users** (required to avoid `Error 403: access_denied`). Restricted-scope refresh tokens expire after 7 days, so you must re-run **Sync ▸ Authorize Health API** about once a week.
4. Note your **Project Number** from IAM & Admin Settings.

### 2. Link Apps Script

1. In your Sheet, open **Extensions ▸ Apps Script**.
2. Go to **Project Settings (⚙)** ▸ **GCP Project ▸ Change project** ▸ paste your **Project Number**.
3. Copy the **Script ID** (under Project Settings).

### 3. Configure OAuth Client

1. GCP Console ▸ **APIs & Services ▸ Credentials ▸ Create Credentials ▸ OAuth client ID** (Web application).
2. Add the **Authorized redirect URI**:
   `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback` (replace `{SCRIPT_ID}`).
3. Copy the **Client ID** and **Client Secret**.
4. In Apps Script **Project Settings (⚙) ▸ Script Properties**, add:
   - `HEALTH_OAUTH_CLIENT_ID`
   - `HEALTH_OAUTH_CLIENT_SECRET`

### 4. Deploy Code

Enable the Apps Script API at [Script Settings](https://script.google.com/home/usersettings), then run:

```bash
npm install
npm run login
cp .clasp.json.example .clasp.json # Paste your Script ID under "scriptId"
npm run push
```

### 5. Set Timezone

In Apps Script **Project Settings (⚙)**, set the project time zone to your local zone. Civil-date filters and synthetic timestamps use this zone. The default in `appsscript.json` is `America/Toronto`.

### 6. Initialize & Authorize

1. In the Apps Script editor, open the `Main` file, select the `setup` function, and click **Run**. The first run prompts for sheet access; approve it. `setup` installs triggers (`onEditTrigger`, `flushIfPending`, `dailyBackstop`) and appends the managed columns (`Exercise Synced At`, `Weight Synced At`, `Created Health IDs`, `Exercise First Edited At`, `Exercises Last Edited At`, `Weight Edited At`, `Matched Health Session`) to the first tab.
2. Refresh your spreadsheet, then select **Sync ▸ Authorize Health API** from the menu. Complete the OAuth consent flow.

The **Sync** menu also exposes:

- **Run now**: sync all dirty rows immediately, waiting for the script lock (the automatic triggers skip when it's held).
- **Resync selected rows**: clears both `Exercise Synced At` and `Weight Synced At` on every selected data row (supports multi-range selections) and resyncs. Unchanged datapoints keep their resource name (idempotent).
- **Resync all rows**: clears both `Synced At` columns for every row and reconciles everything with Google Health. Runs immediately with no confirmation. Unchanged datapoints are left as-is; only changed ones are recreated. If the row count exceeds the per-pass cap, the remainder is deferred to the next poll.
- **Revoke Health API**: clears the stored token.
- **Run setup**: append any missing managed columns and rebuild triggers (after editing timing constants).
- **Run tests**: execute the parser tests inside Apps Script.

---

## Google Cloud caveats

The Google Health API uses **restricted** OAuth scopes. Some gotchas worth knowing before you commit to this setup:

- **7-day refresh token expiry in Testing mode.** With your project's OAuth consent screen set to *Testing*, refresh tokens for restricted scopes expire after 7 days regardless of activity. You'll need to re-run **Sync ▸ Authorize Health API** about once a week. There is no Google-supported way to avoid this for personal Gmail accounts on Testing — move to *In production* under the 100-user cap (next bullet) to escape it.
- **"In production" under the 100-user cap works without verification — and removes the 7-day expiry.** Publishing the consent screen to *In production* keeps restricted-scope refresh tokens alive indefinitely, and a single-user personal sync stays under Google's 100-user cap so **no verification submission is required** (the Verification Center shows a "verification required" banner that you can ignore while under the cap). The catch is consent-screen configuration: the **Branding** home/privacy/terms URLs must point at a domain you've verified in Search Console. An incompletely configured production consent screen (e.g. cleared/missing branding URLs or authorized domains) makes the Google sign-in consent page itself return `Error 500. That's an error.` — so if the flow 500s, check Branding and Authorized Domains before assuming production is unusable.
- **Verification for restricted scopes is impractical for personal use.** If you ever exceed the 100-user cap you'd have to submit for review: Google requires brand verification, scope justification, demo video, and a third-party **CASA security assessment** (~$500–$3000, 6–12 weeks, **annual re-cert required**). Total realistic timeline 2–4 months. Not viable for a one-user personal sync script — staying under the cap avoids it entirely.
- **Cloud Identity Free is an alternative no-fee escape — if you can still sign up.** A custom domain (`you@yourdomain.com`) under [Cloud Identity Free](https://workspace.google.com/signup/gcpfree/welcome) lets you set the consent screen's *User type* to **Internal**, which also removes the 7-day expiry (and the 100-user cap). Google has been actively hiding the free signup flow; as of 2026 it often redirects to paid Workspace. If self-serve signup fails, the remaining options are production-under-the-cap (above), paid Workspace Business Starter (~$7/mo), or living with the weekly re-auth.
- **The unsuffixed `googlehealth.activity_and_fitness` and `googlehealth.health_metrics_and_measurements` scopes are legacy.** They were Google's pre-migration combined read+write scopes, retired on 2026-05-26: *"Replacing read/write scopes with .writeonly. Developers must now explicitly specify read and write permissions"* (per the [Google Health release notes](https://developers.google.com/health/release-notes)). The data plane returns `ACCESS_TOKEN_SCOPE_INSUFFICIENT` if you grant only the unsuffixed scope, and no v4 method lists it in its `scopes` array. The consent screen scope picker still shows their stale pre-migration descriptions, which is misleading. Use the `.readonly` + `.writeonly` suffix variants only.
- **Foreign datapoints cannot be modified by any client.** The Health API enforces ownership on every write path — DELETE returns `403 DATA_POINT_NOT_OWNED_BY_CLIENT` and PATCH returns `400 DATA_POINT_NOT_OWNED_BY_CLIENT` ("Updating data points sourced from other API clients is forbidden"). PUT/POST-to-resource don't exist (404 at Google's frontend). There is no client-side workaround. For workouts your watch/Fitbit logged independently of this script, attach notes in the sheet only — the script's foreign-match logic records the foreign session's resource name in `Matched Health Session` so you can still cross-reference.
- **No newer API version exists.** Only `health:v4` is published in Google's API directory. `v1`/`v1beta`/`v5` all 404; this is the latest and only version.

---

## Configuration & Tuning

Edit [Config.gs](src/Config.gs) to customize:

- `SYNTHETIC_START_HOUR` / `SYNTHETIC_DURATION_HOURS` (default `12` / `1`): synthetic session start hour and duration when edit-derived timing is unavailable (legacy/backfill rows).
- `MIN_EXERCISE_DURATION_MS` / `MAX_EXERCISE_DURATION_MS` (defaults 10 min / 120 min): bounds for the edit-derived interval duration. The floor doubles as the start-only default — a single-edit row (start only, no observed end) records a `MIN`-length session, which also satisfies the Health API's "endTime must be strictly after startTime" requirement; the ceiling caps multi-hour-spanning rows at a plausible workout length.
- `FOREIGN_MATCH_BUFFER_MS` (default 30 min): when checking whether a non-sync-created Strength Training activity (e.g. one logged by a watch/Fitbit) overlaps a row's edit window, allow this much slack on either side. The row **always** writes its own exercise datapoint; an overlapping foreign session only supplies its start/end interval for timing alignment (the manual start/stop is more accurate than the edit-derived window), and the foreign datapoint's resource name is recorded in the `Matched Health Session` column. Rows without same-date exercise edit timestamps get no match and fall through to synthetic/prior timing.
- `BACKSTOP_LOOKBACK_DAYS` / `BACKSTOP_HOUR` (defaults 2 / 4): how many days back (by the row's `Date`) the re-review passes reach. The 5-minute poll re-dirties recent rows that haven't matched a foreign session yet (so a late foreign session re-aligns them within minutes); the daily backstop re-dirties recent rows that *have* matched (so a foreign session whose interval changed after your last edit gets re-borrowed). Both re-dirties are cheap because the exercise re-sync is idempotent — an unchanged row is a read, not a recreate. `BACKSTOP_HOUR` is the local hour the daily trigger fires.
- `ORPHAN_RECONCILE_LOOKBACK_DAYS` (default 7): the same daily trigger scans the last this-many days for sync-created exercise datapoints that no row's `Created Health IDs` references and deletes them. These orphans are leaked by two rare, accepted create windows (a create POST that succeeds server-side but times out client-side and is retried; a 6-minute hard kill landing after the POST returns but before the ID is persisted). Ownership is derived from the listed datapoints themselves, so only datapoints from the same client that created your tracked ones are ever removed — foreign device/watch sessions are never touched. A wider window than the foreign-match backstop since an orphan can otherwise sit indefinitely.
- `POLL_INTERVAL_MIN` (default 5): trigger cadence for `flushIfPending`. Apps Script's minimum is 1 minute.
- `MAX_ROWS_PER_SYNC` (default 100): maximum rows processed per sync pass. Rows are processed newest-first; anything over the cap is deferred to the next poll. At ~2.83s/row observed this leaves comfortable margin under Apps Script's 6-minute execution limit.
- `MAX_BODYWEIGHT_LB` / `MIN_BODYWEIGHT_LB` (defaults 499 / 50): bodyweight values above `MAX` (fat-finger typos like `1850` for `185.0`) or below `MIN` (a rep/set count typed into the Weight column, e.g. `5`) are treated as no-bodyweight and not synced.

Run **Sync ▸ Run setup** after editing timing configurations.

---

## Development

- **Run tests locally**: `npm test` (Node.js runner that simulates the Apps Script globals; fast iteration on parser/formatter changes).
- **Run tests in Apps Script**: **Sync ▸ Run tests** (executes the same suite against the deployed code).
- **Lint markdown**: `npm run lint`.
- **Push changes**: `npm run push`. Pull remote edits back with `npm run pull`.

---

## Troubleshooting

- **403: access_denied / unverified app**: Your email is missing from the GCP OAuth consent screen's **Test Users** list.
- **403: Could not mint UberMint from GaiaMint**: The token contains mixed scopes. Re-run **Sync ▸ Authorize Health API** using the menu.
- **`Health OAuth not configured`**: Ensure `HEALTH_OAUTH_CLIENT_ID` and `HEALTH_OAUTH_CLIENT_SECRET` are set in Apps Script Properties.
- **Redirect URI Mismatch**: Verify the redirect URL in GCP credentials matches your Script ID.
- **Failure emails**: An unrecoverable configuration error (missing required columns, duplicate exercise headers) throws out of the `flushIfPending` trigger, so Apps Script emails the script owner about the failed execution automatically. The dirty flag is left set, so once you fix the misconfig the next poll syncs the backlog without needing an edit or manual sync. While it stays broken the trigger fails every poll; throttle repeat emails by setting the cadence under **Apps Script ▸ Triggers ▸ notifications** (e.g. daily). Transient per-row Health API errors are not unrecoverable; they retry automatically and are visible in the Apps Script **Executions** log.
