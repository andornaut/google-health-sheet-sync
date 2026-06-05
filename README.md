# Google Health Sheet Sync

Google Apps Script to sync strength-training and bodyweight data from Google Sheets to the **Google Health app** via the **Google Health API v4**.

## Features

- **Strength Exercises**: Parses lifts (e.g., `135x5x3`, `*135x5x3` for assisted) and logs them as `STRENGTH_TRAINING` sessions.
- **Bodyweight**: Logs weight data points.
- **Delete+recreate sync**: Each sync pass rebuilds the row's Health datapoint(s) from current cell content. Edits to a synced row clear the `Synced At` stamps and the row is reprocessed on the next poll, with `endTime`, `activeDuration`, and notes refreshed from the current state. (The Google Health API has no functional update mechanism for exercise datapoints — `PATCH` is a server-side no-op and `PUT`/POST-to-resource don't exist; delete+POST is the only path.)
- **Edit-derived timing, date-aware**: The activity's start/end times come from when you first/last edited the row, so the Google Health session reflects when you actually did the workout. Edit timestamps are only trusted when their civil date matches the row's `Date` column — editing an old row today won't shift its Health datapoint to today. On re-sync, the prior datapoint's interval/sample time is fetched and preserved (the content updates but the timing stays put). Rows without edit timestamps (e.g. backfill) fall back to a synthetic noon-ordinal slot on the row's `Date`.
- **Foreign-activity matching**: If a Strength Training session you logged on a watch or in another app already covers the same workout (overlapping edit window when the row's edit timestamps fall on its `Date`, or ordinal pairing among same-date rows and candidates otherwise), the script skips writing its own duplicate exercise datapoint and records the matched session in the row's `Matched Health Session` column. The foreign datapoint is left untouched — its calories, heart rate, and recording method are preserved. (The API enforces this anyway — DELETE on a non-script-owned datapoint returns `403 DATA_POINT_NOT_OWNED_BY_CLIENT`.) Bodyweight still syncs as normal.
- **Two-phase sync**: Weight and exercise sync independently with their own `Synced At` stamps and column-aware dirty tracking. Editing only the Weight column re-pushes just the weight datapoint and leaves the exercise interval untouched; editing only an exercise column re-pushes just the exercise datapoint. Editing the Date column alone on an otherwise-empty row is a no-op (nothing to sync until exercise or weight content lands).
- **Edit-burst debounce**: A dirty row whose `Exercises Last Edited At` is within 60 seconds is skipped on the current poll and retried on the next, so a poll firing mid-edit doesn't push a half-typed row. The weight phase has no debounce — `Weight Edited At` advances on every weight cell edit and the weight datapoint pushes on the next poll.
- **Automated**: Polls every 5 minutes for dirty rows past the debounce.
- **Per-pass cap**: Each sync processes up to 75 rows (newest dates first) to stay under Apps Script's 6-minute execution limit. Remaining rows are picked up on the next poll.

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
| `*135x5x3` | Assisted reps |
| `135x5x3, 145x3x2` | Multiple distinct sets/exercises in one cell |

Unspecified fields are omitted from the Google Health notes. For example, a `Bench press` cell containing `135` renders as `Bench press, 135 lbs`; `135x5` renders as `Bench press, 135 lbs, 5 reps`; `135x5x3` renders as `Bench press, 135 lbs, 3 sets of 5`.

### Example

Trimmed to a few exercise columns and the most relevant managed columns. Health resource names abbreviated as `ex/NNN` / `wt/NNN`; the remaining managed columns `Exercise First Edited At`, `Exercises Last Edited At`, and `Weight Edited At` are omitted here. `Exercise Synced At` and `Weight Synced At` are stamped independently per phase — a row with only bodyweight leaves `Exercise Synced At` blank, and vice versa:

| Date         | Bench press | Deadlift         | Weight | Exercise Synced At   | Weight Synced At     | Created Health IDs | Matched Health Session |
| ------------ | ----------- | ---------------- | ------ | -------------------- | -------------------- | ------------------ | ---------------------- |
| Jan 2, 2026  | 210         |                  | 190.0  | 2026-01-02T18:30:00Z | 2026-01-02T17:45:00Z | [ex/001, wt/002]   |                        |
| Jan 15, 2026 | 215x4       |                  |        | 2026-01-15T20:00:00Z |                      | [ex/004]           |                        |
| Jan 18, 2026 |             | 295x4x6          | 187.5  | 2026-01-18T17:45:00Z | 2026-01-18T17:00:00Z | [ex/005, wt/006]   |                        |
| Jan 24, 2026 | *225        |                  |        | 2026-01-24T18:00:00Z |                      | [ex/007]           |                        |
| Apr 25, 2026 |             | 325x5x3, 335x5x2 |        | 2026-04-25T19:30:00Z |                      | [ex/008]           |                        |
| May 24, 2026 |             | 335x5x5          | 182.4  | 2026-05-24T18:00:00Z | 2026-05-24T17:15:00Z | [wt/009]           | ex/010                 |

The last row's strength session was logged on a watch first; the script skipped writing its own exercise datapoint and recorded the foreign session in `Matched Health Session`. Only the bodyweight was written by the script (see `Created Health IDs`).

---

## Setup

### Prerequisites

- A Google account signed into the Google Health app (with workout history).
- Edit access to the spreadsheet and a GCP project.

### 1. Google Cloud Project (GCP)

1. Create a project at [GCP Console](https://console.cloud.google.com).
2. Enable the **Google Health API**.
3. Configure the **OAuth Consent Screen** (External, Testing):
   - Add your Google email under **Test Users** (required to avoid `Error 403: access_denied`).
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

1. In the Apps Script editor, open the `Main` file, select the `setup` function, and click **Run**. The first run prompts for sheet access; approve it. `setup` installs triggers (`onEditTrigger`, `flushIfPending`) and appends the managed columns (`Exercise Synced At`, `Weight Synced At`, `Created Health IDs`, `Exercise First Edited At`, `Exercises Last Edited At`, `Weight Edited At`, `Matched Health Session`) to the first tab.
2. Refresh your spreadsheet, then select **Sync ▸ Authorize Health API** from the menu. Complete the OAuth consent flow.

The **Sync** menu also exposes:

- **Run now**: sync dirty rows immediately, bypassing the 60-second edit-burst debounce.
- **Resync selected rows**: clears both `Exercise Synced At` and `Weight Synced At` on every selected data row (supports multi-range selections) and resyncs (bypasses debounce).
- **Resync all rows**: clears both `Synced At` columns for every row and re-uploads everything to Google Health (bypasses debounce). Runs immediately with no confirmation. If the row count exceeds the per-pass cap, the remainder is deferred to the next poll.
- **Revoke Health API**: clears the stored token.
- **Run setup**: append any missing managed columns and rebuild triggers (after editing timing constants).
- **Run tests**: execute the parser tests inside Apps Script.

---

## Google Cloud caveats

The Google Health API uses **restricted** OAuth scopes. Some gotchas worth knowing before you commit to this setup:

- **7-day refresh token expiry in Testing mode.** With your project's OAuth consent screen set to *Testing*, refresh tokens for restricted scopes expire after 7 days regardless of activity. You'll need to re-run **Sync ▸ Authorize Health API** about once a week. There is no Google-supported way to avoid this for personal Gmail accounts on Testing.
- **"In production" without verification 500s the consent flow.** Publishing the consent screen to *In production* without completing verification causes the Google sign-in consent page itself to return `Error 500. That's an error.` Stay on Testing unless you've gone through verification.
- **Verification for restricted scopes is impractical for personal use.** Google requires brand verification, scope justification, demo video, and a third-party **CASA security assessment** (~$500–$3000, 6–12 weeks, **annual re-cert required**). Total realistic timeline 2–4 months. Not viable for a one-user personal sync script.
- **Cloud Identity Free is the only no-fee escape hatch — if you can still sign up.** A custom domain (`you@yourdomain.com`) under [Cloud Identity Free](https://workspace.google.com/signup/gcpfree/welcome) lets you set the consent screen's *User type* to **Internal**, which removes the 7-day expiry forever. Google has been actively hiding the free signup flow; as of 2026 it often redirects to paid Workspace. If self-serve signup fails for you, the only alternatives are paid Workspace Business Starter (~$7/mo) or living with the weekly re-auth.
- **The unsuffixed `googlehealth.activity_and_fitness` and `googlehealth.health_metrics_and_measurements` scopes are legacy.** They were Google's pre-migration combined read+write scopes, retired on 2026-05-26: *"Replacing read/write scopes with .writeonly. Developers must now explicitly specify read and write permissions"* (per the [Google Health release notes](https://developers.google.com/health/release-notes)). The data plane returns `ACCESS_TOKEN_SCOPE_INSUFFICIENT` if you grant only the unsuffixed scope, and no v4 method lists it in its `scopes` array. The consent screen scope picker still shows their stale pre-migration descriptions, which is misleading. Use the `.readonly` + `.writeonly` suffix variants only.
- **Foreign datapoints cannot be modified by any client.** The Health API enforces ownership on every write path — DELETE returns `403 DATA_POINT_NOT_OWNED_BY_CLIENT` and PATCH returns `400 DATA_POINT_NOT_OWNED_BY_CLIENT` ("Updating data points sourced from other API clients is forbidden"). PUT/POST-to-resource don't exist (404 at Google's frontend). There is no client-side workaround. For workouts your watch/Fitbit logged independently of this script, attach notes in the sheet only — the script's foreign-match logic records the foreign session's resource name in `Matched Health Session` so you can still cross-reference.
- **No newer API version exists.** Only `health:v4` is published in Google's API directory. `v1`/`v1beta`/`v5` all 404; this is the latest and only version.

---

## Configuration & Tuning

Edit [Config.gs](src/Config.gs) to customize:

- `SYNTHETIC_START_HOUR` / `SYNTHETIC_DURATION_HOURS` (default `12` / `1`): synthetic session start hour and duration when edit-derived timing is unavailable (legacy/backfill rows).
- `LAST_EDIT_QUIESCE_MS` (default 60 sec): edit-burst "still typing" guard for the exercise phase. A row whose `Exercises Last Edited At` is within this window is skipped on the current poll and retried on the next, so a poll firing mid-edit doesn't push a half-typed row. Weight phase has no debounce.
- `MIN_EXERCISE_DURATION_MS` / `MAX_EXERCISE_DURATION_MS` (defaults 5 min / 120 min): bounds for the edit-derived interval duration. The floor satisfies the Health API's "endTime must be strictly after startTime" requirement when only a single edit has happened; the ceiling caps multi-hour-spanning rows at a plausible workout length.
- `FOREIGN_MATCH_BUFFER_MS` (default 30 min): when checking whether a non-sync-created Strength Training activity (e.g. one logged by a watch/Fitbit) overlaps a row's edit window, allow this much slack on either side. The row **always** writes its own exercise datapoint; an overlapping foreign session only supplies its start/end interval for timing alignment (the manual start/stop is more accurate than the edit-derived window), and the foreign datapoint's resource name is recorded in the `Matched Health Session` column. Rows without same-date exercise edit timestamps get no match and fall through to synthetic/prior timing.
- `POLL_INTERVAL_MIN` (default 5): trigger cadence for `flushIfPending`. Apps Script's minimum is 1 minute.
- `MAX_ROWS_PER_SYNC` (default 90): maximum rows processed per sync pass. Rows are processed newest-first; anything over the cap is deferred to the next poll. At ~3.5s/row this leaves comfortable margin under Apps Script's 6-minute execution limit.
- `MAX_BODYWEIGHT_LB` (default 999): bodyweight values above this are treated as typos and ignored.

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
