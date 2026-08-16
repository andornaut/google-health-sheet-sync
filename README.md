# Google Health Sheet Sync

[![CI](https://github.com/andornaut/google-health-sheet-sync/actions/workflows/test.yml/badge.svg)](https://github.com/andornaut/google-health-sheet-sync/actions/workflows/test.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Google Apps Script that syncs strength-training and bodyweight data from Google Sheets to the **Google Health app** via the **Google Health API v4**.

## Features

- **Strength exercises**: Parses lifts (e.g. `135x5x3`, `*135x5x3` for assisted) and logs them as `STRENGTH_TRAINING` sessions.
- **Bodyweight**: Logs weight data points.
- **Idempotent sync**: Each pass reconciles a row's Health datapoint(s) with current cell content. Editing a synced row clears its `Synced At` stamps and reprocesses it. The exercise datapoint is delete+recreated **only when the interval or notes changed**; an unchanged re-sync keeps the existing datapoint and costs just a read. (The Health API has no working update path for exercise datapoints (`PATCH` is a server-side no-op) so delete+POST is the only update mechanism.)
- **Edit-derived timing**: Start/end times come from when you first/last edited the row, so a session reflects when you actually worked out. A single edit records a default-length session; a second edit sets the real end time (`endTime` grows as you log sets, clamped to `[MIN, MAX]`, default 10 min – 2 h). Edit timestamps are only trusted when the first one falls on the row's `Date` (in your local time zone), so editing an old row won't shift its datapoint to today. A later correction, more than the 2 h cap from your first edit, leaves the recorded interval alone; a correction inside that window still extends `endTime`, since that is the same mechanism that grows the session while you log sets. On re-sync the prior interval is preserved. Rows without edit timestamps (backfill) fall back to a synthetic noon slot on `Date`.
- **Foreign-activity matching**: If a Strength Training session from a watch or other app overlaps the row's edit window (timestamps on the row's `Date`, within 10 min slack), the script borrows that session's start/end times for the datapoint it writes and records the match in `Matched Health Session`. The script **always writes its own datapoint** (with reps/sets notes); the foreign datapoint is left untouched (the API forbids modifying it anyway). A backstop runs every 4 hours (`BACKSTOP_INTERVAL_HOURS`) and re-reviews exercise rows from the last 2 days (`BACKSTOP_LOOKBACK_DAYS`), both unmatched (to align a late foreign session) and matched (to re-borrow a foreign interval extended after your last edit). This runs off the 5-minute poll, so foreign re-review doesn't query the Health API every 5 minutes; the tradeoff is that a late foreign session aligns within ~4 hours rather than minutes.
- **Two-phase sync**: Weight and exercise sync independently, with their own `Synced At` stamps and column-aware dirty tracking. Editing only Weight re-pushes only the weight datapoint, and vice versa. Editing only `Date` on an empty row is a no-op.
- **Clearing a cell counts as an edit**: Deleting a logged set or a bodyweight re-syncs the row, so the Health datapoint is rewritten or deleted to match. This works for one cell at a time (Sheets only reports the previous value for single-cell edits); to clear a block of cells at once, select those rows and follow it with **Sync ▸ Resync selected rows**.
- **Sync on edit**: Editing a row syncs it **immediately** via the `onEdit` trigger (under a non-blocking script lock). During an edit burst the lock batches work; an edit that can't grab the lock is caught by the next poll.
- **Automated retry net**: A 5-minute poll re-syncs rows left dirty by lock contention or transient failures (it queries the Health API only when there is pending work). A backstop runs every 4 hours, re-reviewing exercise rows from the last 2 days (matched and unmatched) for foreign-session alignment, and deletes orphaned sync-created datapoints (no row references them).
- **Per-pass cap**: Up to 100 rows per pass (newest first) to stay under Apps Script's 6-minute limit; the remainder is picked up on the next poll.

> [!NOTE]
> The Health API enforces strict ownership. Datapoints this script logs live side-by-side with sessions from your watch (e.g. Pixel Watch, Fitbit).

---

## Spreadsheet Layout

The script always operates on the **first (leftmost) tab**: selected by position, so keep your data tab leftmost. Columns are auto-detected by header name; any header that isn't `Date`, `Weight`, or a managed column is treated as an exercise.

- **`Date`** (leftmost)
- **Exercise columns** (middle; headers are exercise names)
- **`Weight`** (rightmost; bodyweight in lb)

### Grammar for Exercise Cells

One entry per line (newline, comma, or semicolon separated):

| Cell               | Meaning                                                           |
| ------------------ | ----------------------------------------------------------------- |
| `135`              | 135 lb (reps and sets unspecified)                                |
| `135x5`            | 5 reps at 135 lb                                                  |
| `135x5x3`          | 5 reps × 3 sets at 135 lb                                         |
| `135x5x0`          | Start marker: 0 sets done; anchors start time, not shown in notes |
| `*135x5x3`         | Assisted reps                                                     |
| `135x5x3, 145x3x2` | Multiple entries in one cell                                      |

Unspecified fields are omitted from the notes. A `Bench press` cell with `135` renders as `Bench press, 135 lbs`; `135x5` as `Bench press, 135 lbs, 5 reps`; `135x5x3` as `Bench press, 135 lbs, 3 sets of 5`.

A **zero-set** entry (`135x5x0`) is retained but suppressed from notes: it anchors the start time and can match a foreign session, but writes no datapoint until you log a real set. A **zero-rep** entry (`135x0`) is dropped entirely as "not performed".

### Example

Trimmed to a few columns; Health resource names abbreviated. `Exercise Synced At` / `Weight Synced At` are stamped independently per phase.

| Date         | Bench press | Deadlift | Weight | Exercise Synced At   | Weight Synced At     | Created Health IDs | Matched Health Session |
| ------------ | ----------- | -------- | ------ | -------------------- | -------------------- | ------------------ | ---------------------- |
| Jan 2, 2026  | 210         |          | 190.0  | 2026-01-02T18:30:00Z | 2026-01-02T17:45:00Z | [ex/001, wt/002]   |                        |
| Jan 15, 2026 | 215x4       |          |        | 2026-01-15T20:00:00Z |                      | [ex/004]           |                        |
| Jan 18, 2026 |             | 295x4x6  | 187.5  | 2026-01-18T17:45:00Z | 2026-01-18T17:00:00Z | [ex/005, wt/006]   |                        |
| May 24, 2026 |             | 335x5x5  | 182.4  | 2026-05-24T18:00:00Z | 2026-05-24T17:15:00Z | [ex/010, wt/009]   | ex/777                 |

The last row was also logged on a watch: the script still wrote its own datapoint (`ex/010`) but borrowed the watch session's start/end times and recorded the watch session (`ex/777`) in `Matched Health Session`. The watch datapoint is untouched.

---

## Setup

### Prerequisites

- A Google account signed into the Google Health app (with workout history).
- Edit access to the spreadsheet and a GCP project.

### 1. Google Cloud Project

1. Create a project at [GCP Console](https://console.cloud.google.com).
2. Enable the **Google Health API**.
3. Configure the **OAuth Consent Screen** (User type: **External**), then pick a publishing option:
   - **In production** (recommended): set **Branding** home/privacy/terms URLs to a domain you've verified in Search Console, then publish. A single-user sync stays under the 100-user cap, so it runs **without** Google verification and restricted-scope refresh tokens **don't** expire weekly. Ignore the "verification required" banner. (A misconfigured consent screen can make the consent page return `Error 500`, see caveats.)
   - **Testing** (simpler, weekly re-auth): leave status on _Testing_ and add your email under **Test Users** (avoids `Error 403: access_denied`). Restricted-scope refresh tokens expire after 7 days, so re-run **Sync ▸ Authorize Health API** about weekly.
4. Note your **Project Number** from IAM & Admin Settings.

### 2. Link Apps Script

1. In your Sheet, open **Extensions ▸ Apps Script**.
2. **Project Settings (⚙) ▸ GCP Project ▸ Change project** ▸ paste your **Project Number**.
3. Copy the **Script ID** (under Project Settings).

### 3. Configure OAuth Client

1. GCP Console ▸ **APIs & Services ▸ Credentials ▸ Create Credentials ▸ OAuth client ID** (Web application).
2. Add the **Authorized redirect URI**: `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback`.
3. Copy the **Client ID** and **Client Secret**.
4. In Apps Script **Project Settings (⚙) ▸ Script Properties**, add `HEALTH_OAUTH_CLIENT_ID` and `HEALTH_OAUTH_CLIENT_SECRET`.

### 4. Deploy Code

Enable the Apps Script API at [Script Settings](https://script.google.com/home/usersettings), then run:

```bash
npm install
npm run login
cp .clasp.json.example .clasp.json # Paste your Script ID under "scriptId"
npm run push
```

### 5. Set Timezone

In Apps Script **Project Settings (⚙)**, set the project time zone to your local zone (used by date filters and synthetic timestamps). Default in `appsscript.json` is `America/Toronto`.

### 6. Initialize & Authorize

1. In the Apps Script editor, open `Main`, select `setup`, and **Run**. Approve the sheet-access prompt. `setup` installs triggers (`syncOnEdit`, `flushPending`, `backstop`) and appends the managed columns to the first tab.
2. Refresh the spreadsheet, then **Sync ▸ Authorize Health API** and complete the OAuth flow.

The **Sync** menu also exposes:

- **Run now**: sync all dirty rows immediately, waiting for the lock.
- **Resync selected rows**: clears both `Synced At` columns on selected rows and resyncs (unchanged datapoints keep their resource name). Select the rows on the synced tab (the first one); a selection on any other tab is refused.
- **Resync all rows**: clears both `Synced At` columns for every row and reconciles everything. Runs immediately; only changed datapoints are recreated.
- **Revoke Health API**: clears the stored token.
- **Run setup**: append missing managed columns and rebuild triggers (after editing timing constants).
- **Run tests**: run the parser / pure-helper suite inside Apps Script. (The orchestration suite needs the Node test harness, so run it locally with `npm test`.)

---

## Google Cloud caveats

The Google Health API uses **restricted** OAuth scopes. Gotchas worth knowing:

- **7-day refresh-token expiry in Testing mode.** Restricted-scope refresh tokens expire after 7 days regardless of activity, requiring weekly re-auth. No supported workaround for personal Gmail on Testing: move to _In production_ under the 100-user cap to escape it.
- **"In production" under the 100-user cap works without verification.** It keeps refresh tokens alive indefinitely, and a single-user sync stays under the cap so **no verification submission is required** (ignore the "verification required" banner). Catch: the **Branding** home/privacy/terms URLs must point at a domain verified in Search Console. A misconfigured production consent screen makes the sign-in page itself return `Error 500`, check Branding and Authorized Domains first if the flow 500s.
- **Verification for restricted scopes is impractical for personal use.** Exceeding the 100-user cap requires brand verification, scope justification, demo video, and a third-party **CASA security assessment** (~$500–$3000, 6–12 weeks, annual re-cert). Staying under the cap avoids it entirely.
- **Cloud Identity Free is a no-fee alternative, if you can still sign up.** A custom domain under [Cloud Identity Free](https://workspace.google.com/signup/gcpfree/welcome) lets you set _User type_ to **Internal** (removes the 7-day expiry and the cap). Google has been hiding the free signup; as of 2026 it often redirects to paid Workspace. Fallbacks: production-under-the-cap, paid Workspace (~$7/mo), or weekly re-auth.
- **The unsuffixed `googlehealth.*` scopes are legacy.** The combined read+write scopes were retired 2026-05-26 (see [release notes](https://developers.google.com/health/release-notes)); the data plane returns `ACCESS_TOKEN_SCOPE_INSUFFICIENT` for them. Use the `.readonly` + `.writeonly` variants only (the consent screen still shows stale descriptions for the old scopes).
- **Foreign datapoints cannot be modified by any client.** DELETE returns `403 DATA_POINT_NOT_OWNED_BY_CLIENT`, PATCH returns `400`, and PUT/POST-to-resource don't exist. For workouts your watch logged independently, attach notes in the sheet only; the foreign-match logic records the foreign session's name in `Matched Health Session` for cross-reference.
- **No newer API version exists.** Only `health:v4` is published; `v1`/`v1beta`/`v5` all 404.

---

## Configuration & Tuning

Edit [Config.gs](src/Config.gs) to customize:

- `SYNTHETIC_START_HOUR` / `SYNTHETIC_DURATION_HOURS` (`12` / `1`): synthetic session start hour and duration when edit-derived timing is unavailable (backfill rows).
- `MIN_EXERCISE_DURATION_MS` / `MAX_EXERCISE_DURATION_MS` (10 min / 120 min): bounds for the edit-derived duration. The floor doubles as the start-only default and satisfies the API's "endTime after startTime" requirement; the ceiling caps multi-hour rows.
- `FOREIGN_MATCH_BUFFER_MS` (10 min): slack on either side when checking whether a foreign Strength Training session overlaps a row's edit window. An overlapping session only supplies its interval (the script always writes its own datapoint), and its name is recorded in `Matched Health Session`.
- `BACKSTOP_LOOKBACK_DAYS` / `BACKSTOP_INTERVAL_HOURS` (2 / 4): how many days back the backstop's foreign-match re-review reaches, and how often it runs. It re-dirties recent exercise rows (matched and unmatched); the next poll re-syncs them. Cheap because the re-sync is idempotent (an unchanged row is a read, not a recreate).
- `ORPHAN_RECONCILE_LOOKBACK_DAYS` (7): the backstop scans this many days back for sync-created datapoints no row references and deletes them. Ownership is derived from the datapoints themselves, so foreign sessions are never touched.
- `POLL_INTERVAL_MIN` (5): cadence for `flushPending` (Apps Script minimum is 1).
- `MAX_ROWS_PER_SYNC` (100): max rows per pass (newest first); overflow defers to the next poll. ~2.83s/row leaves margin under the 6-minute limit.
- `MAX_BODYWEIGHT_LB` / `MIN_BODYWEIGHT_LB` (499 / 50): values above `MAX` (typos like `1850`) or below `MIN` (a rep count typed into Weight) are treated as no-bodyweight.

Run **Sync ▸ Run setup** after editing timing constants.

---

## Developing

- **Run tests locally**: `npm test` (Node.js runner simulating Apps Script globals).
- **Run tests in Apps Script**: **Sync ▸ Run tests**.
- **Lint markdown**: `npm run lint`.
- **Push / pull**: `npm run push` / `npm run pull`.

---

## Troubleshooting

- **403: access_denied / unverified app**: Your email is missing from the consent screen's **Test Users** list.
- **403: Could not mint UberMint from GaiaMint**: Mixed scopes in the token. Re-run **Sync ▸ Authorize Health API**.
- **`Health OAuth not configured`**: Set `HEALTH_OAUTH_CLIENT_ID` and `HEALTH_OAUTH_CLIENT_SECRET` in Script Properties.
- **Redirect URI Mismatch**: Verify the redirect URL in GCP credentials matches your Script ID.
- **A corrected value on an old row lost its original time**: Correcting an exercise cell keeps the recorded start and end: once your edits stop looking like the workout itself (more than `MAX_EXERCISE_DURATION_MS` apart), the sync reuses the existing datapoint's interval and only rewrites the notes. Bodyweight is the case to watch: typing over the cell PATCHes in place and keeps the original sample time, but clearing it first does not, because the clear deletes the datapoint and the retype creates a new one at the synthetic noon slot on the row's `Date`. The same applies to an exercise cell you clear completely and retype, since the delete removes the interval there was to preserve.
- **Failure emails**: An unrecoverable config error (missing columns, duplicate exercise headers) throws out of `flushPending`, so Apps Script emails the owner. The dirty flag stays set, so the backlog syncs once you fix the misconfig. Throttle repeat emails under **Apps Script ▸ Triggers ▸ notifications**. A row that fails unexpectedly (a Sheets service hiccup mid-write) also emails, naming the row, but only after every other ready row has been attempted (one bad row never blocks the others), and the message carries the pass counts so you can see how much succeeded. Transient per-row API errors retry automatically and appear in the **Executions** log.
