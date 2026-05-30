# Google Health Sheet Sync

Google Apps Script to sync strength-training and bodyweight data from Google Sheets to the **Google Health app** via the **Google Health API v4**.

## Features

- **Strength Exercises**: Parses lifts (e.g., `135x5x3`, `*135x5x3` for assisted) and logs them as `STRENGTH_TRAINING` sessions.
- **Bodyweight**: Logs weight data points.
- **Idempotent Sync**: Deletes previous sync data before creating new entries to prevent duplicates.
- **Edit-derived timing**: The activity's start/end times are taken from when you first/last edited the row, so the Google Health session reflects when you actually did the workout. Rows without edit timestamps (e.g. backfill) fall back to a synthetic noon-ordinal slot.
- **Foreign-activity matching**: If a Strength Training session you logged on a watch or in another app already covers the same workout (overlapping edit window, or 1:1 by ordinal when the row has no edit timestamps), the script skips writing its own duplicate exercise datapoint and records the matched session in the row's `Matched Health Session` column. The foreign datapoint is left untouched — its calories, heart rate, and recording method are preserved. Bodyweight still syncs as normal.
- **Two-phase sync**: Weight and exercise sync independently. Weight syncs immediately on the next poll (no quiesce); exercise waits for the quiesce window so the activity's end time reflects when the workout actually finished. Each phase has its own `Synced At` stamp.
- **Automated**: Polls every 5 minutes for rows whose last edit was at least 45 minutes ago, with an hourly backstop trigger as a safety net. Re-editing a synced row clears both `Exercise Synced At` and `Weight Synced At` and restarts the quiesce window so the next sync reflects the updated content.
- **Per-pass cap**: Each sync processes up to 75 rows (newest dates first) to stay under Apps Script's 6-minute execution limit. Remaining rows are picked up on the next poll.

> [!NOTE]
> The Google Health API enforces strict ownership. Datapoints logged by this script live side-by-side with sessions recorded by your watch (e.g., Pixel Watch, Fitbit).

---

## Spreadsheet Layout

The script always operates on the first tab of the spreadsheet (leftmost; name doesn't matter). Columns are auto-detected by header name and position doesn't matter — any header that isn't `Date`, `Weight`, or a managed column is treated as an exercise. Recommended layout:

- **`Date`** (leftmost)
- **Exercise Columns** (middle; headers are used as exercise names)
- **`Weight`** (rightmost; bodyweight in lb)

### Grammar for Exercise Cells

One entry per line (newline, comma, or semicolon separated):

| Cell | Meaning |
| ------ | --------- |
| `135` | 1 rep at 135 lb |
| `135x5` | 5 reps at 135 lb |
| `135x5x3` | 5 reps × 3 sets at 135 lb |
| `*135x5x3` | Assisted reps |
| `135x5x3, 145x3x2` | Multiple distinct sets/exercises in one cell |

### Example

Trimmed to a few exercise columns and the most relevant managed columns. Health resource names abbreviated as `ex/NNN` / `wt/NNN`; the remaining managed columns `First Edited At` and `Last Edited At` are omitted here. `Exercise Synced At` and `Weight Synced At` are stamped independently per phase — a row with only bodyweight leaves `Exercise Synced At` blank, and vice versa:

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

1. Open `src/Main.gs` in the editor, select the `setup` function, and click **Run**. The first run prompts for sheet access; approve it. `setup` installs triggers (`onEditTrigger`, `flushIfPending`, `backstop`) and appends the managed columns (`Exercise Synced At`, `Weight Synced At`, `Created Health IDs`, `First Edited At`, `Last Edited At`, `Matched Health Session`) to the first tab.
2. Refresh your spreadsheet, then select **Sync ▸ Authorize Health API** from the menu. Complete the OAuth consent flow.

The **Sync** menu also exposes:

- **Run now**: sync dirty rows immediately, bypassing the exercise quiesce window. Weight already syncs without quiesce, so this only changes behavior for exercise content.
- **Force resync current row**: clears both `Exercise Synced At` and `Weight Synced At` on the active row and resyncs (bypasses quiesce).
- **Force resync all rows**: clears both `Synced At` columns for every row and re-uploads everything to Google Health (bypasses quiesce). Runs immediately with no confirmation. If the row count exceeds the per-pass cap, the remainder is deferred to the next poll.
- **Revoke Health API**: clears the stored token.
- **Run setup**: append any missing managed columns and rebuild triggers (after editing timing constants).
- **Run tests**: execute the parser tests inside Apps Script.

---

## Configuration & Tuning

Edit [Config.gs](src/Config.gs) to customize:

- `SYNTHETIC_START_HOUR` / `SYNTHETIC_DURATION_HOURS` (default `12` / `1`): synthetic session start hour and duration when edit-derived timing is unavailable (legacy/backfill rows).
- `LAST_EDIT_QUIESCE_MS` (default 45 min): how long a row must sit idle after its last edit before it's eligible to sync. Lets the activity's end time reflect when the workout actually finished.
- `FOREIGN_MATCH_BUFFER_MS` (default 30 min): when checking whether a non-sync-created Strength Training activity matches a row's edit window, allow this much slack on either side. If a foreign activity overlaps, the row is treated as already represented in Health (no own exercise written) and the foreign datapoint's resource name is recorded in the `Matched Health Session` column.
- `POLL_INTERVAL_MIN` (default 5) and `BACKSTOP_INTERVAL_HOURS` (default 1): trigger cadence.
- `MAX_ROWS_PER_SYNC` (default 75): maximum rows processed per sync pass. Rows are processed newest-first; anything over the cap is deferred to the next poll. At ~3.3s/row this leaves comfortable margin under Apps Script's 6-minute execution limit.
- `SYNC_EXERCISES` / `SYNC_WEIGHT` (default `true`): toggle which datapoint types are written.
- `EXERCISE_ABBREVIATIONS` (cosmetic mapping).

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
