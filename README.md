# Google Health Sheet Sync

Google Apps Script to sync strength-training and bodyweight data from Google Sheets to the **Google Health app** via the **Google Health API v4**.

## Features

- **Strength Exercises**: Parses lifts (e.g., `135x5x3`, `*135x5x3` for assisted) and logs them as `STRENGTH_TRAINING` sessions.
- **Bodyweight**: Logs weight data points.
- **Idempotent Sync**: Deletes previous sync data before creating new entries to prevent duplicates.
- **Edit-derived timing**: The activity's start/end times are taken from when you first/last edited the row, so the Google Health session reflects when you actually did the workout. Rows without edit timestamps (e.g. backfill) fall back to a synthetic noon-ordinal slot.
- **Foreign-activity matching**: If a Strength Training session you logged on a watch or in another app already covers the same workout (overlapping edit window, or 1:1 by ordinal when the row has no edit timestamps), the script skips writing its own duplicate exercise datapoint and records the matched session in the row's `Matched Health Session` column. The foreign datapoint is left untouched — its calories, heart rate, and recording method are preserved. Bodyweight still syncs as normal.
- **Automated**: Polls every 5 minutes for rows whose last edit was at least 45 minutes ago, with an hourly backstop trigger as a safety net.

> [!NOTE]
> The Google Health API enforces strict ownership. Datapoints logged by this script live side-by-side with sessions recorded by your watch (e.g., Pixel Watch, Fitbit).

---

## Spreadsheet Layout

Columns are auto-detected by header name. Required layout:

- **`Date`** (Column 1)
- **Exercise Columns** (Middle columns; headers are used as exercise names)
- **`Weight`** (Last column; bodyweight in lb)

### Grammar for Exercise Cells

One entry per line (newline, comma, or semicolon separated):

| Cell | Meaning |
| ------ | --------- |
| `135` | 1 rep at 135 lb |
| `135x5` | 5 reps at 135 lb |
| `135x5x3` | 5 reps × 3 sets at 135 lb |
| `*135x5x3` | Assisted reps |
| `135x5x3, 145x3x2` | Multiple distinct sets/exercises in one cell |

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

Choose one option:

- **Option A (clasp)**: Enable the Apps Script API at [Script Settings](https://script.google.com/home/usersettings). Then run:

  ```bash
  npm install
  npm run login
  cp .clasp.json.example .clasp.json # Paste your Script ID under "scriptId"
  npm run push
  ```

- **Option B (manual)**: Copy files from `src/` and `test/` into the Apps Script editor.

### 5. Initialize & Authorize

1. Open `src/Main.gs` in the editor, select the `setup` function, and click **Run**. This authorizes sheet access, installs triggers (`onEditTrigger`, `flushIfPending`, `backstop`), and appends the managed columns (`Synced At`, `Health IDs`, `First Edited At`, `Last Edited At`, `Matched Health Session`) to the sheet.
2. Refresh your spreadsheet, then select **Sync ▸ Authorize Health API** from the menu. Complete the OAuth consent flow.

The **Sync** menu also exposes:

- **Run now**: sync dirty rows immediately, bypassing the quiesce window.
- **Force resync current row**: clears `Synced At` on the active row and resyncs (bypasses quiesce).
- **Force resync ALL rows**: clears `Synced At` for every row and re-uploads everything to Google Health (bypasses quiesce). Confirms first.
- **Revoke Health API authorization**: clears the stored token.
- **Re-install triggers**: rebuild triggers after editing timing constants.
- **Run tests**: execute the local parser tests inside Apps Script.

---

## Configuration & Tuning

Edit [Config.gs](file:///home/andornaut/src/github.com/andornaut/google-health-sheet-sync/src/Config.gs) to customize:

- `SYNTHETIC_START_HOUR` / `SYNTHETIC_DURATION_HOURS` (default `12` / `1`): synthetic session start hour and duration when edit-derived timing is unavailable (legacy/backfill rows).
- `LAST_EDIT_QUIESCE_MS` (default 45 min): how long a row must sit idle after its last edit before it's eligible to sync. Lets the activity's end time reflect when the workout actually finished.
- `FOREIGN_MATCH_BUFFER_MS` (default 30 min): when checking whether a non-sync-created Strength Training activity matches a row's edit window, allow this much slack on either side. If a foreign activity overlaps, the row is treated as already represented in Health (no own exercise written) and the foreign datapoint's resource name is recorded in the `Matched Health Session` column.
- `POLL_INTERVAL_MIN` (default 5) and `BACKSTOP_INTERVAL_HOURS` (default 1): trigger cadence.
- `SYNC_EXERCISES` / `SYNC_WEIGHT` (default `true`): toggle which datapoint types are written.
- `EXERCISE_ABBREVIATIONS` (cosmetic mapping).

Run **Sync ▸ Re-install triggers** after editing timing configurations.

---

## Troubleshooting

- **403: access_denied / unverified app**: Your email is missing from the GCP OAuth consent screen's **Test Users** list.
- **403: Could not mint UberMint from GaiaMint**: The token contains mixed scopes. Re-run **Sync ▸ Authorize Health API** using the menu.
- **`Health OAuth not configured`**: Ensure `HEALTH_OAUTH_CLIENT_ID` and `HEALTH_OAUTH_CLIENT_SECRET` are set in Apps Script Properties.
- **Redirect URI Mismatch**: Verify the redirect URL in GCP credentials matches your Script ID.
- **Tests**: Run unit tests from the sheet menu (**Sync ▸ Run tests**).
