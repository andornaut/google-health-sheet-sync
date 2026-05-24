# Google Health Sheet Sync

Google Apps Script to sync strength-training and bodyweight data from Google Sheets to the **Google Health app** via the **Google Health API v4**.

## Features

- **Strength Exercises**: Parses lifts (e.g., `135x5x3`, `*135x5x3` for assisted) and logs them as `STRENGTH_TRAINING` sessions.
- **Bodyweight**: Logs weight data points.
- **Idempotent Sync**: Deletes previous sync data before creating new entries to prevent duplicates.
- **Automated**: Runs ~1 minute after editing (debounced) with an hourly backstop trigger.

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

1. Open `src/Main.gs` in the editor, select the `setup` function, and click **Run** (authorizes sheet access and configures triggers).
2. Refresh your spreadsheet, then select **Sync ▸ Authorize Health API** from the menu. Complete the OAuth consent flow.

---

## Configuration & Tuning

Edit [Config.gs](file:///home/andornaut/src/github.com/andornaut/google-health-sheet-sync/src/Config.gs) to customize:

- `SYNTHETIC_START_HOUR` / `SYNTHETIC_END_HOUR` (default `12` / `13`)
- `DEBOUNCE_MS` (sync delay after last edit, default 60s)
- `EXERCISE_ABBREVIATIONS` (cosmetic mapping)

*Note: Run **Sync ▸ Re-install triggers** after editing timing configurations.*

---

## Troubleshooting

- **403: access_denied / unverified app**: Your email is missing from the GCP OAuth consent screen's **Test Users** list.
- **403: Could not mint UberMint from GaiaMint**: The token contains mixed scopes. Re-run **Sync ▸ Authorize Health API** using the menu.
- **`Health OAuth not configured`**: Ensure `HEALTH_OAUTH_CLIENT_ID` and `HEALTH_OAUTH_CLIENT_SECRET` are set in Apps Script Properties.
- **Redirect URI Mismatch**: Verify the redirect URL in GCP credentials matches your Script ID.
- **Tests**: Run unit tests from the sheet menu (**Sync ▸ Run tests**).
