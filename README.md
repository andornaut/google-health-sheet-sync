# google-health-spreadsheet

Apps Script that syncs strength-training data from a Google Sheet into the **Google Health app** (Pixel Watch and other Wear OS / connected devices that report through Google Health) via the **Google Health API v4**.

## What it does

For each row in the sheet (one row = one workout session), the script:

1. Parses each exercise cell (e.g. `135x5x3` = 135 lb for 5 reps across 3 sets; a leading `*` means assisted).
2. If the row was previously synced, deletes the datapoints created by the prior sync (best-effort `batchDelete`).
3. POSTs a new `STRENGTH_TRAINING` exercise session for the date with the strength detail in `notes`. Server normalizes `displayName` to `"Strength training"` regardless of what we send. Default interval is 12:00–13:00 local (configurable via `SYNTHETIC_START_HOUR` / `SYNTHETIC_END_HOUR` in `src/Config.gs`); when a row's date has multiple sheet entries, each is offset by one hour.
4. POSTs a `weight` datapoint for the row's date (sampled at noon local) using the `Weight` column.
5. Stores the returned resource names in a hidden `Health IDs` column so the next re-sync of the row can `batchDelete` them before creating fresh ones (idempotent).
6. Stamps a hidden `Synced At` column with the sync timestamp.

The Google Health API enforces datapoint ownership: **we cannot modify sessions recorded by your Pixel Watch / Fitbit**. If you record a workout on the watch *and* log it in the sheet, you'll see two entries for that day in the Google Health app — one from the watch (HR, calories, zones) and one from this script (strength detail in notes). They live side-by-side.

Triggered ~1 minute after your last edit (sliding-window debounce) and re-checked hourly as a backstop.

## Spreadsheet shape

The script auto-detects columns by header name. Required headers:

- `Date` (column 1)
- `Weight` (last data column — your bodyweight in lb)
- Any number of exercise columns in between (their header text is used in the synced notes)

Cell grammar, one entry per line:

| Cell | Meaning |
|------|---------|
| `135` | 1 rep at 135 lb (top single) |
| `135x5` | 5 reps at 135 lb |
| `135x5x3` | 5 reps × 3 sets at 135 lb (reps always second) |
| `*135x5x3` | Same, marked assisted (band / spotter etc.) |
| `135x5x3`<br>`145x3x2` | Two distinct entries (separator can be newline, comma, or semicolon) |
| `135x5x3, 145x3x2` | Same as above on one line |
| empty | Exercise skipped |

The script adds a `Synced At` column at the right and hides it on first run.

## Prerequisites

- A Google account that is **already signed into the Google Health app** on your phone (with workout sessions appearing in the app's history). This is what the Health API reads from — if the account isn't onboarded to Google Health, every API call returns 403 `Could not mint UberMint from GaiaMint`. See the troubleshooting section if you're unsure.
- Edit access to the target spreadsheet (using the same Google account).
- Ability to create a Google Cloud project.
- Optional: [`clasp`](https://github.com/google/clasp) for pushing code from this repo (otherwise paste files manually).

## One-time setup

### 1. Google Cloud project

1. Go to <https://console.cloud.google.com> and **Create project** (any name).
2. **APIs & Services ▸ Enabled APIs ▸ + Enable APIs and Services** ▸ search **Google Health API** ▸ Enable.
3. Configure the OAuth consent screen. In the **current Console UI** this lives under **APIs & Services ▸ OAuth consent screen** (left nav), which opens the **Google Auth Platform** with tabs: **Overview / Branding / Audience / Clients / Data Access / Settings**.
   - **Branding** — App name: anything (e.g. `Health Sheet Sync`); User support email + Developer contact: your email.
   - **Audience** —
     - **User type**: External.
     - **Publishing status**: leave at **Testing** — do **not** click "Publish app". (For personal use, no privacy review needed; the 100-user / 7-day refresh-token cap doesn't affect a single-user script.)
     - **Test users** ▸ **+ Add users** ▸ enter the Google account email you'll sign in with (the same one that owns the spreadsheet). **This is required** — without it you'll hit `Error 403: access_denied` ("Access blocked: ... has not completed the Google verification process") when you authorize the script.
   - **Data Access** (Scopes): leave empty. Apps Script declares scopes at runtime from `appsscript.json`.
4. Note the project's **project number** (purely numeric, e.g. `123456789012`). Find it under **IAM & Admin ▸ Settings**, or on the project info card on the Cloud Console home/Dashboard. The **project picker** at the top of the console shows only the Project **ID** (a hyphenated string like `my-project-abc123`), which is **not** the same thing.

### 2. Link Apps Script to the GCP project

1. Open the target spreadsheet ▸ **Extensions ▸ Apps Script**.
2. **Project Settings (⚙)** ▸ **Google Cloud Platform (GCP) Project ▸ Change project** ▸ paste the **project number** from step 1.4 (not the project ID).
3. Copy the **Script ID** from this same Project Settings page (under "IDs", ~57 chars). You'll need it twice below.

### 3. Create an OAuth client for the Google Health API

The Google Health backend **rejects** access tokens that contain any non-Health scopes (returns `403 Could not mint UberMint from GaiaMint`). Apps Script's built-in token always bundles the sheet-access scopes with everything else, so it can't talk to the Health API. The workaround is to run a separate OAuth flow with **only** the Health scopes, using the [apps-script-oauth2 library](https://github.com/googleworkspace/apps-script-oauth2).

1. Cloud Console → **APIs & Services ▸ Credentials ▸ + Create credentials ▸ OAuth client ID**.
   - **Application type**: **Web application**.
   - **Name**: anything (e.g. `Apps Script Health OAuth`).
   - **Authorized redirect URIs** ▸ **+ Add URI**:

     ```text
     https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
     ```

     Replace `{SCRIPT_ID}` with the Script ID from step 2.3.
2. Click **Create**. Copy the **Client ID** and **Client secret** from the dialog (you can also view them later from the Credentials list).
3. In the Apps Script editor → **Project Settings (⚙) ▸ Script Properties ▸ + Add script property**, add **two** rows:
   - `HEALTH_OAUTH_CLIENT_ID` → the client ID from step 3.2
   - `HEALTH_OAUTH_CLIENT_SECRET` → the client secret from step 3.2

### 4. Push the code

**Option A — `clasp` (no global install):**

First, enable the **Apps Script API** for your Google account (one-time, per-user — independent of the GCP project setup above):

1. Visit <https://script.google.com/home/usersettings>
2. Toggle **Google Apps Script API** to **On**. Wait ~1 minute for it to propagate.

Then grab the **Script ID**:

1. In the spreadsheet, **Extensions ▸ Apps Script** to open the editor.
2. **Project Settings (⚙ in the left sidebar)** → under **IDs**, copy the **Script ID** (long alphanumeric string, ~57 chars).

Then from this repo:

```bash
npm install                     # installs @google/clasp as a dev dependency
npm run login                   # opens browser, authorizes clasp against your Google account
cp .clasp.json.example .clasp.json
# edit .clasp.json: paste the Script ID into the "scriptId" field
npm run push                    # uploads src/, test/, appsscript.json to the script
```

**Option B — manual:**

In the Apps Script editor, create one file per `.gs` in `src/` and `test/`, plus replace the contents of `appsscript.json` (you may need to enable "Show appsscript.json manifest" in Project Settings).

### 5. Initialize the spreadsheet integration

In the Apps Script editor:

1. Open **`src/Main.gs`** from the file list on the left. (The function dropdown only lists functions from the file you have open — `Config.gs` shows "No functions" because it only declares constants.)
2. In the toolbar, select function **`setup`** from the dropdown.
3. Click **Run**.

- You'll be prompted to authorize the script. You'll see an **"unverified app"** warning — click **Advanced ▸ Go to (project)** because your email is on the Testing user list.
- This adds and hides the `Synced At` and `Health IDs` columns and installs the three triggers (`onEdit`, 1-minute debounce flush, hourly backstop).

Confirm in **Triggers (⏰)** that all three are listed.

### 6. Authorize the Google Health API (separate from step 5)

Step 5's authorization grants Apps Script access to the **sheet**. The **Google Health API needs its own, isolated grant** (the Health backend refuses tokens that bundle non-Health scopes).

1. Reload the spreadsheet so the **Sync** menu appears.
2. Sheet menu ▸ **Sync ▸ Authorize Health API**. A dialog opens with a link.
3. Click the link, sign in with the **same Google account** that owns your Google Health data, grant the two `googlehealth.*` scopes. Close the success tab.
4. (You can re-run this any time, or use **Sync ▸ Revoke Health API authorization** to clear and start over.)

## Daily use

Edit rows as you normally would. Your last edit triggers a sync about 1 minute later. To force a sync:

- Sheet menu **Sync ▸ Run now** — full pass.
- Sheet menu **Sync ▸ Force resync current row** — clears the active row's `Synced At` and runs immediately.

The custom `Sync` menu appears on sheet open; if missing, run `onOpen` once from the editor.

## API keys / secrets

The Google Health OAuth **client ID and client secret** live in the script's **Script Properties** (set in step 3.3) — not committed to git, not in `.env`. They're only needed by the apps-script-oauth2 library to perform the one-time Health authorization flow. The resulting access + refresh tokens are stored by the library in **User Properties** (per Google account, opaque to you).

Apps Script's built-in OAuth covers the sheet access; nothing extra to configure for that.

## OAuth scopes

### Apps Script's built-in token (declared in `appsscript.json`)

| Scope | Why |
|-------|-----|
| `spreadsheets.currentonly` | Read/write the bound sheet |
| `script.external_request` | `UrlFetchApp` to `health.googleapis.com` |
| `script.scriptapp` | Install triggers |
| `script.container.ui` | Show custom menu + alerts |

### Separate Google Health token (granted via Sync ▸ Authorize Health API)

| Scope                                          | Why                              |
|------------------------------------------------|----------------------------------|
| `googlehealth.activity_and_fitness`            | List + PATCH exercise sessions   |
| `googlehealth.health_metrics_and_measurements` | PATCH bodyweight data points     |

These are kept in a different token because the Google Health backend rejects mixed-scope tokens. See troubleshooting for the `Could not mint UberMint` error if this isn't done correctly.

## Tuning

`src/Config.gs` exposes:

- `DEBOUNCE_MS` — quiet period before a flush (default 60 s; sliding window).
- `DEBOUNCE_CHECK_INTERVAL_MIN` — how often the flush trigger fires (default 1 min).
- `BACKSTOP_INTERVAL_HOURS` — how often the unconditional sync runs (default 1 h).
- `SYNTHETIC_START_HOUR` / `SYNTHETIC_END_HOUR` — interval (local) for the exercise sessions this script creates. Default 12:00–13:00. Each additional sheet row for the same date is offset by one hour.
- `SYNC_EXERCISES` / `SYNC_WEIGHT` — toggles per data type (both default `true`; flip off for debugging).
- `EXERCISE_ABBREVIATIONS` — short codes used in the `displayName` summary (the server will normalize the title to `"Strength training"` anyway, so this is mostly cosmetic).
- `SHEET_NAME` — defaults to `Sheet1`; falls back to the active sheet if missing.

After changing config, re-run `installTriggers` (Sync menu ▸ Re-install triggers) to pick up timing changes.

## Verification

End-to-end (run once after setup):

1. Pick any row with exercise + bodyweight data.
2. Sheet menu ▸ **Sync ▸ Force resync current row**.
3. Open the Google Health app on your phone for that date — a new `Strength training` entry (12:00–13:00 local by default) should appear with your lifts in the notes, and a `Weight` entry should match the sheet value.
4. Edit a cell in that row → the previous Health entries should be deleted and replaced on the next sync (no duplicates).
5. Edit a cell, wait 30 s, edit another cell — sync should fire ~1 minute after the **second** edit (sliding-window debounce).

Unit tests: Sheet menu ▸ **Sync ▸ Run tests** (shows pass/fail in an alert and in **Executions**).

## Troubleshooting

- **`clasp push` says "User has not enabled the Apps Script API"** — flip <https://script.google.com/home/usersettings> ▸ **Google Apps Script API** to On, wait ~1 min, retry.
- **"Access blocked: ... has not completed the Google verification process" (Error 403: access_denied)** — the Google account you're signed in as isn't on the consent screen's Test users list. Go to **APIs & Services ▸ OAuth consent screen ▸ Audience ▸ Test users ▸ + Add users** and add it. Retry in ~30 s.
- **403 on Health API with `Could not mint UberMint from GaiaMint`** — the access token being sent to the Health API contains non-Health scopes, which the Health backend rejects. This means the script is falling back to Apps Script's built-in token instead of using the dedicated Health OAuth flow. Check:
  1. Sync menu ▸ **Authorize Health API** — did you click the link and complete the consent flow? Use **Revoke Health API authorization** then **Authorize** again to redo it.
  2. Script Properties contain `HEALTH_OAUTH_CLIENT_ID` and `HEALTH_OAUTH_CLIENT_SECRET` (Project Settings ▸ Script Properties).
  3. You signed in with the **same Google account** that's onboarded to the Google Health app.
  4. The OAuth client's **Authorized redirect URI** in GCP exactly matches `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback`.
- **`Health OAuth not configured` thrown from `getHealthService`** — Script Properties missing. See step 3.3 of setup.
- **Authorize Health API link 404s or "redirect_uri_mismatch"** — the redirect URI on the OAuth client doesn't match the Script ID. Fix it in GCP ▸ APIs & Services ▸ Credentials ▸ (edit your OAuth client).
- **403 on Health API (other)** — consent screen not configured, wrong GCP project linked to the script, or scopes not granted. Re-authorize.
- **404 on PATCH** — the watch session for that date was deleted in the app. The row's `Synced At` stays blank and is retried on the next pass.
- **Cell can't be parsed** — see **Executions** in the Apps Script editor; the offending line is logged and skipped, the rest of the row still syncs.
- **No triggers firing** — check **Triggers (⏰)**; re-install via Sync menu ▸ Re-install triggers.
- **API in active development** — the Google Health API is GA through May 2026 with possible breaking changes; if a PATCH starts 400-ing, check the [release notes](https://developers.google.com/health/release-notes).
