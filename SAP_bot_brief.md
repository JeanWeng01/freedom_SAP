# SAP Freight Workflow Automation — Documentation

## Overview

A Python + Selenium browser automation bot for SAP Fiori (SAP BTP / SAP Logistics Business Network) that handles repetitive freight workflows for a trucking company processing hundreds of invoices weekly.

The bot has **two operational modes** based on how much human oversight each workflow needs:

- **Autonomous (Tiles 1 & 2):** Runs on a schedule on Railway 24/7. Confirms freight orders and reports stop times. No human input.
- **Human-gated (Tiles 3 & 4):** Triggered manually (locally or via Railway endpoint). Creates invoices and uploads PODs. Designed to allow pauses + human review.

Data backbone is a **Google Sheet** (read/write) and **Google Drive** (POD PDF storage).

---

## Tech stack

- **Python** 3.13 (Railway), 3.12+ locally
- **Selenium WebDriver** — browser automation
- **Edge** (local dev on Windows) and **Chrome** (Railway/Linux), switchable via `config.yaml`
- **Google Sheets API** + **Google Drive API** via `google-api-python-client`
- **Flask** — local web server for status/manual trigger endpoints
- **PyYAML** — config
- **python-dotenv** — secrets from `.env`

---

## Project structure

```
freedom_SAP/
├── Dockerfile                      # Railway image: Chrome + chromedriver + Python deps
├── railway.json                    # Railway deployment config
├── SAP_bot_brief.md               # this doc
└── sap_bot/
    ├── config.yaml                 # browser, dry_run, screenshot flags
    ├── .env                        # SAP login URL/user/password (gitignored)
    ├── service_account.json        # Google credentials (gitignored)
    ├── requirements.txt
    ├── main.py                     # CLI entry point (interactive menu / --tile flags)
    ├── server.py                   # Flask web server (Railway entry point)
    └── bot/
        ├── driver_setup.py         # Edge/Chrome WebDriver init
        ├── login.py                # SAP SSO login
        ├── google_sheets.py        # To Do tab read/write + status updates
        ├── google_drive.py         # POD PDF download from Drive
        ├── tile1_confirmation.py   # Freight Orders for Confirmation
        ├── tile2_reporting.py      # Freight Orders for Reporting
        ├── tile3_invoicing.py      # Invoice Freight Documents
        ├── tile4_pod_upload.py     # Manage Freight Execution / POD upload
        └── utils.py                # screenshots, retries, click_tile, dry-run decorator
```

---

## Configuration

### `sap_bot/config.yaml`

```yaml
browser: "edge"                   # "edge" (local) or "chrome" (Railway)
dry_run: false                    # if true, never clicks Submit/Confirm/Report
step_through: false               # if true, pauses before destructive actions for tiles 3 & 4
screenshot_on_action: false       # screenshots on every action (verbose)
screenshot_on_error: true         # always screenshot on errors
autonomous_interval_minutes: 120  # legacy — actual schedule is in server.py env vars
```

### `sap_bot/.env`

```
SAP_LOGIN_URL=https://...
SAP_LAUNCHPAD_URL=https://lbnlivelwteu10.lbn.cfapps.eu10.hana.ondemand.com/cp.portal/site?sap-language=en#Shell-home
SAP_USERNAME=dispatch@freedomtransportationinc.ca
SAP_PASSWORD=...
```

On Railway, the same vars are set in the Railway environment dashboard. `HEADLESS=true` and `BROWSER=chrome` also set there.

### Google credentials

Service account JSON at `sap_bot/service_account.json`, OR the same JSON passed via `GOOGLE_SERVICE_ACCOUNT_JSON` env var (Railway). Service account must have access to:
- The "To Do"/"Status" Google Sheet (env `GOOGLE_SHEET_ID`)
- The Drive folder containing POD PDFs

### Railway scheduler env vars (in `server.py`)

```
TILE12_RUN_HOURS=9:00,12:00,15:00          # ET, when tiles 1 & 2 auto-run
ENABLE_TILES_34=false                       # if true, tiles 3 & 4 also auto-run
TILE34_RUN_HOURS=9:00,10:00,...            # tile 3 & 4 schedule when enabled
TZ=America/New_York                         # so HH:MM is interpreted as ET
```

---

## Login (`bot/login.py`)

- Navigate to launchpad URL
- If session not active, fill SSO email + password, click login
- Wait for launchpad to render
- Screenshot on success/failure

---

## Google Sheet schema — "To Do" tab

| Col | Header | Purpose | Used by |
|---|---|---|---|
| A | Document_1 | Freight document number, leg 1 | tiles 3, 4 |
| B | Document_2 | Freight document number, leg 2 (blank if single-leg) | tiles 3, 4 |
| C | Notes | Human notes — bot ignores for processing, preserves on move to Status tab | — |
| D | Invoice | Invoice number to enter on the SAP invoice page | tile 3 |
| E | POD_Filename | Bare filename(s), comma-separated for multiple. Bot appends `.pdf` | tile 4 |
| F | Charge_Hrs_1 | Waiting Charges hours for leg 1. Multiplied by `HOURLY_RATE` (42.84) → CAD | tile 3 |
| G | Charge_Hrs_2 | Waiting Charges hours for leg 2 (collective invoices) | tile 3 |
| H | Pause | If `1` or `true`: bot fills the invoice but clicks **Save** instead of **Submit** (human review) | tile 3 |
| I | tile3_timestamp | When tile 3 last touched this row. **I2 specifically** is the local-run heartbeat | bot writes |
| J | tile3_status | `tile3_in_progress` (yellow) → `invoiced` (green) / `drafted, awaiting human action` (orange) / `tile3_error: ...` (red) | bot writes |
| K | tile4_timestamp | When tile 4 last touched this row | bot writes |
| L | tile4_status | `tile4_in_progress` (yellow) → `pod_uploaded` (green) / `pod_error: ...` (red) | bot writes |

### Skip behavior — bot ignores rows whose col J is exactly:
- `invoiced` — already done
- `wait` — manually set by user when row isn't ready yet
- `drafted, awaiting human action` — was Pause=1, already saved as draft for human

Same for tile 4: skips rows where col L is exactly `pod_uploaded`.

### Status colors

| Color | Meaning |
|---|---|
| 🟢 Green | Bot finished successfully |
| 🟡 Yellow | Bot is actively working on this row right now |
| 🟠 Orange | Bot done, awaiting human (paused → drafted) |
| 🔴 Red | Bot errored out |

---

## Tile 1 — Freight Orders for Confirmation (autonomous)

**Goal:** Confirm ALL freight orders with status "New" AND "Updated".

**Flow:**
1. Navigate to tile
2. For each status in `["New", "Updated"]`:
   - Apply Freight Order Status filter
   - Click "All" tab to see unfiltered count
   - Loop until 0 remain (max 10 passes per status):
     - Read expected count from header
     - Scroll to lazy-load all rows
     - Click select-all checkbox → click "Confirm"
     - Refresh and re-check count
3. Returns `{new: N, updated: M, total: N+M}` to the server, which adds to daily totals

**Edge cases:** count not decreasing across passes = stuck loop, breaks out. Max 10 passes. Status filter cleared between New and Updated.

---

## Tile 2 — Freight Orders for Reporting (autonomous)

**Goal:** For each freight order in "Events to Report", copy each stop's Planned Time into Final Time and click Report.

**Flow:**
1. Navigate to tile, click "Events to Report" tab
2. For each row:
   - Click into the order's detail page
   - For each stop's event:
     - Read Planned Time string (e.g. `Mar 17, 2026, 12:00 AM EST` or `Mar 16, 2026, 11:00PM UTC-5`)
     - Strip the trailing timezone via regex (matches `UTC[+-]?\d+` or 2–5 letter codes like EST/EDT)
     - Click "Report Final Time" → popup
     - Paste the stripped date/time into the Final Time field (timezone NOT entered separately — SAP infers it)
     - Click "Report"
3. Returns `{stops_reported: N}` to the server

---

## Tile 3 — Invoice Freight Documents (human-gated)

**Goal:** For each row in the To Do tab, create the invoice (single or collective), enter invoice number, add Waiting Charges if any, submit (or save if paused).

### Per-row flow (`process_row`)

1. **Filter** for `Document_1` (and `Document_2` if collective) in the Freight Document filter.
   - Single doc: type, press Enter
   - Collective: type doc 1, Enter, type doc 2, Enter (both appear in results)
2. Check default "To be Invoiced" tab. If both docs present → select all visible, proceed.
   - If 0 in default tab → click "All" tab, look at all matching rows, identify which are still usable (status `Invoicing in process` or `Not Yet Invoiced`). Skip ones already `Completely invoiced`.
3. Click "Create Invoice" (single) or "Create Collective Invoice" (multiple usable docs)
4. **Two outcomes:**
   - Invoice page opens → proceed to fill (5)
   - Page goes to **Drafts** → return `"drafted"`, then `process_drafted_invoice()` recovers it (Manage Invoices → Drafts tab → filter → click into draft → fill → submit)
5. **Fill (`fill_invoice_and_submit`):**
   - Enter invoice number (col D)
   - If `has_charges_leg1` (col F set): `add_charge(leg_num=1)` — Waiting Charges, amount = hrs × 42.84
   - If collective + `has_charges_leg2` (col G set): `add_charge(leg_num=2)` — same
   - If `Pause=1`: click **Save** (not Submit), return `"paused"` → row marked `drafted, awaiting human action`
   - Otherwise: click **Submit** with verification (see below)

### `add_charge(leg_num)` — adding a Waiting Charges row

For leg 1, the section is already at the top of the page; just click the leg-1 Add ▼ button.

For leg 2, SAP's page is significantly taller now (see "SAP quirks" below). Bot:
1. **Collapses leg 1's section** by clicking its tree-icon toggle (`<span id="...treeicon">` inside the first cell of leg 1's header row in the `sapUiTable`). This shrinks leg 1's content and brings leg 2's section header (with its Add button) up into the visible viewport.
2. Finds leg 2's Add button via doc-aware crawler (every `<button>` with text "Add", picks the one whose closest `Freight Document <NNN>` ancestor matches `document_2`).
3. **Verifies before clicking** with `elementFromPoint` — if the Add button's coords are visually covered by the sticky Save bar, abort with a loud error rather than click blindly.
4. Click via **ActionChains** (real mouse events) — JS click doesn't trigger SAP UI5's tap handler for the Add ▼ dropdown.

After Add: select "Charge" from the dropdown → click the blank charge description input → category popup opens → search "Waiting Charges" → select → enter amount in Rate Amount/Unit field.

### `click_submit` — verifying success

After clicking Submit, bot calls `dismiss_any_popup()`. **If a popup was dismissed**, that means SAP rejected the submission (validation error, etc.) — bot returns `False`. Caller writes `tile3_error: Submit triggered SAP popup — invoice likely still in Drafts` to col J. Prevents the false-positive-"invoiced" bug we hunted down.

### Pause = 1 lifecycle

| Run | Bot does |
|---|---|
| First run after Pause=1 set | Fills invoice, clicks Save (not Submit). Marks col J = `drafted, awaiting human action` (orange). Human reviews and submits manually. |
| Subsequent runs | Skips the row (col J in SKIP_STATUSES) so bot never re-creates the same draft |

---

## Tile 4 — Manage Freight Execution / POD Upload (human-gated)

**Goal:** Upload one or more POD PDFs to the Proof of Delivery section in **Stop 2** of each freight order.

### Per-row flow (`process_item`)

1. Filter by Document_1
2. Click into the first row in the result list (clicks the row container, not the radio button)
3. Verify navigation succeeded by checking for `Stop \d` / `Information` / `Attachments` text
4. Scroll down + click "Expand All" (or click Stop 2 header) to reveal Stop 2's content
5. Find the "Proof of Delivery" section's Report button → click it
6. In the popup:
   - Read the **Planned On** time, strip timezone, paste into Final Time field
   - Upload the PDF file(s) — multi-file via newline-joined paths sent to the file input in one shot
   - Click Report
7. Repeat for `Document_2` if present (collective invoice with both legs needing PODs)

### POD file source — Google Drive

`google_drive.py` downloads POD PDFs from a Drive folder by filename. Filename(s) come from col E of the To Do row. Comma-separated for multiple files. Bot appends `.pdf` automatically (col E should NOT include the extension). Files download to a temp dir and are cleaned up after upload.

---

## Web server (`server.py`)

Flask app, the Railway entry point. Serves on `$PORT` (Railway sets this).

### Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/` | Health check + dry_run flag |
| GET | `/status` | Running state + per-tile last result + `daily_totals` (resets at midnight) |
| GET/POST | `/run/tile1` | Manually trigger tile 1 |
| GET/POST | `/run/tile2` | Tile 2 |
| GET/POST | `/run/tile3` | Tile 3 |
| GET/POST | `/run/tile4` | Tile 4 |
| GET/POST | `/run/all` | Tiles 1, 2, 3, 4 in sequence |
| GET/POST | `/run/invoices` | Tiles 3 & 4 only |

`/status` payload includes `daily_totals` — a running tally of `tile1_new_confirmed`, `tile1_updated_confirmed`, `tile1_total_confirmed`, `tile2_stops_reported` for today (server local time = ET). Resets on first run after midnight. `app.json.sort_keys = False` preserves dict insertion order.

### Global serialization

A single `run_lock = threading.Lock()` serializes ALL tile runs. Two tiles can NEVER run simultaneously, regardless of trigger source. Manual triggers and scheduled runs all queue behind the same lock.

### Schedulers

`scheduler_loop_12` and `scheduler_loop_34` are background threads that watch the clock and fire at HH:MM matches in `TILE12_RUN_HOURS` / `TILE34_RUN_HOURS`. Tile 3 & 4 scheduler is **disabled by default** — enable via `ENABLE_TILES_34=true`.

### Local-run heartbeat

When `main.py` runs locally, it writes the current timestamp to `'To Do'!I2` and refreshes every 5 minutes via a daemon thread. The Railway scheduler checks this cell before firing — if the heartbeat is fresh (< 30 min old), it **skips** the scheduled run to avoid colliding with the local user. On clean exit, `main.py` clears I2.

---

## Railway deployment

Image built from `Dockerfile`:
1. `python:3.13-slim` base
2. Installs `google-chrome-stable` + `chromedriver` matching Chrome major version (downloaded from `googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_<major>`)
3. Installs Python deps from `sap_bot/requirements.txt`
4. `ENV HEADLESS=true PYTHONUNBUFFERED=1 TZ=America/New_York`
5. `CMD python server.py`

`driver_setup.py` checks for `/usr/local/bin/chromedriver` and uses it explicitly when present. Falls back to Selenium Manager for local Windows dev where the path doesn't exist.

**Why chromedriver is baked in:** Selenium Manager's runtime download was failing on Railway (Chrome auto-updated to a version chromedriver hadn't been published for yet). Baking the matching pair into the image means no runtime download = no failure window.

---

## SAP quirks worth knowing

These are non-obvious behaviors of SAP Fiori we've hit and worked around. Documented so future debugging doesn't have to re-derive them.

### 1. Excluded charges (5 zero rows per freight doc)

SAP added 5 placeholder "Excluded Charge" rows (Admin Fee, Ferry Ticket Fee, Additional Handling Weight Surcharge, Shunting Operations, Storage in Transit) to every freight doc's Charges section. Each has rate `0.00`. Caused two bugs in our code:

- The "find first 0.00 rateInput" logic for entering Waiting Charges amount was matching a leftover excluded row instead of the just-added charge row → `Element not currently interactable` because excluded rows are read-only. **Fix:** find the rate input by clone-number proximity to the description input we just filled.
- These rows make each leg's section significantly taller — leg 2's Add button is now well below the fold instead of being naturally on screen.

### 2. ObjectPageLayout scroll strobe

The draft invoice page uses SAP UI5's `sap.uxap.ObjectPageLayout`, which has an "anchor bar" feature that auto-scrolls to align sections with the top of the viewport. When you scroll down, it detects the scroll, decides "you're navigating to section X," and snaps back. **Result: the page is essentially un-scrollable** — even for human users. The bot can't reach leg 2's Add button via any scroll mechanism.

**Workaround:** instead of scrolling, **collapse leg 1's section** by clicking its `treeicon` ▼ toggle. This shrinks leg 1 to a 1-row header, lifting leg 2's section (with its Add button) into the visible viewport without any scroll.

### 3. ActionChains for SAP UI5 controls

SAP UI5 abstracts events through a "tap" handler that wraps the full `mousedown` → `mouseup` → `click` sequence (plus touch events for mobile). **JavaScript-synthetic clicks** (`element.click()` via `execute_script`) do NOT trigger this handler for many controls — including the Add ▼ dropdown and the tree-icon collapse toggle. **Use ActionChains** (real mouse events at the element's coordinates) for these.

The bot's general click strategy is `element.click()` first (handles 95% of cases), ActionChains as fallback. For known SAP UI5 controls (Add ▼, treeicon), we go ActionChains-first and JS only as a last resort.

### 4. "Document saved" / "No changes are made to the invoice" toasts

Both indicate the bot's click landed on the **Save** button (sticky bottom-right of the draft invoice page) instead of the intended target. Almost always means the target was visually covered by the Save button at the click's screen coordinates. The verify-before-click step (`elementFromPoint` check) prevents this.

### 5. SAP's row-click on charge cells opens a side detail panel

If the bot accidentally clicks on a charge row's cell area, SAP opens a side panel showing that row's details (Line Number, Charge Description, Calculation Method, Payment Terms). This was the symptom when our early collapse-toggle attempts hit a wrong element near the row.

### 6. First-call cushion wait

On the very first navigation to tile 3 per session, the bot waits a fixed 7 seconds after the body-text load check before proceeding. Without it, the first filter operation silently no-ops because SAP's filter widgets aren't yet interactive even though the page text is rendered. Subsequent navigations skip this wait.

---

## Safety modes

### Dry-run mode

`dry_run: true` in config OR `--dry-run` CLI flag OR `DRY_RUN=true` env var. The `@destructive_action` decorator on tile 3's `click_submit` (and similar) checks the flag and **skips the actual click** while still navigating, filling fields, and logging. Lets you verify the bot's behavior without touching live data.

### Step-through mode (tiles 3 & 4)

`step_through: true`. Bot pauses before every destructive action and prompts the user via console: `Press Enter to continue, 'skip', or 'quit'`. For row-by-row sign-off during early live testing.

### Loud failure, never silent

We hunted down a class of false-positive bugs where the bot would write `invoiced` to col J even though the underlying SAP submission had silently failed. The fix is now baked in: every step that touches SAP returns a bool / error string, and any failure propagates up to the sheet as `tile3_error: <reason>` instead of being swallowed. Better to leave a row visibly errored for retry than to lie about success.

---

## Logging & screenshots

- Logs written to `sap_bot/logs/sap_bot_<YYYYMMDD>_<HHMMSS>.log` (timestamped per run)
- Old logs (>7 days) auto-cleaned at start of each scheduled cycle
- Screenshots written to `sap_bot/screenshots/` — disabled by default in headless (Railway) for performance, configurable via `screenshot_on_action` in config

---

## CLI reference

```bash
# Interactive menu
python sap_bot/main.py

# Specific tile(s)
python sap_bot/main.py --tile 3
python sap_bot/main.py --tile 1 2

# Dry run
python sap_bot/main.py --tile 3 --dry-run

# Step-through
python sap_bot/main.py --tile 3 --step-through

# Autonomous loop (legacy local mode — Railway uses server.py instead)
python sap_bot/main.py --autonomous
```

---

## Coworker setup checklist

1. Install Python 3.13 (check "Add to PATH" during install)
2. Copy the entire `freedom_SAP/` folder to their machine (anywhere — paths are all relative)
3. `pip install -r sap_bot/requirements.txt`
4. Drop in `sap_bot/.env` and `sap_bot/service_account.json` (transfer privately)
5. Edge browser is preinstalled on Windows; Selenium Manager downloads `msedgedriver` on first run
6. Run `python sap_bot/main.py` or use a `.bat` shortcut

No path edits to the code needed. Everything resolves via `__file__` relative paths.
