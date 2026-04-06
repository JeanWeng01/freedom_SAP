# SAP Freight Workflow Automation — Claude Code Brief

## Project Overview

A Python + Selenium browser automation bot ("the bot") to automate repetitive SAP Fiori portal tasks for a trucking company processing hundreds of freight invoices weekly. The portal runs on SAP BTP / SAP Logistics Business Network / SAP Fiori Launchpad.

The bot is split into **two operational modes** reflecting how much human oversight each workflow requires:

- **Autonomous Mode** (Tiles 1 & 2): Runs continuously or on a schedule with no human input. Should ideally run as a background service so staff never need to manually open these tiles.
- **Human-Gated Mode** (Tiles 3 & 4): Human triggers execution, can pause/stop. Designed for future automation expansion.

---

## Tech Stack

- **Language:** Python 3.x
- **Browser Automation:** Selenium WebDriver
- **Browser:** Microsoft Edge (primary/developer), Chrome (coworkers) — must be configurable via a config file, not hardcoded
- **Excel parsing:** `openpyxl` or `pandas` for reading `SAP_bot.xlsx`
- **Logging:** Python `logging` module, writing timestamped logs to file
- **Scheduling (Tiles 1 & 2):** `schedule` library or Windows Task Scheduler invocation
- **Config:** A `config.yaml` or `config.ini` file for credentials, browser choice, file paths, SAP URL

---

## Project Structure (suggested)

```
sap_bot/
├── config.yaml                  # credentials, browser, SAP URL, paths
├── SAP_bot.xlsx                 # input data file (user-managed)
├── main.py                      # entry point, mode selector
├── bot/
│   ├── driver_setup.py          # WebDriver init (Edge/Chrome switchable)
│   ├── login.py                 # SSO login
│   ├── tile1_confirmation.py    # Freight Orders for Confirmation
│   ├── tile2_reporting.py       # Freight Orders for Reporting
│   ├── tile3_invoicing.py       # Invoice Freight Documents
│   ├── tile4_pod_upload.py      # Manage Freight Execution (POD)
│   ├── excel_reader.py          # reads SAP_bot.xlsx
│   └── utils.py                 # shared helpers (wait, screenshot, retry)
├── logs/                        # auto-created, timestamped log files
└── screenshots/                 # auto-created, error/step screenshots
```

---

## Config File (`config.yaml`)

```yaml
sap_url: "https://your-sap-portal-url.com"
username: "your@email.com"
password: "yourpassword"          # or prompt at runtime for security
browser: "edge"                   # "edge" or "chrome"
excel_path: "SAP_bot.xlsx"
pod_base_path: "C:\\Users\\jean\\Downloads\\"
autonomous_interval_minutes: 15   # how often tiles 1 & 2 auto-run
dry_run: false                    # SAFETY: if true, bot navigates but never clicks Submit/Confirm
```

---

## Login Module

- Navigate to SAP portal URL
- Enter SSO email + password (no MFA)
- Wait for home page / tile dashboard to confirm successful login
- On failure: log error, take screenshot, exit with clear message

---

## Tile 1 — Freight Orders for Confirmation (`tile1_confirmation.py`)

**Goal:** Confirm all "New" freight orders. Runs autonomously.

**Steps:**
1. Navigate to "Freight Orders for Confirmation" tile from home
2. Apply filter: **Freight Order Status = "New"**
3. Navigate to the **"All"** tab
4. Enter a scroll-and-select loop:
   - Scroll to bottom of visible list
   - Check if new items loaded; if yes, keep scrolling until no more items load (i.e., full list is rendered)
   - Check the **"select all" checkbox** (header checkbox)
   - Click **"Confirm"** button (bottom right)
   - Wait for confirmation to process
   - Refresh / re-check if any "New" items remain
   - Repeat until zero "New" items remain
5. Log count of items confirmed per run

**Edge cases:**
- If 0 items: log "nothing to confirm" and exit tile cleanly
- If Confirm button is grayed out or missing: log warning + screenshot
- Handle SAP's lazy-loading list (only ~20 items visible at a time before scrolling)

---

## Tile 2 — Freight Orders for Reporting (`tile2_reporting.py`)

**Goal:** For every item in "Events to Report", copy Planned Time into Final Time for every stop. Runs autonomously.

**Steps:**
1. Navigate to "Freight Orders for Reporting" tile
2. Go to **"Events to Report"** tab
3. For each freight order row in the list:
   a. Click **"Report Final Time"** button on the row (opens the per-order reporting page)
   b. On the reporting page, identify all stops (there may be 2 or more)
   c. For each stop's event row:
      - Read the **Planned Time** value (e.g., `Mar 17, 2026, 12:00 AM EST`)
      - Parse out the **date/time portion** (e.g., `Mar 17, 2026, 12:00 AM`) and **timezone label** (e.g., `EST`)
      - Map timezone label to the correct UTC offset for the dropdown:
        - `EST` → `UTC-5`
        - `EDT` → `UTC-4`
        - `CST` → `UTC-6`
        - `CDT` → `UTC-5`
        - *(extend as needed)*
      - Click **"Report Final Time"** in the Action column for that stop's event
      - In the popup:
        - Enter the parsed date/time string into the **"Final Time"** field
        - Set the **"Time Zone"** dropdown to the mapped UTC offset
        - Leave "Reason for Delay" and "Additional Details" blank
        - Click **"Report"**
      - Wait for popup to close / success confirmation
   d. Repeat for all remaining stops in this freight order
   e. Navigate back to the Events to Report list
4. Repeat for all items in the list
5. Handle pagination / lazy loading same as Tile 1 (scroll to load all, or process page by page)
6. Log each freight order + stop processed

**Edge cases:**
- Items with more than 2 stops: bot must handle N stops generically, not assume exactly 2
- Timezone not in the mapping table: log a warning, take screenshot, skip that item (do not guess)
- Popup fails to open: log + screenshot + continue to next item

---

## Excel File Schema (`SAP_bot.xlsx`)

| Col | Header | Description |
|-----|--------|-------------|
| A | Document_1 | Freight document number, leg 1 |
| B | Invoice | Invoice number to enter |
| C | Has_Charges | Flag — whether extra charges apply (blank = no, any value = yes) |
| D | Charge_Amount | Dollar amount of Waiting Charges for leg 1 |
| E | POD_Base_Path | Base directory path e.g. `C:\Users\jean\Downloads\` |
| F | POD_Filename | Filename only e.g. `VIM64D` (no extension — bot appends `.pdf`) |
| G–J | *(TBD)* | Reserved / future use |
| K | Document_2 | Freight document number, leg 2 (blank if single-leg trip) |

**File path construction:** Bot concatenates col E + col F + `.pdf` to produce the full absolute path (e.g., `C:\Users\jean\Downloads\VIM64D.pdf`). Col E is the folder path only — change this column if files move to a new folder. Col F is the bare filename only, no extension. The bot always appends `.pdf` — never include the extension in col F.

> ⚠️ **Claude Code should validate that the Excel has the expected headers on startup and exit with a clear error if columns are missing or misnamed.**

> ⚠️ **Open item:** Does the Excel need a separate column for leg 2's charge amount (when a collective invoice has 2 legs each with their own charge)? Confirm before building the charges section of Tile 3.

---

## Tile 3 — Invoice Freight Documents (`tile3_invoicing.py`)

**Goal:** For each row in `SAP_bot.xlsx`, filter for freight document(s), create invoice, enter invoice number, add charges if applicable, and submit. **Human-gated.**

**Steps:**
1. Read all rows from `SAP_bot.xlsx`
2. Navigate to "Invoice Freight Documents" tile
3. For each row:

   **Filtering:**
   - If `Document_2` (col K) is **blank** → single invoice:
     - Enter `Document_1` in the Freight Document filter, press Enter
     - One item appears; select its checkbox
     - Click **"Create Invoice"**
   - If `Document_2` is **present** → collective invoice:
     - Click the expand icon on the Freight Document filter field
     - In the popup: set first condition to "contains" + `Document_1`, click "Add Condition", set second to "contains" + `Document_2`, click OK
     - Two items appear; select both checkboxes
     - Click **"Create Collective Invoice"**

   **If invoice page opens successfully:**
   - Go to **"Invoice Details"** tab
   - Enter value from col B into the **"Invoice:"** text field

   **If charges apply** (col C is non-blank):
   - Go to **"Charges"** tab
   - Click **"Add"** → select **"Charge"** from dropdown
   - A new row appears; click the blank field in the first (category) column
   - In the charge category popup, scroll to and select **"Waiting Charges"**
   - In the **"Rate Amount/Unit"** column, enter the amount from col D
   - If `Document_2` exists and leg 2 also has charges, add a second charge row the same way
     *(confirm whether a separate Excel column is needed for leg 2 charge amount)*

   - Click **"Submit"** ← ⚠️ **This submits the live invoice. Dry-run mode must skip this click.**

   **If invoice page did NOT open (went to Drafts instead):**
   - Log the freight document number as "Draft — needs manual review"
   - Continue to next row

4. After each submission, log: freight document(s), invoice number, charges applied, result
5. End-of-run summary: X submitted, Y failed/drafted

---

## Tile 4 — Manage Freight Execution / POD Upload (`tile4_pod_upload.py`)

**Goal:** For each freight order, upload the correct PDF to every "Proof of..." window across all stops. **Human-gated.**

**Steps:**
1. Navigate to "Manage Freight Execution" tile
2. For each item in the list:
   - Click on a **middle column cell** (e.g., "Reporting Status") to open the item — do NOT click the radio button
   - The item detail page shows stops (at minimum: stop 1 pick-up, stop 2 delivery)
   - For each stop:
     - Expand the stop if collapsed
     - Find the window whose header **starts with "Proof of "** (e.g., "Proof of Pick-Up", "Proof of Delivery")
     - Click the **"Report"** button in that window
     - In the popup, click **"Browse..."**
     - Construct file path: col E (base path) + col F (filename) + `.pdf`
     - Upload that file
     - Confirm/close the popup
   - The **same PDF** is uploaded to every "Proof of..." window across all stops for that freight order
3. Navigate back to list, proceed to next item
4. Log each upload: freight order + stop + filename

---

## Safety & Testing Strategy

### Dry-Run Mode
- Set `dry_run: true` in `config.yaml`
- Bot navigates fully, fills in all fields, logs exactly what it *would* do
- **Skips clicking:** Confirm, Submit, Report (final popup), and Browse/upload
- Prints a human-readable action plan: *"Would submit invoice INV-001 for documents 6100181820 + 6100184143 with charge $45.00"*
- Allows full visual verification without touching any live data

### Step-Through Mode (Tiles 3 & 4)
- Bot pauses before every destructive action and prompts:
  *"About to submit invoice INV-001. Press Enter to continue, type 'skip' to skip this item, or 'quit' to stop."*
- Enables row-by-row human sign-off during early live testing

### Screenshot on Every Action
- Bot takes a screenshot before and after every click on Confirm / Submit / Report
- Saved to `screenshots/` with timestamp + freight document number in filename
- Provides an audit trail if anything goes wrong

### Logging
- Every action logged with timestamp: freight order number, action taken, result
- Errors logged with full Python traceback + screenshot filename
- End-of-run summary: items processed, items skipped, items errored

### Recommended Testing Sequence
1. **Dry-run on all 4 tiles** — verify navigation works, fields are found, file paths resolve
2. **Live test Tile 1 only** — Confirmation is low-risk (customer already sent these orders)
3. **Live test Tile 2 only** — Timestamp reporting is low-risk (just logging times)
4. **Step-through live test Tile 4** — Review each POD upload manually before proceeding
5. **Step-through live test Tile 3 with 1 invoice** — Verify the invoice page looks correct before clicking Submit
6. **Small batch live test Tile 3** — Run 3–5 invoices, verify in SAP manually
7. **Full autonomous run**

---

## Hard Rules for Claude Code

- Never hardcode credentials — read from `config.yaml`, or prompt at runtime
- Never click Submit/Confirm/Report when `dry_run: true`
- Never use `time.sleep()` — always use Selenium `WebDriverWait` + `expected_conditions`
- Never silently swallow exceptions — log everything, continue to next item, never crash silently
- Browser (Edge vs Chrome) must be switchable via config only, never hardcoded

---

## Open Items to Resolve During Build

1. **Leg 2 charges:** Is there a separate Excel column for leg 2's charge amount when a collective invoice has 2 legs with separate charges? Or does col D cover both legs?
2. ~~**POD filename extension:** Resolved — bot always appends `.pdf`. Col E = folder path (edit this if files move). Col F = bare filename, no extension, never changes.~~
3. **Excel header row:** Confirm exact column header names so the bot can validate on startup.
4. **Autonomous scheduling:** Should Tiles 1 & 2 loop continuously while running, or be invoked on a fixed interval by Windows Task Scheduler?
5. **Draft invoices:** Placeholder only — log and skip for now, handle in a future module.
