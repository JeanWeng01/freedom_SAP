"""Google Sheets integration for tiles 3 & 4.

Sheet layout:
  "To Do" tab — items for bot to process
    A: Document_1    B: Document_2    C: Invoice       D: POD_Filename
    E: Charge_Hrs_1  F: Charge_Hrs_2  G: Pause (1 = skip submit, save instead)

  "Status" tab — completed items (same A-F columns + processing info)
    A-F: same as To Do
    G: Processed_At_1   (tile 3 timestamp)
    H: Processed_At_2   (tile 4 timestamp)
    I: POD_Uploaded_1    ("done" or error)
    J: POD_Uploaded_2    ("done" or error)
    K: Status            ("done" or "error: ...")
"""

import os
import json
import logging
from datetime import datetime
from dataclasses import dataclass

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

log = logging.getLogger(__name__)

SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "1disuW1t_BzqJpoz_taz04sFKViRLuEM7xY2xNvjQiRQ")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
HOURLY_RATE = 42.84

TODO_TAB = "To Do"
STATUS_TAB = "Status"

# Column indices (0-based) in the To Do tab
COL_DOC1 = 0         # A
COL_DOC2 = 1         # B
# Col C (2): human notes — bot IGNORES this column
COL_INVOICE = 3      # D
COL_POD_FILENAME = 4 # E
COL_CHARGE_HRS1 = 5  # F
COL_CHARGE_HRS2 = 6  # G
COL_PAUSE = 7        # H
# Col I (8): local run heartbeat timestamp (single cell I2)
COL_TODO_STATUS = 9  # J — bot writes: "invoiced", "paused", "tile3_in_progress", etc.


@dataclass
class InvoiceRow:
    """One row from the Google Sheet To Do tab."""
    sheet_row: int  # 1-based row number in the sheet
    document_1: str
    document_2: str | None
    invoice: str
    pod_filename: str
    charge_hours_1: float | None
    charge_hours_2: float | None
    pause: bool  # True = skip submit, save instead
    notes: str = ""  # human notes from col C — bot ignores for processing but preserves when moving to Status

    @property
    def is_collective(self) -> bool:
        return bool(self.document_2)

    @property
    def has_charges_leg1(self) -> bool:
        return bool(self.charge_hours_1)

    @property
    def has_charges_leg2(self) -> bool:
        return bool(self.charge_hours_2)

    @property
    def charge_amount_1(self) -> float | None:
        if self.charge_hours_1:
            return round(self.charge_hours_1 * HOURLY_RATE, 2)
        return None

    @property
    def charge_amount_2(self) -> float | None:
        if self.charge_hours_2:
            return round(self.charge_hours_2 * HOURLY_RATE, 2)
        return None

    @property
    def charge_type_1(self) -> str | None:
        return "Waiting Charges" if self.has_charges_leg1 else None

    @property
    def charge_type_2(self) -> str | None:
        return "Waiting Charges" if self.has_charges_leg2 else None

    @property
    def pod_full_filename(self) -> str:
        """Returns first filename with .pdf suffix (for backwards compat)."""
        names = self.pod_filenames
        return names[0] if names else ""

    @property
    def pod_filenames(self) -> list[str]:
        """Parse POD_Filename cell into list of filenames with .pdf suffix.

        Supports comma-separated multiple files with optional spaces:
          'VIM04E'              → ['VIM04E.pdf']
          'VIM04E,VIM04E_2'     → ['VIM04E.pdf', 'VIM04E_2.pdf']
          'VIM04E, VIM04E_2'    → ['VIM04E.pdf', 'VIM04E_2.pdf']
        """
        if not self.pod_filename:
            return []
        # Accept both comma and semicolon as delimiters for backwards compat
        parts = [p.strip() for p in self.pod_filename.replace(";", ",").split(",")]
        return [p + ".pdf" for p in parts if p]


def _get_credentials():
    """Load Google service account credentials from env var or file."""
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_json:
        info = json.loads(creds_json)
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    creds_file = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
    if os.path.isfile(creds_file):
        return Credentials.from_service_account_file(creds_file, scopes=SCOPES)

    raise RuntimeError(
        "No Google credentials found. Set GOOGLE_SERVICE_ACCOUNT_JSON env var "
        "or place service_account.json in the sap_bot directory."
    )


def _get_sheets_service():
    creds = _get_credentials()
    return build("sheets", "v4", credentials=creds)


def _cell_val(row: list, index: int) -> str | None:
    """Safely get a cell value from a row list, returning None for blank."""
    if index >= len(row):
        return None
    v = row[index]
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    return str(v).strip()


def _float_val(row: list, index: int) -> float | None:
    """Safely get a float from a cell."""
    v = _cell_val(row, index)
    if v is None:
        return None
    try:
        return float(v)
    except ValueError:
        return None


def read_todo_items() -> list[InvoiceRow]:
    """Read all rows from the 'To Do' tab. Skips header row and blank rows."""
    service = _get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=f"'{TODO_TAB}'!A:J",
    ).execute()

    rows = result.get("values", [])
    if len(rows) <= 1:
        log.info("No data rows in To Do tab")
        return []

    items = []
    for i, row in enumerate(rows[1:], start=2):  # skip header, 1-based row numbers
        doc1 = _cell_val(row, COL_DOC1)
        if not doc1:
            continue  # skip blank rows

        pause_val = _cell_val(row, COL_PAUSE)
        is_paused = pause_val == "1" or (pause_val and pause_val.lower() == "true")

        item = InvoiceRow(
            sheet_row=i,
            document_1=doc1,
            document_2=_cell_val(row, COL_DOC2),
            invoice=_cell_val(row, COL_INVOICE) or "",
            pod_filename=_cell_val(row, COL_POD_FILENAME) or "",
            charge_hours_1=_float_val(row, COL_CHARGE_HRS1),
            charge_hours_2=_float_val(row, COL_CHARGE_HRS2),
            pause=is_paused,
            notes=_cell_val(row, 2) or "",  # col C human notes
        )
        items.append(item)

    log.info("Read %d items from To Do tab (%d paused)",
             len(items), sum(1 for x in items if x.pause))
    return items


def move_to_status(item: InvoiceRow, status: str,
                   processed_at_1: str = "", processed_at_2: str = "",
                   pod_uploaded_1: str = "", pod_uploaded_2: str = ""):
    """Move a row from 'To Do' to 'Status' tab with processing info.

    Status tab columns:
      A: Document_1  B: Document_2  C: Notes  D: Invoice  E: POD_Filename
      F: Charge_Hrs_1  G: Charge_Hrs_2
      H: Processed_At_1  I: Processed_At_2
      J: POD_Uploaded_1  K: POD_Uploaded_2
      L: Status
    """
    service = _get_sheets_service()

    row_data = [
        item.document_1,
        item.document_2 or "",
        item.notes,                                                           # C: Notes
        item.invoice,                                                         # D: Invoice
        item.pod_filename,                                                    # E: POD_Filename
        item.charge_hours_1 if item.charge_hours_1 is not None else "",       # F
        item.charge_hours_2 if item.charge_hours_2 is not None else "",       # G
        processed_at_1,                                                       # H
        processed_at_2,                                                       # I
        pod_uploaded_1,                                                       # J
        pod_uploaded_2,                                                       # K
        status,                                                               # L
    ]

    # Append to Status tab
    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=f"'{STATUS_TAB}'!A:L",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [row_data]},
    ).execute()

    # Apply color formatting to the new row
    _apply_status_colors(service, status, pod_uploaded_1, pod_uploaded_2)

    # Delete the row from To Do tab
    _delete_todo_row(service, item.document_1)

    log.info("Moved doc %s to Status tab: %s", item.document_1, status)


def _apply_status_colors(service, status: str, pod1: str, pod2: str):
    """Apply green (success) or red (error) font color to the last row in Status tab."""
    try:
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        status_sheet_id = None
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == STATUS_TAB:
                status_sheet_id = sheet["properties"]["sheetId"]
                break

        if status_sheet_id is None:
            return

        # Get the row count to find the last row
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"'{STATUS_TAB}'!A:A",
        ).execute()
        last_row = len(result.get("values", []))  # 1-based, includes header

        if last_row < 2:
            return

        row_index = last_row - 1  # 0-based for batchUpdate

        # Determine colors for each cell (H through L — cols shifted right by 1 due to notes col)
        requests = []
        cells_to_color = [
            (7, True),                                          # H: Processed_At_1
            (8, True),                                          # I: Processed_At_2
            (9, not pod1.startswith("error")),                  # J: POD_Uploaded_1
            (10, not pod2.startswith("error")),                 # K: POD_Uploaded_2
            (11, not status.startswith("error")),               # L: Status
        ]

        for col_index, is_success in cells_to_color:
            color = {"red": 0.0, "green": 0.6, "blue": 0.0} if is_success else {"red": 0.8, "green": 0.0, "blue": 0.0}
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": status_sheet_id,
                        "startRowIndex": row_index,
                        "endRowIndex": row_index + 1,
                        "startColumnIndex": col_index,
                        "endColumnIndex": col_index + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "foregroundColor": color,
                                "bold": False,
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat",
                }
            })

        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SHEET_ID,
                body={"requests": requests},
            ).execute()

    except Exception as e:
        log.warning("Could not apply status colors: %s", e)


def _delete_todo_row(service, document_1: str):
    """Delete a row from To Do tab by matching Document_1."""
    try:
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        todo_sheet_id = None
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == TODO_TAB:
                todo_sheet_id = sheet["properties"]["sheetId"]
                break

        if todo_sheet_id is None:
            return

        current = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!A:A",
        ).execute()
        current_rows = current.get("values", [])

        for idx, r in enumerate(current_rows):
            if idx == 0:
                continue  # skip header
            if r and str(r[0]).strip() == document_1:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SHEET_ID,
                    body={"requests": [{
                        "deleteDimension": {
                            "range": {
                                "sheetId": todo_sheet_id,
                                "dimension": "ROWS",
                                "startIndex": idx,
                                "endIndex": idx + 1,
                            }
                        }
                    }]},
                ).execute()
                log.info("Deleted row %d (doc %s) from To Do", idx, document_1)
                return

        log.warning("Could not find row to delete for doc %s", document_1)
    except Exception as e:
        log.error("Error deleting row from To Do: %s", e)


def mark_error(item: InvoiceRow, error_msg: str):
    """Move an item to Status tab with an error message."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    move_to_status(item,
                   status=f"error: {error_msg}",
                   processed_at_1=timestamp)


def is_local_run_active(max_age_seconds: int = 1800) -> bool:
    """Check if a local bot run is currently active (by reading a lock cell).

    Local runs write a timestamp to 'To Do'!I2. If that timestamp is fresh
    (within max_age_seconds), consider a local run active.

    max_age_seconds: how recently the local run must have heartbeat to be
        considered active. Default 30 min.
    """
    service = _get_sheets_service()
    try:
        res = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!I2",
        ).execute()
        vals = res.get("values", [])
        if not vals or not vals[0]:
            return False
        ts_str = str(vals[0][0]).strip()
        if not ts_str:
            return False
        from datetime import datetime
        try:
            ts = datetime.strptime(ts_str, "%Y-%m-%d %H:%M:%S")
            age = (datetime.now() - ts).total_seconds()
            return age < max_age_seconds
        except ValueError:
            return False
    except Exception as e:
        log.warning("Could not check local run status: %s", e)
        return False


def write_local_run_heartbeat():
    """Mark that a local run is active by writing current timestamp to 'To Do'!I2."""
    service = _get_sheets_service()
    try:
        from datetime import datetime
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!I2",
            valueInputOption="RAW",
            body={"values": [[ts]]},
        ).execute()
        log.info("Wrote local run heartbeat: %s", ts)
    except Exception as e:
        log.warning("Could not write local run heartbeat: %s", e)


def clear_local_run_heartbeat():
    """Clear the local run heartbeat cell — called when local run ends."""
    service = _get_sheets_service()
    try:
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!I2",
            valueInputOption="RAW",
            body={"values": [[""]]},
        ).execute()
        log.info("Cleared local run heartbeat")
    except Exception as e:
        log.warning("Could not clear local run heartbeat: %s", e)


def read_todo_status(document_1: str) -> str:
    """Read the current col H status for a doc from To Do tab."""
    service = _get_sheets_service()
    try:
        doc_col = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!A:J",
        ).execute().get("values", [])
        for idx, r in enumerate(doc_col):
            if idx == 0:
                continue
            if r and str(r[0]).strip() == document_1:
                return r[9] if len(r) > 9 else ""
    except Exception as e:
        log.warning("Could not read To Do status for %s: %s", document_1, e)
    return ""


def _write_todo_status(document_1: str, status_text: str, color: str = "auto"):
    """Write a status value to column I of a row in To Do.

    color: 'green' | 'red' | 'yellow' | 'auto' (auto picks based on status text)
    """
    service = _get_sheets_service()
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!A:A",
        ).execute()
        rows = result.get("values", [])

        target_idx = None
        for idx, r in enumerate(rows):
            if idx == 0:
                continue
            if r and str(r[0]).strip() == document_1:
                target_idx = idx
                break

        if target_idx is None:
            log.warning("Could not find doc %s in To Do to update status", document_1)
            return

        # Write the value to col I (status column)
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"'{TODO_TAB}'!J{target_idx + 1}",
            valueInputOption="RAW",
            body={"values": [[status_text]]},
        ).execute()
        log.info("Set To Do status for doc %s: '%s'", document_1, status_text)

        # Determine color
        if color == "auto":
            txt = status_text.lower()
            if "error" in txt or "failed" in txt:
                color = "red"
            elif "in_progress" in txt or "processing" in txt or "running" in txt:
                color = "yellow"
            else:
                color = "green"

        color_rgb = {
            "green": {"red": 0.0, "green": 0.6, "blue": 0.0},
            "red":   {"red": 0.8, "green": 0.0, "blue": 0.0},
            "yellow":{"red": 0.85, "green": 0.65, "blue": 0.0},
        }.get(color, {"red": 0.0, "green": 0.0, "blue": 0.0})

        # Apply font color (no bold)
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        todo_sheet_id = None
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == TODO_TAB:
                todo_sheet_id = sheet["properties"]["sheetId"]
                break
        if todo_sheet_id is not None:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SHEET_ID,
                body={"requests": [{
                    "repeatCell": {
                        "range": {
                            "sheetId": todo_sheet_id,
                            "startRowIndex": target_idx,
                            "endRowIndex": target_idx + 1,
                            "startColumnIndex": 9,  # J (Status column)
                            "endColumnIndex": 10,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {
                                    "foregroundColor": color_rgb,
                                    "bold": False,
                                }
                            }
                        },
                        "fields": "userEnteredFormat.textFormat",
                    }
                }]},
            ).execute()
    except Exception as e:
        log.error("Error writing To Do status: %s", e)


def mark_tile3_done(item: InvoiceRow):
    """Mark tile 3 (invoicing) as done — writes 'invoiced' to col H."""
    _write_todo_status(item.document_1, "invoiced")
    log.info("Tile 3 done for doc %s — marked 'invoiced', stays in To Do for tile 4", item.document_1)


def mark_tile3_done_with_note(item: InvoiceRow, note: str):
    """Mark tile 3 as done but with a note (e.g. one leg skipped)."""
    _write_todo_status(item.document_1, f"invoiced ({note})")
    log.info("Tile 3 done for doc %s with note: %s", item.document_1, note)


def mark_tile3_paused(item: InvoiceRow):
    """Mark tile 3 as paused — writes 'paused' to col H."""
    _write_todo_status(item.document_1, "paused")


def mark_fully_done(item: InvoiceRow, pod_uploaded_1: str = "done",
                    pod_uploaded_2: str = "done"):
    """Mark both tile 3 and tile 4 as done — move to Status tab."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    move_to_status(item,
                   status="done",
                   processed_at_1=timestamp,
                   processed_at_2=timestamp,
                   pod_uploaded_1=pod_uploaded_1,
                   pod_uploaded_2=pod_uploaded_2)


def sort_paused_rows_to_top():
    """Sort To Do tab so paused rows (col G = 1) appear at the top."""
    service = _get_sheets_service()
    try:
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        todo_sheet_id = None
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == TODO_TAB:
                todo_sheet_id = sheet["properties"]["sheetId"]
                break

        if todo_sheet_id is None:
            return

        # Sort by column G (Pause) descending — 1s go to top
        service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"requests": [{
                "sortRange": {
                    "range": {
                        "sheetId": todo_sheet_id,
                        "startRowIndex": 1,  # skip header
                        "startColumnIndex": 0,
                        "endColumnIndex": 10,  # A through J
                    },
                    "sortSpecs": [{
                        "dimensionIndex": COL_PAUSE,  # col H
                        "sortOrder": "DESCENDING",
                    }],
                }
            }]},
        ).execute()
        log.info("Sorted To Do: paused rows moved to top")
    except Exception as e:
        log.warning("Could not sort paused rows: %s", e)
