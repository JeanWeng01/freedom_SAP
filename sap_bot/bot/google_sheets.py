"""Google Sheets integration for tiles 3 & 4.

Sheet layout:
  "To Do" tab — items for bot to process
    A: Document_1    B: Invoice       C: POD_Filename   D: Charge_Hrs_1
    E: Charge_Type_1 F: Amount_1      G: (blank spacer)
    H: Document_2    I: Charge_Hrs_2  J: Charge_Type_2  K: Amount_2

  "Status" tab — completed/errored items (same columns + L: Processed_At, M: Status)
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
COL_DOC1 = 0        # A
COL_INVOICE = 1      # B
COL_POD_FILENAME = 2 # C
COL_CHARGE_HRS1 = 3  # D
COL_CHARGE_TYPE1 = 4 # E
COL_AMOUNT1 = 5      # F
# G is blank spacer
COL_DOC2 = 7         # H
COL_CHARGE_HRS2 = 8  # I
COL_CHARGE_TYPE2 = 9 # J
COL_AMOUNT2 = 10     # K


@dataclass
class InvoiceRow:
    """One row from the Google Sheet To Do tab."""
    sheet_row: int  # 1-based row number in the sheet (for writing back)
    raw_values: list  # original row values
    document_1: str
    invoice: str
    pod_filename: str
    charge_hours_1: float | None
    charge_type_1: str | None
    charge_amount_1: float | None
    document_2: str | None
    charge_hours_2: float | None
    charge_type_2: str | None
    charge_amount_2: float | None

    @property
    def is_collective(self) -> bool:
        return bool(self.document_2)

    @property
    def has_charges_leg1(self) -> bool:
        return bool(self.charge_type_1) and bool(self.charge_hours_1)

    @property
    def has_charges_leg2(self) -> bool:
        return bool(self.charge_type_2) and bool(self.charge_hours_2)

    @property
    def pod_full_filename(self) -> str:
        return self.pod_filename + ".pdf" if self.pod_filename else ""


def _get_credentials():
    """Load Google service account credentials from env var or file."""
    # Try JSON string from env var first (Railway)
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_json:
        info = json.loads(creds_json)
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    # Try file path (local development)
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
        range=f"'{TODO_TAB}'!A:K",
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

        charge_hrs1 = _float_val(row, COL_CHARGE_HRS1)
        charge_amt1 = _float_val(row, COL_AMOUNT1)
        # Compute amount from hours if formula result not available
        if charge_amt1 is None and charge_hrs1 is not None:
            charge_amt1 = round(charge_hrs1 * HOURLY_RATE, 2)

        charge_hrs2 = _float_val(row, COL_CHARGE_HRS2)
        charge_amt2 = _float_val(row, COL_AMOUNT2)
        if charge_amt2 is None and charge_hrs2 is not None:
            charge_amt2 = round(charge_hrs2 * HOURLY_RATE, 2)

        item = InvoiceRow(
            sheet_row=i,
            raw_values=row,
            document_1=doc1,
            invoice=_cell_val(row, COL_INVOICE) or "",
            pod_filename=_cell_val(row, COL_POD_FILENAME) or "",
            charge_hours_1=charge_hrs1,
            charge_type_1=_cell_val(row, COL_CHARGE_TYPE1),
            charge_amount_1=charge_amt1,
            document_2=_cell_val(row, COL_DOC2),
            charge_hours_2=charge_hrs2,
            charge_type_2=_cell_val(row, COL_CHARGE_TYPE2),
            charge_amount_2=charge_amt2,
        )
        items.append(item)

    log.info("Read %d items from To Do tab", len(items))
    return items


def move_to_status(item: InvoiceRow, status: str):
    """Move a row from 'To Do' to 'Status' tab with timestamp and status.

    1. Append the row to Status tab (with Processed_At and Status columns)
    2. Delete the row from To Do tab
    """
    service = _get_sheets_service()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Reconstruct row from parsed fields to guarantee column alignment
    # A: Document_1, B: Invoice, C: POD_Filename, D: Charge_Hrs_1,
    # E: Charge_Type_1, F: Amount_1, G: (blank), H: Document_2,
    # I: Charge_Hrs_2, J: Charge_Type_2, K: Amount_2,
    # L: Processed_At, M: Status
    row_data = [
        item.document_1,
        item.invoice,
        item.pod_filename,
        item.charge_hours_1 if item.charge_hours_1 is not None else "",
        item.charge_type_1 or "",
        item.charge_amount_1 if item.charge_amount_1 is not None else "",
        "",  # G: blank spacer
        item.document_2 or "",
        item.charge_hours_2 if item.charge_hours_2 is not None else "",
        item.charge_type_2 or "",
        item.charge_amount_2 if item.charge_amount_2 is not None else "",
        timestamp,  # L: Processed_At
        status,     # M: Status
    ]

    # Append to Status tab
    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=f"'{STATUS_TAB}'!A:M",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [row_data]},
    ).execute()

    # Delete the row from To Do tab
    # Need the sheet's gid (sheetId) — get it from the spreadsheet metadata
    try:
        meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
        todo_sheet_id = None
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == TODO_TAB:
                todo_sheet_id = sheet["properties"]["sheetId"]
                break

        if todo_sheet_id is not None:
            # Re-read current To Do to find the right row index
            # (rows may have shifted if previous items were deleted)
            current = service.spreadsheets().values().get(
                spreadsheetId=SHEET_ID,
                range=f"'{TODO_TAB}'!A:A",
            ).execute()
            current_rows = current.get("values", [])

            # Find the row with matching Document_1
            delete_index = None
            for idx, r in enumerate(current_rows):
                if idx == 0:
                    continue  # skip header
                if r and str(r[0]).strip() == item.document_1:
                    delete_index = idx
                    break

            if delete_index is not None:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SHEET_ID,
                    body={"requests": [{
                        "deleteDimension": {
                            "range": {
                                "sheetId": todo_sheet_id,
                                "dimension": "ROWS",
                                "startIndex": delete_index,
                                "endIndex": delete_index + 1,
                            }
                        }
                    }]},
                ).execute()
                log.info("Moved row %d (doc %s) to Status tab: %s",
                         item.sheet_row, item.document_1, status)
            else:
                log.warning("Could not find row to delete for doc %s", item.document_1)
        else:
            log.warning("Could not find To Do sheet ID for row deletion")
    except Exception as e:
        log.error("Error deleting row from To Do: %s", e)
        # Row was already appended to Status — not critical if delete fails


def mark_error(item: InvoiceRow, error_msg: str):
    """Move an item to Status tab with an error message."""
    move_to_status(item, f"error: {error_msg}")
