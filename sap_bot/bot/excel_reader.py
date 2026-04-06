"""Read and validate SAP_bot.xlsx."""

import os
import logging
from dataclasses import dataclass

import openpyxl

log = logging.getLogger(__name__)

# Expected column layout (1-indexed column positions)
EXPECTED_COLUMNS = {
    1: "Document_1",
    2: "Invoice",
    3: "Charges",        # charge type for leg 1 (blank = no charges)
    4: "Amount",         # computed charge amount for leg 1
    5: "File",           # POD base path
    6: "File",           # POD filename (no extension)
    7: "Charge_Hhrs",   # hours for leg 1 charge calc
    # 8-10: reserved / empty
    11: "Document_2",
    12: "Charges",       # charge type for leg 2
    13: "Amount",        # computed charge amount for leg 2
    14: "Charge_Hhrs",  # hours for leg 2 charge calc
}


@dataclass
class InvoiceRow:
    """One row from SAP_bot.xlsx."""
    row_number: int
    document_1: str
    invoice: str
    charge_type_1: str | None       # e.g. "Waiting Charges" or None
    charge_amount_1: float | None   # computed dollar amount
    pod_base_path: str
    pod_filename: str               # bare filename, no extension
    charge_hours_1: float | None
    document_2: str | None
    charge_type_2: str | None
    charge_amount_2: float | None
    charge_hours_2: float | None

    @property
    def is_collective(self) -> bool:
        return bool(self.document_2)

    @property
    def has_charges_leg1(self) -> bool:
        return bool(self.charge_type_1)

    @property
    def has_charges_leg2(self) -> bool:
        return bool(self.charge_type_2)

    @property
    def pod_full_path(self) -> str:
        return os.path.join(self.pod_base_path, self.pod_filename + ".pdf")


def validate_headers(ws) -> list[str]:
    """Check that expected columns exist. Returns list of error messages (empty = OK)."""
    errors = []
    for col_num, expected_name in EXPECTED_COLUMNS.items():
        actual = ws.cell(row=1, column=col_num).value
        if actual is None:
            errors.append(f"Column {col_num} is empty — expected '{expected_name}'")
        elif str(actual).strip() != expected_name:
            errors.append(
                f"Column {col_num}: expected '{expected_name}', got '{actual}'"
            )
    return errors


def _cell_val(ws, row: int, col: int):
    """Get cell value, returning None for blank cells."""
    v = ws.cell(row=row, column=col).value
    if v is None:
        return None
    if isinstance(v, str) and v.strip() == "":
        return None
    return v


def read_excel(path: str) -> list[InvoiceRow]:
    """Read SAP_bot.xlsx and return a list of InvoiceRow objects.

    Raises SystemExit if file is missing or headers are invalid.
    """
    if not os.path.isfile(path):
        log.error("Excel file not found: %s", path)
        raise SystemExit(f"Excel file not found: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)  # data_only=True reads computed formula values
    ws = wb.active

    # Validate headers
    header_errors = validate_headers(ws)
    if header_errors:
        for err in header_errors:
            log.error("Excel header error: %s", err)
        raise SystemExit(
            "Excel header validation failed:\n  " + "\n  ".join(header_errors)
        )

    rows = []
    for row_num in range(2, ws.max_row + 1):
        doc1 = _cell_val(ws, row_num, 1)
        if doc1 is None:
            continue  # skip blank rows

        charge_amt_1 = _cell_val(ws, row_num, 4)
        charge_amt_2 = _cell_val(ws, row_num, 13)
        charge_hrs_1 = _cell_val(ws, row_num, 7)
        charge_hrs_2 = _cell_val(ws, row_num, 14)

        inv = InvoiceRow(
            row_number=row_num,
            document_1=str(doc1).strip(),
            invoice=str(_cell_val(ws, row_num, 2) or "").strip(),
            charge_type_1=_cell_val(ws, row_num, 3),
            charge_amount_1=float(charge_amt_1) if charge_amt_1 is not None else None,
            pod_base_path=str(_cell_val(ws, row_num, 5) or "").strip(),
            pod_filename=str(_cell_val(ws, row_num, 6) or "").strip(),
            charge_hours_1=float(charge_hrs_1) if charge_hrs_1 is not None else None,
            document_2=str(_cell_val(ws, row_num, 11)).strip() if _cell_val(ws, row_num, 11) else None,
            charge_type_2=_cell_val(ws, row_num, 12),
            charge_amount_2=float(charge_amt_2) if charge_amt_2 is not None else None,
            charge_hours_2=float(charge_hrs_2) if charge_hrs_2 is not None else None,
        )
        rows.append(inv)

    log.info("Loaded %d rows from %s", len(rows), path)
    wb.close()
    return rows
