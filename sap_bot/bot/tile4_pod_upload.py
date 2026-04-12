"""Tile 4 — Manage Freight Execution / POD Upload.

Goal: For each freight order, upload the correct PDF to every "Proof of..." window
across all stops. Human-gated.

Based on screenshots: the detail page has stops (Stop 1, Stop 2, ...) each with
"Proof of Pick-Up" or "Proof of Delivery" sections. Click "Expand All" first,
then for each stop find the "Proof of..." section and click Report → upload PDF.
"""

import os
import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_elements, wait_until_gone,
    scroll_to_load_all, take_screenshot, destructive_action, click_tile,
)
from bot.excel_reader import InvoiceRow

log = logging.getLogger(__name__)

TILE_NAME = "Manage Freight Execution"

# ── Selectors (refined from manage_freight_execution screenshots) ───────────

# ── List page selectors ────────────────────────────────────────────────────

# Tabs on the list page
DOCUMENTS_FOR_REPORTING_TAB = (By.XPATH,
    "//*[contains(text(),'Documents for Reporting')]/.."
)

# Freight document rows in the list table
ORDER_ROWS = (By.CSS_SELECTOR,
    "table tbody tr"
)

# ── Detail page selectors ──────────────────────────────────────────────────

# "Expand All" button (top-right of detail page)
EXPAND_ALL_BTN = (By.XPATH,
    "//button[.//bdi[text()='Expand All'] or .//span[text()='Expand All']]"
    " | //a[contains(text(),'Expand All')]"
)

# Stop sections — "Stop 1 - ...", "Stop 2 - ..." headers
STOP_HEADERS = (By.XPATH,
    "//span[starts-with(text(),'Stop ') and contains(text(),' - ')]"
)

# "Proof of Pick-Up" or "Proof of Delivery" section headers
PROOF_HEADERS = (By.XPATH,
    "//span[starts-with(text(),'Proof of ')]"
)

# "Report" button next to each "Proof of..." section
# In the screenshot, each Proof section has its own "Report" link/button
PROOF_REPORT_BTN = (By.XPATH,
    "//a[text()='Report'] | //button[.//bdi[text()='Report'] or .//span[text()='Report']]"
)

# ── Upload popup selectors ──────────────────────────────────────────────────

# File input for upload (hidden input[type=file] used by Browse button)
FILE_INPUT = (By.CSS_SELECTOR, "input[type='file']")

# Browse button in popup
BROWSE_BUTTON = (By.XPATH,
    "//button[.//bdi[text()='Browse...'] or .//span[text()='Browse...']]"
    " | //button[contains(@id,'browse') or contains(@id,'Browse')]"
)

# Upload / OK / Confirm button in popup
UPLOAD_CONFIRM_BTN = (By.XPATH,
    "//div[contains(@class,'sapMDialog')]//button[.//bdi[text()='OK']]"
    " | //div[contains(@class,'sapMDialog')]//button[.//bdi[text()='Upload']]"
    " | //div[contains(@class,'sapMDialog')]//button[.//bdi[text()='Report']]"
)

# Back / navigation
BACK_BUTTON = (By.CSS_SELECTOR,
    "button[title='Back'], .sapMNavBack, .sapUshellShellHeadItm[title='Back']"
)

# Information / Attachments tabs on detail page
ATTACHMENTS_TAB = (By.XPATH,
    "//*[contains(text(),'Attachments')]/.."
)


def navigate_to_tile(driver: WebDriver):
    """Click the Manage Freight Execution tile."""
    click_tile(driver, TILE_NAME)
    log.info("Tile 4 page loaded")


def click_into_order(driver: WebDriver, row_index: int) -> bool:
    """Click into a freight order from the list by clicking a middle column cell.

    DO NOT click the radio button — click a data cell like 'Current Status' or
    'Departure Location'.
    """
    try:
        rows = driver.find_elements(*ORDER_ROWS)
        if row_index >= len(rows):
            log.warning("Row index %d out of range (%d rows)", row_index, len(rows))
            return False

        row = rows[row_index]
        cells = row.find_elements(By.TAG_NAME, "td")
        # Click on a middle column (e.g. "Reporting Status" — roughly column 5-6)
        target_cell = cells[min(5, len(cells) - 1)] if len(cells) > 1 else row
        target_cell.click()
        log.info("Clicked into freight order at row %d", row_index)
        return True
    except Exception as e:
        log.error("Could not click into order row %d: %s", row_index, e)
        return False


def expand_all_stops(driver: WebDriver):
    """Click 'Expand All' to open all stop sections."""
    try:
        btn = wait_for_element(driver, *EXPAND_ALL_BTN, timeout=10, clickable=True)
        btn.click()
        log.info("Clicked 'Expand All'")
    except TimeoutException:
        log.info("'Expand All' button not found — stops may already be expanded")


@destructive_action("Upload POD {file_path} for {proof_label}")
def upload_pod_file(driver: WebDriver, file_path: str, *, proof_label: str = ""):
    """Upload a PDF file through the file input."""
    take_screenshot(driver, f"tile4_before_upload_{proof_label}")

    # Selenium can send the file path directly to input[type=file]
    try:
        file_input = driver.find_element(*FILE_INPUT)
        file_input.send_keys(file_path)
        log.info("Uploaded file via input: %s", file_path)
    except NoSuchElementException:
        log.error("File input not found for %s", proof_label)
        take_screenshot(driver, f"tile4_file_input_missing_{proof_label}")
        return False

    # Confirm/OK the upload
    try:
        confirm_btn = wait_for_element(driver, *UPLOAD_CONFIRM_BTN, timeout=10, clickable=True)
        confirm_btn.click()
        log.info("Confirmed upload for %s", proof_label)
    except TimeoutException:
        log.info("No confirm button — upload may have auto-completed for %s", proof_label)

    take_screenshot(driver, f"tile4_after_upload_{proof_label}")
    return True


def find_and_upload_proofs(driver: WebDriver, pod_path: str, doc_number: str,
                           *, dry_run: bool = False, step_through: bool = False) -> int:
    """Find all 'Proof of...' sections on the detail page and upload the PDF to each.

    Returns the number of successful uploads.
    """
    # Find all "Proof of ..." headers
    proof_headers = driver.find_elements(*PROOF_HEADERS)
    if not proof_headers:
        log.info("No 'Proof of...' sections found for %s", doc_number)
        return 0

    log.info("Found %d 'Proof of...' sections for %s", len(proof_headers), doc_number)

    # Find all Report buttons — each "Proof of..." section has one
    report_buttons = driver.find_elements(*PROOF_REPORT_BTN)
    # Filter to only the Report buttons that are near/within Proof sections
    # The Report buttons in the screenshot appear as links at the bottom of each proof section

    uploads = 0
    # We process proof sections by finding pairs of (proof_header, report_button)
    # Since the page structure shows Report buttons after each Proof section,
    # we match them positionally
    for i, proof_el in enumerate(proof_headers):
        proof_label = proof_el.text.strip()
        safe_label = proof_label.replace(" ", "_")[:30]
        log.info("Processing: %s (section %d/%d)", proof_label, i + 1, len(proof_headers))

        # Find the closest Report button to this proof section
        # Use the parent container to scope the search
        try:
            # Try to find Report button within the same parent section
            parent = proof_el.find_element(By.XPATH, "./ancestor::div[contains(@class,'sapUiForm') or contains(@class,'sapMPanel') or contains(@class,'sapUiVlt')][1]")
            report_btn = parent.find_element(By.XPATH,
                ".//a[text()='Report'] | .//button[.//bdi[text()='Report']]"
            )
        except NoSuchElementException:
            # Fallback: use positional matching with all Report buttons
            if i < len(report_buttons):
                report_btn = report_buttons[i]
            else:
                log.warning("No Report button found for '%s'", proof_label)
                continue

        # Click the Report button to open the upload dialog
        try:
            report_btn.click()
            log.info("Clicked Report for '%s'", proof_label)
        except Exception as e:
            log.error("Could not click Report for '%s': %s", proof_label, e)
            take_screenshot(driver, f"tile4_report_click_fail_{safe_label}")
            continue

        # Upload the file
        result = upload_pod_file(
            driver,
            pod_path,
            proof_label=f"{safe_label}_{doc_number}",
            dry_run=dry_run,
            step_through=step_through,
        )
        if result and result != "skipped":
            uploads += 1

    return uploads


def process_order(driver: WebDriver, row: InvoiceRow, row_index: int,
                  *, dry_run: bool = False, step_through: bool = False) -> int:
    """Process one freight order: open detail page, expand stops, upload PODs."""
    pod_path = row.pod_full_path
    if not os.path.isfile(pod_path):
        log.error("POD file not found: %s (row %d)", pod_path, row.row_number)
        return 0

    # Click into the order
    if not click_into_order(driver, row_index):
        return 0

    # Wait for detail page to load
    try:
        wait_for_element(driver, *EXPAND_ALL_BTN, timeout=20)
    except TimeoutException:
        log.warning("Detail page may not have loaded fully for %s", row.document_1)

    # Expand all stops
    expand_all_stops(driver)

    # Find and upload to all Proof sections
    uploads = find_and_upload_proofs(
        driver, pod_path, row.document_1,
        dry_run=dry_run, step_through=step_through,
    )

    # Navigate back to list
    try:
        back_btn = wait_for_element(driver, *BACK_BUTTON, timeout=10, clickable=True)
        back_btn.click()
        log.info("Navigated back to list")
    except TimeoutException:
        driver.back()
        log.info("Used browser back")

    return uploads


def run(driver: WebDriver, rows: list[InvoiceRow],
        *, dry_run: bool = False, step_through: bool = False, **_kwargs):
    """Execute the full Tile 4 workflow.

    Returns total number of uploads.
    """
    navigate_to_tile(driver)

    # Switch to "Documents for Reporting" tab if present
    try:
        tab = wait_for_element(driver, *DOCUMENTS_FOR_REPORTING_TAB, timeout=10, clickable=True)
        tab.click()
        log.info("Switched to 'Documents for Reporting' tab")
    except TimeoutException:
        log.info("'Documents for Reporting' tab not found — using default view")

    scroll_to_load_all(driver)

    total_uploads = 0
    for i, row in enumerate(rows):
        log.info("── POD upload row %d/%d (Excel row %d, doc %s) ──",
                 i + 1, len(rows), row.row_number, row.document_1)

        uploads = process_order(driver, row, row_index=0, dry_run=dry_run, step_through=step_through)
        total_uploads += uploads

        # After navigating back, the list re-renders — always use index 0
        # since the processed item moves to a different status

    log.info("Tile 4 complete — %d uploads %s",
             total_uploads, "(dry run)" if dry_run else "")
    return total_uploads
