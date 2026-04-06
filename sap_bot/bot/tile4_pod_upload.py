"""Tile 4 — Manage Freight Execution / POD Upload.

Goal: For each freight order, upload the correct PDF to every "Proof of..." window
across all stops. Human-gated.

NOTE: This is a placeholder — selectors need tuning once a screenshot of the
Manage Freight Execution tile is available.
"""

import os
import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_elements, wait_until_gone,
    scroll_to_load_all, take_screenshot, destructive_action,
)
from bot.excel_reader import InvoiceRow

log = logging.getLogger(__name__)

# ── Selectors (placeholder — need screenshot to refine) ─────────────────────
TILE_SELECTOR = (By.XPATH,
    "//div[contains(@class,'sapUshellTile')]//span[contains(text(),'Manage Freight Execution')]"
    "/ancestor::div[contains(@class,'sapUshellTile')]"
)

# List rows — click a middle column (NOT the radio button)
ORDER_ROWS = (By.CSS_SELECTOR,
    "table tbody tr, .sapMListItems .sapMLIB"
)

# Stop sections on the detail page
STOP_SECTIONS = (By.XPATH,
    "//div[contains(@class,'sapUiForm') or contains(@class,'sapMPanel')]"
    "[.//span[contains(text(),'Stop ')]]"
)

# "Proof of ..." section headers within stops
PROOF_SECTION = (By.XPATH,
    ".//span[starts-with(text(),'Proof of ')]/.."
)

# Report button inside a "Proof of..." section
PROOF_REPORT_BTN = (By.XPATH,
    ".//button[.//bdi[text()='Report'] or .//span[text()='Report']]"
)

# Browse button in the upload popup
BROWSE_BUTTON = (By.CSS_SELECTOR,
    "input[type='file'], button[id*='browse'], button[id*='Browse']"
)

# Confirm/OK button in upload popup
UPLOAD_CONFIRM_BTN = (By.XPATH,
    "//div[contains(@class,'sapMDialog')]//button[.//bdi[text()='OK'] or .//bdi[text()='Upload'] or .//bdi[text()='Confirm']]"
)

BACK_BUTTON = (By.CSS_SELECTOR,
    ".sapMNavBack, button[title='Back'], .sapUshellShellHeadItm[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Manage Freight Execution tile."""
    log.info("Navigating to Manage Freight Execution tile")
    tile = wait_for_element(driver, *TILE_SELECTOR, timeout=30, clickable=True)
    tile.click()
    log.info("Tile 4 page loaded")


@destructive_action("Upload POD file {file_path} for {stop_label}")
def upload_file(driver: WebDriver, file_path: str, *, stop_label: str = ""):
    """Upload a PDF via the browse dialog."""
    take_screenshot(driver, f"tile4_before_upload_{stop_label}")

    # For file input elements, we can send_keys directly with the path
    try:
        file_input = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        file_input.send_keys(file_path)
        log.info("Uploaded file: %s", file_path)
    except NoSuchElementException:
        log.error("File input not found for %s", stop_label)
        take_screenshot(driver, f"tile4_file_input_missing_{stop_label}")
        return False

    # Confirm upload
    try:
        confirm_btn = wait_for_element(driver, *UPLOAD_CONFIRM_BTN, timeout=10, clickable=True)
        confirm_btn.click()
        log.info("Confirmed upload for %s", stop_label)
    except TimeoutException:
        log.warning("No confirm button found — upload may have auto-completed")

    take_screenshot(driver, f"tile4_after_upload_{stop_label}")
    return True


def process_order(driver: WebDriver, row: InvoiceRow,
                  *, dry_run: bool = False, step_through: bool = False) -> int:
    """Process one freight order: open detail, upload POD to all stops."""
    pod_path = row.pod_full_path
    if not os.path.isfile(pod_path):
        log.error("POD file not found: %s (row %d)", pod_path, row.row_number)
        take_screenshot(driver, f"tile4_pod_missing_{row.document_1}")
        return 0

    # Find and click the order row (click middle column, not radio button)
    try:
        order_rows = driver.find_elements(*ORDER_ROWS)
        # TODO: match row to document number — for now process sequentially
        if not order_rows:
            log.warning("No order rows found")
            return 0

        # Click on a middle cell (e.g., 3rd column)
        cells = order_rows[0].find_elements(By.TAG_NAME, "td")
        if len(cells) >= 3:
            cells[2].click()
        else:
            order_rows[0].click()
        log.info("Opened freight order detail for %s", row.document_1)
    except Exception as e:
        log.error("Could not open order: %s", e)
        return 0

    # Find stop sections
    try:
        stops = wait_for_elements(driver, *STOP_SECTIONS, timeout=15)
    except TimeoutException:
        log.error("No stop sections found for %s", row.document_1)
        take_screenshot(driver, f"tile4_no_stops_{row.document_1}")
        driver.back()
        return 0

    uploads = 0
    for i, stop in enumerate(stops):
        stop_label = f"Stop_{i+1}_{row.document_1}"

        # Find "Proof of ..." sections within this stop
        proof_sections = stop.find_elements(*PROOF_SECTION)
        if not proof_sections:
            log.info("No 'Proof of...' section in stop %d — skipping", i + 1)
            continue

        for proof in proof_sections:
            proof_text = proof.text.strip()[:40]
            log.info("Found '%s' in stop %d", proof_text, i + 1)

            # Click Report button
            try:
                report_btn = proof.find_element(*PROOF_REPORT_BTN)
                report_btn.click()
                log.info("Clicked Report for '%s'", proof_text)
            except NoSuchElementException:
                log.warning("No Report button for '%s' — may already be uploaded", proof_text)
                continue

            # Upload the file
            result = upload_file(
                driver,
                pod_path,
                stop_label=f"{proof_text}_{row.document_1}",
                dry_run=dry_run,
                step_through=step_through,
            )
            if result and result != "skipped":
                uploads += 1

    # Navigate back
    try:
        back_btn = wait_for_element(driver, *BACK_BUTTON, timeout=10, clickable=True)
        back_btn.click()
    except TimeoutException:
        driver.back()

    return uploads


def run(driver: WebDriver, rows: list[InvoiceRow],
        *, dry_run: bool = False, step_through: bool = False, **_kwargs):
    """Execute the full Tile 4 workflow.

    Returns total number of uploads.
    """
    navigate_to_tile(driver)
    scroll_to_load_all(driver)

    total_uploads = 0
    for i, row in enumerate(rows):
        log.info("── POD upload row %d/%d (Excel row %d) ──", i + 1, len(rows), row.row_number)
        uploads = process_order(driver, row, dry_run=dry_run, step_through=step_through)
        total_uploads += uploads

    log.info("Tile 4 complete — %d uploads %s",
             total_uploads, "(dry run)" if dry_run else "")
    return total_uploads
