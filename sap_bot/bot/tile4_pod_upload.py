"""Tile 4 — Manage Freight Execution / POD Upload.

Goal: For each freight order, upload the correct PDF from Google Drive
to every "Proof of..." section across all stops. Human-gated.

Workflow per item:
1. Click into the item (middle column, NOT radio button)
2. Click "Expand All" to show all stops
3. For each stop, find "Proof of Pick-Up" or "Proof of Delivery" section
4. Click "Report" → Browse → upload PDF from Google Drive
5. Same PDF uploaded to every Proof section
6. Back to list, next item
"""

import os
import logging
import time as _time
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_page_ready, take_screenshot,
    destructive_action, click_tile,
)
from bot.google_sheets import InvoiceRow, read_todo_items, move_to_status, mark_error
from bot.google_drive import download_file, cleanup_temp_file

log = logging.getLogger(__name__)

TILE_NAME = "Manage Freight Execution"

BACK_BUTTON = (By.CSS_SELECTOR,
    "button[title='Back'], .sapMNavBack, .sapUshellShellHeadItm[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Manage Freight Execution tile and wait for it to load."""
    click_tile(driver, TILE_NAME)
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("""
                var spans = document.querySelectorAll('span');
                for (var i = 0; i < spans.length; i++) {
                    if (spans[i].offsetParent === null) continue;
                    var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
                    if (/Freight Documents \\(/.test(t)) return true;
                    if (t.indexOf('Documents for Reporting') !== -1) return true;
                }
                return false;
            """)
        )
    except TimeoutException:
        log.warning("Tile 4 page content not detected — retrying")
        launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")
        if launchpad_url:
            driver.get(launchpad_url)
            wait_for_page_ready(driver)
        click_tile(driver, TILE_NAME)
        _time.sleep(10)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile4_page_loaded")
    log.info("Tile 4 page loaded")


def dismiss_any_popup(driver: WebDriver) -> bool:
    """Dismiss any popup dialog."""
    result = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            if (d.offsetParent === null) continue;
            var btns = d.querySelectorAll('button');
            for (var j = 0; j < btns.length; j++) {
                var t = btns[j].textContent.replace(/\\xAD/g, '').trim();
                if (t === 'Close' || t === 'OK' || t === 'Cancel') {
                    btns[j].click();
                    return true;
                }
            }
        }
        return null;
    """)
    if result:
        _time.sleep(1)
        return True
    return False


def click_into_first_row(driver: WebDriver) -> bool:
    """Click into the first visible row by clicking a non-interactive cell."""
    row_el = driver.execute_script("""
        var rows = document.querySelectorAll(
            '[role="row"].sapMListTblRow, .sapMListItems .sapMLIB, .sapMListTblRow'
        );
        var dataRows = [];
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].closest('thead')) continue;
            if (rows[i].classList.contains('sapMListTblHeader')) continue;
            if (rows[i].textContent.trim().length > 0) dataRows.push(rows[i]);
        }
        return dataRows.length > 0 ? dataRows[0] : null;
    """)

    if not row_el:
        log.error("No data rows found")
        return False

    try:
        row_el.click()
        log.info("Clicked into first row (native)")
    except Exception:
        ActionChains(driver).move_to_element(row_el).click().perform()
        log.info("Clicked into first row (ActionChains)")

    _time.sleep(2)
    wait_for_page_ready(driver)
    return True


def click_expand_all(driver: WebDriver):
    """Click 'Expand All' button to show all stops."""
    btn = driver.execute_script("""
        var links = document.querySelectorAll('a, button, span');
        for (var i = 0; i < links.length; i++) {
            if (links[i].offsetParent === null) continue;
            var t = links[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Expand All') return links[i];
        }
        return null;
    """)
    if btn:
        try:
            btn.click()
        except Exception:
            ActionChains(driver).move_to_element(btn).click().perform()
        log.info("Clicked Expand All")
        _time.sleep(1)
        wait_for_page_ready(driver)
    else:
        log.info("Expand All not found — stops may already be expanded")


def get_proof_report_buttons(driver: WebDriver) -> list:
    """Find all visible Report buttons inside 'Proof of...' sections."""
    return driver.execute_script("""
        var results = [];
        // Find all "Proof of..." headers
        var spans = document.querySelectorAll('span, h3, h4');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (t.indexOf('Proof of') !== -1) {
                // Find the Report button near this section
                var parent = spans[i].parentElement;
                for (var j = 0; j < 8 && parent; j++) {
                    var btns = parent.querySelectorAll('a, button');
                    for (var k = 0; k < btns.length; k++) {
                        var bt = btns[k].textContent.replace(/\\xAD/g, '').trim();
                        if (bt === 'Report') {
                            results.push({label: t, btn: btns[k]});
                            break;
                        }
                    }
                    if (results.length > 0 && results[results.length-1].label === t) break;
                    parent = parent.parentElement;
                }
            }
        }
        return results;
    """)


@destructive_action("Upload POD for {proof_label}")
def upload_pod(driver: WebDriver, local_path: str, *, proof_label: str = ""):
    """Upload a PDF file after clicking the Report button in a Proof section."""
    take_screenshot(driver, f"tile4_before_upload_{proof_label[:20]}")

    # Find the file input (hidden input[type=file])
    file_input = driver.execute_script("""
        return document.querySelector('input[type="file"]');
    """)

    if file_input:
        file_input.send_keys(local_path)
        log.info("Uploaded file: %s", local_path)
        _time.sleep(2)
        wait_for_page_ready(driver)
    else:
        log.error("File input not found for %s", proof_label)
        take_screenshot(driver, f"tile4_file_input_missing_{proof_label[:20]}")
        return False

    # Confirm upload if dialog appears
    dismiss_any_popup(driver)
    take_screenshot(driver, f"tile4_after_upload_{proof_label[:20]}")
    return True


def click_back(driver: WebDriver):
    """Navigate back with fallbacks."""
    wait_for_page_ready(driver)
    dismiss_any_popup(driver)
    try:
        back_btn = wait_for_element(driver, *BACK_BUTTON, timeout=10, clickable=True)
        back_btn.click()
        wait_for_page_ready(driver)
    except Exception:
        try:
            driver.execute_script("""
                var btn = document.querySelector('a#backBtn, button[title="Back"], .sapMNavBack');
                if (btn) btn.click();
            """)
            wait_for_page_ready(driver)
        except Exception:
            driver.back()
            wait_for_page_ready(driver)


def process_item(driver: WebDriver, row: InvoiceRow, local_pdf_path: str,
                 *, dry_run: bool = False) -> str:
    """Process one item: open detail, expand stops, upload PDF to all Proof sections."""
    click_expand_all(driver)
    take_screenshot(driver, f"tile4_detail_{row.document_1}")

    proof_buttons = get_proof_report_buttons(driver)

    if not proof_buttons:
        log.warning("No 'Proof of...' sections found for %s", row.document_1)
        return "error: no Proof sections found"

    log.info("Found %d Proof sections for %s", len(proof_buttons), row.document_1)

    uploads = 0
    for i, proof in enumerate(proof_buttons):
        label = proof.get("label", f"Proof_{i+1}") if isinstance(proof, dict) else f"Proof_{i+1}"
        safe_label = label.replace(" ", "_")[:25]
        log.info("Processing: %s", label)

        # Click the Report button (native click)
        btn_el = proof.get("btn") if isinstance(proof, dict) else proof
        if btn_el:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_el)
                _time.sleep(0.5)
                btn_el.click()
                log.info("Clicked Report for '%s' (native)", label)
            except Exception:
                try:
                    ActionChains(driver).move_to_element(btn_el).click().perform()
                    log.info("Clicked Report for '%s' (ActionChains)", label)
                except Exception as e:
                    log.error("Could not click Report for '%s': %s", label, e)
                    continue

            _time.sleep(1)
            wait_for_page_ready(driver)

            result = upload_pod(driver, local_pdf_path, proof_label=safe_label,
                                dry_run=dry_run)
            if result and result != "skipped":
                uploads += 1

    if uploads == 0 and not dry_run:
        return "error: no files uploaded"

    return "done" if not dry_run else "dry_run"


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 4 workflow using Google Sheets + Drive."""
    launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")

    rows = read_todo_items()
    if not rows:
        log.info("No items in To Do tab — nothing to upload")
        return {"uploaded": 0, "errors": 0}

    # Filter to rows that have a POD filename
    pod_rows = [r for r in rows if r.pod_filename]
    if not pod_rows:
        log.info("No rows with POD_Filename — nothing to upload")
        return {"uploaded": 0, "errors": 0}

    log.info("Found %d items with POD files to upload", len(pod_rows))
    results = {"uploaded": 0, "errors": 0, "skipped": 0}

    navigate_to_tile(driver)

    for i, row in enumerate(pod_rows):
        log.info("════ POD Upload %d/%d (doc %s, file %s) ════",
                 i + 1, len(pod_rows), row.document_1, row.pod_full_filename)

        # Download PDF from Google Drive
        local_path = download_file(row.pod_full_filename)
        if not local_path:
            log.error("PDF '%s' not found in Google Drive", row.pod_full_filename)
            mark_error(row, f"PDF '{row.pod_full_filename}' not found in Google Drive")
            results["errors"] += 1
            continue

        try:
            # Click into the first row in the list
            if not click_into_first_row(driver):
                log.error("Could not click into row")
                mark_error(row, "could not click into freight order row")
                results["errors"] += 1
                continue

            wait_for_page_ready(driver)
            status = process_item(driver, row, local_path, dry_run=dry_run)
            log.info("Result: %s", status)

            if status == "done":
                results["uploaded"] += 1
                # Don't move to Status here — tile 3 handles the row lifecycle
            elif status == "dry_run":
                results["skipped"] += 1
            elif status.startswith("error:"):
                results["errors"] += 1
                mark_error(row, status.replace("error: ", ""))

            click_back(driver)

        finally:
            cleanup_temp_file(local_path)

        # Re-enter tile for next item
        if i < len(pod_rows) - 1:
            if launchpad_url:
                driver.get(launchpad_url)
                wait_for_page_ready(driver)
            navigate_to_tile(driver)

    log.info("Tile 4 complete: %s", results)
    return results
