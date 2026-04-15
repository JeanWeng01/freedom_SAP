"""Tile 4 — Manage Freight Execution / POD Upload.

Goal: For each freight order, upload the correct PDF to the "Proof of Delivery"
section in Stop 2 only.

Workflow per item:
1. Navigate to Manage Freight Execution tile
2. Filter for the freight document number
3. Click into the item (middle column, not radio button)
4. Expand Stop 2
5. Find "Proof of Delivery" section
6. Click "Report" → popup opens
7. Copy the "Planned On:" time from the popup, paste into the time field
8. Click "Browse..." → upload the PDF from Google Drive
9. Click "Report" button in popup to submit
10. Back to list, next item

After all uploads done, move completed items from To Do → Status tab.
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
from bot.google_sheets import (
    InvoiceRow, read_todo_items, mark_fully_done, mark_error,
)
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


def filter_by_document(driver: WebDriver, doc_number: str):
    """Filter the list by freight document number."""
    input_el = driver.execute_script("""
        var labels = document.querySelectorAll('label, span');
        for (var i = 0; i < labels.length; i++) {
            var clean = labels[i].textContent.replace(/\\xAD/g, '').trim();
            if (clean.indexOf('Freight Document') !== -1) {
                var parent = labels[i].parentElement;
                for (var j = 0; j < 5 && parent; j++) {
                    var input = parent.querySelector('input:not([type="hidden"])');
                    if (input) return input;
                    parent = parent.parentElement;
                }
            }
        }
        var inputs = document.querySelectorAll('input');
        for (var i = 0; i < inputs.length; i++) {
            var ph = (inputs[i].placeholder || '').toLowerCase();
            var al = (inputs[i].getAttribute('aria-label') || '').toLowerCase();
            if (ph.indexOf('freight doc') !== -1 || al.indexOf('freight doc') !== -1) return inputs[i];
        }
        return null;
    """)
    if input_el:
        input_el.click()
        _time.sleep(0.3)
        input_el.clear()
        input_el.send_keys(doc_number)
        from selenium.webdriver.common.keys import Keys
        input_el.send_keys(Keys.RETURN)
        log.info("Filtered for document: %s", doc_number)
        wait_for_page_ready(driver)
    else:
        log.error("Freight Document filter not found")
        take_screenshot(driver, "tile4_filter_missing")


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


def expand_stop_2(driver: WebDriver):
    """Expand Stop 2 on the detail page."""
    # Find and click "Stop 2" header or its expand toggle
    result = driver.execute_script("""
        var spans = document.querySelectorAll('span, button, a');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (/^Stop 2/.test(t)) {
                // Click the expand toggle (usually the parent panel header)
                var panel = spans[i].closest('[class*="sapMPanel"], [class*="sapUiForm"]');
                if (panel) {
                    var toggle = panel.querySelector('[class*="sapMPanelExpandableIcon"], button');
                    if (toggle) { toggle.click(); return 'toggle'; }
                }
                spans[i].click();
                return 'span';
            }
        }
        // Try "Expand All" button as fallback
        var links = document.querySelectorAll('a, button');
        for (var i = 0; i < links.length; i++) {
            if (links[i].offsetParent === null) continue;
            var t = links[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Expand All') { links[i].click(); return 'expand_all'; }
        }
        return null;
    """)
    if result:
        log.info("Expanded Stop 2 (%s)", result)
        _time.sleep(1)
        wait_for_page_ready(driver)
    else:
        log.warning("Could not find Stop 2 to expand — may already be expanded")


def find_proof_of_delivery_report_btn(driver: WebDriver):
    """Find the Report button in the "Proof of Delivery" section."""
    return driver.execute_script("""
        var spans = document.querySelectorAll('span, h3, h4');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (t.indexOf('Proof of Delivery') !== -1) {
                // Walk up to find the Report button in this section
                var parent = spans[i].parentElement;
                for (var j = 0; j < 8 && parent; j++) {
                    var btns = parent.querySelectorAll('a, button');
                    for (var k = 0; k < btns.length; k++) {
                        var bt = btns[k].textContent.replace(/\\xAD/g, '').trim();
                        if (bt === 'Report') return btns[k];
                    }
                    parent = parent.parentElement;
                }
            }
        }
        return null;
    """)


def read_planned_time_from_popup(driver: WebDriver) -> str | None:
    """Read the 'Planned On:' time from the Report popup and strip timezone."""
    import re
    planned = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            if (d.offsetParent === null) continue;
            // Look for "Planned On:" label and its value
            var labels = d.querySelectorAll('label, span');
            for (var j = 0; j < labels.length; j++) {
                var t = labels[j].textContent.replace(/\\xAD/g, '').trim();
                if (t.indexOf('Planned') !== -1 && t.indexOf('On') !== -1) {
                    // The value should be in a nearby element
                    var parent = labels[j].parentElement;
                    for (var k = 0; k < 3 && parent; k++) {
                        var texts = parent.querySelectorAll('span, div');
                        for (var m = 0; m < texts.length; m++) {
                            var val = texts[m].textContent.trim();
                            if (/[A-Z][a-z]{2} \\d{1,2}, \\d{4}/.test(val) && /(AM|PM|UTC)/i.test(val)) {
                                return val;
                            }
                        }
                        parent = parent.parentElement;
                    }
                }
            }
            // Fallback: find any date-like text in the dialog
            var allText = d.querySelectorAll('span, div');
            for (var j = 0; j < allText.length; j++) {
                var val = allText[j].textContent.trim();
                if (/[A-Z][a-z]{2} \\d{1,2}, \\d{4}/.test(val) && /(AM|PM|UTC)/i.test(val) && val.length < 50) {
                    return val;
                }
            }
        }
        return null;
    """)

    if not planned:
        log.error("Could not read Planned On time from popup")
        return None

    # Strip timezone
    stripped = re.sub(r'\s+(UTC[+-]?\d+|[A-Z]{2,5})\s*$', '', planned.strip())
    log.info("Planned On: '%s' → stripped: '%s'", planned, stripped)
    return stripped


@destructive_action("Upload POD and report for {doc_number}")
def upload_and_report(driver: WebDriver, local_paths: list, *, doc_number: str = ""):
    """In the Report popup: fill time, upload PDF(s), click Report.

    local_paths: list of local file paths to upload (supports multiple files).
    """
    take_screenshot(driver, f"tile4_popup_{doc_number}")

    # Read Planned On time from popup
    planned_time = read_planned_time_from_popup(driver)

    # Fill the time into the reporting field
    if planned_time:
        time_input = driver.execute_script("""
            var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
            for (var i = dialogs.length - 1; i >= 0; i--) {
                var d = dialogs[i];
                if (d.offsetParent === null) continue;
                var inputs = d.querySelectorAll('input:not([type="hidden"]):not([type="file"])');
                for (var j = 0; j < inputs.length; j++) {
                    if (inputs[j].offsetParent !== null) return inputs[j];
                }
            }
            return null;
        """)
        if time_input:
            time_input.click()
            _time.sleep(0.3)
            time_input.clear()
            time_input.send_keys(planned_time)
            log.info("Entered time: '%s'", planned_time)
        else:
            log.warning("Time input not found in popup")

    # Upload each PDF via Browse
    # SAP's file input usually accepts one file at a time, but we can send
    # multiple paths by joining with newlines (Selenium supports this for
    # multi-file inputs). We also try sequential uploads as a fallback.
    if not local_paths:
        log.error("No files to upload")
        return False

    file_input = driver.execute_script("""
        return document.querySelector('input[type="file"]');
    """)
    if not file_input:
        log.error("File input not found in popup")
        take_screenshot(driver, f"tile4_no_file_input_{doc_number}")
        return False

    # Try sending all files at once (newline-joined paths for multi-file input)
    uploaded = 0
    for path in local_paths:
        try:
            # Re-find file input each iteration in case popup updated
            fi = driver.execute_script("""
                return document.querySelector('input[type="file"]');
            """)
            if fi:
                fi.send_keys(path)
                log.info("Uploaded file %d/%d: %s",
                         uploaded + 1, len(local_paths), os.path.basename(path))
                uploaded += 1
                _time.sleep(2)
                wait_for_page_ready(driver)
            else:
                log.warning("File input disappeared after upload %d", uploaded)
                break
        except Exception as e:
            log.error("Error uploading %s: %s", os.path.basename(path), e)

    if uploaded == 0:
        log.error("No files were uploaded")
        return False

    log.info("Uploaded %d/%d files", uploaded, len(local_paths))

    take_screenshot(driver, f"tile4_filled_popup_{doc_number}")

    # Click Report button in the popup
    btn = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            if (d.offsetParent === null) continue;
            var btns = d.querySelectorAll('button');
            for (var j = 0; j < btns.length; j++) {
                var t = btns[j].textContent.replace(/\\xAD/g, '').trim();
                if (t === 'Report') return btns[j];
            }
        }
        return null;
    """)
    if btn:
        try:
            btn.click()
            log.info("Clicked Report in popup (native)")
        except Exception:
            ActionChains(driver).move_to_element(btn).click().perform()
            log.info("Clicked Report in popup (ActionChains)")
        _time.sleep(2)
        wait_for_page_ready(driver)
        dismiss_any_popup(driver)
        take_screenshot(driver, f"tile4_after_report_{doc_number}")
        return True
    else:
        log.error("Report button not found in popup")
        take_screenshot(driver, f"tile4_no_report_btn_{doc_number}")
        return False


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


def process_item(driver: WebDriver, row: InvoiceRow, local_pdf_paths: list,
                 *, dry_run: bool = False) -> str:
    """Process one freight document: click in, expand Stop 2, upload POD(s) to Proof of Delivery.

    local_pdf_paths: list of local file paths (1 or more files to upload together).
    """
    doc_number = row.document_1

    # Filter for this specific document
    filter_by_document(driver, doc_number)
    take_screenshot(driver, f"tile4_filtered_{doc_number}")

    # Click into the item
    if not click_into_first_row(driver):
        return "error: could not click into freight order row"

    take_screenshot(driver, f"tile4_detail_{doc_number}")

    # Expand Stop 2
    expand_stop_2(driver)
    take_screenshot(driver, f"tile4_stop2_expanded_{doc_number}")

    # Find the Report button in Proof of Delivery section
    report_btn = find_proof_of_delivery_report_btn(driver)
    if not report_btn:
        log.error("Proof of Delivery Report button not found for %s", doc_number)
        take_screenshot(driver, f"tile4_no_pod_btn_{doc_number}")
        click_back(driver)
        return "error: Proof of Delivery Report button not found"

    # Click the Report button (native click)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", report_btn)
        _time.sleep(0.5)
        report_btn.click()
        log.info("Clicked Report for Proof of Delivery")
    except Exception:
        ActionChains(driver).move_to_element(report_btn).click().perform()
        log.info("Clicked Report (ActionChains)")

    _time.sleep(1)
    wait_for_page_ready(driver)

    # Upload and report
    result = upload_and_report(driver, local_pdf_paths, doc_number=doc_number,
                               dry_run=dry_run)

    click_back(driver)

    if result is None:
        return "dry_run"
    return "done" if result else "error: upload failed"


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 4 workflow using Google Sheets + Drive.

    Reads the same To Do tab as tile 3. For each item with a POD filename,
    uploads the PDF to Proof of Delivery in Stop 2.

    If item has 2 legs (Document_1 + Document_2), uploads for each doc separately.
    After all uploads done, moves completed items to Status tab.
    """
    launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")

    rows = read_todo_items()
    if not rows:
        log.info("No items in To Do tab — nothing to upload")
        return {"uploaded": 0, "errors": 0}

    # Filter to rows with POD filename and not paused
    pod_rows = [r for r in rows if r.pod_filename and not r.pause]
    if not pod_rows:
        log.info("No active rows with POD_Filename — nothing to upload")
        return {"uploaded": 0, "errors": 0}

    log.info("Found %d items with POD files to upload", len(pod_rows))
    results = {"uploaded": 0, "errors": 0, "skipped": 0}

    for i, row in enumerate(pod_rows):
        filenames = row.pod_filenames
        log.info("════ POD Upload %d/%d (doc %s, files %s) ════",
                 i + 1, len(pod_rows), row.document_1, filenames)

        # Download all PDFs from Google Drive
        local_paths = []
        download_errors = []
        for fname in filenames:
            p = download_file(fname)
            if p:
                local_paths.append(p)
            else:
                download_errors.append(fname)

        if download_errors:
            missing = ", ".join(download_errors)
            log.error("Missing PDFs in Drive: %s", missing)
            mark_error(row, f"PDFs not found in Google Drive: {missing}")
            results["errors"] += 1
            # Clean up any that were downloaded
            for p in local_paths:
                cleanup_temp_file(p)
            continue

        try:
            # Navigate to tile for each item
            if launchpad_url:
                driver.get(launchpad_url)
                wait_for_page_ready(driver)
            navigate_to_tile(driver)

            # Process Document_1
            status1 = process_item(driver, row, local_paths, dry_run=dry_run)
            log.info("Doc 1 (%s) result: %s", row.document_1, status1)

            pod1_status = "done" if status1 == "done" else status1
            pod2_status = ""

            # Process Document_2 if it exists
            if row.is_collective and row.document_2:
                log.info("Processing leg 2: %s", row.document_2)

                if launchpad_url:
                    driver.get(launchpad_url)
                    wait_for_page_ready(driver)
                navigate_to_tile(driver)

                row2 = InvoiceRow(
                    sheet_row=row.sheet_row,
                    document_1=row.document_2,
                    document_2=None,
                    invoice=row.invoice,
                    pod_filename=row.pod_filename,
                    charge_hours_1=None,
                    charge_hours_2=None,
                    pause=False,
                )
                status2 = process_item(driver, row2, local_paths, dry_run=dry_run)
                log.info("Doc 2 (%s) result: %s", row.document_2, status2)
                pod2_status = "done" if status2 == "done" else status2

            # Move to Status tab if both uploads succeeded
            if status1 == "done" and (not row.is_collective or pod2_status == "done"):
                results["uploaded"] += 1
                mark_fully_done(row, pod_uploaded_1=pod1_status, pod_uploaded_2=pod2_status)
            elif status1 == "dry_run":
                results["skipped"] += 1
            else:
                results["errors"] += 1
                error_msg = f"doc1: {pod1_status}"
                if pod2_status:
                    error_msg += f", doc2: {pod2_status}"
                mark_error(row, error_msg)

        finally:
            for p in local_paths:
                cleanup_temp_file(p)

    log.info("Tile 4 complete: %s", results)
    return results
