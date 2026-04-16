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
    _write_todo_status, read_todo_status,
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
    """Click into the first visible row by clicking a non-interactive middle cell.

    Same approach as tile 2 — click a cell like 'Reporting Status' or similar,
    NOT the radio button or a link, to navigate into the detail page.
    """
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

    # Use native Selenium click on the row element (same as tile 2)
    try:
        row_el.click()
        log.info("Clicked into first row (native)")
    except Exception:
        try:
            ActionChains(driver).move_to_element(row_el).click().perform()
            log.info("Clicked into first row (ActionChains)")
        except Exception as e:
            log.error("Could not click into row: %s", e)
            return False

    _time.sleep(2)
    wait_for_page_ready(driver)

    # Verify navigation happened (check for detail page indicators)
    on_detail = driver.execute_script("""
        var spans = document.querySelectorAll('span');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (/^Stop \\d/.test(t)) return 'stop_header';
            if (t === 'Information' || t === 'Attachments') return 'detail_tab';
        }
        return null;
    """)
    if on_detail:
        log.info("Navigation to detail confirmed (%s)", on_detail)
        return True

    log.warning("Click may not have navigated to detail page")
    take_screenshot(driver, "tile4_click_no_nav")
    return True  # proceed anyway — screenshots will show what happened


def expand_stop_2(driver: WebDriver):
    """Expand Stop 2 on the detail page.

    First scrolls down to force SAP to render all content, then clicks Expand All
    or the Stop 2 header.
    """
    # Scroll down to ensure stops are rendered (may be below the fold)
    driver.execute_script("""
        // Scroll page and any scrollable containers to bottom
        window.scrollTo(0, document.body.scrollHeight);
        var containers = document.querySelectorAll('.sapMPage, .sapMScrollContainer, [class*="Scroll"]');
        for (var i = 0; i < containers.length; i++) {
            containers[i].scrollTop = containers[i].scrollHeight;
        }
    """)
    _time.sleep(2)

    # Try Expand All first (expands all stops including Stop 2)
    expanded = driver.execute_script("""
        var links = document.querySelectorAll('a, button, span');
        for (var i = 0; i < links.length; i++) {
            if (links[i].offsetParent === null) continue;
            var t = links[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Expand All') {
                links[i].click();
                return 'expand_all_link';
            }
        }
        return null;
    """)

    if expanded:
        log.info("Clicked Expand All")
        _time.sleep(2)
        wait_for_page_ready(driver)

    # After Expand All (or as a fallback), explicitly click Stop 2 header to expand it.
    # "Expand All" often only expands the top-level sections, not nested ones.
    stop2_header = driver.execute_script("""
        var all = document.querySelectorAll('span, div, a, button');
        for (var i = 0; i < all.length; i++) {
            if (all[i].offsetParent === null) continue;
            var t = all[i].textContent.replace(/\\xAD/g, '').trim();
            if (/^Stop 2/.test(t) && t.length < 150) {
                // Prefer the closest panel/section header element
                var header = all[i].closest('[role="button"], [aria-expanded]');
                return header || all[i];
            }
        }
        return null;
    """)

    if stop2_header:
        # Scroll into view first
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", stop2_header)
        _time.sleep(0.5)

        # Check if Stop 2 is already expanded
        is_expanded = driver.execute_script("""
            var el = arguments[0];
            return el.getAttribute('aria-expanded') === 'true';
        """, stop2_header)

        if not is_expanded:
            # Click to expand — try native click, then ActionChains
            try:
                stop2_header.click()
                log.info("Clicked Stop 2 header (native)")
            except Exception:
                try:
                    ActionChains(driver).move_to_element(stop2_header).click().perform()
                    log.info("Clicked Stop 2 header (ActionChains)")
                except Exception as e:
                    log.warning("Could not click Stop 2 header: %s", e)
                    driver.execute_script("arguments[0].click();", stop2_header)
                    log.info("Clicked Stop 2 header (JS)")
            _time.sleep(2)
            wait_for_page_ready(driver)
        else:
            log.info("Stop 2 already expanded")
    else:
        log.warning("Could not find Stop 2 header")

    # Scroll Stop 2 area into view so Proof of Delivery is visible
    driver.execute_script("""
        var all = document.querySelectorAll('span, div');
        for (var i = 0; i < all.length; i++) {
            if (all[i].offsetParent === null) continue;
            var t = all[i].textContent.replace(/\\xAD/g, '').trim();
            if (t.indexOf('Proof of Delivery') !== -1) {
                all[i].scrollIntoView({block: 'center'});
                return;
            }
        }
        // Fallback: scroll to Stop 2 header
        for (var i = 0; i < all.length; i++) {
            if (all[i].offsetParent === null) continue;
            var t = all[i].textContent.replace(/\\xAD/g, '').trim();
            if (/^Stop 2/.test(t) && t.length < 150) {
                all[i].scrollIntoView({block: 'center'});
                return;
            }
        }
    """)
    _time.sleep(1)


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


def upload_and_report(driver: WebDriver, local_paths: list, *, doc_number: str = "",
                      dry_run: bool = False):
    """In the Report popup: fill time, upload PDF(s), click Report.

    Fills the form regardless of dry_run so you can see the result visually.
    Only the final Report button click is guarded by dry_run.
    """
    take_screenshot(driver, f"tile4_popup_{doc_number}")

    # Read Planned On time from popup
    planned_time = read_planned_time_from_popup(driver)

    # Fill the time into the "Enter Final Time" field
    if planned_time:
        time_input = driver.execute_script("""
            var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
            for (var i = dialogs.length - 1; i >= 0; i--) {
                var d = dialogs[i];
                if (d.offsetParent === null) continue;

                // Strategy 1: Find input with placeholder "Enter Final Time"
                var inputs = d.querySelectorAll('input:not([type="hidden"]):not([type="file"])');
                for (var j = 0; j < inputs.length; j++) {
                    if (inputs[j].offsetParent === null) continue;
                    var ph = inputs[j].placeholder || '';
                    var al = inputs[j].getAttribute('aria-label') || '';
                    if (ph.indexOf('Final Time') !== -1 || al.indexOf('Final Time') !== -1) {
                        return inputs[j];
                    }
                }

                // Strategy 2: Find input that's NOT the reason code / timezone
                for (var j = 0; j < inputs.length; j++) {
                    if (inputs[j].offsetParent === null) continue;
                    var ph = (inputs[j].placeholder || '').toLowerCase();
                    var al = (inputs[j].getAttribute('aria-label') || '').toLowerCase();
                    // Skip reason and timezone inputs
                    if (ph.indexOf('reason') !== -1 || al.indexOf('reason') !== -1) continue;
                    if (ph.indexOf('timezone') !== -1 || al.indexOf('timezone') !== -1) continue;
                    if (ph.indexOf('est') !== -1 || al.indexOf('est') !== -1) continue;
                    // Return the first remaining input (likely the Final Time)
                    return inputs[j];
                }
            }
            return null;
        """)
        if time_input:
            # Use focus + clear + send_keys (click can be intercepted by overlays)
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", time_input)
                _time.sleep(0.3)
                driver.execute_script("arguments[0].focus();", time_input)
                _time.sleep(0.3)
                time_input.clear()
                time_input.send_keys(planned_time)
                log.info("Entered Final Time: '%s'", planned_time)
            except Exception as e:
                log.warning("Could not enter time with send_keys: %s — trying JS", e)
                driver.execute_script("""
                    var el = arguments[0];
                    el.value = arguments[1];
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                """, time_input, planned_time)
                log.info("Entered Final Time via JS: '%s'", planned_time)
        else:
            log.warning("'Enter Final Time' input not found in popup")
            # Debug: dump popup inputs
            debug = driver.execute_script("""
                var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
                var results = [];
                for (var i = dialogs.length - 1; i >= 0; i--) {
                    var d = dialogs[i];
                    if (d.offsetParent === null) continue;
                    var inputs = d.querySelectorAll('input:not([type="hidden"])');
                    for (var j = 0; j < inputs.length; j++) {
                        if (inputs[j].offsetParent === null) continue;
                        results.push({
                            placeholder: inputs[j].placeholder,
                            ariaLabel: inputs[j].getAttribute('aria-label'),
                            type: inputs[j].type,
                            value: inputs[j].value
                        });
                    }
                }
                return results;
            """)
            log.info("DEBUG popup inputs: %s", debug)

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

    # SAP's Browse dialog requires ALL files selected at once (Ctrl+click multi-select).
    # Sequential uploads REPLACE the previous one — only the last file sticks.
    # In Selenium, send_keys with newline-joined paths selects multiple files in one action.
    try:
        # Make the file input visible/interactable via JS (some SAP dialogs hide it)
        driver.execute_script("""
            var fi = document.querySelector('input[type="file"]');
            if (fi) {
                fi.style.display = 'block';
                fi.style.visibility = 'visible';
                fi.style.opacity = '1';
                fi.removeAttribute('hidden');
            }
        """)
        _time.sleep(0.3)

        file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')

        # Pass all paths as newline-separated string — Selenium sends them as one selection
        all_paths = "\n".join(local_paths)
        file_input.send_keys(all_paths)
        log.info("Uploaded %d files in one action: %s",
                 len(local_paths),
                 ", ".join(os.path.basename(p) for p in local_paths))
        _time.sleep(2)
        wait_for_page_ready(driver)
    except Exception as e:
        log.error("Multi-file upload failed: %s", e)
        return False

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
        if dry_run:
            log.info("[DRY RUN] Would click Report button in popup — skipping")
            # Cancel the popup instead so the page isn't stuck
            driver.execute_script("""
                var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
                for (var i = dialogs.length - 1; i >= 0; i--) {
                    var d = dialogs[i];
                    if (d.offsetParent === null) continue;
                    var btns = d.querySelectorAll('button');
                    for (var j = 0; j < btns.length; j++) {
                        var t = btns[j].textContent.replace(/\\xAD/g, '').trim();
                        if (t === 'Cancel' || t === 'Close') { btns[j].click(); return; }
                    }
                }
            """)
            _time.sleep(1)
            return None  # dry run result
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

    # Filter to rows with POD filename — Pause flag only affects tile 3, not tile 4
    pod_rows = [r for r in rows if r.pod_filename]
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
            # Don't move to Status — this might be temporary (file not uploaded yet)
            _write_todo_status(row.document_1, f"pod_error: PDFs not in Drive: {missing[:40]}")
            results["errors"] += 1
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

            if status1 == "done" and (not row.is_collective or pod2_status == "done"):
                results["uploaded"] += 1
                current_status = read_todo_status(row.document_1)

                if "invoiced" in current_status.lower():
                    log.info("Both tile 3 (invoiced) and tile 4 (pod uploaded) done — moving to Status")
                    mark_fully_done(row, pod_uploaded_1=pod1_status, pod_uploaded_2=pod2_status)
                else:
                    log.info("POD uploaded but tile 3 not done yet — marking in To Do")
                    _write_todo_status(row.document_1, "pod_uploaded")

            elif status1 == "dry_run":
                results["skipped"] += 1
            else:
                results["errors"] += 1
                error_msg = f"doc1: {pod1_status}"
                if pod2_status:
                    error_msg += f", doc2: {pod2_status}"
                _write_todo_status(row.document_1, f"pod_error: {error_msg[:60]}")

        finally:
            for p in local_paths:
                cleanup_temp_file(p)

    log.info("Tile 4 complete: %s", results)
    return results
