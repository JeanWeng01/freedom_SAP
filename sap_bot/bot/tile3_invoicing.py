"""Tile 3 — Invoice Freight Documents.

Goal: For each row in Google Sheet "To Do" tab, filter for freight document(s),
create invoice, enter invoice number, add charges if applicable, and submit.

Human-gated: runs on Railway 4am-9am, or manually triggered.
"""

import logging
import time as _time
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_page_ready, take_screenshot,
    destructive_action, click_tile,
)
from bot.google_sheets import InvoiceRow, read_todo_items, move_to_status, mark_error

log = logging.getLogger(__name__)

TILE_NAME = "Invoice Freight Documents"

BACK_BUTTON = (By.CSS_SELECTOR,
    "button[title='Back'], .sapMNavBack, .sapUshellShellHeadItm[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Invoice Freight Documents tile and wait for it to load."""
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
                    if (t === 'To be Invoiced' || t === 'To Be Invoiced') return true;
                }
                return false;
            """)
        )
    except TimeoutException:
        log.warning("Tile 3 page content not detected after 30s — retrying")
        import os
        launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")
        if launchpad_url:
            driver.get(launchpad_url)
            wait_for_page_ready(driver)
        click_tile(driver, TILE_NAME)
        _time.sleep(10)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile3_page_loaded")
    log.info("Tile 3 page loaded")


def dismiss_any_popup(driver: WebDriver) -> bool:
    """Dismiss any error/info popup on the page."""
    result = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            if (d.offsetParent === null) continue;
            var btns = d.querySelectorAll('button');
            for (var j = 0; j < btns.length; j++) {
                var t = btns[j].textContent.replace(/\\xAD/g, '').trim();
                if (t === 'Close' || t === 'OK' || t === 'Cancel') {
                    var msgText = d.textContent.substring(0, 200).trim();
                    btns[j].click();
                    return msgText;
                }
            }
        }
        return null;
    """)
    if result:
        log.warning("Dismissed popup: %s", result[:150])
        _time.sleep(1)
        return True
    return False


def filter_single_document(driver: WebDriver, doc_number: str):
    """Filter by a single freight document number using the filter input."""
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
        // Fallback: find any input with freight document in placeholder/label
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
        input_el.send_keys(Keys.RETURN)
        log.info("Filtered for document: %s", doc_number)
        wait_for_page_ready(driver)
    else:
        log.error("Freight Document filter input not found")
        take_screenshot(driver, "tile3_filter_missing")


def filter_collective_documents(driver: WebDriver, doc1: str, doc2: str):
    """Filter for two freight documents by entering both in the filter field."""
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
    if not input_el:
        log.error("Freight Document filter input not found")
        take_screenshot(driver, "tile3_filter_missing")
        return

    # Enter first doc, press Enter, then second doc, press Enter
    input_el.click()
    _time.sleep(0.3)
    input_el.clear()
    input_el.send_keys(doc1)
    input_el.send_keys(Keys.RETURN)
    _time.sleep(1)
    input_el.send_keys(doc2)
    input_el.send_keys(Keys.RETURN)
    log.info("Filtered for documents: %s + %s", doc1, doc2)
    wait_for_page_ready(driver)


def select_all_visible_rows(driver: WebDriver):
    """Select all visible document rows using native Selenium clicks on checkboxes."""
    # Find all visible checkboxes in data rows
    checkboxes = driver.execute_script("""
        var cbs = [];
        var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].offsetParent === null) continue;
            if (rows[i].classList.contains('sapMListTblHeader')) continue;
            var cb = rows[i].querySelector('.sapMCb');
            if (cb && cb.offsetParent !== null) cbs.push(cb);
        }
        return cbs;
    """)

    selected = 0
    for cb in checkboxes:
        try:
            cb.click()  # native Selenium click
            selected += 1
        except Exception:
            try:
                ActionChains(driver).move_to_element(cb).click().perform()
                selected += 1
            except Exception as e:
                log.warning("Could not click checkbox: %s", e)

    log.info("Selected %d rows (native clicks)", selected)
    _time.sleep(0.5)


def click_create_invoice(driver: WebDriver, collective: bool) -> bool:
    """Click Create Invoice or Create Collective Invoice button."""
    button_text = "Create Collective Invoice" if collective else "Create Invoice"
    btn = driver.execute_script("""
        var target = arguments[0];
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].offsetParent === null) continue;
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === target) return btns[i];
        }
        return null;
    """, button_text)

    if not btn:
        log.error("'%s' button not found", button_text)
        take_screenshot(driver, f"tile3_no_{button_text.replace(' ', '_')}_btn")
        return False

    try:
        btn.click()
    except Exception:
        ActionChains(driver).move_to_element(btn).click().perform()
    log.info("Clicked '%s'", button_text)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile3_after_create_invoice")
    return True


def enter_invoice_number(driver: WebDriver, invoice_num: str):
    """Go to Invoice Details tab and enter the invoice number."""
    # Click Invoice Details tab
    driver.execute_script("""
        var spans = document.querySelectorAll('span');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Invoice Details') {
                var tab = spans[i].closest('[role="tab"], [class*="sapMITBFilter"], [class*="sapMITBItem"]');
                if (tab) { tab.click(); return; }
                spans[i].click();
                return;
            }
        }
    """)
    _time.sleep(1)
    wait_for_page_ready(driver)

    # Find the Invoice input field
    inv_input = driver.execute_script("""
        var labels = document.querySelectorAll('label, span');
        for (var i = 0; i < labels.length; i++) {
            var clean = labels[i].textContent.replace(/\\xAD/g, '').trim();
            if (clean === 'Invoice:' || clean === 'Invoice') {
                var parent = labels[i].parentElement;
                for (var j = 0; j < 5 && parent; j++) {
                    var input = parent.querySelector('input:not([type="hidden"])');
                    if (input && input.offsetParent !== null) return input;
                    parent = parent.parentElement;
                }
            }
        }
        return null;
    """)

    if inv_input:
        inv_input.click()
        _time.sleep(0.3)
        inv_input.clear()
        inv_input.send_keys(invoice_num)
        log.info("Entered invoice number: %s", invoice_num)
    else:
        log.error("Invoice input field not found")
        take_screenshot(driver, "tile3_invoice_input_missing")


def add_charge(driver: WebDriver, charge_type: str, amount: float):
    """Add a charge in the Charges tab."""
    # Click Charges tab
    driver.execute_script("""
        var spans = document.querySelectorAll('span');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Charges') {
                var tab = spans[i].closest('[role="tab"], [class*="sapMITBFilter"], [class*="sapMITBItem"]');
                if (tab) { tab.click(); return; }
                spans[i].click();
                return;
            }
        }
    """)
    _time.sleep(1)
    wait_for_page_ready(driver)

    # Click Add button
    add_btn = driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].offsetParent === null) continue;
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Add') return btns[i];
        }
        return null;
    """)
    if add_btn:
        try:
            add_btn.click()
        except Exception:
            ActionChains(driver).move_to_element(add_btn).click().perform()
        log.info("Clicked Add charge")
        _time.sleep(1)
        wait_for_page_ready(driver)
    else:
        log.error("Add button not found in Charges tab")
        take_screenshot(driver, "tile3_add_charge_missing")
        return

    # Select "Charge" from dropdown if prompted
    driver.execute_script("""
        var items = document.querySelectorAll('li, [class*="sapMSLI"], [role="option"]');
        for (var i = 0; i < items.length; i++) {
            if (items[i].offsetParent === null) continue;
            var t = items[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Charge') { items[i].click(); return; }
        }
    """)
    _time.sleep(1)
    wait_for_page_ready(driver)

    # Select charge category (e.g. "Waiting Charges")
    driver.execute_script("""
        var items = document.querySelectorAll('li, [class*="sapMSLI"], [role="option"], span');
        for (var i = 0; i < items.length; i++) {
            if (items[i].offsetParent === null) continue;
            var t = items[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === arguments[0]) { items[i].click(); return; }
        }
    """, charge_type)
    _time.sleep(1)
    wait_for_page_ready(driver)

    # Enter amount in the rate/amount field — find the last visible input with numeric type
    driver.execute_script("""
        var inputs = document.querySelectorAll('input[type="text"], input:not([type="hidden"])');
        var rateInput = null;
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent === null) continue;
            var al = (inputs[i].getAttribute('aria-label') || '').toLowerCase();
            var ph = (inputs[i].placeholder || '').toLowerCase();
            if (al.indexOf('rate') !== -1 || al.indexOf('amount') !== -1 ||
                ph.indexOf('rate') !== -1 || ph.indexOf('amount') !== -1) {
                rateInput = inputs[i];
            }
        }
        if (rateInput) {
            rateInput.focus();
            rateInput.value = '';
            rateInput.dispatchEvent(new Event('input', {bubbles: true}));
        }
        return rateInput;
    """)
    # Use Selenium to type the amount (JS value setting may not trigger SAP bindings)
    rate_inputs = driver.find_elements(By.CSS_SELECTOR, "input")
    for inp in reversed(rate_inputs):
        try:
            al = inp.get_attribute("aria-label") or ""
            if "rate" in al.lower() or "amount" in al.lower():
                inp.clear()
                inp.send_keys(str(amount))
                log.info("Entered charge amount: %s", amount)
                break
        except Exception:
            continue
    _time.sleep(0.5)
    take_screenshot(driver, f"tile3_charge_added_{charge_type[:15]}")


@destructive_action("Submit invoice {invoice_num}")
def click_submit(driver: WebDriver, *, invoice_num: str = ""):
    """Click the Submit button."""
    take_screenshot(driver, f"tile3_before_submit_{invoice_num}")
    btn = driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].offsetParent === null) continue;
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Submit') return btns[i];
        }
        return null;
    """)
    if btn:
        try:
            btn.click()
        except Exception:
            ActionChains(driver).move_to_element(btn).click().perform()
        log.info("Clicked Submit for invoice %s", invoice_num)
        _time.sleep(2)
        wait_for_page_ready(driver)
        dismiss_any_popup(driver)
        take_screenshot(driver, f"tile3_after_submit_{invoice_num}")
    else:
        log.error("Submit button not found")
        take_screenshot(driver, "tile3_submit_missing")


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


def fill_invoice_and_submit(driver: WebDriver, row: InvoiceRow,
                            *, dry_run: bool = False, step_through: bool = False) -> str:
    """Fill invoice number, add charges, and submit. Called when invoice page is open."""
    enter_invoice_number(driver, row.invoice)

    if row.has_charges_leg1:
        add_charge(driver, row.charge_type_1, row.charge_amount_1)

    if row.is_collective and row.has_charges_leg2:
        add_charge(driver, row.charge_type_2, row.charge_amount_2)

    click_submit(driver, invoice_num=row.invoice, dry_run=dry_run, step_through=step_through)

    if dry_run:
        click_back(driver)
        return "dry_run"

    return "submitted"


def navigate_to_manage_invoices(driver: WebDriver):
    """Navigate to the Manage Invoices tile."""
    click_tile(driver, "Manage Invoices")
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("""
                var spans = document.querySelectorAll('span');
                for (var i = 0; i < spans.length; i++) {
                    if (spans[i].offsetParent === null) continue;
                    var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
                    if (t.indexOf('Draft') !== -1 || t.indexOf('Invoicing in Process') !== -1) return true;
                }
                return false;
            """)
        )
    except TimeoutException:
        log.warning("Manage Invoices page content not detected — retrying")
        import os
        launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")
        if launchpad_url:
            driver.get(launchpad_url)
            wait_for_page_ready(driver)
        click_tile(driver, "Manage Invoices")
        _time.sleep(10)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile3_manage_invoices_loaded")
    log.info("Manage Invoices page loaded")


def click_drafts_tab(driver: WebDriver):
    """Click the 'Draft' tab in Manage Invoices."""
    driver.execute_script("""
        var spans = document.querySelectorAll('span');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var clean = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (/^Draft/.test(clean)) {
                var tab = spans[i].closest('[role="tab"], [class*="sapMITBFilter"], [class*="sapMITBItem"]');
                if (tab) { tab.click(); return; }
                spans[i].click();
                return;
            }
        }
    """)
    wait_for_page_ready(driver)
    log.info("Clicked Drafts tab")
    take_screenshot(driver, "tile3_drafts_tab")


def filter_in_manage_invoices(driver: WebDriver, doc1: str, doc2: str = None):
    """Filter for freight document(s) in the Manage Invoices tile."""
    # Find the Freight Document filter input
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
    if not input_el:
        log.error("Filter input not found in Manage Invoices")
        take_screenshot(driver, "tile3_manage_inv_filter_missing")
        return

    input_el.click()
    _time.sleep(0.3)
    input_el.clear()
    input_el.send_keys(doc1)
    input_el.send_keys(Keys.RETURN)
    if doc2:
        _time.sleep(1)
        input_el.send_keys(doc2)
        input_el.send_keys(Keys.RETURN)
    log.info("Filtered in Manage Invoices for: %s%s", doc1, f" + {doc2}" if doc2 else "")
    wait_for_page_ready(driver)


def click_into_draft_row(driver: WebDriver) -> bool:
    """Click into the first visible row in the Drafts list (native click on a safe cell)."""
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
        log.error("No draft rows found")
        take_screenshot(driver, "tile3_no_draft_rows")
        return False

    try:
        row_el.click()
        log.info("Clicked into draft row (native)")
    except Exception:
        ActionChains(driver).move_to_element(row_el).click().perform()
        log.info("Clicked into draft row (ActionChains)")

    _time.sleep(2)
    wait_for_page_ready(driver)
    return True


def process_drafted_invoice(driver: WebDriver, row: InvoiceRow,
                            *, dry_run: bool = False, step_through: bool = False) -> str:
    """Recover a drafted invoice from Manage Invoices tile.

    Flow: go to Manage Invoices → Drafts tab → filter → click into draft →
    fill invoice number + charges → submit.
    """
    import os
    launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")

    doc_desc = row.document_1
    if row.is_collective:
        doc_desc = f"{row.document_1} + {row.document_2}"

    log.info("Recovering drafted invoice for %s", doc_desc)

    try:
        # Go to home first
        if launchpad_url:
            driver.get(launchpad_url)
            wait_for_page_ready(driver)

        # Navigate to Manage Invoices tile
        navigate_to_manage_invoices(driver)

        # Click Drafts tab
        click_drafts_tab(driver)

        # Filter for the freight document(s)
        filter_in_manage_invoices(driver, row.document_1,
                                  row.document_2 if row.is_collective else None)

        take_screenshot(driver, f"tile3_draft_filtered_{row.document_1}")

        # Click into the draft
        if not click_into_draft_row(driver):
            return "error: draft invoice not found in Manage Invoices"

        take_screenshot(driver, f"tile3_draft_opened_{row.document_1}")

        # Check that invoice page opened
        on_invoice = driver.execute_script("""
            var spans = document.querySelectorAll('span');
            for (var i = 0; i < spans.length; i++) {
                if (spans[i].offsetParent === null) continue;
                var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
                if (t === 'Invoice Details') return true;
            }
            return false;
        """)

        if not on_invoice:
            log.error("Draft invoice page did not open for %s", doc_desc)
            take_screenshot(driver, f"tile3_draft_no_page_{row.document_1}")
            click_back(driver)
            return "error: could not open draft invoice"

        # Fill invoice number, add charges, submit
        return fill_invoice_and_submit(driver, row, dry_run=dry_run, step_through=step_through)

    except Exception as e:
        log.error("Error recovering draft for %s: %s", doc_desc, e, exc_info=True)
        take_screenshot(driver, f"tile3_draft_error_{row.document_1}")
        try:
            click_back(driver)
        except Exception:
            pass
        return f"error: draft recovery failed - {str(e)[:80]}"


def process_row(driver: WebDriver, row: InvoiceRow,
                *, dry_run: bool = False, step_through: bool = False) -> str:
    """Process one row from Google Sheet. Returns 'submitted', 'drafted', 'skipped', or error message."""
    doc_desc = row.document_1
    if row.is_collective:
        doc_desc = f"{row.document_1} + {row.document_2}"

    log.info("Processing: %s → invoice %s", doc_desc, row.invoice)

    try:
        # Filter
        if row.is_collective:
            filter_collective_documents(driver, row.document_1, row.document_2)
        else:
            filter_single_document(driver, row.document_1)

        take_screenshot(driver, f"tile3_filtered_{row.document_1}")

        # Select rows
        select_all_visible_rows(driver)

        # Create invoice
        if not click_create_invoice(driver, collective=row.is_collective):
            return "error: Create Invoice button not found"

        # Check if invoice page opened (look for Invoice Details tab)
        on_invoice = driver.execute_script("""
            var spans = document.querySelectorAll('span');
            for (var i = 0; i < spans.length; i++) {
                if (spans[i].offsetParent === null) continue;
                var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
                if (t === 'Invoice Details') return true;
            }
            return false;
        """)

        if not on_invoice:
            # Check if there's an error popup (e.g. "invoice created" → went to drafts)
            error_text = driver.execute_script("""
                var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
                for (var i = dialogs.length - 1; i >= 0; i--) {
                    var d = dialogs[i];
                    if (d.offsetParent === null) continue;
                    return d.textContent.substring(0, 300).trim();
                }
                return null;
            """)
            if error_text:
                log.warning("Popup for %s: %s", doc_desc, error_text[:150])
                dismiss_any_popup(driver)

            log.info("Invoice went to Drafts for %s — will retrieve from Manage Invoices", doc_desc)
            take_screenshot(driver, f"tile3_draft_{row.document_1}")
            return "drafted"

        # Invoice page opened directly — fill it here
        return fill_invoice_and_submit(driver, row, dry_run=dry_run, step_through=step_through)

    except Exception as e:
        log.error("Error processing %s: %s", doc_desc, e, exc_info=True)
        take_screenshot(driver, f"tile3_error_{row.document_1}")
        try:
            click_back(driver)
        except Exception:
            pass
        return f"error: {str(e)[:100]}"


def run(driver: WebDriver, *, dry_run: bool = False, step_through: bool = False, **_kwargs):
    """Execute the full Tile 3 workflow using Google Sheets."""
    import os
    launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")

    rows = read_todo_items()
    if not rows:
        log.info("No items in To Do tab — nothing to invoice")
        return {"submitted": 0, "errors": 0}

    log.info("Found %d items in To Do tab", len(rows))
    results = {"submitted": 0, "drafted": 0, "skipped": 0, "errors": 0}

    navigate_to_tile(driver)

    for i, row in enumerate(rows):
        log.info("════ Invoice %d/%d (doc %s) ════", i + 1, len(rows), row.document_1)

        status = process_row(driver, row, dry_run=dry_run, step_through=step_through)
        log.info("Result: %s", status)

        if status == "submitted":
            results["submitted"] += 1
            move_to_status(row, "done")
        elif status == "dry_run":
            results["skipped"] += 1
        elif status == "drafted":
            # Invoice went to Drafts — recover it from Manage Invoices
            log.info("Attempting to recover drafted invoice for %s", row.document_1)
            draft_status = process_drafted_invoice(driver, row,
                                                    dry_run=dry_run, step_through=step_through)
            log.info("Draft recovery result: %s", draft_status)

            if draft_status == "submitted":
                results["submitted"] += 1
                move_to_status(row, "done (recovered from drafts)")
            elif draft_status == "dry_run":
                results["skipped"] += 1
            else:
                results["errors"] += 1
                mark_error(row, draft_status.replace("error: ", ""))
        elif status.startswith("error:"):
            results["errors"] += 1
            mark_error(row, status.replace("error: ", ""))

        # Navigate back to Invoice Freight Documents tile for next item
        if i < len(rows) - 1:
            if launchpad_url:
                driver.get(launchpad_url)
                wait_for_page_ready(driver)
            navigate_to_tile(driver)

    log.info("Tile 3 complete: %s", results)
    return results
