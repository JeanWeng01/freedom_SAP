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
from bot.google_sheets import (
    InvoiceRow, read_todo_items, mark_tile3_done, mark_tile3_done_with_note,
    mark_tile3_paused, mark_error, mark_fully_done, sort_paused_rows_to_top,
    read_todo_status,
)

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


def count_visible_rows(driver: WebDriver) -> int:
    """Count visible data rows in the current table view."""
    return driver.execute_script("""
        var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
        var n = 0;
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].offsetParent === null) continue;
            if (rows[i].classList.contains('sapMListTblHeader')) continue;
            if (rows[i].textContent.trim().length > 0) n++;
        }
        return n;
    """) or 0


def click_all_tab_tile3(driver: WebDriver):
    """Click the 'All (X)' tab to show filter results regardless of invoicing status.

    The tabs are in an icon tab bar. The 'All' tab label looks like 'All (2)'.
    """
    take_screenshot(driver, "tile3_before_all_tab_search")

    # Wait for tile-specific tab bar to render (NOT the SAP shell tabs)
    # The tile's tabs contain "To be Invoiced" or "Invoicing in Process" — wait for those
    from selenium.webdriver.support.ui import WebDriverWait as _WDW
    try:
        _WDW(driver, 30).until(
            lambda d: d.execute_script("""
                var all = document.querySelectorAll('.sapMITBFilter, .sapMITBItem, span, div');
                for (var i = 0; i < all.length; i++) {
                    var t = all[i].textContent.replace(/\\xAD/g, '').replace(/\\s+/g, ' ').trim();
                    if (t.indexOf('To be Invoiced') !== -1 || t.indexOf('Invoicing in Process') !== -1) {
                        return true;
                    }
                }
                return false;
            """)
        )
        log.info("Tile tab bar rendered (found 'To be Invoiced' or 'Invoicing in Process')")
    except Exception:
        log.warning("Tile tab bar not found after 30s — proceeding anyway")

    # Find the All tab — look specifically within the tabs that include "Invoiced" sibling tabs
    from selenium.webdriver.support.ui import WebDriverWait
    all_tab = None
    js_finder = """
        // Find the tab bar that contains "To be Invoiced" — that's where the "All" tab lives
        // Strategy: find the element containing "To be Invoiced" text, then look for a sibling "All (N)" tab
        var invoiceMarkers = document.querySelectorAll('[role="tab"], .sapMITBFilter, .sapMITBItem');
        var tabBar = null;
        for (var i = 0; i < invoiceMarkers.length; i++) {
            var t = invoiceMarkers[i].textContent.replace(/\\xAD/g, '').replace(/\\s+/g, ' ').trim();
            if (t.indexOf('To be Invoiced') !== -1 || t.indexOf('Invoicing in Process') !== -1) {
                // Walk up to find the common tab container
                tabBar = invoiceMarkers[i].closest('[class*="sapMITBHeader"], [class*="sapMITBContainer"], [role="tablist"]');
                if (!tabBar) tabBar = invoiceMarkers[i].parentElement;
                break;
            }
        }

        if (tabBar) {
            // Look for "All (N)" tab within this bar's siblings
            var tabsInBar = tabBar.querySelectorAll('[role="tab"], .sapMITBFilter, .sapMITBItem');
            for (var i = 0; i < tabsInBar.length; i++) {
                var t = tabsInBar[i].textContent.replace(/\\xAD/g, '').replace(/\\s+/g, ' ').trim();
                if (/^All\\s*\\(\\d+\\)$/.test(t) || /^All\\s*\\(\\d/.test(t)) {
                    return tabsInBar[i];
                }
            }
        }

        // Fallback: any tab-role element with "All (N)" pattern
        var allTabs = document.querySelectorAll('[role="tab"], .sapMITBFilter, .sapMITBItem');
        for (var i = 0; i < allTabs.length; i++) {
            var t = allTabs[i].textContent.replace(/\\xAD/g, '').replace(/\\s+/g, ' ').trim();
            if (/^All\\s*\\(\\d+\\)$/.test(t)) {
                return allTabs[i];
            }
        }

        return null;
    """
    try:
        all_tab = WebDriverWait(driver, 15).until(
            lambda d: d.execute_script(js_finder)
        )
    except Exception:
        pass

    if not all_tab:
        log.warning("'All' tab not found after waiting")
        take_screenshot(driver, "tile3_all_tab_not_found")
        return False

    # Log which element we actually found (for debugging)
    log.info("Found 'All' tab element with text: %s",
             driver.execute_script("return arguments[0].textContent.replace(/\\xAD/g,'').replace(/\\s+/g,' ').trim().substring(0, 50);", all_tab))

    # Native click — JS click doesn't trigger SAP UI5 tab switch
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", all_tab)
        _time.sleep(0.3)
        all_tab.click()
        log.info("Clicked 'All' tab (native)")
    except Exception:
        try:
            ActionChains(driver).move_to_element(all_tab).click().perform()
            log.info("Clicked 'All' tab (ActionChains)")
        except Exception as e:
            log.error("Could not click All tab: %s", e)
            return False

    _time.sleep(2)
    wait_for_page_ready(driver)
    return True


def read_row_doc_and_status(driver: WebDriver) -> list[dict]:
    """Read each visible row's freight document number and Invoicing Status."""
    rows_info = driver.execute_script("""
        var results = [];
        var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].offsetParent === null) continue;
            if (rows[i].classList.contains('sapMListTblHeader')) continue;
            var text = rows[i].textContent.replace(/\\xAD/g, '').trim();
            if (text.length === 0) continue;

            // Find all long digit strings in the row — pick any 10-digit number
            // (freight doc numbers are typically 10 digits like 6100184538)
            var doc = null;
            var digitMatches = text.match(/\\d{10}/g);
            if (digitMatches) {
                // Prefer digits starting with 61 (freight doc prefix)
                for (var j = 0; j < digitMatches.length; j++) {
                    if (digitMatches[j].substring(0, 2) === '61') {
                        doc = digitMatches[j];
                        break;
                    }
                }
                if (!doc) doc = digitMatches[0];  // fallback to first 10-digit match
            }

            // Fallback: scan for 8-12 digit strings
            if (!doc) {
                var m = text.match(/\\d{8,12}/);
                if (m) doc = m[0];
            }

            // Find Invoicing Status — case-insensitive
            var tLower = text.toLowerCase();
            var status = '';
            if (tLower.indexOf('completely invoiced') !== -1) status = 'Completely invoiced';
            else if (tLower.indexOf('invoicing in process') !== -1) status = 'Invoicing in process';
            else if (tLower.indexOf('not yet invoiced') !== -1) status = 'Not Yet Invoiced';

            results.push({doc: doc, status: status, index: i, raw: text.substring(0, 120)});
        }
        return results;
    """)
    return rows_info or []


def select_specific_rows(driver: WebDriver, doc_numbers: list) -> int:
    """Select rows whose freight document number matches one in doc_numbers.

    Returns number of rows actually selected.
    """
    if not doc_numbers:
        return 0

    # Find the rows and their checkboxes
    selected = 0
    for doc in doc_numbers:
        cb = driver.execute_script("""
            var target = arguments[0];
            var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
            for (var i = 0; i < rows.length; i++) {
                if (rows[i].offsetParent === null) continue;
                if (rows[i].classList.contains('sapMListTblHeader')) continue;
                var text = rows[i].textContent.replace(/\\xAD/g, '').trim();
                if (text.indexOf(target) !== -1) {
                    return rows[i].querySelector('.sapMCb');
                }
            }
            return null;
        """, doc)

        if cb:
            try:
                cb.click()
                selected += 1
                log.info("Selected row for doc %s", doc)
            except Exception:
                try:
                    ActionChains(driver).move_to_element(cb).click().perform()
                    selected += 1
                    log.info("Selected row for doc %s (ActionChains)", doc)
                except Exception as e:
                    log.warning("Could not select row for doc %s: %s", doc, e)
        else:
            log.warning("Checkbox not found for doc %s", doc)

    _time.sleep(0.5)
    return selected


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

    # Find the Invoice input field — prefer #invoiceId input on the draft page
    inv_input = driver.execute_script("""
        // Strategy 1: find by input id containing 'invoiceId' (draft invoice page)
        var byId = document.querySelector('input[id*="invoiceId"]:not([type="hidden"])');
        if (byId && byId.offsetParent !== null) return byId;

        // Strategy 2: find by "Invoice:" label
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
        # Scroll into view to avoid click interception from overlapping elements
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inv_input)
        _time.sleep(0.5)

        # Focus + clear + send keys — avoid .click() which may be intercepted
        try:
            driver.execute_script("arguments[0].focus();", inv_input)
            _time.sleep(0.3)
            inv_input.clear()
            inv_input.send_keys(invoice_num)
            log.info("Entered invoice number: %s", invoice_num)
        except Exception as e:
            log.warning("Standard input failed: %s — trying JS value set", e)
            # Fallback: set value via JS and dispatch input event
            driver.execute_script("""
                var el = arguments[0];
                el.value = arguments[1];
                el.dispatchEvent(new Event('input', {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
            """, inv_input, invoice_num)
            log.info("Entered invoice number via JS: %s", invoice_num)
    else:
        log.error("Invoice input field not found")
        take_screenshot(driver, "tile3_invoice_input_missing")


def add_charge(driver: WebDriver, charge_type: str, amount: float, doc_number: str = ""):
    """Add a charge in the Charges tab for a specific freight document."""
    # Find the Charges tab element
    charges_tab = driver.execute_script("""
        var spans = document.querySelectorAll('span, div');
        for (var i = 0; i < spans.length; i++) {
            if (spans[i].offsetParent === null) continue;
            var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Charges') {
                var tab = spans[i].closest('[role="tab"], [class*="sapMITBFilter"], [class*="sapMITBItem"]');
                if (tab) return tab;
                return spans[i];
            }
        }
        return null;
    """)

    if not charges_tab:
        log.error("Charges tab not found")
        take_screenshot(driver, "tile3_charges_tab_missing")
        return

    # Native click on Charges tab
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", charges_tab)
        _time.sleep(0.3)
        charges_tab.click()
        log.info("Clicked Charges tab (native)")
    except Exception:
        ActionChains(driver).move_to_element(charges_tab).click().perform()
        log.info("Clicked Charges tab (ActionChains)")

    _time.sleep(2)
    wait_for_page_ready(driver)
    take_screenshot(driver, f"tile3_charges_tab_open_{charge_type[:15]}")

    # Click Add button for the correct freight doc section
    # Each freight doc has its own "Add" button. Find the one near the target doc number.
    add_btn = driver.execute_script("""
        var targetDoc = arguments[0];

        if (targetDoc) {
            // Find all sections/headers containing the freight doc number
            var all = document.querySelectorAll('*');
            for (var i = 0; i < all.length; i++) {
                if (all[i].offsetParent === null) continue;
                var t = all[i].textContent.replace(/\\xAD/g, '').trim();
                // Match "Freight Document <number>" header
                if (t.indexOf('Freight Document') !== -1 && t.indexOf(targetDoc) !== -1 && t.length < 100) {
                    // Find the "Add" button within this section's parent
                    var parent = all[i];
                    for (var j = 0; j < 10 && parent; j++) {
                        var addBtn = parent.querySelector('button');
                        if (addBtn) {
                            var bt = addBtn.textContent.replace(/\\xAD/g, '').trim();
                            if (bt.indexOf('Add') !== -1) return addBtn;
                        }
                        // Check siblings too
                        var sibling = parent.nextElementSibling;
                        while (sibling) {
                            var btn = sibling.querySelector('button');
                            if (btn) {
                                var st = btn.textContent.replace(/\\xAD/g, '').trim();
                                if (st.indexOf('Add') !== -1) return btn;
                            }
                            sibling = sibling.nextElementSibling;
                        }
                        parent = parent.parentElement;
                    }
                }
            }
        }

        // Fallback: if no doc_number specified or not found, use the last visible Add button
        var btns = document.querySelectorAll('button');
        var lastAdd = null;
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].offsetParent === null) continue;
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Add') lastAdd = btns[i];
        }
        return lastAdd;
    """, doc_number)
    if add_btn:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_btn)
            _time.sleep(0.3)
            add_btn.click()
        except Exception:
            ActionChains(driver).move_to_element(add_btn).click().perform()
        log.info("Clicked Add → new blank charge row")
        _time.sleep(1)
        wait_for_page_ready(driver)
    else:
        log.error("Add button not found in Charges tab")
        take_screenshot(driver, "tile3_add_charge_missing")
        return

    # Some Fiori versions show a dropdown after Add — if "Charge" appears, click it
    charge_option_clicked = driver.execute_script("""
        var items = document.querySelectorAll('li, [role="option"]');
        for (var i = 0; i < items.length; i++) {
            if (items[i].offsetParent === null) continue;
            var t = items[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Charge') { items[i].click(); return true; }
        }
        return false;
    """)
    if charge_option_clicked:
        log.info("Clicked 'Charge' from dropdown")
        _time.sleep(1)
        wait_for_page_ready(driver)

    take_screenshot(driver, f"tile3_blank_charge_row_{charge_type[:15]}")

    # DEBUG: dump HTML around the charges area to find the exact blank field element
    debug_html = driver.execute_script("""
        // Find the area near "Freight Document" text
        var all = document.querySelectorAll('*');
        var results = [];
        for (var i = 0; i < all.length; i++) {
            if (all[i].offsetParent === null) continue;
            var t = all[i].textContent.replace(/\\xAD/g, '').trim();
            if (t.indexOf('Freight Document') !== -1 && t.length < 50) {
                // Dump the parent's HTML
                var parent = all[i].closest('tr, div, section') || all[i].parentElement;
                if (parent) results.push(parent.outerHTML.substring(0, 500));
            }
        }
        // Also find all visible inputs with empty values
        var emptyInputs = [];
        var inputs = document.querySelectorAll('input:not([type="hidden"])');
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent === null) continue;
            if (inputs[i].value.trim() === '' || inputs[i].value.trim().length < 3) {
                emptyInputs.push({
                    id: inputs[i].id,
                    type: inputs[i].type,
                    value: inputs[i].value,
                    ariaLabel: inputs[i].getAttribute('aria-label'),
                    placeholder: inputs[i].placeholder,
                    className: inputs[i].className.substring(0, 80)
                });
            }
        }
        return {freightDocHTML: results, emptyInputs: emptyInputs};
    """)
    log.info("DEBUG charge area — freight doc HTML: %s", str(debug_html.get('freightDocHTML', []))[:500])
    log.info("DEBUG charge area — empty inputs: %s", debug_html.get('emptyInputs', []))

    # Click the BLANK input field in the Charge Description column
    # This field is on the SAME LINE as the freight document header row,
    # NOT in a separate data row. It looks like a text input but is read-only —
    # clicking it opens a category selection popup.
    blank_field = driver.execute_script("""
        // The blank field is a sapMInputBaseInner inside a sapMInput control
        // ID pattern: __input0-__cloneN-inner
        // We need to click the PARENT sapMInput wrapper or its value help icon,
        // not the inner input directly

        // Find the empty sapMInputBaseInner (the one we identified in debug)
        var inputs = document.querySelectorAll('input.sapMInputBaseInner');
        var target = null;
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent === null) continue;
            if (inputs[i].value.trim() === '' && inputs[i].type === 'text') {
                // Skip the search bar
                if (inputs[i].id.indexOf('Search') !== -1 || inputs[i].id.indexOf('search') !== -1) continue;
                target = inputs[i];
                break;
            }
        }
        if (!target) return null;

        // Try to find the value help icon button next to this input
        // SAP places it as a sibling or inside the parent sapMInput wrapper
        var parent = target.closest('.sapMInput, .sapMInputBase, [class*="sapMInput"]');
        if (parent) {
            var vhIcon = parent.querySelector('[class*="ValueHelp"], .sapMInputBaseIcon, .sapUiIcon');
            if (vhIcon && vhIcon.offsetParent !== null) return vhIcon;
            // No icon — return the parent wrapper itself
            return parent;
        }
        return target;
    """)

    if blank_field:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", blank_field)
            _time.sleep(0.5)
            blank_field.click()
            log.info("Clicked blank charge description input (native)")
        except Exception:
            try:
                ActionChains(driver).move_to_element(blank_field).click().perform()
                log.info("Clicked blank field (ActionChains)")
            except Exception:
                driver.execute_script("arguments[0].click();", blank_field)
                log.info("Clicked blank field (JS)")
        _time.sleep(2)
        wait_for_page_ready(driver)
        take_screenshot(driver, f"tile3_after_blank_field_click_{charge_type[:15]}")
    else:
        log.warning("Could not find blank charge description input")

    take_screenshot(driver, f"tile3_category_popup_opening_{charge_type[:15]}")

    # In the popup that opens: search for charge_type in the search bar
    search_input = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            if (d.offsetParent === null) continue;
            // Find the search input in the dialog
            var searchInput = d.querySelector('input[type="search"], input[placeholder*="earch"], input.sapMInputBaseInner');
            if (searchInput && searchInput.offsetParent !== null) return searchInput;
        }
        return null;
    """)

    if search_input:
        try:
            driver.execute_script("arguments[0].focus();", search_input)
            _time.sleep(0.3)
            search_input.clear()
            search_input.send_keys(charge_type)
            log.info("Typed '%s' in charge category search", charge_type)
            _time.sleep(1.5)  # let search results filter
            take_screenshot(driver, f"tile3_charge_searched_{charge_type[:15]}")
        except Exception as e:
            log.warning("Search input type failed: %s", e)
    else:
        log.warning("Search input not found in category popup — charge type may still be visible")

    # Click the matching search result row — must use native click
    # The popup shows a table: Description | Charge Code | Charge Type
    # Find the row element via JS, then click with Selenium native click
    result_row = driver.execute_script("""
        var target = arguments[0].toLowerCase();
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            if (d.offsetParent === null) continue;
            // Search in table rows and list items
            var items = d.querySelectorAll('tr, li, [role="row"], .sapMLIB');
            for (var j = 0; j < items.length; j++) {
                if (items[j].offsetParent === null) continue;
                var t = items[j].textContent.replace(/\\xAD/g, '').trim().toLowerCase();
                if (t.indexOf(target) !== -1 && t.length < 200) {
                    return items[j];
                }
            }
        }
        return null;
    """, charge_type)

    if result_row:
        # Native click on the result row
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", result_row)
            _time.sleep(0.3)
            result_row.click()
            log.info("Selected category row (native click)")
        except Exception:
            try:
                ActionChains(driver).move_to_element(result_row).click().perform()
                log.info("Selected category row (ActionChains)")
            except Exception:
                driver.execute_script("arguments[0].click();", result_row)
                log.info("Selected category row (JS)")
        _time.sleep(2)
        wait_for_page_ready(driver)
    else:
        log.error("Could not find '%s' in category search results", charge_type)
        take_screenshot(driver, f"tile3_category_not_found_{charge_type[:15]}")
        dismiss_any_popup(driver)
        return

    take_screenshot(driver, f"tile3_category_selected_{charge_type[:15]}")

    # Enter the charge amount in the Rate Amount/Unit column
    _time.sleep(3)  # SAP needs time to update the row after category selection
    wait_for_page_ready(driver)

    # Debug: dump all visible inputs to find the rate field
    debug_inputs = driver.execute_script("""
        var inputs = document.querySelectorAll('input:not([type="hidden"]):not([type="file"])');
        var results = [];
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent === null) continue;
            results.push({
                id: inputs[i].id.substring(0, 60),
                value: inputs[i].value,
                ariaLabel: inputs[i].getAttribute('aria-label'),
                type: inputs[i].type
            });
        }
        return results;
    """)
    log.info("DEBUG visible inputs for rate: %s", debug_inputs)

    rate_input_el = driver.execute_script("""
        // The rate input has ID containing 'rateInput' and value '0.00' or '0'
        // For the correct leg, find the one with value 0.00 (the just-added charge row)
        var inputs = document.querySelectorAll('input[id*="rateInput"]');
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent === null) continue;
            var val = inputs[i].value.trim();
            if (val === '0.00' || val === '0' || val === '') {
                return inputs[i];
            }
        }

        // Fallback: any visible input with id containing 'rate' and zero value
        var allInputs = document.querySelectorAll('input');
        for (var i = 0; i < allInputs.length; i++) {
            if (allInputs[i].offsetParent === null) continue;
            var id = (allInputs[i].id || '').toLowerCase();
            var val = allInputs[i].value.trim();
            if (id.indexOf('rate') !== -1 && (val === '0.00' || val === '0' || val === '')) {
                return allInputs[i];
            }
        }

        return null;
    """)

    if rate_input_el:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", rate_input_el)
            _time.sleep(0.3)
            driver.execute_script("arguments[0].focus();", rate_input_el)
            _time.sleep(0.2)
            rate_input_el.clear()
            rate_input_el.send_keys(str(amount))
            log.info("Entered charge amount: %s", amount)
        except Exception as e:
            log.warning("Could not enter amount: %s", e)
    else:
        log.error("Rate Amount input not found in new charge row")
        take_screenshot(driver, f"tile3_rate_input_missing_{charge_type[:15]}")

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


def click_save(driver: WebDriver, invoice_num: str = ""):
    """Click the Save button (next to Submit) for paused items."""
    take_screenshot(driver, f"tile3_before_save_{invoice_num}")
    btn = driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            if (btns[i].offsetParent === null) continue;
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Save') return btns[i];
        }
        return null;
    """)
    if btn:
        try:
            btn.click()
        except Exception:
            ActionChains(driver).move_to_element(btn).click().perform()
        log.info("Clicked Save for invoice %s (paused item)", invoice_num)
        _time.sleep(2)
        wait_for_page_ready(driver)
        dismiss_any_popup(driver)
        take_screenshot(driver, f"tile3_after_save_{invoice_num}")
    else:
        log.error("Save button not found")
        take_screenshot(driver, "tile3_save_missing")


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
    """Fill invoice number, add charges, and submit (or save if paused)."""
    enter_invoice_number(driver, row.invoice)

    if row.has_charges_leg1:
        add_charge(driver, row.charge_type_1, row.charge_amount_1, doc_number=row.document_1)

    if row.is_collective and row.has_charges_leg2:
        add_charge(driver, row.charge_type_2, row.charge_amount_2, doc_number=row.document_2)

    if row.pause:
        # Paused: Save instead of Submit — human will review and submit manually
        click_save(driver, invoice_num=row.invoice)
        click_back(driver)
        return "paused"

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
        expected_docs = [row.document_1]
        if row.is_collective:
            filter_collective_documents(driver, row.document_1, row.document_2)
            expected_docs.append(row.document_2)
        else:
            filter_single_document(driver, row.document_1)

        take_screenshot(driver, f"tile3_filtered_{row.document_1}")

        # Step 1: Check if docs appear in the default "To be Invoiced" tab
        visible_rows = count_visible_rows(driver)
        log.info("Visible rows after filter in default tab: %d", visible_rows)

        usable_docs = []  # docs that can be invoiced
        skipped_docs = []  # docs skipped (already Completely invoiced)

        if visible_rows == len(expected_docs):
            # All expected docs are in To be Invoiced tab — good, use all
            log.info("All %d docs found in 'To be Invoiced' tab", visible_rows)
            usable_docs = expected_docs
            # Use select_all since we want all of them
            select_all_visible_rows(driver)
        else:
            # Not all docs in default tab — check All tab
            log.info("Only %d of %d docs in default tab — checking All tab",
                     visible_rows, len(expected_docs))
            if not click_all_tab_tile3(driver):
                return "error: could not click All tab to check status"

            _time.sleep(2)  # extra time for tab content to render
            wait_for_page_ready(driver)
            take_screenshot(driver, f"tile3_all_tab_{row.document_1}")

            rows_info = read_row_doc_and_status(driver)
            log.info("All tab rows: %s", rows_info)

            # Check status of each expected doc
            for doc in expected_docs:
                found = next((r for r in rows_info if r.get("doc") == doc), None)
                if not found:
                    skipped_docs.append((doc, "not found in All tab"))
                    log.warning("Doc %s not found anywhere", doc)
                elif found.get("status") == "Completely invoiced":
                    skipped_docs.append((doc, "already Completely invoiced"))
                    log.warning("Doc %s is already Completely invoiced — skipping", doc)
                else:
                    # Invoicing in process or Not Yet Invoiced — usable
                    usable_docs.append(doc)

            if not usable_docs:
                # Nothing to invoice for this row
                error_detail = "; ".join(f"{d}: {r}" for d, r in skipped_docs) or "no usable docs"
                return f"error: {error_detail}"

            log.info("Usable docs: %s | Skipped: %s", usable_docs, skipped_docs)

            # Select only the usable docs
            selected = select_specific_rows(driver, usable_docs)
            if selected == 0:
                return "error: could not select any usable docs"

        # Store skipped info on the row for later status reporting
        if skipped_docs:
            row._skipped_note = "skipped: " + "; ".join(f"{d} ({r})" for d, r in skipped_docs)
        else:
            row._skipped_note = ""

        # Create invoice (Collective if 2 usable, single if 1)
        is_collective = len(usable_docs) > 1
        if not click_create_invoice(driver, collective=is_collective):
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
    """Execute the full Tile 3 workflow using Google Sheets.

    Paused items (Pause=1): create invoice, fill details, Save (not Submit), keep in To Do.
    Non-paused items: create invoice, fill details, Submit, keep in To Do for tile 4.
    Items only move to Status tab after both tile 3 and tile 4 complete.
    """
    import os
    launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")

    rows = read_todo_items()
    if not rows:
        log.info("No items in To Do tab — nothing to invoice")
        return {"submitted": 0, "errors": 0}

    # Process non-paused items first, then paused items stay in To Do
    active_rows = [r for r in rows if not r.pause]
    paused_rows = [r for r in rows if r.pause]

    log.info("Found %d items: %d active, %d paused",
             len(rows), len(active_rows), len(paused_rows))

    results = {"submitted": 0, "paused": 0, "drafted": 0, "skipped": 0, "errors": 0}

    if not active_rows and not paused_rows:
        log.info("No items to process")
        return results

    navigate_to_tile(driver)

    # Process all items (active first, then paused)
    all_to_process = active_rows + paused_rows

    for i, row in enumerate(all_to_process):
        pause_label = " [PAUSED]" if row.pause else ""
        log.info("════ Invoice %d/%d (doc %s)%s ════",
                 i + 1, len(all_to_process), row.document_1, pause_label)

        # Mark in-progress (yellow) so user can see activity
        from bot.google_sheets import _write_todo_status
        _write_todo_status(row.document_1, "tile3_in_progress", color="yellow")

        status = process_row(driver, row, dry_run=dry_run, step_through=step_through)
        log.info("Result: %s", status)

        if status == "submitted":
            results["submitted"] += 1
            # Check if tile 4 already ran (col H = "pod_uploaded")
            current_status = read_todo_status(row.document_1)
            if "pod_uploaded" in current_status.lower():
                log.info("Both tile 3 and tile 4 done — moving to Status")
                mark_fully_done(row, pod_uploaded_1="done",
                                pod_uploaded_2="done" if row.is_collective else "")
            else:
                note = getattr(row, '_skipped_note', '')
                if note:
                    mark_tile3_done_with_note(row, note)
                else:
                    mark_tile3_done(row)  # stays in To Do for tile 4
        elif status == "paused":
            results["paused"] += 1
            mark_tile3_paused(row)
            log.info("Paused item saved — marked 'paused' in To Do for human review")
        elif status == "dry_run":
            results["skipped"] += 1
        elif status == "drafted":
            log.info("Attempting to recover drafted invoice for %s", row.document_1)
            draft_status = process_drafted_invoice(driver, row,
                                                    dry_run=dry_run, step_through=step_through)
            log.info("Draft recovery result: %s", draft_status)

            if draft_status == "submitted":
                results["submitted"] += 1
                mark_tile3_done(row)
            elif draft_status == "paused":
                results["paused"] += 1
            elif draft_status == "dry_run":
                results["skipped"] += 1
            else:
                results["errors"] += 1
                # Don't move to Status — leave in To Do with error in col H so it retries next cycle
                from bot.google_sheets import _write_todo_status
                _write_todo_status(row.document_1,
                                   f"tile3_error: {draft_status.replace('error: ', '')}")
        elif status.startswith("error:"):
            results["errors"] += 1
            # Don't move to Status — leave in To Do for retry
            from bot.google_sheets import _write_todo_status
            _write_todo_status(row.document_1,
                               f"tile3_error: {status.replace('error: ', '')}")

        # Navigate back to Invoice Freight Documents tile for next item
        if i < len(all_to_process) - 1:
            if launchpad_url:
                driver.get(launchpad_url)
                wait_for_page_ready(driver)
            navigate_to_tile(driver)

    # Sort paused rows to top of To Do for easy human access
    sort_paused_rows_to_top()

    log.info("Tile 3 complete: %s", results)
    return results
