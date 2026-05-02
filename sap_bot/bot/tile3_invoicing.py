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
    """Click the Invoice Freight Documents tile and wait for it to actually load.

    Uses body text check (not offsetParent) for headless compatibility.
    """
    click_tile(driver, TILE_NAME)
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("""
                var body = document.body ? document.body.textContent.replace(/\\xAD/g, '') : '';
                // Look for text unique to the Invoice Freight Documents tile
                if (body.indexOf('Freight Document') !== -1 && body.indexOf('Invoice') !== -1) return true;
                if (body.indexOf('To be Invoiced') !== -1) return true;
                if (body.indexOf('Invoicing in Process') !== -1) return true;
                if (body.indexOf('Create Invoice') !== -1) return true;
                return false;
            """)
        )
        log.info("Tile 3 page content detected in body text")
    except TimeoutException:
        log.warning("Tile 3 page content not detected after 30s — retrying")
        import os
        launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")
        if launchpad_url:
            driver.get(launchpad_url)
            wait_for_page_ready(driver)
        click_tile(driver, TILE_NAME)
        _time.sleep(15)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile3_page_loaded")
    log.info("Tile 3 page loaded")


def dismiss_any_popup(driver: WebDriver) -> bool:
    """Dismiss any error/info popup on the page."""
    result = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = dialogs.length - 1; i >= 0; i--) {
            var d = dialogs[i];
            // offsetParent check removed — unreliable in headless
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
    """Count data rows in the current table view.

    Does NOT check offsetParent (unreliable in headless Chrome).
    Instead checks for non-empty text content and skips headers.
    """
    return driver.execute_script("""
        var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
        var n = 0;
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].classList.contains('sapMListTblHeader')) continue;
            // Skip "no data" placeholder rows
            var t = rows[i].textContent.trim();
            if (t.length > 0 && t.indexOf('no data') === -1) n++;
        }
        return n;
    """) or 0


def click_all_tab_tile3(driver: WebDriver):
    """Click the 'All (X)' tab to show filter results regardless of invoicing status.

    The tabs are in an icon tab bar. The 'All' tab label looks like 'All (2)'.
    """
    take_screenshot(driver, "tile3_before_all_tab_search")

    # Wait for tile-specific tab bar to render (NOT the SAP shell tabs)
    # Look for ANY text containing "Invoic" (matches "To be Invoiced", "Invoicing in Process",
    # "Completely Invoiced" etc.) — handles soft hyphens and partial rendering
    from selenium.webdriver.support.ui import WebDriverWait as _WDW
    try:
        _WDW(driver, 30).until(
            lambda d: d.execute_script("""
                // Check page text for any invoice-related tab content
                var body = document.body ? document.body.textContent : '';
                body = body.replace(/\\xAD/g, '');
                if (body.indexOf('Invoiced') !== -1 || body.indexOf('Invoicing') !== -1) return true;
                // Also check for the "All (" pattern which appears in the tab bar
                if (/All\\s*\\(\\d/.test(body)) return true;
                return false;
            """)
        )
        log.info("Tile tab bar content detected")
    except Exception:
        log.warning("Tile tab bar text not found after 30s — will try anyway")

    # Find the All tab element
    # Strategy: search the entire DOM for clickable elements whose clean text matches "All (N)"
    # Exclude SAP shell tabs (class sapUshellAnchorItem) which are the top navigation
    from selenium.webdriver.support.ui import WebDriverWait
    all_tab = None
    js_finder = """
        // Search ALL elements, exclude SAP shell tabs
        var all = document.querySelectorAll('*');
        for (var i = 0; i < all.length; i++) {
            // Skip shell navigation tabs (handle SVG elements where className isn't a string)
            var cls = (typeof all[i].className === 'string') ? all[i].className : '';
            if (cls.indexOf('sapUshellAnchor') !== -1) continue;
            var t = all[i].textContent.replace(/\\xAD/g, '').replace(/\\s+/g, ' ').trim();
            // Match "All (N)" — but only if it's a SHORT text (not a parent containing lots of text)
            if (/^All\\s*\\(\\d+\\)$/.test(t) && t.length < 15) {
                // Prefer clicking the closest ITB element or the element itself
                var clickable = all[i].closest('.sapMITBFilter, .sapMITBItem');
                return clickable || all[i];
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
        # Debug: dump what text fragments exist that contain "All"
        debug = driver.execute_script("""
            var results = [];
            var all = document.querySelectorAll('*');
            for (var i = 0; i < all.length; i++) {
                var cls2 = (typeof all[i].className === 'string') ? all[i].className : '';
                if (cls2.indexOf('sapUshellAnchor') !== -1) continue;
                var t = all[i].textContent.replace(/\\xAD/g, '').replace(/\\s+/g, ' ').trim();
                if (/^All/.test(t) && t.length < 20) {
                    results.push({text: t, tag: all[i].tagName, cls: (all[i].className || '').substring(0,40)});
                }
            }
            return results.slice(0, 15);
        """)
        log.warning("'All' tab not found. Elements starting with 'All': %s", debug)
        take_screenshot(driver, "tile3_all_tab_not_found")
        return False

    log.info("Found 'All' tab: %s",
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
            // offsetParent check removed — unreliable in headless
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
                // offsetParent check removed — unreliable in headless
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
            // offsetParent check removed — unreliable in headless
            if (rows[i].classList.contains('sapMListTblHeader')) continue;
            var cb = rows[i].querySelector('.sapMCb');
            if (cb) cbs.push(cb);
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
            // offsetParent check removed — unreliable in headless
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
            // offsetParent check removed — unreliable in headless
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
        if (byId) return byId;

        // Strategy 2: find by "Invoice:" label
        var labels = document.querySelectorAll('label, span');
        for (var i = 0; i < labels.length; i++) {
            var clean = labels[i].textContent.replace(/\\xAD/g, '').trim();
            if (clean === 'Invoice:' || clean === 'Invoice') {
                var parent = labels[i].parentElement;
                for (var j = 0; j < 5 && parent; j++) {
                    var input = parent.querySelector('input:not([type="hidden"])');
                    if (input) return input;
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


def add_charge(driver: WebDriver, charge_type: str, amount: float,
               doc_number: str = "", leg_num: int = 1) -> bool:
    """Add a charge in the Charges tab for a specific freight document.

    leg_num: 1 = leg 1 (top of page, no scroll needed). 2 = leg 2 (typically
    collapsed at the bottom of the page; we scroll to the extreme bottom so
    leg 2's Add button is revealed just above the sticky Save bar, instead of
    being hidden under it).

    Returns True if the charge category was selected AND the amount was entered.
    Returns False on any failure (Charges tab missing, Add btn missing, category
    not found, amount entry rejected, etc.) — caller must abort the invoice.
    """
    # Find the Charges tab element
    charges_tab = driver.execute_script("""
        var spans = document.querySelectorAll('span, div');
        for (var i = 0; i < spans.length; i++) {
            // offsetParent check removed — unreliable in headless
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
        return False

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

    # For leg 2: scroll page to extreme bottom so leg 2's section + Add button
    # are revealed. Leg 2 is typically collapsed by default — that's fine,
    # clicking its Add button will auto-expand it. Without this scroll, leg 2's
    # Add button can be hidden under SAP's sticky Save/Submit/Cancel footer
    # bar, and the click would land on Save (triggering "Document saved" toast
    # instead of opening the Add dropdown).
    # For leg 1: do NOT scroll. Leg 1's Add button is already visible at the
    # top of the page when the Charges tab opens. Scrolling would push it out
    # of frame.
    if leg_num == 2:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        _time.sleep(0.5)
        log.info("Scrolled to page bottom for leg 2 charge add")

    # Find Add button for the target freight doc section.
    # Strategy: find every button whose text is exactly "Add", walk up from each,
    # and pick the one whose CLOSEST "Freight Document <NNN>" ancestor matches
    # the target doc. This avoids the old "first button in subtree" heuristic
    # which was fragile (hitting collapse arrows, trash icons, leg-1 buttons).
    add_btn = driver.execute_script("""
        var targetDoc = arguments[0];

        // Collect every button whose text is exactly "Add"
        var addButtons = [];
        var allBtns = document.querySelectorAll('button');
        for (var i = 0; i < allBtns.length; i++) {
            var t = allBtns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Add') addButtons.push(allBtns[i]);
        }

        if (targetDoc && addButtons.length > 0) {
            // For each Add button, walk up looking for "Freight Document <10digits>"
            // in an ancestor's text. The CLOSEST match determines its section.
            var pattern = /Freight Document\\s*(\\d{10})/;
            for (var k = 0; k < addButtons.length; k++) {
                var ancestor = addButtons[k].parentElement;
                for (var m = 0; m < 15 && ancestor; m++) {
                    var ancestorText = ancestor.textContent.replace(/\\xAD/g, '');
                    var match = ancestorText.match(pattern);
                    if (match) {
                        if (match[1] === targetDoc) return addButtons[k];
                        break;  // wrong section — try next Add button
                    }
                    ancestor = ancestor.parentElement;
                }
            }
        }

        // Fallback (single-doc invoice or doc_number not supplied):
        // last "Add" button on page (rightmost/lowest in DOM)
        if (addButtons.length > 0) return addButtons[addButtons.length - 1];
        return null;
    """, doc_number)
    if add_btn:
        # Click directly. For leg 1, the page wasn't scrolled — Add is already
        # visible at top. For leg 2, we scrolled to the bottom above so Add
        # is in the bottom-right area, just above the sticky Save bar (visible,
        # not covered). If leg 2 happened to be already expanded (Add high on
        # page), the doc-aware crawler still picked the right button; native
        # click + Selenium's auto-scroll handles that case.
        try:
            add_btn.click()
        except Exception:
            ActionChains(driver).move_to_element(add_btn).click().perform()
        log.info("Clicked Add → new blank charge row")
        _time.sleep(1)
        wait_for_page_ready(driver)
    else:
        log.error("Add button not found in Charges tab")
        take_screenshot(driver, "tile3_add_charge_missing")
        return False

    # Some Fiori versions show a dropdown after Add — if "Charge" appears, click it
    charge_option_clicked = driver.execute_script("""
        var items = document.querySelectorAll('li, [role="option"]');
        for (var i = 0; i < items.length; i++) {
            // offsetParent check removed — unreliable in headless
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
            // offsetParent check removed — unreliable in headless
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
            // offsetParent check removed — unreliable in headless
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

        // Find the charge description input — identified by ID containing "__clone"
        // This distinguishes it from filter bar inputs (which have IDs like "FilterBar-...")
        var inputs = document.querySelectorAll('input.sapMInputBaseInner');
        var target = null;
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].value.trim() === '' && inputs[i].type === 'text') {
                var id = inputs[i].id || '';
                // Must have "__clone" in ID (charge table inputs) — skip filter bar inputs
                if (id.indexOf('__clone') === -1) continue;
                // Skip search bars
                if (id.indexOf('Search') !== -1 || id.indexOf('search') !== -1) continue;
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
            if (vhIcon) return vhIcon;
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
            // offsetParent check removed — unreliable in headless
            // Find the search input in the dialog
            var searchInput = d.querySelector('input[type="search"], input[placeholder*="earch"], input.sapMInputBaseInner');
            if (searchInput) return searchInput;
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
            // offsetParent check removed — unreliable in headless
            // Search in table rows and list items
            var items = d.querySelectorAll('tr, li, [role="row"], .sapMLIB');
            for (var j = 0; j < items.length; j++) {
                // offsetParent check removed — unreliable in headless
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
        return False

    take_screenshot(driver, f"tile3_category_selected_{charge_type[:15]}")

    # Enter the charge amount in the Rate Amount/Unit column
    _time.sleep(3)  # SAP needs time to update the row after category selection
    wait_for_page_ready(driver)

    # Debug: dump all visible inputs to find the rate field
    debug_inputs = driver.execute_script("""
        var inputs = document.querySelectorAll('input:not([type="hidden"]):not([type="file"])');
        var results = [];
        for (var i = 0; i < inputs.length; i++) {
            // offsetParent check removed — unreliable in headless
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

    # Find the rate input on the row we just added.
    #
    # Why not "first 0.00 rateInput on page": SAP started rendering 5 read-only
    # "Excluded Charge" rows per freight doc, each with rateInput=0.00. Picking
    # the first one would grab a leg-1 excluded charge (read-only → not interactable).
    #
    # Reliable approach: SAP UI5 assigns clone numbers monotonically by insertion
    # order. The just-added description input (now containing the charge type
    # we just selected) has the highest clone number; its sibling rate input on
    # the same row has the next-highest clone number.
    rate_input_el = driver.execute_script("""
        var chargeType = arguments[0];

        // Step 1: find the description input we just filled (highest clone number)
        var descs = document.querySelectorAll('input.sapMInputBaseInner');
        var newestDesc = null, newestNum = -1;
        for (var i = 0; i < descs.length; i++) {
            if (descs[i].value !== chargeType) continue;
            var m = (descs[i].id || '').match(/__clone(\\d+)/);
            if (!m) continue;
            var n = parseInt(m[1]);
            if (n > newestNum) { newestNum = n; newestDesc = descs[i]; }
        }
        if (!newestDesc) return null;

        // Step 2: find the rate input with the smallest clone number > newestNum
        // (= the rate input on the just-added row)
        var rates = document.querySelectorAll('input[id*="rateInput"]');
        var best = null, bestNum = Infinity;
        for (var j = 0; j < rates.length; j++) {
            var rm = (rates[j].id || '').match(/__clone(\\d+)/);
            if (!rm) continue;
            var rn = parseInt(rm[1]);
            if (rn > newestNum && rn < bestNum) { bestNum = rn; best = rates[j]; }
        }
        if (!best) return null;

        // Step 3: sanity check — the just-added row's rate input must still be empty/zero
        var v = best.value.trim();
        if (v !== '0.00' && v !== '0' && v !== '') return null;
        return best;
    """, charge_type)

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
            log.error("Could not enter amount (aborting invoice): %s", e)
            take_screenshot(driver, f"tile3_amount_entry_failed_{charge_type[:15]}")
            return False
    else:
        log.error("Rate Amount input not found in new charge row")
        take_screenshot(driver, f"tile3_rate_input_missing_{charge_type[:15]}")
        return False

    _time.sleep(0.5)
    take_screenshot(driver, f"tile3_charge_added_{charge_type[:15]}")
    return True


@destructive_action("Submit invoice {invoice_num}")
def click_submit(driver: WebDriver, *, invoice_num: str = "") -> bool:
    """Click the Submit button.

    Returns True if Submit was clicked AND no popup appeared after (likely success).
    Returns False if the button was missing OR a popup was dismissed after Submit
    (popup almost always means SAP rejected the submission). Caller must treat
    False as failure and NOT mark the row as invoiced.
    """
    take_screenshot(driver, f"tile3_before_submit_{invoice_num}")
    btn = driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            // offsetParent check removed — unreliable in headless
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Submit') return btns[i];
        }
        return null;
    """)
    if not btn:
        log.error("Submit button not found")
        take_screenshot(driver, "tile3_submit_missing")
        return False

    try:
        btn.click()
    except Exception:
        ActionChains(driver).move_to_element(btn).click().perform()
    log.info("Clicked Submit for invoice %s", invoice_num)
    _time.sleep(2)
    wait_for_page_ready(driver)
    popup_was_dismissed = dismiss_any_popup(driver)
    take_screenshot(driver, f"tile3_after_submit_{invoice_num}")
    if popup_was_dismissed:
        log.error("Popup appeared after Submit for invoice %s — treating as FAILED submission",
                  invoice_num)
        return False
    return True


def click_save(driver: WebDriver, invoice_num: str = ""):
    """Click the Save button (next to Submit) for paused items."""
    take_screenshot(driver, f"tile3_before_save_{invoice_num}")
    btn = driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            // offsetParent check removed — unreliable in headless
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
    """Fill invoice number, add charges, and submit (or save if paused).

    If any charge add fails OR the Submit click triggers a SAP error popup, this
    returns 'error: ...' so the caller will NOT mark the row as invoiced. This
    prevents false-positive 'invoiced' reporting when the invoice actually
    stayed in Drafts due to a silent failure.
    """
    enter_invoice_number(driver, row.invoice)

    if row.has_charges_leg1:
        if not add_charge(driver, row.charge_type_1, row.charge_amount_1,
                          doc_number=row.document_1, leg_num=1):
            return f"error: failed to add leg 1 charge ({row.charge_type_1})"

    if row.is_collective and row.has_charges_leg2:
        if not add_charge(driver, row.charge_type_2, row.charge_amount_2,
                          doc_number=row.document_2, leg_num=2):
            return f"error: failed to add leg 2 charge ({row.charge_type_2})"

    if row.pause:
        # Paused: Save instead of Submit — human will review and submit manually
        click_save(driver, invoice_num=row.invoice)
        click_back(driver)
        return "paused"

    submit_result = click_submit(driver, invoice_num=row.invoice,
                                 dry_run=dry_run, step_through=step_through)

    if dry_run:
        click_back(driver)
        return "dry_run"

    # submit_result: True = success, False = failed (popup or btn missing),
    # "skipped" = step_through skip (treat as success — user explicitly chose).
    if submit_result is False:
        return "error: Submit triggered SAP popup — invoice likely still in Drafts"

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
                    // offsetParent check removed — unreliable in headless
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
            // offsetParent check removed — unreliable in headless
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
                // offsetParent check removed — unreliable in headless
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
                // offsetParent check removed — unreliable in headless
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
                    // offsetParent check removed — unreliable in headless
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

    # Skip rows whose col J is exactly "invoiced" (already done), "wait"
    # (manually marked by user as not-yet-ready), or "drafted, awaiting human
    # action" (Pause=1 row already saved as draft for human review). Lets the
    # user leave half-done or pending items on the To Do tab without bot
    # re-processing them.
    SKIP_STATUSES = {"invoiced", "wait", "drafted, awaiting human action"}
    skipped = [r for r in rows if r.tile3_status.strip().lower() in SKIP_STATUSES]
    if skipped:
        log.info("Skipping %d rows by col J status: %s",
                 len(skipped),
                 [(r.document_1, r.tile3_status.strip()) for r in skipped])
    rows = [r for r in rows if r.tile3_status.strip().lower() not in SKIP_STATUSES]

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

        # Mark in-progress (yellow) + timestamp
        from bot.google_sheets import write_invoice_status, write_invoice_timestamp
        write_invoice_timestamp(row.document_1)
        write_invoice_status(row.document_1, "tile3_in_progress", color="yellow")

        status = process_row(driver, row, dry_run=dry_run, step_through=step_through)
        log.info("Result: %s", status)

        if status == "submitted":
            results["submitted"] += 1
            note = getattr(row, '_skipped_note', '')
            if note:
                mark_tile3_done_with_note(row, note)
            else:
                mark_tile3_done(row)
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
                mark_tile3_paused(row)
                log.info("Paused item saved via draft recovery — marked in To Do")
            elif draft_status == "dry_run":
                results["skipped"] += 1
            else:
                results["errors"] += 1
                # Don't move to Status — leave in To Do with error in col H so it retries next cycle
                from bot.google_sheets import write_invoice_status
                write_invoice_status(row.document_1,
                                   f"tile3_error: {draft_status.replace('error: ', '')}")
        elif status.startswith("error:"):
            results["errors"] += 1
            from bot.google_sheets import write_invoice_status
            write_invoice_status(row.document_1,
                                 f"tile3_error: {status.replace('error: ', '')}")

        # Navigate back to Invoice Freight Documents tile for next item
        if i < len(all_to_process) - 1:
            if launchpad_url:
                driver.get(launchpad_url)
                wait_for_page_ready(driver)
            navigate_to_tile(driver)

    # Sort paused rows to top of To Do for easy human access
    # sort_paused_rows_to_top()  # disabled — no row reordering

    log.info("Tile 3 complete: %s", results)
    return results
