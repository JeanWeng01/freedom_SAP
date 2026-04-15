"""Tile 1 — Freight Orders for Confirmation.

Goal: Confirm all "New" freight orders. Runs autonomously.

Filter flow:
1. Set "Freight Order Status" filter field to "New"
2. Click the "All" tab to see all filtered results
3. Scroll to load all items, select all, confirm
"""

import re
import logging
import time as _time
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_elements, wait_until_gone,
    wait_for_page_ready, take_screenshot, destructive_action, click_tile,
)

log = logging.getLogger(__name__)

TILE_NAME = "Freight Orders for Confirmation"

# ── Selectors ───────────────────────────────────────────────────────────────

# "Go" button in the filter bar
GO_BUTTON = (By.XPATH,
    "//button[.//bdi[text()='Go'] or .//span[text()='Go']]"
)

# Tabs: All | New | Confirmed | Rejected | Cancelled
ALL_TAB = (By.XPATH,
    "//div[contains(@class,'sapMITBItem') or contains(@class,'sapMITBFilter')]"
    "//span[contains(text(),'All')]/ancestor::div[contains(@class,'sapMITBItem') or contains(@class,'sapMITBFilter')][1]"
)

# The count label like "All Freight Orders (36)"
ORDER_COUNT_LABEL = (By.XPATH,
    "//span[contains(text(),'All Freight Orders')]"
)

# Confirm button (bottom-right, green)
CONFIRM_BUTTON = (By.XPATH,
    "//button//bdi[text()='Confirm']/ancestor::button"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Freight Orders for Confirmation tile and wait for it to actually load."""
    click_tile(driver, TILE_NAME)
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        WebDriverWait(driver, 30).until(
            lambda d: d.execute_script("""
                var inputs = document.querySelectorAll('input[placeholder], input[aria-label]');
                for (var i = 0; i < inputs.length; i++) {
                    if (inputs[i].offsetParent === null) continue;
                    var ph = (inputs[i].placeholder || '').toLowerCase();
                    var al = (inputs[i].getAttribute('aria-label') || '').toLowerCase();
                    if (ph.indexOf('freight') !== -1 || al.indexOf('freight') !== -1) return true;
                }
                var spans = document.querySelectorAll('span');
                for (var i = 0; i < spans.length; i++) {
                    if (spans[i].offsetParent === null) continue;
                    var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
                    if (/All Freight Orders \\(\\d/.test(t)) return true;
                    if (t === 'Freight Order Status') return true;
                }
                return false;
            """)
        )
        log.info("Tile 1 page content detected")
    except TimeoutException:
        log.warning("Tile 1 page content not detected after 30s — retrying")
        try:
            import os
            launchpad_url = os.environ.get("SAP_LAUNCHPAD_URL", "")
            if launchpad_url:
                driver.get(launchpad_url)
                wait_for_page_ready(driver)
            click_tile(driver, TILE_NAME)
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("""
                    var spans = document.querySelectorAll('span');
                    for (var i = 0; i < spans.length; i++) {
                        if (spans[i].offsetParent === null) continue;
                        var t = spans[i].textContent.replace(/\\xAD/g, '').trim();
                        if (/All Freight Orders \\(\\d/.test(t)) return true;
                        if (t === 'Freight Order Status') return true;
                    }
                    return false;
                """)
            )
            log.info("Tile 1 loaded on retry")
        except TimeoutException:
            log.error("Tile 1 failed to load even after retry")
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile1_page_loaded")
    log.info("Tile 1 page loaded")


def _find_status_filter_input(driver: WebDriver):
    """Find the Freight Order Status filter input element via JS."""
    return driver.execute_script("""
    var labels = document.querySelectorAll('label, span');
    for (var i = 0; i < labels.length; i++) {
        var clean = labels[i].textContent.replace(/\\xAD/g, '').trim();
        if (clean === 'Freight Order Status' || clean.indexOf('Freight Order Stat') !== -1) {
            var parent = labels[i].closest('.sapUiVlt, .sapMVBox, .sapUiFormElement, div[class*="filterBar"], div[class*="Filter"]');
            if (!parent) parent = labels[i].parentElement;
            for (var j = 0; j < 5 && parent; j++) {
                var input = parent.querySelector('input:not([type="hidden"])');
                if (input) return input;
                parent = parent.parentElement;
            }
        }
    }
    var inputs = document.querySelectorAll('input');
    for (var i = 0; i < inputs.length; i++) {
        var ph = inputs[i].getAttribute('placeholder') || '';
        var al = inputs[i].getAttribute('aria-label') || '';
        if (ph.indexOf('Freight Order Stat') !== -1 || al.indexOf('Freight Order Stat') !== -1) {
            return inputs[i];
        }
    }
    return null;
    """)


def apply_status_filter(driver: WebDriver, status_value: str):
    """Set the 'Freight Order Status' filter to the given value (e.g. 'New', 'Updated')."""
    log.info("Applying Freight Order Status = '%s' filter", status_value)

    # First, clear any existing filter tokens (e.g. leftover "New" from previous pass)
    driver.execute_script("""
        // Remove all token close buttons in the status filter (SAP multi-input tokens)
        var tokens = document.querySelectorAll('.sapMToken .sapMTokenIcon, .sapMMultiInput .sapMTokenIcon');
        for (var i = tokens.length - 1; i >= 0; i--) {
            if (tokens[i].offsetParent !== null) tokens[i].click();
        }
    """)
    _time.sleep(0.5)

    input_el = _find_status_filter_input(driver)

    if input_el:
        log.info("Found Freight Order Status filter input")
        input_el.click()
        _time.sleep(0.5)
        input_el.clear()
        input_el.send_keys(status_value)
        _time.sleep(1.5)
        take_screenshot(driver, f"tile1_typed_{status_value}_in_filter")

        # Try to select from suggestion list
        try:
            opt = wait_for_element(driver, By.XPATH,
                f"//li[.//span[text()='{status_value}'] or text()='{status_value}']"
                f" | //div[contains(@class,'sapMSLI')]//span[text()='{status_value}']/.."
                f" | //ul[contains(@class,'sapMList')]//span[text()='{status_value}']/..",
                timeout=5, clickable=True
            )
            opt.click()
            log.info("Selected '%s' from suggestion list", status_value)
        except TimeoutException:
            # Try clicking matching option via JS (handles soft hyphens)
            clicked = driver.execute_script("""
                var target = arguments[0];
                var items = document.querySelectorAll('li, [role="option"], [class*="sapMSLI"]');
                for (var i = 0; i < items.length; i++) {
                    if (items[i].offsetParent === null) continue;
                    var t = items[i].textContent.replace(/\\xAD/g, '').trim();
                    if (t === target) { items[i].click(); return true; }
                }
                return false;
            """, status_value)
            if clicked:
                log.info("Selected '%s' via JS", status_value)
            else:
                input_el.send_keys(Keys.RETURN)
                log.info("Pressed Enter to apply '%s' filter", status_value)

        wait_for_page_ready(driver)
    else:
        log.warning("Could not find status filter input via JS")
        take_screenshot(driver, "tile1_filter_input_not_found")

    # Press "Go" button if it exists
    try:
        go_btn = wait_for_element(driver, *GO_BUTTON, timeout=5, clickable=True)
        go_btn.click()
        log.info("Clicked 'Go' button")
        wait_for_page_ready(driver)
    except TimeoutException:
        log.info("No 'Go' button found — filter may have auto-applied")

    take_screenshot(driver, f"tile1_filter_{status_value}_applied")


def click_all_tab(driver: WebDriver):
    """Click the 'All' tab to see all filtered results."""
    try:
        all_tab = wait_for_element(driver, *ALL_TAB, timeout=10, clickable=True)
        all_tab.click()
        wait_for_page_ready(driver)
        log.info("Clicked 'All' tab")
        take_screenshot(driver, "tile1_all_tab")
    except TimeoutException:
        log.warning("'All' tab not found — may already be on it")
        take_screenshot(driver, "tile1_all_tab_missing")


def get_expected_count(driver: WebDriver) -> int | None:
    """Read the count from 'All Freight Orders (N)' label.

    Returns 0 if label shows 'All Freight Orders' without a count (no results).
    """
    try:
        label = driver.find_element(*ORDER_COUNT_LABEL)
        text = label.text.replace('\xad', '').strip()
        match = re.search(r'\(([0-9,]+)\)', text)
        if match:
            count = int(match.group(1).replace(',', ''))
            log.info("Order count from label: %d ('%s')", count, text)
            return count
        # Label says "All Freight Orders" without a number — means 0 results
        log.info("No count in label '%s' — interpreting as 0", text)
        return 0
    except NoSuchElementException:
        log.warning("Order count label not found")
    return None


def get_loaded_row_count(driver: WebDriver) -> int:
    """Count loaded data rows using JS — counts list items in the SAP table."""
    count = driver.execute_script("""
        // SAP Fiori uses sapMLIB (list item base) for table rows
        var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
        if (rows.length > 0) return rows.length;
        // Fallback: count checkboxes in the data area
        var cbs = document.querySelectorAll('table tbody .sapMCb, .sapMListItems .sapMCb');
        return cbs.length;
    """)
    return count or 0


def scroll_and_load_all(driver: WebDriver, expected_count: int | None):
    """Scroll the list container to load all items (SAP lazy-loads ~20 at a time).

    SAP Fiori lists live inside a scrollable container, not the page body.
    We need to find and scroll that container.
    """
    max_scrolls = 200
    stale_attempts = 0
    max_stale = 5  # allow multiple scroll attempts with no new rows before giving up
    last_count = get_loaded_row_count(driver)
    log.info("Initially loaded: %d rows", last_count)

    # Find the scrollable list container
    scroll_js = """
    // Try to find the SAP list's scrollable container
    var list = document.querySelector('.sapMList, .sapMListItems, .sapMTable');
    if (list) {
        var container = list.closest('.sapMScrollContainer, .sapMPage, [class*="Scroll"]');
        if (container) {
            container.scrollTop = container.scrollHeight;
            return 'container';
        }
    }
    // Fallback: scroll the page
    window.scrollTo(0, document.body.scrollHeight);
    return 'page';
    """

    for i in range(max_scrolls):
        if expected_count and last_count >= expected_count:
            log.info("All %d rows loaded", last_count)
            break

        scroll_target = driver.execute_script(scroll_js)
        _time.sleep(2)  # SAP needs time to fetch and render new rows

        new_count = get_loaded_row_count(driver)
        if new_count == last_count:
            stale_attempts += 1
            if stale_attempts >= max_stale:
                log.info("No new rows after %d scroll attempts — %d rows loaded (scrolled: %s)",
                         stale_attempts, new_count, scroll_target)
                break
            # Try scrolling the page body as well
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            _time.sleep(2)
            new_count = get_loaded_row_count(driver)
            if new_count == last_count:
                continue
        else:
            stale_attempts = 0

        log.info("Scroll %d: %d → %d rows (via %s)", i + 1, last_count, new_count, scroll_target)
        last_count = new_count

    final = get_loaded_row_count(driver)
    if expected_count and final < expected_count:
        log.warning("Only loaded %d of %d expected rows", final, expected_count)
    else:
        log.info("Loaded all %d rows", final)
    return final


def click_select_all(driver: WebDriver) -> bool:
    """Select all items using SAP UI5's internal API, then also click the visual
    checkbox to activate the Confirm button.

    Multiple strategies because SAP's checkbox is unreliable with Selenium.
    """
    from selenium.webdriver.common.action_chains import ActionChains

    # Strategy 1: Use SAP UI5 API to select all items programmatically
    sap_selected = driver.execute_script("""
        try {
            // Find the SAP UI5 Table control
            var tables = document.querySelectorAll('[class*="sapMList"], [class*="sapMTable"]');
            for (var i = 0; i < tables.length; i++) {
                if (tables[i].offsetParent === null) continue;
                if (tables[i].id && window.sap && window.sap.ui && window.sap.ui.getCore) {
                    var ctrl = sap.ui.getCore().byId(tables[i].id);
                    if (ctrl && ctrl.selectAll) {
                        ctrl.selectAll(true);
                        return 'sap_selectAll';
                    }
                    if (ctrl && ctrl.getItems) {
                        var items = ctrl.getItems();
                        for (var j = 0; j < items.length; j++) {
                            if (items[j].setSelected) items[j].setSelected(true);
                        }
                        ctrl.fireSelectionChange({selected: true, selectAll: true});
                        return 'sap_setSelected_' + items.length;
                    }
                }
            }
        } catch(e) { return 'sap_error: ' + e.message; }
        return null;
    """)
    if sap_selected:
        log.info("SAP UI5 select all: %s", sap_selected)

    # Strategy 2: Find and click the visual select-all checkbox
    # Find it by its title="Select All" attribute
    cb = driver.execute_script("""
        // Look for element with title "Select All"
        var el = document.querySelector('[title="Select All"], [aria-label="Select All"]');
        if (el) return el;
        // Fallback: header checkbox
        var headerCb = document.querySelector(
            '.sapMListTblHeader .sapMCb, .sapMListSelectAll .sapMCb'
        );
        return headerCb;
    """)

    if cb:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cb)
        _time.sleep(0.5)

        # Double click: deselect then reselect to activate Confirm button
        for click_num in range(2):
            try:
                ActionChains(driver).move_to_element(cb).click().perform()
                log.info("Select All visual click %d (ActionChains)", click_num + 1)
            except Exception:
                try:
                    cb.click()
                    log.info("Select All visual click %d (native)", click_num + 1)
                except Exception:
                    driver.execute_script("arguments[0].click();", cb)
                    log.info("Select All visual click %d (JS)", click_num + 1)
            _time.sleep(1)
    else:
        log.warning("Select All checkbox not found — relying on SAP UI5 API")

    # Verify: check if any rows are now selected
    selected_count = driver.execute_script("""
        var selected = document.querySelectorAll('.sapMLIBSelected, .sapMCbMarkChecked, [aria-selected="true"]');
        return selected.length;
    """)
    log.info("Selected items count: %d", selected_count or 0)

    return (selected_count or 0) > 0 or sap_selected is not None


@destructive_action("Confirm all selected freight orders")
def click_confirm(driver: WebDriver):
    """Click the Confirm button (in the footer toolbar)."""
    # Find Confirm button via JS — it's in the footer and may not be in viewport
    btn = driver.execute_script("""
        var btns = document.querySelectorAll('button');
        for (var i = 0; i < btns.length; i++) {
            var t = btns[i].textContent.replace(/\\xAD/g, '').trim();
            if (t === 'Confirm') return btns[i];
        }
        return null;
    """)
    if not btn:
        log.error("Confirm button not found on page")
        take_screenshot(driver, "tile1_confirm_missing")
        return

    # Scroll it into view
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
    _time.sleep(0.5)
    take_screenshot(driver, "tile1_before_confirm")

    from selenium.webdriver.common.action_chains import ActionChains
    try:
        btn.click()
        log.info("Clicked Confirm button (native)")
    except Exception:
        try:
            ActionChains(driver).move_to_element(btn).click().perform()
            log.info("Clicked Confirm button (ActionChains)")
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
            log.info("Clicked Confirm button (JS)")

    # SAP confirmation can take several minutes for large batches
    # Wait up to 3 minutes for the page to finish processing
    log.info("Waiting for confirmation to process (may take a few minutes)...")
    from selenium.webdriver.support.ui import WebDriverWait
    try:
        # First wait for busy indicator to APPEAR (confirms SAP started processing)
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("""
                var busy = document.querySelectorAll(
                    '.sapUiLocalBusyIndicator, .sapMBusyDialog, .sapUiBLy, .sapMBusyIndicator'
                );
                for (var i = 0; i < busy.length; i++) {
                    if (busy[i].offsetParent !== null) return true;
                }
                return false;
            """)
        )
        log.info("SAP processing started (busy indicator appeared)")
    except Exception:
        log.info("No busy indicator appeared — SAP may have processed instantly")

    # Then wait for it to DISAPPEAR (up to 3 minutes)
    try:
        WebDriverWait(driver, 180).until(
            lambda d: d.execute_script("""
                var busy = document.querySelectorAll(
                    '.sapUiLocalBusyIndicator, .sapMBusyDialog, .sapUiBLy, .sapMBusyIndicator'
                );
                for (var i = 0; i < busy.length; i++) {
                    if (busy[i].offsetParent !== null &&
                        busy[i].style.visibility !== 'hidden' &&
                        busy[i].style.display !== 'none') return false;
                }
                return true;
            """)
        )
        log.info("SAP processing complete (busy indicator gone)")
    except Exception:
        log.warning("Busy indicator still present after 3 minutes — proceeding anyway")

    _time.sleep(2)
    take_screenshot(driver, "tile1_after_confirm")


def confirm_filtered_orders(driver: WebDriver, status_value: str,
                            *, dry_run: bool = False) -> int:
    """Filter by status, select all, confirm. Loops until count drops to 0.

    Exits if: count reaches 0, count doesn't decrease (stuck), or max passes reached.
    """
    log.info("── Confirming '%s' orders ──", status_value)
    total = 0
    max_passes = 10
    prev_expected = None

    for pass_num in range(1, max_passes + 1):
        log.info("'%s' pass %d", status_value, pass_num)

        apply_status_filter(driver, status_value)
        click_all_tab(driver)

        expected = get_expected_count(driver)

        if expected == 0:
            log.info("No '%s' freight orders remaining — done", status_value)
            break

        if expected is None:
            log.warning("Could not determine order count — proceeding cautiously")

        # Safety: if count didn't decrease from last pass, stop (avoid infinite loop)
        if prev_expected is not None and expected is not None and expected >= prev_expected:
            log.warning("Count did not decrease (%s → %s) — stopping to avoid loop",
                        prev_expected, expected)
            break
        prev_expected = expected

        loaded = scroll_and_load_all(driver, expected)
        take_screenshot(driver, f"tile1_loaded_{loaded}_{status_value}_pass{pass_num}")

        if loaded == 0:
            log.info("No rows loaded for '%s' — done", status_value)
            break

        log.info("Loaded %d '%s' rows — selecting all", loaded, status_value)

        driver.execute_script("window.scrollTo(0, 0);")
        _time.sleep(1)

        if click_select_all(driver):
            wait_for_page_ready(driver)
            take_screenshot(driver, f"tile1_selected_{status_value}_pass{pass_num}")
        else:
            take_screenshot(driver, f"tile1_select_failed_{status_value}")
            break

        click_confirm(driver, dry_run=dry_run)
        total += loaded

        if dry_run:
            log.info("[DRY RUN] Would have confirmed %d '%s' orders", loaded, status_value)
            break

        # Refresh page for next pass
        log.info("Refreshing to check for remaining '%s' orders", status_value)
        driver.refresh()
        _time.sleep(3)
        try:
            wait_for_page_ready(driver)
        except Exception:
            _time.sleep(3)

    log.info("Total '%s' confirmed: %d %s",
             status_value, total, "(dry run)" if dry_run else "")
    return total


def clear_status_filter(driver: WebDriver):
    """Clear any existing filter tokens from the Freight Order Status field."""
    driver.execute_script("""
        // Remove all token close buttons
        var tokens = document.querySelectorAll('.sapMToken .sapMTokenIcon, .sapMMultiInput .sapMTokenIcon');
        for (var i = tokens.length - 1; i >= 0; i--) {
            if (tokens[i].offsetParent !== null) tokens[i].click();
        }
    """)
    _time.sleep(0.5)

    # Also clear the input field text
    input_el = _find_status_filter_input(driver)
    if input_el:
        input_el.click()
        input_el.clear()
        _time.sleep(0.3)

    log.info("Cleared status filter")
    wait_for_page_ready(driver)


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 1 workflow — confirm both 'New' and 'Updated' orders.

    Returns dict with per-status counts: {"new": N, "updated": M, "total": N+M}
    """
    navigate_to_tile(driver)

    counts = {}
    for status in ["New", "Updated"]:
        counts[status.lower()] = confirm_filtered_orders(driver, status, dry_run=dry_run)

        # Clear the filter before applying the next one
        log.info("Clearing filter before next status")
        clear_status_filter(driver)

    total = sum(counts.values())
    log.info("Tile 1 complete — New: %d, Updated: %d, Total: %d %s",
             counts["new"], counts["updated"], total,
             "(dry run)" if dry_run else "")
    return {"new": counts["new"], "updated": counts["updated"], "total": total}
