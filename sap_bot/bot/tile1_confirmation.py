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
    """Click the Freight Orders for Confirmation tile from the home page."""
    click_tile(driver, TILE_NAME)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile1_page_loaded")
    log.info("Tile 1 page loaded")


def apply_status_filter_new(driver: WebDriver):
    """Set the 'Freight Order Status' filter to 'New'."""
    log.info("Applying Freight Order Status = 'New' filter")

    js_find_filter = """
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
    """

    input_el = driver.execute_script(js_find_filter)

    if input_el:
        log.info("Found Freight Order Status filter input")
        input_el.click()
        _time.sleep(0.5)
        input_el.clear()
        input_el.send_keys("New")
        _time.sleep(1)
        take_screenshot(driver, "tile1_typed_new_in_filter")

        try:
            new_opt = wait_for_element(driver, By.XPATH,
                "//li[.//span[text()='New'] or text()='New']"
                " | //div[contains(@class,'sapMSLI')]//span[text()='New']/.."
                " | //ul[contains(@class,'sapMList')]//span[text()='New']/..",
                timeout=5, clickable=True
            )
            new_opt.click()
            log.info("Selected 'New' from suggestion list")
        except TimeoutException:
            input_el.send_keys(Keys.RETURN)
            log.info("Pressed Enter to apply 'New' filter")

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

    take_screenshot(driver, "tile1_filter_applied")


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
    """Read the count from 'All Freight Orders (N)' label."""
    try:
        label = driver.find_element(*ORDER_COUNT_LABEL)
        text = label.text.replace('\xad', '').strip()
        match = re.search(r'\(([0-9,]+)\)', text)
        if match:
            count = int(match.group(1).replace(',', ''))
            log.info("Order count from label: %d ('%s')", count, text)
            return count
        log.warning("Could not parse count from label: '%s'", text)
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
    """Click the select-all checkbox in the table header.

    Uses JS to find the checkbox in the column header row — the first checkbox
    that appears before any data rows.
    """
    js_select_all = """
    // Strategy 1: The header checkbox — first .sapMCb inside a column header or before list items
    var headerCb = document.querySelector(
        '.sapMListTblHeader .sapMCb, ' +
        '.sapMListSelectAll .sapMCb, ' +
        'thead .sapMCb, ' +
        '.sapMListHdr .sapMCb'
    );
    if (headerCb) return headerCb;

    // Strategy 2: Find all checkboxes, pick the first one that's NOT inside a data row
    var allCbs = document.querySelectorAll('.sapMCb');
    for (var i = 0; i < allCbs.length; i++) {
        var inDataRow = allCbs[i].closest('.sapMLIB, .sapMListItems, tbody tr');
        if (!inDataRow) {
            // This checkbox is outside data rows — likely the header
            return allCbs[i];
        }
    }

    // Strategy 3: Just get the very first checkbox on the page
    return document.querySelector('.sapMCb');
    """

    cb = driver.execute_script(js_select_all)
    if cb:
        try:
            cb.click()
            log.info("Clicked Select All checkbox (via JS)")
            return True
        except Exception as e:
            log.warning("Click failed, trying JS click: %s", e)
            driver.execute_script("arguments[0].click();", cb)
            log.info("Clicked Select All checkbox (via JS .click())")
            return True

    log.error("Could not find Select All checkbox")
    return False


@destructive_action("Confirm all selected freight orders")
def click_confirm(driver: WebDriver):
    """Click the Confirm button."""
    btn = wait_for_element(driver, *CONFIRM_BUTTON, timeout=15, clickable=True)
    take_screenshot(driver, "tile1_before_confirm")
    btn.click()
    log.info("Clicked Confirm button")
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile1_after_confirm")


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 1 workflow."""
    navigate_to_tile(driver)

    # Step 1: Apply "Freight Order Status = New" filter
    apply_status_filter_new(driver)

    # Step 2: Click "All" tab to see all filtered results
    click_all_tab(driver)

    # Read expected count
    expected = get_expected_count(driver)

    if expected == 0:
        log.info("No 'New' freight orders to confirm — done")
        take_screenshot(driver, "tile1_zero_new")
        return 0

    if expected is None:
        log.warning("Could not determine order count — proceeding cautiously")

    # Step 3: Scroll to load ALL items
    loaded = scroll_and_load_all(driver, expected)
    take_screenshot(driver, f"tile1_loaded_{loaded}_rows")

    if loaded == 0:
        log.info("No rows loaded — nothing to confirm")
        return 0

    log.info("Loaded %d rows (expected %s) — selecting all", loaded, expected)

    # Step 4: Scroll back to top before clicking select-all (header must be visible)
    driver.execute_script("window.scrollTo(0, 0);")
    _time.sleep(1)

    # Step 5: Select all
    if click_select_all(driver):
        wait_for_page_ready(driver)
        take_screenshot(driver, "tile1_all_selected")
    else:
        take_screenshot(driver, "tile1_select_all_failed")
        return 0

    # Step 6: Confirm
    click_confirm(driver, dry_run=dry_run)

    if dry_run:
        log.info("[DRY RUN] Would have confirmed %d orders (label said %s)",
                 loaded, expected)
        take_screenshot(driver, "tile1_dry_run_done")

    log.info("Tile 1 complete — %d orders %s",
             loaded, "would be confirmed (dry run)" if dry_run else "confirmed")
    return loaded
