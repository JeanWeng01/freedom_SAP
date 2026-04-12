"""Tile 2 — Freight Orders for Reporting.

Goal: For every item in "Events to Report", copy Planned Time into Final Time
for every stop. Runs autonomously.

Workflow per item:
1. Click into the row (NOT the Report Final Time button) to open the detail page
2. On the detail page, for EACH stop (usually 2):
   a. Read the Planned Time (e.g. "Mar 16, 2026, 11:00PM UTC-5")
   b. Click "Report Final Time" button in the Action column for that stop
   c. In the popup: paste the date/time part (WITHOUT timezone) into Final Time field
   d. Click "Report" in the popup
   e. Wait for popup to close
3. Only after ALL stops are done, click Back to return to the list
"""

import re
import logging
import time as _time
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_until_gone,
    wait_for_page_ready, take_screenshot, destructive_action, click_tile,
)

log = logging.getLogger(__name__)

TILE_NAME = "Freight Orders for Reporting"

BACK_BUTTON = (By.CSS_SELECTOR,
    "button[title='Back'], .sapMNavBack, .sapUshellShellHeadItm[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Freight Orders for Reporting tile."""
    click_tile(driver, TILE_NAME)
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile2_page_loaded")
    log.info("Tile 2 page loaded")


def go_to_events_tab(driver: WebDriver):
    """Switch to the 'Events to Report' tab (handles soft hyphens)."""
    js_click_tab = """
    var spans = document.querySelectorAll('span');
    for (var i = 0; i < spans.length; i++) {
        var clean = spans[i].textContent.replace(/\\xAD/g, '').trim();
        if (clean.indexOf('Events') !== -1 && clean.indexOf('Report') !== -1
            && clean.indexOf('Reported') === -1) {
            var clickable = spans[i].closest('[role="tab"], [class*="sapMITBFilter"], [class*="sapMITBItem"]');
            if (clickable) { clickable.click(); return 'clicked'; }
            spans[i].click();
            return 'clicked';
        }
    }
    return null;
    """
    result = driver.execute_script(js_click_tab)
    if result:
        log.info("Clicked 'Events to Report' tab")
        wait_for_page_ready(driver)
        take_screenshot(driver, "tile2_events_tab")
    else:
        log.error("Could not find 'Events to Report' tab")
        take_screenshot(driver, "tile2_events_tab_missing")


def get_row_count(driver: WebDriver) -> int:
    """Count visible rows in the Events to Report list."""
    count = driver.execute_script("""
        var rows = document.querySelectorAll('.sapMListItems .sapMLIB, .sapMListTblRow');
        return rows.length;
    """)
    return count or 0


def click_into_first_row(driver: WebDriver) -> bool:
    """Click into the first row to open its detail page.

    Must click a NON-LINK cell — clicking the freight order number link opens a new
    browser tab with the wrong info. Click cells like 'Reporting Status' (text only)
    or 'Drivers and License Plate' (often blank) instead. These trigger in-place
    navigation to the reporting detail page.
    """
    url_before = driver.current_url

    # Find the row via JS, then return its identifier so Selenium can click it
    js_find_row = """
    var rows = document.querySelectorAll(
        '[role="row"].sapMListTblRow, ' +
        '.sapMListItems .sapMLIB, ' +
        '.sapMListTblRow'
    );
    var dataRows = [];
    for (var i = 0; i < rows.length; i++) {
        if (rows[i].closest('thead, [role="rowgroup"][class*="Hdr"]')) continue;
        if (rows[i].classList.contains('sapMListTblHeader')) continue;
        if (rows[i].textContent.trim().length > 0) dataRows.push(rows[i]);
    }
    if (dataRows.length === 0) return null;
    return dataRows[0];
    """
    row_el = driver.execute_script(js_find_row)
    if not row_el:
        log.error("No data rows found")
        take_screenshot(driver, "tile2_no_rows")
        return False

    # Strategy: try several click methods on the row itself (the .sapMLIB element
    # has the press handler in SAP UI5). Use Selenium ActionChains for a real
    # mouse click, which is what SAP UI5 listens for.
    from selenium.webdriver.common.action_chains import ActionChains

    # First try: Selenium native click on the row (real mouse event)
    try:
        row_el.click()
        log.info("Clicked row with native Selenium click")
    except Exception as e:
        log.warning("Native row click failed: %s — trying ActionChains", e)
        try:
            ActionChains(driver).move_to_element(row_el).click().perform()
            log.info("Clicked row with ActionChains")
        except Exception as e2:
            log.error("ActionChains click also failed: %s", e2)
            take_screenshot(driver, "tile2_click_failed")
            return False

    # Verify navigation happened — wait up to 10 seconds for detail page indicators
    _time.sleep(2)
    wait_for_page_ready(driver)

    url_after = driver.current_url
    if url_after != url_before:
        log.info("Navigation confirmed (URL changed: %s)", url_after[-60:])
        return True

    # Check for elements that ONLY exist on the detail page:
    # - "Stop 1 - ..." stop headers
    # - "Information" / "Reporting" / "Notes" tabs (detail page tabs)
    # - "Reporting Status:" label (with colon, in detail header)
    on_detail = driver.execute_script("""
        // Look for "Stop N -" header text (specific to detail page)
        var spans = document.querySelectorAll('span, h3, h4, div');
        for (var i = 0; i < spans.length; i++) {
            var clean = spans[i].textContent.replace(/\\xAD/g, '').trim();
            if (/^Stop \\d+\\s*-/.test(clean)) return 'stop_header';
        }
        // Look for the detail page tabs: "Information", "Notes", "Drivers and License Plate"
        var tabs = document.querySelectorAll('[role="tab"]');
        for (var i = 0; i < tabs.length; i++) {
            var clean = tabs[i].textContent.replace(/\\xAD/g, '').trim();
            if (clean === 'Information' || clean === 'Contacts'
                || clean === 'Drivers and License Plate') {
                return 'detail_tabs';
            }
        }
        return null;
    """)

    if on_detail:
        log.info("Navigation confirmed (%s detected)", on_detail)
        return True

    log.error("Click did not navigate — still on events list page")
    take_screenshot(driver, "tile2_click_no_nav")
    return False


def strip_timezone(planned_time: str) -> str:
    """Strip the timezone suffix from a Planned Time string.

    'Mar 16, 2026, 11:00PM UTC-5' → 'Mar 16, 2026, 11:00PM'
    'Mar 17, 2026, 12:00 AM EST'  → 'Mar 17, 2026, 12:00 AM'
    """
    # Remove trailing timezone: UTC-5, UTC+0, EST, EDT, CST, etc.
    stripped = re.sub(r'\s+(UTC[+-]?\d+|[A-Z]{2,5})\s*$', '', planned_time.strip())
    return stripped.strip()


def get_visible_report_buttons(driver):
    """Return Selenium WebElements for all VISIBLE Report Final Time buttons on the page.

    SAP Fiori is a SPA — old views remain in DOM but hidden. We filter to only
    visible (offsetParent !== null and clientWidth > 0) elements.
    """
    return driver.execute_script("""
        var visible = [];
        var links = document.querySelectorAll('a, button');
        for (var i = 0; i < links.length; i++) {
            var el = links[i];
            // Must be visible: offsetParent set and have layout
            if (el.offsetParent === null) continue;
            if (el.getClientRects().length === 0) continue;
            var clean = el.textContent.replace(/\\xAD/g, '').trim();
            if (clean === 'Report Final Time') visible.push(el);
        }
        return visible;
    """)


def get_visible_planned_times(driver) -> list[str]:
    """Return planned time strings from VISIBLE elements only on the page."""
    return driver.execute_script("""
        var values = [];
        var seen = new Set();
        var nodes = document.querySelectorAll('span, td, [class*="sapMText"]');
        for (var i = 0; i < nodes.length; i++) {
            var el = nodes[i];
            // Must be visible
            if (el.offsetParent === null) continue;
            if (el.getClientRects().length === 0) continue;
            var t = el.textContent.trim();
            // Match "Mon DD, YYYY, HH:MM[AM/PM] TZ" — date with AM/PM or UTC
            if (!/[A-Z][a-z]{2} \\d{1,2}, \\d{4}/.test(t)) continue;
            if (!/(AM|PM|UTC)/i.test(t)) continue;
            if (t.length > 60) continue;
            // Must be a leaf — no child has the same kind of content
            var children = el.querySelectorAll('*');
            var isLeaf = true;
            for (var j = 0; j < children.length; j++) {
                var ct = children[j].textContent.trim();
                if (ct === t) { isLeaf = false; break; }
                if (/(AM|PM|UTC)/i.test(ct) && /\\d{4}/.test(ct) && ct.length < 60) {
                    isLeaf = false; break;
                }
            }
            if (!isLeaf) continue;
            // Avoid duplicates from the same exact value+position
            var key = t + '@' + el.getBoundingClientRect().top;
            if (seen.has(key)) continue;
            seen.add(key);
            values.push(t);
        }
        return values;
    """)


def read_stop_data(driver: WebDriver) -> dict:
    """Read VISIBLE Planned Time values and Report Final Time buttons on the detail page."""
    btns = get_visible_report_buttons(driver)
    times = get_visible_planned_times(driver)

    # On the detail page, each stop has 1 Report button and 1 Planned Time
    # The Planned Times list might also include "Final Time" empty cells — filter to ones that have content
    log.info("Detail page (visible only): %d Report buttons, %d planned times", len(btns), len(times))
    for i, t in enumerate(times):
        log.info("  Planned time %d: '%s'", i, t)

    return {'reportBtns': btns, 'plannedTimes': times, 'reportBtnCount': len(btns)}


@destructive_action("Report Final Time in popup")
def click_popup_report(driver: WebDriver):
    """Click the 'Report' button inside the popup dialog."""
    take_screenshot(driver, "tile2_popup_before_report")
    btn = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = 0; i < dialogs.length; i++) {
            var btns = dialogs[i].querySelectorAll('button');
            for (var j = 0; j < btns.length; j++) {
                var clean = btns[j].textContent.replace(/\\xAD/g, '').trim();
                if (clean === 'Report') return btns[j];
            }
        }
        return null;
    """)
    if btn:
        # Native click — same as row click and Report Final Time button
        from selenium.webdriver.common.action_chains import ActionChains
        try:
            btn.click()
            log.info("Clicked Report in popup (native)")
        except Exception:
            ActionChains(driver).move_to_element(btn).click().perform()
            log.info("Clicked Report in popup (ActionChains)")
        _time.sleep(2)
        wait_for_page_ready(driver)
        take_screenshot(driver, "tile2_popup_after_report")
        return True
    else:
        log.error("Report button not found in popup")
        take_screenshot(driver, "tile2_popup_no_report_btn")
        return False


def process_one_stop(driver: WebDriver, stop_index: int, planned_time_raw: str,
                     *, dry_run: bool = False) -> bool:
    """Click the Report Final Time button for one stop, fill the popup, click Report."""
    datetime_only = strip_timezone(planned_time_raw)
    log.info("Stop %d: Planned='%s' → Final Time='%s'", stop_index + 1, planned_time_raw, datetime_only)

    # Re-fetch visible Report buttons (page may have updated)
    visible_btns = get_visible_report_buttons(driver)
    if stop_index >= len(visible_btns):
        log.error("Stop %d index out of range (only %d visible Report buttons)",
                  stop_index + 1, len(visible_btns))
        take_screenshot(driver, f"tile2_btn_oor_stop{stop_index+1}")
        return False

    btn = visible_btns[stop_index]

    # Use Selenium native click (real mouse event, not JS) — same trick as the row click
    from selenium.webdriver.common.action_chains import ActionChains
    try:
        # Scroll into view first
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        _time.sleep(0.5)
        btn.click()
        log.info("Clicked Report Final Time for stop %d (native click)", stop_index + 1)
    except Exception as e:
        log.warning("Native click failed: %s — trying ActionChains", e)
        try:
            ActionChains(driver).move_to_element(btn).click().perform()
            log.info("Clicked Report Final Time for stop %d (ActionChains)", stop_index + 1)
        except Exception as e2:
            log.error("All click methods failed for stop %d: %s", stop_index + 1, e2)
            take_screenshot(driver, f"tile2_btn_click_failed_stop{stop_index+1}")
            return False
    _time.sleep(1)
    wait_for_page_ready(driver)
    take_screenshot(driver, f"tile2_popup_stop{stop_index+1}")

    # Find the Final Time input field in the popup
    final_input = driver.execute_script("""
        var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
        for (var i = 0; i < dialogs.length; i++) {
            if (dialogs[i].offsetParent === null) continue;  // skip hidden dialogs
            var inputs = dialogs[i].querySelectorAll('input:not([type="hidden"])');
            for (var j = 0; j < inputs.length; j++) {
                if (inputs[j].offsetParent !== null) return inputs[j];
            }
        }
        return null;
    """)

    if not final_input:
        log.error("Final Time input not found in popup for stop %d", stop_index + 1)
        take_screenshot(driver, f"tile2_no_input_stop{stop_index+1}")
        # Try to close popup
        driver.execute_script("""
            var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
            for (var i = 0; i < dialogs.length; i++) {
                var cancelBtn = dialogs[i].querySelector('button');
                if (cancelBtn) { cancelBtn.click(); break; }
            }
        """)
        return False

    # Clear and type the datetime
    final_input.click()
    _time.sleep(0.3)
    final_input.clear()
    final_input.send_keys(datetime_only)
    log.info("Entered Final Time: '%s'", datetime_only)
    _time.sleep(0.5)
    take_screenshot(driver, f"tile2_filled_stop{stop_index+1}")

    # Skip timezone — leave it as default (always the same fixed timezone)

    # Click Report in the popup (or close it in dry run)
    result = click_popup_report(driver, dry_run=dry_run)

    if result is None:
        # Dry run — popup is still open, need to close it
        log.info("Dry run — closing popup via Cancel")
        driver.execute_script("""
            var dialogs = document.querySelectorAll('[class*="sapMDialog"], [role="dialog"]');
            for (var i = dialogs.length - 1; i >= 0; i--) {
                if (dialogs[i].offsetParent === null) continue;
                var btns = dialogs[i].querySelectorAll('button');
                for (var j = 0; j < btns.length; j++) {
                    var t = btns[j].textContent.replace(/\\xAD/g, '').trim();
                    if (t === 'Cancel' || t === 'Close') { btns[j].click(); return; }
                }
                // Last resort — click any button to dismiss
                if (btns.length > 0) btns[btns.length - 1].click();
                return;
            }
        """)
        _time.sleep(1)
        wait_for_page_ready(driver)

    return True


def process_detail_page(driver: WebDriver, *, dry_run: bool = False) -> int:
    """Process all stops on the detail page. Returns number of stops processed."""
    wait_for_page_ready(driver)
    take_screenshot(driver, "tile2_detail_page")

    data = read_stop_data(driver)

    if data['reportBtnCount'] == 0:
        log.warning("No Report Final Time buttons found on detail page")
        return 0

    # We need one planned time per Report button
    # The planned times list may contain extra values (Final Time column etc.)
    # But we expect the first N planned times to correspond to the N stops
    planned_times = data['plannedTimes']
    btn_count = data['reportBtnCount']

    if len(planned_times) < btn_count:
        log.warning("Fewer planned times (%d) than Report buttons (%d)",
                     len(planned_times), btn_count)

    # Cache planned times upfront — do NOT re-scan after each stop, as the popup
    # interaction can change what's visible and corrupt the readings
    cached_times = list(planned_times)
    log.info("Cached planned times for all stops: %s", cached_times)

    stops_done = 0

    for i in range(btn_count):
        if i >= len(cached_times):
            log.error("No planned time for stop %d — skipping", i + 1)
            take_screenshot(driver, f"tile2_no_time_stop{i+1}")
            continue

        log.info("── Stop %d/%d ──", i + 1, btn_count)
        # After reporting stop 1, the Report button for stop 1 may disappear,
        # so stop 2's button shifts to index 0. Always use index 0 for the
        # next unprocessed stop.
        btn_index = 0 if i > 0 else 0  # always click the first visible Report button
        success = process_one_stop(driver, stop_index=btn_index,
                                   planned_time_raw=cached_times[i],
                                   dry_run=dry_run)
        if success:
            stops_done += 1

        if i < btn_count - 1:
            _time.sleep(1)
            wait_for_page_ready(driver)

    log.info("Detail page done: %d/%d stops processed", stops_done, btn_count)
    return stops_done


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 2 workflow."""
    navigate_to_tile(driver)
    go_to_events_tab(driver)

    wait_for_page_ready(driver)
    row_count = get_row_count(driver)

    if row_count == 0:
        log.info("No events to report — done")
        take_screenshot(driver, "tile2_no_events")
        return 0

    log.info("Found %d rows in Events to Report", row_count)
    take_screenshot(driver, "tile2_events_list")

    total_stops = 0

    for i in range(row_count):
        log.info("════ Order %d/%d ════", i + 1, row_count)

        if not click_into_first_row(driver):
            log.error("Could not click into row — stopping")
            break

        wait_for_page_ready(driver)
        stops = process_detail_page(driver, dry_run=dry_run)
        total_stops += stops

        # ALL stops done for this order — now click Back
        try:
            back_btn = wait_for_element(driver, *BACK_BUTTON, timeout=10, clickable=True)
            back_btn.click()
            log.info("Clicked Back to return to list")
            wait_for_page_ready(driver)
        except TimeoutException:
            driver.back()
            log.info("Used browser back")
            wait_for_page_ready(driver)

        # Reported item stays in list until page refresh — refresh to remove it
        # Events to Report tab stays active after refresh, no need to re-click it
        log.info("Refreshing page to clear reported item from list")
        driver.refresh()
        wait_for_page_ready(driver)

        remaining = get_row_count(driver)
        log.info("Remaining rows: %d", remaining)

        if remaining == 0:
            log.info("All events processed")
            break

        if dry_run and i == 0:
            log.info("[DRY RUN] Processed 1 order — %d more would follow", remaining)
            break

    log.info("Tile 2 complete — %d stops %s",
             total_stops, "would be reported (dry run)" if dry_run else "reported")
    return total_stops
