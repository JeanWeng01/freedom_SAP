"""Tile 2 — Freight Orders for Reporting.

Goal: For every item in "Events to Report", copy Planned Time into Final Time
for every stop. Runs autonomously.
"""

import re
import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_elements, wait_until_gone,
    scroll_to_load_all, take_screenshot, destructive_action,
)

log = logging.getLogger(__name__)

# ── Timezone mapping ────────────────────────────────────────────────────────
TZ_MAP = {
    "EST": "UTC-5",
    "EDT": "UTC-4",
    "CST": "UTC-6",
    "CDT": "UTC-5",
    "MST": "UTC-7",
    "MDT": "UTC-6",
    "PST": "UTC-8",
    "PDT": "UTC-7",
    "UTC": "UTC+0",
    "UTC-5": "UTC-5",
    "UTC-4": "UTC-4",
    "UTC-6": "UTC-6",
}

# ── Selectors ───────────────────────────────────────────────────────────────
TILE_SELECTOR = (By.XPATH,
    "//div[contains(@class,'sapUshellTile')]//span[contains(text(),'Freight Orders for Reporting')]"
    "/ancestor::div[contains(@class,'sapUshellTile')]"
)

EVENTS_TO_REPORT_TAB = (By.XPATH,
    "//*[contains(@class,'sapMITBFilter') or contains(@class,'sapMSegBBtn')]"
    "//*[contains(text(),'Events to Report')]/.."
)

# Each freight order row in the "Events to Report" list
ORDER_ROWS = (By.CSS_SELECTOR,
    ".sapMListItems .sapMLIB, table tbody tr.sapUiTableRow, table tbody tr"
)

# "Report Final Time" button on each row in the list view
REPORT_FINAL_TIME_BTN_ROW = (By.XPATH,
    "//button[.//bdi[contains(text(),'Report Final Time')] or .//span[contains(text(),'Report Final Time')]]"
)

# Inside the per-order detail page: stop sections
STOP_SECTIONS = (By.XPATH,
    "//div[contains(@class,'sapUiForm') or contains(@class,'sapMPanel')]"
    "[.//span[starts-with(text(),'Stop')]]"
)

# "Report Final Time" button per stop (in the Action column of the detail page)
STOP_REPORT_BTN = (By.XPATH,
    ".//button[.//bdi[contains(text(),'Report Final Time')] or .//span[contains(text(),'Report Final Time')]]"
    " | .//a[contains(text(),'Report Final Time')]"
)

# Planned Time cell per event row in the detail page
PLANNED_TIME_CELL = (By.XPATH,
    ".//span[contains(text(),'Planned Time')]/../..//span[contains(@class,'sapMText') or contains(@class,'sapMLabel')]"
    " | .//td[preceding-sibling::td[contains(.,'Planned Time')]]"
)

# Popup elements
POPUP_FINAL_TIME_INPUT = (By.CSS_SELECTOR,
    "input[id*='FinalTime'], input[id*='finalTime'], .sapMInputBaseInner[aria-label*='Final Time']"
)
POPUP_TIMEZONE_DROPDOWN = (By.CSS_SELECTOR,
    "select[id*='TimeZone'], select[id*='timeZone'],"
    " div[id*='TimeZone'] .sapMSlt, div[id*='timeZone'] .sapMSlt"
)
POPUP_REPORT_BUTTON = (By.XPATH,
    "//footer//button[.//bdi[text()='Report']]"
    " | //div[contains(@class,'sapMDialog')]//button[.//bdi[text()='Report']]"
)

BACK_BUTTON = (By.CSS_SELECTOR,
    ".sapMNavBack, button[title='Back'], .sapUshellShellHeadItm[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Freight Orders for Reporting tile."""
    log.info("Navigating to Freight Orders for Reporting tile")
    tile = wait_for_element(driver, *TILE_SELECTOR, timeout=30, clickable=True)
    tile.click()
    wait_for_element(driver, *EVENTS_TO_REPORT_TAB, timeout=30)
    log.info("Tile 2 page loaded")


def go_to_events_tab(driver: WebDriver):
    """Switch to the 'Events to Report' tab."""
    tab = wait_for_element(driver, *EVENTS_TO_REPORT_TAB, timeout=15, clickable=True)
    tab.click()
    log.info("Switched to 'Events to Report' tab")


def parse_planned_time(raw_text: str) -> tuple[str, str] | None:
    """Parse a Planned Time string like 'Mar 17, 2026, 12:00 AM EST' or 'Mar 16, 2026, 11:00PM UTC-5'.

    Returns (datetime_str, timezone_label) or None if unparseable.
    datetime_str example: 'Mar 17, 2026, 12:00 AM'
    timezone_label example: 'EST' or 'UTC-5'
    """
    raw_text = raw_text.strip()

    # Try pattern: "Mon DD, YYYY, HH:MM AM/PM TZ"
    match = re.match(
        r'^(.+\d{4},\s*\d{1,2}:\d{2}\s*(?:AM|PM)?)\s+((?:UTC[+-]?\d+|[A-Z]{2,5}))$',
        raw_text, re.IGNORECASE
    )
    if match:
        return match.group(1).strip(), match.group(2).strip().upper()

    log.warning("Could not parse planned time: '%s'", raw_text)
    return None


def select_timezone_in_dropdown(driver: WebDriver, target_utc: str):
    """Select the correct UTC offset in the timezone dropdown."""
    # Try <select> element first
    try:
        select_el = driver.find_element(*POPUP_TIMEZONE_DROPDOWN)
        from selenium.webdriver.support.ui import Select
        sel = Select(select_el)
        for option in sel.options:
            if target_utc in option.text:
                sel.select_by_visible_text(option.text)
                log.info("Selected timezone: %s", option.text)
                return
    except (NoSuchElementException, Exception):
        pass

    # Try SAP UI5 custom dropdown (click to open, then select)
    try:
        tz_trigger = driver.find_element(By.CSS_SELECTOR,
            "div[id*='TimeZone'] .sapMSltArrow, div[id*='timeZone'] .sapMSltArrow,"
            " span[id*='TimeZone'], span[id*='timeZone']"
        )
        tz_trigger.click()
        # Find the option in the opened list
        option = wait_for_element(driver, By.XPATH,
            f"//li[contains(text(),'{target_utc}')] | //div[contains(@class,'sapMSltPicker')]//li[contains(text(),'{target_utc}')]",
            timeout=10, clickable=True
        )
        option.click()
        log.info("Selected timezone via dropdown: %s", target_utc)
    except (TimeoutException, NoSuchElementException) as e:
        log.error("Could not select timezone '%s': %s", target_utc, e)
        take_screenshot(driver, f"tz_select_failed_{target_utc}")


@destructive_action("Report Final Time for a stop")
def click_report_in_popup(driver: WebDriver):
    """Click the 'Report' button in the Final Time popup."""
    take_screenshot(driver, "tile2_before_report")
    btn = wait_for_element(driver, *POPUP_REPORT_BUTTON, timeout=10, clickable=True)
    btn.click()
    log.info("Clicked Report in popup")
    # Wait for popup to close
    wait_until_gone(driver, *POPUP_REPORT_BUTTON, timeout=15)
    take_screenshot(driver, "tile2_after_report")


def process_stop_event(driver: WebDriver, stop_element, stop_label: str,
                       *, dry_run: bool = False):
    """Process one stop's event: read Planned Time, open popup, fill, report."""
    # Read Planned Time
    try:
        planned_el = stop_element.find_element(*PLANNED_TIME_CELL)
        planned_raw = planned_el.text.strip()
    except NoSuchElementException:
        # Try broader search within the stop section
        try:
            cells = stop_element.find_elements(By.CSS_SELECTOR, "span, td")
            planned_raw = None
            for cell in cells:
                text = cell.text.strip()
                if re.search(r'\d{4}', text) and re.search(r'(?:AM|PM|UTC)', text, re.IGNORECASE):
                    planned_raw = text
                    break
            if not planned_raw:
                log.error("Could not find Planned Time for %s", stop_label)
                take_screenshot(driver, f"planned_time_missing_{stop_label}")
                return False
        except Exception as e:
            log.error("Error reading Planned Time for %s: %s", stop_label, e)
            return False

    log.info("%s — Planned Time: %s", stop_label, planned_raw)

    parsed = parse_planned_time(planned_raw)
    if parsed is None:
        log.error("Unparseable Planned Time for %s: '%s'", stop_label, planned_raw)
        take_screenshot(driver, f"unparseable_time_{stop_label}")
        return False

    datetime_str, tz_label = parsed
    tz_utc = TZ_MAP.get(tz_label.upper())
    if tz_utc is None:
        log.warning("Unknown timezone '%s' for %s — skipping", tz_label, stop_label)
        take_screenshot(driver, f"unknown_tz_{tz_label}_{stop_label}")
        return False

    # Click "Report Final Time" for this stop
    try:
        report_btn = stop_element.find_element(*STOP_REPORT_BTN)
        report_btn.click()
        log.info("Opened Report Final Time popup for %s", stop_label)
    except NoSuchElementException:
        log.error("'Report Final Time' button not found for %s", stop_label)
        take_screenshot(driver, f"report_btn_missing_{stop_label}")
        return False

    # Fill the popup
    try:
        final_time_input = wait_for_element(driver, *POPUP_FINAL_TIME_INPUT, timeout=10)
        final_time_input.clear()
        final_time_input.send_keys(datetime_str)
        log.info("Entered Final Time: %s", datetime_str)

        select_timezone_in_dropdown(driver, tz_utc)

        # Click Report
        click_report_in_popup(driver, dry_run=dry_run)
        return True

    except TimeoutException:
        log.error("Popup elements not found for %s", stop_label)
        take_screenshot(driver, f"popup_timeout_{stop_label}")
        return False


def process_order(driver: WebDriver, order_index: int,
                  *, dry_run: bool = False) -> int:
    """Process one freight order: open it, handle all stops, return stop count."""
    # Click "Report Final Time" on the row to open the detail page
    try:
        buttons = driver.find_elements(*REPORT_FINAL_TIME_BTN_ROW)
        if order_index >= len(buttons):
            log.warning("Order index %d out of range (found %d buttons)", order_index, len(buttons))
            return 0
        buttons[order_index].click()
        log.info("Opened freight order detail page (index %d)", order_index)
    except (IndexError, NoSuchElementException) as e:
        log.error("Could not open order at index %d: %s", order_index, e)
        return 0

    # Wait for detail page to load — look for stop sections
    try:
        # Find all stop sections (Stop 1, Stop 2, etc.)
        from selenium.webdriver.support.ui import WebDriverWait
        WebDriverWait(driver, 15).until(
            lambda d: len(d.find_elements(By.XPATH,
                "//span[starts-with(text(),'Stop ') or starts-with(text(),'Stop')]"
            )) > 0
        )
    except TimeoutException:
        log.error("Detail page did not load stops for order index %d", order_index)
        take_screenshot(driver, f"tile2_detail_no_stops_{order_index}")
        # Navigate back
        try:
            driver.find_element(*BACK_BUTTON).click()
        except NoSuchElementException:
            driver.back()
        return 0

    # Find stop sections — look for sections containing "Stop N"
    stop_containers = driver.find_elements(By.XPATH,
        "//div[contains(@class,'sapUiForm') or contains(@class,'sapMPanel') or contains(@class,'sapUiRGrp')]"
        "[.//span[contains(text(),'Stop ')]]"
    )

    if not stop_containers:
        # Fallback: try to find event rows directly
        stop_containers = driver.find_elements(By.XPATH,
            "//table[.//th[contains(text(),'Event') or contains(text(),'Planned')]]"
            "//tbody/tr"
        )

    stops_processed = 0
    for i, stop_el in enumerate(stop_containers):
        stop_label = f"Stop {i + 1}"
        try:
            label_el = stop_el.find_element(By.XPATH, ".//span[contains(text(),'Stop ')]")
            stop_label = label_el.text.strip()[:30]
        except NoSuchElementException:
            pass

        log.info("Processing %s", stop_label)
        success = process_stop_event(driver, stop_el, stop_label, dry_run=dry_run)
        if success:
            stops_processed += 1

    # Navigate back to list
    try:
        back_btn = wait_for_element(driver, *BACK_BUTTON, timeout=10, clickable=True)
        back_btn.click()
        log.info("Navigated back to Events to Report list")
    except TimeoutException:
        driver.back()
        log.info("Used browser back to return to list")

    return stops_processed


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 2 workflow.

    Returns total number of stops reported.
    """
    navigate_to_tile(driver)
    go_to_events_tab(driver)

    # Scroll to load all items first
    scroll_to_load_all(driver)

    total_stops = 0
    order_count = len(driver.find_elements(*REPORT_FINAL_TIME_BTN_ROW))

    if order_count == 0:
        log.info("No events to report — done")
        return 0

    log.info("Found %d freight orders with events to report", order_count)

    # Process orders one at a time (each opens a detail page and returns)
    # After going back, the list re-renders, so we always process index 0
    for i in range(order_count):
        log.info("Processing order %d/%d", i + 1, order_count)
        stops = process_order(driver, order_index=0, dry_run=dry_run)
        total_stops += stops

        if dry_run and i == 0:
            # In dry-run, process just the first order to show what would happen
            log.info("[DRY RUN] Processed 1 order as sample — %d more would follow", order_count - 1)
            break

    log.info("Tile 2 complete — %d stops %s across %d orders",
             total_stops, "would be reported (dry run)" if dry_run else "reported", order_count)
    return total_stops
