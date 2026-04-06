"""Tile 1 — Freight Orders for Confirmation.

Goal: Confirm all "New" freight orders. Runs autonomously.
"""

import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_elements, wait_until_gone,
    scroll_to_load_all, take_screenshot, destructive_action,
)

log = logging.getLogger(__name__)

# ── Selectors (based on screenshot analysis — may need tuning) ──────────────
# The tile link on the home page
TILE_SELECTOR = (By.XPATH,
    "//div[contains(@class,'sapUshellTile')]//span[contains(text(),'Freight Orders for Confirmation')]"
    "/ancestor::div[contains(@class,'sapUshellTile')]"
)

# Filter tabs — "New" tab in the status filter area
NEW_TAB = (By.XPATH, "//div[contains(@class,'sapMITBFilter') or contains(@class,'sapMSegBBtn')]"
           "//span[contains(text(),'New')]/..")
ALL_TAB = (By.XPATH, "//div[contains(@class,'sapMITBFilter') or contains(@class,'sapMSegBBtn')]"
           "//span[contains(text(),'All')]/..")

# Select-all checkbox in the table header
SELECT_ALL_CHECKBOX = (By.CSS_SELECTOR,
    "table thead .sapMCb, .sapUiTableSelectAllCheckBox, .sapMListSelectAll .sapMCb"
)

# Confirm button (bottom-right corner, green button in screenshot)
CONFIRM_BUTTON = (By.XPATH,
    "//button[contains(@class,'sapMBtn')]//bdi[text()='Confirm']/ancestor::button"
    " | //button[contains(@class,'sapMBtn')]//span[text()='Confirm']/ancestor::button"
)

# Row items in the freight order list
ORDER_ROWS = (By.CSS_SELECTOR,
    ".sapMListItems .sapMLIB, .sapUiTableRow, table tbody tr"
)

# Back / navigation button
BACK_BUTTON = (By.CSS_SELECTOR,
    ".sapMNavBack, .sapUshellShellHeadItm[title='Back'], button[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Freight Orders for Confirmation tile from the home page."""
    log.info("Navigating to Freight Orders for Confirmation tile")
    tile = wait_for_element(driver, *TILE_SELECTOR, timeout=30, clickable=True)
    tile.click()
    # Wait for the page to load — look for the "New" tab
    wait_for_element(driver, *NEW_TAB, timeout=30)
    log.info("Tile 1 page loaded")


def apply_new_filter(driver: WebDriver):
    """Click on the 'New' status tab and then the 'All' sub-tab."""
    try:
        new_tab = wait_for_element(driver, *NEW_TAB, timeout=10, clickable=True)
        new_tab.click()
        log.info("Applied 'New' filter")
    except TimeoutException:
        log.warning("'New' tab not found — may already be filtered")

    try:
        all_tab = wait_for_element(driver, *ALL_TAB, timeout=10, clickable=True)
        all_tab.click()
        log.info("Switched to 'All' sub-tab")
    except TimeoutException:
        log.info("'All' sub-tab not found — continuing with default view")


def get_order_count(driver: WebDriver) -> int:
    """Return the number of visible order rows."""
    try:
        rows = driver.find_elements(*ORDER_ROWS)
        return len(rows)
    except NoSuchElementException:
        return 0


@destructive_action("Confirm all selected freight orders")
def click_confirm(driver: WebDriver):
    """Click the Confirm button."""
    btn = wait_for_element(driver, *CONFIRM_BUTTON, timeout=15, clickable=True)
    take_screenshot(driver, "tile1_before_confirm")
    btn.click()
    log.info("Clicked Confirm button")
    # Wait for confirmation to process (button should disappear or page refreshes)
    try:
        wait_until_gone(driver, *CONFIRM_BUTTON, timeout=30)
    except TimeoutException:
        pass
    take_screenshot(driver, "tile1_after_confirm")


def run(driver: WebDriver, *, dry_run: bool = False, **_kwargs):
    """Execute the full Tile 1 workflow.

    Returns the total number of orders confirmed (or that would be confirmed in dry-run).
    """
    navigate_to_tile(driver)
    apply_new_filter(driver)

    total_confirmed = 0
    iteration = 0
    max_iterations = 50  # safety limit

    while iteration < max_iterations:
        iteration += 1
        log.info("Tile 1 — confirmation pass %d", iteration)

        # Scroll to load all items
        scroll_to_load_all(driver)

        count = get_order_count(driver)
        if count == 0:
            log.info("No 'New' freight orders to confirm — done")
            break

        log.info("Found %d freight orders to confirm", count)

        # Select all
        try:
            select_all = wait_for_element(driver, *SELECT_ALL_CHECKBOX, timeout=10, clickable=True)
            select_all.click()
            log.info("Selected all orders")
        except TimeoutException:
            log.error("Could not find 'Select All' checkbox")
            take_screenshot(driver, "tile1_select_all_missing")
            break

        # Click Confirm
        result = click_confirm(driver, dry_run=dry_run)
        total_confirmed += count

        if dry_run:
            log.info("[DRY RUN] Would have confirmed %d orders", count)
            break  # In dry-run, don't loop — we've logged what would happen

        # After confirmation, check if more items remain
        try:
            apply_new_filter(driver)
        except Exception:
            pass  # Page may have auto-refreshed

    log.info("Tile 1 complete — %d orders %s",
             total_confirmed, "would be confirmed (dry run)" if dry_run else "confirmed")
    return total_confirmed
