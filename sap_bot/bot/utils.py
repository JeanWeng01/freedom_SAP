"""Shared helpers: waits, screenshots, retry, dry-run guard, step-through."""

import os
import logging
import functools
from datetime import datetime

from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

log = logging.getLogger(__name__)

SCREENSHOTS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "screenshots")
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)


# ---------------------------------------------------------------------------
# Screenshot helpers
# ---------------------------------------------------------------------------

def take_screenshot(driver: WebDriver, label: str = "") -> str:
    """Save a timestamped screenshot and return the file path."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = label.replace(" ", "_").replace("/", "-")[:60]
    filename = f"{ts}_{safe_label}.png" if safe_label else f"{ts}.png"
    path = os.path.join(SCREENSHOTS_DIR, filename)
    driver.save_screenshot(path)
    log.info("Screenshot saved: %s", path)
    return path


# ---------------------------------------------------------------------------
# Wait helpers
# ---------------------------------------------------------------------------

def wait_for_page_ready(driver: WebDriver, timeout: int = 30):
    """Wait for SAP Fiori page to finish loading.

    Waits for:
    1. Document ready state
    2. SAP UI5 busy indicators to disappear
    3. A short settle time for async rendering
    """
    import time as _time

    # Wait for document ready
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # Wait for SAP UI5 busy indicators to disappear
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.find_elements(
                By.CSS_SELECTOR,
                ".sapUiLocalBusyIndicator, .sapMBusyDialog, "
                ".sapUiBlockLayerTabbable, .sapMBusyIndicator"
            )) == 0 or all(
                not el.is_displayed() for el in d.find_elements(
                    By.CSS_SELECTOR,
                    ".sapUiLocalBusyIndicator, .sapMBusyDialog, "
                    ".sapUiBlockLayerTabbable, .sapMBusyIndicator"
                )
            )
        )
    except TimeoutException:
        log.debug("Busy indicator wait timed out — proceeding")

    # Short settle time for Fiori async rendering
    _time.sleep(2)


def wait_for_element(driver: WebDriver, by: str, value: str,
                     timeout: int = 30, clickable: bool = False) -> WebElement:
    """Wait for an element to be present (or clickable) and return it."""
    condition = (EC.element_to_be_clickable((by, value))
                 if clickable
                 else EC.presence_of_element_located((by, value)))
    return WebDriverWait(driver, timeout).until(condition)


def wait_for_elements(driver: WebDriver, by: str, value: str,
                      timeout: int = 30) -> list[WebElement]:
    """Wait for at least one matching element and return all matches."""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((by, value))
    )


def wait_until_gone(driver: WebDriver, by: str, value: str,
                    timeout: int = 30) -> bool:
    """Wait until an element is no longer present. Returns True if gone."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((by, value))
        )
        return True
    except TimeoutException:
        return False


# ---------------------------------------------------------------------------
# Scroll to load all items (SAP lazy-loading lists)
# ---------------------------------------------------------------------------

def scroll_to_load_all(driver: WebDriver, list_container_css: str = None,
                       pause_timeout: float = 2.0, max_scrolls: int = 100):
    """Scroll to the bottom of a lazy-loading list until no new items load."""
    scroll_target = driver
    last_height = driver.execute_script("return document.body.scrollHeight")

    for i in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        try:
            WebDriverWait(driver, pause_timeout).until(
                lambda d: d.execute_script("return document.body.scrollHeight") > last_height
            )
        except TimeoutException:
            log.debug("Scroll complete after %d scrolls (no new content)", i + 1)
            break
        last_height = driver.execute_script("return document.body.scrollHeight")
    else:
        log.warning("Reached max scroll limit (%d)", max_scrolls)


# ---------------------------------------------------------------------------
# Tile navigation helper
# ---------------------------------------------------------------------------

def click_tile(driver: WebDriver, tile_name: str, timeout: int = 30):
    """Find and click a tile on the SAP Fiori Launchpad home page.

    SAP Fiori inserts soft-hyphen characters (\\xad) into tile text for
    word-wrapping, so exact text matching is unreliable. Instead, we strip
    soft hyphens from the page text and use JavaScript to find the match.
    """
    log.info("Looking for tile: '%s'", tile_name)

    # Wait for the Fiori Launchpad to fully render
    wait_for_page_ready(driver, timeout)
    take_screenshot(driver, f"home_before_tile_{tile_name.replace(' ', '_')[:20]}")

    # Use JavaScript to find spans whose text (stripped of soft hyphens) matches
    js_find = """
    var tileName = arguments[0];
    var spans = document.querySelectorAll('span');
    for (var i = 0; i < spans.length; i++) {
        var clean = spans[i].textContent.replace(/\\xAD/g, '').trim();
        if (clean === tileName) {
            return spans[i];
        }
    }
    // Fallback: partial match
    for (var i = 0; i < spans.length; i++) {
        var clean = spans[i].textContent.replace(/\\xAD/g, '').trim();
        if (clean.indexOf(tileName) !== -1 && clean.length < tileName.length + 20) {
            return spans[i];
        }
    }
    return null;
    """

    # Try with increasing wait
    element = None
    for attempt in range(3):
        element = driver.execute_script(js_find, tile_name)
        if element:
            break
        import time as _time
        _time.sleep(3)

    if element:
        try:
            element.click()
        except Exception:
            driver.execute_script("arguments[0].click();", element)
        log.info("Clicked tile '%s'", tile_name)
        return True

    # Failed — take debug screenshot and log page text
    log.error("Could not find tile '%s' on the page", tile_name)
    take_screenshot(driver, f"tile_not_found_{tile_name.replace(' ', '_')[:20]}")

    try:
        all_spans = driver.find_elements(By.CSS_SELECTOR, "span")
        tile_texts = []
        for s in all_spans:
            t = s.text.replace('\xad', '').strip()
            if len(t) > 3 and t not in tile_texts:
                tile_texts.append(t)
        log.info("Visible span texts on page: %s", tile_texts[:30])
    except Exception:
        pass

    raise TimeoutException(f"Tile '{tile_name}' not found on home page")


# ---------------------------------------------------------------------------
# Dry-run guard
# ---------------------------------------------------------------------------

def destructive_action(action_description: str):
    """Decorator that skips the wrapped function when dry_run is True.

    Usage:
        @destructive_action("Submit invoice {invoice_num}")
        def submit_invoice(driver, invoice_num, *, dry_run=False, step_through=False):
            ...

    The wrapped function MUST accept dry_run and step_through as keyword args.
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            dry_run = kwargs.pop("dry_run", False)
            step_through = kwargs.pop("step_through", False)
            desc = action_description.format(**kwargs) if kwargs else action_description

            if dry_run:
                log.info("[DRY RUN] Would: %s", desc)
                print(f"  [DRY RUN] Would: {desc}")
                return None

            if step_through:
                print(f"\n  About to: {desc}")
                response = input("  Press Enter to continue, 'skip' to skip, 'quit' to stop: ").strip().lower()
                if response == "skip":
                    log.info("[SKIPPED] %s", desc)
                    return "skipped"
                elif response == "quit":
                    log.info("[QUIT] User quit at: %s", desc)
                    raise KeyboardInterrupt("User chose to quit at step-through prompt")

            return func(*args, **kwargs)
        return wrapper
    return decorator


# ---------------------------------------------------------------------------
# Retry helper
# ---------------------------------------------------------------------------

def retry(max_attempts: int = 3, exceptions: tuple = (Exception,)):
    """Retry decorator for flaky operations."""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exc = e
                    log.warning("Attempt %d/%d failed for %s: %s",
                                attempt, max_attempts, func.__name__, e)
            log.error("All %d attempts failed for %s", max_attempts, func.__name__)
            raise last_exc
        return wrapper
    return decorator
