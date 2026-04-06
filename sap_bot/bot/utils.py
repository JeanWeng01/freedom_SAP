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
# Wait helpers (never use time.sleep)
# ---------------------------------------------------------------------------

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
        @functools.wraps(func)
        def wrapper(*args, dry_run=False, step_through=False, **kwargs):
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

            return func(*args, dry_run=False, step_through=False, **kwargs)
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
