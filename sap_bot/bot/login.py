"""SAP BTP OAuth login module.

Login flow:
1. Try navigating directly to the Fiori Launchpad URL.
2. If session is active, the tiles page loads — done.
3. If redirected to login form, fill in credentials and submit.
4. After login, navigate to the Launchpad URL (OAuth may land on /home "Where to?" page).
"""

import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from bot.utils import wait_for_element, take_screenshot

log = logging.getLogger(__name__)

# ── Selectors for SAP BTP Identity Authentication login page ────────────────
EMAIL_INPUT = (By.CSS_SELECTOR, "input[name='j_username'], input[type='email'], input#j_username")
PASSWORD_INPUT = (By.CSS_SELECTOR, "input[name='j_password'], input[type='password'], input#j_password")
LOGIN_BUTTON = (By.CSS_SELECTOR,
    "button[type='submit'], input[type='submit'],"
    " button#logOnFormSubmit, .sapMBtnInner"
)

# Indicator that SAP Fiori Launchpad home page has loaded (tiles visible)
HOME_INDICATOR = (By.CSS_SELECTOR,
    ".sapUshellShell, [class*='launchpad'], [class*='shellTitle'],"
    " .sapUshellTile, div[class*='Tile']"
)


def _wait_for_either(driver: WebDriver, timeout: int = 15) -> str:
    """Wait until either the Launchpad loads OR the login form appears.

    Returns 'launchpad', 'login', or 'unknown'.
    """
    from selenium.webdriver.support.ui import WebDriverWait

    def check(d):
        # Check for launchpad tiles
        tiles = d.find_elements(*HOME_INDICATOR)
        if any(el.is_displayed() for el in tiles):
            return "launchpad"
        # Check for login form
        inputs = d.find_elements(*EMAIL_INPUT)
        if any(el.is_displayed() for el in inputs):
            return "login"
        return False

    try:
        return WebDriverWait(driver, timeout).until(check)
    except TimeoutException:
        return "unknown"


def login(driver: WebDriver, login_url: str, launchpad_url: str,
          username: str, password: str) -> bool:
    """Log into SAP and navigate to the Fiori Launchpad.

    1. Go to launchpad_url — if session is live, tiles load directly.
    2. If redirected to login, fill credentials.
    3. After login, go to launchpad_url again (avoid "Where to?" page).

    Returns True on success.
    """
    # Step 1: Try the launchpad directly (session may still be active)
    log.info("Navigating to Fiori Launchpad: %s", launchpad_url)
    driver.get(launchpad_url)

    result = _wait_for_either(driver, timeout=15)

    if result == "launchpad":
        log.info("Already logged in — Launchpad loaded")
        return True

    # Step 2: Not logged in — go to the OAuth login URL
    log.info("Session not active — navigating to login page")

    if result != "login":
        driver.get(login_url)

    try:
        # ── Email / username ──
        email_field = wait_for_element(driver, *EMAIL_INPUT, timeout=30)
        email_field.clear()
        email_field.send_keys(username)
        log.info("Entered username: %s", username)

        # ── Password ──
        pwd_field = wait_for_element(driver, *PASSWORD_INPUT, timeout=10)
        pwd_field.clear()
        pwd_field.send_keys(password)
        log.info("Entered password")

        # ── Submit ──
        login_btn = wait_for_element(driver, *LOGIN_BUTTON, timeout=10, clickable=True)
        login_btn.click()
        log.info("Clicked login button")

        # Wait a moment for the OAuth redirect chain
        try:
            wait_for_element(driver, *HOME_INDICATOR, timeout=15)
        except TimeoutException:
            pass  # May have landed on "Where to?" page — that's fine

    except TimeoutException as e:
        log.error("Login failed — timed out: %s", e)
        take_screenshot(driver, "login_failed")
        return False
    except Exception as e:
        log.error("Login failed: %s", e)
        take_screenshot(driver, "login_error")
        return False

    # Step 3: Navigate to the Launchpad URL (bypass "Where to?" page)
    log.info("Navigating to Launchpad URL after login")
    driver.get(launchpad_url)

    post_result = _wait_for_either(driver, timeout=30)
    if post_result == "launchpad":
        log.info("Login successful — Launchpad loaded")
        take_screenshot(driver, "login_success")
        return True

    log.error("Login completed but Launchpad did not load (got: %s)", post_result)
    take_screenshot(driver, "login_no_launchpad")
    return False
