"""SAP Fiori portal login module."""

import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from bot.utils import wait_for_element, take_screenshot

log = logging.getLogger(__name__)

# --- Selectors (will need tuning against the real login page) ---
# These are common patterns for SAP BTP / Azure AD login flows.
# Update after inspecting the actual login page.
EMAIL_INPUT = (By.CSS_SELECTOR, "input[type='email'], input[name='loginfmt'], input#j_username")
PASSWORD_INPUT = (By.CSS_SELECTOR, "input[type='password'], input[name='passwd'], input#j_password")
SUBMIT_BUTTON = (By.CSS_SELECTOR, "input[type='submit'], button[type='submit']")
# Indicator that we've reached the SAP Fiori launchpad home
HOME_INDICATOR = (By.CSS_SELECTOR, "[class*='launchpad'], [class*='shellTitle'], .sapUshellShell")


def login(driver: WebDriver, url: str, username: str, password: str) -> bool:
    """Log into the SAP Fiori portal.

    Returns True on success, False on failure.
    """
    log.info("Navigating to SAP portal: %s", url)
    driver.get(url)

    try:
        # --- Email / username step ---
        email_field = wait_for_element(driver, *EMAIL_INPUT, timeout=30)
        email_field.clear()
        email_field.send_keys(username)
        log.info("Entered username")

        submit_btn = wait_for_element(driver, *SUBMIT_BUTTON, timeout=10, clickable=True)
        submit_btn.click()

        # --- Password step ---
        pwd_field = wait_for_element(driver, *PASSWORD_INPUT, timeout=15)
        pwd_field.clear()
        pwd_field.send_keys(password)
        log.info("Entered password")

        submit_btn = wait_for_element(driver, *SUBMIT_BUTTON, timeout=10, clickable=True)
        submit_btn.click()

        # --- Handle possible "Stay signed in?" prompt (Microsoft SSO) ---
        try:
            stay_signed_in = wait_for_element(
                driver, By.CSS_SELECTOR, "input[value='No'], input[value='Yes']",
                timeout=5, clickable=True
            )
            stay_signed_in.click()
            log.info("Dismissed 'Stay signed in?' prompt")
        except TimeoutException:
            pass  # No such prompt — continue

        # --- Wait for home page ---
        wait_for_element(driver, *HOME_INDICATOR, timeout=60)
        log.info("Login successful — SAP Fiori home page loaded")
        take_screenshot(driver, "login_success")
        return True

    except TimeoutException as e:
        log.error("Login failed — timed out waiting for page element: %s", e)
        take_screenshot(driver, "login_failed")
        return False
    except Exception as e:
        log.error("Login failed with unexpected error: %s", e)
        take_screenshot(driver, "login_error")
        return False
