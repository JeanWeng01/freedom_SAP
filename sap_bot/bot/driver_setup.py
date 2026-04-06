"""WebDriver initialization — Edge or Chrome, switchable via config."""

import logging
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService

log = logging.getLogger(__name__)


def create_driver(browser: str) -> webdriver.Remote:
    """Create and return a Selenium WebDriver instance.

    Args:
        browser: "edge" or "chrome" (case-insensitive).
    """
    browser = browser.strip().lower()

    if browser == "edge":
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        driver = webdriver.Edge(service=EdgeService(), options=options)
        log.info("Edge WebDriver initialized")

    elif browser == "chrome":
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        driver = webdriver.Chrome(service=ChromeService(), options=options)
        log.info("Chrome WebDriver initialized")

    else:
        raise ValueError(f"Unsupported browser '{browser}'. Use 'edge' or 'chrome'.")

    driver.implicitly_wait(5)
    return driver
