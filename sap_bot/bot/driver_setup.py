"""WebDriver initialization — Edge or Chrome, switchable via config."""

import os
import logging
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService

log = logging.getLogger(__name__)

HEADLESS = os.environ.get("HEADLESS", "false").lower() == "true"

# Path to chromedriver baked into the Docker image. Falls back to Selenium Manager
# (auto-download) when this file isn't present — e.g. local Windows dev.
CHROMEDRIVER_PATH = "/usr/local/bin/chromedriver"


def create_driver(browser: str) -> webdriver.Remote:
    """Create and return a Selenium WebDriver instance.

    Args:
        browser: "edge" or "chrome" (case-insensitive).

    Headless mode is controlled by the HEADLESS env var (set to "true" on Railway).
    """
    browser = browser.strip().lower()

    if browser == "edge":
        options = webdriver.EdgeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        if HEADLESS:
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--window-size=1920,1080")
        driver = webdriver.Edge(service=EdgeService(), options=options)
        log.info("Edge WebDriver initialized (headless=%s)", HEADLESS)

    elif browser == "chrome":
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        if HEADLESS:
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
        if os.path.isfile(CHROMEDRIVER_PATH):
            service = ChromeService(executable_path=CHROMEDRIVER_PATH)
        else:
            service = ChromeService()
        driver = webdriver.Chrome(service=service, options=options)
        log.info("Chrome WebDriver initialized (headless=%s)", HEADLESS)

    else:
        raise ValueError(f"Unsupported browser '{browser}'. Use 'edge' or 'chrome'.")

    driver.implicitly_wait(5)
    return driver
