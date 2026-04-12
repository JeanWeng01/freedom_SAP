"""Tile 3 — Invoice Freight Documents.

Goal: For each row in SAP_bot.xlsx, filter for freight document(s), create invoice,
enter invoice number, add charges if applicable, and submit. Human-gated.
"""

import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from bot.utils import (
    wait_for_element, wait_for_elements, wait_until_gone,
    take_screenshot, destructive_action, click_tile,
)
from bot.excel_reader import InvoiceRow

log = logging.getLogger(__name__)

TILE_NAME = "Invoice Freight Documents"

# Filter field for Freight Document number
FREIGHT_DOC_FILTER = (By.CSS_SELECTOR,
    "input[placeholder*='Freight Document'], input[id*='FreightDocument'],"
    " input[aria-label*='Freight Document']"
)

# Expand icon on the filter (for multi-value / collective invoice)
FILTER_EXPAND_ICON = (By.CSS_SELECTOR,
    "button[id*='FreightDocument'][id*='vhi'], .sapMSFB .sapUiIcon"
)

# Popup for multi-value filter (Value Help dialog)
VALUE_HELP_DIALOG = (By.CSS_SELECTOR,
    ".sapMDialog[id*='ValueHelp'], .sapMDialog"
)
VALUE_HELP_INPUT_1 = (By.XPATH,
    "(//div[contains(@class,'sapMDialog')]//input[contains(@class,'sapMInputBaseInner')])[1]"
)
VALUE_HELP_INPUT_2 = (By.XPATH,
    "(//div[contains(@class,'sapMDialog')]//input[contains(@class,'sapMInputBaseInner')])[2]"
)
ADD_CONDITION_BUTTON = (By.XPATH,
    "//div[contains(@class,'sapMDialog')]//button[.//bdi[contains(text(),'Add')]]"
)
CONDITION_DROPDOWN = (By.XPATH,
    "//div[contains(@class,'sapMDialog')]//select | //div[contains(@class,'sapMDialog')]//div[contains(@class,'sapMSlt')]"
)
VALUE_HELP_OK = (By.XPATH,
    "//div[contains(@class,'sapMDialog')]//button[.//bdi[text()='OK'] or .//bdi[text()='Go']]"
)

# To Be Invoiced tab
TO_BE_INVOICED_TAB = (By.XPATH,
    "//*[contains(text(),'To be Invoiced')]/.."
)

# Table rows (freight documents listed after filtering)
DOC_ROWS = (By.CSS_SELECTOR,
    "table tbody tr, .sapMListItems .sapMLIB"
)
ROW_CHECKBOX = (By.CSS_SELECTOR,
    ".sapMCb, input[type='checkbox']"
)

# Create Invoice / Create Collective Invoice buttons
CREATE_INVOICE_BTN = (By.XPATH,
    "//button[.//bdi[text()='Create Invoice']]"
)
CREATE_COLLECTIVE_INVOICE_BTN = (By.XPATH,
    "//button[.//bdi[text()='Create Collective Invoice']]"
)

# Invoice Details tab on the invoice page
INVOICE_DETAILS_TAB = (By.XPATH,
    "//*[contains(text(),'Invoice Details')]/.."
)

# Invoice number input field
INVOICE_INPUT = (By.CSS_SELECTOR,
    "input[id*='Invoice'], input[aria-label*='Invoice'],"
    " input[id*='invoice']"
)

# Charges tab
CHARGES_TAB = (By.XPATH,
    "//*[contains(text(),'Charges')]/.."
)

# Add button in charges section
CHARGES_ADD_BTN = (By.XPATH,
    "//button[.//bdi[text()='Add']]"
)

# Charge type dropdown option
CHARGE_OPTION = (By.XPATH,
    "//li[contains(text(),'Charge')] | //div[contains(@class,'sapMSLI')]//span[text()='Charge']/.."
)

# Waiting Charges selection in popup
WAITING_CHARGES_OPTION = (By.XPATH,
    "//*[contains(text(),'Waiting Charges')]"
)

# Rate Amount input (in the newly added charge row)
RATE_AMOUNT_INPUT = (By.CSS_SELECTOR,
    "input[id*='Rate'], input[id*='rate'], input[id*='Amount'], input[id*='amount']"
)

# Submit button
SUBMIT_BUTTON = (By.XPATH,
    "//button[.//bdi[text()='Submit']]"
)

# Drafts indicator — if we land on drafts instead of invoice page
DRAFTS_INDICATOR = (By.XPATH,
    "//*[contains(text(),'Draft')] | //*[contains(text(),'draft')]"
)

BACK_BUTTON = (By.CSS_SELECTOR,
    ".sapMNavBack, button[title='Back'], .sapUshellShellHeadItm[title='Back']"
)


def navigate_to_tile(driver: WebDriver):
    """Click the Invoice Freight Documents tile."""
    click_tile(driver, TILE_NAME)
    wait_for_element(driver, *TO_BE_INVOICED_TAB, timeout=30)
    log.info("Tile 3 page loaded")


def filter_single_document(driver: WebDriver, doc_number: str):
    """Filter by a single freight document number."""
    filter_input = wait_for_element(driver, *FREIGHT_DOC_FILTER, timeout=15)
    filter_input.clear()
    filter_input.send_keys(doc_number)
    filter_input.send_keys(Keys.RETURN)
    log.info("Filtered for document: %s", doc_number)
    # Wait for results to load
    try:
        wait_for_elements(driver, *DOC_ROWS, timeout=15)
    except TimeoutException:
        log.warning("No results after filtering for %s", doc_number)


def filter_collective_documents(driver: WebDriver, doc1: str, doc2: str):
    """Filter for two freight documents using the expand/multi-value popup."""
    # Click the expand icon on the filter
    try:
        expand = wait_for_element(driver, *FILTER_EXPAND_ICON, timeout=10, clickable=True)
        expand.click()
        log.info("Opened filter value help dialog")
    except TimeoutException:
        log.error("Could not open filter expand popup")
        take_screenshot(driver, "filter_expand_failed")
        return

    # Wait for the dialog
    wait_for_element(driver, *VALUE_HELP_DIALOG, timeout=10)

    # Set first condition: "contains" + doc1
    try:
        input1 = wait_for_element(driver, *VALUE_HELP_INPUT_1, timeout=10)
        input1.clear()
        input1.send_keys(doc1)
        log.info("Entered first document: %s", doc1)
    except TimeoutException:
        log.error("First input not found in value help dialog")
        take_screenshot(driver, "value_help_input1_missing")
        return

    # Click "Add Condition"
    try:
        add_btn = wait_for_element(driver, *ADD_CONDITION_BUTTON, timeout=10, clickable=True)
        add_btn.click()
        log.info("Clicked Add Condition")
    except TimeoutException:
        log.error("Add Condition button not found")
        take_screenshot(driver, "add_condition_missing")

    # Set second condition: "contains" + doc2
    try:
        input2 = wait_for_element(driver, *VALUE_HELP_INPUT_2, timeout=10)
        input2.clear()
        input2.send_keys(doc2)
        log.info("Entered second document: %s", doc2)
    except TimeoutException:
        log.error("Second input not found")
        take_screenshot(driver, "value_help_input2_missing")

    # Click OK
    try:
        ok_btn = wait_for_element(driver, *VALUE_HELP_OK, timeout=10, clickable=True)
        ok_btn.click()
        log.info("Confirmed filter dialog")
    except TimeoutException:
        log.error("OK button not found in filter dialog")
        take_screenshot(driver, "value_help_ok_missing")

    # Wait for results
    try:
        wait_for_elements(driver, *DOC_ROWS, timeout=15)
    except TimeoutException:
        log.warning("No results after collective filter for %s + %s", doc1, doc2)


def select_all_rows(driver: WebDriver):
    """Select all visible document rows."""
    rows = driver.find_elements(*DOC_ROWS)
    for row in rows:
        try:
            cb = row.find_element(*ROW_CHECKBOX)
            if not cb.is_selected():
                cb.click()
        except NoSuchElementException:
            pass
    log.info("Selected %d document rows", len(rows))


def add_charge(driver: WebDriver, charge_type: str, amount: float):
    """Add a charge row in the Charges tab."""
    # Click Add
    add_btn = wait_for_element(driver, *CHARGES_ADD_BTN, timeout=10, clickable=True)
    add_btn.click()
    log.info("Clicked Add charge")

    # Select "Charge" from dropdown
    try:
        charge_opt = wait_for_element(driver, *CHARGE_OPTION, timeout=10, clickable=True)
        charge_opt.click()
        log.info("Selected 'Charge' option")
    except TimeoutException:
        log.warning("Charge option dropdown not found — may have auto-selected")

    # Click the category cell and select Waiting Charges
    try:
        wc_option = wait_for_element(driver, *WAITING_CHARGES_OPTION, timeout=10, clickable=True)
        wc_option.click()
        log.info("Selected '%s'", charge_type)
    except TimeoutException:
        log.error("Could not find '%s' in charge category list", charge_type)
        take_screenshot(driver, f"charge_category_missing_{charge_type}")
        return

    # Enter rate amount
    try:
        # Find the last (most recently added) rate input
        rate_inputs = driver.find_elements(*RATE_AMOUNT_INPUT)
        rate_input = rate_inputs[-1] if rate_inputs else None
        if rate_input:
            rate_input.clear()
            rate_input.send_keys(str(amount))
            log.info("Entered charge amount: %s", amount)
        else:
            log.error("Rate amount input not found")
            take_screenshot(driver, "rate_amount_missing")
    except Exception as e:
        log.error("Error entering charge amount: %s", e)
        take_screenshot(driver, "rate_amount_error")


@destructive_action("Submit invoice {invoice_num} for document(s) {doc_desc}")
def click_submit(driver: WebDriver, *, invoice_num: str = "", doc_desc: str = ""):
    """Click the Submit button on the invoice page."""
    take_screenshot(driver, f"tile3_before_submit_{invoice_num}")
    btn = wait_for_element(driver, *SUBMIT_BUTTON, timeout=15, clickable=True)
    btn.click()
    log.info("Clicked Submit for invoice %s", invoice_num)
    take_screenshot(driver, f"tile3_after_submit_{invoice_num}")


def process_row(driver: WebDriver, row: InvoiceRow,
                *, dry_run: bool = False, step_through: bool = False) -> str:
    """Process one Excel row. Returns 'submitted', 'drafted', 'skipped', or 'error'."""
    doc_desc = row.document_1
    if row.is_collective:
        doc_desc = f"{row.document_1} + {row.document_2}"

    log.info("Processing: %s → invoice %s", doc_desc, row.invoice)

    try:
        # Filter
        if row.is_collective:
            filter_collective_documents(driver, row.document_1, row.document_2)
        else:
            filter_single_document(driver, row.document_1)

        # Select rows
        select_all_rows(driver)

        # Click Create Invoice or Create Collective Invoice
        if row.is_collective:
            btn = wait_for_element(driver, *CREATE_COLLECTIVE_INVOICE_BTN, timeout=10, clickable=True)
            btn.click()
            log.info("Clicked Create Collective Invoice")
        else:
            btn = wait_for_element(driver, *CREATE_INVOICE_BTN, timeout=10, clickable=True)
            btn.click()
            log.info("Clicked Create Invoice")

        # Check if we landed on invoice page or drafts
        try:
            wait_for_element(driver, *INVOICE_DETAILS_TAB, timeout=15)
        except TimeoutException:
            # Might have gone to drafts
            log.warning("Invoice page did not open for %s — possibly went to Drafts", doc_desc)
            take_screenshot(driver, f"tile3_draft_{row.document_1}")
            # Navigate back
            try:
                driver.find_element(*BACK_BUTTON).click()
            except NoSuchElementException:
                driver.back()
            return "drafted"

        # Go to Invoice Details tab and enter invoice number
        inv_tab = wait_for_element(driver, *INVOICE_DETAILS_TAB, timeout=10, clickable=True)
        inv_tab.click()

        inv_input = wait_for_element(driver, *INVOICE_INPUT, timeout=10)
        inv_input.clear()
        inv_input.send_keys(row.invoice)
        log.info("Entered invoice number: %s", row.invoice)

        # Handle charges for leg 1
        if row.has_charges_leg1:
            charges_tab = wait_for_element(driver, *CHARGES_TAB, timeout=10, clickable=True)
            charges_tab.click()
            add_charge(driver, row.charge_type_1, row.charge_amount_1)
            log.info("Added charge for leg 1: %s = $%s", row.charge_type_1, row.charge_amount_1)

        # Handle charges for leg 2 (collective invoice)
        if row.is_collective and row.has_charges_leg2:
            add_charge(driver, row.charge_type_2, row.charge_amount_2)
            log.info("Added charge for leg 2: %s = $%s", row.charge_type_2, row.charge_amount_2)

        # Submit
        result = click_submit(
            driver,
            invoice_num=row.invoice,
            doc_desc=doc_desc,
            dry_run=dry_run,
            step_through=step_through,
        )

        if result == "skipped":
            # Navigate back for skipped items
            try:
                driver.find_element(*BACK_BUTTON).click()
            except NoSuchElementException:
                driver.back()
            return "skipped"

        return "submitted"

    except Exception as e:
        log.error("Error processing %s: %s", doc_desc, e, exc_info=True)
        take_screenshot(driver, f"tile3_error_{row.document_1}")
        # Try to navigate back to the list
        try:
            driver.find_element(*BACK_BUTTON).click()
        except Exception:
            try:
                driver.back()
            except Exception:
                pass
        return "error"


def run(driver: WebDriver, rows: list[InvoiceRow],
        *, dry_run: bool = False, step_through: bool = False, **_kwargs):
    """Execute the full Tile 3 workflow.

    Returns dict with counts: submitted, drafted, skipped, error.
    """
    navigate_to_tile(driver)

    results = {"submitted": 0, "drafted": 0, "skipped": 0, "error": 0}

    for i, row in enumerate(rows):
        log.info("── Row %d/%d (Excel row %d) ──", i + 1, len(rows), row.row_number)
        status = process_row(driver, row, dry_run=dry_run, step_through=step_through)
        results[status] += 1
        log.info("Result: %s", status)

    log.info("Tile 3 complete — %s", results)
    print(f"\nTile 3 Summary: {results['submitted']} submitted, "
          f"{results['drafted']} drafted, {results['skipped']} skipped, "
          f"{results['error']} errors")
    return results
