import logging
import sys
import asyncio # Added for async operations
from datetime import datetime, date, timedelta # Corrected import
from typing import TYPE_CHECKING, List, Dict, Any, Optional
from steps.parse_excel import BillingData
import string # Add import

# Switch to async API
if TYPE_CHECKING:
    from playwright.async_api import Page, BrowserContext, PlaywrightException, Locator

logger = logging.getLogger(__name__)
FIRST_ROW_PROVIDER_ID = "0924521"
NPI_NUMBER = "1700585791"
PROVIDER_AUTHORIZED_REPRESENTATIVE = "Nalin Patel"

# --- Form Filling Helper Functions ---

async def _fill_field(page: 'Page', selector: str, value: Optional[str], field_name: str, index: int, press_tab: bool = False) -> None:
    """Fills a single form field, logs the action, and optionally presses Tab."""
    if value is None or value == "":
        logger.debug(f"Tab {index}: Skipping {field_name} as value is empty.")
        return
    try:
        element = page.locator(selector)
        await element.fill(value)
        logger.debug(f"Tab {index}: Filled {field_name} ({selector}): {value}.")
        if press_tab:
            await element.press('Tab')
            await page.wait_for_timeout(100) # Small delay after tab
    except Exception as e:
        logger.warning(f"Tab {index}: Could not fill {field_name} ({selector}): {e}")
        # Decide if this should be a fatal error or just a warning
        # raise RuntimeError(f"Failed to fill {field_name} on tab {index}: {e}")


async def _fill_patient_info(page: 'Page', billing_datum: BillingData, index: int) -> None:
    """Fills the patient information fields."""
    logger.info(f"Tab {index}: Filling patient information...")
    await _fill_field(page, "#ddeclm1500_txtPatientLastName", billing_datum.patient_last_name, "Patient Last Name", index)
    await _fill_field(page, "#ddeclm1500_txtPatientFirstName", billing_datum.patient_first_name, "Patient First Name", index)

    # Patient birth date is pre-formatted by _parse_and_format_birth_date in parse_excel.py
    formatted_dob = billing_datum.patient_birth_date
    if not formatted_dob:
        logger.debug(f"Tab {index}: Patient birth date is empty or could not be parsed. Skipping DOB field.")

    await _fill_field(page, "#ddeclm1500_txtPatientDOB", formatted_dob, "Patient DOB", index)
    await _fill_field(page, "#ddeclm1500_txtInsuredIdNumber", billing_datum.insured_id, "Insured ID", index)
    await _fill_field(page, "#ddeclm1500_txtPatientAccount", billing_datum.insured_id, "Patient Account", index) # Re-using insured ID


async def _fill_prior_auth(page: 'Page', billing_datum: BillingData, index: int) -> None:
    """Fills the prior authorization number if available."""
    logger.info(f"Tab {index}: Filling Prior Authorization...")
    prior_auth = billing_datum.prior_auth_num
    await _fill_field(page, "#ddeclm1500_txtPriorAuthNumber", prior_auth, "Prior Auth Number", index)


async def _fill_diagnosis_codes(page: 'Page', billing_datum: BillingData, index: int) -> None:
    """Fills the diagnosis code fields."""
    logger.info(f"Tab {index}: Filling diagnosis codes...")
    diag_codes_str = billing_datum.diagnosis_code or ''
    # Remove punctuation and quotes
    cleaned_diag_codes_str = diag_codes_str.translate(str.maketrans('', '', string.punctuation)).replace('"', '').replace("'", "")
    diag_codes = cleaned_diag_codes_str.split()
    logger.debug(f"Tab {index}: Original diagnosis string: '{diag_codes_str}'")
    logger.debug(f"Tab {index}: Cleaned diagnosis string: '{cleaned_diag_codes_str}'")
    logger.debug(f"Tab {index}: Found diagnosis codes: {diag_codes}")

    for i, code in enumerate(diag_codes):
        selector = f"#ddeclm1500_txtRelatedInjury{i + 1}"
        field_name = f"Diagnosis Code {i + 1}"
        # Press Tab only if it's not the last code
        press_tab = i < len(diag_codes) - 1
        await _fill_field(page, selector, code, field_name, index, press_tab=press_tab)


async def _fill_service_dates(page: 'Page', from_date_str: str, to_date_str: str, index: int) -> None:
    """Fills the service date fields."""
    logger.info(f"Tab {index}: Filling service dates...")
    await _fill_field(page, "#ddeclm1500_txtFrom01", from_date_str, "From Date", index, press_tab=True)
    await _fill_field(page, "#ddeclm1500_txtTo01", to_date_str, "To Date", index, press_tab=True)


async def _fill_service_details(page: 'Page', index: int) -> None:
    """Fills Place of Service, Provider ID, and NPI numbers."""
    logger.info(f"Tab {index}: Filling service details (POS, Provider ID, NPI)...")
    await _fill_field(page, "#ddeclm1500_txtPlaceOfService01", '99', "Place of Service", index)
    # Note: The original code pressed tab after POS, but the selector logic didn't reflect that. Adding press_tab=True based on observation.
    # await _fill_field(page, "#ddeclm1500_txtPlaceOfService01", '99', "Place of Service", index, press_tab=True) # Consider if tab is needed here

    await _fill_field(page, "#ddeclm1500_txtProviderId1", FIRST_ROW_PROVIDER_ID, "First Row Provider ID", index)
    await _fill_field(page, "#ddeclm1500_txtNPI01", NPI_NUMBER, "NPI Number (Box 24J)", index)
    await _fill_field(page, "#ddeclm1500_txtProvidera", NPI_NUMBER, "Provider NPI (Box 33a)", index) # Original had two fills for NPI


async def _fill_procedures_and_modifiers(page: 'Page', billing_datum: BillingData, index: int) -> None:
    """Fills the CPT/HCPCS code and associated modifiers."""
    logger.info(f"Tab {index}: Filling procedures and modifiers...")
    procedures_str = billing_datum.procedures_services or ''
    proc_parts = procedures_str.split()
    logger.debug(f"Tab {index}: Found procedure parts: {proc_parts}")

    if not proc_parts:
        logger.warning(f"Tab {index}: No procedure/service codes found in billing data ('{procedures_str}'). Skipping CPT/Modifier fields.")
        return

    # Fill CPT/HCPCS code (first part)
    cpt_code = proc_parts[0]
    cpt_selector = "#ddeclm1500_txtCptHcpcs01"
    await _fill_field(page, cpt_selector, cpt_code, "CPT/HCPCS Code", index, press_tab=True)

    # Fill Modifiers (subsequent parts, max 4: A, B, C, D)
    modifier_letters = ['A', 'B', 'C', 'D']
    for i, modifier_code in enumerate(proc_parts[1:]):
        if i >= len(modifier_letters): # Stop if we have more modifiers than fields (A-D)
            logger.warning(f"Tab {index}: More than 4 modifiers found in '{procedures_str}'. Only filling first 4.")
            break

        modifier_letter = modifier_letters[i]
        modifier_selector = f"#ddeclm1500_txtModifier01{modifier_letter}"
        field_name = f"Modifier {modifier_letter}"
        # Press Tab only if it's not the last modifier being filled
        press_tab = i < len(proc_parts[1:]) - 1 and i < len(modifier_letters) - 1
        await _fill_field(page, modifier_selector, modifier_code, field_name, index, press_tab=press_tab)

async def _fill_diagnosis_pointer(page: 'Page', index: int) -> None:
    """Fills the diagnosis pointer field."""
    logger.info(f"Tab {index}: Filling diagnosis pointer...")
    await _fill_field(page, "#ddeclm1500_txtDiag01a", 'a', "Diagnosis Pointer", index)


async def _fill_charges_and_units(page: 'Page', billing_datum: BillingData, index: int) -> None:
    """Fills the total charges and units fields."""
    logger.info(f"Tab {index}: Filling charges and units...")
    # Use Optional[float] for dollars and Optional[int] for units to handle potential None
    charge_value = f"{billing_datum.dollars:.2f}" if billing_datum.dollars is not None else None
    units_value = str(billing_datum.units) if billing_datum.units is not None else None

    await _fill_field(page, "#ddeclm1500_txtCharges01", charge_value, "Charges (Box 24F)", index)
    await _fill_field(page, "#ddeclm1500_txtTotalCharge", charge_value, "Total Charges (Box 28)", index)
    await _fill_field(page, "#ddeclm1500_txtDaysOrUnits01", units_value, "Units (Box 24G)", index)


async def _fill_provider_signature(page: 'Page', today_str: str, index: int) -> None:
    """Fills the provider signature and date fields."""
    logger.info(f"Tab {index}: Filling provider signature fields...")
    await _fill_field(page, "#ddeclm1500_txtPhysicianSignature3", PROVIDER_AUTHORIZED_REPRESENTATIVE, "Provider Authorized Representative", index)
    await _fill_field(page, "#ddeclm1500_txtPhysicianDate3", today_str, "Signature Date", index)


# --- Main Navigation and Orchestration ---

# New async helper function to open and navigate a single tab
async def _open_and_navigate_cms_tab(context: 'BrowserContext', target_url: str, index: int, total: int, billing_datum: BillingData) -> 'Page':
    """
    Opens a new page, navigates, waits for load, and then fills the CMS-1500 form fields using helper functions.

    Args:
        context: The Playwright BrowserContext to use.
        target_url: The URL to navigate the new page to.
        index: The index of this tab (for logging purposes, 1-based).
        total: The total number of tabs being opened (for logging).
        billing_datum: The billing data object for this tab.

    Returns:
        The newly opened, loaded, and filled Playwright Page object.

    Raises:
        RuntimeError: If opening, loading, or filling the form fails significantly.
    """
    logger.info(f"Processing CMS-1500 tab {index}/{total} for {billing_datum.patient_first_name} {billing_datum.patient_last_name}...")
    page: 'Page' | None = None
    try:
        # --- Navigation Part ---
        page = await context.new_page()
        logger.info(f"Tab {index}: New tab created. Navigating to: {target_url}")
        await page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
        logger.debug(f"Tab {index}: Navigation initiated. Waiting for network idle...")
        await page.wait_for_load_state("networkidle", timeout=60000)
        logger.info(f"Tab {index}: CMS-1500 page loaded successfully.")

        # Set the browser tab title
        try:
            new_title = f"Patient: {billing_datum.patient_first_name} {billing_datum.patient_last_name}"
            await page.evaluate(f"document.title = '{new_title}'")
            logger.info(f"Tab {index}: Set tab title to '{new_title}'.")
        except Exception as title_err:
            logger.warning(f"Tab {index}: Could not set tab title: {title_err}") # Not critical

    except Exception as nav_err:
        logger.error(f"Failed to open or load CMS-1500 tab {index}: {nav_err}")
        if page and not page.is_closed():
             await _save_screenshot_and_close(page, f"error_loading_tab_{index}.png", index)
        raise RuntimeError(f"Failed to open or load CMS-1500 tab {index}: {nav_err}") from nav_err

    # --- Form Filling Part (using helper functions) ---
    try:
        logger.info(f"Tab {index}: Starting form filling process...")

        # Date Calculations (Needed by multiple helpers)
        today = date.today() # Use date directly
        days_since_monday = today.weekday() # Monday is 0, Sunday is 6
        last_monday = today - timedelta(days=days_since_monday + 7) # Use timedelta directly
        last_friday = last_monday + timedelta(days=4) # Use timedelta directly
        from_date_str = last_monday.strftime('%m/%d/%y')
        to_date_str = last_friday.strftime('%m/%d/%y')
        today_date_str = today.strftime('%m/%d/%y')
        logger.debug(f"Tab {index}: Calculated From Date: {from_date_str}, To Date: {to_date_str}, Today: {today_date_str}")

        # Call helper functions sequentially
        await _fill_patient_info(page, billing_datum, index)
        await _fill_prior_auth(page, billing_datum, index)
        await _fill_diagnosis_codes(page, billing_datum, index)
        await _fill_service_dates(page, from_date_str, to_date_str, index)
        await _fill_service_details(page, index)
        await _fill_procedures_and_modifiers(page, billing_datum, index)
        await _fill_diagnosis_pointer(page, index)
        await _fill_charges_and_units(page, billing_datum, index)
        await _fill_provider_signature(page, today_date_str, index)

        logger.info(f"Tab {index}: Finished filling form fields successfully.")
        return page # Return the page after successful filling

    except Exception as fill_err:
        # Catching potential errors from helper functions (though they log warnings now)
        # Or errors during orchestration
        logger.error(f"Critical error during form filling orchestration on tab {index}: {fill_err}", exc_info=True)
        if page and not page.is_closed():
             await _save_screenshot_and_close(page, f"error_filling_form_tab_{index}.png", index)
        # Re-raise as RuntimeError to signal failure for this tab
        raise RuntimeError(f"Failed during form filling orchestration on CMS-1500 tab {index}: {fill_err}") from fill_err


async def _save_screenshot_and_close(page: 'Page', filename: str, index: int) -> None:
    """Saves a screenshot and closes the page, logging errors."""
    logger.info(f"Attempting to save screenshot for tab {index} to {filename} and close tab.")
    try:
        await page.screenshot(path=filename)
        logger.info(f"Saved screenshot for tab {index} to {filename}")
    except Exception as screen_err:
        logger.error(f"Could not take error screenshot for tab {index}: {screen_err}")
    finally:
        if not page.is_closed():
            try:
                await page.close()
                logger.info(f"Closed tab {index}.")
            except Exception as close_err:
                logger.error(f"Error closing tab {index} after failure: {close_err}")


# Modified main function to be async
async def input_and_finalize_submission_forms_step(page: 'Page', billing_data: List[BillingData]) -> List['Page']:
    """
    Navigates the initial page to the DDE setup and concurrently opens CMS-1500 forms
    in new tabs for each entry in billing_data using the initial page's context.

    Args:
        page: The async Playwright Page object, assumed to be already logged in on the initial tab
              and belonging to an explicitly created BrowserContext.
        billing_data: The list of parsed billing data objects.

    Returns:
        A list of successfully opened and filled async Playwright Page objects.
        Note: If a tab fails during opening or filling, it will be logged and excluded
              from the returned list, but the process will attempt to continue for other tabs.

    Raises:
        RuntimeError: If login verification fails, DDE navigation fails, terms checkbox fails,
                      or locating the CMS-1500 link fails on the initial page.
                      Individual tab failures during concurrent processing are logged but
                      do not stop the entire process unless all tabs fail.
        ValueError: If billing_data is empty.
    """
    logger.info("Starting DDE form navigation and concurrent tab opening step...")
    if not billing_data:
        logger.error("Billing data list is empty. Cannot open any CMS-1500 tabs.")
        raise ValueError("Billing data cannot be empty for this step.")

    # --- Initial Page Setup ---
    try:
        # 1. Verify Login State
        log_off_link_selector = 'a.nav-link:has-text("Log Off")'
        await page.wait_for_selector(log_off_link_selector, state="visible", timeout=10000)
        logger.info("User is confirmed logged in (Log Off link found).")

        # 2. Navigate to DDE page
        dde_url = "https://www.njmmis.com/dde.aspx"
        logger.info(f"Navigating initial tab to DDE page: {dde_url}")
        await page.goto(dde_url, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_load_state("networkidle", timeout=60000)

        # 3. Check Terms Checkbox
        terms_checkbox_selector = "#termschkbx"
        logger.info("Checking the terms and conditions checkbox on the initial tab.")
        checkbox = page.locator(terms_checkbox_selector)
        await checkbox.wait_for(state="visible", timeout=10000) # Increased timeout slightly
        await checkbox.check()
        await page.wait_for_timeout(500) # Slightly longer delay
        if not await checkbox.is_checked():
            raise RuntimeError("Checkbox did not become checked after clicking.")
        logger.info("Terms checkbox checked successfully.")

        # 4. Locate the CMS-1500 Link
        logger.info("Locating the exact CMS-1500 link.")
        cms_1500_link = page.get_by_role("link", name="CMS-1500", exact=True)
        await cms_1500_link.wait_for(state="visible", timeout=10000) # Increased timeout slightly

    except Exception as setup_err:
        logger.error(f"Failed during initial page setup: {setup_err}", exc_info=True)
        screenshot_filename = f"error_initial_setup_{type(setup_err).__name__}.png"
        try:
            await page.screenshot(path=screenshot_filename)
            logger.info(f"Saved screenshot to {screenshot_filename}")
        except Exception as screen_err:
            logger.error(f"Could not take error screenshot during initial setup failure: {screen_err}")
        raise RuntimeError(f"Failed during initial page setup: {setup_err}") from setup_err

    # --- Concurrent Tab Processing ---
    context: 'BrowserContext' = page.context
    cms_1500_target_url = "https://www.njmmis.com/ddedetails.aspx?ddeForm=1500&type="
    logger.info(f"Target URL for CMS-1500 forms: {cms_1500_target_url}")

    tasks = []
    total_tabs = len(billing_data)
    for i, billing_datum in enumerate(billing_data):
        # Create a task for each tab processing
        task = _open_and_navigate_cms_tab(context, cms_1500_target_url, i + 1, total_tabs, billing_datum)
        tasks.append(task)

    logger.info(f"Attempting to process {total_tabs} CMS-1500 tabs concurrently...")
    # Use return_exceptions=True to get results or exceptions for each task
    results = await asyncio.gather(*tasks, return_exceptions=True)

    successful_pages: List['Page'] = []
    failed_count = 0
    for i, result in enumerate(results):
        if isinstance(result, Exception):
            failed_count += 1
            # Error is already logged within _open_and_navigate_cms_tab or its helpers
            logger.error(f"Tab {i + 1}/{total_tabs} failed processing. See previous logs for details. Error: {result}")
        elif result is not None: # Should be a Page object
            successful_pages.append(result)
            logger.info(f"Tab {i + 1}/{total_tabs} processed successfully.")
        else: # Should not happen if logic is correct, but handle defensively
             failed_count += 1
             logger.error(f"Tab {i + 1}/{total_tabs} returned None unexpectedly after processing.")


    logger.info(f"Concurrent processing finished. Successful tabs: {len(successful_pages)}, Failed tabs: {failed_count}")

    if not successful_pages and billing_data: # If all tabs failed
         logger.error("All CMS-1500 tabs failed during processing.")
         # Optionally raise an error if *no* tabs succeed
         # raise RuntimeError("All CMS-1500 tabs failed processing.")

    return successful_pages
