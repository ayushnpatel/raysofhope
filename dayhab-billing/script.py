import os
import sys
import logging
import argparse
import asyncio # Added for async operations
from typing import TYPE_CHECKING

from dotenv import load_dotenv
# Switch to async playwright
from playwright.async_api import async_playwright

# Import steps (assuming they are now async or compatible)
from steps.login import login_step # Assuming login_step is also async now or refactored
from steps.parse_excel import parse_excel_step # This can remain sync
from steps.input_and_finalize_submission_forms import input_and_finalize_submission_forms_step

if TYPE_CHECKING:
    # Update type hints for async
    from playwright.async_api import Page, Browser, BrowserContext


log_level_name = os.getenv("LOG_LEVEL", "DEBUG").upper()
log_level = getattr(logging, log_level_name, logging.DEBUG)

logging.basicConfig(
    level=log_level,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

logger = logging.getLogger(__name__)

# Change main to async
async def main() -> None:
    """Loads credentials, parses Excel, and runs Playwright automation steps asynchronously."""

    parser = argparse.ArgumentParser(description="Automate Dayhab Billing.")
    parser.add_argument(
        "--excel-file",
        type=str,
        help="Path to the billing Excel file. Overrides EXCEL_FILE_PATH env var."
    )
    args = parser.parse_args()

    logger.info("Loading environment variables from .env file...")
    load_dotenv()
    username = os.getenv("NJMMIS_USERNAME")
    password = os.getenv("NJMMIS_PASSWORD")

    if not username or not password:
        logger.error(
            "Error: NJMMIS_USERNAME and NJMMIS_PASSWORD must be set in the .env file."
        )
        print("Please create a '.env' file in the same directory as the script with:")
        print("NJMMIS_USERNAME=your_username")
        print("NJMMIS_PASSWORD=your_password")
        sys.exit(1)

    # Use async_playwright context manager
    async with async_playwright() as p:
        browser: 'Browser' | None = None # Initialize browser variable
        page: 'Page' | None = None # Initialize page variable
        try:
            # --- Execute Steps --- (now using await)
            logger.info("Executing automation steps asynchronously...")


            logger.info("---- Step 0: Parse Excel (start) ----")
            # 0. Parse Excel Step (this can remain synchronous)
            billing_data = parse_excel_step(file_path_arg=args.excel_file)
            if not billing_data:
                logger.error("Exiting due to issues parsing billing data or empty file.")
                sys.exit(1)
            logger.info(f"Successfully loaded {len(billing_data)} billing records.")
            logger.info("---- Step 0: Parse Excel (end) ----")

            # Launch browser asynchronously
            logger.info("Starting Playwright asynchronously...")
            logger.info("Launching Chromium browser (headed)...")
            # Note: slow_mo is not available in async launch, use waits instead if needed
            browser = await p.chromium.launch(headless=False) # Removed slow_mo
            logger.info("Creating browser context asynchronously...")
            context: 'BrowserContext' = await browser.new_context()
            logger.info("Creating initial page from context asynchronously...")
            page = await context.new_page()

            logger.info("---- Step 1: Login (start) ----")
            # 1. Login Step - Assuming login_step is now async
            # If login_step is not async, it needs refactoring
            # For now, assuming it is:
            await login_step(page, username, password)
            logger.info("---- Step 1: Login (end) ----")

            logger.info("---- Step 2: Navigate to DDE and CMS-1500 Form (start) ----")
            # 2. Navigate to DDE and CMS-1500 Form (Await the async function)
            logger.info("Navigating to DDE page and opening CMS-1500 form tabs concurrently...")
            # Await the coroutine here
            cms_form_pages = await input_and_finalize_submission_forms_step(page, billing_data)
            logger.info(f"Successfully opened {len(cms_form_pages)} CMS-1500 form tabs.")
            logger.info("---- Step 2: Navigate to DDE and CMS-1500 Form (end) ----")

            # We might want to interact with cms_form_pages here in the future
            # Example: await process_all_forms(cms_form_pages, billing_data)

            # Removed the input() prompts as they block the event loop
            input("DDE form navigation step complete. Browser remains open. Press Enter in the terminal to continue...\n") 
            # Consider adding a sleep or other mechanism if a pause is needed for observation
            # await asyncio.sleep(10) # Example: pause for 10 seconds
            logger.info("All async steps completed successfully!")
            # --- --- --- --- --- ---

        except Exception as e:
            logger.exception("An unexpected error occurred during execution:")
            # Attempt to take a screenshot if the page object exists and is not closed
            if page and not page.is_closed():
                try:
                     await page.screenshot(path="runtime_error_screenshot.png")
                     logger.info("Saved screenshot to runtime_error_screenshot.png")
                except Exception as screen_err:
                     logger.error(f"Could not take error screenshot: {screen_err}")

        finally:
            # Keep the browser open for a moment after completion/error for inspection
            # Or close immediately if preferred.
            logger.info("Closing browser...")
            if browser:
                await browser.close()
            logger.info("Browser closed.")


if __name__ == "__main__":
    # Run the async main function using asyncio.run
    asyncio.run(main())
