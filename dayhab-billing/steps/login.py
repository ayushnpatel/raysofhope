import sys
import logging
from typing import TYPE_CHECKING
# Remove sync expect, use async waits if needed
# from playwright.sync_api import expect

if TYPE_CHECKING:
    # Switch to async Page
    from playwright.async_api import Page

# Convert function to async
async def login_step(page: "Page", username: str, password: str) -> None:
    """Navigates to the login page and logs in using provided credentials (asynchronously).

    Args:
        page: The async Playwright Page object.
        username: The username for login.
        password: The password for login.
    """
    logging.debug("Navigating to the login page: https://www.njmmis.com/login.aspx")
    # Use await for goto
    await page.goto("https://www.njmmis.com/login.aspx", wait_until="domcontentloaded")

    # --- IMPORTANT ---
    # Locators for username, password fields, and the login button.
    # These are based on inspecting the site (as of July 2024).
    # Verify these selectors if the script fails, as website structure can change.
    # Use your browser's developer tools (right-click -> Inspect) on the login page.
    username_selector = "#txtUserName"  # Updated based on user feedback
    password_selector = "#txtPassword"
    # The login button is an <input type="submit"> with name="btnSubmit".
    login_button_selector = 'input[name="btnSubmit"]'
    # --- --- --- ---

    logging.info(f"Attempting to log in as {username}...")
    try:
        # Use await for locator actions
        await page.locator(username_selector).fill(username)
        await page.locator(password_selector).fill(password)
        logging.debug("Credentials filled. Clicking login button...")
        await page.locator(login_button_selector).click()

        # Add a wait after click to ensure login completes (e.g., wait for URL change or a known element on the next page)
        # Example: Wait for the URL to no longer be the login page
        logging.debug("Waiting for navigation after login click...")
        await page.wait_for_url("**/default.aspx", timeout=30000) # Example: Adjust URL pattern and timeout as needed

    except Exception as e:
        logging.error(f"Error during login process: {e}")
        try:
            # Use await for screenshot
            await page.screenshot(path="login_error.png")
            logging.info("Saved screenshot to login_error.png")
        except Exception as screen_err:
            logging.error(f"Could not take error screenshot during login failure: {screen_err}")
        raise # Re-raise the exception after attempting screenshot

    logging.info(f"Login successful for {username}.") # Moved success log here after wait

    # Note: Removed the final debug log, as success is confirmed by the wait
    # logging.debug("Login submitted. Waiting for potential navigation or page change...")