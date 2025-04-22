from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException, StaleElementReferenceException, ElementClickInterceptedException
import time
import random
import os
import logging
import schedule
from datetime import datetime
import pathlib
import argparse
import json # Import json for parsing data-m

# Set up logging
# Use 'a' mode for append to keep logs across runs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("ms_rewards_automation.log", mode='a'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

class MicrosoftRewardsBot:
    def __init__(self, user_data_dir=None):
        # Define search terms - expanded list
        self.search_terms = [
            "weather forecast today", "latest news headlines", "easy recipe ideas", "popular movies to stream",
            "best coffee shops near me", "local events this weekend", "tech news today",
            "S&P 500 index", "cheap vacation destinations", "beginner fitness workout",
            "top book recommendations 2024", "DIY home improvement projects", "simple garden ideas",
            "attractions in London", "new music releases 2024", "NBA scores", "dog grooming tips",
            "basic car maintenance checks", "remote job opportunities", "free online courses",
            "ancient history facts", "recent science discoveries", "famous art exhibitions",
            "landscape photography tips", "healthy eating advice", "current fashion trends",
            "video game reviews", "what's trending on social media", "US political news",
            "global environmental issues", "learn a new language", "stock market analysis",
            "famous historical figures", "types of clouds", "world capital cities",
            "different dog breeds", "types of trees", "famous mountains",
            "major rivers of the world", "constellations in the night sky",
            "history of the internet", "how electricity works", "famous scientists",
            "world largest deserts", "types of currency", "meaning of life", # A bit philosophical!
            "best coding languages", "cybersecurity tips", "history of space travel",
            "famous bridges", "longest rivers", "highest waterfalls",
            "types of rocks", "periodic table elements", "famous battles in history",
            "renaissance art", "classical music composers", "modern architecture",
            "types of robots", "artificial intelligence explained", "blockchain technology",
            "sustainable living tips", "renewable energy sources", "ocean currents",
            "types of birds", "insect identification", "dinosaur names",
            "famous philosophers", "economic theories", "political systems",
            "types of poetry", "famous novels", "painting techniques",
            "photography composition rules", "healthy breakfast ideas", "meditation benefits",
            "different types of yoga", "weight loss tips", "muscle building exercises",
            "marathon training plan", "swimming techniques", "football rules",
            "basketball history", "tennis grand slams", "olympic sports",
            "famous explorers", "ancient civilizations", "mythological creatures",
            "types of martial arts", "board game strategy", "card game rules",
            "types of paradoxes", "famous equations", "astrophysics concepts"
        ]
        # Shuffle terms slightly for variation each run
        random.shuffle(self.search_terms)

        self.desktop_search_count = 35 # Defaulting to max possible
        self.mobile_search_count = 25 # Defaulting to max possible

        # Ensure user data directory is set up
        if user_data_dir:
            self.user_data_dir = os.path.abspath(user_data_dir)
            os.makedirs(self.user_data_dir, exist_ok=True) # Ensure directory exists
            logger.info(f"Using user data directory for persistent profile: {self.user_data_dir}")
        else:
            # Create a default hidden directory if none is provided
            default_dir = os.path.join(pathlib.Path.home(), ".ms_rewards_automation_profile")
            self.user_data_dir = os.path.abspath(default_dir)
            os.makedirs(self.user_data_dir, exist_ok=True)
            logger.info(f"No user data directory specified. Using default: {self.user_data_dir}")


        self.driver = None
        self.base_url = "https://rewards.microsoft.com/"
        self.bing_url = "https://www.bing.com/"

    def setup_driver(self):
        """Initialize and configure the Edge webdriver with persistent session using webdriver-manager"""
        try:
            # Ensure a previous driver instance is quit if it exists
            if self.driver:
                 try:
                     # Try to quit gracefully
                     self.driver.quit()
                     logger.info("Quit existing WebDriver instance before setup.")
                 except Exception as quit_err:
                     # If graceful quit fails, it might be a defunct process - just log and continue
                     logger.warning(f"Error quitting existing driver during setup: {quit_err}")
                 self.driver = None # Reset reference

            options = EdgeOptions()
            options.use_chromium = True

            # Set user data directory for persistent session
            # Ensure profile-directory is specified alongside user-data-dir
            options.add_argument(f"user-data-dir={self.user_data_dir}")
            options.add_argument("profile-directory=Default") # Use the 'Default' profile

            options.add_argument("--start-maximized")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-popup-blocking")
            # Avoid potential "Save your password" popups
            options.add_argument("--disable-features=PasswordManager")


            # Optional: Headless mode (uncomment to enable)
            # options.add_argument("--headless")
            # options.add_argument("--disable-gpu") # Needed for headless
            # options.add_argument("--no-sandbox") # Often needed in headless/docker environments
            # options.add_argument("--disable-dev-shm-usage") # Often needed in headless/docker environments

            logger.info("Initializing Edge WebDriver via webdriver-manager...")
            # Use the latest version manager can find
            service = EdgeService(EdgeChromiumDriverManager().install())

            self.driver = webdriver.Edge(service=service, options=options)
            logger.info("Edge WebDriver initialized successfully.")
            # Add a brief initial wait for browser window to settle
            time.sleep(3)
        except Exception as e:
            logger.error(f"Failed to initialize Edge WebDriver: {str(e)}")
            raise # Re-raise the exception to stop the workflow

    def quit_driver(self):
        """Quits the WebDriver instance."""
        if self.driver:
            logger.info("Quitting Edge WebDriver.")
            try:
                # Close all windows and quit the browser process
                self.driver.quit()
                logger.info("WebDriver quit successfully.")
            except Exception as e:
                logger.warning(f"Error during driver quit: {e}")
            self.driver = None # Ensure the reference is cleared

    def dismiss_banners(self):
        """Attempts to dismiss common banners like the 'Enough points to redeem' banner."""
        logger.info("Attempting to dismiss potential banners/popups.")
        # More comprehensive list of potential banner/popup close XPaths
        banner_close_button_xpaths = [
            "//promotional-item//button[contains(@aria-label, 'Close')]", # Specific promotional item
            "//div[contains(@class, 'redeem-banner')]//button[contains(@aria-label, 'Close')]", # Redeem banner
            "//button[contains(@aria-label, 'Close')]", # Generic close button by aria-label
            "//button[text()='Not now']", # Common "Not now" button
            "//button[contains(text(), 'Maybe later')]", # Common "Maybe later" button
            "//button[contains(@class, 'close-button')]", # Common close button class
            "//div[contains(@id, 'banner')]//button[contains(@class, 'close')]", # Generic banner close
            "//div[contains(@role, 'dialog')]//button[contains(@aria-label, 'Close')]", # Modal dialog close
            "//span[contains(@class, 'close-button')]", # Sometimes span is used
            "//button[contains(@class, 'glif-msft-modal-close')]" # Another close button pattern
        ]

        clicked_one = False
        # Iterate and try clicking each potential close button
        for xpath in banner_close_button_xpaths:
            try:
                # Wait briefly for the close button to be clickable
                # Use a very short wait per XPath to not block for too long
                close_button = WebDriverWait(self.driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                # Use JavaScript click for robustness against overlays
                self.driver.execute_script("arguments[0].click();", close_button)
                logger.info(f"Dismissed potential banner/popup using XPath: {xpath}.")
                clicked_one = True
                # Give the element time to disappear
                time.sleep(1)
                # After successfully clicking one, it's possible others appear or the page shifts,
                # so we'll break after the first successful click assuming the most prominent one is handled.
                break
            except TimeoutException:
                logger.debug(f"No dismissible banner found with XPath: {xpath} within timeout.")
                pass # Continue to the next XPath if timeout
            except ElementClickInterceptedException:
                 logger.debug(f"Click on banner close button intercepted with XPath: {xpath}. Element might be behind another, trying next.")
                 pass # Continue to the next XPath
            except Exception as e:
                # Log other unexpected errors but continue trying other XPaths
                logger.debug(f"Error dismissing banner with XPath {xpath}: {e}. Trying next XPath.")
                pass

        if clicked_one:
             logger.info("Attempted to dismiss banners/popups.")
        else:
             logger.debug("No dismissible banners/popups found.")
        time.sleep(1) # Small buffer after attempting dismissal


    def check_login_status(self):
        """Check if already logged in to Microsoft account by waiting for points element on rewards dashboard."""
        try:
            logger.info("Checking login status on rewards page...")
            # Always navigate to the rewards page first
            self.driver.get(self.base_url)
            # Add a brief wait for the page to start loading
            time.sleep(5)

            # Dismiss any banners that appear immediately upon loading
            self.dismiss_banners()

            # Updated XPaths based on provided HTML for points display
            points_xpaths = [
                "//mee-rewards-user-status-banner//p[contains(@class, 'pointsValue')]//span", # Confirmed structure
                "//mee-rewards-user-status-banner//mee-rewards-counter-animation/span", # Confirmed structure (might be same as above)
                "//div[contains(@class, 'points-package')]//span[contains(@class, 'points-label')]", # Common pattern (fallback)
                "//p[contains(@class, 'points')] | //span[contains(@class, 'points')]", # Broader classes (fallback)
                "//div[contains(@class, 'mee-rewards-counter')]//span[string-length(normalize-space()) > 0]", # Counter element (fallback)
                "//mee-rewards-user-status-banner//div[contains(@class, 'pointsBalance')]//span" # Specific to the user status banner (fallback)
            ]

            # Wait for visibility of *any* of these potential points elements
            # Then, wait for the element to have text and retrieve it
            # Basic validation: check if it's non-empty and looks like a number (possibly with commas)
            for xpath in points_xpaths:
                try:
                    logger.debug(f"Checking presence and visibility of points element with XPath: {xpath}")
                    points_element = WebDriverWait(self.driver, 10).until( # Wait up to 10s for visibility
                         EC.visibility_of_element_located((By.XPATH, xpath))
                    )
                    # Additionally check for text presence and validate
                    WebDriverWait(self.driver, 5).until(
                         EC.text_to_be_present_in_element((By.XPATH, xpath), "")
                    )

                    raw_text = points_element.text.strip()
                    aria_label_text = points_element.get_attribute("aria-label") # Check aria-label as seen in HTML

                    # Prioritize aria-label if it's a number, fallback to text
                    text_to_check = aria_label_text if aria_label_text and aria_label_text.replace(",", "").isdigit() else raw_text

                    # Basic validation: check if it's non-empty and looks like a number (possibly with commas)
                    # Also, exclude common non-point texts found in similar elements
                    if text_to_check and text_to_check.replace(",", "").isdigit() and len(text_to_check.replace(",", "")) > 0: # Check if it looks like a number with at least one digit
                         points = text_to_check
                         logger.info(f"Login status confirmed: Points element found and looks valid ('{text_to_check}').")
                         return True # Found a valid points element, logged in
                    else:
                         logger.debug(f"XPath {xpath} found element, but text '{raw_text}' / aria_label '{aria_label_text}' did not look like points. Trying next XPath.")
                         pass # Try next XPath

                except TimeoutException:
                    logger.debug(f"Points element not found/visible with XPath: {xpath} within timeout.")
                    pass # Try next XPath
                except Exception as e:
                    logger.warning(f"Error checking points element with XPath {xpath}: {e}. Trying next XPath.")
                    pass # Try next XPath


            logger.info("Points element not found/visible within timeout using any XPath. Assuming not logged in.")
            return False # No valid points element found after trying all XPaths

        except Exception as e:
            logger.error(f"Error during login status check workflow: {str(e)}")
            # If any general error occurs, assume not logged in or unable to verify
            return False

    def login(self):
        """Handles login process by checking status and waiting for manual login if needed."""
        logger.info("Attempting login or verifying existing session.")

        # First, check if already logged in via the persistent profile.
        # This also navigates to rewards and dismisses banners.
        if self.check_login_status():
            logger.info("Already logged in.")
            return True

        logger.info("Not logged in. Navigating to login page for manual authentication.")
        try:
            # Navigate directly to the Microsoft account login page
            self.driver.get("https://account.microsoft.com/account/")
            logger.info(f"Navigated to: {self.driver.current_url}")
            time.sleep(5) # Give it a few seconds to load/redirect

            # If it redirects quickly back to rewards, re-check status
            if self.base_url in self.driver.current_url:
                 logger.info("Redirected back to rewards page immediately. Re-checking status...")
                 if self.check_login_status(): # check_login_status will also dismiss banners
                      logger.info("Login status confirmed after quick redirect.")
                      return True
                 else:
                      logger.warning("Redirected, but login status could not be confirmed.")
                      # Continue to manual wait if verification failed
                      pass # Fall through to manual login wait
            elif "login.live.com" not in self.driver.current_url.lower() and "account.microsoft.com" not in self.driver.current_url.lower():
                 logger.warning(f"Navigated to unexpected URL after attempting account page: {self.driver.current_url}. User intervention might be required.")
                 # Still fall through to manual login wait, user might need to intervene on a different Microsoft page

        except Exception as e:
            logger.error(f"Failed to navigate to login page or check initial redirect: {e}")
            # Don't return False immediately, still offer chance for manual login if browser is open
            pass

        logger.info("--- Please log in manually in the browser window that opened --- ")
        logger.info("The script will wait up to 5 minutes for you to log in.")
        logger.info(f"Current URL: {self.driver.current_url}")

        # Wait until the URL no longer contains common login/account paths
        try:
            WebDriverWait(self.driver, 300).until( # Wait up to 5 minutes
                lambda d: "login.live.com" not in d.current_url.lower()
                          and "oauth2" not in d.current_url.lower()
                          and "account.microsoft.com" not in d.current_url.lower()
                          and d.current_url != "about:blank" # Ensure it's not a blank page
                          and d.current_url != self.base_url # Also wait if it stays on the rewards page without showing points
            )
            logger.info(f"Detected navigation away from login/account/OAuth page. Current URL: {self.driver.current_url}")
            # Add a small buffer time after redirect
            time.sleep(5)
        except TimeoutException:
            logger.error("Timeout (5 minutes): Still on login/account/OAuth page. Manual login failed or took too long.")
            # The browser is left open with the profile, user can continue manually later
            return False
        except Exception as e:
             logger.error(f"Error while waiting for navigation away from login page: {e}")
             # Treat as login failure for script purposes
             return False


        # Explicitly navigate back to the rewards page AFTER the potential login redirect
        logger.info("Navigating to rewards page to verify login status after manual attempt.")
        try:
            self.driver.get(self.base_url)
            time.sleep(7) # Wait for the rewards page to load properly
            # Dismiss any potential banners that might appear after loading
            self.dismiss_banners()
        except Exception as e:
            logger.error(f"Failed to navigate to rewards page after login attempt: {e}")
            # Treat as login failure for script purposes
            return False

        # Check login status again on the rewards page
        # check_login_status here will dismiss banners again just in case
        if self.check_login_status():
            logger.info("Manual login verified successfully on rewards page.")
            return True
        else:
            logger.error("Login verification failed: Could not confirm login status on rewards page after manual attempt.")
            # The browser is left open with the profile, user can continue manually later
            return False


    def perform_searches(self, count, mobile=False):
        """Perform Bing searches to earn points"""
        try:
            device_type = 'mobile' if mobile else 'desktop'
            logger.info(f"Starting {device_type} searches ({count} searches)...")

            # Set user agent for mobile searches
            if mobile:
                try:
                    # A recent common mobile user agent - using Android as it's common
                    mobile_ua = "Mozilla/5.0 (Linux; Android 10; SM-G975F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Mobile Safari/537.36 EdgA/45.05.4.5058"
                    self.driver.execute_cdp_cmd("Network.setUserAgentOverride", {
                        "userAgent": mobile_ua,
                        "platform": "Android"
                    })
                    logger.info(f"Set mobile user agent: {mobile_ua}")
                    # Reset window size for a more mobile-like experience (optional, user agent is key)
                    self.driver.set_window_size(375, 812) # Example iPhone size
                    logger.info("Set window size to simulate mobile.")

                except Exception as e:
                    logger.error(f"Failed to set mobile user agent or window size via CDP: {str(e)}. Proceeding without specific mobile UA/size.")
                    pass # Don't stop, just warn

            else: # Desktop
                try:
                    # Ensure window is maximized and reset user agent if it was mobile
                    self.driver.set_window_size(1920, 1080) # Or use maximize_window
                    self.driver.maximize_window()
                    # Explicitly set a common desktop UA if you want, or rely on default after reset
                    # desktop_ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0" # Example recent Edge UA
                    # self.driver.execute_cdp_cmd("Network.setUserAgentOverride", {
                    #     "userAgent": desktop_ua,
                    #     "platform": "Windows"
                    # })
                    # logger.info(f"Set desktop user agent: {desktop_ua}")
                    logger.info("Ensured desktop window size/maximization.")

                except Exception as e:
                    logger.warning(f"Failed to reset desktop settings (window size/UA): {str(e)}")
                    pass # Don't stop, just warn


            # Navigate to Bing
            self.driver.get(self.bing_url)
            time.sleep(5) # Wait for Bing page to load

            # Dismiss any banners that appear on Bing (like cookie banners, etc.)
            self.dismiss_banners()

            # Prepare search queries - ensure enough terms are available
            search_queries = self.search_terms[:min(count, len(self.search_terms))]
            if len(search_queries) < count:
                 logger.warning(f"Only {len(search_queries)} search terms available, requested {count}. Performing {len(search_queries)} searches.")


            # Find the search input field - wait for it to be clickable
            # XPaths for the Bing search box - based on common Bing HTML
            search_box_xpaths = [
                 "//textarea[@id='sb_form_q']", # Most common Bing search box
                 "//input[@id='sb_form_q']",    # Older/alternative Bing input
                 "//input[@name='q']",         # Generic search input name
                 "//textarea[@name='q']",      # Generic search textarea name
                 "//input[contains(@class, 'searchbox')]", # Common search box class
                 "//textarea[contains(@class, 'searchbox')]"
            ]

            search_box = None
            # Wait for any of the search box XPaths to be present and clickable
            successful_search_box_xpath = None
            for xpath in search_box_xpaths:
                try:
                    logger.debug(f"Attempting to find {device_type} search box with XPath: {xpath}")
                    search_box = WebDriverWait(self.driver, 15).until( # Increased wait for initial element
                        EC.element_to_be_clickable((By.XPATH, xpath))
                    )
                    logger.info(f"Found {device_type} search box using XPath: {xpath}.")
                    # Store the successful XPath for potential re-finding
                    successful_search_box_xpath = xpath
                    break # Found it, exit loop
                except TimeoutException:
                    logger.debug(f"{device_type} search box not found with XPath: {xpath} within timeout.")
                    pass # Try next XPath
                except Exception as e:
                    logger.debug(f"Error finding {device_type} search box with XPath {xpath}: {e}. Trying next XPath.")
                    pass # Try next XPath

            if not search_box:
                 logger.error(f"Could not find {device_type} search box using any XPath. Skipping searches.")
                 return False # Indicate failure

            # Perform searches in a loop
            for i in range(len(search_queries)):
                query = search_queries[i]
                try:
                    # Make queries slightly unique
                    unique_query = f"{query} {random.randint(1000, 9999)}"

                    # --- Re-find the search box for each search ---
                    # The page reloads after each search, making the previous element stale.
                    # Wait for the search box element to be present and clickable again.
                    # Use the xpath that worked initially if available, or try all again.
                    current_search_box = None
                    if successful_search_box_xpath:
                        try:
                            # Wait for the specific XPath that worked before
                             current_search_box = WebDriverWait(self.driver, 10).until( # Wait up to 10s for element after refresh
                                 EC.element_to_be_clickable((By.XPATH, successful_search_box_xpath))
                             )
                        except TimeoutException:
                            logger.warning(f"Primary search box XPath ({successful_search_box_xpath}) not found after search {i+1}. Trying other XPaths.")
                            pass # Fallback to trying all XPaths
                        except Exception as e:
                             logger.warning(f"Error re-finding primary search box XPath ({successful_search_box_xpath}) after search {i+1}: {e}. Trying other XPaths.")
                             pass

                    # If primary XPath failed, try all XPaths again
                    if not current_search_box:
                        for xpath in search_box_xpaths: # Try all XPaths again
                             try:
                                current_search_box = WebDriverWait(self.driver, 5).until( # Shorter wait per fallback XPath
                                   EC.element_to_be_clickable((By.XPATH, xpath))
                                )
                                # Update successful XPath if a different one worked this time
                                successful_search_box_xpath = xpath
                                logger.debug(f"Re-found {device_type} search box using fallback XPath: {xpath} for search {i+1}.")
                                break # Found it, exit retry loop
                             except TimeoutException:
                                 logger.debug(f"{device_type} search box not found with fallback XPath: {xpath} after search {i+1}.")
                                 pass
                             except Exception as e:
                                 logger.debug(f"Error re-finding {device_type} search box with fallback XPath {xpath} after search {i+1}: {e}. Trying next XPath.")
                                 pass

                    # If search box could not be re-found after retries, stop searching
                    if not current_search_box:
                         logger.error(f"Could not re-find {device_type} search box after search {i+1}. Cannot continue searches.")
                         break # Exit the search loop

                    # Clear the search box and send the query
                    current_search_box.clear(); time.sleep(0.5)
                    current_search_box.send_keys(unique_query); time.sleep(0.5)
                    current_search_box.send_keys(Keys.RETURN)

                    logger.info(f"Completed {device_type} search {i+1}/{len(search_queries)}: '{unique_query}'")

                    # Add a random delay between searches to simulate human behavior
                    time.sleep(random.uniform(7, 12)) # Slightly longer random delay

                    # Optional: Scroll down a bit to simulate real user behavior
                    try:
                        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.3);") # Scroll down 30%
                        time.sleep(random.uniform(1, 3)) # Short random wait after scroll
                         # Scroll back up to potentially see elements at the top on next search
                        self.driver.execute_script("window.scrollTo(0, 0);")
                        time.sleep(random.uniform(0.5, 1.5))
                    except Exception as scroll_err:
                         logger.debug(f"Scroll failed on search results page: {scroll_err}")
                         pass # Ignore scroll errors


                except StaleElementReferenceException:
                     logger.warning(f"Stale element on {device_type} search attempt {i+1}. Re-finding search box (handled at start of loop).")
                     # Stale element is handled by re-finding the element at the start of the next loop iteration.
                     # No need to decrement 'i' or retry here, just continue the loop.
                except Exception as e:
                    logger.error(f"Error during {device_type} search {i+1}: {str(e)}")
                    # If an error occurs during a search (other than stale element), try to continue with the next search.
                    # Re-finding the search box at the start of the loop should handle most recovery.
                    pass # Continue loop to try the next search query

            logger.info(f"Finished attempting {device_type} searches.")
            return True # Indicate completion of attempts
        except Exception as e:
            logger.error(f"General error during {device_type} searches workflow: {str(e)}")
            return False # Indicate failure

    def is_daily_set_item_complete(self, card_element):
        """Checks if a daily set card element visually indicates completion (green checkmark)."""
        # This is the simplified check based on the green checkmark only.
        try:
            # Check for the specific green checkmark icon within the element (confirmed in HTML)
            checkmark_icon = card_element.find_elements(By.XPATH, ".//span[contains(@class, 'mee-icon-SkypeCircleCheck')]")
            if any(el.is_displayed() for el in checkmark_icon):
                logger.debug("Item appears complete via green checkmark icon.")
                return True

            # Fallback checks for completion attributes/classes if the icon isn't found
            state = card_element.get_attribute("state")
            if state and state.lower() == "complete":
                logger.debug(f"Item appears complete via state attribute: {state}")
                return True

            completed_class_elements = card_element.find_elements(By.XPATH, ".//*[contains(@class, 'completed')]")
            if any(el.is_displayed() for el in completed_class_elements):
                 logger.debug("Item appears complete via 'completed' class.")
                 return True


            # If none of the above are found, assume not complete
            return False
        except Exception as e:
            logger.debug(f"Error checking daily set item completion status: {e}. Assuming not complete for now.")
            return False # If checking fails, assume not complete


    def get_daily_set_item_status(self, card_element):
        """Checks the status of a daily set card element: 'completed' or 'actionable'."""
        # Simplifies status to only check for completion. If not completed, it's actionable.

        try:
            # Check for completion (Green checkmark, etc.)
            if self.is_daily_set_item_complete(card_element):
                return "completed"

            # If not completed, it's actionable (ignoring any other icons/states)
            return "actionable"

        except Exception as e:
            logger.warning(f"Error checking daily set item status: {e}. Assuming 'actionable' for now.")
            return "actionable" # If checking fails, assume actionable


    def complete_daily_set(self):
        """Complete daily set activities"""
        try:
            logger.info("Starting daily set tasks...")

            # Ensure on the rewards page and dismiss banners
            if self.base_url not in self.driver.current_url:
                 logger.info("Navigating to rewards dashboard for daily set.")
                 self.driver.get(self.base_url)
                 time.sleep(5)
            self.dismiss_banners() # Dismiss banners before finding elements

            # --- Find Daily Set Container ---
            # Use the confirmed ID
            daily_sets_container_xpath = "//*[@id='daily-sets']"

            found_container = None
            try:
                logger.info(f"Attempting to find daily set container with XPath: {daily_sets_container_xpath}")
                # Wait for visibility of the container itself
                found_container = WebDriverWait(self.driver, 15).until(
                    EC.visibility_of_element_located((By.XPATH, daily_sets_container_xpath))
                )
                logger.info(f"Found daily set container.")
            except TimeoutException:
                logger.error("Could not find the daily set container. Skipping daily set.")
                return False # Indicate failure
            except Exception as e:
                 logger.error(f"Error finding daily set container: {e}. Skipping daily set.")
                 return False

            # --- Define Clickable Daily Set Cards within the container ---
            # Use XPaths based on the provided HTML, targeting the <a> tags
            # that seem to be the clickable cards.
            card_clickable_xpaths_in_container = [
                 ".//div[contains(@class, 'daily-set-item')]/a", # Specific structure confirmed
                 ".//a[contains(@class, 'ds-card-sec')]",       # Alternative targeting the specific class
                 ".//mee-card//a[contains(@href, '')]"          # Fallback: any link within a mee-card
            ]

            # List to store unique identifiers of tasks found
            task_identifiers = []
            try:
                # Wait for *presence* of at least one potential card element matching *any* of the XPaths within the container
                logger.debug("Checking presence of daily set card clickable elements within container...")
                WebDriverWait(found_container, 10).until( # Wait up to 10s for presence
                     EC.presence_of_element_located((By.XPATH, " | ".join(card_clickable_xpaths_in_container)))
                )
                logger.debug("Presence of daily set card element confirmed.")

                # Find all *visible* elements matching the card XPaths within the container
                all_candidate_cards = found_container.find_elements(By.XPATH, " | ".join(card_clickable_xpaths_in_container))
                visible_cards = [card for card in all_candidate_cards if card.is_displayed()]
                logger.info(f"Identified {len(visible_cards)} visible daily set cards within the container.")

                if not visible_cards:
                    logger.info("No visible daily set cards found initially. Daily set likely already completed or not available.")
                    return True # Consider this success if no tasks are found

                # Collect identifiers for processing loop
                for idx, card in enumerate(visible_cards):
                    # Get unique identifiers for the card
                    identifiers = {}
                    identifiers['original_index'] = idx
                    identifiers['href'] = card.get_attribute("href")
                    identifiers['data_bi_id'] = card.get_attribute("data-bi-id")
                    identifiers['data_m_attr'] = card.get_attribute("data-m")

                    # Fallback to text if no reliable identifier found
                    if not identifiers['href'] and not identifiers['data_bi_id'] and not identifiers['data_m_attr']:
                         try:
                            text_element = card.find_elements(By.XPATH, ".//h3 | .//div[contains(@class, 'card-title')] | .//*[string-length(normalize-space()) > 0]")
                            identifiers['id'] = text_element[0].text.strip() if text_element else f"Unknown_DailySet_{idx}"
                            if len(identifiers['id']) > 50: identifiers['id'] = identifiers['id'][:50] + "..."
                         except:
                            identifiers['id'] = f"Unknown_DailySet_{idx}"
                         logger.warning(f"No reliable ID (href, data-bi-id, data-m) for card {idx}, using fallback identifier '{identifiers['id']}'.")
                    else:
                         # Use a primary ID for logging if available
                         identifiers['id'] = identifiers.get('href') or identifiers.get('data_bi_id') or identifiers.get('data_m_attr')[:50] + '...' if identifiers.get('data_m_attr') else f"ID_Found_{idx}"


                    task_identifiers.append(identifiers)

                logger.info(f"Collected {len(task_identifiers)} daily set task identifiers.")


            except TimeoutException:
                logger.error("Timeout waiting for presence of daily set card clickable elements within the container. Skipping daily set.")
                return False
            except Exception as e:
                logger.error(f"Error finding initial visible daily set cards within container: {str(e)}. Skipping daily set.")
                return False


            # --- Process Found Cards using Identifiers ---
            task_statuses = {info['original_index']: 'initial' for info in task_identifiers}

            # Process tasks by their original index order
            for original_index in sorted(task_statuses.keys()):
                 # Find the task_info for this original_index
                 task_info = next((item for item in task_identifiers if item['original_index'] == original_index), None)
                 if not task_info: # Should not happen if logic is correct
                      logger.error(f"Could not find task_info for original index {original_index}. Skipping.")
                      continue

                 offer_id = task_info['id']
                 href_to_find = task_info.get('href')
                 data_bi_id_to_find = task_info.get('data_bi_id')
                 data_m_attr_to_find = task_info.get('data_m_attr')


                 logger.info(f"Processing daily set task (original index {original_index}): '{offer_id}'...")

                 # --- Attempt to process the specific card with retries ---
                 max_retries = 3
                 task_processed_successfully = False

                 for retry_count in range(max_retries):
                     card_element = None
                     current_task_status = 'unknown'

                     try:
                         # --- Navigate back to Rewards Dashboard and Re-find Elements ---
                         logger.debug(f"Navigating back to Rewards dashboard to re-find element (Retry {retry_count+1}/{max_retries})...")
                         self.driver.get(self.base_url)
                         time.sleep(7)
                         self.dismiss_banners()

                         # Re-find the container after navigating back
                         found_container = WebDriverWait(self.driver, 15).until(
                             EC.visibility_of_element_located((By.XPATH, daily_sets_container_xpath))
                         )
                         logger.debug("Re-found daily set container.")

                         # Re-find the SPECIFIC element using its identifier within the fresh container
                         logger.debug(f"Re-finding daily set card by identifiers within container...")
                         all_current_visible_cards = [card for card in found_container.find_elements(By.XPATH, " | ".join(card_clickable_xpaths_in_container)) if card.is_displayed()]

                         found_match = False
                         for current_card in all_current_visible_cards:
                             current_href = current_card.get_attribute("href")
                             current_data_bi_id = current_card.get_attribute("data-bi-id")
                             current_data_m_attr = current_card.get_attribute("data-m")

                             if (href_to_find and current_href == href_to_find) or \
                                (data_bi_id_to_find and current_data_bi_id == data_bi_id_to_find) or \
                                (data_m_attr_to_find and current_data_m_attr == data_m_attr_to_find):
                                 card_element = current_card
                                 logger.debug(f"Matched daily set card by identifier on retry {retry_count+1}.")
                                 found_match = True
                                 WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(card_element))))
                                 break

                         if not found_match and original_index < len(all_current_visible_cards):
                              card_element = all_current_visible_cards[original_index]
                              logger.debug(f"Matched daily set card by index {original_index} on retry {retry_count+1}.")
                              WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(card_element))))


                         if card_element is None:
                                   logger.warning(f"Could not re-find visible element for daily set task '{offer_id}' on retry {retry_count+1}/{max_retries}. Skipping processing for this task.")
                                   break


                         # --- If element re-found, check status and interact ---
                         if card_element:
                             current_task_status = self.get_daily_set_item_status(card_element)
                             if current_task_status == "completed":
                                 logger.info(f"Daily set task '{offer_id}' appears completed.")
                                 task_statuses[original_index] = "completed"
                                 task_processed_successfully = True
                                 break

                             # We do not explicitly check for "locked" status here anymore,
                             # as any non-completed task is considered "actionable" based on the simplified logic.

                             # If status is "actionable"
                             logger.info(f"Daily set task '{offer_id}' is actionable. Attempting interaction (Retry {retry_count+1}/{max_retries})...")

                             # Scroll to the card and click
                             try:
                                 self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card_element)
                                 time.sleep(1)
                             except Exception as scroll_err:
                                 logger.debug(f"Scroll failed for daily set card: {scroll_err}")
                                 pass

                             try:
                                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(card_element))))
                                self.driver.execute_script("arguments[0].click();", card_element)
                                logger.info(f"Clicked daily set task '{offer_id}' successfully.")
                             except Exception as click_err:
                                 logger.warning(f"JS click failed for daily set task '{offer_id}': {click_err}. Retrying.")
                                 if retry_count == max_retries - 1:
                                      logger.error(f"Max retries reached for JS click on daily set task '{offer_id}'. Skipping task.")
                                      break
                                 time.sleep(2)
                                 continue

                             # --- Handle Activity Page (New Tab or In-Page) ---
                             initial_window_handle = self.driver.current_window_handle
                             time.sleep(2)
                             window_handles_after = self.driver.window_handles

                             if len(window_handles_after) > len(self.driver.window_handles):
                                 logger.error("Window handles increased unexpectedly after click. This might indicate a problem with tab handling.")
                                 try:
                                      self.driver.switch_to.window(initial_window_handle)
                                      logger.warning("Switched back to original window. Skipping interaction on potential new tab.")
                                 except:
                                      logger.critical("Failed to switch back to original window. Cannot continue.")
                                      raise
                                 task_statuses[original_index] = "attempted"


                             elif len(window_handles_after) > 1:
                                 new_window_handle = [handle for handle in window_handles_after if handle != initial_window_handle]
                                 if new_window_handle:
                                      self.driver.switch_to.window(new_window_handle[-1])
                                      logger.info(f"Switched to new tab for daily set task '{offer_id}'")

                                      self.handle_activity_page() # Call dedicated handler

                                      logger.info("Closing daily set activity tab and switching back.")
                                      try:
                                          self.driver.close()
                                          self.driver.switch_to.window(initial_window_handle)
                                          time.sleep(3)
                                      except Exception as close_err:
                                           logger.error(f"Error closing activity tab or switching back: {close_err}. Recovery attempt.")
                                           try:
                                                self.driver.get(self.base_url)
                                                time.sleep(7)
                                                self.dismiss_banners()
                                           except:
                                                logger.critical("Failed to navigate back to rewards dashboard after tab error. Cannot reliably continue daily set.")
                                                raise
                                 else:
                                     logger.warning(f"Expected new tab after clicking daily set task '{offer_id}', but no new handle found. Assuming in-page or error.")


                             else:
                                 logger.warning(f"Clicking daily set task '{offer_id}' did not open a new tab or multiple windows already existed. Assuming in-page activity or simple link. Waiting...")
                                 time.sleep(random.uniform(10, 15))
                                 logger.info("Finished waiting after in-page interaction attempt.")

                             task_statuses[original_index] = "attempted"
                             task_processed_successfully = True
                             break

                     # --- Except blocks for retry attempts ---
                     except TimeoutException:
                          logger.warning(f"Timeout waiting for element/page on retry {retry_count+1}/{max_retries} for daily set task '{offer_id}'. Retrying.")
                          time.sleep(2)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for daily set task '{offer_id}' due to Timeout. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                     except StaleElementReferenceException:
                          logger.warning(f"Stale element reference on retry {retry_count+1}/{max_retries} for daily set task '{offer_id}'. Re-finding element and retrying.")
                          time.sleep(2)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for daily set task '{offer_id}' due to Stale Element. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                     except ElementClickInterceptedException as ice:
                          logger.warning(f"Click intercepted on retry {retry_count+1}/{max_retries} for daily set task '{offer_id}': {ice}. Retrying.")
                          time.sleep(2)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for daily set task '{offer_id}' due to Element Click Intercepted. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                     except ElementNotInteractableException as eint_err:
                         logger.warning(f"Element not interactable on retry {retry_count+1}/{max_retries} for daily set task '{offer_id}': {eint_err}. Skipping task.")
                         task_statuses[original_index] = "skipped_not_interactable"
                         break

                     except Exception as e:
                          logger.error(f"Unexpected error on retry {retry_count+1}/{max_retries} for daily set task '{offer_id}': {str(e)}. Retrying.")
                          time.sleep(3)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for daily set task '{offer_id}' due to unexpected error. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                 # --- After the retry loop finishes for a single task ---
                 if not task_processed_successfully:
                      logger.warning(f"Daily set task '{offer_id}' was not successfully processed after {max_retries} retries.")
                 else:
                      if task_statuses[original_index] == 'initial':
                           task_statuses[original_index] = 'attempted'


                 logger.info(f"Finished processing logic for daily set task '{offer_id}'. Final Status: {task_statuses[original_index]}.")


            logger.info("Finished attempting daily set tasks.")
            # Log final status of all tasks attempted
            for original_index, status in task_statuses.items():
                 offer_id = next((item['id'] for item in task_identifiers if item['original_index'] == original_index), 'N/A')
                 logger.info(f"Daily Set Task (Original Index {original_index}, ID: '{offer_id}'): {status}")

            return True
        except Exception as e:
            logger.error(f"General error completing daily set workflow: {str(e)}")
            try:
                 self.driver.get(self.base_url)
                 time.sleep(7)
                 self.dismiss_banners()
            except:
                 logger.warning("Failed to refresh page after general daily set error.")
            raise


    def is_other_activity_complete(self, card_element):
         """Checks if an other activity card element visually indicates completion (green checkmark)."""
         # This is the simplified check based on the green checkmark only.
         try:
             # Check the 'complete' attribute (seen on mee-rewards-points parent)
             try:
                 points_parent = card_element.find_element(By.XPATH, "./ancestor::mee-rewards-points")
                 complete_attr = points_parent.get_attribute("complete")
                 if complete_attr and complete_attr.lower() == "true":
                      logger.debug(f"Item appears complete via mee-rewards-points@complete='true'.")
                      return True
             except NoSuchElementException:
                  pass

             # Check the 'state' attribute if it exists (seen on some elements)
             state = card_element.get_attribute("state")
             if state and state.lower() == "complete":
                 logger.debug(f"Item appears complete via state attribute: {state}")
                 return True

             # Check for the specific green checkmark icon within the element (confirmed in HTML)
             checkmark_icon = card_element.find_elements(By.XPATH, ".//span[contains(@class, 'mee-icon-SkypeCircleCheck')]")
             if any(el.is_displayed() for el in checkmark_icon):
                 logger.debug("Item appears complete via green checkmark icon.")
                 return True

             # Check for a common 'completed' class on the element itself or descendants
             completed_class_elements = card_element.find_elements(By.XPATH, ".//*[contains(@class, 'completed')]")
             if any(el.is_displayed() for el in completed_class_elements):
                  logger.debug("Item appears complete via 'completed' class.")
                  return True

             # If none of the above are found, assume not complete
             return False
         except Exception as e:
             logger.debug(f"Error checking other activity item completion status: {e}. Assuming not complete for now.")
             return False # If checking fails, assume not complete


    def get_other_activity_status(self, card_element):
         """Checks the status of an other activity card element: 'completed' or 'actionable'."""
         # Simplifies status to only check for completion. If not completed, it's actionable.

         try:
             # Check for completion (Green checkmark, etc.)
             if self.is_other_activity_complete(card_element):
                 return "completed"

             # We will NOT explicitly check for "locked" status based on blue lock icon or is-exclusive-locked-item.
             # We WILL check for explicit disabling like aria-disabled="true" or 'locked-card' class
             # because these might make the element truly non-interactable by Selenium.

             try:
                 aria_disabled = card_element.get_attribute("aria-disabled")
                 if aria_disabled and aria_disabled.lower() == 'true':
                     logger.debug("Other Activity item appears explicitly disabled (aria-disabled=true).")
                     return "locked" # Treat explicitly disabled as locked/unavailable

                 # Check for disabling classes on the element itself or parents
                 # Check the rewards-card-container parent div for 'locked-card' class
                 try:
                     parent_card_container = card_element.find_element(By.XPATH, "./ancestor::div[contains(@class, 'rewards-card-container')]")
                     if parent_card_container:
                          locked_status_ng_class = parent_card_container.get_attribute("ng-class")
                          if locked_status_ng_class and "'locked-card'" in locked_status_ng_class:
                               logger.debug("Other Activity item appears locked via parent ng-class 'locked-card'.")
                               return "locked"
                 except NoSuchElementException:
                      pass # Ignore if parent not found

             except Exception as e:
                  logger.debug(f"Error checking explicit disabled status for other activity item: {e}")
                  pass


             # If not completed and not explicitly disabled, it's actionable
             return "actionable"

         except Exception as e:
             logger.warning(f"Error checking other activity item status: {e}. Assuming 'actionable' for now.")
             return "actionable" # If checking fails, assume actionable


    def handle_activity_page(self):
        """Handles basic interactions on an activity page (quizzes, polls, etc.) after clicking a card."""
        logger.info("Attempting interactions on activity page...")
        try:
             # Wait for body to ensure page has loaded
             WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
             time.sleep(random.uniform(3, 6)) # Initial wait

             # Try basic interactions on the new page (e.g., quizzes, polls)
             # Use a broader range of potential interactive elements
             interactive_elements_xpath = "//input[@type='radio'] | //div[contains(@class, 'option') or contains(@class, 'choice')] | //button[contains(text(), 'Submit') or contains(text(), 'Next') or contains(text(), 'Play')] | //a[contains(@class, 'btn') or contains(@class, 'button')] | //button | //a[contains(@href, '')] | //div[@tabindex='0' and (contains(@role, 'button') or contains(@role, 'option'))] | //span[contains(@class, 'answer') or contains(@class, 'option')] | //label[contains(@class, 'option')]"

             interactive_elements = self.driver.find_elements(By.XPATH, interactive_elements_xpath)
             logger.info(f"Found {len(interactive_elements)} potential interactive elements on activity page.")

             if interactive_elements:
                 # Filter for visible and interactable elements before sampling
                 interactable_candidates = [el for el in interactive_elements if el.is_displayed() and el.is_enabled()]
                 logger.debug(f"Found {len(interactable_candidates)} interactable candidates.")

                 if interactable_candidates:
                     # Click a random sample of interactive elements
                     random_elements_to_click = random.sample(interactable_candidates, min(len(interactable_candidates), 5)) # Click up to 5 random elements
                     logger.info(f"Clicking {len(random_elements_to_click)} random interactive elements.")
                     for j, el in enumerate(random_elements_to_click):
                         try:
                              # Re-find the specific element before clicking, in case the list became stale
                              # Use a short wait as presence was just confirmed
                              clickable_interactive_element = WebDriverWait(self.driver, 3).until(
                                  EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(el)))
                              )
                              self.driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});", clickable_interactive_element)
                              time.sleep(0.5) # Small pause after scrolling
                              self.driver.execute_script("arguments[0].click();", clickable_interactive_element)
                              logger.debug(f"Clicked interactive element {j+1}/{len(random_elements_to_click)}")
                              time.sleep(random.uniform(2, 4)) # Wait after clicking an interactive element
                         except TimeoutException:
                             logger.debug(f"Interactive element {j+1} not clickable within timeout during interaction.")
                             pass # Continue trying other random elements
                         except StaleElementReferenceException:
                             logger.debug(f"Interactive element {j+1} became stale during interaction. Skipping interaction for this element.")
                             pass # Continue trying other random elements
                         except ElementClickInterceptedException:
                             logger.debug(f"Click on interactive element {j+1} intercepted during interaction. Skipping interaction for this element.")
                             pass
                         except Exception as interact_err:
                             logger.debug(f"Could not click interactive element {j+1}: {interact_err}. Continuing.")
                             pass # Continue trying other random elements
                 else:
                      logger.info("No visible and interactable common interactive elements found to click.")

             else:
                  logger.info("No common interactive elements found to click on activity page.")

             # Stay on the page for a little longer regardless of interaction attempts
             logger.info("Staying on activity page for sufficient time...")
             time.sleep(random.uniform(5, 10))

        except Exception as e:
            logger.warning(f"Error during activity page interaction: {str(e)}")
            # Don't re-raise, just log and continue, as the task might complete just by visiting


    def complete_other_activities(self):
        """Complete other available point activities"""
        try:
            logger.info("Looking for other point activities...")

            # Ensure on the rewards page and dismiss banners
            if self.base_url not in self.driver.current_url:
                logger.info("Navigating to rewards dashboard for other activities.")
                self.driver.get(self.base_url)
                time.sleep(5)
            self.dismiss_banners() # Dismiss banners before finding elements


            # --- Define Other Activities Container XPath ---
            # Use the confirmed ID for the container
            activities_container_xpath = "//*[@id='more-activities']"

            # --- Define Clickable Other Activities Cards XPaths ---
            card_clickable_xpaths_in_container = [
                ".//div[contains(@class, 'rewards-card-container')]/a[contains(@class, 'ds-card-sec')]", # Specific structure confirmed
                ".//div[contains(@class, 'more-earning-card-item')]/a", # Structure seen in Daily Set, might apply here
                ".//a[contains(@class, 'ds-card-sec')]",       # Alternative targeting the specific class
                ".//div[contains(@class, 'rewards-card')]//mee-card", # Broader class from source, target mee-card
                ".//div[contains(@class, 'promo-item')]//mee-card", # Another common promo pattern
                ".//mee-card//a[contains(@href, '')]"          # Fallback: any link within a mee-card
            ]


            # List to store unique identifiers of tasks found
            task_identifiers = []

            try:
                 logger.debug(f"Attempting to find other activities container ({activities_container_xpath}) or cards.")

                 found_container = None
                 visible_cards = [] # Initialize visible_cards

                 try:
                      # Try to find the container first
                      found_container = WebDriverWait(self.driver, 10).until(
                          EC.visibility_of_element_located((By.XPATH, activities_container_xpath))
                      )
                      logger.info(f"Found other activities container.")

                      # Find visible cards *within* the container
                      all_candidate_cards = found_container.find_elements(By.XPATH, " | ".join([xp for xp in card_clickable_xpaths_in_container if xp.startswith('.//')]))
                      visible_cards = [card for card in all_candidate_cards if card.is_displayed()]
                      logger.info(f"Identified {len(visible_cards)} visible other activity cards within container.")

                 except TimeoutException:
                      logger.warning(f"Other activities container not found or no cards within it. Falling back to searching the entire page.")
                      # Fallback: Search the whole page if the container or cards within it weren't found
                      all_candidate_cards = self.driver.find_elements(By.XPATH, " | ".join([xp for xp in card_clickable_xpaths_in_container if xp.startswith('//')])) # Use absolute XPaths for page search
                      visible_cards = [card for card in all_candidate_cards if card.is_displayed()]
                      logger.info(f"Identified {len(visible_cards)} visible other activity cards using page-wide fallback.")

                 except Exception as e:
                      logger.error(f"Error finding initial other activity elements: {str(e)}. Skipping.")
                      return False


                 if not visible_cards:
                     logger.info("No visible other activity cards found initially. Other activities likely already completed or not available.")
                     return True # Consider this success if no tasks are found

                 # Collect identifiers for processing loop
                 for idx, card in enumerate(visible_cards):
                     # Get unique identifiers for the card
                     identifiers = {}
                     identifiers['original_index'] = idx
                     identifiers['href'] = card.get_attribute("href")
                     identifiers['data_bi_id'] = card.get_attribute("data-bi-id")
                     identifiers['data_m_attr'] = card.get_attribute("data-m")

                     # Fallback to text if no reliable identifier found
                     if not identifiers['href'] and not identifiers['data_bi_id'] and not identifiers['data_m_attr']:
                         try:
                            text_element = card.find_elements(By.XPATH, ".//h3 | .//div[contains(@class, 'card-title')] | .//*[string-length(normalize-space()) > 0]")
                            identifiers['id'] = text_element[0].text.strip() if text_element else f"Unknown_OtherActivity_{idx}"
                            if len(identifiers['id']) > 50: identifiers['id'] = identifiers['id'][:50] + "..."
                         except:
                            identifiers['id'] = f"Unknown_OtherActivity_{idx}"
                         logger.warning(f"No reliable ID (href, data-bi-id, data-m) for card {idx}, using fallback identifier '{identifiers['id']}'.")
                     else:
                         # Use a primary ID for logging if available
                         identifiers['id'] = identifiers.get('href') or identifiers.get('data_bi_id') or identifiers.get('data_m_attr')[:50] + '...' if identifiers.get('data_m_attr') else f"ID_Found_{idx}"


                     task_identifiers.append(identifiers)

                 logger.info(f"Collected {len(task_identifiers)} other activity task identifiers.")


            except Exception as e: # Catch exceptions during the initial finding process
                 logger.error(f"Error during initial scan for other activities: {str(e)}. Skipping.")
                 return False # Indicate failure


            # --- Process Found Cards using Identifiers ---
            task_statuses = {info['original_index']: 'initial' for info in task_identifiers}

            # Process tasks by their original index order
            for original_index in sorted(task_statuses.keys()):
                 # Find the task_info for this original_index
                 task_info = next((item for item in task_identifiers if item['original_index'] == original_index), None)
                 if not task_info: # Should not happen if logic is correct
                      logger.error(f"Could not find task_info for original index {original_index}. Skipping.")
                      continue

                 offer_id = task_info['id']
                 href_to_find = task_info.get('href')
                 data_bi_id_to_find = task_info.get('data_bi_id')
                 data_m_attr_to_find = task_info.get('data_m_attr')


                 logger.info(f"Processing other activity task (original index {original_index}): '{offer_id}'...")

                 # --- Attempt to process the specific card with retries ---
                 max_retries = 3
                 task_processed_successfully = False

                 for retry_count in range(max_retries):
                     card_element = None
                     current_task_status = 'unknown'

                     try:
                         # --- Navigate back to Rewards Dashboard and Re-find Elements ---
                         logger.debug(f"Navigating back to Rewards dashboard to re-find element (Retry {retry_count+1}/{max_retries})...")
                         self.driver.get(self.base_url)
                         time.sleep(7)
                         self.dismiss_banners()

                         # Re-find the container if it was originally found, otherwise search the whole page
                         search_context_retry = None
                         if found_container:
                              try:
                                   search_context_retry = WebDriverWait(self.driver, 15).until(
                                       EC.visibility_of_element_located((By.XPATH, activities_container_xpath))
                                   )
                                   logger.debug("Re-found other activities container.")
                              except TimeoutException:
                                   logger.warning("Other activities container not found after returning. Falling back to page search.")
                                   search_context_retry = self.driver # Fallback to page search context
                              except Exception as e:
                                  logger.warning(f"Error re-finding other activities container: {e}. Falling back to page search.")
                                  search_context_retry = self.driver # Fallback

                         else: # If container wasn't found initially, always search the whole page
                              search_context_retry = self.driver
                              logger.debug("Using page search context to re-find other activities.")


                         # Re-find the SPECIFIC element using its identifier within the fresh context
                         logger.debug(f"Re-finding other activity card by identifiers within context...")
                         current_card_xpaths = [xp for xp in card_clickable_xpaths_in_container if (search_context_retry != self.driver and xp.startswith('.//')) or (search_context_retry == self.driver and xp.startswith('//'))]
                         all_current_visible_cards = [card for card in search_context_retry.find_elements(By.XPATH, " | ".join(current_card_xpaths)) if card.is_displayed()]

                         found_match = False
                         for current_card in all_current_visible_cards:
                             current_href = current_card.get_attribute("href")
                             current_data_bi_id = current_card.get_attribute("data-bi-id")
                             current_data_m_attr = current_card.get_attribute("data-m")

                             if (href_to_find and current_href == href_to_find) or \
                                (data_bi_id_to_find and current_data_bi_id == data_bi_id_to_find) or \
                                (data_m_attr_to_find and current_data_m_attr == data_m_attr_to_find):
                                 card_element = current_card
                                 logger.debug(f"Matched other activity card by identifier on retry {retry_count+1}.")
                                 found_match = True
                                 WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(card_element))))
                                 break

                         if not found_match and original_index < len(all_current_visible_cards):
                              card_element = all_current_visible_cards[original_index]
                              logger.debug(f"Matched other activity card by index {original_index} on retry {retry_count+1}.")
                              WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(card_element))))


                         if card_element is None:
                                   logger.warning(f"Could not re-find visible element for other activity task '{offer_id}' on retry {retry_count+1}/{max_retries}. Skipping processing for this task.")
                                   break


                         # --- If element re-found, check status and interact ---
                         if card_element:
                             current_task_status = self.get_other_activity_status(card_element)
                             if current_task_status == "completed":
                                 logger.info(f"Other activity task '{offer_id}' appears completed.")
                                 task_statuses[original_index] = "completed"
                                 task_processed_successfully = True
                                 break

                             # We do not check for "locked" status here anymore.
                             # If status is "actionable"
                             logger.info(f"Other activity task '{offer_id}' is actionable. Attempting interaction (Retry {retry_count+1}/{max_retries})...")

                             # Scroll to the card and click
                             try:
                                 self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", card_element)
                                 time.sleep(1)
                             except Exception as scroll_err:
                                 logger.debug(f"Scroll failed for other activity card: {scroll_err}")
                                 pass

                             try:
                                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, self.get_element_xpath(card_element))))
                                self.driver.execute_script("arguments[0].click();", card_element)
                                logger.info(f"Clicked other activity task '{offer_id}' successfully.")
                             except Exception as click_err:
                                 logger.warning(f"JS click failed for other activity task '{offer_id}': {click_err}. Retrying.")
                                 if retry_count == max_retries - 1:
                                      logger.error(f"Max retries reached for JS click on other activity task '{offer_id}'. Skipping task.")
                                      break
                                 time.sleep(2)
                                 continue

                             # --- Handle Activity Page (New Tab or In-Page) ---
                             initial_window_handle = self.driver.current_window_handle
                             time.sleep(2)
                             window_handles_after = self.driver.window_handles

                             if len(window_handles_after) > len(self.driver.window_handles):
                                 logger.error("Window handles increased unexpectedly after click. This might indicate a problem with tab handling.")
                                 try:
                                      self.driver.switch_to.window(initial_window_handle)
                                      logger.warning("Switched back to original window. Skipping interaction on potential new tab.")
                                 except:
                                      logger.critical("Failed to switch back to original window. Cannot continue.")
                                      raise
                                 task_statuses[original_index] = "attempted"


                             elif len(window_handles_after) > 1:
                                 new_window_handle = [handle for handle in window_handles_after if handle != initial_window_handle]
                                 if new_window_handle:
                                      self.driver.switch_to.window(new_window_handle[-1])
                                      logger.info(f"Switched to new tab for other activity task '{offer_id}'")

                                      self.handle_activity_page() # Call dedicated handler

                                      logger.info("Closing other activity tab and switching back.")
                                      try:
                                          self.driver.close()
                                          self.driver.switch_to.window(initial_window_handle)
                                          time.sleep(3)
                                      except Exception as close_err:
                                           logger.error(f"Error closing activity tab or switching back: {close_err}. Recovery attempt.")
                                           try:
                                                self.driver.get(self.base_url)
                                                time.sleep(7)
                                                self.dismiss_banners()
                                           except:
                                                logger.critical("Failed to navigate back to rewards dashboard after tab error. Cannot reliably continue other activities.")
                                                raise
                                 else:
                                     logger.warning(f"Expected new tab after clicking other activity task '{offer_id}', but no new handle found. Assuming in-page or error.")


                             else:
                                 logger.warning(f"Clicking other activity task '{offer_id}' did not open a new tab or multiple windows already existed. Assuming in-page activity or simple link. Waiting...")
                                 time.sleep(random.uniform(10, 15))
                                 logger.info("Finished waiting after in-page interaction attempt.")

                             task_statuses[original_index] = "attempted"
                             task_processed_successfully = True
                             break


                     # --- Except blocks for retry attempts ---
                     except TimeoutException:
                          logger.warning(f"Timeout waiting for element/page on retry {retry_count+1}/{max_retries} for other activity task '{offer_id}'. Retrying.")
                          time.sleep(2)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for other activity task '{offer_id}' due to Timeout. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                     except StaleElementReferenceException:
                          logger.warning(f"Stale element reference on retry {retry_count+1}/{max_retries} for other activity task '{offer_id}'. Re-finding element and retrying.")
                          time.sleep(2)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for other activity task '{offer_id}' due to Stale Element. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                     except ElementClickInterceptedException as ice:
                          logger.warning(f"Click intercepted on retry {retry_count+1}/{max_retries} for other activity task '{offer_id}': {ice}. Retrying.")
                          time.sleep(2)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for other activity task '{offer_id}' due to Element Click Intercepted. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                     except ElementNotInteractableException as eint_err:
                         logger.warning(f"Element not interactable on retry {retry_count+1}/{max_retries} for other activity task '{offer_id}': {eint_err}. Skipping task.")
                         task_statuses[original_index] = "skipped_not_interactable"
                         break

                     except Exception as e:
                          logger.error(f"Unexpected error on retry {retry_count+1}/{max_retries} for other activity task '{offer_id}': {str(e)}. Retrying.")
                          time.sleep(3)
                          if retry_count == max_retries - 1:
                               logger.error(f"Max retries reached for other activity task '{offer_id}' due to unexpected error. Skipping task.")
                               task_statuses[original_index] = "failed"
                          continue

                 # --- After the retry loop finishes for a single task ---
                 if not task_processed_successfully:
                      logger.warning(f"Other activity task '{offer_id}' was not successfully processed after {max_retries} retries.")
                 else:
                      if task_statuses[original_index] == 'initial':
                           task_statuses[original_index] = 'attempted'


                 logger.info(f"Finished processing logic for other activity task '{offer_id}'. Final Status: {task_statuses[original_index]}.")


            logger.info("Finished attempting other activities.")
            # Log final status of all tasks attempted
            for original_index, status in task_statuses.items():
                 offer_id = next((item['id'] for item in task_identifiers if item['original_index'] == original_index), 'N/A')
                 logger.info(f"Other Activity Task (Original Index {original_index}, ID: '{offer_id}'): {status}")

            return True
        except Exception as e:
            logger.error(f"General error completing other activities workflow: {str(e)}")
            try:
                 self.driver.get(self.base_url)
                 time.sleep(7)
                 self.dismiss_banners();
            except:
                 logger.warning("Failed to refresh page after general other activities error.")
            raise


    # Helper function to generate robust XPath for a specific element
    def get_element_xpath(self, element):
        """Generates a simple but reasonably robust XPath for a given element."""
        try:
            # Try to get XPath using JavaScript for robustness
            xpath = self.driver.execute_script("""
                getElementXPath = function(element) {
                    if (!element) return null;
                    // Prioritize ID
                    if (element.id) {
                        return "//*[@id='" + element.id + "']";
                    }
                    // Prioritize data-bi-id (common on rewards cards)
                    if (element.hasAttribute('data-bi-id')) {
                         return "//*[@data-bi-id='" + element.getAttribute('data-bi-id') + "']";
                    }
                     // Prioritize data-offer-id (common on some rewards cards)
                    if (element.hasAttribute('data-offer-id')) {
                         return "//*[@data-offer-id='" + element.getAttribute('data-offer-id') + "']";
                    }
                    // Prioritize data-m (contains structured data) - Use the exact string for XPath if it exists
                     if (element.hasAttribute('data-m')) {
                         let dataM = element.getAttribute('data-m');
                         // Use the exact data-m string in XPath, escaping single quotes
                         try {
                            let escapedDataM = dataM.replace(/'/g, "',\"'\",'");
                            return "//*[@data-m='" + escapedDataM + "']";
                         } catch (e) {
                             // console.error("Error escaping data-m for XPath:", e); // Avoid logging JS errors to main log
                             // Fallback to simpler XPath if escaping fails
                         }
                     }
                    // Prioritize href for links
                     if (element.tagName === 'A' && element.hasAttribute('href')) {
                         let href = element.getAttribute('href');
                         // Use contains with the full href for uniqueness if it seems stable
                         return "//a[contains(@href, '" + href + "')]";
                     }
                    // Check for common classes and text content (truncated) - less unique
                    if (element.classList.length > 0) {
                         if (element.textContent && element.textContent.trim().length > 5) {
                              let textPart = element.textContent.trim().substring(0, 20).replace(/'/g, "',\"'\",'"); // Truncate & escape
                               return ".//" + element.tagName.toLowerCase() + "[contains(@class, '" + element.classList[0] + "') and contains(text(), '" + textPart + "')]";
                         }
                         // Fallback to just tag and first class
                         return ".//" + element.tagName.toLowerCase() + "[contains(@class, '" + element.classList[0] + "')]";
                    }

                    // Fallback to tag name (least specific)
                    return ".//" + element.tagName.toLowerCase();
                };
                return getElementXPath(arguments[0]);
            """, element)
            if xpath and "null" not in xpath and xpath.strip() != "":
                return xpath
            else:
                 logger.debug("JS XPath generation failed or returned empty, falling back to tag name.")
                 return f".//{element.tag_name}"

        except Exception as e:
            logger.debug(f"Error getting XPath via JS: {e}. Falling back to tag name.")
            return f".//{element.tag_name}"


    def check_points_balance(self):
        """Check and log current points balance"""
        try:
            logger.info("Checking points balance...")
            # Navigate to the dashboard if not already there
            if self.base_url not in self.driver.current_url:
                 logger.info("Navigating to rewards dashboard to check points.")
                 self.driver.get(self.base_url)
                 time.sleep(5)
            self.dismiss_banners() # Dismiss banners before checking points

            points = "Unknown"

            # Updated XPaths based on provided HTML, prioritizing structure
            points_xpaths = [
                "//mee-rewards-user-status-banner//p[contains(@class, 'pointsValue')]//span", # Confirmed structure
                "//mee-rewards-user-status-banner//mee-rewards-counter-animation/span", # Confirmed structure (might be same as above)
                "//div[contains(@class, 'points-package')]//span[contains(@class, 'points-label')]", # Common pattern (fallback)
                "//p[contains(@class, 'points')] | //span[contains(@class, 'points')]", # Broader classes (fallback)
                "//div[contains(@class, 'mee-rewards-counter')]//span[string-length(normalize-space()) > 0]", # Counter element (fallback)
                "//mee-rewards-user-status-banner//div[contains(@class, 'pointsBalance')]//span" # Specific to the user status banner (fallback)
            ]

            # Iterate through XPaths and try to find the element and get its text
            for xpath in points_xpaths:
                try:
                    logger.debug(f"Attempting points check with XPath: {xpath}")
                    # Wait for the element to be visible and have text
                    points_element = WebDriverWait(self.driver, 10).until( # Shorter wait per XPath
                         EC.visibility_of_element_located((By.XPATH, xpath))
                    )
                    WebDriverWait(self.driver, 5).until(
                         EC.text_to_be_present_in_element((By.XPATH, xpath), "")
                    )

                    raw_text = points_element.text.strip()
                    aria_label_text = points_element.get_attribute("aria-label") # Check aria-label as seen in HTML

                    # Prioritize aria-label if it's a number, fallback to text
                    text_to_check = aria_label_text if aria_label_text and aria_label_text.replace(",", "").isdigit() else raw_text

                    # Basic validation: check if it's non-empty and looks like a number (possibly with commas)
                    # Also, exclude common non-point texts found in similar elements
                    if text_to_check and text_to_check.replace(",", "").isdigit() and len(text_to_check.replace(",", "")) > 0: # Check if it looks like a number with at least one digit
                         points = text_to_check
                         logger.info(f"Current points balance found: {points} (using XPath: {xpath})")
                         return points # Return immediately on success
                    else:
                         logger.debug(f"XPath {xpath} found element, but text '{raw_text}' / aria_label '{aria_label_text}' did not look like points. Trying next XPath.")
                         pass # Try next XPath

                except TimeoutException:
                    logger.debug(f"Points element not found/visible with XPath: {xpath} within timeout.")
                    pass # Try next XPath
                except Exception as e:
                    logger.warning(f"Error during points check with XPath {xpath}: {e}. Trying next XPath.")
                    pass # Try next XPath

            logger.warning("Could not find points balance using any known XPaths.")
            points = "Unknown"

        except Exception as e:
            logger.error(f"General error checking points balance workflow: {str(e)}")
            points = "Error checking" # Indicate failure

        logger.info(f"Final reported points balance check result: {points}")
        return points


    def run_complete_workflow(self, nosearch=False):
        """Run the complete workflow of all tasks"""
        success = False # Assume failure initially
        try:
            logger.info("-" * 40)
            logger.info("Starting complete Microsoft Rewards workflow")
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            logger.info(f"Workflow started at: {current_time}")
            logger.info("-" * 40)


            # Setup the driver instance for this run
            # This must be done inside run_complete_workflow because each scheduled run
            # creates a new bot instance.
            self.setup_driver()

            # Attempt login or verify existing session
            # login method now handles initial navigation and status check
            if self.login():
                # Check initial points balance after successful login
                # This check should now work reliably and banners are dismissed
                initial_points = self.check_points_balance()

                # Conditionally perform searches
                if not nosearch:
                    # Ensure we start from a clean Bing page for desktop searches
                    self.driver.get(self.bing_url)
                    time.sleep(3)
                    self.perform_searches(count=self.desktop_search_count, mobile=False)

                    # Now perform mobile searches
                    # Navigate again to reset state before setting mobile UA
                    self.driver.get(self.bing_url)
                    time.sleep(3)
                    self.perform_searches(count=self.mobile_search_count, mobile=True)

                    # Reset user agent to default desktop after mobile searches (optional but clean)
                    try:
                         self.driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": ""})
                         self.driver.maximize_window()
                         logger.info("Reset user agent to default desktop and maximized window.")
                         time.sleep(2) # Wait after resetting UA and size
                    except Exception as ua_reset_err:
                         logger.warning(f"Failed to reset user agent/window size: {ua_reset_err}")
                         pass


                else:
                    logger.info("Skipping searches due to --nosearch flag.")

                # Daily set and Other activities now handle their own navigation back
                # No explicit navigation to dashboard needed here between searches and activities

                # Complete daily set
                daily_set_success = False
                try:
                    daily_set_success = self.complete_daily_set()
                except Exception as e:
                    logger.error(f"Exception caught during daily set workflow: {str(e)}. Continuing to other activities.")
                    daily_set_success = False


                # Complete other activities
                other_activities_success = False
                try:
                    other_activities_success = self.complete_other_activities()
                except Exception as e:
                    logger.error(f"Exception caught during other activities workflow: {str(e)}. Finishing workflow.")
                    other_activities_success = False


                # Navigate back to rewards dashboard before final points check
                # Redundant check here as check_points_balance now navigates itself, but harmless
                # logger.info("Navigating back to rewards dashboard for final points check.")
                # self.driver.get(self.base_url)
                # time.sleep(7) # Wait longer after navigation
                # self.dismiss_banners() # Dismiss banners before checking final points


                # Check final points balance
                final_points = self.check_points_balance() # This function now handles banner dismissal

                logger.info(f"Workflow completed. Points: {initial_points} -> {final_points}")
                success = daily_set_success and other_activities_success # Overall success if both activities completed successfully
            else:
                logger.error("Workflow aborted due to login failure.")
                success = False # Mark as failure

            # Ensure the driver is quit in the finally block
            return success

        except Exception as e:
            logger.critical(f"Critical unexpected error during complete workflow: {str(e)}")
            success = False # Mark as failure
            # In case of a critical error, it's safer to quit the driver here as well before the finally block
            try:
                 self.quit_driver()
            except:
                 pass # Ignore errors during quit in exception handler
            return success # Return failure status

        finally:
            # This block runs whether there was an exception or not
            # Ensure the driver is quit cleanly regardless of success/failure
            self.quit_driver()
            logger.info("-" * 40)
            logger.info("Workflow process finished. Browser window is closed.")
            logger.info("-" * 40)


# Function to run the bot on schedule
def run_rewards_bot(nosearch=False):
    logger.info("Starting scheduled run of Microsoft Rewards Bot.")

    # Create user data directory for Edge browser persistent profile
    home_dir = pathlib.Path.home();
    # Use a hidden folder specific to this script
    edge_profile_dir = os.path.join(home_dir, ".ms_rewards_automation_profile")
    # The bot constructor will create this directory if it doesn't exist

    # Create a NEW bot instance for each scheduled run
    # This ensures a fresh WebDriver instance is created each time, using the persistent profile.
    bot = MicrosoftRewardsBot(
        user_data_dir=edge_profile_dir
    )

    # Run the workflow
    success = bot.run_complete_workflow(nosearch=nosearch)

    if not success:
         logger.error("Scheduled run finished with errors.")
    else:
         logger.info("Scheduled run completed successfully.")

    logger.info("-" * 40)
    logger.info("Scheduled run process finished. Waiting for next run.")
    logger.info("-" * 40)

    # The bot object and its associated WebDriver instance are quit
    # in the finally block of run_complete_workflow.
    # The browser process should be closed, but the profile data is saved
    # in user_data_dir for the next run.


# Schedule the bot to run daily
def setup_schedule(schedule_time_str="10:00", nosearch=False):
    """Sets up the daily schedule for the bot."""
    try:
        # Validate the time format first
        datetime.strptime(schedule_time_str, '%H:%M').time() # Just check if parsing works

        schedule.every().day.at(schedule_time_str).do(run_rewards_bot, nosearch=nosearch)

        logger.info(f"Scheduler set up. Bot will run daily at {schedule_time_str} {'(without searches)' if nosearch else ''}")
        logger.info("Press Ctrl+C to exit the scheduler.")

        # Run once immediately when the script starts, using the provided arguments
        logger.info("Running workflow immediately on script start...")
        # Pass the parsed arguments to the initial run
        run_rewards_bot(nosearch=args.nosearch) # Use args.nosearch from main

        logger.info("Initial run completed. Entering scheduling loop.")

        # Then keep running on schedule
        while True:
            schedule.run_pending(); time.sleep(60) # Check the schedule every minute

    except ValueError:
        logger.error(f"Invalid time format '{schedule_time_str}'. Please use HH:MM format (e.g., '10:00'). Scheduling aborted.")
        # Do not enter the while loop if time format is invalid
    except KeyboardInterrupt:
        logger.info("Script terminated by user.")
    except Exception as e:
        logger.critical(f"Unhandled error in scheduling loop: {str(e)}")


if __name__ == "__main__":
    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(description='Automate Microsoft Rewards tasks.')
    parser.add_argument('--nosearch', action='store_true',
                        help='Skip the Bing search tasks (both desktop and mobile).')
    parser.add_argument('--time', type=str, default='10:00',
                        help='Specify the daily schedule time in HH:MM format. Default is 10:00.')
    args = parser.parse_args()

    # --- Setup and Run ---
    # Pass the parsed arguments to the schedule setup function
    setup_schedule(schedule_time_str=args.time, nosearch=args.nosearch)
