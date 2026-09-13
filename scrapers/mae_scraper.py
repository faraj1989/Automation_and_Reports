"""MAE Current Alarms scraper - copied verbatim (2026-09-09) from the
sibling NOC Automation Suite (Automation-master, scrapers/mae_scraper_
newerBrowserversion149.py), then wired up to run in this project
(2026-09-10) with this project's own project_config.py/.env once scraping
was brought in-house - the sibling suite's own copy was stopped first to
avoid two scrapers logging into the same MAE account at once.

What it does: headless-Chrome Selenium session that logs into the MAE
portal, opens Current Alarms, and every INTERVAL_SECONDS (default 300s)
refreshes the page, clicks Export -> All -> OK, and saves the resulting
CSV as CurrentAlarms_MAE_<timestamp>.csv under EXPORT_BASE_DIR/<date>/,
pruning older exports for the same day (only the latest current-alarm
snapshot is meaningful, unlike the historical scraper's intentionally
accumulating exports)."""
import os
import sys
import time
import shutil
import winreg
import zipfile
from selenium import webdriver
from selenium.common.exceptions import InvalidSessionIdException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import LoginFailedError, env_int, env_path_str, env_str, keep_only_latest_export, mark_login_failed

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ================== USER CONFIG =====================
USERNAME = env_str("MAE_USERNAME")
PASSWORD = env_str("MAE_PASSWORD")
URL = env_str("MAE_URL")
DOWNLOAD_DIR = env_path_str("MAE_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads"))
EXPORT_BASE_DIR = env_path_str("MAE_EXPORT_BASE_DIR", r"C:\Current_Alarms")
WAIT_TIMEOUT = env_int("MAE_WAIT_TIMEOUT", 45)
INTERVAL_SECONDS = env_int("MAE_INTERVAL_SECONDS", 300)


# ========== ENSURE DIRECTORIES EXIST ==========
def ensure_dir(path):
    """Create directory if it doesn't exist."""
    if not path:
        return path
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        print(f"📁 Created directory: {path}")
    return path


# Create directories before proceeding
DOWNLOAD_DIR = ensure_dir(DOWNLOAD_DIR)
EXPORT_BASE_DIR = ensure_dir(EXPORT_BASE_DIR)

# ========== ERROR SCREENSHOT DIRECTORY ==========
# Get the DATA_ROOT from environment or use default
DATA_ROOT = os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\NOC_Data")
ERROR_SCREENSHOT_DIR = ensure_dir(os.path.join(DATA_ROOT, "Errors"))

# =====================================================

if not USERNAME or not PASSWORD or not URL:
    raise RuntimeError("MAE_USERNAME, MAE_PASSWORD, and MAE_URL must be configured in .env or environment variables.")


def wait_for_file(download_dir, before_files, timeout=180):
    end = time.time() + timeout
    while time.time() < end:
        now = set(os.listdir(download_dir))
        new_files = [f for f in (now - before_files) if not any(x in f for x in [".crdownload", ".tmp", ".partial"])]
        if new_files:
            full = [os.path.join(download_dir, f) for f in new_files]
            return max(full, key=os.path.getmtime)
        time.sleep(2)
    return None


def wait_for_ready_state(timeout=WAIT_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def wait_and_fill(locator, value):
    elem = WebDriverWait(driver, WAIT_TIMEOUT).until(EC.visibility_of_element_located(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    elem.clear()
    elem.send_keys(value)
    return elem


def wait_and_click(locator):
    elem = WebDriverWait(driver, WAIT_TIMEOUT).until(EC.element_to_be_clickable(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
    driver.execute_script("arguments[0].click();", elem)
    return elem


def print_driver_versions():
    caps = driver.capabilities
    browser_version = caps.get("browserVersion") or caps.get("version")
    chromedriver_version = None
    if isinstance(caps.get("chrome"), dict):
        chromedriver_version = caps["chrome"].get("chromedriverVersion")
    chromedriver_version = chromedriver_version or caps.get("chromedriverVersion")
    if chromedriver_version:
        chromedriver_version = chromedriver_version.split(" ")[0]
    print(f"🔧 Browser version: {browser_version}")
    print(f"🔧 ChromeDriver version: {chromedriver_version}")


def create_driver():
    global driver
    try:
        driver.quit()
    except Exception:
        pass

    print("🔧 Starting Chrome browser...")
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print(f"🔧 Using ChromeDriver at: {service.path}")
    except Exception as e:
        print(f"webdriver_manager failed ({e}); falling back to Selenium Manager")
        driver = webdriver.Chrome(options=chrome_options)
    print_driver_versions()


# ========== INITIALIZE BROWSER ==========
chrome_options = Options()
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
})

print("🔧 Initializing ChromeDriver...")
create_driver()


def login_and_navigate():
    """Handles initial login and moving to the Current Alarms page."""
    print("🔐 Logging in...")
    driver.get(URL)
    try:
        wait_for_ready_state()

        # Check if login form is present. find_elements() checks once with no
        # retry - on a slow-rendering SPA login page, "readyState complete"
        # can fire before the form itself is in the DOM, making this look
        # logged-in when it's actually just not loaded yet (confirmed live
        # 2026-09-06 in the sibling nce_active_alarms_scraper.py). A short
        # bounded wait tells the two cases apart.
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "username")))
            has_login_form = True
        except Exception:
            has_login_form = False

        if has_login_form:
            print("📝 Filling login credentials...")

            # Fill username using a visibility wait; fallback to JS assignment
            try:
                username_field = wait_and_fill((By.ID, "username"), USERNAME)
            except Exception:
                username_field = WebDriverWait(driver, WAIT_TIMEOUT).until(
                    EC.presence_of_element_located((By.ID, "username"))
                )
                driver.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                    username_field,
                    USERNAME,
                )

            # Fill password using a visibility wait; fallback to JS assignment
            try:
                password_field = wait_and_fill((By.ID, "value"), PASSWORD)
            except Exception:
                password_field = WebDriverWait(driver, WAIT_TIMEOUT).until(
                    EC.presence_of_element_located((By.ID, "value"))
                )
                driver.execute_script(
                    "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                    password_field,
                    PASSWORD,
                )

            print("🔐 Attempting login click...")

            # Try clicking the actual trigger; prefer clickable wait and JS click,
            # but fall back to alternate container if needed.
            try:
                login_span = WebDriverWait(driver, WAIT_TIMEOUT).until(
                    EC.element_to_be_clickable((By.ID, "submitDataverify"))
                )
                driver.execute_script("arguments[0].click();", login_span)
                print("✅ Login clicked via span element")
            except Exception:
                try:
                    login_div = driver.find_element(By.ID, "btn_outerverify")
                    driver.execute_script("arguments[0].click();", login_div)
                    print("✅ Login clicked via btn_outerverify")
                except Exception as e_click:
                    print(f"⚠️ Login click failed: {e_click}")

            # Wait for login to process
            print("⏳ Waiting for login to complete...")
            time.sleep(8)

            login_form_still = len(driver.find_elements(By.ID, "username")) > 0
            if login_form_still:
                page_text = driver.page_source.lower()
                # Checked BEFORE the credential keywords below: confirmed live
                # 2026-09-10 that MAE's login page carries the literal word
                # "credentials" in its own static boilerplate even on a rate-
                # limit rejection ("Operations are too frequent. Please try
                # later.") - so that keyword alone false-positived a real
                # LoginFailedError (which stops the scraper for the whole day
                # and burns the login-failure marker) on what was actually a
                # transient multi-scraper-login collision, not a bad password.
                rate_limit_phrases = ["too frequent", "try later", "too many requests", "rate limit"]
                if any(phrase in page_text for phrase in rate_limit_phrases):
                    raise RuntimeError("Login rejected by the portal's rate limiter - transient, not a bad password.")
                login_error_keywords = [
                    "invalid username",
                    "invalid password",
                    "incorrect username",
                    "incorrect password",
                    "wrong username",
                    "wrong password",
                    "password expired",
                    "expired password",
                    "credentials",
                    "login failed",
                    "account locked",
                ]
                if any(keyword in page_text for keyword in login_error_keywords):
                    raise LoginFailedError("Login failed: invalid credentials or expired password detected.")
                # Deliberately NOT a LoginFailedError: confirmed live 2026-09-07 (in
                # the sibling NCE scrapers) this branch also covers non-credential
                # cases like a portal's rate limiter kicking in when several
                # scrapers log into the same account at once - a transient
                # condition that should keep retrying, not stop for the day.
                raise RuntimeError("Login appears to have failed; login form still visible.")
        else:
            print("ℹ️ Login form not present; assuming session is already authenticated.")

        # Navigate to Current Alarms
        print("📡 Navigating to Current Alarms...")
        current_alarms_xpath = (
            "//div[contains(text(), 'Current Alarms')] | "
            "//span[contains(text(), 'Current Alarms')] | "
            "//a[contains(text(), 'Current Alarms')] | "
            "//*[contains(text(), 'Current Alarms')]"
        )

        try:
            # Wait for Current Alarms to be present
            current_alarm = WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.XPATH, current_alarms_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", current_alarm)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", current_alarm)
            print("✅ Current Alarms clicked")
        except Exception as e:
            print(f"⚠️ Could not find Current Alarms: {e}")
            # Try alternative - look for any element with "Current Alarms" text using partial match
            try:
                elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Current Alarms')]")
                for elem in elements:
                    if elem.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", elem)
                        print("✅ Current Alarms found and clicked via text search")
                        break
            except Exception as e2:
                print(f"⚠️ Alternative click failed: {e2}")
                # Take screenshot for debugging
                screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR,
                                               f"debug_current_alarms_{time.strftime('%Y%m%d_%H%M%S')}.png")
                driver.save_screenshot(screenshot_path)
                print(f"📸 Screenshot saved as: {screenshot_path}")

        time.sleep(10)
        return True
    except LoginFailedError:
        raise  # not a transient nav failure - let it escape to the top-level handler
    except Exception as e:
        print(f"❌ Login/Navigation failed: {e}")
        return False


# ========== MAIN EXECUTION ==========
try:
    # Initial login
    if not login_and_navigate():
        print("❌ Initial login failed. Exiting...")
        driver.quit()
        exit(1)

    # Main loop
    while True:
        try:
            print(f"\n🔄 [{time.strftime('%H:%M:%S')}] Refreshing page for latest data...")
            driver.refresh()
            time.sleep(12)

            # Check if we were logged out
            if "login" in driver.current_url.lower() or len(driver.find_elements(By.ID, "username")) > 0:
                print("⚠️ Session expired during refresh. Re-logging in...")
                if not login_and_navigate():
                    print("❌ Re-login failed. Exiting...")
                    break

            # Switch to the Alarms Iframe
            driver.switch_to.default_content()
            iframe_xpath = "//iframe[contains(@id,'fmAlarmView')]"
            WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.XPATH, iframe_xpath))
            )
            driver.switch_to.frame(driver.find_element(By.XPATH, iframe_xpath))

            # Find and Click Export
            print("🔍 Finding Export button...")
            export_xpath = "//button[normalize-space()='Export']"
            export_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, export_xpath))
            )

            before_files = set(os.listdir(DOWNLOAD_DIR))
            driver.execute_script("arguments[0].click();", export_btn)
            time.sleep(3)

            # Handle Export Dialog
            print("📦 Handling Export Dialog...")
            all_opt = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//*[text()='All']"))
            )
            driver.execute_script("arguments[0].click();", all_opt)

            ok_btn = driver.find_element(By.XPATH, "//span[text()='OK']/ancestor::button")
            driver.execute_script("arguments[0].click();", ok_btn)

            # Wait for download and move file
            downloaded_file = wait_for_file(DOWNLOAD_DIR, before_files)
            if downloaded_file:
                date_folder = time.strftime("%Y-%m-%d")
                dest_dir = os.path.join(EXPORT_BASE_DIR, date_folder)
                dest_dir = ensure_dir(dest_dir)  # Ensure destination exists
                new_filename = f"CurrentAlarms_MAE_{time.strftime('%Y%m%d_%H%M%S')}.csv"
                dest_path = os.path.join(dest_dir, new_filename)

                if zipfile.is_zipfile(downloaded_file):
                    # MAE zips the export once it crosses a size threshold; unwrap it
                    # so downstream readers always see a plain CSV.
                    with zipfile.ZipFile(downloaded_file) as zf:
                        csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
                        if not csv_names:
                            raise RuntimeError(f"MAE export zip had no CSV inside: {zf.namelist()}")
                        with zf.open(csv_names[0]) as src, open(dest_path, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                    os.remove(downloaded_file)
                else:
                    shutil.move(downloaded_file, dest_path)

                keep_only_latest_export(EXPORT_BASE_DIR, "CurrentAlarms_MAE_*.csv", dest_path)
                print(f"✅ Export Success: {new_filename}")
            else:
                print("❌ Download timed out.")

        except LoginFailedError:
            raise  # not a transient cycle error - let it escape to the top-level handler
        except Exception as e:
            print(f"⚠️ Cycle Error: {e}")
            invalid_session = isinstance(e, InvalidSessionIdException) or "invalid session id" in str(e).lower()
            if not invalid_session:
                try:
                    # Save screenshot to Errors folder
                    screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR, f"error_{time.strftime('%Y%m%d_%H%M%S')}.png")
                    driver.save_screenshot(screenshot_path)
                    print(f"📸 Screenshot saved as: {screenshot_path}")
                except Exception as screenshot_error:
                    print(f"⚠️ Screenshot failed: {screenshot_error}")

            try:
                driver.quit()
            except Exception:
                pass

            create_driver()
            time.sleep(5)
            if not login_and_navigate():
                print("❌ Re-login failed. Exiting...")
                break

        # Countdown
        print(f"😴 Sleeping for {INTERVAL_SECONDS // 60} minutes...")
        for m in range(INTERVAL_SECONDS // 60, 0, -1):
            print(f"Next update in {m} minutes...", end="\r")
            time.sleep(60)
        remaining_seconds = INTERVAL_SECONDS % 60
        if remaining_seconds > 0:
            time.sleep(remaining_seconds)

except KeyboardInterrupt:
    print("\n🛑 Automation stopped by user.")
except LoginFailedError as e:
    print(f"❌ Login failed - stopping for today until the password is fixed: {e}")
    mark_login_failed("MAE Scraper", PASSWORD)
finally:
    print("Browser remains open. Close it manually when finished.")
