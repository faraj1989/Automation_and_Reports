"""Continuous NetEco Historical Alarms exporter - copied verbatim
(2026-09-10) from the sibling NOC Automation Suite (Automation-master),
then wired up to run in this project, same as scrapers/mae_scraper.py.
Produces the rolling-window NetEco_Historical_Alarm_<timestamp>.zip/csv
exports.

Same login/session/export logic as scrapers/neteco_continuous all alrams.py
(same eviewwebsite portal, same NETECO_USERNAME/NETECO_PASSWORD) with only the
report URL swapped from the All Alarms view to Historical Alarms
(fmHistoryAlarm) - historical alarms are already-occurred-and-cleared events,
so this is a separate feed from the live/current alarm scrapers.

The session stays open the whole run: login happens once, then every cycle
just re-navigates to the Historical Alarms URL and clicks Export/All/OK again
so newly occurred/cleared alarms show up without a fresh login each time.
"""
import os
import sys
import shutil
import socket
import time
from urllib.parse import urlparse

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import LoginFailedError, env_int, env_path_str, env_str, mark_login_failed

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ================== USER CONFIGURATION =====================
USERNAME = env_str("NETECO_USERNAME")
PASSWORD = env_str("NETECO_PASSWORD")
URL = env_str("NETECO_URL")
HISTORICAL_ALARMS_URL = env_str("NETECO_HISTORICAL_ALARMS_URL")

DOWNLOAD_DIR = env_path_str(
    "NETECO_HISTORICAL_DOWNLOAD_DIR",
    os.path.join(env_path_str("NETECO_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads")), "Historical"),
)
EXPORT_BASE_DIR = env_path_str("NETECO_HISTORICAL_EXPORT_BASE_DIR", r"C:\Historical_Alarms")

WAIT_TIMEOUT = env_int("NETECO_WAIT_TIMEOUT", 60)
INTERVAL_SECONDS = env_int("NETECO_HISTORICAL_INTERVAL_SECONDS", 420)  # 7 minutes
DOWNLOAD_TIMEOUT_SECONDS = env_int("NETECO_HISTORICAL_DOWNLOAD_TIMEOUT_SECONDS", 600)  # export can take 1+ minutes
DOWNLOAD_STABLE_SECONDS = env_int("NETECO_DOWNLOAD_STABLE_SECONDS", 5)
RETRY_DELAY_SECONDS = env_int("NETECO_RETRY_DELAY_SECONDS", 60)
MAX_CONSECUTIVE_FAILURES = env_int("NETECO_MAX_CONSECUTIVE_FAILURES", 5)
PORT_CHECK_TIMEOUT = env_int("NETECO_PORT_CHECK_TIMEOUT", 5)

TEMP_DOWNLOAD_EXTENSIONS = (".crdownload", ".tmp", ".partial")


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
DATA_ROOT = os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\NOC_Data")
ERROR_SCREENSHOT_DIR = ensure_dir(os.path.join(DATA_ROOT, "Errors"))

# =====================================================

if not USERNAME or not PASSWORD or not URL:
    raise RuntimeError(
        "NETECO_USERNAME, NETECO_PASSWORD, and NETECO_URL must be configured in .env or environment variables."
    )


def log(message):
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}", flush=True)


def wait_and_fill(driver, locator, text, timeout=WAIT_TIMEOUT):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(locator)
    )
    try:
        element.clear()
        element.send_keys(text)
    except WebDriverException:
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
            element,
            text,
        )


def wait_and_click(driver, locator, timeout=WAIT_TIMEOUT):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(locator)
    )
    try:
        element.click()
    except WebDriverException:
        driver.execute_script("arguments[0].click();", element)
    return element


def wait_until_url_port_is_open(url, timeout=PORT_CHECK_TIMEOUT):
    parsed = urlparse(url)
    host = parsed.hostname
    port = parsed.port or (443 if parsed.scheme == "https" else 80)

    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError as exc:
        log(f"Cannot connect to {host}:{port}: {exc}")
        return False


def is_complete_download(path):
    if path.lower().endswith(TEMP_DOWNLOAD_EXTENSIONS):
        return False

    try:
        first_size = os.path.getsize(path)
        time.sleep(DOWNLOAD_STABLE_SECONDS)
        second_size = os.path.getsize(path)
    except OSError:
        return False

    return first_size > 0 and first_size == second_size


def wait_for_download(download_dir, before_files, timeout=DOWNLOAD_TIMEOUT_SECONDS):
    end_time = time.time() + timeout

    while time.time() < end_time:
        current_files = set(os.listdir(download_dir))
        new_files = [
            name
            for name in current_files - before_files
            if not name.lower().endswith(TEMP_DOWNLOAD_EXTENSIONS)
        ]

        newest_first = sorted(
            (os.path.join(download_dir, name) for name in new_files),
            key=os.path.getmtime,
            reverse=True,
        )

        for path in newest_first:
            if os.path.isfile(path) and is_complete_download(path):
                return path

        time.sleep(2)

    return None


def create_driver():
    """Create Chrome driver with proper options - same as the other NetEco scrapers."""
    global driver

    try:
        driver.quit()
    except Exception:
        pass

    log("Starting Chrome...")

    chrome_options = Options()
    chrome_options.set_capability("acceptInsecureCerts", True)
    chrome_options.page_load_strategy = "eager"

    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-insecure-localhost")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--allow-legacy-insecure-renegotiation")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")

    chrome_options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "profile.default_content_settings.popups": 0,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": False,
        },
    )

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        log(f"✅ Chrome started with ChromeDriver: {service.path}")

        caps = driver.capabilities
        browser_version = caps.get("browserVersion") or caps.get("version")
        log(f"🔧 Browser version: {browser_version}")

        driver.set_page_load_timeout(WAIT_TIMEOUT)
        return driver

    except Exception as e:
        log(f"❌ Failed to start Chrome with webdriver_manager: {e}")
        log("   Trying fallback with Selenium Manager...")

        driver = webdriver.Chrome(options=chrome_options)
        log(f"✅ Chrome started with Selenium Manager")
        driver.set_page_load_timeout(WAIT_TIMEOUT)
        return driver


def login_and_navigate(driver):
    log(f"Navigating to: {URL}")
    driver.get(URL)

    WebDriverWait(driver, WAIT_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    log("Filling login form...")
    wait_and_fill(driver, (By.ID, "username"), USERNAME)
    wait_and_fill(driver, (By.ID, "value"), PASSWORD)
    wait_and_click(driver, (By.ID, "btn_outerverify"))
    log("Login submitted.")
    time.sleep(5)

    # Unlike the NCE/MAE-family scrapers, this login flow had no failure
    # detection at all - a bad password just timed out later waiting for
    # Device Management, indistinguishable from any other transient glitch.
    # Checking explicitly lets the caller stop retrying instead of hammering
    # a known-bad password every cycle.
    if driver.find_elements(By.ID, "username"):
        page_text = driver.page_source.lower()
        # Checked BEFORE the credential keywords below: confirmed live
        # 2026-09-10 (on the sibling MAE historical scraper) that a portal's
        # own static page boilerplate can carry the literal word
        # "credentials" even on a rate-limit rejection ("Operations are too
        # frequent. Please try later.") - so that keyword alone
        # false-positived a real LoginFailedError (which stops the scraper
        # for the whole day and burns the login-failure marker) on what was
        # actually a transient multi-scraper-login collision, not a bad
        # password.
        rate_limit_phrases = ["too frequent", "try later", "too many requests", "rate limit"]
        if any(phrase in page_text for phrase in rate_limit_phrases):
            raise RuntimeError("Login rejected by the portal's rate limiter - transient, not a bad password.")
        login_error_keywords = [
            "invalid username", "invalid password", "incorrect username", "incorrect password",
            "wrong username", "wrong password", "password expired", "expired password",
            "credentials", "login failed", "account locked",
        ]
        if any(keyword in page_text for keyword in login_error_keywords):
            raise LoginFailedError("Login failed: invalid credentials or expired password detected.")
        # Deliberately NOT a LoginFailedError: confirmed live 2026-09-07 (in the
        # sibling NCE scrapers) this branch also covers non-credential cases
        # like a portal's rate limiter kicking in when several scrapers log
        # into the same account at once - a transient condition that should
        # keep retrying, not stop for the day.
        raise RuntimeError("Login appears to have failed; login form still visible.")

    dev_mgmt = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.visibility_of_element_located((By.XPATH, "//span[@title='Device Management']"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dev_mgmt)
    time.sleep(2)
    driver.execute_script("arguments[0].click();", dev_mgmt)

    current_alarms_xpath = (
        "//span[normalize-space()='Current Alarms'] | "
        "//a[normalize-space()='Current Alarms']"
    )
    current_alarms_btn = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, current_alarms_xpath))
    )
    driver.execute_script("arguments[0].click();", current_alarms_btn)
    log("Reached Current Alarms (session established).")
    time.sleep(10)


def navigate_to_historical_alarms(driver):
    driver.get(HISTORICAL_ALARMS_URL)
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") in {"interactive", "complete"}
    )
    time.sleep(3)


def click_export_button(driver):
    driver.switch_to.default_content()
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for frame in iframes:
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)
        try:
            export_container = driver.find_element(By.ID, "exportBtn")
            export_button = export_container.find_element(By.TAG_NAME, "button")
            driver.execute_script("arguments[0].click();", export_button)
            return
        except WebDriverException:
            continue

    driver.switch_to.default_content()
    export_container = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "exportBtn"))
    )
    export_button = export_container.find_element(By.TAG_NAME, "button")
    driver.execute_script("arguments[0].click();", export_button)


def run_export_sequence(driver):
    before_files = set(os.listdir(DOWNLOAD_DIR))

    click_export_button(driver)

    all_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "allExport"))
    )
    driver.execute_script("arguments[0].click();", all_option)

    ok_confirm = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "confirmBtn"))
    )
    driver.execute_script("arguments[0].click();", ok_confirm)

    log(f"Waiting for download (timeout {DOWNLOAD_TIMEOUT_SECONDS}s)...")
    downloaded_file = wait_for_download(DOWNLOAD_DIR, before_files)
    if not downloaded_file:
        log("Download timed out. Next export will wait for the normal interval.")
        return False

    dest_dir = os.path.join(EXPORT_BASE_DIR, time.strftime("%Y-%m-%d"))
    dest_dir = ensure_dir(dest_dir)

    extension = os.path.splitext(downloaded_file)[1]
    new_name = f"NetEco_Historical_Alarm_{time.strftime('%Y%m%d_%H%M%S')}{extension}"
    new_path = os.path.join(dest_dir, new_name)

    shutil.move(downloaded_file, new_path)
    log(f"Saved: {new_path}")
    return True


def close_driver(driver):
    if driver:
        try:
            driver.quit()
        except WebDriverException:
            pass


def wait_before_next_export(last_export_attempt_time):
    if not last_export_attempt_time:
        return

    elapsed = time.time() - last_export_attempt_time
    if elapsed < INTERVAL_SECONDS:
        remaining = INTERVAL_SECONDS - elapsed
        log(f"Waiting {remaining / 60:.1f} minutes before next export attempt...")
        time.sleep(remaining)


def main():
    driver = None
    consecutive_failures = 0
    last_export_attempt_time = 0

    while True:
        try:
            wait_before_next_export(last_export_attempt_time)

            if driver is None:
                if not wait_until_url_port_is_open(URL):
                    raise ConnectionError(f"Target is not reachable before opening Chrome: {URL}")

                driver = create_driver()
                login_and_navigate(driver)

            log("Starting export cycle...")
            last_export_attempt_time = time.time()
            navigate_to_historical_alarms(driver)
            run_export_sequence(driver)
            consecutive_failures = 0

        except KeyboardInterrupt:
            log("Stopped by user.")
            close_driver(driver)
            break
        except LoginFailedError as exc:
            log(f"Login failed - stopping for today until the password is fixed: {exc}")
            mark_login_failed("NetEco Historical Alarms Scraper", PASSWORD)
            close_driver(driver)
            break
        except Exception as exc:
            consecutive_failures += 1
            log(f"Error occurred: {exc}")
            log(f"Failure {consecutive_failures}/{MAX_CONSECUTIVE_FAILURES}.")

            try:
                if driver:
                    screenshot_path = os.path.join(ERROR_SCREENSHOT_DIR,
                                                   f"neteco_historical_error_{time.strftime('%Y%m%d_%H%M%S')}.png")
                    driver.save_screenshot(screenshot_path)
                    log(f"📸 Screenshot saved as: {screenshot_path}")
            except Exception as screenshot_error:
                log(f"⚠️ Screenshot failed: {screenshot_error}")

            close_driver(driver)
            driver = None

            if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                raise SystemExit(
                    f"Stopped after {MAX_CONSECUTIVE_FAILURES} consecutive failures. "
                    "Check NetEco service, HTTPS/TLS compatibility, VPN/routing, and credentials."
                )

            log(f"Restarting browser session in {RETRY_DELAY_SECONDS} seconds...")
            time.sleep(RETRY_DELAY_SECONDS)


if __name__ == "__main__":
    main()
