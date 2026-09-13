"""Continuous MAE Historical Alarms exporter - copied verbatim (2026-09-09)
from the sibling NOC Automation Suite (Automation-master), then wired up
to run in this project (2026-09-10), same as scrapers/mae_scraper.py.
Produces the rolling-window HistoricalAlarms_MAE_<timestamp>.zip/csv
exports this project's backend/noc_alarm_processor.py already reads.

Same login system and Export/All/OK flow as scrapers/mae_scraper_newerBrowserversion149.py
(same OSS/MAE portal, same credentials) but pointed at the Historical Alarms
deep link instead of Current Alarms. The URL already deep-links straight into
fmHistoryAlarm, so there is no "click Current Alarms" navigation step here.

Refreshes the page before every export cycle (like the current-alarms
scraper) rather than reusing the same open export dialog/session across
cycles: reusing it left the export dropdown in a stale state after the first
click, so every export after the first silently never triggered a download
and timed out. Only re-logs in if the refresh reveals the session actually
dropped.
"""
import os
import sys
import shutil
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import LoginFailedError, env_int, env_path_str, env_str, mark_login_failed

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ================== USER CONFIG =====================
USERNAME = env_str("MAE_USERNAME")
PASSWORD = env_str("MAE_PASSWORD")
URL = env_str("MAE_HISTORICAL_URL")
DOWNLOAD_DIR = env_path_str(
    "MAE_HISTORICAL_DOWNLOAD_DIR",
    os.path.join(env_path_str("MAE_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads")), "Historical"),
)
EXPORT_BASE_DIR = env_path_str("MAE_HISTORICAL_EXPORT_BASE_DIR", r"C:\Historical_Alarms")
WAIT_TIMEOUT = env_int("MAE_WAIT_TIMEOUT", 45)
DOWNLOAD_TIMEOUT = env_int("MAE_HISTORICAL_DOWNLOAD_TIMEOUT", 600)
INTERVAL_SECONDS = env_int("MAE_HISTORICAL_INTERVAL_SECONDS", 300)

DATA_ROOT = os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\NOC_Data")


def ensure_dir(path):
    if not path:
        return path
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        print(f"Created directory: {path}")
    return path


DOWNLOAD_DIR = ensure_dir(DOWNLOAD_DIR)
EXPORT_BASE_DIR = ensure_dir(EXPORT_BASE_DIR)
ERROR_SCREENSHOT_DIR = ensure_dir(os.path.join(DATA_ROOT, "Errors"))

if not USERNAME or not PASSWORD:
    raise RuntimeError("MAE_USERNAME and MAE_PASSWORD must be configured in .env or environment variables.")


def wait_for_file(download_dir, before_files, timeout):
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


def print_driver_versions():
    caps = driver.capabilities
    browser_version = caps.get("browserVersion") or caps.get("version")
    chromedriver_version = None
    if isinstance(caps.get("chrome"), dict):
        chromedriver_version = caps["chrome"].get("chromedriverVersion")
    chromedriver_version = chromedriver_version or caps.get("chromedriverVersion")
    if chromedriver_version:
        chromedriver_version = chromedriver_version.split(" ")[0]
    print(f"Browser version: {browser_version}")
    print(f"ChromeDriver version: {chromedriver_version}")


def create_driver():
    options = Options()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    })
    new_service = Service(ChromeDriverManager().install())
    new_driver = webdriver.Chrome(service=new_service, options=options)
    print(f"Using ChromeDriver at: {new_service.path}")
    return new_driver


def is_dead_session_error(exc):
    """True when the whole browser/chromedriver process is gone (its session
    port refuses connections) rather than a page-level hiccup - same fix as
    nce_active_alarms_scraper.py / nce_historical_alarms_scraper.py, confirmed
    live 2026-09-13 for this exact scraper: it looped an identical "invalid
    session id" failure every 5-minute cycle for hours with no recovery,
    since retrying selenium calls against a dead session can never succeed
    on its own."""
    message = str(exc)
    return (
        "actively refused" in message
        or "Max retries exceeded" in message
        or "invalid session id" in message.lower()
        or "session deleted" in message.lower()
        or "chrome not reachable" in message.lower()
    )


print("Initializing ChromeDriver...")
driver = create_driver()
print_driver_versions()


def login():
    print("Logging in...")
    driver.get(URL)
    wait_for_ready_state()

    # find_elements() checks once with no retry - on a slow-rendering SPA
    # login page, "readyState complete" can fire before the form itself is
    # in the DOM, making this look logged-in when it's actually just not
    # loaded yet (confirmed live 2026-09-06 in the sibling
    # nce_active_alarms_scraper.py). A short bounded wait tells the two
    # cases apart.
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "username")))
        has_login_form = True
    except Exception:
        has_login_form = False
    if has_login_form:
        print("Filling login credentials...")
        try:
            username_field = wait_and_fill((By.ID, "username"), USERNAME)
        except Exception:
            username_field = WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                username_field, USERNAME,
            )
        try:
            password_field = wait_and_fill((By.ID, "value"), PASSWORD)
        except Exception:
            password_field = WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.presence_of_element_located((By.ID, "value"))
            )
            driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
                password_field, PASSWORD,
            )

        print("Attempting login click...")
        try:
            login_span = WebDriverWait(driver, WAIT_TIMEOUT).until(
                EC.element_to_be_clickable((By.ID, "submitDataverify"))
            )
            driver.execute_script("arguments[0].click();", login_span)
        except Exception:
            try:
                login_div = driver.find_element(By.ID, "btn_outerverify")
                driver.execute_script("arguments[0].click();", login_div)
            except Exception as e_click:
                print(f"Login click failed: {e_click}")

        print("Waiting for login to complete...")
        time.sleep(8)

        if len(driver.find_elements(By.ID, "username")) > 0:
            page_text = driver.page_source.lower()
            # Checked BEFORE the credential keywords below: confirmed live
            # 2026-09-10 that MAE's login page carries the literal word
            # "credentials" in its own static boilerplate even on a rate-
            # limit rejection ("Operations are too frequent. Please try
            # later.") - so that keyword alone false-positived a real
            # LoginFailedError (which stops the scraper for the whole day and
            # burns the login-failure marker) on what was actually a
            # transient multi-scraper-login collision, not a bad password.
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
            # Deliberately NOT a LoginFailedError: confirmed live 2026-09-07 (in
            # the sibling NCE scrapers) this branch also covers non-credential
            # cases like a portal's rate limiter kicking in when several
            # scrapers log into the same account at once - a transient
            # condition that should keep retrying, not stop for the day.
            raise RuntimeError("Login appears to have failed; login form still visible.")
    else:
        print("Login form not present; assuming session is already authenticated.")

    # URL deep-links directly into fmHistoryAlarm - no extra menu click needed.
    time.sleep(10)


EXPORT_XPATH = "//button[normalize-space()='Export']"


def find_frame_with_export(timeout=WAIT_TIMEOUT):
    """Locate the (possibly nested) iframe holding the alarm table's Export
    button. The ossfacewebsite historical-alarms URL wraps the real
    eviewwebsite alarm view inside an iframe whose id is not 'fmAlarmView'
    (that id only applies to the plain eviewwebsite Current Alarms page), so
    this searches every iframe instead of relying on a fixed id."""

    def search(depth=0):
        if driver.find_elements(By.XPATH, EXPORT_XPATH):
            return True
        if depth >= 4:
            return False
        for frame in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(frame)
            except Exception:
                continue
            if search(depth + 1):
                return True
            driver.switch_to.parent_frame()
        return False

    end = time.time() + timeout
    while time.time() < end:
        driver.switch_to.default_content()
        if search():
            return True
        time.sleep(1)
    driver.switch_to.default_content()
    return False


def click_export_and_download():
    """Click Export on the already-open page (assumes the frame holding the
    Export button is already the active Selenium context) and return the
    saved file's final path. No page navigation/refresh happens here."""
    export_btn = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, EXPORT_XPATH))
    )

    before_files = set(os.listdir(DOWNLOAD_DIR))
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(3)

    all_opt = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//*[text()='All']"))
    )
    driver.execute_script("arguments[0].click();", all_opt)

    ok_btn = driver.find_element(By.XPATH, "//span[text()='OK']/ancestor::button")
    driver.execute_script("arguments[0].click();", ok_btn)

    downloaded_file = wait_for_file(DOWNLOAD_DIR, before_files, DOWNLOAD_TIMEOUT)
    if not downloaded_file:
        raise RuntimeError("Download timed out.")

    date_folder = time.strftime("%Y-%m-%d")
    dest_dir = ensure_dir(os.path.join(EXPORT_BASE_DIR, date_folder))
    ext = os.path.splitext(downloaded_file)[1]
    new_filename = f"HistoricalAlarms_MAE_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
    dest_path = os.path.join(dest_dir, new_filename)
    shutil.move(downloaded_file, dest_path)
    return dest_path


def save_error_screenshot(label):
    try:
        screenshot_path = os.path.join(
            ERROR_SCREENSHOT_DIR, f"mae_historical_{label}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        )
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved as: {screenshot_path}")
    except Exception as screenshot_error:
        print(f"Screenshot failed: {screenshot_error}")


def main():
    global driver
    try:
        login()
        print("Locating Export button (searching frames)...")
        if not find_frame_with_export():
            raise RuntimeError("Could not find the Export button in any frame.")

        print(f"Ready. Exporting every {INTERVAL_SECONDS} seconds (page refreshes each cycle).")
        while True:
            try:
                driver.switch_to.default_content()
                driver.refresh()
                wait_for_ready_state()
                time.sleep(5)

                if len(driver.find_elements(By.ID, "username")) > 0:
                    print("Session appears to have expired. Re-logging in...")
                    login()

                if not find_frame_with_export():
                    raise RuntimeError("Could not find the Export button in any frame after refresh.")

                dest_path = click_export_and_download()
                print(f"[{time.strftime('%H:%M:%S')}] Export success: {dest_path}")
            except LoginFailedError:
                raise  # let the outer handler stop the script for today instead of retrying forever
            except Exception as e:
                print(f"[{time.strftime('%H:%M:%S')}] Export cycle failed: {e}")
                if is_dead_session_error(e):
                    print("Browser/ChromeDriver process appears to have died - restarting the browser...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    try:
                        driver = create_driver()
                        print_driver_versions()
                        login()
                        if not find_frame_with_export():
                            print("Could not find the Export button after browser restart; will retry next cycle.")
                    except LoginFailedError:
                        raise
                    except Exception as restart_error:
                        print(f"Browser restart failed: {restart_error}")
                else:
                    save_error_screenshot("cycle_error")

            time.sleep(INTERVAL_SECONDS)
    except KeyboardInterrupt:
        print("Stopped by user.")
    except LoginFailedError as e:
        print(f"Login failed - stopping for today until the password is fixed: {e}")
        mark_login_failed("MAE Historical Alarms Scraper", PASSWORD)
        save_error_screenshot("login_failed")
    except Exception as e:
        print(f"Historical alarms exporter failed: {e}")
        save_error_screenshot("fatal_error")
        raise
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
