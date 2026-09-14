"""NCE Combined scraper - merges nce_active_alarms_scraper.py and
nce_historical_alarms_scraper.py into ONE Chrome process (two tabs, one
login session) instead of two, same architecture validated live in
mae_combined_scraper.py on 2026-09-14 (see that session's chat for the full
design discussion and the two bugs found/fixed there).

NCE is actually SIMPLER than MAE here: both Active and Historical URLs
deep-link straight into their respective alarm views (fmAlarmView /
fmHistoryAlarm) with no extra "click X" navigation step after login - so
both tabs use the exact same login() function, just pointed at different
URLs.

Preemptively includes the two fixes MAE's pilot needed, since NCE's own
docstring says it shares the same login form/portal family as MAE/NetEco
(same submitDataverify/btn_outerverify), making the same "portal syncs
open work-tabs across windows of the same session" issue likely here too:
- find_frame_with_export requires the Export button to be
  visible+enabled, not just present in the DOM, before committing to a
  frame (a hidden Export button belonging to a synced other-tab can
  otherwise cause the recursive search to stop in the wrong frame).
- find_clickable_export_button scans all matches for the first genuinely
  displayed+enabled one, instead of trusting Selenium's default
  first-DOM-order match.

Architecture:
    Create Chrome (one browser, two tabs)
    ├── Tab 1 (Active):     login() via NCE_ACTIVE_URL
    └── Tab 2 (Historical): open via window.open(), login() via
        NCE_HISTORICAL_URL (already authenticated via shared cookies)
    Loop:
        switch to Tab 1 -> refresh -> export Active     -> save
        switch to Tab 2 -> refresh -> export Historical -> save
        sleep(TARGET_INTERVAL - elapsed_this_cycle)

Recovery: a dead-session error (is_dead_session_error, same detector
already proven across the MAE/NCE family) tears down and recreates the
WHOLE driver (both tabs die together if the browser process dies) and
retries only the one task that failed.
"""
import os
import sys
import shutil
import time
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import LoginFailedError, env_int, env_path_str, env_str, keep_only_latest_export, mark_login_failed

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ================== USER CONFIG =====================
USERNAME = env_str("NCE_USERNAME")
PASSWORD = env_str("NCE_PASSWORD")
ACTIVE_URL = env_str("NCE_ACTIVE_URL")
HISTORICAL_URL = env_str("NCE_HISTORICAL_URL")
WAIT_TIMEOUT = env_int("NCE_WAIT_TIMEOUT", 45)
HISTORICAL_DOWNLOAD_TIMEOUT = env_int("NCE_HISTORICAL_DOWNLOAD_TIMEOUT", 600)
ACTIVE_DOWNLOAD_TIMEOUT = env_int("NCE_ACTIVE_DOWNLOAD_TIMEOUT", 600)
# Both originals already use 300s independently.
CYCLE_INTERVAL_SECONDS = env_int("NCE_ACTIVE_INTERVAL_SECONDS", 300)

ACTIVE_EXPORT_BASE_DIR = env_path_str("NCE_ACTIVE_EXPORT_BASE_DIR", r"C:\NCE_Current_Alarms")
HISTORICAL_EXPORT_BASE_DIR = env_path_str("NCE_HISTORICAL_EXPORT_BASE_DIR", r"C:\NCE_Historical_Alarms")
# Shared staging dir for both tabs' raw downloads - see mae_combined_scraper.py's
# identical note on why one shared folder is safe here (no concurrent exports).
DOWNLOAD_DIR = env_path_str(
    "NCE_ACTIVE_DOWNLOAD_DIR",
    os.path.join(os.path.expanduser("~"), "Downloads", "NCE_Active"),
)

DATA_ROOT = os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data")


def ensure_dir(path):
    if not path:
        return path
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        print(f"Created directory: {path}")
    return path


DOWNLOAD_DIR = ensure_dir(DOWNLOAD_DIR)
ACTIVE_EXPORT_BASE_DIR = ensure_dir(ACTIVE_EXPORT_BASE_DIR)
HISTORICAL_EXPORT_BASE_DIR = ensure_dir(HISTORICAL_EXPORT_BASE_DIR)
ERROR_SCREENSHOT_DIR = ensure_dir(os.path.join(DATA_ROOT, "Errors"))

if not USERNAME or not PASSWORD or not ACTIVE_URL or not HISTORICAL_URL:
    raise RuntimeError(
        "NCE_USERNAME, NCE_PASSWORD, NCE_ACTIVE_URL, and NCE_HISTORICAL_URL must be configured in .env."
    )


def wait_for_file(before_files, timeout):
    end = time.time() + timeout
    while time.time() < end:
        now = set(os.listdir(DOWNLOAD_DIR))
        new_files = [f for f in (now - before_files) if not any(x in f for x in [".crdownload", ".tmp", ".partial"])]
        if new_files:
            full = [os.path.join(DOWNLOAD_DIR, f) for f in new_files]
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


def build_chrome_options():
    options = Options()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--renderer-process-limit=1")
    options.add_argument("--js-flags=--max-old-space-size=384")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    })
    return options


def create_driver():
    print("Starting Chrome browser...")
    try:
        service = Service(ChromeDriverManager().install())
        new_driver = webdriver.Chrome(service=service, options=build_chrome_options())
        print(f"Using ChromeDriver at: {service.path}")
    except Exception as e:
        print(f"webdriver_manager failed ({e}); falling back to Selenium Manager")
        new_driver = webdriver.Chrome(options=build_chrome_options())
    return new_driver


def is_dead_session_error(exc):
    message = str(exc)
    return (
        "actively refused" in message
        or "Max retries exceeded" in message
        or "invalid session id" in message.lower()
        or "session deleted" in message.lower()
        or "chrome not reachable" in message.lower()
    )


def login(url):
    """Shared by both tabs - verbatim from nce_active/nce_historical's
    login(), parameterized by URL since the two scripts were otherwise
    identical here."""
    print(f"Logging in ({url})...")
    driver.get(url)
    wait_for_ready_state()

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
            raise RuntimeError("Login appears to have failed; login form still visible.")
    else:
        print("Login form not present; assuming session is already authenticated.")

    # Both URLs deep-link straight into their alarm view - no extra menu click needed.
    time.sleep(10)


def login_active_tab():
    login(ACTIVE_URL)


def login_historical_tab():
    login(HISTORICAL_URL)


EXPORT_XPATH = "//button[normalize-space()='Export']"


def find_frame_with_export(timeout=WAIT_TIMEOUT):
    """Requires the Export match to be visible+enabled, not just present -
    see module docstring for why (lesson from the MAE pilot)."""
    def search(depth=0):
        for elem in driver.find_elements(By.XPATH, EXPORT_XPATH):
            try:
                if elem.is_displayed() and elem.is_enabled():
                    return True
            except Exception:
                continue
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


def find_clickable_export_button(timeout=WAIT_TIMEOUT):
    """Scans all matches for the first genuinely displayed+enabled one -
    see module docstring."""
    end = time.time() + timeout
    while time.time() < end:
        for elem in driver.find_elements(By.XPATH, EXPORT_XPATH):
            try:
                if elem.is_displayed() and elem.is_enabled():
                    return elem
            except Exception:
                continue
        time.sleep(0.5)
    raise TimeoutException(f"No visible+enabled 'Export' button found within {timeout}s")


def export_active_alarms():
    driver.switch_to.default_content()
    driver.refresh()
    wait_for_ready_state()
    time.sleep(5)

    if len(driver.find_elements(By.ID, "username")) > 0:
        print("Active-tab session expired. Re-logging in...")
        login_active_tab()

    if not find_frame_with_export():
        raise RuntimeError("Could not find the Export button in any frame after refresh (Active).")

    export_btn = find_clickable_export_button()
    before_files = set(os.listdir(DOWNLOAD_DIR))
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(3)

    all_opt = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='All']")))
    driver.execute_script("arguments[0].click();", all_opt)
    ok_btn = driver.find_element(By.XPATH, "//span[text()='OK']/ancestor::button")
    driver.execute_script("arguments[0].click();", ok_btn)

    downloaded_file = wait_for_file(before_files, timeout=ACTIVE_DOWNLOAD_TIMEOUT)
    if not downloaded_file:
        raise RuntimeError("Download timed out.")

    date_folder = time.strftime("%Y-%m-%d")
    dest_dir = ensure_dir(os.path.join(ACTIVE_EXPORT_BASE_DIR, date_folder))
    ext = os.path.splitext(downloaded_file)[1]
    new_filename = f"CurrentAlarms_NCE_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
    dest_path = os.path.join(dest_dir, new_filename)
    shutil.move(downloaded_file, dest_path)
    keep_only_latest_export(ACTIVE_EXPORT_BASE_DIR, "CurrentAlarms_NCE_*.*", dest_path)
    print(f"[{time.strftime('%H:%M:%S')}] Export success (Active): {dest_path}")


def export_historical_alarms():
    driver.switch_to.default_content()
    driver.refresh()
    wait_for_ready_state()
    time.sleep(5)

    if len(driver.find_elements(By.ID, "username")) > 0:
        print("Historical-tab session expired. Re-logging in...")
        login_historical_tab()

    if not find_frame_with_export():
        raise RuntimeError("Could not find the Export button in any frame after refresh (Historical).")

    export_btn = find_clickable_export_button()
    before_files = set(os.listdir(DOWNLOAD_DIR))
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(3)

    all_opt = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='All']")))
    driver.execute_script("arguments[0].click();", all_opt)
    ok_btn = driver.find_element(By.XPATH, "//span[text()='OK']/ancestor::button")
    driver.execute_script("arguments[0].click();", ok_btn)

    downloaded_file = wait_for_file(before_files, timeout=HISTORICAL_DOWNLOAD_TIMEOUT)
    if not downloaded_file:
        raise RuntimeError("Download timed out.")

    date_folder = time.strftime("%Y-%m-%d")
    dest_dir = ensure_dir(os.path.join(HISTORICAL_EXPORT_BASE_DIR, date_folder))
    ext = os.path.splitext(downloaded_file)[1]
    new_filename = f"HistoricalAlarms_NCE_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
    dest_path = os.path.join(dest_dir, new_filename)
    shutil.move(downloaded_file, dest_path)
    print(f"[{time.strftime('%H:%M:%S')}] Export success (Historical): {dest_path}")


def save_error_screenshot(label):
    try:
        screenshot_path = os.path.join(
            ERROR_SCREENSHOT_DIR, f"nce_combined_{label}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        )
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved as: {screenshot_path}")
    except Exception as screenshot_error:
        print(f"Screenshot failed: {screenshot_error}")


def open_both_tabs():
    global driver, active_tab, historical_tab
    driver = create_driver()
    print_driver_versions()

    active_tab = driver.current_window_handle
    login_active_tab()
    if not find_frame_with_export():
        raise RuntimeError("Could not find the Export button in the Active tab after initial login.")

    driver.execute_script("window.open('');")
    historical_tab = [h for h in driver.window_handles if h != active_tab][0]
    driver.switch_to.window(historical_tab)
    login_historical_tab()
    if not find_frame_with_export():
        raise RuntimeError("Could not find the Export button in the Historical tab after initial login.")


def recreate_session():
    print("Browser/ChromeDriver process appears to have died - recreating session...")
    try:
        driver.quit()
    except Exception:
        pass
    open_both_tabs()


def run_step(step_name, step_fn):
    try:
        driver.switch_to.window(active_tab if step_name == "Active" else historical_tab)
        step_fn()
    except LoginFailedError:
        raise
    except Exception as e:
        print(f"[{time.strftime('%H:%M:%S')}] {step_name} export failed: {e}")
        if is_dead_session_error(e):
            try:
                recreate_session()
                driver.switch_to.window(active_tab if step_name == "Active" else historical_tab)
                step_fn()
            except LoginFailedError:
                raise
            except Exception as retry_error:
                print(f"{step_name} export failed again after session recovery: {retry_error}")
        else:
            save_error_screenshot(f"{step_name.lower()}_cycle_error")


def main():
    open_both_tabs()
    print(f"Ready. Exporting Active + Historical every {CYCLE_INTERVAL_SECONDS} seconds (elapsed-adjusted).")

    while True:
        cycle_start = time.time()
        run_step("Active", export_active_alarms)
        run_step("Historical", export_historical_alarms)

        elapsed = time.time() - cycle_start
        sleep_time = max(0, CYCLE_INTERVAL_SECONDS - elapsed)
        print(f"Cycle took {elapsed:.1f}s, sleeping {sleep_time:.1f}s "
              f"(target cycle: {CYCLE_INTERVAL_SECONDS}s)")
        time.sleep(sleep_time)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Stopped by user.")
    except LoginFailedError as e:
        print(f"Login failed - stopping for today until the password is fixed: {e}")
        mark_login_failed("NCE Combined Scraper", PASSWORD)
        save_error_screenshot("login_failed")
    except Exception as e:
        print(f"NCE combined scraper failed: {e}")
        save_error_screenshot("fatal_error")
        raise
    finally:
        try:
            driver.quit()
        except Exception:
            pass
