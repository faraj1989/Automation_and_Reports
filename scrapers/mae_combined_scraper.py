"""MAE Combined scraper (pilot, 2026-09-14) - merges mae_scraper.py (Current
Alarms) and mae_historical_alarms_scraper.py (Historical Alarms) into ONE
Chrome process instead of two, to cut Chrome's memory footprint (each
headless Chrome instance costs ~1-1.2GB across its renderer/GPU/network/
crashpad child processes; this project runs 6 of them continuously). Both
views run as separate TABS in the same browser (same origin, same login
session/cookies) rather than one tab re-navigating back and forth, so each
tab's own page-refresh/export logic is an unmodified copy of the original
script's - the only actual behavior change is that both now share one
browser process and one download staging directory instead of two.

Architecture (see chat 2026-09-14 for the full design discussion):

    Create Chrome (one browser, two tabs)
    ├── Tab 1 (Current):  login via MAE_URL, click "Current Alarms"
    └── Tab 2 (Historical): open via window.open(), navigate to
        MAE_HISTORICAL_URL (already authenticated via shared cookies)
    Loop:
        switch to Tab 1 -> refresh -> export Current  -> save
        switch to Tab 2 -> refresh -> export Historical -> save
        sleep(TARGET_INTERVAL - elapsed_this_cycle)   # not a flat sleep -
        # keeps the Current-alarms refresh cadence close to the original
        # 5 minutes instead of drifting to 5:30+ once Historical's own
        # work time is added in.

Recovery (per tab, independent): a dead-session error (chromedriver/browser
process actually gone - see is_dead_session_error, same detector already
proven in mae_historical_alarms_scraper.py) tears down and recreates the
WHOLE driver (both tabs die together if the browser process dies) and
retries only the one task that failed - it does not also redo whichever of
Current/Historical had already succeeded earlier that same cycle.

This is a PILOT for one system only. If it holds up over many real cycles
(exports keep succeeding, RAM drops roughly in line with going from 2
Chrome processes to 1, recovery actually works when forced), the same
two-tab pattern gets replicated to NCE and NetEco.
"""
import os
import sys
import shutil
import time
import zipfile
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
USERNAME = env_str("MAE_USERNAME")
PASSWORD = env_str("MAE_PASSWORD")
CURRENT_URL = env_str("MAE_URL")
HISTORICAL_URL = env_str("MAE_HISTORICAL_URL")
WAIT_TIMEOUT = env_int("MAE_WAIT_TIMEOUT", 45)
HISTORICAL_DOWNLOAD_TIMEOUT = env_int("MAE_HISTORICAL_DOWNLOAD_TIMEOUT", 600)
# Both originals already use 300s independently - the combined cycle target
# is just that same value, not a compromise between two different numbers.
CYCLE_INTERVAL_SECONDS = env_int("MAE_INTERVAL_SECONDS", 300)

CURRENT_EXPORT_BASE_DIR = env_path_str("MAE_EXPORT_BASE_DIR", r"C:\Current_Alarms")
HISTORICAL_EXPORT_BASE_DIR = env_path_str("MAE_HISTORICAL_EXPORT_BASE_DIR", r"C:\Historical_Alarms")
# One shared staging directory for both tabs' raw downloads - Chrome's
# download.default_directory is a browser-level preference, not per-tab, so
# one driver can only have one. Current/Historical exports never run
# concurrently in this script, so there's no risk of mixing up which
# downloaded file belongs to which step; each step still moves its file to
# its own correct *_EXPORT_BASE_DIR immediately after download.
DOWNLOAD_DIR = env_path_str("MAE_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads"))

DATA_ROOT = os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data")
ERROR_SCREENSHOT_DIR = None  # set by ensure_dir below


def ensure_dir(path):
    if not path:
        return path
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        print(f"Created directory: {path}")
    return path


DOWNLOAD_DIR = ensure_dir(DOWNLOAD_DIR)
CURRENT_EXPORT_BASE_DIR = ensure_dir(CURRENT_EXPORT_BASE_DIR)
HISTORICAL_EXPORT_BASE_DIR = ensure_dir(HISTORICAL_EXPORT_BASE_DIR)
ERROR_SCREENSHOT_DIR = ensure_dir(os.path.join(DATA_ROOT, "Errors"))

if not USERNAME or not PASSWORD or not CURRENT_URL or not HISTORICAL_URL:
    raise RuntimeError(
        "MAE_USERNAME, MAE_PASSWORD, MAE_URL, and MAE_HISTORICAL_URL must be configured in .env."
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
    """Same detector already proven in mae_historical_alarms_scraper.py /
    nce_active_alarms_scraper.py / nce_historical_alarms_scraper.py."""
    message = str(exc)
    return (
        "actively refused" in message
        or "Max retries exceeded" in message
        or "invalid session id" in message.lower()
        or "session deleted" in message.lower()
        or "chrome not reachable" in message.lower()
    )


def handle_login_form_if_present():
    """Shared by both tabs' login flows - checked once at whatever page is
    currently loaded in the active tab. Verbatim logic from both original
    scripts (they were already near-identical here)."""
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "username")))
        has_login_form = True
    except Exception:
        has_login_form = False

    if not has_login_form:
        print("Login form not present; assuming session is already authenticated.")
        return

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
        # Checked before the credential keywords: MAE's login page carries
        # the word "credentials" in its own static boilerplate even on a
        # rate-limit rejection - see mae_scraper.py's identical comment.
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


def login_current_tab():
    """Tab 1: full login + the extra 'click Current Alarms' nav step,
    verbatim from mae_scraper.py's login_and_navigate()."""
    print("Logging in (Current Alarms tab)...")
    driver.get(CURRENT_URL)
    wait_for_ready_state()
    handle_login_form_if_present()

    print("Navigating to Current Alarms...")
    current_alarms_xpath = (
        "//div[contains(text(), 'Current Alarms')] | "
        "//span[contains(text(), 'Current Alarms')] | "
        "//a[contains(text(), 'Current Alarms')] | "
        "//*[contains(text(), 'Current Alarms')]"
    )
    try:
        current_alarm = WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, current_alarms_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", current_alarm)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", current_alarm)
        print("Current Alarms clicked")
    except Exception as e:
        print(f"Could not find Current Alarms: {e}")
        elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Current Alarms')]")
        for elem in elements:
            if elem.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", elem)
                print("Current Alarms found and clicked via text search")
                break
    time.sleep(10)


def login_historical_tab():
    """Tab 2: URL deep-links straight into fmHistoryAlarm - no extra menu
    click, verbatim from mae_historical_alarms_scraper.py's login(). Opened
    after Tab 1 has already authenticated, so this normally just hits the
    'already authenticated' branch (shared cookies, same origin)."""
    print("Navigating Historical Alarms tab...")
    driver.get(HISTORICAL_URL)
    wait_for_ready_state()
    handle_login_form_if_present()
    time.sleep(10)


EXPORT_XPATH = "//button[normalize-space()='Export']"


def find_clickable_export_button(timeout=WAIT_TIMEOUT):
    """Root cause found live 2026-09-14 pilot run: this portal syncs its
    open in-app work-tabs across browser windows of the same session (via
    shared localStorage, confirmed by screenshot - Tab 2 showed BOTH
    'Current Alarms' and 'Historical Alarms' as open tabs, even though only
    Historical was ever navigated to in that window). That leaves more than
    one element matching EXPORT_XPATH in the DOM at once (one per open
    work-tab), and Selenium's EC.element_to_be_clickable/find_element
    (singular) always resolves to the first DOM-order match regardless of
    which tab is actually visible - if that happens to be a hidden tab's
    button, the wait hangs for the full timeout every cycle. This scans all
    matches and returns the first one that's actually displayed+enabled,
    which is what the single-tab original scripts got "for free" by only
    ever having exactly one Export button in the DOM."""
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


def find_frame_with_export(timeout=WAIT_TIMEOUT):
    """Recursive iframe search (up to depth 4), adapted from
    mae_historical_alarms_scraper.py for the Historical tab whose iframe
    has no fixed id. Diverges from the original in one way, found live in
    this pilot: the original only checked DOM *presence* of an Export
    button to decide a frame was "the" frame, which is safe when there's
    only ever one Export button in the whole page (the original's
    single-tab world). Here, the portal syncs its open in-app work-tabs
    across browser windows of the same session, so this window's DOM can
    contain a second, permanently-hidden Export button belonging to the
    synced 'Current Alarms' tab in a sibling iframe - presence-only search
    can commit to that iframe first and never reach the real, visible one.
    Requiring visibility+enabled here is what actually fixes it."""
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


def export_current_alarms():
    """Tab 1 export: fixed fmAlarmView iframe id, prune-to-latest output -
    verbatim from mae_scraper.py's main loop body."""
    driver.refresh()
    time.sleep(12)

    if "login" in driver.current_url.lower() or len(driver.find_elements(By.ID, "username")) > 0:
        print("Current-tab session expired during refresh. Re-logging in...")
        login_current_tab()

    driver.switch_to.default_content()
    iframe_xpath = "//iframe[contains(@id,'fmAlarmView')]"
    WebDriverWait(driver, WAIT_TIMEOUT).until(EC.presence_of_element_located((By.XPATH, iframe_xpath)))
    driver.switch_to.frame(driver.find_element(By.XPATH, iframe_xpath))

    print("Finding Export button (Current)...")
    export_btn = find_clickable_export_button(timeout=20)

    before_files = set(os.listdir(DOWNLOAD_DIR))
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(3)

    all_opt = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[text()='All']")))
    driver.execute_script("arguments[0].click();", all_opt)
    ok_btn = driver.find_element(By.XPATH, "//span[text()='OK']/ancestor::button")
    driver.execute_script("arguments[0].click();", ok_btn)

    downloaded_file = wait_for_file(before_files, timeout=180)
    if not downloaded_file:
        raise RuntimeError("Download timed out.")

    date_folder = time.strftime("%Y-%m-%d")
    dest_dir = ensure_dir(os.path.join(CURRENT_EXPORT_BASE_DIR, date_folder))
    new_filename = f"CurrentAlarms_MAE_{time.strftime('%Y%m%d_%H%M%S')}.csv"
    dest_path = os.path.join(dest_dir, new_filename)

    if zipfile.is_zipfile(downloaded_file):
        with zipfile.ZipFile(downloaded_file) as zf:
            csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
            if not csv_names:
                raise RuntimeError(f"MAE export zip had no CSV inside: {zf.namelist()}")
            with zf.open(csv_names[0]) as src, open(dest_path, "wb") as dst:
                shutil.copyfileobj(src, dst)
        os.remove(downloaded_file)
    else:
        shutil.move(downloaded_file, dest_path)

    keep_only_latest_export(CURRENT_EXPORT_BASE_DIR, "CurrentAlarms_MAE_*.csv", dest_path)
    print(f"[{time.strftime('%H:%M:%S')}] Export success (Current): {new_filename}")


def export_historical_alarms():
    """Tab 2 export: recursive frame search, keeps every export (no
    pruning) - verbatim from mae_historical_alarms_scraper.py's
    click_export_and_download()."""
    driver.switch_to.default_content()
    driver.refresh()
    wait_for_ready_state()
    time.sleep(5)

    if len(driver.find_elements(By.ID, "username")) > 0:
        print("Historical-tab session expired. Re-logging in...")
        login_historical_tab()

    if not find_frame_with_export():
        raise RuntimeError("Could not find the Export button in any frame after refresh.")

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
    new_filename = f"HistoricalAlarms_MAE_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
    dest_path = os.path.join(dest_dir, new_filename)
    shutil.move(downloaded_file, dest_path)
    print(f"[{time.strftime('%H:%M:%S')}] Export success (Historical): {dest_path}")


def save_error_screenshot(label):
    try:
        screenshot_path = os.path.join(
            ERROR_SCREENSHOT_DIR, f"mae_combined_{label}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        )
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved as: {screenshot_path}")
    except Exception as screenshot_error:
        print(f"Screenshot failed: {screenshot_error}")


def open_both_tabs():
    """Fresh browser, fresh login on Tab 1, Tab 2 opened+navigated after
    (inherits Tab 1's session cookies - same origin, no separate login)."""
    global driver, current_tab, historical_tab
    driver = create_driver()
    print_driver_versions()

    current_tab = driver.current_window_handle
    login_current_tab()

    driver.execute_script("window.open('');")
    historical_tab = [h for h in driver.window_handles if h != current_tab][0]
    driver.switch_to.window(historical_tab)
    login_historical_tab()
    if not find_frame_with_export():
        raise RuntimeError("Could not find the Export button in the Historical tab after initial login.")


def recreate_session():
    """Recovery: dead session means the whole browser process is gone, so
    both tabs are rebuilt together - see module docstring."""
    print("Browser/ChromeDriver process appears to have died - recreating session...")
    try:
        driver.quit()
    except Exception:
        pass
    open_both_tabs()


def run_step(step_name, step_fn):
    """Runs one export step; on a dead-session error, recreates the whole
    session and retries only this one step once (per the user's recovery
    design: 'retry current task', not the whole cycle)."""
    try:
        driver.switch_to.window(current_tab if step_name == "Current" else historical_tab)
        step_fn()
    except LoginFailedError:
        raise
    except Exception as e:
        print(f"[{time.strftime('%H:%M:%S')}] {step_name} export failed: {e}")
        if is_dead_session_error(e):
            try:
                recreate_session()
                driver.switch_to.window(current_tab if step_name == "Current" else historical_tab)
                step_fn()
            except LoginFailedError:
                raise
            except Exception as retry_error:
                print(f"{step_name} export failed again after session recovery: {retry_error}")
        else:
            save_error_screenshot(f"{step_name.lower()}_cycle_error")


def main():
    open_both_tabs()
    print(f"Ready. Exporting Current + Historical every {CYCLE_INTERVAL_SECONDS} seconds (elapsed-adjusted).")

    while True:
        cycle_start = time.time()
        run_step("Current", export_current_alarms)
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
        mark_login_failed("MAE Combined Scraper", PASSWORD)
        save_error_screenshot("login_failed")
    except Exception as e:
        print(f"MAE combined scraper failed: {e}")
        save_error_screenshot("fatal_error")
        raise
    finally:
        try:
            driver.quit()
        except Exception:
            pass
