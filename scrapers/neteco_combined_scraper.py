"""NetEco Combined scraper - merges "neteco_continuous all alrams.py" (All/
Current Alarms) and neteco_historical_alarms_scraper.py (Historical Alarms)
into ONE Chrome process (two tabs, one login session), same architecture
validated live in mae_combined_scraper.py on 2026-09-14 and replicated in
nce_combined_scraper.py.

NetEco's own UI is structurally different from the MAE/NCE family (ID-based
exportBtn/allExport/confirmBtn instead of text-based Export/All/OK; no
submitDataverify login button, just btn_outerverify; Device Management ->
Current Alarms click-through required before either alarm view is
reachable), so this is adapted rather than copy-pasted, but keeps the same
two-tab/elapsed-time-scheduling/dead-session-recovery shape.

One real adaptation made here: the original login_and_navigate() always
fills the username/password fields unconditionally, assuming a fully cold
browser. That breaks for this script's Tab 2, which opens already
authenticated (shared cookies from Tab 1) and may have no login form to
fill. Added the same "check the form is actually present first" guard the
MAE/NCE family already uses, matching how an already-authenticated session
behaves there - this is the one behavior difference from the originals,
everything else (element IDs, click sequence, export flow) is unchanged.

Architecture:
    Create Chrome (one browser, two tabs)
    ├── Tab 1 (Current): login_and_navigate() via NETECO_URL, click Device
    │   Management -> Current Alarms
    └── Tab 2 (Historical): open via window.open(), login_and_navigate()
        again (already authenticated - skips credential entry, still needs
        the Device Management -> Current Alarms click-through to establish
        routing state, matching what the original standalone script does)
    Loop:
        switch to Tab 1 -> driver.get(ALL_ALARMS_URL)        -> export -> save
        switch to Tab 2 -> driver.get(HISTORICAL_ALARMS_URL) -> export -> save
        sleep(TARGET_INTERVAL - elapsed_this_cycle)

Recovery: dead-session error (same is_dead_session_error detector proven
across the MAE/NCE family) tears down and recreates the WHOLE driver and
retries only the one task that failed - a more surgical recovery than the
originals' "any exception, full restart, give up after 5" model, but
matching the recovery design agreed for all three combined scrapers.
"""
import os
import sys
import shutil
import time
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
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
USERNAME = env_str("NETECO_USERNAME")
PASSWORD = env_str("NETECO_PASSWORD")
URL = env_str("NETECO_URL")
ALL_ALARMS_URL = env_str("NETECO_ALL_ALARMS_URL")
HISTORICAL_ALARMS_URL = env_str("NETECO_HISTORICAL_ALARMS_URL")

WAIT_TIMEOUT = env_int("NETECO_WAIT_TIMEOUT", 60)
CYCLE_INTERVAL_SECONDS = env_int("NETECO_INTERVAL_SECONDS", 300)
CURRENT_DOWNLOAD_TIMEOUT_SECONDS = env_int("NETECO_DOWNLOAD_TIMEOUT_SECONDS", 180)
HISTORICAL_DOWNLOAD_TIMEOUT_SECONDS = env_int("NETECO_HISTORICAL_DOWNLOAD_TIMEOUT_SECONDS", 600)
DOWNLOAD_STABLE_SECONDS = env_int("NETECO_DOWNLOAD_STABLE_SECONDS", 5)

TEMP_DOWNLOAD_EXTENSIONS = (".crdownload", ".tmp", ".partial")

CURRENT_EXPORT_BASE_DIR = env_path_str("NETECO_EXPORT_BASE_DIR", r"C:\Current_Alarms")
HISTORICAL_EXPORT_BASE_DIR = env_path_str("NETECO_HISTORICAL_EXPORT_BASE_DIR", r"C:\Historical_Alarms")
# Shared staging dir for both tabs' raw downloads - see mae_combined_scraper.py's
# identical note on why one shared folder is safe here (no concurrent exports).
DOWNLOAD_DIR = env_path_str(
    "NETECO_ALL_DOWNLOAD_DIR",
    os.path.join(env_path_str("NETECO_DOWNLOAD_DIR", os.path.join(os.path.expanduser("~"), "Downloads")), "AllAlarms"),
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
CURRENT_EXPORT_BASE_DIR = ensure_dir(CURRENT_EXPORT_BASE_DIR)
HISTORICAL_EXPORT_BASE_DIR = ensure_dir(HISTORICAL_EXPORT_BASE_DIR)
ERROR_SCREENSHOT_DIR = ensure_dir(os.path.join(DATA_ROOT, "Errors"))

if not USERNAME or not PASSWORD or not URL or not ALL_ALARMS_URL or not HISTORICAL_ALARMS_URL:
    raise RuntimeError(
        "NETECO_USERNAME, NETECO_PASSWORD, NETECO_URL, NETECO_ALL_ALARMS_URL, and "
        "NETECO_HISTORICAL_ALARMS_URL must be configured in .env."
    )


def wait_and_fill(locator, text, timeout=WAIT_TIMEOUT):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
    try:
        element.clear()
        element.send_keys(text)
    except WebDriverException:
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
            element, text,
        )


def wait_and_click(locator, timeout=WAIT_TIMEOUT):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
    try:
        element.click()
    except WebDriverException:
        driver.execute_script("arguments[0].click();", element)
    return element


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


def wait_for_download(before_files, timeout):
    end_time = time.time() + timeout
    while time.time() < end_time:
        current_files = set(os.listdir(DOWNLOAD_DIR))
        new_files = [name for name in current_files - before_files if not name.lower().endswith(TEMP_DOWNLOAD_EXTENSIONS)]
        newest_first = sorted(
            (os.path.join(DOWNLOAD_DIR, name) for name in new_files),
            key=os.path.getmtime, reverse=True,
        )
        for path in newest_first:
            if os.path.isfile(path) and is_complete_download(path):
                return path
        time.sleep(2)
    return None


def print_driver_versions():
    caps = driver.capabilities
    browser_version = caps.get("browserVersion") or caps.get("version")
    print(f"Browser version: {browser_version}")


def build_chrome_options():
    chrome_options = Options()
    chrome_options.set_capability("acceptInsecureCerts", True)
    chrome_options.page_load_strategy = "eager"
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-insecure-localhost")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--allow-legacy-insecure-renegotiation")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--renderer-process-limit=1")
    chrome_options.add_argument("--js-flags=--max-old-space-size=384")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": False,
    })
    return chrome_options


def create_driver():
    print("Starting Chrome browser...")
    try:
        service = Service(ChromeDriverManager().install())
        new_driver = webdriver.Chrome(service=service, options=build_chrome_options())
        print(f"Using ChromeDriver at: {service.path}")
    except Exception as e:
        print(f"webdriver_manager failed ({e}); falling back to Selenium Manager")
        new_driver = webdriver.Chrome(options=build_chrome_options())
    new_driver.set_page_load_timeout(WAIT_TIMEOUT)
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


def login_and_navigate():
    """Establishes the session on Tab 1 and Tab 2 alike, verbatim from the
    original login_and_navigate() with one change: the credential-fill
    step is now conditional on a login form actually being present (see
    module docstring) - needed for Tab 2, which opens already-authenticated
    via shared cookies."""
    print(f"Navigating to: {URL}")
    driver.get(URL)
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    has_login_form = bool(driver.find_elements(By.ID, "username"))
    if has_login_form:
        print("Filling login form...")
        wait_and_fill((By.ID, "username"), USERNAME)
        wait_and_fill((By.ID, "value"), PASSWORD)
        wait_and_click((By.ID, "btn_outerverify"))
        print("Login submitted.")
        time.sleep(5)

        if driver.find_elements(By.ID, "username"):
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
    print("Reached Current Alarms (session established).")
    time.sleep(10)


def find_clickable(by, value, timeout=10):
    """Scans all matches for the first genuinely displayed+enabled one -
    same defensive lesson as the MAE/NCE combined scrapers, applied here
    even though NetEco's ID-based buttons are less likely to collide than
    MAE's text-based ones."""
    end = time.time() + timeout
    while time.time() < end:
        for elem in driver.find_elements(by, value):
            try:
                if elem.is_displayed() and elem.is_enabled():
                    return elem
            except Exception:
                continue
        time.sleep(0.5)
    raise TimeoutException(f"No visible+enabled element ({by}={value!r}) found within {timeout}s")


def click_export_button():
    driver.switch_to.default_content()
    for frame in driver.find_elements(By.TAG_NAME, "iframe"):
        driver.switch_to.default_content()
        driver.switch_to.frame(frame)
        try:
            export_container = driver.find_element(By.ID, "exportBtn")
            export_button = export_container.find_element(By.TAG_NAME, "button")
            if export_button.is_displayed() and export_button.is_enabled():
                driver.execute_script("arguments[0].click();", export_button)
                return
        except WebDriverException:
            continue

    driver.switch_to.default_content()
    export_container = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "exportBtn")))
    export_button = export_container.find_element(By.TAG_NAME, "button")
    driver.execute_script("arguments[0].click();", export_button)


def run_export_sequence(export_base_dir, filename_prefix, download_timeout, prune=False):
    before_files = set(os.listdir(DOWNLOAD_DIR))
    click_export_button()

    all_option = find_clickable(By.ID, "allExport")
    driver.execute_script("arguments[0].click();", all_option)
    ok_confirm = find_clickable(By.ID, "confirmBtn")
    driver.execute_script("arguments[0].click();", ok_confirm)

    downloaded_file = wait_for_download(before_files, download_timeout)
    if not downloaded_file:
        raise RuntimeError("Download timed out.")

    dest_dir = ensure_dir(os.path.join(export_base_dir, time.strftime("%Y-%m-%d")))
    extension = os.path.splitext(downloaded_file)[1]
    new_name = f"{filename_prefix}_{time.strftime('%Y%m%d_%H%M%S')}{extension}"
    new_path = os.path.join(dest_dir, new_name)
    shutil.move(downloaded_file, new_path)
    if prune:
        keep_only_latest_export(export_base_dir, f"{filename_prefix}_*.*", new_path)
    print(f"[{time.strftime('%H:%M:%S')}] Saved: {new_path}")


def export_current_alarms():
    driver.get(ALL_ALARMS_URL)
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") in {"interactive", "complete"}
    )
    time.sleep(3)
    run_export_sequence(
        CURRENT_EXPORT_BASE_DIR, "NetEco_All_Current_Alarm", CURRENT_DOWNLOAD_TIMEOUT_SECONDS, prune=True,
    )


def export_historical_alarms():
    driver.get(HISTORICAL_ALARMS_URL)
    WebDriverWait(driver, WAIT_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") in {"interactive", "complete"}
    )
    time.sleep(3)
    run_export_sequence(
        HISTORICAL_EXPORT_BASE_DIR, "NetEco_Historical_Alarm", HISTORICAL_DOWNLOAD_TIMEOUT_SECONDS, prune=False,
    )


def save_error_screenshot(label):
    try:
        screenshot_path = os.path.join(
            ERROR_SCREENSHOT_DIR, f"neteco_combined_{label}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        )
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved as: {screenshot_path}")
    except Exception as screenshot_error:
        print(f"Screenshot failed: {screenshot_error}")


def open_both_tabs():
    global driver, current_tab, historical_tab
    driver = create_driver()
    print_driver_versions()

    current_tab = driver.current_window_handle
    login_and_navigate()

    driver.execute_script("window.open('');")
    historical_tab = [h for h in driver.window_handles if h != current_tab][0]
    driver.switch_to.window(historical_tab)
    login_and_navigate()


def recreate_session():
    print("Browser/ChromeDriver process appears to have died - recreating session...")
    try:
        driver.quit()
    except Exception:
        pass
    open_both_tabs()


def run_step(step_name, step_fn):
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
        mark_login_failed("NetEco Combined Scraper", PASSWORD)
        save_error_screenshot("login_failed")
    except Exception as e:
        print(f"NetEco combined scraper failed: {e}")
        save_error_screenshot("fatal_error")
        raise
    finally:
        try:
            driver.quit()
        except Exception:
            pass
