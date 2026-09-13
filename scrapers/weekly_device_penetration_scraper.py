import argparse
import glob
import os
import shutil
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT_DIR))

from project_config import env_int, env_path_str, env_str, load_env_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

load_env_file()

# --- Portal Core Configuration ---
# Same SmartCare portal/login as the SmartCare CEM scraper - only the target
# dashboard (Device Penetration Rate instead of MBB Traffic Analysis) differs.
# Falls back to the SMARTCARE_* credentials/URLs so this works out of the box
# without duplicating login config in .env.
LOGIN_URL = env_str("WEEKLY_DEVICE_PENETRATION_LOGIN_URL", env_str("SMARTCARE_LOGIN_URL"))
TARGET_DASHBOARD_URL = env_str("WEEKLY_DEVICE_PENETRATION_TARGET_DASHBOARD_URL")
EXPORT_TASK_URL = env_str(
    "WEEKLY_DEVICE_PENETRATION_EXPORT_TASK_URL",
    env_str("SMARTCARE_EXPORT_TASK_URL"),
)
USERNAME = env_str("WEEKLY_DEVICE_PENETRATION_USERNAME", env_str("SMARTCARE_USERNAME"))
PASSWORD = env_str("WEEKLY_DEVICE_PENETRATION_PASSWORD", env_str("SMARTCARE_PASSWORD"))

DOWNLOAD_DIR = env_path_str("WEEKLY_DEVICE_PENETRATION_DOWNLOAD_DIR", str(Path.home() / "Downloads" / "Weekly_Device_Penetration"))
OUTPUT_DIR = env_path_str("WEEKLY_DEVICE_PENETRATION_OUTPUT_DIR", str(Path.home() / "Downloads" / "Weekly_Device_Penetration_Exports"))
EXPORT_TASK_TIMEOUT = env_int("WEEKLY_DEVICE_PENETRATION_EXPORT_TASK_TIMEOUT", 300)
POLL_INTERVAL = env_int("WEEKLY_DEVICE_PENETRATION_POLL_INTERVAL", 15)


def init_driver():
    """Initializes Chrome and explicitly allows downloads to the target folder."""
    chrome_options = Options()
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-insecure-localhost")
    #chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        print(f"webdriver_manager failed ({e}); falling back to Selenium Manager")
        driver = webdriver.Chrome(options=chrome_options)

    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": DOWNLOAD_DIR},
    )
    return driver


def snapshot_download_dir(directory, patterns=None):
    """Capture file mtimes so a re-download with the same name is still detectable."""
    patterns = patterns or ("*.xlsx", "*.xls", "*.csv", "*.zip")
    snapshot = {}
    for pattern in patterns:
        for path in glob.glob(os.path.join(directory, pattern)):
            try:
                snapshot[path] = {
                    "mtime": os.path.getmtime(path),
                    "size": os.path.getsize(path),
                }
            except OSError:
                continue
    return snapshot


def wait_for_new_download(directory, snapshot_before_click, timeout=120, expected_task_name=None):
    """Wait for a new or updated file to finish downloading."""
    print(f"Watching download folder for new file: {directory}")
    deadline = time.time() + timeout
    expected_task_name = (expected_task_name or "").strip().lower()

    while time.time() < deadline:
        partials = glob.glob(os.path.join(directory, "*.crdownload"))
        current_snapshot = snapshot_download_dir(directory)
        candidates = []

        for path, meta in current_snapshot.items():
            previous = snapshot_before_click.get(path)
            is_new = previous is None
            is_updated = previous is not None and meta["mtime"] > previous["mtime"] + 0.5
            if not (is_new or is_updated):
                continue

            filename = os.path.basename(path).lower()
            score = 0
            if expected_task_name and expected_task_name in filename:
                score += 10
            if path.lower().endswith(".xlsx"):
                score += 3
            if is_new:
                score += 2
            candidates.append((score, meta["mtime"], path))

        if candidates and not partials:
            _, _, newest = max(candidates, key=lambda item: (item[0], item[1]))
            prev_size = -1
            stable_count = 0
            while time.time() < deadline:
                try:
                    size = os.path.getsize(newest)
                except OSError:
                    time.sleep(2)
                    continue

                if size == prev_size and size > 0:
                    stable_count += 1
                    if stable_count >= 3:
                        print(f"Download complete: {newest} ({size:,} bytes)")
                        return newest
                else:
                    stable_count = 0

                prev_size = size
                time.sleep(2)

        time.sleep(3)

    print("[WARN] Timed out waiting for download file to appear.")
    return None


def parse_portal_datetime(value):
    """Parse portal datetime strings with multiple formats."""
    value = (value or "").strip()
    # Handle HTML entities
    value = value.replace("&nbsp;", " ")
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def open_task_manager_window(driver):
    """Open the async export page in a separate Chrome window in the same session."""
    # Get the original window handle
    original_window = driver.current_window_handle

    # Open a new window
    driver.switch_to.default_content()
    driver.execute_script("window.open('');")

    # Switch to the new window
    new_window = [handle for handle in driver.window_handles if handle != original_window][0]
    driver.switch_to.window(new_window)

    print(f"Opened separate Chrome window for Async Export Task Manager:\n{EXPORT_TASK_URL}")
    driver.get(EXPORT_TASK_URL)

    # Wait for the page to load
    print("Waiting for Async Export page to load...")
    time.sleep(15)

    # Try to handle any popups or authentication dialogs
    try:
        driver.switch_to.alert.dismiss()
    except:
        pass

    return driver


def find_and_click_query_button(driver):
    """Find and click the Query button with improved iframe handling."""
    print("Looking for Query button...")

    # First, try to find if there's an iframe containing the toolbar
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"Found {len(iframes)} iframe(s) on the page")

    query_button = None
    found_in_iframe = False

    # Check each iframe for the Query button
    for idx, iframe in enumerate(iframes):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            time.sleep(1)

            # Look for the Query button in this iframe
            selectors = [
                "#sweet-1005 > button",
                "button.sweet-form-button-highlight",
                "//button[@class='sweet-form-button-highlight']",
                "//button[contains(@class, 'sweet-form-button-highlight')]",
                "//button[.//div[contains(text(), 'Query')]]",
                "//div[@id='sweet-1005']//button",
                "//div[contains(@class, 'sweet-form-button-formerText')][text()='Query']/..",
            ]

            for selector in selectors:
                try:
                    if selector.startswith("#") or selector.startswith("."):
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    else:
                        elements = driver.find_elements(By.XPATH, selector)

                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            query_button = element
                            found_in_iframe = True
                            print(f"Found Query button in iframe {idx} with selector: {selector}")
                            break
                    if query_button:
                        break
                except:
                    continue

            if query_button:
                break

        except Exception as e:
            print(f"Error checking iframe {idx}: {e}")
            continue

    # If not found in iframes, check main page
    if not query_button:
        driver.switch_to.default_content()
        selectors = [
            "#sweet-1005 > button",
            "button.sweet-form-button-highlight",
            "//button[@class='sweet-form-button-highlight']",
            "//button[contains(@class, 'sweet-form-button-highlight')]",
            "//button[.//div[contains(text(), 'Query')]]",
        ]

        for selector in selectors:
            try:
                if selector.startswith("#") or selector.startswith("."):
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                else:
                    elements = driver.find_elements(By.XPATH, selector)

                for element in elements:
                    if element.is_displayed() and element.is_enabled():
                        query_button = element
                        print(f"Found Query button on main page with selector: {selector}")
                        break
                if query_button:
                    break
            except:
                continue

    # If still not found, try JavaScript
    if not query_button:
        print("Trying JavaScript to find Query button...")
        try:
            js_script = """
            // Try to find the Query button
            var button = null;

            // Look for button with sweet-form-button-highlight class
            var buttons = document.querySelectorAll('button.sweet-form-button-highlight');
            for (var i = 0; i < buttons.length; i++) {
                if (buttons[i].textContent.includes('Query')) {
                    button = buttons[i];
                    break;
                }
            }

            // If not found, look for any button with Query text
            if (!button) {
                buttons = document.querySelectorAll('button');
                for (var i = 0; i < buttons.length; i++) {
                    if (buttons[i].textContent.includes('Query')) {
                        button = buttons[i];
                        break;
                    }
                }
            }

            return button;
            """

            element = driver.execute_script(js_script)
            if element:
                query_button = element
                print("Found Query button via JavaScript")
        except Exception as e:
            print(f"JavaScript search failed: {e}")

    if query_button:
        try:
            print("Clicking Query button...")
            driver.execute_script("arguments[0].scrollIntoView(true);", query_button)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", query_button)
            print("Query button clicked successfully!")
            return True
        except Exception as e:
            print(f"Error clicking Query button: {e}")
            return False
    else:
        print("ERROR: Could not find Query button on the page")
        return False


def wait_for_task_completion(driver, task_name, timeout=300):
    """Wait for a specific task to complete."""
    print(f"Waiting for task '{task_name}' to complete...")
    deadline = time.time() + timeout

    while time.time() < deadline:
        try:
            # Find the task row
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.sweet-grid-content-tr")
            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 5:
                        task_name_cell = cells[3] if len(cells) > 3 else None
                        status_cell = cells[4] if len(cells) > 4 else None

                        if task_name_cell:
                            row_task_name = task_name_cell.text.strip()
                            if task_name in row_task_name:
                                status = status_cell.text.strip() if status_cell else ""
                                print(f"Task status: {status}")

                                if "Completed" in status:
                                    print(f"Task '{task_name}' is complete!")
                                    return True
                                elif "Failed" in status or "Error" in status:
                                    print(f"Task '{task_name}' failed with status: {status}")
                                    return False
                                else:
                                    print(f"Task still {status}, waiting...")
                                    time.sleep(10)
                                    break
                except:
                    continue
        except Exception as e:
            print(f"Error checking task status: {e}")

        # Click Query button to refresh the list
        print("Refreshing task list...")
        find_and_click_query_button(driver)
        time.sleep(5)

    print(f"Timeout waiting for task '{task_name}' to complete")
    return False


def extract_task_rows_from_sweet_grid(driver):
    """Extract task rows from the Sweet Grid table."""
    task_rows = []

    # The table uses sweet-grid-content-tr class
    rows = driver.find_elements(By.CSS_SELECTOR, "tr.sweet-grid-content-tr")

    if not rows:
        # Try alternative selector
        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'sweet-grid-content-tr')]")

    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 10:
                # Based on the HTML structure:
                # col 0: checkbox
                # col 1: SN
                # col 2: Name
                # col 3: Task Name (clickable)
                # col 4: Task Status
                # col 5: Task Submit Time
                # col 6: Task Start Time
                # col 7: Task Complete Time
                # col 8: Size
                # col 9: Operation

                task_name_elem = cells[3] if len(cells) > 3 else None
                task_name = ""
                if task_name_elem:
                    try:
                        # Try to get the text from the span
                        name_span = task_name_elem.find_element(By.CSS_SELECTOR, "span")
                        task_name = name_span.text.strip()
                    except:
                        task_name = task_name_elem.text.strip()

                status = cells[4].text.strip() if len(cells) > 4 else ""
                submit_time = cells[5].text.strip() if len(cells) > 5 else ""
                start_time = cells[6].text.strip() if len(cells) > 6 else ""
                complete_time = cells[7].text.strip() if len(cells) > 7 else ""

                # Parse the submit time
                submit_time_clean = submit_time.replace("&nbsp;", " ").strip()
                created_at = parse_portal_datetime(submit_time_clean)

                task_rows.append({
                    'row': row,
                    'cells': cells,
                    'task_name': task_name,
                    'status': status,
                    'submit_time': submit_time,
                    'start_time': start_time,
                    'complete_time': complete_time,
                    'created_at': created_at,
                    'task_name_cell': task_name_elem
                })

        except Exception as e:
            print(f"Error extracting row: {e}")
            continue

    return task_rows


def find_latest_task(driver):
    """Find the latest task from the list."""
    task_rows = extract_task_rows_from_sweet_grid(driver)

    if not task_rows:
        return None

    # Find the latest task by create time
    latest_task = None
    for task in task_rows:
        if task.get('created_at'):
            if latest_task is None or task['created_at'] > latest_task['created_at']:
                latest_task = task

    # If no time could be parsed, just use the first one
    if latest_task is None and task_rows:
        latest_task = task_rows[0]

    return latest_task


def find_task_by_name(driver, target_name):
    """Search for a task by name across iframes and the main document and return its row dict.

    Returns the first matching task dict from extract_task_rows_from_sweet_grid or None.
    """
    target = (target_name or "").strip()
    if not target:
        return None

    # Try current context first
    try:
        rows = extract_task_rows_from_sweet_grid(driver)
        for r in rows:
            if target in (r.get('task_name') or ""):
                return r
    except Exception:
        pass

    # Try inside each iframe
    try:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for idx, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)
                rows = extract_task_rows_from_sweet_grid(driver)
                for r in rows:
                    if target in (r.get('task_name') or ""):
                        return r
            except Exception:
                continue
    except Exception:
        pass

    # Finally try main document again
    try:
        driver.switch_to.default_content()
        rows = extract_task_rows_from_sweet_grid(driver)
        for r in rows:
            if target in (r.get('task_name') or ""):
                return r
    except Exception:
        pass

    return None


def click_task_to_download(driver, latest_task):
    """Click on the task name to trigger download."""
    try:
        # Get the task name cell
        task_name_cell = latest_task['task_name_cell']

        if task_name_cell:
            # Try to find the clickable span or link
            clickable_elements = task_name_cell.find_elements(By.XPATH,
                                                              ".//span[@style='cursor:pointer;color:blue;text-decoration:underline;'] | "
                                                              ".//span[contains(@style, 'cursor')] | "
                                                              ".//a | "
                                                              ".//*[contains(@class, 'link')]")

            if clickable_elements:
                element_to_click = clickable_elements[0]
            else:
                element_to_click = task_name_cell

            # Scroll to element and click
            driver.execute_script("arguments[0].scrollIntoView(true);", element_to_click)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", element_to_click)
            print(f"Clicked latest task: {latest_task['task_name']}")
            return True
        else:
            # Try clicking the row itself
            driver.execute_script("arguments[0].scrollIntoView(true);", latest_task['row'])
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", latest_task['row'])
            print("Clicked the task row as fallback")
            return True

    except Exception as e:
        print(f"[ERROR] Could not click the latest task: {e}")
        return False


def find_and_download_task(driver, wait, timeout=EXPORT_TASK_TIMEOUT):
    """Find and download the latest export task from the Async Export page."""
    print(f"\n--- STEP 9: Async Export Task Manager ---")

    # Reset to default content first
    driver.switch_to.default_content()
    time.sleep(3)

    # Find the task list (might be in an iframe)
    print("Looking for task list...")
    task_rows = []

    # Check if we need to be inside an iframe
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    if iframes:
        for idx, iframe in enumerate(iframes):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)
                time.sleep(2)

                rows = extract_task_rows_from_sweet_grid(driver)
                if rows:
                    task_rows = rows
                    print(f"Found {len(task_rows)} tasks in iframe {idx}")
                    break
            except Exception as e:
                print(f"Error checking iframe {idx}: {e}")
                continue

    # If no tasks found in iframes, check main page
    if not task_rows:
        driver.switch_to.default_content()
        task_rows = extract_task_rows_from_sweet_grid(driver)

    if not task_rows:
        print("[WARN] No export tasks were found. Clicking Query button to load tasks...")
        find_and_click_query_button(driver)
        time.sleep(5)

        # Try again
        task_rows = extract_task_rows_from_sweet_grid(driver)

    if not task_rows:
        print("[ERROR] No tasks found after refresh")
        return False

    print(f"Found {len(task_rows)} task row(s) on Async Export page.")

    # Find the latest task
    latest_task = find_latest_task(driver)

    if latest_task is None:
        print("[ERROR] Could not identify the latest task")
        return False

    print(
        f"Latest task: name='{latest_task['task_name']}', "
        f"submit_time='{latest_task.get('submit_time', 'N/A')}', "
        f"status='{latest_task.get('status', 'N/A')}'"
    )

    task_name = latest_task['task_name']
    status = latest_task.get('status', '')

    # If task is not completed, wait for it
    if "Completed" not in status:
        print(f"Task '{task_name}' is not completed yet (status: {status})")
        print("Will monitor and wait for completion...")

        # Wait for task to complete
        completed = wait_for_task_completion(driver, task_name, timeout=300)

        if not completed:
            print(f"[ERROR] Task '{task_name}' did not complete within timeout")
            return False

        # Refresh the task list after completion
        print("Refreshing task list after completion...")
        find_and_click_query_button(driver)
        time.sleep(5)

        # Re-find the specific task by name (fresh element references)
        latest_task = find_task_by_name(driver, task_name)
        if latest_task is None:
            # Fallback: try to grab the latest task by timestamp
            latest_task = find_latest_task(driver)
            if latest_task is None:
                print("[ERROR] Could not find task after refresh")
                return False

    # Now click the task to download
    pre_existing = snapshot_download_dir(DOWNLOAD_DIR)

    if click_task_to_download(driver, latest_task):
        # Wait for the download
        downloaded_path = wait_for_new_download(
            DOWNLOAD_DIR,
            pre_existing,
            timeout=120,
            expected_task_name=latest_task['task_name'],
        )

        if downloaded_path:
            print(f"\nDownloaded file: {downloaded_path}")
            return downloaded_path

        print("[WARN] Task was clicked, but no downloaded file was detected.")
        return False
    else:
        print("[ERROR] Failed to click the task")
        return False


HISTORICAL_FILENAME = "Weekly_Device_Penetration_Historical.xlsx"
# Confirmed against a real export: Time, Device Model, Device Brand, Device
# Type, Device OS, Device Technology, Number of Users(count), Penetration
# Rate(%) - one row per device model per weekly snapshot Time.
HISTORY_DEDUPE_COLUMNS = ["Time", "Device Model", "Device Brand", "Device Type", "Device OS", "Device Technology"]


def safe_extract_zip(archive: zipfile.ZipFile, destination: Path) -> None:
    """Extract only when every member stays within destination - guards
    against a malicious/corrupt zip using ../ path traversal."""
    destination = destination.resolve()
    for member in archive.infolist():
        target = (destination / member.filename).resolve()
        if target != destination and destination not in target.parents:
            raise ValueError(f"Unsafe ZIP member path: {member.filename}")
    archive.extractall(destination)


def extract_if_zipped(path: Path) -> Path:
    """The portal sometimes delivers this export as a .zip (confirmed - a
    real run left one alongside a plain .xlsx). If it is one, extract its
    data file next to it and return that; otherwise return path unchanged."""
    if path.suffix.lower() != ".zip":
        return path
    extract_dir = path.parent / f"{path.stem}_extracted"
    extract_dir.mkdir(exist_ok=True)
    with zipfile.ZipFile(path) as archive:
        safe_extract_zip(archive, extract_dir)
    data_files = [p for p in extract_dir.iterdir() if p.suffix.lower() in (".xlsx", ".xls", ".csv")]
    if not data_files:
        raise ValueError(f"No .xlsx/.xls/.csv found inside {path}")
    extracted = data_files[0]
    final_path = path.parent / extracted.name
    if final_path.exists():
        final_path.unlink()
    shutil.move(str(extracted), str(final_path))
    extract_dir.rmdir()
    path.unlink()  # the zip's content now lives in final_path - redundant once extracted
    print(f"Extracted {path.name} -> {final_path.name}")
    return final_path


def fold_into_history(data_file: Path, output_dir: Path) -> Path:
    """Append this export's rows into one running historical workbook,
    deduped (by snapshot Time + device identity) so re-processing the same
    export never double-counts."""
    new_data = pd.read_excel(data_file) if data_file.suffix.lower() in (".xlsx", ".xls") else pd.read_csv(data_file)
    history_path = Path(output_dir) / HISTORICAL_FILENAME
    if history_path.exists():
        combined = pd.concat([pd.read_excel(history_path), new_data], ignore_index=True)
    else:
        combined = new_data
    dedupe_columns = [c for c in HISTORY_DEDUPE_COLUMNS if c in combined.columns] or list(combined.columns)
    combined = combined.drop_duplicates(subset=dedupe_columns, keep="last")
    combined.to_excel(history_path, index=False)
    print(f"Historical dataset updated: {history_path} ({len(combined)} total rows)")
    return history_path


def cleanup_old_raw_downloads(output_dir: Path, keep: Path) -> None:
    """Keep only the most recent raw export (for audit) plus the historical
    workbook - delete every older Device_Penetration_Rate_* file so raw
    downloads don't accumulate week over week."""
    keep_resolved = keep.resolve()
    for old in Path(output_dir).glob("Device_Penetration_Rate_*"):
        if old.resolve() == keep_resolved:
            continue
        try:
            old.unlink()
            print(f"Deleted old raw export: {old}")
        except OSError as exc:
            print(f"[WARN] Could not delete old raw export {old}: {exc}")


def parse_args():
    parser = argparse.ArgumentParser(description="Export Weekly Device Penetration data for the analysis pipeline")
    parser.add_argument("--download-dir", default=DOWNLOAD_DIR)
    parser.add_argument("--output-dir", default=OUTPUT_DIR)
    parser.add_argument("--login-attempts", type=int, default=1,
                        help="Accepted for wrapper compatibility; portal retries are handled by Selenium.")
    return parser.parse_args()


def weekly_device_penetration_automation(download_dir=None, output_dir=None):
    """Main automation function."""
    global DOWNLOAD_DIR, OUTPUT_DIR
    DOWNLOAD_DIR = str(Path(download_dir or DOWNLOAD_DIR).expanduser().resolve())
    OUTPUT_DIR = str(Path(output_dir or OUTPUT_DIR).expanduser().resolve())
    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    driver = init_driver()
    wait = WebDriverWait(driver, 45)

    try:
        print(f"Connecting to login page: {LOGIN_URL}")
        driver.get(LOGIN_URL)

        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        username_field.clear()
        username_field.send_keys(USERNAME)

        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(PASSWORD)

        print("Submitting login credentials...")
        driver.find_element(By.ID, "loginButton").click()
        wait.until(EC.url_contains("homepage.html"))
        print("Base authentication successful.")

        print("Routing directly to Device Penetration Rate section...")
        driver.get(TARGET_DASHBOARD_URL)
        time.sleep(10)

        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            frame_index = 1 if len(iframes) >= 2 else 0
            print(f"Switching context to iframe index: {frame_index}")
            driver.switch_to.frame(iframes[frame_index])
            time.sleep(3)

        print("Locating the 'Report Details' tab element...")
        report_details_tab = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'showTab')][@title='Report Details']"))
        )
        report_details_tab.click()
        time.sleep(5)

        print("Locating and clicking 'Query' button...")
        query_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[@title='Query'] | //button[contains(., 'Query')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", query_button)
        time.sleep(1)
        query_button.click()

        print("Query dispatched! Waiting for data loader spinner...")
        try:
            WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.el-loading-mask")))
            print("Loader spinner active. Server compiling rows...")
            WebDriverWait(driver, 120).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.el-loading-mask"))
            )
            print("Loader mask hidden.")
        except Exception:
            print("Loader layout passed or skipped.")

        print("Awaiting table rows to confirm full data render...")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.el-table__body-wrapper tr.el-table__row"))
        )
        print("Data rows confirmed. Stabilising...")
        time.sleep(10)

        print("Locating 'Export' icon button...")
        export_icon = wait.until(
            EC.presence_of_element_located((By.XPATH, "//i[@title='Export' or contains(@class, 'icon-export')]"))
        )
        driver.execute_script("arguments[0].click();", export_icon)
        time.sleep(3)

        print("Selecting 'Excel(All Data)' from dropdown...")
        excel_all_option = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//li[contains(@class,'el-dropdown-menu__item')][contains(text(),'Excel(All Data)')]")
            )
        )
        driver.execute_script("arguments[0].click();", excel_all_option)
        print("Export task successfully requested!")
        time.sleep(10)

        print("Opening a separate Chrome window for Async Export Task Manager...")
        driver = open_task_manager_window(driver)

        # Create a new wait object for the new window
        wait = WebDriverWait(driver, 45)

        # Find and download the task
        downloaded_path = find_and_download_task(driver, wait)

        if downloaded_path:
            source = Path(downloaded_path)
            destination = Path(OUTPUT_DIR) / source.name
            if source.resolve() != destination.resolve():
                if destination.exists():
                    destination = destination.with_name(
                        f"{destination.stem}_{datetime.now():%Y%m%d_%H%M%S}{destination.suffix}"
                    )
                shutil.move(str(source), str(destination))
            print(f"\nAutomation complete. File moved to: {destination}")

            try:
                data_file = extract_if_zipped(destination)
                fold_into_history(data_file, OUTPUT_DIR)
                cleanup_old_raw_downloads(OUTPUT_DIR, keep=data_file)
            except Exception as exc:
                print(f"[WARN] Post-download processing (unzip/history/cleanup) failed: {exc}")
        else:
            print("\n[WARN] Automation finished but the file may not have downloaded. Check the browser window.")
            print("[INFO] Holding browser open for 5 minutes for manual inspection...")
            time.sleep(300)

    except Exception as e:
        print(f"\n[CRITICAL ERROR] Automation sequence broken: {e}")
        import traceback
        traceback.print_exc()
        print("[INFO] Holding browser open for 5 minutes for manual inspection...")
        time.sleep(300)

    finally:
        print("Closing browser session.")
        driver.quit()


if __name__ == "__main__":
    args = parse_args()
    weekly_device_penetration_automation(args.download_dir, args.output_dir)
