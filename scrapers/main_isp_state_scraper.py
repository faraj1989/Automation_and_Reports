"""Continuous Main ISP State exporter (NCE U2000 Performance Instance).

Same NCE portal and login form as scrapers/nce_active_alarms_scraper.py,
but a completely different app within it: opening "U2000" from the portal
launcher loads Huawei's network manager as a Webswing-streamed Java desktop
application - the UI is a live bitmap painted onto a <canvas> element, not a
normal web page. Most of it has no DOM to select against, so navigating it
means real mouse events at pixel coordinates instead of clicking named
elements. This flow was captured with Chrome DevTools Recorder (two passes,
walked through live with the user) rather than read off existing code, since
this app had no scraper before.

Flow: login -> dismiss post-login warning dialog -> open U2000 -> Performance
Instance -> Group Management tab -> select "ISP" in the tree (pixel clicks) ->
select all -> right-click -> View Historical Data -> Table tab -> More ->
Save All -> CSV -> OK -> file lands in DOWNLOAD_DIR via Chrome's normal
download mechanism, same as every other scraper here.

FRAGILITY WARNING: the tree-selection step (find_isp_canvas, the click/
double-click/focus/select-all/right-click sequence) has no stable selector -
there is no search/filter field in this app (confirmed with the user). It
replays fixed pixel coordinates against a locked window size. If Huawei ever
reorders the "Group Management" tree (a new group added above ISP, etc.) or
the window size drifts, this breaks silently - a wrong/empty export, not an
exception. If that happens, re-record the flow. Two tools for this, both
confirmed available in this project (2026-09-06):
  - Chrome DevTools Recorder (F12 -> Recorder) - the original tool this file
    was built with. Exports JSON with click offsets relative to each element,
    which is what canvas_action()'s x/y pixel coordinates come from.
  - `npx playwright codegen <url>` (Playwright is in this project's venv) -
    opens a live browser plus an Inspector window that generates code as you
    click, with a session-storage-backed login so credentials aren't needed
    on every re-record. Worth trying first since it can pick up real
    selectors even inside parts of the app that turn out not to be the
    legacy Webswing canvas (see the 2026-09-06 investigation notes for this
    file - some report pages here are modern real-DOM pages, not canvas).
Either way, record at the exact locked window size above (WINDOW_WIDTH x
WINDOW_HEIGHT) - coordinates captured at a different size are not
transferable - and update ISP_TREE_CLICKS/ISP_LIST_FOCUS/ISP_CONTEXT_MENU_ITEM
below.

There is no deep link into a Webswing app session - unlike the alarm
scrapers, which just refresh their page and re-click Export, every cycle
here must repeat the full click sequence from the portal loading page.

RECURRING POPUP (confirmed live 2026-09-15): NCE intermittently re-surfaces
an "Emergency Maintenance Notification" alarm-reminder dialog on top of the
canvas, seemingly on a periodic nag timer independent of what the scraper is
doing (observed ~2 minutes apart across two separate fresh sessions, and
reproduced by the user manually clicking through the exact same flow live).
When it lands at the same moment as the right-click below, it steals the
click and the context menu never opens, which then fails downstream at
"Could not find the frame containing the Table tab." It renders inside the
Webswing canvas bitmap, so there's no DOM id to dismiss it by - navigate_to_
isp_table() now sends Escape and retries the right-click/View Historical
Data/Table-tab sequence once before giving up. The more permanent fix is on
the NCE side: open the dialog and uncheck "Pop up when notification
received" (or resolve/ack the underlying alarms) so it stops recurring for
this account. A Playwright codegen recording done live with the user the
same day confirmed the rest of the flow works end-to-end once past this
popup - it captured a real 2,873-row CSV download - and surfaced more robust
aria-label/role-based locators (e.g. aria-label="页签项,Table" for the Table
tab) that are now used as the first attempt before the originally recorded
ids, which are kept as fallbacks.
"""
import os
import sys
import shutil
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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
USERNAME = env_str("NCE_USERNAME")
PASSWORD = env_str("NCE_PASSWORD")
URL = env_str("MAIN_ISP_URL")
DOWNLOAD_DIR = env_path_str(
    "MAIN_ISP_DOWNLOAD_DIR",
    os.path.join(os.path.expanduser("~"), "Downloads", "Main_ISP_State"),
)
EXPORT_BASE_DIR = env_path_str("MAIN_ISP_EXPORT_BASE_DIR", r"C:\Main_ISP_State")
WAIT_TIMEOUT = env_int("MAIN_ISP_WAIT_TIMEOUT", 45)
DOWNLOAD_TIMEOUT = env_int("MAIN_ISP_DOWNLOAD_TIMEOUT", 600)
# This flow is far heavier per cycle than the alarm scrapers (full login +
# ~15 sequential UI steps vs. one Export click) - default interval is longer.
INTERVAL_SECONDS = env_int("MAIN_ISP_INTERVAL_SECONDS", 900)
# Locked to the viewport the flow was recorded at - every pixel coordinate
# below is only valid at this exact window size. 1437x877 matches the
# recording of a confirmed, verified-successful manual run (2026-08-30) -
# earlier recordings used 1365x911 and are superseded.
WINDOW_WIDTH, WINDOW_HEIGHT = 1437, 877

# Direct deep link into U2000's Webswing session - previously used in place
# of the portal launcher tile click below, but confirmed live on 2026-09-06
# to no longer reliably reach U2000 (it silently falls back to the portal's
# own default view instead). Kept only as a fallback attempt in
# open_u2000_app(); the tile click is primary now. The fragment after
# "#page=" is base64 for
# "Action=com.huawei.u2000.unitedmgr.topo.action.DoWebTopoAction" - U2000's
# topology view. No session token embedded, so it's safe to hardcode/reuse
# across logins (host is still per-deployment, so it comes from env).
U2000_TOPO_URL = env_str("MAIN_ISP_TOPO_URL")

# Portal home page with the app-tile launcher. Re-confirmed working via a
# fresh Chrome DevTools Recorder pass on 2026-09-06: navigate here, then
# click the "Network Management" (U2000) tile.
PORTAL_HOME_URL = env_str("MAIN_ISP_PORTAL_URL")

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
    raise RuntimeError("NCE_USERNAME and NCE_PASSWORD must be configured in .env or environment variables.")


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


chrome_options = Options()
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument(f"--window-size={WINDOW_WIDTH},{WINDOW_HEIGHT}")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
})

print("Initializing ChromeDriver...")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.set_window_size(WINDOW_WIDTH, WINDOW_HEIGHT)
print(f"Using ChromeDriver at: {service.path}")
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
        except Exception as e_click:
            print(f"Login click failed: {e_click}")

        print("Waiting for login to complete...")
        time.sleep(8)

        if len(driver.find_elements(By.ID, "username")) > 0:
            page_text = driver.page_source.lower()
            login_error_keywords = [
                "invalid username", "invalid password", "incorrect username", "incorrect password",
                "wrong username", "wrong password", "password expired", "expired password",
                "credentials", "login failed", "account locked",
            ]
            if any(keyword in page_text for keyword in login_error_keywords):
                raise LoginFailedError("Login failed: invalid credentials or expired password detected.")
            # Deliberately NOT a LoginFailedError: confirmed live 2026-09-07 this
            # branch also covers non-credential cases like NCE's rate limiter
            # ("Operations are too frequent. Please try later.") when several
            # scrapers log into the same account at once - a transient
            # condition that should keep retrying, not stop for the day.
            raise RuntimeError("Login appears to have failed; login form still visible.")
    else:
        print("Login form not present; assuming session is already authenticated.")

    # Post-login warning/disclaimer dialog - not present on every login (e.g.
    # only once per day on some sessions), so this is best-effort, not a hard
    # requirement.
    try:
        ok_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "login_warn_confirm"))
        )
        driver.execute_script("arguments[0].click();", ok_button)
        print("Dismissed post-login warning dialog.")
    except Exception:
        print("No post-login warning dialog appeared (or already dismissed).")

    time.sleep(3)


def _u2000_loaded(timeout=20):
    """True once a real Webswing canvas is present - the only reliable
    signal that U2000 actually loaded. Earlier this checked for the Monitor
    ribbon button's DOM id instead, but that id turned out to be present
    (and clickable) even on the portal's own default page, which is
    presumably a shared id in the portal's menu-registration markup - so it
    passed the check while still being on the wrong page entirely (confirmed
    live 2026-09-06). A canvas can't lie the same way."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[id^='wrapper-'] canvas"))
        )
        return True
    except Exception:
        return False


def _wait_for_u2000_ready(timeout=60):
    """Wait for U2000's own "Loading... NN% ... Please wait" splash dialog to
    clear instead of a fixed sleep. This used to poll page_source for
    U2000's "System loaded successfully" status-bar message, but that text
    is painted onto the Webswing canvas bitmap (confirmed live 2026-09-10
    via a full DOM/shadow-DOM dump) - it can never appear in page_source, so
    that wait always silently ran out its full timeout every cycle before
    continuing anyway. The Monitor menu item becoming clickable is a real
    DOM signal that U2000 has registered its own menu into the shared
    top-nav shell and is ready to be driven."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[title="Monitor"]'))
        )
    except Exception:
        print("Timed out waiting for U2000's Monitor menu to become clickable - continuing anyway.")
    time.sleep(1)  # let the just-registered menu items settle


def open_u2000_app():
    """Navigate to the portal home page and click the "Network Management"
    (U2000) app tile. A direct deep link into U2000's Webswing session
    (U2000_TOPO_URL) was used here previously and was more deterministic
    when it worked, but was confirmed live on 2026-09-06 to no longer
    reliably reach U2000 - it silently falls back to the portal's default
    view instead, and clicking pixel coordinates meant for U2000 against
    that wrong page is what produced a confusing low-level chromedriver
    error rather than an obvious one. The tile click below was re-confirmed
    working the same day via a fresh Chrome DevTools Recorder pass. The old
    deep link is kept as a fallback attempt in case the tile ever becomes
    unreliable again (see module docstring's FRAGILITY WARNING)."""
    print("Opening U2000 app (portal tile)...")
    driver.get(PORTAL_HOME_URL)
    wait_for_ready_state()
    time.sleep(2)

    tile_selectors = [
        (By.CSS_SELECTOR, "div.brick_container_U2020-F_U2000_App div.appComNameContainer"),
        (By.XPATH, '//*[@id="appComContainer_U2020-F_U2000_App"]/div[2]'),
    ]
    clicked = False
    handles_before = set(driver.window_handles)
    for by, selector in tile_selectors:
        try:
            elem = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((by, selector)))
            driver.execute_script("arguments[0].click();", elem)
            clicked = True
            break
        except Exception:
            continue

    if clicked:
        # The tile opens U2000 in a new popup window, not the current tab -
        # confirmed 2026-09-06 via a Playwright codegen recording
        # (`page.waitForEvent('popup')` around the tile click). Every earlier
        # failure investigating this was the driver continuing to act on the
        # original portal tab, which never navigates anywhere by itself once
        # the tile is clicked.
        try:
            WebDriverWait(driver, 15).until(lambda d: len(d.window_handles) > len(handles_before))
            new_handle = (set(driver.window_handles) - handles_before).pop()
            driver.switch_to.window(new_handle)
        except Exception:
            print("No new window/tab appeared after the tile click.")
        _wait_for_u2000_ready()

    if not clicked or not _u2000_loaded():
        print("Tile click didn't reach U2000 - falling back to the direct deep link...")
        driver.get(U2000_TOPO_URL)
        wait_for_ready_state()
        _wait_for_u2000_ready()

    if not _u2000_loaded():
        save_error_screenshot("u2000_not_loaded")
        raise RuntimeError(
            "U2000 app did not load (neither the portal tile nor the direct deep "
            "link produced a Webswing canvas). See the u2000_not_loaded screenshot."
        )


def find_webswing_canvas():
    """The main Webswing viewport canvas. Its wrapper id (wrapper-<number>)
    looked session-generated across recordings, so match structurally
    instead of hardcoding the exact id."""
    return WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[id^='wrapper-'] canvas:nth-of-type(2)"))
    )


def find_context_menu_canvas():
    # Bumped from 10s to 20s - confirmed live 2026-09-15 this can time out
    # (with an unmessaged, blank-looking TimeoutException) specifically on a
    # right-click retry that follows a popup-interrupted first attempt; the
    # context menu canvas was just slower to appear that time, not absent.
    return WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#webswing-root-container canvas:nth-of-type(3)"))
    )


def canvas_action(canvas, x, y, kind="click"):
    """Real mouse events at a pixel offset from the canvas's top-left corner
    - matches DOM offsetX/offsetY semantics, which is what Chrome DevTools
    Recorder captured. No JS click here: the remote Java app behind Webswing
    needs genuine mouse coordinates, not a synthetic DOM click."""

    def build():
        actions = ActionChains(driver).move_to_element_with_offset(canvas, x, y)
        if kind == "double":
            actions.double_click()
        elif kind == "right":
            actions.context_click()
        else:
            actions.click()
        return actions

    try:
        build().perform()
    except Exception:
        # Confirmed live 2026-09-15: a raw, empty-message WebDriverException
        # from chromedriver on the second of two rapid right-clicks against
        # this canvas (the browser/page were otherwise healthy immediately
        # after - a fresh screenshot showed no blocking overlay), consistent
        # with a transient CDP hiccup rather than a real click failure. One
        # retry after a short pause clears it.
        time.sleep(1)
        build().perform()


def find_frame_containing(by, value, timeout=WAIT_TIMEOUT, max_depth=4):
    """Recursively search every iframe for one containing the given element
    - same defensive approach as scrapers/nce_active_alarms_scraper.py's
    find_frame_with_export(), needed because the "Table" view and everything
    after it (More, Save All, format dropdowns, OK) live inside a nested
    iframe whose depth/position isn't safe to hardcode."""

    def search(depth=0):
        if driver.find_elements(by, value):
            return True
        if depth >= max_depth:
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


def click_by_id(element_id, timeout=WAIT_TIMEOUT):
    elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, element_id)))
    driver.execute_script("arguments[0].click();", elem)
    return elem


def click_by_xpath(xpath, timeout=WAIT_TIMEOUT):
    elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    driver.execute_script("arguments[0].click();", elem)
    return elem


def navigate_to_isp_table():
    """Performance Instance -> select ISP -> View Historical Data -> Table
    tab. Matches a confirmed, verified-successful manual run recorded on
    2026-08-30 at the locked 1437x877 window size - simpler than earlier
    attempts (no separate Group Management tab click or tree drill-down
    turned out to be needed; a single click on the canvas plus Ctrl+A is
    enough). Everything up to and including the right-click is pixel-
    coordinate against the canvas (see module docstring); everything from
    the Table tab onward has real element ids."""
    driver.switch_to.default_content()

    # Clicking Monitor IS required here - a full DOM/shadow-DOM dump
    # confirmed live 2026-09-10 that it's properly scoped to U2000
    # (id="refr.mm.U2020-F_U2000_App.U2020-F_Monitor"), not the shared
    # portal shell. The earlier "don't click Monitor" belief came from
    # clicking Performance Instance via a fixed id (#action-0-2) that turned
    # out to be reused across multiple menu entries (Alarm Logs,
    # Performance Instance, ...) rendered by the same dropdown - whichever
    # one happened to render first in the DOM silently won the click, which
    # looked like "Monitor routes elsewhere" but was really just an
    # ambiguous selector. Clicking by each item's own visible text instead
    # of that shared id resolves it correctly every time.
    print("Opening Monitor menu...")
    click_by_xpath('//span[@title="Monitor"]')
    time.sleep(1)

    print("Clicking Performance Instance...")
    click_by_xpath('//*[normalize-space(text())="Performance Instance"]')
    time.sleep(3)

    print("Selecting ISP and all its objects...")
    canvas = find_webswing_canvas()
    canvas_action(canvas, 418, 199, "click")
    time.sleep(1)
    ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
    time.sleep(1)

    # Screenshot right before the riskiest step - if ISP wasn't actually
    # selected, this is the evidence needed to diagnose a wrong/empty export
    # after the fact rather than just seeing a bad CSV with no explanation.
    save_error_screenshot("before_right_click", is_error=False)

    # Up to three attempts: see the module docstring's RECURRING POPUP note -
    # an NCE alarm-reminder dialog intermittently steals the right-click
    # below, which then fails downstream trying to find the Table tab frame.
    # Confirmed live 2026-09-15 that the very next attempt right after a
    # popup-interrupted one can itself run slower (the context menu canvas
    # taking longer to reappear) and also fail, which is why this is 3
    # attempts and not 2. The selection made above (ISP + Ctrl+A) lives in
    # the app, not in this loop, so retrying just this part doesn't need to
    # redo it.
    last_error = None
    for attempt in (1, 2, 3):
        try:
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.5)
            except Exception:
                pass

            print(f"Right-clicking to open context menu... (attempt {attempt})")
            canvas = find_webswing_canvas()
            canvas_action(canvas, 408, 206, "right")
            time.sleep(1.5)

            print("Clicking 'View Historical Data'...")
            context_canvas = find_context_menu_canvas()
            canvas_action(context_canvas, 164, 119, "click")
            time.sleep(4)  # new tab/panel needs time to load inside its iframe

            print("Switching into the report frame and clicking Table tab...")
            # Prefer the aria-label locator confirmed live via a Playwright
            # codegen recording on 2026-09-15 (Huawei's own accessibility
            # label, "tab item: Table") - it's independent of the generated
            # numeric id, which isn't guaranteed stable across
            # sessions/versions. Falls back to the originally recorded id.
            table_tab_xpath = '//*[@aria-label="页签项,Table"]'
            if find_frame_containing(By.XPATH, table_tab_xpath, timeout=10) or \
                    find_frame_containing(By.ID, "ev_tabItem_10011"):
                try:
                    click_by_xpath(table_tab_xpath, timeout=5)
                except Exception:
                    click_by_id("ev_tabItem_10011")
                time.sleep(2)
                return

            last_error = RuntimeError("Could not find the frame containing the Table tab.")
        except Exception as e:
            # Confirmed live 2026-09-15: an earlier version of this loop only
            # caught the clean "not found" outcome above, so a raised
            # exception here (e.g. find_context_menu_canvas() timing out)
            # escaped the loop entirely and aborted the whole cycle after
            # just 2 of the intended 3 attempts. Every failure mode needs to
            # land back in this loop to actually get retried.
            last_error = e

        save_error_screenshot(f"table_tab_not_found_attempt{attempt}")
        driver.switch_to.default_content()

    raise last_error


def export_csv():
    """More -> Save All -> CSV -> OK. All real, stable element ids - the
    reliable half of this flow. Assumes the current frame (set by
    navigate_to_isp_table) is still active."""
    print("Opening More menu...")
    click_by_xpath('//*[@id="MoreDropDownDropDownButton"]')
    time.sleep(1)

    print("Clicking Save All...")
    try:
        click_by_xpath('//*[@id="ev_popup_10000"]/div/div/span[5]')
    except Exception:
        # aria-label confirmed live via a 2026-09-15 Playwright codegen recording.
        click_by_xpath('//*[@aria-label="选项5,Save All"]')
    time.sleep(2)

    print("Selecting CSV file type...")
    click_by_id("ev_slt_10002")
    time.sleep(1)
    try:
        click_by_xpath('//*[@id="ev_slt_10002_pop"]/div/div/span[1]')
    except Exception:
        try:
            # aria-label confirmed live via a 2026-09-15 Playwright codegen recording.
            click_by_xpath('//*[@aria-label="选项1,CSV files (*.csv)"]')
        except Exception:
            # Fall back to matching by visible text if the popup's item order/ids shift.
            click_by_xpath('//span[contains(text(), "CSV files")]')
    time.sleep(1)

    # Present in the verified working recording (clicking the encoding
    # dropdown without changing its value, already UTF-8 by default) -
    # purpose not fully understood, kept as a non-fatal best-effort step to
    # match that recording exactly rather than risk deviating from it.
    try:
        click_by_id("ev_slt_10003", timeout=5)
        time.sleep(1)
    except Exception:
        print("UTF-8 encoding dropdown not found/clickable - continuing (default is already UTF-8).")

    print("Clicking OK...")
    # Dialog button ids vary by instance (ev_bt_1006/ev_bt_1007 both seen
    # across recordings) - try the ids seen so far, then fall back to
    # matching by visible text.
    for button_id in ("ev_bt_1006", "ev_bt_1007"):
        try:
            click_by_id(button_id, timeout=5)
            break
        except Exception:
            continue
    else:
        try:
            # aria-label confirmed live via a 2026-09-15 Playwright codegen recording.
            click_by_xpath('//*[@aria-label="OK"]')
        except Exception:
            click_by_xpath("//span[text()='OK']/ancestor::button")


def save_error_screenshot(label, is_error=True):
    try:
        prefix = "main_isp" if is_error else "main_isp_debug"
        screenshot_path = os.path.join(
            ERROR_SCREENSHOT_DIR, f"{prefix}_{label}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        )
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved as: {screenshot_path}")
    except Exception as screenshot_error:
        print(f"Screenshot failed: {screenshot_error}")


def run_export_cycle():
    """Full cycle: fresh login through file download. There is no deep link
    or "just re-click Export" shortcut for a Webswing session, so every
    cycle repeats the whole sequence from the portal loading page."""
    login()
    open_u2000_app()
    navigate_to_isp_table()

    before_files = set(os.listdir(DOWNLOAD_DIR))
    export_csv()

    downloaded_file = wait_for_file(DOWNLOAD_DIR, before_files, DOWNLOAD_TIMEOUT)
    if not downloaded_file:
        raise RuntimeError("Download timed out.")

    date_folder = time.strftime("%Y-%m-%d")
    dest_dir = ensure_dir(os.path.join(EXPORT_BASE_DIR, date_folder))
    ext = os.path.splitext(downloaded_file)[1]
    new_filename = f"MainISPState_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
    dest_path = os.path.join(dest_dir, new_filename)
    shutil.move(downloaded_file, dest_path)
    return dest_path


def main():
    print(f"Main ISP State exporter starting. Window locked to {WINDOW_WIDTH}x{WINDOW_HEIGHT}. "
          f"Exporting every {INTERVAL_SECONDS} seconds.")
    try:
        while True:
            try:
                dest_path = run_export_cycle()
                print(f"[{time.strftime('%H:%M:%S')}] Export success: {dest_path}")
            except LoginFailedError:
                raise  # let the outer handler stop the script for today instead of retrying forever
            except Exception as e:
                print(f"[{time.strftime('%H:%M:%S')}] Export cycle failed: {e}")
                save_error_screenshot("cycle_error")

            time.sleep(INTERVAL_SECONDS)
    except KeyboardInterrupt:
        print("Stopped by user.")
    except LoginFailedError as e:
        print(f"Login failed - stopping for today until the password is fixed: {e}")
        mark_login_failed("Main ISP State")
        save_error_screenshot("login_failed")
    except Exception as e:
        print(f"Main ISP State exporter failed: {e}")
        save_error_screenshot("fatal_error")
        raise
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
