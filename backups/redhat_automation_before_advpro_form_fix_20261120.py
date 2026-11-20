"""
Red Hat Exam Scheduler + Adv Pro  ->  "redhat exam 2026.xlsx"

Steps
  1. Open Chrome (Advantage pro profile), Red Hat Exam Scheduler > menu >
     Exam Deliveries. Status = Completed, Time Zone = Asia/Calcutta,
     Effective Date From/To = the day, Location = 2093, then 2337, then 2394.
  2. For every candidate: name, email, contact number, exam code (first 5
     characters, e.g. EX200) and "retake" if the eligibility code says RETAKE,
     otherwise "y".
  3. Open Adv Pro (vetri), Master > Registration > Display, filter by
     PresentPhoneNo and search every phone number: found = "y" (our student),
     not found = "n".
  4. Back up the Excel file, add the rows to the month sheet, save.
  5. Kiosk table at the top of the month sheet (Date | 2093 | 2337 | 2394 | exam number):
     every exam of the day at each location (Completed + Scheduled search, also
     No Show), and the total. Empty cells only.

Usage
  python redhat_automation.py                    # asks for the date (Enter = today)
  python redhat_automation.py --date 2026-10-08  # a specific day
  python redhat_automation.py --skip-advpro      # don't check Adv Pro ("Our Student" left blank)
  python redhat_automation.py --dry-run          # don't save the Excel file
"""
import argparse
import configparser
import datetime as dt
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import warnings
from copy import copy
from pathlib import Path

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

HERE = Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.ini"
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
UPPER, LOWER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"


# ---------------------------------------------------------------- helpers

def log(msg):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def ask_user(msg):
    # In the nightly run nobody can do what is asked, and the Enter key presses fed
    # from a file would make the script carry on as if it had been done.
    if not sys.stdin or not sys.stdin.isatty():
        raise RuntimeError(f"Needs a person at the computer: {msg}")
    input(f"\n>>> {msg}\n    Press Enter when done... ")


def parse_date(text):
    """'08-10-2026', '8/10/2026', '8.10.26', '2026-10-08' or '8-10' (this year). Day comes first."""
    text = text.strip()
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", text):
        return dt.date.fromisoformat("-".join(p.zfill(2) for p in text.split("-")))
    parts = re.split(r"[-/.\s]+", text)
    if not 2 <= len(parts) <= 3 or not all(p.isdigit() for p in parts):
        raise ValueError(text)
    day, month = int(parts[0]), int(parts[1])
    year = int(parts[2]) if len(parts) == 3 else dt.date.today().year
    if year < 100:
        year += 2000
    return dt.date(year, month, day)


def ask_date():
    today = dt.date.today()
    while True:
        text = input(f"\n>>> Which date? Type it as DD-MM-YYYY (e.g. {today:%d-%m-%Y}),\n"
                     f"    or just press Enter for today ({today:%d-%b-%Y}): ").strip(" \t﻿")
        if not text:
            return today
        try:
            date = parse_date(text)
        except ValueError:
            print("    Not a valid date - try again, e.g. 07-10-2026")
            continue
        print(f"    Using {date:%A, %d-%b-%Y}")
        return date


def load_config():
    cfg = configparser.RawConfigParser(inline_comment_prefixes=(";",))
    cfg.read(CONFIG_PATH, encoding="utf-8")
    return cfg


def norm_name(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def clean_cell_text(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def phone10(value):
    digits = re.sub(r"\D", "", str(value or ""))
    return digits[-10:] if len(digits) >= 10 else ""


def clean_phone(raw):
    """Indian mobile -> int like the existing sheet; anything else stays as text."""
    raw = clean_cell_text(raw)
    if not raw:
        return None
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    if len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]
    if len(digits) == 10 and digits[0] in "6789":
        return int(digits)
    return raw


def is_blank(value):
    return value is None or str(value).strip() == ""


def exam_code(text):
    """'EX200V10K: Red Hat Certified System Administrator' -> 'EX200'."""
    text = clean_cell_text(text)
    m = re.match(r"[A-Za-z]{2}\d{3}", text)
    return (m.group(0) if m else text[:5]).upper()


# ---------------------------------------------------------------- chrome

def port_open(port):
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


def shortcut_profile(name):
    """Read --profile-directory from a Desktop shortcut like 'Advantage pro - Chrome.lnk'."""
    for desktop in (Path.home() / "Desktop", Path(r"C:\Users\Public\Desktop")):
        lnk = desktop / name
        if lnk.exists():
            for attempt in range(3):  # PowerShell sometimes comes back empty; the next try works
                args = subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     f"(New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}').Arguments"],
                    capture_output=True, text=True).stdout
                m = re.search(r'--profile-directory=(?:"([^"]+)"|(\S+))', args)
                if m:
                    return m.group(1) or m.group(2)
                time.sleep(2)
            raise RuntimeError(f"No --profile-directory in shortcut {lnk}")
    raise RuntimeError(f"Shortcut not found on the Desktop: {name}")


def chrome_running():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq chrome.exe"], capture_output=True, text=True).stdout
    return "chrome.exe" in out.lower()


# keep Chrome from freezing/throttling tabs in the background
CHROME_FLAGS = ["--disable-background-timer-throttling", "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows",
                "--disable-features=IntensiveWakeUpThrottling,CalculateNativeWinOcclusion"]


def chrome_pages(port):
    """Tabs of the debug-port Chrome, straight from Chrome (never hangs on a frozen tab)."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=10) as r:
            return [t for t in json.load(r) if t.get("type") == "page"]
    except Exception:
        return []


def activate_tab(port, tab_id):
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/json/activate/{tab_id}", timeout=10).close()
    except Exception:
        pass


def close_tab(port, tab_id):
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/json/close/{tab_id}", timeout=10).close()
    except Exception:
        pass


def wake_tabs(port):
    """Chrome freezes background tabs and ChromeDriver hangs on a frozen tab;
    bringing each tab to the front for a moment wakes it up."""
    for t in chrome_pages(port):  # every tab: a frozen chrome://newtab hangs it too
        activate_tab(port, t["id"])
        time.sleep(0.3)


_DRIVERS = []  # ChromeDriver sessions opened by this program (stopped by step_cleanup)


def try_attach(port, target_id, wait):
    """One try at connecting ChromeDriver -> (driver, None) or (None, why it failed)."""
    opts = Options()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    service = Service()
    result = {}

    def connect():
        try:
            driver = webdriver.Chrome(service=service, options=opts)
            driver.switch_to.window(target_id)
            result["driver"] = driver
        except Exception as e:
            result["error"] = f"{type(e).__name__}: {(str(e).splitlines() or [''])[0][:120]}"

    thread = threading.Thread(target=connect, daemon=True)
    thread.start()
    thread.join(wait)
    if "driver" in result:
        _DRIVERS.append(result["driver"])
        return result["driver"], None
    process = getattr(service, "process", None)
    if process:
        try:
            process.kill()  # also frees the thread if it is stuck
        except Exception:
            pass
    return None, result.get("error") or f"no answer in {wait} seconds"


def attach_driver(port, target_id, tries=3, wait=60):
    """ChromeDriver can hang on a frozen tab and Selenium would only give up after
    120 seconds, failing the step - so give up sooner, wake the tabs and try again."""
    for attempt in range(1, tries + 1):
        driver, why = try_attach(port, target_id, wait)
        if driver:
            return driver
        log(f"Could not connect to Chrome (try {attempt} of {tries}): {why}")
        if attempt < tries:
            wake_tabs(port)
            activate_tab(port, target_id)
            time.sleep(2)
    raise RuntimeError(f"Could not connect to Chrome after {tries} tries: {why}")


def open_chrome(cfg, section, start_url, match, close_old=None):
    """Open start_url in the Chrome profile of the section's Desktop shortcut and
    return the Selenium driver switched to that tab. close_old: every other tab
    with this text in its address is closed, so only the new tab is left.

    Chrome is started with a debug port so the script can control it. That only
    works if Chrome was fully closed before the first launch.
    """
    port = cfg.getint("chrome", "debug_port")
    profile = shortcut_profile(cfg.get(section, "shortcut"))
    if not port_open(port) and chrome_running():
        # Chrome often keeps running in the background after its windows are closed
        answer = input("\n>>> Chrome is already running (maybe in the background).\n"
                       "    The script must restart it. Close Chrome now? Open tabs come back\n"
                       "    with Ctrl+Shift+T. [Y/n]: ").strip().lower()
        if answer in ("", "y", "yes"):
            subprocess.run(["taskkill", "/IM", "chrome.exe", "/F"], capture_output=True)
            for _ in range(20):
                if not chrome_running():
                    break
                time.sleep(0.5)
        if chrome_running():
            sys.exit("Chrome is still running - close it (check the system tray) and run again.")
        log("Chrome closed")
    before = {t["id"] for t in chrome_pages(port)}
    log(f"Opening Chrome profile '{profile}' ({cfg.get(section, 'shortcut')})")
    subprocess.Popen([cfg.get("chrome", "path"), f"--profile-directory={profile}",
                      f"--remote-debugging-port={port}", *CHROME_FLAGS, start_url])
    for _ in range(60):
        if port_open(port):
            break
        time.sleep(0.5)
    else:
        raise RuntimeError(f"Chrome did not open the control port {port}")

    # find the tab that just opened start_url (newest first) without touching the others
    target = None
    start = time.time()
    end = start + 30
    while time.time() < end and not target:
        pages = [t for t in chrome_pages(port) if match in t["url"]]
        new = [t for t in pages if t["id"] not in before]
        # give the new tab time to show its address before settling for an old one
        target = (new or (pages if time.time() - start > 10 else []) or [None])[0]
        if not target:
            time.sleep(1)
    if not target:
        ask_user(f"Could not find the new Chrome tab ({match}). Click on it in Chrome.")
        target = next((t for t in chrome_pages(port) if match in t["url"]), None)
        if not target:
            raise RuntimeError(f"No Chrome tab with {match}")

    if close_old:
        old = [t for t in chrome_pages(port) if close_old in t["url"] and t["id"] != target["id"]]
        for t in old:
            close_tab(port, t["id"])
        log(f"Closed {len(old)} old tab(s) of {close_old}")

    log("Waking up Chrome tabs")
    wake_tabs(port)
    activate_tab(port, target["id"])
    driver = attach_driver(port, target["id"])
    log("Connected to Chrome")
    return driver


def visible(el):
    try:
        return el.is_displayed()
    except Exception:
        return False


def click(driver, el):
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


def has_text(driver, text):
    xp = f"//*[normalize-space(translate(., '{UPPER}', '{LOWER}'))='{text.lower()}']"
    return any(visible(e) for e in driver.find_elements(By.XPATH, xp))


def click_text(driver, *texts, timeout=10):
    """Click the visible element whose whole text equals one of `texts` (case-insensitive)."""
    end = time.time() + timeout
    while True:
        for text in texts:
            xp = f"//*[normalize-space(translate(., '{UPPER}', '{LOWER}'))='{text.lower()}']"
            els = [e for e in driver.find_elements(By.XPATH, xp) if visible(e)]
            if els:
                click(driver, els[-1])  # last = innermost element
                return True
        if time.time() > end:
            return False
        time.sleep(0.5)


def wait_for(driver, css, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        els = [e for e in driver.find_elements(By.CSS_SELECTOR, css) if visible(e)]
        if els:
            return els[0]
        time.sleep(0.5)
    return None


def need(driver, css, label, timeout=20):
    """wait_for, but a missing box is an error that says what is missing."""
    el = wait_for(driver, css, timeout=timeout)
    if not el:
        raise RuntimeError(f"{label} box not found on the page: {driver.current_url}")
    return el


def type_into(el, value):
    el.click()
    el.send_keys(Keys.CONTROL, "a")
    el.send_keys(Keys.DELETE)
    el.send_keys(value)


# ---------------------------------------------------------------- red hat

STATUS = "md-select[formcontrolname=searchStatusControl]"
TIMEZONE = "input[formcontrolname=searchTimeZoneControl]"
LOCATION = "input[formcontrolname=searchLocationControl]"
DATE_FROM = "input[formcontrolname=searchFromDateControl]"
DATE_TO = "input[formcontrolname=searchToDateControl]"


def open_exam_deliveries(driver):
    for attempt in (1, 2):
        end = time.time() + 30
        while time.time() < end and not driver.find_elements(By.CSS_SELECTOR, ".menuIconCls"):
            if "sso.redhat.com" in driver.current_url or "login" in driver.current_url.lower():
                ask_user("Please log in to Red Hat in the Chrome window.")
                end = time.time() + 30
            time.sleep(1)
        menu = wait_for(driver, ".menuIconCls", timeout=5)
        if menu:
            click(driver, menu)
            time.sleep(1)
        if click_text(driver, "Exam Deliveries", timeout=10):
            log("Clicked menu > Exam Deliveries")
        if wait_for(driver, STATUS, timeout=20):
            time.sleep(1)
            return
        log(f"'Exam Deliveries' did not open (try {attempt} of 2) - page: {driver.current_url}")
        if attempt == 1:
            driver.refresh()

    page = f"{driver.current_url} ({driver.title})"
    shot = HERE / "logs" / f"exam_deliveries_failed_{dt.datetime.now():%Y%m%d_%H%M%S}.png"
    try:
        shot.parent.mkdir(exist_ok=True)
        driver.save_screenshot(str(shot))
        log(f"Screenshot of the page saved: logs\\{shot.name}")
    except Exception:
        pass
    ask_user(f"Could not open 'Exam Deliveries' (page: {page}). "
             "Click the 3 lines (top left) > Exam Deliveries.")
    if not wait_for(driver, STATUS, timeout=10):
        raise RuntimeError(f"'Exam Deliveries' did not open - page: {page}")
    time.sleep(1)


def pick_option(driver, text, timeout=10):
    """Click an open dropdown/autocomplete option containing `text`."""
    end = time.time() + timeout
    while time.time() < end:
        for o in driver.find_elements(By.CSS_SELECTOR, "md-option, mat-option"):
            if visible(o) and text.lower() in clean_cell_text(o.text).lower():
                click(driver, o)
                return True
        time.sleep(0.5)
    return False


def set_status(driver, status):
    sel = need(driver, STATUS, "Status")
    if clean_cell_text(sel.text).lower().endswith(status.lower()):
        return
    click(driver, sel)
    time.sleep(0.5)
    if not pick_option(driver, status):
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        ask_user(f"Could not choose Status '{status}'. Please choose it in Chrome.")
    log(f"Status = {status}")


def set_autocomplete(driver, css, typed, option_text, label):
    el = need(driver, css, label)
    type_into(el, typed)
    time.sleep(1)
    if not pick_option(driver, option_text):
        ask_user(f"Could not choose {label} '{option_text}'. Please choose it in Chrome.")
    time.sleep(0.5)
    log(f"{label} = {clean_cell_text(el.get_attribute('value'))}")


def set_date(driver, css, date, label):
    el = need(driver, css, label)
    type_into(el, date.isoformat())
    el.send_keys(Keys.TAB)
    time.sleep(0.5)
    value = el.get_attribute("value") or ""
    if value != date.isoformat():
        ask_user(f"{label} date shows '{value}'. Please set it to {date:%d-%b-%Y} in Chrome.")
    log(f"{label} = {date.isoformat()}")


READ_ROWS_JS = r"""
const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
const cell = (r, k) => { const c = r.querySelector('.cdk-column-' + k); return c ? c.innerText.replace(/\s+/g, ' ').trim() : ''; };
const keys = ['candidateName', 'candidateEmail', 'candidateContactNumber', 'examCodeName',
              'eligibilityCode', 'locationName', 'examStartDateTime', 'statusDisplay'];
const rows = [...document.querySelectorAll('md-row, mat-row')].filter(vis)
    .map(r => Object.fromEntries(keys.map(k => [k, cell(r, k)])));
const label = document.querySelector('.mat-paginator-range-label');
const next = document.querySelector('.mat-paginator-navigation-next');
const busy = [...document.querySelectorAll('md-spinner, md-progress-spinner, md-progress-bar, mat-spinner, .loading, .spinner')].some(vis);
return {rows, label: label ? label.innerText.trim() : '', hasNext: !!(next && !next.disabled), busy};
"""


# Count the page's server requests, so after SEARCH / Next we can wait for the
# answer instead of reading the old table.
TRACK_REQUESTS_JS = r"""
if (!window.__rq) {
  window.__rq = {pending: 0, done: 0};
  const send = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function () {
    window.__rq.pending++;
    this.addEventListener('loadend', () => { window.__rq.pending--; window.__rq.done++; });
    return send.apply(this, arguments);
  };
  if (window.fetch) {
    const f = window.fetch;
    window.fetch = function () {
      window.__rq.pending++;
      return f.apply(this, arguments).finally(() => { window.__rq.pending--; window.__rq.done++; });
    };
  }
}
return window.__rq.done;
"""


def requests_done(driver):
    return driver.execute_script(TRACK_REQUESTS_JS)


def read_page(driver, done_before=None, must_ask=True, settle=1.0, timeout=30):
    """Wait for the server answer (if done_before is given) and until the table
    stops changing, then return its rows. must_ask=False: the page may sort
    things out by itself (paging), so don't wait long if it asks nothing."""
    start = time.time()
    end = start + timeout
    if done_before is not None:
        while time.time() < end:
            rq = driver.execute_script("return window.__rq || {pending: 0, done: 0};")
            if rq["pending"] == 0 and (rq["done"] > done_before or (not must_ask and time.time() - start > 3)):
                break
            time.sleep(0.3)
        else:
            # the table still shows the last search: reading it could look like "no exams"
            if must_ask:
                raise RuntimeError("The Red Hat page did not answer the search in time")
            log("  (the page did not answer in time - reading the table anyway)")
    last = None
    while True:
        time.sleep(settle)
        page = driver.execute_script(READ_ROWS_JS)
        key = (page["label"], json.dumps(page["rows"]))
        if not page["busy"] and key == last:
            return page
        if time.time() > end:
            return page
        last = key


def biggest_page_size(driver):
    sel = wait_for(driver, ".mat-paginator-page-size-select", timeout=5)
    if not sel:
        return
    click(driver, sel)
    time.sleep(0.5)
    opts = [o for o in driver.find_elements(By.CSS_SELECTOR, "md-option, mat-option") if visible(o)]
    if opts:
        click(driver, opts[-1])
        time.sleep(1)
    else:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)


def button_text(b):
    # buttons start with an icon-font character (e.g.  before SEARCH)
    return clean_cell_text(re.sub(r"[-]", "", b.text)).upper()


def click_button(driver, text, timeout=5):
    end = time.time() + timeout
    while time.time() < end:
        for b in driver.find_elements(By.TAG_NAME, "button"):
            if visible(b) and button_text(b) == text.upper():
                click(driver, b)
                return True
        time.sleep(0.5)
    log(f"  buttons on the page: {[(clean_cell_text(b.text), visible(b)) for b in driver.find_elements(By.TAG_NAME, 'button')]}")
    return False


def search_rows(driver, code):
    """Search one location (with the Status already chosen) -> every row of every page."""
    set_autocomplete(driver, LOCATION, code, f"- {code}", "Location")
    before = requests_done(driver)
    if not click_button(driver, "SEARCH"):
        ask_user("Could not click SEARCH. Please click it in Chrome.")
        before = None
    log(f"Searching location {code}")
    page = read_page(driver, before)
    if page["hasNext"]:
        before = requests_done(driver)
        biggest_page_size(driver)
        page = read_page(driver, before, must_ask=False)

    rows, seen_pages = list(page["rows"]), {page["label"]}
    while page["hasNext"]:
        before = requests_done(driver)
        nxt = driver.find_element(By.CSS_SELECTOR, ".mat-paginator-navigation-next")
        click(driver, nxt)
        page = read_page(driver, before, must_ask=False)
        if page["label"] in seen_pages:
            break
        seen_pages.add(page["label"])
        rows += page["rows"]
    log(f"  {page['label'] or len(rows)} result(s)")
    return rows


def is_test_booking(r):
    # Red Hat's own test booking, not a candidate
    return exam_code(r["examCodeName"]) == "EX999" or clean_cell_text(r["candidateEmail"]).lower().endswith("@redhat.com")


def search_location(driver, cfg, date, code, rows=None):
    rows = search_rows(driver, code) if rows is None else rows
    people = []
    status = cfg.get("redhat", "status").lower()
    for r in rows:
        name = clean_cell_text(r["candidateName"])
        if is_test_booking(r):
            log(f"  skipped (Red Hat test booking): {name} | {r['candidateEmail']}")
            continue
        # safety: only that location, that day and that status
        if code not in r["locationName"]:
            log(f"  skipped (other location): {name} | {r['locationName']}")
            continue
        if not r["examStartDateTime"].startswith(date.isoformat()):
            log(f"  skipped (other date): {name} | {r['examStartDateTime']}")
            continue
        if r["statusDisplay"].lower() != status:
            log(f"  skipped (status {r['statusDisplay']}): {name}")
            continue
        person = {"location": code, "name": name, "email": clean_cell_text(r["candidateEmail"]),
                  "phone": clean_cell_text(r["candidateContactNumber"]),
                  "code": exam_code(r["examCodeName"]),
                  "status": "retake" if "retake" in r["eligibilityCode"].lower() else "y",
                  "eligibility": r["eligibilityCode"], "our_student": ""}
        people.append(person)
        log(f"  {name} | {person['phone']} | {person['email']} | {person['code']} | "
            f"{person['eligibility']} -> {person['status']}")
    return people


def scrape_redhat(cfg, date):
    driver = open_chrome(cfg, "redhat", cfg.get("redhat", "url"), "rhtapps.redhat.com",
                         close_old="individualexamscheduler")
    open_exam_deliveries(driver)
    set_status(driver, cfg.get("redhat", "status"))
    tz = cfg.get("redhat", "time_zone")
    set_autocomplete(driver, TIMEZONE, tz, tz, "Time Zone")
    set_date(driver, DATE_FROM, date, "From")
    set_date(driver, DATE_TO, date, "To")
    codes = [c.strip() for c in cfg.get("redhat", "locations").split(",") if c.strip()]
    people, kiosk = [], {c: {} for c in codes}
    for code in codes:
        rows = search_rows(driver, code)
        people += search_location(driver, cfg, date, code, rows)
        add_kiosk(kiosk, date, code, rows)
    log(f"{len(people)} completed exam(s) on {date:%d-%b-%Y}")
    # the kiosk table counts every exam of the day, so also the ones still "Scheduled"
    for status in [s.strip() for s in cfg.get("redhat", "kiosk_statuses", fallback="Scheduled").split(",") if s.strip()]:
        if status.lower() == cfg.get("redhat", "status").lower():
            continue
        set_status(driver, status)
        for code in codes:
            add_kiosk(kiosk, date, code, search_rows(driver, code))
    counts = {c: len(v) for c, v in kiosk.items()}
    for c, v in kiosk.items():
        log(f"  kiosk {c}: {len(v)} exam(s)" + (f" - {', '.join(sorted(v.values()))}" if v else ""))
    return people, counts


def add_kiosk(kiosk, date, code, rows):
    """Count each exam of the day at this location once (any status except cancelled)."""
    for r in rows:
        if is_test_booking(r) or code not in r["locationName"] or not r["examStartDateTime"].startswith(date.isoformat()):
            continue
        if "cancel" in r["statusDisplay"].lower():
            continue
        key = (clean_cell_text(r["eligibilityCode"]) or clean_cell_text(r["candidateName"]).lower(),
               exam_code(r["examCodeName"]), r["examStartDateTime"])
        kiosk[code][key] = f"{clean_cell_text(r['candidateName'])} ({r['statusDisplay']})"


# ---------------------------------------------------------------- adv pro

class AdvPro:
    """Drives the Adv Pro desktop program: login, Master > Registration > Display,
    Filter By PresentPhoneNo, and tells if a phone number is registered."""

    GRID = "DataGridWndClass"

    def __init__(self, cfg):
        from pywinauto import Application
        # pywinauto turns its own warnings back on when imported
        warnings.filterwarnings("ignore", message=".*32-bit application.*")
        self.cfg = cfg
        exe = cfg.get("advpro", "exe")
        try:
            self.app = Application(backend="win32").connect(path=Path(exe).name, timeout=2)
            self.opened = False
            log("Adv Pro is already open")
        except Exception:
            log("Opening Adv Pro")
            self.app = Application(backend="win32").start(f'"{exe}"', work_dir=str(Path(exe).parent))
            self.opened = True
        self.main = self.app.window(class_name="ThunderRT6MDIForm")
        self.login()
        self.form = self.open_display()
        self.grid = self.form.child_window(class_name=self.GRID).wrapper_object()
        self.search = self.form.child_window(class_name="ThunderRT6TextBox", control_id=51).wrapper_object()
        combo = self.form.child_window(class_name="ThunderRT6ComboBox", control_id=52).wrapper_object()
        combo.select("PresentPhoneNo")
        time.sleep(0.5)
        if combo.selected_text() != "PresentPhoneNo":
            ask_user("Please choose 'PresentPhoneNo' in 'Filter By' in Adv Pro.")
        log("Adv Pro: Filter By = PresentPhoneNo")

    def login(self):
        login = self.app.window(title="Form1", class_name="ThunderRT6FormDC")
        end = time.time() + 30
        while time.time() < end:
            if self.main.exists() and self.main.is_visible():
                return
            if login.exists() and login.is_visible():
                break
            time.sleep(1)
        else:
            ask_user("Adv Pro did not show the login screen. Log in by hand.")
            return
        user = login.child_window(class_name="ThunderRT6TextBox", control_id=1).wrapper_object()
        pwd = login.child_window(class_name="ThunderRT6TextBox", control_id=2).wrapper_object()
        self.set_text(user, self.cfg.get("advpro", "username"))
        self.set_text(pwd, self.cfg.get("advpro", "password"))
        pwd.set_focus()
        pwd.type_keys("{ENTER}")
        if not self.main.wait("visible", timeout=30):
            ask_user("Adv Pro login did not finish. Log in by hand.")
        log("Adv Pro: logged in")

    def display_form(self):
        for child in self.main.children(class_name="ThunderRT6FormDC"):
            if child.is_visible() and child.children(class_name=self.GRID):
                return self.main.child_window(handle=child.handle)
        return None

    def open_display(self):
        form = self.display_form()
        if not form:
            self.main.set_focus()
            self.main.menu_select("Master->Registration->Display")
            log("Adv Pro: Master > Registration > Display")
            end = time.time() + 20
            while not form and time.time() < end:
                time.sleep(1)
                form = self.display_form()
        if not form:
            ask_user("Open Master > Registration > Display in Adv Pro.")
            form = self.display_form()
        if not form:
            raise RuntimeError("Adv Pro Registration Display screen not found")
        return form

    @staticmethod
    def set_text(ctrl, text):
        import win32con
        import win32gui
        win32gui.SendMessage(ctrl.handle, win32con.WM_SETTEXT, 0, text)

    def grid_has_rows(self):
        """The list shows columns (and a scroll bar) only when something was found.
        Read it two ways: UI Automation, and the screen pixels of the list."""
        from pywinauto.controls.uiawrapper import UIAWrapper
        from pywinauto.uia_element_info import UIAElementInfo
        try:
            uia = len(UIAWrapper(UIAElementInfo(self.grid.handle)).children()) > 0
        except Exception:
            uia = None
        try:
            from PIL import ImageGrab
            r = self.grid.rectangle()
            img = ImageGrab.grab(bbox=(r.left, r.top, r.right, r.bottom), all_screens=True).convert("RGB")
            w, h = img.size
            px = img.load()
            # an empty list is grey right of the first two blank columns
            white = sum(1 for x in range(260, w - 4, 2) for y in range(2, h - 18, 2) if min(px[x, y]) > 245)
            pixels = white > 150
        except Exception:
            pixels = None
        return uia, pixels

    def close(self):
        """Close Adv Pro when this program opened it (left open if someone had it open)."""
        if not getattr(self, "opened", False):
            log("Adv Pro was already open before - left open")
            return
        try:
            self.main.close()
            for _ in range(10):
                time.sleep(0.5)
                if not self.app.is_process_running():
                    break
            else:
                self.app.kill()  # an "exit?" question or a stuck window
        except Exception:
            try:
                self.app.kill()
            except Exception:
                pass
        log("Adv Pro closed")

    def is_student(self, phone):
        p = phone10(phone)
        if not p:
            return ""
        self.set_text(self.search, p)
        wait = self.cfg.getfloat("advpro", "search_wait", fallback=2)
        time.sleep(wait)
        last = self.grid_has_rows()
        for _ in range(6):  # wait until the list stops changing
            time.sleep(0.7)
            now = self.grid_has_rows()
            if now == last:
                break
            last = now
        uia, pixels = last
        if uia is not None and pixels is not None and uia != pixels:
            log(f"  ? Adv Pro {p}: unsure (list check {uia}, screen check {pixels}) - check by hand")
        found = uia if uia is not None else pixels
        if found is None:  # neither check worked - that is not the same as "not our student"
            log(f"  ? Adv Pro {p}: could not read the list - 'Our Student' left blank")
            return ""
        return "y" if found else "n"


def check_advpro(cfg, people):
    todo = [p for p in people if phone10(p["phone"])]
    if not todo:
        return
    adv = AdvPro(cfg)
    try:
        adv.main.set_focus()  # keep it in front for the screen check
        for p in people:
            if not phone10(p["phone"]):
                log(f"  Adv Pro: {p['name']} has no phone number - 'Our Student' left blank")
                continue
            p["our_student"] = adv.is_student(p["phone"])
            log(f"  Adv Pro {phone10(p['phone'])} {p['name']}: {p['our_student']}")
    finally:
        adv.close()


# ---------------------------------------------------------------- excel

def check_excel_closed(path, wait_seconds=300):
    """Wait a while for the Excel file to be closed. An error (not sys.exit), so
    main.py still runs the other steps."""
    end = time.time() + wait_seconds
    told = False
    while True:
        try:
            with open(path, "r+b"):
                return
        except PermissionError:
            if time.time() >= end:
                raise RuntimeError(f"The Excel file is open on some computer. Close it and run again: {path}")
            if not told:
                log(f"The Excel file is open on some computer - close it. Waiting up to {wait_seconds // 60} minutes...")
                told = True
            time.sleep(15)


def month_sheet(wb, date):
    abbr = MONTHS[date.month - 1]
    for ws in wb.worksheets:
        if ws.title.strip().upper().startswith(abbr):
            return ws
    return new_month_sheet(wb, date, abbr)


def new_month_sheet(wb, date, abbr):
    """New month: copy the last month's sheet, with no candidates and no exam counts,
    and the new month's dates (Monday to Friday) in the count table at the top."""
    prev = [ws for ws in wb.worksheets if ws.title.strip().upper()[:3] in MONTHS][-1]
    ws = wb.copy_worksheet(prev)
    ws.title = abbr
    hr = header_row(ws)
    last = last_data_row(ws, hr)
    if last > hr + 1:
        ws.delete_rows(hr + 2, last - hr - 1)
    for c in range(1, ws.max_column + 1):  # this empty row keeps the formatting for the new rows
        ws.cell(hr + 1, c).value = None
    if isinstance(ws["A1"].value, str):
        ws["A1"].value = re.sub(r"^\s*[A-Za-z]+", date.strftime("%B"), ws["A1"].value, count=1)

    table = []
    for r in range(3, hr):
        if any(str(ws.cell(r, c).value or "").strip().lower() == "total" for c in range(1, 6)):
            break
        table.append(r)
    first = date.replace(day=1)
    days = [first + dt.timedelta(days=i) for i in range(31)]
    days = [d for d in days if d.month == date.month and d.weekday() < 5]
    for r in table:
        for c in (1, 2, 3, 4):  # the date and the 3 location counts; "exam number" keeps its formula
            ws.cell(r, c).value = None
    for r, day in zip(table, days):
        ws.cell(r, 1).value = dt.datetime(day.year, day.month, day.day)
    if len(days) > len(table):
        log(f"  ! The count table has room for {len(table)} of the {len(days)} working days - add the rest by hand")
    log(f"Created new sheet {abbr} from {prev.title}")
    return ws


def check_workbook_year(path, date):
    """The sheets are only named by month: a January 2027 run would add its rows to
    January 2026 in the 2026 file."""
    m = re.search(r"20\d\d", Path(path).stem)
    if m and int(m.group(0)) != date.year:
        raise RuntimeError(f"The Excel file is for {m.group(0)} but the date is {date:%d-%b-%Y} - "
                           "start the new year's file and change [excel] path in config.ini")


def header_row(ws):
    """The candidate table header: 'Date' in A and 'Name' in B (the count table above has 2093 in B)."""
    for r in range(1, ws.max_row + 1):
        if (str(ws.cell(r, 1).value or "").strip().lower() == "date"
                and str(ws.cell(r, 2).value or "").strip().lower() == "name"):
            return r
    raise RuntimeError(f"Header row (Date, Name) not found in sheet {ws.title}")


def sheet_columns(ws, hr):
    cols = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(hr, c).value or "").strip().lower()
        if h:
            cols.setdefault(h, c)
    if "exem code" in cols:  # spelled this way in the JAN-SEP sheets
        cols.setdefault("exam code", cols["exem code"])
    for need in ("date", "name", "phone no", "email id", "exam code", "exam status", "our student y/n"):
        if need not in cols:
            raise RuntimeError(f"Column '{need}' not found in sheet {ws.title}")
    return cols


def cell_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


def last_data_row(ws, hr):
    for r in range(ws.max_row, hr, -1):
        if not is_blank(ws.cell(r, 1).value) or not is_blank(ws.cell(r, 2).value):
            return r
    return hr


def safe_save(wb, path):
    """Save to a new file, check it, then swap it in: being stopped halfway
    through a save can't damage the real Excel file."""
    path = Path(path)
    for stale in path.parent.glob(f"~tmp_{path.stem}_*.xlsx"):  # left by a run that was stopped
        stale.unlink(missing_ok=True)
    tmp = path.with_name(f"~tmp_{path.stem}_{os.getpid()}.xlsx")
    try:
        wb.save(tmp)
        openpyxl.load_workbook(tmp, read_only=True).close()
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


def update_excel(path, date, people, backup_dir, dry_run=False):
    wb = openpyxl.load_workbook(path)
    ws = month_sheet(wb, date)
    hr = header_row(ws)
    cols = sheet_columns(ws, hr)

    def day_rows():
        return [r for r in range(hr + 1, last_data_row(ws, hr) + 1)
                if cell_date(ws.cell(r, cols["date"]).value) == date]

    def same_person(r, p):
        # the same person can take two different exams on one day: those are two rows
        code = str(ws.cell(r, cols["exam code"]).value or "").strip().upper()
        if code and p["code"] and code[:5] != p["code"][:5]:
            return False
        email = p["email"].lower()
        row_email = str(ws.cell(r, cols["email id"]).value or "").strip().lower()
        if (email and email == row_email) or norm_name(p["name"]) == norm_name(ws.cell(r, cols["name"]).value):
            return True
        # a shared number (family, company) with another name and email is another person
        return not row_email and phone10(p["phone"]) and phone10(p["phone"]) == phone10(ws.cell(r, cols["phone no"]).value)

    added = filled = 0
    for p in people:
        rows = day_rows()
        existing = next((r for r in rows if not is_blank(ws.cell(r, cols["name"]).value) and same_person(r, p)), None)
        if existing:
            log(f"  already in sheet: {p['name']}")
            # a run whose Adv Pro check failed left this blank - fill it now
            student = ws.cell(existing, cols["our student y/n"])
            if p["our_student"] and is_blank(student.value):
                student.value = p["our_student"]
                filled += 1
                log(f"  row {existing}: {p['name']} our student {p['our_student']}")
            continue
        # use the empty row already made for this date, else add one after that date's rows
        blank = [r for r in rows if is_blank(ws.cell(r, cols["name"]).value)]
        if blank:
            row = blank[0]
        else:
            after = rows[-1] if rows else max(
                [r for r in range(hr + 1, last_data_row(ws, hr) + 1)
                 if (cell_date(ws.cell(r, cols["date"]).value) or dt.date.min) < date] or [hr])
            row = after + 1
            if row <= ws.max_row:
                ws.insert_rows(row)
            style_row = row - 1 if row - 1 > hr else row + 1
            for c in range(1, ws.max_column + 1):
                ws.cell(row, c)._style = copy(ws.cell(style_row, c)._style)
            ws.row_dimensions[row].height = ws.row_dimensions[style_row].height
        ws.cell(row, cols["date"]).value = dt.datetime(date.year, date.month, date.day)
        ws.cell(row, cols["name"]).value = p["name"]
        ws.cell(row, cols["phone no"]).value = clean_phone(p["phone"])
        ws.cell(row, cols["email id"]).value = p["email"] or None
        ws.cell(row, cols["exam code"]).value = p["code"]
        ws.cell(row, cols["exam status"]).value = p["status"]
        ws.cell(row, cols["our student y/n"]).value = p["our_student"] or None
        added += 1
        log(f"  row {row}: {p['name']} | {p['phone']} | {p['email']} | {p['code']} | "
            f"{p['status']} | our student {p['our_student'] or '-'}  ({p['location']})")

    log(f"Sheet {ws.title}: {added} row(s) added" + (f", {filled} 'Our Student' filled in" if filled else ""))
    if dry_run:
        log("Dry run - Excel file NOT saved")
        return
    if not added and not filled:
        log("Nothing changed - Excel file not saved")
        return
    bdir = HERE / backup_dir
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{Path(path).stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    log(f"Backup saved: {backup}")
    safe_save(wb, path)
    log("Excel file saved")


def check_day_complete(cfg, date, skip_advpro=False):
    """Read the saved Excel file back: everyone found on Red Hat has a row for the date,
    with 'Our Student Y/N' if they have a phone number. Raises if not.
    An empty linkedin id or feedback cell is normal (nothing found, or no form filled in)."""
    wb = openpyxl.load_workbook(cfg.get("excel", "path"))
    ws = month_sheet(wb, date)
    hr = header_row(ws)
    cols = sheet_columns(ws, hr)
    rows = [r for r in range(hr + 1, last_data_row(ws, hr) + 1)
            if cell_date(ws.cell(r, cols["date"]).value) == date and not is_blank(ws.cell(r, cols["name"]).value)]
    saved = sorted((HERE / "logs").glob(f"{date}_*.json"), key=lambda p: p.stat().st_mtime)
    people = json.loads(saved[-1].read_text(encoding="utf-8")) if saved else []
    missing, unchecked = [], []
    for p in people:
        email = p["email"].lower()
        row = next((r for r in rows if norm_name(ws.cell(r, cols["name"]).value) == norm_name(p["name"])
                    or (email and email == str(ws.cell(r, cols["email id"]).value or "").strip().lower())), None)
        if not row:
            missing.append(p["name"])
        elif not skip_advpro and phone10(p["phone"]) and is_blank(ws.cell(row, cols["our student y/n"]).value):
            unchecked.append(p["name"])
    problems = []
    if missing:
        problems.append(f"not on the sheet: {', '.join(missing)}")
    if unchecked:
        problems.append(f"no 'Our Student Y/N': {', '.join(unchecked)}")
    if problems:
        raise RuntimeError(f"Excel not complete for {date:%d-%b-%Y} - " + "; ".join(problems))
    log(f"Excel complete for {date:%d-%b-%Y}: {len(rows)} row(s) in sheet {ws.title}, {len(people)} found on Red Hat")


# ---------------------------------------------------------------- kiosk table

def kiosk_table(ws, hr, codes):
    """The count table above the candidates: header row (Date | 2093 | 2337 | 2394 | exam number)
    -> (header row, {location: column}, total column, [table rows])"""
    for r in range(1, hr):
        if str(ws.cell(r, 1).value or "").strip().lower() != "date":
            continue
        heads = {str(ws.cell(r, c).value or "").strip().lower(): c for c in range(2, 12)}
        cols = {code: heads[code] for code in codes if code in heads}
        if not cols:
            continue
        total = next((c for h, c in heads.items() if "exam number" in h or h == "total"), None)
        rows = []
        for rr in range(r + 1, hr):
            if any(str(ws.cell(rr, c).value or "").strip().lower() == "total" for c in range(1, 6)):
                break
            rows.append(rr)
        return r, cols, total, rows
    raise RuntimeError(f"Kiosk table (Date | {' | '.join(codes)}) not found at the top of sheet {ws.title}")


def update_kiosk(path, date, counts, backup_dir, dry_run=False):
    """counts {location: number of exams} -> that date's row of the kiosk table (blank = 0)."""
    if not dry_run:
        check_excel_closed(path)
    wb = openpyxl.load_workbook(path)
    ws = month_sheet(wb, date)
    hr = header_row(ws)
    top, cols, total_col, rows = kiosk_table(ws, hr, list(counts))
    cells = [1, *cols.values()] + ([total_col] if total_col else [])
    row = next((r for r in rows if cell_date(ws.cell(r, 1).value) == date), None)
    if not row:  # e.g. a Saturday: use the first empty row after the last date
        dated = [r for r in rows if cell_date(ws.cell(r, 1).value)]
        free = [r for r in rows if (not dated or r > dated[-1]) and all(is_blank(ws.cell(r, c).value) for c in cells)]
        free = [r for r in free if r < hr - 1]  # keep a blank line above the candidates
        if not free:
            raise RuntimeError(f"No row for {date:%d-%b-%Y} in the kiosk table and no empty row to add it")
        row = free[0]
        if row - 1 > top:
            for c in cells:
                if ws.cell(row - 1, c).has_style:
                    ws.cell(row, c)._style = copy(ws.cell(row - 1, c)._style)
        ws.cell(row, 1).value = dt.datetime(date.year, date.month, date.day)
        log(f"  kiosk table: {date:%d-%b-%Y} added in row {row}")
    total = sum(counts.values())
    wanted = {cols[c]: (n or None) for c, n in counts.items() if c in cols}
    if total_col:
        cell = ws.cell(row, total_col)
        if not (isinstance(cell.value, str) and cell.value.startswith("=")):  # a formula adds them itself
            wanted[total_col] = total or None
    changed = False
    for c, value in wanted.items():
        cell = ws.cell(row, c)
        old = cell.value
        if is_blank(old):
            if value is not None:
                cell.value = value
                changed = True
        elif str(old).strip() != str(value or "").strip():
            log(f"  kiosk table row {row}: '{ws.cell(top, c).value}' already says {old}, "
                f"Red Hat says {value or 0} - left as it is")
    log(f"Kiosk {date:%d-%b-%Y}: " + ", ".join(f"{c}: {n}" for c, n in counts.items()) + f", total {total}"
        + ("" if changed else " (nothing new for the table)"))
    if dry_run:
        log("Dry run - kiosk table NOT saved")
        return
    if not changed:
        return
    bdir = HERE / backup_dir
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{Path(path).stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    safe_save(wb, path)
    log(f"Excel file saved - kiosk table row {row}")


# ---------------------------------------------------------------- step cleanup

def chrome_tab_ids(cfg):
    port = cfg.getint("chrome", "debug_port")
    return {t["id"] for t in chrome_pages(port)} if port_open(port) else set()


def step_cleanup(cfg, tabs_before):
    """After a step: close the Chrome tabs it opened and let go of ChromeDriver,
    so the next step starts clean. Tabs that were open before are left alone."""
    port = cfg.getint("chrome", "debug_port")
    for d in _DRIVERS:
        try:
            d.service.stop()  # only ChromeDriver; Chrome itself stays open
        except Exception:
            pass
    _DRIVERS.clear()
    if not port_open(port):
        return
    pages = chrome_pages(port)
    new = [t for t in pages if t["id"] not in tabs_before]
    if new and len(new) == len(pages):
        new = new[:-1]  # never close the last tab: that would close Chrome
    for t in new:
        close_tab(port, t["id"])
    if new:
        log(f"Closed {len(new)} Chrome tab(s) this step opened")


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="YYYY-MM-DD (default: asks, Enter = today)")
    ap.add_argument("--skip-advpro", action="store_true", help="don't check the phone numbers in Adv Pro")
    ap.add_argument("--dry-run", action="store_true", help="do everything except saving the Excel file")
    args = ap.parse_args()

    cfg = load_config()
    date = dt.date.fromisoformat(args.date) if args.date else ask_date()
    log(f"Date: {date:%d-%b-%Y}   Excel: {cfg.get('excel', 'path')}")
    run(cfg, date, skip_advpro=args.skip_advpro, dry_run=args.dry_run)
    log("Done")


def run(cfg, date, skip_advpro=False, dry_run=False):
    """Red Hat completed exams of `date` + Adv Pro check -> new rows in the Excel file."""
    excel = cfg.get("excel", "path")
    check_workbook_year(excel, date)
    if not dry_run:
        check_excel_closed(excel)

    people, kiosk = scrape_redhat(cfg, date)
    try:
        update_kiosk(excel, date, kiosk, cfg.get("excel", "backup_dir"), dry_run=dry_run)
    except Exception as e:  # the exam rows matter more - keep going
        log(f"!! Kiosk table not updated: {type(e).__name__}: {e}")
    if people and not skip_advpro:
        try:
            check_advpro(cfg, people)
        except Exception as e:  # keep the exams already read from Red Hat
            log(f"!! Adv Pro check failed: {type(e).__name__}: {e}")
            log("   Saving the rows anyway - run this date again to fill the blank 'Our Student' cells")

    logs = HERE / "logs"
    logs.mkdir(exist_ok=True)
    (logs / f"{date}_{dt.datetime.now():%H%M%S}.json").write_text(
        json.dumps(people, indent=2, ensure_ascii=False), encoding="utf-8")

    if not people:
        log("No completed exams found - nothing to write")
        return
    if not dry_run:
        check_excel_closed(excel)
    update_excel(excel, date, people, cfg.get("excel", "backup_dir"), dry_run=dry_run)
    if not skip_advpro:
        unchecked = [p["name"] for p in people if phone10(p["phone"]) and not p["our_student"]]
        if unchecked:  # the rows are saved; this step must run again to fill these in
            raise RuntimeError(f"'Our Student' not checked in Adv Pro for: {', '.join(unchecked)}")


if __name__ == "__main__":
    main()
