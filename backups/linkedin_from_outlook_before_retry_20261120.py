"""
Excel "Email ID"  ->  Outlook contact card  ->  LinkedIn profile URL  ->  Excel "linkedin id"

Steps
  1. Ask the date (Enter = every row of this month).
  2. Read "redhat exam 2026.xlsx": rows with an Email ID but no linkedin id.
  3. Open Chrome with the "Arish - Chrome" profile at outlook.live.com/mail and
     click "New mail".
  4. For each email: put it in "To", hover over it so the contact card shows,
     click the LinkedIn logo, click "See profile on LinkedIn" and take the
     linkedin.com/in/... address from the tab that opens.
  5. Back up the Excel file and write the LinkedIn URLs into "linkedin id".
     The mail is never sent - the draft is discarded at the end.

Usage
  python linkedin_from_outlook.py                     # asks the date
  python linkedin_from_outlook.py --date 2026-10-07   # one day
  python linkedin_from_outlook.py --month 2026-10     # whole month
  python linkedin_from_outlook.py --limit 3           # only the first 3 emails (test)
  python linkedin_from_outlook.py --ask               # when a profile isn't found, wait so you can open it by hand
  python linkedin_from_outlook.py --dry-run           # don't save the Excel file
"""
import argparse
import datetime as dt
import json
import re
import shutil
import sys
import time
import urllib.parse
import urllib.request
from copy import copy
from pathlib import Path

import openpyxl
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import redhat_automation as pv
from redhat_automation import HERE, ask_user, click, click_text, log, visible

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$", re.I)


# ---------------------------------------------------------------- date / excel

def ask_scope():
    """-> (date or None, any day of the month to use)."""
    today = dt.date.today()
    while True:
        text = input(f"\n>>> Which date? Type it as DD-MM-YYYY (e.g. {today:%d-%m-%Y}),\n"
                     f"    or just press Enter for all of {today:%B %Y}: ").strip(" \t﻿")
        if not text:
            print(f"    Using every row of {today:%B %Y}")
            return None, today
        try:
            date = pv.parse_date(text)
        except ValueError:
            print("    Not a valid date - try again, e.g. 07-10-2026")
            continue
        print(f"    Using {date:%A, %d-%b-%Y}")
        return date, date


def linkedin_col(ws, hr):
    for c in range(1, ws.max_column + 1):
        if re.sub(r"\s+", "", str(ws.cell(hr, c).value or "")).lower() == "linkedinid":
            return c
    raise RuntimeError(f"Column 'Linkedin Id' not found in sheet {ws.title}")


def rows_to_do(path, date, month_day):
    wb = openpyxl.load_workbook(path)
    ws = pv.month_sheet(wb, month_day)
    hr = pv.header_row(ws)
    cols = pv.sheet_columns(ws, hr)
    li = linkedin_col(ws, hr)
    todo = []
    # profiles already in the sheet (any day), to catch a stale contact card
    taken = {}
    for r in range(hr + 1, pv.last_data_row(ws, hr) + 1):
        url = linkedin_profile(str(ws.cell(r, li).value or ""))
        if url:
            taken[url.lower()] = pv.clean_cell_text(ws.cell(r, cols["name"]).value)
    for r in range(hr + 1, pv.last_data_row(ws, hr) + 1):
        if date and pv.cell_date(ws.cell(r, cols["date"]).value) != date:
            continue
        name = pv.clean_cell_text(ws.cell(r, cols["name"]).value)
        if not name or not pv.is_blank(ws.cell(r, li).value):
            continue
        email = str(ws.cell(r, cols["email id"]).value or "").strip()
        company = pv.clean_cell_text(ws.cell(r, cols["company name"]).value) if "company name" in cols else ""
        # no (valid) email: Outlook can't help, but the Google search can
        todo.append({"row": r, "name": name, "email": email if EMAIL_RE.match(email) else "",
                     "company": company, "linkedin": ""})
    return ws.title, todo, taken


def save_results(path, sheet, results, backup_dir):
    found = [x for x in results if x["linkedin"]]
    if not found:
        log("No LinkedIn URLs found - Excel file not changed")
        return
    while True:
        try:
            pv.check_excel_closed(path)
            break
        except RuntimeError:
            ask_user(f"The Excel file is open somewhere. Close it so the {len(found)} LinkedIn URL(s) can be saved.")
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    hr = pv.header_row(ws)
    cols = pv.sheet_columns(ws, hr)
    li = linkedin_col(ws, hr)
    written = 0

    def is_row(r, x):
        if x["email"]:
            return str(ws.cell(r, cols["email id"]).value or "").strip().lower() == x["email"].lower()
        return pv.norm_name(ws.cell(r, cols["name"]).value) == pv.norm_name(x["name"])

    for x in found:
        row = x["row"]
        # the row may have moved if someone edited the file meanwhile -> find it by email (or name)
        if not is_row(row, x):
            row = next((r for r in range(hr + 1, ws.max_row + 1) if is_row(r, x)), None)
        if not row or not pv.is_blank(ws.cell(row, li).value):
            continue
        cell = ws.cell(row, li)
        above = ws.cell(row - 1, li)
        if row - 1 > hr and above.has_style:
            cell._style = copy(above._style)
        cell.value = x["linkedin"]
        written += 1
        log(f"  row {row}: {x['name']} -> {x['linkedin']}" + ("  (Google)" if x.get("source") == "google" else ""))
    if not written:
        log("Nothing to write - Excel file not changed")
        return
    bdir = HERE / backup_dir
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{Path(path).stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    log(f"Backup saved: {backup}")
    pv.safe_save(wb, path)
    log(f"Excel file saved - {written} LinkedIn URL(s) written to sheet {sheet}")


# ---------------------------------------------------------------- outlook

TO_BOX_CSS = ['div[contenteditable="true"][aria-label="To"]', 'input[aria-label="To"]',
              '[role="textbox"][aria-label^="To"]', 'div[contenteditable="true"][aria-label*="To"]']


def find_to_box(driver):
    for css in TO_BOX_CSS:
        for el in driver.find_elements(By.CSS_SELECTOR, css):
            if visible(el):
                return el
    return None


def open_compose(driver):
    end = time.time() + 40
    while time.time() < end:
        if find_to_box(driver):
            return find_to_box(driver)
        if "login" in driver.current_url or "signin" in driver.current_url:
            ask_user("Sign in to Outlook in the Arish Chrome window.")
            end = time.time() + 40
        elif click_text(driver, "New mail", "New message", "New email", timeout=3):
            log("Clicked New mail")
            time.sleep(3)
        else:
            time.sleep(1)
    ask_user("Could not open a new mail. Click 'New mail' in Outlook.")
    box = find_to_box(driver)
    if not box:
        raise RuntimeError("The 'To' box of the new mail was not found")
    return box


FIND_PILL_JS = r"""
const email = arguments[0].toLowerCase(), box = arguments[1];
const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
const has = e => [e.getAttribute('aria-label'), e.getAttribute('title'), e.innerText]
                 .some(v => (v || '').toLowerCase().includes(email));
// the recipient "pill" sits next to the To box - search the compose header around it
let area = box;
for (let i = 0; i < 6 && area.parentElement; i++) area = area.parentElement;
const hits = [...area.querySelectorAll('*')].filter(e => vis(e) && e !== box && has(e) && !e.isContentEditable);
// innermost element that matches
return hits.filter(e => !hits.some(o => o !== e && e.contains(o)))[0] || null;
"""

FIND_LINKEDIN_JS = r"""
const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
const label = e => [e.getAttribute('aria-label'), e.getAttribute('title'), e.getAttribute('alt'),
                    e.getAttribute('data-icon-name'), e.innerText].join(' ');
const els = [...document.querySelectorAll('button,a,[role=button],[role=tab],[role=link],img,svg,i')]
            .filter(e => vis(e) && /linked\s*in/i.test(label(e)) && !/see\s+profile|view\s+(full\s+)?profile/i.test(label(e)));
if (!els.length) return null;
const e = els[0];
return e.closest('button,a,[role=button],[role=tab],[role=link]') || e;
"""

FIND_SEE_PROFILE_JS = r"""
const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
const els = [...document.querySelectorAll('a,button,[role=button],[role=link]')].filter(e => vis(e) &&
    /(see|view)\s+(full\s+)?profile\s+on\s+linked\s*in|profile\s+on\s+linked\s*in/i.test(
        (e.innerText || '') + ' ' + (e.getAttribute('aria-label') || '') + ' ' + (e.getAttribute('title') || '')));
return els[0] || null;
"""


def linkedin_profile(url):
    """Any URL (also LinkedIn redirect / sign-in links) -> https://www.linkedin.com/in/<id>/"""
    for _ in range(3):
        m = re.search(r"linkedin\.com/in/([^/?#&\s\"']+)", url or "", re.I)
        if m:
            return f"https://www.linkedin.com/in/{m.group(1)}/"
        url = urllib.parse.unquote(url or "")
    return ""


def wait_for(driver, js, *args, timeout=8):
    end = time.time() + timeout
    while True:
        el = driver.execute_script(js, *args)
        if el or time.time() > end:
            return el
        time.sleep(0.5)


def hover(driver, el):
    """Put the mouse on the element, then move it slightly (the card shows on a light move)."""
    ac = ActionChains(driver)
    ac.move_to_element(el).pause(0.4).move_by_offset(3, 0).pause(0.4).move_by_offset(-3, 1).perform()


def new_linkedin_tab(port, before, timeout=20):
    """Wait for a tab that wasn't open before to show a LinkedIn address; return its URL and close it."""
    end = time.time() + timeout
    seen = None
    while time.time() < end:
        for t in pv.chrome_pages(port):
            if t["id"] in before:
                continue
            seen = t
            if linkedin_profile(t["url"]):
                close_tab(port, t["id"])
                return t["url"]
        time.sleep(0.5)
    if seen:
        log(f"    new tab did not reach a profile: {seen['url'][:100]}")
        close_tab(port, seen["id"])
    return ""


def close_tab(port, tab_id):
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/json/close/{tab_id}", timeout=10).close()
    except Exception:
        pass


def clear_to(driver):
    box = find_to_box(driver)
    if not box:
        return
    click(driver, box)
    for _ in range(4):  # Backspace in an empty To box removes the recipient before it
        box.send_keys(Keys.BACKSPACE)
        time.sleep(0.2)


def close_card(driver):
    box = find_to_box(driver)
    if box:
        ActionChains(driver).move_to_element(box).perform()  # mouse away from the pill
    if wait_for(driver, FIND_LINKEDIN_JS, timeout=0) or wait_for(driver, FIND_SEE_PROFILE_JS, timeout=0):
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.5)


def lookup(driver, port, outlook_tab, email):
    """Email -> LinkedIn profile URL ('' if Outlook shows none)."""
    box = find_to_box(driver) or open_compose(driver)
    clear_to(driver)
    box = find_to_box(driver)
    click(driver, box)
    box.send_keys(email)
    time.sleep(1.5)
    box.send_keys(";")  # turns the typed address into a recipient
    pill = wait_for(driver, FIND_PILL_JS, email, box, timeout=8)
    if not pill:
        log("    recipient not found in the To box")
        return ""

    hover(driver, pill)
    li = wait_for(driver, FIND_LINKEDIN_JS, timeout=6)
    if not li:  # hover again, then try a click on the recipient
        hover(driver, pill)
        li = wait_for(driver, FIND_LINKEDIN_JS, timeout=4)
    if not li:
        click(driver, pill)
        li = wait_for(driver, FIND_LINKEDIN_JS, timeout=6)
    if not li:
        log("    contact card has no LinkedIn logo")
        return ""
    click(driver, li)

    see = wait_for(driver, FIND_SEE_PROFILE_JS, timeout=10)
    if not see:
        log("    no 'See profile on LinkedIn' (no LinkedIn match for this email?)")
        return ""
    direct = linkedin_profile(see.get_attribute("href") or "")
    before = {t["id"] for t in pv.chrome_pages(port)}
    click(driver, see)
    url = new_linkedin_tab(port, before) or direct
    driver.switch_to.window(outlook_tab)
    return linkedin_profile(url)


def manual_lookup(driver, port, outlook_tab, person):
    before = {t["id"] for t in pv.chrome_pages(port)}
    ask_user(f"Open the LinkedIn profile of {person['name']} ({person['email']}) by hand "
             "(hover -> LinkedIn logo -> See profile on LinkedIn).\n"
             "    Leave the LinkedIn tab open. Just press Enter to skip.")
    url = ""
    for t in pv.chrome_pages(port):
        if t["id"] not in before and linkedin_profile(t["url"]):
            url = t["url"]
            close_tab(port, t["id"])
    driver.switch_to.window(outlook_tab)
    return linkedin_profile(url)


def discard_draft(driver):
    clear_to(driver)
    # the Discard button is an icon - find it by its label
    btn = next((e for e in driver.find_elements(By.CSS_SELECTOR, '[aria-label="Discard"]') if visible(e)), None)
    if btn:
        click(driver, btn)
    elif not click_text(driver, "Discard", timeout=3):
        log("Discard button not found - delete the empty draft from Outlook Drafts")
        return
    time.sleep(1)
    click_text(driver, "OK", "Discard", "Yes", timeout=3)  # confirmation, if Outlook asks
    log("Draft discarded")


# ---------------------------------------------------------------- google search -> linkedin
# For the people Outlook had no LinkedIn for: search Google (Arish Chrome) for
#   "<name> @ <company> linkedin profile"
# open the first LinkedIn profile in the results, and keep its URL only when the
# profile's name AND the company from the Excel row match. Otherwise skip the person.

LINKEDIN_LOGIN = ("/login", "/authwall", "/checkpoint", "/uas/", "/signup")

# words that say nothing about which company it is
COMPANY_SKIP = {"pvt", "private", "ltd", "limited", "inc", "llp", "llc", "corp", "corporation",
                "company", "the", "and", "india", "indian", "solutions", "solution", "services",
                "service", "technologies", "technology", "tech", "software", "systems", "group",
                "global", "consulting", "consultancy", "infotech", "info", "enterprises",
                "international", "student", "fresher", "working", "unemployed", "self", "employed",
                "nil", "none", "not", "applicable", "chennai", "bangalore", "bengaluru"}

GOOGLE_RESULTS_JS = r"""
const out = [];
for (const h of document.querySelectorAll('#search a h3, #rso a h3')) {
  const a = h.closest('a');
  if (!a || !a.href) continue;
  let box = a;  // the result block, for the snippet text under the title
  for (let i = 0; i < 6 && box.parentElement; i++) {
    box = box.parentElement;
    if (box.innerText.length > h.innerText.length + 60) break;
  }
  out.push({url: a.href, title: h.innerText, text: box.innerText});
  if (out.length >= 5) break;
}
return out;
"""

PROFILE_NAME_JS = """const h = document.querySelector('main h1') || document.querySelector('h1');
return h ? h.innerText : '';"""


def nobody_at_keyboard():
    return not sys.stdin or not sys.stdin.isatty()


def name_tokens(s):
    return re.findall(r"[a-z]+", str(s or "").lower())


def name_score(a, b):
    """How well name a fits name b ("V.M.Hemanth" vs "HEMANTH V M"). 0 = no match."""
    f, r = name_tokens(a), name_tokens(b)
    if not f or not r or not any(len(t) >= 3 for t in f):
        return 0

    def tok_ok(t):
        if len(t) == 1:  # initial
            return any(x.startswith(t) for x in r)
        return any(t == x or (len(t) >= 3 and len(x) >= 3 and (x.startswith(t) or t.startswith(x)))
                   for x in r)

    if all(tok_ok(t) for t in f):
        return 100 + sum(len(t) for t in f if len(t) >= 3)
    fj, rj = "".join(f), "".join(r)
    if len(fj) >= 5 and (fj in rj or rj in fj):  # "vairamuthu" vs "vaira muthu"
        return 90
    return 0


def names_match(excel_name, profile_name):
    """Every part of the Excel name must be on the profile ("YOGESHWARAN T L" = "Yogeshwaran TL, RHCSA").
    A shorter profile name is OK only when what it leaves out is there as an initial
    ("Karthikeyan R" = "Karthikeyan Rajan", but "Paul Raj" is NOT "Paul Jonathan Raj A")."""
    if name_score(excel_name, profile_name) >= 90:
        return True
    if name_score(profile_name, excel_name) < 100:
        return False
    p = name_tokens(profile_name)

    def covered(t):
        return any(t == x or (len(x) >= 3 and (x.startswith(t) or t.startswith(x)))
                   or (len(x) == 1 and t.startswith(x)) for x in p)

    return all(covered(t) for t in name_tokens(excel_name) if len(t) >= 3)


def company_words(company):
    return [w for w in re.findall(r"[a-z0-9]+", str(company or "").lower())
            if len(w) >= 3 and w not in COMPANY_SKIP]


def company_match(words, text):
    hits = sum(1 for w in words if re.search(rf"\b{re.escape(w)}\b", text or "", re.I))
    return hits >= max(1, (len(words) + 1) // 2)  # at least half of the company's words


def google_captcha(driver):
    if "/sorry/" in driver.current_url:
        return True
    body = driver.execute_script("return document.body ? document.body.innerText : ''") or ""
    return bool(re.search(r"unusual traffic|not a robot", body, re.I))


def google_search(driver, query):
    """-> the first results [{url, title, text}], or None when Google wants a CAPTCHA solved."""
    driver.get("https://www.google.com/search?hl=en&q=" + urllib.parse.quote_plus(query))
    for _ in range(2):
        end = time.time() + 15
        while time.time() < end:
            if google_captcha(driver):
                break
            results = driver.execute_script(GOOGLE_RESULTS_JS)
            if results:
                return results
            time.sleep(1)
        else:
            return []  # page loaded, no results
        if nobody_at_keyboard():
            return None
        ask_user("Google shows \"I'm not a robot\" in the Arish Chrome window - solve it there.")
        driver.get("https://www.google.com/search?hl=en&q=" + urllib.parse.quote_plus(query))
    return None


def read_profile(driver, url):
    """-> (name on the profile, page text), or None when LinkedIn wants a sign-in."""
    for _ in range(2):
        driver.get(url)
        time.sleep(3)
        if not any(k in driver.current_url for k in LINKEDIN_LOGIN):
            break
        if nobody_at_keyboard():
            return None
        ask_user("Sign in to LinkedIn in the Arish Chrome window.")
    else:
        return None
    name, end = "", time.time() + 15
    while not name and time.time() < end:
        name = (driver.execute_script(PROFILE_NAME_JS) or "").strip()
        if not name:
            time.sleep(1)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")  # loads Experience
    time.sleep(1.5)
    text = driver.execute_script(
        "const m = document.querySelector('main'); return (m || document.body).innerText;") or ""
    return name, text


def google_lookup(driver, cfg, people, taken):
    """Fills x['linkedin'] for the people whose Google result matches name + company."""
    delay = cfg.getfloat("linkedin", "google_delay", fallback=8)
    log(f"--- Google search for {len(people)} person(s) still without a LinkedIn Id")
    for i, x in enumerate(people, 1):
        tag = f"[Google {i}/{len(people)}] {x['name']}"
        words = company_words(x.get("company"))
        if not words:
            log(f"{tag}: no usable company name in the sheet ('{x.get('company', '')}') - skipped")
            continue
        query = f"{x['name']} @ {x['company']} linkedin profile"
        log(f"{tag}: {query}")
        try:
            results = google_search(driver, query)
            if results is None:
                log("    Google wants \"I'm not a robot\" solved - Google search stopped for now")
                return
            hit = next((r for r in results if linkedin_profile(r["url"])), None)
            if not hit:
                log("    no LinkedIn profile in the first results - skipped")
                time.sleep(delay)
                continue
            url = linkedin_profile(hit["url"])
            owner = taken.get(url.lower())
            if owner and pv.norm_name(owner) != pv.norm_name(x["name"]):
                log(f"    {url} is already {owner}'s LinkedIn Id in the sheet - skipped")
                time.sleep(delay)
                continue
            prof = read_profile(driver, url)
            if prof is None:
                log("    LinkedIn wants a sign-in - Google search stopped for now")
                return
            name, text = prof
            name = name or re.split(r"\s+[-|–]\s+", hit["title"])[0]
            ok_name = names_match(x["name"], name)
            ok_company = company_match(words, text + "\n" + hit["title"] + "\n" + hit["text"])
            if ok_name and ok_company:
                x["linkedin"], x["source"] = url, "google"
                taken[url.lower()] = x["name"]
                log(f"    match: '{name}' at {x['company']} -> {url}")
            else:
                log(f"    not a match - profile '{name}': name {'ok' if ok_name else 'different'}, "
                    f"company {'found' if ok_company else 'not found'} - skipped")
        except Exception as e:
            log(f"    error: {type(e).__name__}: {str(e).splitlines()[0][:150] if str(e) else ''} - skipped")
        time.sleep(delay)


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="YYYY-MM-DD: only that day")
    ap.add_argument("--month", help="YYYY-MM: every row of that month")
    ap.add_argument("--limit", type=int, help="only the first N emails")
    ap.add_argument("--ask", action="store_true", help="wait for you when a profile isn't found")
    ap.add_argument("--dry-run", action="store_true", help="don't save the Excel file")
    args = ap.parse_args()

    cfg = pv.load_config()
    if args.date:
        date = month_day = dt.date.fromisoformat(args.date)
    elif args.month:
        date, month_day = None, dt.date.fromisoformat(args.month + "-01")
    else:
        date, month_day = ask_scope()
    run(cfg, date, month_day, limit=args.limit, ask=args.ask, dry_run=args.dry_run)
    log("Done")


def run(cfg, date, month_day, limit=None, ask=False, dry_run=False):
    """Find LinkedIn URLs for the rows of `date` (None = the whole month of month_day)."""
    excel = cfg.get("excel", "path")
    pv.check_excel_closed(excel)

    sheet, todo, taken = rows_to_do(excel, date, month_day)
    if limit:
        todo = todo[:limit]
    log(f"Sheet {sheet}: {len(todo)} person(s) without a LinkedIn Id"
        + (f" on {date:%d-%b-%Y}" if date else ""))
    if not todo:
        return
    for x in todo:
        log(f"  row {x['row']}: {x['name']} | {x['email'] or '(no email)'} | {x['company'] or '(no company)'}")

    port = cfg.getint("chrome", "debug_port")

    def start_outlook():
        d = pv.open_chrome(cfg, "linkedin", cfg.get("linkedin", "outlook_url"), "live.com")
        open_compose(d)
        return d, d.current_window_handle

    def back_to_outlook():
        """Return to the Outlook tab; if it was closed, open Outlook again."""
        nonlocal driver, outlook_tab
        try:
            driver.switch_to.window(outlook_tab)
            driver.execute_script("return 1")
        except Exception:
            log("    Outlook tab was closed - opening it again")
            driver, outlook_tab = start_outlook()

    driver, outlook_tab = start_outlook()

    logs = HERE / "logs"
    logs.mkdir(exist_ok=True)
    progress = logs / f"linkedin_{dt.datetime.now():%Y%m%d_%H%M%S}.json"
    try:
        for i, x in enumerate(todo, 1):
            if not x["email"]:
                log(f"[{i}/{len(todo)}] {x['name']}: no email - left for the Google search")
                continue
            log(f"[{i}/{len(todo)}] {x['name']} | {x['email']}")
            try:
                x["linkedin"] = lookup(driver, port, outlook_tab, x["email"])
            except Exception as e:
                log(f"    error: {type(e).__name__}: {(str(e).splitlines() or [''])[0][:150]}")
                x["error"] = True
                back_to_outlook()
            if not x["linkedin"] and ask:
                x["linkedin"] = manual_lookup(driver, port, outlook_tab, x)
            owner = taken.get(x["linkedin"].lower()) if x["linkedin"] else None
            if owner and pv.norm_name(owner) != pv.norm_name(x["name"]):  # same name = a retake, same person
                # Outlook kept showing the card of an earlier person
                log(f"    ! got {x['linkedin']} which is {owner}'s profile - stale contact card, skipped")
                x["linkedin"] = ""
            elif x["linkedin"]:
                taken[x["linkedin"].lower()] = x["name"]
            log(f"    LinkedIn: {x['linkedin'] or '(not found)'}")
            try:
                close_card(driver)
            except Exception:
                back_to_outlook()
            progress.write_text(json.dumps(todo, indent=2), encoding="utf-8")
            time.sleep(1)

        # people Outlook had nothing for -> Google "<name> @ <company> linkedin profile"
        rest = [x for x in todo if not x["linkedin"]]
        if rest and cfg.getboolean("linkedin", "google_search", fallback=True):
            try:
                driver.switch_to.new_window("tab")  # keep the Outlook draft tab as it is
                google_lookup(driver, cfg, rest, taken)
            except KeyboardInterrupt:
                raise
            except Exception as e:
                log(f"Google search stopped: {type(e).__name__}: {str(e).splitlines()[0][:150] if str(e) else ''}")
            finally:
                try:
                    if driver.current_window_handle != outlook_tab:
                        driver.close()
                except Exception:
                    pass
                back_to_outlook()
            progress.write_text(json.dumps(todo, indent=2), encoding="utf-8")
    except KeyboardInterrupt:
        log("Stopped - saving what was found so far")
        raise  # this step is not finished; main.py must not count it as done
    finally:
        try:
            discard_draft(driver)
        except Exception:
            log("Could not discard the draft - delete it from Outlook Drafts")
        found = sum(1 for x in todo if x["linkedin"])
        by_google = sum(1 for x in todo if x.get("source") == "google")
        log(f"LinkedIn found for {found} of {len(todo)} person(s)"
            + (f" ({by_google} by Google search)" if by_google else ""))
        for x in todo:
            if not x["linkedin"]:
                log(f"  ! not found: {x['name']} | {x['email'] or '(no email)'} | {x['company'] or '(no company)'}")
        if dry_run:
            log("Dry run - Excel file NOT saved")
        else:
            save_results(excel, sheet, todo, cfg.get("excel", "backup_dir"))
    errors = sum(1 for x in todo if x.get("error") and not x["linkedin"])
    if errors:  # "no profile" is a normal answer; an error is not an answer
        raise RuntimeError(f"{errors} email(s) could not be looked up")


if __name__ == "__main__":
    main()
