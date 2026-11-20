# Created by YOGESHWARAN TL
"""
No feedback?  ->  LinkedIn profile  ->  Excel "Company Name" + "Designation"

For the candidates who did NOT fill in the ADV PRO feedback form (so their
Company Name / Designation are still blank) but whose "linkedin id" is known:

Steps
  1. Read "redhat exam 2026.xlsx": rows of the date with a linkedin id, the
     "Advantage Pro Feedback form" cell blank, and Company Name or Designation blank.
  2. Open Chrome with the "Arish - Chrome" profile, paste each LinkedIn URL in the
     address bar and read the profile: the current job in Experience
     (role + company, the one that says "Present"); if there is none, the headline
     ("DevOps Engineer at Infosys").
  3. Back up the Excel file and write the company into "Company Name" and the role
     into "Designation". Cells someone already filled are left alone.

Usage
  python company_from_linkedin.py                     # asks the date
  python company_from_linkedin.py --date 2026-10-07   # one day
  python company_from_linkedin.py --month 2026-10     # whole month
  python company_from_linkedin.py --limit 3           # only the first 3 people (test)
  python company_from_linkedin.py --dry-run           # don't save the Excel file
  python company_from_linkedin.py --test <linkedin url>   # only show what it reads from that profile
"""
import argparse
import datetime as dt
import json
import re
import shutil
import time
from copy import copy
from pathlib import Path

import openpyxl

import redhat_automation as pv
from feedback_from_sheet import feedback_columns
from linkedin_from_outlook import LINKEDIN_LOGIN, ask_scope, linkedin_col, linkedin_profile, nobody_at_keyboard
from redhat_automation import HERE, ask_user, log


# ---------------------------------------------------------------- excel

def rows_to_do(path, date, month_day):
    wb = openpyxl.load_workbook(path)
    ws = pv.month_sheet(wb, month_day)
    hr = pv.header_row(ws)
    cols = feedback_columns(ws, hr)
    li = linkedin_col(ws, hr)
    todo = []
    for r in range(hr + 1, pv.last_data_row(ws, hr) + 1):
        if date and pv.cell_date(ws.cell(r, cols["date"]).value) != date:
            continue
        name = pv.clean_cell_text(ws.cell(r, cols["name"]).value)
        url = linkedin_profile(str(ws.cell(r, li).value or ""))
        if not name or not url:
            continue
        if not pv.is_blank(ws.cell(r, cols["feedback"]).value):
            continue  # gave feedback: company / designation come from the form
        need_company = pv.is_blank(ws.cell(r, cols["company"]).value)
        need_role = pv.is_blank(ws.cell(r, cols["designation"]).value)
        if not (need_company or need_role):
            continue
        todo.append({"row": r, "name": name, "email": str(ws.cell(r, cols["email id"]).value or "").strip(),
                     "linkedin": url, "company": "", "designation": ""})
    return ws.title, todo


def save_results(path, sheet, results, backup_dir):
    found = [x for x in results if x["company"] or x["designation"]]
    if not found:
        log("No company / designation found - Excel file not changed")
        return
    while True:
        try:
            pv.check_excel_closed(path)
            break
        except RuntimeError:
            ask_user(f"The Excel file is open somewhere. Close it so {len(found)} row(s) can be saved.")
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    hr = pv.header_row(ws)
    cols = feedback_columns(ws, hr)
    li = linkedin_col(ws, hr)

    def is_row(r, x):  # same LinkedIn id = same person
        return linkedin_profile(str(ws.cell(r, li).value or "")).lower() == x["linkedin"].lower()

    written = 0
    for x in found:
        row = x["row"]
        # the row may have moved if someone edited the file meanwhile -> find it by its LinkedIn id
        if not is_row(row, x):
            row = next((r for r in range(hr + 1, ws.max_row + 1) if is_row(r, x)), None)
        if not row or not pv.is_blank(ws.cell(row, cols["feedback"]).value):
            continue
        changed = []
        for key in ("company", "designation"):
            cell = ws.cell(row, cols[key])
            if x[key] and pv.is_blank(cell.value):
                above = ws.cell(row - 1, cols[key])
                if row - 1 > hr and above.has_style and not cell.has_style:
                    cell._style = copy(above._style)
                cell.value = x[key]
                changed.append(key)
        if changed:
            written += 1
            log(f"  row {row}: {x['name']} -> company '{x['company']}' | designation '{x['designation']}'")
    if not written:
        log("Nothing to write - Excel file not changed")
        return
    bdir = HERE / backup_dir
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{Path(path).stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    log(f"Backup saved: {backup}")
    pv.safe_save(wb, path)
    log(f"Excel file saved - {written} row(s) got company / designation in sheet {sheet}")


# ---------------------------------------------------------------- linkedin profile

# The profile is read as plain text lines (what you see on screen), because LinkedIn
# keeps changing its page code (Nov 2026: no <h1>, no #experience any more).
#   ... <Name> / <Headline> / Message / Connect / ... / Experience /
#   <Role> / <Company · Full-time> / <Feb 2026 - Present · 10 mos> / <place> / ...
PROFILE_JS = r"""
const main = document.querySelector('main') || document.body;
const text = /\bExperience\b/.test(main.innerText || '') ? main.innerText : document.body.innerText;
return {title: document.title, url: location.href,
        lines: (text || '').split('\n').map(s => s.trim()).filter(Boolean)};
"""

EMPLOYMENT = re.compile(r"\b(full[- ]time|part[- ]time|internship|intern|contract|freelance|self[- ]employed|"
                        r"apprenticeship|trainee|seasonal|permanent)\b", re.I)
MONTH = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?"
# "Feb 2026 - Present · 10 mos", "2019 - 2021", "Oct 2014 – Present"
DATE_RANGE = re.compile(rf"^(?:{MONTH}\s+)?(?:19|20)\d\d\s*[-–—]\s*(?:present|(?:{MONTH}\s+)?(?:19|20)\d\d)\b", re.I)
# "Full-time · 3 yrs 2 mos", "3 yrs", "10 mos" (a company with several roles)
TOTAL = re.compile(r"^(?:[\w -]+·\s*)?\d+\s*(?:yrs?|mos?|years?|months?)\b(?:\s*\d+\s*(?:mos?|months?))?$", re.I)
CURRENT = re.compile(r"\bpresent\b", re.I)
# where the Experience part of the page ends
SECTION_END = re.compile(r"^(education|licenses? & certifications|skills|projects|volunteering|honors|"
                         r"courses|languages|recommendations|interests|activity|about|"
                         r"more profiles for you|people also viewed|people you may know|"
                         r"you might like|show all \d+ experiences?|show all)\b", re.I)
NOISE = re.compile(r"^(.+ and \+\d+ skills?|skills?:.*|…?see more|show (more|less).*|\d+ endorsements?)$", re.I)
PLACE = re.compile(r"·\s*(on-site|hybrid|remote)$|^(on-site|hybrid|remote)$|,\s*india\b", re.I)


def clean_company(line):
    """'Infosys · Full-time' -> 'Infosys'"""
    parts = [p.strip() for p in re.split(r"\s+·\s+", line or "") if p.strip()]
    parts = [p for p in parts if not EMPLOYMENT.fullmatch(p)]
    return parts[0] if parts else ""


def profile_name(title):
    """'(2) Seenivasan T | LinkedIn' -> 'Seenivasan T'"""
    return re.sub(r"^\(\d+\)\s*", "", title or "").split(" | ")[0].strip()


def experience_jobs(lines):
    """Lines after 'Experience' -> [(role, company, current?)], newest first."""
    jobs, block = [], []
    for line in lines:
        if SECTION_END.match(line):
            break
        if NOISE.match(line):
            continue
        if not DATE_RANGE.match(line):
            block.append(line)
            continue
        # a date line closes one role; `block` holds what came before it
        words = [x for x in block if not PLACE.search(x)]
        current = bool(CURRENT.search(line))
        total = next((i for i, x in enumerate(words) if TOTAL.match(x)), None)
        if total is not None and total >= 1:
            # company with several roles: [Company, 'Full-time · 3 yrs', (place), Role]
            company = clean_company(words[total - 1])
            role = words[-1] if len(words) > total + 1 else ""
            jobs.append((role, company, current))
        elif len(words) >= 2:
            jobs.append((words[-2], clean_company(words[-1]), current))
        elif len(words) == 1 and jobs:
            # next role at the same company as the one before
            jobs.append((words[0], jobs[-1][1], current))
        elif len(words) == 1:
            jobs.append((words[0], "", current))
        block = []
    return jobs


def from_headline(headline):
    """'DevOps Engineer at Infosys | RHCSA' -> ('DevOps Engineer', 'Infosys')"""
    first = re.split(r"\s+[|•]\s+|\s*\|\s*", headline or "")[0].strip()
    m = re.match(r"(.+?)\s+(?:at|@)\s+(.+)", first, re.I) or re.match(r"(.+?)\s*@\s*(.+)", first)
    if m:
        return m.group(1).strip(" ,-"), m.group(2).strip(" ,-")
    return "", ""


def parse_profile(info):
    """Page text -> {name, headline, jobs}."""
    lines, name = info["lines"], profile_name(info.get("title"))
    exp = next((i for i, x in enumerate(lines) if x.lower() == "experience"), None)
    top = lines[:exp] if exp is not None else lines
    # the headline is the line under the name (the last time the name is shown above Experience)
    at = max((i for i, x in enumerate(top) if name and x.lower() == name.lower()), default=None)
    headline = ""
    if at is not None:
        headline = next((x for x in top[at + 1:at + 4]
                         if not re.match(r"^[·•]|^(he|she|they)/|^\d+(st|nd|rd|th)\b|^(message|connect|follow)$", x, re.I)), "")
    jobs = experience_jobs(lines[exp + 1:]) if exp is not None else []
    return {"name": name, "headline": headline, "jobs": jobs}


def pick_job(p):
    """-> (role, company, how it was found)."""
    now = [j for j in p["jobs"] if j[2] and (j[0] or j[1])]
    if now:
        return now[0][0], now[0][1], "experience"
    role, company = from_headline(p["headline"])
    if role or company:
        return role, company, "headline"
    if p["jobs"]:
        j = p["jobs"][0]
        return "", "", f"no current job (last: {j[0]} at {j[1]})"
    return "", "", "no Experience on the profile"


def open_page(driver, url):
    """Address bar -> url. False when LinkedIn wants a sign-in and nobody can do it."""
    for _ in range(2):
        driver.get(url)
        time.sleep(3)
        if not any(k in driver.current_url for k in LINKEDIN_LOGIN):
            return True
        if nobody_at_keyboard():
            return False
        ask_user("Sign in to LinkedIn in the Arish Chrome window.")
    return False


def read_page(driver):
    end = time.time() + 20
    info = driver.execute_script(PROFILE_JS)
    while time.time() < end and "experience" not in (x.lower() for x in info["lines"]):
        # Experience is loaded only when you scroll down to it
        driver.execute_script("window.scrollBy(0, window.innerHeight);")
        time.sleep(1)
        info = driver.execute_script(PROFILE_JS)
    return parse_profile(info)


def read_job(driver, url):
    """-> (role, company, how) or None when LinkedIn wants a sign-in."""
    if not open_page(driver, url):
        return None
    p = read_page(driver)
    if not p["jobs"]:  # the full list has an address of its own
        if open_page(driver, url.rstrip("/") + "/details/experience/"):
            more = read_page(driver)
            p["jobs"] = more["jobs"]
            p["headline"] = p["headline"] or more["headline"]
    log(f"    profile '{p['name']}' | headline '{p['headline'][:70]}'"
        + (f" | jobs {p['jobs'][:2]}" if p["jobs"] else ""))
    return pick_job(p)


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="YYYY-MM-DD: only that day")
    ap.add_argument("--month", help="YYYY-MM: every row of that month")
    ap.add_argument("--limit", type=int, help="only the first N people")
    ap.add_argument("--dry-run", action="store_true", help="don't save the Excel file")
    ap.add_argument("--test", metavar="URL", help="read one LinkedIn profile and show the result (Excel not used)")
    args = ap.parse_args()

    cfg = pv.load_config()
    if args.test:
        driver = pv.open_chrome(cfg, "linkedin", args.test, "linkedin.com")
        job = read_job(driver, linkedin_profile(args.test) or args.test)
        log(f"Result: {job}")
        return
    if args.date:
        date = month_day = dt.date.fromisoformat(args.date)
    elif args.month:
        date, month_day = None, dt.date.fromisoformat(args.month + "-01")
    else:
        date, month_day = ask_scope()
    run(cfg, date, month_day, limit=args.limit, dry_run=args.dry_run)
    log("Done")


def run(cfg, date, month_day, limit=None, dry_run=False):
    """Company + designation from LinkedIn for the rows of `date` without feedback."""
    excel = cfg.get("excel", "path")
    pv.check_excel_closed(excel)

    sheet, todo = rows_to_do(excel, date, month_day)
    if limit:
        todo = todo[:limit]
    log(f"Sheet {sheet}: {len(todo)} person(s) without feedback who have a LinkedIn Id"
        + (f" on {date:%d-%b-%Y}" if date else ""))
    if not todo:
        return

    driver = pv.open_chrome(cfg, "linkedin", todo[0]["linkedin"], "linkedin.com")
    delay = cfg.getfloat("linkedin", "profile_delay", fallback=5)
    logs = HERE / "logs"
    logs.mkdir(exist_ok=True)
    progress = logs / f"company_{dt.datetime.now():%Y%m%d_%H%M%S}.json"
    errors = 0
    try:
        for i, x in enumerate(todo, 1):
            log(f"[{i}/{len(todo)}] {x['name']} | {x['linkedin']}")
            try:
                job = read_job(driver, x["linkedin"])
            except Exception as e:
                log(f"    error: {type(e).__name__}: {(str(e).splitlines() or [''])[0][:150]}")
                errors += 1
                continue
            if job is None:
                log("    LinkedIn wants a sign-in - stopped for now")
                errors += 1
                break
            role, company, how = job
            x["designation"], x["company"] = pv.clean_cell_text(role), pv.clean_cell_text(company)
            if role or company:
                log(f"    {how}: designation '{x['designation']}' | company '{x['company']}'")
            else:
                log(f"    nothing found ({how})")
            progress.write_text(json.dumps(todo, indent=2), encoding="utf-8")
            time.sleep(delay)
    except KeyboardInterrupt:
        log("Stopped - saving what was found so far")
        raise  # this step is not finished; main.py must not count it as done
    finally:
        found = sum(1 for x in todo if x["company"] or x["designation"])
        log(f"Company / designation found for {found} of {len(todo)} person(s)")
        if dry_run:
            log("Dry run - Excel file NOT saved")
        else:
            save_results(excel, sheet, todo, cfg.get("excel", "backup_dir"))
    if errors:
        raise RuntimeError(f"{errors} LinkedIn profile(s) could not be read")


if __name__ == "__main__":
    main()
