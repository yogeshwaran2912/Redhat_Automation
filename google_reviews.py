# Created by YOGESHWARAN TL
"""
Google reviews (last 24 hours)  ->  Excel "Google Review" + its "Remark"

Shared by Redhat_Automation and PearsonVUE_Automation (the same file is in both folders).

Steps
  1. Chrome ("Advantage pro - Chrome" profile): Google "advantagepro.in" -> the
     Advantage Pro business on Google Maps -> Reviews -> Sort -> Newest.
  2. Read the reviews of the last 24 hours ([review] hours): name, stars, message.
  3. Match each reviewer's name to a candidate of that date in the Excel sheet
     (then to the candidates of the 7 days before, for people who review a day late).
  4. Fill Google Review:  Y  + Remark (next to it) <- their message, or "5 star"
     No review yet -> left blank.  A review that fits more than one candidate -> left blank,
     listed in the report to check by hand.
     "y" / "Received" (real reviews) and hand-typed values are never changed;
     a blank or "Reminder Sent" cell becomes Y when the real review appears.
     Nobody gets Y without a real review.
     Old values written by earlier versions are cleaned on the month sheet:
     "Received" -> Y, "Pending" / "Unable to Verify" -> blank.
  5. Back up the Excel file and save; every change -> logs/review_history.csv;
     daily report -> reports/Google_Reviews_DD-MM-YYYY.txt (and in the log).

Usage
  python google_reviews.py                     # asks the date (Enter = today)
  python google_reviews.py --date 2026-11-18   # a specific day
  python google_reviews.py --dry-run           # don't save the Excel file
  python google_reviews.py --list              # only show the reviews read from Google
"""
import argparse
import datetime as dt
import re
import shutil
import time
import urllib.parse
from copy import copy
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------- google maps

AGE_UNITS = {"minute": 1 / 60, "min": 1 / 60, "hour": 1, "day": 24, "week": 168, "month": 720, "year": 8760}

FIND_CID_JS = r"""
for (const a of document.querySelectorAll('a[href*="/maps/place/"], a[href*="cid="], a[href*="ludocid="]')) {
  const m = a.href.match(/!1s0x[0-9a-f]+:(0x[0-9a-f]+)/i) || a.href.match(/[?&](?:lu)?(?:do)?cid=(\d+)/);
  if (m) return m[1];
}
return null;
"""

CLICK_REVIEWS_TAB_JS = r"""
const t = [...document.querySelectorAll('[role=tab], button')].find(b =>
  /^reviews\b/i.test((b.getAttribute('aria-label') || '').trim()) || /^\s*reviews\s*$/i.test(b.innerText || ''));
if (t) { t.click(); return true; }
return false;
"""

CLICK_SORT_JS = r"""
const b = [...document.querySelectorAll('button')].find(b => /sort reviews|^sort$/i.test((b.getAttribute('aria-label') || b.innerText || '').trim()));
if (b) { b.click(); return true; }
return false;
"""

CLICK_NEWEST_JS = r"""
const e = [...document.querySelectorAll('[role=menuitemradio], [role=menuitem], [role=option]')].find(e => /newest/i.test(e.innerText));
if (e) { e.click(); return true; }
return false;
"""

SCROLL_JS = r"""
const r = document.querySelector('div[data-review-id]');
if (!r) return 0;
let pane = r.parentElement;
while (pane && !(pane.scrollHeight > pane.clientHeight + 20 && /auto|scroll/.test(getComputedStyle(pane).overflowY)))
  pane = pane.parentElement;
if (pane) pane.scrollTop = pane.scrollHeight;
for (const b of document.querySelectorAll('div[data-review-id] button'))  // "More" -> the whole message
  if (/^(see more|more)$/i.test((b.getAttribute('aria-label') || b.innerText || '').trim())) b.click();
return document.querySelectorAll('div[data-review-id][aria-label]').length;
"""

READ_REVIEWS_JS = r"""
return [...document.querySelectorAll('div[data-review-id][aria-label]')].map(r => {
  const stars = r.querySelector('[role=img][aria-label*="star" i]');
  const lines = (r.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
  const when = (r.querySelector('.rsqaWe') || {}).innerText || lines.find(s => /\bago\b/i.test(s)) || '';
  let text = [...r.querySelectorAll('.wiI7pd')].map(e => e.innerText).filter(Boolean)[0];
  if (text === undefined) {  // page code changed: the lines after the time, up to the buttons / owner's reply
    const k = lines.indexOf(when), out = [];
    for (const s of lines.slice(k + 1)) {
      if (/^(like|share|response from the owner|helpful)\b/i.test(s)) break;
      if (!/^(new|more|see more|\d+)$/i.test(s)) out.push(s);
    }
    text = out.join(' ');
  }
  return {id: r.getAttribute('data-review-id'), name: r.getAttribute('aria-label'),
          stars: stars ? stars.getAttribute('aria-label') : '', when, text: text || ''};
});
"""


def age_hours(when):
    """'11 hours ago' -> 11, 'Edited a day ago' -> 24, '' -> None"""
    m = re.search(r"\b(a|an|\d+)\s+(minute|min|hour|day|week|month|year)s?\s+ago", when or "", re.I)
    if not m:
        return None
    n = 1 if m.group(1).lower() in ("a", "an") else int(m.group(1))
    return n * AGE_UNITS[m.group(2).lower()]


def star_count(label):
    m = re.search(r"(\d)(?:\.\d)?\s*star", label or "", re.I)
    return int(m.group(1)) if m else None


def save_config_value(config_path, section, key, value):
    """Update one key in config.ini without losing the comments."""
    lines = Path(config_path).read_text(encoding="utf-8").splitlines()
    current = None
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            current = s[1:-1]
        elif current == section and re.match(rf"{key}\s*=", s):
            lines[i] = f"{key} = {value}"
            break
    Path(config_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def find_maps_url(driver, cfg, h):
    """Google '<search>' -> the business's Google Maps link (saved in config.ini)."""
    url = cfg.get("review", "maps_url", fallback="").strip()
    if url:
        return url
    search = cfg.get("review", "search", fallback="advantagepro.in")
    driver.get("https://www.google.com/search?hl=en&q=" + urllib.parse.quote_plus(search))
    cid = None
    for _ in range(15):
        cid = driver.execute_script(FIND_CID_JS)
        if cid:
            break
        time.sleep(1)
    if not cid:
        raise RuntimeError(f"The business was not found on Google for '{search}' - "
                           "put its Google Maps link in config.ini [review] maps_url")
    cid = str(int(cid, 16)) if cid.lower().startswith("0x") else cid
    url = f"https://www.google.com/maps?cid={cid}&hl=en"
    save_config_value(h.HERE / "config.ini", "review", "maps_url", url)
    h.log(f"Saved the Google Maps link in config.ini: {url}")
    return url


def wait_js(driver, js, timeout=15):
    end = time.time() + timeout
    while True:
        r = driver.execute_script(js)
        if r or time.time() > end:
            return r
        time.sleep(1)


_CACHE = {"hours": -1, "reviews": []}  # one Google visit per run, also when several dates are done


def read_reviews(cfg, h, hours=None):
    """-> reviews of the last `hours` ([review] hours), newest first: [{name, stars, text, hours}]"""
    hours = hours or cfg.getfloat("review", "hours", fallback=24)
    if _CACHE["hours"] >= hours:
        out = [r for r in _CACHE["reviews"] if r["hours"] <= hours]
        h.log(f"{len(out)} Google review(s) in the last {hours:g} hours (read earlier in this run)")
        return out
    driver = h.open_chrome(cfg, "review", "https://www.google.com/search?hl=en&q="
                           + urllib.parse.quote_plus(cfg.get("review", "search", fallback="advantagepro.in")),
                           "google.com")
    driver.get(find_maps_url(driver, cfg, h))
    if not wait_js(driver, CLICK_REVIEWS_TAB_JS, 25):
        raise RuntimeError("The 'Reviews' tab was not found on Google Maps")
    time.sleep(3)
    if not wait_js(driver, CLICK_SORT_JS, 15):
        raise RuntimeError("The 'Sort' button was not found in the reviews")
    time.sleep(1.5)
    if not wait_js(driver, CLICK_NEWEST_JS, 10):
        raise RuntimeError("'Newest' was not found in the Sort menu")
    h.log("Google Maps: Reviews -> Sort -> Newest")
    time.sleep(3)

    reviews, last_count, still = [], -1, 0
    for _ in range(40):  # load more until the reviews are older than the limit
        count = driver.execute_script(SCROLL_JS)
        time.sleep(1.5)
        reviews = driver.execute_script(READ_REVIEWS_JS)
        ages = [age_hours(r["when"]) for r in reviews]
        known = [a for a in ages if a is not None]
        if known and known[-1] > hours:
            break
        still = still + 1 if count == last_count else 0
        if still >= 3:  # nothing more loads
            break
        last_count = count
    driver.execute_script(SCROLL_JS)  # opens "More" on the last ones too
    time.sleep(1)
    reviews = driver.execute_script(READ_REVIEWS_JS)

    skip = [s.strip().lower() for s in cfg.get("review", "skip_reviewers", fallback="Advantage pro").split(",") if s.strip()]
    out, seen = [], set()
    for r in reviews:
        a = age_hours(r["when"])
        if a is None or a > hours or r["id"] in seen:
            continue
        seen.add(r["id"])
        if (r["name"] or "").strip().lower() in skip:
            continue  # the business's own account
        out.append({"name": (r["name"] or "").strip(), "stars": star_count(r["stars"]),
                    "text": re.sub(r"\s+", " ", r["text"] or "").strip(), "when": r["when"], "hours": a})
    h.log(f"{len(out)} Google review(s) in the last {hours:g} hours")
    for r in out:
        h.log(f"  {r['name']} | {r['stars']} star | {r['when']} | {r['text'][:70]}")
    _CACHE.update(hours=hours, reviews=out)
    return out


# ---------------------------------------------------------------- matching

def name_tokens(s):
    return re.findall(r"[a-z]+", str(s or "").lower())


def name_score(a, b):
    """How well name a fits name b ("Aruna N" vs "Aruna Nainappan"). 0 = no match."""
    f, r = name_tokens(a), name_tokens(b)
    if not f or not r or not any(len(t) >= 3 for t in f):
        return 0

    def tok_ok(t):
        if len(t) == 1:  # initial
            return any(x.startswith(t) for x in r)
        # a longer name must add 3+ letters: "Gopal" = "Gopalakrishnan", but "Arun" is not "Aruna"
        return any(t == x or (len(t) >= 3 and len(x) >= 3 and abs(len(t) - len(x)) >= 3
                              and (x.startswith(t) or t.startswith(x))) for x in r)

    if all(tok_ok(t) for t in f):
        return 100 + sum(len(t) for t in f if len(t) >= 3)
    fj, rj = "".join(f), "".join(r)
    if len(fj) >= 6 and (fj in rj or rj in fj):  # "vairamuthu" vs "vaira muthu"
        return 90
    return 0


def strict_score(a, b):
    """Every word of one name is a whole word (or the initial) of the other:
    "Sadhan Kumar K" = "K Sadhan Kumar", but "Arun N" is NOT "Aruna Nainappan"."""
    f, r = name_tokens(a), name_tokens(b)
    if len(f) > len(r):
        f, r = r, f
    if not f or not any(len(t) >= 3 for t in f):
        return 0
    ok = all(t in r if len(t) > 1 else any(x.startswith(t) for x in r) for t in f)
    # the longer name's real words must be there too (as words or initials)
    back = all(x in f or any(len(t) == 1 and x.startswith(t) for t in f) for x in r if len(x) > 1)
    return 100 + sum(len(t) for t in f) if ok and back else 0


def match_score(review_name, excel_name, strict=False):
    """Either name may be the short one ("Madhuri Gopal" / "Madhuri Gopalakrishnan").
    Google names that are only one word must match a whole word of the Excel name.
    strict: whole words only (used for the days before, where a near name is more likely another person)."""
    if strict:
        return strict_score(review_name, excel_name)
    rt = name_tokens(review_name)
    if len([t for t in rt if len(t) >= 3]) < 1:
        return 0
    if len(rt) == 1:
        return 100 if rt[0] in name_tokens(excel_name) and len(rt[0]) >= 4 else 0
    return max(name_score(review_name, excel_name), name_score(excel_name, review_name))


def review_columns(ws, hr):
    """-> (Google Review column, the Remark column right after it)"""
    g = rem = None
    for c in range(1, ws.max_column + 1):
        head = str(ws.cell(hr, c).value or "").strip().lower()
        if head == "google review":
            g = c
        elif g and head == "remark" and not rem:
            rem = c
            break
    if not g or not rem:
        raise RuntimeError(f"Columns 'Google Review' + 'Remark' not found in sheet {ws.title}")
    return g, rem


def name_col(ws, hr):
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(hr, c).value or "").strip().lower() == "name":
            return c
    raise RuntimeError(f"Column 'Name' not found in sheet {ws.title}")


def remark_for(r):
    return r["text"] or (f"{r['stars']} star" if r["stars"] else "")


def match_reviews(ws, hr, h, date, reviews, back_days=7, ambiguous=None):
    """-> ([(row, review)], unmatched reviews, rows of the day)
    ambiguous: list that gets (rows, review) for reviews fitting more than one candidate."""
    g, rem = review_columns(ws, hr)
    nc = name_col(ws, hr)
    rows = []
    for r in range(hr + 1, h.last_data_row(ws, hr) + 1):
        d = h.cell_date(ws.cell(r, 1).value)
        name = h.clean_cell_text(ws.cell(r, nc).value)
        if d and name and date - dt.timedelta(days=back_days) <= d <= date:
            rows.append({"row": r, "date": d, "name": name, "done": review_state(ws.cell(r, g).value, h) == "genuine"})
    day_rows = [x for x in rows if x["date"] == date]
    used, found, unmatched = set(), [], []
    for rv in reviews:
        hit = None
        for strict, group in ((False, day_rows), (True, [x for x in rows if x["date"] != date])):  # that day first
            scored = sorted(((match_score(rv["name"], x["name"], strict), x["date"].toordinal(), x["row"], x)
                             for x in group if x["row"] not in used), key=lambda s: s[:3], reverse=True)
            scored = [s for s in scored if s[0] >= 90]
            if not scored:
                continue
            same = [s for s in scored if s[0] == scored[0][0] and s[1] == scored[0][1]]
            if len(same) > 1:
                h.log(f"  ? review '{rv['name']}' fits more than one row "
                      f"({', '.join(s[3]['name'] for s in same)}) - fill it by hand")
                if ambiguous is not None:
                    ambiguous.append(([s[3] for s in same], rv))
                hit = "ambiguous"
                break
            hit = scored[0][3]
            break
        if hit == "ambiguous":
            continue
        if not hit:
            unmatched.append(rv)
            continue
        used.add(hit["row"])
        found.append((hit, rv))
    return found, unmatched, day_rows


# Google Review column values
#   genuine (never changed): y / yes / Received
#   status (replaced by Received when that person's real review appears):
#     Pending, Reminder Sent, Unable to Verify
#   anything else typed by hand is left alone
RECEIVED, PENDING, REMINDER, UNVERIFIED = "Y", "Pending", "Reminder Sent", "Unable to Verify"
OLD_RECEIVED = "Received"  # written by the version of 19-Nov-2026, now Y
GENUINE = {"y", "yes", "received"}
STATUSES = {PENDING.lower(), REMINDER.lower(), UNVERIFIED.lower()}


def review_state(value, h):
    """-> 'blank' | 'genuine' | 'status' | 'other'"""
    if h.is_blank(value):
        return "blank"
    v = str(value).strip().lower()
    return "genuine" if v in GENUINE else "status" if v in STATUSES else "other"


def history_rows(path, rows):
    """Append to logs/review_history.csv (one line per changed cell, date-wise)."""
    import csv
    new = not path.exists()
    with open(path, "a", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["Run time", "Exam date", "Sheet", "Row", "Candidate", "Old value", "New value",
                        "Remark", "Google reviewer"])
        w.writerows(rows)


def update_excel(cfg, date, reviews, h, dry_run=False):
    path = cfg.get("excel", "path")
    if not dry_run:
        h.check_excel_closed(path)
    wb = openpyxl.load_workbook(path)
    ws = h.month_sheet(wb, date)
    hr = h.header_row(ws)
    g, rem = review_columns(ws, hr)
    ambiguous = []
    found, unmatched, day_rows = match_reviews(ws, hr, h, date, reviews,
                                               cfg.getint("review", "match_days", fallback=7), ambiguous)
    now = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}"
    history, new_reviews = [], []

    def put(r, col, value):
        cell = ws.cell(r, col)
        above = ws.cell(r - 1, col)
        if r - 1 > hr and above.has_style and not cell.has_style:
            cell._style = copy(above._style)
        cell.value = value

    def set_status(x, value, remark="", reviewer=""):
        r = x["row"]
        old = h.clean_cell_text(ws.cell(r, g).value)
        put(r, g, value)
        if remark and h.is_blank(ws.cell(r, rem).value):
            put(r, rem, remark)
        history.append([now, f"{x['date']:%d-%m-%Y}", ws.title, r, x["name"], old, value, remark, reviewer])

    # 1. real reviews -> Received (only blank / status cells; y / Received / hand text never touched)
    for x, rv in found:
        r = x["row"]
        state = review_state(ws.cell(r, g).value, h)
        if state in ("genuine", "other"):
            h.log(f"  row {r} {x['name']}: Google Review already filled "
                  f"('{h.clean_cell_text(ws.cell(r, g).value)}', review by {rv['name']})")
            continue
        set_status(x, RECEIVED, remark_for(rv), rv["name"])
        x["done"] = True
        new_reviews.append((x, rv))
        h.log(f"  row {r} {x['name']} ({x['date']:%d-%b}) <- review by '{rv['name']}': "
              f"{RECEIVED} | {remark_for(rv)[:70]}")

    # 2. a review that fits more than one candidate -> left blank, check by hand (in the report)
    for xs, rv in ambiguous:
        h.log(f"  review by '{rv['name']}' fits more than one candidate ({', '.join(x['name'] for x in xs)})"
              " - left blank, check by hand")

    # 3. the day's candidates with no review yet -> left blank
    waiting = sum(1 for x in day_rows if review_state(ws.cell(x["row"], g).value, h) == "blank")
    if waiting:
        h.log(f"  {waiting} candidate(s) of {date:%d-%b} with no Google review yet - left blank")

    # 4. old values from the earlier version: Received -> Y, Pending / Unable to Verify -> blank
    nc = name_col(ws, hr)
    cleaned = 0
    for r in range(hr + 1, h.last_data_row(ws, hr) + 1):
        old = h.clean_cell_text(ws.cell(r, g).value)
        new = {OLD_RECEIVED.lower(): RECEIVED, PENDING.lower(): None, UNVERIFIED.lower(): None}.get(old.lower(), old)
        if new != old:
            ws.cell(r, g).value = new
            cleaned += 1
            d = h.cell_date(ws.cell(r, 1).value)
            history.append([now, f"{d:%d-%m-%Y}" if d else "", ws.title, r,
                            h.clean_cell_text(ws.cell(r, nc).value), old, new or "", "", ""])
    if cleaned:
        h.log(f"  {cleaned} old cell(s) cleaned: '{OLD_RECEIVED}' -> {RECEIVED}, '{PENDING}' / '{UNVERIFIED}' -> blank")

    if unmatched:
        h.log(f"{len(unmatched)} review(s) did not match a candidate (other courses/exams, or a different name on Google):")
        for rv in unmatched:
            h.log(f"    - {rv['name']} | {rv['stars']} star | {rv['when']}")

    report = daily_report(ws, hr, h, g, date, new_reviews, unmatched, ambiguous)
    for line in report:
        h.log(line)

    if dry_run:
        h.log(f"Dry run - Excel file NOT saved ({len(history)} cell(s) would change)")
        return
    rdir = h.HERE / "reports"
    rdir.mkdir(exist_ok=True)
    rpath = rdir / f"Google_Reviews_{date:%d-%m-%Y}.txt"
    rpath.write_text("\n".join(report) + "\n", encoding="utf-8")
    h.log(f"Report saved: {rpath}")
    if not history:
        h.log("Nothing new to write - Excel file not changed")
        return
    bdir = h.HERE / cfg.get("excel", "backup_dir")
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{Path(path).stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    h.log(f"Backup saved: {backup}")
    h.safe_save(wb, path)
    (h.HERE / "logs").mkdir(exist_ok=True)
    history_rows(h.HERE / "logs" / "review_history.csv", history)
    h.log(f"Excel file saved - {len(history)} Google Review cell(s) changed in sheet {ws.title} "
          f"(history: logs\\review_history.csv)")


def daily_report(ws, hr, h, g, date, new_reviews, unmatched, ambiguous):
    """Counts for the day and for the month sheet (up to the date), from the Google Review column."""
    nc = name_col(ws, hr)

    def counts(test):
        c = {"total": 0, "genuine": 0, "pending": 0, "reminder": 0, "unverified": 0}
        for r in range(hr + 1, h.last_data_row(ws, hr) + 1):
            d = h.cell_date(ws.cell(r, 1).value)
            if not d or not h.clean_cell_text(ws.cell(r, nc).value) or not test(d):
                continue
            c["total"] += 1
            v = ws.cell(r, g).value
            s = review_state(v, h)
            low = h.clean_cell_text(v).lower()
            if s == "genuine":
                c["genuine"] += 1
            elif low == REMINDER.lower():
                c["reminder"] += 1
            elif low == UNVERIFIED.lower():
                c["unverified"] += 1
            elif s in ("blank", "status"):
                c["pending"] += 1
        return c

    def pct(c):
        return f"{100 * c['genuine'] // c['total']}%" if c["total"] else "-"

    day, month = counts(lambda d: d == date), counts(lambda d: d <= date)
    out = [f"===== Google Review report - {date:%d-%b-%Y} (sheet {ws.title}) =====",
           f"                     {'Day':>6} {'Month':>7}",
           f"Total candidates     {day['total']:>6} {month['total']:>7}",
           f"Genuine reviews      {day['genuine']:>6} {month['genuine']:>7}",
           f"Pending              {day['pending']:>6} {month['pending']:>7}",
           f"Reminder Sent        {day['reminder']:>6} {month['reminder']:>7}",
           f"Unable to Verify     {day['unverified']:>6} {month['unverified']:>7}",
           f"Review percentage    {pct(day):>6} {pct(month):>7}",
           f"New reviews received this run: {len(new_reviews)}"]
    out += [f"   + {x['name']} ({x['date']:%d-%b}) - {rv['stars']} star - {remark_for(rv)[:60]}"
            for x, rv in new_reviews]
    if unmatched:
        out.append(f"Reviews not matched to a candidate: {len(unmatched)}")
        out += [f"   ? {rv['name']} - {rv['stars']} star - {rv['when']}" for rv in unmatched]
    if ambiguous:
        out.append(f"Reviews fitting more than one candidate (check by hand): {len(ambiguous)}")
        out += [f"   ? {rv['name']} -> {', '.join(x['name'] for x in xs)}" for xs, rv in ambiguous]
    return out


def run(cfg, date, h, dry_run=False):
    """h = the project's helper module (redhat_automation, or PearsonVUE main)."""
    if not dry_run:
        h.check_excel_closed(cfg.get("excel", "path"))
    # an earlier date (catch-up): look back far enough to reach that day
    since = (dt.datetime.now() - dt.datetime(date.year, date.month, date.day)).total_seconds() / 3600
    reviews = read_reviews(cfg, h, max(cfg.getfloat("review", "hours", fallback=24), round(since + 1)))
    if not reviews:
        h.log("No new Google reviews - still making the report")
    update_excel(cfg, date, reviews or [], h, dry_run=dry_run)


# ---------------------------------------------------------------- main

def helpers():
    try:
        import redhat_automation as h  # Redhat_Automation folder
    except ImportError:
        import main as h  # PearsonVUE_Automation folder
    return h


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="YYYY-MM-DD or DD-MM-YYYY (default: asks, Enter = today)")
    ap.add_argument("--dry-run", action="store_true", help="don't save the Excel file")
    ap.add_argument("--list", action="store_true", help="only show the reviews read from Google")
    args = ap.parse_args()
    h = helpers()
    cfg = h.load_config()
    tabs = h.chrome_tab_ids(cfg)
    try:
        if args.list:
            read_reviews(cfg, h)
            return
        date = h.parse_date(args.date) if args.date else h.ask_date()
        h.log(f"Date: {date:%d-%b-%Y}   Excel: {cfg.get('excel', 'path')}")
        run(cfg, date, h, dry_run=args.dry_run)
    finally:
        h.step_cleanup(cfg, tabs)  # close the Google Maps tab this opened
    h.log("Done")


if __name__ == "__main__":
    main()
