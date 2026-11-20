"""
Google reviews (last 24 hours)  ->  Excel "Google Review" + its "Remark"

Shared by Redhat_Automation and PearsonVUE_Automation (the same file is in both folders).

Steps
  1. Chrome ("Advantage pro - Chrome" profile): Google "advantagepro.in" -> the
     Advantage Pro business on Google Maps -> Reviews -> Sort -> Newest.
  2. Read the reviews of the last 24 hours ([review] hours): name, stars, message.
  3. Match each reviewer's name to a candidate of that date in the Excel sheet
     (then to the candidates of the 7 days before, for people who review a day late).
  4. Fill: Google Review <- "y", Remark (the one next to it) <- their message,
     or "5 star" (the stars they gave) when they wrote no message.
     Only empty cells are filled. Nobody is filled in without a real review.
  5. Back up the Excel file and save; log how many of the day's candidates reviewed.

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


def match_reviews(ws, hr, h, date, reviews, back_days=7):
    """-> ([(row, review)], unmatched reviews, rows of the day)"""
    g, rem = review_columns(ws, hr)
    nc = name_col(ws, hr)
    rows = []
    for r in range(hr + 1, h.last_data_row(ws, hr) + 1):
        d = h.cell_date(ws.cell(r, 1).value)
        name = h.clean_cell_text(ws.cell(r, nc).value)
        if d and name and date - dt.timedelta(days=back_days) <= d <= date:
            rows.append({"row": r, "date": d, "name": name, "done": not h.is_blank(ws.cell(r, g).value)})
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


def update_excel(cfg, date, reviews, h, dry_run=False):
    path = cfg.get("excel", "path")
    if not dry_run:
        h.check_excel_closed(path)
    wb = openpyxl.load_workbook(path)
    ws = h.month_sheet(wb, date)
    hr = h.header_row(ws)
    g, rem = review_columns(ws, hr)
    found, unmatched, day_rows = match_reviews(ws, hr, h, date, reviews,
                                               cfg.getint("review", "match_days", fallback=7))
    written = 0
    for x, rv in found:
        r = x["row"]
        if not h.is_blank(ws.cell(r, g).value):
            h.log(f"  row {r} {x['name']}: Google Review already filled ({rv['name']})")
            continue
        for col, value in ((g, "y"), (rem, remark_for(rv))):
            cell = ws.cell(r, col)
            if value and h.is_blank(cell.value):
                above = ws.cell(r - 1, col)
                if r - 1 > hr and above.has_style and not cell.has_style:
                    cell._style = copy(above._style)
                cell.value = value
        written += 1
        x["done"] = True
        h.log(f"  row {r} {x['name']} ({x['date']:%d-%b}) <- review by '{rv['name']}': y | {remark_for(rv)[:70]}")
    if unmatched:
        h.log(f"{len(unmatched)} review(s) did not match a candidate (other courses/exams, or a different name on Google):")
        for rv in unmatched:
            h.log(f"    - {rv['name']} | {rv['stars']} star | {rv['when']}")
    done = sum(1 for x in day_rows if x["done"])
    if day_rows:
        h.log(f"Google reviews for {date:%d-%b-%Y}: {done} of {len(day_rows)} candidate(s) "
              f"({100 * done // len(day_rows)}%)")
    if dry_run:
        h.log(f"Dry run - Excel file NOT saved ({written} row(s) would be filled)")
        return
    if not written:
        h.log("Nothing new to write - Excel file not changed")
        return
    bdir = h.HERE / cfg.get("excel", "backup_dir")
    bdir.mkdir(exist_ok=True)
    backup = bdir / f"{Path(path).stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    h.log(f"Backup saved: {backup}")
    h.safe_save(wb, path)
    h.log(f"Excel file saved - Google Review filled for {written} row(s) in sheet {ws.title}")


def run(cfg, date, h, dry_run=False):
    """h = the project's helper module (redhat_automation, or PearsonVUE main)."""
    if not dry_run:
        h.check_excel_closed(cfg.get("excel", "path"))
    # an earlier date (catch-up): look back far enough to reach that day
    since = (dt.datetime.now() - dt.datetime(date.year, date.month, date.day)).total_seconds() / 3600
    reviews = read_reviews(cfg, h, max(cfg.getfloat("review", "hours", fallback=24), round(since + 1)))
    if not reviews:
        h.log("No new Google reviews - nothing to write")
        return
    update_excel(cfg, date, reviews, h, dry_run=dry_run)


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
