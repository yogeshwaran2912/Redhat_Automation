# Created by YOGESHWARAN TL
"""
Live status window for the nightly run (same file in Redhat_Automation and PearsonVUE_Automation).

The nightly work runs in the background with no window, so closing a window can never
stop it by accident. This window only SHOWS what it is doing:
  - the window title says RUNNING / WAITING / FINISHED / FAILED and the current step
  - every line the program writes appears here as it happens
Closing this window does NOT stop the automation (stop.bat does that).

Opened by itself by the nightly task (keep_running.py --nightly), and opened again within
10 minutes if it was closed while the work is still running. show_status.bat opens it any time.

  live_view.py                 # tonight's run (after 12 midnight: the night before)
  live_view.py --date 2026-11-19
"""
import argparse
import ctypes
import datetime as dt
import os
import re
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
LOGS = HERE / "logs"
NAME = "PEARSON VUE" if "pearson" in HERE.name.lower() else "RED HAT"
STEP_RE = re.compile(r"=+\s*(?:(\d\d-\w{3}-\d{4})\s+)?Step\s+(\d+/\d+):\s*(.+?)\s*=+")


def set_title(text):
    try:
        ctypes.windll.kernel32.SetConsoleTitleW(f"{NAME} AUTOMATION - {text}")
    except Exception:
        pass


def process_alive(pid):
    PROCESS_QUERY_LIMITED_INFORMATION, STILL_ACTIVE = 0x1000, 259
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not h:
        return False
    try:
        code = ctypes.c_ulong()
        return bool(k32.GetExitCodeProcess(h, ctypes.byref(code))) and code.value == STILL_ACTIVE
    finally:
        k32.CloseHandle(h)


def watcher_running():
    try:
        return process_alive(int((LOGS / "keep_running.pid").read_text().strip()))
    except (OSError, ValueError):
        return False


def night_date():
    now = dt.datetime.now()
    return (now.date() if now.hour >= 12 else now.date() - dt.timedelta(days=1)).isoformat()


class Follow:
    """Reads the lines added to a file since the last call (the file may not exist yet)."""

    def __init__(self, path, from_end):
        self.path, self.pos, self.rest = path, 0, b""
        if from_end:
            try:
                self.pos = path.stat().st_size
            except OSError:
                pass

    def lines(self):
        try:
            size = self.path.stat().st_size
        except OSError:
            return []
        if size < self.pos:
            self.pos = 0
        if size == self.pos:
            return []
        with open(self.path, "rb") as fh:
            fh.seek(self.pos)
            data = self.rest + fh.read()
            self.pos = fh.tell()
        *done, self.rest = data.split(b"\n")
        return [b.decode("utf-8", "replace").rstrip("\r") for b in done]


def show(line):
    print(line, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None)
    args = ap.parse_args()
    date = args.date or night_date()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    os.system("")  # turns on colours in this console
    LOGS.mkdir(exist_ok=True)
    (LOGS / "live_view.pid").write_text(str(os.getpid()))

    program_log = LOGS / f"nightly_{date.replace('-', '')}.txt"
    state_path = LOGS / f"nightly_state_{date}.txt"
    show("=" * 78)
    show(f"  {NAME} AUTOMATION - live status for the night of {date}")
    show("  Closing this window does NOT stop the automation.  To stop it: stop.bat")
    show(f"  Program output: {program_log}")
    show("=" * 78)

    # the program's output so far, then the new lines as they come
    prog = Follow(program_log, from_end=False)
    old = prog.lines()
    if len(old) > 40:
        show(f"   ... ({len(old) - 40} earlier lines in the file above)")
    for line in old[-40:]:
        show(line)
    watch = Follow(LOGS / "keep_running.txt", from_end=True)

    step = ""
    for line in old:
        m = STEP_RE.search(line)
        if m:
            step = (f"{m.group(1)} " if m.group(1) else "") + f"Step {m.group(2)} {m.group(3)}"
    waiting = False
    gone_since = None
    seen = False  # just started: the watcher may need a few seconds to begin
    while True:
        for line in prog.lines():
            show(line)
            m = STEP_RE.search(line)
            if m:
                step = (f"{m.group(1)} " if m.group(1) else "") + f"Step {m.group(2)} {m.group(3)}"
                waiting = False
        for line in watch.lines():
            show(f"\033[96m[watcher] {line}\033[0m")
            low = line.lower()
            if "waiting" in low:
                waiting = True
            elif "started the program" in low:
                waiting = False
        if watcher_running():
            gone_since, seen = None, True
            set_title(("WAITING (other program / daily jobs)" if waiting else "RUNNING")
                      + (f" - {step}" if step and not waiting else ""))
        else:
            gone_since = gone_since or time.time()
            if time.time() - gone_since > (5 if seen else 60):  # let the last lines arrive
                break
        time.sleep(1)

    for line in prog.lines():
        show(line)
    try:
        state = state_path.read_text().strip()
    except OSError:
        state = "not started"
    if state == "finished":
        set_title("FINISHED OK")
        colour, text = "\033[92m", f"FINISHED OK - the Excel file is complete for {date}"
    elif state == "running":
        set_title("NOT RUNNING NOW - the nightly task starts it again within 10 minutes")
        colour, text = "\033[93m", "Not running at the moment - the nightly task starts it again within 10 minutes"
    elif state == "not started":
        set_title(f"NOTHING RUN YET for {date}")
        colour, text = "\033[93m", f"Nothing has run yet for {date}"
    else:
        set_title(state.upper())
        colour, text = "\033[91m", f"{state.upper()} - see logs\\keep_running.txt and logs\\error_*.txt"
    show("")
    show(colour + "=" * 78)
    show("  " + text)
    show("=" * 78 + "\033[0m")
    (LOGS / "live_view.pid").unlink(missing_ok=True)
    try:
        input("\nPress Enter to close this window.")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
