# Created by YOGESHWARAN TL
"""
Runs the program and starts it again until the Excel file is complete for the date.
Nobody has to press anything: every question is answered by itself.

  Closed, Ctrl+C, killed          -> starts again, skipping the steps already finished
  A step failed (Excel open..)    -> tries again after 2, 5, 10, 15, then every 30 minutes
  stop.bat                        -> really stops it, for that date
  08:00 next morning              -> gives up and shows a message: from then on people
                                     need this computer, and a run can close their Chrome

The nightly task runs "--nightly" every 10 minutes until the morning. That starts the work
the first time, and starts it again if this watcher was killed or the computer restarted.

The Red Hat and Pearson VUE programs use the same Chrome, so they take turns: each holds
the ..\\.nightly_running lock for one run, and lets the other go while it waits to retry.

Usage
  run_all.bat                               # today
  run_all.bat --date 04-11-2026             # another day

Log: logs\\keep_running.txt
"""
import argparse
import ctypes
import datetime as dt
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

TITLE = "Red Hat automation"
PROGRAM_ARGS = []

HERE = Path(__file__).resolve().parent
LOGS = HERE / "logs"
PID_FILE = LOGS / "keep_running.pid"
PROGRAM_PID_FILE = LOGS / "keep_running_program.pid"
STOP_FILE = LOGS / "stop_requested"
ENTER_FILE = LOGS / "auto_enter.txt"
CHROME_LOCK = HERE.parent / ".nightly_running"
LOCK_OWNER = CHROME_LOCK / "owner.txt"

GIVE_UP_AT = dt.time(8, 0)
MANUAL_LIMIT = dt.timedelta(hours=4)  # started by hand, in the day: people need this computer
FAILED_DELAYS = [2, 5, 10, 15, 30]  # minutes before each new try; the last one repeats
RESTART_DELAY = 3       # seconds before starting again after a stop
QUICK_EXIT = 60         # seconds: stopping sooner than this counts as a failure, not a person
MAX_RUN = 2 * 60 * 60   # seconds: a run takes about 10 minutes, this long means it hung
EXIT_STEP_FAILED = 3    # main.py: a step failed
EXIT_BAD_OPTIONS = 2    # argparse: starting again can't help

# the Daily_automation jobs use Chrome at 00:00 and 00:06: a new run waits for them
DAILY_JOBS = (dt.time(0, 0), dt.time(0, 15))

NO_WINDOW = subprocess.CREATE_NO_WINDOW


def say(msg):
    LOGS.mkdir(exist_ok=True)
    with open(LOGS / "keep_running.txt", "a", encoding="utf-8") as fh:
        fh.write(f"[{dt.datetime.now():%d-%m-%Y %H:%M:%S}] {msg}\n")


def command_line(pid):
    """Command line of a running process, "" if there is none. Checking it means a
    Windows process that happens to get the same number is never mistaken for ours."""
    query = f'(Get-CimInstance Win32_Process -Filter "ProcessId={int(pid)}").CommandLine'
    return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", query],
                          capture_output=True, text=True, creationflags=NO_WINDOW).stdout


def read_pid(path):
    try:
        return int(path.read_text().strip())
    except (OSError, ValueError):
        return None


def running_watcher():
    """PID of the watcher already running, or None."""
    pid = read_pid(PID_FILE)
    if pid and pid != os.getpid():
        cmd = command_line(pid)
        if "--watcher" in cmd and str(HERE) in cmd:
            return pid
    return None


def kill_left_over_program():
    """The last watcher was killed but the program it started may still be working:
    two at once would fight over Chrome and the Excel file."""
    pid = read_pid(PROGRAM_PID_FILE)
    if pid:
        cmd = command_line(pid)
        if str(HERE) in cmd and "--resume" in cmd:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, creationflags=NO_WINDOW)
            say(f"Stopped a program left running by the last watcher (process {pid})")
    PROGRAM_PID_FILE.unlink(missing_ok=True)


def stop_requested():
    return STOP_FILE.exists()


def pause(seconds):
    """Sleep, but wake up at once for stop.bat."""
    end = time.time() + seconds
    while time.time() < end and not stop_requested():
        time.sleep(min(1, max(0, end - time.time())))


def write_lock_owner(program_pid=""):
    LOCK_OWNER.write_text(f"{os.getpid()} {program_pid}")


def lock_is_stale():
    """Everyone who held the lock is gone (killed, or the computer restarted)."""
    try:
        pids = LOCK_OWNER.read_text().split()
    except OSError:  # just being made, or left by something older
        try:
            return time.time() - CHROME_LOCK.stat().st_mtime > 60
        except OSError:
            return False
    return not any(("--watcher" in cmd or "--resume" in cmd) for cmd in (command_line(p) for p in pids))


def take_chrome_lock():
    """Wait for the other program's run to end -> True, or False if stop.bat was used."""
    told = False
    while not stop_requested():
        try:
            CHROME_LOCK.mkdir()
            write_lock_owner()
            return True
        except FileExistsError:
            pass
        if lock_is_stale():
            say("The other program's lock was left behind - removing it")
            shutil.rmtree(CHROME_LOCK, ignore_errors=True)
            continue
        if not told:
            say("Waiting for the other program to finish with Chrome")
            told = True
        pause(20)
    return False


def release_chrome_lock():
    try:
        if LOCK_OWNER.read_text().split()[0] == str(os.getpid()):
            shutil.rmtree(CHROME_LOCK, ignore_errors=True)
    except (OSError, IndexError):
        pass


def wait_for_daily_jobs():
    told = False
    while DAILY_JOBS[0] <= dt.datetime.now().time() < DAILY_JOBS[1] and not stop_requested():
        if not told:
            say(f"Waiting until {DAILY_JOBS[1]:%H:%M} for the Daily_automation jobs")
            told = True
        pause(20)


def next_give_up_time(now):
    deadline = dt.datetime.combine(now.date(), GIVE_UP_AT)
    return deadline if now < deadline else deadline + dt.timedelta(days=1)


def program():
    return HERE / ("main.py" if (HERE / "main.py").exists() else "main.pyc")


def run_once(date, extra, output):
    """Run the program once -> (exit code, seconds it ran). Code None = stop.bat was used.
    Its questions read Enter from a file, so it never waits for a key press."""
    cmd = [sys.executable, str(program()), *PROGRAM_ARGS, "--date", date, "--resume", *extra]
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    ENTER_FILE.write_text("\n" * 500)
    with open(ENTER_FILE, "rb") as stdin:
        if output:  # nightly: no window to close, output goes to a file
            with open(output, "ab") as out:
                proc = subprocess.Popen(cmd, cwd=HERE, env=env, stdin=stdin, stdout=out, stderr=subprocess.STDOUT,
                                        creationflags=NO_WINDOW)
        else:  # by hand: you can watch it in its own window
            proc = subprocess.Popen(cmd, cwd=HERE, env=env, stdin=stdin, creationflags=subprocess.CREATE_NEW_CONSOLE)
    PROGRAM_PID_FILE.write_text(str(proc.pid))
    write_lock_owner(proc.pid)
    say(f"Started the program (process {proc.pid})")
    started = time.time()
    try:
        while proc.poll() is None:
            if stop_requested():
                # only this process: a Chrome or Adv Pro it opened must stay open
                proc.kill()
                proc.wait()
                return None, time.time() - started
            if time.time() - started > MAX_RUN:
                say(f"The program has been running for {MAX_RUN // 3600} hours - it must have hung, stopping it")
                proc.kill()
                proc.wait()
                return EXIT_STEP_FAILED, time.time() - started
            time.sleep(1)
    finally:
        if proc.poll() is not None:
            PROGRAM_PID_FILE.unlink(missing_ok=True)
    return proc.returncode, time.time() - started


def after_stop():
    """Nothing to put back (only Pearson VUE turns an extension off)."""


def watch(date, extra, output):
    """-> (state, the result to report). State: finished, stopped or gave up."""
    now = dt.datetime.now()
    deadline = next_give_up_time(now) if output else now + MANUAL_LIMIT
    kill_left_over_program()
    failed = restarts = 0
    while True:
        if stop_requested():
            return "stopped", "Stopped by stop.bat"
        if dt.datetime.now() >= deadline:
            return "gave up", (f"Gave up at {deadline:%H:%M}: the Excel file is NOT complete for {date} "
                               f"({failed} failed, {restarts} stopped) - see logs\\error_*.txt")
        if not take_chrome_lock():
            return "stopped", "Stopped by stop.bat"
        try:
            wait_for_daily_jobs()
            if stop_requested():
                return "stopped", "Stopped by stop.bat"
            code, seconds = run_once(date, extra, output)
        finally:
            release_chrome_lock()
        if code is None:
            after_stop()
            return "stopped", "Stopped by stop.bat"
        if code == 0:
            return "finished", f"Finished: the Excel file is complete for {date}"
        if code == EXIT_BAD_OPTIONS:
            return "gave up", f"Gave up: the program refused the options {extra}"
        if code == EXIT_STEP_FAILED or seconds < QUICK_EXIT:
            minutes = FAILED_DELAYS[min(failed, len(FAILED_DELAYS) - 1)]
            failed += 1
            say(f"Not complete (exit code {code}, try {failed}) - trying again in {minutes} minutes")
            pause(minutes * 60)
        else:
            restarts += 1
            say(f"The program was stopped (exit code {code}) - starting again (time {restarts})")
            pause(RESTART_DELAY)


def show_message(text):
    MB_ICONINFORMATION, MB_SETFOREGROUND, MB_TOPMOST = 0x40, 0x10000, 0x40000
    ctypes.windll.user32.MessageBoxW(None, text, TITLE, MB_ICONINFORMATION | MB_SETFOREGROUND | MB_TOPMOST)


def state_file(date):
    return LOGS / f"nightly_state_{date}.txt"


def watcher_main(args, extra):
    LOGS.mkdir(exist_ok=True)
    PID_FILE.write_text(str(os.getpid()))
    STOP_FILE.unlink(missing_ok=True)
    say(f"Watcher started for {args.date}")
    state, result = "gave up", "Gave up: the watcher itself failed - see logs\\keep_running.txt"
    try:
        state, result = watch(args.date, extra, args.output)
    except Exception as e:
        say(f"Watcher error: {type(e).__name__}: {e}")
    finally:
        say(result)
        if args.output:
            state_file(args.date).write_text(state)
        PID_FILE.unlink(missing_ok=True)
        STOP_FILE.unlink(missing_ok=True)
    if not args.output or state == "gave up":
        show_message(result)


def start_watcher(date, extra, output=None):
    """Start the watcher in the background, with no window that could be closed."""
    cmd = [sys.executable, str(Path(__file__).resolve()), "--watcher", "--date", date]
    if output:
        cmd += ["--output", str(output)]
    cmd += extra
    flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    kwargs = dict(cwd=HERE, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:  # keeps it alive if Task Scheduler ends the task
        subprocess.Popen(cmd, creationflags=flags | subprocess.CREATE_BREAKAWAY_FROM_JOB, **kwargs)
    except OSError:
        subprocess.Popen(cmd, creationflags=flags, **kwargs)


def open_live_view(date):
    """Shows tonight's run in a window: title = RUNNING / WAITING / FINISHED + the step, and the
    program's output live. Only a viewer: closing it never stops the work. Opened again by the
    next 10-minute check if it was closed. Never lets a problem here stop the nightly work."""
    try:
        pid = read_pid(LOGS / "live_view.pid")
        if pid and "live_view" in command_line(pid):
            return  # already open
        script = HERE / ("live_view.py" if (HERE / "live_view.py").exists() else "live_view.pyc")
        if not script.exists():
            return
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 4  # SW_SHOWNOACTIVATE: visible, but does not take the keyboard from Adv Pro
        cmd = [sys.executable, str(script), "--date", date]
        kwargs = dict(cwd=HERE, startupinfo=si, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        try:  # keeps it open when Task Scheduler ends the task
            subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_CONSOLE | subprocess.CREATE_BREAKAWAY_FROM_JOB, **kwargs)
        except OSError:
            subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_CONSOLE, **kwargs)
    except Exception as e:
        say(f"Could not open the live status window: {type(e).__name__}: {e}")


def nightly_main():
    """Started every 10 minutes by the task. Does nothing unless work is left for tonight."""
    now = dt.datetime.now()
    # the date stays fixed: after 12 midnight this is still the work of the night before
    date = (now.date() if now.hour >= 12 else now.date() - dt.timedelta(days=1)).isoformat()
    state = state_file(date)
    if running_watcher():
        open_live_view(date)
        return
    if state.exists() and state.read_text().strip() != "running":
        return  # finished, stopped or gave up
    if STOP_FILE.exists():
        STOP_FILE.unlink(missing_ok=True)
        if state.exists():  # tonight's work had started, and its watcher was not there to see stop.bat
            state.write_text("stopped")
            say(f"Stopped by stop.bat - nothing more for {date}")
            return
    if not state.exists():  # first start tonight: every step is done again
        (LOGS / f"progress_{date}.json").unlink(missing_ok=True)
        state.write_text("running")
    start_watcher(date, [], LOGS / f"nightly_{date.replace('-', '')}.txt")
    open_live_view(date)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="DD-MM-YYYY or YYYY-MM-DD (default: today)")
    ap.add_argument("--nightly", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--output", help=argparse.SUPPRESS)
    ap.add_argument("--watcher", action="store_true", help=argparse.SUPPRESS)
    args, extra = ap.parse_known_args()

    LOGS.mkdir(exist_ok=True)
    if args.watcher:
        watcher_main(args, extra)
        return
    if args.nightly:
        nightly_main()
        return

    other = running_watcher()
    if other:
        print(f"The automation is already running (process {other}). Use stop.bat to stop it.")
        sys.exit(1)
    if args.date:
        from redhat_automation import parse_date
        date = parse_date(args.date).isoformat()
    else:
        date = dt.date.today().isoformat()
    (LOGS / f"progress_{date}.json").unlink(missing_ok=True)  # a new run does every step
    start_watcher(date, extra)
    print(f"\nStarted for {date}. The program opens in its own window, and needs no key presses.")
    print("If it is closed or stopped it starts again, until the Excel file is complete.")
    print("To really stop it, run stop.bat. You can close this window.")


if __name__ == "__main__":
    main()
