"""
Everything for one day, in one go  ->  "redhat exam 2026.xlsx"

  Step 1  Exams     Red Hat Exam Scheduler (Completed, locations 2093, 2337, 2394)
                    + Adv Pro "Our Student" check -> new rows
  Step 2  Feedback  ADV PRO FEEDBACK (Responses) Google Sheet -> company,
                    designation, feedback "y", remark
  Step 3  LinkedIn  Outlook contact card -> LinkedIn profile URL

Each step saves the Excel file (with a backup) before the next one starts, so
the next step sees the new rows. If a step fails, the others still run.

Usage
  python main.py                      # asks for the date (Enter = today)
  python main.py --date 2026-10-07    # a specific day
  python main.py --skip-exams         # leave out a step (also --skip-feedback, --skip-linkedin)
  python main.py --skip-advpro        # don't check Adv Pro ("Our Student" left blank)
  python main.py --ask                # LinkedIn: wait for you when a profile isn't found
  python main.py --dry-run            # don't save the Excel file (steps 2-3 then
                                      # won't see the rows step 1 would add)
"""
import argparse
import datetime as dt
import traceback

import feedback_from_sheet
import linkedin_from_outlook
import redhat_automation
from redhat_automation import HERE, ask_date, load_config, log


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="YYYY-MM-DD (default: asks, Enter = today)")
    ap.add_argument("--skip-exams", action="store_true", help="leave out step 1 (Red Hat + Adv Pro)")
    ap.add_argument("--skip-feedback", action="store_true", help="leave out step 2 (feedback sheet)")
    ap.add_argument("--skip-linkedin", action="store_true", help="leave out step 3 (LinkedIn)")
    ap.add_argument("--skip-advpro", action="store_true", help="step 1 without the Adv Pro check")
    ap.add_argument("--ask", action="store_true", help="LinkedIn: wait for you when a profile isn't found")
    ap.add_argument("--dry-run", action="store_true", help="do everything except saving the Excel file")
    args = ap.parse_args()

    cfg = load_config()
    print("=" * 70)
    print("  Red Hat exams -> redhat exam 2026.xlsx")
    print("  The date you give is used for every step:")
    print("    1. Red Hat Exam Scheduler (Advantage pro Chrome) + Adv Pro check")
    print("    2. ADV PRO FEEDBACK Google Sheet (Ranjit Chrome)")
    print("    3. LinkedIn from Outlook (Arish Chrome)")
    print("=" * 70)
    date = dt.date.fromisoformat(args.date) if args.date else ask_date()
    log(f"Date: {date:%d-%b-%Y} for all steps   Excel: {cfg.get('excel', 'path')}")

    steps = [
        ("Exams (Red Hat + Adv Pro)", args.skip_exams,
         lambda: redhat_automation.run(cfg, date, skip_advpro=args.skip_advpro, dry_run=args.dry_run)),
        ("Feedback (Google Sheet)", args.skip_feedback,
         lambda: feedback_from_sheet.run(cfg, date, dry_run=args.dry_run)),
        ("LinkedIn (Outlook)", args.skip_linkedin,
         lambda: linkedin_from_outlook.run(cfg, date, date, ask=args.ask, dry_run=args.dry_run)),
    ]
    results = []
    for i, (name, skip, step) in enumerate(steps, 1):
        if skip:
            results.append((name, "skipped"))
            continue
        log("")
        log(f"========== Step {i}/{len(steps)}: {name} ==========")
        try:
            step()
            results.append((name, "OK"))
        except KeyboardInterrupt:
            results.append((name, "stopped"))
            log("Stopped by you - going to the next step (Ctrl+C again to stop everything)")
        except Exception as e:
            results.append((name, f"FAILED: {type(e).__name__}: {str(e).splitlines()[0][:150] if str(e) else ''}"))
            logs = HERE / "logs"
            logs.mkdir(exist_ok=True)
            err = logs / f"error_{dt.datetime.now():%Y%m%d_%H%M%S}.txt"
            err.write_text(traceback.format_exc(), encoding="utf-8")
            log(f"!! Step {i} failed: {type(e).__name__}: {e}")
            log(f"   details saved in logs\\{err.name} - going to the next step")

    log("")
    log("========== Summary ==========")
    for name, result in results:
        log(f"  {name}: {result}")
    log("Done")


if __name__ == "__main__":
    main()
