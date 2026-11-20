Created by YOGESHWARAN TL

REPAIR TOOL
===========

When the automation shows errors, gets stuck, or does not run at night:
double-click  Repair.bat  and choose:

  1  Quick repair    - closes the stuck program, removes old locks, cleans temporary files,
                       checks Python + packages (installs missing ones), fixes .bat launchers,
                       checks the program files, config.ini, the Excel file on the network,
                       the automation Chrome, and the scheduled tasks (turns them on / makes
                       them again).
  2  Full reset      - Quick repair + puts back ALL program files from the "good copy",
                       like a fresh install. The files it replaces are kept in replaced_<date>.
  3  Save good copy  - do this when everything works well (after a good night run).
                       Full reset brings the programs back to this copy.
  4  Repair + start  - Quick repair, then starts the automation (run_all.bat).

Never deleted: Excel files, logs, reports, backups. config.ini is only put back if it is broken.
Every run writes a log in  Repair\logs\repair_<date>_<time>.txt

At the end it shows  FIXED  and  PROBLEMS LEFT. A problem it cannot fix itself (Excel open on
another PC, network off, Python missing) is written in plain words - fix that, then run again.
