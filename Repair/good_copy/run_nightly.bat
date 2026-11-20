@echo off
@rem Created by YOGESHWARAN TL
rem Started (hidden, by run_hidden.vbs) by Windows Task Scheduler at 11:45 PM and then every
rem 10 minutes until the morning. keep_running.py starts the work for tonight in the background,
rem or starts it again if it was killed; once the Excel file is complete it does nothing.
rem Results: logs\keep_running.txt   Program output: logs\nightly_YYYYMMDD.txt
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if exist keep_running.py (python keep_running.py --nightly) else (python keep_running.pyc --nightly)
