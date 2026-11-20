@echo off
@rem Created by YOGESHWARAN TL
rem Opens the live status window of the nightly run any time (closing it does not stop anything)
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if exist live_view.py (python live_view.py %*) else (python live_view.pyc %*)
