@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist feedback_from_sheet.py (python feedback_from_sheet.py %*) else (python feedback_from_sheet.pyc %*)
pause
