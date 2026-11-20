@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist redhat_automation.py (python redhat_automation.py %*) else (python redhat_automation.pyc %*)
pause
