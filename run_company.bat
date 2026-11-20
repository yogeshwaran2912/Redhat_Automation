@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist company_from_linkedin.py (python company_from_linkedin.py %*) else (python company_from_linkedin.pyc %*)
pause
