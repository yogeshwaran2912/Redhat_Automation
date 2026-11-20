@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist linkedin_from_outlook.py (python linkedin_from_outlook.py %*) else (python linkedin_from_outlook.pyc %*)
pause
