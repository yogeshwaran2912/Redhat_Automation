@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist keep_running.py (python keep_running.py %*) else (python keep_running.pyc %*)
timeout /t 5 >nul
