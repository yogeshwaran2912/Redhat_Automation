@echo off
@rem Created by YOGESHWARAN TL
rem REPAIR TOOL - double-click when the automation gives errors or does not run.
title Repair tool
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
py -3.11 -c "pass" >nul 2>&1
if errorlevel 1 (python repair.py %*) else (py -3.11 repair.py %*)
pause
