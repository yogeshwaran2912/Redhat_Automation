@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if not exist logs mkdir logs
echo stop> logs\stop_requested
echo The automation will stop within a few seconds.
timeout /t 5 >nul
