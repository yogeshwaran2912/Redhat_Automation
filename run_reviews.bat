@echo off
@rem Created by YOGESHWARAN TL
cd /d "%~dp0"
if exist google_reviews.py (python google_reviews.py %*) else (python google_reviews.pyc %*)
pause
