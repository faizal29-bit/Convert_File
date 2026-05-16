@echo off
title LocalConvert
echo ============================================
echo   LocalConvert - http://localhost:5000
echo ============================================
call venv\Scripts\activate
start "" http://localhost:5000
python app.py
pause
