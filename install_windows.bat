@echo off
title LocalConvert - Install
echo ============================================
echo   LocalConvert - Install Dependencies
echo ============================================
echo.

:: Cek Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan!
    echo Download Python di: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Membuat virtual environment...
python -m venv venv
if errorlevel 1 (
    echo [ERROR] Gagal membuat venv
    pause
    exit /b 1
)

echo [2/3] Install dependensi Python...
call venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Gagal install dependensi
    pause
    exit /b 1
)

echo [3/3] Selesai!
echo.
echo ============================================
echo   CATATAN PENTING:
echo ============================================
echo.
echo   1. Install LibreOffice (untuk Word/PPT/Excel ^ PDF):
echo      https://www.libreoffice.org/download/libreoffice/
echo.
echo   2. Install Poppler (untuk PDF ^ PPT):
echo      https://github.com/oschwartz10612/poppler-windows/releases
echo      Ekstrak dan tambahkan folder "bin" ke PATH
echo      Atau letakkan di: C:\poppler\bin
echo.
echo   Jalankan: start_windows.bat
echo ============================================
echo.
pause
