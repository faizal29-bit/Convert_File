#!/bin/bash
set -e

echo "============================================"
echo "  LocalConvert - Install"
echo "============================================"
echo ""

# Cek Python
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 tidak ditemukan!"
    echo "Install: sudo apt install python3 python3-pip python3-venv"
    exit 1
fi

echo "[1/4] Membuat virtual environment..."
python3 -m venv venv

echo "[2/4] Aktifkan venv dan install dependensi..."
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo "[3/4] Cek dependensi sistem..."

# LibreOffice
if command -v libreoffice &> /dev/null; then
    echo "  ✅ LibreOffice: OK"
else
    echo "  ⚠️  LibreOffice tidak ditemukan."
    echo "  Install: sudo apt install libreoffice  (Ubuntu/Debian)"
    echo "           brew install --cask libreoffice  (macOS)"
fi

# Poppler
if command -v pdftoppm &> /dev/null; then
    echo "  ✅ Poppler: OK"
else
    echo "  ⚠️  Poppler tidak ditemukan."
    echo "  Install: sudo apt install poppler-utils  (Ubuntu/Debian)"
    echo "           brew install poppler  (macOS)"
fi

echo ""
echo "[4/4] Selesai!"
echo ""
echo "  Jalankan: bash start_linux.sh"
echo "============================================"
