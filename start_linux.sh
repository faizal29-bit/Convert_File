#!/bin/bash
source venv/bin/activate
echo "============================================"
echo "  LocalConvert - http://localhost:5000"
echo "============================================"
# Buka browser otomatis
if command -v xdg-open &> /dev/null; then
    xdg-open http://localhost:5000 &
elif command -v open &> /dev/null; then
    open http://localhost:5000 &
fi
python app.py
