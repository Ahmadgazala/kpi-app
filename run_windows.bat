@echo off
REM Run without building - portable
echo Starting KPI-2026...
pip install -q matplotlib openpyxl
python local_app.py
pause
