#!/bin/bash
cd "$(dirname "$0")"
pip3 install customtkinter matplotlib openpyxl reportlab arabic-reshaper python-bidi pillow fpdf2 uharfbuzz pyinstaller -q
python3 -m PyInstaller --onefile --windowed --name "KPI-2026" \
  --add-data "KPI-Final 2026.xlsx:." \
  --add-data "kpi_domains.json:." \
  --add-data "NotoNaskhArabic-Regular.ttf:." \
  --add-data "NotoNaskhArabic-Bold.ttf:." \
  --icon "KPI.icns" \
  --collect-all customtkinter \
  --collect-all reportlab \
  --collect-all fpdf \
  --collect-all uharfbuzz \
  --collect-all matplotlib \
  local_app.py --distpath ./dist --workpath ./build --clean
echo "Done! dist/KPI-2026.app"
