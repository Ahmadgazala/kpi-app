# KPI 2026 - Local Desktop App

## Quick Start (macOS)

### Option 1: Double-click App (Easiest) ✅
Just double-click:
```
KPI-2026.app
```
No terminal, no internet. Works offline.

Location: `/Users/ahmad/Desktop/kpi/KPI-2026.app`

If macOS says "untrusted developer":
- Right-click → Open → Open

### Option 2: Run script
```bash
python3 local_app.py
```
or double-click:
```
run_local.command
```

Requirements for script mode:
```bash
pip3 install matplotlib openpyxl
```

### Option 3: Web version (browser)
```bash
streamlit run app.py
```

## How to Use (3 steps)
1. Enter **اسم المؤسسة** at top
2. Go to **📝 التقييم** → select domain from left list → choose 1-5 for each indicator
   - 1 ناشئ (red) → 5 ديناميكي (teal)
   - Description updates live
3. Go to **📈 النتائج** → Click **تحديث الرسوم** → **تصدير Excel**

## Features
- **58 indicators, 9 domains** - exactly as in KPI-Final 2026.xlsx
- Live calculation: المتوسط, المستوى, العلامة/100
- Charts: Bar + Donut gauge + table with gap analysis
- Exports:
  - **Excel** - updates original file (fills الدرجة + العلامة + النتائج sheets)
  - **CSV** - summary table
  - **JSON** - save/load to continue later
- Presets: تصفير / الكل 5 / سيناريو سريع
- Random demo button for testing

## Files
- `local_app.py` - main local app (Tkinter, offline)
- `KPI-2026.app` - macOS bundle (double-click)
- `KPI-2026-local` - single binary
- `KPI-Final 2026.xlsx` - original data
- `kpi_domains.json` - indicators

## Troubleshooting
- If app doesn't open: `chmod +x KPI-2026.app/Contents/MacOS/KPI-2026 && ./KPI-2026.app/Contents/MacOS/KPI-2026`
- For other OS (Windows): `pyinstaller --onefile --windowed local_app.py` on that OS
