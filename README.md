# KPI 2026 - Local App (Mac + Windows)

Works 100% offline. No browser needed.

## 📦 Portable (Works on BOTH Mac & Windows without building)

1. Copy folder `kpi` or zip `KPI-2026-portable.zip` to any computer
2. Install Python 3.10+ from https://www.python.org
3. Run:

**Mac:** Double-click `run_mac.command`  (or Terminal: `python3 local_app.py`)

**Windows:** Double-click `run_windows.bat`  (or CMD: `python local_app.py`)

Auto-installs `matplotlib openpyxl` on first run.

---

## 🖥️ Standalone Executable (No Python needed)

### Mac - Build .app
```bash
./build_mac.sh
# or: python3 build.py
# Output: dist/KPI-2026.app (77MB) -> double-click
```
Already built: `KPI-2026.app` in this folder.

### Windows - Build .exe (must be done ON Windows)
Copy entire folder to Windows PC, then:
```cmd
build_windows.bat
:: or: python build.py
:: Output: dist\KPI-2026.exe -> double-click
```
> Note: .exe must be built on Windows (PyInstaller cannot cross-compile). Use the portable method if you don't want to build.

---

## How to Use
1. Top bar: enter `اسم المؤسسة`
2. **📝 التقييم**: left = 9 domains (58 indicators). For each: choose 1 ناشئ ... 5 ديناميكي. Description updates live.
3. **📈 النتائج**: click `🔄 تحديث الرسوم` -> Bar + Donut charts, gap analysis, table.
4. **💾 تصدير**: Excel (updates original `KPI-Final 2026.xlsx` with الدرجة + النتائج), CSV, JSON save/load.

## Files
- `local_app.py` - Main local app (Tkinter) - **cross-platform**
- `KPI-Final 2026.xlsx` - Template
- `kpi_domains.json` - 58 indicators
- `KPI-2026.app` - Mac executable (built)
- `build_mac.sh` / `build_windows.bat` / `build.py` - Builders
- `run_mac.command` / `run_windows.bat` - Portable launchers
- `KPI-2026-portable.zip` - Shareable zip for both OS

## Requirements
```
matplotlib
openpyxl
```
Included in `requirements.txt`

## Troubleshooting
- **Mac "untrusted"**: Right-click KPI-2026.app -> Open -> Open
- **Windows SmartScreen**: Click "More info" -> Run anyway
- **Arabic fonts**: Auto-detects Segoe UI (Windows) / Helvetica (Mac)
- **Excel not found**: Keep `KPI-Final 2026.xlsx` next to `local_app.py` or next to exe
