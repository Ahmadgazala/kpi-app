# KPI 2026 - Distribution Package

## Contents

### macOS (Ready to run)
- `macOS/KPI-2026.app` - Ready-to-run macOS application

### Windows (Build required)
- Source files and build scripts for Windows

### Source Files (for rebuilding)
- `local_app.py` - Main application
- `kpi_domains.json` - KPI indicators data
- `KPI-Final 2026.xlsx` - Excel template
- `NotoNaskhArabic-Regular.ttf` / `NotoNaskhArabic-Bold.ttf` - Arabic fonts
- `KPI.icns` / `KPI.ico` / `icon.png` - Application icons
- `requirements.txt` - Python dependencies
- `build_mac.sh` / `build_windows.bat` / `build.py` - Build scripts

## Quick Start

### macOS (Ready to use)
1. Copy `macOS/KPI-2026.app` to Applications folder
2. Double-click to run
3. First launch will show onboarding tutorial

### Windows (Build required)
1. Install Python 3.10+ from python.org
2. Run `build_windows.bat` (installs dependencies and builds .exe)
3. Find `dist/KPI-2026.exe` and copy to desired location

## Features
- 58 KPI indicators across 9 domains
- Real-time scoring (1-5 scale)
- Live progress tracking
- Excel export (with all formulas preserved)
- PDF export with correct Arabic text rendering
- Autosave & session restore
- Modern UI with progress tracking

## Requirements
- macOS 10.15+ or Windows 10+
- No Python installation needed for macOS .app
- Windows: Python 3.10+ required for build only

## Data Storage
- Autosave: `~/.kpi_autosave.json` (auto-created)
- Onboarding flag: `~/.kpi_onboarded` (auto-created)

## Support
For issues, check the console output or contact support.
