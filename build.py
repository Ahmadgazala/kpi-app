#!/usr/bin/env python3
"""
Cross-platform builder for KPI-2026 local app
Usage:
  python3 build.py        -> builds for current OS (Mac or Windows)
  python3 build.py --all  -> shows commands for both
"""
import platform, subprocess, pathlib, sys, os

BASE = pathlib.Path(__file__).parent
SEPARATOR = ";" if platform.system() == "Windows" else ":"

def build():
    # Ensure deps
    print(f"Building on {platform.system()} {platform.machine()}...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile" if platform.system() == "Windows" else "--onefile",
        "--windowed",
        "--name", "KPI-2026",
        "--add-data", f"{BASE/'KPI-Final 2026.xlsx'}{SEPARATOR}.",
        "--add-data", f"{BASE/'kpi_domains.json'}{SEPARATOR}.",
        "--add-data", f"{BASE/'NotoNaskhArabic-Regular.ttf'}{SEPARATOR}.",
        "--add-data", f"{BASE/'NotoNaskhArabic-Bold.ttf'}{SEPARATOR}.",
        "--collect-all", "customtkinter",
        "--collect-all", "reportlab",
        "--collect-all", "fpdf",
        "--collect-all", "uharfbuzz",
        "--collect-all", "matplotlib",
        "--collect-all", "mpl_toolkits",
        str(BASE/"local_app.py"),
        "--distpath", str(BASE/"dist"),
        "--workpath", str(BASE/"build"),
        "--clean"
    ]
    # Add platform-specific icon
    if platform.system() == "Darwin":
        ico = BASE / "KPI.icns"
        if ico.exists():
            cmd.insert(4, "--icon")
            cmd.insert(5, str(ico))
    elif platform.system() == "Windows":
        ico = BASE / "KPI.ico"
        if ico.exists():
            cmd.insert(4, "--icon")
            cmd.insert(5, str(ico))
    # On Mac, also build .app bundle via --windowed already does, but provide option
    print("Running:", " ".join(cmd))
    ret = subprocess.call(cmd)
    if ret==0:
        print("\n✓ Build success!")
        print(f"  Output in: {BASE/'dist'}")
        for p in (BASE/"dist").glob("*KPI*"):
            print(f"   - {p} ({p.stat().st_size//1024//1024} MB)")
        if platform.system() == "Darwin":
            print("\nFor Windows: copy this whole folder to a Windows PC and run:")
            print("  build_windows.bat")
        else:
            print("\nFor Mac: copy folder to Mac and run:")
            print("  ./build_mac.sh")
    else:
        print("Build failed")

if __name__ == "__main__":
    if "--help" in sys.argv:
        print(__doc__)
    elif "--all" in sys.argv:
        print("Mac command:\n  pyinstaller --onefile --windowed --name KPI-2026 --add-data 'KPI-Final 2026.xlsx:.' --add-data 'kpi_domains.json:.' local_app.py\n")
        print("Windows command:\n  pyinstaller --onefile --windowed --name KPI-2026 --add-data \"KPI-Final 2026.xlsx;.\" --add-data \"kpi_domains.json;.\" local_app.py")
    else:
        build()
