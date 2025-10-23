# build.ps1
$ErrorActionPreference = "Stop"

# Clean old outputs
# basically redundant but harmless. On a self-hosted runner, 
# it’s essential to avoid stale junk contaminating builds. Keep it.
Remove-Item -Recurse -Force build, dist, Output -ErrorAction Ignore

# Python deps
pip install --upgrade pip
pip install PyQt6 xlwings pynput pyperclip requests beautifulsoup4 pyinstaller

# Verify imports (fail fast)
python - << 'PY'
import importlib, sys
mods = [
  "PyQt6", "PyQt6.QtCore", "PyQt6.QtGui", "PyQt6.QtWidgets",
  "xlwings", "pynput", "pyperclip", "requests", "bs4"
]
failed = []
for m in mods:
    try:
        importlib.import_module(m)
        print(f"OK: {m}")
    except Exception as e:
        failed.append((m, repr(e)))
if failed:
    print("Missing modules:")
    for m,e in failed: print(" -", m, e)
    sys.exit(1)
PY

# Build exe
$opts = @(
  "--clean",
  "--onefile",
  "--noconsole",
  "--icon=assets/logo128.ico",
  "--version-file", "version_info.txt",
  "--specpath", ".",
  "Pub-Xel.py",
  "--add-data", "data;data",
  "--add-data", "ui;ui",
  "--add-data", "assets;assets",
  "--add-data", "mainfunctions.py;.",
  "--add-data", "welcome.py;.",
  "--collect-all", "PyQt6",
  "--collect-all", "xlwings",
  "--hidden-import", "PyQt6.QtCore",
  "--hidden-import", "PyQt6.QtGui",
  "--hidden-import", "PyQt6.QtWidgets",
  "--hidden-import", "PyQt6.QtSvg",
  "--hidden-import", "PyQt6.QtNetwork",
  "--hidden-import", "xlwings",
  "--hidden-import", "pyperclip",
  "--hidden-import", "pynput",
  "--hidden-import", "pynput.keyboard",
  "--hidden-import", "pynput.keyboard._win32"
)

pyinstaller @opts --log-level=DEBUG


# Sanity check
if (-not (Test-Path "dist\Pub-Xel.exe")) {
  Write-Error "PyInstaller output missing"
}