# build.ps1
$ErrorActionPreference = "Stop"

# Clean old outputs
# basically redundant but harmless. On a self-hosted runner, 
# it’s essential to avoid stale junk contaminating builds. Keep it.
Remove-Item -Recurse -Force build, dist, Output -ErrorAction Ignore

# Python deps
pip install --upgrade pip
pip install PyQt6 xlwings pyperclip pynput pyinstaller

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
  "--hidden-import", "pyperclip"
)

pyinstaller @opts --log-level=DEBUG


# Sanity check
if (-not (Test-Path "dist\Pub-Xel.exe")) {
  Write-Error "PyInstaller output missing"
}