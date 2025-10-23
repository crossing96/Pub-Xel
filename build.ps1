# build.ps1
$ErrorActionPreference = "Stop"

# Clean old outputs
# basically redundant but harmless. On a self-hosted runner, 
# it’s essential to avoid stale junk contaminating builds. Keep it.
Remove-Item -Recurse -Force build, dist, Output -ErrorAction Ignore

# Python deps
pip install --upgrade pip
if (Test-Path requirements.txt) { pip install -r requirements.txt }
pip install pyinstaller

# Build exe
pyinstaller --clean --onefile --noconsole --icon=assets/logo128.ico --version-file version_info.txt --specpath ./ Pub-Xel.py `
  --add-data "data;data" `
  --add-data "ui;ui" `
  --add-data "assets;assets" `
  --add-data "mainfunctions.py;." `
  --add-data "welcome.py;."

# Sanity check
if (-not (Test-Path "dist\Pub-Xel.exe")) {
  Write-Error "PyInstaller output missing"
}