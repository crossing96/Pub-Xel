# Run with: 
# pwsh ./build_local.ps1 -Version 1.2.3

# build_local.ps1
param([string]$Version = "0.0.0-local")

$ErrorActionPreference = "Stop"
./build.ps1

# Inno
$iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $iscc)) { throw "Install Inno Setup or add ISCC to PATH" }
& "$iscc" "/DVersion=$Version" "setup.iss"
Write-Host "Installer at: Output\Pub-Xel_Installer_v$Version.exe"

