# Pub-Xel Windows build (PyInstaller onedir).
# Output: dist\Pub-Xel\Pub-Xel.exe
# Installer: compile setup.iss with Inno Setup (see .github/workflows/build-release.yml).

$ErrorActionPreference = "Stop"

Set-Location $PSScriptRoot

# Avoid pulling unrelated packages from a polluted PYTHONPATH (editable installs, etc.).
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue

function Test-PythonExe {
    param([string]$Exe)
    if (-not (Test-Path -LiteralPath $Exe)) { return $false }
    & $Exe -c "import sys" 2>$null | Out-Null
    return $LASTEXITCODE -eq 0
}

function Get-BuildPython {
    # Prefer 3.11 to match CI (.github/workflows/build-release.yml).
    if (Get-Command py -ErrorAction SilentlyContinue) {
        foreach ($ver in @("3.11", "3.12", "3.13")) {
            $prevEap = $ErrorActionPreference
            $ErrorActionPreference = "Continue"
            $exe = & py "-$ver" -c "import sys; print(sys.executable)" 2>$null
            $pyExit = $LASTEXITCODE
            $ErrorActionPreference = $prevEap
            if ($pyExit -eq 0 -and $exe) {
                $resolved = $exe.Trim()
                if (Test-PythonExe $resolved) {
                    Write-Host "Using Python ${ver}: $resolved"
                    return $resolved
                }
            }
        }
    }

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd -and (Test-PythonExe $pythonCmd.Source)) {
        Write-Host "Using Python: $($pythonCmd.Source)"
        return $pythonCmd.Source
    }

    throw "No working Python found. Install Python 3.11+ (recommended: 3.11 for CI parity)."
}

$Python = Get-BuildPython

function Invoke-Python {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)
    & $Python @Args
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed ($LASTEXITCODE): $Python $($Args -join ' ')"
    }
}

# Clean old outputs (essential on CI/self-hosted runners; harmless locally).
Remove-Item -Recurse -Force build, dist, Output -ErrorAction Ignore

Write-Host "Installing dependencies from requirements.txt..."
Invoke-Python -m pip install --upgrade pip
Invoke-Python -m pip install -r requirements.txt

# Verify runtime imports (fail fast before PyInstaller).
$verifyScript = @"
import importlib, sys
mods = [
  "PyQt6", "PyQt6.QtCore", "PyQt6.QtGui", "PyQt6.QtWidgets",
  "xlwings", "pynput", "requests", "bs4"
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
    for m, e in failed:
        print(" -", m, e)
    sys.exit(1)
"@

Set-Content -Path verify_imports.py -Value $verifyScript -Encoding UTF8
try {
    Invoke-Python verify_imports.py
} finally {
    Remove-Item verify_imports.py -Force -ErrorAction Ignore
}

Invoke-Python scripts/smoke_imports.py

Write-Host "Running PyInstaller..."
$pyinstallerOpts = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onedir",
    "--noconsole",
    "--icon=assets/logo128.ico",
    "--version-file", "version_info.txt",
    "--specpath", ".",
    "Pub-Xel.py",
    "--add-data", "data;data",
    "--add-data", "ui;ui",
    "--add-data", "assets;assets",
    "--add-data", "pubxel_core;pubxel_core",
    "--exclude-module", "PyQt6.QtWebEngineCore",
    "--exclude-module", "PyQt6.QtWebEngineWidgets",
    "--exclude-module", "PyQt6.QtWebChannel",
    "--exclude-module", "PyQt6.QtQuick",
    "--exclude-module", "PyQt6.QtQml",
    "--exclude-module", "PyQt6.QtMultimedia",
    "--exclude-module", "PyQt6.QtNetwork",
    "--exclude-module", "PyQt6.QtPrintSupport",
    "--exclude-module", "PyQt6.QtSql",
    "--exclude-module", "PyQt6.QtBluetooth",
    "--exclude-module", "PyQt6.QtNfc",
    "--exclude-module", "PyQt6.QtSensors",
    "--exclude-module", "PyQt6.QtPositioning",
    "--exclude-module", "PyQt6.QtOpenGL",
    "--exclude-module", "PyQt6.QtSvg",
    "--log-level=INFO"
)

Invoke-Python @pyinstallerOpts

$targetDir = Join-Path $PSScriptRoot "dist\Pub-Xel"
$targetExe = Join-Path $targetDir "Pub-Xel.exe"
if (-not (Test-Path -LiteralPath $targetExe)) {
    throw "PyInstaller output missing: $targetExe"
}

Write-Host ""
Write-Host "Build OK."
Write-Host "  Run:       $targetExe"
Write-Host "  Folder:    $targetDir"
Write-Host "  Installer: compile setup.iss with Inno Setup (see CI workflow)."
