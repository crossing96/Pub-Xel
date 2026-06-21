# Sync wiki/ from this repo to the GitHub Wiki git repository.
#
# Usage:
#   .\scripts\sync_wiki.ps1
#   .\scripts\sync_wiki.ps1 -WikiDir C:\path\to\Pub-Xel.wiki
#   .\scripts\sync_wiki.ps1 -Push:$false          # copy only, no git push
#
# First time: create any page on GitHub (Wiki tab) so the .wiki repo exists,
# or let this script clone it after that one-time setup.

param(
    [string]$WikiDir = "",
    [switch]$Push = $true
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
$SourceDir = Join-Path $ProjectRoot "wiki"

if (-not (Test-Path $SourceDir)) {
    throw "Wiki source folder not found: $SourceDir"
}

function Get-WikiRemoteUrl {
    $origin = git -C $ProjectRoot remote get-url origin 2>$null
    if (-not $origin) {
        throw "No git origin found. Set -WikiDir to your Pub-Xel.wiki clone path."
    }
    if ($origin -match "github\.com[:/](.+?)/(.+?)(?:\.git)?$") {
        $owner = $Matches[1]
        $repo = $Matches[2] -replace '\.git$', ''
        return "https://github.com/$owner/$repo.wiki.git"
    }
    throw "Could not derive wiki remote from origin: $origin"
}

if (-not $WikiDir) {
    $WikiDir = Join-Path (Split-Path -Parent $ProjectRoot) "Pub-Xel.wiki"
}

if (-not (Test-Path $WikiDir)) {
    $wikiRemote = Get-WikiRemoteUrl
    Write-Host "Cloning $wikiRemote -> $WikiDir"
    git clone $wikiRemote $WikiDir
} else {
    Write-Host "Updating wiki clone at $WikiDir"
    git -C $WikiDir pull --ff-only
}

python (Join-Path $PSScriptRoot "sync_wiki.py") $SourceDir $WikiDir
if ($LASTEXITCODE -ne 0) {
    throw "sync_wiki.py failed with exit code $LASTEXITCODE"
}

if (-not $Push) {
    Write-Host "Copy complete (-Push:`$false). Review $WikiDir and commit manually."
    exit 0
}

Push-Location $WikiDir
try {
    git add -A
    $status = git status --porcelain
    if (-not $status) {
        Write-Host "Wiki already up to date."
        exit 0
    }
    git commit -m "Sync wiki from main repo"
    git push
    Write-Host "Wiki published."
} finally {
    Pop-Location
}
