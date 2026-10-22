<#
One-time setup for a teammate's machine. Run this once, then the scraper runs on its own every
day at the time you choose, writing into the SAME shared workbook everyone else uses.

Usage (from PowerShell, in this "team-install" folder):
    .\Install-Teammate.ps1 -RunAtTime "08:30"

If your local OneDrive names the shared library folder differently, pass -SharedLibraryRoot
pointing at wherever it resolves on your machine (see Step 1's instructions if it's not found).
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^([01]\d|2[0-3]):[0-5]\d$')]
    [string]$RunAtTime,                 # e.g. "08:30" — pick your own normal morning time, your call

    # IMPORTANT: this must be the folder that results from clicking "Add shortcut to OneDrive"
    # while INSIDE the "06 Sensing (Inputs)" folder itself — not its parent "Global PP+G -
    # Documents" library. Doing it at the wrong level (or replacing a broken/stale shortcut
    # incorrectly) can silently create a local folder that looks right but has no real OneDrive
    # sync registration — writes just pile up locally and never reach SharePoint or your
    # teammates, with no error shown anywhere. Found and fixed 2026-07-28.
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - 06 Sensing (Inputs)",
    [string]$Python = ""                # empty => auto-discover / auto-install
)

$ErrorActionPreference = "Stop"
$teamInstallRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $teamInstallRoot
# Runs BOTH the web scraper and the Signum Iran email pipeline (see run_daily_full_update.ps1).
# The Signum half only does anything if YOU personally receive Andrew Bishop's Signum Iran
# emails in Outlook — if you don't, it simply finds nothing and adds nothing; it won't fail
# your install or affect the web-scraper half.
$orchestrator = Join-Path $projectRoot "run_daily_full_update.ps1"

Write-Host "==== Step 1: checking the shared library is synced ====" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $SharedLibraryRoot)) {
    Write-Host @"
Could not find the shared library at:
  $SharedLibraryRoot

Before continuing:
  1. Open Teams or SharePoint and find 'Global PP+G - Documents'.
  2. Navigate INTO the '06 Sensing (Inputs)' folder itself (not just the top-level library —
     adding the shortcut at the wrong level can silently create a folder that never actually
     syncs, with no error shown).
  3. Click 'Add shortcut to OneDrive' (top toolbar) while you're inside that folder.
  4. Wait a minute, then check under your own OneDrive folder in File Explorer for a new
     folder (something like 'Global PP+G - 06 Sensing (Inputs)'), then re-run this installer.

If it syncs under a different name on your machine, re-run with:
  .\Install-Teammate.ps1 -RunAtTime "$RunAtTime" -SharedLibraryRoot "<the exact path that appeared>"
"@ -ForegroundColor Yellow
    exit 1
}
Write-Host "Found: $SharedLibraryRoot" -ForegroundColor Green

Write-Host "`n==== Step 2: checking Python ====" -ForegroundColor Cyan
function Resolve-Python {
    foreach ($dir in Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending) {
        $candidate = Join-Path $dir.FullName "python.exe"
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    $onPath = (Get-Command python -ErrorAction SilentlyContinue)
    if ($onPath -and $onPath.Source -notlike "*WindowsApps*") { return $onPath.Source }
    return $null
}

$py = if ($Python) { $Python } else { Resolve-Python }
if (-not $py) {
    $winget = (Get-Command winget -ErrorAction SilentlyContinue)
    if (-not $winget) {
        Write-Host "Python was not found and 'winget' isn't available to install it automatically." -ForegroundColor Yellow
        Write-Host "Please install Python 3 yourself from https://python.org/downloads (check 'Add to PATH'), then re-run this installer." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "Python not found — installing via winget (user scope, no admin needed)..." -ForegroundColor Yellow
    winget install --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
    $py = Resolve-Python
    if (-not $py) {
        Write-Host "Install finished but Python still wasn't found. Close and reopen this PowerShell window, then re-run this installer." -ForegroundColor Yellow
        exit 1
    }
}
Write-Host "Using Python: $py" -ForegroundColor Green
& $py --version

Write-Host "`n==== Step 3: installing required Python packages ====" -ForegroundColor Cyan
& $py -m ensurepip --upgrade | Out-Null
& $py -m pip install --disable-pip-version-check -q -r (Join-Path $projectRoot "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "pip install failed." }
Write-Host "Packages installed." -ForegroundColor Green

Write-Host "`n==== Step 4: registering the daily scheduled task ====" -ForegroundColor Cyan
# Task Scheduler rejects a literal colon in -TaskName (Register-ScheduledTask fails with "The
# parameter is incorrect", HRESULT 0x80070057) - confirmed directly while building the sibling
# expert-commentary-scraper tool's installer. $RunAtTime ("HH:mm") always has one, so it can't be
# used as-is in the name, only in the actual trigger time below.
$taskName = "Geopolitics Article Scraper - Daily $($RunAtTime.Replace(':', ''))"
$time = [datetime]::ParseExact($RunAtTime, "HH:mm", $null)
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument ('-NonInteractive -ExecutionPolicy Bypass -File "{0}" -SharedLibraryRoot "{1}" -Python "{2}"' -f $orchestrator, $SharedLibraryRoot, $py)
$trigger = New-ScheduledTaskTrigger -Daily -At $time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description "Daily: scrape think-tank Iran articles and (if you have access) Signum Iran emails into the team's shared geopolitics repository workbook." -Force | Out-Null
Write-Host "Registered task '$taskName'." -ForegroundColor Green

Write-Host "`n==== Step 5: running it once now to confirm everything works ====" -ForegroundColor Cyan
& powershell -NoProfile -ExecutionPolicy Bypass -File $orchestrator -SharedLibraryRoot $SharedLibraryRoot -Python $py
if ($LASTEXITCODE -ne 0) {
    Write-Host "The test run failed — see the messages above. The scheduled task is still registered; fix the issue, then it'll pick up correctly tomorrow (or run the .cmd launcher again to retest)." -ForegroundColor Yellow
    exit 1
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " All set. This machine will now run daily at $RunAtTime." -ForegroundColor Green
Write-Host " That covers the think-tank web scraper, plus your Signum Iran emails too if you" -ForegroundColor Green
Write-Host " receive them (nothing extra to configure — it just finds nothing if you don't)." -ForegroundColor Green
Write-Host " Shared workbook everyone reads:" -ForegroundColor Green
Write-Host "   $SharedLibraryRoot\10 - Geopolitics Sensing & Scraping Tools\Iran Scraping example\Geopolitics Article Repository (review).xlsx" -ForegroundColor Green
Write-Host " To run it manually anytime, double-click 'Run Geopolitics Scraper.cmd' in the project folder." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
