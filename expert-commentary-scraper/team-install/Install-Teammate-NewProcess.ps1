<#
One-time setup for a RESEARCH TEAM teammate's machine, for the NEW Process & Scope (the source
list rebuilt ahead of this handover - see run_new_process_update.ps1's own docstring). This is a
SEPARATE installer from Install-Teammate.ps1, which still exists and is untouched - that one sets
up the OLD Process (run_daily_full_update.ps1), which stays running exactly as it always has for
the user handing this tool off. Running the wrong installer on a teammate's machine would set up
the wrong source scope and workbook column layout entirely, so these are kept as two distinct
scripts rather than one script with a mode flag - less room for a silent wrong choice.

Usage (from PowerShell, in this "team-install" folder):
    .\Install-Teammate-NewProcess.ps1 -RunAtTime "08:30"

If your local OneDrive names the shared library folder differently, pass -SharedLibraryRoot
pointing at wherever it resolves on your machine (see Step 1's instructions if it's not found).
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^([01]\d|2[0-3]):[0-5]\d$')]
    [string]$RunAtTime,                 # e.g. "08:30" - pick your own normal morning time, your call

    # IMPORTANT: this must be the folder that results from clicking "Add shortcut to OneDrive"
    # while INSIDE the "Webscraping & Email Scraping tool" folder itself - not a parent folder.
    # Doing it at the wrong level can silently create a local folder that looks right but has no
    # real OneDrive sync registration - writes just pile up locally and never reach SharePoint or
    # your teammates, with no error shown anywhere.
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1",
    [string]$Python = ""                # empty => auto-discover / auto-install
)

$ErrorActionPreference = "Stop"
$teamInstallRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $teamInstallRoot
# Runs the NEW Process's orchestrator specifically - web scraper + vendor email pipeline, pointed
# at sources_config.new.json / vendors.new.json / the new workbook column layout, writing into the
# "New Process & Scope" subfolder of the shared library. See run_new_process_update.ps1's own
# docstring for the full picture.
$orchestrator = Join-Path $projectRoot "run_new_process_update.ps1"

Write-Host "==== Step 1: checking the shared library is synced ====" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $SharedLibraryRoot)) {
    Write-Host @"
Could not find the shared library at:
  $SharedLibraryRoot

Before continuing:
  1. Open Teams or SharePoint and find the 'Webscraping & Email Scraping tool' folder, under:
     Global Public Policy & Geopolitics site > Shared Documents > 03 Cross-Team Projects >
     08 - Geopolitics Sensing & Scraping Tools (ask your contact for the current link if unsure).
  2. Navigate INTO that folder itself (not a parent folder - adding the shortcut at the wrong
     level can silently create a folder that never actually syncs, with no error shown).
  3. Click 'Add shortcut to OneDrive' (top toolbar) while you're inside that folder.
  4. Wait a minute, then check under your own OneDrive folder in File Explorer for a new
     folder, then re-run this installer.

If it syncs under a different name on your machine, re-run with:
  .\Install-Teammate-NewProcess.ps1 -RunAtTime "$RunAtTime" -SharedLibraryRoot "<the exact path that appeared>"
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
    Write-Host "Python not found - installing via winget (user scope, no admin needed)..." -ForegroundColor Yellow
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

Write-Host "`n==== Step 4: downloading the browser used for a few sources ====" -ForegroundColor Cyan
Write-Host "A handful of sources (WEF, some linked articles) actively block plain web requests, so" -ForegroundColor DarkGray
Write-Host "those use a real (headless) browser instead - this is a one-time ~300MB download." -ForegroundColor DarkGray
& $py -m playwright install chromium
if ($LASTEXITCODE -ne 0) { throw "playwright install chromium failed." }
Write-Host "Browser installed." -ForegroundColor Green

Write-Host "`n==== Step 5: registering the daily scheduled task ====" -ForegroundColor Cyan
# Task Scheduler rejects a literal colon in -TaskName (confirmed directly) - $RunAtTime ("HH:mm")
# always has one, so it can't be used as-is in the name, only in the actual trigger time below.
# "NEW PROCESS" is in the task name on purpose - this machine may eventually also see the Old
# Process's own task registered by someone else; the names must never be confused with each
# other (a real incident during this tool's own build: a stale "0600" task name that actually
# triggered at 9:00 AM caused two unrelated runs to collide - always name a task for what it
# ACTUALLY does, not just what it was first called).
$taskName = "Expert Commentary Scraper - NEW PROCESS - Daily $($RunAtTime.Replace(':', ''))"
$time = [datetime]::ParseExact($RunAtTime, "HH:mm", $null)
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument ('-NonInteractive -ExecutionPolicy Bypass -File "{0}" -SharedLibraryRoot "{1}" -Python "{2}"' -f $orchestrator, $SharedLibraryRoot, $py)
$trigger = New-ScheduledTaskTrigger -Daily -At $time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description "Daily: scrape the New Process & Scope's web sources and (if you have access) vendor emails into the team's shared Expert Commentary repository workbook." -Force | Out-Null
Write-Host "Registered task '$taskName'." -ForegroundColor Green

Write-Host "`n==== Step 6: running it once now to confirm everything works ====" -ForegroundColor Cyan
& powershell -NoProfile -ExecutionPolicy Bypass -File $orchestrator -SharedLibraryRoot $SharedLibraryRoot -Python $py
if ($LASTEXITCODE -ne 0) {
    Write-Host "The test run failed - see the messages above. The scheduled task is still registered; fix the issue, then it'll pick up correctly tomorrow (or run this installer again to retest)." -ForegroundColor Yellow
    exit 1
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " All set. This machine will now run daily at $RunAtTime." -ForegroundColor Green
Write-Host " That covers the New Process & Scope's web sources, plus any vendor emails you receive" -ForegroundColor Green
Write-Host " (Signum, Eurasia Group, Middle East Council on Global Affairs, CSIS, CFR, Atlantic" -ForegroundColor Green
Write-Host " Council) - nothing extra to configure; it just finds nothing for whichever ones you" -ForegroundColor Green
Write-Host " aren't yet subscribed to (see the subscriptions list from the handover deck)." -ForegroundColor Green
Write-Host " Shared workbook everyone reads:" -ForegroundColor Green
Write-Host "   $SharedLibraryRoot\New Process & Scope\Expert Commentary Repository (review).xlsx" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
