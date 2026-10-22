<#
Single daily entry point: runs the web scraper (expert-commentary hub pages) and the vendor
email pipeline (Signum, Eurasia Group, and any future one - see email-scraper/vendors.json), all
writing into the SAME shared team workbook. This is what a scheduled task should point at, not
either sub-script alone.

Each pipeline runs independently - if one fails (e.g. a teammate doesn't have Outlook, or a hub
page is temporarily down), the other still runs and the shared workbook still gets whatever
updates succeeded. Only if BOTH fail does this script itself report failure.
#>
param(
    # 2026-08-18: see the matching comment in run_articles_update.ps1 - re-pointed after the
    # SharePoint site was restructured (a real move, not a rename). Shortcut is now added directly
    # at the "Webscraping & Email Scraping tool" folder level, so SharedSubfolder is empty.
    #
    # 2026-08-23: moved again - see the matching comment in run_articles_update.ps1. Old location
    # left untouched as a revert point; drop the " 1" suffix to point back at it.
    #
    # 2026-10-21, per user request: see the matching comment in run_articles_update.ps1 - the
    # shared folder now splits into "Old Process & Scope" / "New Process & Scope".
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1",
    [string]$SharedSubfolder   = "Old Process & Scope",
    # 2026-08-06: see the matching comment in run_articles_update.ps1 - renamed from
    # "(auto-managed)" to make it unmistakable this file isn't meant to be opened by a person,
    # warning leading the filename so it's the first thing seen.
    [string]$CanonicalFileName = "DO NOT OPEN - Reserved for Automated Updates - Expert Commentary Repository.xlsx",
    [string]$ReviewFileName    = "Expert Commentary Repository (review).xlsx",
    [string]$TeamCopyFileName = "Expert Commentary Repository (team copy).xlsx",
    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Repository\Expert Commentary Repository (review).xlsx",
    [string]$Python = "",

    # See the matching param comment in run_articles_update.ps1. Defaulting False means a missing
    # canonical workbook now fails loudly (reported below as FAILED, nothing overwritten) instead
    # of silently being replaced with an empty one - only pass $true for a genuine first-time
    # setup (Install-Teammate.ps1's own validation-run call site does this).
    [switch]$AllowCreateNewWorkbook,

    # See the matching param comment in run_articles_update.ps1 - a second, independent Teams
    # sync connection kept current the same non-fatal way as the review/team-copy/local mirror.
    [string]$SecondaryTeamsMirrorFolder = ""
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# 2026-10-15, per user request: the scheduled task now retries at 9:30/10:00/10:30 if the laptop
# was logged out at the original 9:00 run (see the matching Task Scheduler repetition change made
# alongside this). Those retries are real Task Scheduler trigger instants, so if 9:00 already ran
# cleanly while logged in, the 9:30/10:00/10:30 ticks would otherwise re-scrape everything for no
# reason. This marker makes each retry a fast no-op once a FULLY clean run (both pipelines OK) has
# already happened today - a partial failure (e.g. a transient SSL error on one site) does NOT set
# the marker, so a later retry still gets a real chance to catch up, which is the whole point.
$successMarkerPath = Join-Path $scriptRoot ".last_full_success_date.txt"
$today = Get-Date -Format "yyyy-MM-dd"
if ((Test-Path $successMarkerPath) -and ((Get-Content $successMarkerPath -Raw).Trim() -eq $today)) {
    Write-Host "Already completed a fully clean run today ($today) - skipping this retry."
    exit 0
}

$webArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    TeamCopyFileName  = $TeamCopyFileName
    LocalMirrorPath   = $LocalMirrorPath
    Python            = $Python
    AllowCreateNewWorkbook = $AllowCreateNewWorkbook
    SecondaryTeamsMirrorFolder = $SecondaryTeamsMirrorFolder
}
$vendorArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    TeamCopyFileName  = $TeamCopyFileName
    LocalMirrorPath   = $LocalMirrorPath
    Python            = $Python
    AllowCreateNewWorkbook = $AllowCreateNewWorkbook
    SecondaryTeamsMirrorFolder = $SecondaryTeamsMirrorFolder
}

$webOk = $false
$vendorOk = $false

Write-Host "==================== Web scraper (expert-commentary hub pages) ===================="
try {
    & (Join-Path $scriptRoot "run_articles_update.ps1") @webArgs
    $webOk = $true
} catch {
    Write-Host ("WEB SCRAPER FAILED: {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== Vendor email pipeline (Signum, Eurasia Group, ...) ===================="
try {
    & (Join-Path $scriptRoot "email-scraper\run_vendor_email_update.ps1") @vendorArgs
    $vendorOk = $true
} catch {
    Write-Host ("VENDOR EMAIL PIPELINE FAILED (non-fatal if you don't have Outlook/vendor access): {0}" -f $_.Exception.Message)
}

# 2026-08-23: rebuilds the team-install distributable zip from whatever this project folder
# currently contains, every day - so any code change made in ANY future session is automatically
# reflected in the shared Setup folder's zip within one day, with nobody needing to remember a
# manual rebuild step (this was missed for days after both the 2026-08-18 and 2026-08-23 path
# migrations). Deliberately non-fatal: a packaging hiccup should never block the real scraping/
# email work above.
$zipOk = $false
Write-Host "`n==================== Distributable zip (team-install package) ===================="
try {
    & (Join-Path $scriptRoot "team-install\Build-DistributableZip.ps1") -SharedLibraryRoot $SharedLibraryRoot -SharedSubfolder $SharedSubfolder
    $zipOk = $true
} catch {
    Write-Host ("ZIP REBUILD FAILED (non-fatal - shared workbook updates above are unaffected): {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== Summary ===================="
$webStatus = if ($webOk) { "OK" } else { "FAILED" }
$vendorStatus = if ($vendorOk) { "OK" } else { "FAILED" }
$zipStatus = if ($zipOk) { "OK" } else { "FAILED" }
Write-Host ("Web scraper:      {0}" -f $webStatus)
Write-Host ("Vendor email:     {0}" -f $vendorStatus)
Write-Host ("Distributable zip: {0}" -f $zipStatus)

if (-not $webOk -and -not $vendorOk) {
    throw "Both the web scraper and the vendor email pipeline failed - see the messages above."
}

if ($webOk -and $vendorOk) {
    Set-Content -Path $successMarkerPath -Value $today
}
