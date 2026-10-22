<#
Orchestrator for the Eurasia Group Iran email pipeline: exports Eurasia Group emails from
Outlook, parses out the Iran-relevant content (whole emails or, for "EG Daily Brief", just the
Iran-relevant sections), and appends new rows into the SAME shared team workbook the web scraper
and Signum pipeline write to — by design (the user explicitly asked for all pipelines' findings
to land in one Excel).

Different SOURCE from the web scraper (email vs web) and a different mailbox/sender than Signum,
so it stays a separate script/pipeline, but shares the exact same shared-library resolution and
output target — see run_daily_article_scraper_update.ps1, two levels up, whose defaults these
mirror.

Requires Outlook desktop open and logged in when this runs (Outlook COM automation — no
password/API key stored anywhere; it just remote-controls your already-logged-in session). If
Outlook isn't available at all on this machine, or the user doesn't receive Eurasia Group emails,
this script exits non-zero and the caller (run_daily_full_update.ps1) treats that as non-fatal.
#>
param(
    # See run_daily_article_scraper_update.ps1 for why this default matters: it must be the
    # folder that resulted from "Add shortcut to OneDrive" clicked on "06 Sensing (Inputs)"
    # itself, which lands under your personal OneDrive folder — not a separate "Deloitte (O365D)"
    # root, which can silently be a disconnected local-only folder with no real sync registration.
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - 06 Sensing (Inputs)",
    [string]$SharedSubfolder   = "10 - Geopolitics Sensing & Scraping Tools\Iran Scraping example",
    [string]$CanonicalFileName = "Geopolitics Article Repository (auto-managed).xlsx",
    [string]$ReviewFileName    = "Geopolitics Article Repository (review).xlsx",
    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Geopolitics Article Repository\Geopolitics Article Repository (review).xlsx",
    [int]$DaysBack = 3,
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Resolve-Python {
    param([string]$Override)
    if ($Override -and (Test-Path -LiteralPath $Override)) { return $Override }
    foreach ($dir in Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending) {
        $candidate = Join-Path $dir.FullName "python.exe"
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    $onPath = (Get-Command python -ErrorAction SilentlyContinue)
    if ($onPath -and $onPath.Source -notlike "*WindowsApps*") { return $onPath.Source }
    $pyLauncher = (Get-Command py -ErrorAction SilentlyContinue)
    if ($pyLauncher) { return $pyLauncher.Source }
    throw "No Python interpreter found. Install Python 3 or pass -Python <path to python.exe>."
}

$stamp = (Get-Date).ToString("yyyy-MM-dd_HHmm")
$outputDir = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Eurasia Iran Reports\daily_runs"
$outputJson = Join-Path $outputDir ("eurasia_export_{0}_{1}.json" -f $stamp, $env:COMPUTERNAME)
$logPath = Join-Path $outputDir "scheduled_task.log"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$transcriptStarted = $false
try {
    Start-Transcript -Path $logPath -Append | Out-Null
    $transcriptStarted = $true

    # Same hard-fail-if-not-synced rule as the web scraper: never silently write somewhere else.
    if (-not (Test-Path -LiteralPath $SharedLibraryRoot)) {
        throw "Shared library not found locally: '$SharedLibraryRoot'. In SharePoint/Teams, navigate INTO the '06 Sensing (Inputs)' folder specifically (not just the parent library) and click 'Add shortcut to OneDrive' there, then re-run."
    }
    $sharedFolder = Join-Path $SharedLibraryRoot $SharedSubfolder
    New-Item -ItemType Directory -Force -Path $sharedFolder | Out-Null
    $workbookPath = Join-Path $sharedFolder $CanonicalFileName
    $reviewPath = Join-Path $sharedFolder $ReviewFileName

    $python = Resolve-Python -Override $Python
    "{0:o} Starting Eurasia Group Iran email update. Workbook: {1}" -f (Get-Date), $workbookPath | Write-Host

    # See run_daily_article_scraper_update.ps1 for why $ErrorActionPreference is relaxed here:
    # a native process's (or, here, a nested script's) stderr text can otherwise become a
    # script-terminating error under PS 5.1 with Start-Transcript active, even on success.
    $previousEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    & (Join-Path $scriptRoot "export_eurasia_iran_emails.ps1") -OutputJson $outputJson -DaysBack $DaysBack

    $appendExitCode = $null
    if (Test-Path -LiteralPath $outputJson) {
        & $python (Join-Path $scriptRoot "append_eurasia_workbook.py") --input-json $outputJson --workbook $workbookPath --runner "$env:COMPUTERNAME ($env:USERNAME)"
        $appendExitCode = $LASTEXITCODE
    }

    $ErrorActionPreference = $previousEap
    if (-not (Test-Path -LiteralPath $outputJson)) { throw "Export did not create $outputJson (check Outlook is open and logged in, and Eurasia Group's emails are reachable)" }
    if ($appendExitCode -ne 0) { throw "Append stage failed with exit code $appendExitCode" }

    "{0:o} Eurasia Group Iran email update completed successfully." -f (Get-Date) | Write-Host

    # --- Refresh the read-only review copy + personal mirror (non-fatal) ---
    try {
        if (Test-Path -LiteralPath $reviewPath) { attrib -R "$reviewPath" 2>$null }
        Copy-Item -LiteralPath $workbookPath -Destination $reviewPath -Force
        attrib +R "$reviewPath" 2>$null
        "{0:o} Review copy refreshed: {1}" -f (Get-Date), $reviewPath | Write-Host
    } catch {
        "{0:o} WARNING: could not refresh review copy ({1}). Shared canonical is still up to date." -f (Get-Date), $_.Exception.Message | Write-Host
    }
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LocalMirrorPath) | Out-Null
        Copy-Item -LiteralPath $reviewPath -Destination $LocalMirrorPath -Force
        "{0:o} Local mirror refreshed: {1}" -f (Get-Date), $LocalMirrorPath | Write-Host
    } catch {
        "{0:o} NOTE: local mirror skipped ({1})." -f (Get-Date), $_.Exception.Message | Write-Host
    }
} catch {
    "{0:o} FAILED: {1}" -f (Get-Date), $_.Exception.Message | Write-Host
    throw
} finally {
    if ($transcriptStarted) { Stop-Transcript | Out-Null }
}
