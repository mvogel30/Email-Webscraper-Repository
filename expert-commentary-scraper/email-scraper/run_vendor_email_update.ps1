<#
Orchestrator for the vendor email pipeline: exports matching Outlook emails from every
configured vendor (see vendors.json - currently Signum, Eurasia Group, Middle East Council on
Global Affairs, and Vriens & Partners; any sender at the configured domain, any subject), parses
them, and appends new rows into the SAME shared team workbook the web scraper writes to.

Different SOURCE from the web scraper (email vs web), so it stays a separate script/pipeline,
but shares the exact same shared-library resolution and output target — see
run_daily_article_scraper_update.ps1, one level up, whose defaults these mirror.

Requires Outlook desktop open and logged in when this runs (Outlook COM automation — no
password/API key stored anywhere; it just remote-controls your already-logged-in session). If
Outlook isn't available at all on this machine, this script exits non-zero and the caller
(run_daily_full_update.ps1) treats that as non-fatal — a teammate without vendor email access
simply gets no vendor-email rows, without breaking the web scraper.
#>
param(
    # 2026-08-18: see the matching comment in run_articles_update.ps1 - the SharePoint site was
    # restructured (a real move, not just a rename), so this now points at a shortcut added
    # directly at the "Webscraping & Email Scraping tool" folder level; SharedSubfolder is empty.
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

    # 2026-08-05: see the matching param comment in run_articles_update.ps1 - a second copy in
    # the same shared folder under a filename nobody opens in Excel, added because the review
    # file's updates weren't propagating to Teams/SharePoint for the user and teammates.
    [string]$TeamCopyFileName = "Expert Commentary Repository (team copy).xlsx",

    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Repository\Expert Commentary Repository (review).xlsx",

    # Widened from 3 to 7 on 2026-08-05: a 3-day window meant one missed run (laptop off,
    # Outlook closed, etc.) created a PERMANENT gap, since every later run only ever looks back
    # from "now" - this is exactly how a batch of real vendor emails (7/29-7/31 Eurasia Group
    # Daily Briefs, several Signum emails) went unscraped until a manual backfill found them.
    # A wider window costs nothing extra: append_rows() dedupes by URL/title+source+date, so
    # re-scanning already-captured emails on every run just skips them.
    [int]$DaysBack = 7,

    [string]$Python = "",

    # See the matching param comment in run_articles_update.ps1 - only pass $true for a genuine
    # first-time setup with no workbook ever created at this path yet.
    [switch]$AllowCreateNewWorkbook,

    # See the matching param comment in run_articles_update.ps1 - a second, independent Teams
    # sync connection this user actually browses to, kept current the same non-fatal way as
    # TeamCopyFileName/LocalMirrorPath. Empty (default) skips this entirely.
    [string]$SecondaryTeamsMirrorFolder = "",

    # See the matching param comment in run_articles_update.ps1 - New Process & Scope pass-
    # through, both default to "" (Old Process behavior, 100% unchanged).
    [string]$VendorsConfig = "",
    [string]$RowSchema = ""
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
$outputDir = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Vendor Emails\daily_runs"
$outputJson = Join-Path $outputDir ("vendor_export_{0}_{1}.json" -f $stamp, $env:COMPUTERNAME)
$logPath = Join-Path $outputDir "scheduled_task.log"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$transcriptStarted = $false
try {
    Start-Transcript -Path $logPath -Append | Out-Null
    $transcriptStarted = $true

    # Same hard-fail-if-not-synced rule as the web scraper: never silently write somewhere else.
    if (-not (Test-Path -LiteralPath $SharedLibraryRoot)) {
        throw "Shared library not found locally: '$SharedLibraryRoot'. In SharePoint/Teams, navigate INTO the 'Webscraping & Email Scraping tool' folder specifically (not just a parent folder) and click 'Add shortcut to OneDrive' there, then re-run."
    }
    $sharedFolder = Join-Path $SharedLibraryRoot $SharedSubfolder
    New-Item -ItemType Directory -Force -Path $sharedFolder | Out-Null
    $workbookPath = Join-Path $sharedFolder $CanonicalFileName
    $reviewPath = Join-Path $sharedFolder $ReviewFileName
    $teamCopyPath = Join-Path $sharedFolder $TeamCopyFileName

    $python = Resolve-Python -Override $Python
    "{0:o} Starting vendor email update. Workbook: {1}" -f (Get-Date), $workbookPath | Write-Host

    # See run_daily_article_scraper_update.ps1 for why $ErrorActionPreference is relaxed here:
    # a native process's (or, here, a nested script's) stderr text can otherwise become a
    # script-terminating error under PS 5.1 with Start-Transcript active, even on success.
    $previousEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    # export_vendor_emails.ps1 is a PowerShell SCRIPT (unlike the python.exe calls below, which
    # aren't subject to PowerShell's own parameter binder at all) - splatting a flat ARRAY here
    # binds POSITIONALLY, not by "-Name" token, which silently mis-bound DaysBack to the JSON path
    # string on this call's first real run. A HASHTABLE splat is what actually binds by name.
    $exportArgs = @{ OutputJson = $outputJson; DaysBack = $DaysBack }
    if ($VendorsConfig) { $exportArgs["VendorsConfig"] = $VendorsConfig }
    & (Join-Path $scriptRoot "export_vendor_emails.ps1") @exportArgs

    $appendExitCode = $null
    if (Test-Path -LiteralPath $outputJson) {
        $appendArgs = @("--input-json", $outputJson, "--workbook", $workbookPath, "--runner", "$env:COMPUTERNAME ($env:USERNAME)")
        if ($AllowCreateNewWorkbook) { $appendArgs += "--allow-create-new-workbook" }
        if ($VendorsConfig) { $appendArgs += @("--vendors-config", $VendorsConfig) }
        if ($RowSchema) { $appendArgs += @("--row-schema", $RowSchema) }
        & $python (Join-Path $scriptRoot "parse_and_append_vendor_emails.py") @appendArgs
        $appendExitCode = $LASTEXITCODE
    }

    $ErrorActionPreference = $previousEap
    if (-not (Test-Path -LiteralPath $outputJson)) { throw "Export did not create $outputJson (check Outlook is open and logged in, and vendor emails are reachable)" }
    if ($appendExitCode -ne 0) { throw "Append stage failed with exit code $appendExitCode" }

    "{0:o} Vendor email update completed successfully." -f (Get-Date) | Write-Host

    # --- Refresh the read-only review copy (non-fatal). Retries once after a short pause: the
    # only real reason this copy fails is someone having the file open in Excel, which often
    # clears within a few seconds, so a single retry catches that without holding up the run. ---
    for ($reviewAttempt = 1; $reviewAttempt -le 2; $reviewAttempt++) {
        try {
            if (Test-Path -LiteralPath $reviewPath) { attrib -R "$reviewPath" 2>$null }
            Copy-Item -LiteralPath $workbookPath -Destination $reviewPath -Force
            attrib +R "$reviewPath" 2>$null
            "{0:o} Review copy refreshed: {1}" -f (Get-Date), $reviewPath | Write-Host
            break
        } catch {
            if ($reviewAttempt -lt 2) {
                Start-Sleep -Seconds 5
                continue
            }
            "{0:o} WARNING: could not refresh review copy ({1}). Shared canonical is still up to date." -f (Get-Date), $_.Exception.Message | Write-Host
        }
    }
    try {
        Copy-Item -LiteralPath $workbookPath -Destination $teamCopyPath -Force
        "{0:o} Team copy refreshed: {1}" -f (Get-Date), $teamCopyPath | Write-Host
    } catch {
        "{0:o} WARNING: could not refresh team copy ({1}). Shared canonical is still up to date." -f (Get-Date), $_.Exception.Message | Write-Host
    }
    # --- Second Teams sync connection (non-fatal, opt-in) - see the matching, more detailed
    # comment and 2026-09-18 bug note in run_articles_update.ps1. Each file gets its own
    # try/catch so a lock on the review copy never also skips the (never-locked) team copy. ---
    if ($SecondaryTeamsMirrorFolder) {
        if (-not (Test-Path -LiteralPath $SecondaryTeamsMirrorFolder)) {
            "{0:o} WARNING: secondary Teams mirror folder not found locally ({1}) - this sync connection may itself be broken. Shared canonical is still up to date." -f (Get-Date), $SecondaryTeamsMirrorFolder | Write-Host
        } else {
            # 2026-10-21: see the matching, more detailed comment in run_articles_update.ps1 - this
            # folder's path plus $SharedSubfolder pushes the canonical filename's full path past
            # Windows' 260-char MAX_PATH limit (confirmed live), so Copy-Item is replaced with a
            # "\\?\"-prefixed .NET File.Copy here, which bypasses that limit reliably.
            $secondaryFolder = Join-Path $SecondaryTeamsMirrorFolder $SharedSubfolder
            New-Item -ItemType Directory -Force -Path $secondaryFolder | Out-Null
            function Copy-ItemLongPathSafe {
                param([string]$Source, [string]$Destination)
                [System.IO.File]::Copy("\\?\$Source", "\\?\$Destination", $true)
            }
            $secondaryCanonical = Join-Path $secondaryFolder $CanonicalFileName
            $secondaryReview = Join-Path $secondaryFolder $ReviewFileName
            $secondaryTeamCopy = Join-Path $secondaryFolder $TeamCopyFileName
            try {
                Copy-ItemLongPathSafe -Source $workbookPath -Destination $secondaryCanonical
                "{0:o} Secondary Teams mirror canonical refreshed: {1}" -f (Get-Date), $secondaryCanonical | Write-Host
            } catch {
                "{0:o} WARNING: could not refresh secondary Teams mirror canonical ({1})." -f (Get-Date), $_.Exception.Message | Write-Host
            }
            for ($secondaryReviewAttempt = 1; $secondaryReviewAttempt -le 2; $secondaryReviewAttempt++) {
                try {
                    if (Test-Path -LiteralPath $secondaryReview) { attrib -R "$secondaryReview" 2>$null }
                    Copy-ItemLongPathSafe -Source $workbookPath -Destination $secondaryReview
                    attrib +R "$secondaryReview" 2>$null
                    "{0:o} Secondary Teams mirror review copy refreshed: {1}" -f (Get-Date), $secondaryReview | Write-Host
                    break
                } catch {
                    if ($secondaryReviewAttempt -lt 2) {
                        Start-Sleep -Seconds 5
                        continue
                    }
                    "{0:o} WARNING: could not refresh secondary Teams mirror review copy ({1})." -f (Get-Date), $_.Exception.Message | Write-Host
                }
            }
            try {
                Copy-ItemLongPathSafe -Source $workbookPath -Destination $secondaryTeamCopy
                "{0:o} Secondary Teams mirror team copy refreshed: {1}" -f (Get-Date), $secondaryTeamCopy | Write-Host
            } catch {
                "{0:o} WARNING: could not refresh secondary Teams mirror team copy ({1})." -f (Get-Date), $_.Exception.Message | Write-Host
            }
        }
    }
    # --- Personal quick-look mirror, sourced from canonical directly (not the review copy) so a
    # locked review copy never causes this to silently propagate stale content instead. ---
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $LocalMirrorPath) | Out-Null
        Copy-Item -LiteralPath $workbookPath -Destination $LocalMirrorPath -Force
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
