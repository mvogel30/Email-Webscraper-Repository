param(
    # Canonical workbook now lives IN the team's shared, synced library — this is the one true
    # file every teammate's install writes to (not a private-per-person copy). It must resolve
    # to the SAME physical library for everyone, so if your local OneDrive names this library
    # differently, override -SharedLibraryRoot at install time (see Install-Teammate.ps1).
    #
    # IMPORTANT (found 2026-07-28): the library MUST be added via "Add shortcut to OneDrive" on
    # the "06 Sensing (Inputs)" folder itself (not the parent "Global PP+G - Documents" library
    # root). A shortcut added at the wrong level, or a stale/broken shortcut re-added incorrectly,
    # can silently create a local folder that LOOKS right (same file names, old content visible)
    # but has NO real OneDrive sync registration — writes just accumulate locally and never reach
    # SharePoint or teammates, with no error at all. Verified correct path lands under your
    # personal OneDrive folder, not a separate "Deloitte (O365D)" root.
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - 06 Sensing (Inputs)",
    [string]$SharedSubfolder   = "10 - Geopolitics Sensing & Scraping Tools\Iran Scraping example",
    [string]$CanonicalFileName = "Geopolitics Article Repository (auto-managed).xlsx",
    [string]$ReviewFileName    = "Geopolitics Article Repository (review).xlsx",

    # Optional personal quick-look mirror in your own OneDrive. Nice-to-have only; never the
    # source of truth, and never required for the run to succeed.
    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Geopolitics Article Repository\Geopolitics Article Repository (review).xlsx",

    [string]$OutputDirectory = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Geopolitics Article Repository\daily_runs",
    [string]$Topic = "",                        # empty => use active_topic from topics.json
    [string]$Python = "",                       # empty => auto-discover
    [string]$RunnerName = "$env:COMPUTERNAME ($env:USERNAME)"   # logged in the workbook's Run Log
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
    if ($pyLauncher) { return $pyLauncher.Source }   # Windows 'py' launcher (defaults to newest Py3)
    throw "No Python interpreter found. Install Python 3 or pass -Python <path to python.exe>."
}

$stamp = (Get-Date).ToString("yyyy-MM-dd_HHmm")
$outputJson = Join-Path $OutputDirectory ("articles_{0}_{1}.json" -f $stamp, $env:COMPUTERNAME)
$logPath = Join-Path $OutputDirectory "scheduled_task.log"
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$transcriptStarted = $false
try {
    Start-Transcript -Path $logPath -Append | Out-Null
    $transcriptStarted = $true

    # The shared library MUST be synced locally before we touch anything. We deliberately do
    # NOT fall back to creating a private local file here: silently writing somewhere else would
    # fork this machine's data away from the team's one shared workbook without telling anyone
    # loudly enough to notice. Fail clearly instead.
    if (-not (Test-Path -LiteralPath $SharedLibraryRoot)) {
        throw "Shared library not found locally: '$SharedLibraryRoot'. In SharePoint/Teams, navigate INTO the '06 Sensing (Inputs)' folder specifically (not just the parent library) and click 'Add shortcut to OneDrive' there, then re-run. (See TEAM-INSTALL-README.md.)"
    }
    $sharedFolder = Join-Path $SharedLibraryRoot $SharedSubfolder
    New-Item -ItemType Directory -Force -Path $sharedFolder | Out-Null
    $workbookPath = Join-Path $sharedFolder $CanonicalFileName
    $reviewPath = Join-Path $sharedFolder $ReviewFileName

    $python = Resolve-Python -Override $Python
    "{0:o} Starting article scraper update. Runner: {1}. Workbook: {2}" -f (Get-Date), $RunnerName, $workbookPath | Write-Host
    "{0:o} Python: {1}" -f (Get-Date), ($python -join ' ') | Write-Host

    # Native processes below may print to stderr (e.g. a per-source WARNING) without that
    # meaning failure. Under $ErrorActionPreference = "Stop" (plus Start-Transcript active),
    # PowerShell 5.1 can treat a native process's stderr text as a script-terminating error
    # even when it exits 0 — which would abort this whole run over a single non-fatal source
    # hiccup. Relax to 'Continue' for these calls specifically; $LASTEXITCODE is still checked
    # explicitly right after, so a genuine failure (non-zero exit) is still caught.
    $previousEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    # --- Fetch stage --- (writes a uniquely-named JSON; never contends with anyone else)
    $fetchArgs = @((Join-Path $scriptRoot "fetch_thinktank_articles.py"), "--output-json", $outputJson)
    if ($Topic) { $fetchArgs += @("--topic", $Topic) }
    & $python @fetchArgs
    $fetchExitCode = $LASTEXITCODE

    # --- Append stage --- (the one step that touches the shared workbook; retries through
    # brief OneDrive sync/Excel locks — see retry_on_lock in append_article_repository_workbook.py)
    $appendExitCode = $null
    if ($fetchExitCode -eq 0) {
        & $python (Join-Path $scriptRoot "append_article_repository_workbook.py") --input-json $outputJson --workbook $workbookPath --runner $RunnerName
        $appendExitCode = $LASTEXITCODE
    }

    $ErrorActionPreference = $previousEap
    if ($fetchExitCode -ne 0) { throw "Fetch stage failed with exit code $fetchExitCode" }
    if (-not (Test-Path -LiteralPath $outputJson)) { throw "Fetch did not create $outputJson" }
    if ($appendExitCode -ne 0) { throw "Append stage failed with exit code $appendExitCode" }

    "{0:o} Repository update completed successfully." -f (Get-Date) | Write-Host

    # --- Publish the read-only review copy (non-fatal) — the file people actually open ---
    try {
        if (Test-Path -LiteralPath $reviewPath) { attrib -R "$reviewPath" 2>$null }
        Copy-Item -LiteralPath $workbookPath -Destination $reviewPath -Force
        attrib +R "$reviewPath" 2>$null
        "{0:o} Review copy refreshed: {1}" -f (Get-Date), $reviewPath | Write-Host
    } catch {
        "{0:o} WARNING: could not refresh review copy ({1}). Shared canonical is still up to date." -f (Get-Date), $_.Exception.Message | Write-Host
    }

    # --- Personal quick-look mirror (non-fatal, optional convenience) ---
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
