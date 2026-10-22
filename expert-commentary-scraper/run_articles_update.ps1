param(
    # Canonical workbook lives IN the team's shared, synced library. If your local OneDrive names
    # this library differently, override -SharedLibraryRoot at install time.
    #
    # 2026-08-18: the SharePoint site was restructured (top folder moved/renamed from
    # "06 Sensing (Inputs)" to "00 Sensing (Inputs)", confirmed a real move, not just a rename -
    # the old local shortcut kept losing files live as OneDrive reconciled with the new remote
    # state). Re-pointed at a shortcut added directly at the "Webscraping & Email Scraping tool"
    # folder level (one click, done once on whichever machine runs this) rather than at a parent
    # folder, so SharedSubfolder is now empty - no extra nesting to resolve underneath it.
    #
    # 2026-08-23: moved again - the team's visible Teams/SharePoint location is now
    # "03 Cross-Team Projects > 08 - Geopolitics Sensing & Scraping Tools > Webscraping & Email
    # Scraping tool". A shortcut to that folder was added to OneDrive; because a folder with the
    # same local name already existed (the old location, kept in place as a revert point),
    # OneDrive suffixed the new one " 1". The old location's files were left untouched on purpose
    # - to revert, just point -SharedLibraryRoot back at the old (non-" 1") path.
    #
    # 2026-10-21, per user request: the shared folder now splits into "Old Process & Scope" (the
    # live pipeline you're reading right now) and "New Process & Scope" (reserved for whatever
    # replaces it) - this default is what makes the daily run land in the right one without
    # anyone needing to pass an extra flag.
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1",
    [string]$SharedSubfolder   = "Old Process & Scope",
    # 2026-08-06: renamed from "(auto-managed)" to make it unmistakable this file is not meant
    # to be opened by a person - scripts write here directly and expect exclusive access. The
    # warning leads the filename (not trailing in parens) so it's the first thing seen when
    # browsing the shared folder.
    [string]$CanonicalFileName = "DO NOT OPEN - Reserved for Automated Updates - Expert Commentary Repository.xlsx",
    [string]$ReviewFileName    = "Expert Commentary Repository (review).xlsx",

    # 2026-08-05: a second copy in the SAME shared folder, under a filename nobody ever opens in
    # Excel. Added because the review file's updates weren't propagating to Teams/SharePoint for
    # the user and teammates viewing it through Teams (while Explorer, reading the local OneDrive
    # cache directly, showed it as current) - the most likely cause is the review file being kept
    # open in Excel by someone, which blocks OneDrive from syncing THAT file specifically. A copy
    # nobody opens should always be free to sync.
    [string]$TeamCopyFileName = "Expert Commentary Repository (team copy).xlsx",

    # Optional personal quick-look mirror in your own OneDrive. Nice-to-have only; never the
    # source of truth, and never required for the run to succeed.
    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Repository\Expert Commentary Repository (review).xlsx",

    [string]$OutputDirectory = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Repository\daily_runs",
    [string]$Python = "",                       # empty => auto-discover
    [string]$RunnerName = "$env:COMPUTERNAME ($env:USERNAME)",  # logged in the workbook's Run Log

    # 2026-09-18, real incident: a transient OneDrive sync gap made the canonical workbook briefly
    # report as missing mid-run, which used to silently create a brand-new EMPTY workbook in its
    # place and save over the real one - real production data loss (1,385+ rows). Fixed at the
    # source (append_article_repository_workbook.get_or_create_workbook) to require this flag
    # before ever creating a new workbook; defaulting False here means every normal run now fails
    # loudly on a missing canonical file instead of silently replacing it. Only pass $true for a
    # genuine first-time setup (see Install-Teammate.ps1's one validation-run call site).
    [switch]$AllowCreateNewWorkbook,

    # 2026-09-18, real incident: discovered this machine has TWO separate local sync connections
    # to what the user considers "the same" Teams folder - this OneDrive personal-shortcut one
    # (SharedLibraryRoot above, what every script actually reads/writes) and a second, independent
    # SharePoint-site-sync one the user actually browses to day-to-day. The second one silently
    # stopped syncing around 2026-09-15 and was NEVER touched by any script, so it drifted stale
    # without anyone noticing until the user pointed it out directly. Rather than trust OneDrive's
    # own dual-sync to eventually reconcile them (already proven unreliable), every run now also
    # copies canonical/review/team-copy directly into this second location - same non-fatal
    # best-effort pattern as TeamCopyFileName/LocalMirrorPath above. Empty string (default) skips
    # this entirely, e.g. on a teammate's machine that only has the one sync connection.
    [string]$SecondaryTeamsMirrorFolder = "",

    # 2026-10-21, New Process & Scope: lets the SAME pipeline run against a different source list
    # and workbook column layout without forking any code - both default to "" (don't pass the
    # underlying --config/--row-schema flag at all), which keeps the Old Process's behavior 100%
    # unchanged. Pass -SourcesConfig pointing at sources_config.new.json and -RowSchema
    # "new-structure" for the new process - see run_new_process_update.ps1.
    [string]$SourcesConfig = "",
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
    # NOT fall back to creating a private local file here — fail clearly instead.
    if (-not (Test-Path -LiteralPath $SharedLibraryRoot)) {
        throw "Shared library not found locally: '$SharedLibraryRoot'. In SharePoint/Teams, navigate INTO the 'Webscraping & Email Scraping tool' folder specifically (not just a parent folder) and click 'Add shortcut to OneDrive' there, then re-run."
    }
    $sharedFolder = Join-Path $SharedLibraryRoot $SharedSubfolder
    New-Item -ItemType Directory -Force -Path $sharedFolder | Out-Null
    $workbookPath = Join-Path $sharedFolder $CanonicalFileName
    $reviewPath = Join-Path $sharedFolder $ReviewFileName
    $teamCopyPath = Join-Path $sharedFolder $TeamCopyFileName

    $python = Resolve-Python -Override $Python
    "{0:o} Starting expert-commentary scraper update. Runner: {1}. Workbook: {2}" -f (Get-Date), $RunnerName, $workbookPath | Write-Host
    "{0:o} Python: {1}" -f (Get-Date), ($python -join ' ') | Write-Host

    # Native processes below may print to stderr (e.g. a per-source WARNING) without that
    # meaning failure - see run_daily_article_scraper_update.ps1 in the sibling tool for why
    # $ErrorActionPreference is relaxed around these calls specifically.
    $previousEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    # --- Fetch stage --- (writes a uniquely-named JSON; never contends with anyone else)
    $fetchArgs = @("--output-json", $outputJson)
    if ($SourcesConfig) { $fetchArgs += @("--config", $SourcesConfig) }
    & $python (Join-Path $scriptRoot "fetch_articles.py") @fetchArgs
    $fetchExitCode = $LASTEXITCODE

    # --- Append stage --- (the one step that touches the shared workbook; retries through
    # brief OneDrive sync/Excel locks — see retry_on_lock in append_article_repository_workbook.py)
    $appendExitCode = $null
    if ($fetchExitCode -eq 0) {
        $appendArgs = @("--input-json", $outputJson, "--workbook", $workbookPath, "--runner", $RunnerName)
        if ($AllowCreateNewWorkbook) { $appendArgs += "--allow-create-new-workbook" }
        if ($RowSchema) { $appendArgs += @("--row-schema", $RowSchema) }
        & $python (Join-Path $scriptRoot "append_article_repository_workbook.py") @appendArgs
        $appendExitCode = $LASTEXITCODE
    }

    $ErrorActionPreference = $previousEap
    if ($fetchExitCode -ne 0) { throw "Fetch stage failed with exit code $fetchExitCode" }
    if (-not (Test-Path -LiteralPath $outputJson)) { throw "Fetch did not create $outputJson" }
    if ($appendExitCode -ne 0) { throw "Append stage failed with exit code $appendExitCode" }

    "{0:o} Repository update completed successfully." -f (Get-Date) | Write-Host

    # --- Publish the read-only review copy (non-fatal) — the file people actually open.
    # Retries once after a short pause: the only real reason this copy fails is someone having
    # the file open in Excel, which often clears within a few seconds (e.g. a stray background
    # window), so a single retry catches that without holding up the rest of the run for long. ---
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

    # --- Team copy in the same shared folder (non-fatal) — see TeamCopyFileName param comment.
    # Never opened by anyone, so it should always be free for OneDrive to sync up to Teams even
    # when the review file above is stuck because someone has it open in Excel. ---
    try {
        Copy-Item -LiteralPath $workbookPath -Destination $teamCopyPath -Force
        "{0:o} Team copy refreshed: {1}" -f (Get-Date), $teamCopyPath | Write-Host
    } catch {
        "{0:o} WARNING: could not refresh team copy ({1}). Shared canonical is still up to date." -f (Get-Date), $_.Exception.Message | Write-Host
    }

    # --- Second Teams sync connection (non-fatal, opt-in via -SecondaryTeamsMirrorFolder) - see
    # the param comment above for why this exists. Skipped entirely when not passed. Each file
    # copied in its OWN try/catch (real bug found and fixed 2026-09-18: an earlier version put all
    # three in one try block, so a lock on the review copy - the one file people might have open -
    # silently skipped the team-copy refresh too, even though team-copy is never locked and should
    # always succeed independently, same reasoning as the primary location's own three separate
    # try blocks below). Review copy also gets the same retry-once-after-a-pause as the primary
    # location's review copy, for the same reason (a transient Excel/OneDrive lock often clears
    # within seconds). ---
    if ($SecondaryTeamsMirrorFolder) {
        if (-not (Test-Path -LiteralPath $SecondaryTeamsMirrorFolder)) {
            "{0:o} WARNING: secondary Teams mirror folder not found locally ({1}) - this sync connection may itself be broken. Shared canonical is still up to date." -f (Get-Date), $SecondaryTeamsMirrorFolder | Write-Host
        } else {
            # 2026-10-21: this folder's own path is already long (Teams/SharePoint site nesting),
            # and adding $SharedSubfolder ("Old Process & Scope") on top pushed the canonical
            # file's full path past Windows' 260-char MAX_PATH limit - confirmed live (261 chars):
            # Copy-Item/Move-Item both fail there with a misleading "Could not find a part of the
            # path". The "\\?\" long-path prefix with .NET's File.Copy bypasses this reliably
            # (confirmed live too) - used for all three files here, not just the one that happens
            # to be long today, so a future rename doesn't quietly reintroduce the same failure.
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

    # --- Personal quick-look mirror (non-fatal, optional convenience). Sourced from the
    # canonical workbook directly, NOT the review copy above - if the review copy's own refresh
    # just failed (still locked after the retry), copying from it here would silently propagate
    # stale content into the mirror while claiming success. Canonical itself is never locked by
    # someone viewing the review copy, since it's a different file. ---
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
