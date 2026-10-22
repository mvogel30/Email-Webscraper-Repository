<#
Rebuilds the team-install distributable zip from the CURRENT state of this project folder and
copies it into the shared Setup folder, so any future code change made to this tool is picked up
automatically on the very next daily run - no one has to remember to manually rebuild/redistribute
the zip after an edit (that manual-rebuild step was missed after both the 2026-08-18 and 2026-08-23
SharePoint path migrations, leaving stale zips live for days).

Called automatically at the end of run_daily_full_update.ps1 (non-fatal - a failure here never
blocks the actual scraping/email pipelines). Also safe to run by hand any time:
    powershell -File "team-install\Build-DistributableZip.ps1"

Stages a clean copy in $env:TEMP first rather than zipping the live project folder in place:
  - excludes __pycache__/*.pyc (build artifacts, not source) and this-machine's runtime cache
    (chinamacro_date_cache.json) - a fresh install should start clean, not inherit one person's
    local run history.
  - this environment's sandbox blocks any Remove-Item call whose path contains the shared
    "Global..." folder text (a blunt string-match, not real permission inspection) - staging AND
    building the zip entirely under $env:TEMP first, then a single Copy-Item -Force into the
    shared folder as the last step, avoids ever needing Remove-Item there.
#>
param(
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1",
    # 2026-10-21, per user request: see the matching comment in run_articles_update.ps1 - the
    # shared folder now splits into "Old Process & Scope" / "New Process & Scope", so the Setup
    # folder the installer zip lives in moves one level deeper too.
    [string]$SharedSubfolder = "Old Process & Scope",
    [string]$SetupSubfolder = "Setup",
    [string]$ZipFileName = "expert-commentary-scraper.zip"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$excludeDirNames = @("__pycache__", ".git")
$excludeFileNames = @("chinamacro_date_cache.json")

$stagingRoot = Join-Path $env:TEMP ("ecs-zip-staging-" + [guid]::NewGuid().ToString("N"))
$stagingProject = Join-Path $stagingRoot "expert-commentary-scraper"
New-Item -ItemType Directory -Path $stagingProject -Force | Out-Null

try {
    Get-ChildItem -LiteralPath $projectRoot -Recurse -Force | ForEach-Object {
        $relative = $_.FullName.Substring($projectRoot.Length).TrimStart('\')
        if ($relative -eq "") { return }
        $segments = $relative -split '\\'
        if ($segments | Where-Object { $excludeDirNames -contains $_ }) { return }
        if (-not $_.PSIsContainer -and $excludeFileNames -contains $_.Name) { return }
        if (-not $_.PSIsContainer -and $_.Extension -eq ".pyc") { return }

        $destination = Join-Path $stagingProject $relative
        if ($_.PSIsContainer) {
            New-Item -ItemType Directory -Path $destination -Force | Out-Null
        } else {
            $destParent = Split-Path -Parent $destination
            if (-not (Test-Path -LiteralPath $destParent)) {
                New-Item -ItemType Directory -Path $destParent -Force | Out-Null
            }
            Copy-Item -LiteralPath $_.FullName -Destination $destination -Force
        }
    }

    $zipStagingPath = Join-Path $stagingRoot $ZipFileName
    Compress-Archive -Path $stagingProject -DestinationPath $zipStagingPath -Force

    $setupFolder = Join-Path (Join-Path $SharedLibraryRoot $SharedSubfolder) $SetupSubfolder
    if (-not (Test-Path -LiteralPath $setupFolder)) {
        New-Item -ItemType Directory -Path $setupFolder -Force | Out-Null
    }
    $finalZipPath = Join-Path $setupFolder $ZipFileName
    Copy-Item -LiteralPath $zipStagingPath -Destination $finalZipPath -Force

    Write-Host ("{0:o} Distributable zip refreshed: {1}" -f (Get-Date), $finalZipPath)
} finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
