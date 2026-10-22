<#
Single daily entry point: runs the web scraper (think-tank sites), the Signum Iran email
pipeline, and the Eurasia Group Iran email pipeline, all writing into the SAME shared team
workbook. This is what the scheduled task (and Install-Teammate.ps1, for teammates) should point
at, not any sub-script alone.

Each pipeline runs independently - if one fails (e.g. a teammate doesn't have Outlook, or
doesn't receive Andrew Bishop's Signum emails or Eurasia Group's emails, or a think-tank site is
temporarily down), the others still run and the shared workbook still gets whatever updates
succeeded. Only if ALL THREE fail does this script itself report failure.
#>
param(
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - 06 Sensing (Inputs)",
    [string]$SharedSubfolder   = "10 - Geopolitics Sensing & Scraping Tools\Iran Scraping example",
    [string]$CanonicalFileName = "Geopolitics Article Repository (auto-managed).xlsx",
    [string]$ReviewFileName    = "Geopolitics Article Repository (review).xlsx",
    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Geopolitics Article Repository\Geopolitics Article Repository (review).xlsx",
    [string]$Topic = "",
    [string]$Python = ""
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$webArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    LocalMirrorPath   = $LocalMirrorPath
    Topic             = $Topic
    Python            = $Python
}
$signumArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    LocalMirrorPath   = $LocalMirrorPath
    Python            = $Python
}
$eurasiaArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    LocalMirrorPath   = $LocalMirrorPath
    Python            = $Python
}

$webOk = $false
$signumOk = $false
$eurasiaOk = $false

Write-Host "==================== Web scraper (think-tank sites) ===================="
try {
    & (Join-Path $scriptRoot "run_daily_article_scraper_update.ps1") @webArgs
    $webOk = $true
} catch {
    Write-Host ("WEB SCRAPER FAILED: {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== Signum Iran email pipeline ===================="
try {
    & (Join-Path $scriptRoot "signum-email-scraper\run_signum_email_update.ps1") @signumArgs
    $signumOk = $true
} catch {
    Write-Host ("SIGNUM PIPELINE FAILED (non-fatal if you don't have Outlook/Signum access): {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== Eurasia Group Iran email pipeline ===================="
try {
    & (Join-Path $scriptRoot "eurasia-email-scraper\run_eurasia_email_update.ps1") @eurasiaArgs
    $eurasiaOk = $true
} catch {
    Write-Host ("EURASIA PIPELINE FAILED (non-fatal if you don't have Outlook/Eurasia Group access): {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== Summary ===================="
$webStatus = if ($webOk) { "OK" } else { "FAILED" }
$signumStatus = if ($signumOk) { "OK" } else { "FAILED" }
$eurasiaStatus = if ($eurasiaOk) { "OK" } else { "FAILED" }
Write-Host ("Web scraper:      {0}" -f $webStatus)
Write-Host ("Signum pipeline:  {0}" -f $signumStatus)
Write-Host ("Eurasia pipeline: {0}" -f $eurasiaStatus)

if (-not $webOk -and -not $signumOk -and -not $eurasiaOk) {
    throw "The web scraper, Signum pipeline, and Eurasia pipeline all failed - see the messages above."
}
