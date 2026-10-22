<#
Entry point for the New Process & Scope (see C:\Users\mivogel\.claude\plans\lively-roaming-moth.md)
- the article-source list being rebuilt ahead of a handover to the research department. Runs the
SAME underlying web scraper and vendor email pipeline as the Old Process (run_daily_full_update.ps1),
just pointed at a different source config, vendor config, workbook column layout, and shared-folder
location - no code was forked to build this, only new config files and a few pass-through params.

Deliberately simpler than the Old Process's orchestrator for now: canonical + review + team-copy +
secondary Teams/SharePoint mirror (added 2026-10-22 - that's the location the user and teammates
actually browse to day-to-day, not the personal OneDrive shortcut) + personal quick-look mirror,
but no distributable-zip rebuild step - this isn't being distributed to anyone yet. Add that back
later if/when this process is actually handed off.
#>
param(
    [string]$SharedLibraryRoot = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1",
    [string]$SharedSubfolder   = "New Process & Scope",
    [string]$CanonicalFileName = "DO NOT OPEN - Reserved for Automated Updates - Expert Commentary Repository.xlsx",
    [string]$ReviewFileName    = "Expert Commentary Repository (review).xlsx",

    # 2026-10-21, real incident: without an explicit override here, run_articles_update.ps1/
    # run_vendor_email_update.ps1 fall back to THEIR OWN default -LocalMirrorPath, which points at
    # the Old Process's shared personal-mirror file - the very first live run of this script
    # silently overwrote it with New Process data (caught and restored from the Old Process's real
    # canonical immediately after). Giving the New Process its own, clearly-named path here makes
    # that collision structurally impossible going forward.
    [string]$LocalMirrorPath = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Repository\Expert Commentary Repository (New Process review).xlsx",

    [string]$Python = "",

    # Only pass this for the very first run ever against a brand-new canonical workbook - same
    # guard the Old Process uses (see run_articles_update.ps1's own param comment).
    [switch]$AllowCreateNewWorkbook,

    [string]$SourcesConfig = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "sources_config.new.json"),
    [string]$VendorsConfig = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "email-scraper\vendors.new.json"),
    [string]$RowSchema = "new-structure",

    # 2026-10-22, per user report: the user (and teammates) actually browse to this second, team-
    # visible SharePoint location day-to-day - not the personal OneDrive shortcut above - so this
    # was the one place that actually mattered and it was empty. Same path the Old Process's own
    # scheduled task already passes; see the matching param in run_articles_update.ps1/
    # run_vendor_email_update.ps1 for why it's non-fatal/best-effort and long-path-safe.
    [string]$SecondaryTeamsMirrorFolder = "$env:USERPROFILE\Deloitte (O365D)\Global PP+G - Documents (1)\03 Cross-Team Projects\08 - Geopolitics Sensing & Scraping Tools\Webscraping & Email Scraping tool"
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Canonical + review is the main deliverable (see module docstring above). LocalMirrorPath is
# explicitly overridden too (see its own param comment - this is NOT optional, it's what prevents
# the Old Process's personal mirror from being overwritten). TeamCopyFileName is left at the
# sub-scripts' own default - it lands harmlessly inside this same isolated SharedSubfolder either
# way. SecondaryTeamsMirrorFolder (added 2026-10-22) points at the SAME root the Old Process's
# task uses, but the sub-scripts join it with $SharedSubfolder ("New Process & Scope") before
# writing, so it lands in its own isolated subfolder there too - no collision with the Old
# Process's "Old Process & Scope" equivalents.
$webArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    LocalMirrorPath   = $LocalMirrorPath
    Python            = $Python
    AllowCreateNewWorkbook = $AllowCreateNewWorkbook
    SourcesConfig     = $SourcesConfig
    RowSchema         = $RowSchema
    SecondaryTeamsMirrorFolder = $SecondaryTeamsMirrorFolder
}
$vendorArgs = @{
    SharedLibraryRoot = $SharedLibraryRoot
    SharedSubfolder   = $SharedSubfolder
    CanonicalFileName = $CanonicalFileName
    ReviewFileName    = $ReviewFileName
    LocalMirrorPath   = $LocalMirrorPath
    Python            = $Python
    AllowCreateNewWorkbook = $AllowCreateNewWorkbook
    VendorsConfig     = $VendorsConfig
    RowSchema         = $RowSchema
    SecondaryTeamsMirrorFolder = $SecondaryTeamsMirrorFolder
}

$webOk = $false
$vendorOk = $false

Write-Host "==================== New Process & Scope: web scraper ===================="
try {
    & (Join-Path $scriptRoot "run_articles_update.ps1") @webArgs
    $webOk = $true
} catch {
    Write-Host ("WEB SCRAPER FAILED: {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== New Process & Scope: vendor email pipeline ===================="
try {
    & (Join-Path $scriptRoot "email-scraper\run_vendor_email_update.ps1") @vendorArgs
    $vendorOk = $true
} catch {
    Write-Host ("VENDOR EMAIL PIPELINE FAILED (non-fatal if you don't have Outlook/vendor access): {0}" -f $_.Exception.Message)
}

Write-Host "`n==================== Summary ===================="
$webStatus = if ($webOk) { "OK" } else { "FAILED" }
$vendorStatus = if ($vendorOk) { "OK" } else { "FAILED" }
Write-Host ("Web scraper:  {0}" -f $webStatus)
Write-Host ("Vendor email: {0}" -f $vendorStatus)

if (-not $webOk -and -not $vendorOk) {
    throw "Both the web scraper and the vendor email pipeline failed - see the messages above."
}
