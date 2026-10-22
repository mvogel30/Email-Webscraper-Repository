<#
Scans this mailbox (via Outlook COM - no password/API key, just automates the already-logged-in
Outlook desktop app, same technique as the sibling geopolitics-article-scraper tool) for emails
from any configured vendor domain, and writes matching messages to ONE JSON file tagged by which
vendor matched.

Filter is deliberately broad and domain-based, not per-person or per-subject: ANY sender at
signumglobal.com is a Signum email, ANY sender at eurasiagroup.net is a Eurasia Group email - no
subject filter, no "only Andrew Bishop" restriction. This is a deliberate change from the sibling
Iran-scraper tool, which locked Signum to one sender + an "Iran:" subject tag; here we want every
report from either firm, and the vendor list in vendors.json is the single source of truth for
"what counts as a Signum/Eurasia email" so adding a third vendor later is a one-line config change,
not a second copy of this whole script.

Outlook's own "Sync Issues" folder (an internal housekeeping area for sync conflicts) is skipped -
it can contain a duplicate copy of a message that already exists in its real folder.
#>
param(
    [string]$VendorsConfig = "$PSScriptRoot\vendors.json",
    [string]$OutputJson = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Expert Commentary Vendor Emails\daily_runs\vendor_export.json",
    [int]$DaysBack = 14,
    [Nullable[datetime]]$Start,
    [Nullable[datetime]]$End
)

$ErrorActionPreference = "Stop"

$Start = if ($Start.HasValue) { $Start.Value } else { (Get-Date).AddDays(-1 * $DaysBack) }
$End = if ($End.HasValue) { $End.Value } else { Get-Date }

$vendorsData = Get-Content -LiteralPath $VendorsConfig -Raw | ConvertFrom-Json
$vendors = @($vendorsData.vendors)
$messages = New-Object System.Collections.Generic.List[object]

function Get-OutlookApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        return New-Object -ComObject Outlook.Application
    }
}

function Get-SmtpAddress($mail) {
    try {
        if ($mail.SenderEmailType -eq "EX" -and $mail.Sender -ne $null) {
            $exchangeUser = $mail.Sender.GetExchangeUser()
            if ($exchangeUser -ne $null -and $exchangeUser.PrimarySmtpAddress) {
                return $exchangeUser.PrimarySmtpAddress
            }
        }
    } catch {}
    return [string]$mail.SenderEmailAddress
}

function Get-MatchingVendor($senderEmail) {
    if (-not $senderEmail) { return $null }
    $lower = $senderEmail.ToLowerInvariant()
    foreach ($vendor in $vendors) {
        # Matches the domain itself ("@cfr.org") AND any subdomain of it ("@e.cfr.org") - added
        # 2026-09-09 after finding a real CFR sender on "publications@e.cfr.org", a different
        # subdomain than the vendor's configured "cfr.org". Purely additive: can only catch MORE
        # legitimate senders under the same root domain, never a false match on an unrelated one.
        $domain = [regex]::Escape($vendor.domain.ToLowerInvariant())
        if ($lower -match ("@([a-z0-9-]+\.)*{0}$" -f $domain)) { return $vendor.name }
    }
    return $null
}

function Visit-Folder {
    param([object]$folder)

    $folderPath = ""
    try { $folderPath = [string]$folder.FolderPath } catch { $folderPath = [string]$folder.Name }
    if ($folderPath -match "Sync Issues") { return }

    $items = $null
    try { $items = $folder.Items } catch {
        Write-Warning ("Skipped Outlook folder {0}: unable to read items: {1}" -f $folderPath, $_.Exception.Message)
    }

    if ($items -ne $null) {
        $isSorted = $false
        try { $items.Sort("[ReceivedTime]", $true); $isSorted = $true } catch { $isSorted = $false }
        try { $itemCount = [int]$items.Count } catch { $itemCount = 0 }

        for ($i = 1; $i -le $itemCount; $i++) {
            try {
                $item = $items.Item($i)
                if ($item.Class -ne 43) { continue }
                if ($item.ReceivedTime -lt $Start) { if ($isSorted) { break } else { continue } }
                if ($item.ReceivedTime -ge $End) { continue }

                $senderEmail = Get-SmtpAddress $item
                $vendorName = Get-MatchingVendor $senderEmail
                if (-not $vendorName) { continue }

                $messages.Add([pscustomobject]@{
                    vendor        = $vendorName
                    entry_id      = [string]$item.EntryID
                    folder        = $folderPath
                    subject       = [string]$item.Subject
                    sender_name   = [string]$item.SenderName
                    sender_email  = $senderEmail
                    received_date = ([datetime]$item.ReceivedTime).ToString("o")
                    body          = [string]$item.Body
                }) | Out-Null
            } catch {
                Write-Warning ("Skipped an item in {0}: {1}" -f $folderPath, $_.Exception.Message)
            }
        }
    }

    try {
        foreach ($child in @($folder.Folders)) { Visit-Folder $child }
    } catch {}
}

$outputParent = Split-Path -Parent $OutputJson
if ($outputParent) { New-Item -ItemType Directory -Force -Path $outputParent | Out-Null }

$outlook = Get-OutlookApplication
$namespace = $outlook.GetNamespace("MAPI")

Write-Host ("Scanning for {0} emails from {1:o} to {2:o}" -f (($vendors | ForEach-Object { $_.name }) -join "/"), $Start, $End)
foreach ($root in @($namespace.Folders)) { Visit-Folder $root }

# De-duplicate on EntryID in case the same message is visible from two paths (e.g. a search
# folder overlapping a real folder) even after skipping Sync Issues.
$deduped = @($messages | Group-Object entry_id | ForEach-Object { $_.Group[0] })

$deduped | Sort-Object received_date | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputJson -Encoding UTF8

$byVendor = @{}
foreach ($vendor in $vendors) { $byVendor[$vendor.name] = @($deduped | Where-Object { $_.vendor -eq $vendor.name }).Count }

[pscustomobject]@{
    output = $OutputJson
    start = $Start.ToString("o")
    end = $End.ToString("o")
    count = $deduped.Count
    by_vendor = $byVendor
} | ConvertTo-Json
