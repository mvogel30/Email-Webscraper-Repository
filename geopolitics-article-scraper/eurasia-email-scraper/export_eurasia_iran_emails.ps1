<#
Scans this mailbox (via Outlook COM — no password/API key, just automates the already-logged-in
Outlook desktop app, same technique as the Signum pipeline and the newsletter tool it was
adapted from) for Eurasia Group emails and writes matching messages to a JSON file.

Filter here is deliberately WIDE at export time: any sender on the eurasiagroup.net domain
(covers "EG Daily Brief" and other formats like "EG Update"), no subject filter — unlike Signum's
"Iran:" subject filter, Eurasia Group doesn't consistently tag Iran content in the subject line.
Iran-relevance is instead decided later, per-email or per-section, by eurasia_parser.py.

Outlook's own "Sync Issues" folder (an internal housekeeping area for sync conflicts) is skipped —
it can contain a duplicate copy of a message that already exists in its real folder.
#>
param(
    [string]$OutputJson = "$env:USERPROFILE\OneDrive - Deloitte (O365D)\Documents\Eurasia Iran Reports\daily_runs\eurasia_export.json",
    [int]$DaysBack = 14,
    [Nullable[datetime]]$Start,
    [Nullable[datetime]]$End
)

$ErrorActionPreference = "Stop"

$Start = if ($Start.HasValue) { $Start.Value } else { (Get-Date).AddDays(-1 * $DaysBack) }
$End = if ($End.HasValue) { $End.Value } else { Get-Date }

$targetDomain = "eurasiagroup.net"
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
                $subject = [string]$item.Subject
                if (-not $senderEmail -or $senderEmail.ToLowerInvariant() -notlike "*@$targetDomain") { continue }

                $messages.Add([pscustomobject]@{
                    file          = ""
                    entry_id      = [string]$item.EntryID
                    folder        = $folderPath
                    subject       = $subject
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

Write-Host ("Scanning for Eurasia Group emails from {0:o} to {1:o}" -f $Start, $End)
foreach ($root in @($namespace.Folders)) { Visit-Folder $root }

# De-duplicate on EntryID in case the same message is visible from two paths (e.g. a search
# folder overlapping a real folder) even after skipping Sync Issues.
$deduped = @($messages | Group-Object entry_id | ForEach-Object { $_.Group[0] })

$deduped | Sort-Object received_date | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputJson -Encoding UTF8

[pscustomobject]@{
    output = $OutputJson
    start = $Start.ToString("o")
    end = $End.ToString("o")
    count = $deduped.Count
} | ConvertTo-Json
