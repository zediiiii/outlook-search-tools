<#
.SYNOPSIS
    Pulls every email exchanged with a given person (sent or received) over the last N days
    from the local Outlook desktop profile, and writes the full text to a file.

.DESCRIPTION
    Walks all mail folders in your Outlook mailbox (Inbox, Sent Items, and any subfolders
    you file mail into), finds every message where the person appears as sender, To, or CC,
    and dumps the complete body text in chronological order. Nothing is summarized or
    truncated -- the output is meant to be fed to an LLM for a catch-up summary.

.PARAMETER Person
    One or more match tokens: an email address, a partial address, or a display name.
    Matching is case-insensitive substring against sender + all recipients.
    Example: -Person jdoe@example.org
    Example: -Person 'jdoe@example.org','Jane Doe'

.PARAMETER Days
    How far back to look. Default 180.

.PARAMETER Since
    Explicit start date, overrides -Days. Example: -Since 2026-01-01

.PARAMETER OutFile
    Where to write the transcript. Defaults to .\context\<person>_<date>.md

.PARAMETER IncludeDeleted
    Also scan Deleted Items, Junk Email and Drafts (skipped by default).

.PARAMETER AllStores
    Also scan other mailboxes / PSTs / online archives attached to the profile,
    not just your primary mailbox.

.PARAMETER TrimQuoted
    Cut the quoted reply history off the bottom of each message. Makes the file far
    smaller but you lose text. Off by default -- everything is included.

.PARAMETER DirectOnly
    Only keep true back-and-forth: messages the person sent, or that you sent to them.
    Drops third-party blasts where you and they merely shared a distribution list.

.PARAMETER PassThru
    Also return an object with the output file path and every message found, so another
    script can show them. Used by FADOE.ps1.

.EXAMPLE
    .\Get-EmailContext.ps1 -Person jdoe@example.org -Days 180

.EXAMPLE
    .\Get-EmailContext.ps1 -Person 'jsmith@example.edu' -Since 2026-01-01 -OutFile C:\temp\jsmith.md
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Person,

    [int]      $Days = 180,
    [datetime] $Since,
    [string]   $OutFile,
    [switch]   $IncludeDeleted,
    [switch]   $AllStores,
    [switch]   $TrimQuoted,
    [switch]   $DirectOnly,
    [switch]   $PassThru
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- setup ----
if (-not $PSBoundParameters.ContainsKey('Since')) {
    $Since = (Get-Date).Date.AddDays(-$Days)
}
$tokens = $Person | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ }

Write-Host "Person   : $($Person -join ', ')"
Write-Host "Since    : $($Since.ToString('yyyy-MM-dd'))"

try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns      = $outlook.GetNamespace('MAPI')
} catch {
    Write-Error "Could not connect to Outlook. Make sure the classic Outlook desktop app is installed and your mail profile is set up. ($_)"
    exit 1
}

# My own addresses, so we can label direction.
$myAddresses = New-Object System.Collections.Generic.HashSet[string]
try {
    $me = $ns.CurrentUser.AddressEntry
    foreach ($a in @($me.Address, $me.GetExchangeUser().PrimarySmtpAddress)) {
        if ($a) { [void]$myAddresses.Add($a.ToLowerInvariant()) }
    }
} catch { }
try {
    foreach ($acct in $ns.Accounts) {
        if ($acct.SmtpAddress) { [void]$myAddresses.Add($acct.SmtpAddress.ToLowerInvariant()) }
    }
} catch { }
Write-Host "You      : $(($myAddresses) -join ', ')"

$PR_SMTP = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$smtpCache = @{}

function Resolve-Smtp {
    param($AddressEntry, [string]$Fallback)
    if (-not $AddressEntry) { return $Fallback }
    $key = $Fallback
    if ($key -and $smtpCache.ContainsKey($key)) { return $smtpCache[$key] }
    $result = $Fallback
    try {
        $t = $AddressEntry.PropertyAccessor.GetProperty($PR_SMTP)
        if ($t) { $result = $t }
    } catch {
        try {
            $eu = $AddressEntry.GetExchangeUser()
            if ($eu -and $eu.PrimarySmtpAddress) { $result = $eu.PrimarySmtpAddress }
        } catch { }
    }
    if ($key) { $smtpCache[$key] = $result }
    return $result
}

# --------------------------------------------------------- folder walk ----
$skipNames = @()
if (-not $IncludeDeleted) {
    $skipNames = @('Deleted Items', 'Junk Email', 'Junk E-mail', 'Conversation History',
                   'Sync Issues', 'RSS Feeds', 'RSS Subscriptions', 'Drafts', 'Outbox')
}

$folders = New-Object System.Collections.ArrayList

function Add-Folders {
    param($Folder, [int]$Depth = 0)
    if ($Depth -gt 12) { return }
    if ($skipNames -contains $Folder.Name) { return }
    try {
        # 0 = olMailItem default item type
        if ($Folder.DefaultItemType -eq 0) { [void]$folders.Add($Folder) }
    } catch { }
    try {
        foreach ($sub in $Folder.Folders) { Add-Folders -Folder $sub -Depth ($Depth + 1) }
    } catch { }
}

$stores = @()
if ($AllStores) {
    foreach ($s in $ns.Stores) { $stores += $s }
} else {
    $stores += $ns.GetDefaultFolder(6).Store   # 6 = olFolderInbox
}

foreach ($store in $stores) {
    try { Add-Folders -Folder $store.GetRootFolder() }
    catch { Write-Warning "Skipped store '$($store.DisplayName)': $_" }
}
Write-Host "Folders  : $($folders.Count) mail folder(s) to scan"

# ------------------------------------------------------------ scanning ----
$stamp = $Since.ToString('yyyy-MM-dd HH:mm')
$dasl  = '@SQL=("urn:schemas:httpmail:datereceived" >= ''' + $stamp + ''' OR "urn:schemas:httpmail:date" >= ''' + $stamp + ''')'

$hits = New-Object System.Collections.ArrayList
$seen = New-Object System.Collections.Generic.HashSet[string]
$fi = 0

foreach ($folder in $folders) {
    $fi++
    Write-Progress -Activity 'Scanning Outlook' -Status $folder.FolderPath -PercentComplete (100 * $fi / [Math]::Max($folders.Count, 1))

    try   { $items = $folder.Items.Restrict($dasl) }
    catch { $items = $folder.Items }

    foreach ($item in $items) {
        try {
            if ($item.Class -ne 43) { continue }   # 43 = olMail

            $when = $null
            try { if ($item.SentOn)       { $when = $item.SentOn } }       catch { }
            if (-not $when) { try { if ($item.ReceivedTime) { $when = $item.ReceivedTime } } catch { } }
            if (-not $when -or $when -lt $Since) { continue }

            # --- sender ---
            $fromName = ''
            $fromAddr = ''
            try { $fromName = [string]$item.SenderName } catch { }
            try { $fromAddr = [string]$item.SenderEmailAddress } catch { }
            if ($fromAddr -like '/o=*') {
                try { $fromAddr = Resolve-Smtp -AddressEntry $item.Sender -Fallback $fromAddr } catch { }
            }

            # --- recipients ---
            $to = @()
            $cc = @()
            $recipParties = @()
            try {
                foreach ($r in $item.Recipients) {
                    $rn = [string]$r.Name
                    $ra = ''
                    try { $ra = [string]$r.Address } catch { }
                    if ($ra -like '/o=*') { $ra = Resolve-Smtp -AddressEntry $r.AddressEntry -Fallback $ra }
                    # if it is still an unresolved Exchange DN, do not clutter the output with it
                    $line = if ($ra -and $ra -ne $rn -and $ra -notlike '/o=*') { "$rn <$ra>" } else { $rn }
                    if ($r.Type -eq 2) { $cc += $line } else { $to += $line }
                    $recipParties += $rn
                    $recipParties += $ra
                }
            } catch {
                try { $to += [string]$item.To } catch { }
                try { $cc += [string]$item.CC } catch { }
                $recipParties += $to
                $recipParties += $cc
            }

            # --- where does this person appear? ---
            $senderHay = ($fromName + ' ' + $fromAddr).ToLowerInvariant()
            $recipHay  = ($recipParties -join ' ').ToLowerInvariant()
            $isFromThem = $false
            $isToThem   = $false
            foreach ($t in $tokens) {
                if ($senderHay.Contains($t)) { $isFromThem = $true }
                if ($recipHay.Contains($t))  { $isToThem   = $true }
            }
            if (-not ($isFromThem -or $isToThem)) { continue }

            $iAmSender = $myAddresses.Contains($fromAddr.ToLowerInvariant())
            if ($DirectOnly -and -not ($isFromThem -or ($iAmSender -and $isToThem))) { continue }

            # --- body ---
            $body = ''
            try { $body = [string]$item.Body } catch { }
            if (-not $body) { try { $body = [string]$item.HTMLBody } catch { } }
            $body = ($body -replace "`r`n", "`n").TrimEnd()

            $subject = ''
            try { $subject = [string]$item.Subject } catch { }

            # --- dedupe (same message filed in two folders) ---
            $key = '{0}|{1}|{2}|{3}' -f $when.ToString('o'), $fromAddr, $subject, $body.Length
            if (-not $seen.Add($key)) { continue }

            $atts = @()
            try { foreach ($a in $item.Attachments) { $atts += $a.FileName } } catch { }

            $dir = if ($iAmSender) { 'SENT by you' } else { 'RECEIVED' }

            [void]$hits.Add([pscustomobject]@{
                When        = $when
                Direction   = $dir
                From        = $(if ($fromAddr -and $fromAddr -ne $fromName) { "$fromName <$fromAddr>" } else { $fromName })
                To          = ($to -join '; ')
                Cc          = ($cc -join '; ')
                Subject     = $subject
                Attachments = ($atts -join '; ')
                Body        = $body
                Folder      = $folder.FolderPath
            })
        } catch { continue }
    }
}
Write-Progress -Activity 'Scanning Outlook' -Completed

$hits = @($hits | Sort-Object When)
Write-Host "Messages : $($hits.Count) found"

if ($hits.Count -eq 0) {
    Write-Warning "No messages matched. Try a shorter token (e.g. just the surname), a larger -Days, or add -AllStores / -IncludeDeleted."
    exit 0
}

# -------------------------------------------------------------- output ----
if (-not $OutFile) {
    $slug = ($Person[0] -replace '[^A-Za-z0-9]+', '-').Trim('-')
    $dir  = Join-Path $PSScriptRoot 'context'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $OutFile = Join-Path $dir ('{0}_{1}.md' -f $slug, (Get-Date -Format 'yyyy-MM-dd'))
}

$sb = New-Object System.Text.StringBuilder
function W { param([string]$s = '') [void]$sb.AppendLine($s) }

$sentCount = @($hits | Where-Object { $_.Direction -like 'SENT*' }).Count
$rcvdCount = @($hits | Where-Object { $_.Direction -eq 'RECEIVED' }).Count

W "# Email context: $($Person -join ' / ')"
W
W "- Window: $($Since.ToString('yyyy-MM-dd')) through $((Get-Date).ToString('yyyy-MM-dd'))"
W "- Messages: $($hits.Count)  (sent by you: $sentCount, received: $rcvdCount)"
W "- First: $($hits[0].When.ToString('yyyy-MM-dd HH:mm'))"
W "- Last:  $($hits[-1].When.ToString('yyyy-MM-dd HH:mm'))"
W "- Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
W
W '## Index'
W
W '| # | Date | Direction | Subject |'
W '|---|------|-----------|---------|'
$i = 0
foreach ($h in $hits) {
    $i++
    $s = ($h.Subject -replace '\|', '/')
    W "| $i | $($h.When.ToString('yyyy-MM-dd HH:mm')) | $($h.Direction) | $s |"
}
W
W '---'
W
W '## Full messages (chronological, oldest first)'
W

$i = 0
foreach ($h in $hits) {
    $i++
    $body = $h.Body
    if ($TrimQuoted) {
        $cut = [regex]::Match($body, '(?m)^\s*(-{3,}\s*Original Message|From:\s|On .{5,80} wrote:|_{10,})')
        if ($cut.Success -and $cut.Index -gt 40) {
            $body = $body.Substring(0, $cut.Index).TrimEnd() + "`n[quoted history trimmed]"
        }
    }
    W "### [$i] $($h.When.ToString('dddd, MMMM d, yyyy - h:mm tt')) - $($h.Direction)"
    W
    W "**Subject:** $($h.Subject)"
    W "**From:** $($h.From)"
    W "**To:** $($h.To)"
    if ($h.Cc)          { W "**Cc:** $($h.Cc)" }
    if ($h.Attachments) { W "**Attachments:** $($h.Attachments)" }
    W "**Folder:** $($h.Folder)"
    W
    W '```text'
    W $body
    W '```'
    W
    W '---'
    W
}

[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
$kb = (Get-Item $OutFile).Length / 1KB
Write-Host ''
Write-Host "Wrote $OutFile"
Write-Host ('Size: {0:N0} KB (roughly {1:N0}k tokens)' -f $kb, ((Get-Item $OutFile).Length / 4000))

if ($PassThru) {
    [pscustomobject]@{ OutFile = $OutFile; Since = $Since; Messages = $hits }
}
