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

.PARAMETER NoIndex
    Scan Outlook directly instead of using the search index (FadoeIndex.ps1). The index is
    used automatically when it's available and -Days is within the last year; it picks the
    matching messages in a second or two and only opens those.

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
    [switch]   $PassThru,
    [switch]   $NoIndex
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

$PR_SMTP        = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$PR_SENDER_SMTP = 'http://schemas.microsoft.com/mapi/proptag/0x5D01001F'
$PR_TARGET      = 'http://schemas.microsoft.com/mapi/proptag/0x8011001F'
$smtpCache = @{}

# Turn an Exchange /o=... address into a real SMTP address. Covers people in your own org and
# outside people who were added to the org directory as contacts (Outlook stores those as /o=
# too, and they have no SMTP property on the directory entry itself).
function Resolve-Smtp {
    param($AddressEntry, [string]$Fallback, $Recipient)
    $key = $Fallback
    if ($key -and $smtpCache.ContainsKey($key)) { return $smtpCache[$key] }
    $result = $Fallback
    # 1. the recipient row on the message usually carries the SMTP address -- no lookup needed
    if ($Recipient) {
        try { $t = [string]$Recipient.PropertyAccessor.GetProperty($PR_SMTP); if ($t) { $result = $t } } catch { }
    }
    # 2. directory users
    if ($result -like '/o=*' -and $AddressEntry) {
        try { $t = [string]$AddressEntry.PropertyAccessor.GetProperty($PR_SMTP); if ($t) { $result = $t } } catch { }
    }
    # 3. directory contacts keep the outside address as their target address ("SMTP:x@y")
    if ($result -like '/o=*' -and $AddressEntry) {
        try { $t = [string]$AddressEntry.PropertyAccessor.GetProperty($PR_TARGET); if ($t) { $result = $t -replace '^(?i)smtp:', '' } } catch { }
    }
    if ($result -like '/o=*' -and $AddressEntry) {
        try { $eu = $AddressEntry.GetExchangeUser(); if ($eu -and $eu.PrimarySmtpAddress) { $result = $eu.PrimarySmtpAddress } } catch { }
    }
    if ($key) { $smtpCache[$key] = $result }
    return $result
}

# Edit distance, for "did you mean" suggestions when a name or address has a typo.
function Get-EditDistance([string]$a, [string]$b) {
    $n = $a.Length; $m = $b.Length
    if ($n -eq 0) { return $m }; if ($m -eq 0) { return $n }
    $prev = New-Object int[] ($m + 1); $cur = New-Object int[] ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($a[$i - 1] -eq $b[$j - 1]) { 0 } else { 1 }
            $cur[$j] = [Math]::Min([Math]::Min($prev[$j] + 1, $cur[$j - 1] + 1), $prev[$j - 1] + $cost)
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    return $prev[$m]
}

# Everyone seen in the scanned mail, for suggestions: address -> name + message count
$people = @{}
function Add-Person([string]$addr, [string]$name) {
    if (-not $addr -or $addr -notlike '*@*') { return }
    $k = $addr.ToLowerInvariant()
    if ($people.ContainsKey($k)) { $people[$k].Count++ }
    else { $people[$k] = @{ Address = $addr; Name = $name; Count = 1 } }
    if ($name -and $name -ne $addr -and $people[$k].Name -eq $addr) { $people[$k].Name = $name }
}

$hits = New-Object System.Collections.ArrayList
$seen = New-Object System.Collections.Generic.HashSet[string]

# ---------------------------------------------------- fast path: the index ----
# The search index already knows every message's sender and recipients, so pick the matching
# messages from it and open only those. Falls back to scanning Outlook below.
$useIndex = $false
$lib = Join-Path $PSScriptRoot 'FadoeIndex.ps1'
if (-not $NoIndex -and (Test-Path $lib) -and $Since -ge (Get-Date).Date.AddDays(-365)) {
    try {
        . $lib
        Initialize-FadoeIndex
        [void](Sync-FadoeIndex)
        $useIndex = $true
    } catch { Write-Warning "Search index unavailable, scanning Outlook directly instead. ($($_.Exception.Message))" }
}

if ($useIndex) {
    $x = $global:FadoeIdx
    $pairRx = '([^;<>]+?)\s*<([^<>\s;]+@[^<>\s;]+)>'
    $matched = New-Object System.Collections.ArrayList
    foreach ($r in $x.Recs.Values) {
        if ($r.W -lt $Since) { continue }
        $fo = $x.Folders[$r.FK]
        if (-not $AllStores -and -not $fo.IsPrimary) { continue }
        if (-not $IncludeDeleted -and ($fo.Path -split '\\')[-1] -eq 'Deleted Items') { continue }

        Add-Person $r.SA $r.SN
        foreach ($m in [regex]::Matches($r.To + ';' + $r.Cc, $pairRx)) { Add-Person $m.Groups[2].Value $m.Groups[1].Value.Trim() }

        $senderHay = ($r.SN + ' ' + $r.SA).ToLowerInvariant()
        $recipHay  = $r.People.ToLowerInvariant()   # everyone on it, Bcc included
        $isFromThem = $false; $isToThem = $false
        foreach ($t in $tokens) {
            if ($senderHay.Contains($t)) { $isFromThem = $true }
            if ($recipHay.Contains($t))  { $isToThem   = $true }
        }
        if (-not ($isFromThem -or $isToThem)) { continue }
        $iAmSender = $myAddresses.Contains(([string]$r.SA).ToLowerInvariant())
        if ($DirectOnly -and -not ($isFromThem -or ($iAmSender -and $isToThem))) { continue }
        [void]$matched.Add($r)
    }
    Write-Host "Folders  : search index ($('{0:N0}' -f $x.Recs.Count) messages); $($matched.Count) involve this person"

    $k = 0
    foreach ($r in @($matched | Sort-Object W)) {
        $k++
        Write-Progress -Activity 'Reading messages' -Status ('{0} of {1}' -f $k, $matched.Count) -PercentComplete (100 * $k / [Math]::Max($matched.Count, 1))
        $body = $null; $atts = @(); $item = $null; $ac = $null
        try {
            $item = $x.Ns.GetItemFromID($r.E, $r.S)
            try { $body = [string]$item.Body } catch { }
            if (-not $body) { try { $body = [string]$item.HTMLBody } catch { } }
            try { $ac = $item.Attachments; foreach ($a in $ac) { $atts += $a.FileName; Release-FadoeCom $a } } catch { }
        } catch { } finally { Release-FadoeCom $ac $item }
        if ($null -eq $body) { $body = $r.Own }   # moved since the last sync: fall back to what the index kept
        $body = ($body -replace "`r`n", "`n").TrimEnd()

        $key = '{0}|{1}|{2}|{3}' -f $r.W.ToString('o'), $r.SA, $r.Subj, $body.Length
        if (-not $seen.Add($key)) { continue }
        $iAmSender = $myAddresses.Contains(([string]$r.SA).ToLowerInvariant())
        [void]$hits.Add([pscustomobject]@{
            When        = $r.W
            Direction   = $(if ($iAmSender) { 'SENT by you' } else { 'RECEIVED' })
            From        = $(if ($r.SA -and $r.SA -ne $r.SN) { "$($r.SN) <$($r.SA)>" } else { $r.SN })
            To          = $r.To
            Cc          = $r.Cc
            Subject     = $r.Subj
            Attachments = ($atts -join '; ')
            Body        = $body
            Folder      = $x.Folders[$r.FK].Path
        })
        if ($k % 100 -eq 0) { Invoke-FadoeCleanup }
    }
    Write-Progress -Activity 'Reading messages' -Completed
}

# --------------------------------------------------------- folder walk ----
# (only when the index isn't being used)
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

if (-not $useIndex) {
    foreach ($store in $stores) {
        try { Add-Folders -Folder $store.GetRootFolder() }
        catch { Write-Warning "Skipped store '$($store.DisplayName)': $_" }
    }
    Write-Host "Folders  : $($folders.Count) mail folder(s) to scan"
}

# ------------------------------------------------------------ scanning ----
$stamp = $Since.ToString('yyyy-MM-dd HH:mm')
$dasl  = '@SQL=("urn:schemas:httpmail:datereceived" >= ''' + $stamp + ''' OR "urn:schemas:httpmail:date" >= ''' + $stamp + ''')'

$fi = 0
$opened = 0

foreach ($folder in $folders) {
    $fi++
    Write-Progress -Activity 'Scanning Outlook' -Status $folder.FolderPath -PercentComplete (100 * $fi / [Math]::Max($folders.Count, 1))

    try   { $items = $folder.Items.Restrict($dasl) }
    catch { $items = $folder.Items }

    foreach ($item in $items) {
        # Outlook caps how many of its objects can be open at once ("exhausted all shared
        # resources"); let go of the messages already read every so often.
        if ((++$opened % 200) -eq 0) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
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
                try { $t = [string]$item.PropertyAccessor.GetProperty($PR_SENDER_SMTP); if ($t) { $fromAddr = $t } } catch { }
            }
            if ($fromAddr -like '/o=*') {
                try { $fromAddr = Resolve-Smtp -AddressEntry $item.Sender -Fallback $fromAddr } catch { }
            }
            Add-Person $fromAddr $fromName

            # --- recipients ---
            $to = @()
            $cc = @()
            $recipParties = @()
            try {
                foreach ($r in $item.Recipients) {
                    $rn = [string]$r.Name
                    $ra = ''
                    try { $ra = [string]$r.Address } catch { }
                    if ($ra -like '/o=*') { $ra = Resolve-Smtp -AddressEntry $r.AddressEntry -Fallback $ra -Recipient $r }
                    Add-Person $ra $rn
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
$items = $null; $folders = $null
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()

$hits = @($hits | Sort-Object When)
Write-Host "Messages : $($hits.Count) found"

if ($hits.Count -eq 0) {
    # Did you mean...? People you actually emailed whose address or name is a near-miss.
    $suggestions = @()
    foreach ($p in $people.Values) {
        $addrL  = $p.Address.ToLowerInvariant()
        $local  = ($addrL -split '@')[0]
        $words  = @(([string]$p.Name).ToLowerInvariant() -split '[^a-z0-9]+' | Where-Object { $_.Length -ge 3 })
        $best = [int]::MaxValue
        foreach ($t in $tokens) {
            $allow = [Math]::Max(1, [Math]::Floor($t.Length / 4))
            $cands = if ($t -like '*@*') { @($addrL) } else { @($local) + $words }
            foreach ($c in $cands) {
                if ([Math]::Abs($c.Length - $t.Length) -gt $allow) { continue }
                $d = Get-EditDistance $t $c
                if ($d -le $allow -and $d -lt $best) { $best = $d }
            }
        }
        if ($best -lt [int]::MaxValue) {
            $suggestions += [pscustomobject]@{ Address = $p.Address; Name = $p.Name; Messages = $p.Count; Distance = $best }
        }
    }
    $suggestions = @($suggestions | Sort-Object Distance, @{ e = 'Messages'; Descending = $true } | Select-Object -First 4)

    Write-Warning "No messages matched. Try a shorter token (e.g. just the surname), a larger -Days, or add -AllStores / -IncludeDeleted."
    foreach ($s in $suggestions) { Write-Host ("Did you mean: {0} <{1}>  ({2} messages)" -f $s.Name, $s.Address, $s.Messages) -ForegroundColor Yellow }
    if ($PassThru) { [pscustomobject]@{ OutFile = $null; Since = $Since; Messages = @(); Suggestions = $suggestions } }
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
