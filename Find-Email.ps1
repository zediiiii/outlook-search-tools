<#
.SYNOPSIS
    Finds the exact message in a reply chain where your search words were actually written,
    and prints a search string that pulls that message up in new Outlook.

.DESCRIPTION
    Outlook's own search matches every reply in a thread, because each reply quotes the ones
    before it. This script ignores quoted reply history and matches only the text each sender
    wrote, then shows one result per conversation: the message where your words appear.

    For each result it prints the exact date and time, sender, subject, where the message sits
    in its thread, the matching passage, any links in it, and a search string you can paste
    into new Outlook's search box. The top result's search string is copied to the clipboard.

    Searches run against a local index (FadoeIndex.ps1). The first run builds it -- about a
    minute per 10,000 messages -- and later runs only pick up what changed, so searches take
    seconds. The index stays in %LOCALAPPDATA%\FADOE on this PC.

    Reads your mailbox through the classic Outlook desktop app (it must be installed, but you
    never have to open it). Read-only. Classic Outlook keeps about a year of mail on your PC,
    so older messages won't be found here -- use Outlook's own search for those.

.PARAMETER Search
    Words to find. Every word must appear. Put a phrase in double quotes to keep it together:
    'budget "site visit"'

.PARAMETER From
    Only messages whose sender name or address contains this text.

.PARAMETER Days
    How far back to look. Default 365 (also the most the index holds).

.PARAMETER Top
    How many conversations to show. Default 10.

.PARAMETER NoClipboard
    Don't copy the top result's search string to the clipboard.

.PARAMETER PassThru
    Also return the results as objects (when, sender, subject, the sender's own text, links,
    thread position, the Outlook search string, IDs) so another script can use them.
    Used by FADOE.ps1.

.PARAMETER SyncOnly
    Just bring the index up to date (or build it) and exit.

.PARAMETER Rebuild
    Re-check every folder, not just the ones whose message count changed.

.PARAMETER NoSync
    Search the index as it is, without checking Outlook for new mail first.

.EXAMPLE
    .\Find-Email.ps1 'kickoff video'

.EXAMPLE
    .\Find-Email.ps1 'budget "site visit"' -From jsmith -Days 90
#>

[CmdletBinding(DefaultParameterSetName = 'Search')]
param(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Search')]
    [string] $Search,

    [Parameter(ParameterSetName = 'Search')] [string] $From,
    [Parameter(ParameterSetName = 'Search')] [int]    $Days = 365,
    [Parameter(ParameterSetName = 'Search')] [int]    $Top = 10,
    [Parameter(ParameterSetName = 'Search')] [switch] $NoClipboard,
    [Parameter(ParameterSetName = 'Search')] [switch] $PassThru,
    [Parameter(ParameterSetName = 'Search')] [switch] $NoSync,
    [Parameter(Mandatory = $true, ParameterSetName = 'Sync')] [switch] $SyncOnly,
    [switch] $Rebuild
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'FadoeIndex.ps1')

# ------------------------------------------------------------ sync only ----
if ($SyncOnly) {
    Initialize-FadoeIndex
    $s = Sync-FadoeIndex -Full:$Rebuild
    Write-Host ('Index up to date: {0:N0} messages ({1:N0} new).' -f $s.Messages, $s.NewMessages)
    $s
    return
}

# ------------------------------------------------------------------ terms ----
# "quoted phrases" stay together; everything else splits on whitespace
$terms = @([regex]::Matches($Search, '"([^"]+)"|(\S+)') | ForEach-Object {
    if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value }
} | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
if ($terms.Count -eq 0) { Write-Error 'Nothing to search for.'; exit 1 }
$Since = (Get-Date).Date.AddDays(-$Days)

# ---------------------------------------------------------------- search ----
Initialize-FadoeIndex
if (-not $NoSync -or $global:FadoeIdx.Recs.Count -eq 0) {
    $s = Sync-FadoeIndex -Full:$Rebuild
    if ($s.NewMessages -gt 200) { Write-Host ('Indexed {0:N0} new messages.' -f $s.NewMessages) -ForegroundColor DarkGray }
}
$found   = Search-FadoeIndex -Terms $terms -From $From -Top $Top -Days $Days
$cands   = $found.Candidates
$results = $found.Results

if ($cands.Count -eq 0) {
    Write-Host "No messages found containing: $($terms -join ', ')  (since $($Since.ToString('yyyy-MM-dd')))"
    Write-Host 'Try fewer or different words. For mail older than a year, use Outlook''s own search.'
    exit 0
}

# Links that show up across several conversations are signatures/boilerplate, not content.
# (dictionary pass rather than Group-Object, which is very slow when thousands match)
$linkRx = '<(https?://[^>\s]+)>'
$convLinks = @{}
foreach ($c in $cands) {
    if ($c.Own.IndexOf('<http', [StringComparison]::Ordinal) -lt 0) { continue }
    $set = $convLinks[$c.Conv]
    if ($null -eq $set) { $set = @{}; $convLinks[$c.Conv] = $set }
    foreach ($m in [regex]::Matches($c.Own, $linkRx)) { $set[$m.Groups[1].Value] = 1 }
}
$linkFreq = @{}
foreach ($set in $convLinks.Values) { foreach ($u in $set.Keys) { $linkFreq[$u] = 1 + [int]$linkFreq[$u] } }
$boiler = '(?i)^(mailto:|tel:)|facebook\.com|twitter\.com|//x\.com|linkedin\.com|instagram\.com|youtube\.com/(user|channel)'

# ------------------------------------------------------------ helpers ----
function Get-OutlookQuery($r) {
    $q = @()
    if ($r.SenderAddr -and $r.SenderAddr -notlike '/o=*') { $q += 'from:' + $r.SenderAddr }
    elseif ($r.SenderName) { $q += 'from:"' + ($r.SenderName -replace '"', '') + '"' }
    $s = ((Get-NormSubject $r.Subject) -replace '"', '' -replace '\s+', ' ').Trim()
    if ($s) { $q += 'subject:"' + $s + '"' }
    $q += 'sent:' + $r.When.ToString('MM/dd/yyyy', [Globalization.CultureInfo]::InvariantCulture)
    $q -join ' AND '
}

function Write-Highlighted([string]$text, [string[]]$words) {
    $rx = ($words | ForEach-Object { [regex]::Escape($_) }) -join '|'
    $last = 0
    foreach ($m in [regex]::Matches($text, $rx, 'IgnoreCase')) {
        Write-Host $text.Substring($last, $m.Index - $last) -NoNewline
        Write-Host $m.Value -NoNewline -ForegroundColor Black -BackgroundColor Yellow
        $last = $m.Index + $m.Length
    }
    Write-Host $text.Substring($last)
}

function Get-Snippet([string]$flat, [int]$at, [int]$radius = 170) {
    if ($flat.Length -eq 0) { return '' }
    $at = [Math]::Min([Math]::Max($at, 0), $flat.Length - 1)
    $start = [Math]::Max(0, $at - [int]($radius / 2))
    $len = [Math]::Min($flat.Length - $start, 2 * $radius)
    $pre = if ($start -gt 0) { '...' } else { '' }
    $post = if ($start + $len -lt $flat.Length) { '...' } else { '' }
    $pre + $flat.Substring($start, $len) + $post
}

# ------------------------------------------------------------- output ----
$convCount = $found.Conversations
Write-Host ''
Write-Host ("{0} conversation(s) contain: {1}   (since {2}; showing {3})" -f `
    $convCount, (($terms | ForEach-Object { "'$_'" }) -join ' + '),
    $Since.ToString('MMM d, yyyy'), $results.Count) -ForegroundColor Cyan

$rows = New-Object System.Collections.ArrayList
$n = 0
foreach ($r in $results) {
    $n++
    Write-Host ''
    Write-Host ("[{0}] {1}" -f $n, $r.When.ToString('ddd MMM d, yyyy \a\t h:mm tt')) -ForegroundColor Green -NoNewline
    Write-Host ("   from {0}" -f $r.SenderName)
    Write-Host ("    Subject : {0}" -f $r.Subject)
    if ($r.To) {
        $toShort = if ($r.To.Length -gt 90) { $r.To.Substring(0, 87) + '...' } else { $r.To }
        Write-Host ("    To      : {0}" -f $toShort)
    }
    $where = $r.Folder
    if ($r.Pos) { $where += ("   |   message {0} of {1} in the thread, oldest first" -f $r.Pos, $r.Total) }
    Write-Host ("    Where   : {0}" -f $where)
    if (-not $r.IsPrimary) {
        Write-Host ("              In '{0}' -- click a folder in that mailbox before searching." -f $r.StoreName) -ForegroundColor Yellow
    }
    if ($r.OwnHits -eq 0) {
        Write-Host '    (your words are only in the subject line of this thread)' -ForegroundColor DarkGray
    }
    if ($r.IsBulk) {
        Write-Host '    (newsletter / mass mailing)' -ForegroundColor DarkGray
    }
    Write-Host '    Passage : ' -NoNewline
    Write-Highlighted (Get-Snippet $r.Flat $r.Start) $terms

    $links = @([regex]::Matches($r.Own, $linkRx) | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique |
        Where-Object { $_ -notmatch $boiler -and [int]$linkFreq[$_] -lt 3 } | Select-Object -First 6)
    foreach ($l in $links) { Write-Host ("    Link    : {0}" -f $l) }

    $q = Get-OutlookQuery $r
    Write-Host '    Search  : ' -NoNewline
    Write-Host $q -ForegroundColor Cyan
    if ($n -eq 1 -and -not $NoClipboard) {
        try { Set-Clipboard -Value $q; Write-Host '              (copied to clipboard)' -ForegroundColor DarkGray } catch { }
    }

    [void]$rows.Add([pscustomobject]@{
        Number             = $n
        When               = $r.When
        Sender             = $r.SenderName
        SenderAddr         = $r.SenderAddr
        To                 = $r.To
        Subject            = $r.Subject
        Folder             = $r.Folder
        Mailbox            = $r.StoreName
        IsPrimary          = $r.IsPrimary
        Position           = $r.Pos
        ThreadCount        = $r.Total
        IsBulk             = $r.IsBulk
        OwnHits            = $r.OwnHits
        Text               = $r.Own
        Links              = $links
        Query              = $q
        Terms              = $terms
        EntryID            = $r.EntryID
        StoreID            = $r.StoreID
        TotalConversations = $convCount
    })
}
Write-Host ''
Write-Host 'Paste a Search line into new Outlook''s search box. If conversations are grouped, open the' -ForegroundColor DarkGray
Write-Host 'thread and look for the message sent at the time shown above.' -ForegroundColor DarkGray

if ($PassThru) { $rows }
