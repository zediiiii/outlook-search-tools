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

    Reads your mailbox through the classic Outlook desktop app (it must be installed, but you
    never have to open it). Read-only. Classic Outlook keeps about a year of mail on your PC,
    so older messages won't be found here -- use Outlook's own search for those.

.PARAMETER Search
    Words to find. Every word must appear. Put a phrase in double quotes to keep it together:
    'budget "site visit"'

.PARAMETER From
    Only messages whose sender name or address contains this text.

.PARAMETER Days
    How far back to look. Default 365.

.PARAMETER Top
    How many conversations to show. Default 10.

.PARAMETER NoClipboard
    Don't copy the top result's search string to the clipboard.

.PARAMETER PassThru
    Also return the results as objects (when, sender, subject, the sender's own text, links,
    thread position, the Outlook search string, IDs) so another script can use them.
    Used by FADOE.ps1.

.EXAMPLE
    .\Find-Email.ps1 'kickoff video'

.EXAMPLE
    .\Find-Email.ps1 'budget "site visit"' -From jsmith -Days 90
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Search,

    [string] $From,
    [int]    $Days = 365,
    [int]    $Top = 10,
    [switch] $NoClipboard,
    [switch] $PassThru
)

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ terms ----
# "quoted phrases" stay together; everything else splits on whitespace
$terms = @([regex]::Matches($Search, '"([^"]+)"|(\S+)') | ForEach-Object {
    if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value }
} | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() })
if ($terms.Count -eq 0) { Write-Error 'Nothing to search for.'; exit 1 }
$Since = (Get-Date).Date.AddDays(-$Days)

try {
    $ns = (New-Object -ComObject Outlook.Application).GetNamespace('MAPI')
} catch {
    Write-Error "Could not connect to Outlook. The classic Outlook desktop app must be installed with your mail profile set up. ($_)"
    exit 1
}
$defaultStoreId = $ns.GetDefaultFolder(6).Store.StoreID   # 6 = olFolderInbox

$PR_SMTP = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
function Get-SenderSmtp($item) {
    $a = ''
    try { $a = [string]$item.SenderEmailAddress } catch { }
    if ($a -like '/o=*') {
        try { $a = [string]$item.Sender.PropertyAccessor.GetProperty($PR_SMTP) }
        catch { try { $a = [string]$item.Sender.GetExchangeUser().PrimarySmtpAddress } catch { } }
    }
    return $a
}

# Text the sender actually wrote: everything above the first quoted-reply header.
$quoteRx = '(?m)^[ \t>]*(From:[ \t]|-{2,}[ \t]*Original Message|_{10,}|On .{5,120} wrote:)'
function Get-OwnText([string]$body) {
    $m = [regex]::Match($body, $quoteRx)
    $own = if ($m.Success) { $body.Substring(0, $m.Index) } else { $body }
    (($own -split "`r?`n") | Where-Object { $_ -notmatch '^\s*>' }) -join "`n"
}

# Newsletters and mass mailings: mailing-list headers first, then an unsubscribe footer
# (footer check skipped on replies/forwards, which may just be quoting a newsletter).
$PR_HEADERS = 'http://schemas.microsoft.com/mapi/proptag/0x007D001F'
function Test-Bulk($item, [string]$subj, [string]$body) {
    try {
        $h = [string]$item.PropertyAccessor.GetProperty($PR_HEADERS)
        if ($h -match '(?im)^(List-Unsubscribe|List-Id):|^Precedence:\s*(bulk|list)') { return $true }
    } catch { }
    if ($subj -match '(?i)^\s*(re|fw|fwd)\s*:') { return $false }
    return ($body -match '(?i)unsubscribe|opt[- ]out|manage (your )?(email )?preferences|update your preferences')
}

function Get-NormSubject([string]$s) {
    ($s -replace '(?i)^\s*((re|fw|fwd)\s*:\s*)+', '').Trim()
}

# Whitespace collapsed, raw <https://...> link targets removed -- what a reader actually sees.
function Get-Flat([string]$own) {
    [regex]::Replace(($own -replace '\s+', ' '), '\s*<https?://[^>]+>', '').Trim()
}

# Smallest stretch of text containing every term (minimum-window search).
# Span is how many characters it covers; Start is where it begins.
function Get-Span([string]$low, [string[]]$words) {
    $ev = New-Object System.Collections.Generic.List[object]
    for ($k = 0; $k -lt $words.Count; $k++) {
        $i = $low.IndexOf($words[$k], [StringComparison]::Ordinal); $c = 0
        while ($i -ge 0 -and $c -lt 300) {
            $ev.Add([pscustomobject]@{ P = $i; K = $k; L = $words[$k].Length })
            $i = $low.IndexOf($words[$k], $i + 1, [StringComparison]::Ordinal); $c++
        }
    }
    if ($ev.Count -eq 0) { return @{ Span = [int]::MaxValue; Start = 0 } }
    $ev = @($ev | Sort-Object P)
    $cnt = @{}; $have = 0; $l = 0
    $best = [int]::MaxValue; $bestStart = $ev[0].P
    for ($r = 0; $r -lt $ev.Count; $r++) {
        $k = $ev[$r].K; $cnt[$k] = 1 + [int]$cnt[$k]; if ($cnt[$k] -eq 1) { $have++ }
        while ($have -eq $words.Count) {
            $span = $ev[$r].P + $ev[$r].L - $ev[$l].P
            if ($span -lt $best) { $best = $span; $bestStart = $ev[$l].P }
            $kl = $ev[$l].K; $cnt[$kl]--; if ($cnt[$kl] -eq 0) { $have-- }; $l++
        }
    }
    @{ Span = $best; Start = $bestStart }
}

# 0 = words form one phrase, 1 = same paragraph, 2 = same message, 3 = not all in the sender's text
function Get-Tier([int]$span) {
    if ($span -le 80) { 0 } elseif ($span -le 400) { 1 } elseif ($span -lt [int]::MaxValue) { 2 } else { 3 }
}

# ------------------------------------------------------------ folder walk ----
$skipNames = @('Junk Email', 'Junk E-mail', 'Sync Issues', 'RSS Feeds', 'RSS Subscriptions',
               'Conversation History', 'Outbox', 'Drafts')
$folders = New-Object System.Collections.ArrayList
function Add-Folders {
    param($Folder, [int]$Depth = 0)
    if ($Depth -gt 12 -or $skipNames -contains $Folder.Name) { return }
    try { if ($Folder.DefaultItemType -eq 0) { [void]$folders.Add($Folder) } } catch { }
    try { foreach ($sub in $Folder.Folders) { Add-Folders -Folder $sub -Depth ($Depth + 1) } } catch { }
}
foreach ($store in $ns.Stores) {
    try { Add-Folders -Folder $store.GetRootFolder() } catch { }
}

# ---------------------------------------------------------------- search ----
$stamp = $Since.ToString('yyyy-MM-dd HH:mm')
$dateClause = '("urn:schemas:httpmail:datereceived" >= ''' + $stamp + ''' OR "urn:schemas:httpmail:date" >= ''' + $stamp + ''')'
$parts = @($dateClause)
foreach ($t in $terms) {
    $e = $t.Replace("'", "''")
    $parts += '("urn:schemas:httpmail:subject" LIKE ''%' + $e + '%'' OR "urn:schemas:httpmail:textdescription" LIKE ''%' + $e + '%'')'
}
$dasl     = '@SQL=' + ($parts -join ' AND ')
$daslDate = '@SQL=' + $dateClause

$cands = New-Object System.Collections.ArrayList
$fi = 0
foreach ($f in $folders) {
    $fi++
    Write-Progress -Activity "Searching for: $Search" -Status $f.FolderPath -PercentComplete (100 * $fi / [Math]::Max($folders.Count, 1))
    $items = $null
    try { $items = $f.Items.Restrict($dasl) }
    catch { try { $items = $f.Items.Restrict($daslDate) } catch { continue } }

    foreach ($it in $items) {
        try {
            if ($it.Class -ne 43) { continue }   # 43 = olMail
            $when = $null
            try { if ($it.SentOn -and $it.SentOn.Year -lt 4000) { $when = $it.SentOn } } catch { }
            if (-not $when) { try { $when = $it.ReceivedTime } catch { } }
            if (-not $when -or $when -lt $Since) { continue }

            $subj = [string]$it.Subject
            $body = [string]$it.Body
            $own  = Get-OwnText $body
            $flat = Get-Flat $own
            $subjL = $subj.ToLowerInvariant()
            $ownL  = $flat.ToLowerInvariant()

            # every term must appear in the subject or in what this sender wrote
            $ok = $true; $ownHits = 0
            foreach ($t in $terms) {
                $inOwn = $ownL.Contains($t)
                if ($inOwn) { $ownHits++ }
                if (-not ($inOwn -or $subjL.Contains($t))) { $ok = $false; break }
            }
            if (-not $ok) { continue }

            $sName = [string]$it.SenderName
            $sAddr = Get-SenderSmtp $it
            if ($From -and ("$sName $sAddr" -notlike "*$From*")) { continue }

            $conv = ''
            try { $conv = [string]$it.ConversationID } catch { }
            if (-not $conv) { $conv = 'subj:' + (Get-NormSubject $subj).ToLowerInvariant() }

            $to = ''
            try { $to = [string]$it.To } catch { }

            $sp = if ($ownHits -eq $terms.Count) { Get-Span $ownL $terms } else { @{ Span = [int]::MaxValue; Start = 0 } }
            if ($ownHits -gt 0 -and $ownHits -lt $terms.Count) {
                # start the passage at the first term the sender did write
                foreach ($t in $terms) { $i = $ownL.IndexOf($t, [StringComparison]::Ordinal); if ($i -ge 0) { $sp.Start = $i; break } }
            }
            # every word in the subject line is the strongest signal; people remember subjects
            $subjAll = $true
            foreach ($t in $terms) { if (-not $subjL.Contains($t)) { $subjAll = $false; break } }
            $tier = if ($subjAll) { -1 } else { Get-Tier $sp.Span }
            # newsletters and mass mailings sink below real correspondence
            $isBulk = Test-Bulk $it $subj $body

            [void]$cands.Add([pscustomobject]@{
                When       = $when
                Subject    = $subj
                SenderName = $sName
                SenderAddr = $sAddr
                To         = $to
                OwnHits    = $ownHits
                Tier       = $tier
                IsBulk     = [bool]$isBulk
                Start      = $sp.Start
                Own        = $own
                Flat       = $flat
                Folder     = $f.FolderPath
                StoreName  = $f.Store.DisplayName
                IsPrimary  = ($f.Store.StoreID -eq $defaultStoreId)
                Conv       = $conv
                EntryID    = $it.EntryID
                StoreID    = $f.StoreID
            })
        } catch { continue }
    }
}
Write-Progress -Activity "Searching for: $Search" -Completed

if ($cands.Count -eq 0) {
    Write-Host "No messages found containing: $($terms -join ', ')  (since $($Since.ToString('yyyy-MM-dd')))"
    Write-Host 'Try fewer or different words. For mail older than a year, use Outlook''s own search.'
    exit 0
}

# One result per conversation: the message where the words were actually written.
# Within a thread: most terms in the sender's own text, then tightest phrase, then earliest.
# Across threads: same, with bulk mail after real correspondence and newest first as the final tiebreak.
$results = foreach ($g in ($cands | Group-Object Conv)) {
    $g.Group | Sort-Object @{ e = 'OwnHits'; Descending = $true }, @{ e = 'Tier'; Descending = $false },
                           @{ e = 'When'; Descending = $false } | Select-Object -First 1
}
$results = @($results | Sort-Object @{ e = 'IsBulk'; Descending = $false }, @{ e = 'OwnHits'; Descending = $true },
                                    @{ e = 'Tier'; Descending = $false }, @{ e = 'When'; Descending = $true } |
             Select-Object -First $Top)

# Links that show up across several conversations are signatures/boilerplate, not content.
$linkRx = '<(https?://[^>\s]+)>'
$linkFreq = @{}
foreach ($g in ($cands | Group-Object Conv)) {
    $seenHere = @{}
    foreach ($c in $g.Group) {
        foreach ($m in [regex]::Matches($c.Own, $linkRx)) { $seenHere[$m.Groups[1].Value] = 1 }
    }
    foreach ($u in $seenHere.Keys) { $linkFreq[$u] = 1 + [int]$linkFreq[$u] }
}
$boiler = '(?i)^(mailto:|tel:)|facebook\.com|twitter\.com|//x\.com|linkedin\.com|instagram\.com|youtube\.com/(user|channel)'

# ------------------------------------------------------------ helpers ----
function Get-ThreadPosition($r) {
    try {
        $item = $ns.GetItemFromID($r.EntryID, $r.StoreID)
        $conv = $item.GetConversation()
        if (-not $conv) { return $null }
        $tbl = $conv.GetTable()
        try { [void]$tbl.Columns.Add('SentOn') } catch { }
        $times = New-Object System.Collections.Generic.HashSet[string]
        while (-not $tbl.EndOfTable) {
            $row = $tbl.GetNextRow()
            try {
                $t = [datetime]$row.Item('SentOn')
                if ($t.Year -lt 4000) { [void]$times.Add($t.ToString('yyyyMMddHHmm')) }
            } catch { }
        }
        if ($times.Count -lt 2) { return $null }
        $mine = $r.When.ToString('yyyyMMddHHmm')
        $pos = @($times | Where-Object { $_ -le $mine }).Count
        return @{ Pos = [Math]::Max($pos, 1); Total = $times.Count }
    } catch { return $null }
}

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
Write-Host ''
Write-Host ("{0} conversation(s) contain: {1}   (since {2}; showing {3})" -f `
    @($cands | Group-Object Conv).Count, (($terms | ForEach-Object { "'$_'" }) -join ' + '),
    $Since.ToString('MMM d, yyyy'), $results.Count) -ForegroundColor Cyan

$convCount = @($cands | Group-Object Conv).Count
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
    $pos = Get-ThreadPosition $r
    if ($pos) { $where += ("   |   message {0} of {1} in the thread, oldest first" -f $pos.Pos, $pos.Total) }
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
        Position           = $(if ($pos) { $pos.Pos } else { $null })
        ThreadCount        = $(if ($pos) { $pos.Total } else { $null })
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
