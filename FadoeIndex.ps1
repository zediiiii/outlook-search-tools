<#
.SYNOPSIS
    FADOE's local search index. Dot-source this file; it is not meant to be run on its own.

.DESCRIPTION
    Searching by reading every message through Outlook takes ~30 seconds. Instead, each message
    from the last year is read ONCE and a small record is kept: sender, recipients, subject, and
    the text the sender actually wrote (quoted reply history dropped). Searches then run in
    memory.

    Keeping it current is cheap: a folder is only re-listed when its message count changes, the
    re-listing uses Outlook's fast table view (IDs only, no message bodies), and only messages
    that are actually new get opened.

    State lives in $global:FadoeIdx, so a long-running host (the FADOE window) keeps the loaded
    index and Outlook connection between searches. A one-shot script pays the ~1 s load.

    PRIVACY: the index holds copies of your email text, like the catch-up files do. It is stored
    in %LOCALAPPDATA%\FADOE\ inside your Windows profile and never leaves this PC. Delete that
    folder at any time; it rebuilds itself on the next search.
#>

$FadoeIndexFormat = 1

$PR_SMTP          = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
$PR_SENDER_SMTP   = 'http://schemas.microsoft.com/mapi/proptag/0x5D01001F'
$PR_TARGET        = 'http://schemas.microsoft.com/mapi/proptag/0x8011001F'
$PR_HEADERS       = 'http://schemas.microsoft.com/mapi/proptag/0x007D001F'
$PR_CONTENT_COUNT = 'http://schemas.microsoft.com/mapi/proptag/0x36020003'

# Folders that never hold mail worth searching (plus shared-mailbox clutter that is slow online).
$FadoeSkipFolders = @('Junk Email', 'Junk E-mail', 'Sync Issues', 'RSS Feeds', 'RSS Subscriptions',
                      'Conversation History', 'Outbox', 'Drafts', 'Yammer Root', 'Files',
                      'ExternalContacts', 'PersonMetadata', 'Social Activity Notifications',
                      'Conversation Action Settings')

# ------------------------------------------------------------------ text ----
# Text the sender actually wrote: everything above the first quoted-reply header.
$FadoeQuoteRx = '(?m)^[ \t>]*(From:[ \t]|-{2,}[ \t]*Original Message|_{10,}|On .{5,120} wrote:)'
function Get-OwnText([string]$body) {
    $m = [regex]::Match($body, $FadoeQuoteRx)
    $own = if ($m.Success) { $body.Substring(0, $m.Index) } else { $body }
    (($own -split "`r?`n") | Where-Object { $_ -notmatch '^\s*>' }) -join "`n"
}

# Whitespace collapsed, raw <https://...> link targets removed -- what a reader actually sees.
function Get-Flat([string]$own) {
    [regex]::Replace(($own -replace '\s+', ' '), '\s*<https?://[^>]+>', '').Trim()
}

function Get-NormSubject([string]$s) {
    ($s -replace '(?i)^\s*((re|fw|fwd)\s*:\s*)+', '').Trim()
}

# Smallest stretch of text containing every term (minimum-window search).
# Occurrences are packed as position*64 + term number into longs and sorted natively -- this
# runs once per matching message, so it has to be cheap even when thousands match.
function Get-Span([string]$low, [string[]]$words) {
    $n = $words.Count
    if ($n -eq 1) {   # one term: it is its own phrase
        $i = $low.IndexOf($words[0], [StringComparison]::Ordinal)
        if ($i -lt 0) { return @{ Span = [int]::MaxValue; Start = 0 } }
        return @{ Span = $words[0].Length; Start = $i }
    }
    $ev = New-Object 'System.Collections.Generic.List[long]'
    for ($k = 0; $k -lt $n -and $k -lt 64; $k++) {
        $i = $low.IndexOf($words[$k], [StringComparison]::Ordinal); $c = 0
        while ($i -ge 0 -and $c -lt 300) {
            $ev.Add([long]$i * 64 + $k)
            $i = $low.IndexOf($words[$k], $i + 1, [StringComparison]::Ordinal); $c++
        }
    }
    if ($ev.Count -eq 0) { return @{ Span = [int]::MaxValue; Start = 0 } }
    $ev.Sort()
    $cnt = New-Object int[] 64; $have = 0; $l = 0
    $best = [int]::MaxValue; $bestStart = [int]($ev[0] -shr 6)
    for ($r = 0; $r -lt $ev.Count; $r++) {
        $k = [int]($ev[$r] -band 63); $cnt[$k]++; if ($cnt[$k] -eq 1) { $have++ }
        while ($have -eq $n) {
            $lp = [int]($ev[$l] -shr 6)
            $span = [int]($ev[$r] -shr 6) + $words[$k].Length - $lp
            if ($span -lt $best) { $best = $span; $bestStart = $lp }
            $kl = [int]($ev[$l] -band 63); $cnt[$kl]--; if ($cnt[$kl] -eq 0) { $have-- }; $l++
        }
    }
    @{ Span = $best; Start = $bestStart }
}

# 0 = words form one phrase, 1 = same paragraph, 2 = same message, 3 = not all in the sender's text
function Get-Tier([int]$span) {
    if ($span -le 80) { 0 } elseif ($span -le 400) { 1 } elseif ($span -lt [int]::MaxValue) { 2 } else { 3 }
}

# ----------------------------------------------------------- COM hygiene ----
# Classic Outlook allows only so many of its objects open at once (folders, messages, property
# readers -- especially in an online-only shared mailbox). Past that it shows "Outlook has
# exhausted all shared resources". So every object opened here is released as soon as it has
# been read, and a full cleanup runs every so often as a backstop.
function Release-FadoeCom {
    foreach ($o in $args) {
        if ($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
        }
    }
}
function Invoke-FadoeCleanup { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect() }

# Read one MAPI property and release the property reader straight away. Throws if missing.
function Get-FadoeProp($obj, [string]$tag) {
    $pa = $null
    try { $pa = $obj.PropertyAccessor; return $pa.GetProperty($tag) }
    finally { Release-FadoeCom $pa }
}

# ------------------------------------------------------------- addresses ----
function Get-FadoeSenderSmtp($item) {
    $a = ''
    try { $a = [string]$item.SenderEmailAddress } catch { }
    if ($a -like '/o=*') {   # the message itself usually carries the SMTP address
        try { $t = [string](Get-FadoeProp $item $PR_SENDER_SMTP); if ($t) { $a = $t } } catch { }
    }
    if ($a -like '/o=*') {
        $snd = $null
        try {
            $snd = $item.Sender
            try { $t = [string](Get-FadoeProp $snd $PR_SMTP); if ($t) { $a = $t } } catch { }
            if ($a -like '/o=*') {   # outside people added to the org directory as contacts
                try { $t = [string](Get-FadoeProp $snd $PR_TARGET); if ($t) { $a = $t -replace '^(?i)smtp:', '' } } catch { }
            }
            if ($a -like '/o=*') {
                $eu = $null
                try { $eu = $snd.GetExchangeUser(); $t = [string]$eu.PrimarySmtpAddress; if ($t) { $a = $t } } catch { } finally { Release-FadoeCom $eu }
            }
        } catch { } finally { Release-FadoeCom $snd }
    }
    return $a
}

function Resolve-FadoeRecipient($r, [string]$fallback) {
    $cache = $global:FadoeIdx.SmtpCache
    if ($fallback -and $cache.ContainsKey($fallback)) { return $cache[$fallback] }
    $result = $fallback
    try { $t = [string](Get-FadoeProp $r $PR_SMTP); if ($t) { $result = $t } } catch { }
    if ($result -like '/o=*') {
        $ae = $null
        try {
            $ae = $r.AddressEntry
            if ($ae) {
                try { $t = [string](Get-FadoeProp $ae $PR_SMTP); if ($t) { $result = $t } } catch { }
                if ($result -like '/o=*') { try { $t = [string](Get-FadoeProp $ae $PR_TARGET); if ($t) { $result = $t -replace '^(?i)smtp:', '' } } catch { } }
                if ($result -like '/o=*') {
                    $eu = $null
                    try { $eu = $ae.GetExchangeUser(); $t = [string]$eu.PrimarySmtpAddress; if ($t) { $result = $t } } catch { } finally { Release-FadoeCom $eu }
                }
            }
        } catch { } finally { Release-FadoeCom $ae }
    }
    if ($fallback) { $cache[$fallback] = $result }
    return $result
}

# Newsletters and mass mailings: mailing-list headers first, then an unsubscribe footer
# (footer check skipped on replies/forwards, which may just be quoting a newsletter).
function Test-FadoeBulk($item, [string]$subj, [string]$body) {
    try {
        $h = [string](Get-FadoeProp $item $PR_HEADERS)
        if ($h -match '(?im)^(List-Unsubscribe|List-Id):|^Precedence:\s*(bulk|list)') { return $true }
    } catch { }
    if ($subj -match '(?i)^\s*(re|fw|fwd)\s*:') { return $false }
    return [bool]($body -match '(?i)unsubscribe|opt[- ]out|manage (your )?(email )?preferences|update your preferences')
}

# ----------------------------------------------------------------- state ----
function Initialize-FadoeIndex {
    param([int]$Days = 365)
    if (-not $global:FadoeIdx) { $global:FadoeIdx = @{} }
    $x = $global:FadoeIdx
    if (-not $x.Ns) {
        try { $x.Ns = (New-Object -ComObject Outlook.Application).GetNamespace('MAPI') }
        catch { throw "Could not connect to Outlook. The classic Outlook desktop app must be installed with your mail profile set up. ($_)" }
        $x.DefaultStoreId = $x.Ns.GetDefaultFolder(6).Store.StoreID   # 6 = olFolderInbox
    }
    if (-not $x.Loaded) {
        $sha = [System.Security.Cryptography.SHA1]::Create()
        $hash = -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes([string]$x.DefaultStoreId)) |
                       Select-Object -First 6 | ForEach-Object { $_.ToString('x2') })
        $x.Dir          = Join-Path $env:LOCALAPPDATA "FADOE\index-$hash"
        $x.Days         = $Days
        $x.Folders      = @{}    # folder EntryID -> @{ S; Path; StoreName; IsPrimary; Count; Synced; Obj }
        $x.Recs         = @{}    # message EntryID -> record
        $x.ByFolder     = @{}    # folder EntryID -> HashSet of message EntryIDs
        $x.SmtpCache    = @{}
        $x.FolderListAt = [datetime]::MinValue
        $x.Dirty        = $false
        Read-FadoeIndex
        $x.Loaded = $true
    }
}

function Read-FadoeIndex {
    $x = $global:FadoeIdx
    $meta = Join-Path $x.Dir 'meta.txt'
    if (-not (Test-Path $meta)) { return }
    $m = @{}
    foreach ($line in [IO.File]::ReadAllLines($meta)) { $kv = $line.Split('=', 2); if ($kv.Count -eq 2) { $m[$kv[0]] = $kv[1] } }
    if ($m['format'] -ne [string]$FadoeIndexFormat) { return }   # older layout: rebuild from scratch
    $x.FolderListAt = [datetime]::new([int64]$m['folderListAt'])
    $enc = [Text.Encoding]::UTF8

    $ff = Join-Path $x.Dir 'folders.tsv'
    if (Test-Path $ff) {
        foreach ($line in [IO.File]::ReadAllLines($ff, $enc)) {
            $p = $line.Split("`t")
            if ($p.Count -lt 7) { continue }
            $x.Folders[$p[0]] = @{ S = $p[1]; Path = $enc.GetString([Convert]::FromBase64String($p[2]))
                                   StoreName = $enc.GetString([Convert]::FromBase64String($p[3]))
                                   IsPrimary = ($p[4] -eq '1'); Count = [int]$p[5]; Synced = ($p[6] -eq '1'); Obj = $null }
            $x.ByFolder[$p[0]] = New-Object 'System.Collections.Generic.HashSet[string]'
        }
    }
    $mf = Join-Path $x.Dir 'messages.tsv'
    if (Test-Path $mf) {
        foreach ($line in [IO.File]::ReadAllLines($mf, $enc)) {
            $p = $line.Split("`t")
            if ($p.Count -lt 14 -or -not $x.Folders.ContainsKey($p[2])) { continue }
            $subj = $enc.GetString([Convert]::FromBase64String($p[5]))
            $flat = $enc.GetString([Convert]::FromBase64String($p[13]))
            $x.Recs[$p[0]] = [pscustomobject]@{
                E = $p[0]; S = $p[1]; FK = $p[2]; W = [datetime]::new([int64]$p[3]); Bulk = ($p[4] -eq '1')
                Subj = $subj; SN = $enc.GetString([Convert]::FromBase64String($p[6])); SA = $p[7]
                To = $enc.GetString([Convert]::FromBase64String($p[8])); Cc = $enc.GetString([Convert]::FromBase64String($p[9]))
                People = $enc.GetString([Convert]::FromBase64String($p[10])); Conv = $p[11]
                Own = $enc.GetString([Convert]::FromBase64String($p[12])); Flat = $flat
                LSubj = $subj.ToLowerInvariant(); LFlat = $flat.ToLowerInvariant()
            }
            [void]$x.ByFolder[$p[2]].Add($p[0])
        }
    }
}

function Save-FadoeIndex {
    $x = $global:FadoeIdx
    New-Item -ItemType Directory -Force -Path $x.Dir | Out-Null
    $u = New-Object Text.UTF8Encoding($false)
    # base64 keeps tabs/newlines in text fields from breaking the one-record-per-line layout
    # (inline .NET calls, not a helper function: ~70k conversions per save)

    $tmp = Join-Path $x.Dir 'folders.tsv.tmp'
    $w = New-Object IO.StreamWriter($tmp, $false, $u)
    try {
        foreach ($k in $x.Folders.Keys) {
            $f = $x.Folders[$k]
            $w.WriteLine((@($k, $f.S, [Convert]::ToBase64String($u.GetBytes([string]$f.Path)),
                            [Convert]::ToBase64String($u.GetBytes([string]$f.StoreName)),
                            [int][bool]$f.IsPrimary, [int]$f.Count, [int][bool]$f.Synced) -join "`t"))
        }
    } finally { $w.Close() }
    Move-Item -LiteralPath $tmp -Destination (Join-Path $x.Dir 'folders.tsv') -Force

    $tmp = Join-Path $x.Dir 'messages.tsv.tmp'
    $w = New-Object IO.StreamWriter($tmp, $false, $u)
    try {
        foreach ($r in $x.Recs.Values) {
            $w.WriteLine((@($r.E, $r.S, $r.FK, $r.W.Ticks, [int][bool]$r.Bulk,
                            [Convert]::ToBase64String($u.GetBytes([string]$r.Subj)),
                            [Convert]::ToBase64String($u.GetBytes([string]$r.SN)), $r.SA,
                            [Convert]::ToBase64String($u.GetBytes([string]$r.To)),
                            [Convert]::ToBase64String($u.GetBytes([string]$r.Cc)),
                            [Convert]::ToBase64String($u.GetBytes([string]$r.People)), $r.Conv,
                            [Convert]::ToBase64String($u.GetBytes([string]$r.Own)),
                            [Convert]::ToBase64String($u.GetBytes([string]$r.Flat))) -join "`t"))
        }
    } finally { $w.Close() }
    Move-Item -LiteralPath $tmp -Destination (Join-Path $x.Dir 'messages.tsv') -Force

    [IO.File]::WriteAllLines((Join-Path $x.Dir 'meta.txt'),
        [string[]]@("format=$FadoeIndexFormat", "folderListAt=$($x.FolderListAt.Ticks)", "days=$($x.Days)", "saved=$((Get-Date).ToString('s'))"))
    $x.Dirty = $false
}

# ------------------------------------------------------------------ sync ----
function Update-FadoeFolderList {
    $x = $global:FadoeIdx
    $found = @{}
    $walk = {
        param($F, [int]$D)
        if ($D -gt 12 -or $FadoeSkipFolders -contains $F.Name) { return }
        try { if ($F.DefaultItemType -eq 0) { $found[$F.EntryID] = $F } } catch { }
        try { foreach ($sub in $F.Folders) { & $walk $sub ($D + 1) } } catch { }
    }
    foreach ($store in $x.Ns.Stores) { try { & $walk $store.GetRootFolder() 0 } catch { } }
    Invoke-FadoeCleanup   # the walk touches every folder in every mailbox; let go of the ones we don't keep

    foreach ($k in @($x.Folders.Keys)) {      # folders that no longer exist: drop them and their mail
        if (-not $found.ContainsKey($k)) {
            if ($x.ByFolder[$k]) { foreach ($e in $x.ByFolder[$k]) { [void]$x.Recs.Remove($e) } }
            $x.Folders.Remove($k); $x.ByFolder.Remove($k); $x.Dirty = $true
        }
    }
    foreach ($k in $found.Keys) {
        $f = $found[$k]
        if ($x.Folders.ContainsKey($k)) { $x.Folders[$k].Obj = $f; $x.Folders[$k].Path = $f.FolderPath; continue }
        $x.Folders[$k] = @{ S = $f.StoreID; Path = $f.FolderPath; StoreName = $f.Store.DisplayName
                            IsPrimary = ($f.StoreID -eq $x.DefaultStoreId); Count = -1; Synced = $false; Obj = $f }
        $x.ByFolder[$k] = New-Object 'System.Collections.Generic.HashSet[string]'
        $x.Dirty = $true
    }
    $x.FolderListAt = Get-Date
}

function New-FadoeRecord($it, [string]$fk, [string]$e, [string]$s) {
    if ($it.Class -ne 43) { return $null }   # 43 = olMail
    $when = $null
    try { if ($it.SentOn -and $it.SentOn.Year -lt 4000) { $when = $it.SentOn } } catch { }
    if (-not $when) { try { $when = $it.ReceivedTime } catch { } }
    if (-not $when) { return $null }

    $subj = ''; try { $subj = [string]$it.Subject } catch { }
    $body = ''; try { $body = [string]$it.Body } catch { }
    $own  = Get-OwnText $body
    $flat = Get-Flat $own
    $sn = ''; try { $sn = [string]$it.SenderName } catch { }
    $sa = Get-FadoeSenderSmtp $it

    $to = New-Object System.Collections.Generic.List[string]
    $cc = New-Object System.Collections.Generic.List[string]
    $people = New-Object System.Collections.Generic.List[string]
    $people.Add($sn); $people.Add($sa)
    $recips = $null
    try {
        $recips = $it.Recipients
        foreach ($r in $recips) {
            try {
                $rn = [string]$r.Name
                $ra = ''; try { $ra = [string]$r.Address } catch { }
                if ($ra -like '/o=*') { $ra = Resolve-FadoeRecipient $r $ra }
                $line = if ($ra -and $ra -ne $rn -and $ra -notlike '/o=*') { "$rn <$ra>" } else { $rn }
                if ($r.Type -eq 2) { $cc.Add($line) } else { $to.Add($line) }
                $people.Add($rn); $people.Add($ra)
            } finally { Release-FadoeCom $r }
        }
    } catch { try { $to.Add([string]$it.To) } catch { } }
    finally { Release-FadoeCom $recips }

    $conv = ''; try { $conv = [string]$it.ConversationID } catch { }
    if (-not $conv) { $conv = 'subj:' + (Get-NormSubject $subj).ToLowerInvariant() }

    [pscustomobject]@{
        E = $e; S = $s; FK = $fk; W = $when; Bulk = (Test-FadoeBulk $it $subj $body)
        Subj = $subj; SN = $sn; SA = $sa; To = ($to -join '; '); Cc = ($cc -join '; ')
        People = (($people | Where-Object { $_ }) -join ' ; '); Conv = $conv
        Own = $own; Flat = $flat; LSubj = $subj.ToLowerInvariant(); LFlat = $flat.ToLowerInvariant()
    }
}

# Bring the index up to date. Unchanged folders cost one property read each.
# -Full re-lists every folder (catches the rare add-and-delete that leaves a count unchanged).
function Sync-FadoeIndex {
    param([switch]$Full)
    Initialize-FadoeIndex
    $x = $global:FadoeIdx
    $act = 'Updating the search index'

    if ($Full -or $x.Folders.Count -eq 0 -or ((Get-Date) - $x.FolderListAt).TotalHours -ge 24) {
        Write-Progress -Activity $act -Status 'Finding your mail folders' -PercentComplete 0
        Update-FadoeFolderList
    }

    $since = (Get-Date).Date.AddDays(-$x.Days)
    $stamp = $since.ToString('yyyy-MM-dd HH:mm')
    $filter = '@SQL=("urn:schemas:httpmail:datereceived" >= ''' + $stamp + ''' OR "urn:schemas:httpmail:date" >= ''' + $stamp + ''')'

    $toFetch = New-Object System.Collections.ArrayList
    $pending = New-Object System.Collections.ArrayList
    $keys = @($x.Folders.Keys); $fi = 0
    foreach ($fk in $keys) {
        $fi++
        $fo = $x.Folders[$fk]
        $f = $fo.Obj
        if (-not $f) {
            try { $f = $x.Ns.GetFolderFromID($fk, $fo.S); $fo.Obj = $f }
            catch { $x.FolderListAt = [datetime]::MinValue; continue }   # gone? re-walk next time
        }
        $count = -1
        try { $count = [int](Get-FadoeProp $f $PR_CONTENT_COUNT) }
        catch {
            $its = $null
            try { $its = $f.Items; $count = $its.Count } catch { $fo.Obj = $null; continue } finally { Release-FadoeCom $its }
        }
        if (-not $Full -and $fo.Synced -and $count -eq $fo.Count) { continue }

        Write-Progress -Activity $act -Status ('Checking ' + $fo.Path) -PercentComplete ([int](10 * $fi / $keys.Count))
        $now = New-Object 'System.Collections.Generic.HashSet[string]'
        $tbl = $null; $cols = $null
        try {
            $tbl = $f.GetTable($filter, 0)
            $cols = $tbl.Columns
            $cols.RemoveAll()
            [void]$cols.Add('EntryID'); [void]$cols.Add('MessageClass')
            while (-not $tbl.EndOfTable) {
                $arr = $tbl.GetArray(5000)
                $rows = $arr.GetLength(0)
                for ($i = 0; $i -lt $rows; $i++) {
                    $cls = [string]($arr[$i, 1])
                    $eid = [string]($arr[$i, 0])
                    if ($cls -like 'IPM.Note*') { [void]$now.Add($eid) }
                }
            }
        } catch { continue }
        finally { Release-FadoeCom $cols $tbl }

        $have = $x.ByFolder[$fk]
        foreach ($e in @($have)) {
            if (-not $now.Contains($e)) { [void]$x.Recs.Remove($e); [void]$have.Remove($e); $x.Dirty = $true }
        }
        foreach ($e in $now) { if (-not $have.Contains($e)) { [void]$toFetch.Add(@($e, $fo.S, $fk)) } }
        [void]$pending.Add(@($fk, $count))
    }

    $k = 0; $fails = 0; $aborted = $false
    foreach ($t in $toFetch) {
        $k++
        if ($k % 20 -eq 1) {
            Write-Progress -Activity $act -Status ('Indexing {0:N0} of {1:N0} new messages' -f $k, $toFetch.Count) `
                           -PercentComplete ([int](10 + 90 * $k / $toFetch.Count))
        }
        $it = $null
        try {
            $it = $x.Ns.GetItemFromID($t[0], $t[1])
            $rec = New-FadoeRecord $it $t[2] $t[0] $t[1]
            if ($rec) { $x.Recs[$t[0]] = $rec; [void]$x.ByFolder[$t[2]].Add($t[0]); $x.Dirty = $true }
            $fails = 0
        } catch {
            # a burst of failures means Outlook is refusing -- stop now, keep what we have, retry next time
            if (++$fails -ge 25) { $aborted = $true; break }
        } finally { Release-FadoeCom $it }
        if ($k % 100 -eq 0) { Invoke-FadoeCleanup }
        if ($k % 2000 -eq 0) { Save-FadoeIndex }   # checkpoint during a first build
    }
    if (-not $aborted) {   # only mark folders done when all their new messages made it in
        foreach ($p in $pending) { $x.Folders[$p[0]].Count = $p[1]; $x.Folders[$p[0]].Synced = $true; $x.Dirty = $true }
    }

    # drop records that have aged out of the window (once a day is plenty)
    if ($x.ExpiredOn -ne (Get-Date).Date) {
        foreach ($r in @($x.Recs.Values | Where-Object { $_.W -lt $since })) {
            [void]$x.Recs.Remove($r.E); if ($x.ByFolder[$r.FK]) { [void]$x.ByFolder[$r.FK].Remove($r.E) }; $x.Dirty = $true
        }
        $x.ExpiredOn = (Get-Date).Date
    }
    Write-Progress -Activity $act -Completed
    if ($x.Dirty) { Save-FadoeIndex }
    Invoke-FadoeCleanup
    if ($aborted) {
        Write-Warning 'Outlook stopped handing over messages partway through. The index kept everything it got and will finish on the next search. If Outlook shows an "exhausted all shared resources" message, click OK and try again.'
    }
    [pscustomobject]@{ Messages = $x.Recs.Count; NewMessages = $toFetch.Count; FoldersChecked = $pending.Count; Incomplete = $aborted }
}

# ---------------------------------------------------------------- search ----
# Returns candidates (every matching message) and results (one per conversation, ranked).
function Search-FadoeIndex {
    param([string[]]$Terms, [string]$From, [int]$Top = 10, [int]$Days = 365)
    $x = $global:FadoeIdx
    $since = (Get-Date).Date.AddDays(-$Days)
    $cands = New-Object System.Collections.ArrayList
    foreach ($r in $x.Recs.Values) {
        if ($r.W -lt $since) { continue }
        # every term must appear in the subject or in what this sender wrote
        $ok = $true; $ownHits = 0
        foreach ($t in $Terms) {
            if ($r.LFlat.Contains($t)) { $ownHits++ } elseif (-not $r.LSubj.Contains($t)) { $ok = $false; break }
        }
        if (-not $ok) { continue }
        if ($From -and ("$($r.SN) $($r.SA)" -notlike "*$From*")) { continue }

        $sp = if ($ownHits -eq $Terms.Count) { Get-Span $r.LFlat $Terms } else { @{ Span = [int]::MaxValue; Start = 0 } }
        if ($ownHits -gt 0 -and $ownHits -lt $Terms.Count) {   # start the passage at a term they did write
            foreach ($t in $Terms) { $i = $r.LFlat.IndexOf($t, [StringComparison]::Ordinal); if ($i -ge 0) { $sp.Start = $i; break } }
        }
        $subjAll = $true
        foreach ($t in $Terms) { if (-not $r.LSubj.Contains($t)) { $subjAll = $false; break } }
        # small record per match; the full result object is only built for the few that are shown
        [void]$cands.Add([pscustomobject]@{
            R = $r; Conv = $r.Conv; Own = $r.Own; When = $r.W; IsBulk = [bool]$r.Bulk
            OwnHits = $ownHits; Tier = $(if ($subjAll) { -1 } else { Get-Tier $sp.Span }); Start = $sp.Start
        })
    }

    # One result per conversation: most terms in the sender's own text, tightest phrase, earliest.
    # (single dictionary pass -- Group-Object is far too slow when thousands of messages match)
    $bestOf = @{}
    foreach ($c in $cands) {
        $b = $bestOf[$c.Conv]
        if ($null -eq $b -or $c.OwnHits -gt $b.OwnHits -or
            ($c.OwnHits -eq $b.OwnHits -and ($c.Tier -lt $b.Tier -or ($c.Tier -eq $b.Tier -and $c.When -lt $b.When)))) {
            $bestOf[$c.Conv] = $c
        }
    }
    # Across conversations: real correspondence before bulk mail, most terms written, tightest
    # phrase, newest first -- packed into one sortable string key and sorted natively.
    $vals = @($bestOf.Values)
    $keys = New-Object string[] $vals.Count
    $maxT = [datetime]::MaxValue.Ticks
    for ($i = 0; $i -lt $vals.Count; $i++) {
        $c = $vals[$i]
        $keys[$i] = '{0}{1:00}{2}{3:D19}' -f [int]$c.IsBulk, (99 - $c.OwnHits), ($c.Tier + 1), ($maxT - $c.When.Ticks)
    }
    # Sort positions, not the objects: PowerShell would hand [Array]::Sort a converted *copy* of
    # an object array and the sort would be lost. int[]/string[] pass straight through. Keys
    # are equal-length digit strings, so the default comparison orders them correctly.
    $order = [int[]](0..([Math]::Max($vals.Count, 1) - 1))
    if ($vals.Count -gt 1) { [Array]::Sort($keys, $order) }
    $results = @(for ($i = 0; $i -lt [Math]::Min($Top, $vals.Count); $i++) {
        $c = $vals[$order[$i]]; $r = $c.R; $fo = $x.Folders[$r.FK]
        [pscustomobject]@{
            When = $r.W; Subject = $r.Subj; SenderName = $r.SN; SenderAddr = $r.SA; To = $r.To
            OwnHits = $c.OwnHits; Tier = $c.Tier; IsBulk = $c.IsBulk; Start = $c.Start
            Own = $r.Own; Flat = $r.Flat; Conv = $r.Conv; EntryID = $r.E; StoreID = $r.S
            Folder = $fo.Path; StoreName = $fo.StoreName; IsPrimary = $fo.IsPrimary; Pos = $null; Total = $null
        }
    })

    # Position in thread, straight from the index (copies in two folders count once).
    $want = @{}; foreach ($r in $results) { $want[$r.Conv] = New-Object 'System.Collections.Generic.HashSet[string]' }
    foreach ($r in $x.Recs.Values) { if ($want.ContainsKey($r.Conv)) { [void]$want[$r.Conv].Add($r.W.ToString('yyyyMMddHHmm')) } }
    foreach ($r in $results) {
        $times = $want[$r.Conv]
        if ($times.Count -ge 2) {
            $mine = $r.When.ToString('yyyyMMddHHmm')
            $r.Pos = [Math]::Max(1, @($times | Where-Object { $_ -le $mine }).Count)
            $r.Total = $times.Count
        }
    }
    [pscustomobject]@{ Candidates = $cands; Results = $results; Conversations = $bestOf.Count }
}
