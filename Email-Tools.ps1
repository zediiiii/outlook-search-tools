<#
.SYNOPSIS
    FADOE - Find A Damn Outlook Email. One menu for both tools in this folder.

.DESCRIPTION
    1  Find an email         Find-Email.ps1: the exact message in a reply chain where your
                             words were written, plus a search string for new Outlook.
    2  Catch up on a person  Get-EmailContext.ps1: every email with them, full text and dates,
                             copied to the clipboard to paste into Claude.

    Double-click the desktop shortcut, or run:
        powershell -ExecutionPolicy Bypass -File Email-Tools.ps1

.PARAMETER InstallShortcut
    Put a "FADOE" shortcut on your desktop that opens this menu, then exit.
#>

[CmdletBinding()]
param([switch] $InstallShortcut)

$here      = $PSScriptRoot
$findTool  = Join-Path $here 'Find-Email.ps1'
$ctxTool   = Join-Path $here 'Get-EmailContext.ps1'
$ctxFolder = Join-Path $here 'context'

# ------------------------------------------------------------ shortcut ----
if ($InstallShortcut) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $lnkPath = Join-Path $desktop 'FADOE.lnk'
    $ws  = New-Object -ComObject WScript.Shell
    $lnk = $ws.CreateShortcut($lnkPath)
    $lnk.TargetPath       = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $lnk.Arguments        = '-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '"'
    $lnk.WorkingDirectory = $here
    $lnk.IconLocation     = (Join-Path $env:SystemRoot 'System32\shell32.dll') + ',22'
    $lnk.Description      = 'FADOE - Find A Damn Outlook Email, or catch up on a person'
    $lnk.Save()
    Write-Host "Shortcut created: $lnkPath"
    return
}

# ------------------------------------------------------------- helpers ----
function Wait-ForMenu { [void](Read-Host 'Press Enter to go back to the menu') }

function Invoke-Find {
    Write-Host ''
    $words = Read-Host 'What words do you remember? (put "quotes" around a phrase)'
    if (-not $words -or -not $words.Trim()) { return }
    $from = Read-Host 'Who sent it? (part of a name or address -- or just press Enter to skip)'

    $params = @{ Search = $words.Trim(); PassThru = $true }
    if ($from -and $from.Trim()) { $params.From = $from.Trim() }

    Write-Host 'Searching... (about 30 seconds)' -ForegroundColor DarkGray
    $res = @(& $findTool @params)
    if ($res.Count -eq 0) { Write-Host ''; Wait-ForMenu; return }

    while ($true) {
        Write-Host ''
        $pick = Read-Host 'Type a result number to copy its search (#1 is already copied), or press Enter for the menu'
        if (-not $pick -or -not $pick.Trim()) { return }
        $n = 0
        if ([int]::TryParse($pick.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $res.Count) {
            Set-Clipboard -Value $res[$n - 1].Query
            Write-Host ("Copied #{0}: {1}" -f $n, $res[$n - 1].Query) -ForegroundColor Green
        } else {
            Write-Host "Pick a number from 1 to $($res.Count)." -ForegroundColor Yellow
        }
    }
}

function Invoke-CatchUp {
    Write-Host ''
    $who = Read-Host 'Whose emails? (their email address works best; separate several with commas)'
    if (-not $who -or -not $who.Trim()) { return }
    $daysIn = Read-Host 'How many days back? (press Enter for 180)'
    $days = 180; $d = 0
    if ($daysIn -and [int]::TryParse($daysIn.Trim(), [ref]$d) -and $d -gt 0) { $days = $d }

    $people = @($who -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not (Test-Path $ctxFolder)) { New-Item -ItemType Directory -Path $ctxFolder | Out-Null }
    $slug = ($people[0] -replace '[^A-Za-z0-9]+', '-').Trim('-')
    $out  = Join-Path $ctxFolder ('{0}_{1}.md' -f $slug, (Get-Date -Format 'yyyy-MM-dd'))
    if (Test-Path $out) { Remove-Item $out }   # same person, same day: regenerate

    Write-Host 'Gathering emails... (can take a minute)' -ForegroundColor DarkGray
    & $ctxTool -Person $people -Days $days -OutFile $out | Out-Null
    if (-not (Test-Path $out)) { Write-Host ''; Wait-ForMenu; return }   # nothing matched

    Set-Clipboard -Value (Get-Content -Raw -Encoding UTF8 $out)
    Write-Host ''
    Write-Host 'The full text is on your clipboard -- paste it into Claude and ask for a summary.' -ForegroundColor Green
    $o = Read-Host 'Type O to show the file in its folder, or press Enter for the menu'
    if ($o -and $o.Trim() -match '^[oO]') {
        Start-Process explorer.exe -ArgumentList ('/select,"' + $out + '"')
    }
}

# ---------------------------------------------------------------- menu ----
try { $Host.UI.RawUI.WindowTitle = 'FADOE - Find A Damn Outlook Email' } catch { }

while ($true) {
    try { Clear-Host } catch { }
    Write-Host ''
    Write-Host '  ==============================================================' -ForegroundColor Cyan
    Write-Host '    FADOE  -  Find A Damn Outlook Email' -ForegroundColor Cyan
    Write-Host '  ==============================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '    1   Find an email'
    Write-Host '        the exact message in a thread + a search to paste into new Outlook' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '    2   Catch up on a person'
    Write-Host '        every email with them, full text, ready to paste into Claude' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '    Q   Quit'
    Write-Host ''
    $choice = Read-Host '  Choose'
    if ($null -eq $choice) { return }
    $choice = $choice.Trim()

    try {
        if ($choice -eq '1')          { Invoke-Find }
        elseif ($choice -eq '2')      { Invoke-CatchUp }
        elseif ($choice -match '^[qQ]') { return }
    } catch {
        Write-Host ''
        Write-Host "Something went wrong: $($_.Exception.Message)" -ForegroundColor Red
        Wait-ForMenu
    }
}
