<#
.SYNOPSIS
    Removes FADOE: its shortcuts and program files.

.DESCRIPTION
    Run it from the install folder (%LOCALAPPDATA%\Programs\FADOE):
        powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Programs\FADOE\Uninstall-FADOE.ps1"

    Only works on a folder the installer created, and only removes shortcuts that point
    into that folder. Your saved catch-up files (the context folder) are left in place,
    because they're your data -- delete them yourself if you don't need them.
#>

[CmdletBinding()]
param([string] $InstallDir)

$ErrorActionPreference = 'Stop'
if (-not $InstallDir) { $InstallDir = $PSScriptRoot }   # not available in param defaults on PS 5.1

if (-not (Test-Path (Join-Path $InstallDir '.fadoe-install'))) {
    Write-Host "This doesn't look like a FADOE install folder ($InstallDir), so nothing was removed." -ForegroundColor Yellow
    return
}

$ws = New-Object -ComObject WScript.Shell
foreach ($folder in 'Desktop', 'Programs') {
    $lnkPath = Join-Path ([Environment]::GetFolderPath($folder)) 'FADOE.lnk'
    if ((Test-Path $lnkPath) -and ($ws.CreateShortcut($lnkPath).Arguments -like "*$InstallDir*")) {
        Remove-Item -LiteralPath $lnkPath
        Write-Host "Removed shortcut: $lnkPath"
    }
}

foreach ($f in 'FADOE.ps1', 'FadoeIndex.ps1', 'Find-Email.ps1', 'Get-EmailContext.ps1', 'README.md', 'FADOE.ico', '.fadoe-install') {
    $p = Join-Path $InstallDir $f
    if (Test-Path $p) { Remove-Item -LiteralPath $p -Force }
}

# The search index is a cache holding copies of your email text -- remove it too.
$cacheRoot = Join-Path $env:LOCALAPPDATA 'FADOE'
if (Test-Path $cacheRoot) {
    Get-ChildItem -LiteralPath $cacheRoot -Directory -Filter 'index-*' | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
    }
    if (-not (Get-ChildItem -LiteralPath $cacheRoot -Force)) { Remove-Item -LiteralPath $cacheRoot }
    Write-Host 'Removed the search index.'
}

$ctx = Join-Path $InstallDir 'context'
if (Test-Path $ctx) {
    Write-Host ''
    Write-Host "Your saved catch-up files are still in: $ctx"
    Write-Host 'They contain copies of your email. Delete that folder yourself if you no longer need them.'
}

Remove-Item -LiteralPath $PSCommandPath -Force
if (-not (Get-ChildItem -LiteralPath $InstallDir -Force)) { Remove-Item -LiteralPath $InstallDir }
Write-Host ''
Write-Host 'FADOE has been removed.' -ForegroundColor Green
