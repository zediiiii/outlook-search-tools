<#
.SYNOPSIS
    Installs FADOE for the current user and adds desktop + Start menu shortcuts.

.DESCRIPTION
    Copies FADOE into %LOCALAPPDATA%\Programs\FADOE (no admin rights needed) and creates
    shortcuts. Run it again from a newer release to upgrade; your saved catch-up files in
    the context folder are left alone.

    Usually started by double-clicking Install.cmd.

.PARAMETER InstallDir
    Where to install. Default: %LOCALAPPDATA%\Programs\FADOE

.PARAMETER NoShortcuts
    Copy the files but don't create shortcuts.
#>

[CmdletBinding()]
param(
    [string] $InstallDir = (Join-Path $env:LOCALAPPDATA 'Programs\FADOE'),
    [switch] $NoShortcuts
)

$ErrorActionPreference = 'Stop'
$src   = $PSScriptRoot
$files = 'FADOE.ps1', 'FadoeIndex.ps1', 'Find-Email.ps1', 'Get-EmailContext.ps1', 'Uninstall-FADOE.ps1', 'README.md'

Write-Host ''
Write-Host '  Installing FADOE - Find A Damn Outlook Email' -ForegroundColor Cyan
Write-Host ''

foreach ($f in $files) {
    if (-not (Test-Path (Join-Path $src $f))) {
        throw "Can't find $f next to the installer. Extract the whole zip first, then run Install.cmd from the extracted folder."
    }
}

if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Outlook.Application')) {
    Write-Host '  Heads up: classic Outlook does not seem to be installed on this computer.' -ForegroundColor Yellow
    Write-Host '  FADOE reads your mailbox through classic Outlook in the background (you can keep' -ForegroundColor Yellow
    Write-Host '  using new Outlook day to day). Install classic Outlook and sign in to it once first.' -ForegroundColor Yellow
    Write-Host ''
}

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
foreach ($f in $files) {
    $dest = Join-Path $InstallDir $f
    Copy-Item -LiteralPath (Join-Path $src $f) -Destination $dest -Force
    Unblock-File -LiteralPath $dest   # drop the "downloaded from the internet" flag
}
$version = ([regex]::Match((Get-Content -Raw (Join-Path $InstallDir 'FADOE.ps1')), "FadoeVersion = '([^']+)'")).Groups[1].Value
# marker the uninstaller checks, so it only ever removes a real install
Set-Content -Path (Join-Path $InstallDir '.fadoe-install') -Value "FADOE $version installed $(Get-Date -Format s)"
Write-Host "  Installed FADOE $version to $InstallDir"

if (-not $NoShortcuts) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $InstallDir 'FADOE.ps1') -InstallShortcut |
        ForEach-Object { Write-Host "  $_" }
}

Write-Host ''
Write-Host '  Done! Open FADOE from the desktop shortcut or the Start menu.' -ForegroundColor Green
Write-Host '  To upgrade later, download the newer release and run Install.cmd again.'
Write-Host ''
