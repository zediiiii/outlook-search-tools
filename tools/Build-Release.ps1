<#
.SYNOPSIS
    Builds dist\FADOE-v<version>.zip -- the download attached to a GitHub release.

.DESCRIPTION
    The version comes from $FadoeVersion in FADOE.ps1. The zip holds one folder with the
    installer and the three scripts; users extract it and double-click Install.cmd.

    To publish (after committing):
        gh release create v<version> dist\FADOE-v<version>.zip --title "FADOE <version>" --notes-file <notes>
#>

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$ver  = ([regex]::Match((Get-Content -Raw (Join-Path $root 'FADOE.ps1')), "FadoeVersion = '([^']+)'")).Groups[1].Value
if (-not $ver) { throw 'Could not read $FadoeVersion from FADOE.ps1' }

$files = 'Install.cmd', 'Install-FADOE.ps1', 'Uninstall-FADOE.ps1',
         'FADOE.ps1', 'Find-Email.ps1', 'Get-EmailContext.ps1', 'README.md'

$dist  = Join-Path $root 'dist'
$name  = "FADOE-v$ver"
$stage = Join-Path ([IO.Path]::GetTempPath()) ("fadoe-build-" + [guid]::NewGuid().ToString('N'))
$inner = Join-Path $stage $name
New-Item -ItemType Directory -Force -Path $inner, $dist | Out-Null
foreach ($f in $files) { Copy-Item -LiteralPath (Join-Path $root $f) -Destination $inner }

$zip = Join-Path $dist "$name.zip"
if (Test-Path $zip) { Remove-Item -LiteralPath $zip }
Compress-Archive -Path $inner -DestinationPath $zip
Remove-Item -LiteralPath $stage -Recurse -Force

Write-Host "Built $zip"
$zip
