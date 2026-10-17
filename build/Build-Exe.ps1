#Requires -Version 5.1
<#
Build Thinka Excel Converter Windows GUI EXE with PS2EXE.
Requires Windows PowerShell 5.1 on Windows 10/11.
Usage:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build\Build-Exe.ps1 -InstallPs2Exe
#>
[CmdletBinding()]
param(
    [switch]$InstallPs2Exe,
    [string]$OutputDirectory = ''
)
$ErrorActionPreference = 'Stop'
$version = '1.0.18'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$source = Join-Path $root 'Thinka_Excel_Converter.ps1'
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "Main source file not found: $source"
}
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or
    $PSVersionTable.PSEdition -ne 'Desktop') {
    throw 'Please build on Windows PowerShell 5.1 (powershell.exe).'
}
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $root 'dist' }
[void](New-Item -ItemType Directory -Force -Path $OutputDirectory)
$output = Join-Path $OutputDirectory 'Thinka_Excel_Converter.exe'

$module = Get-Module -ListAvailable -Name ps2exe |
    Where-Object { $_.Version -eq [version]$version } |
    Select-Object -First 1
if (-not $module) {
    if (-not $InstallPs2Exe) {
        throw "PS2EXE $version is missing. Rerun with -InstallPs2Exe or install it from the PowerShell Gallery."
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion '2.8.5.201' -Scope CurrentUser -Force | Out-Null
    }
    Install-Module -Name ps2exe -RequiredVersion $version -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
}
Import-Module -Name ps2exe -RequiredVersion $version -Force -ErrorAction Stop

Write-Host 'Building Windows GUI executable (no console)...'
Invoke-PS2EXE -inputFile $source -outputFile $output -noConsole -STA -x64 `
    -title 'Thinka Excel Converter' `
    -product 'Thinka CSV Excel Converter' `
    -description 'Thinka Excel bulk device editor with KNX preview and verified CSV export' `
    -version '1.5.0.0'

if (-not (Test-Path -LiteralPath $output -PathType Leaf)) {
    throw 'PS2EXE did not create an EXE file.'
}
# Check PE header and Windows GUI subsystem (2), not console subsystem (3).
$bytes = [IO.File]::ReadAllBytes($output)
if ($bytes.Length -lt 512 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
    throw 'Compiled output is not a Windows PE executable.'
}
$peOffset = [BitConverter]::ToInt32($bytes, 0x3C)
if ($peOffset -lt 0 -or $peOffset + 100 -ge $bytes.Length) {
    throw 'Invalid PE header offset.'
}
if ([BitConverter]::ToUInt32($bytes, $peOffset) -ne 0x00004550) {
    throw 'Invalid PE signature.'
}
$subsystem = [BitConverter]::ToUInt16($bytes, ($peOffset + 24 + 68))
if ($subsystem -ne 2) {
    throw "Unexpected EXE subsystem $subsystem. Expected Windows GUI (2)."
}
Write-Host ("Build succeeded: {0} ({1:N0} bytes), Windows GUI subsystem" -f $output,$bytes.Length)
