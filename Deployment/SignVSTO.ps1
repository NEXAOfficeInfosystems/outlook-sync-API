# SignVSTO.ps1
# This script signs the Outlook Add-in (VSTO) files.
# Run this AFTER building the project in 'Release' mode.

param(
    [string]$ReleasePath = "..\ServerSyncOutlookAddin\bin\Release",
    [string]$PfxPath = "OIS_Sync.pfx",
    [string]$Password = "OisSync2026!"
)

$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
$absolutePfxPath = Join-Path $PSScriptRoot $PfxPath
$absoluteReleasePath = Resolve-Path (Join-Path $PSScriptRoot $ReleasePath)

Write-Host "Signing VSTO files in: $absoluteReleasePath" -ForegroundColor Cyan

# Common paths for Signtool and Mage
$signtool = Get-Command "signtool.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source

if (-not $signtool) {
    # Search for x64 signtool in Windows Kits 10 or 11
    $searchPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\11\bin",
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin",
        "${env:ProgramFiles(x86)}\Microsoft SDKs\ClickOnce\SignTool"
    )

    foreach ($path in $searchPaths) {
        if (Test-Path $path) {
            $signtool = Get-ChildItem $path -Filter "signtool.exe" -Recurse | 
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object -ExpandProperty FullName
            if ($signtool) { break }
        }
    }
}

$mage = Get-Command "mage.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
if (-not $mage) {
    $searchPathsMage = @(
        "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows",
        "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools"
    )
    foreach ($path in $searchPathsMage) {
        if (Test-Path $path) {
            $mage = Get-ChildItem $path -Filter "mage.exe" -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object -ExpandProperty FullName
            if ($mage) { break }
        }
    }
}

if (-not $signtool) {
    Write-Error "signtool.exe not found. Please:"
    Write-Host "1. Install the 'Windows 10/11 SDK' via Visual Studio Installer." -ForegroundColor Yellow
    Write-Host "2. OR run this script from the 'Developer PowerShell for VS 2022'." -ForegroundColor Yellow
    return
}

Write-Host "Using Signtool: $signtool"
if ($mage) { Write-Host "Using Mage: $mage" }

# 1. Sign the DLL
$dllPath = Join-Path $absoluteReleasePath "ServerSyncOutlookAddin.dll"
if (Test-Path $dllPath) {
    & $signtool sign /f $absolutePfxPath /p $Password /t http://timestamp.digicert.com /v $dllPath
}

# 2. Sign the Manifest
$manifestPath = Join-Path $absoluteReleasePath "ServerSyncOutlookAddin.dll.manifest"
if (Test-Path $manifestPath) {
    if ($mage) {
        & $mage -sign $manifestPath -CertFile $absolutePfxPath -Password $Password
    } else {
        & $signtool sign /f $absolutePfxPath /p $Password /t http://timestamp.digicert.com /v $manifestPath
    }
}

# 3. Sign the VSTO file
$vstoPath = Join-Path $absoluteReleasePath "ServerSyncOutlookAddin.vsto"
if (Test-Path $vstoPath) {
    if ($mage) {
        # Update the manifest reference in the VSTO file
        & $mage -update $vstoPath -AppManifest $manifestPath
        & $mage -sign $vstoPath -CertFile $absolutePfxPath -Password $Password
    } else {
        & $signtool sign /f $absolutePfxPath /p $Password /t http://timestamp.digicert.com /v $vstoPath
    }
}

Write-Host "`nSigning Complete!" -ForegroundColor Green

