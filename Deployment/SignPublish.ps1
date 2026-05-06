# SignPublish.ps1
# This script manually signs the VSTO published files.
# It targets both the custom publish path and the default bin\Release\app.publish path.

param(
    [string]$CustomPath = "D:\Outlook Addin publish"
)

$PfxPath = "OIS_Sync.pfx"
$Password = "OisSync2026!"

Write-Host "Manual Signing Start..." -ForegroundColor Cyan

# Find Tools
$signtool = Get-Command "signtool.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
if (-not $signtool) {
    $searchPaths = @(
        "${env:ProgramFiles(x86)}\Windows Kits\11\bin",
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin",
        "${env:ProgramFiles(x86)}\Microsoft SDKs\ClickOnce\SignTool"
    )
    foreach ($path in $searchPaths) {
        if (Test-Path $path) {
            $signtool = Get-ChildItem $path -Filter "signtool.exe" -Recurse | 
                        Where-Object { $_.FullName -like "*\x64\*" -and $_.FullName -notlike "*\arm64\*" } |
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
            $mage = Get-ChildItem $path -Filter "mage.exe" -Recurse | 
                    Where-Object { $_.FullName -notlike "*\arm64\*" } |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Select-Object -ExpandProperty FullName
            if ($mage) { break }
        }
    }
}

if (-not $signtool -or -not $mage) {
    Write-Error "Required tools (signtool or mage) not found."
    return
}

# Determine Paths to sign
$PathsToSign = @()
if (Test-Path $CustomPath) { $PathsToSign += $CustomPath }
$DefaultPath = "..\ServerSyncOutlookAddin\bin\Release\app.publish"
if (Test-Path $DefaultPath) { $PathsToSign += (Resolve-Path $DefaultPath).Path }

if ($PathsToSign.Count -eq 0) {
    Write-Error "No publish folders found to sign!"
    return
}

foreach ($PublishPath in $PathsToSign) {
    Write-Host "`n--- Signing Folder: $PublishPath ---" -ForegroundColor Yellow

    # 1. Sign the setup.exe
    $setupExe = Join-Path $PublishPath "setup.exe"
    if (Test-Path $setupExe) {
        Write-Host "Signing setup.exe..."
        & $signtool sign /f $PfxPath /p $Password /fd sha256 /t http://timestamp.digicert.com /v $setupExe
    }

    # 2. Process Application Files
    $appFilesRoot = Join-Path $PublishPath "Application Files"
    if (Test-Path $appFilesRoot) {
        $appFilesFolder = Get-ChildItem $appFilesRoot -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-Host "Processing Application Files in: $($appFilesFolder.FullName)"

        $manifestFile = Get-ChildItem $appFilesFolder.FullName -Filter "*.dll.manifest" | Select-Object -First 1
        $vstoFile = Join-Path $PublishPath "ServerSyncOutlookAddin.vsto"

        # Temporarily rename .deploy files
        Write-Host "Temporarily removing .deploy extensions..."
        $deployFiles = Get-ChildItem $appFilesFolder.FullName -Filter "*.deploy"
        foreach ($f in $deployFiles) {
            $newName = $f.FullName -replace ".deploy", ""
            if (Test-Path $newName) { Remove-Item $newName -Force }
            Rename-Item $f.FullName $newName
        }

        # Sign the main DLL
        $dllFile = Join-Path $appFilesFolder.FullName "ServerSyncOutlookAddin.dll"
        if (Test-Path $dllFile) {
            & $signtool sign /f $PfxPath /p $Password /fd sha256 /t http://timestamp.digicert.com /v $dllFile
            
            # Verify DLL while it is still named .dll
            Write-Host "Verifying Main DLL... " -NoNewline
            $verifyDll = & $signtool verify /pa /v $dllFile 2>&1
            if ($LASTEXITCODE -eq 0 -or $verifyDll -like "*not trusted*") { 
                Write-Host "[OK]" -ForegroundColor Green 
            } else { 
                Write-Host "[FAILED]" -ForegroundColor Red 
                Write-Host "$verifyDll" -ForegroundColor Gray
            }
        }

        # Update and Sign Application Manifest
        Write-Host "Updating Application Manifest..."
        & $mage -update $manifestFile.FullName -CertFile $PfxPath -Password $Password
        
        Write-Host "Verifying Application Manifest... " -NoNewline
        $verifyAppMan = & $mage -verify $manifestFile.FullName 2>&1
        if ($verifyAppMan -like "*valid signature*") { 
            Write-Host "[OK]" -ForegroundColor Green 
        } else { 
            Write-Host "[FAILED]" -ForegroundColor Red 
            Write-Host "$verifyAppMan" -ForegroundColor Gray
        }

        # Rename back to .deploy
        Write-Host "Restoring .deploy extensions..."
        $noDeployFiles = Get-ChildItem $appFilesFolder.FullName | Where-Object { $_.Extension -ne ".manifest" -and $_.Extension -ne ".pfx" }
        foreach ($f in $noDeployFiles) {
            if (-not $f.Name.EndsWith(".deploy")) {
                Rename-Item $f.FullName ($f.Name + ".deploy")
            }
        }

        # Update and Sign Deployment Manifest (.vsto)
        if (Test-Path $vstoFile) {
            Write-Host "Updating Deployment Manifest (.vsto) with Mage..."
            & $mage -update $vstoFile -AppManifest $manifestFile.FullName -CertFile $PfxPath -Password $Password
            
            Write-Host "Verifying Deployment Manifest (.vsto)... " -NoNewline
            $verifyVsto = & $mage -verify $vstoFile 2>&1
            if ($verifyVsto -like "*valid signature*") { 
                Write-Host "[OK]" -ForegroundColor Green 
            } else { 
                Write-Host "[FAILED]" -ForegroundColor Red 
                Write-Host "$verifyVsto" -ForegroundColor Gray
            }
        }

        # Verify setup.exe
        Write-Host "Verifying setup.exe... " -NoNewline
        $verifySetup = & $signtool verify /pa /v $setupExe 2>&1
        if ($LASTEXITCODE -eq 0 -or $verifySetup -like "*not trusted*") { 
            Write-Host "[OK]" -ForegroundColor Green 
        } else { 
            Write-Host "[FAILED]" -ForegroundColor Red 
            Write-Host "$verifySetup" -ForegroundColor Gray
        }
    }
}

Write-Host "`nMANUAL SIGNING AND VERIFICATION COMPLETE!" -ForegroundColor Green
Write-Host "If all statuses are [OK], you can safely distribute this folder."





