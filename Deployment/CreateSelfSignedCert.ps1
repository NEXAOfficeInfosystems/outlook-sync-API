# CreateSelfSignedCert.ps1
# This script generates a self-signed code signing certificate for OIS Sync.

$subject = "CN=OIS Sync"
$passwordStr = "OisSync2026!" # Change this if needed
$password = ConvertTo-SecureString $passwordStr -AsPlainText -Force
$outDir = $PSScriptRoot

if (-not $outDir) { $outDir = Get-Location }

Write-Host "Generating certificate for $subject..." -ForegroundColor Cyan

# 1. Create the certificate in the local store
$cert = New-SelfSignedCertificate -Type CodeSigning `
    -Subject $subject `
    -KeyExportPolicy Exportable `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotAfter (Get-Date).AddYears(5) `
    -FriendlyName "OIS Sync Code Signing"

# 2. Export to PFX (Private Key + Public Key) - Used for signing during build
$pfxPath = Join-Path $outDir "OIS_Sync.pfx"
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password
Write-Host "Successfully exported PFX to: $pfxPath" -ForegroundColor Green

# 3. Export to CER (Public Key only) - Given to clients to install
$cerPath = Join-Path $outDir "OIS_Sync.cer"
Export-Certificate -Cert $cert -FilePath $cerPath
Write-Host "Successfully exported CER to: $cerPath" -ForegroundColor Green

Write-Host "`nIMPORTANT:" -ForegroundColor Yellow
Write-Host "1. Keep the PFX file secure. It contains your private key."
Write-Host "2. Distribute the .cer file with your installer."
Write-Host "3. The PFX password is: $passwordStr"
