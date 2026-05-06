@echo off
:: InstallCert.bat
:: This script installs the OIS Sync self-signed certificate into the local machine's trusted stores.
:: This is required for self-signed apps to run without security warnings.
:: 
:: IMPORTANT: THIS SCRIPT MUST BE RUN AS ADMINISTRATOR.

set "CERT_FILE=OIS_Sync.cer"

:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo ERROR: This script MUST be run as ADMINISTRATOR.
    echo Please right-click this file and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

if not exist "%~dp0%CERT_FILE%" (
    echo.
    echo ERROR: Certificate file "%CERT_FILE%" not found in this directory.
    echo.
    pause
    exit /b 1
)

echo.
echo Installing OIS Sync Certificate...
echo ---------------------------------

:: Install to Trusted Root Certification Authorities (so the chain is trusted)
echo 1. Adding to Trusted Root...
certutil -addstore -f "Root" "%~dp0%CERT_FILE%"

:: Install to Trusted Publishers (so the code signing is trusted)
echo 2. Adding to Trusted Publishers...
certutil -addstore -f "TrustedPublisher" "%~dp0%CERT_FILE%"

echo.
echo SUCCESS: The certificate has been installed.
echo You can now install the OIS Sync Outlook Add-in and Desktop App.
echo.
pause
