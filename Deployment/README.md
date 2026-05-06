# OIS Sync Deployment Guide

This folder contains the tools and scripts needed to publish the OIS Sync Outlook Add-in and Electron App using a self-signed certificate.

## Developer Instructions (Publishing)

### 1. Generate the Certificate (One-time)
If you haven't already, run the following script to generate your code-signing certificate:
- **File**: `CreateSelfSignedCert.ps1`
- **How to Run**: Right-click and "Run with PowerShell" or run:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\CreateSelfSignedCert.ps1
  ```
- **Output**: `OIS_Sync.pfx` (private key) and `OIS_Sync.cer` (public key).
- **Password**: `OisSync2026!`

### 2. Configure Visual Studio Signing (CRITICAL)
To ensure the Add-in is trusted on client machines, you MUST configure it in Visual Studio:
1. Copy `OIS_Sync.pfx` to the `ServerSyncOutlookAddin` project folder.
2. In Visual Studio, right-click the **ServerSyncOutlookAddin** project -> **Properties**.
3. Go to the **Signing** tab.
4. Check **Sign the ClickOnce manifests**.
5. Click **Select from File...** and choose `OIS_Sync.pfx`.
6. Enter the password: `OisSync2026!`.
7. Go to the **Security** tab and ensure **Enable ClickOnce security settings** is checked.

### 3. Create the Setup and Distribute
1. Build your project in **Release** mode.
2. (Optional) Run `SignVSTO.ps1` if you need to manually sign the output binaries for non-ClickOnce distribution.
3. In Visual Studio, right-click the project -> **Publish...**.
4. Follow the wizard to publish to a folder (e.g., `D:\Outlook Addin publish\`).
5. **Distribution**: Copy the entire contents of the publish folder (including `setup.exe`, `ServerSyncOutlookAddin.vsto`, and the `Application Files` folder) to the client machine or a shared location.

### 4. Sign the Electron App
1. Copy `OIS_Sync.pfx` to the root of the Electron project.
2. Ensure your `package.json` build configuration points to this certificate.
3. Run `npm run build`.

---

## Client Instructions (Installation)

To install the application on a user's machine, follow these exact steps:

1. **Install the Trust Certificate (MUST BE FIRST)**:
   - Copy the `ClientSetup` folder to the client machine.
   - Right-click `InstallCert.bat` and select **Run as Administrator**.
   - This only needs to be done once per machine.

2. **Install the Outlook Add-in**:
   - Navigate to the published folder.
   - Run **setup.exe**.
   - If prompted with a security warning, it should now show **"Publisher: OIS"** (trusted) instead of "Unknown Publisher".

3. **Install the Electron App**:
   - Run the Electron App `-Setup.exe`.

---

## Troubleshooting

- **"Publisher cannot be verified"**: 
  - Ensure the `InstallCert.bat` was run as Administrator.
  - Verify that the Add-in was signed with `OIS_Sync.pfx` in Visual Studio before publishing.
- **Signtool not found**: Run `SignVSTO.ps1` from the **Developer PowerShell for VS 2022**.
- **Installation Aborted**: Check if the client machine has the required .NET Framework (v4.7.2) and VSTO Runtime installed. The `setup.exe` should attempt to install these automatically.

