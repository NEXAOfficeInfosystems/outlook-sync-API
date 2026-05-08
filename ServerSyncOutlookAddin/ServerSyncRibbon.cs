
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ServerSyncOutlookAddin
{
    public partial class ServerSyncRibbon
    {
        private void ServerSyncRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnAttachFromServer_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get the current active inspector (the email window)
                Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector == null) return;

                Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
                if (mailItem == null)
                {
                    MessageBox.Show("Please open an email message to use this feature.", "Context Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Resolve path to DMS Vault (OIS Sync) desktop application
                string electronAppPath;
                string electronArgs;
                bool isDevMode = false;

                GetElectronAppDetails(out electronAppPath, out electronArgs, out isDevMode);

                if (!File.Exists(electronAppPath))
                {
                    MessageBox.Show(
                        $"Could not find OIS Sync at:\n{electronAppPath}\n\nPlease ensure the application is installed.",
                        "OIS Sync Not Found",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Create a temporary file for results
                string tempResultPath = Path.Combine(Path.GetTempPath(), $"outlook_attach_{DateTime.Now.Ticks}.json");

                // Build arguments
                // In dev mode, the first arg is the app directory path, then our flags follow
                // In production, no leading app dir is needed
                string safeSubject = (mailItem.Subject ?? "").Replace("\"", "'");
                string outlookFlags = $"--outlook-invoke --output-file=\"{tempResultPath}\" --email-subject=\"{safeSubject}\"";
                string finalArgs = isDevMode
                    ? $"{electronArgs} {outlookFlags}"
                    : outlookFlags;

                // Launch OIS Sync with Outlook invoke arguments
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = electronAppPath,
                    Arguments = finalArgs,
                    UseShellExecute = false,
                    CreateNoWindow = false   // show window so user can interact with the DMS explorer
                };

                using (Process process = Process.Start(startInfo))
                {
                    var result = WaitForAttachmentResult(tempResultPath, process, TimeSpan.FromMinutes(15));

                    if (result != null)
                    {
                        // User closed the OIS Sync popup without clicking Done & Close.
                        // Nothing to attach — silently clean up and return.
                        if (result.Cancelled)
                        {
                            try { File.Delete(tempResultPath); } catch { }
                            return;
                        }

                        if (result.Files != null && result.Files.Any())
                        {
                            int count = 0;
                            var failed = new List<string>();
                            var failedReasons = new List<string>();

                            foreach (string filePath in result.Files)
                            {
                                string normPath = filePath;
                                try { normPath = Path.GetFullPath(filePath); } catch { }

                                try
                                {
                                    // Check file exists and is readable before handing to Outlook
                                    if (!File.Exists(normPath))
                                        throw new FileNotFoundException($"File not found: {normPath}");

                                    var fi = new FileInfo(normPath);
                                    long sizeMB = fi.Length / (1024 * 1024);

                                    mailItem.Attachments.Add(normPath);
                                    count++;
                                }
                                catch (Exception ex)  // ← capture the actual exception
                                {
                                    failed.Add(Path.GetFileName(normPath));
                                    failedReasons.Add($"{Path.GetFileName(normPath)}: {ex.Message}");
                                }
                            }

                            if (failed.Any())
                            {
                                MessageBox.Show(
                                    $"Attached {count} file(s).\n\nFailed attachments:\n{string.Join("\n", failedReasons)}",
                                    "Partial Attachment",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show(
                                "No files were returned from OIS Sync. Click Attach on at least one document, then click Done & Close.",
                                "No Attachments Selected",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }

                        // Clean up the temp result file
                        try { File.Delete(tempResultPath); } catch { }
                    }
                    else
                    {
                        MessageBox.Show(
                            "Timed out waiting for OIS Sync to finish the attachment selection. Please click Attach on a document and then click Done & Close.",
                            "Attachment Timed Out",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred while launching OIS Sync:\n\n{ex.Message}",
                    "Integration Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private OutlookAttachmentResult WaitForAttachmentResult(string tempResultPath, Process process, TimeSpan timeout)
        {
            DateTime deadlineUtc = DateTime.UtcNow.Add(timeout);

            while (DateTime.UtcNow < deadlineUtc)
            {
                try
                {
                    if (File.Exists(tempResultPath))
                    {
                        string json = File.ReadAllText(tempResultPath);
                        if (!string.IsNullOrWhiteSpace(json))
                        {
                            var result = JsonConvert.DeserializeObject<OutlookAttachmentResult>(json);
                            if (result != null)
                            {
                                // Return immediately whether the user completed or cancelled;
                                // the caller inspects result.Cancelled to decide what to do.
                                return result;
                            }
                        }
                    }
                }
                catch (IOException)
                {
                    // File may still be being written; retry.
                }
                catch (UnauthorizedAccessException)
                {
                    // File may still be locked; retry.
                }
                catch (JsonException)
                {
                    // Partial JSON while write is in progress; retry.
                }

                try
                {
                    if (process != null && process.HasExited)
                    {
                        // In single-instance mode the launcher process exits immediately,
                        // but the already-running OIS Sync instance will write the file later.
                    }
                }
                catch
                {
                    // Ignore process state errors and keep waiting on the temp file.
                }

                Application.DoEvents();
                System.Threading.Thread.Sleep(250);
            }

            return null;
        }

        /// <summary>
        /// Resolves the path to the OIS Sync (DMS Vault) Electron application.
        /// Falls back to developer mode (electron.cmd + app folder) when the
        /// packaged executable is not found.
        /// </summary>
        private void GetElectronAppDetails(out string appPath, out string extraArgs, out bool isDevMode)
        {
            string localAppData    = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string programFiles    = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            // 1. Installed / packaged executable locations (electron-builder output)
            string[] installedPaths = new string[]
            {
                Path.Combine(localAppData,    "Programs", "OIS Sync", "OIS Sync.exe"),
                Path.Combine(programFiles,    "OIS Sync", "OIS Sync.exe"),
                Path.Combine(programFilesX86, "OIS Sync", "OIS Sync.exe"),
                // Common dist output from electron-builder in the dev workspace
                //@"D:\AntiGravity Proj\DMS_Vault_Desktop_Application\dist\win-unpacked\OIS Sync.exe",
            };

            foreach (string p in installedPaths)
            {
                if (File.Exists(p))
                {
                    appPath   = p;
                    extraArgs = string.Empty;
                    isDevMode = false;
                    return;
                }
            }

            // 2. Developer fallback: launch via electron.cmd in the project folder
            string electronCmd = @"D:\AntiGravity Proj\DMS_Vault_Desktop_Application\node_modules\.bin\electron.cmd";
            if (File.Exists(electronCmd))
            {
                appPath   = electronCmd;
                // Pass the DMS Vault app directory as the first positional arg to electron
                extraArgs = "\"D:\\AntiGravity Proj\\DMS_Vault_Desktop_Application\"";
                isDevMode = true;
                return;
            }

            // 3. Last resort: return first installed path (File.Exists check will fail gracefully)
            appPath   = installedPaths[0];
            extraArgs = string.Empty;
            isDevMode = false;
        }
    }

    public class OutlookAttachmentResult
    {
        [JsonProperty("files")]
        public List<string> Files { get; set; }

        [JsonProperty("timestamp")]
        public long Timestamp { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        /// <summary>
        /// True when the user closed the OIS Sync popup without clicking Done &amp; Close.
        /// The Files list will be empty in this case.
        /// </summary>
        [JsonProperty("cancelled")]
        public bool Cancelled { get; set; }

        [JsonProperty("reason")]
        public string Reason { get; set; }
    }
}
