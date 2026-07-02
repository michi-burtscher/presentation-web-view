using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

namespace LiveWebRegionSetup
{
    /// <summary>
    /// Self-contained installer for the Live Web Region PowerPoint add-in.
    /// Extracts the embedded payload to %LOCALAPPDATA%\LiveWebRegion\app and writes
    /// the per-user COM + add-in registration (no admin required).
    /// Run with "/uninstall" to remove it again.
    /// </summary>
    internal static class Program
    {
        private const string Clsid = "{7E9B2C14-3A6D-4F58-9C1E-2B7A5D0F8E31}";
        private const string ProgId = "LiveWebRegion.AddIn";
        private const string AppName = "Live Web Region";
        private const string AssemblyFull = "LiveWebRegion, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null";

        private static string AppDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LiveWebRegion", "app");

        [STAThread]
        private static int Main(string[] args)
        {
            bool uninstall = Array.Exists(args, a => a.Equals("/uninstall", StringComparison.OrdinalIgnoreCase));
            bool update = Array.Exists(args, a => a.Equals("/update", StringComparison.OrdinalIgnoreCase));
            try
            {
                if (update)
                {
                    // Launched by the add-in's Update button: wait for PowerPoint to close,
                    // install, then relaunch PowerPoint.
                    for (int i = 0; i < 90 && IsPowerPointRunning(); i++) System.Threading.Thread.Sleep(1000);
                    if (IsPowerPointRunning())
                    {
                        MessageBox.Show("PowerPoint läuft noch. Bitte schließen – das Update wird anschließend installiert.",
                            AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        for (int i = 0; i < 180 && IsPowerPointRunning(); i++) System.Threading.Thread.Sleep(1000);
                    }
                    if (IsPowerPointRunning())
                    {
                        MessageBox.Show("PowerPoint konnte nicht geschlossen werden. Update abgebrochen.",
                            AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return 1;
                    }
                    System.Threading.Thread.Sleep(1500); // let file locks release
                    ExtractPayload();
                    Register();
                    ReopenPresentations();
                    MessageBox.Show("Update installiert.", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return 0;
                }

                if (IsPowerPointRunning())
                {
                    MessageBox.Show("Bitte PowerPoint schließen und das Setup erneut starten.",
                        AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 1;
                }

                if (uninstall) { Unregister(); RemoveFiles();
                    MessageBox.Show("\"" + AppName + "\" wurde entfernt.", AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return 0;
                }

                if (!WebView2Installed())
                {
                    var r = MessageBox.Show(
                        "Die Microsoft Edge WebView2 Runtime wird benötigt und scheint zu fehlen.\n\n" +
                        "Jetzt die Download-Seite öffnen? Nach der Installation dieses Setup erneut starten.",
                        AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (r == DialogResult.Yes)
                        Process.Start("https://developer.microsoft.com/microsoft-edge/webview2/");
                    return 1;
                }

                ExtractPayload();
                Register();

                MessageBox.Show("\"" + AppName + "\" wurde installiert.\n\nPowerPoint starten → Reiter \"Live Web\".",
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler: " + ex.Message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
        }

        private static bool IsPowerPointRunning() => Process.GetProcessesByName("POWERPNT").Length > 0;

        private static string ReopenListPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "LiveWebRegion", "reopen.txt");

        // Reopen the decks that were open before the update (paths recorded by the add-in).
        // ShellExecute routes each .pptx into a single reused PowerPoint instance.
        private static void ReopenPresentations()
        {
            string[] files = new string[0];
            try { if (File.Exists(ReopenListPath)) files = File.ReadAllLines(ReopenListPath); } catch { }
            try { if (File.Exists(ReopenListPath)) File.Delete(ReopenListPath); } catch { }

            var open = Array.FindAll(files, f => !string.IsNullOrWhiteSpace(f) && File.Exists(f));
            if (open.Length == 0)
            {
                string pp = FindPowerPoint();
                if (pp != null) { try { Process.Start(pp); } catch { } }
                return;
            }
            for (int i = 0; i < open.Length; i++)
            {
                try { Process.Start(new ProcessStartInfo(open[i]) { UseShellExecute = true }); } catch { }
                System.Threading.Thread.Sleep(2000); // let the first instance come up before the rest attach
            }
        }

        private static string FindPowerPoint()
        {
            try
            {
                using (var k = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"))
                    if (k?.GetValue(null) is string p && File.Exists(p)) return p;
            }
            catch { }
            string[] cands =
            {
                @"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
            };
            foreach (var c in cands) if (File.Exists(c)) return c;
            return null;
        }

        private static bool WebView2Installed()
        {
            string[] keys =
            {
                @"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                @"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
            };
            foreach (var k in keys)
            {
                using (var hklm = Registry.LocalMachine.OpenSubKey(k))
                    if (hklm?.GetValue("pv") is string pv && pv != "0.0.0.0") return true;
                using (var hkcu = Registry.CurrentUser.OpenSubKey(k))
                    if (hkcu?.GetValue("pv") is string pv && pv != "0.0.0.0") return true;
            }
            return false;
        }

        private static void ExtractPayload()
        {
            string parent = Path.GetDirectoryName(AppDir);
            Directory.CreateDirectory(parent);

            // 1) Extract into a fresh staging dir next to AppDir. This never touches the
            //    live files, so a still-loaded (locked) DLL cannot corrupt the install.
            string stage = AppDir + ".new";
            string old = AppDir + ".old";
            if (Directory.Exists(stage)) Directory.Delete(stage, true);
            Directory.CreateDirectory(stage);

            var asm = Assembly.GetExecutingAssembly();
            string resName = null;
            foreach (var n in asm.GetManifestResourceNames())
                if (n.EndsWith("payload.zip", StringComparison.OrdinalIgnoreCase)) { resName = n; break; }
            if (resName == null) throw new Exception("Eingebettetes Paket nicht gefunden.");

            string tmp = Path.Combine(Path.GetTempPath(), "LiveWebRegion_payload.zip");
            using (var rs = asm.GetManifestResourceStream(resName))
            using (var fs = File.Create(tmp))
                rs.CopyTo(fs);
            ZipFile.ExtractToDirectory(tmp, stage);
            try { File.Delete(tmp); } catch { }

            // 2) Swap staging into place with atomic directory renames. Retry while the
            //    old LiveWebRegion.dll is still image-locked (the CLR needs a moment to
            //    unload it after PowerPoint exits). A loaded DLL blocks the rename, so if
            //    every retry fails we abort WITHOUT having deleted the existing install.
            try { if (Directory.Exists(old)) Directory.Delete(old, true); } catch { }
            Exception last = null;
            for (int i = 0; i < 40; i++) // ~40 * 750 ms ≈ 30 s grace period
            {
                try
                {
                    if (Directory.Exists(AppDir)) Directory.Move(AppDir, old);
                    Directory.Move(stage, AppDir);
                    last = null;
                    break;
                }
                catch (Exception ex)
                {
                    last = ex;
                    // If the pre-move half-succeeded, restore so AppDir is never missing.
                    if (!Directory.Exists(AppDir) && Directory.Exists(old))
                        try { Directory.Move(old, AppDir); } catch { }
                    System.Threading.Thread.Sleep(750);
                }
            }
            if (last != null)
            {
                try { if (Directory.Exists(stage)) Directory.Delete(stage, true); } catch { }
                throw new Exception(
                    "Die Add-in-Dateien sind noch gesperrt – vermutlich läuft PowerPoint noch. " +
                    "Bitte ALLE PowerPoint-Fenster schließen (Änderungen speichern) und das Update erneut ausführen.", last);
            }
            try { if (Directory.Exists(old)) Directory.Delete(old, true); } catch { }
        }

        private static void RemoveFiles()
        {
            try { if (Directory.Exists(AppDir)) Directory.Delete(AppDir, true); } catch { }
        }

        private static void Register()
        {
            string dll = Path.Combine(AppDir, "LiveWebRegion.dll");
            string codeBase = new Uri(dll).AbsoluteUri;
            string mscoree = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "mscoree.dll");

            using (var clsid = Registry.CurrentUser.CreateSubKey(@"Software\Classes\CLSID\" + Clsid))
            {
                clsid.SetValue(null, ProgId);
                using (var inproc = clsid.CreateSubKey("InprocServer32"))
                {
                    inproc.SetValue(null, mscoree);
                    inproc.SetValue("ThreadingModel", "Both");
                    inproc.SetValue("Class", ProgId);
                    inproc.SetValue("Assembly", AssemblyFull);
                    inproc.SetValue("RuntimeVersion", "v4.0.30319");
                    inproc.SetValue("CodeBase", codeBase);
                    using (var ver = inproc.CreateSubKey("1.0.0.0"))
                    {
                        ver.SetValue("Class", ProgId);
                        ver.SetValue("Assembly", AssemblyFull);
                        ver.SetValue("RuntimeVersion", "v4.0.30319");
                        ver.SetValue("CodeBase", codeBase);
                    }
                }
                using (var pid = clsid.CreateSubKey("ProgId")) pid.SetValue(null, ProgId);
            }

            using (var prog = Registry.CurrentUser.CreateSubKey(@"Software\Classes\" + ProgId))
            {
                prog.SetValue(null, ProgId);
                using (var c = prog.CreateSubKey("CLSID")) c.SetValue(null, Clsid);
            }

            using (var addin = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId))
            {
                addin.SetValue("FriendlyName", AppName);
                addin.SetValue("Description", "Zeigt HTML/JS live in einem Folienbereich an.");
                addin.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                addin.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            }

            ClearDisabledItems();
        }

        private static void Unregister()
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\CLSID\" + Clsid, false); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\" + ProgId, false); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId, false); } catch { }
        }

        // PowerPoint hard-disables an add-in after a crash; clear that so it loads.
        private static void ClearDisabledItems()
        {
            foreach (var ver in new[] { "16.0", "15.0" })
            {
                string path = $@"Software\Microsoft\Office\{ver}\PowerPoint\Resiliency\DisabledItems";
                using (var key = Registry.CurrentUser.OpenSubKey(path, writable: true))
                {
                    if (key == null) continue;
                    foreach (var name in key.GetValueNames())
                    {
                        if (key.GetValue(name) is byte[] b)
                        {
                            string txt = System.Text.Encoding.Unicode.GetString(b).ToLowerInvariant();
                            if (txt.Contains("livewebregion")) { try { key.DeleteValue(name); } catch { } }
                        }
                    }
                }
            }
        }
    }
}
