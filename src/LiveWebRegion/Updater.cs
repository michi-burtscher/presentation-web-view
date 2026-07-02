using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LiveWebRegion
{
    /// <summary>Checks GitHub for a newer release and installs it via the Setup EXE.</summary>
    internal static class Updater
    {
        private sealed class Manifest
        {
            public string Version = "";
            public string Notes = "";
            public string ExeUrl = "";
            public string Sha256 = "";
        }

        // Startup: fetch quietly on a background thread; only inform if newer.
        public static void NotifyIfUpdateAvailable()
        {
            try
            {
                Manifest m = Fetch();
                if (m != null && IsNewer(m.Version, AppInfo.Version))
                {
                    Native.Info("Eine neue Version ist verfügbar: " + m.Version +
                                " (installiert: " + AppInfo.Version + ").\n\n" +
                                "Ribbon „Live Web“ → „Update“ installiert sie.", "Live Web – Update");
                }
            }
            catch (Exception ex) { Log.Error("NotifyIfUpdateAvailable failed", ex); }
        }

        // Ribbon button: check, download, and install (must run on the UI thread).
        public static void CheckAndInstall(dynamic app)
        {
            try
            {
                Manifest m = Fetch();
                if (m == null || string.IsNullOrEmpty(m.Version))
                {
                    Native.Warn("Es konnte kein gültiges Update-Manifest geladen werden.");
                    return;
                }
                if (!IsNewer(m.Version, AppInfo.Version))
                {
                    Native.Info("Sie verwenden bereits die aktuelle Version (" + AppInfo.Version + ").");
                    return;
                }

                string msg = "Neue Version verfügbar: " + m.Version + "  (installiert: " + AppInfo.Version + ")\n\n" +
                             m.Notes + "\n\nJetzt herunterladen und installieren?";
                if (MessageBox.Show(msg, "Live Web – Update", MessageBoxButtons.YesNo, MessageBoxIcon.Information) != DialogResult.Yes)
                    return;

                if (string.IsNullOrEmpty(m.ExeUrl)) { Native.Warn("Im Manifest fehlt die Download-Adresse."); return; }

                string tmp = Path.Combine(Path.GetTempPath(), "LiveWebRegionSetup.exe");
                using (var wc = NewClient()) wc.DownloadFile(m.ExeUrl, tmp);

                if (!string.IsNullOrEmpty(m.Sha256) && !VerifySha256(tmp, m.Sha256))
                {
                    Native.Warn("Prüfsumme der heruntergeladenen Datei stimmt nicht – Update abgebrochen.");
                    try { File.Delete(tmp); } catch { }
                    return;
                }

                if (MessageBox.Show(
                        "Das Update wird jetzt installiert. PowerPoint wird dazu geschlossen und anschließend neu gestartet.\n\n" +
                        "Bitte offene Präsentationen speichern. Fortfahren?",
                        "Live Web – Update", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK)
                    return;

                // Save open decks (so Quit isn't blocked by "save?" dialogs) and record
                // their paths so the installer can reopen them after the update.
                PrepareCloseAndReopen(app);

                Process.Start(new ProcessStartInfo(tmp, "/update") { UseShellExecute = true });
                try { app.Quit(); } catch (Exception ex) { Log.Error("app.Quit for update failed", ex); }
            }
            catch (Exception ex)
            {
                Log.Error("CheckAndInstall failed", ex);
                Native.Warn("Update fehlgeschlagen: " + ex.Message);
            }
        }

        // File the add-in writes and the installer reads to reopen decks post-update.
        internal static string ReopenListPath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "LiveWebRegion", "reopen.txt");

        // Save unsaved-but-on-disk presentations (silent; the user already confirmed the
        // restart) and record all on-disk deck paths for the installer to reopen.
        private static void PrepareCloseAndReopen(dynamic app)
        {
            var paths = new System.Collections.Generic.List<string>();
            try
            {
                dynamic pres = app.Presentations;
                int n = (int)pres.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic p = pres.Item(i);
                    string dir = ""; try { dir = (string)p.Path; } catch { } // empty => never saved
                    if (string.IsNullOrEmpty(dir)) continue;                  // can't reopen unsaved decks
                    try { if ((int)p.Saved == 0) p.Save(); }                  // msoFalse(0) => dirty
                    catch (Exception ex) { Log.Error("Save before update failed", ex); }
                    try { string full = (string)p.FullName; if (!string.IsNullOrEmpty(full)) paths.Add(full); } catch { }
                }
            }
            catch (Exception ex) { Log.Error("Enumerate presentations failed", ex); }

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(ReopenListPath));
                File.WriteAllLines(ReopenListPath, paths);
            }
            catch (Exception ex) { Log.Error("Write reopen list failed", ex); }
        }

        private static Manifest Fetch()
        {
            string json;
            using (var wc = NewClient()) json = wc.DownloadString(AppInfo.ManifestUrl);
            return new Manifest
            {
                Version = Str(json, "version"),
                ExeUrl = Str(json, "exe_url"),
                Sha256 = Str(json, "sha256"),
                Notes = Notes(json)
            };
        }

        private static WebClient NewClient()
        {
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }
            var wc = new WebClient();
            wc.Headers.Add("User-Agent", "LiveWebRegion/" + AppInfo.Version);
            return wc;
        }

        private static string Str(string json, string key)
        {
            var m = Regex.Match(json, "\"" + key + "\"\\s*:\\s*\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string Notes(string json)
        {
            var m = Regex.Match(json, "\"notes\"\\s*:\\s*\\[(.*?)\\]", RegexOptions.Singleline);
            if (!m.Success) return Str(json, "notes");
            var lines = Regex.Matches(m.Groups[1].Value, "\"([^\"]*)\"");
            var sb = new System.Text.StringBuilder();
            foreach (Match l in lines) sb.AppendLine("• " + l.Groups[1].Value);
            return sb.ToString().TrimEnd();
        }

        private static bool IsNewer(string candidate, string current)
        {
            return VersionKey(candidate) > VersionKey(current);
        }

        // First run of digits -> integer (v01 -> 1, v12 -> 12).
        private static int VersionKey(string v)
        {
            var m = Regex.Match(v ?? "", "\\d+");
            return m.Success && int.TryParse(m.Value, out int n) ? n : 0;
        }

        private static bool VerifySha256(string file, string expected)
        {
            try
            {
                using (var sha = SHA256.Create())
                using (var fs = File.OpenRead(file))
                {
                    string actual = BitConverter.ToString(sha.ComputeHash(fs)).Replace("-", "").ToLowerInvariant();
                    return actual == expected.Trim().ToLowerInvariant();
                }
            }
            catch { return false; }
        }
    }
}
