using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LiveWebRegion
{
    /// <summary>
    /// Defines/queries "web regions": shapes tagged with the website they should show.
    /// Storage (shape tags): LiveWebPath = URL / relative / absolute / embedded display
    /// name; LiveWebData = base64 of an embedded HTML file; LiveWebOpts = options.
    /// All access is late-bound (dynamic) against the PowerPoint COM model.
    /// </summary>
    internal static class ShapeRegions
    {
        public const string TagName = "LiveWebPath";
        public const string TagData = "LiveWebData";
        public const string TagOpts = "LiveWebOpts";

        private const int ppSelectionShapes = 2;
        private const int msoShapeRectangle = 1;

        #region selection / insertion

        public static dynamic GetSelectedShape(dynamic app)
        {
            try
            {
                dynamic win = app.ActiveWindow;
                dynamic sel = win.Selection;
                if ((int)sel.Type == ppSelectionShapes)
                {
                    dynamic range = sel.ShapeRange;
                    if ((int)range.Count >= 1) return range.Item(1);
                }
            }
            catch (Exception ex) { Log.Error("GetSelectedShape failed", ex); }
            return null;
        }

        public static dynamic InsertRegionShape(dynamic app)
        {
            try
            {
                dynamic slide = app.ActiveWindow.View.Slide;
                dynamic pres = slide.Parent;
                float sw = (float)pres.PageSetup.SlideWidth;
                float sh = (float)pres.PageSetup.SlideHeight;
                float w = sw * 0.6f, h = sh * 0.6f;
                dynamic shape = slide.Shapes.AddShape(msoShapeRectangle, (sw - w) / 2f, (sh - h) / 2f, w, h);
                try { shape.Select(); } catch { }
                return shape;
            }
            catch (Exception ex) { Log.Error("InsertRegionShape failed", ex); return null; }
        }

        #endregion

        #region queries

        public static bool IsHttpUrl(string v)
        {
            return v != null &&
                   (v.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    v.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
        }

        private static string Tag(dynamic shape, string name)
        {
            try { return (string)shape.Tags.Item(name) ?? ""; }
            catch { return ""; }
        }

        /// <summary>The stored link value (display string); "" if not a region.</summary>
        public static string GetPath(dynamic shape) => Tag(shape, TagName);

        public static bool IsEmbedded(dynamic shape) => !string.IsNullOrEmpty(Tag(shape, TagData));

        public static bool IsRegion(dynamic shape)
        {
            return shape != null && (!string.IsNullOrEmpty(Tag(shape, TagName)) || IsEmbedded(shape));
        }

        public static RegionOptions GetOptions(dynamic shape) => RegionOptions.Parse(Tag(shape, TagOpts));

        #endregion

        #region set / clear

        private const int MaxEmbedBytes = 2 * 1024 * 1024; // safe shape-tag budget

        // One-click installer download (always the latest published release asset).
        private const string DownloadUrl =
            "https://github.com/michi-burtscher/presentation-web-view/releases/latest/download/LiveWebRegionSetup.exe";

        public static void SetRegion(dynamic shape, LinkResult r)
        {
            string value = r.Value ?? "";
            bool wasEmbedded = IsEmbedded(shape);
            string prevPath = GetPath(shape);

            if (IsHttpUrl(value))
            {
                try { shape.Tags.Delete(TagData); } catch { }
                shape.Tags.Add(TagName, value);
            }
            else if (r.Embed && File.Exists(value))
            {
                byte[] bytes = File.ReadAllBytes(value);
                if (bytes.Length > MaxEmbedBytes)
                {
                    Native.Warn("Die Datei ist zu groß zum Einbetten (max. 2 MB). Es wird der Pfad gespeichert.");
                    try { shape.Tags.Delete(TagData); } catch { }
                    shape.Tags.Add(TagName, StoreValue(shape, value));
                }
                else
                {
                    shape.Tags.Add(TagData, Convert.ToBase64String(bytes));
                    shape.Tags.Add(TagName, StoreValue(shape, value)); // keep path for re-edit
                }
            }
            else if (r.Embed && wasEmbedded && SameLocalPath(shape, value, prevPath))
            {
                // File no longer available but the embedded bytes are already stored: keep them.
                shape.Tags.Add(TagName, prevPath);
            }
            else
            {
                try { shape.Tags.Delete(TagData); } catch { }
                shape.Tags.Add(TagName, StoreValue(shape, value));
            }

            shape.Tags.Add(TagOpts, r.Options.ToString());
            try { shape.AlternativeText = "LiveWebRegion: " + value; } catch { }
            ApplyMarker(shape, value, IsEmbedded(shape));
        }

        private static bool SameLocalPath(dynamic shape, string a, string b)
        {
            try { return string.Equals(ResolveToAbsolute(shape, a), ResolveToAbsolute(shape, b), StringComparison.OrdinalIgnoreCase); }
            catch { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        }

        public static void ClearRegion(dynamic shape)
        {
            foreach (var t in new[] { TagName, TagData, TagOpts })
            {
                try { shape.Tags.Delete(t); } catch { }
            }
            try { shape.AlternativeText = ""; } catch { }
            try { shape.TextFrame.TextRange.Text = ""; } catch { }
            try { shape.ActionSettings.Item(1).Hyperlink.Delete(); } catch { }
        }

        // Prefer a path relative to the presentation so the deck stays portable when
        // the HTML travels next to it.
        private static string StoreValue(dynamic shape, string value)
        {
            if (IsHttpUrl(value)) return value;
            if (!File.Exists(value)) return value;
            try
            {
                string presFolder = (string)shape.Parent.Parent.Path; // slide -> presentation
                if (!string.IsNullOrEmpty(presFolder))
                {
                    var baseUri = new Uri(AppendSlash(presFolder));
                    var fileUri = new Uri(value);
                    if (baseUri.IsBaseOf(fileUri) ||
                        string.Equals(Path.GetPathRoot(presFolder), Path.GetPathRoot(value), StringComparison.OrdinalIgnoreCase))
                    {
                        string rel = Uri.UnescapeDataString(baseUri.MakeRelativeUri(fileUri).ToString()).Replace('/', '\\');
                        if (!string.IsNullOrEmpty(rel) && !Path.IsPathRooted(rel)) return rel;
                    }
                }
            }
            catch (Exception ex) { Log.Error("StoreValue relative failed", ex); }
            return value; // fall back to absolute
        }

        #endregion

        #region resolve to a navigable URL

        /// <summary>Resolves a region to the URL/file the WebView2 should navigate to.</summary>
        public static string ResolveTarget(dynamic shape)
        {
            string data = Tag(shape, TagData);
            if (!string.IsNullOrEmpty(data))
            {
                try
                {
                    byte[] bytes = Convert.FromBase64String(data);
                    string file = WriteEmbedTemp(bytes);
                    return new Uri(file).AbsoluteUri;
                }
                catch (Exception ex) { Log.Error("ResolveTarget embed failed", ex); return null; }
            }

            string value = Tag(shape, TagName);
            if (string.IsNullOrEmpty(value)) return null;
            if (IsHttpUrl(value)) return value;

            string path = ResolveToAbsolute(shape, value);
            if (!File.Exists(path)) return null;
            return new Uri(path).AbsoluteUri;
        }

        // Resolve a stored (possibly relative) local path against the presentation folder.
        private static string ResolveToAbsolute(dynamic shape, string path)
        {
            if (string.IsNullOrEmpty(path) || Path.IsPathRooted(path)) return path;
            try
            {
                string presFolder = (string)shape.Parent.Parent.Path;
                if (!string.IsNullOrEmpty(presFolder))
                    return Path.GetFullPath(Path.Combine(presFolder, path));
            }
            catch { }
            return path;
        }

        // Value to pre-fill the dialog with: absolute path for local files (so a round-trip
        // through Normalize/StoreValue is lossless), or the URL as-is.
        private static string ResolveForEditing(dynamic shape)
        {
            string v = GetPath(shape);
            if (string.IsNullOrEmpty(v) || IsHttpUrl(v)) return v;
            return ResolveToAbsolute(shape, v);
        }

        private static string WriteEmbedTemp(byte[] bytes)
        {
            string dir = Path.Combine(Path.GetTempPath(), "LiveWebRegion", "embed");
            Directory.CreateDirectory(dir);
            string hash;
            using (var sha = SHA1.Create())
                hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
            string file = Path.Combine(dir, hash + ".html");
            if (!File.Exists(file)) File.WriteAllBytes(file, bytes);
            return file;
        }

        #endregion

        #region edit-mode marker

        private static string AppendSlash(string folder)
        {
            return folder.EndsWith("\\") || folder.EndsWith("/") ? folder : folder + "\\";
        }

        private static void ApplyMarker(dynamic shape, string value, bool embedded)
        {
            string label;
            if (string.IsNullOrEmpty(value)) label = "Noch kein Link gesetzt";
            else if (embedded) label = Path.GetFileName(value) + " (eingebettet)";
            else if (IsHttpUrl(value)) { try { label = new Uri(value).Host; } catch { label = value; } }
            else label = Path.GetFileName(value);

            // Self-describing card: title, link/host, and a hint that also helps a
            // recipient who does not have the add-in installed.
            try
            {
                dynamic tr = shape.TextFrame.TextRange;
                tr.Text = "🌐  Live Web\n" + label +
                          "\nBenötigt das Add-in „Live Web Region“ für die Live-Ansicht." +
                          "\nDownload: github.com/michi-burtscher/presentation-web-view/releases" +
                          "\n(im Präsentationsmodus genügt ein Klick auf dieses Feld)";
                try { tr.ParagraphFormat.Alignment = 2; } catch { } // ppAlignCenter
                try { tr.Font.Color.RGB = Rgb(30, 58, 138); } catch { }
                try { dynamic p1 = tr.Paragraphs(1, 1); p1.Font.Bold = -1; p1.Font.Size = 18; } catch { }
                try { dynamic p3 = tr.Paragraphs(3, 3); p3.Font.Size = 10; p3.Font.Color.RGB = Rgb(120, 130, 150); } catch { }
                try { dynamic pUrl = tr.Paragraphs(4, 1); pUrl.Font.Bold = -1; pUrl.Font.Color.RGB = Rgb(80, 140, 220); } catch { }
            }
            catch { }

            // Clickable in slideshow for recipients without the add-in: opens the
            // installer download. (With the add-in, the WebView2 overlay covers it.)
            try { shape.ActionSettings.Item(1).Hyperlink.Address = DownloadUrl; } catch { }

            try
            {
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = Rgb(238, 246, 255);
                shape.Fill.Transparency = 0.10f;
            }
            catch { }
            try
            {
                shape.Line.Visible = -1;
                shape.Line.DashStyle = 4; // msoLineDash
                shape.Line.ForeColor.RGB = Rgb(80, 140, 220);
                shape.Line.Weight = 1.25f;
            }
            catch { }
        }

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);

        #endregion

        #region dialog

        /// <summary>Shows the link/options dialog. Returns null on cancel.</summary>
        public static LinkResult EditRegion(dynamic shape)
        {
            string initial = null;
            bool embed = false;
            RegionOptions opts = new RegionOptions();
            if (shape != null && IsRegion(shape))
            {
                initial = ResolveForEditing(shape);
                embed = IsEmbedded(shape);
                opts = GetOptions(shape);
            }
            using (var dlg = new LinkInputDialog(initial, embed, opts))
            {
                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return null;
                string v = dlg.Value;
                if (string.IsNullOrWhiteSpace(v)) return null;
                return new LinkResult { Value = Normalize(v), Embed = dlg.Embed, Options = dlg.Options };
            }
        }

        // Accept "example.com/foo" by assuming https; leave file paths and full URLs as-is.
        private static string Normalize(string v)
        {
            if (IsHttpUrl(v)) return v;
            if (File.Exists(v)) return v;
            bool looksLikeFile = v.Contains("\\") || v.Contains(":") ||
                                 v.EndsWith(".htm", StringComparison.OrdinalIgnoreCase) ||
                                 v.EndsWith(".html", StringComparison.OrdinalIgnoreCase);
            if (!looksLikeFile && v.Contains(".")) return "https://" + v;
            return v;
        }

        #endregion
    }
}
