using System;
using System.IO;

namespace LiveWebRegion
{
    /// <summary>
    /// Defines/queries "web regions": shapes tagged with the HTML file they should
    /// display. All access is late-bound (dynamic) against the PowerPoint COM model.
    /// </summary>
    internal static class ShapeRegions
    {
        public const string TagName = "LiveWebPath";

        private const int ppSelectionShapes = 2;

        /// <summary>Returns the first selected shape, or null if none is selected.</summary>
        public static dynamic GetSelectedShape(dynamic app)
        {
            try
            {
                dynamic win = app.ActiveWindow;
                dynamic sel = win.Selection;
                if ((int)sel.Type == ppSelectionShapes)
                {
                    dynamic range = sel.ShapeRange;
                    if ((int)range.Count >= 1)
                        return range.Item(1);
                }
            }
            catch (Exception ex) { Log.Error("GetSelectedShape failed", ex); }
            return null;
        }

        private const int msoShapeRectangle = 1;

        /// <summary>Inserts a centered rectangle on the current edit slide and selects it.</summary>
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

        /// <summary>HTML path stored on a shape, or "" when it is not a web region.</summary>
        public static string GetPath(dynamic shape)
        {
            try { return (string)shape.Tags.Item(TagName) ?? ""; }
            catch { return ""; }
        }

        public static bool IsRegion(dynamic shape)
        {
            return shape != null && !string.IsNullOrEmpty(GetPath(shape));
        }

        public static bool IsHttpUrl(string v)
        {
            return v != null &&
                   (v.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    v.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Prompts for a URL or local HTML file. Returns null on cancel.</summary>
        public static string PickLink(string initial = null)
        {
            try
            {
                using (var dlg = new LinkInputDialog(initial))
                {
                    if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return null;
                    string v = dlg.Value;
                    if (string.IsNullOrWhiteSpace(v)) return null;
                    return Normalize(v);
                }
            }
            catch (Exception ex) { Log.Error("PickLink failed", ex); return null; }
        }

        // Accept "example.com/foo" by assuming https; leave files and full URLs as-is.
        private static string Normalize(string v)
        {
            if (IsHttpUrl(v)) return v;
            if (File.Exists(v)) return v;
            if (v.Contains(".") && !v.Contains("\\") && !v.Contains(":"))
                return "https://" + v;
            return v;
        }

        /// <summary>Stores the path on the shape and marks it visually in edit mode.</summary>
        public static void SetRegion(dynamic shape, string path)
        {
            shape.Tags.Add(TagName, path);
            try { shape.AlternativeText = "LiveWebRegion: " + path; } catch { }
            ApplyMarker(shape, path);
        }

        public static void ClearRegion(dynamic shape)
        {
            try { shape.Tags.Delete(TagName); } catch { }
            try { shape.AlternativeText = ""; } catch { }
            try { shape.TextFrame.TextRange.Text = ""; } catch { }
        }

        // Light fill + dashed border + label so web regions are recognisable while editing.
        private static void ApplyMarker(dynamic shape, string path)
        {
            string label = path;
            try { label = IsHttpUrl(path) ? new Uri(path).Host : Path.GetFileName(path); } catch { }
            try
            {
                shape.TextFrame.TextRange.Text = "▶ Live Web\n" + label;
            }
            catch { }
            try
            {
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = Rgb(235, 245, 255); // very light blue
                shape.Fill.Transparency = 0.15f;
            }
            catch { }
            try
            {
                shape.Line.Visible = -1;
                shape.Line.DashStyle = 4;            // msoLineDash
                shape.Line.ForeColor.RGB = Rgb(80, 140, 220);
                shape.Line.Weight = 1.5f;
            }
            catch { }
        }

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
    }
}
