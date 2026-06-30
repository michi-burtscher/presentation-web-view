using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace LiveWebRegion
{
    /// <summary>
    /// PowerPoint COM add-in entry point: lifecycle (IDTExtensibility2) and Ribbon
    /// UI (IRibbonExtensibility).
    /// </summary>
    [ComVisible(true)]
    [Guid(ClsidString)]
    [ProgId(ProgIdString)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        public const string ClsidString = "7E9B2C14-3A6D-4F58-9C1E-2B7A5D0F8E31";
        public const string ProgIdString = "LiveWebRegion.AddIn";

        private dynamic _app;     // PowerPoint.Application (late-bound)
        private dynamic _ribbon;  // Office.IRibbonUI (late-bound)
        private OverlayManager _overlays;

        #region IDTExtensibility2

        public void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, IntPtr custom)
        {
            try
            {
                if (Application != IntPtr.Zero)
                    _app = Marshal.GetObjectForIUnknown(Application);
                Log.Info("OnConnection. mode=" + ConnectMode + " host=" + SafeProbe());
            }
            catch (Exception ex) { Log.Error("OnConnection failed", ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, IntPtr custom)
        {
            try { _overlays?.Stop(); } catch { }
            _overlays = null;
            _ribbon = null;
            _app = null;
            Log.Info("OnDisconnection.");
        }

        public void OnAddInsUpdate(IntPtr custom) { }

        public void OnStartupComplete(IntPtr custom)
        {
            try
            {
                _overlays = new OverlayManager(_app);
                _overlays.Start();
            }
            catch (Exception ex) { Log.Error("OverlayManager start failed", ex); }
        }

        public void OnBeginShutdown(IntPtr custom)
        {
            try { _overlays?.Stop(); } catch { }
        }

        private string SafeProbe()
        {
            try { return _app != null ? (string)_app.Name + " " + (string)_app.Version : "<null>"; }
            catch { return "<probe failed>"; }
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                using (Stream s = Assembly.GetExecutingAssembly()
                           .GetManifestResourceStream("LiveWebRegion.Ribbon.xml"))
                {
                    if (s == null) { Log.Error("Ribbon.xml resource missing."); return null; }
                    using (var reader = new StreamReader(s))
                        return reader.ReadToEnd();
                }
            }
            catch (Exception ex) { Log.Error("GetCustomUI failed", ex); return null; }
        }

        #endregion

        #region Ribbon callbacks (params taken as object; Office invokes via IDispatch)

        public void OnRibbonLoad(object ribbon)
        {
            _ribbon = ribbon;
        }

        public void OnSetRegion(object control)
        {
            try
            {
                bool inserted = false;
                dynamic shape = ShapeRegions.GetSelectedShape(_app);
                if (shape == null)
                {
                    // No shape selected: insert a rectangle for the user, then continue.
                    shape = ShapeRegions.InsertRegionShape(_app);
                    inserted = true;
                    if (shape == null)
                    {
                        Native.Warn("Es konnte keine Form eingef&#252;gt werden. Bitte eine Form markieren und erneut versuchen.");
                        return;
                    }
                }

                string path = ShapeRegions.PickLink();
                if (string.IsNullOrEmpty(path))
                {
                    if (inserted) { try { shape.Delete(); } catch { } } // undo the auto-insert on cancel
                    return;
                }

                ShapeRegions.SetRegion(shape, path);
                Log.Info("Region set on shape -> " + path);
                Invalidate();
            }
            catch (Exception ex) { Log.Error("OnSetRegion failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        public void OnChangeFile(object control)
        {
            try
            {
                dynamic shape = ShapeRegions.GetSelectedShape(_app);
                if (!ShapeRegions.IsRegion(shape))
                {
                    Native.Warn("Die markierte Form ist kein Live-Web-Bereich.");
                    return;
                }
                string current = ShapeRegions.GetPath(shape);
                string path = ShapeRegions.PickLink(current);
                if (string.IsNullOrEmpty(path)) return;
                ShapeRegions.SetRegion(shape, path);
                Log.Info("Region file changed -> " + path);
            }
            catch (Exception ex) { Log.Error("OnChangeFile failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        public void OnRemoveRegion(object control)
        {
            try
            {
                dynamic shape = ShapeRegions.GetSelectedShape(_app);
                if (!ShapeRegions.IsRegion(shape))
                {
                    Native.Warn("Die markierte Form ist kein Live-Web-Bereich.");
                    return;
                }
                ShapeRegions.ClearRegion(shape);
                Log.Info("Region removed.");
                Invalidate();
            }
            catch (Exception ex) { Log.Error("OnRemoveRegion failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        public void OnReload(object control)
        {
            try { _overlays?.ForceRefresh(); Log.Info("Overlays refresh requested."); }
            catch (Exception ex) { Log.Error("OnReload failed", ex); }
        }

        public bool OnGetRegionSelected(object control)
        {
            try { return ShapeRegions.IsRegion(ShapeRegions.GetSelectedShape(_app)); }
            catch { return false; }
        }

        public void OnShowHelp(object control) { OpenAsset("help.html"); }
        public void OnShowAbout(object control) { OpenAsset("about.html"); }

        private void OpenAsset(string fileName)
        {
            try
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string path = Path.Combine(dir, "assets", fileName);
                if (File.Exists(path))
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                else
                    Native.Warn("Datei nicht gefunden: " + path);
            }
            catch (Exception ex) { Log.Error("OpenAsset(" + fileName + ") failed", ex); }
        }

        private void Invalidate()
        {
            try { if (_ribbon != null) _ribbon.Invalidate(); } catch { }
        }

        #endregion
    }
}
