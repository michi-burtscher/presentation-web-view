using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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
        private Control _ui;      // UI-thread invoker (deferred dialogs)
        private bool _previewEnabled;

        #region IDTExtensibility2

        public void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, IntPtr custom)
        {
            try
            {
                if (Application != IntPtr.Zero)
                    _app = Marshal.GetObjectForIUnknown(Application);
                Log.Info("OnConnection. mode=" + ConnectMode + " host=" + SafeProbe());

                _ui = new Control();
                var _ = _ui.Handle; // realize handle on the UI thread
            }
            catch (Exception ex) { Log.Error("OnConnection failed", ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, IntPtr custom)
        {
            try { _overlays?.Stop(); } catch { }
            _overlays = null;
            try { _ui?.Dispose(); } catch { }
            _ui = null;
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
                _overlays.PreviewClosedByUser = OnPreviewClosedByUser;
                _overlays.Start();
                // Quiet background check; only informs if a newer version exists.
                System.Threading.Tasks.Task.Run(() => Updater.NotifyIfUpdateAvailable());
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

        public void OnRibbonLoad(object ribbon) { _ribbon = ribbon; }

        // "Fenster erstellen": always insert a new frame, then configure it.
        public void OnCreateFrame(object control)
        {
            try
            {
                dynamic shape = ShapeRegions.InsertRegionShape(_app);
                if (shape == null) { Native.Warn("Es konnte kein Frame eingefügt werden. Ist eine Folie geöffnet?"); return; }

                LinkResult res = ShapeRegions.EditRegion(shape);
                if (res == null) { try { shape.Delete(); } catch { } return; }

                ShapeRegions.SetRegion(shape, res);
                Log.Info("Frame created -> " + res.Value + (res.Embed ? " (embedded)" : ""));
                Invalidate();
            }
            catch (Exception ex) { Log.Error("OnCreateFrame failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        // "Optionen": configure the selected frame.
        public void OnOptions(object control)
        {
            try
            {
                dynamic shape = ShapeRegions.GetSelectedShape(_app);
                if (!ShapeRegions.IsRegion(shape)) { Native.Warn("Bitte zuerst einen Live Web Frame markieren."); return; }
                EditRegionShape(shape);
            }
            catch (Exception ex) { Log.Error("OnOptions failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        public void OnRemoveFrame(object control)
        {
            try
            {
                dynamic shape = ShapeRegions.GetSelectedShape(_app);
                if (!ShapeRegions.IsRegion(shape)) { Native.Warn("Die markierte Form ist kein Live Web Frame."); return; }
                ShapeRegions.ClearRegion(shape);
                Log.Info("Frame removed.");
                Invalidate();
            }
            catch (Exception ex) { Log.Error("OnRemoveFrame failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        public void OnReload(object control)
        {
            try { _overlays?.ForceRefresh(); Log.Info("Overlays refresh requested."); }
            catch (Exception ex) { Log.Error("OnReload failed", ex); }
        }

        public void OnTogglePreview(object control, bool pressed)
        {
            _previewEnabled = pressed;
            try { _overlays?.SetPreview(pressed); } catch (Exception ex) { Log.Error("OnTogglePreview failed", ex); }
        }
        public bool OnGetPreviewPressed(object control) { return _previewEnabled; }

        public bool OnGetRegionSelected(object control)
        {
            try { return ShapeRegions.IsRegion(ShapeRegions.GetSelectedShape(_app)); }
            catch { return false; }
        }

        private void EditRegionShape(dynamic shape)
        {
            try
            {
                LinkResult res = ShapeRegions.EditRegion(shape);
                if (res != null) { ShapeRegions.SetRegion(shape, res); Invalidate(); }
            }
            catch (Exception ex) { Log.Error("EditRegionShape failed", ex); Native.Warn("Fehler: " + ex.Message); }
        }

        private void OnPreviewClosedByUser()
        {
            _previewEnabled = false;
            Invalidate(); // un-press the ribbon toggle
        }

        public void OnCheckUpdate(object control)
        {
            try { Updater.CheckAndInstall(_app); }
            catch (Exception ex) { Log.Error("OnCheckUpdate failed", ex); }
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
