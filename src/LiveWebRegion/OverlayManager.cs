using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace LiveWebRegion
{
    /// <summary>
    /// Detects an active slide show by polling (COM events can't be bound through
    /// the late-bound `dynamic` Application) and keeps a WebView2 overlay positioned
    /// over every tagged web region on the current slide.
    /// </summary>
    internal sealed class OverlayManager
    {
        private readonly dynamic _app;
        private Timer _timer;
        private CoreWebView2Environment _env;
        private readonly List<WebOverlayForm> _overlays = new List<WebOverlayForm>();

        private bool _active;
        private int _builtPos = -1;
        private Control _invoker;   // deferred UI-thread actions

        private const int PollMs = 200;

        public OverlayManager(dynamic app) { _app = app; }

        public void Start()
        {
            _invoker = new Control();
            var _ = _invoker.Handle; // realize handle on the UI thread

            string udf = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LiveWebRegion", "WebView2");
            // Shared environment so all overlays can share one user-data folder.
            CoreWebView2Environment.CreateAsync(null, udf, null)
                .ContinueWith(t =>
                {
                    if (t.IsFaulted) Log.Error("WebView2 env failed", t.Exception);
                    else _env = t.Result;
                });

            _timer = new Timer { Interval = PollMs };
            _timer.Tick += (s, e) => Poll();
            _timer.Start();
            Log.Info("OverlayManager started.");
        }

        public void Stop()
        {
            try { _timer?.Stop(); } catch { }
            CloseAllOverlays();
            try { _invoker?.Dispose(); } catch { }
        }

        // Navigation requested from inside an overlay (WebView2 AcceleratorKeyPressed
        // fires while the page has focus; PowerPoint handles keys natively otherwise).
        private void SlideShow(string action)
        {
            try
            {
                if ((int)_app.SlideShowWindows.Count == 0) return;
                dynamic view = _app.SlideShowWindows.Item(1).View;
                switch (action)
                {
                    case "Next": view.Next(); break;
                    case "Previous": view.Previous(); break;
                    case "Exit":
                        // Hide instantly (we're still on the UI thread, before PowerPoint
                        // blocks it tearing the show down), then exit and dispose later.
                        HideAllOverlays();
                        view.Exit();
                        try { _invoker?.BeginInvoke((Action)EndShow); } catch { }
                        break;
                }
            }
            catch (Exception ex) { Log.Error("SlideShow(" + action + ") failed", ex); }
        }

        /// <summary>Force a rebuild on the next poll (e.g. the "Reload" button).</summary>
        public void ForceRefresh()
        {
            _builtPos = -1;
            foreach (var o in _overlays) o.Reload();
        }

        private void Poll()
        {
            try
            {
                int count = 0;
                try { count = (int)_app.SlideShowWindows.Count; } catch { count = 0; }

                if (count == 0)
                {
                    if (_active) EndShow();
                    return;
                }

                _active = true;
                dynamic view = _app.SlideShowWindows.Item(1).View;
                int pos = (int)view.CurrentShowPosition;

                if (_env != null && pos != _builtPos)
                {
                    if (RebuildForSlide(view)) _builtPos = pos;
                }
            }
            catch (Exception ex) { Log.Error("Poll failed", ex); }
        }

        // Returns true once the slide was processed (so we don't rebuild again);
        // false on a transient state (no slide yet, no show window) to retry.
        private bool RebuildForSlide(dynamic view)
        {
            var windows = Native.GetSlideShowWindows();
            if (windows.Count == 0) return false; // show window not up yet
            Native.ShowWindow show = PickAudience(windows);
            Rectangle winPx = show.Rect;

            dynamic slide;
            try { slide = view.Slide; }
            catch { return false; } // "no slide currently in view" during transitions

            CloseAllOverlays();
            try
            {
                dynamic pres = slide.Parent;
                float slideW = (float)pres.PageSetup.SlideWidth;
                float slideH = (float)pres.PageSetup.SlideHeight;

                // The slide is scaled to fit the show window and centered (letterboxed).
                float s = Math.Min(winPx.Width / slideW, winPx.Height / slideH);
                float renderW = slideW * s, renderH = slideH * s;
                float offX = winPx.Left + (winPx.Width - renderW) / 2f;
                float offY = winPx.Top + (winPx.Height - renderH) / 2f;

                dynamic shapes = slide.Shapes;
                int n = (int)shapes.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic sh = shapes.Item(i);
                    if (!ShapeRegions.IsRegion(sh)) continue;
                    string target = ShapeRegions.ResolveTarget(sh);
                    if (string.IsNullOrEmpty(target)) { Log.Error("Region target unresolved on slide " + _builtPos); continue; }
                    RegionOptions opts = ShapeRegions.GetOptions(sh);

                    int x = (int)Math.Round(offX + (float)sh.Left * s);
                    int y = (int)Math.Round(offY + (float)sh.Top * s);
                    int w = (int)Math.Round((float)sh.Width * s);
                    int h = (int)Math.Round((float)sh.Height * s);
                    var rect = new Rectangle(x, y, w, h);
                    // Grow slightly so the overlay fully covers the shape's edit-mode
                    // marker outline (dashed border straddles the edge).
                    rect.Inflate((int)Math.Ceiling(s) + 1, (int)Math.Ceiling(s) + 1);

                    // Own the overlay to the show window so Windows destroys it the
                    // instant PowerPoint closes the show (Esc), without waiting for poll.
                    var form = new WebOverlayForm(_env, show.Hwnd, opts);
                    form.NavRequested += SlideShow;
                    form.ShowAt(rect, target);
                    _overlays.Add(form);
                    Log.Info("Overlay " + rect + " -> " + target);
                }
                return true;
            }
            catch (Exception ex) { Log.Error("RebuildForSlide failed", ex); return true; }
        }

        // With presenter view there are two "screenClass" windows (presenter + audience).
        // Use the real SlideShowWindow's position (in points) to pick the audience one:
        // a positive Left/Top means it sits on a secondary monitor (further right/down).
        private Native.ShowWindow PickAudience(System.Collections.Generic.List<Native.ShowWindow> windows)
        {
            if (windows.Count == 1) return windows[0];

            bool secondary = false;
            try
            {
                dynamic ssw = _app.SlideShowWindows.Item(1);
                secondary = (float)ssw.Left > 1f || (float)ssw.Top > 1f;
            }
            catch { }

            windows.Sort((a, b) => (a.Rect.Left + a.Rect.Top).CompareTo(b.Rect.Left + b.Rect.Top));
            return secondary ? windows[windows.Count - 1] : windows[0];
        }

        private void EndShow()
        {
            _active = false;
            _builtPos = -1;
            CloseAllOverlays();
            Log.Info("Slide show ended; overlays closed.");
        }

        private void HideAllOverlays()
        {
            foreach (var o in _overlays)
            {
                try { o.Hide(); } catch { }
            }
        }

        private void CloseAllOverlays()
        {
            foreach (var o in _overlays)
            {
                try { o.Hide(); } catch { }
                try { o.Close(); } catch { }
                try { o.Dispose(); } catch { }
            }
            _overlays.Clear();
        }
    }
}
