using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace LiveWebRegion
{
    /// <summary>
    /// Borderless, top-most, non-activating window hosting a WebView2. One per web
    /// region; positioned over the region's screen rectangle during a slide show.
    /// </summary>
    internal sealed class WebOverlayForm : Form
    {
        private readonly WebView2 _web;
        private readonly CoreWebView2Environment _env;
        private readonly IntPtr _ownerHwnd;
        private readonly RegionOptions _opts;
        private Timer _reloadTimer;
        private string _currentUrl;
        private string _pendingUrl;
        private bool _showingError;

        public event Action<string> NavRequested;

        private const uint VK_BACK = 0x08, VK_ESCAPE = 0x1B, VK_PRIOR = 0x21, VK_NEXT = 0x22;
        private const uint VK_LEFT = 0x25, VK_UP = 0x26, VK_RIGHT = 0x27, VK_DOWN = 0x28;

        public WebOverlayForm(CoreWebView2Environment env, IntPtr ownerHwnd, RegionOptions opts)
        {
            _env = env;
            _ownerHwnd = ownerHwnd;
            _opts = opts ?? new RegionOptions();

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = Color.White;

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            _web.NavigationCompleted += OnNavigationCompleted;
            _web.CoreWebView2InitializationCompleted += (s, e) =>
            {
                if (e.IsSuccess)
                {
                    ApplyOptions();
                    if (_opts.Interactive) HookAcceleratorKeys();
                    if (_pendingUrl != null) NavigateNow(_pendingUrl);
                }
                else Log.Error("Overlay WebView2 init failed", e.InitializationException);
            };
        }

        protected override bool ShowWithoutActivation => true;

        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_TRANSPARENT = 0x00000020;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                // Non-interactive: let mouse fall through to the slide show.
                if (_opts != null && !_opts.Interactive) cp.ExStyle |= WS_EX_TRANSPARENT;
                if (_ownerHwnd != IntPtr.Zero) cp.Parent = _ownerHwnd;
                return cp;
            }
        }

        public void ShowAt(Rectangle bounds, string url)
        {
            Bounds = bounds;
            if (!Visible) Show();

            if (_web.CoreWebView2 == null)
            {
                _pendingUrl = url;
                try { _web.EnsureCoreWebView2Async(_env); } catch (Exception ex) { Log.Error("EnsureCoreWebView2 failed", ex); }
            }
            else NavigateNow(url);
        }

        private void ApplyOptions()
        {
            try { _web.ZoomFactor = Math.Max(0.25, _opts.Zoom / 100.0); } catch { }
            try { _web.CoreWebView2.IsMuted = _opts.Mute; } catch { }
            if (_opts.AutoReloadSec > 0)
            {
                _reloadTimer = new Timer { Interval = _opts.AutoReloadSec * 1000 };
                _reloadTimer.Tick += (s, e) => { try { _web.CoreWebView2?.Reload(); } catch { } };
                _reloadTimer.Start();
            }
        }

        private void NavigateNow(string url)
        {
            if (url == _currentUrl) return;
            _currentUrl = url;
            _showingError = false;
            try { _web.CoreWebView2.Navigate(url); } catch (Exception ex) { Log.Error("Navigate failed", ex); }
        }

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess || _showingError) return;
            _showingError = true;
            try { _web.CoreWebView2.NavigateToString(ErrorPage(_currentUrl)); } catch { }
        }

        public void Reload()
        {
            try { _web.CoreWebView2?.Reload(); } catch { }
        }

        private static string ErrorPage(string url)
        {
            string safe = System.Net.WebUtility.HtmlEncode(url ?? "");
            return @"<!DOCTYPE html><html><head><meta charset='utf-8'><style>
                html,body{height:100%;margin:0;font-family:Segoe UI,sans-serif}
                body{display:flex;flex-direction:column;align-items:center;justify-content:center;
                     background:#f1f5f9;color:#334155;text-align:center;padding:24px}
                .icon{font-size:48px;margin-bottom:12px}
                h1{font-size:22px;margin:0 0 8px}
                p{margin:4px 0;color:#64748b}
                code{background:#e2e8f0;border-radius:6px;padding:2px 8px}
                </style></head><body>
                <div class='icon'>🌐⚠️</div>
                <h1>Seite nicht erreichbar</h1>
                <p>Die Web-Inhalte konnten nicht geladen werden.</p>
                <p><code>" + safe + @"</code></p>
                <p>Internetverbindung pr&#252;fen oder Pfad/URL kontrollieren.</p>
                </body></html>";
        }

        #region accelerator keys (navigation while the page has focus)

        private void HookAcceleratorKeys()
        {
            try
            {
                CoreWebView2Controller controller = FindController();
                if (controller == null) { Log.Error("CoreWebView2Controller not found via reflection."); return; }
                controller.AcceleratorKeyPressed += OnAcceleratorKeyPressed;
            }
            catch (Exception ex) { Log.Error("HookAcceleratorKeys failed", ex); }
        }

        private CoreWebView2Controller FindController()
        {
            Type t = _web.GetType();
            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
            foreach (PropertyInfo p in t.GetProperties(Flags))
                if (typeof(CoreWebView2Controller).IsAssignableFrom(p.PropertyType))
                    return p.GetValue(_web) as CoreWebView2Controller;
            foreach (FieldInfo f in t.GetFields(Flags))
                if (typeof(CoreWebView2Controller).IsAssignableFrom(f.FieldType))
                    return f.GetValue(_web) as CoreWebView2Controller;
            return null;
        }

        private void OnAcceleratorKeyPressed(object sender, CoreWebView2AcceleratorKeyPressedEventArgs e)
        {
            if (e.KeyEventKind != CoreWebView2KeyEventKind.KeyDown &&
                e.KeyEventKind != CoreWebView2KeyEventKind.SystemKeyDown)
                return;

            string nav = null;
            switch (e.VirtualKey)
            {
                case VK_ESCAPE: nav = "Exit"; break;
                case VK_RIGHT:
                case VK_DOWN:
                case VK_NEXT: nav = "Next"; break;
                case VK_LEFT:
                case VK_UP:
                case VK_PRIOR:
                case VK_BACK: nav = "Previous"; break;
            }
            if (nav == null) return;

            e.Handled = true;
            NavRequested?.Invoke(nav);
        }

        #endregion
    }
}
