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
        private string _currentUrl;
        private string _pendingUrl;

        // Raised with "Next" / "Previous" / "Exit" when a navigation key is pressed
        // while the page has keyboard focus.
        public event Action<string> NavRequested;

        // Virtual-key codes (Win32).
        private const uint VK_BACK = 0x08, VK_ESCAPE = 0x1B, VK_PRIOR = 0x21, VK_NEXT = 0x22;
        private const uint VK_LEFT = 0x25, VK_UP = 0x26, VK_RIGHT = 0x27, VK_DOWN = 0x28;

        public WebOverlayForm(CoreWebView2Environment env)
        {
            _env = env;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = Color.White;

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            _web.CoreWebView2InitializationCompleted += (s, e) =>
            {
                if (e.IsSuccess)
                {
                    HookAcceleratorKeys();
                    if (_pendingUrl != null) NavigateNow(_pendingUrl);
                }
                else Log.Error("Overlay WebView2 init failed", e.InitializationException);
            };
        }

        // The WinForms wrapper doesn't expose the controller, but AcceleratorKeyPressed
        // (the supported way to handle nav keys while the page has focus) lives on it.
        private void HookAcceleratorKeys()
        {
            try
            {
                CoreWebView2Controller controller = FindController();
                if (controller == null) { Log.Error("CoreWebView2Controller not found via reflection."); return; }
                controller.AcceleratorKeyPressed += OnAcceleratorKeyPressed;
                Log.Info("AcceleratorKeyPressed hooked.");
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
            if (nav == null) return; // let the page keep character/Space/Enter keys

            e.Handled = true;
            NavRequested?.Invoke(nav);
        }

        // Do not steal focus from the running slide show.
        protected override bool ShowWithoutActivation => true;

        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        /// <summary>Position the overlay and (re)load the given file path.</summary>
        public void ShowAt(Rectangle bounds, string filePath)
        {
            Bounds = bounds;
            string url = new Uri(filePath).AbsoluteUri;

            if (!Visible) Show();

            if (_web.CoreWebView2 == null)
            {
                _pendingUrl = url;
                try { _web.EnsureCoreWebView2Async(_env); } catch (Exception ex) { Log.Error("EnsureCoreWebView2 failed", ex); }
            }
            else
            {
                NavigateNow(url);
            }
        }

        private void NavigateNow(string url)
        {
            if (url == _currentUrl) return;
            _currentUrl = url;
            try { _web.CoreWebView2.Navigate(url); } catch (Exception ex) { Log.Error("Navigate failed", ex); }
        }

        public void Reload()
        {
            try { _web.CoreWebView2?.Reload(); } catch { }
        }
    }
}
