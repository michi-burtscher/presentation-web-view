using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace LiveWebRegion
{
    /// <summary>Movable, non-activating preview window used while editing (not in show).</summary>
    internal sealed class PreviewForm : Form
    {
        private readonly WebView2 _web;
        private readonly CoreWebView2Environment _env;
        private string _url;
        private string _pending;

        public PreviewForm(CoreWebView2Environment env)
        {
            _env = env;
            Text = "Live Web – Vorschau";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            ClientSize = new Size(520, 360);
            var wa = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(wa.Right - Width - 40, wa.Top + 80);

            _web = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_web);
            _web.CoreWebView2InitializationCompleted += (s, e) =>
            {
                if (e.IsSuccess && _pending != null) Nav(_pending);
            };
        }

        protected override bool ShowWithoutActivation => true; // don't grab focus from PowerPoint

        public void Navigate(string url)
        {
            if (url == _url) return;
            _url = url;
            if (_web.CoreWebView2 == null)
            {
                _pending = url;
                try { _web.EnsureCoreWebView2Async(_env); } catch (Exception ex) { Log.Error("Preview EnsureCoreWebView2 failed", ex); }
            }
            else Nav(url);
        }

        private void Nav(string u)
        {
            try { _web.CoreWebView2.Navigate(u); } catch (Exception ex) { Log.Error("Preview navigate failed", ex); }
        }
    }
}
