using System;
using System.Drawing;
using System.Windows.Forms;

namespace LiveWebRegion
{
    /// <summary>Asks for a web address (URL) or a local HTML file.</summary>
    internal sealed class LinkInputDialog : Form
    {
        private readonly TextBox _txt;

        public string Value => _txt.Text.Trim();

        public LinkInputDialog(string initial)
        {
            Text = "Web-Link setzen";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 168);
            Font = new Font("Segoe UI", 9f);

            var lbl = new Label { Text = "Adresse (URL) oder lokale HTML-Datei:", Left = 16, Top = 16, Width = 528, Height = 20 };
            // Pre-fill "https://" for convenience so the user can type the rest directly.
            string start = string.IsNullOrEmpty(initial) ? "https://" : initial;
            _txt = new TextBox { Left = 16, Top = 40, Width = 430, Text = start };
            Shown += (s, e) => { _txt.Focus(); _txt.SelectionStart = _txt.TextLength; _txt.SelectionLength = 0; };
            var browse = new Button { Text = "Datei…", Left = 454, Top = 38, Width = 90, Height = 26 };
            browse.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Filter = "Webdateien (*.html;*.htm)|*.html;*.htm|Alle Dateien (*.*)|*.*" })
                {
                    if (ofd.ShowDialog(this) == DialogResult.OK) _txt.Text = ofd.FileName;
                }
            };
            var hint = new Label
            {
                Text = "Beispiele:  https://example.com  ·  C:\\Pfad\\seite.html",
                Left = 16, Top = 72, Width = 528, Height = 20, ForeColor = Color.Gray
            };
            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 360, Top = 122, Width = 84, Height = 28 };
            var cancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, Left = 452, Top = 122, Width = 92, Height = 28 };

            AcceptButton = ok;
            CancelButton = cancel;
            Controls.AddRange(new Control[] { lbl, _txt, browse, hint, ok, cancel });
        }
    }
}
