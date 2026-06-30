using System;
using System.Drawing;
using System.Windows.Forms;

namespace LiveWebRegion
{
    /// <summary>Asks for a web address (URL) or local HTML file, plus per-frame options.</summary>
    internal sealed class LinkInputDialog : Form
    {
        private readonly TextBox _txt;
        private readonly CheckBox _embed;
        private readonly CheckBox _interactive;
        private readonly CheckBox _mute;
        private readonly NumericUpDown _zoom;
        private readonly NumericUpDown _reload;

        public string Value => _txt.Text.Trim();
        public bool Embed => _embed.Checked;

        public RegionOptions Options => new RegionOptions
        {
            Interactive = _interactive.Checked,
            Mute = _mute.Checked,
            Zoom = (int)_zoom.Value,
            AutoReloadSec = (int)_reload.Value
        };

        public LinkInputDialog(string initial, bool embed, RegionOptions opts)
        {
            opts = opts ?? new RegionOptions();
            Text = "Live Web Frame";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 384);
            Font = new Font("Segoe UI", 9f);

            // --- header (card look) ---
            var header = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = Color.FromArgb(30, 58, 138) };
            var hTitle = new Label
            {
                Text = "🌐  Live Web Frame",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 13f, FontStyle.Bold),
                Left = 16, Top = 8, Width = 528, Height = 26, BackColor = Color.Transparent
            };
            var hSub = new Label
            {
                Text = "Website (URL oder HTML-Datei) live im Frame anzeigen – auch im Präsentationsmodus.",
                ForeColor = Color.FromArgb(199, 218, 255),
                Left = 18, Top = 34, Width = 528, Height = 18, BackColor = Color.Transparent
            };
            header.Controls.AddRange(new Control[] { hTitle, hSub });

            int y = 74;
            var lbl = new Label { Text = "Adresse (URL) oder lokale HTML-Datei:", Left = 16, Top = y, Width = 528 };
            string start = string.IsNullOrEmpty(initial) ? "https://" : initial;
            _txt = new TextBox { Left = 16, Top = y + 22, Width = 430, Text = start };
            var browse = new Button { Text = "Datei…", Left = 454, Top = y + 20, Width = 90, Height = 26 };
            browse.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog { Filter = "Webdateien (*.html;*.htm)|*.html;*.htm|Alle Dateien (*.*)|*.*" })
                    if (ofd.ShowDialog(this) == DialogResult.OK) _txt.Text = ofd.FileName;
            };
            var hint = new Label
            {
                Text = "Beispiele:  https://example.com  ·  C:\\Pfad\\seite.html",
                Left = 16, Top = y + 52, Width = 528, ForeColor = Color.Gray
            };
            _embed = new CheckBox
            {
                Text = "Lokale HTML-Datei in die Präsentation einbetten (für Weitergabe)",
                Left = 16, Top = y + 78, Width = 528, Checked = embed
            };

            var grp = new GroupBox { Text = "Optionen", Left = 16, Top = y + 108, Width = 528, Height = 122 };
            _interactive = new CheckBox { Text = "Interaktiv (Klicks/Eingaben an die Seite)", Left = 16, Top = 24, Width = 360, Checked = opts.Interactive };
            _mute = new CheckBox { Text = "Ton stummschalten", Left = 16, Top = 50, Width = 360, Checked = opts.Mute };
            var zlbl = new Label { Text = "Zoom (%):", Left = 16, Top = 80, Width = 70 };
            _zoom = new NumericUpDown { Left = 90, Top = 78, Width = 70, Minimum = 25, Maximum = 500, Increment = 10, Value = Clamp(opts.Zoom, 25, 500) };
            var rlbl = new Label { Text = "Auto-Neuladen alle (Sek., 0 = aus):", Left = 200, Top = 80, Width = 210 };
            _reload = new NumericUpDown { Left = 412, Top = 78, Width = 80, Minimum = 0, Maximum = 86400, Value = Clamp(opts.AutoReloadSec, 0, 86400) };
            grp.Controls.AddRange(new Control[] { _interactive, _mute, zlbl, _zoom, rlbl, _reload });

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 360, Top = y + 240, Width = 84, Height = 30 };
            var cancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, Left = 452, Top = y + 240, Width = 92, Height = 30 };

            AcceptButton = ok;
            CancelButton = cancel;
            Controls.AddRange(new Control[] { lbl, _txt, browse, hint, _embed, grp, ok, cancel, header });
            Shown += (s, e) => { _txt.Focus(); _txt.SelectionStart = _txt.TextLength; };
        }

        private static decimal Clamp(int v, int min, int max) => v < min ? min : (v > max ? max : v);
    }
}
