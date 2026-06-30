using System;
using System.Globalization;

namespace LiveWebRegion
{
    /// <summary>Per-region display options, persisted in a shape tag.</summary>
    internal sealed class RegionOptions
    {
        public bool Interactive = true;   // false => clicks pass through to the slide
        public bool Mute = false;
        public int Zoom = 100;            // percent
        public int AutoReloadSec = 0;     // 0 = off

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "interactive={0};mute={1};zoom={2};reload={3}",
                Interactive ? 1 : 0, Mute ? 1 : 0, Zoom, AutoReloadSec);
        }

        public static RegionOptions Parse(string s)
        {
            var o = new RegionOptions();
            if (string.IsNullOrEmpty(s)) return o;
            foreach (var part in s.Split(';'))
            {
                int eq = part.IndexOf('=');
                if (eq <= 0) continue;
                string k = part.Substring(0, eq).Trim();
                string v = part.Substring(eq + 1).Trim();
                int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n);
                switch (k)
                {
                    case "interactive": o.Interactive = n != 0; break;
                    case "mute": o.Mute = n != 0; break;
                    case "zoom": o.Zoom = n > 0 ? n : 100; break;
                    case "reload": o.AutoReloadSec = n >= 0 ? n : 0; break;
                }
            }
            return o;
        }
    }

    /// <summary>Result of the link dialog.</summary>
    internal sealed class LinkResult
    {
        public string Value;
        public bool Embed;
        public RegionOptions Options = new RegionOptions();
    }
}
