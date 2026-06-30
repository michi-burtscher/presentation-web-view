using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace LiveWebRegion
{
    // Win32 helpers that avoid pulling extra dependencies into the hosted CLR.
    internal static class Native
    {
        #region MessageBox

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        private const uint MB_OK = 0x0;
        private const uint MB_ICONINFORMATION = 0x40;
        private const uint MB_ICONWARNING = 0x30;

        public static void Info(string text, string caption = "Live Web")
        {
            try { MessageBoxW(IntPtr.Zero, text, caption, MB_OK | MB_ICONINFORMATION); } catch { }
        }

        public static void Warn(string text, string caption = "Live Web")
        {
            try { MessageBoxW(IntPtr.Zero, text, caption, MB_OK | MB_ICONWARNING); } catch { }
        }

        #endregion

        #region Window enumeration

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder s, int max);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT r);

        // The PowerPoint slide-show full-screen window has class name "screenClass".
        private const string SlideShowClass = "screenClass";

        /// <summary>
        /// Screen rectangle (physical pixels) of this process's slide-show windows.
        /// Presenter view yields two; the largest one is the audience display.
        /// </summary>
        public static bool TryGetSlideShowRect(out Rectangle rect)
        {
            rect = Rectangle.Empty;
            uint myPid = (uint)Process.GetCurrentProcess().Id;
            Rectangle best = Rectangle.Empty;
            long bestArea = -1;

            EnumWindows((h, l) =>
            {
                if (!IsWindowVisible(h)) return true;
                GetWindowThreadProcessId(h, out uint pid);
                if (pid != myPid) return true;

                var sb = new StringBuilder(64);
                GetClassName(h, sb, sb.Capacity);
                if (sb.ToString() != SlideShowClass) return true;

                if (GetWindowRect(h, out RECT r))
                {
                    var cand = Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
                    long area = (long)cand.Width * cand.Height;
                    if (area > bestArea) { bestArea = area; best = cand; }
                }
                return true;
            }, IntPtr.Zero);

            if (bestArea > 0) { rect = best; return true; }
            return false;
        }

        #endregion
    }
}
