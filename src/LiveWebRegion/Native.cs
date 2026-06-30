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

        public struct ShowWindow
        {
            public IntPtr Hwnd;
            public Rectangle Rect;
        }

        /// <summary>
        /// Slide-show windows (handle + physical-pixel rect) of this process.
        /// Presenter view produces two; the caller picks the audience one.
        /// </summary>
        public static System.Collections.Generic.List<ShowWindow> GetSlideShowWindows()
        {
            var list = new System.Collections.Generic.List<ShowWindow>();
            uint myPid = (uint)Process.GetCurrentProcess().Id;

            EnumWindows((h, l) =>
            {
                if (!IsWindowVisible(h)) return true;
                GetWindowThreadProcessId(h, out uint pid);
                if (pid != myPid) return true;

                var sb = new StringBuilder(64);
                GetClassName(h, sb, sb.Capacity);
                if (sb.ToString() != SlideShowClass) return true;

                if (GetWindowRect(h, out RECT r))
                    list.Add(new ShowWindow { Hwnd = h, Rect = Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom) });
                return true;
            }, IntPtr.Zero);

            return list;
        }

        #endregion
    }
}
