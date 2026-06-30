using System;
using System.IO;

namespace LiveWebRegion
{
    // Lightweight file logger. PowerPoint swallows add-in exceptions silently,
    // so a log file is the primary way to see what is happening.
    internal static class Log
    {
        private static readonly object Gate = new object();

        public static string LogPath
        {
            get
            {
                // In-process inside PowerPoint, GetFolderPath can occasionally come back
                // empty/relative; fall back to the always-rooted temp path so we never
                // silently write to PowerPoint's working directory.
                string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrEmpty(baseDir) || !Path.IsPathRooted(baseDir))
                    baseDir = Path.GetTempPath();
                string dir = Path.Combine(baseDir, "LiveWebRegion");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "addin.log");
            }
        }

        public static void Info(string message)
        {
            Write("INFO", message);
        }

        public static void Error(string message, Exception ex = null)
        {
            Write("ERROR", ex == null ? message : message + " :: " + ex);
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (Gate)
                {
                    File.AppendAllText(
                        LogPath,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}{Environment.NewLine}");
                }
            }
            catch
            {
                // logging must never throw
            }
        }
    }
}
