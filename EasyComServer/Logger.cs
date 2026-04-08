using System;
using System.IO;
using Serilog;
using Serilog.Events;

namespace EasyComServer
{
    public static class Logger
    {
        private static readonly string _fallbackPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "EasyComServer", "easycom.log");

        private static ILogger? _log;
        private static volatile bool _console = true;

        /// <summary>
        /// Initialises the logger. Rolling happens daily and whenever the file
        /// reaches <paramref name="maxSizeMb"/> MB. Up to <paramref name="maxFiles"/>
        /// rotated files are kept (0 = unlimited).
        /// </summary>
        public static void Init(string logFile, bool console,
            int maxSizeMb = 10, int maxFiles = 10)
        {
            _console = console;

            if (!Path.IsPathRooted(logFile))
                logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logFile);

            // Try to create the configured log directory; fall back to ProgramData.
            string resolvedPath = TryEnsureDirectory(logFile) ?? _fallbackPath;
            TryEnsureDirectory(_fallbackPath);

            long?  sizeLimit   = maxSizeMb > 0 ? (long)maxSizeMb * 1024 * 1024 : (long?)null;
            int?   retainCount = maxFiles   > 0 ? maxFiles                       : (int?)null;

            _log = new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.File(
                    path:                    resolvedPath,
                    rollingInterval:         RollingInterval.Day,
                    fileSizeLimitBytes:      sizeLimit,
                    retainedFileCountLimit:  retainCount,
                    rollOnFileSizeLimit:     true,
                    outputTemplate:          "{Timestamp:yyyy-MM-dd HH:mm:ss} {Message:lj}{NewLine}")
                .CreateLogger();

            string sizeInfo  = maxSizeMb > 0 ? $"{maxSizeMb} MB" : "unlimited";
            string filesInfo = maxFiles  > 0 ? maxFiles.ToString() : "unlimited";
            Log($"Logger initialized: {resolvedPath} " +
                $"(rotate daily / {sizeInfo}, keep {filesInfo} files)");
        }

        /// <summary>Toggles console output at runtime (e.g. via SET CONFIGURATION).</summary>
        public static void SetConsole(bool enabled) => _console = enabled;

        public static void Log(string message)
        {
            // File: Serilog handles thread-safety and rotation internally.
            _log?.Information(message);

            // Console: checked separately so SetConsole works without reinitialising.
            if (_console)
                try { Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}"); }
                catch { }
        }

        /// <summary>
        /// Flushes and closes the underlying Serilog logger. Call during service shutdown.
        /// </summary>
        public static void Close()
        {
            if (_log is IDisposable d)
                d.Dispose();
            _log = null;
        }

        private static string? TryEnsureDirectory(string filePath)
        {
            try
            {
                string? dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);
                return filePath;
            }
            catch { return null; }
        }
    }
}
