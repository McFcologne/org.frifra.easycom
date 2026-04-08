using System;
using System.IO;

namespace EasyComServer
{
    public static class Logger
    {
        // Fallback path: ProgramData is always writable for LocalSystem
        private static string _logFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "EasyComServer", "easycom.log");

        private static bool _console = true;
        private static readonly object _lock = new();

        public static void Init(string logFile, bool console)
        {
            _console = console;

            // Convert relative path to absolute, relative to the executable directory
            if (!Path.IsPathRooted(logFile))
                logFile = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, logFile);

            // Create the log directory if it does not exist
            try
            {
                string? dir = Path.GetDirectoryName(logFile);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                _logFile = logFile;
            }
            catch
            {
                // ProgramData fallback remains active
            }

            // Also ensure the ProgramData directory exists
            try
            {
                string fallbackDir = Path.GetDirectoryName(_logFile)!;
                if (!Directory.Exists(fallbackDir))
                    Directory.CreateDirectory(fallbackDir);
            }
            catch { }

            Log($"Logger initialized. Log file: {_logFile}");
        }

        public static void SetConsole(bool enabled) => _console = enabled;

        public static void Log(string message)
        {
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";
            lock (_lock)
            {
                // Always write to the configured log file (safe for Windows Service)
                try
                {
                    File.AppendAllText(_logFile, line + Environment.NewLine);
                }
                catch
                {
                    // Last resort: write directly to ProgramData
                    try
                    {
                        string fallback = Path.Combine(
                            Environment.GetFolderPath(
                                Environment.SpecialFolder.CommonApplicationData),
                            "EasyComServer", "easycom.log");
                        Directory.CreateDirectory(Path.GetDirectoryName(fallback)!);
                        File.AppendAllText(fallback, line + Environment.NewLine);
                    }
                    catch { }
                }

                if (_console)
                    try { Console.WriteLine(line); } catch { }
            }
        }
    }
}
