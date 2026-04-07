using System;
using System.IO;

namespace EasyComServer
{
    public static class Logger
    {
        // Fallback-Pfad: ProgramData ist fuer LocalSystem immer schreibbar
        private static string _logFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "EasyComServer", "easycom.log");

        private static bool _console = true;
        private static readonly object _lock = new();

        public static void Init(string logFile, bool console)
        {
            _console = console;

            // Relativen Pfad -> absolut relativ zur EXE
            if (!Path.IsPathRooted(logFile))
                logFile = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, logFile);

            // Verzeichnis anlegen falls noetig
            try
            {
                string? dir = Path.GetDirectoryName(logFile);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                _logFile = logFile;
            }
            catch
            {
                // Fallback auf ProgramData bleibt aktiv
            }

            // ProgramData-Verzeichnis ebenfalls sicherstellen
            try
            {
                string fallbackDir = Path.GetDirectoryName(_logFile)!;
                if (!Directory.Exists(fallbackDir))
                    Directory.CreateDirectory(fallbackDir);
            }
            catch { }

            Log($"Logger initialisiert. Logdatei: {_logFile}");
        }

        public static void SetConsole(bool enabled) => _console = enabled;

        public static void Log(string message)
        {
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";
            lock (_lock)
            {
                // Immer in ProgramData schreiben (sicher fuer Service)
                try
                {
                    File.AppendAllText(_logFile, line + Environment.NewLine);
                }
                catch
                {
                    // Letzter Ausweg: ProgramData direkt
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
