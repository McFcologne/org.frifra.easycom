using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EasyComServer
{
    public class ServerConfig
    {
        public string DllPath { get; set; } = "EASY_COM.dll";
        public string LogFile { get; set; } = "easycom.log";
        public bool ConsoleLogging { get; set; } = true;
        public int ComIdleTimeoutSeconds { get; set; } = 300;

        // ── Global Basic Auth (applies to all instances) ──────────────────────
        public bool BasicAuthEnabled { get; set; } = false;
        public string BasicAuthUser { get; set; } = "admin";
        public string BasicAuthPass { get; set; } = "";

        public List<InstanceConfig> Instances { get; set; } = new();

        public static ServerConfig Load(string iniPath)
        {
            var cfg = new ServerConfig();

            if (!File.Exists(iniPath))
            {
                WriteDefault(iniPath);
                return cfg;
            }

            string? currentSection = null;
            InstanceConfig? currentInstance = null;

            foreach (var rawLine in File.ReadAllLines(iniPath, Encoding.UTF8))
            {
                string line = rawLine.Trim();
                if (line.StartsWith(";") || line.StartsWith("#") || line.Length == 0)
                    continue;

                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    if (currentInstance != null)
                        cfg.Instances.Add(currentInstance);

                    currentSection = line[1..^1].Trim().ToLower();

                    if (currentSection.StartsWith("instance"))
                    {
                        string instName = currentSection.Length > 8
                            ? currentSection[8..].Trim().Trim(':').Trim()
                            : $"inst{cfg.Instances.Count + 1}";
                        currentInstance = new InstanceConfig { Name = instName };
                    }
                    else
                    {
                        currentInstance = null;
                    }
                    continue;
                }

                var idx = line.IndexOf('=');
                if (idx < 0) continue;
                string key = line[..idx].Trim().ToLower();
                string val = line[(idx + 1)..].Trim();
                int commentIdx = val.IndexOf(';');
                if (commentIdx > 0) val = val[..commentIdx].Trim();

                if (currentInstance != null)
                {
                    switch (key)
                    {
                        case "name": currentInstance.Name = val; break;
                        case "http_enabled": currentInstance.HttpEnabled = ParseBool(val); break;
                        case "http_port": currentInstance.HttpPort = int.Parse(val); break;
                        case "telnet_enabled": currentInstance.TelnetEnabled = ParseBool(val); break;
                        case "telnet_port": currentInstance.TelnetPort = int.Parse(val); break;
                        case "com_port": currentInstance.ComPort = int.Parse(val); break;
                        case "baud_rate": currentInstance.BaudRate = int.Parse(val); break;
                        // Auth settings also read from [instance] for backward compatibility
                        case "basic_auth": cfg.BasicAuthEnabled = ParseBool(val); break;
                        case "auth_user": cfg.BasicAuthUser = val; break;
                        case "auth_pass": cfg.BasicAuthPass = val; break;
                    }
                }
                else
                {
                    switch (key)
                    {
                        case "dll_path": cfg.DllPath = val; break;
                        case "log_file": cfg.LogFile = val; break;
                        case "console_logging": cfg.ConsoleLogging = ParseBool(val); break;
                        case "com_idle_timeout": cfg.ComIdleTimeoutSeconds = int.Parse(val); break;
                        case "basic_auth": cfg.BasicAuthEnabled = ParseBool(val); break;
                        case "auth_user": cfg.BasicAuthUser = val; break;
                        case "auth_pass": cfg.BasicAuthPass = val; break;
                    }
                }
            }

            if (currentInstance != null)
                cfg.Instances.Add(currentInstance);

            if (cfg.Instances.Count == 0)
                cfg.Instances.Add(new InstanceConfig());

            return cfg;
        }

        private static bool ParseBool(string s)
            => s.Equals("true", StringComparison.OrdinalIgnoreCase)
            || s == "1"
            || s.Equals("yes", StringComparison.OrdinalIgnoreCase)
            || s.Equals("enabled", StringComparison.OrdinalIgnoreCase);

        private static void WriteDefault(string path)
        {
            File.WriteAllText(path, @"; EasyComServer configuration file
; ========================================
[global]
dll_path         = EASY_COM.dll
log_file         = easycom.log
console_logging  = true
com_idle_timeout = 300

; HTTP Basic Auth (global — applies to all instances)
; Leave auth_pass empty to disable authentication.
basic_auth       = false
auth_user        = admin
auth_pass        =

[instance]
name             = default
http_enabled     = true
http_port        = 8083
telnet_enabled   = true
telnet_port      = 8023
com_port         = 1
baud_rate        = 9600
", Encoding.UTF8);
        }
    }

    public class InstanceConfig
    {
        public string Name { get; set; } = "default";
        public bool HttpEnabled { get; set; } = true;
        public int HttpPort { get; set; } = 8083;
        public bool TelnetEnabled { get; set; } = true;
        public int TelnetPort { get; set; } = 8023;
        public int ComPort { get; set; } = 0;
        public int BaudRate { get; set; } = 9600;
        public bool HasComConfig => ComPort > 0;
    }
}
