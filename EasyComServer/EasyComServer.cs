using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EasyComServer
{
    public class EasyComService : ServiceBase
    {
        private ServerConfig _config;
        private EasyComWrapper _wrapper; // shared for single-instance setups
        private readonly Dictionary<string, EasyComWrapper> _wrappers = new(); // per-instance
        private List<HttpListener> _httpListeners = new();
        private List<TcpListener> _telnetListeners = new();
        private CancellationTokenSource _cts = new();
        private readonly DateTime _startTime = DateTime.Now;
        private string _dllVersion = "unknown";

        public EasyComService()
        {
            ServiceName = "EasyComServer";
            CanStop = true;
            CanPauseAndContinue = false;
            AutoLog = true;
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                string iniPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "easycom.ini");
                _config = ServerConfig.Load(iniPath);
                Logger.Init(_config.LogFile, _config.ConsoleLogging);
                Logger.Log($"EasyComServer starting, config: {iniPath}");

                // Each instance gets its own EasyComWrapper
                // so COM ports are managed independently
                foreach (var inst in _config.Instances)
                {
                    var wrapper = new EasyComWrapper(_config.DllPath);
                    wrapper.SetComIdleTimeout(_config.ComIdleTimeoutSeconds);
                    if (inst.HasComConfig)
                        wrapper.SetDefaultCom(inst.ComPort, inst.BaudRate);
                    _wrappers[inst.Name] = wrapper;
                }

                // _wrapper points to the first wrapper for backward compatibility
                _wrapper = _wrappers.Count > 0
                    ? _wrappers.Values.First()
                    : new EasyComWrapper(_config.DllPath);

                _dllVersion = _wrapper.GetVersion();
                Logger.Log($"EASY_COM.dll loaded, version: {_dllVersion}");

                foreach (var inst in _config.Instances)
                {
                    if (inst.HttpEnabled)
                        StartHttpListener(inst);
                    if (inst.TelnetEnabled)
                        StartTelnetListener(inst);
                }

                Logger.Log($"Basic Auth: {(_config.BasicAuthEnabled ? $"ENABLED user={_config.BasicAuthUser}" : "disabled")}");
                Logger.Log("EasyComServer started successfully.");
            }
            catch (Exception ex)
            {
                Logger.Log($"FATAL: {ex}");
                throw;
            }
        }

        protected override void OnStop()
        {
            _cts.Cancel();
            foreach (var w in _wrappers.Values) { try { w.Dispose(); } catch { } }
            if (!_wrappers.ContainsValue(_wrapper)) _wrapper?.Dispose();
            Logger.Log("EasyComServer stopped.");
        }

        private void StartHttpListener(InstanceConfig inst)
        {
            var listener = new HttpListener();
            string prefix = $"http://*:{inst.HttpPort}/";
            listener.Prefixes.Add(prefix);
            listener.Start();
            _httpListeners.Add(listener);
            Logger.Log($"HTTP listener started on port {inst.HttpPort} (instance: {inst.Name})");

            Task.Run(async () =>
            {
                var instWrapper = _wrappers.TryGetValue(inst.Name, out var w) ? w : _wrapper;
                var cmdProcessor = new CommandProcessor(instWrapper, _config, inst, _startTime, _dllVersion);
                while (!_cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        var ctx = await listener.GetContextAsync();
                        _ = Task.Run(() => HandleHttp(ctx, cmdProcessor, inst));
                    }
                    catch (HttpListenerException) when (_cts.IsCancellationRequested) { break; }
                    catch (Exception ex) { Logger.Log($"HTTP error: {ex.Message}"); }
                }
            }, _cts.Token);
        }

        private void HandleHttp(HttpListenerContext ctx, CommandProcessor cmdProcessor,
                                 InstanceConfig inst)
        {
            try
            {
                string absPath = ctx.Request.Url?.AbsolutePath ?? "/";
                string query = ctx.Request.Url?.Query ?? "";

                // ── Always set CORS headers (including on 401 responses) ──────
                ctx.Response.AddHeader("Access-Control-Allow-Origin", "*");
                ctx.Response.AddHeader("Access-Control-Allow-Headers",
                    "Authorization, Content-Type");
                ctx.Response.AddHeader("Access-Control-Allow-Methods",
                    "GET, POST, OPTIONS");

                // ── Always allow OPTIONS preflight without authentication ─────
                if (ctx.Request.HttpMethod == "OPTIONS")
                {
                    ctx.Response.StatusCode = 204;
                    ctx.Response.Close();
                    return;
                }

                // ── Basic authentication required for all requests ────────────
                if (_config.BasicAuthEnabled && !CheckBasicAuth(ctx))
                {
                    ctx.Response.StatusCode = 401;
                    ctx.Response.AddHeader("WWW-Authenticate",
                        $"Basic realm=\"EasyComServer-{inst.Name}\"");
                    byte[] deny = Encoding.UTF8.GetBytes("401 Unauthorized\r\n");
                    ctx.Response.ContentLength64 = deny.Length;
                    ctx.Response.OutputStream.Write(deny, 0, deny.Length);
                    ctx.Response.Close();
                    return;
                }

                // ── /easy.cmd?COMMAND  (API endpoint) ─────────────────────────
                if (absPath.EndsWith("easy.cmd", StringComparison.OrdinalIgnoreCase)
                    || absPath == "/")
                {
                    // Parse the query string: strip key=value parameters such as
                    // _=timestamp (jQuery cache buster) and keep only the EASY command.
                    // Example: ?_=1543364225097&WRITE_OBJECT_VALUE%201%204%201%201
                    //      or: ?WRITE_OBJECT_VALUE%201%204%201%201
                    string rawQuery = Uri.UnescapeDataString(query.TrimStart('?')).Trim();
                    string cmd = ParseEasyCmd(rawQuery);

                    // Root path without query → redirect to the web console
                    if (string.IsNullOrEmpty(cmd) && absPath == "/")
                    {
                        ctx.Response.Redirect("/index.html");
                        ctx.Response.Close();
                        return;
                    }

                    string response = cmdProcessor.Execute(cmd);
                    byte[] buf = Encoding.UTF8.GetBytes(response + "\r\n");
                    ctx.Response.ContentType = "text/plain; charset=utf-8";
                    ctx.Response.ContentLength64 = buf.Length;
                    ctx.Response.OutputStream.Write(buf, 0, buf.Length);
                    ctx.Response.Close();
                    return;
                }

                // ── Static files from wwwroot/ ────────────────────────────────
                string wwwroot = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
                string filePath = Path.Combine(wwwroot,
                    absPath.TrimStart('/').Replace('/', Path.DirectorySeparatorChar));

                // Default document handling
                if (Directory.Exists(filePath))
                    filePath = Path.Combine(filePath, "index.html");

                if (File.Exists(filePath))
                {
                    byte[] fileBytes = File.ReadAllBytes(filePath);
                    ctx.Response.ContentType = GetMimeType(filePath);
                    ctx.Response.ContentLength64 = fileBytes.Length;
                    ctx.Response.OutputStream.Write(fileBytes, 0, fileBytes.Length);
                    ctx.Response.Close();
                    return;
                }

                // 404
                ctx.Response.StatusCode = 404;
                byte[] notFound = Encoding.UTF8.GetBytes("404 Not Found\r\n");
                ctx.Response.ContentLength64 = notFound.Length;
                ctx.Response.OutputStream.Write(notFound, 0, notFound.Length);
                ctx.Response.Close();
            }
            catch (Exception ex)
            {
                Logger.Log($"HTTP handler error: {ex.Message}");
                try { ctx.Response.Abort(); } catch { }
            }
        }

        private bool CheckBasicAuth(HttpListenerContext ctx)
        {
            string? authHeader = ctx.Request.Headers["Authorization"];
            if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Basic "))
                return false;
            try
            {
                string decoded = Encoding.UTF8.GetString(
                    Convert.FromBase64String(authHeader.Substring(6)));
                int colon = decoded.IndexOf(':');
                if (colon < 0) return false;
                string user = decoded[..colon];
                string pass = decoded[(colon + 1)..];
                return user == _config.BasicAuthUser && pass == _config.BasicAuthPass;
            }
            catch { return false; }
        }

        /// <summary>
        /// Extracts the EASY command from the query string.
        /// Ignores key=value parameters such as _=timestamp (jQuery cache buster).
        /// Supports both formats:
        ///   /easy.cmd?WRITE_OBJECT_VALUE 1 4 1 1 1|0|10
        ///   /easy.cmd?_=1543364225097&WRITE_OBJECT_VALUE 1 4 1 1 1|0|10
        /// </summary>
        private static string ParseEasyCmd(string rawQuery)
        {
            if (string.IsNullOrWhiteSpace(rawQuery)) return "";

            // Split by & to get individual parts
            var parts = rawQuery.Split('&');
            foreach (var part in parts)
            {
                string p = part.Trim();
                if (string.IsNullOrEmpty(p)) continue;

                // Skip key=value pairs where key contains no spaces
                // (these are HTTP parameters, not EASY commands)
                int eq = p.IndexOf('=');
                if (eq > 0)
                {
                    string key = p[..eq].Trim();
                    // If key has no spaces and is short, it is an HTTP param -> skip
                    if (!key.Contains(' ') && key.Length < 20)
                        continue;
                }

                // This part looks like an EASY command
                return p;
            }

            // Fallback: return raw query as-is
            return rawQuery;
        }

        private static string GetMimeType(string path) =>
            Path.GetExtension(path).ToLower() switch
            {
                ".html" or ".htm" => "text/html; charset=utf-8",
                ".css" => "text/css; charset=utf-8",
                ".js" => "application/javascript; charset=utf-8",
                ".json" => "application/json",
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".svg" => "image/svg+xml",
                ".ico" => "image/x-icon",
                _ => "application/octet-stream",
            };

        private void StartTelnetListener(InstanceConfig inst)
        {
            var listener = new TcpListener(IPAddress.Any, inst.TelnetPort);
            listener.Start();
            _telnetListeners.Add(listener);
            Logger.Log($"Telnet listener started on port {inst.TelnetPort} (instance: {inst.Name})");

            Task.Run(async () =>
            {
                var instWrapper = _wrappers.TryGetValue(inst.Name, out var w) ? w : _wrapper;
                var cmdProcessor = new CommandProcessor(instWrapper, _config, inst, _startTime, _dllVersion);
                while (!_cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        var client = await listener.AcceptTcpClientAsync();
                        _ = Task.Run(() => HandleTelnet(client, cmdProcessor));
                    }
                    catch (ObjectDisposedException) when (_cts.IsCancellationRequested) { break; }
                    catch (Exception ex) { Logger.Log($"Telnet accept error: {ex.Message}"); }
                }
            }, _cts.Token);
        }

        private void HandleTelnet(TcpClient client, CommandProcessor cmdProcessor)
        {
            try
            {
                using (client)
                using (var stream = client.GetStream())
                using (var reader = new StreamReader(stream, Encoding.ASCII))
                using (var writer = new StreamWriter(stream, Encoding.ASCII) { AutoFlush = true })
                {
                    writer.WriteLine($"Moeller EASY(r) Server (EasyComServer .NET)");
                    writer.Write("> ");
                    string? line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (string.IsNullOrEmpty(line)) { writer.Write("> "); continue; }
                        if (line.Equals("quit", StringComparison.OrdinalIgnoreCase) ||
                            line.Equals("exit", StringComparison.OrdinalIgnoreCase)) break;
                        string resp = cmdProcessor.Execute(line);
                        writer.WriteLine(resp);
                        writer.Write("> ");
                    }
                }
            }
            catch (Exception ex) { Logger.Log($"Telnet client error: {ex.Message}"); }
        }

        // Entry point for both service and console mode
        public static void Main(string[] args)
        {
            if (args.Length > 0 && args[0].Equals("--console", StringComparison.OrdinalIgnoreCase))
            {
                var svc = new EasyComService();
                svc.OnStart(args);
                Console.WriteLine("Running in console mode. Press Enter to stop.");
                Console.ReadLine();
                svc.OnStop();
            }
            else
            {
                ServiceBase.Run(new EasyComService());
            }
        }
    }
}
