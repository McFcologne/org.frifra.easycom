using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json.Nodes;

namespace EasyComServer
{
    /// <summary>
    /// Handles RESTful API requests under /api/v1/.
    /// All device operations delegate to the existing CommandProcessor so the
    /// EASY_COM.dll wrapper, connection lifecycle, and logging are untouched.
    /// </summary>
    internal static class RestApiHandler
    {
        public static void Handle(HttpListenerContext ctx, CommandProcessor cmd)
        {
            string method  = ctx.Request.HttpMethod.ToUpperInvariant();
            string absPath = ctx.Request.Url?.AbsolutePath ?? "/";

            // Strip /api/v1 prefix → relative path, e.g. "device/1/clock"
            string rel = absPath;
            if (rel.StartsWith("/api/v1", StringComparison.OrdinalIgnoreCase))
                rel = rel[7..];
            rel = rel.TrimStart('/');

            string[] segs = rel.Split('/', StringSplitOptions.RemoveEmptyEntries);

            try
            {
                (int status, string body) = Route(method, segs, ctx, cmd);
                WriteResponse(ctx, status, body);
            }
            catch (Exception ex)
            {
                Logger.Log($"REST API error: {ex.Message}");
                WriteResponse(ctx, 500, Error(ex.Message));
            }
        }

        // ── Top-level router ──────────────────────────────────────────────────

        private static (int, string) Route(string method, string[] segs,
            HttpListenerContext ctx, CommandProcessor cmd)
        {
            if (segs.Length == 0)
                return (200, Ok("EasyComServer REST API v1"));

            switch (segs[0])
            {
                // /api/v1/system[/error|/connections]
                case "system":
                    if (segs.Length == 1 && method == "GET")
                        return OkCmd(cmd, "show server");
                    if (segs.Length == 2 && method == "GET")
                        return segs[1] switch
                        {
                            "error"       => OkCmd(cmd, "getlastsystemerror"),
                            "connections" => OkCmd(cmd, "show connections"),
                            _             => (404, Error($"Unknown system resource: {segs[1]}"))
                        };
                    break;

                // /api/v1/connection[/baudrate|/waitingtime|/com|/ethernet]
                case "connection":
                    if (segs.Length == 1 && method == "GET")
                        return OkCmd(cmd, "show connections");
                    if (segs.Length == 2)
                        return RouteConnection(method, segs[1], ctx, cmd);
                    break;

                // /api/v1/device/{netId}/...
                case "device":
                    if (segs.Length >= 3 && int.TryParse(segs[1], out int netId))
                        return RouteDevice(method, netId, segs, ctx, cmd);
                    break;
            }

            return (404, Error("Not found"));
        }

        // ── /api/v1/connection/... ────────────────────────────────────────────

        private static (int, string) RouteConnection(string method, string sub,
            HttpListenerContext ctx, CommandProcessor cmd)
        {
            switch (sub)
            {
                case "baudrate":
                    if (method == "GET") return OkCmd(cmd, "getcurrent_baudrate");
                    break;

                case "waitingtime":
                    if (method == "GET") return OkCmd(cmd, "get_userwaitingtime");
                    if (method == "PUT")
                    {
                        var body = ReadBody(ctx);
                        int? ms = GetInt(body, "ms");
                        if (ms == null) return (400, Error("Required: ms (integer milliseconds)"));
                        return OkCmd(cmd, $"set_userwaitingtime {ms}");
                    }
                    break;

                case "com":
                    if (method == "DELETE") return OkCmd(cmd, "close_comport");
                    if (method == "POST")
                    {
                        var body = ReadBody(ctx);
                        int? port = GetInt(body, "port");
                        int? baud = GetInt(body, "baud");
                        if (port == null || baud == null)
                            return (400, Error("Required: port (integer), baud (integer)"));
                        return OkCmd(cmd, $"open_comport {port} {baud}");
                    }
                    break;

                case "ethernet":
                    if (method == "DELETE") return OkCmd(cmd, "close_ethernetport");
                    if (method == "POST")
                    {
                        var body   = ReadBody(ctx);
                        string? ip = GetString(body, "ip");
                        int?  port = GetInt(body, "port");
                        int?  baud = GetInt(body, "baud");
                        bool  nobs = GetBool(body, "noBaudScan") ?? false;
                        if (ip == null || port == null || baud == null)
                            return (400, Error("Required: ip (string), port (integer), baud (integer)"));
                        return OkCmd(cmd, $"open_ethernetport {ip} {port} {baud} {(nobs ? 1 : 0)}");
                    }
                    break;
            }

            return (405, Error("Method not allowed"));
        }

        // ── /api/v1/device/{netId}/... ────────────────────────────────────────

        private static (int, string) RouteDevice(string method, int netId,
            string[] segs, HttpListenerContext ctx, CommandProcessor cmd)
        {
            string resource = segs[2];
            switch (resource)
            {
                // GET|PUT /api/v1/device/{netId}/clock
                case "clock":
                    if (method == "GET")
                        return OkCmd(cmd, $"read_clock {netId}");
                    if (method == "PUT")
                    {
                        var body = ReadBody(ctx);
                        int? y  = GetInt(body, "year");
                        int? mo = GetInt(body, "month");
                        int? d  = GetInt(body, "day");
                        int? h  = GetInt(body, "hour");
                        int? mi = GetInt(body, "minute");
                        if (y == null || mo == null || d == null || h == null || mi == null)
                            return (400, Error("Required: year, month, day, hour, minute"));
                        return OkCmd(cmd, $"write_clock {netId} {y} {mo} {d} {h} {mi}");
                    }
                    break;

                // POST /api/v1/device/{netId}/start
                case "start":
                    if (method == "POST") return OkCmd(cmd, $"start_program {netId}");
                    break;

                // POST /api/v1/device/{netId}/stop
                case "stop":
                    if (method == "POST") return OkCmd(cmd, $"stop_program {netId}");
                    break;

                // POST /api/v1/device/{netId}/unlock
                case "unlock":
                    if (method == "POST")
                    {
                        var body    = ReadBody(ctx);
                        string? pw  = GetString(body, "password");
                        if (pw == null) return (400, Error("Required: password (string)"));
                        return OkCmd(cmd, $"unlock_device {netId} {pw}");
                    }
                    break;

                // POST /api/v1/device/{netId}/lock
                case "lock":
                    if (method == "POST") return OkCmd(cmd, $"lock_device {netId}");
                    break;

                // GET|PUT /api/v1/device/{netId}/objects/{obj}/{index}
                case "objects":
                    if (segs.Length >= 5 &&
                        int.TryParse(segs[3], out int obj) &&
                        int.TryParse(segs[4], out int idx))
                    {
                        if (method == "GET")
                            return OkCmd(cmd, $"read_object_value {netId} {obj} {idx}");
                        if (method == "PUT")
                            return RouteWriteObject(netId, obj, idx, ctx, cmd);
                    }
                    break;

                // GET|PUT /api/v1/device/{netId}/yeartimeswitch/{switchIdx}/{channel}
                case "yeartimeswitch":
                    if (segs.Length >= 5 &&
                        int.TryParse(segs[3], out int ytsIdx) &&
                        int.TryParse(segs[4], out int ytsCh))
                    {
                        if (method == "GET")
                            return OkCmd(cmd, $"read_channel_yeartimeswitch {netId} {ytsIdx} {ytsCh}");
                        if (method == "PUT")
                            return RouteWriteYearTimeSwitch(netId, ytsIdx, ytsCh, ctx, cmd);
                    }
                    break;

                // GET|PUT /api/v1/device/{netId}/7daytimeswitch/{switchIdx}/{channel}
                case "7daytimeswitch":
                    if (segs.Length >= 5 &&
                        int.TryParse(segs[3], out int dtsIdx) &&
                        int.TryParse(segs[4], out int dtsCh))
                    {
                        if (method == "GET")
                            return OkCmd(cmd, $"read_channel_7daytimeswitch {netId} {dtsIdx} {dtsCh}");
                        if (method == "PUT")
                            return RouteWrite7DayTimeSwitch(netId, dtsIdx, dtsCh, ctx, cmd);
                    }
                    break;
            }

            return (404, Error($"Unknown device resource: {resource}"));
        }

        // ── Write helpers ─────────────────────────────────────────────────────

        private static (int, string) RouteWriteObject(int netId, int obj, int idx,
            HttpListenerContext ctx, CommandProcessor cmd)
        {
            var body = ReadBody(ctx);
            if (body == null) return (400, Error("JSON body required"));

            int? length = GetInt(body, "length");
            if (length == null) return (400, Error("Required: length (integer)"));

            // Pulse mode: { "length": 1, "pulse": { "on": 1, "off": 0, "ms": 500 } }
            if (body.ContainsKey("pulse"))
            {
                var pulse = body["pulse"]?.AsObject();
                int? on  = GetInt(pulse, "on");
                int? off = GetInt(pulse, "off");
                int? ms  = GetInt(pulse, "ms");
                if (on == null || off == null || ms == null)
                    return (400, Error("pulse requires: on (int), off (int), ms (int)"));
                return OkCmd(cmd,
                    $"write_object_value {netId} {obj} {idx} {length} {on}|{off}|{ms}");
            }

            // Normal write: { "length": 1, "value": 1 }
            if (!body.ContainsKey("value"))
                return (400, Error("Required: value (integer) or pulse (object)"));
            int? value = GetInt(body, "value");
            if (value == null) return (400, Error("'value' must be an integer"));
            return OkCmd(cmd, $"write_object_value {netId} {obj} {idx} {length} {value}");
        }

        private static (int, string) RouteWriteYearTimeSwitch(int netId, int idx, int ch,
            HttpListenerContext ctx, CommandProcessor cmd)
        {
            var body = ReadBody(ctx);
            if (body == null) return (400, Error("JSON body required"));
            int? onY  = GetInt(body, "onYear");
            int? onM  = GetInt(body, "onMonth");
            int? onD  = GetInt(body, "onDay");
            int? offY = GetInt(body, "offYear");
            int? offM = GetInt(body, "offMonth");
            int? offD = GetInt(body, "offDay");
            if (onY == null || onM == null || onD == null || offY == null || offM == null || offD == null)
                return (400, Error("Required: onYear, onMonth, onDay, offYear, offMonth, offDay"));
            return OkCmd(cmd,
                $"write_channel_yeartimeswitch {netId} {idx} {ch} {onY} {onM} {onD} {offY} {offM} {offD}");
        }

        private static (int, string) RouteWrite7DayTimeSwitch(int netId, int idx, int ch,
            HttpListenerContext ctx, CommandProcessor cmd)
        {
            var body = ReadBody(ctx);
            if (body == null) return (400, Error("JSON body required"));
            int? dy1    = GetInt(body, "dayMask1");
            int? dy2    = GetInt(body, "dayMask2");
            int? onH    = GetInt(body, "onHour");
            int? onMin  = GetInt(body, "onMinute");
            int? offH   = GetInt(body, "offHour");
            int? offMin = GetInt(body, "offMinute");
            if (dy1 == null || dy2 == null || onH == null || onMin == null || offH == null || offMin == null)
                return (400, Error("Required: dayMask1, dayMask2, onHour, onMinute, offHour, offMinute"));
            return OkCmd(cmd,
                $"write_channel_7daytimeswitch {netId} {idx} {ch} {dy1} {dy2} {onH} {onMin} {offH} {offMin}");
        }

        // ── Response builders ─────────────────────────────────────────────────

        private static (int, string) OkCmd(CommandProcessor cmd, string command)
        {
            string result = cmd.Execute(command);
            if (result.StartsWith("ERROR"))
                return (400, Error(result[5..].TrimStart()));
            return (200, Ok(result.Trim()));
        }

        private static string Ok(string result) =>
            new JsonObject { ["ok"] = true, ["result"] = result }.ToJsonString();

        private static string Error(string message) =>
            new JsonObject { ["ok"] = false, ["error"] = message }.ToJsonString();

        private static void WriteResponse(HttpListenerContext ctx, int status, string json)
        {
            byte[] buf = Encoding.UTF8.GetBytes(json);
            ctx.Response.StatusCode    = status;
            ctx.Response.ContentType   = "application/json; charset=utf-8";
            ctx.Response.ContentLength64 = buf.Length;
            ctx.Response.OutputStream.Write(buf, 0, buf.Length);
            ctx.Response.Close();
        }

        // ── JSON body helpers ─────────────────────────────────────────────────

        private static JsonObject? ReadBody(HttpListenerContext ctx)
        {
            try
            {
                using var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8);
                string json = reader.ReadToEnd();
                if (string.IsNullOrWhiteSpace(json)) return null;
                return JsonNode.Parse(json)?.AsObject();
            }
            catch { return null; }
        }

        private static int? GetInt(JsonObject? obj, string key)
        {
            if (obj == null || !obj.ContainsKey(key)) return null;
            try { return obj[key]?.GetValue<int>(); }
            catch { return null; }
        }

        private static string? GetString(JsonObject? obj, string key)
        {
            if (obj == null || !obj.ContainsKey(key)) return null;
            return obj[key]?.GetValue<string>();
        }

        private static bool? GetBool(JsonObject? obj, string key)
        {
            if (obj == null || !obj.ContainsKey(key)) return null;
            try { return obj[key]?.GetValue<bool>(); }
            catch { return null; }
        }
    }
}
