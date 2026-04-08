using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyComServer
{
    public class CommandProcessor
    {
        private readonly EasyComWrapper _wrapper;
        private readonly IReadOnlyDictionary<string, EasyComWrapper> _allWrappers;
        private readonly ServerConfig _config;
        private readonly InstanceConfig _instance;
        private readonly DateTime _startTime;
        private readonly string _dllVersion;

        private static readonly string[] TopCommands =
        {
            "open_comport", "close_comport", "getcurrent_baudrate",
            "set_userwaitingtime", "get_userwaitingtime",
            "open_ethernetport", "close_ethernetport",
            "start_program", "stop_program",
            "read_clock", "write_clock",
            "read_object_value", "write_object_value",
            "read_channel_yeartimeswitch", "write_channel_yeartimeswitch",
            "read_channel_7daytimeswitch", "write_channel_7daytimeswitch",
            "unlock_device", "lock_device",
            "mc_open_comport", "mc_close_comport", "mc_getcurrent_baudrate",
            "mc_set_userwaitingtime", "mc_get_userwaitingtime",
            "mc_open_ethernetport", "mc_close_ethernetport", "mc_closeall",
            "mc_start_program", "mc_stop_program",
            "mc_read_clock", "mc_write_clock",
            "mc_read_object_value", "mc_write_object_value",
            "mc_read_channel_yeartimeswitch", "mc_write_channel_yeartimeswitch",
            "mc_read_channel_7daytimeswitch", "mc_write_channel_7daytimeswitch",
            "mc_unlock_device", "mc_lock_device",
            "getlastsystemerror",
            "show", "set", "help", "quit"
        };

        private static readonly string[] ShowSubCommands =
            { "server", "configuration", "tasks", "connections" };
        private static readonly string[] SetSubCommands = { "configuration" };

        private static int _reqCounter = 0;

        public CommandProcessor(EasyComWrapper wrapper, ServerConfig config,
            InstanceConfig instance, DateTime startTime, string dllVersion,
            IReadOnlyDictionary<string, EasyComWrapper>? allWrappers = null)
        {
            _wrapper = wrapper;
            _allWrappers = allWrappers ?? new Dictionary<string, EasyComWrapper> { [instance.Name] = wrapper };
            _config = config;
            _instance = instance;
            _startTime = startTime;
            _dllVersion = dllVersion;
        }

        public string Execute(string commandLine)
        {
            string reqId = $"{System.Threading.Interlocked.Increment(ref _reqCounter) & 0xFFFF:x4}";
            try
            {
                if (string.IsNullOrWhiteSpace(commandLine)) return "";
                var parts = SplitArgs(commandLine);
                if (parts.Length == 0) return "";

                Logger.Log($"[{reqId}] CMD: {commandLine}");

                // Apply the instance-specific COM port before auto-connect kicks in
                if (_instance.HasComConfig)
                    _wrapper.SetInstanceCom(_instance.ComPort, _instance.BaudRate);

                string verb = Resolve(parts[0], TopCommands);
                string[] args = parts.Skip(1).ToArray();

                string result = verb switch
                {
                    "help" => BuildHelp(),
                    "quit" => "Goodbye.",
                    "show" => ExecuteShow(args),
                    "set" => ExecuteSet(args),
                    "getlastsystemerror" => _wrapper.GetLastSysError(),

                    "open_comport" => RequireArgs(args, 2, v =>
                        _wrapper.OpenComPort(Int(v[0]), Int(v[1]))),
                    "close_comport" => _wrapper.CloseComPort(),
                    "getcurrent_baudrate" => _wrapper.GetCurrentBaudrate(),
                    "set_userwaitingtime" => RequireArgs(args, 1, v =>
                        _wrapper.SetUserWaitingTime(Int(v[0]))),
                    "get_userwaitingtime" => _wrapper.GetUserWaitingTime(),
                    "open_ethernetport" => RequireArgs(args, 4, v =>
                        _wrapper.OpenEthernetPort(v[0], Int(v[1]), Int(v[2]),
                            v[3] == "1" || v[3].ToLower() == "true")),
                    "close_ethernetport" => _wrapper.CloseEthernetPort(),
                    "start_program" => RequireArgs(args, 1, v =>
                        _wrapper.StartProgram(Int(v[0]))),
                    "stop_program" => RequireArgs(args, 1, v =>
                        _wrapper.StopProgram(Int(v[0]))),
                    "read_clock" => RequireArgs(args, 1, v =>
                        _wrapper.ReadClock(Int(v[0]))),
                    "write_clock" => RequireArgs(args, 6, v =>
                        _wrapper.WriteClock(Int(v[0]), Int(v[1]), Int(v[2]),
                            Int(v[3]), Int(v[4]), Int(v[5]))),

                    // READ_OBJECT_VALUE NetId Obj Index  (3 params, no Length)
                    "read_object_value" => RequireArgs(args, 3, v =>
                        _wrapper.ReadObjectValue(Int(v[0]), Int(v[1]), Int(v[2]))),

                    // WRITE_OBJECT_VALUE NetId Obj Index Length Data[|Data2|Ms]
                    "write_object_value" => RequireArgs(args, 5, v =>
                        ExecuteWriteObjectValue(reqId,
                            Int(v[0]), Int(v[1]), Int(v[2]), Int(v[3]), v[4])),

                    "read_channel_yeartimeswitch" => RequireArgs(args, 3, v =>
                        _wrapper.ReadChannelYearTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]))),
                    "write_channel_yeartimeswitch" => RequireArgs(args, 9, v =>
                        _wrapper.WriteChannelYearTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]),
                            Int(v[3]), Int(v[4]), Int(v[5]),
                            Int(v[6]), Int(v[7]), Int(v[8]))),
                    "read_channel_7daytimeswitch" => RequireArgs(args, 3, v =>
                        _wrapper.ReadChannel7DayTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]))),
                    "write_channel_7daytimeswitch" => RequireArgs(args, 9, v =>
                        _wrapper.WriteChannel7DayTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]),
                            Int(v[3]), Int(v[4]), Int(v[5]),
                            Int(v[6]), Int(v[7]), Int(v[8]))),
                    "unlock_device" => RequireArgs(args, 2, v =>
                        _wrapper.UnlockDevice(Int(v[0]), v[1])),
                    "lock_device" => RequireArgs(args, 1, v =>
                        _wrapper.LockDevice(Int(v[0]))),

                    "mc_open_comport" => RequireArgs(args, 2, v =>
                        _wrapper.McOpenComPort(Int(v[0]), Int(v[1]))),
                    "mc_close_comport" => RequireArgs(args, 1, v =>
                        _wrapper.McCloseComPort(Int(v[0]))),
                    "mc_getcurrent_baudrate" => RequireArgs(args, 1, v =>
                        _wrapper.McGetCurrentBaudrate(Int(v[0]))),
                    "mc_set_userwaitingtime" => RequireArgs(args, 2, v =>
                        _wrapper.McSetUserWaitingTime(Int(v[0]), Int(v[1]))),
                    "mc_get_userwaitingtime" => RequireArgs(args, 1, v =>
                        _wrapper.McGetUserWaitingTime(Int(v[0]))),
                    "mc_open_ethernetport" => RequireArgs(args, 5, v =>
                        _wrapper.McOpenEthernetPort(Int(v[0]), v[1], Int(v[2]),
                            Int(v[3]), v[4] == "1" || v[4].ToLower() == "true")),
                    "mc_close_ethernetport" => RequireArgs(args, 1, v =>
                        _wrapper.McCloseEthernetPort(Int(v[0]))),
                    "mc_closeall" => _wrapper.McCloseAll(),
                    "mc_start_program" => RequireArgs(args, 2, v =>
                        _wrapper.McStartProgram(Int(v[0]), Int(v[1]))),
                    "mc_stop_program" => RequireArgs(args, 2, v =>
                        _wrapper.McStopProgram(Int(v[0]), Int(v[1]))),
                    "mc_read_clock" => RequireArgs(args, 2, v =>
                        _wrapper.McReadClock(Int(v[0]), Int(v[1]))),
                    "mc_write_clock" => RequireArgs(args, 7, v =>
                        _wrapper.McWriteClock(Int(v[0]), Int(v[1]), Int(v[2]),
                            Int(v[3]), Int(v[4]), Int(v[5]), Int(v[6]))),

                    // MC_READ_OBJECT_VALUE Handle NetId Obj Index  (4 params)
                    "mc_read_object_value" => RequireArgs(args, 4, v =>
                        _wrapper.McReadObjectValue(
                            Int(v[0]), Int(v[1]), Int(v[2]), Int(v[3]))),

                    // MC_WRITE_OBJECT_VALUE Handle NetId Obj Index Length Data[|Data2|Ms]
                    "mc_write_object_value" => RequireArgs(args, 6, v =>
                        ExecuteMcWriteObjectValue(reqId,
                            Int(v[0]), Int(v[1]), Int(v[2]),
                            Int(v[3]), Int(v[4]), v[5])),

                    "mc_read_channel_yeartimeswitch" => RequireArgs(args, 4, v =>
                        _wrapper.McReadChannelYearTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]), Int(v[3]))),
                    "mc_write_channel_yeartimeswitch" => RequireArgs(args, 10, v =>
                        _wrapper.McWriteChannelYearTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]), Int(v[3]),
                            Int(v[4]), Int(v[5]), Int(v[6]),
                            Int(v[7]), Int(v[8]), Int(v[9]))),
                    "mc_read_channel_7daytimeswitch" => RequireArgs(args, 4, v =>
                        _wrapper.McReadChannel7DayTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]), Int(v[3]))),
                    "mc_write_channel_7daytimeswitch" => RequireArgs(args, 10, v =>
                        _wrapper.McWriteChannel7DayTimeSwitch(
                            Int(v[0]), Int(v[1]), Int(v[2]), Int(v[3]),
                            Int(v[4]), Int(v[5]), Int(v[6]),
                            Int(v[7]), Int(v[8]), Int(v[9]))),
                    "mc_unlock_device" => RequireArgs(args, 3, v =>
                        _wrapper.McUnlockDevice(Int(v[0]), Int(v[1]), v[2])),
                    "mc_lock_device" => RequireArgs(args, 2, v =>
                        _wrapper.McLockDevice(Int(v[0]), Int(v[1]))),

                    _ => $"ERROR Unknown command: '{parts[0]}'. Type HELP for a list."
                };

                Logger.Log($"[{reqId}] RESULT: {result.Split('\n')[0]}");
                return result;
            }
            catch (AmbiguousAbbreviationException ex)
            {
                return $"ERROR Ambiguous command '{ex.Prefix}': matches {ex.Matches}";
            }
            catch (Exception ex)
            {
                Logger.Log($"[{reqId}] EXCEPTION in Execute: {ex.Message}");
                return $"ERROR {ex.Message}";
            }
        }

        // ── WRITE_OBJECT_VALUE with pulse support ─────────────────────────────
        //
        // Syntax:  WRITE_OBJECT_VALUE NetId Obj Index Length Data
        //          WRITE_OBJECT_VALUE NetId Obj Index Length Data1|Data2|Ms
        //
        // Pascal original:
        //   params[0]=NetId, params[1]=Obj, params[2]=Index,
        //   params[3]=Length, params[4]=Data[|Data2|Ms]

        private string ExecuteWriteObjectValue(string reqId, int netId, int obj, int index,
            int length, string dataArg)
        {
            Logger.Log($"[{reqId}] WRITE_OBJECT_VALUE: netId={netId} obj={obj} index={index} length={length} dataArg={dataArg}");

            var p = dataArg.Split('|');

            if (p.Length == 3)
            {
                // Pulse mode: Data1|Data2|Ms
                if (!byte.TryParse(p[0], out byte v1))
                    return $"ERROR Invalid data1: {p[0]}";
                if (!byte.TryParse(p[1], out byte v2))
                    return $"ERROR Invalid data2: {p[1]}";
                if (!int.TryParse(p[2], out int ms) || ms < 0)
                    return $"ERROR Invalid milliseconds: {p[2]}";

                Logger.Log($"[{reqId}] PULSE: writing v1={v1}...");
                string r1 = _wrapper.WriteObjectValue(netId, obj, index, length, v1);
                Logger.Log($"[{reqId}] PULSE: v1={v1} result={r1}");
                if (r1.StartsWith("ERROR")) return r1;

                // Schedule v2 asynchronously so the caller is not blocked during the wait.
                // Other requests may proceed in the meantime; serial exclusivity is
                // maintained by the _lock inside EasyComWrapper.
                Logger.Log($"[{reqId}] PULSE: scheduling v2={v2} in {ms}ms (async)...");
                _ = Task.Run(async () =>
                {
                    // Keep the COM connection alive in chunks to prevent idle-close.
                    var deadline = DateTime.Now.AddMilliseconds(ms);
                    while (true)
                    {
                        int remaining = Math.Max(0,
                            (int)(deadline - DateTime.Now).TotalMilliseconds);
                        if (remaining == 0) break;
                        await Task.Delay(Math.Min(5_000, remaining));
                        _wrapper.KeepAlive();
                    }
                    Logger.Log($"[{reqId}] PULSE: writing v2={v2}...");
                    string r2 = _wrapper.WriteObjectValue(netId, obj, index, length, v2);
                    Logger.Log($"[{reqId}] PULSE: v2={v2} result={r2}");
                });

                return $"OK PULSE {v1}->{v2} ({ms}ms) async";
            }

            // Normal write
            if (!byte.TryParse(dataArg, out byte val))
                return $"ERROR Invalid data: {dataArg}";
            string res = _wrapper.WriteObjectValue(netId, obj, index, length, val);
            Logger.Log($"[{reqId}] WRITE: val={val} result={res}");
            return res;
        }

        private string ExecuteMcWriteObjectValue(string reqId, int handle, int netId, int obj,
            int index, int length, string dataArg)
        {
            Logger.Log($"[{reqId}] MC_WRITE_OBJECT_VALUE: handle={handle} netId={netId} obj={obj} index={index} length={length} dataArg={dataArg}");

            var p = dataArg.Split('|');

            if (p.Length == 3)
            {
                if (!byte.TryParse(p[0], out byte v1))
                    return $"ERROR Invalid data1: {p[0]}";
                if (!byte.TryParse(p[1], out byte v2))
                    return $"ERROR Invalid data2: {p[1]}";
                if (!int.TryParse(p[2], out int ms) || ms < 0)
                    return $"ERROR Invalid milliseconds: {p[2]}";

                Logger.Log($"[{reqId}] MC PULSE: writing v1={v1}...");
                string r1 = _wrapper.McWriteObjectValue(handle, netId, obj, index, length, v1);
                Logger.Log($"[{reqId}] MC PULSE: v1={v1} result={r1}");
                if (r1.StartsWith("ERROR")) return r1;

                Logger.Log($"[{reqId}] MC PULSE: scheduling v2={v2} in {ms}ms (async)...");
                int capturedHandle = handle;
                _ = Task.Run(async () =>
                {
                    var deadline = DateTime.Now.AddMilliseconds(ms);
                    while (true)
                    {
                        int remaining = Math.Max(0,
                            (int)(deadline - DateTime.Now).TotalMilliseconds);
                        if (remaining == 0) break;
                        await Task.Delay(Math.Min(5_000, remaining));
                        _wrapper.KeepAlive();
                    }
                    Logger.Log($"[{reqId}] MC PULSE: writing v2={v2}...");
                    string r2 = _wrapper.McWriteObjectValue(
                        capturedHandle, netId, obj, index, length, v2);
                    Logger.Log($"[{reqId}] MC PULSE: v2={v2} result={r2}");
                });

                return $"OK PULSE {v1}->{v2} ({ms}ms) async";
            }

            if (!byte.TryParse(dataArg, out byte val))
                return $"ERROR Invalid data: {dataArg}";
            return _wrapper.McWriteObjectValue(handle, netId, obj, index, length, val);
        }

        // ── SHOW ──────────────────────────────────────────────────────────────

        private string GetAllConnections()
        {
            var sb = new StringBuilder();
            foreach (var kv in _allWrappers)
            {
                string conns = kv.Value.GetOpenConnections();
                if (conns == "none")
                    sb.AppendLine($"[{kv.Key}] none");
                else
                    foreach (var line in conns.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                        sb.AppendLine($"[{kv.Key}] {line.Trim()}");
            }
            return sb.Length == 0 ? "none" : sb.ToString().TrimEnd();
        }

        private string ExecuteShow(string[] args)
        {
            if (args.Length == 0)
                return "Usage: SHOW SERVER | CONFIGURATION <var> | TASKS | CONNECTIONS";
            string sub = Resolve(args[0], ShowSubCommands);
            return sub switch
            {
                "server" => BuildShowServer(),
                "connections" => GetAllConnections(),
                "configuration" => args.Length > 1
                    ? ShowConfigVar(args[1])
                    : "Usage: SHOW CONFIGURATION <variable>",
                "tasks" => "HTTP and Telnet listener threads active.",
                _ => $"ERROR Unknown SHOW subcommand: {args[0]}"
            };
        }

        private string BuildShowServer()
        {
            var elapsed = DateTime.Now - _startTime;
            string e = $"{(int)elapsed.TotalDays} days " +
                       $"{elapsed.Hours:00}:{elapsed.Minutes:00}:{elapsed.Seconds:00}";
            var sb = new StringBuilder();
            sb.AppendLine($"Moeller EASY(r) Server (EasyComServer .NET) {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            sb.AppendLine($"Server directory:       {AppDomain.CurrentDomain.BaseDirectory}");
            sb.AppendLine($"Elapsed time:           {e}");
            sb.AppendLine($"Moeller EasyCom DLL:    {_dllVersion}");
            sb.AppendLine($"Instance:               {_instance.Name}");
            sb.AppendLine($"HTTP port:              {(_instance.HttpEnabled ? _instance.HttpPort.ToString() : "disabled")}");
            sb.AppendLine($"Telnet port:            {(_instance.TelnetEnabled ? _instance.TelnetPort.ToString() : "disabled")}");
            sb.AppendLine($"COM idle timeout:       {_config.ComIdleTimeoutSeconds}s");
            sb.AppendLine($"Console Logging:        {(_config.ConsoleLogging ? "Enabled" : "Disabled")}");
            string conns = _wrapper.GetOpenConnections();
            if (conns == "none" && _wrapper.GetDefaultComInfo() != "not configured")
                conns = $"none (configured: {_wrapper.GetDefaultComInfo()})";
            sb.Append($"Open Connections:       {conns}");
            return sb.ToString().TrimEnd();
        }

        private string ShowConfigVar(string name) =>
            name.ToLower() switch
            {
                "dll_path" or "dllpath" => $"dll_path={_config.DllPath}",
                "com_idle_timeout" => $"com_idle_timeout={_config.ComIdleTimeoutSeconds}",
                "log_file" => $"log_file={_config.LogFile}",
                "console_logging" => $"console_logging={_config.ConsoleLogging}",
                "com_port" => $"com_port={_instance.ComPort}",
                "baud_rate" => $"baud_rate={_instance.BaudRate}",
                _ => $"ERROR Unknown variable: {name}"
            };

        // ── SET ───────────────────────────────────────────────────────────────

        private string ExecuteSet(string[] args)
        {
            if (args.Length < 2) return "Usage: SET CONFIGURATION \"variable=value\"";
            string sub = Resolve(args[0], SetSubCommands);
            if (sub != "configuration") return $"ERROR Unknown SET subcommand: {args[0]}";
            string kv = string.Join(" ", args.Skip(1)).Trim('"', ' ');
            int idx = kv.IndexOf('=');
            if (idx < 0) return "ERROR Expected variable=value";
            string key = kv[..idx].Trim().ToLower();
            string val = kv[(idx + 1)..].Trim();
            return key switch
            {
                "console_logging" => SetConsoleLogging(val),
                "com_idle_timeout" => SetComIdleTimeout(val),
                "com_port" => SetDefaultCom(val, null),
                "baud_rate" => SetDefaultCom(null, val),
                _ => $"ERROR Unknown configuration variable: {key}"
            };
        }

        private string SetConsoleLogging(string val)
        {
            bool on = val.Equals("true", StringComparison.OrdinalIgnoreCase)
                   || val == "1"
                   || val.Equals("enabled", StringComparison.OrdinalIgnoreCase);
            _config.ConsoleLogging = on;
            Logger.SetConsole(on);
            return $"OK console_logging={on}";
        }

        private string SetComIdleTimeout(string val)
        {
            if (!int.TryParse(val, out int sec))
                return "ERROR Expected integer seconds";
            _config.ComIdleTimeoutSeconds = sec;
            _wrapper.SetComIdleTimeout(sec);
            return $"OK com_idle_timeout={sec}";
        }

        private string SetDefaultCom(string? port, string? baud)
        {
            int comPort = port != null && int.TryParse(port, out int p)
                ? p : _instance.ComPort;
            int baudRate = baud != null && int.TryParse(baud, out int b)
                ? b : _instance.BaudRate;
            _instance.ComPort = comPort;
            _instance.BaudRate = baudRate;
            _wrapper.SetDefaultCom(comPort, baudRate);
            return $"OK com_port={comPort} baud_rate={baudRate}";
        }

        // ── HELP ──────────────────────────────────────────────────────────────

        private static string BuildHelp() =>
            "EasyComServer - Command Reference\r\n" +
            "Abbreviations allowed (e.g. SH SER = SHOW SERVER)\r\n" +
            "\r\n" +
            "--- Connection ---\r\n" +
            "OPEN_COMPORT <port> <baud>\r\n" +
            "CLOSE_COMPORT\r\n" +
            "GETCURRENT_BAUDRATE\r\n" +
            "SET_USERWAITINGTIME <ms>\r\n" +
            "GET_USERWAITINGTIME\r\n" +
            "OPEN_ETHERNETPORT <ip> <port> <baud> <nobaudscan>\r\n" +
            "CLOSE_ETHERNETPORT\r\n" +
            "GETLASTSYSTEMERROR\r\n" +
            "\r\n" +
            "--- Device ---\r\n" +
            "START_PROGRAM <netid>\r\n" +
            "STOP_PROGRAM <netid>\r\n" +
            "READ_CLOCK <netid>\r\n" +
            "WRITE_CLOCK <netid> <y> <m> <d> <h> <min>\r\n" +
            "UNLOCK_DEVICE <netid> <password>\r\n" +
            "LOCK_DEVICE <netid>\r\n" +
            "\r\n" +
            "--- Process image ---\r\n" +
            "READ_OBJECT_VALUE <netid> <obj> <index>\r\n" +
            "WRITE_OBJECT_VALUE <netid> <obj> <index> <len> <data>\r\n" +
            "WRITE_OBJECT_VALUE ... <d1>|<d2>|<ms>  (pulse)\r\n" +
            "  e.g. WRITE_OBJECT_VALUE 1 4 1 1 1|0|10\r\n" +
            "\r\n" +
            "--- Time switches ---\r\n" +
            "READ_CHANNEL_YEARTIMESWITCH <netid> <idx> <ch>\r\n" +
            "WRITE_CHANNEL_YEARTIMESWITCH <netid> <idx> <ch>\r\n" +
            "  <onY> <onM> <onD> <offY> <offM> <offD>\r\n" +
            "READ_CHANNEL_7DAYTIMESWITCH <netid> <idx> <ch>\r\n" +
            "WRITE_CHANNEL_7DAYTIMESWITCH <netid> <idx> <ch>\r\n" +
            "  <dy1> <dy2> <onH> <onM> <offH> <offM>\r\n" +
            "\r\n" +
            "--- MC_ multi-connection ---\r\n" +
            "MC_OPEN_COMPORT <port> <baud>  -> HANDLE\r\n" +
            "MC_CLOSE_COMPORT <handle>\r\n" +
            "MC_CLOSEALL\r\n" +
            "MC_OPEN_ETHERNETPORT <h> <ip> <port> <baud> <nobs>\r\n" +
            "MC_CLOSE_ETHERNETPORT <handle>\r\n" +
            "MC_GETCURRENT_BAUDRATE <handle>\r\n" +
            "MC_SET_USERWAITINGTIME <handle> <ms>\r\n" +
            "MC_GET_USERWAITINGTIME <handle>\r\n" +
            "MC_START_PROGRAM <handle> <netid>\r\n" +
            "MC_STOP_PROGRAM <handle> <netid>\r\n" +
            "MC_READ_CLOCK <handle> <netid>\r\n" +
            "MC_WRITE_CLOCK <h> <netid> <y> <m> <d> <h> <min>\r\n" +
            "MC_READ_OBJECT_VALUE <h> <netid> <obj> <index>\r\n" +
            "MC_WRITE_OBJECT_VALUE <h> <netid> <obj> <idx> <len> <data>[|d2|ms]\r\n" +
            "MC_UNLOCK_DEVICE <handle> <netid> <password>\r\n" +
            "MC_LOCK_DEVICE <handle> <netid>\r\n" +
            "\r\n" +
            "--- Server ---\r\n" +
            "SHOW SERVER\r\n" +
            "SHOW CONNECTIONS\r\n" +
            "SHOW CONFIGURATION <var>\r\n" +
            "SET CONFIGURATION \"var=value\"\r\n" +
            "  HELP\n" +
            "  QUIT";

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string Resolve(string input, string[] candidates)
        {
            string lower = input.ToLower();
            if (candidates.Contains(lower)) return lower;
            var matches = candidates.Where(c => c.StartsWith(lower)).ToList();
            if (matches.Count == 1) return matches[0];
            if (matches.Count > 1)
                throw new AmbiguousAbbreviationException(input,
                    string.Join(", ", matches));
            return lower;
        }

        private static string RequireArgs(string[] args, int count,
            Func<string[], string> fn)
        {
            if (args.Length < count)
                return $"ERROR Requires {count} argument(s), got {args.Length}.";
            return fn(args);
        }

        private static int Int(string s)
        {
            if (s.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                return Convert.ToInt32(s, 16);
            return int.Parse(s);
        }

        private static string[] SplitArgs(string line)
        {
            var args = new List<string>();
            var current = new StringBuilder();
            bool inQuote = false;
            foreach (char c in line)
            {
                if (c == '"') { inQuote = !inQuote; continue; }
                if (c == ' ' && !inQuote)
                {
                    if (current.Length > 0) { args.Add(current.ToString()); current.Clear(); }
                }
                else current.Append(c);
            }
            if (current.Length > 0) args.Add(current.ToString());
            return args.ToArray();
        }
    }

    public class AmbiguousAbbreviationException : Exception
    {
        public string Prefix { get; }
        public string Matches { get; }
        public AmbiguousAbbreviationException(string prefix, string matches)
            : base($"Ambiguous: '{prefix}' matches {matches}")
        { Prefix = prefix; Matches = matches; }
    }
}
