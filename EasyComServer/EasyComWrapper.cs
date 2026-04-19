using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace EasyComServer
{
    /// <summary>
    /// P/Invoke wrapper for EASY_COM.dll (Delphi/Pascal 32-bit DLL).
    ///
    /// IMPORTANT: Delphi pushes ALL parameters as 32-bit integers on the stack,
    /// regardless of the declared type (byte, word, integer).
    /// Therefore ALL P/Invoke parameters must be declared as 'int', not byte/ushort.
    /// Using byte/ushort causes stack misalignment and 0xC0000005 Access Violations.
    ///
    /// Write_Object_Value and MC_Write_Object_Value expect the value parameter
    /// as a pointer (ref int) — confirmed by disassembly (TEST EDI,EDI null check).
    /// </summary>
    public class EasyComWrapper : IDisposable
    {
        private readonly string _dllPath;

        // ── P/Invoke: Single-connection ───────────────────────────────────────
        // ALL parameters are int (Delphi pushes everything as 32-bit on stack)

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Open_ComPort")]
        private static extern int Open_ComPort(int comPortNr, int baudRate);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Close_ComPort")]
        private static extern int Close_ComPort();

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "GetCurrent_Baudrate")]
        private static extern int GetCurrent_Baudrate(ref int baudRate);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Set_UserWaitingTime")]
        private static extern int Set_UserWaitingTime(int timeoutMs);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Get_UserWaitingTime")]
        private static extern int Get_UserWaitingTime(ref int timeoutMs);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Open_EthernetPort", CharSet = CharSet.Ansi)]
        private static extern int Open_EthernetPort(string ipAddress, int ipPort, int baudRate, int noBaudRateScan);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Close_EthernetPort")]
        private static extern int Close_EthernetPort();

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Start_Program")]
        private static extern int Start_Program(int netId, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Stop_Program")]
        private static extern int Stop_Program(int netId, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Read_Clock")]
        private static extern int Read_Clock(int netId,
            ref int year, ref int month, ref int day, ref int hour, ref int min);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Write_Clock")]
        private static extern int Write_Clock(int netId,
            int year, int month, int day, int hour, int min);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Read_Object_Value")]
        private static extern int Read_Object_Value(int netId, int obj, int index, ref int value);

        // Last parameter is a pointer (ref) — confirmed by disassembly (TEST EDI,EDI)
        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Write_Object_Value")]
        private static extern int Write_Object_Value(int netId, int obj, int index, int length, ref int value);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Read_Channel_YearTimeSwitch")]
        private static extern int Read_Channel_YearTimeSwitch(int netId, int index, int channel,
            ref int onYear, ref int onMonth, ref int onDay,
            ref int offYear, ref int offMonth, ref int offDay);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Write_Channel_YearTimeSwitch")]
        private static extern int Write_Channel_YearTimeSwitch(int netId, int index, int channel,
            int onYear, int onMonth, int onDay, int offYear, int offMonth, int offDay);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Read_Channel_7DayTimeSwitch")]
        private static extern int Read_Channel_7DayTimeSwitch(int netId, int index, int channel,
            ref int dy1, ref int dy2, ref int onHour, ref int onMin, ref int offHour, ref int offMin);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Write_Channel_7DayTimeSwitch")]
        private static extern int Write_Channel_7DayTimeSwitch(int netId, int index, int channel,
            int dy1, int dy2, int onHour, int onMin, int offHour, int offMin);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Unlock_Device", CharSet = CharSet.Ansi)]
        private static extern int Unlock_Device(int netId,
            [MarshalAs(UnmanagedType.LPStr)] string password, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "Lock_Device")]
        private static extern int Lock_Device(int netId, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "GetLastSystemError")]
        private static extern int GetLastSystemError_Native();

        // ── P/Invoke: Multi-connection (MC_) ──────────────────────────────────

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Open_ComPort")]
        private static extern int MC_Open_ComPort(ref int handle, int comPortNr, int baudRate);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Close_ComPort")]
        private static extern int MC_Close_ComPort(int handle);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_GetCurrent_Baudrate")]
        private static extern int MC_GetCurrent_Baudrate(int handle, ref int baudRate);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Set_UserWaitingTime")]
        private static extern int MC_Set_UserWaitingTime(int handle, int timeoutMs);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Get_UserWaitingTime")]
        private static extern int MC_Get_UserWaitingTime(int handle, ref int timeoutMs);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Open_EthernetPort", CharSet = CharSet.Ansi)]
        private static extern int MC_Open_EthernetPort(int handle, string ipAddress, int ipPort,
            int baudRate, int noBaudRateScan);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Close_EthernetPort")]
        private static extern int MC_Close_EthernetPort(int handle);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_CloseAll")]
        private static extern int MC_CloseAll();

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Start_Program")]
        private static extern int MC_Start_Program(int handle, int netId, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Stop_Program")]
        private static extern int MC_Stop_Program(int handle, int netId, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Read_Clock")]
        private static extern int MC_Read_Clock(int handle, int netId,
            ref int year, ref int month, ref int day, ref int hour, ref int min);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Write_Clock")]
        private static extern int MC_Write_Clock(int handle, int netId,
            int year, int month, int day, int hour, int min);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Read_Object_Value")]
        private static extern int MC_Read_Object_Value(int handle, int netId, int obj,
            int index, ref int value);

        // Last parameter is a pointer (ref) — confirmed by disassembly (TEST EDI,EDI)
        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Write_Object_Value")]
        private static extern int MC_Write_Object_Value(int handle, int netId, int obj,
            int index, int length, ref int value);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Read_Channel_YearTimeSwitch")]
        private static extern int MC_Read_Channel_YearTimeSwitch(int handle, int netId,
            int index, int channel,
            ref int onYear, ref int onMonth, ref int onDay,
            ref int offYear, ref int offMonth, ref int offDay);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Write_Channel_YearTimeSwitch")]
        private static extern int MC_Write_Channel_YearTimeSwitch(int handle, int netId,
            int index, int channel,
            int onYear, int onMonth, int onDay, int offYear, int offMonth, int offDay);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Read_Channel_7DayTimeSwitch")]
        private static extern int MC_Read_Channel_7DayTimeSwitch(int handle, int netId,
            int index, int channel,
            ref int dy1, ref int dy2, ref int onHour, ref int onMin,
            ref int offHour, ref int offMin);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Write_Channel_7DayTimeSwitch")]
        private static extern int MC_Write_Channel_7DayTimeSwitch(int handle, int netId,
            int index, int channel,
            int dy1, int dy2, int onHour, int onMin, int offHour, int offMin);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Unlock_Device", CharSet = CharSet.Ansi)]
        private static extern int MC_Unlock_Device(int handle, int netId,
            [MarshalAs(UnmanagedType.LPStr)] string password, ref int errorDetail);

        [DllImport("EASY_COM.dll", CallingConvention = CallingConvention.StdCall, EntryPoint = "MC_Lock_Device")]
        private static extern int MC_Lock_Device(int handle, int netId, ref int errorDetail);

        // ── Error code lookup ─────────────────────────────────────────────────

        private static readonly Dictionary<int, string> EasyErrors = new()
        {
            {  0, "OK" },
            {  1, "ERROR Invalid parameter" },
            {  2, "ERROR Device adverse response" },
            {  3, "ERROR Device no response" },
            {  4, "ERROR Device password locked" },
            {  5, "ERROR Device type unknown" },
            {  6, "ERROR Windows system error" },
            {  7, "ERROR COM port not open" },
            {  8, "ERROR COM port does not exist" },
            {  9, "ERROR COM port cannot be accessed" },
            { 10, "ERROR COM general error" },
            { 11, "ERROR Internal error" },
            { 12, "ERROR Function block does not exist" },
            { 13, "ERROR Function block no const parameter" },
            { 14, "ERROR TCP port does not exist" },
            { 15, "ERROR TCP port no baudrate scan possible" },
            { 16, "ERROR Handle invalid" },
        };

        private static string EasyErrorString(int code)
        {
            if (EasyErrors.TryGetValue(code, out string? msg)) return msg;
            return $"ERROR Unknown ({code})";
        }

        // ── Connection state ──────────────────────────────────────────────────

        private readonly object _lock = new();
        private bool _singleConnected = false;
        private DateTime _singleLastActivity = DateTime.MinValue;
        private readonly Dictionary<int, ConnectionInfo> _connections = new();
        private Timer? _idleTimer;
        private int _comIdleSeconds;
        private int _defaultComPort = 0;
        private int _defaultBaudRate = 9600;
        private int _connectedComPort = 0;   // actual port currently open
        private int _connectedBaudRate = 0;  // actual baud rate currently open

        public EasyComWrapper(string dllPath)
        {
            _dllPath = dllPath;
        }

        /// <summary>Sets the default COM port used for auto-connect.</summary>
        public void SetDefaultCom(int comPort, int baudRate)
        {
            _defaultComPort = comPort;
            _defaultBaudRate = baudRate;
            Logger.Log($"Default COM config: COM{comPort} @ {baudRate} baud.");
        }

        /// <summary>
        /// Updates the default COM port for the next auto-connect.
        /// Called by CommandProcessor before each command to apply
        /// the instance-specific port. Has no effect if a connection
        /// is already open.
        /// </summary>
        public void SetInstanceCom(int comPort, int baudRate)
        {
            lock (_lock)
            {
                if (!_singleConnected)
                {
                    _defaultComPort = comPort;
                    _defaultBaudRate = baudRate;
                }
            }
        }

        /// <summary>
        /// Configures the idle timeout. After this many seconds without activity
        /// the COM port is closed automatically and re-opened on the next command.
        /// Pass 0 to disable auto-close.
        /// </summary>
        public void SetComIdleTimeout(int seconds)
        {
            _comIdleSeconds = seconds;
            _idleTimer?.Dispose();
            _idleTimer = null;
            if (seconds > 0)
            {
                _idleTimer = new Timer(CheckIdleConnections, null,
                    TimeSpan.FromSeconds(10), TimeSpan.FromSeconds(10));
                Logger.Log($"COM idle timeout set to {seconds}s.");
            }
        }

        private string? EnsureComConnected(int comPortOverride = 0, int baudRateOverride = 0)
        {
            lock (_lock)
            {
                if (_singleConnected)
                {
                    _singleLastActivity = DateTime.Now;
                    return null;
                }

                int comPort = comPortOverride > 0 ? comPortOverride : _defaultComPort;
                int baudRate = baudRateOverride > 0 ? baudRateOverride : _defaultBaudRate;

                if (comPort <= 0)
                    return "ERROR No default COM port configured. " +
                           "Add com_port=N and baud_rate=N to easycom.ini.";

                Logger.Log($"Auto-connecting to COM{comPort} @ {baudRate} baud...");
                int rc = Open_ComPort(comPort, baudRate);
                if (rc != 0)
                    return $"{EasyErrorString(rc)} (auto-connect COM{comPort})";

                _singleConnected = true;
                _singleLastActivity = DateTime.Now;
                _connectedComPort = comPort;
                _connectedBaudRate = baudRate;
                // Update defaults so KeepAlive and idle timer know which port is open
                _defaultComPort = comPort;
                _defaultBaudRate = baudRate;
                Logger.Log($"Auto-connected to COM{comPort} @ {baudRate} baud.");
                return null;
            }
        }

        private void CheckIdleConnections(object? state)
        {
            if (_comIdleSeconds <= 0) return;
            var now = DateTime.Now;
            lock (_lock)
            {
                if (_singleConnected &&
                    (now - _singleLastActivity).TotalSeconds >= _comIdleSeconds)
                {
                    try { Close_ComPort(); } catch { }
                    _singleConnected = false;
                    Logger.Log($"Idle-closed COM{_defaultComPort} " +
                               $"(idle for {_comIdleSeconds}s). Will auto-reconnect on next command.");
                }
                var toClose = new List<int>();
                foreach (var kv in _connections)
                    if (kv.Value.Type == ConnectionType.Com &&
                        (now - kv.Value.LastActivity).TotalSeconds >= _comIdleSeconds)
                        toClose.Add(kv.Key);
                foreach (var h in toClose)
                {
                    try { MC_Close_ComPort(h); } catch { }
                    _connections.Remove(h);
                    Logger.Log($"Idle-closed MC COM handle={h}.");
                }
            }
        }

        public string GetVersion()
        {
            try
            {
                var vi = System.Diagnostics.FileVersionInfo.GetVersionInfo(_dllPath);

                // Always use numeric parts — FileVersion string may contain commas
                // instead of dots (e.g. "2, 4, 2, 2010" from EASY_COM.dll)
                if (vi.FileMajorPart > 0 || vi.FileMinorPart > 0)
                {
                    string ver = $"{vi.FileMajorPart}.{vi.FileMinorPart}.{vi.FileBuildPart}.{vi.FilePrivatePart}";
                    bool supported = vi.FileMajorPart == 2
                                  && vi.FileMinorPart >= 3
                                  && vi.FileMinorPart <= 5;
                    return supported
                        ? $"{ver} (supported)"
                        : $"{ver} (not tested — supported range: 2.3.x–2.5.x)";
                }

                return "unknown";
            }
            catch { return "unknown"; }
        }

        public string GetDefaultComInfo()
            => _defaultComPort > 0
                ? $"COM{_defaultComPort} @ {_defaultBaudRate} baud"
                : "not configured";

        public string GetComInfo()
        {
            lock (_lock)
            {
                if (_singleConnected)
                    return $"COM{_connectedComPort} @ {_connectedBaudRate} baud (connected)";
                if (_defaultComPort > 0)
                    return $"COM{_defaultComPort} @ {_defaultBaudRate} baud (auto-connect on next command)";
                return "not configured";
            }
        }

        public bool IsSingleConnected { get { lock (_lock) return _singleConnected; } }

        /// <summary>
        /// Refreshes the last-activity timestamp to prevent idle-close
        /// during a pulse busy-wait. Call continuously in the wait loop.
        /// </summary>
        public void KeepAlive()
        {
            lock (_lock)
            {
                if (_singleConnected)
                    _singleLastActivity = DateTime.Now;
            }
        }

        private void TouchHandle(int handle)
        {
            lock (_lock)
            {
                if (_connections.TryGetValue(handle, out var ci))
                    ci.LastActivity = DateTime.Now;
            }
        }

        public string GetOpenConnections()
        {
            lock (_lock)
            {
                var sb = new StringBuilder();
                if (_singleConnected)
                    sb.AppendLine($"  COM{_defaultComPort} @ {_defaultBaudRate} baud" +
                                  $"  idle={(DateTime.Now - _singleLastActivity).TotalSeconds:F0}s");
                foreach (var kv in _connections)
                    sb.AppendLine($"  Handle={kv.Key} {kv.Value.Type} {kv.Value.Target}" +
                                  $"  since={kv.Value.OpenedAt:HH:mm:ss}" +
                                  $"  idle={(DateTime.Now - kv.Value.LastActivity).TotalSeconds:F0}s");
                return sb.Length == 0 ? "none" : sb.ToString().TrimEnd();
            }
        }

        // ── Single-connection API ─────────────────────────────────────────────

        public string OpenComPort(int comPortNr, int baudRate)
        {
            lock (_lock)
            {
                if (_singleConnected) { Close_ComPort(); _singleConnected = false; }
                int rc = Open_ComPort(comPortNr, baudRate);
                if (rc == 0)
                {
                    _singleConnected = true;
                    _singleLastActivity = DateTime.Now;
                    _connectedComPort = comPortNr;
                    _connectedBaudRate = baudRate;
                    _defaultComPort = comPortNr;
                    _defaultBaudRate = baudRate;
                    return $"OK OPEN_COMPORT COM{comPortNr} {baudRate}";
                }
                return EasyErrorString(rc);
            }
        }

        public string CloseComPort()
        {
            lock (_lock)
            {
                int rc = Close_ComPort();
                if (rc == 0) { _singleConnected = false; _connectedComPort = 0; _connectedBaudRate = 0; }
                return rc == 0 ? "OK CLOSE_COMPORT" : EasyErrorString(rc);
            }
        }

        public string GetCurrentBaudrate()
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int baud = 0;
            int rc = GetCurrent_Baudrate(ref baud);
            return rc == 0 ? $"OK\r\n{baud}" : EasyErrorString(rc);
        }

        public string SetUserWaitingTime(int ms)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int rc = Set_UserWaitingTime(ms);
            return rc == 0 ? $"OK\r\n{ms}" : EasyErrorString(rc);
        }

        public string GetUserWaitingTime()
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int w = 0;
            int rc = Get_UserWaitingTime(ref w);
            return rc == 0 ? $"OK\r\n{w}" : EasyErrorString(rc);
        }

        public string OpenEthernetPort(string ip, int port, int baud, bool noBaudScan)
        {
            int rc = Open_EthernetPort(ip, port, baud, noBaudScan ? 1 : 0);
            return rc == 0 ? $"OK OPEN_ETHERNETPORT {ip}:{port}" : EasyErrorString(rc);
        }

        public string CloseEthernetPort()
        {
            int rc = Close_EthernetPort();
            return rc == 0 ? "OK CLOSE_ETHERNETPORT" : EasyErrorString(rc);
        }

        public string StartProgram(int netId)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int det = 0;
            int rc = Start_Program(netId, ref det);
            return rc == 0 ? $"OK START_PROGRAM NET_ID={netId}"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string StopProgram(int netId)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int det = 0;
            int rc = Stop_Program(netId, ref det);
            return rc == 0 ? $"OK STOP_PROGRAM NET_ID={netId}"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string ReadClock(int netId)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int yr = 0, mo = 0, dy = 0, h = 0, mi = 0;
            int rc = Read_Clock(netId, ref yr, ref mo, ref dy, ref h, ref mi);
            return rc == 0 ? $"OK\r\n{yr} {mo} {dy} {h} {mi}" : EasyErrorString(rc);
        }

        public string WriteClock(int netId, int year, int month, int day, int hour, int min)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int rc = Write_Clock(netId, year, month, day, hour, min);
            return rc == 0 ? "OK WRITE_CLOCK" : EasyErrorString(rc);
        }

        public string ReadObjectValue(int netId, int obj, int index)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int value = 0;
            int rc = Read_Object_Value(netId, obj, index, ref value);
            return rc == 0 ? $"OK\r\n{value}" : EasyErrorString(rc);
        }

        public string WriteObjectValue(int netId, int obj, int index, int length, int value)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int v = value;
            int rc = Write_Object_Value(netId, obj, index, length, ref v);
            return rc == 0 ? $"OK\r\n{v}" : EasyErrorString(rc);
        }

        /// <summary>
        /// Direct write without EnsureComConnected — for pulse use only.
        /// The connection must already be open. Updates LastActivity to
        /// prevent idle-close during the pulse wait.
        /// </summary>
        internal string WriteObjectValueDirect(int netId, int obj, int index, int length, int value)
        {
            lock (_lock) { _singleLastActivity = DateTime.Now; }
            int v = value;
            int rc = Write_Object_Value(netId, obj, index, length, ref v);
            return rc == 0 ? $"OK\r\n{v}" : EasyErrorString(rc);
        }

        public string ReadChannelYearTimeSwitch(int netId, int index, int channel)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int onY = 0, onM = 0, onD = 0, offY = 0, offM = 0, offD = 0;
            int rc = Read_Channel_YearTimeSwitch(netId, index, channel,
                ref onY, ref onM, ref onD, ref offY, ref offM, ref offD);
            return rc == 0
                ? $"OK\r\n{onY} {onM} {onD} {offY} {offM} {offD}"
                : EasyErrorString(rc);
        }

        public string WriteChannelYearTimeSwitch(int netId, int index, int channel,
            int onY, int onM, int onD, int offY, int offM, int offD)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int rc = Write_Channel_YearTimeSwitch(netId, index, channel,
                onY, onM, onD, offY, offM, offD);
            return rc == 0
                ? $"OK\r\n{onY} {onM} {onD} {offY} {offM} {offD}"
                : EasyErrorString(rc);
        }

        public string ReadChannel7DayTimeSwitch(int netId, int index, int channel)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int dy1 = 0, dy2 = 0, onH = 0, onM = 0, offH = 0, offM = 0;
            int rc = Read_Channel_7DayTimeSwitch(netId, index, channel,
                ref dy1, ref dy2, ref onH, ref onM, ref offH, ref offM);
            return rc == 0
                ? $"OK\r\n{dy1} {dy2} {onH} {onM} {offH} {offM}"
                : EasyErrorString(rc);
        }

        public string WriteChannel7DayTimeSwitch(int netId, int index, int channel,
            int dy1, int dy2, int onH, int onM, int offH, int offM)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int rc = Write_Channel_7DayTimeSwitch(netId, index, channel,
                dy1, dy2, onH, onM, offH, offM);
            return rc == 0
                ? $"OK\r\n{dy1} {dy2} {onH} {onM} {offH} {offM}"
                : EasyErrorString(rc);
        }

        public string UnlockDevice(int netId, string password)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int det = 0;
            int rc = Unlock_Device(netId, password, ref det);
            return rc == 0 ? "OK UNLOCK_DEVICE"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string LockDevice(int netId)
        {
            var err = EnsureComConnected(); if (err != null) return err;
            int det = 0;
            int rc = Lock_Device(netId, ref det);
            return rc == 0 ? "OK LOCK_DEVICE"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string GetLastSysError()
        {
            int n = GetLastSystemError_Native();
            return $"OK\r\n{Marshal.GetPInvokeErrorMessage(n)} ({n})";
        }

        // ── MC_ multi-connection API ──────────────────────────────────────────

        public string McOpenComPort(int comPortNr, int baudRate)
        {
            int handle = 0;
            int rc = MC_Open_ComPort(ref handle, comPortNr, baudRate);
            if (rc == 0)
            {
                lock (_lock)
                {
                    _connections[handle] = new ConnectionInfo
                    {
                        Handle = handle,
                        Type = ConnectionType.Com,
                        Target = $"COM{comPortNr}@{baudRate}",
                        LastActivity = DateTime.Now
                    };
                }
                return $"OK MC_OPEN_COMPORT\r\n{handle}";
            }
            return EasyErrorString(rc);
        }

        public string McCloseComPort(int handle)
        {
            int rc = MC_Close_ComPort(handle);
            if (rc == 0) { lock (_lock) { _connections.Remove(handle); } }
            return rc == 0 ? $"OK MC_CLOSE_COMPORT\r\n{handle}" : EasyErrorString(rc);
        }

        public string McGetCurrentBaudrate(int handle)
        {
            TouchHandle(handle);
            int baud = 0;
            int rc = MC_GetCurrent_Baudrate(handle, ref baud);
            return rc == 0 ? $"OK\r\n{baud}" : EasyErrorString(rc);
        }

        public string McSetUserWaitingTime(int handle, int ms)
        {
            TouchHandle(handle);
            int rc = MC_Set_UserWaitingTime(handle, ms);
            return rc == 0 ? $"OK\r\n{ms}" : EasyErrorString(rc);
        }

        public string McGetUserWaitingTime(int handle)
        {
            TouchHandle(handle);
            int w = 0;
            int rc = MC_Get_UserWaitingTime(handle, ref w);
            return rc == 0 ? $"OK\r\n{w}" : EasyErrorString(rc);
        }

        public string McOpenEthernetPort(int handle, string ip, int port, int baud, bool noBaudScan)
        {
            int rc = MC_Open_EthernetPort(handle, ip, port, baud, noBaudScan ? 1 : 0);
            if (rc == 0)
            {
                lock (_lock)
                {
                    _connections[handle] = new ConnectionInfo
                    {
                        Handle = handle,
                        Type = ConnectionType.Ethernet,
                        Target = $"{ip}:{port}",
                        LastActivity = DateTime.Now
                    };
                }
                return $"OK MC_OPEN_ETHERNETPORT HANDLE={handle}";
            }
            return EasyErrorString(rc);
        }

        public string McCloseEthernetPort(int handle)
        {
            int rc = MC_Close_EthernetPort(handle);
            if (rc == 0) { lock (_lock) { _connections.Remove(handle); } }
            return rc == 0 ? "OK MC_CLOSE_ETHERNETPORT" : EasyErrorString(rc);
        }

        public string McCloseAll()
        {
            int rc = MC_CloseAll();
            if (rc == 0) { lock (_lock) { _connections.Clear(); } }
            return rc == 0 ? "OK MC_CLOSEALL" : EasyErrorString(rc);
        }

        public string McStartProgram(int handle, int netId)
        {
            TouchHandle(handle);
            int det = 0;
            int rc = MC_Start_Program(handle, netId, ref det);
            return rc == 0 ? "OK MC_START_PROGRAM"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string McStopProgram(int handle, int netId)
        {
            TouchHandle(handle);
            int det = 0;
            int rc = MC_Stop_Program(handle, netId, ref det);
            return rc == 0 ? "OK MC_STOP_PROGRAM"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string McReadClock(int handle, int netId)
        {
            TouchHandle(handle);
            int yr = 0, mo = 0, dy = 0, hh = 0, mi = 0;
            int rc = MC_Read_Clock(handle, netId, ref yr, ref mo, ref dy, ref hh, ref mi);
            return rc == 0 ? $"OK\r\n{yr} {mo} {dy} {hh} {mi}" : EasyErrorString(rc);
        }

        public string McWriteClock(int handle, int netId, int year, int month,
            int day, int hour, int min)
        {
            TouchHandle(handle);
            int rc = MC_Write_Clock(handle, netId, year, month, day, hour, min);
            return rc == 0 ? "OK MC_WRITE_CLOCK" : EasyErrorString(rc);
        }

        public string McReadObjectValue(int handle, int netId, int obj, int index)
        {
            TouchHandle(handle);
            int value = 0;
            int rc = MC_Read_Object_Value(handle, netId, obj, index, ref value);
            return rc == 0 ? $"OK\r\n{value}" : EasyErrorString(rc);
        }

        public string McWriteObjectValue(int handle, int netId, int obj, int index,
            int length, int value)
        {
            TouchHandle(handle);
            int v = value;
            int rc = MC_Write_Object_Value(handle, netId, obj, index, length, ref v);
            return rc == 0 ? $"OK\r\n{v}" : EasyErrorString(rc);
        }

        /// <summary>
        /// Direct MC write without TouchHandle — for pulse use only.
        /// Updates LastActivity to prevent idle-close during the pulse wait.
        /// </summary>
        internal string McWriteObjectValueDirect(int handle, int netId, int obj,
            int index, int length, int value)
        {
            lock (_lock)
            {
                if (_connections.TryGetValue(handle, out var ci))
                    ci.LastActivity = DateTime.Now;
            }
            int v = value;
            int rc = MC_Write_Object_Value(handle, netId, obj, index, length, ref v);
            return rc == 0 ? $"OK\r\n{v}" : EasyErrorString(rc);
        }

        public string McReadChannelYearTimeSwitch(int handle, int netId, int index, int channel)
        {
            TouchHandle(handle);
            int onY = 0, onM = 0, onD = 0, offY = 0, offM = 0, offD = 0;
            int rc = MC_Read_Channel_YearTimeSwitch(handle, netId, index, channel,
                ref onY, ref onM, ref onD, ref offY, ref offM, ref offD);
            return rc == 0
                ? $"OK\r\n{onY} {onM} {onD} {offY} {offM} {offD}"
                : EasyErrorString(rc);
        }

        public string McWriteChannelYearTimeSwitch(int handle, int netId, int index, int channel,
            int onY, int onM, int onD, int offY, int offM, int offD)
        {
            TouchHandle(handle);
            int rc = MC_Write_Channel_YearTimeSwitch(handle, netId, index, channel,
                onY, onM, onD, offY, offM, offD);
            return rc == 0
                ? $"OK\r\n{onY} {onM} {onD} {offY} {offM} {offD}"
                : EasyErrorString(rc);
        }

        public string McReadChannel7DayTimeSwitch(int handle, int netId, int index, int channel)
        {
            TouchHandle(handle);
            int dy1 = 0, dy2 = 0, onH = 0, onM = 0, offH = 0, offM = 0;
            int rc = MC_Read_Channel_7DayTimeSwitch(handle, netId, index, channel,
                ref dy1, ref dy2, ref onH, ref onM, ref offH, ref offM);
            return rc == 0
                ? $"OK\r\n{dy1} {dy2} {onH} {onM} {offH} {offM}"
                : EasyErrorString(rc);
        }

        public string McWriteChannel7DayTimeSwitch(int handle, int netId, int index, int channel,
            int dy1, int dy2, int onH, int onM, int offH, int offM)
        {
            TouchHandle(handle);
            int rc = MC_Write_Channel_7DayTimeSwitch(handle, netId, index, channel,
                dy1, dy2, onH, onM, offH, offM);
            return rc == 0
                ? $"OK\r\n{dy1} {dy2} {onH} {onM} {offH} {offM}"
                : EasyErrorString(rc);
        }

        public string McUnlockDevice(int handle, int netId, string password)
        {
            TouchHandle(handle);
            int det = 0;
            int rc = MC_Unlock_Device(handle, netId, password, ref det);
            return rc == 0 ? "OK MC_UNLOCK_DEVICE"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public string McLockDevice(int handle, int netId)
        {
            TouchHandle(handle);
            int det = 0;
            int rc = MC_Lock_Device(handle, netId, ref det);
            return rc == 0 ? "OK MC_LOCK_DEVICE"
                 : rc == 2 ? $"{EasyErrorString(rc)}\r\nDevice error: {det}"
                 : EasyErrorString(rc);
        }

        public void Dispose()
        {
            _idleTimer?.Dispose();
            try
            {
                lock (_lock)
                {
                    if (_singleConnected) { Close_ComPort(); _singleConnected = false; }
                    MC_CloseAll();
                    _connections.Clear();
                }
            }
            catch { }
        }
    }

    public class ConnectionInfo
    {
        public int Handle { get; set; }
        public ConnectionType Type { get; set; }
        public string Target { get; set; } = "";
        public DateTime OpenedAt { get; set; } = DateTime.Now;
        public DateTime LastActivity { get; set; } = DateTime.Now;
    }

    public enum ConnectionType { Com, Ethernet }
}
