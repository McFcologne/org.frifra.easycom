using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EasyComConfigurator
{
    // ── Data model ────────────────────────────────────────────────────────────

    internal class GlobalCfg
    {
        public string DllPath        { get; set; } = "EASY_COM.dll";
        public string LogFile        { get; set; } = "logs/easycom.log";
        public bool   ConsoleLogging { get; set; } = true;
        public int    LogMaxSizeMb   { get; set; } = 10;
        public int    LogMaxFiles    { get; set; } = 10;
        public int    ComIdleTimeout { get; set; } = 300;
        public bool   BasicAuth      { get; set; } = false;
        public string AuthUser       { get; set; } = "admin";
        public string AuthPass       { get; set; } = "";

        public GlobalCfg Clone() => (GlobalCfg)MemberwiseClone();
    }

    internal class InstanceCfg
    {
        public string Name          { get; set; } = "default";
        public bool   HttpEnabled   { get; set; } = true;
        public int    HttpPort      { get; set; } = 8083;
        public bool   TelnetEnabled { get; set; } = true;
        public int    TelnetPort    { get; set; } = 8023;
        public int    ComPort       { get; set; } = 1;
        public int    BaudRate      { get; set; } = 9600;

        public InstanceCfg Clone() => (InstanceCfg)MemberwiseClone();
    }

    // ── Change classification ─────────────────────────────────────────────────

    internal record Change(string Label, string Key, string Value);

    // ── Main form ─────────────────────────────────────────────────────────────

    public class MainForm : Form
    {
        private const string SvcName = "EasyComServer";

        // Live-injectable via SET CONFIGURATION (no restart needed)
        private static readonly HashSet<string> LiveGlobal = new()
            { "console_logging", "com_idle_timeout", "basic_auth", "auth_user", "auth_pass" };
        private static readonly HashSet<string> LiveInstance = new()
            { "com_port", "baud_rate" };

        // ── Model ─────────────────────────────────────────────────────────────
        private string _iniPath = "";
        private GlobalCfg _global = new();
        private readonly List<InstanceCfg> _instances = new();

        // Snapshot of the last loaded/saved state for change detection
        private GlobalCfg _snap = new();
        private readonly List<InstanceCfg> _snapInstances = new();

        // ── Controls ──────────────────────────────────────────────────────────
        private TextBox _pathBox     = null!;
        private ListBox _list        = null!;
        private Panel   _globalPanel = null!;
        private Panel   _instPanel   = null!;
        private Label   _lblSvcStatus = null!;

        // Global fields
        private TextBox       _gDllPath    = null!;
        private TextBox       _gLogFile    = null!;
        private NumericUpDown _gLogSizeMb  = null!;
        private NumericUpDown _gLogFiles   = null!;
        private CheckBox      _gConsoleLog = null!;
        private NumericUpDown _gIdleTime   = null!;
        private CheckBox      _gBasicAuth  = null!;
        private TextBox       _gAuthUser   = null!;
        private TextBox       _gAuthPass   = null!;

        // Instance fields
        private TextBox       _iName       = null!;
        private CheckBox      _iHttpOn     = null!;
        private NumericUpDown _iHttpPort   = null!;
        private CheckBox      _iTelnetOn   = null!;
        private NumericUpDown _iTelnetPort = null!;
        private NumericUpDown _iComPort    = null!;
        private ComboBox      _iBaud       = null!;

        private bool _loading  = false;
        private int  _activeIdx = 0;

        private readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(3) };

        // ─────────────────────────────────────────────────────────────────────

        public MainForm()
        {
            BuildUi();

            string exe = AppDomain.CurrentDomain.BaseDirectory;
            foreach (var candidate in new[]
            {
                Path.Combine(exe, "easycom.ini"),
                Path.Combine(exe, @"..\EasyComServer\bin\Debug\net8.0-windows\easycom.ini"),
                Path.Combine(exe, @"..\EasyComServer\bin\Release\net8.0-windows\easycom.ini"),
                Path.Combine(exe, @"..\EasyComServer\easycom.ini"),
            })
            {
                string full = Path.GetFullPath(candidate);
                if (File.Exists(full)) { TryLoad(full); break; }
            }

            RefreshServiceStatus();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _http.Dispose();
            base.Dispose(disposing);
        }

        // ── UI construction ───────────────────────────────────────────────────

        private void BuildUi()
        {
            Text          = "EasyComServer Konfigurator";
            Size          = new Size(860, 620);
            MinimumSize   = new Size(700, 520);
            Font          = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterScreen;
            string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "easycom.ico");
            if (File.Exists(icoPath)) Icon = new Icon(icoPath);

            // Top: ini path
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 46 };
            var lblPath = new Label { Text = "Konfigurationsdatei:", AutoSize = true, Location = new Point(8, 15) };
            var btnBrowse = new Button { Text = "…", Width = 32, Height = 23, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            _pathBox = new TextBox { Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, Location = new Point(168, 12) };
            btnBrowse.Click  += (_, __) => BrowseOpen();
            _pathBox.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) TryLoad(_pathBox.Text.Trim()); };
            pnlTop.Controls.AddRange(new Control[] { lblPath, _pathBox, btnBrowse });
            pnlTop.Resize += (_, __) =>
            {
                _pathBox.Width = pnlTop.Width - 168 - 44;
                btnBrowse.Left = pnlTop.Width - 40;
                btnBrowse.Top  = 11;
            };

            // Bottom: service + save
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 88 };
            var grpSvc = new GroupBox { Text = "Windows-Dienst", Location = new Point(8, 6), Size = new Size(500, 52) };
            _lblSvcStatus = new Label { Text = "…", AutoSize = true, Location = new Point(8, 22) };
            var btnStart   = Btn("▶ Start",    new Point(130, 18), 80);
            var btnStop    = Btn("■ Stop",     new Point(214, 18), 80);
            var btnRestart = Btn("↺ Neustart", new Point(298, 18), 90);
            var btnRefresh = Btn("↻",          new Point(394, 18), 30);
            btnRefresh.Font    = new Font(btnRefresh.Font!.FontFamily, 8f);
            btnStart.Click    += (_, __) => ServiceAction(s => s.Start());
            btnStop.Click     += (_, __) => ServiceAction(s => s.Stop());
            btnRestart.Click  += (_, __) => ServiceAction(s =>
            {
                if (s.Status == ServiceControllerStatus.Running)
                { s.Stop(); s.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(15)); }
                s.Start();
            });
            btnRefresh.Click += (_, __) => RefreshServiceStatus();
            grpSvc.Controls.AddRange(new Control[] { _lblSvcStatus, btnStart, btnStop, btnRestart, btnRefresh });

            var btnSave  = new Button { Text = "💾  Speichern", Size = new Size(130, 30), Anchor = AnchorStyles.Bottom | AnchorStyles.Right, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            var btnClose = new Button { Text = "Schließen",     Size = new Size(90,  30), Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            btnSave.Click  += async (_, __) => await SaveAndApply();
            btnClose.Click += (_, __) => Close();
            pnlBottom.Controls.AddRange(new Control[] { grpSvc, btnSave, btnClose });
            pnlBottom.Resize += (_, __) =>
            {
                btnSave.Location  = new Point(pnlBottom.Width - 234, 28);
                btnClose.Location = new Point(pnlBottom.Width - 100, 28);
            };

            // Center: split list / editor
            var split = new SplitContainer { Dock = DockStyle.Fill };
            Load += (_, __) =>
            {
                split.Panel1MinSize    = 160;
                split.Panel2MinSize    = 320;
                split.SplitterDistance = 210;
            };

            _list = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, Font = new Font("Segoe UI", 9.5f) };
            _list.SelectedIndexChanged += ListSelectionChanged;

            var pnlBtns  = new Panel { Dock = DockStyle.Bottom, Height = 34 };
            var btnAdd    = Btn("+ Hinzufügen", new Point(4,   5), 108);
            var btnRemove = Btn("− Entfernen",  new Point(116, 5), 88);
            btnAdd.Click    += (_, __) => AddInstance();
            btnRemove.Click += (_, __) => RemoveInstance();
            pnlBtns.Controls.AddRange(new Control[] { btnAdd, btnRemove });

            split.Panel1.Controls.Add(_list);
            split.Panel1.Controls.Add(pnlBtns);

            _globalPanel = BuildGlobalPanel();
            _globalPanel.Dock = DockStyle.Fill;

            _instPanel = BuildInstancePanel();
            _instPanel.Dock    = DockStyle.Fill;
            _instPanel.Visible = false;

            split.Panel2.Controls.Add(_instPanel);
            split.Panel2.Controls.Add(_globalPanel);

            Controls.Add(split);
            Controls.Add(pnlBottom);
            Controls.Add(pnlTop);
        }

        // ── Field panels ──────────────────────────────────────────────────────

        private Panel BuildGlobalPanel()
        {
            var scroll = new Panel { AutoScroll = true };
            var tbl = MakeTbl();
            Section(tbl, "Global");
            _gDllPath    = Field(tbl, "DLL-Pfad:");
            _gLogFile    = Field(tbl, "Log-Datei:");
            _gLogSizeMb  = Num(tbl, "Max. Log-Größe (MB):", 0, 9999);
            _gLogFiles   = Num(tbl, "Max. Log-Dateien:", 0, 999);
            _gConsoleLog = Chk(tbl, "Konsolenausgabe:");
            _gIdleTime   = Num(tbl, "COM Idle-Timeout (s):", 0, 86400);
            Section(tbl, "HTTP Basic Auth");
            _gBasicAuth = Chk(tbl, "Aktiviert:");
            _gAuthUser  = Field(tbl, "Benutzer:");
            _gAuthPass  = Field(tbl, "Passwort / Hash:");
            scroll.Controls.Add(tbl);
            return scroll;
        }

        private Panel BuildInstancePanel()
        {
            var scroll = new Panel { AutoScroll = true };
            var tbl = MakeTbl();
            Section(tbl, "Instanz");
            _iName = Field(tbl, "Name:");
            Section(tbl, "HTTP");
            _iHttpOn   = Chk(tbl, "Aktiviert:");
            _iHttpPort = Num(tbl, "Port:", 1, 65535);
            Section(tbl, "Telnet");
            _iTelnetOn   = Chk(tbl, "Aktiviert:");
            _iTelnetPort = Num(tbl, "Port:", 1, 65535);
            Section(tbl, "COM-Port");
            _iComPort = Num(tbl, "COM-Port (0 = auto):", 0, 256);
            _iBaud    = new ComboBox { Dock = DockStyle.Fill, Margin = new Padding(0, 3, 4, 3), DropDownStyle = ComboBoxStyle.DropDown };
            foreach (var b in new[] { "1200","2400","4800","9600","19200","38400","57600","115200" })
                _iBaud.Items.Add(b);
            Row(tbl, "Baudrate:", _iBaud);
            scroll.Controls.Add(tbl);
            return scroll;
        }

        // ── TableLayout helpers ───────────────────────────────────────────────

        private static TableLayoutPanel MakeTbl()
        {
            var t = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, Padding = new Padding(12, 8, 12, 8) };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return t;
        }
        private static void Section(TableLayoutPanel t, string title)
        {
            var l = new Label { Text = title, Font = new Font("Segoe UI", 9f, FontStyle.Bold), AutoSize = true, Padding = new Padding(0, 10, 0, 2) };
            t.Controls.Add(l); t.SetColumnSpan(l, 2);
        }
        private static TextBox Field(TableLayoutPanel t, string lbl) { var c = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0,3,4,3) }; Row(t,lbl,c); return c; }
        private static NumericUpDown Num(TableLayoutPanel t, string lbl, int min, int max) { var c = new NumericUpDown { Minimum=min, Maximum=max, Dock=DockStyle.Fill, Margin=new Padding(0,3,4,3) }; Row(t,lbl,c); return c; }
        private static CheckBox Chk(TableLayoutPanel t, string lbl) { var c = new CheckBox { AutoSize=true, Margin=new Padding(0,3,4,3) }; Row(t,lbl,c); return c; }
        private static void Row(TableLayoutPanel t, string lbl, Control c)
        {
            t.Controls.Add(new Label { Text=lbl, AutoSize=true, Anchor=AnchorStyles.Left|AnchorStyles.Top, Padding=new Padding(0,5,6,0) });
            t.Controls.Add(c);
        }
        private static Button Btn(string text, Point loc, int w) => new() { Text=text, Location=loc, Width=w, Height=24 };

        // ── List selection ────────────────────────────────────────────────────

        private void ListSelectionChanged(object? sender, EventArgs e)
        {
            if (_loading || _list.SelectedIndex < 0) return;
            CommitCurrent();
            _activeIdx = _list.SelectedIndex;
            LoadSelected();
        }

        // ── Model ↔ Form ──────────────────────────────────────────────────────

        private void CommitCurrent()
        {
            if (_activeIdx == 0)
            {
                _global.DllPath        = _gDllPath.Text.Trim();
                _global.LogFile        = _gLogFile.Text.Trim();
                _global.LogMaxSizeMb   = (int)_gLogSizeMb.Value;
                _global.LogMaxFiles    = (int)_gLogFiles.Value;
                _global.ConsoleLogging = _gConsoleLog.Checked;
                _global.ComIdleTimeout = (int)_gIdleTime.Value;
                _global.BasicAuth      = _gBasicAuth.Checked;
                _global.AuthUser       = _gAuthUser.Text.Trim();
                _global.AuthPass       = _gAuthPass.Text.Trim();
            }
            else
            {
                int i = _activeIdx - 1;
                if (i < 0 || i >= _instances.Count) return;
                var inst = _instances[i];
                inst.Name          = _iName.Text.Trim();
                inst.HttpEnabled   = _iHttpOn.Checked;
                inst.HttpPort      = (int)_iHttpPort.Value;
                inst.TelnetEnabled = _iTelnetOn.Checked;
                inst.TelnetPort    = (int)_iTelnetPort.Value;
                inst.ComPort       = (int)_iComPort.Value;
                if (int.TryParse(_iBaud.Text, out int baud)) inst.BaudRate = baud;
                _list.Items[_activeIdx] = InstLabel(inst);
            }
        }

        private void LoadSelected()
        {
            _loading = true;
            try
            {
                if (_activeIdx == 0)
                {
                    _globalPanel.Visible = true; _instPanel.Visible = false;
                    _globalPanel.BringToFront();
                    _gDllPath.Text       = _global.DllPath;
                    _gLogFile.Text       = _global.LogFile;
                    _gLogSizeMb.Value    = Clamp(_global.LogMaxSizeMb, 0, 9999);
                    _gLogFiles.Value     = Clamp(_global.LogMaxFiles,  0, 999);
                    _gConsoleLog.Checked = _global.ConsoleLogging;
                    _gIdleTime.Value     = Clamp(_global.ComIdleTimeout, 0, 86400);
                    _gBasicAuth.Checked  = _global.BasicAuth;
                    _gAuthUser.Text      = _global.AuthUser;
                    _gAuthPass.Text      = _global.AuthPass;
                }
                else
                {
                    int i = _activeIdx - 1;
                    if (i < 0 || i >= _instances.Count) return;
                    var inst = _instances[i];
                    _instPanel.Visible = true; _globalPanel.Visible = false;
                    _instPanel.BringToFront();
                    _iName.Text        = inst.Name;
                    _iHttpOn.Checked   = inst.HttpEnabled;
                    _iHttpPort.Value   = Clamp(inst.HttpPort,   1, 65535);
                    _iTelnetOn.Checked = inst.TelnetEnabled;
                    _iTelnetPort.Value = Clamp(inst.TelnetPort, 1, 65535);
                    _iComPort.Value    = Clamp(inst.ComPort,    0, 256);
                    _iBaud.Text        = inst.BaudRate.ToString();
                }
            }
            finally { _loading = false; }
        }

        // ── Add / Remove ──────────────────────────────────────────────────────

        private void AddInstance()
        {
            CommitCurrent();
            var inst = new InstanceCfg
            {
                Name       = $"instance{_instances.Count + 1}",
                HttpPort   = 8083 + _instances.Count,
                TelnetPort = 8023 + _instances.Count,
                ComPort    = _instances.Count + 1,
            };
            _instances.Add(inst);
            _loading = true; _list.Items.Add(InstLabel(inst)); _loading = false;
            _list.SelectedIndex = _list.Items.Count - 1;
        }

        private void RemoveInstance()
        {
            if (_activeIdx <= 0) { MessageBox.Show("Global-Einstellungen können nicht entfernt werden.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            int i = _activeIdx - 1;
            if (MessageBox.Show($"Instanz \"{_instances[i].Name}\" wirklich entfernen?", "Bestätigung", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            _instances.RemoveAt(i);
            _loading = true; _list.Items.RemoveAt(_activeIdx); _loading = false;
            _list.SelectedIndex = 0;
        }

        // ── INI read ──────────────────────────────────────────────────────────

        private void TryLoad(string path)
        {
            if (!File.Exists(path)) { MessageBox.Show($"Datei nicht gefunden:\n{path}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            LoadIni(path);
        }

        private void LoadIni(string path)
        {
            _iniPath = path; _pathBox.Text = path;
            _global  = new GlobalCfg(); _instances.Clear();

            string? section = null; InstanceCfg? cur = null;

            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                string line = raw.Trim();
                if (line.StartsWith(";") || line.StartsWith("#") || line.Length == 0) continue;
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    if (cur != null) { _instances.Add(cur); cur = null; }
                    section = line[1..^1].Trim().ToLower();
                    if (section.StartsWith("instance")) cur = new InstanceCfg();
                    continue;
                }
                int eq = line.IndexOf('='); if (eq < 0) continue;
                string key = line[..eq].Trim().ToLower();
                string val = line[(eq+1)..].Trim();
                int ci = val.IndexOf(';'); if (ci > 0) val = val[..ci].Trim();

                if (cur != null) switch (key)
                {
                    case "name":           cur.Name          = val; break;
                    case "http_enabled":   cur.HttpEnabled   = Bool(val); break;
                    case "http_port":      if (int.TryParse(val, out int hp)) cur.HttpPort = hp; break;
                    case "telnet_enabled": cur.TelnetEnabled = Bool(val); break;
                    case "telnet_port":    if (int.TryParse(val, out int tp)) cur.TelnetPort = tp; break;
                    case "com_port":       if (int.TryParse(val, out int cp)) cur.ComPort = cp; break;
                    case "baud_rate":      if (int.TryParse(val, out int br)) cur.BaudRate = br; break;
                    case "basic_auth":     _global.BasicAuth = Bool(val); break;
                    case "auth_user":      _global.AuthUser  = val; break;
                    case "auth_pass":      _global.AuthPass  = val; break;
                }
                else switch (key)
                {
                    case "dll_path":         _global.DllPath        = val; break;
                    case "log_file":         _global.LogFile        = val; break;
                    case "console_logging":  _global.ConsoleLogging = Bool(val); break;
                    case "log_max_size_mb":  if (int.TryParse(val, out int ls)) _global.LogMaxSizeMb = ls; break;
                    case "log_max_files":    if (int.TryParse(val, out int lf)) _global.LogMaxFiles  = lf; break;
                    case "com_idle_timeout": if (int.TryParse(val, out int it)) _global.ComIdleTimeout = it; break;
                    case "basic_auth":       _global.BasicAuth = Bool(val); break;
                    case "auth_user":        _global.AuthUser  = val; break;
                    case "auth_pass":        _global.AuthPass  = val; break;
                }
            }
            if (cur != null) _instances.Add(cur);
            if (_instances.Count == 0) _instances.Add(new InstanceCfg());

            TakeSnapshot();
            RefreshList(0);
        }

        // ── INI write ─────────────────────────────────────────────────────────

        private void WriteIniFile()
        {
            var sb = new StringBuilder();
            sb.AppendLine("; EasyComServer configuration file");
            sb.AppendLine($"; Gespeichert von EasyComConfigurator — {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            sb.AppendLine();
            sb.AppendLine("[global]");
            sb.AppendLine($"dll_path         = {_global.DllPath}");
            sb.AppendLine($"log_file         = {_global.LogFile}");
            sb.AppendLine($"log_max_size_mb  = {_global.LogMaxSizeMb}");
            sb.AppendLine($"log_max_files    = {_global.LogMaxFiles}");
            sb.AppendLine($"console_logging  = {(_global.ConsoleLogging ? "true" : "false")}");
            sb.AppendLine($"com_idle_timeout = {_global.ComIdleTimeout}");
            sb.AppendLine();
            sb.AppendLine("; HTTP Basic Auth (global — applies to all instances)");
            sb.AppendLine($"basic_auth       = {(_global.BasicAuth ? "true" : "false")}");
            sb.AppendLine($"auth_user        = {_global.AuthUser}");
            sb.AppendLine($"auth_pass        = {_global.AuthPass}");
            foreach (var inst in _instances)
            {
                sb.AppendLine();
                sb.AppendLine($"[instance: {inst.Name}]");
                sb.AppendLine($"name             = {inst.Name}");
                sb.AppendLine($"http_enabled     = {(inst.HttpEnabled   ? "true" : "false")}");
                sb.AppendLine($"http_port        = {inst.HttpPort}");
                sb.AppendLine($"telnet_enabled   = {(inst.TelnetEnabled ? "true" : "false")}");
                sb.AppendLine($"telnet_port      = {inst.TelnetPort}");
                sb.AppendLine($"com_port         = {inst.ComPort}");
                sb.AppendLine($"baud_rate        = {inst.BaudRate}");
            }
            File.WriteAllText(_iniPath, sb.ToString(), Encoding.UTF8);
        }

        // ── Save & live-inject ────────────────────────────────────────────────

        private async Task SaveAndApply()
        {
            if (string.IsNullOrEmpty(_iniPath)) { BrowseSave(); if (string.IsNullOrEmpty(_iniPath)) return; }
            CommitCurrent();

            // Collect changes before writing
            var (liveGlobal, liveInst, needRestart) = CollectChanges();

            try { WriteIniFile(); }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Speichern:\n{ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Credentials to use for authenticated requests (old values still active in service)
            string? authUser = _snap.BasicAuth ? _snap.AuthUser : null;
            string? authPass = _snap.BasicAuth ? _snap.AuthPass : null;

            // Send global live changes to every instance (they share one ServerConfig)
            var liveOk   = new List<string>();
            var liveFail = new List<string>();

            if (liveGlobal.Count > 0)
            {
                var target = _instances.Count > 0 ? _instances[0] : null;
                if (target != null)
                {
                    foreach (var ch in liveGlobal)
                    {
                        bool ok = await TrySendSetConfig(target, ch.Key, ch.Value, authUser, authPass);
                        (ok ? liveOk : liveFail).Add(ch.Label);
                    }
                }
            }

            // Send per-instance live changes
            foreach (var (instIdx, changes) in liveInst)
            {
                if (instIdx >= _instances.Count) continue;
                var inst = _instances[instIdx];
                foreach (var ch in changes)
                {
                    bool ok = await TrySendSetConfig(inst, ch.Key, ch.Value, authUser, authPass);
                    (ok ? liveOk : liveFail).Add($"{inst.Name}/{ch.Label}");
                }
            }

            TakeSnapshot();
            ShowSaveResult(liveOk, liveFail, needRestart);
        }

        // ── Change detection ──────────────────────────────────────────────────

        private (List<Change> liveGlobal,
                 List<(int instIdx, List<Change> changes)> liveInst,
                 List<string> needRestart)
            CollectChanges()
        {
            var liveGlobal  = new List<Change>();
            var liveInst    = new List<(int, List<Change>)>();
            var needRestart = new List<string>();

            // Global
            if (_global.ConsoleLogging != _snap.ConsoleLogging) liveGlobal.Add(new("console_logging", "console_logging", _global.ConsoleLogging ? "true" : "false"));
            if (_global.ComIdleTimeout != _snap.ComIdleTimeout) liveGlobal.Add(new("com_idle_timeout", "com_idle_timeout", _global.ComIdleTimeout.ToString()));
            if (_global.BasicAuth      != _snap.BasicAuth)      liveGlobal.Add(new("basic_auth", "basic_auth", _global.BasicAuth ? "true" : "false"));
            if (_global.AuthUser       != _snap.AuthUser)       liveGlobal.Add(new("auth_user",  "auth_user",  _global.AuthUser));
            if (_global.AuthPass       != _snap.AuthPass)       liveGlobal.Add(new("auth_pass",  "auth_pass",  _global.AuthPass));

            if (_global.DllPath      != _snap.DllPath)      needRestart.Add("dll_path");
            if (_global.LogFile      != _snap.LogFile)       needRestart.Add("log_file");
            if (_global.LogMaxSizeMb != _snap.LogMaxSizeMb) needRestart.Add("log_max_size_mb");
            if (_global.LogMaxFiles  != _snap.LogMaxFiles)   needRestart.Add("log_max_files");

            // Instances: compare by index (add/remove always triggers restart)
            if (_instances.Count != _snapInstances.Count)
            {
                needRestart.Add("Anzahl Instanzen geändert");
            }
            else
            {
                for (int i = 0; i < _instances.Count; i++)
                {
                    var cur  = _instances[i];
                    var snap = _snapInstances[i];
                    var instLive = new List<Change>();

                    if (cur.ComPort  != snap.ComPort)  instLive.Add(new("com_port",  "com_port",  cur.ComPort.ToString()));
                    if (cur.BaudRate != snap.BaudRate)  instLive.Add(new("baud_rate", "baud_rate", cur.BaudRate.ToString()));

                    if (instLive.Count > 0) liveInst.Add((i, instLive));

                    if (cur.Name          != snap.Name)          needRestart.Add($"{cur.Name}/name");
                    if (cur.HttpEnabled   != snap.HttpEnabled)   needRestart.Add($"{cur.Name}/http_enabled");
                    if (cur.HttpPort      != snap.HttpPort)      needRestart.Add($"{cur.Name}/http_port");
                    if (cur.TelnetEnabled != snap.TelnetEnabled) needRestart.Add($"{cur.Name}/telnet_enabled");
                    if (cur.TelnetPort    != snap.TelnetPort)    needRestart.Add($"{cur.Name}/telnet_port");
                }
            }

            return (liveGlobal, liveInst, needRestart);
        }

        // ── HTTP live-inject ──────────────────────────────────────────────────

        private async Task<bool> TrySendSetConfig(InstanceCfg inst, string key, string value,
            string? authUser, string? authPass)
        {
            try
            {
                string cmd     = $"SET CONFIGURATION \"{key}={value}\"";
                string url     = $"http://localhost:{inst.HttpPort}/easy.cmd?{Uri.EscapeDataString(cmd)}";
                var    request = new HttpRequestMessage(HttpMethod.Get, url);

                if (authUser != null)
                {
                    string creds   = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{authUser}:{authPass}"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Basic", creds);
                }

                var resp = await _http.SendAsync(request);
                string body = await resp.Content.ReadAsStringAsync();
                return body.TrimStart().StartsWith("OK", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        // ── Result dialog ─────────────────────────────────────────────────────

        private void ShowSaveResult(List<string> liveOk, List<string> liveFail, List<string> needRestart)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Konfiguration gespeichert.");

            if (liveOk.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("✅ Live aktualisiert (kein Neustart nötig):");
                foreach (var s in liveOk) sb.AppendLine($"   • {s}");
            }

            if (liveFail.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("⚠ Live-Update fehlgeschlagen (Service nicht erreichbar?):");
                foreach (var s in liveFail) sb.AppendLine($"   • {s}");
            }

            if (needRestart.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("🔄 Neustart erforderlich:");
                foreach (var s in needRestart) sb.AppendLine($"   • {s}");
                sb.AppendLine();
                sb.Append("Dienst jetzt neu starten?");

                if (MessageBox.Show(sb.ToString(), "EasyComConfigurator",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ServiceAction(s =>
                    {
                        if (s.Status == ServiceControllerStatus.Running)
                        { s.Stop(); s.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(15)); }
                        s.Start();
                    });
            }
            else
            {
                MessageBox.Show(sb.ToString(), "EasyComConfigurator",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // ── Service control ───────────────────────────────────────────────────

        private void ServiceAction(Action<ServiceController> action)
        {
            try
            {
                using var svc = new ServiceController(SvcName);
                action(svc);
                System.Threading.Thread.Sleep(800);
                RefreshServiceStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Dienst-Zugriff:\n{ex.Message}\n\nBitte als Administrator ausführen.",
                    "Dienst-Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RefreshServiceStatus();
            }
        }

        private void RefreshServiceStatus()
        {
            try
            {
                using var svc = new ServiceController(SvcName);
                svc.Refresh();
                (string text, Color color) = svc.Status switch
                {
                    ServiceControllerStatus.Running      => ("● Läuft",     Color.Green),
                    ServiceControllerStatus.Stopped      => ("○ Gestoppt",  Color.OrangeRed),
                    ServiceControllerStatus.StartPending => ("↑ Startet …", Color.DarkOrange),
                    ServiceControllerStatus.StopPending  => ("↓ Stoppt …",  Color.DarkOrange),
                    ServiceControllerStatus.Paused       => ("⏸ Pausiert",  Color.Gray),
                    _ => (svc.Status.ToString(), Color.Gray),
                };
                _lblSvcStatus.Text      = text;
                _lblSvcStatus.ForeColor = color;
            }
            catch
            {
                _lblSvcStatus.Text      = "○ Nicht installiert";
                _lblSvcStatus.ForeColor = Color.Gray;
            }
        }

        // ── Snapshot ──────────────────────────────────────────────────────────

        private void TakeSnapshot()
        {
            _snap = _global.Clone();
            _snapInstances.Clear();
            foreach (var inst in _instances) _snapInstances.Add(inst.Clone());
        }

        // ── List helpers ──────────────────────────────────────────────────────

        private void RefreshList(int selectIdx)
        {
            _loading = true;
            _list.Items.Clear();
            _list.Items.Add("⚙  Global");
            foreach (var inst in _instances) _list.Items.Add(InstLabel(inst));
            _loading = false;
            _list.SelectedIndex = Math.Max(0, Math.Min(selectIdx, _list.Items.Count - 1));
        }

        private static string InstLabel(InstanceCfg inst) => $"▶  {inst.Name}  (:{inst.HttpPort})";

        // ── Browse ────────────────────────────────────────────────────────────

        private void BrowseOpen()
        {
            using var dlg = new OpenFileDialog { Title = "easycom.ini öffnen", Filter = "INI-Dateien (*.ini)|*.ini|Alle Dateien (*.*)|*.*", FileName = "easycom.ini" };
            if (dlg.ShowDialog() == DialogResult.OK) TryLoad(dlg.FileName);
        }

        private void BrowseSave()
        {
            using var dlg = new SaveFileDialog { Title = "easycom.ini speichern unter", Filter = "INI-Dateien (*.ini)|*.ini|Alle Dateien (*.*)|*.*", FileName = "easycom.ini" };
            if (dlg.ShowDialog() == DialogResult.OK) _iniPath = dlg.FileName;
        }

        // ── Utilities ─────────────────────────────────────────────────────────

        private static bool Bool(string s) =>
            s.Equals("true",    StringComparison.OrdinalIgnoreCase)
            || s.Equals("yes",  StringComparison.OrdinalIgnoreCase)
            || s.Equals("enabled", StringComparison.OrdinalIgnoreCase)
            || s == "1";

        private static decimal Clamp(int v, int min, int max) => Math.Max(min, Math.Min(max, v));
    }
}
