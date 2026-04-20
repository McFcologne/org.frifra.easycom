(function () {
  var cfg = null;
  var selIdx = 0;
  var toastEl = null;
  var toastTimer = null;

  function ensureToast() {
    if (toastEl) return;
    toastEl = document.createElement("div");
    toastEl.className = "cfg-toast";
    document.body.appendChild(toastEl);
  }

  function toast(msg, type) {
    ensureToast();
    toastEl.textContent = msg;
    toastEl.className = "cfg-toast show " + type;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove("show"); }, 7000);
  }

  // ── Open / close ───────────────────────────────────────────────────────────
  window.openSettings = async function () {
    document.getElementById("cfg-overlay").classList.add("show");
    await loadConfig();
  };

  window.closeSettings = function () {
    document.getElementById("cfg-overlay").classList.remove("show");
  };

  document.addEventListener("DOMContentLoaded", function () {
    var ov = document.getElementById("cfg-overlay");
    if (ov) {
      ov.addEventListener("mousedown", function (e) {
        if (e.target === ov) closeSettings();
      });
    }
  });

  // ── Load config ────────────────────────────────────────────────────────────
  async function loadConfig() {
    try {
      var r = await fetch("/api/v1/config");
      var j = await r.json();
      if (!j.ok) { toast("Fehler: " + j.error, "err"); return; }
      cfg = j.config;
      document.getElementById("cfg-ini-path").value = cfg.ini_path || "";
      selIdx = 0;
      renderList();
      renderEditor();
      updateSvcStatus();
    } catch (e) { toast("Konfiguration nicht ladbar: " + e, "err"); }
  }

  // ── List ───────────────────────────────────────────────────────────────────
  function renderList() {
    var lb = document.getElementById("cfg-list");
    lb.innerHTML = "";
    addLi(lb, 0, "\u2699\u2002Global");
    (cfg.instances || []).forEach(function (inst, i) {
      addLi(lb, i + 1, "\u25B6\u2002" + inst.name + "  (:" + inst.http_port + ")");
    });
  }

  function addLi(lb, idx, text) {
    var li = document.createElement("li");
    li.textContent = text;
    if (selIdx === idx) li.className = "sel";
    li.onclick = function () { selectItem(idx); };
    lb.appendChild(li);
  }

  function selectItem(idx) { commitCurrent(); selIdx = idx; renderList(); renderEditor(); }

  // ── Editor ─────────────────────────────────────────────────────────────────
  function renderEditor() {
    var p = document.getElementById("cfg-editor");
    if (selIdx === 0) {
      p.innerHTML = globalFormHtml();
    } else {
      var inst = cfg.instances[selIdx - 1];
      if (inst) p.innerHTML = instanceFormHtml(inst);
    }
  }

  function globalFormHtml() {
    return sh("Global") +
      ft("dll_path",        "DLL-Pfad:",               cfg.dll_path) +
      ft("log_file",        "Log-Datei:",               cfg.log_file) +
      fn("log_max_size_mb", "Max. Log-Gr\u00f6\u00dfe (MB):", cfg.log_max_size_mb, 0, 9999) +
      fn("log_max_files",   "Max. Log-Dateien:",         cfg.log_max_files,   0, 999) +
      fc("console_logging", "Konsolenausgabe:",          cfg.console_logging) +
      fn("com_idle_timeout","COM Idle-Timeout (s):",     cfg.com_idle_timeout, 0, 86400) +
      sh("HTTP Basic Auth") +
      fc("basic_auth",      "Aktiviert:",                cfg.basic_auth) +
      ft("auth_user",       "Benutzer:",                 cfg.auth_user) +
      fp("auth_pass",       "Passwort / Hash:",           cfg.auth_pass);
  }

  function instanceFormHtml(inst) {
    return sh("Instanz") + ft("i_name", "Name:", inst.name) +
      sh("HTTP") +
      fc("i_http_enabled",   "Aktiviert:",          inst.http_enabled) +
      fn("i_http_port",      "Port:",                inst.http_port, 1, 65535) +
      sh("Telnet") +
      fc("i_telnet_enabled", "Aktiviert:",          inst.telnet_enabled) +
      fn("i_telnet_port",    "Port:",                inst.telnet_port, 1, 65535) +
      sh("COM-Port") +
      fn("i_com_port",       "COM-Port (0\u202f=\u202fauto):", inst.com_port, 0, 256) +
      fs("i_baud_rate",      "Baudrate:",             inst.baud_rate,
         [1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200]);
  }

  function sh(t)        { return "<div class=\"cfg-sec\">" + esc(t) + "</div>"; }
  function ft(id, l, v) { return "<div class=\"cfg-row\"><label>" + l + "</label><input type=\"text\" id=\"" + id + "\" value=\"" + esc(v || "") + "\"></div>"; }
  function fp(id, l, v) {
    return "<div class=\"cfg-row\"><label>" + l + "</label>" +
      "<div class=\"cfg-pass-wrap\">" +
      "<input type=\"password\" id=\"" + id + "\" value=\"" + esc(v || "") + "\">" +
      "<button type=\"button\" class=\"cfg-pass-btn\" onclick=\"cfgTogglePass('" + id + "',this)\" tabindex=\"-1\">👁</button>" +
      "</div></div>";
  }
  function fn(id, l, v, mn, mx) {
    return "<div class=\"cfg-row\"><label>" + l + "</label><input type=\"number\" id=\"" + id + "\" value=\"" + (+v || 0) + "\" min=\"" + mn + "\" max=\"" + mx + "\"></div>";
  }
  function fc(id, l, v) { return "<div class=\"cfg-row\"><label>" + l + "</label><input type=\"checkbox\" id=\"" + id + "\"" + (v ? " checked" : "") + "></div>"; }
  function fs(id, l, v, opts) {
    var os = opts.map(function (o) {
      return "<option value=\"" + o + "\"" + (o == v ? " selected" : "") + ">" + o + "</option>";
    }).join("");
    return "<div class=\"cfg-row\"><label>" + l + "</label><select id=\"" + id + "\">" + os + "</select></div>";
  }

  // ── Commit form → cfg ──────────────────────────────────────────────────────
  function commitCurrent() {
    if (!cfg) return;
    var dg = function (id) { var e = document.getElementById(id); return e ? e.value : null; };
    var ng = function (id) { var e = document.getElementById(id); return e ? +e.value : null; };
    var bg = function (id) { var e = document.getElementById(id); return e ? e.checked : null; };
    if (selIdx === 0) {
      if (dg("dll_path")        !== null) cfg.dll_path        = dg("dll_path");
      if (dg("log_file")         !== null) cfg.log_file         = dg("log_file");
      if (ng("log_max_size_mb") !== null) cfg.log_max_size_mb = ng("log_max_size_mb");
      if (ng("log_max_files")   !== null) cfg.log_max_files   = ng("log_max_files");
      if (bg("console_logging") !== null) cfg.console_logging = bg("console_logging");
      if (ng("com_idle_timeout")!== null) cfg.com_idle_timeout= ng("com_idle_timeout");
      if (bg("basic_auth")      !== null) cfg.basic_auth      = bg("basic_auth");
      if (dg("auth_user")        !== null) cfg.auth_user        = dg("auth_user");
      if (dg("auth_pass")        !== null) cfg.auth_pass        = dg("auth_pass");
    } else {
      var inst = cfg.instances && cfg.instances[selIdx - 1];
      if (!inst) return;
      if (dg("i_name")           !== null) inst.name           = dg("i_name");
      if (bg("i_http_enabled")   !== null) inst.http_enabled   = bg("i_http_enabled");
      if (ng("i_http_port")       !== null) inst.http_port       = ng("i_http_port");
      if (bg("i_telnet_enabled") !== null) inst.telnet_enabled = bg("i_telnet_enabled");
      if (ng("i_telnet_port")    !== null) inst.telnet_port    = ng("i_telnet_port");
      if (ng("i_com_port")        !== null) inst.com_port        = ng("i_com_port");
      if (ng("i_baud_rate")       !== null) inst.baud_rate       = ng("i_baud_rate");
      renderList();
    }
  }

  // ── Add / Remove ───────────────────────────────────────────────────────────
  window.cfgAddInstance = function () {
    commitCurrent();
    var n = cfg.instances.length;
    cfg.instances.push({
      name: "instance" + (n + 1),
      http_enabled: true,   http_port: 8083 + n,
      telnet_enabled: true, telnet_port: 8023 + n,
      com_port: n + 1,      baud_rate: 9600
    });
    renderList();
    selectItem(cfg.instances.length);
  };

  window.cfgRemoveInstance = function () {
    if (selIdx === 0) {
      toast("Global-Einstellungen k\u00f6nnen nicht entfernt werden.", "warn");
      return;
    }
    var inst = cfg.instances[selIdx - 1];
    if (!confirm("Instanz \"" + inst.name + "\" wirklich entfernen?")) return;
    cfg.instances.splice(selIdx - 1, 1);
    selectItem(0);
  };

  // ── Save ───────────────────────────────────────────────────────────────────
  window.cfgSave = async function () {
    commitCurrent();
    try {
      var r = await fetch("/api/v1/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dll_path: cfg.dll_path,       log_file: cfg.log_file,
          console_logging: cfg.console_logging,
          log_max_size_mb: cfg.log_max_size_mb,
          log_max_files:   cfg.log_max_files,
          com_idle_timeout: cfg.com_idle_timeout,
          basic_auth: cfg.basic_auth,   auth_user: cfg.auth_user,
          auth_pass: cfg.auth_pass,     instances: cfg.instances
        })
      });
      var j = await r.json();
      if (!j.ok) { toast("Fehler: " + (j.error || "unbekannt"), "err"); return; }
      var msg     = j.result || "Gespeichert";
      var live    = j.live_applied     || [];
      var restart = j.restart_required || [];
      if (live.length)    msg += "\n\u2705 Live: "                   + live.join(", ");
      if (restart.length) msg += "\n\uD83D\uDD04 Neustart erforderlich: " + restart.join(", ");
      toast(msg, restart.length ? "warn" : "ok");
    } catch (e) { toast("Speichern fehlgeschlagen: " + e, "err"); }
  };

  // ── Service status ─────────────────────────────────────────────────────────
  async function updateSvcStatus() {
    try {
      var r = await fetch("/api/v1/system");
      var j = await r.json();
      var first = ((j.result || "").split("\n")[0] || "").slice(0, 80);
      document.getElementById("cfg-svc-status").textContent = first;
    } catch (e) {
      document.getElementById("cfg-svc-status").textContent = "nicht erreichbar";
    }
  }

  // ── Password field helpers ─────────────────────────────────────────────────
  window.cfgTogglePass = function (id, btn) {
    var inp = document.getElementById(id);
    if (!inp) return;
    inp.type = inp.type === "password" ? "text" : "password";
    btn.textContent = inp.type === "password" ? "👁" : "🙈";
  };

// ── Utility ────────────────────────────────────────────────────────────────
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
}());
