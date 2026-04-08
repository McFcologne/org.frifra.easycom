let cfg = { url: 'http://localhost:8083' };
let cmdHistory = [];
let histIdx = -1;
let cmdCount = 0;

const COMMANDS = [
  'open_comport','close_comport','getcurrent_baudrate',
  'set_userwaitingtime','get_userwaitingtime',
  'open_ethernetport','close_ethernetport',
  'start_program','stop_program',
  'read_clock','write_clock',
  'read_object_value','write_object_value',
  'read_channel_yeartimeswitch','write_channel_yeartimeswitch',
  'read_channel_7daytimeswitch','write_channel_7daytimeswitch',
  'unlock_device','lock_device',
  'mc_open_comport','mc_close_comport','mc_getcurrent_baudrate',
  'mc_set_userwaitingtime','mc_get_userwaitingtime',
  'mc_open_ethernetport','mc_close_ethernetport','mc_closeall',
  'mc_start_program','mc_stop_program',
  'mc_read_clock','mc_write_clock',
  'mc_read_object_value','mc_write_object_value',
  'mc_read_channel_yeartimeswitch','mc_write_channel_yeartimeswitch',
  'mc_read_channel_7daytimeswitch','mc_write_channel_7daytimeswitch',
  'mc_unlock_device','mc_lock_device',
  'getlastsystemerror',
  'show server','show connections','show configuration','show tasks',
  'set configuration','help',
];

function appendLine(type, text) {
  const out = document.getElementById('output');
  if (!out) return;
  const span = document.createElement('span');
  span.className = 'line line-' + type;
  span.textContent = text;
  out.appendChild(span);
  out.scrollTop = out.scrollHeight;
}

function clearOutput() {
  document.getElementById('output').innerHTML = '';
  appendLine('sys', 'Console cleared - ' + new Date().toLocaleString('en'));
}

function setDot(state, label) {
  document.getElementById('status-dot').className = 'dot ' + state;
  document.getElementById('dot-label').textContent = label;
}

function applySettings() {
  cfg.url = document.getElementById('cfg-url').value.trim().replace(/\/$/, '');
  document.getElementById('sb-url').textContent = cfg.url;
  document.getElementById('server-url-display').textContent = cfg.url;
  appendLine('sys', 'Connecting to ' + cfg.url + ' ...');
  runCmd('show server');
}

function openCom() {
  const port = document.getElementById('cfg-comport').value;
  const baud = document.getElementById('cfg-baud').value;
  runCmd('open_comport ' + port + ' ' + baud);
}

function setIdle() {
  runCmd('set configuration "com_idle_timeout=' + document.getElementById('cfg-idle').value + '"');
}

function switchTab(name, btn) {
  document.querySelectorAll('.rp-section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.rp-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  btn.classList.add('active');
}

function promptAndRun(base, fields) {
  const vals = [];
  for (const f of fields) {
    const v = window.prompt(base.toUpperCase() + '\n' + f + ':');
    if (v === null) return;
    vals.push(v.trim());
  }
  runCmd(base + ' ' + vals.join(' '));
}

async function sendCmd() {
  const input = document.getElementById('cmd-input');
  const cmd = input.value.trim();
  if (!cmd) return;
  input.value = '';
  histIdx = -1;
  runCmd(cmd);
}

async function runCmd(cmd) {
  if (!cmd) return;
  if (cmdHistory.length === 0 || cmdHistory[0] !== cmd) {
    cmdHistory.unshift(cmd);
    if (cmdHistory.length > 80) cmdHistory.pop();
    renderHistory();
  }
  cmdCount++;
  document.getElementById('sb-count').textContent = cmdCount;

  appendLine('sep', '\u2500'.repeat(52));
  appendLine('cmd', '> ' + cmd);
  setDot('busy', 'sending ...');

  try {
    const res = await fetch(cfg.url + '/easy.cmd?' + encodeURIComponent(cmd), {
      credentials: 'include',
      signal: AbortSignal.timeout(8000)
    });

    if (res.status === 401) {
      setDot('err', '401 Unauthorized');
      appendLine('err', 'ERROR 401 - Please reload the page and enter your credentials');
      return;
    }

    const text = (await res.text()).trim();
    text.split(/\r?\n/).forEach(l => { if (l) appendLine(classifyLine(l), l); });
    setDot('ok', 'online');
    document.getElementById('status-msg').textContent =
      new Date().toLocaleTimeString('en') + ' - ' + cmd.split(' ')[0].toUpperCase();

  } catch (e) {
    setDot('err', 'error');
    appendLine('err', 'ERROR ' + e.message);
  }
}

function classifyLine(l) {
  if (l.startsWith('OK'))    return 'ok';
  if (l.startsWith('ERROR')) return 'err';
  if (l.match(/^\d/))        return 'info';
  return 'info';
}

function renderHistory() {
  const ul = document.getElementById('history-list');
  ul.innerHTML = '';
  cmdHistory.forEach(cmd => {
    const li = document.createElement('li');
    li.textContent = cmd;
    li.onclick = () => {
      document.getElementById('cmd-input').value = cmd;
      document.getElementById('cmd-input').focus();
    };
    ul.appendChild(li);
  });
}

document.getElementById('cmd-input').addEventListener('keydown', e => {
  if (e.key === 'Enter') { sendCmd(); return; }

  if (e.key === 'ArrowUp') {
    e.preventDefault();
    if (histIdx < cmdHistory.length - 1) histIdx++;
    document.getElementById('cmd-input').value = cmdHistory[histIdx] || '';
    return;
  }
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    if (histIdx > 0) histIdx--;
    else { histIdx = -1; document.getElementById('cmd-input').value = ''; return; }
    document.getElementById('cmd-input').value = cmdHistory[histIdx] || '';
    return;
  }

  if (e.key === 'Tab') {
    e.preventDefault();
    const val = document.getElementById('cmd-input').value.toLowerCase();
    if (!val) return;
    const matches = COMMANDS.filter(c => c.startsWith(val));
    if (matches.length === 1) {
      document.getElementById('cmd-input').value = matches[0] + ' ';
    } else if (matches.length > 1) {
      appendLine('sys', matches.join('  '));
    }
  }
});

async function loadServerConfig() {
  try {
    const r1 = await fetch(cfg.url + '/easy.cmd?' + encodeURIComponent('show configuration com_port'), { credentials: 'include' });
    if (r1.ok) {
      const t1 = (await r1.text()).trim().split('\n');
      const line = t1.find(l => l.startsWith('com_port='));
      if (line) document.getElementById('cfg-comport').value = line.split('=')[1].trim();
    }
    const r2 = await fetch(cfg.url + '/easy.cmd?' + encodeURIComponent('show configuration baud_rate'), { credentials: 'include' });
    if (r2.ok) {
      const t2 = (await r2.text()).trim().split('\n');
      const line = t2.find(l => l.startsWith('baud_rate='));
      if (line) {
        const val = line.split('=')[1].trim();
        const sel = document.getElementById('cfg-baud');
        for (let i = 0; i < sel.options.length; i++) {
          if (sel.options[i].value === val) { sel.selectedIndex = i; break; }
        }
      }
    }
    const r3 = await fetch(cfg.url + '/easy.cmd?' + encodeURIComponent('show configuration com_idle_timeout'), { credentials: 'include' });
    if (r3.ok) {
      const t3 = (await r3.text()).trim().split('\n');
      const line = t3.find(l => l.startsWith('com_idle_timeout='));
      if (line) document.getElementById('cfg-idle').value = line.split('=')[1].trim();
    }
  } catch(e) { /* ignore if server not yet reachable */ }
}

(function init() {
  cfg.url = location.protocol + '//' + location.host;
  document.getElementById('cfg-url').value = cfg.url;
  document.getElementById('sb-url').textContent = cfg.url;
  document.getElementById('server-url-display').textContent = cfg.url;

  appendLine('sys', 'EASY COM Server Console');
  appendLine('sys', 'Server: ' + cfg.url);
  appendLine('sys', new Date().toLocaleString('en'));
  appendLine('sep', '\u2500'.repeat(52));

  loadServerConfig();
  runCmd('show server');
  document.getElementById('cmd-input').focus();

  // Background health poll — keeps dot current between commands
  setInterval(async () => {
    if (document.getElementById('status-dot').classList.contains('busy')) return;
    try {
      const r = await fetch(cfg.url + '/easy.cmd?' + encodeURIComponent('show server'), {
        credentials: 'include', signal: AbortSignal.timeout(4000)
      });
      setDot(r.ok ? 'ok' : 'err', r.ok ? 'online' : 'offline');
    } catch {
      setDot('err', 'offline');
    }
  }, 15000);
})();
