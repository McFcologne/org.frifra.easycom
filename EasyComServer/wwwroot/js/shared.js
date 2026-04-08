(async function poll() {
  const dot = document.getElementById('status-dot');
  const lbl = document.getElementById('dot-label');
  try {
    const r = await fetch(
      location.protocol + '//' + location.host + '/easy.cmd?' + encodeURIComponent('show server'),
      { credentials: 'include', signal: AbortSignal.timeout(4000) }
    );
    dot.className = 'dot ' + (r.ok ? 'ok' : 'err');
    lbl.textContent = r.ok ? 'online' : 'offline';
  } catch {
    dot.className = 'dot err';
    lbl.textContent = 'offline';
  }
  setTimeout(poll, 15000);
})();
