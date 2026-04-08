document.getElementById('server-url-display').textContent =
  location.protocol + '//' + location.host;

fetch('./readme.md')
  .then(r => {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.text();
  })
  .then(md => {
    document.getElementById('content').innerHTML = marked.parse(md);
  })
  .catch(err => {
    document.getElementById('content').innerHTML =
      '<p style="color:#f85149">Failed to load README.md: ' + err.message + '</p>';
  });
