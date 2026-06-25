// Embedded dashboard HTML served by HTTPServer at GET /.
let dashboardHTML = #"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Outlook Capture Dashboard</title>
<style>
  :root {
    --bg: #0f1117; --surface: #1a1d27; --border: #2e3147;
    --accent: #5b8af7; --green: #34d399; --red: #f87171;
    --yellow: #fbbf24; --text: #e4e6f0; --muted: #6b7280;
    --font: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--text); font-family: var(--font); font-size: 13px; }
  header {
    display: flex; align-items: center; gap: 12px;
    padding: 14px 20px; background: var(--surface);
    border-bottom: 1px solid var(--border);
    position: sticky; top: 0; z-index: 10;
  }
  header h1 { font-size: 16px; font-weight: 600; }
  #status-dot { width: 9px; height: 9px; border-radius: 50%; background: var(--muted); flex-shrink: 0; }
  #status-dot.connected { background: var(--green); }
  #status-dot.error { background: var(--red); }
  #event-count { margin-left: auto; color: var(--muted); }
  .toolbar {
    display: flex; gap: 8px; padding: 10px 20px;
    background: var(--surface); border-bottom: 1px solid var(--border);
  }
  .toolbar input {
    flex: 1; background: var(--bg); border: 1px solid var(--border);
    color: var(--text); border-radius: 6px; padding: 5px 10px; font-size: 12px;
  }
  .toolbar button {
    background: var(--border); color: var(--text); border: none;
    border-radius: 6px; padding: 5px 12px; cursor: pointer; font-size: 12px;
  }
  .toolbar button:hover { background: var(--accent); }
  #table-wrap { overflow-x: auto; padding: 16px 20px; }
  table { width: 100%; border-collapse: collapse; }
  thead th {
    text-align: left; padding: 8px 10px; font-size: 11px; font-weight: 600;
    text-transform: uppercase; letter-spacing: .05em; color: var(--muted);
    border-bottom: 1px solid var(--border); white-space: nowrap;
  }
  tbody tr { border-bottom: 1px solid var(--border); transition: background .1s; }
  tbody tr:hover { background: var(--surface); }
  tbody td { padding: 8px 10px; vertical-align: top; max-width: 260px; }
  .badge {
    display: inline-block; padding: 2px 7px; border-radius: 12px;
    font-size: 11px; font-weight: 600; white-space: nowrap;
  }
  .badge-mail       { background: #1e3a5f; color: #93c5fd; }
  .badge-meeting    { background: #1e3a2f; color: #6ee7b7; }
  .badge-appointment{ background: #3b2e1a; color: #fcd34d; }
  .badge-select     { background: #2a2a3e; color: var(--muted); }
  .badge-reply      { background: #1e3a5f; color: #93c5fd; }
  .badge-send       { background: #1e3a2f; color: #6ee7b7; }
  .badge-forward    { background: #2d2a3e; color: #c4b5fd; }
  .badge-close      { background: #3b1a1a; color: #fca5a5; }
  .badge-compose    { background: #1a3b2d; color: #6ee7b7; }
  .badge-save       { background: #2a3b1a; color: #86efac; }
  .badge-meetingAccept   { background: #1a3b25; color: #34d399; }
  .badge-meetingTentative{ background: #3b3b1a; color: #fbbf24; }
  .badge-meetingDecline  { background: #3b1a1a; color: #f87171; }
  .ts { color: var(--muted); font-size: 11px; white-space: nowrap; }
  .subject { font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .recipients { color: var(--muted); font-size: 11px; }
  .widget-path { font-size: 10px; color: var(--muted); line-height: 1.6; }
  .widget-node { display: flex; gap: 4px; flex-wrap: wrap; }
  .widget-node .role { color: var(--accent); }
  .widget-node .wtitle { color: var(--text); }
  .chevron { color: var(--muted); }
  .screenshot-thumb { width: 80px; height: 50px; object-fit: cover; border-radius: 4px; cursor: pointer; border: 1px solid var(--border); }
  .screenshot-thumb:hover { border-color: var(--accent); }
  .no-shot { color: var(--muted); font-style: italic; font-size: 11px; }
  #lightbox {
    display: none; position: fixed; inset: 0; background: rgba(0,0,0,.85);
    z-index: 100; align-items: center; justify-content: center;
  }
  #lightbox.open { display: flex; }
  #lightbox img { max-width: 90vw; max-height: 90vh; border-radius: 8px; }
  #lightbox-close {
    position: absolute; top: 20px; right: 24px; font-size: 28px; cursor: pointer;
    color: var(--text); line-height: 1;
  }
  .empty-row td { text-align: center; padding: 40px; color: var(--muted); }
</style>
</head>
<body>

<header>
  <div id="status-dot"></div>
  <h1>Outlook Capture Dashboard</h1>
  <span id="status-label" style="color:var(--muted);font-size:12px;">Connecting…</span>
  <span id="event-count" style="font-size:12px;">0 events</span>
</header>

<div class="toolbar">
  <input id="filter-input" type="text" placeholder="Filter by subject, action, recipient…">
  <button onclick="clearFilter()">Clear</button>
  <button onclick="scrollToBottom()">↓ Latest</button>
  <button onclick="exportJSON()">Export JSON</button>
</div>

<div id="table-wrap">
  <table id="events-table">
    <thead>
      <tr>
        <th>Time</th>
        <th>Action</th>
        <th>Kind</th>
        <th>Subject</th>
        <th>Recipients</th>
        <th>Widget Path</th>
        <th>Screenshot</th>
      </tr>
    </thead>
    <tbody id="tbody">
      <tr class="empty-row"><td colspan="7">No events yet — waiting for Outlook activity…</td></tr>
    </tbody>
  </table>
</div>

<div id="lightbox">
  <span id="lightbox-close" onclick="closeLightbox()">✕</span>
  <img id="lightbox-img" src="" alt="Screenshot">
</div>

<script>
const tbody = document.getElementById('tbody');
const dot = document.getElementById('status-dot');
const statusLabel = document.getElementById('status-label');
const countEl = document.getElementById('event-count');
const filterInput = document.getElementById('filter-input');
let allEvents = [];
let ws = null;
let reconnectTimer = null;

// ── helpers ──────────────────────────────────────────────────────────────────
function fmtTime(iso) {
  try {
    const d = new Date(iso);
    return d.toLocaleTimeString([], {hour:'2-digit',minute:'2-digit',second:'2-digit'});
  } catch { return iso; }
}

function badgeClass(action) {
  const map = {
    reply:'reply',replyAll:'reply',forward:'forward',send:'send',
    save:'save',compose:'compose',select:'select',close:'close',
    meetingAccept:'meetingAccept',meetingTentative:'meetingTentative',
    meetingDecline:'meetingDecline'
  };
  return 'badge badge-' + (map[action] || 'select');
}

function widgetHTML(nodes) {
  if (!nodes || !nodes.length) return '<span style="color:var(--muted);font-size:11px;">—</span>';
  return nodes.slice(0, 5).map((n, i) => {
    const parts = [];
    if (n.role) parts.push(`<span class="role">${esc(n.role)}</span>`);
    if (n.title) parts.push(`<span class="wtitle">"${esc(n.title)}"</span>`);
    const html = `<span class="widget-node">${parts.join(' ')}</span>`;
    return i < nodes.length - 1 && i < 4 ? html + '<span class="chevron">›</span>' : html;
  }).join('');
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function recipientsHTML(meta) {
  if (!meta) return '—';
  const to = (meta.to || []).join(', ');
  const cc = (meta.cc || []).join(', ');
  let out = to ? `<div>To: ${esc(to)}</div>` : '';
  if (cc) out += `<div style="color:var(--muted)">Cc: ${esc(cc)}</div>`;
  return out || '—';
}

function makeRow(ev) {
  const tr = document.createElement('tr');
  tr.dataset.id = ev.id;
  const shotHtml = ev.screenshotPath
    ? `<img class="screenshot-thumb" src="http://127.0.0.1:8080/${esc(ev.screenshotPath)}"
          onclick="openLightbox(this.src)" alt="screenshot">`
    : '<span class="no-shot">—</span>';
  const kindClass = ev.itemKind === 'mail' ? 'badge-mail'
                  : ev.itemKind === 'meeting' ? 'badge-meeting'
                  : ev.itemKind === 'appointment' ? 'badge-appointment'
                  : 'badge-select';
  tr.innerHTML = `
    <td class="ts">${fmtTime(ev.timestamp)}</td>
    <td><span class="${badgeClass(ev.action)}">${esc(ev.action)}</span></td>
    <td><span class="badge ${kindClass}">${esc(ev.itemKind)}</span></td>
    <td class="subject" title="${esc(ev.metadata?.subject||'')}">${esc(ev.metadata?.subject||'—')}</td>
    <td class="recipients">${recipientsHTML(ev.metadata)}</td>
    <td class="widget-path">${widgetHTML(ev.widgetPath)}</td>
    <td>${shotHtml}</td>`;
  return tr;
}

function renderAll() {
  const q = filterInput.value.toLowerCase();
  const filtered = q ? allEvents.filter(ev =>
    JSON.stringify(ev).toLowerCase().includes(q)
  ) : allEvents;
  tbody.innerHTML = '';
  if (!filtered.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="7">No matching events.</td></tr>';
    return;
  }
  filtered.forEach(ev => tbody.appendChild(makeRow(ev)));
  countEl.textContent = `${allEvents.length} event${allEvents.length!==1?'s':''}`;
}

function addEvent(ev) {
  allEvents.push(ev);
  renderAll();
  scrollToBottom();
}

// ── WebSocket ─────────────────────────────────────────────────────────────────
function connect() {
  clearTimeout(reconnectTimer);
  try { ws && ws.close(); } catch {}
  ws = new WebSocket('ws://127.0.0.1:8765');
  ws.onopen = () => {
    dot.className = 'connected';
    statusLabel.textContent = 'Live';
    ws.send(JSON.stringify({type:'subscribe'}));
  };
  ws.onmessage = ({data}) => {
    try {
      const msg = JSON.parse(data);
      if (msg.type === 'broadcast' && msg.event) addEvent(msg.event);
    } catch {}
  };
  ws.onclose = ws.onerror = () => {
    dot.className = 'error';
    statusLabel.textContent = 'Reconnecting…';
    reconnectTimer = setTimeout(connect, 3000);
  };
}

// ── Initial load ──────────────────────────────────────────────────────────────
async function loadInitial() {
  try {
    const r = await fetch('http://127.0.0.1:8080/api/events');
    const events = await r.json();
    allEvents = events;
    renderAll();
  } catch (e) {
    console.warn('Could not load initial events:', e);
  }
}

// ── UI helpers ────────────────────────────────────────────────────────────────
function scrollToBottom() {
  window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});
}
function clearFilter() { filterInput.value = ''; renderAll(); }
function exportJSON() {
  const a = document.createElement('a');
  a.href = 'data:application/json,' + encodeURIComponent(JSON.stringify(allEvents, null, 2));
  a.download = 'outlook-capture-events.json';
  a.click();
}
function openLightbox(src) {
  document.getElementById('lightbox-img').src = src;
  document.getElementById('lightbox').classList.add('open');
}
function closeLightbox() { document.getElementById('lightbox').classList.remove('open'); }
document.getElementById('lightbox').addEventListener('click', e => {
  if (e.target === e.currentTarget) closeLightbox();
});
filterInput.addEventListener('input', renderAll);

// ── Boot ──────────────────────────────────────────────────────────────────────
loadInitial();
connect();
</script>
</body>
</html>
"""#
