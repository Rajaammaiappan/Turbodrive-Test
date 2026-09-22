# -*- coding: utf-8 -*-
"""HTML templates for tool_hub.py — kept in a separate file to keep the
Flask app itself readable. No external CDN links are used anywhere
(corporate machines may have no internet access) — all CSS/JS is inline.
"""

BASE_CSS = """
:root {
  --bg: #f4f7fb;
  --surface: #ffffff;
  --surface2: #eef3f8;
  --surface3: #e3ebf3;
  --border: #cbd5e1;
  --accent: #2f6fed;
  --accent2: #4f8bff;
  --green: #22c55e;
  --red: #ef4444;
  --amber: #f59e0b;
  --text: #172033;
  --muted: #5f6b7a;
  --mono: 'Consolas', monospace;
}
* { box-sizing: border-box; }
body {
  margin:0; background:var(--bg); color:var(--text);
  font-family: 'Segoe UI', system-ui, sans-serif; font-size:14px;
}
a { color: var(--accent2); }
header {
  display:flex; align-items:center; gap:18px; padding:14px 26px;
  background:linear-gradient(90deg, #ffffff, #edf4fb);
  border-bottom:1px solid var(--border); position:sticky; top:0; z-index:50;
}
.header-logo { display:flex; align-items:center; gap:8px; }
.logo-rr {
  width:34px; height:34px; border-radius:50%; background:#0033a0;
  color:#fff; display:flex; align-items:center; justify-content:center;
  font-weight:800; font-size:12px; border:2px solid #fff;
}
.logo-rr-img { height:34px; width:auto; display:block; }
.logo-alten {
  color:var(--text); font-weight:800; letter-spacing:2px; font-size:14px;
  border-left:1px solid var(--border); padding-left:10px;
}
.logo-alten-img {
  height:28px; width:auto; display:block; margin-left:10px; padding-left:10px;
  border-left:1px solid var(--border);
}
.header-title { flex:1; }
.header-title h1 { margin:0; font-size:22px; font-weight:800; letter-spacing:.5px; display:flex; align-items:baseline; gap:10px; }
.header-title .brand-tagline { font-size:11.5px; font-weight:500; font-style:italic; color:var(--accent2); letter-spacing:0; }
.header-title p { margin:2px 0 0; font-size:12px; color:var(--muted); }
.header-actions { display:flex; align-items:center; gap:10px; }
.header-actions input {
  background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:6px; padding:7px 10px; font-size:13px; width:150px;
}
.icon-btn {
  background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:6px; padding:7px 10px; font-size:13px; cursor:pointer;
  display:flex; align-items:center; gap:6px; text-decoration:none;
}
.icon-btn:hover { border-color:var(--accent2); }
.badge-count {
  background:var(--red); color:#fff; border-radius:10px; font-size:10px;
  padding:1px 6px; margin-left:2px;
}
main { max-width:1400px; margin:0 auto; padding:22px 26px 60px; }

.stats-row { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-bottom:22px; }
.stat-card {
  background:linear-gradient(160deg, var(--surface), var(--surface2));
  border:1px solid var(--border); border-radius:10px; padding:16px 18px;
}
.stat-card .v { font-size:26px; font-weight:800; color:var(--accent2); }
.stat-card .l { font-size:12px; color:var(--muted); margin-top:4px; }

.toolbar {
  display:flex; align-items:center; gap:12px; margin-bottom:16px; flex-wrap:wrap;
}
.toolbar input[type=text] {
  background:var(--surface); border:1px solid var(--border); color:var(--text);
  border-radius:8px; padding:10px 14px; font-size:13px; flex:1; min-width:220px;
}
.chip {
  background:var(--surface); border:1px solid var(--border); color:var(--muted);
  border-radius:20px; padding:7px 14px; font-size:12px; cursor:pointer; white-space:nowrap;
}
.chip.active { background:var(--accent); color:#fff; border-color:var(--accent); }
.chip-row { display:flex; gap:8px; flex-wrap:wrap; }

.section-title {
  font-size:15px; font-weight:700; margin:26px 0 12px; display:flex;
  align-items:center; gap:8px; color:var(--text);
}
.section-title .muted { font-size:12px; color:var(--muted); font-weight:400; }

.quick-launch { display:flex; gap:10px; overflow-x:auto; padding-bottom:6px; }
.ql-card {
  background:var(--surface); border:1px solid var(--border); border-radius:10px;
  padding:10px 14px; min-width:150px; cursor:pointer; flex-shrink:0;
}
.ql-card:hover { border-color:var(--accent2); }
.ql-card .icon { font-size:20px; }
.ql-card .name { font-size:12.5px; font-weight:600; margin-top:4px; }
.ql-card .cnt { font-size:11px; color:var(--muted); }

.cards {
  display:grid; grid-template-columns:repeat(auto-fill, minmax(270px,1fr));
  gap:16px; margin-top:6px;
}
.card {
  background:var(--surface); border:1px solid var(--border); border-radius:12px;
  padding:16px; display:flex; flex-direction:column; gap:10px; position:relative;
  transition:border-color .15s, transform .1s;
}
.card:hover { border-color:var(--accent2); }
.card.disabled { opacity:.55; }
.card .top-row { display:flex; align-items:flex-start; justify-content:flex-end; }
.card .fav-btn {
  background:none; border:none; color:var(--muted); font-size:18px; cursor:pointer;
  padding:0; line-height:1; position:absolute; top:12px; right:14px;
}
.card .fav-btn.active { color:var(--amber); }
.card .desc { font-size:12.5px; color:var(--muted); min-height:32px; }
.card .meta { font-size:11px; color:var(--muted); display:flex; justify-content:space-between; }
.card .status-pill {
  font-size:10px; padding:2px 8px; border-radius:10px; display:inline-block;
  width:fit-content;
}
.status-pill.active { background:rgba(34,197,94,.15); color:var(--green); }
.status-pill.disabled { background:rgba(239,68,68,.15); color:var(--red); }
.card .open-box {
  background:var(--surface2); border:1.5px solid var(--border); border-radius:8px;
  padding:12px 14px; cursor:pointer; display:flex; align-items:center; gap:10px;
  transition:border-color .15s, background .15s;
}
.card .open-box:hover { border-color:var(--accent2); background:var(--surface3); }
.card .open-box .icon { font-size:22px; width:38px; height:38px; border-radius:8px;
  background:var(--surface3); display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.card .open-box .title-wrap { flex:1; min-width:0; }
.card .open-box .name { font-size:15px; font-weight:700; }
.card .open-box .cat { font-size:10.5px; color:var(--accent2); text-transform:uppercase; letter-spacing:.5px; }
.card .open-box .go { color:var(--muted); font-size:16px; flex-shrink:0; }
.card .actions { display:flex; gap:8px; margin-top:2px; justify-content:flex-end; }
.card .actions button {
  background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:7px; width:36px; height:34px; font-size:14px; cursor:pointer;
  display:flex; align-items:center; justify-content:center; position:relative;
}
.card .actions button:hover { border-color:var(--accent2); color:var(--accent2); }
.card .actions button .num {
  position:absolute; top:-6px; right:-6px; background:var(--accent); color:#fff;
  border-radius:50%; width:15px; height:15px; font-size:9px; display:flex;
  align-items:center; justify-content:center; font-weight:700;
}

.empty-state { text-align:center; color:var(--muted); padding:50px 0; font-size:13px; }

.analytics-grid { display:grid; grid-template-columns:1.3fr 1fr; gap:16px; }
.panel {
  background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:16px;
}
.panel h3 { margin:0 0 12px; font-size:13.5px; }
.bar-row { display:flex; align-items:center; gap:8px; margin-bottom:8px; font-size:12px; }
.bar-row .lbl { width:130px; flex-shrink:0; color:var(--muted); overflow:hidden;
  text-overflow:ellipsis; white-space:nowrap; }
.bar-track { flex:1; background:var(--surface2); border-radius:4px; height:16px; overflow:hidden; }
.bar-fill { height:100%; background:linear-gradient(90deg, var(--accent), var(--accent2)); }
.bar-row .val { width:28px; text-align:right; color:var(--text); font-weight:600; }

.activity-list { max-height:340px; overflow-y:auto; }
.activity-item {
  display:flex; justify-content:space-between; gap:10px; padding:7px 0;
  border-bottom:1px solid var(--border); font-size:12px;
}
.activity-item:last-child { border-bottom:none; }
.activity-item .who { color:var(--text); font-weight:600; }
.activity-item .what { color:var(--muted); }
.activity-item .when { color:var(--muted); font-family:var(--mono); font-size:11px; white-space:nowrap;}

.sparkline { display:flex; align-items:flex-end; gap:3px; height:60px; margin-top:8px; }
.spark-bar { flex:1; background:var(--accent2); border-radius:2px 2px 0 0; min-height:2px; }

/* Modal */
.modal-overlay {
  display:none; position:fixed; inset:0; background:rgba(0,0,0,.6);
  z-index:200; align-items:center; justify-content:center; padding:20px;
}
.modal-overlay.open { display:flex; }
.modal {
  background:var(--surface); border:1px solid var(--border); border-radius:12px;
  width:100%; max-width:760px; max-height:88vh; overflow-y:auto; padding:22px 24px;
}
.modal h2 { margin:0 0 4px; font-size:18px; }
.modal .sub { color:var(--muted); font-size:12px; margin-bottom:16px; }
.modal .close-x {
  float:right; background:none; border:none; color:var(--muted); font-size:20px; cursor:pointer;
}
.about-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:10px; }
.about-block h4 { margin:0 0 6px; font-size:12.5px; color:var(--accent2); text-transform:uppercase; letter-spacing:.5px;}
.about-block ul { margin:4px 0; padding-left:18px; font-size:13px; }
.about-block p { margin:4px 0; font-size:13px; color:var(--text); }
.change-history { margin-top:16px; }
.ch-row { display:flex; gap:10px; padding:6px 0; border-bottom:1px solid var(--border); font-size:12.5px; }
.ch-row .v { color:var(--accent2); font-weight:700; width:50px; flex-shrink:0; }
.ch-row .d { color:var(--muted); width:90px; flex-shrink:0; font-family:var(--mono); font-size:11px; }

/* DB dashboard modal */
.modal.wide { max-width:1100px; }
.db-toolbar { display:flex; gap:10px; margin:10px 0 14px; flex-wrap:wrap; align-items:center; }
.db-toolbar input[type=text] {
  background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:6px; padding:8px 10px; font-size:13px; flex:1; min-width:200px;
}
.db-toolbar button {
  background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:6px; padding:8px 12px; font-size:12.5px; cursor:pointer;
}
.db-toolbar button.green { background:var(--green); border-color:var(--green); color:#04240f; font-weight:700;}
.db-toolbar button.blue { background:var(--accent); border-color:var(--accent); color:#fff; font-weight:700;}
.db-stats { display:flex; gap:14px; margin-bottom:12px; flex-wrap:wrap; }
.db-stat { background:var(--surface2); border:1px solid var(--border); border-radius:8px;
  padding:8px 14px; font-size:12px; }
.db-stat b { display:block; font-size:16px; color:var(--accent2); }
.table-wrap { overflow:auto; max-height:420px; border:1px solid var(--border); border-radius:8px; }
table.data-tbl { border-collapse:collapse; width:100%; font-size:12px; }
table.data-tbl th {
  background:var(--surface3); position:sticky; top:0; text-align:left; padding:8px 10px;
  border-bottom:1px solid var(--border); white-space:nowrap;
}
table.data-tbl td { padding:7px 10px; border-bottom:1px solid var(--border); white-space:nowrap;
  max-width:220px; overflow:hidden; text-overflow:ellipsis; }
table.data-tbl tr:hover td { background:var(--surface2); }
.pager { display:flex; gap:10px; align-items:center; margin-top:10px; font-size:12px; }
.pager button { background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:6px; padding:6px 10px; cursor:pointer; }

footer { text-align:center; color:var(--muted); font-size:11px; padding:20px; }
.toast {
  position:fixed; bottom:20px; right:20px; background:var(--surface3); border:1px solid var(--border);
  color:var(--text); padding:10px 16px; border-radius:8px; font-size:13px; z-index:300; display:none;
}
"""

DASHBOARD_JS = r"""
var CURRENT_USER = localStorage.getItem('hub_user') || '';
var ALL_TOOLS = [];
var FAVORITES = [];
var ACTIVE_CATEGORY = 'All';
var DB_STATE = { tool_id:null, page:0, q:'' };

document.addEventListener('DOMContentLoaded', function() {
  loadTheme();
  document.getElementById('userName').value = CURRENT_USER;
  loadEverything();
  document.getElementById('userName').addEventListener('change', function(e){
    CURRENT_USER = e.target.value.trim();
    localStorage.setItem('hub_user', CURRENT_USER);
    loadEverything();
  });
  document.getElementById('searchBox').addEventListener('input', renderCards);
  document.getElementById('favOnly').addEventListener('change', renderCards);
});

function toast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.style.display = 'block';
  setTimeout(function(){ t.style.display = 'none'; }, 2200);
}

function loadEverything() {
  Promise.all([
    fetch('/api/tools').then(r=>r.json()),
    fetch('/api/stats').then(r=>r.json()),
    fetch('/api/analytics').then(r=>r.json()),
    fetch('/api/favorites?user=' + encodeURIComponent(CURRENT_USER||'anonymous')).then(r=>r.json())
  ]).then(function(res){
    ALL_TOOLS = res[0].tools;
    FAVORITES = res[3].favorites;
    renderStats(res[1]);
    renderCategories();
    renderCards();
    renderQuickLaunch();
    renderAnalytics(res[2]);
    renderNotifications();
  });
}

function renderStats(s) {
  document.getElementById('statTools').textContent = s.total_tools;
  document.getElementById('statUsers').textContent = s.active_users;
  document.getElementById('statAuto').textContent = s.total_automations;
  document.getElementById('statHours').textContent = s.hours_saved;
}

function renderCategories() {
  var cats = ['All'].concat(Array.from(new Set(ALL_TOOLS.map(t=>t.category))).sort());
  var wrap = document.getElementById('catChips');
  wrap.innerHTML = cats.map(function(c){
    return '<div class="chip' + (c===ACTIVE_CATEGORY?' active':'') + '" data-cat="' + c + '">' + c + '</div>';
  }).join('');
  Array.from(wrap.children).forEach(function(el){
    el.addEventListener('click', function(){
      ACTIVE_CATEGORY = el.getAttribute('data-cat');
      renderCategories(); renderCards();
    });
  });
}

function renderQuickLaunch() {
  var favTools = ALL_TOOLS.filter(t => FAVORITES.indexOf(t.id) !== -1);
  var topUsed = ALL_TOOLS.slice().sort((a,b)=>b.usage_count-a.usage_count).slice(0,6);
  var combined = favTools.concat(topUsed.filter(t=>favTools.indexOf(t)===-1)).slice(0,8);
  var wrap = document.getElementById('quickLaunch');
  if (!combined.length) { wrap.innerHTML = '<div class="empty-state">No quick-launch tools yet — use a tool or star a favorite.</div>'; return; }
  wrap.innerHTML = combined.map(function(t){
    return '<div class="ql-card" onclick="openTool(\'' + t.id + '\')">' +
      '<div class="icon">' + t.icon + '</div>' +
      '<div class="name">' + t.tool_name + '</div>' +
      '<div class="cnt">' + t.usage_count + ' uses</div></div>';
  }).join('');
}

function renderCards() {
  var q = document.getElementById('searchBox').value.toLowerCase();
  var favOnly = document.getElementById('favOnly').checked;
  var list = ALL_TOOLS.filter(function(t){
    if (ACTIVE_CATEGORY !== 'All' && t.category !== ACTIVE_CATEGORY) return false;
    if (favOnly && FAVORITES.indexOf(t.id) === -1) return false;
    if (q && (t.tool_name.toLowerCase().indexOf(q)===-1 && t.description.toLowerCase().indexOf(q)===-1 && t.owner.toLowerCase().indexOf(q)===-1)) return false;
    return true;
  });
  var wrap = document.getElementById('cards');
  if (!list.length) { wrap.innerHTML = '<div class="empty-state">No tools match your search/filter.</div>'; return; }
  wrap.innerHTML = list.map(cardHtml).join('');
}

function cardHtml(t) {
  var isFav = FAVORITES.indexOf(t.id) !== -1;
  var disabled = t.status !== 'active';
  var openAttr = disabled ? '' : ' onclick="openTool(\'' + t.id + '\')"';
  return '<div class="card' + (disabled?' disabled':'') + '">' +
    '<button class="fav-btn' + (isFav?' active':'') + '" onclick="toggleFav(\'' + t.id + '\', this)">' + (isFav?'\u2605':'\u2606') + '</button>' +
    '<div class="open-box"' + openAttr + ' title="Click to open ' + t.tool_name + '">' +
      '<div class="icon">' + t.icon + '</div>' +
      '<div class="title-wrap">' +
        '<div class="name">' + t.tool_name + '</div>' +
        '<div class="cat">' + t.category + '</div>' +
      '</div>' +
      '<div class="go">' + (disabled ? '\u26D4' : '\u25B6') + '</div>' +
    '</div>' +
    '<div class="desc">' + t.description + '</div>' +
    '<span class="status-pill ' + (disabled?'disabled':'active') + '">' + (disabled?'Disabled':'Active') + '</span>' +
    '<div class="meta"><span>Owner: ' + t.owner + '</span><span>v' + t.version + '</span></div>' +
    '<div class="meta"><span>Updated: ' + t.last_updated + '</span><span>' + t.usage_count + ' uses</span></div>' +
    '<div class="actions">' +
      '<button onclick="openGuide(\'' + t.id + '\')" title="User Guide">\uD83D\uDCD8<span class="num">1</span></button>' +
      '<button onclick="openAbout(\'' + t.id + '\')" title="About Tool">\u2139<span class="num">2</span></button>' +
      '<button onclick="openDbDashboard(\'' + t.id + '\')" title="Database / Dashboard">\uD83D\uDCCA<span class="num">3</span></button>' +
    '</div>' +
  '</div>';
}

function toggleFav(id, el) {
  fetch('/api/favorites/toggle', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({user: CURRENT_USER||'anonymous', tool_id: id})})
    .then(r=>r.json()).then(function(j){
      if (j.favorite) { FAVORITES.push(id); } else { FAVORITES = FAVORITES.filter(x=>x!==id); }
      renderCards(); renderQuickLaunch();
    });
}

function logEvent(id, action) {
  fetch('/api/log', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({tool_id:id, action:action, user: CURRENT_USER||'anonymous'})});
}

function openTool(id) {
  var t = ALL_TOOLS.find(x=>x.id===id);
  if (!t) return;
  if (t.status !== 'active') { toast('This tool is currently disabled.'); return; }
  logEvent(id, 'open');
  fetch('/api/resolve_open', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({tool_id:id, field:'tool_url'})})
    .then(r=>r.json()).then(function(j){
      if (!j.ok) { toast(j.error || 'Could not open this tool.'); return; }
      if (j.mode === 'url') window.open(j.url, '_blank');
      else toast('Launching ' + t.tool_name + '…');
    })
    .catch(function(){ toast('Could not reach the hub server.'); });
  setTimeout(loadEverything, 600);
}
function openGuide(id) {
  var t = ALL_TOOLS.find(x=>x.id===id);
  if (!t) return;
  logEvent(id, 'guide');
  fetch('/api/resolve_open', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({tool_id:id, field:'guide_url'})})
    .then(r=>r.json()).then(function(j){
      if (!j.ok) { toast(j.error || 'No user guide configured for this tool yet.'); return; }
      if (j.mode === 'url') window.open(j.url, '_blank');
      else toast('Opening user guide…');
    })
    .catch(function(){ toast('Could not reach the hub server.'); });
}

function openAbout(id) {
  var t = ALL_TOOLS.find(x=>x.id===id);
  if (!t) return;
  logEvent(id, 'about');
  document.getElementById('aboutTitle').textContent = t.icon + ' ' + t.tool_name;
  document.getElementById('aboutSub').textContent = t.category + ' \u2022 v' + t.version + ' \u2022 Owner: ' + t.owner;
  var feats = (t.features||[]).map(f=>'<li>'+f+'</li>').join('') || '<li>—</li>';
  var bens = (t.benefits||[]).map(f=>'<li>'+f+'</li>').join('') || '<li>—</li>';
  document.getElementById('aboutBody').innerHTML =
    '<div class="about-grid">' +
      '<div class="about-block"><h4>Purpose</h4><p>' + (t.purpose||'—') + '</p>' +
        '<h4 style="margin-top:12px">Scope</h4><p>' + (t.scope||'—') + '</p></div>' +
      '<div class="about-block"><h4>Features</h4><ul>' + feats + '</ul></div>' +
      '<div class="about-block"><h4>Benefits</h4><ul>' + bens + '</ul></div>' +
      '<div class="about-block"><h4>Contact</h4><p>Owner: ' + t.owner + '<br>Email: ' + (t.contact||'—') + '</p></div>' +
    '</div>' +
    '<div class="change-history"><h4 style="color:var(--accent2);font-size:12.5px;text-transform:uppercase;">Change History</h4>' +
    (t.change_history||[]).map(c=>'<div class="ch-row"><div class="v">v'+c.version+'</div><div class="d">'+c.date+'</div><div>'+c.notes+'</div></div>').join('') +
    '</div>';
  document.getElementById('aboutOverlay').classList.add('open');
}

function closeModal(id) { document.getElementById(id).classList.remove('open'); }

function openDbDashboard(id) {
  var t = ALL_TOOLS.find(x=>x.id===id);
  if (!t) return;
  logEvent(id, 'dashboard');
  if (t.db_viewer_url) {
    // Tool has its own dedicated DB viewer app -> open that instead of the built-in dashboard.
    window.open(t.db_viewer_url, '_blank');
    return;
  }
  DB_STATE = { tool_id:id, page:0, q:'' };
  document.getElementById('dbTitle').textContent = t.icon + ' ' + t.tool_name + ' — Data Dashboard';
  document.getElementById('dbSearch').value = '';
  document.getElementById('dbOverlay').classList.add('open');
  loadDbData();
}

function loadDbData() {
  var url = '/api/tool_db/' + DB_STATE.tool_id + '?q=' + encodeURIComponent(DB_STATE.q) + '&page=' + DB_STATE.page;
  fetch(url).then(r=>r.json()).then(function(j){
    if (!j.ok) {
      document.getElementById('dbStats').innerHTML = '';
      document.getElementById('dbTableWrap').innerHTML = '<div class="empty-state">' + (j.error || 'No data available for this tool yet.') + '</div>';
      document.getElementById('dbPager').innerHTML = '';
      return;
    }
    document.getElementById('dbStats').innerHTML =
      '<div class="db-stat"><b>' + j.total + '</b>Total Records</div>' +
      '<div class="db-stat"><b>' + j.rows.length + '</b>Rows Shown</div>' +
      '<div class="db-stat"><b>' + (j.page+1) + '</b>Page</div>' +
      '<div class="db-stat"><b>' + (j.recent_activity && j.recent_activity[0] ? j.recent_activity[0] : '—') + '</b>Most Recent Record</div>';
    if (!j.rows.length) {
      document.getElementById('dbTableWrap').innerHTML = '<div class="empty-state">No matching records.</div>';
    } else {
      var head = '<tr>' + j.columns.map(c=>'<th>'+c+'</th>').join('') + '</tr>';
      var body = j.rows.map(function(row){
        return '<tr>' + j.columns.map(function(c){
          var v = row[c]; v = (v===null||v===undefined) ? '' : String(v);
          return '<td title="'+v.replace(/"/g,'&quot;')+'">'+v+'</td>';
        }).join('') + '</tr>';
      }).join('');
      document.getElementById('dbTableWrap').innerHTML =
        '<div class="table-wrap"><table class="data-tbl"><thead>'+head+'</thead><tbody>'+body+'</tbody></table></div>';
    }
    var totalPages = Math.max(1, Math.ceil(j.total / j.page_size));
    document.getElementById('dbPager').innerHTML =
      '<button onclick="dbPage(-1)" ' + (j.page<=0?'disabled':'') + '>\u2190 Prev</button>' +
      '<span>Page ' + (j.page+1) + ' of ' + totalPages + '</span>' +
      '<button onclick="dbPage(1)" ' + (j.page+1>=totalPages?'disabled':'') + '>Next \u2192</button>';
  });
}
function dbPage(delta) { DB_STATE.page = Math.max(0, DB_STATE.page + delta); loadDbData(); }
function dbSearch() { DB_STATE.q = document.getElementById('dbSearch').value; DB_STATE.page = 0; loadDbData(); }
function dbExport(fmt) {
  window.location = '/api/tool_db/' + DB_STATE.tool_id + '/export?format=' + fmt + '&q=' + encodeURIComponent(DB_STATE.q);
}

function renderAnalytics(a) {
  var maxUsed = Math.max(1, ...a.most_used.map(x=>x.count));
  document.getElementById('mostUsedWrap').innerHTML = a.most_used.length ? a.most_used.map(function(m){
    var pct = Math.round((m.count/maxUsed)*100);
    return '<div class="bar-row"><div class="lbl">' + m.tool_name + '</div>' +
      '<div class="bar-track"><div class="bar-fill" style="width:'+pct+'%"></div></div>' +
      '<div class="val">' + m.count + '</div></div>';
  }).join('') : '<div class="empty-state">No usage yet.</div>';

  document.getElementById('recentActivity').innerHTML = a.recent.length ? a.recent.map(function(r){
    return '<div class="activity-item"><div><span class="who">' + r.user + '</span> <span class="what">' + actionLabel(r.action) + ' ' + r.tool_name + '</span></div>' +
      '<div class="when">' + r.ts.replace('T',' ') + '</div></div>';
  }).join('') : '<div class="empty-state">No activity yet.</div>';

  var maxDaily = Math.max(1, ...a.daily.map(x=>x.count));
  document.getElementById('sparkline').innerHTML = a.daily.map(function(d){
    var h = Math.max(2, Math.round((d.count/maxDaily)*56));
    return '<div class="spark-bar" style="height:'+h+'px" title="'+d.date+': '+d.count+'"></div>';
  }).join('');

  var maxCat = Math.max(1, ...a.categories.map(x=>x.count));
  document.getElementById('catBreakdown').innerHTML = a.categories.map(function(c){
    var pct = Math.round((c.count/maxCat)*100);
    return '<div class="bar-row"><div class="lbl">' + c.category + '</div>' +
      '<div class="bar-track"><div class="bar-fill" style="width:'+pct+'%"></div></div>' +
      '<div class="val">' + c.count + '</div></div>';
  }).join('');
}
function actionLabel(a) {
  return {open:'opened', guide:'viewed guide for', about:'viewed About for', dashboard:'viewed dashboard for'}[a] || a;
}

var THEME_PRESETS = {
  slate: {accent:'#d97706', accent2:'#f59e0b', bg:'#f4f7fb', surface:'#ffffff', surface2:'#eef3f8', surface3:'#e3ebf3', border:'#cbd5e1', text:'#172033', muted:'#5f6b7a'},
  white: {accent:'#2563eb', accent2:'#3b82f6', bg:'#ffffff', surface:'#ffffff', surface2:'#f1f5f9', surface3:'#e2e8f0', border:'#cbd5e1', text:'#172033', muted:'#64748b'},
  ice: {accent:'#0284c7', accent2:'#38bdf8', bg:'#eff8ff', surface:'#ffffff', surface2:'#e0f2fe', surface3:'#bae6fd', border:'#bae6fd', text:'#082f49', muted:'#52748a'},
  solar: {accent:'#ea580c', accent2:'#fb923c', bg:'#fff7ed', surface:'#ffffff', surface2:'#ffedd5', surface3:'#fed7aa', border:'#fdba74', text:'#431407', muted:'#9a6048'},
  fresh: {accent:'#059669', accent2:'#10b981', bg:'#f0fdf4', surface:'#ffffff', surface2:'#dcfce7', surface3:'#bbf7d0', border:'#a7f3d0', text:'#064e3b', muted:'#4f806f'},
  gulf: {accent:'#d97706', accent2:'#fb923c', bg:'#fffbeb', surface:'#ffffff', surface2:'#fef3c7', surface3:'#fde68a', border:'#fcd34d', text:'#451a03', muted:'#8a6a35'},
  corporate: {accent:'#0284c7', accent2:'#38bdf8', bg:'#f0f9ff', surface:'#ffffff', surface2:'#e0f2fe', surface3:'#bae6fd', border:'#bae6fd', text:'#082f49', muted:'#52748a'},
  emerald: {accent:'#059669', accent2:'#34d399', bg:'#f0fdf4', surface:'#ffffff', surface2:'#dcfce7', surface3:'#bbf7d0', border:'#a7f3d0', text:'#064e3b', muted:'#4f806f'},
  sunburst: {accent:'#ea580c', accent2:'#fb923c', bg:'#fff7ed', surface:'#ffffff', surface2:'#ffedd5', surface3:'#fed7aa', border:'#fdba74', text:'#431407', muted:'#9a6048'},
  dark: {accent:'#2f6fed', accent2:'#4f8bff', bg:'#0b0e14', surface:'#12161f', surface2:'#1a2030', surface3:'#212940', border:'#262f42', text:'#e7ecf5', muted:'#8b96ac'}
};

function setThemeColors(theme, save) {
  var root = document.documentElement;
  Object.keys(theme).forEach(function(key){ root.style.setProperty('--' + key, theme[key]); });
  ['themeAccent','themeAccent2','themeBg','themeSurface'].forEach(function(id, index){
    var input = document.getElementById(id);
    if (input) input.value = [theme.accent, theme.accent2, theme.bg, theme.surface][index];
  });
  if (save) localStorage.setItem('engage_theme', JSON.stringify(theme));
}

function applyThemePreset(name) {
  var preset = THEME_PRESETS[name];
  if (!preset) return;
  setThemeColors(preset, true);
  document.querySelectorAll('.theme-choice').forEach(function(button){
    button.classList.toggle('selected', button.getAttribute('data-preset') === name || button.getAttribute('data-mode') === name);
  });
}

function updateCustomTheme() {
  var current = JSON.parse(localStorage.getItem('engage_theme') || '{}');
  current.accent = document.getElementById('themeAccent').value;
  current.accent2 = document.getElementById('themeAccent2').value;
  current.bg = document.getElementById('themeBg').value;
  current.surface = document.getElementById('themeSurface').value;
  setThemeColors(current, true);
  document.querySelectorAll('.theme-choice').forEach(function(button){ button.classList.remove('selected'); });
}

function toggleThemePanel() { document.getElementById('themePanel').classList.toggle('open'); }
function resetTheme() { applyThemePreset('slate'); }
function loadTheme() {
  var saved = localStorage.getItem('engage_theme');
  if (saved) {
    try { setThemeColors(JSON.parse(saved), false); } catch (e) { applyThemePreset('slate'); }
  } else { applyThemePreset('slate'); }
  ['themeAccent','themeAccent2','themeBg','themeSurface'].forEach(function(id){
    document.getElementById(id).addEventListener('input', updateCustomTheme);
  });
}

function renderNotifications() {
  var since = new Date(); since.setDate(since.getDate()-30);
  var recentlyUpdated = ALL_TOOLS.filter(function(t){
    var d = new Date(t.last_updated); return d >= since;
  }).sort((a,b)=> new Date(b.last_updated) - new Date(a.last_updated));
  var badge = document.getElementById('notifBadge');
  badge.textContent = recentlyUpdated.length;
  badge.style.display = recentlyUpdated.length ? 'inline-block' : 'none';
  document.getElementById('notifList').innerHTML = recentlyUpdated.length ?
    recentlyUpdated.map(t => '<div class="activity-item"><div><span class="who">'+t.tool_name+'</span> <span class="what">updated to v'+t.version+'</span></div><div class="when">'+t.last_updated+'</div></div>').join('') :
    '<div class="empty-state">No recent updates.</div>';
}
function toggleNotif() {
  var p = document.getElementById('notifPanel');
  p.style.display = (p.style.display === 'block') ? 'none' : 'block';
}
"""

PAGE_DASHBOARD = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ENGAGE — Engineering Gateway for Automation, Guidance & Enterprise Tools</title>
<style>""" + BASE_CSS + """
.notif-wrap { position:relative; }
#notifPanel {
  display:none; position:absolute; right:0; top:42px; width:320px; max-height:400px;
  overflow-y:auto; background:var(--surface); border:1px solid var(--border);
  border-radius:10px; padding:12px; box-shadow:0 10px 30px rgba(0,0,0,.5); z-index:100;
}
.theme-float-btn {
  position:fixed; right:22px; bottom:22px; z-index:180; padding:12px 18px;
  border:2px solid var(--amber); border-radius:24px; background:var(--surface);
  color:var(--text); font-weight:700; cursor:pointer; box-shadow:0 8px 24px rgba(15,23,42,.18);
}
.theme-float-btn:hover { background:var(--surface2); }
.theme-panel {
  display:none; position:fixed; right:22px; bottom:78px; z-index:190; width:310px;
  max-height:calc(100vh - 100px); overflow-y:auto; padding:16px;
  background:var(--surface); border:1px solid var(--border); border-radius:12px;
  box-shadow:0 16px 40px rgba(15,23,42,.25);
}
.theme-panel.open { display:block; }
.theme-panel h2 { margin:0 0 14px; font-size:15px; }
.theme-group { margin-top:16px; }
.theme-group h3 { margin:0 0 9px; font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.4px; }
.theme-presets, .theme-modes { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
.theme-choice { min-height:48px; padding:7px 6px; border:1px solid var(--border); border-radius:8px;
  background:var(--surface2); color:var(--text); font-size:11px; font-weight:700; cursor:pointer; }
.theme-choice:hover, .theme-choice.selected { border-color:var(--amber); }
.theme-dots { display:flex; justify-content:center; gap:4px; margin-bottom:4px; }
.theme-dot { width:10px; height:10px; border-radius:50%; border:1px solid var(--border); }
.theme-color-row { display:flex; align-items:center; justify-content:space-between; gap:8px; margin:10px 0; font-size:12px; font-weight:600; }
.theme-color-row input[type=color] { width:34px; height:26px; padding:1px; border:1px solid var(--border); background:var(--surface2); cursor:pointer; }
.theme-reset { width:100%; margin-top:12px; padding:8px; border:1px solid var(--border); border-radius:7px;
  background:var(--surface2); color:var(--text); cursor:pointer; }
@media (max-width:600px) {
  .theme-float-btn { right:12px; bottom:12px; }
  .theme-panel { right:12px; bottom:64px; width:calc(100vw - 24px); }
}
</style>
</head>
<body>
<header>
  <div class="header-logo">
    {% if rr_logo %}<img class="logo-rr-img" src="{{ rr_logo }}" alt="Rolls-Royce">{% else %}<div class="logo-rr">RR</div>{% endif %}
    {% if alten_logo %}<img class="logo-alten-img" src="{{ alten_logo }}" alt="ALTEN">{% else %}<div class="logo-alten">ALTEN</div>{% endif %}
  </div>
  <div class="header-title">
    <h1>ENGAGE <span class="brand-tagline">Launch. Automate. Accelerate.</span></h1>
    <p>Engineering Gateway for Automation, Guidance &amp; Enterprise Tools — Internal Use Only</p>
  </div>
  <div class="header-actions">
    <input type="text" id="userName" placeholder="Your name…">
    <div class="notif-wrap">
      <button class="icon-btn" onclick="toggleNotif()">\U0001F514 <span id="notifBadge" class="badge-count" style="display:none">0</span></button>
      <div id="notifPanel"><h3 style="margin:4px 0 10px;font-size:13px;">Recent tool updates</h3><div id="notifList"></div></div>
    </div>
    <a class="icon-btn" href="/admin">\u2699 Admin</a>
  </div>
</header>
<main>
  <div class="stats-row">
    <div class="stat-card"><div class="v" id="statTools">–</div><div class="l">Total Tools</div></div>
    <div class="stat-card"><div class="v" id="statUsers">–</div><div class="l">Active Users</div></div>
    <div class="stat-card"><div class="v" id="statAuto">–</div><div class="l">Total Automations Run</div></div>
    <div class="stat-card"><div class="v" id="statHours">–</div><div class="l">Est. Hours Saved</div></div>
  </div>

  <div class="toolbar">
    <input type="text" id="searchBox" placeholder="\U0001F50D Search tools by name, description, or owner…">
    <label style="display:flex;align-items:center;gap:6px;font-size:12.5px;color:var(--muted);">
      <input type="checkbox" id="favOnly"> Favorites only
    </label>
  </div>
  <div class="chip-row" id="catChips"></div>

  <div class="section-title">\u26A1 Quick Launch <span class="muted">— your favorites and most-used tools</span></div>
  <div class="quick-launch" id="quickLaunch"></div>

  <div class="section-title">\U0001F9F0 All Tools</div>
  <div class="cards" id="cards"></div>

  <div class="section-title">\U0001F4C8 Usage Analytics <span class="muted">— for management visibility</span></div>
  <div class="analytics-grid">
    <div class="panel">
      <h3>Most Used Tools</h3>
      <div id="mostUsedWrap"></div>
      <h3 style="margin-top:18px">Usage — Last 14 Days</h3>
      <div class="sparkline" id="sparkline"></div>
      <h3 style="margin-top:18px">Tools by Category</h3>
      <div id="catBreakdown"></div>
    </div>
    <div class="panel">
      <h3>Recent Access History</h3>
      <div class="activity-list" id="recentActivity"></div>
    </div>
  </div>
</main>
<footer>&copy; 2026 Alten-Rolls-Royce. All rights reserved. Confidential &ndash; Internal Use Only.</footer>

<button class="theme-float-btn" onclick="toggleThemePanel()">&#127912; Theme &amp; Colors</button>
<aside class="theme-panel" id="themePanel" aria-label="Theme and Colors">
  <h2>&#127912; Theme Customizer <button class="close-x" onclick="toggleThemePanel()">&times;</button></h2>
  <div class="theme-group">
    <h3>Quick Color Presets</h3>
    <div class="theme-presets">
      <button class="theme-choice" data-preset="gulf" onclick="applyThemePreset('gulf')"><span class="theme-dots"><i class="theme-dot" style="background:#f59e0b"></i><i class="theme-dot" style="background:#fb923c"></i></span>Golden Brand</button>
      <button class="theme-choice" data-preset="corporate" onclick="applyThemePreset('corporate')"><span class="theme-dots"><i class="theme-dot" style="background:#0284c7"></i><i class="theme-dot" style="background:#38bdf8"></i></span>Corporate Blue</button>
      <button class="theme-choice" data-preset="emerald" onclick="applyThemePreset('emerald')"><span class="theme-dots"><i class="theme-dot" style="background:#059669"></i><i class="theme-dot" style="background:#34d399"></i></span>Eco Emerald</button>
      <button class="theme-choice" data-preset="sunburst" onclick="applyThemePreset('sunburst')"><span class="theme-dots"><i class="theme-dot" style="background:#ea580c"></i><i class="theme-dot" style="background:#fb923c"></i></span>Sunburst</button>
    </div>
  </div>
  <div class="theme-group">
    <h3>Light Theme Combinations</h3>
    <div class="theme-presets">
      <button class="theme-choice" data-preset="slate" onclick="applyThemePreset('slate')"><span class="theme-dots"><i class="theme-dot" style="background:#f59e0b"></i><i class="theme-dot" style="background:#f8fafc"></i></span>Clean Slate Gold</button>
      <button class="theme-choice" data-preset="ice" onclick="applyThemePreset('ice')"><span class="theme-dots"><i class="theme-dot" style="background:#0284c7"></i><i class="theme-dot" style="background:#f8fafc"></i></span>Ice Blue Tech</button>
      <button class="theme-choice" data-preset="solar" onclick="applyThemePreset('solar')"><span class="theme-dots"><i class="theme-dot" style="background:#ea580c"></i><i class="theme-dot" style="background:#fff7ed"></i></span>Warm Solar Light</button>
      <button class="theme-choice" data-preset="fresh" onclick="applyThemePreset('fresh')"><span class="theme-dots"><i class="theme-dot" style="background:#059669"></i><i class="theme-dot" style="background:#f0fdf4"></i></span>Emerald Fresh Light</button>
    </div>
  </div>
  <div class="theme-group">
    <h3>Custom Brand Colors</h3>
    <label class="theme-color-row">Primary Accent Color <input type="color" id="themeAccent"></label>
    <label class="theme-color-row">Secondary Highlight <input type="color" id="themeAccent2"></label>
    <label class="theme-color-row">Page Background <input type="color" id="themeBg"></label>
    <label class="theme-color-row">Card / Section Background <input type="color" id="themeSurface"></label>
  </div>
  <div class="theme-group">
    <h3>Page Background Mode</h3>
    <div class="theme-modes">
      <button class="theme-choice" data-mode="slate" onclick="applyThemePreset('slate')">Slate</button>
      <button class="theme-choice" data-mode="white" onclick="applyThemePreset('white')">Pure White</button>
      <button class="theme-choice" data-mode="solar" onclick="applyThemePreset('solar')">Warm Solar</button>
      <button class="theme-choice" data-mode="ice" onclick="applyThemePreset('ice')">Ice Blue</button>
      <button class="theme-choice" data-mode="dark" onclick="applyThemePreset('dark')">Dark Mode</button>
    </div>
  </div>
  <button class="theme-reset" onclick="resetTheme()">Reset to Default</button>
</aside>

<!-- About modal -->
<div class="modal-overlay" id="aboutOverlay">
  <div class="modal">
    <button class="close-x" onclick="closeModal('aboutOverlay')">&times;</button>
    <h2 id="aboutTitle"></h2>
    <div class="sub" id="aboutSub"></div>
    <div id="aboutBody"></div>
  </div>
</div>

<!-- DB dashboard modal -->
<div class="modal-overlay" id="dbOverlay">
  <div class="modal wide">
    <button class="close-x" onclick="closeModal('dbOverlay')">&times;</button>
    <h2 id="dbTitle"></h2>
    <div class="db-toolbar">
      <input type="text" id="dbSearch" placeholder="Search records…" onkeydown="if(event.key==='Enter') dbSearch()">
      <button onclick="dbSearch()">Search</button>
      <button class="green" onclick="dbExport('xlsx')">\u2B07 Export Excel</button>
      <button onclick="dbExport('csv')">\u2B07 Export CSV</button>
    </div>
    <div class="db-stats" id="dbStats"></div>
    <div id="dbTableWrap"></div>
    <div class="pager" id="dbPager"></div>
  </div>
</div>

<div class="toast" id="toast"></div>
<script>""" + DASHBOARD_JS + """</script>
</body>
</html>"""


ADMIN_JS = r"""
function closeModal(id) { document.getElementById(id).classList.remove('open'); }
var PASSCODE = sessionStorage.getItem('hub_admin_pass') || '';
var TOOLS = [];

document.addEventListener('DOMContentLoaded', function(){
  if (PASSCODE) { unlock(); } else { document.getElementById('loginBox').style.display='block'; }
  document.getElementById('loginBtn').addEventListener('click', function(){
    PASSCODE = document.getElementById('passInput').value;
    fetch('/api/admin/login', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({passcode: PASSCODE})}).then(r=>r.json()).then(function(j){
        if (j.ok) { sessionStorage.setItem('hub_admin_pass', PASSCODE); unlock(); }
        else { document.getElementById('loginErr').textContent = 'Incorrect passcode.'; }
      });
  });
  document.getElementById('addForm').addEventListener('submit', function(e){
    e.preventDefault(); submitAdd();
  });
});

function unlock() {
  document.getElementById('loginBox').style.display='none';
  document.getElementById('adminBody').style.display='block';
  loadTools(); loadUsage();
}

function loadTools() {
  fetch('/api/tools').then(r=>r.json()).then(function(j){
    TOOLS = j.tools;
    document.getElementById('toolsTbl').innerHTML = TOOLS.map(rowHtml).join('');
  });
}

function rowHtml(t) {
  return '<tr>' +
    '<td>' + t.icon + ' ' + t.tool_name + '</td>' +
    '<td>' + t.category + '</td>' +
    '<td>' + t.owner + '</td>' +
    '<td>v' + t.version + '</td>' +
    '<td>' + t.last_updated + '</td>' +
    '<td>' + t.usage_count + '</td>' +
    '<td><span class="status-pill ' + (t.status==='active'?'active':'disabled') + '">' + t.status + '</span></td>' +
    '<td class="row-actions">' +
      '<button onclick="editTool(\'' + t.id + '\')">Edit</button>' +
      '<button onclick="showHistory(\'' + t.id + '\')">History</button>' +
      '<button onclick="toggleTool(\'' + t.id + '\')">' + (t.status==='active'?'Disable':'Enable') + '</button>' +
      '<button class="danger" onclick="deleteTool(\'' + t.id + '\')">Delete</button>' +
    '</td></tr>';
}

function toggleTool(id) {
  fetch('/api/admin/tools/' + id + '/toggle', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({passcode: PASSCODE})}).then(r=>r.json()).then(function(){ loadTools(); });
}
function deleteTool(id) {
  if (!confirm('Delete this tool from the hub? This cannot be undone.')) return;
  fetch('/api/admin/tools/' + id + '?passcode=' + encodeURIComponent(PASSCODE), {method:'DELETE'})
    .then(r=>r.json()).then(function(){ loadTools(); });
}
function editTool(id) {
  var t = TOOLS.find(x=>x.id===id);
  if (!t) return;
  fillForm(t);
  document.getElementById('formTitle').textContent = 'Edit Tool: ' + t.tool_name;
  document.getElementById('editingId').value = t.id;
  window.scrollTo({top:0, behavior:'smooth'});
}
function resetForm() {
  document.getElementById('addForm').reset();
  document.getElementById('editingId').value = '';
  document.getElementById('formTitle').textContent = 'Add New Tool';
}
function fillForm(t) {
  var f = document.getElementById('addForm');
  f.tool_name.value = t.tool_name; f.category.value = t.category; f.icon.value = t.icon;
  f.description.value = t.description; f.owner.value = t.owner; f.version.value = t.version;
  f.tool_url.value = t.tool_url; f.guide_url.value = t.guide_url; f.db_path.value = t.db_path;
  f.db_viewer_url.value = t.db_viewer_url || '';
  f.hours_saved_per_use.value = t.hours_saved_per_use; f.purpose.value = t.purpose;
  f.scope.value = t.scope; f.contact.value = t.contact;
  f.features.value = (t.features||[]).join('\n');
  f.benefits.value = (t.benefits||[]).join('\n');
  f.change_notes.value = ''; f.changed_by.value = '';
}

function submitAdd() {
  var f = document.getElementById('addForm');
  var id = document.getElementById('editingId').value;
  var payload = {
    passcode: PASSCODE,
    tool_name: f.tool_name.value, category: f.category.value, icon: f.icon.value || '\ud83d\udd27',
    description: f.description.value, owner: f.owner.value, version: f.version.value,
    tool_url: f.tool_url.value, guide_url: f.guide_url.value, db_path: f.db_path.value,
    db_viewer_url: f.db_viewer_url.value,
    hours_saved_per_use: parseFloat(f.hours_saved_per_use.value || 0),
    purpose: f.purpose.value, scope: f.scope.value, contact: f.contact.value,
    features: f.features.value.split('\n').map(s=>s.trim()).filter(Boolean),
    benefits: f.benefits.value.split('\n').map(s=>s.trim()).filter(Boolean),
    db_type: 'sqlite', status: 'active',
    change_notes: f.change_notes.value, changed_by: f.changed_by.value
  };
  var url = id ? ('/api/admin/tools/' + id) : '/api/admin/tools';
  var method = id ? 'PUT' : 'POST';
  fetch(url, {method:method, headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)})
    .then(r=>r.json()).then(function(j){
      if (!j.ok) { alert(j.error || 'Save failed'); return; }
      resetForm(); loadTools();
    });
}

function loadUsage() {
  fetch('/api/admin/usage?passcode=' + encodeURIComponent(PASSCODE)).then(r=>r.json()).then(function(j){
    if (!j.ok) return;
    document.getElementById('usageTbl').innerHTML = j.usage.map(function(u){
      return '<tr><td>' + u.tool_name + '</td><td>' + u.open + '</td><td>' + u.guide +
        '</td><td>' + u.about + '</td><td>' + u.dashboard + '</td></tr>';
    }).join('');
  });
}

var HISTORY_TOOL_ID = null;
function showHistory(id) {
  HISTORY_TOOL_ID = id;
  var t = TOOLS.find(x=>x.id===id);
  document.getElementById('historyTitle').textContent = 'Version History — ' + (t ? t.tool_name : id);
  document.getElementById('historyBody').innerHTML = '<div class="empty-state">Loading…</div>';
  document.getElementById('historyOverlay').classList.add('open');
  fetch('/api/admin/tools/' + id + '/history?passcode=' + encodeURIComponent(PASSCODE))
    .then(r=>r.json()).then(function(j){
      if (!j.ok) { document.getElementById('historyBody').innerHTML = '<div class="empty-state">' + (j.error||'Could not load history.') + '</div>'; return; }
      var chHtml = '<h4 style="color:var(--accent2);font-size:12px;text-transform:uppercase;margin:14px 0 6px;">Change Log (shown in About popup)</h4>' +
        (j.change_history.length ? j.change_history.map(c =>
          '<div class="ch-row"><div class="v">v'+c.version+'</div><div class="d">'+c.date+'</div><div>'+c.notes+'</div></div>'
        ).join('') : '<div class="empty-state">No change log entries yet.</div>');
      var auditHtml = '<h4 style="color:var(--accent2);font-size:12px;text-transform:uppercase;margin:18px 0 6px;">Full Audit Trail</h4>' +
        (j.audit.length ? j.audit.map(auditRowHtml).join('') : '<div class="empty-state">No audit history yet.</div>');
      document.getElementById('historyBody').innerHTML = chHtml + auditHtml;
    });
}
function auditRowHtml(a) {
  var verLabel = a.from_version && a.to_version && a.from_version !== a.to_version
    ? ('v' + a.from_version + ' \u2192 v' + a.to_version)
    : (a.to_version ? 'v' + a.to_version : '');
  var fields = (a.field_changes || []).map(function(fc){
    return '<div style="font-size:11px;color:var(--muted);padding-left:8px;">&bull; ' + fc.field + ': "' +
      (fc.old==null?'—':fc.old) + '" \u2192 "' + (fc.new==null?'—':fc.new) + '"</div>';
  }).join('');
  var rollbackBtn = a.has_snapshot ?
    '<button style="margin-left:8px;" onclick="rollbackTo(' + a.id + ')">Roll back to this point</button>' : '';
  return '<div class="audit-row">' +
    '<div style="display:flex;justify-content:space-between;align-items:center;">' +
      '<div><b>' + a.change_type.replace('_',' ') + '</b> ' + verLabel + ' <span style="color:var(--muted)">by ' + (a.changed_by||'—') + '</span></div>' +
      '<div style="color:var(--muted);font-family:var(--mono);font-size:11px;">' + a.ts.replace('T',' ') + rollbackBtn + '</div>' +
    '</div>' +
    (a.notes ? '<div style="font-size:12px;margin-top:4px;">' + a.notes + '</div>' : '') +
    fields +
  '</div>';
}
function rollbackTo(auditId) {
  if (!confirm('Roll this tool back to that earlier version? This will restore its fields and add a new "rollback" entry to the version history.')) return;
  var by = prompt('Your name (for the audit trail):', '') || 'admin';
  fetch('/api/admin/tools/' + HISTORY_TOOL_ID + '/rollback', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({passcode: PASSCODE, audit_id: auditId, changed_by: by})})
    .then(r=>r.json()).then(function(j){
      if (!j.ok) { alert(j.error || 'Rollback failed'); return; }
      loadTools();
      showHistory(HISTORY_TOOL_ID);
    });
}
"""

PAGE_ADMIN = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ENGAGE — Admin</title>
<style>""" + BASE_CSS + """
.form-grid { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.form-grid .full { grid-column:1 / -1; }
.form-grid label { display:block; font-size:11.5px; color:var(--muted); margin-bottom:4px; }
.form-grid input, .form-grid textarea, .form-grid select {
  width:100%; background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:6px; padding:8px 10px; font-size:13px;
}
.form-grid textarea { resize:vertical; min-height:60px; }
table.admin-tbl { width:100%; border-collapse:collapse; font-size:12.5px; margin-top:10px; }
table.admin-tbl th { text-align:left; padding:8px 10px; background:var(--surface3); }
table.admin-tbl td { padding:8px 10px; border-bottom:1px solid var(--border); }
.row-actions button { background:var(--surface2); border:1px solid var(--border); color:var(--text);
  border-radius:5px; padding:5px 9px; font-size:11.5px; cursor:pointer; margin-right:5px; }
.row-actions button.danger { border-color:var(--red); color:var(--red); }
.audit-row { border:1px solid var(--border); border-radius:8px; padding:10px 12px; margin-bottom:8px; background:var(--surface2); }
.history-modal { background:var(--surface); border:1px solid var(--border); border-radius:12px;
  width:100%; max-width:700px; max-height:85vh; overflow-y:auto; padding:22px 24px; }
#loginBox { max-width:340px; margin:80px auto; text-align:center; }
#loginBox input { width:100%; padding:10px; border-radius:6px; border:1px solid var(--border);
  background:var(--surface2); color:var(--text); margin:10px 0; }
#loginBox button { background:var(--accent); border:none; color:#fff; padding:10px 18px;
  border-radius:6px; cursor:pointer; font-weight:600; }
.btn-row { display:flex; gap:10px; margin-top:14px; }
.btn-row button { background:var(--accent); border:none; color:#fff; padding:10px 18px;
  border-radius:6px; cursor:pointer; font-weight:600; }
.btn-row button.ghost { background:var(--surface2); color:var(--text); border:1px solid var(--border); }
</style>
</head>
<body>
<header>
  <div class="header-logo">
    {% if rr_logo %}<img class="logo-rr-img" src="{{ rr_logo }}" alt="Rolls-Royce">{% else %}<div class="logo-rr">RR</div>{% endif %}
    {% if alten_logo %}<img class="logo-alten-img" src="{{ alten_logo }}" alt="ALTEN">{% else %}<div class="logo-alten">ALTEN</div>{% endif %}
  </div>
  <div class="header-title"><h1>ENGAGE <span class="brand-tagline">Launch. Automate. Accelerate.</span></h1><p>Admin — Manage tools, URLs, guides, and view usage metrics</p></div>
  <div class="header-actions"><a class="icon-btn" href="/">\u2190 Back to Hub</a></div>
</header>
<main>
  <div id="loginBox" class="panel" style="display:none">
    <h3>Admin Access</h3>
    <input type="password" id="passInput" placeholder="Enter admin passcode">
    <div id="loginErr" style="color:var(--red);font-size:12px;"></div>
    <button id="loginBtn">Unlock</button>
  </div>

  <div id="adminBody" style="display:none">
    <div class="section-title" id="formTitle">Add New Tool</div>
    <div class="panel">
      <form id="addForm">
        <input type="hidden" id="editingId">
        <div class="form-grid">
          <div><label>Tool Name</label><input name="tool_name" required></div>
          <div><label>Category</label>
            <select name="category">
              <option>Automation</option><option>Quality</option><option>Engineering</option>
              <option>AI</option><option>Analytics</option><option>Other</option>
            </select></div>
          <div><label>Icon (emoji)</label><input name="icon" placeholder="\U0001F527"></div>
          <div><label>Owner</label><input name="owner"></div>
          <div><label>Version</label><input name="version" placeholder="1.0"></div>
          <div><label>Estimated Hours Saved / Use</label><input name="hours_saved_per_use" type="number" step="0.1"></div>
          <div><label>Tool URL (http(s):// link, OR a local/UNC path to a .bat/.exe — opens with OS default handler)</label><input name="tool_url" placeholder="http://127.0.0.1:5010  or  \\\\server\\share\\run_tool.bat"></div>
          <div><label>User Guide (http(s):// link, OR a local/UNC path to .pdf/.docx/.html)</label><input name="guide_url" placeholder="http://.../guide.pdf  or  \\\\server\\share\\guide.pdf"></div>
          <div class="full"><label>SQLite DB path for built-in Data Dashboard (relative to hub folder, optional)</label>
            <input name="db_path" placeholder="e.g. bdc_usage_log.db"></div>
          <div class="full"><label>External DB Viewer URL (optional — if set, icon 3 opens this instead of the built-in dashboard)</label>
            <input name="db_viewer_url" placeholder="e.g. http://127.0.0.1:5002"></div>
          <div class="full"><label>Short Description (shown on card)</label><input name="description"></div>
          <div class="full"><label>Purpose</label><textarea name="purpose"></textarea></div>
          <div class="full"><label>Scope</label><textarea name="scope"></textarea></div>
          <div><label>Features (one per line)</label><textarea name="features"></textarea></div>
          <div><label>Benefits (one per line)</label><textarea name="benefits"></textarea></div>
          <div class="full"><label>Contact Email</label><input name="contact"></div>
          <div><label>Your Name (for version history)</label><input name="changed_by" placeholder="e.g. Raja"></div>
          <div class="full"><label>Change Notes (what changed / why — recorded in version history)</label>
            <textarea name="change_notes" placeholder="e.g. Bumped to v1.1, fixed export bug"></textarea></div>
        </div>
        <div class="btn-row">
          <button type="submit">Save Tool</button>
          <button type="button" class="ghost" onclick="resetForm()">Clear / New</button>
        </div>
      </form>
    </div>

    <div class="section-title">All Tools</div>
    <div class="panel">
      <table class="admin-tbl">
        <thead><tr><th>Tool</th><th>Category</th><th>Owner</th><th>Version</th><th>Updated</th><th>Uses</th><th>Status</th><th>Actions</th></tr></thead>
        <tbody id="toolsTbl"></tbody>
      </table>
    </div>

    <div class="section-title">Usage Metrics by Action</div>
    <div class="panel">
      <table class="admin-tbl">
        <thead><tr><th>Tool</th><th>Opens</th><th>Guide Views</th><th>About Views</th><th>Dashboard Views</th></tr></thead>
        <tbody id="usageTbl"></tbody>
      </table>
    </div>
  </div>
</main>
<footer>&copy; 2026 Alten-Rolls-Royce. All rights reserved. Confidential &ndash; Internal Use Only.</footer>

<div class="modal-overlay" id="historyOverlay">
  <div class="history-modal">
    <button class="close-x" onclick="closeModal('historyOverlay')">&times;</button>
    <h2 id="historyTitle">Version History</h2>
    <div class="sub">Every version bump and field-level change is recorded here, with rollback support.</div>
    <div id="historyBody"></div>
  </div>
</div>

<script>""" + ADMIN_JS + """</script>
</body>
</html>"""
