// ---------- DOM helpers ----------
const $ = (id) => document.getElementById(id);

// Convert GitHub "blob" URL to raw.githubusercontent.com
function toRawGithub(url){
  try{
    const u = new URL(url);
    if (u.hostname === 'raw.githubusercontent.com') return url;
    if (u.hostname === 'github.com'){
      const parts = u.pathname.split('/');
      const i = parts.indexOf('blob');
      if (i !== -1){
        const owner = parts[1], repo = parts[2], branch = parts[i+1];
        const path = parts.slice(i+2).join('/');
        return `https://raw.githubusercontent.com/${owner}/${repo}/${branch}/${path}`;
      }
    }
  }catch(e){}
  return url;
}

function escapeHtml(s){
  return s.replace(/[&<>"']/g, ch=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[ch]));
}

function setKpi(id, text){ $(id).innerHTML = `<span class="pill">${escapeHtml(text)}</span>`; }

function renderTable(rows, mountId, limit=200){
  if (!rows || rows.length===0){
    $(mountId).innerHTML = `<div class="small muted">No rows.</div>`;
    return;
  }
  const headers = Object.keys(rows[0]);
  const head = `<tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
  const body = rows.slice(0,limit).map(r=>{
    return `<tr>${headers.map(h=>`<td>${escapeHtml(String(r[h]))}</td>`).join('')}</tr>`;
  }).join('');
  $(mountId).innerHTML = `<div class="small muted">Showing ${Math.min(rows.length,limit)} of ${rows.length} rows</div>
    <div style="overflow:auto;max-height:380px;border:1px solid rgba(255,255,255,.06);border-radius:10px">
      <table><thead>${head}</thead><tbody>${body}</tbody></table>
    </div>`;
}

// ---------- Excel loading ----------
async function fetchSheet(url){
  const raw = toRawGithub(url);
  const res = await fetch(raw, {mode:'cors'});
  if (!res.ok) throw new Error(`HTTP ${res.status} on ${raw}`);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), {type:'array'});
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
  return {rows, sheetName, count: rows.length};
}

// ---------- Analytics (Setlist) ----------
function pickColumn(headers, candidates){
  const lower = headers.map(h=>h.toLowerCase());
  for (const cand of candidates){
    const idx = lower.findIndex(h=> h===cand || h.includes(cand));
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function parseDateFlexible(val){
  if (!val) return null;
  if (typeof val === 'number'){
    const epoch = new Date(Date.UTC(1899,11,30)); // Excel serial date base
    const ms = val * 86400000;
    return new Date(epoch.getTime()+ms);
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function buildSongStats(setlistRows){
  if (!setlistRows || setlistRows.length===0) return null;
  const headers = Object.keys(setlistRows[0]);
  const songCol = pickColumn(headers, ['song title','song','title','songs']);
  const dateCol = pickColumn(headers, ['date','service date','when','week']);

  const counts = new Map();
  const timeline = new Map(); // yyyy-mm -> count
  let total = 0;

  for(const r of setlistRows){
    const name = (r[songCol] ?? '').toString().trim();
    if (name){
      counts.set(name, (counts.get(name)||0)+1);
      total++;
    }
    if (dateCol){
      const dt = parseDateFlexible(r[dateCol]);
      if (dt){
        const key = `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}`;
        timeline.set(key, (timeline.get(key)||0)+1);
      }
    }
  }
  const freq = Array.from(counts.entries()).sort((a,b)=>b[1]-a[1]);
  const top = freq[0] || ['—', 0];
  const least = freq.length ? freq.at(-1) : ['—', 0];
  const unique = counts.size;

  const tl = Array.from(timeline.entries()).sort((a,b)=> a[0].localeCompare(b[0]));
  return {freq, top, least, unique, total, timeline: tl};
}

// ---------- Charts ----------
let barChart, pieChart, lineChart;
function upsertChart(ctxId, type, data, options){
  const existing = {barChart, pieChart, lineChart}[ctxId];
  if (existing) existing.destroy();
  const ctx = document.getElementById(ctxId).getContext('2d');
  const chart = new Chart(ctx, { type, data, options });
  if (ctxId==='barChart') barChart = chart;
  if (ctxId==='pieChart') pieChart = chart;
  if (ctxId==='lineChart') lineChart = chart;
}

function drawCharts(stats){
  if (!stats){
    upsertChart('barChart','bar',{labels:[],datasets:[{label:'No data',data:[]}]},{});
    upsertChart('pieChart','pie',{labels:[],datasets:[{label:'No data',data:[]}]},{});
    upsertChart('lineChart','line',{labels:[],datasets:[{label:'No data',data:[]}]},{});
    return;
  }
  const topN = stats.freq.slice(0,10);
  const labels = topN.map(d=>d[0]);
  const values = topN.map(d=>d[1]);

  upsertChart('barChart','bar',{
    labels,
    datasets:[{label:'Play Count (Top 10)', data: values}]
  },{
    responsive:true,
    plugins:{legend:{display:false}},
    scales:{y:{beginAtZero:true}}
  });

  const pieN = stats.freq.slice(0,6);
  upsertChart('pieChart','pie',{
    labels: pieN.map(d=>d[0]),
    datasets:[{label:'Top Share', data: pieN.map(d=>d[1])}]
  },{responsive:true});

  const tlLabels = stats.timeline.map(d=>d[0]);
  const tlValues = stats.timeline.map(d=>d[1]);
  upsertChart('lineChart','line',{
    labels: tlLabels,
    datasets:[{label:'Total Plays / Month', data: tlValues, tension: .25}]
  },{
    responsive:true,
    plugins:{legend:{display:true}},
    scales:{y:{beginAtZero:true}}
  });
}

// ---------- Main load ----------
async function loadAll(){
  const btn = $('loadBtn');
  const status = $('status');
  btn.disabled = true;
  status.innerHTML = `<span class="loader"></span> Loading Excel files from GitHub…`;

  const urls = {
    announcements: $('annUrl').value.trim(),
    bible: $('bibleUrl').value.trim(),
    members: $('membersUrl').value.trim(),
    setlist: $('setlistUrl').value.trim()
  };

  try{
    const [ann, bible, members, setlist] = await Promise.all([
      fetchSheet(urls.announcements),
      fetchSheet(urls.bible),
      fetchSheet(urls.members),
      fetchSheet(urls.setlist)
    ]);

    // Tables + KPIs
    setKpi('annKpi', `${ann.count} rows • sheet “${ann.sheetName}”`);
    renderTable(ann.rows, 'annTable');

    setKpi('bibleKpi', `${bible.count} rows • sheet “${bible.sheetName}”`);
    renderTable(bible.rows, 'bibleTable');

    setKpi('membersKpi', `${members.count} rows • sheet “${members.sheetName}”`);
    renderTable(members.rows, 'membersTable');

    setKpi('setlistKpi', `${setlist.count} rows • sheet “${setlist.sheetName}”`);
    renderTable(setlist.rows, 'setlistTable');

    // Analytics
    const stats = buildSongStats(setlist.rows);
    $('topSong').innerHTML = `Top song: <strong>${stats?.top?.[0] ?? '—'}</strong> (${stats?.top?.[1] ?? 0})`;
    $('leastSong').innerHTML = `Least played: <strong>${stats?.least?.[0] ?? '—'}</strong> (${stats?.least?.[1] ?? 0})`;
    $('totalPlays').innerHTML = `Total plays: <strong>${stats?.total ?? 0}</strong>`;
    $('uniqueSongs').innerHTML = `Unique songs: <strong>${stats?.unique ?? 0}</strong>`;
    drawCharts(stats);

    status.innerHTML = `<span class="ok">Loaded successfully.</span>`;
  }catch(err){
    console.error(err);
    status.innerHTML = `<span class="err">Load failed:</span> <span class="small">${escapeHtml(err.message)}</span><br>
    <span class="small">Tip: Keep the repo public. This page auto-converts “blob” links to raw URLs.</span>`;
  }finally{
    btn.disabled = false;
  }
}

// Wire up
window.addEventListener('DOMContentLoaded', () => {
  $('loadBtn').addEventListener('click', loadAll);
  loadAll();
});
