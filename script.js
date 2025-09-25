/* =========================
   HFBC Praise & Worship — FULL script.js
   - Top 10 only (excludes NA / N.A. / N/A / None / contains "Church Close")
   - 3D pie under Song Analytics
   - Coming Up & Previous Set side-by-side
========================= */

/* ---------- CONFIG ---------- */
const GITHUB = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };
const PATHS = {
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  setlist: "setlist/setlist.xlsx",
  bible: "bible_study/bible_study.xlsx",
  special: "special_practice/special_practice.xlsx"
};

const PRACTICE = {
  thursday: { dow: 4, time: "6:00pm–8:00pm" },
  sunday:   { dow: 0, time: "8:40am–9:30am" }
};

/* ---------- UTIL ---------- */
const $ = (sel, root=document) => root.querySelector(sel);

// Dates include weekday: "Sunday, Sep 28, 2025"
const fmtDate = d =>
  d ? d.toLocaleDateString(undefined, {
        weekday: "long", month: "short", day: "numeric", year: "numeric"
      })
    : "";

const escapeHtml = (s="") => String(s).replace(/[&<>"']/g, m => ({
  "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"
}[m]));

function todayLocalMidnight(){
  const t = new Date();
  return new Date(t.getFullYear(), t.getMonth(), t.getDate());
}
function rawUrl(pathRel){
  return `https://raw.githubusercontent.com/${GITHUB.owner}/${GITHUB.repo}/${GITHUB.branch}/${pathRel}`;
}
async function fetchWB(pathRel){
  const res = await fetch(rawUrl(pathRel) + `?nocache=${Date.now()}`);
  if(!res.ok) throw new Error(`Fetch failed: ${pathRel} (${res.status})`);
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array", cellDates: true });
}
function aoaFromWB(wb){
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
}

// Robust Excel/JS/string date → local midnight
function toLocalDate(val){
  if(val == null || val === "") return null;

  if (val instanceof Date && !isNaN(val)) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  if (typeof val === "number") {
    const o = XLSX.SSF.parse_date_code(val);
    if (o && o.y && o.m && o.d) return new Date(o.y, o.m - 1, o.d);
  }

  const s = String(val).trim();
  const mdyyyy = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/;
  const ymd    = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/;

  let m;
  if ((m = s.match(mdyyyy))) {
    const M = +m[1], D = +m[2], Y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return new Date(Y, M - 1, D);
  }
  if ((m = s.match(ymd))) {
    return new Date(+m[1], +m[2] - 1, +m[3]);
  }

  const d = new Date(s);
  if(!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}
function safeLabel(s){ return String(s || "").replace(/\s+/g," ").trim(); }

function findFirst(haystack, needles){
  const L = haystack.map(h => String(h ?? "").toLowerCase().trim());
  for(const n of needles){
    const i = L.indexOf(n.toLowerCase());
    if (i !== -1) return i;
  }
  return -1;
}

/* ---------- EXCLUSION RULES (analytics) ---------- */
function isExcludedSong(song){
  const s = String(song || "").trim().toLowerCase();
  if (!s) return true;
  if (s === "na" || s === "n/a" || s === "n.a." || s === "n.a" || s === "none") return true;
  if (s.includes("church close")) return true;
  return false;
}

/* ---------- LAYOUT HELPERS ---------- */
// Build a 2-column grid for Coming Up / Previous setlists inline
function ensureSetlistsGrid(){
  // Find existing single-column cards (best-effort)
  const nextCard = $("#setlist-next")?.closest(".card") || $("#setlist-next")?.parentElement;
  const prevCard = $("#setlist-prev")?.closest(".card") || $("#setlist-prev")?.parentElement;
  const parent = (nextCard && nextCard.parentElement) || document.body;

  // Hide old single-column blocks if they exist
  if (nextCard) nextCard.style.display = "none";
  if (prevCard) prevCard.style.display = "none";

  // Create grid
  const grid = document.createElement("div");
  grid.id = "setlists-grid";
  grid.style.cssText = "display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:16px;align-items:start;margin-top:8px;";

  // Two empty cards for Next & Previous
  grid.innerHTML = `
    <div class="card">
      <h2 class="section-title">Coming Up Set</h2>
      <div id="set-next-meta" class="dim" style="margin-bottom:6px;">—</div>
      <div id="set-next-table" class="table like-card">Loading…</div>
    </div>
    <div class="card">
      <h2 class="section-title">Previous Set</h2>
      <div id="set-prev-meta" class="dim" style="margin-bottom:6px;">—</div>
      <div id="set-prev-table" class="table like-card">Loading…</div>
    </div>
  `;

  parent.appendChild(grid);
}

// Put 3D pie **under** Song Analytics (below the Top 10 list)
function ensurePieUnderAnalytics(){
  // Remove any old charts row/canvas pies
  document.querySelectorAll(".charts-2, #pieChart, canvas#pieChart").forEach(n => n.remove());

  // Remove previous Google pie container if we’re re-rendering
  const old = document.getElementById("pieChart3D");
  if (old) old.closest(".chart-card").parentElement.remove();

  const listEl = document.getElementById("top5");
  if (!listEl) return;
  const host = listEl.closest("div") || listEl.parentElement;

  const wrap = document.createElement("div");
  wrap.innerHTML = `
    <h3 class="subhead" style="margin-top:12px;">Plays by Song (3D Pie)</h3>
    <div class="chart-card"><div id="pieChart3D" style="width:100%;height:320px;"></div></div>
  `;
  host.appendChild(wrap);
}

/* ---------- Google Charts loader (3D pie) ---------- */
let gchartsLoaded = false, gchartsLoading = null;
function loadGoogleCharts(){
  if (gchartsLoaded) return Promise.resolve();
  if (gchartsLoading) return gchartsLoading;
  gchartsLoading = new Promise((resolve, reject)=>{
    const s = document.createElement('script');
    s.src = "https://www.gstatic.com/charts/loader.js";
    s.onload = () => {
      try{
        google.charts.load('current', { packages:['corechart'] });
        google.charts.setOnLoadCallback(() => { gchartsLoaded = true; resolve(); });
      }catch(e){ reject(e); }
    };
    s.onerror = () => reject(new Error("Failed to load Google Charts"));
    document.head.appendChild(s);
  });
  return gchartsLoading;
}

/* ---------- WORSHIP PRACTICE (formatted line) ---------- */
function nextOccurrence(targetDow){
  const today = todayLocalMidnight();
  const wd = today.getDay();
  let delta = (targetDow - wd + 7) % 7;
  if(delta === 0) delta = 7;
  const d = new Date(today);
  d.setDate(today.getDate() + delta);
  return d;
}
function weekdayName(i){ return ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][i]; }

/* REPLACED: show lines like "Thursday 09/25/2025 6:00pm–8:00pm" */
function loadPractice(){
  function mmddyyyy(d){
    const mm = String(d.getMonth()+1).padStart(2,"0");
    const dd = String(d.getDate()).padStart(2,"0");
    const yyyy = d.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }
  const thDate = nextOccurrence(PRACTICE.thursday.dow);
  const suDate = nextOccurrence(PRACTICE.sunday.dow);

  const thLine = `${weekdayName(thDate.getDay())} ${mmddyyyy(thDate)} ${PRACTICE.thursday.time}`;
  const suLine = `${weekdayName(suDate.getDay())} ${mmddyyyy(suDate)} ${PRACTICE.sunday.time}`;

  const html = `<ul class="reminders">
    <li><strong>Thursday Practice:</strong> ${thLine}</li>
    <li><strong>Sunday Practice:</strong> ${suLine}</li>
  </ul>`;
  $("#reminders-table").innerHTML = html;
}

/* ---------- SPECIAL PRACTICE (future-only) ---------- */
async function loadSpecialPractice(){
  try{
    const wb = await fetchWB(PATHS.special);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){
      $("#special-practice-table").innerHTML = `<p class="dim">No upcoming special practices.</p>`;
      return;
    }
    const hdrRaw = aoa[0].map(h => String(h).trim());
    const hdr = hdrRaw.map(h => h.toLowerCase());
    const idxDate = findFirst(hdr, ["date","service date"]);
    const idxTime = findFirst(hdr, ["time","start time"]);
    const idxNotes = findFirst(hdr, ["notes","note","description"]);

    const rows = [];
    for(let i=1;i<aoa.length;i++){
      const r = aoa[i]; if(!r || r.every(c => String(c ?? "").trim()==="")) continue;
      const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
      if(!d || d < todayLocalMidnight()) continue; // future only
      rows.push({
        date: d,
        time: idxTime !== -1 ? String(r[idxTime]) : "",
        notes: idxNotes !== -1 ? String(r[idxNotes]) : ""
      });
    }
    rows.sort((a,b)=> a.date - b.date);

    if(rows.length === 0){
      $("#special-practice-table").innerHTML = `<p class="dim">No upcoming special practices.</p>`;
      return;
    }

    const out = [["Date","Time","Notes"]];
    rows.forEach(r => out.push([fmtDate(r.date), r.time, r.notes]));

    renderAOATable(out, "#special-practice-table");
  }catch(e){
    console.error(e);
    $("#special-practice-table").innerHTML = `<p class="dim">Unable to load special practices.</p>`;
  }
}

/* ---------- ANNOUNCEMENTS (bilingual if two columns) ---------- */
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length === 0){ $("#announcements-table").innerHTML = `<p class="dim">No data.</p>`; return; }

    const hdrRaw = aoa[0].map(h => String(h).trim());
    const hdr = hdrRaw.map(h => h.toLowerCase());
    const idxDate = hdr.findIndex(h => ["date","service date"].includes(h));
    const idxEn = findFirst(hdr, ["english","announcement en","announcement (en)","announcement english","en","message en"]);
    const idxHm = findFirst(hdr, ["hmong","announcement hm","announcement (hmong)","announcement hmong","hm","message hm"]);

    if(idxEn !== -1 && idxHm !== -1){
      const out = [];
      const head = [];
      if(idxDate !== -1) head.push("Date");
      head.push("English","Hmong");
      out.push(head);

      for(let i=1;i<aoa.length;i++){
        const r = aoa[i]; if(!r || r.every(c => String(c ?? "").trim()==="")) continue;
        const row = [];
        if(idxDate !== -1){
          const d = toLocalDate(r[idxDate]);
          row.push(d ? fmtDate(d) : String(r[idxDate] ?? ""));
        }
        row.push(String(r[idxEn] ?? ""));
        row.push(String(r[idxHm] ?? ""));
        out.push(row);
      }
      renderAOATable(out, "#announcements-table");
      return;
    }

    // Single-language fallback
    const out = [["Message"]];
    for(let i=1;i<aoa.length;i++){
      const r = aoa[i]; if(!r || r.every(c => String(c ?? "").trim()==="")) continue;
      out.push([String(r[0] ?? "")]);
    }
    renderAOATable(out, "#announcements-table");
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<p class="dim">Unable to load announcements.</p>`;
  }
}

/* ---------- BIBLE VERSES (last 4 weeks) ---------- */
async function loadBibleVerses(){
  try{
    const wb = await fetchWB(PATHS.bible);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){ $("#bible-verse-table").innerHTML = `<p class="dim">No data.</p>`; return; }

    const hdrRaw = aoa[0].map(h => String(h).trim());
    const hdr = hdrRaw.map(h => h.toLowerCase());
    const idxDate = findFirst(hdr, ["date","service date"]);
    const idxRef  = findFirst(hdr, ["reference","ref","verse ref","bible verse ref","bible ref"]);
    const idxText = findFirst(hdr, ["text","verse","bible verse"]);

    const rows = [];
    for(let i=1;i<aoa.length;i++){
      const r = aoa[i]; if(!r || r.every(c => String(c ?? "").trim()==="")) continue;
      const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
      if(!d) continue;
      rows.push({ date: d, ref: idxRef !== -1 ? String(r[idxRef] ?? "") : "", text: idxText !== -1 ? String(r[idxText] ?? "") : "" });
    }

    // last 4 weeks (prefer future->past up to 4 rows)
    const today = todayLocalMidnight();
    let recent = rows
      .filter(x => x.date <= today)
      .sort((a,b)=> b.date - a.date)
      .slice(0,4);

    if(recent.length === 0){
      recent = rows
        .filter(x => x.date <= today)
        .sort((a,b)=> b.date - a.date)
        .slice(0,4);
    }

    if(recent.length === 0){
      $("#bible-verse-table").innerHTML = `<p class="dim">No verses in the last 4 weeks.</p>`;
      return;
    }

    const showRef = recent.some(x => x.ref);
    const showText = recent.some(x => x.text);
    const head = ["Date"];
    if(showRef) head.push("Reference");
    if(showText) head.push("Verse");

    const out = [head];
    for(const r of recent){
      const row = [fmtDate(r.date)];
      if(showRef) row.push(r.ref || "—");
      if(showText) row.push(r.text || "—");
      out.push(row);
    }
    renderAOATable(out, "#bible-verse-table");
  }catch(e){
    console.error(e);
    $("#bible-verse-table").innerHTML = `<p class="dim">Unable to load verses.</p>`;
  }
}

/* ---------- MEMBERS ---------- */
async function loadMembers(){
  try{
    const wb = await fetchWB(PATHS.members);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){ $("#members-table").innerHTML = `<p class="dim">No data.</p>`; return; }

    const hdrRaw = aoa[0].map(h => String(h).trim());
    const hdr = hdrRaw.map(h => h.toLowerCase());
    const idxName  = findFirst(hdr, ["name","member name"]);
    const idxRole  = findFirst(hdr, ["role","part","voice","instrument"]);
    const idxAvail = findFirst(hdr, ["availability","available"]);

    const out = [["Name","Role","Availability"]];
    for(let i=1;i<aoa.length;i++){
      const r = aoa[i]; if(!r || r.every(c => String(c ?? "").trim()==="")) continue;
      out.push([
        idxName !== -1 ? String(r[idxName] ?? "") : "",
        idxRole !== -1 ? String(r[idxRole] ?? "") : "",
        idxAvail !== -1 ? String(r[idxAvail] ?? "") : ""
      ]);
    }
    renderAOATable(out, "#members-table");
  }catch(e){
    console.error(e);
    $("#members-table").innerHTML = `<p class="dim">Unable to load members.</p>`;
  }
}

/* ---------- SETLISTS + ANALYTICS ---------- */
function renderAOATable(aoa, selector){
  const html = `
    <table>
      ${aoa[0] ? `<thead><tr>${aoa[0].map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>` : ""}
      <tbody>
        ${aoa.slice(1).map(r=>`<tr>${r.map(c=>`<td>${escapeHtml(String(c))}</td>`).join("")}</tr>`).join("")}
      </tbody>
    </table>
  `;
  $(selector).innerHTML = html;
}
function renderSetlistCard(group, metaSel, tableSel){
  if(!group){ $(metaSel).textContent = "—"; $(tableSel).innerHTML = `<p class="dim">No data.</p>`; return; }
  $(metaSel).textContent =
    `${group.date ? "Service Date: " + fmtDate(group.date) + " · " : ""}` +
    `${group.sermon ? "Sermon: " + group.sermon : "Sermon: —"}`;
  const out = [["Song"]];
  group.rows.forEach(r => out.push([r.song]));
  renderAOATable(out, tableSel);
}
function pickGroup(rows, hdr){
  const idxService = findFirst(hdr, ["service date","date"]);
  const idxSong    = findFirst(hdr, ["song","title","song title"]);
  const idxSermon  = findFirst(hdr, ["sermon","topic","theme"]);

  const entries = [];
  for(let i=1;i<rows.length;i++){
    const r = rows[i]; if(!r || r.every(c => String(c ?? "").trim()==="")) continue;
    const d = idxService !== -1 ? toLocalDate(r[idxService]) : null;
    const s = idxSong    !== -1 ? String(r[idxSong] ?? "") : "";
    const sermon = idxSermon !== -1 ? String(r[idxSermon] ?? "") : "";
    if (!d || !s) continue;
    entries.push({ date: d, song: s, sermon });
  }
  if(entries.length === 0) return null;

  // Group by date
  const map = new Map();
  for(const e of entries){
    const key = e.date.toISOString();
    if(!map.has(key)) map.set(key, { date:e.date, sermon:e.sermon || "", rows:[] });
    map.get(key).rows.push({ song: e.song });
  }

  // Sort groups by date desc
  const groups = [...map.values()].sort((a,b)=> b.date - a.date);
  return { next: groups[0], prev: groups[1] || null, groups };
}

async function loadSetlistsAndAnalytics(){
  try{
    const wb = await fetchWB(PATHS.setlist);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){
      $("#set-next-table").innerHTML = `<p class="dim">No data.</p>`;
      $("#set-prev-table").innerHTML = `<p class="dim">No data.</p>`;
      return;
    }

    const hdrRaw = aoa[0].map(h => String(h).trim());
    const hdr = hdrRaw.map(h => h.toLowerCase());
    const idxSong = findFirst(hdr, ["song","title","song title"]);

    const { next, prev, groups } = pickGroup(aoa, hdr);
    renderSetlistCard(next, "#set-next-meta", "#set-next-table");
    renderSetlistCard(prev, "#set-prev-meta", "#set-prev-table");

    // --- Analytics (Top 10, Exclusions) ---
    const counts = new Map();
    if(idxSong !== -1){
      for(const r of aoa.slice(1)){
        const sRaw = String(r[idxSong] ?? "").trim();
        if (!sRaw || isExcludedSong(sRaw)) continue;
        counts.set(sRaw, (counts.get(sRaw)||0)+1);
      }
    }
    const songCounts = [...counts.entries()].sort((a,b)=>b[1]-a[1]);

    const leftHeader = $("#top5")?.previousElementSibling;
    if(leftHeader) leftHeader.textContent = "Top 10 – Most Played";

    const top10 = songCounts.slice(0,10);
    $("#top5").innerHTML = top10.length
      ? top10.map(([s,c])=>`<li>${escapeHtml(safeLabel(s))} — ${c}</li>`).join("")
      : `<li class="dim">No data</li>`;

    // Ensure the 3D pie lives **under** Song Analytics, then draw it
    ensurePieUnderAnalytics();
    await drawSongsPie3D(songCounts.slice(0,7)); // top 7 slices for readability
  }catch(e){
    console.error(e);
    $("#set-next-table").innerHTML = `<p class="dim">Unable to load <code>${PATHS.setlist}</code>.</p>`;
    $("#set-prev-table").innerHTML = `<p class="dim">—</p>`;
  }
}

function renderSetlistGroup(group, metaSel, tableSel){
  if(!group){ $(metaSel).textContent = "—"; $(tableSel).innerHTML = `<p class="dim">No data.</p>`; return; }
  $(metaSel).textContent =
    `${group.date ? "Service Date: " + fmtDate(group.date) + " · " : ""}` +
    `${group.sermon ? "Sermon: " + group.sermon : "Sermon: —"}`;

  const header = ["Song"];
  let html = `<table><thead><tr>${header.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>`;
  group.rows.forEach((row, i)=>{
    const id = `${tableSel.replace("#","")}-story-${i}`;
    html += `<tr>
      <td>${escapeHtml(row.song)}
        <span class="story-btn" style="margin-left:8px" data-song="${escapeHtml(row.song)}" data-target="${id}">Story</span>
        <div id="${id}" class="story-popup" style="display:none">Loading…</div>
      </td>
    </tr>`;
  });
  html += `</tbody></table>`;
  $(tableSel).innerHTML = html;
}

/* ---------- Chart: Google 3D Pie ---------- */
async function drawSongsPie3D(entries){
  const container = document.getElementById("pieChart3D");
  if (!container || !entries || entries.length === 0) return;

  await loadGoogleCharts();

  const dataArr = [["Song","Plays"]];
  for(const [label, count] of entries){
    dataArr.push([safeLabel(label), Number(count) || 0]);
  }
  const data = google.visualization.arrayToDataTable(dataArr);

  const options = {
    is3D: true,
    backgroundColor: 'transparent',
    pieSliceText: 'percentage',
    legend: { textStyle: { color: '#cbd5e1' } },
    titleTextStyle: { color: '#cbd5e1' },
    chartArea: { left: 10, top: 10, width: '95%', height: '85%' }
  };

  const chart = new google.visualization.PieChart(container);
  chart.draw(data, options);

  window.addEventListener('resize', () => chart.draw(data, options));
}

/* ===========================
   Data Analyst — additive block
   Reads weekly CSVs from /data and renders inside Song Analytics card.
   =========================== */
async function daFetchCSV(url) {
  const bust = url.includes("?") ? `&t=${Date.now()}` : `?t=${Date.now()}`;
  const res = await fetch(url + bust, { cache: "no-store" });
  if (!res.ok) { console.warn("CSV fetch failed:", url, res.status); return []; }
  const text = await res.text();
  if (!text.trim()) return [];
  const rows = text.trim().split("\n").map(r => r.split(","));
  const header = rows.shift();
  return rows.map(r => Object.fromEntries(header.map((h,i)=>[h.trim(), (r[i]??"").trim()])));
}
function daMountHost() {
  const analyticsCard = document.querySelector("#analytics-section .card");
  if (!analyticsCard) return null;
  const host = document.createElement("div");
  host.id = "da-grid";
  host.className = "charts charts-2";
  host.innerHTML = `
    <div class="chart-card">
      <h3 class="subhead">Hymnal Coverage</h3>
      <div id="da-coverage" class="table like-card"><p class="dim" style="padding:8px;margin:0">Loading…</p></div>
    </div>
    <div class="chart-card">
      <h3 class="subhead">Usage by Source</h3>
      <div id="da-by-source" class="table like-card"><p class="dim" style="padding:8px;margin:0">Loading…</p></div>
    </div>
    <div class="chart-card" style="grid-column:1/-1">
      <h3 class="subhead">Unused Hymnal Numbers</h3>
      <div id="da-unused" class="table like-card"><p class="dim" style="padding:8px;margin:0">Loading…</p></div>
    </div>`;
  analyticsCard.appendChild(host);
  // ensure tables don't paint white
  host.querySelectorAll("table").forEach(t => t.style.background = "transparent");
  return host;
}
function daRenderCoverage(rows){
  const obj = Object.fromEntries(rows.map(r => [r.metric, r.value]));
  const html = `
    <table style="background:transparent">
      <tbody>
        <tr><th>Used</th><td>${obj.hymnal_coverage_used ?? "—"} / 352</td></tr>
        <tr><th>Unused</th><td>${obj.hymnal_coverage_unused ?? "—"}</td></tr>
        <tr><th>Coverage</th><td>${obj.hymnal_coverage_percent ?? "—"}%</td></tr>
      </tbody>
    </table>`;
  document.getElementById("da-coverage").innerHTML = html;
}
function daRenderBySource(rows){
  const html = rows.length ? `
    <table style="background:transparent">
      <thead><tr><th>Source</th><th>Count</th></tr></thead>
      <tbody>${rows.map(r => `<tr><td>${r.source_final}</td><td>${r.count}</td></tr>`).join("")}</tbody>
    </table>` : `<p class="dim" style="padding:8px;margin:0">No data.</p>`;
  document.getElementById("da-by-source").innerHTML = html;
}
function daRenderUnused(rows){
  if (!rows.length){ document.getElementById("da-unused").innerHTML = `<p class="dim" style="padding:8px;margin:0">All hymnal numbers have been used at least once.</p>`; return; }
  const nums = rows.map(r => parseInt(r.unused_number,10)).filter(n => !Number.isNaN(n)).sort((a,b)=>a-b).slice(0,50);
  document.getElementById("da-unused").innerHTML = `
    <table style="background:transparent">
      <thead><tr><th>Unused (first 50)</th></tr></thead>
      <tbody><tr><td>${nums.map(n=>`#${n}`).join(", ")}</td></tr></tbody>
    </table>`;
}
async function loadDataAnalyst(){
  if (!document.getElementById("da-grid")) daMountHost();
  const [cov, bysrc, unused] = await Promise.all([
    daFetchCSV("data/kpi_hymnal_coverage.csv"),
    daFetchCSV("data/kpi_by_source.csv"),
    daFetchCSV("data/hymnal_unused.csv"),
  ]);
  daRenderCoverage(cov); daRenderBySource(bysrc); daRenderUnused(unused);
}

/* ---------- BOOT ---------- */
document.addEventListener("DOMContentLoaded", async ()=>{
  // Layout scaffolding
  ensureSetlistsGrid();
  ensurePieUnderAnalytics();

  // Hide "Least Popular" and Bar chart (keep Top-10 + 3D Pie only)
  const bottom = document.getElementById('bottom5'); if (bottom) bottom.closest('div').style.display='none';
  const bar = document.getElementById('barChart'); if (bar) bar.closest('.chart-card').style.display='none';

  // Practice & sections
  loadPractice();
  try{ await loadMembers(); }catch{}
  try{ await loadSpecialPractice(); }catch{}
  try{ await loadAnnouncements(); }catch{}
  try{ await loadBibleVerses(); }catch{}
  try{ await loadSetlistsAndAnalytics(); }catch{}
  try{ await loadDataAnalyst(); }catch{}
});
