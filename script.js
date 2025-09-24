// ---------- CONFIG ----------
const GITHUB = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };

const PATHS = {
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  reminders: "reminders/reminders.xlsx", // optional
  setlist: "setlist/setlist.xlsx"
};

// ---------- UTIL ----------
const $ = (sel, root=document) => root.querySelector(sel);
const fmtDate = d => d ? d.toLocaleDateString(undefined, { month:"short", day:"numeric", year:"numeric" }) : "—";
const escapeHtml = s => s.replace(/[&<>"']/g, m => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[m]));

function rawUrl(pathRel){
  return `https://raw.githubusercontent.com/${GITHUB.owner}/${GITHUB.repo}/${GITHUB.branch}/${pathRel}`;
}

async function fetchWB(pathRel){
  const url = rawUrl(pathRel) + `?nocache=${Date.now()}`;
  const res = await fetch(url);
  if(!res.ok) throw new Error(`Fetch failed: ${url} (${res.status})`);
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array", cellDates: true }); // let xlsx give us JS Dates when available
}

function aoaFromWB(wb){
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
}

// Robust Excel/JS/string date → local Date at midnight (prevents TZ shifts)
function toLocalDate(val){
  if(val == null || val === "") return null;

  // 1) Already a JS Date
  if (val instanceof Date && !isNaN(val)) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }

  // 2) Excel serial number
  if (typeof val === "number") {
    const o = XLSX.SSF.parse_date_code(val);
    if (o && o.y && o.m && o.d) return new Date(o.y, o.m - 1, o.d);
  }

  // 3) Common string formats (m/d/yyyy, yyyy-mm-dd, etc.)
  const s = String(val).trim();
  const mdyyyy = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/; // 10/4/2025 or 10-4-2025
  const ymd = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/;       // 2025-10-04

  let m;
  if ((m = s.match(mdyyyy))) {
    const M = +m[1], D = +m[2], Y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return new Date(Y, M - 1, D);
  }
  if ((m = s.match(ymd))) {
    return new Date(+m[1], +m[2] - 1, +m[3]);
  }

  // 4) Fallback to Date parser, then normalize to midnight local
  const d = new Date(s);
  if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());

  return null;
}

function renderAOATable(aoa, targetSel){
  const el = $(targetSel);
  if(!aoa || aoa.length === 0){ el.innerHTML = `<p class="dim">No data.</p>`; return; }
  const header = aoa[0];
  let html = `<table><thead><tr>${header.map(h=>`<th>${escapeHtml(String(h))}</th>`).join("")}</tr></thead><tbody>`;
  for(let i=1;i<aoa.length;i++){
    const row = aoa[i];
    if(row.every(c=>String(c).trim()==="")) continue;
    html += `<tr>${header.map((_,j)=>`<td>${escapeHtml(String(row[j] ?? ""))}</td>`).join("")}</tr>`;
  }
  html += `</tbody></table>`;
  el.innerHTML = html;
}

function sameYMD(a,b){ return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }

// ---------- LOADERS (Announcements / Reminders / Members) ----------
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = aoaFromWB(wb);
    // Format any "Date" column
    const hdr = aoa[0].map(x=>String(x).toLowerCase());
    const idxDate = hdr.findIndex(h=>["date","service date"].includes(h));
    const out = idxDate === -1 ? aoa : aoa.map((row,i)=>{
      if(i===0) return row;
      const r = row.slice();
      const d = toLocalDate(r[idxDate]);
      r[idxDate] = d ? fmtDate(d) : (r[idxDate] ?? "");
      return r;
    });
    renderAOATable(out, "#announcements-table");
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<p class="dim">Unable to load <code>${PATHS.announcements}</code>.</p>`;
  }
}

async function loadReminders(){
  try{
    const wb = await fetchWB(PATHS.reminders);
    const aoa = aoaFromWB(wb);
    renderAOATable(aoa, "#reminders-table");
  }catch(e){
    // keep static two reminders silently
  }
}

async function loadMembers(){
  try{
    const wb = await fetchWB(PATHS.members);
    const aoa = aoaFromWB(wb);
    renderAOATable(aoa, "#members-table");
  }catch(e){
    console.error(e);
    $("#members-table").innerHTML = `<p class="dim">Unable to load <code>${PATHS.members}</code>.</p>`;
  }
}

// ---------- SETLIST (Two most recent dates; Sermon; Story toggle; NO "Key") ----------
async function loadSetlistsAndAnalytics(){
  try{
    const wb = await fetchWB(PATHS.setlist);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){ $("#setlist-this").innerHTML=`<p class="dim">No rows.</p>`; return; }

    const hdrRaw = aoa[0].map(h=>String(h));
    const hdr = hdrRaw.map(h=>h.trim().toLowerCase());
    const idxDate   = hdr.findIndex(h=>["date","service date"].includes(h));
    const idxSermon = hdr.findIndex(h=>["sermon","sermon topic","topic"].includes(h));
    const idxSong   = hdr.findIndex(h=>["song","title","song title"].includes(h));
    const idxNotes  = hdr.findIndex(h=>["notes","note","comment"].includes(h));

    const rows = aoa.slice(1).filter(r => r.some(c => String(c).trim()!==""));

    // Group rows by local date
    const groups = new Map(); // key -> { date, sermon, rows:[] }
    for(const r of rows){
      const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
      const key = d ? d.toISOString().slice(0,10) : "__nodate__";
      const sermon = idxSermon !== -1 ? String(r[idxSermon] ?? "").trim() : "";
      if(!groups.has(key)) groups.set(key, { date:d, sermon: sermon || "", rows:[] });
      const g = groups.get(key);
      if(sermon && !g.sermon) g.sermon = sermon;
      g.rows.push(r);
    }

    // Sort by date desc (nodate last)
    const arr = [...groups.values()].sort((a,b)=>{
      if(a.date && b.date) return b.date - a.date;
      if(a.date && !b.date) return -1;
      if(!a.date && b.date) return 1;
      return 0;
    });

    const thisG = arr[0];
    const lastG = arr[1];

    renderSetlistGroup(thisG, "#setlist-this-meta", "#setlist-this", idxSong, idxNotes);
    renderSetlistGroup(lastG, "#setlist-last-meta", "#setlist-last", idxSong, idxNotes);

    // --- Analytics based on ALL rows ---
    if(idxSong !== -1){
      const counts = new Map();
      for(const r of rows){
        const s = String(r[idxSong] ?? "").trim();
        if(s) counts.set(s, (counts.get(s)||0)+1);
      }
      const sorted = [...counts.entries()].sort((a,b)=>b[1]-a[1]);
      const top5 = sorted.slice(0,5);
      const bottom5 = sorted.slice(-5).reverse();

      $("#top5").innerHTML = top5.length ? top5.map(([s,c])=>`<li>${escapeHtml(s)} — ${c}</li>`).join("") : `<li class="dim">No data</li>`;
      $("#bottom5").innerHTML = bottom5.length ? bottom5.map(([s,c])=>`<li>${escapeHtml(s)} — ${c}</li>`).join("") : `<li class="dim">No data</li>`;

      drawPie(sorted.slice(0,7));
      drawBar(top5);
    }else{
      $("#top5").innerHTML = `<li class="dim">Add a "Song" header to enable analytics.</li>`;
      $("#bottom5").innerHTML = `<li class="dim">Add a "Song" header to enable analytics.</li>`;
    }
  }catch(e){
    console.error(e);
    $("#setlist-this").innerHTML = `<p class="dim">Unable to load <code>${PATHS.setlist}</code>.</p>`;
    $("#setlist-last").innerHTML = `<p class="dim">—</p>`;
  }
}

function renderSetlistGroup(group, metaSel, tableSel, idxSong, idxNotes){
  if(!group){ $(metaSel).textContent = "—"; $(tableSel).innerHTML = `<p class="dim">No data.</p>`; return; }
  $(metaSel).textContent = `${group.date ? "Service Date: " + fmtDate(group.date) + " · " : ""}${group.sermon ? "Sermon: " + group.sermon : "Sermon: —"}`;

  // Header WITHOUT "Key"
  const header = ["Song","Notes",""];
  // Build all rows for this date group
  const rows = group.rows.map((r,i)=>{
    const song = idxSong !== -1 ? String(r[idxSong] ?? "") : "";
    const note = idxNotes !== -1 ? String(r[idxNotes] ?? "") : "";
    return { song, note, i };
  });

  // Render table + story toggles
  let html = `<table><thead><tr>${header.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>`;
  rows.forEach(({song,note,i})=>{
    html += `<tr>
      <td>${escapeHtml(song)}
        <span class="story-btn" data-song="${escapeHtml(song)}" data-target="story-${tableSel}-${i}">Story</span>
        <div id="story-${tableSel}-${i}" class="story" style="display:none;"></div>
      </td>
      <td>${escapeHtml(note)}</td>
      <td></td>
    </tr>`;
  });
  html += `</tbody></table>`;
  $(tableSel).innerHTML = html;

  // Wire story buttons
  $(tableSel).querySelectorAll(".story-btn").forEach(btn=>{
    btn.addEventListener("click", async ()=>{
      const box = document.getElementById(btn.getAttribute("data-target"));
      const song = btn.getAttribute("data-song");
      if(box.style.display==="none"){
        box.style.display = "block";
        box.innerHTML = `<span class="dim">Fetching story…</span>`;
        const text = await fetchSongStory(song);
        box.innerHTML = text ? escapeHtml(text) : `<span class="dim">No summary found.</span>`;
      }else{
        box.style.display = "none";
      }
    });
  });
}

// ---------- Song story (Wikipedia) ----------
async function fetchSongStory(query){
  try{
    let summary = await wikipediaSummary(query);
    if(summary) return summary;

    const alt = await wikipediaOpenSearch(query);
    if(alt){
      summary = await wikipediaSummary(alt);
      if(summary) return summary;
    }
    summary = await wikipediaSummary(`${query} (hymn)`);
    return summary || "";
  }catch{ return ""; }
}

async function wikipediaSummary(title){
  const url = `https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(title)}`;
  const res = await fetch(url, { headers:{ "accept":"application/json" } });
  if(!res.ok) return "";
  const j = await res.json();
  return j?.extract || "";
}

async function wikipediaOpenSearch(q){
  const url = `https://en.wikipedia.org/w/api.php?action=opensearch&search=${encodeURIComponent(q)}&limit=1&namespace=0&format=json&origin=*`;
  const res = await fetch(url);
  if(!res.ok) return "";
  const j = await res.json();
  return j?.[1]?.[0] || "";
}

// ---------- Charts (pie & bar) ----------
let pieInst, barInst;
function destroyChart(inst){ if(inst){ inst.destroy(); } }

function drawPie(entries){
  const ctx = $("#pieChart");
  destroyChart(pieInst);
  pieInst = new Chart(ctx, {
    type: "pie",
    data: { labels: entries.map(e=>e[0]), datasets:[{ data: entries.map(e=>e[1]) }] },
    options: { responsive:true, plugins:{ legend:{ position:"bottom" } } }
  });
}

function drawBar(entries){
  const ctx = $("#barChart");
  destroyChart(barInst);
  barInst = new Chart(ctx, {
    type: "bar",
    data: { labels: entries.map(e=>e[0]), datasets:[{ data: entries.map(e=>e[1]) }] },
    options: { responsive:true, scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } } } }
  });
}

// ---------- BOOT ----------
document.addEventListener("DOMContentLoaded", async ()=>{
  await loadAnnouncements();
  await loadReminders();
  await loadMembers();
  await loadSetlistsAndAnalytics();
});
