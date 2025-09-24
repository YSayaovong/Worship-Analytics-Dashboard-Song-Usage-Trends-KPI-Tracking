// ---------- CONFIG ----------
const GITHUB = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };

const PATHS = {
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  reminders: "reminders/reminders.xlsx", // optional; else we keep static two reminders
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
  return XLSX.read(ab, { type: "array" });
}

function aoaFromWB(wb){
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
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

function parseDateLoose(val){
  if(val == null) return null;
  if(typeof val === "number"){
    try{
      const base = new Date(Date.UTC(1899,11,30));
      const ms = val * 86400000;
      return new Date(base.getTime() + ms);
    }catch{ return null; }
  }
  const s = String(val).trim();
  if(!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function sameYMD(a,b){ return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }

// ---------- LOADERS ----------
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = aoaFromWB(wb);
    // If first column is an Excel serial date, render formatted
    const hdr = aoa[0].map(x=>String(x).toLowerCase());
    const idxDate = hdr.findIndex(h=>["date","service date"].includes(h));
    if(idxDate !== -1){
      const out = aoa.map((row, i)=>{
        if(i===0) return row;
        const r = row.slice();
        const d = parseDateLoose(r[idxDate]);
        r[idxDate] = d ? fmtDate(d) : (r[idxDate] ?? "");
        return r;
      });
      renderAOATable(out, "#announcements-table");
    }else{
      renderAOATable(aoa, "#announcements-table");
    }
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

// ---------- SETLIST with sermon + two weeks + stories ----------
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
    const idxKey    = hdr.findIndex(h=>h==="key");
    const idxNotes  = hdr.findIndex(h=>["notes","comment"].includes(h));

    const rows = aoa.slice(1).filter(r => r.some(c => String(c).trim()!==""));

    // Group rows by date
    let groups = new Map(); // dateKey -> { date:Date|null, sermon?:string, rows:[] }
    for(const r of rows){
      const d = idxDate !== -1 ? parseDateLoose(r[idxDate]) : null;
      const key = d ? d.toISOString().slice(0,10) : "__nodate__";
      const sermon = idxSermon !== -1 ? String(r[idxSermon] ?? "").trim() : "";
      if(!groups.has(key)) groups.set(key, { date:d, sermon: sermon || "", rows:[] });
      const g = groups.get(key);
      // keep first non-empty sermon seen for the date
      if(sermon && !g.sermon) g.sermon = sermon;
      g.rows.push(r);
    }

    // Sort groups by date desc with nodate last
    const arr = [...groups.entries()].sort((a,b)=>{
      const da = a[1].date, db = b[1].date;
      if(da && db) return db - da;
      if(da && !db) return -1;
      if(!da && db) return 1;
      return 0;
    });

    const thisG = arr[0]?.[1];
    const lastG = arr[1]?.[1];

    renderSetlistGroup(thisG, "#setlist-this-meta", "#setlist-this", idxSong, idxKey, idxNotes);
    renderSetlistGroup(lastG, "#setlist-last-meta", "#setlist-last", idxSong, idxKey, idxNotes);

    // Analytics from all rows
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

      // Charts (pie top 7, bar top 5)
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

function renderSetlistGroup(group, metaSel, tableSel, idxSong, idxKey, idxNotes){
  if(!group){ $(metaSel).textContent = "—"; $(tableSel).innerHTML = `<p class="dim">No data.</p>`; return; }
  $(metaSel).textContent = `${group.date ? "Service Date: " + fmtDate(group.date) + " · " : ""}${group.sermon ? "Sermon: " + group.sermon : "Sermon: —"}`;

  // Build minimal AOA
  const header = ["Song","Key","Notes",""];
  const rows = group.rows.map(r=>{
    const song = idxSong !== -1 ? String(r[idxSong] ?? "") : "";
    const key  = idxKey  !== -1 ? String(r[idxKey]  ?? "") : "";
    const note = idxNotes!== -1 ? String(r[idxNotes]?? "") : "";
    return [song, key, note, "story-slot"]; // placeholder token
  });

  // Render table then attach story buttons
  let html = `<table><thead><tr>${header.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>`;
  rows.forEach((r, i)=>{
    html += `<tr>
      <td>${escapeHtml(r[0])} <span class="story-btn" data-song="${escapeHtml(r[0])}" data-target="story-${i}">Story</span>
          <div id="story-${i}" class="story" style="display:none;"></div>
      </td>
      <td>${escapeHtml(r[1])}</td>
      <td>${escapeHtml(r[2])}</td>
      <td></td>
    </tr>`;
  });
  html += `</tbody></table>`;
  $(tableSel).innerHTML = html;

  // Wire up story fetchers
  $(tableSel).querySelectorAll(".story-btn").forEach(btn=>{
    btn.addEventListener("click", async ()=>{
      const target = btn.getAttribute("data-target");
      const song = btn.getAttribute("data-song");
      const box = document.getElementById(target);
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
    // 1) try direct summary
    let summary = await wikipediaSummary(query);
    if(summary) return summary;

    // 2) opensearch -> top title -> summary
    const alt = await wikipediaOpenSearch(query);
    if(alt){
      summary = await wikipediaSummary(alt);
      if(summary) return summary;
    }

    // 3) Try with "(hymn)" suffix if looks like a hymn
    summary = await wikipediaSummary(`${query} (hymn)`);
    return summary || "";
  }catch{
    return "";
  }
}

async function wikipediaSummary(title){
  const url = `https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(title)}`;
  const res = await fetch(url, { headers:{ "accept":"application/json" } });
  if(!res.ok) return "";
  const j = await res.json();
  if(j?.extract) return j.extract;
  return "";
}

async function wikipediaOpenSearch(q){
  const url = `https://en.wikipedia.org/w/api.php?action=opensearch&search=${encodeURIComponent(q)}&limit=1&namespace=0&format=json&origin=*`;
  const res = await fetch(url);
  if(!res.ok) return "";
  const j = await res.json();
  const title = j?.[1]?.[0];
  return title || "";
}

// ---------- Charts (pie & bar only) ----------
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
