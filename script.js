/* =========================
   CONFIG
========================= */
const GITHUB = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };
const PATHS = {
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  reminders: "reminders/reminders.xlsx",   // optional; falls back to two static reminders
  setlist: "setlist/setlist.xlsx",
  bible: "bible_study/bible_study.xlsx",   // last 4 weeks
  special: "special_practice/special_practice.xlsx" // Special Practice (DATE | SPECIAL PRACTICE)
};

/* =========================
   UTIL
========================= */
const $ = (sel, root=document) => root.querySelector(sel);
const fmtDate = d => d ? d.toLocaleDateString(undefined, { month:"short", day:"numeric", year:"numeric" }) : "—";
const escapeHtml = s => s.replace(/[&<>"']/g, m => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[m]));

function todayLocalMidnight(){
  const t = new Date();
  return new Date(t.getFullYear(), t.getMonth(), t.getDate());
}

function rawUrl(pathRel){
  return `https://raw.githubusercontent.com/${GITHUB.owner}/${GITHUB.repo}/${GITHUB.branch}/${pathRel}`;
}

async function fetchWB(pathRel){
  const url = rawUrl(pathRel) + `?nocache=${Date.now()}`;
  const res = await fetch(url);
  if(!res.ok) throw new Error(`Fetch failed: ${url} (${res.status})`);
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
  const ymd = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/;

  let m;
  if ((m = s.match(mdyyyy))) {
    const M = +m[1], D = +m[2], Y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return new Date(Y, M - 1, D);
  }
  if ((m = s.match(ymd))) {
    return new Date(+m[1], +m[2] - 1, +m[3]);
  }

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

const safeLabel = (s) => {
  const v = String(s ?? "").trim();
  return v || "Unknown";
};

/* =========================
   ANNOUNCEMENTS (bilingual if present)
========================= */
function findFirst(headers, candidates){
  for(const c of candidates){
    const i = headers.indexOf(c);
    if(i !== -1) return i;
  }
  return -1;
}

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
        const r = aoa[i]; if(!r) continue;
        if(r.every(c => String(c ?? "").trim()==="")) continue;
        const row = [];
        if(idxDate !== -1){
          const d = toLocalDate(r[idxDate]);
          row.push(d ? fmtDate(d) : String(r[idxDate] ?? ""));
        }
        row.push(String(r[idxEn] ?? ""), String(r[idxHm] ?? ""));
        out.push(row);
      }
      renderAOATable(out, "#announcements-table");
      return;
    }

    const out2 = idxDate === -1 ? aoa : aoa.map((r,i)=>{
      if(i===0) return r;
      const rr = r.slice();
      const d = toLocalDate(rr[idxDate]);
      rr[idxDate] = d ? fmtDate(d) : (rr[idxDate] ?? "");
      return rr;
    });
    renderAOATable(out2, "#announcements-table");
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<p class="dim">Unable to load <code>${PATHS.announcements}</code>.</p>`;
  }
}

/* =========================
   SPECIAL PRACTICE (from special_practice.xlsx)
   Supports headers: DATE | SPECIAL PRACTICE
   - Future dates only (>= today)
   - Shows Date + Details
   - Deduplicates by date+details
========================= */
async function loadSpecialPractice(){
  try{
    const wb = await fetchWB(PATHS.special);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){
      $("#special-practice-table").innerHTML = `<p class="dim">No upcoming special practices.</p>`;
      return;
    }

    // Normalize headers
    const hdrRaw = aoa[0].map(h => String(h).trim());
    const hdr = hdrRaw.map(h => h.toLowerCase());

    // Column indices
    const idxDate  = findFirst(hdr, ["date","service date","practice date"]);
    const idxText  = findFirst(hdr, [
      "special practice","special_practice","practice","details","reason","notes","note","description","desc","title","topic"
    ]);

    const today = todayLocalMidnight();
    const seen = new Set();
    const items = [];

    for(let i=1;i<aoa.length;i++){
      const r = aoa[i]; if(!r) continue;

      const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
      if(!d || d < today) continue; // only upcoming

      let text = idxText !== -1 ? String(r[idxText] ?? "").trim() : "";
      if(!text){
        // fallback: first non-empty non-date cell
        text = String(r.find((cell, j) => j !== idxDate && String(cell ?? "").trim() !== "") ?? "").trim();
      }
      if(!text) continue;

      const key = `${d.getFullYear()}-${d.getMonth()+1}-${d.getDate()}|${text.toLowerCase()}`;
      if(seen.has(key)) continue;
      seen.add(key);

      items.push({ date: d, text });
    }

    items.sort((a,b)=> a.date - b.date);

    if(items.length === 0){
      $("#special-practice-table").innerHTML = `<p class="dim">No upcoming special practices.</p>`;
      return;
    }

    const out = [["Date","Details"]];
    for(const it of items){
      out.push([fmtDate(it.date), it.text]);
    }
    renderAOATable(out, "#special-practice-table");
  }catch(e){
    console.error(e);
    $("#special-practice-table").innerHTML = `<p class="dim">Unable to load special practices.</p>`;
  }
}

/* =========================
   WORSHIP PRACTICE / MEMBERS
========================= */
async function loadReminders(){
  try{
    const wb = await fetchWB(PATHS.reminders);
    const aoa = aoaFromWB(wb);
    renderAOATable(aoa, "#reminders-table");
  }catch(e){
    /* fallback stays */
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

/* =========================
   WEEKLY BIBLE VERSES (last 4 weeks)
========================= */
async function loadBibleVerses(){
  try{
    const wb = await fetchWB(PATHS.bible);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){ $("#bible-verse-table").innerHTML = `<p class="dim">No verses found.</p>`; return; }

    const hdrRaw = aoa[0].map(h=>String(h));
    const hdr = hdrRaw.map(h=>h.trim().toLowerCase());
    const idxDate  = findFirst(hdr, ["date","service date"]);
    const idxVerse = findFirst(hdr, ["verse","bible verse","scripture","scripture text","topic","title"]);
    const idxRef   = findFirst(hdr, ["reference","passage","scripture reference","book/chapter","book chapter"]);

    const rows = aoa.slice(1)
      .filter(r => r && r.some(c => String(c).trim()!==""))
      .map(r => {
        const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
        const verse = idxVerse !== -1 ? String(r[idxVerse] ?? "").trim() : "";
        const ref = idxRef !== -1 ? String(r[idxRef] ?? "").trim() : "";
        return { date: d, verse, ref };
      })
      .filter(x => x.date && (x.verse || x.ref));

    const today = todayLocalMidnight();
    const fourWeeksAgo = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 28);

    let recent = rows
      .filter(x => x.date >= fourWeeksAgo && x.date <= today)
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
    const header = showRef ? ["Date","Verse","Reference"] : ["Date","Verse"];
    const out = [header];
    for(const r of recent){
      out.push(showRef ? [fmtDate(r.date), r.verse || r.ref || "", r.ref] : [fmtDate(r.date), r.verse || r.ref || ""]);
    }
    renderAOATable(out, "#bible-verse-table");
  }catch(e){
    console.error(e);
    $("#bible-verse-table").innerHTML = `<p class="dim">Unable to load <code>${PATHS.bible}</code>.</p>`;
  }
}

/* =========================
   SETLIST (Coming Up / Previous) — NO "Notes"
========================= */
async function loadSetlistsAndAnalytics(){
  try{
    const wb = await fetchWB(PATHS.setlist);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length < 2){ $("#setlist-next").innerHTML=`<p class="dim">No rows.</p>`; return; }

    const hdrRaw = aoa[0].map(h=>String(h));
    const hdr = hdrRaw.map(h=>h.trim().toLowerCase());
    const idxDate   = findFirst(hdr, ["date","service date"]);
    const idxSermon = findFirst(hdr, ["sermon","sermon topic","topic"]);
    const idxSong   = findFirst(hdr, ["song","title","song title"]);

    const rows = aoa.slice(1).filter(r => r.some(c => String(c).trim()!==""));

    // Group by date (only rows with a song)
    const groups = [];
    for(const r of rows){
      const date = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
      const song = idxSong !== -1 ? String(r[idxSong] ?? "").trim() : "";
      const sermon = idxSermon !== -1 ? String(r[idxSermon] ?? "").trim() : "";
      if(!song) continue;

      const key = date ? date.getTime() : NaN;
      let g = groups.find(x => (x.date && date && x.date.getTime()===key));
      if(!g){
        g = { date, sermon: sermon || "", rows: [] };
        groups.push(g);
      }
      if(sermon && !g.sermon) g.sermon = sermon;
      g.rows.push({ song });
    }

    const dated = groups.filter(g=>g.date).sort((a,b)=> a.date - b.date);
    const today = todayLocalMidnight();

    const next = dated.find(g => g.date > today) || null;
    const prev = [...dated].filter(g => g.date <= today).slice(-1)[0] || null;

    renderSetlistGroup(next, "#setlist-next-meta", "#setlist-next");
    renderSetlistGroup(prev, "#setlist-prev-meta", "#setlist-prev");

    // --- Analytics from ALL rows (by song) ---
    if(idxSong !== -1){
      const counts = new Map();
      for(const r of rows){
        const s = String(r[idxSong] ?? "").trim();
        if(s) counts.set(s, (counts.get(s)||0)+1);
      }
      const sorted = [...counts.entries()].sort((a,b)=>b[1]-a[1]);
      const top5 = sorted.slice(0,5);
      const bottom5 = sorted.slice(-5).reverse();

      $("#top5").innerHTML = top5.length
        ? top5.map(([s,c])=>`<li>${escapeHtml(safeLabel(s))} — ${c}</li>`).join("")
        : `<li class="dim">No data</li>`;

      $("#bottom5").innerHTML = bottom5.length
        ? bottom5.map(([s,c])=>`<li>${escapeHtml(safeLabel(s))} — ${c}</li>`).join("")
        : `<li class="dim">No data</li>`;

      drawPie(sorted.slice(0,7));
      drawBar(top5);
    }else{
      $("#top5").innerHTML = `<li class="dim">Add a "Song" header to enable analytics.</li>`;
      $("#bottom5").innerHTML = `<li class="dim">Add a "Song" header to enable analytics.</li>`;
    }
  }catch(e){
    console.error(e);
    $("#setlist-next").innerHTML = `<p class="dim">Unable to load <code>${PATHS.setlist}</code>.</p>`;
    $("#setlist-prev").innerHTML = `<p class="dim">—</p>`;
  }
}

function renderSetlistGroup(group, metaSel, tableSel){
  if(!group){ $(metaSel).textContent = "—"; $(tableSel).innerHTML = `<p class="dim">No data.</p>`; return; }
  $(metaSel).textContent = `${group.date ? "Service Date: " + fmtDate(group.date) + " · " : ""}${group.sermon ? "Sermon: " + group.sermon : "Sermon: —"}`;

  // Columns: Song (with Story toggle)
  const header = ["Song"];
  let html = `<table><thead><tr>${header.map(h=>`<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody>`;
  group.rows.forEach((row, i)=>{
    const id = `${tableSel.replace("#","")}-story-${i}`;
    html += `<tr>
      <td>${escapeHtml(row.song)}
        <span class="story-btn" style="margin-left:8px" data-song="${escapeHtml(row.song)}" data-target="${id}">Story</span>
        <div id="${id}" class="story" style="display:none;"></div>
      </td>
    </tr>`;
  });
  html += `</tbody></table>`;
  $(tableSel).innerHTML = html;

  // Story toggles (Wikipedia summary)
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

/* =========================
   Song story (Wikipedia)
========================= */
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

/* =========================
   Charts (pie with %; bar colors by quarter)
========================= */
let pieInst, barInst;
function destroyChart(inst){ if(inst){ inst.destroy(); } }

function drawPie(entries){
  const ctx = $("#pieChart");
  destroyChart(pieInst);

  const labels = entries.map(e => safeLabel(e[0]));
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b)=>a+(+b||0), 0) || 1;

  pieInst = new Chart(ctx, {
    type: "pie",
    data: { labels, datasets:[{ data }] },
    options: {
      responsive: true,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            generateLabels(chart){
              const base = Chart.defaults.plugins.legend.labels.generateLabels(chart);
              const ds   = chart.data.datasets[0]?.data || [];
              const lbs  = chart.data.labels || [];
              const sum  = ds.reduce((a,b)=>a+(+b||0),0) || 1;
              return base.map((item, i) => {
                const val = +ds[i] || 0;
                const pct = Math.round((val/sum)*1000)/10;
                return { ...item, text: `${safeLabel(lbs[i])} — ${pct}%` };
              });
            }
          }
        },
        tooltip: {
          callbacks: {
            label(ctx){
              const val = +ctx.parsed || 0;
              const pct = Math.round((val/total)*1000)/10;
              return `${safeLabel(ctx.label)}: ${val} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function drawBar(entries){
  const ctx = $("#barChart");
  destroyChart(barInst);

  const labels = entries.map(e => safeLabel(e[0]));
  const data   = entries.map(e => e[1]);

  const quarterColors = ["#4e79a7","#f28e2b","#e15759","#76b7b2"];
  const bg = labels.map((_, i) => quarterColors[i % 4]);
  const border = bg;

  barInst = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "Plays",
        data,
        backgroundColor: bg,
        borderColor: border,
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } } },
      plugins: {
        tooltip: {
          callbacks: {
            title(items){ return items.length ? safeLabel(items[0].label) : ""; },
            label(ctx){ return `Plays: ${+ctx.parsed.y || 0}`; }
          }
        }
      }
    }
  });
}

/* =========================
   BOOT
========================= */
document.addEventListener("DOMContentLoaded", async ()=>{
  try{ await loadReminders(); }catch(e){ /* ignore */ }
  try{ await loadMembers(); }catch(e){ console.error(e); $("#members-table").innerHTML = `<p class="dim">Error loading members.</p>`; }

  try{ await loadSpecialPractice(); }catch(e){ console.error(e); $("#special-practice-table").innerHTML = `<p class="dim">Error.</p>`; }
  try{ await loadAnnouncements(); }catch(e){ console.error(e); $("#announcements-table").innerHTML = `<p class="dim">Error loading announcements.</p>`; }

  try{ await loadBibleVerses(); }catch(e){ console.error(e); $("#bible-verse-table").innerHTML = `<p class="dim">Error loading verses.</p>`; }
  try{ await loadSetlistsAndAnalytics(); }catch(e){ console.error(e); $("#setlist-next").innerHTML = `<p class="dim">Error.</p>`; }
});
