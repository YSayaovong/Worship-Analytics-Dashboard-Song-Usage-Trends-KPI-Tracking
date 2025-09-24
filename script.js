/* =========================
   HFBC Praise & Worship — FULL script.js
   - Second pie chart: Most Popular Topics (from setlist.xlsx)
   - "Least Popular" songs list replaced with "Top 5 – Most Popular Topics"
   - All dates include weekday (e.g., "Sunday, Sep 28, 2025")
========================= */

/* ---------- CONFIG ---------- */
const GITHUB = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };
const PATHS = {
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  setlist: "setlist/setlist.xlsx",
  bible: "bible_study/bible_study.xlsx",            // Weekly Bible Verses (last 4 weeks)
  special: "special_practice/special_practice.xlsx" // Special Practice (DATE | SPECIAL PRACTICE)
};

// Worship Practice (static times; roll forward every Thu/Sun at local 12:00 AM)
const PRACTICE = {
  thursday: { dow: 4, time: "6:00pm–8:00pm" },
  sunday:   { dow: 0, time: "8:40am–9:30am" }
};

/* ---------- UTIL ---------- */
const $ = (sel, root=document) => root.querySelector(sel);

// Include weekday in all dates: "Sunday, Sep 28, 2025"
const fmtDate = d =>
  d ? d.toLocaleDateString(undefined, {
        weekday: "long", month: "short", day: "numeric", year: "numeric"
      })
    : "—";

const escapeHtml = s => String(s ?? "").replace(/[&<>"']/g, m => ({
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
    if(!row || row.every(c=>String(c).trim()==="")) continue;
    html += `<tr>${header.map((_,j)=>`<td>${escapeHtml(String(row[j] ?? ""))}</td>`).join("")}</tr>`;
  }
  html += `</tbody></table>`;
  el.innerHTML = html;
}

const safeLabel = (s) => {
  const v = String(s ?? "").trim();
  return v || "Unknown";
};

function findFirst(headers, candidates){
  for(const c of candidates){
    const i = headers.indexOf(c);
    if(i !== -1) return i;
  }
  return -1;
}

/* ---------- WORSHIP PRACTICE (table like Special Practice) ---------- */
/* Always show the NEXT Thursday and NEXT Sunday relative to today’s midnight.
   If today is Thu or Sun, roll to the FOLLOWING week (post-midnight behavior). */
function nextOccurrence(targetDow){
  const today = todayLocalMidnight();
  const wd = today.getDay();
  let delta = (targetDow - wd + 7) % 7;
  if(delta === 0) delta = 7; // roll a full week if same day
  const d = new Date(today);
  d.setDate(today.getDate() + delta);
  return d;
}
function loadPractice(){
  const rows = [
    ["Date","Time"],
    [fmtDate(nextOccurrence(PRACTICE.thursday.dow)), PRACTICE.thursday.time],
    [fmtDate(nextOccurrence(PRACTICE.sunday.dow)),   PRACTICE.sunday.time],
  ];
  renderAOATable(rows, "#reminders-table");
}

/* ---------- SPECIAL PRACTICE (special_practice.xlsx) ---------- */
/* Expects headers: DATE | SPECIAL PRACTICE (any case). Future-only.
   Render as Date | Time (we map the sheet’s text into the “Time” column). */
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
    const idxDate = findFirst(hdr, ["date","service date","practice date"]);
    const idxText = findFirst(hdr, ["special practice","special_practice","details","time","reason","notes","note","description","desc","title","topic"]);

    const today = todayLocalMidnight();
    const seen = new Set();
    const items = [];

    for(let i=1;i<aoa.length;i++){
      const r = aoa[i]; if(!r) continue;
      const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
      if(!d || d < today) continue;

      let text = idxText !== -1 ? String(r[idxText] ?? "").trim() : "";
      if(!text){
        text = String(r.find((cell, j) => j !== idxDate && String(cell ?? "").trim() !== "") ?? "").trim();
      }
      if(!text) continue;

      const key = `${d.getTime()}|${text.toLowerCase()}`;
      if(seen.has(key)) continue; seen.add(key);
      items.push({ date: d, text });
    }
    items.sort((a,b)=> a.date - b.date);

    if(items.length === 0){
      $("#special-practice-table").innerHTML = `<p class="dim">No upcoming special practices.</p>`;
      return;
    }

    const out = [["Date","Time"]];
    items.forEach(it => out.push([fmtDate(it.date), it.text]));
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
        row.push(String(r[idxEn] ?? ""), String(r[idxHm] ?? ""));
        out.push(row);
      }
      renderAOATable(out, "#announcements-table");
      return;
    }

    // Fallback: render whole sheet, format Date if present
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

/* ---------- MEMBERS ---------- */
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

/* ---------- WEEKLY BIBLE VERSES (last 4 weeks) ---------- */
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

/* ---------- SETLISTS (Coming Up / Previous) + Analytics ---------- */
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

    const rows = aoa.slice(1).filter(r => r && r.some(c => String(c).trim()!==""));

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

    // --- Analytics (Songs Pie + Topics Pie + Lists) ---
    // SONG COUNTS (all rows)
    let songCounts = [];
    if(idxSong !== -1){
      const counts = new Map();
      for(const r of rows){
        const s = String(r[idxSong] ?? "").trim();
        if(s) counts.set(s, (counts.get(s)||0)+1);
      }
      songCounts = [...counts.entries()].sort((a,b)=>b[1]-a[1]);
    }

    // TOPIC COUNTS (de-dup per service date to avoid inflating)
    let topicCounts = [];
    if(idxSermon !== -1){
      const tCounts = new Map();
      const seen = new Set(); // date|topic
      for(const r of rows){
        const d = idxDate !== -1 ? toLocalDate(r[idxDate]) : null;
        const t = String(r[idxSermon] ?? "").trim();
        if(!t) continue;
        const key = `${d ? d.getTime() : "nodate"}|${t.toLowerCase()}`;
        if(seen.has(key)) continue;
        seen.add(key);
        tCounts.set(t, (tCounts.get(t)||0)+1);
      }
      topicCounts = [...tCounts.entries()].sort((a,b)=>b[1]-a[1]);
    }

    // Lists
    const top5Songs = songCounts.slice(0,5);
    $("#top5").innerHTML = top5Songs.length
      ? top5Songs.map(([s,c])=>`<li>${escapeHtml(safeLabel(s))} — ${c}</li>`).join("")
      : `<li class="dim">No data</li>`;

    const topicsHeader = $("#bottom5")?.previousElementSibling;
    if(topicsHeader) topicsHeader.textContent = "Top 5 – Most Popular Topics";
    const top5Topics = topicCounts.slice(0,5);
    $("#bottom5").innerHTML = top5Topics.length
      ? top5Topics.map(([t,c])=>`<li>${escapeHtml(safeLabel(t))} — ${c}</li>`).join("")
      : `<li class="dim">No topics</li>`;

    // Charts
    drawSongsPie(songCounts.slice(0,7));   // left pie: songs
    drawTopicsPie(topicCounts.slice(0,7)); // right pie: topics
    updateSecondChartTitle("Most Popular Topics (Pie)");
  }catch(e){
    console.error(e);
    $("#setlist-next").innerHTML = `<p class="dim">Unable to load <code>${PATHS.setlist}</code>.</p>`;
    $("#setlist-prev").innerHTML = `<p class="dim">—</p>`;
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

/* ---------- Song story (Wikipedia) ---------- */
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
  const res = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(title)}`, { headers:{ "accept":"application/json" } });
  if(!res.ok) return "";
  const j = await res.json();
  return j?.extract || "";
}
async function wikipediaOpenSearch(q){
  const res = await fetch(`https://en.wikipedia.org/w/api.php?action=opensearch&search=${encodeURIComponent(q)}&limit=1&namespace=0&format=json&origin=*`);
  if(!res.ok) return "";
  const j = await res.json();
  return j?.[1]?.[0] || "";
}

/* ---------- Charts ---------- */
let songsPieInst, topicsPieInst;
function destroyChart(inst){ if(inst){ inst.destroy(); } }

// Left pie: Plays by Song
function drawSongsPie(entries){
  const ctx = $("#pieChart");
  destroyChart(songsPieInst);

  const labels = entries.map(e => safeLabel(e[0]));
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b)=>a+(+b||0), 0) || 1;

  songsPieInst = new Chart(ctx, {
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

// Right pie: Most Popular Topics
function drawTopicsPie(entries){
  const ctx = $("#barChart"); // reuse existing canvas slot
  destroyChart(topicsPieInst);

  const labels = entries.map(e => safeLabel(e[0]));
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b)=>a+(+b||0), 0) || 1;

  topicsPieInst = new Chart(ctx, {
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

// Adjust the title above the right-hand chart card
function updateSecondChartTitle(text){
  const card = $("#barChart")?.closest(".chart-card");
  const h3 = card?.querySelector("h3");
  if(h3) h3.textContent = text;
}

/* ---------- BOOT ---------- */
document.addEventListener("DOMContentLoaded", async ()=>{
  loadPractice(); // render computed Thursday/Sunday table
  try{ await loadMembers(); }catch{}
  try{ await loadSpecialPractice(); }catch{}
  try{ await loadAnnouncements(); }catch{}
  try{ await loadBibleVerses(); }catch{}
  try{ await loadSetlistsAndAnalytics(); }catch{}
});
