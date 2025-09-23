/* Hmong First Baptist Church – Praise & Worship
   Static site (no backend) powered by Excel files on GitHub.

   Sources it will read automatically each week:
     • /setlist/setlist.xlsx           (preferred "This Coming Week" if present)
     • /setlists/YYYY-MM-DD.xlsx       (archive and fallback for Next/Last)
     • /announcements/announcements.xlsx

   Weekly Excel conventions:
     - First sheet (any name) rows = songs. Suggested headers:
         Song | Key | Notes | Link | CCLI | Sermon | YouTube
     - Optional sheet "Meta" with rows:
         Field | Value   e.g.,  Sermon | <text>    YouTube | <url>   Date | 2025-10-05
     - Optional sheet "Team" (not displayed here, but supported if you want to extend):
         Role | Members

   Announcements Excel:
     - announcements.xlsx with headers like: Date | Title | Details
*/

// >>>>>> EDIT REPO INFO <<<<<
const GH = {
  owner: "YSayaovong",
  repo:  "HFBC_Praise_Worship",
  branch:"main"
};

const PATHS = {
  specialCurrent: "setlist/setlist.xlsx",          // optional, used for "This Coming Week"
  setlistsDir:    "setlists",                      // dated archive files
  announcements:  "announcements/announcements.xlsx"
};

// ---------- helpers ----------
const apiURL = (p) => `https://api.github.com/repos/${GH.owner}/${GH.repo}/contents/${encodeURIComponent(p)}?ref=${encodeURIComponent(GH.branch)}`;
const rawURL = (p) => `https://raw.githubusercontent.com/${GH.owner}/${GH.repo}/${GH.branch}/${p}`;
const $ = (s) => document.querySelector(s);

async function existsOnGitHub(path){
  const r = await fetch(apiURL(path), { headers:{ "Accept":"application/vnd.github+json" }});
  return r.ok;
}
async function listDir(path){
  const r = await fetch(apiURL(path), { headers:{ "Accept":"application/vnd.github+json" }});
  if(!r.ok) return [];
  return r.json();
}
async function fetchWB(path){
  const r = await fetch(rawURL(path));
  if(!r.ok) throw new Error(`Fetch error ${r.status} for ${path}`);
  const ab = await r.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}
function firstSheetAOA(wb){
  const sh = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sh, { header:1, defval:"" });
}
function toDateFromName(name){ // 2025-10-05.xlsx
  const m = /^(\d{4})-(\d{2})-(\d{2})\.xlsx$/.exec(name);
  return m ? new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`) : null;
}
function niceDate(d){
  return d.toLocaleDateString(undefined,{ weekday:"long", year:"numeric", month:"long", day:"numeric" });
}
function extractYouTubeId(url){
  try{
    if(!url) return "";
    if(url.includes("watch?v=")) return new URL(url).searchParams.get("v") || "";
    if(url.includes("youtu.be/")) return url.split("youtu.be/")[1].split(/[?&#]/)[0];
    if(url.includes("/embed/")) return url.split("/embed/")[1].split(/[?&#]/)[0];
    return "";
  }catch{ return ""; }
}
function renderTable(aoa, targetSel){
  const el = $(targetSel);
  if(!aoa || aoa.length===0){ el.textContent = "No data."; return; }
  let html = "<table>";
  aoa.forEach((row,i)=>{
    html += "<tr>";
    row.forEach(cell => { html += (i===0? "<th>":"<td>") + String(cell) + (i===0? "</th>":"</td>"); });
    html += "</tr>";
  });
  html += "</table>";
  el.innerHTML = html;
}
function renderSermon(targetSel, sermon){
  const el = $(targetSel);
  el.textContent = sermon ? `Sermon: ${sermon}` : "";
}
function renderYouTube(targetSel, url){
  const el = $(targetSel);
  el.innerHTML = "";
  const id = extractYouTubeId(url);
  if(!id){ el.textContent = "No video for this week."; return; }
  el.innerHTML = `<iframe src="https://www.youtube.com/embed/${id}" allowfullscreen></iframe>`;
}

// Extract Sermon/YouTube/Date from Meta or columns
function parseMeta(wb, aoa){
  let sermon="", youtube="", serviceDate=null;

  if(wb.Sheets["Meta"]){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Meta"], { header:1, defval:"" });
    for(const r of rows){
      const k = String(r[0]||"").toLowerCase();
      const v = r[1] || "";
      if(k==="sermon")  sermon = v || sermon;
      if(k==="youtube") youtube = v || youtube;
      if(k==="date" || k==="servicedate" || k==="service date"){
        const d = new Date(v);
        if(!isNaN(d)) serviceDate = d;
      }
    }
  }
  if(aoa && aoa[0]){
    const hdr = aoa[0].map(h=>String(h).toLowerCase());
    const sIdx = hdr.indexOf("sermon");
    const yIdx = hdr.indexOf("youtube");
    const dIdx = hdr.indexOf("date")>-1 ? hdr.indexOf("date") : (hdr.indexOf("servicedate")>-1 ? hdr.indexOf("servicedate") : hdr.indexOf("service date"));
    if(!sermon && sIdx>=0) sermon = aoa[1]?.[sIdx] || "";
    if(!youtube && yIdx>=0) youtube = aoa[1]?.[yIdx] || "";
    if(!serviceDate && dIdx>=0){
      const d = new Date(aoa[1]?.[dIdx]);
      if(!isNaN(d)) serviceDate = d;
    }
  }
  return { sermon, youtube, serviceDate };
}

// ---------- Announcements ----------
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = firstSheetAOA(wb);
    if(aoa.length>1){
      const header = aoa[0];
      const rows = aoa.slice(1).filter(r => r.some(c => String(c).trim()!==""));
      rows.sort((a,b)=> new Date(b[0]) - new Date(a[0])); // sort by Date desc if present
      renderTable([header, ...rows], "#announcements-table");
    }else{
      renderTable(aoa, "#announcements-table");
    }
  }catch(e){
    $("#announcements-table").innerHTML = `<p class="dim">Add <code>${PATHS.announcements}</code> with headers like: Date | Title | Details.</p>`;
  }
}

// ---------- Setlists (Next & Last) ----------
async function getDatedFiles(){
  const items = await listDir(PATHS.setlistsDir);
  return items
    .filter(it => it.type==="file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(it.name) && toDateFromName(it.name))
    .map(it => ({ name: it.name, date: toDateFromName(it.name) }))
    .sort((a,b)=> a.date - b.date); // oldest -> newest
}

async function loadSetlists(){
  const dated = await getDatedFiles();
  const specialExists = await existsOnGitHub(PATHS.specialCurrent);

  // Load special (preferred for "This Coming Week")
  let nextRendered = false;
  let specialMeta = null;

  if(specialExists){
    try{
      const wb = await fetchWB(PATHS.specialCurrent);
      const aoa = firstSheetAOA(wb);
      const { sermon, youtube, serviceDate } = parseMeta(wb, aoa);

      $("#next-date").textContent = serviceDate ? niceDate(serviceDate) : "This Week";
      renderTable(aoa, "#next-setlist");
      renderSermon("#next-sermon", sermon);
      renderYouTube("#youtube-video", youtube);
      nextRendered = true;

      // Choose "Last Week" based on special date if available; else use latest dated file
      let last = null;
      if(serviceDate){
        const before = dated.filter(f => f.date < serviceDate);
        last = before[before.length - 1] || null;
      }else{
        last = dated[dated.length - 1] || null;
      }

      if(last){
        $("#last-date").textContent = niceDate(last.date);
        const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
        const aoa2 = firstSheetAOA(wb2);
        renderTable(aoa2, "#last-setlist");
        const meta2 = parseMeta(wb2, aoa2);
        renderSermon("#last-sermon", meta2.sermon);
      }else{
        $("#last-date").textContent = "—";
        $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
      }

      specialMeta = { wb, aoa, serviceDate };
    }catch(e){
      // If special fails, fall back to dated logic
      nextRendered = false;
    }
  }

  if(!nextRendered){
    // Use dated files to determine next / last by today
    if(dated.length===0){
      $("#next-date").textContent = "No setlists yet.";
      $("#next-setlist").innerHTML = `<p class="dim">Upload weekly files to <code>${PATHS.setlistsDir}/YYYY-MM-DD.xlsx</code> or provide <code>${PATHS.specialCurrent}</code>.</p>`;
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">—</p>`;
      $("#youtube-video").textContent = "No video for this week.";
      return { dated, specialMeta };
    }

    const today = new Date(); today.setHours(0,0,0,0);
    let nextIdx = dated.findIndex(f => f.date >= today);
    if(nextIdx === -1) nextIdx = dated.length - 1;

    const next = dated[nextIdx];
    const wb = await fetchWB(`${PATHS.setlistsDir}/${next.name}`);
    const aoa = firstSheetAOA(wb);
    renderTable(aoa, "#next-setlist");
    $("#next-date").textContent = niceDate(next.date);
    const meta = parseMeta(wb, aoa);
    renderSermon("#next-sermon", meta.sermon);
    renderYouTube("#youtube-video", meta.youtube);

    const last = dated[nextIdx - 1] || dated[dated.length - 2] || null;
    if(last){
      $("#last-date").textContent = niceDate(last.date);
      const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
      const aoa2 = firstSheetAOA(wb2);
      renderTable(aoa2, "#last-setlist");
      const meta2 = parseMeta(wb2, aoa2);
      renderSermon("#last-sermon", meta2.sermon);
    }else{
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
    }
  }

  return { dated, specialMeta };
}

// ---------- Analytics (Top/Bottom & Library) ----------
async function buildAnalytics(dated, specialMeta){
  // If there are no dated files, optionally include the special current for analytics
  const useSpecialOnly = dated.length === 0 && specialMeta;

  if(dated.length===0 && !useSpecialOnly){
    $("#top5").innerHTML = `<li class="dim">No data yet.</li>`;
    $("#bottom5").innerHTML = `<li class="dim">No data yet.</li>`;
    $("#library-table").innerHTML = `<p class="dim">No setlists found.</p>`;
    return;
  }

  const songCounts = new Map();  // title -> plays
  const songKeys   = new Map();  // title -> Set(keys)

  async function accumulateFromWB(wb){
    const aoa = firstSheetAOA(wb);
    if(!aoa || aoa.length===0) return;
    const headers = aoa[0].map(h => String(h).toLowerCase());
    const ti = headers.indexOf("song") !== -1 ? headers.indexOf("song") : headers.indexOf("title");
    const ki = headers.indexOf("key");
    if(ti === -1) return;

    for(const row of aoa.slice(1)){
      const title = (row[ti] || "").toString().trim();
      if(!title) continue;
      const key = ki>=0 ? (row[ki] || "").toString().trim() : "";

      songCounts.set(title, (songCounts.get(title)||0) + 1);
      if(!songKeys.has(title)) songKeys.set(title, new Set());
      if(key) songKeys.get(title).add(key);
    }
  }

  if(useSpecialOnly){
    await accumulateFromWB(specialMeta.wb);
  }else{
    for(const f of dated){
      const wb = await fetchWB(`${PATHS.setlistsDir}/${f.name}`);
      await accumulateFromWB(wb);
    }
  }

  const entries = Array.from(songCounts.entries()); // [title, count]
  entries.sort((a,b)=> b[1]-a[1] || a[0].localeCompare(b[0]));

  const top5 = entries.slice(0,5);
  $("#top5").innerHTML = top5.length
    ? top5.map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`).join("")
    : `<li class="dim">No data yet.</li>`;

  const bottom5 = entries.slice().sort((a,b)=> a[1]-b[1] || a[0].localeCompare(b[0])).slice(0,5);
  $("#bottom5").innerHTML = bottom5.length
    ? bottom5.map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`).join("")
    : `<li class="dim">No data yet.</li>`;

  const header = ["Song","Keys Used","Plays"];
  let html = "<table><tr>" + header.map(h=>`<th>${h}</th>`).join("") + "</tr>";
  const allTitles = Array.from(songCounts.keys()).sort((a,b)=> a.localeCompare(b));
  for(const t of allTitles){
    const keysUsed = songKeys.get(t) ? Array.from(songKeys.get(t)).sort().join(", ") : "";
    const plays = songCounts.get(t) || 0;
    html += `<tr><td>${t}</td><td>${keysUsed}</td><td>${plays}</td></tr>`;
  }
  html += "</table>";
  $("#library-table").innerHTML = html;
}

// ---------- boot ----------
document.addEventListener("DOMContentLoaded", async ()=>{
  try{
    await loadAnnouncements();
    const { dated, specialMeta } = await loadSetlists();
    await buildAnalytics(dated, specialMeta);
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<p class="dim">Unable to load announcements. Ensure <code>${PATHS.announcements}</code> exists.</p>`;
    $("#next-date").textContent = "—";
    $("#next-setlist").innerHTML = `<p class="dim">Unable to load setlists. Ensure files exist in <code>${PATHS.setlistsDir}/</code> or <code>${PATHS.specialCurrent}</code>.</p>`;
    $("#last-setlist").innerHTML = `<p class="dim">—</p>`;
    $("#youtube-video").textContent = "—";
    $("#top5").innerHTML = `<li class="dim">No data.</li>`;
    $("#bottom5").innerHTML = `<li class="dim">No data.</li>`;
    $("#library-table").innerHTML = `<p class="dim">No data.</p>`;
  }
});
