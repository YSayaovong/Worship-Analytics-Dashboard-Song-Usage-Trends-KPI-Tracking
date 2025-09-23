/* Hmong First Baptist Church – Praise & Worship
   Static site (no backend) powered by Excel files on GitHub.

   Reads weekly content:
     • /setlist/setlist.xlsx           (preferred "This Coming Week")
     • /setlists/YYYY-MM-DD.xlsx       (archive, also used for Last Week & analytics)
     • /announcements/announcements.xlsx

   Weekly Excel (recommended headers on first sheet):
     Song | Key | Notes | Link | CCLI | Sermon | YouTube | Date (optional)
   Optional sheet "Meta":
     Field | Value    e.g.,  Sermon | <text>   YouTube | <url>   Date | 2025-10-05
*/

// ---- repo config ----
const GH = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };
const PATHS = {
  specialCurrent: "setlist/setlist.xlsx",
  setlistsDir: "setlists",
  announcements: "announcements/announcements.xlsx"
};

// ---- basic helpers ----
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
  return d.toLocaleDateString(undefined,{ month:"short", day:"numeric", year:"numeric" }); // e.g., "Sep 20, 2025"
}

// ---- Excel date + text helpers ----
function excelSerialToDate(val){
  if (typeof val === "number") {
    const o = XLSX.SSF.parse_date_code(val);
    if (o) return new Date(o.y, o.m - 1, o.d);
  }
  if (typeof val === "string" && val.trim() !== "") {
    const num = Number(val);
    if (!Number.isNaN(num)) {
      const o = XLSX.SSF.parse_date_code(num);
      if (o) return new Date(o.y, o.m - 1, o.d);
    }
    const d = new Date(val);
    if (!Number.isNaN(d)) return d;
  }
  return null;
}
function normalizeAnnouncementText(s){
  let out = String(s ?? "");
  out = out.replace(/:(?=\d)/g, ": "); // add space after ":" before a digit
  out = out.replace(/([A-Za-z])(\d{1,2}\/\d{1,2}\/\d{2,4})/g, "$1 $2"); // word then date
  return out.trim();
}
function extractYouTubeId(url){
  try{
    if(!url) return "";
    const u = new URL(url);
    if (u.hostname.includes("youtube.com")) {
      if (u.pathname.startsWith("/watch")) return u.searchParams.get("v") || "";
      if (u.pathname.startsWith("/embed/")) return u.pathname.split("/embed/")[1].split(/[?&#]/)[0] || "";
      if (u.pathname.startsWith("/shorts/")) return u.pathname.split("/shorts/")[1].split(/[?&#]/)[0] || "";
    }
    if (u.hostname.includes("youtu.be")) return u.pathname.slice(1).split(/[?&#]/)[0] || "";
  }catch{}
  if (url.includes("watch?v=")) return url.split("watch?v=")[1].split("&")[0];
  if (url.includes("youtu.be/")) return url.split("youtu.be/")[1].split(/[?&#]/)[0];
  if (url.includes("/embed/"))  return url.split("/embed/")[1].split(/[?&#]/)[0];
  if (url.includes("/shorts/")) return url.split("/shorts/")[1].split(/[?&#]/)[0];
  return "";
}

// ---- smart table renderer (formats DATE columns) ----
function renderTable(aoa, targetSel){
  const el = $(targetSel);
  if(!el){ return; }
  if(!aoa || aoa.length===0){ el.textContent = "No data."; return; }

  const headers = (aoa[0] || []).map(h => String(h));
  const dateCols = new Set();
  headers.forEach((h,i)=>{ if(/date/i.test(h)) dateCols.add(i); });

  function prettyDateCell(v){
    const d = excelSerialToDate(v) || new Date(v);
    return (d && !isNaN(d)) ? niceDate(d) : String(v);
  }

  let html = "<table>";
  aoa.forEach((row,rIdx)=>{
    html += "<tr>";
    row.forEach((cell,cIdx)=>{
      let out = String(cell);
      if(rIdx>0 && dateCols.has(cIdx)) out = prettyDateCell(cell);
      html += (rIdx===0? `<th>${out}</th>` : `<td>${out}</td>`);
    });
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
  const embed = `https://www.youtube.com/embed/${id}`;
  el.innerHTML = `
    <iframe
      src="${embed}"
      allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
      allowfullscreen
    ></iframe>
    <div class="small"><a href="${url}" target="_blank" rel="noopener">Open on YouTube</a></div>
  `;
}

// ---- parse Meta/columns for Sermon/YouTube/Date ----
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
        const d = excelSerialToDate(v) || new Date(v);
        if(!isNaN(d)) serviceDate = d;
      }
    }
  }
  if(aoa && aoa[0]){
    const hdr = aoa[0].map(h=>String(h).toLowerCase());
    const sIdx = hdr.indexOf("sermon");
    const yIdx = hdr.indexOf("youtube");
    const dIdx = hdr.indexOf("date")>-1 ? hdr.indexOf("date")
               : (hdr.indexOf("servicedate")>-1 ? hdr.indexOf("servicedate") : hdr.indexOf("service date"));
    if(!sermon && sIdx>=0) sermon = aoa[1]?.[sIdx] || "";
    if(!youtube && yIdx>=0) youtube = aoa[1]?.[yIdx] || "";
    if(!serviceDate && dIdx>=0){
      const d = excelSerialToDate(aoa[1]?.[dIdx]) || new Date(aoa[1]?.[dIdx]);
      if(!isNaN(d)) serviceDate = d;
    }
  }
  return { sermon, youtube, serviceDate };
}

// ---- announcements ----
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = firstSheetAOA(wb);
    if(aoa.length===0){ $("#announcements-table").textContent = "No announcements."; return; }

    const headers = aoa[0].map(h => String(h));
    const dateIdx = headers.findIndex(h => /date/i.test(h));
    const textIdx = headers.findIndex(h => /(announcement|details?)/i.test(h));
    const idxDate = dateIdx >= 0 ? dateIdx : 0;
    const idxText = textIdx >= 0 ? textIdx : 1;

    const rows = aoa.slice(1)
      .filter(r => r.some(c => String(c).trim()!==""))
      .map(r => {
        const d = excelSerialToDate(r[idxDate]) || new Date(r[idxDate]);
        const ds = (!isNaN(d)) ? niceDate(d) : String(r[idxDate]);
        const txt = normalizeAnnouncementText(r[idxText] || "");
        return { d: (!isNaN(d) ? d : null), ds, txt };
      });

    rows.sort((a,b)=>{
      if(a.d && b.d) return b.d - a.d;
      if(a.d) return -1;
      if(b.d) return 1;
      return 0;
    });

    const pretty = [["DATE","ANNOUNCEMENT"], ...rows.map(r => [r.ds, r.txt])];
    renderTable(pretty, "#announcements-table");
  }catch(e){
    $("#announcements-table").innerHTML = `<p class="dim">Add <code>${PATHS.announcements}</code> with headers like: Date | Title/Announcement | Details.</p>`;
  }
}

// ---- setlists (This Coming Week & Last Week) ----
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

  let nextRendered = false;
  let specialMeta = null;

  if (specialExists) {
    try {
      const wb = await fetchWB(PATHS.specialCurrent);
      const aoa = firstSheetAOA(wb);
      const { sermon, youtube, serviceDate } = parseMeta(wb, aoa);

      $("#next-date").textContent = serviceDate ? niceDate(serviceDate) : niceDate(new Date());
      renderTable(aoa, "#next-setlist");
      renderSermon("#next-sermon", sermon);
      renderYouTube("#youtube-video", youtube);
      nextRendered = true;

      let last = null;
      if (serviceDate) {
        const before = dated.filter(f => f.date < serviceDate);
        last = before[before.length - 1] || null;
      } else {
        last = dated[dated.length - 1] || null;
      }

      if (last) {
        $("#last-date").textContent = niceDate(last.date);
        const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
        const aoa2 = firstSheetAOA(wb2);
        renderTable(aoa2, "#last-setlist");
        const meta2 = parseMeta(wb2, aoa2);
        renderSermon("#last-sermon", meta2.sermon);
      } else {
        $("#last-date").textContent = "—";
        $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
      }

      specialMeta = { wb, aoa, serviceDate };
    } catch {
      nextRendered = false;
    }
  }

  if (!nextRendered) {
    if (dated.length === 0) {
      $("#next-date").textContent = "No setlists yet.";
      $("#next-setlist").innerHTML = `<p class="dim">Upload weekly files to <code>${PATHS.setlistsDir}/YYYY-MM-DD.xlsx</code> or provide <code>${PATHS.specialCurrent}</code>.</p>`;
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">—</p>`;
      $("#youtube-video").textContent = "No video for this week.";
      return { dated, specialMeta };
    }

    const today = new Date(); today.setHours(0,0,0,0);
    let nextIdx = dated.findIndex(f => f.date >= today);
    if (nextIdx === -1) nextIdx = dated.length - 1;

    const next = dated[nextIdx];
    const wb = await fetchWB(`${PATHS.setlistsDir}/${next.name}`);
    const aoa = firstSheetAOA(wb);
    renderTable(aoa, "#next-setlist");
    $("#next-date").textContent = niceDate(next.date);
    const meta = parseMeta(wb, aoa);
    renderSermon("#next-sermon", meta.sermon);
    renderYouTube("#youtube-video", meta.youtube);

    const last = dated[nextIdx - 1] || (dated.length >= 2 ? dated[dated.length - 2] : null);
    if (last) {
      $("#last-date").textContent = niceDate(last.date);
      const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
      const aoa2 = firstSheetAOA(wb2);
      renderTable(aoa2, "#last-setlist");
      const meta2 = parseMeta(wb2, aoa2);
      renderSermon("#last-sermon", meta2.sermon);
    } else {
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
    }
  }

  return { dated, specialMeta };
}

// ---- analytics (Current Year, auto after load) ----
async function buildAnalytics(dated, specialMeta){
  const currentYear = new Date().getFullYear();

  const datedThisYear = (dated || []).filter(
    f => f.date && f.date.getFullYear() === currentYear
  );

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

  // Include special "current" setlist if:
  //  - it has a serviceDate in the current year, OR
  //  - no date provided (assume current year to reflect newest songs)
  const includeSpecial =
    !!specialMeta && (
      !specialMeta.serviceDate ||
      specialMeta.serviceDate.getFullYear() === currentYear
    );

  if (datedThisYear.length === 0 && includeSpecial) {
    await accumulateFromWB(specialMeta.wb);
  } else {
    for (const f of datedThisYear) {
      const wb = await fetchWB(`${PATHS.setlistsDir}/${f.name}`);
      await accumulateFromWB(wb);
    }
    if (includeSpecial) {
      await accumulateFromWB(specialMeta.wb);
    }
  }

  if (songCounts.size === 0) {
    $("#top5").innerHTML = `<li class="dim">No data yet for ${currentYear}.</li>`;
    $("#bottom5").innerHTML = `<li class="dim">No data yet for ${currentYear}.</li>`;
    $("#library-table").innerHTML = `<p class="dim">No setlists found for ${currentYear}.</p>`;
    return;
  }

  const entries = Array.from(songCounts.entries()); // [title, count]
  entries.sort((a,b)=> b[1]-a[1] || a[0].localeCompare(b[0]));

  const top5 = entries.slice(0,5);
  $("#top5").innerHTML = top5
    .map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`)
    .join("");

  const bottom5 = entries.slice().sort((a,b)=> a[1]-b[1] || a[0].localeCompare(b[0])).slice(0,5);
  $("#bottom5").innerHTML = bottom5
    .map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`)
    .join("");

  // Library (current year)
  const header = ["Song","Keys Used","Plays"];
  let html = "<table><tr>" + header.map(h=>`<th>${h}</th>`).join("") + "</tr>";
  const titles = Array.from(songCounts.keys()).sort((a,b)=> a.localeCompare(b));
  for (const t of titles) {
    const keysUsed = songKeys.get(t) ? Array.from(songKeys.get(t)).sort().join(", ") : "";
    const plays = songCounts.get(t) || 0;
    html += `<tr><td>${t}</td><td>${keysUsed}</td><td>${plays}</td></tr>`;
  }
  html += "</table>";
  $("#library-table").innerHTML = html;
}

// ---- boot ----
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
