/* Hmong First Baptist Church – Praise & Worship
   Single-page site with portfolio-style Projects section.
   Loads weekly Excel setlists + announcements directly from GitHub.
   Excel conventions:
     - Weekly setlist file: setlists/YYYY-MM-DD.xlsx
       Sheet 1 (any name) columns suggestion: Song | Key | Notes | Link | CCLI | YouTube (optional)
       Optional sheet "Meta" with rows: Field | Value (e.g., "YouTube" | <url>)
       Optional sheet "Team" with columns: Role | Members
     - Announcements file: announcements/announcements.xlsx (first row headers; include Date | Title | Details)
*/

// >>> EDIT THIS with your repo details <<<
const GH = {
  owner: "YSayaovong",            // GitHub username/org
  repo:  "hfb-pw-site",           // repository name
  branch:"main"
};
const PATHS = {
  setlists: "setlists",
  ann:      "announcements/announcements.xlsx"
};

// ---------- helpers ----------
const apiURL = (p) => `https://api.github.com/repos/${GH.owner}/${GH.repo}/contents/${encodeURIComponent(p)}?ref=${encodeURIComponent(GH.branch)}`;
const rawURL = (p) => `https://raw.githubusercontent.com/${GH.owner}/${GH.repo}/${GH.branch}/${p}`;
const $ = (sel) => document.querySelector(sel);

async function listDir(path){
  const res = await fetch(apiURL(path), { headers:{ "Accept":"application/vnd.github+json" }});
  if(!res.ok) return [];
  return res.json();
}
async function fetchWB(path){
  const res = await fetch(rawURL(path));
  if(!res.ok) throw new Error(`Fetch error: ${res.status} ${res.statusText}`);
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type:"array" });
}
function firstSheetAOA(wb){
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header:1, defval:"" });
}
function toNiceDateFromName(name){
  const m = name.match(/^(\d{4})-(\d{2})-(\d{2})\.xlsx$/);
  if(!m) return name;
  const d = new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`);
  return d.toLocaleDateString(undefined,{ weekday:"long", year:"numeric", month:"long", day:"numeric" });
}
function sortSetlistNamesDesc(files){
  return files.sort((a,b)=> b.localeCompare(a)); // YYYY-MM-DD.xlsx
}
function renderTable(aoa, targetSel, caption){
  const target = $(targetSel);
  if(!target) return;
  if(!aoa || aoa.length===0){ target.textContent = "No data."; return; }
  let html = caption ? `<p class="dim">${caption}</p>` : "";
  html += "<table>";
  aoa.forEach((row,i)=>{
    html += "<tr>";
    row.forEach(cell=>{
      const safe = String(cell);
      html += i===0 ? `<th>${safe}</th>` : `<td>${safe}</td>`;
    });
    html += "</tr>";
  });
  html += "</table>";
  target.innerHTML = html;
}
function embedYouTube(url){
  const host = $("#youtube-video");
  host.innerHTML = "";
  if(!url) { host.textContent = "No video for this week."; return; }
  // Extract id for common formats
  let id = "";
  try{
    if(url.includes("youtube.com/watch?v=")) id = new URL(url).searchParams.get("v");
    else if(url.includes("youtu.be/")) id = url.split("youtu.be/")[1].split(/[?&#]/)[0];
    else if(url.includes("/embed/")) id = url.split("/embed/")[1].split(/[?&#]/)[0];
  }catch(_){}
  if(!id){ host.textContent = "Invalid YouTube URL."; return; }
  host.innerHTML = `<iframe src="https://www.youtube.com/embed/${id}" allowfullscreen></iframe>`;
}

// ---------- main sections ----------
async function loadLatestSetlist(){
  const items = await listDir(PATHS.setlists);
  const files = items.filter(i=>i.type==="file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(i.name)).map(i=>i.name);
  if(files.length===0){
    $("#next-date").textContent = "No setlists found.";
    $("#next-setlist").textContent = "Upload a file like setlists/2025-10-05.xlsx";
    return [];
  }
  sortSetlistNamesDesc(files);
  const latest = files[0];
  $("#next-date").textContent = toNiceDateFromName(latest);

  const wb = await fetchWB(`${PATHS.setlists}/${latest}`);
  const aoa = firstSheetAOA(wb);
  renderTable(aoa, "#next-setlist");

  // YouTube: Meta sheet or YouTube column on first sheet
  let yt = "";
  if(wb.Sheets["Meta"]){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Meta"], { header:1, defval:"" });
    const row = rows.find(r => (r[0]||"").toString().toLowerCase()==="youtube");
    yt = row ? row[1] : "";
  }
  if(!yt && aoa[0]){
    const col = aoa[0].findIndex(h=> (h||"").toString().toLowerCase()==="youtube");
    if(col>=0) yt = aoa[1]?.[col] || "";
  }
  embedYouTube(yt);

  // Team (optional)
  if(wb.Sheets["Team"]){
    const t = XLSX.utils.sheet_to_json(wb.Sheets["Team"], { header:1, defval:"" });
    const body = t.slice(1).filter(r=>r[0]||r[1]).map(r=>`<tr><td>${r[0]}</td><td>${r[1]}</td></tr>`).join("");
    $("#team-list").innerHTML = `<table><tr><th>Role</th><th>Members</th></tr>${body}</table>`;
  }else{
    $("#team-list").innerHTML = `<p class="dim">Add a "Team" sheet (Role | Members) in the weekly Excel to display roster.</p>`;
  }

  return files; // for archive & library
}

async function loadArchive(files){
  if(!files || files.length===0){
    $("#archive-list").textContent = "No archive yet.";
    return;
  }
  const html = files.map(f=>{
    const nice = toNiceDateFromName(f);
    return `<div><a href="#" data-file="${f}" class="archive-link">${nice}</a> <span class="dim">(${f})</span></div>`;
  }).join("");
  $("#archive-list").innerHTML = html;

  // click handlers: load selected setlist into Next section (and scroll)
  $("#archive-list").querySelectorAll(".archive-link").forEach(a=>{
    a.addEventListener("click", async (e)=>{
      e.preventDefault();
      const file = a.getAttribute("data-file");
      const wb = await fetchWB(`${PATHS.setlists}/${file}`);
      const aoa = firstSheetAOA(wb);
      $("#next-date").textContent = toNiceDateFromName(file);
      renderTable(aoa, "#next-setlist");

      // YouTube again for selected
      let yt = "";
      if(wb.Sheets["Meta"]){
        const rows = XLSX.utils.sheet_to_json(wb.Sheets["Meta"], { header:1, defval:"" });
        const row = rows.find(r => (r[0]||"").toString().toLowerCase()==="youtube");
        yt = row ? row[1] : "";
      }
      if(!yt && aoa[0]){
        const col = aoa[0].findIndex(h=> (h||"").toString().toLowerCase()==="youtube");
        if(col>=0) yt = aoa[1]?.[col] || "";
      }
      embedYouTube(yt);

      // Team for selected
      if(wb.Sheets["Team"]){
        const t = XLSX.utils.sheet_to_json(wb.Sheets["Team"], { header:1, defval:"" });
        const body = t.slice(1).filter(r=>r[0]||r[1]).map(r=>`<tr><td>${r[0]}</td><td>${r[1]}</td></tr>`).join("");
        $("#team-list").innerHTML = `<table><tr><th>Role</th><th>Members</th></tr>${body}</table>`;
      }else{
        $("#team-list").innerHTML = `<p class="dim">No "Team" sheet in this file.</p>`;
      }

      // Scroll up to the Next section
      document.querySelector("#next").scrollIntoView({ behavior:"smooth" });
    });
  });
}

async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.ann);
    const aoa = firstSheetAOA(wb);
    // Try to sort rows (after header) by Date desc if first column parses
    if(aoa.length>1){
      const header = aoa[0];
      const rows = aoa.slice(1).filter(r=> r.some(c=>String(c).trim()!==""));
      rows.sort((a,b)=> new Date(b[0]) - new Date(a[0]));
      renderTable([header, ...rows], "#announcements-table");
    }else{
      renderTable(aoa, "#announcements-table");
    }
  }catch(e){
    $("#announcements-table").innerHTML = `<p class="dim">Add <code>announcements/announcements.xlsx</code> with columns like: Date | Title | Details.</p>`;
  }
}

async function buildSongLibrary(files){
  try{
    if(!files || files.length===0){
      $("#library-table").innerHTML = `<p class="dim">No setlists yet.</p>`;
      return;
    }
    const map = new Map(); // title -> { key, ccli, link }
    for(const f of files){
      const wb = await fetchWB(`${PATHS.setlists}/${f}`);
      const aoa = firstSheetAOA(wb);
      if(!aoa || aoa.length===0) continue;
      const headers = aoa[0].map(h => String(h).toLowerCase());
      const ti = headers.indexOf("song")>-1 ? headers.indexOf("song") : headers.indexOf("title");
      const ki = headers.indexOf("key");
      const ci = headers.indexOf("ccli");
      const li = headers.indexOf("link");
      for(const row of aoa.slice(1)){
        const title = (row[ti]||"").toString().trim();
        if(!title) continue;
        if(!map.has(title)){
          map.set(title, {
            key: (ki>=0 ? row[ki] : "") || "",
            ccli: (ci>=0 ? row[ci] : "") || "",
            link: (li>=0 ? row[li] : "") || ""
          });
        }
      }
    }
    const header = ["Song","Key","CCLI","Ref"];
    const rows = Array.from(map.entries()).sort((a,b)=> a[0].localeCompare(b[0])).map(([t,meta])=>[
      t, meta.key, meta.ccli, meta.link
    ]);
    // Render
    let html = "<table><tr>";
    header.forEach(h=> html += `<th>${h}</th>`);
    html += "</tr>";
    rows.forEach(r=>{
      html += "<tr>";
      html += `<td>${r[0]}</td>`;
      html += `<td>${r[1]||""}</td>`;
      html += `<td>${r[2]||""}</td>`;
      html += `<td>${r[3] ? `<a href="${r[3]}" target="_blank" rel="noopener">Link</a>` : ""}</td>`;
      html += "</tr>";
    });
    html += "</table>";
    $("#library-table").innerHTML = html || `<p class="dim">No songs found in setlists.</p>`;
  }catch(e){
    $("#library-table").innerHTML = `<p class="dim">Unable to build library.</p>`;
  }
}

// ---------- init ----------
document.addEventListener("DOMContentLoaded", async ()=>{
  // Starter fallbacks appear until files are added
  try{
    const files = await loadLatestSetlist();   // sets Next Service + YouTube + Team (if “Team” sheet)
    await loadArchive(files);                  // clickable archive
    await loadAnnouncements();                 // announcements table
    await buildSongLibrary(files);             // derived library
  }catch(e){
    console.error(e);
    // Friendly starters
    $("#next-date").textContent = "Add your first weekly Excel";
    $("#next-setlist").innerHTML = `<p class="dim">Place <code>setlists/YYYY-MM-DD.xlsx</code> (first sheet = songs). Optional sheet <code>Meta</code> → <code>YouTube</code> for embedded video.</p>`;
    $("#archive-list").innerHTML = `<p class="dim">Archive appears after at least one weekly file is added.</p>`;
    $("#announcements-table").innerHTML = `<p class="dim">Add <code>announcements/announcements.xlsx</code> with columns e.g. Date | Title | Details.</p>`;
    $("#library-table").innerHTML = `<p class="dim">Library builds after weekly files exist.</p>`;
    $("#team-list").innerHTML = `<p class="dim">Add a "Team" sheet (Role | Members) to show roster.</p>`;
  }
});
