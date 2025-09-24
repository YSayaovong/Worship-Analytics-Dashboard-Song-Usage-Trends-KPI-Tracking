// --------- Utilities ---------
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const fmtDate = d => new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});

function tableFromRows(headers, rows){
  const thead = `<thead><tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr></thead>`;
  const body  = rows.length
    ? rows.map(r=>`<tr>${r.map(c=>`<td>${c ?? ""}</td>`).join("")}</tr>`).join("")
    : `<tr><td class="dim" colspan="${headers.length}">No data</td></tr>`;
  return `<table>${thead}<tbody>${body}</tbody></table>`;
}

async function loadExcel(relPath){
  try{
    const res = await fetch(relPath, {cache:"no-store"});
    if(!res.ok) throw new Error(`HTTP ${res.status} for ${relPath}`);
    const ab = await res.arrayBuffer();
    const wb = XLSX.read(ab, {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, {defval:""});
  }catch(err){
    console.error("Excel load failed:", relPath, err);
    return null;
  }
}

// --------- Wikipedia helpers (fallback when Credit is blank) ---------
async function wikiSummary(title){
  try{
    // Find the best page
    const s = await fetch(`https://en.wikipedia.org/w/api.php?action=opensearch&format=json&origin=*&limit=1&search=${encodeURIComponent(title)}`);
    const sj = await s.json();
    const page = (sj[1] && sj[1][0]) ? sj[1][0] : title;
    // Pull summary
    const r = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(page)}`);
    if(!r.ok) return null;
    return await r.json(); // {title, extract, description, ...}
  }catch{ return null; }
}

// Light heuristic for authors if not obvious in Excel
function inferAuthors(title, extract){
  const curated = {
    "Be Thou My Vision":"Ancient Irish text (attrib. Dallán Forgaill); English vers. Eleanor Hull (1912)",
    "You Raise Me Up":"Rolf Løvland (music), Brendan Graham (lyrics)",
    "10,000 Reasons":"Matt Redman, Jonas Myrin"
  };
  if(curated[title]) return curated[title];
  const m = /by ([A-Z][A-Za-z’' .-]+(?:,? [A-Z][A-Za-z’' .-]+)*)/.exec(extract||"");
  return m ? m[1] : "";
}

// --------- Sections: Announcements / Bible / Members ---------
async function renderHomeBlocks(){
  // Announcements
  const ann = await loadExcel("announcements/announcements.xlsx");
  if(ann){
    const rows = ann.sort((a,b)=> new Date(b.Date||b.date) - new Date(a.Date||a.date))
                    .map(r=>[fmtDate(r.Date||r.date), r.Announcement || r.Text || ""]);
    $("#announcements-table").innerHTML = tableFromRows(["Date","Announcement"], rows);
  }else{
    $("#announcements-table").textContent = "Failed to load announcements.";
  }

  // Bible study / reminders (render as-is)
  const bible = await loadExcel("bible_study/bible_study.xlsx");
  if(bible && bible.length){
    const headers = Object.keys(bible[0]);
    const rows = bible.map(r=> headers.map(h=> r[h]));
    $("#bible-study-table").innerHTML = tableFromRows(headers, rows);
  }else{
    $("#bible-study-table").textContent = "Failed to load reminders.";
  }

  // Members
  const members = await loadExcel("members/members.xlsx");
  if(members){
    const rows = members.map(r=>[ r.Name || r.NAME || "", r.Role || r.ROLE || "" ]);
    $("#members-table").innerHTML = tableFromRows(["Name","Role"], rows);
  }else{
    $("#members-table").textContent = "Failed to load members.";
  }
}

// --------- Songs & Analytics (with online enrichment) ---------
async function renderSongs(){
  const raw = await loadExcel("setlist/setlist.xlsx");
  if(!raw){
    $("#next-week-table").textContent = "Failed to load setlist.";
    $("#last-week-table").textContent = "Failed to load setlist.";
    return;
  }

  // Normalize expected columns
  const rows = raw.map(r=>({
    Date: r.Date || r.Sunday || r.ServiceDate || "",
    Song: r.Song || r.Title || r.Name || "",
    Topic: r.Topic || r.Theme || "",
    Credit: r.Credit || r.Note || "",
    CCLI: r.CCLI || r.Ccli || "",
    PublicDomain: `${r.PublicDomain||""}`.toLowerCase()==="true",
    Year: r.Year || r.Published || ""
  })).filter(r=> r.Song);

  // Determine last & next
  const sorted = [...rows].sort((a,b)=> new Date(a.Date)-new Date(b.Date));
  const uniqDates = [...new Set(sorted.map(r=>r.Date).filter(Boolean))]
                    .sort((a,b)=> new Date(a)-new Date(b));
  const nextDate = uniqDates.at(-1);
  const lastDate = uniqDates.length>1 ? uniqDates.at(-2) : null;

  async function enrich(r){
    if(r.Credit){ // prefer Excel credit/note
      return {...r, Story:r.Credit, Author:""};
    }
    const sum = await wikiSummary(r.Song);
    const story = sum?.extract || "";
    const author = inferAuthors(r.Song, story);
    return {...r, Story: story, Author: author};
  }

  // Next week
  if(nextDate){
    $("#next-date").textContent = fmtDate(nextDate);
    const nextRows = sorted.filter(r=>r.Date===nextDate);
    const enriched = await Promise.all(nextRows.map(enrich));
    $("#next-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enriched.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  }else{
    $("#next-week-table").innerHTML = `<div class="dim">No upcoming set found.</div>`;
  }

  // Last week
  if(lastDate){
    $("#last-date").textContent = fmtDate(lastDate);
    const lastRows = sorted.filter(r=>r.Date===lastDate);
    const enrichedL = await Promise.all(lastRows.map(enrich));
    $("#last-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enrichedL.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  }else{
    $("#last-week-table").innerHTML = `<div class="dim">No prior week found.</div>`;
  }

  // ----- Analytics for current year -----
  const yearNow = new Date().getFullYear();
  const yearRows = rows.filter(r=> new Date(r.Date).getFullYear()===yearNow);

  // plays per song
  const plays = new Map();
  for(const r of yearRows){
    const name = r.Song.trim();
    if(!name) continue;
    plays.set(name, (plays.get(name)||0)+1);
  }
  const ranking = [...plays.entries()].sort((a,b)=> b[1]-a[1]);
  const top5 = ranking.slice(0,5);
  const bottom5 = ranking.slice(-5).reverse();

  $("#top5").innerHTML = top5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;
  $("#bottom5").innerHTML = bottom5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;

  // Charts
  const labels = top5.map(([s])=>s);
  const data   = top5.map(([,c])=>c);
  if(window.Chart){
    new Chart($("#barChart"), { type:"bar", data:{ labels, datasets:[{ label:"Plays", data }] }, options:{ plugins:{legend:{display:false}} }});
    new Chart($("#pieChart"), { type:"pie", data:{ labels, datasets:[{ data }] }});
  }

  // Library & CCLI tables
  const lib = new Map();
  for(const r of yearRows){
    const s = r.Song.trim(); if(!s) continue;
    const e = lib.get(s) || {plays:0, ccli:"", pd:false, year:null};
    e.plays += 1;
    if(r.CCLI) e.ccli = r.CCLI;
    if(r.PublicDomain) e.pd = true;
    if(r.Year) e.year = Number(r.Year);
    lib.set(s, e);
  }

  $("#library-table").innerHTML = tableFromRows(
    ["Song","Plays","Status"],
    [...lib.entries()].sort((a,b)=> b[1].plays-a[1].plays).map(([s,v])=>{
      const isPD = v.pd || (v.year && v.year <= 1929);
      const status = v.ccli && !isPD ? "Report to CCLI" : (isPD ? "Public Domain" : "Unknown");
      return [s, v.plays, status];
    })
  );

  const ccliRows = [], pdRows = [];
  for(const [s,v] of lib){
    const isPD = v.pd || (v.year && v.year <= 1929);
    if(v.ccli && !isPD) ccliRows.push([s, v.plays, v.ccli]);
    else if(isPD)       pdRows.push([s, v.plays, v.pd ? "PublicDomain: TRUE" : (v.year ? `Year: ${v.year}` : "—")]);
  }
  $("#ccli-report").innerHTML = tableFromRows(["Song","Plays","CCLI"], ccliRows);
  $("#ccli-pd").innerHTML     = tableFromRows(["Song","Plays","Basis"], pdRows);

  // Export CSV (CCLI section)
  $("#btn-export-ccli").onclick = () => {
    const data = ccliRows.map(([song,plays,ccli])=>({Song:song,Plays:plays,CCLI:ccli}));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CCLI");
    XLSX.writeFile(wb, "ccli_report.csv");
  };
}

// --------- Init ---------
(async function(){
  await renderHomeBlocks();
  await renderSongs();
  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
})();
