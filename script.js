/* ================== CONFIG ================== */
const OWNER  = "YSayaovong";
const REPO   = "HFBC_Praise_Worship";
const BRANCH = "main";
const SHEETS = {
  announcements: "announcements/announcements.xlsx",
  bible:         "bible_study/bible_study.xlsx",
  members:       "members/members.xlsx",
  setlist:       "setlist/setlist.xlsx",
};

/* CDN mirrors for resilience */
const cdnMirrors = [
  (p)=>`https://cdn.jsdelivr.net/gh/${OWNER}/${REPO}@${BRANCH}/${p}`,
  (p)=>`https://fastly.jsdelivr.net/gh/${OWNER}/${REPO}@${BRANCH}/${p}`,
  (p)=>`https://gcore.jsdelivr.net/gh/${OWNER}/${REPO}@${BRANCH}/${p}`,
];

/* ================== UTILITIES ================== */
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

function diag(msg){
  const box = $("#diag"); box.hidden = false;
  const line = document.createElement("div");
  line.textContent = msg;
  box.appendChild(line);
}

/* ================== Robust Excel loader ================== */
async function loadExcel(path){
  // Warn if opened as file://
  if (location.protocol === "file:") {
    diag("You opened index.html via file://. Please serve over HTTP(S) (GitHub Pages, Netlify, or `python -m http.server`).");
  }
  // Try mirrors
  let lastErr;
  for (const m of cdnMirrors) {
    const url = m(path) + `?v=${Date.now()}`; // cache-buster
    try {
      const res = await fetch(url, {mode:"cors"});
      if(!res.ok) throw new Error(`HTTP ${res.status} ${url}`);
      const ab = await res.arrayBuffer();
      return parseXLSX(ab);
    } catch (e) {
      lastErr = e;
      diag(`CDN fetch failed: ${e.message}`);
    }
  }
  // Fallback: GitHub API (base64)
  try{
    const api = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${encodeURIComponent(path)}?ref=${BRANCH}`;
    const r = await fetch(api, {headers:{Accept:"application/vnd.github+json"}});
    if(!r.ok) throw new Error(`GitHub API ${r.status} ${api}`);
    const j = await r.json();
    const bin = Uint8Array.from(atob(j.content.replace(/\n/g,"")), c=>c.charCodeAt(0));
    return parseXLSX(bin.buffer);
  }catch(e){
    diag(`GitHub API fallback failed: ${e.message}`);
    throw lastErr || e;
  }
}
function parseXLSX(ab){
  const wb = XLSX.read(ab, {type:"array"});
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {defval:""});
}

/* ================== Wikipedia helpers ================== */
async function wikiSummary(title){
  try{
    const s = await fetch(`https://en.wikipedia.org/w/api.php?action=opensearch&format=json&origin=*&limit=1&search=${encodeURIComponent(title)}`);
    const sj = await s.json();
    const page = (sj[1] && sj[1][0]) ? sj[1][0] : title;
    const r = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(page)}`);
    if(!r.ok) return null;
    return await r.json(); // {title, extract}
  }catch{ return null; }
}
function inferAuthors(title, extract){
  const curated = {
    "Be Thou My Vision":"Ancient Irish text (attrib. Dallán Forgaill); English vers. Eleanor Hull (1912)",
    "You Raise Me Up":"Rolf Løvland (music), Brendan Graham (lyrics)",
    "10,000 Reasons":"Matt Redman, Jonas Myrin",
  };
  if(curated[title]) return curated[title];
  const m = /by ([A-Z][A-Za-z’' .-]+(?:,? [A-Z][A-Za-z’' .-]+)*)/.exec(extract||"");
  return m ? m[1] : "";
}

/* ================== HOME BLOCKS ================== */
async function renderHomeBlocks(){
  try{
    const rows = await loadExcel(SHEETS.announcements);
    const a = rows.sort((x,y)=> new Date(y.Date||y.date) - new Date(x.Date||x.date))
                  .map(r=>[fmtDate(r.Date||r.date), r.Announcement || r.Text || ""]);
    $("#announcements-table").innerHTML = tableFromRows(["Date","Announcement"], a);
  }catch(e){
    $("#announcements-table").innerHTML = `<div class="dim">Failed to load announcements.</div>`;
  }

  try{
    const rows = await loadExcel(SHEETS.bible);
    const headers = Object.keys(rows[0]||{});
    $("#bible-study-table").innerHTML = tableFromRows(headers, rows.map(r=> headers.map(h=> r[h])));
  }catch(e){
    $("#bible-study-table").innerHTML = `<div class="dim">Failed to load reminders.</div>`;
  }

  try{
    const rows = await loadExcel(SHEETS.members);
    const m = rows.map(r=>[ r.Name || r.NAME || "", r.Role || r.ROLE || "" ]);
    $("#members-table").innerHTML = tableFromRows(["Name","Role"], m);
  }catch(e){
    $("#members-table").innerHTML = `<div class="dim">Failed to load members.</div>`;
  }
}

/* ================== SONGS & ANALYTICS ================== */
async function renderSongs(){
  let raw;
  try{ raw = await loadExcel(SHEETS.setlist); }
  catch(e){
    $("#next-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    $("#last-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    return;
  }

  const rows = raw.map(r=>({
    Date: r.Date || r.Sunday || r.ServiceDate || "",
    Song: r.Song || r.Title || r.Name || "",
    Topic: r.Topic || r.Theme || "",
    Credit: r.Credit || r.Note || "",
    CCLI: r.CCLI || r.Ccli || "",
    PublicDomain: String(r.PublicDomain||"").toLowerCase()==="true",
    Year: r.Year || r.Published || ""
  })).filter(r=> r.Song);

  const sorted = [...rows].sort((a,b)=> new Date(a.Date)-new Date(b.Date));
  const dates  = [...new Set(sorted.map(r=>r.Date).filter(Boolean))].sort((a,b)=> new Date(a)-new Date(b));
  const nextDate = dates.at(-1);
  const lastDate = dates.length>1 ? dates.at(-2) : null;

  async function enrich(r){
    if(r.Credit) return {...r, Story:r.Credit, Author:""};
    const sum = await wikiSummary(r.Song);
    const story = sum?.extract || "";
    const author = inferAuthors(r.Song, story);
    return {...r, Story:story, Author:author};
  }

  if(nextDate){
    $("#next-date").textContent = fmtDate(nextDate);
    const nextRows = sorted.filter(r=> r.Date===nextDate);
    const enriched = await Promise.all(nextRows.map(enrich));
    $("#next-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enriched.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  }else{
    $("#next-week-table").innerHTML = `<div class="dim">No upcoming set found.</div>`;
  }

  if(lastDate){
    $("#last-date").textContent = fmtDate(lastDate);
    const lastRows = sorted.filter(r=> r.Date===lastDate);
    const enrichedL = await Promise.all(lastRows.map(enrich));
    $("#last-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enrichedL.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  }else{
    $("#last-week-table").innerHTML = `<div class="dim">No prior week found.</div>`;
  }

  // Analytics (current year)
  const yearNow = new Date().getFullYear();
  const yearRows = rows.filter(r=> new Date(r.Date).getFullYear()===yearNow);

  const plays = new Map();
  for(const r of yearRows){ const s=r.Song.trim(); if(s) plays.set(s,(plays.get(s)||0)+1); }
  const rank=[...plays.entries()].sort((a,b)=>b[1]-a[1]);
  const top5=rank.slice(0,5), bottom5=rank.slice(-5).reverse();

  $("#top5").innerHTML = top5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;
  $("#bottom5").innerHTML = bottom5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;

  if(window.Chart){
    const labels = top5.map(([s])=>s), data = top5.map(([,c])=>c);
    new Chart($("#barChart"), {type:"bar",data:{labels,datasets:[{label:"Plays",data}]},options:{plugins:{legend:{display:false}}}});
    new Chart($("#pieChart"), {type:"pie",data:{labels,datasets:[{data}]}} );
  }

  // Library & CCLI
  const lib=new Map();
  for(const r of yearRows){
    const s=r.Song.trim(); if(!s) continue;
    const e=lib.get(s)||{plays:0,ccli:"",pd:false,year:null};
    e.plays++; if(r.CCLI) e.ccli=r.CCLI; if(r.PublicDomain) e.pd=true; if(r.Year) e.year=+r.Year;
    lib.set(s,e);
  }
  $("#library-table").innerHTML = tableFromRows(
    ["Song","Plays","Status"],
    [...lib.entries()].sort((a,b)=>b[1].plays-a[1].plays).map(([s,v])=>{
      const isPD = v.pd || (v.year && v.year<=1929);
      const status = v.ccli && !isPD ? "Report to CCLI" : (isPD ? "Public Domain" : "Unknown");
      return [s,v.plays,status];
    })
  );

  const ccli=[], pd=[];
  for(const [s,v] of lib){
    const isPD=v.pd || (v.year && v.year<=1929);
    if(v.ccli && !isPD) ccli.push([s,v.plays,v.ccli]);
    else if(isPD)       pd.push([s,v.plays,v.pd?"PublicDomain: TRUE":(v.year?`Year: ${v.year}`:"—")]);
  }
  $("#ccli-report").innerHTML = tableFromRows(["Song","Plays","CCLI"], ccli);
  $("#ccli-pd").innerHTML     = tableFromRows(["Song","Plays","Basis"], pd);

  $("#btn-export-ccli").onclick = ()=>{
    const data=ccli.map(([Song,Plays,CCLI])=>({Song,Plays,CCLI}));
    const ws=XLSX.utils.json_to_sheet(data), wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"CCLI"); XLSX.writeFile(wb,"ccli_report.csv");
  };
}

/* ================== INIT ================== */
(async function(){
  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
  await renderHomeBlocks();
  await renderSongs();
})();
