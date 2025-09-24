/* ================== CONFIG (GitHub -> jsDelivr) ================== */
const OWNER  = "YSayaovong";
const REPO   = "HFBC_Praise_Worship";
const BRANCH = "main";
const SHEETS = {
  announcements: "announcements/announcements.xlsx",
  bible:         "bible_study/bible_study.xlsx",
  members:       "members/members.xlsx",
  setlist:       "setlist/setlist.xlsx",
};
const jsd = p => `https://cdn.jsdelivr.net/gh/${OWNER}/${REPO}@${BRANCH}/${p}`;

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
function showError(el, msg, err){
  console.error(msg, err || "");
  el.innerHTML = `<div class="dim">${msg}</div>`;
}

/* ---- robust Excel loader: jsDelivr with GitHub API fallback ---- */
async function loadExcel(path){
  try{
    return await loadExcelFromURL(jsd(path));
  }catch(e1){
    try{
      // Fallback: GitHub API (base64 content)
      const api = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${encodeURIComponent(path)}?ref=${BRANCH}`;
      const r = await fetch(api, {headers:{Accept:"application/vnd.github+json"}});
      if(!r.ok) throw new Error(`GitHub API ${r.status}`);
      const j = await r.json();
      const bin = Uint8Array.from(atob(j.content.replace(/\n/g,"")), c=>c.charCodeAt(0));
      return parseXLSX(bin.buffer);
    }catch(e2){
      throw e2;
    }
  }
}
async function loadExcelFromURL(url){
  const res = await fetch(url, {mode:"cors", cache:"no-store"});
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const ab = await res.arrayBuffer();
  return parseXLSX(ab);
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
  // Announcements
  try{
    const rows = await loadExcel(SHEETS.announcements);
    const a = rows.sort((x,y)=> new Date(y.Date||y.date) - new Date(x.Date||x.date))
                  .map(r=>[fmtDate(r.Date||r.date), r.Announcement || r.Text || ""]);
    $("#announcements-table").innerHTML = tableFromRows(["Date","Announcement"], a);
  }catch(e){ showError($("#announcements-table"), "Failed to load announcements.", e); }

  // Bible / reminders (render as-is)
  try{
    const rows = await loadExcel(SHEETS.bible);
    const headers = Object.keys(rows[0]||{});
    $("#bible-study-table").innerHTML = tableFromRows(headers, rows.map(r=> headers.map(h=> r[h])));
  }catch(e){ showError($("#bible-study-table"), "Failed to load reminders.", e); }

  // Members
  try{
    const rows = await loadExcel(SHEETS.members);
    const m = rows.map(r=>[ r.Name || r.NAME || "", r.Role || r.ROLE || "" ]);
    $("#members-table").innerHTML = tableFromRows(["Name","Role"], m);
  }catch(e){ showError($("#members-table"), "Failed to load members.", e); }
}

/* ================== SONGS & ANALYTICS ================== */
async function renderSongs(){
  let raw;
  try{
    raw = await loadExcel(SHEETS.setlist);
  }catch(e){
    showError($("#next-week-table"), "Failed to load setlist.", e);
    showError($("#last-week-table"), "Failed to load setlist.", e);
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

  // dates
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

  // next week
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

  // last week
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

  // analytics (current year)
  const yearNow = new Date().getFullYear();
  const yearRows = rows.filter(r=> new Date(r.Date).getFullYear()===yearNow);

  const plays = new Map();
  for(const r of yearRows){
    const name = r.Song.trim(); if(!name) continue;
    plays.set(name, (plays.get(name)||0)+1);
  }
  const ranking = [...plays.entries()].sort((a,b)=> b[1]-a[1]);
  const top5 = ranking.slice(0,5), bottom5 = ranking.slice(-5).reverse();

  $("#top5").innerHTML = top5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;
  $("#bottom5").innerHTML = bottom5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;

  if(window.Chart){
    const labels = top5.map(([s])=>s);
    const data   = top5.map(([,c])=>c);
    new Chart($("#barChart"), { type:"bar", data:{ labels, datasets:[{ label:"Plays", data }] }, options:{ plugins:{legend:{display:false}} }});
    new Chart($("#pieChart"), { type:"pie", data:{ labels, datasets:[{ data }] }});
  }

  // library & CCLI
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

  // CSV export (only the “Report to CCLI” bucket)
  $("#btn-export-ccli").onclick = () => {
    const data = ccliRows.map(([Song,Plays,CCLI])=>({Song,Plays,CCLI}));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CCLI");
    XLSX.writeFile(wb, "ccli_report.csv");
  };
}

/* ================== INIT ================== */
(async function(){
  try{
    await renderHomeBlocks();
    await renderSongs();
  }finally{
    $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
  }
})();
