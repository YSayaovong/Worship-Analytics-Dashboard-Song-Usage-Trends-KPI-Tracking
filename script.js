/* ================= Excel sources via jsDelivr (CORS-friendly) ================ */
const GITHUB_OWNER = "YSayaovong";
const GITHUB_REPO  = "HFBC_Praise_Worship";
const BRANCH       = "main";
const XLSX_SOURCES = {
  announcements: jsv("announcements/announcements.xlsx"),
  bible:         jsv("bible_study/bible_study.xlsx"),
  members:       jsv("members/members.xlsx"),
  setlist:       jsv("setlist/setlist.xlsx"),
};
function jsv(path){
  return `https://cdn.jsdelivr.net/gh/${GITHUB_OWNER}/${GITHUB_REPO}@${BRANCH}/${path}`;
}

/* ================= Utilities ================= */
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const fmtDate = d => new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});

function tableFromRows(headers, rows){
  const thead = `<thead><tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${
    rows.length ? rows.map(r=>`<tr>${r.map(c=>`<td>${c ?? ""}</td>`).join("")}</tr>`).join("")
                : `<tr><td colspan="${headers.length}" class="dim">No data</td></tr>`
  }</tbody>`;
  return `<table>${thead}${tbody}</table>`;
}

async function fetchExcel(url){
  const res = await fetch(url, {mode:"cors", cache:"no-store"});
  if(!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  const ab  = await res.arrayBuffer();
  const wb  = XLSX.read(ab, {type:"array"});
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {defval:""});
}

/* ================= Wikipedia helpers (for Story/Authors) ================= */
async function wikiSummary(title){
  try{
    const s = await fetch(`https://en.wikipedia.org/w/api.php?action=opensearch&format=json&origin=*&search=${encodeURIComponent(title)}&limit=1`);
    const sj = await s.json();
    const page = (sj[1] && sj[1][0]) ? sj[1][0] : title;
    const r = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(page)}`);
    if(!r.ok) return null;
    return await r.json(); // has .extract, .title
  }catch{ return null; }
}
function inferAuthors(title, extract){
  const map = {
    "Be Thou My Vision": "Ancient Irish text (attrib. Dallán Forgaill); English vers. Eleanor Hull (1912)",
    "You Raise Me Up": "Rolf Løvland (music), Brendan Graham (lyrics)",
    "10,000 Reasons": "Matt Redman, Jonas Myrin",
  };
  if(map[title]) return map[title];
  const m = /by ([A-Z][A-Za-z’' .-]+(?:,? [A-Z][A-Za-z’' .-]+)*)/.exec(extract || "");
  return m ? m[1] : "";
}

/* ================= Home ================= */
async function renderHome(){
  // Announcements
  try{
    const rows = await fetchExcel(XLSX_SOURCES.announcements);
    const aRows = rows
      .sort((a,b)=> new Date(b.Date||b.date) - new Date(a.Date||a.date))
      .map(a=>[ fmtDate(a.Date || a.date), a.Announcement || a.Text || a.text ]);
    $("#announcements-table").innerHTML = tableFromRows(["Date","Announcement"], aRows);
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<div class="dim">Failed to load announcements.</div>`;
  }

  // Bible study (render raw table as provided)
  try{
    const rows = await fetchExcel(XLSX_SOURCES.bible);
    const headers = Object.keys(rows[0]||{});
    $("#bible-study-table").innerHTML = tableFromRows(headers, rows.map(r=> headers.map(h=> r[h])));
  }catch(e){
    console.error(e);
    $("#bible-study-table").innerHTML = `<div class="dim">Failed to load reminders.</div>`;
  }

  // Members
  try{
    const rows = await fetchExcel(XLSX_SOURCES.members);
    const mRows = rows.map(r=>[ r.Name || r.NAME || "", r.Role || r.ROLE || "" ]);
    $("#members-table").innerHTML = tableFromRows(["Name","Role"], mRows);
  }catch(e){
    console.error(e);
    $("#members-table").innerHTML = `<div class="dim">Failed to load members.</div>`;
  }

  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
}

/* ================= Reporting ================= */
async function renderReporting(){
  let rows = [];
  try{
    rows = await fetchExcel(XLSX_SOURCES.setlist);
  }catch(e){
    console.error(e);
    $("#next-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    $("#last-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    return;
  }

  // Normalize
  const norm = rows.map(r=>({
    Date: r.Date || r.Sunday || r.ServiceDate || "",
    Song: r.Song || r.Title || r.Name || "",
    Topic: r.Topic || r.Theme || "",
    Credit: r.Credit || r.Note || "",
    CCLI: r.CCLI || r.Ccli || "",
    PublicDomain: String(r.PublicDomain||"").toLowerCase()==="true",
    Year: r.Year || r.Published || ""
  })).filter(r=> r.Song);

  // Dates
  const byDate = [...norm].sort((a,b)=> new Date(a.Date) - new Date(b.Date));
  const uniqueDates = [...new Set(byDate.map(r=>r.Date).filter(Boolean))].sort((a,b)=> new Date(a)-new Date(b));
  const nextDate = uniqueDates[uniqueDates.length-1];
  const lastDate = uniqueDates.length>1 ? uniqueDates[uniqueDates.length-2] : null;

  async function enrichRow(r){
    if(r.Credit) return {...r, Story:r.Credit, Author:""};
    const sum = await wikiSummary(r.Song);
    const story = sum?.extract || "";
    const author = inferAuthors(r.Song, story);
    return {...r, Story: story, Author: author};
  }

  if(nextDate){
    $("#next-date").textContent = fmtDate(nextDate);
    const nextRows = byDate.filter(r=> r.Date===nextDate);
    const enriched = await Promise.all(nextRows.map(enrichRow));
    $("#next-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enriched.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  }else{
    $("#next-week-table").innerHTML = `<div class="dim">No upcoming setlist found.</div>`;
  }

  if(lastDate){
    $("#last-date").textContent = fmtDate(lastDate);
    const lastRows = byDate.filter(r=> r.Date===lastDate);
    const enriched = await Promise.all(lastRows.map(enrichRow));
    $("#last-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enriched.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  }else{
    $("#last-week-table").innerHTML = `<div class="dim">No prior week found.</div>`;
  }

  const yearNow = new Date().getFullYear();
  buildAnalytics(norm.filter(r=> new Date(r.Date).getFullYear()===yearNow));

  $("#btn-export-ccli")?.addEventListener("click", exportCcliCsv);
  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
}

function buildAnalytics(rows){
  const plays = new Map();
  for(const r of rows){
    const s = r.Song.trim();
    if(!s) continue;
    plays.set(s, (plays.get(s)||0)+1);
  }
  const sorted = [...plays.entries()].sort((a,b)=> b[1]-a[1]);
  const top5 = sorted.slice(0,5);
  const bottom5 = sorted.slice(-5).reverse();

  $("#top5").innerHTML = top5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;
  $("#bottom5").innerHTML = bottom5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;

  const labels = top5.map(([s])=>s);
  const data = top5.map(([,c])=>c);
  const barCtx = document.getElementById("barChart");
  const pieCtx = document.getElementById("pieChart");
  if(barCtx && window.Chart){
    new Chart(barCtx, { type:"bar", data:{ labels, datasets:[{ label:"Plays", data }] }, options:{ plugins:{legend:{display:false}} } });
  }
  if(pieCtx && window.Chart){
    new Chart(pieCtx, { type:"pie", data:{ labels, datasets:[{ data }] } });
  }

  const lib = new Map();
  for(const r of rows){
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
    else if(isPD)       pdRows.push([s, v.plays, v.pd ? "PublicDomain: TRUE" : (v.year? `Year: ${v.year}` : "—")]);
  }
  $("#ccli-report").innerHTML = tableFromRows(["Song","Plays","CCLI"], ccliRows);
  $("#ccli-pd").innerHTML    = tableFromRows(["Song","Plays","Basis"], pdRows);
}

function exportCcliCsv(){
  const rows = Array.from(document.querySelectorAll("#ccli-report table tr"))
    .map(tr=> Array.from(tr.children).map(td=> `"${td.textContent.replaceAll('"','""')}"`));
  const csv = rows.map(r=> r.join(",")).join("\n");
  const blob = new Blob([csv], {type:"text/csv"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "ccli_report.csv"; a.click();
  URL.revokeObjectURL(url);
}

/* ================= Router ================= */
(function init(){
  const page = (location.pathname || "").split("/").pop() || "index.html";
  if(page==="index.html")      renderHome();
  else if(page==="reporting.html") renderReporting();
  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
})();
