/* ========= Repo config ========= */
const OWNER  = "YSayaovong";
const REPO   = "HFBC_Praise_Worship";
const BRANCH = "main";

/* ========= Helpers ========= */
const $  = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));
const fmtDate = d => new Date(d).toLocaleDateString("en-US", { month:"short", day:"numeric", year:"numeric" });
function tableFromRows(headers, rows){
  const thead = `<thead><tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${
    rows.length ? rows.map(r=>`<tr>${r.map(c=>`<td>${c ?? ""}</td>`).join("")}</tr>`).join("")
                : `<tr><td colspan="${headers.length}" class="dim">No data</td></tr>`
  }</tbody>`;
  return `<table>${thead}${tbody}</table>`;
}

/* ========= Excel fetch (robust: jsDelivr, then GitHub API base64) ========= */
async function fetchExcel(pathRelative){
  const jsd = `https://cdn.jsdelivr.net/gh/${OWNER}/${REPO}@${BRANCH}/${pathRelative}`;
  try{
    const rows = await fetchExcelFromURL(jsd);
    return rows;
  }catch(e1){
    // Fallback: GitHub API (base64)
    const api = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${encodeURIComponent(pathRelative)}?ref=${BRANCH}`;
    const r = await fetch(api, {headers:{'Accept':'application/vnd.github+json'}});
    if(!r.ok) throw new Error(`GitHub API ${r.status} ${pathRelative}`);
    const j = await r.json();
    const bin = Uint8Array.from(atob(j.content.replace(/\n/g,"")), c=>c.charCodeAt(0));
    return parseXlsxArrayBuffer(bin.buffer);
  }
}
async function fetchExcelFromURL(url){
  const r = await fetch(url, {mode:"cors", cache:"no-store"});
  if(!r.ok) throw new Error(`HTTP ${r.status} @ ${url}`);
  const ab = await r.arrayBuffer();
  return parseXlsxArrayBuffer(ab);
}
function parseXlsxArrayBuffer(ab){
  // Load SheetJS dynamically when needed on pages that use Excel
  if(typeof XLSX === "undefined") throw new Error("XLSX not loaded on this page");
  const wb = XLSX.read(ab, {type:"array"});
  // choose first non-empty sheet
  let ws;
  for(const name of wb.SheetNames){
    const candidate = wb.Sheets[name];
    const json = XLSX.utils.sheet_to_json(candidate, {defval:""});
    if(json.length){ ws = candidate; break; }
  }
  if(!ws) ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {defval:""});
}

/* ========= Wikipedia helpers for story/authors ========= */
async function wikiSummary(title){
  try{
    const s = await fetch(`https://en.wikipedia.org/w/api.php?action=opensearch&format=json&origin=*&search=${encodeURIComponent(title)}&limit=1`);
    const sj = await s.json();
    const page = (sj[1] && sj[1][0]) ? sj[1][0] : title;
    const r = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(page)}`);
    if(!r.ok) return null;
    return await r.json();
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

/* ========= HOME ========= */
async function renderHome(){
  // Announcements
  try{
    const rows = await fetchExcel("announcements/announcements.xlsx");
    const a = rows.sort((x,y)=> new Date(y.Date||y.date) - new Date(x.Date||x.date))
                  .map(r=>[ fmtDate(r.Date || r.date), r.Announcement || r.Text || "" ]);
    $("#announcements-table").innerHTML = tableFromRows(["Date","Announcement"], a);
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<div class="dim">Failed to load announcements.</div>`;
  }

  // Bible study / reminders (render whatever columns provided)
  try{
    const rows = await fetchExcel("bible_study/bible_study.xlsx");
    const headers = Object.keys(rows[0]||{});
    $("#bible-study-table").innerHTML = tableFromRows(headers, rows.map(r=> headers.map(h=> r[h])));
  }catch(e){
    console.error(e);
    $("#bible-study-table").innerHTML = `<div class="dim">Failed to load reminders.</div>`;
  }

  // Members
  try{
    const rows = await fetchExcel("members/members.xlsx");
    const m = rows.map(r=>[ r.Name || r.NAME || "", r.Role || r.ROLE || "" ]);
    $("#members-table").innerHTML = tableFromRows(["Name","Role"], m);
  }catch(e){
    console.error(e);
    $("#members-table").innerHTML = `<div class="dim">Failed to load members.</div>`;
  }

  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
}

/* ========= REPORTING ========= */
async function renderReporting(){
  let rows = [];
  try{
    rows = await fetchExcel("setlist/setlist.xlsx");
  }catch(e){
    console.error(e);
    $("#next-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    $("#last-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    return;
  }

  // Normalize columns
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
  const dates = [...new Set(byDate.map(r=>r.Date).filter(Boolean))].sort((a,b)=> new Date(a)-new Date(b));
  const nextDate = dates[dates.length-1];
  const lastDate = dates.length>1 ? dates[dates.length-2] : null;

  async function enrich(r){
    if(r.Credit) return {...r, Story:r.Credit, Author:""};
    const sum = await wikiSummary(r.Song);
    const story = sum?.extract || "";
    const author = inferAuthors(r.Song, story);
    return {...r, Story:story, Author:author};
  }

  if(nextDate){
    $("#next-date").textContent = fmtDate(nextDate);
    const nextRows = byDate.filter(r=> r.Date===nextDate);
    const e = await Promise.all(nextRows.map(enrich));
    $("#next-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      e.map(s=>[ fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||"" ])
    );
  }else{
    $("#next-week-table").innerHTML = `<div class="dim">No upcoming setlist found.</div>`;
  }

  if(lastDate){
    $("#last-date").textContent = fmtDate(lastDate);
    const lastRows = byDate.filter(r=> r.Date===lastDate);
    const e = await Promise.all(lastRows.map(enrich));
    $("#last-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      e.map(s=>[ fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||"" ])
    );
  }else{
    $("#last-week-table").innerHTML = `<div class="dim">No prior week found.</div>`;
  }

  // Analytics (current year)
  const yearNow = new Date().getFullYear();
  buildAnalytics(norm.filter(r=> new Date(r.Date).getFullYear()===yearNow));

  $("#btn-export-ccli")?.addEventListener("click", exportCcliCsv);
  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
}

function buildAnalytics(rows){
  const plays = new Map();
  for(const r of rows){
    const s = r.Song.trim(); if(!s) continue;
    plays.set(s, (plays.get(s)||0)+1);
  }
  const sorted = [...plays.entries()].sort((a,b)=> b[1]-a[1]);
  const top5 = sorted.slice(0,5);
  const bottom5 = sorted.slice(-5).reverse();

  $("#top5").innerHTML = top5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;
  $("#bottom5").innerHTML = bottom5.map(([s,c])=>`<li>${s} — <span class="dim">${c}</span></li>`).join("") || `<li class="dim">No data</li>`;

  const labels = top5.map(([s])=>s);
  const data   = top5.map(([,c])=>c);
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
  const url  = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "ccli_report.csv"; a.click();
  URL.revokeObjectURL(url);
}

/* ========= HYMNAL: list PDFs from repo path /song_pdf ========= */
async function renderHymnal(){
  // List files via GitHub API
  const api = `https://api.github.com/repos/${OWNER}/${REPO}/contents/song_pdf?ref=${BRANCH}`;
  let files = [];
  try{
    const r = await fetch(api, {headers:{'Accept':'application/vnd.github+json'}});
    if(r.ok){
      const j = await r.json();
      files = (Array.isArray(j) ? j : []).filter(x => x.type === "file" && /\.pdf$/i.test(x.name));
    }
  }catch(e){ console.error(e); }

  // Build sets
  const byNumber = new Map();   // number -> {name, html_url}
  const extras   = [];          // non-numbered
  for(const f of files){
    const m = /^(\d+)\.pdf$/i.exec(f.name);
    if(m){
      const n = Number(m[1]);
      byNumber.set(n, f.download_url || f.html_url || f.path);
    }else{
      extras.push({ name: f.name.replace(/\.pdf$/i,""), url: f.download_url || f.html_url || f.path });
    }
  }

  // Hymnal list #1–#352: show only those that exist, in order, open in new tab
  const hymnalUL = $("#hymnal-pdf-list");
  hymnalUL.innerHTML = "";
  for(let n=1; n<=352; n++){
    if(byNumber.has(n)){
      const url = byNumber.get(n);
      const li = document.createElement("li");
      li.innerHTML = `<a target="_blank" rel="noopener" href="${url}">#${n} — Song #${n}</a>`;
      hymnalUL.appendChild(li);
    }
  }
  if(!hymnalUL.children.length){
    hymnalUL.innerHTML = `<li class="dim">No hymnal PDFs found in /song_pdf/.</li>`;
  }

  // Extra (named) PDFs
  const extraUL = $("#extra-pdf-list");
  extras.sort((a,b)=> a.name.localeCompare(b.name));
  extraUL.innerHTML = extras.map(e=> `<li><a target="_blank" rel="noopener" href="${e.url}">${e.name}</a></li>`).join("")
                     || `<li class="dim">No additional PDFs.</li>`;

  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
}

/* ========= Router ========= */
(function init(){
  const page = (location.pathname || "").split("/").pop() || "index.html";
  if(page === "index.html"){
    // SheetJS is needed here; ensure it exists
    if(typeof XLSX === "undefined"){
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js";
      s.onload = renderHome;
      document.head.appendChild(s);
    } else renderHome();
  }else if(page === "reporting.html"){
    if(typeof XLSX === "undefined"){
      const s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js";
      s.onload = renderReporting;
      document.head.appendChild(s);
    } else renderReporting();
  }else if(page === "hymnal.html"){
    renderHymnal();
  }
  $$("#year").forEach(el=> el.textContent = new Date().getFullYear());
})();
