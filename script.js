/* ========= CONFIG: Excel sources (converted to raw URLs) ========= */
const XLSX_SOURCES = {
  announcements: ghRaw("https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx"),
  bible:         ghRaw("https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/bible_study/bible_study.xlsx"),
  members:       ghRaw("https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/members/members.xlsx"),
  setlist:       ghRaw("https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/setlist.xlsx"),
};

// Turn any GitHub blob link into a raw link
function ghRaw(url){
  return url.replace("https://github.com/","https://raw.githubusercontent.com/").replace("/blob/","/");
}

/* ========= UTILITIES ========= */
const $  = sel => document.querySelector(sel);
const $$ = sel => Array.from(document.querySelectorAll(sel));
const fmtDate = d => new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});

function tableFromRows(headers, rows){
  const thead = `<thead><tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.length ? rows.map(r=>`<tr>${r.map(c=>`<td>${c ?? ""}</td>`).join("")}</tr>`).join("") : `<tr><td colspan="${headers.length}" class="dim">No data</td></tr>`}</tbody>`;
  return `<table>${thead}${tbody}</table>`;
}

async function fetchExcel(url){
  const res = await fetch(url, {cache:"no-store"});
  const ab  = await res.arrayBuffer();
  const wb  = XLSX.read(ab, {type:"array"});
  // Use first worksheet by default
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {defval:""});
}

/* ========= WIKIPEDIA HELPERS (for story/author when not in Excel) ========= */
// Get best page title, then summary
async function wikiSummary(title){
  // Search first to resolve exact page
  const s = await fetch(`https://en.wikipedia.org/w/api.php?action=opensearch&format=json&origin=*&search=${encodeURIComponent(title)}&limit=1`);
  const sj = await s.json();
  const page = (sj[1] && sj[1][0]) ? sj[1][0] : title;
  const r = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(page)}`);
  if(!r.ok) return null;
  return await r.json(); // has .title, .extract, .description
}

// naive author extraction heuristic, with curated fallbacks
function inferAuthors(title, extract){
  const map = {
    "Be Thou My Vision": "Ancient Irish text (attrib. Dallán Forgaill); English vers. Eleanor Hull (1912)",
    "You Raise Me Up": "Rolf Løvland (music), Brendan Graham (lyrics)",
    "10,000 Reasons": "Matt Redman, Jonas Myrin"
  };
  if(map[title]) return map[title];
  // crude pattern hunt
  const m = /by ([A-Z][A-Za-z’' .-]+(?:,? [A-Z][A-Za-z’' .-]+)*)/.exec(extract || "");
  return m ? m[1] : "";
}

/* ========= PAGES ========= */
async function renderHome(){
  // Announcements
  try{
    const rows = await fetchExcel(XLSX_SOURCES.announcements);
    const aRows = rows
      .sort((a,b)=> new Date(b.Date||b.date) - new Date(a.Date||a.date))
      .map(a=>[ fmtDate(a.Date || a.date), a.Announcement || a.Text || a.text ]);
    $("#announcements-table").innerHTML = tableFromRows(["Date","Announcement"], aRows);
  }catch(e){
    $("#announcements-table").innerHTML = `<div class="dim">Failed to load announcements.</div>`;
  }

  // Bible study / reminders block
  try{
    const rows = await fetchExcel(XLSX_SOURCES.bible);
    // Expect columns: Day, Time, Note  (flexible: Thursday/Sunday rows)
    const headers = Object.keys(rows[0]||{});
    $("#bible-study-table").innerHTML = tableFromRows(headers, rows.map(r=>headers.map(h=>r[h])));
  }catch(e){
    $("#bible-study-table").innerHTML = `<div class="dim">Failed to load reminders.</div>`;
  }

  // Members
  try{
    const rows = await fetchExcel(XLSX_SOURCES.members);
    // Expect columns: Name, Role
    const headers = Object.keys(rows[0]||{});
    const mRows = rows.map(r=>[r.Name || r.NAME || "", r.Role || r.ROLE || ""]);
    $("#members-table").innerHTML = tableFromRows(["Name","Role"], mRows);
  }catch(e){
    $("#members-table").innerHTML = `<div class="dim">Failed to load members.</div>`;
  }

  $$("#year").forEach(el=>el.textContent = new Date().getFullYear());
}

async function renderReporting(){
  // Load full setlist from Excel (all rows)
  let rows = [];
  try{
    rows = await fetchExcel(XLSX_SOURCES.setlist);
  }catch(e){
    $("#next-week-table").innerHTML = `<div class="dim">Failed to load setlist.</div>`;
    return;
  }

  // Normalize: expected columns -> Date, Song, Topic, Credit, CCLI, PublicDomain, Year
  const norm = rows.map(r=>({
    Date: r.Date || r.date || r.Sunday || r.ServiceDate || "",
    Song: r.Song || r.Title || r.Name || "",
    Topic: r.Topic || r.Theme || "",
    Credit: r.Credit || r.Note || "",
    CCLI: r.CCLI || r.Ccli || "",
    PublicDomain: (String(r.PublicDomain||"").toLowerCase()==="true"),
    Year: r.Year || r.Published || ""
  })).filter(r=>r.Song);

  // Identify next & last by date
  const byDate = [...norm].sort((a,b)=> new Date(a.Date) - new Date(b.Date));
  const uniqueDates = [...new Set(byDate.map(r=>r.Date).filter(Boolean))].sort((a,b)=> new Date(a)-new Date(b));
  const nextDate = uniqueDates[uniqueDates.length-1];
  const lastDate = uniqueDates.length>1 ? uniqueDates[uniqueDates.length-2] : null;

  // Helper to enrich a row with story/author when Credit empty
  async function enrichRow(r){
    if(r.Credit) return {...r, Story:r.Credit, Author:""}; // prefer your note
    const sum = await wikiSummary(r.Song);
    const story = sum?.extract || "";
    const author = inferAuthors(r.Song, story);
    return {...r, Story: story, Author: author};
  }

  // Render “This Coming Week”
  if(nextDate){
    $("#next-date").textContent = fmtDate(nextDate);
    const nextRows = byDate.filter(r=>r.Date===nextDate);
    const enrichedNext = await Promise.all(nextRows.map(enrichRow));
    $("#next-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enrichedNext.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  } else {
    $("#next-week-table").innerHTML = `<div class="dim">No upcoming setlist found.</div>`;
  }

  // Render “Last Week”
  if(lastDate){
    $("#last-date").textContent = fmtDate(lastDate);
    const lastRows = byDate.filter(r=>r.Date===lastDate);
    const enrichedLast = await Promise.all(lastRows.map(enrichRow));
    $("#last-week-table").innerHTML = tableFromRows(
      ["Date","Song","Author","Credit/Story","Topic"],
      enrichedLast.map(s=>[fmtDate(s.Date), s.Song, s.Author||"", s.Story||"", s.Topic||""])
    );
  } else {
    $("#last-week-table").innerHTML = `<div class="dim">No prior week found.</div>`;
  }

  // Build analytics for current year
  const yearNow = new Date().getFullYear();
  const currentYearRows = norm.filter(r=> new Date(r.Date).getFullYear() === yearNow);

  buildAnalytics(currentYearRows);

  // CSV export (from CCLI table)
  $("#btn-export-ccli")?.addEventListener("click", exportCcliCsv);

  $$("#year").forEach(el=>el.textContent = new Date().getFullYear());
}

function buildAnalytics(rows){
  // Aggregate plays per song
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

  // Charts
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

  // Library table + CCLI buckets
  const lib = new Map(); // song -> {plays, ccli, pd, year}
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
    [...lib.entries()].sort((a,b)=> b[1].plays - a[1].plays).map(([s,v])=>{
      const isPD = v.pd || (v.year && v.year <= 1929);
      const status = v.ccli && !isPD ? "Report to CCLI" : (isPD ? "Public Domain" : "Unknown");
      return [s, v.plays, status];
    })
  );

  const ccliRows = [], pdRows = [];
  for(const [s,v] of lib){
    const isPD = v.pd || (v.year && v.year <= 1929);
    if(v.ccli && !isPD){
      ccliRows.push([s, v.plays, v.ccli]);
    } else if(isPD){
      pdRows.push([s, v.plays, v.pd ? "PublicDomain: TRUE" : (v.year? `Year: ${v.year}` : "—")]);
    }
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

/* Hymnal mock index (1–352) */
function renderHymnal(){
  const index = Array.from({length:352}, (_,i)=>({number: i+1, title:`Song #${i+1}`, pdf:`pdf/${i+1}.pdf`}));
  const ul = $("#song-index");
  ul.innerHTML = index.slice(0,200).map(it=>`<li data-pdf="${it.pdf}"><strong>#${it.number}</strong> — ${it.title}</li>`).join("");
  ul.addEventListener("click", e=>{
    const li = e.target.closest("li"); if(!li) return;
    $("#pdf-frame").src = li.dataset.pdf;
  });
  $("#song-search").addEventListener("input", e=>{
    const q = e.target.value.toLowerCase();
    ul.innerHTML = index.filter(it=> String(it.number).includes(q) || it.title.toLowerCase().includes(q))
      .slice(0,200).map(it=>`<li data-pdf="${it.pdf}"><strong>#${it.number}</strong> — ${it.title}</li>`).join("");
  });
  $$("#year").forEach(el=>el.textContent = new Date().getFullYear());
}

/* ========= ROUTER ========= */
(function init(){
  const path = (location.pathname || "").split("/").pop() || "index.html";
  if(path === "index.html") renderHome();
  else if(path === "reporting.html") renderReporting();
  else if(path === "hymnal.html") renderHymnal();
  $$("#year").forEach(el=>el.textContent = new Date().getFullYear());
})();
