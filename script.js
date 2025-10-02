// ====== CONFIG: RAW sources (new repo) ======
const SRC = {
  announcements: "https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/announcements/announcements.xlsx",
  members:       "https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/members/members.xlsx",
  setlist:       "https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/setlist/setlist.xlsx",
  addPractice:   "https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/special_practice/special_practice.xlsx",
  training:      "https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/special_practice/training.xlsx",
  bibleStudy:    "https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/bible_study/bible_study.xlsx",
};

// ====== Utilities ======
async function fetchXlsxRows(rawUrl, sheet=0){
  const res = await fetch(rawUrl + "?v=" + Date.now());
  if(!res.ok) throw new Error("Fetch failed: " + rawUrl);
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, {type: "array"});
  const ws = typeof sheet==="number" ? wb.Sheets[wb.SheetNames[sheet]] : wb.Sheets[sheet] || wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}
function excelToDate(v){
  if(v==null||v==="") return null;
  if(typeof v==="number"){ const d=XLSX.SSF.parse_date_code(v); return new Date(d.y, d.m-1, d.d, d.H||0, d.M||0, d.S||0); }
  const d=new Date(v); return isNaN(d.getTime())?null:d;
}
const DAY_ABBR = ["Sun","Mon","Tues","Wed","Thurs","Fri","Sat"];
const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
const fmtDateOnly = dt => `${DAY_ABBR[dt.getDay()]}, ${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;
const norm = s => String(s||"").toLowerCase().replace(/[^a-z0-9]+/g,"");
function normMap(r){ const m={}; Object.keys(r||{}).forEach(k=>m[norm(k)]=r[k]); return m; }
function val(m, keys){ for(const k of keys){ const v=m[k]; if(v!=null && String(v)!=="") return v; } return ""; }
function findByIncludes(m, subs){ for(const k of Object.keys(m)){ const a=k.toLowerCase(); if(subs.every(s=>a.includes(s))) return m[k]; } return ""; }

// ====== Weekly Practices (Thu 6–8 PM, Sun 8:40–9:30 AM; no rollover until 11:59 PM) ======
async function renderWeeklyPractices(){
  const tbody = document.getElementById("weekly-practice-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  const now = new Date();
  function nextDow(target){ // allow same-day
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const delta = (target - d.getDay() + 7) % 7;
    d.setDate(d.getDate() + delta); d.setHours(0,0,0,0);
    return d;
  }
  const thurs = nextDow(4);
  const sun   = nextDow(0);
  const rows = [
    { date: thurs, time: "6:00–8:00 PM" },
    { date: sun,   time: "8:40–9:30 AM" },
  ];
  rows.forEach(r=>{ const tr=document.createElement("tr"); tr.innerHTML=`<td>${fmtDateOnly(r.date)}</td><td>${r.time}</td>`; tbody.appendChild(tr); });
}

// ====== Additional Practice ======
async function renderAdditionalPractice(){
  const tbody = document.getElementById("additional-practice-body");
  if(!tbody) return; tbody.innerHTML="";
  try{
    const rows = await fetchXlsxRows(SRC.addPractice);
    rows.forEach(r=>{
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day","servicedate"]));
      const t = val(m, ["time","starttime","practice"]);
      if(!d||!t) return;
      const tr=document.createElement("tr"); tr.innerHTML=`<td>${fmtDateOnly(d)}</td><td>${t}</td>`; tbody.appendChild(tr);
    });
    if(!tbody.children.length) tbody.innerHTML = `<tr><td colspan="2">No additional practices listed.</td></tr>`;
  }catch(e){ console.error(e); tbody.innerHTML = `<tr><td colspan="2">Could not load special practice sheet.</td></tr>`; }
}

// ====== Training ======
async function renderTraining(){
  const tbody = document.getElementById("training-body");
  if(!tbody) return; tbody.innerHTML="";
  try{
    const rows = await fetchXlsxRows(SRC.training);
    rows.forEach(r=>{
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day"]));
      const t = val(m, ["time","starttime"]);
      const passage = val(m, ["passage","topic","study"]);
      const verse = val(m, ["bibleverse","verse","reference"]);
      if(!d || (!t && !passage && !verse)) return;
      const tr=document.createElement("tr"); tr.innerHTML=`<td>${fmtDateOnly(d)}</td><td>${t||""}</td><td>${passage||""}</td><td>${verse||""}</td>`; tbody.appendChild(tr);
    });
    if(!tbody.children.length) tbody.innerHTML = `<tr><td colspan="4">No training entries found.</td></tr>`;
  }catch(e){ console.error(e); tbody.innerHTML = `<tr><td colspan="4">Could not load training sheet.</td></tr>`; }
}

// ====== Members ======
async function renderMembers(){
  const leaderList = document.getElementById("leader-list");
  const musicianList = document.getElementById("musician-list");
  const singersList = document.getElementById("singers-list");
  if(!leaderList||!musicianList||!singersList) return;
  leaderList.innerHTML = musicianList.innerHTML = singersList.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(SRC.members);
    const leaders=[], musicians=[], singers=[];
    rows.forEach(r=>{
      const m = normMap(r);
      const name = val(m, ["name","member","person"]) || "";
      const role = (val(m, ["role","position","type"]) || "").toLowerCase();
      if(!name) return;
      if(role.includes("leader")) leaders.push(name);
      else if(role.includes("singer") || role.includes("vocal")) singers.push(name);
      else musicians.push(name);
    });
    const addAll=(ul,arr)=>{ if(!arr.length) ul.innerHTML="<li class='muted'>None listed</li>"; else arr.forEach(n=>{ const li=document.createElement("li"); li.textContent=n; ul.appendChild(li); }); };
    addAll(leaderList, leaders); addAll(musicianList, musicians); addAll(singersList, singers);
  }catch(e){ console.error(e); leaderList.innerHTML = musicianList.innerHTML = singersList.innerHTML = "<li class='muted'>Could not load members sheet.</li>"; }
}

// ====== Announcements (English + Hmong) ======
async function renderAnnouncements(){
  const tbody = document.getElementById("announcements-body");
  if(!tbody) return; tbody.innerHTML="";
  try{
    const rows = await fetchXlsxRows(SRC.announcements);
    const today = new Date(); const limit = 31*24*60*60*1000;
    const items = rows.map(r=>{
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day"]));
      const en = val(m, ["announcement","announcementenglish","english","announcement_(english)"]);
      // robust Hmong detection + explicit header
      let hm = val(m, ["hmong","lus_tshaj_tawm","lustshajtawm","lus","tshaj"]);
      if(!hm) hm = findByIncludes(m, ["hmong"]) || findByIncludes(m, ["hmoob"]) || findByIncludes(m, ["lus","tshaj"]) || r["LUS TSHAJ TAWM"];
      return { d, en, hm };
    }).filter(x=>x.d && (today - x.d) <= limit).sort((a,b)=>b.d-a.d);
    items.forEach(it=>{ const tr=document.createElement("tr"); tr.innerHTML=`<td>${fmtDateOnly(it.d)}</td><td>${it.en||""}</td><td>${it.hm||""}</td>`; tbody.appendChild(tr); });
    if(!tbody.children.length) tbody.innerHTML = `<tr><td colspan="3">No announcements from the last 31 days.</td></tr>`;
  }catch(e){ console.error(e); tbody.innerHTML = `<tr><td colspan="3">Could not load announcements sheet.</td></tr>`; }
}

// ====== Bible Study ======
function getWeekRange(date){
  const start=new Date(date); start.setDate(start.getDate()-start.getDay()); start.setHours(0,0,0,0);
  const end=new Date(start); end.setDate(start.getDate()+6); end.setHours(23,59,59,999);
  return {start,end};
}
async function renderBibleStudy(){
  const tbody=document.getElementById("bible-study-body");
  if(!tbody) return; tbody.innerHTML="";
  try{
    const rows = await fetchXlsxRows(SRC.bibleStudy);
    const {start,end} = getWeekRange(new Date());
    const prev3=new Date(start); prev3.setDate(prev3.getDate()-21);
    const items = rows.map(r=>{
      const m=normMap(r);
      const d=excelToDate(val(m, ["date","day"]));
      const topic=val(m, ["topic","passage","study"]);
      const verse=val(m, ["bibleverse","verse","reference"]);
      return {d,topic,verse};
    }).filter(x=>x.d && x.d>=prev3 && x.d<=end).sort((a,b)=>b.d-a.d);
    items.forEach(it=>{ const tr=document.createElement("tr"); tr.innerHTML=`<td>${fmtDateOnly(it.d)}</td><td>${it.topic||""}</td><td>${it.verse||""}</td>`; tbody.appendChild(tr); });
    if(!tbody.children.length) tbody.innerHTML = `<tr><td colspan="3">No bible study entries for the last 3 weeks.</td></tr>`;
  }catch(e){ console.error(e); tbody.innerHTML = `<tr><td colspan="3">Could not load bible study sheet.</td></tr>`; }
}

// ====== Setlist ======
function normSetlistRow(r){
  const m=normMap(r);
  const date=excelToDate(val(m, ["date","day","servicedate"]));
  const song=String(val(m, ["song","title"])).trim();
  const topic=String(val(m, ["topic","notes"])).trim();
  return {date,song,topic};
}
function dedupeByTitle(list){
  const seen=new Set(); return list.filter(it=>{const k=(it.song||"").toLowerCase(); if(!k||seen.has(k)) return false; seen.add(k); return true;});
}
async function renderSetlist(){
  const upHead=document.getElementById("setlist-up-head");
  const upBody=document.getElementById("setlist-up-body");
  const lsHead=document.getElementById("setlist-last-head");
  const lsBody=document.getElementById("setlist-last-body");
  if(!upHead||!upBody||!lsHead||!lsBody) return;
  try{
    const rows=(await fetchXlsxRows(SRC.setlist)).map(normSetlistRow).filter(x=>x.date&&x.song);
    const byDate=new Map();
    for(const r of rows){ const k=r.date.toISOString().slice(0,10); if(!byDate.has(k)) byDate.set(k,[]); byDate.get(k).push(r); }
    const dates=Array.from(byDate.keys()).map(d=>new Date(d)).sort((a,b)=>a-b);
    const now=new Date(); const {start,end}=getWeekRange(now);
    const upcomingDates = dates.filter(d=>d>=start && d<=end);
    const {start:lsStart,end:lsEnd}=getWeekRange(new Date(start.getTime()-7*24*60*60*1000));
    const lastDates = dates.filter(d=>d>=lsStart && d<=lsEnd);
    const renderBlock=(dateObjs,headEl,bodyEl,msg)=>{
      headEl.innerHTML="<tr><th>Date</th><th>Song</th><th>Topic</th></tr>";
      bodyEl.innerHTML="";
      if(!dateObjs.length){ bodyEl.innerHTML=`<tr><td colspan='3'>${msg}</td></tr>`; return; }
      dateObjs.forEach(d=>{ const k=d.toISOString().slice(0,10); const list=dedupeByTitle(byDate.get(k)||[]);
        list.forEach(({date,song,topic})=>{ const tr=document.createElement("tr"); tr.innerHTML=`<td>${fmtDateOnly(date)}</td><td>${song}</td><td>${topic}</td>`; bodyEl.appendChild(tr); });
      });
    };
    renderBlock(upcomingDates, upHead, upBody, "No songs listed for this week.");
    renderBlock(lastDates, lsHead, lsBody, "No songs found for last week.");
  }catch(e){ console.error("Setlist error:", e); }
}

// ====== Analytics (52 weeks + all-time) ======
function loadGoogle(){ return new Promise(res=>{ google.charts.load("current",{packages:["corechart"]}); google.charts.setOnLoadCallback(res); }); }
function isExcludedSong(name){ const s=String(name||"").trim().toLowerCase(); return !s || s==="na" || s==="n/a" || s.includes("church close"); }
function inLastWeeks(d, w){ const today=new Date(); const start=new Date(today.getFullYear(),today.getMonth(),today.getDate()); start.setDate(start.getDate()-(w*7-1)); return d>=start && d<=today; }
async function computeCountsWindow(allRows, weeks){
  const rows=allRows.filter(r=>{ const m=normMap(r); const d=excelToDate(val(m, ["date","day","servicedate"])); return d && (weeks>=9999 || inLastWeeks(d,weeks)); });
  const byDate=new Map();
  rows.forEach(r=>{ const m=normMap(r); const d=excelToDate(val(m, ["date","day","servicedate"])); const t=String(val(m, ["song","title"])).trim(); if(!d||isExcludedSong(t)) return; const k=d.toISOString().slice(0,10); if(!byDate.has(k)) byDate.set(k,new Set()); byDate.get(k).add(t.toLowerCase()); });
  const counts=new Map(); byDate.forEach(set=>set.forEach(t=>counts.set(t,(counts.get(t)||0)+1)));
  const titleCase=new Map(); rows.forEach(r=>{ const m=normMap(r); const t=String(val(m, ["song","title"])).trim(); if(!isExcludedSong(t)){ const k=t.toLowerCase(); if(!titleCase.has(k)) titleCase.set(k,t); } });
  return {counts,titleCase};
}
function drawAnalytics(dataArray, chartId, tableId){
  const colors=["#1f77b4","#ff7f0e","#2ca02c","#d62728","#9467bd","#8c564b","#e377c2","#7f7f7f","#bcbd22","#17becf"];
  const el=document.getElementById(chartId);
  if(el){ if(!dataArray.length) el.innerHTML="No data."; else{ const data=google.visualization.arrayToDataTable([["Song","Plays"],...dataArray]); const opts={is3D:true,backgroundColor:"transparent",legend:"none",colors,chartArea:{width:"95%",height:"88%"}}; new google.visualization.PieChart(el).draw(data,opts); } }
  const tbody=document.getElementById(tableId);
  if(tbody){ tbody.innerHTML=""; if(!dataArray.length) tbody.innerHTML="<tr><td colspan='2'>No data found.</td></tr>"; else dataArray.forEach(([name,plays],i)=>{ const color=colors[i%colors.length]; const tr=document.createElement("tr"); tr.innerHTML=`<td><span class='dot' style='background:${color}'></span>${name}</td><td>${plays}</td>`; tbody.appendChild(tr); }); }
}
async function renderAnalytics(){
  await loadGoogle();
  const slRows = await fetchXlsxRows(SRC.setlist);
  const w = await computeCountsWindow(slRows, 52);
  const a = await computeCountsWindow(slRows, 9999);
  const topW = Array.from(w.counts.entries()).map(([k,v])=>[w.titleCase.get(k)||k,v]).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0])).slice(0,10);
  const topA = Array.from(a.counts.entries()).map(([k,v])=>[a.titleCase.get(k)||k,v]).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0])).slice(0,10);
  drawAnalytics(topW, "chart-top10-played", "table-top10-window");
  drawAnalytics(topA, "chart-top10-alltime", "table-top10-alltime");
}

// ====== Init ======
document.addEventListener("DOMContentLoaded", async () => {
  try {
    await renderWeeklyPractices();
    await Promise.all([
      renderAdditionalPractice(),
      renderTraining(),
      renderMembers(),
      renderAnnouncements(),
      renderBibleStudy(),
      renderSetlist()
    ]);
    await renderAnalytics();
  } catch(e) { console.error("Init error:", e); }
});
