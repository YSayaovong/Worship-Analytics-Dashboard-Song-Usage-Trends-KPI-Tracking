 /***** CONFIG: GitHub sources *****/
const SRC = {
  // Updated announcements location (new repo)
  announcements: "https://github.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/blob/main/announcements/announcements.xlsx",
  // Existing sources (unchanged)
  members:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/members/members.xlsx",
  setlist:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/setlist.xlsx",
  catalog:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/songs_catalog.csv",
  addPractice:  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/special_practice.xlsx",
  training:     "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/training.xlsx",
  bibleStudy:   "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/bible_study/bible_study.xlsx",
};

/***** UTILITIES *****/
const toRaw = (blobUrl) =>
  blobUrl.replace("https://github.com/", "https://raw.githubusercontent.com/").replace("/blob/", "/");

async function fetchXlsxRows(blobUrl, sheetNameOrIndex = 0) {
  const url = toRaw(blobUrl) + "?v=" + Date.now();
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Fetch failed: ${url}`);
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const sheet =
    typeof sheetNameOrIndex === "number"
      ? wb.Sheets[wb.SheetNames[sheetNameOrIndex]]
      : wb.Sheets[sheetNameOrIndex] || wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function excelToDate(val) {
  if (val == null || val === "") return null;
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (!d) return null;
    return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

const DAY_ABBR = ["Sun","Mon","Tues","Wed","Thurs","Fri","Sat"];
const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
const fmtDateOnly = (dt) => `${DAY_ABBR[dt.getDay()]}, ${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;
const norm = (s) => String(s||"").toLowerCase().replace(/[^a-z0-9]+/g,"");

/***** WEEK RANGE HELPER *****/
function getWeekRange(date) {
  const start = new Date(date);
  start.setDate(start.getDate() - start.getDay());
  start.setHours(0,0,0,0);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  end.setHours(23,59,59,999);
  return { start, end };
}

/***** Helpers for robust column picking *****/
function normMap(row){
  const m = {}; Object.keys(row||{}).forEach(k => m[norm(k)] = row[k]); return m;
}
function val(m, keys){
  for(const k of keys){ const v = m[k]; if(v!=null && String(v)!=="") return v; }
  return "";
}
function findByIncludes(m, substrings){
  const keys = Object.keys(m);
  for(const key of keys){
    const k = key.toLowerCase();
    let ok = true;
    for(const sub of substrings){
      if(!k.includes(sub)) { ok = false; break; }
    }
    if(ok) return m[key];
  }
  return "";
}

/***** SETLIST (includes all events in the week) *****/
function normSetlistRow(r) {
  const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
  const date = excelToDate(m.date ?? m.day ?? m.servicedate);
  const song = String(m.song ?? m.title ?? "").trim();
  const topic = String(m.topic ?? m.notes ?? "").trim();
  return { date, song, topic };
}
function dedupeByTitle(list) {
  const seen = new Set();
  return list.filter(item => {
    const key = (item.song || "").toLowerCase();
    if (!key) return false;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
async function renderSetlist() {
  const upHead  = document.getElementById("setlist-up-head");
  const upBody  = document.getElementById("setlist-up-body");
  const lsHead  = document.getElementById("setlist-last-head");
  const lsBody  = document.getElementById("setlist-last-body");
  if (!upHead || !upBody || !lsHead || !lsBody) return;

  try {
    const rows = (await fetchXlsxRows(SRC.setlist)).map(normSetlistRow).filter(x => x.date && x.song);

    const byDate = new Map();
    for (const r of rows) {
      const key = r.date.toISOString().slice(0,10);
      if (!byDate.has(key)) byDate.set(key, []);
      byDate.get(key).push(r);
    }
    const dates = Array.from(byDate.keys()).map(d => new Date(d)).sort((a,b)=>a-b);

    const now = new Date();
    const { start, end } = getWeekRange(now);
    const upcomingDates = dates.filter(d => d >= start && d <= end);

    const { start: lastStart, end: lastEnd } = getWeekRange(new Date(start.getTime() - 7*24*60*60*1000));
    const lastDates = dates.filter(d => d >= lastStart && d <= lastEnd);

    const renderBlock = (dateObjs, headEl, bodyEl, emptyMsg) => {
      headEl.innerHTML = `<tr><th>Date</th><th>Song</th><th>Topic</th></tr>`;
      bodyEl.innerHTML = "";
      if (!dateObjs.length) { bodyEl.innerHTML = `<tr><td colspan="3">${emptyMsg}</td></tr>`; return; }
      dateObjs.forEach(dateObj => {
        const key = dateObj.toISOString().slice(0,10);
        const list = dedupeByTitle(byDate.get(key) || []);
        list.forEach(({date, song, topic}) => {
          const tr = document.createElement("tr");
          tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${song}</td><td>${topic}</td>`;
          bodyEl.appendChild(tr);
        });
      });
    };

    renderBlock(upcomingDates, upHead, upBody, "No songs listed for this week.");
    renderBlock(lastDates, lsHead, lsBody, "No songs found for last week.");
  } catch (e) {
    console.error("Setlist error:", e);
  }
}

/***** ANALYTICS *****/
function loadGoogle() {
  return new Promise((resolve) => { google.charts.load("current", { packages: ["corechart"] }); google.charts.setOnLoadCallback(resolve); });
}
function isExcludedSong(name) {
  const s = String(name||"").trim().toLowerCase();
  if (!s) return true;
  if (s === "na" || s === "n/a") return true;
  if (s.includes("church close")) return true;
  return false;
}
async function computeCountsWindow(allRows, weeksWindow) {
  const rows = allRows.filter(r => {
    const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
    const date = excelToDate(m.date ?? m.day ?? m.servicedate);
    return date && (weeksWindow >= 9999 || inLastWeeks(date, weeksWindow));
  });

  const byDate = new Map();
  rows.forEach(r => {
    const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
    const date = excelToDate(m.date ?? m.day ?? m.servicedate);
    const title = String(m.song ?? m.title ?? "").trim();
    if (!date || isExcludedSong(title)) return;
    const key = date.toISOString().slice(0,10);
    if (!byDate.has(key)) byDate.set(key, new Set());
    byDate.get(key).add(title.toLowerCase());
  });

  const counts = new Map();
  byDate.forEach(set => set.forEach(t => counts.set(t, (counts.get(t) || 0) + 1)));

  const titleCase = new Map();
  rows.forEach(r => {
    const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
    const t = String(m.song ?? m.title ?? "").trim();
    if (!isExcludedSong(t)) {
      const key = t.toLowerCase();
      if (!titleCase.has(key)) titleCase.set(key, t);
    }
  });

  return { counts, titleCase };
}
function inLastWeeks(d, weeks) {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  start.setDate(start.getDate() - (weeks * 7 - 1));
  return d >= start && d <= today;
}
async function renderAnalytics52Weeks() {
  await loadGoogle();
  const slRows = await fetchXlsxRows(SRC.setlist);
  const { counts, titleCase } = await computeCountsWindow(slRows, 52);
  const displayCounts = Array.from(counts.entries()).map(([k,v]) => [titleCase.get(k) || k, v]);
  const topWindow = displayCounts.sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0])).slice(0,10);
  drawAnalytics(topWindow, "chart-top10-played", "table-top10-window");
}
async function renderAnalyticsAllTime() {
  await loadGoogle();
  const slRows = await fetchXlsxRows(SRC.setlist);
  const { counts, titleCase } = await computeCountsWindow(slRows, 9999);
  const displayCounts = Array.from(counts.entries()).map(([k,v]) => [titleCase.get(k) || k, v]);
  const topAll = displayCounts.sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0])).slice(0,10);
  drawAnalytics(topAll, "chart-top10-alltime", "table-top10-alltime");
}
function drawAnalytics(dataArray, chartId, tableId) {
  const colors = ["#1f77b4","#ff7f0e","#2ca02c","#d62728","#9467bd","#8c564b","#e377c2","#7f7f7f","#bcbd22","#17becf"];
  const elPie = document.getElementById(chartId);
  if (elPie) {
    if (!dataArray.length) {
      elPie.innerHTML = "No data.";
    } else {
      const data = google.visualization.arrayToDataTable([["Song","Plays"], ...dataArray]);
      const options = { is3D: true, backgroundColor: "transparent", legend: "none", colors, chartArea: { width: "95%", height: "88%" } };
      const chart = new google.visualization.PieChart(elPie);
      chart.draw(data, options);
    }
  }
  const tbody = document.getElementById(tableId);
  if (tbody) {
    tbody.innerHTML = "";
    if (!dataArray.length) {
      tbody.innerHTML = `<tr><td colspan="2">No data found.</td></tr>`;
    } else {
      dataArray.forEach(([name, plays], idx) => {
        const color = colors[idx % colors.length];
        const tr = document.createElement("tr");
        tr.innerHTML = `<td><span class="dot" style="background:${color}"></span>${name}</td><td>${plays}</td>`;
        tbody.appendChild(tr);
      });
    }
  }
}

/***** INIT *****/
document.addEventListener("DOMContentLoaded", async () => {
  try {
    renderWeeklyPractices();
    await Promise.all([
      renderAdditionalPractice(),
      renderTraining(),
      renderMembers(),
      renderAnnouncements(),
      renderBibleStudy(),
      renderSetlist()
    ]);
    await renderAnalytics52Weeks();
    await renderAnalyticsAllTime();
  } catch (e) {
    console.error("Init error:", e);
  }
});

/***** ---------- SAFE RENDERERS ---------- *****/

/***** Weekly Practices (show same-day until 11:59 PM, no premature rollover) *****/
async function renderWeeklyPractices(){
  const tbody = document.getElementById("weekly-practice-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  const now = new Date();
  // Return the next occurrence of target DOW, allowing same-day (delta can be 0)
  function nextDow(target){ // 0=Sun ... 6=Sat
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
    const delta = (target - d.getDay() + 7) % 7; // <-- no "|| 7" so same-day stays
    d.setDate(d.getDate() + delta);
    d.setHours(0,0,0,0);
    return d;
  }
  const thurs = nextDow(4);
  const sun = nextDow(0);
  const fmt = (dt) => `${DAY_ABBR[dt.getDay()]}, ${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;
  const rows = [
    { date: thurs, time: "7:00 PM" },
    { date: sun,   time: "12:30 PM" },
  ];
  rows.forEach(r => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${fmt(r.date)}</td><td>${r.time}</td>`;
    tbody.appendChild(tr);
  });
}

/***** Additional Practice *****/
async function renderAdditionalPractice(){
  const tbody = document.getElementById("additional-practice-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(SRC.addPractice);
    rows.forEach(r => {
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day","servicedate"]));
      const t = val(m, ["time","starttime","practice"]);
      if(!d || !t) return;
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(d)}</td><td>${t}</td>`;
      tbody.appendChild(tr);
    });
    if(!tbody.children.length){
      tbody.innerHTML = `<tr><td colspan="2">No additional practices listed.</td></tr>`;
    }
  }catch(e){
    console.error("Additional practice error:", e);
    tbody.innerHTML = `<tr><td colspan="2">Could not load special practice sheet.</td></tr>`;
  }
}

/***** Training (already positioned after Members in HTML) *****/
async function renderTraining(){
  const tbody = document.getElementById("training-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(SRC.training);
    rows.forEach(r => {
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day"]));
      const t = val(m, ["time","starttime"]);
      const passage = val(m, ["passage","topic","study"]);
      const verse = val(m, ["bibleverse","verse","reference"]);
      if(!d || (!t && !passage && !verse)) return;
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(d)}</td><td>${t||""}</td><td>${passage||""}</td><td>${verse||""}</td>`;
      tbody.appendChild(tr);
    });
    if(!tbody.children.length){
      tbody.innerHTML = `<tr><td colspan="4">No training entries found.</td></tr>`;
    }
  }catch(e){
    console.error("Training error:", e);
    tbody.innerHTML = `<tr><td colspan="4">Could not load training sheet.</td></tr>`;
  }
}

/***** Members *****/
async function renderMembers(){
  const leaderList = document.getElementById("leader-list");
  const musicianList = document.getElementById("musician-list");
  const singersList = document.getElementById("singers-list");
  if(!leaderList || !musicianList || !singersList) return;
  leaderList.innerHTML = musicianList.innerHTML = singersList.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(SRC.members);
    const leaders = [], musicians = [], singers = [];
    rows.forEach(r => {
      const m = normMap(r);
      const name = val(m, ["name","member","person"]) || "";
      const role = (val(m, ["role","position","type"]) || "").toLowerCase();
      if(!name) return;
      if(role.includes("leader")) leaders.push(name);
      else if(role.includes("singer") || role.includes("vocal")) singers.push(name);
      else musicians.push(name);
    });
    const addAll = (ul, arr) => {
      if(!arr.length){ ul.innerHTML = "<li class='muted'>None listed</li>"; return; }
      arr.forEach(n => { const li = document.createElement("li"); li.textContent = n; ul.appendChild(li); });
    };
    addAll(leaderList, leaders);
    addAll(musicianList, musicians);
    addAll(singersList, singers);
  }catch(e){
    console.error("Members error:", e);
    leaderList.innerHTML = musicianList.innerHTML = singersList.innerHTML = "<li class='muted'>Could not load members sheet.</li>";
  }
}

/***** Announcements (improved Hmong detection) *****/
async function renderAnnouncements(){
  const tbody = document.getElementById("announcements-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(SRC.announcements);
    const today = new Date();
    const thirtyOneDays = 31 * 24 * 60 * 60 * 1000;

    const items = rows.map(r => {
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day"]));
      const en = val(m, ["announcementenglish","announcement","english"]);
      // robust hmong: try known keys; otherwise find any key containing both 'hmong' or ('lus' + 'tshaj')
      let hm = val(m, ["hmong","lustshajtawm","lus_tshaj_tawm","lus","tshaj"]);
      if(!hm){
        // Try by includes
        hm = findByIncludes(m, ["hmong"]) || findByIncludes(m, ["lus","tshaj"]);
      }
      return { d, en, hm };
    }).filter(x => x.d && (today - x.d) <= thirtyOneDays)
      .sort((a,b) => b.d - a.d);

    items.forEach(it => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(it.d)}</td><td>${it.en||""}</td><td>${it.hm||""}</td>`;
      tbody.appendChild(tr);
    });
    if(!tbody.children.length){
      tbody.innerHTML = `<tr><td colspan="3">No announcements from the last 31 days.</td></tr>`;
    }
  }catch(e){
    console.error("Announcements error:", e);
    tbody.innerHTML = `<tr><td colspan="3">Could not load announcements sheet.</td></tr>`;
  }
}

/***** Bible Study *****/
async function renderBibleStudy(){
  const tbody = document.getElementById("bible-study-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(SRC.bibleStudy);
    const { start, end } = getWeekRange(new Date());
    const prev3 = new Date(start); prev3.setDate(prev3.getDate() - 21);
    const items = rows.map(r => {
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day"]));
      const topic = val(m, ["topic","passage","study"]);
      const verse = val(m, ["bibleverse","verse","reference"]);
      return { d, topic, verse };
    }).filter(x => x.d && x.d >= prev3 && x.d <= end)
      .sort((a,b) => b.d - a.d);
    items.forEach(it => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(it.d)}</td><td>${it.topic||""}</td><td>${it.verse||""}</td>`;
      tbody.appendChild(tr);
    });
    if(!tbody.children.length){
      tbody.innerHTML = `<tr><td colspan="3">No bible study entries for the last 3 weeks.</td></tr>`;
    }
  }catch(e){
    console.error("Bible study error:", e);
    tbody.innerHTML = `<tr><td colspan="3">Could not load bible study sheet.</td></tr>`;
  }
}
