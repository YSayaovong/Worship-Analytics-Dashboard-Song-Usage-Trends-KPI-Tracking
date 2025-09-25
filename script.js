/***** CONFIG: GitHub sources (blob URLs) *****/
const SRC = {
  announcements: "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx",
  members:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/members/members.xlsx",
  setlist:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/setlist.xlsx",
  catalog:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/songs_catalog.csv",
  addPractice:  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/special_practice.xlsx",
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

async function fetchCsv(blobUrl) {
  const url = toRaw(blobUrl) + "?v=" + Date.now();
  const txt = await (await fetch(url)).text();
  return new Promise((resolve) => {
    Papa.parse(txt, { header: true, skipEmptyLines: true, complete: (r) => resolve(r.data) });
  });
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
function withinLastNDays(dt, n = 31) { const t = new Date(); const s = new Date(); s.setDate(t.getDate() - n); return dt >= s && dt <= t; }
const norm = (s) => String(s||"").toLowerCase().replace(/[^a-z0-9]+/g,"");

/***** WEEKLY PRACTICES (rolling) *****/
function nextWeeklyOccurrence(targetWeekday, startH, startM, endH, endM) {
  const now = new Date();
  const occ = new Date(now);
  const delta = (targetWeekday - now.getDay() + 7) % 7;
  occ.setDate(now.getDate() + delta);
  occ.setHours(startH, startM, 0, 0);
  const end = new Date(occ); end.setHours(endH, endM, 0, 0);
  if (now > end) occ.setDate(occ.getDate() + 7);
  const toHM = (h,m)=> new Date(0,0,0,h,m).toLocaleTimeString([], {hour:"numeric", minute:"2-digit"}).toLowerCase();
  return `${fmtDateOnly(occ)} ${toHM(startH,startM)} to ${toHM(endH,endM)}`;
}
function renderWeeklyPractices() {
  const ul = document.getElementById("weekly-practice-list");
  if (!ul) return;
  ul.innerHTML = "";
  [nextWeeklyOccurrence(4, 18, 0, 20, 0),  // Thurs 6–8pm
   nextWeeklyOccurrence(0, 8, 40, 9, 30)]  // Sun 8:40–9:30am
   .forEach(t => { const li=document.createElement("li"); li.textContent=t; ul.appendChild(li); });
}

/***** ADDITIONAL PRACTICE (Date + Time from special_practice.xlsx) *****/
function detectTimeInRow(r) {
  for (const [k,v] of Object.entries(r)) {
    const nk = norm(k);
    if ((nk.includes("additional") || nk.includes("time")) && String(v).trim()) return String(v).trim();
  }
  const timeLike = Object.values(r).map(v => String(v||"")).find(s => /\d{1,2}:\d{2}\s*(am|pm)?\s*-\s*\d{1,2}:\d{2}\s*(am|pm)?/i.test(s));
  if (timeLike) return timeLike.trim();
  const vals = Object.values(r);
  if (vals.length === 2) {
    const [a,b] = vals;
    const aDate = excelToDate(a); const bDate = excelToDate(b);
    if (aDate && !bDate) return String(b||"").trim();
    if (bDate && !aDate) return String(a||"").trim();
  }
  return "";
}
async function renderAdditionalPractice() {
  const tbody = document.getElementById("additional-practice-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.addPractice);
    const out = [];
    rows.forEach(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      const date = excelToDate(m.date ?? m.practicedate ?? m.day ?? m.practice ?? m.practice_dt);
      if (!date) return;
      if ([0,4].includes(date.getDay())) return;
      const time = detectTimeInRow(r);
      out.push({ date, time });
    });

    out.sort((a,b)=>a.date-b.date);
    tbody.innerHTML = "";
    if (!out.length) { tbody.innerHTML = `<tr><td colspan="2">No additional practices listed.</td></tr>`; return; }
    out.forEach(({date,time}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${time || "-"}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Additional Practice error:", e);
    tbody.innerHTML = `<tr><td colspan="2">Unable to load Additional Practice.</td></tr>`;
  }
}

/***** MEMBERS *****/
async function renderMembers() {
  const tbody = document.getElementById("members-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.members);
    const mapped = rows.map(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      return { name: String(m.name ?? m.member ?? m.fullname ?? "").trim(),
               role: String(m.role ?? m.position ?? "").trim() };
    }).filter(x => x.name || x.role);

    tbody.innerHTML = "";
    if (!mapped.length) { tbody.innerHTML = `<tr><td colspan="2">No members listed.</td></tr>`; return; }
    mapped.forEach(({name, role}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${name}</td><td>${role}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Members error:", e);
    tbody.innerHTML = `<tr><td colspan="2">Unable to load members.</td></tr>`;
  }
}

/***** ANNOUNCEMENTS (Date | English | Hmong; newest; 31 days) *****/
async function renderAnnouncements() {
  const tbody = document.getElementById("announcements-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.announcements);
    const data = rows.map(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      const date = excelToDate(m.date ?? m.announcementdate);
      const english = String(m.announcement ?? m.english ?? m.message ?? "").trim();
      const hmong   = String(m.lustshajtawm ?? m.hmong ?? "").trim();
      return { date, english, hmong };
    }).filter(x => x.date && (x.english || x.hmong))
      .filter(x => withinLastNDays(x.date, 31))
      .sort((a,b)=>b.date-a.date);

    tbody.innerHTML = "";
    if (!data.length) { tbody.innerHTML = `<tr><td colspan="3">No announcements in the last 31 days.</td></tr>`; return; }
    data.forEach(({date, english, hmong}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${english}</td><td>${hmong}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Announcements error:", e);
    tbody.innerHTML = `<tr><td colspan="3">Unable to load announcements.</td></tr>`;
  }
}

/***** SETLIST (Date, Song, Topic; split upcoming/last; no repeated titles per date) *****/
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
    const today = new Date();
    const todayOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const upcomingDate = dates.find(d => d >= todayOnly);
    const lastDate = [...dates].filter(d => d < todayOnly).pop();

    const renderBlock = (dateObj, headEl, bodyEl) => {
      headEl.innerHTML = `<tr><th>Date</th><th>Song</th><th>Topic</th></tr>`;
      bodyEl.innerHTML = "";
      if (!dateObj) { bodyEl.innerHTML = `<tr><td colspan="3">No data.</td></tr>`; return; }
      const key = dateObj.toISOString().slice(0,10);
      const list = dedupeByTitle(byDate.get(key) || []);
      if (!list.length) { bodyEl.innerHTML = `<tr><td colspan="3">No songs for this date.</td></tr>`; return; }
      list.forEach(({date, song, topic}) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${song}</td><td>${topic}</td>`;
        bodyEl.appendChild(tr);
      });
    };

    renderBlock(upcomingDate, upHead, upBody);
    renderBlock(lastDate,  lsHead, lsBody);
  } catch (e) {
    console.error("Setlist error:", e);
    upHead.innerHTML = "<tr><th>Error</th></tr>";
    upBody.innerHTML = "<tr><td>Unable to load setlist.</td></tr>";
    lsHead.innerHTML = "<tr><th>Error</th></tr>";
    lsBody.innerHTML = "<tr><td>Unable to load setlist.</td></tr>";
  }
}

/***** ANALYTICS — Top 10 Played (all time) + Top 10 counts for current year *****/
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
async function computeSongCounts(allRows) {
  // Build counts with per-date dedupe
  const byDate = new Map();
  allRows.forEach(r => {
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
  allRows.forEach(r => {
    const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
    const t = String(m.song ?? m.title ?? "").trim();
    if (!isExcludedSong(t)) {
      const key = t.toLowerCase();
      if (!titleCase.has(key)) titleCase.set(key, t);
    }
  });

  return { counts, titleCase };
}
async function renderTopCharts() {
  await loadGoogle();
  const slRows = await fetchXlsxRows(SRC.setlist);

  // All-time counts
  const { counts, titleCase } = await computeSongCounts(slRows);
  const displayCounts = Array.from(counts.entries()).map(([k,v]) => [titleCase.get(k) || k, v]);
  const topPlayed = displayCounts
    .sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0]))
    .slice(0,10);

  // Draw all-time pie
  const el1 = document.getElementById("chart-top10-played");
  if (el1) {
    if (!topPlayed.length) { el1.innerHTML = "No data."; }
    else {
      const data = google.visualization.arrayToDataTable([["Song","Plays"], ...topPlayed]);
      const options = { is3D:true, backgroundColor:"transparent",
        legend:{ textStyle:{ color:"#e5e7eb" } }, chartArea:{ width:"90%", height:"80%" } };
      const chart = new google.visualization.PieChart(el1);
      chart.draw(data, options);
    }
  }

  // Current year counts
  const year = new Date().getFullYear();
  const rowsThisYear = slRows.filter(r => {
    const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
    const d = excelToDate(m.date ?? m.day ?? m.servicedate);
    return d && d.getFullYear() === year;
  });
  const { counts: countsYear, titleCase: titleCaseYear } = await computeSongCounts(rowsThisYear);
  const displayYear = Array.from(countsYear.entries())
    .map(([k,v]) => [titleCaseYear.get(k) || k, v])
    .sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0]))
    .slice(0,10);

  // Render table for this year
  const tbody = document.getElementById("table-top10-year");
  if (tbody) {
    tbody.innerHTML = "";
    if (!displayYear.length) {
      tbody.innerHTML = `<tr><td colspan="2">No songs played this year.</td></tr>`;
    } else {
      displayYear.forEach(([name, plays]) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${name}</td><td>${plays}</td>`;
        tbody.appendChild(tr);
      });
    }
  }
}

/***** BIBLE STUDY LOG *****/
async function renderBibleStudy() {
  const tbody = document.getElementById("bible-study-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.bibleStudy);
    const items = rows.map(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      const date   = excelToDate(m.date ?? m.studydate ?? m.sessiondate ?? m.day);
      const topic  = String(m.topic ?? m.passage ?? m.study ?? m.series ?? "").trim();
      const leader = String(m.leader ?? m.teacher ?? m.speaker ?? "").trim();
      const notes  = String(m.notes ?? m.note ?? "").trim();
      return { date, topic, leader, notes };
    }).filter(x => x.date || x.topic || x.leader || x.notes)
      .sort((a,b) => (b.date?.getTime()||0) - (a.date?.getTime()||0));

    tbody.innerHTML = "";
    if (!items.length) { tbody.innerHTML = `<tr><td colspan="4">No bible study entries found.</td></tr>`; return; }
    items.forEach(({date, topic, leader, notes}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${date ? fmtDateOnly(date) : "-"}</td><td>${topic||"-"}</td><td>${leader||"-"}</td><td>${notes||""}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Bible Study error:", e);
    tbody.innerHTML = `<tr><td colspan="4">Unable to load bible study log.</td></tr>`;
  }
}

/***** INIT *****/
document.addEventListener("DOMContentLoaded", async () => {
  try {
    renderWeeklyPractices();
    await Promise.all([
      renderAdditionalPractice(),
      renderMembers(),
      renderAnnouncements(),
      renderSetlist(),
      renderBibleStudy()
    ]);
    await renderTopCharts();
  } catch (e) {
    console.error("Init error:", e);
  }
});
