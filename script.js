/***** CONFIG: GitHub sources (blob URLs) *****/
const SRC = {
  announcements: "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx",
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

/***** TIME NORMALIZATION *****/
function normalizeTimeRange(str) {
  if (!str) return "";
  let s = String(str).toLowerCase();
  s = s.replace(/[–—]/g, "-");
  s = s.replace(/\s*to\s*/g, "-");
  s = s.replace(/\s+/g, "");
  return s;     // e.g., "6:00pm-8:00pm"
}

/***** WEEKLY PRACTICES (rolling; now renders as table) *****/
const WEEKLY_RANGES_BY_WEEKDAY = { 4: "6:00pm-8:00pm", 0: "8:40am-9:30am" };

function toHM(h, m) {
  return new Date(0,0,0,h,m).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }).toLowerCase();
}
function nextWeeklyOccurrenceParts(targetWeekday, startH, startM, endH, endM) {
  const now = new Date();
  const occ = new Date(now);
  const delta = (targetWeekday - now.getDay() + 7) % 7;
  occ.setDate(now.getDate() + delta);
  occ.setHours(startH, startM, 0, 0);
  const end = new Date(occ); end.setHours(endH, endM, 0, 0);
  if (now > end) occ.setDate(occ.getDate() + 7);
  return { date: fmtDateOnly(occ), time: `${toHM(startH,startM)} - ${toHM(endH,endM)}` };
}
function renderWeeklyPractices() {
  const tbody = document.getElementById("weekly-practice-body");
  if (!tbody) return;
  const items = [
    nextWeeklyOccurrenceParts(4, 18, 0, 20, 0),  // Thurs 6–8pm
    nextWeeklyOccurrenceParts(0, 8, 40, 9, 30),  // Sun 8:40–9:30am
  ];
  tbody.innerHTML = "";
  items.forEach(({date,time}) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${date}</td><td>${time}</td>`;
    tbody.appendChild(tr);
  });
}

/***** ADDITIONAL PRACTICE (exclude Thu/Sun regular times) *****/
function detectTimeInRow(r) {
  for (const [k,v] of Object.entries(r)) {
    const nk = norm(k);
    if ((nk.includes("additional") || nk.includes("time")) && String(v).trim()) return String(v).trim();
  }
  const timeLike = Object.values(r).map(v => String(v||""))
    .find(s => /\d{1,2}:\d{2}\s*(am|pm)?\s*(?:-|to|–|—)\s*\d{1,2}:\d{2}\s*(am|pm)?/i.test(s));
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

      const rawTime = detectTimeInRow(r);
      const normalized = normalizeTimeRange(rawTime);
      const weekday = date.getDay();
      const weeklyRange = WEEKLY_RANGES_BY_WEEKDAY[weekday];
      if (weeklyRange && normalized && normalized === weeklyRange) return;

      out.push({ date, time: rawTime });
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

/***** TRAINING (Date · Time · Passage · Bible Verse; last 52 weeks) *****/
function detectVerse(r) {
  for (const k of Object.keys(r)) {
    const nk = norm(k);
    if (nk.includes("verse")) {
      const v = String(r[k] ?? "").trim();
      if (v) return v;
    }
  }
  for (const v of Object.values(r)) {
    const s = String(v ?? "").trim();
    if (/\b([1-3]?\s?[A-Za-z]+)\s+\d{1,3}:\d{1,3}\b/.test(s)) return s;
  }
  return "";
}
function detectPassage(r) {
  const candidates = ["passage","topic","study","scripture","reading","reference"];
  for (const k of Object.keys(r)) {
    const nk = norm(k);
    if (candidates.some(c => nk.includes(c))) {
      const v = String(r[k] ?? "").trim();
      if (v) return v;
    }
  }
  return "";
}
function inLastWeeks(d, weeks) {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  start.setDate(start.getDate() - (weeks * 7 - 1));
  return d >= start && d <= today;
}
async function renderTraining() {
  const tbody = document.getElementById("training-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.training);
    const items = rows.map(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      const date  = excelToDate(m.date ?? m.day ?? m.trainingdate ?? m.sessiondate);
      const time  = detectTimeInRow(r);
      const pass  = detectPassage(r);
      const verse = detectVerse(r);
      return { date, time, pass, verse };
    }).filter(x => x.date && inLastWeeks(x.date, 52))
      .sort((a,b)=>a.date-b.date);

    tbody.innerHTML = "";
    if (!items.length) { tbody.innerHTML = `<tr><td colspan="4">No trainings in the last 52 weeks.</td></tr>`; return; }
    items.forEach(({date,time,pass,verse}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${time || "-"}</td><td>${pass || "-"}</td><td>${verse || "-"}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Training error:", e);
    tbody.innerHTML = `<tr><td colspan="4">Unable to load training.</td></tr>`;
  }
}

/***** MEMBERS — THREE CATEGORIES *****/
async function renderMembers() {
  const leaderUl   = document.getElementById("leader-list");
  const musicianUl = document.getElementById("musician-list");
  const singersUl  = document.getElementById("singers-list");
  if (!leaderUl || !musicianUl || !singersUl) return;

  try {
    const rows = await fetchXlsxRows(SRC.members);
    const leader = [], musicians = [], singers = [];

    rows.forEach(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      const name = String(m.name ?? m.member ?? m.fullname ?? "").trim();
      const role = String(m.role ?? m.position ?? "").trim();
      const rlow = role.toLowerCase();

      if (!name && !role) return;
      if (rlow.includes("leader"))        leader.push(name || role);
      else if (rlow.includes("singer"))   singers.push(name || role);
      else                                musicians.push(name ? `${name} – ${role || "Musician"}` : role);
    });

    leader.sort((a,b)=>a.localeCompare(b));
    musicians.sort((a,b)=>a.localeCompare(b));
    singers.sort((a,b)=>a.localeCompare(b));

    const fill = (ul, list, emptyText="-") => {
      ul.innerHTML = "";
      if (!list.length) { ul.innerHTML = `<li>${emptyText}</li>`; return; }
      list.forEach(t => { const li=document.createElement("li"); li.textContent=t; ul.appendChild(li); });
    };
    fill(leaderUl, leader);
    fill(musicianUl, musicians);
    fill(singersUl, singers);
  } catch (e) {
    console.error("Members error:", e);
    leaderUl.innerHTML   = `<li>Unable to load members.</li>`;
    musicianUl.innerHTML = `<li>Unable to load members.</li>`;
    singersUl.innerHTML  = `<li>Unable to load members.</li>`;
  }
}

/***** ANNOUNCEMENTS *****/
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
      .filter(x => { const now=new Date(); const start = new Date(now); start.setDate(now.getDate()-31); return x.date>=start && x.date<=now; })
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

/***** SETLIST (between Bible Study and Analytics) *****/
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
function sundayCutoff(now = new Date()) {
  const d = new Date(now);
  const diff = (0 - d.getDay() + 7) % 7;
  const thisSunday = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff, 12, 30, 0, 0);
  if (d.getDay() === 0) thisSunday.setDate(d.getDate());
  return thisSunday;
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
    const todayISO = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const cutoff = sundayCutoff(now);

    const firstOnOrAfter = (target) => dates.find(d => d >= target);
    const firstAfter     = (target) => dates.find(d => d > target);
    const lastBefore     = (target) => [...dates].filter(d => d < target).pop();
    const lastOnOrBefore = (target) => [...dates].filter(d => d <= target).pop();

    let upDate, lastDate;
    if (now < cutoff) {
      upDate   = firstOnOrAfter(todayISO);
      lastDate = lastBefore(todayISO);
    } else {
      upDate   = firstAfter(todayISO);
      lastDate = lastOnOrBefore(todayISO);
    }

    const renderBlock = (dateObj, headEl, bodyEl, emptyMsg) => {
      headEl.innerHTML = `<tr><th>Date</th><th>Song</th><th>Topic</th></tr>`;
      bodyEl.innerHTML = "";
      if (!dateObj) { bodyEl.innerHTML = `<tr><td colspan="3">${emptyMsg}</td></tr>`; return; }
      const key = dateObj.toISOString().slice(0,10);
      const list = dedupeByTitle(byDate.get(key) || []);
      if (!list.length) { bodyEl.innerHTML = `<tr><td colspan="3">${emptyMsg}</td></tr>`; return; }
      list.forEach(({date, song, topic}) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${song}</td><td>${topic}</td>`;
        bodyEl.appendChild(tr);
      });
    };

    renderBlock(upDate,   upHead, upBody, "No songs listed yet.");
    renderBlock(lastDate, lsHead, lsBody, "No songs found for last week.");
  } catch (e) {
    console.error("Setlist error:", e);
    upHead.innerHTML = "<tr><th>Error</th></tr>";
    upBody.innerHTML = "<tr><td>Unable to load setlist.</td></tr>";
    lsHead.innerHTML = "<tr><th>Error</th></tr>";
    lsBody.innerHTML = "<tr><td>Unable to load setlist.</td></tr>";
  }
}

/***** ANALYTICS — last 52 weeks *****/
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
    return date && inLastWeeks(date, weeksWindow);
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

  const topWindow = displayCounts
    .sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0]))
    .slice(0,10);

  const colors = ["#1f77b4","#ff7f0e","#2ca02c","#d62728","#9467bd","#8c564b","#e377c2","#7f7f7f","#bcbd22","#17becf"];

  const elPie = document.getElementById("chart-top10-played");
  if (elPie) {
    if (!topWindow.length) {
      elPie.innerHTML = "No songs in the last 52 weeks.";
    } else {
      const data = google.visualization.arrayToDataTable([["Song","Plays"], ...topWindow]);
      const options = {
        is3D: true,
        backgroundColor: "transparent",
        legend: "none",
        colors,
        chartArea: { width: "95%", height: "88%" }
      };
      const chart = new google.visualization.PieChart(elPie);
      chart.draw(data, options);
    }
  }

  const tbody = document.getElementById("table-top10-window");
  if (tbody) {
    tbody.innerHTML = "";
    if (!topWindow.length) {
      tbody.innerHTML = `<tr><td colspan="2">No songs in the last 52 weeks.</td></tr>`;
    } else {
      topWindow.forEach(([name, plays], idx) => {
        const color = colors[idx % colors.length];
        const tr = document.createElement("tr");
        tr.innerHTML = `<td><span class="dot" style="background:${color}"></span>${name}</td><td>${plays}</td>`;
        tbody.appendChild(tr);
      });
    }
  }
}

/***** BIBLE STUDY (rolling last 4 weeks) *****/
async function renderBibleStudy() {
  const tbody = document.getElementById("bible-study-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.bibleStudy);
    const items = rows.map(r => {
      const m = {}; Object.keys(r).forEach(k => m[norm(k)] = r[k]);
      const date   = excelToDate(m.date ?? m.studydate ?? m.sessiondate ?? m.day);
      const topic  = String(m.topic ?? m.passage ?? m.study ?? m.series ?? "").trim();
      const verse  = String(m.verse ?? m["bibleverse"] ?? m["bible verse"] ?? "").trim();
      return { date, topic, verse };
    }).filter(x => x.date)
      .sort((a,b) => (b.date?.getTime()||0) - (a.date?.getTime()||0));

    const filtered = items.filter(x => x.date <= new Date() && inLastWeeks(x.date, 4)).slice(0,4);

    tbody.innerHTML = "";
    if (!filtered.length) { tbody.innerHTML = `<tr><td colspan="3">No entries for the last 4 weeks.</td></tr>`; return; }
    filtered.forEach(({date, topic, verse}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${topic || "-"}</td><td>${verse || "-"}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Bible Study error:", e);
    tbody.innerHTML = `<tr><td colspan="3">Unable to load bible study log.</td></tr>`;
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
  } catch (e) {
    console.error("Init error:", e);
  }
});
