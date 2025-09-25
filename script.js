/***** CONFIG: GitHub sources (blob URLs) *****/
const SRC = {
  announcements: "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx",
  members:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/members/members.xlsx",
  setlist:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/setlist.xlsx",
  catalog:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/songs_catalog.csv",
  kpiTop10:     "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/data/kpi_top10.csv",
  addPractice:  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/special_practice.xlsx",
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

// Excel/str/serial → Date (local)
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

// Formatting
const DAY_ABBR = ["Sun","Mon","Tues","Wed","Thurs","Fri","Sat"];
const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
function fmtDateOnly(dt) {
  return `${DAY_ABBR[dt.getDay()]}, ${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;
}
function withinLastNDays(dt, n = 31) {
  const today = new Date(); const start = new Date(); start.setDate(today.getDate() - n);
  return dt >= start && dt <= today;
}

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

/***** ADDITIONAL PRACTICE (non-Thu/Sun; reads "ADDITIONAL PRACTICE") *****/
async function renderAdditionalPractice() {
  const tbody = document.getElementById("additional-practice-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.addPractice);
    const normalized = rows.map((r) => {
      const o = {}; Object.keys(r).forEach(k => o[k.trim().toLowerCase()] = r[k]);
      const date = excelToDate(o.date ?? o["practice date"] ?? o.day ?? "");
      // robust: accept multiple header variants for the time column
      const time = String(
        o["additional practice"] ?? o["additionalpractice"] ??
        o["time"] ?? o["practice time"] ?? o["notes"] ?? ""
      ).trim();
      return { date, time };
    }).filter(x => x.date && ![0,4].includes(x.date.getDay())) // exclude Sun(0), Thu(4)
      .sort((a,b) => a.date - b.date);

    tbody.innerHTML = "";
    if (!normalized.length) { tbody.innerHTML = `<tr><td colspan="2">No additional practices listed.</td></tr>`; return; }
    normalized.forEach(({date,time}) => {
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
      const o = {}; Object.keys(r).forEach(k => o[k.trim().toLowerCase()] = r[k]);
      return {
        name: String(o.name ?? o.member ?? o["full name"] ?? "").trim(),
        role: String(o.role ?? o.position ?? "").trim()
      };
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
    const data = rows.map((r) => {
      const o = {}; Object.keys(r).forEach(k => o[k.trim().toLowerCase()] = r[k]);
      const date = excelToDate(o.date ?? o["announcement date"] ?? "");
      const english = String(o.announcement ?? o.english ?? o.message ?? "").trim();
      const hmong   = String(o["lus tshaj tawm"] ?? o.hmong ?? "").trim();
      return { date, english, hmong };
    }).filter(x => x.date && (x.english || x.hmong))
      .filter(x => withinLastNDays(x.date, 31))
      .sort((a,b) => b.date - a.date);

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

/***** SETLIST (only Date, Song, Topic; split upcoming/last) *****/
function normalizeSetlistRow(r) {
  const o = {}; Object.keys(r).forEach(k => o[k.trim().toLowerCase()] = r[k]);
  const date = excelToDate(o.date ?? o.day ?? o["service date"] ?? "");
  const song = String(o.song ?? o.title ?? "").trim();
  const topic = String(o.topic ?? o.notes ?? "").trim();
  return { date, song, topic };
}
async function renderSetlist() {
  const upHead  = document.getElementById("setlist-up-head");
  const upBody  = document.getElementById("setlist-up-body");
  const lsHead  = document.getElementById("setlist-last-head");
  const lsBody  = document.getElementById("setlist-last-body");
  if (!upHead || !upBody || !lsHead || !lsBody) return;

  try {
    const rows = (await fetchXlsxRows(SRC.setlist)).map(normalizeSetlistRow).filter(x => x.date && x.song);

    // Group by date (YYYY-MM-DD)
    const byDate = new Map();
    for (const r of rows) {
      const key = r.date.toISOString().slice(0,10);
      if (!byDate.has(key)) byDate.set(key, []);
      byDate.get(key).push(r);
    }
    const dates = Array.from(byDate.keys()).map(d => new Date(d)).sort((a,b)=>a-b);
    const today = new Date();
    const todayDateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const upcomingDate = dates.find(d => d >= todayDateOnly);
    const lastDate = [...dates].filter(d => d < todayDateOnly).pop();

    function renderBlock(dateObj, headEl, bodyEl) {
      headEl.innerHTML = `<tr><th>Date</th><th>Song</th><th>Topic</th></tr>`;
      bodyEl.innerHTML = "";
      if (!dateObj) { bodyEl.innerHTML = `<tr><td colspan="3">No data.</td></tr>`; return; }
      const key = dateObj.toISOString().slice(0,10);
      const list = byDate.get(key) || [];
      if (!list.length) { bodyEl.innerHTML = `<tr><td colspan="3">No songs for this date.</td></tr>`; return; }
      list.forEach(({date, song, topic}) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${song}</td><td>${topic}</td>`;
        bodyEl.appendChild(tr);
      });
    }

    renderBlock(upcomingDate, upHead, upBody);
    renderBlock(lastDate, lsHead, lsBody);
  } catch (e) {
    console.error("Setlist error:", e);
    upHead.innerHTML = "<tr><th>Error</th></tr>";
    upBody.innerHTML = "<tr><td>Unable to load setlist.</td></tr>";
    lsHead.innerHTML = "<tr><th>Error</th></tr>";
    lsBody.innerHTML = "<tr><td>Unable to load setlist.</td></tr>";
  }
}

/***** ANALYTICS *****/
function loadGoogle() {
  return new Promise((resolve) => { google.charts.load("current", { packages: ["corechart"] }); google.charts.setOnLoadCallback(resolve); });
}

// Top 10 (3D pie)
async function renderTop10Pie() {
  const el = document.getElementById("chart-top10");
  if (!el) return;
  try {
    await loadGoogle();
    const rows = await fetchCsv(SRC.kpiTop10);
    const headers = rows.length ? Object.keys(rows[0]) : [];
    let nameKey = headers.find(h => /song|title/i.test(h)) || headers[0];
    let valKey  = headers.find(h => /play|count|times|value|freq|uses/i.test(h)) || headers[1];

    const dataArr = [["Song", "Plays"]];
    rows.forEach(r => {
      const name = String(r?.[nameKey] ?? "").trim();
      let valRaw = (r?.[valKey] ?? "0").toString().trim().replace(/,/g,"");
      const val = Number(valRaw);
      if (name && isFinite(val)) dataArr.push([name, val]);
    });

    if (dataArr.length <= 1) { el.innerHTML = "Unable to render Top 10."; return; }
    const data = google.visualization.arrayToDataTable(dataArr);
    const options = { is3D: true, backgroundColor: "transparent", legend: { textStyle: { color: "#e5e7eb" } }, chartArea: { width: "90%", height: "80%" } };
    const chart = new google.visualization.PieChart(el);
    chart.draw(data, options);
  } catch (e) {
    console.error("Top10 Pie error:", e);
    el.innerHTML = "Unable to render Top 10.";
  }
}

// Played vs Not Played — from songs_catalog.csv only
async function renderPlayedRatio() {
  const kpiVal = document.getElementById("kpi-played-ratio");
  const kpiSub = document.getElementById("kpi-played-sub");
  const el = document.getElementById("chart-played-ratio");
  if (!kpiVal || !kpiSub || !el) return;

  try {
    await loadGoogle();
    const rows = await fetchCsv(SRC.catalog);

    // Identify ID-like columns to ignore when deciding "played"
    const idLike = new Set(["number", "no", "index", "#"]);

    let played = 0, total = 0;
    rows.forEach(row => {
      total += 1;
      // "Played" if any non-ID column has a non-blank value
      const hasName = Object.entries(row).some(([k,v]) => {
        const key = String(k).trim().toLowerCase();
        if (idLike.has(key)) return false;
        const val = String(v ?? "").trim();
        return val.length > 0;
      });
      if (hasName) played += 1;
    });

    const notPlayed = Math.max(total - played, 0);
    const pct = total ? Math.round((played / total) * 100) : 0;

    kpiVal.textContent = `${pct}%`;
    kpiSub.textContent = `Played ${played} of ${total}`;

    const data = google.visualization.arrayToDataTable([
      ["Type","Count"],
      ["Played", played],
      ["Not Played", notPlayed],
    ]);
    const options = { is3D: true, backgroundColor: "transparent", legend: { textStyle: { color: "#e5e7eb" } }, chartArea: { width: "90%", height: "80%" } };
    const chart = new google.visualization.PieChart(el);
    chart.draw(data, options);
  } catch (e) {
    console.error("Played Ratio error:", e);
    kpiVal.textContent = "—";
    kpiSub.textContent = "Unable to compute";
    el.innerHTML = "Unable to render Played/Not Played.";
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
      renderSetlist()
    ]);
    await Promise.all([
      renderTop10Pie(),
      renderPlayedRatio()
    ]);
  } catch (e) {
    console.error("Init error:", e);
  }
});
