/***** CONFIG: GitHub sources (blob URLs) *****/
const SRC = {
  announcements: "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx",
  bibleStudy:   "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/bible_study/bible_study.xlsx",
  hymnalUnused: "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/data/hymnal_unused.csv",
  kpiBySource:  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/data/kpi_by_source.csv",
  kpiCoverage:  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/data/kpi_hymnal_coverage.csv",
  kpiTop10:     "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/data/kpi_top10.csv",
  members:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/members/members.xlsx",
  setlist:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/setlist.xlsx",
  catalog:      "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/setlist/songs_catalog.csv",
  addPractice:  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/special_practice.xlsx",
};

/***** UTILITIES *****/
const toRaw = (blobUrl) =>
  blobUrl.replace("https://github.com/", "https://raw.githubusercontent.com/").replace("/blob/", "/");

async function fetchXlsxRows(blobUrl, sheetNameOrIndex = 0) {
  const url = toRaw(blobUrl) + "?v=" + Date.now(); // cache-bust
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
  const url = toRaw(blobUrl) + "?v=" + Date.now(); // cache-bust
  const txt = await (await fetch(url)).text();
  return new Promise((resolve) => {
    Papa.parse(txt, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data),
    });
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

// Date formatting
const DAY_ABBR = ["Sun","Mon","Tues","Wed","Thurs","Fri","Sat"];
const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
function fmtDateOnly(dt) {
  return `${DAY_ABBR[dt.getDay()]}, ${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;
}
function fmtTime(h, m) {
  const d = new Date(); d.setHours(h, m, 0, 0);
  return d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }).toLowerCase();
}
function fmtDateRange(dt, startH, startM, endH, endM) {
  return `${fmtDateOnly(dt)} ${fmtTime(startH, startM)} to ${fmtTime(endH, endM)}`;
}

function withinLastNDays(dt, n = 31) {
  const today = new Date();
  const start = new Date();
  start.setDate(today.getDate() - n);
  return dt >= start && dt <= today;
}

/***** WEEKLY PRACTICES (rolling next occurrences) *****/
// Thursday 6:00pm–8:00pm (weekday=4), Sunday 8:40am–9:30am (weekday=0)
function nextWeeklyOccurrence(targetWeekday, startH, startM, endH, endM) {
  const now = new Date();
  const occ = new Date(now);
  const delta = (targetWeekday - now.getDay() + 7) % 7;
  occ.setDate(now.getDate() + delta);
  occ.setHours(startH, startM, 0, 0);
  // if today and already past end time, push 7 days
  const end = new Date(occ); end.setHours(endH, endM, 0, 0);
  if (now > end) { occ.setDate(occ.getDate() + 7); }
  return occ;
}

function renderWeeklyPractices() {
  const ul = document.getElementById("weekly-practice-list");
  if (!ul) return;

  const thurs = nextWeeklyOccurrence(4, 18, 0, 20, 0);      // Thurs 6–8pm
  const sun   = nextWeeklyOccurrence(0, 8, 40, 9, 30);      // Sun 8:40–9:30am

  ul.innerHTML = "";
  const li1 = document.createElement("li");
  li1.textContent = `${fmtDateRange(thurs, 18, 0, 20, 0)}`.replace(/^Thursdays?/, "Thurs");
  const li2 = document.createElement("li");
  li2.textContent = `${fmtDateRange(sun, 8, 40, 9, 30)}`.replace(/^Sundays?/, "Sun");
  // Ensure abbreviations
  li1.textContent = li1.textContent.replace(/^Thu,/, "Thurs,");
  ul.appendChild(li1);
  ul.appendChild(li2);
}

/***** ADDITIONAL PRACTICE (non-Thu/Sun) *****/
async function renderAdditionalPractice() {
  const tbody = document.getElementById("additional-practice-body");
  if (!tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.addPractice);
    const normalized = rows.map((r) => {
      const o = {}; Object.keys(r).forEach(k => o[k.trim().toLowerCase()] = r[k]);
      const date = excelToDate(o.date ?? o["practice date"] ?? o.day ?? "");
      const time = String(o.time ?? o["practice time"] ?? "").trim();
      const notes = String(o.notes ?? o.note ?? "").trim();
      return { date, time, notes };
    }).filter(x => x.date && ![0,4].includes(x.date.getDay())) // exclude Sun(0) & Thu(4)
      .sort((a,b) => a.date - b.date);

    tbody.innerHTML = "";
    if (!normalized.length) {
      tbody.innerHTML = `<tr><td colspan="3">No additional practices listed.</td></tr>`;
      return;
    }
    normalized.forEach(({date,time,notes}) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDateOnly(date)}</td><td>${time || "-"}</td><td>${notes || ""}</td>`;
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Additional Practice error:", e);
    tbody.innerHTML = `<tr><td colspan="3">Unable to load Additional Practice.</td></tr>`;
  }
}

/***** ANNOUNCEMENTS (EN+HM, newest, 31d) *****/
async function renderAnnouncements() {
  const ul = document.getElementById("announcements-list");
  if (!ul) return;
  try {
    const rows = await fetchXlsxRows(SRC.announcements);
    const data = rows.map((r) => {
      const o = {}; Object.keys(r).forEach(k => o[k.trim().toLowerCase()] = r[k]);
      const date = excelToDate(o.date ?? o["announcement date"] ?? o.created ?? "");
      const english = String(o.english ?? o.message ?? o.announcement ?? "").trim();
      const hmong   = String(o.hmong ?? "").trim();
      return { date, english, hmong };
    }).filter(x => x.date && (x.english || x.hmong))
      .filter(x => withinLastNDays(x.date, 31))
      .sort((a,b) => b.date - a.date);

    ul.innerHTML = "";
    if (!data.length) { ul.innerHTML = `<li>No announcements in the last 31 days.</li>`; return; }
    data.forEach(({date, english, hmong}) => {
      const li = document.createElement("li");
      let html = `<strong>${fmtDateOnly(date)}:</strong>`;
      if (english) html += `<br><em>(English)</em> ${english}`;
      if (hmong)   html += `<br><em>(Hmong)</em> ${hmong}`;
      li.innerHTML = html;
      ul.appendChild(li);
    });
  } catch (e) {
    console.error("Announcements error:", e);
    ul.innerHTML = `<li>Unable to load announcements.</li>`;
  }
}

/***** BIBLE STUDY (generic table) *****/
async function renderBibleStudy() {
  const thead = document.getElementById("bible-study-head");
  const tbody = document.getElementById("bible-study-body");
  if (!thead || !tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.bibleStudy);
    thead.innerHTML = ""; tbody.innerHTML = "";
    if (!rows.length) { thead.innerHTML = "<tr><th>Info</th></tr>"; tbody.innerHTML = "<tr><td>No data</td></tr>"; return; }
    const headers = Object.keys(rows[0]);
    thead.innerHTML = `<tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr>`;
    rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = headers.map(h => `<td>${r[h]}</td>`).join("");
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Bible Study error:", e);
    thead.innerHTML = "<tr><th>Error</th></tr>"; tbody.innerHTML = "<tr><td>Unable to load Bible Study.</td></tr>";
  }
}

/***** MEMBERS (Name, Role) *****/
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

/***** SETLIST (generic table) *****/
async function renderSetlist() {
  const thead = document.getElementById("setlist-head");
  const tbody = document.getElementById("setlist-body");
  if (!thead || !tbody) return;
  try {
    const rows = await fetchXlsxRows(SRC.setlist);
    thead.innerHTML = ""; tbody.innerHTML = "";
    if (!rows.length) { thead.innerHTML = "<tr><th>Info</th></tr>"; tbody.innerHTML = "<tr><td>No data</td></tr>"; return; }
    const headers = Object.keys(rows[0]);
    thead.innerHTML = `<tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr>`;
    rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = headers.map(h => `<td>${r[h]}</td>`).join("");
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Setlist error:", e);
    thead.innerHTML = "<tr><th>Error</th></tr>"; tbody.innerHTML = "<tr><td>Unable to load Setlist.</td></tr>";
  }
}

/***** SONGS CATALOG (CSV table) *****/
async function renderCatalog() {
  const thead = document.getElementById("catalog-head");
  const tbody = document.getElementById("catalog-body");
  if (!thead || !tbody) return;
  try {
    const rows = await fetchCsv(SRC.catalog);
    thead.innerHTML = ""; tbody.innerHTML = "";
    if (!rows.length) { thead.innerHTML = "<tr><th>Info</th></tr>"; tbody.innerHTML = "<tr><td>No data</td></tr>"; return; }
    const headers = Object.keys(rows[0]);
    thead.innerHTML = `<tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr>`;
    rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = headers.map(h => `<td>${r[h]}</td>`).join("");
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Catalog error:", e);
    thead.innerHTML = "<tr><th>Error</th></tr>"; tbody.innerHTML = "<tr><td>Unable to load Songs Catalog.</td></tr>";
  }
}

/***** ANALYTICS *****/
// 1) Hymnal Coverage KPI
async function renderKpiCoverage() {
  const valEl = document.getElementById("kpi-coverage");
  const subEl = document.getElementById("kpi-coverage-sub");
  if (!valEl || !subEl) return;
  try {
    const rows = await fetchCsv(SRC.kpiCoverage);
    // Expecting columns like: total_hymns, used_hymns, coverage_pct
    if (!rows.length) { valEl.textContent = "—"; subEl.textContent = "No data"; return; }
    const r = rows[0];
    const pct = Number(r.coverage_pct ?? r.coverage ?? r.pct ?? 0);
    const used = r.used_hymns ?? r.used ?? "—";
    const total = r.total_hymns ?? r.total ?? "—";
    valEl.textContent = isFinite(pct) ? `${pct}%` : `${pct}`;
    subEl.textContent = `Used ${used} of ${total}`;
  } catch (e) {
    console.error("Coverage KPI error:", e);
    valEl.textContent = "—";
    subEl.textContent = "Unable to load";
  }
}

// 2) Top 10 (3D pie)
async function renderTop10Pie() {
  const el = document.getElementById("chart-top10");
  if (!el) return;
  try {
    const rows = await fetchCsv(SRC.kpiTop10);
    // Expect columns: Song, Plays (or similar)
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const nameKey = headers.find(h => /song|title/i.test(h)) || "Song";
    const valKey  = headers.find(h => /play|count|times/i.test(h)) || "Plays";
    const dataArr = [["Song", "Plays"]];
    rows.forEach(r => {
      const name = String(r[nameKey] ?? "").trim();
      const val = Number(r[valKey] ?? 0);
      if (name && isFinite(val)) dataArr.push([name, val]);
    });

    const data = google.visualization.arrayToDataTable(dataArr);
    const options = { is3D: true, backgroundColor: "transparent", legend: { textStyle: { color: "#e5e7eb" } }, chartArea: { width: "90%", height: "80%" } };
    const chart = new google.visualization.PieChart(el);
    chart.draw(data, options);
  } catch (e) {
    console.error("Top10 Pie error:", e);
    el.innerHTML = "Unable to render Top 10.";
  }
}

// 3) KPI by Source (bar)
async function renderBySourceBar() {
  const el = document.getElementById("chart-by-source");
  if (!el) return;
  try {
    const rows = await fetchCsv(SRC.kpiBySource);
    // Expect columns: Source, Count
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const srcKey = headers.find(h => /source|category/i.test(h)) || "Source";
    const valKey = headers.find(h => /count|plays|value/i.test(h)) || "Count";
    const dataArr = [[srcKey, valKey]];
    rows.forEach(r => {
      const k = String(r[srcKey] ?? "").trim();
      const v = Number(r[valKey] ?? 0);
      if (k && isFinite(v)) dataArr.push([k, v]);
    });

    const data = google.visualization.arrayToDataTable(dataArr);
    const options = {
      backgroundColor: "transparent",
      legend: { position: "none" },
      hAxis: { textStyle: { color: "#e5e7eb" }, gridlines: { color: "#1f2937" } },
      vAxis: { textStyle: { color: "#e5e7eb" }, gridlines: { color: "#1f2937" }, minValue: 0 },
      chartArea: { width: "85%", height: "70%" },
    };
    const chart = new google.visualization.ColumnChart(el);
    chart.draw(data, options);
  } catch (e) {
    console.error("BySource Bar error:", e);
    el.innerHTML = "Unable to render KPI by Source.";
  }
}

// 4) Unused Hymnal table
async function renderUnusedTable() {
  const thead = document.getElementById("unused-head");
  const tbody = document.getElementById("unused-body");
  if (!thead || !tbody) return;
  try {
    const rows = await fetchCsv(SRC.hymnalUnused);
    thead.innerHTML = ""; tbody.innerHTML = "";
    if (!rows.length) { thead.innerHTML = "<tr><th>Info</th></tr>"; tbody.innerHTML = "<tr><td>No data</td></tr>"; return; }
    const headers = Object.keys(rows[0]);
    thead.innerHTML = `<tr>${headers.map(h=>`<th>${h}</th>`).join("")}</tr>`;
    rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = headers.map(h => `<td>${r[h]}</td>`).join("");
      tbody.appendChild(tr);
    });
  } catch (e) {
    console.error("Unused Hymnal error:", e);
    thead.innerHTML = "<tr><th>Error</th></tr>"; tbody.innerHTML = "<tr><td>Unable to load Unused Hymnal.</td></tr>";
  }
}

/***** INIT *****/
function loadGoogle() {
  return new Promise((resolve) => {
    google.charts.load("current", { packages: ["corechart"] });
    google.charts.setOnLoadCallback(resolve);
  });
}

document.addEventListener("DOMContentLoaded", async () => {
  try {
    renderWeeklyPractices();
    await Promise.all([
      renderAdditionalPractice(),
      renderAnnouncements(),
      renderBibleStudy(),
      renderMembers(),
      renderSetlist(),
      renderCatalog(),
      renderKpiCoverage(),
      renderUnusedTable()
    ]);
    await loadGoogle();
    await Promise.all([renderTop10Pie(), renderBySourceBar()]);
  } catch (e) {
    console.error("Init error:", e);
  }
});
