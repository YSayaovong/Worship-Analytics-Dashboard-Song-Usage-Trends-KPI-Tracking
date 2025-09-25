// ======= Config: Excel sources on GitHub =======
const ANNOUNCEMENTS_XLSX =
  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx";
const SPECIAL_PRACTICE_XLSX =
  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/special_practice.xlsx";

// ======= Helpers =======
const toRawGitHub = (blobUrl) =>
  blobUrl.replace("https://github.com/", "https://raw.githubusercontent.com/").replace("/blob/", "/");

async function fetchSheetJSON(githubBlobUrl, sheetNameOrIndex = 0) {
  const rawUrl = toRawGitHub(githubBlobUrl);
  const res = await fetch(rawUrl);
  if (!res.ok) throw new Error(`Failed to fetch: ${rawUrl}`);
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
    // local time to avoid GMT strings
    return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(dt) {
  if (!dt) return "";
  const opts = { weekday: "short", year: "numeric", month: "short", day: "numeric" };
  return dt.toLocaleDateString(undefined, opts);
}

function withinLastNDays(dt, n = 31) {
  const today = new Date();
  const start = new Date();
  start.setDate(today.getDate() - n);
  return dt >= start && dt <= today;
}

// ======= 1) Worship Practice: strip label words only =======
function sanitizeWorshipPracticeLabels() {
  const list = document.getElementById("worship-practice-list");
  if (!list) return;
  [...list.querySelectorAll("li")].forEach((li) => {
    const cleaned = li.textContent.trim().replace(
      /^\s*(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s+practice\s*:\s*/i,
      ""
    );
    li.textContent = cleaned;
  });
}

// ======= 2) Special Practice (Date, Time, Notes) =======
async function renderSpecialPractice() {
  const tbody = document.getElementById("special-practice-body");
  if (!tbody) return;

  try {
    const rows = await fetchSheetJSON(SPECIAL_PRACTICE_XLSX);

    const normalized = rows
      .map((r) => {
        const o = {};
        Object.keys(r).forEach((k) => (o[k.trim().toLowerCase()] = r[k]));
        const date = excelToDate(o.date ?? o["practice date"] ?? o.day ?? "");
        const time = String(o.time ?? o["practice time"] ?? "").trim();
        const notes = String(o.notes ?? o.note ?? "").trim();
        return { date, time, notes };
      })
      .filter((x) => x.date)
      .sort((a, b) => a.date - b.date);

    tbody.innerHTML = "";
    if (!normalized.length) {
      tbody.innerHTML = `<tr><td colspan="3">No special practices listed.</td></tr>`;
      return;
    }

    normalized.forEach(({ date, time, notes }) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${fmtDate(date)}</td><td>${time || "-"}</td><td>${notes || ""}</td>`;
      tbody.appendChild(tr);
    });
  } catch (err) {
    console.error("Special Practice error:", err);
    tbody.innerHTML = `<tr><td colspan="3">Unable to load Special Practice.</td></tr>`;
  }
}

// ======= 3) Announcements (English + Hmong; newest first; last 31 days) =======
async function renderAnnouncements() {
  const list = document.getElementById("announcements-list");
  if (!list) return;

  try {
    const rows = await fetchSheetJSON(ANNOUNCEMENTS_XLSX);

    const normalized = rows
      .map((r) => {
        const o = {};
        Object.keys(r).forEach((k) => (o[k.trim().toLowerCase()] = r[k]));
        const date = excelToDate(o.date ?? o["announcement date"] ?? o.created ?? "");
        const english = String(o.english ?? o.message ?? o.announcement ?? "").trim();
        const hmong = String(o.hmong ?? "").trim();
        return { date, english, hmong };
      })
      .filter(({ date, english, hmong }) => date && (english || hmong))
      .filter(({ date }) => withinLastNDays(date, 31))
      .sort((a, b) => b.date - a.date);

    list.innerHTML = "";
    if (!normalized.length) {
      const li = document.createElement("li");
      li.textContent = "No announcements in the last 31 days.";
      list.appendChild(li);
      return;
    }

    normalized.forEach(({ date, english, hmong }) => {
      const li = document.createElement("li");
      let html = `<strong>${fmtDate(date)}:</strong>`;
      if (english) html += ` <br><em>(English)</em> ${english}`;
      if (hmong)   html += ` <br><em>(Hmong)</em> ${hmong}`;
      li.innerHTML = html;
      list.appendChild(li);
    });
  } catch (err) {
    console.error("Announcements error:", err);
    list.innerHTML = `<li>Unable to load announcements.</li>`;
  }
}

// ======= Init =======
document.addEventListener("DOMContentLoaded", () => {
  sanitizeWorshipPracticeLabels();
  renderSpecialPractice();
  renderAnnouncements();
});
