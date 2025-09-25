// ===== Configuration: Excel sources on GitHub =====
const ANNOUNCEMENTS_XLSX =
  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/announcements/announcements.xlsx";
const SPECIAL_PRACTICE_XLSX =
  "https://github.com/YSayaovong/HFBC_Praise_Worship/blob/main/special_practice/special_practice.xlsx";

// ---- Helpers ----
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
  return XLSX.utils.sheet_to_json(sheet, { defval: "" }); // array of row objects
}

// Excel date normalization: supports ISO strings, human strings, and Excel serials.
function excelToDate(val) {
  if (val == null || val === "") return null;
  if (typeof val === "number") {
    const d = XLSX.SSF.parse_date_code(val);
    if (!d) return null;
    return new Date(Date.UTC(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0));
  }
  // Strings – try Date parse
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(dt, withTime = false) {
  if (!dt) return "";
  const dateOpts = { weekday: "short", year: "numeric", month: "short", day: "numeric" };
  const timeOpts = { hour: "numeric", minute: "2-digit" };
  if (withTime) {
    return `${dt.toLocaleDateString(undefined, dateOpts)} ${dt.toLocaleTimeString(undefined, timeOpts)}`;
  }
  return dt.toLocaleDateString(undefined, dateOpts);
}

function withinLastNDays(dt, n = 31) {
  const today = new Date();
  const past = new Date(today);
  past.setDate(today.getDate() - n);
  return dt >= past && dt <= today;
}

// ---- Renderers ----
async function loadSpecialPractice() {
  try {
    const rows = await fetchSheetJSON(SPECIAL_PRACTICE_XLSX);
    // Expected columns (case-insensitive): Date | Time | Notes
    const tbody = document.getElementById("special-practice-body");
    tbody.innerHTML = "";

    rows
      .map((r) => {
        // normalize column names
        const obj = {};
        Object.keys(r).forEach((k) => (obj[k.trim().toLowerCase()] = r[k]));
        const date = excelToDate(obj.date ?? obj["practice date"] ?? obj["day"] ?? "");
        const time = String(obj.time ?? obj["practice time"] ?? "").trim();
        const notes = String(obj.notes ?? obj.note ?? "").trim();
        return { date, time, notes };
      })
      .filter((x) => x.date) // require a valid date
      .sort((a, b) => a.date - b.date) // upcoming first
      .forEach(({ date, time, notes }) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${fmtDate(date, false)}</td>
          <td>${time || "-"}</td>
          <td>${notes || ""}</td>`;
        tbody.appendChild(tr);
      });
  } catch (e) {
    console.error("Special Practice load error:", e);
    document.getElementById("special-practice-body").innerHTML =
      `<tr><td colspan="3">Unable to load Special Practice.</td></tr>`;
  }
}

async function loadAnnouncements() {
  try {
    const rows = await fetchSheetJSON(ANNOUNCEMENTS_XLSX);
    // Expected columns (case-insensitive): Date | Message (or Announcement)
    const list = document.getElementById("announcements-list");
    list.innerHTML = "";

    const normalized = rows
      .map((r) => {
        const obj = {};
        Object.keys(r).forEach((k) => (obj[k.trim().toLowerCase()] = r[k]));
        const date = excelToDate(obj.date ?? obj["announcement date"] ?? obj["created"] ?? "");
        const msg = String(obj.message ?? obj.announcement ?? obj.note ?? "").trim();
        return { date, msg };
      })
      .filter(({ date, msg }) => date && msg)
      .filter(({ date }) => withinLastNDays(date, 31)) // keep last 31 days only
      .sort((a, b) => b.date - a.date); // newest first

    if (!normalized.length) {
      const li = document.createElement("li");
      li.textContent = "No announcements in the last 31 days.";
      list.appendChild(li);
      return;
    }

    normalized.forEach(({ date, msg }) => {
      const li = document.createElement("li");
      li.innerHTML = `<strong>${fmtDate(date, false)}:</strong> ${msg}`;
      list.appendChild(li);
    });
  } catch (e) {
    console.error("Announcements load error:", e);
    const list = document.getElementById("announcements-list");
    list.innerHTML = `<li>Unable to load announcements.</li>`;
  }
}

// ---- Init ----
document.addEventListener("DOMContentLoaded", async () => {
  await Promise.all([loadSpecialPractice(), loadAnnouncements()]);
});
