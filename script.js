/* ===== Hmong FBC P&W — Static Frontend with Excel ingestion =====
 * - Lists weekly Excel setlists from the GitHub repo via GitHub API
 * - Renders latest setlist on Home
 * - Renders archive and any specific setlist by query string
 * - Renders announcements from announcements.xlsx
 * --------------------------------------------------------------- */

/** >>>> CONFIG: set these to your repo details <<<< */
const GH = {
  owner: "YSayaovong",             // <-- your GitHub username/org
  repo:  "hfb-pw-site",            // <-- your repository name
  branch:"main"                    // <-- branch hosting Pages
};

const PATHS = {
  setlists: "setlists",            // folder with weekly .xlsx files named YYYY-MM-DD.xlsx
  announcements: "announcements/announcements.xlsx" // single Excel file for announcements
};

/** Utility: GitHub API & Raw URLs */
const ghApiUrl = (path) =>
  `https://api.github.com/repos/${GH.owner}/${GH.repo}/contents/${encodeURIComponent(path)}?ref=${encodeURIComponent(GH.branch)}`;

const ghRawUrl = (path) =>
  `https://raw.githubusercontent.com/${GH.owner}/${GH.repo}/${GH.branch}/${path}`;

/** List directory contents via GitHub API (returns array of file objects) */
async function listDirectory(path) {
  const res = await fetch(ghApiUrl(path), { headers: { "Accept": "application/vnd.github+json" } });
  if (!res.ok) throw new Error(`GitHub API error: ${res.status} ${res.statusText}`);
  return res.json();
}

/** Fetch an Excel file from raw.githubusercontent.com, parse to workbook */
async function fetchWorkbook(rawPath) {
  const res = await fetch(ghRawUrl(rawPath));
  if (!res.ok) throw new Error(`Fetch error for ${rawPath}: ${res.status} ${res.statusText}`);
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}

/** Convert first sheet to array-of-arrays */
function firstSheetToAOA(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

/** Render a simple table to a target element */
function renderTable(aoa, targetId, caption) {
  const el = document.getElementById(targetId);
  if (!el) return;
  if (!aoa || aoa.length === 0) { el.textContent = "No data found."; return; }

  let html = "<div class='table-wrap'>";
  if (caption) html += `<p class="meta">${caption}</p>`;
  html += "<table>";
  aoa.forEach((row, i) => {
    html += "<tr>";
    row.forEach(cell => {
      const safe = String(cell);
      html += (i === 0) ? `<th>${safe}</th>` : `<td>${safe}</td>`;
    });
    html += "</tr>";
  });
  html += "</table></div>";
  el.innerHTML = html;
}

/** Format YYYY-MM-DD to friendly date */
function friendlyDateFromFilename(name) {
  // name like 2025-10-05.xlsx
  const m = name.match(/^(\d{4})-(\d{2})-(\d{2})\.xlsx$/);
  if (!m) return name;
  const d = new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`);
  return d.toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
}

/** Compare by date embedded in filename */
function compareSetlistNamesDesc(a, b) {
  const da = a.replace(".xlsx",""), db = b.replace(".xlsx","");
  return db.localeCompare(da); // YYYY-MM-DD sorts correctly as strings
}

/** Extract list of setlist filenames matching YYYY-MM-DD.xlsx */
async function getSetlistFiles() {
  const items = await listDirectory(PATHS.setlists);
  return items
    .filter(it => it.type === "file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(it.name))
    .map(it => it.name)
    .sort(compareSetlistNamesDesc); // newest first
}

/** Home: render latest setlist + latest announcements */
async function renderHome() {
  try {
    // Latest setlist
    const files = await getSetlistFiles();
    if (files.length === 0) {
      document.getElementById("latest-setlist").textContent = "No setlists found.";
    } else {
      const latest = files[0];
      const wb = await fetchWorkbook(`${PATHS.setlists}/${latest}`);
      const aoa = firstSheetToAOA(wb);
      const caption = `${friendlyDateFromFilename(latest)} • (${latest})`;
      renderTable(aoa, "latest-setlist", caption);
    }

    // Announcements (limit 5 rows after header)
    await renderAnnouncements("latest-announcements", 5);
  } catch (err) {
    console.error(err);
    const el = document.getElementById("latest-setlist");
    if (el) el.textContent = "Error loading latest setlist.";
    const an = document.getElementById("latest-announcements");
    if (an) an.textContent = "Error loading announcements.";
  }
}

/** Archive: list all setlists with links to setlist.html?file=... */
async function renderArchive() {
  try {
    const ul = document.getElementById("archive-list");
    if (!ul) return;

    const files = await getSetlistFiles();
    if (files.length === 0) {
      ul.innerHTML = "<li>No setlists found.</li>";
      return;
    }

    ul.innerHTML = files.map(name => {
      const nice = friendlyDateFromFilename(name);
      const href = `./setlist.html?file=${encodeURIComponent(name)}`;
      return `<li><a href="${href}">${nice}</a> <span class="dim">(${name})</span></li>`;
    }).join("");
  } catch (err) {
    console.error(err);
    const ul = document.getElementById("archive-list");
    if (ul) ul.innerHTML = "<li>Error loading archive.</li>";
  }
}

/** Setlist page: render by ?file=YYYY-MM-DD.xlsx */
async function renderSetlistFromQuery() {
  try {
    const params = new URLSearchParams(window.location.search);
    const file = params.get("file");
    const titleEl = document.getElementById("setlist-title");
    if (!file || !/^\d{4}-\d{2}-\d{2}\.xlsx$/.test(file)) {
      document.getElementById("setlist-view").textContent = "Invalid or missing setlist file.";
      if (titleEl) titleEl.textContent = "Setlist";
      return;
    }
    if (titleEl) titleEl.textContent = `Setlist — ${friendlyDateFromFilename(file)}`;

    const wb = await fetchWorkbook(`${PATHS.setlists}/${file}`);
    const aoa = firstSheetToAOA(wb);
    renderTable(aoa, "setlist-view", `(${file})`);
  } catch (err) {
    console.error(err);
    const el = document.getElementById("setlist-view");
    if (el) el.textContent = "Error loading setlist.";
  }
}

/** Announcements: read a single Excel file and render (optional limit) */
async function renderAnnouncements(targetId, limit = null) {
  try {
    const wb = await fetchWorkbook(PATHS.announcements);
    const aoa = firstSheetToAOA(wb);
    if (!aoa || aoa.length === 0) {
      document.getElementById(targetId).textContent = "No announcements.";
      return;
    }

    // Expect first row to be headers (e.g., Date | Title | Details)
    let rows = aoa.slice(1); // skip header for limiting/sorting
    // Try to sort by Date desc if first column looks like a date
    rows.sort((a, b) => {
      const da = new Date(a[0]); const db = new Date(b[0]);
      return db - da;
    });

    if (limit) rows = rows.slice(0, limit);
    const table = [aoa[0], ...rows]; // reattach header
    renderTable(table, targetId);
  } catch (err) {
    console.error(err);
    const el = document.getElementById(targetId);
    if (el) el.textContent = "Error loading announcements.";
  }
}

/** Announcements page: full render */
async function renderAnnouncementsPage() {
  await renderAnnouncements("announcements-table", null);
}

/** Expose public API for inline scripts */
window.HFBPW = {
  renderHome,
  renderArchive,
  renderSetlistFromQuery,
  renderAnnouncementsPage
};

/** Auto-run Home when index.html loads (exists check prevents errors on other pages) */
document.addEventListener("DOMContentLoaded", () => {
  if (document.getElementById("latest-setlist")) {
    HFBPW.renderHome();
  }
});
