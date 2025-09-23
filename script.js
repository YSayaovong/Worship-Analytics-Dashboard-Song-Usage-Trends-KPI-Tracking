/* Hmong First Baptist Church – Praise & Worship
   Static site: auto-loads Excel setlists from GitHub and shows latest + archive.
   Adds YouTube embed if present in the Excel file.
*/

// >>> CONFIG: set your repo info <<<
const GH = {
  owner: "YSayaovong",   // GitHub username
  repo: "hfb-pw-site",   // Repository name
  branch: "main"         // Branch with Pages
};
const SETLIST_DIR = "setlists";

// --- GitHub helpers ---
const ghApiUrl = (path) =>
  `https://api.github.com/repos/${GH.owner}/${GH.repo}/contents/${path}?ref=${GH.branch}`;
const ghRawUrl = (path) =>
  `https://raw.githubusercontent.com/${GH.owner}/${GH.repo}/${GH.branch}/${path}`;

// --- Excel helpers ---
async function fetchWorkbook(path) {
  const res = await fetch(ghRawUrl(path));
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}
function firstSheetToAOA(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

// --- Render helpers ---
function renderTable(aoa, target) {
  if (!aoa || aoa.length === 0) {
    target.textContent = "No data.";
    return;
  }
  let html = "<table>";
  aoa.forEach((row, i) => {
    html += "<tr>";
    row.forEach(cell => {
      html += i === 0 ? `<th>${cell}</th>` : `<td>${cell}</td>`;
    });
    html += "</tr>";
  });
  html += "</table>";
  target.innerHTML = html;
}
function embedYouTube(url) {
  const ytDiv = document.getElementById("youtube-video");
  ytDiv.innerHTML = "";
  if (!url) return;
  const videoId = url.split("v=")[1]?.split("&")[0] || url.split("/").pop();
  if (!videoId) return;
  ytDiv.innerHTML = `<iframe src="https://www.youtube.com/embed/${videoId}" allowfullscreen></iframe>`;
}

// --- Main ---
async function init() {
  try {
    // list setlist files
    const res = await fetch(ghApiUrl(SETLIST_DIR));
    const items = await res.json();
    const files = items
      .filter(f => f.type === "file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(f.name))
      .map(f => f.name)
      .sort((a,b) => b.localeCompare(a)); // newest first

    if (files.length === 0) {
      document.getElementById("latest-setlist").textContent = "No setlists found.";
      return;
    }

    // Latest
    const latest = files[0];
    const wb = await fetchWorkbook(`${SETLIST_DIR}/${latest}`);
    const aoa = firstSheetToAOA(wb);
    renderTable(aoa, document.getElementById("latest-setlist"));

    // Try YouTube
    let ytUrl = "";
    if (wb.Sheets["Meta"]) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets["Meta"], { header: 1, defval: "" });
      const ytRow = rows.find(r => r[0]?.toLowerCase() === "youtube");
      ytUrl = ytRow ? ytRow[1] : "";
    } else if (aoa[0].includes("YouTube")) {
      const col = aoa[0].indexOf("YouTube");
      ytUrl = aoa[1]?.[col] || "";
    }
    embedYouTube(ytUrl);

    // Archive
    const list = document.getElementById("archive-list");
    list.innerHTML = files.map(f => {
      const nice = f.replace(".xlsx", "");
      return `<li><a href="#" onclick="loadSetlist('${f}')">${nice}</a></li>`;
    }).join("");
  } catch (err) {
    console.error(err);
    document.getElementById("latest-setlist").textContent = "Error loading setlist.";
  }
}

// load from archive link
async function loadSetlist(file) {
  const wb = await fetchWorkbook(`${SETLIST_DIR}/${file}`);
  const aoa = firstSheetToAOA(wb);
  renderTable(aoa, document.getElementById("latest-setlist"));

  // YouTube again
  let ytUrl = "";
  if (wb.Sheets["Meta"]) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Meta"], { header: 1, defval: "" });
    const ytRow = rows.find(r => r[0]?.toLowerCase() === "youtube");
    ytUrl = ytRow ? ytRow[1] : "";
  } else if (aoa[0].includes("YouTube")) {
    const col = aoa[0].indexOf("YouTube");
    ytUrl = aoa[1]?.[col] || "";
  }
  embedYouTube(ytUrl);
  window.location.hash = "#home";
}

document.addEventListener("DOMContentLoaded", init);
