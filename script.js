// ====== CONFIG: GitHub source (Announcements only) ======
const ANNOUNCEMENTS_BLOB = "https://github.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/blob/main/announcements/announcements.xlsx";

// Convert GitHub 'blob' URL to raw content URL
const toRaw = (blobUrl) =>
  blobUrl.replace("https://github.com/", "https://raw.githubusercontent.com/").replace("/blob/", "/");

// Fetch rows from an Excel worksheet
async function fetchXlsxRows(blobUrl, sheetNameOrIndex = 0) {
  const url = toRaw(blobUrl) + "?v=" + Date.now();  // cache-bust
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

// Helpers
const DAY_ABBR = ["Sun","Mon","Tues","Wed","Thurs","Fri","Sat"];
const MONTH_ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"];
const fmtDateOnly = (dt) => `${DAY_ABBR[dt.getDay()]}, ${MONTH_ABBR[dt.getMonth()]} ${dt.getDate()}, ${dt.getFullYear()}`;

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
function norm(s){ return String(s||"").toLowerCase().replace(/[^a-z0-9]+/g,""); }
function normMap(row){ const m = {}; Object.keys(row||{}).forEach(k => m[norm(k)] = row[k]); return m; }
function val(m, keys){ for(const k of keys){ const v = m[k]; if(v!=null && String(v)!=="") return v; } return ""; }

// ====== Announcements renderer (English & Hmong) ======
async function renderAnnouncements(){
  const tbody = document.getElementById("announcements-body");
  if(!tbody) return;
  tbody.innerHTML = "";
  try{
    const rows = await fetchXlsxRows(ANNOUNCEMENTS_BLOB);
    const today = new Date();
    const THIRTY_ONE_DAYS = 31 * 24 * 60 * 60 * 1000;

    // Map rows into {d, en, hm} with flexible column names
    const items = rows.map(r => {
      const m = normMap(r);
      const d = excelToDate(val(m, ["date","day"]));
      const en = val(m, ["announcementenglish","announcement","english"]);
      const hm = val(m, ["hmong","lus","tshaj","lus tshaj tawm","lus_tshaj_tawm"]);
      return { d, en, hm };
    })
    .filter(x => x.d && (today - x.d) <= THIRTY_ONE_DAYS)
    .sort((a,b) => b.d - a.d);

    // Render
    items.forEach(it => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${fmtDateOnly(it.d)}</td>
        <td>${it.en || ""}</td>
        <td>${it.hm || ""}</td>
      `;
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

// ====== Init ======
document.addEventListener("DOMContentLoaded", () => {
  renderAnnouncements();
});
