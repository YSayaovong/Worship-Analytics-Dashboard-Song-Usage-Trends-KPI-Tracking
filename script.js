// ---------- CONFIG ----------
// Your public repo details (adjust if you rename or move)
const GITHUB = {
  owner: "YSayaovong",
  repo: "HFBC_Praise_Worship",
  branch: "main"
};

// Relative paths inside the repo
const PATHS = {
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  setlist: "setlist/setlist.xlsx" // expected to contain multiple rows; first sheet, header row
};

// ---------- HELPERS ----------
const $ = (sel, root=document) => root.querySelector(sel);

function rawUrl(pathRel){
  // Build raw.githubusercontent URL
  return `https://raw.githubusercontent.com/${GITHUB.owner}/${GITHUB.repo}/${GITHUB.branch}/${pathRel}`;
}

async function fetchWB(pathRel){
  const url = rawUrl(pathRel) + `?nocache=${Date.now()}`;
  const res = await fetch(url);
  if(!res.ok) throw new Error(`Fetch failed: ${url} (${res.status})`);
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}

function sheetToAOA(wb){
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(ws, { header:1, defval:"" }); // array of arrays
}

function renderAOATable(aoa, targetSel){
  const el = $(targetSel);
  if(!aoa || aoa.length === 0){
    el.innerHTML = `<p class="dim">No data found.</p>`;
    return;
  }
  let html = `<table><thead><tr>`;
  const header = aoa[0];
  header.forEach(h => html += `<th>${escapeHtml(String(h))}</th>`);
  html += `</tr></thead><tbody>`;
  for(let i=1;i<aoa.length;i++){
    const row = aoa[i];
    // Skip completely empty rows
    if(row.every(cell => String(cell).trim() === "")) continue;
    html += `<tr>` + header.map((_, cIdx)=>`<td>${escapeHtml(String(row[cIdx] ?? ""))}</td>`).join("") + `</tr>`;
  }
  html += `</tbody></table>`;
  el.innerHTML = html;
}

function escapeHtml(s){
  return s.replace(/[&<>"']/g, m => ({
    "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"
  }[m]));
}

// ---------- LOADERS ----------
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = sheetToAOA(wb);
    renderAOATable(aoa, "#announcements-table");
  }catch(err){
    console.error("Announcements error:", err);
    $("#announcements-table").innerHTML =
      `<p class="dim">Unable to load <code>${PATHS.announcements}</code>. Ensure the file exists and is public.</p>`;
  }
}

async function loadMembers(){
  try{
    const wb = await fetchWB(PATHS.members);
    const aoa = sheetToAOA(wb);
    renderAOATable(aoa, "#members-table");
  }catch(err){
    console.error("Members error:", err);
    $("#members-table").innerHTML =
      `<p class="dim">Unable to load <code>${PATHS.members}</code>. Ensure the file exists and is public.</p>`;
  }
}

async function loadSetlistAndAnalytics(){
  try{
    const wb = await fetchWB(PATHS.setlist);
    const aoa = sheetToAOA(wb);
    if(!aoa || aoa.length < 2){
      $("#setlist-table").innerHTML = `<p class="dim">No setlist rows found.</p>`;
      return;
    }

    // Try to detect common headers
    const headers = aoa[0].map(h => String(h).trim().toLowerCase());
    const idxDate = headers.findIndex(h => ["date","service date"].includes(h));
    const idxSong = headers.findIndex(h => ["song","title","song title"].includes(h));
    const idxKey  = headers.findIndex(h => h === "key");
    const idxNotes = headers.findIndex(h => ["notes","comment"].includes(h));

    // Render the visible "This Week" by taking the most recent date if Date column exists,
    // otherwise just render all rows after header.
    let rows = aoa.slice(1).filter(r => r.some(cell => String(cell).trim() !== ""));
    if(idxDate !== -1){
      // Parse dates and group by date
      const parsed = rows.map(r => ({
        date: parseDateLoose(r[idxDate]),
        row: r
      })).filter(o => o.date !== null);

      // Get latest date
      if(parsed.length){
        const latestDate = parsed.reduce((a,b)=> (a.date > b.date ? a : b)).date;
        const latestRows = parsed.filter(o => sameYMD(o.date, latestDate)).map(o => o.row);

        // Build a display AOA
        const displayHeader = [];
        if(idxSong !== -1) displayHeader.push("Song");
        if(idxKey  !== -1) displayHeader.push("Key");
        if(idxNotes!== -1) displayHeader.push("Notes");

        const displayRows = latestRows.map(r => {
          const out = [];
          if(idxSong !== -1) out.push(r[idxSong] ?? "");
          if(idxKey  !== -1) out.push(r[idxKey]  ?? "");
          if(idxNotes!== -1) out.push(r[idxNotes]?? "");
          return out;
        });

        $("#setlist-meta").textContent = `Service Date: ${latestDate.toLocaleDateString()}`;
        renderAOATable([displayHeader, ...displayRows], "#setlist-table");
      }else{
        // Fallback: render entire sheet
        renderAOATable(aoa, "#setlist-table");
      }
    }else{
      // No date col: render entire sheet
      $("#setlist-meta").textContent = "";
      renderAOATable(aoa, "#setlist-table");
    }

    // ----- Analytics -----
    // Count plays per song (across the whole sheet)
    if(idxSong !== -1){
      const counts = new Map();
      const perDate = new Map(); // dateStr -> count rows that day
      for(const r of rows){
        const title = String(r[idxSong] ?? "").trim();
        if(title){
          counts.set(title, (counts.get(title) || 0) + 1);
        }
        if(idxDate !== -1){
          const d = parseDateLoose(r[idxDate]);
          if(d){
            const key = d.toISOString().slice(0,10);
            perDate.set(key, (perDate.get(key)||0)+1);
          }
        }
      }

      // Top/Bottom 5
      const sorted = [...counts.entries()].sort((a,b)=> b[1]-a[1]);
      const top5 = sorted.slice(0,5);
      const bottom5 = sorted.slice(-5).reverse();

      $("#top5").innerHTML = top5.length ? top5.map(([s,c])=>`<li>${escapeHtml(s)} — ${c}</li>`).join("") : `<li class="dim">No data</li>`;
      $("#bottom5").innerHTML = bottom5.length ? bottom5.map(([s,c])=>`<li>${escapeHtml(s)} — ${c}</li>`).join("") : `<li class="dim">No data</li>`;

      // Charts
      drawPieChart(sorted.slice(0,7)); // top 7 for readability
      drawBarChart(top5);
      drawLineChart([...perDate.entries()].sort((a,b)=> a[0].localeCompare(b[0])));
    }else{
      $("#top5").innerHTML = `<li class="dim">Add a "Song" header to enable analytics.</li>`;
      $("#bottom5").innerHTML = `<li class="dim">Add a "Song" header to enable analytics.</li>`;
    }

  }catch(err){
    console.error("Setlist error:", err);
    $("#setlist-table").innerHTML =
      `<p class="dim">Unable to load <code>${PATHS.setlist}</code>. Ensure the file exists and is public.</p>`;
    $("#top5").innerHTML = `<li class="dim">No data</li>`;
    $("#bottom5").innerHTML = `<li class="dim">No data</li>`;
  }
}

// ---------- Date helpers ----------
function parseDateLoose(val){
  if(val == null) return null;
  // If Excel serial date
  if(typeof val === "number"){
    try{
      return XLSX.SSF.parse_date_code(val)
        ? excelSerialToDate(val)
        : null;
    }catch{ return null; }
  }
  // Try native Date parsing
  const s = String(val).trim();
  if(!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function excelSerialToDate(serial){
  // Excel serial to JS Date (assuming 1900 system)
  const utc_days  = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400; // seconds
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 1e-7;
  let totalSeconds = Math.floor(86400 * fractional_day);
  const seconds = totalSeconds % 60;
  totalSeconds = Math.floor(totalSeconds / 60);
  const minutes = totalSeconds % 60;
  const hours = Math.floor(totalSeconds / 60);
  date_info.setHours(hours, minutes, seconds);
  return date_info;
}

function sameYMD(a,b){
  return a.getFullYear()===b.getFullYear() &&
         a.getMonth()===b.getMonth() &&
         a.getDate()===b.getDate();
}

// ---------- Charts ----------
let pieInst, barInst, lineInst;

function destroyChart(inst){
  if(inst){ inst.destroy(); }
}

function drawPieChart(entries){
  const ctx = $("#pieChart");
  destroyChart(pieInst);
  const labels = entries.map(([s])=>s);
  const data = entries.map(([,c])=>c);
  pieInst = new Chart(ctx, {
    type: "pie",
    data: { labels, datasets: [{ data }] },
    options: { responsive: true, plugins:{ legend:{ position:"bottom" } } }
  });
}

function drawBarChart(entries){
  const ctx = $("#barChart");
  destroyChart(barInst);
  const labels = entries.map(([s])=>s);
  const data = entries.map(([,c])=>c);
  barInst = new Chart(ctx, {
    type: "bar",
    data: { labels, datasets: [{ data }] },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true, ticks: { precision:0 } } }
    }
  });
}

function drawLineChart(datePairs){
  const ctx = $("#lineChart");
  destroyChart(lineInst);
  const labels = datePairs.map(([d])=>d);
  const data = datePairs.map(([,c])=>c);
  lineInst = new Chart(ctx, {
    type: "line",
    data: { labels, datasets: [{ data, tension: 0.2, fill:false }] },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true, ticks: { precision:0 } } }
    }
  });
}

// ---------- BOOT ----------
document.addEventListener("DOMContentLoaded", async () => {
  await loadAnnouncements();
  await loadMembers();
  await loadSetlistAndAnalytics();
});
