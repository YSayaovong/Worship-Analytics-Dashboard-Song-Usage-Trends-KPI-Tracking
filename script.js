/* Hmong First Baptist Church – Praise & Worship
   Static site (no backend) powered by Excel files on GitHub.

   Reads weekly content:
     • /setlist/setlist.xlsx              (preferred "This Coming Week")
     • /setlists/YYYY-MM-DD.xlsx          (archive, used for Last Week & analytics)
     • /announcements/announcements.xlsx  (hide >31 days)
     • /members/members.xlsx              (team roster)

   Analytics: current calendar year only; Song + Plays (no keys).
   Charts: Bar + Pie via Chart.js.
*/

// ---- repo config ----
const GH = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };
const PATHS = {
  specialCurrent: "setlist/setlist.xlsx",
  setlistsDir: "setlists",
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx"
};

// ---- basic helpers ----
const apiURL = (p) => `https://api.github.com/repos/${GH.owner}/${GH.repo}/contents/${encodeURIComponent(p)}?ref=${encodeURIComponent(GH.branch)}`;
const rawURL = (p) => `https://raw.githubusercontent.com/${GH.owner}/${GH.repo}/${GH.branch}/${p}`;
const $ = (s) => document.querySelector(s);

async function existsOnGitHub(path){
  const r = await fetch(apiURL(path), { headers:{ "Accept":"application/vnd.github+json" }});
  return r.ok;
}
async function listDir(path){
  const r = await fetch(apiURL(path), { headers:{ "Accept":"application/vnd.github+json" }});
  if(!r.ok) return [];
  return r.json();
}
async function fetchWB(path){
  const r = await fetch(rawURL(path));
  if(!r.ok) throw new Error(`Fetch error ${r.status} for ${path}`);
  const ab = await r.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}
function firstSheetAOA(wb){
  const sh = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sh, { header:1, defval:"" });
}
function toDateFromName(name){
  const m = /^(\d{4})-(\d{2})-(\d{2})\.xlsx$/.exec(name);
  return m ? new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`) : null;
}
function niceDate(d){
  return d.toLocaleDateString(undefined,{ month:"short", day:"numeric", year:"numeric" }); // "Sep 20, 2025"
}
function weekRangeSunday(d){
  const start = new Date(d); start.setHours(0,0,0,0);
  start.setDate(start.getDate() - start.getDay()); // Sunday
  const end = new Date(start); end.setDate(start.getDate() + 6); end.setHours(23,59,59,999);
  return { start, end };
}

// ---- Excel helpers ----
function excelSerialToDate(val){
  if (typeof val === "number") {
    const o = XLSX.SSF.parse_date_code(val);
    if (o) return new Date(o.y, o.m - 1, o.d);
  }
  if (typeof val === "string" && val.trim() !== "") {
    const num = Number(val);
    if (!Number.isNaN(num)) {
      const o = XLSX.SSF.parse_date_code(num);
      if (o) return new Date(o.y, o.m - 1, o.d);
    }
    const d = new Date(val);
    if (!Number.isNaN(d)) return d;
  }
  return null;
}
function normalizeAnnouncementText(s){
  let out = String(s ?? "");
  out = out.replace(/:(?=\d)/g, ": ");
  out = out.replace(/([A-Za-z])(\d{1,2}\/\d{1,2}\/\d{2,4})/g, "$1 $2");
  return out.trim();
}

// ---- smart table renderer (formats DATE columns) ----
function renderTable(aoa, targetSel){
  const el = $(targetSel);
  if(!el){ return; }
  if(!aoa || aoa.length===0){ el.textContent = "No data."; return; }

  const headers = (aoa[0] || []).map(h => String(h));
  const dateCols = new Set();
  headers.forEach((h,i)=>{ if(/date/i.test(h)) dateCols.add(i); });

  function prettyDateCell(v){
    const d = excelSerialToDate(v) || new Date(v);
    return (d && !isNaN(d)) ? niceDate(d) : String(v);
  }

  let html = "<table>";
  aoa.forEach((row,rIdx)=>{
    html += "<tr>";
    row.forEach((cell,cIdx)=>{
      let out = String(cell);
      if(rIdx>0 && dateCols.has(cIdx)) out = prettyDateCell(cell);
      html += (rIdx===0? `<th>${out}</th>` : `<td>${out}</td>`);
    });
    html += "</tr>";
  });
  html += "</table>";
  el.innerHTML = html;
}

function renderSermon(targetSel, sermon){
  const el = $(targetSel);
  el.textContent = sermon ? `Sermon: ${sermon}` : "";
}

// ---- parse Meta/columns for Sermon/Date ----
function parseMeta(wb, aoa){
  let sermon="", serviceDate=null;

  if(wb.Sheets["Meta"]){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets["Meta"], { header:1, defval:"" });
    for(const r of rows){
      const k = String(r[0]||"").toLowerCase();
      const v = r[1] || "";
      if(k==="sermon")  sermon = v || sermon;
      if(k==="date" || k==="servicedate" || k==="service date"){
        const d = excelSerialToDate(v) || new Date(v);
        if(!isNaN(d)) serviceDate = d;
      }
    }
  }
  if(aoa && aoa[0]){
    const hdr = aoa[0].map(h=>String(h).toLowerCase());
    const sIdx = hdr.indexOf("sermon");
    const dIdx = hdr.indexOf("date")>-1 ? hdr.indexOf("date")
               : (hdr.indexOf("servicedate")>-1 ? hdr.indexOf("servicedate") : hdr.indexOf("service date"));
    if(!sermon && sIdx>=0) sermon = aoa[1]?.[sIdx] || "";
    if(!serviceDate && dIdx>=0){
      const d = excelSerialToDate(aoa[1]?.[dIdx]) || new Date(aoa[1]?.[dIdx]);
      if(!isNaN(d)) serviceDate = d;
    }
  }
  return { sermon, serviceDate };
}

// ---- announcements (hide older than 31 days) ----
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = firstSheetAOA(wb);
    if(aoa.length===0){ $("#announcements-table").textContent = "No announcements."; return; }

    const headers = aoa[0].map(h => String(h));
    const dateIdx = headers.findIndex(h => /date/i.test(h));
    const textIdx = headers.findIndex(h => /(announcement|details?)/i.test(h));
    const idxDate = dateIdx >= 0 ? dateIdx : 0;
    const idxText = textIdx >= 0 ? textIdx : 1;

    const cutoff = new Date(); cutoff.setHours(0,0,0,0); cutoff.setDate(cutoff.getDate() - 31);

    let rows = aoa.slice(1)
      .filter(r => r.some(c => String(c).trim()!==""))
      .map(r => {
        const d = excelSerialToDate(r[idxDate]) || new Date(r[idxDate]);
        const validDate = (d && !isNaN(d)) ? d : null;
        const ds = validDate ? niceDate(validDate) : String(r[idxDate]);
        const txt = normalizeAnnouncementText(r[idxText] || "");
        return { d: validDate, ds, txt };
      })
      .filter(r => r.d && r.d >= cutoff);

    rows.sort((a,b)=> b.d - a.d);

    if (rows.length === 0){
      $("#announcements-table").innerHTML = `<p class="dim">No announcements in the last 31 days.</p>`;
      return;
    }

    const pretty = [["DATE","ANNOUNCEMENT"], ...rows.map(r => [r.ds, r.txt])];
    renderTable(pretty, "#announcements-table");
  }catch(e){
    $("#announcements-table").innerHTML = `<p class="dim">Add <code>${PATHS.announcements}</code> with headers like: Date | Announcement.</p>`;
  }
}

// ---- setlists (This Coming Week & Last Week) ----
async function getDatedFiles(){
  const items = await listDir(PATHS.setlistsDir);
  return items
    .filter(it => it.type==="file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(it.name) && toDateFromName(it.name))
    .map(it => ({ name: it.name, date: toDateFromName(it.name) }))
    .sort((a,b)=> a.date - b.date); // oldest -> newest
}

async function loadSetlists(){
  const dated = await getDatedFiles();
  const specialExists = await existsOnGitHub(PATHS.specialCurrent);

  let nextRendered = false;
  let specialMeta = null;

  if (specialExists) {
    try {
      const wb = await fetchWB(PATHS.specialCurrent);
      const aoa = firstSheetAOA(wb);
      const { sermon, serviceDate } = parseMeta(wb, aoa);

      $("#next-date").textContent = serviceDate ? niceDate(serviceDate) : niceDate(new Date());
      renderTable(aoa, "#next-setlist");
      renderSermon("#next-sermon", sermon);
      nextRendered = true;

      let last = null;
      if (serviceDate) {
        const before = dated.filter(f => f.date < serviceDate);
        last = before[before.length - 1] || null;
      } else {
        last = dated[dated.length - 1] || null;
      }

      if (last) {
        $("#last-date").textContent = niceDate(last.date);
        const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
        const aoa2 = firstSheetAOA(wb2);
        renderTable(aoa2, "#last-setlist");
        const meta2 = parseMeta(wb2, aoa2);
        renderSermon("#last-sermon", meta2.sermon);
      } else {
        $("#last-date").textContent = "—";
        $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
      }

      specialMeta = { wb, aoa, serviceDate };
    } catch {
      nextRendered = false;
    }
  }

  if (!nextRendered) {
    if (dated.length === 0) {
      $("#next-date").textContent = "No setlists yet.";
      $("#next-setlist").innerHTML = `<p class="dim">Upload weekly files to <code>${PATHS.setlistsDir}/YYYY-MM-DD.xlsx</code> or provide <code>${PATHS.specialCurrent}</code>.</p>`;
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">—</p>`;
      return { dated, specialMeta };
    }

    const today = new Date(); today.setHours(0,0,0,0);
    let nextIdx = dated.findIndex(f => f.date >= today);
    if (nextIdx === -1) nextIdx = dated.length - 1;

    const next = dated[nextIdx];
    const wb = await fetchWB(`${PATHS.setlistsDir}/${next.name}`);
    const aoa = firstSheetAOA(wb);
    renderTable(aoa, "#next-setlist");
    $("#next-date").textContent = niceDate(next.date);
    const meta = parseMeta(wb, aoa);
    renderSermon("#next-sermon", meta.sermon);

    const last = dated[nextIdx - 1] || (dated.length >= 2 ? dated[dated.length - 2] : null);
    if (last) {
      $("#last-date").textContent = niceDate(last.date);
      const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
      const aoa2 = firstSheetAOA(wb2);
      renderTable(aoa2, "#last-setlist");
      const meta2 = parseMeta(wb2, aoa2);
      renderSermon("#last-sermon", meta2.sermon);
    } else {
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
    }
  }

  return { dated, specialMeta };
}

// ---- members (flexible columns) ----
async function loadMembers(){
  try{
    const wb = await fetchWB(PATHS.members);
    const aoa = firstSheetAOA(wb);
    if(!aoa || aoa.length === 0){
      $("#members-table").textContent = "No members listed.";
      return;
    }

    // Prefer common columns ordering if present
    const headers = aoa[0].map(h=>String(h));
    const preferred = ["Name","Role","Section","Instrument","Part","Phone","Email","Notes"];
    const lowerHeaders = headers.map(h=>h.toLowerCase());

    const orderIdx = [];
    preferred.forEach(p=>{
      const idx = lowerHeaders.indexOf(p.toLowerCase());
      if(idx>=0) orderIdx.push(idx);
    });
    headers.forEach((_,i)=>{ if(!orderIdx.includes(i)) orderIdx.push(i); });

    const ordered = aoa.map(row => orderIdx.map(i => row[i] ?? ""));

    renderTable(ordered, "#members-table");
  }catch(e){
    $("#members-table").innerHTML = `<p class="dim">Unable to load members. Ensure <code>${PATHS.members}</code> exists.</p>`;
  }
}

// ---- analytics (Current Year, no keys) + charts ----
async function buildAnalytics(dated, specialMeta){
  const currentYear = new Date().getFullYear();
  const datedThisYear = (dated || []).filter(
    f => f.date && f.date.getFullYear() === currentYear
  );

  const songCounts = new Map();  // title -> plays

  async function accumulateFromWB(wb){
    const aoa = firstSheetAOA(wb);
    if(!aoa || aoa.length===0) return;
    const headers = aoa[0].map(h => String(h).toLowerCase());
    const ti = headers.indexOf("song") !== -1 ? headers.indexOf("song") : headers.indexOf("title");
    if(ti === -1) return;

    for(const row of aoa.slice(1)){
      const title = (row[ti] || "").toString().trim();
      if(!title) continue;
      songCounts.set(title, (songCounts.get(title)||0) + 1);
    }
  }

  const includeSpecial =
    !!specialMeta && (
      !specialMeta.serviceDate ||
      specialMeta.serviceDate.getFullYear() === currentYear
    );

  if (datedThisYear.length === 0 && includeSpecial) {
    await accumulateFromWB(specialMeta.wb);
  } else {
    for (const f of datedThisYear) {
      const wb = await fetchWB(`${PATHS.setlistsDir}/${f.name}`);
      await accumulateFromWB(wb);
    }
    if (includeSpecial) await accumulateFromWB(specialMeta.wb);
  }

  if (songCounts.size === 0) {
    $("#top5").innerHTML = `<li class="dim">No data yet for ${currentYear}.</li>`;
    $("#bottom5").innerHTML = `<li class="dim">No data yet for ${currentYear}.</li>`;
    $("#library-table").innerHTML = `<p class="dim">No setlists found for ${currentYear}.</p>`;
    renderCharts([]);
    return;
  }

  const entries = Array.from(songCounts.entries()); // [title, count]
  entries.sort((a,b)=> b[1]-a[1] || a[0].localeCompare(b[0]));

  const top5 = entries.slice(0,5);
  $("#top5").innerHTML = top5
    .map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`)
    .join("");

  const bottom5 = entries.slice().sort((a,b)=> a[1]-b[1] || a[0].localeCompare(b[0])).slice(0,5);
  $("#bottom5").innerHTML = bottom5
    .map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`)
    .join("");

  // Library (current year) — Song + Plays only
  const header = ["Song","Plays"];
  let html = "<table><tr>" + header.map(h=>`<th>${h}</th>`).join("") + "</tr>";
  const titles = Array.from(songCounts.keys()).sort((a,b)=> a.localeCompare(b));
  for (const t of titles) {
    const plays = songCounts.get(t) || 0;
    html += `<tr><td>${t}</td><td>${plays}</td></tr>`;
  }
  html += "</table>";
  $("#library-table").innerHTML = html;

  // Charts
  renderCharts(entries);
}

// ---- charts (Chart.js) ----
let barChart, pieChart;
function renderCharts(entries){
  const barEl = document.getElementById("barChart");
  const pieEl = document.getElementById("pieChart");

  if (barChart) { barChart.destroy(); barChart = null; }
  if (pieChart) { pieChart.destroy(); pieChart = null; }

  if (!entries || entries.length === 0){
    if (barEl) barEl.replaceWith(barEl.cloneNode(true));
    if (pieEl) pieEl.replaceWith(pieEl.cloneNode(true));
    return;
  }

  const topN = entries.slice(0,10);
  const labels = topN.map(([t])=>t);
  const counts = topN.map(([,c])=>c);

  if (barEl){
    barChart = new Chart(barEl.getContext("2d"), {
      type: "bar",
      data: { labels, datasets: [{ label: "Plays (Top 10)", data: counts }] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: { y: { beginAtZero: true, ticks: { precision: 0 } } },
        plugins: { legend: { display: false } }
      }
    });
  }

  const pieN = entries.slice(0,8);
  const pieLabels = pieN.map(([t])=>t);
  const pieCounts = pieN.map(([,c])=>c);

  if (pieEl){
    pieChart = new Chart(pieEl.getContext("2d"), {
      type: "pie",
      data: { labels: pieLabels, datasets: [{ data: pieCounts }] },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }
}

// ---- boot ----
document.addEventListener("DOMContentLoaded", async ()=>{
  try{
    await loadAnnouncements();
    await loadMembers();
    const { dated, specialMeta } = await loadSetlists();
    await buildAnalytics(dated, specialMeta);
  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<p class="dim">Unable to load announcements. Ensure <code>${PATHS.announcements}</code> exists.</p>`;
    $("#members-table").innerHTML = `<p class="dim">Unable to load members. Ensure <code>${PATHS.members}</code> exists.</p>`;
    $("#next-date").textContent = "—";
    $("#next-setlist").innerHTML = `<p class="dim">Unable to load setlists. Ensure files exist in <code>${PATHS.setlistsDir}/</code> or <code>${PATHS.specialCurrent}</code>.</p>`;
    $("#last-setlist").innerHTML = `<p class="dim">—</p>`;
    $("#top5").innerHTML = `<li class="dim">No data.</li>`;
    $("#bottom5").innerHTML = `<li class="dim">No data.</li>`;
    $("#library-table").innerHTML = `<p class="dim">No data.</p>`;
    renderCharts([]);
  }
});
