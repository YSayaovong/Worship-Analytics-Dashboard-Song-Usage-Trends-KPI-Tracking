/* HFBC – Praise & Worship (Static, GitHub-hosted)
   Pages:
     - index.html (Home): Announcements, Reminders, Bible Study (4 weeks), Members
     - reporting.html: Weekly Songs (+stories), Analytics, CCLI buckets, CSV export
     - hymnal.html: FlipHTML5 embed
*/

const GH = { owner: "YSayaovong", repo: "HFBC_Praise_Worship", branch: "main" };
const PATHS = {
  specialCurrent: "setlist/setlist.xlsx",
  setlistsDir: "setlists",
  announcements: "announcements/announcements.xlsx",
  members: "members/members.xlsx",
  bibleStudy: "bible_study/bible_study.xlsx"
};

// ---------- tiny DOM helpers ----------
const $ = (s, r=document) => r.querySelector(s);

// ---------- GitHub fetch helpers ----------
const apiURL = (p) => `https://api.github.com/repos/${GH.owner}/${GH.repo}/contents/${encodeURIComponent(p)}?ref=${encodeURIComponent(GH.branch)}`;
const rawURL = (p) => `https://raw.githubusercontent.com/${GH.owner}/${GH.repo}/${GH.branch}/${p}`;

async function listDir(path){
  const r = await fetch(apiURL(path), { headers:{ "Accept":"application/vnd.github+json" }});
  if(!r.ok) return [];
  return r.json();
}
async function fetchWB(path){
  const r = await fetch(rawURL(path));
  if(!r.ok) throw new Error(`Fetch error ${r.status} for ${path}`);
  const ab = await r.arrayBuffer();
  return XLSX.read(ab, { type:"array" });
}
function firstSheetAOA(wb){
  const sh = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sh, { header:1, defval:"" });
}
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
function niceDate(d){ return d.toLocaleDateString(undefined,{ month:"short", day:"numeric", year:"numeric" }); }
function normalizeAnnouncementText(s){
  let out = String(s ?? "");
  out = out.replace(/:(?=\d)/g, ": ");
  out = out.replace(/([A-Za-z])(\d{1,2}\/\d{1,2}\/\d{2,4})/g, "$1 $2");
  return out.trim();
}

// ---------- generic table renderer ----------
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
      let out = String(cell ?? "");
      if(rIdx>0 && dateCols.has(cIdx)) out = prettyDateCell(cell);
      html += (rIdx===0? `<th>${out}</th>` : `<td>${out}</td>`);
    });
    html += "</tr>";
  });
  html += "</table>";
  el.innerHTML = html;
}

// ---------- announcements (hide > 31 days) ----------
async function loadAnnouncements(){
  const target = "#announcements-table";
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = firstSheetAOA(wb);
    if(aoa.length===0){ $(target).textContent = "No announcements."; return; }

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
      .filter(r => r.d && r.d >= cutoff)
      .sort((a,b)=> b.d - a.d);

    if (rows.length === 0){
      $(target).innerHTML = `<p class="dim">No announcements in the last 31 days.</p>`;
      return;
    }

    const pretty = [["DATE","ANNOUNCEMENT"], ...rows.map(r => [r.ds, r.txt])];
    renderTable(pretty, target);
  }catch(e){
    $(target).innerHTML = `<p class="dim">Add <code>${PATHS.announcements}</code> with headers like: Date | Announcement.</p>`;
  }
}

// ---------- bible study (last 4 weeks) ----------
async function loadBibleStudy(){
  const target = "#bible-study-table";
  try{
    const wb = await fetchWB(PATHS.bibleStudy);
    const aoa = firstSheetAOA(wb);
    if(!aoa || aoa.length === 0){ $(target).textContent = "No Bible study items."; return; }

    const headers = aoa[0].map(h => String(h));
    const di = headers.findIndex(h => /date/i.test(h));
    if (di === -1){ renderTable(aoa, target); return; }

    const today = new Date(); today.setHours(23,59,59,999);
    const fourWeeksAgo = new Date(today); fourWeeksAgo.setDate(today.getDate() - 28); fourWeeksAgo.setHours(0,0,0,0);

    const rows = aoa
      .slice(1)
      .filter(r => r.some(c => String(c).trim()!==""))
      .map(r => ({ r, d: excelSerialToDate(r[di]) || new Date(r[di]) }))
      .filter(x => x.d && !isNaN(x.d) && x.d >= fourWeeksAgo && x.d <= today)
      .sort((a,b) => b.d - a.d)
      .map(x => x.r);

    if (rows.length === 0){
      $(target).innerHTML = `<p class="dim">No Bible study items in the past 4 weeks.</p>`;
      return;
    }
    renderTable([headers, ...rows], target);
  }catch(e){
    $(target).innerHTML = `<p class="dim">Unable to load. Ensure <code>${PATHS.bibleStudy}</code> exists.</p>`;
  }
}

// ---------- setlists & analytics helpers ----------
function toDateFromName(name){
  const m = /^(\d{4})-(\d{2})-(\d{2})\.xlsx$/.exec(name);
  return m ? new Date(`${m[1]}-${m[2]}-${m[3]}T00:00:00`) : null;
}
async function getDatedFiles(){
  const items = await listDir("setlists");
  return items
    .filter(it => it.type==="file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(it.name) && toDateFromName(it.name))
    .map(it => ({ name: it.name, date: toDateFromName(it.name) }))
    .sort((a,b)=> a.date - b.date);
}
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

// ---------- stories (Wikipedia) ----------
const STORY_TTL_MS = 1000 * 60 * 60 * 24 * 30; // 30 days
function getStoryCache(){ try{ return JSON.parse(localStorage.getItem("songStoryCache")||"{}"); }catch{ return {}; } }
function setStoryCache(cache){ localStorage.setItem("songStoryCache", JSON.stringify(cache)); }
async function fetchWikipediaStory(title){
  const key = title.trim().toLowerCase();
  const cache = getStoryCache();
  const now = Date.now();
  if (cache[key] && (now - cache[key].ts) < STORY_TTL_MS) return cache[key].data;

  const q = `${title} song hymn meaning`;
  try{
    const sr = await fetch(`https://en.wikipedia.org/api/rest_v1/search/page?q=${encodeURIComponent(q)}&limit=5`);
    if(sr.ok){
      const sdata = await sr.json();
      const candidate = (sdata?.pages||[])[0];
      if(candidate && candidate.title){
        const sum = await fetch(`https://en.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(candidate.title)}`);
        if(sum.ok){
          const j = await sum.json();
          const data = {
            summary: j.extract || "",
            url: j.content_urls?.desktop?.page || `https://en.wikipedia.org/wiki/${encodeURIComponent(candidate.title)}`
          };
          cache[key] = { ts: now, data }; setStoryCache(cache);
          return data;
        }
      }
    }
  }catch(_e){ /* ignore */ }
  const empty = { summary: "", url: "" };
  cache[key] = { ts: now, data: empty }; setStoryCache(cache);
  return empty;
}
function storyCellHTML(text, url){
  if(!text && !url) return `<span class="story-cell dim">—</span>`;
  const short = text && text.length>220 ? (text.slice(0,217) + "…") : (text || "—");
  const link = url ? ` <a href="${url}" target="_blank" rel="noopener">Source</a>` : "";
  return `<span class="story-cell">${short}${link}</span>`;
}
async function renderSetlistWithStories(containerSel, aoa){
  const el = $(containerSel);
  if(!aoa || aoa.length===0){ el.innerHTML = `<p class="dim">No data.</p>`; return; }

  const headers = aoa[0].map(h => String(h));
  const lower = headers.map(h=>h.toLowerCase());
  const songIdx = lower.indexOf("song") !== -1 ? lower.indexOf("song") : lower.indexOf("title");
  const storyIdxExcel = lower.indexOf("story");

  const finalHeaders = [...headers, "Story"];
  let html = "<table><tr>" + finalHeaders.map(h=>`<th class="small">${h}</th>`).join("") + "</tr>";

  const rowsMeta = [];
  aoa.slice(1).forEach((row, rIdx)=>{
    const cells = row.map(c => `<td class="small">${String(c ?? "")}</td>`).join("");
    const placeholderId = `${containerSel.replace("#","")}-story-${rIdx}`;
    const storyProvided = storyIdxExcel >= 0 ? String(row[storyIdxExcel]||"").trim() : "";
    const storyCell = storyProvided
      ? storyCellHTML(storyProvided, "")
      : `<span id="${placeholderId}" class="story-cell dim">Searching…</span>`;
    html += `<tr>${cells}<td>${storyCell}</td></tr>`;
    rowsMeta.push({ title: songIdx>=0 ? (row[songIdx]||"").toString().trim() : "", placeholderId, storyProvided });
  });
  html += "</table>";
  el.innerHTML = html;

  for(const m of rowsMeta){
    if(!m.title || m.storyProvided) continue;
    const { summary, url } = await fetchWikipediaStory(m.title);
    const target = document.getElementById(m.placeholderId);
    if(target){ target.outerHTML = storyCellHTML(summary, url); }
  }
}

// ---------- load setlists into Reporting page ----------
async function loadSetlistsIntoReporting(){
  const items = await listDir(PATHS.setlistsDir);
  const dated = items
    .filter(it => it.type==="file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(it.name) && toDateFromName(it.name))
    .map(it => ({ name: it.name, date: toDateFromName(it.name) }))
    .sort((a,b)=> a.date - b.date);

  let nextRendered = false;

  // Prefer special current week if provided
  try{
    const r = await fetch(rawURL(PATHS.specialCurrent), { method:"HEAD" });
    if (r.ok){
      const wb = await fetchWB(PATHS.specialCurrent);
      const aoa = firstSheetAOA(wb);
      const { sermon, serviceDate } = parseMeta(wb, aoa);
      $("#next-date").textContent = serviceDate ? niceDate(serviceDate) : niceDate(new Date());
      await renderSetlistWithStories("#next-setlist", aoa);
      $("#next-sermon").textContent = sermon ? `Sermon: ${sermon}` : "";

      // Last week from archive
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
        await renderSetlistWithStories("#last-setlist", aoa2);
        const meta2 = parseMeta(wb2, aoa2);
        $("#last-sermon").textContent = meta2.sermon ? `Sermon: ${meta2.sermon}` : "";
      } else {
        $("#last-date").textContent = "—";
        $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
      }

      nextRendered = true;
    }
  }catch(_e){ /* ignore */ }

  if (!nextRendered) {
    if (dated.length === 0) {
      $("#next-date").textContent = "No setlists yet.";
      $("#next-setlist").innerHTML = `<p class="dim">Upload weekly files to <code>${PATHS.setlistsDir}/YYYY-MM-DD.xlsx</code> or provide <code>${PATHS.specialCurrent}</code>.</p>`;
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">—</p>`;
      return { dated, specialMeta: null };
    }

    const today = new Date(); today.setHours(0,0,0,0);
    let nextIdx = dated.findIndex(f => f.date >= today);
    if (nextIdx === -1) nextIdx = dated.length - 1;

    const next = dated[nextIdx];
    const wb = await fetchWB(`${PATHS.setlistsDir}/${next.name}`);
    const aoa = firstSheetAOA(wb);
    await renderSetlistWithStories("#next-setlist", aoa);
    $("#next-date").textContent = niceDate(next.date);
    const meta = parseMeta(wb, aoa);
    $("#next-sermon").textContent = meta.sermon ? `Sermon: ${meta.sermon}` : "";

    const last = dated[nextIdx - 1] || (dated.length >= 2 ? dated[dated.length - 2] : null);
    if (last) {
      $("#last-date").textContent = niceDate(last.date);
      const wb2 = await fetchWB(`${PATHS.setlistsDir}/${last.name}`);
      const aoa2 = firstSheetAOA(wb2);
      await renderSetlistWithStories("#last-setlist", aoa2);
      const meta2 = parseMeta(wb2, aoa2);
      $("#last-sermon").textContent = meta2.sermon ? `Sermon: ${meta2.sermon}` : "";
    } else {
      $("#last-date").textContent = "—";
      $("#last-setlist").innerHTML = `<p class="dim">No prior week found.</p>`;
    }
  }

  return { dated, specialMeta: null };
}

// ---------- analytics + CCLI status + CSV export (Reporting page) ----------
async function buildAnalyticsAndCCLI(dated, specialMeta){
  const currentYear = new Date().getFullYear();
  const PD_YEAR_CUTOFF = currentYear - 96;

  // If 'dated' not provided, recompute from folder
  if (!dated){
    const items = await listDir("setlists");
    dated = items
      .filter(it => it.type==="file" && /^\d{4}-\d{2}-\d{2}\.xlsx$/.test(it.name) && toDateFromName(it.name))
      .map(it => ({ name: it.name, date: toDateFromName(it.name) }))
      .sort((a,b)=> a.date - b.date);
  }
  const datedThisYear = (dated || []).filter(f => f.date && f.date.getFullYear() === currentYear);

  const songInfo = new Map(); // title -> { plays, ccli, hasCCLI, pdFlag, year }

  function coerceYear(v){
    const n = Number(String(v||"").match(/\d{4}/)?.[0] || NaN);
    return Number.isFinite(n) ? n : null;
  }

  async function accumulateFromWB(wb){
    const aoa = firstSheetAOA(wb);
    if(!aoa || aoa.length===0) return;

    const headers = aoa[0].map(h => String(h).toLowerCase());
    const ti = headers.indexOf("song") !== -1 ? headers.indexOf("song") : headers.indexOf("title");
    if(ti === -1) return;
    const ccliIdx = headers.findIndex(h => /ccli/.test(h));
    const pdIdx   = headers.findIndex(h => /(public\s*domain|ispublicdomain|\bpd\b)/.test(h));
    const yrIdx   = headers.findIndex(h => /^year|yearpublished/.test(h));

    for(const row of aoa.slice(1)){
      const title = (row[ti] || "").toString().trim();
      if(!title) continue;
      const info = songInfo.get(title) || { plays:0, ccli:"", hasCCLI:false, pdFlag:false, year:null };
      info.plays += 1;
      if (ccliIdx >= 0 && row[ccliIdx] != null){
        const c = String(row[ccliIdx]).trim();
        if (c) { info.ccli ||= c; info.hasCCLI = info.hasCCLI || /\d+/.test(c) || c.length>0; }
      }
      if (pdIdx >= 0 && row[pdIdx] != null){
        if (["1","true","yes","y","public domain","pd"].includes(String(row[pdIdx]).trim().toLowerCase())) info.pdFlag = true;
      }
      if (yrIdx >= 0 && row[yrIdx] != null){
        const y = coerceYear(row[yrIdx]);
        if (y && !info.year) info.year = y;
      }
      songInfo.set(title, info);
    }
  }

  // Aggregate this year's archive
  for (const f of datedThisYear) {
    const wb = await fetchWB(`${PATHS.setlistsDir}/${f.name}`);
    await accumulateFromWB(wb);
  }

  // Rankings
  const entries = Array.from(songInfo.entries()).map(([t,i])=>[t,i.plays]).sort((a,b)=> b[1]-a[1] || a[0].localeCompare(b[0]));
  const top5 = entries.slice(0,5);
  $("#top5").innerHTML = top5.map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`).join("");
  const bottom5 = entries.slice().sort((a,b)=> a[1]-b[1] || a[0].localeCompare(b[0])).slice(0,5);
  $("#bottom5").innerHTML = bottom5.map(([t,c]) => `<li><strong>${t}</strong> — ${c} play${c>1?"s":""}</li>`).join("");

  function classify(info){
    const pdByFlag = info.pdFlag === true;
    const pdByYear = info.year && info.year <= PD_YEAR_CUTOFF;
    if (pdByFlag || pdByYear) return "Public Domain";
    if (info.hasCCLI) return "Report (Licensed)";
    return "Needs Review";
  }

  // Library table
  const titles = Array.from(songInfo.keys()).sort((a,b)=> a.localeCompare(b));
  let html = "<table><tr><th>Song</th><th>Plays</th><th>Status</th></tr>";
  for (const t of titles) {
    const info = songInfo.get(t);
    const status = classify(info);
    html += `<tr><td>${t}</td><td>${info.plays}</td><td>${status}</td></tr>`;
  }
  html += "</table>";
  $("#library-table").innerHTML = html;

  // Buckets + CSV export data
  const reportRows = [["Song","Plays","CCLI"]];
  const pdRows = [["Song","Plays","Basis"]];

  for (const t of titles){
    const info = songInfo.get(t);
    const status = classify(info);
    if (status === "Public Domain"){
      const basis = info.pdFlag ? "Explicit PD" : (info.year ? `Year ≤ ${PD_YEAR_CUTOFF}` : "PD heuristic");
      pdRows.push([t, String(info.plays), basis]);
    } else if (status === "Report (Licensed)") {
      reportRows.push([t, String(info.plays), info.ccli || ""]);
    }
  }

  renderTable(reportRows, "#ccli-report");
  renderTable(pdRows, "#ccli-pd");

  // Charts
  renderCharts(entries);

  // CSV export
  $("#export-csv")?.addEventListener("click", ()=>{
    const csv = reportRows.map(r => r.map(x => `"${String(x).replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `ccli_report_${currentYear}.csv`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  });
}

// ---------- charts (Chart.js) ----------
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
      options: { responsive:true, maintainAspectRatio:false, scales:{ y:{ beginAtZero:true, ticks:{ precision:0 } } }, plugins:{ legend:{ display:false } } }
    });
  }
  const pieN = entries.slice(0,8);
  const pieLabels = pieN.map(([t])=>t);
  const pieCounts = pieN.map(([,c])=>c);
  if (pieEl){
    pieChart = new Chart(pieEl.getContext("2d"), {
      type: "pie",
      data: { labels: pieLabels, datasets: [{ data: pieCounts }] },
      options: { responsive:true, maintainAspectRatio:false }
    });
  }
}

// ---------- Page dispatcher ----------
document.addEventListener("DOMContentLoaded", async ()=>{
  const page = document.body.getAttribute("data-page");

  if (page === "home"){
    await loadAnnouncements();
    await loadBibleStudy();
    // Members loaded on Home (kept here to ensure order)
    const target = "#members-table";
    try{
      const wb = await fetchWB(PATHS.members);
      const aoa = firstSheetAOA(wb);
      if(!aoa || aoa.length === 0){ $(target).textContent = "No members listed."; return; }
      renderTable(aoa, target);
    }catch{
      $(target).innerHTML = `<p class="dim">Unable to load members. Ensure <code>${PATHS.members}</code> exists.</p>`;
    }
  }

  if (page === "reporting"){
    const { dated, specialMeta } = await loadSetlistsIntoReporting();
    await buildAnalyticsAndCCLI(dated, specialMeta);
  }

  // No library page anymore
});
