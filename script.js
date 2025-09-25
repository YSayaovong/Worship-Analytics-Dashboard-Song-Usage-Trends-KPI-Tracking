/* ===========================
   HFBC — Script (Additive)
   Keeps existing layout. Adds Data Analyst widgets that read /data/*.csv
   =========================== */

/* ---------- Small utilities ---------- */
function $(sel, root = document) { return root.querySelector(sel); }
function $all(sel, root = document) { return [...root.querySelectorAll(sel)]; }

function el(tag, opts = {}) {
  const n = document.createElement(tag);
  if (opts.className) n.className = opts.className;
  if (opts.text) n.textContent = opts.text;
  if (opts.html) n.innerHTML = opts.html;
  if (opts.attrs) Object.entries(opts.attrs).forEach(([k,v]) => n.setAttribute(k, v));
  return n;
}

/* ---------- Minimal CSV reader (works for simple KPI files) ---------- */
async function fetchCSV(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`Fetch failed: ${url} (${res.status})`);
  const text = await res.text();

  // basic CSV split—sufficient for our simple KPI files without commas in values
  const rows = text.trim().split("\n").map(r => r.split(","));
  const header = rows.shift();
  return rows.map(r =>
    Object.fromEntries(header.map((h, i) => [h.trim(), (r[i] ?? "").trim()]))
  );
}

/* ---------- Data Analyst UI block (non-destructive) ---------- */
function mountDataAnalystHost() {
  // Prefer to append under your existing “Song Analytics” card
  // Adjust this selector if your analytics card uses a different id/class
  const analyticsCard =
    $("#analytics-section .card") ||
    $("#song-analytics .card") ||
    $(".analytics .card");

  const mountPoint = analyticsCard || $("main") || document.body;

  // Container
  const host = el("section", { className: "da-host" });
  host.style.marginTop = "16px";
  host.innerHTML = `
    <h2 class="subhead">Data Analyst</h2>
    <div class="charts charts-2" id="da-grid">
      <div class="chart-card">
        <h3 class="subhead">Hymnal Coverage</h3>
        <div id="da-coverage" class="dim">Loading…</div>
      </div>
      <div class="chart-card">
        <h3 class="subhead">Usage by Source</h3>
        <div id="da-by-source" class="dim">Loading…</div>
      </div>
      <div class="chart-card" style="grid-column: 1 / -1;">
        <h3 class="subhead">Unused Hymnal Numbers</h3>
        <p class="dim" style="margin-top:-4px">Showing first 50 (ascending).</p>
        <div id="da-unused" class="like-card">Loading…</div>
      </div>
      <div class="chart-card" style="grid-column: 1 / -1;">
        <h3 class="subhead">Top 10 Most Played</h3>
        <ol id="da-top10" class="like-list" style="margin:0;"></ol>
      </div>
    </div>
  `;

  mountPoint.appendChild(host);
}

/* ---------- Renderers ---------- */
function renderCoverage(rows) {
  const obj = Object.fromEntries(rows.map(r => [r.metric, r.value]));
  const html = `
    <div class="table">
      <table>
        <tbody>
          <tr><th>Used</th><td>${obj.hymnal_coverage_used ?? "—"} / 352</td></tr>
          <tr><th>Unused</th><td>${obj.hymnal_coverage_unused ?? "—"}</td></tr>
          <tr><th>Coverage</th><td>${obj.hymnal_coverage_percent ?? "—"}%</td></tr>
        </tbody>
      </table>
    </div>
  `;
  $("#da-coverage").innerHTML = html;
}

function renderBySource(rows) {
  // rows like: [{source_final:"Hymnal", count:"45"}, ...]
  if (!rows?.length) {
    $("#da-by-source").textContent = "No data";
    return;
  }
  const html = `
    <div class="table">
      <table>
        <thead><tr><th>Source</th><th>Count</th></tr></thead>
        <tbody>
          ${rows.map(r => `<tr><td>${r.source_final}</td><td>${r.count}</td></tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
  $("#da-by-source").innerHTML = html;
}

function renderUnused(rows) {
  if (!rows?.length) {
    $("#da-unused").textContent = "All hymnal numbers have been used at least once.";
    return;
  }
  const list = rows
    .map(r => parseInt(r.unused_number, 10))
    .filter(n => !Number.isNaN(n))
    .sort((a,b) => a - b)
    .slice(0, 50);

  $("#da-unused").innerHTML = list.length
    ? `<div class="tag-row">${list.map(n => `<span class="tag">#${n}</span>`).join(" ")}</div>`
    : "All hymnal numbers have been used at least once.";
}

function renderTop10(rows) {
  const ol = $("#da-top10");
  if (!rows?.length) {
    ol.innerHTML = `<li class="dim">No data</li>`;
    return;
  }
  ol.innerHTML = rows
    .map(r => `<li>${r.title_norm || r.title || "Unknown"} <span class="dim">(${r.plays})</span></li>`)
    .join("");
}

/* ---------- Loader (pull weekly from /data/*.csv) ---------- */
async function loadDataAnalystKpis() {
  // Ensure host exists
  if (!$("#da-grid")) mountDataAnalystHost();

  // Parallel fetches
  const tasks = [
    fetchCSV("data/kpi_hymnal_coverage.csv").catch(() => []),
    fetchCSV("data/kpi_by_source.csv").catch(() => []),
    fetchCSV("data/hymnal_unused.csv").catch(() => []),
    fetchCSV("data/kpi_top10.csv").catch(() => []), // optional
  ];

  const [cov, bysrc, unused, top10] = await Promise.all(tasks);

  try { renderCoverage(cov); } catch(e) { console.warn("Coverage render:", e); }
  try { renderBySource(bysrc); } catch(e) { console.warn("BySource render:", e); }
  try { renderUnused(unused); } catch(e) { console.warn("Unused render:", e); }
  try { renderTop10(top10); } catch(e) { console.warn("Top10 render:", e); }
}

/* ---------- Boot ---------- */
document.addEventListener("DOMContentLoaded", () => {
  // Run after your existing page initializes.
  // If your other code also runs on DOMContentLoaded, both will execute.
  loadDataAnalystKpis();
});

/* ---------- Optional light styles (non-invasive) ----------
   Comment out if you prefer your own classes only.
*/
(function injectLightStyles() {
  const css = `
    .charts { display: grid; gap: 16px; }
    .charts-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    @media (max-width: 800px){ .charts-2 { grid-template-columns: 1fr; } }
    .chart-card { border: 1px solid #e5e7eb; border-radius: 12px; padding: 12px; background: #fff; }
    .subhead { margin: 0 0 8px; font-size: 1.1rem; }
    .dim { color: #6b7280; }
    .table table { width: 100%; border-collapse: collapse; }
    .table th, .table td { text-align: left; padding: 6px 8px; border-bottom: 1px solid #f0f0f0; }
    .tag-row { display: flex; flex-wrap: wrap; gap: 8px; }
    .tag { padding: 4px 8px; border-radius: 999px; border: 1px solid #e5e7eb; background: #f9fafb; font-size: 0.9rem; }
    .like-list { padding-left: 20px; }
  `;
  const style = el("style", { html: css });
  document.head.appendChild(style);
})();
