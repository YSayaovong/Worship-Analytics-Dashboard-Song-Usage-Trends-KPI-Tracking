// --- Simple CSV loader ---
async function fetchCSV(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`Fetch failed: ${url} (${res.status})`);
  const text = await res.text();

  const rows = text.trim().split("\n").map((r) => r.split(","));
  const header = rows.shift();
  return rows.map((r) =>
    Object.fromEntries(header.map((h, i) => [h.trim(), (r[i] ?? "").trim()]))
  );
}

// Hymnal Coverage
fetchCSV("data/kpi_hymnal_coverage.csv").then((rows) => {
  const obj = Object.fromEntries(rows.map((r) => [r.metric, r.value]));
  document.getElementById("coverageUsed").textContent = obj.hymnal_coverage_used;
  document.getElementById("coverageUnused").textContent = obj.hymnal_coverage_unused;
  document.getElementById("coveragePercent").textContent = obj.hymnal_coverage_percent + "%";
});

// Usage by Source
fetchCSV("data/kpi_by_source.csv").then((rows) => {
  const div = document.getElementById("bySource");
  div.innerHTML = rows
    .map((r) => `${r.source_final}: ${r.count}`)
    .join(" | ");
});

// Top 10 Most Played
fetchCSV("data/kpi_top10.csv").then((rows) => {
  const ol = document.getElementById("top10List");
  ol.innerHTML = rows.map((r) => `<li>${r.title_norm} (${r.plays})</li>`).join("");
});

// Unused Hymnal Numbers
fetchCSV("data/hymnal_unused.csv").then((rows) => {
  const ul = document.getElementById("unusedList");
  ul.innerHTML = rows
    .slice(0, 50) // show first 50
    .map((r) => `<li>#${r.unused_number}</li>`)
    .join("");
});
