// --- Replace your existing loadAnnouncements() with this ---
async function loadAnnouncements(){
  try{
    const wb = await fetchWB(PATHS.announcements);
    const aoa = aoaFromWB(wb);
    if(!aoa || aoa.length === 0){ $("#announcements-table").innerHTML = `<p class="dim">No data.</p>`; return; }

    // Normalize headers
    const headersRaw = aoa[0].map(h => String(h).trim());
    const headers = headersRaw.map(h => h.toLowerCase());

    // Detect columns
    const idxDate = headers.findIndex(h => ["date","service date"].includes(h));

    // English
    const idxEn = [
      "announcement en","english","announcement (en)","announcement english","en"
    ].map(k => headers.indexOf(k)).find(i => i !== -1);

    // Hmong
    const idxHm = [
      "announcement hm","hmong","announcement (hmong)","announcement hmong","hm"
    ].map(k => headers.indexOf(k)).find(i => i !== -1);

    // If both language columns are present, build a 3-column bilingual table
    if(idxEn !== undefined && idxEn !== -1 && idxHm !== undefined && idxHm !== -1){
      const out = [["Date","English","Hmong"]];
      for(let i=1;i<aoa.length;i++){
        const row = aoa[i];
        if(!row || row.every(c => String(c ?? "").trim()==="")) continue;
        const d = idxDate !== -1 ? toLocalDate(row[idxDate]) : null;
        out.push([
          idxDate !== -1 ? (d ? fmtDate(d) : String(row[idxDate] ?? "")) : "",
          String(row[idxEn] ?? ""),
          String(row[idxHm] ?? "")
        ]);
      }
      renderAOATable(out, "#announcements-table");
      return;
    }

    // Otherwise, keep your previous behavior (format Date if present; render all columns)
    const idxDate2 = idxDate;
    const out2 = idxDate2 === -1 ? aoa : aoa.map((r, i) => {
      if(i===0) return r;
      const rr = r.slice();
      const d = toLocalDate(rr[idxDate2]);
      rr[idxDate2] = d ? fmtDate(d) : (rr[idxDate2] ?? "");
      return rr;
    });
    renderAOATable(out2, "#announcements-table");

  }catch(e){
    console.error(e);
    $("#announcements-table").innerHTML = `<p class="dim">Unable to load <code>${PATHS.announcements}</code>.</p>`;
  }
}
