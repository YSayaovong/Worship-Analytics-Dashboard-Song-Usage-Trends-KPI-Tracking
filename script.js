const SRC={announcements:"https://raw.githubusercontent.com/YSayaovong/Worship-Analytics-Dashboard-Song-Usage-Trends-KPI-Tracking/main/announcements/announcements.xlsx"};
function excelToDate(v){if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);return new Date(d.y,d.m-1,d.d);}return new Date(v);}
function fmt(d){return d.toDateString();}
async function fetchRows(url){const r=await fetch(url);const ab=await r.arrayBuffer();const wb=XLSX.read(ab,{type:'array'});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});}
async function renderAnnouncements(){const tb=document.getElementById("announcements-body");tb.innerHTML="";const rows=await fetchRows(SRC.announcements);rows.forEach(r=>{const d=excelToDate(r.DATE);const en=r.ANNOUNCEMENT;const hm=r["LUS TSHAJ TAWM"];const tr=document.createElement("tr");tr.innerHTML=`<td>${fmt(d)}</td><td>${en}</td><td>${hm}</td>`;tb.appendChild(tr);});}
document.addEventListener("DOMContentLoaded",renderAnnouncements);
