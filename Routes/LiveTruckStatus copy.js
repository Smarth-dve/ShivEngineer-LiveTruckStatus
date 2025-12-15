const express = require("express");
const router = express.Router();
const sql = require("mssql/msnodesqlv8");
const escapeHtml = require("escape-html");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const dbConfig = require("../Config/dbConfig");
const PdfPrinter = require("pdfmake");

/* ================= HELPERS ================= */

function buildWhere(req, r, tags = []) {
  let w = "WHERE 1=1";

  if (req.query.search) {
    w += " AND (TRUCK_REG_NO LIKE @s OR CARD_NO LIKE @s OR CUSTOMER_NAME LIKE @s)";
    r.input("s", sql.VarChar, `%${req.query.search}%`);
    tags.push(`Search-${req.query.search}`);
  }

  if (req.query.bay) {
    w += " AND BAY_NO=@b";
    r.input("b", sql.Int, req.query.bay);
    tags.push(`Bay-${req.query.bay}`);
  }

  if (req.query.processType) {
    w += " AND PROCESS_TYPE=@p";
    r.input("p", sql.Int, req.query.processType);
    tags.push(`Type-${req.query.processType}`);
  }

  return w;
}

function statusMap(v) {
  switch (v) {
    case 0: return { t: "Pending", c: "pending" };
    case 1: return { t: "Completed", c: "completed" };
    case 2: return { t: "In Progress", c: "in-progress" };
    case 16: return { t: "In Process", c: "in-process" };
    default: return { t: String(v), c: "unknown" };
  }
}

/* ================= MAIN PAGE ================= */

router.get("/LiveTruckStatus", async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);

    const page = +req.query.page || 1;
    const limit = 20;
    const offset = (page - 1) * limit;

    const cr = pool.request();
    const wc = buildWhere(req, cr);
    const total = (await cr.query(`SELECT COUNT(*) t FROM COMMON_VIEW ${wc}`))
      .recordset[0].t;
    const pages = Math.ceil(total / limit);

    const dr = pool.request();
    const wd = buildWhere(req, dr);

    const rs = await dr.query(`
      SELECT
        TRUCK_REG_NO, CARD_NO, PROCESS_STATUS, BAY_NO,
        CUSTOMER_NAME, ITEM_DESCRIPTION, NET_WEIGHT, EXIT_GATE_TIME,
        (SELECT cv.* FROM COMMON_VIEW cv
         WHERE cv.TRUCK_REG_NO = main.TRUCK_REG_NO
           AND cv.CARD_NO = main.CARD_NO
         FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS FULL_ROW
      FROM COMMON_VIEW main
      ${wd}
      ORDER BY FAN_TIME_OUT DESC
      OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY
    `);

    const hasData = rs.recordset.length > 0;

    let html = `
<!DOCTYPE html>
<html>
<head>
<title>Live Truck Status</title>
<link rel="stylesheet" href="/Css/Home.css">
<link rel="stylesheet" href="/Css/Page.css">
<link rel="stylesheet" href="/Css/LiveTruckStatus.css">
<link href="https://fonts.googleapis.com/css?family=DM Sans" rel="stylesheet">
</head>
<body>

${fs.readFileSync(path.join(__dirname, "../public/Css/navbar.html"), "utf8")}

<h2>LIVE TRUCK STATUS</h2>

<form class="filter-bar">
<input name="search" placeholder="Truck / Card / Customer" value="${escapeHtml(req.query.search || "")}">
<select name="bay"><option value="">All Bays</option><option>1</option><option>2</option><option>3</option><option>4</option></select>
<select name="processType"><option value="">All Types</option><option value="1">Loading</option><option value="0">Unloading</option></select>
<button>Apply</button>
<a href="/LiveTruckStatus">Reset</a>
</form>

<div class="export-bar">
  <div class="right-actions">
    <button class="dark-toggle" onclick="toggleDarkMode()" id="darkBtn">
  <span id="darkIcon">🌙</span>
</button>

<a href="#"
   class="btn excel ${!hasData ? "disabled" : ""}"
   ${!hasData ? 'onclick="return false"' : 'onclick="openExcelModal(event)"'}>
   Excel
</a>

<a href="#"
   class="btn pdf ${!hasData ? "disabled" : ""}"
   ${!hasData ? 'onclick="return false"' : 'onclick="openPdfModal(event)"'}>
   PDF
</a>

  </div>
</div>

<table>
<thead>
<tr>
<th>#</th><th>Truck</th><th>Card</th><th>Status</th>
<th>Bay</th><th>Customer</th><th>Product</th>
<th>Net Wt</th><th>Exit</th><th>Action</th>
</tr>
</thead>
<tbody>
`;

    rs.recordset.forEach((r, i) => {
      const st = statusMap(r.PROCESS_STATUS);
      html += `
<tr>
<td>${offset + i + 1}</td>
<td>${r.TRUCK_REG_NO}</td>
<td>${r.CARD_NO}</td>
<td class="status ${st.c}">${st.t}</td>
<td>${r.BAY_NO || ""}</td>
<td>${escapeHtml(r.CUSTOMER_NAME || "")}</td>
<td>${escapeHtml(r.ITEM_DESCRIPTION || "")}</td>
<td>${r.NET_WEIGHT || ""}</td>
<td>${r.EXIT_GATE_TIME || ""}</td>
<td>
<button class="view-btn" onclick='openModal(${escapeHtml(JSON.stringify(r.FULL_ROW))})'>View</button>
</td>
</tr>`;
    });

    html += `
</tbody></table>

<div class="pagination">
${page > 1 ? `<a class="page-btn" href="?page=${page - 1}">Prev</a>` : ""}
${Array.from({ length: pages }, (_, i) =>
  `<a class="page-btn ${page === i + 1 ? "active" : ""}" href="?page=${i + 1}">${i + 1}</a>`
).join("")}
${page < pages ? `<a class="page-btn" href="?page=${page + 1}">Next</a>` : ""}
</div>

<!-- VIEW MODAL -->
<div id="modal" class="modal">
  <div class="modal-content">
    <div class="modal-header">
      <h3>Truck Full Details</h3>
      <span class="modal-close" onclick="closeModal()">✕</span>
    </div>
    <div class="modal-body" id="modalBody"></div>
  </div>
</div>

<!-- EXPORT MODALS -->
${exportModalHtml("pdf")}
${exportModalHtml("excel")}

<script>
function closeAllModals(){
  document.querySelectorAll('.modal').forEach(m=>m.style.display='none');
}
document.addEventListener('keydown',e=>{
  if(e.key==='Escape') closeAllModals();
});

function openModal(data){
  const obj = JSON.parse(data);
  let h='<table class="detail-table">';
  for(const k in obj) h+=\`<tr><th>\${k}</th><td>\${obj[k]}</td></tr>\`;
  h+='</table>';
  modalBody.innerHTML=h;
  modal.style.display='block';
}
function closeModal(){ modal.style.display='none'; }

function openPdfModal(e){ openExportModal(e,'pdfModal'); }
function openExcelModal(e){ openExportModal(e,'excelModal'); }

function openExportModal(e,id){
  e.preventDefault();
  closeAllModals();
  const m=document.getElementById(id);
  m.style.display='block';
  const r=e.target.getBoundingClientRect();
  m.querySelector('.modal-content').style.marginTop=(r.bottom+window.scrollY+10)+'px';
}

function exportData(type,scope){
  const p=new URLSearchParams(window.location.search);
  if(scope==='page') p.set('scope','page');
  else { p.set('scope','all'); p.delete('page'); }
  closeAllModals();
  location.href='/LiveTruckStatus/'+type+'?'+p.toString();
}

function toggleDarkMode(){
  const body = document.body;
  const icon = document.getElementById('darkIcon');

  body.classList.toggle('dark');

  if(body.classList.contains('dark')){
    icon.textContent = '☀️';
    localStorage.setItem('lts_dark','1');
  } else {
    icon.textContent = '🌙';
    localStorage.removeItem('lts_dark');
  }
}

// restore on load
(function(){
  if(localStorage.getItem('lts_dark')){
    document.body.classList.add('dark');
    const icon = document.getElementById('darkIcon');
    if(icon) icon.textContent = '☀️';
  }
})();
</script>

</body></html>`;

    res.send(html);
  } catch (e) {
    console.error(e);
    res.status(500).send("Live Truck Status Error");
  }
});

/* ================= EXCEL EXPORT ================= */

router.get("/LiveTruckStatus/excel", async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);
    const r = pool.request();
    const tags = [];

    const where = buildWhere(req, r, tags);

    const limit = 20;
    const page = +req.query.page || 1;
    const offset = (page - 1) * limit;
    let paging = "";

    if (req.query.scope === "page") {
      paging = ` OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY `;
      tags.push(`Page-${page}`);
    }

    const rs = await r.query(`
      SELECT * FROM COMMON_VIEW
      ${where}
      ORDER BY FAN_TIME_OUT DESC
      ${paging}
    `);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Live Truck Status");

    ws.columns = Object.keys(rs.recordset[0]).map(k=>({
      header:k.replace(/_/g," "),
      key:k,
      width:22
    }));

    rs.recordset.forEach(r=>ws.addRow(r));

    ws.getRow(1).eachCell(c=>{
      c.font={bold:true,color:{argb:"FFFFFFFF"}};
      c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1F4E78"}};
    });

    const d = new Date().toLocaleDateString("en-GB").split("/").join("_");
    const f = tags.length ? "_"+tags.join("_") : "";

    res.setHeader("Content-Disposition",`attachment; filename="LiveTruckStatus${f}_${d}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error("Excel error:", e);
    res.status(500).send("Excel export failed");
  }
});

/* ================= PDF EXPORT ================= */

router.get("/LiveTruckStatus/pdf", async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);
    const r = pool.request();
    const tags = [];

    const where = buildWhere(req, r, tags);

    const limit = 20;
    const page = +req.query.page || 1;
    const offset = (page - 1) * limit;
    let paging = "";

    if (req.query.scope === "page") {
      paging = ` OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY `;
      tags.push(`Page-${page}`);
    }

    const rs = await r.query(`
      SELECT * FROM COMMON_VIEW
      ${where}
      ORDER BY FAN_TIME_OUT DESC
      ${paging}
    `);

    const cols = Object.keys(rs.recordset[0]);
    const widths = cols.map(c =>
      c.includes("TIME")||c.includes("DATE")?80:
      c.includes("NAME")||c.includes("ADDRESS")?100:
      c.includes("WEIGHT")?70:55
    );

    const body = [ cols.map(c=>({text:c.replace(/_/g," "),style:"th"})) ];
    rs.recordset.forEach(r=>{
      body.push(cols.map(c=>String(r[c]??"")));
    });

    const printer = new PdfPrinter({Helvetica:{normal:"Helvetica",bold:"Helvetica-Bold"}});
    const doc = {
      pageSize:"A3",
      pageOrientation:"landscape",
      defaultStyle:{font:"Helvetica",fontSize:8},
      footer:(p,t)=>({text:`Page ${p} of ${t}`,alignment:"right",margin:[0,0,20,0]}),
      content:[
        {text:"Live Truck Status Report",fontSize:16,bold:true},
        {text:`Generated on: ${new Date().toLocaleString()}`,margin:[0,5,0,10]},
        {table:{headerRows:1,widths,body},layout:"lightHorizontalLines"}
      ],
      styles:{th:{bold:true,fillColor:"#1F4E78",color:"white"}}
    };

    const d = new Date().toLocaleDateString("en-GB").split("/").join("_");
    const f = tags.length ? "_"+tags.join("_") : "";

    res.setHeader("Content-Disposition",`attachment; filename="LiveTruckStatus${f}_${d}.pdf"`);
    res.setHeader("Content-Type","application/pdf");

    const pdf = printer.createPdfKitDocument(doc);
    pdf.pipe(res);
    pdf.end();
  } catch (e) {
    console.error("PDF error:", e);
    res.status(500).send("PDF export failed");
  }
});

function exportModalHtml(type){
  return `
<div id="${type}Modal" class="modal">
  <div class="modal-content pdf-modal">
    <div class="modal-header">
      <h3>Export ${type.toUpperCase()}</h3>
      <span class="modal-close" onclick="closeAllModals()">✕</span>
    </div>
    <div class="modal-body pdf-options">
      <button class="btn ${type}" onclick="exportData('${type}','page')">Current Page</button>
      <button class="btn ${type}" onclick="exportData('${type}','all')">All Pages</button>
    </div>
  </div>
</div>`;
}

module.exports = router;
