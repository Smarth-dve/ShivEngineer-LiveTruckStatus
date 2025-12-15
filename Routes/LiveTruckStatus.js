const express = require("express");
const router = express.Router();
const sql = require("mssql/msnodesqlv8");
const escapeHtml = require("escape-html");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");
const dbConfig = require("../Config/dbConfig");
const PdfPrinter = require("pdfmake");

/* ---------------- HELPERS ---------------- */

function dateSuffix() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, "0")}_${String(
    d.getMonth() + 1
  ).padStart(2, "0")}_${String(d.getFullYear()).slice(-2)}`;
}

function statusMap(v) {
  switch (v) {
    case 0:
      return { t: "Pending", c: "pending" };
    case 1:
      return { t: "Completed", c: "completed" };
    case 2:
      return { t: "In Progress", c: "in-progress" };
    case 16:
      return { t: "In Process", c: "in-process" };
    default:
      return { t: String(v), c: "unknown" };
  }
}

function buildWhere(req, r) {
  let w = "WHERE 1=1";

  if (req.query.search) {
    w += ` AND (TRUCK_REG_NO LIKE @s OR CARD_NO LIKE @s OR CUSTOMER_NAME LIKE @s)`;
    r.input("s", sql.VarChar, `%${req.query.search}%`);
  }

  if (req.query.bay) {
    w += " AND BAY_NO=@b";
    r.input("b", sql.Int, req.query.bay);
  }

  if (req.query.processType) {
    w += " AND PROCESS_TYPE=@p";
    r.input("p", sql.Int, req.query.processType);
  }

  return w;
}

/* ---------------- MAIN PAGE ---------------- */

router.get("/LiveTruckStatus", async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);

    const page = +req.query.page || 1;
    const limit = 20;
    const offset = (page - 1) * limit;

    const cr = pool.request();
    const wc = buildWhere(req, cr);
    const count = await cr.query(`SELECT COUNT(*) t FROM COMMON_VIEW ${wc}`);
    const total = count.recordset[0].t;
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
<input name="search" placeholder="Truck / Card / Customer" value="${escapeHtml(
      req.query.search || ""
    )}">
<select name="bay"><option value="">All Bays</option><option>1</option><option>2</option><option>3</option><option>4</option></select>
<select name="processType"><option value="">All Types</option><option value="1">Loading</option><option value="0">Unloading</option></select>
<button>Apply</button>
<a href="/LiveTruckStatus">Reset</a>
</form>

<div class="export-bar">
  <div class="right-actions">
    <button class="dark-toggle" onclick="toggleDarkMode()" id="darkBtn" title="Toggle Dark Mode">
      🌙
    </button>

    <a href="/LiveTruckStatus/excel?${new URLSearchParams(
      req.query
    )}" class="btn excel">
      Excel
    </a>

    <a href="#" class="btn pdf" onclick="downloadPdf()">PDF</a>
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
<button class="view-btn" onclick='openModal(${escapeHtml(
        JSON.stringify(r.FULL_ROW)
      )})'>
View
</button>
</td>
</tr>
`;
    });

    html += `
</tbody>
</table>

<div class="pagination">
${
  pages > 1
    ? `
      ${page > 1 ? `<a class="page-btn" href="?page=${page - 1}">Prev</a>` : ""}

      ${Array.from(
        { length: pages },
        (_, i) => `
        <a class="page-btn ${page === i + 1 ? "active" : ""}"
           href="?page=${i + 1}">
          ${i + 1}
        </a>
      `
      ).join("")}

      ${
        page < pages
          ? `<a class="page-btn" href="?page=${page + 1}">Next</a>`
          : ""
      }
    `
    : ""
}
</div>


<!-- MODAL -->
<div id="modal" class="modal">
  <div class="modal-content">
    <div class="modal-header">
      <h3>Truck Full Details</h3>
      <span class="modal-close" onclick="closeModal()">✕</span>
    </div>
    <div class="modal-body" id="modalBody"></div>
  </div>
</div>


<script>
function openModal(data){
  const obj = JSON.parse(data);
  let html = '<table class="detail-table">';
  for (const k in obj) {
    html += '<tr><th>'+k+'</th><td>'+obj[k]+'</td></tr>';
  }
  html += '</table>';
  document.getElementById('modalBody').innerHTML = html;
  document.getElementById('modal').style.display='block';
}
function closeModal(){
  document.getElementById('modal').style.display='none';
}
</script>
<script>
function toggleDarkMode() {
  const body = document.body;
  const btn = document.getElementById('darkBtn');

  body.classList.toggle('dark');

  if (body.classList.contains('dark')) {
    btn.textContent = '☀️';
    localStorage.setItem('lts_dark', '1');
  } else {
    btn.textContent = '🌙';
    localStorage.removeItem('lts_dark');
  }
}

// restore state on load
(function () {
  if (localStorage.getItem('lts_dark')) {
    document.body.classList.add('dark');
    const btn = document.getElementById('darkBtn');
    if (btn) btn.textContent = '☀️';
  }
})();
</script>
<script>
function downloadPdf() {
  const params = new URLSearchParams(window.location.search);
  const page = params.get("page") || 1;

  const isCurrent = confirm(
    "Export PDF\n\nOK  → Current Page\nCancel → All Pages"
  );

  if (isCurrent) {
    params.set("scope", "page");
    params.set("page", page);
  } else {
    params.set("scope", "all");
    params.delete("page");
  }

  window.location.href = "/LiveTruckStatus/pdf?" + params.toString();
}
</script>



</body></html>
`;

    res.send(html);
  } catch (e) {
    console.error(e);
    res.status(500).send("Live Truck Status Error");
  }
});

router.get("/LiveTruckStatus/excel", async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);
    const r = pool.request();

    let where = "WHERE 1=1";
    let filterTag = [];

    if (req.query.search) {
      where +=
        " AND (TRUCK_REG_NO LIKE @s OR CARD_NO LIKE @s OR CUSTOMER_NAME LIKE @s)";
      r.input("s", sql.VarChar, `%${req.query.search}%`);
      filterTag.push(`Search-${req.query.search}`);
    }

    if (req.query.bay) {
      where += " AND BAY_NO=@b";
      r.input("b", sql.Int, req.query.bay);
      filterTag.push(`Bay-${req.query.bay}`);
    }

    if (req.query.processType) {
      where += " AND PROCESS_TYPE=@p";
      r.input("p", sql.Int, req.query.processType);
      filterTag.push(`Type-${req.query.processType}`);
    }

    const result = await r.query(`
      SELECT *
      FROM COMMON_VIEW
      ${where}
      ORDER BY FAN_TIME_OUT DESC
    `);

    if (!result.recordset.length) {
      return res.status(404).send("No data available for export");
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Live Truck Status");

    /* ===== Columns ===== */
    ws.columns = Object.keys(result.recordset[0]).map((k) => ({
      header: k.replace(/_/g, " "),
      key: k,
      width: 22,
    }));

    /* ===== Data ===== */
    result.recordset.forEach((row) => ws.addRow(row));

    /* ===== Header Styling ===== */
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF2F5597" }, // professional blue
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    /* ===== Freeze + Filter ===== */
    ws.views = [{ state: "frozen", ySplit: 1 }];
    const lastColumnNumber = ws.columnCount;
    const lastColumnLetter = ws.getColumn(lastColumnNumber).letter;

    ws.autoFilter = {
      from: "A1",
      to: `${lastColumnLetter}1`,
    };

    /* ===== Number formatting ===== */
    ws.eachRow((row, rowNum) => {
      if (rowNum > 1) {
        row.eachCell((cell) => {
          if (typeof cell.value === "number") {
            cell.numFmt = "#,##0";
          }
        });
      }
    });

    /* ===== Filename ===== */
    const datePart = new Date()
      .toLocaleDateString("en-GB")
      .split("/")
      .join("_");

    const filterPart = filterTag.length ? `_${filterTag.join("_")}` : "";

    const fileName = `LiveTruckStatus${filterPart}_${datePart}.xlsx`;

    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Excel error:", err);
    res.status(500).send("Error generating Excel file");
  }
});

router.get("/LiveTruckStatus/pdf", async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);
    const r = pool.request();

    let where = "WHERE 1=1";
    let filterTag = [];

    if (req.query.search) {
      where += " AND (TRUCK_REG_NO LIKE @s OR CARD_NO LIKE @s OR CUSTOMER_NAME LIKE @s)";
      r.input("s", sql.VarChar, `%${req.query.search}%`);
      filterTag.push(`Search-${req.query.search}`);
    }

    if (req.query.bay) {
      where += " AND BAY_NO=@b";
      r.input("b", sql.Int, req.query.bay);
      filterTag.push(`Bay-${req.query.bay}`);
    }

    if (req.query.processType) {
      where += " AND PROCESS_TYPE=@p";
      r.input("p", sql.Int, req.query.processType);
      filterTag.push(`Type-${req.query.processType}`);
    }

    /* ---------- PAGE VS ALL ---------- */
    const limit = 20;
    const page = parseInt(req.query.page || 1);
    const offset = (page - 1) * limit;

    let pagingSql = "";
    if (req.query.scope === "page") {
      pagingSql = ` OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY `;
      filterTag.push(`Page-${page}`);
    }

    const result = await r.query(`
      SELECT *
      FROM COMMON_VIEW
      ${where}
      ORDER BY FAN_TIME_OUT DESC
      ${pagingSql}
    `);

    if (!result.recordset.length) {
      return res.status(404).send("No data available");
    }

    /* ---------- PDF SETUP ---------- */

    const columns = Object.keys(result.recordset[0]);

    const widths = columns.map(col => {
      if (col.includes("TIME") || col.includes("DATE")) return 80;
      if (col.includes("NAME") || col.includes("ADDRESS")) return 100;
      if (col.includes("WEIGHT") || col.includes("CAPACITY")) return 70;
      return 55;
    });

    const body = [
      columns.map(c => ({
        text: c.replace(/_/g, " "),
        style: "tableHeader"
      }))
    ];

    result.recordset.forEach(row => {
      let fillColor = null;

      if (row.BLACKLIST_STATUS === 1) fillColor = "#FEE2E2";
      else if (row.PROCESS_STATUS === 16) fillColor = "#E6FFFA";
      else if (row.PROCESS_STATUS === 15) fillColor = "#FFF7ED";

      body.push(
        columns.map(col => ({
          text: String(row[col] ?? ""),
          fillColor
        }))
      );
    });

    const PdfPrinter = require("pdfmake");
    const printer = new PdfPrinter({
      Helvetica: {
        normal: "Helvetica",
        bold: "Helvetica-Bold"
      }
    });

    const docDefinition = {
      pageSize: "A3",
      pageOrientation: "landscape",
      pageMargins: [20, 50, 20, 40],

      defaultStyle: {
        font: "Helvetica",
        fontSize: 8
      },

      footer: (currentPage, pageCount) => ({
        text: `Page ${currentPage} of ${pageCount}`,
        alignment: "right",
        margin: [0, 0, 20, 0],
        fontSize: 8
      }),

      content: [
        {
          text: "Live Truck Status Report",
          style: "title"
        },
        {
          text: `Generated on: ${new Date().toLocaleString()}`,
          margin: [0, 0, 0, 10]
        },
        {
          table: {
            headerRows: 1,
            widths,
            body
          },
          layout: "lightHorizontalLines"
        }
      ],

      styles: {
        title: {
          fontSize: 16,
          bold: true,
          margin: [0, 0, 0, 10]
        },
        tableHeader: {
          bold: true,
          fillColor: "#1F4E78",
          color: "white"
        }
      }
    };

    const datePart = new Date().toLocaleDateString("en-GB").split("/").join("_");
    const filterPart = filterTag.length ? `_${filterTag.join("_")}` : "";
    const fileName = `LiveTruckStatus${filterPart}_${datePart}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);

    const pdfDoc = printer.createPdfKitDocument(docDefinition);
    pdfDoc.pipe(res);
    pdfDoc.end();

  } catch (err) {
    console.error("PDF error:", err);
    res.status(500).send("PDF generation failed");
  }
});
module.exports = router;
