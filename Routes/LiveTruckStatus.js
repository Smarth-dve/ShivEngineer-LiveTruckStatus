const express = require('express');
const router = express.Router();
const sql = require('mssql/msnodesqlv8');
const escapeHtml = require('escape-html');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const dbConfig = require('../Config/dbConfig');

/*
  LIVE TRUCK STATUS
  SAME STRUCTURE & BEHAVIOR AS tees.js
  SOURCE: COMMON_VIEW_3
*/

router.get('/LiveTruckStatus', async (req, res) => {
  try {
    const pool = await sql.connect(dbConfig);

    /* pagination */
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 20;
    const offset = (page - 1) * limit;

    /* filters */
    const search = req.query.search ? `%${req.query.search}%` : null;
    const bay = req.query.bay || null;
    const type = req.query.type || null;

    /* sorting */
    const sort = req.query.sort || 'FAN_TIMEOUT';
    const order = req.query.order === 'ASC' ? 'ASC' : 'DESC';

    const allowedSortMap = {
      TRUCK_REG_NO: '[TRUCK REG NO]',
      CARD_NO: '[CARD NO]',
      STATUS: '[STATUS]',
      TYPE: '[TYPE]',
      BAY: '[BAY]',
      BATCH_STATUS: '[BATCH_STATUS]',
      ENTRY_GATE_TIME: '[ENTRY GATE TIME]',
      FAN_TIMEOUT: '[FAN TIMEOUT]'
    };

    const sortColumn = allowedSortMap[sort] || '[FAN TIMEOUT]';

    /* where clause */
    let whereClause = 'WHERE 1=1';

    if (search) {
      whereClause += `
        AND (
          [TRUCK REG NO] LIKE @search OR
          [CARD NO] LIKE @search OR
          [STATUS] LIKE @search
        )
      `;
    }

    if (bay) whereClause += ' AND [BAY] = @bay';
    if (type) whereClause += ' AND [TYPE] = @type';

    /* count query */
    const countReq = pool.request();
    if (search) countReq.input('search', sql.VarChar, search);
    if (bay) countReq.input('bay', sql.Int, bay);
    if (type) countReq.input('type', sql.VarChar, type);

    const countResult = await countReq.query(`
      SELECT COUNT(*) AS total
      FROM COMMON_VIEW_3
      ${whereClause}
    `);

    const totalRecords = countResult.recordset[0].total;
    const totalPages = Math.ceil(totalRecords / limit);

    /* data query */
    const dataReq = pool.request();
    if (search) dataReq.input('search', sql.VarChar, search);
    if (bay) dataReq.input('bay', sql.Int, bay);
    if (type) dataReq.input('type', sql.VarChar, type);

    const dataResult = await dataReq.query(`
      SELECT
        [TRUCK REG NO],
        [CARD NO],
        [STATUS],
        [TYPE],
        [BAY],
        [BATCH_STATUS],
        [ENTRY GATE TIME],
        [FAN TIMEOUT]
      FROM COMMON_VIEW_3
      ${whereClause}
      ORDER BY ${sortColumn} ${order}
      OFFSET ${offset} ROWS
      FETCH NEXT ${limit} ROWS ONLY
    `);

    const rows = dataResult.recordset;

    /* HTML */
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

${fs.readFileSync(path.join(__dirname, '../public/Css/navbar.html'), 'utf8')}

<h2>LIVE TRUCK STATUS</h2>

<div class="top-bar">
  <form method="GET" action="/LiveTruckStatus">
    <input type="text" name="search" placeholder="Search Truck / Card / Status"
      value="${escapeHtml(req.query.search || '')}">

    <select name="bay">
      <option value="">All Bays</option>
      <option value="1">Bay 1</option>
      <option value="2">Bay 2</option>
      <option value="3">Bay 3</option>
      <option value="4">Bay 4</option>
    </select>

    <select name="type">
      <option value="">All Types</option>
      <option value="Loading">Loading</option>
      <option value="UnLoading">UnLoading</option>
    </select>

    <button type="submit">Search</button>
    <a class="btn-reset" href="/LiveTruckStatus">Refresh</a>
  </form>

  <div class="export-bar">
    <a href="/LiveTruckStatus/excel" class="export-btn excel">Excel</a>
    <a href="/LiveTruckStatus/pdf" class="export-btn pdf">PDF</a>
  </div>
</div>

<table>
<thead>
<tr>
  <th>Sr</th>
  <th>Truck Reg No</th>
  <th>Card No</th>
  <th>Status</th>
  <th>Type</th>
  <th>Bay</th>
  <th>Batch</th>
  <th>Entry Gate</th>
  <th>Fan Timeout</th>
</tr>
</thead>
<tbody>
`;

    rows.forEach((r, i) => {
      const statusClass = r['STATUS']
        ? r['STATUS'].replace(/\s+/g, '-').toLowerCase()
        : '';

      html += `
<tr>
  <td>${offset + i + 1}</td>
  <td>${escapeHtml(r['TRUCK REG NO'] || '')}</td>
  <td>${escapeHtml(r['CARD NO'] || '')}</td>
  <td class="status-col ${statusClass}">${escapeHtml(r['STATUS'] || '')}</td>
  <td>${escapeHtml(r['TYPE'] || '')}</td>
  <td>${r['BAY'] ?? ''}</td>
  <td>${r['BATCH_STATUS'] ?? ''}</td>
  <td>${r['ENTRY GATE TIME'] ?? ''}</td>
  <td>${r['FAN TIMEOUT'] ?? ''}</td>
</tr>`;
    });

    html += `
</tbody>
</table>

<div class="pagination">
Page ${page} of ${totalPages}<br/>
`;

    for (let p = 1; p <= totalPages; p++) {
      html += `<a href="/LiveTruckStatus?page=${p}&limit=${limit}">${p}</a>`;
    }

    html += `
</div>

<label class="auto-refresh">
  <input type="checkbox" id="autoRefresh"> Auto Refresh (10s)
</label>

<script>
let interval;
document.getElementById('autoRefresh').addEventListener('change', e => {
  if (e.target.checked) interval = setInterval(() => location.reload(), 10000);
  else clearInterval(interval);
});
</script>

</body>
</html>
`;

    res.send(html);

  } catch (err) {
    console.error(err);
    res.status(500).send('Live Truck Status error');
  }
});

/* ===== EXCEL EXPORT ===== */
router.get('/LiveTruckStatus/excel', async (req, res) => {
  const pool = await sql.connect(dbConfig);
  const result = await pool.request().query('SELECT * FROM COMMON_VIEW_3');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Live Truck Status');

  ws.columns = Object.keys(result.recordset[0]).map(k => ({
    header: k,
    key: k
  }));

  result.recordset.forEach(r => ws.addRow(r));

  res.setHeader('Content-Disposition', 'attachment; filename=LiveTruckStatus.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

/* ===== PDF EXPORT ===== */
router.get('/LiveTruckStatus/pdf', async (req, res) => {
  const pool = await sql.connect(dbConfig);
  const result = await pool.request().query('SELECT * FROM COMMON_VIEW_3');

  const doc = new PDFDocument({ size: 'A4', margin: 30 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', 'attachment; filename=LiveTruckStatus.pdf');

  doc.pipe(res);
  doc.fontSize(16).text('LIVE TRUCK STATUS REPORT', { align: 'center' });
  doc.moveDown();

  result.recordset.forEach((r, i) => {
    doc.fontSize(9).text(
      `${i + 1}. Truck:${r['TRUCK REG NO']} | Card:${r['CARD NO']} | Status:${r['STATUS']}`
    );
  });

  doc.end();
});

module.exports = router;
