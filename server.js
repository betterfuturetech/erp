const express = require('express');
const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 5000;

// Ensure data directory
const dataDir = path.join(__dirname, 'data');
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

// Init DB
const dbPath = path.join(dataDir, 'invoices.db');
const db = new Database(dbPath);

db.prepare(`CREATE TABLE IF NOT EXISTS invoices (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL,
  invoiceNo TEXT,
  customer TEXT,
  item TEXT,
  quantity REAL,
  unitPrice REAL,
  taxRate REAL,
  taxAmount REAL,
  discount REAL,
  grandTotal REAL,
  notes TEXT,
  paymentType TEXT DEFAULT 'cash',
  receivableStatus TEXT DEFAULT 'paid',
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)`).run();

// Migrate existing DBs: add columns if missing
try { db.prepare(`ALTER TABLE invoices ADD COLUMN paymentType TEXT DEFAULT 'cash'`).run(); } catch(_) {}
try { db.prepare(`ALTER TABLE invoices ADD COLUMN receivableStatus TEXT DEFAULT 'paid'`).run(); } catch(_) {}

app.use(require('cors')());
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true }));

// Serve static files from repository root (index.html, sales-billing.html)
app.use(express.static(path.join(__dirname)));

// API: create record
app.post('/api/records', (req, res) => {
  try {
    const r = req.body;
    const paymentType = (r.paymentType === 'credit') ? 'credit' : 'cash';
    const receivableStatus = (paymentType === 'credit') ? 'open' : 'paid';
    const stmt = db.prepare(`INSERT INTO invoices (date, invoiceNo, customer, item, quantity, unitPrice, taxRate, taxAmount, discount, grandTotal, notes, paymentType, receivableStatus) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
    const info = stmt.run(r.date, r.invoiceNo, r.customer, r.item, r.quantity, r.unitPrice, r.taxRate, r.taxAmount, r.discount, r.grandTotal, r.notes, paymentType, receivableStatus);
    const inserted = db.prepare('SELECT * FROM invoices WHERE id = ?').get(info.lastInsertRowid);
    res.status(201).json(inserted);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to save record' });
  }
});

// API: mark credit invoice as paid
app.patch('/api/records/:id/mark-paid', (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'Invalid id' });
  const info = db.prepare(`UPDATE invoices SET receivableStatus = 'paid' WHERE id = ? AND paymentType = 'credit'`).run(id);
  if (info.changes === 0) return res.status(404).json({ error: 'Record not found or not a credit invoice' });
  res.json({ updated: info.changes });
});

// API: get all open receivables
app.get('/api/receivables', (req, res) => {
  const rows = db.prepare(`SELECT * FROM invoices WHERE paymentType = 'credit' ORDER BY receivableStatus ASC, date DESC`).all();
  res.json(rows);
});

// API: list records by date
app.get('/api/records', (req, res) => {
  const date = req.query.date;
  if (!date) return res.status(400).json({ error: 'Missing date parameter' });
  const rows = db.prepare('SELECT * FROM invoices WHERE date = ? ORDER BY id DESC').all(date);
  res.json(rows);
});

// API: delete single record
app.delete('/api/records/:id', (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'Invalid id' });
  const info = db.prepare('DELETE FROM invoices WHERE id = ?').run(id);
  res.json({ deleted: info.changes });
});

// API: delete all records (useful for reset) - protected by a simple query param to reduce accidental use
app.delete('/api/records', (req, res) => {
  const confirm = req.query.confirm === 'true';
  if (!confirm) return res.status(400).json({ error: 'Missing confirm=true query parameter' });
  const info = db.prepare('DELETE FROM invoices').run();
  res.json({ deleted: info.changes });
});

// API: stats
app.get('/api/stats', (req, res) => {
  const date = req.query.date;
  if (!date) return res.status(400).json({ error: 'Missing date parameter' });
  const todayCount = db.prepare('SELECT COUNT(*) AS c FROM invoices WHERE date = ?').get(date).c;
  const todayTotal = db.prepare('SELECT COALESCE(SUM(grandTotal), 0) AS s FROM invoices WHERE date = ?').get(date).s;
  const totalRecords = db.prepare('SELECT COUNT(*) AS c FROM invoices').get().c;
  const openReceivables = db.prepare(`SELECT COALESCE(SUM(grandTotal), 0) AS s FROM invoices WHERE paymentType = 'credit' AND receivableStatus = 'open'`).get().s;
  const openReceivablesCount = db.prepare(`SELECT COUNT(*) AS c FROM invoices WHERE paymentType = 'credit' AND receivableStatus = 'open'`).get().c;
  res.json({ todayCount, todayTotal, totalRecords, openReceivables, openReceivablesCount });
});

// Helper: style header row
function styleHeader(sheet, color = '2563EB') {
  const row = sheet.getRow(1);
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + color } };
  row.alignment = { vertical: 'middle' };
  row.height = 20;
}

// API: export full bookkeeping workbook
app.get('/api/export', async (req, res) => {
  const date = req.query.date;
  const exportAll = req.query.all === 'true';

  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Better Future Tech ERP';
    workbook.created = new Date();

    // --- Sheet 1: Sales Journal (all invoices for the date or all time) ---
    const salesRows = exportAll
      ? db.prepare('SELECT * FROM invoices ORDER BY date DESC, id DESC').all()
      : (date ? db.prepare('SELECT * FROM invoices WHERE date = ? ORDER BY id DESC').all(date) : []);

    const salesSheet = workbook.addWorksheet('Sales Journal');
    salesSheet.columns = [
      { header: 'Date', key: 'date', width: 12 },
      { header: 'Invoice No', key: 'invoiceNo', width: 15 },
      { header: 'Customer', key: 'customer', width: 25 },
      { header: 'Item / Service', key: 'item', width: 28 },
      { header: 'Qty', key: 'quantity', width: 8 },
      { header: 'Unit Price', key: 'unitPrice', width: 13 },
      { header: 'Subtotal', key: 'subtotal', width: 13 },
      { header: 'Tax %', key: 'taxRate', width: 8 },
      { header: 'Tax Amount', key: 'taxAmount', width: 13 },
      { header: 'Discount', key: 'discount', width: 12 },
      { header: 'Grand Total', key: 'grandTotal', width: 14 },
      { header: 'Payment Type', key: 'paymentType', width: 14 },
      { header: 'Notes', key: 'notes', width: 30 },
    ];
    salesRows.forEach(r => {
      const row = salesSheet.addRow({
        date: r.date, invoiceNo: r.invoiceNo, customer: r.customer, item: r.item,
        quantity: r.quantity, unitPrice: r.unitPrice, subtotal: r.quantity * r.unitPrice,
        taxRate: r.taxRate, taxAmount: r.taxAmount, discount: r.discount,
        grandTotal: r.grandTotal, paymentType: (r.paymentType || 'cash').toUpperCase(),
        notes: r.notes
      });
      ['unitPrice','subtotal','taxAmount','discount','grandTotal'].forEach(k => {
        const cell = row.getCell(k); cell.numFmt = '"$"#,##0.00';
      });
    });
    styleHeader(salesSheet, '2563EB');
    // Total row
    if (salesRows.length) {
      salesSheet.addRow({});
      const tot = salesSheet.addRow({ customer: 'TOTAL', grandTotal: salesRows.reduce((s,r) => s + (r.grandTotal||0), 0) });
      tot.font = { bold: true };
      tot.getCell('grandTotal').numFmt = '"$"#,##0.00';
    }

    // --- Sheet 2: Cash Receipts ---
    const cashRows = exportAll
      ? db.prepare(`SELECT * FROM invoices WHERE paymentType = 'cash' ORDER BY date DESC`).all()
      : (date ? db.prepare(`SELECT * FROM invoices WHERE date = ? AND paymentType = 'cash' ORDER BY id DESC`).all(date) : []);

    const cashSheet = workbook.addWorksheet('Cash Receipts');
    cashSheet.columns = [
      { header: 'Date', key: 'date', width: 12 },
      { header: 'Invoice No', key: 'invoiceNo', width: 15 },
      { header: 'Customer', key: 'customer', width: 25 },
      { header: 'Item / Service', key: 'item', width: 28 },
      { header: 'Grand Total', key: 'grandTotal', width: 14 },
      { header: 'Notes', key: 'notes', width: 30 },
    ];
    cashRows.forEach(r => {
      const row = cashSheet.addRow({
        date: r.date, invoiceNo: r.invoiceNo, customer: r.customer,
        item: r.item, grandTotal: r.grandTotal, notes: r.notes
      });
      row.getCell('grandTotal').numFmt = '"$"#,##0.00';
    });
    styleHeader(cashSheet, '16A34A');
    if (cashRows.length) {
      cashSheet.addRow({});
      const tot = cashSheet.addRow({ customer: 'TOTAL CASH', grandTotal: cashRows.reduce((s,r) => s + (r.grandTotal||0), 0) });
      tot.font = { bold: true };
      tot.getCell('grandTotal').numFmt = '"$"#,##0.00';
    }

    // --- Sheet 3: Accounts Receivable ---
    const arRows = db.prepare(`SELECT * FROM invoices WHERE paymentType = 'credit' ORDER BY receivableStatus ASC, date DESC`).all();

    const arSheet = workbook.addWorksheet('Accounts Receivable');
    arSheet.columns = [
      { header: 'Date', key: 'date', width: 12 },
      { header: 'Invoice No', key: 'invoiceNo', width: 15 },
      { header: 'Customer', key: 'customer', width: 25 },
      { header: 'Item / Service', key: 'item', width: 28 },
      { header: 'Grand Total', key: 'grandTotal', width: 14 },
      { header: 'Status', key: 'status', width: 10 },
      { header: 'Notes', key: 'notes', width: 30 },
    ];
    arRows.forEach(r => {
      const row = arSheet.addRow({
        date: r.date, invoiceNo: r.invoiceNo, customer: r.customer,
        item: r.item, grandTotal: r.grandTotal,
        status: (r.receivableStatus || 'open').toUpperCase(), notes: r.notes
      });
      row.getCell('grandTotal').numFmt = '"$"#,##0.00';
      if ((r.receivableStatus || 'open') === 'open') {
        row.getCell('status').font = { color: { argb: 'FFDC2626' }, bold: true };
      } else {
        row.getCell('status').font = { color: { argb: 'FF16A34A' }, bold: true };
      }
    });
    styleHeader(arSheet, 'B45309');
    if (arRows.length) {
      arSheet.addRow({});
      const openTotal = arRows.filter(r => (r.receivableStatus||'open') === 'open').reduce((s,r) => s + (r.grandTotal||0), 0);
      const collectedTotal = arRows.filter(r => r.receivableStatus === 'paid').reduce((s,r) => s + (r.grandTotal||0), 0);
      const r1 = arSheet.addRow({ customer: 'TOTAL OUTSTANDING', grandTotal: openTotal, status: 'OPEN' });
      r1.font = { bold: true }; r1.getCell('grandTotal').numFmt = '"$"#,##0.00';
      const r2 = arSheet.addRow({ customer: 'TOTAL COLLECTED', grandTotal: collectedTotal, status: 'PAID' });
      r2.font = { bold: true }; r2.getCell('grandTotal').numFmt = '"$"#,##0.00';
    }

    // --- Sheet 4: Summary ---
    const allRows = db.prepare('SELECT * FROM invoices').all();
    const totalSales = allRows.reduce((s,r) => s + (r.grandTotal||0), 0);
    const totalCash = allRows.filter(r => (r.paymentType||'cash') === 'cash').reduce((s,r) => s + (r.grandTotal||0), 0);
    const totalCredit = allRows.filter(r => r.paymentType === 'credit').reduce((s,r) => s + (r.grandTotal||0), 0);
    const totalOutstanding = allRows.filter(r => r.paymentType === 'credit' && (r.receivableStatus||'open') === 'open').reduce((s,r) => s + (r.grandTotal||0), 0);
    const totalCollected = allRows.filter(r => r.paymentType === 'credit' && r.receivableStatus === 'paid').reduce((s,r) => s + (r.grandTotal||0), 0);

    const sumSheet = workbook.addWorksheet('Summary');
    sumSheet.columns = [
      { header: 'Category', key: 'category', width: 35 },
      { header: 'Amount', key: 'amount', width: 18 },
    ];
    styleHeader(sumSheet, '4B5563');
    const summaryData = [
      { category: 'Total Sales (All Time)', amount: totalSales },
      { category: 'Cash Receipts (All Time)', amount: totalCash },
      { category: 'Credit Sales (All Time)', amount: totalCredit },
      { category: '  — Collected / Paid', amount: totalCollected },
      { category: '  — Outstanding (Open AR)', amount: totalOutstanding },
    ];
    summaryData.forEach(d => {
      const row = sumSheet.addRow(d);
      row.getCell('amount').numFmt = '"$"#,##0.00';
      if (d.category.includes('Outstanding')) {
        row.font = { bold: true, color: { argb: 'FFDC2626' } };
      }
    });

    if (!salesRows.length && !exportAll) {
      // still export summary even with no date data
    }

    const filename = exportAll ? 'Bookkeeping_Workbook_All.xlsx' : `Bookkeeping_Workbook_${date}.xlsx`;
    const buf = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to generate workbook' });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server listening on http://0.0.0.0:${PORT}`);
});
