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
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)`).run();

app.use(require('cors')());
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true }));

// Serve static files from repository root (index.html, sales-billing.html)
app.use(express.static(path.join(__dirname)));

// API: create record
app.post('/api/records', (req, res) => {
  try {
    const r = req.body;
    const stmt = db.prepare(`INSERT INTO invoices (date, invoiceNo, customer, item, quantity, unitPrice, taxRate, taxAmount, discount, grandTotal, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
    const info = stmt.run(r.date, r.invoiceNo, r.customer, r.item, r.quantity, r.unitPrice, r.taxRate, r.taxAmount, r.discount, r.grandTotal, r.notes);
    const inserted = db.prepare('SELECT * FROM invoices WHERE id = ?').get(info.lastInsertRowid);
    res.status(201).json(inserted);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to save record' });
  }
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
  res.json({ todayCount, todayTotal, totalRecords });
});

// API: export to Excel
app.get('/api/export', async (req, res) => {
  const date = req.query.date;
  if (!date) return res.status(400).json({ error: 'Missing date parameter' });

  const rows = db.prepare('SELECT * FROM invoices WHERE date = ? ORDER BY id DESC').all(date);
  if (!rows || rows.length === 0) return res.status(404).json({ error: 'No records for this date' });

  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sales Report');

    sheet.columns = [
      { header: 'Date', key: 'date', width: 12 },
      { header: 'Invoice No', key: 'invoiceNo', width: 15 },
      { header: 'Customer Name', key: 'customer', width: 25 },
      { header: 'Product/Service', key: 'item', width: 25 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit Price', key: 'unitPrice', width: 12 },
      { header: 'Subtotal', key: 'subtotal', width: 12 },
      { header: 'Tax Rate (%)', key: 'taxRate', width: 12 },
      { header: 'Tax Amount', key: 'taxAmount', width: 12 },
      { header: 'Discount', key: 'discount', width: 12 },
      { header: 'Grand Total', key: 'grandTotal', width: 15 },
      { header: 'Notes', key: 'notes', width: 30 }
    ];

    rows.forEach(r => {
      sheet.addRow({
        date: r.date,
        invoiceNo: r.invoiceNo,
        customer: r.customer,
        item: r.item,
        quantity: r.quantity,
        unitPrice: r.unitPrice,
        subtotal: (r.quantity * r.unitPrice),
        taxRate: r.taxRate,
        taxAmount: r.taxAmount,
        discount: r.discount,
        grandTotal: r.grandTotal,
        notes: r.notes
      });
    });

    // Format header
    sheet.getRow(1).font = { bold: true };

    const buf = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Disposition', `attachment; filename="Sales_Report_${date}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to generate Excel' });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server listening on http://0.0.0.0:${PORT}`);
});
