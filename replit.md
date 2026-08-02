# Better Future Tech ERP

A lightweight ERP/sales-billing web app with a static HTML frontend and an Express + SQLite backend.

## Stack
- **Frontend**: Static HTML/CSS/JS (`index.html`, `sales-billing.html`)
- **Backend**: Node.js + Express (`server.js`)
- **Database**: SQLite via `better-sqlite3` (stored at `data/invoices.db`)
- **Excel export**: `exceljs`

## How to run
The app is configured to start automatically via the "Start application" workflow.

Manual start:
```
npm start
```

The server listens on port **5000** (or `$PORT` if set) and serves both the static files and the REST API.

## API endpoints
| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/records` | Create an invoice record |
| GET | `/api/records?date=YYYY-MM-DD` | List records for a date |
| DELETE | `/api/records/:id` | Delete a single record |
| DELETE | `/api/records?confirm=true` | Delete all records |
| GET | `/api/stats?date=YYYY-MM-DD` | Stats for a date |
| GET | `/api/export?date=YYYY-MM-DD` | Download Excel report |

## User preferences
