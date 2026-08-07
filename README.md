# Better Future Tech ERP - Local Server

This repository includes a static frontend (index.html, sales-billing.html) and a small Express server to provide persistence and Excel export for the Sales/Billing app.

How to run locally on your machine and expose to the LAN:

1. Install Node.js (>=14) and npm.
2. From the repository root:

   npm install

3. Start the server:

   npm start

4. Find your machine IP on the local network (e.g., 192.168.1.10) and open in a browser:

   http://<your-ip>:3000/index.html

Notes:
- The server binds to 0.0.0.0 so other machines on your LAN can access it.
- Data is stored in data/invoices.db (SQLite); exported Excel files are generated on-demand.
- To reset all records (dangerous), send a DELETE request to:

  http://<your-ip>:3000/api/records?confirm=true


If you want, I can also:
- Add a simple systemd service or Dockerfile for easier deployment on an always-on machine.
- Add validation and authentication for the API.
