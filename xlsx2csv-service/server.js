// server.js
const express = require('express');
const XLSX = require('xlsx');

const app = express();

// Max payload in MB (default: 20). Override with env MAX_PAYLOAD_MB
const MAX_MB = parseInt(process.env.MAX_PAYLOAD_MB || '20', 10);
app.use(express.json({ limit: `${MAX_MB}mb` }));

// Health check
app.get('/healthz', (_req, res) => res.status(200).send('ok'));

/**
 * POST /convert
 * Body (JSON):
 * {
 *   "data": "<BASE64 XLSX>",           // required
 *   "sheet": "Hoja1" | 1,              // optional (name or 1-based index). Default: first sheet
 *   "delimiter": ",",                  // optional, default: ","
 *   "response": "base64" | "text"      // optional, default: "base64"
 * }
 *
 * Responses:
 * - response="base64" (default):
 *   { "mime_type":"text/csv","data":"<BASE64 CSV>" }
 * - response="text":
 *   Content-Type: text/csv ; body = CSV
 */
app.post('/convert', (req, res) => {
  try {
    const { data, sheet, delimiter, response } = req.body || {};
    if (!data || typeof data !== 'string') {
      return res.status(400).json({ error: 'Missing "data" (base64 XLSX) in request body' });
    }

    // Strip data URI prefix if present
    const b64 = data.replace(/^data:.*?;base64,/, '');
    let buf;
    try {
      buf = Buffer.from(b64, 'base64');
    } catch (e) {
      return res.status(400).json({ error: 'Invalid base64 in "data"' });
    }

    // Read workbook
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

    // Select sheet
    let ws;
    if (typeof sheet === 'number') {
      const idx = Math.max(1, Math.floor(sheet)) - 1; // 1-based -> 0-based
      const name = wb.SheetNames[idx];
      if (!name) {
        return res.status(400).json({ error: `Sheet index ${sheet} not found. Available: ${wb.SheetNames.join(', ')}` });
      }
      ws = wb.Sheets[name];
    } else if (typeof sheet === 'string') {
      ws = wb.Sheets[sheet];
      if (!ws) {
        return res.status(400).json({ error: `Sheet "${sheet}" not found. Available: ${wb.SheetNames.join(', ')}` });
      }
    } else {
      ws = wb.Sheets[wb.SheetNames[0]];
    }

    // XLSX -> CSV
    const FS = (typeof delimiter === 'string' && delimiter.length) ? delimiter : ',';
    const csv = XLSX.utils.sheet_to_csv(ws, {
      FS,          // field separator
      RS: '\n',    // row separator
      strip: true, // trim cells
      blankrows: false,
    });

    if ((response || 'base64') === 'text') {
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      return res.status(200).send(csv);
    } else {
      const csvB64 = Buffer.from(csv, 'utf8').toString('base64');
      return res.status(200).json({ mime_type: 'text/csv', data: csvB64 });
    }
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'conversion_failed', detail: String(err && err.message || err) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`xlsx2csv service listening on :${PORT}`);
});
