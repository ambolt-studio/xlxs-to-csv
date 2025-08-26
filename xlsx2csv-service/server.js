// server.js — conversión genérica XLSX/XLS → CSV “tabular” sin perder filas ni headers
const express = require('express');
const XLSX = require('xlsx');

const app = express();
const MAX_MB = parseInt(process.env.MAX_PAYLOAD_MB || '20', 10);
app.use(express.json({ limit: `${MAX_MB}mb` }));

// ───────── utils ─────────
function b64buf(s) {
  if (typeof s !== 'string') throw new Error('data must be base64 string');
  const clean = s.replace(/^data:.*?;base64,/, '');
  return Buffer.from(clean, 'base64');
}

// CSV escape (RFC-4180)
function esc(v) {
  if (v === null || v === undefined) v = '';
  const s = String(v);
  return (s.includes('"') || s.includes('\n') || s.includes('\r') || s.includes(',') || s.includes(';') || s.includes('\t'))
    ? `"${s.replace(/"/g, '""')}"`
    : s;
}

// date → ISO (YYYY-MM-DD o YYYY-MM-DD HH:MM:SS)
function dateToISO(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const pad = n => String(n).padStart(2,'0');
  const y = d.getFullYear(), m = pad(d.getMonth()+1), day = pad(d.getDate());
  const hh = pad(d.getHours()), mm = pad(d.getMinutes()), ss = pad(d.getSeconds());
  // si no hay parte horaria, solo fecha
  return (d.getHours()===0 && d.getMinutes()===0 && d.getSeconds()===0)
    ? `${y}-${m}-${day}`
    : `${y}-${m}-${day} ${hh}:${mm}:${ss}`;
}

// renderiza celda SIN separadores de miles, respetando tipos
function cellToString(cell) {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  switch (cell.t) {
    case 'd': return dateToISO(cell.v);            // ya es Date si usamos cellDates:true
    case 'n': return Number.isFinite(cell.v) ? String(cell.v) : String(cell.v ?? ''); // 50000.12, sin “,” de miles
    case 'b': return cell.v ? 'TRUE' : 'FALSE';
    default:  return String(cell.v);
  }
}

// expande merges opcionalmente copiando el valor del topleft
function buildMergeFillMap(ws) {
  const map = new Map();
  const merges = ws['!merges'] || [];
  for (const m of merges) {
    const { s, e } = m; // start/end {r,c}
    const topLeft = XLSX.utils.encode_cell(s);
    for (let r = s.r; r <= e.r; r++) {
      for (let c = s.c; c <= e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (addr !== topLeft) map.set(addr, topLeft);
      }
    }
  }
  return map;
}

// ───────── endpoints ─────────
app.get('/healthz', (_req, res) => res.status(200).send('ok'));

// lista hojas + preview (debug)
app.post('/sheets', (req, res) => {
  try {
    const buf = b64buf(req.body?.data);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
    const sheets = wb.SheetNames.map(n => {
      const ws = wb.Sheets[n];
      const ref = ws['!ref'] || null;
      const sample = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }).slice(0, 5);
      return { name: n, ref, sample };
    });
    res.json({ sheets });
  } catch (e) {
    res.status(400).json({ error: String(e?.message || e) });
  }
});

/**
 * POST /convert
 * Body:
 * {
 *   "data": "<BASE64 XLSX/XLS>",   // requerido
 *   "sheet": 1 | "Hoja",           // opcional (1-based o nombre). default: 1ra hoja
 *   "delimiter": ",",              // opcional: "," ";" "\t"  (default: ",")
 *   "force_quotes": false,         // opcional: si true, comilla siempre
 *   "fill_merges": false,          // opcional: si true, replica valor del topleft sobre celdas merged
 *   "response": "base64"|"text"    // opcional (default: "base64")
 * }
 */
app.post('/convert', (req, res) => {
  try {
    const body = req.body || {};
    const buf  = b64buf(body.data);
    const wb   = XLSX.read(buf, { type: 'buffer', cellDates: true });

    // hoja
    let ws;
    if (typeof body.sheet === 'number') {
      const idx = Math.max(1, Math.floor(body.sheet)) - 1;
      const name = wb.SheetNames[idx];
      if (!name) return res.status(400).json({ error: `Sheet index ${body.sheet} not found. Available: ${wb.SheetNames.join(', ')}` });
      ws = wb.Sheets[name];
    } else if (typeof body.sheet === 'string' && body.sheet.trim() !== '') {
      ws = wb.Sheets[body.sheet];
      if (!ws) return res.status(400).json({ error: `Sheet "${body.sheet}" not found. Available: ${wb.SheetNames.join(', ')}` });
    } else {
      ws = wb.Sheets[wb.SheetNames[0]];
    }

    // rango usado exacto
    const ref = ws['!ref'];
    if (!ref) {
      const empty = '';
      return (body.response === 'text')
        ? res.type('text/csv').send(empty)
        : res.json({ mime_type: 'text/csv', data: Buffer.from(empty, 'utf8').toString('base64') });
    }
    const range = XLSX.utils.decode_range(ref);

    // opcional: relleno de merges
    const mergeMap = body.fill_merges ? buildMergeFillMap(ws) : new Map();

    // recorremos celda por celda para asegurar ancho constante y sin pérdidas
    const rows = [];
    for (let R = range.s.r; R <= range.e.r; R++) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const src  = mergeMap.get(addr) || addr;
        const cell = ws[src];
        row.push(cellToString(cell));
      }
      rows.push(row);
    }

    // delimitador
    const FS = (typeof body.delimiter === 'string' && body.delimiter.length) ? body.delimiter : ',';

    // a CSV
    const lines = rows.map(r => r.map(v => body.force_quotes ? `"${String(v).replace(/"/g,'""')}"` : esc(v)).join(FS));
    const csv = lines.join('\n');

    if (body.response === 'text') {
      res.type('text/csv; charset=utf-8').send(csv);
    } else {
      res.json({ mime_type: 'text/csv', data: Buffer.from(csv, 'utf8').toString('base64') });
    }
  } catch (e) {
    res.status(400).json({ error: String(e?.message || e) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log('xlsx2csv listening on :' + PORT));

