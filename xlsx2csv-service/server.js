// server.js — XLSX/XLS → CSV genérico con normalización por encabezado
// - No pierde filas ni headers.
// - Alinea todas las filas al ancho de la fila de encabezados detectada.
// - Sin separadores de miles (usa valores "raw") y con comillas cuando hace falta.
//
// POST /convert
// Body JSON:
// {
//   "data": "<BASE64 XLSX/XLS>",      // requerido
//   "sheet": 1 | "NombreHoja",        // opcional (1 = primera). Por defecto, primera hoja
//   "delimiter": ",",                 // opcional (",", ";", "\t"). Default ","
//   "force_quotes": true,             // opcional. Default: true
//   "response": "base64" | "text"     // opcional. Default: "base64"
// }
//
// GET  /healthz  → "ok"
// POST /sheets   → { sheets: [{ name, ref, sample: [...] }, ... ] } (debug)
//
// Requiere: npm i express xlsx
//

const express = require('express');
const XLSX = require('xlsx');

const app = express();
const MAX_MB = parseInt(process.env.MAX_PAYLOAD_MB || '20', 10);
app.use(express.json({ limit: `${MAX_MB}mb` }));

// ───────────── helpers ─────────────
function decodeBase64Data(str) {
  if (!str || typeof str !== 'string') throw new Error('Missing base64 data');
  const b64 = str.replace(/^data:.*?;base64,/, '');
  return Buffer.from(b64, 'base64');
}

function pickSheet(wb, sheetParam) {
  // Permite número 1-based o nombre exacto
  if (typeof sheetParam === 'number') {
    const idx = Math.max(1, Math.floor(sheetParam)) - 1;
    const name = wb.SheetNames[idx];
    if (!name) throw new Error(`Sheet index ${sheetParam} not found. Available: ${wb.SheetNames.join(', ')}`);
    return wb.Sheets[name];
  }
  if (typeof sheetParam === 'string' && sheetParam.trim() !== '') {
    const ws = wb.Sheets[sheetParam];
    if (!ws) throw new Error(`Sheet "${sheetParam}" not found. Available: ${wb.SheetNames.join(', ')}`);
    return ws;
  }
  return wb.Sheets[wb.SheetNames[0]];
}

// CSV escape (RFC 4180). Si force=true, cita siempre.
function csvField(value, force = false) {
  let s = value == null ? '' : String(value);
  if (force || /[",\n\r\t;]/.test(s)) {
    s = `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

// Detecta la fila de headers: la más "densa" en las primeras N filas
function findHeaderRow(aoa, scanRows = 50) {
  const limit = Math.min(scanRows, aoa.length);
  let headerRow = 0;
  let bestCount = -1;
  for (let r = 0; r < limit; r++) {
    const row = aoa[r] || [];
    const count = row.reduce((acc, v) => acc + (String(v ?? '').trim() !== '' ? 1 : 0), 0);
    if (count > bestCount) { bestCount = count; headerRow = r; }
  }
  return headerRow;
}

// Calcula el ancho (columna izquierda y derecha) a partir de la fila de headers
function computeTableBounds(headerRowArr) {
  let left = 0;
  while (left < headerRowArr.length && String(headerRowArr[left] ?? '').trim() === '') left++;
  let right = headerRowArr.length - 1;
  while (right >= 0 && String(headerRowArr[right] ?? '').trim() === '') right--;
  return (right < left) ? { left: 0, right: Math.max(0, headerRowArr.length - 1) } : { left, right };
}

// ───────────── endpoints ─────────────

app.get('/healthz', (_req, res) => res.status(200).send('ok'));

// Debug: lista hojas y muestra 5 filas crudas de cada una
app.post('/sheets', (req, res) => {
  try {
    const buf = decodeBase64Data(req.body?.data);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
    const sheets = wb.SheetNames.map(name => {
      const ws = wb.Sheets[name];
      const ref = ws['!ref'] || null;
      const sample = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false }).slice(0, 5);
      return { name, ref, sample };
    });
    res.json({ sheets });
  } catch (e) {
    res.status(400).json({ error: String(e?.message || e) });
  }
});

// Conversión genérica con normalización por encabezado
app.post('/convert', (req, res) => {
  try {
    const body = req.body || {};
    const buf  = decodeBase64Data(body.data);
    const wb   = XLSX.read(buf, { type: 'buffer', cellDates: true });
    const ws   = pickSheet(wb, body.sheet);

    const ref = ws['!ref'];
    if (!ref) {
      if (body.response === 'text') return res.type('text/csv; charset=utf-8').send('');
      return res.json({ mime_type: 'text/csv', data: '' });
    }

    // 1) Volcamos todo como AOA "raw": sin miles, sin formatos
    const aoaFull = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: true,
      defval: '',
      blankrows: false,
    });

    if (!Array.isArray(aoaFull) || !aoaFull.length) {
      if (body.response === 'text') return res.type('text/csv; charset=utf-8').send('');
      return res.json({ mime_type: 'text/csv', data: '' });
    }

    // 2) Detectamos la fila de headers y los límites de tabla (left/right)
    const headerRowIdx = findHeaderRow(aoaFull, 50);
    const headerRowArr = aoaFull[headerRowIdx] || [];
    let { left, right } = computeTableBounds(headerRowArr);

    // Fallback: si el header quedó de ancho 0, ampliamos con la fila más ancha de las siguientes 50
    if (right <= left && aoaFull.length > headerRowIdx + 1) {
      let maxLen = headerRowArr.length;
      for (let r = headerRowIdx + 1; r < Math.min(aoaFull.length, headerRowIdx + 51); r++) {
        maxLen = Math.max(maxLen, (aoaFull[r] || []).length);
      }
      right = Math.max(right, maxLen - 1);
    }

    // 3) Recortamos todas las filas al ancho de la tabla y eliminamos filas completamente vacías
    const sliceRow = (row) => {
      const out = [];
      for (let c = left; c <= right; c++) out.push(row[c] ?? '');
      return out;
    };

    const headers = sliceRow(headerRowArr);
    const table = [headers];
    for (let r = headerRowIdx + 1; r < aoaFull.length; r++) {
      const row = sliceRow(aoaFull[r] || []);
      const allEmpty = row.every(v => String(v ?? '').trim() === '');
      if (!allEmpty) table.push(row);
    }

    // 4) Serializamos CSV RFC-4180
    const FS = (typeof body.delimiter === 'string' && body.delimiter.length) ? body.delimiter : ',';
    const forceQuotes = body.force_quotes !== false; // por defecto true

    const lines = table.map(r => r.map(v => csvField(v, forceQuotes)).join(FS));
    const csv = lines.join('\n');

    // 5) Respuesta
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
app.listen(PORT, () => {
  console.log(`xlsx2csv listening on :${PORT}`);
});

