// server.js
const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');

const app = express();
const MAX_MB = parseInt(process.env.MAX_PAYLOAD_MB || '20', 10);
app.use(express.json({ limit: `${MAX_MB}mb` }));

// ───────────── utils ─────────────
function decodeBase64Data(str) {
  if (!str || typeof str !== 'string') throw new Error('Missing base64 data');
  const b64 = str.replace(/^data:.*?;base64,/, '');
  return Buffer.from(b64, 'base64');
}
function coerceSheetParam(sheet) {
  if (typeof sheet === 'string' && /^\d+$/.test(sheet.trim())) return parseInt(sheet.trim(), 10);
  return sheet;
}
function pickSheet(wb, sheetParam) {
  const sheet = coerceSheetParam(sheetParam);
  if (typeof sheet === 'number') {
    const idx = Math.max(1, Math.floor(sheet)) - 1;
    const name = wb.SheetNames[idx];
    if (!name) throw new Error(`Sheet index ${sheet} not found. Available: ${wb.SheetNames.join(', ')}`);
    return wb.Sheets[name];
  }
  if (typeof sheet === 'string') {
    const ws = wb.Sheets[sheet];
    if (!ws) throw new Error(`Sheet "${sheet}" not found. Available: ${wb.SheetNames.join(', ')}`);
    return ws;
  }
  return wb.Sheets[wb.SheetNames[0]];
}
function whichSoffice() {
  return new Promise((resolve) => { execFile('soffice', ['--version'], (err) => resolve(!err)); });
}

// CSV safe-escape
function esc(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// Busca automáticamente la fila de headers: prioridad a una que contenga "date"
function findHeaderRow(aoa) {
  if (!Array.isArray(aoa)) return -1;
  let bestIdx = -1, bestScore = -1;
  for (let i = 0; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const nonEmpty = row.filter(c => (c !== null && c !== undefined && String(c).trim() !== '')).length;
    const hasDate = row.some(c => String(c).trim().toLowerCase() === 'date');
    const score = (hasDate ? 100 : 0) + nonEmpty; // favorece “date” y densidad
    if (score > bestScore) { bestScore = score; bestIdx = i; }
    if (hasDate && nonEmpty >= 3) return i; // match “fuerte”
  }
  return bestIdx;
}

// ───────────── endpoints ─────────────
app.get('/healthz', (_req, res) => res.status(200).send('ok'));

// Debug: lista hojas + preview de primeras filas
app.post('/sheets', (req, res) => {
  try {
    const { data } = req.body || {};
    if (!data) return res.status(400).json({ error: 'Missing "data" (base64 XLSX)' });
    const buf = decodeBase64Data(data);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

    const sheets = wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name];
      const ref = ws['!ref'] || null;
      const sample = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false }).slice(0, 5);
      return { name, ref, sample };
    });
    res.json({ sheets });
  } catch (e) {
    res.status(500).json({ error: 'sheets_failed', detail: String(e?.message || e) });
  }
});

/**
 * POST /convert
 * Body:
 * {
 *   data: "<BASE64 XLSX/XLS>",
 *   sheet: 1 | "brex cta",                 // opcional
 *   delimiter: "," | ";",                  // opcional (default ",")
 *   response: "base64" | "text",           // opcional (default "base64")
 *   header_row: number | "auto",           // opcional (default "auto")
 *   skip_regex: "^(Vendor Name)",          // opcional (default salta Vendor Name)
 *   raw: true|false,                       // opcional (default true)  -> números sin miles
 *   force_quotes: true|false               // opcional (default true)  -> todo entre comillas
 * }
 */
app.post('/convert', async (req, res) => {
  try {
    let {
      data, sheet, delimiter, response,
      header_row, skip_regex, raw, force_quotes,
    } = req.body || {};
    if (!data || typeof data !== 'string') {
      return res.status(400).json({ error: 'Missing "data" (base64 XLSX) in request body' });
    }
    sheet = coerceSheetParam(sheet);

    // defaults robustos para bancos
    const FS = (typeof delimiter === 'string' && delimiter.length) ? delimiter : ',';
    raw = (typeof raw === 'boolean') ? raw : true;
    force_quotes = (typeof force_quotes === 'boolean') ? force_quotes : true;
    const skipRe = new RegExp(skip_regex || '^(Vendor Name)', 'i');

    const buf = decodeBase64Data(data);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
    const ws = pickSheet(wb, sheet);

    // AOA con defval para conservar vacíos
    const aoa = XLSX.utils.sheet_to_json(ws, {
      header: 1, raw, defval: '', blankrows: false,
    });

    if (!Array.isArray(aoa) || !aoa.length) {
      const emptyCsv = '';
      if ((response || 'base64') === 'text') {
        res.setHeader('Content-Type', 'text/csv; charset=utf-8');
        return res.status(200).send(emptyCsv);
      }
      return res.status(200).json({ mime_type: 'text/csv', data: Buffer.from(emptyCsv, 'utf8').toString('base64') });
    }

    // Detectar header row
    let hr;
    if (typeof header_row === 'number') {
      hr = Math.max(1, Math.floor(header_row)) - 1;
    } else {
      hr = findHeaderRow(aoa);
    }
    const headersRaw = (aoa[hr] || []).map(h => String(h || '').trim());
    // Normalizar headers vacíos
    const headers = headersRaw.map((h, i) => h || `col_${i+1}`);

    // Filtrar filas de datos a partir de hr+1
    const dataRows = aoa.slice(hr + 1).filter(row => {
      const first = String((row[0] ?? '')).trim();
      const allEmpty = row.every(c => (String(c || '').trim() === ''));
      if (allEmpty) return false;
      if (skipRe.test(first)) return false; // e.g., "Vendor Name : …"
      return true;
    });

    // Alinear ancho y construir CSV con headers
    let csv = headers.map(esc).join(FS) + '\n' +
      dataRows.map(r => headers.map((_, i) => esc(r[i] ?? '')).join(FS)).join('\n');

    // Fallback LibreOffice si quedó vacío y está habilitado
    if ((!csv || !csv.trim()) && String(process.env.USE_LIBREOFFICE || '') === '1') {
      const hasLO = await whichSoffice();
      if (hasLO) {
        // Generamos un libro con solo las filas útiles
        const ws2 = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
        const csvLO = await new Promise((resolve) => {
          const tmpDir = '/tmp';
          const stamp = Date.now() + '_' + Math.random().toString(36).slice(2);
          const inX = path.join(tmpDir, `in_${stamp}.xlsx`);
          const outCsv = path.join(tmpDir, `in_${stamp}.csv`);
          const wb2 = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb2, ws2, 'Sheet1');
          const buf2 = XLSX.write(wb2, { type: 'buffer', bookType: 'xlsx' });
          fs.writeFileSync(inX, buf2);
          execFile(
            'soffice',
            ['--headless', '--convert-to', 'csv:"Text - txt - csv (StarCalc)":44,34,0', '--outdir', tmpDir, inX],
            () => {
              try {
                const text = fs.existsSync(outCsv) ? fs.readFileSync(outCsv, 'utf8') : '';
                fs.unlinkSync(inX); fs.unlinkSync(outCsv);
                resolve(text);
              } catch { resolve(''); }
            }
          );
        });
        if (csvLO && csvLO.trim()) csv = csvLO;
      }
    }

    // Forzar comillas si se pidió
    if (force_quotes) {
      // Ya escapamos con esc(), que pone comillas si hace falta.
      // Para “forzar”, envolvemos todo de nuevo si no tiene comillas.
      const lines = csv.split('\n').map(line => {
        if (!line) return line;
        const cols = line.split(FS).map(c => {
          if (!/^".*"$/.test(c)) return `"${c.replace(/^"|"$/g, '')}"`;
          return c;
        });
        return cols.join(FS);
      });
      csv = lines.join('\n');
    }

    if ((response || 'base64') === 'text') {
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      return res.status(200).send(csv);
    }
    return res.status(200).json({ mime_type: 'text/csv', data: Buffer.from(csv, 'utf8').toString('base64') });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'conversion_failed', detail: String(err?.message || err) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`xlsx2csv service listening on :${PORT}`);
});
