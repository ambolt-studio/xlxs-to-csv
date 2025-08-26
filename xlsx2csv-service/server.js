// server.js
const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');

const app = express();

// Max payload en MB (default: 20). Podés ajustar con MAX_PAYLOAD_MB
const MAX_MB = parseInt(process.env.MAX_PAYLOAD_MB || '20', 10);
app.use(express.json({ limit: `${MAX_MB}mb` }));

// ───────────────────────── utils ─────────────────────────
function decodeBase64Data(str) {
  if (!str || typeof str !== 'string') throw new Error('Missing base64 data');
  const b64 = str.replace(/^data:.*?;base64,/, '');
  return Buffer.from(b64, 'base64');
}

function coerceSheetParam(sheet) {
  // "1" -> 1 (1-based). Si es string no-numérico, se devuelve tal cual.
  if (typeof sheet === 'string' && /^\d+$/.test(sheet.trim())) {
    return parseInt(sheet.trim(), 10);
  }
  return sheet;
}

function pickSheet(wb, sheetParam) {
  const sheet = coerceSheetParam(sheetParam);
  let ws;
  if (typeof sheet === 'number') {
    const idx = Math.max(1, Math.floor(sheet)) - 1; // 1-based → 0-based
    const name = wb.SheetNames[idx];
    if (!name) {
      throw new Error(`Sheet index ${sheet} not found. Available: ${wb.SheetNames.join(', ')}`);
    }
    ws = wb.Sheets[name];
  } else if (typeof sheet === 'string') {
    ws = wb.Sheets[sheet];
    if (!ws) {
      throw new Error(`Sheet "${sheet}" not found. Available: ${wb.SheetNames.join(', ')}`);
    }
  } else {
    ws = wb.Sheets[wb.SheetNames[0]]; // primera hoja
  }
  return ws;
}

function csvFromSheet_XLSX(ws, FS) {
  // Intento 1: conversión directa
  return XLSX.utils.sheet_to_csv(ws, {
    FS,          // separador de campos
    RS: '\n',    // separador de filas
    strip: true,
    blankrows: false,
    raw: false,  // usa valores “mostrados” (cuidado con fechas si querés raw)
    defval: '',  // conserva vacíos
  });
}

function csvFromSheet_AOA(ws, FS) {
  // Intento 2: AOA → CSV manual (rescata contenido cuando !ref está raro)
  const aoa = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: false,
    defval: '',
    blankrows: false,
  });
  if (!Array.isArray(aoa) || !aoa.length) return '';

  const esc = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return aoa.map(row => row.map(esc).join(FS)).join('\n');
}

function whichSoffice() {
  return new Promise((resolve) => {
    execFile('soffice', ['--version'], (err) => resolve(!err));
  });
}

async function csvFromSheet_LibreOffice(ws, FS) {
  // Intento 3 (opcional): usar LibreOffice para evaluar fórmulas
  // Estrategia: creamos un workbook de 1 hoja con 'ws' y lo convertimos.
  const tmpDir = '/tmp';
  const stamp = Date.now() + '_' + Math.random().toString(36).slice(2);
  const inX = path.join(tmpDir, `in_${stamp}.xlsx`);
  const outCsv = path.join(tmpDir, `in_${stamp}.csv`);

  const wb1 = XLSX.utils.book_new();
  // Usamos nombre seguro 'Sheet1' para evitar caracteres problemáticos
  XLSX.utils.book_append_sheet(wb1, ws, 'Sheet1');
  const buf1 = XLSX.write(wb1, { type: 'buffer', bookType: 'xlsx' });
  fs.writeFileSync(inX, buf1);

  try {
    // Forzamos separador coma con filtro CSV de LO (44 = ',', 34 = '"', 0 = UTF-8)
    // Ver doc: "Text - txt - csv (StarCalc)"
    await new Promise((resolve, reject) => {
      execFile(
        'soffice',
        [
          '--headless',
          '--convert-to',
          'csv:"Text - txt - csv (StarCalc)":44,34,0',
          '--outdir',
          tmpDir,
          inX,
        ],
        (err) => (err ? reject(err) : resolve())
      );
    });
    if (fs.existsSync(outCsv)) {
      return fs.readFileSync(outCsv, 'utf8');
    }
    // Fallback: algunos LO generan nombre por defecto
    const alt = path.join(tmpDir, path.basename(inX, '.xlsx') + '.csv');
    return fs.existsSync(alt) ? fs.readFileSync(alt, 'utf8') : '';
  } finally {
    // Limpieza best-effort
    try { fs.unlinkSync(inX); } catch {}
    try { fs.unlinkSync(outCsv); } catch {}
  }
}

// ───────────────────────── endpoints ─────────────────────────

app.get('/healthz', (_req, res) => res.status(200).send('ok'));

// Debug: lista hojas y muestra sample de primeras filas
app.post('/sheets', (req, res) => {
  try {
    const { data } = req.body || {};
    if (!data) return res.status(400).json({ error: 'Missing "data" (base64 XLSX)' });
    const buf = decodeBase64Data(data);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

    const sheets = wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name];
      const ref = ws['!ref'] || null;
      const sample = XLSX.utils.sheet_to_json(ws, {
        header: 1, raw: false, defval: '', blankrows: false,
      }).slice(0, 5);
      return { name, ref, sample };
    });
    res.json({ sheets });
  } catch (e) {
    res.status(500).json({ error: 'sheets_failed', detail: String(e?.message || e) });
  }
});

// Conversión XLSX/XLS → CSV
app.post('/convert', async (req, res) => {
  try {
    let { data, sheet, delimiter, response } = req.body || {};
    if (!data || typeof data !== 'string') {
      return res.status(400).json({ error: 'Missing "data" (base64 XLSX) in request body' });
    }
    // Normalizamos sheet (acepta "2" → 2)
    sheet = coerceSheetParam(sheet);

    const buf = decodeBase64Data(data);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

    const ws = pickSheet(wb, sheet);
    const FS = (typeof delimiter === 'string' && delimiter.length) ? delimiter : ',';

    // 1) XLSX → CSV nativo
    let csv = csvFromSheet_XLSX(ws, FS);

    // 2) Fallback AOA si quedó vacío/blanco
    if (!csv || !csv.trim()) {
      csv = csvFromSheet_AOA(ws, FS);
    }

    // 3) Fallback LibreOffice (opcional, si está habilitado y disponible)
    if ((!csv || !csv.trim()) && String(process.env.USE_LIBREOFFICE || '') === '1') {
      const hasLO = await whichSoffice();
      if (hasLO) {
        try {
          csv = await csvFromSheet_LibreOffice(ws, FS);
        } catch (e) {
          // seguimos sin romper si falla LO
          console.error('LibreOffice fallback failed:', e?.message || e);
        }
      }
    }

    // Respuesta
    if ((response || 'base64') === 'text') {
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      return res.status(200).send(csv || '');
    } else {
      const csvB64 = Buffer.from(csv || '', 'utf8').toString('base64');
      return res.status(200).json({ mime_type: 'text/csv', data: csvB64 });
    }
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'conversion_failed', detail: String(err?.message || err) });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`xlsx2csv service listening on :${PORT}`);
});
