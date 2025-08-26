app.post('/convert', (req, res) => {
  try {
    const body = req.body || {};
    const buf  = Buffer.from(String(body.data || '').replace(/^data:.*?;base64,/, ''), 'base64');
    if (!buf.length) return res.status(400).json({ error: 'Missing "data" (base64 XLSX)' });

    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

    // Selección de hoja (número 1-based o nombre)
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

    const ref = ws['!ref'];
    if (!ref) {
      const empty = '';
      return (body.response === 'text')
        ? res.type('text/csv; charset=utf-8').send(empty)
        : res.json({ mime_type: 'text/csv', data: Buffer.from(empty, 'utf8').toString('base64') });
    }

    // ---------- 1) Leemos TODAS las celdas como AOA (crudo, sin miles) ----------
    const aoaFull = XLSX.utils.sheet_to_json(ws, {
      header: 1, raw: true, defval: '', blankrows: false,
    });
    if (!aoaFull.length) {
      const empty = '';
      return (body.response === 'text')
        ? res.type('text/csv; charset=utf-8').send(empty)
        : res.json({ mime_type: 'text/csv', data: Buffer.from(empty, 'utf8').toString('base64') });
    }

    // ---------- 2) Detectamos fila de header y límites de la tabla ----------
    // Estrategia: elegir la fila con mayor cantidad de celdas no vacías
    // dentro de las primeras N filas (heights habituales). Luego tomamos
    // el primer y último índice no vacío de ESA fila como ancho de la tabla.
    const SCAN_ROWS = Math.min(50, aoaFull.length);
    let headerRow = 0, bestCount = -1;
    for (let r = 0; r < SCAN_ROWS; r++) {
      const row = aoaFull[r] || [];
      const count = row.reduce((acc, v) => acc + (String(v).trim() !== '' ? 1 : 0), 0);
      if (count > bestCount) { bestCount = count; headerRow = r; }
    }
    const hdr = aoaFull[headerRow] || [];
    let left = 0;
    while (left < hdr.length && String(hdr[left]).trim() === '') left++;
    let right = hdr.length - 1;
    while (right >= 0 && String(hdr[right]).trim() === '') right--;

    if (right < left) {
      // No pudimos definir ancho: devolvemos todo sin normalizar
      left = 0; right = (aoaFull[0] || []).length - 1;
    }

    // ---------- 3) Cortamos por ese sub-rango y dejamos headers + datos ----------
    const sliceRow = (row) => {
      const out = [];
      for (let c = left; c <= right; c++) out.push(row[c] ?? '');
      return out;
    };
    const headers = sliceRow(hdr).map(x => {
      const s = String(x || '').trim();
      return s || ''; // si querés, podés poner col_1, col_2…
    });

    const table = [headers];
    for (let r = headerRow + 1; r < aoaFull.length; r++) {
      table.push(sliceRow(aoaFull[r] || []));
    }

    // ---------- 4) Serializamos CSV RFC-4180 ----------
    const FS = (typeof body.delimiter === 'string' && body.delimiter.length) ? body.delimiter : ',';
    const forceQuotes = body.force_quotes === true;

    const esc = (v) => {
      if (v === null || v === undefined) v = '';
      // Normalizamos números: sin separador de miles (ya viene crudo de XLSX)
      // Fechas: si XLSX las interpretó como Date, sheet_to_json(raw:true) trae serial; pero
      // esta AOA conserva string renderizado. Si querés ISO estricto, convertí antes.
      const s = String(v);
      const needs = forceQuotes || /[",\n\r\t;]/.test(s);
      return needs ? `"${s.replace(/"/g, '""')}"` : s;
    };

    const csv = table.map(row => row.map(esc).join(FS)).join('\n');

    // ---------- 5) Respuesta ----------
    if (body.response === 'text') {
      return res.type('text/csv; charset=utf-8').send(csv);
    }
    return res.json({ mime_type: 'text/csv', data: Buffer.from(csv, 'utf8').toString('base64') });

  } catch (e) {
    return res.status(400).json({ error: String(e?.message || e) });
  }
});
