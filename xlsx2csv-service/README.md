# xlsx2csv-service

Microservicio mínimo (Node + Express + `xlsx`) para convertir **XLSX → CSV** vía **REST**. Ideal para usar desde n8n y luego enviar el CSV a Gemini como `text/csv`.

## Endpoints

### `GET /healthz`
Devuelve `ok` para health checks.

### `POST /convert`
Body (JSON):
```json
{
  "data": "<BASE64 XLSX>",
  "sheet": "Hoja1 or 1",
  "delimiter": ",",
  "response": "base64 | text"
}
```
- **data**: requerido. XLSX en base64 (se admite data URI con prefijo).
- **sheet**: opcional. Nombre de hoja o índice 1-based (1 = primera).
- **delimiter**: opcional, por defecto `,`.
- **response**: `base64` (default) devuelve `{ mime_type, data }`; `text` devuelve `text/csv`.

## Ejemplos

### cURL (respuesta base64 para inline en Gemini)
```bash
curl -X POST "$URL/convert" \
  -H "Content-Type: application/json" \
  -d '{ "data":"<BASE64_XLSX>", "sheet":1, "response":"base64" }'
```

Respuesta:
```json
{ "mime_type": "text/csv", "data": "<BASE64_CSV>" }
```

### cURL (respuesta como text/csv)
```bash
curl -X POST "$URL/convert" \
  -H "Content-Type: application/json" \
  -d '{ "data":"<BASE64_XLSX>", "response":"text" }' -o out.csv
```

### n8n (HTTP Request → POST JSON)
URL: `https://TU-APP.up.railway.app/convert`
```json
{
  "data": "={{$json.data}}",
  "sheet": "={{$json.sheet_name || 1}}",
  "delimiter": ",",
  "response": "base64"
}
```
Luego, en tu llamada a Gemini:
```json
{
  "contents": [{
    "parts": [
      { "text": "Instrucciones…" },
      { "inline_data": { "mime_type": "text/csv", "data": "={{$json.data}}" } }
    ]
  }],
  "generationConfig": { "temperature": 0, "responseMimeType": "application/json" }
}
```

## Deploy en Railway

1. Crear repo con estos archivos (`server.js`, `package.json`, `Dockerfile`).
2. En Railway: **New Project → Deploy from GitHub** → elegí el repo.
3. Variables opcionales:
   - `MAX_PAYLOAD_MB` (por ej. `50` para XLSX grandes).
4. Railway expondrá una URL pública, p. ej. `https://tu-app.up.railway.app`.

> También podés correr localmente:
```bash
npm install
npm start
# POST a http://localhost:8080/convert
```

## Notas sobre datos bancarios
- **Ceros a la izquierda**: si en Excel están como números, puede perderse el cero. Guardar como **texto** en la hoja fuente para preservarlos.
- **Fechas**: `sheet_to_csv` usa el valor mostrado. Si la celda es un serial de Excel, considera normalizar luego a ISO-8601 si lo necesitás.
- **Varias hojas**: usá `sheet` para seleccionar.
