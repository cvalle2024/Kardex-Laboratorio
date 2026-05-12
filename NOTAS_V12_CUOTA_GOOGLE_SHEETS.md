# V12 — Corrección para cuota 429 de Google Sheets

Esta versión reduce el error:

`APIError: [429]: Quota exceeded for quota metric 'Read requests'`

## Cambios técnicos

- Ya no usa `open_by_key()` ni `worksheet()` de gspread para abrir la hoja en cada arranque.
- Usa la API REST directa de Google Sheets con `values:batchGet`.
- Lee todas las pestañas en una sola solicitud por lote.
- Mantiene caché de datos en Streamlit.
- Incluye reintentos con espera cuando Google responde 429.
- Crea/valida encabezados con operaciones de escritura, evitando lecturas de metadata.

## Secret opcional

Puede ajustar los reintentos agregando en Streamlit Secrets:

```toml
GSHEETS_MAX_RETRIES = 3
GSHEETS_RETRY_SECONDS = 20
AUTO_CREATE_SHEETS = true
```

Para producción mantenga:

```toml
USE_GOOGLE_SHEETS = true
ALLOW_LOCAL_FALLBACK = false
```

Si ya apareció el error 429, espere 1 a 2 minutos antes de reiniciar o refrescar muchas veces la app.
