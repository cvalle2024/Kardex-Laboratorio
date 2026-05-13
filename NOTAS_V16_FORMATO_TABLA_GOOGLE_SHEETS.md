# V16 - Formato tipo tabla en Google Sheets

Esta versión agrega formato visual automático a las pestañas de Google Sheets.

## Qué se aplica

- Encabezados con fondo oscuro y texto blanco.
- Fila 1 congelada.
- Filtros activos por columna.
- Bordes en el rango de la tabla.
- Ajuste de texto.
- Autoajuste de columnas.
- Formato de fechas en columnas que contienen `fecha`.
- Formato numérico para columnas de cantidad, stock, saldo, entradas y salidas.
- Formato de porcentaje en columnas que contienen `porcentaje`.
- Anchos controlados para columnas largas como observaciones y detalle de salidas.

## Cómo usarlo

El formato se aplica automáticamente cuando se guardan datos en Google Sheets.

También se puede aplicar manualmente desde:

- `3️⃣ Administración → Diagnóstico → 🎨 Aplicar formato tabla a Google Sheets`
- `5️⃣ Kardex consolidado → 🎨 Formato tabla`

## Configuración opcional en Streamlit Secrets

```toml
FORMAT_GOOGLE_SHEETS_AS_TABLE = true
```

Si por alguna razón se desea desactivar el formato automático:

```toml
FORMAT_GOOGLE_SHEETS_AS_TABLE = false
```

## Nota técnica

Google Sheets no tiene el mismo objeto `Tabla` estructurada de Excel en todas las operaciones de la API. Por eso el sistema aplica un formato equivalente: encabezados, filtros, fila congelada, bordes, formatos de número/fecha y anchos ordenados.
