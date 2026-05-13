# V15 — Corrección de lectura de estructura Google Sheets

## Problema corregido
En versiones anteriores, el sistema intentaba leer la hoja `Kardex_Consolidado` como si fuera una hoja base obligatoria al iniciar.

Como `Kardex_Consolidado` es una hoja calculada desde `Movimientos` y `Productos`, si esa pestaña física todavía no existía en Google Sheets, o si Google limitaba lecturas de estructura, el sistema podía detenerse con el mensaje:

> La conexión se creó, pero falló al leer la estructura de la base.

## Corrección aplicada
- El sistema ahora solo lee al iniciar las hojas base:
  - Productos
  - Proveedores
  - Solicitantes
  - Personal
  - Usuarios
  - Movimientos
  - Config
- `Kardex_Consolidado` se calcula dentro del sistema desde los movimientos.
- `Kardex_Consolidado` se crea/actualiza físicamente en Google Sheets al sincronizar.
- El guardado en Google Sheets ahora asegura que la pestaña exista antes de escribir.

## Qué hacer después de desplegar
1. Subir la V15 a GitHub.
2. Hacer commit.
3. Reiniciar la app en Streamlit Cloud.
4. Entrar al módulo `5️⃣ Kardex consolidado`.
5. Presionar `🔄 Actualizar hoja Kardex_Consolidado` una vez.

Con eso la hoja física quedará creada/actualizada en Google Sheets.
