# V25 - Fix estructura Google Sheets después de agregar CRUD, roles y auditoría

Esta versión corrige el problema donde la conexión con Google Sheets se crea correctamente,
pero el sistema se detiene al leer la estructura de la base después de actualizar a la versión
con CRUD, permisos y auditoría.

## Causa principal

La V24 agregó dos hojas nuevas:

- Permisos_Usuarios
- Auditoria_Cambios

Si la base de Google Sheets fue creada con una versión anterior, esas pestañas pueden no existir
y Google Sheets puede devolver error al leer por lote.

## Corrección

- Se agregó migración automática de estructura.
- Si falta una pestaña base, el sistema la crea automáticamente y vuelve a intentar la lectura.
- No borra registros existentes.
- Solo crea pestañas faltantes y actualiza encabezados esperados en la fila 1.
- Kardex_Consolidado sigue siendo una hoja calculada; no es obligatoria para iniciar.

## Secret recomendado

```toml
AUTO_MIGRATE_GOOGLE_SHEETS = true
AUTO_CREATE_SHEETS = true
USE_GOOGLE_SHEETS = true
ALLOW_LOCAL_FALLBACK = false
```

## Después de subir a GitHub

1. Subir esta versión al repositorio.
2. Hacer commit.
3. En Streamlit Cloud presionar Reboot app.
4. Esperar 1 minuto antes de probar.
5. Entrar nuevamente al sistema.

