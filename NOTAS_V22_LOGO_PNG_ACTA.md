# V22 - Logo PNG en acta PDF

Corrección aplicada:

- El acta de entrega ahora usa el archivo PNG real `assets/logo_vihca.png`.
- El logo también queda embebido en base64 dentro de `app.py` como respaldo para Streamlit Cloud.
- Si el archivo PNG no existe en el servidor, el sistema reconstruye automáticamente el logo en `/tmp/kardex_logo_vihca.png`.
- Se eliminó el encabezado textual anterior que podía verse deformado o montado sobre el logo.

Para cambiar el logo en el futuro, reemplace:

```text
assets/logo_vihca.png
```

por el PNG oficial del programa y mantenga el mismo nombre.
