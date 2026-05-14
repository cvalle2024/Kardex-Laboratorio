# V21 - Corrección de acta PDF y selección en devolución

## Cambios incluidos

1. **Logo del acta de entrega**
   - Se agregó el archivo `assets/logo_vihca.png` dentro del paquete del sistema.
   - El PDF ahora usa el logo institucional como imagen local, sin depender de internet, Drive ni rutas temporales de Streamlit Cloud.
   - Si por alguna razón el archivo no existe, el sistema usa un encabezado textual de respaldo.

2. **Selección de salida para devolución**
   - La tabla de salidas ahora permite selección de fila con clic usando `st.dataframe` con `selection_mode="single-row"`.
   - Al seleccionar una fila, el formulario de devolución se llena automáticamente con producto, marca, lote, unidad, vencimiento y sitio.
   - Se mantiene un selector de respaldo debajo de la tabla para navegadores o versiones donde el clic no marque la fila.

3. **Validación más segura**
   - Se agregó verificación para evitar errores cuando una selección no coincide con ninguna fila.
