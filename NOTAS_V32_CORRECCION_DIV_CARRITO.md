# V32 — Corrección de contenedor en el carrito de salida

En **Registrar movimientos > Salida – Entrega a un sitio o unidad > Armar carrito de salida**, se eliminó el par de etiquetas HTML `<div class='form-card'>` y `</div>` que se enviaban en invocaciones separadas a `st.markdown`. En Streamlit ese patrón no garantiza un único contenedor HTML envolviendo widgets.

El área del carrito se agrupa ahora con `with st.container(border=True):`, sin etiquetas DIV manuales. Se mantienen la búsqueda en dos columnas (producto_id y Producto), sincronización, selección de lotes y validaciones de stock.

Comprobación: análisis sintáctico completo de `app.py` mediante `ast.parse`.
