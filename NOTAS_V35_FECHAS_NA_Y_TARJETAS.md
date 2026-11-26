# V35 — Fechas opcionales y tarjetas HTML

- Ingreso: calendarios independientes de elaboración y vencimiento, inicialmente vacíos; casillas N/A independientes. Si la casilla está marcada o la fecha queda vacía se guarda el texto N/A. La casilla prevalece sobre una fecha seleccionada.
- Excel: se conserva el texto N/A al leer movimientos y Kardex consolidado. En los cálculos, vencimientos vacíos y N/A se tratan como la misma clave de lote; no se generan alertas de vencimiento sin fecha. No se alteran los registros históricos directamente.
- Edición: se permite conservar o retirar fechas sin inventar una fecha; los campos históricos vacíos se conservan al editar otros datos.
- Salida: se corrige la tarjeta del responsable, cargo, departamento y teléfono mediante st.html. Se aplica la misma solución a la ficha del producto y al resumen del movimiento seleccionado; se escapan textos dinámicos en esas tarjetas.
- Sin cambios en permisos, órdenes de compra, reglas de confirmación, cantidades, estructura de hojas ni generación de actas.

Pruebas locales: compilación Python completa; ejecución aislada de funciones de fechas, stock y consolidado con ingresos, salidas, devoluciones y ajustes; saldo esperado 7; alertas con fechas reales; lectura/escritura Excel con N/A. No se ejecutó la interfaz completa de Streamlit ni se conectó a Google Sheets de producción (dependencias/credenciales no disponibles en el entorno de revisión).
