# V37 — Normalización de fechas de lotes

- Unifica seriales de Excel/Google Sheets, ISO y DD/MM/AAAA en una clave ISO; preserva N/A y campos vacíos.
- Las fechas de elaboración y vencimiento se normalizan al calcular y al guardar en Google Sheets.
- No cambia las reglas de movimientos, agrupación por producto/lote/marca/unidad ni el carrito V36.
- Prueba con exportable 26/11/2026: 436 movimientos, 42 productos; 62 a 56 grupos; 7 a 1 saldos negativos. Ambos lotes Histo quedan 10 entradas, 9 salidas, 1 saldo.
- Actualización: reemplazar app.py, reiniciar Streamlit, entrar a Kardex consolidado por lote y pulsar Actualizar hoja Kardex_Consolidado. No volver a registrar salidas.
- Pendiente independiente: PRD-0034 tiene una variante de marca que genera un grupo negativo. Verificar contra el soporte original antes de editar el movimiento.
- Pruebas locales de cálculo y sintaxis; conexión Google Sheets/Streamlit de producción no probada.
