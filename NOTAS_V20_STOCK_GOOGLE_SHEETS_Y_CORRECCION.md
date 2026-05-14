# V20 — Stock verificado contra Google Sheets y corrección de selección

## Mejoras incluidas

1. **Stock operativo verificado contra Google Sheets**
   - En el módulo de movimientos, el sistema refresca las hojas operativas desde Google Sheets con caché corta.
   - La salida por carrito usa la hoja `Movimientos` y el catálogo `Productos` como fuente oficial para calcular stock.
   - La hoja `Kardex_Consolidado` se lee como diagnóstico visual, pero no reemplaza la bitácora transaccional.

2. **Reconocimiento flexible de tipos de movimiento**
   - Ahora reconoce como entradas valores escritos como `Ingreso`, `Entrada`, `Recepción`, `Devolución`, `Corrección entrada` y `Ajuste entrada`.
   - Reconoce como salidas valores como `Salida`, `Egreso`, `Entrega`, `Despacho`, `Corrección salida` y `Ajuste salida`.
   - Esto ayuda si existen registros editados manualmente o importados desde versiones anteriores.

3. **Diagnóstico cuando no aparece stock**
   - Si no hay stock para salida, muestra conteo de productos, movimientos, movimientos que suman/restan stock, lotes calculados y saldos en `Kardex_Consolidado`.
   - Si hay saldo en `Kardex_Consolidado` pero no existen ingresos válidos en `Movimientos`, el sistema lo advierte para evitar salidas sin respaldo transaccional.

4. **Corrección del error `single positional indexer is out-of-bounds`**
   - Ocurría cuando después de filtrar lotes para `Corrección salida` ya no quedaban filas, pero el sistema intentaba seleccionar `.iloc[0]`.
   - Ahora valida si existen lotes antes de mostrar el selector.

## Recomendación operativa

Para que un producto aparezca en salida, debe existir en `Movimientos` una entrada válida con cantidad numérica y lote. El catálogo `Productos` solo define el producto; no crea existencia.
