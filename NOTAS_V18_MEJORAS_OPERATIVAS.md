# V18 - Mejoras operativas de movimientos y acta de entrega

Esta versión incorpora mejoras solicitadas para el flujo real de Kardex:

## Salidas tipo carrito
- La opción **Salida** permite seleccionar varios reactivos/insumos antes de guardar.
- Al guardar, cada insumo del carrito se registra como una fila individual en la hoja `Movimientos`.
- Todos los registros de una misma salida quedan vinculados con el mismo `acta_entrega_id`.

## Nueva variable en ingresos
- Se agregó la columna `orden_compra` en la hoja `Movimientos`.
- En el formulario de ingreso se agregó el campo **Orden de compra**.

## Devoluciones guiadas
- La opción **Devolución** muestra una tabla con salidas registradas.
- Al seleccionar una salida, el sistema llena automáticamente producto, marca, lote, unidad y vencimiento.
- Solo quedan habilitados los campos operativos: cantidad a devolver, quién devuelve, personal que recibe y observación.

## Correcciones
- Se reemplazó la palabra **Ajuste** por **Corrección**.
- Se mantienen compatibilidades internas con registros antiguos que digan `Ajuste entrada` o `Ajuste salida`.
- Al seleccionar el producto/lote, el sistema autocompleta la información y permite registrar cantidad, responsable y justificación.

## Acta de entrega en PDF
- Al guardar una salida, el sistema puede generar un PDF de acta de entrega.
- El acta agrupa todos los insumos enviados a un mismo sitio/unidad solicitante.
- El diseño sigue la estructura de la referencia: fecha, destinatario, texto institucional, tabla de insumos, firmas y nota de sello/firma.
- Se agregó `reportlab` en `requirements.txt`.

## Google Sheets
- Las nuevas columnas se crean en la estructura de la base.
- La hoja `Movimientos` ahora incluye `orden_compra` y `acta_entrega_id`.
- El Kardex consolidado incluye `orden_compra_ingreso`.
