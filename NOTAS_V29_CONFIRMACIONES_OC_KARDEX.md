# NOTAS V29 — Confirmaciones, orden de Kardex y orden de compra en acta

## Cambios incluidos

1. **Confirmación antes de crear catálogos**
   - Productos
   - Proveedores
   - Solicitantes / unidades
   - Personal

   Al presionar guardar, el sistema ahora muestra una confirmación con el resumen del registro. Solo guarda cuando el usuario confirma.

2. **Tabla Kardex consolidado por lote reordenada**
   - Nuevo orden principal:
     1. estado
     2. entrada_total
     3. salida_total
     4. saldo_actual
     5. producto
     6. marca
     7. lote
     8. unidad
     9. resto de trazabilidad

   El mismo orden se aplica a la hoja física `Kardex_Consolidado` cuando se sincroniza.

3. **Orden de compra en acta de entrega**
   - Se corrigió el cálculo de stock para conservar la `orden_compra_ingreso` del ingreso original del lote.
   - Cuando se agrega un lote al carrito de salida, el sistema hereda esa orden de compra.
   - Al generar el PDF del acta, la columna “ORDEN DE COMPRA” toma el valor heredado del ingreso.

## Causa del fallo detectado

La salida intentaba leer `orden_compra_ingreso` desde la tabla de stock, pero el cálculo de stock no estaba trayendo esa columna desde los movimientos de ingreso. Por eso el carrito quedaba con orden de compra vacía y el acta imprimía “—”.
