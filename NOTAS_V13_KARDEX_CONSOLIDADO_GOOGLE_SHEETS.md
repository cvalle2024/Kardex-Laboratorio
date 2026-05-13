# Notas V13 — Kardex consolidado guardado en Google Sheets

## Problema corregido

En versiones anteriores, la pantalla **Kardex consolidado** existía dentro de Streamlit, pero la tabla no quedaba guardada físicamente como pestaña en Google Sheets.

## Cambio aplicado

La versión V13 agrega una hoja física llamada:

```text
Kardex_Consolidado
```

Esta hoja se crea automáticamente en Google Sheets o Excel local y se actualiza con la vista consolidada por producto/lote.

## Columnas principales

- estado
- producto_id
- producto
- marca
- lote
- unidad
- fecha_ingreso
- proveedor_ingreso
- fecha_elaboracion
- fecha_vencimiento
- entrada_total
- salida_total
- saldo_actual
- porcentaje_consumido
- numero_salidas
- fecha_ultima_salida
- ultimo_entregado_a
- ultimo_personal_entrega
- detalle_salidas
- observacion_ingreso
- dias_para_vencer
- stock_minimo

## Cuándo se actualiza

La hoja `Kardex_Consolidado` se actualiza automáticamente cuando:

1. Se registra un movimiento de ingreso, salida, devolución o ajuste.
2. Se importa el Kardex anterior.
3. Se edita el catálogo de productos.

También se agregó un botón manual en:

```text
5️⃣ Kardex consolidado → 🔄 Actualizar hoja Kardex_Consolidado
```

Y otro botón en:

```text
3️⃣ Administración → Diagnóstico → 🔄 Crear/actualizar hoja Kardex_Consolidado
```

## Lógica

La hoja `Movimientos` sigue siendo la bitácora transaccional oficial.

La hoja `Kardex_Consolidado` no se digita manualmente; se recalcula desde `Movimientos` y `Productos` para mostrar:

```text
Entrada total - Salida total = Saldo actual
```

Esto permite revisar directamente en Google Sheets cuál es el stock actual en existencia por producto, marca, lote, vencimiento y unidad.
