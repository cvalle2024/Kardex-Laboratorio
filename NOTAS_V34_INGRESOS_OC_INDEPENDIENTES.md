# V34 — Reingreso de productos existentes por nuevas órdenes de compra

- La orden de compra pertenece a cada **movimiento de ingreso**, no a la identidad del producto.
- En `Registrar movimientos > Ingreso`, seleccione el producto existente mediante ID, código o nombre, registre lote, cantidad y OC nueva. El mismo producto_id/lote puede recibirse de nuevo; cada recepción se guarda con un movimiento_id nuevo.
- `Ajuste (+)` permanece reservado para diferencias de inventario o correcciones de existencias, no para registrar compras normales.
- En `Agregar producto`, la validación impide crear otro registro con el mismo nombre normalizado aunque se escriba una orden distinta. La antigua columna de OC del catálogo se conserva por compatibilidad, pero es histórica/opcional: la OC operativa va en Movimientos.
- No se modifican automáticamente ni fusionan productos históricos duplicados por la antigua regla nombre+OC; deben revisarse antes de una consolidación/migración de IDs.
- El Kardex existente agrega stock por producto, marca, lote, vencimiento y unidad. Si varias órdenes alimentan exactamente el mismo lote, sus cantidades se suman; la OC exacta de cada compra permanece en la hoja Movimientos. El campo `orden_compra_ingreso` en vistas consolidadas puede mostrar solo la primera OC y NO debe interpretarse como un desglose de adquisiciones. Una salida del lote compartido no se asigna automáticamente a una OC individual.
- Antes de poner en producción, validar flujo con copia de Google Sheets y confirmar stock, salidas y actas para lotes financiados por órdenes distintas.
