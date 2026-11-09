# V33 — Tres selectores y clave compuesta de productos

## Movimientos
- En **Ingreso**, **Salida / carrito** y **Corrección entrada/salida (ajustes)** se muestran tres columnas sincronizadas: `producto_id`, `Código del producto`, `Producto`.
- Las etiquetas de código y nombre muestran también el ID para desambiguar códigos/nombres repetidos. En nombre se muestra además la orden de compra.
- La selección se resuelve internamente por `producto_id`, no por nombre.
- En salida, la opción `Todos` sigue mostrando todos los lotes con existencias; en ajustes, la selección del producto precede al lote.
- La orden de compra de ingreso se carga desde la ficha del catálogo y se comprueba que corresponda al mismo producto, también en la confirmación.

## Catálogo de productos
- Se agrega **`orden_compra`** al FINAL de la estructura `Productos`, sin alterar la posición de las columnas existentes.
- La identidad lógica se valida usando **(nombre_producto, orden_compra)** con insensibilidad a mayúsculas y espacios repetidos. Nombres iguales con órdenes diferentes se admiten como productos diferentes y reciben diferentes `producto_id`.
- La misma combinación se bloquea tanto al registrar como al editar. Se comprueba de nuevo justo antes del guardado.
- Los registros anteriores no se eliminan ni fusionan: las órdenes de compra anteriores pueden aparecer vacías hasta que se completen en **Catálogos > Productos > Editar**.
- Para nuevas altas, nombre y orden de compra son obligatorios. El código interno no es la clave única.

## Google Sheets
- La ampliación de la hoja Productos verifica si necesita añadir una columna al final antes de escribir el encabezado. No se reordena ni se borra la información anterior en la migración del encabezado.

## Limitaciones
- La validación de duplicados es a nivel de aplicación; en usos simultáneos con múltiples sesiones, Google Sheets no ofrece una restricción única transaccional. Es prudente evitar guardar dos altas idénticas al mismo instante.
- Esta revisión supera la comprobación de sintaxis de Python y pruebas aisladas de la clave compuesta; no se probó el flujo completo contra una cuenta real de Google Sheets.
