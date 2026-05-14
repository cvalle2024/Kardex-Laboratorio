# V19 — Diagnóstico de stock para salidas

Esta versión mejora el módulo de salida cuando el sistema no encuentra lotes con stock disponible.

## Cambio principal

Antes el sistema solo mostraba:

> No hay lotes con stock disponible para registrar salidas.

Ahora muestra un diagnóstico con:

- Productos registrados en catálogo.
- Movimientos registrados.
- Movimientos que suman stock: Ingreso, Devolución, Corrección entrada.
- Movimientos que restan stock: Salida, Corrección salida.
- Lotes calculados en stock.
- Lotes con saldo mayor a cero.

## Importante

Los productos del catálogo no generan existencia por sí solos. Para que un producto aparezca en el carrito de salida debe existir al menos un movimiento de ingreso con cantidad disponible.

## Botón agregado

En la pantalla de salidas se agregó:

- Actualizar datos desde Google Sheets

Esto limpia caché y vuelve a leer la base cuando se están haciendo pruebas.
