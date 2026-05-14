# V23 - Texto dinámico por categoría en acta de entrega

## Mejora agregada

En la generación del acta de entrega, el texto:

> se hace entrega formal de los siguientes reactivos e insumos

ahora se calcula automáticamente según la categoría registrada en el catálogo de productos.

Ejemplos:

- Si todos los productos son categoría `Insumo`, el acta dirá: `los siguientes insumos`.
- Si todos los productos son categoría `Reactivo`, el acta dirá: `los siguientes reactivos`.
- Si hay mezcla de categorías, el acta dirá: `los siguientes reactivos e insumos`, `reactivos, insumos y equipos`, etc.

## Campo editable antes de generar el acta

En el formulario de salida se agregó el campo:

`Texto de categoría para el acta`

El sistema lo llena automáticamente según el carrito, pero el usuario puede ajustarlo antes de guardar y generar el PDF.

## No cambia la base transaccional

La categoría se toma desde la hoja `Productos`. No se agregó como columna obligatoria en `Movimientos`, para no romper la estructura de la base ya desplegada.
