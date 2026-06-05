# V30 - Validación de ingreso por producto/lote

## Cambio principal

Se agregó una validación en el módulo **Registrar movimientos > Ingreso** para evitar duplicar por error un producto que ya fue ingresado anteriormente con el mismo lote.

## Reglas aplicadas

1. **Solo permite guardar ingresos de productos activos en catálogo.**
   - Si el producto fue desactivado o ya no existe en catálogo, el ingreso no se guarda.
   - El formulario de ingreso sigue tomando producto, marca y unidad desde el catálogo activo.

2. **Bloquea un segundo ingreso para el mismo producto y lote.**
   - La comparación usa `producto_id + lote`.
   - El lote se normaliza para evitar duplicados por diferencias de mayúsculas, espacios, guiones, barras o guiones bajos.
   - Ejemplo: `AB-2026-001`, `ab 2026 001` y `AB/2026/001` se consideran el mismo lote.

3. **Si el lote ya existe, se muestra un mensaje operativo.**
   - El sistema informa entrada acumulada, salida acumulada y saldo actual.
   - Indica que, para sumar más unidades al mismo lote, debe usarse **Corrección entrada / Ajuste (+)**.

4. **La validación se ejecuta dos veces.**
   - Al presionar **Guardar ingreso**.
   - Al confirmar en la ventana emergente.
   Esto evita que se guarde un duplicado si otro usuario agregó el mismo producto/lote mientras estaba pendiente la confirmación.

## Mejora técnica adicional

La confirmación de ingreso ahora guarda todos los datos del formulario en un payload temporal antes de confirmar. Esto evita pérdida de datos por el comportamiento `clear_on_submit=True` de Streamlit y hace más segura la ventana emergente de confirmación.
