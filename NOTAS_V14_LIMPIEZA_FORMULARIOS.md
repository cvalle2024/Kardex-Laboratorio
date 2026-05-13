# V14 - Limpieza de formularios y mensajes de guardado

Cambios principales:

- Los formularios de registro ahora usan `clear_on_submit=True`.
- El formulario de movimientos usa una clave dinámica (`mov_form_nonce`) para limpiar correctamente campos con claves persistentes como marca, lote, fechas, costo y observación.
- Se agregó un sistema de mensajes temporales tipo `flash_message` para que la confirmación de guardado no desaparezca al hacer `rerun()`.
- Después de guardar, el usuario verá mensajes claros como:
  - Movimiento guardado correctamente.
  - Producto guardado correctamente.
  - Proveedor guardado correctamente.
  - Solicitante guardado correctamente.
  - Personal guardado correctamente.
  - Usuario creado correctamente.
  - Importación completada.
- Los mensajes también indican cuando el formulario quedó limpio para continuar registrando.

Objetivo:

Evitar que, después de guardar un registro, los campos queden llenos y el usuario piense que el movimiento/producto/proveedor no se guardó o que debe borrar manualmente los campos.
