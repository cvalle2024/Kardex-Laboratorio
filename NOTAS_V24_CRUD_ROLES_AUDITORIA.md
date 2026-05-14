# V24 - CRUD controlado, permisos por usuario y auditoría

Esta versión agrega una capa profesional de gestión CRUD sin eliminar físicamente registros sensibles del Kardex.

## Cambios principales

1. **Catálogos con CRUD controlado**
   - Productos, proveedores, solicitantes y personal ahora permiten consultar, editar, desactivar y reactivar registros.
   - No se eliminan físicamente registros: se usa desactivación lógica para proteger la trazabilidad histórica.
   - Cada modificación pide motivo obligatorio y se registra en auditoría.

2. **Permisos especiales por usuario**
   - Se agrega la hoja `Permisos_Usuarios`.
   - El administrador puede habilitar o negar permisos individuales por usuario.
   - Un usuario Operador puede recibir permiso especial para editar o desactivar productos, proveedores, solicitantes, personal o gestionar movimientos.

3. **Movimientos protegidos**
   - Los movimientos no se eliminan.
   - Se pueden anular con motivo obligatorio si el usuario tiene permiso.
   - Los movimientos anulados no afectan el cálculo de stock ni el Kardex consolidado.
   - Se permite edición administrativa limitada de movimientos: fecha, proveedor, solicitante, personal, orden de compra y observación.
   - Para cambios de cantidad o tipo de movimiento se debe usar corrección o anulación, no edición directa.

4. **Auditoría del sistema**
   - Se agrega la hoja `Auditoria_Cambios`.
   - Registra usuario, rol, acción, módulo, registro, campo, valor anterior, valor nuevo, motivo y detalle.

5. **Roles incluidos**
   - Administrador
   - Supervisor
   - Operador
   - Consulta

## Recomendación operativa

- Producción: mantener `ALLOW_LOCAL_FALLBACK = false`.
- En Google Sheets se crearán las nuevas hojas automáticamente: `Permisos_Usuarios` y `Auditoria_Cambios`.
- Después de subir la versión, hacer Reboot en Streamlit Cloud.
