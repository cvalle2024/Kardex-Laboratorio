# V17 - Seguridad de sesión y navegación por rol

Cambios incluidos:

- Cierre automático de sesión por inactividad.
- Tiempo predeterminado: 15 minutos.
- Ajuste opcional en Secrets: `SESSION_TIMEOUT_MINUTES = 15`.
- El navegador fuerza una recarga después del periodo de inactividad para que Streamlit cierre la sesión.
- Los usuarios que no son Administrador no ven el módulo **3️⃣ Administración** en la navegación ni en accesos rápidos.
- La pantalla de login ya no muestra el usuario ni la contraseña inicial.
- Se mantiene visible el PATH temporal generado porque forma parte de la validación dinámica solicitada.
