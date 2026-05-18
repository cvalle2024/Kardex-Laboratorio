# V26 — Arquitectura modular profesional

Se reorganizó el sistema Kardex PRO en una estructura modular.

## Cambios clave

- `app.py` ahora solo inicia la aplicación.
- Se creó el paquete `kardex_app`.
- Se separó la lógica en capas:
  - Core/utilidades
  - Storage/Google Sheets/Excel
  - Servicios de Kardex y reportes
  - Autenticación y permisos
  - Componentes UI
  - Páginas funcionales

## Beneficios

- Mejor mantenimiento.
- Menos riesgo de dañar todo el sistema al modificar una sección.
- Estructura más profesional para GitHub y despliegue en Streamlit Cloud.
- Facilita agregar nuevas funciones: compras, actas, auditoría avanzada, reportes PDF y permisos.
