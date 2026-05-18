# Arquitectura modular — Sistema Kardex PRO V26

Esta versión separa la aplicación en módulos funcionales para evitar que el sistema dependa de un único `app.py` extenso.

## Estructura principal

```text
app.py                         # Punto de entrada Streamlit
kardex_app/
  common.py                    # Configuración, constantes, columnas, datos iniciales
  main.py                      # Orquestador general de navegación
  auth.py                      # Login, PATH temporal, roles, permisos y sesión
  core/
    utils.py                   # Utilidades generales, limpieza, fechas, secretos
  storage/
    backends.py                # Excel local y Google Sheets
  services/
    kardex.py                  # Cálculo de stock, Kardex consolidado y KPIs
    reports.py                 # Reportes Excel, importación legacy y acta PDF
  ui/
    components.py              # Tema, tarjetas, secciones y componentes visuales
  pages/
    inicio.py                  # Ruta operativa del Kardex
    catalogos.py               # CRUD de catálogos
    movimientos.py             # Ingresos, salidas, devoluciones y correcciones
    kardex.py                  # Kardex consolidado
    stock.py                   # Stock y alertas
    dashboard.py               # Dashboard ejecutivo
    reportes.py                # Exportación de reportes
    importar.py                # Importación de Kardex anterior
    admin.py                   # Usuarios, permisos, auditoría y diagnóstico
```

## Principios aplicados

- `app.py` queda solo como entrada de ejecución.
- La lógica de conexión está separada de la lógica de negocio.
- El cálculo de stock y Kardex está centralizado en `services/kardex.py`.
- Los reportes y el acta PDF están centralizados en `services/reports.py`.
- Cada pantalla de Streamlit vive en un archivo independiente dentro de `pages/`.
- Los permisos, roles y auditoría quedan separados en `auth.py` y el módulo administrativo.

## Flujo del sistema

```text
Login → Catálogos → Administración opcional → Movimientos → Kardex/Stock → Reportes
```

## Recomendación de mantenimiento

Cuando se agregue una nueva funcionalidad, ubicarla así:

- Nueva pantalla: `kardex_app/pages/`
- Nueva regla de inventario: `kardex_app/services/kardex.py`
- Nuevo reporte o PDF: `kardex_app/services/reports.py`
- Nueva conexión o persistencia: `kardex_app/storage/backends.py`
- Nuevos permisos o login: `kardex_app/auth.py`
- Nuevo componente visual reusable: `kardex_app/ui/components.py`
