# KARDEX PRO | Sistema en Streamlit

Sistema web para control de reactivos e insumos con base de datos en Excel local o Google Sheets.

## Novedades de esta versión

- Nueva navegación tipo **Ruta Kardex**, ordenada según el flujo real del proceso.
- Nuevo módulo **Inicio / Ruta del Kardex** con guía de pasos, estado de configuración y acceso rápido al siguiente paso recomendado.
- Menú principal reordenado: Inicio, Catálogos, Administración, Movimientos, Kardex consolidado, Stock, Dashboard, Reportes e Importación.
- Nueva vista **Kardex consolidado por lote**: una sola fila por producto/lote donde se ve ingreso, salida acumulada, último destinatario, fecha última salida y saldo actual.
- Exportación de **Kardex_Consolidado** dentro del reporte completo en Excel.
- Login inicial con **usuario, contraseña y PATH temporal generado automáticamente**.
- Formularios amigables y separados para:
  - Productos
  - Proveedores
  - Solicitantes / unidades
  - Personal
  - Usuarios del sistema
- Formularios de movimiento por bloques: datos generales, producto/lote, origen/destino y observación.
- Panel de administración para crear usuarios, revisar la seguridad PATH automática y verificar la estructura de la base.
- Validaciones más claras para evitar salidas mayores al stock disponible.
- Mantenimiento de catálogos con formulario ordenado y tabla editable para ajustes masivos.

## Acceso inicial

Al ejecutar por primera vez, el sistema crea un usuario administrador inicial:

```text
Usuario: admin
Contraseña: admin123
```

En la pantalla de login el sistema mostrará un apartado llamado **Código PATH generado**. Copie ese código y péguelo en el campo **Pegar PATH generado** para completar el acceso. El código es temporal y se puede regenerar desde el mismo login.

## Qué incluye

- Ruta de navegación guiada para que el usuario avance en el orden correcto del Kardex.
- Dashboard ejecutivo con KPIs de stock, productos activos, movimientos del mes y alertas.
- Registro de movimientos: Ingreso, Salida, Devolución, Ajuste entrada y Ajuste salida.
- Control de stock por producto, marca, lote, vencimiento y unidad.
- Validación para evitar salidas mayores al stock disponible.
- Alertas automáticas por vencimiento, productos vencidos y stock bajo.
- Catálogos editables: Productos, Proveedores, Solicitantes y Personal.
- Gestión básica de usuarios y roles: Administrador, Operador y Consulta.
- Reportes filtrables y exportación completa a Excel, incluyendo Kardex consolidado.
- Importador del Kardex anterior con macros usando la hoja `MOVIMIENTO`.
- Plantilla de base de datos: `data/kardex_db_plantilla.xlsx`.

## Instalación local

1. Cree una carpeta para el sistema y copie todos estos archivos.
2. Abra una terminal dentro de la carpeta.
3. Instale dependencias:

```bash
pip install -r requirements.txt
```

4. Ejecute el sistema:

```bash
streamlit run app.py
```

El sistema creará automáticamente `data/kardex_db.xlsx` si no existe. Si falta alguna hoja, también la crea automáticamente.

## Uso con Excel local

No requiere configuración adicional. La información se guarda en:

```text
data/kardex_db.xlsx
```

La plantilla inicial está en:

```text
data/kardex_db_plantilla.xlsx
```

## Uso con Google Sheets

1. Cree una hoja de Google Sheets vacía.
2. Copie el ID de la hoja desde la URL.
3. Cree una cuenta de servicio en Google Cloud y descargue el JSON de credenciales.
4. Comparta la hoja de Google con el correo `client_email` de la cuenta de servicio.
5. Copie `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml`.
6. Cambie:

```toml
USE_GOOGLE_SHEETS = true
GOOGLE_SHEET_ID = "ID_DE_LA_HOJA"
```

7. Complete los datos de `[gcp_service_account]` con el JSON de la cuenta de servicio.

## Modelo de datos

### Productos
Catálogo maestro de reactivos e insumos. Incluye stock mínimo y días de alerta por vencimiento.

### Proveedores
Catálogo de proveedores con datos generales, representante, teléfono, correo y dirección.

### Solicitantes
Unidades, áreas, sitios o personas que solicitan productos.

### Personal
Personal que entrega, recibe o registra movimientos.

### Usuarios
Usuarios del sistema con rol, contraseña cifrada con SHA-256 y validación de PATH temporal generado automáticamente en el login.

### Movimientos
Historial transaccional. Esta hoja es la fuente principal para calcular stock.

Tipos positivos de stock:
- Ingreso
- Devolución
- Ajuste entrada

Tipos negativos de stock:
- Salida
- Ajuste salida

### Stock actual
No se registra manualmente. El sistema lo calcula agrupando movimientos por producto, marca, lote, vencimiento y unidad.

### Kardex consolidado
No reemplaza la hoja `Movimientos`; es una vista calculada para lectura operativa. Agrupa cada producto/lote en una sola fila y muestra:

- Fecha de ingreso
- Proveedor de ingreso
- Entrada total
- Salida acumulada
- Último destinatario
- Fecha última salida
- Detalle de salidas
- Saldo actual

## Ruta operativa recomendada

1. Iniciar sesión con usuario, contraseña y PATH temporal.
2. Completar catálogos base: productos, proveedores, solicitantes y personal.
3. Administrar usuarios y roles si el usuario tiene rol de Administrador.
4. Registrar movimientos: ingresos, salidas, devoluciones y ajustes.
5. Revisar Kardex consolidado y stock/alertas.
6. Generar reportes y exportar a Excel.
7. Importar Kardex anterior solo cuando se necesite migrar información histórica.

## Recomendación operativa

Para inventarios de reactivos con vencimiento, se recomienda registrar siempre:

- Lote
- Fecha de vencimiento
- Unidad
- Cantidad
- Proveedor o solicitante según aplique
- Personal que entrega o recibe

Esto permite trazabilidad, alertas y reportes consistentes.

## Versión 7.0 - Lista para GitHub, Streamlit Cloud y Google Sheets

Cambios principales:

- Se agregó `.gitignore` para evitar subir credenciales y la base local generada.
- Se agregó `DEPLOY_STREAMLIT_CLOUD.md` con pasos para GitHub + Streamlit Community Cloud + Google Sheets.
- Se mejoró la lectura de `USE_GOOGLE_SHEETS` para aceptar valores booleanos o texto en Streamlit Secrets.
- Se mantiene Excel local como modo de respaldo si Google Sheets no está configurado.

## Versión 6.0 - Navegación guiada y estructura Kardex

Cambios principales:

- Se agregó una pantalla inicial llamada **Inicio / Ruta del Kardex**.
- El menú lateral ahora está ordenado en la secuencia lógica del trabajo real.
- La pantalla inicial muestra el avance de configuración: productos, proveedores, solicitantes, personal y movimientos.
- El sistema sugiere el siguiente paso: completar catálogos, administrar usuarios, registrar ingresos o revisar reportes.
- Se mantienen las mejoras del formulario de movimientos: primero producto/marca/lote y autocompletado desde catálogo.
- Se conserva la vista de Kardex consolidado y el reporte exportable en Excel.
