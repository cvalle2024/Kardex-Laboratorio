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


### Novedad V13: Kardex consolidado físico en Google Sheets

Además de verse dentro del sistema, la vista **Kardex consolidado por lote** ahora se guarda también como una pestaña real en la base de datos:

```text
Kardex_Consolidado
```

Esta pestaña muestra el stock actual por producto/lote con entrada total, salida acumulada, saldo actual, último destinatario, fecha de última salida y detalle de entregas.

La hoja se actualiza automáticamente al guardar movimientos, importar datos o modificar productos. También puede sincronizarse manualmente desde el menú **5️⃣ Kardex consolidado** o desde **3️⃣ Administración → Diagnóstico**.


### Novedad V16: formato tipo tabla en Google Sheets

Las pestañas de Google Sheets ahora pueden quedar ordenadas visualmente con formato tipo tabla:

- Encabezado oscuro con texto blanco.
- Filtros por columna.
- Fila de encabezado congelada.
- Bordes y ajuste de texto.
- Autoajuste de columnas.
- Formatos de fecha, número y porcentaje.
- Botón manual en **3️⃣ Administración → Diagnóstico → 🎨 Aplicar formato tabla a Google Sheets**.
- Botón manual en **5️⃣ Kardex consolidado → 🎨 Formato tabla**.

El formato también se aplica automáticamente al guardar información en Google Sheets.

Se puede activar/desactivar con este secreto:

```toml
FORMAT_GOOGLE_SHEETS_AS_TABLE = true
```



### Novedad V17: seguridad de sesión y navegación por rol

- Cierre automático de sesión por inactividad. El tiempo predeterminado es de 15 minutos.
- Se puede ajustar desde Streamlit Secrets con:

```toml
SESSION_TIMEOUT_MINUTES = 15
```

- Los usuarios que no tienen rol **Administrador** ya no ven el módulo **Administración** en el menú lateral ni en los accesos rápidos.
- La pantalla de login ya no muestra el usuario ni la contraseña inicial; solo solicita las credenciales asignadas y el PATH temporal generado automáticamente.

## Acceso inicial

Al ejecutar por primera vez, el sistema crea un usuario administrador inicial:

```text
Usuario: admin
Contraseña: admin123
```

En la pantalla de login el sistema mostrará un apartado llamado **Código PATH generado**. Copie ese código y péguelo en el campo **Pegar PATH generado** para completar el acceso. El código es temporal y se puede regenerar desde el mismo login. Por seguridad, la pantalla de login no muestra el usuario ni la contraseña inicial.

## Qué incluye

- Ruta de navegación guiada para que el usuario avance en el orden correcto del Kardex.
- Dashboard ejecutivo con KPIs de stock, productos activos, movimientos del mes y alertas.
- Registro de movimientos: Ingreso, Salida, Devolución, Corrección entrada y Corrección salida.
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
- Corrección entrada

Tipos negativos de stock:
- Salida
- Corrección salida

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


### Nota sobre Google Sheets

En `GOOGLE_SHEET_ID` puede pegar el ID puro o la URL completa de la hoja. El sistema extrae automáticamente el ID para evitar el error `APIError: [400]: Request contains an invalid argument`.

## Corrección V9: guardado obligatorio en Google Sheets

Esta versión evita una confusión importante: si `USE_GOOGLE_SHEETS = true` y la conexión falla, el sistema se detiene en vez de guardar silenciosamente en Excel local. Así se evita hacer pruebas pensando que se guardan en Google Sheets cuando realmente se estaban guardando en un archivo temporal local.

También se agregó en **Administración → Diagnóstico** el botón **Probar escritura en Google Sheets**, que agrega una fila de prueba en la pestaña `Config`.

### Secrets mínimos en Streamlit Cloud

```toml
USE_GOOGLE_SHEETS = true
ALLOW_LOCAL_FALLBACK = false
GOOGLE_SHEET_ID = "PEGUE_AQUI_EL_ID_O_URL_COMPLETA_DE_SU_HOJA_GOOGLE"

[gcp_service_account]
type = "service_account"
project_id = "TU_PROJECT_ID"
private_key_id = "TU_PRIVATE_KEY_ID"
private_key = "-----BEGIN PRIVATE KEY-----\nTU_PRIVATE_KEY\n-----END PRIVATE KEY-----\n"
client_email = "kardex-app@tu-proyecto.iam.gserviceaccount.com"
client_id = "TU_CLIENT_ID"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "TU_CLIENT_X509_CERT_URL"
universe_domain = "googleapis.com"
```

### Checklist si no guarda

1. La aplicación debe mostrar **Base activa: Google Sheets**.
2. El Google Sheet debe estar compartido como **Editor** con el `client_email` de la cuenta de servicio.
3. En Google Cloud deben estar habilitadas **Google Sheets API** y **Google Drive API**.
4. La hoja debe ser un Google Sheet real, no un Excel `.xlsx` solo subido a Drive.
5. En **Administración → Diagnóstico**, presione **Probar escritura en Google Sheets** y revise si aparece una fila en la pestaña `Config`.


## Mejora V11: control de cuota Google Sheets

Esta versión reduce el error `APIError [429]: Quota exceeded` usando:

- Lectura por lotes de todas las pestañas con una sola petición.
- Caché de datos por 60 segundos para evitar recargar la base en cada interacción de Streamlit.
- Reutilización de la conexión de Google Sheets con `st.cache_resource`.
- Diagnóstico sin re-leer todas las pestañas desde Google Sheets.

Si aparece un error 429, espere al menos 60 segundos y reinicie/recargue la app.

## V18 - Salidas por carrito, orden de compra, devoluciones y acta PDF

La versión V18 mejora el flujo operativo:

- Salidas con carrito de múltiples insumos.
- Registro individual por cada insumo en Google Sheets.
- Campo nuevo `orden_compra` en ingresos.
- Devolución desde una tabla de salidas históricas.
- Cambio de `Ajuste` a `Corrección`.
- Acta de entrega PDF generada al guardar una salida.

Para el acta PDF, asegúrese de que `requirements.txt` incluya:

```text
reportlab>=4.2.0
```

## Nota V19 — Salidas y stock disponible

El módulo de salidas solo muestra productos con saldo disponible calculado desde la hoja Movimientos. Los productos registrados en el catálogo no aparecen en salida hasta que se registre un movimiento de Ingreso, Devolución o Corrección entrada. Si no hay lotes disponibles, el sistema muestra una tabla de diagnóstico para orientar al usuario.


## V22 - Logo PNG en acta PDF

El acta de entrega usa `assets/logo_vihca.png` como logo oficial. También se incluye respaldo embebido para evitar problemas de rutas en Streamlit Cloud.

## Actualización V24 - CRUD, permisos y auditoría

La V24 agrega gestión profesional de catálogos y movimientos:

- CRUD controlado en catálogos.
- Permisos especiales por usuario desde Administración.
- Anulación lógica de movimientos en lugar de eliminación física.
- Auditoría completa de cambios.
- Nuevas hojas en la base: `Permisos_Usuarios` y `Auditoria_Cambios`.

Regla clave: los movimientos del Kardex no se eliminan; se anulan o corrigen para conservar trazabilidad.


## V25 - Corrección de estructura Google Sheets

Esta versión agrega migración automática para crear las hojas nuevas `Permisos_Usuarios` y `Auditoria_Cambios` si la base fue creada con una versión anterior. Se recomienda tener en Secrets: `AUTO_MIGRATE_GOOGLE_SHEETS = true`.


## Versión V26 — Arquitectura modular profesional

Esta versión reorganiza el sistema para que deje de depender de un único script largo. El archivo `app.py` queda como punto de entrada y la lógica se separa en módulos: configuración, utilidades, almacenamiento, servicios de Kardex, reportes, autenticación, componentes UI y páginas. Consulte `ARCHITECTURE.md` para el detalle técnico.
