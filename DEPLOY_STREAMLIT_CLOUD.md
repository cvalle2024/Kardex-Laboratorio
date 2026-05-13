# Despliegue en GitHub + Streamlit Community Cloud + Google Sheets

Esta versión está preparada para trabajar en línea usando:

- GitHub como repositorio del código.
- Streamlit Community Cloud como servidor web.
- Google Sheets como base de datos.
- Streamlit Secrets para guardar credenciales sin subirlas a GitHub.

## 1. Archivos que debe subir a GitHub

Suba estos archivos y carpetas:

```text
app.py
requirements.txt
README.md
MODELO_LOGICO.md
DEPLOY_STREAMLIT_CLOUD.md
.streamlit/config.toml
.streamlit/secrets.toml.example
data/kardex_db_plantilla.xlsx
.gitignore
```

No suba:

```text
.streamlit/secrets.toml
*.json
data/kardex_db.xlsx
```

El archivo `.gitignore` ya queda configurado para evitar subir credenciales y la base Excel local.

## 2. Crear la hoja de Google Sheets

1. Cree una hoja vacía en Google Sheets.
2. Copie el ID de la hoja desde la URL.

Ejemplo:

```text
https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit#gid=0
```

El sistema puede crear automáticamente las hojas internas cuando se conecte por primera vez:

- Productos
- Proveedores
- Solicitantes
- Personal
- Usuarios
- Movimientos
- Config

## 3. Crear cuenta de servicio en Google Cloud

1. Entre a Google Cloud Console.
2. Cree o seleccione un proyecto.
3. Active Google Sheets API.
4. Cree una cuenta de servicio.
5. Genere una llave JSON.
6. Copie el valor `client_email` del JSON.
7. Comparta la hoja de Google Sheets con ese correo como **Editor**.

Importante: la cuenta de servicio no tiene acceso a la hoja hasta que usted comparta el Google Sheet con el `client_email`.

## 4. Secretos para Streamlit Cloud

En Streamlit Cloud, entre a:

```text
App → Settings → Secrets
```

Pegue este bloque, reemplazando los datos por los del JSON de Google Cloud:

```toml
USE_GOOGLE_SHEETS = true
ALLOW_LOCAL_FALLBACK = false
FORMAT_GOOGLE_SHEETS_AS_TABLE = true
GOOGLE_SHEET_ID = "PEGUE_AQUI_EL_ID_O_URL_COMPLETA_DE_SU_HOJA_GOOGLE"

[gcp_service_account]
type = "service_account"
project_id = "su-project-id"
private_key_id = "su-private-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\nPEGUE_AQUI_LA_LLAVE_PRIVADA\n-----END PRIVATE KEY-----\n"
client_email = "cuenta-servicio@su-proyecto.iam.gserviceaccount.com"
client_id = "su-client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/cuenta-servicio%40su-proyecto.iam.gserviceaccount.com"
universe_domain = "googleapis.com"
```

## 5. Desplegar en Streamlit Cloud

1. Cree un repositorio en GitHub.
2. Suba los archivos del sistema.
3. Entre a Streamlit Community Cloud.
4. Seleccione **Create app**.
5. Elija el repositorio de GitHub.
6. Seleccione la rama, por ejemplo `main`.
7. En archivo principal escriba:

```text
app.py
```

8. En **Advanced settings**, pegue los secretos del paso 4.
9. Presione **Deploy**.

## 6. Credenciales iniciales del sistema

```text
Usuario: admin
Contraseña: admin123
```

En el login, el sistema genera un PATH temporal. Copie el código generado y péguelo en el campo correspondiente.

## 7. Flujo de trabajo recomendado

1. Login.
2. Catálogos base.
3. Administración de usuarios.
4. Registro de ingresos.
5. Registro de salidas.
6. Kardex consolidado.
7. Stock y alertas.
8. Reportes y exportación.

## 8. Recomendación de seguridad

- Use repositorio privado si el sistema manejará información sensible.
- No suba archivos JSON de credenciales a GitHub.
- No suba `.streamlit/secrets.toml`.
- Cambie la contraseña del usuario `admin` después del primer ingreso.
- En Google Sheets, comparta la hoja solo con la cuenta de servicio y personas autorizadas.


## Corrección de error 400: Request contains an invalid argument

Esta versión acepta en `GOOGLE_SHEET_ID` tanto el ID puro como la URL completa del Google Sheet.

Ejemplo válido con ID:

```toml
GOOGLE_SHEET_ID = "1AbCDefGhIjK123456789"
```

Ejemplo válido con URL completa:

```toml
GOOGLE_SHEET_ID = "https://docs.google.com/spreadsheets/d/1AbCDefGhIjK123456789/edit#gid=0"
```

Si el error persiste, revise:

1. Que el archivo sea realmente un Google Sheet, no un Excel cargado en Drive sin convertir.
2. Que la hoja esté compartida con el `client_email` de la cuenta de servicio como Editor.
3. Que `private_key` conserve `-----BEGIN PRIVATE KEY-----`, `-----END PRIVATE KEY-----` y los saltos de línea `\n`.
4. Que Google Sheets API esté habilitada en el proyecto de Google Cloud.

## Importante V9: evitar guardado local accidental

En versiones anteriores, si Google Sheets fallaba, el sistema podía usar Excel local como respaldo. En Streamlit Cloud eso puede confundir, porque el usuario registra datos pero no aparecen en Google Sheets.

En V9, use:

```toml
USE_GOOGLE_SHEETS = true
ALLOW_LOCAL_FALLBACK = false
FORMAT_GOOGLE_SHEETS_AS_TABLE = true
```

Con esta configuración, si Google Sheets falla, el sistema se detiene y muestra el error real.


## Si aparece APIError [429] Quota exceeded

Google Sheets limita las lecturas por minuto. La V11 ya minimiza las lecturas; si el error aparece durante pruebas intensas:

1. Espere 60 segundos.
2. Evite presionar varias veces los botones de recarga/reboot.
3. Verifique que esté usando esta V11 o superior.
4. Mantenga `ALLOW_LOCAL_FALLBACK = false` en producción para no guardar datos fuera de Google Sheets.
