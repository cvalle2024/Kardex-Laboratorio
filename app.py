from __future__ import annotations

import hashlib
import re
import secrets
import string
import time
import uuid
from io import BytesIO
import base64
from pathlib import Path
from typing import Dict, List, Tuple
from urllib.parse import quote


import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
APP_TITLE = "KARDEX PRO | Reactivos e Insumos"
APP_SUBTITLE = "Inventario inteligente con trazabilidad por lote, stock, vencimientos, alertas y reportes"
DATA_DIR = Path("data")
DB_FILE = DATA_DIR / "kardex_db.xlsx"
TODAY = pd.Timestamp.today().normalize()
DEFAULT_PATH_KEY = "DINAMICO"
DEFAULT_ADMIN_PASSWORD = "admin123"
PATH_TTL_MINUTES = 10
PATH_LENGTH = 8
SESSION_TIMEOUT_MINUTES = 15

KARDEX_CONSOLIDADO_COLUMNS: List[str] = [
    "estado", "producto_id", "producto", "marca", "lote", "unidad",
    "fecha_ingreso", "proveedor_ingreso", "orden_compra_ingreso", "fecha_elaboracion", "fecha_vencimiento",
    "entrada_total", "salida_total", "saldo_actual", "porcentaje_consumido",
    "numero_salidas", "fecha_ultima_salida", "ultimo_entregado_a", "ultimo_personal_entrega",
    "detalle_salidas", "observacion_ingreso", "dias_para_vencer", "stock_minimo",
]

SHEET_COLUMNS: Dict[str, List[str]] = {
    "Productos": [
        "producto_id", "codigo_producto", "nombre_producto", "categoria", "marca_default",
        "unidad_default", "stock_minimo", "dias_alerta_vencimiento", "activo", "observacion"
    ],
    "Proveedores": [
        "proveedor_id", "proveedor", "descripcion", "ruc", "representante", "telefono",
        "correo", "direccion", "activo"
    ],
    "Solicitantes": [
        "solicitante_id", "unidad_solicitante", "departamento", "municipio", "responsable",
        "telefono", "correo", "activo"
    ],
    "Personal": [
        "personal_id", "nombre", "cargo", "correo", "activo"
    ],
    "Usuarios": [
        "usuario_id", "usuario", "nombre", "rol", "password_hash", "path_verificacion", "activo", "fecha_creacion"
    ],
    "Movimientos": [
        "movimiento_id", "fecha", "tipo_movimiento", "producto_id", "producto", "marca", "lote",
        "proveedor", "orden_compra", "solicitante", "personal", "fecha_elaboracion", "fecha_vencimiento",
        "unidad", "cantidad", "costo_total", "observacion", "usuario_registro", "fecha_registro", "acta_entrega_id"
    ],
    # Hoja física en Google Sheets/Excel. Se calcula automáticamente desde Movimientos.
    # Esta hoja permite ver el stock actual en la misma base, sin depender solo de la vista de Streamlit.
    "Kardex_Consolidado": KARDEX_CONSOLIDADO_COLUMNS,
    "Config": ["clave", "valor"],
}

# Hojas que se leen al iniciar el sistema.
# Kardex_Consolidado es una hoja calculada/sincronizada desde Movimientos,
# por eso NO debe ser requisito de lectura al arrancar. Si la pestaña no existe
# todavía en Google Sheets, el sistema debe poder iniciar y crearla al sincronizar.
CORE_SHEETS: List[str] = [
    "Productos", "Proveedores", "Solicitantes", "Personal", "Usuarios", "Movimientos", "Config"
]
CALCULATED_SHEETS: List[str] = ["Kardex_Consolidado"]

# Formato visual de tablas en Google Sheets. Google Sheets no maneja las
# "Tablas" exactamente igual que Excel, pero sí permite dejar cada pestaña
# con encabezados destacados, filtros, fila congelada, bordes, formatos
# numéricos/fecha y anchos ordenados mediante la API.
TABLE_HEADER_COLOR = {"red": 0.043, "green": 0.047, "blue": 0.063}  # #0B0C10
TABLE_HEADER_TEXT = {"red": 1, "green": 1, "blue": 1}
TABLE_BORDER_COLOR = {"red": 0.82, "green": 0.86, "blue": 0.90}
TABLE_BODY_COLOR = {"red": 1, "green": 1, "blue": 1}
TABLE_ALT_COLOR = {"red": 0.965, "green": 0.976, "blue": 0.988}

TIPOS_MOVIMIENTO = ["Ingreso", "Salida", "Devolución", "Corrección entrada", "Corrección salida"]
# Se mantienen los valores legacy Ajuste entrada/salida para no romper bases creadas con versiones anteriores.
TIPOS_POSITIVOS = {"Ingreso", "Devolución", "Corrección entrada", "Ajuste entrada"}
TIPOS_NEGATIVOS = {"Salida", "Corrección salida", "Ajuste salida"}
CATEGORIAS_DEFAULT = ["Reactivo", "Insumo", "Equipo", "Material", "Papelería", "Otro"]
UNIDADES_DEFAULT = [
    "TABLETAS", "FRASCOS", "RESMAS", "C/U", "SET", "ROLLOS", "PRUEBAS", "KIT",
    "UNIDAD", "CAJAS", "CAJITAS", "BOLSAS", "AMPOLLAS", "TUBOS", "PLACAS", "ML", "L"
]
ROLES = ["Administrador", "Operador", "Consulta"]

# Orden de navegación recomendado según el flujo real de un Kardex.
PAGE_INICIO = "1️⃣ Inicio / Ruta del Kardex"
PAGE_CATALOGOS = "2️⃣ Catálogos base"
PAGE_ADMIN = "3️⃣ Administración"
PAGE_MOVIMIENTOS = "4️⃣ Registrar movimientos"
PAGE_KARDEX = "5️⃣ Kardex consolidado"
PAGE_STOCK = "6️⃣ Stock y alertas"
PAGE_DASHBOARD = "7️⃣ Dashboard ejecutivo"
PAGE_REPORTES = "8️⃣ Reportes y exportación"
PAGE_IMPORTAR = "9️⃣ Importar Kardex anterior"

NAV_PAGES = [
    PAGE_INICIO,
    PAGE_CATALOGOS,
    PAGE_ADMIN,
    PAGE_MOVIMIENTOS,
    PAGE_KARDEX,
    PAGE_STOCK,
    PAGE_DASHBOARD,
    PAGE_REPORTES,
    PAGE_IMPORTAR,
]


def hash_password(password: str) -> str:
    return hashlib.sha256(str(password).encode("utf-8")).hexdigest()


INITIAL_DATA: Dict[str, pd.DataFrame] = {
    "Productos": pd.DataFrame([
        {
            "producto_id": "PRD-0001",
            "codigo_producto": "TOXO-IGG-IGM",
            "nombre_producto": "Toxoplasmosis IgG/IgM",
            "categoria": "Reactivo",
            "marca_default": "Determine",
            "unidad_default": "FRASCOS",
            "stock_minimo": 5,
            "dias_alerta_vencimiento": 90,
            "activo": "Sí",
            "observacion": "Producto migrado como referencia desde el Kardex anterior."
        }
    ], columns=SHEET_COLUMNS["Productos"]),
    "Proveedores": pd.DataFrame([
        {
            "proveedor_id": "PROV-0001",
            "proveedor": "Abbott Guatemala",
            "descripcion": "Proveedor de referencia",
            "ruc": "",
            "representante": "N/A",
            "telefono": "",
            "correo": "",
            "direccion": "Guatemala",
            "activo": "Sí"
        }
    ], columns=SHEET_COLUMNS["Proveedores"]),
    "Solicitantes": pd.DataFrame([
        {
            "solicitante_id": "SOL-0001",
            "unidad_solicitante": "Hospital San Isidro",
            "departamento": "Colón",
            "municipio": "Tocoa",
            "responsable": "Daniela Navas",
            "telefono": "8798-9856",
            "correo": "",
            "activo": "Sí"
        }
    ], columns=SHEET_COLUMNS["Solicitantes"]),
    "Personal": pd.DataFrame([
        {
            "personal_id": "PER-0001",
            "nombre": "Vania Vallecillo",
            "cargo": "",
            "correo": "",
            "activo": "Sí"
        }
    ], columns=SHEET_COLUMNS["Personal"]),
    "Usuarios": pd.DataFrame([
        {
            "usuario_id": "USR-0001",
            "usuario": "admin",
            "nombre": "Administrador del sistema",
            "rol": "Administrador",
            "password_hash": hash_password(DEFAULT_ADMIN_PASSWORD),
            "path_verificacion": "DINAMICO",
            "activo": "Sí",
            "fecha_creacion": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    ], columns=SHEET_COLUMNS["Usuarios"]),
    "Movimientos": pd.DataFrame([
        {
            "movimiento_id": "MOV-000001",
            "fecha": "2025-09-02",
            "tipo_movimiento": "Ingreso",
            "producto_id": "PRD-0001",
            "producto": "Toxoplasmosis IgG/IgM",
            "marca": "Determine",
            "lote": "LOTE-REFERENCIA",
            "proveedor": "Abbott Guatemala",
            "solicitante": "",
            "personal": "",
            "fecha_elaboracion": "2020-05-25",
            "fecha_vencimiento": "2028-05-21",
            "unidad": "FRASCOS",
            "cantidad": 15,
            "costo_total": 1200,
            "observacion": "Registro base tomado como ejemplo del archivo anterior.",
            "usuario_registro": "Sistema",
            "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
        {
            "movimiento_id": "MOV-000002",
            "fecha": "2025-09-02",
            "tipo_movimiento": "Salida",
            "producto_id": "PRD-0001",
            "producto": "Toxoplasmosis IgG/IgM",
            "marca": "Determine",
            "lote": "LOTE-REFERENCIA",
            "proveedor": "",
            "solicitante": "Hospital San Isidro",
            "personal": "Vania Vallecillo",
            "fecha_elaboracion": "2020-05-25",
            "fecha_vencimiento": "2028-05-21",
            "unidad": "FRASCOS",
            "cantidad": 4,
            "costo_total": 0,
            "observacion": "Registro base tomado como ejemplo del archivo anterior.",
            "usuario_registro": "Sistema",
            "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
    ], columns=SHEET_COLUMNS["Movimientos"]),
    "Kardex_Consolidado": pd.DataFrame(columns=SHEET_COLUMNS["Kardex_Consolidado"]),
    "Config": pd.DataFrame([
        {"clave": "dias_alerta_global", "valor": "90"},
        {"clave": "moneda", "valor": "L"},
        {"clave": "institucion", "valor": "Proyecto VIHCA"},
        {"clave": "path_verificacion", "valor": "DINAMICO"},
        {"clave": "path_ttl_minutos", "valor": str(PATH_TTL_MINUTES)},
        {"clave": "version_sistema", "valor": "14.0"},
    ], columns=SHEET_COLUMNS["Config"]),
}

# ============================================================
# UTILIDADES
# ============================================================
def rerun() -> None:
    if hasattr(st, "rerun"):
        st.rerun()
    else:  # pragma: no cover
        st.experimental_rerun()


def set_flash(message: str, level: str = "success") -> None:
    """Guarda un mensaje temporal para mostrarlo después del rerun.

    Streamlit vuelve a ejecutar la app después de guardar; si usamos st.success()
    antes del rerun, el usuario casi no alcanza a ver la confirmación.
    Este helper conserva el mensaje y lo muestra una vez en la siguiente carga.
    """
    st.session_state["flash_message"] = {"level": level, "message": message}


def show_flash() -> None:
    flash = st.session_state.pop("flash_message", None)
    if not flash:
        return
    level = flash.get("level", "success")
    message = flash.get("message", "")
    if level == "success":
        st.success(message, icon="✅")
    elif level == "warning":
        st.warning(message, icon="⚠️")
    elif level == "error":
        st.error(message, icon="🚫")
    else:
        st.info(message, icon="ℹ️")


def bump_form_nonce(name: str) -> None:
    """Cambia la clave de un formulario para limpiar textboxes tras guardar."""
    st.session_state[name] = int(st.session_state.get(name, 0)) + 1


def clean_str(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_number(series, default: float = 0) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(default)


def to_date(series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True, format="mixed")


def format_date(value) -> str:
    if pd.isna(value) or value == "":
        return ""
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except Exception:
        return clean_str(value)


def active_mask(df: pd.DataFrame, col: str = "activo") -> pd.Series:
    if df.empty:
        return pd.Series([], dtype=bool, index=df.index)
    if col not in df.columns:
        return pd.Series([True] * len(df), index=df.index)
    return df[col].fillna("Sí").astype(str).str.lower().str.strip().isin(["sí", "si", "true", "1", "activo", ""])


def ensure_columns(df: pd.DataFrame, sheet: str) -> pd.DataFrame:
    columns = SHEET_COLUMNS[sheet]
    if df is None or df.empty:
        df = pd.DataFrame(columns=columns)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df[columns].copy()


def next_code(prefix: str, df: pd.DataFrame, id_col: str, width: int = 4) -> str:
    if df.empty or id_col not in df.columns:
        return f"{prefix}-{'1'.zfill(width)}"
    numbers = (
        df[id_col].astype(str).str.extract(r"(\d+)$")[0]
        .pipe(pd.to_numeric, errors="coerce")
        .dropna()
    )
    next_num = int(numbers.max()) + 1 if len(numbers) else 1
    return f"{prefix}-{str(next_num).zfill(width)}"


def safe_secret(key: str, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def nested_secret(section: str, key: str, default=None):
    """Lee valores anidados de st.secrets sin romper la app si no existen."""
    try:
        sec = st.secrets.get(section, None)
        if sec is None:
            return default
        return sec.get(key, default)
    except Exception:
        return default


def exception_detail(exc: Exception) -> str:
    """Devuelve un detalle útil de errores gspread/APIError para mostrar en pantalla."""
    parts = [str(exc)]
    response = getattr(exc, "response", None)
    if response is not None:
        try:
            parts.append(f"status={getattr(response, 'status_code', '')}")
            parts.append(response.text[:900])
        except Exception:
            pass
    return " | ".join([p for p in parts if p])


def as_bool(value, default: bool = False) -> bool:
    """Convierte valores de st.secrets a booleano de forma segura.

    En Streamlit Cloud es común pegar secretos como texto. Esta función evita que
    una cadena como "false" se interprete erróneamente como True por Python.
    """
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"true", "1", "yes", "sí", "si", "activo", "google_sheets"}:
        return True
    if text in {"false", "0", "no", "excel", "local", ""}:
        return False
    return default


def extract_google_sheet_id(value: str) -> str:
    """Acepta un ID puro o una URL completa de Google Sheets y devuelve solo el ID.

    Esto evita el error frecuente de Google Sheets API:
    APIError [400]: Request contains an invalid argument, que aparece cuando
    open_by_key recibe la URL completa en lugar del spreadsheetId.
    """
    text = clean_str(value)
    if not text:
        return ""

    # Formato URL: https://docs.google.com/spreadsheets/d/<ID>/edit#gid=0
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", text)
    if match:
        return match.group(1)

    # Formato compartido u otros enlaces que traen parámetros.
    text = text.split("?")[0].split("#")[0].strip()
    text = text.strip("/")
    return text


def column_letter(n: int) -> str:
    """Convierte 1 -> A, 27 -> AA para rangos de Google Sheets."""
    result = ""
    n = max(int(n), 1)
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def sheet_range_for(sheet: str) -> str:
    last_col = column_letter(len(SHEET_COLUMNS[sheet]))
    # Se usa rango de columnas completas para que Google devuelva solo las filas con datos.
    return f"'{sheet}'!A:{last_col}"


def values_to_dataframe(values: list, sheet: str) -> pd.DataFrame:
    """Convierte el resultado crudo de values_batch_get en DataFrame con columnas esperadas."""
    columns = SHEET_COLUMNS[sheet]
    if not values:
        return pd.DataFrame(columns=columns)
    rows = values[1:] if len(values) > 1 else []
    normalized_rows = [(list(row) + [""] * len(columns))[:len(columns)] for row in rows]
    return ensure_columns(pd.DataFrame(normalized_rows, columns=columns), sheet)


def mark_data_dirty() -> None:
    """Invalida caché de datos después de guardar para refrescar Google Sheets sin exceso de lecturas."""
    try:
        st.session_state["data_refresh_token"] = st.session_state.get("data_refresh_token", 0) + 1
    except Exception:
        pass
    try:
        st.cache_data.clear()
    except Exception:
        pass


def normalize_service_account_info(creds_info) -> dict:
    """Normaliza credenciales pegadas en Streamlit Secrets.

    Corrige casos comunes:
    - private_key con saltos de línea escapados como \n
    - valores tipo AttrDict de Streamlit
    """
    info = dict(creds_info)
    private_key = clean_str(info.get("private_key", ""))
    if "\\n" in private_key:
        private_key = private_key.replace("\\n", "\n")
    info["private_key"] = private_key
    return info


def diagnose_gsheets_error(exc: Exception) -> str:
    msg = str(exc)
    low = msg.lower()
    tips = []
    if "invalid argument" in low or "[400]" in low:
        tips.append("Verifique que GOOGLE_SHEET_ID sea el ID real de la hoja o una URL válida de Google Sheets. Esta versión ya extrae el ID si pega la URL completa.")
        tips.append("Confirme que la hoja exista y no sea un archivo Excel subido a Drive sin convertir a Google Sheets.")
        tips.append("Revise que los secretos TOML no tengan comillas faltantes, especialmente en private_key.")
    if "permission" in low or "forbidden" in low or "403" in low:
        tips.append("Comparta el Google Sheet con el client_email de la cuenta de servicio y permiso Editor.")
    if "not found" in low or "404" in low:
        tips.append("Revise que el ID/URL de la hoja sea correcto y que la cuenta de servicio tenga acceso.")
    if "private key" in low or "could not deserialize" in low:
        tips.append("Revise el campo private_key; debe conservar BEGIN PRIVATE KEY, END PRIVATE KEY y los saltos de línea \n.")
    if not tips:
        tips.append("Revise GOOGLE_SHEET_ID, gcp_service_account y que Google Sheets API esté habilitada en Google Cloud.")
    return " ".join(tips)


def normalize_for_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d")
        df[col] = df[col].replace({np.nan: "", pd.NaT: ""})
    return df.astype(str)


def get_config_value(data: Dict[str, pd.DataFrame], key: str, default: str = "") -> str:
    cfg = ensure_columns(data.get("Config", pd.DataFrame()), "Config")
    match = cfg[cfg["clave"].astype(str) == key]
    if match.empty:
        return default
    return clean_str(match.iloc[0]["valor"]) or default


def set_config_value(storage, data: Dict[str, pd.DataFrame], key: str, value: str) -> None:
    cfg = ensure_columns(data.get("Config", pd.DataFrame()), "Config")
    if (cfg["clave"].astype(str) == key).any():
        cfg.loc[cfg["clave"].astype(str) == key, "valor"] = value
    else:
        cfg = pd.concat([cfg, pd.DataFrame([{"clave": key, "valor": value}])], ignore_index=True)
    storage.save("Config", ensure_columns(cfg, "Config"))

# ============================================================
# CAPA DE ALMACENAMIENTO: EXCEL LOCAL O GOOGLE SHEETS
# ============================================================
class LocalExcelStorage:
    def __init__(self, path: Path):
        self.path = path
        DATA_DIR.mkdir(exist_ok=True)
        self.ensure_database()

    def ensure_database(self) -> None:
        if not self.path.exists():
            with pd.ExcelWriter(self.path, engine="openpyxl") as writer:
                for sheet, columns in SHEET_COLUMNS.items():
                    base_df = INITIAL_DATA.get(sheet, pd.DataFrame(columns=columns))
                    ensure_columns(base_df, sheet).to_excel(writer, index=False, sheet_name=sheet)
            return

        existing = pd.ExcelFile(self.path).sheet_names
        with pd.ExcelWriter(self.path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            for sheet, columns in SHEET_COLUMNS.items():
                if sheet not in existing:
                    base_df = INITIAL_DATA.get(sheet, pd.DataFrame(columns=columns))
                    ensure_columns(base_df, sheet).to_excel(writer, index=False, sheet_name=sheet)

    def load(self, sheet: str) -> pd.DataFrame:
        try:
            df = pd.read_excel(self.path, sheet_name=sheet, dtype=object)
        except Exception:
            df = pd.DataFrame(columns=SHEET_COLUMNS[sheet])
        return ensure_columns(df, sheet)

    def save(self, sheet: str, df: pd.DataFrame) -> None:
        df = ensure_columns(df, sheet)
        with pd.ExcelWriter(self.path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet)
        mark_data_dirty()

    def append_row(self, sheet: str, row: dict) -> None:
        self.append_rows(sheet, [row])

    def append_rows(self, sheet: str, rows: list[dict]) -> None:
        if not rows:
            return
        df = self.load(sheet)
        df = pd.concat([df, pd.DataFrame(rows)], ignore_index=True)
        self.save(sheet, df)
        mark_data_dirty()


class GoogleSheetsStorage:
    """Almacenamiento Google Sheets optimizado por API REST directa.

    Mejoras V13:
    - Ya no usa open_by_key() ni worksheet(), porque esas llamadas hacen lecturas de metadata.
    - Lee todas las pestañas con values:batchGet en una sola petición.
    - Guarda con endpoints de values update/append.
    - Incluye reintentos con espera si Google devuelve 429 por cuota.
    """
    def __init__(self):
        try:
            from google.oauth2.service_account import Credentials
            from google.auth.transport.requests import AuthorizedSession
        except ImportError as exc:
            raise RuntimeError("Falta instalar google-auth para usar Google Sheets.") from exc

        sheet_id_raw = safe_secret("GOOGLE_SHEET_ID", "") or nested_secret("google_sheets", "spreadsheet_id", "")
        sheet_id = extract_google_sheet_id(sheet_id_raw)
        creds_info = safe_secret("gcp_service_account", None) or safe_secret("google_service_account", None)
        if not sheet_id or not creds_info:
            raise RuntimeError(
                "No se encontró GOOGLE_SHEET_ID/gcp_service_account en Secrets. "
                "También se acepta [google_sheets].spreadsheet_id y [google_service_account] como respaldo."
            )

        creds_dict = normalize_service_account_info(creds_info)
        client_email = clean_str(creds_dict.get("client_email", ""))
        private_key = clean_str(creds_dict.get("private_key", ""))
        if not client_email or "@" not in client_email:
            raise RuntimeError("El campo client_email de gcp_service_account está vacío o no parece un correo válido.")
        if "BEGIN PRIVATE KEY" not in private_key or "END PRIVATE KEY" not in private_key:
            raise RuntimeError("El campo private_key no tiene el formato correcto. Debe incluir BEGIN PRIVATE KEY y END PRIVATE KEY.")

        self.sheet_id = sheet_id
        self.client_email = client_email
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        try:
            credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        except Exception as exc:
            raise RuntimeError(
                "No se pudieron cargar las credenciales de la cuenta de servicio. "
                f"Detalle: {exception_detail(exc)}"
            ) from exc
        self.session = AuthorizedSession(credentials)
        self.base_url = f"https://sheets.googleapis.com/v4/spreadsheets/{self.sheet_id}"
        self._sheet_ids: Dict[str, int] = {}

        # Prepara pestañas con operaciones de escritura. Esto evita la llamada de metadata
        # que estaba disparando el error 429 en open_by_key()/fetch_sheet_metadata.
        auto_create = as_bool(safe_secret("AUTO_CREATE_SHEETS", True), True)
        if auto_create:
            self.ensure_database_write_only()

    def _request(self, method: str, url: str, *, ok_empty: bool = False, **kwargs):
        """Ejecuta una llamada REST con reintentos para 429/5xx."""
        max_attempts = int(safe_secret("GSHEETS_MAX_RETRIES", 3) or 3)
        base_wait = int(safe_secret("GSHEETS_RETRY_SECONDS", 20) or 20)
        last_detail = ""
        for attempt in range(max_attempts + 1):
            try:
                response = self.session.request(method, url, timeout=40, **kwargs)
            except Exception as exc:
                last_detail = exception_detail(exc)
                if attempt >= max_attempts:
                    raise RuntimeError(f"No se pudo contactar Google Sheets. Detalle: {last_detail}") from exc
                time.sleep(min(base_wait * (attempt + 1), 60))
                continue

            if response.status_code in (429, 500, 502, 503, 504):
                last_detail = response.text[:1000]
                if attempt < max_attempts:
                    wait = min(base_wait * (attempt + 1), 60)
                    try:
                        st.warning(f"Google Sheets está limitando solicitudes temporalmente ({response.status_code}). Reintentando en {wait} segundos...")
                    except Exception:
                        pass
                    time.sleep(wait)
                    continue

            if not response.ok:
                detail = response.text[:1600]
                raise RuntimeError(f"APIError: [{response.status_code}]: {detail}")

            if ok_empty or not response.text:
                return {}
            try:
                return response.json()
            except Exception:
                return {}

        raise RuntimeError(
            "Google Sheets sigue devolviendo error de cuota o disponibilidad después de varios reintentos. "
            f"Último detalle: {last_detail}"
        )

    def _refresh_sheet_ids(self) -> None:
        """Lee una sola vez los sheetId necesarios para aplicar formato visual.

        Las operaciones de formato de Google Sheets requieren sheetId numérico.
        Se cachea dentro del objeto para no consumir lecturas de metadata en cada acción.
        """
        url = f"{self.base_url}?fields=sheets.properties(sheetId,title)"
        result = self._request("get", url)
        mapping = {}
        for item in result.get("sheets", []) if isinstance(result, dict) else []:
            props = item.get("properties", {})
            title = props.get("title")
            sheet_id = props.get("sheetId")
            if title and sheet_id is not None:
                mapping[str(title)] = int(sheet_id)
        if mapping:
            self._sheet_ids.update(mapping)

    def _get_sheet_id(self, sheet: str) -> int:
        if sheet not in self._sheet_ids:
            self._refresh_sheet_ids()
        if sheet not in self._sheet_ids:
            raise RuntimeError(f"No se encontró el sheetId de la pestaña '{sheet}' para aplicar formato.")
        return int(self._sheet_ids[sheet])

    def _format_requests_for_sheet(self, sheet: str, data_rows: int = 0) -> list:
        """Construye solicitudes batchUpdate para dejar una pestaña con formato tipo tabla."""
        sheet_id = self._get_sheet_id(sheet)
        columns = SHEET_COLUMNS[sheet]
        col_count = len(columns)
        # Incluye encabezado + filas con datos y deja un bloque visual disponible
        # para que nuevos registros con append sigan cayendo dentro del formato/filtro.
        total_rows = max(int(data_rows) + 1, 200)
        grid_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": total_rows, "startColumnIndex": 0, "endColumnIndex": col_count}
        header_range = {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": col_count}
        body_range = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": max(total_rows, 2), "startColumnIndex": 0, "endColumnIndex": col_count}

        border = {"style": "SOLID", "width": 1, "color": TABLE_BORDER_COLOR}
        requests = [
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            {
                "repeatCell": {
                    "range": header_range,
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": TABLE_HEADER_COLOR,
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP",
                            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": TABLE_HEADER_TEXT},
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)",
                }
            },
            {
                "repeatCell": {
                    "range": body_range,
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": TABLE_BODY_COLOR,
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP",
                            "textFormat": {"fontSize": 9},
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,verticalAlignment,wrapStrategy,textFormat)",
                }
            },
            {"setBasicFilter": {"filter": {"range": grid_range}}},
            {
                "updateBorders": {
                    "range": grid_range,
                    "top": border,
                    "bottom": border,
                    "left": border,
                    "right": border,
                    "innerHorizontal": border,
                    "innerVertical": border,
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                    "properties": {"pixelSize": 34},
                    "fields": "pixelSize",
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": col_count}
                }
            },
        ]

        # Formatos por tipo de columna.
        for idx, col in enumerate(columns):
            key = col.lower()
            col_range = {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": max(total_rows, 2), "startColumnIndex": idx, "endColumnIndex": idx + 1}
            if "fecha" in key:
                requests.append({
                    "repeatCell": {
                        "range": col_range,
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}},
                        "fields": "userEnteredFormat.numberFormat",
                    }
                })
            elif "porcentaje" in key:
                requests.append({
                    "repeatCell": {
                        "range": col_range,
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}},
                        "fields": "userEnteredFormat.numberFormat",
                    }
                })
            elif any(term in key for term in ["cantidad", "costo", "total", "saldo", "stock", "dias", "numero", "salida", "entrada"]):
                requests.append({
                    "repeatCell": {
                        "range": col_range,
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.##"}}},
                        "fields": "userEnteredFormat.numberFormat",
                    }
                })

            # Anchos máximos para columnas descriptivas que suelen crecer demasiado.
            if any(term in key for term in ["observacion", "detalle", "direccion"]):
                requests.append({
                    "updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": idx, "endIndex": idx + 1},
                        "properties": {"pixelSize": 320 if "detalle" in key else 260},
                        "fields": "pixelSize",
                    }
                })
            elif any(term in key for term in ["producto", "proveedor", "solicitante", "entregado"]):
                requests.append({
                    "updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": idx, "endIndex": idx + 1},
                        "properties": {"pixelSize": 210},
                        "fields": "pixelSize",
                    }
                })

        return requests

    def apply_table_format(self, sheet: str, data_rows: int = 0, strict: bool = False) -> None:
        """Aplica encabezado, filtros, fila congelada, bordes y formatos numéricos.

        No bloquea el guardado cuando falla, salvo que strict=True. Así se evita perder
        registros por un problema menor de presentación.
        """
        if not as_bool(safe_secret("FORMAT_GOOGLE_SHEETS_AS_TABLE", True), True):
            return
        try:
            requests = self._format_requests_for_sheet(sheet, data_rows=data_rows)
            self._request("post", f"{self.base_url}:batchUpdate", json={"requests": requests}, ok_empty=True)
        except Exception as exc:
            if strict:
                raise RuntimeError(f"No se pudo aplicar formato tabla en '{sheet}'. Detalle: {exception_detail(exc)}") from exc
            try:
                st.warning(f"Los datos se guardaron, pero no se pudo aplicar el formato tabla en '{sheet}'. Detalle: {exc}")
            except Exception:
                pass

    def apply_table_format_all(self, data: Dict[str, pd.DataFrame] | None = None, strict: bool = False) -> None:
        """Aplica formato tipo tabla a todas las pestañas de la base en una sola operación."""
        if not as_bool(safe_secret("FORMAT_GOOGLE_SHEETS_AS_TABLE", True), True):
            return
        try:
            all_requests = []
            for sheet in SHEET_COLUMNS.keys():
                rows = 0
                if data is not None and sheet in data:
                    rows = len(ensure_columns(data.get(sheet, pd.DataFrame()), sheet))
                all_requests.extend(self._format_requests_for_sheet(sheet, data_rows=rows))
            if all_requests:
                self._request("post", f"{self.base_url}:batchUpdate", json={"requests": all_requests}, ok_empty=True)
        except Exception as exc:
            if strict:
                raise RuntimeError(f"No se pudo aplicar formato tabla a Google Sheets. Detalle: {exception_detail(exc)}") from exc
            try:
                st.warning(f"No se pudo aplicar formato tabla a todas las pestañas. Detalle: {exc}")
            except Exception:
                pass

    def _values_url(self, sheet: str, suffix: str = "") -> str:
        rng = quote(sheet_range_for(sheet), safe="")
        return f"{self.base_url}/values/{rng}{suffix}"

    def _create_sheet_and_header(self, sheet: str) -> None:
        """Crea una pestaña y escribe encabezados sin depender de metadata.

        Es segura para ejecutar varias veces: si la pestaña ya existe, se ignora
        el error esperado de duplicado y se reescribe únicamente el encabezado.
        """
        if sheet not in SHEET_COLUMNS:
            raise RuntimeError(f"La hoja '{sheet}' no está definida en SHEET_COLUMNS.")

        body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet,
                            "gridProperties": {
                                "rowCount": 1000,
                                "columnCount": max(len(SHEET_COLUMNS[sheet]), 10),
                            },
                        }
                    }
                }
            ]
        }
        try:
            result = self._request("post", f"{self.base_url}:batchUpdate", json=body)
            replies = result.get("replies", []) if isinstance(result, dict) else []
            if replies:
                new_sheet_id = replies[0].get("addSheet", {}).get("properties", {}).get("sheetId")
                if new_sheet_id is not None:
                    self._sheet_ids[sheet] = int(new_sheet_id)
        except Exception as exc:
            msg = str(exc).lower()
            # Errores esperados si la pestaña ya existe. Se ignoran.
            if "already exists" not in msg and "already exist" not in msg and "duplicate" not in msg:
                if "429" in msg or "quota" in msg or "resource_exhausted" in msg:
                    raise RuntimeError(
                        "No se pudo crear/verificar estructura por cuota de Google Sheets. "
                        f"Detalle: {exception_detail(exc)}"
                    ) from exc
                if "unable to parse range" not in msg:
                    raise

        header_range = quote(f"'{sheet}'!A1:{column_letter(len(SHEET_COLUMNS[sheet]))}1", safe="")
        url = f"{self.base_url}/values/{header_range}?valueInputOption=USER_ENTERED"
        try:
            self._request("put", url, json={"values": [SHEET_COLUMNS[sheet]]}, ok_empty=True)
        except Exception as exc:
            raise RuntimeError(
                f"No se pudo escribir encabezados en la pestaña '{sheet}'. "
                f"Detalle API: {exception_detail(exc)}. Revisión: {diagnose_gsheets_error(exc)}"
            ) from exc

    def ensure_database_write_only(self) -> None:
        """Crea pestañas faltantes y escribe encabezados sin leer metadata.

        V15: la hoja Kardex_Consolidado se crea, pero ya no se lee al iniciar
        porque es una hoja calculada. Esto evita detener el sistema si esa pestaña
        física aún no existe o si Google Sheets limita lecturas de estructura.
        """
        for sheet in SHEET_COLUMNS.keys():
            self._create_sheet_and_header(sheet)

    def load_many(self, sheets: list[str]) -> Dict[str, pd.DataFrame]:
        """Carga varias pestañas con una sola lectura real a la API."""
        ranges = [sheet_range_for(sheet) for sheet in sheets]
        params = []
        for rng in ranges:
            params.append(("ranges", rng))
        params.append(("majorDimension", "ROWS"))
        try:
            result = self._request("get", f"{self.base_url}/values:batchGet", params=params)
        except Exception as exc:
            raise RuntimeError(
                "No se pudo leer Google Sheets por lote. "
                f"Detalle API: {exception_detail(exc)}. Revisión: {diagnose_gsheets_error(exc)}"
            ) from exc

        value_ranges = result.get("valueRanges", []) if isinstance(result, dict) else []
        out = {}
        for idx, sheet in enumerate(sheets):
            values = []
            if idx < len(value_ranges):
                values = value_ranges[idx].get("values", [])
            out[sheet] = values_to_dataframe(values, sheet)
        return out

    def load(self, sheet: str) -> pd.DataFrame:
        return self.load_many([sheet]).get(sheet, pd.DataFrame(columns=SHEET_COLUMNS[sheet]))

    def save(self, sheet: str, df: pd.DataFrame) -> None:
        df = normalize_for_sheet(ensure_columns(df, sheet))
        try:
            # Asegura que la pestaña exista antes de limpiar/actualizar. Esto es clave
            # para Kardex_Consolidado, porque es una hoja calculada que puede no existir
            # todavía en bases Google Sheets creadas con versiones anteriores.
            self._create_sheet_and_header(sheet)
            clear_url = self._values_url(sheet, ":clear")
            self._request("post", clear_url, json={}, ok_empty=True)
            values = [SHEET_COLUMNS[sheet]] + df.values.tolist()
            update_url = f"{self._values_url(sheet)}?valueInputOption=USER_ENTERED"
            self._request("put", update_url, json={"values": values}, ok_empty=True)
            self.apply_table_format(sheet, data_rows=len(df), strict=False)
            mark_data_dirty()
        except Exception as exc:
            raise RuntimeError(
                f"No se pudo guardar la pestaña '{sheet}'. Detalle API: {exception_detail(exc)}. "
                f"Revisión: {diagnose_gsheets_error(exc)}"
            ) from exc

    def append_row(self, sheet: str, row: dict) -> None:
        return self.append_rows(sheet, [row])

    def append_rows(self, sheet: str, rows: list[dict]) -> dict:
        if not rows:
            return {}
        df = pd.DataFrame(rows)
        df = normalize_for_sheet(ensure_columns(df, sheet))
        values = df.values.tolist()
        append_range = quote(f"'{sheet}'!A1", safe="")
        url = f"{self.base_url}/values/{append_range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS&includeValuesInResponse=false"
        try:
            self._create_sheet_and_header(sheet)
            response = self._request("post", url, json={"values": values})
            # Se actualiza el formato de la pestaña para que los nuevos registros queden dentro del filtro y la tabla visual.
            self.apply_table_format(sheet, data_rows=0, strict=False)
            mark_data_dirty()
            return response
        except Exception as exc:
            raise RuntimeError(
                f"No se pudieron agregar filas en '{sheet}'. Detalle API: {exception_detail(exc)}. "
                f"Revisión: {diagnose_gsheets_error(exc)}"
            ) from exc

    def test_write(self) -> dict:
        stamp = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            response = self.append_row("Config", {"clave": "ultima_prueba_escritura", "valor": stamp})
        except Exception as exc:
            raise RuntimeError(
                f"No se pudo ejecutar prueba de escritura. Detalle API: {exception_detail(exc)}. "
                f"Revisión: {diagnose_gsheets_error(exc)}"
            ) from exc
        return {"timestamp": stamp, "response": response}

    def info(self) -> dict:
        return {"sheet_id": getattr(self, "sheet_id", ""), "client_email": getattr(self, "client_email", "")}

@st.cache_resource(show_spinner=False)
def get_google_storage_cached():
    return GoogleSheetsStorage()


def get_storage():
    use_gsheets = as_bool(safe_secret("USE_GOOGLE_SHEETS", False), False)
    allow_local_fallback = as_bool(safe_secret("ALLOW_LOCAL_FALLBACK", False), False)
    if use_gsheets:
        try:
            return get_google_storage_cached(), "Google Sheets"
        except Exception as exc:
            st.error(
                "No se pudo conectar a Google Sheets. Para evitar que los registros se guarden solo en Excel local, "
                "el sistema se detuvo. "
                f"Detalle técnico: {exc}. "
                f"Revisión sugerida: {diagnose_gsheets_error(exc)}"
            )
            st.info(
                "Si desea permitir respaldo temporal en Excel local, agregue en Secrets: ALLOW_LOCAL_FALLBACK = true. "
                "Para producción, manténgalo en false."
            )
            if not allow_local_fallback:
                st.stop()
    return LocalExcelStorage(DB_FILE), "Excel local"

# ============================================================
# CÁLCULOS DE KARDEX
# ============================================================
@st.cache_data(ttl=60, show_spinner="Cargando base desde Google Sheets...")
def load_all_cached(_storage, mode: str, refresh_token: int) -> Dict[str, pd.DataFrame]:
    """Carga la base usando caché para evitar exceder cuota de Google Sheets.

    refresh_token cambia después de cada guardado; ttl evita releer en cada rerun.
    """
    # Solo se leen las hojas transaccionales/base.
    # Kardex_Consolidado se calcula desde Movimientos y Productos; no se lee como requisito
    # para que la app no se detenga si la pestaña física aún no existe.
    sheets = CORE_SHEETS
    if hasattr(_storage, "load_many"):
        data = _storage.load_many(sheets)
    else:
        data = {sheet: _storage.load(sheet) for sheet in sheets}
    data["Kardex_Consolidado"] = pd.DataFrame(columns=SHEET_COLUMNS["Kardex_Consolidado"])
    return data


def load_all(storage, mode: str = "") -> Dict[str, pd.DataFrame]:
    if mode == "Google Sheets":
        refresh_token = st.session_state.get("data_refresh_token", 0)
        return load_all_cached(storage, mode, refresh_token)
    data = {sheet: storage.load(sheet) for sheet in CORE_SHEETS}
    data["Kardex_Consolidado"] = pd.DataFrame(columns=SHEET_COLUMNS["Kardex_Consolidado"])
    return data


def calcular_stock(df_mov: pd.DataFrame, df_prod: pd.DataFrame) -> pd.DataFrame:
    df = ensure_columns(df_mov, "Movimientos")
    if df.empty:
        return pd.DataFrame(columns=[
            "producto_id", "producto", "marca", "lote", "fecha_vencimiento", "unidad",
            "ingreso_total", "salida_total", "stock_actual", "costo_total_ingresos",
            "stock_minimo", "dias_alerta_vencimiento", "dias_para_vencer", "estado"
        ])

    df = df.copy()
    df["cantidad_num"] = to_number(df["cantidad"])
    df["costo_total_num"] = to_number(df["costo_total"])
    df["tipo_movimiento"] = df["tipo_movimiento"].astype(str).str.strip()
    df["entrada"] = np.where(df["tipo_movimiento"].isin(TIPOS_POSITIVOS), df["cantidad_num"], 0)
    df["salida"] = np.where(df["tipo_movimiento"].isin(TIPOS_NEGATIVOS), df["cantidad_num"], 0)
    df["costo_ingreso"] = np.where(df["tipo_movimiento"].isin(TIPOS_POSITIVOS), df["costo_total_num"], 0)

    group_cols = ["producto_id", "producto", "marca", "lote", "fecha_vencimiento", "unidad"]
    stock = (
        df.groupby(group_cols, dropna=False, as_index=False)
        .agg(
            ingreso_total=("entrada", "sum"),
            salida_total=("salida", "sum"),
            costo_total_ingresos=("costo_ingreso", "sum"),
        )
    )
    stock["stock_actual"] = stock["ingreso_total"] - stock["salida_total"]

    prod = ensure_columns(df_prod, "Productos")[["producto_id", "stock_minimo", "dias_alerta_vencimiento"]].copy()
    prod["stock_minimo"] = to_number(prod["stock_minimo"])
    prod["dias_alerta_vencimiento"] = to_number(prod["dias_alerta_vencimiento"], 90)
    stock = stock.merge(prod, on="producto_id", how="left")
    stock["stock_minimo"] = to_number(stock["stock_minimo"])
    stock["dias_alerta_vencimiento"] = to_number(stock["dias_alerta_vencimiento"], 90)

    venc = to_date(stock["fecha_vencimiento"])
    stock["dias_para_vencer"] = (venc - TODAY).dt.days

    def estado(row) -> str:
        stock_actual = float(row.get("stock_actual", 0) or 0)
        dias = row.get("dias_para_vencer")
        minimo = float(row.get("stock_minimo", 0) or 0)
        alerta = float(row.get("dias_alerta_vencimiento", 90) or 90)
        if stock_actual <= 0:
            return "Sin stock"
        if pd.notna(dias) and dias < 0:
            return "Vencido"
        if pd.notna(dias) and dias <= alerta:
            return "Por vencer"
        if stock_actual <= minimo:
            return "Stock bajo"
        return "Disponible"

    stock["estado"] = stock.apply(estado, axis=1)
    stock = stock.sort_values(["estado", "producto", "fecha_vencimiento", "lote"], na_position="last")
    return stock


def kardex_consolidado_columns() -> List[str]:
    return KARDEX_CONSOLIDADO_COLUMNS.copy()


def first_non_empty(series) -> str:
    for value in series:
        text = clean_str(value)
        if text:
            return text
    return ""


def last_non_empty(series) -> str:
    for value in reversed(list(series)):
        text = clean_str(value)
        if text:
            return text
    return ""


def calcular_kardex_consolidado(df_mov: pd.DataFrame, df_prod: pd.DataFrame) -> pd.DataFrame:
    """Genera una vista de una fila por producto/lote.

    La hoja Movimientos se mantiene como bitácora transaccional. Esta vista consolida las
    entradas y salidas para que el usuario vea, en una misma fila, cuánto ingresó, cuánto
    salió, a quién se entregó por última vez y cuánto queda disponible.
    """
    columns = kardex_consolidado_columns()
    df = ensure_columns(df_mov, "Movimientos")
    if df.empty:
        return pd.DataFrame(columns=columns)

    df = df.copy()
    df["cantidad_num"] = to_number(df["cantidad"])
    df["fecha_dt"] = to_date(df["fecha"])
    df["fecha_registro_dt"] = to_date(df["fecha_registro"])
    df["tipo_movimiento"] = df["tipo_movimiento"].astype(str).str.strip()
    df["entrada"] = np.where(df["tipo_movimiento"].isin(TIPOS_POSITIVOS), df["cantidad_num"], 0)
    df["salida"] = np.where(df["tipo_movimiento"].isin(TIPOS_NEGATIVOS), df["cantidad_num"], 0)

    group_cols = ["producto_id", "producto", "marca", "lote", "fecha_vencimiento", "unidad"]
    base = (
        df.groupby(group_cols, dropna=False, as_index=False)
        .agg(
            entrada_total=("entrada", "sum"),
            salida_total=("salida", "sum"),
            movimientos_registrados=("movimiento_id", "count"),
        )
    )
    base["saldo_actual"] = base["entrada_total"] - base["salida_total"]
    base["porcentaje_consumido"] = np.where(
        base["entrada_total"] > 0,
        base["salida_total"] / base["entrada_total"],
        0,
    )

    positivos = df[df["tipo_movimiento"].isin(TIPOS_POSITIVOS)].copy()
    if not positivos.empty:
        positivos = positivos.sort_values(["fecha_dt", "fecha_registro_dt"], na_position="last")
        ingreso_info = (
            positivos.groupby(group_cols, dropna=False)
            .agg(
                fecha_ingreso=("fecha", first_non_empty),
                proveedor_ingreso=("proveedor", first_non_empty),
                orden_compra_ingreso=("orden_compra", first_non_empty),
                fecha_elaboracion=("fecha_elaboracion", first_non_empty),
                observacion_ingreso=("observacion", first_non_empty),
            )
            .reset_index()
        )
    else:
        ingreso_info = pd.DataFrame(columns=group_cols + ["fecha_ingreso", "proveedor_ingreso", "orden_compra_ingreso", "fecha_elaboracion", "observacion_ingreso"])

    negativos = df[df["tipo_movimiento"].isin(TIPOS_NEGATIVOS)].copy()
    if not negativos.empty:
        negativos = negativos.sort_values(["fecha_dt", "fecha_registro_dt"], na_position="last")
        salida_info = (
            negativos.groupby(group_cols, dropna=False)
            .agg(
                numero_salidas=("movimiento_id", "count"),
                fecha_ultima_salida=("fecha", last_non_empty),
                ultimo_entregado_a=("solicitante", last_non_empty),
                ultimo_personal_entrega=("personal", last_non_empty),
            )
            .reset_index()
        )

        def build_detalle_salidas(g: pd.DataFrame) -> str:
            partes = []
            g = g.sort_values(["fecha_dt", "fecha_registro_dt"], na_position="last")
            for _, r in g.iterrows():
                fecha_txt = format_date(r.get("fecha", ""))
                cantidad_txt = f"{float(r.get('cantidad_num', 0) or 0):,.0f}"
                unidad_txt = clean_str(r.get("unidad", ""))
                solicitante_txt = clean_str(r.get("solicitante", "")) or "Sin solicitante"
                personal_txt = clean_str(r.get("personal", ""))
                responsable_txt = f" | Entrega: {personal_txt}" if personal_txt else ""
                partes.append(f"{fecha_txt}: {cantidad_txt} {unidad_txt} a {solicitante_txt}{responsable_txt}")
            return "\n".join(partes)

        detalles_rows = []
        for key, group in negativos.groupby(group_cols, dropna=False):
            key_tuple = key if isinstance(key, tuple) else (key,)
            row_detalle = dict(zip(group_cols, key_tuple))
            row_detalle["detalle_salidas"] = build_detalle_salidas(group)
            detalles_rows.append(row_detalle)
        detalles = pd.DataFrame(detalles_rows, columns=group_cols + ["detalle_salidas"])
        salida_info = salida_info.merge(detalles, on=group_cols, how="left")
    else:
        salida_info = pd.DataFrame(columns=group_cols + ["numero_salidas", "fecha_ultima_salida", "ultimo_entregado_a", "ultimo_personal_entrega", "detalle_salidas"])

    kardex = base.merge(ingreso_info, on=group_cols, how="left").merge(salida_info, on=group_cols, how="left")

    prod = ensure_columns(df_prod, "Productos")[["producto_id", "stock_minimo", "dias_alerta_vencimiento"]].copy()
    prod["stock_minimo"] = to_number(prod["stock_minimo"])
    prod["dias_alerta_vencimiento"] = to_number(prod["dias_alerta_vencimiento"], 90)
    kardex = kardex.merge(prod, on="producto_id", how="left")
    kardex["stock_minimo"] = to_number(kardex["stock_minimo"])
    kardex["dias_alerta_vencimiento"] = to_number(kardex["dias_alerta_vencimiento"], 90)

    venc = to_date(kardex["fecha_vencimiento"])
    kardex["dias_para_vencer"] = (venc - TODAY).dt.days

    def estado_consolidado(row) -> str:
        saldo = float(row.get("saldo_actual", 0) or 0)
        dias = row.get("dias_para_vencer")
        minimo = float(row.get("stock_minimo", 0) or 0)
        alerta = float(row.get("dias_alerta_vencimiento", 90) or 90)
        if saldo <= 0:
            return "Consumido / sin stock"
        if pd.notna(dias) and dias < 0:
            return "Vencido"
        if pd.notna(dias) and dias <= alerta:
            return "Por vencer"
        if saldo <= minimo:
            return "Stock bajo"
        if float(row.get("salida_total", 0) or 0) > 0:
            return "Con salidas"
        return "Disponible sin salidas"

    kardex["estado"] = kardex.apply(estado_consolidado, axis=1)
    for col in ["fecha_ingreso", "proveedor_ingreso", "orden_compra_ingreso", "fecha_elaboracion", "observacion_ingreso", "fecha_ultima_salida", "ultimo_entregado_a", "ultimo_personal_entrega", "detalle_salidas"]:
        if col not in kardex.columns:
            kardex[col] = ""
        kardex[col] = kardex[col].fillna("")
    kardex["numero_salidas"] = to_number(kardex.get("numero_salidas", pd.Series(dtype=float))).astype(int)
    kardex = kardex.sort_values(["estado", "producto", "fecha_vencimiento", "lote"], na_position="last")
    return kardex[columns]


def sync_kardex_consolidado_sheet(storage, data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Actualiza la hoja física Kardex_Consolidado en Google Sheets/Excel.

    La tabla consolidada no se digita manualmente: se recalcula desde Movimientos y
    Productos para reflejar entrada total, salida acumulada y saldo actual por lote.
    """
    movimientos = ensure_columns(data.get("Movimientos", pd.DataFrame()), "Movimientos")
    productos = ensure_columns(data.get("Productos", pd.DataFrame()), "Productos")
    kardex = calcular_kardex_consolidado(movimientos, productos)
    storage.save("Kardex_Consolidado", ensure_columns(kardex, "Kardex_Consolidado"))
    return kardex


def resumen_kpis(stock: pd.DataFrame, movimientos: pd.DataFrame) -> Dict[str, int | float]:
    stock_pos = stock[stock["stock_actual"] > 0].copy() if not stock.empty else stock
    mov_mes = movimientos.copy()
    if not mov_mes.empty:
        mov_mes["fecha_dt"] = to_date(mov_mes["fecha"])
        mov_mes = mov_mes[mov_mes["fecha_dt"].dt.to_period("M") == TODAY.to_period("M")]
    return {
        "lotes_activos": int(len(stock_pos)),
        "productos_activos": int(stock_pos["producto"].nunique()) if not stock_pos.empty else 0,
        "stock_total": float(stock_pos["stock_actual"].sum()) if not stock_pos.empty else 0,
        "vencidos": int(((stock["estado"] == "Vencido") & (stock["stock_actual"] > 0)).sum()) if not stock.empty else 0,
        "por_vencer": int(((stock["estado"] == "Por vencer") & (stock["stock_actual"] > 0)).sum()) if not stock.empty else 0,
        "stock_bajo": int(((stock["estado"] == "Stock bajo") & (stock["stock_actual"] > 0)).sum()) if not stock.empty else 0,
        "movimientos_mes": int(len(mov_mes)),
    }

# ============================================================
# EXPORTACIONES / IMPORTACIÓN LEGACY
# ============================================================
def build_excel_report(data: Dict[str, pd.DataFrame], stock: pd.DataFrame, kardex: pd.DataFrame) -> bytes:
    output = BytesIO()
    movimientos = data["Movimientos"].copy()
    alertas_vencimiento = stock[(stock["estado"].isin(["Vencido", "Por vencer"])) & (stock["stock_actual"] > 0)].copy()
    stock_bajo = stock[(stock["estado"] == "Stock bajo") & (stock["stock_actual"] > 0)].copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kardex.to_excel(writer, index=False, sheet_name="Kardex_Consolidado")
        stock.to_excel(writer, index=False, sheet_name="Stock_Actual")
        movimientos.to_excel(writer, index=False, sheet_name="Movimientos")
        alertas_vencimiento.to_excel(writer, index=False, sheet_name="Alertas_Vencimiento")
        stock_bajo.to_excel(writer, index=False, sheet_name="Stock_Bajo")
        data["Productos"].to_excel(writer, index=False, sheet_name="Catalogo_Productos")
        data["Proveedores"].to_excel(writer, index=False, sheet_name="Catalogo_Proveedores")
        data["Solicitantes"].to_excel(writer, index=False, sheet_name="Catalogo_Solicitantes")
        data["Personal"].to_excel(writer, index=False, sheet_name="Catalogo_Personal")
    output.seek(0)
    return output.read()


def parse_legacy_kardex(uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(uploaded_file, sheet_name="MOVIMIENTO", header=3, dtype=object)
    raw = raw.dropna(how="all")
    raw.columns = [clean_str(c).upper() for c in raw.columns]
    rows = []
    products = []

    for _, r in raw.iterrows():
        producto = clean_str(r.get("MEDICAMENTO", ""))
        if not producto:
            continue
        producto_id = "PRD-" + str(abs(hash(producto)) % 999999).zfill(6)
        products.append({
            "producto_id": producto_id,
            "codigo_producto": producto[:20].upper().replace(" ", "-")[:25],
            "nombre_producto": producto,
            "categoria": "Reactivo/Insumo",
            "marca_default": clean_str(r.get("MARCA", "")),
            "unidad_default": clean_str(r.get("PRESENTACION", "")),
            "stock_minimo": 5,
            "dias_alerta_vencimiento": 90,
            "activo": "Sí",
            "observacion": "Producto importado desde Kardex anterior.",
        })
        base = {
            "fecha": format_date(r.get("FECHA", "")),
            "producto_id": producto_id,
            "producto": producto,
            "marca": clean_str(r.get("MARCA", "")),
            "lote": clean_str(r.get("LOTE", "")),
            "proveedor": clean_str(r.get("PROVEEDOR", "")),
            "solicitante": clean_str(r.get("NOMBRE DEL PACIENTE A QUIEN SE LE ENTREGA", "")),
            "personal": clean_str(r.get("PERSONAL QUE ENTREGA MEDICAMENTO", "")),
            "fecha_elaboracion": format_date(r.get("FECHA DE ELABORACION", "")),
            "fecha_vencimiento": format_date(r.get("FECHA DE VENCIMIENTO", "")),
            "unidad": clean_str(r.get("PRESENTACION", "")),
            "costo_total": r.get("$ COSTO TOTAL", ""),
            "observacion": clean_str(r.get("OBSERVACION", "")),
            "usuario_registro": "Importación legacy",
            "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        for tipo, col in [("Ingreso", "INGRESO"), ("Salida", "SALIDA"), ("Devolución", "DEVUELVE")]:
            qty = pd.to_numeric(r.get(col, 0), errors="coerce")
            if pd.notna(qty) and qty > 0:
                item = base.copy()
                item["movimiento_id"] = "MOV-" + uuid.uuid4().hex[:10].upper()
                item["tipo_movimiento"] = tipo
                item["cantidad"] = qty
                rows.append(item)

    mov = pd.DataFrame(rows, columns=SHEET_COLUMNS["Movimientos"])
    prod = pd.DataFrame(products, columns=SHEET_COLUMNS["Productos"]).drop_duplicates("nombre_producto")
    return mov, prod

# ============================================================
# UI / ESTILO
# ============================================================
def apply_theme() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="wide")
    st.markdown(
        """
        <style>
        :root{
            --bg:#050816; --panel:#0B1020; --panel2:#111827; --line:rgba(148,163,184,.22);
            --text:#E5E7EB; --muted:#94A3B8; --cyan:#38BDF8; --green:#22C55E; --orange:#F97316; --red:#EF4444;
            --purple:#8B5CF6; --blue:#2563EB;
        }
        .stApp{
            background: radial-gradient(circle at top left, rgba(56,189,248,.12), transparent 28%),
                        radial-gradient(circle at top right, rgba(139,92,246,.10), transparent 30%),
                        linear-gradient(180deg, #030712 0%, #0B1020 55%, #030712 100%);
        }
        .block-container{padding-top:1rem; padding-bottom:2rem; max-width:1480px;}
        section[data-testid="stSidebar"]{background:rgba(2,6,23,.96); border-right:1px solid rgba(148,163,184,.18);}
        .hero{
            background: radial-gradient(circle at top left, rgba(56,189,248,.28), transparent 32%),
                        linear-gradient(135deg, #070B18 0%, #0F172A 62%, #111827 100%);
            border:1px solid var(--line); border-radius:26px; padding:24px 26px; margin-bottom:18px;
            box-shadow:0 18px 50px rgba(0,0,0,.25);
        }
        .hero h1{font-size:34px; margin:0; color:white; letter-spacing:.3px;}
        .hero p{margin:7px 0 0 0; color:var(--muted); font-size:14px;}
        .section-title{font-size:24px; font-weight:900; color:#F8FAFC; margin:6px 0 4px 0;}
        .section-subtitle{font-size:13px; color:#94A3B8; margin-bottom:16px;}
        .form-card{
            border:1px solid rgba(148,163,184,.22); border-radius:22px; padding:18px 20px; margin:8px 0 18px 0;
            background:linear-gradient(180deg, rgba(15,23,42,.82), rgba(2,6,23,.70)); box-shadow:0 14px 35px rgba(0,0,0,.20);
        }
        .mini-card{
            border:1px solid rgba(148,163,184,.18); border-radius:18px; padding:14px 16px; margin-bottom:12px;
            background:rgba(15,23,42,.60);
        }
        .mini-title{font-weight:800; color:#F8FAFC; font-size:15px; margin-bottom:2px;}
        .mini-sub{color:#94A3B8; font-size:12px;}
        .kpi-card{
            border:1px solid var(--line); border-radius:20px; padding:17px 19px;
            background:linear-gradient(180deg, rgba(15,23,42,.96), rgba(2,6,23,.94));
            box-shadow:0 14px 35px rgba(0,0,0,.18); min-height:120px;
        }
        .kpi-label{color:#94A3B8; font-size:12px; text-transform:uppercase; letter-spacing:.08em;}
        .kpi-value{color:#F8FAFC; font-size:32px; font-weight:900; margin-top:8px;}
        .kpi-note{color:#CBD5E1; font-size:12px; margin-top:6px;}
        .pill{display:inline-block; padding:5px 10px; border-radius:999px; background:rgba(56,189,248,.12); color:#BAE6FD; border:1px solid rgba(56,189,248,.22); font-size:12px; margin-right:6px;}
        .alert-red{border-left:4px solid #EF4444; padding:12px 15px; background:rgba(239,68,68,.10); border-radius:14px; color:#FEE2E2;}
        .alert-orange{border-left:4px solid #F97316; padding:12px 15px; background:rgba(249,115,22,.10); border-radius:14px; color:#FFEDD5;}
        .alert-green{border-left:4px solid #22C55E; padding:12px 15px; background:rgba(34,197,94,.10); border-radius:14px; color:#DCFCE7;}
        .login-wrap{max-width:540px; margin:7vh auto 0 auto;}
        .login-card{border:1px solid rgba(148,163,184,.22); border-radius:28px; padding:28px; background:linear-gradient(180deg, rgba(15,23,42,.93), rgba(2,6,23,.88)); box-shadow:0 24px 70px rgba(0,0,0,.35);}
        .login-title{font-size:31px; font-weight:950; color:#F8FAFC; margin:0;}
        .login-sub{font-size:14px; color:#94A3B8; margin:6px 0 18px 0;}
        .field-note{font-size:12px; color:#94A3B8; margin-top:-8px; margin-bottom:10px;}
        div[data-testid="stMetricValue"]{font-size:28px;}
        .stButton>button{border-radius:12px; font-weight:700;}
        .stDownloadButton>button{border-radius:12px; font-weight:700;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(storage_mode: str, user_name: str = "") -> None:
    user_badge = f"<span class='pill'>Usuario: {user_name}</span>" if user_name else ""
    st.markdown(
        f"""
        <div class="hero">
            <h1>📦 {APP_TITLE}</h1>
            <p>{APP_SUBTITLE}</p>
            <div style="margin-top:12px"><span class='pill'>Base activa: {storage_mode}</span>{user_badge}<span class='pill'>Versión 17.0</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str = "") -> None:
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='section-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def card_start(title: str, subtitle: str = "") -> None:
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="mini-title">{title}</div>
            <div class="mini-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(label: str, value, note: str = "") -> None:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# AUTENTICACIÓN
# ============================================================
def render_login(storage, data: Dict[str, pd.DataFrame], mode: str) -> bool:
    if st.session_state.get("auth_ok"):
        return True

    usuarios = ensure_columns(data["Usuarios"], "Usuarios")
    dynamic_path = ensure_dynamic_login_path()
    created_at = pd.Timestamp(st.session_state.get("login_path_created_at", pd.Timestamp.now()))
    expires_at = created_at + pd.Timedelta(minutes=PATH_TTL_MINUTES)
    remaining_seconds = max(0, int((expires_at - pd.Timestamp.now()).total_seconds()))
    remaining_min = remaining_seconds // 60
    remaining_sec = remaining_seconds % 60

    st.markdown("<div class='login-wrap'><div class='login-card'>", unsafe_allow_html=True)
    st.markdown(
        f"<p class='login-title'>🔐 Acceso Kardex PRO</p>"
        f"<p class='login-sub'>{APP_SUBTITLE}<br>Base activa: <b>{mode}</b></p>",
        unsafe_allow_html=True,
    )

    st.markdown("#### Código PATH generado")
    st.code(dynamic_path, language="text")
    st.caption(f"Copie el código anterior y péguelo abajo. ⏱️ Vigencia aproximada: {remaining_min:02d}:{remaining_sec:02d} minutos. Si vence, genere uno nuevo.")
    if st.button("🔄 Generar nuevo PATH", use_container_width=True):
        ensure_dynamic_login_path(force_new=True)
        rerun()

    with st.form("login_form"):
        usuario = st.text_input("Usuario", placeholder="Ingrese su usuario")
        password = st.text_input("Contraseña", type="password", placeholder="Ingrese su contraseña")
        path_key = st.text_input(
            "Pegar PATH generado",
            placeholder="Copie el código generado arriba y péguelo aquí",
            help="Debe coincidir exactamente con el código PATH generado en esta pantalla.",
        )
        submitted = st.form_submit_button("Ingresar al sistema", use_container_width=True)

    st.markdown(
        "<div class='field-note'>Ingrese sus credenciales asignadas por el administrador. "
        "El PATH temporal se genera automáticamente en esta pantalla y se copia/pega para validar el acceso.</div>",
        unsafe_allow_html=True,
    )
    st.markdown("</div></div>", unsafe_allow_html=True)

    if not submitted:
        return False

    if not usuario or not password or not path_key:
        st.error("Debe ingresar usuario, contraseña y pegar el PATH generado.")
        return False

    match = usuarios[usuarios["usuario"].astype(str).str.lower().str.strip() == usuario.lower().strip()]
    if match.empty:
        st.error("Usuario no encontrado.")
        return False

    user_row = match.iloc[0]
    if clean_str(user_row.get("activo", "Sí")).lower() not in ["sí", "si", "true", "1", "activo", ""]:
        st.error("El usuario está inactivo.")
        return False

    saved_hash = clean_str(user_row.get("password_hash", ""))
    valid_password = saved_hash == hash_password(password)
    valid_path = path_key.strip().upper() == dynamic_path.upper()

    if valid_password and valid_path:
        st.session_state["auth_ok"] = True
        st.session_state["usuario"] = clean_str(user_row.get("usuario", usuario))
        st.session_state["nombre_usuario"] = clean_str(user_row.get("nombre", usuario))
        st.session_state["rol"] = clean_str(user_row.get("rol", "Operador"))
        clear_dynamic_login_path()
        set_flash("Acceso autorizado. Bienvenido al sistema Kardex PRO.")
        rerun()
    else:
        if not valid_password:
            st.error("Contraseña incorrecta.")
        elif not valid_path:
            st.error("El PATH no coincide con el código generado. Copie el código mostrado arriba y péguelo nuevamente.")
    return False

def logout() -> None:
    for key in ["auth_ok", "usuario", "nombre_usuario", "rol", "login_path_code", "login_path_created_at", "last_activity_ts"]:
        st.session_state.pop(key, None)
    rerun()


def require_admin() -> bool:
    if st.session_state.get("rol") != "Administrador":
        st.warning("Este módulo requiere rol de Administrador.")
        return False
    return True


def get_session_timeout_minutes() -> int:
    """Tiempo máximo de inactividad antes de cerrar sesión.

    Puede cambiarse en Streamlit Secrets con:
    SESSION_TIMEOUT_MINUTES = 15
    """
    raw = safe_secret("SESSION_TIMEOUT_MINUTES", SESSION_TIMEOUT_MINUTES)
    try:
        value = int(float(raw))
    except Exception:
        value = SESSION_TIMEOUT_MINUTES
    return max(1, value)


def clear_auth_state() -> None:
    """Limpia únicamente el estado de autenticación de la sesión actual."""
    for key in [
        "auth_ok", "usuario", "nombre_usuario", "rol",
        "login_path_code", "login_path_created_at", "last_activity_ts"
    ]:
        st.session_state.pop(key, None)


def enforce_inactivity_timeout() -> bool:
    """Cierra sesión si la sesión autenticada superó el tiempo de inactividad."""
    if not st.session_state.get("auth_ok"):
        return False

    now = time.time()
    timeout_seconds = get_session_timeout_minutes() * 60
    last_activity = st.session_state.get("last_activity_ts")

    if last_activity is not None:
        try:
            elapsed = now - float(last_activity)
        except Exception:
            elapsed = 0
        if elapsed > timeout_seconds:
            clear_auth_state()
            set_flash(
                f"Sesión cerrada por inactividad después de {get_session_timeout_minutes()} minutos. Ingrese nuevamente.",
                "warning",
            )
            rerun()
            return False

    st.session_state["last_activity_ts"] = now
    return True


def inject_inactivity_watcher() -> None:
    """Recarga la app cuando el navegador queda sin actividad.

    Streamlit solo ejecuta Python cuando hay una interacción o recarga. Este pequeño
    script detecta inactividad del lado del navegador y fuerza una recarga; al recargar,
    Python valida el tiempo transcurrido y cierra la sesión si corresponde.
    """
    timeout_ms = get_session_timeout_minutes() * 60 * 1000
    components.html(
        f"""
        <script>
        (function() {{
            const timeoutMs = {timeout_ms};
            const parentWindow = window.parent;
            const timerKey = "__kardexInactivityTimer";
            const resetTimer = function() {{
                if (parentWindow[timerKey]) {{
                    clearTimeout(parentWindow[timerKey]);
                }}
                parentWindow[timerKey] = setTimeout(function() {{
                    parentWindow.location.reload();
                }}, timeoutMs + 1000);
            }};
            const events = ["mousemove", "mousedown", "keydown", "scroll", "touchstart", "click"];
            events.forEach(function(eventName) {{
                parentWindow.addEventListener(eventName, resetTimer, {{passive: true}});
            }});
            resetTimer();
        }})();
        </script>
        """,
        height=0,
        width=0,
    )


def allowed_nav_pages_for_role(role: str) -> List[str]:
    """Devuelve los módulos visibles según el rol del usuario."""
    pages = list(NAV_PAGES)
    if clean_str(role) != "Administrador":
        pages = [p for p in pages if p != PAGE_ADMIN]
    return pages


def generate_dynamic_path_code(length: int = PATH_LENGTH) -> str:
    """Genera un PATH temporal fácil de leer y suficientemente aleatorio para el login."""
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # evita I, O, 0 y 1
    token = "".join(secrets.choice(alphabet) for _ in range(length))
    return f"KDX-{token[:4]}-{token[4:]}"


def ensure_dynamic_login_path(force_new: bool = False) -> str:
    """Mantiene un PATH temporal en la sesión del navegador y lo renueva al vencer."""
    now = pd.Timestamp.now()
    created = st.session_state.get("login_path_created_at")
    expired = True
    if created is not None:
        try:
            expired = (now - pd.Timestamp(created)).total_seconds() > PATH_TTL_MINUTES * 60
        except Exception:
            expired = True

    if force_new or "login_path_code" not in st.session_state or expired:
        st.session_state["login_path_code"] = generate_dynamic_path_code()
        st.session_state["login_path_created_at"] = now.isoformat()
    return st.session_state["login_path_code"]


def clear_dynamic_login_path() -> None:
    st.session_state.pop("login_path_code", None)
    st.session_state.pop("login_path_created_at", None)

# ============================================================
# PÁGINAS
# ============================================================
def nav_go(page_name: str) -> None:
    st.session_state["page"] = page_name
    rerun()


def page_inicio_operativo(data: Dict[str, pd.DataFrame], stock: pd.DataFrame, kardex: pd.DataFrame, mode: str) -> None:
    section_header(
        "🧭 Ruta de trabajo del Kardex",
        "Guía de navegación ordenada según las acciones reales: configurar catálogos, administrar accesos, registrar movimientos, revisar stock y exportar reportes."
    )

    productos = ensure_columns(data["Productos"], "Productos")
    proveedores = ensure_columns(data["Proveedores"], "Proveedores")
    solicitantes = ensure_columns(data["Solicitantes"], "Solicitantes")
    personal_df = ensure_columns(data["Personal"], "Personal")
    usuarios = ensure_columns(data["Usuarios"], "Usuarios")
    movimientos = ensure_columns(data["Movimientos"], "Movimientos")

    productos_activos = int(active_mask(productos).sum()) if not productos.empty else 0
    proveedores_activos = int(active_mask(proveedores).sum()) if not proveedores.empty else 0
    solicitantes_activos = int(active_mask(solicitantes).sum()) if not solicitantes.empty else 0
    personal_activo = int(active_mask(personal_df).sum()) if not personal_df.empty else 0
    usuarios_activos = int(active_mask(usuarios).sum()) if not usuarios.empty else 0
    movimientos_total = int(len(movimientos))

    # Siguiente acción sugerida para que el usuario no se pierda.
    if productos_activos == 0 or proveedores_activos == 0 or solicitantes_activos == 0 or personal_activo == 0:
        siguiente = "Complete primero los catálogos base antes de registrar ingresos o salidas."
        siguiente_tipo = "alert-orange"
        boton_texto = "Ir a catálogos base"
        boton_page = PAGE_CATALOGOS
    elif st.session_state.get("rol") == "Administrador" and usuarios_activos <= 1:
        siguiente = "Ya hay catálogos mínimos. Ahora puede crear usuarios operativos o continuar con movimientos."
        siguiente_tipo = "alert-green"
        boton_texto = "Ir a administración"
        boton_page = PAGE_ADMIN
    elif movimientos_total == 0:
        siguiente = "Los catálogos están listos. El siguiente paso operativo es registrar ingresos iniciales."
        siguiente_tipo = "alert-green"
        boton_texto = "Registrar movimiento"
        boton_page = PAGE_MOVIMIENTOS
    else:
        siguiente = "El sistema ya tiene movimientos. Revise Kardex consolidado, stock, alertas y reportes."
        siguiente_tipo = "alert-green"
        boton_texto = "Ver Kardex consolidado"
        boton_page = PAGE_KARDEX

    st.markdown(f"<div class='{siguiente_tipo}'><b>Siguiente acción sugerida:</b> {siguiente}</div>", unsafe_allow_html=True)
    st.write("")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi_card("Productos", productos_activos, "Catálogo base")
    with c2: kpi_card("Proveedores", proveedores_activos, "Para ingresos")
    with c3: kpi_card("Solicitantes", solicitantes_activos, "Para salidas")
    with c4: kpi_card("Personal", personal_activo, "Recibe / entrega")
    with c5: kpi_card("Movimientos", movimientos_total, "Bitácora")

    st.write("")
    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    st.markdown("#### Flujo recomendado de uso")
    if st.session_state.get("rol") == "Administrador":
        flujo_html = """
        <div style="display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px;">
            <div class="mini-card"><div class="mini-title">1. Acceso seguro</div><div class="mini-sub">El usuario se loguea con usuario, contraseña y PATH temporal generado automáticamente.</div></div>
            <div class="mini-card"><div class="mini-title">2. Catálogos base</div><div class="mini-sub">Se registran productos, marcas, unidades, proveedores, solicitantes y personal.</div></div>
            <div class="mini-card"><div class="mini-title">3. Administración</div><div class="mini-sub">Solo administrador: usuarios, roles, seguridad PATH y diagnóstico de base.</div></div>
            <div class="mini-card"><div class="mini-title">4. Operación diaria</div><div class="mini-sub">Ingresos, salidas, devoluciones y ajustes. El formulario toma datos del catálogo.</div></div>
            <div class="mini-card"><div class="mini-title">5. Control Kardex</div><div class="mini-sub">Una fila por producto/lote con entrada, salida acumulada, saldo y último destino.</div></div>
            <div class="mini-card"><div class="mini-title">6. Reportes</div><div class="mini-sub">Stock, alertas, movimientos, Kardex consolidado y exportación a Excel.</div></div>
        </div>
        """
    else:
        flujo_html = """
        <div style="display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px;">
            <div class="mini-card"><div class="mini-title">1. Acceso seguro</div><div class="mini-sub">El usuario se loguea con sus credenciales y PATH temporal generado automáticamente.</div></div>
            <div class="mini-card"><div class="mini-title">2. Catálogos base</div><div class="mini-sub">Consulte o complete los catálogos permitidos según su rol operativo.</div></div>
            <div class="mini-card"><div class="mini-title">3. Operación diaria</div><div class="mini-sub">Registre ingresos, salidas, devoluciones y ajustes. El formulario toma datos del catálogo.</div></div>
            <div class="mini-card"><div class="mini-title">4. Control Kardex</div><div class="mini-sub">Revise existencia por producto/lote, salida acumulada, saldo y último destino.</div></div>
            <div class="mini-card"><div class="mini-title">5. Stock y alertas</div><div class="mini-sub">Consulte vencimientos, stock bajo y existencias disponibles.</div></div>
            <div class="mini-card"><div class="mini-title">6. Reportes</div><div class="mini-sub">Genere reportes operativos y exportaciones autorizadas.</div></div>
        </div>
        """
    st.markdown(flujo_html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    st.markdown("#### Accesos rápidos según la ruta")
    if st.session_state.get("rol") == "Administrador":
        a, b, c, d = st.columns(4)
        if a.button("📚 Catálogos base", use_container_width=True):
            nav_go(PAGE_CATALOGOS)
        if b.button("🔐 Administración", use_container_width=True):
            nav_go(PAGE_ADMIN)
        if c.button("🔁 Registrar movimientos", use_container_width=True):
            nav_go(PAGE_MOVIMIENTOS)
        if d.button("📊 Reportes", use_container_width=True):
            nav_go(PAGE_REPORTES)
    else:
        a, b, c = st.columns(3)
        if a.button("📚 Catálogos base", use_container_width=True):
            nav_go(PAGE_CATALOGOS)
        if b.button("🔁 Registrar movimientos", use_container_width=True):
            nav_go(PAGE_MOVIMIENTOS)
        if c.button("📊 Reportes", use_container_width=True):
            nav_go(PAGE_REPORTES)
    st.write("")
    if st.button(f"➡️ {boton_texto}", use_container_width=True):
        nav_go(boton_page)
    st.markdown("</div>", unsafe_allow_html=True)

    card_start("Regla importante del sistema", "Para evitar errores, los movimientos dependen de los catálogos.")
    st.markdown(
        """
        <div class='field-note' style='font-size:13px; margin-top:4px;'>
        Primero se crea o verifica el producto en catálogo. Después, al registrar un ingreso o salida, el sistema toma automáticamente nombre del producto, marca predeterminada, unidad, stock mínimo y días de alerta.
        Las salidas se hacen contra lotes existentes para mantener trazabilidad real por vencimiento y saldo disponible.
        </div>
        """,
        unsafe_allow_html=True,
    )


def page_dashboard(data: Dict[str, pd.DataFrame], stock: pd.DataFrame) -> None:
    section_header("🏠 Dashboard ejecutivo", "Resumen general del inventario, alertas y movimientos recientes.")
    kpis = resumen_kpis(stock, data["Movimientos"])
    cols = st.columns(4)
    with cols[0]: kpi_card("Stock total", f"{kpis['stock_total']:,.0f}", "Unidades disponibles en lotes activos")
    with cols[1]: kpi_card("Productos activos", kpis["productos_activos"], "Productos con existencia")
    with cols[2]: kpi_card("Alertas", kpis["por_vencer"] + kpis["vencidos"] + kpis["stock_bajo"], "Vencidos, por vencer o stock bajo")
    with cols[3]: kpi_card("Movimientos del mes", kpis["movimientos_mes"], "Entradas, salidas y ajustes")

    st.write("")
    c1, c2 = st.columns([1.35, 1])
    stock_pos = stock[stock["stock_actual"] > 0].copy() if not stock.empty else stock
    with c1:
        card_start("Stock por producto", "Top de productos con mayor existencia disponible.")
        if stock_pos.empty:
            st.info("No hay stock activo registrado todavía.")
        else:
            plot_df = stock_pos.groupby("producto", as_index=False)["stock_actual"].sum().sort_values("stock_actual", ascending=False).head(15)
            fig = px.bar(plot_df, x="stock_actual", y="producto", orientation="h", text="stock_actual")
            fig.update_layout(height=430, margin=dict(l=10, r=10, t=20, b=10), yaxis={"categoryorder": "total ascending"})
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        card_start("Estado del inventario", "Distribución de stock por estado operativo.")
        if stock.empty:
            st.info("Sin datos de inventario.")
        else:
            estado_df = stock_pos.groupby("estado", as_index=False)["stock_actual"].sum()
            fig = px.pie(estado_df, values="stock_actual", names="estado", hole=.58)
            fig.update_layout(height=430, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)

    card_start("Alertas críticas", "Lotes vencidos, próximos a vencer o con stock bajo.")
    alertas = stock[(stock["stock_actual"] > 0) & (stock["estado"].isin(["Vencido", "Por vencer", "Stock bajo"]))].copy()
    if alertas.empty:
        st.markdown('<div class="alert-green">✅ No hay alertas críticas activas.</div>', unsafe_allow_html=True)
    else:
        st.dataframe(alertas[["estado", "producto", "marca", "lote", "fecha_vencimiento", "dias_para_vencer", "stock_actual", "stock_minimo", "unidad"]], use_container_width=True, hide_index=True)


# ============================================================
# MÓDULO OPERATIVO DE MOVIMIENTOS V18
# ============================================================
def spanish_month_name(month: int) -> str:
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
        7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    return meses.get(int(month), "")


def fecha_larga_es(value) -> str:
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        dt = TODAY
    return f"{dt.day} de {spanish_month_name(dt.month)} de {dt.year}"


def next_movement_ids(df: pd.DataFrame, n: int) -> list[str]:
    first = next_code("MOV", df, "movimiento_id", 6)
    try:
        base = int(str(first).split("-")[-1])
    except Exception:
        base = int(time.time()) % 900000
    return [f"MOV-{base + i:06d}" for i in range(n)]


def get_first_match(df: pd.DataFrame, col: str, value: str) -> dict:
    if df is None or df.empty or col not in df.columns:
        return {}
    m = df[df[col].astype(str) == str(value)]
    if m.empty:
        return {}
    return {k: clean_str(v) for k, v in m.iloc[0].to_dict().items()}


def build_acta_entrega_pdf(
    salida_rows: list[dict],
    solicitante_info: dict,
    personal_info: dict,
    fecha_entrega,
    recibe_nombre: str = "",
    recibe_cargo: str = "",
    ciudad: str = "Tegucigalpa",
    observacion: str = "",
) -> bytes:
    """Genera acta de entrega en PDF basada en el diseño del acta de referencia.

    La salida se agrupa en una sola acta por sitio/solicitante y lista todos los
    productos del carrito de salida. Cada producto ya queda registrado de forma
    individual en la hoja Movimientos.
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    except Exception as exc:
        raise RuntimeError("Falta instalar reportlab. Agregue reportlab en requirements.txt.") from exc

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.85 * inch,
        leftMargin=0.85 * inch,
        topMargin=0.58 * inch,
        bottomMargin=0.60 * inch,
    )
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("ActaNormal", parent=styles["Normal"], fontName="Helvetica", fontSize=10.5, leading=14, alignment=TA_LEFT)
    normal_center = ParagraphStyle("ActaCenter", parent=normal, alignment=TA_CENTER)
    normal_right = ParagraphStyle("ActaRight", parent=normal, alignment=TA_RIGHT)
    bold = ParagraphStyle("ActaBold", parent=normal, fontName="Helvetica-Bold")
    small = ParagraphStyle("ActaSmall", parent=normal, fontSize=8.7, leading=11)
    small_center = ParagraphStyle("ActaSmallCenter", parent=small, alignment=TA_CENTER)
    note_style = ParagraphStyle("ActaNote", parent=normal, fontSize=8.8, leading=12, italic=True)

    sitio = clean_str(solicitante_info.get("unidad_solicitante", "")) or clean_str(salida_rows[0].get("solicitante", "")) or "Sitio receptor"
    responsable_catalogo = clean_str(solicitante_info.get("responsable", ""))
    recibe_nombre = clean_str(recibe_nombre) or responsable_catalogo or "Responsable del sitio"
    recibe_cargo = clean_str(recibe_cargo) or "Responsable / receptor"
    entrega_nombre = clean_str(personal_info.get("nombre", "")) or clean_str(salida_rows[0].get("personal", "")) or "Personal que entrega"
    entrega_cargo = clean_str(personal_info.get("cargo", "")) or "Personal asignado"

    story = []
    # Logo simplificado en texto para evitar depender de archivos externos en Streamlit Cloud.
    story.append(Paragraph("<b><font color='#0B4F8A' size='25'>VIHCA</font></b> <font color='#B91C1C' size='24'>♦</font>", normal_center))
    story.append(Paragraph("<font color='#B91C1C' size='7'><b>PROGRAMA REGIONAL DE VIH</b></font>", normal_center))
    story.append(Spacer(1, 0.18 * inch))
    story.append(Paragraph(f"{ciudad} {fecha_larga_es(fecha_entrega)}", normal_right))
    story.append(Spacer(1, 0.18 * inch))

    story.append(Paragraph(f"<b>{recibe_nombre}</b>", bold))
    story.append(Paragraph(sitio, normal))
    story.append(Spacer(1, 0.20 * inch))

    story.append(Paragraph(
        "Reciba un cordial saludo de parte del Programa Regional Centroamericano de VIH Asociado al "
        "Centro de Estudios en Salud de la Universidad del Valle de Guatemala.", normal
    ))
    story.append(Spacer(1, 0.12 * inch))
    story.append(Paragraph(
        f"El motivo de la presente es para hacer de su conocimiento que como parte del apoyo que el "
        f"Programa de VIH brinda en Honduras, se hace entrega formal de los siguientes reactivos e insumos "
        f"para el procesamiento y uso en {sitio}.", normal
    ))
    story.append(Spacer(1, 0.24 * inch))

    data_table = [[
        Paragraph("<b>DESCRIPCIÓN</b>", normal_center),
        Paragraph("<b>PRESENTACIÓN / CANTIDAD</b>", normal_center),
        Paragraph("<b>FECHA DE<br/>VENCIMIENTO</b>", normal_center),
        Paragraph("<b>LOTE</b>", normal_center),
    ]]
    for r in salida_rows:
        desc = clean_str(r.get("producto", ""))
        marca = clean_str(r.get("marca", ""))
        if marca:
            desc = f"{desc}<br/><font size='8'>Marca: {marca}</font>"
        cantidad = float(r.get("cantidad", 0) or 0)
        cantidad_txt = f"{cantidad:,.0f}" if cantidad.is_integer() else f"{cantidad:,.2f}"
        unidad = clean_str(r.get("unidad", ""))
        data_table.append([
            Paragraph(desc, normal_center),
            Paragraph(f"{cantidad_txt} {unidad}", normal_center),
            Paragraph(format_date(r.get("fecha_vencimiento", "")), normal_center),
            Paragraph(clean_str(r.get("lote", "")), normal_center),
        ])

    table = Table(data_table, colWidths=[2.30 * inch, 1.85 * inch, 1.55 * inch, 1.15 * inch], hAlign="CENTER")
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1.1, colors.black),
        ("BOX", (0, 0), (-1, -1), 1.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(table)
    story.append(Spacer(1, 0.40 * inch))

    firma_table = Table([
        ["____________________________", "____________________________"],
        [Paragraph("<b>Entregué conforme</b>", small_center), Paragraph("<b>Recibí conforme</b>", small_center)],
        [Paragraph(f"<b>{entrega_nombre}</b>", small_center), Paragraph(f"<b>{recibe_nombre}</b>", small_center)],
        [Paragraph(entrega_cargo, small_center), Paragraph(recibe_cargo, small_center)],
        [Paragraph("Programa Regional Centroamericano de VIH<br/>Asociado al Centro de Estudios en Salud, de la<br/>Universidad del Valle de Guatemala", small_center), Paragraph(sitio, small_center)],
    ], colWidths=[3.25 * inch, 3.25 * inch])
    firma_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(firma_table)
    story.append(Spacer(1, 0.35 * inch))

    note = (
        "<b>Nota:</b> se hace constar que, para la validación de este proceso y la posterior gestión del retorno "
        "del acta de entrega, es requisito indispensable que el documento cuente con la <b>firma y el sello oficial "
        "del laboratorio receptor</b>, para fines de inventario y control de calidad. (favor entregarla al personal "
        "de VIHCA asignado, después de su recepción y firma)"
    )
    if clean_str(observacion):
        note += f"<br/><br/><b>Observación:</b> {clean_str(observacion)}"
    story.append(Paragraph(note, note_style))

    doc.build(story)
    return buffer.getvalue()


def page_movimiento(storage, data: Dict[str, pd.DataFrame], stock: pd.DataFrame) -> None:
    section_header(
        "➕ Registrar movimientos",
        "Ingreso, salida por carrito, devolución desde salidas históricas y correcciones de inventario con formularios guiados."
    )

    productos = ensure_columns(data["Productos"], "Productos")
    proveedores = ensure_columns(data["Proveedores"], "Proveedores")
    solicitantes = ensure_columns(data["Solicitantes"], "Solicitantes")
    personal_df = ensure_columns(data["Personal"], "Personal")
    movimientos = ensure_columns(data["Movimientos"], "Movimientos")

    productos = productos[active_mask(productos)].copy()
    proveedores = proveedores[active_mask(proveedores)].copy()
    solicitantes = solicitantes[active_mask(solicitantes)].copy()
    personal_df = personal_df[active_mask(personal_df)].copy()

    if productos.empty:
        st.warning("Primero registre al menos un producto en el catálogo. El formulario tomará de ahí la marca, unidad, categoría, stock mínimo y días de alerta.")
        return

    def _catalog_row_by_id(producto_id_value: str) -> pd.Series:
        match = productos[productos["producto_id"].astype(str) == str(producto_id_value)]
        if not match.empty:
            return match.iloc[0]
        return pd.Series({col: "" for col in SHEET_COLUMNS["Productos"]})

    def _product_card(prod_row: pd.Series, titulo: str = "Ficha del producto tomada del catálogo") -> None:
        codigo = clean_str(prod_row.get("codigo_producto", "")) or "Sin código"
        nombre = clean_str(prod_row.get("nombre_producto", "")) or "Sin nombre"
        categoria = clean_str(prod_row.get("categoria", "")) or "Sin categoría"
        marca_default = clean_str(prod_row.get("marca_default", "")) or "Sin marca predeterminada"
        unidad_default = clean_str(prod_row.get("unidad_default", "")) or "Sin unidad predeterminada"
        stock_minimo = clean_str(prod_row.get("stock_minimo", "")) or "0"
        dias_alerta = clean_str(prod_row.get("dias_alerta_vencimiento", "")) or "0"
        obs = clean_str(prod_row.get("observacion", ""))
        st.markdown(
            f"""
            <div class="mini-card" style="border-color:rgba(56,189,248,.30); background:linear-gradient(180deg, rgba(8,47,73,.38), rgba(15,23,42,.70));">
                <div class="mini-title">{titulo}</div>
                <div style="color:#F8FAFC; font-size:18px; font-weight:900; margin:3px 0 8px 0;">{nombre}</div>
                <div style="display:grid; grid-template-columns:repeat(6,minmax(0,1fr)); gap:8px;">
                    <div class="field-note"><b>Código</b><br>{codigo}</div>
                    <div class="field-note"><b>Categoría</b><br>{categoria}</div>
                    <div class="field-note"><b>Marca</b><br>{marca_default}</div>
                    <div class="field-note"><b>Unidad</b><br>{unidad_default}</div>
                    <div class="field-note"><b>Stock mínimo</b><br>{stock_minimo}</div>
                    <div class="field-note"><b>Alerta venc.</b><br>{dias_alerta} días</div>
                </div>
                {f'<div class="field-note" style="margin-top:10px;"><b>Observación:</b> {obs}</div>' if obs else ''}
            </div>
            """,
            unsafe_allow_html=True,
        )

    def _sync_after_save(rows: list[dict], message: str) -> None:
        sync_data = dict(data)
        sync_data["Movimientos"] = ensure_columns(
            pd.concat([ensure_columns(data["Movimientos"], "Movimientos"), pd.DataFrame(rows)], ignore_index=True),
            "Movimientos",
        )
        try:
            sync_kardex_consolidado_sheet(storage, sync_data)
            set_flash(message)
        except Exception as exc:
            set_flash(f"Registros guardados. No se pudo actualizar Kardex_Consolidado: {exc}", "warning")
        bump_form_nonce("mov_form_nonce")
        mark_data_dirty()
        rerun()

    tipo = st.radio(
        "Tipo de operación",
        TIPOS_MOVIMIENTO,
        horizontal=True,
        help="Salida permite seleccionar varios insumos en un carrito. Devolución se basa en las salidas registradas. Ajuste ahora se llama Corrección."
    )

    # ========================================================
    # INGRESO
    # ========================================================
    if tipo == "Ingreso":
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        st.markdown("#### 1) Producto / datos del catálogo")
        productos_ord = productos.sort_values(["nombre_producto", "codigo_producto"], na_position="last")
        labels = productos_ord.apply(
            lambda r: f"{clean_str(r['nombre_producto'])} | Código: {clean_str(r['codigo_producto']) or 'Sin código'} | Marca: {clean_str(r['marca_default']) or 'Sin marca'}",
            axis=1,
        ).tolist()
        producto_label = st.selectbox("Producto / reactivo / insumo registrado en catálogo", labels)
        prod_row = productos_ord.iloc[labels.index(producto_label)]
        producto_id = clean_str(prod_row["producto_id"])
        producto = clean_str(prod_row["nombre_producto"])
        marca_default = clean_str(prod_row.get("marca_default", ""))
        unidad_default = clean_str(prod_row.get("unidad_default", ""))
        _product_card(prod_row)
        st.markdown("</div>", unsafe_allow_html=True)

        mov_nonce = int(st.session_state.get("mov_form_nonce", 0))
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        with st.form(f"frm_ingreso_{mov_nonce}", clear_on_submit=True):
            st.markdown("#### 2) Datos del ingreso")
            c1, c2, c3 = st.columns(3)
            fecha = c1.date_input("Fecha del ingreso", value=TODAY.date())
            usuario = c2.text_input("Usuario que registra", value=st.session_state.get("nombre_usuario", "Usuario"))
            cantidad = c3.number_input("Cantidad", min_value=0.0, step=1.0, format="%.2f")

            st.markdown("#### 3) Datos del lote / presentación")
            unidades = sorted(set([u for u in UNIDADES_DEFAULT + [unidad_default] if clean_str(u)]))
            unidad_index = unidades.index(unidad_default) if unidad_default in unidades else 0
            a, b, c = st.columns(3)
            marca = a.text_input("Marca", value=marca_default, help="Se carga desde el catálogo del producto.")
            lote = b.text_input("Lote *", placeholder="Ejemplo: AB-2026-001")
            unidad = c.selectbox("Unidad", unidades, index=unidad_index)
            d, e, f = st.columns(3)
            fecha_elaboracion_dt = d.date_input("Fecha de elaboración", value=TODAY.date())
            fecha_vencimiento_dt = e.date_input("Fecha de vencimiento", value=(TODAY + pd.Timedelta(days=365)).date())
            costo_total = f.number_input("Costo total", min_value=0.0, step=1.0, format="%.2f")

            st.markdown("#### 4) Proveedor / compra / responsable")
            g, h, i = st.columns(3)
            proveedor = g.selectbox("Proveedor *", [""] + proveedores["proveedor"].dropna().astype(str).sort_values().tolist())
            orden_compra = h.text_input("Orden de compra", placeholder="Ejemplo: OC-2026-001")
            personal = i.selectbox("Personal que recibe", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            observacion = st.text_area("Observación", placeholder="Detalle de factura, donación, compra u otra referencia.")
            submitted = st.form_submit_button("💾 Guardar ingreso", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if st.session_state.get("rol") == "Consulta":
                st.error("El rol Consulta no puede registrar movimientos.")
                return
            if cantidad <= 0:
                st.error("La cantidad debe ser mayor a cero.")
                return
            if not lote:
                st.error("El lote es obligatorio.")
                return
            if not proveedor:
                st.error("Para ingresos debe seleccionar proveedor.")
                return
            row = {
                "movimiento_id": next_code("MOV", movimientos, "movimiento_id", 6),
                "fecha": pd.to_datetime(fecha).strftime("%Y-%m-%d"),
                "tipo_movimiento": "Ingreso",
                "producto_id": producto_id,
                "producto": producto,
                "marca": marca,
                "lote": lote,
                "proveedor": proveedor,
                "orden_compra": orden_compra,
                "solicitante": "",
                "personal": personal,
                "fecha_elaboracion": fecha_elaboracion_dt.strftime("%Y-%m-%d"),
                "fecha_vencimiento": fecha_vencimiento_dt.strftime("%Y-%m-%d"),
                "unidad": unidad,
                "cantidad": cantidad,
                "costo_total": costo_total,
                "observacion": observacion,
                "usuario_registro": usuario,
                "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                "acta_entrega_id": "",
            }
            storage.append_row("Movimientos", row)
            _sync_after_save([row], "Ingreso guardado correctamente. El formulario quedó limpio y el Kardex consolidado fue actualizado.")
        return

    # ========================================================
    # SALIDA CON CARRITO
    # ========================================================
    if tipo == "Salida":
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        st.markdown("#### 1) Armar carrito de salida")
        disponibles = stock[stock["stock_actual"] > 0].copy() if not stock.empty else pd.DataFrame()
        if disponibles.empty:
            st.error("No hay lotes con stock disponible para registrar salidas.")
            st.markdown("</div>", unsafe_allow_html=True)
            return

        if "salida_cart" not in st.session_state:
            st.session_state["salida_cart"] = []

        disponibles = disponibles.sort_values(["producto", "fecha_vencimiento", "lote"], na_position="last").reset_index(drop=True)
        disponibles["label"] = disponibles.apply(
            lambda r: f"{r['producto']} | Marca: {r['marca']} | Lote: {r['lote']} | Vence: {r['fecha_vencimiento']} | Stock: {float(r['stock_actual']):,.0f} {r['unidad']}",
            axis=1,
        )

        with st.form("frm_add_cart_salida", clear_on_submit=True):
            cprod, clote, ccant = st.columns([1, 1.9, .8])
            filtro_producto = cprod.selectbox("Producto", ["Todos"] + sorted(disponibles["producto"].dropna().astype(str).unique().tolist()))
            lotes_view = disponibles.copy()
            if filtro_producto != "Todos":
                lotes_view = lotes_view[lotes_view["producto"].astype(str) == filtro_producto]
            label = clote.selectbox("Lote disponible", lotes_view["label"].tolist())
            selected_stock_row = lotes_view[lotes_view["label"] == label].iloc[0]
            cantidad_item = ccant.number_input("Cantidad", min_value=0.0, step=1.0, format="%.2f")
            add_item = st.form_submit_button("➕ Agregar al carrito", use_container_width=True)

        if add_item:
            if cantidad_item <= 0:
                st.error("Ingrese una cantidad mayor a cero.")
            else:
                item_key = "|".join([
                    clean_str(selected_stock_row.get("producto_id", "")), clean_str(selected_stock_row.get("lote", "")),
                    clean_str(selected_stock_row.get("marca", "")), clean_str(selected_stock_row.get("fecha_vencimiento", "")),
                ])
                ya_en_carrito = sum(float(x.get("cantidad", 0) or 0) for x in st.session_state["salida_cart"] if x.get("item_key") == item_key)
                disponible = float(selected_stock_row.get("stock_actual", 0) or 0)
                if ya_en_carrito + cantidad_item > disponible:
                    st.error(f"No puede agregar {cantidad_item:,.0f}; ya tiene {ya_en_carrito:,.0f} en el carrito y el stock disponible es {disponible:,.0f}.")
                else:
                    st.session_state["salida_cart"].append({
                        "item_key": item_key,
                        "producto_id": clean_str(selected_stock_row["producto_id"]),
                        "producto": clean_str(selected_stock_row["producto"]),
                        "marca": clean_str(selected_stock_row["marca"]),
                        "lote": clean_str(selected_stock_row["lote"]),
                        "fecha_vencimiento": format_date(selected_stock_row["fecha_vencimiento"]),
                        "unidad": clean_str(selected_stock_row["unidad"]),
                        "cantidad": cantidad_item,
                        "stock_disponible": disponible,
                    })
                    st.success("Producto agregado al carrito.")
                    rerun()

        cart = st.session_state.get("salida_cart", [])
        if cart:
            st.markdown("#### 2) Productos seleccionados para la salida")
            cart_df = pd.DataFrame(cart).drop(columns=["item_key"], errors="ignore")
            cart_df.insert(0, "item_no", range(1, len(cart_df) + 1))
            st.dataframe(cart_df, use_container_width=True, hide_index=True)
            r1, r2 = st.columns([1, 2])
            quitar = r1.selectbox("Quitar item", [""] + [str(i) for i in range(1, len(cart) + 1)])
            if r2.button("🗑️ Quitar del carrito", use_container_width=True, disabled=(not quitar)):
                idx = int(quitar) - 1
                st.session_state["salida_cart"].pop(idx)
                rerun()
        else:
            st.info("Agregue uno o varios productos al carrito. Al guardar la salida, cada producto se registrará como una fila individual en Movimientos.")
        st.markdown("</div>", unsafe_allow_html=True)

        if st.session_state.get("last_acta_pdf"):
            st.download_button(
                "📄 Descargar última acta de entrega generada",
                data=st.session_state["last_acta_pdf"],
                file_name=st.session_state.get("last_acta_filename", "acta_entrega.pdf"),
                mime="application/pdf",
                use_container_width=True,
            )

        mov_nonce = int(st.session_state.get("mov_form_nonce", 0))
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        with st.form(f"frm_confirmar_salida_{mov_nonce}", clear_on_submit=True):
            st.markdown("#### 3) Datos de entrega / sitio receptor")
            a, b, c = st.columns(3)
            fecha = a.date_input("Fecha de salida", value=TODAY.date())
            usuario = b.text_input("Usuario que registra", value=st.session_state.get("nombre_usuario", "Usuario"))
            solicitante = c.selectbox("Sitio / unidad solicitante *", [""] + solicitantes["unidad_solicitante"].dropna().astype(str).sort_values().tolist())
            solicitante_info = get_first_match(solicitantes, "unidad_solicitante", solicitante)
            d, e, f = st.columns(3)
            personal = d.selectbox("Personal que entrega *", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            recibe_default = clean_str(solicitante_info.get("responsable", "")) if solicitante_info else ""
            recibe_nombre = e.text_input("Persona que recibe / firma", value=recibe_default)
            recibe_cargo = f.text_input("Cargo de quien recibe", value="Responsable del sitio")
            observacion = st.text_area("Observación para movimientos y acta", placeholder="Detalle de solicitud, referencia o comentario de entrega.")
            generar_acta = st.checkbox("Generar acta de entrega en PDF", value=True)
            submitted = st.form_submit_button("💾 Guardar salida y generar acta", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if st.session_state.get("rol") == "Consulta":
                st.error("El rol Consulta no puede registrar movimientos.")
                return
            cart = st.session_state.get("salida_cart", [])
            if not cart:
                st.error("Debe agregar al menos un producto al carrito.")
                return
            if not solicitante:
                st.error("Debe seleccionar el sitio/unidad solicitante.")
                return
            if not personal:
                st.error("Debe seleccionar el personal que entrega.")
                return
            # Validación final contra stock actual en pantalla.
            for item in cart:
                same = [x for x in cart if x.get("item_key") == item.get("item_key")]
                qty_same = sum(float(x.get("cantidad", 0) or 0) for x in same)
                if qty_same > float(item.get("stock_disponible", 0) or 0):
                    st.error(f"La cantidad del lote {item.get('lote')} supera el stock disponible.")
                    return
            acta_id = f"ACTA-{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}"
            ids = next_movement_ids(movimientos, len(cart))
            rows = []
            for mov_id, item in zip(ids, cart):
                rows.append({
                    "movimiento_id": mov_id,
                    "fecha": pd.to_datetime(fecha).strftime("%Y-%m-%d"),
                    "tipo_movimiento": "Salida",
                    "producto_id": item["producto_id"],
                    "producto": item["producto"],
                    "marca": item["marca"],
                    "lote": item["lote"],
                    "proveedor": "",
                    "orden_compra": "",
                    "solicitante": solicitante,
                    "personal": personal,
                    "fecha_elaboracion": "",
                    "fecha_vencimiento": item["fecha_vencimiento"],
                    "unidad": item["unidad"],
                    "cantidad": item["cantidad"],
                    "costo_total": 0,
                    "observacion": observacion,
                    "usuario_registro": usuario,
                    "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "acta_entrega_id": acta_id,
                })
            storage.append_rows("Movimientos", rows)
            if generar_acta:
                personal_info = get_first_match(personal_df, "nombre", personal)
                pdf_bytes = build_acta_entrega_pdf(rows, solicitante_info, personal_info, fecha, recibe_nombre, recibe_cargo, observacion=observacion)
                st.session_state["last_acta_pdf"] = pdf_bytes
                st.session_state["last_acta_filename"] = f"acta_entrega_{acta_id}_{clean_str(solicitante).replace(' ', '_')}.pdf"
                st.session_state["last_acta_id"] = acta_id
            st.session_state["salida_cart"] = []
            _sync_after_save(rows, f"Salida guardada correctamente: {len(rows)} insumo(s) registrados individualmente. Acta: {acta_id}.")
        return

    # ========================================================
    # DEVOLUCIÓN DESDE SALIDAS REGISTRADAS
    # ========================================================
    if tipo == "Devolución":
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        st.markdown("#### 1) Seleccionar salida relacionada")
        salidas = movimientos[movimientos["tipo_movimiento"].astype(str).isin(["Salida", "Corrección salida", "Ajuste salida"])].copy()
        if salidas.empty:
            st.info("Aún no hay salidas registradas para devolver.")
            st.markdown("</div>", unsafe_allow_html=True)
            return
        salidas["fecha_dt"] = to_date(salidas["fecha"])
        salidas = salidas.sort_values(["fecha_dt", "producto", "lote"], ascending=[False, True, True]).reset_index(drop=True)
        vista_salidas = salidas[["movimiento_id", "fecha", "producto", "marca", "lote", "fecha_vencimiento", "unidad", "cantidad", "solicitante", "personal", "acta_entrega_id"]].copy()
        st.dataframe(vista_salidas, use_container_width=True, hide_index=True)
        salidas["label"] = salidas.apply(lambda r: f"{r['movimiento_id']} | {r['fecha']} | {r['producto']} | Lote {r['lote']} | Cant. {r['cantidad']} {r['unidad']} | {r['solicitante']}", axis=1)
        label = st.selectbox("Salida a devolver", salidas["label"].tolist())
        salida_row = salidas[salidas["label"] == label].iloc[0]
        prod_row = _catalog_row_by_id(clean_str(salida_row["producto_id"]))
        if clean_str(prod_row.get("nombre_producto", "")) == "":
            prod_row["nombre_producto"] = salida_row["producto"]
            prod_row["marca_default"] = salida_row["marca"]
            prod_row["unidad_default"] = salida_row["unidad"]
        _product_card(prod_row, "Producto seleccionado desde una salida registrada")
        st.markdown("</div>", unsafe_allow_html=True)

        mov_nonce = int(st.session_state.get("mov_form_nonce", 0))
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        with st.form(f"frm_devolucion_{mov_nonce}", clear_on_submit=True):
            st.markdown("#### 2) Datos de devolución")
            a, b, c = st.columns(3)
            fecha = a.date_input("Fecha de devolución", value=TODAY.date())
            usuario = b.text_input("Usuario que registra", value=st.session_state.get("nombre_usuario", "Usuario"))
            max_qty = float(salida_row.get("cantidad", 0) or 0)
            cantidad = c.number_input("Cantidad a devolver", min_value=0.0, max_value=max_qty, step=1.0, format="%.2f")
            d, e, f = st.columns(3)
            d.text_input("Producto", value=clean_str(salida_row["producto"]), disabled=True)
            e.text_input("Lote", value=clean_str(salida_row["lote"]), disabled=True)
            f.text_input("Unidad", value=clean_str(salida_row["unidad"]), disabled=True)
            g, h = st.columns(2)
            solicitante_default = clean_str(salida_row.get("solicitante", ""))
            opciones_solic = [""] + solicitantes["unidad_solicitante"].dropna().astype(str).sort_values().tolist()
            idx_solic = opciones_solic.index(solicitante_default) if solicitante_default in opciones_solic else 0
            quien_devuelve = g.selectbox("Quién devuelve *", opciones_solic, index=idx_solic)
            personal = h.selectbox("Personal que recibe *", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            observacion = st.text_area("Observación", value=f"Devolución relacionada con salida {clean_str(salida_row['movimiento_id'])}")
            submitted = st.form_submit_button("💾 Guardar devolución", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if st.session_state.get("rol") == "Consulta":
                st.error("El rol Consulta no puede registrar movimientos.")
                return
            if cantidad <= 0:
                st.error("La cantidad a devolver debe ser mayor a cero.")
                return
            if not quien_devuelve:
                st.error("Seleccione quién devuelve.")
                return
            row = {
                "movimiento_id": next_code("MOV", movimientos, "movimiento_id", 6),
                "fecha": pd.to_datetime(fecha).strftime("%Y-%m-%d"),
                "tipo_movimiento": "Devolución",
                "producto_id": clean_str(salida_row["producto_id"]),
                "producto": clean_str(salida_row["producto"]),
                "marca": clean_str(salida_row["marca"]),
                "lote": clean_str(salida_row["lote"]),
                "proveedor": "",
                "orden_compra": "",
                "solicitante": quien_devuelve,
                "personal": personal,
                "fecha_elaboracion": clean_str(salida_row.get("fecha_elaboracion", "")),
                "fecha_vencimiento": format_date(salida_row.get("fecha_vencimiento", "")),
                "unidad": clean_str(salida_row["unidad"]),
                "cantidad": cantidad,
                "costo_total": 0,
                "observacion": observacion,
                "usuario_registro": usuario,
                "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                "acta_entrega_id": clean_str(salida_row.get("acta_entrega_id", "")),
            }
            storage.append_row("Movimientos", row)
            _sync_after_save([row], "Devolución guardada correctamente. El lote fue actualizado en Kardex consolidado.")
        return

    # ========================================================
    # CORRECCIÓN ENTRADA / SALIDA
    # ========================================================
    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    st.markdown("#### 1) Seleccionar producto/lote para corrección")
    disponibles = stock.copy() if not stock.empty else pd.DataFrame()
    if disponibles.empty:
        st.error("No hay lotes registrados para aplicar una corrección.")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    if tipo == "Corrección salida":
        disponibles = disponibles[disponibles["stock_actual"] > 0].copy()
    disponibles = disponibles.sort_values(["producto", "fecha_vencimiento", "lote"], na_position="last").reset_index(drop=True)
    disponibles["label"] = disponibles.apply(
        lambda r: f"{r['producto']} | Marca: {r['marca']} | Lote: {r['lote']} | Vence: {r['fecha_vencimiento']} | Stock: {float(r['stock_actual']):,.0f} {r['unidad']}",
        axis=1,
    )
    label = st.selectbox("Producto/lote a corregir", disponibles["label"].tolist())
    selected = disponibles[disponibles["label"] == label].iloc[0]
    prod_row = _catalog_row_by_id(clean_str(selected["producto_id"]))
    if clean_str(prod_row.get("nombre_producto", "")) == "":
        prod_row["nombre_producto"] = selected["producto"]
        prod_row["marca_default"] = selected["marca"]
        prod_row["unidad_default"] = selected["unidad"]
    _product_card(prod_row, "Producto seleccionado para corrección")
    st.info(f"Stock actual del lote: {float(selected.get('stock_actual', 0) or 0):,.0f} {clean_str(selected.get('unidad', ''))}.")
    st.markdown("</div>", unsafe_allow_html=True)

    mov_nonce = int(st.session_state.get("mov_form_nonce", 0))
    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    with st.form(f"frm_correccion_{mov_nonce}", clear_on_submit=True):
        st.markdown("#### 2) Datos de corrección")
        a, b, c = st.columns(3)
        fecha = a.date_input("Fecha de corrección", value=TODAY.date())
        usuario = b.text_input("Usuario que registra", value=st.session_state.get("nombre_usuario", "Usuario"))
        cantidad = c.number_input("Cantidad", min_value=0.0, step=1.0, format="%.2f")
        d, e, f = st.columns(3)
        d.text_input("Producto", value=clean_str(selected["producto"]), disabled=True)
        e.text_input("Lote", value=clean_str(selected["lote"]), disabled=True)
        f.text_input("Unidad", value=clean_str(selected["unidad"]), disabled=True)
        g, h = st.columns(2)
        persona_relacionada = g.selectbox("Persona / sitio relacionado *", [""] + solicitantes["unidad_solicitante"].dropna().astype(str).sort_values().tolist())
        personal = h.selectbox("Personal responsable *", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
        observacion = st.text_area("Justificación de la corrección *", placeholder="Explique la razón de la corrección para auditoría.")
        submitted = st.form_submit_button("💾 Guardar corrección", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        if st.session_state.get("rol") == "Consulta":
            st.error("El rol Consulta no puede registrar movimientos.")
            return
        if cantidad <= 0:
            st.error("La cantidad debe ser mayor a cero.")
            return
        if tipo == "Corrección salida" and cantidad > float(selected.get("stock_actual", 0) or 0):
            st.error("La corrección de salida no puede superar el stock actual del lote.")
            return
        if not persona_relacionada:
            st.error("Seleccione la persona o sitio relacionado con la corrección.")
            return
        if not personal:
            st.error("Seleccione el personal responsable.")
            return
        if not observacion:
            st.error("La justificación es obligatoria para registrar una corrección.")
            return
        row = {
            "movimiento_id": next_code("MOV", movimientos, "movimiento_id", 6),
            "fecha": pd.to_datetime(fecha).strftime("%Y-%m-%d"),
            "tipo_movimiento": tipo,
            "producto_id": clean_str(selected["producto_id"]),
            "producto": clean_str(selected["producto"]),
            "marca": clean_str(selected["marca"]),
            "lote": clean_str(selected["lote"]),
            "proveedor": "",
            "orden_compra": "",
            "solicitante": persona_relacionada,
            "personal": personal,
            "fecha_elaboracion": "",
            "fecha_vencimiento": format_date(selected.get("fecha_vencimiento", "")),
            "unidad": clean_str(selected["unidad"]),
            "cantidad": cantidad,
            "costo_total": 0,
            "observacion": observacion,
            "usuario_registro": usuario,
            "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            "acta_entrega_id": "",
        }
        storage.append_row("Movimientos", row)
        _sync_after_save([row], "Corrección guardada correctamente. El Kardex consolidado fue actualizado.")

def page_kardex_consolidado(kardex: pd.DataFrame, storage=None, data: Dict[str, pd.DataFrame] | None = None, mode: str = "") -> None:
    section_header(
        "📋 Kardex consolidado por lote",
        "Una fila por producto/lote: ingreso, salida acumulada, último destinatario, fecha de entrega y saldo actual."
    )

    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    c_sync, c_format, c_msg = st.columns([.95, .95, 2.1])
    with c_sync:
        if storage is not None and data is not None and st.button("🔄 Actualizar hoja Kardex_Consolidado", use_container_width=True):
            try:
                sync_kardex_consolidado_sheet(storage, data)
                st.success("Hoja Kardex_Consolidado actualizada correctamente en la base.")
            except Exception as exc:
                st.error(f"No se pudo actualizar la hoja Kardex_Consolidado: {exc}")
    with c_format:
        if storage is not None and hasattr(storage, "apply_table_format") and st.button("🎨 Formato tabla", use_container_width=True):
            try:
                storage.apply_table_format("Kardex_Consolidado", data_rows=len(kardex), strict=True)
                st.success("Formato tipo tabla aplicado en Kardex_Consolidado.")
            except Exception as exc:
                st.error(f"No se pudo aplicar formato tabla: {exc}")
    with c_msg:
        st.caption("Esta tabla se calcula desde Movimientos. Use el botón para crear/actualizar la pestaña Kardex_Consolidado en Google Sheets y dejarla con formato tabla.")
    st.markdown("</div>", unsafe_allow_html=True)

    if kardex.empty:
        st.info("No hay movimientos registrados para consolidar.")
        return

    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns([1, 1, 1, .9])
    estados = sorted(kardex["estado"].dropna().astype(str).unique().tolist())
    estado = c1.multiselect("Estado", estados, default=estados)
    productos = c2.multiselect("Producto", sorted(kardex["producto"].dropna().astype(str).unique().tolist()))
    lote = c3.text_input("Buscar lote")
    vista = c4.selectbox("Vista", ["Con stock", "Todos", "Con salidas", "Sin salidas", "Sin stock"])
    st.markdown("</div>", unsafe_allow_html=True)

    df = kardex.copy()
    if estado:
        df = df[df["estado"].isin(estado)]
    if productos:
        df = df[df["producto"].isin(productos)]
    if lote:
        df = df[df["lote"].astype(str).str.contains(lote, case=False, na=False)]
    if vista == "Con stock":
        df = df[df["saldo_actual"] > 0]
    elif vista == "Con salidas":
        df = df[df["salida_total"] > 0]
    elif vista == "Sin salidas":
        df = df[df["salida_total"] <= 0]
    elif vista == "Sin stock":
        df = df[df["saldo_actual"] <= 0]

    total_ingresado = float(df["entrada_total"].sum()) if not df.empty else 0
    total_salida = float(df["salida_total"].sum()) if not df.empty else 0
    total_saldo = float(df["saldo_actual"].sum()) if not df.empty else 0
    k1, k2, k3, k4 = st.columns(4)
    with k1: kpi_card("Lotes en vista", f"{len(df):,}", "Filas consolidadas")
    with k2: kpi_card("Entrada total", f"{total_ingresado:,.0f}", "Cantidad acumulada")
    with k3: kpi_card("Salida total", f"{total_salida:,.0f}", "Cantidad entregada")
    with k4: kpi_card("Saldo actual", f"{total_saldo:,.0f}", "Cantidad restante")

    st.write("")
    card_start("Tabla consolidada", "La salida se refleja en la misma fila del ingreso/lote. Use 'Detalle salidas' para ver entregas múltiples.")
    cols_view = [
        "estado", "producto", "marca", "lote", "unidad", "fecha_ingreso", "proveedor_ingreso",
        "fecha_elaboracion", "fecha_vencimiento", "entrada_total", "salida_total", "saldo_actual",
        "numero_salidas", "fecha_ultima_salida", "ultimo_entregado_a", "ultimo_personal_entrega",
        "detalle_salidas", "dias_para_vencer", "stock_minimo",
    ]
    st.dataframe(
        df[cols_view],
        use_container_width=True,
        hide_index=True,
        column_config={
            "entrada_total": st.column_config.NumberColumn("Entrada", format="%.0f"),
            "salida_total": st.column_config.NumberColumn("Salida acumulada", format="%.0f"),
            "saldo_actual": st.column_config.NumberColumn("Saldo actual", format="%.0f"),
            "numero_salidas": st.column_config.NumberColumn("No. salidas", format="%d"),
            "detalle_salidas": st.column_config.TextColumn("Detalle salidas", width="large"),
        },
    )

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Kardex_Consolidado")
    output.seek(0)
    st.download_button(
        "⬇️ Descargar Kardex consolidado filtrado",
        data=output.read(),
        file_name=f"kardex_consolidado_{TODAY.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def page_stock(stock: pd.DataFrame) -> None:
    section_header("📦 Stock y alertas", "Consulta por producto, lote, vencimiento y estado del inventario.")
    if stock.empty:
        st.info("No hay movimientos registrados.")
        return
    col1, col2, col3, col4 = st.columns([1, 1, 1, .8])
    estado = col1.multiselect("Estado", sorted(stock["estado"].dropna().unique()), default=sorted(stock["estado"].dropna().unique()))
    producto = col2.multiselect("Producto", sorted(stock["producto"].dropna().unique()))
    lote = col3.text_input("Buscar lote")
    solo_con_stock = col4.toggle("Solo con stock", value=True)

    df = stock.copy()
    if estado:
        df = df[df["estado"].isin(estado)]
    if producto:
        df = df[df["producto"].isin(producto)]
    if lote:
        df = df[df["lote"].astype(str).str.contains(lote, case=False, na=False)]
    if solo_con_stock:
        df = df[df["stock_actual"] > 0]

    st.dataframe(
        df[["estado", "producto", "marca", "lote", "fecha_vencimiento", "dias_para_vencer", "unidad", "ingreso_total", "salida_total", "stock_actual", "stock_minimo"]],
        use_container_width=True,
        hide_index=True,
    )


def page_reportes(data: Dict[str, pd.DataFrame], stock: pd.DataFrame, kardex: pd.DataFrame) -> None:
    section_header("📊 Reportes", "Filtre, analice y exporte movimientos, stock y alertas.")
    movimientos = data["Movimientos"].copy()
    movimientos["fecha_dt"] = to_date(movimientos["fecha"])

    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    fecha_min = col1.date_input("Desde", value=(TODAY - pd.Timedelta(days=90)).date())
    fecha_max = col2.date_input("Hasta", value=TODAY.date())
    tipo = col3.multiselect("Tipo de movimiento", TIPOS_MOVIMIENTO, default=TIPOS_MOVIMIENTO)
    producto = st.multiselect("Producto", sorted(movimientos["producto"].dropna().astype(str).unique().tolist()))
    st.markdown("</div>", unsafe_allow_html=True)

    filtro = movimientos.copy()
    filtro = filtro[(filtro["fecha_dt"] >= pd.to_datetime(fecha_min)) & (filtro["fecha_dt"] <= pd.to_datetime(fecha_max))]
    if tipo:
        filtro = filtro[filtro["tipo_movimiento"].isin(tipo)]
    if producto:
        filtro = filtro[filtro["producto"].isin(producto)]

    card_start("Movimientos filtrados", f"Registros encontrados: {len(filtro):,}")
    st.dataframe(filtro.drop(columns=["fecha_dt"], errors="ignore"), use_container_width=True, hide_index=True)

    card_start("Resumen por producto y tipo", "Suma de cantidades filtradas.")
    if filtro.empty:
        st.info("No hay movimientos para el filtro seleccionado.")
    else:
        res = filtro.assign(cantidad_num=to_number(filtro["cantidad"])).groupby(["producto", "tipo_movimiento"], as_index=False)["cantidad_num"].sum()
        fig = px.bar(res, x="producto", y="cantidad_num", color="tipo_movimiento", barmode="group")
        fig.update_layout(height=430, margin=dict(l=10, r=10, t=20, b=120), xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

    report_bytes = build_excel_report(data, stock, kardex)
    st.download_button(
        "⬇️ Descargar reporte completo en Excel",
        data=report_bytes,
        file_name=f"reporte_kardex_{TODAY.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ------------------------- CATÁLOGOS -------------------------
def save_new_product(storage, data, values: dict) -> None:
    df = ensure_columns(data["Productos"], "Productos")
    row = {
        "producto_id": next_code("PRD", df, "producto_id", 4),
        **values,
    }
    storage.append_row("Productos", row)


def product_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo producto", "Registre productos con stock mínimo y días de alerta por vencimiento.")
    with st.form("frm_producto", clear_on_submit=True):
        c1, c2 = st.columns([1, 2])
        codigo = c1.text_input("Código interno", placeholder="Ejemplo: HIV-RAP-001")
        nombre = c2.text_input("Nombre del producto *", placeholder="Nombre completo del reactivo o insumo")
        c3, c4, c5 = st.columns(3)
        categoria = c3.selectbox("Categoría", CATEGORIAS_DEFAULT)
        marca = c4.text_input("Marca predeterminada")
        unidad = c5.selectbox("Unidad predeterminada", UNIDADES_DEFAULT)
        c6, c7, c8 = st.columns(3)
        minimo = c6.number_input("Stock mínimo", min_value=0.0, step=1.0, value=5.0)
        dias_alerta = c7.number_input("Días alerta vencimiento", min_value=1, step=1, value=90)
        activo = c8.selectbox("Estado", ["Sí", "No"])
        observacion = st.text_area("Observación")
        submitted = st.form_submit_button("➕ Guardar producto", use_container_width=True)
    if submitted:
        if not nombre:
            st.error("El nombre del producto es obligatorio.")
            return
        save_new_product(storage, data, {
            "codigo_producto": codigo or nombre[:18].upper().replace(" ", "-"),
            "nombre_producto": nombre,
            "categoria": categoria,
            "marca_default": marca,
            "unidad_default": unidad,
            "stock_minimo": minimo,
            "dias_alerta_vencimiento": dias_alerta,
            "activo": activo,
            "observacion": observacion,
        })
        set_flash("Producto guardado correctamente. El formulario quedó limpio para registrar un nuevo producto.")
        rerun()


def provider_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo proveedor", "Formulario ordenado por datos generales, contacto y ubicación.")
    with st.form("frm_proveedor", clear_on_submit=True):
        st.markdown("##### Datos generales")
        c1, c2 = st.columns([2, 1])
        proveedor = c1.text_input("Nombre del proveedor *", placeholder="Nombre comercial o razón social")
        ruc = c2.text_input("RUC / RTN")
        descripcion = st.text_input("Descripción / tipo de proveedor", placeholder="Ejemplo: Distribuidor de reactivos, insumos médicos, papelería")
        st.markdown("##### Contacto")
        c3, c4, c5 = st.columns(3)
        representante = c3.text_input("Representante")
        telefono = c4.text_input("Teléfono")
        correo = c5.text_input("Correo")
        st.markdown("##### Ubicación y estado")
        c6, c7 = st.columns([2, 1])
        direccion = c6.text_area("Dirección")
        activo = c7.selectbox("Estado", ["Sí", "No"])
        submitted = st.form_submit_button("➕ Guardar proveedor", use_container_width=True)
    if submitted:
        if not proveedor:
            st.error("El nombre del proveedor es obligatorio.")
            return
        df = ensure_columns(data["Proveedores"], "Proveedores")
        row = {
            "proveedor_id": next_code("PROV", df, "proveedor_id", 4),
            "proveedor": proveedor,
            "descripcion": descripcion,
            "ruc": ruc,
            "representante": representante,
            "telefono": telefono,
            "correo": correo,
            "direccion": direccion,
            "activo": activo,
        }
        storage.append_row("Proveedores", row)
        set_flash("Proveedor guardado correctamente. El formulario quedó limpio para registrar un nuevo proveedor.")
        rerun()


def requester_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo solicitante / unidad", "Registre unidades, áreas o sitios que pueden solicitar productos.")
    with st.form("frm_solicitante", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        unidad = c1.text_input("Unidad solicitante *", placeholder="Ejemplo: Hospital, laboratorio, componente, sitio")
        departamento = c2.text_input("Departamento")
        municipio = c3.text_input("Municipio")
        c4, c5, c6 = st.columns(3)
        responsable = c4.text_input("Responsable")
        telefono = c5.text_input("Teléfono")
        correo = c6.text_input("Correo")
        activo = st.selectbox("Estado", ["Sí", "No"])
        submitted = st.form_submit_button("➕ Guardar solicitante", use_container_width=True)
    if submitted:
        if not unidad:
            st.error("La unidad solicitante es obligatoria.")
            return
        df = ensure_columns(data["Solicitantes"], "Solicitantes")
        row = {
            "solicitante_id": next_code("SOL", df, "solicitante_id", 4),
            "unidad_solicitante": unidad,
            "departamento": departamento,
            "municipio": municipio,
            "responsable": responsable,
            "telefono": telefono,
            "correo": correo,
            "activo": activo,
        }
        storage.append_row("Solicitantes", row)
        set_flash("Solicitante guardado correctamente. El formulario quedó limpio para registrar una nueva unidad solicitante.")
        rerun()


def staff_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo personal", "Usuarios operativos que reciben, entregan o registran movimientos.")
    with st.form("frm_personal", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([2, 1, 1.4, .8])
        nombre = c1.text_input("Nombre completo *")
        cargo = c2.text_input("Cargo")
        correo = c3.text_input("Correo")
        activo = c4.selectbox("Estado", ["Sí", "No"])
        submitted = st.form_submit_button("➕ Guardar personal", use_container_width=True)
    if submitted:
        if not nombre:
            st.error("El nombre es obligatorio.")
            return
        df = ensure_columns(data["Personal"], "Personal")
        row = {
            "personal_id": next_code("PER", df, "personal_id", 4),
            "nombre": nombre,
            "cargo": cargo,
            "correo": correo,
            "activo": activo,
        }
        storage.append_row("Personal", row)
        set_flash("Personal guardado correctamente. El formulario quedó limpio para registrar un nuevo personal.")
        rerun()


def catalog_editor(storage, data: Dict[str, pd.DataFrame], sheet: str, title: str) -> None:
    card_start(title, "Busque, revise y edite la tabla completa cuando necesite ajustes masivos.")
    df = ensure_columns(data[sheet], sheet)
    search = st.text_input(f"Buscar en {title}", key=f"search_{sheet}")
    view = df.copy()
    if search:
        mask = view.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        view = view[mask]
    edited = st.data_editor(view, num_rows="dynamic", use_container_width=True, key=f"editor_{sheet}")
    if st.button(f"💾 Guardar cambios en {title}", key=f"save_{sheet}", use_container_width=True):
        if search:
            st.warning("Para guardar edición masiva, quite el filtro de búsqueda para evitar perder filas no visibles.")
            return
        edited_df = ensure_columns(pd.DataFrame(edited), sheet)
        storage.save(sheet, edited_df)
        if sheet == "Productos":
            sync_data = dict(data)
            sync_data["Productos"] = edited_df
            try:
                sync_kardex_consolidado_sheet(storage, sync_data)
                set_flash(f"{title} actualizado correctamente. Kardex consolidado sincronizado.")
            except Exception as exc:
                set_flash(f"{title} actualizado, pero no se pudo sincronizar Kardex_Consolidado: {exc}", "warning")
        else:
            set_flash(f"{title} actualizado correctamente.")
        rerun()


def page_catalogos(storage, data: Dict[str, pd.DataFrame]) -> None:
    section_header("⚙️ Catálogos", "Formularios independientes para mantener ordenados productos, proveedores, solicitantes y personal.")
    if st.session_state.get("rol") == "Consulta":
        st.info("Su rol es de consulta. Puede visualizar catálogos, pero no registrar cambios.")

    tab_prod, tab_prov, tab_sol, tab_per = st.tabs(["📦 Productos", "🚚 Proveedores", "🏥 Solicitantes", "👤 Personal"])

    with tab_prod:
        if st.session_state.get("rol") != "Consulta":
            product_form(storage, data)
        catalog_editor(storage, data, "Productos", "Listado de productos")
    with tab_prov:
        if st.session_state.get("rol") != "Consulta":
            provider_form(storage, data)
        catalog_editor(storage, data, "Proveedores", "Listado de proveedores")
    with tab_sol:
        if st.session_state.get("rol") != "Consulta":
            requester_form(storage, data)
        catalog_editor(storage, data, "Solicitantes", "Listado de solicitantes")
    with tab_per:
        if st.session_state.get("rol") != "Consulta":
            staff_form(storage, data)
        catalog_editor(storage, data, "Personal", "Listado de personal")


def page_importar(storage, data: Dict[str, pd.DataFrame]) -> None:
    section_header("🧾 Importar Kardex anterior", "Convierte la hoja MOVIMIENTO del archivo con macros a la nueva estructura.")
    if st.session_state.get("rol") == "Consulta":
        st.warning("El rol Consulta no puede importar registros.")
        return
    st.write("Suba el archivo `.xlsm` anterior para leer la hoja MOVIMIENTO y convertir los registros a la nueva estructura.")
    uploaded = st.file_uploader("Archivo KARDEX anterior", type=["xlsm", "xlsx"])
    if uploaded:
        try:
            mov_new, prod_new = parse_legacy_kardex(uploaded)
            st.success(f"Se detectaron {len(mov_new)} movimientos y {len(prod_new)} productos únicos.")
            st.dataframe(mov_new.head(50), use_container_width=True, hide_index=True)
            if st.button("Importar registros detectados", use_container_width=True):
                prod_actual = data["Productos"]
                prod_final = pd.concat([prod_actual, prod_new], ignore_index=True).drop_duplicates("nombre_producto", keep="first")
                mov_final = pd.concat([data["Movimientos"], mov_new], ignore_index=True)
                prod_final = ensure_columns(prod_final, "Productos")
                mov_final = ensure_columns(mov_final, "Movimientos")
                storage.save("Productos", prod_final)
                storage.save("Movimientos", mov_final)
                sync_data = dict(data)
                sync_data["Productos"] = prod_final
                sync_data["Movimientos"] = mov_final
                try:
                    sync_kardex_consolidado_sheet(storage, sync_data)
                    set_flash("Importación completada. Kardex consolidado actualizado en la hoja de base.")
                except Exception as exc:
                    set_flash(f"Importación completada, pero no se pudo actualizar Kardex_Consolidado: {exc}", "warning")
                rerun()
        except Exception as exc:
            st.error(f"No se pudo importar el archivo. Revise que tenga la hoja MOVIMIENTO con la estructura esperada. Detalle: {exc}")


def page_admin(storage, data: Dict[str, pd.DataFrame], mode: str) -> None:
    section_header("🔐 Administración", "Gestión de usuarios, seguridad PATH automática y diagnóstico de conexión a la base.")
    if not require_admin():
        return

    tab1, tab2, tab3 = st.tabs(["Usuarios", "Seguridad PATH", "Diagnóstico"])

    with tab1:
        card_start("Crear usuario", "Asigne rol y credenciales de acceso.")
        with st.form("frm_usuario", clear_on_submit=True):
            c1, c2, c3 = st.columns([1, 1.5, 1])
            usuario = c1.text_input("Usuario *")
            nombre = c2.text_input("Nombre completo *")
            rol = c3.selectbox("Rol", ROLES)
            c4, c5 = st.columns([1.4, .7])
            password = c4.text_input("Contraseña inicial *", type="password")
            activo = c5.selectbox("Estado", ["Sí", "No"])
            st.caption("El PATH del usuario se valida con código temporal generado automáticamente en el login; no se asigna manualmente.")
            submitted = st.form_submit_button("➕ Crear usuario", use_container_width=True)
        if submitted:
            if not usuario or not nombre or not password:
                st.error("Usuario, nombre y contraseña son obligatorios.")
            else:
                users = ensure_columns(data["Usuarios"], "Usuarios")
                exists = users["usuario"].astype(str).str.lower().str.strip().eq(usuario.lower().strip()).any()
                if exists:
                    st.error("Ya existe un usuario con ese nombre.")
                else:
                    row = {
                        "usuario_id": next_code("USR", users, "usuario_id", 4),
                        "usuario": usuario,
                        "nombre": nombre,
                        "rol": rol,
                        "password_hash": hash_password(password),
                        "path_verificacion": "DINAMICO",
                        "activo": activo,
                        "fecha_creacion": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    storage.append_row("Usuarios", row)
                    set_flash("Usuario creado correctamente. El formulario quedó limpio para crear otro usuario.")
                    rerun()

        card_start("Usuarios registrados", "Por seguridad no se muestra la contraseña ni el hash completo.")
        users_view = ensure_columns(data["Usuarios"], "Usuarios").copy()
        if not users_view.empty:
            users_view["password_hash"] = users_view["password_hash"].astype(str).str[:10] + "..."
            users_view["path_verificacion"] = "Automático"
        st.dataframe(users_view, use_container_width=True, hide_index=True)

    with tab2:
        card_start("PATH automático temporal", "El sistema genera un código PATH visible en la pantalla de login. El usuario lo copia y lo pega para completar el acceso.")
        st.markdown(
            """
            <div class='alert-green'>
            ✅ El PATH ya no es una clave fija ni se configura manualmente. Cada sesión de login genera un código temporal con vigencia limitada.
            </div>
            """,
            unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Formato", "KDX-XXXX-XXXX")
        with c2:
            st.metric("Vigencia", f"{PATH_TTL_MINUTES} min")
        with c3:
            st.metric("Modo", "Automático")
        st.info("En la pantalla de login aparecerá el campo 'Código PATH generado'. El usuario debe copiar ese código y pegarlo en 'Pegar PATH generado'.")

    with tab3:
        card_start("Diagnóstico de base y estructura", "Verifica hojas cargadas, conexión activa y evita lecturas innecesarias a Google Sheets.")
        if mode == "Google Sheets":
            st.info("Modo optimizado V16: conexión por API REST directa, lectura por lotes, caché y formato tipo tabla en Google Sheets.")
        expected = list(SHEET_COLUMNS.keys())
        status_rows = []
        for sheet in expected:
            df = ensure_columns(data.get(sheet, pd.DataFrame()), sheet)
            if sheet == "Kardex_Consolidado":
                estado_hoja = "Calculada desde Movimientos; se sincroniza a Google Sheets con el botón"
            else:
                estado_hoja = "OK"
            status_rows.append({"Hoja": sheet, "Estado": estado_hoja, "Filas cargadas/calculadas": len(df), "Columnas esperadas": len(SHEET_COLUMNS[sheet])})
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)
        if mode == "Excel local":
            st.warning(f"Base activa: Excel local. Ruta: {DB_FILE.resolve()}. Si está en Streamlit Cloud, estos datos no se verán en Google Sheets.")
        else:
            info = storage.info() if hasattr(storage, "info") else {}
            st.success("Base activa: Google Sheets. Los registros deben guardarse en las pestañas del archivo conectado.")
            st.caption(f"Google Sheet ID: {info.get('sheet_id', '')}")
            st.caption(f"Cuenta de servicio: {info.get('client_email', '')}")
            if st.button("🧪 Probar escritura en Google Sheets", use_container_width=True):
                try:
                    result = storage.test_write()
                    st.success(f"Prueba guardada correctamente en la pestaña Config: {result['timestamp']}")
                    st.json(result.get("response", {}))
                except Exception as exc:
                    st.error(f"No se pudo escribir en Google Sheets: {exc}")
                    st.info(diagnose_gsheets_error(exc))
            if st.button("🔄 Crear/actualizar hoja Kardex_Consolidado", use_container_width=True):
                try:
                    sync_kardex_consolidado_sheet(storage, data)
                    st.success("Kardex_Consolidado actualizado correctamente en Google Sheets.")
                except Exception as exc:
                    st.error(f"No se pudo sincronizar Kardex_Consolidado: {exc}")

            if hasattr(storage, "apply_table_format_all") and st.button("🎨 Aplicar formato tabla a Google Sheets", use_container_width=True):
                try:
                    data_fmt = dict(data)
                    data_fmt["Kardex_Consolidado"] = calcular_kardex_consolidado(data.get("Movimientos", pd.DataFrame()), data.get("Productos", pd.DataFrame()))
                    storage.apply_table_format_all(data_fmt, strict=True)
                    st.success("Formato tipo tabla aplicado correctamente a las pestañas de Google Sheets.")
                except Exception as exc:
                    st.error(f"No se pudo aplicar el formato tabla: {exc}")
                    st.info(diagnose_gsheets_error(exc))

# ============================================================
# APP PRINCIPAL
# ============================================================
def main() -> None:
    apply_theme()
    storage, mode = get_storage()
    try:
        data = load_all(storage, mode)
    except Exception as exc:
        st.error(
            "La conexión se creó, pero falló al leer la estructura de la base. "
            "El sistema se detuvo para evitar guardar datos en un lugar incorrecto."
        )
        st.code(exception_detail(exc))
        st.info(diagnose_gsheets_error(exc))
        st.stop()

    if not render_login(storage, data, mode):
        show_flash()
        return

    if not enforce_inactivity_timeout():
        return
    inject_inactivity_watcher()

    stock = calcular_stock(data["Movimientos"], data["Productos"])
    kardex = calcular_kardex_consolidado(data["Movimientos"], data["Productos"])
    # Mantiene disponible la tabla calculada dentro del diccionario de datos
    # sin exigir leerla desde Google Sheets al iniciar.
    data["Kardex_Consolidado"] = ensure_columns(kardex, "Kardex_Consolidado")
    hero(mode, st.session_state.get("nombre_usuario", ""))
    show_flash()

    allowed_pages = allowed_nav_pages_for_role(st.session_state.get("rol", ""))
    if "page" not in st.session_state or st.session_state["page"] not in allowed_pages:
        st.session_state["page"] = PAGE_INICIO

    with st.sidebar:
        st.markdown("### 📌 Ruta Kardex")
        st.caption(f"Conectado como: **{st.session_state.get('nombre_usuario', '')}**")
        st.caption(f"Rol: **{st.session_state.get('rol', '')}**")
        st.divider()
        st.markdown("**Orden recomendado de navegación**")
        st.caption("Siga la ruta de arriba hacia abajo para no perderse en el proceso.")
        page = st.radio(
            "Seleccione módulo",
            allowed_pages,
            index=allowed_pages.index(st.session_state["page"]),
            label_visibility="collapsed",
        )
        st.session_state["page"] = page
        st.divider()
        st.markdown("**Estructura lógica**")
        st.caption("1. Acceso → 2. Catálogos → 3. Movimientos → 4. Kardex/Stock → 5. Reportes")
        st.divider()
        if st.button("Cerrar sesión", use_container_width=True):
            logout()
        st.caption("Sistema Kardex PRO: control por lote, vencimiento, stock mínimo y trazabilidad.")

    if page == PAGE_INICIO:
        page_inicio_operativo(data, stock, kardex, mode)
    elif page == PAGE_CATALOGOS:
        page_catalogos(storage, data)
    elif page == PAGE_ADMIN and st.session_state.get("rol") == "Administrador":
        page_admin(storage, data, mode)
    elif page == PAGE_MOVIMIENTOS:
        page_movimiento(storage, data, stock)
    elif page == PAGE_KARDEX:
        page_kardex_consolidado(kardex, storage, data, mode)
    elif page == PAGE_STOCK:
        page_stock(stock)
    elif page == PAGE_DASHBOARD:
        page_dashboard(data, stock)
    elif page == PAGE_REPORTES:
        page_reportes(data, stock, kardex)
    elif page == PAGE_IMPORTAR:
        page_importar(storage, data)


if __name__ == "__main__":
    main()
