from __future__ import annotations

import hashlib
import re
import secrets
import string
import time
import uuid
import unicodedata
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
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DB_FILE = DATA_DIR / "kardex_db.xlsx"
ASSETS_DIR = APP_DIR / "assets"
ACTA_LOGO_PATH = ASSETS_DIR / "logo_vihca.png"
ACTA_LOGO_B64 = """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"""
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
        "unidad_default", "stock_minimo", "dias_alerta_vencimiento", "activo", "observacion",
        "estado_registro", "creado_por", "fecha_creacion", "modificado_por", "fecha_modificacion", "motivo_modificacion"
    ],
    "Proveedores": [
        "proveedor_id", "proveedor", "descripcion", "ruc", "representante", "telefono",
        "correo", "direccion", "activo",
        "estado_registro", "creado_por", "fecha_creacion", "modificado_por", "fecha_modificacion", "motivo_modificacion"
    ],
    "Solicitantes": [
        "solicitante_id", "unidad_solicitante", "departamento", "municipio", "responsable",
        "telefono", "correo", "activo",
        "estado_registro", "creado_por", "fecha_creacion", "modificado_por", "fecha_modificacion", "motivo_modificacion"
    ],
    "Personal": [
        "personal_id", "nombre", "cargo", "correo", "activo",
        "estado_registro", "creado_por", "fecha_creacion", "modificado_por", "fecha_modificacion", "motivo_modificacion"
    ],
    "Usuarios": [
        "usuario_id", "usuario", "nombre", "rol", "password_hash", "path_verificacion", "activo", "fecha_creacion"
    ],
    "Movimientos": [
        "movimiento_id", "fecha", "tipo_movimiento", "producto_id", "producto", "marca", "lote",
        "proveedor", "orden_compra", "solicitante", "personal", "fecha_elaboracion", "fecha_vencimiento",
        "unidad", "cantidad", "costo_total", "observacion", "usuario_registro", "fecha_registro", "acta_entrega_id",
        "estado_movimiento", "anulado_por", "fecha_anulacion", "motivo_anulacion",
        "modificado_por", "fecha_modificacion", "motivo_modificacion"
    ],
    # Hoja física en Google Sheets/Excel. Se calcula automáticamente desde Movimientos.
    # Esta hoja permite ver el stock actual en la misma base, sin depender solo de la vista de Streamlit.
    "Kardex_Consolidado": KARDEX_CONSOLIDADO_COLUMNS,
    "Permisos_Usuarios": [
        "usuario", "permiso", "valor", "fecha_asignacion", "asignado_por", "estado"
    ],
    "Auditoria_Cambios": [
        "auditoria_id", "fecha_evento", "usuario", "rol", "accion", "modulo", "registro_id",
        "campo", "valor_anterior", "valor_nuevo", "motivo", "detalle"
    ],
    "Config": ["clave", "valor"],
}

# Hojas que se leen al iniciar el sistema.
# Kardex_Consolidado es una hoja calculada/sincronizada desde Movimientos,
# por eso NO debe ser requisito de lectura al arrancar. Si la pestaña no existe
# todavía en Google Sheets, el sistema debe poder iniciar y crearla al sincronizar.
CORE_SHEETS: List[str] = [
    "Productos", "Proveedores", "Solicitantes", "Personal", "Usuarios",
    "Movimientos", "Permisos_Usuarios", "Auditoria_Cambios", "Config"
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
ROLES = ["Administrador", "Supervisor", "Operador", "Consulta"]

PERMISSION_LABELS = {
    "crear_productos": "Crear productos",
    "editar_productos": "Editar productos",
    "desactivar_productos": "Desactivar/reactivar productos",
    "crear_proveedores": "Crear proveedores",
    "editar_proveedores": "Editar proveedores",
    "desactivar_proveedores": "Desactivar/reactivar proveedores",
    "crear_solicitantes": "Crear solicitantes/sitios",
    "editar_solicitantes": "Editar solicitantes/sitios",
    "desactivar_solicitantes": "Desactivar/reactivar solicitantes/sitios",
    "crear_personal": "Crear personal",
    "editar_personal": "Editar personal",
    "desactivar_personal": "Desactivar/reactivar personal",
    "registrar_movimientos": "Registrar ingresos, salidas, devoluciones y correcciones",
    "anular_movimientos": "Anular movimientos",
    "editar_movimientos": "Editar datos administrativos de movimientos",
    "exportar_reportes": "Exportar reportes",
    "ver_auditoria": "Ver auditoría",
}

CATALOG_PERMISSION_BASE = {
    "Productos": "productos",
    "Proveedores": "proveedores",
    "Solicitantes": "solicitantes",
    "Personal": "personal",
}

ROLE_DEFAULT_PERMISSIONS = {
    "Administrador": set(PERMISSION_LABELS.keys()),
    "Supervisor": {
        "crear_productos", "crear_proveedores", "crear_solicitantes", "crear_personal",
        "editar_productos", "editar_proveedores", "editar_solicitantes", "editar_personal",
        "registrar_movimientos", "anular_movimientos", "editar_movimientos",
        "exportar_reportes", "ver_auditoria",
    },
    "Operador": {
        "crear_productos", "crear_proveedores", "crear_solicitantes", "crear_personal",
        "registrar_movimientos", "exportar_reportes",
    },
    "Consulta": {"exportar_reportes"},
}

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
        {"clave": "version_sistema", "valor": "24.0"},
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
    """Guarda un mensaje para mostrarlo como toast animado tras el rerun."""
    st.session_state["flash_message"] = {"level": level, "message": message}


def show_flash() -> None:
    """Muestra el mensaje pendiente como toast animado en esquina superior derecha."""
    flash = st.session_state.pop("flash_message", None)
    if not flash:
        return
    level   = flash.get("level", "success")
    message = flash.get("message", "")
    icons   = {"success": "✅", "warning": "⚠️", "error": "❌", "info": "ℹ️"}
    icon    = icons.get(level, "ℹ️")
    _id     = f"toast_{level}_{hash(message) & 0xFFFF}"
    st.markdown(
        f"""
        <div id="{_id}" class="toast toast-{level}" style="position:fixed;">
            <span class="toast-icon">{icon}</span>
            <div style="flex:1;">{message}</div>
            <div class="toast-progress"></div>
        </div>
        <script>
        (function(){{
            var el = document.getElementById('{_id}');
            if(!el) return;
            setTimeout(function(){{
                el.style.animation = 'toastOut .4s ease forwards';
                setTimeout(function(){{ el.remove(); }}, 400);
            }}, 4200);
        }})();
        </script>
        """,
        unsafe_allow_html=True,
    )


def bump_form_nonce(name: str) -> None:
    """Cambia la clave de un formulario para limpiar textboxes tras guardar."""
    st.session_state[name] = int(st.session_state.get(name, 0)) + 1


def confirm_pending(key: str) -> bool:
    """Returns True if a confirmation is pending for the given key."""
    return st.session_state.get(f"_confirm_pending_{key}", False)


def set_confirm_pending(key: str, payload: dict | None = None) -> None:
    """Mark an action as pending confirmation with optional payload data."""
    st.session_state[f"_confirm_pending_{key}"] = True
    st.session_state[f"_confirm_payload_{key}"] = payload or {}


def clear_confirm(key: str) -> None:
    """Clear the confirmation state."""
    st.session_state.pop(f"_confirm_pending_{key}", None)
    st.session_state.pop(f"_confirm_payload_{key}", None)


def get_confirm_payload(key: str) -> dict:
    """Get the payload stored for a pending confirmation."""
    return st.session_state.get(f"_confirm_payload_{key}", {})


def render_confirm_box(
    key: str,
    title: str,
    body: str,
    confirm_label: str = "✅ Confirmar",
    cancel_label: str  = "✖ Cancelar",
    danger: bool = True,
) -> bool:
    """Show an inline confirmation card. Returns True when user confirms."""
    accent   = "#EF4444" if danger else "#38BDF8"
    icon     = "⚠️" if danger else "❓"
    confirmed = False
    st.markdown(
        f"""
        <div style="border:1.5px solid {accent}55; border-radius:18px; padding:22px 24px; margin:12px 0;
                    background:linear-gradient(160deg,rgba(15,23,42,.96),rgba(3,7,18,.92));
                    box-shadow:0 16px 50px rgba(0,0,0,.40); animation:slideUp .22s ease;">
            <div style="font-size:18px; font-weight:900; color:#F8FAFC; margin-bottom:10px;">
                {icon} {title}
            </div>
            <div style="font-size:13.5px; color:#94A3B8; line-height:1.65; margin-bottom:18px;">
                {body}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    if c1.button(confirm_label, key=f"_confirm_yes_{key}", use_container_width=True, type="primary"):
        clear_confirm(key)
        confirmed = True
    if c2.button(cancel_label, key=f"_confirm_no_{key}", use_container_width=True):
        clear_confirm(key)
        st.rerun()
    return confirmed


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
        mask = pd.Series([True] * len(df), index=df.index)
    else:
        mask = df[col].fillna("Sí").astype(str).str.lower().str.strip().isin(["sí", "si", "true", "1", "activo", ""])
    if "estado_registro" in df.columns:
        mask = mask & ~df["estado_registro"].fillna("Activo").astype(str).str.lower().str.strip().isin(["inactivo", "desactivado", "desactivada", "eliminado", "eliminada"])
    return mask


def ensure_columns(df: pd.DataFrame, sheet: str) -> pd.DataFrame:
    columns = SHEET_COLUMNS[sheet]
    if df is None or df.empty:
        df = pd.DataFrame(columns=columns)
    for col in columns:
        if col not in df.columns:
            df[col] = ""

    # Valores por defecto para mantener compatibilidad con bases creadas en versiones anteriores.
    if sheet in ["Productos", "Proveedores", "Solicitantes", "Personal"]:
        if "estado_registro" in df.columns:
            df["estado_registro"] = df["estado_registro"].replace({np.nan: "", None: ""})
            df.loc[df["estado_registro"].astype(str).str.strip().eq(""), "estado_registro"] = "Activo"
        if "activo" in df.columns:
            df["activo"] = df["activo"].replace({np.nan: "", None: ""})
            df.loc[df["activo"].astype(str).str.strip().eq(""), "activo"] = "Sí"
    if sheet == "Movimientos":
        if "estado_movimiento" in df.columns:
            df["estado_movimiento"] = df["estado_movimiento"].replace({np.nan: "", None: ""})
            df.loc[df["estado_movimiento"].astype(str).str.strip().eq(""), "estado_movimiento"] = "Vigente"
    if sheet == "Permisos_Usuarios":
        if "estado" in df.columns:
            df["estado"] = df["estado"].replace({np.nan: "", None: ""})
            df.loc[df["estado"].astype(str).str.strip().eq(""), "estado"] = "Activo"
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
    if "invalid argument" in low or "[400]" in low or "unable to parse range" in low:
        tips.append("Verifique que GOOGLE_SHEET_ID sea el ID real de la hoja o una URL válida de Google Sheets. Esta versión ya extrae el ID si pega la URL completa.")
        tips.append("Si el error menciona unable to parse range, puede faltar una pestaña en Google Sheets; active AUTO_MIGRATE_GOOGLE_SHEETS=true o cree la estructura desde Administración.")
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

        # Prepara solo las pestañas BASE con operaciones de escritura.
        # No fuerza Kardex_Consolidado al iniciar porque es calculada.
        # Esto evita detener la app cuando se agregan nuevas hojas en versiones recientes
        # y también reduce llamadas innecesarias a Google Sheets.
        auto_create = as_bool(safe_secret("AUTO_CREATE_SHEETS", True), True)
        if auto_create:
            self.ensure_sheets_write_only(CORE_SHEETS)

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

    def ensure_sheets_write_only(self, sheets: list[str] | None = None) -> None:
        """Crea pestañas faltantes y escribe encabezados sin leer metadata.

        Esta función se usa como migración automática cuando una versión nueva del
        sistema agrega hojas como Permisos_Usuarios o Auditoria_Cambios. No borra
        registros existentes: solo crea pestañas faltantes y actualiza la fila 1
        con los encabezados esperados.
        """
        target_sheets = list(sheets or SHEET_COLUMNS.keys())
        for sheet in target_sheets:
            if sheet in SHEET_COLUMNS:
                self._create_sheet_and_header(sheet)

    def ensure_database_write_only(self) -> None:
        """Crea todas las pestañas de la base sin leer metadata."""
        self.ensure_sheets_write_only(list(SHEET_COLUMNS.keys()))

    def load_many(self, sheets: list[str], _retry_migration: bool = True) -> Dict[str, pd.DataFrame]:
        """Carga varias pestañas con una sola lectura real a la API.

        Si Google devuelve error porque una pestaña agregada por una versión nueva
        todavía no existe, el sistema crea automáticamente las hojas base y vuelve
        a intentar la lectura una vez. Esto evita detener el sistema con el mensaje
        general de "falló al leer la estructura de la base".
        """
        ranges = [sheet_range_for(sheet) for sheet in sheets]
        params = []
        for rng in ranges:
            params.append(("ranges", rng))
        params.append(("majorDimension", "ROWS"))
        try:
            result = self._request("get", f"{self.base_url}/values:batchGet", params=params)
        except Exception as exc:
            detail = exception_detail(exc)
            low = detail.lower()
            should_migrate = (
                _retry_migration
                and as_bool(safe_secret("AUTO_MIGRATE_GOOGLE_SHEETS", True), True)
                and any(term in low for term in [
                    "unable to parse range",
                    "not found",
                    "bad request",
                    "[400]",
                    "requested entity was not found",
                    "cannot find grid",
                ])
            )
            if should_migrate:
                try:
                    self.ensure_sheets_write_only(sheets)
                    time.sleep(2)
                    return self.load_many(sheets, _retry_migration=False)
                except Exception as mig_exc:
                    raise RuntimeError(
                        "No se pudo leer Google Sheets y tampoco se pudo completar la migración automática de estructura. "
                        f"Detalle lectura inicial: {detail}. Detalle migración: {exception_detail(mig_exc)}. "
                        f"Revisión: {diagnose_gsheets_error(mig_exc)}"
                    ) from mig_exc
            raise RuntimeError(
                "No se pudo leer Google Sheets por lote. "
                f"Detalle API: {detail}. Revisión: {diagnose_gsheets_error(exc)}"
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



def normalize_movement_text(value) -> str:
    """Normaliza textos de tipo de movimiento para reconocer bases antiguas/manuales.

    Google Sheets puede contener valores escritos como Entrada, ENTRADA, ingreso,
    salida de insumo, ajuste salida, corrección entrada, etc. Esta función elimina
    acentos, espacios dobles y diferencias de mayúsculas para clasificar mejor.
    """
    text = clean_str(value).lower()
    text = ''.join(ch for ch in unicodedata.normalize('NFKD', text) if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text).strip()
    return text


def movimiento_sign(value) -> int:
    """Devuelve 1 si suma stock, -1 si resta stock y 0 si no clasifica.

    Se mantiene compatibilidad con las versiones anteriores que usaban Ajuste,
    y con registros escritos manualmente en Google Sheets como Entrada/Salida.
    """
    text = normalize_movement_text(value)
    if not text:
        return 0

    # Primero se revisan los casos que restan para evitar confusiones con frases largas.
    if any(token in text for token in ["salida", "egreso", "entrega", "despacho"]):
        return -1

    if "devol" in text:
        return 1

    if any(token in text for token in ["ingreso", "entrada", "recepcion", "recepción", "compra"]):
        return 1

    # Compatibilidad con ajustes/correcciones escritos de diferentes formas.
    if any(token in text for token in ["ajuste", "correccion", "corrección"]):
        if any(token in text for token in ["salida", "resta", "negativo", "disminucion", "disminución"]):
            return -1
        if any(token in text for token in ["entrada", "suma", "positivo", "aumento"]):
            return 1

    # Último respaldo con las constantes históricas exactas.
    if clean_str(value) in TIPOS_POSITIVOS:
        return 1
    if clean_str(value) in TIPOS_NEGATIVOS:
        return -1
    return 0


def stock_empty_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=[
        "producto_id", "producto", "marca", "lote", "fecha_vencimiento", "unidad",
        "ingreso_total", "salida_total", "stock_actual", "costo_total_ingresos",
        "stock_minimo", "dias_alerta_vencimiento", "dias_para_vencer", "estado"
    ])


def calcular_stock(df_mov: pd.DataFrame, df_prod: pd.DataFrame) -> pd.DataFrame:
    df = ensure_columns(df_mov, "Movimientos")
    if df.empty:
        return stock_empty_frame()

    df = df.copy()
    df = df[valid_movement_mask(df)].copy()
    if df.empty:
        return stock_empty_frame()
    df["cantidad_num"] = to_number(df["cantidad"])
    df["costo_total_num"] = to_number(df["costo_total"])
    df["tipo_movimiento"] = df["tipo_movimiento"].astype(str).str.strip()
    df["mov_sign"] = df["tipo_movimiento"].apply(movimiento_sign)
    df["entrada"] = np.where(df["mov_sign"] > 0, df["cantidad_num"], 0)
    df["salida"] = np.where(df["mov_sign"] < 0, df["cantidad_num"], 0)
    df["costo_ingreso"] = np.where(df["mov_sign"] > 0, df["costo_total_num"], 0)

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


@st.cache_data(ttl=20, show_spinner="Verificando stock directamente en Google Sheets...")
def load_operational_data_cached(_storage, refresh_token: int) -> Dict[str, pd.DataFrame]:
    """Carga hojas operativas con caché corta para tomar lo que existe en Google Sheets.

    Esta lectura se usa especialmente en Movimientos/Salidas para que el carrito
    calcule el stock con la base real y no únicamente con datos que pudieron quedar
    cacheados al iniciar la app.
    """
    sheets = ["Productos", "Proveedores", "Solicitantes", "Personal", "Movimientos"]
    if hasattr(_storage, "load_many"):
        loaded = _storage.load_many(sheets)
    else:
        loaded = {sheet: _storage.load(sheet) for sheet in sheets}

    # Kardex_Consolidado es diagnóstico; si no existe no debe romper la operación.
    try:
        loaded["Kardex_Consolidado"] = _storage.load("Kardex_Consolidado")
    except Exception:
        loaded["Kardex_Consolidado"] = pd.DataFrame(columns=SHEET_COLUMNS["Kardex_Consolidado"])
    return loaded


def calcular_stock_desde_kardex_consolidado(df_kardex: pd.DataFrame, df_prod: pd.DataFrame) -> pd.DataFrame:
    """Convierte la hoja física Kardex_Consolidado a formato de stock para diagnóstico.

    No reemplaza la fuente oficial de transacciones. La fuente oficial sigue siendo
    Movimientos; esta conversión permite detectar si existe una hoja consolidada con
    saldos que no están respaldados por movimientos válidos.
    """
    kd = ensure_columns(df_kardex, "Kardex_Consolidado")
    if kd.empty:
        return stock_empty_frame()
    out = pd.DataFrame({
        "producto_id": kd["producto_id"],
        "producto": kd["producto"],
        "marca": kd["marca"],
        "lote": kd["lote"],
        "fecha_vencimiento": kd["fecha_vencimiento"],
        "unidad": kd["unidad"],
        "ingreso_total": to_number(kd["entrada_total"]),
        "salida_total": to_number(kd["salida_total"]),
        "stock_actual": to_number(kd["saldo_actual"]),
        "costo_total_ingresos": 0,
    })
    prod = ensure_columns(df_prod, "Productos")[["producto_id", "stock_minimo", "dias_alerta_vencimiento"]].copy()
    prod["stock_minimo"] = to_number(prod["stock_minimo"])
    prod["dias_alerta_vencimiento"] = to_number(prod["dias_alerta_vencimiento"], 90)
    out = out.merge(prod, on="producto_id", how="left")
    out["stock_minimo"] = to_number(out["stock_minimo"])
    out["dias_alerta_vencimiento"] = to_number(out["dias_alerta_vencimiento"], 90)
    venc = to_date(out["fecha_vencimiento"])
    out["dias_para_vencer"] = (venc - TODAY).dt.days
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
    out["estado"] = out.apply(estado, axis=1)
    return out.sort_values(["estado", "producto", "fecha_vencimiento", "lote"], na_position="last")


def refresh_operational_data_and_stock(storage, data: Dict[str, pd.DataFrame], current_stock: pd.DataFrame) -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame, pd.DataFrame]:
    """Refresca Movimientos y catálogos desde la base real para salidas/correcciones.

    Devuelve:
    - data actualizado
    - stock calculado desde Movimientos, que es la fuente oficial
    - stock calculado desde Kardex_Consolidado solo para diagnóstico
    """
    try:
        refresh_token = st.session_state.get("data_refresh_token", 0)
        live = load_operational_data_cached(storage, refresh_token)
        for sheet, df in live.items():
            if sheet in SHEET_COLUMNS:
                data[sheet] = ensure_columns(df, sheet)
        stock_mov = calcular_stock(data.get("Movimientos", pd.DataFrame()), data.get("Productos", pd.DataFrame()))
        stock_kardex = calcular_stock_desde_kardex_consolidado(data.get("Kardex_Consolidado", pd.DataFrame()), data.get("Productos", pd.DataFrame()))
        return data, stock_mov, stock_kardex
    except Exception as exc:
        st.warning(
            "No se pudo refrescar la información operativa directamente desde la base. "
            "Se usará la información ya cargada en memoria. Detalle: " + exception_detail(exc)
        )
        return data, current_stock, stock_empty_frame()


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
    df = df[valid_movement_mask(df)].copy()
    if df.empty:
        return pd.DataFrame(columns=columns)
    df["cantidad_num"] = to_number(df["cantidad"])
    df["fecha_dt"] = to_date(df["fecha"])
    df["fecha_registro_dt"] = to_date(df["fecha_registro"])
    df["tipo_movimiento"] = df["tipo_movimiento"].astype(str).str.strip()
    df["mov_sign"] = df["tipo_movimiento"].apply(movimiento_sign)
    df["entrada"] = np.where(df["mov_sign"] > 0, df["cantidad_num"], 0)
    df["salida"] = np.where(df["mov_sign"] < 0, df["cantidad_num"], 0)

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
        if "Permisos_Usuarios" in data:
            data["Permisos_Usuarios"].to_excel(writer, index=False, sheet_name="Permisos_Usuarios")
        if "Auditoria_Cambios" in data:
            data["Auditoria_Cambios"].to_excel(writer, index=False, sheet_name="Auditoria_Cambios")
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
        /* ── Variables ──────────────────────────────────────────────── */
        :root{
            --bg:#050816; --panel:#0B1020; --panel2:#0F172A; --line:rgba(148,163,184,.18);
            --text:#E5E7EB; --muted:#94A3B8; --dim:#64748B;
            --cyan:#38BDF8; --green:#22C55E; --orange:#F97316; --red:#EF4444;
            --purple:#8B5CF6; --blue:#3B82F6; --yellow:#EAB308;
            --radius-sm:10px; --radius-md:16px; --radius-lg:22px; --radius-xl:28px;
            --shadow-sm:0 4px 14px rgba(0,0,0,.20);
            --shadow-md:0 12px 32px rgba(0,0,0,.28);
            --shadow-lg:0 20px 50px rgba(0,0,0,.35);
        }
        /* ── Base ───────────────────────────────────────────────────── */
        .stApp{
            background: radial-gradient(ellipse 80% 50% at 10% 0%, rgba(56,189,248,.10) 0%, transparent 50%),
                        radial-gradient(ellipse 60% 40% at 90% 0%, rgba(139,92,246,.08) 0%, transparent 50%),
                        linear-gradient(180deg, #030712 0%, #0A0F1E 60%, #030712 100%);
        }
        .block-container{padding-top:.75rem; padding-bottom:2.5rem; max-width:1500px;}
        /* ── Sidebar ────────────────────────────────────────────────── */
        section[data-testid="stSidebar"]{
            background:linear-gradient(180deg, rgba(3,7,18,.98), rgba(10,15,30,.98));
            border-right:1px solid var(--line);
        }
        section[data-testid="stSidebar"] .stRadio label{
            border-radius:var(--radius-sm); padding:8px 10px; transition:background .15s; display:block;
        }
        /* ── Hero banner ────────────────────────────────────────────── */
        .hero{
            background: radial-gradient(ellipse 70% 80% at 0% 50%, rgba(56,189,248,.22) 0%, transparent 55%),
                        linear-gradient(135deg, #07111F 0%, #0D1B2E 50%, #0F172A 100%);
            border:1px solid rgba(56,189,248,.18); border-radius:var(--radius-xl);
            padding:22px 28px; margin-bottom:20px; box-shadow:var(--shadow-md);
            position:relative; overflow:hidden;
        }
        .hero::after{
            content:""; position:absolute; top:-40px; right:-40px;
            width:200px; height:200px; border-radius:50%;
            background:radial-gradient(circle, rgba(139,92,246,.12), transparent 70%);
        }
        .hero h1{font-size:32px; margin:0; color:#F8FAFC; letter-spacing:.2px; font-weight:950;}
        .hero p{margin:6px 0 0 0; color:var(--muted); font-size:13.5px; line-height:1.5;}
        /* ── Section headers ────────────────────────────────────────── */
        .section-title{
            font-size:22px; font-weight:900; color:#F8FAFC; margin:4px 0 2px 0;
            display:flex; align-items:center; gap:8px;
        }
        .section-subtitle{font-size:13px; color:#64748B; margin-bottom:18px; line-height:1.5;}
        .section-divider{border:none; border-top:1px solid var(--line); margin:16px 0;}
        /* ── Cards ──────────────────────────────────────────────────── */
        .card{
            border:1px solid var(--line); border-radius:var(--radius-lg);
            padding:20px 22px; margin:0 0 16px 0;
            background:linear-gradient(170deg, rgba(15,23,42,.88), rgba(5,10,24,.82));
            box-shadow:var(--shadow-sm); transition:border-color .2s;
        }
        .card:hover{border-color:rgba(148,163,184,.30);}
        .card-title{
            font-size:15px; font-weight:800; color:#F8FAFC; margin-bottom:4px;
            display:flex; align-items:center; gap:7px;
        }
        .card-sub{font-size:12px; color:var(--muted); margin-bottom:14px; line-height:1.4;}
        /* ── Legacy form-card kept for compatibility ─────────────────── */
        .form-card{
            border:1px solid var(--line); border-radius:var(--radius-lg); padding:20px 22px;
            margin:8px 0 18px 0; background:linear-gradient(170deg, rgba(15,23,42,.88), rgba(5,10,24,.82));
            box-shadow:var(--shadow-sm);
        }
        /* ── Step badges ────────────────────────────────────────────── */
        .step-badge{
            display:inline-flex; align-items:center; justify-content:center;
            width:26px; height:26px; border-radius:50%; background:rgba(56,189,248,.18);
            border:1.5px solid rgba(56,189,248,.45); color:#38BDF8;
            font-size:12px; font-weight:900; margin-right:8px; flex-shrink:0;
        }
        .step-title{font-size:14px; font-weight:800; color:#E2E8F0; display:flex; align-items:center;}
        .step-divider{border:none; border-top:1px solid rgba(148,163,184,.12); margin:14px 0 12px 0;}
        /* ── KPI cards ──────────────────────────────────────────────── */
        .kpi-card{
            border:1px solid var(--line); border-radius:var(--radius-md); padding:16px 18px;
            background:linear-gradient(160deg, rgba(15,23,42,.95), rgba(3,7,18,.90));
            box-shadow:var(--shadow-sm); position:relative; overflow:hidden; min-height:110px;
        }
        .kpi-accent{
            position:absolute; top:0; right:0; width:50px; height:50px; border-radius:0 var(--radius-md) 0 50px;
            opacity:.15;
        }
        .kpi-label{color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.10em; font-weight:600;}
        .kpi-value{color:#F8FAFC; font-size:30px; font-weight:950; margin:8px 0 4px 0; line-height:1;}
        .kpi-note{color:#CBD5E1; font-size:11.5px; line-height:1.4;}
        .kpi-trend{font-size:11px; font-weight:700; margin-top:5px;}
        /* ── Mini info card ─────────────────────────────────────────── */
        .mini-card{
            border:1px solid var(--line); border-radius:var(--radius-md); padding:13px 15px;
            margin-bottom:12px; background:rgba(15,23,42,.60);
        }
        .mini-title{font-weight:800; color:#F8FAFC; font-size:14.5px; margin-bottom:2px;}
        .mini-sub{color:var(--muted); font-size:12px;}
        /* ── Alerts ─────────────────────────────────────────────────── */
        .alert{border-radius:var(--radius-sm); padding:12px 16px; margin:8px 0; font-size:13px; line-height:1.5;}
        .alert-red{border-left:3px solid var(--red); background:rgba(239,68,68,.10); color:#FEE2E2;}
        .alert-orange{border-left:3px solid var(--orange); background:rgba(249,115,22,.10); color:#FFEDD5;}
        .alert-green{border-left:3px solid var(--green); background:rgba(34,197,94,.10); color:#DCFCE7;}
        .alert-blue{border-left:3px solid var(--cyan); background:rgba(56,189,248,.10); color:#BAE6FD;}
        .alert-purple{border-left:3px solid var(--purple); background:rgba(139,92,246,.10); color:#EDE9FE;}
        /* ── Confirm modal overlay ───────────────────────────────────── */
        .confirm-overlay{
            position:fixed; inset:0; background:rgba(0,0,0,.65); z-index:9998;
            display:flex; align-items:center; justify-content:center;
            backdrop-filter:blur(4px); animation:fadeIn .18s ease;
        }
        .confirm-box{
            background:linear-gradient(160deg, #0F172A, #070E1C);
            border:1px solid rgba(148,163,184,.25); border-radius:var(--radius-xl);
            padding:28px 32px; max-width:480px; width:90%;
            box-shadow:0 30px 80px rgba(0,0,0,.55); animation:slideUp .22s ease;
        }
        .confirm-title{font-size:20px; font-weight:900; color:#F8FAFC; margin-bottom:8px;}
        .confirm-body{font-size:13.5px; color:#94A3B8; line-height:1.6; margin-bottom:20px;}
        /* ── Toast ──────────────────────────────────────────────────── */
        .toast{
            position:fixed; top:18px; right:18px; z-index:9999;
            max-width:400px; min-width:280px;
            border-radius:var(--radius-md); padding:14px 18px;
            box-shadow:0 12px 40px rgba(0,0,0,.45);
            display:flex; align-items:flex-start; gap:12px;
            animation:toastIn .3s cubic-bezier(.22,1,.36,1);
            font-size:14px; font-weight:500; line-height:1.5;
        }
        .toast-icon{font-size:20px; flex-shrink:0; margin-top:1px;}
        .toast-success{background:linear-gradient(135deg,#052e16,#052e16); border:1px solid #166534; color:#dcfce7;}
        .toast-warning{background:linear-gradient(135deg,#1c1009,#1c1009); border:1px solid #92400e; color:#fef3c7;}
        .toast-error  {background:linear-gradient(135deg,#1a0505,#1a0505); border:1px solid #7f1d1d; color:#fee2e2;}
        .toast-info   {background:linear-gradient(135deg,#071928,#071928); border:1px solid #1e3a5f; color:#bae6fd;}
        .toast-progress{
            position:absolute; bottom:0; left:0; height:3px; border-radius:0 0 var(--radius-md) var(--radius-md);
            animation:toastBar 4s linear forwards;
        }
        .toast-success .toast-progress{background:#22c55e;}
        .toast-warning .toast-progress{background:#f97316;}
        .toast-error   .toast-progress{background:#ef4444;}
        .toast-info    .toast-progress{background:#38bdf8;}
        /* ── Badges/pills ───────────────────────────────────────────── */
        .pill{display:inline-block; padding:4px 10px; border-radius:999px;
              background:rgba(56,189,248,.12); color:#BAE6FD;
              border:1px solid rgba(56,189,248,.22); font-size:11.5px; margin-right:5px; font-weight:600;}
        .badge{display:inline-block; padding:2px 8px; border-radius:6px; font-size:11px; font-weight:700;}
        .badge-green{background:rgba(34,197,94,.18); color:#4ade80; border:1px solid rgba(34,197,94,.30);}
        .badge-red  {background:rgba(239,68,68,.18);  color:#f87171; border:1px solid rgba(239,68,68,.30);}
        .badge-orange{background:rgba(249,115,22,.18);color:#fb923c; border:1px solid rgba(249,115,22,.30);}
        .badge-blue {background:rgba(56,189,248,.18); color:#38bdf8; border:1px solid rgba(56,189,248,.30);}
        .badge-purple{background:rgba(139,92,246,.18);color:#a78bfa; border:1px solid rgba(139,92,246,.30);}
        /* ── Login ──────────────────────────────────────────────────── */
        .login-wrap{max-width:520px; margin:8vh auto 0 auto;}
        .login-card{
            border:1px solid rgba(148,163,184,.22); border-radius:var(--radius-xl);
            padding:32px; background:linear-gradient(170deg, rgba(15,23,42,.95), rgba(3,7,18,.90));
            box-shadow:var(--shadow-lg);
        }
        .login-title{font-size:28px; font-weight:950; color:#F8FAFC; margin:0 0 4px 0;}
        .login-sub{font-size:13.5px; color:var(--muted); margin:0 0 20px 0;}
        /* ── Field notes ────────────────────────────────────────────── */
        .field-note{font-size:11.5px; color:var(--muted); margin-top:-6px; margin-bottom:8px;}
        /* ── Buttons ────────────────────────────────────────────────── */
        div[data-testid="stMetricValue"]{font-size:26px;}
        .stButton>button{
            border-radius:var(--radius-sm); font-weight:700; font-size:13.5px;
            transition:transform .12s, box-shadow .12s;
        }
        .stButton>button:hover{transform:translateY(-1px); box-shadow:0 6px 20px rgba(0,0,0,.30);}
        .stDownloadButton>button{border-radius:var(--radius-sm); font-weight:700;}
        /* ── Animations ─────────────────────────────────────────────── */
        @keyframes fadeIn{from{opacity:0}to{opacity:1}}
        @keyframes slideUp{from{transform:translateY(24px);opacity:0}to{transform:translateY(0);opacity:1}}
        @keyframes toastIn{from{transform:translateX(110%);opacity:0}to{transform:translateX(0);opacity:1}}
        @keyframes toastOut{to{transform:translateX(110%);opacity:0}}
        @keyframes toastBar{from{width:100%}to{width:0%}}
        /* ── Kardex table colors ────────────────────────────────────── */
        .kdx-row-ok  {background:rgba(34,197,94,.06);}
        .kdx-row-warn{background:rgba(249,115,22,.06);}
        .kdx-row-crit{background:rgba(239,68,68,.06);}
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
            <div style="margin-top:12px"><span class='pill'>Base activa: {storage_mode}</span>{user_badge}<span class='pill'>Versión 24.0</span></div>
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


def kpi_card(label: str, value, note: str = "", color: str = "#38BDF8", trend: str = "") -> None:
    trend_html = f'<div class="kpi-trend" style="color:{color};">{trend}</div>' if trend else ""
    st.markdown(
        f"""
        <div class="kpi-card" style="border-color:{color}22; border-top:2px solid {color}55;">
            <div class="kpi-accent" style="background:{color};"></div>
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="color:{color};">{value}</div>
            <div class="kpi-note">{note}</div>
            {trend_html}
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


def current_username() -> str:
    return clean_str(st.session_state.get("usuario", "Sistema")) or "Sistema"


def current_user_display() -> str:
    return clean_str(st.session_state.get("nombre_usuario", current_username())) or current_username()


def current_role() -> str:
    return clean_str(st.session_state.get("rol", ""))


def is_admin() -> bool:
    return current_role() == "Administrador"


def permission_value_is_yes(value) -> bool:
    return clean_str(value).lower() in {"sí", "si", "true", "1", "activo", "yes"}


def user_has_permission(data: Dict[str, pd.DataFrame], permission: str) -> bool:
    """Permisos efectivos por rol + excepciones asignadas por el administrador.

    - Administrador siempre tiene todos los permisos.
    - Supervisor/Operador/Consulta reciben permisos base.
    - La hoja Permisos_Usuarios permite activar o negar permisos específicos por usuario.
    """
    role = current_role()
    user = current_username().lower().strip()
    if role == "Administrador":
        return True

    allowed = permission in ROLE_DEFAULT_PERMISSIONS.get(role, set())
    permisos = ensure_columns(data.get("Permisos_Usuarios", pd.DataFrame()), "Permisos_Usuarios")
    if not permisos.empty and user:
        p = permisos[
            (permisos["usuario"].astype(str).str.lower().str.strip() == user)
            & (permisos["permiso"].astype(str).str.strip() == permission)
            & (permisos["estado"].astype(str).str.lower().str.strip().isin(["activo", "sí", "si", "true", "1", ""]))
        ]
        if not p.empty:
            allowed = permission_value_is_yes(p.iloc[-1].get("valor", "No"))
    return bool(allowed)


def require_permission(data: Dict[str, pd.DataFrame], permission: str, action_text: str = "realizar esta acción") -> bool:
    if user_has_permission(data, permission):
        return True
    label = PERMISSION_LABELS.get(permission, permission)
    st.warning(
        f"No tiene permiso para {action_text}. Permiso requerido: {label}. "
        "Solicite al administrador que le habilite este permiso si corresponde."
    )
    return False


def append_audit(storage, accion: str, modulo: str, registro_id: str = "", campo: str = "", valor_anterior: str = "", valor_nuevo: str = "", motivo: str = "", detalle: str = "") -> None:
    """Registra auditoría sin detener la operación si falla la escritura."""
    try:
        row = {
            "auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}",
            "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            "usuario": current_username(),
            "rol": current_role(),
            "accion": accion,
            "modulo": modulo,
            "registro_id": clean_str(registro_id),
            "campo": clean_str(campo),
            "valor_anterior": clean_str(valor_anterior),
            "valor_nuevo": clean_str(valor_nuevo),
            "motivo": clean_str(motivo),
            "detalle": clean_str(detalle),
        }
        storage.append_row("Auditoria_Cambios", row)
    except Exception:
        # La auditoría no debe romper el guardado principal.
        pass


def append_audit_rows(storage, rows: list[dict]) -> None:
    if not rows:
        return
    try:
        storage.append_rows("Auditoria_Cambios", rows)
    except Exception:
        pass


def catalog_id_col(sheet: str) -> str:
    return {
        "Productos": "producto_id",
        "Proveedores": "proveedor_id",
        "Solicitantes": "solicitante_id",
        "Personal": "personal_id",
    }.get(sheet, "")


def catalog_display_col(sheet: str) -> str:
    return {
        "Productos": "nombre_producto",
        "Proveedores": "proveedor",
        "Solicitantes": "unidad_solicitante",
        "Personal": "nombre",
    }.get(sheet, "")


def catalog_permission(sheet: str, action: str) -> str:
    base = CATALOG_PERMISSION_BASE.get(sheet, sheet.lower())
    return f"{action}_{base}"


def valid_movement_mask(df: pd.DataFrame) -> pd.Series:
    if df.empty:
        return pd.Series([], dtype=bool, index=df.index)
    if "estado_movimiento" not in df.columns:
        return pd.Series([True] * len(df), index=df.index)
    return ~df["estado_movimiento"].astype(str).str.lower().str.strip().isin(["anulado", "anulada", "cancelado", "cancelada"])


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



def resolve_acta_logo_path() -> Path | None:
    """Devuelve una ruta local válida para el logo PNG del acta.

    En Streamlit Cloud algunas rutas pueden no existir si el archivo no fue subido
    al repositorio. Para evitar que el PDF use el logo textual de respaldo, el logo
    oficial también queda embebido en base64 dentro del código y se reconstruye
    automáticamente en /tmp si hace falta.
    """
    candidates = [
        ACTA_LOGO_PATH,
        APP_DIR / "logo_vihca.png",
        APP_DIR / "logo.png",
        ASSETS_DIR / "logo.png",
    ]
    for candidate in candidates:
        try:
            if candidate.exists() and candidate.stat().st_size > 0:
                return candidate
        except Exception:
            pass

    try:
        tmp_logo = Path("/tmp/kardex_logo_vihca.png")
        if not tmp_logo.exists() or tmp_logo.stat().st_size == 0:
            tmp_logo.write_bytes(base64.b64decode(ACTA_LOGO_B64))
        return tmp_logo
    except Exception:
        return None


def get_first_match(df: pd.DataFrame, col: str, value: str) -> dict:
    if df is None or df.empty or col not in df.columns:
        return {}
    m = df[df[col].astype(str) == str(value)]
    if m.empty:
        return {}
    return {k: clean_str(v) for k, v in m.iloc[0].to_dict().items()}


def join_spanish(items: list[str]) -> str:
    """Une una lista corta en español: A, B y C."""
    items = [clean_str(x) for x in items if clean_str(x)]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} y {items[1]}"
    return f"{', '.join(items[:-1])} y {items[-1]}"


def categoria_acta_plural(categoria: str) -> str:
    """Convierte la categoría del catálogo en el texto que se imprime en el acta."""
    cat = clean_str(categoria).strip().lower()
    if not cat:
        return "productos"

    # Normalización flexible por si el catálogo tiene variantes escritas por usuarios.
    if "react" in cat or "prueba" in cat or "test" in cat:
        return "reactivos"
    if "insumo" in cat:
        return "insumos"
    if "equipo" in cat:
        return "equipos"
    if "material" in cat:
        return "materiales"
    if "papeler" in cat:
        return "materiales de papelería"
    if "medic" in cat or "fárm" in cat or "farm" in cat:
        return "medicamentos"
    if "control" in cat:
        return "controles"
    if "calibr" in cat:
        return "calibradores"
    if "consum" in cat:
        return "consumibles"

    # Fallback conservador: si ya viene en plural, se respeta; si no, se agrega "s".
    if cat.endswith("s"):
        return cat
    return f"{cat}s"


def acta_tipo_entrega_desde_categorias(rows: list[dict]) -> str:
    """Determina el texto: reactivos, insumos, equipos, etc., según categoría de los productos."""
    categorias = []
    for r in rows or []:
        categoria = clean_str(r.get("categoria", ""))
        plural = categoria_acta_plural(categoria)
        if plural and plural not in categorias:
            categorias.append(plural)

    if not categorias:
        return "productos"

    # Si hay demasiadas categorías, se usa una frase compacta para no saturar la redacción.
    if len(categorias) > 4:
        principales = categorias[:3]
        return f"{join_spanish(principales)} y otros productos"

    return join_spanish(categorias)


def build_acta_entrega_pdf(
    salida_rows: list[dict],
    solicitante_info: dict,
    personal_info: dict,
    fecha_entrega,
    recibe_nombre: str = "",
    recibe_cargo: str = "",
    ciudad: str = "Tegucigalpa",
    observacion: str = "",
    tipo_entrega_texto: str = "",
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
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
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
    tipo_entrega = clean_str(tipo_entrega_texto) or acta_tipo_entrega_desde_categorias(salida_rows)

    story = []
    # Logo institucional desde PNG real.
    # Primero intenta assets/logo_vihca.png; si no existe en Streamlit Cloud, reconstruye
    # automáticamente el PNG desde base64 para evitar el encabezado textual deformado.
    logo_path = resolve_acta_logo_path()
    if logo_path:
        try:
            logo = Image(str(logo_path), width=1.95 * inch, height=0.73 * inch, hAlign="CENTER")
            story.append(logo)
        except Exception:
            story.append(Paragraph("<b>PROGRAMA REGIONAL DE VIH</b>", normal_center))
    else:
        story.append(Paragraph("<b>PROGRAMA REGIONAL DE VIH</b>", normal_center))
    story.append(Spacer(1, 0.16 * inch))
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
        f"Programa de VIH brinda en Honduras, se hace entrega formal de los siguientes {tipo_entrega} "
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
        [Paragraph("Programa VIHCA<br/>Asociado al Centro de Estudios en Salud, de la<br/>Universidad del Valle de Guatemala", small_center), Paragraph(sitio, small_center)],
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
        "📋 Movimientos de inventario",
        "Registre ingresos, salidas y devoluciones. Si cometió un error en un movimiento ya guardado, use la opción Editar movimiento."
    )

    if not require_permission(data, "registrar_movimientos", "registrar movimientos"):
        return

    # V20: antes de registrar salidas/correcciones se verifica la base real.
    # En Google Sheets se hace una lectura por lote con caché corta para evitar cuota,
    # pero evitando que el carrito use datos viejos de una carga inicial.
    data, stock, stock_kardex_sheet = refresh_operational_data_and_stock(storage, data, stock)

    productos = ensure_columns(data["Productos"], "Productos")
    proveedores = ensure_columns(data["Proveedores"], "Proveedores")
    solicitantes = ensure_columns(data["Solicitantes"], "Solicitantes")
    personal_df = ensure_columns(data["Personal"], "Personal")
    movimientos = ensure_columns(data["Movimientos"], "Movimientos")

    productos = productos[active_mask(productos)].copy()
    proveedores = proveedores[active_mask(proveedores)].copy()
    solicitantes = solicitantes[active_mask(solicitantes)].copy()
    personal_df = personal_df[active_mask(personal_df)].copy()

    if storage.__class__.__name__ == "GoogleSheetsStorage":
        st.caption("Stock operativo verificado contra Google Sheets: hoja Movimientos + catálogo Productos. La hoja Kardex_Consolidado se usa como respaldo visual/diagnóstico, no como fuente transaccional principal.")

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
            audit_rows = []
            for r in rows:
                audit_rows.append({
                    "auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}",
                    "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "usuario": current_username(),
                    "rol": current_role(),
                    "accion": "CREAR_MOVIMIENTO",
                    "modulo": "Movimientos",
                    "registro_id": clean_str(r.get("movimiento_id", "")),
                    "campo": "",
                    "valor_anterior": "",
                    "valor_nuevo": clean_str(r.get("tipo_movimiento", "")),
                    "motivo": "",
                    "detalle": f"{clean_str(r.get('producto', ''))} | Lote: {clean_str(r.get('lote', ''))} | Cantidad: {clean_str(r.get('cantidad', ''))}",
                })
            append_audit_rows(storage, audit_rows)
            sync_kardex_consolidado_sheet(storage, sync_data)
            set_flash(message)
        except Exception as exc:
            set_flash(f"Registros guardados. No se pudo actualizar Kardex_Consolidado o auditoría: {exc}", "warning")
        bump_form_nonce("mov_form_nonce")
        mark_data_dirty()
        rerun()

    # ── Guía rápida de opciones ───────────────────────────────────────────
    can_manage = (user_has_permission(data, "anular_movimientos")
                  or user_has_permission(data, "editar_movimientos"))

    st.markdown(
        """
        <div style="display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin:0 0 18px 0;">
            <div style="border:1px solid rgba(56,189,248,.25); border-radius:14px; padding:12px 14px; background:rgba(56,189,248,.06);">
                <div style="font-size:15px; font-weight:800; color:#38BDF8; margin-bottom:4px;">📥 Ingreso</div>
                <div style="font-size:12px; color:#94A3B8;">Productos que llegan al inventario desde un proveedor.</div>
            </div>
            <div style="border:1px solid rgba(249,115,22,.25); border-radius:14px; padding:12px 14px; background:rgba(249,115,22,.06);">
                <div style="font-size:15px; font-weight:800; color:#F97316; margin-bottom:4px;">📤 Salida</div>
                <div style="font-size:12px; color:#94A3B8;">Productos que salen del inventario hacia un sitio o unidad.</div>
            </div>
            <div style="border:1px solid rgba(34,197,94,.25); border-radius:14px; padding:12px 14px; background:rgba(34,197,94,.06);">
                <div style="font-size:15px; font-weight:800; color:#22C55E; margin-bottom:4px;">↩️ Devolución</div>
                <div style="font-size:12px; color:#94A3B8;">Un sitio retorna productos que le habían sido entregados.</div>
            </div>
            <div style="border:1px solid rgba(139,92,246,.25); border-radius:14px; padding:12px 14px; background:rgba(139,92,246,.06);">
                <div style="font-size:15px; font-weight:800; color:#8B5CF6; margin-bottom:4px;">🔧 Ajuste de inventario</div>
                <div style="font-size:12px; color:#94A3B8;">Agrega o descuenta unidades por diferencia de conteo físico. <b style="color:#E5E7EB;">Crea un nuevo registro</b>, no modifica el anterior.</div>
            </div>
            <div style="border:1px solid rgba(239,68,68,.25); border-radius:14px; padding:12px 14px; background:rgba(239,68,68,.06); grid-column:span 2;">
                <div style="font-size:15px; font-weight:800; color:#EF4444; margin-bottom:4px;">✏️ Editar movimiento existente</div>
                <div style="font-size:12px; color:#94A3B8;">
                    ¿Se equivocó en la cantidad, el lote o el sitio de un movimiento ya guardado?
                    <b style="color:#F8FAFC;">Esta opción modifica directamente ese registro</b>
                    y el stock se recalcula automáticamente.
                    <br><span style="color:#EF4444;">→ Si registró 110 pero eran 80, aquí puede corregirlo y quedará como 80.</span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _OPCIONES_LABEL = {
        "Ingreso":            "📥 Ingreso — Nueva entrada al inventario",
        "Salida":             "📤 Salida — Entrega a un sitio o unidad",
        "Devolución":         "↩️ Devolución — Retorno de productos entregados",
        "Corrección entrada": "🔧 Ajuste (+) — Agregar unidades por conteo físico",
        "Corrección salida":  "🔧 Ajuste (–) — Descontar unidades por conteo físico",
        "Editar movimiento":  "✏️ Editar movimiento — Corregir un registro ya guardado",
    }

    opciones_movimiento = list(TIPOS_MOVIMIENTO)
    if can_manage:
        opciones_movimiento.append("Editar movimiento")

    tipo = st.radio(
        "¿Qué desea hacer?",
        opciones_movimiento,
        format_func=lambda x: _OPCIONES_LABEL.get(x, x),
        horizontal=False,
        help="Seleccione la operación. Si cometió un error en un registro existente, elija 'Editar movimiento'.",
    )

    if tipo == "Editar movimiento":
        render_movement_crud_controls(storage, data)
        return

    # ========================================================
    # INGRESO
    # ========================================================
    if tipo == "Ingreso":
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        st.markdown(
            "<div class='step-title'><span class='step-badge'>1</span>Producto del catálogo</div>",
            unsafe_allow_html=True
        )
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
            st.markdown(
                "<div class='step-title'><span class='step-badge'>2</span>Datos del ingreso</div><hr class='step-divider'>",
                unsafe_allow_html=True
            )
            c1, c2, c3 = st.columns(3)
            fecha = c1.date_input("Fecha del ingreso", value=TODAY.date())
            usuario = c2.text_input("Usuario que registra", value=st.session_state.get("nombre_usuario", "Usuario"))
            cantidad = c3.number_input("Cantidad", min_value=0.0, step=1.0, format="%.2f")

            st.markdown(
                "<div class='step-title'><span class='step-badge'>3</span>Datos del lote</div><hr class='step-divider'>",
                unsafe_allow_html=True
            )
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

            st.markdown(
                "<div class='step-title'><span class='step-badge'>4</span>Proveedor y logística</div><hr class='step-divider'>",
                unsafe_allow_html=True
            )
            g, h, i = st.columns(3)
            proveedor = g.selectbox("Proveedor *", [""] + proveedores["proveedor"].dropna().astype(str).sort_values().tolist())
            orden_compra = h.text_input("Orden de compra", placeholder="Ejemplo: OC-2026-001")
            personal = i.selectbox("Personal que recibe", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            observacion = st.text_area("Observación", placeholder="Detalle de factura, donación, compra u otra referencia.")
            submitted = st.form_submit_button("💾 Guardar ingreso →", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if not require_permission(data, "registrar_movimientos", "registrar movimientos"):
                return
            if cantidad <= 0:
                st.error("⚠️ La cantidad debe ser mayor a cero.")
                return
            if not lote:
                st.error("⚠️ El lote es obligatorio.")
                return
            if not proveedor:
                st.error("⚠️ Para ingresos debe seleccionar proveedor.")
                return
            set_confirm_pending(f"guardar_ingreso_{lote}_{producto_id}", {
                "producto": producto, "lote": lote, "cantidad": cantidad,
                "proveedor": proveedor, "unidad": unidad,
            })

        confirm_key_ing = f"guardar_ingreso_{lote}_{producto_id}" if submitted or confirm_pending(f"guardar_ingreso_{lote}_{producto_id}") else ""
        if confirm_key_ing and confirm_pending(confirm_key_ing):
            _pay = get_confirm_payload(confirm_key_ing)
            confirmed_ing = render_confirm_box(
                key=confirm_key_ing,
                title="¿Confirmar ingreso al inventario?",
                body=(
                    f"<b>Producto:</b> {_pay.get('producto','')}<br>"
                    f"<b>Lote:</b> {_pay.get('lote','')}<br>"
                    f"<b>Cantidad:</b> {_pay.get('cantidad',0):g} {_pay.get('unidad','')}<br>"
                    f"<b>Proveedor:</b> {_pay.get('proveedor','')}"
                ),
                confirm_label="✅ Confirmar ingreso",
                cancel_label="↩ Cancelar",
                danger=False,
            )
            if not confirmed_ing:
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
            productos_count = len(productos)
            movimientos_count = len(movimientos)
            tipos_mov = movimientos["tipo_movimiento"].astype(str).str.strip() if not movimientos.empty else pd.Series(dtype=str)
            signos_mov = tipos_mov.apply(movimiento_sign) if not movimientos.empty else pd.Series(dtype=int)
            ingresos_count = int((signos_mov > 0).sum()) if not movimientos.empty else 0
            salidas_count = int((signos_mov < 0).sum()) if not movimientos.empty else 0
            stock_rows = len(stock) if not stock.empty else 0
            kardex_sheet_rows = len(stock_kardex_sheet) if 'stock_kardex_sheet' in locals() and not stock_kardex_sheet.empty else 0
            kardex_sheet_positive = int((to_number(stock_kardex_sheet["stock_actual"]) > 0).sum()) if 'stock_kardex_sheet' in locals() and not stock_kardex_sheet.empty else 0

            st.error("No hay lotes con stock disponible para registrar salidas.")
            st.warning(
                "Tener productos registrados en el catálogo no significa que exista stock disponible. "
                "La salida solo se habilita cuando el sistema encuentra movimientos de ingreso, devolución o corrección de entrada con saldo mayor a cero."
            )
            diag_df = pd.DataFrame([
                {"Validación": "Productos en catálogo", "Cantidad": productos_count, "Interpretación": "Son datos maestros; no generan existencia por sí solos."},
                {"Validación": "Movimientos registrados", "Cantidad": movimientos_count, "Interpretación": "Bitácora transaccional usada para calcular stock."},
                {"Validación": "Movimientos que suman stock", "Cantidad": ingresos_count, "Interpretación": "Ingreso, Devolución o Corrección entrada."},
                {"Validación": "Movimientos que restan stock", "Cantidad": salidas_count, "Interpretación": "Salida o Corrección salida."},
                {"Validación": "Lotes calculados en stock", "Cantidad": stock_rows, "Interpretación": "Filas generadas desde movimientos por producto/lote."},
                {"Validación": "Lotes con saldo > 0 desde Movimientos", "Cantidad": 0, "Interpretación": "Por eso no aparece el carrito de salida."},
                {"Validación": "Filas en hoja Kardex_Consolidado", "Cantidad": kardex_sheet_rows, "Interpretación": "Hoja calculada/visual en Google Sheets."},
                {"Validación": "Saldos > 0 en Kardex_Consolidado", "Cantidad": kardex_sheet_positive, "Interpretación": "Si aquí hay saldos pero no en Movimientos, debe reconstruirse la hoja desde Movimientos o revisar tipos/cantidades."},
            ])
            st.dataframe(diag_df, use_container_width=True, hide_index=True)

            if productos_count > 0 and ingresos_count == 0:
                st.info(
                    "Siguiente paso recomendado: registre primero un movimiento de tipo Ingreso para el producto/lote. "
                    "Después de guardar el ingreso, el lote aparecerá automáticamente en Salida."
                )
            elif ingresos_count > 0 and stock_rows == 0:
                st.info(
                    "Hay ingresos registrados, pero no se logró construir stock. Revise que los movimientos tengan producto, lote, unidad, cantidad y tipo_movimiento correctos."
                )
            elif stock_rows > 0:
                st.info(
                    "El sistema sí calculó lotes desde Movimientos, pero todos tienen saldo cero o negativo. Revise el Kardex consolidado para confirmar consumo total."
                )

            if kardex_sheet_positive > 0 and ingresos_count == 0:
                st.warning(
                    "Se detectan saldos positivos en la hoja física Kardex_Consolidado, pero no hay ingresos válidos en Movimientos. "
                    "Para hacer salidas seguras, el sistema necesita movimientos de ingreso/devolución/corrección entrada, porque Movimientos es la bitácora oficial. "
                    "Revise que los tipos de movimiento estén escritos como Ingreso, Entrada, Devolución o Corrección entrada y que cantidad sea numérica."
                )

            if st.button("🔄 Actualizar datos desde Google Sheets", use_container_width=True):
                mark_data_dirty()
                rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            return

        if "salida_cart" not in st.session_state:
            st.session_state["salida_cart"] = []

        disponibles = disponibles.sort_values(["producto", "fecha_vencimiento", "lote"], na_position="last").reset_index(drop=True)
        disponibles["label"] = disponibles.apply(
            lambda r: f"{r['producto']} | Marca: {r['marca']} | Lote: {r['lote']} | Vence: {r['fecha_vencimiento']} | Stock: {float(r['stock_actual']):,.0f} {r['unidad']}",
            axis=1,
        )

        # ── Selector de producto FUERA del form → rerun inmediato al cambiar ──
        productos_lista = ["Todos"] + sorted(disponibles["producto"].dropna().astype(str).unique().tolist())
        filtro_producto = st.selectbox(
            "Producto",
            productos_lista,
            key="sel_producto_carrito",
        )

        lotes_view = disponibles.copy()
        if filtro_producto != "Todos":
            lotes_view = lotes_view[lotes_view["producto"].astype(str) == filtro_producto]

        # Mostrar info del lote seleccionado antes del form
        lotes_opciones = lotes_view["label"].tolist()

        # ── Form de cantidad: solo lote y cantidad, el producto ya está fijo ──
        with st.form("frm_add_cart_salida", clear_on_submit=True):
            clote, ccant = st.columns([2.5, .8])
            label = clote.selectbox(
                "Lote disponible",
                lotes_opciones if lotes_opciones else ["— Sin lotes disponibles —"],
                help="Lotes filtrados según el producto seleccionado arriba.",
            )
            cantidad_item = ccant.number_input("Cantidad", min_value=0.0, step=1.0, format="%.2f")
            add_item = st.form_submit_button("➕ Agregar al carrito", use_container_width=True)

        # Buscar la fila seleccionada DESPUÉS del form
        selected_stock_row = None
        if lotes_opciones and label in lotes_view["label"].values:
            selected_stock_row = lotes_view[lotes_view["label"] == label].iloc[0]

        if add_item:
            if not lotes_opciones or selected_stock_row is None:
                st.error("⚠️ No hay lotes disponibles para el producto seleccionado.")
            elif cantidad_item <= 0:
                st.error("⚠️ Ingrese una cantidad mayor a cero.")
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
                    prod_info_item = _catalog_row_by_id(clean_str(selected_stock_row.get("producto_id", "")))
                    st.session_state["salida_cart"].append({
                        "item_key": item_key,
                        "producto_id": clean_str(selected_stock_row["producto_id"]),
                        "producto": clean_str(selected_stock_row["producto"]),
                        "marca": clean_str(selected_stock_row["marca"]),
                        "lote": clean_str(selected_stock_row["lote"]),
                        "fecha_vencimiento": format_date(selected_stock_row["fecha_vencimiento"]),
                        "unidad": clean_str(selected_stock_row["unidad"]),
                        "categoria": clean_str(prod_info_item.get("categoria", "")),
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

        # ── Paso 3: selección de sitio FUERA del form para auto-relleno ──────
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        st.markdown(
            "<div class='step-title'><span class='step-badge'>3</span>Sitio receptor y datos de entrega</div><hr class='step-divider'>",
            unsafe_allow_html=True,
        )

        # Fecha del acta: por defecto es hoy, pero el usuario puede cambiarla
        fecha_col, _ = st.columns([1, 2])
        fecha_acta = fecha_col.date_input(
            "📅 Fecha del acta de entrega",
            value=pd.Timestamp.now().date(),
            help="Se usa en el acta PDF y en el registro del movimiento. Por defecto es la fecha actual, puede cambiarla si es necesario.",
            key=f"fecha_acta_salida_{mov_nonce}",
        )
        fecha = pd.Timestamp(fecha_acta)  # convertir a Timestamp para compatibilidad

        # Selector de sitio fuera del form → rerun inmediato al cambiar
        col_solic, col_pers = st.columns(2)
        solicitante = col_solic.selectbox(
            "Sitio / unidad solicitante *",
            [""] + solicitantes[active_mask(solicitantes)]["unidad_solicitante"].dropna().astype(str).sort_values().tolist(),
            key=f"sel_solic_salida_{mov_nonce}",
            help="Al seleccionar el sitio se cargan automáticamente el encargado y su cargo.",
        )

        # Buscar info del solicitante en la base
        solicitante_info   = get_first_match(solicitantes, "unidad_solicitante", solicitante) if solicitante else {}
        responsable_auto   = clean_str(solicitante_info.get("responsable", "")) if solicitante_info else ""
        dpto_auto          = clean_str(solicitante_info.get("departamento", "")) if solicitante_info else ""
        tel_auto           = clean_str(solicitante_info.get("telefono", ""))     if solicitante_info else ""

        # Buscar cargo del responsable en la tabla Personal
        personal_resp_row  = get_first_match(personal_df, "nombre", responsable_auto) if responsable_auto else {}
        cargo_auto         = clean_str(personal_resp_row.get("cargo", "")) if personal_resp_row else ""
        if not cargo_auto:
            cargo_auto = "Responsable del sitio"

        # Mostrar tarjeta del encargado si ya seleccionó un sitio
        if solicitante and responsable_auto:
            st.markdown(
                f"""<div style="border:1px solid rgba(56,189,248,.25); border-radius:14px; padding:12px 16px;
                                margin:4px 0 12px 0; background:rgba(56,189,248,.06); display:flex; gap:16px; flex-wrap:wrap;">
                    <div style="font-size:13px;">
                        <span style="color:#64748B; font-size:11px; text-transform:uppercase; letter-spacing:.06em;">Encargado del sitio</span><br>
                        <b style="color:#F8FAFC; font-size:15px;">👤 {responsable_auto}</b>
                    </div>
                    <div style="font-size:13px;">
                        <span style="color:#64748B; font-size:11px; text-transform:uppercase; letter-spacing:.06em;">Cargo</span><br>
                        <b style="color:#CBD5E1;">{cargo_auto}</b>
                    </div>
                    {"" if not dpto_auto else f'<div style="font-size:13px;"><span style="color:#64748B; font-size:11px; text-transform:uppercase; letter-spacing:.06em;">Departamento</span><br><b style="color:#CBD5E1;">{dpto_auto}</b></div>'}
                    {"" if not tel_auto else f'<div style="font-size:13px;"><span style="color:#64748B; font-size:11px; text-transform:uppercase; letter-spacing:.06em;">Teléfono</span><br><b style="color:#CBD5E1;">{tel_auto}</b></div>'}
                </div>""",
                unsafe_allow_html=True,
            )
        elif solicitante and not responsable_auto:
            st.markdown(
                "<div class='alert-orange'>⚠️ Este sitio no tiene un encargado registrado. "
                "Puede agregarlo en el catálogo de Solicitantes y el campo se llenará automáticamente.</div>",
                unsafe_allow_html=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

        # ── Formulario principal de salida ────────────────────────────────────
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        with st.form(f"frm_confirmar_salida_{mov_nonce}", clear_on_submit=True):
            st.markdown(
                "<div class='step-title'><span class='step-badge'>4</span>Datos de entrega y observaciones</div><hr class='step-divider'>",
                unsafe_allow_html=True,
            )
            d1, d2 = st.columns(2)
            usuario  = d1.text_input("Registrado por", value=st.session_state.get("nombre_usuario", "Usuario"))
            personal = d2.selectbox(
                "Personal que entrega *",
                [""] + personal_df[active_mask(personal_df)]["nombre"].dropna().astype(str).sort_values().tolist(),
            )

            # Campos de receptor: pre-rellenos y bloqueados desde la base de datos
            st.markdown("<hr class='step-divider'>", unsafe_allow_html=True)
            st.markdown(
                "<span style='font-size:12px; color:#64748B;'>Los siguientes campos se rellenan automáticamente "
                "desde el catálogo del sitio receptor y no son editables.</span>",
                unsafe_allow_html=True,
            )
            r1, r2 = st.columns(2)
            recibe_nombre = r1.text_input(
                "Persona que recibe (del catálogo)",
                value=responsable_auto,
                disabled=True,
                help="Se obtiene automáticamente del responsable registrado para este sitio.",
            )
            recibe_cargo  = r2.text_input(
                "Cargo",
                value=cargo_auto,
                disabled=True,
                help="Se obtiene del cargo registrado en Personal para este responsable.",
            )
            observacion = st.text_area("Observación para movimientos y acta", placeholder="Detalle de solicitud, referencia o comentario de entrega.")
            tipo_entrega_auto = acta_tipo_entrega_desde_categorias(cart)
            tipo_entrega_texto = st.text_input(
                "Texto de categoría para el acta",
                value=tipo_entrega_auto,
                help="Se calcula automáticamente según la categoría del producto en el catálogo. Puede ajustarlo antes de generar el PDF. Ejemplo: reactivos, insumos, equipos, materiales."
            )
            st.caption(f"En el acta se imprimirá: 'se hace entrega formal de los siguientes {tipo_entrega_texto}...'")
            generar_acta = st.checkbox("Generar acta de entrega en PDF", value=True)
            submitted = st.form_submit_button("📤 Solicitar salida →", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if submitted:
            if not require_permission(data, "registrar_movimientos", "registrar movimientos"):
                return
            cart = st.session_state.get("salida_cart", [])
            if not cart:
                st.error("⚠️ Debe agregar al menos un producto al carrito.")
                return
            if not solicitante:
                st.error("⚠️ Debe seleccionar el sitio/unidad solicitante.")
                return
            if not personal:
                st.error("⚠️ Debe seleccionar el personal que entrega.")
                return
            total_sal = sum(float(x.get("cantidad", 0)) for x in cart)
            summary_sal = ", ".join(f"{x['producto']} x{float(x['cantidad']):g}" for x in cart[:3])
            if len(cart) > 3: summary_sal += f" y {len(cart)-3} más"
            set_confirm_pending(f"conf_sal_{solicitante}", {
                "solicitante": solicitante, "personal": personal,
                "n_items": len(cart), "total": total_sal, "resumen": summary_sal,
            })

        _ck_sal = f"conf_sal_{solicitante}" if (submitted or confirm_pending(f"conf_sal_{solicitante}")) else ""
        if _ck_sal and confirm_pending(_ck_sal):
            _p_sal = get_confirm_payload(_ck_sal)
            _confirmed_sal = render_confirm_box(
                key=_ck_sal,
                title=f"¿Confirmar salida a {_p_sal.get('solicitante','')}?",
                body=(
                    f"<b>Destinatario:</b> {_p_sal.get('solicitante','')}<br>"
                    f"<b>Entrega:</b> {_p_sal.get('personal','')}<br>"
                    f"<b>Productos:</b> {_p_sal.get('resumen','')}<br>"
                    f"<b>Total unidades:</b> {_p_sal.get('total',0):g}"
                ),
                confirm_label="✅ Confirmar salida",
                cancel_label="↩ Cancelar",
                danger=False,
            )
            if not _confirmed_sal:
                return
            cart = st.session_state.get("salida_cart", [])
            # Validación final contra stock actual en pantalla.
            for item in cart:
                same = [x for x in cart if x.get("item_key") == item.get("item_key")]
                qty_same = sum(float(x.get("cantidad", 0) or 0) for x in same)
                if qty_same > float(item.get("stock_disponible", 0) or 0):
                    st.error(f"La cantidad del lote {item.get('lote')} supera el stock disponible.")
                    return
            fecha_ts   = pd.Timestamp.now()   # timestamp del momento exacto de guardado (para el ID del acta)
            acta_id    = f"ACTA-{fecha_ts.strftime('%Y%m%d%H%M%S')}"
            ids = next_movement_ids(movimientos, len(cart))
            rows = []
            for mov_id, item in zip(ids, cart):
                rows.append({
                    "movimiento_id": mov_id,
                    "fecha": fecha.strftime("%Y-%m-%d"),  # fecha seleccionada por el usuario en el acta
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
                    "categoria": clean_str(item.get("categoria", "")),
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
                pdf_bytes = build_acta_entrega_pdf(rows, solicitante_info, personal_info, fecha, recibe_nombre, recibe_cargo, observacion=observacion, tipo_entrega_texto=tipo_entrega_texto)
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
        st.caption("Seleccione una fila de la tabla. Si su navegador no marca la fila con clic, use el selector de respaldo debajo de la tabla.")
        selection_event = st.dataframe(
            vista_salidas,
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="devolucion_salidas_table",
        )
        salidas["label"] = salidas.apply(lambda r: f"{r['movimiento_id']} | {r['fecha']} | {r['producto']} | Lote {r['lote']} | Cant. {r['cantidad']} {r['unidad']} | {r['solicitante']}", axis=1)
        labels_salidas = salidas["label"].tolist()

        selected_position = None
        try:
            rows_selected = getattr(getattr(selection_event, "selection", None), "rows", None)
            if rows_selected is None and isinstance(selection_event, dict):
                rows_selected = selection_event.get("selection", {}).get("rows", [])
            if rows_selected:
                selected_position = int(rows_selected[0])
        except Exception:
            selected_position = None

        if selected_position is not None and 0 <= selected_position < len(labels_salidas):
            label = labels_salidas[selected_position]
            st.success(f"Salida seleccionada desde la tabla: {label}")
        else:
            label = st.selectbox("Salida a devolver", labels_salidas, key="devolucion_salida_label_fallback")

        salida_match = salidas[salidas["label"] == label]
        if salida_match.empty:
            st.warning("No se pudo identificar la salida seleccionada. Actualice los datos desde Google Sheets e intente nuevamente.")
            st.markdown("</div>", unsafe_allow_html=True)
            return
        salida_row = salida_match.iloc[0]
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
            if not require_permission(data, "registrar_movimientos", "registrar movimientos"):
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
        if disponibles.empty:
            st.error("No hay lotes con stock mayor a cero para registrar una Corrección salida.")
            st.info("El sistema verificó la base y encontró lotes, pero ninguno tiene saldo disponible para restar. Revise ingresos, salidas y cantidades en Movimientos.")
            st.markdown("</div>", unsafe_allow_html=True)
            return
    disponibles = disponibles.sort_values(["producto", "fecha_vencimiento", "lote"], na_position="last").reset_index(drop=True)
    disponibles["label"] = disponibles.apply(
        lambda r: f"{r['producto']} | Marca: {r['marca']} | Lote: {r['lote']} | Vence: {r['fecha_vencimiento']} | Stock: {float(r['stock_actual']):,.0f} {r['unidad']}",
        axis=1,
    )
    labels_disponibles = disponibles["label"].dropna().astype(str).tolist()
    if not labels_disponibles:
        st.error("No hay productos/lotes válidos para seleccionar en la corrección.")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    label = st.selectbox("Producto/lote a corregir", labels_disponibles)
    selected_match = disponibles[disponibles["label"].astype(str) == str(label)]
    if selected_match.empty:
        st.warning("La selección ya no está disponible. Actualice datos desde Google Sheets e intente nuevamente.")
        st.markdown("</div>", unsafe_allow_html=True)
        return
    selected = selected_match.iloc[0]
    prod_row = _catalog_row_by_id(clean_str(selected["producto_id"]))
    if clean_str(prod_row.get("nombre_producto", "")) == "":
        prod_row["nombre_producto"] = selected["producto"]
        prod_row["marca_default"] = selected["marca"]
        prod_row["unidad_default"] = selected["unidad"]
    _product_card(prod_row, "Producto seleccionado para corrección")
    st.info(f"Stock actual del lote: {float(selected.get('stock_actual', 0) or 0):,.0f} {clean_str(selected.get('unidad', ''))}.")
    st.markdown("</div>", unsafe_allow_html=True)

    mov_nonce = int(st.session_state.get("mov_form_nonce", 0))

    # Explicación contextual según tipo de ajuste
    if tipo == "Corrección entrada":
        st.markdown(
            "<div class='alert-green' style='margin-bottom:14px;'>"
            "🔧 <b>Ajuste de entrada:</b> Se agregará una cantidad adicional al stock del lote seleccionado. "
            "Use esto cuando el conteo físico muestra más unidades de las que indica el sistema. "
            "Este ajuste crea un <b>nuevo movimiento</b>, no modifica el ingreso original."
            "</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<div class='alert-orange' style='margin-bottom:14px;'>"
            "🔧 <b>Ajuste de salida:</b> Se descontará una cantidad del stock del lote seleccionado. "
            "Use esto cuando el conteo físico muestra menos unidades de las que indica el sistema. "
            "Este ajuste crea un <b>nuevo movimiento</b>, no modifica la salida original."
            "</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    with st.form(f"frm_correccion_{mov_nonce}", clear_on_submit=True):
        st.markdown("#### 2) Datos del ajuste")
        a, b, c = st.columns(3)
        fecha = a.date_input("Fecha del ajuste", value=TODAY.date())
        usuario = b.text_input("Registrado por", value=st.session_state.get("nombre_usuario", "Usuario"))
        cantidad = c.number_input("Cantidad a ajustar", min_value=0.0, step=1.0, format="%.2f",
                                  help="Ingrese solo la diferencia, no el total. Ej: si faltan 5 unidades, escriba 5.")
        d, e, f = st.columns(3)
        d.text_input("Producto", value=clean_str(selected["producto"]), disabled=True)
        e.text_input("Lote", value=clean_str(selected["lote"]), disabled=True)
        f.text_input("Unidad", value=clean_str(selected["unidad"]), disabled=True)
        observacion = st.text_area(
            "Justificación del ajuste *",
            placeholder="Ej: diferencia encontrada en conteo físico del día 15/05/2026, lote verificado por bodeguero.",
        )
        submitted = st.form_submit_button("💾 Guardar ajuste de inventario", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        if st.session_state.get("rol") == "Consulta":
            st.error("El rol Consulta no puede registrar movimientos.")
            return
        if cantidad <= 0:
            st.error("La cantidad debe ser mayor a cero.")
            return
        if tipo == "Corrección salida" and cantidad > float(selected.get("stock_actual", 0) or 0):
            st.error(f"No puede ajustar más de {float(selected.get('stock_actual', 0) or 0):g} unidades (stock actual del lote).")
            return
        if not observacion:
            st.error("La justificación es obligatoria para registrar un ajuste.")
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
            "solicitante": "",
            "personal": clean_str(usuario),
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
        _sync_after_save([row], "✅ Ajuste de inventario guardado. El stock y Kardex consolidado fueron actualizados.")

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
    total_salida    = float(df["salida_total"].sum())  if not df.empty else 0
    total_saldo     = float(df["saldo_actual"].sum())  if not df.empty else 0
    lotes_criticos  = int((df["dias_para_vencer"].fillna(999).astype(float) <= 30).sum()) if not df.empty else 0
    lotes_bajo_min  = int((
        (to_number(df["saldo_actual"]) > 0) &
        (to_number(df["saldo_actual"]) < to_number(df["stock_minimo"]))
    ).sum()) if not df.empty else 0
    pct_consumido   = round(total_salida / total_ingresado * 100, 1) if total_ingresado > 0 else 0.0

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1: kpi_card("Lotes", f"{len(df):,}", "en la vista actual", color="#38BDF8")
    with k2: kpi_card("Ingresado", f"{total_ingresado:,.0f}", "unidades totales", color="#22C55E")
    with k3: kpi_card("Entregado", f"{total_salida:,.0f}", "unidades salidas", color="#F97316")
    with k4: kpi_card("Saldo", f"{total_saldo:,.0f}", "unidades disponibles", color="#8B5CF6",
                       trend=f"Consumido: {pct_consumido}%")
    with k5: kpi_card("⏰ Por vencer", f"{lotes_criticos}", "en ≤ 30 días", color="#EAB308")
    with k6: kpi_card("📉 Bajo mínimo", f"{lotes_bajo_min}", "lotes con stock bajo", color="#EF4444")

    st.markdown("")

    # ── Color-coded table ──────────────────────────────────────────────────
    card_start(
        "Tabla consolidada por lote",
        "Verde = disponible · Naranja = cerca de vencer o bajo mínimo · Rojo = vencido o sin stock",
    )

    if not df.empty:
        def _row_class(row: pd.Series) -> str:
            dias = float(pd.to_numeric(row.get("dias_para_vencer", 9999), errors="coerce") or 9999)
            saldo = float(pd.to_numeric(row.get("saldo_actual", 0), errors="coerce") or 0)
            min_s = float(pd.to_numeric(row.get("stock_minimo", 0), errors="coerce") or 0)
            if dias <= 0 or saldo <= 0:
                return "🔴"
            if dias <= 30 or (min_s > 0 and saldo < min_s):
                return "🟡"
            return "🟢"
        df = df.copy()
        df.insert(0, "🚦", df.apply(_row_class, axis=1))

    cols_view = [
        "🚦", "estado", "producto", "marca", "lote", "unidad",
        "fecha_ingreso", "proveedor_ingreso",
        "fecha_vencimiento", "dias_para_vencer",
        "entrada_total", "salida_total", "saldo_actual", "stock_minimo",
        "numero_salidas", "fecha_ultima_salida", "ultimo_entregado_a",
        "detalle_salidas",
    ]
    available_cols = [c for c in cols_view if c in df.columns]
    st.dataframe(
        df[available_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "🚦":             st.column_config.TextColumn("", width=30),
            "entrada_total":  st.column_config.NumberColumn("📥 Entrada",  format="%.0f"),
            "salida_total":   st.column_config.NumberColumn("📤 Salida",   format="%.0f"),
            "saldo_actual":   st.column_config.NumberColumn("📦 Saldo",    format="%.0f"),
            "stock_minimo":   st.column_config.NumberColumn("Mín.",        format="%.0f"),
            "dias_para_vencer": st.column_config.NumberColumn("Días vence", format="%d"),
            "numero_salidas": st.column_config.NumberColumn("# Salidas",   format="%d"),
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
            "estado_registro": "Activo",
            "creado_por": current_username(),
            "fecha_creacion": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            "estado_registro": "Activo" if activo == "Sí" else "Inactivo",
            "creado_por": current_username(),
            "fecha_creacion": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            "estado_registro": "Activo" if activo == "Sí" else "Inactivo",
            "creado_por": current_username(),
            "fecha_creacion": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            "estado_registro": "Activo" if activo == "Sí" else "Inactivo",
            "creado_por": current_username(),
            "fecha_creacion": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        storage.append_row("Personal", row)
        set_flash("Personal guardado correctamente. El formulario quedó limpio para registrar un nuevo personal.")
        rerun()


def catalog_editor(storage, data: Dict[str, pd.DataFrame], sheet: str, title: str) -> None:
    """CRUD controlado de catálogos.

    No elimina físicamente registros: los desactiva/reactiva para conservar trazabilidad
    y evitar romper movimientos históricos.
    """
    card_start(title, "Consulte, edite o desactive registros según sus permisos. Los cambios quedan auditados.")
    df = ensure_columns(data[sheet], sheet)
    id_col = catalog_id_col(sheet)
    name_col = catalog_display_col(sheet)
    base = CATALOG_PERMISSION_BASE.get(sheet, sheet.lower())

    search = st.text_input(f"Buscar en {title}", key=f"search_{sheet}")
    view = df.copy()
    if search:
        mask = view.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        view = view[mask]

    st.dataframe(view, use_container_width=True, hide_index=True)

    can_edit = user_has_permission(data, f"editar_{base}")
    can_deactivate = user_has_permission(data, f"desactivar_{base}")
    if not can_edit and not can_deactivate:
        st.info("Puede consultar este catálogo. Para modificar o desactivar registros, solicite permiso al administrador.")
        return

    if df.empty or not id_col or not name_col:
        st.info("No hay registros para gestionar.")
        return

    labels = df.apply(lambda r: f"{clean_str(r.get(name_col, ''))} | ID: {clean_str(r.get(id_col, ''))} | Estado: {clean_str(r.get('activo', 'Sí'))}", axis=1).tolist()
    selected_label = st.selectbox(f"Seleccionar registro para gestionar - {title}", labels, key=f"sel_crud_{sheet}")
    selected_idx = labels.index(selected_label)
    selected = df.iloc[selected_idx].copy()
    registro_id = clean_str(selected.get(id_col, ""))

    editable_cols = [
        c for c in SHEET_COLUMNS[sheet]
        if c not in {id_col, "creado_por", "fecha_creacion", "modificado_por", "fecha_modificacion", "motivo_modificacion"}
    ]
    values = {}
    with st.form(f"frm_crud_{sheet}_{registro_id}"):
        st.markdown("##### Datos del registro")
        cols = st.columns(2)
        for idx, col in enumerate(editable_cols):
            current = selected.get(col, "")
            label = col.replace("_", " ").title()
            with cols[idx % 2]:
                if col in ["activo"]:
                    opts = ["Sí", "No"]
                    cur = clean_str(current) or "Sí"
                    values[col] = st.selectbox(label, opts, index=opts.index(cur) if cur in opts else 0, disabled=not can_edit)
                elif col == "estado_registro":
                    opts = ["Activo", "Inactivo"]
                    cur = clean_str(current) or "Activo"
                    values[col] = st.selectbox(label, opts, index=opts.index(cur) if cur in opts else 0, disabled=not can_edit)
                elif col in ["stock_minimo", "dias_alerta_vencimiento"]:
                    values[col] = st.number_input(label, min_value=0.0, value=float(to_number(pd.Series([current])).iloc[0]), step=1.0, disabled=not can_edit)
                elif col == "observacion" or "direccion" in col or "descripcion" in col:
                    values[col] = st.text_area(label, value=clean_str(current), disabled=not can_edit)
                else:
                    values[col] = st.text_input(label, value=clean_str(current), disabled=not can_edit)
        motivo = st.text_area("Motivo del cambio *", placeholder="Explique por qué se edita, desactiva o reactiva este registro.")
        c1, c2, c3 = st.columns(3)
        save_btn = c1.form_submit_button("💾 Guardar cambios", disabled=not can_edit, use_container_width=True)
        deactivate_btn = c2.form_submit_button("🚫 Desactivar", disabled=not can_deactivate, use_container_width=True)
        reactivate_btn = c3.form_submit_button("✅ Reactivar", disabled=not can_deactivate, use_container_width=True)

    if save_btn or deactivate_btn or reactivate_btn:
        if not motivo:
            st.error("Debe ingresar el motivo del cambio para auditoría.")
            return
        new_df = df.copy()
        mask = new_df[id_col].astype(str) == registro_id
        if not mask.any():
            st.error("No se encontró el registro seleccionado.")
            return

        audit_rows = []
        action = "EDITAR_REGISTRO"
        if save_btn:
            for col, new_value in values.items():
                old_value = clean_str(new_df.loc[mask, col].iloc[0]) if col in new_df.columns else ""
                new_text = clean_str(new_value)
                if old_value != new_text:
                    new_df.loc[mask, col] = new_value
                    audit_rows.append({
                        "auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}",
                        "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "usuario": current_username(),
                        "rol": current_role(),
                        "accion": action,
                        "modulo": sheet,
                        "registro_id": registro_id,
                        "campo": col,
                        "valor_anterior": old_value,
                        "valor_nuevo": new_text,
                        "motivo": motivo,
                        "detalle": f"Edición de {title}",
                    })
        elif deactivate_btn:
            action = "DESACTIVAR_REGISTRO"
            if "activo" in new_df.columns:
                old_value = clean_str(new_df.loc[mask, "activo"].iloc[0])
                new_df.loc[mask, "activo"] = "No"
                audit_rows.append({"auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}", "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"), "usuario": current_username(), "rol": current_role(), "accion": action, "modulo": sheet, "registro_id": registro_id, "campo": "activo", "valor_anterior": old_value, "valor_nuevo": "No", "motivo": motivo, "detalle": f"Desactivación lógica de {title}"})
            if "estado_registro" in new_df.columns:
                old_value = clean_str(new_df.loc[mask, "estado_registro"].iloc[0])
                new_df.loc[mask, "estado_registro"] = "Inactivo"
                audit_rows.append({"auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}", "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"), "usuario": current_username(), "rol": current_role(), "accion": action, "modulo": sheet, "registro_id": registro_id, "campo": "estado_registro", "valor_anterior": old_value, "valor_nuevo": "Inactivo", "motivo": motivo, "detalle": f"Desactivación lógica de {title}"})
        elif reactivate_btn:
            action = "REACTIVAR_REGISTRO"
            if "activo" in new_df.columns:
                old_value = clean_str(new_df.loc[mask, "activo"].iloc[0])
                new_df.loc[mask, "activo"] = "Sí"
                audit_rows.append({"auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}", "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"), "usuario": current_username(), "rol": current_role(), "accion": action, "modulo": sheet, "registro_id": registro_id, "campo": "activo", "valor_anterior": old_value, "valor_nuevo": "Sí", "motivo": motivo, "detalle": f"Reactivación de {title}"})
            if "estado_registro" in new_df.columns:
                old_value = clean_str(new_df.loc[mask, "estado_registro"].iloc[0])
                new_df.loc[mask, "estado_registro"] = "Activo"
                audit_rows.append({"auditoria_id": f"AUD-{uuid.uuid4().hex[:12].upper()}", "fecha_evento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"), "usuario": current_username(), "rol": current_role(), "accion": action, "modulo": sheet, "registro_id": registro_id, "campo": "estado_registro", "valor_anterior": old_value, "valor_nuevo": "Activo", "motivo": motivo, "detalle": f"Reactivación de {title}"})

        if "modificado_por" in new_df.columns:
            new_df.loc[mask, "modificado_por"] = current_username()
        if "fecha_modificacion" in new_df.columns:
            new_df.loc[mask, "fecha_modificacion"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        if "motivo_modificacion" in new_df.columns:
            new_df.loc[mask, "motivo_modificacion"] = motivo

        storage.save(sheet, ensure_columns(new_df, sheet))
        append_audit_rows(storage, audit_rows)
        if sheet == "Productos":
            sync_data = dict(data)
            sync_data["Productos"] = ensure_columns(new_df, sheet)
            try:
                sync_kardex_consolidado_sheet(storage, sync_data)
            except Exception:
                pass
        set_flash(f"{title}: cambio guardado correctamente con auditoría.")
        rerun()

def page_catalogos(storage, data: Dict[str, pd.DataFrame]) -> None:
    section_header("⚙️ Catálogos", "Formularios independientes para mantener ordenados productos, proveedores, solicitantes y personal.")
    if st.session_state.get("rol") == "Consulta":
        st.info("Su rol es de consulta. Puede visualizar catálogos, pero no registrar cambios.")

    tab_prod, tab_prov, tab_sol, tab_per = st.tabs(["📦 Productos", "🚚 Proveedores", "🏥 Solicitantes", "👤 Personal"])

    with tab_prod:
        if user_has_permission(data, "crear_productos"):
            product_form(storage, data)
        else:
            st.info("Puede consultar productos. Para crear productos solicite permiso al administrador.")
        catalog_editor(storage, data, "Productos", "Listado de productos")
    with tab_prov:
        if user_has_permission(data, "crear_proveedores"):
            provider_form(storage, data)
        else:
            st.info("Puede consultar proveedores. Para crear proveedores solicite permiso al administrador.")
        catalog_editor(storage, data, "Proveedores", "Listado de proveedores")
    with tab_sol:
        if user_has_permission(data, "crear_solicitantes"):
            requester_form(storage, data)
        else:
            st.info("Puede consultar solicitantes. Para crear solicitantes solicite permiso al administrador.")
        catalog_editor(storage, data, "Solicitantes", "Listado de solicitantes")
    with tab_per:
        if user_has_permission(data, "crear_personal"):
            staff_form(storage, data)
        else:
            st.info("Puede consultar personal. Para crear personal solicite permiso al administrador.")
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



# ------------------------- GESTIÓN CONTROLADA Y AUDITORÍA -------------------------
def render_movement_crud_controls(storage, data: Dict[str, pd.DataFrame]) -> None:
    """Gestión completa de movimientos: búsqueda, edición contextual por tipo y anulación lógica.

    - Los campos del formulario de edición se adaptan al tipo de movimiento (Ingreso, Salida, Devolución, Corrección).
    - No se eliminan filas: la anulación cambia estado_movimiento → Anulado y el stock se recalcula.
    - Todo cambio queda en Auditoria_Cambios con usuario, campo, valor anterior/nuevo y motivo.
    """
    section_header(
        "🧾 Gestión de movimientos",
        "Busque, revise, edite o anule movimientos registrados. El formulario de edición se adapta automáticamente al tipo.",
    )

    movimientos = ensure_columns(data.get("Movimientos", pd.DataFrame()), "Movimientos")
    if movimientos.empty:
        st.info("Aún no hay movimientos registrados. Registre ingresos o salidas primero.")
        return

    personal_df  = ensure_columns(data.get("Personal", pd.DataFrame()), "Personal")
    proveedores  = ensure_columns(data.get("Proveedores", pd.DataFrame()), "Proveedores")
    solicitantes = ensure_columns(data.get("Solicitantes", pd.DataFrame()), "Solicitantes")

    # ── 1. TABLA RESUMEN + FILTROS ─────────────────────────────────────────
    card_start("Movimientos registrados", "Use los filtros para encontrar el movimiento que desea gestionar.")

    col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
    filtro_texto  = col_f1.text_input("🔍 Buscar por producto, lote, ID o proveedor",
                                       key="buscar_mov_crud",
                                       placeholder="Escriba cualquier término...")
    estados_disponibles = sorted(movimientos["estado_movimiento"].dropna().astype(str).unique().tolist())
    filtro_estado = col_f2.multiselect("Estado", estados_disponibles, default=estados_disponibles, key="filtro_estado_crud")
    filtro_tipo   = col_f3.selectbox("Tipo de movimiento", ["Todos"] + TIPOS_MOVIMIENTO, key="filtro_tipo_crud")

    view = movimientos.copy()
    if filtro_estado:
        view = view[view["estado_movimiento"].isin(filtro_estado)]
    if filtro_tipo != "Todos":
        view = view[view["tipo_movimiento"].astype(str) == filtro_tipo]
    if filtro_texto:
        view = view[view.astype(str).apply(
            lambda col: col.str.contains(filtro_texto, case=False, na=False)
        ).any(axis=1)]

    cols_tabla = ["movimiento_id", "fecha", "tipo_movimiento", "producto", "lote",
                  "cantidad", "unidad", "proveedor", "solicitante",
                  "estado_movimiento", "usuario_registro", "modificado_por"]
    st.dataframe(view[cols_tabla], use_container_width=True, hide_index=True)

    if view.empty:
        st.warning("No hay movimientos con esos filtros. Ajuste los criterios de búsqueda.")
        return

    st.markdown("---")

    # ── 2. SELECCIÓN DEL MOVIMIENTO ────────────────────────────────────────
    st.markdown("#### ✏️ Seleccione el movimiento a gestionar")

    def _label(r: pd.Series) -> str:
        estado = clean_str(r.get("estado_movimiento", "Vigente")).upper()
        tipo   = clean_str(r.get("tipo_movimiento", ""))
        prod   = clean_str(r.get("producto", ""))
        lote   = clean_str(r.get("lote", ""))
        cant   = clean_str(r.get("cantidad", ""))
        unid   = clean_str(r.get("unidad", ""))
        fecha  = clean_str(r.get("fecha", ""))
        mid    = clean_str(r.get("movimiento_id", ""))
        return f"[{estado}] {mid} | {fecha} | {tipo} | {prod} | Lote: {lote} | {cant} {unid}"

    labels = view.apply(_label, axis=1).tolist()
    sel_label = st.selectbox("Movimiento", labels, key="sel_mov_crud",
                             help="Seleccione el movimiento que desea editar o anular.")
    sel_idx   = labels.index(sel_label)
    selected  = view.iloc[sel_idx].copy()
    mov_id    = clean_str(selected.get("movimiento_id", ""))
    tipo_mov  = clean_str(selected.get("tipo_movimiento", ""))
    est_mov   = clean_str(selected.get("estado_movimiento", "Vigente"))
    cant_actual = float(pd.to_numeric(selected.get("cantidad", 0), errors="coerce") or 0)

    # ── 3. TARJETA DE RESUMEN DEL MOVIMIENTO SELECCIONADO ─────────────────
    color_estado = "#22C55E" if est_mov.lower() == "vigente" else "#EF4444"
    icono_tipo = {
        "Ingreso": "📥", "Salida": "📤", "Devolución": "↩️",
        "Corrección entrada": "➕", "Corrección salida": "➖",
    }.get(tipo_mov, "📋")

    prov_val  = clean_str(selected.get("proveedor", ""))
    solic_val = clean_str(selected.get("solicitante", ""))
    pers_val  = clean_str(selected.get("personal", ""))
    extra_label = ""
    if tipo_mov == "Ingreso":
        extra_label = f"<b>Proveedor:</b> {prov_val or '—'}"
    elif tipo_mov in ("Salida", "Devolución"):
        extra_label = f"<b>Sitio/unidad:</b> {solic_val or '—'} &nbsp;|&nbsp; <b>Personal:</b> {pers_val or '—'}"
    elif tipo_mov in ("Corrección entrada", "Corrección salida"):
        extra_label = f"<b>Observación:</b> {clean_str(selected.get('observacion', ''))[:60]}"

    st.markdown(
        f"""
        <div style="border:1px solid {color_estado}33; border-radius:18px; padding:16px 20px; margin:12px 0;
                    background:linear-gradient(135deg,rgba(15,23,42,.90),rgba(2,6,23,.80));">
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:12px;">
                <span style="font-size:22px;">{icono_tipo}</span>
                <span style="font-size:18px; font-weight:900; color:#F8FAFC;">{mov_id}</span>
                <span style="padding:3px 12px; border-radius:999px; background:{color_estado}22;
                             color:{color_estado}; border:1px solid {color_estado}55; font-size:12px; font-weight:700;">
                    {est_mov}
                </span>
                <span style="padding:3px 12px; border-radius:999px; background:rgba(56,189,248,.15);
                             color:#38BDF8; border:1px solid rgba(56,189,248,.30); font-size:12px;">
                    {tipo_mov}
                </span>
            </div>
            <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:12px;">
                <div style="font-size:12px; color:#94A3B8;">
                    <b style="color:#E5E7EB; display:block; margin-bottom:2px;">Producto</b>
                    {clean_str(selected.get('producto',''))}
                </div>
                <div style="font-size:12px; color:#94A3B8;">
                    <b style="color:#E5E7EB; display:block; margin-bottom:2px;">Lote</b>
                    {clean_str(selected.get('lote',''))}
                </div>
                <div style="font-size:12px; color:#94A3B8;">
                    <b style="color:#E5E7EB; display:block; margin-bottom:2px;">Cantidad registrada</b>
                    <span style="font-size:20px; font-weight:900; color:#38BDF8;">{cant_actual:g}</span>
                    <span style="color:#64748B;"> {clean_str(selected.get('unidad',''))}</span>
                </div>
                <div style="font-size:12px; color:#94A3B8;">
                    <b style="color:#E5E7EB; display:block; margin-bottom:2px;">Fecha</b>
                    {clean_str(selected.get('fecha',''))}
                </div>
            </div>
            {f'<div style="margin-top:10px; font-size:12px; color:#94A3B8; border-top:1px solid rgba(148,163,184,.15); padding-top:10px;">{extra_label}</div>' if extra_label else ''}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── 4. TABS DE ACCIÓN ──────────────────────────────────────────────────
    tab_editar, tab_anular = st.tabs(["✏️ Editar movimiento", "🚫 Anular movimiento"])

    # ─────────────────────────────────────────────────────────────────────
    # TAB EDITAR: campos inteligentes según tipo de movimiento
    # ─────────────────────────────────────────────────────────────────────
    with tab_editar:
        if not user_has_permission(data, "editar_movimientos"):
            st.info("No tiene permiso para editar movimientos. Solicítelo al administrador.")
        elif est_mov.lower() == "anulado":
            st.markdown(
                "<div class='alert-red'>⛔ Este movimiento está <b>anulado</b> y ya no puede editarse. "
                "Los movimientos anulados son de solo lectura para preservar la trazabilidad.</div>",
                unsafe_allow_html=True,
            )
        else:
            # Advertencia contextual por tipo
            tips = {
                "Ingreso":            "💡 Para un ingreso puede corregir la cantidad, lote, marca, proveedor y fecha. El producto no cambia para preservar el Kardex.",
                "Salida":             "💡 Para una salida puede corregir la cantidad, el sitio receptor, el personal y la fecha. El lote no cambia porque identifica el stock afectado.",
                "Devolución":         "💡 Para una devolución puede corregir la cantidad, quién devolvió, el personal receptor y la fecha.",
                "Corrección entrada": "💡 Para correcciones solo se puede ajustar la cantidad, la fecha y la observación.",
                "Corrección salida":  "💡 Para correcciones solo se puede ajustar la cantidad, la fecha y la observación.",
            }
            st.markdown(
                f"<div class='alert-orange' style='margin-bottom:14px;'>{tips.get(tipo_mov, '💡 Edite solo los campos necesarios.')} "
                f"<br>📋 Todos los cambios quedan registrados en auditoría con su usuario, la fecha y el motivo.</div>",
                unsafe_allow_html=True,
            )

            # ── Valores actuales ──────────────────────────────────────────
            val_cant    = float(pd.to_numeric(selected.get("cantidad", 0), errors="coerce") or 0)
            val_fecha   = clean_str(selected.get("fecha", ""))
            val_lote    = clean_str(selected.get("lote", ""))
            val_marca   = clean_str(selected.get("marca", ""))
            val_unidad  = clean_str(selected.get("unidad", ""))
            val_prov    = clean_str(selected.get("proveedor", ""))
            val_oc      = clean_str(selected.get("orden_compra", ""))
            val_solic   = clean_str(selected.get("solicitante", ""))
            val_pers    = clean_str(selected.get("personal", ""))
            val_obs     = clean_str(selected.get("observacion", ""))
            val_costo   = float(pd.to_numeric(selected.get("costo_total", 0), errors="coerce") or 0)

            # Listas para selectbox
            lista_proveedores  = [""] + proveedores[active_mask(proveedores)]["proveedor"].dropna().astype(str).sort_values().tolist()
            lista_solicitantes = [""] + solicitantes[active_mask(solicitantes)]["unidad_solicitante"].dropna().astype(str).sort_values().tolist()
            lista_personal     = [""] + personal_df[active_mask(personal_df)]["nombre"].dropna().astype(str).sort_values().tolist()
            lista_unidades     = sorted(set([u for u in UNIDADES_DEFAULT + [val_unidad] if clean_str(u)]))
            unidad_idx         = lista_unidades.index(val_unidad) if val_unidad in lista_unidades else 0
            prov_idx           = lista_proveedores.index(val_prov) if val_prov in lista_proveedores else 0
            solic_idx          = lista_solicitantes.index(val_solic) if val_solic in lista_solicitantes else 0
            pers_idx           = lista_personal.index(val_pers) if val_pers in lista_personal else 0

            with st.form(f"frm_editar_mov_{mov_id}"):

                # ── CAMPO COMÚN: Cantidad ─────────────────────────────────
                st.markdown("##### 📦 Cantidad")
                col_cant, col_fecha = st.columns(2)
                nueva_cantidad = col_cant.number_input(
                    "Cantidad *",
                    min_value=0.0,
                    value=val_cant,
                    step=1.0,
                    format="%.2f",
                    help="Corrija la cantidad si fue ingresada incorrectamente.",
                )
                nueva_fecha = col_fecha.text_input("Fecha del movimiento", value=val_fecha,
                                                    help="Formato YYYY-MM-DD")

                # ── CAMPOS SEGÚN TIPO ─────────────────────────────────────
                new_values: dict = {}

                if tipo_mov == "Ingreso":
                    st.markdown("##### 📋 Datos del lote y proveedor")
                    c1, c2, c3 = st.columns(3)
                    nuevo_lote  = c1.text_input("Lote", value=val_lote)
                    nueva_marca = c2.text_input("Marca", value=val_marca)
                    nueva_unidad = c3.selectbox("Unidad", lista_unidades, index=unidad_idx)

                    st.markdown("##### 🏭 Proveedor y logística")
                    d1, d2 = st.columns(2)
                    nuevo_prov = d1.selectbox("Proveedor", lista_proveedores, index=prov_idx)
                    nuevo_oc   = d2.text_input("Orden de compra", value=val_oc)
                    nuevo_pers = st.selectbox("Personal que recibió", lista_personal, index=pers_idx,
                                              help="Persona del equipo interno que recibió el ingreso.")
                    nueva_obs  = st.text_area("Observación", value=val_obs)
                    nuevo_costo = st.number_input("Costo total", min_value=0.0, value=val_costo, step=1.0, format="%.2f")
                    new_values = {
                        "cantidad": nueva_cantidad, "fecha": nueva_fecha,
                        "lote": nuevo_lote, "marca": nueva_marca, "unidad": nueva_unidad,
                        "proveedor": nuevo_prov, "orden_compra": nuevo_oc,
                        "personal": nuevo_pers, "observacion": nueva_obs, "costo_total": nuevo_costo,
                    }

                elif tipo_mov == "Salida":
                    st.markdown("##### 📬 Destinatario de la salida")
                    d1, d2 = st.columns(2)
                    nuevo_solic = d1.selectbox("Sitio / unidad que recibió *", lista_solicitantes, index=solic_idx,
                                               help="¿A qué sitio o unidad fue entregada esta salida?")
                    nuevo_pers  = d2.selectbox("Personal que entregó", lista_personal, index=pers_idx,
                                               help="Persona del equipo que realizó la entrega.")
                    nueva_obs = st.text_area("Observación", value=val_obs)
                    new_values = {
                        "cantidad": nueva_cantidad, "fecha": nueva_fecha,
                        "solicitante": nuevo_solic, "personal": nuevo_pers,
                        "observacion": nueva_obs,
                    }

                elif tipo_mov == "Devolución":
                    st.markdown("##### ↩️ Datos de la devolución")
                    d1, d2 = st.columns(2)
                    nuevo_solic = d1.selectbox("Quién devolvió *", lista_solicitantes, index=solic_idx,
                                               help="Sitio o unidad que realizó la devolución.")
                    nuevo_pers  = d2.selectbox("Personal que recibió la devolución", lista_personal, index=pers_idx)
                    nueva_obs = st.text_area("Observación", value=val_obs)
                    new_values = {
                        "cantidad": nueva_cantidad, "fecha": nueva_fecha,
                        "solicitante": nuevo_solic, "personal": nuevo_pers,
                        "observacion": nueva_obs,
                    }

                elif tipo_mov in ("Corrección entrada", "Corrección salida"):
                    st.markdown("##### 🛠️ Datos de la corrección")
                    nueva_obs = st.text_area(
                        "Justificación de la corrección *", value=val_obs,
                        placeholder="Explique por qué se corrige esta cantidad.",
                    )
                    new_values = {
                        "cantidad": nueva_cantidad, "fecha": nueva_fecha,
                        "observacion": nueva_obs,
                    }

                else:
                    # Tipo no reconocido: formulario genérico
                    nueva_obs = st.text_area("Observación", value=val_obs)
                    new_values = {"cantidad": nueva_cantidad, "fecha": nueva_fecha, "observacion": nueva_obs}

                # ── MOTIVO (siempre obligatorio) ───────────────────────────
                st.markdown("---")
                st.markdown(
                    f"<div style='font-size:12px; color:#94A3B8; margin-bottom:6px;'>"
                    f"👤 Modificación registrada por: <b style='color:#F8FAFC;'>{current_username()}</b></div>",
                    unsafe_allow_html=True,
                )
                motivo = st.text_area(
                    "Motivo de la modificación *",
                    placeholder="Explique brevemente por qué realiza este cambio. Ej: error de digitación, cantidad incorrecta al registrar.",
                )
                submitted_edit = st.form_submit_button("💾 Guardar cambios", use_container_width=True)

            confirm_key_edit = f"edit_{mov_id}"

            if submitted_edit:
                if not motivo:
                    st.error("El motivo de modificación es obligatorio.")
                    return
                set_confirm_pending(confirm_key_edit, {
                    "new_values": new_values, "motivo": motivo,
                    "tipo_mov": tipo_mov, "mov_id": mov_id,
                })
                st.rerun()

            if confirm_pending(confirm_key_edit):
                payload_e = get_confirm_payload(confirm_key_edit)
                nv_e = payload_e.get("new_values", {})
                motivo_e = payload_e.get("motivo", "")
                new_cant  = nv_e.get("cantidad", val_cant)
                diff_cant = float(pd.to_numeric(new_cant, errors="coerce") or 0) - val_cant
                diff_txt  = (f" → <b style='color:#22C55E;'>+{diff_cant:g}</b>" if diff_cant > 0
                             else f" → <b style='color:#EF4444;'>{diff_cant:g}</b>") if diff_cant != 0 else ""
                confirmed_edit = render_confirm_box(
                    key=confirm_key_edit,
                    title=f"¿Confirmar modificación de {mov_id}?",
                    body=(
                        f"<b>Tipo:</b> {tipo_mov} &nbsp;|&nbsp; "
                        f"<b>Producto:</b> {clean_str(selected.get('producto',''))}<br>"
                        f"<b>Cantidad:</b> {val_cant:g} {val_unidad}{diff_txt}<br>"
                        f"<b>Motivo:</b> {motivo_e}<br><br>"
                        "El Kardex y el stock se recalcularán automáticamente."
                    ),
                    confirm_label="💾 Sí, guardar cambios",
                    cancel_label="↩ Cancelar",
                    danger=False,
                )
                if not confirmed_edit:
                    return
                new_values = nv_e
                motivo = motivo_e
                # Guardar cambios con auditoría
                new_df = movimientos.copy()
                mask = new_df["movimiento_id"].astype(str) == mov_id
                audit_rows = []
                cantidad_cambio = False
                for col, val in new_values.items():
                    old_raw  = new_df.loc[mask, col].iloc[0]
                    is_num   = col in ("cantidad", "costo_total")
                    old_str  = str(float(pd.to_numeric(old_raw, errors="coerce") or 0)) if is_num else clean_str(old_raw)
                    new_str  = str(float(pd.to_numeric(val, errors="coerce") or 0)) if is_num else clean_str(val)
                    if old_str != new_str:
                        new_df.loc[mask, col] = val
                        if col == "cantidad":
                            cantidad_cambio = True
                        audit_rows.append({
                            "auditoria_id":   f"AUD-{uuid.uuid4().hex[:12].upper()}",
                            "fecha_evento":   pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "usuario":        current_username(),
                            "rol":            current_role(),
                            "accion":         "EDITAR_MOVIMIENTO",
                            "modulo":         "Movimientos",
                            "registro_id":    mov_id,
                            "campo":          col,
                            "valor_anterior": old_str,
                            "valor_nuevo":    new_str,
                            "motivo":         motivo,
                            "detalle":        f"Edición de {tipo_mov} {mov_id} — campo '{col}'",
                        })
                if not audit_rows:
                    st.info("No se detectaron cambios. Verifique que modificó al menos un campo.")
                    return
                new_df.loc[mask, "modificado_por"]      = current_username()
                new_df.loc[mask, "fecha_modificacion"]  = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                new_df.loc[mask, "motivo_modificacion"] = motivo
                storage.save("Movimientos", ensure_columns(new_df, "Movimientos"))
                append_audit_rows(storage, audit_rows)
                if cantidad_cambio or any(k in new_values for k in ("lote", "marca", "unidad")):
                    sync_data = dict(data)
                    sync_data["Movimientos"] = ensure_columns(new_df, "Movimientos")
                    try:
                        sync_kardex_consolidado_sheet(storage, sync_data)
                    except Exception:
                        pass
                campos_editados = ", ".join(r["campo"] for r in audit_rows)
                set_flash(f"✅ Movimiento {mov_id} actualizado. Campos modificados: {campos_editados}.")
                rerun()

    # ─────────────────────────────────────────────────────────────────────
    # TAB ANULAR: anulación lógica con confirmación doble
    # ─────────────────────────────────────────────────────────────────────
    with tab_anular:
        if not user_has_permission(data, "anular_movimientos"):
            st.info("No tiene permiso para anular movimientos. Solicítelo al administrador.")
        elif est_mov.lower() == "anulado":
            st.markdown(
                "<div class='alert-red'>⛔ Este movimiento ya está <b>anulado</b>. "
                "Solo puede consultarse en auditoría; ya no afecta el stock.</div>",
                unsafe_allow_html=True,
            )
        else:
            # Descripción del impacto según tipo
            impacto_texto = {
                "Ingreso":            f"Al anular, se <b>restará {cant_actual:g} {val_unidad}</b> del stock del lote {val_lote}.",
                "Salida":             f"Al anular, se <b>devolverán {cant_actual:g} {val_unidad}</b> al stock del lote {val_lote}.",
                "Devolución":         f"Al anular, se <b>revertirá la devolución</b> de {cant_actual:g} {val_unidad} del lote {val_lote}.",
                "Corrección entrada": f"Al anular, se <b>revertirá la corrección</b> de +{cant_actual:g} {val_unidad}.",
                "Corrección salida":  f"Al anular, se <b>revertirá la corrección</b> de -{cant_actual:g} {val_unidad}.",
            }.get(tipo_mov, f"Se anulará el movimiento de {cant_actual:g} {val_unidad}.")

            st.markdown(
                f"""<div class='alert-orange' style='margin-bottom:14px;'>
                    ⚠️ <b>Impacto de la anulación:</b> {impacto_texto}<br>
                    <span style='font-size:12px;'>El movimiento <b>no se elimina</b>: queda con estado Anulado
                    en la base y en auditoría. El stock y Kardex se recalculan automáticamente.</span>
                </div>""",
                unsafe_allow_html=True,
            )

            confirm_key_anular = f"anular_{mov_id}"

            # Step 1: Motivo form
            if not confirm_pending(confirm_key_anular):
                with st.form(f"frm_anular_{mov_id}"):
                    st.markdown(
                        f"**Movimiento:** `{mov_id}` &nbsp;|&nbsp; **Tipo:** {tipo_mov} "
                        f"&nbsp;|&nbsp; **Producto:** {clean_str(selected.get('producto',''))} "
                        f"&nbsp;|&nbsp; **Cantidad:** {cant_actual:g} {val_unidad}"
                    )
                    motivo_anulacion_input = st.text_area(
                        "Motivo de la anulación *",
                        placeholder="Explique el motivo. Ej: cantidad incorrecta, movimiento duplicado, error de lote.",
                        key=f"motivo_anular_{mov_id}",
                    )
                    step1 = st.form_submit_button("🚫 Solicitar anulación →", use_container_width=True)
                if step1:
                    if not motivo_anulacion_input:
                        st.error("El motivo de anulación es obligatorio antes de continuar.")
                    else:
                        set_confirm_pending(confirm_key_anular, {
                            "motivo": motivo_anulacion_input,
                            "mov_id": mov_id,
                        })
                        st.rerun()

            # Step 2: Confirmation card
            else:
                payload = get_confirm_payload(confirm_key_anular)
                motivo_anulacion = payload.get("motivo", "")
                confirmed = render_confirm_box(
                    key=confirm_key_anular,
                    title=f"¿Confirmar anulación de {mov_id}?",
                    body=(
                        f"Esta acción marcará el movimiento como <b>Anulado</b>. "
                        f"No se elimina, pero <b>dejará de afectar el stock</b>.<br><br>"
                        f"<b>Movimiento:</b> {mov_id} — {tipo_mov}<br>"
                        f"<b>Producto:</b> {clean_str(selected.get('producto',''))}<br>"
                        f"<b>Cantidad:</b> {cant_actual:g} {val_unidad}<br>"
                        f"<b>Motivo:</b> {motivo_anulacion}"
                    ),
                    confirm_label="🚫 Sí, anular definitivamente",
                    cancel_label="↩ Cancelar",
                    danger=True,
                )
                if confirmed:
                    new_df = movimientos.copy()
                    mask = new_df["movimiento_id"].astype(str) == mov_id
                    new_df.loc[mask, "estado_movimiento"] = "Anulado"
                    new_df.loc[mask, "anulado_por"]       = current_username()
                    new_df.loc[mask, "fecha_anulacion"]   = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_df.loc[mask, "motivo_anulacion"]  = motivo_anulacion
                    storage.save("Movimientos", ensure_columns(new_df, "Movimientos"))
                    append_audit(storage, "ANULAR_MOVIMIENTO", "Movimientos", mov_id,
                                 "estado_movimiento", "Vigente", "Anulado",
                                 motivo_anulacion, f"Anulación lógica de {tipo_mov} {mov_id}")
                    sync_data = dict(data)
                    sync_data["Movimientos"] = ensure_columns(new_df, "Movimientos")
                    try:
                        sync_kardex_consolidado_sheet(storage, sync_data)
                    except Exception:
                        pass
                    set_flash(f"Movimiento {mov_id} anulado. El stock fue recalculado.", "success")
                    rerun()



def admin_permissions_tab(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Permisos especiales por usuario", "Active o niegue permisos específicos sin cambiar el rol general.")
    usuarios = ensure_columns(data.get("Usuarios", pd.DataFrame()), "Usuarios")
    permisos = ensure_columns(data.get("Permisos_Usuarios", pd.DataFrame()), "Permisos_Usuarios")
    usuarios_activos = usuarios[active_mask(usuarios)].copy()
    if usuarios_activos.empty:
        st.info("No hay usuarios activos.")
        return
    user_labels = usuarios_activos.apply(lambda r: f"{clean_str(r['usuario'])} | {clean_str(r['nombre'])} | Rol: {clean_str(r['rol'])}", axis=1).tolist()
    selected_label = st.selectbox("Usuario", user_labels, key="perm_user")
    selected_user = clean_str(usuarios_activos.iloc[user_labels.index(selected_label)]["usuario"])
    current = permisos[(permisos["usuario"].astype(str).str.lower().str.strip() == selected_user.lower()) & (permisos["estado"].astype(str).str.lower().str.strip().isin(["activo", "sí", "si", "true", "1", ""]))]
    current_yes = set(current[current["valor"].astype(str).str.lower().str.strip().isin(["sí", "si", "true", "1", "yes"])] ["permiso"].astype(str).tolist())
    options = list(PERMISSION_LABELS.keys())
    selected_perms = st.multiselect(
        "Permisos habilitados para este usuario",
        options,
        default=[p for p in options if p in current_yes],
        format_func=lambda p: PERMISSION_LABELS.get(p, p),
    )
    st.caption("Los permisos seleccionados se guardan como excepciones explícitas. Si desmarca un permiso, quedará negado para ese usuario.")
    if st.button("💾 Guardar permisos del usuario", use_container_width=True):
        now = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        remaining = permisos[permisos["usuario"].astype(str).str.lower().str.strip() != selected_user.lower()].copy()
        rows = []
        for perm in options:
            rows.append({
                "usuario": selected_user,
                "permiso": perm,
                "valor": "Sí" if perm in selected_perms else "No",
                "fecha_asignacion": now,
                "asignado_por": current_username(),
                "estado": "Activo",
            })
        new_df = pd.concat([remaining, pd.DataFrame(rows)], ignore_index=True)
        storage.save("Permisos_Usuarios", ensure_columns(new_df, "Permisos_Usuarios"))
        append_audit(storage, "ASIGNAR_PERMISOS", "Permisos_Usuarios", selected_user, "permisos", ", ".join(sorted(current_yes)), ", ".join(sorted(selected_perms)), "Asignación desde administración", f"Permisos actualizados para {selected_user}")
        set_flash("Permisos actualizados correctamente.")
        rerun()
    st.dataframe(permisos, use_container_width=True, hide_index=True)


def admin_audit_tab(data: Dict[str, pd.DataFrame]) -> None:
    card_start("Auditoría del sistema", "Registro de creación, edición, anulación y cambios de permisos.")
    aud = ensure_columns(data.get("Auditoria_Cambios", pd.DataFrame()), "Auditoria_Cambios")
    if aud.empty:
        st.info("Aún no hay eventos de auditoría registrados.")
        return
    c1, c2, c3 = st.columns(3)
    accion = c1.multiselect("Acción", sorted(aud["accion"].dropna().astype(str).unique().tolist()))
    usuario = c2.multiselect("Usuario", sorted(aud["usuario"].dropna().astype(str).unique().tolist()))
    modulo = c3.multiselect("Módulo", sorted(aud["modulo"].dropna().astype(str).unique().tolist()))
    view = aud.copy()
    if accion:
        view = view[view["accion"].isin(accion)]
    if usuario:
        view = view[view["usuario"].isin(usuario)]
    if modulo:
        view = view[view["modulo"].isin(modulo)]
    st.dataframe(view.sort_values("fecha_evento", ascending=False), use_container_width=True, hide_index=True)

def page_admin(storage, data: Dict[str, pd.DataFrame], mode: str) -> None:
    section_header("🔐 Administración", "Gestión de usuarios, seguridad PATH automática y diagnóstico de conexión a la base.")
    if not require_admin():
        return

    tab1, tab_perm, tab_crud, tab_audit, tab2, tab3 = st.tabs(["Usuarios", "Permisos", "Movimientos", "Auditoría", "Seguridad PATH", "Diagnóstico"])

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

    with tab_perm:
        admin_permissions_tab(storage, data)

    with tab_crud:
        render_movement_crud_controls(storage, data)

    with tab_audit:
        admin_audit_tab(data)

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
        st.warning(
            "Sugerencia V25: verifique que en Secrets esté AUTO_MIGRATE_GOOGLE_SHEETS = true. "
            "Esta versión crea automáticamente las hojas nuevas de auditoría/permisos si faltan."
        )
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
