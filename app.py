from __future__ import annotations

import hashlib
import secrets
import string
import uuid
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

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
        "proveedor", "solicitante", "personal", "fecha_elaboracion", "fecha_vencimiento",
        "unidad", "cantidad", "costo_total", "observacion", "usuario_registro", "fecha_registro"
    ],
    "Config": ["clave", "valor"],
}

TIPOS_MOVIMIENTO = ["Ingreso", "Salida", "Devolución", "Ajuste entrada", "Ajuste salida"]
TIPOS_POSITIVOS = {"Ingreso", "Devolución", "Ajuste entrada"}
TIPOS_NEGATIVOS = {"Salida", "Ajuste salida"}
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
    "Config": pd.DataFrame([
        {"clave": "dias_alerta_global", "valor": "90"},
        {"clave": "moneda", "valor": "L"},
        {"clave": "institucion", "valor": "Proyecto VIHCA"},
        {"clave": "path_verificacion", "valor": "DINAMICO"},
        {"clave": "path_ttl_minutos", "valor": str(PATH_TTL_MINUTES)},
        {"clave": "version_sistema", "valor": "7.0"},
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
                for sheet, df in INITIAL_DATA.items():
                    ensure_columns(df, sheet).to_excel(writer, index=False, sheet_name=sheet)
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

    def append_row(self, sheet: str, row: dict) -> None:
        df = self.load(sheet)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        self.save(sheet, df)


class GoogleSheetsStorage:
    def __init__(self):
        try:
            import gspread
            from google.oauth2.service_account import Credentials
        except ImportError as exc:
            raise RuntimeError("Falta instalar gspread y google-auth para usar Google Sheets.") from exc

        sheet_id = safe_secret("GOOGLE_SHEET_ID", "")
        creds_info = safe_secret("gcp_service_account", None)
        if not sheet_id or not creds_info:
            raise RuntimeError("No se encontró GOOGLE_SHEET_ID o gcp_service_account en los secretos de Streamlit.")

        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = Credentials.from_service_account_info(dict(creds_info), scopes=scopes)
        client = gspread.authorize(credentials)
        self.book = client.open_by_key(sheet_id)
        self.ensure_database()

    def ensure_database(self) -> None:
        existing = [ws.title for ws in self.book.worksheets()]
        for sheet, columns in SHEET_COLUMNS.items():
            if sheet not in existing:
                ws = self.book.add_worksheet(title=sheet, rows=1000, cols=max(len(columns), 10))
                ws.update([columns])
                if not INITIAL_DATA[sheet].empty:
                    data = normalize_for_sheet(INITIAL_DATA[sheet])
                    ws.append_rows(data.values.tolist(), value_input_option="USER_ENTERED")
            else:
                ws = self.book.worksheet(sheet)
                values = ws.get_all_values()
                if not values:
                    ws.update([columns])

    def load(self, sheet: str) -> pd.DataFrame:
        ws = self.book.worksheet(sheet)
        records = ws.get_all_records()
        df = pd.DataFrame(records)
        return ensure_columns(df, sheet)

    def save(self, sheet: str, df: pd.DataFrame) -> None:
        ws = self.book.worksheet(sheet)
        df = normalize_for_sheet(ensure_columns(df, sheet))
        ws.clear()
        values = [SHEET_COLUMNS[sheet]] + df.values.tolist()
        ws.update(values, value_input_option="USER_ENTERED")

    def append_row(self, sheet: str, row: dict) -> None:
        ws = self.book.worksheet(sheet)
        df = pd.DataFrame([row])
        df = normalize_for_sheet(ensure_columns(df, sheet))
        ws.append_row(df.iloc[0].tolist(), value_input_option="USER_ENTERED")


def get_storage():
    use_gsheets = as_bool(safe_secret("USE_GOOGLE_SHEETS", False), False)
    if use_gsheets:
        try:
            return GoogleSheetsStorage(), "Google Sheets"
        except Exception as exc:
            st.error(f"No se pudo conectar a Google Sheets. Se usará Excel local. Detalle: {exc}")
    return LocalExcelStorage(DB_FILE), "Excel local"

# ============================================================
# CÁLCULOS DE KARDEX
# ============================================================
def load_all(storage) -> Dict[str, pd.DataFrame]:
    return {sheet: storage.load(sheet) for sheet in SHEET_COLUMNS.keys()}


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
    return [
        "estado", "producto_id", "producto", "marca", "lote", "unidad",
        "fecha_ingreso", "proveedor_ingreso", "fecha_elaboracion", "fecha_vencimiento",
        "entrada_total", "salida_total", "saldo_actual", "porcentaje_consumido",
        "numero_salidas", "fecha_ultima_salida", "ultimo_entregado_a", "ultimo_personal_entrega",
        "detalle_salidas", "observacion_ingreso", "dias_para_vencer", "stock_minimo",
    ]


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
                fecha_elaboracion=("fecha_elaboracion", first_non_empty),
                observacion_ingreso=("observacion", first_non_empty),
            )
            .reset_index()
        )
    else:
        ingreso_info = pd.DataFrame(columns=group_cols + ["fecha_ingreso", "proveedor_ingreso", "fecha_elaboracion", "observacion_ingreso"])

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
    for col in ["fecha_ingreso", "proveedor_ingreso", "fecha_elaboracion", "observacion_ingreso", "fecha_ultima_salida", "ultimo_entregado_a", "ultimo_personal_entrega", "detalle_salidas"]:
        if col not in kardex.columns:
            kardex[col] = ""
        kardex[col] = kardex[col].fillna("")
    kardex["numero_salidas"] = to_number(kardex.get("numero_salidas", pd.Series(dtype=float))).astype(int)
    kardex = kardex.sort_values(["estado", "producto", "fecha_vencimiento", "lote"], na_position="last")
    return kardex[columns]


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
            <div style="margin-top:12px"><span class='pill'>Base activa: {storage_mode}</span>{user_badge}<span class='pill'>Versión 6.0</span></div>
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
        usuario = st.text_input("Usuario", placeholder="Ejemplo: admin")
        password = st.text_input("Contraseña", type="password", placeholder="Ingrese su contraseña")
        path_key = st.text_input(
            "Pegar PATH generado",
            placeholder="Copie el código generado arriba y péguelo aquí",
            help="Debe coincidir exactamente con el código PATH generado en esta pantalla.",
        )
        submitted = st.form_submit_button("Ingresar al sistema", use_container_width=True)

    st.markdown(
        "<div class='field-note'>Usuario inicial: <b>admin</b> | Contraseña inicial: <b>admin123</b>. "
        "El PATH ya no es fijo: se genera automáticamente en esta pantalla y se copia/pega para validar el acceso.</div>",
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
        st.success("Acceso autorizado.")
        rerun()
    else:
        if not valid_password:
            st.error("Contraseña incorrecta.")
        elif not valid_path:
            st.error("El PATH no coincide con el código generado. Copie el código mostrado arriba y péguelo nuevamente.")
    return False

def logout() -> None:
    for key in ["auth_ok", "usuario", "nombre_usuario", "rol", "login_path_code", "login_path_created_at"]:
        st.session_state.pop(key, None)
    rerun()


def require_admin() -> bool:
    if st.session_state.get("rol") != "Administrador":
        st.warning("Este módulo requiere rol de Administrador.")
        return False
    return True


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
    st.markdown(flujo_html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    st.markdown("#### Accesos rápidos según la ruta")
    a, b, c, d = st.columns(4)
    if a.button("📚 Catálogos base", use_container_width=True):
        nav_go(PAGE_CATALOGOS)
    if b.button("🔐 Administración", use_container_width=True):
        nav_go(PAGE_ADMIN)
    if c.button("🔁 Registrar movimientos", use_container_width=True):
        nav_go(PAGE_MOVIMIENTOS)
    if d.button("📊 Reportes", use_container_width=True):
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


def page_movimiento(storage, data: Dict[str, pd.DataFrame], stock: pd.DataFrame) -> None:
    section_header(
        "➕ Registrar movimiento",
        "Formulario guiado: primero seleccione el producto/lote y luego complete cantidades, proveedor, solicitante y responsable."
    )

    productos = ensure_columns(data["Productos"], "Productos")
    proveedores = ensure_columns(data["Proveedores"], "Proveedores")
    solicitantes = ensure_columns(data["Solicitantes"], "Solicitantes")
    personal_df = ensure_columns(data["Personal"], "Personal")

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

    tipo = st.radio(
        "Tipo de movimiento",
        TIPOS_MOVIMIENTO,
        horizontal=True,
        help="Seleccione si está ingresando, entregando, recibiendo devolución o ajustando inventario."
    )
    is_salida = tipo in TIPOS_NEGATIVOS

    selected_stock_row = None
    proveedor = ""
    solicitante = ""
    personal = ""
    fecha_elaboracion = ""
    costo_total = 0.0

    # ========================================================
    # 1) Producto primero: aquí se toma la información del catálogo
    # ========================================================
    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    st.markdown("#### 1) Producto / marca / lote")

    if is_salida:
        disponibles = stock[stock["stock_actual"] > 0].copy() if not stock.empty else pd.DataFrame()
        if disponibles.empty:
            st.error("No hay lotes con stock disponible para registrar salidas.")
            st.markdown("</div>", unsafe_allow_html=True)
            return

        disponibles = disponibles.sort_values(["producto", "fecha_vencimiento", "lote"], na_position="last")
        cprod, clote = st.columns([1, 1.55])
        filtro_producto = cprod.selectbox(
            "Producto con existencia",
            ["Todos"] + sorted(disponibles["producto"].dropna().astype(str).unique().tolist()),
            help="Primero elija el producto. Luego seleccione el lote específico que entregará."
        )
        if filtro_producto != "Todos":
            disponibles = disponibles[disponibles["producto"].astype(str) == filtro_producto]

        if disponibles.empty:
            st.warning("No hay lotes disponibles para el producto seleccionado.")
            st.markdown("</div>", unsafe_allow_html=True)
            return

        disponibles["label"] = disponibles.apply(
            lambda r: f"{r['producto']} | Marca: {r['marca']} | Lote: {r['lote']} | Vence: {r['fecha_vencimiento']} | Stock: {float(r['stock_actual']):,.0f} {r['unidad']}",
            axis=1,
        )
        label = clote.selectbox("Lote disponible para salida", disponibles["label"].tolist())
        selected_stock_row = disponibles[disponibles["label"] == label].iloc[0]

        producto_id = clean_str(selected_stock_row["producto_id"])
        producto = clean_str(selected_stock_row["producto"])
        marca = clean_str(selected_stock_row["marca"])
        lote = clean_str(selected_stock_row["lote"])
        fecha_vencimiento = format_date(selected_stock_row["fecha_vencimiento"])
        unidad = clean_str(selected_stock_row["unidad"])
        stock_disponible = float(selected_stock_row["stock_actual"])
        prod_row = _catalog_row_by_id(producto_id)
        if clean_str(prod_row.get("nombre_producto", "")) == "":
            prod_row["nombre_producto"] = producto
            prod_row["marca_default"] = marca
            prod_row["unidad_default"] = unidad
        _product_card(prod_row, "Ficha del producto seleccionado")
        st.info(f"Stock disponible para este lote: {stock_disponible:,.0f} {unidad}. El sistema no permitirá registrar una salida mayor a este saldo.")

    else:
        productos = productos.sort_values(["nombre_producto", "codigo_producto"], na_position="last")
        labels = productos.apply(
            lambda r: f"{clean_str(r['nombre_producto'])} | Código: {clean_str(r['codigo_producto']) or 'Sin código'} | Marca: {clean_str(r['marca_default']) or 'Sin marca'}",
            axis=1,
        ).tolist()
        producto_label = st.selectbox(
            "Producto / reactivo / insumo registrado en catálogo",
            labels,
            help="Al seleccionar el producto, el sistema toma automáticamente marca, unidad, categoría, stock mínimo y días de alerta registrados en el catálogo."
        )
        prod_row = productos.iloc[labels.index(producto_label)]
        producto_id = clean_str(prod_row["producto_id"])
        producto = clean_str(prod_row["nombre_producto"])
        marca_default = clean_str(prod_row.get("marca_default", ""))
        unidad_default = clean_str(prod_row.get("unidad_default", ""))
        _product_card(prod_row)
        st.caption("La marca y la unidad se cargan desde el catálogo. Puede modificarlas solo si este lote específico viene con una presentación diferente.")

    st.markdown("</div>", unsafe_allow_html=True)

    # ========================================================
    # 2) Datos del movimiento: datos generales y lote
    # ========================================================
    st.markdown("<div class='form-card'>", unsafe_allow_html=True)
    with st.form("frm_movimiento", clear_on_submit=False):
        st.markdown("#### 2) Datos del movimiento")
        col_a, col_b, col_c = st.columns([1, 1, 1])
        fecha = col_a.date_input("Fecha del movimiento", value=TODAY.date())
        usuario = col_b.text_input("Usuario que registra", value=st.session_state.get("nombre_usuario", "Usuario"))
        cantidad = col_c.number_input("Cantidad", min_value=0.0, step=1.0, format="%.2f")

        if is_salida:
            st.markdown("#### 3) Datos automáticos del lote seleccionado")
            l1, l2, l3, l4 = st.columns(4)
            l1.text_input("Producto", value=producto, disabled=True)
            l2.text_input("Marca", value=marca, disabled=True)
            l3.text_input("Lote", value=lote, disabled=True)
            l4.text_input("Unidad", value=unidad, disabled=True)
            v1, v2 = st.columns([1, 2])
            v1.text_input("Fecha de vencimiento", value=fecha_vencimiento, disabled=True)
            v2.text_input("Stock disponible", value=f"{stock_disponible:,.0f} {unidad}", disabled=True)
            fecha_elaboracion = ""
            costo_total = 0.0
        else:
            st.markdown("#### 3) Datos del lote / presentación")
            unidades = sorted(set([u for u in UNIDADES_DEFAULT + [unidad_default] if clean_str(u)]))
            unidad_index = unidades.index(unidad_default) if unidad_default in unidades else 0
            col1, col2, col3 = st.columns([1, 1, 1])
            marca = col1.text_input(
                "Marca",
                value=marca_default,
                key=f"mov_marca_{tipo}_{producto_id}",
                help="Se carga desde el catálogo del producto."
            )
            lote = col2.text_input("Lote *", placeholder="Ejemplo: AB-2026-001", key=f"mov_lote_{tipo}_{producto_id}")
            unidad = col3.selectbox(
                "Unidad",
                unidades,
                index=unidad_index,
                key=f"mov_unidad_{tipo}_{producto_id}",
                help="Se carga desde el catálogo del producto."
            )
            col4, col5, col6 = st.columns([1, 1, 1])
            fecha_elaboracion_dt = col4.date_input("Fecha de elaboración", value=TODAY.date(), key=f"mov_elab_{tipo}_{producto_id}")
            fecha_vencimiento_dt = col5.date_input("Fecha de vencimiento", value=(TODAY + pd.Timedelta(days=365)).date(), key=f"mov_venc_{tipo}_{producto_id}")
            costo_total = col6.number_input("Costo total", min_value=0.0, step=1.0, format="%.2f", key=f"mov_costo_{tipo}_{producto_id}")
            fecha_elaboracion = fecha_elaboracion_dt.strftime("%Y-%m-%d")
            fecha_vencimiento = fecha_vencimiento_dt.strftime("%Y-%m-%d")

        st.markdown("#### 4) Origen / destino / responsable")
        colx, coly, colz = st.columns([1, 1, 1])
        if tipo == "Ingreso":
            proveedor = colx.selectbox("Proveedor *", [""] + proveedores["proveedor"].dropna().astype(str).sort_values().tolist())
            personal = coly.selectbox("Personal que recibe", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            colz.text_input("Destino", value="Bodega / Inventario", disabled=True)
        elif tipo in ["Salida", "Ajuste salida"]:
            solicitante = colx.selectbox("Entregado a / solicitante *", [""] + solicitantes["unidad_solicitante"].dropna().astype(str).sort_values().tolist())
            personal = coly.selectbox("Personal que entrega", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            colz.text_input("Origen", value="Bodega / Inventario", disabled=True)
        elif tipo == "Devolución":
            solicitante = colx.selectbox("Quién devuelve", [""] + solicitantes["unidad_solicitante"].dropna().astype(str).sort_values().tolist())
            personal = coly.selectbox("Personal que recibe", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            colz.text_input("Destino", value="Bodega / Inventario", disabled=True)
        else:
            personal = colx.selectbox("Personal responsable", [""] + personal_df["nombre"].dropna().astype(str).sort_values().tolist())
            coly.text_input("Tipo de ajuste", value=tipo, disabled=True)
            colz.text_input("Destino", value="Bodega / Inventario", disabled=True)

        st.markdown("#### 5) Observación")
        observacion = st.text_area("Observación / justificación", placeholder="Detalle de factura, solicitud, ajuste, entrega o comentario relevante.")
        submitted = st.form_submit_button("💾 Guardar movimiento", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    if submitted:
        if st.session_state.get("rol") == "Consulta":
            st.error("El rol Consulta no puede registrar movimientos.")
            return
        if cantidad <= 0:
            st.error("La cantidad debe ser mayor a cero.")
            return
        if is_salida and selected_stock_row is not None and cantidad > float(selected_stock_row["stock_actual"]):
            st.error("La salida no puede ser mayor al stock disponible del lote seleccionado.")
            return
        if not lote:
            st.error("El lote es obligatorio.")
            return
        if tipo == "Ingreso" and not proveedor:
            st.error("Para ingresos debe seleccionar proveedor.")
            return
        if tipo in ["Salida", "Ajuste salida"] and not solicitante:
            st.error("Para salidas debe seleccionar solicitante/unidad.")
            return

        mov_df = data["Movimientos"]
        row = {
            "movimiento_id": next_code("MOV", mov_df, "movimiento_id", 6),
            "fecha": pd.to_datetime(fecha).strftime("%Y-%m-%d"),
            "tipo_movimiento": tipo,
            "producto_id": producto_id,
            "producto": producto,
            "marca": marca,
            "lote": lote,
            "proveedor": proveedor,
            "solicitante": solicitante,
            "personal": personal,
            "fecha_elaboracion": fecha_elaboracion,
            "fecha_vencimiento": fecha_vencimiento,
            "unidad": unidad,
            "cantidad": cantidad,
            "costo_total": costo_total,
            "observacion": observacion,
            "usuario_registro": usuario,
            "fecha_registro": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        storage.append_row("Movimientos", row)
        st.success("Movimiento guardado correctamente.")
        rerun()


def page_kardex_consolidado(kardex: pd.DataFrame) -> None:
    section_header(
        "📋 Kardex consolidado por lote",
        "Una fila por producto/lote: ingreso, salida acumulada, último destinatario, fecha de entrega y saldo actual."
    )
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
    with st.form("frm_producto"):
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
        st.success("Producto guardado correctamente.")
        rerun()


def provider_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo proveedor", "Formulario ordenado por datos generales, contacto y ubicación.")
    with st.form("frm_proveedor"):
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
        st.success("Proveedor guardado correctamente.")
        rerun()


def requester_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo solicitante / unidad", "Registre unidades, áreas o sitios que pueden solicitar productos.")
    with st.form("frm_solicitante"):
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
        st.success("Solicitante guardado correctamente.")
        rerun()


def staff_form(storage, data: Dict[str, pd.DataFrame]) -> None:
    card_start("Nuevo personal", "Usuarios operativos que reciben, entregan o registran movimientos.")
    with st.form("frm_personal"):
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
        st.success("Personal guardado correctamente.")
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
        storage.save(sheet, ensure_columns(pd.DataFrame(edited), sheet))
        st.success(f"{title} actualizado correctamente.")
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
                storage.save("Productos", ensure_columns(prod_final, "Productos"))
                storage.save("Movimientos", ensure_columns(mov_final, "Movimientos"))
                st.success("Importación completada.")
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
        with st.form("frm_usuario"):
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
                    st.success("Usuario creado correctamente.")
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
        card_start("Diagnóstico de base y estructura", "Verifica que existan las hojas obligatorias y muestra el PATH de almacenamiento activo.")
        expected = list(SHEET_COLUMNS.keys())
        status_rows = []
        for sheet in expected:
            try:
                df = storage.load(sheet)
                status_rows.append({"Hoja": sheet, "Estado": "OK", "Filas": len(df), "Columnas esperadas": len(SHEET_COLUMNS[sheet])})
            except Exception as exc:
                status_rows.append({"Hoja": sheet, "Estado": f"Error: {exc}", "Filas": 0, "Columnas esperadas": len(SHEET_COLUMNS[sheet])})
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)
        if mode == "Excel local":
            st.info(f"Ruta de base local: {DB_FILE.resolve()}")
        else:
            st.info("Base conectada a Google Sheets mediante secretos de Streamlit.")

# ============================================================
# APP PRINCIPAL
# ============================================================
def main() -> None:
    apply_theme()
    storage, mode = get_storage()
    data = load_all(storage)

    if not render_login(storage, data, mode):
        return

    stock = calcular_stock(data["Movimientos"], data["Productos"])
    kardex = calcular_kardex_consolidado(data["Movimientos"], data["Productos"])
    hero(mode, st.session_state.get("nombre_usuario", ""))

    if "page" not in st.session_state or st.session_state["page"] not in NAV_PAGES:
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
            NAV_PAGES,
            index=NAV_PAGES.index(st.session_state["page"]),
            label_visibility="collapsed",
        )
        st.session_state["page"] = page
        st.divider()
        st.markdown("**Estructura lógica**")
        st.caption("1. Acceso → 2. Catálogos → 3. Administración → 4. Movimientos → 5. Kardex/Stock → 6. Reportes")
        st.divider()
        if st.button("Cerrar sesión", use_container_width=True):
            logout()
        st.caption("Sistema Kardex PRO: control por lote, vencimiento, stock mínimo y trazabilidad.")

    if page == PAGE_INICIO:
        page_inicio_operativo(data, stock, kardex, mode)
    elif page == PAGE_CATALOGOS:
        page_catalogos(storage, data)
    elif page == PAGE_ADMIN:
        page_admin(storage, data, mode)
    elif page == PAGE_MOVIMIENTOS:
        page_movimiento(storage, data, stock)
    elif page == PAGE_KARDEX:
        page_kardex_consolidado(kardex)
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
