"""Punto de entrada del Sistema Kardex PRO.

Este archivo debe permanecer en la raíz del repositorio, junto a la carpeta
`kardex_app/`. En Streamlit Cloud seleccione este archivo como Main file path:

    app.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

try:
    from kardex_app.main import main
except ModuleNotFoundError as exc:
    if exc.name == "kardex_app":
        raise ModuleNotFoundError(
            "No se encontró el paquete 'kardex_app'. Verifique que en GitHub se haya subido "
            "la carpeta completa 'kardex_app/' en la misma raíz donde está app.py. "
            "No suba solo app.py; debe subir toda la estructura modular del sistema."
        ) from exc
    raise


if __name__ == "__main__":
    main()
