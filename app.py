"""Punto de entrada del Sistema Kardex PRO.

Este archivo debe permanecer en la raíz del repositorio. La carpeta modular
`kardex_app/` debe estar junto a este archivo. Como respaldo, este arranque
busca automáticamente la carpeta `kardex_app` si por error quedó dentro de una
subcarpeta al subir el proyecto a GitHub.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent


def _add_path(path: Path) -> None:
    """Agrega una ruta a sys.path si todavía no está registrada."""
    path_str = str(path.resolve())
    if path_str not in sys.path:
        sys.path.insert(0, path_str)


def _find_kardex_package_root(base_dir: Path) -> Path | None:
    """Busca la carpeta padre que contiene kardex_app/main.py.

    Esto ayuda cuando en GitHub quedó una estructura como:

        app.py
        sistema_kardex_streamlit_v28_modular_autodetect/
            kardex_app/
                main.py

    En producción lo correcto sigue siendo tener:

        app.py
        kardex_app/
            main.py
    """
    direct_pkg = base_dir / "kardex_app" / "main.py"
    if direct_pkg.exists():
        return base_dir

    # Búsqueda limitada para evitar recorrer carpetas grandes innecesariamente.
    ignored_dirs = {".git", "__pycache__", ".streamlit", "data", "assets", ".venv", "venv"}
    max_depth = 3
    base_depth = len(base_dir.parts)

    for current_root, dirs, files in os.walk(base_dir):
        current = Path(current_root)
        depth = len(current.parts) - base_depth
        if depth > max_depth:
            dirs[:] = []
            continue

        dirs[:] = [d for d in dirs if d not in ignored_dirs]
        if current.name == "kardex_app" and "main.py" in files:
            return current.parent

    return None


package_root = _find_kardex_package_root(ROOT_DIR)
if package_root is not None:
    _add_path(package_root)
else:
    _add_path(ROOT_DIR)

try:
    from kardex_app.main import main
except ModuleNotFoundError as exc:
    if exc.name == "kardex_app":
        raise ModuleNotFoundError(
            "No se encontró el paquete modular 'kardex_app'. Revise GitHub: debe existir "
            "la carpeta 'kardex_app/' completa, con archivos como 'main.py', 'common.py', "
            "'auth.py', 'pages/', 'services/' y 'storage/', en la misma raíz donde está app.py. "
            "Si el ZIP quedó como una carpeta dentro del repositorio, copie el contenido de esa "
            "carpeta a la raíz del repositorio y haga commit nuevamente."
        ) from exc
    raise


if __name__ == "__main__":
    main()
