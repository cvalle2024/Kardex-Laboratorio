# V27 — Corrección de arranque modular en Streamlit Cloud

Esta versión corrige el arranque del sistema modular.

## Problema corregido

En algunos despliegues de Streamlit Cloud, el archivo `app.py` puede ejecutarse sin reconocer de inmediato la carpeta `kardex_app/` como paquete local, generando error en:

```python
from kardex_app.main import main
```

## Correcciones aplicadas

- `app.py` ahora agrega explícitamente la raíz del proyecto a `sys.path`.
- Si falta la carpeta `kardex_app/`, el sistema muestra un mensaje técnico más claro.
- Se agregó nuevamente la navegación hacia `9️⃣ Importar Kardex anterior` en `kardex_app/main.py`.
- Se eliminaron carpetas `__pycache__` del paquete distribuible.

## Estructura obligatoria en GitHub

La raíz del repositorio debe verse así:

```text
app.py
requirements.txt
kardex_app/
assets/
data/
.streamlit/
README.md
```

No debe subirse solo `app.py`. La carpeta `kardex_app/` completa es obligatoria.
