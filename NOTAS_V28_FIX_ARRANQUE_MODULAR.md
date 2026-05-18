# V28 - Fix de arranque modular

Esta versión mejora el archivo `app.py` para detectar automáticamente la carpeta `kardex_app/` si por error quedó dentro de una subcarpeta al subir el sistema a GitHub.

## Estructura ideal en GitHub

```text
app.py
requirements.txt
kardex_app/
assets/
data/
.streamlit/
README.md
```

## Problema corregido

Error típico:

```text
ModuleNotFoundError: No se encontró el paquete 'kardex_app'
```

Este error ocurre cuando Streamlit ejecuta `app.py`, pero no encuentra la carpeta modular `kardex_app/`.

## Recomendación

Aunque V28 busca la carpeta automáticamente, la estructura profesional recomendada sigue siendo subir el contenido del ZIP a la raíz del repositorio, no dejar todo dentro de una carpeta adicional.
