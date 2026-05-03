# AORIX POS / Kyte Clone Pro Render

Proyecto Flask listo para GitHub y Render.

## Usuarios demo
- admin1 / admin123
- admin2 / admin123
- vendedor1 / venta123

## Render
Build command: `pip install -r requirements.txt`
Start command: `gunicorn app:app --bind 0.0.0.0:$PORT`

Se agregó `.python-version` porque Render ahora recomienda fijar la versión de Python con ese archivo o con la variable `PYTHON_VERSION`.
