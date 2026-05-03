# AORIX / KYTE POS PRO - Render

Proyecto Flask listo para GitHub + Render.

## Usuarios demo
- admin1 / admin123
- admin2 / admin123
- vendedor1 / venta123

## Render
Build Command:
```bash
pip install -r requirements.txt
```

Start Command:
```bash
gunicorn --bind 0.0.0.0:$PORT app:app --workers 1 --threads 4 --timeout 120
```

## Importante
Sube TODO el contenido del ZIP: `app.py`, `templates/`, `static/`, `requirements.txt`, `Procfile`, `.python-version`.
No subas solo `app.py`.
