# AORIX POS Pro - Render

## Render
Build Command:
```bash
pip install -r requirements.txt
```
Start Command:
```bash
gunicorn --bind 0.0.0.0:$PORT app:app --workers 1 --threads 4 --timeout 120
```

Usuarios demo:
- admin1 / admin123
- admin2 / admin123
- vendedor1 / venta123

Esta versión incluye plantillas internas dentro de app.py para evitar definitivamente `TemplateNotFound: error.html`.
