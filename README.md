# Kyte Checkout Pro para Render

Proyecto Flask POS/Checkout inspirado en el flujo de Kyte Web, listo para GitHub + Render.

## Render
Build Command:
```bash
pip install -r requirements.txt
```
Start Command:
```bash
gunicorn --bind 0.0.0.0:$PORT app:app --workers 1 --threads 4 --timeout 120
```

## Usuarios demo
- admin1 / admin123
- vendedor1 / venta123
