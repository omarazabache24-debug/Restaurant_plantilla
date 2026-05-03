# Negocio 2.0 - AORIX PRO

Aplicación Flask lista para GitHub y Render.

## Mejoras incluidas
- Creación y actualización de usuarios desde **Usuarios / Admin**.
- Control por roles:
  - **ADMIN**: acceso total.
  - Usuarios no ADMIN: solo **Venta**, **Pedido**, **Cierre** y **Salir**.
- Interfaz mejorada con logo AORIX, pestañas superiores y menú lateral.
- Vista responsive para celular: menú compacto, botones grandes y tablas con scroll.
- Persistencia local/Render usando SQLite en `/data` si está disponible.

## Usuarios iniciales
- admin / admin123
- caja / caja123
- mozo / mozo123

## Ejecutar local
```bash
pip install -r requirements.txt
python app.py
```

## Render
Subir el repositorio a GitHub y crear Web Service en Render usando:
```bash
gunicorn app:app
```

## Corrección Render
Este ZIP ya viene con `requirements.txt`, `runtime.txt`, `Procfile`, `render.yaml` y `app.py` en la RAÍZ del proyecto. No subir la carpeta contenedora; subir directamente estos archivos al repositorio.
