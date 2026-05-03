# AORIX POS Pro - Clon Checkout tipo Kyte para Render

Sistema POS/Checkout responsive inspirado en funciones públicas de Kyte: ventas, productos, inventario, clientes, usuarios, comprobantes y cierre de día.

## Usuarios demo
- Admin: `admin1` / `admin123`
- Vendedor: `vendedor1` / `venta123`

## Ejecutar local
```bash
pip install -r requirements.txt
python app.py
```
Abrir: http://127.0.0.1:5000

## Subir a GitHub
```bash
git init
git add .
git commit -m "POS checkout pro listo para Render"
git branch -M main
git remote add origin TU_URL_DE_GITHUB
git push -u origin main
```

## Render
- New Web Service
- Conectar repositorio
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`
- Variable recomendada: `SECRET_KEY=una_clave_larga_segura`

## Módulos incluidos
- Login con roles admin/vendedor
- Checkout tipo POS
- Carrito dinámico
- Clientes
- Productos e inventario
- Stock automático al vender
- Historial de ventas
- Comprobante imprimible / PDF desde navegador
- Cierre de día
- Responsive para celular
