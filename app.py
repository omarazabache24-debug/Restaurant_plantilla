# -*- coding: utf-8 -*-
"""
RESTAURANTE AORIX - APP WEB PRO COMPLETA
Adaptación Flask lista para GitHub + Render.
Integra la lógica principal del Restaurante_AORIX desktop:
- Panel principal
- Ventas / pedidos
- Inventario
- Recetas con descuento de insumos
- Caja
- Delivery / clientes
- Indicadores
- Reportes + Excel/CSV
- Administrador / usuarios / sucursales
- Log

Usuarios iniciales:
- admin / admin123
- caja / caja123
"""

import csv
import os
import sqlite3
import uuid
from datetime import datetime
from functools import wraps
from io import BytesIO, StringIO
from zoneinfo import ZoneInfo

from flask import (
    Flask, flash, redirect, render_template_string, request,
    send_file, session, url_for
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    from openpyxl import Workbook, load_workbook
    OPENPYXL = True
except Exception:
    Workbook = None
    load_workbook = None
    OPENPYXL = False

# =========================
# CONFIG RENDER / LOCAL
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PERSIST_DIR = os.getenv("PERSIST_DIR", "/data" if os.path.isdir("/data") else os.path.join(BASE_DIR, "data"))
os.makedirs(PERSIST_DIR, exist_ok=True)
DB_PATH = os.path.join(PERSIST_DIR, "restaurante_aorix_pro.db")
STATIC_DIR = os.path.join(BASE_DIR, "static")
CATALOG_DIR = os.path.join(STATIC_DIR, "catalogo")
os.makedirs(CATALOG_DIR, exist_ok=True)
APP_TZ = ZoneInfo(os.getenv("APP_TIMEZONE", "America/Lima"))

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "aorix-restaurante-pro-2026")

APP_TITLE = "EL TORO RESTAURANT GRILL"
APP_SUBTITLE = "Sistema de ventas, pedidos, inventario y catálogo digital"
BRAND = "EL TORO - Restaurant Grill"

# =========================
# HELPERS DB
# =========================
def now():
    return datetime.now(APP_TZ)

def today():
    return now().date().isoformat()

def hour():
    return now().strftime("%H:%M:%S")

def money(v):
    try:
        return "S/ {:,.2f}".format(float(v or 0))
    except Exception:
        return "S/ 0.00"

def clean(v):
    return str(v or "").strip()

def up(v):
    return clean(v).upper()

def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def q_all(sql, params=()):
    with conn() as c:
        return c.execute(sql, params).fetchall()

def q_one(sql, params=()):
    rows = q_all(sql, params)
    return rows[0] if rows else None

def q_exec(sql, params=()):
    with conn() as c:
        cur = c.execute(sql, params)
        c.commit()
        return cur.lastrowid

def log_event(accion, detalle=""):
    try:
        q_exec(
            "INSERT INTO logs(fecha,hora,usuario,accion,detalle) VALUES(?,?,?,?,?)",
            (today(), hour(), session.get("user", "sistema"), accion, detalle),
        )
    except Exception:
        pass

# =========================
# INIT DB
# =========================
def init_db():
    with conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS usuarios(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT UNIQUE NOT NULL,
                nombre TEXT DEFAULT '',
                clave_hash TEXT NOT NULL,
                rol TEXT DEFAULT 'CAJA',
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS sucursales(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS productos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                nombre TEXT UNIQUE NOT NULL,
                categoria TEXT DEFAULT 'PLATOS',
                tipo TEXT DEFAULT 'VENTA',
                unidad TEXT DEFAULT 'PLATO',
                precio REAL DEFAULT 0,
                costo REAL DEFAULT 0,
                stock REAL DEFAULT 0,
                stock_min REAL DEFAULT 0,
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS insumos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                unidad TEXT DEFAULT 'UND',
                stock REAL DEFAULT 0,
                stock_min REAL DEFAULT 0,
                costo REAL DEFAULT 0,
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS recetas(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                producto_id INTEGER NOT NULL,
                insumo_id INTEGER NOT NULL,
                cantidad REAL DEFAULT 0,
                observacion TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS clientes(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                telefono TEXT DEFAULT '',
                direccion TEXT DEFAULT '',
                referencia TEXT DEFAULT '',
                notas TEXT DEFAULT '',
                activo INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS pedidos(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT UNIQUE,
                fecha TEXT,
                hora TEXT,
                mesa TEXT DEFAULT '',
                cliente TEXT DEFAULT '',
                telefono TEXT DEFAULT '',
                direccion TEXT DEFAULT '',
                referencia TEXT DEFAULT '',
                servicio TEXT DEFAULT 'SALÓN',
                metodo_pago TEXT DEFAULT 'EFECTIVO',
                subtotal REAL DEFAULT 0,
                descuento REAL DEFAULT 0,
                total REAL DEFAULT 0,
                estado TEXT DEFAULT 'PENDIENTE',
                pagado TEXT DEFAULT 'NO',
                usuario TEXT DEFAULT '',
                observacion TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS pedido_detalle(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                pedido_id INTEGER,
                producto_id INTEGER,
                producto TEXT,
                cantidad REAL,
                precio REAL,
                total REAL
            );
            CREATE TABLE IF NOT EXISTS ventas(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                hora TEXT,
                pedido_id INTEGER DEFAULT NULL,
                cliente TEXT DEFAULT '',
                servicio TEXT DEFAULT 'SALÓN',
                metodo_pago TEXT DEFAULT 'EFECTIVO',
                subtotal REAL DEFAULT 0,
                descuento REAL DEFAULT 0,
                total REAL DEFAULT 0,
                usuario TEXT DEFAULT '',
                estado TEXT DEFAULT 'PAGADO'
            );
            CREATE TABLE IF NOT EXISTS venta_detalle(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                venta_id INTEGER,
                producto_id INTEGER,
                producto TEXT,
                cantidad REAL,
                precio REAL,
                total REAL
            );
            CREATE TABLE IF NOT EXISTS caja(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                hora TEXT,
                tipo TEXT,
                concepto TEXT,
                monto REAL,
                usuario TEXT,
                venta_id INTEGER DEFAULT NULL,
                estado TEXT DEFAULT 'OK'
            );
            CREATE TABLE IF NOT EXISTS contexto(
                clave TEXT PRIMARY KEY,
                valor TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS catalogo_publico(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                descripcion TEXT DEFAULT '',
                precio REAL DEFAULT 0,
                categoria TEXT DEFAULT 'PLATOS',
                imagen TEXT DEFAULT '',
                destacado INTEGER DEFAULT 0,
                activo INTEGER DEFAULT 1,
                creado TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS logs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT,
                hora TEXT,
                usuario TEXT,
                accion TEXT,
                detalle TEXT
            );
            """
        )
        c.commit()

    defaults = [("Sucursal Principal",), ("Delivery",)]
    for (nombre,) in defaults:
        if not q_one("SELECT id FROM sucursales WHERE nombre=?", (nombre,)):
            q_exec("INSERT INTO sucursales(nombre,activo) VALUES(?,1)", (nombre,))

    for usuario, nombre, clave, rol in [
        ("admin", "Administrador", "admin123", "ADMIN"),
        ("caja", "Caja", "caja123", "CAJA"),
        ("mozo", "Mozo", "mozo123", "MESERO"),
    ]:
        if not q_one("SELECT id FROM usuarios WHERE usuario=?", (usuario,)):
            q_exec(
                "INSERT INTO usuarios(usuario,nombre,clave_hash,rol,activo) VALUES(?,?,?,?,1)",
                (usuario, nombre, generate_password_hash(clave), rol),
            )

    for k, v in {
        "sucursal": "Sucursal Principal",
        "turno": "MAÑANA",
        "dia_abierto": today(),
        "caja_abierta": "0",
        "monto_apertura": "0",
    }.items():
        if not q_one("SELECT clave FROM contexto WHERE clave=?", (k,)):
            q_exec("INSERT INTO contexto(clave,valor) VALUES(?,?)", (k, v))

    if not q_one("SELECT id FROM productos LIMIT 1"):
        demo_productos = [
            ("P001", "1/2 POLLO", "PLATOS", "VENTA", "PLATO", 35, 10, 40, 5),
            ("P002", "1/4 POLLO", "PLATOS", "VENTA", "PLATO", 18, 8, 0, 0),
            ("P003", "1/4 POLLO BROSTER", "PLATOS", "VENTA", "PLATO", 22, 0, 0, 0),
            ("P004", "1/8 POLLO", "PLATOS", "VENTA", "PLATO", 12, 0, 0, 0),
            ("P005", "ADICIONAL CARNE", "ADICIONALES", "VENTA", "UND", 5, 0, 0, 0),
            ("P006", "GASEOSA PERSONAL", "BEBIDAS", "VENTA", "UND", 4, 2, 30, 5),
        ]
        for p in demo_productos:
            q_exec(
                "INSERT INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)",
                p,
            )

    if not q_one("SELECT id FROM catalogo_publico LIMIT 1"):
        for titulo, descripcion, precio, categoria in [
            ("PARRILLA FAMILIAR", "Especial de la casa para compartir", 69.90, "PARRILLAS"),
            ("1/2 POLLO", "Pollo a la brasa con papas y ensalada", 35.00, "PLATOS"),
            ("PIZZA AMERICANA", "Pizza familiar estilo El Toro", 42.00, "PIZZAS"),
        ]:
            q_exec("INSERT INTO catalogo_publico(titulo,descripcion,precio,categoria,imagen,destacado,activo) VALUES(?,?,?,?,?,?,1)", (titulo, descripcion, precio, categoria, "toro_logo.png", 1))

    if not q_one("SELECT id FROM insumos LIMIT 1"):
        for nombre, unidad, stock, stock_min, costo in [
            ("POLLO ENTERO", "UND", 20, 3, 20),
            ("PAPA KG", "KG", 30, 5, 3),
            ("ENSALADA", "PORCION", 50, 10, 1),
            ("MAYONESA", "PORCION", 50, 10, 0.5),
            ("AJI", "PORCION", 50, 10, 0.5),
            ("ENVASE", "UND", 50, 10, 1),
        ]:
            q_exec(
                "INSERT INTO insumos(nombre,unidad,stock,stock_min,costo,activo) VALUES(?,?,?,?,?,1)",
                (nombre, unidad, stock, stock_min, costo),
            )

init_db()

# =========================
# AUTH
# =========================
def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("rol") != "ADMIN":
            flash("Solo el administrador puede ingresar a esta opción.", "error")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)
    return wrapper

def is_admin():
    return session.get("rol") == "ADMIN"

def safe_home():
    return "dashboard" if is_admin() else "ventas"

def admin_only_redirect():
    flash("Acceso restringido: tu usuario solo tiene Venta, Pedido, Cierre y Salir.", "error")
    return redirect(url_for("ventas"))

# =========================
# UI
# =========================
BASE_HTML = r'''
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{title}}</title>
<style>
:root{--navy:#082238;--navy2:#0B2D4A;--green:#66D17B;--green2:#16a34a;--blue:#0d73b8;--bg:#f2f4f7;--line:#cbd5e1;--text:#111827;--muted:#64748b;--red:#dc2626;--orange:#f97316}*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text)}a{text-decoration:none;color:inherit}button,.btn{border:0;background:#e5e7eb;border-radius:4px;padding:10px 16px;font-weight:800;cursor:pointer;box-shadow:inset 0 0 0 1px #999;color:#0f172a;display:inline-block}.btn-primary,button.primary{background:#0B2D4A;color:white}.btn-green{background:#66D17B;color:#082238}.btn-red{background:#fecaca;color:#991b1b}.btn-orange{background:#fb923c;color:white}.btn-blue{background:#0d73b8;color:white}input,select,textarea{width:100%;border:1px solid #b6b6b6;background:white;padding:8px 9px;min-height:34px;color:#0f172a}label{font-weight:800;color:#334155}.login-page{min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#061b2b,#0B2D4A);padding:22px}.login-card{width:min(430px,94vw);background:white;border-radius:16px;padding:28px;box-shadow:0 25px 70px rgba(0,0,0,.28);text-align:center}.logo{font-size:44px;font-weight:950;color:#0B2D4A;letter-spacing:-2px}.logo span{color:#f97316}.login-card label{text-align:left;display:block;margin:12px 0 5px}.login-card button{width:100%;margin-top:16px;background:#0B2D4A;color:white}.hint{font-size:12px;color:#64748b;margin-top:14px;line-height:1.5}.app{display:grid;grid-template-columns:210px minmax(0,1fr);min-height:100vh}.side{background:#082238;color:white;height:100vh;position:sticky;top:0;overflow:auto}.brand{padding:16px;text-align:center;border-bottom:1px solid rgba(255,255,255,.15)}.brand .logo{font-size:34px;color:white}.brand small{color:#dbeafe;font-weight:800}.nav a{display:flex;align-items:center;gap:8px;padding:12px 14px;font-weight:900;border-bottom:1px solid rgba(255,255,255,.05);color:#eef6ff}.nav a.on,.nav a:hover{background:#66D17B;color:#082238}.main{min-width:0}.topbar{background:#082238;color:white;padding:14px 20px;text-align:center}.topbar h1{margin:0;font-size:28px}.topbar p{margin:3px 0 0;color:#dbeafe;font-weight:700}.content{padding:14px;max-width:1700px;margin:0 auto}.flash{padding:12px 14px;border:1px solid #bfdbfe;background:#eff6ff;color:#1d4ed8;font-weight:800;margin-bottom:10px}.flash.error{background:#fff1f2;border-color:#fecaca;color:#991b1b}.flash.ok{background:#ecfdf5;border-color:#bbf7d0;color:#166534}.tabs{display:flex;gap:0;flex-wrap:wrap;background:#ececec;border-bottom:1px solid #aaa;margin:-14px -14px 14px}.tabs a{padding:16px 24px;font-size:18px;font-weight:950;border:1px solid #aaa;border-bottom:0;background:#e5e5e5;color:#082238}.tabs a.on{background:#66D17B}.panel{border:1px solid #999;background:#f8fafc;padding:14px;margin-bottom:14px}.panel legend,.box-title{font-weight:900;color:#0f172a}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;align-items:end}.grid5{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;align-items:end}.grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}.actions{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.kpi{background:white;border:1px solid #222;padding:24px;text-align:center}.kpi h3{margin:0 0 16px;color:#334155}.kpi b{font-size:32px;color:#082238}.kpi .red{color:#dc2626}.table-wrap{overflow:auto;border:1px solid #999;background:white}.table-wrap.small{max-height:360px}.table-wrap table{width:100%;border-collapse:collapse;min-width:900px}th{background:#0B2D4A;color:white;font-size:15px}th,td{border:1px solid #b6b6b6;padding:9px;text-align:center;white-space:nowrap}.row-ok td{background:#e8fff1;color:#047857}.row-bad td{background:#ffe1e1;color:#dc2626}.row-sel td{background:#51728d!important;color:white!important}.badge{display:inline-block;padding:5px 10px;border-radius:999px;font-weight:950;font-size:12px}.ok{background:#dcfce7;color:#166534}.warn{background:#fef3c7;color:#92400e}.off{background:#fee2e2;color:#991b1b}.muted{color:#64748b}.report-box{width:100%;min-height:420px;font-family:Consolas,monospace;white-space:pre;background:white}.chart{height:360px;border:1px solid #cbd5e1;background:white;padding:18px;display:flex;align-items:end;gap:18px}.bar{width:70px;background:#66D17B;border:1px solid #0f766e;display:flex;align-items:flex-start;justify-content:center;font-weight:900;padding-top:4px;color:#082238}.bar-wrap{text-align:center}.mobile-card{display:none}@media(max-width:900px){.app{display:block}.side{height:auto;position:relative}.brand{display:none}.nav{display:grid;grid-template-columns:repeat(2,1fr)}.nav a{justify-content:center;text-align:center}.topbar h1{font-size:22px}.tabs{overflow-x:auto;flex-wrap:nowrap}.tabs a{font-size:14px;padding:12px}.grid,.grid2,.grid5,.kpis{grid-template-columns:1fr}.content{padding:10px}.panel{padding:10px}.table-wrap{max-height:55vh}button,.btn{width:100%;text-align:center}.desktop-only{display:none}.mobile-card{display:block;background:white;border:1px solid #cbd5e1;padding:12px;margin-bottom:10px;border-radius:10px}.mobile-card b{display:inline-block;min-width:115px;color:#64748b}}

/* ===== MEJORAS PRO NEGOCIO 2.0: LOGO, PESTAÑAS Y CELULAR ===== */
.logo{font-family:"Segoe UI",Arial,sans-serif}.side{background:radial-gradient(circle at 90% 95%,rgba(102,209,123,.18),transparent 28%),linear-gradient(180deg,#061b2b,#041827)!important}.brand{padding:18px 12px!important}.brand .logo{font-size:38px!important}.brand small{display:block;margin-top:4px}.nav a{border-radius:10px;margin:6px 10px;border-bottom:0!important;min-height:42px}.nav a.on,.nav a:hover{background:linear-gradient(90deg,#66D17B,#0d73b8)!important;color:white!important}.topbar{background:linear-gradient(135deg,#061b2b,#0B2D4A)!important}.topbar h1{font-size:34px!important;letter-spacing:.2px}.tabs{background:white!important;border:1px solid var(--line)!important;border-radius:14px!important;margin:0 0 14px!important;padding:8px!important;gap:8px!important;box-shadow:0 8px 22px rgba(15,35,55,.07)}.tabs a{border:0!important;border-radius:10px!important;background:#eef2f7!important;font-size:15px!important;padding:12px 16px!important}.tabs a.on{background:linear-gradient(90deg,#66D17B,#0d73b8)!important;color:white!important}.panel{border-radius:14px!important;border:1px solid var(--line)!important;background:white!important;box-shadow:0 8px 20px rgba(15,35,55,.06)}button,.btn{border-radius:10px!important;box-shadow:0 6px 16px rgba(15,35,55,.12)!important}input,select,textarea{border-radius:10px!important}.table-wrap{border-radius:12px!important;border:1px solid var(--line)!important}th{background:#0B2D4A!important}.role-note{background:#ecfdf5;border:1px solid #bbf7d0;color:#166534;font-weight:900;border-radius:12px;padding:12px;margin-bottom:14px}.admin-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.danger-zone{border-color:#fecaca!important;background:#fff7f7!important}
@media(max-width:900px){.app{display:block!important}.side{height:auto!important;position:relative!important}.brand{display:block!important}.brand .logo{font-size:30px!important}.nav{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:8px!important;padding:8px!important}.nav a{margin:0!important;justify-content:center!important;text-align:center!important;font-size:13px!important;padding:11px 8px!important}.topbar{padding:16px 12px!important}.topbar h1{font-size:25px!important}.topbar p{font-size:13px!important}.tabs{overflow-x:auto!important;flex-wrap:nowrap!important;padding:7px!important}.tabs a{min-width:max-content!important;font-size:13px!important;padding:10px 12px!important}.content{padding:10px!important}.admin-grid{grid-template-columns:1fr!important}.grid,.grid2,.grid5,.kpis{grid-template-columns:1fr!important}.actions{display:grid!important;grid-template-columns:1fr!important}.table-wrap{max-height:58vh!important;-webkit-overflow-scrolling:touch}.table-wrap table{min-width:820px!important}button,.btn{width:100%!important}.login-card{width:min(410px,94vw)!important;padding:24px!important}.login-card .logo{font-size:38px!important}}
@media(max-width:430px){.nav{grid-template-columns:1fr!important}.topbar h1{font-size:22px!important}.tabs a{font-size:12px!important}.content{padding:8px!important}.panel{padding:10px!important}}

/* ===== ULTRA MEJORA AORIX RESTAURANTE - WEB + CELULAR ===== */
:root{--dark:#071827;--dark2:#0B2D4A;--accent:#ff7a18;--ok:#22c55e;--sky:#0ea5e9;--panel:#ffffff;--soft:#f5f7fb;--ink:#0f172a}
body{background:linear-gradient(180deg,#eef4f8,#f7fafc)!important}.login-page{background:radial-gradient(circle at 10% 10%,rgba(255,122,24,.18),transparent 28%),linear-gradient(160deg,#06111f,#0B2D4A 58%,#03131f)!important}.login-card{border-radius:28px!important;padding:34px 30px!important;background:rgba(255,255,255,.96)!important}.login-card .logo{font-size:54px!important}.login-card h2{font-size:28px;margin:10px 0 4px}.app{grid-template-columns:230px minmax(0,1fr)!important}.side{background:linear-gradient(180deg,#061b2b,#05121f 55%,#062c25)!important;box-shadow:8px 0 25px rgba(2,8,23,.12)}.brand{min-height:176px;display:grid;place-items:center;align-content:center}.brand .logo{font-size:44px!important}.nav{padding:10px}.nav a{font-size:15px!important;letter-spacing:.1px;min-height:48px;border-radius:14px!important}.nav a.on,.nav a:hover{background:linear-gradient(90deg,#22c55e,#0ea5e9)!important;color:#fff!important;transform:translateX(2px)}.topbar{padding:18px 20px!important;background:linear-gradient(135deg,#071827,#0B2D4A)!important}.topbar h1{font-size:38px!important}.content{padding:18px!important}.tabs{position:sticky;top:0;z-index:7;border-radius:18px!important;padding:10px!important;gap:10px!important}.tabs a{border-radius:14px!important;font-size:15px!important;padding:13px 18px!important}.tabs a.on{box-shadow:0 10px 18px rgba(34,197,94,.20)!important}.panel{border-radius:22px!important;padding:20px!important;border:1px solid #dbe5ef!important}.box-title{font-size:20px}.grid,.grid2,.grid5{gap:16px!important}.kpis{gap:16px!important}.kpi{border:0!important;border-radius:22px!important;box-shadow:0 10px 25px rgba(15,35,55,.08);background:linear-gradient(180deg,#fff,#f8fafc)!important}.kpi h3{font-size:16px}.kpi b{font-size:36px!important}.actions button,.actions .btn,button,.btn{border:0!important;border-radius:14px!important;box-shadow:0 10px 20px rgba(15,35,55,.10)!important;min-height:46px}.primary,.btn-primary,button.primary{background:linear-gradient(90deg,#0B2D4A,#0d73b8)!important;color:white!important}.btn-green{background:linear-gradient(90deg,#22c55e,#0ea5e9)!important;color:white!important}.btn-red{background:linear-gradient(90deg,#ef4444,#991b1b)!important;color:white!important}.btn-orange{background:linear-gradient(90deg,#fb923c,#f97316)!important;color:white!important}input,select,textarea{border-radius:14px!important;min-height:45px!important;border:1px solid #cbd5e1!important}.table-wrap{border-radius:18px!important}.table-wrap table{min-width:950px}th{background:#082238!important;height:44px}.report-card-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:14px}.mini-report{background:#fff;border-radius:18px;padding:18px;border:1px solid #e2e8f0;box-shadow:0 8px 18px rgba(15,35,55,.07)}.mini-report b{font-size:24px}.chart-pro{height:310px;background:linear-gradient(180deg,#fff,#f8fafc);border-radius:18px;border:1px solid #e2e8f0;padding:18px;display:flex;align-items:end;gap:14px;overflow:auto}.bar{border-radius:12px 12px 0 0!important;background:linear-gradient(180deg,#0ea5e9,#22c55e)!important;border:0!important;min-width:76px;color:white!important}.search-row{display:grid;grid-template-columns:minmax(260px,1fr) 170px 160px;gap:12px;align-items:end;margin:12px 0}.mobile-only{display:none}
@media(max-width:900px){body{background:#111!important;color:#e5e7eb!important}.app{display:block!important}.side{height:auto!important;position:relative!important;border-bottom:1px solid rgba(255,255,255,.08)}.brand{display:block!important;min-height:auto!important;padding:22px 12px!important}.brand .logo{font-size:42px!important}.nav{display:grid!important;grid-template-columns:repeat(2,1fr)!important;gap:9px!important;padding:10px 14px 18px!important}.nav a{background:rgba(255,255,255,.05)!important;margin:0!important;justify-content:center!important;text-align:center!important;font-size:15px!important}.main{background:#141414!important}.topbar{padding:22px 12px!important}.topbar h1{font-size:32px!important;line-height:1.1}.topbar p{font-size:16px}.content{padding:12px!important}.tabs{position:relative!important;background:#1f1f1f!important;border-color:#3f3f46!important;border-radius:24px!important;overflow-x:auto!important;flex-wrap:nowrap!important}.tabs a{background:#2d2d30!important;color:#e5e7eb!important;min-width:max-content}.tabs a.on{background:linear-gradient(90deg,#064e3b,#0B2D4A)!important}.panel{background:#1f1f1f!important;border-color:#424242!important;color:#e5e7eb!important}.box-title,label{color:#f8fafc!important}.grid,.grid2,.grid5,.kpis,.report-card-grid,.search-row{grid-template-columns:1fr!important}.actions{display:grid!important;grid-template-columns:1fr!important}.actions input,.actions select{max-width:none!important}.kpi{background:#202020!important;color:#fff!important}.kpi h3{color:#e5e7eb}.kpi b{color:#fff!important}input,select,textarea{background:#4a4a4a!important;color:white!important;border-color:#666!important;font-size:16px!important;min-height:58px!important}button,.btn{width:100%!important;min-height:58px!important;font-size:16px!important}.table-wrap{background:#202020!important;border-color:#555!important;max-height:55vh!important}.table-wrap table{min-width:820px!important}td{color:#f3f4f6!important;background:#202020}.desktop-only{display:none!important}.mobile-only{display:block!important}.report-box{background:#202020!important;color:#e5e7eb!important}.chart-pro{background:#202020!important;border-color:#555!important}.flash{border-radius:16px!important}}
@media(max-width:430px){.nav{grid-template-columns:1fr!important}.brand small{font-size:14px}.topbar h1{font-size:28px!important}.content{padding:10px!important}.panel{padding:18px!important}.tabs a{font-size:15px!important;padding:13px 16px!important}.box-title{font-size:22px}}


/* ===== AORIX CADENA RESTAURANTES ULTRA V2 ===== */
body{font-size:16px}.content{max-width:1800px}.panel{padding:20px!important}.box-title{font-size:22px}.section-title{display:flex;align-items:center;gap:10px;font-size:22px;font-weight:950;color:#071827;margin:0 0 14px}.hint-card{background:#eff6ff;border:1px solid #bfdbfe;color:#0f3b68;border-radius:14px;padding:14px;font-weight:750;line-height:1.45}.clean-grid{display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:14px;align-items:end}.clean-grid-4{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:14px;align-items:end}.btn-danger{background:linear-gradient(90deg,#ef4444,#b91c1c)!important;color:white!important}.btn-success{background:linear-gradient(90deg,#22c55e,#0ea5e9)!important;color:white!important}.btn-warning{background:linear-gradient(90deg,#fb923c,#f97316)!important;color:white!important}.sku-help{font-size:12px;color:#64748b;font-weight:700;margin-top:6px}.analytics-grid{display:grid;grid-template-columns:repeat(4,minmax(170px,1fr));gap:16px;margin-bottom:16px}.analytics-card{background:white;border:1px solid #dbe7f3;border-radius:18px;padding:18px;box-shadow:0 12px 28px rgba(15,35,55,.08)}.analytics-card span{display:block;color:#64748b;font-weight:900;font-size:13px}.analytics-card b{display:block;font-size:32px;color:#071827;margin-top:8px}.analytics-card small{color:#64748b;font-weight:750}.chart-pro{min-height:300px;border:1px solid #dbe7f3;border-radius:18px;background:linear-gradient(180deg,#fff,#f8fbff);padding:16px;display:flex;align-items:end;gap:14px;overflow:auto}.bar{min-width:72px;border-radius:12px 12px 4px 4px;background:linear-gradient(180deg,#22c55e,#0ea5e9)!important;color:white!important;border:0!important;box-shadow:0 10px 24px rgba(14,165,233,.22)}.bar-wrap small{display:block;margin-top:8px;font-weight:900;color:#334155}.report-card-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:14px}.mini-report{background:white;border:1px solid #dbe7f3;border-radius:16px;padding:18px;font-weight:900;color:#64748b;box-shadow:0 10px 24px rgba(15,35,55,.07)}.mini-report b{display:block;margin-top:7px;color:#071827;font-size:28px}.ops-strip{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.ops-strip .hint-card{background:#f0fdf4;border-color:#bbf7d0;color:#14532d}.pedido-actions{display:grid;grid-template-columns:1.5fr 1fr auto;gap:12px;align-items:end}.pedido-actions-2{display:grid;grid-template-columns:1.2fr 1.2fr auto;gap:12px;align-items:end}.service-note{display:grid;grid-template-columns:1fr 1fr;gap:14px}.main-context{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;align-items:end}.mobile-bottom{display:none}.dash-kpi{border:0!important;border-radius:18px!important;background:linear-gradient(180deg,#fff,#f8fbff);box-shadow:0 14px 32px rgba(15,35,55,.08)!important}.dash-kpi b{font-size:36px}@media(max-width:900px){.side{position:sticky!important;top:0!important;z-index:20}.brand{min-height:auto!important;padding:14px!important}.nav{grid-template-columns:repeat(4,1fr)!important}.nav a{font-size:12px!important;min-height:44px!important}.content{padding:10px!important}.tabs{position:relative!important;overflow-x:auto!important}.analytics-grid,.report-card-grid,.clean-grid,.clean-grid-4,.ops-strip,.pedido-actions,.pedido-actions-2,.service-note,.main-context{grid-template-columns:1fr!important}.panel{padding:16px!important}.box-title,.section-title{font-size:20px!important}.kpis{grid-template-columns:1fr!important}.table-wrap table{min-width:760px!important}.mobile-bottom{display:grid;position:fixed;bottom:0;left:0;right:0;z-index:30;grid-template-columns:repeat(4,1fr);background:#071827;padding:8px;gap:6px}.mobile-bottom a{color:white;text-align:center;font-weight:900;font-size:12px;background:rgba(255,255,255,.08);border-radius:12px;padding:10px 4px}.content{padding-bottom:78px!important}}@media(max-width:520px){.nav{grid-template-columns:repeat(2,1fr)!important}.topbar h1{font-size:24px!important}.login-card .logo{font-size:46px!important}.analytics-card b{font-size:26px}.chart-pro{min-height:240px}.bar{min-width:58px}.tabs a{font-size:12px!important}.side .brand small{font-size:11px}}


/* ===== AORIX FOOD APP THEME V3 - inspiración móvil tipo food app / Kyte ===== */
:root{--food-dark:#0c0b1b;--food-card:#171727;--food-card2:#1f1f2d;--food-line:#323348;--food-coral:#ff4f5f;--food-coral2:#ff7a59;--food-lime:#65d77b;--food-blue:#0b2d4a;--food-text:#f8fafc;--food-muted:#aab2c5;}
.login-page{background:radial-gradient(circle at 15% 12%,rgba(255,79,95,.28),transparent 24%),radial-gradient(circle at 80% 86%,rgba(101,215,123,.16),transparent 24%),linear-gradient(135deg,#ec7a9c 0%,#a8bad6 42%,#13233b 100%)!important;}
.login-card{background:#10101f!important;color:#fff!important;border:1px solid rgba(255,255,255,.08)!important;border-radius:22px!important;box-shadow:0 24px 60px rgba(0,0,0,.35)!important;position:relative;overflow:hidden;}
.login-card:after{content:"15%\A descuento\A primera compra";white-space:pre;position:absolute;left:22px;right:22px;bottom:22px;min-height:88px;border-radius:18px;background:linear-gradient(120deg,var(--food-coral),#ff6b6b);color:#fff;font-weight:950;font-size:22px;display:flex;align-items:center;padding:18px;text-align:left;opacity:.96;}
.login-card .logo{color:#fff!important;margin-bottom:10px}.login-card h2{color:#fff!important}.login-card .muted,.login-card .hint{color:#cbd5e1!important}.login-card label{color:#e5e7eb!important}.login-card input{background:#1b1b2b!important;border:1px solid #35364b!important;color:#fff!important;border-radius:14px!important}.login-card button{background:linear-gradient(90deg,var(--food-coral),var(--food-coral2))!important;border-radius:999px!important;margin-bottom:120px!important;color:#fff!important;box-shadow:0 12px 28px rgba(255,79,95,.28)!important;}
.topbar{background:linear-gradient(135deg,#071827,#0b2d4a 60%,#06111f)!important;box-shadow:0 10px 26px rgba(11,45,74,.20)}
.tabs{border-radius:22px!important;border:1px solid #dbe7f3!important;box-shadow:0 12px 28px rgba(15,35,55,.07)!important}.tabs a{border-radius:15px!important;min-height:50px;display:flex;align-items:center;justify-content:center}.panel{border-radius:22px!important;border:1px solid #dce7f5!important;box-shadow:0 18px 42px rgba(15,35,55,.08)!important}.kpi,.analytics-card,.mini-report{border:0!important;border-radius:22px!important;box-shadow:0 20px 45px rgba(15,35,55,.08)!important}.kpi b,.analytics-card b,.mini-report b{letter-spacing:-.8px}.btn-primary,.primary,button.primary{background:linear-gradient(90deg,#0b2d4a,#0ea5e9)!important;color:white!important}.btn-green,.btn-success{background:linear-gradient(90deg,#65d77b,#0ea5e9)!important;color:white!important}.btn-red,.btn-danger{background:linear-gradient(90deg,#ff4f5f,#b91c1c)!important;color:white!important}.btn-orange,.btn-warning{background:linear-gradient(90deg,#ff7a59,#f97316)!important;color:white!important}button,.btn{border:0!important;border-radius:14px!important}.admin-grid{grid-template-columns:1fr!important}.admin-panels{display:grid;grid-template-columns:1fr 1fr;gap:16px}.users-clean-table table{min-width:720px}.delivery-advice{background:linear-gradient(135deg,#fff7ed,#fff);border:1px solid #fed7aa;border-radius:18px;padding:14px;font-weight:800;color:#9a3412}.food-card-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px}.food-card{background:linear-gradient(160deg,#fff,#f8fbff);border:1px solid #dbe7f3;border-radius:20px;padding:14px;box-shadow:0 14px 32px rgba(15,35,55,.08)}.food-card .thumb{height:74px;border-radius:16px;background:linear-gradient(135deg,#fff1f2,#fff7ed);display:grid;place-items:center;font-size:32px;margin-bottom:10px}.food-card b{display:block;color:#071827;font-size:16px}.food-card small{color:#64748b;font-weight:800}.report-box{border-radius:18px!important;border:1px solid #dbe7f3!important;min-height:220px!important;background:#f8fbff!important}.chart-pro{background:linear-gradient(180deg,#171727,#10101f)!important;border-color:#323348!important;color:#fff!important}.chart-pro .muted{color:#cbd5e1!important}.bar{background:linear-gradient(180deg,#ff4f5f,#ff7a59)!important;border-radius:18px 18px 6px 6px!important}.bar-wrap small{color:#e5e7eb!important}.table-wrap th{background:#071827!important}.table-wrap tr:hover td{background:#f8fafc}
@media(max-width:900px){body{background:#151515!important;color:#f8fafc!important}.app{display:block!important}.side{position:relative!important;height:auto!important;background:#050b16!important}.brand{background:#050b16!important;border-bottom:1px solid #1f2937!important}.brand .logo{font-size:42px!important}.brand small{color:#e5e7eb!important}.nav{display:grid!important;grid-template-columns:1fr!important;gap:10px!important;padding:16px!important}.nav a{background:transparent!important;color:#e5e7eb!important;border-radius:18px!important;min-height:58px!important;font-size:20px!important;justify-content:center!important}.nav a.on{background:linear-gradient(90deg,#073b16,#061b3e)!important;color:#fff!important}.topbar{background:#050b16!important;padding:30px 18px!important}.topbar h1{font-size:30px!important}.topbar p{font-size:17px!important;color:#cbd5e1!important}.content{background:#1d1d1d!important;padding:16px!important}.tabs{background:#202020!important;border:1px solid #4b5563!important;border-radius:24px!important;padding:10px!important;gap:10px!important}.tabs a{background:#20242c!important;color:#dbeafe!important;border-radius:18px!important;min-width:110px!important}.tabs a.on{background:linear-gradient(90deg,#073b16,#061b3e)!important}.panel,.food-card,.analytics-card,.mini-report,.kpi{background:#202020!important;color:#f8fafc!important;border:1px solid #4b5563!important;border-radius:24px!important}.box-title,.section-title,h2,h3,label{color:#f8fafc!important}.muted,.sku-help,.hint-card small{color:#cbd5e1!important}.hint-card,.role-note,.delivery-advice{background:#102017!important;color:#bbf7d0!important;border-color:#166534!important}.clean-grid,.clean-grid-4,.main-context,.service-note,.pedido-actions,.pedido-actions-2,.admin-panels,.grid,.grid2,.grid5,.kpis,.analytics-grid,.report-card-grid{grid-template-columns:1fr!important}.actions{display:grid!important;grid-template-columns:1fr!important}input,select,textarea{background:#4a4a4a!important;color:#fff!important;border:1px solid #5b5b5b!important;border-radius:20px!important;min-height:62px!important;font-size:18px!important;padding:14px 18px!important}button,.btn{min-height:58px!important;border-radius:18px!important;font-size:18px!important}.table-wrap{background:#202020!important;border:1px solid #4b5563!important;border-radius:20px!important;max-height:55vh!important}.table-wrap td{background:#202020;color:#fff}.table-wrap table{min-width:780px!important}.mobile-bottom{background:#101010!important}.mobile-bottom a{font-size:13px!important;background:#232323!important}.login-card:after{display:none}.login-card button{margin-bottom:20px!important}}



/* ==========================================================
   AORIX FOOD APP V4 - INSPIRACIÓN KYTE / APP DE COMIDA
   Web limpio + celular presentable + colores coral/dark
   ========================================================== */
:root{
  --kyte-bg:#f6f7fb; --kyte-card:#ffffff; --kyte-dark:#11111f; --kyte-dark2:#171827;
  --kyte-coral:#ff4d57; --kyte-coral2:#ff7a59; --kyte-gold:#ffb23f;
  --kyte-green:#15c77a; --kyte-blue:#2478ff; --kyte-line:#e7ebf2; --kyte-text:#101828;
}
body{background:var(--kyte-bg)!important;color:var(--kyte-text)!important;font-family:Inter,"Segoe UI",Arial,sans-serif!important;}
.app{grid-template-columns:224px minmax(0,1fr)!important;background:var(--kyte-bg)!important;}
.side{background:linear-gradient(180deg,#0b1220,#070a16 70%,#0b2d29)!important;border-right:1px solid rgba(255,255,255,.06)!important;}
.brand{min-height:168px!important;padding:20px 14px!important;}
.brand .logo{font-size:44px!important;letter-spacing:-2.5px!important;}
.brand small{font-size:12px!important;opacity:.92!important;line-height:1.35!important;}
.nav{padding:12px 10px 20px!important;}
.nav a{border-radius:18px!important;margin:7px 0!important;padding:13px 15px!important;min-height:48px!important;color:#f6f7fb!important;font-size:15px!important;}
.nav a.on,.nav a:hover{background:linear-gradient(100deg,var(--kyte-coral),#ff7a59)!important;color:white!important;box-shadow:0 15px 30px rgba(255,77,87,.28)!important;transform:none!important;}
.topbar{background:linear-gradient(180deg,#0c2136,#09192b)!important;padding:22px 20px!important;border-bottom:1px solid rgba(255,255,255,.08)!important;}
.topbar h1{font-size:38px!important;letter-spacing:.4px!important;}
.topbar p{font-size:15px!important;color:#dce7f4!important;}
.content{padding:22px!important;max-width:1560px!important;}
.tabs{background:white!important;border:1px solid var(--kyte-line)!important;border-radius:22px!important;padding:12px!important;gap:10px!important;box-shadow:0 14px 38px rgba(16,24,40,.07)!important;}
.tabs a{border-radius:16px!important;background:#f1f4f8!important;color:#0d1b2f!important;font-weight:900!important;border:1px solid transparent!important;padding:13px 20px!important;}
.tabs a.on{background:linear-gradient(100deg,var(--kyte-coral),var(--kyte-coral2))!important;color:white!important;box-shadow:0 14px 28px rgba(255,77,87,.24)!important;}
.panel{background:white!important;border:1px solid var(--kyte-line)!important;border-radius:24px!important;padding:22px!important;box-shadow:0 16px 40px rgba(16,24,40,.06)!important;margin-bottom:18px!important;}
.section-title,.box-title{font-size:23px!important;color:#0b1220!important;margin-bottom:12px!important;}
input,select,textarea{border:1px solid #d8e0ec!important;border-radius:16px!important;min-height:50px!important;padding:12px 15px!important;background:#fff!important;font-size:15px!important;}
input:focus,select:focus,textarea:focus{outline:none!important;border-color:var(--kyte-coral)!important;box-shadow:0 0 0 4px rgba(255,77,87,.12)!important;}
button,.btn{border-radius:16px!important;border:0!important;min-height:50px!important;padding:12px 20px!important;font-weight:950!important;box-shadow:0 14px 28px rgba(16,24,40,.08)!important;background:#eef2f7!important;color:#0b1220!important;}
button.primary,.btn-primary,.btn-success,.btn-green{background:linear-gradient(100deg,var(--kyte-coral),var(--kyte-coral2))!important;color:#fff!important;box-shadow:0 16px 28px rgba(255,77,87,.23)!important;}
.btn-warning,.btn-orange{background:linear-gradient(100deg,var(--kyte-gold),#ff7a18)!important;color:#fff!important;}
.btn-danger,.btn-red{background:linear-gradient(100deg,#ef4444,#b91c1c)!important;color:#fff!important;}
.clean-grid,.clean-grid-4,.grid,.grid2,.grid5{gap:16px!important;}
.kpis{gap:18px!important;}
.kpi,.analytics-card,.mini-report{border:1px solid var(--kyte-line)!important;border-radius:24px!important;background:#fff!important;box-shadow:0 16px 40px rgba(16,24,40,.06)!important;}
.kpi b,.analytics-card b,.mini-report b{color:#0b1220!important;}
.table-wrap{border:1px solid var(--kyte-line)!important;border-radius:20px!important;box-shadow:none!important;}
th{background:#111827!important;color:white!important;height:48px!important;font-size:14px!important;}
th,td{border-color:#e6eaf1!important;padding:12px 10px!important;}
.hint-card,.sku-help,.service-note .hint-card{border-radius:18px!important;border:1px solid #ffe1e3!important;background:#fff6f7!important;color:#7f1d1d!important;padding:14px!important;}
.report-card-grid{display:grid!important;grid-template-columns:repeat(4,minmax(160px,1fr))!important;gap:16px!important;margin-bottom:18px!important;}
.chart-pro{min-height:300px!important;border-radius:22px!important;background:linear-gradient(180deg,#fff,#fff8f8)!important;border:1px solid #ffe1e3!important;padding:18px!important;display:flex!important;align-items:end!important;gap:16px!important;overflow:auto!important;}
.bar{background:linear-gradient(180deg,var(--kyte-coral),var(--kyte-coral2))!important;border:0!important;color:#fff!important;border-radius:14px 14px 6px 6px!important;min-width:72px!important;box-shadow:0 12px 26px rgba(255,77,87,.22)!important;}
/* Login estilo app food */
.login-page{background:radial-gradient(circle at 12% 14%,rgba(255,77,87,.25),transparent 26%),linear-gradient(145deg,#121321,#080914 70%,#111827)!important;}
.login-card{background:#11111f!important;color:white!important;border:1px solid rgba(255,255,255,.08)!important;border-radius:30px!important;box-shadow:0 28px 80px rgba(0,0,0,.45)!important;}
.login-card .logo{color:#fff!important;font-size:52px!important;}.login-card .logo span{color:#ff7a18!important}.login-card p,.login-card .muted,.login-card .hint{color:#cbd5e1!important}.login-card label{color:#fff!important}.login-card input{background:#1b1c2b!important;border-color:#2f3347!important;color:white!important}.login-card button{background:linear-gradient(100deg,var(--kyte-coral),var(--kyte-coral2))!important;border-radius:999px!important;color:white!important;}
/* mobile tipo app: sin sidebar gigante, navegación superior compacta */
@media(max-width:900px){
  .app{display:block!important;background:#101016!important;}
  body{background:#101016!important;color:#f8fafc!important;}
  .side{height:auto!important;position:relative!important;background:#090b16!important;border-right:0!important;box-shadow:none!important;}
  .brand{display:block!important;min-height:auto!important;padding:18px 14px 10px!important;}
  .brand .logo{font-size:42px!important;}.brand small{font-size:13px!important;}
  .nav{display:grid!important;grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:9px!important;padding:10px 12px 16px!important;}
  .nav a{margin:0!important;justify-content:center!important;text-align:center!important;min-height:48px!important;border-radius:20px!important;background:#171827!important;color:#f8fafc!important;font-size:13px!important;padding:10px 7px!important;}
  .nav a.on,.nav a:hover{background:linear-gradient(100deg,var(--kyte-coral),var(--kyte-coral2))!important;box-shadow:0 12px 24px rgba(255,77,87,.25)!important;}
  .topbar{background:#090b16!important;padding:18px 12px 16px!important;border-top:1px solid rgba(255,255,255,.06)!important;}
  .topbar h1{font-size:28px!important;line-height:1.05!important;}.topbar p{font-size:14px!important;}
  .content{padding:14px!important;background:#101016!important;max-width:none!important;}
  .tabs{background:#1b1c22!important;border-color:#3a3d48!important;border-radius:28px!important;overflow-x:auto!important;flex-wrap:nowrap!important;padding:10px!important;margin-bottom:14px!important;}
  .tabs a{background:#22242d!important;color:#e5e7eb!important;border-radius:18px!important;min-width:max-content!important;padding:12px 18px!important;}
  .tabs a.on{background:linear-gradient(100deg,var(--kyte-coral),var(--kyte-coral2))!important;color:white!important;}
  .panel{background:#1f1f20!important;color:#f8fafc!important;border-color:#3b3f4a!important;border-radius:26px!important;padding:18px!important;box-shadow:none!important;}
  .section-title,.box-title,label{color:#f8fafc!important;}
  input,select,textarea{background:#4a4a4d!important;color:#fff!important;border-color:#606169!important;border-radius:18px!important;min-height:56px!important;font-size:16px!important;}
  input::placeholder{color:#c5c8d1!important;}
  button,.btn{width:100%!important;border-radius:18px!important;min-height:56px!important;font-size:16px!important;}
  .clean-grid,.clean-grid-4,.grid,.grid2,.grid5,.kpis,.analytics-grid,.report-card-grid{grid-template-columns:1fr!important;gap:12px!important;}
  .kpi,.analytics-card,.mini-report{background:#1f1f20!important;color:#f8fafc!important;border-color:#3b3f4a!important;border-radius:24px!important;}
  .kpi b,.analytics-card b,.mini-report b{color:#fff!important;}
  .table-wrap{border-color:#3b3f4a!important;background:#1f1f20!important;max-height:55vh!important;-webkit-overflow-scrolling:touch!important;}
  .table-wrap table{min-width:820px!important;} th{background:#010a22!important;} td{color:#f8fafc!important;border-color:#3b3f4a!important;}
  .hint-card,.sku-help{background:#261d1f!important;color:#fecdd3!important;border-color:#7f1d1d!important;}
}
@media(max-width:430px){.nav{grid-template-columns:repeat(2,minmax(0,1fr))!important}.content{padding:10px!important}.panel{padding:16px!important}.topbar h1{font-size:24px!important}.login-card{width:94vw!important;padding:26px 22px!important}}



/* ===== EL TORO / KYTE FOOD V6 ===== */
:root{--toro-red:#ff1744;--toro-red2:#ff5b4f;--toro-black:#07070c;--toro-card:#11131d;--toro-mint:#2dd4bf;}
body{background:#f4f6f9!important;overflow-x:hidden!important}.topbar{background:linear-gradient(135deg,#06070d,#111827 55%,#38010a)!important;border-bottom:4px solid var(--toro-red)!important}.brand{background:linear-gradient(180deg,#07070c,#111827)!important}.brand .logo{display:flex!important;align-items:center;justify-content:center;gap:8px}.brand-logo-img{width:78px;height:78px;object-fit:contain;display:block;margin:0 auto 8px;filter:drop-shadow(0 8px 20px rgba(255,23,68,.28))}.nav a.on,.tabs a.on{background:linear-gradient(100deg,var(--toro-red),var(--toro-red2))!important;color:#fff!important;box-shadow:0 16px 35px rgba(255,23,68,.28)!important}.btn-success,.btn-green,.primary,button.primary{background:linear-gradient(100deg,var(--toro-red),var(--toro-red2))!important;color:#fff!important;border:0!important}.btn-warning,.btn-orange{background:linear-gradient(100deg,#ff9f1c,#ff5b00)!important;color:#fff!important;border:0!important}.btn-danger,.btn-red{background:linear-gradient(100deg,#b91c1c,#ff1744)!important;color:#fff!important;border:0!important}.panel{border:1px solid #e5e7eb!important;border-radius:26px!important;box-shadow:0 16px 45px rgba(15,23,42,.08)!important;background:#fff!important}.section-title{font-size:24px!important;font-weight:950!important;color:#07070c!important}.catalog-hero{border-radius:30px;background:linear-gradient(135deg,#07070c,#151827 60%,#43000c);color:#fff;padding:28px;display:grid;grid-template-columns:1.1fr .9fr;gap:22px;align-items:center;overflow:hidden;position:relative}.catalog-hero:after{content:"";position:absolute;right:-80px;top:-80px;width:260px;height:260px;border-radius:50%;background:rgba(255,23,68,.28)}.catalog-hero h2{font-size:36px;margin:0 0 8px}.catalog-phone{background:#0b0c14;border:1px solid rgba(255,255,255,.12);border-radius:30px;padding:18px;box-shadow:0 22px 60px rgba(0,0,0,.35)}.phone-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}.phone-food{background:#171927;border-radius:18px;padding:10px}.phone-food img{width:100%;height:90px;object-fit:cover;border-radius:14px}.phone-food b{display:block;color:#fff;font-size:12px;margin-top:6px}.phone-food small{color:#ffccd3}.catalog-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:18px}.catalog-card{background:#11131d;border-radius:24px;padding:12px;color:white;box-shadow:0 18px 45px rgba(0,0,0,.18);position:relative}.catalog-card img{width:100%;height:150px;object-fit:cover;border-radius:18px}.catalog-card h3{margin:10px 0 4px;font-size:18px}.catalog-card p{color:#cbd5e1;font-size:13px;min-height:38px}.price-pill{display:inline-block;background:linear-gradient(100deg,var(--toro-red),var(--toro-red2));padding:8px 12px;border-radius:999px;font-weight:950}.qr-box{display:grid;grid-template-columns:220px 1fr;gap:18px;align-items:center;background:#fff7f8;border:1px solid #fecdd3;border-radius:24px;padding:18px}.qr-img{width:180px;height:180px;border:10px solid white;border-radius:22px;box-shadow:0 12px 30px rgba(0,0,0,.12)}.upload-drop{border:2px dashed #fecdd3;background:#fff7f8;border-radius:24px;padding:18px}.analytics-hero{background:linear-gradient(135deg,#07070c,#141827)!important;color:white!important;border:0!important}.analytics-hero .section-title{color:white!important}.chart-pro{min-height:310px;border-radius:24px;background:linear-gradient(180deg,#fff,#fff7f8)!important;border:1px solid #fee2e2!important;padding:20px!important;display:flex;align-items:end;gap:18px;overflow:auto}.bar{background:linear-gradient(180deg,var(--toro-red2),var(--toro-red))!important;border:0!important;border-radius:16px 16px 4px 4px!important;color:#fff!important}.admin-panels{display:grid;grid-template-columns:1.1fr .9fr;gap:18px}.admin-mini-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}.admin-mini{border-radius:24px;background:linear-gradient(135deg,#111827,#07070c);color:white;padding:20px;min-height:120px}.admin-mini b{display:block;font-size:18px;margin-top:8px}.admin-mini small{color:#cbd5e1}.load-day-panel{margin-top:18px!important;background:#fff7ed!important;border-color:#fed7aa!important}.product-toolbar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:16px}.product-toolbar input{max-width:520px}.pedido-stack{display:grid;gap:18px}.pedido-card{border:1px solid #e5e7eb;border-radius:24px;background:#fff;padding:18px}.pedido-card h3{margin:0 0 14px}.pedido-grid{display:grid;grid-template-columns:minmax(280px,1fr) 220px auto;gap:14px;align-items:end}.pedido-grid-2{display:grid;grid-template-columns:minmax(280px,1fr) minmax(280px,1fr) auto;gap:14px;align-items:end}
@media(max-width:1100px){.catalog-grid{grid-template-columns:repeat(2,1fr)}.catalog-hero,.admin-panels,.qr-box{grid-template-columns:1fr}.admin-mini-grid{grid-template-columns:1fr}.pedido-grid,.pedido-grid-2{grid-template-columns:1fr}}
@media(max-width:900px){.catalog-grid{grid-template-columns:1fr}.catalog-hero h2{font-size:28px}.catalog-phone{display:none}.brand-logo-img{width:92px;height:92px}.side{background:#07070c!important}.content{background:#101016!important}.panel{background:#1f1f20!important;border-color:#333845!important;color:#f8fafc!important}.section-title{color:#f8fafc!important}.catalog-card{background:#11131d!important}.qr-box,.upload-drop,.load-day-panel{background:#1f1f20!important;border-color:#333845!important}.product-toolbar input{max-width:none}.table-wrap{overflow-x:auto!important}}

</style>
</head>
<body>
{% if session.get('user') %}
<div class="app">
  <aside class="side">
    <div class="brand"><img class="brand-logo-img" src="{{url_for('static', filename='toro_logo.png')}}" onerror="this.style.display='none'"><div class="logo">EL <span>TORO</span></div><small>Restaurant Grill</small><br><small>{{session.get('user')}} - {{session.get('rol')}}</small></div>
    <nav class="nav">
      {% if session.get('rol') == 'ADMIN' %}
      <a class="{{'on' if active=='dashboard' else ''}}" href="{{url_for('dashboard')}}">📊 Panel principal</a>
      {% endif %}
      <a class="{{'on' if active=='ventas' else ''}}" href="{{url_for('ventas')}}">🧾 Venta</a>
      <a class="{{'on' if active=='pedidos' else ''}}" href="{{url_for('pedidos')}}">🚚 Pedido</a>
      <a class="{{'on' if active=='cierre' else ''}}" href="{{url_for('cierre')}}">🔒 Cierre</a>
      {% if session.get('rol') == 'ADMIN' %}
      <a class="{{'on' if active=='inventario' else ''}}" href="{{url_for('inventario')}}">📦 Inventario</a>
      <a class="{{'on' if active=='catalogo' else ''}}" href="{{url_for('catalogo_admin')}}">🖼️ Catálogo / QR</a>
      <a class="{{'on' if active=='recetas' else ''}}" href="{{url_for('recetas')}}">🍽️ Recetas</a>
      <a class="{{'on' if active=='caja' else ''}}" href="{{url_for('caja')}}">💵 Caja</a>
      <a class="{{'on' if active=='delivery' else ''}}" href="{{url_for('delivery')}}">🛵 Delivery</a>
      <a class="{{'on' if active=='indicadores' else ''}}" href="{{url_for('indicadores')}}">📈 Indicadores</a>
      <a class="{{'on' if active=='reportes' else ''}}" href="{{url_for('reportes')}}">📄 Reportes</a>
      <a class="{{'on' if active=='admin' else ''}}" href="{{url_for('admin')}}">⚙️ Usuarios / Admin</a>
      <a class="{{'on' if active=='log' else ''}}" href="{{url_for('logs')}}">🧾 Log</a>
      {% endif %}
      <a href="{{url_for('logout')}}">🚪 Salir</a>
    </nav>
  </aside>
  <main class="main">
    <header class="topbar"><h1>{{title}}</h1><p>{{subtitle}}</p></header>
    <div class="content">
      <div class="tabs">
        {% for key,label,endpoint in tabs %}<a class="{{'on' if active==key else ''}}" href="{{url_for(endpoint)}}">{{label}}</a>{% endfor %}
      </div>
      {% with msgs=get_flashed_messages(with_categories=true) %}{% for cat,msg in msgs %}<div class="flash {{cat}}">{{msg}}</div>{% endfor %}{% endwith %}
      {{content|safe}}
      <div class="mobile-bottom"><a href="{{url_for('ventas')}}">Venta</a><a href="{{url_for('pedidos')}}">Pedido</a><a href="{{url_for('cierre')}}">Cierre</a><a href="{{url_for('logout')}}">Salir</a></div>
    </div>
  </main>
</div>
{% else %}
{{content|safe}}
{% endif %}
</body>
</html>
'''

def get_ctx(k, default=""):
    r = q_one("SELECT valor FROM contexto WHERE clave=?", (k,))
    return r["valor"] if r else default

def set_ctx(k, v):
    if q_one("SELECT clave FROM contexto WHERE clave=?", (k,)):
        q_exec("UPDATE contexto SET valor=? WHERE clave=?", (str(v), k))
    else:
        q_exec("INSERT INTO contexto(clave,valor) VALUES(?,?)", (k, str(v)))

def tabs():
    if session.get("rol") != "ADMIN":
        return [
            ("ventas", "Venta", "ventas"),
            ("pedidos", "Pedido", "pedidos"),
            ("cierre", "Cierre", "cierre"),
        ]
    return [
        ("dashboard", "Panel Principal", "dashboard"),
        ("ventas", "Ventas", "ventas"),
        ("pedidos", "Pedidos", "pedidos"),
        ("cierre", "Cierre", "cierre"),
        ("inventario", "Inventario", "inventario"),
        ("catalogo", "Catálogo / QR", "catalogo_admin"),
        ("recetas", "Recetas", "recetas"),
        ("caja", "Caja", "caja"),
        ("delivery", "Delivery", "delivery"),
        ("indicadores", "Indicadores", "indicadores"),
        ("reportes", "Reportes", "reportes"),
        ("admin", "Usuarios / Admin", "admin"),
        ("log", "Log", "logs"),
    ]

def page(content, active="dashboard"):
    return render_template_string(
        BASE_HTML,
        content=content,
        active=active,
        title=APP_TITLE,
        subtitle=APP_SUBTITLE,
        brand=BRAND,
        tabs=tabs(),
        money=money,
    )

def select_options(rows, value_key="id", text_key="nombre", selected=None):
    out = []
    for r in rows:
        val = str(r[value_key])
        sel = "selected" if selected is not None and str(selected) == val else ""
        out.append(f'<option value="{val}" {sel}>{r[text_key]}</option>')
    return "".join(out)

# =========================
# LÓGICA NEGOCIO
# =========================
def descontar_receta(producto_id, cantidad_producto):
    for r in q_all("SELECT * FROM recetas WHERE producto_id=?", (producto_id,)):
        q_exec(
            "UPDATE insumos SET stock=COALESCE(stock,0)-? WHERE id=?",
            (float(r["cantidad"] or 0) * float(cantidad_producto or 0), r["insumo_id"]),
        )

def descontar_producto(producto_id, cantidad):
    q_exec("UPDATE productos SET stock=COALESCE(stock,0)-? WHERE id=?", (float(cantidad or 0), producto_id))
    descontar_receta(producto_id, cantidad)

def crear_venta_desde_pedido(pedido_id, metodo_pago="EFECTIVO"):
    p = q_one("SELECT * FROM pedidos WHERE id=?", (pedido_id,))
    if not p:
        return None
    if p["pagado"] == "SI":
        return None
    vid = q_exec(
        "INSERT INTO ventas(fecha,hora,pedido_id,cliente,servicio,metodo_pago,subtotal,descuento,total,usuario,estado) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
        (today(), hour(), pedido_id, p["cliente"], p["servicio"], metodo_pago, p["subtotal"], p["descuento"], p["total"], session.get("user", ""), "PAGADO"),
    )
    for d in q_all("SELECT * FROM pedido_detalle WHERE pedido_id=?", (pedido_id,)):
        q_exec(
            "INSERT INTO venta_detalle(venta_id,producto_id,producto,cantidad,precio,total) VALUES(?,?,?,?,?,?)",
            (vid, d["producto_id"], d["producto"], d["cantidad"], d["precio"], d["total"]),
        )
        descontar_producto(d["producto_id"], d["cantidad"])
    q_exec("UPDATE pedidos SET pagado='SI', estado='PAGADO' WHERE id=?", (pedido_id,))
    q_exec(
        "INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario,venta_id) VALUES(?,?,?,?,?,?,?)",
        (today(), hour(), "INGRESO", f"VENTA PEDIDO {p['codigo']}", p["total"], session.get("user", ""), vid),
    )
    log_event("VENTA", f"Pedido cobrado {p['codigo']}")
    return vid

# =========================
# ROUTES AUTH
# =========================
@app.route("/", methods=["GET", "POST"])
def login():
    if session.get("user"):
        return redirect(url_for(safe_home()))
    if request.method == "POST":
        usuario = clean(request.form.get("usuario"))
        clave = request.form.get("clave", "")
        r = q_one("SELECT * FROM usuarios WHERE usuario=? AND activo=1", (usuario,))
        if r and check_password_hash(r["clave_hash"], clave):
            session["user"] = r["usuario"]
            session["nombre"] = r["nombre"]
            session["rol"] = r["rol"]
            log_event("LOGIN", "Ingreso correcto")
            return redirect(url_for(safe_home()))
        flash("Usuario o clave incorrectos.", "error")
    html = """
    <div class="login-page">
      <form class="login-card" method="post">
        <div style="font-size:44px;margin-bottom:8px">🍽️</div>
        <div class="logo">AOR<span>IX</span></div>
        <h2>Food Bank</h2>
        <p class="muted">Restaurante · Pizzería · Delivery · Caja</p>
        <label>Usuario</label><input name="usuario" placeholder="Ingrese su usuario" autofocus>
        <label>Clave</label><input name="clave" type="password" placeholder="Ingrese su clave">
        <button>Ingresar</button>
        <div class="hint">Acceso seguro creado por el administrador.</div>
      </form>
    </div>
    """
    return page(html)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# =========================
# DASHBOARD
# =========================
@app.route("/dashboard", methods=["GET", "POST"])
@login_required
def dashboard():
    if not is_admin():
        return redirect(url_for("ventas"))
    if request.method == "POST":
        set_ctx("sucursal", request.form.get("sucursal", "Sucursal Principal"))
        set_ctx("turno", request.form.get("turno", "MAÑANA"))
        set_ctx("dia_abierto", request.form.get("fecha", today()))
        flash("Contexto guardado.", "ok")
        return redirect(url_for("dashboard"))
    f = get_ctx("dia_abierto", today())
    ventas_hoy = q_one("SELECT COALESCE(SUM(total),0) t, COUNT(*) c FROM ventas WHERE fecha=?", (f,))
    pedidos_act = q_one("SELECT COUNT(*) c FROM pedidos WHERE fecha=? AND estado NOT IN ('ENTREGADO','PAGADO')", (f,))["c"]
    mesas = q_one("SELECT COUNT(DISTINCT mesa) c FROM pedidos WHERE fecha=? AND mesa<>'' AND estado NOT IN ('ENTREGADO','PAGADO')", (f,))["c"]
    stock_bajo = q_one("SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min")["c"]
    sucursales = q_all("SELECT * FROM sucursales WHERE activo=1 ORDER BY nombre")
    opts_suc = select_options(sucursales, "nombre", "nombre", get_ctx("sucursal"))
    html = f'''
    <form method="post" class="panel"><div class="section-title">🏪 Contexto de operación</div>
      <div class="main-context"><div><label>Sucursal</label><select name="sucursal">{opts_suc}</select></div><div><label>Fecha de trabajo</label><input type="date" name="fecha" value="{f}"></div><div><label>Turno</label><select name="turno"><option>MAÑANA</option><option>TARDE</option><option>NOCHE</option></select></div></div><br>
      <div class="actions"><button class="btn-primary">Guardar contexto</button><a class="btn" href="{url_for('dashboard')}">Recargar panel</a><a class="btn btn-success" href="{url_for('ventas')}">Ir a ventas</a></div>
      <div class="sku-help">Usuario y rol se controlan internamente por login, por eso ya no se muestran en el panel.</div>
    </form>
    <div class="actions" style="margin:12px 0"><h2 style="color:#047857;margin:0">🟢 DÍA ABIERTO: {f}</h2><a class="btn btn-danger" href="{url_for('cierre')}">🔒 Ir a cierre</a><a class="btn" href="{url_for('reportes',fi=f,ff=f)}">Ver reporte</a></div>
    <div class="panel"><div class="section-title">📊 Indicadores de hoy</div><div class="kpis"><div class="kpi dash-kpi"><h3>Ventas hoy</h3><b>{money(ventas_hoy['t'])}</b><p class="muted">Pedidos pagados: {ventas_hoy['c']}</p></div><div class="kpi dash-kpi"><h3>Pedidos activos</h3><b>{pedidos_act}</b><p class="muted">Pendientes/preparación/listos</p></div><div class="kpi dash-kpi"><h3>Mesas ocupadas</h3><b class="red">{mesas}</b><p class="muted">Salón en atención</p></div><div class="kpi dash-kpi"><h3>Stock bajo</h3><b class="red">{stock_bajo}</b><p class="muted">Productos por reponer</p></div></div></div>
    <div class="ops-strip"><div class="hint-card">✅ Ventas: salón, recojo o delivery.</div><div class="hint-card">✅ Pedidos: cocina, estados y quitar ítems.</div><div class="hint-card">✅ Indicadores: ventas, ticket y top productos.</div></div>'''
    return page(html, "dashboard")

@app.route("/cierre", methods=["GET", "POST"])
@login_required
def cierre():
    f = get_ctx("dia_abierto", today())
    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "cerrar":
            set_ctx("dia_cerrado", f)
            log_event("CIERRE", f"Día cerrado {f}")
            flash("Día cerrado correctamente. Ya puedes revisar el resumen y exportar reportes.", "ok")
        elif accion == "reabrir" and is_admin():
            set_ctx("dia_cerrado", "")
            log_event("REABRIR", f"Día reabierto {f}")
            flash("Día reabierto correctamente.", "ok")
        return redirect(url_for("cierre"))
    cerrado = get_ctx("dia_cerrado", "") == f
    ventas_hoy = q_one("SELECT COALESCE(SUM(total),0) t, COUNT(*) c FROM ventas WHERE fecha=?", (f,))
    pedidos = q_one("SELECT COUNT(*) c FROM pedidos WHERE fecha=?", (f,))["c"]
    pendientes = q_one("SELECT COUNT(*) c FROM pedidos WHERE fecha=? AND estado NOT IN ('PAGADO','ENTREGADO')", (f,))["c"]
    caja_ing = q_one("SELECT COALESCE(SUM(monto),0) t FROM caja WHERE fecha=? AND tipo='INGRESO'", (f,))["t"]
    caja_egr = q_one("SELECT COALESCE(SUM(monto),0) t FROM caja WHERE fecha=? AND tipo='EGRESO'", (f,))["t"]
    estado = "CERRADO" if cerrado else "ABIERTO"
    cls = "off" if cerrado else "ok"
    admin_btn = f'<button name="accion" value="reabrir" class="btn-orange">🔓 Reabrir día</button>' if is_admin() else ''
    html = f"""
    <div class="panel">
      <div class="box-title">🔒 Cierre del día</div><br>
      <div class="role-note">Estado actual: <span class="badge {cls}">{estado}</span> · Fecha: <b>{f}</b></div>
      <div class="kpis">
        <div class="kpi"><h3>Ventas del día</h3><b>{money(ventas_hoy['t'])}</b></div>
        <div class="kpi"><h3>Pedidos</h3><b>{pedidos}</b></div>
        <div class="kpi"><h3>Pendientes</h3><b class="red">{pendientes}</b></div>
        <div class="kpi"><h3>Caja neta</h3><b>{money(float(caja_ing or 0)-float(caja_egr or 0))}</b></div>
      </div><br>
      <form method="post" class="actions">
        <button name="accion" value="cerrar" class="btn-red">🔒 Cerrar día</button>
        {admin_btn}
        <a class="btn" href="{url_for('reportes',fi=f,ff=f)}">📄 Ver reporte del día</a>
        <a class="btn" href="{url_for('export_excel',fi=f,ff=f)}">📊 Exportar Excel</a>
      </form>
    </div>
    <div class="panel"><div class="box-title">Recomendación operativa</div><p>Antes de cerrar, verifica pedidos pendientes, caja y stock bajo. El cierre es lógico: no borra datos y el administrador puede reabrir el día.</p></div>
    """
    return page(html, "cierre")

# =========================
# VENTAS
# =========================
@app.route("/ventas", methods=["GET", "POST"])
@login_required
def ventas():
    if request.method == "POST":
        accion = request.form.get("accion", "guardar_pedido")
        if accion == "importar_productos" and "archivo" in request.files:
            count = importar_productos(request.files["archivo"])
            flash(f"Carga inicio día completada: {count} productos actualizados.", "ok")
            return redirect(url_for("ventas"))
        if accion == "guardar_pedido":
            producto_id = int(request.form.get("producto_id") or 0)
            prod = q_one("SELECT * FROM productos WHERE id=? AND activo=1", (producto_id,))
            if not prod:
                flash("Selecciona un producto válido.", "error")
                return redirect(url_for("ventas"))
            cantidad = float(request.form.get("cantidad") or 1)
            descuento = float(request.form.get("descuento") or 0)
            precio = float(prod["precio"] or 0)
            subtotal = precio * cantidad
            total = max(subtotal - descuento, 0)
            codigo = "PED-" + now().strftime("%Y%m%d-%H%M%S")
            pedido_id = q_exec("INSERT INTO pedidos(codigo,fecha,hora,mesa,cliente,telefono,direccion,referencia,servicio,metodo_pago,subtotal,descuento,total,estado,pagado,usuario,observacion) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (codigo, today(), hour(), request.form.get("mesa", ""), up(request.form.get("cliente") or "CLIENTE GENERAL"), clean(request.form.get("telefono")), up(request.form.get("direccion")), up(request.form.get("referencia")), request.form.get("servicio", "SALÓN"), request.form.get("metodo_pago", "EFECTIVO"), subtotal, descuento, total, "PENDIENTE", "NO", session.get("user"), ""))
            q_exec("INSERT INTO pedido_detalle(pedido_id,producto_id,producto,cantidad,precio,total) VALUES(?,?,?,?,?,?)", (pedido_id, producto_id, prod["nombre"], cantidad, precio, total))
            log_event("PEDIDO", f"Creado {codigo}")
            flash("Pedido guardado correctamente.", "ok")
            return redirect(url_for("ventas"))
        if accion == "cobrar_ticket":
            pedido_id = int(request.form.get("pedido_id") or 0)
            vid = crear_venta_desde_pedido(pedido_id, request.form.get("metodo_pago", "EFECTIVO"))
            flash("Pedido cobrado y ticket generado." if vid else "No se pudo cobrar: verifica que el pedido exista y no esté pagado.", "ok" if vid else "error")
            return redirect(url_for("ventas"))
    buscar_prod = clean(request.args.get("buscar", ""))
    if buscar_prod:
        like = f"%{buscar_prod.upper()}%"
        productos = q_all("SELECT * FROM productos WHERE activo=1 AND (UPPER(nombre) LIKE ? OR UPPER(categoria) LIKE ? OR UPPER(codigo) LIKE ?) ORDER BY nombre", (like, like, like))
    else:
        productos = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY CASE WHEN stock>0 THEN 0 ELSE 1 END,nombre")
    opts_prod = "".join(f'<option value="{p["id"]}">{p["codigo"] or p["id"]} · {p["nombre"]} · {money(p["precio"])} · Stock {p["stock"]}</option>' for p in productos)
    tr_prod = "".join(f'<tr class="{"row-ok" if float(r["stock"] or 0)>0 else "row-bad"}"><td>{r["codigo"]}</td><td>{r["nombre"]}</td><td>{r["categoria"]}</td><td>{money(r["precio"])}</td><td>{r["stock"]}</td><td>{"DISPONIBLE" if float(r["stock"] or 0)>0 else "SIN STOCK"}</td></tr>' for r in productos[:150]) or '<tr><td colspan="6">Sin productos.</td></tr>'
    pedidos_pend = q_all("SELECT * FROM pedidos WHERE pagado='NO' ORDER BY id DESC LIMIT 40")
    opts_ped = "".join(f'<option value="{p["id"]}">{p["codigo"]} · {p["cliente"]} · {money(p["total"])} · {p["estado"]}</option>' for p in pedidos_pend)
    detalle = q_all("SELECT d.*,p.codigo FROM pedido_detalle d JOIN pedidos p ON p.id=d.pedido_id ORDER BY d.id DESC LIMIT 30")
    tr_det = "".join(f'<tr><td>{r["codigo"]}</td><td>{r["producto"]}</td><td>{r["cantidad"]}</td><td>{money(r["precio"])}</td><td>{money(r["total"])}</td></tr>' for r in detalle) or '<tr><td colspan="5">Sin detalle.</td></tr>'
    html = f'''<div class="panel"><div class="section-title">🧾 Nueva venta / pedido</div><div class="hint-card">Registra ventas por salón, recojo o delivery. La carga del día ahora está más abajo, después del catálogo, para no estorbar la venta diaria.</div><br><form method="get" class="product-toolbar"><input name="buscar" value="{buscar_prod}" placeholder="Buscar producto por nombre, código o categoría"><button>🔎 Buscar producto</button><a class="btn" href="{url_for('ventas')}">Limpiar</a></form><form method="post"><input type="hidden" name="accion" value="guardar_pedido"><div class="clean-grid-4"><div><label>Mesa</label><select name="mesa"><option></option><option>MESA 1</option><option>MESA 2</option><option>MESA 3</option><option>MESA 4</option><option>MESA 5</option><option>MESA 6</option></select></div><div><label>Tipo servicio</label><select name="servicio"><option>SALÓN</option><option>DELIVERY</option><option>RECOJO</option></select></div><div><label>Cliente</label><input name="cliente" placeholder="Cliente general"></div><div><label>Teléfono</label><input name="telefono" placeholder="Celular"></div><div><label>Dirección</label><input name="direccion" placeholder="Solo para delivery"></div><div><label>Referencia</label><input name="referencia" placeholder="Referencia"></div><div><label>Producto</label><select name="producto_id">{opts_prod}</select></div><div><label>Cantidad</label><input name="cantidad" type="number" step="0.01" value="1"></div><div><label>Método pago</label><select name="metodo_pago"><option>EFECTIVO</option><option>YAPE</option><option>PLIN</option><option>TARJETA</option><option>TRANSFERENCIA</option></select></div><div><label>Descuento</label><input name="descuento" type="number" step="0.01" value="0.00"></div></div><br><div class="actions"><button class="primary">Guardar pedido</button><button type="reset">Limpiar venta</button><a class="btn" href="{url_for('inventario')}">Nuevo producto</a><a class="btn" href="{url_for('pedidos')}">Ver pedidos</a></div></form></div><div class="service-note"><div class="panel"><div class="box-title">Cobro rápido</div><br><form method="post" class="actions"><input type="hidden" name="accion" value="cobrar_ticket"><select name="pedido_id" style="max-width:580px">{opts_ped}</select><select name="metodo_pago" style="max-width:220px"><option>EFECTIVO</option><option>YAPE</option><option>PLIN</option><option>TARJETA</option><option>TRANSFERENCIA</option></select><button class="btn-success">Cobrar y ticket</button></form></div><div class="hint-card"><b>Delivery:</b> no se elimina porque sirve como base de clientes frecuentes y direcciones. El tipo servicio en Ventas solo marca la operación; el módulo Delivery funciona como CRM simple.</div></div><div class="panel"><div class="section-title">🍽️ Catálogo de productos</div><div class="table-wrap small"><table><thead><tr><th>Código</th><th>Producto</th><th>Categoría</th><th>Precio</th><th>Stock</th><th>Estado</th></tr></thead><tbody>{tr_prod}</tbody></table></div></div><div class="panel load-day-panel"><div class="section-title">📥 Carga de inicio de día</div><div class="hint-card">Importa Excel/CSV para actualizar productos, precios y stock antes de iniciar ventas. Esta sección queda abajo para mantener limpia la pantalla de venta.</div><br><form method="post" enctype="multipart/form-data" class="actions"><input type="hidden" name="accion" value="importar_productos"><input type="file" name="archivo" accept=".xlsx,.csv" style="max-width:420px"><button class="btn-warning">📥 Cargar día / importar Excel productos</button><a class="btn" href="{url_for('plantilla_inventario')}">📄 Descargar plantilla</a></form></div><div class="panel"><div class="section-title">🧩 Últimos ítems registrados</div><div class="table-wrap small"><table><thead><tr><th>Pedido</th><th>Producto</th><th>Cantidad</th><th>Precio</th><th>Subtotal</th></tr></thead><tbody>{tr_det}</tbody></table></div></div>'''
    return page(html, "ventas")

@app.route("/pedidos", methods=["GET", "POST"])
@login_required
def pedidos():
    if request.method == "POST":
        accion = request.form.get("accion")
        pedido_id = int(request.form.get("pedido_id") or 0)
        if not pedido_id:
            flash("Selecciona un pedido.", "error")
            return redirect(url_for("pedidos"))
        if accion == "estado":
            q_exec("UPDATE pedidos SET estado=? WHERE id=?", (request.form.get("estado", "PENDIENTE"), pedido_id))
            flash("Estado actualizado.", "ok")
        elif accion == "pagado":
            crear_venta_desde_pedido(pedido_id, request.form.get("metodo_pago", "EFECTIVO"))
            flash("Pedido marcado como pagado.", "ok")
        elif accion == "quitar_item":
            item_id = int(request.form.get("item_id") or 0)
            if item_id:
                q_exec("DELETE FROM pedido_detalle WHERE id=? AND pedido_id=?", (item_id, pedido_id))
                total = q_one("SELECT COALESCE(SUM(total),0) t FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))["t"]
                q_exec("UPDATE pedidos SET subtotal=?, total=max(?-COALESCE(descuento,0),0) WHERE id=?", (total, total, pedido_id))
                flash("Item retirado del pedido.", "ok")
        elif accion == "limpiar":
            q_exec("DELETE FROM pedido_detalle WHERE pedido_id=?", (pedido_id,))
            q_exec("DELETE FROM pedidos WHERE id=?", (pedido_id,))
            flash("Pedido eliminado.", "ok")
        return redirect(url_for("pedidos"))
    estado = request.args.get("estado", "TODOS")
    rows = q_all("SELECT * FROM pedidos ORDER BY id DESC LIMIT 200") if estado == "TODOS" else q_all("SELECT * FROM pedidos WHERE estado=? ORDER BY id DESC", (estado,))
    opts_p = '<option value="">Selecciona pedido</option>' + "".join(f'<option value="{r["id"]}">{r["codigo"]} · {r["cliente"] or "CLIENTE GENERAL"} · {r["estado"]} · {money(r["total"])}</option>' for r in rows)
    selected_first = rows[0]["id"] if rows else 0
    detalles = q_all("SELECT d.*,p.codigo FROM pedido_detalle d JOIN pedidos p ON p.id=d.pedido_id ORDER BY d.id DESC LIMIT 120")
    item_opts = '<option value="">Selecciona item a quitar</option>' + "".join(f'<option value="{r["id"]}">#{r["id"]} · {r["codigo"]} · {r["producto"]} · Cant. {r["cantidad"]}</option>' for r in detalles)
    trs = "".join(f'<tr><td>{r["id"]}</td><td>{r["codigo"]}</td><td>{r["fecha"]}</td><td>{r["hora"]}</td><td>{r["mesa"]}</td><td>{r["cliente"]}</td><td>{r["servicio"]}</td><td><span class="badge warn">{r["estado"]}</span></td><td>{money(r["total"])}</td><td>{r["pagado"]}</td></tr>' for r in rows) or '<tr><td colspan="10">Sin pedidos.</td></tr>'
    trd = "".join(f'<tr><td>{r["pedido_id"]}</td><td>{r["codigo"]}</td><td>{r["id"]}</td><td>{r["producto"]}</td><td>{r["cantidad"]}</td><td>{money(r["precio"])}</td><td>{money(r["total"])}</td></tr>' for r in detalles) or '<tr><td colspan="7">Sin detalle.</td></tr>'
    html = f'''<div class="panel"><div class="section-title">🚚 Control de pedidos / cocina</div><form method="get" class="clean-grid"><div><label>Filtrar por estado</label><select name="estado"><option>TODOS</option><option>PENDIENTE</option><option>PREPARACIÓN</option><option>LISTO</option><option>ENTREGADO</option><option>PAGADO</option></select></div><button>Refrescar</button><a class="btn" href="{url_for('ventas')}">Nuevo pedido</a></form><br><form method="post" class="pedido-actions"><div><label>Pedido</label><select name="pedido_id">{opts_p}</select></div><div><label>Cambiar a estado</label><select name="estado"><option>PREPARACIÓN</option><option>LISTO</option><option>ENTREGADO</option><option>PAGADO</option></select></div><button name="accion" value="estado" class="btn-success">Actualizar estado</button></form><br><form method="post" class="pedido-actions"><div><label>Pedido para cobrar o limpiar</label><select name="pedido_id">{opts_p}</select></div><div><label>Método pago</label><select name="metodo_pago"><option>EFECTIVO</option><option>YAPE</option><option>PLIN</option><option>TARJETA</option></select></div><div class="actions"><button name="accion" value="pagado" class="btn-success">Marcar pagado</button><button name="accion" value="limpiar" class="btn-danger" onclick="return confirm('¿Eliminar pedido completo?')">Eliminar pedido</button><a class="btn" href="{url_for('ticket', pedido_id=selected_first)}">Imprimir ticket</a></div></form><br><form method="post" class="pedido-actions-2"><div><label>Pedido</label><select name="pedido_id">{opts_p}</select></div><div><label>Ítem del pedido</label><select name="item_id">{item_opts}</select></div><button name="accion" value="quitar_item" class="btn-warning">➖ Quitar ítem</button></form><div class="sku-help">Se ordenó la pestaña: cada bloque tiene una función clara y se retiraron campos vacíos.</div></div><div class="panel"><div class="section-title">📋 Listado de pedidos</div><div class="table-wrap"><table><thead><tr><th>ID</th><th>Código</th><th>Fecha</th><th>Hora</th><th>Mesa</th><th>Cliente</th><th>Servicio</th><th>Estado</th><th>Total</th><th>Pagado</th></tr></thead><tbody>{trs}</tbody></table></div></div><div class="panel"><div class="section-title">🧾 Detalle de ítems</div><div class="table-wrap small"><table><thead><tr><th>Pedido ID</th><th>Código</th><th>Item</th><th>Producto</th><th>Cantidad</th><th>Precio</th><th>Subtotal</th></tr></thead><tbody>{trd}</tbody></table></div></div>'''
    return page(html, "pedidos")

@app.route("/ticket/<int:pedido_id>")
@login_required
def ticket(pedido_id):
    p = q_one("SELECT * FROM pedidos WHERE id=?", (pedido_id,)) if pedido_id else q_one("SELECT * FROM pedidos ORDER BY id DESC LIMIT 1")
    if not p:
        return "Sin pedido"
    det = q_all("SELECT * FROM pedido_detalle WHERE pedido_id=?", (p["id"],))
    lines = ["RESTAURANTE AORIX", "TICKET", f"Pedido: {p['codigo']}", f"Fecha: {p['fecha']} {p['hora']}", f"Cliente: {p['cliente']}", "-" * 32]
    for d in det:
        lines.append(f"{d['cantidad']} x {d['producto']} {money(d['total'])}")
    lines += ["-" * 32, f"TOTAL: {money(p['total'])}"]
    bio = BytesIO("\n".join(lines).encode("utf-8"))
    return send_file(bio, as_attachment=True, download_name=f"ticket_{p['codigo']}.txt", mimetype="text/plain")

# =========================
# INVENTARIO
# =========================
@app.route("/inventario", methods=["GET", "POST"])
@login_required
@admin_required
def inventario():
    if request.method == "POST":
        accion = request.form.get("accion", "producto")
        if accion == "producto":
            nombre = up(request.form.get("nombre"))
            if not nombre:
                flash("Producto obligatorio.", "error")
                return redirect(url_for("inventario"))
            codigo = clean(request.form.get("codigo")) or "P" + now().strftime("%H%M%S")
            existe = q_one("SELECT id FROM productos WHERE nombre=?", (nombre,))
            data = (codigo, nombre, up(request.form.get("categoria") or "PLATOS"), request.form.get("tipo", "VENTA"), up(request.form.get("unidad") or "PLATO"), float(request.form.get("precio") or 0), float(request.form.get("costo") or 0), float(request.form.get("stock") or 0), float(request.form.get("stock_min") or 0))
            if existe:
                q_exec("UPDATE productos SET codigo=?,categoria=?,tipo=?,unidad=?,precio=?,costo=?,stock=?,stock_min=?,activo=1 WHERE nombre=?", data[:1]+data[2:]+(nombre,))
            else:
                q_exec("INSERT INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)", data)
            flash("Producto guardado.", "ok")
        elif accion == "stock_in":
            pid = int(request.form.get("producto_id"))
            q_exec("UPDATE productos SET stock=stock+? WHERE id=?", (float(request.form.get("cantidad") or 0), pid))
            flash("Entrada de stock registrada.", "ok")
        elif accion == "stock_out":
            pid = int(request.form.get("producto_id"))
            q_exec("UPDATE productos SET stock=stock-? WHERE id=?", (float(request.form.get("cantidad") or 0), pid))
            flash("Salida de stock registrada.", "ok")
        elif accion == "importar" and "archivo" in request.files:
            file = request.files["archivo"]
            count = importar_productos(file)
            flash(f"Importación completada: {count} productos.", "ok")
        return redirect(url_for("inventario"))

    productos = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY CASE WHEN stock<=stock_min THEN 0 ELSE 1 END,nombre")
    opts_prod = select_options(productos)
    trp = "".join(f'<tr class="{"row-bad" if float(p["stock"] or 0)<=float(p["stock_min"] or 0) else ""}"><td>{p["id"]}</td><td>{p["nombre"]}</td><td>{p["categoria"]}</td><td>{p["tipo"]}</td><td>{p["unidad"]}</td><td>{money(p["precio"])}</td><td>{money(p["costo"])}</td><td>{p["stock"]}</td><td>{p["stock_min"]}</td><td>{"SIN STOCK" if float(p["stock"] or 0)<=0 else "OK"}</td></tr>' for p in productos)
    stock_bajo = q_one("SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min")["c"]
    html = f"""
    <div class="panel"><div class="section-title">📦 Productos e inventario</div><div class="hint-card"><b>Código/SKU:</b> puede usar letras y números. Recomendado para cadena de restaurantes: PIZ-MUZ-01, BEB-COCA-500, PL-001. Evita espacios y símbolos raros.</div><br>
      <form method="post" class="grid5"><input type="hidden" name="accion" value="producto">
        <div><label>Código/SKU</label><input name="codigo" placeholder="Ej. PIZ-MUZ-01"><div class="sku-help">Letras + números recomendado.</div></div><div><label>Nombre producto</label><input name="nombre" placeholder="Nombre producto"></div>
        <div><label>Categoría</label><select name="categoria"><option>PIZZAS</option><option>PLATOS</option><option>PARRILLAS</option><option>BEBIDAS</option><option>ADICIONALES</option><option>INSUMOS</option></select></div>
        <div><label>Tipo</label><select name="tipo"><option>VENTA</option><option>INSUMO</option></select></div><div><label>Unidad</label><select name="unidad"><option>PLATO</option><option>UND</option><option>KG</option><option>PORCION</option></select></div>
        <div><label>Precio venta</label><input name="precio" type="number" step="0.01"></div><div><label>Costo</label><input name="costo" type="number" step="0.01"></div><div><label>Stock</label><input name="stock" type="number" step="0.01"></div><div><label>Stock mínimo</label><input name="stock_min" type="number" step="0.01"></div><button>Guardar producto</button>
      </form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="stock_in"><select name="producto_id" style="max-width:380px">{opts_prod}</select><input name="cantidad" type="number" step="0.01" placeholder="Cantidad" style="max-width:160px"><button>Entrada stock</button></form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="stock_out"><select name="producto_id" style="max-width:380px">{opts_prod}</select><input name="cantidad" type="number" step="0.01" placeholder="Cantidad" style="max-width:160px"><button>Salida stock</button><a class="btn" href="{url_for('plantilla_inventario')}">Descargar plantilla</a><a class="btn" href="{url_for('export_inventario')}">Exportar inventario</a><b style="color:#c2410c">Alertas: {stock_bajo} con stock bajo</b></form><br>
      <form method="post" enctype="multipart/form-data" class="actions"><input type="hidden" name="accion" value="importar"><input type="file" name="archivo" accept=".xlsx,.csv" style="max-width:340px"><button class="btn-orange">Importar Excel</button></form>
    </div>
    <div class="panel"><div class="section-title">📋 Stock y productos</div><div class="table-wrap"><table><thead><tr><th>ID</th><th>Producto</th><th>Categoría</th><th>Tipo</th><th>Unidad</th><th>Precio</th><th>Costo</th><th>Stock</th><th>Mínimo</th><th>Estado</th></tr></thead><tbody>{trp}</tbody></table></div></div>
    """
    return page(html, "inventario")

def importar_productos(file_storage):
    count = 0
    filename = (file_storage.filename or "").lower()
    if filename.endswith(".xlsx") and OPENPYXL:
        wb = load_workbook(file_storage, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return 0
        headers = [up(h).replace(" ", "_") for h in rows[0]]
        for row in rows[1:]:
            d = dict(zip(headers, row))
            nombre = up(d.get("PRODUCTO") or d.get("NOMBRE") or "")
            if not nombre:
                continue
            codigo = clean(d.get("CODIGO") or d.get("CÓDIGO") or "") or "P" + str(count + 1).zfill(4)
            q_exec("INSERT OR REPLACE INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)", (codigo, nombre, up(d.get("CATEGORIA") or "PLATOS"), up(d.get("TIPO") or "VENTA"), up(d.get("UNIDAD") or "UND"), float(d.get("PRECIO") or 0), float(d.get("COSTO") or 0), float(d.get("STOCK") or 0), float(d.get("MINIMO") or d.get("STOCK_MIN") or 0)))
            count += 1
    else:
        content = file_storage.read().decode("utf-8-sig", errors="ignore")
        for d in csv.DictReader(StringIO(content)):
            nombre = up(d.get("PRODUCTO") or d.get("NOMBRE") or "")
            if not nombre:
                continue
            codigo = clean(d.get("CODIGO") or "") or "P" + str(count + 1).zfill(4)
            q_exec("INSERT OR REPLACE INTO productos(codigo,nombre,categoria,tipo,unidad,precio,costo,stock,stock_min,activo) VALUES(?,?,?,?,?,?,?,?,?,1)", (codigo, nombre, up(d.get("CATEGORIA") or "PLATOS"), up(d.get("TIPO") or "VENTA"), up(d.get("UNIDAD") or "UND"), float(d.get("PRECIO") or 0), float(d.get("COSTO") or 0), float(d.get("STOCK") or 0), float(d.get("MINIMO") or 0)))
            count += 1
    return count

# =========================
# RECETAS
# =========================
@app.route("/recetas", methods=["GET", "POST"])
@login_required
@admin_required
def recetas():
    if request.method == "POST":
        accion = request.form.get("accion", "agregar")
        if accion == "insumo":
            nombre = up(request.form.get("nombre"))
            if nombre:
                q_exec("INSERT OR REPLACE INTO insumos(nombre,unidad,stock,stock_min,costo,activo) VALUES(?,?,?,?,?,1)", (nombre, up(request.form.get("unidad") or "UND"), float(request.form.get("stock") or 0), float(request.form.get("stock_min") or 0), float(request.form.get("costo") or 0)))
                flash("Insumo guardado.", "ok")
        else:
            q_exec("INSERT INTO recetas(producto_id,insumo_id,cantidad,observacion) VALUES(?,?,?,?)", (int(request.form.get("producto_id")), int(request.form.get("insumo_id")), float(request.form.get("cantidad") or 0), up(request.form.get("observacion"))))
            flash("Insumo agregado a receta.", "ok")
        return redirect(url_for("recetas"))
    productos = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY nombre")
    insumos = q_all("SELECT * FROM insumos WHERE activo=1 ORDER BY nombre")
    opts_p = select_options(productos)
    opts_i = select_options(insumos)
    rows = q_all("SELECT r.id,p.nombre producto,i.nombre insumo,i.unidad,r.cantidad,r.observacion FROM recetas r JOIN productos p ON p.id=r.producto_id JOIN insumos i ON i.id=r.insumo_id ORDER BY p.nombre,i.nombre")
    trs = "".join(f'<tr><td>{r["id"]}</td><td>{r["producto"]}</td><td>{r["insumo"]}</td><td>{r["unidad"]}</td><td>{r["cantidad"]}</td><td>{r["observacion"]}</td></tr>' for r in rows) or '<tr><td colspan="6">Sin receta.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">Recetas detalladas</div><br>
      <form method="post" class="grid5"><input type="hidden" name="accion" value="agregar"><div><label>Producto de venta</label><select name="producto_id">{opts_p}</select></div><div><label>Insumo</label><select name="insumo_id">{opts_i}</select></div><div><label>Cantidad del insumo</label><input name="cantidad" type="number" step="0.0001" value="1"></div><div><label>Observación</label><input name="observacion"></div><button>Agregar insumo</button></form><br>
      <form method="post" class="grid5"><input type="hidden" name="accion" value="insumo"><div><label>Nuevo insumo</label><input name="nombre" placeholder="Ej. AJI"></div><div><label>Unidad</label><select name="unidad"><option>PORCION</option><option>UND</option><option>KG</option></select></div><div><label>Stock</label><input name="stock" type="number" step="0.01"></div><div><label>Mínimo</label><input name="stock_min" type="number" step="0.01"></div><button>Guardar insumo</button></form>
    </div>
    <div class="panel"><div class="box-title">Detalle de receta</div><br><div class="table-wrap"><table><thead><tr><th>ID</th><th>Producto</th><th>Insumo</th><th>Unidad</th><th>Cantidad</th><th>Observación</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    """
    return page(html, "recetas")

# =========================
# CAJA
# =========================
@app.route("/caja", methods=["GET", "POST"])
@login_required
@admin_required
def caja():
    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "abrir":
            monto = float(request.form.get("monto_apertura") or 0)
            set_ctx("caja_abierta", "1")
            set_ctx("monto_apertura", monto)
            q_exec("INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario) VALUES(?,?,?,?,?,?)", (today(), hour(), "INGRESO", "APERTURA CAJA", monto, session.get("user")))
            flash("Caja abierta.", "ok")
        elif accion == "cerrar":
            set_ctx("caja_abierta", "0")
            flash("Caja cerrada.", "ok")
        elif accion == "gasto":
            q_exec("INSERT INTO caja(fecha,hora,tipo,concepto,monto,usuario) VALUES(?,?,?,?,?,?)", (today(), hour(), "EGRESO", up(request.form.get("concepto")), float(request.form.get("monto") or 0), session.get("user")))
            flash("Gasto registrado.", "ok")
        return redirect(url_for("caja"))
    f = request.args.get("fecha", today())
    rows = q_all("SELECT * FROM caja WHERE fecha=? ORDER BY id DESC", (f,))
    ingresos = q_one("SELECT COALESCE(SUM(monto),0) t FROM caja WHERE fecha=? AND tipo='INGRESO'", (f,))["t"]
    egresos = q_one("SELECT COALESCE(SUM(monto),0) t FROM caja WHERE fecha=? AND tipo='EGRESO'", (f,))["t"]
    trs = "".join(f'<tr><td>{r["hora"]}</td><td>{r["tipo"]}</td><td>{r["concepto"]}</td><td>{money(r["monto"])}</td><td>{r["usuario"]}</td></tr>' for r in rows) or '<tr><td colspan="5">Sin movimientos.</td></tr>'
    estado = "ABIERTA" if get_ctx("caja_abierta", "0") == "1" else "SIN ABRIR"
    html = f"""
    <div class="panel"><div class="box-title">Caja</div><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="abrir"><label>Monto apertura:</label><input name="monto_apertura" type="number" step="0.01" value="100.00" style="max-width:180px"><button>Abrir caja</button></form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="cerrar"><button>Cerrar caja</button></form><br>
      <form method="post" class="actions"><input type="hidden" name="accion" value="gasto"><label>Egreso:</label><input name="concepto" placeholder="Concepto" style="max-width:240px"><input name="monto" type="number" step="0.01" value="0.00" style="max-width:180px"><button>Registrar gasto</button></form>
    </div>
    <div class="panel"><div class="box-title">Resumen de caja</div><h2 style="color:#dc2626">Caja: {estado}</h2><div class="grid"><div>Apertura: <b>{money(get_ctx('monto_apertura','0'))}</b></div><div>Efectivo sistema: <b>{money(ingresos-egresos)}</b></div><div>Ingresos: <b>{money(ingresos)}</b></div><div>Gastos: <b>{money(egresos)}</b></div></div></div>
    <div class="panel"><form method="get" class="actions"><input type="date" name="fecha" value="{f}" style="max-width:180px"><button>Filtrar</button></form><br><div class="table-wrap"><table><thead><tr><th>Hora</th><th>Tipo</th><th>Concepto</th><th>Monto</th><th>Usuario</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    """
    return page(html, "caja")

# =========================
# DELIVERY / CLIENTES
# =========================
@app.route("/delivery", methods=["GET", "POST"])
@login_required
@admin_required
def delivery():
    if request.method == "POST":
        q_exec("INSERT INTO clientes(nombre,telefono,direccion,referencia,notas,activo) VALUES(?,?,?,?,?,1)", (up(request.form.get("nombre")), clean(request.form.get("telefono")), up(request.form.get("direccion")), up(request.form.get("referencia")), up(request.form.get("notas"))))
        flash("Cliente guardado.", "ok")
        return redirect(url_for("delivery"))
    rows = q_all("SELECT * FROM clientes WHERE activo=1 ORDER BY id DESC")
    trs = "".join(f'<tr><td>{r["id"]}</td><td>{r["nombre"]}</td><td>{r["telefono"]}</td><td>{r["direccion"]}</td><td>{r["referencia"]}</td><td>{r["notas"]}</td></tr>' for r in rows) or '<tr><td colspan="6">Sin clientes.</td></tr>'
    html = f"""
    <div class="delivery-advice">🛵 Delivery no debe retirarse: úsalo como CRM de clientes frecuentes, teléfonos, direcciones y referencias. En Ventas solo eliges el tipo de servicio; aquí administras la base de clientes.</div><br>
    <div class="panel"><div class="section-title">🛵 Clientes delivery / frecuentes</div><form method="post" class="clean-grid">
      <div><label>Nombre</label><input name="nombre" placeholder="Cliente"></div><div><label>Teléfono</label><input name="telefono" placeholder="Celular"></div><div><label>Dirección</label><input name="direccion" placeholder="Dirección"></div><div><label>Referencia</label><input name="referencia" placeholder="Referencia"></div><div><label>Notas</label><input name="notas" placeholder="Ej. portón negro, paga Yape"></div><button class="btn-success">Guardar cliente</button><a class="btn" href="{url_for('ventas')}">Cargar en venta</a><a class="btn" href="{url_for('delivery')}">Refrescar</a></form></div>
    <div class="panel"><div class="section-title">📒 Base de clientes frecuentes</div><div class="table-wrap"><table><thead><tr><th>ID</th><th>Nombre</th><th>Teléfono</th><th>Dirección</th><th>Referencia</th><th>Notas</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    """
    return page(html, "delivery")


# =========================
# CATÁLOGO ONLINE / IMÁGENES / QR
# =========================
def catalog_public_url():
    try:
        return url_for("menu_publico", _external=True)
    except Exception:
        return "/menu"

def save_catalog_image(file_storage):
    if not file_storage or not getattr(file_storage, "filename", ""):
        return ""
    filename = secure_filename(file_storage.filename)
    ext = os.path.splitext(filename)[1].lower()
    if ext not in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
        raise ValueError("Formato no permitido. Usa PNG, JPG, JPEG, WEBP o GIF.")
    new_name = f"{uuid.uuid4().hex}{ext}"
    path = os.path.join(CATALOG_DIR, new_name)
    file_storage.save(path)
    return f"catalogo/{new_name}"

def qr_svg_data(text):
    try:
        import qrcode
        import qrcode.image.svg
        factory = qrcode.image.svg.SvgImage
        img = qrcode.make(text, image_factory=factory, box_size=10)
        bio = BytesIO()
        img.save(bio)
        import base64
        return "data:image/svg+xml;base64," + base64.b64encode(bio.getvalue()).decode("ascii")
    except Exception:
        return ""

@app.route("/catalogo", methods=["GET", "POST"])
@login_required
@admin_required
def catalogo_admin():
    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "config":
            set_ctx("negocio_nombre", up(request.form.get("negocio_nombre")) or "EL TORO RESTAURANT GRILL")
            set_ctx("catalogo_slug", clean(request.form.get("catalogo_slug")) or "el-toro-restaurant-grill")
            flash("Datos del catálogo actualizados.", "ok")
        elif accion == "crear_item":
            try:
                imagen = save_catalog_image(request.files.get("imagen"))
                q_exec("INSERT INTO catalogo_publico(titulo,descripcion,precio,categoria,imagen,destacado,activo) VALUES(?,?,?,?,?,?,1)", (
                    up(request.form.get("titulo")),
                    clean(request.form.get("descripcion")),
                    float(request.form.get("precio") or 0),
                    up(request.form.get("categoria") or "PLATOS"),
                    imagen,
                    1 if request.form.get("destacado") else 0,
                ))
                flash("Producto con imagen agregado al catálogo.", "ok")
            except Exception as ex:
                flash(str(ex), "error")
        elif accion == "desactivar_item":
            q_exec("UPDATE catalogo_publico SET activo=0 WHERE id=?", (int(request.form.get("id") or 0),))
            flash("Imagen/producto retirado del catálogo.", "ok")
        return redirect(url_for("catalogo_admin"))

    negocio = get_ctx("negocio_nombre", "EL TORO RESTAURANT GRILL")
    slug = get_ctx("catalogo_slug", "el-toro-restaurant-grill")
    url = catalog_public_url()
    qr = qr_svg_data(url)
    items = q_all("SELECT * FROM catalogo_publico WHERE activo=1 ORDER BY destacado DESC,id DESC")

    cards = ""
    for i in items:
        img = url_for("static", filename=(i["imagen"] or "toro_logo.png"))
        cards += f"""<div class='catalog-card'><img src='{img}'><h3>{i['titulo']}</h3><p>{i['descripcion'] or i['categoria']}</p><span class='price-pill'>{money(i['precio'])}</span><form method='post' style='margin-top:10px'><input type='hidden' name='accion' value='desactivar_item'><input type='hidden' name='id' value='{i['id']}'><button class='btn-danger' onclick=\"return confirm('¿Retirar del catálogo?')\">Retirar</button></form></div>"""
    if not cards:
        cards = "<div class='hint-card'>Aún no tienes imágenes. Carga tus platos, pizzas, combos o bebidas para publicarlos.</div>"

    phone_items = ""
    for i in items[:4]:
        img = url_for("static", filename=(i["imagen"] or "toro_logo.png"))
        phone_items += f"""<div class='phone-food'><img src='{img}'><b>{i['titulo']}</b><small>{money(i['precio'])}</small></div>"""
    if not phone_items:
        phone_items = f"""<div class='phone-food'><img src='{url_for('static', filename='toro_logo.png')}'><b>Sube tu plato</b><small>Desde S/ 0.00</small></div>"""
    qr_html = f"<img class='qr-img' src='{qr}'>" if qr else "<div class='qr-img' style='display:grid;place-items:center'>QR</div>"

    html = f"""
    <div class="catalog-hero">
      <div><h2>Catálogo online con imágenes y QR</h2><p>Publica platos, pizzas, combos y promociones con estilo tipo app. Tus clientes pueden abrir el link o escanear el QR.</p><div class="actions"><a class="btn-success" href="{url_for('menu_publico')}" target="_blank">Abrir catálogo público</a><button onclick="navigator.clipboard.writeText('{url}')" class="btn-warning">Copiar link</button></div></div>
      <div class="catalog-phone"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px"><b>{negocio}</b><span>☰</span></div><div class="phone-grid">{phone_items}</div></div>
    </div><br>
    <div class="grid2">
      <div class="panel upload-drop"><div class="section-title">Cargar imagen de producto</div><form method="post" enctype="multipart/form-data" class="clean-grid"><input type="hidden" name="accion" value="crear_item"><div><label>Nombre del plato</label><input name="titulo" required placeholder="Ej. Parrilla familiar / Pizza americana"></div><div><label>Categoría</label><select name="categoria"><option>PLATOS</option><option>PIZZAS</option><option>PARRILLAS</option><option>BEBIDAS</option><option>COMBOS</option><option>POSTRES</option></select></div><div><label>Precio</label><input name="precio" type="number" step="0.01" value="0.00"></div><div><label>Imagen</label><input name="imagen" type="file" accept="image/*" required></div><div style="grid-column:1/-1"><label>Descripción</label><textarea name="descripcion" placeholder="Ingredientes, tamaño, promoción, etc."></textarea></div><label style="display:flex;gap:8px;align-items:center"><input type="checkbox" name="destacado" style="width:auto"> Destacar</label><button class="btn-success">Agregar al catálogo</button></form></div>
      <div class="panel"><div class="section-title">Compartir catálogo</div><div class="qr-box">{qr_html}<div><b>Link público</b><input readonly value="{url}" onclick="this.select()"><p class="muted">Agrega este link en WhatsApp, Facebook, Instagram o imprímelo en mesa.</p><a class="btn-success" target="_blank" href="https://wa.me/?text={url}">Compartir por WhatsApp</a></div></div><hr><form method="post" class="clean-grid"><input type="hidden" name="accion" value="config"><div><label>Nombre del negocio</label><input name="negocio_nombre" value="{negocio}"></div><div><label>Enlace corto</label><input name="catalogo_slug" value="{slug}"></div><button class="btn-warning">Guardar configuración</button></form></div>
    </div><br>
    <div class="panel"><div class="section-title">Productos publicados</div><div class="catalog-grid">{cards}</div></div>
    """
    return page(html, "catalogo")

@app.route("/menu")
def menu_publico():
    negocio = get_ctx("negocio_nombre", "EL TORO RESTAURANT GRILL")
    items = q_all("SELECT * FROM catalogo_publico WHERE activo=1 ORDER BY destacado DESC,id DESC")
    css = BASE_HTML.split("<style>")[1].split("</style>")[0]
    cards = ""
    for i in items:
        img = url_for("static", filename=(i["imagen"] or "toro_logo.png"))
        cards += f"""<div class='catalog-card'><img src='{img}'><h3>{i['titulo']}</h3><p>{i['descripcion'] or i['categoria']}</p><span class='price-pill'>{money(i['precio'])}</span></div>"""
    if not cards:
        cards = "<div class='hint-card'>Catálogo en preparación.</div>"
    logo = url_for('static', filename='toro_logo.png')
    return f"""<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{negocio}</title><style>{css}</style></head><body style='background:#101016;color:white'><div style='max-width:1100px;margin:auto;padding:24px'><div class='catalog-hero'><div><img src='{logo}' style='width:120px'><h2>{negocio}</h2><p>Menú digital · Platos · Pizzas · Parrillas · Promociones</p></div></div><br><div class='catalog-grid'>{cards}</div></div></body></html>"""

# =========================
# INDICADORES
# =========================
@app.route("/indicadores")
@login_required
@admin_required
def indicadores():
    try:
        fi = request.args.get("fi", today()); ff = request.args.get("ff", fi)
        ventas = q_one("SELECT COALESCE(SUM(total),0) t, COUNT(*) c FROM ventas WHERE fecha BETWEEN ? AND ?", (fi, ff)) or {"t":0,"c":0}
        pendientes = q_one("SELECT COUNT(*) c FROM pedidos WHERE fecha BETWEEN ? AND ? AND estado NOT IN ('PAGADO','ENTREGADO')", (fi, ff))["c"]
        stock_bajo = q_one("SELECT COUNT(*) c FROM productos WHERE activo=1 AND stock<=stock_min")["c"]
        ticket = (float(ventas["t"] or 0) / int(ventas["c"] or 1)) if int(ventas["c"] or 0) else 0
        rows = q_all("SELECT fecha periodo, COALESCE(SUM(total),0) ventas, COUNT(*) pedidos FROM ventas WHERE fecha BETWEEN ? AND ? GROUP BY fecha ORDER BY fecha", (fi, ff))
        pagos = q_all("SELECT metodo_pago, COALESCE(SUM(total),0) total FROM ventas WHERE fecha BETWEEN ? AND ? GROUP BY metodo_pago ORDER BY total DESC", (fi, ff))
        top = q_all("SELECT d.producto producto, SUM(d.cantidad) cant, SUM(d.total) total FROM venta_detalle d JOIN ventas v ON v.id=d.venta_id WHERE v.fecha BETWEEN ? AND ? GROUP BY d.producto ORDER BY cant DESC LIMIT 7", (fi, ff))
        maxv = max([float(r["ventas"] or 0) for r in rows] + [1])
        bars = "".join(f'<div class="bar-wrap"><div class="bar" style="height:{max(28, int((float(r["ventas"] or 0)/maxv)*250))}px">{money(r["ventas"])}</div><small>{r["periodo"]}</small></div>' for r in rows) or '<div class="muted">Sin ventas en el periodo. Registra ventas para ver la gráfica.</div>'
        trs = "".join(f'<tr><td>{r["periodo"]}</td><td>{money(r["ventas"])}</td><td>{r["pedidos"]}</td><td>{money(float(r["ventas"] or 0)/int(r["pedidos"] or 1))}</td></tr>' for r in rows) or '<tr><td colspan="4">Sin detalle.</td></tr>'
        trp = "".join(f'<tr><td>{p["metodo_pago"]}</td><td>{money(p["total"])}</td></tr>' for p in pagos) or '<tr><td colspan="2">Sin pagos.</td></tr>'
        trt = "".join(f'<tr><td>{t["producto"]}</td><td>{t["cant"]}</td><td>{money(t["total"])}</td></tr>' for t in top) or '<tr><td colspan="3">Sin productos vendidos.</td></tr>'
        html = f'''<div class="panel analytics-hero"><div class="section-title">📈 Indicadores para cadena de restaurantes y pizzerías</div><form method="get" class="clean-grid"><div><label>Fecha inicio</label><input type="date" name="fi" value="{fi}"></div><div><label>Fecha fin</label><input type="date" name="ff" value="{ff}"></div><div class="actions"><button class="btn-success">Actualizar</button><a class="btn" href="{url_for('indicadores')}">Hoy</a></div></form></div><div class="analytics-grid"><div class="analytics-card"><span>Ventas netas</span><b>{money(ventas['t'])}</b><small>Ingreso del periodo</small></div><div class="analytics-card"><span>Pedidos pagados</span><b>{ventas['c']}</b><small>Transacciones cobradas</small></div><div class="analytics-card"><span>Ticket promedio</span><b>{money(ticket)}</b><small>Venta promedio por pedido</small></div><div class="analytics-card"><span>Pendientes / stock bajo</span><b>{pendientes} / {stock_bajo}</b><small>Operación y abastecimiento</small></div></div><div class="grid2"><div class="panel"><div class="section-title">📊 Ventas por día</div><div class="chart-pro">{bars}</div></div><div class="panel"><div class="section-title">💳 Ventas por método de pago</div><div class="table-wrap small"><table><thead><tr><th>Método</th><th>Total</th></tr></thead><tbody>{trp}</tbody></table></div><br><div class="hint-card">Mide ticket promedio, top productos, ventas por día y pagos para controlar rentabilidad, stock y velocidad de atención.</div></div></div><div class="grid2"><div class="panel"><div class="section-title">🏆 Top productos vendidos</div><div class="table-wrap small"><table><thead><tr><th>Producto</th><th>Cantidad</th><th>Total</th></tr></thead><tbody>{trt}</tbody></table></div></div><div class="panel"><div class="section-title">📅 Detalle por periodo</div><div class="table-wrap small"><table><thead><tr><th>Periodo</th><th>Ventas S/</th><th>Pedidos</th><th>Ticket promedio</th></tr></thead><tbody>{trs}</tbody></table></div></div></div>'''
        return page(html, "indicadores")
    except Exception as ex:
        log_event("ERROR INDICADORES", str(ex))
        return page(f'<div class="panel"><h2>Indicadores no disponibles</h2><p>Se evitó el error controlado. Detalle técnico registrado en Log.</p><a class="btn" href="{url_for("dashboard")}">Volver al panel</a></div>', "indicadores")

@app.route("/reportes")
@login_required
@admin_required
def reportes():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    ventas = q_all("SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora", (fi, ff))
    stock = q_all("SELECT * FROM productos WHERE activo=1 AND stock<=stock_min ORDER BY nombre")
    total = sum(float(v["total"] or 0) for v in ventas)
    pedidos_pagados = len(ventas)
    ticket_prom = total / pedidos_pagados if pedidos_pagados else 0
    trs = "".join(f'<tr><td>{v["fecha"]}</td><td>{v["hora"]}</td><td>{v["cliente"] or "CLIENTE GENERAL"}</td><td>{v["servicio"]}</td><td>{v["metodo_pago"]}</td><td>{money(v["total"])}</td><td>{v["usuario"]}</td></tr>' for v in ventas) or '<tr><td colspan="7">Sin ventas en este periodo.</td></tr>'
    trs_stock = "".join(f'<tr class="row-bad"><td>{s["nombre"]}</td><td>{s["categoria"]}</td><td>{s["stock"]} {s["unidad"]}</td><td>{s["stock_min"]}</td></tr>' for s in stock) or '<tr><td colspan="4">Sin stock bajo.</td></tr>'
    html = f"""
    <div class="panel"><div class="box-title">📄 Reportería ejecutiva por sucursal</div><br><form method="get" class="actions"><label>Fecha inicio:</label><input type="date" name="fi" value="{fi}" style="max-width:170px"><label>Fecha fin:</label><input type="date" name="ff" value="{ff}" style="max-width:170px"><button>Generar</button><a class="btn btn-green" href="{url_for('export_excel',fi=fi,ff=ff)}">Exportar Excel</a><a class="btn" href="{url_for('export_csv',fi=fi,ff=ff)}">Exportar CSV</a></form></div>
    <div class="report-card-grid"><div class="mini-report">Ventas<br><b>{money(total)}</b></div><div class="mini-report">Pedidos pagados<br><b>{pedidos_pagados}</b></div><div class="mini-report">Ticket promedio<br><b>{money(ticket_prom)}</b></div><div class="mini-report">Stock bajo<br><b>{len(stock)}</b></div></div>
    <div class="panel"><div class="box-title">Ventas del periodo</div><br><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Cliente</th><th>Servicio</th><th>Pago</th><th>Total</th><th>Usuario</th></tr></thead><tbody>{trs}</tbody></table></div></div>
    <div class="panel"><div class="box-title">Alertas de stock bajo</div><br><div class="table-wrap small"><table><thead><tr><th>Producto</th><th>Categoría</th><th>Stock</th><th>Mínimo</th></tr></thead><tbody>{trs_stock}</tbody></table></div></div>
    """
    return page(html, "reportes")

@app.route("/export_excel")
@login_required
@admin_required
def export_excel():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    rows = q_all("SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora", (fi, ff))
    if OPENPYXL:
        wb = Workbook()
        ws = wb.active
        ws.title = "Ventas"
        ws.append(["ID", "Fecha", "Hora", "Cliente", "Servicio", "Pago", "Subtotal", "Descuento", "Total", "Usuario", "Estado"])
        for r in rows:
            ws.append([r["id"], r["fecha"], r["hora"], r["cliente"], r["servicio"], r["metodo_pago"], r["subtotal"], r["descuento"], r["total"], r["usuario"], r["estado"]])
        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)
        return send_file(bio, as_attachment=True, download_name=f"reporte_aorix_{fi}_{ff}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return redirect(url_for("export_csv", fi=fi, ff=ff))

@app.route("/export_csv")
@login_required
@admin_required
def export_csv():
    fi = request.args.get("fi", today())
    ff = request.args.get("ff", fi)
    rows = q_all("SELECT * FROM ventas WHERE fecha BETWEEN ? AND ? ORDER BY fecha,hora", (fi, ff))
    out = StringIO()
    w = csv.writer(out)
    w.writerow(["ID", "Fecha", "Hora", "Cliente", "Servicio", "Pago", "Total"])
    for r in rows:
        w.writerow([r["id"], r["fecha"], r["hora"], r["cliente"], r["servicio"], r["metodo_pago"], r["total"]])
    return send_file(BytesIO(out.getvalue().encode("utf-8-sig")), as_attachment=True, download_name=f"reporte_aorix_{fi}_{ff}.csv", mimetype="text/csv")

@app.route("/export_inventario")
@login_required
@admin_required
def export_inventario():
    rows = q_all("SELECT * FROM productos WHERE activo=1 ORDER BY nombre")
    out = StringIO()
    w = csv.writer(out)
    w.writerow(["CODIGO", "PRODUCTO", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "MINIMO"])
    for r in rows:
        w.writerow([r["codigo"], r["nombre"], r["categoria"], r["tipo"], r["unidad"], r["precio"], r["costo"], r["stock"], r["stock_min"]])
    return send_file(BytesIO(out.getvalue().encode("utf-8-sig")), as_attachment=True, download_name="inventario_aorix.csv", mimetype="text/csv")

@app.route("/plantilla_inventario")
@login_required
@admin_required
def plantilla_inventario():
    out = StringIO()
    w = csv.writer(out)
    w.writerow(["CODIGO", "PRODUCTO", "CATEGORIA", "TIPO", "UNIDAD", "PRECIO", "COSTO", "STOCK", "MINIMO"])
    w.writerow(["P001", "1/2 POLLO", "PLATOS", "VENTA", "PLATO", 35, 10, 40, 5])
    return send_file(BytesIO(out.getvalue().encode("utf-8-sig")), as_attachment=True, download_name="plantilla_inventario_aorix.csv", mimetype="text/csv")

# =========================
# ADMIN / LOG
# =========================
@app.route("/admin", methods=["GET", "POST"])
@login_required
@admin_required
def admin():
    if request.method == "POST":
        accion = request.form.get("accion")
        if accion == "sucursal":
            nombre = up(request.form.get("nombre"))
            if nombre:
                q_exec("INSERT OR IGNORE INTO sucursales(nombre,activo) VALUES(?,1)", (nombre,))
                flash("Sucursal registrada.", "ok")
        elif accion == "usuario":
            usuario = clean(request.form.get("usuario"))
            clave = clean(request.form.get("clave"))
            nombre = up(request.form.get("nombre"))
            rol = request.form.get("rol", "MESERO")
            if usuario and clave:
                q_exec("INSERT OR REPLACE INTO usuarios(usuario,nombre,clave_hash,rol,activo) VALUES(?,?,?,?,1)", (usuario, nombre, generate_password_hash(clave), rol))
                flash("Usuario guardado.", "ok")
        elif accion == "desactivar_usuario":
            usuario = clean(request.form.get("usuario"))
            if usuario and usuario != session.get("user"):
                q_exec("UPDATE usuarios SET activo=0 WHERE usuario=?", (usuario,))
                log_event("USUARIO DESACTIVADO", usuario)
                flash("Usuario desactivado correctamente.", "ok")
        elif accion == "sembrar":
            init_db()
            flash("Demo sembrada / verificada.", "ok")
        return redirect(url_for("admin"))
    sucursales = q_all("SELECT * FROM sucursales ORDER BY nombre")
    usuarios = q_all("SELECT usuario,nombre,rol,activo FROM usuarios ORDER BY usuario")
    tr_s = "".join(f'<tr><td>{s["id"]}</td><td>{s["nombre"]}</td><td>{s["activo"]}</td></tr>' for s in sucursales)
    tr_u = "".join(f'<tr><td>{u["usuario"]}</td><td>{u["nombre"]}</td><td>{u["rol"]}</td><td>{u["activo"]}</td><td>{"" if u["usuario"]==session.get("user") else f"<form method=\"post\" style=\"margin:0\"><input type=\"hidden\" name=\"accion\" value=\"desactivar_usuario\"><input type=\"hidden\" name=\"usuario\" value=\"{u["usuario"]}\"><button class=\"btn-red\" onclick=\"return confirm(\'¿Desactivar usuario?\')\">Desactivar</button></form>"}</td></tr>' for u in usuarios)
    html = f"""
    <div class="role-note">✅ Administración limpia: crea usuarios, sucursales y limita accesos. No ADMIN solo verá Venta, Pedido, Cierre y Salir.</div>
    <div class="admin-panels">
      <div class="panel"><div class="section-title">👤 Crear / actualizar usuario</div>
        <form method="post" class="clean-grid">
          <input type="hidden" name="accion" value="usuario">
          <div><label>Usuario</label><input name="usuario" required placeholder="ej: vendedor1"></div>
          <div><label>Nombre</label><input name="nombre" placeholder="Nombre completo"></div>
          <div><label>Clave</label><input name="clave" type="password" required placeholder="Clave"></div>
          <div><label>Rol</label><select name="rol"><option>OPERADOR</option><option>MESERO</option><option>CAJA</option><option>COCINA</option><option>ADMIN</option></select></div>
          <button class="btn-blue">Guardar usuario</button>
        </form>
      </div>
      <div class="panel"><div class="section-title">🏬 Sucursales</div>
        <form method="post" class="actions"><input type="hidden" name="accion" value="sucursal"><input name="nombre" placeholder="Nueva sucursal"><button class="btn-green">Registrar sucursal</button><button name="accion" value="sembrar" class="btn-orange">Sembrar demo</button></form>
        <br><div class="hint-card">Pensado para cadena: sucursales, usuarios por rol y control por turno.</div>
      </div>
    </div>
    <div class="food-card-grid">
      <div class="food-card"><div class="thumb">🔐</div><b>Roles claros</b><small>Admin controla todo; operadores trabajan ventas/pedidos/cierre.</small></div>
      <div class="food-card"><div class="thumb">🏪</div><b>Sucursales</b><small>Base para reportes por local y turnos.</small></div>
      <div class="food-card"><div class="thumb">📱</div><b>Celular</b><small>Botones grandes y navegación tipo app.</small></div>
    </div><br>
    <div class="grid2">
      <div class="panel"><div class="section-title">🏬 Sucursales creadas</div><div class="table-wrap small"><table><thead><tr><th>ID</th><th>Sucursal</th><th>Activo</th></tr></thead><tbody>{tr_s}</tbody></table></div></div>
      <div class="panel users-clean-table"><div class="section-title">👥 Usuarios creados</div><div class="table-wrap small"><table><thead><tr><th>Usuario</th><th>Nombre</th><th>Rol</th><th>Activo</th><th>Acción</th></tr></thead><tbody>{tr_u}</tbody></table></div></div>
    </div>
    <div class="panel"><div class="section-title">📌 Resumen funcional</div><textarea class="report-box" readonly>NEGOCIO 2.0 - AORIX FOOD APP THEME

MEJORAS ACTIVAS:
- Login moderno estilo app de comida.
- Panel de administración ordenado, sin espacios vacíos.
- Vista celular con modo oscuro, botones grandes y navegación simple.
- Roles: ADMIN completo; usuarios operativos limitados.
- Diseño preparado para restaurante, pizzería y cadena con sucursales.

Base actual: {DB_PATH}</textarea></div>
    """
    return page(html, "admin")

@app.route("/logs")
@login_required
@admin_required
def logs():
    rows = q_all("SELECT * FROM logs ORDER BY id DESC LIMIT 300")
    trs = "".join(f'<tr><td>{r["fecha"]}</td><td>{r["hora"]}</td><td>{r["usuario"]}</td><td>{r["accion"]}</td><td>{r["detalle"]}</td></tr>' for r in rows) or '<tr><td colspan="5">Sin log.</td></tr>'
    return page(f'<div class="panel"><div class="box-title">Log del sistema</div><br><div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Hora</th><th>Usuario</th><th>Acción</th><th>Detalle</th></tr></thead><tbody>{trs}</tbody></table></div></div>', "log")

@app.errorhandler(500)
def err500(e):
    try:
        log_event("ERROR 500", str(e))
    except Exception:
        pass
    return page('<div class="panel"><h2>Error interno controlado</h2><p>No se perdió información. Revisa los logs de Render o vuelve al panel.</p><a class="btn btn-primary" href="/dashboard">Volver al panel</a></div>'), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=False)
