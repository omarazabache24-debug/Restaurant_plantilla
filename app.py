
import os
import sqlite3
import logging
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB = os.path.join(BASE_DIR, "database", "pos.db")
DB_PATH = os.environ.get("DATABASE_PATH", DEFAULT_DB)
SECRET_KEY = os.environ.get("SECRET_KEY", "aorix-pos-render-cambiar")

os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = SECRET_KEY
app.config["JSON_AS_ASCII"] = False

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s in %(module)s: %(message)s")


def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('admin','seller')) DEFAULT 'seller',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS products(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT UNIQUE,
        name TEXT NOT NULL,
        category TEXT DEFAULT 'General',
        price REAL NOT NULL DEFAULT 0,
        cost REAL NOT NULL DEFAULT 0,
        stock INTEGER NOT NULL DEFAULT 0,
        min_stock INTEGER NOT NULL DEFAULT 5,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        email TEXT,
        notes TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS sales(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_code TEXT UNIQUE NOT NULL,
        customer_id INTEGER,
        user_id INTEGER,
        subtotal REAL NOT NULL,
        discount REAL NOT NULL DEFAULT 0,
        total REAL NOT NULL,
        payment_method TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'PAGADO',
        created_at TEXT NOT NULL,
        FOREIGN KEY(customer_id) REFERENCES customers(id),
        FOREIGN KEY(user_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS sale_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_id INTEGER NOT NULL,
        product_id INTEGER NOT NULL,
        qty INTEGER NOT NULL,
        unit_price REAL NOT NULL,
        total REAL NOT NULL,
        FOREIGN KEY(sale_id) REFERENCES sales(id) ON DELETE CASCADE,
        FOREIGN KEY(product_id) REFERENCES products(id)
    );
    CREATE TABLE IF NOT EXISTS day_closings(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        closed_by INTEGER NOT NULL,
        total_sales REAL NOT NULL,
        total_orders INTEGER NOT NULL,
        cash_total REAL NOT NULL,
        card_total REAL NOT NULL,
        transfer_total REAL NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY(closed_by) REFERENCES users(id)
    );
    """)
    if cur.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"] == 0:
        cur.executemany("INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)", [
            ("admin1", generate_password_hash("admin123"), "admin", now_iso()),
            ("admin2", generate_password_hash("admin123"), "admin", now_iso()),
            ("vendedor1", generate_password_hash("venta123"), "seller", now_iso()),
        ])
    if cur.execute("SELECT COUNT(*) AS n FROM products").fetchone()["n"] == 0:
        products = [
            ("P001", "Menú Ejecutivo", "Comedor", 12.00, 7.50, 200, 20),
            ("P002", "Combo Almuerzo", "Promos", 15.00, 9.00, 150, 15),
            ("P003", "Cena", "Comedor", 11.00, 6.80, 150, 15),
            ("P004", "Bebida Personal", "Bebidas", 3.50, 1.80, 300, 30),
            ("P005", "Postre", "Comedor", 4.00, 2.00, 100, 10),
        ]
        cur.executemany("INSERT INTO products(sku,name,category,price,cost,stock,min_stock) VALUES(?,?,?,?,?,?,?)", products)
    if cur.execute("SELECT COUNT(*) AS n FROM customers").fetchone()["n"] == 0:
        cur.execute("INSERT INTO customers(name,phone,email,notes) VALUES(?,?,?,?)", ("Cliente General", "", "", "Venta rápida"))
    conn.commit(); conn.close()

init_db()


def get_user():
    uid = session.get("user_id")
    if not uid:
        return None
    conn = get_conn()
    user = conn.execute("SELECT id,username,role,active,created_at FROM users WHERE id=? AND active=1", (uid,)).fetchone()
    conn.close()
    return user


@app.context_processor
def inject_globals():
    return {"me": get_user(), "app_name": "AORIX POS Pro"}


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Solo administrador puede realizar esta acción.", "error")
            return redirect(url_for("checkout"))
        return fn(*args, **kwargs)
    return wrapper


@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.exception("Error interno: %s", e)
    return render_template("error.html", error=str(e)), 500


@app.route("/healthz", methods=["GET", "HEAD"])
def healthz():
    return "OK", 200


@app.route("/", methods=["GET", "HEAD", "POST"])
def login():
    if request.method == "HEAD":
        return "", 200
    if session.get("user_id") and request.method == "GET":
        return redirect(url_for("checkout"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_conn()
        user = conn.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        conn.close()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("checkout"))
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/checkout")
@login_required
def checkout():
    conn = get_conn()
    products = conn.execute("SELECT * FROM products WHERE active=1 ORDER BY category,name").fetchall()
    customers = conn.execute("SELECT * FROM customers ORDER BY name").fetchall()
    conn.close()
    return render_template("checkout.html", products=products, customers=customers)


@app.route("/api/products")
@login_required
def api_products():
    q = "%" + request.args.get("q", "").strip() + "%"
    conn = get_conn()
    rows = conn.execute("""
        SELECT * FROM products
        WHERE active=1 AND (name LIKE ? OR sku LIKE ? OR category LIKE ?)
        ORDER BY name
    """, (q, q, q)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/checkout", methods=["POST"])
@login_required
def api_checkout():
    data = request.get_json(silent=True) or {}
    items = data.get("items") or []
    if not items:
        return jsonify({"ok": False, "error": "Carrito vacío"}), 400
    try:
        discount = max(float(data.get("discount") or 0), 0)
    except ValueError:
        discount = 0
    payment = data.get("payment_method") or "Efectivo"
    customer_id = data.get("customer_id") or None
    conn = get_conn(); cur = conn.cursor()
    try:
        subtotal = 0.0
        checked = []
        for item in items:
            pid = int(item.get("product_id"))
            qty = int(item.get("qty", 0))
            if qty <= 0:
                raise ValueError("Cantidad inválida")
            product = cur.execute("SELECT * FROM products WHERE id=? AND active=1", (pid,)).fetchone()
            if not product:
                raise ValueError("Producto no encontrado")
            if int(product["stock"]) < qty:
                raise ValueError(f"Stock insuficiente: {product['name']}")
            line_total = qty * float(product["price"])
            subtotal += line_total
            checked.append((product, qty, line_total))
        total = max(subtotal - discount, 0)
        code = "V" + datetime.now().strftime("%Y%m%d%H%M%S%f")
        cur.execute("""
            INSERT INTO sales(sale_code,customer_id,user_id,subtotal,discount,total,payment_method,created_at)
            VALUES(?,?,?,?,?,?,?,?)
        """, (code, customer_id, session["user_id"], subtotal, discount, total, payment, now_iso()))
        sale_id = cur.lastrowid
        for product, qty, line_total in checked:
            cur.execute("INSERT INTO sale_items(sale_id,product_id,qty,unit_price,total) VALUES(?,?,?,?,?)",
                        (sale_id, product["id"], qty, product["price"], line_total))
            cur.execute("UPDATE products SET stock=stock-? WHERE id=?", (qty, product["id"]))
        conn.commit()
        return jsonify({"ok": True, "sale_id": sale_id, "sale_code": code, "total": total, "receipt_url": url_for("receipt", sale_id=sale_id)})
    except Exception as e:
        conn.rollback()
        return jsonify({"ok": False, "error": str(e)}), 400
    finally:
        conn.close()


@app.route("/receipt/<int:sale_id>")
@login_required
def receipt(sale_id):
    conn = get_conn()
    sale = conn.execute("""
        SELECT s.*, c.name AS customer, u.username AS seller
        FROM sales s
        LEFT JOIN customers c ON c.id=s.customer_id
        LEFT JOIN users u ON u.id=s.user_id
        WHERE s.id=?
    """, (sale_id,)).fetchone()
    if not sale:
        conn.close(); flash("Venta no encontrada.", "error"); return redirect(url_for("sales"))
    items = conn.execute("""
        SELECT si.*, p.name, p.sku
        FROM sale_items si JOIN products p ON p.id=si.product_id
        WHERE si.sale_id=?
    """, (sale_id,)).fetchall()
    conn.close()
    return render_template("receipt.html", sale=sale, items=items)


@app.route("/products", methods=["GET", "POST"])
@login_required
@admin_required
def products():
    conn = get_conn()
    if request.method == "POST":
        form = request.form
        try:
            conn.execute("""INSERT INTO products(sku,name,category,price,cost,stock,min_stock)
                          VALUES(?,?,?,?,?,?,?)""",
                         (form.get("sku") or None, form.get("name"), form.get("category") or "General",
                          float(form.get("price") or 0), float(form.get("cost") or 0),
                          int(form.get("stock") or 0), int(form.get("min_stock") or 0)))
            conn.commit(); flash("Producto creado correctamente.", "ok")
        except sqlite3.IntegrityError:
            flash("El SKU ya existe. Usa otro código.", "error")
    rows = conn.execute("SELECT * FROM products ORDER BY id DESC").fetchall()
    conn.close()
    return render_template("products.html", products=rows)


@app.route("/customers", methods=["GET", "POST"])
@login_required
def customers():
    conn = get_conn()
    if request.method == "POST":
        form = request.form
        conn.execute("INSERT INTO customers(name,phone,email,notes) VALUES(?,?,?,?)",
                     (form.get("name"), form.get("phone"), form.get("email"), form.get("notes")))
        conn.commit(); flash("Cliente creado correctamente.", "ok")
    rows = conn.execute("SELECT * FROM customers ORDER BY id DESC").fetchall()
    conn.close()
    return render_template("customers.html", customers=rows)


@app.route("/sales")
@login_required
def sales():
    conn = get_conn()
    rows = conn.execute("""
        SELECT s.*, c.name AS customer, u.username AS seller
        FROM sales s
        LEFT JOIN customers c ON c.id=s.customer_id
        LEFT JOIN users u ON u.id=s.user_id
        ORDER BY s.id DESC LIMIT 500
    """).fetchall()
    conn.close()
    return render_template("sales.html", sales=rows)


@app.route("/users", methods=["GET", "POST"])
@login_required
@admin_required
def users():
    conn = get_conn()
    if request.method == "POST":
        form = request.form
        try:
            conn.execute("INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
                         (form.get("username"), generate_password_hash(form.get("password") or "123456"), form.get("role") or "seller", now_iso()))
            conn.commit(); flash("Usuario creado correctamente.", "ok")
        except sqlite3.IntegrityError:
            flash("Ese usuario ya existe.", "error")
    rows = conn.execute("SELECT id,username,role,active,created_at FROM users ORDER BY id DESC").fetchall()
    conn.close()
    return render_template("users.html", users=rows)


@app.route("/users/<int:user_id>/toggle", methods=["POST"])
@login_required
@admin_required
def toggle_user(user_id):
    if user_id == session.get("user_id"):
        flash("No puedes desactivar tu propio usuario.", "error")
        return redirect(url_for("users"))
    conn = get_conn()
    conn.execute("UPDATE users SET active=CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?", (user_id,))
    conn.commit(); conn.close()
    flash("Estado de usuario actualizado.", "ok")
    return redirect(url_for("users"))


@app.route("/close-day", methods=["GET", "POST"])
@login_required
def close_day():
    conn = get_conn()
    stats = conn.execute("""
        SELECT COUNT(*) AS n,
               COALESCE(SUM(total),0) AS total,
               COALESCE(SUM(CASE WHEN payment_method='Efectivo' THEN total ELSE 0 END),0) AS cash,
               COALESCE(SUM(CASE WHEN payment_method='Tarjeta' THEN total ELSE 0 END),0) AS card,
               COALESCE(SUM(CASE WHEN payment_method='Transferencia' THEN total ELSE 0 END),0) AS transfer
        FROM sales
        WHERE date(created_at)=date('now','localtime')
    """).fetchone()
    if request.method == "POST":
        conn.execute("""INSERT INTO day_closings(closed_by,total_sales,total_orders,cash_total,card_total,transfer_total,created_at)
                      VALUES(?,?,?,?,?,?,?)""",
                     (session["user_id"], stats["total"], stats["n"], stats["cash"], stats["card"], stats["transfer"], now_iso()))
        conn.commit(); flash("Día cerrado correctamente.", "ok")
    closings = conn.execute("""
        SELECT dc.*, u.username FROM day_closings dc
        JOIN users u ON u.id=dc.closed_by
        ORDER BY dc.id DESC LIMIT 100
    """).fetchall()
    conn.close()
    return render_template("close_day.html", stats=stats, closings=closings)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
