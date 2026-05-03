import os, sqlite3, json, secrets
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, g, flash, abort
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get('DATABASE_PATH', os.path.join(BASE_DIR, 'instance', 'kyte_checkout.db'))
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'render-demo-' + secrets.token_hex(16))
app.config.update(JSON_AS_ASCII=False)


def db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(exc=None):
    con = g.pop('db', None)
    if con is not None:
        con.close()

def q(sql, args=(), one=False):
    cur = db().execute(sql, args)
    rows = cur.fetchall()
    return (rows[0] if rows else None) if one else rows

def execq(sql, args=()):
    con = db(); cur = con.execute(sql, args); con.commit(); return cur

def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.executescript('''
    CREATE TABLE IF NOT EXISTS users(
      id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'vendedor', active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS customers(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT DEFAULT '', email TEXT DEFAULT '', created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS products(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, category TEXT NOT NULL, price REAL NOT NULL,
      stock INTEGER NOT NULL DEFAULT 0, image TEXT DEFAULT '', active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS sales(
      id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE NOT NULL, customer_id INTEGER, user_id INTEGER,
      subtotal REAL NOT NULL, discount REAL NOT NULL, total REAL NOT NULL, payment TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'pagado', created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS sale_items(
      id INTEGER PRIMARY KEY AUTOINCREMENT, sale_id INTEGER NOT NULL, product_id INTEGER NOT NULL,
      name TEXT NOT NULL, qty INTEGER NOT NULL, price REAL NOT NULL, total REAL NOT NULL);
    CREATE TABLE IF NOT EXISTS day_closings(
      id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, date TEXT NOT NULL, cash_total REAL NOT NULL,
      card_total REAL NOT NULL, transfer_total REAL NOT NULL, total REAL NOT NULL, notes TEXT DEFAULT '', created_at TEXT NOT NULL);
    ''')
    now = datetime.now().isoformat(timespec='seconds')
    users = cur.execute('SELECT COUNT(*) FROM users').fetchone()[0]
    if users == 0:
        cur.executemany('INSERT INTO users(username,password_hash,role,active,created_at) VALUES(?,?,?,?,?)', [
            ('admin1', generate_password_hash('admin123'), 'admin', 1, now),
            ('vendedor1', generate_password_hash('venta123'), 'vendedor', 1, now),
        ])
    cust = cur.execute('SELECT COUNT(*) FROM customers').fetchone()[0]
    if cust == 0:
        cur.executemany('INSERT INTO customers(name,phone,email,created_at) VALUES(?,?,?,?)', [
            ('Cliente general','','',now), ('Cliente WhatsApp','','',now), ('Mesa rápida','','',now)
        ])
    prod = cur.execute('SELECT COUNT(*) FROM products').fetchone()[0]
    if prod == 0:
        products = [
            ('Menú Ejecutivo','Comidas',12,200,'🍽️'), ('Combo Almuerzo','Combos',15,150,'🥘'), ('Cena','Comidas',11,150,'🍛'),
            ('Bebida Personal','Bebidas',3.5,300,'🥤'), ('Postre','Postres',4,100,'🍮'), ('Entrada','Comidas',5,120,'🥗'),
            ('Promo Familiar','Promos',35,60,'🛍️'), ('Café','Bebidas',4.5,90,'☕'), ('Agua','Bebidas',2.5,280,'💧')]
        cur.executemany('INSERT INTO products(name,category,price,stock,image,active,created_at) VALUES(?,?,?,?,?,?,?)',
                        [(n,c,p,s,img,1,now) for n,c,p,s,img in products])
    con.commit(); con.close()

with app.app_context():
    init_db()

def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper

def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if session.get('role') != 'admin':
            flash('Solo administrador.', 'error'); return redirect(url_for('checkout'))
        return f(*args, **kwargs)
    return wrapper

@app.errorhandler(Exception)
def handle_exception(e):
    code = getattr(e, 'code', 500)
    app.logger.exception('ERROR CONTROLADO')
    return render_template('error.html', error=str(e), code=code), code

@app.route('/healthz', methods=['GET','HEAD'])
def healthz(): return 'ok', 200

@app.route('/', methods=['GET','HEAD'])
def index():
    return redirect(url_for('checkout') if session.get('user_id') else url_for('login'))

@app.route('/login', methods=['GET','POST','HEAD'])
def login():
    if request.method == 'POST':
        user = q('SELECT * FROM users WHERE username=? AND active=1', (request.form.get('username','').strip(),), True)
        if user and check_password_hash(user['password_hash'], request.form.get('password','')):
            session.clear(); session['user_id']=user['id']; session['username']=user['username']; session['role']=user['role']
            return redirect(url_for('checkout'))
        flash('Usuario o contraseña incorrectos.', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/checkout')
@login_required
def checkout():
    return render_template('checkout.html')

@app.route('/api/bootstrap')
@login_required
def api_bootstrap():
    products = [dict(r) for r in q('SELECT * FROM products WHERE active=1 ORDER BY category,name')]
    customers = [dict(r) for r in q('SELECT * FROM customers ORDER BY id')]
    cats = sorted({p['category'] for p in products})
    return jsonify(products=products, customers=customers, categories=cats, user={'username':session['username'],'role':session['role']})

@app.route('/api/checkout', methods=['POST'])
@login_required
def api_checkout():
    data = request.get_json(force=True)
    items = data.get('items') or []
    if not items: return jsonify(ok=False, error='Carrito vacío'), 400
    discount = max(float(data.get('discount') or 0), 0)
    payment = data.get('payment') or 'Efectivo'
    customer_id = int(data.get('customer_id') or 1)
    subtotal = 0.0
    checked = []
    for item in items:
        pid = int(item['id']); qty = int(item.get('qty') or 1)
        prod = q('SELECT * FROM products WHERE id=? AND active=1', (pid,), True)
        if not prod: return jsonify(ok=False, error='Producto no existe'), 400
        if prod['stock'] < qty: return jsonify(ok=False, error=f'Stock insuficiente: {prod["name"]}'), 400
        line = round(prod['price'] * qty, 2); subtotal += line
        checked.append((prod, qty, line))
    total = max(round(subtotal - discount, 2), 0)
    code = 'V' + datetime.now().strftime('%Y%m%d%H%M%S')
    cur = execq('INSERT INTO sales(code,customer_id,user_id,subtotal,discount,total,payment,status,created_at) VALUES(?,?,?,?,?,?,?,?,?)',
          (code, customer_id, session['user_id'], subtotal, discount, total, payment, 'pagado', datetime.now().isoformat(timespec='seconds')))
    sale_id = cur.lastrowid
    for prod, qty, line in checked:
        execq('INSERT INTO sale_items(sale_id,product_id,name,qty,price,total) VALUES(?,?,?,?,?,?)',
              (sale_id, prod['id'], prod['name'], qty, prod['price'], line))
        execq('UPDATE products SET stock=stock-? WHERE id=?', (qty, prod['id']))
    return jsonify(ok=True, sale_id=sale_id, code=code, receipt=url_for('receipt', sale_id=sale_id))

@app.route('/receipt/<int:sale_id>')
@login_required
def receipt(sale_id):
    sale = q('SELECT s.*, c.name customer, u.username user FROM sales s LEFT JOIN customers c ON c.id=s.customer_id LEFT JOIN users u ON u.id=s.user_id WHERE s.id=?', (sale_id,), True)
    if not sale: abort(404)
    items = q('SELECT * FROM sale_items WHERE sale_id=?', (sale_id,))
    return render_template('receipt.html', sale=sale, items=items)

@app.route('/sales')
@login_required
def sales():
    rows = q('SELECT s.*, c.name customer, u.username user FROM sales s LEFT JOIN customers c ON c.id=s.customer_id LEFT JOIN users u ON u.id=s.user_id ORDER BY s.id DESC LIMIT 200')
    return render_template('sales.html', rows=rows)

@app.route('/products', methods=['GET','POST'])
@login_required
@admin_required
def products():
    if request.method == 'POST':
        execq('INSERT INTO products(name,category,price,stock,image,active,created_at) VALUES(?,?,?,?,?,?,?)',
              (request.form['name'], request.form['category'], float(request.form['price']), int(request.form['stock']), request.form.get('image','📦'), 1, datetime.now().isoformat(timespec='seconds')))
        return redirect(url_for('products'))
    rows = q('SELECT * FROM products ORDER BY category,name')
    return render_template('products.html', rows=rows)

@app.route('/customers', methods=['GET','POST'])
@login_required
def customers():
    if request.method == 'POST':
        execq('INSERT INTO customers(name,phone,email,created_at) VALUES(?,?,?,?)',
              (request.form['name'], request.form.get('phone',''), request.form.get('email',''), datetime.now().isoformat(timespec='seconds')))
        return redirect(url_for('customers'))
    rows = q('SELECT * FROM customers ORDER BY id DESC')
    return render_template('customers.html', rows=rows)

@app.route('/users', methods=['GET','POST'])
@login_required
@admin_required
def users():
    if request.method == 'POST':
        execq('INSERT INTO users(username,password_hash,role,active,created_at) VALUES(?,?,?,?,?)',
              (request.form['username'], generate_password_hash(request.form['password']), request.form['role'], 1, datetime.now().isoformat(timespec='seconds')))
        return redirect(url_for('users'))
    rows = q('SELECT id,username,role,active,created_at FROM users ORDER BY id DESC')
    return render_template('users.html', rows=rows)

@app.route('/close-day', methods=['GET','POST'])
@login_required
def close_day():
    today = datetime.now().strftime('%Y-%m-%d')
    sums = {r['payment']: r['total'] for r in q("SELECT payment, SUM(total) total FROM sales WHERE substr(created_at,1,10)=? GROUP BY payment", (today,))}
    total = sum(v or 0 for v in sums.values())
    if request.method == 'POST':
        execq('INSERT INTO day_closings(user_id,date,cash_total,card_total,transfer_total,total,notes,created_at) VALUES(?,?,?,?,?,?,?,?)',
              (session['user_id'], today, sums.get('Efectivo',0) or 0, sums.get('Tarjeta',0) or 0, sums.get('Transferencia',0) or 0, total, request.form.get('notes',''), datetime.now().isoformat(timespec='seconds')))
        flash('Día cerrado correctamente.', 'ok')
    closes = q('SELECT d.*, u.username FROM day_closings d LEFT JOIN users u ON u.id=d.user_id ORDER BY d.id DESC LIMIT 50')
    return render_template('close_day.html', sums=sums, total=total, closes=closes, today=today)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
