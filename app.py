import os
import sqlite3
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
from werkzeug.security import generate_password_hash, check_password_hash

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, 'database', 'pos.db')
RECEIPTS_DIR = os.path.join(APP_DIR, 'receipts')
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
os.makedirs(RECEIPTS_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'cambia-esta-clave-en-render')


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()
    c = conn.cursor()
    c.executescript('''
    CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'seller',
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS products(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sku TEXT UNIQUE,
        name TEXT NOT NULL,
        category TEXT,
        price REAL NOT NULL DEFAULT 0,
        cost REAL NOT NULL DEFAULT 0,
        stock INTEGER NOT NULL DEFAULT 0,
        min_stock INTEGER NOT NULL DEFAULT 5,
        active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        email TEXT,
        notes TEXT
    );
    CREATE TABLE IF NOT EXISTS sales(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_code TEXT UNIQUE NOT NULL,
        customer_id INTEGER,
        user_id INTEGER,
        subtotal REAL NOT NULL,
        discount REAL NOT NULL,
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
        FOREIGN KEY(sale_id) REFERENCES sales(id),
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
    ''')
    if not c.execute('SELECT id FROM users LIMIT 1').fetchone():
        c.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)',
                  ('admin1', generate_password_hash('admin123'), 'admin', datetime.now().isoformat()))
        c.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)',
                  ('vendedor1', generate_password_hash('venta123'), 'seller', datetime.now().isoformat()))
    if not c.execute('SELECT id FROM products LIMIT 1').fetchone():
        demo = [
            ('P001','Menú Ejecutivo','Comedor',12.00,7.50,80,10),
            ('P002','Bebida Personal','Bebidas',3.50,1.80,120,20),
            ('P003','Postre','Comedor',4.00,2.00,60,10),
            ('P004','Combo Almuerzo','Promos',15.00,9.00,50,8),
            ('P005','Cena','Comedor',11.00,6.80,70,10),
        ]
        c.executemany('INSERT INTO products(sku,name,category,price,cost,stock,min_stock) VALUES(?,?,?,?,?,?,?)', demo)
    if not c.execute('SELECT id FROM customers LIMIT 1').fetchone():
        c.execute('INSERT INTO customers(name,phone,email,notes) VALUES(?,?,?,?)', ('Cliente General','','','Venta rápida'))
    conn.commit(); conn.close()


def current_user():
    if 'user_id' not in session: return None
    conn=db(); user=conn.execute('SELECT * FROM users WHERE id=?', (session['user_id'],)).fetchone(); conn.close(); return user


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get('role') != 'admin':
            flash('Solo administrador puede realizar esta acción.', 'error')
            return redirect(url_for('checkout'))
        return fn(*args, **kwargs)
    return wrapper

@app.context_processor
def inject_user():
    return {'me': current_user()}

@app.route('/', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        username=request.form.get('username','').strip()
        password=request.form.get('password','')
        conn=db(); user=conn.execute('SELECT * FROM users WHERE username=? AND active=1', (username,)).fetchone(); conn.close()
        if user and check_password_hash(user['password_hash'], password):
            session['user_id']=user['id']; session['username']=user['username']; session['role']=user['role']
            return redirect(url_for('checkout'))
        flash('Usuario o contraseña incorrectos.', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/checkout')
@login_required
def checkout():
    conn=db()
    products=conn.execute('SELECT * FROM products WHERE active=1 ORDER BY category,name').fetchall()
    customers=conn.execute('SELECT * FROM customers ORDER BY name').fetchall()
    conn.close()
    return render_template('checkout.html', products=products, customers=customers)

@app.route('/api/products')
@login_required
def api_products():
    q = '%' + request.args.get('q','').strip() + '%'
    conn=db(); rows=conn.execute('SELECT * FROM products WHERE active=1 AND (name LIKE ? OR sku LIKE ? OR category LIKE ?) ORDER BY name', (q,q,q)).fetchall(); conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/checkout', methods=['POST'])
@login_required
def api_checkout():
    data=request.get_json(force=True)
    items=data.get('items', [])
    if not items: return jsonify({'ok': False, 'error': 'Carrito vacío'}), 400
    discount=float(data.get('discount') or 0)
    payment=data.get('payment_method','Efectivo')
    customer_id=data.get('customer_id') or None
    conn=db(); c=conn.cursor()
    try:
        subtotal=0; checked=[]
        for it in items:
            pid=int(it['product_id']); qty=int(it['qty'])
            p=c.execute('SELECT * FROM products WHERE id=? AND active=1', (pid,)).fetchone()
            if not p: raise ValueError('Producto no encontrado')
            if qty <= 0: raise ValueError('Cantidad inválida')
            if p['stock'] < qty: raise ValueError(f"Stock insuficiente: {p['name']}")
            line=qty*float(p['price']); subtotal += line; checked.append((p,qty,line))
        total=max(subtotal-discount, 0)
        code='V' + datetime.now().strftime('%Y%m%d%H%M%S')
        c.execute('INSERT INTO sales(sale_code,customer_id,user_id,subtotal,discount,total,payment_method,created_at) VALUES(?,?,?,?,?,?,?,?)',
                  (code, customer_id, session['user_id'], subtotal, discount, total, payment, datetime.now().isoformat()))
        sale_id=c.lastrowid
        for p, qty, line in checked:
            c.execute('INSERT INTO sale_items(sale_id,product_id,qty,unit_price,total) VALUES(?,?,?,?,?)', (sale_id,p['id'],qty,p['price'],line))
            c.execute('UPDATE products SET stock=stock-? WHERE id=?', (qty,p['id']))
        conn.commit()
        return jsonify({'ok': True, 'sale_id': sale_id, 'sale_code': code, 'total': total, 'receipt_url': url_for('receipt', sale_id=sale_id)})
    except Exception as e:
        conn.rollback(); return jsonify({'ok': False, 'error': str(e)}), 400
    finally:
        conn.close()

@app.route('/receipt/<int:sale_id>')
@login_required
def receipt(sale_id):
    conn=db()
    sale=conn.execute('SELECT s.*, c.name customer, u.username seller FROM sales s LEFT JOIN customers c ON c.id=s.customer_id LEFT JOIN users u ON u.id=s.user_id WHERE s.id=?', (sale_id,)).fetchone()
    items=conn.execute('SELECT si.*, p.name, p.sku FROM sale_items si JOIN products p ON p.id=si.product_id WHERE si.sale_id=?', (sale_id,)).fetchall(); conn.close()
    return render_template('receipt.html', sale=sale, items=items)

@app.route('/products', methods=['GET','POST'])
@login_required
@admin_required
def products():
    conn=db()
    if request.method=='POST':
        f=request.form
        conn.execute('INSERT INTO products(sku,name,category,price,cost,stock,min_stock) VALUES(?,?,?,?,?,?,?)',
                     (f.get('sku'), f.get('name'), f.get('category'), float(f.get('price') or 0), float(f.get('cost') or 0), int(f.get('stock') or 0), int(f.get('min_stock') or 0)))
        conn.commit(); flash('Producto creado.', 'ok')
    rows=conn.execute('SELECT * FROM products ORDER BY id DESC').fetchall(); conn.close()
    return render_template('products.html', products=rows)

@app.route('/customers', methods=['GET','POST'])
@login_required
def customers():
    conn=db()
    if request.method=='POST':
        f=request.form; conn.execute('INSERT INTO customers(name,phone,email,notes) VALUES(?,?,?,?)', (f.get('name'), f.get('phone'), f.get('email'), f.get('notes'))); conn.commit(); flash('Cliente creado.', 'ok')
    rows=conn.execute('SELECT * FROM customers ORDER BY id DESC').fetchall(); conn.close()
    return render_template('customers.html', customers=rows)

@app.route('/sales')
@login_required
def sales():
    conn=db(); rows=conn.execute('SELECT s.*, c.name customer, u.username seller FROM sales s LEFT JOIN customers c ON c.id=s.customer_id LEFT JOIN users u ON u.id=s.user_id ORDER BY s.id DESC LIMIT 200').fetchall(); conn.close()
    return render_template('sales.html', sales=rows)

@app.route('/users', methods=['GET','POST'])
@login_required
@admin_required
def users():
    conn=db()
    if request.method=='POST':
        f=request.form
        conn.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)', (f.get('username'), generate_password_hash(f.get('password')), f.get('role','seller'), datetime.now().isoformat()))
        conn.commit(); flash('Usuario creado.', 'ok')
    rows=conn.execute('SELECT id,username,role,active,created_at FROM users ORDER BY id DESC').fetchall(); conn.close()
    return render_template('users.html', users=rows)

@app.route('/close-day', methods=['GET','POST'])
@login_required
def close_day():
    conn=db()
    today=datetime.now().date().isoformat()
    stats=conn.execute("SELECT COUNT(*) n, COALESCE(SUM(total),0) total, COALESCE(SUM(CASE WHEN payment_method='Efectivo' THEN total ELSE 0 END),0) cash, COALESCE(SUM(CASE WHEN payment_method='Tarjeta' THEN total ELSE 0 END),0) card, COALESCE(SUM(CASE WHEN payment_method='Transferencia' THEN total ELSE 0 END),0) transfer FROM sales WHERE date(created_at)=date('now','localtime')").fetchone()
    if request.method=='POST':
        conn.execute('INSERT INTO day_closings(closed_by,total_sales,total_orders,cash_total,card_total,transfer_total,created_at) VALUES(?,?,?,?,?,?,?)', (session['user_id'], stats['total'], stats['n'], stats['cash'], stats['card'], stats['transfer'], datetime.now().isoformat()))
        conn.commit(); flash('Día cerrado correctamente.', 'ok')
    closings=conn.execute('SELECT dc.*, u.username FROM day_closings dc JOIN users u ON u.id=dc.closed_by ORDER BY dc.id DESC LIMIT 50').fetchall(); conn.close()
    return render_template('close_day.html', stats=stats, closings=closings, today=today)

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
else:
    init_db()
