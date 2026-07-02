import os
from datetime import datetime, date
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, Response
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import csv, io
from openpyxl import load_workbook
from pdf_generator import generate_pi_pdf

basedir = os.path.abspath(os.path.dirname(__file__))
os.makedirs(os.path.join(basedir, 'instance'), exist_ok=True)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'change-me-in-production')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get(
    'DATABASE_URL', f"sqlite:///{os.path.join(basedir, 'instance', 'app.db')}"
).replace("postgres://", "postgresql://", 1)
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = "Please log in."

GST_RATE = 0.18
DISPATCH_UNITS = ["FORTUNAPET", "UNIT 1", "DABASPETE", "MAHIMAPURA", "DADRA"]
UOM_OPTIONS    = ["Nos", "KG", "Pieces", "BAG", "TON", "BOX"]
DOC_STATUSES   = ["pending", "delivered", "lost"]

# ── MODELS ──────────────────────────────────────────────────────────────────

class User(UserMixin, db.Model):
    id             = db.Column(db.Integer, primary_key=True)
    username       = db.Column(db.String(80), unique=True, nullable=False)
    full_name      = db.Column(db.String(120), nullable=False)
    password_hash  = db.Column(db.String(255), nullable=False)
    role           = db.Column(db.String(20), nullable=False, default='staff')
    active         = db.Column(db.Boolean, default=True)
    monthly_target = db.Column(db.Float, default=0)
    created_at     = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, pw):   self.password_hash = generate_password_hash(pw)
    def check_password(self, pw): return check_password_hash(self.password_hash, pw)
    def is_admin(self):           return self.role == 'admin'


class Customer(db.Model):
    id             = db.Column(db.Integer, primary_key=True)
    name           = db.Column(db.String(200), nullable=False)
    contact_person = db.Column(db.String(120))
    phone          = db.Column(db.String(30))
    created_by_id  = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    created_at     = db.Column(db.DateTime, default=datetime.utcnow)
    created_by     = db.relationship('User', backref='customers')


class Item(db.Model):
    id            = db.Column(db.Integer, primary_key=True)
    name          = db.Column(db.String(200), nullable=False)
    default_uom   = db.Column(db.String(20))
    default_price = db.Column(db.Float, default=0)
    created_by_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    created_at    = db.Column(db.DateTime, default=datetime.utcnow)
    created_by    = db.relationship('User', backref='items')


class Document(db.Model):
    id              = db.Column(db.Integer, primary_key=True)
    doc_type        = db.Column(db.String(10), default='PI')
    quote_no        = db.Column(db.String(20), unique=True, nullable=False)
    doc_date        = db.Column(db.Date, default=date.today)
    dispatch_from   = db.Column(db.String(50))
    customer_id     = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    item_desc       = db.Column(db.String(300), nullable=False)
    packaging       = db.Column(db.String(50))
    qty             = db.Column(db.Float, default=0)
    uom             = db.Column(db.String(20))
    price           = db.Column(db.Float, default=0)
    base_amount     = db.Column(db.Float, default=0)
    gst_applied     = db.Column(db.Boolean, default=False)
    gst_amount      = db.Column(db.Float, default=0)
    freight_charges = db.Column(db.Float, default=0)
    total_amount    = db.Column(db.Float, default=0)
    follow_up_date  = db.Column(db.Date)
    status          = db.Column(db.String(20), default='pending')
    lost_reason     = db.Column(db.Text)
    notes           = db.Column(db.Text)
    created_by_id   = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    created_at      = db.Column(db.DateTime, default=datetime.utcnow)
    customer        = db.relationship('Customer', backref='documents')
    created_by      = db.relationship('User', backref='documents')

    def recalc(self):
        self.base_amount  = round((self.qty or 0) * (self.price or 0), 2)
        self.gst_amount   = round(self.base_amount * GST_RATE, 2) if self.gst_applied else 0
        self.total_amount = round(self.base_amount + self.gst_amount + (self.freight_charges or 0), 2)


class Counter(db.Model):
    id    = db.Column(db.Integer, primary_key=True)
    name  = db.Column(db.String(50), unique=True, nullable=False)
    value = db.Column(db.Integer, nullable=False)


def next_quote_number():
    c = Counter.query.filter_by(name='quote_no').first()
    if not c:
        c = Counter(name='quote_no', value=260001)
        db.session.add(c)
    else:
        c.value += 1
    db.session.commit()
    return str(c.value)


@login_manager.user_loader
def load_user(uid): return User.query.get(int(uid))


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin():
            flash("Admin access required.", "danger")
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return wrapper


def visible_docs():
    if current_user.is_admin(): return Document.query
    return Document.query.filter_by(created_by_id=current_user.id)

def visible_custs():
    if current_user.is_admin(): return Customer.query
    return Customer.query.filter_by(created_by_id=current_user.id)

def _parse_date(v):
    if not v: return None
    try: return datetime.strptime(v, '%Y-%m-%d').date()
    except: return None

# ── AUTH ────────────────────────────────────────────────────────────────────

@app.route('/login', methods=['GET','POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('dashboard'))
    if request.method == 'POST':
        u = User.query.filter_by(username=request.form.get('username','').strip()).first()
        if u and u.active and u.check_password(request.form.get('password','')):
            login_user(u); return redirect(url_for('dashboard'))
        flash("Invalid username or password.", "danger")
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user(); return redirect(url_for('login'))

# ── DASHBOARD ───────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def dashboard():
    docs      = visible_docs().all()
    pending   = [d for d in docs if d.status == 'pending']
    delivered = [d for d in docs if d.status == 'delivered']
    lost      = [d for d in docs if d.status == 'lost']
    today     = date.today()
    overdue   = [d for d in pending if d.follow_up_date and d.follow_up_date < today]
    due_soon  = [d for d in pending if d.follow_up_date and 0 <= (d.follow_up_date-today).days <= 2]
    return render_template('dashboard.html',
        total_amount    = sum(d.total_amount or 0 for d in docs),
        pending_amount  = sum(d.total_amount or 0 for d in pending),
        delivered_count = len(delivered),
        pending_count   = len(pending),
        lost_count      = len(lost),
        overdue=overdue, due_soon=due_soon,
        pending_docs=sorted(pending, key=lambda d: d.follow_up_date or date.max)[:10],
        today=today)

# ── DOCUMENTS ───────────────────────────────────────────────────────────────

@app.route('/documents')
@login_required
def documents():
    q      = request.args.get('q','').strip()
    sf     = request.args.get('status','')
    query  = visible_docs()
    if sf: query = query.filter(Document.status == sf)
    docs   = query.order_by(Document.created_at.desc()).all()
    if q:
        ql   = q.lower()
        docs = [d for d in docs if ql in d.item_desc.lower()
                or ql in (d.customer.name.lower() if d.customer else '')
                or ql in d.quote_no.lower()]
    custs = visible_custs().order_by(Customer.name).all()
    items = Item.query.order_by(Item.name).all()
    return render_template('documents.html',
        docs=docs, customers=custs, items=items,
        dispatch_units=DISPATCH_UNITS, uom_options=UOM_OPTIONS,
        doc_statuses=DOC_STATUSES,
        q=q, status_filter=sf, today=date.today().isoformat())


@app.route('/documents/new', methods=['POST'])
@login_required
def new_document():
    cid   = request.form.get('customer_id')
    item  = request.form.get('item_desc','').strip()
    disp  = request.form.get('dispatch_from')
    if not cid or not item or not disp:
        flash("Customer, dispatch location and item are required.", "danger")
        return redirect(url_for('documents'))
    doc = Document(
        doc_type        = request.form.get('doc_type','PI'),
        quote_no        = next_quote_number(),
        doc_date        = _parse_date(request.form.get('doc_date')) or date.today(),
        dispatch_from   = disp,
        customer_id     = int(cid),
        item_desc       = item,
        packaging       = request.form.get('packaging','').strip(),
        qty             = float(request.form.get('qty') or 0),
        uom             = request.form.get('uom'),
        price           = float(request.form.get('price') or 0),
        gst_applied     = request.form.get('gst_applied') == 'on',
        freight_charges = float(request.form.get('freight_charges') or 0),
        follow_up_date  = _parse_date(request.form.get('follow_up_date')),
        status          = request.form.get('status','pending'),
        lost_reason     = request.form.get('lost_reason','').strip(),
        notes           = request.form.get('notes','').strip(),
        created_by_id   = current_user.id,
    )
    doc.recalc()
    db.session.add(doc); db.session.commit()
    flash(f"Document {doc.quote_no} created.", "success")
    return redirect(url_for('documents'))


@app.route('/documents/<int:doc_id>/edit', methods=['GET','POST'])
@login_required
def edit_document(doc_id):
    doc = Document.query.get_or_404(doc_id)
    if not current_user.is_admin() and doc.created_by_id != current_user.id:
        flash("You can only edit your own documents.", "danger")
        return redirect(url_for('documents'))
    custs = visible_custs().order_by(Customer.name).all()
    items = Item.query.order_by(Item.name).all()
    if request.method == 'POST':
        doc.doc_type        = request.form.get('doc_type', doc.doc_type)
        doc.doc_date        = _parse_date(request.form.get('doc_date')) or doc.doc_date
        doc.dispatch_from   = request.form.get('dispatch_from', doc.dispatch_from)
        doc.customer_id     = int(request.form.get('customer_id', doc.customer_id))
        doc.item_desc       = request.form.get('item_desc', doc.item_desc).strip()
        doc.packaging       = request.form.get('packaging','').strip()
        doc.qty             = float(request.form.get('qty') or 0)
        doc.uom             = request.form.get('uom', doc.uom)
        doc.price           = float(request.form.get('price') or 0)
        doc.gst_applied     = request.form.get('gst_applied') == 'on'
        doc.freight_charges = float(request.form.get('freight_charges') or 0)
        doc.follow_up_date  = _parse_date(request.form.get('follow_up_date'))
        doc.status          = request.form.get('status', doc.status)
        doc.lost_reason     = request.form.get('lost_reason','').strip()
        doc.notes           = request.form.get('notes','').strip()
        doc.recalc()
        db.session.commit()
        flash(f"Document {doc.quote_no} updated.", "success")
        return redirect(url_for('documents'))
    return render_template('edit_document.html', doc=doc,
        customers=custs, items=items,
        dispatch_units=DISPATCH_UNITS, uom_options=UOM_OPTIONS, doc_statuses=DOC_STATUSES)


@app.route('/documents/<int:doc_id>/delete', methods=['POST'])
@login_required
def delete_document(doc_id):
    doc = Document.query.get_or_404(doc_id)
    if not current_user.is_admin() and doc.created_by_id != current_user.id:
        flash("You can only delete your own documents.", "danger")
        return redirect(url_for('documents'))
    db.session.delete(doc); db.session.commit()
    flash("Document deleted.", "success")
    return redirect(url_for('documents'))


@app.route('/documents/<int:doc_id>/pdf')
@login_required
def download_document_pdf(doc_id):
    doc = Document.query.get_or_404(doc_id)
    if not current_user.is_admin() and doc.created_by_id != current_user.id:
        flash("Access denied.", "danger")
        return redirect(url_for('documents'))
    buf      = generate_pi_pdf(doc)
    filename = f"{doc.doc_type}_{doc.quote_no}_{doc.customer.name[:20].replace(' ','_')}.pdf"
    return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=filename)

# ── CUSTOMERS ───────────────────────────────────────────────────────────────

@app.route('/customers')
@login_required
def customers():
    return render_template('customers.html', customers=visible_custs().order_by(Customer.name).all())


@app.route('/customers/new', methods=['POST'])
@login_required
def new_customer():
    name = request.form.get('name','').strip()
    if not name:
        flash("Customer name is required.", "danger")
        return redirect(url_for('customers'))
    db.session.add(Customer(
        name           = name,
        contact_person = request.form.get('contact_person','').strip(),
        phone          = request.form.get('phone','').strip(),
        created_by_id  = current_user.id))
    db.session.commit()
    flash(f"Customer '{name}' added.", "success")
    return redirect(url_for('customers'))


@app.route('/customers/<int:cid>/delete', methods=['POST'])
@login_required
def delete_customer(cid):
    c = Customer.query.get_or_404(cid)
    if not current_user.is_admin() and c.created_by_id != current_user.id:
        flash("Access denied.", "danger")
        return redirect(url_for('customers'))
    db.session.delete(c); db.session.commit()
    flash("Customer deleted.", "success")
    return redirect(url_for('customers'))


@app.route('/customers/bulk-import', methods=['POST'])
@login_required
def bulk_import_customers():
    file = request.files.get('file')
    if not file or not file.filename:
        flash("Please choose a file.", "danger"); return redirect(url_for('customers'))
    rows = _read_tabular_file(file)
    if rows is None:
        flash("Use .csv or .xlsx", "danger"); return redirect(url_for('customers'))
    n = 0
    for row in rows:
        name = (row.get('name') or row.get('Name') or row.get('Customer Name') or '').strip()
        if not name: continue
        db.session.add(Customer(
            name=name,
            contact_person=(row.get('contact_person') or row.get('Contact Person') or '').strip(),
            phone=(row.get('phone') or row.get('Phone') or '').strip(),
            created_by_id=current_user.id)); n += 1
    db.session.commit()
    flash(f"Imported {n} customers.", "success")
    return redirect(url_for('customers'))


@app.route('/customers/template')
@login_required
def customer_template():
    return _csv_response(['name','contact_person','phone'],
        [['Example Traders Pvt Ltd','Ramesh Kumar','+919876543210']],
        'customer_import_template.csv')

# ── ITEMS ────────────────────────────────────────────────────────────────────

@app.route('/items')
@login_required
def items():
    return render_template('items.html', items=Item.query.order_by(Item.name).all(), uom_options=UOM_OPTIONS)


@app.route('/items/new', methods=['POST'])
@login_required
def new_item():
    name = request.form.get('name','').strip()
    if not name:
        flash("Item name required.", "danger"); return redirect(url_for('items'))
    db.session.add(Item(name=name,
        default_uom=request.form.get('default_uom'),
        default_price=float(request.form.get('default_price') or 0),
        created_by_id=current_user.id))
    db.session.commit()
    flash(f"Item '{name}' added.", "success")
    return redirect(url_for('items'))


@app.route('/items/<int:iid>/delete', methods=['POST'])
@login_required
def delete_item(iid):
    db.session.delete(Item.query.get_or_404(iid)); db.session.commit()
    flash("Item deleted.", "success"); return redirect(url_for('items'))


@app.route('/items/bulk-import', methods=['POST'])
@login_required
def bulk_import_items():
    file = request.files.get('file')
    if not file or not file.filename:
        flash("Choose a file.", "danger"); return redirect(url_for('items'))
    rows = _read_tabular_file(file)
    if rows is None:
        flash("Use .csv or .xlsx", "danger"); return redirect(url_for('items'))
    n = 0
    for row in rows:
        name = (row.get('name') or row.get('Name') or row.get('Item') or '').strip()
        if not name: continue
        db.session.add(Item(name=name,
            default_uom=(row.get('uom') or row.get('UOM') or '').strip(),
            default_price=float(row.get('price') or row.get('Price') or 0),
            created_by_id=current_user.id)); n += 1
    db.session.commit()
    flash(f"Imported {n} items.", "success")
    return redirect(url_for('items'))


@app.route('/items/template')
@login_required
def item_template():
    return _csv_response(['name','uom','price'],
        [['28 mm caps - K blue caps - 20 box','Nos','0.32']],
        'item_import_template.csv')

# ── ADMIN: USERS + REPORTS + TARGETS ────────────────────────────────────────

@app.route('/admin/users')
@login_required
@admin_required
def admin_users():
    return render_template('admin_users.html', users=User.query.order_by(User.username).all())


@app.route('/admin/users/new', methods=['POST'])
@login_required
@admin_required
def admin_new_user():
    username  = request.form.get('username','').strip().lower()
    full_name = request.form.get('full_name','').strip()
    password  = request.form.get('password','')
    role      = request.form.get('role','staff')
    target    = float(request.form.get('monthly_target') or 0)
    if not username or not password or not full_name:
        flash("Username, full name and password required.", "danger")
        return redirect(url_for('admin_users'))
    if User.query.filter_by(username=username).first():
        flash("Username already exists.", "danger")
        return redirect(url_for('admin_users'))
    u = User(username=username, full_name=full_name, role=role, monthly_target=target)
    u.set_password(password)
    db.session.add(u); db.session.commit()
    flash(f"User '{username}' created.", "success")
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/toggle', methods=['POST'])
@login_required
@admin_required
def admin_toggle_user(uid):
    u = User.query.get_or_404(uid)
    if u.id == current_user.id:
        flash("Cannot deactivate your own account.", "danger")
        return redirect(url_for('admin_users'))
    u.active = not u.active; db.session.commit()
    flash(f"User '{u.username}' {'activated' if u.active else 'deactivated'}.", "success")
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/reset-password', methods=['POST'])
@login_required
@admin_required
def admin_reset_password(uid):
    u  = User.query.get_or_404(uid)
    pw = request.form.get('new_password','')
    if len(pw) < 6:
        flash("Password must be at least 6 characters.", "danger")
        return redirect(url_for('admin_users'))
    u.set_password(pw); db.session.commit()
    flash(f"Password reset for '{u.username}'.", "success")
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/set-target', methods=['POST'])
@login_required
@admin_required
def admin_set_target(uid):
    u = User.query.get_or_404(uid)
    u.monthly_target = float(request.form.get('monthly_target') or 0)
    db.session.commit()
    flash(f"Target updated for '{u.full_name}'.", "success")
    return redirect(url_for('admin_reports'))


@app.route('/admin/reports')
@login_required
@admin_required
def admin_reports():
    staff_users  = User.query.filter_by(role='staff').order_by(User.full_name).all()
    selected_id  = request.args.get('user_id', type=int)
    selected_user = User.query.get(selected_id) if selected_id else None
    summaries = []
    for u in staff_users:
        udocs     = Document.query.filter_by(created_by_id=u.id).all()
        delivered = sum(d.total_amount or 0 for d in udocs if d.status == 'delivered')
        summaries.append(dict(
            user          = u,
            total_docs    = len(udocs),
            total_amount  = sum(d.total_amount or 0 for d in udocs),
            delivered_amount = delivered,
            pending_count = len([d for d in udocs if d.status == 'pending']),
            lost_count    = len([d for d in udocs if d.status == 'lost']),
            target        = u.monthly_target,
            target_pct    = round(delivered / u.monthly_target * 100, 1) if u.monthly_target else 0,
        ))
    detail_docs = Document.query.filter_by(created_by_id=selected_user.id)\
        .order_by(Document.created_at.desc()).all() if selected_user else []
    return render_template('admin_reports.html',
        summaries=summaries, staff_users=staff_users,
        selected_user=selected_user, detail_docs=detail_docs)

# ── HELPERS ──────────────────────────────────────────────────────────────────

def _read_tabular_file(fs):
    fn = fs.filename.lower()
    if fn.endswith('.csv'):
        return list(csv.DictReader(io.StringIO(fs.stream.read().decode('utf-8-sig'))))
    if fn.endswith('.xlsx'):
        wb = load_workbook(fs, read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        try: hdr = [str(h).strip() if h else '' for h in next(rows)]
        except StopIteration: return []
        return [{hdr[i]: ('' if i >= len(r) or r[i] is None else str(r[i])) for i in range(len(hdr))} for r in rows]
    return None


def _csv_response(headers, rows, filename):
    out = io.StringIO()
    csv.writer(out).writerow(headers)
    csv.writer(out).writerows(rows)
    return Response(out.getvalue(), mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename={filename}'})

# ── STARTUP ───────────────────────────────────────────────────────────────────

with app.app_context():
    db.create_all()
    with db.engine.connect() as conn:
        for stmt in [
            "ALTER TABLE document ADD COLUMN lost_reason TEXT",
            "ALTER TABLE user ADD COLUMN monthly_target FLOAT DEFAULT 0",
        ]:
            try: conn.execute(db.text(stmt)); conn.commit()
            except: pass
    if not User.query.filter_by(username='admin').first():
        a = User(username='admin', full_name='Administrator', role='admin')
        a.set_password('admin123'); db.session.add(a); db.session.commit()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',5000)), debug=False)
