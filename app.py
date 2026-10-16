from __future__ import annotations

import sqlite3
import sys
import unicodedata
import os
import re
import secrets
import hmac
import csv
import io
import hashlib
import threading
from decimal import Decimal, InvalidOperation
from functools import wraps
from io import BytesIO
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import barcode
from barcode.codex import Code128
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.graphics.barcode import code128
from reportlab.graphics.barcode.qr import QrCodeWidget
from reportlab.graphics.shapes import Drawing
from barcode.writer import SVGWriter
from openpyxl import Workbook
from openpyxl.styles import Font
from flask import (
    Flask,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash


# Si se ejecuta como .exe (PyInstaller) se usa la carpeta del ejecutable;
# en desarrollo, la carpeta del proyecto.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

DATABASE = BASE_DIR / "data" / "panol.db"
BARCODE_DIR = BASE_DIR / "static" / "barcodes"
BACKUP_DIR = BASE_DIR / "data" / "backups"
PAGINATION_DEFAULT_SIZE = 25
PAGINATION_MAX_SIZE = 100
BACKUP_RETENTION_DEFAULT = 30
LABELS_MAX_COUNT = 500

LOGIN_MAX_ATTEMPTS = 5
LOGIN_LOCKOUT_MINUTES = 15
_login_attempts: dict[str, tuple[int, datetime | None]] = {}
_login_lock = threading.Lock()

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)
def _load_or_create_secret_key() -> str:
    env_key = os.environ.get("PANOL_SECRET_KEY", "")
    if env_key and len(env_key) >= 32:
        return env_key
    key_file = BASE_DIR / "data" / "secret.key"
    key_file.parent.mkdir(parents=True, exist_ok=True)
    if key_file.exists():
        stored = key_file.read_text().strip()
        if len(stored) >= 32:
            return stored
    new_key = secrets.token_hex(32)
    key_file.write_text(new_key)
    return new_key

app.config["SECRET_KEY"] = _load_or_create_secret_key()
app.config["DATABASE"] = DATABASE
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=8)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("PANOL_HTTPS", "").lower() in {"1", "true", "yes"}
app.config["PAGINATION_DEFAULT_SIZE"] = PAGINATION_DEFAULT_SIZE
app.config["PAGINATION_MAX_SIZE"] = PAGINATION_MAX_SIZE

IDENTIFIER_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{1,39}$")


def backup_retention() -> int:
    """Return the configured number of managed backups to retain."""
    try:
        return max(1, int(os.environ.get("PANOL_BACKUP_RETENTION", BACKUP_RETENTION_DEFAULT)))
    except (TypeError, ValueError):
        return BACKUP_RETENTION_DEFAULT


def create_backup(
    database_path: Path | str | None = None,
    backup_dir: Path | str | None = None,
    retention: int | None = None,
) -> Path:
    """Create a consistent SQLite backup and prune only managed backup files."""
    source_path = Path(database_path or DATABASE)
    target_dir = Path(backup_dir or os.environ.get("PANOL_BACKUP_DIR", str(BACKUP_DIR)))
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    backup_path = target_dir / f"panol-{timestamp}.db"
    source = sqlite3.connect(source_path)
    destination = sqlite3.connect(backup_path)
    try:
        source.backup(destination)
        destination.commit()
    except Exception:
        backup_path.unlink(missing_ok=True)
        raise
    finally:
        source.close()
        destination.close()

    keep = backup_retention() if retention is None else max(1, int(retention))
    managed = sorted(
        target_dir.glob("panol-*.db"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for old_backup in managed[keep:]:
        old_backup.unlink()
    return backup_path


def normalize_text(value: str, max_length: int = 120) -> str:
    return " ".join(value.strip().split())[:max_length]


def validate_identifier(value: str, label: str) -> str | None:
    if not IDENTIFIER_PATTERN.fullmatch(value):
        return f"{label} debe tener entre 2 y 40 caracteres y solo usar letras, números, guion, guion bajo o punto."
    return None


def parse_unit_cost(value: Any, *, allow_empty: bool = True) -> Decimal | None:
    """Parse a finite, non-negative unit cost without introducing binary rounding."""
    text = str(value or "").strip()
    if not text and allow_empty:
        return None
    try:
        cost = Decimal(text)
    except (InvalidOperation, ValueError):
        raise ValueError("El costo unitario debe ser un número finito no negativo.") from None
    if not cost.is_finite() or cost < 0:
        raise ValueError("El costo unitario debe ser un número finito no negativo.")
    return cost


def cost_for_db(value: Decimal | None) -> str | None:
    """Use a decimal string at the SQLite boundary instead of a binary float."""
    return None if value is None else format(value, "f")


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        DATABASE.parent.mkdir(parents=True, exist_ok=True)
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def pagination_params(page_param: str = "page") -> tuple[int, int, int]:
    """Return a bounded page, page size and SQL offset from query parameters."""
    try:
        # Keep ``page`` as a backwards-compatible fallback for independently
        # paginated sections that expose a more specific parameter.
        page = max(1, int(request.args.get(page_param, request.args.get("page", "1"))))
    except (TypeError, ValueError):
        page = 1
    try:
        page_size = int(request.args.get("page_size", str(app.config["PAGINATION_DEFAULT_SIZE"])))
    except (TypeError, ValueError):
        page_size = app.config["PAGINATION_DEFAULT_SIZE"]
    page_size = max(1, min(page_size, app.config["PAGINATION_MAX_SIZE"]))
    return page, page_size, (page - 1) * page_size


def pagination_context(page: int, page_size: int, total: int) -> dict[str, int]:
    return {
        "page": page,
        "page_size": page_size,
        "total": total,
        "pages": max(1, (total + page_size - 1) // page_size),
    }


@app.teardown_appcontext
def close_db(_: BaseException | None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db() -> None:
    db = get_db()
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            full_name TEXT NOT NULL,
            legajo TEXT NOT NULL UNIQUE,
            phone TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            brand TEXT NOT NULL DEFAULT '',
            category TEXT NOT NULL,
            product_type TEXT NOT NULL DEFAULT 'material',
            presentation TEXT NOT NULL DEFAULT 'unidad',
            unit TEXT NOT NULL DEFAULT 'unidad',
            minimum_stock INTEGER NOT NULL DEFAULT 0 CHECK(minimum_stock >= 0),
            stock INTEGER NOT NULL DEFAULT 0 CHECK(stock >= 0),
            reorder_quantity INTEGER NOT NULL DEFAULT 0 CHECK(reorder_quantity >= 0),
            unit_cost REAL NOT NULL DEFAULT 0 CHECK(unit_cost >= 0),
            supplier_id INTEGER REFERENCES suppliers(id),
            location TEXT,
            retornable INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            contact TEXT NOT NULL DEFAULT '',
            phone TEXT NOT NULL DEFAULT '',
            email TEXT NOT NULL DEFAULT '',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL REFERENCES products(id),
            employee_id INTEGER REFERENCES employees(id),
            tool_instance_id INTEGER,
            movement_type TEXT NOT NULL CHECK(movement_type IN ('entrada', 'salida', 'devolucion', 'ajuste')),
            quantity INTEGER NOT NULL CHECK(quantity > 0),
            notes TEXT,
            movement_date TEXT NOT NULL DEFAULT CURRENT_DATE,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS tool_instances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL REFERENCES products(id),
            serial_number TEXT NOT NULL UNIQUE,
            status TEXT NOT NULL DEFAULT 'disponible'
                CHECK(status IN ('disponible', 'entregada', 'reparacion', 'danada', 'perdida', 'baja')),
            current_employee_id INTEGER REFERENCES employees(id),
            notes TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS tool_instance_status_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            instance_id INTEGER NOT NULL REFERENCES tool_instances(id),
            previous_status TEXT NOT NULL,
            new_status TEXT NOT NULL,
            employee_id INTEGER REFERENCES employees(id),
            notes TEXT NOT NULL DEFAULT '',
            changed_by INTEGER REFERENCES users(id),
            changed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'admin',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER REFERENCES users(id),
            action TEXT NOT NULL,
            entity_type TEXT NOT NULL,
            entity_id INTEGER,
            details TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS movement_reversals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_movement_id INTEGER NOT NULL UNIQUE REFERENCES movements(id),
            reversal_movement_id INTEGER NOT NULL UNIQUE REFERENCES movements(id),
            user_id INTEGER REFERENCES users(id),
            reason TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS inventory_counts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            count_date TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'open' CHECK(status IN ('open', 'closed')),
            created_by INTEGER NOT NULL REFERENCES users(id),
            closed_by INTEGER REFERENCES users(id),
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            closed_at TEXT
        );
        CREATE TABLE IF NOT EXISTS inventory_count_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            count_id INTEGER NOT NULL REFERENCES inventory_counts(id) ON DELETE CASCADE,
            product_id INTEGER NOT NULL REFERENCES products(id),
            theoretical_stock INTEGER NOT NULL CHECK(theoretical_stock >= 0),
            counted_stock INTEGER CHECK(counted_stock >= 0),
            UNIQUE(count_id, product_id)
        );
        CREATE TABLE IF NOT EXISTS purchase_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_number TEXT NOT NULL UNIQUE,
            supplier_id INTEGER NOT NULL REFERENCES suppliers(id),
            status TEXT NOT NULL DEFAULT 'borrador'
                CHECK(status IN ('borrador', 'solicitada', 'parcial', 'recibida', 'cancelada')),
            order_date TEXT NOT NULL DEFAULT CURRENT_DATE,
            requested_at TEXT,
            received_at TEXT,
            cancelled_at TEXT,
            notes TEXT NOT NULL DEFAULT '',
            created_by INTEGER NOT NULL REFERENCES users(id),
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS purchase_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_order_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
            product_id INTEGER NOT NULL REFERENCES products(id),
            quantity_requested INTEGER NOT NULL CHECK(quantity_requested > 0),
            quantity_received INTEGER NOT NULL DEFAULT 0 CHECK(quantity_received >= 0),
            unit_cost NUMERIC CHECK(unit_cost IS NULL OR unit_cost >= 0),
            UNIQUE(purchase_order_id, product_id)
        );
        CREATE TABLE IF NOT EXISTS purchase_receipts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_order_id INTEGER NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
            receipt_key TEXT NOT NULL UNIQUE,
            movement_date TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_products_category_name ON products(category, name);
        CREATE INDEX IF NOT EXISTS idx_employees_full_name ON employees(full_name);
        CREATE INDEX IF NOT EXISTS idx_movements_employee_type ON movements(employee_id, movement_type);
        CREATE INDEX IF NOT EXISTS idx_tool_status_history_instance
            ON tool_instance_status_history(instance_id, changed_at DESC);
        CREATE INDEX IF NOT EXISTS idx_purchase_orders_status ON purchase_orders(status, order_date DESC);
        CREATE INDEX IF NOT EXISTS idx_audit_log_action_entity
            ON audit_log(action, entity_type, id DESC);
        CREATE INDEX IF NOT EXISTS idx_audit_log_user_id
            ON audit_log(user_id, id DESC);
        """
    )
    columns = {row["name"] for row in db.execute("PRAGMA table_info(employees)")}
    if "legajo" not in columns:
        db.execute("ALTER TABLE employees RENAME COLUMN document_number TO legajo")
    columns = {row["name"] for row in db.execute("PRAGMA table_info(products)")}
    if "reorder_quantity" not in columns:
        db.execute("ALTER TABLE products ADD COLUMN reorder_quantity INTEGER NOT NULL DEFAULT 0")
    if "supplier_id" not in columns:
        db.execute("ALTER TABLE products ADD COLUMN supplier_id INTEGER")
    if "brand" not in columns:
        db.execute("ALTER TABLE products ADD COLUMN brand TEXT NOT NULL DEFAULT ''")
    if "product_type" not in columns:
        db.execute("ALTER TABLE products ADD COLUMN product_type TEXT NOT NULL DEFAULT 'material'")
    if "presentation" not in columns:
        db.execute("ALTER TABLE products ADD COLUMN presentation TEXT NOT NULL DEFAULT 'unidad'")
    if "unit_cost" not in columns:
        db.execute("ALTER TABLE products ADD COLUMN unit_cost REAL NOT NULL DEFAULT 0 CHECK(unit_cost >= 0)")
    columns = {row["name"] for row in db.execute("PRAGMA table_info(movements)")}
    if "tool_instance_id" not in columns:
        db.execute("ALTER TABLE movements ADD COLUMN tool_instance_id INTEGER")
    if "movement_date" not in columns:
        db.execute("ALTER TABLE movements ADD COLUMN movement_date TEXT")
        db.execute("UPDATE movements SET movement_date = date(created_at) WHERE movement_date IS NULL")
    if "remito" not in columns:
        db.execute("ALTER TABLE movements ADD COLUMN remito TEXT")
    user_columns = {row["name"] for row in db.execute("PRAGMA table_info(users)")}
    if "role" not in user_columns:
        db.execute("ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'admin'")
    product_columns = {row["name"] for row in db.execute("PRAGMA table_info(products)")}
    if "retornable" not in product_columns:
        db.execute("ALTER TABLE products ADD COLUMN retornable INTEGER NOT NULL DEFAULT 0")
    if "active" not in product_columns:
        db.execute("ALTER TABLE products ADD COLUMN active INTEGER NOT NULL DEFAULT 1")
    if db.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        admin_username = normalize_text(os.environ.get("PANOL_ADMIN_USER", "admin"), 40)
        admin_password = os.environ.get("PANOL_ADMIN_PASSWORD", "")
        generated_password = not admin_password
        if generated_password:
            # Sin contraseña configurada no se usa ningún valor fijo: se genera una aleatoria.
            admin_password = secrets.token_urlsafe(12)
        if not admin_username or len(admin_password) < 8:
            raise RuntimeError(
                "Configure PANOL_ADMIN_USER y PANOL_ADMIN_PASSWORD (mínimo 8 caracteres) "
                "para crear el usuario administrador inicial."
            )
        if generated_password:
            print(
                f"[Inventario Pañol] Usuario inicial: {admin_username} | "
                f"Contraseña generada (guárdela y cámbiela): {admin_password}",
                file=sys.stderr,
            )
        db.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'admin')",
            (admin_username, generate_password_hash(admin_password)),
        )
    db.commit()


@app.before_request
def require_login() -> Any:
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
        submitted_token = request.form.get("csrf_token", "") or request.headers.get("X-CSRF-Token", "")
        if not hmac.compare_digest(submitted_token, session["csrf_token"]):
            return "Solicitud no válida. Actualice la página e inténtelo nuevamente.", 400
    if request.endpoint in {"login", "static"}:
        return None
    if "user_id" not in session:
        return redirect(url_for("login", next=request.full_path))
    user = get_db().execute(
        "SELECT id, username, role, active FROM users WHERE id = ?",
        (session["user_id"],),
    ).fetchone()
    if user is None or not user["active"]:
        session.clear()
        flash("La cuenta no está disponible.", "error")
        return redirect(url_for("login"))
    g.user = user
    return None


def role_required(*roles: str):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if getattr(g, "user", None) is None or g.user["role"] not in roles:
                flash("No tiene permisos para realizar esta acción.", "error")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapped
    return decorator


def audit(action: str, entity_type: str, entity_id: int | None, details: str) -> None:
    user_id = g.user["id"] if getattr(g, "user", None) else None
    get_db().execute(
        "INSERT INTO audit_log (user_id, action, entity_type, entity_id, details) VALUES (?, ?, ?, ?, ?)",
        (user_id, action, entity_type, entity_id, normalize_text(details, 1000)),
    )


def can_write() -> bool:
    return getattr(g, "user", None) is not None and g.user["role"] in {
        "admin", "supervisor", "pañolero"
    }


def count_row(count_id: int) -> sqlite3.Row | None:
    return get_db().execute(
        "SELECT * FROM inventory_counts WHERE id = ?", (count_id,)
    ).fetchone()


@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        ip = request.remote_addr or "unknown"
        now = datetime.now()
        with _login_lock:
            _attempts, _locked_until = _login_attempts.get(ip, (0, None))
            if _locked_until and now < _locked_until:
                _remaining = max(1, int((_locked_until - now).total_seconds() / 60) + 1)
                flash(f"Demasiados intentos fallidos. Intente en {_remaining} minuto(s).", "error")
                return render_template("login.html", username=""), 429
        username = normalize_text(request.form.get("username", ""), 40)
        password = request.form.get("password", "")
        user = get_db().execute(
            "SELECT id, username, password_hash, role FROM users WHERE username = ? AND active = 1",
            (username,),
        ).fetchone()
        if user is None or not check_password_hash(user["password_hash"], password):
            with _login_lock:
                _count, _ = _login_attempts.get(ip, (0, None))
                _new_count = _count + 1
                _lockout = datetime.now() + timedelta(minutes=LOGIN_LOCKOUT_MINUTES) if _new_count >= LOGIN_MAX_ATTEMPTS else None
                _login_attempts[ip] = (_new_count, _lockout)
            flash("Usuario o contraseña incorrectos.", "error")
            return render_template("login.html", username=username), 401
        with _login_lock:
            _login_attempts.pop(ip, None)
        session.clear()
        session.permanent = True
        session["csrf_token"] = secrets.token_urlsafe(32)
        session["user_id"] = user["id"]
        session["username"] = user["username"]
        session["role"] = user["role"]
        destination = request.args.get("next", "")
        if not destination.startswith("/") or destination.startswith("//"):
            destination = url_for("dashboard")
        return redirect(destination)
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    flash("Sesión cerrada correctamente.", "success")
    return redirect(url_for("login"))


def generate_barcode(code: str) -> None:
    if validate_identifier(code, "El código"):
        raise ValueError("No se puede generar un código de barras para un identificador inválido.")
    BARCODE_DIR.mkdir(parents=True, exist_ok=True)
    Code128(code, writer=SVGWriter()).save(str(BARCODE_DIR / code))


_FOOTER_TEXT = "Programa Administrativo de Recursos en Obra  |  Desarrollado por: José Azuaje"


def _pie_pagina(canvas, doc):
    """Pie de página común para todos los reportes PDF."""
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColorRGB(0.55, 0.55, 0.55)
    canvas.drawCentredString(doc.pagesize[0] / 2, 8 * mm, _FOOTER_TEXT)
    canvas.restoreState()


def build_labels_pdf(records: list[dict[str, str]]) -> BytesIO:
    """Build a self-contained printable sheet without reading user-provided files."""
    output = BytesIO()
    document = SimpleDocTemplate(
        output, pagesize=A4, rightMargin=12 * mm, leftMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    label_style = ParagraphStyle("label", parent=styles["Normal"], fontSize=10,
                                 leading=12, alignment=TA_CENTER)
    code_style = ParagraphStyle("label-code", parent=label_style, fontName="Helvetica-Bold")
    cells = []
    for record in records:
        barcode_value = record["barcode"]
        barcode_drawing = code128.Code128(barcode_value, barHeight=14 * mm, barWidth=0.28 * mm)
        barcode_drawing.hAlign = "CENTER"
        qr_size = 22 * mm
        qr = QrCodeWidget(barcode_value)
        bounds = qr.getBounds()
        qr_w = bounds[2] - bounds[0]
        qr_h = bounds[3] - bounds[1]
        qr_drawing = Drawing(qr_size, qr_size,
                             transform=[qr_size / qr_w, 0, 0, qr_size / qr_h, 0, 0])
        qr_drawing.add(qr)
        codes_tbl = Table(
            [[barcode_drawing, qr_drawing]],
            colWidths=[30 * mm, 23 * mm],
            rowHeights=[24 * mm],
        )
        codes_tbl.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 1),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        codes_tbl.hAlign = "CENTER"
        cells.append([
            Paragraph(record["name"], label_style),
            Paragraph(f'{record["kind"]}: {record["identifier"]}', code_style),
            codes_tbl,
        ])
    table = Table([cells[i:i + 3] for i in range(0, len(cells), 3)],
                  colWidths=[59 * mm] * 3, rowHeights=[48 * mm] * ((len(cells) + 2) // 3))
    table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#9aa4ad")),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cfd6de")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    document.build(
        [Paragraph("Etiquetas de inventario", styles["Title"]), Spacer(1, 4 * mm), table],
        onFirstPage=_pie_pagina, onLaterPages=_pie_pagina,
    )
    output.seek(0)
    return output


@app.context_processor
def inject_helpers() -> dict[str, Any]:
    def category_class(category: str) -> str:
        normalized = "".join(
            character for character in unicodedata.normalize("NFD", category.lower())
            if unicodedata.category(character) != "Mn"
        )
        categories = {
            "electricidad": "category-electricidad",
            "limpieza": "category-limpieza",
            "higiene y seguridad": "category-higiene",
            "higiene": "category-higiene",
            "seguridad e higiene": "category-higiene",
            "ferreteria": "category-ferreteria",
            "durlock": "category-durlock",
            "colocacion": "category-colocacion",
        }
        return categories.get(normalized, "category-default")

    try:
        db = get_db()
        alert_count = db.execute(
            "SELECT COUNT(*) FROM products WHERE active = 1 AND stock <= minimum_stock"
        ).fetchone()[0]
    except Exception:
        alert_count = 0

    return {
        "low_stock": lambda product: product["stock"] <= product["minimum_stock"],
        "category_class": category_class,
        "csrf_token": session.get("csrf_token", ""),
        "sidebar_alert_count": alert_count,
    }


@app.after_request
def add_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(self)")
    return response


@app.route("/")
def dashboard():
    db = get_db()
    stats = {
        "products": db.execute("SELECT COUNT(*) FROM products WHERE active = 1").fetchone()[0],
        "employees": db.execute("SELECT COUNT(*) FROM employees WHERE active = 1").fetchone()[0],
        "low_stock": db.execute(
            "SELECT COUNT(*) FROM products WHERE active = 1 AND stock <= minimum_stock"
        ).fetchone()[0],
        "pending_purchase_orders": db.execute(
            """SELECT COUNT(*) FROM purchase_orders
               WHERE status IN ('solicitada', 'parcial')"""
        ).fetchone()[0],
        "movements": db.execute(
            "SELECT COUNT(*) FROM movements WHERE movement_date = date('now', 'localtime')"
        ).fetchone()[0],
        "tools_delivered": db.execute(
            "SELECT COUNT(*) FROM tool_instances WHERE status = 'entregada'"
        ).fetchone()[0],
        "inventory_value": db.execute(
            "SELECT COALESCE(SUM(stock * unit_cost), 0) FROM products WHERE active = 1"
        ).fetchone()[0],
        "movements_month": db.execute(
            """
            SELECT COUNT(*) FROM movements
            WHERE movement_date >= date('now', 'localtime', 'start of month')
              AND movement_date < date('now', 'localtime', 'start of month', '+1 month')
            """
        ).fetchone()[0],
        # The instance status is updated atomically with each delivery/return and
        # therefore represents the current net balance per employee and tool.
        "pending_returns": db.execute(
            "SELECT COUNT(*) FROM tool_instances WHERE status = 'entregada'"
        ).fetchone()[0],
    }
    latest = db.execute(
        """
        SELECT m.*, p.name AS product_name, e.full_name AS employee_name
        FROM movements m
        JOIN products p ON p.id = m.product_id
        LEFT JOIN employees e ON e.id = m.employee_id
        ORDER BY m.movement_date DESC, m.id DESC LIMIT 8
        """
    ).fetchall()
    low_products = db.execute(
        "SELECT * FROM products WHERE active = 1 AND stock <= minimum_stock ORDER BY stock, name LIMIT 8"
    ).fetchall()
    pending_orders = db.execute(
        """SELECT po.id, po.order_number, po.status, po.order_date,
                  s.name AS supplier_name,
                  SUM(i.quantity_requested) AS quantity_requested,
                  SUM(i.quantity_received) AS quantity_received,
                  SUM(i.quantity_requested - i.quantity_received) AS quantity_pending
           FROM purchase_orders po
           JOIN suppliers s ON s.id = po.supplier_id
           JOIN purchase_order_items i ON i.purchase_order_id = po.id
           WHERE po.status IN ('solicitada', 'parcial')
           GROUP BY po.id, po.order_number, po.status, po.order_date, s.name
           ORDER BY po.order_date ASC, po.id ASC
           LIMIT 8"""
    ).fetchall()
    return render_template(
        "dashboard.html",
        stats=stats,
        latest=latest,
        low_products=low_products,
        pending_orders=pending_orders,
    )


@app.route("/productos", methods=["GET", "POST"])
def products():
    db = get_db()
    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite modificar productos.", "error")
            return redirect(url_for("products"))
        tipo = normalize_text(request.form.get("tipo", "")).lower()
        if tipo not in {"material", "herramienta"}:
            tipo = ""
        data = {
            key: normalize_text(request.form.get(key, ""))
            for key in ("code", "name", "brand", "category", "product_type", "presentation", "unit", "location")
        }
        if tipo:
            data["product_type"] = tipo
        identifier_error = validate_identifier(data["code"], "El código")
        if identifier_error:
            flash(identifier_error, "error")
            return redirect(url_for("products", tipo=tipo))
        try:
            minimum_stock = int(request.form.get("minimum_stock", "0"))
            stock = int(request.form.get("stock", "0"))
            reorder_quantity = int(request.form.get("reorder_quantity", "0"))
            supplier_id = int(request.form["supplier_id"]) if request.form.get("supplier_id") else None
            unit_cost = parse_unit_cost(request.form.get("unit_cost", "")) or 0
        except ValueError:
            flash("Los stocks deben ser enteros y el costo unitario un número finito no negativo.", "error")
            return redirect(url_for("products", tipo=tipo))
        if data["product_type"] not in {"material", "herramienta"}:
            flash("El tipo de producto no es válido.", "error")
            return redirect(url_for("products", tipo=tipo))
        if not data["name"] or not data["category"] or not data["presentation"] or minimum_stock < 0 or stock < 0 or reorder_quantity < 0:
            flash("Código, nombre, categoría y presentación son obligatorios; los stocks no pueden ser negativos.", "error")
            return redirect(url_for("products", tipo=tipo))
        try:
            with db:
                retornable = 1 if request.form.get("retornable") else 0
                cursor = db.execute(
                    """INSERT INTO products
                       (code, name, brand, category, product_type, presentation, unit, minimum_stock, stock, reorder_quantity, unit_cost, supplier_id, location, retornable)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (data["code"], data["name"], data["brand"], data["category"], data["product_type"],
                     data["presentation"], data["unit"] or "unidad", minimum_stock, stock, reorder_quantity,
                     cost_for_db(unit_cost), supplier_id, data["location"], retornable),
                )
                audit("create", "product", cursor.lastrowid, f"Producto {data['code']} registrado")
            generate_barcode(data["code"])
            flash("Producto registrado correctamente.", "success")
        except sqlite3.IntegrityError:
            flash("El código del producto ya existe.", "error")
        return redirect(url_for("products", tipo=tipo))
    tipo = request.args.get("tipo", "").strip().lower()
    if tipo not in {"material", "herramienta"}:
        tipo = ""
    search = request.args.get("q", "").strip()
    category = request.args.get("category", "").strip()
    show_deleted = request.args.get("eliminados") == "1" and session.get("role") == "admin"
    stock_bajo = request.args.get("stock_bajo") == "1"
    query = """
        SELECT * FROM products
        WHERE (name LIKE ? OR code LIKE ? OR brand LIKE ? OR category LIKE ?)
        AND active = ?
    """
    params: list = [f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%", 0 if show_deleted else 1]
    if category:
        query += " AND category = ?"
        params.append(category)
    if tipo:
        query += " AND product_type = ?"
        params.append(tipo)
    if stock_bajo:
        query += " AND stock <= minimum_stock"
    page, page_size, offset = pagination_params()
    total = db.execute(f"SELECT COUNT(*) FROM ({query})", params).fetchone()[0]
    query += " ORDER BY name, id LIMIT ? OFFSET ?"
    products_list = db.execute(query, [*params, page_size, offset]).fetchall()
    suppliers = db.execute("SELECT id, name FROM suppliers WHERE active = 1 ORDER BY name").fetchall()
    all_categories = db.execute("SELECT DISTINCT category FROM products WHERE active = 1 ORDER BY category").fetchall()
    deleted_count = 0
    if session.get("role") == "admin":
        deleted_count_q = "SELECT COUNT(*) FROM products WHERE active = 0"
        deleted_count_params: list = []
        if tipo:
            deleted_count_q += " AND product_type = ?"
            deleted_count_params.append(tipo)
        deleted_count = db.execute(deleted_count_q, deleted_count_params).fetchone()[0]
    return render_template(
        "products.html", products=products_list, search=search, category=category, suppliers=suppliers,
        pagination=pagination_context(page, page_size, total), tipo=tipo,
        all_categories=all_categories, deleted_count=deleted_count, show_deleted=show_deleted,
        stock_bajo=stock_bajo,
    )


@app.route("/productos/importar", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero")
def import_products():
    if request.method == "POST":
        upload = request.files.get("file")
        if upload is None or not upload.filename:
            flash("Seleccione un archivo CSV.", "error")
            return redirect(url_for("import_products"))
        if not upload.filename.lower().endswith(".csv"):
            flash("El archivo debe tener formato CSV.", "error")
            return redirect(url_for("import_products"))
        try:
            content = upload.read(5 * 1024 * 1024 + 1)
            if len(content) > 5 * 1024 * 1024:
                raise ValueError("El archivo supera el límite de 5 MB.")
            text = content.decode("utf-8-sig")
            reader = csv.DictReader(io.StringIO(text))
            required_headers = {
                "code", "name", "category", "product_type", "presentation",
                "unit", "stock", "minimum_stock",
            }
            headers = set(reader.fieldnames or [])
            missing = required_headers - headers
            if missing:
                raise ValueError(f"Faltan columnas obligatorias: {', '.join(sorted(missing))}.")
            rows = list(reader)
            if not rows:
                raise ValueError("El archivo no contiene filas de datos.")
            if len(rows) > 10000:
                raise ValueError("El archivo no puede superar 10.000 filas.")
        except (UnicodeDecodeError, csv.Error, ValueError) as error:
            flash(f"No se pudo leer el archivo: {error}", "error")
            return redirect(url_for("import_products"))

        db = get_db()
        errors: list[str] = []
        prepared: list[dict[str, Any]] = []
        seen_codes: set[str] = set()
        for line_number, row in enumerate(rows, start=2):
            code = normalize_text(row.get("code") or "", 40)
            name = normalize_text(row.get("name") or "", 120)
            product_type = normalize_text(row.get("product_type") or "", 20).lower()
            try:
                stock = int((row.get("stock") or "0").replace(",", "").replace(".", "").strip())
                minimum_stock = int((row.get("minimum_stock") or "0").replace(",", "").replace(".", "").strip())
                reorder_quantity = int((row.get("reorder_quantity") or "0").replace(",", "").replace(".", "").strip())
            except (ValueError, TypeError):
                errors.append(f"Fila {line_number}: stock, mínimo y reposición deben ser enteros.")
                continue
            if validate_identifier(code, "El código"):
                errors.append(f"Fila {line_number}: código inválido '{code}'.")
            if not name or not (row.get("category") or "").strip() or not (row.get("presentation") or "").strip():
                errors.append(f"Fila {line_number}: nombre, categoría y presentación son obligatorios.")
            if product_type not in {"material", "herramienta"}:
                errors.append(f"Fila {line_number}: product_type debe ser material o herramienta.")
            if stock < 0 or minimum_stock < 0 or reorder_quantity < 0:
                errors.append(f"Fila {line_number}: los valores numéricos no pueden ser negativos.")
            if code in seen_codes:
                errors.append(f"Fila {line_number}: el código {code} está repetido en el archivo.")
            seen_codes.add(code)
            prepared.append({
                "code": code, "name": name, "brand": normalize_text(row.get("brand") or ""),
                "category": normalize_text(row.get("category") or ""),
                "product_type": product_type,
                "presentation": normalize_text(row.get("presentation") or ""),
                "unit": normalize_text(row.get("unit") or "") or "unidad",
                "stock": stock, "minimum_stock": minimum_stock,
                "reorder_quantity": reorder_quantity,
                "location": normalize_text(row.get("location") or ""),
                "supplier_name": normalize_text(row.get("supplier_name") or ""),
            })
        if errors:
            flash("Importación cancelada: " + " ".join(errors[:10]), "error")
            return redirect(url_for("import_products"))

        try:
            with db:
                for item in prepared:
                    supplier_id = None
                    if item["supplier_name"]:
                        supplier = db.execute(
                            "SELECT id FROM suppliers WHERE name = ?",
                            (item["supplier_name"],),
                        ).fetchone()
                        if supplier is None:
                            cursor = db.execute(
                                "INSERT INTO suppliers (name) VALUES (?)",
                                (item["supplier_name"],),
                            )
                            supplier_id = cursor.lastrowid
                        else:
                            supplier_id = supplier["id"]
                    existing = db.execute(
                        "SELECT id FROM products WHERE code = ?", (item["code"],)
                    ).fetchone()
                    if existing:
                        db.execute(
                            """UPDATE products SET name=?, brand=?, category=?, product_type=?,
                               presentation=?, unit=?, stock=?, minimum_stock=?,
                               reorder_quantity=?, supplier_id=?, location=? WHERE id=?""",
                            (
                                item["name"], item["brand"], item["category"], item["product_type"],
                                item["presentation"], item["unit"], item["stock"], item["minimum_stock"],
                                item["reorder_quantity"], supplier_id, item["location"], existing["id"],
                            ),
                        )
                        audit("update", "product", existing["id"], f"Importación masiva: {item['code']}")
                    else:
                        cursor = db.execute(
                            """INSERT INTO products
                               (code, name, brand, category, product_type, presentation, unit,
                                stock, minimum_stock, reorder_quantity, supplier_id, location)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                            (
                                item["code"], item["name"], item["brand"], item["category"],
                                item["product_type"], item["presentation"], item["unit"],
                                item["stock"], item["minimum_stock"], item["reorder_quantity"],
                                supplier_id, item["location"],
                            ),
                        )
                        audit("create", "product", cursor.lastrowid, f"Importación masiva: {item['code']}")
            for item in prepared:
                generate_barcode(item["code"])
        except sqlite3.IntegrityError as error:
            flash(f"Importación cancelada por un conflicto de datos: {error}", "error")
            return redirect(url_for("import_products"))
        flash(f"Importación completada: {len(prepared)} productos procesados.", "success")
        return redirect(url_for("products"))
    return render_template("import_products.html")


@app.route("/productos/importar/plantilla")
@role_required("admin", "supervisor", "pañolero")
def products_import_template():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "code", "name", "brand", "category", "product_type", "presentation",
        "unit", "stock", "minimum_stock", "reorder_quantity", "supplier_name", "location",
    ])
    writer.writerow([
        "MAT-001", "Guantes de trabajo", "3M", "Higiene y seguridad", "material",
        "Caja x 10", "unidad", "100", "20", "50", "Proveedor ejemplo", "Estante A-01",
    ])
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name="plantilla-productos-panol.csv",
    )


@app.route("/productos/<int:product_id>/editar", methods=["GET", "POST"])
def edit_product(product_id: int):
    db = get_db()
    product = db.execute("SELECT * FROM products WHERE id = ?", (product_id,)).fetchone()
    if product is None:
        flash("El producto no existe.", "error")
        return redirect(url_for("products"))
    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite modificar productos.", "error")
            return redirect(url_for("products"))
        data = {
            key: normalize_text(request.form.get(key, ""))
            for key in ("code", "name", "brand", "category", "product_type", "presentation", "unit", "location")
        }
        identifier_error = validate_identifier(data["code"], "El código")
        if identifier_error:
            flash(identifier_error, "error")
            return redirect(url_for("edit_product", product_id=product_id))
        try:
            minimum_stock = int(request.form.get("minimum_stock", "0"))
            stock = int(request.form.get("stock", "0"))
            reorder_quantity = int(request.form.get("reorder_quantity", "0"))
            supplier_id = int(request.form["supplier_id"]) if request.form.get("supplier_id") else None
            unit_cost = parse_unit_cost(request.form.get("unit_cost", "")) or 0
        except ValueError:
            flash("Los stocks deben ser enteros y el costo unitario un número finito no negativo.", "error")
            return redirect(url_for("edit_product", product_id=product_id))
        retornable = 1 if request.form.get("retornable") else 0
        if (
            not data["code"] or not data["name"] or not data["category"]
            or not data["presentation"] or data["product_type"] not in {"material", "herramienta"}
            or minimum_stock < 0 or stock < 0 or reorder_quantity < 0
        ):
            flash("Revisá los datos obligatorios y los valores de stock.", "error")
            return redirect(url_for("edit_product", product_id=product_id))
        try:
            with db:
                db.execute(
                    """UPDATE products SET code = ?, name = ?, brand = ?, category = ?, product_type = ?,
                       presentation = ?, unit = ?, minimum_stock = ?, stock = ?, reorder_quantity = ?, unit_cost = ?, supplier_id = ?, location = ?, retornable = ?
                       WHERE id = ?""",
                    (
                        data["code"], data["name"], data["brand"], data["category"], data["product_type"],
                        data["presentation"], data["unit"] or "unidad", minimum_stock, stock, reorder_quantity,
                        cost_for_db(unit_cost), supplier_id, data["location"], retornable, product_id,
                    ),
                )
                audit("update", "product", product_id, f"Producto {data['code']} modificado")
            if product["code"] != data["code"]:
                old_barcode = BARCODE_DIR / f"{product['code']}.svg"
                if old_barcode.exists():
                    old_barcode.unlink()
            generate_barcode(data["code"])
            flash("Producto modificado correctamente.", "success")
            return redirect(url_for("products"))
        except sqlite3.IntegrityError:
            flash("El código del producto ya existe.", "error")
            return redirect(url_for("edit_product", product_id=product_id))
    suppliers = db.execute("SELECT id, name FROM suppliers WHERE active = 1 ORDER BY name").fetchall()
    all_categories = db.execute("SELECT DISTINCT category FROM products WHERE active = 1 ORDER BY category").fetchall()
    product_history = db.execute(
        """SELECT m.movement_date, m.movement_type, m.quantity, m.notes, m.remito,
                  e.full_name AS employee_name, m.created_at, m.id
           FROM movements m
           LEFT JOIN employees e ON e.id = m.employee_id
           WHERE m.product_id = ?
           ORDER BY m.movement_date DESC, m.id DESC
           LIMIT 30""",
        (product_id,),
    ).fetchall()
    barcode_url = url_for("barcodes", filename=f"{product['code']}.svg") if (BARCODE_DIR / f"{product['code']}.svg").exists() else None
    return render_template(
        "edit_product.html",
        product=product,
        suppliers=suppliers,
        all_categories=all_categories,
        product_history=product_history,
        barcode_url=barcode_url,
    )


@app.route("/productos/<int:product_id>/historial.csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def product_history_csv(product_id: int):
    db = get_db()
    product = db.execute("SELECT name, code FROM products WHERE id = ?", (product_id,)).fetchone()
    if product is None:
        flash("El producto no existe.", "error")
        return redirect(url_for("products"))
    rows = db.execute(
        """SELECT m.movement_date, m.movement_type, m.quantity,
                  COALESCE(e.full_name, '') AS employee_name,
                  COALESCE(m.remito, '') AS remito,
                  COALESCE(m.notes, '') AS notes,
                  m.created_at
           FROM movements m
           LEFT JOIN employees e ON e.id = m.employee_id
           WHERE m.product_id = ?
           ORDER BY m.movement_date DESC, m.id DESC""",
        (product_id,),
    ).fetchall()
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Fecha", "Tipo", "Cantidad", "Responsable", "Remito", "Notas", "Registrado"])
    for row in rows:
        writer.writerow([row["movement_date"], row["movement_type"], row["quantity"],
                         row["employee_name"], row["remito"], row["notes"], row["created_at"]])
    filename = f"historial-{product['code']}.csv"
    return Response(
        "﻿" + output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.route("/empleados/<int:employee_id>/historial.csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def employee_history_csv(employee_id: int):
    db = get_db()
    employee = db.execute("SELECT full_name, code FROM employees WHERE id = ?", (employee_id,)).fetchone()
    if employee is None:
        flash("El empleado no existe.", "error")
        return redirect(url_for("employees"))
    rows = db.execute(
        """SELECT m.movement_date, m.movement_type, m.quantity,
                  p.name AS product_name, p.code AS product_code,
                  COALESCE(m.remito, '') AS remito,
                  COALESCE(m.notes, '') AS notes,
                  m.created_at
           FROM movements m
           JOIN products p ON p.id = m.product_id
           WHERE m.employee_id = ?
           ORDER BY m.movement_date DESC, m.id DESC""",
        (employee_id,),
    ).fetchall()
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Fecha", "Tipo", "Producto", "Código", "Cantidad", "Remito", "Notas", "Registrado"])
    for row in rows:
        writer.writerow([row["movement_date"], row["movement_type"], row["product_name"],
                         row["product_code"], row["quantity"], row["remito"], row["notes"], row["created_at"]])
    filename = f"historial-empleado-{employee['code']}.csv"
    return Response(
        "﻿" + output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.route("/productos/<int:product_id>/eliminar", methods=["POST"])
@role_required("admin")
def deactivate_product(product_id: int):
    db = get_db()
    product = db.execute("SELECT * FROM products WHERE id = ?", (product_id,)).fetchone()
    if product is None:
        flash("El producto no existe.", "error")
        return redirect(url_for("products"))
    if product["active"] == 0:
        flash("El producto ya estaba dado de baja.", "warning")
        return redirect(url_for("edit_product", product_id=product_id))
    delivered = db.execute(
        """SELECT COUNT(*) FROM tool_instances
           WHERE product_id = ? AND status = 'entregada'""",
        (product_id,),
    ).fetchone()[0]
    if delivered > 0:
        flash(
            f"No se puede dar de baja: hay {delivered} unidad(es) de esta herramienta actualmente entregadas.",
            "error",
        )
        return redirect(url_for("edit_product", product_id=product_id))
    with db:
        db.execute("UPDATE products SET active = 0 WHERE id = ?", (product_id,))
        audit("update", "product", product_id, f"Baja lógica: {product['name']} ({product['code']})")
    flash(f"'{product['name']}' dado de baja. El historial de movimientos se conserva.", "success")
    return redirect(url_for("products", tipo=product["product_type"]))


@app.route("/productos/<int:product_id>/restaurar", methods=["POST"])
@role_required("admin")
def restore_product(product_id: int):
    db = get_db()
    product = db.execute("SELECT * FROM products WHERE id = ?", (product_id,)).fetchone()
    if product is None:
        flash("El producto no existe.", "error")
        return redirect(url_for("products"))
    with db:
        db.execute("UPDATE products SET active = 1 WHERE id = ?", (product_id,))
        audit("update", "product", product_id, f"Restauración: {product['name']} ({product['code']})")
    flash(f"'{product['name']}' restaurado correctamente.", "success")
    return redirect(url_for("edit_product", product_id=product_id))


@app.route("/empleados", methods=["GET", "POST"])
def employees():
    db = get_db()
    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite modificar empleados.", "error")
            return redirect(url_for("employees"))
        code = normalize_text(request.form.get("code", ""))
        name = normalize_text(request.form.get("full_name", ""))
        legajo = normalize_text(request.form.get("legajo", ""))
        phone = normalize_text(request.form.get("phone", ""), 30)
        code_error = validate_identifier(code, "El código")
        legajo_error = validate_identifier(legajo, "El legajo")
        if code_error or legajo_error:
            flash(code_error or legajo_error, "error")
            return redirect(url_for("employees"))
        if not code or not name or not legajo:
            flash("Código, nombre y legajo son obligatorios.", "error")
            return redirect(url_for("employees"))
        try:
            with db:
                cursor = db.execute(
                    "INSERT INTO employees (code, full_name, legajo, phone) VALUES (?, ?, ?, ?)",
                    (code, name, legajo, phone),
                )
                audit("create", "employee", cursor.lastrowid, f"Empleado {code} registrado")
            generate_barcode(code)
            flash("Empleado registrado correctamente.", "success")
        except sqlite3.IntegrityError:
            flash("El código o legajo del empleado ya existe.", "error")
        return redirect(url_for("employees"))
    search = request.args.get("q", "").strip()
    employee_query = """SELECT e.*,
                               (SELECT COUNT(*) FROM tool_instances ti
                                WHERE ti.current_employee_id = e.id AND ti.status = 'entregada') AS tools_out
                        FROM employees e
                        WHERE e.full_name LIKE ? OR e.code LIKE ? OR e.legajo LIKE ?"""
    employee_params = (f"%{search}%", f"%{search}%", f"%{search}%")
    page, page_size, offset = pagination_params()
    total = db.execute(f"SELECT COUNT(*) FROM ({employee_query})", employee_params).fetchone()[0]
    employees_list = db.execute(
        employee_query + " ORDER BY e.full_name, e.id LIMIT ? OFFSET ?",
        (*employee_params, page_size, offset),
    ).fetchall()
    return render_template(
        "employees.html", employees=employees_list, search=search,
        pagination=pagination_context(page, page_size, total),
    )


@app.route("/empleados/importar", methods=["GET", "POST"])
@role_required("admin", "supervisor")
def import_employees():
    if request.method == "POST":
        upload = request.files.get("file")
        if upload is None or not upload.filename:
            flash("Seleccione un archivo CSV.", "error")
            return redirect(url_for("import_employees"))
        if not upload.filename.lower().endswith(".csv"):
            flash("El archivo debe tener formato CSV.", "error")
            return redirect(url_for("import_employees"))
        try:
            content = upload.read(5 * 1024 * 1024 + 1)
            if len(content) > 5 * 1024 * 1024:
                raise ValueError("El archivo supera el límite de 5 MB.")
            text = content.decode("utf-8-sig")
            reader = csv.DictReader(io.StringIO(text))
            required_headers = {"code", "full_name", "legajo"}
            headers = set(reader.fieldnames or [])
            missing = required_headers - headers
            if missing:
                raise ValueError(f"Faltan columnas obligatorias: {', '.join(sorted(missing))}.")
            rows = list(reader)
            if not rows:
                raise ValueError("El archivo no contiene filas de datos.")
            if len(rows) > 5000:
                raise ValueError("El archivo no puede superar 5.000 filas.")
        except (UnicodeDecodeError, csv.Error, ValueError) as error:
            flash(f"No se pudo leer el archivo: {error}", "error")
            return redirect(url_for("import_employees"))

        errors: list[str] = []
        prepared: list[dict[str, Any]] = []
        seen_codes: set[str] = set()
        seen_legajos: set[str] = set()
        for line_number, row in enumerate(rows, start=2):
            code = normalize_text(row.get("code", ""), 40)
            full_name = normalize_text(row.get("full_name", ""), 120)
            legajo = normalize_text(row.get("legajo", "") or code, 40)
            phone = normalize_text(row.get("phone", ""), 30)
            active_raw = (row.get("active") or "1").strip()
            active = 0 if active_raw in ("0", "false", "False", "no", "No") else 1
            if validate_identifier(code, "El código"):
                errors.append(f"Fila {line_number}: código inválido '{code}'.")
                continue
            if validate_identifier(legajo, "El legajo"):
                errors.append(f"Fila {line_number}: legajo inválido '{legajo}'.")
                continue
            if not code or not full_name or not legajo:
                errors.append(f"Fila {line_number}: código, nombre y legajo son obligatorios.")
                continue
            if code in seen_codes:
                errors.append(f"Fila {line_number}: el código '{code}' está repetido en el archivo.")
                continue
            if legajo in seen_legajos:
                errors.append(f"Fila {line_number}: el legajo '{legajo}' está repetido en el archivo.")
                continue
            seen_codes.add(code)
            seen_legajos.add(legajo)
            prepared.append({
                "code": code, "full_name": full_name, "legajo": legajo,
                "phone": phone, "active": active,
            })
        if errors:
            flash("Importación cancelada: " + " ".join(errors[:10]), "error")
            return redirect(url_for("import_employees"))

        db = get_db()
        created = 0
        updated = 0
        try:
            with db:
                for item in prepared:
                    existing = db.execute(
                        "SELECT id FROM employees WHERE code = ?", (item["code"],)
                    ).fetchone()
                    if existing:
                        db.execute(
                            """UPDATE employees SET full_name=?, legajo=?, phone=?, active=?
                               WHERE id=?""",
                            (item["full_name"], item["legajo"], item["phone"],
                             item["active"], existing["id"]),
                        )
                        audit("update", "employee", existing["id"],
                              f"Importación masiva: {item['code']}")
                        updated += 1
                    else:
                        cursor = db.execute(
                            """INSERT INTO employees (code, full_name, legajo, phone, active)
                               VALUES (?, ?, ?, ?, ?)""",
                            (item["code"], item["full_name"], item["legajo"],
                             item["phone"], item["active"]),
                        )
                        audit("create", "employee", cursor.lastrowid,
                              f"Importación masiva: {item['code']}")
                        created += 1
            for item in prepared:
                generate_barcode(item["code"])
        except sqlite3.IntegrityError as error:
            flash(f"Importación cancelada por conflicto de datos: {error}", "error")
            return redirect(url_for("import_employees"))
        except Exception as error:
            app.logger.exception("Error inesperado en importación masiva")
            flash(f"Error al importar: {error}", "error")
            return redirect(url_for("import_employees"))
        flash(
            f"Importación completada: {created} nuevos, {updated} actualizados.",
            "success",
        )
        return redirect(url_for("employees"))
    return render_template("import_employees.html")


@app.route("/empleados/importar/plantilla")
@role_required("admin", "supervisor")
def employees_import_template():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["code", "full_name", "legajo", "phone", "active"])
    writer.writerow(["EMP-001", "Juan Pérez", "12345", "+54 11 1234-5678", "1"])
    writer.writerow(["EMP-002", "María García", "67890", "", "1"])
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name="plantilla-empleados-panol.csv",
    )


@app.route("/empleados/<int:employee_id>/editar", methods=["GET", "POST"])
def edit_employee(employee_id: int):
    db = get_db()
    employee = db.execute("SELECT * FROM employees WHERE id = ?", (employee_id,)).fetchone()
    if employee is None:
        flash("El empleado no existe.", "error")
        return redirect(url_for("employees"))
    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite modificar empleados.", "error")
            return redirect(url_for("employees"))
        code = normalize_text(request.form.get("code", ""))
        name = normalize_text(request.form.get("full_name", ""))
        legajo = normalize_text(request.form.get("legajo", ""))
        phone = normalize_text(request.form.get("phone", ""), 30)
        code_error = validate_identifier(code, "El código")
        legajo_error = validate_identifier(legajo, "El legajo")
        if code_error or legajo_error:
            flash(code_error or legajo_error, "error")
            return redirect(url_for("edit_employee", employee_id=employee_id))
        if not code or not name or not legajo:
            flash("Código, nombre y legajo son obligatorios.", "error")
            return redirect(url_for("edit_employee", employee_id=employee_id))
        try:
            with db:
                db.execute(
                    "UPDATE employees SET code = ?, full_name = ?, legajo = ?, phone = ? WHERE id = ?",
                    (code, name, legajo, phone, employee_id),
                )
                audit("update", "employee", employee_id, f"Empleado {code} modificado")
            if employee["code"] != code:
                old_barcode = BARCODE_DIR / f"{employee['code']}.svg"
                if old_barcode.exists():
                    old_barcode.unlink()
            generate_barcode(code)
            flash("Empleado modificado correctamente.", "success")
            return redirect(url_for("employees"))
        except sqlite3.IntegrityError:
            flash("El código o legajo del empleado ya existe.", "error")
            return redirect(url_for("edit_employee", employee_id=employee_id))
    current_tools = db.execute(
        """SELECT ti.serial_number, p.name AS product_name, p.code AS product_code, p.id AS product_id
           FROM tool_instances ti
           JOIN products p ON p.id = ti.product_id
           WHERE ti.status = 'entregada' AND ti.current_employee_id = ?
           ORDER BY p.name, ti.serial_number""",
        (employee_id,),
    ).fetchall()
    employee_history = db.execute(
        """SELECT m.movement_date, m.movement_type, m.quantity, m.remito, m.notes,
                  p.name AS product_name, p.code AS product_code, p.product_type, m.id
           FROM movements m
           JOIN products p ON p.id = m.product_id
           WHERE m.employee_id = ?
           ORDER BY m.movement_date DESC, m.id DESC
           LIMIT 30""",
        (employee_id,),
    ).fetchall()
    emp_barcode_url = url_for("barcodes", filename=f"{employee['code']}.svg") if (BARCODE_DIR / f"{employee['code']}.svg").exists() else None
    return render_template(
        "edit_employee.html",
        employee=employee,
        current_tools=current_tools,
        employee_history=employee_history,
        emp_barcode_url=emp_barcode_url,
    )


@app.route("/movimientos", methods=["GET", "POST"])
def movements():
    db = get_db()
    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite registrar movimientos.", "error")
            return redirect(url_for("movements"))
        product_type = request.form.get("product_type", "")
        movement_type = request.form.get("movement_type", "")
        movement_date = request.form.get("movement_date", "").strip()
        notes = normalize_text(request.form.get("notes", ""), 500)
        remito = normalize_text(request.form.get("remito", ""), 60) or None
        try:
            date.fromisoformat(movement_date)
        except ValueError:
            flash("La fecha del movimiento no es válida.", "error")
            return redirect(url_for("movements"))
        if movement_type not in {"entrada", "salida", "devolucion", "ajuste"}:
            flash("Tipo de movimiento inválido.", "error")
            return redirect(url_for("movements"))

        if product_type == "herramienta":
            try:
                employee_id = int(request.form["employee_id"]) if request.form.get("employee_id") else None
            except ValueError:
                flash("Responsable inválido.", "error")
                return redirect(url_for("movements"))
            if movement_type == "salida" and employee_id is None:
                flash("El responsable es obligatorio para entregar herramientas.", "error")
                return redirect(url_for("movements"))
            if employee_id is not None:
                emp = db.execute(
                    "SELECT id FROM employees WHERE id = ? AND active = 1", (employee_id,)
                ).fetchone()
                if emp is None:
                    flash("El responsable seleccionado no existe o está inactivo.", "error")
                    return redirect(url_for("movements"))
            product_ids_raw = request.form.getlist("product_id[]")
            tool_instance_ids_raw = request.form.getlist("tool_instance_id[]")
            quantities_raw = request.form.getlist("quantity[]")
            unit_serials_raw = request.form.getlist("unit_serial[]")
            if not product_ids_raw:
                flash("Debe seleccionar al menos una herramienta.", "error")
                return redirect(url_for("movements"))
            validated = []
            for i, pid_raw in enumerate(product_ids_raw):
                try:
                    pid = int(pid_raw)
                    qty = int(quantities_raw[i]) if i < len(quantities_raw) and quantities_raw[i] else 1
                    tid_raw = tool_instance_ids_raw[i] if i < len(tool_instance_ids_raw) else ""
                    tid = int(tid_raw) if tid_raw else None
                    unit_serial = normalize_text(
                        unit_serials_raw[i] if i < len(unit_serials_raw) else "", 80
                    )
                except ValueError:
                    flash(f"Fila {i + 1}: herramienta o cantidad inválidos.", "error")
                    return redirect(url_for("movements"))
                if qty <= 0:
                    flash(f"Fila {i + 1}: la cantidad debe ser mayor a 0.", "error")
                    return redirect(url_for("movements"))
                product = db.execute(
                    "SELECT * FROM products WHERE id = ? AND product_type = 'herramienta'", (pid,)
                ).fetchone()
                if product is None:
                    flash(f"Fila {i + 1}: herramienta no encontrada.", "error")
                    return redirect(url_for("movements"))
                create_new_serial = False
                if unit_serial and tid is None:
                    existing_inst = db.execute(
                        "SELECT id FROM tool_instances WHERE product_id = ? AND serial_number = ?",
                        (pid, unit_serial),
                    ).fetchone()
                    if existing_inst:
                        tid = existing_inst["id"]
                    else:
                        create_new_serial = True
                if tid is not None:
                    if movement_type not in {"entrada", "salida", "devolucion"}:
                        flash(
                            "Las unidades individualizadas solo admiten ingreso, entrega o devolución.",
                            "error",
                        )
                        return redirect(url_for("movements"))
                    if qty != 1:
                        flash("Las herramientas individualizadas se mueven de a una unidad.", "error")
                        return redirect(url_for("movements"))
                    instance = db.execute(
                        "SELECT * FROM tool_instances WHERE id = ? AND product_id = ?", (tid, pid)
                    ).fetchone()
                    if instance is None:
                        flash(f"Fila {i + 1}: la unidad no pertenece al producto indicado.", "error")
                        return redirect(url_for("movements"))
                    if movement_type == "salida":
                        if instance["status"] in {"reparacion", "danada", "perdida", "baja"}:
                            flash(
                                f"No se puede entregar {product['name']}: estado '{instance['status']}'.",
                                "error",
                            )
                            return redirect(url_for("movements"))
                        if instance["status"] == "entregada":
                            flash(
                                f"{product['name']} ya está entregada. Registre su devolución primero.",
                                "error",
                            )
                            return redirect(url_for("movements"))
                    elif movement_type == "devolucion":
                        if instance["status"] != "entregada":
                            flash(f"{product['name']} no está entregada actualmente.", "error")
                            return redirect(url_for("movements"))
                    elif instance["status"] == "entregada":
                        flash(
                            f"No se puede ingresar {product['name']}: registre primero su devolución.",
                            "error",
                        )
                        return redirect(url_for("movements"))
                delta = qty if movement_type in {"entrada", "devolucion"} else -qty
                new_stock = product["stock"] + delta
                if new_stock < 0:
                    flash(
                        f"Stock insuficiente para {product['name']}. Disponible: {product['stock']}.",
                        "error",
                    )
                    return redirect(url_for("movements"))
                validated.append({
                    "product": product,
                    "tool_instance_id": tid,
                    "create_serial": unit_serial if create_new_serial else None,
                    "quantity": qty,
                    "new_stock": new_stock,
                })
            movement_ids: list[int] = []
            with db:
                for v in validated:
                    if v["create_serial"]:
                        db.execute(
                            "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)",
                            (v["product"]["id"], v["create_serial"]),
                        )
                        v["tool_instance_id"] = db.execute("SELECT last_insert_rowid()").fetchone()[0]
                    db.execute(
                        """INSERT INTO movements
                           (product_id, employee_id, tool_instance_id, movement_type, quantity, notes, movement_date, remito)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (v["product"]["id"], employee_id, v["tool_instance_id"],
                         movement_type, v["quantity"], notes, movement_date, remito),
                    )
                    movement_ids.append(db.execute("SELECT last_insert_rowid()").fetchone()[0])
                    db.execute(
                        "UPDATE products SET stock = ? WHERE id = ?",
                        (v["new_stock"], v["product"]["id"]),
                    )
                    if v["tool_instance_id"] is not None:
                        next_status = "entregada" if movement_type == "salida" else "disponible"
                        next_employee = employee_id if movement_type == "salida" else None
                        db.execute(
                            "UPDATE tool_instances SET status = ?, current_employee_id = ? WHERE id = ?",
                            (next_status, next_employee, v["tool_instance_id"]),
                        )
                    audit(
                        "create", "movement", movement_ids[-1],
                        f"{movement_type} de {v['quantity']} unidad(es) de {v['product']['code']}",
                    )
            return redirect(url_for("movements_batch_receipt", ids=",".join(str(i) for i in movement_ids)))

        # Materials (multi-item via product_id[] array)
        try:
            employee_id = int(request.form["employee_id"]) if request.form.get("employee_id") else None
        except ValueError:
            flash("Responsable inválido.", "error")
            return redirect(url_for("movements"))
        product_ids_raw = request.form.getlist("product_id[]")
        quantities_raw = request.form.getlist("quantity[]")
        if not product_ids_raw:
            flash("Debe escanear al menos un material.", "error")
            return redirect(url_for("movements"))
        if movement_type == "salida" and employee_id is None:
            flash("El responsable es obligatorio para entregar materiales.", "error")
            return redirect(url_for("movements"))
        if employee_id is not None:
            emp = db.execute(
                "SELECT id FROM employees WHERE id = ? AND active = 1", (employee_id,)
            ).fetchone()
            if emp is None:
                flash("El responsable seleccionado no existe o está inactivo.", "error")
                return redirect(url_for("movements"))
        validated = []
        for i, pid_raw in enumerate(product_ids_raw):
            try:
                pid = int(pid_raw)
                qty = int(quantities_raw[i]) if i < len(quantities_raw) and quantities_raw[i] else 1
            except ValueError:
                flash(f"Fila {i + 1}: producto o cantidad inválidos.", "error")
                return redirect(url_for("movements"))
            if qty <= 0:
                flash(f"Fila {i + 1}: la cantidad debe ser mayor a 0.", "error")
                return redirect(url_for("movements"))
            product = db.execute(
                "SELECT * FROM products WHERE id = ? AND product_type = 'material'", (pid,)
            ).fetchone()
            if product is None:
                flash(f"Fila {i + 1}: material no encontrado.", "error")
                return redirect(url_for("movements"))
            delta = qty if movement_type in {"entrada", "devolucion"} else -qty
            new_stock = product["stock"] + delta
            if new_stock < 0:
                flash(
                    f"Stock insuficiente para {product['name']}. Disponible: {product['stock']}.",
                    "error",
                )
                return redirect(url_for("movements"))
            validated.append({"product": product, "quantity": qty, "new_stock": new_stock})
        movement_ids: list[int] = []
        with db:
            for v in validated:
                db.execute(
                    """INSERT INTO movements
                       (product_id, employee_id, tool_instance_id, movement_type, quantity, notes, movement_date, remito)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (v["product"]["id"], employee_id, None, movement_type, v["quantity"], notes, movement_date, remito),
                )
                movement_ids.append(db.execute("SELECT last_insert_rowid()").fetchone()[0])
                db.execute(
                    "UPDATE products SET stock = ? WHERE id = ?",
                    (v["new_stock"], v["product"]["id"]),
                )
                audit(
                    "create", "movement", movement_ids[-1],
                    f"{movement_type} de {v['quantity']} unidad(es) del material {v['product']['code']}",
                )
        return redirect(url_for("movements_batch_receipt", ids=",".join(str(i) for i in movement_ids)))
    products_list = db.execute("SELECT * FROM products WHERE active = 1 ORDER BY name").fetchall()
    material_products = [product for product in products_list if product["product_type"] == "material"]
    tool_products = [product for product in products_list if product["product_type"] == "herramienta"]
    tool_instances = db.execute(
        """SELECT ti.*, p.name AS product_name, p.code AS product_code
           FROM tool_instances ti JOIN products p ON p.id = ti.product_id
           WHERE ti.status NOT IN ('baja', 'perdida') ORDER BY p.name, ti.serial_number"""
    ).fetchall()
    employees_list = db.execute("SELECT * FROM employees WHERE active = 1 ORDER BY full_name").fetchall()
    all_employees = db.execute("SELECT id, full_name FROM employees ORDER BY full_name").fetchall()
    search = request.args.get("q", "").strip()
    movement_type = request.args.get("movement_type", "").strip()
    employee_id_filter = request.args.get("employee_id", "").strip()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    history_query = """SELECT m.*, p.name AS product_name, p.code AS product_code,
                              p.product_type, e.full_name AS employee_name
                       FROM movements m JOIN products p ON p.id = m.product_id
                       LEFT JOIN employees e ON e.id = m.employee_id"""
    history_conditions: list[str] = []
    history_params: list[object] = []
    if search:
        history_conditions.append(
            "(p.name LIKE ? OR p.code LIKE ? OR e.full_name LIKE ? OR m.notes LIKE ? OR m.remito LIKE ?)"
        )
        search_pattern = f"%{search}%"
        history_params.extend([search_pattern] * 5)
    if movement_type in {"entrada", "salida", "devolucion", "ajuste"}:
        history_conditions.append("m.movement_type = ?")
        history_params.append(movement_type)
    else:
        movement_type = ""
    try:
        employee_id_value = int(employee_id_filter) if employee_id_filter else None
    except ValueError:
        employee_id_value = None
        employee_id_filter = ""
    if employee_id_value is not None:
        history_conditions.append("m.employee_id = ?")
        history_params.append(employee_id_value)
    try:
        date.fromisoformat(date_from)
    except ValueError:
        date_from = ""
    try:
        date.fromisoformat(date_to)
    except ValueError:
        date_to = ""
    if date_from:
        history_conditions.append("m.movement_date >= ?")
        history_params.append(date_from)
    if date_to:
        history_conditions.append("m.movement_date <= ?")
        history_params.append(date_to)
    if history_conditions:
        history_query += " WHERE " + " AND ".join(history_conditions)
    page, page_size, offset = pagination_params()
    total = db.execute(
        f"SELECT COUNT(*) FROM ({history_query})", history_params
    ).fetchone()[0]
    history_query += " ORDER BY m.id DESC LIMIT ? OFFSET ?"
    history = db.execute(history_query, [*history_params, page_size, offset]).fetchall()
    return render_template(
        "movements.html", products=products_list, material_products=material_products,
        tool_products=tool_products, tool_instances=tool_instances, employees=employees_list, history=history,
        all_employees=all_employees,
        today=date.today().isoformat(), search=search, movement_type=movement_type,
        employee_id_filter=employee_id_filter,
        date_from=date_from, date_to=date_to,
        pagination=pagination_context(page, page_size, total),
    )


@app.route("/movimientos/comprobante-lote")
def movements_batch_receipt():
    raw = request.args.get("ids", "")
    try:
        ids = [int(i) for i in raw.split(",") if i.strip()][:50]
    except ValueError:
        flash("Parámetro de comprobante inválido.", "error")
        return redirect(url_for("movements"))
    if not ids:
        return redirect(url_for("movements"))
    placeholders = ",".join("?" * len(ids))
    db = get_db()
    rows = db.execute(
        f"""SELECT m.*, p.name AS product_name, p.code AS product_code,
                   p.unit, p.product_type,
                   e.full_name AS employee_name, e.legajo,
                   ti.serial_number
            FROM movements m
            JOIN products p ON p.id = m.product_id
            LEFT JOIN employees e ON e.id = m.employee_id
            LEFT JOIN tool_instances ti ON ti.id = m.tool_instance_id
            WHERE m.id IN ({placeholders})
            ORDER BY m.id""",
        ids,
    ).fetchall()
    if not rows:
        flash("Comprobante no encontrado.", "error")
        return redirect(url_for("movements"))
    return render_template("movement_batch_receipt.html", movements=rows, ids_raw=raw)


@app.route("/movimientos/exportar.csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def movements_export_csv():
    db = get_db()
    search = request.args.get("q", "").strip()
    movement_type = request.args.get("movement_type", "").strip()
    employee_id_filter = request.args.get("employee_id", "").strip()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    query = """SELECT m.movement_date, m.movement_type, m.quantity,
                      p.name AS product_name, p.code AS product_code, p.product_type,
                      COALESCE(e.full_name, '') AS employee_name,
                      COALESCE(m.remito, '') AS remito,
                      COALESCE(m.notes, '') AS notes,
                      m.created_at
               FROM movements m
               JOIN products p ON p.id = m.product_id
               LEFT JOIN employees e ON e.id = m.employee_id"""
    conditions: list[str] = []
    params: list[object] = []
    if search:
        conditions.append("(p.name LIKE ? OR p.code LIKE ? OR e.full_name LIKE ? OR m.notes LIKE ? OR m.remito LIKE ?)")
        params.extend([f"%{search}%"] * 5)
    if movement_type in {"entrada", "salida", "devolucion", "ajuste"}:
        conditions.append("m.movement_type = ?")
        params.append(movement_type)
    try:
        emp_id = int(employee_id_filter) if employee_id_filter else None
    except ValueError:
        emp_id = None
    if emp_id is not None:
        conditions.append("m.employee_id = ?")
        params.append(emp_id)
    for field, value in (("m.movement_date >= ?", date_from), ("m.movement_date <= ?", date_to)):
        try:
            date.fromisoformat(value)
            conditions.append(field)
            params.append(value)
        except ValueError:
            pass
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " ORDER BY m.movement_date DESC, m.id DESC"
    rows = db.execute(query, params).fetchall()
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Fecha", "Tipo", "Producto", "Código", "Tipo producto", "Cantidad",
                     "Responsable", "Remito", "Notas", "Registrado"])
    for row in rows:
        writer.writerow([row["movement_date"], row["movement_type"], row["product_name"],
                         row["product_code"], row["product_type"], row["quantity"],
                         row["employee_name"], row["remito"], row["notes"], row["created_at"]])
    parts = []
    if date_from:
        parts.append(date_from)
    if date_to:
        parts.append(date_to)
    suffix = "-".join(parts) or "completo"
    return Response(
        "﻿" + output.getvalue(),
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="movimientos-{suffix}.csv"'},
    )


@app.route("/api/codigos/<path:code>")
def lookup_code(code: str):
    normalized_code = normalize_text(code, 40)
    db = get_db()
    product = db.execute(
        "SELECT id, code, name, product_type, stock, unit FROM products WHERE code = ? AND active = 1",
        (normalized_code,),
    ).fetchone()
    if product is not None:
        return jsonify({"type": "product", "item": dict(product)})
    employee = db.execute(
        "SELECT id, code, full_name, legajo FROM employees WHERE code = ? AND active = 1",
        (normalized_code,),
    ).fetchone()
    if employee is not None:
        return jsonify({"type": "employee", "item": dict(employee)})
    return jsonify({"error": "Código no encontrado o registro inactivo."}), 404


@app.route("/api/dashboard-stats")
def api_dashboard_stats():
    db = get_db()
    return jsonify({
        "products": db.execute("SELECT COUNT(*) FROM products").fetchone()[0],
        "employees": db.execute("SELECT COUNT(*) FROM employees WHERE active = 1").fetchone()[0],
        "low_stock": db.execute(
            "SELECT COUNT(*) FROM products WHERE stock <= minimum_stock"
        ).fetchone()[0],
        "movements": db.execute(
            "SELECT COUNT(*) FROM movements WHERE movement_date = date('now', 'localtime')"
        ).fetchone()[0],
        "pending_purchase_orders": db.execute(
            "SELECT COUNT(*) FROM purchase_orders WHERE status IN ('solicitada', 'parcial')"
        ).fetchone()[0],
        "inventory_value": float(
            db.execute("SELECT COALESCE(SUM(stock * unit_cost), 0) FROM products").fetchone()[0] or 0
        ),
        "tools_delivered": db.execute(
            "SELECT COUNT(*) FROM tool_instances WHERE status = 'entregada'"
        ).fetchone()[0],
        "movements_month": db.execute(
            """SELECT COUNT(*) FROM movements
               WHERE movement_date >= date('now', 'localtime', 'start of month')
                 AND movement_date < date('now', 'localtime', 'start of month', '+1 month')"""
        ).fetchone()[0],
        "pending_returns": db.execute(
            "SELECT COUNT(*) FROM tool_instances WHERE status = 'entregada'"
        ).fetchone()[0],
    })


@app.route("/api/movements-chart")
def api_movements_chart():
    db = get_db()
    rows = db.execute(
        """SELECT movement_date, movement_type, SUM(quantity) AS total
           FROM movements
           WHERE movement_date >= date('now', 'localtime', '-29 days')
           GROUP BY movement_date, movement_type
           ORDER BY movement_date"""
    ).fetchall()
    today = date.today()
    dates = [(today - timedelta(days=i)).isoformat() for i in range(29, -1, -1)]
    buckets: dict[str, dict[str, int]] = {
        d: {"entrada": 0, "salida": 0, "devolucion": 0} for d in dates
    }
    for row in rows:
        if row["movement_date"] in buckets and row["movement_type"] in buckets[row["movement_date"]]:
            buckets[row["movement_date"]][row["movement_type"]] = row["total"]
    labels = [d[5:].replace("-", "/") for d in dates]
    return jsonify({
        "labels": labels,
        "entrada": [buckets[d]["entrada"] for d in dates],
        "salida": [buckets[d]["salida"] for d in dates],
        "devolucion": [buckets[d]["devolucion"] for d in dates],
    })


@app.route("/api/stock-by-category")
def api_stock_by_category():
    db = get_db()
    rows = db.execute(
        """SELECT category, SUM(stock) AS total_stock
           FROM products
           WHERE category <> ''
           GROUP BY category
           ORDER BY total_stock DESC
           LIMIT 8"""
    ).fetchall()
    return jsonify({
        "labels": [row["category"] for row in rows],
        "values": [row["total_stock"] or 0 for row in rows],
    })


@app.route("/api/product-movements-chart/<int:product_id>")
def api_product_movements_chart(product_id: int):
    db = get_db()
    rows = db.execute(
        """SELECT movement_date, movement_type, SUM(quantity) AS total
           FROM movements
           WHERE product_id = ?
             AND movement_date >= date('now', 'localtime', '-59 days')
           GROUP BY movement_date, movement_type
           ORDER BY movement_date""",
        (product_id,)
    ).fetchall()
    today = date.today()
    dates = [(today - timedelta(days=i)).isoformat() for i in range(59, -1, -1)]
    buckets: dict[str, dict[str, int]] = {
        d: {"entrada": 0, "salida": 0, "devolucion": 0} for d in dates
    }
    for row in rows:
        if row["movement_date"] in buckets and row["movement_type"] in buckets[row["movement_date"]]:
            buckets[row["movement_date"]][row["movement_type"]] = row["total"]
    labels = [d[5:].replace("-", "/") for d in dates]
    return jsonify({
        "labels": labels,
        "entrada": [buckets[d]["entrada"] for d in dates],
        "salida": [buckets[d]["salida"] for d in dates],
        "devolucion": [buckets[d]["devolucion"] for d in dates],
    })


@app.route("/movimientos/<int:movement_id>/editar", methods=["GET", "POST"])
def edit_material_movement(movement_id: int):
    db = get_db()
    movement = db.execute(
        """SELECT m.*, p.name AS product_name, p.code AS product_code,
                  p.product_type, e.full_name AS employee_name
           FROM movements m
           JOIN products p ON p.id = m.product_id
           LEFT JOIN employees e ON e.id = m.employee_id
           WHERE m.id = ?""",
        (movement_id,),
    ).fetchone()
    if movement is None:
        flash("El movimiento no existe.", "error")
        return redirect(url_for("movements"))
    if movement["product_type"] != "material":
        flash("Solo se pueden editar responsable y nota de movimientos de materiales.", "error")
        return redirect(url_for("movements"))
    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite corregir movimientos.", "error")
            return redirect(url_for("movements"))
        employee_id_value = request.form.get("employee_id", "").strip()
        try:
            employee_id = int(employee_id_value) if employee_id_value else None
        except ValueError:
            flash("El responsable seleccionado no es válido.", "error")
            return redirect(url_for("edit_material_movement", movement_id=movement_id))
        notes = request.form.get("notes", "").strip()
        if movement["movement_type"] == "salida" and employee_id is None:
            flash("El responsable es obligatorio para entregar materiales.", "error")
            return redirect(url_for("edit_material_movement", movement_id=movement_id))
        if employee_id is not None:
            employee = db.execute(
                "SELECT id FROM employees WHERE id = ? AND active = 1", (employee_id,)
            ).fetchone()
            if employee is None:
                flash("El responsable seleccionado no existe o está inactivo.", "error")
                return redirect(url_for("edit_material_movement", movement_id=movement_id))
        with db:
            db.execute(
                "UPDATE movements SET employee_id = ?, notes = ? WHERE id = ?",
                (employee_id, normalize_text(notes, 500), movement_id),
            )
            audit("update", "movement", movement_id, "Responsable y nota corregidos")
        flash("Responsable y nota del movimiento actualizados.", "success")
        return redirect(url_for("movements"))
    employees_list = db.execute(
        "SELECT id, full_name, legajo FROM employees WHERE active = 1 ORDER BY full_name"
    ).fetchall()
    return render_template(
        "edit_material_movement.html",
        movement=movement,
        employees=employees_list,
    )


@app.route("/movimientos/<int:movement_id>/comprobante")
def movement_receipt(movement_id: int):
    movement = get_db().execute(
        """SELECT m.*, p.name AS product_name, p.code AS product_code,
                  p.brand, p.category, p.presentation, p.unit,
                  e.full_name AS employee_name, e.legajo,
                  ti.serial_number,
                  (SELECT u2.username
                     FROM audit_log a2
                     LEFT JOIN users u2 ON u2.id = a2.user_id
                    WHERE a2.entity_type = 'movement'
                      AND a2.entity_id = m.id
                      AND a2.action = 'create'
                    ORDER BY a2.id ASC
                    LIMIT 1) AS created_by_username
           FROM movements m
           JOIN products p ON p.id = m.product_id
           LEFT JOIN employees e ON e.id = m.employee_id
           LEFT JOIN tool_instances ti ON ti.id = m.tool_instance_id
           WHERE m.id = ?""",
        (movement_id,),
    ).fetchone()
    if movement is None:
        flash("El movimiento no existe.", "error")
        return redirect(url_for("movements"))

    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=f"Comprobante de movimiento #{movement_id}",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReceiptTitle", parent=styles["Title"], alignment=TA_CENTER,
        textColor=colors.HexColor("#0d47a1"), spaceAfter=5,
    )
    subtitle_style = ParagraphStyle(
        "ReceiptSubtitle", parent=styles["Normal"], alignment=TA_CENTER,
        textColor=colors.HexColor("#6b7785"), spaceAfter=18,
    )
    heading_style = ParagraphStyle(
        "ReceiptHeading", parent=styles["Heading2"],
        textColor=colors.HexColor("#1565c0"), spaceBefore=14, spaceAfter=7,
    )
    type_labels = {
        "entrada": "Entrada / ingreso",
        "salida": "Entrega / salida",
        "devolucion": "Devolución",
        "ajuste": "Ajuste",
    }
    product_label = f"{movement['product_name']} ({movement['product_code']})"
    if movement["brand"]:
        product_label += f" - {movement['brand']}"
    details = [
        ["Comprobante", f"#{movement_id}"],
        ["Tipo de movimiento", type_labels.get(movement["movement_type"], movement["movement_type"])],
        ["Fecha del movimiento", movement["movement_date"]],
        ["Fecha de registro", movement["created_at"]],
        ["Producto", product_label],
        ["Categoría", movement["category"]],
        ["Presentación", movement["presentation"]],
        ["Cantidad", f"{movement['quantity']} {movement['unit']}"],
    ]
    if movement["serial_number"]:
        details.append(["Número de serie", movement["serial_number"]])
    details.extend([
        ["Empleado / responsable", movement["employee_name"] or "No corresponde"],
        ["Legajo", movement["legajo"] or "—"],
        ["Observaciones", movement["notes"] or "—"],
    ])
    if movement["remito"]:
        details.append(["N° de remito", movement["remito"]])
    story = [
        Paragraph("PAÑOL | Obra Demo", title_style),
        Paragraph("Comprobante de movimiento de inventario", subtitle_style),
        Paragraph("Detalle de la operación", heading_style),
        make_pdf_table(details, [55 * mm, 105 * mm]),
        Spacer(1, 18),
        Paragraph(
            "Este comprobante documenta el movimiento registrado en el sistema. "
            "Las correcciones posteriores se realizan mediante una operación trazable.",
            styles["Normal"],
        ),
        Spacer(1, 30),
        Table(
            [["Firma entrega / registra", "Firma recibe / devuelve"],
             ["\n\n____________________________", "\n\n____________________________"]],
            colWidths=[80 * mm, 80 * mm],
            style=TableStyle([
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]),
        ),
    ]
    document.build(story, onFirstPage=_pie_pagina, onLaterPages=_pie_pagina)
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"comprobante-movimiento-{movement_id}.pdf",
    )


@app.route("/reportes")
@role_required("admin", "supervisor", "pañolero", "consulta")
def reports():
    db = get_db()
    employee_id = request.args.get("employee_id", type=int)
    category = request.args.get("category", "").strip()
    pending, inventory, categories, employees_list = get_report_data(db, employee_id, category)
    tools = [product for product in inventory if product["product_type"] == "herramienta"]
    materials = [product for product in inventory if product["product_type"] == "material"]
    return render_template(
        "reports.html",
        pending=pending,
        inventory=inventory,
        tools=tools,
        materials=materials,
        categories=categories,
        employees=employees_list,
        employee_id=employee_id,
        category=category,
    )


@app.route("/reportes/valorizacion")
@role_required("admin", "supervisor", "pañolero", "consulta")
def inventory_valuation():
    db = get_db()
    category = request.args.get("category", "").strip()
    product_type = request.args.get("product_type", "").strip()
    products_value = valuation_rows(db, category, product_type)
    categories = db.execute("SELECT DISTINCT category FROM products WHERE active = 1 ORDER BY category").fetchall()
    total_value = sum(float(row["total_value"] or 0) for row in products_value)
    return render_template(
        "valuation.html", products=products_value, categories=categories,
        category=category, product_type=product_type, total_value=total_value,
    )


def valuation_rows(db: sqlite3.Connection, category: str, product_type: str):
    """Return valuation rows using the same validated filters as the HTML report."""
    query = (
        "SELECT name, code, category, product_type, stock, unit, unit_cost, "
        "stock * unit_cost AS total_value FROM products WHERE active = 1"
    )
    params: list[str] = []
    if category:
        query += " AND category = ?"
        params.append(category)
    if product_type in {"material", "herramienta"}:
        query += " AND product_type = ?"
        params.append(product_type)
    query += " ORDER BY category, name"
    return db.execute(query, params).fetchall()


def valuation_download_name(extension: str, category: str, product_type: str) -> str:
    """Build a filesystem-safe download name without trusting query-string values."""
    parts = ["valorizacion-inventario"]
    for value in (category, product_type):
        if value:
            safe = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
            safe = re.sub(r"[^A-Za-z0-9_-]+", "-", safe).strip("-_")
            if safe:
                parts.append(safe[:40])
    return "-".join(parts) + extension


@app.route("/reportes/valorizacion/csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def valuation_csv():
    category = request.args.get("category", "").strip()
    product_type = request.args.get("product_type", "").strip()
    rows = valuation_rows(get_db(), category, product_type)
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Producto", "Código", "Categoría", "Tipo", "Stock", "Unidad",
                     "Costo unitario", "Valor total"])
    for row in rows:
        writer.writerow([
            row["name"], row["code"], row["category"], row["product_type"],
            row["stock"], row["unit"], row["unit_cost"] or 0, row["total_value"] or 0,
        ])
    output = BytesIO(csv_buffer.getvalue().encode("utf-8-sig"))
    return send_file(
        output,
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name=valuation_download_name(".csv", category, product_type),
    )


@app.route("/reportes/valorizacion/excel")
@role_required("admin", "supervisor", "pañolero", "consulta")
def valuation_excel():
    category = request.args.get("category", "").strip()
    product_type = request.args.get("product_type", "").strip()
    rows = valuation_rows(get_db(), category, product_type)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Valorización"
    sheet.append(["Producto", "Código", "Categoría", "Tipo", "Stock", "Unidad",
                  "Costo unitario", "Valor total"])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append([
            row["name"], row["code"], row["category"], row["product_type"],
            row["stock"], row["unit"], row["unit_cost"] or 0, row["total_value"] or 0,
        ])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in sheet.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 35)
        sheet.column_dimensions[column[0].column_letter].width = width
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=valuation_download_name(".xlsx", category, product_type),
    )


def report_export_rows(db: sqlite3.Connection, employee_id: int | None, category: str):
    pending, inventory, _, employees_list = get_report_data(db, employee_id, category)
    return pending, inventory, employees_list


@app.route("/reportes/csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def reports_csv():
    pending, inventory, employees_list = report_export_rows(
        get_db(), request.args.get("employee_id", type=int), request.args.get("category", "").strip()
    )
    output = BytesIO()
    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Sección", "Producto / empleado", "Código", "Categoría / legajo", "Stock / cantidad", "Estado / fecha", "Proveedor"])
    for product in inventory:
        writer.writerow([
            "Inventario", product["name"], product["code"], product["category"],
            f"{product['stock']} {product['unit']}", "Reponer" if product["stock"] <= product["minimum_stock"] else "Disponible",
            product["supplier_id"] or "",
        ])
    for item in pending:
        writer.writerow([
            "Pendientes", f"{item['full_name']} - {item['product_name']}",
            item["product_code"], item["legajo"], item["pending_quantity"], item["last_delivery"], "",
        ])
    for employee in employees_list:
        writer.writerow(["Empleados", employee["full_name"], employee["code"], employee["legajo"], "", "", ""])
    output.write(csv_buffer.getvalue().encode("utf-8-sig"))
    output.seek(0)
    return send_file(output, mimetype="text/csv; charset=utf-8", as_attachment=True,
                     download_name="reporte-panol.csv")


@app.route("/reportes/excel")
@role_required("admin", "supervisor", "pañolero", "consulta")
def reports_excel():
    pending, inventory, employees_list = report_export_rows(
        get_db(), request.args.get("employee_id", type=int), request.args.get("category", "").strip()
    )
    workbook = Workbook()
    sheets = [
        ("Inventario", ["Producto", "Código", "Tipo", "Marca", "Categoría", "Presentación", "Stock", "Unidad", "Mínimo", "Proveedor"]),
        ("Pendientes", ["Empleado", "Legajo", "Herramienta", "Código", "Categoría", "Cantidad pendiente", "Última entrega"]),
        ("Empleados", ["Nombre completo", "Legajo", "Código"]),
    ]
    for index, (title, headers) in enumerate(sheets):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = title
        sheet.append(headers)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
    inventory_sheet, pending_sheet, employees_sheet = workbook.worksheets
    for product in inventory:
        inventory_sheet.append([
            product["name"], product["code"], product["product_type"], product["brand"],
            product["category"], product["presentation"], product["stock"], product["unit"],
            product["minimum_stock"], product["supplier_id"] or "",
        ])
    for item in pending:
        pending_sheet.append([
            item["full_name"], item["legajo"], item["product_name"], item["product_code"],
            item["category"], item["pending_quantity"], item["last_delivery"],
        ])
    for employee in employees_list:
        employees_sheet.append([employee["full_name"], employee["legajo"], employee["code"]])
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            width = min(max(len(str(cell.value or "")) for cell in column) + 2, 35)
            sheet.column_dimensions[column[0].column_letter].width = width
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="reporte-panol.xlsx")


@app.route("/reportes/pendientes")
@role_required("admin", "supervisor", "pañolero", "consulta")
def pending_returns():
    db = get_db()
    employee_id = request.args.get("employee_id", type=int)

    instance_query = """
        SELECT ti.id, ti.serial_number,
               p.name AS product_name, p.code AS product_code, p.category,
               e.full_name, e.legajo, e.code AS employee_code,
               last_m.movement_date AS delivery_date,
               CAST(julianday('now') - julianday(COALESCE(last_m.movement_date, '1900-01-01')) AS INTEGER) AS days_out
        FROM tool_instances ti
        JOIN products p ON p.id = ti.product_id
        LEFT JOIN employees e ON e.id = ti.current_employee_id
        LEFT JOIN (
            SELECT tool_instance_id, MAX(movement_date) AS movement_date
            FROM movements WHERE movement_type = 'salida' GROUP BY tool_instance_id
        ) last_m ON last_m.tool_instance_id = ti.id
        WHERE ti.status = 'entregada'
    """
    instance_params: list[Any] = []
    if employee_id:
        instance_query += " AND e.id = ?"
        instance_params.append(employee_id)
    instance_query += " ORDER BY days_out DESC, e.full_name, p.name"
    instances = db.execute(instance_query, instance_params).fetchall()

    qty_query = """
        SELECT e.id AS employee_id, e.full_name, e.code AS employee_code, e.legajo,
               p.id AS product_id, p.name AS product_name, p.code AS product_code, p.category, p.unit,
               SUM(CASE WHEN m.movement_type = 'salida' THEN m.quantity
                        WHEN m.movement_type = 'devolucion' THEN -m.quantity
                        ELSE 0 END) AS pending_quantity,
               MAX(m.movement_date) AS last_delivery
        FROM movements m
        JOIN employees e ON e.id = m.employee_id
        JOIN products p ON p.id = m.product_id
        WHERE m.movement_type IN ('salida', 'devolucion')
          AND p.product_type = 'herramienta'
          AND m.tool_instance_id IS NULL
    """
    qty_params: list[Any] = []
    if employee_id:
        qty_query += " AND e.id = ?"
        qty_params.append(employee_id)
    qty_query += " GROUP BY e.id, p.id HAVING pending_quantity > 0 ORDER BY e.full_name, p.name"
    qty_pending = db.execute(qty_query, qty_params).fetchall()

    mat_query = """
        SELECT e.id AS employee_id, e.full_name, e.code AS employee_code, e.legajo,
               p.id AS product_id, p.name AS product_name, p.code AS product_code, p.category, p.unit,
               SUM(CASE WHEN m.movement_type = 'salida' THEN m.quantity
                        WHEN m.movement_type = 'devolucion' THEN -m.quantity
                        ELSE 0 END) AS pending_quantity,
               MAX(m.movement_date) AS last_delivery
        FROM movements m
        JOIN employees e ON e.id = m.employee_id
        JOIN products p ON p.id = m.product_id
        WHERE m.movement_type IN ('salida', 'devolucion')
          AND p.product_type = 'material'
          AND p.retornable = 1
    """
    mat_params: list[Any] = []
    if employee_id:
        mat_query += " AND e.id = ?"
        mat_params.append(employee_id)
    mat_query += " GROUP BY e.id, p.id HAVING pending_quantity > 0 ORDER BY e.full_name, p.name"
    mat_pending = db.execute(mat_query, mat_params).fetchall()

    employees_list = db.execute(
        "SELECT id, full_name, legajo FROM employees WHERE active = 1 ORDER BY full_name"
    ).fetchall()
    return render_template(
        "pending_returns.html",
        instances=instances,
        qty_pending=qty_pending,
        mat_pending=mat_pending,
        employees=employees_list,
        employee_id=employee_id,
        today=date.today().isoformat(),
    )


@app.route("/devolucion-rapida", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero")
def devolucion_rapida():
    db = get_db()

    if request.method == "POST":
        if not can_write():
            flash("Su rol no permite registrar movimientos.", "error")
            return redirect(url_for("devolucion_rapida"))

        try:
            employee_id = int(request.form["employee_id"])
        except (KeyError, ValueError):
            flash("Empleado inválido.", "error")
            return redirect(url_for("devolucion_rapida"))

        emp = db.execute(
            "SELECT id, full_name FROM employees WHERE id = ? AND active = 1", (employee_id,)
        ).fetchone()
        if emp is None:
            flash("El empleado no existe o está inactivo.", "error")
            return redirect(url_for("devolucion_rapida"))

        movement_date = request.form.get("movement_date", "").strip()
        try:
            date.fromisoformat(movement_date)
        except ValueError:
            flash("La fecha del movimiento no es válida.", "error")
            return redirect(url_for("devolucion_rapida", employee_id=employee_id))

        notes = normalize_text(request.form.get("notes", ""), 500)
        registered = 0
        errors: list[str] = []

        # --- Herramientas individualizadas ---
        instance_ids = request.form.getlist("instance_ids[]")
        for tid_raw in instance_ids:
            try:
                tid = int(tid_raw)
            except ValueError:
                continue
            instance = db.execute(
                """SELECT ti.*, p.name AS product_name, p.stock, p.id AS product_id
                   FROM tool_instances ti JOIN products p ON p.id = ti.product_id
                   WHERE ti.id = ? AND ti.status = 'entregada' AND ti.current_employee_id = ?""",
                (tid, employee_id),
            ).fetchone()
            if instance is None:
                errors.append(f"Unidad #{tid}: no encontrada o ya devuelta.")
                continue
            with db:
                db.execute(
                    """INSERT INTO movements
                       (product_id, employee_id, tool_instance_id, movement_type, quantity,
                        movement_date, notes)
                       VALUES (?, ?, ?, 'devolucion', 1, ?, ?)""",
                    (instance["product_id"], employee_id, tid, movement_date, notes),
                )
                db.execute(
                    "UPDATE products SET stock = stock + 1 WHERE id = ?",
                    (instance["product_id"],),
                )
                db.execute(
                    "UPDATE tool_instances SET status = 'disponible', current_employee_id = NULL WHERE id = ?",
                    (tid,),
                )
            registered += 1

        # --- Herramientas por cantidad y materiales retornables ---
        qty_pids = request.form.getlist("qty_pid[]")
        qty_qtys = request.form.getlist("qty_qty[]")
        for i, pid_raw in enumerate(qty_pids):
            try:
                pid = int(pid_raw)
                qty = int(qty_qtys[i]) if i < len(qty_qtys) else 0
            except ValueError:
                continue
            if qty <= 0:
                continue
            product = db.execute(
                "SELECT id, name, stock FROM products WHERE id = ?", (pid,)
            ).fetchone()
            if product is None:
                errors.append(f"Producto #{pid}: no encontrado.")
                continue
            with db:
                db.execute(
                    """INSERT INTO movements
                       (product_id, employee_id, movement_type, quantity, movement_date, notes)
                       VALUES (?, ?, 'devolucion', ?, ?, ?)""",
                    (pid, employee_id, qty, movement_date, notes),
                )
                db.execute(
                    "UPDATE products SET stock = stock + ? WHERE id = ?",
                    (qty, pid),
                )
            registered += 1

        for msg in errors:
            flash(msg, "error")
        if registered:
            flash(
                f"Se registraron {registered} devolución(es) para {emp['full_name']}.",
                "success",
            )
        else:
            flash("No se seleccionó ningún ítem para devolver.", "warning")
        return redirect(url_for("devolucion_rapida", employee_id=employee_id))

    # --- GET ---
    employee_id = request.args.get("employee_id", type=int)
    employees_list = db.execute(
        "SELECT id, full_name, legajo FROM employees WHERE active = 1 ORDER BY full_name"
    ).fetchall()

    instances: list[Any] = []
    qty_pending: list[Any] = []
    mat_pending: list[Any] = []

    if employee_id:
        instances = db.execute(
            """SELECT ti.id, ti.serial_number,
                      p.name AS product_name, p.code AS product_code, p.category,
                      last_m.movement_date AS delivery_date,
                      CAST(julianday('now') - julianday(COALESCE(last_m.movement_date,'1900-01-01'))
                           AS INTEGER) AS days_out
               FROM tool_instances ti
               JOIN products p ON p.id = ti.product_id
               LEFT JOIN (
                   SELECT tool_instance_id, MAX(movement_date) AS movement_date
                   FROM movements WHERE movement_type = 'salida' GROUP BY tool_instance_id
               ) last_m ON last_m.tool_instance_id = ti.id
               WHERE ti.status = 'entregada' AND ti.current_employee_id = ?
               ORDER BY days_out DESC, p.name""",
            (employee_id,),
        ).fetchall()

        qty_pending = db.execute(
            """SELECT p.id AS product_id, p.name AS product_name, p.code AS product_code,
                      p.category, p.unit,
                      SUM(CASE WHEN m.movement_type='salida' THEN m.quantity
                               WHEN m.movement_type='devolucion' THEN -m.quantity
                               ELSE 0 END) AS pending_quantity,
                      MAX(m.movement_date) AS last_delivery
               FROM movements m
               JOIN products p ON p.id = m.product_id
               WHERE m.movement_type IN ('salida','devolucion')
                 AND p.product_type = 'herramienta'
                 AND m.tool_instance_id IS NULL
                 AND m.employee_id = ?
               GROUP BY p.id HAVING pending_quantity > 0
               ORDER BY p.name""",
            (employee_id,),
        ).fetchall()

        mat_pending = db.execute(
            """SELECT p.id AS product_id, p.name AS product_name, p.code AS product_code,
                      p.category, p.unit,
                      SUM(CASE WHEN m.movement_type='salida' THEN m.quantity
                               WHEN m.movement_type='devolucion' THEN -m.quantity
                               ELSE 0 END) AS pending_quantity,
                      MAX(m.movement_date) AS last_delivery
               FROM movements m
               JOIN products p ON p.id = m.product_id
               WHERE m.movement_type IN ('salida','devolucion')
                 AND p.product_type = 'material'
                 AND p.retornable = 1
                 AND m.employee_id = ?
               GROUP BY p.id HAVING pending_quantity > 0
               ORDER BY p.name""",
            (employee_id,),
        ).fetchall()

    return render_template(
        "devolucion_rapida.html",
        employees=employees_list,
        employee_id=employee_id,
        instances=instances,
        qty_pending=qty_pending,
        mat_pending=mat_pending,
        today=date.today().isoformat(),
    )


@app.route("/reportes/pendientes/csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def pending_returns_csv():
    db = get_db()
    employee_id = request.args.get("employee_id", type=int)

    instance_query = """
        SELECT ti.serial_number, p.name AS product_name, p.code AS product_code, p.category,
               e.full_name, e.legajo,
               last_m.movement_date AS delivery_date,
               CAST(julianday('now') - julianday(COALESCE(last_m.movement_date, '1900-01-01')) AS INTEGER) AS days_out
        FROM tool_instances ti
        JOIN products p ON p.id = ti.product_id
        LEFT JOIN employees e ON e.id = ti.current_employee_id
        LEFT JOIN (
            SELECT tool_instance_id, MAX(movement_date) AS movement_date
            FROM movements WHERE movement_type = 'salida' GROUP BY tool_instance_id
        ) last_m ON last_m.tool_instance_id = ti.id
        WHERE ti.status = 'entregada'
    """
    qty_query = """
        SELECT e.full_name, e.legajo, p.name AS product_name, p.code AS product_code,
               p.category, p.unit,
               SUM(CASE WHEN m.movement_type = 'salida' THEN m.quantity
                        WHEN m.movement_type = 'devolucion' THEN -m.quantity
                        ELSE 0 END) AS pending_quantity,
               MAX(m.movement_date) AS last_delivery
        FROM movements m
        JOIN employees e ON e.id = m.employee_id
        JOIN products p ON p.id = m.product_id
        WHERE m.movement_type IN ('salida', 'devolucion')
          AND p.product_type = 'herramienta' AND m.tool_instance_id IS NULL
    """
    params: list[Any] = []
    if employee_id:
        instance_query += " AND e.id = ?"
        qty_query += " AND e.id = ?"
        params.append(employee_id)
    instance_query += " ORDER BY days_out DESC, e.full_name, p.name"
    qty_query += " GROUP BY e.id, p.id HAVING pending_quantity > 0 ORDER BY e.full_name, p.name"
    instances = db.execute(instance_query, params).fetchall()
    qty_rows = db.execute(qty_query, params).fetchall()

    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerow(["Tipo", "Empleado", "Legajo", "Herramienta", "Código",
                     "Categoría", "N° serie", "Cantidad pendiente", "Fecha entrega", "Días fuera"])
    for r in instances:
        writer.writerow(["Serie", r["full_name"] or "", r["legajo"] or "",
                         r["product_name"], r["product_code"], r["category"],
                         r["serial_number"], 1, r["delivery_date"] or "", r["days_out"] or ""])
    for r in qty_rows:
        writer.writerow(["Cantidad", r["full_name"], r["legajo"],
                         r["product_name"], r["product_code"], r["category"],
                         "", r["pending_quantity"], r["last_delivery"], ""])
    output = BytesIO(csv_buffer.getvalue().encode("utf-8-sig"))
    today_str = date.today().strftime("%Y-%m-%d")
    return send_file(output, mimetype="text/csv; charset=utf-8", as_attachment=True,
                     download_name=f"pendientes-devolucion-{today_str}.csv")


@app.route("/reportes/catalogo")
@role_required("admin", "supervisor", "pañolero", "consulta")
def catalogo_report():
    db = get_db()
    tipo = request.args.get("tipo", "").strip()
    categorias = request.args.getlist("categoria")
    stock_filter = request.args.get("stock_filter", "").strip()

    query = "SELECT * FROM products WHERE active = 1"
    params: list[Any] = []
    if tipo in {"material", "herramienta"}:
        query += " AND product_type = ?"
        params.append(tipo)
    if categorias:
        placeholders = ",".join("?" * len(categorias))
        query += f" AND category IN ({placeholders})"
        params.extend(categorias)
    if stock_filter == "disponible":
        query += " AND stock > minimum_stock"
    elif stock_filter == "bajo":
        query += " AND stock > 0 AND stock <= minimum_stock"
    elif stock_filter == "sin":
        query += " AND stock = 0"
    query += " ORDER BY category, product_type, name"
    products = db.execute(query, params).fetchall()

    grouped: dict[str, list] = {}
    for p in products:
        grouped.setdefault(p["category"], []).append(p)

    all_categories = db.execute(
        "SELECT DISTINCT category FROM products ORDER BY category"
    ).fetchall()
    total_mat = sum(1 for p in products if p["product_type"] == "material")
    total_her = sum(1 for p in products if p["product_type"] == "herramienta")
    return render_template(
        "catalogo_report.html",
        grouped=grouped,
        all_categories=all_categories,
        tipo=tipo,
        categorias=categorias,
        stock_filter=stock_filter,
        total=len(products),
        total_mat=total_mat,
        total_her=total_her,
    )


@app.route("/reportes/catalogo/csv")
@role_required("admin", "supervisor", "pañolero", "consulta")
def catalogo_report_csv():
    db = get_db()
    tipo = request.args.get("tipo", "").strip()
    categorias = request.args.getlist("categoria")
    stock_filter = request.args.get("stock_filter", "").strip()

    query = "SELECT * FROM products WHERE 1 = 1"
    params: list[Any] = []
    if tipo in {"material", "herramienta"}:
        query += " AND product_type = ?"
        params.append(tipo)
    if categorias:
        placeholders = ",".join("?" * len(categorias))
        query += f" AND category IN ({placeholders})"
        params.extend(categorias)
    if stock_filter == "disponible":
        query += " AND stock > minimum_stock"
    elif stock_filter == "bajo":
        query += " AND stock > 0 AND stock <= minimum_stock"
    elif stock_filter == "sin":
        query += " AND stock = 0"
    query += " ORDER BY category, product_type, name"
    products = db.execute(query, params).fetchall()

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["Categoría", "Tipo", "Código", "Nombre", "Marca",
                     "Presentación", "Unidad", "Stock", "Stock mínimo", "Ubicación"])
    for p in products:
        writer.writerow([
            p["category"], p["product_type"], p["code"], p["name"],
            p["brand"] or "", p["presentation"], p["unit"],
            p["stock"], p["minimum_stock"], p["location"] or "",
        ])
    output = BytesIO(buf.getvalue().encode("utf-8-sig"))
    today_str = date.today().strftime("%Y-%m-%d")
    return send_file(
        output,
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name=f"catalogo-inventario-{today_str}.csv",
    )


@app.route("/reportes/pendientes/asignar-unidad", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def assign_unit_to_pending():
    db = get_db()
    try:
        employee_id = int(request.form["employee_id"])
        product_id = int(request.form["product_id"])
    except (KeyError, ValueError):
        flash("Empleado o herramienta inválidos.", "error")
        return redirect(url_for("pending_returns"))
    serial = normalize_text(request.form.get("serial_number", ""), 80)
    if not serial:
        flash("Debe ingresar un número de unidad.", "error")
        return redirect(url_for("pending_returns"))
    product = db.execute(
        "SELECT * FROM products WHERE id = ? AND product_type = 'herramienta'", (product_id,)
    ).fetchone()
    if not product:
        flash("Herramienta no encontrada.", "error")
        return redirect(url_for("pending_returns"))
    employee = db.execute(
        "SELECT * FROM employees WHERE id = ? AND active = 1", (employee_id,)
    ).fetchone()
    if not employee:
        flash("Empleado no encontrado.", "error")
        return redirect(url_for("pending_returns"))
    existing = db.execute(
        "SELECT * FROM tool_instances WHERE product_id = ? AND serial_number = ?",
        (product_id, serial),
    ).fetchone()
    if existing:
        if existing["status"] == "entregada" and existing["current_employee_id"] != employee_id:
            flash(f"La unidad '{serial}' ya está entregada a otro empleado.", "error")
            return redirect(url_for("pending_returns"))
        if existing["status"] in {"reparacion", "danada", "perdida", "baja"}:
            flash(f"La unidad '{serial}' tiene estado '{existing['status']}' y no puede asignarse.", "error")
            return redirect(url_for("pending_returns"))
    movement = db.execute(
        """SELECT id FROM movements
           WHERE employee_id = ? AND product_id = ? AND movement_type = 'salida'
             AND tool_instance_id IS NULL
           ORDER BY movement_date DESC, id DESC LIMIT 1""",
        (employee_id, product_id),
    ).fetchone()
    if not movement:
        flash("No se encontró un movimiento de salida pendiente para asignar.", "error")
        return redirect(url_for("pending_returns"))
    with db:
        if existing:
            instance_id = existing["id"]
            db.execute(
                "UPDATE tool_instances SET status = 'entregada', current_employee_id = ? WHERE id = ?",
                (employee_id, instance_id),
            )
        else:
            db.execute(
                """INSERT INTO tool_instances (product_id, serial_number, status, current_employee_id)
                   VALUES (?, ?, 'entregada', ?)""",
                (product_id, serial, employee_id),
            )
            instance_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        db.execute(
            "UPDATE movements SET tool_instance_id = ? WHERE id = ?",
            (instance_id, movement["id"]),
        )
        audit(
            "update", "movement", movement["id"],
            f"Unidad '{serial}' asignada retroactivamente a {product['name']} → {employee['full_name']}",
        )
    flash(f"Unidad '{serial}' asignada correctamente a {employee['full_name']} — {product['name']}.", "success")
    return redirect(url_for("pending_returns"))


def get_report_data(
    db: sqlite3.Connection, employee_id: int | None, category: str
) -> tuple[list[sqlite3.Row], list[sqlite3.Row], list[sqlite3.Row], list[sqlite3.Row]]:
    """Get the report datasets shared by HTML and PDF exports."""

    pending_query = """
        SELECT e.id AS employee_id, e.full_name, e.code AS employee_code,
               e.legajo, p.name AS product_name, p.code AS product_code,
               p.category, p.presentation, p.unit,
               SUM(CASE WHEN m.movement_type = 'salida' THEN m.quantity
                        WHEN m.movement_type = 'devolucion' THEN -m.quantity
                        ELSE 0 END) AS pending_quantity,
               MAX(m.movement_date) AS last_delivery
        FROM movements m
        JOIN employees e ON e.id = m.employee_id
        JOIN products p ON p.id = m.product_id
        WHERE m.movement_type IN ('salida', 'devolucion')
          AND p.product_type = 'herramienta'
    """
    pending_params: list[Any] = []
    if employee_id:
        pending_query += " AND e.id = ?"
        pending_params.append(employee_id)
    if category:
        pending_query += " AND p.category = ?"
        pending_params.append(category)
    pending_query += """
        GROUP BY e.id, p.id
        HAVING pending_quantity > 0
        ORDER BY e.full_name, p.name
    """
    pending = db.execute(pending_query, pending_params).fetchall()

    inventory_query = "SELECT * FROM products WHERE 1 = 1"
    inventory_params: list[str] = []
    if category:
        inventory_query += " AND category = ?"
        inventory_params.append(category)
    inventory_query += " ORDER BY category, name"
    inventory = db.execute(inventory_query, inventory_params).fetchall()
    categories = db.execute("SELECT DISTINCT category FROM products ORDER BY category").fetchall()
    employees_list = db.execute(
        "SELECT id, full_name, legajo, code FROM employees WHERE active = 1 ORDER BY full_name"
    ).fetchall()
    return pending, inventory, categories, employees_list


@app.route("/reportes/pdf")
@role_required("admin", "supervisor", "pañolero", "consulta")
def reports_pdf():
    db = get_db()
    employee_id = request.args.get("employee_id", type=int)
    category = request.args.get("category", "").strip()
    pending, inventory, _, employees_list = get_report_data(db, employee_id, category)
    tools = [product for product in inventory if product["product_type"] == "herramienta"]
    materials = [product for product in inventory if product["product_type"] == "material"]

    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Reporte de pañol - Obra Demo",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle", parent=styles["Title"], alignment=TA_CENTER,
        textColor=colors.HexColor("#0d47a1"), spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "ReportSubtitle", parent=styles["Normal"], alignment=TA_CENTER,
        textColor=colors.HexColor("#6b7785"), spaceAfter=14,
    )
    heading_style = ParagraphStyle(
        "ReportHeading", parent=styles["Heading2"], textColor=colors.HexColor("#1565c0"),
        spaceBefore=10, spaceAfter=6,
    )
    story = [
        Paragraph("PAÑOL | Obra Demo", title_style),
        Paragraph(
            f"Reporte A4 generado el {date.today().strftime('%d/%m/%Y')}",
            subtitle_style,
        ),
        Paragraph("Listado de herramientas", heading_style),
    ]
    tool_rows = [["Herramienta", "Marca", "Categoría", "Presentación", "Ubicación", "Stock", "Mínimo", "Estado"]]
    tool_rows.extend(
        [
            [
                f"{product['name']} ({product['code']})",
                product["brand"] or "—",
                product["category"],
                product["presentation"],
                product["location"] or "Sin ubicación",
                f"{product['stock']} {product['unit']}",
                str(product["minimum_stock"]),
                "Reponer" if product["stock"] <= product["minimum_stock"] else "Disponible",
            ]
            for product in tools
        ]
    )
    if len(tool_rows) == 1:
        tool_rows.append(["No hay herramientas registradas.", "", "", "", "", "", "", ""])
    story.append(make_pdf_table(tool_rows, [48 * mm, 30 * mm, 34 * mm, 38 * mm, 38 * mm, 23 * mm, 20 * mm, 28 * mm]))
    story.extend([Spacer(1, 8), Paragraph("Herramientas pendientes de devolución", heading_style)])
    pending_rows = [["Empleado / legajo", "Herramienta", "Categoría", "Cantidad", "Última entrega"]]
    pending_rows.extend(
        [
            [
                f"{item['full_name']} / {item['legajo']}",
                f"{item['product_name']} ({item['product_code']})",
                item["category"],
                f"{item['pending_quantity']} {item['unit']}",
                item["last_delivery"],
            ]
            for item in pending
        ]
    )
    if len(pending_rows) == 1:
        pending_rows.append(["No hay herramientas pendientes.", "", "", "", ""])
    story.append(make_pdf_table(pending_rows, [55 * mm, 65 * mm, 42 * mm, 30 * mm, 42 * mm]))
    story.extend([Spacer(1, 8), Paragraph("Listado de materiales", heading_style)])
    material_rows = [["Material", "Marca", "Categoría", "Presentación", "Ubicación", "Stock", "Mínimo", "Estado"]]
    material_rows.extend(
        [
            [
                f"{product['name']} ({product['code']})",
                product["brand"] or "—",
                product["category"],
                product["presentation"],
                product["location"] or "Sin ubicación",
                f"{product['stock']} {product['unit']}",
                str(product["minimum_stock"]),
                "Reponer" if product["stock"] <= product["minimum_stock"] else "Disponible",
            ]
            for product in materials
        ]
    )
    if len(material_rows) == 1:
        material_rows.append(["No hay materiales registrados.", "", "", "", "", "", "", ""])
    story.append(make_pdf_table(material_rows, [48 * mm, 30 * mm, 34 * mm, 38 * mm, 38 * mm, 23 * mm, 20 * mm, 28 * mm]))
    story.extend([Spacer(1, 8), Paragraph("Listado de empleados", heading_style)])
    employee_rows = [["Nombre completo", "Legajo", "Código"]]
    employee_rows.extend(
        [[employee["full_name"], employee["legajo"], employee["code"]] for employee in employees_list]
    )
    if len(employee_rows) == 1:
        employee_rows.append(["No hay empleados activos registrados.", "", ""])
    story.append(make_pdf_table(employee_rows, [95 * mm, 45 * mm, 45 * mm]))
    document.build(story, onFirstPage=_pie_pagina, onLaterPages=_pie_pagina)
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="reporte-panol.pdf",
    )


@app.route("/reportes/pdf/materiales")
@role_required("admin", "supervisor", "pañolero", "consulta")
def reports_pdf_materiales():
    db = get_db()
    category = request.args.get("category", "").strip()
    query = "SELECT * FROM products WHERE active = 1 AND product_type = 'material'"
    params: list[str] = []
    if category:
        query += " AND category = ?"
        params.append(category)
    query += " ORDER BY category, name"
    materials = db.execute(query, params).fetchall()

    buffer = BytesIO()
    filter_text = f" — Categoría: {category}" if category else ""
    document = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=12 * mm, leftMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title="Stock de materiales - Pañol Demo",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "MatTitle", parent=styles["Title"], alignment=TA_CENTER,
        textColor=colors.HexColor("#0d47a1"), spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "MatSubtitle", parent=styles["Normal"], alignment=TA_CENTER,
        textColor=colors.HexColor("#6b7785"), spaceAfter=14,
    )
    heading_style = ParagraphStyle(
        "MatHeading", parent=styles["Heading2"], textColor=colors.HexColor("#1565c0"),
        spaceBefore=10, spaceAfter=6,
    )
    story = [
        Paragraph("PAÑOL | Obra Demo", title_style),
        Paragraph(
            f"Stock de materiales{filter_text} — {date.today().strftime('%d/%m/%Y')}",
            subtitle_style,
        ),
        Paragraph("Stock de materiales", heading_style),
    ]
    rows = [["Material", "Marca", "Categoría", "Presentación", "Ubicación", "Stock", "Mínimo", "Estado"]]
    rows.extend([
        [
            f"{p['name']} ({p['code']})",
            p["brand"] or "—",
            p["category"],
            p["presentation"],
            p["location"] or "Sin ubicación",
            f"{p['stock']} {p['unit']}",
            str(p["minimum_stock"]),
            "Reponer" if p["stock"] <= p["minimum_stock"] else "Disponible",
        ]
        for p in materials
    ])
    if len(rows) == 1:
        rows.append(["No hay materiales registrados.", "", "", "", "", "", "", ""])
    story.append(make_pdf_table(rows, [48 * mm, 30 * mm, 34 * mm, 38 * mm, 38 * mm, 23 * mm, 20 * mm, 28 * mm]))
    document.build(story, onFirstPage=_pie_pagina, onLaterPages=_pie_pagina)
    buffer.seek(0)
    return send_file(
        buffer, mimetype="application/pdf", as_attachment=True,
        download_name="stock-materiales-demo.pdf",
    )


@app.route("/reportes/pdf/herramientas")
@role_required("admin", "supervisor", "pañolero", "consulta")
def reports_pdf_herramientas():
    db = get_db()
    category = request.args.get("category", "").strip()
    query = "SELECT * FROM products WHERE active = 1 AND product_type = 'herramienta'"
    params: list[str] = []
    if category:
        query += " AND category = ?"
        params.append(category)
    query += " ORDER BY category, name"
    tools = db.execute(query, params).fetchall()

    buffer = BytesIO()
    filter_text = f" — Categoría: {category}" if category else ""
    document = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=12 * mm, leftMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
        title="Stock de herramientas - Pañol Demo",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "HerTitle", parent=styles["Title"], alignment=TA_CENTER,
        textColor=colors.HexColor("#0d47a1"), spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "HerSubtitle", parent=styles["Normal"], alignment=TA_CENTER,
        textColor=colors.HexColor("#6b7785"), spaceAfter=14,
    )
    heading_style = ParagraphStyle(
        "HerHeading", parent=styles["Heading2"], textColor=colors.HexColor("#1565c0"),
        spaceBefore=10, spaceAfter=6,
    )
    story = [
        Paragraph("PAÑOL | Obra Demo", title_style),
        Paragraph(
            f"Stock de herramientas{filter_text} — {date.today().strftime('%d/%m/%Y')}",
            subtitle_style,
        ),
        Paragraph("Stock de herramientas", heading_style),
    ]
    rows = [["Herramienta", "Marca", "Categoría", "Presentación", "Ubicación", "Stock", "Mínimo", "Estado"]]
    rows.extend([
        [
            f"{p['name']} ({p['code']})",
            p["brand"] or "—",
            p["category"],
            p["presentation"],
            p["location"] or "Sin ubicación",
            f"{p['stock']} {p['unit']}",
            str(p["minimum_stock"]),
            "Reponer" if p["stock"] <= p["minimum_stock"] else "Disponible",
        ]
        for p in tools
    ])
    if len(rows) == 1:
        rows.append(["No hay herramientas registradas.", "", "", "", "", "", "", ""])
    story.append(make_pdf_table(rows, [48 * mm, 30 * mm, 34 * mm, 38 * mm, 38 * mm, 23 * mm, 20 * mm, 28 * mm]))
    document.build(story, onFirstPage=_pie_pagina, onLaterPages=_pie_pagina)
    buffer.seek(0)
    return send_file(
        buffer, mimetype="application/pdf", as_attachment=True,
        download_name="stock-herramientas-demo.pdf",
    )


@app.route("/movimientos/<int:movement_id>/revertir", methods=["POST"])
@role_required("admin", "supervisor")
def reverse_movement(movement_id: int):
    db = get_db()
    movement = db.execute(
        """SELECT m.*, p.code AS product_code, p.stock
           FROM movements m JOIN products p ON p.id = m.product_id
           WHERE m.id = ?""",
        (movement_id,),
    ).fetchone()
    if movement is None:
        flash("El movimiento no existe.", "error")
        return redirect(url_for("movements"))
    if db.execute(
        "SELECT 1 FROM movement_reversals WHERE original_movement_id = ?",
        (movement_id,),
    ).fetchone():
        flash("El movimiento ya fue revertido.", "error")
        return redirect(url_for("movements"))
    reason = normalize_text(request.form.get("reason", ""), 500)
    if not reason:
        flash("Debe indicar el motivo de la reversión.", "error")
        return redirect(url_for("movements"))
    reverse_type = {
        "entrada": "salida",
        "salida": "devolucion",
        "devolucion": "salida",
        "ajuste": "entrada",
    }[movement["movement_type"]]
    new_stock = movement["stock"] + (
        movement["quantity"] if reverse_type in {"entrada", "devolucion"} else -movement["quantity"]
    )
    if new_stock < 0:
        flash("No se puede revertir: el stock actual no permite deshacer este movimiento.", "error")
        return redirect(url_for("movements"))
    with db:
        cursor = db.execute(
            """INSERT INTO movements
               (product_id, employee_id, tool_instance_id, movement_type, quantity, notes, movement_date)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (
                movement["product_id"],
                movement["employee_id"],
                movement["tool_instance_id"],
                reverse_type,
                movement["quantity"],
                f"Reversión del movimiento #{movement_id}: {reason}",
                date.today().isoformat(),
            ),
        )
        reversal_id = cursor.lastrowid
        db.execute("UPDATE products SET stock = ? WHERE id = ?", (new_stock, movement["product_id"]))
        if movement["tool_instance_id"] is not None:
            instance = db.execute(
                "SELECT id, serial_number, status, current_employee_id FROM tool_instances WHERE id = ?",
                (movement["tool_instance_id"],),
            ).fetchone()
            if instance is not None:
                next_status = "entregada" if reverse_type == "salida" else "disponible"
                next_employee = movement["employee_id"] if reverse_type == "salida" else None
                db.execute(
                    "UPDATE tool_instances SET status = ?, current_employee_id = ? WHERE id = ?",
                    (next_status, next_employee, movement["tool_instance_id"]),
                )
                db.execute(
                    """INSERT INTO tool_instance_status_history
                       (instance_id, previous_status, new_status, employee_id, notes, changed_by)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (
                        movement["tool_instance_id"],
                        instance["status"],
                        next_status,
                        next_employee,
                        f"Reversión de movimiento #{movement_id}: {reason}",
                        g.user["id"],
                    ),
                )
        db.execute(
            """INSERT INTO movement_reversals
               (original_movement_id, reversal_movement_id, user_id, reason)
               VALUES (?, ?, ?, ?)""",
            (movement_id, reversal_id, g.user["id"], reason),
        )
        audit(
            "reverse",
            "movement",
            movement_id,
            f"Revertido con movimiento #{reversal_id}. Motivo: {reason}",
        )
    flash("Movimiento revertido y stock actualizado.", "success")
    return redirect(url_for("movements"))


@app.route("/usuarios", methods=["GET", "POST"])
@role_required("admin")
def users():
    db = get_db()
    if request.method == "POST":
        username = normalize_text(request.form.get("username", ""), 40)
        password = request.form.get("password", "")
        role = request.form.get("role", "")
        if not username or not IDENTIFIER_PATTERN.fullmatch(username):
            flash("El usuario debe tener entre 2 y 40 caracteres válidos.", "error")
        elif len(password) < 8:
            flash("La contraseña debe tener al menos 8 caracteres.", "error")
        elif role not in {"admin", "supervisor", "pañolero", "consulta"}:
            flash("El rol seleccionado no es válido.", "error")
        else:
            try:
                with db:
                    cursor = db.execute(
                        "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                        (username, generate_password_hash(password), role),
                    )
                    audit("create", "user", cursor.lastrowid, f"Usuario {username} creado con rol {role}")
                flash("Usuario creado correctamente.", "success")
            except sqlite3.IntegrityError:
                flash("El nombre de usuario ya existe.", "error")
        return redirect(url_for("users"))
    users_list = db.execute(
        "SELECT id, username, role, active, created_at FROM users ORDER BY username"
    ).fetchall()
    return render_template("users.html", users=users_list)


@app.route("/proveedores", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero")
def suppliers():
    db = get_db()
    if request.method == "POST":
        name = normalize_text(request.form.get("name", ""), 120)
        contact = normalize_text(request.form.get("contact", ""), 120)
        phone = normalize_text(request.form.get("phone", ""), 40)
        email = normalize_text(request.form.get("email", ""), 120)
        if not name:
            flash("El nombre del proveedor es obligatorio.", "error")
            return redirect(url_for("suppliers"))
        if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
            flash("El formato del email no es válido.", "error")
            return redirect(url_for("suppliers"))
        try:
            with db:
                cursor = db.execute(
                    "INSERT INTO suppliers (name, contact, phone, email) VALUES (?, ?, ?, ?)",
                    (name, contact, phone, email),
                )
                audit("create", "supplier", cursor.lastrowid, f"Proveedor {name} registrado")
            flash("Proveedor registrado correctamente.", "success")
        except sqlite3.IntegrityError:
            flash("Ya existe un proveedor con ese nombre.", "error")
        return redirect(url_for("suppliers"))
    search = request.args.get("q", "").strip()
    status = request.args.get("status", "").strip().lower()
    if status not in {"active", "inactive"}:
        status = ""
    page, page_size, offset = pagination_params()

    conditions = []
    params = []
    if search:
        pattern = f"%{search}%"
        conditions.append("(name LIKE ? OR contact LIKE ? OR phone LIKE ? OR email LIKE ?)")
        params.extend([pattern] * 4)
    if status:
        conditions.append("active = ?")
        params.append(1 if status == "active" else 0)
    where = f" WHERE {' AND '.join(conditions)}" if conditions else ""
    total = db.execute(f"SELECT COUNT(*) FROM suppliers{where}", params).fetchone()[0]
    rows = db.execute(
        f"SELECT * FROM suppliers{where} ORDER BY name LIMIT ? OFFSET ?",
        [*params, page_size, offset],
    ).fetchall()
    return render_template(
        "suppliers.html",
        suppliers=rows,
        search=search,
        status=status,
        pagination=pagination_context(page, page_size, total),
    )


PURCHASE_ORDER_STATUSES = {"borrador", "solicitada", "parcial", "recibida", "cancelada"}


def _purchase_order(order_id: int):
    return get_db().execute(
        """SELECT po.*, s.name AS supplier_name, u.username AS creator
           FROM purchase_orders po JOIN suppliers s ON s.id = po.supplier_id
           JOIN users u ON u.id = po.created_by WHERE po.id = ?""", (order_id,)
    ).fetchone()


@app.route("/ordenes-compra", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero", "consulta")
def purchase_orders():
    db = get_db()
    if request.method == "POST":
        if not can_write():
            flash("Su rol solo permite visualizar órdenes de compra.", "error")
            return redirect(url_for("purchase_orders"))
        try:
            supplier_id = int(request.form["supplier_id"])
        except (KeyError, ValueError):
            flash("Seleccione un proveedor válido.", "error")
            return redirect(url_for("purchase_orders"))
        supplier = db.execute("SELECT id FROM suppliers WHERE id = ? AND active = 1", (supplier_id,)).fetchone()
        notes = normalize_text(request.form.get("notes", ""), 1000)
        if supplier is None:
            flash("El proveedor no existe o está inactivo.", "error")
            return redirect(url_for("purchase_orders"))
        order_date = request.form.get("order_date", date.today().isoformat())
        try:
            date.fromisoformat(order_date)
        except ValueError:
            flash("La fecha de la orden no es válida.", "error")
            return redirect(url_for("purchase_orders"))
        with db:
            sequence = db.execute(
                "SELECT COUNT(*) FROM purchase_orders WHERE order_number LIKE ?",
                (f"OC-{order_date.replace('-', '')}-%",),
            ).fetchone()[0] + 1
            order_number = f"OC-{order_date.replace('-', '')}-{sequence:03d}"
            cursor = db.execute(
                """INSERT INTO purchase_orders
                   (order_number, supplier_id, order_date, notes, created_by)
                   VALUES (?, ?, ?, ?, ?)""",
                (order_number, supplier_id, order_date, notes, g.user["id"]),
            )
            audit("create", "purchase_order", cursor.lastrowid, f"Orden {order_number} creada")
        flash("Orden de compra creada como borrador.", "success")
        return redirect(url_for("purchase_order_detail", order_id=cursor.lastrowid))
    status = request.args.get("status", "").strip()
    search = request.args.get("q", "").strip()
    supplier_id_filter = request.args.get("supplier_id", "").strip()
    query = """SELECT po.*, s.name AS supplier_name
               FROM purchase_orders po JOIN suppliers s ON s.id = po.supplier_id"""
    params = []
    conditions = []
    if status in PURCHASE_ORDER_STATUSES:
        conditions.append("po.status = ?")
        params.append(status)
    else:
        status = ""
    try:
        supplier_id = int(supplier_id_filter)
    except (TypeError, ValueError):
        supplier_id = None
    if supplier_id is not None and supplier_id > 0:
        conditions.append("po.supplier_id = ?")
        params.append(supplier_id)
    else:
        supplier_id_filter = ""
    if search:
        conditions.append("(po.order_number LIKE ? OR po.notes LIKE ?)")
        params.extend((f"%{search}%", f"%{search}%"))
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    page, page_size, offset = pagination_params()
    total = db.execute(f"SELECT COUNT(*) FROM ({query})", params).fetchone()[0]
    rows = db.execute(
        query + " ORDER BY po.order_date DESC, po.id DESC LIMIT ? OFFSET ?",
        [*params, page_size, offset],
    ).fetchall()
    suppliers_list = db.execute(
        "SELECT id, name FROM suppliers WHERE active = 1 ORDER BY name"
    ).fetchall()
    filter_suppliers = db.execute("SELECT id, name FROM suppliers ORDER BY name").fetchall()
    return render_template(
        "purchase_orders.html",
        orders=rows,
        suppliers=suppliers_list,
        filter_suppliers=filter_suppliers,
        status=status,
        supplier_id_filter=supplier_id_filter,
        search=search,
        pagination=pagination_context(page, page_size, total),
        today=date.today().isoformat(),
    )


@app.route("/ordenes-compra/nueva")
@role_required("admin", "supervisor", "pañolero")
def purchase_order_new():
    return redirect(url_for("purchase_orders"))


@app.route("/ordenes-compra/<int:order_id>")
@role_required("admin", "supervisor", "pañolero", "consulta")
def purchase_order_detail(order_id: int):
    order = _purchase_order(order_id)
    if order is None:
        return "Orden no encontrada", 404
    items = get_db().execute(
        """SELECT i.*, p.code, p.name, p.unit FROM purchase_order_items i
           JOIN products p ON p.id = i.product_id
           WHERE i.purchase_order_id = ? ORDER BY p.name""", (order_id,)
    ).fetchall()
    products_list = get_db().execute("SELECT id, code, name, unit FROM products ORDER BY name").fetchall()
    receipt_token = session.setdefault(
        f"purchase_receipt_token_{order_id}", secrets.token_urlsafe(24)
    )
    return render_template(
        "purchase_order_detail.html",
        order=order,
        items=items,
        products=products_list,
        receipt_token=receipt_token,
    )


@app.route("/ordenes-compra/<int:order_id>/items", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def purchase_order_add_item(order_id: int):
    db = get_db()
    order = _purchase_order(order_id)
    if order is None:
        return "Orden no encontrada", 404
    if order["status"] != "borrador":
        flash("Solo se pueden agregar ítems en borrador.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    try:
        product_id = int(request.form["product_id"])
        quantity = int(request.form["quantity_requested"])
        cost_text = request.form.get("unit_cost", "").strip()
        unit_cost = parse_unit_cost(cost_text)
    except (KeyError, ValueError):
        flash("Producto, cantidad y costo deben ser válidos.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    product = db.execute("SELECT id FROM products WHERE id = ?", (product_id,)).fetchone()
    if product is None or quantity <= 0:
        flash("Producto, cantidad o costo inválidos.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    try:
        with db:
            db.execute(
                """INSERT INTO purchase_order_items
                   (purchase_order_id, product_id, quantity_requested, unit_cost)
                   VALUES (?, ?, ?, ?)                   """, (order_id, product_id, quantity, cost_for_db(unit_cost))
            )
            audit("create", "purchase_order_item", order_id, f"Ítem agregado a orden {order['order_number']}")
        flash("Ítem agregado.", "success")
    except sqlite3.IntegrityError:
        flash("El producto ya está incluido en la orden.", "error")
    return redirect(url_for("purchase_order_detail", order_id=order_id))


@app.route("/ordenes-compra/<int:order_id>/items/<int:item_id>/editar", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def purchase_order_edit_item(order_id: int, item_id: int):
    db = get_db()
    order = _purchase_order(order_id)
    try:
        quantity = int(request.form["quantity_requested"])
        cost_text = request.form.get("unit_cost", "").strip()
        unit_cost = parse_unit_cost(cost_text)
    except (KeyError, ValueError):
        flash("Cantidad o costo inválidos.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    item = db.execute("SELECT * FROM purchase_order_items WHERE id = ? AND purchase_order_id = ?", (item_id, order_id)).fetchone()
    if order is None or item is None or order["status"] != "borrador" or quantity <= 0:
        flash("El ítem no se puede editar.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    with db:
        db.execute("UPDATE purchase_order_items SET quantity_requested = ?, unit_cost = ? WHERE id = ?", (quantity, cost_for_db(unit_cost), item_id))
        audit("update", "purchase_order_item", item_id, f"Ítem de orden {order['order_number']} modificado")
    flash("Ítem actualizado.", "success")
    return redirect(url_for("purchase_order_detail", order_id=order_id))


@app.route("/ordenes-compra/<int:order_id>/solicitar", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def purchase_order_request(order_id: int):
    db = get_db()
    order = _purchase_order(order_id)
    items = db.execute("SELECT COUNT(*) FROM purchase_order_items WHERE purchase_order_id = ?", (order_id,)).fetchone()[0] if order else 0
    if order is None:
        return "Orden no encontrada", 404
    if order["status"] != "borrador" or not items:
        flash("La orden debe ser un borrador con al menos un ítem.", "error")
    else:
        with db:
            db.execute("UPDATE purchase_orders SET status = 'solicitada', requested_at = CURRENT_TIMESTAMP WHERE id = ?", (order_id,))
            audit("update", "purchase_order", order_id, f"Orden {order['order_number']} solicitada")
        flash("Orden solicitada.", "success")
    return redirect(url_for("purchase_order_detail", order_id=order_id))


@app.route("/ordenes-compra/<int:order_id>/cancelar", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def purchase_order_cancel(order_id: int):
    db = get_db()
    order = _purchase_order(order_id)
    if order is None:
        return "Orden no encontrada", 404
    if order["status"] in {"recibida", "cancelada"}:
        flash("La orden ya no se puede cancelar.", "error")
    else:
        with db:
            db.execute("UPDATE purchase_orders SET status = 'cancelada', cancelled_at = CURRENT_TIMESTAMP WHERE id = ?", (order_id,))
            audit("update", "purchase_order", order_id, f"Orden {order['order_number']} cancelada")
        flash("Orden cancelada.", "success")
    return redirect(url_for("purchase_order_detail", order_id=order_id))


@app.route("/ordenes-compra/<int:order_id>/recibir", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def purchase_order_receive(order_id: int):
    db = get_db()
    order = _purchase_order(order_id)
    if order is None:
        return "Orden no encontrada", 404
    if order["status"] not in {"solicitada", "parcial"}:
        flash("No se puede recibir una orden cancelada, en borrador o totalmente recibida.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    items = db.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = ?", (order_id,)).fetchall()
    received = {}
    try:
        for item in items:
            raw = request.form.get(
                f"received_{item['id']}",
                request.form.get(f"quantity_received_{item['id']}", request.form.get(f"quantity_{item['id']}", "0")),
            )
            received[item["id"]] = int(raw or 0)
            if received[item["id"]] < 0 or received[item["id"]] > item["quantity_requested"] - item["quantity_received"]:
                raise ValueError
    except ValueError:
        flash("La recepción debe ser un entero no negativo y no superar lo pendiente.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    if not any(received.values()):
        flash("Indique al menos una cantidad recibida.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    movement_date = request.form.get("movement_date", "").strip() or date.today().isoformat()
    try:
        movement_date = date.fromisoformat(movement_date).isoformat()
    except ValueError:
        flash("La fecha de recepción no es válida.", "error")
        return redirect(url_for("purchase_order_detail", order_id=order_id))
    receipt_token = request.form.get("receipt_token", "").strip()
    if receipt_token:
        receipt_key = f"token:{receipt_token}"
    else:
        payload = "|".join(
            f"{item_id}:{received[item_id]}" for item_id in sorted(received)
        )
        receipt_key = "legacy:" + hashlib.sha256(
            f"{order_id}|{movement_date}|{payload}".encode()
        ).hexdigest()
    with db:
        try:
            db.execute(
                """INSERT INTO purchase_receipts
                   (purchase_order_id, receipt_key, movement_date)
                   VALUES (?, ?, ?)""",
                (order_id, receipt_key, movement_date),
            )
        except sqlite3.IntegrityError:
            flash("La recepción ya fue registrada; se ignoró el reenvío.", "error")
            return redirect(url_for("purchase_order_detail", order_id=order_id))
        for item in items:
            qty = received[item["id"]]
            if not qty:
                continue
            db.execute(
                """INSERT INTO movements
                   (product_id, movement_type, quantity, notes, movement_date)
                   VALUES (?, 'entrada', ?, ?, ?)""",
                (item["product_id"], qty, f"Recepción orden {order['order_number']}", movement_date),
            )
            db.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (qty, item["product_id"]))
            if item["unit_cost"] is not None:
                cost = parse_unit_cost(str(item["unit_cost"]), allow_empty=False)
                product = db.execute(
                    "SELECT unit_cost FROM products WHERE id = ?", (item["product_id"],)
                ).fetchone()
                old_cost = parse_unit_cost(str(product["unit_cost"]), allow_empty=False)
                if old_cost != cost:
                    db.execute(
                        "UPDATE products SET unit_cost = ? WHERE id = ?",
                        (cost_for_db(cost), item["product_id"]),
                    )
                    audit(
                        "update",
                        "product",
                        item["product_id"],
                        f"Costo unitario actualizado de {old_cost} a {cost} por recepción de {order['order_number']}",
                    )
            db.execute("UPDATE purchase_order_items SET quantity_received = quantity_received + ? WHERE id = ?", (qty, item["id"]))
            audit("create", "movement", db.execute("SELECT last_insert_rowid()").fetchone()[0], f"Entrada por orden {order['order_number']}")
        remaining = db.execute(
            "SELECT COUNT(*) FROM purchase_order_items WHERE purchase_order_id = ? AND quantity_received < quantity_requested", (order_id,)
        ).fetchone()[0]
        new_status = "parcial" if remaining else "recibida"
        db.execute("UPDATE purchase_orders SET status = ?, received_at = CASE WHEN ? = 'recibida' THEN CURRENT_TIMESTAMP ELSE received_at END WHERE id = ?", (new_status, new_status, order_id))
        audit("update", "purchase_order", order_id, f"Recepción de orden {order['order_number']}: {new_status}")
    session.pop(f"purchase_receipt_token_{order_id}", None)
    flash("Recepción registrada y stock actualizado.", "success")
    return redirect(url_for("purchase_order_detail", order_id=order_id))


@app.route("/reposicion")
@role_required("admin", "supervisor", "pañolero", "consulta")
def replenishment():
    db = get_db()
    search = request.args.get("q", "").strip()
    category = request.args.get("category", "").strip()
    supplier = request.args.get("supplier", request.args.get("supplier_id", "")).strip()
    product_type = request.args.get("product_type", "").strip()
    sort = request.args.get("sort", "").strip()

    where = ["p.active = 1", "p.stock <= p.minimum_stock"]
    params: list[Any] = []
    if search:
        where.append("(p.name LIKE ? OR p.code LIKE ?)")
        params.extend((f"%{search}%", f"%{search}%"))
    if category:
        where.append("p.category = ?")
        params.append(category)
    if supplier:
        try:
            supplier_id = int(supplier)
        except ValueError:
            supplier_id = None
        if supplier_id is None:
            where.append("1 = 0")
        else:
            where.append("p.supplier_id = ?")
            params.append(supplier_id)
    if product_type:
        where.append("p.product_type = ?")
        params.append(product_type)
    where_sql = " AND ".join(where)
    from_sql = "FROM products p LEFT JOIN suppliers s ON s.id = p.supplier_id"
    total = db.execute(
        f"SELECT COUNT(*) {from_sql} WHERE {where_sql}", params
    ).fetchone()[0]

    page, page_size, offset = pagination_params()
    order_by = {
        "faltante": "p.minimum_stock - p.stock DESC, p.name, p.id",
        "name": "p.name, p.id",
    }.get(sort, "CASE WHEN s.name IS NULL THEN 1 ELSE 0 END, p.name, p.id")
    rows = db.execute(
        f"""SELECT p.*, s.name AS supplier_name, s.contact, s.phone, s.email,
                   CASE WHEN p.reorder_quantity > 0 THEN p.reorder_quantity
                        ELSE MAX(p.minimum_stock - p.stock, 1) END AS suggested_quantity
            {from_sql}
            WHERE {where_sql}
            ORDER BY {order_by}
            LIMIT ? OFFSET ?""",
        [*params, page_size, offset],
    ).fetchall()
    categories = db.execute(
        "SELECT DISTINCT category FROM products WHERE category <> '' ORDER BY category"
    ).fetchall()
    suppliers = db.execute("SELECT id, name FROM suppliers ORDER BY name").fetchall()
    return render_template(
        "replenishment.html",
        products=rows,
        search=search,
        category=category,
        supplier=supplier,
        product_type=product_type,
        sort=sort,
        categories=categories,
        suppliers=suppliers,
        pagination=pagination_context(page, page_size, total),
    )


@app.route("/herramientas", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero")
def tool_instances():
    db = get_db()
    if request.method == "POST":
        try:
            product_id = int(request.form["product_id"])
        except (KeyError, ValueError):
            flash("Debe seleccionar una herramienta.", "error")
            return redirect(url_for("tool_instances"))
        serial_number = normalize_text(request.form.get("serial_number", ""), 60)
        product = db.execute(
            "SELECT id, name, product_type FROM products WHERE id = ?",
            (product_id,),
        ).fetchone()
        if product is None or product["product_type"] != "herramienta":
            flash("El producto seleccionado no es una herramienta.", "error")
            return redirect(url_for("tool_instances"))
        if not serial_number:
            flash("El número de serie es obligatorio.", "error")
            return redirect(url_for("tool_instances"))
        try:
            with db:
                cursor = db.execute(
                    "INSERT INTO tool_instances (product_id, serial_number, notes) VALUES (?, ?, ?)",
                    (product_id, serial_number, normalize_text(request.form.get("notes", ""), 500)),
                )
                audit("create", "tool_instance", cursor.lastrowid, f"Unidad {serial_number} creada")
            flash("Herramienta individual registrada.", "success")
        except sqlite3.IntegrityError:
            flash("El número de serie ya existe.", "error")
        return redirect(url_for("tool_instances"))
    search = request.args.get("q", "").strip()
    status = request.args.get("status", "").strip()
    valid_statuses = {"disponible", "entregada", "reparacion", "danada", "perdida", "baja"}
    if status not in valid_statuses:
        status = ""
    employee_id = request.args.get("employee_id", "").strip()
    try:
        employee_id_value = int(employee_id) if employee_id else None
    except ValueError:
        employee_id = ""
        employee_id_value = None
    query = """SELECT ti.*, p.name AS product_name, p.code AS product_code,
                      e.full_name AS employee_name
               FROM tool_instances ti JOIN products p ON p.id = ti.product_id
               LEFT JOIN employees e ON e.id = ti.current_employee_id
               WHERE (p.name LIKE ? OR p.code LIKE ? OR ti.serial_number LIKE ?)"""
    params: list[object] = [f"%{search}%"] * 3
    if status:
        query += " AND ti.status = ?"
        params.append(status)
    if employee_id_value is not None:
        query += " AND ti.current_employee_id = ?"
        params.append(employee_id_value)
    page, page_size, offset = pagination_params()
    total = db.execute(f"SELECT COUNT(*) FROM ({query})", params).fetchone()[0]
    instances = db.execute(
        query + " ORDER BY p.name, ti.serial_number, ti.id LIMIT ? OFFSET ?",
        [*params, page_size, offset],
    ).fetchall()
    products_list = db.execute(
        "SELECT id, name, brand FROM products WHERE product_type = 'herramienta' ORDER BY name"
    ).fetchall()
    employees_list = db.execute(
        "SELECT id, full_name, code FROM employees ORDER BY full_name, id"
    ).fetchall()
    return render_template(
        "tool_instances.html",
        instances=instances,
        products=products_list,
        employees=employees_list,
        search=search,
        status=status,
        employee_id=employee_id,
        pagination=pagination_context(page, page_size, total),
    )


@app.route("/herramientas/<int:instance_id>/historial")
@role_required("admin", "supervisor", "pañolero")
def tool_instance_history(instance_id: int):
    db = get_db()
    instance = db.execute(
        """SELECT ti.id, ti.serial_number, ti.status, ti.notes, ti.created_at,
                  p.name AS product_name, p.code AS product_code,
                  e.full_name AS employee_name
           FROM tool_instances ti
           JOIN products p ON p.id = ti.product_id
           LEFT JOIN employees e ON e.id = ti.current_employee_id
           WHERE ti.id = ?""",
        (instance_id,),
    ).fetchone()
    if instance is None:
        flash("La herramienta no existe.", "error")
        return redirect(url_for("tool_instances"))
    movement_page, page_size, movement_offset = pagination_params("movement_page")
    status_page, _, status_offset = pagination_params("status_page")
    movement_total = db.execute(
        "SELECT COUNT(*) FROM movements WHERE tool_instance_id = ?",
        (instance_id,),
    ).fetchone()[0]
    status_total = db.execute(
        "SELECT COUNT(*) FROM tool_instance_status_history WHERE instance_id = ?",
        (instance_id,),
    ).fetchone()[0]
    movements = db.execute(
        """SELECT m.id, m.movement_date, m.created_at, m.movement_type,
                  m.quantity, m.notes, e.full_name AS employee_name,
                  (SELECT u.username
                     FROM audit_log a
                     LEFT JOIN users u ON u.id = a.user_id
                    WHERE a.entity_type = 'movement'
                      AND a.entity_id = m.id
                      AND a.action = 'create'
                    ORDER BY a.id DESC
                    LIMIT 1) AS username
           FROM movements m
           LEFT JOIN employees e ON e.id = m.employee_id
           WHERE m.tool_instance_id = ?
           ORDER BY m.movement_date DESC, m.id DESC
           LIMIT ? OFFSET ?""",
        (instance_id, page_size, movement_offset),
    ).fetchall()
    status_history = db.execute(
        """SELECT h.previous_status, h.new_status, h.employee_id, h.notes,
                  h.changed_at, u.username
           FROM tool_instance_status_history h
           LEFT JOIN users u ON u.id = h.changed_by
           WHERE h.instance_id = ?
           ORDER BY h.changed_at DESC, h.id DESC
           LIMIT ? OFFSET ?""",
        (instance_id, page_size, status_offset),
    ).fetchall()
    return render_template(
        "tool_instance_history.html",
        instance=instance,
        movements=movements,
        status_history=status_history,
        movement_pagination=pagination_context(movement_page, page_size, movement_total),
        status_pagination=pagination_context(status_page, page_size, status_total),
    )


@app.route("/etiquetas", methods=["GET"])
@role_required("admin", "supervisor", "pañolero", "consulta")
def labels():
    db = get_db()
    return render_template(
        "labels.html",
        products=db.execute("SELECT id, name, code FROM products WHERE active = 1 ORDER BY name LIMIT ?",
                            (LABELS_MAX_COUNT,)).fetchall(),
        employees=db.execute(
            "SELECT id, full_name, code FROM employees WHERE active = 1 ORDER BY full_name LIMIT ?",
            (LABELS_MAX_COUNT,)).fetchall(),
        tools=db.execute(
            """SELECT ti.id, ti.serial_number, p.name AS product_name
               FROM tool_instances ti JOIN products p ON p.id = ti.product_id
               ORDER BY p.name, ti.serial_number LIMIT ?""",
            (LABELS_MAX_COUNT,)).fetchall(),
        max_count=LABELS_MAX_COUNT,
    )


@app.route("/etiquetas/pdf", methods=["POST"])
@role_required("admin", "supervisor", "pañolero", "consulta")
def labels_pdf():
    queries = {
        "products": ("SELECT id, name, code AS identifier FROM products WHERE id IN ({})", "Producto"),
        "employees": (
            "SELECT id, full_name AS name, code AS identifier FROM employees "
            "WHERE active = 1 AND id IN ({})", "Empleado",
        ),
        "tools": (
            "SELECT ti.id, p.name AS name, ti.serial_number AS identifier "
            "FROM tool_instances ti JOIN products p ON p.id = ti.product_id "
            "WHERE ti.id IN ({})", "Herramienta",
        ),
    }
    label_type = request.form.get("label_type", "")
    if label_type not in queries:
        return "Tipo de etiqueta no válido.", 400
    raw_ids = request.form.getlist("ids")
    if not raw_ids or len(raw_ids) > LABELS_MAX_COUNT:
        return f"Debe seleccionar entre 1 y {LABELS_MAX_COUNT} registros.", 400
    try:
        ids = [int(value) for value in raw_ids]
    except (TypeError, ValueError):
        return "Los identificadores no son válidos.", 400
    if any(value <= 0 for value in ids) or len(set(ids)) != len(ids):
        return "Los identificadores no son válidos.", 400
    query, kind = queries[label_type]
    rows = get_db().execute(query.format(",".join("?" for _ in ids)), ids).fetchall()
    if len(rows) != len(ids):
        return "Uno o más registros no existen o no están disponibles.", 400
    by_id = {row["id"]: row for row in rows}
    records = [{"name": str(by_id[item_id]["name"]), "kind": kind,
                "identifier": str(by_id[item_id]["identifier"]),
                "barcode": str(by_id[item_id]["identifier"])} for item_id in ids]
    return send_file(build_labels_pdf(records), mimetype="application/pdf",
                     as_attachment=True, download_name="etiquetas-inventario.pdf")


@app.route("/herramientas/<int:instance_id>/estado", methods=["POST"])
@role_required("admin", "supervisor", "pañolero")
def update_tool_instance_status(instance_id: int):
    db = get_db()
    status = request.form.get("status", "")
    if status not in {"disponible", "reparacion", "danada", "perdida", "baja"}:
        flash("Estado inválido.", "error")
        return redirect(url_for("tool_instances"))
    reason = normalize_text(
        request.form.get("reason", "") or request.form.get("notes", ""), 1000
    )
    if status in {"reparacion", "danada", "perdida", "baja"} and not reason:
        flash("Debe indicar un motivo para este estado.", "error")
        return redirect(url_for("tool_instances"))
    instance = db.execute(
        "SELECT serial_number, status, current_employee_id FROM tool_instances WHERE id = ?",
        (instance_id,),
    ).fetchone()
    if instance is None:
        flash("La herramienta no existe.", "error")
        return redirect(url_for("tool_instances"))
    if instance["status"] == status:
        flash("La herramienta ya se encuentra en ese estado.", "info")
        return redirect(url_for("tool_instances"))
    with db:
        db.execute(
            "UPDATE tool_instances SET status = ?, current_employee_id = NULL WHERE id = ?",
            (status, instance_id),
        )
        db.execute(
            """INSERT INTO tool_instance_status_history
               (instance_id, previous_status, new_status, employee_id, notes, changed_by)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                instance_id,
                instance["status"],
                status,
                instance["current_employee_id"],
                reason,
                g.user["id"],
            ),
        )
        audit(
            "update",
            "tool_instance",
            instance_id,
            f"Unidad {instance['serial_number']} pasó de {instance['status']} a {status}"
            + (f": {reason}" if reason else ""),
        )
    flash("Estado de la herramienta actualizado.", "success")
    return redirect(url_for("tool_instances"))


@app.route("/usuarios/<int:user_id>/contrasena", methods=["POST"])
@role_required("admin")
def change_user_password(user_id: int):
    db = get_db()
    user = db.execute("SELECT id, username FROM users WHERE id = ?", (user_id,)).fetchone()
    if user is None:
        flash("El usuario no existe.", "error")
        return redirect(url_for("users"))
    new_password = request.form.get("new_password", "")
    confirm_password = request.form.get("confirm_password", "")
    if len(new_password) < 8:
        flash("La contraseña debe tener al menos 8 caracteres.", "error")
        return redirect(url_for("users"))
    if new_password != confirm_password:
        flash("Las contraseñas no coinciden.", "error")
        return redirect(url_for("users"))
    with db:
        db.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (generate_password_hash(new_password), user_id),
        )
        audit("update", "user", user_id,
              f"Contraseña de {user['username']} actualizada por admin")
    flash(f"Contraseña de '{user['username']}' actualizada correctamente.", "success")
    return redirect(url_for("users"))


@app.route("/usuarios/<int:user_id>/estado", methods=["POST"])
@role_required("admin")
def toggle_user(user_id: int):
    db = get_db()
    if user_id == g.user["id"]:
        flash("No puede desactivar su propio usuario.", "error")
        return redirect(url_for("users"))
    user = db.execute("SELECT username, active FROM users WHERE id = ?", (user_id,)).fetchone()
    if user is None:
        flash("El usuario no existe.", "error")
        return redirect(url_for("users"))
    new_status = 0 if user["active"] else 1
    with db:
        db.execute("UPDATE users SET active = ? WHERE id = ?", (new_status, user_id))
        audit("update", "user", user_id, f"Usuario {user['username']} {'activado' if new_status else 'desactivado'}")
    flash("Estado del usuario actualizado.", "success")
    return redirect(url_for("users"))


@app.route("/auditoria")
@role_required("admin", "supervisor")
def audit_report():
    db = get_db()
    action = normalize_text(request.args.get("action", ""), 80)
    # ``entity``/``q`` are the public names used by the template; accept the
    # column-oriented aliases too so bookmarked URLs remain interoperable.
    entity_type = normalize_text(
        request.args.get("entity", request.args.get("entity_type", "")), 80
    )
    search = normalize_text(request.args.get("q", request.args.get("search", "")), 120)
    filters = []
    params = []
    if action:
        filters.append("a.action = ?")
        params.append(action)
    if entity_type:
        filters.append("a.entity_type = ?")
        params.append(entity_type)
    if search:
        filters.append("(u.username LIKE ? OR a.details LIKE ?)")
        search_pattern = f"%{search}%"
        params.extend((search_pattern, search_pattern))
    where = f" WHERE {' AND '.join(filters)}" if filters else ""
    total = db.execute(
        f"""SELECT COUNT(*)
            FROM audit_log a LEFT JOIN users u ON u.id = a.user_id{where}""",
        params,
    ).fetchone()[0]
    page, page_size, offset = pagination_params()
    entries = db.execute(
        f"""SELECT a.*, u.username
            FROM audit_log a LEFT JOIN users u ON u.id = a.user_id
            {where}
            ORDER BY a.id DESC LIMIT ? OFFSET ?""",
        [*params, page_size, offset],
    ).fetchall()
    return render_template(
        "audit.html",
        entries=entries,
        action=action,
        entity_type=entity_type,
        search=search,
        pagination=pagination_context(page, page_size, total),
    )


@app.route("/inventarios", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero")
def inventory_counts():
    db = get_db()
    if request.method == "POST":
        name = normalize_text(request.form.get("name", ""), 120)
        count_date = request.form.get("count_date", "").strip()
        try:
            date.fromisoformat(count_date)
        except ValueError:
            flash("La fecha del inventario no es válida.", "error")
            return redirect(url_for("inventory_counts"))
        if not name:
            flash("El nombre del inventario es obligatorio.", "error")
            return redirect(url_for("inventory_counts"))
        with db:
            cursor = db.execute(
                """INSERT INTO inventory_counts (name, count_date, created_by)
                   VALUES (?, ?, ?)""",
                (name, count_date, g.user["id"]),
            )
            count_id = cursor.lastrowid
            products = db.execute("SELECT id, stock FROM products ORDER BY id").fetchall()
            db.executemany(
                """INSERT INTO inventory_count_items
                   (count_id, product_id, theoretical_stock)
                   VALUES (?, ?, ?)""",
                [(count_id, product["id"], product["stock"]) for product in products],
            )
            audit("create", "inventory_count", count_id, f"Inventario físico '{name}' iniciado")
        flash("Inventario físico creado con el stock teórico actual.", "success")
        return redirect(url_for("inventory_count_detail", count_id=count_id))
    search = request.args.get("q", "").strip()
    status = request.args.get("status", "").strip().lower()
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    if status not in {"open", "closed"}:
        status = ""
    filters = []
    params: list[str] = []
    if search:
        filters.append("c.name LIKE ?")
        params.append(f"%{search}%")
    if status:
        filters.append("c.status = ?")
        params.append(status)
    if date_from:
        filters.append("c.count_date >= ?")
        params.append(date_from)
    if date_to:
        filters.append("c.count_date <= ?")
        params.append(date_to)
    where = f" WHERE {' AND '.join(filters)}" if filters else ""
    count_query = f"FROM inventory_counts c JOIN users u ON u.id = c.created_by{where}"
    total = db.execute(f"SELECT COUNT(*) {count_query}", params).fetchone()[0]
    page, page_size, offset = pagination_params()
    counts = db.execute(
        f"""SELECT c.*, u.username AS created_by_name,
                   COUNT(i.id) AS item_count,
                   SUM(CASE WHEN i.counted_stock IS NOT NULL THEN 1 ELSE 0 END) AS counted_items
            FROM inventory_counts c
            JOIN users u ON u.id = c.created_by
            LEFT JOIN inventory_count_items i ON i.count_id = c.id
            {where}
            GROUP BY c.id
            ORDER BY c.id DESC
            LIMIT ? OFFSET ?""",
        [*params, page_size, offset],
    ).fetchall()
    return render_template(
        "inventory_counts.html", counts=counts, today=date.today().isoformat(),
        search=search, status=status, date_from=date_from, date_to=date_to,
        pagination=pagination_context(page, page_size, total),
    )


@app.route("/inventarios/<int:count_id>", methods=["GET", "POST"])
@role_required("admin", "supervisor", "pañolero")
def inventory_count_detail(count_id: int):
    db = get_db()
    count = count_row(count_id)
    if count is None:
        flash("El inventario no existe.", "error")
        return redirect(url_for("inventory_counts"))
    if request.method == "POST":
        if count["status"] != "open":
            flash("El inventario ya está cerrado.", "error")
            return redirect(url_for("inventory_count_detail", count_id=count_id))
        values: list[tuple[int, int, int]] = []
        for item in db.execute(
            "SELECT id, product_id FROM inventory_count_items WHERE count_id = ?",
            (count_id,),
        ).fetchall():
            raw_value = request.form.get(f"counted_{item['id']}", "").strip()
            try:
                counted = int(raw_value)
            except ValueError:
                flash("Todas las cantidades contadas deben ser números enteros.", "error")
                return redirect(url_for("inventory_count_detail", count_id=count_id))
            if counted < 0:
                flash("Las cantidades contadas no pueden ser negativas.", "error")
                return redirect(url_for("inventory_count_detail", count_id=count_id))
            values.append((counted, item["id"], count_id))
        with db:
            db.executemany(
                "UPDATE inventory_count_items SET counted_stock = ? WHERE id = ? AND count_id = ?",
                values,
            )
            audit("update", "inventory_count", count_id, "Conteo físico actualizado")
        flash("Conteo guardado correctamente.", "success")
        return redirect(url_for("inventory_count_detail", count_id=count_id))
    items = db.execute(
        """SELECT i.*, p.code, p.name, p.brand, p.category, p.unit
           FROM inventory_count_items i JOIN products p ON p.id = i.product_id
           WHERE i.count_id = ? ORDER BY p.name""",
        (count_id,),
    ).fetchall()
    return render_template("inventory_count_detail.html", count=count, items=items)


@app.route("/inventarios/<int:count_id>/cerrar", methods=["POST"])
@role_required("admin", "supervisor")
def close_inventory_count(count_id: int):
    db = get_db()
    count = count_row(count_id)
    if count is None:
        flash("El inventario no existe.", "error")
        return redirect(url_for("inventory_counts"))
    if count["status"] != "open":
        flash("El inventario ya está cerrado.", "error")
        return redirect(url_for("inventory_count_detail", count_id=count_id))
    items = db.execute(
        """SELECT i.*, p.code, p.stock, p.name
           FROM inventory_count_items i JOIN products p ON p.id = i.product_id
           WHERE i.count_id = ?""",
        (count_id,),
    ).fetchall()
    if not items or any(item["counted_stock"] is None for item in items):
        flash("Debe completar el conteo de todos los productos antes de cerrar.", "error")
        return redirect(url_for("inventory_count_detail", count_id=count_id))
    with db:
        for item in items:
            difference = item["counted_stock"] - item["stock"]
            if difference:
                movement_type = "entrada" if difference > 0 else "salida"
                quantity = abs(difference)
                cursor = db.execute(
                    """INSERT INTO movements
                       (product_id, movement_type, quantity, notes, movement_date)
                       VALUES (?, ?, ?, ?, ?)""",
                    (
                        item["product_id"],
                        movement_type,
                        quantity,
                        f"Ajuste por inventario físico #{count_id}: {count['name']}",
                        count["count_date"],
                    ),
                )
                db.execute(
                    "UPDATE products SET stock = ? WHERE id = ?",
                    (item["counted_stock"], item["product_id"]),
                )
                audit(
                    "adjust",
                    "product",
                    item["product_id"],
                    f"Inventario #{count_id}: stock {item['stock']} -> {item['counted_stock']}; movimiento #{cursor.lastrowid}",
                )
        db.execute(
            "UPDATE inventory_counts SET status = 'closed', closed_by = ?, closed_at = CURRENT_TIMESTAMP WHERE id = ?",
            (g.user["id"], count_id),
        )
        audit("close", "inventory_count", count_id, f"Inventario físico '{count['name']}' cerrado")
    flash("Inventario cerrado y diferencias aplicadas al stock.", "success")
    return redirect(url_for("inventory_counts"))


def _send_created_backup() -> Any:
    try:
        backup_path = create_backup()
    except (OSError, sqlite3.Error, ValueError) as exc:
        return f"No se pudo crear la copia de seguridad: {exc}", 500
    audit("backup", "database", None, f"Copia de seguridad {backup_path.name} generada")
    get_db().commit()
    return send_file(
        backup_path,
        as_attachment=True,
        download_name=backup_path.name,
        mimetype="application/octet-stream",
    )


@app.route("/backup")
@role_required("admin")
def download_backup():
    return _send_created_backup()


@app.route("/backup/manual", methods=["POST"])
@role_required("admin")
def manual_backup():
    try:
        backup_path = create_backup()
    except (OSError, sqlite3.Error, ValueError) as exc:
        flash(f"No se pudo crear la copia de seguridad: {exc}", "error")
        return redirect(url_for("dashboard")), 500
    audit("backup", "database", None, f"Copia de seguridad {backup_path.name} generada")
    get_db().commit()
    return send_file(
        backup_path,
        as_attachment=True,
        download_name=backup_path.name,
        mimetype="application/octet-stream",
    )


@app.route("/admin/vacuum", methods=["POST"])
@role_required("admin")
def vacuum_db():
    db = get_db()
    size_before = DATABASE.stat().st_size
    db.execute("VACUUM")
    db.commit()
    size_after = DATABASE.stat().st_size
    freed_kb = max(0, (size_before - size_after) // 1024)
    audit("vacuum", "database", None, f"VACUUM ejecutado; liberados {freed_kb} KB")
    flash(f"Base de datos compactada correctamente. Espacio liberado: {freed_kb} KB.", "success")
    return redirect(url_for("dashboard"))


@app.cli.command("backup")
def backup_command() -> None:
    """Create one backup for external schedulers such as Windows Task Scheduler."""
    try:
        backup_path = create_backup()
    except (OSError, sqlite3.Error, ValueError) as exc:
        raise RuntimeError(f"No se pudo crear la copia de seguridad: {exc}") from exc
    audit("backup", "database", None, f"Copia de seguridad {backup_path.name} generada")
    get_db().commit()
    print(backup_path)


def make_pdf_table(rows: list[list[str]], widths: list[float]) -> Table:
    table = Table(rows, colWidths=widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1565c0")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cfd6de")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f7f9")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


@app.route("/barcodes/<path:filename>")
def barcodes(filename: str):
    return send_from_directory(BARCODE_DIR, filename)


if os.environ.get("PANOL_SKIP_INIT_DB", "").lower() not in {"1", "true", "yes"}:
    with app.app_context():
        init_db()


if __name__ == "__main__":
    import threading
    import webbrowser
    debug = os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    if not debug:
        threading.Timer(1.5, lambda: webbrowser.open("http://localhost:5000")).start()
    app.run(debug=debug)
