"""
Generacion automatica de ordenes de compra borrador — Inventario Paniol.
Crea una OC borrador por proveedor para productos bajo stock minimo
que no esten ya cubiertos por una OC activa (borrador/solicitada/parcial).
"""
import os
import sqlite3
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

SCRIPT_DIR = Path(os.environ.get("PANOL_HOME") or Path(__file__).resolve().parent.parent)
DATABASE   = SCRIPT_DIR / "data" / "panol.db"
LOG_FILE   = SCRIPT_DIR / "data" / "ordenes_auto.log"


def log(msg: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}\n"
    print(line, end="")
    try:
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except OSError:
        pass


def get_admin_user_id(db: sqlite3.Connection) -> int | None:
    row = db.execute(
        "SELECT id FROM users WHERE role = 'admin' AND active = 1 ORDER BY id LIMIT 1"
    ).fetchone()
    return row[0] if row else None


def query_productos_sin_oc(db: sqlite3.Connection) -> list:
    """Productos bajo minimo sin cobertura en ninguna OC activa, con proveedor asignado."""
    return db.execute(
        """SELECT p.id AS product_id, p.name, p.code, p.stock, p.minimum_stock,
                  p.supplier_id, s.name AS supplier_name,
                  CASE WHEN p.reorder_quantity > 0 THEN p.reorder_quantity
                       ELSE MAX(p.minimum_stock - p.stock, 1) END AS qty
           FROM products p
           JOIN suppliers s ON s.id = p.supplier_id
           WHERE p.active = 1
             AND p.stock <= p.minimum_stock
             AND NOT EXISTS (
                 SELECT 1 FROM purchase_order_items poi
                 JOIN purchase_orders po ON po.id = poi.purchase_order_id
                 WHERE poi.product_id = p.id
                   AND po.status IN ('borrador', 'solicitada', 'parcial')
             )
           ORDER BY s.name, p.name"""
    ).fetchall()


def next_order_number(db: sqlite3.Connection, today_str: str) -> str:
    prefix = f"OC-{today_str.replace('-', '')}-"
    count = db.execute(
        "SELECT COUNT(*) FROM purchase_orders WHERE order_number LIKE ?",
        (f"{prefix}%",),
    ).fetchone()[0]
    return f"{prefix}{count + 1:03d}"


def crear_oc(
    db: sqlite3.Connection,
    supplier_id: int,
    productos: list,
    user_id: int,
    today: str,
) -> int:
    order_number = next_order_number(db, today)
    notes = f"Generada automaticamente el {today}. Revisar antes de solicitar."
    cursor = db.execute(
        """INSERT INTO purchase_orders
           (order_number, supplier_id, order_date, notes, created_by)
           VALUES (?, ?, ?, ?, ?)""",
        (order_number, supplier_id, today, notes, user_id),
    )
    order_id = cursor.lastrowid
    for p in productos:
        db.execute(
            """INSERT INTO purchase_order_items
               (purchase_order_id, product_id, quantity_requested)
               VALUES (?, ?, ?)""",
            (order_id, p["product_id"], p["qty"]),
        )
    return order_id


if __name__ == "__main__":
    if not DATABASE.exists():
        log(f"ERROR: Base de datos no encontrada en {DATABASE}")
        sys.exit(1)

    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row

    try:
        user_id = get_admin_user_id(db)
        if user_id is None:
            log("ERROR: no se encontro ningun usuario admin activo en la BD")
            sys.exit(1)

        productos = query_productos_sin_oc(db)

        if not productos:
            log("OK  ningún producto requiere OC nueva")
            sys.exit(0)

        por_proveedor: dict[int, list] = defaultdict(list)
        nombres: dict[int, str] = {}
        for p in productos:
            por_proveedor[p["supplier_id"]].append(p)
            nombres[p["supplier_id"]] = p["supplier_name"]

        today = date.today().isoformat()
        creadas = []

        with db:
            for supplier_id, items in por_proveedor.items():
                oid = crear_oc(db, supplier_id, items, user_id, today)
                creadas.append((nombres[supplier_id], len(items), oid))

        for nombre, n, oid in creadas:
            log(f"OC creada  id={oid}  proveedor={nombre}  items={n}")

        log(f"OK  {len(creadas)} OC(s) borrador creada(s) para {len(productos)} producto(s)")

    finally:
        db.close()
