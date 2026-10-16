#!/usr/bin/env python3
"""Carga datos FICTICIOS de ejemplo para probar o mostrar Inventario Pañol.

Crea (o completa) la base data/panol.db con proveedores, empleados, materiales,
herramientas con número de serie y un historial de movimientos de los últimos días.
Todos los nombres, teléfonos y correos son inventados.

Uso:
    python seed_demo.py            # solo si la base está vacía
    python seed_demo.py --reset    # borra data/panol.db y la vuelve a crear

Usuario inicial: admin. La contraseña se toma de PANOL_ADMIN_PASSWORD; si no está
definida, para esta base DE EJEMPLO se usa "Demo-Panol-2026". No use datos ni
contraseñas de ejemplo en una instalación real.
"""

from __future__ import annotations

import argparse
import os
import random
import sqlite3
import sys
from datetime import date, timedelta
from pathlib import Path

os.environ.setdefault("PANOL_SKIP_INIT_DB", "1")
os.environ.setdefault("PANOL_ADMIN_PASSWORD", "Demo-Panol-2026")
sys.path.insert(0, str(Path(__file__).resolve().parent))

import app as panol  # noqa: E402

DIAS_HISTORIAL = 35

PROVEEDORES = [
    ("Ferretería Central (demo)", "Mostrador", "11-0000-0001", "ventas@ferreteria-demo.test"),
    ("Eléctrica Norte (demo)", "Ventas", "11-0000-0002", "contacto@electrica-demo.test"),
    ("Seguridad Industrial Sur (demo)", "Atención", "11-0000-0003", "info@seguridad-demo.test"),
]

EMPLEADOS = [
    ("EMP-001", "Ana Ejemplo", "LEG-001"),
    ("EMP-002", "Bruno Ejemplo", "LEG-002"),
    ("EMP-003", "Carla Ejemplo", "LEG-003"),
    ("EMP-004", "Diego Ejemplo", "LEG-004"),
    ("EMP-005", "Elena Ejemplo", "LEG-005"),
    ("EMP-006", "Fabián Ejemplo", "LEG-006"),
]

# (código, nombre, marca, categoría, presentación, unidad, stock mínimo, stock inicial,
#  cantidad de reposición, costo unitario, índice de proveedor, ubicación)
MATERIALES = [
    ("ELE-001", "Cable unipolar 2,5 mm", "Marca A", "Electricidad", "rollo x 100 m", "rollo", 4, 12, 6, 48500.0, 1, "Estante A1"),
    ("ELE-002", "Tomacorriente doble", "Marca B", "Electricidad", "unidad", "unidad", 10, 40, 20, 3200.0, 1, "Estante A2"),
    ("ELE-003", "Cinta aisladora", "Marca B", "Electricidad", "unidad", "unidad", 10, 14, 20, 950.0, 1, "Estante A2"),
    ("LIM-001", "Detergente industrial 5 L", "Marca C", "Limpieza", "bidón", "bidón", 3, 9, 6, 7800.0, 0, "Estante B1"),
    ("LIM-002", "Bolsa residuo 90x120", "Marca C", "Limpieza", "paquete x 10", "paquete", 5, 22, 10, 2100.0, 0, "Estante B1"),
    ("SEG-001", "Guantes de seguridad", "Marca D", "Higiene y seguridad", "par", "par", 20, 60, 40, 1800.0, 2, "Estante C1"),
    ("SEG-002", "Anteojos de seguridad", "Marca D", "Higiene y seguridad", "unidad", "unidad", 10, 25, 20, 2600.0, 2, "Estante C1"),
    ("SEG-003", "Casco de obra", "Marca D", "Higiene y seguridad", "unidad", "unidad", 5, 8, 10, 6900.0, 2, "Estante C2"),
    ("FER-001", "Tornillo autoperforante 6x1", "Marca E", "Ferretería", "caja x 100", "caja", 8, 30, 15, 2900.0, 0, "Estante D1"),
    ("FER-002", "Mecha para hormigón 8 mm", "Marca E", "Ferretería", "unidad", "unidad", 6, 5, 12, 2400.0, 0, "Estante D2"),
    ("FER-003", "Taco Fischer 8 mm", "Marca E", "Ferretería", "caja x 50", "caja", 6, 18, 12, 2200.0, 0, "Estante D1"),
    ("DUR-001", "Placa de yeso 12,5 mm", "Marca F", "Durlock", "placa", "placa", 15, 70, 30, 9800.0, 0, "Depósito 1"),
    ("DUR-002", "Perfil montante 70 mm", "Marca F", "Durlock", "barra 2,6 m", "barra", 20, 90, 40, 3100.0, 0, "Depósito 1"),
    ("COL-001", "Adhesivo cerámico 30 kg", "Marca G", "Colocación", "bolsa", "bolsa", 6, 4, 12, 8700.0, 0, "Depósito 2"),
    ("COL-002", "Pastina blanca 1 kg", "Marca G", "Colocación", "bolsa", "bolsa", 5, 16, 10, 1500.0, 0, "Depósito 2"),
]

# (código, nombre, marca, categoría, costo, ubicación, números de serie)
HERRAMIENTAS = [
    ("HER-001", "Taladro percutor 13 mm", "Marca H", "Ferretería", 85000.0, "Pañol", ["TP-001", "TP-002", "TP-003"]),
    ("HER-002", "Amoladora angular 4 1/2", "Marca H", "Ferretería", 72000.0, "Pañol", ["AA-001", "AA-002"]),
    ("HER-003", "Nivel láser", "Marca I", "Colocación", 135000.0, "Pañol", ["NL-001"]),
    ("HER-004", "Atornillador eléctrico", "Marca H", "Durlock", 64000.0, "Pañol", ["AT-001", "AT-002", "AT-003"]),
    ("HER-005", "Alargue 25 m", "Marca J", "Electricidad", 21000.0, "Pañol", ["AL-001", "AL-002", "AL-003", "AL-004"]),
]


def cargar(db: sqlite3.Connection, rng: random.Random) -> dict[str, int]:
    hoy = date.today()
    inicio = hoy - timedelta(days=DIAS_HISTORIAL)

    ids_proveedor = []
    for nombre, contacto, telefono, correo in PROVEEDORES:
        cur = db.execute(
            "INSERT INTO suppliers (name, contact, phone, email) VALUES (?, ?, ?, ?)",
            (nombre, contacto, telefono, correo),
        )
        ids_proveedor.append(cur.lastrowid)

    ids_empleado = []
    for codigo, nombre, legajo in EMPLEADOS:
        cur = db.execute(
            "INSERT INTO employees (code, full_name, legajo) VALUES (?, ?, ?)",
            (codigo, nombre, legajo),
        )
        ids_empleado.append(cur.lastrowid)
        panol.generate_barcode(codigo)

    contador = {"movimientos": 0}

    def movimiento(producto_id, tipo, cantidad, dia, empleado_id=None, instancia_id=None, nota=""):
        db.execute(
            """INSERT INTO movements
               (product_id, employee_id, tool_instance_id, movement_type, quantity, notes, movement_date)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (producto_id, empleado_id, instancia_id, tipo, cantidad, nota, dia.isoformat()),
        )
        contador["movimientos"] += 1

    # Materiales: stock inicial y movimientos coherentes con el stock disponible.
    for (codigo, nombre, marca, cat, present, unidad, minimo, inicial, repo, costo, prov, ubic) in MATERIALES:
        cur = db.execute(
            """INSERT INTO products
               (code, name, brand, category, product_type, presentation, unit, minimum_stock,
                stock, reorder_quantity, unit_cost, supplier_id, location)
               VALUES (?, ?, ?, ?, 'material', ?, ?, ?, 0, ?, ?, ?, ?)""",
            (codigo, nombre, marca, cat, present, unidad, minimo, repo, costo,
             ids_proveedor[prov], ubic),
        )
        pid = cur.lastrowid
        panol.generate_barcode(codigo)
        stock = inicial
        movimiento(pid, "entrada", inicial, inicio, nota="Stock inicial de ejemplo")
        for _ in range(rng.randint(2, 5)):
            dia = inicio + timedelta(days=rng.randint(1, DIAS_HISTORIAL))
            maximo = max(1, stock // 3)
            if stock <= 1:
                break
            cantidad = rng.randint(1, maximo)
            movimiento(pid, "salida", cantidad, dia, rng.choice(ids_empleado), nota="Entrega de ejemplo")
            stock -= cantidad
        db.execute("UPDATE products SET stock = ? WHERE id = ?", (stock, pid))

    # Herramientas: unidades con número de serie, algunas entregadas.
    entregas_pendientes = 0
    for (codigo, nombre, marca, cat, costo, ubic, series) in HERRAMIENTAS:
        cur = db.execute(
            """INSERT INTO products
               (code, name, brand, category, product_type, presentation, unit, minimum_stock,
                stock, reorder_quantity, unit_cost, location, retornable)
               VALUES (?, ?, ?, ?, 'herramienta', 'unidad', 'unidad', 1, 0, 1, ?, ?, 1)""",
            (codigo, nombre, marca, cat, costo, ubic),
        )
        pid = cur.lastrowid
        panol.generate_barcode(codigo)
        total = len(series)
        movimiento(pid, "entrada", total, inicio, nota="Alta de herramientas de ejemplo")
        stock = total
        for indice, serie in enumerate(series):
            inst = db.execute(
                "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)", (pid, serie)
            ).lastrowid
            if indice == 0 and rng.random() < 0.9:
                empleado = rng.choice(ids_empleado)
                dia_salida = hoy - timedelta(days=rng.choice([1, 2, 6, 9]))
                movimiento(pid, "salida", 1, dia_salida, empleado, inst, "Entrega de ejemplo")
                stock -= 1
                if entregas_pendientes < 3 or dia_salida <= hoy - timedelta(days=6):
                    db.execute(
                        "UPDATE tool_instances SET status = 'entregada', current_employee_id = ? WHERE id = ?",
                        (empleado, inst),
                    )
                    entregas_pendientes += 1
                else:
                    movimiento(pid, "devolucion", 1, dia_salida + timedelta(days=1), empleado, inst, "Devolución de ejemplo")
                    stock += 1
            elif indice == 1:
                dia_salida = hoy - timedelta(days=rng.randint(10, 20))
                empleado = rng.choice(ids_empleado)
                movimiento(pid, "salida", 1, dia_salida, empleado, inst, "Entrega de ejemplo")
                movimiento(pid, "devolucion", 1, dia_salida + timedelta(days=2), empleado, inst, "Devolución de ejemplo")
        db.execute("UPDATE products SET stock = ? WHERE id = ?", (stock, pid))

    return {
        "proveedores": len(PROVEEDORES),
        "empleados": len(EMPLEADOS),
        "materiales": len(MATERIALES),
        "herramientas": len(HERRAMIENTAS),
        "movimientos": contador["movimientos"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Carga datos ficticios de ejemplo.")
    parser.add_argument("--reset", action="store_true", help="Borra data/panol.db y la recrea")
    args = parser.parse_args()

    database = Path(panol.DATABASE)
    if args.reset and database.exists():
        database.unlink()

    with panol.app.app_context():
        panol.init_db()
        db = panol.get_db()
        if db.execute("SELECT COUNT(*) FROM products").fetchone()[0] > 0:
            print("La base ya tiene productos. Use --reset para recrearla con datos de ejemplo.", file=sys.stderr)
            return 1
        with db:
            resumen = cargar(db, random.Random(2026))

    print("Datos de ejemplo cargados:")
    for clave, valor in resumen.items():
        print(f"  {clave}: {valor}")
    print(f"Base: {database}")
    print("Usuario: admin | Contraseña: la de PANOL_ADMIN_PASSWORD (por defecto, solo en este ejemplo: Demo-Panol-2026)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
