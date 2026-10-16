"""Pruebas funcionales aisladas para la aplicación Flask del pañol."""

import os
import sys
import csv
import io
from datetime import date
from pathlib import Path

import pytest

os.environ["PANOL_SKIP_INIT_DB"] = "1"
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app as app_module
from openpyxl import load_workbook


@pytest.fixture()
def client(tmp_path, monkeypatch):
    """Crea una base y un directorio de códigos completamente temporales."""
    database = tmp_path / "panol-test.db"
    barcode_dir = tmp_path / "barcodes"
    backup_dir = tmp_path / "backups"
    monkeypatch.setattr(app_module, "DATABASE", database)
    monkeypatch.setattr(app_module, "BARCODE_DIR", barcode_dir)
    monkeypatch.setattr(app_module, "BACKUP_DIR", backup_dir)
    app_module.app.config.update(TESTING=True, DATABASE=database)
    with app_module.app.app_context():
        app_module.init_db()
    with app_module.app.test_client() as test_client:
        yield test_client


def csrf_token(client):
    with client.session_transaction() as session:
        return session["csrf_token"]


def login(client, username="admin", password="admin1234"):
    client.get("/login")
    response = client.post(
        "/login",
        data={
            "username": username,
            "password": password,
            "csrf_token": csrf_token(client),
        },
        follow_redirects=False,
    )
    assert response.status_code == 302
    return csrf_token(client)


def create_product(client, **overrides):
    data = {
        "code": "MAT-001",
        "name": "Cable prueba",
        "brand": "Marca",
        "category": "Electricidad",
        "product_type": "material",
        "presentation": "rollo",
        "unit": "unidad",
        "stock": "5",
        "minimum_stock": "1",
        "reorder_quantity": "2",
        "location": "A-01",
        "csrf_token": csrf_token(client),
    }
    data.update(overrides)
    response = client.post("/productos", data=data)
    assert response.status_code == 302
    with app_module.app.app_context():
        return app_module.get_db().execute(
            "SELECT id FROM products WHERE code = ?", (data["code"],)
        ).fetchone()["id"]


def create_employee(client, **overrides):
    data = {
        "code": "EMP-001",
        "full_name": "Ana Prueba",
        "legajo": "LEG-001",
        "phone": "1111",
        "csrf_token": csrf_token(client),
    }
    data.update(overrides)
    response = client.post("/empleados", data=data)
    assert response.status_code == 302
    with app_module.app.app_context():
        return app_module.get_db().execute(
            "SELECT id FROM employees WHERE code = ?", (data["code"],)
        ).fetchone()["id"]


def test_login_correcto_e_incorrecto(client):
    client.get("/login")
    incorrecto = client.post(
        "/login",
        data={
            "username": "admin",
            "password": "incorrecta",
            "csrf_token": csrf_token(client),
        },
    )
    assert incorrecto.status_code == 401
    assert b"incorrectos" in incorrecto.data

    correcto = client.post(
        "/login",
        data={
            "username": "admin",
            "password": "admin1234",
            "csrf_token": csrf_token(client),
        },
    )
    assert correcto.status_code == 302
    assert correcto.headers["Location"].endswith("/")


def test_proveedores_filtra_por_busqueda_y_estado(client):
    login(client)
    with app_module.app.app_context():
        db = app_module.get_db()
        db.executemany(
            """INSERT INTO suppliers (name, contact, phone, email, active)
               VALUES (?, ?, ?, ?, ?)""",
            [
                ("Proveedor activo", "Ana Contacto", "111", "activo@example.com", 1),
                ("Proveedor inactivo", "Otro Contacto", "222", "inactivo@example.com", 0),
            ],
        )
        db.commit()

    response = client.get("/proveedores?q=activo&status=active")
    assert response.status_code == 200
    assert b"Proveedor activo" in response.data
    assert b"Proveedor inactivo" not in response.data
    assert b'name="q" value="activo"' in response.data


def test_proveedores_pagina_y_preserva_filtros(client):
    login(client)
    with app_module.app.app_context():
        db = app_module.get_db()
        db.executemany(
            "INSERT INTO suppliers (name, active) VALUES (?, 1)",
            [(f"Proveedor {index:02d}",) for index in range(3)],
        )
        db.commit()

    response = client.get("/proveedores?q=Proveedor&status=active&page=2&page_size=1")
    assert response.status_code == 200
    assert b"Pagina 2" in response.data or b"P\xc3\xa1gina 2" in response.data
    assert b"Proveedor 01" in response.data
    assert b"q=Proveedor" in response.data
    assert b"status=active" in response.data
    assert b"page_size=1" in response.data


def test_inventarios_lista_filtra_por_nombre_estado_y_pagina(client):
    login(client)
    for name, count_date in (
        ("Inventario abierto", "2026-09-01"),
        ("Inventario cerrado", "2026-09-02"),
        ("Otra jornada", "2026-09-03"),
    ):
        response = client.post(
            "/inventarios",
            data={"name": name, "count_date": count_date, "csrf_token": csrf_token(client)},
        )
        assert response.status_code == 302

    with app_module.app.app_context():
        closed_id = app_module.get_db().execute(
            "SELECT id FROM inventory_counts WHERE name = ?", ("Inventario cerrado",)
        ).fetchone()["id"]
    with app_module.app.app_context():
        app_module.get_db().execute(
            "UPDATE inventory_counts SET status = 'closed' WHERE id = ?", (closed_id,)
        )
        app_module.get_db().commit()

    listing = client.get("/inventarios?q=Inventario&status=open")
    assert listing.status_code == 200
    assert b"Inventario abierto" in listing.data
    assert b"Inventario cerrado" not in listing.data

    listing = client.get("/inventarios?status=closed")
    assert b"Inventario cerrado" in listing.data
    assert b"Inventario abierto" not in listing.data

    page = client.get("/inventarios?page=2&page_size=1")
    assert page.status_code == 200
    assert b"P\xc3\xa1gina 2 de 3" in page.data
    assert b"page=1" in page.data


def test_dashboard_muestra_producto_bajo_minimo_y_orden_solicitada(client):
    login(client)
    product_id = create_product(
        client,
        code="DASH-LOW",
        name="Producto bajo mínimo",
        stock="1",
        minimum_stock="5",
    )
    with app_module.app.app_context():
        db = app_module.get_db()
        supplier_id = db.execute(
            "INSERT INTO suppliers (name) VALUES (?)", ("Proveedor dashboard",)
        ).lastrowid
        order_id = db.execute(
            """INSERT INTO purchase_orders
               (order_number, supplier_id, status, created_by)
               VALUES (?, ?, 'solicitada', ?)""",
            ("OC-DASH-001", supplier_id, 1),
        ).lastrowid
        db.execute(
            """INSERT INTO purchase_order_items
               (purchase_order_id, product_id, quantity_requested)
               VALUES (?, ?, ?)""",
            (order_id, product_id, 4),
        )
        db.commit()

    response = client.get("/")

    assert response.status_code == 200
    assert b"Producto bajo m\xc3\xadnimo" in response.data
    assert b"OC-DASH-001" in response.data
    assert b"Proveedor dashboard" in response.data


def test_reposicion_filtra_por_texto_categoria_proveedor_y_tipo(client):
    login(client)
    create_product(
        client, code="REP-MAT", name="Cable reposición", category="Electricidad",
        product_type="material", stock="0", minimum_stock="5",
    )
    create_product(
        client, code="REP-TOOL", name="Taladro reposición", category="Herramientas",
        product_type="herramienta", stock="0", minimum_stock="2",
    )
    with app_module.app.app_context():
        db = app_module.get_db()
        supplier_id = db.execute(
            "INSERT INTO suppliers (name) VALUES (?)", ("Proveedor reposición",)
        ).lastrowid
        db.execute(
            "UPDATE products SET supplier_id = ? WHERE code = ?", (supplier_id, "REP-MAT")
        )
        db.commit()

    response = client.get(
        "/reposicion?q=Cable&category=Electricidad"
        f"&supplier={supplier_id}&product_type=material"
    )
    assert response.status_code == 200
    assert b"Cable reposici\xc3\xb3n" in response.data
    assert b"Taladro reposici\xc3\xb3n" not in response.data
    assert b'name="q" value="Cable"' in response.data
    assert b'name="supplier"' in response.data


def test_reposicion_pagina_y_orden_conservan_filtros(client):
    login(client)
    for index in range(3):
        create_product(
            client, code=f"REP-PAGE-{index}", name=f"Reposición {index}",
            stock="0", minimum_stock=str(index + 1),
        )

    response = client.get("/reposicion?sort=faltante&page=2&page_size=1")
    assert response.status_code == 200
    assert b"P\xc3\xa1gina 2 de 3" in response.data
    assert b"sort=faltante" in response.data
    assert b"page_size=1" in response.data


def test_etiquetas_pdf_producto_y_herramienta(client):
    login(client)
    product_id = create_product(client)
    tool_product_id = create_product(
        client, code="TAL-001", name="Taladro", product_type="herramienta"
    )
    with app_module.app.app_context():
        db = app_module.get_db()
        tool_id = db.execute(
            "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)",
            (tool_product_id, "SER-001"),
        ).lastrowid
        db.commit()
    for label_type, item_id in (("products", product_id), ("tools", tool_id)):
        response = client.post("/etiquetas/pdf", data={
            "label_type": label_type, "ids": str(item_id), "csrf_token": csrf_token(client)
        })
        assert response.status_code == 200
        assert response.mimetype == "application/pdf"
        assert response.data.startswith(b"%PDF")


def test_etiquetas_rechaza_ids_invalidos(client):
    login(client)
    for invalid_id in ("no-es-un-id", "999999"):
        response = client.post("/etiquetas/pdf", data={
            "label_type": "products", "ids": invalid_id, "csrf_token": csrf_token(client)
        })
        assert response.status_code == 400


def test_costo_se_actualiza_al_recibir_compra(client):
    login(client)
    product_id = create_product(client, code="COST-001")
    client.post("/proveedores", data={"name": "Proveedor costo", "csrf_token": csrf_token(client)})
    with app_module.app.app_context():
        supplier_id = app_module.get_db().execute(
            "SELECT id FROM suppliers WHERE name = 'Proveedor costo'"
        ).fetchone()["id"]
    order = client.post("/ordenes-compra", data={
        "supplier_id": supplier_id, "order_date": date.today().isoformat(),
        "csrf_token": csrf_token(client),
    })
    order_id = int(order.headers["Location"].rsplit("/", 1)[-1])
    client.post(f"/ordenes-compra/{order_id}/items", data={
        "product_id": product_id, "quantity_requested": "3", "unit_cost": "12.50",
        "csrf_token": csrf_token(client),
    })
    client.post(f"/ordenes-compra/{order_id}/solicitar", data={"csrf_token": csrf_token(client)})
    client.post(f"/ordenes-compra/{order_id}/recibir", data={
        "received_1": "2", "movement_date": "2026-09-17",
        "csrf_token": csrf_token(client),
    })
    with app_module.app.app_context():
        db = app_module.get_db()
        row = db.execute(
            "SELECT stock, unit_cost FROM products WHERE id = ?", (product_id,)
        ).fetchone()
        movement = db.execute(
            "SELECT movement_date FROM movements WHERE product_id = ?", (product_id,)
        ).fetchone()
        cost_audit = db.execute(
            """SELECT details FROM audit_log
               WHERE action = 'update' AND entity_type = 'product' AND entity_id = ?
               ORDER BY id DESC LIMIT 1""",
            (product_id,),
        ).fetchone()
    assert row["unit_cost"] == 12.5
    assert row["stock"] == 7
    assert movement["movement_date"] == "2026-09-17"
    assert "Costo unitario actualizado" in cost_audit["details"]


def test_valorizacion_filtra_y_calcula_total(client):
    login(client)
    create_product(client, code="VAL-001", stock="4", unit_cost="3.25", category="Valorización")
    response = client.get("/reportes/valorizacion?category=Valorizaci%C3%B3n")
    assert response.status_code == 200
    assert b"13.00" in response.data
    assert b"VAL-001" in response.data


def test_exportaciones_valorizacion_conservan_filtros(client):
    login(client)
    create_product(
        client, code="VAL-MAT", name="Material filtrado", stock="4",
        unit_cost="3.25", category="Valorización", product_type="material",
    )
    create_product(
        client, code="VAL-TOOL", name="Herramienta excluida",
        category="Valorización", product_type="herramienta",
    )

    query = "?category=Valorizaci%C3%B3n&product_type=material"
    csv_response = client.get(f"/reportes/valorizacion/csv{query}")
    assert csv_response.status_code == 200
    assert csv_response.data.startswith(b"\xef\xbb\xbf")
    csv_rows = list(csv.reader(io.StringIO(csv_response.data.decode("utf-8-sig"))))
    assert csv_rows[1][1] == "VAL-MAT"
    assert all("VAL-TOOL" not in row for row in csv_rows)
    assert "valorizacion-inventario-Valorizacion-material.csv" in (
        csv_response.headers["Content-Disposition"]
    )

    excel_response = client.get(f"/reportes/valorizacion/excel{query}")
    assert excel_response.status_code == 200
    workbook = load_workbook(io.BytesIO(excel_response.data), read_only=True)
    sheet = workbook["Valorización"]
    assert sheet["B2"].value == "VAL-MAT"
    assert sheet["B3"].value is None
    assert "valorizacion-inventario-Valorizacion-material.xlsx" in (
        excel_response.headers["Content-Disposition"]
    )


def test_rechaza_costo_invalido(client):
    login(client)
    response = client.post("/productos", data={
        "code": "BAD-COST", "name": "Costo inválido", "category": "Pruebas",
        "product_type": "material", "presentation": "unidad", "stock": "0",
        "minimum_stock": "0", "reorder_quantity": "0", "unit_cost": "NaN",
        "csrf_token": csrf_token(client),
    })
    assert response.status_code == 302
    with app_module.app.app_context():
        assert app_module.get_db().execute(
            "SELECT COUNT(*) FROM products WHERE code = 'BAD-COST'"
        ).fetchone()[0] == 0


def test_crea_backup_sqlite_y_aplica_retencion(tmp_path, monkeypatch):
    database = tmp_path / "source.db"
    backup_dir = tmp_path / "backups"
    connection = app_module.sqlite3.connect(database)
    connection.execute("CREATE TABLE sample (value TEXT)")
    connection.execute("INSERT INTO sample VALUES ('ok')")
    connection.commit()
    connection.close()
    monkeypatch.setenv("PANOL_BACKUP_RETENTION", "2")

    created = [app_module.create_backup(database, backup_dir) for _ in range(3)]

    assert all(path.name.startswith("panol-") for path in created)
    assert len(list(backup_dir.glob("panol-*.db"))) == 2
    copied = app_module.sqlite3.connect(created[-1])
    assert copied.execute("SELECT value FROM sample").fetchone()[0] == "ok"
    copied.close()


def test_backup_manual_requiere_admin_y_descarga_copia(client):
    login(client)
    response = client.post(
        "/backup/manual",
        data={"csrf_token": csrf_token(client)},
    )
    assert response.status_code == 200
    assert response.mimetype == "application/octet-stream"
    assert response.data.startswith(b"SQLite format 3")


def test_proteccion_de_rutas_y_csrf_en_post(client):
    protegido = client.get("/productos")
    assert protegido.status_code == 302
    assert "/login" in protegido.headers["Location"]

    login(client)
    sin_csrf = client.post(
        "/productos",
        data={"code": "MAT-002", "name": "Sin token"},
    )
    assert sin_csrf.status_code == 400
    assert b"Solicitud no v" in sin_csrf.data


def test_alta_listado_y_filtro_de_productos_y_empleados(client):
    login(client)
    create_product(client)
    create_product(
        client,
        code="MAT-002",
        name="Guantes prueba",
        category="Limpieza",
        stock="2",
    )
    create_employee(client)
    create_employee(
        client, code="EMP-002", full_name="Bruno Filtro", legajo="LEG-002"
    )

    productos = client.get("/productos?q=Guantes&category=Limpieza")
    assert productos.status_code == 200
    assert b"Guantes prueba" in productos.data
    assert b"Cable prueba" not in productos.data

    empleados = client.get("/empleados?q=Bruno")
    assert empleados.status_code == 200
    assert b"Bruno Filtro" in empleados.data
    assert b"Ana Prueba" not in empleados.data


def test_movimientos_salida_devolucion_ajuste_y_filtros(client):
    login(client)
    product_id = create_product(client)
    employee_id = create_employee(client)
    today = date.today().isoformat()

    def movement(kind, quantity):
        response = client.post(
            "/movimientos",
            data={
                "csrf_token": csrf_token(client),
                "product_id[]": product_id,
                "product_type": "material",
                "employee_id": employee_id,
                "movement_type": kind,
                "quantity[]": quantity,
                "movement_date": today,
                "notes": f"nota {kind}",
            },
        )
        assert response.status_code == 302

    movement("salida", 2)
    movement("devolucion", 1)
    movement("ajuste", 1)

    with app_module.app.app_context():
        db = app_module.get_db()
        stock = db.execute(
            "SELECT stock FROM products WHERE id = ?", (product_id,)
        ).fetchone()["stock"]
        total = db.execute("SELECT COUNT(*) AS n FROM movements").fetchone()["n"]
    assert stock == 3
    assert total == 3

    listado = client.get("/movimientos?movement_type=salida&q=nota")
    assert listado.status_code == 200
    assert b"nota salida" in listado.data
    assert b"nota devolucion" not in listado.data


def test_transiciones_estrictas_de_unidad_individualizada(client):
    login(client)
    product_id = create_product(
        client,
        code="TOOL-001",
        name="Taladro prueba",
        product_type="herramienta",
        stock="1",
    )
    employee_id = create_employee(client)
    with app_module.app.app_context():
        db = app_module.get_db()
        instance_id = db.execute(
            "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)",
            (product_id, "SER-001"),
        ).lastrowid
        db.commit()

    def movement(kind, **extra):
        data = {
            "csrf_token": csrf_token(client),
            "product_id[]": product_id,
            "product_type": "herramienta",
            "employee_id": employee_id,
            "tool_instance_id[]": instance_id,
            "movement_type": kind,
            "quantity[]": 1,
            "movement_date": date.today().isoformat(),
        }
        data.update(extra)
        return client.post("/movimientos", data=data, follow_redirects=True)

    salida = movement("salida")
    assert salida.status_code == 200
    with app_module.app.app_context():
        instance = app_module.get_db().execute(
            "SELECT status, current_employee_id FROM tool_instances WHERE id = ?",
            (instance_id,),
        ).fetchone()
    assert (instance["status"], instance["current_employee_id"]) == (
        "entregada",
        employee_id,
    )

    segunda_salida = movement("salida")
    assert b"ya est\xc3\xa1 entregada" in segunda_salida.data
    assert movement("devolucion").status_code == 200
    with app_module.app.app_context():
        instance = app_module.get_db().execute(
            "SELECT status, current_employee_id FROM tool_instances WHERE id = ?",
            (instance_id,),
        ).fetchone()
    assert (instance["status"], instance["current_employee_id"]) == ("disponible", None)

    devolucion_sin_salida = movement("devolucion")
    assert b"no est\xc3\xa1 entregada" in devolucion_sin_salida.data
    cantidad_invalida = movement("salida", **{"quantity[]": 2})
    assert b"se mueven de a una unidad" in cantidad_invalida.data

    with app_module.app.app_context():
        db = app_module.get_db()
        db.execute(
            "UPDATE tool_instances SET status = 'reparacion' WHERE id = ?",
            (instance_id,),
        )
        db.commit()
    bloqueada = movement("salida")
    assert b"No se puede entregar" in bloqueada.data
    assert b"reparacion" in bloqueada.data


def test_unidad_individual_debe_pertenecer_al_producto(client):
    login(client)
    product_id = create_product(
        client,
        code="TOOL-002",
        name="Amoladora prueba",
        product_type="herramienta",
        stock="1",
    )
    other_product_id = create_product(
        client,
        code="TOOL-003",
        name="Sierra prueba",
        product_type="herramienta",
        stock="1",
    )
    employee_id = create_employee(client)
    with app_module.app.app_context():
        instance_id = app_module.get_db().execute(
            "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)",
            (other_product_id, "SER-002"),
        ).lastrowid
        app_module.get_db().commit()
    response = client.post(
        "/movimientos",
        data={
            "csrf_token": csrf_token(client),
            "product_id[]": product_id,
            "product_type": "herramienta",
            "employee_id": employee_id,
            "tool_instance_id[]": instance_id,
            "movement_type": "salida",
            "quantity[]": 1,
            "movement_date": date.today().isoformat(),
        },
        follow_redirects=True,
    )
    assert b"no pertenece al producto" in response.data


def test_historial_de_estados_registra_transicion_motivo_y_no_duplica(client):
    login(client)
    product_id = create_product(
        client,
        code="TOOL-HISTORY",
        name="Llave de historial",
        product_type="herramienta",
        stock="1",
    )
    with app_module.app.app_context():
        db = app_module.get_db()
        instance_id = db.execute(
            "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)",
            (product_id, "SER-HISTORY"),
        ).lastrowid
        db.commit()

    endpoint = f"/herramientas/{instance_id}/estado"
    missing_reason = client.post(
        endpoint,
        data={"csrf_token": csrf_token(client), "status": "perdida"},
        follow_redirects=True,
    )
    assert b"motivo" in missing_reason.data

    changed = client.post(
        endpoint,
        data={
            "csrf_token": csrf_token(client),
            "status": "perdida",
            "reason": "No fue localizada durante el inventario",
        },
        follow_redirects=False,
    )
    assert changed.status_code == 302
    unchanged = client.post(
        endpoint,
        data={
            "csrf_token": csrf_token(client),
            "status": "perdida",
            "reason": "Intento repetido",
        },
        follow_redirects=True,
    )
    assert b"ya se encuentra" in unchanged.data

    history = client.get(f"/herramientas/{instance_id}/historial")
    assert history.status_code == 200
    assert b"No fue localizada durante el inventario" in history.data
    with app_module.app.app_context():
        row = app_module.get_db().execute(
            """SELECT previous_status, new_status, notes, changed_by
               FROM tool_instance_status_history WHERE instance_id = ?""",
            (instance_id,),
        ).fetchall()
    assert len(row) == 1
    assert (row[0]["previous_status"], row[0]["new_status"]) == (
        "disponible",
        "perdida",
    )
    assert row[0]["changed_by"] is not None


def test_historial_de_herramienta_pagina_movimientos_y_estados_independientes(client):
    login(client)
    product_id = create_product(
        client,
        code="TOOL-PAGING",
        name="Herramienta paginada",
        product_type="herramienta",
        stock="1",
    )
    with app_module.app.app_context():
        db = app_module.get_db()
        instance_id = db.execute(
            "INSERT INTO tool_instances (product_id, serial_number) VALUES (?, ?)",
            (product_id, "SER-PAGING"),
        ).lastrowid
        for index in range(3):
            db.execute(
                """INSERT INTO movements
                   (product_id, tool_instance_id, movement_type, quantity, notes,
                    movement_date)
                   VALUES (?, ?, 'entrada', 1, ?, ?)""",
                (product_id, instance_id, f"Movimiento {index}", f"2026-09-{index + 1:02d}"),
            )
            db.execute(
                """INSERT INTO tool_instance_status_history
                   (instance_id, previous_status, new_status, notes)
                   VALUES (?, ?, ?, ?)""",
                (instance_id, "disponible", "reparacion", f"Cambio {index}"),
            )
        db.commit()

    first = client.get(
        f"/herramientas/{instance_id}/historial"
        "?movement_page=1&status_page=1&page_size=2"
    )
    assert first.status_code == 200
    assert b"Movimiento 2" in first.data
    assert b"Movimiento 0" not in first.data
    assert b"Cambio 2" in first.data
    assert b"Cambio 0" not in first.data
    assert b"movement_page=2" in first.data
    assert b"status_page=2" in first.data

    second_movements = client.get(
        f"/herramientas/{instance_id}/historial"
        "?movement_page=2&status_page=1&page_size=2"
    )
    assert b"Movimiento 0" in second_movements.data
    assert b"Movimiento 2" not in second_movements.data
    assert b"Cambio 2" in second_movements.data


def test_herramientas_filtra_por_texto_estado_responsable_y_pagina(client):
    login(client)
    product_id = create_product(
        client,
        code="TOOL-FILTER",
        name="Taladro filtrable",
        product_type="herramienta",
        stock="3",
    )
    employee_id = create_employee(
        client, code="EMP-TOOL", full_name="Responsable herramienta"
    )
    with app_module.app.app_context():
        db = app_module.get_db()
        db.executemany(
            """INSERT INTO tool_instances
               (product_id, serial_number, status, current_employee_id)
               VALUES (?, ?, ?, ?)""",
            [
                (product_id, "SER-DISPONIBLE", "disponible", None),
                (product_id, "SER-ENTREGADA", "entregada", employee_id),
                (product_id, "SER-REPARACION", "reparacion", None),
            ],
        )
        db.commit()

    text_match = client.get("/herramientas?q=SER-ENTREGADA")
    assert text_match.status_code == 200
    assert b"SER-ENTREGADA" in text_match.data
    assert b"SER-DISPONIBLE" not in text_match.data

    status_match = client.get("/herramientas?status=reparacion")
    assert b"SER-REPARACION" in status_match.data
    assert b"SER-ENTREGADA" not in status_match.data

    employee_match = client.get(f"/herramientas?employee_id={employee_id}")
    assert b"SER-ENTREGADA" in employee_match.data
    assert b"SER-DISPONIBLE" not in employee_match.data

    page = client.get(
        "/herramientas?q=Taladro&status=disponible&page=1&page_size=1"
    )
    assert b"SER-DISPONIBLE" in page.data
    assert b'name="q" value="Taladro"' in page.data
    assert b'name="page_size" value="1"' in page.data


def test_comprobante_pdf_y_historial_inexistente(client):
    login(client)
    product_id = create_product(client)
    employee_id = create_employee(client)
    response = client.post(
        "/movimientos",
        data={
            "csrf_token": csrf_token(client),
            "product_id[]": product_id,
            "product_type": "material",
            "employee_id": employee_id,
            "movement_type": "salida",
            "quantity[]": 1,
            "movement_date": date.today().isoformat(),
        },
    )
    assert response.status_code == 302
    with app_module.app.app_context():
        movement_id = app_module.get_db().execute(
            "SELECT id FROM movements ORDER BY id DESC LIMIT 1"
        ).fetchone()["id"]

    receipt = client.get(f"/movimientos/{movement_id}/comprobante")
    assert receipt.status_code == 200
    assert receipt.mimetype == "application/pdf"
    assert receipt.data.startswith(b"%PDF")

    missing = client.get("/herramientas/99999/historial")
    assert missing.status_code == 302
    assert missing.headers["Location"].endswith("/herramientas")


def test_exportaciones_csv_y_excel(client):
    login(client)
    create_product(client)
    create_employee(client)

    csv_response = client.get("/reportes/csv")
    assert csv_response.status_code == 200
    assert csv_response.mimetype == "text/csv"
    assert b"Inventario" in csv_response.data

    excel_response = client.get("/reportes/excel")
    assert excel_response.status_code == 200
    assert excel_response.mimetype == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert excel_response.data[:2] == b"PK"


def test_paginacion_conserva_filtros_y_limita_tamano(client):
    login(client)
    with app_module.app.app_context():
        db = app_module.get_db()
        db.executemany(
            """INSERT INTO products
               (code, name, category, product_type, presentation)
               VALUES (?, ?, ?, 'material', 'unidad')""",
            [
                (f"P-{i:03d}", f"Producto {i:03d}", "Electricidad" if i < 30 else "Limpieza")
                for i in range(30)
            ],
        )
        db.executemany(
            "INSERT INTO employees (code, full_name, legajo) VALUES (?, ?, ?)",
            [(f"E-{i:03d}", f"Empleado {i:03d}", f"L-{i:03d}") for i in range(30)],
        )
        product_id = db.execute("SELECT id FROM products ORDER BY id LIMIT 1").fetchone()["id"]
        employee_id = db.execute("SELECT id FROM employees ORDER BY id LIMIT 1").fetchone()["id"]
        db.executemany(
            """INSERT INTO movements
               (product_id, employee_id, movement_type, quantity, notes, movement_date)
               VALUES (?, ?, 'salida', 1, ?, CURRENT_DATE)""",
            [(product_id, employee_id, f"Nota {i:03d}") for i in range(30)],
        )
        db.commit()

    products_page = client.get("/productos?q=Producto&category=Electricidad&page=2&page_size=10")
    assert products_page.status_code == 200
    assert b"Producto 010" in products_page.data
    assert b"Producto 000" not in products_page.data
    assert b"page=1" in products_page.data

    employees_page = client.get("/empleados?q=Empleado&page=2&page_size=10")
    assert employees_page.status_code == 200
    assert b"Empleado 010" in employees_page.data
    assert b"Empleado 000" not in employees_page.data

    movements_page = client.get(
        "/movimientos?q=Nota&movement_type=salida&employee_id="
        f"{employee_id}&page=2&page_size=10"
    )
    assert movements_page.status_code == 200
    assert b"Nota 010" in movements_page.data
    assert b"Nota 000" not in movements_page.data

    bounded = client.get("/empleados?q=Empleado&page=1&page_size=9999")
    assert bounded.status_code == 200
    assert b'name="page_size" value="100"' in bounded.data


def test_orden_compra_recibe_parcial_y_bloquea_sobre_recepcion(client):
    login(client)
    product_id = create_product(client, code="OC-MAT-001", name="Material para compra", stock="0")
    with app_module.app.app_context():
        db = app_module.get_db()
        supplier_id = db.execute(
            "INSERT INTO suppliers (name, active) VALUES (?, 1)",
            ("Proveedor de prueba",),
        ).lastrowid
        db.commit()

    created = client.post(
        "/ordenes-compra",
        data={
            "csrf_token": csrf_token(client),
            "supplier_id": supplier_id,
            "order_date": date.today().isoformat(),
            "notes": "Compra de prueba",
        },
    )
    assert created.status_code == 302
    order_id = int(created.headers["Location"].rsplit("/", 1)[-1])

    item = client.post(
        f"/ordenes-compra/{order_id}/items",
        data={
            "csrf_token": csrf_token(client),
            "product_id": product_id,
            "quantity_requested": "5",
            "unit_cost": "10.50",
        },
    )
    assert item.status_code == 302
    requested = client.post(
        f"/ordenes-compra/{order_id}/solicitar",
        data={"csrf_token": csrf_token(client)},
    )
    assert requested.status_code == 302

    with app_module.app.app_context():
        item_id = app_module.get_db().execute(
            "SELECT id FROM purchase_order_items WHERE purchase_order_id = ?",
            (order_id,),
        ).fetchone()["id"]

    partial = client.post(
        f"/ordenes-compra/{order_id}/recibir",
        data={
            "csrf_token": csrf_token(client),
            f"received_{item_id}": "2",
        },
    )
    assert partial.status_code == 302
    with app_module.app.app_context():
        db = app_module.get_db()
        order = db.execute(
            "SELECT status FROM purchase_orders WHERE id = ?", (order_id,)
        ).fetchone()
        stock = db.execute(
            "SELECT stock FROM products WHERE id = ?", (product_id,)
        ).fetchone()["stock"]
    assert order["status"] == "parcial"
    assert stock == 2

    remaining = client.post(
        f"/ordenes-compra/{order_id}/recibir",
        data={
            "csrf_token": csrf_token(client),
            f"received_{item_id}": "3",
        },
    )
    assert remaining.status_code == 302
    with app_module.app.app_context():
        db = app_module.get_db()
        order = db.execute(
            "SELECT status FROM purchase_orders WHERE id = ?", (order_id,)
        ).fetchone()
        stock = db.execute(
            "SELECT stock FROM products WHERE id = ?", (product_id,)
        ).fetchone()["stock"]
    assert order["status"] == "recibida"
    assert stock == 5

    duplicate = client.post(
        f"/ordenes-compra/{order_id}/recibir",
        data={
            "csrf_token": csrf_token(client),
            f"received_{item_id}": "1",
        },
        follow_redirects=True,
    )
    assert b"No se puede recibir" in duplicate.data


def test_ordenes_compra_filtra_por_estado_y_proveedor(client):
    login(client)
    with app_module.app.app_context():
        db = app_module.get_db()
        supplier_a = db.execute(
            "INSERT INTO suppliers (name, active) VALUES (?, 1)", ("Proveedor A",)
        ).lastrowid
        supplier_b = db.execute(
            "INSERT INTO suppliers (name, active) VALUES (?, 1)", ("Proveedor B",)
        ).lastrowid
        db.executemany(
            """INSERT INTO purchase_orders
               (order_number, supplier_id, status, notes, created_by)
               VALUES (?, ?, ?, ?, 1)""",
            [
                ("OC-FILTRO-001", supplier_a, "solicitada", "Urgente",),
                ("OC-FILTRO-002", supplier_b, "borrador", "Normal",),
            ],
        )
        db.commit()

    response = client.get(
        f"/ordenes-compra?status=solicitada&supplier_id={supplier_a}&q=Urgente"
    )
    assert response.status_code == 200
    assert b"OC-FILTRO-001" in response.data
    assert b"OC-FILTRO-002" not in response.data
    assert b'name="status"' in response.data


def test_ordenes_compra_pagina_y_conserva_filtros(client):
    login(client)
    with app_module.app.app_context():
        db = app_module.get_db()
        supplier_id = db.execute(
            "INSERT INTO suppliers (name, active) VALUES (?, 1)", ("Proveedor páginas",)
        ).lastrowid
        db.executemany(
            """INSERT INTO purchase_orders
               (order_number, supplier_id, status, notes, created_by)
               VALUES (?, ?, 'borrador', ?, 1)""",
            [
                ("OC-PAG-001", supplier_id, "Nota 1"),
                ("OC-PAG-002", supplier_id, "Nota 2"),
            ],
        )
        db.commit()

    response = client.get(
        f"/ordenes-compra?status=borrador&supplier_id={supplier_id}"
        "&page=2&page_size=1"
    )
    assert response.status_code == 200
    assert b"OC-PAG-001" in response.data
    assert b"status=borrador" in response.data
    assert f"supplier_id={supplier_id}".encode() in response.data
