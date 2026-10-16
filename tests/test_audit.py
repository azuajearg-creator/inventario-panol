"""Pruebas aisladas de filtros y paginación de auditoría."""

import os
import sys
from pathlib import Path

import pytest

os.environ["PANOL_SKIP_INIT_DB"] = "1"
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app as app_module


@pytest.fixture()
def audit_client(tmp_path, monkeypatch):
    database = tmp_path / "audit-test.db"
    monkeypatch.setattr(app_module, "DATABASE", database)
    app_module.app.config.update(TESTING=True, DATABASE=database)
    with app_module.app.app_context():
        app_module.init_db()
        db = app_module.get_db()
        user_id = db.execute(
            "SELECT id FROM users WHERE username = 'admin'"
        ).fetchone()["id"]
        db.executemany(
            """INSERT INTO audit_log
               (user_id, action, entity_type, details)
               VALUES (?, ?, ?, ?)""",
            [
                (user_id, "create", "product", "Producto cable creado"),
                (user_id, "update", "employee", "Ana actualizada"),
                (user_id, "delete", "product", "Producto obsoleto eliminado"),
            ],
        )
        db.commit()
    with app_module.app.test_client() as client:
        client.get("/login")
        with client.session_transaction() as session:
            csrf_token = session["csrf_token"]
        response = client.post(
            "/login",
            data={"username": "admin", "password": "admin1234", "csrf_token": csrf_token},
        )
        assert response.status_code == 302
        yield client


def test_auditoria_filtra_por_accion_entidad_y_texto(audit_client):
    response = audit_client.get("/auditoria?action=update&entity=employee&q=Ana")

    assert response.status_code == 200
    assert b"Ana actualizada" in response.data
    assert b"Producto cable creado" not in response.data
    assert b"Producto obsoleto eliminado" not in response.data


def test_auditoria_pagina_y_preserva_filtros(audit_client):
    response = audit_client.get("/auditoria?entity=product&page=2&page_size=1")

    assert response.status_code == 200
    assert b"Producto cable creado" in response.data
    assert b"Producto obsoleto eliminado" not in response.data
    assert b"entity=product" in response.data
    assert b"page_size=1" in response.data
