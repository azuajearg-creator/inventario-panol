"""El usuario administrador inicial no debe tener una contraseña fija conocida."""

import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app as app_module
from werkzeug.security import check_password_hash

MARCADOR = "Contraseña generada (guárdela y cámbiela): "


def test_sin_variable_se_genera_una_contrasena_aleatoria(tmp_path, monkeypatch, capsys):
    monkeypatch.delenv("PANOL_ADMIN_PASSWORD", raising=False)
    monkeypatch.delenv("PANOL_ADMIN_USER", raising=False)
    database = tmp_path / "panol-test.db"
    monkeypatch.setattr(app_module, "DATABASE", database)
    app_module.app.config.update(TESTING=True, DATABASE=database)

    with app_module.app.app_context():
        app_module.init_db()

    salida = capsys.readouterr().err
    assert MARCADOR in salida
    generada = salida.split(MARCADOR, 1)[1].strip()
    assert len(generada) >= 12

    conexion = sqlite3.connect(database)
    (hash_guardado,) = conexion.execute(
        "SELECT password_hash FROM users WHERE username = 'admin'"
    ).fetchone()
    conexion.close()
    assert check_password_hash(hash_guardado, generada)
    # Ninguna de las contraseñas fijas históricas debe funcionar.
    for antigua in ("admin1234", "AdminSeguro2024", "AdminAdmin"):
        assert not check_password_hash(hash_guardado, antigua)


def test_con_variable_se_usa_la_contrasena_configurada(tmp_path, monkeypatch, capsys):
    monkeypatch.setenv("PANOL_ADMIN_PASSWORD", "clave-de-prueba-123")
    database = tmp_path / "panol-test2.db"
    monkeypatch.setattr(app_module, "DATABASE", database)
    app_module.app.config.update(TESTING=True, DATABASE=database)

    with app_module.app.app_context():
        app_module.init_db()

    assert MARCADOR not in capsys.readouterr().err
    conexion = sqlite3.connect(database)
    (hash_guardado,) = conexion.execute(
        "SELECT password_hash FROM users WHERE username = 'admin'"
    ).fetchone()
    conexion.close()
    assert check_password_hash(hash_guardado, "clave-de-prueba-123")
