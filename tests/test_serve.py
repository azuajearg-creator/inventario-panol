"""Pruebas de configuración del punto de entrada de producción."""

import os

os.environ["PANOL_SKIP_INIT_DB"] = "1"

import pytest

from serve import DEFAULT_HOST, DEFAULT_PORT, DEFAULT_THREADS, server_config


def test_server_config_usa_valores_seguros_por_defecto():
    assert server_config({}) == {
        "host": DEFAULT_HOST,
        "port": DEFAULT_PORT,
        "threads": DEFAULT_THREADS,
    }


def test_server_config_lee_variables_de_entorno():
    assert server_config(
        {"PANOL_HOST": "192.168.1.10", "PANOL_PORT": "9000", "PANOL_THREADS": "8"}
    ) == {"host": "192.168.1.10", "port": 9000, "threads": 8}


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("PANOL_PORT", "0"),
        ("PANOL_PORT", "65536"),
        ("PANOL_PORT", "no-es-un-numero"),
        ("PANOL_THREADS", "0"),
        ("PANOL_THREADS", "1001"),
        ("PANOL_THREADS", "no-es-un-numero"),
    ],
)
def test_server_config_rechaza_valores_invalidos(name, value):
    with pytest.raises(ValueError, match=name):
        server_config({name: value})
