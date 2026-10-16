"""Punto de entrada para ejecutar la aplicación con Waitress en producción."""

from __future__ import annotations

import os
from collections.abc import Mapping

from app import app

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8000
DEFAULT_THREADS = 4
MAX_THREADS = 1000


def _positive_integer(value: str, name: str, maximum: int) -> int:
    """Convierte una variable de entorno en un entero dentro de un rango seguro."""
    try:
        parsed = int(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{name} debe ser un entero") from error
    if not 1 <= parsed <= maximum:
        raise ValueError(f"{name} debe estar entre 1 y {maximum}")
    return parsed


def server_config(environ: Mapping[str, str] | None = None) -> dict[str, int | str]:
    """Devuelve la configuración validada de Waitress sin iniciar un servidor."""
    environment = os.environ if environ is None else environ
    host = environment.get("PANOL_HOST", DEFAULT_HOST).strip() or DEFAULT_HOST
    port = _positive_integer(
        environment.get("PANOL_PORT", str(DEFAULT_PORT)), "PANOL_PORT", 65535
    )
    threads = _positive_integer(
        environment.get("PANOL_THREADS", str(DEFAULT_THREADS)),
        "PANOL_THREADS",
        MAX_THREADS,
    )
    return {"host": host, "port": port, "threads": threads}


def main() -> None:
    """Inicia la aplicación Flask usando el servidor WSGI Waitress."""
    import warnings
    from waitress import serve

    if not os.environ.get("PANOL_SECRET_KEY"):
        warnings.warn(
            "PANOL_SECRET_KEY no está configurada; se está utilizando la clave por defecto. "
            "Configure PANOL_SECRET_KEY en entornos de producción para mayor seguridad.",
            stacklevel=2,
        )
    serve(app, **server_config())


if __name__ == "__main__":
    main()
