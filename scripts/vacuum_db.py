"""
Mantenimiento semanal de la BD SQLite — Inventario Paniol.
Ejecuta VACUUM para compactar el archivo y recuperar espacio libre.
"""
import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(os.environ.get("PANOL_HOME") or Path(__file__).resolve().parent.parent)
DATABASE   = SCRIPT_DIR / "data" / "panol.db"
LOG_FILE   = SCRIPT_DIR / "data" / "vacuum.log"


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


if __name__ == "__main__":
    if not DATABASE.exists():
        log(f"ERROR: Base de datos no encontrada en {DATABASE}")
        sys.exit(1)

    size_before = DATABASE.stat().st_size

    try:
        db = sqlite3.connect(DATABASE, timeout=10)
        try:
            db.execute("VACUUM")
            db.execute("ANALYZE")
            db.commit()
        finally:
            db.close()
    except sqlite3.OperationalError as exc:
        log(f"ERROR: la BD esta bloqueada o no accesible — {exc}")
        sys.exit(1)

    size_after = DATABASE.stat().st_size
    freed_kb   = max(0, (size_before - size_after) // 1024)
    size_kb    = size_after // 1024

    log(f"OK  VACUUM+ANALYZE completado — tamano actual: {size_kb} KB, liberados: {freed_kb} KB")
