"""
Backup autónomo para Inventario Pañol.
Se ejecuta desde el Programador de tareas de Windows.
No requiere que la aplicación esté corriendo.
"""
import sqlite3
import os
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(os.environ.get("PANOL_HOME") or Path(__file__).resolve().parent.parent)
DATABASE    = SCRIPT_DIR / "data" / "panol.db"
BACKUP_DIR  = SCRIPT_DIR / "data" / "backups"
LOG_FILE    = SCRIPT_DIR / "data" / "backup.log"
RETENTION   = max(1, int(os.environ.get("PANOL_BACKUP_RETENTION", 30)))


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


def create_backup() -> tuple[Path, int]:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp   = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = BACKUP_DIR / f"panol-{timestamp}.db"

    src = sqlite3.connect(DATABASE)
    dst = sqlite3.connect(backup_path)
    try:
        src.backup(dst)
        dst.commit()
    except Exception:
        backup_path.unlink(missing_ok=True)
        raise
    finally:
        src.close()
        dst.close()

    # Retención: conservar solo los últimos RETENTION backups
    managed = sorted(
        BACKUP_DIR.glob("panol-*.db"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    deleted = 0
    for old in managed[RETENTION:]:
        old.unlink()
        deleted += 1

    return backup_path, deleted


if __name__ == "__main__":
    if not DATABASE.exists():
        log(f"ERROR: Base de datos no encontrada en {DATABASE}")
        sys.exit(1)

    try:
        backup_path, deleted = create_backup()
        size_kb = backup_path.stat().st_size // 1024
        msg = f"OK  {backup_path.name}  ({size_kb} KB)"
        if deleted:
            msg += f"  — {deleted} backup(s) antiguo(s) eliminado(s)"
        log(msg)
    except Exception as exc:
        log(f"ERROR: {exc}")
        sys.exit(1)
