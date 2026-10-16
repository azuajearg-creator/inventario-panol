"""
Recordatorio de devoluciones pendientes — Inventario Paniol.
Notifica herramientas y materiales retornables con mas de N dias sin devolver.
"""
import os
import sqlite3
import subprocess
import smtplib
import sys
from configparser import ConfigParser
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

SCRIPT_DIR = Path(os.environ.get("PANOL_HOME") or Path(__file__).resolve().parent.parent)
DATABASE    = SCRIPT_DIR / "data" / "panol.db"
CONFIG_FILE = SCRIPT_DIR / "data" / "alerta_config.ini"
LOG_FILE    = SCRIPT_DIR / "data" / "devoluciones.log"

CREATE_NO_WINDOW = 0x08000000


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


def load_config() -> ConfigParser:
    cfg = ConfigParser()
    cfg["email"] = {
        "sender": "",
        "recipient": "",
        "app_password": "",
        "smtp_host": "smtp.gmail.com",
        "smtp_port": "587",
    }
    cfg["devoluciones"] = {"dias_umbral": "3"}
    if CONFIG_FILE.exists():
        cfg.read(CONFIG_FILE, encoding="utf-8")
    return cfg


def query_pendientes(db: sqlite3.Connection, umbral: int) -> dict:
    # Herramientas con numero de serie (instancia individual)
    instancias = db.execute(
        """SELECT ti.id, ti.serial_number,
                  p.name AS product_name, p.code AS product_code,
                  e.full_name, e.legajo,
                  last_m.movement_date AS delivery_date,
                  CAST(julianday('now') - julianday(
                      COALESCE(last_m.movement_date, '1900-01-01')
                  ) AS INTEGER) AS days_out
           FROM tool_instances ti
           JOIN products p ON p.id = ti.product_id
           LEFT JOIN employees e ON e.id = ti.current_employee_id
           LEFT JOIN (
               SELECT tool_instance_id, MAX(movement_date) AS movement_date
               FROM movements WHERE movement_type = 'salida'
               GROUP BY tool_instance_id
           ) last_m ON last_m.tool_instance_id = ti.id
           WHERE ti.status = 'entregada'
             AND CAST(julianday('now') - julianday(
                     COALESCE(last_m.movement_date, '1900-01-01')
                 ) AS INTEGER) >= ?
           ORDER BY days_out DESC, e.full_name, p.name""",
        (umbral,),
    ).fetchall()

    # Herramientas sin instancia (por cantidad)
    herramientas_qty = db.execute(
        """SELECT e.full_name, e.legajo,
                  p.name AS product_name, p.code AS product_code, p.unit,
                  SUM(CASE WHEN m.movement_type = 'salida' THEN m.quantity
                           WHEN m.movement_type = 'devolucion' THEN -m.quantity
                           ELSE 0 END) AS pending_quantity,
                  MAX(m.movement_date) AS last_delivery,
                  CAST(julianday('now') - julianday(MAX(m.movement_date)) AS INTEGER) AS days_out
           FROM movements m
           JOIN employees e ON e.id = m.employee_id
           JOIN products p ON p.id = m.product_id
           WHERE m.movement_type IN ('salida', 'devolucion')
             AND p.product_type = 'herramienta'
             AND m.tool_instance_id IS NULL
           GROUP BY e.id, p.id
           HAVING pending_quantity > 0
              AND CAST(julianday('now') - julianday(MAX(m.movement_date)) AS INTEGER) >= ?
           ORDER BY days_out DESC, e.full_name, p.name""",
        (umbral,),
    ).fetchall()

    # Materiales retornables
    materiales = db.execute(
        """SELECT e.full_name, e.legajo,
                  p.name AS product_name, p.code AS product_code, p.unit,
                  SUM(CASE WHEN m.movement_type = 'salida' THEN m.quantity
                           WHEN m.movement_type = 'devolucion' THEN -m.quantity
                           ELSE 0 END) AS pending_quantity,
                  MAX(m.movement_date) AS last_delivery,
                  CAST(julianday('now') - julianday(MAX(m.movement_date)) AS INTEGER) AS days_out
           FROM movements m
           JOIN employees e ON e.id = m.employee_id
           JOIN products p ON p.id = m.product_id
           WHERE m.movement_type IN ('salida', 'devolucion')
             AND p.product_type = 'material'
             AND p.retornable = 1
           GROUP BY e.id, p.id
           HAVING pending_quantity > 0
              AND CAST(julianday('now') - julianday(MAX(m.movement_date)) AS INTEGER) >= ?
           ORDER BY days_out DESC, e.full_name, p.name""",
        (umbral,),
    ).fetchall()

    return {
        "instancias":       instancias,
        "herramientas_qty": herramientas_qty,
        "materiales":       materiales,
    }


def show_notification(title: str, body: str) -> None:
    t = title.replace('"', '`"')
    b = body.replace('"', '`"')
    ps = (
        'Add-Type -AssemblyName System.Windows.Forms; '
        '$n = New-Object System.Windows.Forms.NotifyIcon; '
        '$n.Icon = [System.Drawing.SystemIcons]::Warning; '
        '$n.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning; '
        f'$n.BalloonTipTitle = "{t}"; '
        f'$n.BalloonTipText  = "{b}"; '
        '$n.Visible = $True; '
        '$n.ShowBalloonTip(10000); '
        'Start-Sleep -Seconds 12; '
        '$n.Dispose()'
    )
    try:
        subprocess.Popen(
            ["powershell", "-WindowStyle", "Hidden", "-Command", ps],
            creationflags=CREATE_NO_WINDOW,
        )
    except Exception as exc:
        log(f"WARN notificacion Windows fallida: {exc}")


def _filas_html(rows, cols: list) -> str:
    return "".join(
        "<tr>" + "".join(f"<td>{r[c]}</td>" for c in cols) + "</tr>"
        for r in rows
    )


def send_email(cfg: ConfigParser, pendientes: dict, umbral: int) -> None:
    app_password = cfg.get("email", "app_password", fallback="").strip()
    if not app_password:
        log("WARN email omitido: completar app_password en data\\alerta_config.ini")
        return

    sender    = cfg.get("email", "sender")
    recipient = cfg.get("email", "recipient")
    host      = cfg.get("email", "smtp_host")
    port      = cfg.getint("email", "smtp_port")
    fecha     = datetime.now().strftime("%d/%m/%Y %H:%M")

    inst = pendientes["instancias"]
    hqty = pendientes["herramientas_qty"]
    mats = pendientes["materiales"]
    total = len(inst) + len(hqty) + len(mats)

    bloques = []
    if inst:
        bloques.append(
            f"<h3>Herramientas con numero de serie ({len(inst)})</h3>"
            "<table><tr><th>Empleado</th><th>Legajo</th><th>Herramienta</th>"
            "<th>N/S</th><th>Fecha salida</th><th>Dias fuera</th></tr>"
            + _filas_html(inst, ["full_name", "legajo", "product_name",
                                 "serial_number", "delivery_date", "days_out"])
            + "</table>"
        )
    if hqty:
        bloques.append(
            f"<h3>Herramientas por cantidad ({len(hqty)})</h3>"
            "<table><tr><th>Empleado</th><th>Legajo</th><th>Herramienta</th>"
            "<th>Cant.</th><th>Unidad</th><th>Ult. salida</th><th>Dias fuera</th></tr>"
            + _filas_html(hqty, ["full_name", "legajo", "product_name",
                                  "pending_quantity", "unit", "last_delivery", "days_out"])
            + "</table>"
        )
    if mats:
        bloques.append(
            f"<h3>Materiales retornables ({len(mats)})</h3>"
            "<table><tr><th>Empleado</th><th>Legajo</th><th>Material</th>"
            "<th>Cant.</th><th>Unidad</th><th>Ult. salida</th><th>Dias fuera</th></tr>"
            + _filas_html(mats, ["full_name", "legajo", "product_name",
                                  "pending_quantity", "unit", "last_delivery", "days_out"])
            + "</table>"
        )

    html = (
        "<!DOCTYPE html><html lang='es'><head><meta charset='utf-8'>"
        "<style>"
        "body{font-family:Arial,sans-serif;font-size:14px;color:#222}"
        "h2{color:#c0392b}h3{color:#2c3e50;margin-top:24px}"
        "table{border-collapse:collapse;width:100%;margin-bottom:16px}"
        "th{background:#2c3e50;color:#fff;padding:8px 10px;text-align:left}"
        "td{padding:7px 10px;border-bottom:1px solid #ddd}"
        "tr:nth-child(even) td{background:#f9f9f9}"
        "</style></head><body>"
        "<h2>Devoluciones Pendientes - Paniol</h2>"
        f"<p>Fecha: <strong>{fecha}</strong> | "
        f"Umbral: <strong>{umbral} dias</strong> | "
        f"Items: <strong>{total}</strong></p>"
        + "".join(bloques)
        + "<p style='color:#777;font-size:12px;margin-top:20px'>"
          "Generado automaticamente por Inventario Paniol.</p>"
          "</body></html>"
    )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Paniol: {total} devolucion(es) pendiente(s) (+{umbral} dias) - {fecha}"
    msg["From"]    = sender
    msg["To"]      = recipient
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP(host, port, timeout=20) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender, app_password)
            smtp.sendmail(sender, [recipient], msg.as_bytes())
        log(f"OK  email enviado a {recipient} ({total} item(s))")
    except Exception as exc:
        log(f"ERROR email: {exc}")


if __name__ == "__main__":
    if not DATABASE.exists():
        log(f"ERROR: Base de datos no encontrada en {DATABASE}")
        sys.exit(1)

    cfg    = load_config()
    umbral = cfg.getint("devoluciones", "dias_umbral", fallback=3)

    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    try:
        pendientes = query_pendientes(db, umbral)
    finally:
        db.close()

    inst  = pendientes["instancias"]
    hqty  = pendientes["herramientas_qty"]
    mats  = pendientes["materiales"]
    total = len(inst) + len(hqty) + len(mats)

    if total == 0:
        log(f"OK  sin devoluciones pendientes de mas de {umbral} dias")
        sys.exit(0)

    log(f"RECORDATORIO  {total} devolucion(es) con mas de {umbral} dias")

    empleados = sorted({r["full_name"] for r in list(inst) + list(hqty) + list(mats)})
    nombres   = ", ".join(empleados[:3])
    sufijo    = f" (+{len(empleados) - 3} mas)" if len(empleados) > 3 else ""
    show_notification(
        f"Paniol: {total} devolucion(es) pendiente(s)",
        f"{nombres}{sufijo}",
    )

    send_email(cfg, pendientes, umbral)
