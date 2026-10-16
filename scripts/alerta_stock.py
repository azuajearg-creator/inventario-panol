"""
Alerta automática de stock mínimo — Inventario Pañol.
Consulta la BD, muestra notificación Windows y envía email.
"""
import sqlite3
import os
import sys
import subprocess
import smtplib
from configparser import ConfigParser
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

SCRIPT_DIR = Path(os.environ.get("PANOL_HOME") or Path(__file__).resolve().parent.parent)
DATABASE    = SCRIPT_DIR / "data" / "panol.db"
CONFIG_FILE = SCRIPT_DIR / "data" / "alerta_config.ini"
LOG_FILE    = SCRIPT_DIR / "data" / "alerta_stock.log"

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
        "sender":       "",
        "recipient":    "",
        "app_password": "",
        "smtp_host":    "smtp.gmail.com",
        "smtp_port":    "587",
    }
    if CONFIG_FILE.exists():
        cfg.read(CONFIG_FILE, encoding="utf-8")
    return cfg


def query_stock_bajo(db: sqlite3.Connection) -> list:
    return db.execute(
        """SELECT p.name, p.code, p.category, p.product_type,
                  p.stock, p.minimum_stock,
                  CASE WHEN p.reorder_quantity > 0 THEN p.reorder_quantity
                       ELSE MAX(p.minimum_stock - p.stock, 1) END AS suggested_quantity,
                  s.name AS supplier_name
           FROM products p
           LEFT JOIN suppliers s ON s.id = p.supplier_id
           WHERE p.active = 1 AND p.stock <= p.minimum_stock
           ORDER BY (p.minimum_stock - p.stock) DESC, p.name"""
    ).fetchall()


def show_notification(title: str, body: str) -> None:
    """Muestra un balloon tip en la bandeja del sistema de Windows."""
    t = title.replace('"', '`"').replace("'", "`'")
    b = body.replace('"',  '`"').replace("'", "`'")
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


def send_email(cfg: ConfigParser, rows: list) -> None:
    app_password = cfg.get("email", "app_password", fallback="").strip()
    if not app_password:
        log("WARN email omitido: completar app_password en data\\alerta_config.ini")
        return

    sender    = cfg.get("email", "sender")
    recipient = cfg.get("email", "recipient")
    host      = cfg.get("email", "smtp_host")
    port      = cfg.getint("email", "smtp_port")
    fecha     = datetime.now().strftime("%d/%m/%Y %H:%M")

    filas_html = "".join(
        f"<tr>"
        f"<td>{r['name']}</td>"
        f"<td>{r['code'] or '—'}</td>"
        f"<td>{r['category'] or '—'}</td>"
        f"<td style=\"color:{'#c0392b' if r['stock'] == 0 else '#e67e22'};font-weight:bold\">{r['stock']}</td>"
        f"<td>{r['minimum_stock']}</td>"
        f"<td>{r['suggested_quantity']}</td>"
        f"<td>{r['supplier_name'] or '—'}</td>"
        f"</tr>"
        for r in rows
    )

    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8">
<style>
body  {{font-family:Arial,sans-serif;font-size:14px;color:#222}}
h2    {{color:#c0392b}}
table {{border-collapse:collapse;width:100%}}
th    {{background:#2c3e50;color:#fff;padding:8px 10px;text-align:left}}
td    {{padding:7px 10px;border-bottom:1px solid #ddd}}
tr:nth-child(even) td {{background:#f9f9f9}}
</style>
</head>
<body>
<h2>&#9888; Alerta de Stock M&#237;nimo &mdash; Pa&#241;ol</h2>
<p>Fecha: <strong>{fecha}</strong> &nbsp;|&nbsp;
   Productos afectados: <strong>{len(rows)}</strong></p>
<table>
  <tr>
    <th>Producto</th><th>C&#243;digo</th><th>Categor&#237;a</th>
    <th>Stock actual</th><th>M&#237;nimo</th>
    <th>Sugerido pedir</th><th>Proveedor</th>
  </tr>
  {filas_html}
</table>
<p style="color:#777;font-size:12px;margin-top:20px">
  Generado autom&#225;ticamente por Inventario Pa&#241;ol.
</p>
</body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Paniol: {len(rows)} producto(s) bajo stock minimo - {fecha}"
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
        log(f"OK  email enviado a {recipient} ({len(rows)} producto(s))")
    except Exception as exc:
        log(f"ERROR email: {exc}")


if __name__ == "__main__":
    if not DATABASE.exists():
        log(f"ERROR: Base de datos no encontrada en {DATABASE}")
        sys.exit(1)

    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    try:
        rows = query_stock_bajo(db)
    finally:
        db.close()

    if not rows:
        log("OK  sin productos bajo stock minimo — no se envia alerta")
        sys.exit(0)

    log(f"ALERTA  {len(rows)} producto(s) bajo stock minimo")

    cfg = load_config()

    # Notificación Windows (siempre, si hay sesión abierta)
    nombres = ", ".join(r["name"] for r in rows[:3])
    sufijo  = f" (+{len(rows) - 3} mas)" if len(rows) > 3 else ""
    show_notification(
        f"Paniol — {len(rows)} producto(s) bajo minimo",
        f"{nombres}{sufijo}",
    )

    # Email (solo si app_password está configurada)
    send_email(cfg, rows)
