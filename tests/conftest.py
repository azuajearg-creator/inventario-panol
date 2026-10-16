"""Configuración común de las pruebas.

Se ejecuta antes de importar `app`:
- PANOL_SKIP_INIT_DB evita crear la base real al importar el módulo.
- PANOL_ADMIN_PASSWORD define una contraseña conocida para poder iniciar sesión.
  La aplicación ya no trae una contraseña fija: si no se define, genera una aleatoria.
"""

import os

os.environ.setdefault("PANOL_SKIP_INIT_DB", "1")
os.environ.setdefault("PANOL_ADMIN_PASSWORD", "admin1234")
