@echo off
rem Recordatorio de devoluciones pendientes — Inventario Paniol
rem Invocado por el Programador de tareas de Windows

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%..\\.venv\\Scripts\\python.exe
if not exist "%PYTHON%" set PYTHON=python

"%PYTHON%" "%SCRIPT_DIR%recordatorio_devoluciones.py"
exit /b %errorlevel%
