@echo off
rem Mantenimiento semanal de BD SQLite — Inventario Paniol
rem Invocado por el Programador de tareas de Windows

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%..\\.venv\\Scripts\\python.exe
if not exist "%PYTHON%" set PYTHON=python

"%PYTHON%" "%SCRIPT_DIR%vacuum_db.py"
exit /b %errorlevel%
