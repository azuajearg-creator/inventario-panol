@echo off
rem Alerta de stock mínimo — Inventario Pañol
rem Invocado por el Programador de tareas de Windows

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%..\\.venv\\Scripts\\python.exe
if not exist "%PYTHON%" set PYTHON=python

"%PYTHON%" "%SCRIPT_DIR%alerta_stock.py"
exit /b %errorlevel%
