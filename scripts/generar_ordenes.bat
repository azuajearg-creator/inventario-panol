@echo off
rem Generacion automatica de OC borrador — Inventario Paniol
rem Invocado por el Programador de tareas de Windows

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%..\\.venv\\Scripts\\python.exe
if not exist "%PYTHON%" set PYTHON=python

"%PYTHON%" "%SCRIPT_DIR%generar_ordenes.py"
exit /b %errorlevel%
