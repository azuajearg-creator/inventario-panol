@echo off
rem Copia de seguridad automática — Inventario Pañol
rem Invocado por el Programador de tareas de Windows

set SCRIPT_DIR=%~dp0
set PYTHON=%SCRIPT_DIR%..\\.venv\\Scripts\\python.exe

rem Fallback al Python del sistema si no existe el venv
if not exist "%PYTHON%" set PYTHON=python

"%PYTHON%" "%SCRIPT_DIR%hacer_backup.py"
exit /b %errorlevel%
