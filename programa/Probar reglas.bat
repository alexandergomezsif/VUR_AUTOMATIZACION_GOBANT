@echo off
REM Ejecuta las pruebas automaticas de las reglas del programa.
REM Usar despues de cualquier cambio en vur_valorizacion.py.
cd /d "%~dp0"
where py >nul 2>nul && (py pruebas_reglas.py) || (python pruebas_reglas.py)
echo.
pause
