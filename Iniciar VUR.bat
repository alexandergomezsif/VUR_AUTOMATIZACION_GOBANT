@echo off
REM ============================================================
REM  Abre el programa "Consulta de valorizacion en VUR".
REM  Doble clic sobre este archivo (o sobre el acceso directo).
REM ============================================================
cd /d "%~dp0"

REM 0. Ubicar el programa: carpeta "programa" (organizada) o esta misma carpeta
set "APP=%~dp0programa\vur_valorizacion.py"
if not exist "%APP%" set "APP=%~dp0vur_valorizacion.py"
if not exist "%APP%" (
  echo.
  echo  No se encontro vur_valorizacion.py en esta carpeta ni en "programa".
  echo.
  pause
  exit /b 1
)

REM 1. Buscar Python (pyw/pythonw abren el programa sin ventana negra)
set "PY="
where pyw >nul 2>nul && set "PY=pyw"
if not defined PY where pythonw >nul 2>nul && set "PY=pythonw"
if not defined PY (
  echo.
  echo  No se encontro Python en este equipo.
  echo  Instalelo desde https://www.python.org/downloads/
  echo  y marque la casilla "Add python.exe to PATH" durante la instalacion.
  echo.
  pause
  exit /b 1
)

REM 2. Version de consola del mismo Python (pyw -> py, pythonw -> python)
set "PYC=%PY:w=%"

REM 3. Verificar librerias; si faltan, instalarlas una sola vez
%PYC% -c "import selenium, pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Faltan librerias del programa. Instalando selenium, pandas y openpyxl...
  echo.
  %PYC% -m pip install --user selenium pandas openpyxl xlrd
  if errorlevel 1 (
    echo.
    echo  No se pudieron instalar las librerias. Revise la conexion a internet.
    pause
    exit /b 1
  )
)

REM 4. Abrir el programa
start "" %PY% "%APP%"
