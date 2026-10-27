@echo off
REM ============================================================
REM  Crea (o actualiza) en el Escritorio el acceso directo
REM  "Consulta VUR". Basta con ejecutarlo una vez (doble clic).
REM ============================================================
cd /d "%~dp0"
set "ICONO=%~dp0programa\icono_vur.ico"
if not exist "%ICONO%" set "ICONO=%~dp0icono_vur.ico"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('Desktop'); $s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d 'Consulta VUR.lnk')); $s.TargetPath=(Join-Path '%~dp0' 'Iniciar VUR.bat'); $s.WorkingDirectory='%~dp0'; $s.IconLocation='%ICONO%'; $s.WindowStyle=7; $s.Description='Consulta de valorizacion en VUR'; $s.Save()"
if errorlevel 1 (
  echo.
  echo  No se pudo crear el acceso directo.
) else (
  echo.
  echo  Listo: en su Escritorio quedo el acceso directo "Consulta VUR".
)
if /i not "%~1"=="/sinpausa" (
  echo.
  pause
)
