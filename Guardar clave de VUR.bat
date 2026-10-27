@echo off
REM ============================================================
REM  Abre Chrome NORMAL (sin control del programa) con el mismo
REM  perfil que usa el aplicativo, en la pagina de ingreso de VUR.
REM  Ahi Chrome si ofrece "Guardar contrasena".
REM  Usarlo una vez, con el aplicativo CERRADO.
REM ============================================================
cd /d "%~dp0"
set "PERFIL=%~dp0programa\selenium_profile"
if not exist "%PERFIL%" set "PERFIL=%~dp0selenium_profile"
echo.
echo  1. Asegurese de que el programa de consulta VUR este CERRADO.
echo  2. Se abrira Chrome en la pagina de ingreso de VUR.
echo  3. Escriba usuario y contrasena, resuelva el captcha e ingrese.
echo  4. Cuando Chrome pregunte "Guardar contrasena", pulse GUARDAR.
echo  5. Cierre esa ventana de Chrome.
echo.
pause
start "" chrome --user-data-dir="%PERFIL%" "https://www.vur.gov.co/portal/pages/vur/inicio.jsf"
if errorlevel 1 (
  start "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe" --user-data-dir="%PERFIL%" "https://www.vur.gov.co/portal/pages/vur/inicio.jsf"
)
