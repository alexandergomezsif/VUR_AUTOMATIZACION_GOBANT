@echo off
REM ================================================================
REM  INSTALADOR - Consulta de valorizacion en VUR (version BETA)
REM  Gobernacion de Antioquia - Secretaria de Infraestructura
REM
REM  1. Descargue el ZIP del repositorio y extraigalo (clic derecho >
REM     "Extraer todo").
REM  2. Haga doble clic en este archivo.
REM  El instalador pregunta la carpeta, instala Python y las librerias
REM  si faltan, copia el programa, crea las carpetas de trabajo y el
REM  acceso directo. Ejecutarlo de nuevo ACTUALIZA sin borrar datos.
REM ================================================================
setlocal
title Instalador VUR - Gobernacion de Antioquia
set "VUR_ORIGEN=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText('%~f0'); $m='#==INICIO'+'_POWERSHELL=='; iex $t.Substring($t.IndexOf($m)+$m.Length)"
echo.
pause
exit /b

#==INICIO_POWERSHELL==
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$UrlPrograma = 'https://github.com/alexandergomezsif/VUR_AUTOMATIZACION_GOBANT/raw/main/vur_valorizacion.py'
$Librerias = @('selenium>=4.11', 'pandas>=2.0', 'openpyxl>=3.1', 'xlrd>=2.0')
$CfgDir = Join-Path $env:APPDATA 'VUR_Valorizacion'
$Cfg = Join-Path $CfgDir 'config.json'
$Usuario = Join-Path $env:LOCALAPPDATA 'VUR_Valorizacion'
$Fecha = Get-Date -Format 'yyyyMMdd_HHmm'

function Paso($n, $texto) { Write-Host ''; Write-Host " [$n/7] $texto" -ForegroundColor Green }
function Aviso($texto) { Write-Host "       $texto" -ForegroundColor Yellow }
function Falla($texto) { Write-Host ''; Write-Host " ERROR: $texto" -ForegroundColor Red; exit 1 }

Write-Host ''
Write-Host ' ==========================================================' -ForegroundColor DarkGreen
Write-Host '   Consulta de valorizacion en VUR  -  Instalador BETA' -ForegroundColor White
Write-Host '   Gobernacion de Antioquia - Secretaria de Infraestructura' -ForegroundColor Gray
Write-Host ' ==========================================================' -ForegroundColor DarkGreen
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---- 1. Carpeta ---------------------------------------------------------------
Paso 1 'Carpeta del aplicativo'
$anterior = $null
if (Test-Path $Cfg) {
    try { $anterior = (Get-Content $Cfg -Raw | ConvertFrom-Json).carpeta } catch { $anterior = $null }
}
$destino = $null
if ($anterior -and (Test-Path $anterior)) {
    Write-Host "       Ya esta instalado en: $anterior"
    $r = Read-Host '       Actualizar esa instalacion? (S = si, N = elegir otra carpeta)'
    if ($r -notmatch '^[nN]') { $destino = $anterior }
}
if (-not $destino) {
    Write-Host '       Se abrira una ventana para elegir DONDE crear la carpeta VUR_valorizacion.'
    Add-Type -AssemblyName System.Windows.Forms
    $dueno = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Elija donde crear la carpeta VUR_valorizacion (por ejemplo Documentos o C:\)'
    $dlg.SelectedPath = [Environment]::GetFolderPath('MyDocuments')
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog($dueno) -ne [System.Windows.Forms.DialogResult]::OK) { Falla 'Instalacion cancelada: no se eligio carpeta.' }
    $base = $dlg.SelectedPath
    if ((Split-Path $base -Leaf) -match '^vur_valorizacion$' -or (Test-Path (Join-Path $base 'vur_valorizacion.py')) -or
        (Test-Path (Join-Path $base 'programa\vur_valorizacion.py'))) {
        $destino = $base
    } else {
        $destino = Join-Path $base 'VUR_valorizacion'
    }
}
Write-Host "       Carpeta: $destino"
foreach ($sub in @('', 'consultas', 'datos', 'historico', 'respaldos')) {
    New-Item -ItemType Directory -Force -Path (Join-Path $destino $sub) | Out-Null
}

# ---- 2. Python ----------------------------------------------------------------
Paso 2 'Python'
function Buscar-Python {
    foreach ($c in @('py', 'python')) {
        try {
            $v = & $c -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
            if ($LASTEXITCODE -eq 0 -and $v -match '^3\.(\d+)$' -and [int]$Matches[1] -ge 10) { return $c }
        } catch {}
    }
    $lanzador = Join-Path $env:LOCALAPPDATA 'Programs\Python\Launcher\py.exe'
    if (Test-Path $lanzador) { return $lanzador }
    return $null
}
$py = Buscar-Python
if (-not $py) {
    Aviso 'No se encontro Python 3.10 o superior.'
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        $r = Read-Host '       Instalar Python 3.12 ahora (solo para su usuario, sin permisos de administrador)? (S/N)'
        if ($r -match '^[sS]') {
            winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
            $env:Path = [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + [Environment]::GetEnvironmentVariable('Path', 'Machine')
            $py = Buscar-Python
        }
    }
    if (-not $py) { Falla 'Instale Python desde https://www.python.org/downloads/ (marque "Add python.exe to PATH") y ejecute de nuevo este instalador.' }
}
Write-Host "       Python listo: $(& $py --version)"

# ---- 3. Programa ------------------------------------------------------------------
Paso 3 'Programa'
$nuevo = Join-Path $env:VUR_ORIGEN 'vur_valorizacion.py'
if (Test-Path $nuevo) {
    Write-Host '       Se usa vur_valorizacion.py del ZIP descargado.'
} else {
    Aviso 'No se encontro vur_valorizacion.py junto al instalador (el ZIP no se extrajo); se descarga de GitHub.'
    $nuevo = Join-Path $env:TEMP 'vur_valorizacion_descargado.py'
    try { Invoke-WebRequest -UseBasicParsing -Uri $UrlPrograma -OutFile $nuevo }
    catch { Falla "No se pudo descargar el programa: $($_.Exception.Message). Extraiga el ZIP completo y ejecute el instalador desde esa carpeta." }
}

# ---- 4. Instalar (respaldando lo anterior) ------------------------------------------
Paso 4 'Instalando archivos'
$respaldo = Join-Path $destino "respaldos\version_anterior_$Fecha"
$actual = Join-Path $destino 'vur_valorizacion.py'
if ((Test-Path $actual) -and ((Resolve-Path $actual).Path -ne (Resolve-Path $nuevo).Path)) {
    New-Item -ItemType Directory -Force -Path $respaldo | Out-Null
    Copy-Item $actual $respaldo -Force
}
# Versiones anteriores (carpeta "programa" y varios .bat): el perfil de Chrome pasa a la carpeta
# del usuario y lo demas se mueve a respaldos. No se borra nada.
$viejo = Join-Path $destino 'programa'
if (Test-Path $viejo) {
    $perfilViejo = Join-Path $viejo 'selenium_profile'
    $perfilNuevo = Join-Path $Usuario 'perfil_chrome'
    if ((Test-Path $perfilViejo) -and -not (Test-Path $perfilNuevo)) {
        New-Item -ItemType Directory -Force -Path $Usuario | Out-Null
        try { Move-Item $perfilViejo $perfilNuevo } catch { Aviso 'No se pudo mover el perfil de Chrome (cierre Chrome y el programa VUR y repita).' }
    }
    New-Item -ItemType Directory -Force -Path $respaldo | Out-Null
    Move-Item $viejo (Join-Path $respaldo 'programa') -Force
}
foreach ($f in @('Iniciar VUR.bat', 'Crear acceso directo.bat', 'Guardar clave de VUR.bat', 'Organizar carpeta.bat',
                 'Descargar manual de identidad.bat', 'LEAME.txt', 'requirements.txt')) {
    $p = Join-Path $destino $f
    if (Test-Path $p) { New-Item -ItemType Directory -Force -Path $respaldo | Out-Null; Move-Item $p $respaldo -Force }
}
if (Test-Path $respaldo) { Write-Host "       Version anterior guardada en respaldos\version_anterior_$Fecha" }
if ((Resolve-Path $nuevo).Path -ne (Join-Path (Resolve-Path $destino).Path 'vur_valorizacion.py')) { Copy-Item $nuevo $actual -Force }
Write-Host '       vur_valorizacion.py instalado.'

# ---- 5. Librerias ---------------------------------------------------------------
Paso 5 'Instalando librerias de Python (selenium, pandas, openpyxl, xlrd)'
& $py -m pip install --user --upgrade --disable-pip-version-check -q @Librerias
if ($LASTEXITCODE -ne 0) { Falla 'No se pudieron instalar las librerias (revise internet o el proxy de la entidad).' }

# ---- 6. Configuracion y acceso directo --------------------------------------------
Paso 6 'Configuracion y acceso directo'
New-Item -ItemType Directory -Force -Path $CfgDir | Out-Null
@{ carpeta = $destino; instalado = (Get-Date -Format 'yyyy-MM-dd HH:mm'); origen = 'Instalar_VUR.bat' } |
    ConvertTo-Json | Set-Content -Path $Cfg -Encoding UTF8
& $py $actual --recursos | Out-Null
$pyw = (& $py -c "import sys, os; print(os.path.join(os.path.dirname(sys.executable), 'pythonw.exe'))").Trim()
if (-not (Test-Path $pyw)) { $pyw = (& $py -c "import sys; print(sys.executable)").Trim() }
$escritorio = [Environment]::GetFolderPath('Desktop')
$atajo = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $escritorio 'Consulta VUR.lnk'))
$atajo.TargetPath = $pyw
$atajo.Arguments = '"' + $actual + '"'
$atajo.WorkingDirectory = $destino
$icono = Join-Path $Usuario 'recursos\icono_vur.ico'
if (Test-Path $icono) { $atajo.IconLocation = $icono }
$atajo.Description = 'Consulta de valorizacion en VUR'
$atajo.Save()
Write-Host '       Acceso directo "Consulta VUR" creado en el Escritorio.'

# ---- 7. Verificacion ------------------------------------------------------------
Paso 7 'Verificando el programa (pruebas automaticas)'
# En Windows PowerShell 5.1, leer la salida de error de un programa con 2>&1 se toma como error si
# ErrorActionPreference es 'Stop'; por eso se relaja solo para este paso.
$ErrorActionPreference = 'Continue'
$salida = & $py $actual --probar 2>&1 | Out-String
$ErrorActionPreference = 'Stop'
if ($salida -match 'TODAS LAS PRUEBAS PASARON') { Write-Host '       Todas las pruebas pasaron.' }
else { Aviso 'Algunas pruebas fallaron. Avise al desarrollador antes de usar el programa.'; Write-Host $salida }
$chrome = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe") | Where-Object { Test-Path $_ }
if (-not $chrome) { Aviso 'No se encontro Google Chrome. El programa lo necesita para entrar a VUR.' }

Write-Host ''
Write-Host ' ==========================================================' -ForegroundColor DarkGreen
Write-Host "   LISTO. Programa instalado en: $destino" -ForegroundColor White
Write-Host '   Abralo con el acceso directo "Consulta VUR" del Escritorio.' -ForegroundColor White
Write-Host '   Ya puede borrar el ZIP y la carpeta descargada.' -ForegroundColor White
Write-Host ' ==========================================================' -ForegroundColor DarkGreen
$r = Read-Host '   Abrir el programa ahora? (S/N)'
if ($r -match '^[sS]') { Start-Process -FilePath $pyw -ArgumentList ('"' + $actual + '"') -WorkingDirectory $destino }
