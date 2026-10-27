@echo off
REM ================================================================
REM  INSTALADOR - Consulta de valorizacion en VUR (version BETA)
REM  Gobernacion de Antioquia - Secretaria de Infraestructura
REM
REM  Un solo archivo: descarguelo y haga doble clic. El instalador
REM  pregunta en que carpeta instalar, descarga el programa desde
REM  GitHub, instala Python y las librerias si faltan, crea las
REM  carpetas de trabajo y el acceso directo del Escritorio.
REM  Ejecutarlo de nuevo ACTUALIZA el programa sin borrar consultas.
REM ================================================================
setlocal
title Instalador VUR - Gobernacion de Antioquia
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText('%~f0'); $m='#==INICIO'+'_POWERSHELL=='; iex $t.Substring($t.IndexOf($m)+$m.Length)"
echo.
pause
exit /b

#==INICIO_POWERSHELL==
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Zip = 'https://github.com/alexandergomezsif/VUR_AUTOMATIZACION_GOBANT/archive/refs/heads/main.zip'
$CfgDir = Join-Path $env:APPDATA 'VUR_Valorizacion'
$Cfg = Join-Path $CfgDir 'config.json'

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
if ($anterior -and (Test-Path (Join-Path $anterior 'programa'))) {
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
    if ((Split-Path $base -Leaf) -match '^vur_valorizacion$' -or (Test-Path (Join-Path $base 'programa'))) {
        $destino = $base
    } else {
        $destino = Join-Path $base 'VUR_valorizacion'
    }
}
Write-Host "       Carpeta: $destino"
New-Item -ItemType Directory -Force -Path $destino | Out-Null

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

# ---- 3. Descargar el programa ---------------------------------------------------
Paso 3 'Descargando el programa desde GitHub'
$tmp = Join-Path $env:TEMP ('vur_instalador_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
try {
    Invoke-WebRequest -UseBasicParsing -Uri $Zip -OutFile (Join-Path $tmp 'vur.zip')
} catch { Falla "No se pudo descargar el programa: $($_.Exception.Message). Revise internet o el proxy de la entidad." }
Expand-Archive -Path (Join-Path $tmp 'vur.zip') -DestinationPath $tmp -Force
$origen = Get-ChildItem $tmp -Directory | Select-Object -First 1

# ---- 4. Copiar (respaldando la version anterior) --------------------------------
Paso 4 'Instalando archivos'
foreach ($sub in @('consultas', 'datos', 'historico', 'respaldos', 'programa')) {
    New-Item -ItemType Directory -Force -Path (Join-Path $destino $sub) | Out-Null
}
$viejo = Join-Path $destino 'programa\vur_valorizacion.py'
if (Test-Path $viejo) {
    $copia = Join-Path $destino ('respaldos\vur_valorizacion_' + (Get-Date -Format 'yyyyMMdd_HHmm') + '.py')
    Copy-Item $viejo $copia -Force
    Write-Host "       Version anterior respaldada en respaldos\"
}
# Solo se copian archivos del programa: consultas, historico y el perfil de Chrome no se tocan.
Get-ChildItem $origen.FullName -Force | Where-Object { $_.Name -notin @('.gitignore', '.gitattributes', '.github') } |
    ForEach-Object { Copy-Item $_.FullName -Destination $destino -Recurse -Force }
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

# ---- 5. Librerias ---------------------------------------------------------------
Paso 5 'Instalando librerias de Python (selenium, pandas, openpyxl)'
& $py -m pip install --user --upgrade --disable-pip-version-check -q -r (Join-Path $destino 'requirements.txt')
if ($LASTEXITCODE -ne 0) { Falla 'No se pudieron instalar las librerias (revise internet o el proxy de la entidad).' }

# ---- 6. Configuracion y acceso directo --------------------------------------------
Paso 6 'Configuracion y acceso directo'
New-Item -ItemType Directory -Force -Path $CfgDir | Out-Null
@{ carpeta = $destino; instalado = (Get-Date -Format 'yyyy-MM-dd HH:mm'); origen = 'Instalar_VUR.bat' } |
    ConvertTo-Json | Set-Content -Path $Cfg -Encoding UTF8
$escritorio = [Environment]::GetFolderPath('Desktop')
$atajo = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $escritorio 'Consulta VUR.lnk'))
$atajo.TargetPath = Join-Path $destino 'Iniciar VUR.bat'
$atajo.WorkingDirectory = $destino
$atajo.IconLocation = Join-Path $destino 'programa\icono_vur.ico'
$atajo.WindowStyle = 7
$atajo.Description = 'Consulta de valorizacion en VUR'
$atajo.Save()
Write-Host '       Acceso directo "Consulta VUR" creado en el Escritorio.'

# ---- 7. Verificacion ------------------------------------------------------------
Paso 7 'Verificando el programa (pruebas automaticas)'
Push-Location (Join-Path $destino 'programa')
$salida = & $py pruebas_reglas.py 2>&1 | Out-String
Pop-Location
if ($salida -match 'TODAS LAS PRUEBAS PASARON') { Write-Host '       Todas las pruebas pasaron.' }
else { Aviso 'Algunas pruebas fallaron. Avise al desarrollador antes de usar el programa.'; Write-Host $salida }
$chrome = @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe") | Where-Object { Test-Path $_ }
if (-not $chrome) { Aviso 'No se encontro Google Chrome. El programa lo necesita para entrar a VUR.' }

Write-Host ''
Write-Host ' ==========================================================' -ForegroundColor DarkGreen
Write-Host "   LISTO. Programa instalado en: $destino" -ForegroundColor White
Write-Host '   Abralo con el acceso directo "Consulta VUR" del Escritorio.' -ForegroundColor White
Write-Host ' ==========================================================' -ForegroundColor DarkGreen
$r = Read-Host '   Abrir el programa ahora? (S/N)'
if ($r -match '^[sS]') { Start-Process -FilePath (Join-Path $destino 'Iniciar VUR.bat') -WorkingDirectory $destino -WindowStyle Minimized }
