# Consulta de valorización en VUR

Aplicativo de escritorio para revisar si una matrícula inmobiliaria tiene, o tuvo, gravamen por
valorización. Revisa en la Ventanilla Única de Registro (VUR) el folio y sus matrículas matriz.

**Gobernación de Antioquia** · Secretaría de Infraestructura · Dirección Administrativa y Financiera
Desarrollado por Alexander Gómez A. · Versión BETA

## Instalación en un equipo nuevo (pruebas BETA)

1. Descargue el instalador: [Instalar_VUR.bat](https://github.com/alexandergomezsif/VUR_AUTOMATIZACION_GOBANT/raw/main/Instalar_VUR.bat).
   Si el navegador lo abre como texto, use clic derecho → «Guardar enlace como…».
2. Haga doble clic en `Instalar_VUR.bat`. Si Windows muestra «Windows protegió su PC», pulse
   **Más información** → **Ejecutar de todas formas**; el aviso sale porque el archivo viene de internet.
3. El instalador:
   1. pregunta en qué carpeta instalar;
   2. descarga el programa;
   3. instala Python (si falta) y las librerías;
   4. crea las carpetas `consultas`, `datos`, `historico` y `respaldos`;
   5. crea el acceso directo «Consulta VUR» en el Escritorio;
   6. ejecuta las pruebas automáticas.
4. Para **actualizar**, ejecute de nuevo el instalador y responda «S». No borra consultas ni histórico,
   y guarda la versión anterior en `respaldos`.

Requiere Windows 10 u 11, Google Chrome, conexión a internet y permiso para ejecutar PowerShell.

## Cómo se usa

1. Doble clic en `Iniciar VUR.bat`. La primera vez instala las librerías que faltan.
2. Paso 1: seleccionar el Excel con la columna de matrículas.
3. Paso 2: «Iniciar consultas». El funcionario inicia sesión en VUR y resuelve el captcha.
4. Paso 3: «Guardar Excel». El archivo trae las hojas «Resultados» y «Resumen».

Detalle de la carpeta y del histórico de consultas: ver `LEAME.txt`.

## Requisitos

- Windows 10 u 11, Google Chrome y Python 3.10 o superior.
- Librerías: `selenium`, `pandas`, `openpyxl` (y `xlrd` para archivos .xls).

## Estructura

| Ruta | Contenido |
|---|---|
| `programa/vur_valorizacion.py` | Aplicativo: interfaz, consulta a VUR, reglas de análisis, histórico |
| `programa/datos_valorizacion.py` | Proyectos por municipio (generado; no editar a mano) |
| `programa/generar_datos_valorizacion.py` | Regenera el archivo anterior desde el Excel de proyectos |
| `programa/pruebas_reglas.py` | Pruebas automáticas (`Probar reglas.bat`) |
| `programa/fuentes/` | Tipografía Poppins (licencia SIL OFL 1.1) |
| `*.bat` | Abrir el programa, crear el acceso directo, guardar la clave en Chrome |

## Lo que NO se sube a este repositorio

El archivo `.gitignore` excluye todo lo que contiene datos personales o de sesión:

- los Excel de consultas y resultados;
- el histórico;
- el Excel de proyectos;
- el perfil de Chrome (`selenium_profile`, que guarda la sesión y la clave de VUR).

Estos datos están protegidos por la Ley 1581 de 2012 y deben quedarse en el equipo o en la carpeta
autorizada.

## Después de cualquier cambio al código

Ejecutar `programa/Probar reglas.bat`. Todas las pruebas deben pasar antes de usar el programa.
