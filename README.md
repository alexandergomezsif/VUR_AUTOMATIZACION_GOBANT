# Consulta de valorización en VUR

Aplicativo de escritorio para revisar si una matrícula inmobiliaria tiene, o tuvo, gravamen por
valorización. Revisa en la Ventanilla Única de Registro (VUR) el folio y sus matrículas matriz.

**Gobernación de Antioquia** · Secretaría de Infraestructura · Dirección Administrativa y Financiera
Desarrollado por Alexander Gómez A. · Versión BETA

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
