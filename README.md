# Consulta de valorización en VUR

Aplicativo de escritorio para revisar si una matrícula inmobiliaria tiene, o tuvo, gravamen por
valorización. Revisa en la Ventanilla Única de Registro (VUR) el folio y sus matrículas matriz.

**Gobernación de Antioquia** · Secretaría de Infraestructura · Dirección Administrativa y Financiera
Desarrollado por Alexander Gómez A. · Versión BETA

## Instalación

1. Descargue el ZIP: [VUR_AUTOMATIZACION_GOBANT-main.zip](https://github.com/alexandergomezsif/VUR_AUTOMATIZACION_GOBANT/archive/refs/heads/main.zip).
   También puede usar el botón verde **Code → Download ZIP**.
2. Clic derecho sobre el ZIP → **Extraer todo**.
3. En la carpeta extraída, doble clic en **`Instalar_VUR.bat`**. Si Windows muestra «Windows protegió su PC»,
   pulse **Más información → Ejecutar de todas formas**.
4. Siga las instrucciones en pantalla. El instalador:
   1. pregunta en qué carpeta instalar;
   2. instala Python, si falta, y las librerías;
   3. copia el programa;
   4. crea las carpetas de trabajo y el acceso directo **«Consulta VUR»** en el Escritorio;
   5. ejecuta las pruebas automáticas.
5. Al terminar puede borrar el ZIP y la carpeta extraída.

**Actualizar:** descargue el ZIP nuevo y ejecute otra vez `Instalar_VUR.bat`. No borra consultas
ni histórico, y guarda la versión anterior en `respaldos`.

**Requisitos:** Windows 10 u 11, Google Chrome, internet, permiso para ejecutar PowerShell y un
usuario propio de VUR.

## Uso

1. Paso 1: **Seleccionar Excel** con la columna de matrículas (por ejemplo `001-123456`).
2. Paso 2: **Iniciar consultas**. Se abre Chrome; usted inicia sesión en VUR y resuelve el captcha.
3. Paso 3: **Guardar Excel**. El archivo trae las hojas «Resultados» y «Resumen». Con **Comparar** se
   ve qué cambió entre dos consultas del mismo lote.

Antes de consultar VUR, el programa revisa el **histórico**. Si una matrícula se consultó hace menos
de 30 días con confiabilidad ALTA, toma ese resultado (columna *Fuente* = Histórico) y no gasta
cupo de VUR. Para consultarla de nuevo, marque **«Forzar consulta en VUR»**.

## Carpetas en el equipo

| Carpeta | Contenido |
|---|---|
| `VUR_valorizacion\vur_valorizacion.py` | El aplicativo completo |
| `consultas\` | Excel de matrículas, resultados y comparaciones |
| `historico\` | Consultas anteriores. `registros\` no se edita; `Historico_consultas.xlsx` es para leer |
| `datos\`, `respaldos\` | Excel de proyectos (solo el desarrollador) y versiones anteriores |
| `%LOCALAPPDATA%\VUR_Valorizacion\` | Perfil de Chrome del programa y recursos (logo, icono, fuentes) |

## Repositorio

| Archivo | Para qué |
|---|---|
| `Instalar_VUR.bat` | Instalador: lo único que usan los profesionales |
| `vur_valorizacion.py` | Todo el aplicativo: código, datos de proyectos por municipio, logo, icono, tipografía Poppins (SIL OFL 1.1) y pruebas |
| `desarrollo/generar_datos_valorizacion.py` | Actualiza los datos de proyectos dentro del aplicativo cuando cambia el Excel |

**Datos personales:** el archivo `.gitignore` impide subir Excel de consultas, el histórico o el perfil
de Chrome, porque contienen datos protegidos por la Ley 1581 de 2012.

### Para el desarrollador

```
python vur_valorizacion.py --probar        # pruebas automáticas: todas deben pasar
python desarrollo/generar_datos_valorizacion.py "ruta\INFORMACION PROYECTOS DE VALORIZACION (1).xlsx"
```
