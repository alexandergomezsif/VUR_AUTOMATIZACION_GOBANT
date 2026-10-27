"""
Automatización VUR - Consulta de gravámenes por valorización.

El aplicativo usa Chrome (Selenium) solo para que el usuario inicie sesión en
VUR. Las consultas se hacen a los mismos servicios JSON que usa la página de
VUR, dentro de esa sesión:
    webresources/consultaIndices/datosBasicos/tierras/<círculo>-<número>
    webresources/consultaIndices/estadoJuridico/tierras/<círculo>-<número>
Así no se depende de leer tablas en pantalla.

Flujo por matrícula:
  1. Datos básicos: propietarios (todos), dirección, referencia catastral,
     municipio, Nupre, fecha de apertura del folio y matrículas matriz.
  2. Clasificación del municipio con la hoja "Proy x Mpio" del Excel de
     proyectos (incorporada en datos_valorizacion.py):
       000 / NO APLICA  -> no se consulta estado jurídico.
       001 / CONDONADA  -> se consulta estado jurídico.
       Sin clasificar   -> se consulta estado jurídico por precaución.
  3. Estado jurídico: se analizan las anotaciones. Es de valorización la que
     tiene alguna PALABRAS_VALORIZACION en su especificación. Una anotación
     posterior con "Se cancela anotación No: N" la marca como cancelada.
  4. Si no hay anotación de valorización y el folio se abrió DESPUÉS de la
     fecha límite del municipio (proyecto más antiguo), se revisan sus
     matrículas matriz como consultas nuevas, nivel por nivel y sin tope de
     niveles, hasta encontrar la valorización (y a quién se impuso) o llegar a
     folios sin matriz o anteriores a la fecha límite.

Archivos necesarios en la misma carpeta:
  vur_valorizacion.py, datos_valorizacion.py
  (generar_datos_valorizacion.py solo para regenerar los datos cuando cambie el Excel)
  logo_gobernacion.png  (logo oficial; si falta, se muestra el nombre en texto)
  fuentes/Poppins-Regular.ttf, fuentes/Poppins-Bold.ttf, fuentes/OFL-Poppins.txt
                        (tipografía del manual; si faltan, se usa Segoe UI)

Requisitos:
  pip install --user selenium pandas openpyxl xlrd
  (Selenium >= 4.6 gestiona ChromeDriver automáticamente.)
"""

import html
import json
import math
import os
import queue
import re
import threading
import tkinter as tk
import unicodedata
from collections import deque
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    InvalidSessionIdException,
    JavascriptException,
    NoAlertPresentException,
    NoSuchWindowException,
    TimeoutException,
    WebDriverException,
)

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

URL_INICIO = "https://www.vur.gov.co/portal/pages/vur/inicio.jsf"
URL_API = "https://www.vur.gov.co/portal/PantallasVUR/webresources/consultaIndices/"
TIPO_CONSULTA = "tierras"  # el mismo que usan los módulos "...Tierras" del menú de VUR

PERFIL_CHROME = Path(__file__).resolve().parent / "selenium_profile"

MAX_INTENTOS = 3
GUARDADO_PARCIAL_CADA = 10
TIMEOUT_API = 120  # segundos por consulta a VUR

# Palabras que identifican una anotación de valorización en su ESPECIFICACION.
# "GRAVAMEN" no se usa sola: también aparece en hipotecas ("NATURALEZA JURIDICA
# DEL ACTO: GRAVAMEN"), lo que producía falsos positivos.
PALABRAS_VALORIZACION = ["VALORIZA", "INVIAS"]

# Revisión de matrículas matriz.
# No hay tope de niveles: se sube hasta que el folio no tenga matriz o sea
# anterior a la fecha límite. MAX_FOLIOS_SEGURIDAD solo protege contra un
# recorrido descontrolado (cada folio son 2 consultas del cupo diario de VUR).
MAX_FOLIOS_SEGURIDAD = 300
# Si la revisión no se puede completar (tope de seguridad o una matriz sin
# respuesta), el estado queda "NO GRAVADA - REVISION INCOMPLETA".
ESTADO_INCOMPLETO = "NO GRAVADA - REVISION INCOMPLETA"
# Fecha límite del municipio: la más antigua de sus proyectos, tomando para cada
# proyecto la Fecha de Inicio (hoja Vencimiento) o, si no la tiene, la fecha de
# la resolución (hoja Proyectos Valorización) o el 1 de enero de su año.
# Si el folio se abrió después de esa fecha, se revisa su matriz.

# Orden de columnas del Excel de resultados.
COLUMNAS_SALIDA = [
    "Matricula_Original", "Matricula_Limpia",
    "Propietario", "Total_Propietarios", "Tipo_Identificacion", "Numero_Identificacion",
    "Propietarios_Actuales",
    "Direccion_Inmueble", "Matricula_VUR", "Referencia_Catastral", "Departamento", "Municipio",
    "Nupre", "Tipo_Predio", "Estado_Folio", "Fecha_Apertura_Folio", "Matriculas_Matriz",
    "Clasificacion_Municipio", "Proyectos_Municipio", "Fecha_Limite_Revision",
    "Estado", "Confiabilidad", "Motivos_Revision", "Segunda_Pasada",
    "Matricula_Con_Anotacion", "Nivel_Matriz", "Detalle_Valorizacion", "Propietario_Gravado",
    "Matricula_Cancelacion", "Observaciones", "Afectaciones_Valorizacion",
    "Cadena_Revisada", "Folios_Revisados", "Notas_Proceso", "Intentos", "Fecha_Consulta", "Version_Aplicativo",
    "Fuente", "Funcionario", "Equipo",
]

# Versión de las reglas de análisis. Se escribe en cada fila del Excel para saber
# con qué reglas se obtuvo cada resultado.
VERSION_APLICATIVO = "2026.10.26-5"
# Versión de las REGLAS de análisis. Se cambia solo cuando cambia cómo se clasifica una
# matrícula; así los resultados guardados en el histórico con reglas anteriores no se
# reutilizan y se vuelven a consultar. Los cambios de pantalla no la modifican.
VERSION_REGLAS = "2026.10.26-4"

# Confiabilidad de cada resultado:
#   ALTA    -> todo lo que sustenta el estado viene explícito en VUR.
#   REVISAR -> algún dato se dedujo, faltó información o hay una alerta; los
#              motivos quedan en Motivos_Revision. Debe verlo un profesional.
# Segunda pasada: al terminar el lote, las matrículas cuyo resultado pudo deberse a
# una falla momentánea de VUR se consultan de nuevo tras una pausa. No se repiten
# las marcadas REVISAR por el contenido (propietario deducido, afectación...),
# porque consultar otra vez daría lo mismo.
ESPERA_SEGUNDA_PASADA = 60  # segundos de pausa antes de la segunda pasada


def necesita_segunda_pasada(reg):
    """True si el resultado pudo deberse a una respuesta fallida o vacía de VUR."""
    if reg.get("Estado") in ("SIN DATOS EN VUR", "ERROR"):
        return "Formato de matrícula inválido" not in (reg.get("Notas_Proceso") or "")
    if reg.get("Estado") == ESTADO_INCOMPLETO:
        notas = reg.get("Notas_Proceso") or ""
        return "no devolvió" in notas or "no se pudo revisar" in notas
    return "no devolvió el estado jurídico" in (reg.get("Notas_Proceso") or "")


CONFIABILIDAD_ALTA = "ALTA"
CONFIABILIDAD_REVISAR = "REVISAR"

# ---------------------------------------------------------------------------
# Datos de valorización por municipio (hoja "Proy x Mpio")
# ---------------------------------------------------------------------------

try:
    from datos_valorizacion import FUENTE_DATOS, MUNICIPIOS_VALORIZACION
except ImportError as exc:  # pragma: no cover - mensaje para el usuario
    raise SystemExit(
        "Falta datos_valorizacion.py junto a vur_valorizacion.py. Genérelo con: "
        "python generar_datos_valorizacion.py (lee el Excel de proyectos de la carpeta 'datos')"
    ) from exc

# Nombre como lo escribe VUR -> nombre en la hoja "Proy x Mpio".
ALIAS_MUNICIPIOS = {
    "SAN ANDRES DE CUERQUIA": "SAN ANDRES DE C.",
    "SAN JOSE DE LA MONTAÑA": "SAN JOSE DE LA M.",
    "SAN PEDRO DE LOS MILAGROS": "SAN PEDRO DE LOS M.",
    "SANTA FE DE ANTIOQUIA": "SANTA FE DE ANT",
    "SANTA ROSA DE OSOS": "SANTA ROSA",
    "DONMATIAS": "DON MATIAS",
    "EL CARMEN DE VIBORAL": "CARMEN DE VIBORAL",
    "EL SANTUARIO": "SANTUARIO",
    "EL RETIRO": "RETIRO",
    "EL PEÑOL": "PEÑOL",
    "CAROLINA DEL PRINCIPE": "CAROLINA",
    "SAN VICENTE FERRER": "SAN VICENTE",
}

ETIQUETAS_ESTADO_MUNICIPIO = {
    "000": "000 - NO GRAVADO",
    "001": "001 - BUSCAR SI TIENE GRAVAMEN",
    "NO_APLICA": "NO APLICA",
    "CONDONADA": "CONDONADA",
}

# Estados en los que NO se consulta el estado jurídico. MEDELLÍN figura como
# "NO APLICA" en la hoja; ARBOLETES (CONDONADA) sí se consulta para dejar
# constancia de la anotación y de su cancelación.
ESTADOS_SIN_CONSULTA = {"000", "NO_APLICA"}

# ---------------------------------------------------------------------------
# Utilidades puras
# ---------------------------------------------------------------------------


def normalizar(texto):
    """Mayúsculas, sin tildes ni diéresis, espacios colapsados."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFD", str(texto))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t).strip().upper()


def clave_municipio(texto):
    """Clave de comparación: normalizado, sin código numérico inicial ni punto final."""
    # VUR envía la Ñ dañada ("PE\ufffdOL"); es la única letra no ASCII de los municipios
    # que llega así, por lo que se reemplaza por N (igual que la Ñ normalizada).
    t = re.sub(r"^[\d\s\-]+", "", normalizar(str(texto or "").replace("\ufffd", "N")))
    return t.rstrip(".").strip()


def _construir_clasificacion():
    clasificacion = {clave_municipio(n): info for n, info in MUNICIPIOS_VALORIZACION.items()}
    for alias, nombre in ALIAS_MUNICIPIOS.items():
        destino = clave_municipio(nombre)
        if destino not in clasificacion:
            raise ValueError(f"Alias apunta a municipio inexistente: {nombre}")
        clasificacion[clave_municipio(alias)] = clasificacion[destino]
    return clasificacion


CLASIFICACION_MUNICIPIOS = _construir_clasificacion()


def clasificar_municipio(municipio):
    """Devuelve (estado, etiqueta, info) o (None, '', None) si no está en la hoja."""
    info = CLASIFICACION_MUNICIPIOS.get(clave_municipio(municipio))
    if not info:
        return None, "", None
    return info["estado"], ETIQUETAS_ESTADO_MUNICIPIO[info["estado"]], info


def describir_proyectos(info):
    """'008 URRAO - BETULIA (res. 1187, 1998, vence 2007-05-01)' separados por ' | '."""
    if not info:
        return ""
    partes = []
    for p in info.get("proyectos", []):
        detalle = [x for x in (
            f"res. {p['resolucion']}" if p.get("resolucion") else "",
            str(p["anio"]) if p.get("anio") else "",
            f"vence {p['fecha_vencimiento']}" if p.get("fecha_vencimiento") else "",
        ) if x]
        texto = f"{p.get('codigo') or ''} {p.get('obra', '')}".strip()
        partes.append(texto + (f" ({', '.join(detalle)})" if detalle else ""))
    return " | ".join(partes)


def parsear_fecha(texto):
    """Acepta dd/mm/aaaa, dd-mm-aaaa y aaaa-mm-dd. Devuelve date o None."""
    t = (texto or "").strip()
    for patron, orden in ((r"(\d{4})-(\d{1,2})-(\d{1,2})", "amd"), (r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", "dma")):
        m = re.search(patron, t)
        if m:
            a, b, c = (int(x) for x in m.groups())
            try:
                return date(c, b, a) if orden == "dma" else date(a, b, c)
            except ValueError:
                return None
    return None


def fecha_limite_municipio(info):
    """Fecha más antigua entre los proyectos del municipio (ver configuración)."""
    fechas = []
    for p in (info or {}).get("proyectos", []):
        f = parsear_fecha(p.get("fecha_inicio")) or parsear_fecha(p.get("fecha_resolucion"))
        if not f and p.get("anio"):
            f = date(int(p["anio"]), 1, 1)
        if f:
            fechas.append(f)
    return min(fechas) if fechas else None


def limpiar_matricula(valor):
    t = re.sub(r"^MI\s*", "", normalizar(valor))
    return t.replace(" ", "")


def separar_matricula(matricula):
    """'01N-0005202' -> ('01N', '5202'). El servicio de VUR no acepta ceros a la izquierda."""
    m = re.fullmatch(r"([0-9A-Z]{3})-(\d+)", matricula or "")
    return (m.group(1), str(int(m.group(2)))) if m else None


def html_a_texto(valor):
    t = re.sub(r"<br\s*/?>", "\n", valor or "", flags=re.IGNORECASE)
    t = re.sub(r"<[^>]+>", " ", t)
    t = html.unescape(t).replace("\r", "\n")
    return "\n".join(re.sub(r"[ \t]+", " ", l).strip() for l in t.split("\n") if l.strip())


RE_ANOTACION = re.compile(r"ANOTACION:\s*Nro\s*(\d+)\s*Fecha:\s*([\d/-]+)", re.IGNORECASE)
RE_CANCELA = re.compile(r"Se cancela anotaci\S*n No:\s*([\d,\s]+)", re.IGNORECASE)


def parsear_anotacion(valor_html):
    """Convierte una entrada de textoAnotaciones de VUR en un dict."""
    texto = html_a_texto(valor_html)
    m = RE_ANOTACION.search(texto)
    espec = re.search(r"ESPECIFICACION:\s*(.+)", texto)
    naturaleza = re.search(r"NATURALEZA JURIDICA DEL ACTO:\s*(.+)", texto)
    cancela = []
    for grupo in RE_CANCELA.findall(texto):
        cancela.extend(int(n) for n in re.findall(r"\d+", grupo))
    lineas = texto.split("\n")
    especificacion = espec.group(1).strip() if espec else ""
    if not especificacion and naturaleza:
        # Algunas anotaciones antiguas no traen ESPECIFICACION y solo dicen la naturaleza
        # del acto; se usa esa para clasificarlas y no dejar pasar una valorización.
        especificacion = naturaleza.group(1).strip()
    return {
        "nro": int(m.group(1)) if m else None,
        "fecha": m.group(2) if m else "",
        "especificacion": especificacion,
        "sin_especificacion": not espec,
        "naturaleza": naturaleza.group(1).strip() if naturaleza else "",
        "cancela": cancela,
        "de": [l[3:].strip() for l in lineas if l.startswith("DE:")],
        "a": [l[2:].strip() for l in lineas if l.startswith("A:")],
        "texto": texto,
    }


# Códigos de especificación de VUR observados en gravámenes de valorización reales:
#   430 "GRAVAMEN POR CONTRIBUCION DE VALORIZACION" (tabla de códigos anterior)
#   0212 "VALORIZACION ..." (tabla actual)
# Una anotación que menciona valorización con otro código no se descarta: se
# cuenta como gravamen pero el resultado queda marcado para revisión.
CODIGOS_GRAVAMEN_VALORIZACION = {"430", "212"}


def _codigo(especificacion):
    m = re.match(r"\s*(\d{3,4})\b", especificacion or "")
    return m.group(1).lstrip("0") if m else ""


def es_cancelacion(anotacion):
    """Anotación que cancela otra: código 08xx, '(CANCELACION)' o 'Se cancela anotación No:'."""
    espec = normalizar(anotacion["especificacion"])
    return bool(anotacion["cancela"]) or "(CANCELACION)" in espec or bool(re.match(r"0?8\d\d\b", espec)) \
        or bool(re.match(r"(\d{3,4}\s+)?CANCELACION", espec))  # p. ej. "0723 CANCELACION VALORIZACION"


def _menciona_valorizacion(anotacion):
    espec = normalizar(anotacion["especificacion"])
    return any(normalizar(p) in espec for p in PALABRAS_VALORIZACION)


def _menciona_afectacion(anotacion):
    texto = normalizar(anotacion["especificacion"] + " " + anotacion["naturaleza"])
    return "AFECTACION" in texto or "INENAJENAB" in texto or "MEDIDA CAUTELAR" in texto


def tipo_valorizacion(anotacion):
    """None si no tiene relación con valorización; si no: 'cancelacion', 'afectacion',
    'gravamen' (código conocido) o 'dudosa' (menciona valorización con código no reconocido)."""
    if not _menciona_valorizacion(anotacion):
        return None
    if es_cancelacion(anotacion):
        return "cancelacion"
    if _menciona_afectacion(anotacion):
        return "afectacion"
    if _codigo(anotacion["especificacion"]) in CODIGOS_GRAVAMEN_VALORIZACION:
        return "gravamen"
    return "dudosa"


def es_valorizacion(anotacion):
    """Gravamen por valorización (no sus cancelaciones ni las afectaciones de inenajenabilidad)."""
    return tipo_valorizacion(anotacion) in ("gravamen", "dudosa")


def _referencias(texto):
    """Números de obra/resolución citados ('OBRA 808', 'RESOLUCION 939') para vincular cancelaciones."""
    return set(re.findall(r"(?:OBRA|RESOLUCION|RES\.?|OFICIO)\s*(?:N[O°.]*\s*)?(\d{2,})", normalizar(texto)))


def analizar_folio(anotaciones):
    """Clasifica las anotaciones de un folio.

    Devuelve dict con:
      valorizacion: gravámenes por valorización (con 'cancelada_por' y 'vinculo':
                    'numero' si la cancelación cita su número, 'texto' si se dedujo).
      afectaciones: afectaciones de inenajenabilidad por valorización (no deciden el
                    estado; se informan al profesional), con 'cancelada_por'.
      cancelaciones: cancelaciones de los gravámenes, o de una valorización
                    inscrita en otro folio (la cita no existe en este folio).
      estado: 'GRAVADA', 'GRAVAMEN CANCELADO' o 'NO GRAVADA'.

    Algunas cancelaciones de VUR no traen "Se cancela anotación No:" (p. ej.
    01N-5052127, anotación 6 "CANCELACION ... DE VALORIZACION OBRA 808"). En
    ese caso se vinculan a las anotaciones vigentes anteriores del mismo tipo
    que citen la misma obra o resolución; si no comparten referencia, a todas
    las vigentes anteriores de ese tipo. El vínculo queda marcado 'texto'.
    """
    valorizacion = [dict(a, cancelada_por=None, vinculo="") for a in anotaciones if es_valorizacion(a)]
    afectaciones = [dict(a, cancelada_por=None, vinculo="") for a in anotaciones
                    if tipo_valorizacion(a) == "afectacion"]
    por_nro = {a["nro"]: a for a in valorizacion + afectaciones}
    nros_folio = {a["nro"] for a in anotaciones}
    cancelaciones = []
    for a in anotaciones:
        if not es_cancelacion(a):
            continue
        objetivos = [n for n in a["cancela"] if n in por_nro]
        vinculo = "numero"
        if not objetivos and a["cancela"]:
            if _menciona_valorizacion(a) and not any(n in nros_folio for n in a["cancela"]):
                cancelaciones.append(a)  # cita una anotación de otro folio (gravamen en la matriz)
            continue
        if not objetivos and _menciona_valorizacion(a):
            grupo = afectaciones if _menciona_afectacion(a) else valorizacion
            previas = [x for x in grupo if x["cancelada_por"] is None and (x["nro"] or 0) < (a["nro"] or 10**9)]
            refs = _referencias(a["texto"])
            comunes = [x for x in previas if refs & _referencias(x["texto"])]
            objetivos = [x["nro"] for x in (comunes or previas)]
            vinculo = "texto"
            if not objetivos:
                if not _menciona_afectacion(a):
                    cancelaciones.append(a)  # cancela una valorización inscrita en otro folio
                continue
        if not objetivos:
            continue
        if any(por_nro[n] in valorizacion for n in objetivos):
            cancelaciones.append(a)
        for n in objetivos:
            por_nro[n]["cancelada_por"] = a["nro"]
            por_nro[n]["vinculo"] = vinculo
            por_nro[n]["cancelacion"] = a  # se informa junto a la afectación cancelada
    if not valorizacion:
        estado = "GRAVAMEN CANCELADO" if cancelaciones else "NO GRAVADA"
    elif any(a["cancelada_por"] is None for a in valorizacion):
        estado = "GRAVADA"
    else:
        estado = "GRAVAMEN CANCELADO"
    return {"valorizacion": valorizacion, "afectaciones": afectaciones,
            "cancelaciones": cancelaciones, "estado": estado}


def analizar_anotaciones(anotaciones):
    """Compatibilidad: (valorizacion, cancelaciones, estado)."""
    r = analizar_folio(anotaciones)
    return r["valorizacion"], r["cancelaciones"], r["estado"]


def describir_no_reconocidas(etiqueta, anotaciones):
    """Explica qué anotaciones no se pudieron leer y si podrían ocultar una valorización.

    Devuelve (motivos, notas). Las entradas vacías que a veces trae VUR se ignoran.
    """
    motivos, notas = [], []
    for i, a in enumerate(anotaciones, 1):
        if not a["texto"].strip():
            continue
        if a["nro"] is None:
            falla = f"entrada {i} sin número de anotación"
        elif not a["especificacion"]:
            falla = f"anotación {a['nro']} sin especificación ni naturaleza del acto"
        else:
            if a.get("sin_especificacion"):
                notas.append(f"{etiqueta}: la anotación {a['nro']} no trae ESPECIFICACION; "
                             "se clasificó por la naturaleza del acto")
            continue
        menciona = any(normalizar(p) in normalizar(a["texto"]) for p in PALABRAS_VALORIZACION)
        aviso = " y MENCIONA VALORIZACIÓN: revisarla en VUR" if menciona else ""
        motivos.append(f"{etiqueta}: {falla}{aviso}")
        notas.append(f"Texto de la {falla} de {etiqueta}: «{a['texto'][:300]}»")
    return motivos, notas


def es_adquisicion(anotacion):
    return "MODO DE ADQUISICION" in normalizar(anotacion["naturaleza"] + " " + anotacion["especificacion"])


def propietarios_en_fecha(anotaciones, nro_limite=None):
    """Personas 'A:' de la última anotación de adquisición (compraventa, adjudicación...)
    anterior a la anotación nro_limite (o la última del folio si no se indica).
    Devuelve (personas, nro_anotacion) o ([], None)."""
    candidatas = [a for a in anotaciones if es_adquisicion(a) and a["a"]
                  and (nro_limite is None or (a["nro"] or 0) < nro_limite)]
    if not candidatas:
        return [], None
    ultima = max(candidatas, key=lambda a: a["nro"] or 0)
    return ultima["a"], ultima["nro"]


def propietarios_hasta_fecha(anotaciones, fecha):
    """Personas 'A:' de la última adquisición inscrita en o antes de 'fecha'. ([], None) si no hay."""
    candidatas = []
    for a in anotaciones:
        f = parsear_fecha(a["fecha"])
        if es_adquisicion(a) and a["a"] and f and f <= fecha:
            candidatas.append((f, a["nro"] or 0, a))
    if not candidatas:
        return [], None
    ultima = max(candidatas, key=lambda x: x[:2])[2]
    return ultima["a"], ultima["nro"]


MAX_MATRICES_PROPIETARIO = 3  # en un englobe se miran como máximo las 3 primeras matrices


def propietario_desde_matrices(api, folio, db_folio, fecha_gravamen):
    """Deduce a quién se le impuso un gravamen que no nombra al propietario.

    Caso: el gravamen está en un folio sin compra anterior (p. ej. 01N-5052127 de
    Bello). Ese folio nació de su matriz, así que el dueño al abrirse es el dueño
    de la matriz en esa fecha. Se busca la última adquisición de la matriz hasta
    la apertura del folio (o hasta la fecha del gravamen si no hay apertura).

    Devuelve una lista de textos "PERSONA (dueño de la matriz X según anotación N)".
    Es siempre una deducción: el resultado debe quedar para revisión.
    """
    corte = parsear_fecha((db_folio or {}).get("fechaAperturaFolio")) or parsear_fecha(fecha_gravamen)
    if not corte:
        return []
    hallados = []
    for m in ((db_folio or {}).get("matriculasMatriz") or [])[:MAX_MATRICES_PROPIETARIO]:
        partes = separar_matricula(m)
        if not partes:
            continue
        ej = api.estado_juridico(*partes)
        anotaciones = [parsear_anotacion(t) for t in (ej or {}).get("textoAnotaciones") or []]
        personas, nro = propietarios_hasta_fecha(anotaciones, corte)
        hallados.extend(f"{limpiar_persona(p)} (dueño de la matriz {m} según su anotación {nro}, "
                        f"deducido)" for p in personas)
    return hallados


def limpiar_persona(texto):
    """'MONTOYA RAMIREZ JOSE   CC 123  X' -> 'MONTOYA RAMIREZ JOSE (CC 123) X'."""
    t = re.sub(r"\s+", " ", texto or "").strip()
    m = re.search(r"\b(CC|NIT|TI|CE|PA)\s*([\d.\-]+)", t)
    if m:
        nombre, resto = t[:m.start()].strip(), t[m.end():].strip()
        t = f"{nombre} ({m.group(1)} {m.group(2)})" + (f" {resto}" if resto else "")
    # En VUR, "X" al final indica titular del derecho real de dominio e "I", dominio incompleto.
    t = re.sub(r"\s+X$", " [titular]", t)
    return re.sub(r"\s+I$", " [dominio incompleto]", t)


def _tipo_documento(p):
    tipo = p.get("tipoDocumento")
    return (tipo.get("nombre") if isinstance(tipo, dict) else tipo) or ""


def formatear_propietarios(propietarios):
    """Lista JSON de VUR -> 'NOMBRE (CC 123) 50% [titular]; ...'."""
    partes = []
    for p in propietarios or []:
        nombre = re.sub(r"\s+", " ", p.get("nombre") or "").strip()
        ident = " ".join(x for x in (_tipo_documento(p), (p.get("numeroDocumento") or "").strip()) if x)
        texto = nombre + (f" ({ident})" if ident else "")
        pct = p.get("porcentajeParticipacion")
        if pct not in (None, ""):
            texto += f" {pct}" if str(pct).endswith("%") else f" {pct}%"
        if (p.get("titular") or "").strip().upper() == "X":
            texto += " [titular]"
        partes.append(texto)
    return "; ".join(partes)


def extraer_campos(db):
    """Columnas de datos básicos a partir del JSON de VUR."""
    props = db.get("propietarios") or []
    primero = props[0] if props else {}
    return {
        "Propietario": re.sub(r"\s+", " ", primero.get("nombre") or "").strip(),
        "Total_Propietarios": len(props),
        "Tipo_Identificacion": _tipo_documento(primero),
        "Numero_Identificacion": (primero.get("numeroDocumento") or "").strip(),
        "Propietarios_Actuales": formatear_propietarios(props),
        "Direccion_Inmueble": re.sub(r"\s+", " ", db.get("direccion") or "").strip(),
        "Matricula_VUR": f"{db.get('circulo', '')}-{db.get('numeroMatricula', '')}",
        "Referencia_Catastral": db.get("referenciaCatastral") or "",
        "Departamento": (db.get("departamento") or {}).get("nombre", ""),
        "Municipio": ((db.get("municipio") or {}).get("nombre") or "").replace("\ufffd", "Ñ"),
        "Nupre": db.get("nupre") or "",
        "Tipo_Predio": db.get("tipoPredio") or "",
        "Estado_Folio": db.get("estadoFolio") or "",
        "Fecha_Apertura_Folio": db.get("fechaAperturaFolio") or "",
        "Matriculas_Matriz": ", ".join(db.get("matriculasMatriz") or []),
    }


LIMITE_CELDA_EXCEL = 32000  # Excel admite 32.767 caracteres por celda.


def recortar(texto):
    return texto if len(texto) <= LIMITE_CELDA_EXCEL else texto[:LIMITE_CELDA_EXCEL] + " …[recortado]"


def leer_excel(ruta):
    df = pd.read_excel(ruta, dtype=str)
    columna = next((c for c in df.columns if "MATRICULA" in normalizar(c)), None)
    if columna is None:
        raise ValueError("No se encontró una columna cuyo encabezado contenga 'Matricula'.")
    datos = df[[columna]].rename(columns={columna: "Matricula_Original"})
    datos["Matricula_Original"] = datos["Matricula_Original"].fillna("").astype(str).str.strip()
    datos = datos[(datos["Matricula_Original"] != "") & (datos["Matricula_Original"].str.lower() != "nan")]
    datos["Matricula_Limpia"] = datos["Matricula_Original"].map(limpiar_matricula)
    return datos.reset_index(drop=True)


def mensaje_corto(error):
    texto = getattr(error, "msg", None) or str(error)
    return texto.strip().split("\n")[0][:200]


# ---------------------------------------------------------------------------
# Revisión de una matrícula y su cadena de matrices (independiente del navegador)
# ---------------------------------------------------------------------------


def revisar_cadena(api, circulo, numero, db_inicial, fecha_limite, log=lambda m: None):
    """Determina si la matrícula tiene o tuvo valorización, subiendo por sus matrices.

    `api` debe ofrecer datos_basicos(circ, num) y estado_juridico(circ, num).
    Recorre las matrices por niveles (todas las del nivel 1, luego las del
    nivel 2...) sin tope de niveles. Una rama se deja de subir solo cuando:
      - el folio no tiene matriz, o
      - el folio se abrió en o antes de la fecha límite del municipio (ningún
        proyecto pudo inscribirse en un folio anterior a él), o
      - ya se encontró la anotación de valorización.
    MAX_FOLIOS_SEGURIDAD solo evita un recorrido descontrolado; si se alcanza,
    el resultado se marca como revisión incompleta.

    Si un folio solo tiene la CANCELACION de una valorización inscrita antes
    (en la matriz), se sigue subiendo para encontrar el gravamen original y
    a quién se le impuso.

    Devuelve dict con estado, matricula (folio con el gravamen), valorizacion,
    cancelaciones, matricula_cancelacion, cadena, niveles y notas.
    """
    cola = deque([(circulo, numero, db_inicial, 0)])
    vistos = {f"{circulo}-{numero}"}
    cadena, notas = [], []
    folios = 0
    cancelacion_previa = None   # (folio, anotaciones) de una cancelación sin su gravamen
    anotaciones_por_folio = {}
    datos_por_folio = {}   # datos básicos de cada folio revisado (para buscar dueños en su matriz)
    afectaciones = []      # (folio, anotación) de afectaciones de inenajenabilidad por valorización
    no_reconocidas = []    # folios con anotaciones que no se pudieron leer
    aviso_sin_fecha = False
    incompleta = False
    nivel_max = 0

    def resultado(estado, matricula="", valorizacion=(), cancelaciones=(), matricula_canc="", nivel=None):
        return {"estado": estado, "matricula": matricula, "nivel": nivel, "valorizacion": list(valorizacion),
                "cancelaciones": list(cancelaciones), "matricula_cancelacion": matricula_canc,
                "cadena": cadena, "niveles": nivel_max, "notas": list(dict.fromkeys(notas)),
                "anotaciones_folio": anotaciones_por_folio.get(matricula, []),
                "afectaciones": afectaciones, "no_reconocidas": no_reconocidas,
                "anotaciones_inicial": anotaciones_por_folio.get(f"{circulo}-{numero}", []),
                "datos_folio": datos_por_folio.get(matricula, {})}

    while cola:
        circ, num, db, nivel = cola.popleft()
        nivel_max = max(nivel_max, nivel)
        etiqueta = f"{circ}-{num}"
        apertura = parsear_fecha(db.get("fechaAperturaFolio"))
        ej = api.estado_juridico(circ, num)
        if ej is None:
            ej = api.estado_juridico(circ, num)  # un reintento ante respuesta vacía
        if ej is None:
            notas.append(f"{etiqueta}: VUR no devolvió el estado jurídico (dos intentos)")
            incompleta = True
        anotaciones = [parsear_anotacion(t) for t in (ej or {}).get("textoAnotaciones") or []]
        anotaciones_por_folio[etiqueta] = anotaciones
        datos_por_folio[etiqueta] = db
        analisis = analizar_folio(anotaciones)
        valorizacion, cancelaciones, estado = analisis["valorizacion"], analisis["cancelaciones"], analisis["estado"]
        afectaciones.extend((etiqueta, a) for a in analisis["afectaciones"])
        motivos_formato, notas_formato = describir_no_reconocidas(etiqueta, anotaciones)
        no_reconocidas.extend(motivos_formato)
        notas.extend(notas_formato)
        folios += 1
        cadena.append(f"{etiqueta} (nivel {nivel}, apertura {db.get('fechaAperturaFolio') or '?'}, "
                      f"{len(anotaciones)} anotaciones: {estado})")
        log(f"  nivel {nivel} {etiqueta}: {estado} ({len(anotaciones)} anotaciones)")

        if valorizacion:
            if cancelacion_previa and estado == "GRAVADA":
                # Gravamen vigente en la matriz pero cancelado en el folio hijo.
                canc_folio, canc = cancelacion_previa
                return resultado("GRAVAMEN CANCELADO", etiqueta, valorizacion, canc, canc_folio, nivel)
            return resultado(estado, etiqueta, valorizacion, cancelaciones, etiqueta if cancelaciones else "", nivel)
        if cancelaciones and not cancelacion_previa:
            cancelacion_previa = (etiqueta, cancelaciones)
            notas.append(f"{etiqueta}: tiene cancelación de valorización sin el gravamen original; "
                         "se busca el gravamen en sus matrices")

        matrices = db.get("matriculasMatriz") or []
        if not matrices:
            continue
        if fecha_limite and apertura and apertura <= fecha_limite and not cancelacion_previa:
            notas.append(f"{etiqueta}: folio abierto el {apertura:%d/%m/%Y}, en o antes de la fecha límite "
                         f"{fecha_limite:%d/%m/%Y}; no se suben sus matrices")
            continue
        if not apertura and not aviso_sin_fecha:
            notas.append(f"{etiqueta}: sin fecha de apertura; se revisan sus matrices por precaución")
        if not fecha_limite and not aviso_sin_fecha:
            notas.append("Municipio sin fecha de proyecto: se suben todas las matrices")
            aviso_sin_fecha = True
        for m in matrices:
            if m in vistos:
                continue
            if folios + len(cola) >= MAX_FOLIOS_SEGURIDAD:
                notas.append(f"Tope de seguridad de {MAX_FOLIOS_SEGURIDAD} folios alcanzado; "
                             f"sin revisar: {', '.join(x for x in matrices if x not in vistos)}")
                incompleta = True
                cola.clear()
                break
            partes = separar_matricula(m)
            if not partes:
                notas.append(f"Matriz con formato no reconocido: {m}")
                continue
            vistos.add(m)
            db_m = api.datos_basicos(*partes) or api.datos_basicos(*partes)  # un reintento
            if not db_m:
                notas.append(f"Matriz {m}: VUR no devolvió datos (dos intentos); no se pudo revisar")
                incompleta = True
                continue
            cola.append((partes[0], partes[1], db_m, nivel + 1))

    if cancelacion_previa:
        canc_folio, canc = cancelacion_previa
        notas.append("No se encontró el gravamen original en las matrices revisadas")
        return resultado("GRAVAMEN CANCELADO", "", (), canc, canc_folio)
    return resultado(ESTADO_INCOMPLETO if incompleta else "NO GRAVADA")


# ---------------------------------------------------------------------------
# Excepciones de control
# ---------------------------------------------------------------------------


class SesionExpirada(Exception):
    pass


class NavegadorCerrado(Exception):
    pass


class ProcesoDetenido(Exception):
    pass


class LimiteDiario(Exception):
    pass


class ErrorVUR(Exception):
    pass


# ---------------------------------------------------------------------------
# Cliente VUR (sesión de Chrome + servicios JSON)
# ---------------------------------------------------------------------------

JS_FETCH = r"""
const url = arguments[0];
const listo = arguments[arguments.length - 1];
fetch(url, {credentials: 'include'})
  .then(async r => listo({status: r.status, url: r.url,
                          tipo: r.headers.get('content-type') || '', texto: await r.text()}))
  .catch(e => listo({error: String(e)}));
"""


class VURClient:
    def __init__(self, log):
        self.log = log
        self.driver = None
        self.consultas = 0          # peticiones reales a VUR (cuentan para el cupo diario)
        self.consultas_cache = 0    # respuestas reutilizadas sin consultar a VUR
        self._cache = {}

    # --- Navegador ---------------------------------------------------------

    def iniciar_navegador(self):
        opciones = webdriver.ChromeOptions()
        opciones.add_argument(f"--user-data-dir={PERFIL_CHROME}")
        opciones.add_argument("--start-maximized")
        opciones.set_capability("unhandledPromptBehavior", "accept")
        # Mantiene activo el administrador de contraseñas de Chrome en este perfil, para que
        # rellene el usuario y la clave guardados con "Guardar clave de VUR.bat".
        opciones.add_experimental_option("prefs", {
            "credentials_enable_service": True,
            "profile.password_manager_enabled": True,
        })
        self.driver = webdriver.Chrome(options=opciones)
        self.driver.set_script_timeout(TIMEOUT_API)

    def cerrar_navegador(self):
        if self.driver:
            try:
                self.driver.quit()
            except WebDriverException:
                pass
            self.driver = None

    def vivo(self):
        try:
            _ = self.driver.window_handles
            return True
        except (InvalidSessionIdException, NoSuchWindowException, WebDriverException):
            return False

    def abrir_inicio(self):
        self.driver.switch_to.default_content()
        self.driver.get(URL_INICIO)

    def cerrar_alerta(self):
        try:
            self.driver.switch_to.alert.accept()
        except (NoAlertPresentException, WebDriverException):
            pass

    def recuperar(self):
        if not self.vivo():
            raise NavegadorCerrado("El navegador se cerró o perdió la sesión de Selenium.")
        self.cerrar_alerta()

    # --- Servicios JSON ------------------------------------------------------

    def _asegurar_origen(self):
        """fetch() necesita estar en una página de www.vur.gov.co para enviar la sesión."""
        url = (self.driver.current_url or "").lower()
        if "siteminder" in url or "login" in url:
            raise SesionExpirada("VUR muestra la página de inicio de sesión.")
        if "vur.gov.co" not in url:
            self.abrir_inicio()

    def api_get(self, ruta):
        """Consulta un servicio; si la sesión parece vencida, recarga el inicio y reintenta una vez.

        Se comprobó en VUR que volver a cargar inicio.jsf renueva la sesión de
        los servicios sin pedir usuario y contraseña de nuevo.

        Las respuestas con datos se guardan en memoria durante la ejecución:
        varias matrículas comparten matrices (p. ej. tres de El Retiro suben
        por 017-81855, 017-63800 y 017-7991), y un reintento no vuelve a gastar
        cupo en los folios ya consultados. Las respuestas vacías no se guardan,
        para que un fallo temporal de VUR pueda reintentarse.
        """
        if ruta in self._cache:
            self.consultas_cache += 1
            return self._cache[ruta]
        try:
            datos = self._api_get(ruta)
        except SesionExpirada:
            self.log("Sesión de VUR inactiva; recargando la página de inicio...")
            self.abrir_inicio()
            datos = self._api_get(ruta)
        if datos:
            self._cache[ruta] = datos
        return datos

    def _api_get(self, ruta):
        self._asegurar_origen()
        try:
            r = self.driver.execute_async_script(JS_FETCH, URL_API + ruta)
        except JavascriptException as e:
            raise ErrorVUR(f"Error de JavaScript: {mensaje_corto(e)}") from e
        self.consultas += 1
        if not r or r.get("error"):
            raise ErrorVUR(f"Fallo de red: {(r or {}).get('error')}")
        destino = (r.get("url") or "").lower()
        if "siteminder" in destino or "login" in destino or "json" not in r.get("tipo", ""):
            raise SesionExpirada("La sesión de VUR no está activa (respuesta de inicio de sesión).")
        if r["status"] != 200:
            raise ErrorVUR(f"VUR respondió HTTP {r['status']}")
        datos = json.loads(r["texto"] or "null")
        primero = datos[0] if isinstance(datos, list) and datos else datos
        if isinstance(primero, dict) and primero.get("estado") == -9999:
            raise LimiteDiario("VUR: el usuario excedió el número de peticiones por día.")
        return datos

    def datos_basicos(self, circulo, numero):
        return self.api_get(f"datosBasicos/{TIPO_CONSULTA}/{circulo}-{numero}") or None  # {} = no existe

    def estado_juridico(self, circulo, numero):
        return self.api_get(f"estadoJuridico/{TIPO_CONSULTA}/{circulo}-{numero}") or None


ERRORES_RECUPERABLES = (ErrorVUR, TimeoutException, WebDriverException, json.JSONDecodeError)

# ---------------------------------------------------------------------------
# Identidad visual (Manual de Identidad Visual Gobernación de Antioquia 2026)
# ---------------------------------------------------------------------------
# Colores primarios del manual: verde #018D38, verde oscuro #0B5640, blanco y
# verde menta #3AF9A2 (solo piezas digitales). Secundarios: azul #3561AB,
# naranja #F28E18, morado #8B4A97. Tipografía institucional: Poppins (si no
# está instalada en el equipo se usa Segoe UI).
# El logo NO se dibuja en el código: se carga el archivo oficial que entregue
# la Dirección de Comunicaciones, guardado junto al script como ARCHIVO_LOGO.
# El manual pide mínimo 150 x 67 px en pantalla, sobre fondo blanco sin
# contenedor (o contenedor blanco sobre color) y con área de protección libre.
# Por eso el encabezado es blanco y el logo va directamente sobre él.

COLOR = {
    "verde": "#018D38",
    "verde_oscuro": "#0B5640",
    "menta": "#3AF9A2",
    "blanco": "#FFFFFF",
    "azul": "#3561AB",
    "naranja": "#F28E18",
    "morado": "#8B4A97",
    "fondo": "#F3F6F4",       # neutro de trabajo con leve tinte verde
    "linea": "#D5DED9",
    "texto": "#1C2723",
    "texto_suave": "#5B6B64",
    "rojo": "#B3261E",        # solo para errores (color semántico, no institucional)
}

TITULO_APLICATIVO = "Consulta de valorización en VUR"
SECRETARIA = "Secretaría de Infraestructura"
DEPENDENCIA = "Dirección Administrativa y Financiera"
CREDITO = "Desarrollado por Alexander Gómez A."

CARPETA_APP = Path(__file__).resolve().parent
ARCHIVO_LOGO = CARPETA_APP / "logo_gobernacion.png"
# Organización de la carpeta (ver LEAME.txt): el programa vive en "programa\" y los
# Excel de matrículas y resultados en "consultas\", al lado. Si la carpeta aún no se
# ha organizado, todo sigue funcionando desde la carpeta del script.
CARPETA_RAIZ = CARPETA_APP.parent if CARPETA_APP.name.lower() == "programa" else CARPETA_APP
CARPETA_CONSULTAS = (CARPETA_RAIZ / "consultas" if CARPETA_RAIZ != CARPETA_APP and (CARPETA_RAIZ / "consultas").is_dir()
                     else CARPETA_APP)
# Histórico de consultas. Hoy es una carpeta local; para compartirlo entre funcionarios
# (carpeta de red o nube) basta con definir la variable de entorno VUR_HISTORICO con
# la nueva ruta, sin cambiar el código.
CARPETA_HISTORICO = Path(os.environ["VUR_HISTORICO"]) if os.environ.get("VUR_HISTORICO") else CARPETA_RAIZ / "historico"
CARPETA_FUENTES = CARPETA_APP / "fuentes"  # Poppins (licencia SIL OFL 1.1, ver OFL-Poppins.txt)


def cargar_fuentes_privadas():
    """Registra Poppins solo para este programa (Windows), sin instalarla en el equipo.

    Debe llamarse antes de crear la ventana de Tk. En otros sistemas, o si
    faltan los archivos, la interfaz usa Segoe UI / Helvetica.
    """
    import sys
    if sys.platform != "win32" or not CARPETA_FUENTES.is_dir():
        return 0
    import ctypes
    FR_PRIVATE = 0x10
    cargadas = 0
    for ttf in CARPETA_FUENTES.glob("*.ttf"):
        cargadas += bool(ctypes.windll.gdi32.AddFontResourceExW(str(ttf), FR_PRIVATE, 0))
    return cargadas
URL_LOGO_OFICIAL = "https://www.antioquia.gov.co/images/logos/Logo-gob-antioquia-ant.png"  # 284 x 110 px
LOGO_ALTO_MAX = 110  # px; reducirlo a la mitad lo dejaría bajo el mínimo del manual (150 x 67 px)
# Versión para la interfaz: el mismo logo oficial sin el margen transparente que
# trae el PNG, reducido con alta calidad a 157 x 70 px (cumple el mínimo 150 x 67).
ARCHIVO_LOGO_APP = CARPETA_APP / "logo_gobernacion_app.png"


def asegurar_logo(log=lambda m: None):
    """Descarga el logo oficial del portal de la Gobernación si no está en la carpeta."""
    if ARCHIVO_LOGO.exists():
        return True
    import urllib.request
    try:
        with urllib.request.urlopen(URL_LOGO_OFICIAL, timeout=10) as r:
            datos = r.read(2_000_000)
        if not datos.startswith(b"\x89PNG\r\n\x1a\n"):
            log("El archivo descargado como logo no es un PNG; se omite.")
            return False
        ARCHIVO_LOGO.write_bytes(datos)
        log(f"Logo oficial descargado de antioquia.gov.co en {ARCHIVO_LOGO.name}.")
        return True
    except OSError as e:
        log(f"No se pudo descargar el logo oficial ({type(e).__name__}); se muestra el nombre en texto.")
        return False

# Estado -> (etiqueta corta del contador, color del texto, color de fondo de fila)
ESTILO_ESTADO = {
    "GRAVADA": ("Gravadas", "#A3570A", "#FDEBD3"),  # naranja institucional oscurecido para texto legible
    "GRAVAMEN CANCELADO": ("Canceladas", COLOR["azul"], "#E1E9F6"),
    "NO GRAVADA": ("No gravadas", COLOR["verde"], "#E3F4EA"),
    ESTADO_INCOMPLETO: ("Incompletas", COLOR["morado"], "#EFE4F2"),
    "SIN DATOS EN VUR": ("Sin datos", COLOR["texto_suave"], "#ECEFED"),
    "ERROR": ("Error", COLOR["rojo"], "#F8E1DF"),
}


def exportar_excel(registros, ruta):
    """Escribe el Excel de resultados con encabezado institucional.

    Todas las celdas se escriben como texto o número, nunca como fórmula: un
    texto de VUR que empiece por "=" no se ejecuta al abrir el archivo.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append(COLUMNAS_SALIDA)
    for reg in registros:
        ws.append([reg.get(c, "") for c in COLUMNAS_SALIDA])
    for fila in ws.iter_rows(min_row=2):
        for celda in fila:
            if isinstance(celda.value, str) and celda.value.startswith("="):
                celda.data_type = "s"
    encabezado = PatternFill("solid", fgColor=COLOR["verde_oscuro"][1:])
    for celda in ws[1]:
        celda.fill = encabezado
        celda.font = Font(bold=True, color="FFFFFF")
        celda.alignment = Alignment(vertical="center", wrap_text=True)
    anchos = {"Observaciones": 80, "Cadena_Revisada": 60, "Detalle_Valorizacion": 50,
              "Propietarios_Actuales": 45, "Propietario_Gravado": 40, "Notas_Proceso": 50,
              "Proyectos_Municipio": 45, "Motivos_Revision": 55, "Afectaciones_Valorizacion": 60}
    for i, col in enumerate(COLUMNAS_SALIDA, 1):
        ws.column_dimensions[get_column_letter(i)].width = anchos.get(col, max(12, len(col) + 2))
        if col in anchos:
            for celda in ws.iter_cols(min_col=i, max_col=i, min_row=2):
                for c in celda:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    col_estado = COLUMNAS_SALIDA.index("Estado") + 1
    for fila in range(2, ws.max_row + 1):
        celda = ws.cell(row=fila, column=col_estado)
        estilo = ESTILO_ESTADO.get(celda.value)
        if estilo:
            celda.fill = PatternFill("solid", fgColor=estilo[2][1:])
            celda.font = Font(bold=True, color=estilo[1][1:])
    col_conf = COLUMNAS_SALIDA.index("Confiabilidad") + 1
    for fila in range(2, ws.max_row + 1):
        celda = ws.cell(row=fila, column=col_conf)
        if celda.value == CONFIABILIDAD_REVISAR:
            celda.fill = PatternFill("solid", fgColor="FDEBD3")
            celda.font = Font(bold=True, color="A3570A")
        elif celda.value == CONFIABILIDAD_ALTA:
            celda.font = Font(bold=True, color=COLOR["verde"][1:])
    _hoja_resumen(wb, registros)
    wb.save(ruta)


def _hoja_resumen(wb, registros):
    """Hoja 'Resumen': totales por estado y lista de lo que debe revisar el profesional."""
    from collections import Counter
    from openpyxl.styles import Alignment, Font, PatternFill

    ws = wb.create_sheet("Resumen")  # 2.ª hoja: "Resultados" sigue siendo la primera para otros programas
    wb.active = wb.index(ws)          # pero el archivo abre en el resumen
    titulo = Font(bold=True, size=13, color=COLOR["verde_oscuro"][1:])
    negrita = Font(bold=True)
    encabezado = PatternFill("solid", fgColor=COLOR["verde_oscuro"][1:])
    ws.append([f"{TITULO_APLICATIVO} - {SECRETARIA}"])
    ws["A1"].font = titulo
    ws.append([DEPENDENCIA])
    ws.append(["Generado", f"{datetime.now():%d/%m/%Y %H:%M}"])
    ws.append(["Versión del programa", VERSION_APLICATIVO])
    ws.append(["Matrículas en el archivo", len(registros)])
    ws.append([])
    ws.append(["Estado", "Cantidad"])
    for celda in ws[ws.max_row]:
        celda.fill, celda.font = encabezado, Font(bold=True, color="FFFFFF")
    conteo = Counter(r.get("Estado", "") for r in registros)
    for estado in list(ESTILO_ESTADO) + sorted(e for e in conteo if e not in ESTILO_ESTADO):
        if conteo.get(estado):
            ws.append([estado or "(sin estado)", conteo[estado]])
            estilo = ESTILO_ESTADO.get(estado)
            if estilo:
                ws.cell(ws.max_row, 1).font = Font(bold=True, color=estilo[1][1:])
    revisar = [r for r in registros if r.get("Confiabilidad") == CONFIABILIDAD_REVISAR]
    ws.append(["Requieren revisión del profesional", len(revisar)])
    ws.cell(ws.max_row, 1).font = negrita
    if revisar:
        ws.append([])
        ws.append(["Matrícula", "Municipio", "Estado", "Por qué revisar"])
        for celda in ws[ws.max_row]:
            celda.fill, celda.font = encabezado, Font(bold=True, color="FFFFFF")
        for r in revisar:
            ws.append([r.get("Matricula_Limpia", ""), r.get("Municipio", ""), r.get("Estado", ""),
                       r.get("Motivos_Revision", "")])
            ws.cell(ws.max_row, 4).alignment = Alignment(wrap_text=True, vertical="top")
    for fila in ws.iter_rows():
        for celda in fila:
            if isinstance(celda.value, str) and celda.value.startswith("="):
                celda.data_type = "s"
    for col, ancho in zip("ABCD", (34, 22, 30, 90)):
        ws.column_dimensions[col].width = ancho


# ---------------------------------------------------------------------------
# Comparación de dos Excel de resultados
# ---------------------------------------------------------------------------

CAMPOS_COMPARADOS = [  # (columna, nombre para el usuario)
    ("Estado", "Estado"), ("Confiabilidad", "Confiabilidad"), ("Municipio", "Municipio"),
    ("Matricula_Con_Anotacion", "Folio con la anotación"), ("Nivel_Matriz", "Nivel de matriz"),
    ("Propietario_Gravado", "Propietario gravado"), ("Matricula_Cancelacion", "Folio de la cancelación"),
    ("Afectaciones_Valorizacion", "Afectaciones"), ("Folios_Revisados", "Folios revisados"),
    ("Propietarios_Actuales", "Propietarios actuales"),
]


def _leer_resultados(ruta):
    """Lee la hoja 'Resultados' (o la primera) de un Excel generado por el programa."""
    hojas = pd.ExcelFile(ruta).sheet_names
    df = pd.read_excel(ruta, sheet_name="Resultados" if "Resultados" in hojas else 0, dtype=str).fillna("")
    if "Matricula_Limpia" not in df.columns:
        raise ValueError(f"{Path(ruta).name} no parece un Excel de resultados del programa "
                         "(no tiene la columna Matricula_Limpia).")
    return {fila["Matricula_Limpia"].strip(): fila for fila in df.to_dict("records") if fila["Matricula_Limpia"].strip()}


def _anotaciones_en_cadena(texto):
    """Total de anotaciones que dice la columna Cadena_Revisada ('..., 13 anotaciones: ...')."""
    return sum(int(n) for n in re.findall(r"(\d+) anotaciones", texto or ""))


def _explicar_cambio(antes, ahora):
    """Causa probable de que una matrícula cambie entre dos consultas (es una pista, no una certeza)."""
    causas = []
    if antes.get("Estado") in ("SIN DATOS EN VUR", "ERROR", ESTADO_INCOMPLETO):
        causas.append("la consulta anterior no fue completa (sin datos, error o matrices sin respuesta)")
    if ahora.get("Estado") in ("SIN DATOS EN VUR", "ERROR", ESTADO_INCOMPLETO):
        causas.append("la consulta nueva no fue completa: repetirla")
    n_antes, n_ahora = _anotaciones_en_cadena(antes.get("Cadena_Revisada")), _anotaciones_en_cadena(ahora.get("Cadena_Revisada"))
    if n_antes and n_ahora and n_antes != n_ahora:
        causas.append(f"VUR muestra otra cantidad de anotaciones ({n_antes} -> {n_ahora}): hubo inscripciones nuevas")
    v_antes, v_ahora = antes.get("Version_Aplicativo", ""), ahora.get("Version_Aplicativo", "")
    if v_antes != v_ahora:
        causas.append(f"cambiaron las reglas del programa (versión {v_antes or 'anterior a 2026.10.23'} -> {v_ahora or '?'})")
    return "; ".join(causas) or "sin causa evidente: revisar manualmente en VUR"


def comparar_resultados(ruta_antes, ruta_ahora, ruta_salida):
    """Compara dos Excel de resultados y escribe un Excel con las diferencias.

    Devuelve un resumen (dict). Si los archivos no comparten ninguna matrícula,
    lanza ValueError: casi siempre significa que son de lotes distintos.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    antes, ahora = _leer_resultados(ruta_antes), _leer_resultados(ruta_ahora)
    comunes = [m for m in ahora if m in antes]
    if not comunes:
        raise ValueError("Los dos archivos no tienen ninguna matrícula en común: probablemente son de lotes "
                         "distintos. Elija el resultado anterior del mismo Excel de matrículas.")
    cambios, detalle = [], []
    for m in comunes:
        a, b = antes[m], ahora[m]
        campos = [(col, nombre) for col, nombre in CAMPOS_COMPARADOS
                  if col in a and col in b  # una columna nueva en una sola versión no es un cambio
                  and str(a[col]).strip() != str(b[col]).strip()]
        if not campos:
            continue
        cambios.append([m, b.get("Municipio") or a.get("Municipio", ""), a.get("Estado", ""), b.get("Estado", ""),
                        a.get("Confiabilidad", ""), b.get("Confiabilidad", ""), ", ".join(n for _, n in campos),
                        _explicar_cambio(a, b), b.get("Motivos_Revision", "")])
        for col, nombre in campos:
            detalle.append([m, nombre, a.get(col) or "(vacío)", b.get(col) or "(vacío)"])
    solo_antes = [m for m in antes if m not in ahora]
    solo_ahora = [m for m in ahora if m not in antes]
    cambio_estado = sum(1 for c in cambios if c[2] != c[3])
    resumen = {"comunes": len(comunes), "iguales": len(comunes) - len(cambios), "con_cambios": len(cambios),
               "cambio_estado": cambio_estado, "solo_antes": len(solo_antes), "solo_ahora": len(solo_ahora)}

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    enc = PatternFill("solid", fgColor=COLOR["verde_oscuro"][1:])
    blanco = Font(bold=True, color="FFFFFF")
    ws.append(["Comparación de resultados VUR"])
    ws["A1"].font = Font(bold=True, size=13, color=COLOR["verde_oscuro"][1:])
    for fila in (["Resultado anterior", Path(ruta_antes).name], ["Resultado nuevo", Path(ruta_ahora).name],
                 ["Generado", f"{datetime.now():%d/%m/%Y %H:%M}  (versión {VERSION_APLICATIVO})"], [],
                 ["Matrículas en ambos archivos", resumen["comunes"]], ["Sin ningún cambio", resumen["iguales"]],
                 ["Con algún cambio", resumen["con_cambios"]], ["   de ellas, cambió el estado", cambio_estado],
                 ["Solo en el anterior", resumen["solo_antes"]], ["Solo en el nuevo", resumen["solo_ahora"]]):
        ws.append(fila)
    if solo_antes or solo_ahora:
        ws.append([])
        ws.append(["Solo en el anterior", ", ".join(solo_antes) or "—"])
        ws.append(["Solo en el nuevo", ", ".join(solo_ahora) or "—"])
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 34, 70

    def hoja(titulo, columnas, filas, anchos):
        h = wb.create_sheet(titulo)
        h.append(columnas)
        for celda in h[1]:
            celda.fill, celda.font = enc, blanco
        for f in filas:
            h.append([recortar(str(x)) for x in f])
        for i, ancho in enumerate(anchos):
            h.column_dimensions[chr(65 + i)].width = ancho
            for celda in h.iter_cols(min_col=i + 1, max_col=i + 1, min_row=2):
                for c in celda:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        h.freeze_panes = "B2"
        return h

    hc = hoja("Cambios", ["Matrícula", "Municipio", "Estado antes", "Estado ahora", "Confianza antes",
                          "Confianza ahora", "Qué cambió", "Causa probable", "Motivos de revisión (ahora)"],
              cambios, (14, 18, 22, 22, 12, 12, 34, 55, 55))
    for fila in hc.iter_rows(min_row=2):
        if fila[2].value != fila[3].value:
            for c in fila[2:4]:
                c.font = Font(bold=True, color="A3570A")
    hoja("Detalle", ["Matrícula", "Campo", "Antes", "Ahora"], detalle, (14, 22, 70, 70))
    for h in wb.worksheets:
        for fila in h.iter_rows():
            for celda in fila:
                if isinstance(celda.value, str) and celda.value.startswith("="):
                    celda.data_type = "s"
    wb.save(ruta_salida)
    return resumen


# ---------------------------------------------------------------------------
# Histórico de consultas (evita repetir en VUR matrículas ya consultadas)
# ---------------------------------------------------------------------------
# Estructura de la carpeta:
#   registros\AAAA-MM-DD_HHMMSS_usuario_equipo.jsonl  una línea JSON por consulta a VUR.
#        Cada sesión escribe SOLO su propio archivo y nunca modifica los de otros: así
#        varios funcionarios (o varias ventanas) pueden trabajar a la vez sin dañar nada.
#   Historico_consultas.xlsx  última consulta válida de cada matrícula, para leer y
#        filtrar. Se regenera al terminar cada proceso; no se edita a mano.

VIGENCIA_HISTORICO_DIAS = 30   # un resultado más antiguo se vuelve a consultar en VUR
ESTADOS_REUTILIZABLES = {"NO GRAVADA", "GRAVADA", "GRAVAMEN CANCELADO"}
ESTADOS_VALIDOS = ESTADOS_REUTILIZABLES | {ESTADO_INCOMPLETO}


def _fecha_registro(reg):
    try:
        return datetime.strptime(str(reg.get("Fecha_Consulta", ""))[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def _prioridad(reg):
    """Orden para elegir el registro que representa a una matrícula: primero los
    resultados válidos (un ERROR de hoy no tapa un resultado bueno de ayer), luego el más reciente."""
    return (reg.get("Estado") in ESTADOS_VALIDOS, _fecha_registro(reg) or datetime.min)


def evaluar_historico(reg, ahora=None):
    """¿Se puede reutilizar este registro del histórico sin consultar VUR?

    Devuelve (True/False, explicación para el usuario). Solo se reutiliza si:
    se analizó con las reglas vigentes, el estado es definitivo, la confiabilidad
    es ALTA y la consulta tiene menos de VIGENCIA_HISTORICO_DIAS días.
    """
    if not reg:
        return False, ""
    fecha = _fecha_registro(reg)
    cuando = f"{fecha:%d/%m/%Y}" if fecha else "fecha desconocida"
    quien = reg.get("Funcionario") or "otro usuario"
    if reg.get("Version_Reglas") != VERSION_REGLAS:
        return False, f"consultada el {cuando} con reglas anteriores del programa; se consulta de nuevo"
    if reg.get("Estado") not in ESTADOS_REUTILIZABLES or reg.get("Confiabilidad") != CONFIABILIDAD_ALTA:
        return False, f"la consulta del {cuando} quedó en {reg.get('Estado')} / {reg.get('Confiabilidad')}; se consulta de nuevo"
    dias = ((ahora or datetime.now()) - fecha).days if fecha else 10**6
    if dias > VIGENCIA_HISTORICO_DIAS:
        return False, f"consultada hace {dias} días (vigencia {VIGENCIA_HISTORICO_DIAS}); se consulta de nuevo"
    hace = "hoy" if dias == 0 else f"hace {dias} día{'s' if dias != 1 else ''}"
    return True, f"consultada el {cuando} por {quien} ({hace})"


def _usuario_equipo():
    import getpass
    import platform
    try:
        usuario = getpass.getuser()
    except Exception:  # noqa: BLE001 - nunca debe impedir consultar
        usuario = "usuario"
    return usuario, os.environ.get("COMPUTERNAME") or platform.node() or "equipo"


class Historico:
    def __init__(self, carpeta=None, log=lambda m: None):
        self.carpeta = Path(carpeta or CARPETA_HISTORICO)
        self.dir_registros = self.carpeta / "registros"
        self.archivo_excel = self.carpeta / "Historico_consultas.xlsx"
        self.log = log
        self.ultimo = {}       # matrícula -> registro que la representa
        self.total_lineas = 0
        self._archivo_sesion = None
        self._lock = threading.Lock()

    def cargar(self):
        """Lee todos los registros. Las líneas dañadas se ignoran (y se avisa)."""
        self.ultimo, self.total_lineas, dañadas = {}, 0, 0
        if not self.dir_registros.is_dir():
            return 0
        for archivo in sorted(self.dir_registros.glob("*.jsonl")):
            try:
                lineas = archivo.read_text(encoding="utf-8").splitlines()
            except OSError as e:
                self.log(f"Histórico: no se pudo leer {archivo.name} ({type(e).__name__})")
                continue
            for linea in lineas:
                if not linea.strip():
                    continue
                try:
                    reg = json.loads(linea)
                except json.JSONDecodeError:
                    dañadas += 1
                    continue
                self._indexar(reg)
                self.total_lineas += 1
        if dañadas:
            self.log(f"Histórico: se ignoraron {dañadas} línea(s) dañadas")
        return len(self.ultimo)

    def _indexar(self, reg):
        m = reg.get("Matricula_Limpia")
        if m and (m not in self.ultimo or _prioridad(reg) >= _prioridad(self.ultimo[m])):
            self.ultimo[m] = reg

    def buscar(self, matricula):
        return self.ultimo.get(matricula)

    def agregar(self, registro):
        """Guarda una consulta hecha en VUR (se escribe en disco de inmediato)."""
        reg = {k: v for k, v in registro.items() if k != "Fuente"}
        reg["Version_Reglas"] = VERSION_REGLAS
        with self._lock:
            if self._archivo_sesion is None:
                usuario, equipo = _usuario_equipo()
                limpio = re.sub(r"[^A-Za-z0-9.-]+", "_", f"{usuario}_{equipo}")
                self.dir_registros.mkdir(parents=True, exist_ok=True)
                self._archivo_sesion = self.dir_registros / f"{datetime.now():%Y-%m-%d_%H%M%S}_{limpio}.jsonl"
            with open(self._archivo_sesion, "a", encoding="utf-8") as f:
                f.write(json.dumps(reg, ensure_ascii=False, default=str) + "\n")
            self._indexar(reg)
            self.total_lineas += 1

    def exportar_excel(self):
        """Regenera Historico_consultas.xlsx. Si está abierto en Excel, se avisa y se deja para después."""
        if not self.ultimo:
            return False
        regs = sorted(self.ultimo.values(), key=lambda r: r.get("Matricula_Limpia", ""))
        temporal = self.archivo_excel.with_name("~Historico_consultas_nuevo.xlsx")
        try:
            self.carpeta.mkdir(parents=True, exist_ok=True)
            exportar_excel([dict(r, Fuente="Histórico") for r in regs], temporal)
            os.replace(temporal, self.archivo_excel)
            return True
        except OSError as e:
            self.log(f"Histórico: no se pudo actualizar {self.archivo_excel.name} (¿está abierto en Excel?): "
                     f"{type(e).__name__}. Se actualizará en el próximo proceso.")
            try:
                temporal.unlink()
            except OSError:
                pass
            return False


# ---------------------------------------------------------------------------
# Interfaz Tkinter
# ---------------------------------------------------------------------------

COLUMNAS_TABLA = [  # (columna del registro, título, ancho)
    ("Matricula_Limpia", "Matrícula", 110),
    ("Municipio", "Municipio", 120),
    ("Estado", "Estado", 255),
    ("Confiabilidad", "Confianza", 100),
    ("Matricula_Con_Anotacion", "Folio hallado", 140),
    ("Nivel_Matriz", "Nivel", 60),
    ("Propietario_Gravado", "Propietario gravado", 200),
    ("Folios_Revisados", "Folios", 62),
    ("Fuente", "Fuente", 92),
]


def valores_tabla(reg):
    """Valores de una fila de la tabla; la fuente se muestra corta (el detalle va en el Excel)."""
    valores = []
    for col, _, _ in COLUMNAS_TABLA:
        valor = reg.get(col, "")
        if col == "Fuente" and str(valor).startswith("Histórico"):
            valor = "Histórico"
        valores.append(valor)
    return valores


# ---------------------------------------------------------------------------
# Componentes visuales con relieve (Tk puro, sin librerías adicionales)
# ---------------------------------------------------------------------------
# Tkinter no trae botones redondeados, sombras ni transparencias. Aquí se
# "pintan" píxel a píxel en una imagen (tk.PhotoImage): bordes suavizados
# (antialiasing), sombra difusa debajo y un brillo de 1 px arriba que da la
# sensación de volumen. Como Tk no mezcla transparencias entre widgets, cada
# componente recibe el color del fondo sobre el que va y la mezcla se calcula
# aquí. Las imágenes se guardan en caché: cada combinación se pinta una vez.

SOMBRA_RGB = (11, 40, 28)  # verde casi negro: la sombra conserva el tono institucional


def _rgb(color):
    color = color.lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def _hex(rgb):
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(v + 0.5))) for v in rgb)


def _mezcla(a, b, t):
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t)


def mezclar(c1, c2, t):
    """Color intermedio entre dos colores '#rrggbb' (t=0 -> c1, t=1 -> c2). Simula transparencia."""
    return _hex(_mezcla(_rgb(c1), _rgb(c2), t))


def _distancia(x, y, ancho, alto, radio):
    """Distancia con signo de (x, y) al borde de un rectángulo redondeado: negativa = dentro."""
    qx = abs(x - ancho / 2) - (ancho / 2 - radio)
    qy = abs(y - alto / 2) - (alto / 2 - radio)
    return math.hypot(max(qx, 0.0), max(qy, 0.0)) + min(max(qx, qy), 0.0) - radio


def _cobertura(d):
    return 0.0 if d >= 0.5 else 1.0 if d <= -0.5 else 0.5 - d


def pintar_en(filas, x0, y0, ancho, alto, radio, arriba, abajo=None, borde=None,
              sombra=(0.0, 0.0, 1.0), brillo=0.0, anillo=None):
    """Pinta una forma redondeada sobre una matriz de colores RGB ya existente.

    arriba/abajo: colores del degradado vertical. sombra: (opacidad, desplazamiento
    hacia abajo, difuminado en px). brillo: opacidad del reflejo blanco superior.
    anillo: color del anillo de foco de teclado (accesibilidad), o None.
    """
    arr, abj = _rgb(arriba), _rgb(abajo or arriba)
    bor = _rgb(borde) if borde else None
    ani = _rgb(anillo) if anillo else None
    alfa_s, dy, difu = sombra
    blanco = (255, 255, 255)
    for py, fila in enumerate(filas):
        sy = py + 0.5 - y0
        t = min(max(sy / alto, 0.0), 1.0)
        relleno = _mezcla(arr, abj, t)
        for px in range(len(fila)):
            sx = px + 0.5 - x0
            col = fila[px]
            if alfa_s > 0:
                ds = _distancia(sx, sy - dy, ancho, alto, radio)
                if ds < difu:
                    col = _mezcla(col, SOMBRA_RGB, alfa_s * (1 - max(ds, 0.0) / difu) ** 2)
            d = _distancia(sx, sy, ancho, alto, radio)
            if ani is not None:
                a = min(max(1.5 - abs(d - 3.0), 0.0), 1.0)
                if a:
                    col = _mezcla(col, ani, a)
            cob = _cobertura(d)
            if cob > 0:
                c = relleno
                if brillo:
                    h = cob - _cobertura(_distancia(sx, sy - 1.2, ancho, alto, radio))
                    if h > 0:
                        c = _mezcla(c, blanco, brillo * h)
                if bor is not None:
                    b = cob - _cobertura(d + 1.0)
                    if b > 0:
                        c = _mezcla(c, bor, b)
                col = _mezcla(col, c, cob)
            fila[px] = col
    return filas


def lienzo(ancho, alto, fondo):
    f = _rgb(fondo)
    return [[f] * ancho for _ in range(alto)]


def a_imagen(maestro, filas):
    img = tk.PhotoImage(master=maestro, width=len(filas[0]), height=len(filas))
    img.put(" ".join("{" + " ".join(_hex(c) for c in fila) + "}" for fila in filas), to=(0, 0))
    return img


_CACHE_IMAGENES = {}


def imagen_cache(maestro, clave, fabricar):
    """Devuelve la imagen guardada para 'clave' o la fabrica una sola vez."""
    if clave not in _CACHE_IMAGENES:
        _CACHE_IMAGENES[clave] = a_imagen(maestro, fabricar())
    return _CACHE_IMAGENES[clave]


# Apariencia de los botones: (arriba, abajo, borde, texto, sombra, brillo) por estado.
_VERDE_TXT = "#FFFFFF"
APARIENCIA_BOTON = {
    "primario": {
        "normal": ("#078F3E", "#01722E", "#01652A", _VERDE_TXT, (0.30, 2.0, 5.0), 0.30),
        "hover": ("#0A9A44", "#017A31", "#01692B", _VERDE_TXT, (0.34, 3.0, 5.0), 0.34),
        "presionado": ("#015F26", "#016F2D", "#015524", _VERDE_TXT, (0.14, 0.5, 2.0), 0.0),
        "deshabilitado": ("#E6EBE8", "#E6EBE8", "#D9E0DC", "#97A69F", (0.04, 1.0, 2.0), 0.0),
    },
    "secundario": {
        "normal": ("#FFFFFF", "#F2F6F3", "#CBD5CF", "#0B5640", (0.12, 1.5, 4.0), 0.0),
        "hover": ("#FFFFFF", "#E9F4ED", "#9CC9AE", "#0B5640", (0.17, 2.5, 5.0), 0.0),
        "presionado": ("#DFEDE4", "#EAF3ED", "#9CC9AE", "#0B5640", (0.06, 0.5, 2.0), 0.0),
        "deshabilitado": ("#F5F7F6", "#F5F7F6", "#E1E7E3", "#A7B3AD", (0.0, 0.0, 1.0), 0.0),
    },
    "peligro": {
        "normal": ("#FFFFFF", "#F8F4F3", "#DCC6C3", "#B3261E", (0.12, 1.5, 4.0), 0.0),
        "hover": ("#FFFFFF", "#FBEAE8", "#E2A9A4", "#B3261E", (0.17, 2.5, 5.0), 0.0),
        "presionado": ("#F4D9D6", "#F9E7E5", "#E2A9A4", "#9C1F18", (0.06, 0.5, 2.0), 0.0),
        "deshabilitado": ("#F5F7F6", "#F5F7F6", "#E1E7E3", "#A7B3AD", (0.0, 0.0, 1.0), 0.0),
    },
}


class BotonPro(tk.Canvas):
    """Botón redondeado con degradado, sombra, efecto al pasar el mouse y al presionar.

    Se usa igual que un botón normal: configure(state="disabled"/"normal"),
    cget("state"), invoke(). Responde a Tab, Enter y barra espaciadora.
    """
    ALTO, RADIO, MARGEN = 36, 10, 8

    def __init__(self, padre, texto, comando=None, variante="secundario", fondo="#FFFFFF",
                 fuente=None, ancho=None, state="normal"):
        from tkinter import font as tkfont
        self._fuente = tkfont.Font(root=padre, font=fuente) if fuente else tkfont.nametofont("TkDefaultFont")
        self._ancho = ancho or max(92, self._fuente.measure(texto) + 40)
        m = self.MARGEN
        super().__init__(padre, width=self._ancho + 2 * m, height=self.ALTO + 2 * m, bg=fondo,
                         highlightthickness=0, bd=0, takefocus=1)
        self._fondo, self._variante, self._comando = fondo, variante, comando
        self._estado = state
        self._sobre = self._presionado = self._foco = False
        self._img = self.create_image(0, 0, anchor="nw")
        self._txt = self.create_text((self._ancho + 2 * m) / 2, (self.ALTO + 2 * m) / 2, text=texto,
                                     font=self._fuente)
        for evento, accion in (("<Enter>", self._entrar), ("<Leave>", self._salir),
                               ("<ButtonPress-1>", self._presionar), ("<ButtonRelease-1>", self._soltar),
                               ("<FocusIn>", lambda e: self._enfocar(True)),
                               ("<FocusOut>", lambda e: self._enfocar(False)),
                               ("<KeyPress-space>", lambda e: self.invoke()),
                               ("<KeyPress-Return>", lambda e: self.invoke())):
            self.bind(evento, accion, add="+")
        self._pintar()

    @property
    def activo(self):
        return self._estado != "disabled"

    def _modo(self):
        if not self.activo:
            return "deshabilitado"
        if self._presionado and self._sobre:
            return "presionado"
        return "hover" if self._sobre else "normal"

    def _pintar(self):
        modo = self._modo()
        arriba, abajo, borde, texto, sombra, brillo = APARIENCIA_BOTON[self._variante][modo]
        anillo = COLOR["verde"] if self._foco and self.activo else None
        m, w, h = self.MARGEN, self._ancho, self.ALTO
        clave = ("boton", w, self._variante, modo, bool(anillo), self._fondo)
        img = imagen_cache(self, clave, lambda: pintar_en(
            lienzo(w + 2 * m, h + 2 * m, self._fondo), m, m, w, h, self.RADIO,
            arriba, abajo, borde, sombra, brillo, anillo))
        self.itemconfigure(self._img, image=img)
        self.itemconfigure(self._txt, fill=texto)
        desplazamiento = 1 if modo == "presionado" else 0
        self.coords(self._txt, (w + 2 * m) / 2, (h + 2 * m) / 2 + desplazamiento)
        super().configure(cursor="hand2" if self.activo else "arrow", takefocus=1 if self.activo else 0)

    def _entrar(self, _e):
        self._sobre = True
        self._pintar()

    def _salir(self, _e):
        self._sobre = False
        self._pintar()

    def _presionar(self, _e):
        if self.activo:
            self._presionado = True
            self._pintar()

    def _soltar(self, _e):
        ejecutar = self._presionado and self._sobre and self.activo
        self._presionado = False
        self._pintar()
        if ejecutar:
            self.invoke()

    def _enfocar(self, valor):
        self._foco = valor
        self._pintar()

    def invoke(self):
        if self.activo and self._comando:
            self._comando()

    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        repintar = False
        if "state" in kw:
            self._estado = str(kw.pop("state"))
            self._presionado = self._presionado and self.activo
            repintar = True
        if "text" in kw:
            self.itemconfigure(self._txt, text=kw.pop("text"))
        if "command" in kw:
            self._comando = kw.pop("command")
        resultado = super().configure(**kw) if kw else None
        if repintar:
            self._pintar()
        return resultado

    config = configure

    def cget(self, clave):
        if clave == "state":
            return self._estado
        if clave == "text":
            return self.itemcget(self._txt, "text")
        return super().cget(clave)

    __getitem__ = cget


class Panel(tk.Canvas):
    """Tarjeta blanca con esquinas redondeadas, borde fino y sombra suave.

    Los widgets se ponen dentro de panel.interior. Para que cambiar el tamaño de
    la ventana sea fluido, la tarjeta se arma en 9 piezas: 4 esquinas pintadas y
    bordes/centro como líneas de color (técnica "9-slice").
    """
    RADIO, MARGEN = 14, 10
    SOMBRA = (0.07, 2.0, 8.0)

    def __init__(self, padre, fondo, relleno=(16, 10), ajustar_alto=True):
        super().__init__(padre, bg=fondo, highlightthickness=0, bd=0)
        self._fondo, self._ajustar = fondo, ajustar_alto
        self._px, self._py = relleno
        r, m = self.RADIO, self.MARGEN
        self._k = k = r + m
        self._base = pintar_en(lienzo(2 * k + 1, 2 * k + 1, fondo), m, m, 2 * r + 1, 2 * r + 1, r,
                               COLOR["blanco"], borde="#E1E7E3", sombra=self.SOMBRA)
        b = self._base
        self._esquinas = [a_imagen(self, [fila[x0:x0 + k] for fila in b[y0:y0 + k]])
                          for x0, y0 in ((0, 0), (k + 1, 0), (0, k + 1), (k + 1, k + 1))]
        self.interior = tk.Frame(self, bg=COLOR["blanco"])
        self._ventana = self.create_window(k + self._px - r, k + self._py - r, window=self.interior, anchor="nw")
        self.bind("<Configure>", self._redibujar)
        if ajustar_alto:
            self.interior.bind("<Configure>", self._ajustar_altura, add="+")

    def _desde_borde(self):
        return self._k - self.RADIO  # = MARGEN: espacio de la sombra alrededor

    def _ajustar_altura(self, evento):
        alto = evento.height + 2 * (self._desde_borde() + self._py)
        if int(self.cget("height")) != alto:
            super().configure(height=alto)

    def _redibujar(self, evento):
        w, h, k, b = evento.width, evento.height, self._k, self._base
        if w < 2 * k + 2 or h < 2 * k + 2:
            return
        self.delete("fondo")
        for img, (x, y) in zip(self._esquinas, ((0, 0), (w - k, 0), (0, h - k), (w - k, h - k))):
            self.create_image(x, y, image=img, anchor="nw", tags="fondo")
        for i in range(k):
            self.create_rectangle(k, i, w - k, i + 1, fill=_hex(b[i][k]), width=0, tags="fondo")
            self.create_rectangle(k, h - k + i, w - k, h - k + i + 1, fill=_hex(b[k + 1 + i][k]), width=0,
                                  tags="fondo")
            self.create_rectangle(i, k, i + 1, h - k, fill=_hex(b[k][i]), width=0, tags="fondo")
            self.create_rectangle(w - k + i, k, w - k + i + 1, h - k, fill=_hex(b[k][k + 1 + i]), width=0,
                                  tags="fondo")
        self.create_rectangle(k, k, w - k, h - k, fill=_hex(b[k][k]), width=0, tags="fondo")
        self.tag_lower("fondo")
        e = self._desde_borde()
        self.coords(self._ventana, e + self._px, e + self._py)
        opciones = {"width": max(w - 2 * (e + self._px), 1)}
        if not self._ajustar:
            opciones["height"] = max(h - 2 * (e + self._py), 1)
        self.itemconfigure(self._ventana, **opciones)


class Segmentado(tk.Canvas):
    """Control segmentado estilo Apple (p. ej. Resultados | Registro del proceso)."""
    ALTO, RADIO, MARGEN, HOLGURA = 30, 9, 5, 3

    def __init__(self, padre, opciones, al_cambiar, fondo, fuente, fuente_sel):
        from tkinter import font as tkfont
        self._f, self._fs = tkfont.Font(root=padre, font=fuente), tkfont.Font(root=padre, font=fuente_sel)
        self._seg = max(self._fs.measure(o) for o in opciones) + 32
        self._ancho = self._seg * len(opciones) + 2 * self.HOLGURA
        m = self.MARGEN
        super().__init__(padre, width=self._ancho + 2 * m, height=self.ALTO + 2 * m, bg=fondo,
                         highlightthickness=0, bd=0, takefocus=1, cursor="hand2")
        self._opciones, self._al_cambiar, self._fondo = opciones, al_cambiar, fondo
        self.seleccion, self._foco = 0, False
        self._img = self.create_image(0, 0, anchor="nw")
        self._textos = [self.create_text(m + self.HOLGURA + self._seg * (i + 0.5), m + self.ALTO / 2, text=o)
                        for i, o in enumerate(opciones)]
        self.bind("<ButtonPress-1>", self._clic)
        self.bind("<KeyPress-Left>", lambda e: self.seleccionar(self.seleccion - 1))
        self.bind("<KeyPress-Right>", lambda e: self.seleccionar(self.seleccion + 1))
        self.bind("<FocusIn>", lambda e: self._enfocar(True))
        self.bind("<FocusOut>", lambda e: self._enfocar(False))
        self._pintar()

    def _pintar(self):
        m, hg, sel = self.MARGEN, self.HOLGURA, self.seleccion
        pista = mezclar(self._fondo, COLOR["verde_oscuro"], 0.07)  # gris verdoso translúcido

        def fabricar():
            filas = lienzo(self._ancho + 2 * m, self.ALTO + 2 * m, self._fondo)
            pintar_en(filas, m, m, self._ancho, self.ALTO, self.RADIO, pista,
                      anillo=COLOR["verde"] if self._foco else None)
            return pintar_en(filas, m + hg + self._seg * sel, m + hg, self._seg, self.ALTO - 2 * hg,
                             self.RADIO - 2, "#FFFFFF", "#FBFCFB", "#DCE3DF", (0.14, 1.0, 3.0))
        self.itemconfigure(self._img, image=imagen_cache(
            self, ("seg", self._ancho, sel, self._foco, self._fondo), fabricar))
        for i, t in enumerate(self._textos):
            self.itemconfigure(t, font=self._fs if i == sel else self._f,
                               fill=COLOR["verde_oscuro"] if i == sel else COLOR["texto_suave"])

    def _clic(self, evento):
        i = int((evento.x - self.MARGEN - self.HOLGURA) // self._seg)
        if 0 <= i < len(self._opciones):
            self.seleccionar(i)

    def _enfocar(self, valor):
        self._foco = valor
        self._pintar()

    def seleccionar(self, i):
        i = max(0, min(len(self._opciones) - 1, i))
        if i != self.seleccion:
            self.seleccion = i
            self._pintar()
            self._al_cambiar(i)


class Chip(tk.Canvas):
    """Contador en forma de píldora de color: número en negrita y nombre del estado."""
    ALTO = 26

    def __init__(self, padre, titulo, color, fondo_chip, fondo, fuente, fuente_num):
        from tkinter import font as tkfont
        self._fn = tkfont.Font(root=padre, font=fuente_num)
        f = tkfont.Font(root=padre, font=fuente)
        self._ancho = 12 + self._fn.measure("0000") + 6 + f.measure(titulo) + 12
        super().__init__(padre, width=self._ancho, height=self.ALTO, bg=fondo, highlightthickness=0, bd=0)
        img = imagen_cache(self, ("chip", self._ancho, fondo_chip, fondo), lambda: pintar_en(
            lienzo(self._ancho, self.ALTO, fondo), 0, 0, self._ancho, self.ALTO, self.ALTO / 2,
            fondo_chip, borde=mezclar(fondo_chip, color, 0.16)))
        self.create_image(0, 0, image=img, anchor="nw")
        self._num = self.create_text(12, self.ALTO / 2, text="0", anchor="w", font=self._fn, fill=color)
        self._tit = self.create_text(0, self.ALTO / 2, text=titulo, anchor="w", font=f, fill=color)
        self.valor = 0
        self.fijar(0)

    def fijar(self, n):
        self.valor = n
        self.itemconfigure(self._num, text=str(n))
        self.coords(self._tit, 12 + self._fn.measure(str(n)) + 6, self.ALTO / 2)


class BarraProgreso(tk.Canvas):
    """Barra de progreso delgada con extremos redondeados."""
    GROSOR = 6

    def __init__(self, padre, fondo):
        super().__init__(padre, height=self.GROSOR + 4, bg=fondo, highlightthickness=0, bd=0)
        self.valor, self.maximo = 0, 1
        self.bind("<Configure>", lambda e: self._pintar())

    def fijar(self, valor=None, maximo=None):
        if maximo is not None:
            self.maximo = max(maximo, 1)
        if valor is not None:
            self.valor = valor
        self._pintar()

    def _pintar(self):
        self.delete("all")
        w, y, g = self.winfo_width(), (self.GROSOR + 4) / 2, self.GROSOR
        if w < 2 * g:
            return
        self.create_line(g / 2, y, w - g / 2, y, width=g, capstyle="round", fill="#E3E9E5")
        fraccion = min(max(self.valor / self.maximo, 0), 1)
        if fraccion > 0:
            self.create_line(g / 2, y, g / 2 + (w - g) * fraccion, y, width=g, capstyle="round",
                             fill=COLOR["verde"])


class Insignia(tk.Canvas):
    """Círculo verde con el número del paso (1, 2, 3)."""
    TAM = 24

    def __init__(self, padre, numero, fondo, fuente):
        t = self.TAM
        super().__init__(padre, width=t, height=t, bg=fondo, highlightthickness=0, bd=0)
        img = imagen_cache(self, ("insignia", fondo), lambda: pintar_en(
            lienzo(t, t, fondo), 0, 0, t, t, t / 2, "#0A9341", "#01722E", brillo=0.25))
        self.create_image(0, 0, image=img, anchor="nw")
        self.create_text(t / 2, t / 2, text=str(numero), fill="#FFFFFF", font=fuente)


class Ayuda:
    """Globo de ayuda que aparece al dejar el mouse sobre un control."""

    def __init__(self, widget, texto, fuente):
        self.widget, self.texto, self.fuente = widget, texto, fuente
        self._ventana = self._tarea = None
        widget.bind("<Enter>", self._programar, add="+")
        widget.bind("<Leave>", self._ocultar, add="+")
        widget.bind("<ButtonPress>", self._ocultar, add="+")

    def _programar(self, _e):
        self._tarea = self.widget.after(650, self._mostrar)

    def _mostrar(self):
        if self._ventana:
            return
        x = self.widget.winfo_rootx() + 8
        y = self.widget.winfo_rooty() + self.widget.winfo_height() - 4
        self._ventana = v = tk.Toplevel(self.widget)
        v.wm_overrideredirect(True)
        v.wm_geometry(f"+{x}+{y}")
        marco = tk.Frame(v, bg=COLOR["linea"], padx=1, pady=1)
        marco.pack()
        tk.Label(marco, text=self.texto, bg="#FFFFFF", fg=COLOR["texto"], font=self.fuente, justify="left",
                 wraplength=320, padx=10, pady=6).pack()

    def _ocultar(self, _e=None):
        if self._tarea:
            self.widget.after_cancel(self._tarea)
            self._tarea = None
        if self._ventana:
            self._ventana.destroy()
            self._ventana = None


def colorear_barra_titulo(ventana, fondo="#FFFFFF", texto="#1C2723"):
    """Windows 11: pinta la barra de título del color del encabezado (en Windows 10 no hace nada)."""
    import sys
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ventana.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(ventana.winfo_id())

        def colorref(color):
            r, g, b = _rgb(color)
            return ctypes.c_int(r | (g << 8) | (b << 16))
        for atributo, valor in ((34, colorref(COLOR["linea"])), (35, colorref(fondo)), (36, colorref(texto))):
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, atributo, ctypes.byref(valor), ctypes.sizeof(valor))
    except Exception:  # noqa: BLE001 - estético; nunca debe impedir abrir el programa
        pass


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{TITULO_APLICATIVO} · {SECRETARIA} · {DEPENDENCIA} · Gobernación de Antioquia")
        self.root.geometry("1180x860")
        self.root.minsize(1120, 640)
        self.root.configure(bg=COLOR["fondo"])

        self.cola = queue.Queue()
        self.detener = threading.Event()
        self.lock = threading.Lock()
        self.df = None
        self.resultados = []
        self.ruta_excel = None
        self.hilo = None
        self.cliente = None
        self._logo = None

        self._fuentes()
        self._estilos()
        self.forzar_vur = tk.BooleanVar(value=False)
        self._forzar = False
        self.n_historico = 0
        self._construir_ui()
        colorear_barra_titulo(self.root)
        self.usuario, self.equipo = _usuario_equipo()
        self.historico = Historico(log=self.log_hilo)
        try:
            n = self.historico.cargar()
            self._log(f"Histórico: {n} matrícula(s) consultadas antes, en {self.historico.carpeta}")
        except Exception as e:  # noqa: BLE001 - el histórico nunca debe impedir abrir el programa
            self._log(f"Histórico no disponible: {type(e).__name__}: {e}")
        self.root.protocol("WM_DELETE_WINDOW", self._al_cerrar)
        self.root.after(100, self._procesar_cola)

    # --- Apariencia ------------------------------------------------------------

    def _fuentes(self):
        from tkinter import font as tkfont
        familias = set(tkfont.families(self.root))
        base = next((f for f in ("Poppins", "Segoe UI", "Helvetica") if f in familias), "TkDefaultFont")
        self.f_titulo = (base, 15, "bold")
        self.f_subtitulo = (base, 10)
        self.f_subtitulo_negrita = (base, 10, "bold")
        self.f_pequena_negrita = (base, 9, "bold")
        self.f_seccion = (base, 11, "bold")
        self.f_boton = (base, 10, "bold")
        self.f_texto = (base, 10)
        self.f_pequena = (base, 9)
        self.f_mini = (base, 8)
        self.f_numero = (base, 10, "bold")
        self.f_mono = ("Consolas", 9) if "Consolas" in familias else ("Courier New", 9)

    def _estilos(self):
        st = ttk.Style(self.root)
        st.theme_use("clam")
        c = COLOR
        st.configure(".", background=c["fondo"], foreground=c["texto"], font=self.f_texto)
        st.configure("Treeview", background=c["blanco"], fieldbackground=c["blanco"], rowheight=28,
                     font=self.f_texto, borderwidth=0, relief="flat")
        st.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])  # sin marco alrededor
        st.configure("Treeview.Heading", background="#F7F9F8", foreground=c["texto_suave"],
                     font=self.f_pequena_negrita, relief="flat", borderwidth=0, padding=(8, 7))
        st.map("Treeview.Heading", background=[("active", "#EDF3EF")])
        st.map("Treeview", background=[("selected", c["verde"])], foreground=[("selected", c["blanco"])])
        st.configure("Vertical.TScrollbar", gripcount=0, background="#D5DDD8", troughcolor=c["blanco"],
                     bordercolor=c["blanco"], lightcolor="#D5DDD8", darkcolor="#D5DDD8",
                     arrowcolor=c["texto_suave"], relief="flat", arrowsize=11)
        st.map("Vertical.TScrollbar", background=[("active", "#B9C6BE")])
        st.configure("Opcion.TCheckbutton", background=c["blanco"], foreground=c["texto"], font=self.f_pequena,
                     indicatorcolor=c["blanco"], indicatorbackground=c["blanco"])
        st.map("Opcion.TCheckbutton", indicatorcolor=[("selected", c["verde"]), ("pressed", "#E3F4EA")],
               background=[("active", c["blanco"])])

    def _cargar_logo(self, contenedor):
        """Muestra el logo oficial sin caja: el PNG tiene fondo transparente y va sobre blanco."""
        avisos = []
        asegurar_logo(avisos.append)
        for aviso in avisos:
            self.root.after(500, lambda a=aviso: self._log(a))
        for archivo in (ARCHIVO_LOGO_APP, ARCHIVO_LOGO):
            if not archivo.exists():
                continue
            try:
                img = tk.PhotoImage(file=str(archivo))
                factor = max(1, -(-img.height() // LOGO_ALTO_MAX))  # división hacia arriba
                self._logo = img.subsample(factor, factor) if factor > 1 else img
                tk.Label(contenedor, image=self._logo, bg=COLOR["blanco"], bd=0).pack(side="left", padx=(24, 18),
                                                                                     pady=10)
                return
            except tk.TclError:
                continue
        tk.Label(contenedor, text="GOBERNACIÓN\nDE ANTIOQUIA", bg=COLOR["blanco"], fg=COLOR["verde_oscuro"],
                 font=self.f_seccion, justify="left").pack(side="left", padx=(24, 18), pady=14)
        self.root.after(600, lambda: self._log(
            f"Logo: guarde el archivo oficial en PNG como {ARCHIVO_LOGO.name} junto al script para mostrarlo."))

    # --- Construcción de la interfaz -------------------------------------------

    def _boton(self, padre, texto, comando, variante="secundario", ayuda=None, **kw):
        b = BotonPro(padre, texto, comando, variante=variante, fondo=COLOR["blanco"], fuente=self.f_boton, **kw)
        if ayuda:
            Ayuda(b, ayuda, self.f_pequena)
        return b

    def _paso(self, padre, columna, numero):
        grupo = tk.Frame(padre, bg=COLOR["blanco"])
        grupo.grid(row=0, column=columna, sticky="w")
        Insignia(grupo, numero, COLOR["blanco"], self.f_pequena_negrita).pack(side="left", padx=(0, 4))
        return grupo

    def _divisor(self, padre, columna):
        tk.Frame(padre, bg=COLOR["linea"], width=1).grid(row=0, column=columna, sticky="ns", padx=10, pady=10)

    def _construir_ui(self):
        c = COLOR
        # Encabezado: blanco, como pide el manual para el logo (fondo blanco sin contenedor).
        cab = tk.Frame(self.root, bg=c["blanco"])
        cab.pack(fill="x")
        self._cargar_logo(cab)
        tk.Frame(cab, bg=c["linea"], width=1).pack(side="left", fill="y", pady=20, padx=(0, 18))
        textos = tk.Frame(cab, bg=c["blanco"])
        textos.pack(side="left")
        tk.Label(textos, text=TITULO_APLICATIVO, bg=c["blanco"], fg=c["verde_oscuro"],
                 font=self.f_titulo).pack(anchor="w")
        tk.Label(textos, text=f"{SECRETARIA}  ·  {DEPENDENCIA}", bg=c["blanco"],
                 fg=c["verde"], font=self.f_subtitulo_negrita).pack(anchor="w")
        tk.Frame(self.root, bg=c["verde"], height=2).pack(fill="x")  # filete institucional

        # Pie
        pie = tk.Frame(self.root, bg=c["fondo"])
        pie.pack(side="bottom", fill="x", padx=24, pady=(0, 8))
        tk.Label(pie, text=f"Proyectos: {FUENTE_DATOS}", bg=c["fondo"], fg=c["texto_suave"],
                 font=self.f_mini).pack(side="left")
        tk.Label(pie, text=CREDITO, bg=c["fondo"], fg=c["verde_oscuro"],
                 font=self.f_pequena_negrita).pack(side="right")
        tk.Label(pie, text=f"Versión {VERSION_APLICATIVO}   ·", bg=c["fondo"], fg=c["texto_suave"],
                 font=self.f_mini).pack(side="right", padx=(0, 6))

        cuerpo = tk.Frame(self.root, bg=c["fondo"])
        cuerpo.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        # 1) Barra de acciones: los tres pasos en una sola fila compacta.
        barra = Panel(cuerpo, c["fondo"], relleno=(14, 6))
        barra.pack(fill="x")
        fila = barra.interior
        fila.columnconfigure(0, weight=1)

        g1 = self._paso(fila, 0, 1)
        self.btn_seleccionar = self._boton(
            g1, "Seleccionar Excel", self.seleccionar_excel,
            ayuda="Elija el Excel con la columna de matrículas (por ejemplo 001-123456).")
        self.btn_seleccionar.pack(side="left")
        self._divisor(fila, 1)

        g2 = self._paso(fila, 2, 2)
        self.btn_iniciar = self._boton(
            g2, "Iniciar consultas", self.iniciar, "primario",
            ayuda="Abre Chrome en VUR. Usted inicia sesión y resuelve el captcha; luego el programa consulta "
                  "todas las matrículas y sus matrices.")
        self.btn_iniciar.pack(side="left")
        self.btn_detener = self._boton(g2, "Detener", self.solicitar_detener, "peligro", state="disabled",
                                       ayuda="Se detiene al terminar la matrícula actual y guarda el parcial.")
        self.btn_detener.pack(side="left")
        self.btn_diagnostico = self._boton(
            g2, "Diagnosticar", self.diagnosticar,
            ayuda="Descarga la respuesta completa de VUR de una matrícula a un archivo JSON, para soporte técnico.")
        self.btn_diagnostico.pack(side="left")
        self._divisor(fila, 3)

        g3 = self._paso(fila, 4, 3)
        self.btn_guardar = self._boton(g3, "Guardar Excel", self.guardar,
                                       ayuda="Guarda todos los resultados con el formato institucional. Además, durante la "
                                             f"consulta se guarda un parcial automático cada {GUARDADO_PARCIAL_CADA} "
                                             "matrículas.")
        self.btn_guardar.pack(side="left")
        self.btn_comparar = self._boton(
            g3, "Comparar", self.comparar,
            ayuda="Compara dos Excel de resultados del mismo lote (por ejemplo el de hoy con uno anterior) y "
                  "genera un Excel con lo que cambió y la causa probable.")
        self.btn_comparar.pack(side="left")

        # 2) Estado de la consulta: progreso, avisos y contadores en poco espacio.
        estado = Panel(cuerpo, c["fondo"], relleno=(18, 8))
        estado.pack(fill="x", pady=(0, 0))
        e = estado.interior
        sup = tk.Frame(e, bg=c["blanco"])
        sup.pack(fill="x")
        self.lbl_archivo = tk.Label(sup, text="Sin archivo cargado", bg=c["blanco"], fg=c["texto_suave"],
                                    font=self.f_pequena)
        self.lbl_archivo.pack(side="left", padx=(0, 14))
        self.lbl_progreso = tk.Label(sup, text="Sin consulta en curso", bg=c["blanco"], fg=c["texto"],
                                     font=self.f_subtitulo_negrita)
        self.lbl_progreso.pack(side="left")
        self.lbl_revisar = tk.Label(sup, text="", bg=c["blanco"], fg="#A3570A", font=self.f_pequena_negrita)
        self.lbl_revisar.pack(side="left", padx=(16, 0))
        self.n_revisar = 0
        self.lbl_consultas = tk.Label(sup, text="", bg=c["blanco"], fg=c["texto_suave"], font=self.f_pequena)
        self.lbl_consultas.pack(side="right")
        self.progreso = BarraProgreso(e, c["blanco"])
        self.progreso.pack(fill="x", pady=(6, 8))
        chips = tk.Frame(e, bg=c["blanco"])
        chips.pack(fill="x")
        self.contadores = {}
        for estado_reg, (titulo, color, fondo) in ESTILO_ESTADO.items():
            chip = Chip(chips, titulo, color, fondo, c["blanco"], self.f_pequena, self.f_numero)
            chip.pack(side="left", padx=(0, 8))
            self.contadores[estado_reg] = chip
        self.chk_forzar = ttk.Checkbutton(chips, text="Forzar consulta en VUR", variable=self.forzar_vur,
                                          style="Opcion.TCheckbutton", takefocus=1)
        self.chk_forzar.pack(side="right")
        Ayuda(self.chk_forzar, "Marcado: consulta todas las matrículas en VUR aunque estén en el histórico. "
                               f"Sin marcar: las consultadas hace menos de {VIGENCIA_HISTORICO_DIAS} días con "
                               "resultado ALTA se toman del histórico, sin gastar cupo de VUR.", self.f_pequena)

        # 3) Resultados: ocupan todo el espacio restante.
        resultados = Panel(cuerpo, c["fondo"], relleno=(2, 2), ajustar_alto=False)
        resultados.pack(fill="both", expand=True)
        r = resultados.interior
        r.rowconfigure(2, weight=1)
        r.columnconfigure(0, weight=1)
        cab_r = tk.Frame(r, bg=c["blanco"])
        cab_r.grid(row=0, column=0, sticky="ew", padx=(8, 14), pady=(2, 2))
        self.segmentos = Segmentado(cab_r, ["Resultados", "Registro del proceso"], self._mostrar_vista,
                                    c["blanco"], self.f_texto, self.f_subtitulo_negrita)
        self.segmentos.pack(side="left")
        tk.Label(cab_r, text="Doble clic o Enter sobre una fila para ver el detalle completo",
                 bg=c["blanco"], fg=c["texto_suave"], font=self.f_pequena).pack(side="right")
        tk.Frame(r, bg=c["linea"], height=1).grid(row=1, column=0, sticky="ew")

        self.vista_tabla = tk.Frame(r, bg=c["blanco"])
        self.vista_tabla.grid(row=2, column=0, sticky="nsew", padx=(1, 1), pady=(0, 10))
        self.tabla = ttk.Treeview(self.vista_tabla, columns=[col[0] for col in COLUMNAS_TABLA], show="headings",
                                  selectmode="browse")
        for col, titulo, ancho in COLUMNAS_TABLA:
            self.tabla.heading(col, text=titulo.upper(), anchor="w")
            self.tabla.column(col, width=ancho, anchor="w", stretch=col in ("Propietario_Gravado", "Estado"))
        for estado_reg, (_, color, fondo) in ESTILO_ESTADO.items():
            self.tabla.tag_configure(estado_reg, background=fondo, foreground=COLOR["texto"])
        barra_v = ttk.Scrollbar(self.vista_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra_v.set)
        barra_v.pack(side="right", fill="y")  # primero la barra: así nunca queda fuera de la ventana
        self.tabla.pack(side="left", fill="both", expand=True)
        self.tabla.bind("<Double-1>", self._ver_detalle)
        self.tabla.bind("<Return>", self._ver_detalle)
        # En Tk 8.6 el color de estado de la fila tapa el resaltado de selección; por eso
        # la fila elegida recibe temporalmente la etiqueta "seleccion" (verde, texto blanco).
        self.tabla.tag_configure("seleccion", background=c["verde"], foreground=c["blanco"])
        self._fila_marcada = None  # (iid, etiquetas originales)
        self.tabla.bind("<<TreeviewSelect>>", self._marcar_seleccion)
        self.lbl_vacio = tk.Label(self.vista_tabla, bg=c["blanco"], fg=c["texto_suave"], font=self.f_texto,
                                  justify="center", text="Aún no hay resultados.\n"
                                  "Seleccione el Excel de matrículas (paso 1) y pulse «Iniciar consultas» (paso 2).")
        self.lbl_vacio.place(relx=0.5, rely=0.45, anchor="center")

        self.vista_log = tk.Frame(r, bg=c["blanco"])
        self.vista_log.grid(row=2, column=0, sticky="nsew", padx=(1, 1), pady=(0, 10))
        self.txt_log = tk.Text(self.vista_log, height=12, state="disabled", font=self.f_mono, relief="flat",
                               bd=0, bg=COLOR["blanco"], fg=COLOR["texto"], padx=14, pady=10, wrap="word",
                               highlightthickness=0)
        barra_log = ttk.Scrollbar(self.vista_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=barra_log.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        barra_log.pack(side="right", fill="y")
        self.vista_tabla.tkraise()

    def _marcar_seleccion(self, _evento=None):
        if self._fila_marcada and self.tabla.exists(self._fila_marcada[0]):
            self.tabla.item(self._fila_marcada[0], tags=self._fila_marcada[1])
        self._fila_marcada = None
        sel = self.tabla.selection()
        if sel:
            self._fila_marcada = (sel[0], self.tabla.item(sel[0], "tags"))
            self.tabla.item(sel[0], tags=("seleccion",))

    def _mostrar_vista(self, indice):
        (self.vista_tabla if indice == 0 else self.vista_log).tkraise()

    # --- Registro, progreso y tabla ---------------------------------------------

    def _log(self, mensaje):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{datetime.now():%H:%M:%S}] {mensaje}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def log_hilo(self, mensaje):
        self.cola.put(("log", mensaje))

    def _en_ejecucion(self, activo):
        estado = "disabled" if activo else "normal"
        for b in (self.btn_seleccionar, self.btn_iniciar, self.btn_diagnostico, self.btn_guardar, self.btn_comparar):
            b.configure(state=estado)
        self.chk_forzar.configure(state=estado)
        self.btn_detener.configure(state="normal" if activo else "disabled")

    def _sumar_contador(self, estado, delta):
        if estado in self.contadores:
            chip = self.contadores[estado]
            chip.fijar(chip.valor + delta)

    def _texto_revisar(self):
        self.lbl_revisar.configure(
            text=f"⚑ {self.n_revisar} {'requiere' if self.n_revisar == 1 else 'requieren'} revisión del profesional"
            if self.n_revisar else "")

    def _reiniciar_vista(self):
        self.tabla.delete(*self.tabla.get_children())
        self._fila_marcada = None
        for chip in self.contadores.values():
            chip.fijar(0)
        self.progreso.fijar(0, len(self.df))
        self.lbl_progreso.configure(text=f"0 de {len(self.df)} matrículas")
        self.lbl_consultas.configure(text="")
        self.n_revisar = 0
        self._texto_revisar()
        self.lbl_vacio.configure(text=f"Archivo listo: {len(self.df)} matrículas.\n"
                                      "Pulse «Iniciar consultas» (paso 2) para comenzar.")
        self.lbl_vacio.place(relx=0.5, rely=0.45, anchor="center")

    def _agregar_fila(self, reg):
        self.lbl_vacio.place_forget()
        estado = reg.get("Estado", "")
        valores = valores_tabla(reg)
        iid = self.tabla.insert("", "end", values=valores, tags=(estado,) if estado in ESTILO_ESTADO else ())
        self.tabla.see(iid)
        self._sumar_contador(estado, 1)
        if reg.get("Confiabilidad") == CONFIABILIDAD_REVISAR:
            self.n_revisar += 1
            self._texto_revisar()

    def _actualizar_fila(self, indice, anterior, nuevo):
        """Reemplaza una fila de la tabla tras la segunda pasada y ajusta los contadores."""
        hijos = self.tabla.get_children()
        if indice >= len(hijos):
            return
        estado = nuevo.get("Estado", "")
        etiquetas = (estado,) if estado in ESTILO_ESTADO else ()
        if self._fila_marcada and self._fila_marcada[0] == hijos[indice]:
            self._fila_marcada = (hijos[indice], etiquetas)  # conserva el resaltado de selección
            etiquetas = ("seleccion",)
        self.tabla.item(hijos[indice], values=valores_tabla(nuevo),
                        tags=etiquetas)
        for reg, delta in ((anterior, -1), (nuevo, 1)):
            self._sumar_contador(reg.get("Estado", ""), delta)
            if reg.get("Confiabilidad") == CONFIABILIDAD_REVISAR:
                self.n_revisar += delta
        self._texto_revisar()

    def _ver_detalle(self, _evento=None):
        sel = self.tabla.selection()
        if not sel:
            return
        indice = self.tabla.index(sel[0])
        with self.lock:
            if indice >= len(self.resultados):
                return
            reg = dict(self.resultados[indice])
        c = COLOR
        ven = tk.Toplevel(self.root)
        ven.title(f"Detalle {reg.get('Matricula_Limpia', '')}")
        ven.geometry("860x600")
        ven.configure(bg=c["blanco"])
        cab = tk.Frame(ven, bg=c["blanco"])
        cab.pack(fill="x", padx=24, pady=(16, 12))
        tk.Label(cab, text=f"{reg.get('Matricula_Limpia', '')}  ·  {reg.get('Municipio', '')}", bg=c["blanco"],
                 fg=c["verde_oscuro"], font=self.f_titulo).pack(anchor="w")
        linea = tk.Frame(cab, bg=c["blanco"])
        linea.pack(anchor="w", pady=(6, 0))
        estado = reg.get("Estado", "")
        _, color, fondo = ESTILO_ESTADO.get(estado, ("", c["texto"], "#ECEFED"))
        tk.Label(linea, text=f"  {estado or '—'}  ", bg=fondo, fg=color, font=self.f_pequena_negrita,
                 pady=3).pack(side="left")
        confiab = reg.get("Confiabilidad", "")
        tk.Label(linea, text=f"Confiabilidad: {confiab or '—'}", bg=c["blanco"],
                 fg="#A3570A" if confiab == CONFIABILIDAD_REVISAR else c["texto_suave"],
                 font=self.f_pequena_negrita).pack(side="left", padx=(12, 0))
        tk.Frame(ven, bg=c["verde"], height=2).pack(fill="x")
        marco = tk.Frame(ven, bg=c["blanco"])
        marco.pack(fill="both", expand=True)
        txt = tk.Text(marco, wrap="word", font=self.f_texto, relief="flat", bd=0, padx=24, pady=10,
                      bg=c["blanco"], fg=c["texto"], highlightthickness=0, spacing1=1, spacing3=1)
        barra = ttk.Scrollbar(marco, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=barra.set)
        txt.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        txt.tag_configure("t", font=self.f_pequena_negrita, foreground=c["verde"], spacing1=14, spacing3=4)
        for campo, titulo in (("Fuente", "Fuente del resultado"),
                              ("Motivos_Revision", "Por qué requiere revisión"),
                              ("Segunda_Pasada", "Segunda pasada"),
                              ("Propietarios_Actuales", "Propietarios actuales"),
                              ("Propietario_Gravado", "Propietario gravado"),
                              ("Detalle_Valorizacion", "Anotación de valorización"),
                              ("Observaciones", "Anotaciones completas"),
                              ("Afectaciones_Valorizacion", "Afectaciones por valorización (no definen el paz y salvo)"),
                              ("Cadena_Revisada", "Cadena de matrices revisada"),
                              ("Proyectos_Municipio", "Proyectos del municipio"),
                              ("Notas_Proceso", "Notas del proceso")):
            valor = str(reg.get(campo, "") or "—")
            if campo == "Cadena_Revisada":
                valor = valor.replace(" -> ", "\n→ ")
            txt.insert("end", titulo.upper() + "\n", "t")
            txt.insert("end", valor + "\n")
        txt.configure(state="disabled")
        colorear_barra_titulo(ven)

    def _procesar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self._log(dato)
                elif tipo == "progreso":
                    hechas, consultas, reutilizadas, del_historico = dato
                    total = len(self.df) if self.df is not None else 0
                    self.progreso.fijar(hechas, total)
                    self.lbl_progreso.configure(text=f"{hechas} de {total} matrículas")
                    self.lbl_consultas.configure(
                        text=f"{del_historico} del histórico · {consultas} consultas a VUR · "
                             f"{reutilizadas} reutilizadas")
                elif tipo == "resultado":
                    self._agregar_fila(dato)
                elif tipo == "actualizar":
                    self._actualizar_fila(*dato)
                elif tipo == "estado":
                    self.lbl_progreso.configure(text=dato)
                elif tipo == "confirmar":
                    mensaje, evento, respuesta = dato
                    respuesta["ok"] = messagebox.askokcancel("VUR", mensaje, parent=self.root)
                    evento.set()
                elif tipo == "info":
                    messagebox.showinfo("VUR", dato, parent=self.root)
                elif tipo == "fin":
                    self._en_ejecucion(False)
        except queue.Empty:
            pass
        self.root.after(100, self._procesar_cola)

    def _esperar_confirmacion(self, mensaje):
        """Llamado desde el hilo de trabajo: muestra el diálogo en el hilo principal y espera."""
        evento, respuesta = threading.Event(), {}
        self.cola.put(("confirmar", (mensaje, evento, respuesta)))
        while not evento.wait(0.5):
            if self.detener.is_set():
                return False
        return respuesta.get("ok", False)

    # --- Acciones de botones ------------------------------------------------

    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")], initialdir=str(CARPETA_CONSULTAS),
                                          title="Excel de matrículas")
        if not ruta:
            return
        try:
            df = leer_excel(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{type(e).__name__}: {e}")
            return
        self.df, self.ruta_excel = df, Path(ruta)
        with self.lock:
            self.resultados = []
        repetidas = int(df["Matricula_Limpia"].duplicated().sum())
        extra = f" · {repetidas} repetidas" if repetidas else ""
        unicas = list(dict.fromkeys(df["Matricula_Limpia"]))
        en_historico = [m for m in unicas if self.historico.buscar(m)]
        vigentes = [m for m in en_historico if evaluar_historico(self.historico.buscar(m))[0]]
        hist = f" · {len(en_historico)} en histórico ({len(vigentes)} vigentes)" if en_historico else ""
        self.lbl_archivo.configure(text=f"{self.ruta_excel.name} · {len(df)} matrículas{extra}{hist}")
        self._reiniciar_vista()
        self._log(f"Archivo cargado: {self.ruta_excel.name} ({len(df)} matrículas{extra}{hist})")
        if en_historico:
            self._avisar_historico(en_historico, vigentes)

    def _avisar_historico(self, en_historico, vigentes):
        lista = ", ".join(en_historico[:12]) + (" …" if len(en_historico) > 12 else "")
        texto = (f"{len(en_historico)} matrícula(s) de este archivo ya fueron consultadas antes:\n{lista}\n\n"
                 f"• {len(vigentes)} tienen un resultado vigente (menos de {VIGENCIA_HISTORICO_DIAS} días, "
                 "confiabilidad ALTA): se tomarán del histórico sin consultar VUR.\n"
                 f"• {len(en_historico) - len(vigentes)} se consultarán de nuevo (resultado antiguo o que "
                 "requería revisión).\n\n"
                 "Para consultarlas todas en VUR, marque «Forzar consulta en VUR».\n\n"
                 f"¿Abrir ahora el Excel del histórico ({self.historico.archivo_excel.name})?")
        if messagebox.askyesno("Matrículas ya consultadas", texto, parent=self.root):
            if not self.historico.archivo_excel.exists():
                self.historico.exportar_excel()
            try:
                os.startfile(self.historico.archivo_excel)  # solo Windows
            except (AttributeError, OSError) as e:
                self._log(f"No se pudo abrir el histórico: {type(e).__name__}")

    def _validar_carga(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("VUR", "Primero cargue un Excel con matrículas.")
            return False
        return not (self.hilo and self.hilo.is_alive())

    def iniciar(self):
        if not self._validar_carga():
            return
        self.detener.clear()
        self._forzar = self.forzar_vur.get()  # se lee aquí: las variables de Tk no se tocan desde el hilo
        with self.lock:
            self.resultados = []
        self._reiniciar_vista()
        self._en_ejecucion(True)
        self.hilo = threading.Thread(target=self._flujo_consultas, daemon=True)
        self.hilo.start()

    def diagnosticar(self):
        if not self._validar_carga():
            return
        sugerida = next((m for m in self.df["Matricula_Limpia"] if separar_matricula(m)), "")
        texto = simpledialog.askstring("Diagnóstico VUR", "Matrícula a diagnosticar:",
                                       initialvalue=sugerida, parent=self.root)
        if not texto:
            return
        partes = separar_matricula(limpiar_matricula(texto))
        if not partes:
            messagebox.showwarning("VUR", f"Formato de matrícula inválido: {texto}")
            return
        self.detener.clear()
        self._en_ejecucion(True)
        self.hilo = threading.Thread(target=self._flujo_diagnostico, args=(partes,), daemon=True)
        self.hilo.start()

    def solicitar_detener(self):
        self.detener.set()
        self._log("Detención solicitada: se detendrá al terminar la matrícula actual.")

    def guardar(self):
        with self.lock:
            datos = list(self.resultados)
        if not datos:
            messagebox.showinfo("VUR", "No hay resultados para guardar.")
            return
        sugerido = f"{self.ruta_excel.stem}_resultados.xlsx" if self.ruta_excel else "resultados_vur.xlsx"
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=sugerido,
                                            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        try:
            exportar_excel(datos, ruta)
            self._log(f"Resultados guardados en {ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar (¿está abierto en Excel?):\n{type(e).__name__}: {e}")

    def comparar(self):
        tipos = [("Excel", "*.xlsx")]
        ruta_ahora = filedialog.askopenfilename(title="1 de 2: resultado NUEVO", filetypes=tipos,
                                                initialdir=str(CARPETA_CONSULTAS), parent=self.root)
        if not ruta_ahora:
            return
        ruta_antes = filedialog.askopenfilename(title="2 de 2: resultado ANTERIOR (mismo lote)", filetypes=tipos,
                                                initialdir=str(Path(ruta_ahora).parent), parent=self.root)
        if not ruta_antes:
            return
        if Path(ruta_antes).resolve() == Path(ruta_ahora).resolve():
            messagebox.showwarning("Comparar", "Eligió el mismo archivo dos veces.", parent=self.root)
            return
        salida = Path(ruta_ahora).with_name(f"comparacion_{Path(ruta_ahora).stem}_vs_{Path(ruta_antes).stem}.xlsx")
        try:
            r = comparar_resultados(ruta_antes, ruta_ahora, salida)
        except ValueError as e:
            messagebox.showwarning("Comparar", str(e), parent=self.root)
            return
        except Exception as e:  # archivo abierto en Excel, dañado, etc.
            messagebox.showerror("Comparar", f"No se pudo comparar:\n{type(e).__name__}: {e}", parent=self.root)
            return
        self._log(f"Comparación guardada en {salida}")
        texto = (f"Matrículas en ambos: {r['comunes']}\nSin cambios: {r['iguales']}\n"
                 f"Con cambios: {r['con_cambios']} (cambió el estado en {r['cambio_estado']})\n"
                 f"Solo en el anterior: {r['solo_antes']} · solo en el nuevo: {r['solo_ahora']}\n\n"
                 f"Archivo: {salida.name}\n\n¿Abrirlo ahora?")
        if messagebox.askyesno("Comparación lista", texto, parent=self.root):
            try:
                import os
                os.startfile(salida)  # solo Windows
            except (AttributeError, OSError):
                pass

    def _al_cerrar(self):
        if self.hilo and self.hilo.is_alive():
            if not messagebox.askokcancel("VUR", "Hay un proceso en curso. ¿Cerrar de todos modos?\n"
                                                 "Se guardará el parcial con lo consultado hasta ahora."):
                return
            self.detener.set()
            self.hilo.join(timeout=5)
        if self.cliente:
            self.cliente.cerrar_navegador()  # evita dejar Chrome abierto al cerrar la ventana
        self.root.destroy()

    # --- Hilo de trabajo ----------------------------------------------------

    def _iniciar_sesion(self, cliente):
        self.log_hilo("Iniciando Chrome...")
        cliente.iniciar_navegador()
        cliente.abrir_inicio()
        return self._esperar_confirmacion(
            "Inicie sesión en VUR en la ventana de Chrome (incluido el captcha).\n\n"
            "Pulse Aceptar cuando haya ingresado."
        )

    def _guardar_parcial(self):
        with self.lock:
            datos = list(self.resultados)
        if not datos or not self.ruta_excel:
            return
        ruta = self.ruta_excel.with_name(f"{self.ruta_excel.stem}_resultados_parcial.xlsx")
        try:
            exportar_excel(datos, ruta)
            self.log_hilo(f"Guardado parcial: {ruta.name} ({len(datos)} registros)")
        except Exception as e:
            self.log_hilo(f"No se pudo guardar el parcial (¿está abierto en Excel?): {type(e).__name__}: {e}")

    def _flujo_consultas(self):
        cliente = self.cliente = VURClient(self.log_hilo)
        sesion_vur = False
        self.n_historico = 0
        try:
            self.historico.cargar()  # puede haber consultas nuevas de otra ventana o funcionario
        except Exception as e:  # noqa: BLE001
            self.log_hilo(f"Histórico no disponible: {type(e).__name__}: {e}")
        self._previo = dict(self.historico.ultimo)  # foto del histórico antes de este proceso
        if self._forzar:
            self.log_hilo("Forzar consulta en VUR: no se usa el histórico en este proceso.")
        try:
            for i, fila in self.df.iterrows():
                if self.detener.is_set():
                    raise ProcesoDetenido()
                resultado = None if self._forzar else self._desde_historico(fila)
                if resultado is None:
                    if not sesion_vur:  # Chrome se abre solo si de verdad hay que consultar VUR
                        if not self._iniciar_sesion(cliente):
                            self.log_hilo("Proceso cancelado por el usuario.")
                            return
                        sesion_vur = True
                    resultado = self._consultar_y_registrar(cliente, fila)
                else:
                    self.n_historico += 1
                with self.lock:
                    self.resultados.append(resultado)
                self.cola.put(("resultado", resultado))
                self.cola.put(("progreso", (i + 1, cliente.consultas, cliente.consultas_cache, self.n_historico)))
                if (i + 1) % GUARDADO_PARCIAL_CADA == 0:
                    self._guardar_parcial()
            if sesion_vur:
                self._segunda_pasada(cliente)
            self.log_hilo(f"Consultas finalizadas: {self.n_historico} tomadas del histórico, {cliente.consultas} "
                          f"peticiones a VUR, {cliente.consultas_cache} respuestas reutilizadas.")
            with self.lock:
                por_revisar = sum(1 for x in self.resultados if x.get("Confiabilidad") == CONFIABILIDAD_REVISAR)
            del_hist = (f"{self.n_historico} matrícula(s) se tomaron del histórico sin consultar VUR "
                        "(columna Fuente).\n\n") if self.n_historico else ""
            self.cola.put(("info", f"Consultas finalizadas.\n\n{del_hist}{por_revisar} matrícula(s) requieren "
                                   "revisión del profesional (columna Confiabilidad = REVISAR).\n\n"
                                   "Use 'Guardar Excel' para guardar los resultados."))
        except ProcesoDetenido:
            self.log_hilo("Proceso detenido.")
        except LimiteDiario as e:
            self.log_hilo(f"Proceso detenido: {e}")
            self.cola.put(("info", f"{e}\nLos resultados obtenidos se guardaron en el archivo parcial."))
        except NavegadorCerrado as e:
            self.log_hilo(f"Proceso interrumpido: {e}")
        except Exception as e:
            self.log_hilo(f"Error general: {type(e).__name__}: {mensaje_corto(e)}")
        finally:
            self._guardar_parcial()
            cliente.cerrar_navegador()
            if self.historico.exportar_excel():
                self.log_hilo(f"Histórico actualizado: {self.historico.archivo_excel}")
            self.cola.put(("fin", None))

    def _desde_historico(self, fila):
        """Resultado tomado del histórico si es reutilizable; None si hay que consultar VUR."""
        previo = self._previo.get(fila["Matricula_Limpia"])
        reutilizable, explicacion = evaluar_historico(previo)
        if not reutilizable:
            if explicacion:
                self.log_hilo(f"{fila['Matricula_Limpia']}: {explicacion}")
            return None
        registro = {c: previo.get(c, "") for c in COLUMNAS_SALIDA}
        registro.update(Matricula_Original=fila["Matricula_Original"], Fuente=f"Histórico: {explicacion}")
        self.log_hilo(f"{fila['Matricula_Limpia']}: tomada del histórico ({explicacion}) -> {registro['Estado']}")
        return registro

    def _consultar_y_registrar(self, cliente, fila):
        """Consulta VUR, compara con la consulta anterior (si la hay) y guarda en el histórico."""
        nuevo = self._procesar_matricula(cliente, fila)
        nuevo.update(Fuente="VUR", Funcionario=self.usuario, Equipo=self.equipo)
        previo = self._previo.get(nuevo["Matricula_Limpia"])
        if previo and previo.get("Estado") in ESTADOS_VALIDOS and nuevo.get("Estado") in ESTADOS_VALIDOS \
                and previo["Estado"] != nuevo["Estado"]:
            fecha = _fecha_registro(previo)
            cambio = (f"Cambió respecto a la consulta del {fecha:%d/%m/%Y} " if fecha else
                      "Cambió respecto a la consulta anterior ") + f"({previo['Estado']} -> {nuevo['Estado']})"
            if previo.get("Version_Reglas") == VERSION_REGLAS:
                # Mismas reglas y otro resultado: VUR cambió (inscripción nueva). Lo debe ver el profesional.
                nuevo["Motivos_Revision"] = "; ".join(x for x in (nuevo.get("Motivos_Revision"), cambio) if x)
                nuevo["Confiabilidad"] = CONFIABILIDAD_REVISAR
            else:
                nuevo["Notas_Proceso"] = "; ".join(x for x in (nuevo.get("Notas_Proceso"),
                                                               cambio + " por cambio de reglas del programa") if x)
        try:
            self.historico.agregar(nuevo)
        except OSError as e:
            self.log_hilo(f"No se pudo guardar en el histórico: {type(e).__name__}: {e}")
        return nuevo

    def _segunda_pasada(self, cliente):
        """Vuelve a consultar las matrículas con resultado dudoso por fallas de VUR."""
        import time
        with self.lock:
            pendientes = [i for i, r in enumerate(self.resultados) if necesita_segunda_pasada(r)]
        if not pendientes:
            return
        self.log_hilo(f"Segunda pasada: {len(pendientes)} matrícula(s) con respuesta incompleta de VUR. "
                      f"Se reintentan en {ESPERA_SEGUNDA_PASADA} s.")
        self.cola.put(("estado", f"Segunda pasada en {ESPERA_SEGUNDA_PASADA} s: {len(pendientes)} matrícula(s)…"))
        for _ in range(ESPERA_SEGUNDA_PASADA):
            if self.detener.is_set():
                raise ProcesoDetenido()
            time.sleep(1)
        for n, i in enumerate(pendientes, 1):
            if self.detener.is_set():
                raise ProcesoDetenido()
            with self.lock:
                anterior = dict(self.resultados[i])
            self.cola.put(("estado", f"Segunda pasada: {n} de {len(pendientes)}"))
            nuevo = self._consultar_y_registrar(cliente, self.df.iloc[i])
            if nuevo["Estado"] == anterior["Estado"]:
                nuevo["Segunda_Pasada"] = f"Consultada de nuevo: se confirmó {nuevo['Estado']}"
            else:
                nuevo["Segunda_Pasada"] = f"Consultada de nuevo: cambió de {anterior['Estado']} a {nuevo['Estado']}"
            self.log_hilo(f"{nuevo['Matricula_Limpia']}: {nuevo['Segunda_Pasada']}")
            with self.lock:
                self.resultados[i] = nuevo
            self.cola.put(("actualizar", (i, anterior, nuevo)))
        self._guardar_parcial()

    def _procesar_matricula(self, cliente, fila):
        original, limpia = fila["Matricula_Original"], fila["Matricula_Limpia"]
        registro = {col: "" for col in COLUMNAS_SALIDA}
        registro.update(Matricula_Original=original, Matricula_Limpia=limpia, Intentos=0,
                        Version_Aplicativo=VERSION_APLICATIVO, Confiabilidad=CONFIABILIDAD_REVISAR)

        partes = separar_matricula(limpia)
        if not partes:
            registro.update(Estado="ERROR", Notas_Proceso="Formato de matrícula inválido",
                            Motivos_Revision="Formato de matrícula inválido",
                            Fecha_Consulta=f"{datetime.now():%Y-%m-%d %H:%M:%S}")
            self.log_hilo(f"{original}: formato inválido, se omite.")
            return registro
        circulo, numero = partes

        intentos, ultimo_error = 0, ""
        while intentos < MAX_INTENTOS:
            if self.detener.is_set():
                raise ProcesoDetenido()
            intentos += 1
            registro["Intentos"] = intentos
            self.log_hilo(f"{limpia}: intento {intentos}/{MAX_INTENTOS}")
            try:
                return self._consultar(cliente, circulo, numero, registro)
            except SesionExpirada:
                intentos -= 1  # la expiración de sesión no consume intentos
                cliente.abrir_inicio()
                if not self._esperar_confirmacion("La sesión de VUR expiró. Inicie sesión de nuevo y pulse Aceptar."):
                    raise ProcesoDetenido()
            except ERRORES_RECUPERABLES as e:
                ultimo_error = f"{type(e).__name__}: {mensaje_corto(e)}"
                self.log_hilo(f"{limpia}: {ultimo_error}")
                cliente.recuperar()

        registro.update(Estado="ERROR", Fecha_Consulta=f"{datetime.now():%Y-%m-%d %H:%M:%S}",
                        Motivos_Revision="La consulta a VUR falló; repetirla",
                        Notas_Proceso=f"Falló tras {MAX_INTENTOS} intentos. Último error: {ultimo_error}")
        return registro

    def _consultar(self, cliente, circulo, numero, registro):
        registro["Fecha_Consulta"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
        registro["Version_Aplicativo"] = VERSION_APLICATIVO
        db = cliente.datos_basicos(circulo, numero)
        if not db:
            # VUR a veces responde vacío por una falla momentánea (024-0024613 respondió vacío
            # el 22/10/2026 y con datos el 23/10/2026). Se espera y se reintenta antes de concluir.
            import time
            for espera in (3, 10):
                self.log_hilo(f"{circulo}-{numero}: VUR respondió sin datos; reintento en {espera} s")
                time.sleep(espera)
                db = cliente.datos_basicos(circulo, numero)
                if db:
                    break
        if not db:
            registro.update(Estado="SIN DATOS EN VUR", Notas_Proceso="VUR no devolvió datos para la matrícula",
                            Confiabilidad=CONFIABILIDAD_REVISAR, Version_Aplicativo=VERSION_APLICATIVO,
                            Motivos_Revision="VUR no devolvió datos en 3 intentos: verificar la matrícula y "
                                             "consultarla manualmente (puede ser una falla temporal de VUR)")
            self.log_hilo(f"{circulo}-{numero}: sin datos en VUR")
            return dict(registro)

        campos = extraer_campos(db)
        municipio = campos["Municipio"]
        estado_mpio, etiqueta, info = clasificar_municipio(municipio)
        limite = fecha_limite_municipio(info)
        notas = [f"Hoja Proy x Mpio: {info['nota']}"] if info and info.get("nota") else []

        registro.update(campos)
        registro.update(
            Proyectos_Municipio=describir_proyectos(info),
            Fecha_Limite_Revision=f"{limite:%d/%m/%Y}" if limite else "",
        )

        if estado_mpio in ESTADOS_SIN_CONSULTA:
            registro.update(
                Clasificacion_Municipio=etiqueta,
                Estado="NO GRAVADA" if estado_mpio == "000" else etiqueta,
                Confiabilidad=CONFIABILIDAD_ALTA if estado_mpio == "000" else CONFIABILIDAD_REVISAR,
                Motivos_Revision="" if estado_mpio == "000" else f"Municipio {etiqueta} en la hoja Proy x Mpio",
                Notas_Proceso="; ".join(notas + [f"Municipio {etiqueta} en {FUENTE_DATOS}; no se consultó estado jurídico"]),
            )
            self.log_hilo(f"{circulo}-{numero}: {registro['Estado']} | {municipio}")
            return dict(registro)

        motivos = []
        if estado_mpio is None:
            motivos.append("Municipio no está en la hoja Proy x Mpio")
        if estado_mpio == "CONDONADA":
            motivos.append("Municipio con proyecto condonado")
        if estado_mpio is None:
            etiqueta = "SIN CLASIFICAR"
            notas.append(f"Municipio '{municipio or 'no leído'}' no está en la hoja Proy x Mpio; se consultó por precaución")

        r = revisar_cadena(cliente, circulo, numero, db, limite, self.log_hilo)
        gravados, textos, detalle = [], [], []
        for a in r["valorizacion"]:
            if a["a"]:
                gravados.extend(limpiar_persona(p) for p in a["a"])
            else:
                # La anotación no nombra al propietario ("A:"). 1) Se toma el dueño a esa fecha:
                # la última adquisición anterior en el mismo folio. 2) Si el folio no tiene
                # compra anterior, el dueño de su matriz cuando el folio se abrió.
                personas, nro_adq = propietarios_en_fecha(r["anotaciones_folio"], a["nro"])
                desde_matriz = [] if personas else propietario_desde_matrices(
                    cliente, r["matricula"], r.get("datos_folio"), a["fecha"])
                if personas:
                    gravados.extend(f"{limpiar_persona(p)} (propietario según anotación {nro_adq})" for p in personas)
                    motivos.append("Propietario gravado deducido de la adquisición anterior (VUR no lo indica)")
                elif desde_matriz:
                    gravados.extend(desde_matriz)
                    motivos.append("Propietario gravado deducido de la matriz del folio (VUR no lo indica y el "
                                   "folio no tiene compra anterior): verificar")
                else:
                    motivos.append("Propietario gravado no identificado")
                    notas.append(f"Anotación {a['nro']} de {r['matricula']}: VUR no indica el propietario; "
                                 "no hay adquisición anterior en ese folio ni en su matriz")
            marca = f" [cancelada por la anotación {a['cancelada_por']}]" if a["cancelada_por"] else ""
            if a.get("vinculo") == "texto":
                marca += " (cancelación deducida del texto: la anotación no cita el número; verificar)"
                motivos.append(f"Cancelación de la anotación {a['nro']} deducida del texto (sin número)")
            if tipo_valorizacion(a) == "dudosa":
                motivos.append(f"Anotación {a['nro']}: código de valorización no reconocido ({_codigo(a['especificacion']) or 'sin código'})")
            textos.append(f"{r['matricula']} - {a['texto']}{marca}")
            detalle.append(f"{r['matricula']} anotación {a['nro']} del {a['fecha']}: {a['especificacion']}{marca}")
        folio_canc = r["matricula_cancelacion"] or r["matricula"]
        for a in r["cancelaciones"]:
            textos.append(f"{folio_canc} - {a['texto']}")
            detalle.append(f"{folio_canc} anotación {a['nro']} del {a['fecha']}: {a['especificacion']}")
        if not registro.get("Propietarios_Actuales"):
            personas, nro_adq = propietarios_en_fecha(r["anotaciones_inicial"])
            if personas:
                registro["Propietarios_Actuales"] = "; ".join(limpiar_persona(p) for p in personas) + \
                    f" (según anotación {nro_adq} del estado jurídico; datos básicos no trae propietarios)"
                registro["Propietario"] = limpiar_persona(personas[0])
                registro["Total_Propietarios"] = len(personas)
                motivos.append("Propietarios actuales tomados del estado jurídico (datos básicos no los trae)")
        if r["estado"] == "NO GRAVADA":
            notas.append("NO ESTÁ GRAVADA EN VUR")
        elif r["estado"] == ESTADO_INCOMPLETO:
            notas.append("No se halló valorización, pero la revisión de matrices quedó incompleta: verificar manualmente")
            motivos.append("Revisión de matrices incompleta")
        if r["estado"] == "GRAVAMEN CANCELADO" and not r["valorizacion"]:
            motivos.append("Se halló la cancelación pero no el gravamen original")
        afect = []
        for folio, a in r["afectaciones"]:
            marca = f" [cancelada por la anotación {a['cancelada_por']}]" if a["cancelada_por"] else " [VIGENTE]"
            if a.get("vinculo") == "texto":
                marca += " (cancelación deducida del texto; verificar)"
            bloque = f"{folio} anotación {a['nro']} del {a['fecha']}: {a['especificacion']}{marca}\n{a['texto']}"
            canc = a.get("cancelacion")
            if canc:
                bloque += (f"\n--- Cancelación: {folio} anotación {canc['nro']} del {canc['fecha']}: "
                           f"{canc['especificacion']}\n{canc['texto']}")
            afect.append(bloque)
        if afect:
            # Afectación vigente, o cuya cancelación se dedujo del texto -> revisión del profesional.
            # Afectación cancelada con número -> no exige revisión, pero queda a la vista.
            dudosas = [a for _, a in r["afectaciones"] if not a["cancelada_por"] or a.get("vinculo") == "texto"]
            if dudosas:
                motivos.append("Tiene afectación por valorización VIGENTE (o su cancelación no es clara): no define "
                               "el paz y salvo, pero debe tenerse en cuenta")
            else:
                notas.append("Tuvo afectación por valorización ya cancelada (ver Afectaciones_Valorizacion); "
                             "no define el paz y salvo")
        if r["no_reconocidas"]:
            motivos.append("Anotaciones que no se pudieron leer: " + "; ".join(r["no_reconocidas"]))
        motivos = list(dict.fromkeys(motivos))

        registro.update(
            Clasificacion_Municipio=etiqueta,
            Estado=r["estado"],
            Matricula_Con_Anotacion=r["matricula"],
            Nivel_Matriz="" if r["nivel"] is None else r["nivel"],
            Detalle_Valorizacion=recortar(" | ".join(detalle)),
            Propietario_Gravado="; ".join(dict.fromkeys(gravados)),
            Matricula_Cancelacion=r["matricula_cancelacion"],
            Observaciones=recortar("\n\n".join(textos)),
            Afectaciones_Valorizacion=recortar("\n\n".join(afect)),
            Confiabilidad=CONFIABILIDAD_REVISAR if motivos else CONFIABILIDAD_ALTA,
            Motivos_Revision="; ".join(motivos),
            Cadena_Revisada=recortar(" -> ".join(r["cadena"])),
            Folios_Revisados=len(r["cadena"]),
            Notas_Proceso="; ".join(notas + r["notas"]),
        )
        self.log_hilo(f"{circulo}-{numero}: {r['estado']} | {municipio} | {len(r['cadena'])} folio(s) revisado(s)")
        return dict(registro)

    def _flujo_diagnostico(self, partes):
        """Guarda las respuestas JSON de VUR de una matrícula para revisión."""
        cliente = self.cliente = VURClient(self.log_hilo)
        try:
            if not self._iniciar_sesion(cliente):
                return
            etiqueta = f"{partes[0]}-{partes[1]}"
            ruta = self.ruta_excel.with_name(f"diagnostico_{etiqueta}.json")
            datos = {"datosBasicos": cliente.datos_basicos(*partes),
                     "estadoJuridico": cliente.estado_juridico(*partes)}
            with open(ruta, "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=2)
            self.log_hilo(f"Diagnóstico guardado en {ruta}")
            self.cola.put(("info", f"Diagnóstico guardado en:\n{ruta}"))
        except Exception as e:
            self.log_hilo(f"Error en diagnóstico: {type(e).__name__}: {mensaje_corto(e)}")
        finally:
            cliente.cerrar_navegador()
            self.cola.put(("fin", None))


if __name__ == "__main__":
    cargar_fuentes_privadas()
    raiz = tk.Tk()
    App(raiz)
    raiz.mainloop()
