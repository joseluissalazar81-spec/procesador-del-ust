#!/usr/bin/env python3
"""
revisar_planificaciones.py
Revisa y corrige automáticamente las planificaciones didácticas UST
en todas las carpetas de asignaturas dentro de ~/Desktop/DEL.

Estructura esperada por asignatura:
  ~/Desktop/DEL/<Asignatura>/
    Correcciones/   → planificación ya revisada por DEL (*.xlsx)  [opcional]
    Enviado a DEL/  → planificación original + escala (Escala*.xlsx)
    Revisado/       → carpeta de salida (se crea si no existe)

Lógica de fuente:
  1. Si existe un .xlsx en Correcciones/  → lo usa como base (ya fue revisado por DEL)
  2. Si no hay Correcciones/ o está vacía → busca el .xlsx en Enviado a DEL/
     que NO sea la escala de apreciación, y lo trata como planificación original.

Correcciones aplicadas automáticamente:
  [Síntesis didáctica]
    - Elimina bloques DIAGNÓSTICA y FORMATIVA de la columna
      "Procedimientos de evaluación", dejando solo SUMATIVA.

  [Planificación por unidades]
    Alta prioridad — alineación al programa oficial:
      - Procedimiento 'Cuestionario' → 'Pruebas escritas u orales'
      - Procedimiento 'Tarea'        → 'Producciones del estudiante'
      - Procedimiento 'Prueba'       → 'Pruebas escritas u orales'
      - Instrumento 'Pauta' (exacto) → 'Pauta de observación'
      - Instrumento 'Rúbrica'        → 'Pauta de observación'
        (solo cuando el procedimiento es Producciones/Actuación)
      - Tipo 'Formativo'             → 'Formativa'
      - Tipo 'Diagnóstico'           → 'Diagnóstica'  (normaliza mayúsculas)
      - Datos evaluativos huérfanos en filas sin Tipo de evaluación → eliminados

    Media prioridad — columna J (Medio de entrega):
      - Desarrollo + Sincrónico con evaluación  → 'No aplica'
      - Trabajo independiente / No presencial   → 'Buzón de tareas'
      - Preparación con Diagnóstico             → 'Cuestionario'

    Media prioridad — columna M (Individual/Grupal):
      - Procedimiento contiene 'equipo' o 'grupo'  → 'Grupal'
      - Unidades donde las sumativas son 'Grupal'  → Grupal en toda la unidad
      - Resto                                      → 'Individual'

    Baja prioridad:
      - % evaluación = None en filas de evaluación → 0
"""

import os
import re
import glob
import shutil
import tempfile
import openpyxl
from openpyxl.styles import Font, PatternFill
from datetime import datetime


# ══════════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN Y VERIFICACIÓN DEL PROGRAMA OFICIAL (PDF)
# ══════════════════════════════════════════════════════════════════════════

def extraer_programa_pdf(pdf_path):
    """
    Extrae datos clave del programa de asignatura UST desde un PDF.
    Devuelve un dict con: codigo, asignatura, carrera, creditos, area,
    horas_tpe, unidades (lista), ponderaciones (dict unidad→%), pct_examen.
    """
    try:
        import pdfplumber
    except ImportError:
        return {'_error': 'pdfplumber no instalado'}

    programa = {}
    try:
        texto = ''
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                texto += (page.extract_text() or '') + '\n'

        # Código
        m = re.search(r'CÓDIGO\s*:\s*([A-Z]+-[A-Z0-9]+)', texto, re.IGNORECASE)
        programa['codigo'] = m.group(1).strip() if m else None

        # Asignatura (entre ASIGNATURA: y CÓDIGO:)
        m = re.search(r'ASIGNATURA\s*:\s*(.+?)(?=CÓDIGO\s*:)', texto,
                      re.IGNORECASE | re.DOTALL)
        if m:
            programa['asignatura'] = re.sub(r'\s+', ' ', m.group(1)).strip()

        # Carrera (entre CARRERA: y ASIGNATURA:)
        m = re.search(r'CARRERA\s*:\s*(.+?)(?=ASIGNATURA\s*:)', texto,
                      re.IGNORECASE | re.DOTALL)
        if m:
            programa['carrera'] = re.sub(r'\s+', ' ', m.group(1)).strip()

        # Créditos
        m = re.search(r'(\d+)\s+CRÉDITOS', texto, re.IGNORECASE)
        programa['creditos'] = int(m.group(1)) if m else None

        # Área de conocimiento
        m = re.search(r'AREA[A]?\s+DE\s+CONOCIMIENTO\s*:\s*(.+?)(?:\n|$)',
                      texto, re.IGNORECASE)
        programa['area'] = m.group(1).strip() if m else None

        # Horas TPE
        m = re.search(r'HORAS\s+TPE\*?\s*:\s*(\d+)', texto, re.IGNORECASE)
        programa['horas_tpe'] = int(m.group(1)) if m else None

        # Unidades (sección V UNIDADES DE APRENDIZAJE — primera aparición)
        m_sec = re.search(
            r'V\s+UNIDADES\s+DE\s+APRENDIZAJE\s*\n(.*?)(?:VI\s+DESGLOSE|VII\s+METODOLOGÍA)',
            texto, re.IGNORECASE | re.DOTALL
        )
        bloque_unidades = m_sec.group(1) if m_sec else texto
        unidades_raw = re.findall(
            r'UNIDAD\s+([IVX]+)\s+(.*?)\s+(\d+)\s+HORAS\s+PEDAGÓGICAS',
            bloque_unidades, re.IGNORECASE
        )
        # Deduplicar conservando orden
        vistos = set()
        programa['unidades'] = []
        for n, nom, h in unidades_raw:
            key = n.upper()
            if key not in vistos:
                vistos.add(key)
                programa['unidades'].append({
                    'numero': key,
                    'nombre': re.sub(r'\s+', ' ', nom).strip(),
                    'horas': int(h)
                })

        # Ponderaciones por unidad (sección PONDERACIÓN)
        ponderaciones = re.findall(r'[Uu]nidad\s+([IVX]+)\s+(\d+)%', texto)
        programa['ponderaciones'] = {n.upper(): int(p) for n, p in ponderaciones}

        # Examen LGA
        m = re.search(r'[Ee]xamen[^%\n]{0,60}?(\d+)%', texto)
        programa['pct_examen'] = int(m.group(1)) if m else None

    except Exception as e:
        programa['_error'] = str(e)

    return programa


def verificar_contra_programa(wb, programa, log):
    """
    Cruza la planificación con datos extraídos del programa oficial.
    Verifica: código, créditos, área, nombre asignatura y ponderaciones.
    Devuelve número de discrepancias.
    """
    if not programa or programa.get('_error'):
        log.append('\n  [Verificación contra programa] — datos no disponibles (omitida)')
        return 0

    errores = []
    oks = []

    ws_sint = (wb['Síntesis didáctica']
               if 'Síntesis didáctica' in wb.sheetnames else None)
    ws_plan = (wb['Planificación por unidades']
               if 'Planificación por unidades' in wb.sheetnames else None)

    def sv(ws, r, c):
        v = ws.cell(r, c).value
        return str(v).strip() if v else ''

    def norm(s):
        return re.sub(r'\s+', ' ', str(s or '')).strip().upper()

    STOPWORDS = {'Y', 'E', 'DE', 'DEL', 'LA', 'LAS', 'LOS', 'EL', 'CON',
                 'EN', 'A', 'O', 'U', 'AL'}

    if ws_sint:
        # Código
        if programa.get('codigo'):
            val = sv(ws_sint, 7, 1)
            if norm(val) == norm(programa['codigo']):
                oks.append(f'Código "{programa["codigo"]}"')
            elif val:
                errores.append(
                    f'Código: planificación="{val}" vs programa="{programa["codigo"]}"')

        # Créditos
        if programa.get('creditos') is not None:
            val = ws_sint.cell(7, 3).value
            try:
                if int(val) == int(programa['creditos']):
                    oks.append(f'Créditos ({programa["creditos"]})')
                else:
                    errores.append(
                        f'Créditos: planificación={val} vs programa={programa["creditos"]}')
            except (TypeError, ValueError):
                if val:
                    errores.append(f'Créditos: valor en planificación="{val}" no es numérico')

        # Área de conocimiento
        if programa.get('area'):
            val = sv(ws_sint, 11, 1)
            if val and (norm(programa['area']) in norm(val) or
                        norm(val) in norm(programa['area'])):
                oks.append(f'Área "{programa["area"]}"')
            elif val:
                errores.append(
                    f'Área: planificación="{val}" vs programa="{programa["area"]}"')

        # Nombre asignatura (coincidencia por palabras clave)
        if programa.get('asignatura'):
            val = sv(ws_sint, 4, 4)
            palabras_prog = set(norm(programa['asignatura']).split()) - STOPWORDS
            palabras_plan = set(norm(val).split()) - STOPWORDS
            if palabras_prog:
                ratio = len(palabras_prog & palabras_plan) / len(palabras_prog)
                if ratio >= 0.7:
                    oks.append(f'Nombre asignatura (similitud {ratio:.0%})')
                elif val:
                    errores.append(
                        f'Nombre: planificación="{val}" vs programa="{programa["asignatura"]}"')

    # Ponderaciones sumativas por unidad
    if ws_plan and programa.get('ponderaciones'):
        # Mapa para convertir número arábigo → romano (Unidad 1 → I, etc.)
        ARABE_A_ROMANO = {'1': 'I', '2': 'II', '3': 'III', '4': 'IV',
                          '5': 'V', '6': 'VI', '7': 'VII', '8': 'VIII'}

        def num_unidad(texto):
            """Extrae número de unidad como romano desde texto con arábigo o romano."""
            # Romano: "UNIDAD II", "Unidad III"
            m = re.search(r'\bUNIDAD\s+([IVX]+)\b', str(texto or ''), re.IGNORECASE)
            if m:
                return m.group(1).upper()
            # Arábigo: "Unidad 1", "Unidad 2 – Nombre"
            m = re.search(r'\bUNIDAD\s+(\d+)\b', str(texto or ''), re.IGNORECASE)
            if m:
                return ARABE_A_ROMANO.get(m.group(1))
            return None

        pcts_plan = {}
        unidad_actual = None
        for row in ws_plan.iter_rows(min_row=4, values_only=True):
            unidad = row[COL['UNIDAD'] - 1]
            tipo   = row[COL['TIPO']   - 1]
            pct    = row[COL['PCT']    - 1]
            if unidad:
                unidad_actual = str(unidad)
            if (tipo and 'sumativa' in str(tipo).lower()
                    and isinstance(pct, (int, float)) and pct > 0):
                num = num_unidad(unidad_actual)
                if num:
                    pcts_plan[num] = pcts_plan.get(num, 0) + pct

        for num, pct_prog in programa['ponderaciones'].items():
            pct_p = pcts_plan.get(num)
            if pct_p is None:
                errores.append(
                    f'Unidad {num}: ponderación {pct_prog}% no encontrada en planificación '
                    f'(¿falta la etiqueta "Unidad {num}" en col. Unidad/módulo?)')
            else:
                if pct_p < 1:
                    pct_p = round(pct_p * 100)
                if abs(pct_p - pct_prog) <= 1:
                    oks.append(f'Unidad {num}: ponderación {pct_prog}%')
                else:
                    errores.append(
                        f'Unidad {num}: planificación={pct_p}% vs programa={pct_prog}%')

    log.append('\n  [Verificación contra programa oficial]')
    for o in oks:
        log.append(f'    ✅ {o}')
    for e in errores:
        log.append(f'    ❌ {e}')
    if not oks and not errores:
        log.append('    ⚠️  Sin campos verificables (revisa el PDF extraído)')

    return len(errores)

# ── Constantes de columnas (1-based en openpyxl) ──────────────────────────
COL = {
    'RA':        1,
    'UNIDAD':    2,
    'SEMANA':    3,
    'NOMBRE':    4,
    'MODALIDAD': 5,
    'MOMENTO':   6,
    'ACTIVIDAD': 7,
    'RECURSOS':  8,
    'CONTENIDOS':9,
    'MEDIO':    10,   # J — Medio de entrega
    'TIPO':     11,   # K — Tipo de evaluación
    'PROC':     12,   # L — Procedimiento
    'INDIV':    13,   # M — Individual/Grupal
    'INSTR':    14,   # N — Instrumento
    'PCT':      15,   # O — % evaluación
}

# ── Mapas de corrección ────────────────────────────────────────────────────
PROC_MAP = {
    'Cuestionario': 'Pruebas escritas u orales',
    'Tarea':        'Producciones del estudiante',
    'Prueba':       'Pruebas escritas u orales',
}

INSTR_MAP = {
    'Pauta':   'Pauta de observación',   # solo coincidencia exacta
    'Rúbrica': 'Pauta de observación',   # cuando procedimiento es Producciones/Actuación
}

TIPO_MAP = {
    'Formativo':   'Formativa',
    'Diagnóstico': 'Diagnóstica',
}

PROC_INSTR_RÚBRICA_CONTEXTOS = {
    'Producciones del estudiante',
    'Actuación del estudiante',
    'Informe de análisis de casos',
}

# ── Escala de apreciación estándar UST (41 criterios) ─────────────────────
# Fuente: Escala institucional DEL — idéntica para todas las asignaturas.
# Cada ítem: (sección, N° dentro de sección, criterio, modo de verificación)
#   modo: 'auto' = verificable por el script
#         'manual' = requiere revisión humana
ESCALA_ESTANDAR = [
    # ── Sección 1: Síntesis Didáctica (15 ítems) ──────────────────────────
    ('Síntesis Didáctica', 1,
     'Establece correctamente la escuela (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 2,
     'Menciona correctamente el programa académico (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 3,
     'Nombre de asignatura correcto (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 4,
     'Código correcto (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 5,
     'Versión año correcto (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 6,
     'N° de créditos correcto (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 7,
     'Área del conocimiento OCDE correcta (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 8,
     'Requisitos correctos (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 9,
     'Carga académica bien estructurada (Conforme a Programa Oficial)',
     'auto'),
    ('Síntesis Didáctica', 10,
     'Presenta competencias específicas conforme a programa oficial',
     'auto'),
    ('Síntesis Didáctica', 11,
     'Estructura módulos o unidades con horas pedagógicas correspondientes',
     'auto'),
    ('Síntesis Didáctica', 12,
     'Presenta títulos de módulos o unidades académicas',
     'auto'),
    ('Síntesis Didáctica', 13,
     'Presenta resultados de aprendizaje esperados',
     'auto'),
    ('Síntesis Didáctica', 14,
     'Presenta procedimientos de evaluación acorde a resultados esperados '
     '(solo SUMATIVA en la síntesis)',
     'auto'),
    ('Síntesis Didáctica', 15,
     'Diferencia resultados de aprendizaje con porcentajes según programa oficial',
     'auto'),

    # ── Sección 2: Planificación por unidades (3 ítems) ───────────────────
    ('Planificación por unidades', 1,
     'Presenta las unidades de aprendizaje diferenciadas correctamente '
     'para cada módulo, conforme a programa oficial',
     'auto'),
    ('Planificación por unidades', 2,
     'Presenta RA correctamente vinculados a los módulos, conforme a programa oficial',
     'auto'),
    ('Planificación por unidades', 3,
     'Las distintas sesiones se encuentran bien diferenciadas y con '
     'establecimiento de horas para cada momento',
     'auto'),

    # ── Sección 3: Estrategias metodológicas (11 ítems) ───────────────────
    ('Estrategias metodológicas', 1,
     'Se presentan los contenidos de actividades',
     'auto'),
    ('Estrategias metodológicas', 2,
     'Se presentan los contenidos de las evaluaciones',
     'auto'),
    ('Estrategias metodológicas', 3,
     'Las actividades son coherentes con los contenidos',
     'manual'),
    ('Estrategias metodológicas', 4,
     'Las evaluaciones son coherentes con los contenidos',
     'manual'),
    ('Estrategias metodológicas', 5,
     'Las actividades son consistentes con los Resultados de Aprendizaje',
     'manual'),
    ('Estrategias metodológicas', 6,
     'El trabajo independiente refleja autonomía del estudiante '
     '(tiempos asignados coherentes con la actividad)',
     'auto'),
    ('Estrategias metodológicas', 7,
     'Se establece momento de Preparación para cada actividad o evaluación',
     'auto'),
    ('Estrategias metodológicas', 8,
     'Se establece momento de Desarrollo para cada actividad o evaluación',
     'auto'),
    ('Estrategias metodológicas', 9,
     'Se establece momento de Trabajo Independiente para cada actividad o evaluación',
     'auto'),
    ('Estrategias metodológicas', 10,
     'Se presentan los recursos didácticos para cada actividad o evaluación',
     'auto'),
    ('Estrategias metodológicas', 11,
     'Los recursos didácticos son consistentes con los resultados de aprendizaje',
     'manual'),

    # ── Sección 4: Estrategias evaluativas (3 ítems) ──────────────────────
    ('Estrategias evaluativas', 1,
     'El tipo, procedimiento e instrumento de evaluación son coherentes '
     'y consistentes con el programa oficial',
     'auto'),
    ('Estrategias evaluativas', 2,
     'Existe consistencia entre los tipos de evaluaciones y los RA '
     'conforme al programa oficial (diagnóstica, formativa y sumativa por unidad)',
     'auto'),
    ('Estrategias evaluativas', 3,
     'Porcentajes de evaluaciones son consistentes conforme al programa oficial',
     'auto'),

    # ── Sección 5: Carga académica (9 ítems) ──────────────────────────────
    ('Carga académica', 1,
     'Distribución de horas presenciales conforme a programa oficial',
     'auto'),
    ('Carga académica', 2,
     'Total de horas presenciales conforme a programa oficial',
     'auto'),
    ('Carga académica', 3,
     'Distribución de horas virtuales conforme a programa oficial',
     'auto'),
    ('Carga académica', 4,
     'Total de horas virtuales conforme a programa oficial',
     'auto'),
    ('Carga académica', 5,
     'Diferenciación correcta entre horas sincrónicas y asincrónicas',
     'auto'),
    ('Carga académica', 6,
     'Asignación de horas de preparación consistentes con la carga del estudiante',
     'manual'),
    ('Carga académica', 7,
     'Distribución de horas TPE conforme a programa oficial',
     'auto'),
    ('Carga académica', 8,
     'Total horas TPE conforme a programa oficial',
     'auto'),
    ('Carga académica', 9,
     'Asignación de horas de TPE consistentes con la carga del estudiante',
     'manual'),
]


# ── Colores ───────────────────────────────────────────────────────────────
_AZUL_FUENTE    = '0070C0'
_FILL_FORMATIVA = PatternFill(fill_type='solid', fgColor='E8D5F5')  # lila pastel
_FILL_SUMATIVA  = PatternFill(fill_type='solid', fgColor='FFF2CC')  # amarillo pastel
_FILL_NONE      = PatternFill(fill_type=None)

def aplicar_azul(cell):
    """Aplica color de fuente azul (#0070C0) sin tocar el fondo de la celda."""
    f = cell.font
    cell.font = Font(name=f.name, size=f.size, bold=f.bold,
                     italic=f.italic, underline=f.underline, color=_AZUL_FUENTE)


def aplicar_colores_evaluacion(ws):
    """
    Colorea el fondo de filas según tipo de evaluación:
      Formativa / Formativo → lila pastel  (#E8D5F5)
      Sumativa              → amarillo pastel (#FFF2CC)
    Se aplica después de todas las correcciones para no interferir con
    el color de fuente azul de las celdas corregidas.
    Las celdas fusionadas (MergedCell) se omiten para evitar errores.
    """
    from openpyxl.cell.cell import MergedCell
    n_cols = ws.max_column or 15
    for row in ws.iter_rows(min_row=4, max_col=n_cols):
        r = row[0].row
        tipo = ws.cell(r, COL['TIPO']).value
        if not tipo:
            continue
        t = str(tipo).lower()
        if 'formativ' in t:
            fill = _FILL_FORMATIVA
        elif 'sumativ' in t:
            fill = _FILL_SUMATIVA
        else:
            continue
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.fill = fill


# ══════════════════════════════════════════════════════════════════════════
#  FUNCIONES AUXILIARES
# ══════════════════════════════════════════════════════════════════════════

def encontrar_archivo(carpeta, patron):
    """Busca el primer archivo que coincida con el patrón glob en una carpeta."""
    resultados = glob.glob(os.path.join(carpeta, patron))
    return resultados[0] if resultados else None


def encontrar_planificacion_enviada(carpeta_enviado):
    """
    En 'Enviado a DEL/', devuelve el .xlsx que NO es la escala de apreciación.
    Devuelve (path, escala_path) o (None, None).
    Si hay más de un xlsx que no es escala, usa el más reciente y advierte en stderr.
    """
    xlsx_files = glob.glob(os.path.join(carpeta_enviado, '*.xlsx'))
    planes = []
    escala_path = None
    for f in xlsx_files:
        nombre = os.path.basename(f).lower()
        if nombre.startswith('escala'):
            escala_path = f
        else:
            planes.append(f)
    if len(planes) > 1:
        planes.sort(key=os.path.getmtime, reverse=True)
        import sys
        print(
            f'  ⚠️  Múltiples planificaciones en "{carpeta_enviado}" — '
            f'se usa la más reciente: {os.path.basename(planes[0])}',
            file=sys.stderr
        )
    plan_path = planes[0] if planes else None
    return plan_path, escala_path


def leer_observaciones_escala(escala_path):
    """
    Lee la escala de apreciación y extrae las observaciones de mejora
    (columna F / 'Solicitud de Mejora') y comentarios de 2ª validación que tienen texto.
    Devuelve lista de (criterio, observacion).
    """
    if not escala_path:
        return []
    try:
        wb = openpyxl.load_workbook(escala_path, data_only=True)
        ws = wb.active
        observaciones = []
        for row in ws.iter_rows(min_row=4, values_only=True):
            criterio   = row[0]  if len(row) > 0  else None
            obs_1      = row[5]  if len(row) > 5  else None   # col F — 1ª validación
            obs_2      = row[9]  if len(row) > 9  else None   # col J — 2ª validación
            resultado  = row[2]  if len(row) > 2  else None   # col C — Si/X
            if not criterio:
                continue
            # Recoger observaciones reales (no encabezados de columna)
            obs_texto = []
            if obs_1 and str(obs_1).strip() and str(obs_1).strip().lower() not in (
                    'observaciones', 'solicitud de mejora/ observaciones',
                    'comentarios de postgrado post-1° validación'):
                obs_texto.append(str(obs_1).strip())
            if obs_2 and str(obs_2).strip() and str(obs_2).strip().lower() not in (
                    'comentarios 2° validación', 'ok'):
                obs_texto.append(f'[2ª val.] {str(obs_2).strip()}')
            if obs_texto:
                observaciones.append((str(criterio).strip(), ' | '.join(obs_texto)))
        return observaciones
    except Exception:
        return []


# ══════════════════════════════════════════════════════════════════════════
#  INSTANCIA 2 — APLICAR CORRECCIONES DE ESCALA DEL
# ══════════════════════════════════════════════════════════════════════════

_OBS_IGNORAR = {
    'si', 'parcialmente', 'no', 'observaciones',
    'solicitud de mejora/ observaciones',
    'comentarios de postgrado post-1° validación',
    'comentarios de postgrado post-1Â° validaciÃ³n',
    'comentarios 2° validación', 'comentarios 2Â° validaciÃ³n',
    'ok', 'identificación', 'resultados', 'validación inicial',
    'sanción validador(a) inicial',
}

_SECCIONES_ESCALA = (
    'síntesis didáctica', 'planificación por unidades',
    'estrategias metodológicas', 'estrategias evaluativas',
    'carga académica',
)


def leer_escala_completa(escala_path):
    """
    Lee la escala de apreciación completa.
    Devuelve lista de dicts:
      { seccion, criterio, estado ('Si'|'Parcialmente'|'No'), obs_texto }
    Solo incluye filas con marca de estado (X/x en col C, D o E).
    """
    if not escala_path:
        return []
    try:
        wb_e = openpyxl.load_workbook(escala_path, data_only=True)
        ws_e = wb_e.active
    except Exception:
        return []

    criterios = []
    seccion_actual = None

    for row in ws_e.iter_rows(min_row=4, values_only=True):
        criterio = row[0] if len(row) > 0 else None
        si_val   = row[2] if len(row) > 2 else None   # col C — Sí
        parc_val = row[3] if len(row) > 3 else None   # col D — Parcialmente
        no_val   = row[4] if len(row) > 4 else None   # col E — No
        obs_f    = row[5] if len(row) > 5 else None   # col F — Solicitud de Mejora
        obs_i    = row[8] if len(row) > 8 else None   # col I — Post-1ª validación
        obs_j    = row[9] if len(row) > 9 else None   # col J — 2ª validación

        if not criterio:
            continue
        c_lower = str(criterio).lower().strip()

        # Detectar encabezado de sección
        for sec in _SECCIONES_ESCALA:
            if c_lower.startswith(sec):
                seccion_actual = str(criterio).strip()
                break

        # Saltar filas sin marca de estado
        si_ok   = bool(si_val   and str(si_val).strip().upper()   == 'X')
        parc_ok = bool(parc_val and str(parc_val).strip().upper() == 'X')
        no_ok   = bool(no_val   and str(no_val).strip().upper()   == 'X')
        if not (si_ok or parc_ok or no_ok):
            continue

        if c_lower in _OBS_IGNORAR:
            continue

        estado = 'No' if no_ok else ('Parcialmente' if parc_ok else 'Si')

        # Construir texto de observación consolidado
        partes = []
        for o in (obs_f, obs_i, obs_j):
            if o:
                t = str(o).strip()
                if t and t.lower() not in _OBS_IGNORAR:
                    partes.append(t)
        obs_texto = ' | '.join(partes) if partes else None

        criterios.append({
            'seccion':   seccion_actual,
            'criterio':  str(criterio).strip(),
            'estado':    estado,
            'obs_texto': obs_texto,
        })

    return criterios


def _col_para_criterio(criterio_texto):
    """
    Mapea el texto de un criterio a la(s) columna(s) relevantes de la planificación.
    Devuelve lista de claves de COL (pueden ser varias).
    """
    t = str(criterio_texto).lower()
    cols = []
    if 'instrumento'                               in t: cols.append('INSTR')
    if 'procedimiento'                             in t: cols.append('PROC')
    if 'tipo de evaluación' in t or 'tipos de ev' in t: cols.append('TIPO')
    if 'recurso'                                   in t: cols.append('RECURSOS')
    if 'resultado' in t and 'aprendizaje'          in t: cols.append('RA')
    if 'contenido'                                 in t: cols.append('CONTENIDOS')
    if 'medio de entrega'                          in t: cols.append('MEDIO')
    if 'porcentaje'                                in t: cols.append('PCT')
    if 'individual' in t or 'grupal'               in t: cols.append('INDIV')
    if ('redacción' in t or 'momentos' in t
            or ('actividad' in t and 'estrategia' not in t)):
                                                         cols.append('ACTIVIDAD')
    return cols


def _escribir_notas_i2(wb, fallidos):
    """Escribe/reemplaza la hoja NOTAS_CORRECCIONES_DEL con el resumen de I2."""
    HOJA = 'NOTAS_CORRECCIONES_DEL'
    if HOJA in wb.sheetnames:
        del wb[HOJA]
    ws_n = wb.create_sheet(HOJA)

    ws_n.cell(1, 1, 'OBSERVACIONES REVISORA DEL').font = Font(bold=True, size=12)
    ws_n.cell(2, 1, 'Sección').font   = Font(bold=True)
    ws_n.cell(2, 2, 'Criterio').font  = Font(bold=True)
    ws_n.cell(2, 3, 'Estado').font    = Font(bold=True)
    ws_n.cell(2, 4, 'Observación').font = Font(bold=True)

    fill_no   = PatternFill(fill_type='solid', fgColor='F8D7DA')
    fill_parc = PatternFill(fill_type='solid', fgColor='FFF3CD')

    for i, item in enumerate(fallidos, start=3):
        ws_n.cell(i, 1, item['seccion']  or '—')
        ws_n.cell(i, 2, item['criterio'])
        ws_n.cell(i, 3, item['estado'])
        ws_n.cell(i, 4, item['obs_texto'] or '—')
        fill = fill_no if item['estado'] == 'No' else fill_parc
        for c in range(1, 5):
            ws_n.cell(i, c).fill = fill

    ws_n.column_dimensions['B'].width = 60
    ws_n.column_dimensions['D'].width = 80


def aplicar_obs_escala_i2(wb, criterios, log):
    """
    Para cada criterio ❌ o Parcialmente con texto de observación:
    - Inyecta el texto en azul en las celdas vacías de la(s) columna(s) mapeadas.
    - Escribe resumen en hoja NOTAS_CORRECCIONES_DEL.
    Devuelve número de anotaciones aplicadas.
    """
    from openpyxl.cell.cell import MergedCell

    ws_plan = wb['Planificación por unidades'] if 'Planificación por unidades' in wb.sheetnames else None
    if not ws_plan:
        return 0

    fallidos = [c for c in criterios
                if c['estado'] in ('No', 'Parcialmente') and c['obs_texto']]

    if not fallidos:
        log.append('  [I2] Sin observaciones de la revisora con texto.')
        _escribir_notas_i2(wb, [])
        return 0

    log.append(f'\n  [Instancia 2 — Observaciones revisora DEL ({len(fallidos)} criterios)]')
    total_anotaciones = 0

    for item in fallidos:
        criterio = item['criterio']
        obs      = item['obs_texto']
        estado   = item['estado']
        cols     = _col_para_criterio(criterio)
        icono    = '❌' if estado == 'No' else '⚠️'

        log.append(f'  {icono} {criterio[:75]}')
        if obs:
            log.append(f'      → {obs[:140]}')

        if not cols:
            log.append('      [sin mapeo automático — ver hoja NOTAS_CORRECCIONES_DEL]')
            continue

        obs_corta = obs[:100] + ('…' if len(obs) > 100 else '')

        for col_key in cols:
            col_num = COL.get(col_key)
            if not col_num:
                continue

            for row in ws_plan.iter_rows(min_row=4):
                r = row[0].row
                # Solo filas con actividad o evaluación (no filas vacías)
                tiene_contenido = (ws_plan.cell(r, COL['ACTIVIDAD']).value
                                   or ws_plan.cell(r, COL['TIPO']).value)
                if not tiene_contenido:
                    continue

                target = ws_plan.cell(r, col_num)
                if isinstance(target, MergedCell):
                    continue

                if not target.value:
                    target.value = f'[DEL: {obs_corta}]'
                    aplicar_azul(target)
                    total_anotaciones += 1

    _escribir_notas_i2(wb, fallidos)
    log.append(f'  Anotaciones inyectadas: {total_anotaciones}  '
               f'| Resumen completo en hoja NOTAS_CORRECCIONES_DEL')
    return total_anotaciones


def procesar_instancia2(plan_bytes, escala_bytes,
                        plan_nombre='planificacion.xlsx',
                        programa=None, es_as=False,
                        instancia_num=2):
    """
    Flujo Instancia 2 / 3: aplica correcciones automáticas + observaciones de la
    revisora DEL (escala completada) sobre la planificación del docente.

    Parámetros
    ----------
    plan_bytes    : bytes del .xlsx del docente
    escala_bytes  : bytes de la escala completada por la revisora
    plan_nombre   : nombre original del archivo (para el nombre de salida)
    programa      : dict extraído por extraer_programa_pdf (opcional)
    es_as         : bool — activa verificación hitos A+Se
    instancia_num : 2 ó 3 — controla encabezado y nombre del archivo de salida

    Devuelve
    --------
    (log: list[str], ok: bool, output_bytes: bytes, output_nombre: str)
    """
    from io import BytesIO

    log = [
        f'\n{"═"*70}',
        f'  INSTANCIA {instancia_num} — Aplicar correcciones de escala DEL',
        f'{"═"*70}',
        f'  Plan   : {plan_nombre}',
    ]

    # ── Leer escala ────────────────────────────────────────────────────────
    try:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_e:
            tmp_e.write(escala_bytes)
            tmp_escala = tmp_e.name
        criterios = leer_escala_completa(tmp_escala)
        os.unlink(tmp_escala)
    except Exception as exc:
        log.append(f'  ❌ Error leyendo escala: {exc}')
        return log, False, b'', ''

    n_total   = len(criterios)
    n_fallidos = sum(1 for c in criterios if c['estado'] in ('No', 'Parcialmente'))
    n_ok       = n_total - n_fallidos
    log.append(f'  Escala  : {n_total} criterios leídos | '
               f'{n_ok} ✅ | {n_fallidos} con observaciones')

    # ── Cargar planificación ───────────────────────────────────────────────
    try:
        wb = openpyxl.load_workbook(BytesIO(plan_bytes))
    except Exception as exc:
        log.append(f'  ❌ Error cargando planificación: {exc}')
        return log, False, b'', ''

    hojas = wb.sheetnames
    log.append(f'  Hojas   : {hojas}')

    # ── Correcciones automáticas (misma lógica I1) ─────────────────────────
    total_sint = 0
    if 'Síntesis didáctica' in hojas:
        log.append('\n  [Síntesis didáctica — correcciones automáticas]')
        total_sint = corregir_sintesis(wb['Síntesis didáctica'], log)
        if total_sint == 0:
            log.append('    Sin cambios necesarios.')

    total_plan = 0
    if 'Planificación por unidades' in hojas:
        log.append('\n  [Planificación por unidades — correcciones automáticas]')
        a, mj, mm, b = corregir_planificacion(wb['Planificación por unidades'], log)
        total_leng   = corregir_lenguaje_actividades(
            wb['Planificación por unidades'], log)
        total_plan   = a + mj + mm + b + total_leng
        if total_plan == 0:
            log.append('    Sin cambios necesarios.')

    # ── Observaciones de la revisora ───────────────────────────────────────
    n_obs = aplicar_obs_escala_i2(wb, criterios, log)

    # ── Colores de fila (Formativa/Sumativa) ──────────────────────────────
    if 'Planificación por unidades' in hojas:
        aplicar_colores_evaluacion(wb['Planificación por unidades'])

    # ── Verificación 41 criterios ──────────────────────────────────────────
    resultados_escala = verificar_escala(wb)
    log.extend(formatear_verificacion(resultados_escala, 'archivo'))

    # ── Verificación vs programa ───────────────────────────────────────────
    verificar_contra_programa(wb, programa, log)

    # ── Verificación A+Se ──────────────────────────────────────────────────
    if es_as and 'Planificación por unidades' in hojas:
        verificar_as(wb['Planificación por unidades'], log)

    # ── Guardar ────────────────────────────────────────────────────────────
    total = total_sint + total_plan + n_obs
    log.append(f'\n  RESUMEN I2:')
    log.append(f'    Auto-correcciones I1     : {total_sint + total_plan}')
    log.append(f'    Anotaciones revisora DEL : {n_obs}')
    log.append(f'    TOTAL                    : {total}')
    log.append(f'    🔵 Texto azul = corrección automática o anotación revisora')
    log.append(f'    📋 Hoja NOTAS_CORRECCIONES_DEL = resumen completo de observaciones')

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    output_bytes = buf.read()

    output_nombre = nombre_instancia(plan_nombre, instancia_num)

    return log, True, output_bytes, output_nombre


def verificar_escala(wb):
    """
    Verifica automáticamente los criterios marcados como 'auto' en ESCALA_ESTANDAR
    contra el contenido del workbook.
    Devuelve lista de (seccion, n, criterio, estado, detalle)
    donde estado es: '✅', '❌', '⚠️ manual'
    """
    resultados = []

    ws_sint = wb['Síntesis didáctica'] if 'Síntesis didáctica' in wb.sheetnames else None
    ws_plan = wb['Planificación por unidades'] if 'Planificación por unidades' in wb.sheetnames else None

    # Leer síntesis en dict para acceso fácil
    sint = {}
    if ws_sint:
        # Leer valores de celdas clave por posición
        def sv(r, c): return ws_sint.cell(r, c).value

        sint = {
            'escuela':      sv(4, 1),
            'programa':     sv(4, 2),
            'asignatura':   sv(4, 4),
            'codigo':       sv(7, 1),
            'version':      sv(7, 2),
            'creditos':     sv(7, 3),
            'area':         sv(11, 1),
            'requisitos':   sv(11, 2),
            'hrs_pres':     sv(9, 4),
            'hrs_sinc':     sv(9, 5),
            'hrs_asinc':    sv(9, 6),
            'hrs_tpe':      sv(9, 7),
        }
        # Competencias (filas 13-15 col B)
        sint['competencias'] = [ws_sint.cell(r, 2).value for r in range(13, 16)
                                if ws_sint.cell(r, 2).value]
        # Unidades (filas 18-21 col B y G)
        sint['unidades'] = [(ws_sint.cell(r, 2).value, ws_sint.cell(r, 7).value)
                            for r in range(18, 22)
                            if ws_sint.cell(r, 2).value]
        # RA + procedimientos (filas 25-29 col A y D)
        sint['ra_procs'] = [(ws_sint.cell(r, 1).value,
                             ws_sint.cell(r, 4).value,
                             ws_sint.cell(r, 7).value)
                            for r in range(25, 32)
                            if ws_sint.cell(r, 1).value]

    def ok(val, nombre):
        if val:
            return ('✅', f'{nombre}: presente')
        return ('❌', f'{nombre}: vacío o ausente')

    def check(seccion, n, criterio, modo, estado, detalle):
        resultados.append((seccion, n, criterio, estado, detalle))

    # ── Sección 1: Síntesis Didáctica ─────────────────────────────────────
    if ws_sint:
        e, d = ok(sint.get('escuela'), 'Escuela')
        check('Síntesis Didáctica', 1, ESCALA_ESTANDAR[0][2], 'auto', e, d)

        e, d = ok(sint.get('programa'), 'Programa')
        check('Síntesis Didáctica', 2, ESCALA_ESTANDAR[1][2], 'auto', e, d)

        e, d = ok(sint.get('asignatura'), 'Nombre asignatura')
        check('Síntesis Didáctica', 3, ESCALA_ESTANDAR[2][2], 'auto', e, d)

        e, d = ok(sint.get('codigo'), 'Código')
        check('Síntesis Didáctica', 4, ESCALA_ESTANDAR[3][2], 'auto', e, d)

        e, d = ok(sint.get('version'), 'Versión año')
        check('Síntesis Didáctica', 5, ESCALA_ESTANDAR[4][2], 'auto', e, d)

        e, d = ok(sint.get('creditos'), 'N° créditos')
        check('Síntesis Didáctica', 6, ESCALA_ESTANDAR[5][2], 'auto', e, d)

        e, d = ok(sint.get('area'), 'Área OCDE')
        check('Síntesis Didáctica', 7, ESCALA_ESTANDAR[6][2], 'auto', e, d)

        e, d = ok(sint.get('requisitos'), 'Requisitos')
        check('Síntesis Didáctica', 8, ESCALA_ESTANDAR[7][2], 'auto', e, d)

        horas = [sint.get('hrs_pres'), sint.get('hrs_sinc'),
                 sint.get('hrs_asinc'), sint.get('hrs_tpe')]
        if all(h is not None for h in horas):
            e, d = '✅', f'Horas: pres={horas[0]} sinc={horas[1]} asinc={horas[2]} TPE={horas[3]}'
        else:
            e, d = '❌', f'Faltan horas: {horas}'
        check('Síntesis Didáctica', 9, ESCALA_ESTANDAR[8][2], 'auto', e, d)

        n_comp = len(sint.get('competencias', []))
        e = '✅' if n_comp > 0 else '❌'
        check('Síntesis Didáctica', 10, ESCALA_ESTANDAR[9][2], 'auto',
              e, f'{n_comp} competencia(s) declarada(s)')

        unidades = sint.get('unidades', [])
        n_u = len(unidades)
        todas_con_horas = all(h for _, h in unidades)
        e = '✅' if n_u > 0 and todas_con_horas else ('⚠️' if n_u > 0 else '❌')
        check('Síntesis Didáctica', 11, ESCALA_ESTANDAR[10][2], 'auto',
              e, f'{n_u} unidad(es); todas con horas={todas_con_horas}')

        e = '✅' if n_u > 0 else '❌'
        check('Síntesis Didáctica', 12, ESCALA_ESTANDAR[11][2], 'auto',
              e, f'{n_u} título(s) de unidad presentes')

        ra_procs = sint.get('ra_procs', [])
        n_ra = len(ra_procs)
        e = '✅' if n_ra > 0 else '❌'
        check('Síntesis Didáctica', 13, ESCALA_ESTANDAR[12][2], 'auto',
              e, f'{n_ra} RA declarado(s)')

        # Criterio 14: procedimientos solo SUMATIVA
        procs_malos = [(r, p) for r, p, _ in ra_procs
                       if p and ('DIAGNÓSTICA' in str(p) or 'FORMATIVA' in str(p))]
        e = '✅' if not procs_malos else '❌'
        d = 'Solo SUMATIVA en procedimientos' if not procs_malos \
            else f'Contiene DIAGNÓSTICA/FORMATIVA en {len(procs_malos)} celda(s)'
        check('Síntesis Didáctica', 14, ESCALA_ESTANDAR[13][2], 'auto', e, d)

        # Criterio 15: porcentajes
        pcts = [p for _, _, p in ra_procs if isinstance(p, (int, float))]
        total_pct = sum(pcts)
        e = '✅' if abs(total_pct - 1.0) < 0.01 or abs(total_pct - 100) < 1 else '❌'
        check('Síntesis Didáctica', 15, ESCALA_ESTANDAR[14][2], 'auto',
              e, f'Suma porcentajes = {total_pct}')
    else:
        for i in range(15):
            check('Síntesis Didáctica', i+1, ESCALA_ESTANDAR[i][2], 'auto',
                  '❌', 'Hoja no encontrada')

    # ── Sección 2: Planificación por unidades ─────────────────────────────
    if ws_plan:
        unidades_plan = set()
        ra_plan = set()
        momentos = set()
        filas_eval = 0
        filas_sin_horas = 0
        sesiones_con_prep = sesiones_con_des = sesiones_con_ti = 0
        semana_actual = None

        for row in ws_plan.iter_rows(min_row=4, values_only=True):
            unidad  = row[COL['UNIDAD']  - 1]
            ra      = row[COL['RA']      - 1]
            semana  = row[COL['SEMANA']  - 1]
            momento = row[COL['MOMENTO'] - 1]
            tipo    = row[COL['TIPO']    - 1]
            hrs_p   = row[COL['MODALIDAD'] - 1]  # col E modalidad

            if unidad: unidades_plan.add(str(unidad))
            if ra:     ra_plan.add(str(ra)[:40])
            if semana and semana != 'Total carga académica del estudiante':
                semana_actual = semana
            if momento:
                m = str(momento).lower()
                if 'preparación' in m:   sesiones_con_prep += 1
                if 'desarrollo' in m:    sesiones_con_des  += 1
                if 'independiente' in m: sesiones_con_ti   += 1
            if tipo:
                filas_eval += 1

        n_u = len(unidades_plan) - 1 if '' in unidades_plan else len(unidades_plan)
        e = '✅' if n_u > 0 else '❌'
        check('Planificación por unidades', 1, ESCALA_ESTANDAR[15][2], 'auto',
              e, f'{n_u} unidad(es) identificada(s)')

        e = '✅' if len(ra_plan) > 0 else '❌'
        check('Planificación por unidades', 2, ESCALA_ESTANDAR[16][2], 'auto',
              e, f'{len(ra_plan)} RA vinculado(s) a unidades')

        todos = sesiones_con_prep > 0 and sesiones_con_des > 0 and sesiones_con_ti > 0
        e = '✅' if todos else '❌'
        check('Planificación por unidades', 3, ESCALA_ESTANDAR[17][2], 'auto',
              e, f'Prep={sesiones_con_prep} Des={sesiones_con_des} TI={sesiones_con_ti}')
    else:
        for i in range(3):
            check('Planificación por unidades', i+1, ESCALA_ESTANDAR[15+i][2], 'auto',
                  '❌', 'Hoja no encontrada')

    # ── Sección 3: Estrategias metodológicas ──────────────────────────────
    idx = 18  # índice en ESCALA_ESTANDAR
    if ws_plan:
        filas_actividad = filas_recursos = filas_contenido_eval = 0
        filas_eval_sin_contenido = 0
        prep_rows = des_rows = ti_rows = 0
        filas_eval_total = 0

        for row in ws_plan.iter_rows(min_row=4, values_only=True):
            momento   = str(row[COL['MOMENTO']   - 1] or '').lower()
            actividad = row[COL['ACTIVIDAD'] - 1]
            recursos  = row[COL['RECURSOS']  - 1]
            contenido = row[COL['CONTENIDOS']- 1]
            tipo      = row[COL['TIPO']      - 1]

            if actividad: filas_actividad += 1
            if recursos:  filas_recursos  += 1
            if tipo:
                filas_eval_total += 1
                if contenido: filas_contenido_eval += 1
            if 'preparación'   in momento: prep_rows += 1
            if 'desarrollo'    in momento: des_rows  += 1
            if 'independiente' in momento: ti_rows   += 1

        e = '✅' if filas_actividad > 0 else '❌'
        check('Estrategias metodológicas', 1, ESCALA_ESTANDAR[idx][2], 'auto',
              e, f'{filas_actividad} filas con instrucciones de actividad')
        idx += 1

        # Criterio 2: contenidos de evaluaciones
        # Las filas de evaluación deben tener instrucciones en col G (Actividad)
        filas_eval_con_instruc = 0
        for row in ws_plan.iter_rows(min_row=4, values_only=True):
            tipo_r     = row[COL['TIPO']     - 1]
            actividad_r= row[COL['ACTIVIDAD']- 1]
            if tipo_r and actividad_r:
                filas_eval_con_instruc += 1
        if filas_eval_total > 0:
            e = '✅' if filas_eval_con_instruc >= filas_eval_total * 0.8 else '⚠️'
            d = f'{filas_eval_con_instruc}/{filas_eval_total} filas de evaluación con instrucciones'
        else:
            e, d = '⚠️', 'No se detectaron filas de evaluación'
        check('Estrategias metodológicas', 2, ESCALA_ESTANDAR[idx][2], 'auto', e, d)
        idx += 1

        # Criterios 3, 4, 5 — manual
        for i in range(3):
            check('Estrategias metodológicas', 3+i, ESCALA_ESTANDAR[idx][2], 'manual',
                  '⚠️ manual', 'Requiere revisión humana')
            idx += 1

        # Criterio 6: TPE autonomía — verificable estructuralmente
        e = '✅' if ti_rows > 0 else '❌'
        check('Estrategias metodológicas', 6, ESCALA_ESTANDAR[idx][2], 'auto',
              e, f'{ti_rows} filas con Trabajo Independiente')
        idx += 1

        # Criterios 7, 8, 9: momentos
        for label, count, n in [('Preparación', prep_rows, 7),
                                  ('Desarrollo',  des_rows,  8),
                                  ('T. Independiente', ti_rows, 9)]:
            e = '✅' if count > 0 else '❌'
            check('Estrategias metodológicas', n, ESCALA_ESTANDAR[idx][2], 'auto',
                  e, f'{count} filas con momento {label}')
            idx += 1

        # Criterio 10: recursos didácticos
        e = '✅' if filas_recursos > 0 else '❌'
        check('Estrategias metodológicas', 10, ESCALA_ESTANDAR[idx][2], 'auto',
              e, f'{filas_recursos} filas con recursos didácticos')
        idx += 1

        # Criterio 11 — manual
        check('Estrategias metodológicas', 11, ESCALA_ESTANDAR[idx][2], 'manual',
              '⚠️ manual', 'Requiere revisión humana')
        idx += 1

    # ── Sección 4: Estrategias evaluativas ────────────────────────────────
    if ws_plan:
        tipos_por_unidad = {}
        procs_invalidos  = []
        instrs_invalidos = []
        pct_total = 0

        unidad_actual = None
        for row in ws_plan.iter_rows(min_row=4, values_only=True):
            unidad = row[COL['UNIDAD'] - 1]
            tipo   = row[COL['TIPO']   - 1]
            proc   = row[COL['PROC']   - 1]
            instr  = row[COL['INSTR']  - 1]
            pct    = row[COL['PCT']    - 1]

            if unidad: unidad_actual = str(unidad)
            if not tipo: continue

            t = str(tipo).lower()
            if unidad_actual not in tipos_por_unidad:
                tipos_por_unidad[unidad_actual] = set()
            tipos_por_unidad[unidad_actual].add(t)

            if proc in PROC_MAP:
                procs_invalidos.append(f'F? proc="{proc}"')
            if instr in ('Pauta', 'Rúbrica'):
                instrs_invalidos.append(f'F? instr="{instr}"')
            if isinstance(pct, (int, float)) and pct > 0:
                pct_total += pct

        # Criterio 1: tipo/proc/instr coherentes
        errores = len(procs_invalidos) + len(instrs_invalidos)
        e = '✅' if errores == 0 else '❌'
        d = 'Sin errores de procedimiento/instrumento' if errores == 0 \
            else f'{errores} error(es): {procs_invalidos + instrs_invalidos}'
        check('Estrategias evaluativas', 1, ESCALA_ESTANDAR[idx][2], 'auto', e, d)
        idx += 1

        # Criterio 2: todos los tipos presentes por unidad
        # Excluir filas de LGA / Examen final (no son unidades temáticas)
        EXCLUIR_UNIDADES = {'evaluación final', 'lga', 'examen'}
        faltantes = []
        for u, tipos in tipos_por_unidad.items():
            if any(ex in str(u).lower() for ex in EXCLUIR_UNIDADES):
                continue
            for t_req in ('diagnóstica', 'formativa', 'sumativa'):
                if not any(t_req in t for t in tipos):
                    faltantes.append(f'{u[:35]}: falta {t_req}')
        e = '✅' if not faltantes else '❌'
        d = 'Diagnóstica/Formativa/Sumativa presentes por unidad' if not faltantes \
            else '; '.join(faltantes)
        check('Estrategias evaluativas', 2, ESCALA_ESTANDAR[idx][2], 'auto', e, d)
        idx += 1

        # Criterio 3: porcentajes sumativos de unidades suman 100%
        # Excluir el examen final (que es adicional al 100%)
        pct_unidades = 0
        _unidad_actual_pct = None
        for row in ws_plan.iter_rows(min_row=4, values_only=True):
            tipo_r   = row[COL['TIPO']   - 1]
            pct_r    = row[COL['PCT']    - 1]
            unidad_r = row[COL['UNIDAD'] - 1]
            if unidad_r:
                _unidad_actual_pct = str(unidad_r)
            es_excluida = any(ex in str(_unidad_actual_pct or '').lower()
                              for ex in EXCLUIR_UNIDADES)
            if (tipo_r and 'sumativa' in str(tipo_r).lower()
                    and isinstance(pct_r, (int, float)) and pct_r > 0
                    and not es_excluida):
                pct_unidades += pct_r
        e = '✅' if abs(pct_unidades - 100) < 1 or abs(pct_unidades - 1.0) < 0.01 else '⚠️'
        check('Estrategias evaluativas', 3, ESCALA_ESTANDAR[idx][2], 'auto',
              e, f'Suma % sumativos (sin examen) = {pct_unidades}')
        idx += 1

    # ── Sección 5: Carga académica ─────────────────────────────────────────
    if ws_sint:
        def hv(key): return sint.get(key)

        # Criterios 1-5: presencia de valores de horas
        for n, key, label in [
            (1, 'hrs_pres',  'Horas presenciales'),
            (2, 'hrs_pres',  'Total horas presenciales'),
            (3, 'hrs_sinc',  'Horas virtuales sincrónicas'),
            (4, 'hrs_asinc', 'Horas virtuales asincrónicas'),
            (5, None,        'Diferenciación sincrónicas/asincrónicas'),
        ]:
            if key:
                v = hv(key)
                e = '✅' if v is not None else '❌'
                d = f'{label}: {v}'
            else:
                s, a = hv('hrs_sinc'), hv('hrs_asinc')
                e = '✅' if (s is not None and a is not None) else '❌'
                d = f'Sincrónicas={s} Asincrónicas={a}'
            check('Carga académica', n, ESCALA_ESTANDAR[idx][2], 'auto', e, d)
            idx += 1

        # Criterio 6 — manual
        check('Carga académica', 6, ESCALA_ESTANDAR[idx][2], 'manual',
              '⚠️ manual', 'Requiere revisión humana')
        idx += 1

        # Criterios 7-8: horas TPE
        tpe = hv('hrs_tpe')
        for n, label in [(7, 'Distribución horas TPE'), (8, 'Total horas TPE')]:
            e = '✅' if tpe is not None else '❌'
            check('Carga académica', n, ESCALA_ESTANDAR[idx][2], 'auto',
                  e, f'{label}: {tpe}')
            idx += 1

        # Criterio 9 — manual
        check('Carga académica', 9, ESCALA_ESTANDAR[idx][2], 'manual',
              '⚠️ manual', 'Requiere revisión humana')

    return resultados


def formatear_verificacion(resultados, fuente_escala):
    """
    Formatea los resultados de verificación para el reporte.
    fuente_escala: 'archivo' | 'estandar'
    """
    lineas = []
    lineas.append('\n  [Verificación — Escala Estándar UST (41 criterios — meta: 100%)]')
    if fuente_escala == 'estandar':
        lineas.append('  Fuente criterios: estándar institucional integrado (sin archivo de escala)')
    else:
        lineas.append('  Fuente criterios: estándar institucional integrado + archivo de escala disponible')

    seccion_actual = None
    total = len(resultados)
    aprobados = sum(1 for _, _, _, e, _ in resultados if e == '✅')
    manuales  = sum(1 for _, _, _, e, _ in resultados if 'manual' in str(e))
    fallidos  = total - aprobados - manuales

    for seccion, n, criterio, estado, detalle in resultados:
        if seccion != seccion_actual:
            lineas.append(f'\n    ── {seccion} ──')
            seccion_actual = seccion
        texto_criterio = criterio[:65] + '…' if len(criterio) > 65 else criterio
        lineas.append(f'    {estado} {n:02d}. {texto_criterio}')
        if estado != '✅':
            lineas.append(f'         → {detalle}')

    lineas.append(f'\n    Resultado: {aprobados}✅  {fallidos}❌  {manuales}⚠️ manual  '
                  f'(de {total} criterios)')
    return lineas


def _base_limpia(nombre):
    """
    Elimina todos los sufijos de instancia del nombre de archivo (iterativo).
    Quita cualquier combinación de: _I1, _I2, _I3, _REVISADO, _FINAL.
    """
    base, ext = os.path.splitext(nombre)
    patron = re.compile(r'(_I[123]|_REVISADO|_FINAL)+$', re.IGNORECASE)
    prev = None
    while prev != base:
        prev = base
        base = patron.sub('', base)
    return base, ext


def nombre_instancia(nombre_original, n):
    """
    Genera nombre limpio con sufijo _I1 / _I2 / _I3.
    Elimina cualquier sufijo de instancia previo antes de añadir el nuevo.
    Ejemplo: 'plan_I1_REVISADO_I2.xlsx' → 'plan_I3.xlsx'
    """
    base, ext = _base_limpia(nombre_original)
    return f'{base}_I{n}{ext}'


def nombre_revisado(nombre_original):
    """Alias de compatibilidad — genera nombre _I1."""
    return nombre_instancia(nombre_original, 1)


def solo_sumativa(texto):
    """Extrae únicamente el bloque SUMATIVA de un texto de procedimientos."""
    match = re.search(r'(SUMATIVA\s*\([^)]+\):.*)', texto, re.DOTALL)
    return match.group(1).strip() if match else texto


def tiene_bloques_extra(texto):
    """Devuelve True si el texto contiene DIAGNÓSTICA o FORMATIVA."""
    return 'DIAGNÓSTICA' in texto or 'FORMATIVA' in texto


def inferir_medio(modalidad, momento, tipo):
    """Infiere el Medio de entrega según modalidad, momento y tipo de evaluación."""
    if not tipo:
        return None
    m = str(momento or '').lower()
    mod = str(modalidad or '').lower()
    t = str(tipo or '').lower()
    if 'diagnós' in t and 'preparación' in m:
        return 'Cuestionario'
    if 'desarrollo' in m and ('sincrónico' in mod or 'presencial' in mod):
        return 'No aplica'
    if 'independiente' in m or 'no presencial' in mod:
        return 'Buzón de tareas'
    return 'No aplica'


def inferir_indiv_grupal(procedimiento, unidad_es_grupal):
    """Infiere Individual/Grupal según procedimiento y contexto de la unidad."""
    proc = str(procedimiento or '').lower()
    if 'equipo' in proc or 'grupo' in proc or 'colaborativo' in proc:
        return 'Grupal'
    if unidad_es_grupal:
        return 'Grupal'
    return 'Individual'


def detectar_unidades_grupales(ws):
    """
    Detecta qué unidades (col B) tienen al menos una sumativa con M='Grupal'.
    Devuelve un set de valores de col B.
    """
    grupales = set()
    for row in ws.iter_rows(min_row=4, values_only=False):
        tipo  = row[COL['TIPO']  - 1].value
        indiv = row[COL['INDIV'] - 1].value
        unidad = row[COL['UNIDAD'] - 1].value
        if tipo and 'sumativa' in str(tipo).lower() and indiv == 'Grupal' and unidad:
            grupales.add(unidad)
    return grupales


# ══════════════════════════════════════════════════════════════════════════
#  CORRECCIÓN SÍNTESIS DIDÁCTICA
# ══════════════════════════════════════════════════════════════════════════

def corregir_sintesis(ws, log):
    """Elimina DIAGNÓSTICA y FORMATIVA de procedimientos de evaluación."""
    cambios = 0
    for row in ws.iter_rows(min_row=1, values_only=False):
        for cell in row:
            if cell.column == 4 and cell.value and isinstance(cell.value, str):
                if tiene_bloques_extra(cell.value):
                    nuevo = solo_sumativa(cell.value)
                    log.append(f'    [Síntesis F{cell.row}] Eliminados bloques DIAGNÓSTICA/FORMATIVA')
                    cell.value = nuevo
                    aplicar_azul(cell)
                    cambios += 1
    return cambios


# ══════════════════════════════════════════════════════════════════════════
#  CORRECCIÓN PLANIFICACIÓN POR UNIDADES
# ══════════════════════════════════════════════════════════════════════════

def corregir_planificacion(ws, log):
    """Aplica todas las correcciones a la hoja Planificación por unidades."""
    cambios_alta = 0
    cambios_media_j = 0
    cambios_media_m = 0
    cambios_baja = 0

    # Pre-scan: detectar unidades grupales (para inferir M)
    unidades_grupales = detectar_unidades_grupales(ws)

    # Rastrear unidad actual (col B puede estar vacía por merge)
    unidad_actual = None

    for row in ws.iter_rows(min_row=4, values_only=False):
        r = row[0].row

        # Leer valores clave
        unidad_cell   = ws.cell(r, COL['UNIDAD'])
        modalidad     = ws.cell(r, COL['MODALIDAD']).value
        momento       = ws.cell(r, COL['MOMENTO']).value
        tipo_cell     = ws.cell(r, COL['TIPO'])
        proc_cell     = ws.cell(r, COL['PROC'])
        indiv_cell    = ws.cell(r, COL['INDIV'])
        instr_cell    = ws.cell(r, COL['INSTR'])
        pct_cell      = ws.cell(r, COL['PCT'])
        medio_cell    = ws.cell(r, COL['MEDIO'])

        if unidad_cell.value:
            unidad_actual = unidad_cell.value

        tipo  = tipo_cell.value
        proc  = proc_cell.value
        instr = instr_cell.value

        tiene_eval = bool(tipo)

        # ── ALTA: datos huérfanos (proc/instr sin tipo) ────────────────────
        if not tipo and (proc or instr):
            actividad = ws.cell(r, COL['ACTIVIDAD']).value
            if not actividad:
                # Fila sin actividad ni tipo: limpiar proc/instr huérfanos
                proc_cell.value = None
                instr_cell.value = None
                log.append(f'    [Plan F{r}] Datos huérfanos eliminados (L y N sin tipo evaluación)')
                cambios_alta += 1
            # En ambos casos, si no hay tipo no hay correcciones evaluativas que aplicar
            continue

        if not tiene_eval:
            continue

        # ── ALTA: normalizar Tipo ──────────────────────────────────────────
        if tipo in TIPO_MAP:
            nuevo_tipo = TIPO_MAP[tipo]
            log.append(f'    [Plan F{r}] Tipo: "{tipo}" → "{nuevo_tipo}"')
            tipo_cell.value = nuevo_tipo
            aplicar_azul(tipo_cell)
            tipo = nuevo_tipo
            cambios_alta += 1

        # ── ALTA: normalizar Procedimiento ────────────────────────────────
        if proc in PROC_MAP:
            nuevo_proc = PROC_MAP[proc]
            log.append(f'    [Plan F{r}] Procedimiento: "{proc}" → "{nuevo_proc}"')
            proc_cell.value = nuevo_proc
            aplicar_azul(proc_cell)
            proc = nuevo_proc
            cambios_alta += 1

        # ── ALTA: normalizar Instrumento ──────────────────────────────────
        if instr == 'Pauta':
            log.append(f'    [Plan F{r}] Instrumento: "Pauta" → "Pauta de observación"')
            instr_cell.value = 'Pauta de observación'
            aplicar_azul(instr_cell)
            instr = 'Pauta de observación'
            cambios_alta += 1
        elif instr == 'Rúbrica' and proc in PROC_INSTR_RÚBRICA_CONTEXTOS:
            log.append(f'    [Plan F{r}] Instrumento: "Rúbrica" → "Pauta de observación" (proc={proc})')
            instr_cell.value = 'Pauta de observación'
            aplicar_azul(instr_cell)
            instr = 'Pauta de observación'
            cambios_alta += 1

        # ── MEDIA: Medio de entrega (J) ────────────────────────────────────
        if medio_cell.value is None:
            nuevo_medio = inferir_medio(modalidad, momento, tipo)
            if nuevo_medio:
                medio_cell.value = nuevo_medio
                aplicar_azul(medio_cell)
                log.append(f'    [Plan F{r}] Medio de entrega: None → "{nuevo_medio}"')
                cambios_media_j += 1

        # ── MEDIA: Individual/Grupal (M) ───────────────────────────────────
        if indiv_cell.value is None:
            es_grupal = unidad_actual in unidades_grupales
            nuevo_indiv = inferir_indiv_grupal(proc, es_grupal)
            indiv_cell.value = nuevo_indiv
            aplicar_azul(indiv_cell)
            log.append(f'    [Plan F{r}] Individual/Grupal: None → "{nuevo_indiv}"')
            cambios_media_m += 1

        # ── BAJA: % = None → 0 ────────────────────────────────────────────
        if pct_cell.value is None:
            pct_cell.value = 0
            aplicar_azul(pct_cell)
            log.append(f'    [Plan F{r}] %: None → 0')
            cambios_baja += 1

    # ── Colorear filas por tipo de evaluación (al final, sin pisar fuente azul)
    aplicar_colores_evaluacion(ws)

    return cambios_alta, cambios_media_j, cambios_media_m, cambios_baja


# ══════════════════════════════════════════════════════════════════════════
#  CORRECCIÓN DE LENGUAJE EN ACTIVIDADES
# ══════════════════════════════════════════════════════════════════════════

# Verbos en 3ª persona plural al inicio de ítems → 2ª persona singular (imperativo)
_VERBOS_PLURAL_SINGULAR = {
    'presenten': 'presenta',      'elaboren': 'elabora',       'construyan': 'construye',
    'identifiquen': 'identifica', 'analicen': 'analiza',       'revisen': 'revisa',
    'discutan': 'discute',        'compartan': 'comparte',     'entreguen': 'entrega',
    'suban': 'sube',              'respondan': 'responde',     'completen': 'completa',
    'diseñen': 'diseña',          'apliquen': 'aplica',        'realicen': 'realiza',
    'observen': 'observa',        'seleccionen': 'selecciona', 'documenten': 'documenta',
    'redacten': 'redacta',        'propongan': 'propone',      'formulen': 'formula',
    'calculen': 'calcula',        'interpreten': 'interpreta', 'reflexionen': 'reflexiona',
    'investiguen': 'investiga',   'describan': 'describe',     'expliquen': 'explica',
}

# Ítems numerados que describen acción del docente en lugar del estudiante
_PAT_ACCION_DOCENTE = re.compile(
    r'^\d+\.\s+(la docente|el o la docente|el docente|la\/el docente)\s+\w',
    re.IGNORECASE
)

# Ítems numerados que son frases nominales (sin verbo imperativo)
_PAT_FRASE_NOMINAL = re.compile(
    r'^\d+\.\s+'
    r'(puesta en|revisión de|análisis de|presentación de|trabajo en|'
    r'discusión sobre|reflexión sobre|síntesis de|manejo de|introducción a)',
    re.IGNORECASE
)

# Líneas informales sin número que tampoco son título ni Propósito
# Detecta: notas con "+", keywords sueltos, notas entre paréntesis
_PAT_LINEA_INFORMAL = re.compile(
    r'^(?!'                          # NO empieza con…
    r'\d+\.'                         #   ítem numerado
    r'|Propósito'                    #   "Propósito:"
    r'|[A-ZÁÉÍÓÚ][a-záéíóúü]+'     #   palabra con mayúscula (título del momento)
    r')',
    re.IGNORECASE
)

# Verbos imperativos 2ª persona válidos para iniciar ítems
_VERBOS_IMPERATIVO = {
    'accede','activa','agrega','amplía','analiza','anota','aplica','argumenta',
    'atiende','ayuda','busca','calcula','categoriza','clasifica','colabora',
    'comenta','compara','comparte','completa','conecta','construye','consulta',
    'contesta','contrasta','contribuye','coopera','corrige','crea','cuestiona',
    'deduce','define','describe','desarrolla','destaca','diferencia','diseña',
    'documenta','elabora','elabora','emplea','enumera','escribe','escucha',
    'esquematiza','evalúa','examina','explica','explora','expone','familiarízate',
    'formula','grafica','identifica','implementa','infiere','interpreta',
    'investiga','justifica','lee','lista','localiza','mapea','marca','mide',
    'nota','observa','optimiza','organiza','participa','planifica','practica',
    'prepara','presenta','prioriza','profundiza','propone','publica','recibe',
    'recuerda','redacta','reflexiona','registra','relaciona','repasa','responde',
    'retoma','revisa','resume','selecciona','señala','sintetiza','sistematiza',
    'socializa','soluciona','subraya','sube','tabula','toma','trabaja','ubica',
    'usa','utiliza','valida','valora','verifica',
}

# Imperativas con acento incorrecto → forma correcta
_ACENTOS_IMPERATIVO = {
    'clasifíca': 'clasifica', 'identifíca': 'identifica',
    'analíza':   'analiza',   'organíza':   'organiza',
    'sintetíza': 'sintetiza', 'especifíca': 'especifica',
    'calífca':   'califica',  'justifíca':  'justifica',
}


def corregir_lenguaje_actividades(ws, log):
    """
    Corrige incoherencias de lenguaje en la columna Actividad (G):
      - "la docente" / "el docente" → "el o la docente" (lenguaje inclusivo UST)
      - Espacio faltante tras punto numerado: "2.Verbo" → "2. Verbo"
      - Verbos 3ª persona plural → 2ª persona singular (imperativo)
    Además registra advertencias para revisión manual:
      - Ítems que describen acción del docente (no instrucción al estudiante)
      - Ítems que son frases nominales sin verbo imperativo
    Devuelve el número de celdas modificadas.
    """
    celdas_modificadas = 0
    advertencias = []

    for row in ws.iter_rows(min_row=4, values_only=False):
        r = row[0].row
        momento = ws.cell(r, COL['MOMENTO']).value
        if not momento:
            continue  # filas sin momento (p.ej. filas de totales al final)

        cell = ws.cell(r, COL['ACTIVIDAD'])
        texto = cell.value
        if not isinstance(texto, str):
            continue

        texto_nuevo = texto
        cambios_celda = []

        # ── 1. Espacio faltante: "2.Verbo" → "2. Verbo" ──────────────────
        t2 = re.sub(r'(\d+)\.([^\s\d\n])', r'\1. \2', texto_nuevo)
        if t2 != texto_nuevo:
            cambios_celda.append('espacio faltante tras punto numerado')
            texto_nuevo = t2

        # ── 2a. "de/del el/la (o el/la)* docente" → "del o la docente" ─────
        # Cubre: "de el docente", "de la docente", "de el o la docente",
        #        "de el o el o la docente", "del o el o la docente", etc.
        t2 = re.sub(
            r'\bde(?:l)?\s+(?:el|la)(?:\s+o\s+(?:el|la))*\s+docente\b',
            'del o la docente',
            texto_nuevo, flags=re.IGNORECASE
        )
        if t2 != texto_nuevo:
            cambios_celda.append('"de el o… docente" → "del o la docente"')
            texto_nuevo = t2

        # ── 2b. "el/la docente" sueltos (no precedidos por "o ") ──────────
        # Lookbehind fijo: evita romper "el o la docente" ya corregido
        t2 = re.sub(r'(?<!o )\bel docente\b', 'el o la docente', texto_nuevo, flags=re.IGNORECASE)
        t2 = re.sub(r'(?<!o )la docente\b',   'el o la docente', t2,          flags=re.IGNORECASE)
        if t2 != texto_nuevo:
            cambios_celda.append('"la/el docente" → "el o la docente"')
            texto_nuevo = t2

        # ── 3. Verbos 3ª plural → 2ª singular al inicio de ítems ─────────
        def _plural_a_singular(m):
            verbo = m.group(2).lower()
            if verbo in _VERBOS_PLURAL_SINGULAR:
                singular = _VERBOS_PLURAL_SINGULAR[verbo]
                singular = singular[0].upper() + singular[1:]
                cambios_celda.append(f'"{m.group(2)}" → "{singular}" (3ª plural → 2ª singular)')
                return f'{m.group(1)}. {singular}'
            return m.group(0)

        t2 = re.sub(r'(\d+)\.\s+([A-ZÁÉÍÓÚ][a-záéíóúü]+)\b', _plural_a_singular, texto_nuevo)
        if t2 != texto_nuevo:
            texto_nuevo = t2

        # ── 3b. Imperativas con acento incorrecto (clasifíca → clasifica) ──
        for mal, bien in _ACENTOS_IMPERATIVO.items():
            patron_acento = re.compile(r'\b' + re.escape(mal) + r'\b', re.IGNORECASE)
            t2 = patron_acento.sub(bien, texto_nuevo)
            if t2 != texto_nuevo:
                cambios_celda.append(f'acento incorrecto: "{mal}" → "{bien}"')
                texto_nuevo = t2

        # ── Aplicar cambios si los hubo ───────────────────────────────────
        if texto_nuevo != texto:
            cell.value = texto_nuevo
            aplicar_azul(cell)
            for c in cambios_celda:
                log.append(f'    [Plan F{r}] Lenguaje actividad: {c}')
            celdas_modificadas += 1

        # ── 4. Advertencias por revisión manual ───────────────────────────
        for linea in texto_nuevo.split('\n'):
            linea = linea.strip()
            if not linea:
                continue

            # a) Acción descrita como tarea del docente
            if _PAT_ACCION_DOCENTE.match(linea):
                advertencias.append(
                    f'    [Plan F{r} {str(momento)[:12]}] ⚠️  Ítem describe acción del docente '
                    f'(reescribir como instrucción al estudiante): "{linea[:90]}"'
                )

            # b) Frase nominal sin verbo imperativo
            elif _PAT_FRASE_NOMINAL.match(linea):
                advertencias.append(
                    f'    [Plan F{r} {str(momento)[:12]}] ⚠️  Ítem nominal sin verbo '
                    f'(iniciar con verbo imperativo: Revisa, Lee, Aplica…): "{linea[:90]}"'
                )

            # c) Línea con "+" como separador de actividades
            elif '+' in linea and not linea.startswith('Propósito'):
                advertencias.append(
                    f'    [Plan F{r} {str(momento)[:12]}] ⚠️  Usa "+" como separador '
                    f'(reemplazar por ítems numerados con verbo imperativo): "{linea[:90]}"'
                )

            # d) Línea no numerada y no título que no empieza con verbo imperativo
            elif (linea and not re.match(r'^\d+\.', linea)
                      and not linea.startswith('Propósito')
                      and not re.match(r'^[A-ZÁÉÍÓÚ][a-záéíóúü ]+$', linea)):
                primer_word = re.split(r'[\s,.]', linea)[0].lower().strip('¡!')
                if primer_word and primer_word not in _VERBOS_IMPERATIVO:
                    advertencias.append(
                        f'    [Plan F{r} {str(momento)[:12]}] ⚠️  Línea no inicia con verbo '
                        f'imperativo (Revisa, Lee, Aplica…): "{linea[:90]}"'
                    )

    if advertencias:
        log.append('\n  [Advertencias de lenguaje — requieren revisión manual]')
        log.extend(advertencias)

    return celdas_modificadas


# ══════════════════════════════════════════════════════════════════════════
#  VERIFICACIÓN A+Se (APRENDIZAJE + SERVICIO E-LEARNING)
# ══════════════════════════════════════════════════════════════════════════

HITOS_AS = [
    # (id, descripción, modo)
    ('RE1',   'Reflexión Estructurada 1 presente (diagnóstico)',               'auto'),
    ('RE2',   'Reflexión Estructurada 2 presente (diseño del servicio)',        'auto'),
    ('RE3',   'Reflexión Estructurada 3 presente (cierre/evaluación)',          'auto'),
    ('SC',    'Socio Comunitario mencionado en actividades',                    'auto'),
    ('SC_S7', 'Estudiantes conocen al SC antes de semana 7',                   'auto'),
    ('ENC',   'Encuestas A+S programadas en semanas 17-18',                    'auto'),
    ('GRUP',  'Trabajo grupal/equipo presente (duplas o grupos)',               'auto'),
    ('PROD',  'Evaluación sumativa incluye producto o actuación del estudiante','auto'),
    ('RE3_S', 'Reflexión Estructurada 3 en momento sincrónico',                'auto'),
    ('ETAP',  'Las 4 etapas A+S representadas en la planificación',            'manual'),
    ('CIER',  'Actividad de cierre con el socio comunitario presente',         'manual'),
]


def verificar_as(ws_plan, log):
    """
    Verifica los hitos obligatorios A+Se según el lineamiento UST 2025
    en la hoja Planificación por unidades.
    Devuelve (n_ok, n_error, n_manual).
    """
    # Recolectar datos por fila
    filas = []
    semana_actual = None
    for row in ws_plan.iter_rows(min_row=4, values_only=True):
        semana    = row[COL['SEMANA']    - 1]
        momento   = row[COL['MOMENTO']  - 1]
        actividad = row[COL['ACTIVIDAD']- 1]
        tipo      = row[COL['TIPO']     - 1]
        proc      = row[COL['PROC']     - 1]
        indiv     = row[COL['INDIV']    - 1]
        if semana:
            try:
                semana_actual = int(re.sub(r'[^\d]', '', str(semana)))
            except ValueError:
                pass
        if actividad or tipo:
            filas.append({
                'semana':    semana_actual,
                'momento':   str(momento   or '').lower(),
                'actividad': str(actividad or '').lower(),
                'tipo':      str(tipo      or '').lower(),
                'proc':      str(proc      or '').lower(),
                'indiv':     str(indiv     or '').lower(),
            })

    resultados = {}

    # ── RE1 / RE2 / RE3 — contar Reflexiones Estructuradas ───────────────
    def es_re(texto):
        return bool(re.search(
            r'reflexi[oó]n\s+estructurada|RE\s*[123]',
            texto, re.IGNORECASE
        ))

    filas_re = [f for f in filas if es_re(f['actividad'])]
    n_re = len(filas_re)

    resultados['RE1'] = (
        ('✅', 'Reflexión Estructurada 1 encontrada') if n_re >= 1
        else ('❌', 'No se encontró ninguna Reflexión Estructurada en actividades')
    )
    resultados['RE2'] = (
        ('✅', 'Reflexión Estructurada 2 encontrada') if n_re >= 2
        else ('❌', f'Solo {n_re} reflexión(es) encontrada(s) — se necesitan 3')
    )
    resultados['RE3'] = (
        ('✅', 'Reflexión Estructurada 3 encontrada') if n_re >= 3
        else ('❌', f'Solo {n_re} reflexión(es) encontrada(s) — se necesitan 3')
    )

    # ── SC — Socio Comunitario mencionado ─────────────────────────────────
    def menciona_sc(texto):
        return bool(re.search(r'socio\s+comunitario', texto, re.IGNORECASE))

    filas_sc = [f for f in filas if menciona_sc(f['actividad'])]
    resultados['SC'] = (
        ('✅', f'"Socio comunitario" encontrado ({len(filas_sc)} fila(s))') if filas_sc
        else ('❌', '"Socio comunitario" no aparece en ninguna actividad')
    )

    # ── SC_S7 — SC conocido antes de semana 7 ────────────────────────────
    primera_semana_sc = next(
        (f['semana'] for f in filas_sc if f['semana'] is not None), None
    )
    if primera_semana_sc is None:
        resultados['SC_S7'] = ('❌', 'Socio comunitario sin semana asignada')
    elif primera_semana_sc <= 7:
        resultados['SC_S7'] = ('✅', f'SC en semana {primera_semana_sc} (≤ 7 ✓)')
    else:
        resultados['SC_S7'] = ('⚠️',
            f'SC en semana {primera_semana_sc} — lineamiento recomienda ≤ semana 7')

    # ── ENC — Encuestas A+S en semanas 17-18 ─────────────────────────────
    filas_enc = [f for f in filas
                 if re.search(r'encuesta', f['actividad'], re.IGNORECASE)]
    semanas_enc = [f['semana'] for f in filas_enc if f['semana'] is not None]
    if not semanas_enc:
        resultados['ENC'] = ('⚠️',
            'Encuesta A+S no encontrada — debe programarse en semanas 17-18')
    elif all(17 <= s <= 18 for s in semanas_enc):
        resultados['ENC'] = ('✅', f'Encuesta A+S en semana(s) {semanas_enc}')
    else:
        resultados['ENC'] = ('⚠️',
            f'Encuesta en semana(s) {semanas_enc} — debe estar en semanas 17-18')

    # ── GRUP — Trabajo grupal presente ───────────────────────────────────
    tiene_grupal = any('grupal' in f['indiv'] for f in filas)
    resultados['GRUP'] = (
        ('✅', 'Trabajo grupal/equipo presente') if tiene_grupal
        else ('❌', 'Sin trabajo grupal — A+S requiere trabajo en duplas o equipos')
    )

    # ── PROD — Sumativa con producto o actuación ──────────────────────────
    tiene_prod = any(
        'sumativa' in f['tipo'] and
        any(k in f['proc'] for k in ('producciones', 'actuación', 'actuacion', 'informe'))
        for f in filas
    )
    resultados['PROD'] = (
        ('✅', 'Evaluación sumativa incluye producto o actuación del estudiante')
        if tiene_prod
        else ('❌',
              'Sumativa sin producto/actuación — A+S evalúa el servicio entregado')
    )

    # ── RE3_S — Reflexión 3 en momento sincrónico ─────────────────────────
    re3_sinc = any(
        es_re(f['actividad']) and 'sincr' in f['momento']
        for f in filas
    )
    resultados['RE3_S'] = (
        ('✅', 'Reflexión Estructurada 3 en momento sincrónico') if re3_sinc
        else ('⚠️', 'No se confirmó RE3 sincrónica — el lineamiento lo recomienda')
    )

    # ── Criterios manuales ────────────────────────────────────────────────
    resultados['ETAP'] = ('⚠️ manual',
        'Verificar que las 4 etapas A+S estén representadas (Diagnóstico, '
        'Planificación, Implementación, Evaluación/Cierre)')
    resultados['CIER'] = ('⚠️ manual',
        'Verificar que haya una actividad de cierre con el socio comunitario')

    # ── Reporte ───────────────────────────────────────────────────────────
    log.append('\n  [Verificación A+Se — Hitos obligatorios (Lineamiento UST 2025)]')
    n_ok = n_err = n_manual = 0
    for hito_id, desc, _ in HITOS_AS:
        estado, detalle = resultados.get(hito_id, ('⚠️', 'No verificado'))
        desc_corta = desc[:65] + '…' if len(desc) > 65 else desc
        log.append(f'    {estado} {desc_corta}')
        if estado != '✅':
            log.append(f'         → {detalle}')
        if estado == '✅':
            n_ok += 1
        elif 'manual' in str(estado):
            n_manual += 1
        else:
            n_err += 1

    log.append(
        f'\n    Resultado A+Se: {n_ok}✅  {n_err}❌  {n_manual}⚠️ manual'
        f'  (de {len(HITOS_AS)} hitos)'
    )
    return n_ok, n_err, n_manual


# ══════════════════════════════════════════════════════════════════════════
#  PROCESAMIENTO POR ASIGNATURA
# ══════════════════════════════════════════════════════════════════════════

def procesar_asignatura(carpeta_asig, dry_run=False, programa=None, es_as=False):
    nombre = os.path.basename(carpeta_asig)
    log = [f'\n{"═"*70}', f'  ASIGNATURA: {nombre}', f'{"═"*70}']

    carpeta_correcciones = os.path.join(carpeta_asig, 'Correcciones')
    carpeta_enviado      = os.path.join(carpeta_asig, 'Enviado a DEL')
    carpeta_revisado     = os.path.join(carpeta_asig, 'Revisado')

    # ── Determinar fuente ──────────────────────────────────────────────────
    plan_path   = encontrar_archivo(carpeta_correcciones, '*.xlsx')
    escala_path = None
    es_original = False

    if plan_path:
        fuente = 'Correcciones'
        log.append(f'  Fuente : Correcciones/ (revisada por DEL)')
        # intentar encontrar escala igual
        _, escala_path = encontrar_planificacion_enviada(carpeta_enviado)
    else:
        # Buscar en Enviado a DEL
        if os.path.isdir(carpeta_enviado):
            plan_path, escala_path = encontrar_planificacion_enviada(carpeta_enviado)
        if plan_path:
            fuente = 'Enviado a DEL'
            es_original = True
            log.append(f'  Fuente : Enviado a DEL/ (planificación original — sin revisión DEL)')
        else:
            log.append('  ⚠️  No se encontró ninguna planificación (.xlsx) — omitida')
            return log, False

    log.append(f'  Archivo: {os.path.basename(plan_path)}')

    # ── Observaciones específicas del archivo de escala (contexto adicional) ─
    obs_escala = leer_observaciones_escala(escala_path)
    if escala_path:
        log.append(f'  Escala : {os.path.basename(escala_path)}')
        if obs_escala:
            log.append('  [Observaciones DEL registradas en la escala — solo referencia]')
            for criterio, obs in obs_escala:
                log.append(f'    • {criterio[:65]}')
                log.append(f'      → {obs[:130]}')
    else:
        log.append('  Escala : [no disponible — se aplican los 41 criterios estándar UST]')

    # ── Cargar libro ───────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(plan_path)
    hojas = wb.sheetnames
    log.append(f'\n  Hojas detectadas: {hojas}')

    total_sintesis = 0
    total_alta = 0
    total_media_j = 0
    total_media_m = 0
    total_baja = 0

    # ── Síntesis didáctica ─────────────────────────────────────────────────
    if 'Síntesis didáctica' in hojas:
        log.append('\n  [Síntesis didáctica]')
        n = corregir_sintesis(wb['Síntesis didáctica'], log)
        total_sintesis = n
        if n == 0:
            log.append('    Sin cambios necesarios.')
    else:
        log.append('\n  ⚠️  Hoja "Síntesis didáctica" no encontrada')

    # ── Planificación por unidades ─────────────────────────────────────────
    total_lenguaje = 0
    if 'Planificación por unidades' in hojas:
        log.append('\n  [Planificación por unidades]')
        a, mj, mm, b = corregir_planificacion(wb['Planificación por unidades'], log)
        total_alta, total_media_j, total_media_m, total_baja = a, mj, mm, b
        total_lenguaje = corregir_lenguaje_actividades(wb['Planificación por unidades'], log)
        if a + mj + mm + b + total_lenguaje == 0:
            log.append('    Sin cambios necesarios.')
    else:
        log.append('\n  ⚠️  Hoja "Planificación por unidades" no encontrada')

    # ── Verificación contra escala (post-correcciones) ─────────────────────
    fuente_escala = 'archivo' if escala_path else 'estandar'
    resultados_escala = verificar_escala(wb)
    log.extend(formatear_verificacion(resultados_escala, fuente_escala))

    # ── Verificación contra programa oficial (si se proveyó PDF) ───────────
    n_discrepancias = verificar_contra_programa(wb, programa, log)

    # ── Verificación A+Se (si corresponde) ──────────────────────────────
    n_as_ok = n_as_err = n_as_manual = 0
    if es_as:
        if 'Planificación por unidades' in hojas:
            n_as_ok, n_as_err, n_as_manual = verificar_as(
                wb['Planificación por unidades'], log
            )
        else:
            log.append('\n  ⚠️  A+Se: hoja "Planificación por unidades" no encontrada')

    # ── Guardar ────────────────────────────────────────────────────────────
    if not dry_run:
        os.makedirs(carpeta_revisado, exist_ok=True)
        # Si viene de original, renombrar para distinguirla
        nombre_salida = (nombre_revisado(os.path.basename(plan_path))
                         if es_original else os.path.basename(plan_path))
        dest = os.path.join(carpeta_revisado, nombre_salida)
        wb.save(dest)
        log.append(f'\n  💾 Guardado en: Revisado/{nombre_salida}')

    total = total_sintesis + total_alta + total_media_j + total_media_m + total_baja + total_lenguaje
    log.append(f'\n  RESUMEN DE CAMBIOS:')
    log.append(f'    Síntesis didáctica      : {total_sintesis} correcciones')
    log.append(f'    Alta  (proc/instr/tipo) : {total_alta} correcciones')
    log.append(f'    Media (Medio entrega J) : {total_media_j} correcciones')
    log.append(f'    Media (Indiv/Grupal  M) : {total_media_m} correcciones')
    log.append(f'    Baja  (% vacíos → 0)   : {total_baja} correcciones')
    log.append(f'    Lenguaje (actividades)  : {total_lenguaje} celdas corregidas')
    log.append(f'    ─────────────────────────────')
    log.append(f'    TOTAL                  : {total} correcciones')
    log.append(f'    🔵 Celdas modificadas marcadas con fuente azul (#0070C0)')
    if programa:
        log.append(f'    📄 Verificación vs programa: {n_discrepancias} discrepancia(s) encontrada(s)')
    if es_as:
        log.append(f'    📌 Verificación A+Se: {n_as_ok}✅  {n_as_err}❌  {n_as_manual}⚠️ manual')
    if es_original:
        log.append(f'    ⚠️  Revisar manualmente las observaciones de la escala')

    return log, True


# ══════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='Revisa y corrige planificaciones didácticas UST en ~/Desktop/DEL'
    )
    parser.add_argument(
        '--base', type=str, default=None,
        help='Carpeta raíz con las asignaturas (por defecto: ~/Desktop/DEL). '
             'Útil en Google Colab: --base /content/drive/MyDrive/DEL'
    )
    parser.add_argument(
        '--dry-run', action='store_true',
        help='Muestra correcciones sin guardar archivos'
    )
    parser.add_argument(
        '--asignatura', type=str, default=None,
        help='Procesa solo esta asignatura (nombre exacto de la carpeta)'
    )
    parser.add_argument(
        '--log', type=str, default=None,
        help='Guarda el reporte en este archivo (ej: reporte.txt)'
    )
    args = parser.parse_args()

    base = os.path.expanduser(args.base) if args.base else os.path.expanduser('~/Desktop/DEL')
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')

    lineas = [
        f'REPORTE DE REVISIÓN — PLANIFICACIONES UST',
        f'Generado: {timestamp}',
        f'Carpeta base: {base}',
        f'Modo: {"DRY-RUN (sin guardar)" if args.dry_run else "CORRECCIÓN APLICADA"}',
    ]

    # Detectar carpetas de asignaturas
    if args.asignatura:
        carpetas = [os.path.join(base, args.asignatura)]
    else:
        carpetas = sorted([
            os.path.join(base, d)
            for d in os.listdir(base)
            if os.path.isdir(os.path.join(base, d)) and not d.startswith('.')
        ])

    if not carpetas:
        print('No se encontraron carpetas de asignaturas en', base)
        return

    procesadas = 0
    omitidas = 0

    for carpeta in carpetas:
        log, ok = procesar_asignatura(carpeta, dry_run=args.dry_run)
        lineas.extend(log)
        if ok:
            procesadas += 1
        else:
            omitidas += 1

    lineas.append(f'\n{"═"*70}')
    lineas.append(f'  TOTAL: {procesadas} asignatura(s) procesada(s), {omitidas} omitida(s)')
    lineas.append(f'{"═"*70}')

    reporte = '\n'.join(lineas)
    print(reporte)

    # Guardar log si se solicitó
    log_path = args.log or os.path.join(base, f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.txt')
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write(reporte)
    print(f'\n📄 Reporte guardado en: {log_path}')


if __name__ == '__main__':
    main()
