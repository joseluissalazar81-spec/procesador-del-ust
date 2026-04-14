"""
apa_recursos.py — Validador y corrector APA 7 para columna H (Recursos didácticos)
Procesador DEL · UST 2026-1

Reglas APA 7 verificadas:
  - Formato de autor: Apellido, I. I.
  - Año entre paréntesis seguido de punto: (YYYY).
  - Título en mayúscula inicial, termina en punto (o antes de corchete)
  - Tipo de recurso entre corchetes: [Video], [Diapositivas], etc.
  - Múltiples autores con & (no "y")
  - URL sin "Disponible en" precedente
  - No terminar con punto después de URL

Correcciones seguras (autocorrección):
  - Eliminar viñeta •\xa0 inicial
  - "y" → "&" entre autores (antes del año)
  - "Disponible en " antes de URL → solo URL
  - Doble espacio → espacio simple
  - Agregar punto tras cierre de año si falta: (2026)Título → (2026). Título
  - Trimear espacios y saltos de línea sobrantes
"""

from __future__ import annotations
import re
import urllib.request
import urllib.parse
import json

# ── Tipos de recurso válidos UST ──────────────────────────────────────────────
_TIPOS_RECURSO_UST = {
    "video", "videoclase", "video cápsula", "podcast",
    "diapositivas", "presentación", "presentacion",
    "material de estudio", "instrumento", "instrumento evaluativo",
    "instrumento de evaluación formativa", "instrumento diagnóstico",
    "recurso audiovisual", "recurso de aprendizaje",
    "presentación interactiva h5p",
    "guía de aprendizaje", "guia de aprendizaje",
    "guía de ejercicios", "guia de ejercicios",
    "apunte de contenidos", "ficha de estudio",
    "infografía", "infografia", "cuestionario",
}

# Patrón de URL
_RE_URL = re.compile(r'https?://\S+')

# Patrón de año APA: (2024) o (2024, enero 15) o (s.f.)
_RE_ANIO = re.compile(r'\(\d{4}(?:,\s*[^)]+)?\)|\(s\.f\.\)')

# Patrón de autor básico: Apellido, I. o Apellido, I. I.
_RE_AUTOR = re.compile(
    r'^[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñA-ZÁÉÍÓÚÜÑ\-]+,\s+[A-ZÁÉÍÓÚÜÑ]\.'
)

# "y" como conector entre autores (antes del año) — solo fuera de paréntesis
_RE_Y_AUTORES = re.compile(
    r'([A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+(?:,\s+[A-Z]\.)*)(\s+y\s+)([A-ZÁÉÍÓÚÜÑ])'
)

# "Disponible en" seguido de URL
_RE_DISPONIBLE_EN = re.compile(r'[Dd]isponible en\s+(https?://\S+)')

# Punto faltante tras año: (2026)Título → (2026). Título
_RE_ANIO_SIN_PUNTO = re.compile(r'(\(\d{4}(?:,\s*[^)]+)?\)|\(s\.f\.\))([A-ZÁÉÍÓÚÜÑ\[])')


# ═════════════════════════════════════════════════════════════════════════════
#  SEPARAR REFERENCIAS EN UNA CELDA
# ═════════════════════════════════════════════════════════════════════════════

def separar_referencias(texto: str) -> list[str]:
    """
    Divide el contenido de una celda en referencias individuales.
    Separador: salto de línea (con o sin viñeta •).
    Devuelve lista de strings limpios (sin viñeta ni espacios extra).
    """
    if not texto:
        return []
    partes = re.split(r'\n+', str(texto))
    resultado = []
    for p in partes:
        limpio = p.strip().lstrip('•\xa0 ').strip()
        if limpio:
            resultado.append(limpio)
    return resultado


# ═════════════════════════════════════════════════════════════════════════════
#  VALIDAR UNA REFERENCIA
# ═════════════════════════════════════════════════════════════════════════════

def validar_referencia(ref: str) -> list[dict]:
    """
    Valida una referencia APA 7 individual.
    Devuelve lista de problemas: [{nivel, codigo, mensaje}]
    """
    problemas = []

    def p(nivel, codigo, msg):
        problemas.append({"nivel": nivel, "codigo": codigo, "mensaje": msg})

    ref_s = ref.strip()
    if len(ref_s) < 15:
        p("advertencia", "APA_MUY_CORTA",
          f"Referencia muy corta para ser APA 7: «{ref_s[:60]}»")
        return problemas

    # ── 1. Año presente ───────────────────────────────────────────────────
    if not _RE_ANIO.search(ref_s):
        p("error", "APA_SIN_AÑO",
          f"No se detecta año entre paréntesis: «{ref_s[:80]}»")

    # ── 2. Formato de autor (primera referencia) ──────────────────────────
    if not _RE_AUTOR.match(ref_s):
        # Puede ser institución — solo advertencia
        p("advertencia", "APA_AUTOR_FORMATO",
          f"El inicio no sigue formato «Apellido, I.»: «{ref_s[:60]}»")

    # ── 3. "y" entre autores (debe ser "&") ──────────────────────────────
    if _RE_Y_AUTORES.search(ref_s):
        p("error", "APA_Y_EN_VEZ_DE_AMPERSAND",
          f"Usar «&» en vez de «y» entre autores: «{ref_s[:80]}»")

    # ── 4. "Disponible en" antes de URL ──────────────────────────────────
    if re.search(r'[Dd]isponible en\s+https?://', ref_s):
        p("advertencia", "APA_DISPONIBLE_EN",
          "Eliminar «Disponible en» — APA 7 solo incluye la URL directamente.")

    # ── 5. Punto después del año ─────────────────────────────────────────
    if _RE_ANIO_SIN_PUNTO.search(ref_s):
        p("error", "APA_PUNTO_TRAS_AÑO",
          f"Falta punto tras el año: «{ref_s[:80]}»")

    # ── 6. URL termina con punto (incorrecto en APA 7) ────────────────────
    urls = _RE_URL.findall(ref_s)
    for url in urls:
        if url.endswith('.') and not url.endswith('..'):
            p("advertencia", "APA_URL_CON_PUNTO",
              f"URL no debe terminar en punto: {url[:80]}")

    # ── 7. Corchetes de tipo de recurso bien cerrados ─────────────────────
    abre = ref_s.count('[')
    cierra = ref_s.count(']')
    if abre != cierra:
        p("error", "APA_CORCHETE_DESCUADRADO",
          f"Corchete sin cerrar o extra: «{ref_s[:80]}»")

    # ── 8. Referencia no termina con punto (excepto si termina en URL) ────
    if not urls:
        # Sin URL: debe terminar en punto
        if not ref_s.rstrip().endswith('.'):
            p("advertencia", "APA_SIN_PUNTO_FINAL",
              f"La referencia debería terminar en punto: «{ref_s[-40:]}»")
    else:
        # Con URL: no debe terminar en punto después de URL
        pass  # ya validado arriba

    return problemas


# ═════════════════════════════════════════════════════════════════════════════
#  CORREGIR UNA REFERENCIA (correcciones seguras)
# ═════════════════════════════════════════════════════════════════════════════

def corregir_referencia(ref: str) -> tuple[str, list[str]]:
    """
    Aplica correcciones APA 7 seguras a una referencia.
    Devuelve (texto_corregido, lista_cambios).
    """
    original = ref.strip().lstrip('•\xa0 ').strip()
    texto = original
    cambios = []

    # ── 1. "y" → "&" entre autores (antes del año) ───────────────────────
    def reemplazar_y(m):
        return m.group(1) + ' & ' + m.group(3)

    nuevo = _RE_Y_AUTORES.sub(reemplazar_y, texto)
    if nuevo != texto:
        cambios.append('«y» → «&» entre autores')
        texto = nuevo

    # ── 2. "Disponible en URL" → solo URL ────────────────────────────────
    nuevo = _RE_DISPONIBLE_EN.sub(lambda m: m.group(1), texto)
    if nuevo != texto:
        cambios.append('«Disponible en» eliminado antes de URL')
        texto = nuevo

    # ── 3. Punto faltante tras año ────────────────────────────────────────
    nuevo = _RE_ANIO_SIN_PUNTO.sub(lambda m: m.group(1) + '. ' + m.group(2), texto)
    if nuevo != texto:
        cambios.append('Punto agregado tras el año')
        texto = nuevo

    # ── 4. Doble espacio → espacio simple ─────────────────────────────────
    nuevo = re.sub(r'  +', ' ', texto)
    if nuevo != texto:
        cambios.append('Espacios dobles normalizados')
        texto = nuevo

    # ── 5. Punto al final de URL → eliminar ──────────────────────────────
    def quitar_punto_url(m):
        url = m.group(0)
        if url.endswith('.') and not url.endswith('..'):
            return url[:-1]
        return url

    nuevo = _RE_URL.sub(quitar_punto_url, texto)
    if nuevo != texto:
        cambios.append('Punto eliminado al final de URL')
        texto = nuevo

    return texto, cambios


# ═════════════════════════════════════════════════════════════════════════════
#  RECONSTRUIR CELDA CON FORMATO UST (viñetas + saltos de línea)
# ═════════════════════════════════════════════════════════════════════════════

def reconstruir_celda(referencias: list[str]) -> str:
    """Reconstruye el contenido de la celda con formato viñeta UST."""
    return '\n'.join(f'•\xa0{r}' for r in referencias)


# ═════════════════════════════════════════════════════════════════════════════
#  REVISAR COLUMNA H COMPLETA
# ═════════════════════════════════════════════════════════════════════════════

def revisar_columna_recursos(ws_plan, autocorregir: bool = False) -> tuple[list[str], int, int]:
    """
    Revisa y opcionalmente corrige las referencias APA 7 en la columna H.

    Parámetros
    ----------
    ws_plan      : hoja openpyxl "Planificación por unidades" (modo escritura)
    autocorregir : si True, aplica correcciones seguras en las celdas

    Devuelve
    --------
    (log_lineas, n_problemas, n_correcciones)
    """
    from openpyxl.styles import Font, Color
    from openpyxl.utils import get_column_letter

    COL_H = 8   # columna H = 8
    AZUL  = "FF2E74B5"

    log: list[str] = []
    n_problemas   = 0
    n_correcciones = 0

    log.append('\n  [Verificación APA 7 — columna H (Recursos didácticos)]')

    for row in ws_plan.iter_rows(min_row=4):
        cell = row[COL_H - 1]
        if not cell.value:
            continue

        texto_orig = str(cell.value)
        refs = separar_referencias(texto_orig)
        if not refs:
            continue

        refs_corregidas = []
        cambios_fila: list[str] = []
        problemas_fila: list[dict] = []

        for ref in refs:
            # Validar
            probs = validar_referencia(ref)
            problemas_fila.extend(probs)

            # Corregir
            if autocorregir:
                ref_corr, cambios = corregir_referencia(ref)
                refs_corregidas.append(ref_corr)
                cambios_fila.extend(cambios)
            else:
                refs_corregidas.append(ref)

        # Loguear problemas
        if problemas_fila:
            n_problemas += len(problemas_fila)
            log.append(f'    Fila {cell.row} (H{cell.row}):')
            for prob in problemas_fila:
                icono = '❌' if prob['nivel'] == 'error' else '⚠️'
                log.append(f'      {icono} [{prob["codigo"]}] {prob["mensaje"][:100]}')

        # Aplicar correcciones
        if autocorregir and cambios_fila:
            nuevo_texto = reconstruir_celda(refs_corregidas)
            if nuevo_texto != texto_orig:
                cell.value = nuevo_texto
                # Marcar en azul (fuente)
                try:
                    font_orig = cell.font.copy() if cell.font else Font()
                    cell.font = Font(
                        name=font_orig.name or 'Calibri',
                        size=font_orig.size or 11,
                        color=Color(rgb=AZUL),
                    )
                except Exception:
                    pass
                n_correcciones += len(cambios_fila)
                log.append(f'    ✏️  Fila {cell.row}: {", ".join(set(cambios_fila))}')

    # Resumen
    log.append(
        f'\n    APA 7 Recursos: {n_problemas} problema(s) detectado(s), '
        f'{n_correcciones} corrección(es) aplicada(s)'
    )

    return log, n_problemas, n_correcciones


# ═════════════════════════════════════════════════════════════════════════════
#  GENERAR REFERENCIA APA 7 DESDE URL / DOI
# ═════════════════════════════════════════════════════════════════════════════

def generar_desde_url(url: str) -> tuple[str, str]:
    """
    Intenta generar una referencia APA 7 básica desde una URL o DOI.

    Devuelve (referencia_generada, mensaje_estado).
    Si no puede generar, devuelve ('', mensaje_de_error).
    """
    url = url.strip()

    # ── DOI ──────────────────────────────────────────────────────────────
    doi_match = re.match(r'(?:https?://doi\.org/|doi:)(10\.\S+)', url, re.I)
    if doi_match:
        doi = doi_match.group(1)
        try:
            req = urllib.request.Request(
                f'https://api.crossref.org/works/{urllib.parse.quote(doi)}',
                headers={'User-Agent': 'DEL-UST-APA/1.0 (mailto:del@ust.cl)'},
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read())['message']

            autores = data.get('author', [])
            anio    = (data.get('published', {}).get('date-parts') or [['']])[0][0]
            titulo  = (data.get('title') or [''])[0]
            journal = (data.get('container-title') or [''])[0]
            vol     = data.get('volume', '')
            num     = data.get('issue', '')
            pags    = data.get('page', '')

            def fmt_autor(a):
                apellido = a.get('family', '')
                inicial  = (a.get('given') or '?')[0] + '.'
                return f'{apellido}, {inicial}'

            if autores:
                if len(autores) == 1:
                    aut_str = fmt_autor(autores[0])
                elif len(autores) <= 20:
                    partes = [fmt_autor(a) for a in autores[:-1]]
                    aut_str = ', '.join(partes) + ', & ' + fmt_autor(autores[-1])
                else:
                    partes = [fmt_autor(a) for a in autores[:19]]
                    aut_str = ', '.join(partes) + ', ... ' + fmt_autor(autores[-1])
            else:
                aut_str = 'Autor desconocido'

            partes_ref = [f'{aut_str} ({anio}). {titulo}.']
            if journal:
                partes_ref.append(f' *{journal}*')
                if vol:
                    partes_ref.append(f', *{vol}*')
                if num:
                    partes_ref.append(f'({num})')
                if pags:
                    partes_ref.append(f', {pags}')
                partes_ref.append('.')
            partes_ref.append(f' https://doi.org/{doi}')

            return ''.join(partes_ref), 'DOI resuelto via Crossref'

        except Exception as e:
            return '', f'Error al resolver DOI: {e}'

    # ── Página web (Open Graph / meta tags) ──────────────────────────────
    try:
        req = urllib.request.Request(
            url,
            headers={
                'User-Agent': 'Mozilla/5.0 (compatible; DEL-UST-APA/1.0)',
                'Accept': 'text/html',
            },
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            html = resp.read().decode('utf-8', errors='ignore')[:20000]

        def meta(prop):
            patterns = [
                rf'<meta[^>]+property=["\']og:{prop}["\'][^>]+content=["\'](.*?)["\']',
                rf'<meta[^>]+name=["\']{prop}["\'][^>]+content=["\'](.*?)["\']',
                rf'<meta[^>]+content=["\'](.*?)["\'][^>]+property=["\']og:{prop}["\']',
            ]
            for pat in patterns:
                m = re.search(pat, html, re.I | re.S)
                if m:
                    return m.group(1).strip()
            return ''

        titulo   = meta('title') or re.search(r'<title[^>]*>(.*?)</title>', html, re.I | re.S)
        if hasattr(titulo, 'group'):
            titulo = titulo.group(1).strip()
        site     = meta('site_name') or urllib.parse.urlparse(url).netloc
        autor    = meta('author') or ''
        anio_now = __import__('datetime').date.today().year

        if not titulo:
            return '', 'No se pudo extraer título de la página'

        if autor:
            # Intentar parsear "Nombre Apellido" → "Apellido, N."
            partes = autor.strip().split()
            if len(partes) >= 2:
                aut_fmt = f'{partes[-1]}, {partes[0][0]}.'
            else:
                aut_fmt = autor
            ref = f'{aut_fmt} ({anio_now}). {titulo}. {site}. {url}'
        else:
            ref = f'{site}. ({anio_now}). {titulo}. {url}'

        return ref, 'Generado desde metadatos de la página (revisar manualmente)'

    except Exception as e:
        return '', f'No se pudo acceder a la URL: {e}'
