"""
apa_llm.py — Revisión APA 7 con LLM · Procesador DEL · UST
Soporta: Ollama (local, por defecto), Claude API, OpenAI, Grok.

Uso mínimo:
    resultado = revisar_referencias_llm(texto_celda)
    # → {"referencias": [...], "cambios": [...], "ok": True}
"""

from __future__ import annotations
import json
import re
import urllib.request
import urllib.parse

# ── Modelos Ollama locales ────────────────────────────────────────────────────
# NOTA: 16 GB RAM — no correr ambos al mismo tiempo.
MODELO_VISION = "qwen2.5vl:latest"    # Para imágenes (requiere modelo multimodal)
MODELO_TEXTO  = "qwen3.5:latest"      # Para texto largo

# ── Configuración de backends ────────────────────────────────────────────────

BACKENDS = ["ollama", "claude", "openai", "grok"]

BACKEND_DEFAULTS = {
    "ollama": {
        "url":   "http://localhost:11434/v1/chat/completions",
        "model": MODELO_TEXTO,
        "key":   "ollama",
    },
    "claude": {
        "url":   "https://api.anthropic.com/v1/messages",
        "model": "claude-haiku-4-5-20251001",   # más económico
        "key":   "",   # requiere ANTHROPIC_API_KEY
    },
    "openai": {
        "url":   "https://api.openai.com/v1/chat/completions",
        "model": "gpt-4o-mini",
        "key":   "",   # requiere OPENAI_API_KEY
    },
    "grok": {
        "url":   "https://api.x.ai/v1/chat/completions",
        "model": "grok-3-mini",
        "key":   "",   # requiere XAI_API_KEY
    },
}

# ── Prompt del sistema ────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """Eres un experto en norma APA 7 para planificaciones didácticas e-learning de la Universidad Santo Tomás (UST), Chile.
Tu tarea es revisar y corregir referencias y recursos bibliográficos en columna H del Excel de planificación.

CONTEXTO UST — cada recurso declarado en la planificación debe tener 5 campos:
  1. Título: nombre del recurso entre comillas
  2. Autoría: Apellido, I. (año) en formato APA 7
  3. Tipo: entre corchetes — [Video], [Diapositivas], [Artículo], [H5P], [Guía], [Podcast], [Libro], etc.
  4. Acceso: "Disponible en el aula virtual" o URL completa (sin "Disponible en" antes de URL)
  5. Extensión: N slides / N pgs. / N min.

Ejemplo CORRECTO UST:
  "Presentación S2 Ansiedad Precompetitiva". González, M. (2019). [Diapositivas, 16 slides]. Disponible en el aula virtual.

Ejemplo INCOMPLETO (reportar como error):
  González et al. (2019) — faltan tipo, acceso y extensión.

REGLAS APA 7 adicionales:
1. Autor: Apellido, I. I. (iniciales con punto). Varios autores: coma entre ellos, & antes del último.
2. Año entre paréntesis seguido de punto: (2024).
3. Título en redonda, mayúscula solo en primera palabra y nombres propios.
4. Revista o libro en cursiva (indicar con *cursiva*).
5. DOI con formato: https://doi.org/...
6. URLs directas sin "Disponible en" precedente (excepto recursos del aula virtual UST).
7. Sin punto después de URL.
8. Usar & (no "y") entre coautores.

RESPONDE ÚNICAMENTE con este JSON exacto, sin texto adicional:
{
  "referencias": [
    {
      "original": "texto original",
      "corregida": "texto corregido con los 5 campos completos",
      "cambios": ["descripción del cambio 1", "descripción del cambio 2"],
      "ok": true
    }
  ]
}

- ok=true → referencia correcta (no se modifica)
- ok=false → tiene errores (reportar qué falta en "cambios")
- Si no puedes completar un campo (ej: no sabes la extensión), indícalo en "cambios" para revisión manual."""

_USER_TEMPLATE = """Revisa y corrige estas referencias/recursos bibliográficos según APA 7 y el formato de 5 campos UST:

{referencias}

Devuelve solo el JSON con las correcciones."""


# ── Funciones de llamada por backend ─────────────────────────────────────────

def _llamar_openai_compat(url: str, api_key: str, model: str,
                           mensajes: list[dict], timeout: int = 30) -> str:
    """Llama a cualquier API compatible con OpenAI chat/completions."""
    payload = json.dumps({
        "model":       model,
        "messages":    mensajes,
        "temperature": 0.1,
        "max_tokens":  2048,
    }).encode("utf-8")

    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {api_key}",
    }

    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read())

    return data["choices"][0]["message"]["content"]


def _llamar_ollama(system: str, user_msg: str, model: str,
                   timeout: int = 120) -> str:
    """Llama a Ollama usando la API nativa /api/chat con think=False."""
    payload = json.dumps({
        "model":   model,
        "think":   False,
        "stream":  False,
        "options": {"num_predict": 2048, "temperature": 0.1},
        "messages": [
            {"role": "system", "content": system},
            {"role": "user",   "content": user_msg},
        ],
    }).encode("utf-8")

    headers = {"Content-Type": "application/json"}
    req = urllib.request.Request(
        "http://localhost:11434/api/chat",
        data=payload, headers=headers, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read())

    return data["message"]["content"]


def _llamar_claude(api_key: str, model: str, mensajes: list[dict],
                   system: str, timeout: int = 30) -> str:
    """Llama a la API nativa de Anthropic (Messages API)."""
    payload = json.dumps({
        "model":      model,
        "max_tokens": 2048,
        "system":     system,
        "messages":   mensajes,
    }).encode("utf-8")

    headers = {
        "Content-Type":      "application/json",
        "x-api-key":         api_key,
        "anthropic-version": "2023-06-01",
    }

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload, headers=headers, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read())

    return data["content"][0]["text"]


# ── Función principal ─────────────────────────────────────────────────────────

def revisar_referencias_llm(
    texto_celda: str,
    backend:  str = "ollama",
    model:    str | None = None,
    api_key:  str = "",
    timeout:  int = 150,
) -> dict:
    """
    Revisa las referencias APA 7 de una celda usando LLM.

    Parámetros
    ----------
    texto_celda : contenido completo de la celda H
    backend     : "ollama" | "claude" | "openai" | "grok"
    model       : override del modelo (None = usa el default del backend)
    api_key     : clave API (no requerida para ollama)
    timeout     : segundos máximo de espera

    Devuelve
    --------
    {
      "referencias": [{"original", "corregida", "cambios", "ok"}, ...],
      "error": str | None,
    }
    """
    # Separar referencias de la celda
    from apa_recursos import separar_referencias
    refs = separar_referencias(texto_celda)
    if not refs:
        return {"referencias": [], "error": None}

    cfg   = BACKEND_DEFAULTS.get(backend, BACKEND_DEFAULTS["ollama"])
    model = model or cfg["model"]
    key   = api_key or cfg["key"]

    # Construir prompt
    texto_refs = "\n".join(f"{i+1}. {r}" for i, r in enumerate(refs))
    user_msg   = _USER_TEMPLATE.format(referencias=texto_refs)

    try:
        if backend == "claude":
            respuesta = _llamar_claude(
                api_key=key, model=model,
                mensajes=[{"role": "user", "content": user_msg}],
                system=_SYSTEM_PROMPT, timeout=timeout,
            )
        elif backend == "ollama":
            # Usa API nativa de Ollama (evita bug de content vacío con qwen3)
            respuesta = _llamar_ollama(
                system=_SYSTEM_PROMPT, user_msg=user_msg,
                model=model, timeout=timeout,
            )
        else:
            # OpenAI, Grok — formato OpenAI-compatible
            url = cfg["url"]
            mensajes = [
                {"role": "system",  "content": _SYSTEM_PROMPT},
                {"role": "user",    "content": user_msg},
            ]
            respuesta = _llamar_openai_compat(
                url=url, api_key=key, model=model,
                mensajes=mensajes, timeout=timeout,
            )

        # Extraer JSON de la respuesta
        json_match = re.search(r'\{[\s\S]*\}', respuesta)
        if not json_match:
            return {"referencias": [], "error": f"LLM no devolvió JSON válido: {respuesta[:200]}"}

        data = json.loads(json_match.group(0))
        return {"referencias": data.get("referencias", []), "error": None}

    except Exception as e:
        return {"referencias": [], "error": str(e)}


# ── Procesar columna H completa con LLM ──────────────────────────────────────

def revisar_columna_recursos_llm(
    ws_plan,
    backend:      str = "ollama",
    model:        str | None = None,
    api_key:      str = "",
    autocorregir: bool = False,
    timeout:      int  = 150,
) -> tuple[list[str], int, int]:
    """
    Revisa la columna H (Recursos) con LLM fila por fila.

    Devuelve (log_lineas, n_problemas, n_correcciones).
    """
    from openpyxl.styles import Font, Color
    from apa_recursos import reconstruir_celda

    COL_H = 8
    AZUL  = "FF2E74B5"

    log: list[str] = []
    n_problemas    = 0
    n_correcciones = 0

    cfg   = BACKEND_DEFAULTS.get(backend, BACKEND_DEFAULTS["ollama"])
    modelo_usado = model or cfg["model"]
    log.append(f'\n  [Verificación APA 7 con LLM — {backend} / {modelo_usado}]')

    for row in ws_plan.iter_rows(min_row=4):
        cell = row[COL_H - 1]
        if not cell.value:
            continue

        resultado = revisar_referencias_llm(
            str(cell.value),
            backend=backend, model=model,
            api_key=api_key, timeout=timeout,
        )

        if resultado["error"]:
            log.append(f'    ⚠️  Fila {cell.row}: error LLM — {resultado["error"][:80]}')
            continue

        refs_data = resultado["referencias"]
        if not refs_data:
            continue

        tiene_problemas = any(not r.get("ok", True) for r in refs_data)
        if tiene_problemas:
            log.append(f'    Fila {cell.row} (H{cell.row}):')

        cambios_fila = []
        refs_corregidas = []

        for r in refs_data:
            ok      = r.get("ok", True)
            cambios = r.get("cambios", [])
            corr    = r.get("corregida", r.get("original", ""))

            if not ok:
                n_problemas += 1
                for c in cambios:
                    log.append(f'      ❌ {c[:120]}')

            elif cambios:
                for c in cambios:
                    log.append(f'      ✏️  {c[:120]}')

            refs_corregidas.append(corr)
            if cambios and corr != r.get("original", ""):
                cambios_fila.extend(cambios)

        # Aplicar correcciones si se solicitó
        if autocorregir and cambios_fila:
            nuevo_texto = reconstruir_celda(refs_corregidas)
            if nuevo_texto != str(cell.value):
                cell.value = nuevo_texto
                try:
                    font_orig = cell.font if cell.font else Font()
                    cell.font = Font(
                        name=font_orig.name or 'Calibri',
                        size=font_orig.size or 11,
                        color=Color(rgb=AZUL),
                    )
                except Exception:
                    pass
                n_correcciones += len(cambios_fila)
                log.append(
                    f'    ✅ Fila {cell.row}: {len(cambios_fila)} corrección(es) aplicada(s)'
                )

    log.append(
        f'\n    APA 7 LLM: {n_problemas} problema(s), '
        f'{n_correcciones} corrección(es) aplicada(s)'
    )
    return log, n_problemas, n_correcciones


# ── Análisis visual de documentos con Qwen2.5-VL ─────────────────────────────

_VISION_SYSTEM = """Eres un asistente experto en diseño instruccional y planificación didáctica para
educación superior a distancia (UST). Analizas imágenes de documentos académicos.
Responde siempre en español. Sé preciso, estructurado y orientado a mejorar la calidad
de los recursos e-learning."""

_VISION_PROMPTS = {
    "planificacion": (
        "Analiza esta planificación didáctica. Identifica:\n"
        "1. Unidades y semanas visibles\n"
        "2. Tipos de recursos (T1 videoclase, T2 Genially, T3 guía, T4 foro/quiz/tarea)\n"
        "3. Momentos de aprendizaje (Preparación, Desarrollo, Trabajo Independiente)\n"
        "4. Posibles inconsistencias o campos vacíos\n"
        "5. Sugerencias de mejora\n\n"
        "Organiza tu respuesta con encabezados claros."
    ),
    "programa": (
        "Analiza este programa de asignatura. Extrae:\n"
        "1. Nombre y código de la asignatura\n"
        "2. Unidades temáticas y sus objetivos\n"
        "3. Resultados de aprendizaje\n"
        "4. Sistema de evaluación y ponderaciones\n"
        "5. Bibliografía relevante\n\n"
        "Estructura la respuesta en secciones."
    ),
    "genially": (
        "Analiza este recurso Genially/infografía educativa. Evalúa:\n"
        "1. Coherencia con estándares UST (paleta verde, tipografía, estructura)\n"
        "2. Claridad del contenido educativo\n"
        "3. Uso adecuado de elementos visuales\n"
        "4. Propuesta de mejoras concretas\n\n"
        "Sé específico en las observaciones."
    ),
    "libre": (
        "Analiza este documento educativo y describe su contenido, "
        "estructura y calidad en el contexto de e-learning universitario."
    ),
}


def analizar_imagen_llm(
    imagen_bytes: bytes,
    tipo_analisis: str = "planificacion",
    prompt_extra: str = "",
    model: str = MODELO_VISION,
    timeout: int = 90,
) -> dict:
    """
    Analiza una imagen de documento con Qwen2.5-VL (Ollama local).

    Parámetros
    ----------
    imagen_bytes   : bytes de la imagen (PNG, JPG, etc.)
    tipo_analisis  : "planificacion" | "programa" | "genially" | "libre"
    prompt_extra   : prompt adicional del usuario (se agrega al template)
    model          : modelo Ollama (por defecto qwen2.5vl:latest)
    timeout        : segundos máximo

    Devuelve
    --------
    {"resultado": str, "error": str | None}
    """
    import base64

    b64 = base64.b64encode(imagen_bytes).decode("utf-8")
    prompt_base = _VISION_PROMPTS.get(tipo_analisis, _VISION_PROMPTS["libre"])
    prompt_final = prompt_base + (f"\n\n{prompt_extra}" if prompt_extra.strip() else "")

    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": _VISION_SYSTEM},
            {
                "role": "user",
                "content": [
                    {"type": "text",      "text": prompt_final},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                ],
            },
        ],
        "temperature": 0.2,
        "max_tokens":  2048,
    }).encode("utf-8")

    headers = {
        "Content-Type":  "application/json",
        "Authorization": "Bearer ollama",
    }

    try:
        req = urllib.request.Request(
            "http://localhost:11434/v1/chat/completions",
            data=payload, headers=headers, method="POST",
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read())
        texto = data["choices"][0]["message"]["content"]
        return {"resultado": texto, "error": None}
    except Exception as e:
        return {"resultado": "", "error": str(e)}
