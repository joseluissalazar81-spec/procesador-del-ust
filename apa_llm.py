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

# ── Configuración de backends ────────────────────────────────────────────────

BACKENDS = ["ollama", "claude", "openai", "grok"]

BACKEND_DEFAULTS = {
    "ollama": {
        "url":   "http://localhost:11434/v1/chat/completions",
        "model": "gemma3:12b",
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

_SYSTEM_PROMPT = """Eres un experto en norma APA 7 para documentos académicos en español.
Tu tarea es revisar y corregir referencias bibliográficas según APA 7.

Reglas que debes aplicar:
1. Formato de autor: Apellido, I. I. (iniciales con punto). Varios autores separados con coma y & antes del último.
2. Año entre paréntesis seguido de punto: (2024).
3. Título del artículo/recurso en redonda (no cursiva), mayúscula solo en primera palabra y nombres propios.
4. Nombre de la revista/libro en cursiva (indica con *cursiva*).
5. Volumen en cursiva, número de issue entre paréntesis sin cursiva.
6. DOI con formato: https://doi.org/...
7. URLs directas sin "Disponible en" precedente.
8. Sin punto después de URL.
9. Usar & (no "y") entre coautores.
10. Recursos digitales incluyen tipo entre corchetes: [Video], [Diapositivas], [Material de estudio], etc.

RESPONDE ÚNICAMENTE con un JSON con esta estructura exacta, sin texto adicional:
{
  "referencias": [
    {
      "original": "texto original de la referencia",
      "corregida": "texto corregido en APA 7",
      "cambios": ["descripción cambio 1", "descripción cambio 2"],
      "ok": true
    }
  ]
}

Si una referencia ya está correcta, pon ok=true y corregida igual al original.
Si no puedes corregirla, pon ok=false y explica en cambios."""

_USER_TEMPLATE = """Revisa y corrige estas referencias bibliográficas en APA 7:

{referencias}

Devuelve el JSON con las correcciones."""


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
    timeout:  int = 45,
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
        else:
            # Ollama, OpenAI, Grok — todos usan formato OpenAI-compatible
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
    timeout:      int  = 45,
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
