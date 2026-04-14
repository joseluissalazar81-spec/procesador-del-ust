"""
dict_ust.py — Diccionario UST personalizable · Procesador DEL
Gestiona correcciones personalizadas que se mezclan con los mapas base.

El archivo dict_ust.json se crea automáticamente junto a este módulo.
Formato JSON:
{
  "PROC_MAP":  { "término incorrecto": "término correcto", ... },
  "INSTR_MAP": { ... },
  "TIPO_MAP":  { ... }
}
"""

from __future__ import annotations
import json
import os

_JSON_PATH = os.path.join(os.path.dirname(__file__), "dict_ust.json")

# Mapas base integrados en revisar_planificaciones.py
_BASE = {
    "PROC_MAP": {
        "Cuestionario": "Pruebas escritas u orales",
        "Tarea":        "Producciones del estudiante",
        "Prueba":       "Pruebas escritas u orales",
    },
    "INSTR_MAP": {
        "Pauta":   "Pauta de observación",
        "Rúbrica": "Pauta de observación",
    },
    "TIPO_MAP": {
        "Formativo":   "Formativa",
        "Diagnóstico": "Diagnóstica",
    },
}

_ETIQUETAS = {
    "PROC_MAP":  "Procedimiento de evaluación",
    "INSTR_MAP": "Instrumento de evaluación",
    "TIPO_MAP":  "Tipo de evaluación",
}


# ── Lectura / escritura del JSON ───────────────────────────────────────────

def _cargar_json() -> dict:
    if not os.path.exists(_JSON_PATH):
        return {"PROC_MAP": {}, "INSTR_MAP": {}, "TIPO_MAP": {}}
    try:
        with open(_JSON_PATH, encoding="utf-8") as f:
            data = json.load(f)
        for key in ("PROC_MAP", "INSTR_MAP", "TIPO_MAP"):
            data.setdefault(key, {})
        return data
    except Exception:
        return {"PROC_MAP": {}, "INSTR_MAP": {}, "TIPO_MAP": {}}


def _guardar_json(data: dict):
    with open(_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── API pública ────────────────────────────────────────────────────────────

def obtener_dict_completo() -> dict:
    """
    Devuelve el diccionario combinado: base + personalizaciones.
    Las entradas personalizadas sobreescriben las base con el mismo término.
    """
    custom = _cargar_json()
    resultado = {}
    for key in ("PROC_MAP", "INSTR_MAP", "TIPO_MAP"):
        resultado[key] = {**_BASE[key], **custom[key]}
    return resultado


def obtener_entradas_custom() -> dict:
    """Solo las entradas personalizadas (sin la base)."""
    return _cargar_json()


def agregar_entrada(mapa: str, incorrecto: str, correcto: str):
    """Agrega o actualiza una entrada en el diccionario personalizado."""
    if mapa not in ("PROC_MAP", "INSTR_MAP", "TIPO_MAP"):
        raise ValueError(f"Mapa desconocido: {mapa}")
    data = _cargar_json()
    data[mapa][incorrecto.strip()] = correcto.strip()
    _guardar_json(data)


def eliminar_entrada(mapa: str, incorrecto: str):
    """Elimina una entrada del diccionario personalizado (no afecta la base)."""
    data = _cargar_json()
    data[mapa].pop(incorrecto.strip(), None)
    _guardar_json(data)


def exportar_json() -> bytes:
    """Devuelve el diccionario personalizado como bytes JSON (para descarga)."""
    data = _cargar_json()
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


def importar_json(contenido: bytes) -> int:
    """
    Importa un JSON exportado previamente.
    Hace merge: las nuevas entradas se suman a las existentes.
    Devuelve el número de entradas importadas.
    """
    nuevo = json.loads(contenido.decode("utf-8"))
    actual = _cargar_json()
    total = 0
    for key in ("PROC_MAP", "INSTR_MAP", "TIPO_MAP"):
        for k, v in nuevo.get(key, {}).items():
            actual[key][k] = v
            total += 1
    _guardar_json(actual)
    return total


def etiqueta(mapa: str) -> str:
    return _ETIQUETAS.get(mapa, mapa)


MAPAS = list(_ETIQUETAS.keys())
