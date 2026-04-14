"""
db_historial.py — Historial SQLite para Procesador DEL · UST
Registra cada planificación procesada y permite consultar el historial.
"""

import sqlite3
import os
from datetime import datetime

_DB_PATH = os.path.join(os.path.dirname(__file__), "historial_del.db")


def _conectar() -> sqlite3.Connection:
    conn = sqlite3.connect(_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_db():
    """Crea la tabla si no existe."""
    with _conectar() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS historial (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_hora          TEXT    NOT NULL,
                instancia           INTEGER NOT NULL,
                codigo_asignatura   TEXT,
                nombre_asignatura   TEXT,
                archivo_nombre      TEXT,
                total_correcciones  INTEGER DEFAULT 0,
                criterios_ok        INTEGER DEFAULT 0,
                criterios_error     INTEGER DEFAULT 0,
                criterios_manual    INTEGER DEFAULT 0,
                discrepancias_prog  INTEGER DEFAULT 0,
                lt_errores          INTEGER DEFAULT 0,
                lt_correcciones     INTEGER DEFAULT 0,
                tiene_as            INTEGER DEFAULT 0,
                as_ok               INTEGER DEFAULT 0,
                as_error            INTEGER DEFAULT 0,
                bloom_ok            INTEGER DEFAULT 0,
                bloom_debil         INTEGER DEFAULT 0
            )
        """)
        conn.commit()


def registrar(
    instancia: int,
    archivo_nombre: str,
    metricas: dict,
    programa: dict | None = None,
):
    """
    Inserta un registro en el historial.

    Parameters
    ----------
    instancia       : 1, 2 o 3
    archivo_nombre  : nombre del .xlsx procesado
    metricas        : dict devuelto por parsear_log() de app_del.py
    programa        : dict devuelto por extraer_programa_pdf() (puede ser None)
    """
    inicializar_db()
    codigo = (programa or {}).get("codigo") or ""
    nombre = (programa or {}).get("asignatura") or ""

    with _conectar() as conn:
        conn.execute("""
            INSERT INTO historial (
                fecha_hora, instancia, codigo_asignatura, nombre_asignatura,
                archivo_nombre, total_correcciones,
                criterios_ok, criterios_error, criterios_manual,
                discrepancias_prog, lt_errores, lt_correcciones,
                tiene_as, as_ok, as_error,
                bloom_ok, bloom_debil
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            instancia,
            codigo,
            nombre,
            archivo_nombre,
            metricas.get("total_correcciones", 0),
            metricas.get("criterios_ok", 0),
            metricas.get("criterios_error", 0),
            metricas.get("criterios_manual", 0),
            metricas.get("discrepancias_prog", 0),
            metricas.get("lt_errores", 0),
            metricas.get("lt_correcciones", 0),
            int(metricas.get("tiene_as", False)),
            metricas.get("as_ok", 0),
            metricas.get("as_error", 0),
            metricas.get("bloom_ok", 0),
            metricas.get("bloom_debil", 0),
        ))
        conn.commit()


def obtener_historial(limit: int = 200) -> list[dict]:
    """Devuelve los registros más recientes, del más nuevo al más antiguo."""
    inicializar_db()
    with _conectar() as conn:
        rows = conn.execute(
            "SELECT * FROM historial ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def contar_por_codigo(codigo: str) -> int:
    """Cuántas veces se ha procesado una asignatura por código."""
    inicializar_db()
    if not codigo:
        return 0
    with _conectar() as conn:
        row = conn.execute(
            "SELECT COUNT(*) FROM historial WHERE codigo_asignatura = ?",
            (codigo,),
        ).fetchone()
    return row[0] if row else 0


def resumen_errores() -> list[dict]:
    """
    Agrega los tipos de error más frecuentes:
    correcciones totales, errores de criterios y discrepancias.
    """
    inicializar_db()
    with _conectar() as conn:
        rows = conn.execute("""
            SELECT
                codigo_asignatura,
                nombre_asignatura,
                COUNT(*)            AS veces_procesada,
                SUM(total_correcciones)  AS total_corr,
                SUM(criterios_error)     AS total_crit_err,
                SUM(discrepancias_prog)  AS total_disc,
                SUM(lt_errores)          AS total_lt,
                SUM(bloom_ok)            AS total_bloom_ok,
                SUM(bloom_debil)         AS total_bloom_debil
            FROM historial
            GROUP BY codigo_asignatura
            ORDER BY total_corr DESC
        """).fetchall()
    return [dict(r) for r in rows]
