"""
app_del.py — Procesador de Planificaciones DEL · UST 2026-1
Interfaz Streamlit para el equipo de Dirección de Educación a Distancia.

Ejecutar localmente:
    streamlit run app_del.py

Desplegar en Streamlit Cloud:
    1. Subir este archivo + revisar_planificaciones.py + requirements.txt a GitHub
    2. Conectar el repositorio en share.streamlit.io
"""

import streamlit as st
import sys, os, glob, shutil, tempfile, re
from io import BytesIO

# ── Configuración de página ───────────────────────────────────────────────
st.set_page_config(
    page_title="Procesador DEL — UST",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── Cargar script revisor ─────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
try:
    import revisar_planificaciones as rp
    SCRIPT_OK = True
except ImportError:
    SCRIPT_OK = False

# ── Estilos ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { max-width: 780px; padding-top: 2rem; }
    .stButton > button { font-weight: 600; }
    .metric-card {
        background: #f8f9fa; border-radius: 10px;
        padding: 1rem 1.2rem; text-align: center;
    }
    .metric-card .value { font-size: 2rem; font-weight: 700; }
    .metric-card .label { font-size: 0.85rem; color: #666; margin-top: 2px; }
    .tag-ok    { background:#d4edda; color:#155724; border-radius:4px;
                 padding:2px 8px; font-size:0.82rem; }
    .tag-error { background:#f8d7da; color:#721c24; border-radius:4px;
                 padding:2px 8px; font-size:0.82rem; }
    .tag-warn  { background:#fff3cd; color:#856404; border-radius:4px;
                 padding:2px 8px; font-size:0.82rem; }
</style>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════
#  UTILIDADES
# ═════════════════════════════════════════════════════════════════════════

def parsear_log(log: list[str]) -> dict:
    """Extrae métricas clave del log para mostrar en el resumen."""
    resultado = {
        "total_correcciones": 0,
        "criterios_ok": 0, "criterios_error": 0, "criterios_manual": 0,
        "discrepancias_prog": 0,
        "as_ok": 0, "as_error": 0, "as_manual": 0,
        "tiene_as": False,
        "correcciones_detalle": [],
        "discrepancias_detalle": [],
        "as_detalle": [],
    }
    seccion = None
    for linea in log:
        # Total correcciones
        if "TOTAL" in linea and "correcciones" in linea:
            m = re.search(r"TOTAL\s*:\s*(\d+)", linea)
            if m:
                resultado["total_correcciones"] = int(m.group(1))

        # Resultado escala UST
        m = re.search(r"Resultado:\s*(\d+)✅\s*(\d+)❌\s*(\d+)⚠️\s*manual", linea)
        if m:
            resultado["criterios_ok"]     = int(m.group(1))
            resultado["criterios_error"]  = int(m.group(2))
            resultado["criterios_manual"] = int(m.group(3))

        # Discrepancias vs programa
        m = re.search(r"Verificación vs programa:\s*(\d+)", linea)
        if m:
            resultado["discrepancias_prog"] = int(m.group(1))

        # Resultado A+S
        m = re.search(r"Resultado A\+Se:\s*(\d+)✅\s*(\d+)❌\s*(\d+)⚠️\s*manual", linea)
        if m:
            resultado["tiene_as"]  = True
            resultado["as_ok"]     = int(m.group(1))
            resultado["as_error"]  = int(m.group(2))
            resultado["as_manual"] = int(m.group(3))

        # Sección actual para detalles
        if "[Planificación por unidades]" in linea or "[Síntesis didáctica]" in linea:
            seccion = "correcciones"
        elif "[Verificación contra programa" in linea:
            seccion = "programa"
        elif "[Verificación A+Se" in linea:
            seccion = "as"

        # Líneas de detalle
        stripped = linea.strip()
        if seccion == "correcciones" and stripped.startswith("[Plan") and "→" in stripped:
            resultado["correcciones_detalle"].append(stripped)
        elif seccion == "programa" and stripped and stripped[0] in ("✅", "❌", "⚠"):
            resultado["discrepancias_detalle"].append(stripped)
        elif seccion == "as" and stripped and stripped[0] in ("✅", "❌", "⚠"):
            resultado["as_detalle"].append(stripped)

    return resultado


def tag(texto: str, tipo: str) -> str:
    css = {"ok": "tag-ok", "error": "tag-error", "warn": "tag-warn"}.get(tipo, "tag-warn")
    return f'<span class="{css}">{texto}</span>'


# ═════════════════════════════════════════════════════════════════════════
#  CABECERA
# ═════════════════════════════════════════════════════════════════════════

st.markdown("## 📋 Procesador de Planificaciones DEL")
st.caption("Universidad Santo Tomás · Dirección de Educación a Distancia · 2026-1")
st.divider()

if not SCRIPT_OK:
    st.error("No se encontró `revisar_planificaciones.py` en la misma carpeta que esta app.")
    st.stop()

# ═════════════════════════════════════════════════════════════════════════
#  PASO 1 — ARCHIVOS OBLIGATORIOS
# ═════════════════════════════════════════════════════════════════════════

st.markdown("### 1 · Archivos obligatorios")

col_pdf, col_xlsx = st.columns(2)
with col_pdf:
    pdf_file = st.file_uploader(
        "📄 Programa de asignatura",
        type="pdf",
        help="PDF del programa oficial de la asignatura.",
    )
with col_xlsx:
    xlsx_file = st.file_uploader(
        "📊 Planificación didáctica",
        type="xlsx",
        help="Archivo Excel con la planificación por unidades.",
    )

# Mostrar resumen del programa si ya se subió
if pdf_file:
    with st.spinner("Leyendo programa..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
            tmp_pdf.write(pdf_file.read())
            tmp_pdf_path = tmp_pdf.name
        pdf_file.seek(0)
        programa = rp.extraer_programa_pdf(tmp_pdf_path)
        os.unlink(tmp_pdf_path)

    if programa.get("_error"):
        st.warning(f"No se pudo leer el PDF: {programa['_error']}")
        programa = {}
    else:
        with st.expander("✅ Programa leído — verificar datos extraídos", expanded=False):
            c1, c2, c3 = st.columns(3)
            c1.metric("Código", programa.get("codigo") or "—")
            c2.metric("Créditos", programa.get("creditos") or "—")
            c3.metric("Área", programa.get("area") or "—")
            unidades = programa.get("unidades", [])
            if unidades:
                st.markdown("**Unidades y ponderaciones:**")
                ponds = programa.get("ponderaciones", {})
                for u in unidades:
                    pct = ponds.get(u["numero"], "—")
                    st.markdown(
                        f"- Unidad {u['numero']}: **{u['nombre']}** "
                        f"· {u['horas']}h pedagógicas · ponderación: {pct}%"
                    )
else:
    programa = {}

# ═════════════════════════════════════════════════════════════════════════
#  PASO 2 — OPCIONES
# ═════════════════════════════════════════════════════════════════════════

st.markdown("### 2 · Opciones")
st.caption("Solo marcar si aplica a esta asignatura.")

col_as, col_dec = st.columns(2)
with col_as:
    es_as = st.checkbox(
        "📌 Lineamiento A+S",
        help="Activa la verificación de los 11 hitos obligatorios "
             "del lineamiento Aprendizaje + Servicio (UST 2025).",
    )
with col_dec:
    st.checkbox(
        "📜 Decreto de actualización",
        disabled=True,
        help="Próximamente: permite subir un decreto que actualice "
             "campos del programa oficial.",
    )

# ═════════════════════════════════════════════════════════════════════════
#  PASO 3 — PROCESAR
# ═════════════════════════════════════════════════════════════════════════

st.markdown("### 3 · Revisar")

listo = bool(pdf_file and xlsx_file)
procesar = st.button(
    "▶ Revisar planificación",
    disabled=not listo,
    type="primary",
    use_container_width=True,
    help="Sube el PDF y el .xlsx para activar este botón." if not listo else "",
)

if not listo:
    st.caption(
        "⬆ Sube el PDF del programa y la planificación .xlsx para continuar."
    )

# ═════════════════════════════════════════════════════════════════════════
#  PROCESAMIENTO Y RESULTADOS
# ═════════════════════════════════════════════════════════════════════════

if procesar and listo:
    output_bytes = None
    output_name  = None
    log          = []
    ok           = False

    with st.spinner("Procesando planificación..."):
        with tempfile.TemporaryDirectory() as tmp:
            carpeta_asig = os.path.join(tmp, "asig")
            carpeta_env  = os.path.join(carpeta_asig, "Enviado a DEL")
            os.makedirs(carpeta_env)

            # Guardar xlsx
            xlsx_path = os.path.join(carpeta_env, xlsx_file.name)
            with open(xlsx_path, "wb") as f:
                f.write(xlsx_file.getvalue())

            # Guardar PDF y extraer programa si aún no se hizo
            if not programa:
                pdf_path = os.path.join(tmp, pdf_file.name)
                with open(pdf_path, "wb") as f:
                    f.write(pdf_file.getvalue())
                programa = rp.extraer_programa_pdf(pdf_path)

            # Procesar
            log, ok = rp.procesar_asignatura(
                carpeta_asig,
                programa=programa,
                es_as=es_as,
            )

            if ok:
                salidas = glob.glob(os.path.join(carpeta_asig, "Revisado", "*.xlsx"))
                if salidas:
                    with open(salidas[0], "rb") as f:
                        output_bytes = f.read()
                    output_name = os.path.basename(salidas[0])

    # ── Resultados ────────────────────────────────────────────────────────
    st.divider()
    st.markdown("### Resultados")

    if not ok or not output_bytes:
        st.error("El procesamiento falló. Revisa el log.")
    else:
        metricas = parsear_log(log)

        # Métricas principales
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(
                "Correcciones aplicadas",
                metricas["total_correcciones"],
                help="Celdas modificadas automáticamente (marcadas en azul en el archivo).",
            )
        with col2:
            total_criterios = (metricas["criterios_ok"] + metricas["criterios_error"]
                               + metricas["criterios_manual"])
            st.metric(
                "Criterios UST",
                f"{metricas['criterios_ok']}/{total_criterios}",
                delta=f"{metricas['criterios_error']} con error"
                      if metricas["criterios_error"] else "sin errores",
                delta_color="inverse" if metricas["criterios_error"] else "normal",
            )
        with col3:
            st.metric(
                "Discrepancias vs programa",
                metricas["discrepancias_prog"],
                delta="revisar" if metricas["discrepancias_prog"] else "todo coincide",
                delta_color="inverse" if metricas["discrepancias_prog"] else "normal",
            )

        # A+S métricas adicionales
        if metricas["tiene_as"]:
            st.markdown("**Verificación A+Se:**")
            total_as = metricas["as_ok"] + metricas["as_error"] + metricas["as_manual"]
            st.progress(
                metricas["as_ok"] / total_as if total_as else 0,
                text=f"{metricas['as_ok']}✅  {metricas['as_error']}❌  "
                     f"{metricas['as_manual']}⚠️ manual  (de {total_as} hitos)",
            )

        # Discrepancias vs programa (si hay)
        if metricas["discrepancias_detalle"]:
            with st.expander("📊 Detalle — verificación contra programa", expanded=True):
                for linea in metricas["discrepancias_detalle"]:
                    st.markdown(linea)

        # Hitos A+S (si aplica)
        if metricas["tiene_as"] and metricas["as_detalle"]:
            with st.expander("📌 Detalle — hitos A+Se", expanded=metricas["as_error"] > 0):
                for linea in metricas["as_detalle"]:
                    st.markdown(linea)

        # Botón de descarga
        st.divider()
        st.download_button(
            label=f"⬇ Descargar {output_name}",
            data=output_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption(
            "El archivo incluye todas las correcciones marcadas en azul. "
            "Las filas Formativa tienen fondo lila y las Sumativa fondo amarillo."
        )

        # Log completo
        with st.expander("🔍 Ver log completo de correcciones"):
            log_texto = "\n".join(log)
            st.code(log_texto, language=None)
            st.download_button(
                "⬇ Descargar log (.txt)",
                data=log_texto,
                file_name=output_name.replace(".xlsx", "_log.txt"),
                mime="text/plain",
            )
