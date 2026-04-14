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
import streamlit.components.v1 as components
import sys, os, glob, shutil, tempfile, re
from io import BytesIO
import db_historial as hist
import validar_planificacion as vp
import dict_ust
import apa_recursos as apa
import apa_llm

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

# ── Estilos institucionales UST ───────────────────────────────────────────
st.markdown("""
<style>
    /* ── Paleta UST ── */
    /* Verde institucional : #006633  */
    /* Verde claro pastel  : #E8F5EE  */
    /* Verde medio         : #C8E6D4  */
    /* Acento dorado       : #C8A951  */

    /* ── Fondo general ── */
    .stApp {
        background-color: #F0F7F4;
    }

    /* ── Contenedor principal ── */
    .block-container {
        max-width: 820px;
        padding-top: 1.8rem;
        background-color: #FFFFFF;
        border-radius: 12px;
        padding-left: 2.5rem;
        padding-right: 2.5rem;
        box-shadow: 0 2px 12px rgba(0,102,51,0.08);
    }

    /* ── Cabecera de la app ── */
    .del-header {
        background: linear-gradient(135deg, #006633 0%, #009450 100%);
        color: #FFFFFF;
        border-radius: 10px;
        padding: 1.2rem 1.6rem 1rem;
        margin-bottom: 1.2rem;
    }
    .del-header h2 {
        color: #FFFFFF !important;
        margin: 0 0 0.15rem 0;
        font-size: 1.45rem;
        font-weight: 700;
        letter-spacing: 0.01em;
    }
    .del-header .sub {
        color: #C8E6D4;
        font-size: 0.82rem;
        margin: 0;
    }
    .del-header .badge {
        display: inline-block;
        background: rgba(255,255,255,0.18);
        border: 1px solid rgba(255,255,255,0.35);
        border-radius: 20px;
        padding: 1px 10px;
        font-size: 0.72rem;
        color: #FFFFFF;
        margin-top: 0.5rem;
    }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background-color: #E8F5EE;
        border-radius: 8px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px;
        padding: 0.4rem 1rem;
        font-weight: 500;
        color: #004d26;
        background-color: transparent;
    }
    .stTabs [aria-selected="true"] {
        background-color: #006633 !important;
        color: #FFFFFF !important;
        font-weight: 600;
    }

    /* ── Botón primario ── */
    .stButton > button[kind="primary"],
    div[data-testid="stDownloadButton"] button[kind="primary"] {
        background-color: #006633;
        border: none;
        color: #FFFFFF;
        font-weight: 600;
        border-radius: 8px;
        transition: background-color 0.2s;
    }
    .stButton > button[kind="primary"]:hover,
    div[data-testid="stDownloadButton"] button[kind="primary"]:hover {
        background-color: #004d26;
    }

    /* ── Botones secundarios ── */
    .stButton > button:not([kind="primary"]) {
        border: 1.5px solid #006633;
        color: #006633;
        border-radius: 8px;
        font-weight: 500;
    }

    /* ── Métricas ── */
    div[data-testid="stMetric"] {
        background-color: #F0F7F4;
        border: 1px solid #C8E6D4;
        border-radius: 10px;
        padding: 0.7rem 1rem;
    }
    div[data-testid="stMetricLabel"] {
        color: #004d26 !important;
        font-weight: 600;
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    div[data-testid="stMetricValue"] {
        color: #006633 !important;
        font-weight: 700;
    }

    /* ── Expanders ── */
    .streamlit-expanderHeader {
        background-color: #F0F7F4;
        border: 1px solid #C8E6D4;
        border-radius: 8px;
        color: #004d26;
        font-weight: 500;
    }

    /* ── Dividers ── */
    hr {
        border-color: #C8E6D4 !important;
    }

    /* ── Info / success banners ── */
    div[data-testid="stAlert"] {
        border-radius: 8px;
    }

    /* ── File uploader ── */
    div[data-testid="stFileUploader"] {
        border: 1.5px dashed #009450;
        border-radius: 10px;
        padding: 0.5rem;
        background-color: #F7FCF9;
    }

    /* ── Tags ── */
    .tag-ok    { background:#C8E6D4; color:#004d26; border-radius:4px;
                 padding:2px 9px; font-size:0.82rem; font-weight:500; }
    .tag-error { background:#FDDCDC; color:#8B0000; border-radius:4px;
                 padding:2px 9px; font-size:0.82rem; font-weight:500; }
    .tag-warn  { background:#FFF3CD; color:#7A5700; border-radius:4px;
                 padding:2px 9px; font-size:0.82rem; font-weight:500; }

    /* ── Captions ── */
    .stCaption { color: #4a7c5e; }

    /* ── Checkbox ── */
    .stCheckbox label { color: #004d26; font-weight: 500; }

    /* colores de botones: aplicados via JS (components.html) */
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
        "lt_errores": 0,
        "lt_revisadas": 0,
        "lt_correcciones": 0,
        "lt_detalle": [],
        "lt_ejecutado": False,
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

        # LanguageTool — resumen (nuevo formato con correcciones)
        m = re.search(r"LanguageTool:\s*(\d+)\s*celda\(s\)\s*revisadas?,\s*(\d+)\s*error\(es?\)(?:,\s*(\d+)\s*correcci.+?\s*aplicada\(s\))?", linea)
        if m:
            resultado["lt_revisadas"] = int(m.group(1))
            resultado["lt_errores"]   = int(m.group(2))
            resultado["lt_correcciones"] = int(m.group(3)) if m.group(3) else 0
            resultado["lt_ejecutado"] = True

        # Sección actual para detalles
        if "[Planificación por unidades]" in linea or "[Síntesis didáctica]" in linea:
            seccion = "correcciones"
        elif "[Verificación contra programa" in linea:
            seccion = "programa"
        elif "[Verificación A+Se" in linea:
            seccion = "as"
        elif "LanguageTool es" in linea or "autocorrección" in linea:
            seccion = "lt"

        # Líneas de detalle
        stripped = linea.strip()
        if seccion == "correcciones" and stripped.startswith("[Plan") and "→" in stripped:
            resultado["correcciones_detalle"].append(stripped)
        elif seccion == "programa" and stripped and stripped[0] in ("✅", "❌", "⚠"):
            resultado["discrepancias_detalle"].append(stripped)
        elif seccion == "as" and stripped and stripped[0] in ("✅", "❌", "⚠"):
            resultado["as_detalle"].append(stripped)
        elif seccion == "lt" and stripped and stripped[0] in ("✅", "⚠", "ℹ"):
            resultado["lt_detalle"].append(stripped)

    return resultado


def _datos_desde_xlsx(xlsx_bytes: bytes) -> dict:
    """Lee código y nombre de asignatura desde la hoja Síntesis didáctica del Excel."""
    try:
        import openpyxl as _oxl
        wb = _oxl.load_workbook(BytesIO(xlsx_bytes), read_only=True, data_only=True)
        ws = next(
            (wb[s] for s in wb.sheetnames
             if 'síntesis' in s.lower() or 'sintesis' in s.lower()),
            None
        )
        if ws is None:
            return {}
        codigo     = str(ws.cell(7, 1).value or "").strip()
        asignatura = str(ws.cell(4, 4).value or "").strip()
        wb.close()
        return {"codigo": codigo or None, "asignatura": asignatura or None}
    except Exception:
        return {}


def _nombre_descarga(programa: dict | None, instancia: int,
                     xlsx_bytes: bytes | None = None) -> str:
    """
    Genera el nombre del archivo de descarga:
      {CODIGO}_{Nombre_Asignatura}_Revisado_I{N}.xlsx
    Primero lee del dict programa (PDF); si falta algo usa el Excel como fallback.
    """
    prog = dict(programa or {})
    # Fallback al Excel si el programa no tiene código o nombre
    if xlsx_bytes and (not prog.get("codigo") or not prog.get("asignatura")):
        xlsx_datos = _datos_desde_xlsx(xlsx_bytes)
        prog.setdefault("codigo",     xlsx_datos.get("codigo"))
        prog.setdefault("asignatura", xlsx_datos.get("asignatura"))

    codigo = (prog.get("codigo") or "").strip()
    nombre = (prog.get("asignatura") or "").strip()
    nombre_limpio = re.sub(r'[\\/:*?"<>|]', '', nombre).strip()
    nombre_limpio = re.sub(r'\s+', '_', nombre_limpio)
    partes = [p for p in [codigo, nombre_limpio] if p]
    base   = "_".join(partes) if partes else "Planificacion"
    return f"{base}_Revisado_I{instancia}.xlsx"


def tag(texto: str, tipo: str) -> str:
    css = {"ok": "tag-ok", "error": "tag-error", "warn": "tag-warn"}.get(tipo, "tag-warn")
    return f'<span class="{css}">{texto}</span>'


# ═════════════════════════════════════════════════════════════════════════
#  CABECERA
# ═════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="del-header">
  <h2>📋 Procesador de Planificaciones DEL</h2>
  <p class="sub">Universidad Santo Tomás · Dirección de Educación a Distancia</p>
  <span class="badge">Semestre 2026-1</span>
</div>
""", unsafe_allow_html=True)

# ── Colores de botones vía JS (CSS no puede atravesar el DOM de Streamlit) ──
components.html("""
<script>
(function () {
    var COLOR_MAP = [
        { label: 'Revisar planificación',             bg: '#4A9068', hov: '#357a52' },
        { label: 'Aplicar correcciones (Instancia 2)',bg: '#4A7FA5', hov: '#356488' },
        { label: 'Aplicar correcciones (Instancia 3)',bg: '#8B6EAF', hov: '#70578f' },
        { label: 'Exportar historial',                bg: '#C07A3A', hov: '#a0622c' },
        { label: 'Agregar al diccionario',            bg: '#5A7A8A', hov: '#456070' },
        { label: 'Nueva',                             bg: '#6c757d', hov: '#545b62' },
    ];
    var done = new WeakSet();

    function paint() {
        try {
            var doc = window.parent.document;
            doc.querySelectorAll('button').forEach(function (btn) {
                if (done.has(btn)) return;
                var txt = (btn.innerText || btn.textContent || '').trim();
                COLOR_MAP.forEach(function (c) {
                    if (txt.indexOf(c.label) !== -1) {
                        btn.style.setProperty('background-color', c.bg, 'important');
                        btn.style.setProperty('border-color',     c.bg, 'important');
                        btn.style.setProperty('color',            '#ffffff', 'important');
                        btn.addEventListener('mouseenter', function () {
                            btn.style.setProperty('background-color', c.hov, 'important');
                        });
                        btn.addEventListener('mouseleave', function () {
                            btn.style.setProperty('background-color', c.bg, 'important');
                        });
                        done.add(btn);
                    }
                });
            });
        } catch (e) { /* cross-origin bloqueado — silencioso */ }
    }

    paint();
    new MutationObserver(paint).observe(
        window.parent.document.body,
        { childList: true, subtree: true }
    );
})();
</script>
""", height=0, scrolling=False)

if not SCRIPT_OK:
    st.error("No se encontró `revisar_planificaciones.py` en la misma carpeta que esta app.")
    st.stop()

# ── Botón de limpieza ─────────────────────────────────────────────────────
_col_sp, _col_btn = st.columns([6, 1])
with _col_btn:
    if st.button("🔄 Nueva", help="Limpia todos los archivos y resultados para procesar otra planificación", use_container_width=True):
        st.session_state.clear()
        st.rerun()

# ═════════════════════════════════════════════════════════════════════════
#  SELECTOR DE INSTANCIA
# ═════════════════════════════════════════════════════════════════════════

tab_i1, tab_i2, tab_i3, tab_hist, tab_dict = st.tabs([
    "Instancia 1 — Revisión previa al envío",
    "Instancia 2 — 1ª revisora DEL",
    "Instancia 3 — 2ª revisora / aprobación final",
    "📊 Historial",
    "📖 Diccionario UST",
])

# ═════════════════════════════════════════════════════════════════════════
#  INSTANCIA 1
# ═════════════════════════════════════════════════════════════════════════

with tab_i1:

    st.markdown("### 1 · Archivos obligatorios")

    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        pdf_file = st.file_uploader(
            "📄 Programa de asignatura",
            type="pdf",
            key="i1_pdf",
            help="PDF del programa oficial de la asignatura.",
        )
    with col_xlsx:
        xlsx_file = st.file_uploader(
            "📊 Planificación didáctica",
            type="xlsx",
            key="i1_xlsx",
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

    # ── Validaciones preventivas ──────────────────────────────────────────
    if xlsx_file:
        problemas = vp.validar_xlsx(xlsx_file.getvalue())
        xlsx_file.seek(0)
        errores   = [p for p in problemas if p["nivel"] == "error"]
        advertencias = [p for p in problemas if p["nivel"] == "advertencia"]
        if errores or advertencias:
            with st.expander(
                f"{'🔴' if errores else '🟡'} Validación previa — "
                f"{len(errores)} error(es), {len(advertencias)} advertencia(s)",
                expanded=bool(errores),
            ):
                for p in errores:
                    st.error(f"**[{p['codigo']}]** {p['mensaje']}", icon="🔴")
                for p in advertencias:
                    st.warning(f"**[{p['codigo']}]** {p['mensaje']}", icon="⚠️")
                if errores:
                    st.caption(
                        "Corrige los errores marcados en rojo antes de procesar."
                    )
        else:
            st.success("Archivo validado sin problemas estructurales.", icon="✅")

    st.markdown("### 2 · Opciones")
    st.caption("Solo marcar si aplica a esta asignatura.")

    col_as, col_dec = st.columns(2)
    with col_as:
        es_as = st.checkbox(
            "📌 Lineamiento A+S",
            key="i1_as",
            help="Activa la verificación de los 11 hitos obligatorios "
                 "del lineamiento Aprendizaje + Servicio (UST 2025).",
        )
    with col_dec:
        st.checkbox(
            "📜 Decreto de actualización",
            disabled=True,
            key="i1_dec",
            help="Próximamente: permite subir un decreto que actualice "
                 "campos del programa oficial.",
        )

    usar_lt_i1 = st.checkbox(
        "🔤 Revisión ortográfica y gramatical (LanguageTool)",
        value=True,
        key="i1_lt",
        help="Revisa ortografía y gramática de las celdas de actividad "
             "usando la API gratuita de LanguageTool en español. "
             "Requiere conexión a internet. Aumenta el tiempo de procesamiento.",
    )

    autocorregir_lt_i1 = st.checkbox(
        "✍️ Aplicar correcciones automáticas (conservador)",
        value=True,
        key="i1_autocorr",
        help="Aplica automáticamente las correcciones SEGURAS detectadas por LanguageTool "
             "(solo cambios unívocos: tildes faltantes, errores ortográficos claros). "
             "Los cambios se marcan en azul en el archivo descargado.",
    )

    st.markdown("**Referencias bibliográficas (columna H)**")
    col_apa1, col_apa2 = st.columns(2)
    with col_apa1:
        usar_apa_i1 = st.checkbox(
            "📚 Validar referencias APA 7 (reglas)",
            value=True,
            key="i1_apa",
            help="Verifica que las referencias de la columna H cumplan APA 7: "
                 "formato de autor, año entre paréntesis, & entre autores, "
                 "URL sin 'Disponible en', etc.",
        )
    with col_apa2:
        autocorr_apa_i1 = st.checkbox(
            "✍️ Corregir APA automáticamente",
            value=False,
            key="i1_apa_corr",
            help="Aplica correcciones APA 7 seguras: elimina 'Disponible en', "
                 "'y' → '&' entre autores, punto tras año, etc. "
                 "Los cambios se marcan en azul.",
        )

    # ── Revisión APA con LLM ──────────────────────────────────────────────
    with st.expander("🤖 Revisión APA 7 con LLM (más precisa)", expanded=False):
        st.caption(
            "Usa un modelo de lenguaje para revisar y corregir referencias APA 7 "
            "con mayor profundidad que las reglas automáticas. "
            "Ollama es gratuito y corre localmente."
        )

        col_llm1, col_llm2 = st.columns([2, 2])
        with col_llm1:
            usar_llm_i1 = st.checkbox(
                "Activar revisión LLM",
                value=False,
                key="i1_llm_activo",
            )
            backend_i1 = st.selectbox(
                "Backend",
                apa_llm.BACKENDS,
                key="i1_llm_backend",
                help="Ollama = local y gratis. Los demás requieren API key.",
            )
        with col_llm2:
            # Modelo por defecto según backend
            modelo_default = apa_llm.BACKEND_DEFAULTS[backend_i1]["model"]
            modelo_i1 = st.text_input(
                "Modelo",
                value=modelo_default,
                key="i1_llm_model",
                help="Para Ollama: nombre del modelo instalado (ej: gemma3:12b). "
                     "Para otros backends: nombre del modelo de la API.",
            )
            apikey_i1 = st.text_input(
                "API Key",
                value="",
                type="password",
                key="i1_llm_key",
                placeholder="No requerida para Ollama",
                help="Pega aquí tu API key si usas Claude, OpenAI o Grok.",
            )
        autocorr_llm_i1 = st.checkbox(
            "Aplicar correcciones LLM al archivo",
            value=False,
            key="i1_llm_autocorr",
            help="Si está activo, las correcciones del LLM se escriben en la "
                 "columna H del archivo descargado (marcadas en azul).",
        )
        if backend_i1 == "ollama":
            st.info(
                f"Modelo por defecto: **{modelo_default}**. "
                "Si no lo tienes instalado, ejecuta en terminal: "
                f"`ollama pull {modelo_default}`",
                icon="💡",
            )

    st.markdown("### 3 · Revisar")

    listo_i1 = bool(pdf_file and xlsx_file)
    procesar_i1 = st.button(
        "▶ Revisar planificación",
        disabled=not listo_i1,
        type="primary",
        key="btn_i1",
        use_container_width=True,
        help="Sube el PDF y el .xlsx para activar este botón." if not listo_i1 else "",
    )

    if not listo_i1:
        st.caption("⬆ Sube el PDF del programa y la planificación .xlsx para continuar.")

    if procesar_i1 and listo_i1:
        output_bytes = None
        output_name  = None
        log          = []
        ok           = False

        rp._LANGUAGETOOL_ACTIVO = usar_lt_i1
        rp._LANGUAGETOOL_AUTOCORREGIR = autocorregir_lt_i1

        spinner_msg = ("Procesando planificación + revisión ortográfica (puede tardar)..."
                       if usar_lt_i1 else "Procesando planificación...")
        with st.spinner(spinner_msg):
            with tempfile.TemporaryDirectory() as tmp:
                carpeta_asig = os.path.join(tmp, "asig")
                carpeta_env  = os.path.join(carpeta_asig, "Enviado a DEL")
                os.makedirs(carpeta_env)

                xlsx_path = os.path.join(carpeta_env, xlsx_file.name)
                with open(xlsx_path, "wb") as f:
                    f.write(xlsx_file.getvalue())

                if not programa:
                    pdf_path = os.path.join(tmp, pdf_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_file.getvalue())
                    programa = rp.extraer_programa_pdf(pdf_path)

                log, ok = rp.procesar_asignatura(
                    carpeta_asig,
                    programa=programa,
                    es_as=es_as,
                )

                if ok:
                    salidas = glob.glob(os.path.join(carpeta_asig, "Revisado", "*.xlsx"))
                    if salidas:
                        # ── APA 7 en columna H ──────────────────────────
                        if usar_apa_i1:
                            import openpyxl as _oxl
                            _wb_apa = _oxl.load_workbook(salidas[0])
                            if "Planificación por unidades" in _wb_apa.sheetnames:
                                _ws_apa = _wb_apa["Planificación por unidades"]
                                _apa_log, _apa_probs, _apa_corr = \
                                    apa.revisar_columna_recursos(
                                        _ws_apa,
                                        autocorregir=autocorr_apa_i1,
                                    )
                                log.extend(_apa_log)
                                if autocorr_apa_i1 and _apa_corr:
                                    _wb_apa.save(salidas[0])
                            _wb_apa.close()
                        # ── APA con LLM ──────────────────────────────
                        if usar_llm_i1:
                            import openpyxl as _oxl2
                            _wb_llm = _oxl2.load_workbook(salidas[0])
                            if "Planificación por unidades" in _wb_llm.sheetnames:
                                _ws_llm = _wb_llm["Planificación por unidades"]
                                _llm_log, _llm_probs, _llm_corr = \
                                    apa_llm.revisar_columna_recursos_llm(
                                        _ws_llm,
                                        backend=backend_i1,
                                        model=modelo_i1 or None,
                                        api_key=apikey_i1,
                                        autocorregir=autocorr_llm_i1,
                                    )
                                log.extend(_llm_log)
                                if autocorr_llm_i1 and _llm_corr:
                                    _wb_llm.save(salidas[0])
                            _wb_llm.close()
                        # ────────────────────────────────────────────────
                        with open(salidas[0], "rb") as f:
                            output_bytes = f.read()
                        output_name = _nombre_descarga(programa, 1, output_bytes)

        st.divider()
        st.markdown("### Resultados")

        if not ok or not output_bytes:
            st.error("El procesamiento falló. Revisa el log.")
        else:
            metricas = parsear_log(log)
            hist.registrar(instancia=1, archivo_nombre=output_name,
                           metricas=metricas, programa=programa or None)
            veces = hist.contar_por_codigo((programa or {}).get("codigo", ""))
            if veces >= 3:
                st.warning(
                    f"⚠️ Esta asignatura ya fue procesada **{veces} veces**. "
                    "Verifica si corresponde a una nueva versión.",
                    icon="🔁",
                )

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(
                    "Correcciones aplicadas",
                    metricas["total_correcciones"],
                    help="Celdas modificadas automáticamente (marcadas en azul).",
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

            # ── Métricas APA ────────────────────────────────────────────
            if usar_apa_i1:
                _apa_probs_log = sum(
                    1 for l in log
                    if "[APA_" in l and ("❌" in l or "⚠️" in l)
                )
                _apa_corr_log = sum(
                    1 for l in log if l.strip().startswith("✏️") and "Fila" in l
                )
                col_a1, col_a2 = st.columns(2)
                with col_a1:
                    st.metric(
                        "Problemas APA 7 (col. H)",
                        _apa_probs_log,
                        delta="revisar referencias" if _apa_probs_log else "referencias OK",
                        delta_color="inverse" if _apa_probs_log else "normal",
                        help="Errores y advertencias detectados en columna H.",
                    )
                with col_a2:
                    st.metric(
                        "Correcciones APA aplicadas",
                        _apa_corr_log,
                        help="Cambios APA 7 seguros aplicados en columna H.",
                    )
                # Detalle APA en expander
                _apa_detalle = [
                    l for l in log
                    if "APA" in l or ("Fila" in l and "H" in l and "✏️" in l)
                ]
                if _apa_detalle:
                    with st.expander(
                        f"📚 Detalle APA 7 — columna H ({_apa_probs_log} problema(s))",
                        expanded=_apa_probs_log > 0,
                    ):
                        for linea in _apa_detalle:
                            st.markdown(linea)

            # ── Métricas LLM ─────────────────────────────────────────────
            if usar_llm_i1:
                _llm_probs_log = sum(
                    1 for l in log if "APA 7 LLM" not in l
                    and "❌" in l and "LLM" not in l
                    and "[APA_" not in l
                )
                _llm_corr_log = sum(
                    1 for l in log
                    if l.strip().startswith("✅") and "corrección(es) aplicada" in l
                )
                # Leer directamente del log LLM
                for l in reversed(log):
                    m = re.search(r'APA 7 LLM: (\d+) problema', l)
                    if m:
                        _llm_probs_log = int(m.group(1))
                        break
                for l in reversed(log):
                    m = re.search(r'APA 7 LLM:.*?(\d+) corrección', l)
                    if m:
                        _llm_corr_log = int(m.group(1))
                        break

                col_l1, col_l2 = st.columns(2)
                with col_l1:
                    st.metric(
                        f"Problemas APA — LLM ({backend_i1})",
                        _llm_probs_log,
                        delta="revisar" if _llm_probs_log else "referencias OK",
                        delta_color="inverse" if _llm_probs_log else "normal",
                    )
                with col_l2:
                    st.metric("Correcciones LLM aplicadas", _llm_corr_log)

                _llm_detalle = [
                    l for l in log
                    if "LLM" in l or ("Fila" in l and "corrección(es) aplicada" in l)
                ]
                if _llm_detalle:
                    with st.expander(
                        f"🤖 Detalle APA LLM ({_llm_probs_log} problema(s))",
                        expanded=_llm_probs_log > 0,
                    ):
                        for linea in _llm_detalle:
                            st.markdown(linea)

            if metricas["tiene_as"]:
                st.markdown("**Verificación A+Se:**")
                total_as = metricas["as_ok"] + metricas["as_error"] + metricas["as_manual"]
                st.progress(
                    metricas["as_ok"] / total_as if total_as else 0,
                    text=f"{metricas['as_ok']}✅  {metricas['as_error']}❌  "
                         f"{metricas['as_manual']}⚠️ manual  (de {total_as} hitos)",
                )

            if metricas["lt_ejecutado"]:
                lt_err = metricas["lt_errores"]
                lt_rev = metricas["lt_revisadas"]
                lt_corr = metricas["lt_correcciones"]
                delta_text = f"{lt_rev} celda(s) revisadas"
                if lt_corr > 0:
                    delta_text += f", {lt_corr} corregidas"
                st.metric(
                    "Errores ortográficos/gramaticales",
                    lt_err,
                    delta=delta_text,
                    delta_color="off" if lt_corr == 0 else "normal",
                    help="LanguageTool API (es). Las correcciones automáticas se aplican solo a cambios seguros.",
                )
                if metricas["lt_detalle"]:
                    with st.expander(
                        f"🔤 Detalle ortografía/gramática ({lt_err} error(es), {lt_corr} corregidos)",
                        expanded=lt_err > 0 or lt_corr > 0,
                    ):
                        for linea in metricas["lt_detalle"]:
                            st.markdown(linea)

            if metricas["discrepancias_detalle"]:
                with st.expander("📊 Detalle — verificación contra programa", expanded=True):
                    for linea in metricas["discrepancias_detalle"]:
                        st.markdown(linea)

            if metricas["tiene_as"] and metricas["as_detalle"]:
                with st.expander("📌 Detalle — hitos A+Se",
                                 expanded=metricas["as_error"] > 0):
                    for linea in metricas["as_detalle"]:
                        st.markdown(linea)

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

            with st.expander("🔍 Ver log completo de correcciones"):
                log_texto = "\n".join(log)
                st.code(log_texto, language=None)
                st.download_button(
                    "⬇ Descargar log (.txt)",
                    data=log_texto,
                    file_name=output_name.replace(".xlsx", "_log.txt"),
                    mime="text/plain",
                )

# ═════════════════════════════════════════════════════════════════════════
#  INSTANCIA 2 e INSTANCIA 3 (lógica compartida)
# ═════════════════════════════════════════════════════════════════════════

# ═════════════════════════════════════════════════════════════════════════
#  INSTANCIA 3
# ═════════════════════════════════════════════════════════════════════════

def _render_instancia_escala(tab, instancia_num, key_prefix):
    """Renderiza el formulario de Instancia 2 o 3 (lógica idéntica, distinto número)."""
    with tab:
        st.markdown("### 1 · Archivos")
        if instancia_num == 2:
            st.caption(
                "Sube la escala completada por la **1ª revisora DEL** "
                "y la planificación del docente."
            )
        else:
            st.caption(
                "Sube la escala completada por la **2ª revisora DEL** "
                "y el archivo `_I2_REVISADO.xlsx` generado en Instancia 2."
            )

        col_e, col_p = st.columns(2)
        with col_e:
            escala_f = st.file_uploader(
                "📋 Escala de apreciación completada",
                type="xlsx",
                key=f"{key_prefix}_escala",
                help="Archivo Excel de la escala completada por la revisora DEL.",
            )
        with col_p:
            plan_f = st.file_uploader(
                "📊 Planificación",
                type="xlsx",
                key=f"{key_prefix}_plan",
                help=("Planificación del docente o _I2_REVISADO.xlsx "
                      if instancia_num == 3 else
                      "Planificación del docente o _REVISADO de Instancia 1."),
            )

        if escala_f:
            with st.spinner("Leyendo escala..."):
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_e:
                    tmp_e.write(escala_f.getvalue())
                    tmp_e_path = tmp_e.name
                crits = rp.leer_escala_completa(tmp_e_path)
                os.unlink(tmp_e_path)

            n_ok_e   = sum(1 for c in crits if c["estado"] == "Si")
            n_parc_e = sum(1 for c in crits if c["estado"] == "Parcialmente")
            n_no_e   = sum(1 for c in crits if c["estado"] == "No")

            with st.expander(
                f"✅ Escala leída — {len(crits)} criterios: "
                f"{n_ok_e} ✅  {n_parc_e} ⚠️ parcial  {n_no_e} ❌",
                expanded=False,
            ):
                for c in crits:
                    icono = ("✅" if c["estado"] == "Si"
                             else ("⚠️" if c["estado"] == "Parcialmente" else "❌"))
                    st.markdown(f"{icono} **{c['criterio'][:80]}**")
                    if c["obs_texto"]:
                        st.caption(f"↳ {c['obs_texto'][:160]}")

        st.markdown("### 2 · Opciones")
        col_as_x, col_pdf_x = st.columns(2)
        with col_as_x:
            es_as_x = st.checkbox(
                "📌 Lineamiento A+S",
                key=f"{key_prefix}_as",
                help="Activa la verificación de los 11 hitos A+Se.",
            )
        with col_pdf_x:
            pdf_x = st.file_uploader(
                "📄 Programa (opcional)",
                type="pdf",
                key=f"{key_prefix}_pdf",
                help="PDF del programa para cruzar correcciones.",
            )

        usar_lt_x = st.checkbox(
            "🔤 Revisión ortográfica y gramatical (LanguageTool)",
            value=True,
            key=f"{key_prefix}_lt",
            help="Revisa ortografía y gramática de las actividades usando "
                 "LanguageTool API en español. Requiere conexión a internet.",
        )

        autocorregir_lt_x = st.checkbox(
            "✍️ Aplicar correcciones automáticas (conservador)",
            value=True,
            key=f"{key_prefix}_autocorr",
            help="Aplica automáticamente las correcciones SEGURAS detectadas por LanguageTool "
                 "(solo cambios unívocos: tildes faltantes, errores ortográficos claros).",
        )

        programa_x = {}
        if pdf_x:
            with st.spinner("Leyendo programa..."):
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_px:
                    tmp_px.write(pdf_x.read())
                    tmp_px_path = tmp_px.name
                programa_x = rp.extraer_programa_pdf(tmp_px_path)
                os.unlink(tmp_px_path)
            if programa_x.get("_error"):
                st.warning(f"No se pudo leer el PDF: {programa_x['_error']}")
                programa_x = {}

        st.markdown(f"### 3 · Aplicar correcciones")

        listo_x = bool(escala_f and plan_f)
        procesar_x = st.button(
            f"▶ Aplicar correcciones (Instancia {instancia_num})",
            disabled=not listo_x,
            type="primary",
            key=f"btn_{key_prefix}",
            use_container_width=True,
            help="Sube la escala y la planificación para activar."
                 if not listo_x else "",
        )
        if not listo_x:
            st.caption("⬆ Sube la escala completada y la planificación para continuar.")

        if procesar_x and listo_x:
            rp._LANGUAGETOOL_ACTIVO = usar_lt_x
            rp._LANGUAGETOOL_AUTOCORREGIR = autocorregir_lt_x
            spinner_x = (f"Aplicando correcciones Instancia {instancia_num} + revisión ortográfica..."
                         if usar_lt_x else
                         f"Aplicando correcciones Instancia {instancia_num}...")
            with st.spinner(spinner_x):
                log_x, ok_x, out_bytes_x, out_name_x = rp.procesar_instancia2(
                    plan_bytes=plan_f.getvalue(),
                    escala_bytes=escala_f.getvalue(),
                    plan_nombre=plan_f.name,
                    programa=programa_x or None,
                    es_as=es_as_x,
                    instancia_num=instancia_num,
                )
                if ok_x:
                    out_name_x = _nombre_descarga(programa_x, instancia_num, out_bytes_x)

            st.divider()
            st.markdown("### Resultados")

            if not ok_x or not out_bytes_x:
                st.error("El procesamiento falló. Revisa el log.")
                with st.expander("Log de error"):
                    st.code("\n".join(log_x), language=None)
            else:
                metricas_x = parsear_log(log_x)
                hist.registrar(instancia=instancia_num, archivo_nombre=out_name_x,
                               metricas=metricas_x, programa=programa_x or None)
                veces_x = hist.contar_por_codigo((programa_x or {}).get("codigo", ""))
                if veces_x >= 3:
                    st.warning(
                        f"⚠️ Esta asignatura ya fue procesada **{veces_x} veces**. "
                        "Verifica si corresponde a una nueva versión.",
                        icon="🔁",
                    )

                n_anot_x = 0
                for linea in log_x:
                    m = re.search(r"Anotaciones inyectadas:\s*(\d+)", linea)
                    if m:
                        n_anot_x = int(m.group(1))
                        break

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(
                        "Auto-correcciones",
                        metricas_x["total_correcciones"],
                        help="Correcciones estándar aplicadas.",
                    )
                with col2:
                    st.metric(
                        "Anotaciones revisora",
                        n_anot_x,
                        help="Textos de la escala inyectados en azul.",
                    )
                with col3:
                    total_crit_x = (metricas_x["criterios_ok"]
                                    + metricas_x["criterios_error"]
                                    + metricas_x["criterios_manual"])
                    st.metric(
                        "Criterios UST",
                        f"{metricas_x['criterios_ok']}/{total_crit_x}",
                        delta=f"{metricas_x['criterios_error']} con error"
                              if metricas_x["criterios_error"] else "sin errores",
                        delta_color="inverse" if metricas_x["criterios_error"] else "normal",
                    )

                if metricas_x["lt_ejecutado"]:
                    lt_err_x = metricas_x["lt_errores"]
                    lt_rev_x = metricas_x["lt_revisadas"]
                    lt_corr_x = metricas_x["lt_correcciones"]
                    delta_text_x = f"{lt_rev_x} celda(s) revisadas"
                    if lt_corr_x > 0:
                        delta_text_x += f", {lt_corr_x} corregidas"
                    st.metric(
                        "Errores ortográficos/gramaticales",
                        lt_err_x,
                        delta=delta_text_x,
                        delta_color="off" if lt_corr_x == 0 else "normal",
                        help="LanguageTool API (es). Las correcciones automáticas se aplican solo a cambios seguros.",
                    )
                    if metricas_x["lt_detalle"]:
                        with st.expander(
                            f"🔤 Detalle ortografía/gramática ({lt_err_x} error(es), {lt_corr_x} corregidos)",
                            expanded=lt_err_x > 0 or lt_corr_x > 0,
                        ):
                            for linea in metricas_x["lt_detalle"]:
                                st.markdown(linea)

                if metricas_x["discrepancias_detalle"]:
                    with st.expander("📊 Verificación contra programa", expanded=True):
                        for linea in metricas_x["discrepancias_detalle"]:
                            st.markdown(linea)

                if instancia_num == 3 and metricas_x["criterios_error"] == 0:
                    st.success(
                        "Sin criterios con error automático. "
                        "La planificación está lista para aprobación final.",
                        icon="🎓",
                    )

                st.info(
                    "La hoja **NOTAS_CORRECCIONES_DEL** en el archivo descargado "
                    "contiene el resumen completo de observaciones de la revisora.",
                    icon="📋",
                )

                st.divider()
                st.download_button(
                    label=f"⬇ Descargar {out_name_x}",
                    data=out_bytes_x,
                    file_name=out_name_x,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
                st.caption(
                    "Azul = corrección automática o anotación de la revisora. "
                    "Lila = Formativa · Amarillo = Sumativa. "
                    "Hoja NOTAS_CORRECCIONES_DEL = resumen completo."
                )

                with st.expander("🔍 Ver log completo"):
                    log_x_texto = "\n".join(log_x)
                    st.code(log_x_texto, language=None)
                    st.download_button(
                        "⬇ Descargar log (.txt)",
                        data=log_x_texto,
                        file_name=out_name_x.replace(".xlsx", "_log.txt"),
                        mime="text/plain",
                    )


# Renderizar I2 e I3 con la función compartida
_render_instancia_escala(tab_i2, instancia_num=2, key_prefix="i2")
_render_instancia_escala(tab_i3, instancia_num=3, key_prefix="i3")

# ═════════════════════════════════════════════════════════════════════════
#  HISTORIAL
# ═════════════════════════════════════════════════════════════════════════

with tab_hist:
    st.markdown("### Historial de planificaciones procesadas")

    registros = hist.obtener_historial()

    if not registros:
        st.info("Aún no hay registros. Procesa una planificación para comenzar el historial.")
    else:
        # ── Métricas globales ───────────────────────────────────────────
        total_proc   = len(registros)
        total_corr   = sum(r["total_correcciones"] for r in registros)
        total_crit_e = sum(r["criterios_error"] for r in registros)

        c1, c2, c3 = st.columns(3)
        c1.metric("Planificaciones procesadas", total_proc)
        c2.metric("Correcciones acumuladas", total_corr)
        c3.metric("Criterios con error (acum.)", total_crit_e)

        st.divider()

        # ── Tabla de registros ──────────────────────────────────────────
        import pandas as pd

        df = pd.DataFrame(registros)
        df = df.rename(columns={
            "id":                  "ID",
            "fecha_hora":          "Fecha",
            "instancia":           "I",
            "codigo_asignatura":   "Código",
            "nombre_asignatura":   "Asignatura",
            "archivo_nombre":      "Archivo",
            "total_correcciones":  "Corr.",
            "criterios_ok":        "✅",
            "criterios_error":     "❌",
            "criterios_manual":    "⚠️",
            "discrepancias_prog":  "Disc.",
            "lt_errores":          "LT err.",
            "lt_correcciones":     "LT corr.",
            "tiene_as":            "A+S",
            "as_ok":               "A+S✅",
            "as_error":            "A+S❌",
        })
        df["A+S"] = df["A+S"].map({0: "", 1: "Sí"})
        st.dataframe(
            df.drop(columns=["A+S✅", "A+S❌"], errors="ignore"),
            use_container_width=True,
            hide_index=True,
        )

        # ── Resumen por asignatura ──────────────────────────────────────
        st.divider()
        st.markdown("#### Resumen por asignatura")
        resumen = hist.resumen_errores()
        if resumen:
            df_res = pd.DataFrame(resumen).rename(columns={
                "codigo_asignatura":  "Código",
                "nombre_asignatura":  "Asignatura",
                "veces_procesada":    "Veces",
                "total_corr":         "Corr. totales",
                "total_crit_err":     "Crit. error",
                "total_disc":         "Discrepancias",
                "total_lt":           "LT errores",
            })
            st.dataframe(df_res, use_container_width=True, hide_index=True)

        # ── Exportar CSV ────────────────────────────────────────────────
        st.divider()
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇ Exportar historial completo (.csv)",
            data=csv_bytes,
            file_name="historial_del.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )

# ═════════════════════════════════════════════════════════════════════════
#  DICCIONARIO UST
# ═════════════════════════════════════════════════════════════════════════

with tab_dict:
    st.markdown("### Diccionario UST personalizable")
    st.caption(
        "Las entradas base (en gris) vienen del script y no se pueden borrar desde aquí. "
        "Las entradas personalizadas (en verde) se suman a las base y tienen prioridad."
    )

    completo = dict_ust.obtener_dict_completo()
    custom   = dict_ust.obtener_entradas_custom()

    # ── Ver diccionario completo ──────────────────────────────────────────
    for mapa in dict_ust.MAPAS:
        entradas = completo[mapa]
        n_custom = len(custom.get(mapa, {}))
        with st.expander(
            f"**{dict_ust.etiqueta(mapa)}** — {len(entradas)} entradas "
            f"({n_custom} personalizadas)",
            expanded=False,
        ):
            import pandas as pd
            filas = []
            for inc, corr in sorted(entradas.items()):
                es_custom = inc in custom.get(mapa, {})
                filas.append({
                    "Término incorrecto": inc,
                    "→ Corrección":       corr,
                    "Origen":             "✏️ Personalizada" if es_custom else "📋 Base",
                })
            if filas:
                st.dataframe(pd.DataFrame(filas), use_container_width=True,
                             hide_index=True)

    st.divider()

    # ── Agregar nueva entrada ─────────────────────────────────────────────
    st.markdown("#### Agregar entrada")
    col_m, col_i, col_c = st.columns([2, 3, 3])
    with col_m:
        mapa_sel = st.selectbox(
            "Mapa",
            dict_ust.MAPAS,
            format_func=dict_ust.etiqueta,
            key="dict_mapa",
        )
    with col_i:
        nuevo_inc  = st.text_input("Término incorrecto", key="dict_inc",
                                   placeholder="ej: Examen")
    with col_c:
        nuevo_corr = st.text_input("Corrección UST",     key="dict_corr",
                                   placeholder="ej: Pruebas escritas u orales")

    if st.button("➕ Agregar al diccionario", key="btn_dict_add", type="primary", use_container_width=True):
        if nuevo_inc.strip() and nuevo_corr.strip():
            dict_ust.agregar_entrada(mapa_sel, nuevo_inc, nuevo_corr)
            st.success(f"Entrada agregada: «{nuevo_inc}» → «{nuevo_corr}»")
            st.rerun()
        else:
            st.warning("Completa ambos campos antes de agregar.")

    st.divider()

    # ── Eliminar entrada personalizada ────────────────────────────────────
    st.markdown("#### Eliminar entrada personalizada")
    todas_custom = [
        (mapa, inc)
        for mapa in dict_ust.MAPAS
        for inc in custom.get(mapa, {})
    ]
    if todas_custom:
        opciones = [f"{dict_ust.etiqueta(m)} → «{i}»" for m, i in todas_custom]
        sel_idx  = st.selectbox("Selecciona entrada a eliminar", range(len(opciones)),
                                format_func=lambda i: opciones[i], key="dict_del_sel")
        if st.button("🗑 Eliminar", key="btn_dict_del", type="secondary"):
            mapa_d, inc_d = todas_custom[sel_idx]
            dict_ust.eliminar_entrada(mapa_d, inc_d)
            st.success(f"Entrada «{inc_d}» eliminada.")
            st.rerun()
    else:
        st.caption("No hay entradas personalizadas para eliminar.")

    st.divider()

    # ── Exportar / Importar ───────────────────────────────────────────────
    st.markdown("#### Exportar / Importar")
    col_exp, col_imp = st.columns(2)
    with col_exp:
        st.download_button(
            "⬇ Exportar diccionario (.json)",
            data=dict_ust.exportar_json(),
            file_name="dict_ust.json",
            mime="application/json",
        )
    with col_imp:
        json_up = st.file_uploader("⬆ Importar diccionario (.json)",
                                   type="json", key="dict_import")
        if json_up:
            n = dict_ust.importar_json(json_up.getvalue())
            st.success(f"{n} entrada(s) importadas correctamente.")
            st.rerun()

    st.divider()

    # ── Generador de referencia APA 7 desde URL / DOI ─────────────────────
    st.markdown("#### Generar referencia APA 7 desde URL o DOI")
    st.caption(
        "Ingresa una URL o DOI para generar automáticamente la referencia APA 7. "
        "Requiere conexión a internet. Revisa el resultado antes de usarlo."
    )

    url_input = st.text_input(
        "URL o DOI",
        key="apa_url_input",
        placeholder="https://doi.org/10.1016/... o https://www.sitio.cl/articulo",
    )

    if st.button("🔍 Generar referencia", key="btn_apa_gen"):
        if url_input.strip():
            with st.spinner("Consultando metadatos..."):
                ref_gen, estado = apa.generar_desde_url(url_input.strip())
            if ref_gen:
                st.success(f"Referencia generada ({estado}):")
                st.code(ref_gen, language=None)
                st.caption(
                    "Copia esta referencia y pégala en la columna H del Excel. "
                    "Verifica que el año, nombre de autores y tipo de recurso sean correctos."
                )
            else:
                st.warning(f"No se pudo generar la referencia: {estado}")
        else:
            st.warning("Ingresa una URL o DOI para continuar.")
