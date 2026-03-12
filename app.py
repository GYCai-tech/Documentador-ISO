"""
app.py — Interfaz web para el Asistente ISO de GÓMEZ Y CRESPO S.A.
Ejecutar: streamlit run app.py
"""

import json
import os
import sys
import tempfile

import streamlit as st

WORKDIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, WORKDIR)

# ── Configuración de página ────────────────────────────────────────
st.set_page_config(
    page_title="GYC · Asistente ISO",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizado ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Crimson+Pro:ital,wght@0,400;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

:root {
    --bg:        #0c0f18;
    --surface:   #141824;
    --border:    #232840;
    --border-hi: #2e3654;
    --primary:   #4f7ef0;
    --accent:    #b8d44a;
    --text:      #dde3f0;
    --muted:     #5a6484;
    --danger:    #e05c5c;
}

/* ── Estructura base ── */
#MainMenu, footer, header, .stDeployButton { visibility: hidden; display: none; }
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapseButton"] { visibility: visible !important; display: flex !important; }
.stApp { background: var(--bg); }
.main .block-container {
    max-width: 920px;
    padding: 2.5rem 2rem 4rem;
    font-family: 'IBM Plex Sans', sans-serif;
}

/* ── Header ── */
.gyc-header {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    border-bottom: 1px solid var(--border);
    padding-bottom: 1.75rem;
    margin-bottom: 2.25rem;
}
.gyc-wordmark {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.4rem;
}
.gyc-title {
    font-family: 'Crimson Pro', serif;
    font-size: 2.4rem;
    font-weight: 700;
    color: var(--text);
    line-height: 1;
    letter-spacing: -0.01em;
}
.gyc-badge {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.6rem;
    letter-spacing: 0.12em;
    color: var(--muted);
    border: 1px solid var(--border);
    padding: 0.3rem 0.6rem;
    border-radius: 3px;
    margin-top: 0.25rem;
    text-transform: uppercase;
    white-space: nowrap;
}

/* ── Barra de fase ── */
.phase-track {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 6px;
    margin-bottom: 0.5rem;
}
.phase-seg {
    height: 2px;
    border-radius: 1px;
    background: var(--border);
    transition: background 0.4s;
}
.phase-seg.done    { background: var(--accent); }
.phase-seg.active  { background: var(--primary); }
.phase-labels {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 6px;
    margin-bottom: 2rem;
}
.phase-lbl {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.6rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    text-align: center;
    color: var(--muted);
}
.phase-lbl.done   { color: var(--accent); }
.phase-lbl.active { color: var(--primary); }

/* ── Cards ── */
.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1.25rem;
}
.card-accent-left  { border-left: 3px solid var(--primary); }
.card-accent-green { border-left: 3px solid var(--accent); }
.card-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.6rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.4rem;
}
.card-title {
    font-family: 'Crimson Pro', serif;
    font-size: 1.25rem;
    font-weight: 600;
    color: var(--text);
}
.card-sub {
    font-size: 0.8rem;
    color: var(--muted);
    margin-top: 0.2rem;
}

/* ── Mensajes del chat ── */
[data-testid="stChatMessage"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    padding: 1rem 1.25rem !important;
    margin-bottom: 0.5rem !important;
}
[data-testid="stChatMessage"] p,
[data-testid="stChatMessage"] li,
[data-testid="stChatMessage"] span {
    color: var(--text) !important;
    font-size: 0.95rem !important;
    line-height: 1.75 !important;
}
/* Mensaje del asistente: acento izquierdo verde */
[data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-assistant"]) {
    border-left: 3px solid var(--accent) !important;
    background: #161b2a !important;
}
/* Mensaje del usuario: acento izquierdo azul */
[data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) {
    border-left: 3px solid var(--primary) !important;
}

/* ── Input de chat ── */
[data-testid="stChatInput"] textarea {
    background: var(--surface) !important;
    border: 1px solid var(--border-hi) !important;
    border-radius: 8px !important;
    color: var(--text) !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 0.875rem !important;
}
[data-testid="stChatInput"] textarea:focus {
    border-color: var(--primary) !important;
    box-shadow: 0 0 0 3px rgba(79,126,240,0.15) !important;
}

/* ── Text input de formulario ── */
.stTextInput label { color: var(--muted) !important; font-size: 0.8rem !important; }
.stTextInput input {
    background: var(--surface) !important;
    border: 1px solid var(--border-hi) !important;
    border-radius: 6px !important;
    color: var(--text) !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 0.9rem !important;
    padding: 0.65rem 0.9rem !important;
}
.stTextInput input:focus {
    border-color: var(--primary) !important;
    box-shadow: 0 0 0 3px rgba(79,126,240,0.15) !important;
}

/* ── Botones ── */
.stButton > button, .stDownloadButton > button {
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.02em !important;
    border-radius: 6px !important;
    border: none !important;
    padding: 0.6rem 1.4rem !important;
    transition: opacity 0.2s !important;
}
.stButton > button[kind="primary"] {
    background: var(--primary) !important;
    color: #fff !important;
}
.stButton > button[kind="secondary"] {
    background: var(--surface) !important;
    color: var(--text) !important;
    border: 1px solid var(--border-hi) !important;
}
.stButton > button:hover, .stDownloadButton > button:hover { opacity: 0.85 !important; }
.stDownloadButton > button {
    background: var(--accent) !important;
    color: #0c0f18 !important;
    font-weight: 600 !important;
}

/* ── Spinner ── */
.stSpinner > div { border-top-color: var(--primary) !important; }

/* ── Separador ── */
hr { border-color: var(--border) !important; }

/* ── Alertas ── */
.stAlert { background: var(--surface) !important; border-radius: 8px !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebar"] * { color: var(--text) !important; }
[data-testid="stSidebar"] .stSlider label,
[data-testid="stSidebar"] .stNumberInput label {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.72rem !important;
    letter-spacing: 0.05em !important;
    color: var(--muted) !important;
}
[data-testid="stSidebar"] .stSlider [data-baseweb="slider"] div[role="slider"] {
    background: var(--primary) !important;
    border-color: var(--primary) !important;
}
[data-testid="stSidebar"] .stSlider [data-baseweb="slider"] div[data-testid="stThumbValue"] {
    background: var(--primary) !important;
    color: #fff !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.65rem !important;
}
[data-testid="stSidebar"] .stSlider [data-baseweb="slider"] > div > div:first-child {
    background: var(--border) !important;
}
[data-testid="stSidebar"] .stSlider [data-baseweb="slider"] > div > div:nth-child(2) {
    background: var(--primary) !important;
}
[data-testid="stSidebar"] .stNumberInput input {
    background: var(--bg) !important;
    border: 1px solid var(--border-hi) !important;
    border-radius: 5px !important;
    color: var(--text) !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important;
}
[data-testid="stSidebar"] hr {
    border-color: var(--border) !important;
    margin: 1rem 0 !important;
}
.sidebar-section-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.6rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.75rem;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid var(--border);
}
/* ── Textarea de prompts en sidebar ── */
[data-testid="stSidebar"] textarea {
    background: var(--bg) !important;
    border: 1px solid var(--border-hi) !important;
    border-radius: 5px !important;
    color: var(--text) !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.72rem !important;
    line-height: 1.6 !important;
}
[data-testid="stSidebar"] textarea:focus {
    border-color: var(--primary) !important;
}

.sidebar-model-tag {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.65rem;
    color: var(--muted);
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0.15rem 0.45rem;
    display: inline-block;
    margin-bottom: 0.75rem;
}
</style>
""", unsafe_allow_html=True)


# ── Historial de procedimientos ────────────────────────────────────
HISTORIAL_PATH = os.path.join(WORKDIR, "historial.json")

def load_historial() -> list[dict]:
    if not os.path.exists(HISTORIAL_PATH):
        return []
    try:
        with open(HISTORIAL_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []

def append_historial(data: dict, doc_path: str):
    from datetime import datetime
    historial = load_historial()
    historial.append({
        "codigo":           data.get("codigo", ""),
        "nombre":           data.get("nombre", ""),
        "fecha_generacion": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "ruta_docx":        doc_path,
    })
    with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)


# ── Prompts por defecto (importados del módulo) ────────────────────
def _default_prompts():
    import asistente_iso
    return asistente_iso.INTERVIEW_SYSTEM, asistente_iso.DRAFT_SYSTEM


# ── Session state ──────────────────────────────────────────────────
_DEFAULTS = {
    "fase":        "inicio",   # inicio | entrevista | redactando | revision_chat | revision_redactando | listo
    "modo":        "nuevo",    # "nuevo" | "revision"
    "mensajes":    [],
    "chat":        None,
    "log":         [],
    "rag_context": "",
    "doc_path":    None,
    "topic":       "",
    "data":        None,   # JSON del procedimiento generado por Gemini Pro
    # Modo revisión
    "docx_original_text": "",
    "revision_mensajes":  [],
    "revision_chat":      None,
    "revision_log":       [],
    # Parámetros Flash (entrevista)
    "flash_temperature":  0.7,
    "flash_top_k":        40,
    "flash_top_p":        0.95,
    "flash_max_tokens":   2048,
    # Parámetros Pro (redacción final)
    "pro_temperature":    0.3,
    "pro_top_k":          20,
    "pro_top_p":          0.85,
    "pro_max_tokens":     8192,
    # Prompts (inicializados lazy)
    "interview_prompt":   None,
    "draft_prompt":       None,
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# Inicialización lazy de prompts
if st.session_state.interview_prompt is None or st.session_state.draft_prompt is None:
    _ip, _dp = _default_prompts()
    st.session_state.interview_prompt = _ip
    st.session_state.draft_prompt     = _dp


# ── Sidebar: configuración de modelos ─────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown('<div class="sidebar-section-title">Configuración de modelos</div>', unsafe_allow_html=True)

        # ── Flash ───────────────────────────────────────────────────
        st.markdown("**Entrevista**")
        st.markdown('<span class="sidebar-model-tag">gemini-2.5-flash</span>', unsafe_allow_html=True)

        st.session_state.flash_temperature = st.slider(
            "Temperature", 0.0, 1.0,
            value=st.session_state.flash_temperature,
            step=0.05, key="sl_flash_temp",
            help="Creatividad de las respuestas. Valores altos → más variado.",
        )
        st.session_state.flash_top_k = st.slider(
            "Top-K", 1, 100,
            value=st.session_state.flash_top_k,
            step=1, key="sl_flash_topk",
            help="Número de tokens candidatos en cada paso.",
        )
        st.session_state.flash_top_p = st.slider(
            "Top-P", 0.5, 1.0,
            value=st.session_state.flash_top_p,
            step=0.01, key="sl_flash_topp",
            help="Nucleus sampling: masa de probabilidad acumulada.",
        )
        st.session_state.flash_max_tokens = st.number_input(
            "Max tokens", 256, 8192,
            value=st.session_state.flash_max_tokens,
            step=256, key="ni_flash_tokens",
        )

        st.markdown("---")

        # ── Pro ─────────────────────────────────────────────────────
        st.markdown("**Redacción final**")
        st.markdown('<span class="sidebar-model-tag">gemini-2.5-pro</span>', unsafe_allow_html=True)

        st.session_state.pro_temperature = st.slider(
            "Temperature", 0.0, 1.0,
            value=st.session_state.pro_temperature,
            step=0.05, key="sl_pro_temp",
            help="Valores bajos → texto más fiel al estilo ISO de referencia.",
        )
        st.session_state.pro_top_k = st.slider(
            "Top-K", 1, 100,
            value=st.session_state.pro_top_k,
            step=1, key="sl_pro_topk",
        )
        st.session_state.pro_top_p = st.slider(
            "Top-P", 0.5, 1.0,
            value=st.session_state.pro_top_p,
            step=0.01, key="sl_pro_topp",
        )
        st.session_state.pro_max_tokens = st.number_input(
            "Max tokens", 1024, 32768,
            value=st.session_state.pro_max_tokens,
            step=512, key="ni_pro_tokens",
        )

        st.markdown("---")

        if st.button("Restablecer valores", use_container_width=True, type="secondary"):
            for k, v in {
                "flash_temperature": 0.7, "flash_top_k": 40,
                "flash_top_p": 0.95,      "flash_max_tokens": 2048,
                "pro_temperature":   0.3,  "pro_top_k": 20,
                "pro_top_p":         0.85, "pro_max_tokens": 8192,
            }.items():
                st.session_state[k] = v
            st.rerun()

        st.markdown("---")

        # ── Prompts del agente ───────────────────────────────────────
        st.markdown('<div class="sidebar-section-title">Prompts del agente</div>', unsafe_allow_html=True)

        with st.expander("Entrevista · Flash"):
            st.session_state.interview_prompt = st.text_area(
                "interview_prompt",
                value=st.session_state.interview_prompt,
                height=320,
                key="ta_interview_prompt",
                label_visibility="collapsed",
                help="System prompt del agente de entrevista (gemini-2.5-flash).",
            )

        with st.expander("Redacción · Pro"):
            st.session_state.draft_prompt = st.text_area(
                "draft_prompt",
                value=st.session_state.draft_prompt,
                height=320,
                key="ta_draft_prompt",
                label_visibility="collapsed",
                help="System prompt del agente redactor ISO (gemini-2.5-pro).",
            )

        if st.button("Restablecer prompts", use_container_width=True, type="secondary"):
            _ip, _dp = _default_prompts()
            st.session_state.interview_prompt = _ip
            st.session_state.draft_prompt     = _dp
            st.rerun()

        st.markdown("---")

        # ── Historial ───────────────────────────────────────────────
        st.markdown('<div class="sidebar-section-title">Historial</div>', unsafe_allow_html=True)

        entries = load_historial()
        if not entries:
            st.markdown(
                '<div style="font-size:0.75rem;color:#5a6484">Sin procedimientos generados aún.</div>',
                unsafe_allow_html=True,
            )
        else:
            for entry in reversed(entries):
                fname = os.path.basename(entry["ruta_docx"])
                label = f"{entry['codigo']} · {entry['fecha_generacion']}"
                if os.path.exists(entry["ruta_docx"]):
                    with open(entry["ruta_docx"], "rb") as fh:
                        st.download_button(
                            label=label,
                            data=fh,
                            file_name=fname,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"hist_{entry['ruta_docx']}",
                            use_container_width=True,
                        )
                else:
                    st.markdown(
                        f'<div style="font-size:0.72rem;color:#5a6484">{label} (no encontrado)</div>',
                        unsafe_allow_html=True,
                    )


def apply_model_configs():
    """Sobreescribe configs y prompts del módulo asistente_iso con los valores del sidebar."""
    import asistente_iso
    from vertexai.generative_models import GenerationConfig

    asistente_iso.CHAT_CONFIG = GenerationConfig(
        temperature=st.session_state.flash_temperature,
        top_k=int(st.session_state.flash_top_k),
        top_p=st.session_state.flash_top_p,
        max_output_tokens=int(st.session_state.flash_max_tokens),
    )
    asistente_iso.DRAFT_CONFIG = GenerationConfig(
        temperature=st.session_state.pro_temperature,
        top_k=int(st.session_state.pro_top_k),
        top_p=st.session_state.pro_top_p,
        max_output_tokens=int(st.session_state.pro_max_tokens),
    )
    asistente_iso.INTERVIEW_SYSTEM = st.session_state.interview_prompt
    asistente_iso.DRAFT_SYSTEM     = st.session_state.draft_prompt


# ── Preview del documento (eliminada) ─────────────────────────────
def generate_doc_html(data: dict) -> str:  # noqa: unused — mantenido por compatibilidad
    """Legacy: generaba HTML de preview. Ya no se usa en la UI."""
    codigo = data.get("codigo", "PC-XX")
    nombre = data.get("nombre", "")
    fecha  = data.get("fecha", "")
    rev    = data.get("revision", "00")
    elab   = data.get("elaborado_por", "")
    apr    = data.get("aprobado_por", "")

    def section(title, content_html):
        return f"""
        <div style="background:#E9EFB1;padding:4px 8px;font-weight:bold;
                    font-size:9.5pt;margin-top:14px;margin-bottom:4px">
            {title}
        </div>
        {content_html}"""

    def row(cells, header=False):
        tds = "".join(
            f'<td style="border:1px solid #aaa;padding:3px 6px;'
            f'{"font-weight:bold;background:#95B3D7;" if header else ""}">{c}</td>'
            for c in cells
        )
        return f"<tr>{tds}</tr>"

    # ── Control de revisiones ────────────────────────────────────────
    hist_rows = "".join(
        row([h.get("rev",""), h.get("fecha",""), h.get("descripcion",""),
             h.get("revisado",""), h.get("elaborado","")], False)
        for h in data.get("historial", [])
    )
    hist_rows += row(["REV","FECHA","DESCRIPCIÓN DE LOS CAMBIOS",
                       "REVISADO Y APROBADO","ELABORADO"], header=True)
    revisiones = f'<table style="width:100%;border-collapse:collapse;font-size:8.5pt">{hist_rows}</table>'

    # ── Índice ───────────────────────────────────────────────────────
    secciones = ["Objeto.", "Alcance.", "Responsabilidades."]
    idx_lines = [f'<div style="color:#4f7ef0;font-weight:bold;font-size:9pt;'
                 f'padding:1px 0">{i+1}. {s}</div>'
                 for i, s in enumerate(secciones)]
    idx_lines.append(
        f'<div style="color:#4f7ef0;font-weight:bold;font-size:9pt;padding:1px 0">'
        f'{len(secciones)+1}. Desarrollo.</div>'
    )
    for item in data.get("desarrollo", []):
        idx_lines.append(
            f'<div style="color:#4f7ef0;font-size:8.5pt;padding:1px 0 1px 16px">'
            f'{item["num"]} {item["titulo"]}.</div>'
        )
    for i, s in enumerate(["Archivo.", "Diagrama de Flujo.", "Referencias.", "Anexos."]):
        idx_lines.append(
            f'<div style="color:#4f7ef0;font-weight:bold;font-size:9pt;padding:1px 0">'
            f'{len(secciones)+2+i}. {s}</div>'
        )
    indice_html = "".join(idx_lines)

    # ── Responsabilidades ────────────────────────────────────────────
    resp_html = ""
    for rol in data.get("responsabilidades", []):
        resp_html += f'<div style="font-weight:bold;font-size:9pt;margin-top:6px">{rol["cargo"]}</div>'
        for t in rol.get("tareas", []):
            resp_html += f'<div style="font-size:8.5pt;padding:1px 0 1px 12px">• {t}</div>'

    # ── Desarrollo ───────────────────────────────────────────────────
    dev_html = ""
    for item in data.get("desarrollo", []):
        dev_html += (
            f'<div style="font-weight:bold;font-size:9pt;margin-top:8px">'
            f'{item["num"]}  {item["titulo"]}</div>'
            f'<div style="font-size:8.5pt;text-align:justify;margin-top:2px">'
            f'{item["descripcion"]}</div>'
        )

    # ── Archivo ──────────────────────────────────────────────────────
    arch_rows = row(["Documento", "Responsable", "Lugar"], header=True)
    for f_ in data.get("archivo", []):
        arch_rows += row([f_.get("documento",""), f_.get("responsable",""), f_.get("lugar","")])
    archivo_html = f'<table style="width:100%;border-collapse:collapse;font-size:8.5pt">{arch_rows}</table>'

    # ── Referencias y Anexos ─────────────────────────────────────────
    refs = "".join(f'<div style="font-size:8.5pt;padding:1px 0">• {r}</div>'
                   for r in data.get("referencias", []))
    anex = "".join(f'<div style="font-size:8.5pt;padding:1px 0">• {a}</div>'
                   for a in data.get("anexos", [])) or '<div style="font-size:8.5pt;color:#888">No aplica.</div>'

    return f"""
    <div style="font-family:Verdana,sans-serif;font-size:9pt;color:#111;
                background:#fff;padding:12px;line-height:1.5">

      <!-- HEADER -->
      <div style="background:#00FFFF;border:1px solid #aaa;
                  display:grid;grid-template-columns:auto 1fr;margin-bottom:8px">
        <div style="padding:6px 10px;border-right:1px solid #aaa;
                    font-weight:bold;font-size:9pt;display:flex;align-items:center">
          GYC
        </div>
        <div style="padding:6px 10px;font-weight:bold;font-size:10pt;
                    text-align:center;align-self:center">
          {codigo}: {nombre}
        </div>
      </div>
      <div style="background:#E9EFB1;border:1px solid #aaa;
                  display:grid;grid-template-columns:1fr 1fr;
                  font-size:8.5pt;margin-bottom:12px">
        <div style="padding:3px 8px;border-right:1px solid #aaa">
          Elaborado: {elab}
        </div>
        <div style="padding:3px 8px">
          Revisado y Aprobado: {apr}
        </div>
      </div>

      <!-- METADATOS -->
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;
                  border:1px solid #aaa;font-size:8.5pt;margin-bottom:12px">
        <div style="background:#95B3D7;padding:3px 8px;text-align:center;border-right:1px solid #aaa">
          <b>FECHA:</b> {fecha}
        </div>
        <div style="background:#E9EFB1;padding:3px 8px;text-align:center;border-right:1px solid #aaa">
          <b>REVISIÓN:</b> {rev}
        </div>
        <div style="background:#95B3D7;padding:3px 8px;text-align:center">
          <b>PÁGINAS:</b> {data.get("paginas", "")}
        </div>
      </div>

      {section("CONTROL DE REVISIONES", revisiones)}
      {section("ÍNDICE", indice_html)}
      {section("OBJETO", f'<div style="font-size:8.5pt;text-align:justify">{data.get("objeto","")}</div>')}
      {section("ALCANCE", f'<div style="font-size:8.5pt;text-align:justify">{data.get("alcance","")}</div>')}
      {section("RESPONSABILIDADES", resp_html)}
      {section("DESARROLLO", dev_html)}
      {section("ARCHIVO", archivo_html)}
      {section("DIAGRAMA DE FLUJO",
               '<div style="font-size:8.5pt;color:#888;font-style:italic">[Ver documento Word]</div>')}
      {section("REFERENCIAS", refs)}
      {section("ANEXOS", anex)}
    </div>
    """


def render_preview():
    """Panel derecho: muestra la preview del documento según la fase actual."""
    data  = st.session_state.get("data")
    fase  = st.session_state.fase
    topic = st.session_state.get("topic", "")

    st.markdown(
        '<div style="font-family:\'IBM Plex Mono\',monospace;font-size:0.6rem;'
        'letter-spacing:0.18em;text-transform:uppercase;color:#b8d44a;'
        'margin-bottom:0.75rem;padding-bottom:0.4rem;border-bottom:1px solid #232840">'
        'Vista previa del documento</div>',
        unsafe_allow_html=True,
    )

    if data:
        html = generate_doc_html(data)
        st.html(f'<div style="height:78vh;overflow-y:auto;border:1px solid #232840;'
                f'border-radius:6px">{html}</div>')
    elif fase == "inicio":
        st.markdown(
            '<div style="height:78vh;border:1px dashed #232840;border-radius:6px;'
            'display:flex;align-items:center;justify-content:center;'
            'color:#5a6484;font-family:\'IBM Plex Mono\',monospace;font-size:0.75rem">'
            'El documento aparecerá aquí</div>',
            unsafe_allow_html=True,
        )
    elif fase == "entrevista":
        n = len(st.session_state.get("mensajes", []))
        pct = min(int((n / 20) * 100), 95)
        st.markdown(
            f'<div style="height:78vh;border:1px solid #232840;border-radius:6px;'
            f'padding:1.5rem;display:flex;flex-direction:column;gap:1rem">'
            f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:0.65rem;'
            f'color:#5a6484;text-transform:uppercase;letter-spacing:0.1em">Recopilando información</div>'
            f'<div style="font-family:\'Crimson Pro\',serif;font-size:1.2rem;'
            f'font-weight:600;color:#dde3f0">{topic}</div>'
            f'<div style="background:#232840;border-radius:4px;height:4px;margin-top:0.5rem">'
            f'<div style="background:#4f7ef0;width:{pct}%;height:100%;border-radius:4px;'
            f'transition:width 0.5s"></div></div>'
            f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:0.6rem;'
            f'color:#5a6484">{n} intercambios · La preview aparecerá al terminar la entrevista</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
    elif fase == "redactando":
        st.markdown(
            '<div style="height:78vh;border:1px solid #232840;border-radius:6px;'
            'display:flex;align-items:center;justify-content:center;flex-direction:column;gap:1rem">'
            '<div style="font-family:\'IBM Plex Mono\',monospace;font-size:0.7rem;'
            'color:#4f7ef0;text-transform:uppercase;letter-spacing:0.1em">Generando documento...</div>'
            '</div>',
            unsafe_allow_html=True,
        )


# ── RAG index (cacheado — se construye una sola vez) ───────────────
@st.cache_resource(show_spinner="Cargando índice de documentos...")
def get_rag_index():
    import vertexai
    from asistente_iso import PROJECT_ID, LOCATION, load_or_build_index
    vertexai.init(project=PROJECT_ID, location=LOCATION)
    return load_or_build_index()


# ── Helper: barra de progreso de fases ────────────────────────────
def render_phases(fase_actual: str):
    if st.session_state.get("modo") == "revision":
        fases  = ["inicio", "revision_chat", "revision_redactando", "listo"]
        labels = ["Inicio", "Cambios", "Redacción", "Descarga"]
    else:
        fases  = ["inicio", "entrevista", "redactando", "listo"]
        labels = ["Inicio", "Entrevista", "Redacción", "Descarga"]
    idx = fases.index(fase_actual) if fase_actual in fases else 0

    segs = "".join(
        f'<div class="phase-seg {"done" if i < idx else "active" if i == idx else ""}"></div>'
        for i in range(4)
    )
    lbls = "".join(
        f'<div class="phase-lbl {"done" if i < idx else "active" if i == idx else ""}">{labels[i]}</div>'
        for i in range(4)
    )
    st.markdown(
        f'<div class="phase-track">{segs}</div>'
        f'<div class="phase-labels">{lbls}</div>',
        unsafe_allow_html=True,
    )


render_sidebar()

# ══════════════════════════════════════════════════════════════════
# CABECERA
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<div class="gyc-header">
  <div>
    <div class="gyc-wordmark">Gómez y Crespo S.A.</div>
    <div class="gyc-title">Asistente ISO</div>
  </div>
  <div class="gyc-badge">ISO 9001 · ISO 14001</div>
</div>
""", unsafe_allow_html=True)

render_phases(st.session_state.fase)

# ══════════════════════════════════════════════════════════════════
# Flujo de fases (columna única)
# ══════════════════════════════════════════════════════════════════
# ── FASE 1 — INICIO ───────────────────────────────────────────
if st.session_state.fase == "inicio":

    modo_sel = st.radio(
        "modo",
        options=["Crear nuevo procedimiento", "Revisar procedimiento existente"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if modo_sel == "Crear nuevo procedimiento":
        st.markdown("""
        <div class="card card-accent-left">
          <div class="card-label">Instrucciones</div>
          <div class="card-sub" style="color:#dde3f0;font-size:0.875rem;line-height:1.6">
            Describe brevemente el procedimiento que quieres documentar.<br>
            El asistente te guiará sección a sección mediante una entrevista
            y generará automáticamente el documento Word con formato GYC.
          </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("form_inicio", clear_on_submit=False):
            topic = st.text_input(
                "Procedimiento",
                placeholder="Ej: gestión de reclamaciones de clientes",
                label_visibility="visible",
            )
            col_a, col_b = st.columns([3, 1])
            with col_b:
                submit = st.form_submit_button("Comenzar →", type="primary", use_container_width=True)

        if submit and topic.strip():
            with st.spinner("Inicializando modelos y contexto RAG..."):
                from asistente_iso import retrieve_context, init_interview
                apply_model_configs()
                rag_index   = get_rag_index()
                rag_context = retrieve_context(topic.strip(), rag_index)
                chat, log   = init_interview(topic.strip())

            st.session_state.update({
                "fase":        "entrevista",
                "modo":        "nuevo",
                "topic":       topic.strip(),
                "rag_context": rag_context,
                "chat":        chat,
                "log":         log,
                "mensajes":    [{"role": "assistant", "content": log[-1]["text"]}],
            })
            st.rerun()

        elif submit:
            st.warning("Escribe una descripción del procedimiento antes de continuar.")

    else:  # Revisar procedimiento existente
        st.markdown("""
        <div class="card card-accent-left">
          <div class="card-label">Revisión de procedimiento</div>
          <div class="card-sub" style="color:#dde3f0;font-size:0.875rem;line-height:1.6">
            Sube el .docx del procedimiento existente. El asistente lo leerá
            y te guiará para introducir los cambios de la nueva revisión.
          </div>
        </div>
        """, unsafe_allow_html=True)

        uploaded = st.file_uploader("Procedimiento .docx a revisar", type=["docx"])

        if uploaded:
            col_a, col_b = st.columns([3, 1])
            with col_b:
                if st.button("Iniciar revisión →", type="primary", use_container_width=True):
                    with st.spinner("Leyendo procedimiento y cargando contexto RAG..."):
                        from asistente_iso import extract_text_from_docx_bytes, retrieve_context, init_revision_chat
                        apply_model_configs()
                        docx_text   = extract_text_from_docx_bytes(uploaded.read())
                        rag_index   = get_rag_index()
                        rag_context = retrieve_context("revisión de procedimiento existente", rag_index)
                        chat, log   = init_revision_chat(docx_text)

                    st.session_state.update({
                        "fase":               "revision_chat",
                        "modo":               "revision",
                        "docx_original_text": docx_text,
                        "rag_context":        rag_context,
                        "revision_chat":      chat,
                        "revision_log":       log,
                        "revision_mensajes":  [{"role": "assistant", "content": log[-1]["text"]}],
                    })
                    st.rerun()

# ── FASE 2 — ENTREVISTA ───────────────────────────────────────
elif st.session_state.fase == "entrevista":

    st.markdown(
        f'<div class="card card-accent-left">'
        f'<div class="card-label">Procedimiento en curso</div>'
        f'<div class="card-title">{st.session_state.topic}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    for msg in st.session_state.mensajes:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_input = st.chat_input("Tu respuesta...")

    if user_input:
        st.session_state.mensajes.append({"role": "user", "content": user_input})
        st.session_state.log.append({"role": "user", "text": user_input})

        with st.spinner(""):
            response = st.session_state.chat.send_message(user_input)
            reply    = response.text

        st.session_state.log.append({"role": "assistant", "text": reply})

        if "RECOPILADO" in reply:
            reply_clean = reply.replace("RECOPILADO", "").strip()
            if reply_clean:
                st.session_state.mensajes.append({"role": "assistant", "content": reply_clean})
            st.session_state.fase = "redactando"
        else:
            st.session_state.mensajes.append({"role": "assistant", "content": reply})

        st.rerun()

# ── FASE REVISIÓN — CHAT DE CAMBIOS ───────────────────────────
elif st.session_state.fase == "revision_chat":

    st.markdown(
        '<div class="card card-accent-left">'
        '<div class="card-label">Revisión en curso</div>'
        '<div class="card-sub" style="color:#dde3f0;font-size:0.875rem">'
        'Describe los cambios que quieres introducir. El asistente los confirmará sección a sección.'
        '</div></div>',
        unsafe_allow_html=True,
    )

    for msg in st.session_state.revision_mensajes:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    user_input = st.chat_input("Describe los cambios...")

    if user_input:
        st.session_state.revision_mensajes.append({"role": "user", "content": user_input})
        st.session_state.revision_log.append({"role": "user", "text": user_input})

        with st.spinner(""):
            response = st.session_state.revision_chat.send_message(user_input)
            reply    = response.text

        st.session_state.revision_log.append({"role": "assistant", "text": reply})

        if "CAMBIOS_RECOPILADOS" in reply:
            reply_clean = reply.replace("CAMBIOS_RECOPILADOS", "").strip()
            if reply_clean:
                st.session_state.revision_mensajes.append({"role": "assistant", "content": reply_clean})
            st.session_state.fase = "revision_redactando"
        else:
            st.session_state.revision_mensajes.append({"role": "assistant", "content": reply})

        st.rerun()

# ── FASE REVISIÓN — REDACTANDO ─────────────────────────────────
elif st.session_state.fase == "revision_redactando":

    for msg in st.session_state.revision_mensajes[-3:]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="card card-accent-left">
      <div class="card-label">Procesando revisión</div>
      <div class="card-sub" style="color:#dde3f0;font-size:0.875rem">
        Gemini Pro está aplicando los cambios acordados y generando la nueva revisión del procedimiento.
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Generando procedimiento revisado..."):
        from asistente_iso import draft_revision_with_gemini_pro, extract_json, complete_with_defaults
        apply_model_configs()
        from json_a_ficha import generar_ficha

        transcript = "\n".join(
            f"{'Consultor' if m['role'] == 'assistant' else 'Usuario'}: {m['text']}"
            for m in st.session_state.revision_log
        )
        draft = draft_revision_with_gemini_pro(
            transcript,
            st.session_state.docx_original_text,
            st.session_state.rag_context,
        )
        data = extract_json(draft)

    if data:
        data = complete_with_defaults(data)
        tmp = tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", dir=WORKDIR, delete=False, encoding="utf-8",
        )
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp.close()
        with st.spinner("Generando documento Word revisado..."):
            doc_path = generar_ficha(tmp.name)
        os.unlink(tmp.name)
        append_historial(data, doc_path)
        st.session_state.data     = data
        st.session_state.doc_path = doc_path
        st.session_state.fase     = "listo"
        st.rerun()
    else:
        st.error("No se pudo extraer el JSON. Vuelve a intentarlo.")
        st.session_state.fase = "revision_chat"
        st.rerun()

# ── FASE 3 — REDACTANDO ───────────────────────────────────────
elif st.session_state.fase == "redactando":

    for msg in st.session_state.mensajes[-3:]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="card card-accent-left">
      <div class="card-label">Procesando</div>
      <div class="card-sub" style="color:#dde3f0;font-size:0.875rem">
        Gemini Pro está redactando el procedimiento completo y generando el diagrama de flujo.
        Esto puede tardar entre 30 y 60 segundos.
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Redactando procedimiento ISO..."):
        from asistente_iso import draft_with_gemini_pro, extract_json, complete_with_defaults
        apply_model_configs()
        from json_a_ficha import generar_ficha

        transcript = "\n".join(
            f"{'Consultor' if m['role'] == 'assistant' else 'Usuario'}: {m['text']}"
            for m in st.session_state.log
        )
        draft = draft_with_gemini_pro(transcript, st.session_state.rag_context)
        data  = extract_json(draft)

    if data:
        data = complete_with_defaults(data)
        tmp = tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", dir=WORKDIR,
            delete=False, encoding="utf-8",
        )
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp.close()

        with st.spinner("Generando documento Word y diagrama..."):
            doc_path = generar_ficha(tmp.name)
        os.unlink(tmp.name)

        append_historial(data, doc_path)
        st.session_state.data     = data
        st.session_state.doc_path = doc_path
        st.session_state.fase     = "listo"
        st.rerun()
    else:
        st.error("No se pudo extraer el JSON del borrador. Vuelve a intentarlo.")
        st.session_state.fase = "entrevista"
        st.rerun()

# ── FASE 4 — LISTO ────────────────────────────────────────────
elif st.session_state.fase == "listo":

    doc_path = st.session_state.doc_path
    fname    = os.path.basename(doc_path)

    st.markdown(f"""
    <div class="card card-accent-green">
      <div class="card-label">Documento generado</div>
      <div class="card-title">{fname}</div>
      <div class="card-sub">
        Incluye header/footer corporativo · índice · desarrollo · diagrama de flujo Mermaid
      </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        with open(doc_path, "rb") as f:
            st.download_button(
                label="Descargar .docx",
                data=f,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
            )

    with col2:
        if st.button("Nuevo procedimiento", use_container_width=True, type="secondary"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    # ── Regeneración de secciones ──────────────────────────────
    st.markdown("---")
    st.markdown(
        '<div class="card-label" style="margin-bottom:1rem">Ajustar secciones</div>',
        unsafe_allow_html=True,
    )

    SECCIONES_REGEN = {
        "objeto":            "Objeto",
        "alcance":           "Alcance",
        "responsabilidades": "Responsabilidades",
        "desarrollo":        "Desarrollo",
        "archivo":           "Archivo",
    }

    data = st.session_state.data

    for sec_key, sec_label in SECCIONES_REGEN.items():
        with st.expander(sec_label):
            val = data.get(sec_key, "")
            if isinstance(val, str):
                st.markdown(
                    f'<div style="font-size:0.85rem;color:#dde3f0;line-height:1.6">{val}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.json(val)

            instruction = st.text_input(
                "Indicación (opcional)",
                placeholder="Ej: incluir referencia al sistema AHORA",
                key=f"regen_instr_{sec_key}",
            )

            if st.button(f"Regenerar {sec_label}", key=f"btn_regen_{sec_key}", type="secondary"):
                with st.spinner(f"Regenerando {sec_label}..."):
                    from asistente_iso import regenerate_section
                    apply_model_configs()
                    new_val = regenerate_section(sec_key, data, instruction)

                if new_val is not None:
                    st.session_state.data[sec_key] = new_val
                    tmp = tempfile.NamedTemporaryFile(
                        mode="w", suffix=".json", dir=WORKDIR,
                        delete=False, encoding="utf-8",
                    )
                    json.dump(st.session_state.data, tmp, ensure_ascii=False, indent=2)
                    tmp.close()
                    with st.spinner("Actualizando documento Word..."):
                        from json_a_ficha import generar_ficha
                        new_doc_path = generar_ficha(tmp.name)
                    os.unlink(tmp.name)
                    append_historial(st.session_state.data, new_doc_path)
                    st.session_state.doc_path = new_doc_path
                    st.success(f"{sec_label} regenerado.")
                    st.rerun()
                else:
                    st.error(f"No se pudo regenerar {sec_label}. Vuelve a intentarlo.")
