"""
asistente_iso.py  ·  v3
Asistente conversacional ISO 9001/14001 para GÓMEZ Y CRESPO S.A.

Pipeline multi-modelo (todo Vertex AI):
  1. text-embedding-004   (Vertex AI) → RAG sobre procedimientos existentes
  2. gemini-1.5-flash-002 (Vertex AI) → entrevista iterativa, muchos turnos cortos
  3. gemini-1.5-pro-002   (Vertex AI) → redacción ISO de calidad, un único llamado
  4. json_a_ficha.py                  → genera el .docx formateado

Requisitos previos:
  pip install -r requirements.txt
  gcloud auth application-default login
  Ajusta PROJECT_ID y LOCATION abajo.
"""

import glob
import io
import json
import math
import os
import re
import sys
import tempfile

import vertexai
from vertexai.generative_models import GenerativeModel, ChatSession, GenerationConfig
from vertexai.language_models import TextEmbeddingModel
from docx import Document as DocxReader
from pypdf import PdfReader

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────────────────────

PROJECT_ID  = "documentacion-iso"  # GCP Project ID
LOCATION    = "us-central1"

# Entrevista: Flash es barato para muchos turnos cortos
CHAT_MODEL  = "gemini-2.5-flash"   # razonamiento y redacción colaborativa

# Redacción final: Pro para mayor calidad de escritura en un único llamado
# Alternativa más barata: "gemini-2.0-flash-001"
DRAFT_MODEL = "gemini-2.5-pro"

EMBED_MODEL = "text-embedding-004"

# Procedimientos a excluir del índice RAG (nombres de archivo exactos)
EXCLUDE_FROM_RAG = {"PC-02.docx"}

# ──────────────────────────────────────────────────────────────────────────────
# PARÁMETROS DE GENERACIÓN
# ──────────────────────────────────────────────────────────────────────────────

# Gemini Flash — entrevista conversacional
# temperature alta → respuestas más variadas y naturales
# top_k / top_p controlan el muestreo del vocabulario
CHAT_CONFIG = GenerationConfig(
    temperature=0.7,        # 0.0 = determinista, 1.0 = muy creativo
    top_k=40,               # considera los 40 tokens más probables en cada paso
    top_p=0.95,             # nucleus sampling: masa de probabilidad acumulada
    max_output_tokens=2048,
)

# Gemini Pro — redacción ISO final
# temperature baja → texto más preciso, consistente y fiel al estilo de referencia
DRAFT_CONFIG = GenerationConfig(
    temperature=0.3,        # más bajo = más fiel al estilo ISO existente
    top_k=20,
    top_p=0.85,
    max_output_tokens=8192,
)

WORKDIR    = os.path.dirname(os.path.abspath(__file__))
BC_DIR     = os.path.join(WORKDIR, "BC")
INDEX_FILE = os.path.join(WORKDIR, "rag_index.json")

# Campos fijos de GYC — no se preguntan, se añaden automáticamente
DEFAULTS = {
    "revision":      "00",
    "paginas":       5,
    "elaborado_por": "Responsable de Calidad y Medio Ambiente",
    "aprobado_por":  "Gerencia",
}

# ──────────────────────────────────────────────────────────────────────────────
# RAG — extracción, indexación y recuperación
# ──────────────────────────────────────────────────────────────────────────────

def extract_text_from_docx(path: str) -> str:
    try:
        doc = DocxReader(path)
        return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        print(f"  [RAG] No se pudo leer {os.path.basename(path)}: {e}")
        return ""


def extract_text_from_pdf(path: str) -> str:
    try:
        reader = PdfReader(path)
        pages  = [page.extract_text() or "" for page in reader.pages]
        return "\n".join(p.strip() for p in pages if p.strip())
    except Exception as e:
        print(f"  [RAG] No se pudo leer {os.path.basename(path)}: {e}")
        return ""


def extract_text_from_md(path: str) -> str:
    try:
        with open(path, encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        print(f"  [RAG] No se pudo leer {os.path.basename(path)}: {e}")
        return ""


def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    if ext == ".docx":
        return extract_text_from_docx(path)
    if ext == ".md":
        return extract_text_from_md(path)
    return ""


def chunk_text(text: str, size: int = 250, overlap: int = 40) -> list[str]:
    words = text.split()
    chunks, i = [], 0
    while i < len(words):
        chunk = " ".join(words[i : i + size])
        if chunk:
            chunks.append(chunk)
        i += size - overlap
    return chunks


def cosine_similarity(a: list[float], b: list[float]) -> float:
    dot    = sum(x * y for x, y in zip(a, b))
    norm_a = math.sqrt(sum(x * x for x in a))
    norm_b = math.sqrt(sum(x * x for x in b))
    return dot / (norm_a * norm_b) if norm_a and norm_b else 0.0


def collect_files() -> list[str]:
    """Recoge todos los archivos a indexar de WORKDIR y BC_DIR."""
    files = []

    # Procedimientos del directorio raíz (PC-XX.docx, excluye outputs y exclusiones)
    for f in glob.glob(os.path.join(WORKDIR, "PC-*.docx")):
        if not re.search(r"PC-\d{2,3}_.+\.docx$", f) \
                and os.path.basename(f) not in EXCLUDE_FROM_RAG:
            files.append(f)

    # Carpeta BC: .docx, .pdf y .md, excluye plantillas y archivos temporales de Word
    if os.path.isdir(BC_DIR):
        for ext in ("*.docx", "*.pdf", "*.md"):
            for f in glob.glob(os.path.join(BC_DIR, ext)):
                name = os.path.basename(f)
                if name.startswith("~$"):          # archivo temporal de Word
                    continue
                if "PLANTILLA" in name.upper():    # plantilla vacía
                    continue
                files.append(f)

    return sorted(files)


def build_rag_index() -> list[dict]:
    print("[RAG] Construyendo índice de documentos...")

    files = collect_files()
    if not files:
        print("[RAG] No se encontraron documentos. El asistente funcionará sin contexto.\n")
        return []

    embed_model = TextEmbeddingModel.from_pretrained(EMBED_MODEL)
    index = []

    for path in files:
        text = extract_text(path)
        if not text:
            continue
        chunks = chunk_text(text)
        label  = os.path.relpath(path, WORKDIR).replace("\\", "/")
        print(f"  {label}: {len(chunks)} chunks")

        for i in range(0, len(chunks), 10):
            batch      = chunks[i : i + 10]
            embeddings = embed_model.get_embeddings(batch)
            for chunk, emb in zip(batch, embeddings):
                index.append({
                    "source":    label,
                    "text":      chunk,
                    "embedding": emb.values,
                })

    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)

    print(f"[RAG] Indice guardado: {len(index)} chunks -> {INDEX_FILE}\n")
    return index


def load_or_build_index() -> list[dict]:
    if os.path.exists(INDEX_FILE):
        print("[RAG] Cargando índice existente (borra rag_index.json para reconstruir)...")
        with open(INDEX_FILE, encoding="utf-8") as f:
            return json.load(f)
    return build_rag_index()


def retrieve_context(query: str, index: list[dict], top_k: int = 6) -> str:
    if not index:
        return ""

    embed_model = TextEmbeddingModel.from_pretrained(EMBED_MODEL)
    q_emb       = embed_model.get_embeddings([query])[0].values

    scored = [(cosine_similarity(q_emb, item["embedding"]), item) for item in index]
    scored.sort(key=lambda x: x[0], reverse=True)

    lines = ["=== Fragmentos de procedimientos existentes de GYC ===\n"]
    for score, item in scored[:top_k]:
        lines.append(f"[{item['source']} · similitud {score:.2f}]\n{item['text']}\n")
    return "\n".join(lines)


# ──────────────────────────────────────────────────────────────────────────────
# FASE 1 — GEMINI FLASH: entrevista iterativa (barato)
# ──────────────────────────────────────────────────────────────────────────────

INTERVIEW_SYSTEM = """\
Eres un consultor experto en calidad ISO que ayuda a documentar procedimientos para
GÓMEZ Y CRESPO S.A. (fabricante de equipamiento agroganadero, ISO 9001 e ISO 14001,
ERP/CRM: AHORA, sede en Ourense). Conoces la empresa y sus procesos.

Tu forma de trabajar es colaborativa: no haces una lista de preguntas genéricas,
sino que para cada sección PROPONES un texto redactado basándote en lo que ya sabes
de GYC, y luego preguntas si es correcto o hay que ajustarlo. Si el usuario dice
"sí" o "vale", lo das por bueno y pasas a la siguiente sección.

Reglas:
- Trata UNA sección cada vez. No avances hasta confirmar la anterior.
- Para cada sección: propón un borrador concreto y pregunta "¿Es así, o lo ajustamos?".
- Si necesitas un dato que no puedes inferir (un cargo específico, un número de registro,
  un plazo concreto), pregúntalo directamente antes de proponer el texto.
- Sé específico: menciona AHORA cuando sea el sistema usado, usa los cargos reales
  de GYC (Responsable de Calidad y Medio Ambiente, Gerencia, Responsable de Compras…),
  cita los procedimientos relacionados que correspondan.
- No menciones explícitamente la norma ISO ni sus cláusulas en la conversación;
  el cumplimiento normativo ya está garantizado por el contexto.
- Cuando propongas texto, imita el estilo, tono y nivel de detalle de los
  procedimientos existentes que tienes en el contexto RAG: misma estructura de frases,
  mismo vocabulario técnico, misma forma de describir las acciones.
- Si el usuario amplía o corrige algo, incorpora el cambio y confirma el texto final.
- Cuando tengas TODAS las secciones confirmadas (código, nombre, fecha, objeto, alcance,
  responsabilidades, desarrollo, archivo, referencias, anexos), escribe únicamente:

RECOPILADO

No escribas nada más tras RECOPILADO. No redactes el documento completo.\
"""


def init_interview(topic: str) -> tuple[ChatSession, list[dict]]:
    """Inicia la sesión de entrevista con Gemini Flash."""
    model = GenerativeModel(
        CHAT_MODEL,
        system_instruction=INTERVIEW_SYSTEM,
        generation_config=CHAT_CONFIG,
    )
    chat = model.start_chat()
    log  = []

    response = chat.send_message(
        f"Quiero documentar el siguiente procedimiento: {topic}"
    )
    log.append({"role": "assistant", "text": response.text})
    return chat, log


# ──────────────────────────────────────────────────────────────────────────────
# FASE 2 — CLAUDE HAIKU: redacción ISO de calidad (un único llamado)
# ──────────────────────────────────────────────────────────────────────────────

DRAFT_SYSTEM = """\
Eres el Redactor Jefe de Procedimientos ISO de GÓMEZ Y CRESPO S.A., empresa española
fabricante de equipamiento agroganadero (avicultura, cunicultura, minifundio),
certificada ISO 9001:2015 e ISO 14001:2015, sede en Ourense.
ERP/CRM corporativo: AHORA. Elabora el Responsable de Calidad y Medio Ambiente; aprueba Gerencia.

Tu misión es redactar el procedimiento con un nivel de detalle intermedio: suficiente para
entender el flujo y las responsabilidades, pero sin entrar en operativa interna específica.

Reglas de redacción:
- Escribe en español formal ISO. Frases claras y directas.
- Cada apartado del desarrollo debe explicar QUÉ se hace, QUIÉN lo hace y el resultado esperado.
  3-4 frases por apartado. Desarrolla con lógica y coherencia, pero sin inventar pasos
  operativos concretos, plazos exactos ni datos que no hayan aparecido en la entrevista.
- Si necesitas completar algo no mencionado, usa fórmulas genéricas del estilo ISO
  ("según corresponda", "de acuerdo con los criterios establecidos", "cuando proceda").
- No añadas requisitos normativos, cláusulas ni referencias que no estén en la entrevista.\
"""

_CLAUDE_PROMPT_TPL = """\
A continuación tienes el contexto de los procedimientos existentes y la transcripción
de la entrevista donde se recopiló la información para el nuevo procedimiento.

--- PROCEDIMIENTOS EXISTENTES (contexto RAG) ---
{rag_context}

--- TRANSCRIPCIÓN DE LA ENTREVISTA ---
{transcript}

--- INSTRUCCIONES ---
Redacta el procedimiento completo en español formal ISO 9001. Usa los mismos cargos,
referencias y terminología que aparecen en los procedimientos existentes.
El campo "historial" debe reflejar la fecha real del procedimiento.
Nivel de detalle: explica cada paso con claridad pero de forma general. No inventes
operativa específica, plazos ni datos que no hayan sido mencionados en la transcripción.

Cuando hayas terminado de redactar, escribe exactamente:

FINALIZADO

E inmediatamente después, sin ningún texto adicional, el bloque JSON con este
esquema exacto (sin cambiar ninguna clave):

```json
{{
  "codigo": "PC-XX",
  "nombre": "NOMBRE DEL PROCEDIMIENTO EN MAYÚSCULAS",
  "fecha": "DD/MM/AA",
  "historial": [
    {{
      "rev": "00",
      "fecha": "DD/MM/AA",
      "descripcion": "Nuevo lanzamiento documental en revisión 00",
      "revisado": "",
      "elaborado": ""
    }}
  ],
  "objeto": "Texto redactado del objeto...",
  "alcance": "Texto redactado del alcance...",
  "responsabilidades": [
    {{
      "cargo": "Nombre del cargo",
      "tareas": ["Tarea 1 redactada.", "Tarea 2 redactada."]
    }}
  ],
  "desarrollo": [
    {{
      "num": "4.1.",
      "titulo": "Título del apartado",
      "descripcion": "Descripción detallada y redactada del apartado."
    }}
  ],
  "archivo": [
    {{
      "documento": "Nombre del registro",
      "responsable": "Cargo custodio",
      "lugar": "Lugar o sistema de archivo"
    }}
  ],
  "referencias": ["PC-02: «Procesos Relacionados con los Clientes»"],
  "anexos": ["Anexo 1, PC-XX: Nombre del anexo"],
  "diagrama_mermaid": "flowchart TD\n    A([Inicio]) --> B[Paso 1]\n    B --> C{{¿Decisión?}}\n    C -- Sí --> D[Paso 2]\n    C -- No --> E[Corrección]\n    E --> B\n    D --> F([Fin])"
}}
```

El campo "diagrama_mermaid" debe contener un diagrama flowchart TD en sintaxis Mermaid que represente
fielmente el flujo del procedimiento: usa rectángulos para actividades, rombos para decisiones
({{texto}}), óvalos/estadios para inicio/fin (([texto])). Incluye los cargos responsables como
etiquetas en los nodos cuando sea relevante. El diagrama debe ser coherente con las secciones
"desarrollo" y "responsabilidades".\
"""


def draft_with_gemini_pro(transcript: str, rag_context: str) -> str:
    """Envía la transcripción a Gemini Pro para redactar el procedimiento ISO."""
    print("\n[Gemini Pro] Redactando procedimiento ISO... ", end="", flush=True)

    prompt = _CLAUDE_PROMPT_TPL.format(
        rag_context=rag_context or "No hay procedimientos existentes indexados.",
        transcript=transcript,
    )

    model    = GenerativeModel(
        DRAFT_MODEL,
        system_instruction=DRAFT_SYSTEM,
        generation_config=DRAFT_CONFIG,
    )
    response = model.generate_content(prompt)

    print("OK\n")
    return response.text


# ──────────────────────────────────────────────────────────────────────────────
# EXTRACCIÓN DE JSON Y GENERACIÓN DEL DOCX
# ──────────────────────────────────────────────────────────────────────────────

def extract_json(text: str) -> dict | None:
    m = re.search(r"```json\s*(\{.*?\})\s*```", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    m = re.search(r'(\{.*?"codigo".*?\})', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    return None


def complete_with_defaults(data: dict) -> dict:
    return {**DEFAULTS, **data}


# ──────────────────────────────────────────────────────────────────────────────
# FEATURE 3 — REGENERACIÓN DE SECCIÓN PARCIAL
# ──────────────────────────────────────────────────────────────────────────────

_REGEN_SYSTEM = """\
Eres el Redactor Jefe de Procedimientos ISO de GÓMEZ Y CRESPO S.A.
Tu tarea es regenerar ÚNICAMENTE la sección indicada de un procedimiento existente,
manteniendo coherencia con el resto del documento que se te proporciona como contexto.
No modifiques ni devuelvas ninguna otra sección.
Responde solo con el valor JSON de la sección pedida, sin texto adicional ni bloque de código.\
"""

_REGEN_PROMPT_TPL = """\
Procedimiento actual completo (no modificar el resto):
{current_json}

Sección a regenerar: {section_key}
Instrucción del usuario: {user_instruction}

Devuelve únicamente el valor JSON para la clave "{section_key}".
- objeto / alcance → string de texto
- responsabilidades → lista de objetos {{"cargo": ..., "tareas": [...]}}
- desarrollo → lista de objetos {{"num": ..., "titulo": ..., "descripcion": ...}}
- archivo → lista de objetos {{"documento": ..., "responsable": ..., "lugar": ...}}
\
"""


def regenerate_section(section_key: str, current_data: dict, user_instruction: str = ""):
    """Regenera una sola sección del JSON. Devuelve el nuevo valor o None si falla."""
    prompt = _REGEN_PROMPT_TPL.format(
        current_json=json.dumps(current_data, ensure_ascii=False, indent=2),
        section_key=section_key,
        user_instruction=user_instruction or "Mejora la redacción manteniendo el contenido.",
    )
    model    = GenerativeModel(DRAFT_MODEL, system_instruction=_REGEN_SYSTEM,
                               generation_config=DRAFT_CONFIG)
    response = model.generate_content(prompt)
    raw = response.text.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r'(\[.*\]|\{.*\}|"[^"]*")', raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass
    return None


# ──────────────────────────────────────────────────────────────────────────────
# FEATURE 7 — MODO REVISIÓN DE PROCEDIMIENTO EXISTENTE
# ──────────────────────────────────────────────────────────────────────────────

def extract_text_from_docx_bytes(data: bytes) -> str:
    """Extrae texto de un .docx recibido como bytes (desde st.file_uploader)."""
    try:
        doc = DocxReader(io.BytesIO(data))
        return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        print(f"  [RAG] Error extrayendo bytes docx: {e}")
        return ""


_REVISION_SYSTEM = """\
Eres un consultor ISO que ayuda a revisar y actualizar procedimientos existentes
de GÓMEZ Y CRESPO S.A. Se te proporciona el texto completo de un procedimiento.
Tu misión es entender los cambios que el usuario quiere introducir, confirmando
sección por sección qué debe modificarse y qué debe conservarse.
Sé conciso y directo. Cuando hayas recopilado todos los cambios necesarios, escribe únicamente:

CAMBIOS_RECOPILADOS

No redactes el procedimiento revisado en el chat.\
"""

_REVISION_PROMPT_TPL = """\
--- PROCEDIMIENTO ORIGINAL ---
{original_text}

--- PROCEDIMIENTOS DE REFERENCIA (RAG) ---
{rag_context}

--- CAMBIOS ACORDADOS EN LA REVISIÓN ---
{transcript}

--- INSTRUCCIONES ---
Genera el procedimiento REVISADO completo en formato JSON, aplicando exactamente los cambios
acordados. Incrementa la revisión en el historial (R00→R01, R01→R02, etc.) y añade una
entrada describiendo los cambios aplicados. Mantén sin cambios todo lo que no haya sido
mencionado. Nivel de detalle intermedio: claro pero sin inventar operativa específica.

Cuando termines, escribe exactamente FINALIZADO y a continuación el bloque JSON.\
"""


def init_revision_chat(docx_text: str) -> tuple[ChatSession, list[dict]]:
    """Inicia sesión de chat Flash para recopilar cambios sobre un procedimiento existente."""
    model = GenerativeModel(
        CHAT_MODEL,
        system_instruction=_REVISION_SYSTEM,
        generation_config=CHAT_CONFIG,
    )
    chat = model.start_chat()
    log  = []
    response = chat.send_message(
        f"Este es el procedimiento existente que vamos a revisar:\n\n{docx_text}\n\n"
        f"¿Qué cambios quieres introducir en esta revisión?"
    )
    log.append({"role": "assistant", "text": response.text})
    return chat, log


def draft_revision_with_gemini_pro(transcript: str, original_text: str, rag_context: str) -> str:
    """Genera el procedimiento revisado con los cambios acordados."""
    print("\n[Gemini Pro] Generando revisión ISO... ", end="", flush=True)
    prompt = _REVISION_PROMPT_TPL.format(
        original_text=original_text,
        rag_context=rag_context or "No hay procedimientos indexados.",
        transcript=transcript,
    )
    model    = GenerativeModel(DRAFT_MODEL, system_instruction=DRAFT_SYSTEM,
                               generation_config=DRAFT_CONFIG)
    response = model.generate_content(prompt)
    print("OK\n")
    return response.text


def generate_docx(data: dict) -> None:
    sys.path.insert(0, WORKDIR)
    from json_a_ficha import generar_ficha  # noqa: PLC0415

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", dir=WORKDIR,
        delete=False, encoding="utf-8",
    )
    json.dump(data, tmp, ensure_ascii=False, indent=2)
    tmp.close()

    try:
        out_path = generar_ficha(tmp.name)
        print(f"[OK] Documento generado: {out_path}\n")
    finally:
        os.unlink(tmp.name)


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 62)
    print("  ASISTENTE ISO — GÓMEZ Y CRESPO S.A.")
    print("  Entrevista: Gemini Flash  |  Redacción: Gemini Pro")
    print("  Escribe 'salir' en cualquier momento para cancelar.")
    print("=" * 62 + "\n")

    # Inicializar Vertex AI
    vertexai.init(project=PROJECT_ID, location=LOCATION)

    # RAG
    index = load_or_build_index()

    # Tema inicial
    print("¿Sobre qué procedimiento vas a trabajar hoy? (descripción breve)\n")
    try:
        topic = input("Tú: ").strip()
    except (KeyboardInterrupt, EOFError):
        return

    if topic.lower() == "salir":
        return

    rag_context = retrieve_context(topic, index)
    if rag_context:
        print("\n[RAG] Contexto de procedimientos existentes cargado.\n")

    # ── Fase 1: Gemini Flash — entrevista ─────────────────────────────────────
    try:
        chat, log = init_interview(topic)
    except Exception as e:
        print(f"[ERROR] No se pudo conectar con Vertex AI: {e}")
        sys.exit(1)

    print(f"\nConsultor: {log[-1]['text']}\n")

    while True:
        try:
            user_input = input("Tú: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nSesión cancelada.")
            return

        if not user_input:
            continue
        if user_input.lower() == "salir":
            print("Sesión finalizada.")
            return

        log.append({"role": "user", "text": user_input})
        response = chat.send_message(user_input)
        reply    = response.text
        log.append({"role": "assistant", "text": reply})
        print(f"\nConsultor: {reply}\n")

        if "RECOPILADO" in reply:
            break

    # ── Fase 2: Claude — redacción ISO ────────────────────────────────────────
    transcript = "\n".join(
        f"{'Consultor' if m['role'] == 'assistant' else 'Usuario'}: {m['text']}"
        for m in log
    )

    try:
        draft = draft_with_gemini_pro(transcript, rag_context)
    except Exception as e:
        print(f"[ERROR] Fallo al llamar a Gemini Pro: {e}")
        sys.exit(1)

    # ── Extraer JSON y generar .docx ──────────────────────────────────────────
    data = extract_json(draft)
    if data:
        data = complete_with_defaults(data)
        generate_docx(data)
    else:
        print("[AVISO] No se encontró JSON válido en la respuesta de Claude.")
        fallback = os.path.join(WORKDIR, "ultimo_output.txt")
        with open(fallback, "w", encoding="utf-8") as f:
            f.write(draft)
        print(f"[INFO] Respuesta de Claude guardada en {fallback}")


if __name__ == "__main__":
    main()
