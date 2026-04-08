# asistente.py — Lógica de negocio: RAG + entrevista + redacción ISO
# Usa Google AI Studio (google-genai), NO Vertex AI.

import os
import json
import math
import re
import io
import time
import requests
from dotenv import load_dotenv
from google import genai
from docx import Document as DocxReader
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter

# ── 1. Configuración ───────────────────────────────────────────────────────────

load_dotenv()
client = genai.Client(
    api_key=os.getenv("GOOGLE_API_KEY"),
    http_options={"api_version": "v1"},
)

CHAT_MODEL  = "gemini-2.0-flash"
DRAFT_MODEL = "gemini-1.5-pro-latest"
EMBED_MODEL = "gemini-embedding-001"

# ── 5. System prompt (debe ir antes de las funciones que lo usan) ──────────────

SYSTEM_PROMPT = """
# Instrucciones — GPT Documentador ISO de GÓMEZ Y CRESPO S.A.

Eres un consultor experto en calidad ISO integrado en el sistema documental de GÓMEZ Y CRESPO S.A. (fabricante de equipamiento agroganadero, ISO 9001:2015 e ISO 14001:2015, ERP: AHORA, sede en Ourense). Redactas procedimientos ISO en español formal.

## Contexto de la empresa

- Cargos habituales: Gerencia, Responsable de Calidad y Medio Ambiente, Responsable de Compras, Responsable de Producción, Departamento Técnico, Administración.
- ERP/CRM corporativo: AHORA.
- Elabora siempre: Responsable de Calidad y Medio Ambiente.
- Aprueba siempre: Gerencia.
- No menciones cláusulas ISO en el documento; el cumplimiento normativo ya está implícito.

## Flujo para crear un procedimiento nuevo

1. Consulta los archivos de conocimiento antes de redactar cualquier sección para conocer el estilo, vocabulario y procedimientos relacionados de GYC. Imita ese estilo.
2. Propón el código del procedimiento consultando los archivos de conocimiento para identificar el siguiente código disponible. Pide confirmación.
3. Entrevista colaborativa — trabaja sección por sección en este orden. En cada sección:
   - redacta el texto completo y definitivo tal como aparecerá en el documento,
   - luego pregunta: "¿Es así, o lo ajustamos?"
   - no avances hasta confirmar.

Orden de secciones:
1. Código y nombre
2. Objeto
3. Alcance
4. Definiciones y abreviaturas
5. Responsabilidades
6. Entradas y salidas
7. Desarrollo
8. Archivo
9. Referencias
10. Anexos

## Reglas de redacción

- Usa siempre tercera persona + futuro de obligación.
- Usa tono formal, claro y narrativo.
- Nombra siempre el cargo completo.
- Menciona AHORA cuando sea relevante.
- No inventes datos.
- Usa negritas inline con **texto**.
- En el desarrollo, cada subapartado llevará un subtítulo en negrita como primera frase.
- Durante la entrevista del Desarrollo, pregunta explícitamente:
  - si hay casos alternativos o excepciones,
  - qué documentos o formularios internos se generan o consultan.

## Comportamiento durante la entrevista

- No hagas preguntas genéricas.
- Guía la conversación sección por sección.
- Haz una sola pregunta o bloque corto por turno.
- Si falta información importante, pídela.
- Si el usuario da información incompleta, propón una redacción provisional y pide confirmación.
- Mantén siempre la conversación en español.
"""

# ── 2. RAG — extracción de texto ───────────────────────────────────────────────

def extract_text_from_docx(path: str) -> str:
    doc = DocxReader(path)
    textos = []
    for p in doc.paragraphs:
        if p.text.strip():
            textos.append(p.text)
    return "\n".join(textos)


def extract_text_from_pdf(path: str) -> str:
    pdf = PdfReader(path)
    textos = []
    for t in pdf.pages:
        text = t.extract_text()
        if text:
            textos.append(text)
    return "\n".join(textos)

# ── 3. RAG — chunking y embeddings ────────────────────────────────────────────

splitter = RecursiveCharacterTextSplitter(
    chunk_size=250,
    chunk_overlap=40,
    separators=["\n\n", "\n", " ", ""]
)

def chunking(text: str) -> list[str]:
    return splitter.split_text(text)


def embed_text(text: str) -> list[float]:
    """Llama directamente a la REST API de Google para embeddings."""
    api_key = os.getenv("GOOGLE_API_KEY")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{EMBED_MODEL}:embedContent?key={api_key}"
    payload = {
        "model": f"models/{EMBED_MODEL}",
        "content": {"parts": [{"text": text}]}
    }
    for attempt in range(3):
        response = requests.post(url, json=payload)
        if response.status_code == 429:
            time.sleep(2 ** attempt)  # 1s, 2s, 4s
            continue
        response.raise_for_status()
        return response.json()["embedding"]["values"]
    response.raise_for_status()


def generate_embeddings(chunks: list[str], source: str) -> list[dict]:
    results = []
    for chunk in chunks:
        emb = embed_text(chunk)
        results.append({
            "source": source,
            "text": chunk,
            "embedding": emb
        })
    return results


def cosine_similarity(a: list[float], b: list[float]) -> float:
    dot    = sum(x * y for x, y in zip(a, b))
    norm_a = math.sqrt(sum(x * x for x in a))
    norm_b = math.sqrt(sum(y * y for y in b))
    return dot / (norm_a * norm_b)


def retrieve(query: str, index: list[dict], top_k: int = 5) -> list[str]:
    query_emb = embed_text(query)
    scores = []
    for item in index:
        sim = cosine_similarity(query_emb, item["embedding"])
        scores.append((sim, item["text"]))
    scores.sort(reverse=True)
    return [text for _, text in scores[:top_k]]

# ── 4. RAG — construcción y carga del índice ──────────────────────────────────

def save_index(index: list[dict], path: str = "rag_index.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def load_index(path: str = "rag_index.json") -> list[dict] | None:
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_rag_index(folder_path: str = "base-conocimiento") -> list[dict]:
    index = []
    for filename in os.listdir(folder_path):
        path = os.path.join(folder_path, filename)
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(path)
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(path)
        else:
            continue
        chunks = chunking(text)
        index.extend(generate_embeddings(chunks, source=filename))
    save_index(index)
    return index

# ── 5. Fase entrevista — Gemini Flash ─────────────────────────────────────────

def _build_initial_prompt(topic: str) -> str:
    return f"""
El usuario quiere crear un procedimiento nuevo sobre este proceso:

{topic}

Empieza la entrevista colaborativa.
Haz primero una propuesta breve para la sección "Código y nombre", y después pregunta:
"¿Es así, o lo ajustamos?"

Si todavía no hay datos suficientes para cerrar esa sección, haz la pregunta mínima necesaria para poder proponerla.
""".strip()


def init_interview(topic: str):
    chat = client.chats.create(
        model=CHAT_MODEL,
        config={
            "system_instruction": SYSTEM_PROMPT,
            "temperature": 0.7
        }
    )
    first_input = _build_initial_prompt(topic)
    response = chat.send_message(first_input)
    log = [
        {"role": "user",      "content": topic},
        {"role": "assistant", "content": response.text}
    ]
    return chat, log


def continue_interview(chat, user_message: str, log: list):
    response = chat.send_message(user_message)
    log.append({"role": "user",      "content": user_message})
    log.append({"role": "assistant", "content": response.text})
    return response.text, log


def transcript_from_log(log: list) -> str:
    lines = []
    for turn in log:
        prefix = "Usuario" if turn["role"] == "user" else "Asistente"
        lines.append(f"{prefix}: {turn['content']}")
    return "\n".join(lines)

# ── 6. Fase redacción — Gemini Pro ────────────────────────────────────────────

DRAFT_SYSTEM_PROMPT = """
Eres el Redactor Jefe de Procedimientos ISO de GÓMEZ Y CRESPO S.A.
Tu única misión es convertir la transcripción de una entrevista en un procedimiento ISO completo en formato JSON.

Reglas:
- Escribe en español formal ISO. Tercera persona, futuro de obligación.
- Nivel de detalle intermedio: explica qué se hace, quién lo hace y el resultado esperado.
- No inventes datos que no aparezcan en la transcripción.
- Usa los cargos reales de GYC y menciona AHORA cuando sea relevante.
- El diagrama_mermaid debe representar fielmente el flujo del procedimiento.

Cuando termines de redactar escribe exactamente:

FINALIZADO

E inmediatamente después el bloque JSON, sin texto adicional.
"""

_DRAFT_PROMPT_TPL = """\
--- CONTEXTO: PROCEDIMIENTOS EXISTENTES DE GYC ---
{rag_context}

--- TRANSCRIPCIÓN DE LA ENTREVISTA ---
{transcript}

--- INSTRUCCIONES ---
Redacta el procedimiento completo en formato JSON con esta estructura exacta:

```json
{{
  "codigo": "PC-XX",
  "nombre": "NOMBRE EN MAYÚSCULAS",
  "fecha": "DD/MM/AA",
  "revision": "00",
  "paginas": 5,
  "elaborado_por": "Responsable de Calidad y Medio Ambiente",
  "aprobado_por": "Gerencia",
  "historial": [
    {{
      "rev": "00",
      "fecha": "DD/MM/AA",
      "descripcion": "Nuevo lanzamiento documental en revisión 00",
      "revisado": "",
      "elaborado": ""
    }}
  ],
  "objeto": "...",
  "alcance": "...",
  "responsabilidades": [
    {{"cargo": "Nombre del cargo", "tareas": ["Tarea 1.", "Tarea 2."]}}
  ],
  "desarrollo": [
    {{"num": "4.1.", "titulo": "Título del apartado", "descripcion": "Descripción."}}
  ],
  "archivo": [
    {{"documento": "Nombre del registro", "responsable": "Cargo", "lugar": "Lugar"}}
  ],
  "referencias": ["PC-02: «Procesos Relacionados con los Clientes»"],
  "anexos": ["Anexo 1, PC-XX: Nombre del anexo"],
  "diagrama_mermaid": "flowchart TD\\n    A([Inicio]) --> B[Paso 1]\\n    B --> C([Fin])"
}}
```
"""


def draft_procedure(transcript: str, rag_context: str = "") -> str:
    prompt = _DRAFT_PROMPT_TPL.format(
        rag_context=rag_context or "No hay procedimientos existentes indexados.",
        transcript=transcript,
    )
    response = client.models.generate_content(
        model=DRAFT_MODEL,
        contents=prompt,
        config={
            "system_instruction": DRAFT_SYSTEM_PROMPT,
            "temperature": 0.3,
        }
    )
    return response.text


# ── 7. Extracción de JSON y generación de .docx ───────────────────────────────

DEFAULTS = {
    "revision":      "00",
    "paginas":       5,
    "elaborado_por": "Responsable de Calidad y Medio Ambiente",
    "aprobado_por":  "Gerencia",
}


def extract_json(text: str) -> dict | None:
    """Extrae el primer bloque ```json ... ``` del texto y lo parsea."""
    m = re.search(r"```json\s*(\{.*?\})\s*```", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    # Fallback: busca cualquier { } si no hay bloque marcado
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return None


def add_defaults(data: dict) -> dict:
    """Rellena campos fijos de GYC que no se preguntan en la entrevista."""
    return {**DEFAULTS, **data}


def generate_docx(data: dict) -> str:
    """Guarda el JSON en un archivo temporal y llama a json_a_ficha.generar_ficha()."""
    import tempfile
    from json_a_ficha import generar_ficha

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", delete=False, encoding="utf-8"
    ) as tmp:
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp_path = tmp.name

    try:
        out_path = generar_ficha(tmp_path)
    finally:
        os.unlink(tmp_path)

    return out_path
