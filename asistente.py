# asistente.py — Lógica de negocio: RAG + entrevista + redacción ISO
# Usa Google AI Studio (google-genai), NO Vertex AI.

import os
import json
import time
from datetime import datetime, timezone
import numpy as np
from dotenv import load_dotenv
from google import genai
from docx import Document as DocxReader
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from pydantic import BaseModel


# ── Esquema de documento ISO (contrato de salida forzado) ─────────────────────

class HistorialEntry(BaseModel):
    rev: str
    fecha: str
    descripcion: str
    revisado: str = ""
    elaborado: str = ""

class Definicion(BaseModel):
    termino: str
    descripcion: str

class Responsabilidad(BaseModel):
    cargo: str
    tareas: list[str]

class ApartadoDesarrollo(BaseModel):
    num: str
    titulo: str
    descripcion: str

class RegistroArchivo(BaseModel):
    documento: str
    responsable: str
    lugar: str

class ProcedimientoBase(BaseModel):
    """Schema para el paso 1: todo excepto responsabilidades."""
    codigo: str
    nombre: str
    fecha: str
    revision: str = "00"
    paginas: int = 5
    elaborado_por: str = "Responsable de Calidad"
    aprobado_por: str = "Gerencia"
    historial: list[HistorialEntry]
    objeto: str
    alcance: str
    definiciones: list[Definicion]
    desarrollo: list[ApartadoDesarrollo]
    archivo: list[RegistroArchivo]
    referencias: list[str]
    anexos: list[str]
    diagrama_mermaid: str

class ListaResponsabilidades(BaseModel):
    """Schema para el paso 2: responsabilidades derivadas del Desarrollo."""
    responsabilidades: list[Responsabilidad]

class Procedimiento(ProcedimientoBase):
    """Schema completo: base + responsabilidades."""
    responsabilidades: list[Responsabilidad]

# ── 1. Configuración ───────────────────────────────────────────────────────────

load_dotenv()
client = genai.Client(
    api_key=os.getenv("GOOGLE_API_KEY"),
    http_options={"api_version": "v1beta"},
)

CHAT_MODEL  = "gemini-2.5-flash"
DRAFT_MODEL = "gemini-2.5-pro"
EMBED_MODEL = "gemini-embedding-001"

# ── 5. System prompt (debe ir antes de las funciones que lo usan) ──────────────

SYSTEM_PROMPT = """
# Instrucciones — GPT Documentador ISO de GÓMEZ Y CRESPO S.A.

Eres un consultor experto en calidad ISO integrado en el sistema documental de GÓMEZ Y CRESPO S.A. (fabricante de equipamiento agroganadero, ISO 9001:2015 e ISO 14001:2015, sede en Ourense). Redactas procedimientos ISO en español formal.

## Contexto de la empresa

- Cargos habituales: Gerencia, Responsable de Calidad, Responsable de Compras, Responsable de Producción, Departamento Técnico, Administración.
- ERP/CRM corporativo: ERP (no menciones el nombre comercial del ERP, refiérete siempre a él como "el ERP").
- Elabora siempre: Responsable de Calidad.
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

### Instrucciones específicas para la sección Responsabilidades

La sección Responsabilidades se redacta **a partir del Desarrollo**: extrae de él los cargos que aparecen mencionados explícitamente y describe sus tareas según lo que el propio Desarrollo ya establece. No añadas cargos genéricos (p. ej. "Todo el Personal") salvo que estén mencionados de forma específica en el Desarrollo.

Cuando trabajes la sección Responsabilidades:
- Incluye únicamente los cargos que aparezcan mencionados de forma explícita en el Desarrollo confirmado.
- Para cada cargo, redacta sus tareas extrayéndolas directamente de lo que el Desarrollo describe para ese cargo.
- Nunca dejes un cargo sin al menos dos tareas concretas; si el Desarrollo solo menciona una acción para ese cargo, profundiza preguntando al usuario antes de redactar.
- No añadas cargos "de relleno" ni responsabilidades genéricas que no estén respaldadas por el Desarrollo.

## Reglas de redacción

- Usa siempre tercera persona + futuro de obligación.
- Usa tono formal, claro y narrativo.
- Nombra siempre el cargo completo.
- Menciona el ERP cuando sea relevante, pero nunca por su nombre comercial.
- No inventes datos.
- Usa negritas inline con **texto**.
- En el desarrollo, cada subapartado llevará un subtítulo en negrita como primera frase, seguido de texto narrativo en párrafos. No uses listas de puntos ni guiones dentro del desarrollo; redacta siempre en prosa continua.
- Durante la entrevista del Desarrollo, pregunta explícitamente:
  - si hay casos alternativos o excepciones,
  - qué documentos o formularios internos se generan o consultan,
  - los plazos o frecuencias relevantes,
  - los criterios de aceptación o rechazo si aplica,
  - qué ocurre si el proceso falla o hay una incidencia.

## Comportamiento durante la entrevista

- No hagas preguntas genéricas.
- Guía la conversación sección por sección.
- En cada sección propón siempre un borrador completo antes de preguntar.
- Haz preguntas de profundización: "¿En qué plazo?", "¿Quién valida?", "¿Qué registro queda?", "¿Hay excepciones?".
- Si el usuario confirma sin añadir nada, pregunta al menos una cosa más antes de avanzar para asegurarte de que no falta detalle.
- Si falta información importante, pídela con una pregunta concreta.
- Si el usuario da información incompleta, propón una redacción provisional detallada y pide confirmación.
- Mantén siempre la conversación en español.

## Finalización

Cuando todas las secciones estén confirmadas por el usuario, escribe exactamente en una línea:

FINALIZADO

No añadas ningún texto después de esa palabra.
"""

SYSTEM_PROMPT_EXPRESS = """
# Instrucciones — Modo Express · GPT Documentador ISO de GÓMEZ Y CRESPO S.A.

Eres un consultor experto en calidad ISO de GÓMEZ Y CRESPO S.A. (fabricante de equipamiento agroganadero, ISO 9001:2015 e ISO 14001:2015, sede en Ourense). Elabora siempre: Responsable de Calidad. Aprueba siempre: Gerencia. ERP corporativo: refiérete a él como "el ERP". No menciones cláusulas ISO.

## Tu misión en modo express

Recopilar toda la información necesaria en el mínimo de turnos posible, luego cerrar la entrevista.

## Flujo

1. En tu primer mensaje, lanza un bloque de preguntas clave que cubra todas las secciones del procedimiento:
   - ¿Cuál es el objeto y alcance del procedimiento?
   - ¿Quién o quiénes son los responsables principales?
   - ¿Cuáles son los pasos principales del proceso, de inicio a fin?
   - ¿Qué documentos, registros o formularios se generan o consultan?
   - ¿Hay casos especiales, excepciones o situaciones de fallo relevantes?
   - ¿Hay referencias a otros procedimientos o normativa interna?
   - ¿Hay anexos?

2. Con la respuesta del usuario, redacta un borrador completo de todas las secciones y pregunta: "¿Lo dejamos así o ajustamos algo?"

3. Aplica los ajustes si los hay, confirma y escribe:

FINALIZADO

No añadas ningún texto después de esa palabra.

## Reglas de redacción

- Tercera persona, futuro de obligación, tono formal.
- Nombra siempre el cargo completo.
- No inventes datos.
"""

# ── 2. RAG — extracción de texto ───────────────────────────────────────────────

def extract_text_from_docx(path: str) -> str:
    doc = DocxReader(path)
    textos = []
    for p in doc.paragraphs:
        if p.text.strip():
            textos.append(p.text)
    # Incluye texto de tablas (responsables, registros, pasos…)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    textos.append(cell.text.strip())
    return "\n".join(textos)


def extract_text_from_pdf(path: str) -> str:
    try:
        import pdfplumber
        textos = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=3, y_tolerance=3)
                if text:
                    textos.append(text)
                # Preserva tablas como filas pipe-separadas
                for table in page.extract_tables():
                    for row in table:
                        cells = [str(c).strip() if c else "" for c in row]
                        row_text = " | ".join(c for c in cells if c)
                        if row_text:
                            textos.append(row_text)
        return "\n".join(textos)
    except Exception:
        # Fallback a PyPDF si pdfplumber falla en algún PDF corrupto
        pdf = PdfReader(path)
        return "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())


def extract_text_from_md(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def extract_text_from_doc(path: str) -> str:
    import subprocess
    import docx2txt
    try:
        return docx2txt.process(path) or ""
    except Exception:
        # Formato .doc binario (OLE) — usar antiword
        try:
            result = subprocess.run(
                ["antiword", path], capture_output=True, text=True, timeout=30
            )
            return result.stdout or ""
        except Exception:
            return ""


def index_single_file(path: str, filename: str) -> list[dict]:
    """Extrae texto, trocea y genera embeddings para un único archivo."""
    try:
        if filename.endswith(".pdf"):
            text = extract_text_from_pdf(path)
        elif filename.endswith(".docx"):
            text = extract_text_from_docx(path)
        elif filename.endswith(".doc"):
            text = extract_text_from_doc(path)
        elif filename.endswith(".md"):
            text = extract_text_from_md(path)
        else:
            return []
        chunks = chunking(text)
        if not chunks:
            return []
        return generate_embeddings(chunks, filename)
    except Exception as e:
        print(f"[index_single_file] Error procesando '{filename}': {e}")
        return []

# ── 3. RAG — chunking y embeddings ────────────────────────────────────────────

splitter = RecursiveCharacterTextSplitter(
    chunk_size=600,
    chunk_overlap=80,
    separators=["\n\n", "\n", " ", ""]
)

def chunking(text: str) -> list[str]:
    return splitter.split_text(text)


def embed_text(text: str) -> list[float]:
    """Embede un único texto (usado para queries en retrieve)."""
    return embed_batch([text])[0]


def embed_batch(texts: list[str], batch_size: int = 100) -> list[list[float]]:
    """Embede una lista de textos en lotes usando el SDK de Google GenAI."""
    results = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i + batch_size]
        for attempt in range(3):
            try:
                response = client.models.embed_content(
                    model=EMBED_MODEL,
                    contents=batch,
                )
                results.extend(e.values for e in response.embeddings)
                break
            except Exception as e:
                err = str(e)
                if "429" in err or "RESOURCE_EXHAUSTED" in err:
                    time.sleep(2 ** attempt)
                    continue
                raise
        else:
            raise RuntimeError(
                f"Rate-limit persistente al embeder lote {i // batch_size + 1}"
            )
    return results


def generate_embeddings(chunks: list[str], source: str) -> list[dict]:
    embeddings = embed_batch(chunks)
    indexed_at = datetime.now(timezone.utc).isoformat()
    return [
        {"source": source, "text": chunk, "embedding": emb, "indexed_at": indexed_at}
        for chunk, emb in zip(chunks, embeddings)
    ]


_query_embedding_cache: dict[str, list[float]] = {}

def retrieve(query: str, index: list[dict], top_k: int = 5) -> list[str]:
    if query not in _query_embedding_cache:
        _query_embedding_cache[query] = embed_text(query)
    query_vec = np.array(_query_embedding_cache[query])
    matrix    = np.array([item["embedding"] for item in index])
    norms  = np.linalg.norm(matrix, axis=1) * np.linalg.norm(query_vec)
    scores = matrix @ query_vec / np.where(norms == 0, 1e-10, norms)
    top_idx = np.argsort(scores)[-top_k:][::-1]
    label = query[:60].replace("\n", " ")
    for rank, i in enumerate(top_idx):
        src = index[int(i)].get("source", "?")
        print(f"[retrieval] rank={rank+1} score={scores[int(i)]:.3f} source={src} query={label!r}")
    return [index[int(i)]["text"] for i in top_idx]

# ── 4. RAG — construcción y carga del índice ──────────────────────────────────

def save_index(index: list[dict], path: str = "rag_index.json"):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)  # sin indent: los embeddings son arrays de ~768 floats


def load_index(path: str = "rag_index.json") -> list[dict] | None:
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


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


def init_interview(topic: str, system_prompt: str = None):
    chat = client.chats.create(
        model=CHAT_MODEL,
        config={
            "system_instruction": system_prompt or SYSTEM_PROMPT,
            "temperature": 0.7,
            "thinking_config": {"thinking_budget": 0},
        }
    )
    first_input = _build_initial_prompt(topic)
    response = chat.send_message(first_input)
    # Se registra first_input (no topic crudo) para que la transcripción
    # refleje exactamente lo que recibió el modelo redactor.
    log = [
        {"role": "user",      "content": first_input},
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


# ── 5b. Flujo edición de procedimiento existente ──────────────────────────────

EDIT_SYSTEM_PROMPT = """
# Instrucciones — Modo Edición · GPT Documentador ISO de GÓMEZ Y CRESPO S.A.

Eres un consultor experto en calidad ISO integrado en el sistema documental de GÓMEZ Y CRESPO S.A. (fabricante de equipamiento agroganadero, ISO 9001:2015 e ISO 14001:2015, sede en Ourense). Tu misión es revisar y actualizar un procedimiento existente.

## Contexto de la empresa

- Cargos habituales: Gerencia, Responsable de Calidad, Responsable de Compras, Responsable de Producción, Departamento Técnico, Administración.
- ERP/CRM corporativo: refiérete siempre a él como "el ERP".
- Elabora siempre: Responsable de Calidad.
- Aprueba siempre: Gerencia.
- No menciones cláusulas ISO en el documento; el cumplimiento normativo ya está implícito.

## Flujo de edición

1. En tu primer mensaje, lista brevemente las secciones del procedimiento cargado y pregunta al usuario qué secciones quiere modificar.
2. Para cada sección que el usuario indique:
   - Muestra el texto actual entre comillas.
   - Propón la nueva redacción completa tal como aparecerá en el documento.
   - Pregunta: "¿Es así, o lo ajustamos?"
   - No avances hasta confirmar.
3. Las secciones no mencionadas se mantienen exactamente igual que en el original.
4. Al terminar los cambios solicitados, pregunta: "¿Hay alguna sección más que quieras modificar?"
5. Cuando el usuario confirme que no hay más cambios, escribe exactamente en una línea:

FINALIZADO

No añadas ningún texto después de esa palabra.

## Reglas de redacción

- Usa siempre tercera persona + futuro de obligación.
- Usa tono formal, claro y narrativo.
- Nombra siempre el cargo completo.
- Menciona el ERP cuando sea relevante, pero nunca por su nombre comercial.
- No inventes datos.
- Usa negritas inline con **texto**.
- En el desarrollo, cada subapartado llevará un subtítulo en negrita como primera frase, seguido de texto narrativo en párrafos. No uses listas de puntos ni guiones dentro del desarrollo; redacta siempre en prosa continua.
- Mantén siempre la conversación en español.
"""


def _build_edit_prompt(document_text: str) -> str:
    return f"""\
He cargado el siguiente procedimiento para revisarlo:

--- DOCUMENTO ACTUAL ---
{document_text}
--- FIN DEL DOCUMENTO ---

Lista las secciones que contiene y pregúntame qué quiero modificar.
""".strip()


def init_edit_interview(document_text: str, system_prompt: str = None):
    chat = client.chats.create(
        model=CHAT_MODEL,
        config={
            "system_instruction": system_prompt or EDIT_SYSTEM_PROMPT,
            "temperature": 0.7,
            "thinking_config": {"thinking_budget": 0},
        }
    )
    first_input = _build_edit_prompt(document_text)
    response = chat.send_message(first_input)
    log = [
        {"role": "user",      "content": first_input},
        {"role": "assistant", "content": response.text}
    ]
    return chat, log


# ── 6. Fase redacción — Gemini Pro ────────────────────────────────────────────

DRAFT_SYSTEM_PROMPT = """
Eres el Redactor Jefe de Procedimientos ISO de GÓMEZ Y CRESPO S.A.
Tu única misión es convertir la transcripción de una entrevista en un procedimiento ISO completo en formato JSON.

Reglas:
- Escribe en español formal ISO. Tercera persona, futuro de obligación.
- Nivel de detalle ALTO: cada paso del desarrollo debe explicar qué se hace, quién lo hace, cómo se hace, en qué plazo si se mencionó, qué registro o documento se genera y cuál es el resultado esperado.
- Desarrolla cada subapartado en al menos 3-5 frases completas en prosa continua. Está terminantemente prohibido usar listas de puntos, guiones o enumeraciones dentro del Desarrollo; todo el contenido debe redactarse como párrafos narrativos fluidos.
- El apartado "Desarrollo" debe ser el más extenso del documento: desglosa el proceso en tantos subapartados como pasos tenga, con subtítulos en negrita.
- Las responsabilidades se generan en un paso separado a partir del Desarrollo; no las incluyas ni las menciones en tu redacción.
- Objeto y Alcance: redáctalos con suficiente contexto para que alguien ajeno a la empresa entienda el propósito y límites del procedimiento.
- Definiciones: incluye todas las siglas, términos técnicos y nombres de sistemas mencionados en la transcripción.
- No inventes datos que no aparezcan en la transcripción.
- Usa los cargos reales de GYC y menciona el ERP cuando sea relevante, nunca por su nombre comercial.
- El diagrama_mermaid debe representar fielmente el flujo completo del procedimiento, incluyendo decisiones y caminos alternativos si los hay. Usa SIEMPRE "flowchart TD" (top-down, vertical). Incluye un nodo por cada subapartado del Desarrollo más los nodos de inicio y fin. Nunca uses flowchart LR ni diagramas de un solo nivel.

## Reglas anti-sobrecompromiso (crítico para auditorías ISO)

- NUNCA escribas plazos concretos (24h, 48h, 5 días...) salvo que el usuario los haya confirmado explícitamente. Usa "en el plazo establecido", "de manera periódica", "cuando proceda".
- NUNCA escribas frecuencias concretas (diariamente, semanalmente...) salvo que se hayan mencionado. Usa "con la periodicidad definida", "según necesidad".
- En responsabilidades, describe QUÉ hace cada cargo, no cuándo ni con qué exactitud de tiempo.
- Prefiere "velará por", "se asegurará de", "supervisará" sobre verbos que implican comprobación sistemática con frecuencia fija.
- Si algo no se mencionó en la entrevista, omítelo o descríbelo de forma genérica. No rellenes huecos con prácticas habituales que podrían no cumplirse.
- El objetivo es que el documento describa fielmente lo que se hace, sin imponer más de lo que la empresa puede garantizar siempre.
"""

_DRAFT_PROMPT_TPL = """\
--- CONTEXTO: PROCEDIMIENTOS EXISTENTES DE GYC ---
{rag_context}

--- TRANSCRIPCIÓN DE LA ENTREVISTA ---
{transcript}

--- INSTRUCCIONES ---
Redacta el procedimiento ISO completo basándote exclusivamente en la transcripción de la entrevista y el contexto de procedimientos existentes.
Aplica todas las reglas del sistema: prosa narrativa en el desarrollo, sin listas, tercera persona, futuro de obligación, anti-sobrecompromiso.
"""


_RESPONSABILIDADES_SYSTEM = """
Eres un redactor de procedimientos ISO. Tu única tarea es extraer las responsabilidades
de los cargos que aparecen en el texto del Desarrollo de un procedimiento.

Reglas estrictas:
- Incluye ÚNICAMENTE los cargos mencionados de forma explícita en el Desarrollo.
- Escribe el nombre del cargo siempre con la primera letra de cada palabra en mayúscula
  (ej. "Responsable de Calidad", "Responsable del Área Afectada").
- Para cada cargo, agrupa sus acciones en 2 a 5 tareas concretas. No hagas una tarea
  por cada frase del Desarrollo; sintetiza las acciones relacionadas en una sola tarea.
- No añadas cargos que no aparezcan en el Desarrollo, aunque sean habituales en la empresa.
- No inventes tareas. Cada tarea debe estar respaldada por algo escrito en el Desarrollo.
- Redacta cada tarea en tercera persona, futuro de obligación (velará por, registrará, notificará…).
""".strip()

_RESPONSABILIDADES_PROMPT = """
A continuación tienes el texto completo del apartado DESARROLLO de un procedimiento ISO.
Extrae los cargos y sus responsabilidades tal y como se describen en ese texto.

--- DESARROLLO ---
{desarrollo}
"""


def _derive_responsabilidades(desarrollo: list[dict]) -> list[dict]:
    """Paso 2: deriva responsabilidades del texto real del Desarrollo (Gemini Flash)."""
    desarrollo_texto = "\n\n".join(
        f"{s['num']} {s['titulo']}\n{s['descripcion']}"
        for s in desarrollo
    )
    prompt = _RESPONSABILIDADES_PROMPT.format(desarrollo=desarrollo_texto)
    response = client.models.generate_content(
        model=CHAT_MODEL,
        contents=prompt,
        config={
            "system_instruction": _RESPONSABILIDADES_SYSTEM,
            "temperature": 0.2,
            "response_mime_type": "application/json",
            "response_schema": ListaResponsabilidades,
        },
    )
    if getattr(response, "parsed", None) is not None:
        return response.parsed.model_dump()["responsabilidades"]
    return ListaResponsabilidades.model_validate_json(response.text).responsabilidades


def draft_procedure(transcript: str, rag_context: str = "", draft_system_prompt: str = None) -> dict:
    """
    Genera el procedimiento en dos pasos:
      1. Gemini Pro redacta todo el documento (sin responsabilidades).
      2. Gemini Flash deriva las responsabilidades del Desarrollo generado.
    Retorna un dict validado contra el esquema Procedimiento completo.
    """
    prompt = _DRAFT_PROMPT_TPL.format(
        rag_context=rag_context or "No hay procedimientos existentes indexados.",
        transcript=transcript,
    )
    config = {
        "system_instruction": draft_system_prompt or DRAFT_SYSTEM_PROMPT,
        "temperature": 1,
        "response_mime_type": "application/json",
        "response_schema": ProcedimientoBase,
    }
    models_to_try = [DRAFT_MODEL, CHAT_MODEL]
    last_error = None
    for model in models_to_try:
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config=config,
            )
            if getattr(response, "parsed", None) is not None:
                base = response.parsed.model_dump()
            else:
                base = ProcedimientoBase.model_validate_json(response.text).model_dump()

            # Paso 2: responsabilidades derivadas del Desarrollo real
            responsabilidades = _derive_responsabilidades(base["desarrollo"])
            data = {**base, "responsabilidades": responsabilidades}
            return Procedimiento.model_validate(data).model_dump()
        except Exception as e:
            last_error = e
            if "503" in str(e) or "UNAVAILABLE" in str(e):
                continue
            raise
    raise last_error


# ── 7. Extracción de JSON y generación de .docx ───────────────────────────────

def add_defaults(data: dict) -> dict:
    """Fuerza los campos fijos de GYC ignorando lo que haya generado el modelo."""
    today = datetime.now().strftime("%d/%m/%y")
    overrides = {
        "revision":      "00",
        "paginas":       5,
        "elaborado_por": "Responsable de Calidad",
        "aprobado_por":  "Gerencia",
        "fecha":         today,
    }
    return {**data, **overrides}


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
