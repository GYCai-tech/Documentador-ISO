# app.py — Interfaz Chainlit para el Asistente ISO de GÓMEZ Y CRESPO S.A.

import asyncio
import os
import chainlit as cl
from chainlit.input_widget import TextInput

from asistente import (
    load_index, save_index, retrieve,
    extract_text_from_docx, extract_text_from_pdf, extract_text_from_md,
    index_single_file, chunking, generate_embeddings,
    init_interview, continue_interview, transcript_from_log,
    init_edit_interview,
    draft_procedure, add_defaults, generate_docx,
    SYSTEM_PROMPT, SYSTEM_PROMPT_EXPRESS, DRAFT_SYSTEM_PROMPT, EDIT_SYSTEM_PROMPT,
)
from auditoria import registrar_generacion

RAG_INDEX_PATH = os.environ.get("RAG_CACHE_DIR", ".") + "/rag_index.json"
FOLDER_PATH    = "base-conocimiento"


# ── Indexado con progreso ──────────────────────────────────────────────────────

async def _build_index_with_progress(folder_path: str) -> list[dict]:
    """Construye el índice RAG mostrando el progreso por archivo con TaskList."""
    files = [
        f for f in os.listdir(folder_path)
        if f.endswith(".docx") or f.endswith(".doc") or f.endswith(".pdf") or f.endswith(".md")
        or f.endswith(".xlsx") or f.endswith(".xls")
    ]

    task_list = cl.TaskList()
    task_list.status = "Indexando base de conocimiento..."
    await task_list.send()

    tasks = {}
    for filename in files:
        task = cl.Task(title=filename, status=cl.TaskStatus.READY)
        await task_list.add_task(task)
        tasks[filename] = task
    await task_list.update()

    index = []
    for filename in files:
        task = tasks[filename]
        task.status = cl.TaskStatus.RUNNING
        await task_list.update()

        path = os.path.join(folder_path, filename)
        entries = await asyncio.to_thread(index_single_file, path, filename)
        index.extend(entries)

        task.status = cl.TaskStatus.DONE
        task.title  = f"{filename} — {len(entries)} fragmentos"
        await task_list.update()

    await asyncio.to_thread(save_index, index, RAG_INDEX_PATH)
    task_list.status = f"Listo — {len(index)} fragmentos indexados"
    await task_list.update()

    return index


# ── Settings ───────────────────────────────────────────────────────────────────

@cl.on_settings_update
async def on_settings_update(settings: dict):
    cl.user_session.set("system_prompt",       settings.get("system_prompt",       SYSTEM_PROMPT))
    cl.user_session.set("draft_system_prompt", settings.get("draft_system_prompt", DRAFT_SYSTEM_PROMPT))


# ── Chat start ─────────────────────────────────────────────────────────────────

@cl.on_chat_start
async def on_chat_start():
    cl.user_session.set("system_prompt",       SYSTEM_PROMPT)
    cl.user_session.set("draft_system_prompt", DRAFT_SYSTEM_PROMPT)

    await cl.ChatSettings([
        TextInput(
            id="system_prompt",
            label="Prompt — Entrevistador (Flash)",
            initial=SYSTEM_PROMPT,
            multiline=True,
        ),
        TextInput(
            id="draft_system_prompt",
            label="Prompt — Redactor (Pro)",
            initial=DRAFT_SYSTEM_PROMPT,
            multiline=True,
        ),
    ]).send()

    index = load_index(RAG_INDEX_PATH)
    if index is not None:
        await cl.Message(content=f"Base de conocimiento cargada — {len(index)} fragmentos indexados.").send()
    else:
        index = await _build_index_with_progress(FOLDER_PATH)
    cl.user_session.set("rag_index", index)

    user     = cl.user_session.get("user")
    username = user.identifier if user else None
    greeting = f"Bienvenido, **{username}**.\n\n" if username else ""

    cl.user_session.set("phase", "menu")
    await cl.Message(
        content=(
            "# GYC · Asistente ISO\n\n"
            f"{greeting}"
            "¿Qué quieres hacer?"
        ),
        actions=[
            cl.Action(name="action_nuevo",   payload={"value": "nuevo"},   label="Crear nuevo procedimiento"),
            cl.Action(name="action_revisar", payload={"value": "revisar"}, label="Revisar procedimiento existente"),
            cl.Action(name="action_subir",   payload={"value": "subir"},   label="Subir documentos a la base de conocimiento"),
        ],
    ).send()


# ── Mensajes entrantes ─────────────────────────────────────────────────────────

@cl.on_message
async def on_message(msg: cl.Message):
    phase = cl.user_session.get("phase", "idle")

    if phase == "get_topic":
        await handle_topic(msg.content)

    elif phase == "interview":
        await handle_interview(msg.content)

    elif phase in ("upload", "drafting", "processing"):
        await cl.Message(content="Por favor, espera a que termine la operación actual.").send()

    elif phase in ("menu", "mode_select"):
        await cl.Message(content="Por favor, selecciona una opción con los botones de arriba.").send()

    elif phase == "idle":
        await cl.Message(content="Inicia una nueva sesión para continuar.").send()


# ── Action callbacks — menú principal ─────────────────────────────────────────

@cl.action_callback("action_nuevo")
async def on_action_nuevo(action: cl.Action):
    if cl.user_session.get("phase") not in ("menu", "post_upload"):
        await action.remove()
        return
    await action.remove()
    cl.user_session.set("phase", "mode_select")
    await cl.Message(
        content="¿Cómo quieres trabajar?",
        actions=[
            cl.Action(name="action_detallado", payload={"value": "detallado"}, label="Detallado — sección por sección"),
            cl.Action(name="action_express",   payload={"value": "express"},   label="Express — borrador rápido"),
        ],
    ).send()


@cl.action_callback("action_revisar")
async def on_action_revisar(action: cl.Action):
    await action.remove()
    cl.user_session.set("phase", "upload")

    uploaded = await cl.AskFileMessage(
        content="Sube el procedimiento que quieres editar (DOCX):",
        accept=["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
        max_files=1,
        max_size_mb=20,
    ).send()

    if not uploaded:
        cl.user_session.set("phase", "idle")
        await cl.Message(content="No se subió ningún archivo. Sesión finalizada.").send()
        return

    f = uploaded[0]
    doc_text = await asyncio.to_thread(extract_text_from_docx, f.path)

    import re
    rev_match = re.search(r'[Rr]ev(?:isi[oó]n)?\.?\s*[:.]?\s*(\d{2})', doc_text)
    cl.user_session.set("edit_revision", rev_match.group(1) if rev_match else None)
    cl.user_session.set("edit_mode", True)
    cl.user_session.set("topic", f.name)

    index       = cl.user_session.get("rag_index", [])
    rag_context = "\n\n".join(retrieve(f.name, index)) if index else ""
    cl.user_session.set("rag_context", rag_context)

    thinking = await cl.Message(content="").send()
    chat, log = await asyncio.to_thread(init_edit_interview, doc_text)
    cl.user_session.set("chat", chat)
    cl.user_session.set("log",  log)
    cl.user_session.set("phase", "interview")

    thinking.content = log[-1]["content"]
    await thinking.update()
    await _offer_approval()


@cl.action_callback("action_subir")
async def on_action_subir(action: cl.Action):
    await action.remove()
    await handle_upload()


# ── Action callbacks — selección de modo ──────────────────────────────────────

@cl.action_callback("action_detallado")
async def on_action_detallado(action: cl.Action):
    if cl.user_session.get("phase") != "mode_select":
        await action.remove()
        return
    await action.remove()
    cl.user_session.set("modo", "detallado")
    cl.user_session.set("system_prompt", SYSTEM_PROMPT)
    cl.user_session.set("phase", "get_topic")
    await cl.Message(content="Describe brevemente el procedimiento que quieres documentar:").send()


@cl.action_callback("action_express")
async def on_action_express(action: cl.Action):
    if cl.user_session.get("phase") != "mode_select":
        await action.remove()
        return
    await action.remove()
    cl.user_session.set("modo", "express")
    cl.user_session.set("system_prompt", SYSTEM_PROMPT_EXPRESS)
    cl.user_session.set("phase", "get_topic")
    await cl.Message(content="Describe brevemente el procedimiento que quieres documentar:").send()


# ── Action callback — aprobación durante entrevista ───────────────────────────

@cl.action_callback("action_ok")
async def on_action_ok(action: cl.Action):
    if cl.user_session.get("phase") != "interview":
        await action.remove()
        return

    # Comprueba que el botón es el más reciente (evita doble procesado con botones viejos)
    action_token   = action.payload.get("token", -1)
    current_token  = cl.user_session.get("approval_token", 0)
    if action_token != current_token:
        await action.remove()
        return

    cl.user_session.set("phase", "processing")
    await action.remove()

    chat = cl.user_session.get("chat")
    log  = cl.user_session.get("log")

    thinking = await cl.Message(content="").send()
    reply, log = await asyncio.to_thread(continue_interview, chat, "ok", log)
    cl.user_session.set("log", log)
    thinking.content = reply
    await thinking.update()

    if "FINALIZADO" in reply or _interview_complete(log):
        await generate_and_deliver()
        return

    cl.user_session.set("phase", "interview")
    await _offer_approval()


# ── Action callbacks — subida de documentos ───────────────────────────────────

@cl.action_callback("action_upload_mas")
async def on_action_upload_mas(action: cl.Action):
    await action.remove()
    await handle_upload()


@cl.action_callback("action_upload_volver")
async def on_action_upload_volver(action: cl.Action):
    await action.remove()
    total = len(cl.user_session.get("rag_index", []))
    await cl.Message(content=f"Base de conocimiento actualizada — **{total} fragmentos** en total.").send()
    cl.user_session.set("phase", "post_upload")
    await cl.Message(
        content="¿Qué quieres hacer ahora?",
        actions=[
            cl.Action(name="action_nuevo",  payload={"value": "nuevo"},  label="Crear nuevo procedimiento"),
            cl.Action(name="action_cerrar", payload={"value": "cerrar"}, label="Terminar sesión"),
        ],
    ).send()


@cl.action_callback("action_cerrar")
async def on_action_cerrar(action: cl.Action):
    await action.remove()
    cl.user_session.set("phase", "idle")
    await cl.Message(content="Sesión terminada. Recarga la página para empezar de nuevo.").send()


# ── Handlers ───────────────────────────────────────────────────────────────────

async def handle_topic(topic: str):
    cl.user_session.set("topic", topic)
    cl.user_session.set("phase", "processing")

    index       = cl.user_session.get("rag_index", [])
    rag_context = "\n\n".join(retrieve(topic, index)) if index else ""
    cl.user_session.set("rag_context", rag_context)

    modo = cl.user_session.get("modo", "detallado")
    default_prompt = SYSTEM_PROMPT_EXPRESS if modo == "express" else SYSTEM_PROMPT
    system_prompt  = cl.user_session.get("system_prompt", default_prompt)

    thinking = await cl.Message(content="").send()
    chat, log = await asyncio.to_thread(init_interview, topic, system_prompt)
    cl.user_session.set("chat", chat)
    cl.user_session.set("log",  log)
    cl.user_session.set("phase", "interview")

    thinking.content = log[-1]["content"]
    await thinking.update()

    await _offer_approval()


async def _offer_approval():
    """Muestra un botón OK no bloqueante — el usuario puede escribir libremente en lugar de pulsarlo."""
    token = cl.user_session.get("approval_token", 0) + 1
    cl.user_session.set("approval_token", token)
    await cl.Message(
        content="¿Continuamos con la siguiente sección?",
        actions=[
            cl.Action(name="action_ok", payload={"value": "ok", "token": token}, label="✓ OK, así está bien"),
        ],
    ).send()


async def handle_interview(user_input: str):
    cl.user_session.set("phase", "processing")
    chat = cl.user_session.get("chat")
    log  = cl.user_session.get("log")

    thinking = await cl.Message(content="").send()
    reply, log = await asyncio.to_thread(continue_interview, chat, user_input, log)
    cl.user_session.set("log", log)

    thinking.content = reply
    await thinking.update()

    if "FINALIZADO" in reply or _interview_complete(log):
        await generate_and_deliver()
        return

    cl.user_session.set("phase", "interview")
    await _offer_approval()


def _interview_complete(log: list) -> bool:
    last = log[-1]["content"] if log else ""
    return "FINALIZADO" in last or "procedimiento completo" in last.lower()


async def handle_upload():
    cl.user_session.set("phase", "upload")

    uploaded = await cl.AskFileMessage(
        content="Sube uno o más documentos (PDF, DOCX, DOC, MD, XLSX, XLS):",
        accept=[
            "application/pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/msword",
            "text/markdown",
            "text/plain",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/vnd.ms-excel",
        ],
        max_files=10,
        max_size_mb=20,
    ).send()

    if not uploaded:
        cl.user_session.set("phase", "idle")
        return

    index = cl.user_session.get("rag_index", [])
    for f in uploaded:
        msg = await cl.Message(content=f"Indexando **{f.name}**...").send()
        # Purga entradas obsoletas del mismo archivo antes de re-indexar
        before = len(index)
        index  = [e for e in index if e.get("source") != f.name]
        purged = before - len(index)
        try:
            entries = await asyncio.to_thread(index_single_file, f.path, f.name)
        except Exception as e:
            msg.content = f"**{f.name}** — error al indexar: {e}"
            await msg.update()
            continue
        if entries:
            index.extend(entries)
            indexed_at  = entries[0].get("indexed_at", "")[:10]
            purge_note  = f", reemplazó {purged} fragmentos anteriores" if purged else ""
            msg.content = f"**{f.name}** — {len(entries)} fragmentos indexados ({indexed_at}{purge_note})."
        else:
            msg.content = f"**{f.name}** — formato no soportado u omitido."
        await msg.update()

    cl.user_session.set("rag_index", index)
    await asyncio.to_thread(save_index, index, RAG_INDEX_PATH)

    await cl.Message(
        content="¿Quieres subir más documentos?",
        actions=[
            cl.Action(name="action_upload_mas",    payload={"value": "mas"},    label="Subir más"),
            cl.Action(name="action_upload_volver", payload={"value": "volver"}, label="Volver al menú"),
        ],
    ).send()


async def generate_and_deliver():
    cl.user_session.set("phase", "drafting")

    log                 = cl.user_session.get("log")
    rag_context         = cl.user_session.get("rag_context", "")
    transcript          = transcript_from_log(log)
    draft_system_prompt = cl.user_session.get("draft_system_prompt", DRAFT_SYSTEM_PROMPT)
    tema                = cl.user_session.get("topic", "")
    index               = cl.user_session.get("rag_index", [])
    user                = cl.user_session.get("user")
    username            = user.identifier if user else "anónimo"

    status = await cl.Message(content="Redactando procedimiento ISO, un momento...").send()

    # Re-retrieval con el transcript completo: captura vocabulario específico que
    # el tema inicial no contenía (nombres de registros, cargos, pasos concretos).
    if index and transcript:
        draft_query    = transcript[-1200:] if len(transcript) > 1200 else transcript
        extra_chunks   = await asyncio.to_thread(retrieve, draft_query, index, top_k=3)
        existing_texts = set(rag_context.split("\n\n")) if rag_context else set()
        additions      = [c for c in extra_chunks if c not in existing_texts]
        if additions:
            rag_context = (rag_context + "\n\n" + "\n\n".join(additions)).strip()

    data = None
    for attempt in range(2):
        try:
            data = await asyncio.to_thread(draft_procedure, transcript, rag_context, draft_system_prompt)
            break
        except Exception:
            if attempt == 0:
                status.content = "Error en la generación, reintentando..."
                await status.update()

    if not data:
        status.content = "No se pudo generar el procedimiento. Puedes seguir editando o iniciar una nueva sesión."
        await status.update()
        cl.user_session.set("phase", "interview")
        return

    data         = add_defaults(data)
    edit_revision = cl.user_session.get("edit_revision")
    if edit_revision is not None:
        data["revision"] = edit_revision
    out_path = await asyncio.to_thread(generate_docx, data)

    codigo = data.get("codigo", "PC-XX")
    nombre = data.get("nombre", "")

    await asyncio.to_thread(registrar_generacion, codigo, nombre, tema, out_path, username)

    status.content = f"Procedimiento **{codigo} — {nombre}** generado correctamente."
    await status.update()

    await cl.Message(
        content="Tu procedimiento está listo:",
        elements=[cl.File(name=os.path.basename(out_path), path=out_path)]
    ).send()

    cl.user_session.set("phase", "idle")
    cl.user_session.set("edit_mode", False)
    cl.user_session.set("edit_revision", None)
