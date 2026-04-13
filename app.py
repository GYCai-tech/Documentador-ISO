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
    draft_procedure, extract_json, add_defaults, generate_docx,
    SYSTEM_PROMPT, DRAFT_SYSTEM_PROMPT,
)

RAG_INDEX_PATH = os.environ.get("RAG_CACHE_DIR", ".") + "/rag_index.json"
FOLDER_PATH    = "base-conocimiento"


# ── Indexado con progreso ──────────────────────────────────────────────────────

async def _build_index_with_progress(folder_path: str) -> list[dict]:
    """Construye el índice RAG mostrando el progreso por archivo con TaskList."""
    files = [
        f for f in os.listdir(folder_path)
        if f.endswith(".docx") or f.endswith(".doc") or f.endswith(".pdf") or f.endswith(".md")
    ]

    task_list = cl.TaskList()
    task_list.status = "Indexando base de conocimiento..."
    await task_list.send()

    # Crea una tarea por archivo
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


# ── Inicio de sesión ───────────────────────────────────────────────────────────

@cl.on_settings_update
async def on_settings_update(settings: dict):
    cl.user_session.set("system_prompt",       settings.get("system_prompt",       SYSTEM_PROMPT))
    cl.user_session.set("draft_system_prompt", settings.get("draft_system_prompt", DRAFT_SYSTEM_PROMPT))


@cl.on_chat_start
async def on_chat_start():
    # Inicializa los prompts en sesión y muestra el panel de configuración
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

    # Carga o construye el índice RAG
    index = load_index(RAG_INDEX_PATH)
    if index is not None:
        await cl.Message(content=f"Base de conocimiento cargada — {len(index)} fragmentos indexados.").send()
    else:
        index = await _build_index_with_progress(FOLDER_PATH)
    cl.user_session.set("rag_index", index)

    # Bienvenida y selección de modo
    await cl.Message(
        content=(
            "# GYC · Asistente ISO\n\n"
            "Bienvenido al asistente de documentación ISO 9001 de **Gómez y Crespo S.A.**\n\n"
            "¿Qué quieres hacer?"
        )
    ).send()

    res = await cl.AskActionMessage(
        content="Selecciona una opción:",
        actions=[
            cl.Action(name="nuevo",   payload={"value": "nuevo"},   label="Crear nuevo procedimiento"),
            cl.Action(name="revisar", payload={"value": "revisar"}, label="Revisar procedimiento existente"),
            cl.Action(name="subir",   payload={"value": "subir"},   label="Subir documentos a la base de conocimiento"),
        ],
    ).send()

    value = res.get("payload", {}).get("value") if res else None
    if value == "nuevo":
        cl.user_session.set("phase", "get_topic")
        await cl.Message(content="Describe brevemente el procedimiento que quieres documentar:").send()
    elif value == "subir":
        cl.user_session.set("phase", "upload")
        await handle_upload()
    else:
        cl.user_session.set("phase", "idle")
        await cl.Message(content="Función de revisión próximamente disponible.").send()


# ── Mensajes entrantes ─────────────────────────────────────────────────────────

@cl.on_message
async def on_message(msg: cl.Message):
    phase = cl.user_session.get("phase", "idle")

    if phase == "get_topic":
        await handle_topic(msg.content)

    elif phase == "interview":
        await handle_interview(msg.content)

    elif phase == "idle":
        await cl.Message(content="Inicia una nueva sesión para continuar.").send()


# ── Handlers ──────────────────────────────────────────────────────────────────

async def handle_topic(topic: str):
    cl.user_session.set("topic", topic)

    # Recupera contexto RAG relevante
    index       = cl.user_session.get("rag_index", [])
    rag_context = "\n\n".join(retrieve(topic, index)) if index else ""
    cl.user_session.set("rag_context", rag_context)

    # Inicia entrevista
    system_prompt = cl.user_session.get("system_prompt", SYSTEM_PROMPT)
    thinking = await cl.Message(content="").send()
    chat, log = await asyncio.to_thread(init_interview, topic, system_prompt)
    cl.user_session.set("chat", chat)
    cl.user_session.set("log",  log)
    cl.user_session.set("phase", "interview")

    thinking.content = log[-1]["content"]
    await thinking.update()


async def handle_interview(user_input: str):
    chat = cl.user_session.get("chat")
    log  = cl.user_session.get("log")

    # Envía el mensaje y espera respuesta
    thinking = await cl.Message(content="").send()
    reply, log = await asyncio.to_thread(continue_interview, chat, user_input, log)
    cl.user_session.set("log", log)

    thinking.content = reply
    await thinking.update()

    # Detecta fin de entrevista
    if "FINALIZADO" in reply or _interview_complete(log):
        await generate_and_deliver()


def _interview_complete(log: list) -> bool:
    """Detecta si el asistente ha marcado la entrevista como completa."""
    last = log[-1]["content"] if log else ""
    return "FINALIZADO" in last or "procedimiento completo" in last.lower()


async def handle_upload():
    while True:
        uploaded = await cl.AskFileMessage(
            content="Sube uno o más documentos (PDF, DOCX, DOC, MD):",
            accept=[
                "application/pdf",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/msword",
                "text/markdown",
                "text/plain",
            ],
            max_files=10,
            max_size_mb=20,
        ).send()

        if not uploaded:
            break

        index = cl.user_session.get("rag_index", [])
        for f in uploaded:
            msg = await cl.Message(content=f"Indexando **{f.name}**...").send()
            entries = await asyncio.to_thread(index_single_file, f.path, f.name)
            if entries:
                index.extend(entries)
                msg.content = f"**{f.name}** — {len(entries)} fragmentos indexados."
            else:
                msg.content = f"**{f.name}** — formato no soportado, omitido."
            await msg.update()

        cl.user_session.set("rag_index", index)
        await asyncio.to_thread(save_index, index, RAG_INDEX_PATH)

        res = await cl.AskActionMessage(
            content="¿Quieres subir más documentos?",
            actions=[
                cl.Action(name="mas",    payload={"value": "mas"},    label="Subir más"),
                cl.Action(name="volver", payload={"value": "volver"}, label="Volver al menú"),
            ],
        ).send()

        if not res or res.get("payload", {}).get("value") != "mas":
            break

    total = len(cl.user_session.get("rag_index", []))
    cl.user_session.set("phase", "idle")
    await cl.Message(content=f"Base de conocimiento actualizada — **{total} fragmentos** en total.").send()


async def generate_and_deliver():
    cl.user_session.set("phase", "drafting")

    log               = cl.user_session.get("log")
    rag_context       = cl.user_session.get("rag_context", "")
    transcript        = transcript_from_log(log)
    draft_system_prompt = cl.user_session.get("draft_system_prompt", DRAFT_SYSTEM_PROMPT)

    status = await cl.Message(content="Redactando procedimiento ISO, un momento...").send()

    # Llama al modelo de redacción
    draft = await asyncio.to_thread(draft_procedure, transcript, rag_context, draft_system_prompt)
    data  = extract_json(draft)

    if not data:
        status.content = "No se pudo extraer el JSON del procedimiento. Inténtalo de nuevo."
        await status.update()
        return

    data     = add_defaults(data)
    out_path = await asyncio.to_thread(generate_docx, data)

    codigo = data.get("codigo", "PC-XX")
    nombre = data.get("nombre", "")
    status.content = f"Procedimiento **{codigo} — {nombre}** generado correctamente."
    await status.update()

    # Ofrece descarga
    await cl.Message(
        content="Tu procedimiento está listo:",
        elements=[cl.File(name=os.path.basename(out_path), path=out_path)]
    ).send()

    cl.user_session.set("phase", "idle")
