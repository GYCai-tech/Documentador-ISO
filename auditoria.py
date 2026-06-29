import hashlib
import os
import sqlite3
from datetime import datetime, timezone

DB_PATH = os.path.join(os.environ.get("RAG_CACHE_DIR", "."), "auditoria.db")


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS generaciones (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            usuario   TEXT DEFAULT 'anónimo',
            codigo    TEXT,
            nombre    TEXT,
            tema      TEXT,
            docx_hash TEXT,
            docx_path TEXT
        )
    """)
    # Migración: añade columna usuario si la tabla existía sin ella
    try:
        conn.execute("ALTER TABLE generaciones ADD COLUMN usuario TEXT DEFAULT 'anónimo'")
    except sqlite3.OperationalError:
        pass  # Ya existe
    conn.commit()
    return conn


def registrar_generacion(
    codigo: str,
    nombre: str,
    tema: str,
    docx_path: str,
    usuario: str = "anónimo",
) -> str | None:
    """Registra la generación de un documento en el audit trail. Devuelve timestamp UTC o None."""
    try:
        ts = datetime.now(timezone.utc).isoformat()
        with open(docx_path, "rb") as f:
            docx_hash = hashlib.sha256(f.read()).hexdigest()
        with _get_conn() as conn:
            conn.execute(
                "INSERT INTO generaciones (timestamp, usuario, codigo, nombre, tema, docx_hash, docx_path) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (ts, usuario, codigo, nombre, tema, docx_hash, docx_path),
            )
        return ts
    except Exception as e:
        print(f"[auditoria] Error registrando generación: {e}")
        return None
