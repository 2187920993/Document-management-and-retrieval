import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path

from .config import DATABASE_URL, STORAGE_DIR


def _sqlite_path() -> Path:
    if DATABASE_URL.startswith("sqlite:///"):
        return Path(DATABASE_URL.removeprefix("sqlite:///"))
    return Path("./data/app.db")


def is_postgres() -> bool:
    return DATABASE_URL.startswith(("postgresql://", "postgres://"))


def connect():
    if is_postgres():
        import psycopg
        conn = psycopg.connect(DATABASE_URL, row_factory=psycopg.rows.dict_row)
        conn.autocommit = True
        return conn
    path = _sqlite_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


@contextmanager
def db():
    conn = connect()
    try:
        yield conn
        if not is_postgres():
            conn.commit()
    finally:
        conn.close()


def init_db():
    STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    with db() as conn:
        if is_postgres():
            conn.execute("""
                CREATE TABLE IF NOT EXISTS categories (
                id SERIAL PRIMARY KEY, name TEXT NOT NULL UNIQUE, parent_id INTEGER REFERENCES categories(id) ON DELETE SET NULL, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS documents (
                    id TEXT PRIMARY KEY, name TEXT NOT NULL, original_name TEXT NOT NULL, extension TEXT NOT NULL,
                    size_bytes BIGINT NOT NULL, mime_type TEXT NOT NULL, category_id INTEGER REFERENCES categories(id) ON DELETE SET NULL,
                    storage_key TEXT NOT NULL UNIQUE, content_hash TEXT, archived BOOLEAN NOT NULL DEFAULT FALSE,
                    storage_status TEXT NOT NULL, index_status TEXT NOT NULL, index_error TEXT,
                    content_version INTEGER NOT NULL DEFAULT 1, uploaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS index_jobs (
                    id BIGSERIAL PRIMARY KEY, document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    status TEXT NOT NULL DEFAULT 'pending', attempts INTEGER NOT NULL DEFAULT 0,
                    available_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), error TEXT, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS document_chunks (
                    id BIGSERIAL PRIMARY KEY, document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    content_version INTEGER NOT NULL, chunk_index INTEGER NOT NULL, text TEXT NOT NULL, vector JSONB NOT NULL,
                    UNIQUE(document_id, content_version, chunk_index)
                )
            """)
            conn.execute("CREATE TABLE IF NOT EXISTS app_settings (key TEXT PRIMARY KEY, value TEXT NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())")
            conn.execute("ALTER TABLE documents ADD COLUMN IF NOT EXISTS storage_path TEXT")
            conn.execute("ALTER TABLE documents ADD COLUMN IF NOT EXISTS content_hash TEXT")
            conn.execute("ALTER TABLE categories ADD COLUMN IF NOT EXISTS parent_id INTEGER REFERENCES categories(id) ON DELETE SET NULL")
            conn.execute("UPDATE documents SET storage_path = storage_key WHERE storage_path IS NULL")
            conn.execute("ALTER TABLE categories DROP CONSTRAINT IF EXISTS categories_name_key")
            conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS categories_parent_name_unique ON categories ((COALESCE(parent_id, 0)), name)")
        else:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE, parent_id INTEGER REFERENCES categories(id) ON DELETE SET NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS documents (
                    id TEXT PRIMARY KEY, name TEXT NOT NULL, original_name TEXT NOT NULL, extension TEXT NOT NULL,
                    size_bytes INTEGER NOT NULL, mime_type TEXT NOT NULL, category_id INTEGER REFERENCES categories(id) ON DELETE SET NULL,
                    storage_key TEXT NOT NULL UNIQUE, content_hash TEXT, archived INTEGER NOT NULL DEFAULT 0,
                    storage_status TEXT NOT NULL, index_status TEXT NOT NULL, index_error TEXT,
                    content_version INTEGER NOT NULL DEFAULT 1, uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS index_jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    status TEXT NOT NULL DEFAULT 'pending', attempts INTEGER NOT NULL DEFAULT 0,
                    available_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, error TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS document_chunks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
                    content_version INTEGER NOT NULL, chunk_index INTEGER NOT NULL, text TEXT NOT NULL, vector TEXT NOT NULL,
                    UNIQUE(document_id, content_version, chunk_index)
                );
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY, value TEXT NOT NULL, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
            """)
            columns = {row["name"] for row in conn.execute("PRAGMA table_info(documents)").fetchall()}
            category_columns = {row["name"] for row in conn.execute("PRAGMA table_info(categories)").fetchall()}
            if "parent_id" not in category_columns:
                conn.execute("ALTER TABLE categories ADD COLUMN parent_id INTEGER REFERENCES categories(id) ON DELETE SET NULL")
            if "storage_path" not in columns:
                conn.execute("ALTER TABLE documents ADD COLUMN storage_path TEXT")
            if "content_hash" not in columns:
                conn.execute("ALTER TABLE documents ADD COLUMN content_hash TEXT")
            conn.execute("UPDATE documents SET storage_path = storage_key WHERE storage_path IS NULL")
            # Older SQLite databases used UNIQUE(name), which prevented two
            # different parent categories from sharing a child name. Rebuild
            # the small category table once, preserving IDs and references.
            indexes = conn.execute("PRAGMA index_list(categories)").fetchall()
            has_hierarchical_index = any(row[1] == "categories_parent_name_unique" for row in indexes)
            if not has_hierarchical_index:
                conn.execute("PRAGMA foreign_keys = OFF")
                conn.execute("CREATE TABLE categories_migrated (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, parent_id INTEGER REFERENCES categories_migrated(id) ON DELETE SET NULL, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)")
                conn.execute("INSERT INTO categories_migrated (id, name, parent_id, created_at) SELECT id, name, parent_id, created_at FROM categories")
                conn.execute("DROP TABLE categories")
                conn.execute("ALTER TABLE categories_migrated RENAME TO categories")
                conn.execute("CREATE UNIQUE INDEX categories_parent_name_unique ON categories (COALESCE(parent_id, 0), name)")
                conn.execute("PRAGMA foreign_keys = ON")
