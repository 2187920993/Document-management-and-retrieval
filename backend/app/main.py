import io
import json
import base64
import hashlib
import math
import mimetypes
import re
import shutil
import uuid
import zipfile
from html import escape as html_escape
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse

from .config import ALLOWED_EXTENSIONS, EMBEDDING_DIM, MAX_BATCH_DOWNLOAD_BYTES, MAX_FILE_BYTES, POSTGRES_HOST_PATH, PROJECT_ROOT, STORAGE_DIR, STORAGE_HOST_PATH, STORAGE_LAYOUT
from .db import db, init_db, is_postgres
from .embeddings import EMBEDDING_VERSION, cosine, embed, tokenize
from .indexer import extract_text, index_document
from .schemas import Category, CategoryCreate, Document, DocumentBatchDelete, DocumentBatchRename, DocumentUpdate, Health, SearchResponse, SearchResult
from .settings import get_setting, get_settings, set_settings

app = FastAPI(title="文件管理与知识检索平台", version="0.1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])


class DuplicateUploadError(Exception):
    def __init__(self, filename: str, existing_name: str | None = None, existing_id: str | None = None):
        self.filename = filename
        self.existing_name = existing_name
        self.existing_id = existing_id
        super().__init__(filename)


def _preview_text_html(value) -> str:
    return html_escape("" if value is None else str(value)).replace("\n", "<br>")


def _preview_table_html(rows: list[list[object]], limit: int = 12000) -> str:
    rendered: list[str] = []
    for row_index, row in enumerate(rows):
        cells = "".join(f"<{('th' if row_index == 0 else 'td')}>{_preview_text_html(value)}</{('th' if row_index == 0 else 'td')}>" for value in row)
        rendered.append(f"<tr>{cells}</tr>")
        if len("".join(rendered)) >= limit:
            break
    return f"<table><tbody>{''.join(rendered)}</tbody></table>" if rendered else ""


def _pptx_preview_html(path: Path) -> str:
    from pptx import Presentation

    presentation = Presentation(str(path))
    sections: list[str] = []
    for slide_index, slide in enumerate(presentation.slides, 1):
        parts = [f"<section><h3>第 {slide_index} 页</h3>"]
        table_rows: list[list[object]] = []
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                table_rows.extend([[cell.text for cell in row.cells] for row in shape.table.rows])
            elif hasattr(shape, "text") and str(shape.text).strip():
                parts.append(f"<p>{_preview_text_html(shape.text)}</p>")
        if table_rows:
            parts.append(_preview_table_html(table_rows))
        parts.append("</section>")
        sections.append("".join(parts))
        if len("".join(sections)) >= 100000:
            break
    return "".join(sections)[:100000]


def _spreadsheet_preview_html(path: Path, extension: str) -> str:
    sections: list[str] = []
    if extension == ".xlsx":
        from openpyxl import load_workbook

        workbook = load_workbook(filename=str(path), read_only=True, data_only=True)
        sheets = ((sheet.title, sheet.iter_rows(values_only=True)) for sheet in workbook.worksheets)
    else:
        import xlrd

        workbook = xlrd.open_workbook(filename=str(path), on_demand=True)
        sheets = ((sheet.name, (sheet.row_values(index) for index in range(sheet.nrows))) for sheet in workbook.sheets())
    for sheet_name, row_iter in sheets:
        rows: list[list[object]] = []
        for row in row_iter:
            values = list(row)
            if any(value not in (None, "") for value in values):
                rows.append(values)
            if len(rows) >= 5000:
                break
        sections.append(f"<section><h3>{_preview_text_html(sheet_name)}</h3>{_preview_table_html(rows)}</section>")
        if len("".join(sections)) >= 100000:
            break
    return "".join(sections)[:100000]


MIGRATION_DIR = PROJECT_ROOT / "runtime"
MIGRATION_REQUEST = MIGRATION_DIR / ".storage-migration.request"
MIGRATION_DB_PATH = MIGRATION_DIR / ".storage-migration.db"
MIGRATION_STORAGE_PATH = MIGRATION_DIR / ".storage-migration.storage"
MIGRATION_STATUS = MIGRATION_DIR / ".storage-migration.status"
BACKGROUND_DIR = PROJECT_ROOT / "runtime" / "ui-background"
BACKGROUND_MAX_BYTES = 10 * 1024 * 1024
BACKGROUND_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".gif": "image/gif",
}


def _background_path() -> Path | None:
    """Return the single persisted workspace background, if one exists."""
    for extension in BACKGROUND_TYPES:
        path = BACKGROUND_DIR / f"background{extension}"
        if path.is_file():
            return path
    return None


def _is_image_payload(payload: bytes, extension: str) -> bool:
    signatures = {
        ".png": payload.startswith(b"\x89PNG\r\n\x1a\n"),
        ".jpg": payload.startswith(b"\xff\xd8\xff"),
        ".jpeg": payload.startswith(b"\xff\xd8\xff"),
        ".gif": payload.startswith((b"GIF87a", b"GIF89a")),
        ".webp": len(payload) >= 12 and payload[:4] == b"RIFF" and payload[8:12] == b"WEBP",
    }
    return bool(signatures.get(extension))


def _read_migration_status() -> dict:
    if not MIGRATION_STATUS.exists():
        return {"status": "idle"}
    try:
        value = json.loads(MIGRATION_STATUS.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {"status": "idle"}
    except (OSError, ValueError, TypeError):
        return {"status": "unknown", "message": "无法读取迁移状态"}


def _validate_host_path(value: object, field: str) -> str:
    path = str(value or "").strip().strip('"').strip("'").strip()
    if not path:
        raise ValueError(f"{field} 不能为空")
    if len(path) > 1024 or any(char in path for char in ("\x00", "\r", "\n", "\"", "<", ">", "|", "?", "*", "&", "^", "%", "`", ";")):
        raise ValueError(f"{field} 包含不支持的字符")
    if path.startswith("-"):
        raise ValueError(f"{field} 不能以 - 开头")
    # Store Windows paths with forward slashes so the BAT status JSON remains
    # valid and Docker Compose accepts them consistently.
    return path.replace("\\", "/")


def _write_migration_text(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    temporary.write_text(value, encoding="utf-8")
    temporary.replace(path)


def qmark(sql: str) -> str:
    return sql.replace("?", "%s") if is_postgres() else sql


def row_to_document(row) -> Document:
    def get(key):
        return row[key]
    return Document(
        id=get("id"), name=get("name"), original_name=get("original_name"), extension=get("extension"),
        size_bytes=get("size_bytes"), mime_type=get("mime_type"), storage_path=get("storage_path") if "storage_path" in row.keys() else get("storage_key"), category_id=get("category_id"),
        storage_conflict=False,
        category_name=get("category_name"), storage_status=get("storage_status"), index_status=get("index_status"),
        index_error=get("index_error"), archived=bool(get("archived")), uploaded_at=get("uploaded_at"),
    )


def select_documents(where: str = "", args: tuple = (), order: str = "uploaded_at DESC"):
    sql = """SELECT d.*, c.name AS category_name FROM documents d LEFT JOIN categories c ON c.id = d.category_id"""
    if where:
        sql += " WHERE " + where
    sql += " ORDER BY " + order
    with db() as conn:
        return conn.execute(qmark(sql), args).fetchall()


def _validate_extension(filename: str) -> str:
    extension = Path(filename or "").suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        supported = ", ".join(sorted(ALLOWED_EXTENSIONS))
        raise HTTPException(status_code=400, detail=f"不支持的文件类型 {extension or '(无扩展名)'}，支持：{supported}")
    return extension


def _safe_storage_name(name: str | None, document_id: str, extension: str) -> str:
    candidate = Path(name or "").name
    candidate = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", candidate).strip(" .")
    if not candidate:
        candidate = f"{document_id}{extension}"
    if not Path(candidate).suffix:
        candidate += extension
    return candidate


def _strip_generated_collision_suffix(name: str) -> str:
    """Remove repeated collision suffixes created by older migrations."""
    path = Path(name)
    stem = re.sub(r"(?:__(?:doc_)?[0-9a-f]{4,8}(?:_\d+)?)+$", "", path.stem, flags=re.IGNORECASE)
    return f"{stem or path.stem}{path.suffix}"


def _safe_folder_name(name: str | None, fallback: str) -> str:
    candidate = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", (name or "").strip()).strip(" .")
    return candidate or fallback


def _category_folder(category_id: int | None) -> str:
    if category_id is None:
        return "未分类"
    parts: list[str] = []
    seen: set[int] = set()
    current = category_id
    with db() as conn:
        while current is not None and current not in seen:
            seen.add(current)
            row = conn.execute(qmark("SELECT id, name, parent_id FROM categories WHERE id = ?"), (current,)).fetchone()
            if not row:
                break
            parts.append(_safe_folder_name(row["name"], "未分类"))
            current = row["parent_id"]
    return "/".join(reversed(parts)) or "未分类"


def _storage_key(document_id: str, extension: str, category_id: int | None = None, archived: bool = False, source_name: str | None = None) -> str:
    filename = _safe_storage_name(source_name, document_id, extension)
    if get_setting("storage_layout") != "folders":
        return filename
    area = "归档" if archived else "资料"
    category = _category_folder(category_id)
    return f"{area}/{category}/{filename}"


def _existing_storage_name(row) -> str:
    """Keep source names when moving; recognize legacy document-id filenames."""
    old_name = Path(row["storage_key"]).name
    legacy_name = f"{row['id']}{row['extension']}"
    source_name = row["original_name"] if old_name == legacy_name else old_name
    return _strip_generated_collision_suffix(source_name)


def _collision_safe_key(key: str, document_id: str) -> str:
    relative = Path(key)
    normalized_name = _strip_generated_collision_suffix(relative.name)
    key = str(relative.with_name(normalized_name)).replace("\\", "/")
    path = STORAGE_DIR / key
    if not path.exists():
        return key
    suffix = path.suffix
    candidate = path.with_name(f"{path.stem}__{document_id[:8]}{suffix}")
    counter = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}__{document_id[:8]}_{counter}{suffix}")
        counter += 1
    return str(candidate.relative_to(STORAGE_DIR)).replace("\\", "/")


def migrate_storage_layout() -> None:
    """Normalize persisted files to source names and archive/category folders."""
    rows = select_documents()
    moved: list[tuple[str, str]] = []
    updates: list[tuple[str, str, str]] = []
    try:
        for row in rows:
            source_name = _existing_storage_name(row)
            new_key = _collision_safe_key(_storage_key(row["id"], row["extension"], row["category_id"], bool(row["archived"]), source_name), row["id"])
            old_key = row["storage_key"]
            if new_key != old_key:
                source = STORAGE_DIR / old_key
                target = STORAGE_DIR / new_key
                if source.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(source), str(target))
                    moved.append((old_key, new_key))
            updates.append((new_key, new_key, row["id"]))
        with db() as conn:
            for storage_key, storage_path, document_id in updates:
                conn.execute(qmark("UPDATE documents SET storage_key = ?, storage_path = ? WHERE id = ?"), (storage_key, storage_path, document_id))
    except Exception:
        for old_key, new_key in reversed(moved):
            source, target = STORAGE_DIR / new_key, STORAGE_DIR / old_key
            if source.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(source), str(target))
        raise


def _duplicate_filename(filename: str, occupied: set[str]) -> str:
    """Give an intentional content duplicate a visible copy name."""
    path = Path(filename)
    stem, suffix = path.stem, path.suffix
    occupied_folded = {name.casefold() for name in occupied}
    candidate = f"{stem} (副本){suffix}"
    counter = 2
    while candidate.casefold() in occupied_folded:
        candidate = f"{stem} (副本 {counter}){suffix}"
        counter += 1
    return candidate


def _find_content_duplicate(content_hash: str):
    with db() as conn:
        return conn.execute(qmark("SELECT id, name FROM documents WHERE content_hash = ? ORDER BY uploaded_at ASC LIMIT 1"), (content_hash,)).fetchone()


async def _store_uploaded_file(file: UploadFile, category_id: int | None, duplicate_action: str = "prompt", content: bytes | None = None, seen_hashes: set[str] | None = None) -> Document:
    filename = file.filename or "未命名文件"
    extension = _validate_extension(filename)
    if content is None:
        content = await file.read()
    max_file_bytes = int(get_setting("max_file_bytes"))
    if len(content) > max_file_bytes:
        raise HTTPException(status_code=413, detail=f"文件超过 {max_file_bytes // 1024 // 1024} MiB 限制")
    content_hash = hashlib.sha256(content).hexdigest()
    duplicate = _find_content_duplicate(content_hash)
    if duplicate or (seen_hashes is not None and content_hash in seen_hashes):
        if duplicate_action != "rename":
            raise DuplicateUploadError(filename, duplicate["name"] if duplicate else "本次上传中的同内容文件", duplicate["id"] if duplicate else None)
        with db() as conn:
            occupied_rows = conn.execute(qmark("SELECT name FROM documents WHERE name LIKE ?"), (f"{Path(filename).stem}%{extension}",)).fetchall()
        filename = _duplicate_filename(filename, {str(row["name"]) for row in occupied_rows})
    document_id = f"doc_{uuid.uuid4().hex}"
    storage_key = _collision_safe_key(_storage_key(document_id, extension, category_id=category_id, source_name=filename), document_id)
    path = STORAGE_DIR / storage_key
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    mime_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    try:
        with db() as conn:
            category = conn.execute(qmark("SELECT id FROM categories WHERE id = ?"), (category_id,)).fetchone() if category_id is not None else None
            if category_id is not None and not category:
                raise HTTPException(status_code=400, detail="分类不存在")
            sql = qmark("INSERT INTO documents (id, name, original_name, extension, size_bytes, mime_type, category_id, storage_key, storage_path, content_hash, storage_status, index_status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'stored', 'pending')")
            conn.execute(sql, (document_id, filename, filename, extension, len(content), mime_type, category_id, storage_key, storage_key, content_hash))
            conn.execute(qmark("INSERT INTO index_jobs (document_id) VALUES (?)"), (document_id,))
    except Exception:
        path.unlink(missing_ok=True)
        raise
    return row_to_document(select_documents("d.id = ?", (document_id,))[0])


@app.on_event("startup")
def startup():
    init_db()
    backfill_content_hashes()
    ensure_embedding_version()


def backfill_content_hashes() -> None:
    """Populate fingerprints for documents created before duplicate checking existed."""
    rows = select_documents("d.content_hash IS NULL")
    updates: list[tuple[str, str]] = []
    for row in rows:
        path = STORAGE_DIR / row["storage_key"]
        try:
            if path.is_file():
                updates.append((hashlib.sha256(path.read_bytes()).hexdigest(), row["id"]))
        except OSError:
            continue
    if updates:
        with db() as conn:
            for content_hash, document_id in updates:
                conn.execute(qmark("UPDATE documents SET content_hash = ? WHERE id = ? AND content_hash IS NULL"), (content_hash, document_id))


def ensure_embedding_version() -> None:
    """Queue a one-time rebuild when local vector features change."""
    if get_setting("embedding_version") == EMBEDDING_VERSION:
        return
    with db() as conn:
        rows = conn.execute("SELECT id FROM documents").fetchall()
        for row in rows:
            document_id = row["id"]
            pending = conn.execute(qmark("SELECT 1 FROM index_jobs WHERE document_id = ? AND status IN ('pending', 'processing') LIMIT 1"), (document_id,)).fetchone()
            if not pending:
                conn.execute(qmark("INSERT INTO index_jobs (document_id) VALUES (?)"), (document_id,))
            conn.execute(qmark("UPDATE documents SET index_status = 'pending', index_error = NULL WHERE id = ?"), (document_id,))
        if is_postgres():
            conn.execute("INSERT INTO app_settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()", ("embedding_version", EMBEDDING_VERSION))
        else:
            conn.execute("INSERT INTO app_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP", ("embedding_version", EMBEDDING_VERSION))


@app.get("/api/health", response_model=Health)
def health():
    with db() as conn:
        conn.execute("SELECT 1").fetchone()
    return Health(status="ok", database="postgres" if is_postgres() else "sqlite", embedding_dimension=EMBEDDING_DIM)


@app.get("/api/storage/status")
def storage_status():
    layout = get_setting("storage_layout")
    if layout == "folders":
        behavior = "按源文件名保存；资料/归档下使用系统分类名分目录"
    else:
        behavior = "按源文件名保存到统一目录；分类只更新数据库归属"
    return {"layout": layout, "behavior": behavior, "storage_dir": str(STORAGE_DIR), "host_path": STORAGE_HOST_PATH, "database_host_path": POSTGRES_HOST_PATH, "migration": _read_migration_status()}


def sync_storage_status(rows):
    """Refresh storage flags for rows returned to the UI.

    The check is deliberately based on the persisted storage key and does not
    remove records or index data.  Updating the row in memory as well makes
    the response immediately reflect a file that was removed externally.
    """
    updates: list[tuple[str, str]] = []
    for row in rows:
        exists = (STORAGE_DIR / row["storage_key"]).is_file()
        next_status = "stored" if exists else "missing"
        if next_status != row["storage_status"]:
            row["storage_status"] = next_status
            updates.append((next_status, row["id"]))
    if updates:
        with db() as conn:
            for status, document_id in updates:
                conn.execute(qmark("UPDATE documents SET storage_status = ? WHERE id = ?"), (status, document_id))


@app.post("/api/storage/scan")
def scan_storage():
    """Check persisted files without deleting database records or index data."""
    rows = select_documents()
    missing_ids: list[str] = []
    restored_ids: list[str] = []
    with db() as conn:
        for row in rows:
            exists = (STORAGE_DIR / row["storage_key"]).is_file()
            previous = row["storage_status"]
            next_status = "stored" if exists else "missing"
            if next_status != previous:
                conn.execute(qmark("UPDATE documents SET storage_status = ? WHERE id = ?"), (next_status, row["id"]))
            if next_status == "missing":
                missing_ids.append(row["id"])
            elif previous == "missing":
                restored_ids.append(row["id"])
    return {"checked": len(rows), "missing_count": len(missing_ids), "missing_ids": missing_ids, "restored_count": len(restored_ids)}


@app.get("/api/storage/migration-status")
def storage_migration_status():
    return _read_migration_status()


@app.post("/api/storage/migrate")
def request_storage_migration(payload: dict):
    """Queue a host-side directory move requested from the settings page.

    PostgreSQL's data directory cannot be moved safely from inside the API
    container. The BAT watcher, started with the platform, stops the stack,
    copies and verifies both directories, updates .env, then starts Compose
    again. The request files are exchanged through the project mount.
    """
    try:
        database_path = _validate_host_path(payload.get("database_host_path"), "数据库目录")
        storage_path = _validate_host_path(payload.get("storage_host_path"), "文件目录")
    except (AttributeError, TypeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    status = _read_migration_status()
    if status.get("status") in {"queued", "processing"} or MIGRATION_REQUEST.exists():
        raise HTTPException(status_code=409, detail="已有存储迁移正在处理，请等待完成")
    if database_path == POSTGRES_HOST_PATH and storage_path == STORAGE_HOST_PATH:
        return {"status": "idle", "message": "数据库和文件目录未改变"}
    request_id = uuid.uuid4().hex
    try:
        _write_migration_text(MIGRATION_DB_PATH, database_path)
        _write_migration_text(MIGRATION_STORAGE_PATH, storage_path)
        _write_migration_text(MIGRATION_STATUS, json.dumps({
            "id": request_id,
            "status": "queued",
            "database_host_path": database_path,
            "storage_host_path": storage_path,
            "message": "等待启动平台.bat 的迁移监视器处理",
        }, ensure_ascii=False))
        _write_migration_text(MIGRATION_REQUEST, request_id)
    except OSError as exc:
        raise HTTPException(status_code=503, detail=f"无法写入迁移请求：{exc}") from exc
    return {"status": "queued", "id": request_id, "database_host_path": database_path, "storage_host_path": storage_path}


@app.get("/api/settings")
def read_settings():
    return get_settings()


@app.get("/api/settings/background")
def read_background():
    path = _background_path()
    if not path:
        return {"enabled": False, "url": None, "version": None}
    return {
        "enabled": True,
        "url": "/api/settings/background/file",
        "version": str(path.stat().st_mtime_ns),
        "filename": path.name,
    }


@app.post("/api/settings/background")
async def upload_background(file: UploadFile = File(...)):
    filename = file.filename or ""
    extension = Path(filename).suffix.lower()
    if extension not in BACKGROUND_TYPES:
        raise HTTPException(status_code=400, detail="背景图片仅支持 PNG、JPG、JPEG、WEBP 或 GIF")
    payload = await file.read(BACKGROUND_MAX_BYTES + 1)
    if len(payload) > BACKGROUND_MAX_BYTES:
        raise HTTPException(status_code=413, detail="背景图片不能超过 10 MB")
    if not payload or not _is_image_payload(payload, extension):
        raise HTTPException(status_code=400, detail="文件内容不是有效的图片")
    BACKGROUND_DIR.mkdir(parents=True, exist_ok=True)
    temporary = BACKGROUND_DIR / f".background-{uuid.uuid4().hex}{extension}"
    target = BACKGROUND_DIR / f"background{extension}"
    try:
        temporary.write_bytes(payload)
        for old_extension in BACKGROUND_TYPES:
            old_path = BACKGROUND_DIR / f"background{old_extension}"
            if old_path != target:
                old_path.unlink(missing_ok=True)
        temporary.replace(target)
    except OSError as exc:
        temporary.unlink(missing_ok=True)
        raise HTTPException(status_code=500, detail=f"背景图片保存失败：{exc}") from exc
    return {
        "enabled": True,
        "url": "/api/settings/background/file",
        "version": str(target.stat().st_mtime_ns),
        "filename": target.name,
    }


@app.delete("/api/settings/background")
def delete_background():
    removed = False
    for extension in BACKGROUND_TYPES:
        path = BACKGROUND_DIR / f"background{extension}"
        if path.exists():
            path.unlink(missing_ok=True)
            removed = True
    return {"enabled": False, "removed": removed}


@app.get("/api/settings/background/file")
def background_file():
    path = _background_path()
    if not path:
        raise HTTPException(status_code=404, detail="尚未设置背景图片")
    return FileResponse(path, media_type=BACKGROUND_TYPES[path.suffix.lower()], headers={"Content-Disposition": "inline"})


@app.put("/api/settings")
def update_settings(payload: dict):
    try:
        previous_layout = get_setting("storage_layout")
        saved = set_settings(payload)
        if saved["storage_layout"] != previous_layout or payload.get("storage_layout") is not None:
            migrate_storage_layout()
        return saved
    except (ValueError, TypeError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/categories", response_model=list[Category])
def list_categories():
    with db() as conn:
        return [Category(id=row["id"], name=row["name"], parent_id=row["parent_id"]) for row in conn.execute("SELECT id, name, parent_id FROM categories ORDER BY name").fetchall()]


@app.get("/api/database/overview")
def database_overview():
    """A safe, read-only metadata view for the database button in the UI."""
    with db() as conn:
        categories = [dict(row) for row in conn.execute("SELECT id, name, parent_id, created_at FROM categories ORDER BY name").fetchall()]
        documents = [dict(row) for row in conn.execute("SELECT id, name, original_name, extension, size_bytes, storage_path, category_id, index_status, archived, uploaded_at FROM documents ORDER BY uploaded_at DESC").fetchall()]
    for row in documents:
        row["archived"] = bool(row["archived"])
    return {"database": "postgres" if is_postgres() else "sqlite", "categories": categories, "documents": documents}


@app.post("/api/categories", response_model=Category, status_code=201)
def create_category(payload: CategoryCreate):
    name = payload.name.strip()
    if payload.parent_id is not None:
        with db() as conn:
            if not conn.execute(qmark("SELECT id FROM categories WHERE id = ?"), (payload.parent_id,)).fetchone():
                raise HTTPException(status_code=400, detail="父分类不存在")
    try:
        with db() as conn:
            if is_postgres():
                row = conn.execute("INSERT INTO categories (name, parent_id) VALUES (%s, %s) RETURNING id, name, parent_id", (name, payload.parent_id)).fetchone()
            else:
                cur = conn.execute("INSERT INTO categories (name, parent_id) VALUES (?, ?)", (name, payload.parent_id))
                row = {"id": cur.lastrowid, "name": name, "parent_id": payload.parent_id}
    except Exception as exc:
        raise HTTPException(status_code=409, detail="分类已存在") from exc
    return Category(id=row["id"], name=row["name"], parent_id=row["parent_id"])


@app.delete("/api/categories/{category_id}")
def delete_category(category_id: int):
    """Delete a category and move all of its documents to the default category."""
    with db() as conn:
        category = conn.execute(qmark("SELECT id FROM categories WHERE id = ?"), (category_id,)).fetchone()
    if not category:
        raise HTTPException(status_code=404, detail="分类不存在")
    rows = select_documents("d.category_id = ?", (category_id,))
    moved: list[tuple[str, str]] = []
    try:
        if get_setting("storage_layout") == "folders":
            for row in rows:
                old_key = row["storage_key"]
                new_key = _collision_safe_key(_storage_key(row["id"], row["extension"], None, bool(row["archived"]), _existing_storage_name(row)), row["id"])
                if old_key == new_key:
                    continue
                source, target = STORAGE_DIR / old_key, STORAGE_DIR / new_key
                if source.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(source), str(target))
                    moved.append((old_key, new_key))
        with db() as conn:
            for row in rows:
                new_key = _collision_safe_key(_storage_key(row["id"], row["extension"], None, bool(row["archived"]), _existing_storage_name(row)), row["id"])
                conn.execute(qmark("UPDATE documents SET category_id = NULL, storage_key = ?, storage_path = ? WHERE id = ?"), (new_key, new_key, row["id"]))
            conn.execute(qmark("DELETE FROM categories WHERE id = ?"), (category_id,))
    except Exception:
        for old_key, new_key in reversed(moved):
            source, target = STORAGE_DIR / new_key, STORAGE_DIR / old_key
            if source.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(source), str(target))
        raise HTTPException(status_code=500, detail="删除分类失败，文件已恢复原路径")
    return {"deleted": category_id, "documents_moved_to_default": len(rows)}


@app.patch("/api/categories/{category_id}", response_model=Category)
def rename_category(category_id: int, payload: CategoryCreate):
    name = payload.name.strip()
    with db() as conn:
        if not conn.execute(qmark("SELECT id FROM categories WHERE id = ?"), (category_id,)).fetchone():
            raise HTTPException(status_code=404, detail="分类不存在")
        if payload.parent_id == category_id:
            raise HTTPException(status_code=400, detail="分类不能成为自己的父分类")
        if payload.parent_id is not None and not conn.execute(qmark("SELECT id FROM categories WHERE id = ?"), (payload.parent_id,)).fetchone():
            raise HTTPException(status_code=400, detail="父分类不存在")
        ancestor = payload.parent_id
        visited: set[int] = set()
        while ancestor is not None and ancestor not in visited:
            if ancestor == category_id:
                raise HTTPException(status_code=400, detail="父分类不能是当前分类的下级分类")
            visited.add(ancestor)
            parent_row = conn.execute(qmark("SELECT parent_id FROM categories WHERE id = ?"), (ancestor,)).fetchone()
            ancestor = parent_row["parent_id"] if parent_row else None
        try:
            conn.execute(qmark("UPDATE categories SET name = ?, parent_id = ? WHERE id = ?"), (name, payload.parent_id, category_id))
        except Exception as exc:
            raise HTTPException(status_code=409, detail="分类名称已存在") from exc
        row = conn.execute(qmark("SELECT id, name, parent_id FROM categories WHERE id = ?"), (category_id,)).fetchone()
    if get_setting("storage_layout") == "folders":
        migrate_storage_layout()
    return Category(id=row["id"], name=row["name"], parent_id=row["parent_id"])


@app.get("/api/documents", response_model=list[Document])
def list_documents(q: str = "", category_id: int | None = None, archived: bool = False):
    clauses = ["d.archived = ?"]
    args: list = [archived]
    if q.strip():
        clauses.append("(LOWER(d.name) LIKE ? OR LOWER(d.original_name) LIKE ?)")
        needle = f"%{q.strip().lower()}%"
        args.extend([needle, needle])
    if category_id is not None:
        clauses.append("d.category_id = ?")
        args.append(category_id)
    # Synchronize every record, including documents hidden by the current
    # category/archive filter, then return the requested view.
    sync_storage_status(select_documents())
    rows = select_documents(" AND ".join(clauses), tuple(args))
    # Keep the visible list truthful even when a user removes a stored file
    # outside the platform.  This is intentionally a status update only: the
    # document record and its index remain available until the user removes
    # the index explicitly.
    sync_storage_status(rows)
    return [row_to_document(r) for r in rows]


@app.post("/api/documents", response_model=Document, status_code=201)
async def upload_document(file: UploadFile = File(...), category_id: int | None = Form(None), duplicate_action: str = Form("prompt")):
    try:
        return await _store_uploaded_file(file, category_id, duplicate_action)
    except DuplicateUploadError as duplicate:
        raise HTTPException(status_code=409, detail={"code": "duplicate", "filename": duplicate.filename, "existing_name": duplicate.existing_name, "existing_id": duplicate.existing_id})


@app.post("/api/documents/batch")
async def upload_documents_batch(files: list[UploadFile] = File(...), category_id: int | None = Form(None), duplicate_action: str = Form("prompt")):
    if not files:
        raise HTTPException(status_code=400, detail="至少选择一个文件")
    if duplicate_action not in {"prompt", "skip", "rename"}:
        raise HTTPException(status_code=400, detail="duplicate_action 必须是 prompt、skip 或 rename")
    uploaded: list[Document] = []
    failed: list[dict[str, str]] = []
    duplicates: list[dict] = []
    seen_hashes: set[str] = set()
    for file_index, file in enumerate(files):
        try:
            content = await file.read()
            uploaded.append(await _store_uploaded_file(file, category_id, duplicate_action, content, seen_hashes))
            seen_hashes.add(hashlib.sha256(content).hexdigest())
        except DuplicateUploadError as duplicate:
            duplicates.append({"index": file_index, "name": duplicate.filename, "existing_name": duplicate.existing_name, "existing_id": duplicate.existing_id})
        except HTTPException as exc:
            failed.append({"name": file.filename or "未命名文件", "error": str(exc.detail)})
        except Exception as exc:
            failed.append({"name": file.filename or "未命名文件", "error": f"保存失败：{exc}"})
    return {"uploaded": uploaded, "failed": failed, "duplicates": duplicates, "total": len(files)}


@app.get("/api/documents/download-batch")
def download_documents_batch(ids: list[str] = Query(min_length=1)):
    placeholders = ",".join("?" for _ in ids)
    rows = select_documents(f"d.id IN ({placeholders})", tuple(ids))
    if not rows:
        raise HTTPException(status_code=404, detail="没有找到可下载的文件")
    total = 0
    archive = io.BytesIO()
    used_names: set[str] = set()
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
        for row in rows:
            path = STORAGE_DIR / row["storage_key"]
            if not path.exists():
                continue
            total += path.stat().st_size
            if total > int(get_setting("max_batch_download_bytes")):
                raise HTTPException(status_code=413, detail="批量下载总大小超过 1 GiB 限制")
            name = row["original_name"] or row["name"]
            stem = Path(name).stem
            suffix = Path(name).suffix
            candidate = name
            counter = 2
            while candidate in used_names:
                candidate = f"{stem} ({counter}){suffix}"
                counter += 1
            used_names.add(candidate)
            bundle.write(path, arcname=candidate)
    if not used_names:
        raise HTTPException(status_code=404, detail="选中的文件原始内容不可用")
    archive.seek(0)
    headers = {"Content-Disposition": "attachment; filename=documents.zip"}
    return StreamingResponse(archive, media_type="application/zip", headers=headers)


@app.get("/api/documents/{document_id}", response_model=Document)
def get_document(document_id: str):
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    return row_to_document(rows[0])


@app.delete("/api/documents/batch")
def delete_documents_batch(payload: DocumentBatchDelete):
    """Permanently remove selected documents, their bytes, and indexed chunks."""
    placeholders = ",".join("?" for _ in payload.ids)
    rows = select_documents(f"d.id IN ({placeholders})", tuple(payload.ids))
    if not rows:
        raise HTTPException(status_code=404, detail="没有找到可移除的文件")
    for row in rows:
        (STORAGE_DIR / row["storage_key"]).unlink(missing_ok=True)
    with db() as conn:
        conn.execute(qmark(f"DELETE FROM documents WHERE id IN ({placeholders})"), tuple(row["id"] for row in rows))
    return {"removed": len(rows), "missing": len(payload.ids) - len(rows)}


@app.patch("/api/documents/batch/rename")
def rename_documents_batch(payload: DocumentBatchRename):
    """Batch-update source names and optionally extensions, moving platform files safely."""
    if not any((payload.find_text, payload.replace_text, payload.prefix, payload.suffix, payload.extension)):
        raise HTTPException(status_code=400, detail="请填写文件名替换、前缀、后缀或扩展名")
    placeholders = ",".join("?" for _ in payload.ids)
    rows = select_documents(f"d.id IN ({placeholders})", tuple(payload.ids))
    if not rows:
        raise HTTPException(status_code=404, detail="没有找到可修改的文件")
    if len(rows) != len(set(payload.ids)):
        raise HTTPException(status_code=404, detail="部分文件不存在，请刷新资料列表后重试")
    requested_extension = payload.extension.strip().lower() if payload.extension is not None else None
    if requested_extension:
        if not requested_extension.startswith("."):
            requested_extension = "." + requested_extension
        if requested_extension not in ALLOWED_EXTENSIONS:
            raise HTTPException(status_code=400, detail=f"不支持的扩展名 {requested_extension}，请从支持的类型中选择")
    plans: list[tuple[object, str, str, str, str]] = []
    planned_keys: set[str] = set()
    moved: list[tuple[Path, Path]] = []
    try:
        for row in rows:
            old_name = row["original_name"] or row["name"]
            old_path = Path(old_name)
            extension = requested_extension or row["extension"]
            stem = old_path.stem
            if payload.find_text:
                stem = stem.replace(payload.find_text, payload.replace_text)
            stem = f"{payload.prefix}{stem}{payload.suffix}"
            new_name = _safe_storage_name(f"{stem}{extension}", row["id"], extension)
            new_key = _storage_key(row["id"], extension, row["category_id"], bool(row["archived"]), new_name)
            if new_key != row["storage_key"]:
                new_key = _collision_safe_key(new_key, row["id"])
                if new_key in planned_keys:
                    raise HTTPException(status_code=409, detail=f"批量修改后出现同名文件：{new_name}")
                source = STORAGE_DIR / row["storage_key"]
                target = STORAGE_DIR / new_key
                if not source.exists():
                    raise HTTPException(status_code=404, detail=f"原始文件不存在：{old_name}")
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(source), str(target))
                moved.append((source, target))
            planned_keys.add(new_key)
            mime_type = mimetypes.guess_type(new_name)[0] or row["mime_type"]
            plans.append((row, new_name, extension, new_key, mime_type))
        with db() as conn:
            for row, new_name, extension, new_key, mime_type in plans:
                conn.execute(qmark("UPDATE documents SET name = ?, original_name = ?, extension = ?, mime_type = ?, storage_key = ?, storage_path = ?, index_status = 'pending', index_error = NULL, content_version = content_version + 1 WHERE id = ?"), (new_name, new_name, extension, mime_type, new_key, new_key, row["id"]))
                conn.execute(qmark("DELETE FROM document_chunks WHERE document_id = ?"), (row["id"],))
                conn.execute(qmark("DELETE FROM index_jobs WHERE document_id = ? AND status != 'processing'"), (row["id"],))
                conn.execute(qmark("INSERT INTO index_jobs (document_id) VALUES (?)"), (row["id"],))
    except HTTPException:
        for source, target in reversed(moved):
            if target.exists():
                source.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(target), str(source))
        raise
    except Exception as exc:
        for source, target in reversed(moved):
            if target.exists():
                source.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(target), str(source))
        raise HTTPException(status_code=500, detail=f"批量修改文件名失败：{exc}") from exc
    return {"updated": [row_to_document(select_documents("d.id = ?", (row["id"],))[0]) for row, *_ in plans]}


@app.get("/api/documents/{document_id}/download")
def download_document(document_id: str):
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    row = rows[0]
    path = STORAGE_DIR / row["storage_key"]
    if not path.exists():
        raise HTTPException(status_code=404, detail="原始文件不存在")
    return FileResponse(path, media_type=row["mime_type"], filename=row["original_name"])


@app.get("/api/documents/{document_id}/preview")
def preview_document(document_id: str):
    """Return bounded plain text plus a rich preview descriptor."""
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    row = rows[0]
    path = STORAGE_DIR / row["storage_key"]
    if not path.exists():
        raise HTTPException(status_code=404, detail="原始文件不存在")
    text = extract_text(path, row["extension"])
    rich_type = None
    rich_html = None
    if row["extension"] == ".pdf":
        # The browser can render these formats directly. HTML is sandboxed by
        # the client preview so uploaded scripts cannot execute in the app UI.
        rich_type = "pdf"
    elif row["extension"] in {".html", ".htm"}:
        rich_type = "iframe"
    elif row["extension"] == ".docx":
        try:
            from docx import Document as WordDocument

            document = WordDocument(str(path))
            parts: list[str] = []
            for paragraph in document.paragraphs:
                style_name = (paragraph.style.name or "").lower()
                heading_level = next((level for level in range(1, 7) if f"heading {level}" in style_name), None)
                tag = f"h{heading_level}" if heading_level else "p"
                runs: list[str] = []
                for run in paragraph.runs:
                    value = html_escape(run.text or "")
                    if run.bold:
                        value = f"<strong>{value}</strong>"
                    if run.italic:
                        value = f"<em>{value}</em>"
                    runs.append(value)
                content = "".join(runs) or html_escape(paragraph.text or "")
                parts.append(f"<{tag}>{content}</{tag}>")
            for table in document.tables:
                rows_html = []
                for table_row in table.rows:
                    cells = "".join(f"<td>{html_escape(cell.text or '')}</td>" for cell in table_row.cells)
                    rows_html.append(f"<tr>{cells}</tr>")
                if rows_html:
                    parts.append(f"<table><tbody>{''.join(rows_html)}</tbody></table>")
            candidate = "".join(parts)
            rich_html = candidate[:100000]
            rich_type = "html" if rich_html else None
        except Exception:
            # Text extraction remains available even if the optional rich
            # conversion cannot parse a particular DOCX file.
            rich_type = None
    elif row["extension"] == ".pptx":
        try:
            rich_html = _pptx_preview_html(path)
            rich_type = "pptx" if rich_html else None
        except Exception:
            rich_type = None
    elif row["extension"] in {".xlsx", ".xls"}:
        try:
            rich_html = _spreadsheet_preview_html(path, row["extension"])
            rich_type = "sheet" if rich_html else None
        except Exception:
            rich_type = None
    if not text:
        return {
            "document": row_to_document(row),
            "previewable": False,
            "text": "",
            "truncated": False,
            "reason": "该文件没有可提取的文本内容",
            "rich_available": bool(rich_type),
            "rich_type": rich_type,
            "rich_html": rich_html,
        }
    limit = 50000
    return {
        "document": row_to_document(row),
        "previewable": True,
        "text": text[:limit],
        "truncated": len(text) > limit,
        "rich_available": bool(rich_type),
        "rich_type": rich_type,
        "rich_html": rich_html,
    }


@app.get("/api/documents/{document_id}/preview-file")
def preview_file(document_id: str):
    """Stream browser-renderable originals with inline disposition."""
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    row = rows[0]
    if row["extension"] not in {".pdf", ".html", ".htm", ".pptx", ".xlsx", ".xls"}:
        raise HTTPException(status_code=415, detail="该文件不支持原文内嵌预览")
    path = STORAGE_DIR / row["storage_key"]
    if not path.exists():
        raise HTTPException(status_code=404, detail="原始文件不存在")
    filename = quote(row["original_name"] or row["name"])
    return FileResponse(
        path,
        media_type=row["mime_type"],
        headers={"Content-Disposition": f"inline; filename*=UTF-8''{filename}"},
    )


@app.get("/api/documents/{document_id}/preview-data")
def preview_data(document_id: str):
    """Return PDF bytes as JSON so download managers cannot hijack preview fetches."""
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    row = rows[0]
    if row["extension"] != ".pdf":
        raise HTTPException(status_code=415, detail="该文件不是 PDF")
    path = STORAGE_DIR / row["storage_key"]
    if not path.exists():
        raise HTTPException(status_code=404, detail="原始文件不存在")
    return {
        "data": base64.b64encode(path.read_bytes()).decode("ascii"),
        "mime_type": "application/pdf",
        "filename": row["original_name"] or row["name"],
    }


@app.patch("/api/documents/{document_id}", response_model=Document)
def update_document(document_id: str, payload: DocumentUpdate):
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    fields, args = [], []
    category_requested = "category_id" in getattr(payload, "model_fields_set", getattr(payload, "__fields_set__", set()))
    if category_requested:
        if payload.category_id is not None:
            with db() as conn:
                if not conn.execute(qmark("SELECT id FROM categories WHERE id = ?"), (payload.category_id,)).fetchone():
                    raise HTTPException(status_code=400, detail="分类不存在")
        fields.append("category_id = ?"); args.append(payload.category_id)
    if payload.archived is not None:
        fields.append("archived = ?"); args.append(payload.archived)
    if fields:
        old_row = rows[0]
        next_category = payload.category_id if category_requested else old_row["category_id"]
        next_archived = payload.archived if payload.archived is not None else bool(old_row["archived"])
        old_key = old_row["storage_key"]
        desired_key = _storage_key(document_id, old_row["extension"], next_category, next_archived, _existing_storage_name(old_row))
        storage_conflict = desired_key != old_key and (STORAGE_DIR / desired_key).exists()
        next_key = _collision_safe_key(desired_key, document_id)
        moved = False
        if get_setting("storage_layout") == "folders" and next_key != old_key:
            source = STORAGE_DIR / old_key
            target = STORAGE_DIR / next_key
            if source.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(source), str(target))
                moved = True
            fields.extend(["storage_key = ?", "storage_path = ?"])
            args.extend([next_key, next_key])
        args.append(document_id)
        try:
            with db() as conn:
                conn.execute(qmark(f"UPDATE documents SET {', '.join(fields)} WHERE id = ?"), tuple(args))
        except Exception:
            if moved:
                (STORAGE_DIR / old_key).parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(STORAGE_DIR / next_key), str(STORAGE_DIR / old_key))
            raise
    result = get_document(document_id)
    result.storage_conflict = storage_conflict if 'storage_conflict' in locals() else False
    return result


@app.post("/api/documents/{document_id}/index-retry", response_model=Document)
def retry_index(document_id: str):
    doc = get_document(document_id)
    if doc.index_status not in {"failed", "not_supported"}:
        return doc
    with db() as conn:
        conn.execute(qmark("UPDATE documents SET index_status = 'pending', index_error = NULL WHERE id = ?"), (document_id,))
        conn.execute(qmark("INSERT INTO index_jobs (document_id) VALUES (?)"), (document_id,))
    return get_document(document_id)


@app.delete("/api/documents/{document_id}/index", response_model=Document)
def delete_index(document_id: str):
    """Remove only indexed chunks/jobs; keep the document database record."""
    rows = select_documents("d.id = ?", (document_id,))
    if not rows:
        raise HTTPException(status_code=404, detail="文件不存在")
    with db() as conn:
        conn.execute(qmark("DELETE FROM document_chunks WHERE document_id = ?"), (document_id,))
        conn.execute(qmark("DELETE FROM index_jobs WHERE document_id = ? AND status != 'processing'"), (document_id,))
        conn.execute(qmark("UPDATE documents SET index_status = 'not_indexed', index_error = ? WHERE id = ?"), ("原文件不存在，已删除索引", document_id))
    return get_document(document_id)


def keyword_results(query: str, category_id: int | None, archived: bool, scope: str = "all"):
    normalized_query = re.sub(r"\s+", "", query.strip().lower())
    tokens = tokenize(query)
    ascii_tokens = [token for token in tokens if re.fullmatch(r"[a-z0-9_]+", token)]
    chinese_chars = "".join(token for token in tokens if re.fullmatch(r"[\u4e00-\u9fff]", token))
    chinese_bigrams = [chinese_chars[index:index + 2] for index in range(len(chinese_chars) - 1)]
    # Filename search is available immediately after the bytes are stored. Body
    # matching additionally uses document_chunks when the background index is ready.
    clauses = ["d.archived = ?"]
    args: list = [archived]
    if category_id is not None:
        clauses.append("d.category_id = ?"); args.append(category_id)
    rows = select_documents(" AND ".join(clauses), tuple(args))
    found = []
    for row in rows:
        with db() as conn:
            chunks = conn.execute(qmark("SELECT text FROM document_chunks WHERE document_id = ? ORDER BY chunk_index"), (row["id"],)).fetchall()
        text = "\n".join(chunk["text"] for chunk in chunks)
        filename_haystack = re.sub(r"\s+", "", f"{row['name']} {row['original_name']}".lower())
        content_haystack = re.sub(r"\s+", "", text.lower())
        haystack = filename_haystack if scope == "filename" else content_haystack if scope == "content" else f"{filename_haystack}{content_haystack}"
        phrase_match = bool(normalized_query and normalized_query in haystack)
        score = 1.0 if phrase_match else 0.0
        matched_terms: list[str] = []
        if phrase_match:
            matched_terms = [normalized_query]
        elif chinese_bigrams:
            # Chinese is tokenized as individual characters for embeddings. For
            # keyword search require most adjacent pairs, avoiding the old
            # "any two characters" false-positive behaviour on long queries.
            chinese_hits = [term for term in chinese_bigrams if term in haystack]
            ratio = len(chinese_hits) / len(chinese_bigrams)
            ascii_ok = not ascii_tokens or all(token in haystack for token in ascii_tokens)
            if ascii_ok and ratio >= (0.7 if len(chinese_bigrams) >= 3 else 1.0):
                matched_terms = chinese_hits + [token for token in ascii_tokens if token in haystack]
                score = ratio
            else:
                matched_terms = []
        elif ascii_tokens:
            matched_terms = [token for token in ascii_tokens if token in haystack]
            if len(matched_terms) == len(ascii_tokens):
                score = 0.8
            else:
                matched_terms = []
        if score > 0:
            snippets = []
            search_terms = [normalized_query] if phrase_match else matched_terms[:3]
            if scope == "filename":
                snippets = [row["name"]]
            else:
                for token in search_terms:
                    match = re.search(re.escape(token), text, re.IGNORECASE)
                    if match:
                        snippets.append(text[max(0, match.start() - 90):match.end() + 180].replace("\n", " "))
            found.append(SearchResult(document=row_to_document(row), score=round(score, 4), snippets=snippets, match_type=scope))
    return sorted(found, key=lambda item: item.score or 0, reverse=True)


# The local hash vector is deliberately dependency-free, but a vector alone
# cannot know that common Chinese product terms such as “隐藏”和“归档” are
# related. These small, transparent expansions make local semantic search useful
# for paraphrases while preserving the real vector retrieval as one signal.
SEMANTIC_EXPANSIONS = {
    "代码提交": ("代码", "提交", "发布", "上线", "检查", "核对", "复核", "验证", "质量"),
    "合并代码": ("合并", "代码", "提交", "审阅", "评审", "检查", "测试", "质量"),
    "提交前": ("提交", "提交前", "发布前", "审阅", "检查", "测试", "核对", "验证"),
    "提交": ("提交", "发布", "上线", "检查", "核对", "复核", "验证", "质量"),
    "搜不到": ("搜索", "检索", "索引", "可检索", "worker", "任务", "失败", "排查"),
    "正文搜不到": ("正文", "搜索", "检索", "索引", "可检索", "任务", "失败", "排查"),
    "不能搜索": ("搜索", "检索", "索引", "可检索", "任务", "失败", "排查"),
    "上传": ("上传", "保存", "文件", "资料", "存储"),
    "网络请求": ("请求", "网络", "重发", "重复", "幂等", "相同键", "记录"),
    "请求重放": ("请求", "重放", "重发", "重复", "幂等", "相同键", "记录"),
    "发两遍": ("重发", "重复", "幂等", "相同", "两份", "记录"),
    "两份记录": ("两份", "重复", "幂等", "相同", "记录"),
    "旧项目": ("旧项目", "归档", "隐藏", "恢复", "找回", "历史资料"),
    "旧项目资料": ("旧项目", "资料", "归档", "隐藏", "恢复", "找回", "历史"),
    "归档恢复": ("归档", "恢复", "找回", "历史", "资料", "默认列表"),
    "隐藏": ("隐藏", "归档", "恢复", "找回", "历史资料"),
    "找回": ("找回", "恢复", "归档", "历史", "资料"),
    "文件类型": ("文件类型", "文件格式", "格式", "筛选", "分类"),
    "筛选": ("筛选", "过滤", "分类", "格式"),
    "发布执行": ("发布", "执行", "发布检查", "清单", "规范", "记录", "回退"),
    "通用规范": ("通用", "规范", "发布检查", "执行记录", "关联资料"),
    "故障": ("故障", "失败", "异常", "问题", "排查", "修复"),
    "约稿": ("约稿", "稿约", "征稿", "投稿", "稿件", "征文", "投稿要求", "约稿要求"),
}


def semantic_terms(query: str) -> dict[str, float]:
    compact = re.sub(r"\s+", "", query.lower())
    terms: dict[str, float] = {}
    # Keep ASCII identifiers as whole terms and Chinese runs as weighted
    # bigrams, which handles small wording changes without matching every
    # individual character in a long question.
    for token in tokenize(query):
        if re.fullmatch(r"[a-z0-9_]+", token):
            terms[token] = max(terms.get(token, 0.0), 1.0)
    for run in re.findall(r"[\u4e00-\u9fff]{2,}", compact):
        for index in range(len(run) - 1):
            bigram = run[index:index + 2]
            terms[bigram] = max(terms.get(bigram, 0.0), 0.35)
            # Many Chinese compounds have a natural reverse form, such as
            # “约稿” and “稿约”. Keep this as a lower-weight semantic signal.
            reversed_bigram = bigram[::-1]
            if reversed_bigram != bigram:
                terms[reversed_bigram] = max(terms.get(reversed_bigram, 0.0), 0.28)
    # Longer phrases win when an expansion overlaps a shorter trigger.
    for phrase in sorted(SEMANTIC_EXPANSIONS, key=len, reverse=True):
        if phrase in compact:
            for term in SEMANTIC_EXPANSIONS[phrase]:
                terms[term] = max(terms.get(term, 0.0), 1.0)
    return terms


def semantic_lexical_score(query: str, text: str, terms: dict[str, float]) -> tuple[float, list[str]]:
    haystack = re.sub(r"\s+", "", text.lower())
    hits = [term for term, weight in terms.items() if term in haystack]
    total = sum(terms.values()) or 1.0
    score = sum(terms[term] for term in hits) / total
    # The final meaningful phrase in a Chinese question is often the subject
    # being looked up (for example, “内心” in “三角形内心”). Give that exact
    # phrase a small anchor bonus so a document containing it is not displaced
    # by unrelated documents that happen to share generic words like “数学”.
    anchor_terms = [run[-2:] for run in re.findall(r"[\u4e00-\u9fff]{2,}", re.sub(r"\s+", "", query))]
    anchor_hits = [term for term in anchor_terms if term in haystack]
    if anchor_hits:
        score = min(1.0, score + 0.22)
    if re.sub(r"\s+", "", query.lower()) in haystack:
        score = min(1.0, score + 0.45)
    return min(1.0, score), hits


def semantic_snippet(text: str, hits: list[str]) -> str:
    clean = text.replace("\n", " ").strip()
    for term in sorted(hits, key=len, reverse=True):
        match = re.search(re.escape(term), clean, re.IGNORECASE)
        if match:
            return clean[max(0, match.start() - 120):match.end() + 260]
    return clean[:420]


def _search_tokens(text: str) -> list[str]:
    """Tokenize Chinese and ASCII text for BM25 without an extra dependency."""
    compact = re.sub(r"\s+", "", text.lower())
    tokens = tokenize(text)
    for run in re.findall(r"[\u4e00-\u9fff]{2,}", compact):
        tokens.extend(run[index:index + 2] for index in range(len(run) - 1))
    return list(dict.fromkeys(token for token in tokens if token))


def _bm25_score(query_tokens: list[str], document_tokens: list[str], document_frequency: dict[str, int], total_documents: int, average_length: float) -> float:
    """Compute BM25 for one chunk using the standard k1/b parameters."""
    if not query_tokens or not document_tokens or not total_documents:
        return 0.0
    frequencies: dict[str, int] = {}
    for token in document_tokens:
        frequencies[token] = frequencies.get(token, 0) + 1
    length = len(document_tokens)
    norm_length = length / (average_length or 1.0)
    score = 0.0
    for token in query_tokens:
        tf = frequencies.get(token, 0)
        if not tf:
            continue
        df = document_frequency.get(token, 0)
        idf = math.log(1.0 + (total_documents - df + 0.5) / (df + 0.5))
        score += idf * (tf * 2.0) / (tf + 1.2 * (0.75 + 0.25 * norm_length))
    return score


def _rrf_rank(ranks: list[int], constant: int = 60) -> float:
    """Reciprocal Rank Fusion score for one document across recall lists."""
    return sum(1.0 / (constant + rank) for rank in ranks if rank > 0)


def _semantic_rerank_score(vector_score: float, bm25_score: float, rrf_score: float, lexical_score: float, title_score: float, exact_phrase: bool) -> float:
    """Small deterministic reranker applied after vector/BM25 fusion.

    This keeps ranking explainable and dependency-free while rewarding exact
    phrases, title evidence and lexical coverage after the two recall stages.
    """
    vector_part = max(0.0, min(1.0, (vector_score + 1.0) / 2.0))
    bm25_part = max(0.0, min(1.0, bm25_score / 8.0))
    rrf_part = max(0.0, min(1.0, rrf_score / 0.032))
    score = 0.28 * rrf_part + 0.23 * vector_part + 0.22 * bm25_part + 0.17 * lexical_score + 0.10 * title_score
    if exact_phrase:
        score += 0.12
    return min(1.0, score)


@app.get("/api/search", response_model=SearchResponse)
def search(q: str = Query(min_length=1), mode: str = "keyword", category_id: int | None = None, archived: bool = False):
    if mode in {"keyword", "filename", "content"}:
        results = keyword_results(q, category_id, archived, "all" if mode == "keyword" else mode)
        return SearchResponse(query=q, mode=mode, results=results)
    if mode != "semantic":
        raise HTTPException(status_code=400, detail="mode 仅支持 filename、content 或 semantic")

    # 1. Query embedding and candidate document scope.
    query_vector = embed(q)
    query_terms = semantic_terms(q)
    bm25_query_tokens = _search_tokens(q)
    # Only feed high-confidence semantic expansions into BM25.  Low-weight
    # Chinese bigrams remain useful for vector features but are too generic to
    # become keyword candidates for a long natural-language question.
    for term, weight in query_terms.items():
        if weight >= 0.8 and term not in bm25_query_tokens:
            bm25_query_tokens.append(term)
    clauses = ["d.archived = ?"]
    args: list = [archived]
    if category_id is not None:
        clauses.append("d.category_id = ?")
        args.append(category_id)
    rows = select_documents(" AND ".join(clauses), tuple(args))
    by_doc = {row["id"]: row for row in rows}
    if not by_doc:
        return SearchResponse(query=q, mode=mode, results=[])

    # 2. Load persisted chunks. The vector list is the ANN-compatible recall
    # boundary; argpartition can be introduced later without changing ranking.
    placeholders = ",".join("?" for _ in by_doc)
    with db() as conn:
        chunks = conn.execute(qmark(f"SELECT document_id, chunk_index, text, vector FROM document_chunks WHERE document_id IN ({placeholders})"), tuple(by_doc)).fetchall()
    if not chunks:
        return SearchResponse(query=q, mode=mode, results=[])

    prepared: list[dict] = []
    document_frequency: dict[str, int] = {}
    total_length = 0
    for chunk in chunks:
        body_text = str(chunk["text"] or "")
        row = by_doc[chunk["document_id"]]
        title_text = f"{row['name']} {row['original_name']}"
        text = f"{title_text}\n{body_text}"
        tokens = _search_tokens(text)
        for token in set(tokens):
            document_frequency[token] = document_frequency.get(token, 0) + 1
        total_length += len(tokens)
        vector = chunk["vector"] if isinstance(chunk["vector"], list) else json.loads(chunk["vector"])
        prepared.append({
            "document_id": chunk["document_id"],
            "chunk_index": chunk["chunk_index"],
            "text": body_text,
            "title": title_text,
            "tokens": tokens,
            "vector_score": cosine(query_vector, vector),
        })
    average_length = total_length / max(1, len(prepared))

    # 3. Independent recalls: vector similarity and BM25 over the same chunks.
    for item in prepared:
        item["bm25_score"] = _bm25_score(bm25_query_tokens, item["tokens"], document_frequency, len(prepared), average_length)
    vector_ranked = sorted(prepared, key=lambda item: item["vector_score"], reverse=True)[:80]
    bm25_ranked = sorted((item for item in prepared if item["bm25_score"] > 0), key=lambda item: item["bm25_score"], reverse=True)[:80]

    # 4. RRF at document level, retaining the best chunk from each recall list.
    vector_doc_rank: dict[str, int] = {}
    bm25_doc_rank: dict[str, int] = {}
    for rank, item in enumerate(vector_ranked, 1):
        vector_doc_rank.setdefault(item["document_id"], rank)
    for rank, item in enumerate(bm25_ranked, 1):
        bm25_doc_rank.setdefault(item["document_id"], rank)
    candidate_ids = set(vector_doc_rank) | set(bm25_doc_rank)
    candidates: list[dict] = []
    query_compact = re.sub(r"\s+", "", q.lower())
    for document_id in candidate_ids:
        items = [item for item in prepared if item["document_id"] == document_id]
        best = max(items, key=lambda item: (item["vector_score"], item["bm25_score"]))
        body_score, body_hits = semantic_lexical_score(q, best["text"], query_terms)
        title_score, title_hits = semantic_lexical_score(q, best["title"], query_terms)
        body_haystack = re.sub(r"\s+", "", best["text"].lower())
        title_haystack = re.sub(r"\s+", "", best["title"].lower())
        exact_phrase = bool(query_compact and (query_compact in body_haystack or query_compact in title_haystack))
        rrf_score = _rrf_rank([vector_doc_rank.get(document_id, 0), bm25_doc_rank.get(document_id, 0)])
        lexical_score = max(body_score, title_score)
        rerank_score = _semantic_rerank_score(best["vector_score"], best["bm25_score"], rrf_score, lexical_score, title_score, exact_phrase)
        candidates.append({
            "document_id": document_id,
            "score": rerank_score,
            "rrf": rrf_score,
            "text": best["text"] if body_score >= title_score else f"文件名：{by_doc[document_id]['name']}\n{best['text']}",
            "hits": list(dict.fromkeys(body_hits + title_hits + bm25_query_tokens[:4])),
            "bm25": best["bm25_score"],
        })

    # 5. Reranked results, with a relative floor to suppress weak hash hits.
    candidates.sort(key=lambda item: item["score"], reverse=True)
    top_score = candidates[0]["score"] if candidates else 0.0
    results = [
        SearchResult(
            document=row_to_document(by_doc[item["document_id"]]),
            score=round(item["score"], 4),
            snippets=[semantic_snippet(item["text"], item["hits"])],
            match_type="semantic",
        )
        for item in candidates
        if item["score"] >= max(0.20, top_score * 0.62) and (item["bm25"] > 0 or item["rrf"] > 0)
    ][:50]
    return SearchResponse(query=q, mode=mode, results=results)


@app.post("/api/documents/{document_id}/index-now", response_model=Document, include_in_schema=False)
def index_now(document_id: str):
    get_document(document_id)
    index_document(document_id)
    return get_document(document_id)
