import json
import re

from .config import EMBEDDING_API_KEY, EMBEDDING_API_URL, EMBEDDING_MODEL, EMBEDDING_PROVIDER, INDEX_POLL_SECONDS, MAX_BATCH_DOWNLOAD_BYTES, MAX_FILE_BYTES, STORAGE_LAYOUT
from .db import db, is_postgres

# Kept in the database so the browser settings panel survives container restarts.
DEFAULTS = {
    "storage_layout": STORAGE_LAYOUT,
    "embedding_provider": EMBEDDING_PROVIDER,
    "embedding_api_url": EMBEDDING_API_URL,
    "embedding_model": EMBEDDING_MODEL,
    "max_file_bytes": str(MAX_FILE_BYTES),
    "max_batch_download_bytes": str(MAX_BATCH_DOWNLOAD_BYTES),
    "index_poll_seconds": str(INDEX_POLL_SECONDS),
    "sort_method": "modified_desc",
    "page_size": "20",
    "semantic_result_limit": "5",
    "workspace_opacity": "0.96",
    "background_opacity": "0.38",
    "accent_color": "#173b45",
    "accent_opacity": "1",
    "modal_animation": "fade",
    "title_text": "文件管理与知识检索平台",
    "eyebrow_text": "Documents Workspace",
    "eyebrow_color": "#9ad4c8",
    "subtitle_text": "Made By KOKONA",
    "title_color": "#f5fbfa",
    "subtitle_color": "#9fc4c9",
    "embedding_version": "local-hash-semantic-v3",
}


def _sql(sql: str) -> str:
    return sql.replace("?", "%s") if is_postgres() else sql


def get_setting(key: str) -> str:
    with db() as conn:
        row = conn.execute(_sql("SELECT value FROM app_settings WHERE key = ?"), (key,)).fetchone()
    return row["value"] if row else DEFAULTS.get(key, "")


def get_settings() -> dict:
    with db() as conn:
        rows = conn.execute("SELECT key, value FROM app_settings").fetchall()
    values = DEFAULTS | {row["key"]: row["value"] for row in rows}
    return {
        "storage_layout": values["storage_layout"],
        "embedding_provider": values["embedding_provider"],
        "embedding_api_url": values["embedding_api_url"],
        "embedding_model": values["embedding_model"],
        "embedding_api_key_set": bool(get_setting("embedding_api_key") or EMBEDDING_API_KEY),
        "max_file_bytes": int(values["max_file_bytes"]),
        "max_batch_download_bytes": int(values["max_batch_download_bytes"]),
        "index_poll_seconds": float(values["index_poll_seconds"]),
        "sort_method": values["sort_method"],
        "page_size": int(values.get("page_size", "20")),
        "semantic_result_limit": int(values.get("semantic_result_limit", "5")),
        "workspace_opacity": float(values.get("workspace_opacity", "0.96")),
        "background_opacity": float(values.get("background_opacity", "0.38")),
        "accent_color": values.get("accent_color", "#173b45"),
        "accent_opacity": float(values.get("accent_opacity", "1")),
        "modal_animation": values.get("modal_animation", "fade"),
        "title_text": values.get("title_text", "文件管理与知识检索平台"),
        "eyebrow_text": values.get("eyebrow_text", "Documents Workspace"),
        "eyebrow_color": values.get("eyebrow_color", "#9ad4c8"),
        "subtitle_text": values.get("subtitle_text", "Made By KOKONA"),
        "title_color": values.get("title_color", "#f5fbfa"),
        "subtitle_color": values.get("subtitle_color", "#9fc4c9"),
    }


def set_settings(payload: dict) -> dict:
    allowed = {"storage_layout", "embedding_provider", "embedding_api_url", "embedding_model", "max_file_bytes", "max_batch_download_bytes", "index_poll_seconds", "embedding_api_key", "sort_method", "page_size", "semantic_result_limit", "workspace_opacity", "background_opacity", "accent_color", "accent_opacity", "modal_animation", "title_text", "eyebrow_text", "eyebrow_color", "subtitle_text", "title_color", "subtitle_color"}
    values = {key: value for key, value in payload.items() if key in allowed and value is not None}
    if values.get("storage_layout") not in {None, "flat", "folders"}:
        raise ValueError("storage_layout 仅支持 flat 或 folders")
    if values.get("embedding_provider") not in {None, "local", "remote", "openai", "openai_compatible"}:
        raise ValueError("embedding_provider 不受支持")
    for key in ("max_file_bytes", "max_batch_download_bytes"):
        if key in values and (int(values[key]) < 1 or int(values[key]) > 10 * 1024 * 1024 * 1024):
            raise ValueError(f"{key} 超出范围")
    if "index_poll_seconds" in values and not 0.2 <= float(values["index_poll_seconds"]) <= 60:
        raise ValueError("index_poll_seconds 应在 0.2 到 60 秒之间")
    if "sort_method" in values and values["sort_method"] not in {"name_asc", "name_desc", "size_asc", "size_desc", "modified_asc", "modified_desc", "category_asc", "category_desc", "status_asc", "status_desc"}:
        raise ValueError("sort_method 不受支持")
    if "page_size" in values and int(values["page_size"]) not in {0, 10, 20, 50, 100, 200}:
        raise ValueError("page_size 仅支持 0、10、20、50、100 或 200；0 表示不分页")
    if "semantic_result_limit" in values and not 1 <= int(values["semantic_result_limit"]) <= 50:
        raise ValueError("semantic_result_limit 应在 1 到 50 之间")
    for key in ("workspace_opacity", "background_opacity", "accent_opacity"):
        if key in values and not 0 <= float(values[key]) <= 1:
            raise ValueError(f"{key} 应在 0 到 1 之间")
    if "accent_color" in values:
        color = str(values["accent_color"]).strip()
        if not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
            raise ValueError("accent_color 必须是六位十六进制颜色")
        values["accent_color"] = color
    if "modal_animation" in values and values["modal_animation"] not in {"fade", "scale", "slide", "none"}:
        raise ValueError("modal_animation 不受支持")
    for key in ("title_color", "eyebrow_color", "subtitle_color"):
        if key in values:
            color = str(values[key]).strip()
            if not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
                raise ValueError(f"{key} 必须是六位十六进制颜色")
            values[key] = color
    for key in ("title_text", "eyebrow_text", "subtitle_text"):
        if key in values:
            values[key] = str(values[key]).strip()[:120]
    with db() as conn:
        for key, value in values.items():
            value = str(value)
            if is_postgres():
                conn.execute("INSERT INTO app_settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()", (key, value))
            else:
                conn.execute("INSERT INTO app_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP", (key, value))
    return get_settings()
