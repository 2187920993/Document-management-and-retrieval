import os
from pathlib import Path


DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./data/app.db")
STORAGE_DIR = Path(os.getenv("STORAGE_DIR", "./storage"))
STORAGE_HOST_PATH = os.getenv("STORAGE_HOST_PATH", "./runtime/storage")
POSTGRES_HOST_PATH = os.getenv("POSTGRES_HOST_PATH", "./runtime/postgres")
# The compose project is mounted here so the host-side BAT watcher can receive
# a migration request from the settings page and safely restart the stack.
PROJECT_ROOT = Path(os.getenv("PROJECT_ROOT", "/project"))
MAX_FILE_BYTES = int(os.getenv("MAX_FILE_BYTES", 20 * 1024 * 1024))
MAX_BATCH_DOWNLOAD_BYTES = int(os.getenv("MAX_BATCH_DOWNLOAD_BYTES", 1024 * 1024 * 1024))
EMBEDDING_DIM = int(os.getenv("EMBEDDING_DIM", "384"))
EMBEDDING_PROVIDER = os.getenv("EMBEDDING_PROVIDER", "local").lower()
EMBEDDING_API_URL = os.getenv("EMBEDDING_API_URL", "").strip()
EMBEDDING_API_KEY = os.getenv("EMBEDDING_API_KEY", "")
EMBEDDING_MODEL = os.getenv("EMBEDDING_MODEL", "").strip()
INDEX_POLL_SECONDS = float(os.getenv("INDEX_POLL_SECONDS", "1"))
STORAGE_LAYOUT = os.getenv("STORAGE_LAYOUT", "flat").lower()

# Formats accepted by the platform. Text-like formats are indexed directly; PDF and
# DOCX use their extractors when available, while other accepted binary formats can
# still be stored and downloaded with an explicit not_supported index status.
ALLOWED_EXTENSIONS = {
    ".pdf", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".txt", ".md", ".markdown", ".log", ".csv", ".tsv", ".json",
    ".yaml", ".yml", ".xml", ".html", ".htm", ".css", ".js", ".jsx", ".ts", ".tsx",
    ".py", ".java", ".go", ".rs", ".sql", ".ini", ".conf", ".toml", ".vue", ".svelte",
}

TEXT_EXTENSIONS = ALLOWED_EXTENSIONS - {".pdf", ".docx", ".ppt", ".pptx", ".xls", ".xlsx"}
