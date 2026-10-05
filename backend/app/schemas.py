from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


class CategoryCreate(BaseModel):
    name: str = Field(min_length=1, max_length=100)
    parent_id: int | None = None


class Category(BaseModel):
    id: int
    name: str
    parent_id: int | None = None


class DocumentUpdate(BaseModel):
    category_id: int | None = None
    archived: bool | None = None
    tags: list[str] = []


class DocumentBatchDelete(BaseModel):
    ids: list[str] = Field(min_length=1, max_length=100)


class DocumentBatchRename(BaseModel):
    ids: list[str] = Field(min_length=1, max_length=100)
    find_text: str = Field(default="", max_length=200)
    replace_text: str = Field(default="", max_length=200)
    prefix: str = Field(default="", max_length=100)
    suffix: str = Field(default="", max_length=100)
    extension: str | None = Field(default=None, max_length=20)


class Document(BaseModel):
    id: str
    name: str
    original_name: str
    extension: str
    size_bytes: int
    mime_type: str
    storage_path: str | None = None
    storage_conflict: bool = False
    category_id: int | None
    category_name: str | None = None
    storage_status: str
    index_status: str
    index_error: str | None = None
    archived: bool
    uploaded_at: datetime | str


class SearchResult(BaseModel):
    document: Document
    score: float | None = None
    snippets: list[str] = []
    match_type: str


class SearchResponse(BaseModel):
    query: str
    mode: str
    results: list[SearchResult]


class Health(BaseModel):
    status: str
    database: str
    embedding_dimension: int
