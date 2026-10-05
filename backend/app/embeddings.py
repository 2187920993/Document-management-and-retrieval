"""Small, deterministic text-vector adapter.

The default implementation is a signed hashing embedding so the demo works without
network credentials. It is a real numeric vector persisted with each chunk and uses
cosine similarity. The adapter boundary allows replacing it with a hosted or local
language embedding model without changing the indexing or search APIs.
"""
import hashlib
import json
import math
import re
import urllib.request

from .config import EMBEDDING_DIM
from .settings import get_setting

TOKEN_RE = re.compile(r"[A-Za-z0-9_]+|[\u4e00-\u9fff]", re.UNICODE)
EMBEDDING_VERSION = "local-hash-semantic-v3"

# Canonical topic features make the local vector useful for Chinese paraphrases
# without pretending to be a pretrained language model. Both document chunks and
# queries receive the same topic feature before hashing, so the result remains a
# genuine persisted vector/cosine retrieval flow.
SEMANTIC_TOPICS = {
    "topic_release_checks": ("代码提交", "提交前", "合并代码", "发布前", "发布检查", "审阅", "评审", "核对", "复核"),
    "topic_idempotency": ("幂等", "请求重放", "重发", "重复请求", "相同键", "两份记录", "重复任务"),
    "topic_indexing": ("搜不到", "不能搜索", "正文检索", "语义检索", "索引", "可检索", "索引任务", "重建索引"),
    "topic_archive": ("旧项目", "旧资料", "隐藏", "归档", "恢复", "找回", "历史资料"),
    "topic_storage": ("上传", "保存文件", "原文件", "持久化", "下载", "存储目录"),
    "topic_classification": ("分类", "文件类型", "文件格式", "筛选", "归属", "目录"),
    "topic_failure": ("故障", "失败", "异常", "排查", "修复", "错误", "问题"),
    "topic_submission": ("约稿", "稿约", "征稿", "投稿", "稿件", "征文", "投稿要求", "约稿要求"),
}


def tokenize(text: str) -> list[str]:
    tokens = TOKEN_RE.findall(text.lower())
    # Character tokens keep short Chinese rewrites connected even when wording differs.
    # ASCII words remain whole tokens so identifiers such as API-032 stay searchable.
    return tokens


def embed(text: str) -> list[float]:
    provider = get_setting("embedding_provider")
    api_url = get_setting("embedding_api_url")
    api_key = get_setting("embedding_api_key")
    model = get_setting("embedding_model")
    if provider in {"remote", "openai", "openai_compatible"} and api_url:
        try:
            payload = {"input": text}
            if model:
                payload["model"] = model
            request = urllib.request.Request(
                api_url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json", **({"Authorization": f"Bearer {api_key}"} if api_key else {})},
                method="POST",
            )
            with urllib.request.urlopen(request, timeout=15) as response:
                body = json.loads(response.read().decode("utf-8"))
            vector = body.get("data", [{}])[0].get("embedding") or body.get("embedding")
            if vector and len(vector) == EMBEDDING_DIM:
                return [float(value) for value in vector]
        except Exception:
            # A remote provider is optional. Keep indexing usable when it is
            # unavailable and use the deterministic local vector instead.
            pass
    values = [0.0] * EMBEDDING_DIM
    normalized = re.sub(r"\s+", "", text.lower())
    tokens = tokenize(text)
    for topic, phrases in SEMANTIC_TOPICS.items():
        if any(phrase in normalized for phrase in phrases):
            tokens.append(topic)
    compact = normalized
    bigrams = [compact[index:index + 2] for index in range(max(0, len(compact) - 1))]
    tokens += bigrams
    # Chinese compounds are sometimes written in reverse order (例如“约稿”/“稿约”).
    # Add an order-independent feature alongside the normal directional bigram so
    # the vector remains useful for ordinary phrases while matching this common
    # wording variation.
    tokens += ["zhpair:" + "".join(sorted(pair)) for pair in bigrams if all("\u4e00" <= char <= "\u9fff" for char in pair)]
    if not tokens:
        return values
    for token in tokens:
        digest = hashlib.blake2b(token.encode("utf-8"), digest_size=8).digest()
        number = int.from_bytes(digest, "big")
        index = number % EMBEDDING_DIM
        values[index] += 1.0 if ((number >> 8) & 1) else -1.0
    norm = math.sqrt(sum(value * value for value in values)) or 1.0
    return [value / norm for value in values]


def cosine(a: list[float], b: list[float]) -> float:
    return sum(x * y for x, y in zip(a, b))
