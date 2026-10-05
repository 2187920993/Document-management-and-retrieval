import io
import os
import zipfile

os.environ["DATABASE_URL"] = "sqlite:///./data/test.db"
os.environ["STORAGE_DIR"] = "./storage-test"

from fastapi.testclient import TestClient
from app.main import app


def test_health_and_category(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    with TestClient(app) as client:
        response = client.get("/api/health")
        assert response.status_code == 200
        assert response.json()["status"] == "ok"
        response = client.post("/api/categories", json={"name": "技术规范"})
        assert response.status_code == 201


def test_upload_index_search_download_and_archive(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    with TestClient(app) as client:
        category = client.post("/api/categories", json={"name": "项目文件"}).json()
        content = "发布前需要检查数据库备份、索引任务和回退方案。"
        uploaded = client.post(
            "/api/documents",
            files={"file": ("发布检查.md", content.encode("utf-8"), "text/markdown")},
            data={"category_id": str(category["id"])},
        )
        assert uploaded.status_code == 201
        document = uploaded.json()
        assert document["storage_status"] == "stored"
        assert document["index_status"] == "pending"

        indexed = client.post(f"/api/documents/{document['id']}/index-now")
        assert indexed.status_code == 200
        assert indexed.json()["index_status"] == "ready"
        assert client.get(f"/api/documents/{document['id']}/download").content == content.encode("utf-8")
        keyword = client.get("/api/search", params={"q": "数据库备份", "mode": "keyword"})
        assert keyword.status_code == 200 and keyword.json()["results"][0]["document"]["id"] == document["id"]
        semantic = client.get("/api/search", params={"q": "发布之前要做哪些准备", "mode": "semantic"})
        assert semantic.status_code == 200 and semantic.json()["results"]

        archived = client.patch(f"/api/documents/{document['id']}", json={"archived": True})
        assert archived.status_code == 200
        assert client.get("/api/documents").json() == []
        assert client.get("/api/documents", params={"archived": True}).json()[0]["id"] == document["id"]


def test_batch_upload_broad_formats_keyword_search_and_zip(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    with TestClient(app) as client:
        files = [
            ("files", ("notes.log", b"worker queue retry marker", "text/plain")),
            ("files", ("config.json", b'{"queue": "worker"}', "application/json")),
        ]
        response = client.post("/api/documents/batch", files=files)
        assert response.status_code == 200
        payload = response.json()
        assert len(payload["uploaded"]) == 2 and payload["failed"] == []
        for document in payload["uploaded"]:
            result = client.post(f"/api/documents/{document['id']}/index-now")
            assert result.json()["index_status"] == "ready"

        search = client.get("/api/search", params={"q": "worker", "mode": "keyword"})
        assert {item["document"]["extension"] for item in search.json()["results"]} == {".log", ".json"}

        ids = [document["id"] for document in payload["uploaded"]]
        archive = client.get("/api/documents/download-batch", params=[("ids", ids[0]), ("ids", ids[1])])
        assert archive.status_code == 200
        with zipfile.ZipFile(io.BytesIO(archive.content)) as bundle:
            assert set(bundle.namelist()) == {"notes.log", "config.json"}
