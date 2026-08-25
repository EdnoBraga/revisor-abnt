from io import BytesIO

from docx import Document
from fastapi.testclient import TestClient

from app.main import app


def sample_docx() -> bytes:
    document = Document()
    document.add_paragraph("CAPA")
    document.add_heading("1 INTRODUÇÃO", level=1)
    document.add_paragraph("Texto de teste para revisão.")
    document.add_heading("REFERÊNCIAS", level=1)
    document.add_paragraph("SILVA, Maria. Título de teste. Brasília: Editora, 2026.")
    content = BytesIO()
    document.save(content)
    return content.getvalue()


def test_health():
    client = TestClient(app)
    assert client.get("/api/health").json() == {"status": "ok"}


def test_docx_revision_creates_downloadable_copy(tmp_path, monkeypatch):
    monkeypatch.setenv("REVISOR_ABNT_JOBS_DIR", str(tmp_path / "jobs"))
    client = TestClient(app)
    response = client.post(
        "/api/revisions",
        files={"document": ("tcc.docx", sample_docx(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        data={"insert_page_numbers": "false"},
    )
    assert response.status_code == 202
    job = client.get(response.json()["status_url"]).json()
    assert job["status"] == "completed"
    download = client.get(job["download_url"])
    assert download.status_code == 200
    assert download.content[:2] == b"PK"
