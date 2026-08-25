from __future__ import annotations

from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.services.document_processor import (
    ProcessingError,
    create_job_from_upload,
    cleanup_expired_jobs,
    get_job,
    job_file,
    process_job,
)


APP_ROOT = Path(__file__).resolve().parent
app = FastAPI(title="Revisor ABNT", version="0.2.0")
app.mount("/static", StaticFiles(directory=APP_ROOT / "static"), name="static")


def public_job(job: dict) -> dict:
    job_id = job["id"]
    return {
        "id": job_id,
        "status": job["status"],
        "error": job.get("error"),
        "original_filename": job["original_filename"],
        "status_url": f"/api/revisions/{job_id}",
        "download_url": f"/api/revisions/{job_id}/download" if job["status"] == "completed" else None,
        "audit_url": f"/api/revisions/{job_id}/audit" if job["status"] == "completed" else None,
        "format_report_url": f"/api/revisions/{job_id}/format-report" if job["status"] == "completed" else None,
        "review_summary": job.get("review_summary"),
    }


@app.get("/", include_in_schema=False)
def home() -> FileResponse:
    return FileResponse(APP_ROOT / "static" / "index.html")


@app.get("/privacidade", include_in_schema=False)
def privacy() -> FileResponse:
    return FileResponse(APP_ROOT / "static" / "privacy.html")


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "expired_jobs_removed": cleanup_expired_jobs()}


@app.post("/api/revisions", status_code=status.HTTP_202_ACCEPTED)
async def create_revision(
    background_tasks: BackgroundTasks,
    document: UploadFile = File(description="Arquivo DOCX ou DOC de até 25 MB"),
    document_type: str = Form("tcc"),
    citation_system: str = Form("auto"),
    font: str = Form("Times New Roman"),
    toc_mode: str = Form("audit"),
    pagination_mode: str = Form("audit"),
    order_references: bool = Form(True),
    privacy_acknowledged: bool = Form(...),
) -> dict:
    if not privacy_acknowledged:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Confirme a leitura do tratamento de documentos para continuar.",
        )
    if document_type not in {"tcc", "article"}:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail="Tipo de documento inválido.")
    if citation_system not in {"auto", "author-date", "numeric"}:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail="Sistema de citação inválido.")
    if font not in {"Times New Roman", "Arial"}:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail="Fonte de revisão inválida.")
    if toc_mode not in {"audit", "insert-if-empty"} or pagination_mode not in {"audit", "request"}:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail="Opção de estrutura inválida.")
    try:
        job = await create_job_from_upload(document)
    except ProcessingError as exc:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(exc)) from exc
    background_tasks.add_task(
        process_job,
        job["id"],
        document_type=document_type,
        citation_system=citation_system,
        font=font,
        toc_mode=toc_mode,
        pagination_mode=pagination_mode,
        order_references=order_references,
    )
    return public_job(job)


@app.get("/api/revisions/{job_id}")
def revision_status(job_id: str) -> dict:
    try:
        return public_job(get_job(job_id))
    except ProcessingError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc


@app.get("/api/revisions/{job_id}/download")
def download_revision(job_id: str) -> FileResponse:
    try:
        path = job_file(job_id, "trabalho-revisado-abnt.docx")
    except ProcessingError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="trabalho-revisado-abnt.docx")


@app.get("/api/revisions/{job_id}/audit")
def download_audit(job_id: str) -> FileResponse:
    try:
        path = job_file(job_id, "relatorio-auditoria.json")
    except ProcessingError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    return FileResponse(path, media_type="application/json", filename="relatorio-auditoria.json")


@app.get("/api/revisions/{job_id}/format-report")
def download_format_report(job_id: str) -> FileResponse:
    try:
        path = job_file(job_id, "trabalho-revisado-abnt_abnt_report.json")
    except ProcessingError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    return FileResponse(path, media_type="application/json", filename="relatorio-revisao-abnt.json")
