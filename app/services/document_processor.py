"""Armazenamento isolado e processamento conservador de revisões DOCX."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import uuid
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[2]
SKILL_DIR = PROJECT_ROOT / "revisor-abnt-docx"
AUDIT_SCRIPT = SKILL_DIR / "scripts" / "audit_docx_abnt.py"
FORMAT_SCRIPT = SKILL_DIR / "scripts" / "format_docx_abnt.py"
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
ALLOWED_EXTENSIONS = {".doc", ".docx"}


class ProcessingError(RuntimeError):
    """Erro seguro para apresentar ao usuário da aplicação."""


def _jobs_root() -> Path:
    root = Path(os.getenv("REVISOR_ABNT_JOBS_DIR", PROJECT_ROOT / "data" / "jobs"))
    root.mkdir(parents=True, exist_ok=True)
    return root


def _retention_hours() -> int:
    """Return a bounded retention period so uploaded files cannot accumulate forever."""
    try:
        value = int(os.getenv("REVISOR_ABNT_RETENTION_HOURS", "24"))
    except ValueError:
        value = 24
    return min(max(value, 1), 168)


# Um job saudavel processa em segundos a poucos minutos (o subprocesso de
# formatacao tem timeout de 180s, a conversao de .doc tem 120s). Se um job
# ainda estiver em "uploaded"/"processing" muito depois disso, o processo que
# deveria terminá-lo provavelmente morreu (crash, reinício do container) e o
# job ficaria preso para sempre sem isto -- BackgroundTasks nao tem retomada
# nem supervisor proprio.
STALE_PROCESSING_MINUTES = 30


def cleanup_expired_jobs() -> int:
    """Remove jobs finalizados apos a retencao e jobs travados ha muito tempo em processamento."""
    now = datetime.now(UTC).timestamp()
    threshold = now - (_retention_hours() * 3600)
    stale_threshold = now - (STALE_PROCESSING_MINUTES * 60)
    removed = 0
    for candidate in _jobs_root().iterdir():
        if not candidate.is_dir():
            continue
        try:
            uuid.UUID(candidate.name)
            job_path = candidate / "job.json"
            job = json.loads(job_path.read_text(encoding="utf-8"))
            created_at = datetime.fromisoformat(job["created_at"]).timestamp()
        except (OSError, ValueError, KeyError, json.JSONDecodeError):
            continue
        status = job.get("status")
        if status in {"completed", "failed"} and created_at < threshold:
            shutil.rmtree(candidate, ignore_errors=True)
            removed += 1
        elif status in {"uploaded", "processing"} and created_at < stale_threshold:
            shutil.rmtree(candidate, ignore_errors=True)
            removed += 1
    return removed


def _job_dir(job_id: str) -> Path:
    try:
        normalized = str(uuid.UUID(job_id))
    except ValueError as exc:
        raise ProcessingError("Identificador de revisão inválido.") from exc
    return _jobs_root() / normalized


def _job_path(job_id: str) -> Path:
    return _job_dir(job_id) / "job.json"


def _now() -> str:
    return datetime.now(UTC).isoformat()


def _write_job(job_id: str, payload: dict[str, Any]) -> None:
    path = _job_path(job_id)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def get_job(job_id: str) -> dict[str, Any]:
    path = _job_path(job_id)
    if not path.is_file():
        raise ProcessingError("Revisão não encontrada.")
    return json.loads(path.read_text(encoding="utf-8"))


def update_job(job_id: str, **changes: Any) -> dict[str, Any]:
    job = get_job(job_id)
    job.update(changes)
    job["updated_at"] = _now()
    _write_job(job_id, job)
    return job


async def create_job_from_upload(upload) -> dict[str, Any]:
    filename = upload.filename or "documento"
    extension = Path(filename).suffix.lower()
    if extension not in ALLOWED_EXTENSIONS:
        raise ProcessingError("Envie um arquivo .docx ou .doc.")

    cleanup_expired_jobs()
    job_id = str(uuid.uuid4())
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=False)
    original_path = job_dir / f"original{extension}"
    total = 0
    try:
        with original_path.open("wb") as destination:
            while chunk := await upload.read(1024 * 1024):
                total += len(chunk)
                if total > MAX_UPLOAD_BYTES:
                    raise ProcessingError("O arquivo excede o limite de 25 MB.")
                destination.write(chunk)
    except Exception:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    finally:
        await upload.close()

    job = {
        "id": job_id,
        "status": "uploaded",
        "original_filename": filename,
        "original_extension": extension,
        "size_bytes": total,
        "created_at": _now(),
        "updated_at": _now(),
        "error": None,
        "output_filename": None,
    }
    _write_job(job_id, job)
    return job


def _is_valid_docx(path: Path) -> bool:
    if not zipfile.is_zipfile(path):
        return False
    with zipfile.ZipFile(path) as archive:
        return "[Content_Types].xml" in archive.namelist() and "word/document.xml" in archive.namelist()


def _convert_doc_to_docx(source: Path, work_dir: Path) -> Path:
    configured = os.getenv("SOFFICE_BIN", "soffice")
    soffice = shutil.which(configured) if not Path(configured).is_file() else configured
    if not soffice:
        raise ProcessingError(
            "Arquivos .doc exigem LibreOffice no servidor. Instale o conversor ou envie o arquivo em .docx."
        )
    completed = subprocess.run(
        [str(soffice), "--headless", "--convert-to", "docx", "--outdir", str(work_dir), str(source)],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    converted = work_dir / f"{source.stem}.docx"
    if completed.returncode != 0 or not converted.is_file():
        raise ProcessingError("Não foi possível converter o arquivo .doc para .docx com segurança.")
    return converted


def _run(command: list[str], label: str) -> None:
    completed = subprocess.run(command, capture_output=True, text=True, timeout=180, check=False)
    if completed.returncode != 0:
        details = (completed.stderr or completed.stdout).strip().splitlines()
        raise ProcessingError(f"Falha durante {label}: {details[-1] if details else 'erro interno'}")


def process_job(
    job_id: str,
    *,
    document_type: str = "tcc",
    citation_system: str = "auto",
    font: str = "Times New Roman",
    toc_mode: str = "audit",
    pagination_mode: str = "audit",
    order_references: bool = True,
    institution: str = "generic",
) -> None:
    """Executa a revisão em diretório próprio; nunca modifica o arquivo original."""
    try:
        job = update_job(job_id, status="processing", error=None)
        job_dir = _job_dir(job_id)
        source = job_dir / f"original{job['original_extension']}"
        input_docx = _convert_doc_to_docx(source, job_dir) if source.suffix.lower() == ".doc" else source
        if not _is_valid_docx(input_docx):
            raise ProcessingError("O arquivo enviado não possui uma estrutura DOCX válida.")

        audit_path = job_dir / "relatorio-auditoria.json"
        output_path = job_dir / "trabalho-revisado-abnt.docx"
        _run(
            [
                sys.executable, str(AUDIT_SCRIPT), str(input_docx), "--out", str(audit_path),
                "--document-type", document_type, "--citation-system", citation_system,
                "--institution", institution,
            ],
            "a auditoria",
        )
        format_command = [
            sys.executable, str(FORMAT_SCRIPT), str(input_docx), "--out", str(output_path),
            "--document-type", document_type, "--citation-system", citation_system,
            "--font", font, "--toc-mode", toc_mode, "--pagination-mode", pagination_mode,
            "--institution", institution,
        ]
        if not order_references:
            format_command.append("--do-not-order-references")
        _run(format_command, "a formatação")
        format_report = output_path.with_name(f"{output_path.stem}_abnt_report.json")
        review_summary = None
        if format_report.is_file():
            report_data = json.loads(format_report.read_text(encoding="utf-8"))
            review_summary = {
                "actions_applied": len(report_data.get("actions_applied", [])),
                "issues_remaining": len(report_data.get("issues_remaining", [])),
            }
        update_job(
            job_id,
            status="completed",
            output_filename="trabalho-revisado-abnt.docx",
            audit_filename=audit_path.name,
            format_report_filename=format_report.name if format_report.is_file() else None,
            review_options={
                "document_type": document_type,
                "citation_system": citation_system,
                "font": font,
                "toc_mode": toc_mode,
                "pagination_mode": pagination_mode,
                "order_references": order_references,
                "institution": institution,
            },
            review_summary=review_summary,
        )
    except ProcessingError as exc:
        update_job(job_id, status="failed", error=str(exc))
    except Exception:
        update_job(job_id, status="failed", error="Ocorreu um erro interno durante a revisão.")


def job_file(job_id: str, filename: str) -> Path:
    allowed = {"trabalho-revisado-abnt.docx", "relatorio-auditoria.json", "trabalho-revisado-abnt_abnt_report.json"}
    if filename not in allowed:
        raise ProcessingError("Arquivo de resultado inválido.")
    path = _job_dir(job_id) / filename
    if not path.is_file():
        raise ProcessingError("Resultado ainda não está disponível.")
    return path
