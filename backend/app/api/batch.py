"""Batch report generation endpoints."""

import asyncio
import json
import uuid
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException
from sqlalchemy.orm import Session

from app.config import settings
from app.database import SessionLocal, get_db
from app.models.task import Task, TaskResult
from app.models.upload import Upload
from app.schemas.clinical_info import PatientInfo
from app.schemas.common import ApiResponse
from app.services import clinical_info_service
from app.services.file_manager import ensure_report_dir
from app.services.task_manager import submit_batch_task

router = APIRouter(prefix="/reports", tags=["reports-batch"])


def _persist_patient_snapshot(snapshot: dict | None) -> None:
    if not snapshot:
        return
    sample_id = str(snapshot.get("sample_id") or "").strip()
    if not sample_id:
        return
    patient = PatientInfo(
        sample_id=sample_id,
        patient_name=snapshot.get("patient_name"),
        gender=snapshot.get("gender"),
        age=str(snapshot.get("age")) if snapshot.get("age") not in (None, "") else None,
        cancer_type=snapshot.get("cancer_type"),
        sample_type=snapshot.get("sample_type"),
        report_number=snapshot.get("report_number"),
        pathology_id=snapshot.get("pathology_id"),
        hospital=snapshot.get("hospital"),
        department=snapshot.get("department"),
        collection_date=snapshot.get("collection_date"),
        receive_date=snapshot.get("receive_date"),
        report_date=snapshot.get("report_date"),
        issuer=snapshot.get("issuer"),
        reviewer=snapshot.get("reviewer"),
        signature_image_path=snapshot.get("signature_image_path"),
        project_type=snapshot.get("project_type"),
        project_name=snapshot.get("project_name"),
    )
    clinical_info_service.upsert_patient(patient)


async def _on_batch_complete(task_id: str, result: dict):
    """Callback when batch task completes — update DB."""
    db = SessionLocal()
    try:
        task = db.query(Task).filter(Task.id == task_id).first()
        if not task:
            return

        if result.get("success"):
            report = result.get("report", {})
            task.completed_files = report.get("successes", 0)
            task.failed_files = report.get("failures", 0)
            task.status = "failed" if task.failed_files else "completed"
            task.output_path = result.get("output_root")
            task.errors = json.dumps(
                [
                    f"{item.get('excel_filename')}: {err}"
                    for item in report.get("results", [])
                    for err in (item.get("errors") or [])
                ],
                ensure_ascii=False,
            )
            task.warnings = json.dumps(
                [
                    f"{item.get('excel_filename')}: {warn}"
                    for item in report.get("results", [])
                    for warn in (item.get("warnings") or [])
                ],
                ensure_ascii=False,
            )
            db.query(TaskResult).filter(TaskResult.task_id == task_id).delete()
            for item in report.get("results", []):
                status = "completed" if item.get("ok") else "failed"
                db.add(
                    TaskResult(
                        task_id=task_id,
                        file_index=int(item.get("index") or 0),
                        excel_filename=str(item.get("excel_filename") or ""),
                        status=status,
                        output_path=item.get("output_docx"),
                        duration_seconds=item.get("duration_seconds"),
                        errors=json.dumps(item.get("errors") or [], ensure_ascii=False),
                        warnings=json.dumps(
                            item.get("warnings") or [], ensure_ascii=False
                        ),
                        validation_summary=json.dumps(
                            item.get("validation") or {}, ensure_ascii=False
                        ),
                    )
                )
                if status == "completed":
                    try:
                        _persist_patient_snapshot(item.get("patient_snapshot"))
                    except Exception:
                        pass
        else:
            task.status = "failed"
            task.errors = json.dumps([result.get("error", "Unknown error")], ensure_ascii=False)

        task.completed_at = datetime.utcnow()
        if task.started_at:
            task.duration_seconds = (task.completed_at - task.started_at).total_seconds()
        db.commit()
    finally:
        db.close()


@router.post("/batch", response_model=ApiResponse)
async def batch_generate(
    upload_ids: list[str] = [],
    input_dir: Optional[str] = None,
    project_type: Optional[str] = None,
    highlight: bool = False,
    template_contract: str = "warn",
    background_tasks: BackgroundTasks = BackgroundTasks(),
    db: Session = Depends(get_db),
):
    """
    Submit a batch generation task.

    Provide either upload_ids (list of previously uploaded files)
    or input_dir (directory path containing Excel files).
    """
    task_id = str(uuid.uuid4())
    output_dir = ensure_report_dir(task_id)

    # Resolve input paths
    input_paths: list[str] = []
    if upload_ids:
        for uid in upload_ids:
            upload = db.query(Upload).filter(Upload.id == uid).first()
            if upload:
                input_paths.append(upload.stored_path)
    elif input_dir:
        input_paths.append(input_dir)
    else:
        raise HTTPException(status_code=400, detail="请提供 upload_ids 或 input_dir")

    total_files = len(input_paths) if upload_ids else 0  # unknown for dir

    # Create task record
    task = Task(
        id=task_id,
        task_type="batch",
        status="running",
        project_type=project_type,
        total_files=total_files,
        started_at=datetime.utcnow(),
    )
    db.add(task)
    db.commit()

    # Submit to background
    asyncio.create_task(
        submit_batch_task(
            task_id=task_id,
            inputs=input_paths,
            output_root=str(output_dir),
            config_dir=settings.upstream_config_dir,
            template=None,
            project_type=project_type,
            template_contract=template_contract,
            highlight=highlight,
            on_complete=_on_batch_complete,
        )
    )

    return ApiResponse(data={
        "task_id": task_id,
        "status": "running",
        "total_files": total_files,
    })
