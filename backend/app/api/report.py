"""Report generation and download endpoints."""

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.dependencies import get_bridge
from app.models.task import Task
from app.models.upload import Upload
from app.schemas.clinical_info import PatientInfo
from app.schemas.common import ApiResponse
from app.schemas.report import GenerateRequest, GenerateResponse, TaskStatus
from app.services import clinical_info_service as clinical_info_service
from app.services.file_manager import ensure_report_dir
from app.services.reportgen_bridge import ReportGenBridge

router = APIRouter(prefix="/reports", tags=["reports"])


def _persist_patient_snapshot(clinical_info: dict | None) -> None:
    """将本次报告中的患者基础信息回写到 patient_info.yaml。"""
    if not clinical_info:
        return

    sample_id = str(clinical_info.get("sample_id") or "").strip()
    if not sample_id:
        return

    patient = PatientInfo(
        sample_id=sample_id,
        patient_name=clinical_info.get("patient_name"),
        gender=clinical_info.get("gender"),
        age=str(clinical_info.get("age")) if clinical_info.get("age") not in (None, "") else None,
        cancer_type=clinical_info.get("cancer_type"),
        sample_type=clinical_info.get("sample_type"),
        report_number=clinical_info.get("report_number"),
        pathology_id=clinical_info.get("pathology_id"),
        hospital=clinical_info.get("hospital"),
        department=clinical_info.get("department"),
        collection_date=clinical_info.get("collection_date"),
        receive_date=clinical_info.get("receive_date"),
        report_date=clinical_info.get("report_date"),
        issuer=clinical_info.get("issuer"),
        reviewer=clinical_info.get("reviewer"),
        signature_image_path=clinical_info.get("signature_image_path"),
        project_type=clinical_info.get("project_type"),
        project_name=clinical_info.get("project_name"),
    )
    clinical_info_service.upsert_patient(patient)


def _project_name_from_type(
    bridge: ReportGenBridge, project_type: Optional[str]
) -> Optional[str]:
    if not project_type:
        return None
    for entry in bridge.detector.project_types:
        if entry.get("id") == project_type:
            return entry.get("name")
    return None


@router.post("/generate", response_model=ApiResponse[GenerateResponse])
def generate_report(
    req: GenerateRequest,
    db: Session = Depends(get_db),
    bridge: ReportGenBridge = Depends(get_bridge),
):
    """Generate a single report (synchronous, 2-5s)."""
    upload = db.query(Upload).filter(Upload.id == req.upload_id).first()
    if not upload:
        raise HTTPException(status_code=404, detail="上传记录不存在")

    task_id = str(uuid.uuid4())
    output_dir = ensure_report_dir(task_id)
    resolved_project_type = req.project_type or upload.detected_project_type
    resolved_project_name = (
        req.project_name
        or upload.detected_project_name
        or _project_name_from_type(bridge, resolved_project_type)
    )
    clinical_snapshot = dict(req.clinical_info or {})
    if resolved_project_type:
        clinical_snapshot["project_type"] = resolved_project_type
    if resolved_project_name:
        clinical_snapshot["project_name"] = resolved_project_name

    # Create task record
    task = Task(
        id=task_id,
        upload_id=req.upload_id,
        task_type="single",
        status="running",
        project_type=resolved_project_type,
        clinical_info_snapshot=json.dumps(clinical_snapshot, ensure_ascii=False)
        if clinical_snapshot
        else None,
        started_at=datetime.utcnow(),
    )
    db.add(task)
    db.commit()

    try:
        result = bridge.generate_report(
            excel_path=upload.stored_path,
            output_dir=str(output_dir),
            template_name=req.template_name,
            clinical_info=req.clinical_info,
            project_type=resolved_project_type,
            project_name=resolved_project_name,
            strict_mode=req.strict_mode,
            template_contract_mode=req.template_contract_mode,
        )

        success = result.get("success", False)
        task.status = "completed" if success else "failed"
        task.output_path = result.get("output_file")
        task.duration_seconds = result.get("duration")
        task.errors = json.dumps(result.get("errors", []), ensure_ascii=False)
        task.warnings = json.dumps(result.get("warnings", []), ensure_ascii=False)
        task.completed_at = datetime.utcnow()
        db.commit()

        if success:
            try:
                _persist_patient_snapshot(clinical_snapshot)
            except Exception:
                # 患者信息回写失败不影响报告生成结果
                pass

        return ApiResponse(
            data=GenerateResponse(
                task_id=task_id,
                success=success,
                output_file=result.get("output_file"),
                duration_seconds=result.get("duration"),
                errors=result.get("errors", []),
                warnings=result.get("warnings", []),
            )
        )
    except Exception as e:
        task.status = "failed"
        task.errors = json.dumps([str(e)], ensure_ascii=False)
        task.completed_at = datetime.utcnow()
        db.commit()
        return ApiResponse(
            success=False,
            data=GenerateResponse(
                task_id=task_id,
                success=False,
                errors=[str(e)],
            ),
            error=str(e),
        )


@router.get("/{task_id}", response_model=ApiResponse[TaskStatus])
def get_task_status(task_id: str, db: Session = Depends(get_db)):
    task = db.query(Task).filter(Task.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")

    return ApiResponse(
        data=TaskStatus(
            id=task.id,
            task_type=task.task_type,
            status=task.status,
            project_type=task.project_type,
            total_files=task.total_files,
            completed_files=task.completed_files,
            failed_files=task.failed_files,
            output_path=task.output_path,
            created_at=task.created_at,
            started_at=task.started_at,
            completed_at=task.completed_at,
            duration_seconds=task.duration_seconds,
            errors=json.loads(task.errors) if task.errors else [],
            warnings=json.loads(task.warnings) if task.warnings else [],
        )
    )


@router.get("/{task_id}/download")
def download_report(task_id: str, db: Session = Depends(get_db)):
    task = db.query(Task).filter(Task.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")
    if not task.output_path:
        raise HTTPException(status_code=404, detail="报告文件不存在")

    file_path = Path(task.output_path)
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="报告文件已被删除")

    return FileResponse(
        path=str(file_path),
        filename=file_path.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
