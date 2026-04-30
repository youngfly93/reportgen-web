"""Task queue management endpoints."""

import json
import shutil
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from app.config import settings
from app.database import get_db
from app.models.task import Task, TaskResult
from app.models.upload import Upload
from app.schemas.common import ApiResponse

router = APIRouter(prefix="/tasks", tags=["tasks"])


def _loads_json(value: str | None) -> dict | list:
    if not value:
        return {}
    try:
        return json.loads(value)
    except Exception:
        return {}


def _derive_sample_id(snapshot: dict, upload: Upload | None, task: Task) -> str | None:
    sample_id = str(snapshot.get("sample_id") or "").strip()
    if sample_id:
        return sample_id

    candidate_paths = [task.output_path]
    if upload:
        candidate_paths.append(upload.original_filename)

    for candidate in candidate_paths:
        if not candidate:
            continue
        stem = Path(candidate).stem
        parts = [part for part in stem.split("_") if part]
        for part in parts:
            if any(ch.isdigit() for ch in part) and len(part) >= 8:
                return part
    return None


def _serialize_task(task: Task, upload: Upload | None = None) -> dict:
    clinical_info_snapshot = _loads_json(task.clinical_info_snapshot)
    warnings = _loads_json(task.warnings)
    errors = _loads_json(task.errors)
    sample_id = _derive_sample_id(clinical_info_snapshot, upload, task)

    return {
        "id": task.id,
        "task_type": task.task_type,
        "status": task.status,
        "project_type": task.project_type,
        "project_name": clinical_info_snapshot.get("project_name")
        or (upload.detected_project_name if upload else None),
        "sample_id": sample_id,
        "sample_label": sample_id or task.id,
        "upload_id": task.upload_id,
        "original_filename": upload.original_filename if upload else None,
        "clinical_info_snapshot": clinical_info_snapshot,
        "total_files": task.total_files,
        "completed_files": task.completed_files,
        "failed_files": task.failed_files,
        "output_path": task.output_path,
        "created_at": task.created_at.isoformat() if task.created_at else None,
        "started_at": task.started_at.isoformat() if task.started_at else None,
        "completed_at": task.completed_at.isoformat() if task.completed_at else None,
        "duration_seconds": task.duration_seconds,
        "errors": errors if isinstance(errors, list) else [],
        "warnings": warnings if isinstance(warnings, list) else [],
        "editable": task.task_type == "single" and bool(task.upload_id),
    }


def _serialize_task_result(result: TaskResult) -> dict:
    return {
        "id": result.id,
        "task_id": result.task_id,
        "file_index": result.file_index,
        "excel_filename": result.excel_filename,
        "status": result.status,
        "output_path": result.output_path,
        "duration_seconds": result.duration_seconds,
        "errors": _loads_json(result.errors) if result.errors else [],
        "warnings": _loads_json(result.warnings) if result.warnings else [],
        "validation_summary": _loads_json(result.validation_summary),
    }


@router.get("", response_model=ApiResponse)
def list_tasks(
    status: str = Query(None, description="Filter by status"),
    task_type: str = Query(None, description="Filter by type: single|batch"),
    page: int = 1,
    page_size: int = 20,
    db: Session = Depends(get_db),
):
    query = db.query(Task).order_by(Task.created_at.desc())
    if status:
        query = query.filter(Task.status == status)
    if task_type:
        query = query.filter(Task.task_type == task_type)

    total = query.count()
    tasks = query.offset((page - 1) * page_size).limit(page_size).all()
    upload_ids = {t.upload_id for t in tasks if t.upload_id}
    uploads = {}
    if upload_ids:
        uploads = {
            u.id: u
            for u in db.query(Upload).filter(Upload.id.in_(upload_ids)).all()
        }

    items = []
    for t in tasks:
        items.append(_serialize_task(t, uploads.get(t.upload_id)))

    return ApiResponse(data={"items": items, "total": total, "page": page, "page_size": page_size})


@router.get("/stats", response_model=ApiResponse)
def task_stats(db: Session = Depends(get_db)):
    total = db.query(Task).count()
    completed = db.query(Task).filter(Task.status == "completed").count()
    failed = db.query(Task).filter(Task.status == "failed").count()
    running = db.query(Task).filter(Task.status == "running").count()
    pending = db.query(Task).filter(Task.status == "pending").count()

    return ApiResponse(data={
        "total": total,
        "completed": completed,
        "failed": failed,
        "running": running,
        "pending": pending,
    })


@router.get("/{task_id}", response_model=ApiResponse)
def get_task(task_id: str, db: Session = Depends(get_db)):
    task = db.query(Task).filter(Task.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")
    upload = None
    if task.upload_id:
        upload = db.query(Upload).filter(Upload.id == task.upload_id).first()
    return ApiResponse(data=_serialize_task(task, upload))


@router.get("/{task_id}/results", response_model=ApiResponse)
def get_task_results(task_id: str, db: Session = Depends(get_db)):
    task = db.query(Task).filter(Task.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")
    results = (
        db.query(TaskResult)
        .filter(TaskResult.task_id == task_id)
        .order_by(TaskResult.file_index.asc())
        .all()
    )
    return ApiResponse(data=[_serialize_task_result(result) for result in results])


@router.delete("/{task_id}", response_model=ApiResponse)
def cancel_task(task_id: str, db: Session = Depends(get_db)):
    task = db.query(Task).filter(Task.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")
    if task.status not in ("pending", "running"):
        raise HTTPException(status_code=400, detail="只能取消待执行或执行中的任务")

    task.status = "cancelled"
    db.commit()
    return ApiResponse(data={"id": task_id, "status": "cancelled"})


@router.delete("/{task_id}/record", response_model=ApiResponse)
def delete_task_record(task_id: str, db: Session = Depends(get_db)):
    task = db.query(Task).filter(Task.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")

    output_path = Path(task.output_path) if task.output_path else None
    db.delete(task)
    db.commit()

    if output_path and output_path.exists():
        try:
            output_path.unlink()
        except IsADirectoryError:
            pass
        except Exception:
            pass

    task_dir = settings.report_dir / task_id
    if task_dir.exists():
        shutil.rmtree(task_dir, ignore_errors=True)

    return ApiResponse(data={"id": task_id, "deleted": True})
