"""WebSocket endpoint for batch task progress streaming."""

from collections import defaultdict
from typing import Any

from fastapi import APIRouter, WebSocket, WebSocketDisconnect
from jose import JWTError, jwt

from app.config import settings
from app.database import SessionLocal
from app.dependencies import ALGORITHM
from app.models.user import User

router = APIRouter()

# Active WebSocket connections per task_id
_connections: dict[str, list[WebSocket]] = defaultdict(list)


def _is_authorized_token(token: str | None) -> bool:
    if not token:
        return False
    try:
        payload = jwt.decode(token, settings.secret_key, algorithms=[ALGORITHM])
        user_id = int(payload.get("sub", 0))
    except (JWTError, ValueError):
        return False

    db = SessionLocal()
    try:
        return (
            db.query(User)
            .filter(User.id == user_id, User.is_active.is_(True))
            .first()
            is not None
        )
    finally:
        db.close()


async def broadcast_progress(task_id: str, message: dict[str, Any]) -> None:
    """Broadcast a progress message to all connected clients for a task."""
    import json

    dead = []
    for ws in _connections.get(task_id, []):
        try:
            await ws.send_text(json.dumps(message, ensure_ascii=False))
        except Exception:
            dead.append(ws)
    # Cleanup dead connections
    for ws in dead:
        _connections[task_id].remove(ws)


@router.websocket("/ws/tasks/{task_id}/progress")
async def task_progress_ws(websocket: WebSocket, task_id: str):
    if not _is_authorized_token(websocket.query_params.get("token")):
        await websocket.close(code=1008)
        return

    await websocket.accept()
    _connections[task_id].append(websocket)
    try:
        while True:
            # Keep connection alive; client may send ping/cancel
            data = await websocket.receive_text()
            if data == "cancel":
                # TODO: implement task cancellation
                pass
    except WebSocketDisconnect:
        pass
    finally:
        if websocket in _connections[task_id]:
            _connections[task_id].remove(websocket)
        if not _connections[task_id]:
            del _connections[task_id]
