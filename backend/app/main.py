"""FastAPI application factory."""

import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.api.router import api_router
from app.config import settings
from app.dependencies import pwd_context
from app.models import User  # noqa: F401 — ensure models are registered
from app.ws.progress import router as ws_router

_log = logging.getLogger("reportgen-web")

_INSECURE_SECRET_KEYS = {"", "change-me-in-production"}
_INSECURE_ADMIN_PASSWORDS = {"", "admin123"}


def _validate_security_settings() -> None:
    """Fail closed unless explicitly allowed for local development."""
    problems = []
    if settings.secret_key in _INSECURE_SECRET_KEYS:
        problems.append("RG_WEB_SECRET_KEY 未设置或仍为默认值")
    if settings.default_admin_password in _INSECURE_ADMIN_PASSWORDS:
        problems.append("RG_WEB_DEFAULT_ADMIN_PASSWORD 未设置或仍为默认值")

    if not problems:
        return
    message = "；".join(problems)
    if settings.allow_insecure_defaults:
        _log.warning("Insecure development settings enabled: %s", message)
        return
    raise RuntimeError(f"{message}。如仅本地开发，请显式设置 RG_WEB_ALLOW_INSECURE_DEFAULTS=true。")


def _run_database_migrations() -> None:
    """Upgrade the configured database to the latest Alembic revision."""
    from alembic.config import Config

    from alembic import command

    backend_dir = Path(__file__).resolve().parents[1]
    alembic_cfg = Config(str(backend_dir / "alembic.ini"))
    alembic_cfg.set_main_option("script_location", str(backend_dir / "alembic"))
    alembic_cfg.set_main_option("sqlalchemy.url", settings.database_url.replace("%", "%%"))
    command.upgrade(alembic_cfg, "head")


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application startup/shutdown events."""
    # Ensure storage directories exist
    for d in [
        settings.upload_dir,
        settings.report_dir,
        settings.preview_dir,
        settings.signature_dir,
        settings.storage_root / "db",
    ]:
        d.mkdir(parents=True, exist_ok=True)

    _validate_security_settings()

    # Schema evolution is managed by Alembic migrations, not implicit create_all.
    _run_database_migrations()

    # Seed default admin if no users exist
    from app.database import SessionLocal

    db = SessionLocal()
    try:
        if db.query(User).count() == 0:
            admin = User(
                username=settings.default_admin_username,
                password_hash=pwd_context.hash(settings.default_admin_password),
                display_name="管理员",
                role="admin",
            )
            db.add(admin)
            db.commit()
        else:
            admin = (
                db.query(User)
                .filter(User.username == settings.default_admin_username)
                .first()
            )
            if admin and pwd_context.verify("admin123", admin.password_hash):
                admin.password_hash = pwd_context.hash(settings.default_admin_password)
                db.commit()
                _log.warning("Rotated legacy default admin password hash")
    finally:
        db.close()

    yield


def create_app() -> FastAPI:
    app = FastAPI(
        title="基因组Panel自动化报告系统",
        description="Genomic panel report automation web platform",
        version="0.1.0",
        lifespan=lifespan,
    )

    # Same-origin frontend/API traffic does not need CORS. Configure explicit
    # trusted origins only when the frontend is hosted on a different origin.
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origin_list,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.get("/healthz", include_in_schema=False)
    def healthz():
        return {"status": "ok"}

    # API routes
    app.include_router(api_router)

    # WebSocket routes
    app.include_router(ws_router)

    # Serve frontend static files if built
    static_dir = Path(__file__).parent.parent / "static"
    if static_dir.exists():
        app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")

    return app


app = create_app()
