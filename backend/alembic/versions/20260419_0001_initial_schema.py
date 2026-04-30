"""Initial web application schema.

Revision ID: 20260419_0001
Revises:
Create Date: 2026-04-19
"""

import sqlalchemy as sa

from alembic import op

revision = "20260419_0001"
down_revision = None
branch_labels = None
depends_on = None


def _has_table(name: str) -> bool:
    return sa.inspect(op.get_bind()).has_table(name)


def upgrade() -> None:
    if not _has_table("users"):
        op.create_table(
            "users",
            sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
            sa.Column("username", sa.String(length=50), nullable=False),
            sa.Column("password_hash", sa.String(length=128), nullable=False),
            sa.Column("display_name", sa.String(length=100), nullable=False),
            sa.Column("role", sa.String(length=20), nullable=False),
            sa.Column("is_active", sa.Boolean(), nullable=True),
            sa.Column("created_at", sa.DateTime(), server_default=sa.func.now()),
            sa.Column("last_login_at", sa.DateTime(), nullable=True),
            sa.UniqueConstraint("username"),
        )

    if not _has_table("uploads"):
        op.create_table(
            "uploads",
            sa.Column("id", sa.String(length=36), primary_key=True),
            sa.Column("user_id", sa.Integer(), nullable=True),
            sa.Column("original_filename", sa.String(length=255), nullable=False),
            sa.Column("stored_path", sa.String(length=500), nullable=False),
            sa.Column("file_size_bytes", sa.Integer(), nullable=True),
            sa.Column("sheet_names", sa.Text(), nullable=True),
            sa.Column("detected_project_type", sa.String(length=50), nullable=True),
            sa.Column("detected_project_name", sa.String(length=200), nullable=True),
            sa.Column("detection_confidence", sa.Float(), nullable=True),
            sa.Column("uploaded_at", sa.DateTime(), server_default=sa.func.now()),
            sa.Column("status", sa.String(length=20), nullable=True),
        )

    if not _has_table("tasks"):
        op.create_table(
            "tasks",
            sa.Column("id", sa.String(length=36), primary_key=True),
            sa.Column("user_id", sa.Integer(), nullable=True),
            sa.Column("upload_id", sa.String(length=36), nullable=True),
            sa.Column("task_type", sa.String(length=20), nullable=False),
            sa.Column("status", sa.String(length=20), nullable=False),
            sa.Column("project_type", sa.String(length=50), nullable=True),
            sa.Column("template_path", sa.String(length=500), nullable=True),
            sa.Column("clinical_info_snapshot", sa.Text(), nullable=True),
            sa.Column("output_path", sa.String(length=500), nullable=True),
            sa.Column("preview_path", sa.String(length=500), nullable=True),
            sa.Column("context_json_path", sa.String(length=500), nullable=True),
            sa.Column("total_files", sa.Integer(), nullable=True),
            sa.Column("completed_files", sa.Integer(), nullable=True),
            sa.Column("failed_files", sa.Integer(), nullable=True),
            sa.Column("created_at", sa.DateTime(), server_default=sa.func.now()),
            sa.Column("started_at", sa.DateTime(), nullable=True),
            sa.Column("completed_at", sa.DateTime(), nullable=True),
            sa.Column("duration_seconds", sa.Float(), nullable=True),
            sa.Column("errors", sa.Text(), nullable=True),
            sa.Column("warnings", sa.Text(), nullable=True),
        )

    if not _has_table("task_results"):
        op.create_table(
            "task_results",
            sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
            sa.Column("task_id", sa.String(length=36), nullable=False),
            sa.Column("file_index", sa.Integer(), nullable=False),
            sa.Column("excel_filename", sa.String(length=255), nullable=False),
            sa.Column("status", sa.String(length=20), nullable=False),
            sa.Column("output_path", sa.String(length=500), nullable=True),
            sa.Column("duration_seconds", sa.Float(), nullable=True),
            sa.Column("errors", sa.Text(), nullable=True),
            sa.Column("warnings", sa.Text(), nullable=True),
            sa.Column("validation_summary", sa.Text(), nullable=True),
        )

    if not _has_table("audit_logs"):
        op.create_table(
            "audit_logs",
            sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
            sa.Column("user_id", sa.Integer(), nullable=True),
            sa.Column("action", sa.String(length=50), nullable=False),
            sa.Column("resource_type", sa.String(length=50), nullable=True),
            sa.Column("resource_id", sa.String(length=100), nullable=True),
            sa.Column("details", sa.Text(), nullable=True),
            sa.Column("ip_address", sa.String(length=45), nullable=True),
            sa.Column("created_at", sa.DateTime(), server_default=sa.func.now()),
        )


def downgrade() -> None:
    for table in ("audit_logs", "task_results", "tasks", "uploads", "users"):
        if _has_table(table):
            op.drop_table(table)
