"""Shared template library independent of accounts and projects."""

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects.postgresql import JSONB

revision = "e5f6a7b8c9d0"
down_revision = "d3e4f5a6b7c8"
branch_labels = None
depends_on = None


def upgrade():
    if sa.inspect(op.get_bind()).has_table("templates"):
        return
    op.create_table(
        "templates",
        sa.Column("id", sa.String(36), primary_key=True),
        sa.Column("name", sa.String(200), nullable=False),
        sa.Column("filename", sa.String(255), nullable=False),
        sa.Column("storage_key", sa.String(512), nullable=False),
        sa.Column("sha256", sa.String(64), nullable=False),
        sa.Column("size_bytes", sa.Integer(), nullable=False),
        sa.Column("slide_count", sa.Integer(), nullable=True),
        sa.Column("origin", sa.String(24), nullable=False),
        sa.Column("preparation_status", sa.String(32), nullable=False),
        sa.Column("preparation_stage", sa.String(64), nullable=True),
        sa.Column("preparation_error", sa.Text(), nullable=True),
        sa.Column("prepared_key", sa.String(512), nullable=True),
        sa.Column("prepared_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("metadata", sa.JSON().with_variant(JSONB, "postgresql"), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.UniqueConstraint("storage_key"),
    )
    op.create_index("ix_templates_sha256", "templates", ["sha256"])
    op.create_index("ix_templates_preparation_status", "templates", ["preparation_status"])


def downgrade():
    op.drop_index("ix_templates_preparation_status", table_name="templates")
    op.drop_index("ix_templates_sha256", table_name="templates")
    op.drop_table("templates")
