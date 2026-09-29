"""Persist the safe OpenRouter role manifest reported by the worker."""
from alembic import op
import sqlalchemy as sa

revision = "c2d8e3f4a5b6"
down_revision = "b1c7d2e3f4a5"
branch_labels = None
depends_on = None


def upgrade():
    op.add_column("worker_capabilities", sa.Column("configuration_json", sa.JSON(), nullable=False, server_default="{}"))


def downgrade():
    op.drop_column("worker_capabilities", "configuration_json")
