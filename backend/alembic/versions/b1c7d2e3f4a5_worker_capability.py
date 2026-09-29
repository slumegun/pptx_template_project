"""worker capability report without API credentials

Revision ID: b1c7d2e3f4a5
Revises: 8793ca2210cf
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa

revision: str = "b1c7d2e3f4a5"
down_revision: Union[str, Sequence[str], None] = "8793ca2210cf"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "worker_capabilities",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("model_mode", sa.String(length=32), nullable=False),
        sa.Column("text_model", sa.String(length=160), nullable=True),
        sa.Column("vision_model", sa.String(length=160), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("id"),
    )


def downgrade() -> None:
    op.drop_table("worker_capabilities")

