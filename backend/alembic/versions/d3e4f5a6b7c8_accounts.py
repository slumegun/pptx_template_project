"""Accounts, sessions and ownership of presentation workspaces."""

from alembic import op
import sqlalchemy as sa

revision = "d3e4f5a6b7c8"
down_revision = "c2d8e3f4a5b6"
branch_labels = None
depends_on = None


def upgrade():
    bind = op.get_bind()
    if not sa.inspect(bind).has_table("users"):
        op.create_table(
            "users",
            sa.Column("id", sa.String(36), primary_key=True),
            sa.Column("email", sa.String(320), nullable=True),
            sa.Column("password_hash", sa.String(256), nullable=True),
            sa.Column("vk_id", sa.String(80), nullable=True),
            sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
            sa.UniqueConstraint("email"),
            sa.UniqueConstraint("vk_id"),
        )
    if "ix_users_email" not in {index["name"] for index in sa.inspect(bind).get_indexes("users")}:
        op.create_index("ix_users_email", "users", ["email"])
    if not sa.inspect(bind).has_table("session_tokens"):
        op.create_table(
            "session_tokens",
            sa.Column("id", sa.String(64), primary_key=True),
            sa.Column("user_id", sa.String(36), sa.ForeignKey("users.id", ondelete="CASCADE"), nullable=False),
            sa.Column("expires_at", sa.DateTime(timezone=True), nullable=False),
            sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        )
    if "ix_session_tokens_user_id" not in {index["name"] for index in sa.inspect(bind).get_indexes("session_tokens")}:
        op.create_index("ix_session_tokens_user_id", "session_tokens", ["user_id"])
    if "owner_user_id" not in {column["name"] for column in sa.inspect(bind).get_columns("projects")}:
        with op.batch_alter_table("projects") as batch:
            batch.add_column(sa.Column("owner_user_id", sa.String(36), nullable=True))
            batch.create_foreign_key("fk_projects_owner_user_id", "users", ["owner_user_id"], ["id"])
    if "ix_projects_owner_user_id" not in {index["name"] for index in sa.inspect(bind).get_indexes("projects")}:
        op.create_index("ix_projects_owner_user_id", "projects", ["owner_user_id"])


def downgrade():
    op.drop_index("ix_projects_owner_user_id", table_name="projects")
    with op.batch_alter_table("projects") as batch:
        batch.drop_column("owner_user_id")
    op.drop_index("ix_session_tokens_user_id", table_name="session_tokens")
    op.drop_table("session_tokens")
    op.drop_index("ix_users_email", table_name="users")
    op.drop_table("users")
