"""add user password hash

Revision ID: 20260409_03
Revises: 20260409_02
Create Date: 2026-04-09 03:10:00.000000

"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "20260409_03"
down_revision = "20260409_02"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column(
        "users",
        sa.Column("password_hash", sa.String(length=512), nullable=False, server_default="!"),
    )
    op.alter_column("users", "password_hash", server_default=None)


def downgrade() -> None:
    op.drop_column("users", "password_hash")