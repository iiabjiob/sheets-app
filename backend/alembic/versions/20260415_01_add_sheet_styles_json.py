"""add sheet styles json column

Revision ID: 20260415_01
Revises: 20260413_01
Create Date: 2026-04-15 00:00:00.000000

"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "20260415_01"
down_revision = "20260413_01"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.add_column(
        "sheets",
        sa.Column(
            "styles_json",
            sa.JSON(),
            nullable=False,
            server_default=sa.text("'[]'"),
        ),
    )
    op.alter_column("sheets", "styles_json", server_default=None)


def downgrade() -> None:
    op.drop_column("sheets", "styles_json")