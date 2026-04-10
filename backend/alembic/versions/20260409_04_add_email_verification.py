"""add email verification fields

Revision ID: 20260409_04
Revises: 20260409_03
Create Date: 2026-04-09 05:00:00.000000

"""

from __future__ import annotations

from datetime import UTC, datetime

from alembic import op
import sqlalchemy as sa


revision = "20260409_04"
down_revision = "20260409_03"
branch_labels = None
depends_on = None


def upgrade() -> None:
    now = datetime.now(UTC)
    op.add_column("users", sa.Column("email_verified_at", sa.DateTime(timezone=True), nullable=True))
    op.add_column("users", sa.Column("email_verification_token", sa.String(length=128), nullable=True))
    op.add_column("users", sa.Column("email_verification_sent_at", sa.DateTime(timezone=True), nullable=True))

    op.execute(
        sa.text(
            """
            UPDATE users
            SET email_verified_at = :now,
                email_verification_token = NULL,
                email_verification_sent_at = NULL
            WHERE email_verified_at IS NULL
            """
        ).bindparams(now=now)
    )


def downgrade() -> None:
    op.drop_column("users", "email_verification_sent_at")
    op.drop_column("users", "email_verification_token")
    op.drop_column("users", "email_verified_at")