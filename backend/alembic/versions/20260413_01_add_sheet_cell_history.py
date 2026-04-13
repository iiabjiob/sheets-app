"""add sheet cell history tables

Revision ID: 20260413_01
Revises: 20260409_04
Create Date: 2026-04-13 00:00:00.000000

"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "20260413_01"
down_revision = "20260409_04"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "sheet_revisions",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("workspace_id", sa.String(length=32), nullable=False),
        sa.Column("workbook_id", sa.String(length=32), nullable=True),
        sa.Column("sheet_id", sa.String(length=32), nullable=False),
        sa.Column("user_id", sa.String(length=32), nullable=True),
        sa.Column("action_type", sa.String(length=64), nullable=False),
        sa.Column("payload_json", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["sheet_id"], ["sheets.id"], ondelete="CASCADE"),
        sa.ForeignKeyConstraint(["user_id"], ["users.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["workbook_id"], ["workbooks.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["workspace_id"], ["workspaces.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index(op.f("ix_sheet_revisions_sheet_id"), "sheet_revisions", ["sheet_id"], unique=False)
    op.create_index(op.f("ix_sheet_revisions_user_id"), "sheet_revisions", ["user_id"], unique=False)
    op.create_index(op.f("ix_sheet_revisions_workbook_id"), "sheet_revisions", ["workbook_id"], unique=False)
    op.create_index(op.f("ix_sheet_revisions_workspace_id"), "sheet_revisions", ["workspace_id"], unique=False)

    op.create_table(
        "sheet_cell_revisions",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("revision_id", sa.String(length=32), nullable=False),
        sa.Column("sheet_id", sa.String(length=32), nullable=False),
        sa.Column("record_id", sa.String(length=32), nullable=False),
        sa.Column("column_key", sa.String(length=120), nullable=False),
        sa.Column("previous_value_json", sa.JSON(), nullable=True),
        sa.Column("next_value_json", sa.JSON(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["revision_id"], ["sheet_revisions.id"], ondelete="CASCADE"),
        sa.ForeignKeyConstraint(["sheet_id"], ["sheets.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index(op.f("ix_sheet_cell_revisions_column_key"), "sheet_cell_revisions", ["column_key"], unique=False)
    op.create_index(op.f("ix_sheet_cell_revisions_record_id"), "sheet_cell_revisions", ["record_id"], unique=False)
    op.create_index(op.f("ix_sheet_cell_revisions_revision_id"), "sheet_cell_revisions", ["revision_id"], unique=False)
    op.create_index(op.f("ix_sheet_cell_revisions_sheet_id"), "sheet_cell_revisions", ["sheet_id"], unique=False)
    op.create_index(
        "ix_sheet_cell_revisions_sheet_record_column_created",
        "sheet_cell_revisions",
        ["sheet_id", "record_id", "column_key", "created_at"],
        unique=False,
    )


def downgrade() -> None:
    op.drop_index("ix_sheet_cell_revisions_sheet_record_column_created", table_name="sheet_cell_revisions")
    op.drop_index(op.f("ix_sheet_cell_revisions_sheet_id"), table_name="sheet_cell_revisions")
    op.drop_index(op.f("ix_sheet_cell_revisions_revision_id"), table_name="sheet_cell_revisions")
    op.drop_index(op.f("ix_sheet_cell_revisions_record_id"), table_name="sheet_cell_revisions")
    op.drop_index(op.f("ix_sheet_cell_revisions_column_key"), table_name="sheet_cell_revisions")
    op.drop_table("sheet_cell_revisions")

    op.drop_index(op.f("ix_sheet_revisions_workspace_id"), table_name="sheet_revisions")
    op.drop_index(op.f("ix_sheet_revisions_workbook_id"), table_name="sheet_revisions")
    op.drop_index(op.f("ix_sheet_revisions_user_id"), table_name="sheet_revisions")
    op.drop_index(op.f("ix_sheet_revisions_sheet_id"), table_name="sheet_revisions")
    op.drop_table("sheet_revisions")