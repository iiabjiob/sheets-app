from __future__ import annotations

from datetime import datetime
from typing import Any

from sqlalchemy import JSON, DateTime, ForeignKey, Index, String
from sqlalchemy.orm import Mapped, mapped_column

from app.infrastructure.db.database import Base
from app.models.common import utc_now


class SheetRevisionModel(Base):
    __tablename__ = "sheet_revisions"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    workspace_id: Mapped[str] = mapped_column(
        ForeignKey("workspaces.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    workbook_id: Mapped[str | None] = mapped_column(
        ForeignKey("workbooks.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    sheet_id: Mapped[str] = mapped_column(
        ForeignKey("sheets.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    user_id: Mapped[str | None] = mapped_column(
        ForeignKey("users.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    action_type: Mapped[str] = mapped_column(String(64), nullable=False, default="sheet_grid_updated")
    payload_json: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False, default=dict)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)


class SheetCellRevisionModel(Base):
    __tablename__ = "sheet_cell_revisions"
    __table_args__ = (
        Index(
            "ix_sheet_cell_revisions_sheet_record_column_created",
            "sheet_id",
            "record_id",
            "column_key",
            "created_at",
        ),
    )

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    revision_id: Mapped[str] = mapped_column(
        ForeignKey("sheet_revisions.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    sheet_id: Mapped[str] = mapped_column(
        ForeignKey("sheets.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    record_id: Mapped[str] = mapped_column(String(32), nullable=False, index=True)
    column_key: Mapped[str] = mapped_column(String(120), nullable=False, index=True)
    previous_value_json: Mapped[Any | None] = mapped_column(JSON, nullable=True)
    next_value_json: Mapped[Any | None] = mapped_column(JSON, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)