from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Any

from sqlalchemy import JSON, DateTime, ForeignKey, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.infrastructure.db.database import Base
from app.models.common import utc_now

if TYPE_CHECKING:
    from app.models.sheet import SheetModel, SheetRecordModel
    from app.models.user import UserModel
    from app.models.workbook import WorkbookModel
    from app.models.workspace import WorkspaceModel


class ActivityLogModel(Base):
    __tablename__ = "activity_logs"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    workspace_id: Mapped[str | None] = mapped_column(
        ForeignKey("workspaces.id", ondelete="CASCADE"),
        nullable=True,
        index=True,
    )
    workbook_id: Mapped[str | None] = mapped_column(
        ForeignKey("workbooks.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    sheet_id: Mapped[str | None] = mapped_column(
        ForeignKey("sheets.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    record_id: Mapped[str | None] = mapped_column(
        ForeignKey("sheet_records.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    user_id: Mapped[str | None] = mapped_column(
        ForeignKey("users.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    action_type: Mapped[str] = mapped_column(String(64), nullable=False, index=True)
    payload_json: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False, default=dict)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    workspace: Mapped[WorkspaceModel | None] = relationship(back_populates="activity_logs")
    workbook: Mapped[WorkbookModel | None] = relationship(back_populates="activity_logs")
    sheet: Mapped[SheetModel | None] = relationship(back_populates="activity_logs")
    record: Mapped[SheetRecordModel | None] = relationship(back_populates="activity_logs")
    user: Mapped[UserModel | None] = relationship(back_populates="activity_logs")