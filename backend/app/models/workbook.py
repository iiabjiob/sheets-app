from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

from sqlalchemy import Boolean, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.infrastructure.db.database import Base
from app.models.common import utc_now

if TYPE_CHECKING:
    from app.models.activity import ActivityLogModel
    from app.models.sheet import SheetModel
    from app.models.user import UserModel
    from app.models.workspace import WorkspaceModel


class WorkbookModel(Base):
    __tablename__ = "workbooks"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    workspace_id: Mapped[str] = mapped_column(
        ForeignKey("workspaces.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    name: Mapped[str] = mapped_column(String(120), nullable=False)
    description: Mapped[str | None] = mapped_column(Text, nullable=True)
    position: Mapped[int] = mapped_column(Integer, nullable=False, default=0, index=True)
    created_by: Mapped[str] = mapped_column(ForeignKey("users.id", ondelete="RESTRICT"), nullable=False)
    updated_by: Mapped[str] = mapped_column(ForeignKey("users.id", ondelete="RESTRICT"), nullable=False)
    archived: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    workspace: Mapped[WorkspaceModel] = relationship(back_populates="workbooks")
    creator: Mapped[UserModel] = relationship(back_populates="created_workbooks", foreign_keys=[created_by])
    updater: Mapped[UserModel] = relationship(back_populates="updated_workbooks", foreign_keys=[updated_by])
    sheets: Mapped[list[SheetModel]] = relationship(
        back_populates="workbook",
        cascade="all, delete-orphan",
        order_by="SheetModel.position",
    )
    activity_logs: Mapped[list[ActivityLogModel]] = relationship(back_populates="workbook")