from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

from sqlalchemy import Boolean, DateTime, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.infrastructure.db.database import Base
from app.models.common import utc_now

if TYPE_CHECKING:
    from app.models.activity import ActivityLogModel
    from app.models.sheet import SheetRecordModel, SheetViewModel
    from app.models.workbook import WorkbookModel
    from app.models.workspace import WorkspaceMemberModel, WorkspaceModel


class UserModel(Base):
    __tablename__ = "users"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    email: Mapped[str] = mapped_column(String(255), nullable=False, unique=True, index=True)
    full_name: Mapped[str] = mapped_column(String(120), nullable=False)
    password_hash: Mapped[str] = mapped_column(String(512), nullable=False, default="!")
    email_verified_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    email_verification_token: Mapped[str | None] = mapped_column(String(128), nullable=True)
    email_verification_sent_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    avatar_url: Mapped[str | None] = mapped_column(String(512), nullable=True)
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    owned_workspaces: Mapped[list[WorkspaceModel]] = relationship(back_populates="owner")
    workspace_memberships: Mapped[list[WorkspaceMemberModel]] = relationship(back_populates="user")
    created_workbooks: Mapped[list[WorkbookModel]] = relationship(
        back_populates="creator",
        foreign_keys="WorkbookModel.created_by",
    )
    updated_workbooks: Mapped[list[WorkbookModel]] = relationship(
        back_populates="updater",
        foreign_keys="WorkbookModel.updated_by",
    )
    created_records: Mapped[list[SheetRecordModel]] = relationship(
        back_populates="creator",
        foreign_keys="SheetRecordModel.created_by",
    )
    updated_records: Mapped[list[SheetRecordModel]] = relationship(
        back_populates="updater",
        foreign_keys="SheetRecordModel.updated_by",
    )
    created_views: Mapped[list[SheetViewModel]] = relationship(back_populates="creator")
    activity_logs: Mapped[list[ActivityLogModel]] = relationship(back_populates="user")