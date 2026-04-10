from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

from sqlalchemy import DateTime, ForeignKey, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.infrastructure.db.database import Base
from app.models.common import utc_now

if TYPE_CHECKING:
    from app.models.activity import ActivityLogModel
    from app.models.user import UserModel
    from app.models.workbook import WorkbookModel


class WorkspaceModel(Base):
    __tablename__ = "workspaces"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    name: Mapped[str] = mapped_column(String(80), nullable=False)
    slug: Mapped[str] = mapped_column(String(120), nullable=False, unique=True, index=True)
    owner_id: Mapped[str] = mapped_column(ForeignKey("users.id", ondelete="RESTRICT"), nullable=False, index=True)
    description: Mapped[str] = mapped_column(Text, nullable=False)
    color: Mapped[str] = mapped_column(String(16), nullable=False)
    position: Mapped[int] = mapped_column(Integer, nullable=False, default=0, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    owner: Mapped[UserModel] = relationship(back_populates="owned_workspaces")
    members: Mapped[list[WorkspaceMemberModel]] = relationship(
        back_populates="workspace",
        cascade="all, delete-orphan",
    )
    workbooks: Mapped[list[WorkbookModel]] = relationship(
        back_populates="workspace",
        cascade="all, delete-orphan",
        order_by="WorkbookModel.position",
    )
    activity_logs: Mapped[list[ActivityLogModel]] = relationship(back_populates="workspace")


class WorkspaceMemberModel(Base):
    __tablename__ = "workspace_members"
    __table_args__ = (
        UniqueConstraint("workspace_id", "user_id", name="uq_workspace_members_workspace_user"),
    )

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    workspace_id: Mapped[str] = mapped_column(
        ForeignKey("workspaces.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    user_id: Mapped[str] = mapped_column(
        ForeignKey("users.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    role: Mapped[str] = mapped_column(String(24), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    workspace: Mapped[WorkspaceModel] = relationship(back_populates="members")
    user: Mapped[UserModel] = relationship(back_populates="workspace_memberships")