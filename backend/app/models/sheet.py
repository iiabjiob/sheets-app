from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, Any

from sqlalchemy import Boolean, JSON, DateTime, ForeignKey, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.infrastructure.db.database import Base
from app.models.common import utc_now

if TYPE_CHECKING:
    from app.models.activity import ActivityLogModel
    from app.models.user import UserModel
    from app.models.workbook import WorkbookModel


class SheetModel(Base):
    __tablename__ = "sheets"
    __table_args__ = (UniqueConstraint("workbook_id", "key", name="uq_sheets_workbook_key"),)

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    workbook_id: Mapped[str] = mapped_column(
        ForeignKey("workbooks.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    key: Mapped[str] = mapped_column(String(120), nullable=False)
    name: Mapped[str] = mapped_column(String(80), nullable=False)
    kind: Mapped[str] = mapped_column(String(24), nullable=False, default="data")
    source_sheet_id: Mapped[str | None] = mapped_column(
        ForeignKey("sheets.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )
    position: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    config_json: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False, default=dict)
    styles_json: Mapped[list[dict[str, Any]]] = mapped_column(JSON, nullable=False, default=list)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    workbook: Mapped[WorkbookModel] = relationship(back_populates="sheets")
    source_sheet: Mapped[SheetModel | None] = relationship(remote_side=[id], back_populates="derived_sheets")
    derived_sheets: Mapped[list[SheetModel]] = relationship(back_populates="source_sheet")
    columns: Mapped[list[SheetColumnModel]] = relationship(
        back_populates="sheet",
        cascade="all, delete-orphan",
        order_by="SheetColumnModel.position",
    )
    records: Mapped[list[SheetRecordModel]] = relationship(
        back_populates="sheet",
        cascade="all, delete-orphan",
        order_by="SheetRecordModel.position",
    )
    views: Mapped[list[SheetViewModel]] = relationship(
        back_populates="sheet",
        cascade="all, delete-orphan",
        order_by="SheetViewModel.created_at",
    )
    activity_logs: Mapped[list[ActivityLogModel]] = relationship(back_populates="sheet")


class SheetColumnModel(Base):
    __tablename__ = "sheet_columns"
    __table_args__ = (UniqueConstraint("sheet_id", "key", name="uq_sheet_columns_sheet_key"),)

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    sheet_id: Mapped[str] = mapped_column(
        ForeignKey("sheets.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    key: Mapped[str] = mapped_column(String(120), nullable=False)
    title: Mapped[str] = mapped_column(String(120), nullable=False)
    type: Mapped[str] = mapped_column(String(32), nullable=False)
    position: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    width: Mapped[int | None] = mapped_column(Integer, nullable=True)
    nullable: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    editable: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    computed: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    expression: Mapped[str | None] = mapped_column(Text, nullable=True)
    settings_json: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False, default=dict)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    sheet: Mapped[SheetModel] = relationship(back_populates="columns")


class SheetRecordModel(Base):
    __tablename__ = "sheet_records"

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    sheet_id: Mapped[str] = mapped_column(
        ForeignKey("sheets.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    position: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    values_json: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False, default=dict)
    display_values_json: Mapped[dict[str, Any] | None] = mapped_column(JSON, nullable=True)
    created_by: Mapped[str] = mapped_column(ForeignKey("users.id", ondelete="RESTRICT"), nullable=False)
    updated_by: Mapped[str] = mapped_column(ForeignKey("users.id", ondelete="RESTRICT"), nullable=False)
    archived: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    sheet: Mapped[SheetModel] = relationship(back_populates="records")
    creator: Mapped[UserModel] = relationship(back_populates="created_records", foreign_keys=[created_by])
    updater: Mapped[UserModel] = relationship(back_populates="updated_records", foreign_keys=[updated_by])
    activity_logs: Mapped[list[ActivityLogModel]] = relationship(back_populates="record")


class SheetViewModel(Base):
    __tablename__ = "sheet_views"
    __table_args__ = (UniqueConstraint("sheet_id", "key", name="uq_sheet_views_sheet_key"),)

    id: Mapped[str] = mapped_column(String(32), primary_key=True)
    sheet_id: Mapped[str] = mapped_column(
        ForeignKey("sheets.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    key: Mapped[str] = mapped_column(String(120), nullable=False)
    name: Mapped[str] = mapped_column(String(120), nullable=False)
    type: Mapped[str] = mapped_column(String(24), nullable=False, default="grid")
    is_default: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    created_by: Mapped[str] = mapped_column(ForeignKey("users.id", ondelete="RESTRICT"), nullable=False)
    config_json: Mapped[dict[str, Any]] = mapped_column(JSON, nullable=False, default=dict)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=utc_now)

    sheet: Mapped[SheetModel] = relationship(back_populates="views")
    creator: Mapped[UserModel] = relationship(back_populates="created_views")