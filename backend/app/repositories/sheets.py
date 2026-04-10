from __future__ import annotations

from sqlalchemy import Select, select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import joinedload, selectinload

from app.models import SheetModel, WorkbookModel


class SheetRepository:
    async def get_by_workspace_and_id(
        self,
        session: AsyncSession,
        workspace_id: str,
        sheet_id: str,
        *,
        include_grid: bool,
    ) -> SheetModel | None:
        statement = self._query(include_grid=include_grid).where(
            WorkbookModel.workspace_id == workspace_id,
            SheetModel.id == sheet_id,
        )
        result = await session.execute(statement)
        return result.scalar_one_or_none()

    def _query(self, *, include_grid: bool) -> Select[tuple[SheetModel]]:
        options = [
            joinedload(SheetModel.workbook).joinedload(WorkbookModel.workspace),
            selectinload(SheetModel.records),
        ]

        if include_grid:
            options.extend(
                [
                    selectinload(SheetModel.columns),
                    selectinload(SheetModel.records),
                    selectinload(SheetModel.views),
                ]
            )

        return (
            select(SheetModel)
            .join(SheetModel.workbook)
            .execution_options(populate_existing=True)
            .options(*options)
        )


sheet_repository = SheetRepository()
