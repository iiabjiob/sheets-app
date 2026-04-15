from __future__ import annotations

from sqlalchemy import Select, func, select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.models import SheetModel, SheetRecordModel, WorkbookModel, WorkspaceMemberModel, WorkspaceModel


class WorkspaceRepository:
    async def count(self, session: AsyncSession) -> int:
        result = await session.execute(select(func.count(WorkspaceModel.id)))
        return int(result.scalar_one())

    async def load_many(
        self,
        session: AsyncSession,
        *,
        user_id: str,
        include_grid: bool = False,
        include_members: bool = False,
        include_records: bool = True,
    ) -> list[WorkspaceModel]:
        result = await session.execute(
            self._query(
                user_id=user_id,
                include_grid=include_grid,
                include_members=include_members,
                include_records=include_records,
            )
        )
        return list(result.scalars().unique().all())

    async def get_by_id(
        self,
        session: AsyncSession,
        workspace_id: str,
        *,
        user_id: str,
        include_grid: bool = False,
        include_members: bool = False,
        include_records: bool = True,
    ) -> WorkspaceModel | None:
        statement = self._query(
            user_id=user_id,
            include_grid=include_grid,
            include_members=include_members,
            include_records=include_records,
        ).where(WorkspaceModel.id == workspace_id)
        result = await session.execute(statement)
        return result.scalar_one_or_none()

    async def existing_slugs(
        self,
        session: AsyncSession,
        *,
        exclude_workspace_id: str | None = None,
    ) -> set[str]:
        statement = select(WorkspaceModel.slug)
        if exclude_workspace_id is not None:
            statement = statement.where(WorkspaceModel.id != exclude_workspace_id)
        result = await session.execute(statement)
        return {value for value in result.scalars().all()}

    async def next_position(self, session: AsyncSession, user_id: str) -> int:
        result = await session.execute(
            select(func.min(WorkspaceModel.position))
            .select_from(WorkspaceModel)
            .join(WorkspaceMemberModel, WorkspaceMemberModel.workspace_id == WorkspaceModel.id)
            .where(WorkspaceMemberModel.user_id == user_id)
        )
        current_min = result.scalar_one_or_none()
        return (current_min - 1) if current_min is not None else 0

    def _query(
        self,
        *,
        user_id: str | None = None,
        include_grid: bool = False,
        include_members: bool = False,
        include_records: bool = True,
    ) -> Select[tuple[WorkspaceModel]]:
        options = [selectinload(WorkspaceModel.workbooks).selectinload(WorkbookModel.sheets)]

        if include_records or include_grid:
            options.extend(
                [
                    selectinload(WorkspaceModel.workbooks)
                    .selectinload(WorkbookModel.sheets)
                    .selectinload(SheetModel.records)
                    .selectinload(SheetRecordModel.creator),
                    selectinload(WorkspaceModel.workbooks)
                    .selectinload(WorkbookModel.sheets)
                    .selectinload(SheetModel.records)
                    .selectinload(SheetRecordModel.updater),
                ]
            )

        if include_members:
            options.append(selectinload(WorkspaceModel.members))

        if include_grid:
            options.extend(
                [
                    selectinload(WorkspaceModel.workbooks)
                    .selectinload(WorkbookModel.sheets)
                    .selectinload(SheetModel.columns),
                    selectinload(WorkspaceModel.workbooks)
                    .selectinload(WorkbookModel.sheets)
                    .selectinload(SheetModel.views),
                ]
            )

        statement = (
            select(WorkspaceModel)
            .execution_options(populate_existing=True)
            .options(*options)
            .order_by(WorkspaceModel.position.asc(), WorkspaceModel.created_at.asc())
        )

        if user_id is not None:
            statement = statement.join(WorkspaceMemberModel).where(WorkspaceMemberModel.user_id == user_id)

        return statement


workspace_repository = WorkspaceRepository()
