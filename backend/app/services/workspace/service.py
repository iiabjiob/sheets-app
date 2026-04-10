from __future__ import annotations

from copy import deepcopy
from uuid import uuid4

from fastapi import HTTPException, status
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models import ActivityLogModel, SheetModel, WorkbookModel, WorkspaceMemberModel, WorkspaceModel
from app.repositories import workspace_repository
from app.services.workspace.builders import (
    build_columns_from_grid_input,
    build_records,
    build_sheet,
    build_workbook,
    clone_sheet,
    clone_workbook,
    normalize_grid_rows,
)
from app.services.workspace.naming import copy_name, timestamp, unique_key, unique_slug
from app.services.workspace.ordering import (
    reindex_sheets,
    reindex_workspaces,
    resolve_target_index,
    sort_sheets,
    sort_workbooks,
)
from app.services.workspace.serialization import (
    serialize_sheet,
    serialize_sheet_workbook_context,
    serialize_workspace,
    workspace_color,
)


class WorkspaceService:
    async def list_workspaces(self, session: AsyncSession, user_id: str) -> list[dict[str, object]]:
        workspaces = await self._load_workspaces(session, user_id=user_id)
        return [serialize_workspace(workspace) for workspace in workspaces]

    async def get_sheet_document(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
    ) -> dict[str, object]:
        workspace = await self._require_workspace(session, user_id, workspace_id, include_grid=True)
        sheet = self._find_sheet(workspace, sheet_id)
        if sheet is None:
            raise self._not_found("Sheet")

        return {
            "workspace": serialize_workspace(workspace),
            "sheet": serialize_sheet(sheet, include_grid=True),
            "workbook": serialize_sheet_workbook_context(
                workspace,
                workbook_id=sheet.workbook_id,
            ),
        }

    async def create_workspace(self, session: AsyncSession, user_id: str, name: str) -> dict[str, object]:
        normalized_name = name.strip()
        if not normalized_name:
            raise self._validation_error("Workspace name must not be empty.")

        existing_workspaces = await self._load_workspaces(session, user_id=user_id)
        existing_slugs = {workspace.slug for workspace in existing_workspaces}
        created_at = timestamp()
        position = await self._next_workspace_position(session, user_id)
        workbook_id = f"wb_{uuid4().hex[:8]}"
        primary_sheet = build_sheet(
            workbook_id=workbook_id,
            user_id=user_id,
            name=f"{normalized_name} sheet",
            position=0,
        )
        workspace = WorkspaceModel(
            id=f"ws_{uuid4().hex[:8]}",
            name=normalized_name,
            slug=unique_slug(normalized_name, existing_slugs),
            owner_id=user_id,
            description="New workspace for planning sheets and execution views.",
            color=workspace_color(position),
            position=position,
            created_at=created_at,
            updated_at=created_at,
        )
        workspace.members = [
            WorkspaceMemberModel(
                id=f"wm_{uuid4().hex[:8]}",
                workspace_id=workspace.id,
                user_id=user_id,
                role="owner",
                created_at=created_at,
            )
        ]
        workspace.workbooks = [
            build_workbook(
                workspace_id=workspace.id,
                user_id=user_id,
                name=f"{normalized_name} workbook",
                description="Primary workbook for grid, gantt, and derived sheet views.",
                position=0,
                workbook_id=workbook_id,
                sheets=[primary_sheet],
            )
        ]

        session.add(workspace)
        self._append_activity(
            session,
            workspace_id=workspace.id,
            workbook_id=workbook_id,
            sheet_id=primary_sheet.id,
            user_id=user_id,
            action_type="workspace_created",
            payload={"workspaceName": normalized_name},
        )
        await session.commit()

        reloaded_workspace = await self._require_workspace(session, user_id, workspace.id, include_grid=False)
        return serialize_workspace(reloaded_workspace)

    async def delete_workspace(self, session: AsyncSession, user_id: str, workspace_id: str) -> None:
        workspace = await self._require_workspace(session, user_id, workspace_id, include_members=True)
        await session.delete(workspace)
        await session.commit()

    async def rename_workspace(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        name: str,
    ) -> dict[str, object]:
        normalized_name = name.strip()
        if not normalized_name:
            raise self._validation_error("Workspace name must not be empty.")

        workspace = await self._require_workspace(session, user_id, workspace_id)
        existing_slugs = await self._existing_workspace_slugs(session, exclude_workspace_id=workspace_id)
        workspace.name = normalized_name
        workspace.slug = unique_slug(normalized_name, existing_slugs)
        workspace.updated_at = timestamp()
        self._append_activity(
            session,
            workspace_id=workspace.id,
            user_id=user_id,
            action_type="workspace_renamed",
            payload={"workspaceName": normalized_name},
        )
        await session.commit()
        reloaded_workspace = await self._require_workspace(session, user_id, workspace.id)
        return serialize_workspace(reloaded_workspace)

    async def duplicate_workspace(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
    ) -> dict[str, object]:
        source_workspace = await self._require_workspace(
            session,
            user_id,
            workspace_id,
            include_grid=True,
            include_members=True,
        )
        existing_workspaces = await self._load_workspaces(
            session,
            user_id=user_id,
            include_grid=True,
            include_members=True,
        )
        source_index = next(
            (index for index, workspace in enumerate(existing_workspaces) if workspace.id == workspace_id),
            None,
        )
        if source_index is None:
            raise self._not_found("Workspace")

        duplicate_name = copy_name(
            source_workspace.name,
            {workspace.name.lower() for workspace in existing_workspaces},
        )
        created_at = timestamp()
        duplicate_workspace = WorkspaceModel(
            id=f"ws_{uuid4().hex[:8]}",
            name=duplicate_name,
            slug=unique_slug(duplicate_name, {workspace.slug for workspace in existing_workspaces}),
            owner_id=user_id,
            description=source_workspace.description,
            color=source_workspace.color,
            position=source_workspace.position,
            created_at=created_at,
            updated_at=created_at,
        )
        duplicate_workspace.members = [
            WorkspaceMemberModel(
                id=f"wm_{uuid4().hex[:8]}",
                workspace_id=duplicate_workspace.id,
                user_id=user_id,
                role="owner",
                created_at=created_at,
            )
        ]
        duplicate_workspace.workbooks = [
            clone_workbook(workbook, workspace_id=duplicate_workspace.id, position=index)
            for index, workbook in enumerate(sort_workbooks(source_workspace.workbooks))
        ]

        session.add(duplicate_workspace)
        ordered_workspaces = list(existing_workspaces)
        ordered_workspaces.insert(source_index + 1, duplicate_workspace)
        reindex_workspaces(ordered_workspaces)
        self._append_activity(
            session,
            workspace_id=duplicate_workspace.id,
            user_id=user_id,
            action_type="workspace_duplicated",
            payload={"sourceWorkspaceId": source_workspace.id},
        )
        await session.commit()

        reloaded_workspace = await self._require_workspace(session, user_id, duplicate_workspace.id)
        return serialize_workspace(reloaded_workspace)

    async def move_workspace(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        direction: str,
    ) -> None:
        workspaces = await self._load_workspaces(session, user_id=user_id)
        source_index = next(
            (index for index, workspace in enumerate(workspaces) if workspace.id == workspace_id),
            None,
        )
        if source_index is None:
            raise self._not_found("Workspace")

        target_index = resolve_target_index(source_index, len(workspaces), direction)
        if target_index is None:
            if direction in {"up", "down"}:
                return
            raise self._validation_error("Direction must be either 'up' or 'down'.")

        workspaces[source_index], workspaces[target_index] = workspaces[target_index], workspaces[source_index]
        reindex_workspaces(workspaces)
        updated_at = timestamp()
        workspaces[source_index].updated_at = updated_at
        workspaces[target_index].updated_at = updated_at
        self._append_activity(
            session,
            workspace_id=workspace_id,
            user_id=user_id,
            action_type="workspace_moved",
            payload={"direction": direction},
        )
        await session.commit()

    async def create_sheet(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        name: str,
    ) -> dict[str, object]:
        normalized_name = name.strip()
        if not normalized_name:
            raise self._validation_error("Sheet name must not be empty.")

        workspace = await self._require_workspace(session, user_id, workspace_id, include_records=False)
        workbook = self._primary_workbook(workspace)
        if workbook is None:
            raise self._not_found("Workbook")

        existing_keys = {sheet.key for sheet in workbook.sheets}
        workbook.updated_by = user_id
        sheet = build_sheet(
            workbook_id=workbook.id,
            user_id=user_id,
            name=normalized_name,
            key=unique_key(normalized_name, existing_keys),
            position=len(sort_sheets(workbook.sheets)),
        )
        workbook.sheets.append(sheet)
        workbook.updated_at = timestamp()
        workspace.updated_at = workbook.updated_at
        session.add(sheet)
        self._append_activity(
            session,
            workspace_id=workspace.id,
            workbook_id=workbook.id,
            sheet_id=sheet.id,
            user_id=user_id,
            action_type="sheet_created",
            payload={"sheetName": normalized_name},
        )
        await session.commit()

        reloaded_sheet = await self._require_sheet(session, user_id, workspace.id, sheet.id, include_grid=False)
        return serialize_sheet(reloaded_sheet)

    async def update_sheet_grid(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
        *,
        columns: list[dict[str, object]],
        rows: list[dict[str, object]],
    ) -> dict[str, object]:
        if not columns:
            raise self._validation_error("Sheet must contain at least one column.")

        seen_keys: set[str] = set()
        normalized_columns: list[dict[str, object]] = []
        for index, column in enumerate(columns):
            key = str(column.get("key") or "").strip()
            label = str(column.get("label") or "").strip()
            if not key:
                raise self._validation_error(f"Column {index + 1} is missing a key.")
            if not label:
                raise self._validation_error(f"Column {index + 1} is missing a label.")
            if key in seen_keys:
                raise self._validation_error(f"Column key '{key}' must be unique.")

            seen_keys.add(key)
            expression = str(column.get("expression") or "").strip()
            requested_column_type = str(column.get("column_type") or "").strip().lower()
            computed = bool(column.get("computed", False) or expression or requested_column_type == "formula")
            normalized_columns.append(
                {
                    "key": key,
                    "label": label,
                    "data_type": str(column.get("data_type") or "text"),
                    "column_type": str(column.get("column_type") or ("formula" if computed else "")),
                    "width": column.get("width"),
                    "editable": False if computed else bool(column.get("editable", True)),
                    "computed": computed,
                    "expression": expression or None,
                    "options": column.get("options") if isinstance(column.get("options"), list) else [],
                    "settings": column.get("settings") if isinstance(column.get("settings"), dict) else {},
                }
            )

        sheet = await self._require_sheet(session, user_id, workspace_id, sheet_id, include_grid=True)
        updated_at = timestamp()
        writable_column_keys = [
            str(column["key"]) for column in normalized_columns if not bool(column.get("computed"))
        ]
        normalized_rows = normalize_grid_rows(rows=rows, column_keys=writable_column_keys)
        workbook = await self._touch_sheet_parents(
            session,
            user_id=user_id,
            workspace_id=workspace_id,
            workbook_id=sheet.workbook_id,
            updated_at=updated_at,
        )

        # Replace child collections in two phases so unique constraints on existing
        # sheet columns/records are released before new rows are inserted.
        sheet.columns.clear()
        sheet.records.clear()
        await session.flush()

        sheet.columns = build_columns_from_grid_input(sheet_id=sheet.id, columns=normalized_columns)
        sheet.records = build_records(
            sheet_id=sheet.id,
            user_id=user_id,
            rows=normalized_rows,
        )
        sheet.updated_at = updated_at
        self._append_activity(
            session,
            workspace_id=workspace_id,
            workbook_id=sheet.workbook_id,
            sheet_id=sheet.id,
            user_id=user_id,
            action_type="sheet_grid_updated",
            payload={
                "columnCount": len(sheet.columns),
                "rowCount": len(sheet.records),
            },
        )
        workbook.updated_at = updated_at
        await session.commit()

        reloaded_sheet = await self._require_sheet(session, user_id, workspace_id, sheet.id, include_grid=True)
        return serialize_sheet(reloaded_sheet, include_grid=True)

    async def delete_sheet(self, session: AsyncSession, user_id: str, workspace_id: str, sheet_id: str) -> None:
        sheet = await self._require_sheet(session, user_id, workspace_id, sheet_id, include_grid=False)
        updated_at = timestamp()
        await self._touch_sheet_parents(
            session,
            user_id=user_id,
            workspace_id=workspace_id,
            workbook_id=sheet.workbook_id,
            updated_at=updated_at,
        )
        await session.delete(sheet)
        await session.commit()

    async def rename_sheet(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
        name: str,
    ) -> dict[str, object]:
        normalized_name = name.strip()
        if not normalized_name:
            raise self._validation_error("Sheet name must not be empty.")

        sheet = await self._require_sheet(session, user_id, workspace_id, sheet_id, include_grid=False)
        updated_at = timestamp()
        await self._touch_sheet_parents(
            session,
            user_id=user_id,
            workspace_id=workspace_id,
            workbook_id=sheet.workbook_id,
            updated_at=updated_at,
        )
        sheet.name = normalized_name
        sheet.updated_at = updated_at
        self._append_activity(
            session,
            workspace_id=workspace_id,
            workbook_id=sheet.workbook_id,
            sheet_id=sheet.id,
            user_id=user_id,
            action_type="sheet_renamed",
            payload={"sheetName": normalized_name},
        )
        await session.commit()
        reloaded_sheet = await self._require_sheet(session, user_id, workspace_id, sheet.id, include_grid=False)
        return serialize_sheet(reloaded_sheet)

    async def duplicate_sheet(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
    ) -> dict[str, object]:
        workspace = await self._require_workspace(session, user_id, workspace_id, include_grid=True)
        sheet = self._find_sheet(workspace, sheet_id)
        if sheet is None:
            raise self._not_found("Sheet")

        workbook = sheet.workbook
        ordered_sheets = sort_sheets(workbook.sheets)
        source_index = next((index for index, item in enumerate(ordered_sheets) if item.id == sheet_id), None)
        if source_index is None:
            raise self._not_found("Sheet")

        duplicate_name = copy_name(sheet.name, {item.name.lower() for item in ordered_sheets})
        duplicate_sheet = clone_sheet(
            sheet,
            workbook_id=workbook.id,
            position=sheet.position,
            name=duplicate_name,
            key=unique_key(duplicate_name, {item.key for item in ordered_sheets}),
        )

        session.add(duplicate_sheet)
        ordered_sheets.insert(source_index + 1, duplicate_sheet)
        reindex_sheets(ordered_sheets)
        workbook.updated_by = user_id
        workbook.updated_at = timestamp()
        workspace.updated_at = workbook.updated_at
        self._append_activity(
            session,
            workspace_id=workspace.id,
            workbook_id=workbook.id,
            sheet_id=duplicate_sheet.id,
            user_id=user_id,
            action_type="sheet_duplicated",
            payload={"sourceSheetId": sheet.id},
        )
        await session.commit()
        reloaded_sheet = await self._require_sheet(session, user_id, workspace_id, duplicate_sheet.id, include_grid=False)
        return serialize_sheet(reloaded_sheet)

    async def move_sheet(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
        direction: str,
    ) -> None:
        workspace = await self._require_workspace(session, user_id, workspace_id, include_grid=True)
        sheet = self._find_sheet(workspace, sheet_id)
        if sheet is None:
            raise self._not_found("Sheet")

        workbook = sheet.workbook
        ordered_sheets = sort_sheets(workbook.sheets)
        source_index = next((index for index, item in enumerate(ordered_sheets) if item.id == sheet_id), None)
        if source_index is None:
            raise self._not_found("Sheet")

        target_index = resolve_target_index(source_index, len(ordered_sheets), direction)
        if target_index is None:
            if direction in {"up", "down"}:
                return
            raise self._validation_error("Direction must be either 'up' or 'down'.")

        ordered_sheets[source_index], ordered_sheets[target_index] = (
            ordered_sheets[target_index],
            ordered_sheets[source_index],
        )
        reindex_sheets(ordered_sheets)
        workbook.updated_by = user_id
        workbook.updated_at = timestamp()
        workspace.updated_at = workbook.updated_at
        self._append_activity(
            session,
            workspace_id=workspace.id,
            workbook_id=workbook.id,
            sheet_id=sheet.id,
            user_id=user_id,
            action_type="sheet_moved",
            payload={"direction": direction},
        )
        await session.commit()

    async def _load_workspaces(
        self,
        session: AsyncSession,
        *,
        user_id: str,
        include_grid: bool = False,
        include_members: bool = False,
        include_records: bool = True,
    ) -> list[WorkspaceModel]:
        return await workspace_repository.load_many(
            session,
            user_id=user_id,
            include_grid=include_grid,
            include_members=include_members,
            include_records=include_records,
        )

    async def _require_workspace(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        *,
        include_grid: bool = False,
        include_members: bool = False,
        include_records: bool = True,
    ) -> WorkspaceModel:
        workspace = await workspace_repository.get_by_id(
            session,
            workspace_id,
            user_id=user_id,
            include_grid=include_grid,
            include_members=include_members,
            include_records=include_records,
        )
        if workspace is None:
            raise self._not_found("Workspace")
        return workspace

    async def _require_sheet(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
        *,
        include_grid: bool,
    ) -> SheetModel:
        workspace = await self._require_workspace(
            session,
            user_id,
            workspace_id,
            include_grid=include_grid,
        )
        sheet = self._find_sheet(workspace, sheet_id)
        if sheet is None:
            raise self._not_found("Sheet")
        return sheet

    async def _touch_sheet_parents(
        self,
        session: AsyncSession,
        *,
        user_id: str,
        workspace_id: str,
        workbook_id: str,
        updated_at,
    ) -> WorkbookModel:
        workbook = await session.scalar(
            select(WorkbookModel).where(
                WorkbookModel.id == workbook_id,
                WorkbookModel.workspace_id == workspace_id,
            )
        )
        if workbook is None:
            raise self._not_found("Workbook")

        workspace = await session.get(WorkspaceModel, workspace_id)
        if workspace is None:
            raise self._not_found("Workspace")

        workbook.updated_by = user_id
        workbook.updated_at = updated_at
        workspace.updated_at = updated_at
        return workbook

    def _primary_workbook(self, workspace: WorkspaceModel) -> WorkbookModel | None:
        candidates = [workbook for workbook in sort_workbooks(workspace.workbooks) if not workbook.archived]
        return candidates[0] if candidates else None

    def _find_sheet(self, workspace: WorkspaceModel, sheet_id: str) -> SheetModel | None:
        for workbook in sort_workbooks(workspace.workbooks):
            for sheet in sort_sheets(workbook.sheets):
                if sheet.id == sheet_id:
                    return sheet
        return None

    async def _existing_workspace_slugs(
        self,
        session: AsyncSession,
        *,
        exclude_workspace_id: str | None = None,
    ) -> set[str]:
        return await workspace_repository.existing_slugs(
            session,
            exclude_workspace_id=exclude_workspace_id,
        )

    async def _next_workspace_position(self, session: AsyncSession, user_id: str) -> int:
        return await workspace_repository.next_position(session, user_id)

    def _append_activity(
        self,
        session: AsyncSession,
        *,
        action_type: str,
        workspace_id: str | None,
        workbook_id: str | None = None,
        sheet_id: str | None = None,
        record_id: str | None = None,
        user_id: str | None = None,
        payload: dict[str, object] | None = None,
    ) -> None:
        session.add(
            ActivityLogModel(
                id=f"log_{uuid4().hex[:8]}",
                workspace_id=workspace_id,
                workbook_id=workbook_id,
                sheet_id=sheet_id,
                record_id=record_id,
                user_id=user_id,
                action_type=action_type,
                payload_json=deepcopy(payload or {}),
                created_at=timestamp(),
            )
        )

    def _validation_error(self, detail: str) -> HTTPException:
        return HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=detail)

    def _not_found(self, entity: str) -> HTTPException:
        return HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=f"{entity} not found.")


workspace_service = WorkspaceService()
