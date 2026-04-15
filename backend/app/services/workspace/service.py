from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from uuid import uuid4

from fastapi import HTTPException, status
from sqlalchemy import func, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models import (
    ActivityLogModel,
    SheetCellRevisionModel,
    SheetColumnModel,
    SheetModel,
    SheetRecordModel,
    SheetRevisionModel,
    UserModel,
    WorkbookModel,
    WorkspaceMemberModel,
    WorkspaceModel,
)
from app.repositories import workspace_repository
from app.services.workspace.builders import (
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
from app.services.workspace.validation import normalize_and_validate_formula_alias


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
        styles: list[dict[str, object]] | None = None,
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
            settings = (
                deepcopy(column.get("settings"))
                if isinstance(column.get("settings"), dict)
                else {}
            )
            raw_formula_alias = column.get(
                "formula_alias",
                settings.get("formulaAlias", settings.get("formula_alias")),
            )
            try:
                formula_alias = normalize_and_validate_formula_alias(raw_formula_alias)
            except ValueError as error:
                raise self._validation_error(str(error)) from error

            if formula_alias:
                settings["formulaAlias"] = formula_alias
            else:
                settings.pop("formulaAlias", None)
                settings.pop("formula_alias", None)
            normalized_columns.append(
                {
                    "key": key,
                    "label": label,
                    "formula_alias": formula_alias,
                    "data_type": str(column.get("data_type") or "text"),
                    "column_type": str(column.get("column_type") or ("formula" if computed else "")),
                    "width": column.get("width"),
                    "editable": False if computed else bool(column.get("editable", True)),
                    "computed": computed,
                    "expression": expression or None,
                    "options": column.get("options") if isinstance(column.get("options"), list) else [],
                    "settings": settings,
                }
            )

        sheet = await self._require_sheet(session, user_id, workspace_id, sheet_id, include_grid=True)
        updated_at = timestamp()
        previous_columns = self._build_activity_columns_snapshot(sheet.columns)
        previous_row_ids = [record.id for record in sorted(sheet.records, key=lambda record: record.position) if not record.archived]
        previous_styles = deepcopy(sheet.styles_json or [])
        writable_column_keys = [
            str(column["key"]) for column in normalized_columns if not bool(column.get("computed"))
        ]
        normalized_rows = normalize_grid_rows(rows=rows, column_keys=writable_column_keys)
        normalized_styles = deepcopy(styles) if styles is not None else None
        workbook = await self._touch_sheet_parents(
            session,
            user_id=user_id,
            workspace_id=workspace_id,
            workbook_id=sheet.workbook_id,
            updated_at=updated_at,
        )

        current_columns_by_key = {column.key: column for column in sheet.columns}
        current_records_by_id = {
            record.id: record for record in sheet.records if not record.archived
        }
        revision_id = f"rev_{uuid4().hex[:8]}"
        cell_revisions = self._build_sheet_cell_revisions(
            sheet_id=sheet.id,
            revision_id=revision_id,
            writable_column_keys=writable_column_keys,
            current_records_by_id=current_records_by_id,
            next_rows=normalized_rows,
            created_at=updated_at,
        )
        styles_changed = normalized_styles is not None and previous_styles != normalized_styles
        activity_events = self._build_sheet_grid_activity_events(
            previous_columns=previous_columns,
            next_columns=normalized_columns,
            previous_row_ids=previous_row_ids,
            next_rows=normalized_rows,
            cell_revisions=cell_revisions,
            previous_styles=previous_styles,
            next_styles=normalized_styles,
        )

        self._sync_sheet_columns(
            session,
            sheet=sheet,
            normalized_columns=normalized_columns,
            existing_columns_by_key=current_columns_by_key,
            updated_at=updated_at,
        )
        self._sync_sheet_records(
            session,
            sheet=sheet,
            user_id=user_id,
            normalized_rows=normalized_rows,
            existing_records_by_id=current_records_by_id,
            updated_at=updated_at,
        )
        if normalized_styles is not None:
            sheet.styles_json = normalized_styles
        sheet.updated_at = updated_at
        if cell_revisions or styles_changed:
            self._append_sheet_revision(
                session,
                revision_id=revision_id,
                workspace_id=workspace_id,
                workbook_id=sheet.workbook_id,
                sheet_id=sheet.id,
                user_id=user_id,
                created_at=updated_at,
                payload={
                    "cellChangeCount": len(cell_revisions),
                    "rowCount": len(normalized_rows),
                    "styleRuleCount": len(sheet.styles_json or []),
                    "stylesChanged": styles_changed,
                },
            )
            if cell_revisions:
                await session.flush()
                session.add_all(cell_revisions)
        for action_type, payload in activity_events:
            self._append_activity(
                session,
                workspace_id=workspace_id,
                workbook_id=sheet.workbook_id,
                sheet_id=sheet.id,
                user_id=user_id,
                action_type=action_type,
                payload=payload,
            )
        workbook.updated_at = updated_at
        await session.commit()

        reloaded_sheet = await self._require_sheet(session, user_id, workspace_id, sheet.id, include_grid=True)
        return serialize_sheet(reloaded_sheet, include_grid=True)

    async def get_sheet_cell_history(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
        *,
        record_id: str,
        column_key: str,
    ) -> dict[str, object]:
        await self._require_sheet(session, user_id, workspace_id, sheet_id, include_grid=False)
        history_result = await session.execute(
            select(SheetCellRevisionModel, SheetRevisionModel, UserModel)
            .join(SheetRevisionModel, SheetRevisionModel.id == SheetCellRevisionModel.revision_id)
            .outerjoin(UserModel, UserModel.id == SheetRevisionModel.user_id)
            .where(
                SheetCellRevisionModel.sheet_id == sheet_id,
                SheetCellRevisionModel.record_id == record_id,
                SheetCellRevisionModel.column_key == column_key,
            )
            .order_by(SheetCellRevisionModel.created_at.desc(), SheetCellRevisionModel.id.desc())
        )

        items = []
        for cell_revision, revision, actor in history_result.all():
            items.append(
                {
                    "id": cell_revision.id,
                    "revision_id": revision.id,
                    "record_id": cell_revision.record_id,
                    "column_key": cell_revision.column_key,
                    "previous_value": deepcopy(cell_revision.previous_value_json),
                    "next_value": deepcopy(cell_revision.next_value_json),
                    "changed_at": cell_revision.created_at.isoformat(),
                    "actor": None
                    if actor is None
                    else {
                        "id": actor.id,
                        "email": actor.email,
                        "full_name": actor.full_name,
                        "avatar_url": actor.avatar_url,
                    },
                }
            )

        return {"items": items}

    async def get_sheet_activity(
        self,
        session: AsyncSession,
        user_id: str,
        workspace_id: str,
        sheet_id: str,
        *,
        created_from: datetime | None = None,
        created_to: datetime | None = None,
        action_types: list[str] | None = None,
        user_ids: list[str] | None = None,
        limit: int = 200,
    ) -> dict[str, object]:
        await self._require_sheet(session, user_id, workspace_id, sheet_id, include_grid=False)

        scope_conditions = [
            ActivityLogModel.workspace_id == workspace_id,
            ActivityLogModel.sheet_id == sheet_id,
        ]
        filtered_conditions = list(scope_conditions)

        normalized_action_types = sorted({value.strip() for value in (action_types or []) if value.strip()})
        normalized_user_ids = sorted({value.strip() for value in (user_ids or []) if value.strip()})
        if created_from is not None:
            filtered_conditions.append(ActivityLogModel.created_at >= created_from)
        if created_to is not None:
            filtered_conditions.append(ActivityLogModel.created_at <= created_to)
        if normalized_action_types:
            filtered_conditions.append(ActivityLogModel.action_type.in_(normalized_action_types))
        if normalized_user_ids:
            filtered_conditions.append(ActivityLogModel.user_id.in_(normalized_user_ids))

        activity_result = await session.execute(
            select(ActivityLogModel, UserModel)
            .outerjoin(UserModel, UserModel.id == ActivityLogModel.user_id)
            .where(*filtered_conditions)
            .order_by(ActivityLogModel.created_at.desc(), ActivityLogModel.id.desc())
            .limit(limit)
        )
        action_result = await session.execute(
            select(ActivityLogModel.action_type, func.count(ActivityLogModel.id))
            .where(*scope_conditions)
            .group_by(ActivityLogModel.action_type)
            .order_by(func.count(ActivityLogModel.id).desc(), ActivityLogModel.action_type.asc())
        )
        collaborator_result = await session.execute(
            select(
                UserModel.id,
                UserModel.email,
                UserModel.full_name,
                UserModel.avatar_url,
                func.count(ActivityLogModel.id),
            )
            .join(UserModel, UserModel.id == ActivityLogModel.user_id)
            .where(*scope_conditions, ActivityLogModel.user_id.is_not(None))
            .group_by(UserModel.id, UserModel.email, UserModel.full_name, UserModel.avatar_url)
            .order_by(
                func.count(ActivityLogModel.id).desc(),
                UserModel.full_name.asc(),
                UserModel.email.asc(),
            )
        )

        items = [
            {
                "id": activity.id,
                "action_type": activity.action_type,
                "workbook_id": activity.workbook_id,
                "sheet_id": activity.sheet_id,
                "record_id": activity.record_id,
                "payload": deepcopy(activity.payload_json or {}),
                "created_at": activity.created_at.isoformat(),
                "actor": None
                if actor is None
                else {
                    "id": actor.id,
                    "email": actor.email,
                    "full_name": actor.full_name,
                    "avatar_url": actor.avatar_url,
                },
            }
            for activity, actor in activity_result.all()
        ]
        actions = [
            {
                "action_type": action_type,
                "count": count,
            }
            for action_type, count in action_result.all()
        ]
        collaborators = [
            {
                "actor": {
                    "id": collaborator_id,
                    "email": email,
                    "full_name": full_name,
                    "avatar_url": avatar_url,
                },
                "count": count,
            }
            for collaborator_id, email, full_name, avatar_url, count in collaborator_result.all()
        ]

        return {
            "items": items,
            "actions": actions,
            "collaborators": collaborators,
        }

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

    def _append_sheet_revision(
        self,
        session: AsyncSession,
        *,
        revision_id: str,
        workspace_id: str,
        workbook_id: str | None,
        sheet_id: str,
        user_id: str | None,
        created_at,
        payload: dict[str, object] | None = None,
    ) -> None:
        session.add(
            SheetRevisionModel(
                id=revision_id,
                workspace_id=workspace_id,
                workbook_id=workbook_id,
                sheet_id=sheet_id,
                user_id=user_id,
                action_type="sheet_grid_updated",
                payload_json=deepcopy(payload or {}),
                created_at=created_at,
            )
        )

    def _build_activity_columns_snapshot(
        self,
        columns: list[SheetColumnModel],
    ) -> list[dict[str, object]]:
        snapshot: list[dict[str, object]] = []

        for column in sorted(columns, key=lambda item: item.position):
            settings = deepcopy(column.settings_json or {})
            raw_formula_alias = settings.get("formulaAlias", settings.get("formula_alias"))
            try:
                formula_alias = normalize_and_validate_formula_alias(raw_formula_alias)
            except ValueError:
                formula_alias = str(raw_formula_alias).strip() or None if raw_formula_alias is not None else None

            raw_options = settings.get("options")
            options = [str(option).strip() for option in raw_options if str(option).strip()] if isinstance(raw_options, list) else []
            snapshot.append(
                {
                    "key": column.key,
                    "label": column.title,
                    "formula_alias": formula_alias,
                    "column_type": column.type,
                    "editable": column.editable,
                    "computed": column.computed,
                    "expression": column.expression,
                    "options": options,
                }
            )

        return snapshot

    def _build_sheet_grid_activity_events(
        self,
        *,
        previous_columns: list[dict[str, object]],
        next_columns: list[dict[str, object]],
        previous_row_ids: list[str],
        next_rows: list[dict[str, object]],
        cell_revisions: list[SheetCellRevisionModel],
        previous_styles: list[dict[str, object]],
        next_styles: list[dict[str, object]] | None,
    ) -> list[tuple[str, dict[str, object]]]:
        events: list[tuple[str, dict[str, object]]] = []
        row_payload = self._build_sheet_row_change_activity_payload(
            previous_row_ids=previous_row_ids,
            next_rows=next_rows,
        )
        added_row_ids = set(row_payload.get("insertedRowIds", [])) if row_payload else set()
        removed_row_ids = set(row_payload.get("deletedRowIds", [])) if row_payload else set()
        column_payload = self._build_sheet_column_change_activity_payload(
            previous_columns=previous_columns,
            next_columns=next_columns,
        )
        cell_payload = self._build_sheet_cell_change_activity_payload(
            cell_revisions=cell_revisions,
            added_row_ids=added_row_ids,
            removed_row_ids=removed_row_ids,
            previous_columns=previous_columns,
            next_columns=next_columns,
            next_rows=next_rows,
        )
        style_payload = self._build_sheet_style_change_activity_payload(
            previous_styles=previous_styles,
            next_styles=next_styles,
        )

        if column_payload is not None:
            events.append(("sheet_columns_changed", column_payload))
        if row_payload is not None:
            events.append(("sheet_rows_changed", row_payload))
        if cell_payload is not None:
            events.append(("sheet_cells_changed", cell_payload))
        if style_payload is not None:
            events.append(("sheet_styles_changed", style_payload))

        return events

    def _build_sheet_column_change_activity_payload(
        self,
        *,
        previous_columns: list[dict[str, object]],
        next_columns: list[dict[str, object]],
    ) -> dict[str, object] | None:
        previous_keys = [str(column["key"]) for column in previous_columns]
        next_keys = [str(column["key"]) for column in next_columns]
        previous_by_key = {str(column["key"]): column for column in previous_columns}
        next_by_key = {str(column["key"]): column for column in next_columns}

        added_keys = [column_key for column_key in next_keys if column_key not in previous_by_key]
        removed_keys = [column_key for column_key in previous_keys if column_key not in next_by_key]
        common_keys = [column_key for column_key in previous_keys if column_key in next_by_key]
        renamed_columns = [
            {
                "key": column_key,
                "previousLabel": previous_by_key[column_key].get("label"),
                "nextLabel": next_by_key[column_key].get("label"),
            }
            for column_key in common_keys
            if previous_by_key[column_key].get("label") != next_by_key[column_key].get("label")
        ]
        alias_changed_columns = [
            {
                "key": column_key,
                "previousAlias": previous_by_key[column_key].get("formula_alias"),
                "nextAlias": next_by_key[column_key].get("formula_alias"),
            }
            for column_key in common_keys
            if previous_by_key[column_key].get("formula_alias") != next_by_key[column_key].get("formula_alias")
        ]
        reconfigured_columns = [
            column_key
            for column_key in common_keys
            if any(
                previous_by_key[column_key].get(field_name) != next_by_key[column_key].get(field_name)
                for field_name in ("column_type", "editable", "computed", "expression", "options")
            )
        ]
        reordered = common_keys != [column_key for column_key in next_keys if column_key in previous_by_key]

        if not (added_keys or removed_keys or renamed_columns or alias_changed_columns or reconfigured_columns or reordered):
            return None

        return {
            "addedCount": len(added_keys),
            "removedCount": len(removed_keys),
            "renamedCount": len(renamed_columns),
            "aliasChangedCount": len(alias_changed_columns),
            "reconfiguredCount": len(reconfigured_columns),
            "reordered": reordered,
            "addedColumns": added_keys[:5],
            "removedColumns": removed_keys[:5],
            "renamedColumns": renamed_columns[:5],
            "aliasChangedColumns": alias_changed_columns[:5],
            "reconfiguredColumns": reconfigured_columns[:5],
        }

    def _build_sheet_row_change_activity_payload(
        self,
        *,
        previous_row_ids: list[str],
        next_rows: list[dict[str, object]],
    ) -> dict[str, object] | None:
        next_row_ids = [str(row.get("id")) for row in next_rows]
        previous_row_id_set = set(previous_row_ids)
        next_row_id_set = set(next_row_ids)
        inserted_row_ids = [row_id for row_id in next_row_ids if row_id not in previous_row_id_set]
        deleted_row_ids = [row_id for row_id in previous_row_ids if row_id not in next_row_id_set]
        inserted_row_numbers = [index + 1 for index, row_id in enumerate(next_row_ids) if row_id not in previous_row_id_set]
        deleted_row_numbers = [index + 1 for index, row_id in enumerate(previous_row_ids) if row_id not in next_row_id_set]
        previous_common_row_ids = [row_id for row_id in previous_row_ids if row_id in next_row_id_set]
        next_common_row_ids = [row_id for row_id in next_row_ids if row_id in previous_row_id_set]
        reordered = previous_common_row_ids != next_common_row_ids

        if not (inserted_row_ids or deleted_row_ids or reordered):
            return None

        return {
            "insertedCount": len(inserted_row_ids),
            "deletedCount": len(deleted_row_ids),
            "reordered": reordered,
            "insertedRowNumbers": inserted_row_numbers[:5],
            "deletedRowNumbers": deleted_row_numbers[:5],
            "insertedRowIds": inserted_row_ids[:5],
            "deletedRowIds": deleted_row_ids[:5],
        }

    def _build_sheet_cell_change_activity_payload(
        self,
        *,
        cell_revisions: list[SheetCellRevisionModel],
        added_row_ids: set[str],
        removed_row_ids: set[str],
        previous_columns: list[dict[str, object]],
        next_columns: list[dict[str, object]],
        next_rows: list[dict[str, object]],
    ) -> dict[str, object] | None:
        relevant_revisions = [
            revision
            for revision in cell_revisions
            if revision.record_id not in added_row_ids and revision.record_id not in removed_row_ids
        ]
        if not relevant_revisions:
            return None

        affected_row_ids = sorted({revision.record_id for revision in relevant_revisions})
        affected_column_keys = sorted({revision.column_key for revision in relevant_revisions})
        row_number_by_id = {
            str(row.get("id")): index + 1
            for index, row in enumerate(next_rows)
            if row.get("id") is not None
        }
        column_label_by_key: dict[str, str] = {}
        for column in previous_columns:
            column_key = str(column.get("key") or "").strip()
            column_label = str(column.get("label") or column_key).strip()
            if column_key:
                column_label_by_key[column_key] = column_label or column_key
        for column in next_columns:
            column_key = str(column.get("key") or "").strip()
            column_label = str(column.get("label") or column_key).strip()
            if column_key:
                column_label_by_key[column_key] = column_label or column_key

        return {
            "changedCellCount": len(relevant_revisions),
            "affectedRowCount": len(affected_row_ids),
            "affectedColumnCount": len(affected_column_keys),
            "affectedRowIds": affected_row_ids[:5],
            "affectedColumns": affected_column_keys[:5],
            "affectedColumnLabels": [column_label_by_key.get(column_key, column_key) for column_key in affected_column_keys[:5]],
            "sampleChanges": [
                {
                    "recordId": revision.record_id,
                    "rowNumber": row_number_by_id.get(revision.record_id),
                    "columnKey": revision.column_key,
                    "columnLabel": column_label_by_key.get(revision.column_key, revision.column_key),
                    "previousValue": deepcopy(revision.previous_value_json),
                    "nextValue": deepcopy(revision.next_value_json),
                }
                for revision in relevant_revisions[:5]
            ],
        }

    def _build_sheet_style_change_activity_payload(
        self,
        *,
        previous_styles: list[dict[str, object]],
        next_styles: list[dict[str, object]] | None,
    ) -> dict[str, object] | None:
        if next_styles is None or previous_styles == next_styles:
            return None

        return {
            "previousRuleCount": len(previous_styles),
            "nextRuleCount": len(next_styles),
        }

    def _build_sheet_cell_revisions(
        self,
        *,
        sheet_id: str,
        revision_id: str,
        writable_column_keys: list[str],
        current_records_by_id: dict[str, SheetRecordModel],
        next_rows: list[dict[str, object]],
        created_at,
    ) -> list[SheetCellRevisionModel]:
        next_rows_by_id = {str(row.get("id")): row for row in next_rows}
        record_ids = set(current_records_by_id) | set(next_rows_by_id)
        revisions: list[SheetCellRevisionModel] = []

        for record_id in record_ids:
            previous_values = deepcopy(current_records_by_id.get(record_id).values_json or {}) if record_id in current_records_by_id else {}
            next_values = deepcopy(next_rows_by_id.get(record_id) or {})

            for column_key in writable_column_keys:
                previous_value = deepcopy(previous_values.get(column_key))
                next_value = deepcopy(next_values.get(column_key))
                if previous_value == next_value:
                    continue

                revisions.append(
                    SheetCellRevisionModel(
                        id=f"cellrev_{uuid4().hex[:8]}",
                        revision_id=revision_id,
                        sheet_id=sheet_id,
                        record_id=record_id,
                        column_key=column_key,
                        previous_value_json=previous_value,
                        next_value_json=next_value,
                        created_at=created_at,
                    )
                )

        return revisions

    def _sync_sheet_columns(
        self,
        session: AsyncSession,
        *,
        sheet: SheetModel,
        normalized_columns: list[dict[str, object]],
        existing_columns_by_key: dict[str, SheetColumnModel],
        updated_at,
    ) -> None:
        next_column_keys = {str(column["key"]) for column in normalized_columns}

        for column in list(sheet.columns):
            if column.key not in next_column_keys:
                session.delete(column)

        for position, normalized_column in enumerate(normalized_columns):
            column_key = str(normalized_column["key"])
            existing_column = existing_columns_by_key.get(column_key)
            settings = deepcopy(
                normalized_column.get("settings")
                if isinstance(normalized_column.get("settings"), dict)
                else {}
            )
            options = normalized_column.get("options")
            if isinstance(options, list):
                normalized_options = [str(option).strip() for option in options if str(option).strip()]
                if normalized_options:
                    settings["options"] = normalized_options
                else:
                    settings.pop("options", None)

            width = normalized_column.get("width")
            resolved_width = int(width) if isinstance(width, int) else None

            if existing_column is None:
                sheet.columns.append(
                    SheetColumnModel(
                        id=f"col_{uuid4().hex[:8]}",
                        sheet_id=sheet.id,
                        key=column_key,
                        title=str(normalized_column["label"]),
                        type=str(normalized_column["column_type"]),
                        position=position,
                        width=resolved_width,
                        nullable=True,
                        editable=bool(normalized_column["editable"]),
                        computed=bool(normalized_column["computed"]),
                        expression=normalized_column.get("expression") or None,
                        settings_json=settings,
                        created_at=updated_at,
                        updated_at=updated_at,
                    )
                )
                continue

            existing_column.title = str(normalized_column["label"])
            existing_column.type = str(normalized_column["column_type"])
            existing_column.position = position
            existing_column.width = resolved_width
            existing_column.editable = bool(normalized_column["editable"])
            existing_column.computed = bool(normalized_column["computed"])
            existing_column.expression = normalized_column.get("expression") or None
            existing_column.settings_json = settings
            existing_column.updated_at = updated_at

    def _sync_sheet_records(
        self,
        session: AsyncSession,
        *,
        sheet: SheetModel,
        user_id: str,
        normalized_rows: list[dict[str, object]],
        existing_records_by_id: dict[str, SheetRecordModel],
        updated_at,
    ) -> None:
        next_record_ids = {str(row.get("id")) for row in normalized_rows}

        for record in list(sheet.records):
            if record.id not in next_record_ids:
                session.delete(record)

        for position, normalized_row in enumerate(normalized_rows):
            record_id = str(normalized_row.get("id"))
            row_values = deepcopy(normalized_row)
            row_values.pop("id", None)
            existing_record = existing_records_by_id.get(record_id)

            if existing_record is None:
                sheet.records.append(
                    SheetRecordModel(
                        id=record_id,
                        sheet_id=sheet.id,
                        position=position,
                        values_json=row_values,
                        display_values_json=None,
                        created_by=user_id,
                        updated_by=user_id,
                        archived=False,
                        created_at=updated_at,
                        updated_at=updated_at,
                    )
                )
                continue

            existing_record.position = position
            existing_record.values_json = row_values
            existing_record.updated_by = user_id
            existing_record.updated_at = updated_at
            existing_record.archived = False

    def _validation_error(self, detail: str) -> HTTPException:
        return HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=detail)

    def _not_found(self, entity: str) -> HTTPException:
        return HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=f"{entity} not found.")


workspace_service = WorkspaceService()
