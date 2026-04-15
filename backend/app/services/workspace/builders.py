from __future__ import annotations

from copy import deepcopy
from uuid import uuid4

from app.models import (
    SheetColumnModel,
    SheetModel,
    SheetRecordModel,
    SheetViewModel,
    WorkbookModel,
    WorkspaceMemberModel,
    WorkspaceModel,
)
from app.services.workspace.constants import (
    DEFAULT_STARTER_COLUMN_COUNT,
    DEFAULT_STARTER_COLUMN_WIDTH,
    DEFAULT_VIEW_CONFIG,
)
from app.services.workspace.fixtures import SEED_WORKSPACE_FIXTURES
from app.services.workspace.naming import slugify, timestamp
from app.services.workspace.ordering import sort_columns, sort_records, sort_sheets, sort_workbooks


def default_column_key(index: int) -> str:
    return f"column_{index}"


def default_column_title(index: int) -> str:
    return f"Column {index}"


INITIAL_SHEET_PLACEHOLDER_ROW_BUDGET = 50
PLACEHOLDER_ROW_ID_PREFIX = "__datagrid_placeholder__:"


def build_default_column_blueprints() -> list[dict[str, object]]:
    return [
        {
            "key": default_column_key(index),
            "title": default_column_title(index),
            "type": "text",
            "width": DEFAULT_STARTER_COLUMN_WIDTH,
            "nullable": True,
            "editable": True,
            "computed": False,
            "expression": None,
            "settings_json": {},
        }
        for index in range(1, DEFAULT_STARTER_COLUMN_COUNT + 1)
    ]


def infer_blueprint_from_key(key: str) -> dict[str, object]:
    normalized_key = key.strip().lower()
    title = " ".join(part for part in normalized_key.replace("-", "_").split("_") if part).title() or key
    blueprint = {
        "key": key,
        "title": title,
        "type": "text",
        "width": DEFAULT_STARTER_COLUMN_WIDTH,
        "nullable": True,
        "editable": True,
        "computed": False,
        "expression": None,
        "settings_json": {},
    }

    if normalized_key == "task":
        blueprint["title"] = "Task"
        blueprint["width"] = 280
    elif normalized_key == "owner":
        blueprint["title"] = "Owner"
        blueprint["type"] = "user"
        blueprint["width"] = 140
    elif normalized_key == "status":
        blueprint["title"] = "Status"
        blueprint["type"] = "status"
        blueprint["width"] = 140
        blueprint["settings_json"] = {
            "options": ["Draft", "Ready", "In review", "In progress", "Blocked", "Pending"]
        }
    elif normalized_key == "timeline":
        blueprint["title"] = "Timeline"
        blueprint["type"] = "date"
        blueprint["width"] = 140
    elif normalized_key == "progress":
        blueprint["title"] = "Progress"
        blueprint["type"] = "percent"
        blueprint["width"] = 120
        blueprint["settings_json"] = {"min": 0, "max": 100}

    return blueprint


def infer_column_blueprints(rows: list[dict[str, object]] | None) -> list[dict[str, object]]:
    if not rows:
        return build_default_column_blueprints()

    ordered_keys: list[str] = []
    seen_keys: set[str] = set()

    for row in rows:
        for key in row:
            normalized_key = str(key)
            if normalized_key == "id" or normalized_key in seen_keys:
                continue
            seen_keys.add(normalized_key)
            ordered_keys.append(normalized_key)

    if not ordered_keys:
        return build_default_column_blueprints()

    return [infer_blueprint_from_key(key) for key in ordered_keys]


def build_columns(sheet_id: str, rows: list[dict[str, object]] | None = None) -> list[SheetColumnModel]:
    created_at = timestamp()
    return [
        SheetColumnModel(
            id=f"col_{uuid4().hex[:8]}",
            sheet_id=sheet_id,
            key=blueprint["key"],
            title=blueprint["title"],
            type=blueprint["type"],
            position=position,
            width=blueprint["width"],
            nullable=blueprint["nullable"],
            editable=blueprint["editable"],
            computed=blueprint["computed"],
            expression=blueprint["expression"],
            settings_json=deepcopy(blueprint["settings_json"]),
            created_at=created_at,
            updated_at=created_at,
        )
        for position, blueprint in enumerate(infer_column_blueprints(rows))
    ]


def grid_data_type_to_column_type(value: str) -> str:
    if value == "number":
        return "number"
    if value == "date":
        return "date"
    if value == "status":
        return "status"
    if value == "currency":
        return "currency"
    return "text"


def build_columns_from_grid_input(
    *,
    sheet_id: str,
    columns: list[dict[str, object]],
) -> list[SheetColumnModel]:
    created_at = timestamp()
    built_columns: list[SheetColumnModel] = []

    for position, column in enumerate(columns):
        settings = deepcopy(column.get("settings") if isinstance(column.get("settings"), dict) else {})
        options = column.get("options")
        if isinstance(options, list):
            normalized_options = [str(option).strip() for option in options if str(option).strip()]
            if normalized_options:
                settings["options"] = normalized_options
            else:
                settings.pop("options", None)

        width = column.get("width")
        resolved_width = int(width) if isinstance(width, int) else None
        data_type = str(column.get("data_type") or "text")
        column_type = str(column.get("column_type") or grid_data_type_to_column_type(data_type))
        expression = str(column.get("expression") or "").strip()
        computed = bool(column.get("computed", False) or expression or column_type == "formula")

        built_columns.append(
            SheetColumnModel(
                id=f"col_{uuid4().hex[:8]}",
                sheet_id=sheet_id,
                key=str(column["key"]),
                title=str(column["label"]),
                type=column_type,
                position=position,
                width=resolved_width,
                nullable=True,
                editable=False if computed else bool(column.get("editable", True)),
                computed=computed,
                expression=expression or None,
                settings_json=settings,
                created_at=created_at,
                updated_at=created_at,
            )
        )

    return built_columns


def build_records(
    *,
    sheet_id: str,
    user_id: str,
    rows: list[dict[str, object]],
) -> list[SheetRecordModel]:
    created_at = timestamp()
    records: list[SheetRecordModel] = []

    for position, row in enumerate(rows):
        row_values = deepcopy(row)
        record_id = str(row_values.pop("id", f"rec_{uuid4().hex[:8]}"))
        records.append(
            SheetRecordModel(
                id=record_id,
                sheet_id=sheet_id,
                position=position,
                values_json=row_values,
                display_values_json=None,
                created_by=user_id,
                updated_by=user_id,
                archived=False,
                created_at=created_at,
                updated_at=created_at,
            )
        )

    return records


def normalize_grid_rows(
    *,
    rows: list[dict[str, object]],
    column_keys: list[str],
) -> list[dict[str, object]]:
    normalized_rows: list[dict[str, object]] = []

    for row in rows:
        raw_row_id = str(row.get("id") or "").strip()
        row_payload: dict[str, object] = {
            "id": (
                f"rec_{uuid4().hex[:8]}"
                if not raw_row_id or raw_row_id.startswith(PLACEHOLDER_ROW_ID_PREFIX)
                else raw_row_id
            )
        }

        for column_key in column_keys:
            if column_key in row:
                row_payload[column_key] = deepcopy(row[column_key])

        if not row_has_meaningful_values(row_payload, column_keys):
            continue

        normalized_rows.append(row_payload)

    return normalized_rows


def row_has_meaningful_values(row: dict[str, object], column_keys: list[str]) -> bool:
    for column_key in column_keys:
        value = row.get(column_key)

        if value is None:
            continue

        if isinstance(value, str) and not value.strip():
            continue

        if isinstance(value, (list, dict)) and not value:
            continue

        return True

    return False


def build_default_view(sheet_id: str, user_id: str) -> SheetViewModel:
    created_at = timestamp()
    return SheetViewModel(
        id=f"view_{uuid4().hex[:8]}",
        sheet_id=sheet_id,
        key="grid_default",
        name="Grid",
        type="grid",
        is_default=True,
        created_by=user_id,
        config_json=deepcopy(DEFAULT_VIEW_CONFIG),
        created_at=created_at,
        updated_at=created_at,
    )


def build_sheet(
    *,
    workbook_id: str,
    user_id: str,
    name: str,
    position: int,
    rows: list[dict[str, object]] | None = None,
    key: str | None = None,
    kind: str = "data",
    source_sheet_id: str | None = None,
    config_json: dict[str, object] | None = None,
    sheet_id: str | None = None,
) -> SheetModel:
    created_at = timestamp()
    resolved_sheet_id = sheet_id or f"sheet_{uuid4().hex[:8]}"
    resolved_rows = rows if rows is not None else []
    resolved_config = deepcopy(config_json or {})
    if config_json is None and rows is None:
        resolved_config["initial_placeholder_row_budget"] = INITIAL_SHEET_PLACEHOLDER_ROW_BUDGET

    sheet = SheetModel(
        id=resolved_sheet_id,
        workbook_id=workbook_id,
        key=key or slugify(name).replace("-", "_"),
        name=name,
        kind=kind,
        source_sheet_id=source_sheet_id,
        position=position,
        config_json=resolved_config,
        styles_json=[],
        created_at=created_at,
        updated_at=created_at,
    )
    sheet.columns = build_columns(resolved_sheet_id, resolved_rows)
    sheet.records = build_records(sheet_id=resolved_sheet_id, user_id=user_id, rows=resolved_rows)
    sheet.views = [build_default_view(resolved_sheet_id, user_id)]
    return sheet


def build_workbook(
    *,
    workspace_id: str,
    user_id: str,
    name: str,
    description: str | None,
    position: int,
    sheets: list[SheetModel],
    workbook_id: str | None = None,
) -> WorkbookModel:
    created_at = timestamp()
    return WorkbookModel(
        id=workbook_id or f"wb_{uuid4().hex[:8]}",
        workspace_id=workspace_id,
        name=name,
        description=description,
        position=position,
        created_by=user_id,
        updated_by=user_id,
        archived=False,
        created_at=created_at,
        updated_at=created_at,
        sheets=sheets,
    )


def clone_sheet(
    source: SheetModel,
    *,
    workbook_id: str,
    position: int,
    name: str | None = None,
    key: str | None = None,
) -> SheetModel:
    created_at = timestamp()
    sheet = SheetModel(
        id=f"sheet_{uuid4().hex[:8]}",
        workbook_id=workbook_id,
        key=key or source.key,
        name=name or source.name,
        kind=source.kind,
        source_sheet_id=None,
        position=position,
        config_json=deepcopy(source.config_json or {}),
        styles_json=deepcopy(source.styles_json or []),
        created_at=created_at,
        updated_at=created_at,
    )
    sheet.columns = [
        SheetColumnModel(
            id=f"col_{uuid4().hex[:8]}",
            sheet_id=sheet.id,
            key=column.key,
            title=column.title,
            type=column.type,
            position=index,
            width=column.width,
            nullable=column.nullable,
            editable=column.editable,
            computed=column.computed,
            expression=column.expression,
            settings_json=deepcopy(column.settings_json or {}),
            created_at=created_at,
            updated_at=created_at,
        )
        for index, column in enumerate(sort_columns(source.columns))
    ]
    sheet.records = [
        SheetRecordModel(
            id=f"rec_{uuid4().hex[:8]}",
            sheet_id=sheet.id,
            position=index,
            values_json=deepcopy(record.values_json or {}),
            display_values_json=deepcopy(record.display_values_json) if record.display_values_json else None,
            created_by=record.created_by,
            updated_by=record.updated_by,
            archived=record.archived,
            created_at=created_at,
            updated_at=created_at,
        )
        for index, record in enumerate(sort_records(source.records))
    ]
    sheet.views = [
        SheetViewModel(
            id=f"view_{uuid4().hex[:8]}",
            sheet_id=sheet.id,
            key=view.key,
            name=view.name,
            type=view.type,
            is_default=view.is_default,
            created_by=view.created_by,
            config_json=deepcopy(view.config_json or {}),
            created_at=created_at,
            updated_at=created_at,
        )
        for view in source.views
    ]
    return sheet


def clone_workbook(
    source: WorkbookModel,
    *,
    workspace_id: str,
    position: int,
) -> WorkbookModel:
    created_at = timestamp()
    workbook = WorkbookModel(
        id=f"wb_{uuid4().hex[:8]}",
        workspace_id=workspace_id,
        name=source.name,
        description=source.description,
        position=position,
        created_by=source.created_by,
        updated_by=source.updated_by,
        archived=source.archived,
        created_at=created_at,
        updated_at=created_at,
    )
    workbook.sheets = [
        clone_sheet(sheet, workbook_id=workbook.id, position=index)
        for index, sheet in enumerate(sort_sheets(source.sheets))
    ]
    return workbook


def build_seed_workspaces(owner_id: str) -> list[WorkspaceModel]:
    seed_now = timestamp()
    workspaces: list[WorkspaceModel] = []

    for workspace_fixture in SEED_WORKSPACE_FIXTURES:
        workspace = WorkspaceModel(
            id=str(workspace_fixture["id"]),
            name=str(workspace_fixture["name"]),
            slug=str(workspace_fixture["slug"]),
            owner_id=owner_id,
            description=str(workspace_fixture["description"]),
            color=str(workspace_fixture["color"]),
            position=int(workspace_fixture["position"]),
            created_at=seed_now,
            updated_at=seed_now,
        )
        workspace.members = [
            WorkspaceMemberModel(
                id=str(workspace_fixture["member_id"]),
                workspace_id=workspace.id,
                user_id=owner_id,
                role="owner",
                created_at=seed_now,
            )
        ]

        workbook_fixture = workspace_fixture["workbook"]
        workbook_sheets = [
            build_sheet(
                workbook_id=str(workbook_fixture["id"]),
                user_id=owner_id,
                name=str(sheet_fixture["name"]),
                key=str(sheet_fixture["key"]),
                position=int(sheet_fixture["position"]),
                rows=deepcopy(sheet_fixture["rows"]),
                sheet_id=str(sheet_fixture["id"]),
            )
            for sheet_fixture in workbook_fixture["sheets"]
        ]
        workspace.workbooks = [
            build_workbook(
                workspace_id=workspace.id,
                user_id=owner_id,
                name=str(workbook_fixture["name"]),
                description=str(workbook_fixture["description"]),
                position=int(workbook_fixture["position"]),
                workbook_id=str(workbook_fixture["id"]),
                sheets=workbook_sheets,
            )
        ]
        workspaces.append(workspace)

    return workspaces
