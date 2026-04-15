from __future__ import annotations

from copy import deepcopy

from app.models import SheetColumnModel, SheetModel, SheetRecordModel, WorkspaceModel
from app.services.workspace.constants import WORKSPACE_COLORS
from app.services.workspace.ordering import sort_columns, sort_records, sort_sheets, sort_workbooks
from app.services.workspace.validation import sanitize_formula_alias

SYSTEM_GRID_COLUMN_TYPES = frozenset({"created_by", "created_at", "updated_by", "updated_at"})
ROW_CREATED_BY_META_KEY = "__record_created_by"
ROW_CREATED_AT_META_KEY = "__record_created_at"
ROW_UPDATED_BY_META_KEY = "__record_updated_by"
ROW_UPDATED_AT_META_KEY = "__record_updated_at"


def column_type_to_grid_data_type(value: str) -> str:
    if value in {"number", "duration", "percent", "formula"}:
        return "number"
    if value in {"date", "datetime", "created_at", "updated_at"}:
        return "date"
    if value == "status":
        return "status"
    return "text"


def resolve_system_grid_column_kind(column: dict[str, object]) -> str | None:
    column_type = str(column.get("column_type") or "").strip()
    if column_type in SYSTEM_GRID_COLUMN_TYPES:
        return column_type

    settings = column.get("settings") if isinstance(column.get("settings"), dict) else {}
    if settings.get("system") is True:
        legacy_key = str(column.get("key") or "").strip()
        if legacy_key in SYSTEM_GRID_COLUMN_TYPES:
            return legacy_key

    return None


def resolve_record_creator_name(record: SheetRecordModel) -> str:
    creator = getattr(record, "creator", None)
    if creator is not None:
        full_name = str(getattr(creator, "full_name", "") or "").strip()
        if full_name:
            return full_name

        email = str(getattr(creator, "email", "") or "").strip()
        if email:
            return email

    return str(record.created_by)


def resolve_record_updater_name(record: SheetRecordModel) -> str:
    updater = getattr(record, "updater", None)
    if updater is not None:
        full_name = str(getattr(updater, "full_name", "") or "").strip()
        if full_name:
            return full_name

        email = str(getattr(updater, "email", "") or "").strip()
        if email:
            return email

    return str(record.updated_by)


def resolve_system_row_value(record: SheetRecordModel, kind: str) -> object:
    if kind == "created_by":
        return resolve_record_creator_name(record)
    if kind == "created_at":
        return record.created_at.isoformat()
    if kind == "updated_by":
        return resolve_record_updater_name(record)
    if kind == "updated_at":
        return record.updated_at.isoformat()
    return None


def row_payload(
    record: SheetRecordModel,
    *,
    writable_column_keys: set[str] | None = None,
    system_columns: dict[str, str] | None = None,
) -> dict[str, object]:
    payload: dict[str, object] = {
        "id": record.id,
        ROW_CREATED_BY_META_KEY: resolve_record_creator_name(record),
        ROW_CREATED_AT_META_KEY: record.created_at.isoformat(),
        ROW_UPDATED_BY_META_KEY: resolve_record_updater_name(record),
        ROW_UPDATED_AT_META_KEY: record.updated_at.isoformat(),
    }
    row_values = deepcopy(record.values_json or {})

    if writable_column_keys is not None:
        row_values = {
            key: value for key, value in row_values.items() if key in writable_column_keys
        }

    payload.update(row_values)

    for column_key, system_kind in sorted((system_columns or {}).items()):
        payload[column_key] = resolve_system_row_value(record, system_kind)

    return payload


def column_options(settings_json: dict[str, object] | None) -> list[str]:
    raw_options = (settings_json or {}).get("options")
    if not isinstance(raw_options, list):
        return []

    options: list[str] = []
    for option in raw_options:
        normalized_option = str(option).strip()
        if normalized_option:
            options.append(normalized_option)

    return options


def serialize_sheet_column(column: SheetColumnModel) -> dict[str, object]:
    settings = deepcopy(column.settings_json or {})
    raw_formula_alias = settings.pop("formulaAlias", settings.pop("formula_alias", None))
    formula_alias = sanitize_formula_alias(raw_formula_alias)
    is_system_column = column.type in SYSTEM_GRID_COLUMN_TYPES

    return {
        "key": column.key,
        "label": column.title,
        "formula_alias": formula_alias,
        "data_type": column_type_to_grid_data_type(column.type),
        "column_type": column.type,
        "width": column.width,
        "editable": False if is_system_column else column.editable,
        "computed": column.computed,
        "expression": column.expression,
        "options": column_options(settings),
        "settings": settings,
    }


def serialize_sheet(sheet: SheetModel, *, include_grid: bool = False) -> dict[str, object]:
    active_records = [record for record in sort_records(sheet.records) if not record.archived]
    raw_initial_placeholder_row_budget = (sheet.config_json or {}).get("initial_placeholder_row_budget")
    data: dict[str, object] = {
        "id": sheet.id,
        "key": sheet.key,
        "name": sheet.name,
        "icon": "grid" if sheet.kind in {"data", "derived"} else sheet.kind,
        "kind": sheet.kind,
        "row_count": len(active_records),
        "initial_placeholder_row_budget": max(
            0,
            raw_initial_placeholder_row_budget if isinstance(raw_initial_placeholder_row_budget, int) else 0,
        ),
        "updated_at": sheet.updated_at.isoformat(),
    }

    if include_grid:
        ordered_columns = sort_columns(sheet.columns)
        serialized_columns = [serialize_sheet_column(column) for column in ordered_columns]
        writable_column_keys = {
            str(column.get("key") or "")
            for column in serialized_columns
            if str(column.get("key") or "")
            and not bool(column.get("computed"))
            and not str(column.get("expression") or "").strip()
            and resolve_system_grid_column_kind(column) is None
        }
        system_columns = {
            str(column.get("key") or ""): system_kind
            for column in serialized_columns
            if str(column.get("key") or "")
            for system_kind in [resolve_system_grid_column_kind(column)]
            if system_kind is not None
        }
        data["columns"] = serialized_columns
        data["rows"] = [
            row_payload(
                record,
                writable_column_keys=writable_column_keys,
                system_columns=system_columns,
            )
            for record in active_records
        ]
        data["styles"] = deepcopy(sheet.styles_json or [])

    return data


def serialize_workbook_context_sheet(sheet: SheetModel) -> dict[str, object]:
    return serialize_sheet(sheet, include_grid=True)


def serialize_sheet_workbook_context(
    workspace: WorkspaceModel,
    *,
    workbook_id: str | None = None,
) -> dict[str, object]:
    return {
        "sheets": [
            serialize_workbook_context_sheet(sheet)
            for workbook in sort_workbooks(workspace.workbooks)
            if workbook_id is None or workbook.id == workbook_id
            for sheet in sort_sheets(workbook.sheets)
        ]
    }


def serialize_workspace(workspace: WorkspaceModel) -> dict[str, object]:
    flattened_sheets = [
        sheet
        for workbook in sort_workbooks(workspace.workbooks)
        for sheet in sort_sheets(workbook.sheets)
    ]
    return {
        "id": workspace.id,
        "name": workspace.name,
        "description": workspace.description,
        "color": workspace.color,
        "sheet_count": len(flattened_sheets),
        "sheets": [serialize_sheet(sheet) for sheet in flattened_sheets],
    }


def workspace_color(position: int) -> str:
    return WORKSPACE_COLORS[abs(position) % len(WORKSPACE_COLORS)]
