from __future__ import annotations

from app.models import SheetColumnModel, SheetModel, SheetRecordModel, WorkbookModel, WorkspaceModel


def sort_workbooks(workbooks: list[WorkbookModel]) -> list[WorkbookModel]:
    return sorted(workbooks, key=lambda item: (item.position, item.created_at, item.id))


def sort_sheets(sheets: list[SheetModel]) -> list[SheetModel]:
    return sorted(sheets, key=lambda item: (item.position, item.created_at, item.id))


def sort_columns(columns: list[SheetColumnModel]) -> list[SheetColumnModel]:
    return sorted(columns, key=lambda item: (item.position, item.created_at, item.id))


def sort_records(records: list[SheetRecordModel]) -> list[SheetRecordModel]:
    return sorted(records, key=lambda item: (item.position, item.created_at, item.id))


def reindex_workspaces(workspaces: list[WorkspaceModel]) -> None:
    for index, workspace in enumerate(workspaces):
        workspace.position = index


def reindex_sheets(sheets: list[SheetModel]) -> None:
    for index, sheet in enumerate(sheets):
        sheet.position = index


def resolve_target_index(index: int, length: int, direction: str) -> int | None:
    if direction == "up":
        return index - 1 if index > 0 else None
    if direction == "down":
        return index + 1 if index < length - 1 else None
    return None