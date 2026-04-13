from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

GridColumnDataType = Literal["text", "number", "currency", "date", "status"]
GridColumnType = Literal[
    "text",
    "number",
    "currency",
    "date",
    "datetime",
    "duration",
    "percent",
    "status",
    "user",
    "formula",
]


class GridColumn(BaseModel):
    key: str
    label: str
    data_type: GridColumnDataType = "text"
    column_type: GridColumnType = "text"
    width: int | None = None
    editable: bool = True
    computed: bool = False
    expression: str | None = None
    options: list[str] = Field(default_factory=list)
    settings: dict[str, Any] = Field(default_factory=dict)


class SheetSummary(BaseModel):
    id: str
    key: str
    name: str
    icon: str = "grid"
    kind: str = "data"
    row_count: int = 0
    updated_at: str


class SheetDetail(SheetSummary):
    columns: list[GridColumn] = Field(default_factory=list)
    rows: list[dict[str, Any]] = Field(default_factory=list)


class SheetCreateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=80)


class SheetUpdateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=80)


class SheetGridUpdateRequest(BaseModel):
    columns: list[GridColumn] = Field(min_length=1)
    rows: list[dict[str, Any]] = Field(default_factory=list)


class SheetCellHistoryActor(BaseModel):
    id: str | None = None
    email: str | None = None
    full_name: str
    avatar_url: str | None = None


class SheetCellHistoryEntry(BaseModel):
    id: str
    revision_id: str
    record_id: str
    column_key: str
    previous_value: Any | None = None
    next_value: Any | None = None
    changed_at: str
    actor: SheetCellHistoryActor | None = None


class SheetCellHistoryResponse(BaseModel):
    items: list[SheetCellHistoryEntry] = Field(default_factory=list)
