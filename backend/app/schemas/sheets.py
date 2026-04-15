from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator, model_validator

from app.services.workspace.validation import normalize_and_validate_formula_alias

GridColumnDataType = Literal["text", "number", "currency", "date", "status"]
GridColumnType = Literal[
    "text",
    "number",
    "currency",
    "checkbox",
    "created_at",
    "created_by",
    "date",
    "datetime",
    "duration",
    "percent",
    "select",
    "status",
    "updated_at",
    "updated_by",
    "user",
    "formula",
]


class GridColumn(BaseModel):
    key: str
    label: str
    formula_alias: str | None = None
    data_type: GridColumnDataType = "text"
    column_type: GridColumnType = "text"
    width: int | None = None
    editable: bool = True
    computed: bool = False
    expression: str | None = None
    options: list[str] = Field(default_factory=list)
    settings: dict[str, Any] = Field(default_factory=dict)

    @field_validator("formula_alias", mode="before")
    @classmethod
    def validate_formula_alias(cls, value: object) -> str | None:
        return normalize_and_validate_formula_alias(value)


def _normalize_optional_style_string(value: object) -> str | None:
    if value is None:
        return None

    normalized = str(value).strip()
    return normalized or None


class SheetStyleRange(BaseModel):
    start_row: int = Field(ge=0)
    end_row: int = Field(ge=0)
    start_column: int = Field(ge=0)
    end_column: int = Field(ge=0)

    @model_validator(mode="after")
    def validate_bounds(self) -> "SheetStyleRange":
        if self.end_row < self.start_row:
            raise ValueError("Style range end_row must be greater than or equal to start_row.")
        if self.end_column < self.start_column:
            raise ValueError("Style range end_column must be greater than or equal to start_column.")
        return self


class SheetCellStyle(BaseModel):
    font_family: str | None = None
    font_size: int | None = Field(default=None, ge=1, le=512)
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strikethrough: bool | None = None
    text_color: str | None = None
    background_color: str | None = None
    horizontal_align: Literal["left", "center", "right"] | None = None
    vertical_align: Literal["top", "middle", "bottom"] | None = None
    wrap_mode: Literal["overflow", "clip", "wrap"] | None = None
    number_format: str | None = None

    @field_validator(
        "font_family",
        "text_color",
        "background_color",
        "number_format",
        mode="before",
    )
    @classmethod
    def normalize_optional_strings(cls, value: object) -> str | None:
        return _normalize_optional_style_string(value)

    @model_validator(mode="after")
    def validate_not_empty(self) -> "SheetCellStyle":
        if not self.model_dump(exclude_none=True):
            raise ValueError("Style rule must define at least one style property.")
        return self


class SheetStyleRule(BaseModel):
    range: SheetStyleRange
    style: SheetCellStyle


class SheetSummary(BaseModel):
    id: str
    key: str
    name: str
    icon: str = "grid"
    kind: str = "data"
    row_count: int = 0
    initial_placeholder_row_budget: int = 0
    updated_at: str


class SheetDetail(SheetSummary):
    columns: list[GridColumn] = Field(default_factory=list)
    rows: list[dict[str, Any]] = Field(default_factory=list)
    styles: list[SheetStyleRule] = Field(default_factory=list)


class SheetCreateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=80)


class SheetUpdateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=80)


class SheetGridUpdateRequest(BaseModel):
    columns: list[GridColumn] = Field(min_length=1)
    rows: list[dict[str, Any]] = Field(default_factory=list)
    styles: list[SheetStyleRule] | None = None


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


class SheetActivityActor(BaseModel):
    id: str | None = None
    email: str | None = None
    full_name: str
    avatar_url: str | None = None


class SheetActivityEntry(BaseModel):
    id: str
    action_type: str
    workbook_id: str | None = None
    sheet_id: str | None = None
    record_id: str | None = None
    payload: dict[str, Any] = Field(default_factory=dict)
    created_at: str
    actor: SheetActivityActor | None = None


class SheetActivityActionOption(BaseModel):
    action_type: str
    count: int = Field(ge=0)


class SheetActivityCollaboratorOption(BaseModel):
    actor: SheetActivityActor
    count: int = Field(ge=0)


class SheetActivityResponse(BaseModel):
    items: list[SheetActivityEntry] = Field(default_factory=list)
    actions: list[SheetActivityActionOption] = Field(default_factory=list)
    collaborators: list[SheetActivityCollaboratorOption] = Field(default_factory=list)
