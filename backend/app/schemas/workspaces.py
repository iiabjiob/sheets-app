from __future__ import annotations

from pydantic import BaseModel, Field

from app.schemas.sheets import SheetSummary


class WorkspaceSummary(BaseModel):
    id: str
    name: str
    description: str
    color: str
    sheet_count: int = 0
    sheets: list[SheetSummary] = Field(default_factory=list)


class WorkspaceCollectionResponse(BaseModel):
    items: list[WorkspaceSummary] = Field(default_factory=list)


class WorkspaceCreateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=80)


class WorkspaceUpdateRequest(BaseModel):
    name: str = Field(min_length=1, max_length=80)