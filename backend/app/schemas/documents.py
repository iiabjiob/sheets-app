from __future__ import annotations

from pydantic import BaseModel, Field

from app.schemas.sheets import SheetDetail
from app.schemas.workspaces import WorkspaceSummary


class SheetWorkbookContextResponse(BaseModel):
    sheets: list[SheetDetail] = Field(default_factory=list)


class SheetDocumentResponse(BaseModel):
    workspace: WorkspaceSummary
    sheet: SheetDetail
    workbook: SheetWorkbookContextResponse
