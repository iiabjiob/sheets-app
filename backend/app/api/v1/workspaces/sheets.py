from __future__ import annotations

from fastapi import APIRouter, Depends, Response, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.deps import get_current_user
from app.infrastructure.db.database import get_db
from app.models import UserModel
from app.schemas.common import MoveRequest
from app.schemas.documents import SheetDocumentResponse
from app.schemas.sheets import (
    SheetCreateRequest,
    SheetDetail,
    SheetGridUpdateRequest,
    SheetSummary,
    SheetUpdateRequest,
)
from app.services.workspace import workspace_service

router = APIRouter(tags=["Sheets"])


@router.post("/{workspace_id}/sheets", response_model=SheetSummary, status_code=status.HTTP_201_CREATED)
async def create_sheet(
    workspace_id: str,
    payload: SheetCreateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetSummary:
    return SheetSummary.model_validate(
        await workspace_service.create_sheet(db, current_user.id, workspace_id, payload.name)
    )


@router.delete("/{workspace_id}/sheets/{sheet_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_sheet(
    workspace_id: str,
    sheet_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> Response:
    await workspace_service.delete_sheet(db, current_user.id, workspace_id, sheet_id)
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.patch("/{workspace_id}/sheets/{sheet_id}", response_model=SheetSummary)
async def rename_sheet(
    workspace_id: str,
    sheet_id: str,
    payload: SheetUpdateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetSummary:
    return SheetSummary.model_validate(
        await workspace_service.rename_sheet(db, current_user.id, workspace_id, sheet_id, payload.name)
    )


@router.post(
    "/{workspace_id}/sheets/{sheet_id}/duplicate",
    response_model=SheetSummary,
    status_code=status.HTTP_201_CREATED,
)
async def duplicate_sheet(
    workspace_id: str,
    sheet_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetSummary:
    return SheetSummary.model_validate(
        await workspace_service.duplicate_sheet(db, current_user.id, workspace_id, sheet_id)
    )


@router.post("/{workspace_id}/sheets/{sheet_id}/move", status_code=status.HTTP_204_NO_CONTENT)
async def move_sheet(
    workspace_id: str,
    sheet_id: str,
    payload: MoveRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> Response:
    await workspace_service.move_sheet(db, current_user.id, workspace_id, sheet_id, payload.direction)
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.get("/{workspace_id}/sheets/{sheet_id}", response_model=SheetDocumentResponse)
async def get_sheet_document(
    workspace_id: str,
    sheet_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetDocumentResponse:
    return SheetDocumentResponse.model_validate(
        await workspace_service.get_sheet_document(db, current_user.id, workspace_id, sheet_id)
    )


@router.put("/{workspace_id}/sheets/{sheet_id}/grid", response_model=SheetDetail)
async def update_sheet_grid(
    workspace_id: str,
    sheet_id: str,
    payload: SheetGridUpdateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetDetail:
    return SheetDetail.model_validate(
        await workspace_service.update_sheet_grid(
            db,
            current_user.id,
            workspace_id,
            sheet_id,
            columns=[column.model_dump() for column in payload.columns],
            rows=payload.rows,
        )
    )
