from __future__ import annotations

from datetime import datetime

from fastapi import APIRouter, Depends, Query, Response, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.deps import get_current_user
from app.infrastructure.db.database import get_db
from app.models import UserModel
from app.schemas.common import MoveRequest
from app.schemas.documents import SheetDocumentResponse
from app.schemas.sheets import (
    SheetActivityResponse,
    SheetCreateRequest,
    SheetCellHistoryResponse,
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


@router.get("/{workspace_id}/sheets/{sheet_id}/cells/history", response_model=SheetCellHistoryResponse)
async def get_sheet_cell_history(
    workspace_id: str,
    sheet_id: str,
    record_id: str,
    column_key: str,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetCellHistoryResponse:
    return SheetCellHistoryResponse.model_validate(
        await workspace_service.get_sheet_cell_history(
            db,
            current_user.id,
            workspace_id,
            sheet_id,
            record_id=record_id,
            column_key=column_key,
        )
    )


@router.get("/{workspace_id}/sheets/{sheet_id}/activity", response_model=SheetActivityResponse)
async def get_sheet_activity(
    workspace_id: str,
    sheet_id: str,
    created_from: datetime | None = Query(default=None),
    created_to: datetime | None = Query(default=None),
    action_type: list[str] | None = Query(default=None),
    user_id: list[str] | None = Query(default=None),
    limit: int = Query(default=200, ge=1, le=500),
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> SheetActivityResponse:
    return SheetActivityResponse.model_validate(
        await workspace_service.get_sheet_activity(
            db,
            current_user.id,
            workspace_id,
            sheet_id,
            created_from=created_from,
            created_to=created_to,
            action_types=action_type,
            user_ids=user_id,
            limit=limit,
        )
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
            styles=None
            if payload.styles is None
            else [style_rule.model_dump(exclude_none=True) for style_rule in payload.styles],
        )
    )
