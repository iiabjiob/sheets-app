from __future__ import annotations

from fastapi import APIRouter, Depends, Response, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.deps import get_current_user
from app.infrastructure.db.database import get_db
from app.models import UserModel
from app.schemas.common import MoveRequest
from app.schemas.workspaces import (
    WorkspaceCollectionResponse,
    WorkspaceCreateRequest,
    WorkspaceSummary,
    WorkspaceUpdateRequest,
)
from app.services.workspace import workspace_service

router = APIRouter(tags=["Workspaces"])


@router.get("/", response_model=WorkspaceCollectionResponse)
async def list_workspaces(
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> WorkspaceCollectionResponse:
    return WorkspaceCollectionResponse(items=await workspace_service.list_workspaces(db, current_user.id))


@router.post("/", response_model=WorkspaceSummary, status_code=status.HTTP_201_CREATED)
async def create_workspace(
    payload: WorkspaceCreateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> WorkspaceSummary:
    return WorkspaceSummary.model_validate(
        await workspace_service.create_workspace(db, current_user.id, payload.name)
    )


@router.delete("/{workspace_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_workspace(
    workspace_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> Response:
    await workspace_service.delete_workspace(db, current_user.id, workspace_id)
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.patch("/{workspace_id}", response_model=WorkspaceSummary)
async def rename_workspace(
    workspace_id: str,
    payload: WorkspaceUpdateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> WorkspaceSummary:
    return WorkspaceSummary.model_validate(
        await workspace_service.rename_workspace(db, current_user.id, workspace_id, payload.name)
    )


@router.post("/{workspace_id}/duplicate", response_model=WorkspaceSummary, status_code=status.HTTP_201_CREATED)
async def duplicate_workspace(
    workspace_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> WorkspaceSummary:
    return WorkspaceSummary.model_validate(
        await workspace_service.duplicate_workspace(db, current_user.id, workspace_id)
    )


@router.post("/{workspace_id}/move", status_code=status.HTTP_204_NO_CONTENT)
async def move_workspace(
    workspace_id: str,
    payload: MoveRequest,
    db: AsyncSession = Depends(get_db),
    current_user: UserModel = Depends(get_current_user),
) -> Response:
    await workspace_service.move_workspace(db, current_user.id, workspace_id, payload.direction)
    return Response(status_code=status.HTTP_204_NO_CONTENT)