from __future__ import annotations

from fastapi import APIRouter

from app.api.v1.workspaces.sheets import router as sheets_router
from app.api.v1.workspaces.workspaces import router as workspaces_router

router = APIRouter(prefix="/api/v1/workspaces")
router.include_router(workspaces_router)
router.include_router(sheets_router)