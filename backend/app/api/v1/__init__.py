from app.api.v1.auth.router import router as auth_router
from app.api.v1.health.router import router as health_router
from app.api.v1.workspaces.router import router as workspace_router

__all__ = ["auth_router", "health_router", "workspace_router"]
