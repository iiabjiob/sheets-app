import asyncio
from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from contextlib import asynccontextmanager

from app.api.v1.auth.router import router as auth_router
from app.api.v1.health.router import router as health_router
from app.api.v1.workspaces.router import router as workspace_router


from app.infrastructure.db.database import engine, initialize_database
from sqlalchemy import text


from app.core.config import get_settings
from app.core.logger import get_logger


settings = get_settings()
logger = get_logger("core")


async def check_database_connection(max_attempts: int = 10, base_delay: float = 1.5):
    """Ping the DB with retries so we can survive slow compose DNS/startup."""

    for attempt in range(1, max_attempts + 1):
        try:
            async with engine.begin() as conn:
                await conn.execute(text("SELECT 1"))
                logger.info("✅ Connected to the database!")
                return
        except Exception as e:
            logger.error(
                "💥 Database connection failed (attempt %s/%s): %s",
                attempt,
                max_attempts,
                e,
            )
            if attempt == max_attempts:
                raise
            await asyncio.sleep(base_delay * attempt)


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("🚀 Starting FastAPI application...")

    # Healthchecks
    await check_database_connection()
    await initialize_database()
    logger.info("🗃️ Database schema verified and initial seed ensured")

    try:
        yield
    finally:
        logger.info("🛑 Shutting down FastAPI application...")


app = FastAPI(
    title=settings.app_name,
    description=settings.description,
    version=settings.app_version,
    debug=settings.debug,
    lifespan=lifespan,
)

# Routers
logger.info("🔗 Registering REST API routers...")
app.include_router(health_router)
app.include_router(auth_router)
app.include_router(workspace_router)

logger.info("✅ REST API routers registered")

logger.info(f"✅ FastAPI application is up and running at version {settings.app_version}")


def _format_validation_message(error: dict[str, object]) -> str:
    error_type = str(error.get("type") or "")
    location = error.get("loc")
    field = location[-1] if isinstance(location, (list, tuple)) and location else None
    field_label = field.replace("_", " ").capitalize() if isinstance(field, str) else "Field"
    context = error.get("ctx") if isinstance(error.get("ctx"), dict) else {}

    if error_type == "missing":
        return f"{field_label} is required."
    if error_type == "string_too_short":
        min_length = context.get("min_length")
        return f"{field_label} must contain at least {min_length} characters."
    if error_type == "string_too_long":
        max_length = context.get("max_length")
        return f"{field_label} must contain at most {max_length} characters."
    if error_type == "string_pattern_mismatch":
        return f"{field_label} has an invalid format."

    message = str(error.get("msg") or "Invalid request.").strip()
    if field and isinstance(field, str) and field not in {"body", "query", "path"}:
        return f"{field_label}: {message}"
    return message


@app.exception_handler(RequestValidationError)
async def validation_exception_handler(_: Request, exc: RequestValidationError) -> JSONResponse:
    messages = [_format_validation_message(error) for error in exc.errors()]
    detail = "; ".join(dict.fromkeys(messages)) if messages else "Invalid request."
    return JSONResponse(status_code=422, content={"detail": detail})