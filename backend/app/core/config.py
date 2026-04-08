from pathlib import Path
import os
from urllib.parse import quote
from pydantic_settings import BaseSettings, SettingsConfigDict


REPO_ROOT = Path(__file__).resolve().parents[2]
APP_ENV = os.getenv("APP_ENV", "development").lower()
ENV_FILE = None if APP_ENV in ("production", "prod") else REPO_ROOT / ".env.dev"

print("ENV_FILE resolved to:", ENV_FILE)

if ENV_FILE and not ENV_FILE.exists():
    raise FileNotFoundError("❌ ENV FILE NOT FOUND: .env.dev")


model_config: dict = {"extra": "allow"}
if ENV_FILE:
    model_config["env_file"] = str(ENV_FILE)


class Settings(BaseSettings):
    model_config = SettingsConfigDict(**model_config)

    # ---- BASE ----
    app_version: str = "1.0.0"
    app_name: str = "Project Backend"
    description: str = "Backend for Project application"
    app_env: str = "production"
    debug: bool = False
    debug_level: str = "INFO"

    # ---- DB ----
    postgres_user: str
    postgres_password: str
    postgres_host: str
    postgres_port: int
    postgres_db: str

    # ---- Database URL ----
    @property
    def database_url(self) -> str:
        user = quote(self.postgres_user)
        password = quote(self.postgres_password)
        return (
            f"postgresql+asyncpg://{user}:{password}"
            f"@{self.postgres_host}:{self.postgres_port}/{self.postgres_db}"
        )

def get_settings() -> Settings:
    return Settings()
