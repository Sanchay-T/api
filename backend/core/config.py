import logging

from pydantic import Field
from pydantic_settings import BaseSettings

logger = logging.getLogger(__name__)


class Settings(BaseSettings):
    """Central project configuration pulled from environment variables.

    This module must stay *extremely* small – only simple primitives so it can be
    imported anywhere without side‑effects (except for the ready banner)."""

    # FastAPI server network
    api_host: str = Field("127.0.0.1", env="API_HOST")
    api_port: int = Field(7500, env="API_PORT")

    # Database – defaults to Sqlite for local dev.  In prod set DATABASE_URL env var, e.g.
    # postgres+asyncpg://user:pass@host:port/dbname?sslmode=require
    database_url: str = Field(
        "sqlite+aiosqlite:///./db.sqlite3",
        env="DATABASE_URL",
    )

    # JSON Web Token secret key (change in prod!)
    jwt_secret: str = Field("CHANGE_ME", env="JWT_SECRET")

    # Observability
    sentry_dsn: str | None = Field(None, alias="SENTRY_DSN", env="SENTRY_DSN")

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


# Instantiate *once* so downstream modules can do `from backend.core.config import settings`.
settings = Settings()

logger.info("✔ core.config ready (DB=%s)", settings.database_url.split(":")[0])
