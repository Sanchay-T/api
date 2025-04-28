import logging
import uuid
from typing import AsyncGenerator

from fastapi import Depends
from fastapi_users import schemas
from fastapi_users_db_sqlalchemy import (
    SQLAlchemyBaseUserTableUUID,
    SQLAlchemyUserDatabase,
)
from sqlalchemy.ext.asyncio import AsyncSession

from backend.core.database import get_db
from backend.models.base import Base

logger = logging.getLogger(__name__)


class User(SQLAlchemyBaseUserTableUUID, Base):
    """User table (UUID primary key).

    Additional custom columns can be added later (e.g., `full_name`)."""

    __tablename__ = "users"  # explicit name keeps things predictable


# ---------- Pydantic schemas ---------- #


class UserRead(schemas.BaseUser[uuid.UUID]):
    pass


class UserCreate(schemas.BaseUserCreate):
    pass


class UserUpdate(schemas.BaseUserUpdate):
    pass


# ---------- Dependency helpers ---------- #


async def get_user_db(
    session: AsyncSession = Depends(get_db),
) -> AsyncGenerator[SQLAlchemyUserDatabase, None]:
    """Yield the FastAPI‑Users database adapter."""

    yield SQLAlchemyUserDatabase(session, User)


logger.info("✔ models.user ready")
