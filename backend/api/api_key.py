import logging
import secrets

from fastapi import APIRouter, Depends, Header, HTTPException, status
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from backend.auth import fastapi_users
from backend.core.database import get_db
from backend.models.api_key import APIKey
from backend.models.user import User

logger = logging.getLogger(__name__)

router = APIRouter(tags=["api-key"], prefix="/users")

current_user = fastapi_users.current_user(active=True)


@router.post("/api-key", summary="Generate a new API key for the authenticated user")
async def create_api_key(
    user: User = Depends(current_user), session: AsyncSession = Depends(get_db)
):
    key_value = secrets.token_urlsafe(32)
    api_key = APIKey(key=key_value, user_id=str(user.id))
    session.add(api_key)
    await session.commit()
    logger.info(
        "🔑  API key issued for user %s …%s",
        str(user.email).split("@")[0],
        key_value[-4:],
    )
    return {"api_key": key_value}


async def get_api_key(
    x_api_key: str | None = Header(default=None, alias="X-API-Key"),
    session: AsyncSession = Depends(get_db),
):
    if not x_api_key:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED, detail="API key missing"
        )

    query = select(APIKey).where(APIKey.key == x_api_key)
    result = await session.execute(query)
    api_key_row = result.scalar_one_or_none()
    if not api_key_row:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN, detail="Invalid API key"
        )
    return api_key_row
