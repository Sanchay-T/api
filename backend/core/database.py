import logging
import ssl

from sqlalchemy.ext.asyncio import (
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)

from backend.core.config import settings

logger = logging.getLogger(__name__)

# asyncpg does not understand the `sslmode` query param that DigitalOcean
# suggests for libpq‑based drivers.  Instead we pass an explicit `ssl=True`
# via `connect_args` **only** when the URL scheme is Postgres.
connect_args: dict = {}
if settings.database_url.startswith("postgresql"):
    # Load the CA certificate provided by DigitalOcean
    ca_path = "backend/ca/ca-certificate.crt"
    ssl_ctx = ssl.create_default_context(cafile=ca_path)
    # Optional: You might need this if the hostname doesn't match exactly,
    # but usually create_default_context handles it. Try without first.
    # ssl_ctx.check_hostname = True
    # ssl_ctx.verify_mode = ssl.CERT_REQUIRED # This is the default
    connect_args["ssl"] = ssl_ctx

engine = create_async_engine(
    settings.database_url, echo=False, future=True, connect_args=connect_args
)
SessionLocal = async_sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)


async def get_db() -> AsyncSession:  # pragma: no cover
    async with SessionLocal() as session:
        yield session


logger.info("✔ core.database ready (engine=%s)", settings.database_url.split(":")[0])
