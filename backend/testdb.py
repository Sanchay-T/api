import asyncio
import ssl

import sqlalchemy as sa
from sqlalchemy.ext.asyncio import create_async_engine

PG_DSN = (
    "postgresql+asyncpg://doadmin:AVNS_vbnP6NCTEDbqx_sav0-@"
    "db-postgresql-blr1-71405-do-user-21241169-0.l.db.ondigitalocean.com:25060/defaultdb"
)

ssl_ctx = ssl.create_default_context(cafile="backend/ca/ca-certificate.crt")


async def main():
    engine = create_async_engine(PG_DSN, connect_args={"ssl": ssl_ctx})
    async with engine.connect() as conn:
        val = await conn.scalar(sa.text("SELECT 1"))
        print("DB says:", val)
    await engine.dispose()


asyncio.run(main())
