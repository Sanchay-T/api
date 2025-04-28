import logging
import os
from typing import List

import sentry_sdk
from sentry_sdk.integrations.logging import LoggingIntegration
from sentry_sdk.integrations.sqlalchemy import SqlalchemyIntegration

try:
    from sentry_sdk.integrations.fastapi import FastAPIIntegration
except ImportError:  # sentry-sdk>=2.0 renamed integration
    from sentry_sdk.integrations.starlette import (
        StarletteIntegration as FastAPIIntegration,
    )

from backend.core.config import settings

_logger = logging.getLogger(__name__)


def init_sentry(
    extra_integrations: List[object] | None = None,
) -> None:  # pragma: no cover
    """Initialise Sentry once for the whole application.

    We pull the DSN from *settings.sentry_dsn* (env var ``SENTRY_DSN``).
    The function is safe to call multiple times – it will no‑op if Sentry
    has already been initialised or if no DSN is provided.

    Parameters
    ----------
    extra_integrations:
        Optional additional Sentry integrations (e.g. CeleryIntegration())
        you may want to pass from the entry‑point.
    """

    if not settings.sentry_dsn:
        _logger.info("Sentry disabled – no DSN provided")
        return

    if getattr(sentry_sdk.Hub.current, "client", None):
        # Already configured
        return

    base_integrations = [
        FastAPIIntegration(),
        SqlalchemyIntegration(),
        LoggingIntegration(level=logging.INFO, event_level=logging.ERROR),
    ]
    if extra_integrations:
        base_integrations.extend(extra_integrations)

    sentry_sdk.init(
        dsn=settings.sentry_dsn,
        integrations=base_integrations,
        # Capture traces at 20 % – adjust via env if needed
        traces_sample_rate=float(os.getenv("SENTRY_TRACES_SAMPLE_RATE", 0.2)),
        environment=os.getenv("APP_ENV", "local"),
        release=os.getenv("SENTRY_RELEASE"),
        _experiments={"auto_enabling_integrations": True},
    )
    _logger.info("✔ Sentry initialised (env=%s)", os.getenv("APP_ENV", "local"))
