import logging
import sys

import structlog

from .request_id import request_id_ctx_var

LOG_FORMAT = (
    "% (levelname)s | %(name)s | %(message)s"  # placeholder not used with structlog
)


class RequestIdProcessor:
    def __call__(self, logger, method_name, event_dict):
        rid = request_id_ctx_var.get(None)
        if rid:
            event_dict["request_id"] = rid
        return event_dict


def init_logging() -> None:
    if structlog.is_configured():
        return

    structlog.configure(
        processors=[
            structlog.contextvars.merge_contextvars,
            RequestIdProcessor(),
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.StackInfoRenderer(),
            structlog.processors.format_exc_info,
            structlog.dev.ConsoleRenderer(colors=True),
        ],
        wrapper_class=structlog.make_filtering_bound_logger(logging.INFO),
        cache_logger_on_first_use=True,
    )

    logging.basicConfig(format="%(message)s", stream=sys.stdout, level=logging.INFO)
    structlog.get_logger(__name__).info("✔ core.logging ready (structlog)")
