"""
Structured logging configuration for the xlsx_parser.

Provides a JSON-structured logging format with workbook/sheet/block context
fields for observability and debugging in production.
"""

from __future__ import annotations

import json
import logging
import sys
from datetime import datetime, timezone


class StructuredFormatter(logging.Formatter):
    """JSON-structured log formatter with context fields."""

    def format(self, record: logging.LogRecord) -> str:
        log_entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "module": record.module,
            "function": record.funcName,
            "line": record.lineno,
        }
        # Add context fields if present
        for field in ("workbook_id", "sheet_name", "block_id", "stage"):
            if hasattr(record, field):
                log_entry[field] = getattr(record, field)

        if record.exc_info and record.exc_info[1]:
            log_entry["exception"] = str(record.exc_info[1])

        return json.dumps(log_entry)


def configure_logging(
    level: int = logging.INFO,
    structured: bool = False,
) -> None:
    """
    Configure logging for the xlsx_parser package.

    Args:
        level: Logging level (default INFO).
        structured: If True, use JSON-structured output. Otherwise, standard format.
    """
    root_logger = logging.getLogger("xlsx_parser")
    root_logger.setLevel(level)

    # Remove existing handlers
    root_logger.handlers.clear()

    handler = logging.StreamHandler(sys.stderr)
    handler.setLevel(level)

    if structured:
        handler.setFormatter(StructuredFormatter())
    else:
        handler.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))

    root_logger.addHandler(handler)
