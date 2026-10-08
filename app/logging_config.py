"""
SlideGuard Production Logging Configuration.
Sets up rotating file logging and clean console output without exposing secrets or raw XML payloads.
"""

import sys
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

from app.config import AppConfig


class SensitiveDataFilter(logging.Filter):
    """Filter that masks sensitive tokens or keys if accidentally logged."""

    def filter(self, record: logging.LogRecord) -> bool:
        msg = str(record.getMessage())
        # Mask potential long API keys matching standard patterns
        if "AIza" in msg:
            record.msg = "[MASKED_KEY]"
            record.args = ()
        return True


class SafeStreamHandler(logging.StreamHandler):
    """Stream handler that safely falls back to replacement if console lacks Unicode glyphs."""

    def emit(self, record: logging.LogRecord):
        try:
            super().emit(record)
        except UnicodeEncodeError:
            try:
                msg = self.format(record)
                stream = self.stream
                encoding = getattr(stream, "encoding", "utf-8") or "utf-8"
                safe_msg = msg.encode(encoding, errors="replace").decode(encoding)
                stream.write(safe_msg + self.terminator)
                self.flush()
            except Exception:
                self.handleError(record)
        except Exception:
            self.handleError(record)


def setup_production_logging(config: Optional[AppConfig] = None) -> Path:
    """
    Configures application-wide logging with both a rotating file handler and console output.
    Returns the path to the active rotating log file.
    """
    cfg = config or AppConfig.load_from_env()
    log_dir = cfg.log_dir
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "slideguard.log"

    root_logger = logging.getLogger()
    numeric_level = getattr(logging, cfg.log_level, logging.INFO)
    root_logger.setLevel(numeric_level)

    # Clear existing handlers to avoid duplicates
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)

    # 1. Console Stream Handler (Safe with Windows ANSI code pages)
    console_handler = SafeStreamHandler(sys.stdout)
    console_handler.setLevel(numeric_level)
    console_fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
        datefmt="%H:%M:%S"
    )
    console_handler.setFormatter(console_fmt)
    console_handler.addFilter(SensitiveDataFilter())
    root_logger.addHandler(console_handler)

    # 2. Rotating File Handler
    try:
        file_handler = RotatingFileHandler(
            filename=log_file,
            maxBytes=cfg.max_log_size_bytes,
            backupCount=cfg.backup_log_count,
            encoding="utf-8"
        )
        file_handler.setLevel(numeric_level)
        file_fmt = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] [%(name)s] [%(threadName)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        file_handler.setFormatter(file_fmt)
        file_handler.addFilter(SensitiveDataFilter())
        root_logger.addHandler(file_handler)
    except Exception as e:
        console_handler.handle(
            logging.LogRecord("slideguard", logging.WARNING, "", 0, f"Could not initialize rotating log file: {e}", (), None)
        )

    logging.getLogger("slideguard.init").info(
        "Production logging initialized at level %s. Log path: %s", cfg.log_level, log_file
    )
    return log_file
