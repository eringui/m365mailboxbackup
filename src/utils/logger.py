import logging
import os
import re
from logging.handlers import RotatingFileHandler
from pathlib import Path


_WINDOWS_PATH = re.compile(r"[A-Za-z]:\\[^\r\n\"']+")


def compact_log_text(text, max_length=260):
    if not text:
        return text

    def shorten(match):
        value = match.group(0)
        parts = value.split("\\")
        if len(parts) <= 2:
            return value
        return "…\\" + "\\".join(parts[-2:])

    compacted = _WINDOWS_PATH.sub(shorten, str(text))
    return compacted if len(compacted) <= max_length else compacted[: max_length - 1] + "…"


class CompactFormatter(logging.Formatter):
    def format(self, record):
        # Preserve structured progress lines intact so downstream parsers can
        # consume full JSON payloads. Only compact ordinary log lines.
        try:
            message = record.getMessage() if hasattr(record, 'getMessage') else str(record.msg)
        except Exception:
            message = str(record.msg)
        if isinstance(message, str) and message.startswith((
            "[PROGRESS] ", "[PST-PROGRESS] ", "[PST-EVENT] "
        )):
            # Return the full formatted line without compacting/truncation
            return super().format(record)
        return compact_log_text(super().format(record))


def setup_logger():
    base_dir = Path(__file__).resolve().parents[2]
    logs_dir = base_dir / "logs"

    logs_dir.mkdir(exist_ok=True)

    log_file = logs_dir / "m365_mailbox_backup.log"

    logger = logging.getLogger("m365_mailbox_backup")

    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    formatter = CompactFormatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8"
    )

    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

def setup_pst_logger(operation_id, logs_root=None):
    base_dir = Path(__file__).resolve().parents[2]
    logs_dir = Path(logs_root or (base_dir / "logs")) / "pst"
    logs_dir.mkdir(parents=True, exist_ok=True)
    safe_id = "".join(c if c.isalnum() or c in "._-" else "_" for c in str(operation_id))
    logger = logging.getLogger(f"m365_mailbox_backup.pst.{safe_id}")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    if not logger.handlers:
        formatter = CompactFormatter("%(asctime)s | %(levelname)s | %(message)s")
        handler = RotatingFileHandler(
            logs_dir / f"{safe_id}.log", maxBytes=10 * 1024 * 1024,
            backupCount=5, encoding="utf-8"
        )
        handler.setFormatter(formatter)
        level_name = str(os.getenv("M365_PST_FILE_LOG_LEVEL", "INFO")).upper()
        handler.setLevel(getattr(logging, level_name, logging.INFO))
        logger.addHandler(handler)
    return logger
