import logging
import os
import re
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path

_WINDOWS_PATH = re.compile(r"[A-Za-z]:\\[^\r\n\"']+")

def application_root():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]

def resolve_logs_root():
    configured = os.getenv("M365_LOGS_ROOT", "").strip()
    candidates = []
    if configured:
        candidates.append(Path(configured).expanduser())
    candidates.append(application_root() / "logs")
    local = os.getenv("LOCALAPPDATA", "").strip()
    if local:
        candidates.append(Path(local) / "M365 Mailbox Backup" / "logs")
    candidates.append(Path(tempfile.gettempdir()) / "M365 Mailbox Backup" / "logs")
    last_error = None
    for candidate in candidates:
        try:
            candidate = candidate.resolve()
            candidate.mkdir(parents=True, exist_ok=True)
            probe = candidate / ".write_probe"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            os.environ["M365_LOGS_ROOT"] = str(candidate)
            return candidate
        except Exception as error:
            last_error = error
    raise RuntimeError(f"Nenhum diretório de logs está disponível: {last_error}")

def compact_log_text(text, max_length=520):
    if not text: return text
    def shorten(match):
        value=match.group(0); parts=value.split("\\")
        return value if len(parts)<=3 else "…\\"+"\\".join(parts[-3:])
    compacted=_WINDOWS_PATH.sub(shorten,str(text))
    return compacted if len(compacted)<=max_length else compacted[:max_length-1]+"…"

class CompactFormatter(logging.Formatter):
    def format(self, record):
        try: message=record.getMessage()
        except Exception: message=str(record.msg)
        formatted=super().format(record)
        if isinstance(message,str) and message.startswith(("[PROGRESS] ","[PST-PROGRESS] ","[PST-EVENT] ","[ROOT-CAUSE] ","[REPAIR-PROGRESS] ","[REPAIR-EVENT] ","[REPAIR-SUMMARY] ")):
            return formatted
        return compact_log_text(formatted)

def _utf8_console_stream():
    stream=getattr(sys,"stderr",None)
    reconfigure=getattr(stream,"reconfigure",None)
    if callable(reconfigure):
        try: reconfigure(encoding="utf-8",errors="backslashreplace",line_buffering=True)
        except Exception: pass
    return stream

def _file_handler(path, formatter, level=logging.INFO):
    path.parent.mkdir(parents=True,exist_ok=True)
    handler=RotatingFileHandler(path,maxBytes=10*1024*1024,backupCount=5,encoding="utf-8",delay=False)
    handler.setLevel(level); handler.setFormatter(formatter)
    return handler

def _reset_logger(logger):
    for handler in list(logger.handlers):
        try: handler.flush(); handler.close()
        except Exception: pass
        logger.removeHandler(handler)

def setup_logger():
    logs=resolve_logs_root(); formatter=CompactFormatter("%(asctime)s | %(levelname)s | %(message)s")
    logger=logging.getLogger("m365_mailbox_backup"); logger.setLevel(logging.INFO); logger.propagate=False
    expected=(logs/"m365_mailbox_backup.log").resolve()
    valid=any(isinstance(h,RotatingFileHandler) and Path(h.baseFilename).resolve()==expected for h in logger.handlers)
    if not valid:
        _reset_logger(logger); logger.addHandler(_file_handler(expected,formatter))
        stream=_utf8_console_stream()
        if stream is not None:
            console=logging.StreamHandler(stream); console.setFormatter(formatter); logger.addHandler(console)
    operation_id=os.getenv("M365_OPERATION_ID","").strip()
    if operation_id:
        safe=re.sub(r"[^A-Za-z0-9_.-]+","_",operation_id)
        op_path=(logs/"backup"/f"{safe}.log").resolve()
        if not any(isinstance(h,RotatingFileHandler) and Path(h.baseFilename).resolve()==op_path for h in logger.handlers):
            logger.addHandler(_file_handler(op_path,formatter))
    if not getattr(logger,"_startup_written",False):
        logger.info("Logger EML inicializado | pid=%s | logs=%s",os.getpid(),logs)
        logger._startup_written=True
    return logger

def setup_report_logger():
    main=setup_logger(); logger=logging.getLogger("m365_mailbox_backup.report")
    logger.setLevel(logging.INFO); logger.propagate=False
    if not logger.handlers:
        for h in main.handlers:
            if isinstance(h,RotatingFileHandler): logger.addHandler(h)
        stream=_utf8_console_stream()
        if stream is not None:
            console=logging.StreamHandler(stream); console.setFormatter(logging.Formatter("%(message)s")); logger.addHandler(console)
    return logger

def setup_pst_logger(operation_id, logs_root=None):
    logs=Path(logs_root).expanduser().resolve() if logs_root else resolve_logs_root()
    safe="".join(c if c.isalnum() or c in "._-" else "_" for c in str(operation_id or "pst"))
    logger=logging.getLogger(f"m365_mailbox_backup.pst.{safe}"); logger.setLevel(logging.DEBUG); logger.propagate=False
    path=(logs/"pst"/f"{safe}.log").resolve(); formatter=CompactFormatter("%(asctime)s | %(levelname)s | %(message)s")
    valid=any(isinstance(h,RotatingFileHandler) and Path(h.baseFilename).resolve()==path for h in logger.handlers)
    if not valid:
        _reset_logger(logger)
        level_name=os.getenv("M365_PST_FILE_LOG_LEVEL","INFO").upper()
        logger.addHandler(_file_handler(path,formatter,getattr(logging,level_name,logging.INFO)))
    if not getattr(logger,"_startup_written",False):
        logger.info("Logger PST inicializado | operação=%s | pid=%s | arquivo=%s",operation_id,os.getpid(),path)
        logger._startup_written=True
    return logger
