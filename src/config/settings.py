import os
import sys
from pathlib import Path
from dotenv import load_dotenv


BASE_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parents[2]
)

ENV_FILE = Path(
    os.getenv("M365_BACKUP_ENV_PATH", str(BASE_DIR / ".env"))
).expanduser().resolve()

load_dotenv(ENV_FILE)


TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

GRAPH_SCOPE = os.getenv(
    "GRAPH_SCOPE",
    "https://graph.microsoft.com/.default"
)

GRAPH_URL = os.getenv(
    "GRAPH_URL",
    "https://graph.microsoft.com/v1.0"
)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"


DEFAULT_OUTPUT_DIR = BASE_DIR / "output" / "backups"

OUTPUT_DIR = Path(
    os.getenv(
        "M365_BACKUP_OUTPUT_ROOT",
        str(DEFAULT_OUTPUT_DIR)
    )
).expanduser().resolve()



def _bounded_int_env(name, default, minimum, maximum):
    try:
        value = int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(maximum, value))


def _bounded_float_env(name, default, minimum, maximum):
    try:
        value = float(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(maximum, value))


EML_DOWNLOAD_WORKERS = _bounded_int_env(
    "M365_EML_DOWNLOAD_WORKERS", 24, 1, 64
)
EML_DOWNLOAD_CHUNK_SIZE = _bounded_int_env(
    "M365_EML_DOWNLOAD_CHUNK_SIZE_MB", 1, 1, 16
) * 1024 * 1024
EML_MAX_PENDING_DOWNLOADS = _bounded_int_env(
    "M365_EML_MAX_PENDING_DOWNLOADS", 96, 4, 256
)

# Controle preventivo e adaptativo de throttling para downloads MIME.
MIME_MAX_CONCURRENCY = _bounded_int_env(
    "M365_MIME_MAX_CONCURRENCY", 6, 1, 16
)
MIME_RESUME_MAX_CONCURRENCY = _bounded_int_env(
    "M365_RESUME_MIME_MAX_CONCURRENCY", 6, 1, 8
)
MIME_MIN_INTERVAL_SECONDS = _bounded_float_env(
    "M365_MIME_MIN_INTERVAL_SECONDS", 0.15, 0.0, 10.0
)
THROTTLE_SAFETY_SECONDS = _bounded_float_env(
    "M365_THROTTLE_SAFETY_SECONDS", 1.0, 0.0, 30.0
)
THROTTLE_JITTER_MAX_SECONDS = _bounded_float_env(
    "M365_THROTTLE_JITTER_MAX_SECONDS", 1.0, 0.0, 10.0
)
THROTTLE_RECOVERY_SECONDS = _bounded_int_env(
    "M365_THROTTLE_RECOVERY_SECONDS", 90, 30, 3600
)
ADAPTIVE_THROTTLING = os.getenv(
    "M365_ADAPTIVE_THROTTLING", "1"
).strip().lower() not in {"0", "false", "no", "off"}

# PyrateLimiter: limites preventivos por processo, mailbox e entre processos.
RATE_LIMITER_ENABLED = os.getenv(
    "M365_RATE_LIMITER_ENABLED", "1"
).strip().lower() not in {"0", "false", "no", "off"}
RATE_LIMITER_PROFILE = os.getenv("M365_RATE_LIMITER_PROFILE", "automatic").strip().lower()
GLOBAL_MIME_RATE_SECOND = _bounded_int_env(
    "M365_GLOBAL_MIME_RATE_SECOND", 10, 1, 100
)
GLOBAL_MIME_RATE_MINUTE = _bounded_int_env(
    "M365_GLOBAL_MIME_RATE_MINUTE", 480, 1, 6000
)
MAILBOX_MIME_RATE_SECOND = _bounded_int_env(
    "M365_MAILBOX_MIME_RATE_SECOND", 6, 1, 50
)
MAILBOX_MIME_RATE_MINUTE = _bounded_int_env(
    "M365_MAILBOX_MIME_RATE_MINUTE", 240, 1, 3000
)
RATE_LIMITER_WAIT_TIMEOUT = _bounded_float_env(
    "M365_RATE_LIMITER_WAIT_TIMEOUT", 300.0, 1.0, 3600.0
)
RATE_LIMITER_DB = Path(
    os.getenv("M365_RATE_LIMITER_DB", str(BASE_DIR / "_gui_state" / "graph_rate_limits.sqlite3"))
).expanduser().resolve()



# Pipeline EML -> PST: preparacao paralela e gravacao COM serial.
PST_PREPARE_WORKERS = _bounded_int_env("M365_PST_PREPARE_WORKERS", 3, 1, 8)
PST_PREPARE_QUEUE_SIZE = _bounded_int_env("M365_PST_PREPARE_QUEUE_SIZE", 12, 2, 100)
PST_LARGE_EML_MB = _bounded_int_env("M365_PST_LARGE_EML_MB", 25, 1, 500)

PST_VERIFICATION_BATCH_SIZE = _bounded_int_env(
    "M365_PST_VERIFICATION_BATCH_SIZE", 25, 1, 500
)
PST_VERIFICATION_FINAL_RETRIES = _bounded_int_env(
    "M365_PST_VERIFICATION_FINAL_RETRIES", 5, 1, 20
)

PST_MEMORY_BUDGET_MB = _bounded_int_env("M365_PST_MEMORY_BUDGET_MB", 512, 128, 8192)
PST_MIN_PREPARE_WORKERS = _bounded_int_env("M365_PST_MIN_PREPARE_WORKERS", 1, 1, 8)
PST_MAX_PREPARE_WORKERS = _bounded_int_env("M365_PST_MAX_PREPARE_WORKERS", 4, 1, 8)
PST_COM_SLOW_SECONDS = _bounded_float_env("M365_PST_COM_SLOW_SECONDS", 8.0, 1.0, 120.0)
PST_ADAPTIVE_ENABLED = os.getenv("M365_PST_ADAPTIVE_ENABLED", "1").strip().lower() not in {"0", "false", "no", "off"}

# PST fast resume
PST_FAST_RESUME_ENABLED = os.getenv("M365_PST_FAST_RESUME_ENABLED", "1").strip().lower() not in {"0", "false", "no", "off"}
PST_RESUME_INITIAL_BATCH = _bounded_int_env("M365_PST_RESUME_INITIAL_BATCH", 2, 1, 50)
PST_RESUME_INITIAL_QUEUE = _bounded_int_env("M365_PST_RESUME_INITIAL_QUEUE", 2, 1, 50)
PST_RESUME_QUERY_BATCH = _bounded_int_env("M365_PST_RESUME_QUERY_BATCH", 1000, 50, 10000)
PST_RESUME_TARGET_SECONDS = _bounded_float_env("M365_PST_RESUME_TARGET_SECONDS", 5.0, 1.0, 120.0)
PST_RESUME_FIRST_COMMIT_TARGET_SECONDS = _bounded_float_env("M365_PST_RESUME_FIRST_COMMIT_TARGET_SECONDS", 10.0, 1.0, 120.0)
PST_CAPACITY_PREFLIGHT = os.getenv("M365_PST_CAPACITY_PREFLIGHT", "1").strip().lower() not in {"0", "false", "no", "off"}
PST_RETRY_FAILED_ON_RESUME = os.getenv("M365_PST_RETRY_FAILED_ON_RESUME", "1").strip().lower() not in {"0", "false", "no", "off"}

def validate_settings():
    required_settings = {
        "TENANT_ID": TENANT_ID,
        "CLIENT_ID": CLIENT_ID,
        "CLIENT_SECRET": CLIENT_SECRET,
        "GRAPH_SCOPE": GRAPH_SCOPE,
        "GRAPH_URL": GRAPH_URL,
    }

    missing = [
        key for key, value in required_settings.items()
        if not value
    ]

    if missing:
        raise EnvironmentError(
            "Configurações ausentes no arquivo .env: "
            + ", ".join(missing)
        )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    return True
