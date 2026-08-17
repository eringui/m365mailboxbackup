import argparse
import csv
import json
import os
import sys
import tempfile
import zipfile
import email
import email.policy
import email.utils
import email.header
import html
import hashlib
import gc
import re
import shutil
import time
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

try:
    import psutil
except ImportError:
    psutil = None

try:
    from src.services.checkpoint_store import CheckpointStore
    from src.utils.logger import setup_pst_logger, setup_logger, setup_report_logger
except ImportError:
    from checkpoint_store import CheckpointStore
    from logger import setup_pst_logger, setup_logger, setup_report_logger

_report_logger = setup_report_logger()


def _report(message=""):
    _report_logger.info(message)



class PstCapacityError(RuntimeError):
    """Fatal destination-capacity error. The conversion must stop immediately."""


def safe_regex_group(match, group_name_or_index, default=""):
    if not match:
        return default

    try:
        value = match.group(group_name_or_index)
    except (IndexError, KeyError):
        return default

    if value is None:
        return default

    return value


class PstExportService:
    SOURCE_KEY_DASL = (
        "http://schemas.microsoft.com/mapi/string/"
        "{4D365042-4143-4B55-5053-544B45593130}/M365BackupSourceKey"
    )

    MAPI_PROPS = {
        "message_flags": "http://schemas.microsoft.com/mapi/proptag/0x0E070003",
        "delivery_time": "http://schemas.microsoft.com/mapi/proptag/0x0E060040",
        "client_submit_time": "http://schemas.microsoft.com/mapi/proptag/0x00390040",
        "sender_name": "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F",
        "sender_name_alt": "http://schemas.microsoft.com/mapi/proptag/0x0042001F",
        "sender_email": "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F",
        "sender_email_alt": "http://schemas.microsoft.com/mapi/proptag/0x0065001F",
        "sender_addr_type": "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F",
        "sender_addr_type_alt": "http://schemas.microsoft.com/mapi/proptag/0x0064001F",
        "transport_headers": "http://schemas.microsoft.com/mapi/proptag/0x007D001F",
        "internet_message_id": "http://schemas.microsoft.com/mapi/proptag/0x1035001F",
        "attachment_mime_tag": "http://schemas.microsoft.com/mapi/proptag/0x370E001F",
        "attachment_content_id": "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "attachment_content_location": "http://schemas.microsoft.com/mapi/proptag/0x3713001F",
        "attachment_hidden": "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
        "attachment_flags": "http://schemas.microsoft.com/mapi/proptag/0x37140003",
    }

    def __init__(self, logger=None):
        self.logger = logger
        self.operation_id = os.getenv("M365_OPERATION_ID", "pst-manual")
        self.pst_logger = setup_pst_logger(self.operation_id)
        self.visible_metadata = True
        self.import_attachments = True
        self.image_max_width = 700
        self.folder_mode = "preserve"
        self.root_folder_name = ""
        self.verification_level = "balanced"
        self.verification_batch_size = self._env_int("M365_PST_VERIFICATION_BATCH_SIZE", 25, 1, 500)
        self.verification_final_retries = self._env_int("M365_PST_VERIFICATION_FINAL_RETRIES", 5, 1, 20)
        self._pending_verifications = deque()
        self._verification_metrics = {
            "saved": 0, "verified": 0, "pending": 0, "reconciled": 0,
            "attempts": 0, "audit_failures": 0
        }
        self.checkpoint_batch_size = self._env_int("M365_PST_CHECKPOINT_BATCH_SIZE", 50, 1, 1000)
        self.progress_interval_seconds = self._env_float("M365_PST_PROGRESS_INTERVAL_SECONDS", 1.0, 0.2, 60.0)
        self.gc_interval = self._env_int("M365_PST_GC_INTERVAL", 100, 10, 5000)
        self._last_progress_at = 0.0
        self._last_progress_current = -1
        self._folder_cache = {}
        self.prepare_workers = self._env_int("M365_PST_PREPARE_WORKERS", 3, 1, 8)
        self.prepare_queue_size = self._env_int("M365_PST_PREPARE_QUEUE_SIZE", 12, 2, 100)
        self.large_eml_mb = self._env_int("M365_PST_LARGE_EML_MB", 25, 1, 500)
        self.performance_profile = os.getenv("M365_PST_PERFORMANCE_PROFILE", "balanced").strip().lower()
        self.adaptive_enabled = os.getenv("M365_PST_ADAPTIVE_ENABLED", "1").strip().lower() not in {"0", "false", "no", "off"}
        self.memory_budget_mb = self._env_int("M365_PST_MEMORY_BUDGET_MB", 512, 128, 8192)
        self.min_prepare_workers = self._env_int("M365_PST_MIN_PREPARE_WORKERS", 1, 1, 8)
        self.max_prepare_workers = self._env_int("M365_PST_MAX_PREPARE_WORKERS", self.prepare_workers, 1, 8)
        self.com_slow_seconds = self._env_float("M365_PST_COM_SLOW_SECONDS", 8.0, 1.0, 120.0)
        self._adaptive_queue_limit = self.prepare_queue_size
        self._adaptive_worker_limit = self.prepare_workers
        self._last_adaptive_at = 0.0
        self._com_samples = deque(maxlen=30)
        self._prepare_samples = deque(maxlen=30)
        self.fast_resume_enabled = os.getenv("M365_PST_FAST_RESUME_ENABLED", "1").strip().lower() not in {"0", "false", "no", "off"}
        self.resume_initial_batch = self._env_int("M365_PST_RESUME_INITIAL_BATCH", 2, 1, 50)
        self.resume_initial_queue = self._env_int("M365_PST_RESUME_INITIAL_QUEUE", 2, 1, 50)
        self.resume_query_batch = self._env_int("M365_PST_RESUME_QUERY_BATCH", 1000, 50, 10000)
        self.resume_target_seconds = self._env_float("M365_PST_RESUME_TARGET_SECONDS", 5.0, 1.0, 120.0)
        self.resume_first_commit_target_seconds = self._env_float(
            "M365_PST_RESUME_FIRST_COMMIT_TARGET_SECONDS", 10.0, 1.0, 120.0
        )
        self.capacity_preflight = os.getenv("M365_PST_CAPACITY_PREFLIGHT", "1").strip().lower() not in {"0", "false", "no", "off"}
        self.disk_critical_gb = self._env_float("M365_DISK_CRITICAL_GB", 10, 1, 100000)
        self.retry_failed_on_resume = os.getenv("M365_PST_RETRY_FAILED_ON_RESUME", "1").strip().lower() not in {"0", "false", "no", "off"}
        self._resume_started_at = 0.0
        self._resume_metrics = {
            "checkpoint_seconds": 0.0, "outlook_seconds": 0.0,
            "pending_query_seconds": 0.0, "first_item_seconds": 0.0,
            "first_selected_seconds": 0.0, "first_prepared_seconds": 0.0,
            "first_committed_seconds": 0.0, "first_source_position": 0,
            "first_relative_path": "",
            "first_commit_target_seconds": self.resume_first_commit_target_seconds,
            "first_commit_target_met": False,
            "total_seconds": 0.0, "skipped_before_parse": 0,
            "eligible_items": 0, "capacity_blocked_items": 0,
            "failure_reason": ""
        }
        self._fast_resume_ramp_pending = False
        self._pipeline_metrics = {
            "prepared": 0, "consumed": 0, "queue_depth": 0,
            "large_messages": 0, "prepare_seconds": 0.0,
            "com_seconds": 0.0, "queue_wait_seconds": 0.0,
            "queue_bytes": 0, "peak_queue_bytes": 0, "peak_rss_bytes": 0,
            "adaptive_adjustments": 0, "effective_workers": self.prepare_workers,
            "effective_queue_limit": self.prepare_queue_size, "bottleneck": "calculando",
            "eta_seconds": 0.0, "memory_pressure": False
        }

    @staticmethod
    def _env_int(name, default, minimum, maximum):
        try:
            value = int(str(os.getenv(name, default)).strip())
        except (TypeError, ValueError):
            value = int(default)
        return max(minimum, min(maximum, value))

    @staticmethod
    def _env_float(name, default, minimum, maximum):
        try:
            value = float(str(os.getenv(name, default)).strip())
        except (TypeError, ValueError):
            value = float(default)
        return max(minimum, min(maximum, value))

    def log_debug(self, message):
        self.pst_logger.debug(message)
        if self.logger and hasattr(self.logger, "debug"):
            self.logger.debug(message)


    def log_info(self, message):
        self.pst_logger.info(message)
        if self.logger:
            self.logger.info(message)

    def log_error(self, message):
        self.pst_logger.error(message)
        if self.logger:
            self.logger.error(message)

    def log_stage(self, category, message, **details):
        payload = {"category": category, "message": message, **details}
        structured = "[PST-EVENT] " + json.dumps(
            payload, ensure_ascii=False, separators=(",", ":")
        )
        if category in {"eml", "mapi", "attachment", "pst", "verification", "checkpoint", "success"}:
            self.log_debug(structured)
            return
        self.log_info(structured)
        self.log_info(f"[PST-STAGE] {message}")

    def ensure_disk_capacity(self, pst_path):
        """Abort early when free disk space drops below the critical threshold.

        M365_PST_CAPACITY_PREFLIGHT existed as a flag but had no actual disk
        check anywhere; PstCapacityError plugs into the same checkpoint-preserving
        abort path already used for "PST reached its maximum size" COM errors.
        """
        if not self.capacity_preflight:
            return
        try:
            usage = shutil.disk_usage(str(Path(pst_path).resolve().parent))
        except OSError:
            return
        free_gb = usage.free / (1024 ** 3)
        if free_gb < self.disk_critical_gb:
            raise PstCapacityError(
                f"Espaço livre em disco insuficiente para continuar a gravação do "
                f"PST: {free_gb:.1f} GB livres, limite crítico configurado é "
                f"{self.disk_critical_gb:.0f} GB (M365_DISK_CRITICAL_GB)."
            )

    def acquire_destination_lock(self, pst_path):
        pst_path = self.normalize_path(pst_path)
        lock_path = pst_path.with_suffix(pst_path.suffix + ".conversion.lock")
        lock_path.parent.mkdir(parents=True, exist_ok=True)
        for attempt in range(2):
            try:
                descriptor = os.open(
                    str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY
                )
                with os.fdopen(descriptor, "w", encoding="utf-8") as lock_file:
                    lock_file.write(json.dumps({
                        "pid": os.getpid(),
                        "pst_path": str(pst_path),
                        "created_at": datetime.now().isoformat()
                    }, ensure_ascii=False))
                return lock_path
            except FileExistsError:
                stale = False
                try:
                    data = self.load_json_safely(lock_path, {})
                    pid = int(data.get("pid", 0) or 0)
                    if pid <= 0:
                        stale = True
                    elif psutil is not None and not psutil.pid_exists(pid):
                        # No Windows, os.kill(pid, 0) não é uma verificação
                        # inofensiva: ele chama TerminateProcess(pid, 0) e
                        # mataria de fato um processo real com esse PID.
                        # psutil.pid_exists() apenas consulta, sem matar nada.
                        stale = True
                except Exception:
                    stale = True
                if stale and attempt == 0:
                    try:
                        lock_path.unlink(missing_ok=True)
                    except Exception:
                        pass
                    continue
                raise RuntimeError(
                    "Este PST já está sendo processado por outra conversão: "
                    f"{pst_path}"
                )
        raise RuntimeError(f"Não foi possível bloquear o destino PST: {pst_path}")

    def release_destination_lock(self, lock_path):
        if lock_path is None:
            return
        try:
            Path(lock_path).unlink(missing_ok=True)
        except Exception:
            pass

    def normalize_path(self, path_value):
        return Path(path_value).expanduser().resolve()

    def get_checkpoint_path(self, pst_path):
        return self.normalize_path(pst_path).with_suffix(".pst_checkpoint.json")

    def load_json_safely(self, path, default=None):
        path = Path(path)
        if not path.exists():
            return {} if default is None else default
        try:
            with open(path, "r", encoding="utf-8") as file:
                return json.load(file)
        except Exception:
            return {} if default is None else default

    def save_json_atomic(self, path, data):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_suffix(path.suffix + ".tmp")
        with open(temporary, "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=2)
            file.flush()
            os.fsync(file.fileno())
        os.replace(temporary, path)

    def build_eml_key(self, mail_folders_root, eml_file):
        relative = Path(eml_file).relative_to(Path(mail_folders_root))
        stat = Path(eml_file).stat()
        return relative.as_posix(), {
            "relative_path": relative.as_posix(),
            "size": int(stat.st_size),
            "mtime_ns": int(stat.st_mtime_ns)
        }

    def get_eml_import_rate_per_second(self):
        try:
            raw_value = os.getenv("M365_EML_IMPORT_RATE_PER_SECOND")
            if raw_value is None:
                raw_value = os.getenv("M365_PST_EML_IMPORT_RATE_PER_SECOND")
            if raw_value is None:
                return 10
            value = float(str(raw_value).strip())
            if value <= 0:
                return 10
            return value
        except Exception:
            return 10

    def enforce_eml_rate_limit(self, last_processed_at, rate_per_second):
        if not rate_per_second or rate_per_second <= 0:
            return time.monotonic()
        interval = 1.0 / float(rate_per_second)
        now = time.monotonic()
        if last_processed_at is not None:
            elapsed = now - last_processed_at
            if elapsed < interval:
                time.sleep(interval - elapsed)
        return time.monotonic()

    def load_or_create_pst_checkpoint(
        self,
        backup_root,
        pst_path,
        total_eml
    ):
        checkpoint_path = self.get_checkpoint_path(pst_path)
        normalized_backup = str(self.normalize_path(backup_root))
        normalized_pst = str(self.normalize_path(pst_path))
        checkpoint = self.load_json_safely(checkpoint_path, {})

        compatible = (
            isinstance(checkpoint, dict)
            and checkpoint.get("backup_root") == normalized_backup
            and checkpoint.get("pst_path") == normalized_pst
        )
        if not compatible:
            checkpoint = {
                "version": 1,
                "backup_root": normalized_backup,
                "pst_path": normalized_pst,
                "created_at": datetime.now().isoformat(),
                "updated_at": datetime.now().isoformat(),
                "status": "running",
                "total_eml": int(total_eml),
                "imported_eml": {},
                "failed_eml": {}
            }
        else:
            checkpoint.setdefault("imported_eml", {})
            checkpoint.setdefault("failed_eml", {})
            checkpoint["total_eml"] = int(total_eml)
            checkpoint["status"] = "running"
            checkpoint["updated_at"] = datetime.now().isoformat()

        self.save_json_atomic(checkpoint_path, checkpoint)
        return checkpoint_path, checkpoint

    def emit_pst_progress(self, checkpoint_path, checkpoint):
        imported = checkpoint.get("imported_eml", {})
        payload = {
            "current": len(imported),
            "total": int(checkpoint.get("total_eml", 0) or 0),
            "failed": len(checkpoint.get("failed_eml", {})),
            "status": checkpoint.get("status"),
            "checkpoint_path": str(checkpoint_path)
        }
        self.log_info(
            "[PST-PROGRESS] " + json.dumps(
                payload,
                ensure_ascii=False,
                separators=(",", ":")
            )
        )

    def write_pst_checkpoint_summary(self, path, store, total_eml):
        counts = store.counts()
        operation = store.get_operation()
        summary = {
            "version": 2,
            "engine": "sqlite-wal+mapi-idempotency",
            "backup_root": operation.get("source_root"),
            "pst_path": operation.get("destination_path"),
            "status": operation.get("status"),
            "total_eml": int(total_eml),
            "imported_count": counts["completed"],
            "failed_count": counts["failed"],
            "processing_count": counts["processing"],
            "sqlite_path": str(store.path),
            "updated_at": datetime.now().isoformat()
        }
        self.save_json_atomic(path, summary)

    def emit_sqlite_pst_progress(self, sqlite_path, store, total_eml):
        counts = store.counts()
        payload = {
            "current": counts["completed"],
            "total": int(total_eml),
            "failed": counts["failed"],
            "processing": counts["processing"],
            "status": store.get_operation().get("status"),
            "checkpoint_path": str(sqlite_path),
            "folder": (
                f"Preparação MIME: {self._pipeline_metrics.get('prepared', 0)} preparados · "
                f"fila {self._pipeline_metrics.get('queue_depth', 0)}"
            ),
            "page": int(self._pipeline_metrics.get("queue_depth", 0) or 0),
            "pipeline": {
                **self._pipeline_metrics,
                "workers": self.prepare_workers,
                "queue_limit": self.prepare_queue_size
            },
            "verification": {
                **self._verification_metrics,
                "mode": self.verification_level,
                "batch_size": self.verification_batch_size
            },
            "resume": dict(self._resume_metrics)
        }
        self.log_info(
            "[PST-PROGRESS] " + json.dumps(
                payload,
                ensure_ascii=False,
                separators=(",", ":")
            )
        )

    def extract_zip_if_needed(self, source_path):
        source_path = self.normalize_path(source_path)

        if source_path.is_dir():
            return source_path, None

        if not source_path.exists():
            raise FileNotFoundError(
                f"Caminho não encontrado: {source_path}"
            )

        if source_path.suffix.lower() != ".zip":
            raise ValueError(
                "O caminho informado precisa ser uma pasta de backup ou um arquivo .zip."
            )

        temp_dir = tempfile.TemporaryDirectory()
        extract_to = Path(temp_dir.name)

        self.log_info(
            f"Extraindo ZIP para pasta temporária: {extract_to}"
        )

        with zipfile.ZipFile(source_path, "r") as zip_file:
            zip_file.extractall(extract_to)

        return extract_to, temp_dir

    def find_mail_folders_root(self, backup_root):
        backup_root = self.normalize_path(backup_root)

        direct_path = backup_root / "mail" / "folders"

        if direct_path.exists():
            return direct_path

        candidates = []

        for path in backup_root.rglob("mail"):
            folders_path = path / "folders"

            if folders_path.exists():
                candidates.append(folders_path)

        if candidates:
            return candidates[0]

        raise FileNotFoundError(
            f"Não foi encontrada a estrutura mail/folders dentro de: {backup_root}"
        )

    def find_eml_files(self, backup_root):
        mail_folders_root = self.find_mail_folders_root(backup_root)

        eml_files = []

        for eml_file in mail_folders_root.rglob("*.eml"):
            if eml_file.is_file():
                eml_files.append(eml_file)

        return sorted(eml_files), mail_folders_root

    def get_pst_folder_parts(self, mail_folders_root, eml_file):
        mail_folders_root = Path(mail_folders_root)
        eml_file = Path(eml_file)

        try:
            relative = eml_file.relative_to(mail_folders_root)
        except ValueError:
            return ["Imported EML"]

        parts = list(relative.parts)
        folder_parts = []

        for part in parts[:-1]:
            if part.lower() == "eml":
                break

            clean_part = str(part).strip()

            if clean_part:
                folder_parts.append(clean_part)

        if not folder_parts:
            folder_parts = ["Imported EML"]
        if self.folder_mode == "single":
            folder_parts = [self.root_folder_name or "Todos os E-mails"]
        elif self.root_folder_name:
            folder_parts = [self.root_folder_name] + folder_parts
        return folder_parts

    def sanitize_outlook_folder_name(self, folder_name):
        folder_name = str(folder_name or "").strip()

        if not folder_name:
            folder_name = "sem_nome"

        invalid_chars = [
            "\\",
            "/",
            ":",
            "*",
            "?",
            '"',
            "<",
            ">",
            "|"
        ]

        for invalid_char in invalid_chars:
            folder_name = folder_name.replace(invalid_char, "_")

        folder_name = re.sub(
            r"\s+",
            " ",
            folder_name
        ).strip()

        if not folder_name:
            folder_name = "sem_nome"

        return folder_name[:240]

    def get_or_create_outlook_folder(self, root_folder, folder_parts):
        normalized = tuple(self.sanitize_outlook_folder_name(part) for part in folder_parts)
        full_key = "/".join(part.casefold() for part in normalized)
        cached = self._folder_cache.get(full_key)
        if cached is not None:
            try:
                _ = cached.Name
                return cached
            except Exception:
                self._folder_cache.pop(full_key, None)
        current_folder = root_folder
        traversed = []
        for safe_name in normalized:
            traversed.append(safe_name.casefold())
            partial_key = "/".join(traversed)
            cached = self._folder_cache.get(partial_key)
            if cached is not None:
                try:
                    _ = cached.Name
                    current_folder = cached
                    continue
                except Exception:
                    self._folder_cache.pop(partial_key, None)
            found_folder = None
            try:
                for existing_folder in current_folder.Folders:
                    if str(existing_folder.Name).casefold() == safe_name.casefold():
                        found_folder = existing_folder
                        break
            except Exception:
                found_folder = None
            if found_folder is None:
                found_folder = current_folder.Folders.Add(safe_name)
            current_folder = found_folder
            self._folder_cache[partial_key] = current_folder
        self._folder_cache[full_key] = current_folder
        return current_folder

    def create_or_attach_pst(self, namespace, pst_path, display_name=None):
        pst_path = self.normalize_path(pst_path)
        pst_path.parent.mkdir(parents=True, exist_ok=True)

        for store in namespace.Stores:
            try:
                store_path = getattr(store, "FilePath", None)

                if store_path and Path(store_path).resolve() == pst_path:
                    pst_root = store.GetRootFolder()

                    if display_name:
                        try:
                            pst_root.Name = display_name
                        except Exception:
                            pass

                    return pst_root
            except Exception:
                continue

        self.log_info(
            f"Criando/anexando PST: {pst_path}"
        )

        try:
            namespace.AddStoreEx(str(pst_path), 2)
        except Exception as error:
            raise RuntimeError(
                f"Não foi possível criar o PST Unicode: {error}"
            ) from error

        pst_root = None

        for store in namespace.Stores:
            try:
                store_path = getattr(store, "FilePath", None)

                if store_path and Path(store_path).resolve() == pst_path:
                    pst_root = store.GetRootFolder()
                    break
            except Exception:
                continue

        if pst_root is None:
            raise RuntimeError(
                "O PST foi solicitado ao Outlook, mas não foi possível localizar o store criado."
            )

        if display_name:
            try:
                pst_root.Name = display_name
            except Exception:
                pass

        return pst_root

    def decode_header_value(self, value):
        if not value:
            return ""

        try:
            return str(
                email.header.make_header(
                    email.header.decode_header(value)
                )
            )
        except Exception:
            return str(value)

    def parse_recipient_header(self, msg, header_name):
        header_value = msg.get(header_name)

        if not header_value:
            return ""

        try:
            addresses = email.utils.getaddresses([header_value])
        except Exception:
            return self.decode_header_value(header_value)

        result = []

        for name, address in addresses:
            if address:
                result.append(address)
            elif name:
                result.append(self.decode_header_value(name))

        return "; ".join(result)

    def get_message_body(self, msg):
        html_body = ""
        text_body = ""

        try:
            html_part = msg.get_body(
                preferencelist=("html",)
            )

            if html_part:
                html_body = html_part.get_content()
        except Exception:
            html_body = ""

        try:
            text_part = msg.get_body(
                preferencelist=("plain",)
            )

            if text_part:
                text_body = text_part.get_content()
        except Exception:
            text_body = ""

        if html_body or text_body:
            return text_body or "", html_body or ""

        try:
            if msg.is_multipart():
                for part in msg.walk():
                    if part.is_multipart():
                        continue

                    content_type = part.get_content_type()
                    disposition = str(
                        part.get("Content-Disposition", "")
                    ).lower()

                    if "attachment" in disposition:
                        continue

                    payload = part.get_payload(decode=True)

                    if not payload:
                        continue

                    charset = part.get_content_charset() or "utf-8"

                    decoded = payload.decode(
                        charset,
                        errors="replace"
                    )

                    if content_type == "text/html" and not html_body:
                        html_body = decoded

                    if content_type == "text/plain" and not text_body:
                        text_body = decoded

            else:
                payload = msg.get_payload(decode=True)

                if payload:
                    charset = msg.get_content_charset() or "utf-8"
                    text_body = payload.decode(
                        charset,
                        errors="replace"
                    )
        except Exception:
            pass

        return text_body or "", html_body or ""

    def build_visible_metadata_header(self, msg):
        from_value = self.decode_header_value(
            msg.get("From", "")
        )

        to_value = self.decode_header_value(
            msg.get("To", "")
        )

        cc_value = self.decode_header_value(
            msg.get("Cc", "")
        )

        date_value = self.decode_header_value(
            msg.get("Date", "")
        )

        subject_value = self.decode_header_value(
            msg.get("Subject", "")
        )

        lines = [
            "----- Metadados originais do EML -----",
            f"De: {from_value}",
            f"Para: {to_value}",
            f"Cc: {cc_value}",
            f"Data: {date_value}",
            f"Assunto: {subject_value}",
            "--------------------------------------",
            ""
        ]

        return "\n".join(lines)

    def build_visible_metadata_header_html(self, msg):
        text_header = self.build_visible_metadata_header(msg)

        return (
            "<div style='font-family: Segoe UI, Arial, sans-serif; "
            "font-size: 10pt; border: 1px solid #dddddd; "
            "padding: 8px; margin-bottom: 10px; background: #f7f7f7;'>"
            "<pre style='white-space: pre-wrap; margin: 0;'>"
            + html.escape(text_header)
            + "</pre></div>"
        )

    def make_html_images_proportional(self, html_body):
        if not html_body:
            return html_body

        def update_img_tag(match):
            tag = match.group(0)

            tag = re.sub(
                r'\swidth=["\']?\d+["\']?',
                "",
                tag,
                flags=re.IGNORECASE
            )

            tag = re.sub(
                r'\sheight=["\']?\d+["\']?',
                "",
                tag,
                flags=re.IGNORECASE
            )

            style_match = re.search(
                r'style=(["\'])(.*?)\1',
                tag,
                flags=re.IGNORECASE | re.DOTALL
            )

            proportional_style = (
                f"max-width: {self.image_max_width}px; "
                "width: auto; "
                "height: auto; "
                "object-fit: contain;"
            )

            if style_match:
                current_style = safe_regex_group(style_match, 2, "").strip()

                if current_style and not current_style.endswith(";"):
                    current_style += ";"

                new_style = f'style="{current_style} {proportional_style}"'

                tag = re.sub(
                    r'style=(["\']).*?\1',
                    new_style,
                    tag,
                    count=1,
                    flags=re.IGNORECASE | re.DOTALL
                )
            else:
                if tag.endswith("/>"):
                    tag = tag[:-2] + f' style="{proportional_style}" />'
                else:
                    tag = tag[:-1] + f' style="{proportional_style}">'

            return tag

        return re.sub(
            r"<img\b[^>]*>",
            update_img_tag,
            html_body,
            flags=re.IGNORECASE
        )

    def wrap_html_body_for_outlook(self, msg, html_body):
        if not html_body:
            html_body = ""

        html_body = self.make_html_images_proportional(
            html_body
        )

        metadata_html = (
            self.build_visible_metadata_header_html(msg) if self.visible_metadata else ""
        )

        # CSS dentro de f-string exige chaves literais duplicadas.
        # Mantemos apenas image_max_width como interpolação Python.
        css = f"""
<style>
    body {{
        font-family: Segoe UI, Arial, sans-serif;
        font-size: 10pt;
    }}

    img {{
        max-width: {self.image_max_width}px !important;
        width: auto !important;
        height: auto !important;
        object-fit: contain !important;
    }}

    table {{
        max-width: 100%;
    }}
</style>
"""

        if "<head" in html_body.lower():
            html_body = re.sub(
                r"(<head[^>]*>)",
                r"\1" + css,
                html_body,
                count=1,
                flags=re.IGNORECASE
            )
        else:
            html_body = css + html_body

        if "<body" in html_body.lower():
            html_body = re.sub(
                r"(<body[^>]*>)",
                r"\1" + metadata_html,
                html_body,
                count=1,
                flags=re.IGNORECASE
            )
        else:
            html_body = metadata_html + html_body

        return html_body

    def sanitize_attachment_name(self, filename, index):
        if not filename:
            filename = f"attachment_{index}.bin"

        filename = self.decode_header_value(filename)

        invalid_chars = [
            "\\",
            "/",
            ":",
            "*",
            "?",
            '"',
            "<",
            ">",
            "|"
        ]

        for invalid_char in invalid_chars:
            filename = filename.replace(invalid_char, "_")

        filename = re.sub(
            r"\s+",
            " ",
            filename
        ).strip()

        if not filename:
            filename = f"attachment_{index}.bin"

        return filename[:180]

    def get_unique_temp_file_path(self, temp_dir, filename):
        temp_dir = Path(temp_dir)
        candidate = temp_dir / filename

        if not candidate.exists():
            return candidate

        stem = candidate.stem
        suffix = candidate.suffix

        for index in range(1, 10000):
            candidate = temp_dir / f"{stem}_{index}{suffix}"

            if not candidate.exists():
                return candidate

        return temp_dir / f"{stem}_{datetime.now().timestamp()}{suffix}"

    def extract_attachments_to_temp(self, msg, temp_dir):
        attachments = []
        index = 0

        try:
            for part in msg.walk():
                if part.is_multipart():
                    continue

                filename = part.get_filename()
                content_type = part.get_content_type()
                content_id = part.get("Content-ID", "")
                content_location = part.get("Content-Location", "")
                disposition = str(
                    part.get("Content-Disposition", "")
                ).lower()

                is_inline = (
                    "inline" in disposition
                    or bool(content_id)
                )

                is_attachment = (
                    filename is not None
                    or "attachment" in disposition
                    or is_inline
                )

                if not is_attachment:
                    continue

                payload = part.get_payload(decode=True)

                if payload is None:
                    continue

                index += 1

                safe_filename = self.sanitize_attachment_name(
                    filename,
                    index
                )

                attachment_path = self.get_unique_temp_file_path(
                    temp_dir,
                    safe_filename
                )

                with open(attachment_path, "wb") as file:
                    file.write(payload)

                clean_content_id = str(content_id or "").strip()

                if clean_content_id.startswith("<") and clean_content_id.endswith(">"):
                    clean_content_id = clean_content_id[1:-1]

                attachments.append(
                    {
                        "path": attachment_path,
                        "filename": attachment_path.name,
                        "content_type": content_type,
                        "content_id": clean_content_id,
                        "content_location": content_location,
                        "is_inline": is_inline
                    }
                )
        except Exception:
            pass

        return attachments

    def set_mapi_property_safely(self, accessor, prop_name, value):
        try:
            accessor.SetProperty(
                prop_name,
                value
            )
        except Exception:
            pass

    def apply_mapi_properties_to_received_mail(self, mail_item, msg):
        try:
            accessor = mail_item.PropertyAccessor
        except Exception:
            return

        self.set_mapi_property_safely(
            accessor,
            self.MAPI_PROPS["message_flags"],
            1
        )

        try:
            date_header = msg.get("Date")

            if date_header:
                parsed_date = email.utils.parsedate_to_datetime(
                    date_header
                )

                if parsed_date.tzinfo is not None:
                    parsed_date = parsed_date.astimezone().replace(
                        tzinfo=None
                    )

                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["delivery_time"],
                    parsed_date
                )

                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["client_submit_time"],
                    parsed_date
                )
        except Exception:
            pass

        try:
            from_header = msg.get("From", "")
            sender_name, sender_email = email.utils.parseaddr(
                from_header
            )

            sender_name = self.decode_header_value(sender_name or sender_email)
            sender_email = sender_email or ""

            if sender_name:
                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["sender_name"],
                    sender_name
                )

                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["sender_name_alt"],
                    sender_name
                )

            if sender_email:
                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["sender_email"],
                    sender_email
                )

                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["sender_email_alt"],
                    sender_email
                )

                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["sender_addr_type"],
                    "SMTP"
                )

                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["sender_addr_type_alt"],
                    "SMTP"
                )
        except Exception:
            pass

        try:
            headers = ""

            for key, value in msg.items():
                headers += f"{key}: {value}\r\n"

            if headers:
                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["transport_headers"],
                    headers
                )
        except Exception:
            pass

        try:
            message_id = msg.get("Message-ID")

            if message_id:
                self.set_mapi_property_safely(
                    accessor,
                    self.MAPI_PROPS["internet_message_id"],
                    message_id
                )
        except Exception:
            pass

    def attach_files_to_mail_item(self, mail_item, attachments):
        for attachment in attachments:
            try:
                added_attachment = mail_item.Attachments.Add(
                    str(attachment["path"])
                )

                try:
                    attachment_accessor = added_attachment.PropertyAccessor

                    if attachment.get("content_type"):
                        self.set_mapi_property_safely(
                            attachment_accessor,
                            self.MAPI_PROPS["attachment_mime_tag"],
                            attachment["content_type"]
                        )

                    if attachment.get("content_id"):
                        self.set_mapi_property_safely(
                            attachment_accessor,
                            self.MAPI_PROPS["attachment_content_id"],
                            attachment["content_id"]
                        )

                    if attachment.get("content_location"):
                        self.set_mapi_property_safely(
                            attachment_accessor,
                            self.MAPI_PROPS["attachment_content_location"],
                            attachment["content_location"]
                        )

                    if attachment.get("is_inline"):
                        self.set_mapi_property_safely(
                            attachment_accessor,
                            self.MAPI_PROPS["attachment_hidden"],
                            True
                        )

                        self.set_mapi_property_safely(
                            attachment_accessor,
                            self.MAPI_PROPS["attachment_flags"],
                            4
                        )
                except Exception as metadata_error:
                    self.log_error(
                        "[PST-EVENT] " + json.dumps(
                            {
                                "category": "attachment_metadata_error",
                                "message": "Anexo adicionado, mas metadados MAPI não aplicados",
                                "attachment": str(attachment.get("path")),
                                "error": str(metadata_error)
                            },
                            ensure_ascii=False
                        )
                    )

            except Exception as attachment_error:
                self.log_error(
                    "[PST-EVENT] " + json.dumps(
                        {
                            "category": "attachment_error",
                            "message": "Falha ao anexar arquivo ao item do PST; anexo não incluído",
                            "attachment": str(attachment.get("path")),
                            "error": str(attachment_error)
                        },
                        ensure_ascii=False
                    )
                )

    def build_source_key(self, backup_root, relative_path, file_size, modified_ns):
        payload = (
            f"{self.normalize_path(backup_root)}|{relative_path}|"
            f"{int(file_size)}|{int(modified_ns)}"
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def set_source_key(self, mail_item, source_key):
        accessor = mail_item.PropertyAccessor
        accessor.SetProperty(self.SOURCE_KEY_DASL, str(source_key))

    def folder_contains_source_key(self, target_folder, source_key):
        try:
            for item in target_folder.Items:
                try:
                    value = item.PropertyAccessor.GetProperty(self.SOURCE_KEY_DASL)
                    if str(value) == str(source_key):
                        return True
                except Exception:
                    continue
        except Exception:
            return False
        return False

    def find_item_by_source_key(self, target_folder, source_key):
        try:
            for item in target_folder.Items:
                try:
                    value = item.PropertyAccessor.GetProperty(self.SOURCE_KEY_DASL)
                    if str(value) == str(source_key):
                        return {
                            "entry_id": str(getattr(item, "EntryID", "") or ""),
                            "store_id": str(getattr(item.Parent, "StoreID", "") or ""),
                            "folder_name": str(getattr(target_folder, "Name", "") or ""),
                            "verified": True,
                        }
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def _discard_unsaved_or_failed_item(self, mail_item):
        if mail_item is None:
            return
        try:
            mail_item.Delete()
        except Exception:
            pass

    def _verify_item_in_target_pst(
        self, namespace, target_folder, mail_item, source_key
    ):
        entry_id = str(getattr(mail_item, "EntryID", "") or "")
        target_store_id = str(getattr(target_folder, "StoreID", "") or "")
        item_store_id = str(getattr(mail_item.Parent, "StoreID", "") or "")
        if not entry_id:
            raise RuntimeError("O Outlook não retornou EntryID após salvar no PST.")
        if not target_store_id or item_store_id != target_store_id:
            raise RuntimeError(
                "O item foi salvo fora do store PST de destino."
            )
        confirmed = namespace.GetItemFromID(entry_id, target_store_id)
        confirmed_parent_store = str(
            getattr(confirmed.Parent, "StoreID", "") or ""
        )
        if confirmed_parent_store != target_store_id:
            raise RuntimeError("A confirmação MAPI apontou para outro store.")
        confirmed_key = confirmed.PropertyAccessor.GetProperty(
            self.SOURCE_KEY_DASL
        )
        if str(confirmed_key) != str(source_key):
            raise RuntimeError("A chave de origem não foi confirmada no PST.")
        return {
            "entry_id": entry_id,
            "store_id": target_store_id,
            "folder_name": str(getattr(target_folder, "Name", "") or "")
        }

    def _verify_pst_item(self, namespace, target_folder, entry_id, store_id, source_key, final=False):
        entry_id = str(entry_id or "")
        store_id = str(store_id or "")
        target_store_id = str(getattr(target_folder, "StoreID", "") or "")
        if not entry_id or not store_id:
            return {"entry_id": entry_id, "store_id": store_id, "verified": False,
                    "verification_error": "EntryID ou StoreID ausente."}
        if target_store_id and store_id != target_store_id:
            return {"entry_id": entry_id, "store_id": store_id, "verified": False,
                    "wrong_store": True, "verification_error": "Item salvo fora do PST de destino."}
        if self.verification_level == "quick" and not final:
            return {"entry_id": entry_id, "store_id": store_id, "verified": False,
                    "pending": True, "verification_status": "saved"}
        attempts = self.verification_final_retries if (final or self.verification_level == "complete") else 1
        last_error = None
        for attempt in range(1, attempts + 1):
            self._verification_metrics["attempts"] += 1
            try:
                confirmed = namespace.GetItemFromID(entry_id, store_id)
                confirmed_key = confirmed.PropertyAccessor.GetProperty(self.SOURCE_KEY_DASL)
                if str(confirmed_key) == str(source_key):
                    return {"entry_id": entry_id, "store_id": store_id, "verified": True,
                            "verification_status": "verified"}
                last_error = RuntimeError("Chave de origem divergente.")
            except Exception as error:
                last_error = error
            if attempt < attempts:
                time.sleep(min(0.05 * attempt, 0.25))
        if final and self.folder_contains_source_key(target_folder, source_key):
            return {"entry_id": entry_id, "store_id": store_id, "verified": True,
                    "verification_status": "reconciled"}
        return {"entry_id": entry_id, "store_id": store_id, "verified": False,
                "pending": not final,
                "verification_status": "pending" if not final else "failed",
                "verification_error": str(last_error or "Confirmação MAPI indisponível.")}

    def _queue_verification(self, pending):
        self._pending_verifications.append(pending)
        self._verification_metrics["pending"] = len(self._pending_verifications)

    def drain_pending_verifications(self, namespace, store, result, final=False):
        if not self._pending_verifications:
            return 0
        limit = len(self._pending_verifications) if final else min(
            self.verification_batch_size, len(self._pending_verifications)
        )
        verified_now = 0
        remaining = deque()
        for _ in range(limit):
            pending = self._pending_verifications.popleft()
            identity = self._verify_pst_item(
                namespace, pending["target_folder"], pending["entry_id"],
                pending["store_id"], pending["source_key"], final=final
            )
            if identity.get("verified"):
                status = identity.get("verification_status", "verified")
                store.mark_completed(
                    pending["source_key"], relative_path=pending["relative_path"],
                    file_size=pending["file_size"], modified_ns=pending["modified_ns"],
                    output_path=pending["output_path"], source_key=pending["source_key"],
                    destination_entry_id=identity.get("entry_id"),
                    destination_store_id=identity.get("store_id"),
                    destination_folder=pending["destination_folder"],
                    verification_status=status
                )
                result["eml_imported"] += 1
                verified_now += 1
                self._verification_metrics["verified"] += 1
                if status == "reconciled":
                    self._verification_metrics["reconciled"] += 1
            elif final:
                error = identity.get("verification_error") or "Auditoria final não confirmou o item."
                store.mark_failed(
                    pending["source_key"], error, relative_path=pending["relative_path"],
                    file_size=pending["file_size"], modified_ns=pending["modified_ns"],
                    output_path=pending["output_path"], source_key=pending["source_key"],
                    destination_entry_id=pending["entry_id"],
                    destination_store_id=pending["store_id"],
                    destination_folder=pending["destination_folder"],
                    verification_status="failed"
                )
                result["eml_failed"] += 1
                self._verification_metrics["audit_failures"] += 1
                result["errors"].append(
                    f"Auditoria MAPI falhou: {pending['relative_path']} | {error}"
                )
            else:
                remaining.append(pending)
        while self._pending_verifications:
            remaining.append(self._pending_verifications.popleft())
        self._pending_verifications = remaining
        self._verification_metrics["pending"] = len(remaining)
        return verified_now

    def apply_performance_profile(self):
        profiles = {
            "conservative": {"workers": 2, "queue": 6, "memory": 256},
            "balanced": {"workers": 4, "queue": 12, "memory": 512},
            "performance": {"workers": 6, "queue": 24, "memory": 1024},
        }
        preset = profiles.get(self.performance_profile)
        if preset:
            self.max_prepare_workers = min(self.max_prepare_workers, preset["workers"])
            self.prepare_queue_size = min(self.prepare_queue_size, preset["queue"])
            self.memory_budget_mb = min(self.memory_budget_mb, preset["memory"])
        self._adaptive_worker_limit = max(self.min_prepare_workers, min(self.prepare_workers, self.max_prepare_workers))
        self._adaptive_queue_limit = self.prepare_queue_size

    def _memory_snapshot(self):
        rss = 0
        available = 0
        percent = 0.0
        if psutil is not None:
            try:
                rss = int(psutil.Process(os.getpid()).memory_info().rss)
                vm = psutil.virtual_memory()
                available = int(vm.available)
                percent = float(vm.percent)
            except Exception:
                pass
        self._pipeline_metrics["peak_rss_bytes"] = max(
            int(self._pipeline_metrics.get("peak_rss_bytes", 0) or 0), rss
        )
        return rss, available, percent

    def _update_adaptive_limits(self, queue_depth=0, queue_bytes=0):
        rss, available, system_percent = self._memory_snapshot()
        budget = self.memory_budget_mb * 1024 * 1024
        pressure = queue_bytes >= budget * 0.85 or system_percent >= 88.0 or (available and available < 512 * 1024 * 1024)
        self._pipeline_metrics["memory_pressure"] = bool(pressure)
        now = time.monotonic()
        if not self.adaptive_enabled or now - self._last_adaptive_at < 1.0:
            return
        old = (self._adaptive_worker_limit, self._adaptive_queue_limit)
        avg_com = sum(self._com_samples) / len(self._com_samples) if self._com_samples else 0.0
        avg_prepare = sum(self._prepare_samples) / len(self._prepare_samples) if self._prepare_samples else 0.0
        if pressure or avg_com >= self.com_slow_seconds:
            self._adaptive_worker_limit = max(self.min_prepare_workers, self._adaptive_worker_limit - 1)
            self._adaptive_queue_limit = max(2, self._adaptive_queue_limit - 2)
        elif queue_depth <= 1 and avg_prepare > avg_com and self._adaptive_worker_limit < self.max_prepare_workers:
            self._adaptive_worker_limit += 1
            self._adaptive_queue_limit = min(self.prepare_queue_size, self._adaptive_queue_limit + 1)
        elif queue_depth >= max(2, self._adaptive_queue_limit - 1) and avg_com > avg_prepare * 1.5:
            self._adaptive_queue_limit = max(2, self._adaptive_queue_limit - 1)
        if old != (self._adaptive_worker_limit, self._adaptive_queue_limit):
            self._pipeline_metrics["adaptive_adjustments"] += 1
            self._last_adaptive_at = now
        self._pipeline_metrics["effective_workers"] = self._adaptive_worker_limit
        self._pipeline_metrics["effective_queue_limit"] = self._adaptive_queue_limit
        if pressure:
            bottleneck = "memória"
        elif avg_com > max(avg_prepare * 1.5, 0.25):
            bottleneck = "Outlook COM"
        elif avg_prepare > max(avg_com * 1.5, 0.25):
            bottleneck = "parsing MIME"
        elif queue_depth >= self._adaptive_queue_limit:
            bottleneck = "backpressure"
        else:
            bottleneck = "equilibrado"
        self._pipeline_metrics["bottleneck"] = bottleneck

    def record_com_timing(self, elapsed):
        elapsed = max(0.0, float(elapsed or 0.0))
        self._com_samples.append(elapsed)
        self._pipeline_metrics["com_seconds"] += elapsed
        if elapsed >= self.com_slow_seconds:
            self.log_info("[PST-EVENT] " + json.dumps({
                "category": "slow_com", "message": "Operação COM lenta detectada; backpressure reforçado.",
                "seconds": round(elapsed, 3)
            }, ensure_ascii=False, separators=(",", ":")))

    def performance_recommendations(self):
        metrics = self._pipeline_metrics
        recommendations = []
        if metrics.get("memory_pressure") or metrics.get("peak_queue_bytes", 0) >= self.memory_budget_mb * 1024 * 1024 * 0.85:
            recommendations.append("A fila se aproximou do orçamento de memória; mantenha ou reduza fila e workers.")
        if metrics.get("bottleneck") == "Outlook COM":
            recommendations.append("O Outlook foi o gargalo principal; aumentar workers MIME não deve produzir ganho relevante.")
        elif metrics.get("bottleneck") == "parsing MIME":
            recommendations.append("O parsing MIME foi o gargalo; o perfil Desempenho pode ajudar se houver memória disponível.")
        if int(metrics.get("large_messages", 0) or 0) > 0:
            recommendations.append("Mensagens grandes foram processadas de forma exclusiva para limitar picos de memória.")
        if not recommendations:
            recommendations.append("Pipeline equilibrado; mantenha o perfil atual para a próxima conversão.")
        return recommendations

    def prepare_eml(self, index, eml_file):
        started = time.perf_counter()
        eml_file = self.normalize_path(eml_file)
        if not eml_file.is_file() or eml_file.stat().st_size <= 0:
            raise RuntimeError(f"Arquivo EML ausente ou vazio: {eml_file}")
        size = int(eml_file.stat().st_size)
        with open(eml_file, "rb") as file:
            message = email.message_from_binary_file(file, policy=email.policy.default)
        elapsed = time.perf_counter() - started
        self._prepare_samples.append(elapsed)
        return {
            "index": int(index), "eml_file": eml_file, "message": message,
            "size": size, "large": size >= self.large_eml_mb * 1024 * 1024,
            "prepare_seconds": elapsed
        }

    def iter_prepared_emls(self, eml_files):
        """Pipeline ordenado com backpressure por itens, bytes e pressão do sistema."""
        pending = deque()
        def indexed_entries():
            for fallback_index, value in enumerate(eml_files, start=1):
                if isinstance(value, tuple) and len(value) == 2:
                    yield int(value[0]), Path(value[1])
                else:
                    yield fallback_index, Path(value)
        iterator = iter(indexed_entries())
        deferred = None
        queued_estimated_bytes = 0
        large_pending = 0
        budget = self.memory_budget_mb * 1024 * 1024
        with ThreadPoolExecutor(max_workers=self.max_prepare_workers, thread_name_prefix="pst-mime") as executor:
            exhausted = False
            while pending or deferred is not None or not exhausted:
                self._update_adaptive_limits(len(pending), queued_estimated_bytes)
                target = max(1, min(self._adaptive_queue_limit, self._adaptive_worker_limit * 3))
                while not exhausted and len(pending) < target:
                    if deferred is not None:
                        index, eml_file = deferred
                        deferred = None
                    else:
                        try:
                            index, eml_file = next(iterator)
                        except StopIteration:
                            exhausted = True
                            break
                    try:
                        original_size = max(1, int(Path(eml_file).stat().st_size))
                    except Exception:
                        original_size = 1
                    estimated = max(original_size, int(original_size * 2.5))
                    is_large = original_size >= self.large_eml_mb * 1024 * 1024
                    blocked = pending and (
                        queued_estimated_bytes + estimated > budget
                        or (is_large and large_pending >= 1)
                    )
                    if blocked:
                        deferred = (index, eml_file)
                        break
                    future = executor.submit(self.prepare_eml, index, eml_file)
                    pending.append((index, eml_file, future, estimated, is_large, time.perf_counter()))
                    queued_estimated_bytes += estimated
                    large_pending += 1 if is_large else 0
                self._pipeline_metrics["queue_depth"] = len(pending)
                self._pipeline_metrics["queue_bytes"] = queued_estimated_bytes
                self._pipeline_metrics["peak_queue_bytes"] = max(
                    int(self._pipeline_metrics.get("peak_queue_bytes", 0) or 0), queued_estimated_bytes
                )
                if not pending:
                    if deferred is not None:
                        index, eml_file = deferred
                        deferred = None
                        future = executor.submit(self.prepare_eml, index, eml_file)
                        pending.append((index, eml_file, future, 1, False, time.perf_counter()))
                    else:
                        continue
                index, eml_file, future, estimated, was_large, queued_at = pending.popleft()
                queued_estimated_bytes = max(0, queued_estimated_bytes - estimated)
                large_pending = max(0, large_pending - (1 if was_large else 0))
                try:
                    prepared = future.result()
                except Exception as error:
                    prepared = {"index": index, "eml_file": eml_file, "message": None,
                                "size": 0, "large": was_large, "prepare_seconds": 0.0, "error": str(error)}
                self._pipeline_metrics["prepared"] += 1
                self._pipeline_metrics["consumed"] += 1
                self._pipeline_metrics["prepare_seconds"] += float(prepared.get("prepare_seconds", 0.0) or 0.0)
                self._pipeline_metrics["queue_wait_seconds"] += max(0.0, time.perf_counter() - queued_at)
                if prepared.get("large"):
                    self._pipeline_metrics["large_messages"] += 1
                self._pipeline_metrics["queue_depth"] = len(pending)
                self._pipeline_metrics["queue_bytes"] = queued_estimated_bytes
                yield index, eml_file, prepared
                prepared = None
                if was_large:
                    gc.collect()

    @staticmethod
    def is_pst_capacity_error(error):
        text = str(error or "").casefold()
        return any(token in text for token in (
            "atingiu o tamanho máximo", "atingiu o tamanho maximo",
            "maximum size", "maximum limit", "-2147219956", "0x8004060c"
        ))

    def build_fast_resume_file_list(self, eml_files, mail_folders_root, store):
        """Filter completed EML before MIME parsing using cheap file metadata."""
        started = time.perf_counter()
        completed = store.completed_signatures()
        eligible = []
        skipped = 0
        for original_index, eml_file in enumerate(eml_files, start=1):
            path = Path(eml_file)
            try:
                relative = path.relative_to(mail_folders_root).as_posix()
                stat = path.stat()
                signature = (relative, int(stat.st_size), int(stat.st_mtime_ns))
            except (OSError, ValueError):
                eligible.append((original_index, path))
                continue
            if signature in completed:
                skipped += 1
            else:
                eligible.append((original_index, path))
        self._resume_metrics["pending_query_seconds"] = time.perf_counter() - started
        self._resume_metrics["skipped_before_parse"] = skipped
        self._resume_metrics["eligible_items"] = len(eligible)
        return eligible

    def iter_resume_prepared_emls(self, eml_files):
        """Prepare the first pending EML synchronously, then start the normal pipeline."""
        entries = iter(eml_files)
        try:
            first = next(entries)
        except StopIteration:
            return
        if isinstance(first, tuple) and len(first) == 2:
            first_index, first_path = int(first[0]), Path(first[1])
        else:
            first_index, first_path = 1, Path(first)
        selected = time.perf_counter() - self._resume_started_at
        self._resume_metrics["first_selected_seconds"] = selected
        self._resume_metrics["first_source_position"] = first_index
        prepared = self.prepare_eml(first_index, first_path)
        self._pipeline_metrics["prepared"] += 1
        self._pipeline_metrics["consumed"] += 1
        self._pipeline_metrics["prepare_seconds"] += float(prepared.get("prepare_seconds", 0.0) or 0.0)
        self._resume_metrics["first_prepared_seconds"] = time.perf_counter() - self._resume_started_at
        yield first_index, first_path, prepared
        yield from self.iter_prepared_emls(entries)

    def import_eml_to_folder(self, outlook, namespace, eml_file, target_folder, source_key, prepared=None):
        eml_file = self.normalize_path(eml_file)
        item = None
        item_saved = False
        try:
            self.log_stage("eml", "Lendo arquivo EML", eml=str(eml_file))
            if prepared is not None:
                if prepared.get("error"):
                    raise RuntimeError(
                        f"Falha na preparação MIME: {prepared.get('error')}"
                    )
                msg = prepared["message"]
            else:
                if not eml_file.is_file() or eml_file.stat().st_size <= 0:
                    raise RuntimeError("Arquivo EML ausente ou vazio.")
                with open(eml_file, "rb") as file:
                    msg = email.message_from_binary_file(file, policy=email.policy.default)
            self.log_stage("mapi", "Criando item diretamente na pasta do PST", eml=str(eml_file))
            item = target_folder.Items.Add("IPM.Note")
            item.Subject = self.decode_header_value(msg.get("Subject", "(sem assunto)")) or "(sem assunto)"
            for header, attribute in (("To", "To"), ("Cc", "CC"), ("Bcc", "BCC")):
                recipients = self.parse_recipient_header(msg, header)
                if recipients:
                    setattr(item, attribute, recipients)
            text_body, html_body = self.get_message_body(msg)
            if html_body:
                item.HTMLBody = self.wrap_html_body_for_outlook(msg, html_body)
            else:
                prefix = self.build_visible_metadata_header(msg) if self.visible_metadata else ""
                item.Body = prefix + (text_body or "")
            with tempfile.TemporaryDirectory() as temp_dir:
                attachments = (
                    self.extract_attachments_to_temp(msg, temp_dir)
                    if self.import_attachments else []
                )
                self.log_stage("attachment", f"Adicionando {len(attachments)} anexo(s)", eml=str(eml_file))
                self.attach_files_to_mail_item(item, attachments)
                self.log_stage("mapi", "Aplicando propriedades MAPI", eml=str(eml_file))
                self.apply_mapi_properties_to_received_mail(item, msg)
                self.set_source_key(item, source_key)
                self.log_stage("pst", "Salvando item no PST", eml=str(eml_file))
                item.Save()
                item_saved = True
                saved_entry_id = str(getattr(item, "EntryID", "") or "")
                # Determine actual store where the item was saved (may differ from target_folder)
                try:
                    saved_store_id = str(getattr(item.Parent, "StoreID", "") or "")
                except Exception:
                    saved_store_id = str(getattr(target_folder, "StoreID", "") or "")

                # Verify and, if needed, move the item into the intended PST folder to avoid ending up in Drafts
                self.log_stage("verification", "Confirmando EntryID, StoreID e chave de origem", eml=str(eml_file))
                identity = self._verify_pst_item(
                    namespace,
                    target_folder,
                    saved_entry_id,
                    saved_store_id,
                    source_key
                )

                target_store_id = str(getattr(target_folder, "StoreID", "") or "")

                needs_relocation = (
                    bool(identity.get("wrong_store"))
                    or str(identity.get("store_id") or "") != target_store_id
                    or (self.verification_level == "complete" and not identity.get("verified", False))
                )
                if needs_relocation:
                    # Attempt to relocate the item into the target PST folder
                    try:
                        current_item = None
                        try:
                            current_item = namespace.GetItemFromID(saved_entry_id, saved_store_id)
                        except Exception:
                            current_item = None

                        already_in_target_folder = False
                        if current_item is not None:
                            try:
                                already_in_target_folder = (
                                    str(getattr(current_item.Parent, "EntryID", "") or "")
                                    == str(getattr(target_folder, "EntryID", "") or "")
                                )
                            except Exception:
                                already_in_target_folder = False

                        if already_in_target_folder:
                            # O item já está na pasta correta; a verificação por
                            # source_key só não confirmou a tempo (Outlook lento).
                            # Mover um item para a pasta onde ele já está pode
                            # falhar no COM do Outlook e levar à exclusão indevida
                            # de um item que foi salvo corretamente — não faz
                            # sentido tentar.
                            self.log_stage(
                                "verification",
                                "Item já está na pasta correta; ignorando relocação",
                                eml=str(eml_file)
                            )
                        elif current_item is not None:
                            try:
                                moved = current_item.Move(target_folder)
                                moved.Save()
                                moved_entry_id = str(getattr(moved, "EntryID", "") or "")
                                moved_store_id = str(getattr(moved.Parent, "StoreID", "") or "")

                                # Re-verify moved item
                                identity = self._verify_pst_item(
                                    namespace,
                                    target_folder,
                                    moved_entry_id,
                                    moved_store_id,
                                    source_key
                                )

                                if not identity.get("verified", True):
                                    raise RuntimeError("Falha ao mover item para o PST de destino.")

                                # Update saved ids to moved ones for downstream logging
                                saved_entry_id = moved_entry_id
                                saved_store_id = moved_store_id
                            except Exception as move_error:
                                # The item may have been created in Drafts before Move failed.
                                # Remove only this temporary item to avoid orphaned drafts.
                                try:
                                    if current_item is not None:
                                        current_item.Delete()
                                except Exception:
                                    pass
                                self.log_error(
                                    "[PST-EVENT] " + json.dumps(
                                        {
                                            "category": "move_error",
                                            "message": "Falha ao mover item para pasta PST de destino",
                                            "error": str(move_error),
                                            "eml": str(eml_file)
                                        },
                                        ensure_ascii=False
                                    )
                                )
                                if self.is_pst_capacity_error(move_error):
                                    raise PstCapacityError(str(move_error)) from move_error
                                return {"success": False, "error": str(move_error)}
                    finally:
                        item = None
                else:
                    # A referencia COM e liberada agora; a coleta completa ocorre em lote.
                    item = None
            if identity.get("verified"):
                self.log_stage("success", "EML confirmado no PST", eml=str(eml_file))
            else:
                self.log_stage("verification", "EML salvo e aguardando confirmação em lote", eml=str(eml_file))
            return {"success": True, "error": None, **identity}
        except PstCapacityError:
            if item is not None:
                self._discard_unsaved_or_failed_item(item)
            raise
        except Exception as error:
            if item is not None and not item_saved:
                try:
                    item.Delete()
                except Exception:
                    pass

            self.log_error(
                "[PST-EVENT] " + json.dumps(
                    {
                        "category": "error",
                        "message": "Falha ao importar EML",
                        "eml": str(eml_file),
                        "error": str(error)
                    },
                    ensure_ascii=False
                )
            )
            # O PST atingiu o tamanho máximo também pode acontecer aqui, no
            # item.Save() principal, não só no Move() de relocação — precisa
            # interromper a conversão imediatamente em vez de tentar salvar
            # (e falhar) todos os itens restantes um a um.
            if self.is_pst_capacity_error(error):
                raise PstCapacityError(str(error)) from error
            return {"success": False, "error": str(error)}

    def detach_pst(self, namespace, pst_root):
        try:
            namespace.RemoveStore(pst_root)

            return {
                "success": True,
                "error": None
            }
        except Exception as error:
            return {
                "success": False,
                "error": str(error)
            }

    def detach_store_by_path_if_attached(self, namespace, pst_path):
        """Detach an existing PST from the Outlook profile before it is deleted.

        Needed before --existing-action replace removes the file: deleting a
        .pst that is still attached to the profile fails with a Windows
        sharing-violation PermissionError.
        """
        pst_path = Path(pst_path).resolve()
        for store in namespace.Stores:
            try:
                store_path = getattr(store, "FilePath", None)
                if store_path and Path(store_path).resolve() == pst_path:
                    return self.detach_pst(namespace, store.GetRootFolder())
            except Exception as error:
                return {"success": False, "error": str(error)}
        return {"success": True, "error": None}

    def write_report(self, result):
        pst_path = self.normalize_path(result["pst_path"])

        report_json = pst_path.with_suffix(".pst_report.json")
        report_csv = pst_path.with_suffix(".pst_report.csv")
        failed_csv = pst_path.with_suffix(".pst_failed_imports.csv")

        result["report_path"] = str(report_json)
        result["csv_report_path"] = str(report_csv)
        result["failed_csv_report_path"] = str(failed_csv)

        report_json.parent.mkdir(parents=True, exist_ok=True)

        with open(report_json, "w", encoding="utf-8") as json_file:
            json.dump(
                result,
                json_file,
                ensure_ascii=False,
                indent=4
            )

        with open(
            report_csv,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as csv_file:
            fieldnames = [
                "index",
                "mail_folder",
                "eml_file",
                "success",
                "error"
            ]

            writer = csv.DictWriter(
                csv_file,
                fieldnames=fieldnames
            )

            writer.writeheader()

            for item in result.get("items", []):
                writer.writerow(
                    {
                        "index": item.get("index"),
                        "mail_folder": item.get("mail_folder"),
                        "eml_file": item.get("eml_file"),
                        "success": item.get("success"),
                        "error": item.get("error")
                    }
                )

        failed_items = [
            item for item in result.get("items", [])
            if not item.get("success")
        ]

        with open(
            failed_csv,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as csv_file:
            fieldnames = [
                "index",
                "mail_folder",
                "eml_file",
                "error"
            ]

            writer = csv.DictWriter(
                csv_file,
                fieldnames=fieldnames
            )

            writer.writeheader()

            for item in failed_items:
                writer.writerow(
                    {
                        "index": item.get("index"),
                        "mail_folder": item.get("mail_folder"),
                        "eml_file": item.get("eml_file"),
                        "error": item.get("error")
                    }
                )

    def flush_pst_state(self, checkpoint_path, sqlite_path, store, total_eml, current=0, force=False, failed=False, folder_changed=False):
        now = time.monotonic()
        batch_due = int(current or 0) > 0 and int(current) % self.checkpoint_batch_size == 0
        time_due = now - self._last_progress_at >= self.progress_interval_seconds
        if not (force or failed or folder_changed or batch_due or time_due):
            return False
        self.write_pst_checkpoint_summary(checkpoint_path, store, total_eml)
        if force or failed or folder_changed or time_due or int(current or 0) != self._last_progress_current:
            self.emit_sqlite_pst_progress(sqlite_path, store, total_eml)
            self._last_progress_at = now
            self._last_progress_current = int(current or 0)
        return True

    def export_backup_to_pst(
        self,
        backup_root,
        pst_path,
        pst_display_name="M365 Mailbox Backup",
        detach_after=False,
        existing_action="resume",
        folder_mode="preserve",
        root_folder_name="",
        visible_metadata=True,
        import_attachments=True,
        image_max_width=700,
        verification_level="balanced",
        verification_batch_size=None,
        prepare_workers=None,
        prepare_queue_size=None,
        large_eml_mb=None,
        performance_profile=None,
        adaptive_enabled=None,
        memory_budget_mb=None,
        min_prepare_workers=None,
        max_prepare_workers=None,
        com_slow_seconds=None
    ):
        started_at = datetime.now()
        self._resume_started_at = time.perf_counter()
        self.folder_mode = folder_mode if folder_mode in ("preserve", "single") else "preserve"
        self.root_folder_name = str(root_folder_name or "").strip()
        self.visible_metadata = bool(visible_metadata)
        self.import_attachments = bool(import_attachments)
        self.image_max_width = max(200, min(2000, int(image_max_width or 700)))
        self.verification_level = (
            verification_level if verification_level in ("quick", "balanced", "complete") else "balanced"
        )
        if verification_batch_size is not None:
            self.verification_batch_size = max(1, min(500, int(verification_batch_size)))
        if prepare_workers is not None:
            self.prepare_workers = max(1, min(8, int(prepare_workers)))
        if prepare_queue_size is not None:
            self.prepare_queue_size = max(2, min(100, int(prepare_queue_size)))
        if large_eml_mb is not None:
            self.large_eml_mb = max(1, min(500, int(large_eml_mb)))
        if performance_profile is not None:
            self.performance_profile = str(performance_profile).strip().lower()
        if adaptive_enabled is not None:
            self.adaptive_enabled = bool(adaptive_enabled)
        if memory_budget_mb is not None:
            self.memory_budget_mb = max(128, min(8192, int(memory_budget_mb)))
        if min_prepare_workers is not None:
            self.min_prepare_workers = max(1, min(8, int(min_prepare_workers)))
        if max_prepare_workers is not None:
            self.max_prepare_workers = max(self.min_prepare_workers, min(8, int(max_prepare_workers)))
        if com_slow_seconds is not None:
            self.com_slow_seconds = max(1.0, min(120.0, float(com_slow_seconds)))
        self.apply_performance_profile()

        result = {
            "success": False,
            "partial_success": False,
            "phase": "8C",
            "engine": "Outlook Classic COM",
            "started_at": started_at.isoformat(),
            "finished_at": None,
            "duration_seconds": 0,
            "backup_root": str(backup_root),
            "pst_path": str(pst_path),
            "pst_display_name": pst_display_name,
            "detach_after": detach_after,
            "eml_found": 0,
            "eml_imported": 0,
            "eml_failed": 0,
            "items": [],
            "errors": [],
            "report_path": None,
            "csv_report_path": None,
            "failed_csv_report_path": None,
            "checkpoint_path": None,
            "resumed_imported": 0,
            "eml_skipped_checkpoint": 0,
            "performance": {}
        }

        temp_dir = None
        pythoncom_module = None
        checkpoint_store = None
        checkpoint_path = None
        sqlite_path = None
        total_emls = 0
        destination_lock_path = None
        self._folder_cache.clear()
        self._pending_verifications.clear()
        self._verification_metrics.update({
            "saved": 0, "verified": 0, "pending": 0, "reconciled": 0,
            "attempts": 0, "audit_failures": 0
        })
        self._com_samples.clear()
        self._prepare_samples.clear()
        self._pipeline_metrics.update({
            "prepared": 0, "consumed": 0, "queue_depth": 0,
            "large_messages": 0, "prepare_seconds": 0.0, "com_seconds": 0.0,
            "queue_wait_seconds": 0.0, "queue_bytes": 0, "peak_queue_bytes": 0,
            "peak_rss_bytes": 0, "adaptive_adjustments": 0,
            "effective_workers": self._adaptive_worker_limit,
            "effective_queue_limit": self._adaptive_queue_limit,
            "bottleneck": "calculando", "eta_seconds": 0.0, "memory_pressure": False
        })
        self._last_progress_at = 0.0
        self._last_progress_current = -1

        try:
            if os.name != "nt":
                raise RuntimeError(
                    "A exportação PST via Outlook Classic precisa rodar no Windows. "
                    "Execute com PowerShell/CMD usando o Python do Windows, não no WSL."
                )

            try:
                import pythoncom
                import win32com.client

                pythoncom_module = pythoncom
            except ImportError as error:
                raise ImportError(
                    "pywin32 não está instalado. Rode no PowerShell: py -m pip install pywin32"
                ) from error

            backup_root, temp_dir = self.extract_zip_if_needed(
                backup_root
            )

            eml_files, mail_folders_root = self.find_eml_files(
                backup_root
            )

            result["eml_found"] = len(eml_files)

            if not eml_files:
                result["errors"].append(
                    "Nenhum arquivo .eml encontrado no backup informado."
                )

                return result

            pst_path = self.normalize_path(pst_path)
            existing_action = str(existing_action or "resume").lower()
            if pst_path.exists():
                if existing_action == "cancel":
                    raise FileExistsError(f"O arquivo PST já existe: {pst_path}")
                if existing_action == "number":
                    base = pst_path.with_suffix("")
                    for number in range(2, 10000):
                        candidate = Path(f"{base}_{number:03d}.pst")
                        if not candidate.exists():
                            pst_path = candidate
                            result["pst_path"] = str(candidate)
                            break
            # Impede que duas conversões apontando para o mesmo PST corrompam
            # o mesmo perfil do Outlook simultaneamente (o COM não é seguro
            # para reentrância concorrente sobre o mesmo Outlook.Application).
            destination_lock_path = self.acquire_destination_lock(pst_path)
            self.log_stage("precheck", "Inicializando automação COM do Outlook Classic")
            pythoncom.CoInitialize()

            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as error:
                raise RuntimeError(
                    "Não foi possível iniciar a automação COM do Outlook Classic "
                    f"({error}). Verifique se o Outlook Classic (não o novo Outlook "
                    "para Windows, que não suporta automação COM) está instalado, "
                    "com um perfil de e-mail já configurado, e se a arquitetura do "
                    "Python (32/64 bits) é a mesma do Outlook instalado."
                ) from error

            if existing_action == "replace" and pst_path.exists():
                # É preciso desanexar o PST do perfil antes de apagar o arquivo:
                # excluir um .pst ainda aberto no Outlook falha com um
                # PermissionError de violação de compartilhamento do Windows.
                detach_result = self.detach_store_by_path_if_attached(namespace, pst_path)
                if not detach_result["success"]:
                    self.log_info(
                        f"Não foi possível desanexar o PST existente do perfil: "
                        f"{detach_result['error']}"
                    )
                sidecars = (
                    pst_path, pst_path.with_suffix(".pst_checkpoint.sqlite3"),
                    pst_path.with_suffix(".pst_checkpoint.json"),
                    pst_path.with_suffix(".pst_report.json"),
                    pst_path.with_suffix(".pst_report.csv"),
                    pst_path.with_suffix(".pst_failed_imports.csv")
                )
                delete_error = None
                # Mesmo já desanexado, o Outlook (ou indexação/antivírus)
                # costuma manter o arquivo brevemente ocupado logo após o
                # RemoveStore; poucas novas tentativas com espera curta
                # absorvem essa janela sem precisar de intervenção manual.
                for attempt in range(6):
                    delete_error = None
                    for sidecar in sidecars:
                        try:
                            sidecar.unlink(missing_ok=True)
                        except OSError as error:
                            delete_error = error
                    if delete_error is None:
                        break
                    time.sleep(min(0.5 * (attempt + 1), 3.0))
                if delete_error is not None:
                    raise RuntimeError(
                        f"Não foi possível substituir o PST existente em '{pst_path}' "
                        f"porque o arquivo ainda está em uso, provavelmente aberto no "
                        f"Outlook: {delete_error}. Feche o Outlook Classic e tente de "
                        "novo, ou use --existing-action resume/number."
                    ) from delete_error

            self.ensure_disk_capacity(pst_path)

            self.log_stage("outlook", "Abrindo ou anexando o arquivo PST ao perfil")
            outlook_attach_started = time.perf_counter()
            pst_root = self.create_or_attach_pst(
                namespace=namespace,
                pst_path=pst_path,
                display_name=pst_display_name
            )

            total_emls = len(eml_files)
            checkpoint_path = self.get_checkpoint_path(pst_path)
            sqlite_path = self.normalize_path(pst_path).with_suffix(
                ".pst_checkpoint.sqlite3"
            )
            checkpoint_store = CheckpointStore(
                path=sqlite_path,
                operation_type="pst_import",
                source_root=backup_root,
                destination_path=pst_path
            )
            checkpoint_store.set_operation(status="running", total_items=total_emls)

            legacy = self.load_json_safely(checkpoint_path, {})
            legacy_imported = legacy.get("imported_eml", {}) if isinstance(legacy, dict) else {}
            if not isinstance(legacy_imported, dict):
                legacy_imported = {}

            checkpoint_ready_at = time.perf_counter()
            counts = checkpoint_store.counts()
            verification_counts = checkpoint_store.verification_counts()
            self._verification_metrics.update({
                "saved": int(counts.get("completed", 0)) + int(counts.get("verification_pending", 0)),
                "verified": int(verification_counts.get("verified", 0)),
                "pending": int(verification_counts.get("pending", 0)),
                "reconciled": int(verification_counts.get("reconciled", 0)),
                "audit_failures": int(verification_counts.get("failed", 0)),
            })
            self._resume_metrics["checkpoint_seconds"] = max(0.0, checkpoint_ready_at - self._resume_started_at - self._resume_metrics["outlook_seconds"])
            if self.fast_resume_enabled and existing_action == "resume":
                eml_files = self.build_fast_resume_file_list(
                    eml_files, mail_folders_root, checkpoint_store
                )
                self._fast_resume_ramp_pending = bool(eml_files)
                self._adaptive_queue_limit = min(self._adaptive_queue_limit, self.resume_initial_queue)
                self._adaptive_worker_limit = min(self._adaptive_worker_limit, self.resume_initial_batch)
                self.log_stage(
                    "fast_resume",
                    f"Retomada rápida: {self._resume_metrics['skipped_before_parse']} concluídos ignorados antes do MIME; "
                    f"{self._resume_metrics['eligible_items']} elegíveis.",
                    **self._resume_metrics
                )
            result["checkpoint_path"] = str(sqlite_path)
            result["resumed_imported"] = counts["completed"]
            result["eml_imported"] = int(counts.get("completed", 0))
            result["eml_skipped_checkpoint"] = int(self._resume_metrics.get("skipped_before_parse", 0))
            self.flush_pst_state(
                checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                current=counts["completed"], force=True
            )

            rate_per_second = self.get_eml_import_rate_per_second()
            last_eml_processed_at = None
            last_mail_folder = None
            processed_since_gc = 0

            self.log_stage(
                "pipeline",
                f"Pipeline {self.performance_profile}: até {self.max_prepare_workers} worker(s), "
                f"fila {self.prepare_queue_size}, memória {self.memory_budget_mb} MB"
            )
            prepared_iterator = (
                self.iter_resume_prepared_emls(eml_files)
                if self.fast_resume_enabled and existing_action == "resume" and eml_files
                else self.iter_prepared_emls(eml_files)
            )
            first_pending_commit = bool(
                self.fast_resume_enabled and existing_action == "resume" and eml_files
            )
            for index, eml_file, prepared in prepared_iterator:
                eml_key, eml_signature = self.build_eml_key(
                    mail_folders_root,
                    eml_file
                )
                source_key = self.build_source_key(
                    backup_root,
                    eml_key,
                    eml_signature["size"],
                    eml_signature["mtime_ns"]
                )
                legacy_entry = legacy_imported.get(eml_key)
                if (
                    isinstance(legacy_entry, dict)
                    and int(legacy_entry.get("size", -1)) == eml_signature["size"]
                    and int(legacy_entry.get("mtime_ns", -1)) == eml_signature["mtime_ns"]
                ):
                    checkpoint_store.mark_completed(
                        source_key,
                        relative_path=eml_key,
                        file_size=eml_signature["size"],
                        modified_ns=eml_signature["mtime_ns"],
                        output_path=str(pst_path)
                    )

                if checkpoint_store.is_completed(
                    source_key,
                    eml_signature["size"],
                    eml_signature["mtime_ns"]
                ):
                    result["eml_skipped_checkpoint"] += 1
                    result["eml_imported"] += 1
                    continue

                if existing_action == "resume" and not self.retry_failed_on_resume:
                    previous_status = checkpoint_store.get_item(source_key)
                    if previous_status and previous_status.get("status") == "failed":
                        result["eml_failed"] += 1
                        continue

                folder_parts = self.get_pst_folder_parts(
                    mail_folders_root=mail_folders_root,
                    eml_file=eml_file
                )

                mail_folder = "/".join(folder_parts)
                last_eml_processed_at = self.enforce_eml_rate_limit(
                    last_eml_processed_at,
                    rate_per_second
                )

                item_result = {
                    "index": index,
                    "mail_folder": mail_folder,
                    "eml_file": str(eml_file),
                    "success": False,
                    "error": None
                }

                try:
                    if index == 1 or index % self.checkpoint_batch_size == 0:
                        self.ensure_disk_capacity(pst_path)

                    target_folder = self.get_or_create_outlook_folder(
                        root_folder=pst_root,
                        folder_parts=folder_parts
                    )

                    previous = checkpoint_store.get_item(source_key)
                    existing_identity = None
                    if previous and previous.get("status") in ("processing", "failed", "verification_pending"):
                        existing_identity = self.find_item_by_source_key(target_folder, source_key)
                    if existing_identity:
                        checkpoint_store.mark_completed(
                            source_key, relative_path=eml_key,
                            file_size=eml_signature["size"],
                            modified_ns=eml_signature["mtime_ns"],
                            output_path=str(pst_path), source_key=source_key,
                            destination_entry_id=existing_identity.get("entry_id"),
                            destination_store_id=existing_identity.get("store_id"),
                            destination_folder=existing_identity.get("folder_name"),
                            verification_status="reconciled"
                        )
                        result["eml_skipped_checkpoint"] += 1
                        result["eml_imported"] += 1
                        self.flush_pst_state(
                            checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                            current=result["eml_imported"]
                        )
                        continue

                    checkpoint_store.mark_processing(
                        source_key,
                        relative_path=eml_key,
                        file_size=eml_signature["size"],
                        modified_ns=eml_signature["mtime_ns"],
                        output_path=str(pst_path)
                    )
                    checkpoint_store.set_sequence_number(source_key, index)

                    com_started = time.perf_counter()
                    import_result = self.import_eml_to_folder(
                        outlook=outlook,
                        namespace=namespace,
                        eml_file=eml_file,
                        target_folder=target_folder,
                        source_key=source_key,
                        prepared=prepared
                    )
                    self.record_com_timing(time.perf_counter() - com_started)
                    processed_for_eta = max(1, result["eml_imported"] + result["eml_failed"] + 1)
                    elapsed_for_eta = max(0.001, (datetime.now() - started_at).total_seconds())
                    remaining_for_eta = max(0, total_emls - processed_for_eta)
                    self._pipeline_metrics["eta_seconds"] = remaining_for_eta * (elapsed_for_eta / processed_for_eta)

                    item_result["success"] = import_result["success"]
                    item_result["error"] = import_result["error"]

                    if import_result["success"]:
                        if first_pending_commit and not import_result.get("verified"):
                            immediate_identity = self._verify_pst_item(
                                namespace, target_folder,
                                import_result.get("entry_id"),
                                import_result.get("store_id"),
                                source_key, final=True
                            )
                            import_result.update(immediate_identity)
                            if not import_result.get("verified"):
                                import_result["success"] = False
                                import_result["error"] = (
                                    immediate_identity.get("verification_error")
                                    or "O primeiro EML da retomada foi salvo, mas não pôde ser confirmado no PST."
                                )
                                item_result["success"] = False
                                item_result["error"] = import_result["error"]
                        if self._fast_resume_ramp_pending:
                            self._fast_resume_ramp_pending = False
                            self._adaptive_worker_limit = max(
                                self.min_prepare_workers, min(self.prepare_workers, self.max_prepare_workers)
                            )
                            self._adaptive_queue_limit = self.prepare_queue_size
                        self._verification_metrics["saved"] += 1
                        if import_result.get("verified"):
                            result["eml_imported"] += 1
                            checkpoint_store.mark_completed(
                                source_key, relative_path=eml_key,
                                file_size=eml_signature["size"],
                                modified_ns=eml_signature["mtime_ns"],
                                output_path=str(pst_path),
                                destination_entry_id=import_result.get("entry_id"),
                                destination_store_id=import_result.get("store_id"),
                                destination_folder=mail_folder, source_key=source_key,
                                verification_status=import_result.get("verification_status", "verified")
                            )
                            self._verification_metrics["verified"] += 1
                            if first_pending_commit:
                                committed = time.perf_counter() - self._resume_started_at
                                self._resume_metrics["first_item_seconds"] = committed
                                self._resume_metrics["first_committed_seconds"] = committed
                                self._resume_metrics["total_seconds"] = committed
                                self._resume_metrics["first_source_position"] = int(index)
                                self._resume_metrics["first_relative_path"] = eml_key
                                self._resume_metrics["first_commit_target_met"] = bool(
                                    committed < self.resume_first_commit_target_seconds
                                )
                                first_pending_commit = False
                                self.log_stage(
                                    "resume_first_commit",
                                    f"Primeiro EML de continuação confirmado no PST em {committed:.2f}s "
                                    f"(posição original {index}).",
                                    **self._resume_metrics
                                )
                                self.flush_pst_state(
                                    checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                                    current=result["eml_imported"], force=True
                                )
                        else:
                            checkpoint_store.mark_verification_pending(
                                source_key, relative_path=eml_key,
                                file_size=eml_signature["size"],
                                modified_ns=eml_signature["mtime_ns"],
                                output_path=str(pst_path), source_key=source_key,
                                destination_entry_id=import_result.get("entry_id"),
                                destination_store_id=import_result.get("store_id"),
                                destination_folder=mail_folder
                            )
                            self._queue_verification({
                                "source_key": source_key, "relative_path": eml_key,
                                "file_size": eml_signature["size"],
                                "modified_ns": eml_signature["mtime_ns"],
                                "output_path": str(pst_path),
                                "entry_id": import_result.get("entry_id"),
                                "store_id": import_result.get("store_id"),
                                "destination_folder": mail_folder,
                                "target_folder": target_folder
                            })
                            if self.verification_level == "balanced" and len(self._pending_verifications) >= self.verification_batch_size:
                                self.drain_pending_verifications(
                                    namespace, checkpoint_store, result, final=False
                                )
                        folder_changed = mail_folder != last_mail_folder
                        self.flush_pst_state(
                            checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                            current=result["eml_imported"], folder_changed=folder_changed
                        )
                        last_mail_folder = mail_folder
                        if index % self.checkpoint_batch_size == 0 or index == total_emls:
                            self.log_info(f"[OK] {index}/{total_emls} | Pasta: {mail_folder}")
                    else:
                        result["eml_failed"] += 1
                        checkpoint_store.mark_failed(
                            source_key,
                            import_result["error"],
                            relative_path=eml_key,
                            file_size=eml_signature["size"],
                            modified_ns=eml_signature["mtime_ns"],
                            output_path=str(pst_path),
                            destination_entry_id=import_result.get("entry_id"),
                            destination_store_id=import_result.get("store_id")
                        )
                        self.flush_pst_state(
                            checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                            current=result["eml_imported"], failed=True
                        )

                        error_message = (
                            f"Falha ao importar EML: {eml_file} | {import_result['error']}"
                        )

                        result["errors"].append(error_message)

                        self.log_error(error_message)

                except PstCapacityError as error:
                    self._resume_metrics["failure_reason"] = "capacity_blocked"
                    checkpoint_store.mark_failed(
                        source_key, str(error), relative_path=eml_key,
                        file_size=eml_signature["size"], modified_ns=eml_signature["mtime_ns"],
                        output_path=str(pst_path), verification_status="capacity_blocked"
                    )
                    checkpoint_store.mark_failure_class(source_key, "capacity_blocked", retryable=False)
                    self._resume_metrics["capacity_blocked_items"] += 1
                    checkpoint_store.set_operation(status="capacity_blocked", total_items=total_emls)
                    self.flush_pst_state(
                        checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                        current=result["eml_imported"], force=True, failed=True
                    )
                    raise RuntimeError(
                        "O PST recusou novas gravações por limite de capacidade. "
                        "A conversão foi interrompida imediatamente e o checkpoint foi preservado."
                    ) from error
                except Exception as error:
                    item_result["success"] = False
                    item_result["error"] = str(error)

                    result["eml_failed"] += 1
                    checkpoint_store.mark_failed(
                        source_key,
                        str(error),
                        relative_path=eml_key,
                        file_size=eml_signature["size"],
                        modified_ns=eml_signature["mtime_ns"],
                        output_path=str(pst_path)
                    )
                    self.flush_pst_state(
                        checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                        current=result["eml_imported"], failed=True
                    )

                    error_message = (
                        f"Falha ao processar EML: {eml_file} | {error}"
                    )

                    result["errors"].append(error_message)

                    self.log_error(error_message)

                result["items"].append(item_result)
                processed_since_gc += 1
                if processed_since_gc >= self.gc_interval:
                    gc.collect()
                    processed_since_gc = 0

            self.log_stage("audit", "Executando auditoria final das confirmações MAPI", pending=len(self._pending_verifications))
            self.drain_pending_verifications(
                namespace, checkpoint_store, result, final=True
            )
            if self._pending_verifications:
                raise RuntimeError("A auditoria final terminou com confirmações pendentes.")
            gc.collect()
            if detach_after:
                detach_result = self.detach_pst(
                    namespace=namespace,
                    pst_root=pst_root
                )

                if not detach_result["success"]:
                    result["errors"].append(
                        f"Não foi possível remover o PST do Outlook após exportação: {detach_result['error']}"
                    )

            result["success"] = (
                result["eml_found"] > 0
                and result["eml_imported"] > 0
                and result["eml_failed"] == 0
            )

            result["partial_success"] = (
                result["eml_found"] > 0
                and result["eml_imported"] > 0
                and result["eml_failed"] > 0
            )

            checkpoint_store.set_operation(
                status=(
                    "completed" if result["success"] else "completed_with_errors"
                ),
                total_items=total_emls
            )
            checkpoint_store.compact()
            self.flush_pst_state(
                checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                current=result["eml_imported"], force=True
            )

        except Exception as error:
            result["errors"].append(str(error))

        finally:
            finished_at = datetime.now()

            result["finished_at"] = finished_at.isoformat()
            result["duration_seconds"] = round(
                (finished_at - started_at).total_seconds(),
                2
            )
            result["performance"] = {
                **self._pipeline_metrics,
                "profile": self.performance_profile,
                "adaptive_enabled": self.adaptive_enabled,
                "memory_budget_mb": self.memory_budget_mb,
                "recommendations": self.performance_recommendations(),
                "resume": dict(self._resume_metrics)
            }

            if checkpoint_store is not None and checkpoint_path is not None and sqlite_path is not None:
                try:
                    self.flush_pst_state(
                        checkpoint_path, sqlite_path, checkpoint_store, total_emls,
                        current=result.get("eml_imported", 0), force=True,
                        failed=bool(result.get("errors"))
                    )
                except Exception:
                    pass
            try:
                self.write_report(result)
            except Exception as report_error:
                result["errors"].append(
                    f"Falha ao gerar relatório PST: {report_error}"
                )

            if temp_dir:
                temp_dir.cleanup()

            self.release_destination_lock(destination_lock_path)

            if checkpoint_store is not None:
                try:
                    checkpoint_store.close()
                except Exception:
                    pass
            if pythoncom_module is not None:
                try:
                    pythoncom_module.CoUninitialize()
                except Exception:
                    pass

        return result


def print_pst_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 8C — CONVERSÃO EML PARA PST")
    _report("=" * 70)
    _report(f"Engine: {result.get('engine')}")
    _report(f"Backup origem: {result.get('backup_root')}")
    _report(f"PST destino: {result.get('pst_path')}")
    _report(f"Nome PST: {result.get('pst_display_name')}")
    _report(f"EML encontrados: {result.get('eml_found')}")
    _report(f"EML importados: {result.get('eml_imported')}")
    _report(f"EML com falha: {result.get('eml_failed')}")
    _report(f"Duração: {result.get('duration_seconds')} segundos")

    if result.get("report_path"):
        _report(f"Relatório JSON: {result.get('report_path')}")

    if result.get("csv_report_path"):
        _report(f"Relatório CSV: {result.get('csv_report_path')}")

    if result.get("failed_csv_report_path"):
        _report(f"Falhas CSV: {result.get('failed_csv_report_path')}")

    if result.get("success"):
        _report("")
        _report("[SUCESSO] PST gerado com sucesso.")
    elif result.get("partial_success"):
        _report("")
        _report("[ATENÇÃO] PST gerado parcialmente.")
        _report("Alguns EML falharam, mas a conversão continuou.")
        _report("Revise o CSV de falhas.")
    else:
        _report("")
        _report("[FALHA] A geração do PST não conseguiu importar nenhum EML ou terminou com erro crítico.")

    if result.get("errors"):
        _report("")
        _report("ERROS/AVISOS")
        _report("-" * 70)

        for error in result.get("errors", []):
            _report(f"- {error}")

    _report("=" * 70)


def main():
    parser = argparse.ArgumentParser(
        description="Converte backups EML do M365 Mailbox Backup para PST usando Outlook Classic."
    )

    parser.add_argument(
        "--backup-root",
        required=True,
        help="Pasta raiz do backup ou arquivo .zip contendo os .eml."
    )

    parser.add_argument(
        "--pst-path",
        required=True,
        help="Caminho do arquivo PST de saída."
    )

    parser.add_argument(
        "--pst-display-name",
        default="M365 Mailbox Backup",
        help="Nome de exibição do PST no Outlook."
    )

    parser.add_argument(
        "--detach-after",
        action="store_true",
        help="Remove o PST do Outlook ao finalizar."
    )
    parser.add_argument("--existing-action", choices=("resume", "number", "replace", "cancel"), default="resume")
    parser.add_argument("--folder-mode", choices=("preserve", "single"), default="preserve")
    parser.add_argument("--root-folder-name", default="")
    parser.add_argument("--hide-visible-metadata", action="store_true")
    parser.add_argument("--skip-attachments", action="store_true")
    parser.add_argument("--image-max-width", type=int, default=700)
    parser.add_argument("--verification-level", choices=("quick", "balanced", "complete"), default="balanced")
    parser.add_argument("--verification-batch-size", type=int, default=None)
    parser.add_argument("--import-rate", type=float, default=10.0)
    parser.add_argument("--prepare-workers", type=int, default=None)
    parser.add_argument("--prepare-queue-size", type=int, default=None)
    parser.add_argument("--large-eml-mb", type=int, default=None)
    parser.add_argument("--performance-profile", choices=("conservative", "balanced", "performance", "custom"), default="balanced")
    parser.add_argument("--disable-adaptive", action="store_true")
    parser.add_argument("--memory-budget-mb", type=int, default=None)
    parser.add_argument("--min-prepare-workers", type=int, default=None)
    parser.add_argument("--max-prepare-workers", type=int, default=None)
    parser.add_argument("--com-slow-seconds", type=float, default=None)

    args = parser.parse_args()

    os.environ["M365_EML_IMPORT_RATE_PER_SECOND"] = str(args.import_rate)
    # Sem um logger real (console + arquivo), log_info/log_error só escreviam
    # em logs/pst/<id>.log; era um _report() bruto que fazia o [PST-PROGRESS]
    # chegar ao stdout monitorado pelo coordenador. Usar o mesmo logger da
    # fase de backup elimina o _report() e mantém esse contrato.
    service = PstExportService(logger=setup_logger())

    result = service.export_backup_to_pst(
        backup_root=args.backup_root,
        pst_path=args.pst_path,
        pst_display_name=args.pst_display_name,
        detach_after=args.detach_after, existing_action=args.existing_action,
        folder_mode=args.folder_mode, root_folder_name=args.root_folder_name,
        visible_metadata=not args.hide_visible_metadata,
        import_attachments=not args.skip_attachments,
        image_max_width=args.image_max_width,
        verification_level=args.verification_level,
        verification_batch_size=args.verification_batch_size,
        prepare_workers=args.prepare_workers,
        prepare_queue_size=args.prepare_queue_size,
        large_eml_mb=args.large_eml_mb,
        performance_profile=args.performance_profile,
        adaptive_enabled=not args.disable_adaptive,
        memory_budget_mb=args.memory_budget_mb,
        min_prepare_workers=args.min_prepare_workers,
        max_prepare_workers=args.max_prepare_workers,
        com_slow_seconds=args.com_slow_seconds
    )

    print_pst_result(result)

    if result.get("eml_imported", 0) <= 0:
        sys.exit(1)


if __name__ == "__main__":
    main()