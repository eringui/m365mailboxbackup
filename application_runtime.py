"""Runtime stability, configuration, diagnostics, profiles and integrity helpers."""

import configparser
import faulthandler
import json
import os
import platform
import shutil
import socket
import sqlite3
import sys
import threading
import traceback
from datetime import datetime
from email.parser import BytesHeaderParser
from pathlib import Path

APP_VERSION = "5.1.3"
SCHEMA_VERSION = 2
DEFAULT_SETTINGS = {
    "appearance": {"theme": "automatic", "font_size": 10, "compact_tables": False},
    "paths": {
        "backup_root": "output/backups", "pst_root": "output/pst",
        "logs_root": "logs", "reports_root": "output/reports", "temp_root": "_temp_gui_jobs"
    },
    "backup": {
        "parallel": 2, "download_workers": 24, "max_pending_downloads": 96,
        "mime_max_concurrency": 6, "resume_mime_max_concurrency": 6,
        "mime_min_interval_seconds": 0.15, "adaptive_throttling": True,
        "throttle_safety_seconds": 1.0, "throttle_jitter_max_seconds": 1.0,
        "throttle_recovery_seconds": 90,
        "rate_limiter_enabled": True, "rate_limiter_profile": "performance_2x",
        "global_mime_rate_second": 10, "global_mime_rate_minute": 480,
        "mailbox_mime_rate_second": 6, "mailbox_mime_rate_minute": 240,
        "performance_revision": 2,
        "rate_limiter_wait_timeout": 300,
        "chunk_size_mb": 1, "checkpoint_batch": 100, "page_size": 100,
        "use_delta": True, "save_next_link": True, "validate_eml": True,
        "preserve_removed": True, "export_attachments": False,
        "skip_calendar": True, "skip_contacts": True, "skip_tasks": True
    },
    "graph": {
        "scope": "https://graph.microsoft.com/.default",
        "url": "https://graph.microsoft.com/v1.0", "timeout_seconds": 60,
        "max_retries": 8, "max_throttle_wait_seconds": 300
    },
    "storage": {
        "warning_free_gb": 50, "critical_free_gb": 10,
        "safety_margin_percent": 15, "auto_pause": True, "check_seconds": 15
    },
    "pst": {
        "parallel": 1, "detach_after": False, "existing_action": "resume",
        "last_source_root": "", "last_display_name": "M365 Mailbox Backup",
        "folder_mode": "preserve", "root_folder_name": "",
        "visible_metadata": True, "import_attachments": True,
        "image_max_width": 700, "verification_level": "balanced",
        "verification_batch_size": 25, "verification_final_retries": 5,
        "open_folder_after": False, "log_visual_limit": 5000,
        "eml_import_rate_per_second": 10,
        "prepare_workers": 3, "prepare_queue_size": 12,
        "large_eml_mb": 25, "performance_profile": "balanced",
        "adaptive_enabled": True, "memory_budget_mb": 512,
        "min_prepare_workers": 1, "max_prepare_workers": 4,
        "com_slow_seconds": 8.0,
        "resume_first_commit_target_seconds": 10.0
    },
    "logs": {"level": "NORMAL", "max_mb": 10, "backups": 5, "event_retention_days": 30},
    "ui": {"start_page": 0, "refresh_ms": 1000, "confirm_close": True}
}


def deep_merge(base, override):
    result = json.loads(json.dumps(base))
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


class AppSettings:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.state_dir = self.root / "_gui_state"
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.path = self.state_dir / "settings.json"
        self.data = self.load()

    def load(self):
        try:
            loaded = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            loaded = {}
        data = deep_merge(DEFAULT_SETTINGS, loaded)
        backup = data.setdefault("backup", {})
        revision = int(backup.get("performance_revision", 0) or 0)
        if revision < 2:
            upgrades = {
                "download_workers": (8, 24),
                "max_pending_downloads": (24, 96),
                "mime_max_concurrency": (3, 6),
                "resume_mime_max_concurrency": (3, 6),
                "mime_min_interval_seconds": (0.35, 0.15),
                "throttle_recovery_seconds": (300, 90),
                "global_mime_rate_second": (5, 10),
                "global_mime_rate_minute": (240, 480),
                "mailbox_mime_rate_second": (3, 6),
                "mailbox_mime_rate_minute": (120, 240),
            }
            for key, (old_value, new_value) in upgrades.items():
                if backup.get(key) == old_value:
                    backup[key] = new_value
            if backup.get("rate_limiter_profile") == "automatic":
                backup["rate_limiter_profile"] = "performance_2x"
            backup["performance_revision"] = 2
        return data

    def save(self, data=None):
        if data is not None:
            self.data = deep_merge(DEFAULT_SETTINGS, data)
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, self.path)
        return self.data

    def get(self, section, key, default=None):
        return self.data.get(section, {}).get(key, default)

    def set(self, section, key, value):
        self.data.setdefault(section, {})[key] = value

    def resolved_path(self, key):
        value = Path(str(self.get("paths", key, ""))).expanduser()
        return value.resolve() if value.is_absolute() else (self.root / value).resolve()

    def build_environment(self):
        """Build the exact environment consumed by backend child processes."""
        return {
            "M365_BACKUP_OUTPUT_ROOT": str(self.resolved_path("backup_root")),
            "M365_LOGS_ROOT": str(self.resolved_path("logs_root")),
            "M365_EML_DOWNLOAD_WORKERS": str(self.get("backup", "download_workers", 24)),
            "M365_EML_DOWNLOAD_CHUNK_SIZE_MB": str(self.get("backup", "chunk_size_mb", 1)),
            "M365_EML_MAX_PENDING_DOWNLOADS": str(self.get("backup", "max_pending_downloads", 96)),
            "M365_MIME_MAX_CONCURRENCY": str(self.get("backup", "mime_max_concurrency", 6)),
            "M365_RESUME_MIME_MAX_CONCURRENCY": str(self.get("backup", "resume_mime_max_concurrency", 6)),
            "M365_MIME_MIN_INTERVAL_SECONDS": str(self.get("backup", "mime_min_interval_seconds", 0.15)),
            "M365_ADAPTIVE_THROTTLING": "1" if self.get("backup", "adaptive_throttling", True) else "0",
            "M365_THROTTLE_SAFETY_SECONDS": str(self.get("backup", "throttle_safety_seconds", 1.0)),
            "M365_THROTTLE_JITTER_MAX_SECONDS": str(self.get("backup", "throttle_jitter_max_seconds", 1.0)),
            "M365_THROTTLE_RECOVERY_SECONDS": str(self.get("backup", "throttle_recovery_seconds", 90)),
            "M365_RATE_LIMITER_ENABLED": "1" if self.get("backup", "rate_limiter_enabled", True) else "0",
            "M365_RATE_LIMITER_PROFILE": str(self.get("backup", "rate_limiter_profile", "performance_2x")),
            "M365_GLOBAL_MIME_RATE_SECOND": str(self.get("backup", "global_mime_rate_second", 10)),
            "M365_GLOBAL_MIME_RATE_MINUTE": str(self.get("backup", "global_mime_rate_minute", 480)),
            "M365_MAILBOX_MIME_RATE_SECOND": str(self.get("backup", "mailbox_mime_rate_second", 6)),
            "M365_MAILBOX_MIME_RATE_MINUTE": str(self.get("backup", "mailbox_mime_rate_minute", 240)),
            "M365_RATE_LIMITER_WAIT_TIMEOUT": str(self.get("backup", "rate_limiter_wait_timeout", 300)),
            "M365_RATE_LIMITER_DB": str(self.state_dir / "graph_rate_limits.sqlite3"),
            "GRAPH_SCOPE": str(self.get("graph", "scope", "https://graph.microsoft.com/.default")),
            "GRAPH_URL": str(self.get("graph", "url", "https://graph.microsoft.com/v1.0")),
            "M365_GRAPH_TIMEOUT_SECONDS": str(self.get("graph", "timeout_seconds", 60)),
            "M365_CHECKPOINT_BATCH_SIZE": str(self.get("backup", "checkpoint_batch", 100)),
            "M365_MESSAGE_PAGE_SIZE": str(self.get("backup", "page_size", 100)),
            "M365_DISK_CRITICAL_GB": str(self.get("storage", "critical_free_gb", 10)),
            "M365_DISK_WARNING_GB": str(self.get("storage", "warning_free_gb", 50)),
            "M365_DISK_AUTO_PAUSE": "1" if self.get("storage", "auto_pause", True) else "0",
            "M365_EML_IMPORT_RATE_PER_SECOND": str(self.get("pst", "eml_import_rate_per_second", 10)),
            "M365_PST_PREPARE_WORKERS": str(self.get("pst", "prepare_workers", 3)),
            "M365_PST_PREPARE_QUEUE_SIZE": str(self.get("pst", "prepare_queue_size", 12)),
            "M365_PST_LARGE_EML_MB": str(self.get("pst", "large_eml_mb", 25)),
            "M365_PST_VERIFICATION_BATCH_SIZE": str(self.get("pst", "verification_batch_size", 25)),
            "M365_PST_VERIFICATION_FINAL_RETRIES": str(self.get("pst", "verification_final_retries", 5)),
            "M365_PST_PERFORMANCE_PROFILE": str(self.get("pst", "performance_profile", "balanced")),
            "M365_PST_ADAPTIVE_ENABLED": "1" if self.get("pst", "adaptive_enabled", True) else "0",
            "M365_PST_MEMORY_BUDGET_MB": str(self.get("pst", "memory_budget_mb", 512)),
            "M365_PST_MIN_PREPARE_WORKERS": str(self.get("pst", "min_prepare_workers", 1)),
            "M365_PST_MAX_PREPARE_WORKERS": str(self.get("pst", "max_prepare_workers", 4)),
            "M365_PST_COM_SLOW_SECONDS": str(self.get("pst", "com_slow_seconds", 8.0)),
            "M365_PST_RESUME_FIRST_COMMIT_TARGET_SECONDS": str(
                self.get("pst", "resume_first_commit_target_seconds", 10.0)
            ),
        }

    def validate_runtime_settings(self):
        errors = []
        for key in ("backup_root", "pst_root"):
            if not str(self.get("paths", key, "")).strip():
                errors.append(f"Caminho obrigatório não informado: {key}")
        if int(self.get("storage", "critical_free_gb", 10)) > int(self.get("storage", "warning_free_gb", 50)):
            errors.append("O limite crítico de disco não pode ser maior que o limite de alerta.")
        if int(self.get("pst", "min_prepare_workers", 1)) > int(self.get("pst", "max_prepare_workers", 4)):
            errors.append("Workers mínimos do PST não podem superar os workers máximos.")
        return errors

    def snapshot_for_operation(self, operation_type, overrides=None):
        environment = self.build_environment()
        overrides = dict(overrides or {})
        if operation_type == "pst":
            aliases = {
                "prepare_workers": "M365_PST_PREPARE_WORKERS",
                "prepare_queue_size": "M365_PST_PREPARE_QUEUE_SIZE",
                "large_eml_mb": "M365_PST_LARGE_EML_MB",
                "verification_batch_size": "M365_PST_VERIFICATION_BATCH_SIZE",
                "performance_profile": "M365_PST_PERFORMANCE_PROFILE",
                "memory_budget_mb": "M365_PST_MEMORY_BUDGET_MB",
                "min_prepare_workers": "M365_PST_MIN_PREPARE_WORKERS",
                "max_prepare_workers": "M365_PST_MAX_PREPARE_WORKERS",
                "com_slow_seconds": "M365_PST_COM_SLOW_SECONDS",
                "resume_first_commit_target_seconds": "M365_PST_RESUME_FIRST_COMMIT_TARGET_SECONDS",
            }
            for option, env_name in aliases.items():
                if option in overrides:
                    environment[env_name] = str(overrides[option])
            if "adaptive_enabled" in overrides:
                environment["M365_PST_ADAPTIVE_ENABLED"] = "1" if overrides["adaptive_enabled"] else "0"
            if "import_rate" in overrides:
                environment["M365_EML_IMPORT_RATE_PER_SECOND"] = str(overrides["import_rate"])
        return {
            "schema": 1, "operation_type": str(operation_type),
            "captured_at": datetime.now().isoformat(),
            "settings_path": str(self.path), "environment": environment
        }

    def apply_environment(self):
        mapping = self.build_environment()
        os.environ.update(mapping)
        return mapping

    def export_profile(self, path):
        profile = deep_merge(self.data, {})
        profile.pop("credentials", None)
        Path(path).write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")

    def import_profile(self, path):
        profile = json.loads(Path(path).read_text(encoding="utf-8"))
        profile.pop("credentials", None)
        self.data = deep_merge(self.data, profile)
        return self.save()


class CredentialStore:
    """DPAPI-backed credentials on Windows, with an env-compatible fallback."""
    def __init__(self, root):
        self.path = Path(root) / "_gui_state" / "credentials.bin"
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def save(self, tenant_id, client_id, client_secret):
        payload = json.dumps({
            "tenant_id": tenant_id.strip(), "client_id": client_id.strip(),
            "client_secret": client_secret, "updated_at": datetime.now().isoformat()
        }).encode("utf-8")
        try:
            import win32crypt
            payload = win32crypt.CryptProtectData(payload, "M365 Mailbox Backup", None, None, None, 0)
        except Exception:
            pass
        self.path.write_bytes(payload)

    def load(self):
        if not self.path.exists():
            return {}
        payload = self.path.read_bytes()
        try:
            import win32crypt
            payload = win32crypt.CryptUnprotectData(payload, None, None, None, 0)[1]
        except Exception:
            pass
        try:
            return json.loads(payload.decode("utf-8"))
        except Exception:
            return {}

    def clear(self):
        self.path.unlink(missing_ok=True)

    def apply_environment(self):
        values = self.load()
        if values:
            os.environ["TENANT_ID"] = values.get("tenant_id", "")
            os.environ["CLIENT_ID"] = values.get("client_id", "")
            os.environ["CLIENT_SECRET"] = values.get("client_secret", "")
        return values

    def import_env(self, env_path):
        values = {}
        path = Path(env_path)
        if path.exists():
            for line in path.read_text(encoding="utf-8-sig").splitlines():
                if "=" in line and not line.lstrip().startswith("#"):
                    key, value = line.split("=", 1)
                    values[key.strip()] = value.strip().strip('"').strip("'")
        if values.get("TENANT_ID") and values.get("CLIENT_ID") and values.get("CLIENT_SECRET"):
            self.save(values["TENANT_ID"], values["CLIENT_ID"], values["CLIENT_SECRET"])
        return values


class SingleInstance:
    def __init__(self, port=48765):
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            self.socket.bind(("127.0.0.1", port))
            self.socket.listen(1)
            self.primary = True
        except OSError:
            self.primary = False

    def close(self):
        try:
            self.socket.close()
        except Exception:
            pass


class CrashReporter:
    def __init__(self, root):
        self.root = Path(root)
        self.log_dir = self.root / "logs"
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.crash_log = self.log_dir / "application_crash.log"
        self.marker = self.root / "_gui_state" / "unclean_shutdown.json"
        self.marker.parent.mkdir(parents=True, exist_ok=True)
        self.previous_unclean = self.marker.exists()
        self.marker.write_text(json.dumps({"started_at": datetime.now().isoformat(), "pid": os.getpid()}), encoding="utf-8")
        try:
            stream = open(self.crash_log, "a", encoding="utf-8")
            faulthandler.enable(stream)
            self._fault_stream = stream
        except Exception:
            self._fault_stream = None
        sys.excepthook = self.handle_exception
        if hasattr(threading, "excepthook"):
            threading.excepthook = lambda args: self.handle_exception(args.exc_type, args.exc_value, args.exc_traceback)

    def handle_exception(self, exc_type, exc_value, exc_traceback):
        with open(self.crash_log, "a", encoding="utf-8") as file:
            file.write("\n" + "=" * 80 + "\n")
            file.write(f"{datetime.now().isoformat()} | M365 Mailbox Backup {APP_VERSION}\n")
            file.write(f"Windows: {platform.platform()} | Python: {sys.version}\n")
            traceback.print_exception(exc_type, exc_value, exc_traceback, file=file)
        try:
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
        except Exception:
            pass

    def clean_shutdown(self):
        self.marker.unlink(missing_ok=True)
        if self._fault_stream:
            self._fault_stream.close()


class EnvironmentDiagnostics:
    def __init__(self, root, settings, credentials):
        self.root, self.settings, self.credentials = Path(root), settings, credentials

    def run(self, include_graph=False):
        results = []
        def add(name, ok, details):
            results.append({"name": name, "ok": bool(ok), "details": str(details)})
        creds = self.credentials.load()
        add("Credenciais", all(creds.get(k) for k in ("tenant_id", "client_id", "client_secret")) or (self.root / ".env").exists(), "Credenciais protegidas ou .env")
        for key in ("backup_root", "pst_root", "logs_root"):
            path = self.settings.resolved_path(key)
            try:
                path.mkdir(parents=True, exist_ok=True)
                probe = path / ".write_test"
                probe.write_text("ok", encoding="utf-8"); probe.unlink()
                add(f"Destino {key}", True, path)
            except Exception as error:
                add(f"Destino {key}", False, error)
        usage = shutil.disk_usage(self.settings.resolved_path("backup_root"))
        add("Espaço em disco", usage.free > int(self.settings.get("storage", "critical_free_gb", 10)) * 1024**3, f"{usage.free / 1024**3:.1f} GB livres")
        try:
            connection = sqlite3.connect(self.root / "_gui_state" / "diagnostic.sqlite3")
            connection.execute("CREATE TABLE IF NOT EXISTS probe(id INTEGER)"); connection.close()
            add("SQLite", True, "Leitura e escrita disponíveis")
        except Exception as error:
            add("SQLite", False, error)
        outlook = False
        try:
            import win32com.client
            win32com.client.Dispatch("Outlook.Application")
            outlook = True
        except Exception:
            pass
        add("Outlook Classic", outlook, "Disponível" if outlook else "PST indisponível; backup EML continua disponível")
        add("Sistema", True, f"{platform.platform()} | Python {platform.python_version()} | App {APP_VERSION}")
        return results


class DatabaseMigrator:
    def __init__(self, root):
        self.root = Path(root)
        self.backup_dir = self.root / "_gui_state" / "migrations"
        self.backup_dir.mkdir(parents=True, exist_ok=True)

    def validate_and_backup(self, database_path):
        path = Path(database_path)
        if not path.exists():
            return None
        with sqlite3.connect(path) as connection:
            result = connection.execute("PRAGMA integrity_check").fetchone()[0]
            if result != "ok":
                raise RuntimeError(f"Banco inconsistente: {result}")
        destination = self.backup_dir / f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
        shutil.copy2(path, destination)
        return destination


class IntegrityValidator:
    def validate_backup(self, backup_root, complete=False):
        root = Path(backup_root)
        report = {"root": str(root), "eml": 0, "valid": 0, "invalid": [], "partial": []}
        for path in root.rglob("*.part"):
            report["partial"].append(str(path))
        for path in root.rglob("*.eml"):
            report["eml"] += 1
            try:
                if path.stat().st_size <= 0:
                    raise ValueError("arquivo vazio")
                with open(path, "rb") as file:
                    header = BytesHeaderParser().parse(file)
                    if not any(header.get(key) for key in ("Subject", "From", "Date", "Message-ID")):
                        raise ValueError("cabeçalho MIME não reconhecido")
                    if complete:
                        while file.read(1024 * 1024):
                            pass
                report["valid"] += 1
            except Exception as error:
                report["invalid"].append({"path": str(path), "error": str(error)})
        report["status"] = "integral" if not report["invalid"] and not report["partial"] else "inconsistente"
        return report
