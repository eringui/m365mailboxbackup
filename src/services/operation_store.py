import json
import sqlite3
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path


def utc_now():
    return datetime.now(timezone.utc).isoformat()


class OperationStore:
    """Persistent source of truth for coordinator-managed operations."""

    def __init__(self, path):
        self.path = Path(path).expanduser().resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._local = threading.local()
        self._initialize()

    def connection(self):
        connection = getattr(self._local, "connection", None)
        if connection is None:
            connection = sqlite3.connect(str(self.path), timeout=3, isolation_level=None)
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA journal_mode=WAL")
            connection.execute("PRAGMA synchronous=NORMAL")
            connection.execute("PRAGMA busy_timeout=3000")
            self._local.connection = connection
        return connection

    def _initialize(self):
        self.connection().executescript(
            """
            CREATE TABLE IF NOT EXISTS operations (
                operation_id TEXT PRIMARY KEY,
                operation_token TEXT NOT NULL,
                operation_type TEXT NOT NULL,
                mailbox TEXT,
                status TEXT NOT NULL,
                queue_position INTEGER NOT NULL DEFAULT 0,
                pid INTEGER,
                process_created_at TEXT,
                source_path TEXT,
                destination_path TEXT,
                backup_path TEXT,
                command_json TEXT NOT NULL,
                options_json TEXT NOT NULL DEFAULT '{}',
                current_items INTEGER NOT NULL DEFAULT 0,
                total_items INTEGER NOT NULL DEFAULT 0,
                failed_items INTEGER NOT NULL DEFAULT 0,
                downloaded_bytes INTEGER NOT NULL DEFAULT 0,
                current_folder TEXT,
                current_page INTEGER NOT NULL DEFAULT 0,
                sync_mode TEXT,
                rate_limiter_enabled INTEGER NOT NULL DEFAULT 0,
                rate_limiter_profile TEXT,
                rate_limiter_wait_seconds REAL NOT NULL DEFAULT 0,
                rate_limiter_wait_events INTEGER NOT NULL DEFAULT 0,
                mime_rate_second REAL NOT NULL DEFAULT 0,
                mime_concurrency INTEGER NOT NULL DEFAULT 0,
                pst_saved_items INTEGER NOT NULL DEFAULT 0,
                pst_verified_items INTEGER NOT NULL DEFAULT 0,
                pst_pending_verifications INTEGER NOT NULL DEFAULT 0,
                pst_reconciled_items INTEGER NOT NULL DEFAULT 0,
                pst_verification_attempts INTEGER NOT NULL DEFAULT 0,
                pst_audit_failures INTEGER NOT NULL DEFAULT 0,
                pst_verification_mode TEXT,
                pst_performance_profile TEXT,
                pst_bottleneck TEXT,
                pst_effective_workers INTEGER NOT NULL DEFAULT 0,
                pst_effective_queue_limit INTEGER NOT NULL DEFAULT 0,
                pst_queue_bytes INTEGER NOT NULL DEFAULT 0,
                pst_peak_queue_bytes INTEGER NOT NULL DEFAULT 0,
                pst_peak_rss_bytes INTEGER NOT NULL DEFAULT 0,
                pst_prepare_seconds REAL NOT NULL DEFAULT 0,
                pst_com_seconds REAL NOT NULL DEFAULT 0,
                pst_queue_wait_seconds REAL NOT NULL DEFAULT 0,
                pst_adaptive_adjustments INTEGER NOT NULL DEFAULT 0,
                pst_eta_seconds REAL NOT NULL DEFAULT 0,
                pst_memory_pressure INTEGER NOT NULL DEFAULT 0,
                pst_resume_total_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_checkpoint_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_outlook_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_pending_query_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_first_item_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_first_selected_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_first_prepared_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_first_committed_seconds REAL NOT NULL DEFAULT 0,
                pst_resume_first_source_position INTEGER NOT NULL DEFAULT 0,
                pst_resume_first_commit_target_seconds REAL NOT NULL DEFAULT 10,
                pst_resume_first_commit_target_met INTEGER NOT NULL DEFAULT 0,
                pst_resume_skipped_before_parse INTEGER NOT NULL DEFAULT 0,
                pst_resume_eligible_items INTEGER NOT NULL DEFAULT 0,
                pst_resume_failure_reason TEXT,
                pst_capacity_blocked_items INTEGER NOT NULL DEFAULT 0,
                pause_requested INTEGER NOT NULL DEFAULT 0,
                cancel_requested INTEGER NOT NULL DEFAULT 0,
                heartbeat_at TEXT,
                started_at TEXT,
                finished_at TEXT,
                last_error TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_operations_status ON operations(status);
            CREATE INDEX IF NOT EXISTS idx_operations_queue ON operations(queue_position);

            CREATE TABLE IF NOT EXISTS operation_events (
                event_id INTEGER PRIMARY KEY AUTOINCREMENT,
                operation_id TEXT,
                event_type TEXT NOT NULL,
                severity TEXT NOT NULL DEFAULT 'info',
                friendly_message TEXT,
                technical_details TEXT,
                payload_json TEXT NOT NULL DEFAULT '{}',
                created_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_events_operation ON operation_events(operation_id, event_id);

            CREATE TABLE IF NOT EXISTS application_state (
                key TEXT PRIMARY KEY,
                value_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            """
        )
        existing_columns = {
            row[1] for row in self.connection().execute("PRAGMA table_info(operations)")
        }
        migrations = {
            "rate_limiter_enabled": "INTEGER NOT NULL DEFAULT 0",
            "rate_limiter_profile": "TEXT",
            "rate_limiter_wait_seconds": "REAL NOT NULL DEFAULT 0",
            "rate_limiter_wait_events": "INTEGER NOT NULL DEFAULT 0",
            "mime_rate_second": "REAL NOT NULL DEFAULT 0",
            "mime_concurrency": "INTEGER NOT NULL DEFAULT 0",
            "pst_saved_items": "INTEGER NOT NULL DEFAULT 0",
            "pst_verified_items": "INTEGER NOT NULL DEFAULT 0",
            "pst_pending_verifications": "INTEGER NOT NULL DEFAULT 0",
            "pst_reconciled_items": "INTEGER NOT NULL DEFAULT 0",
            "pst_verification_attempts": "INTEGER NOT NULL DEFAULT 0",
            "pst_audit_failures": "INTEGER NOT NULL DEFAULT 0",
            "pst_verification_mode": "TEXT",
            "pst_performance_profile": "TEXT",
            "pst_bottleneck": "TEXT",
            "pst_effective_workers": "INTEGER NOT NULL DEFAULT 0",
            "pst_effective_queue_limit": "INTEGER NOT NULL DEFAULT 0",
            "pst_queue_bytes": "INTEGER NOT NULL DEFAULT 0",
            "pst_peak_queue_bytes": "INTEGER NOT NULL DEFAULT 0",
            "pst_peak_rss_bytes": "INTEGER NOT NULL DEFAULT 0",
            "pst_prepare_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_com_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_queue_wait_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_adaptive_adjustments": "INTEGER NOT NULL DEFAULT 0",
            "pst_eta_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_memory_pressure": "INTEGER NOT NULL DEFAULT 0",
            "pst_resume_total_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_checkpoint_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_outlook_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_pending_query_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_first_item_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_first_selected_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_first_prepared_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_first_committed_seconds": "REAL NOT NULL DEFAULT 0",
            "pst_resume_first_source_position": "INTEGER NOT NULL DEFAULT 0",
            "pst_resume_first_commit_target_seconds": "REAL NOT NULL DEFAULT 10",
            "pst_resume_first_commit_target_met": "INTEGER NOT NULL DEFAULT 0",
            "pst_resume_skipped_before_parse": "INTEGER NOT NULL DEFAULT 0",
            "pst_resume_eligible_items": "INTEGER NOT NULL DEFAULT 0",
            "pst_resume_failure_reason": "TEXT",
            "pst_capacity_blocked_items": "INTEGER NOT NULL DEFAULT 0",
        }
        for column, declaration in migrations.items():
            if column not in existing_columns:
                self.connection().execute(
                    f"ALTER TABLE operations ADD COLUMN {column} {declaration}"
                )
        if self.get_setting("backup_concurrency") is None:
            self.set_setting("backup_concurrency", 2)
        if self.get_setting("pst_concurrency") is None:
            self.set_setting("pst_concurrency", 0)

    def create_operation(self, operation_type, command, mailbox=None, source_path=None,
                         destination_path=None, backup_path=None, options=None,
                         status="queued", operation_id=None):
        operation_id = operation_id or f"{operation_type}-{uuid.uuid4().hex[:12]}"
        token = uuid.uuid4().hex
        now = utc_now()
        row = self.connection().execute(
            "SELECT COALESCE(MAX(queue_position), 0) + 1 AS next_position FROM operations"
        ).fetchone()
        position = int(row["next_position"])
        self.connection().execute(
            """
            INSERT INTO operations (
                operation_id, operation_token, operation_type, mailbox, status,
                queue_position, source_path, destination_path, backup_path,
                command_json, options_json, created_at, updated_at, heartbeat_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                operation_id, token, operation_type, mailbox, status, position,
                source_path, destination_path, backup_path,
                json.dumps(command, ensure_ascii=False),
                json.dumps(options or {}, ensure_ascii=False), now, now, now
            )
        )
        self.add_event(operation_id, "operation_created", "info", "Operação adicionada à fila.")
        return self.get_operation(operation_id)

    def _decode(self, row):
        if not row:
            return None
        item = dict(row)
        item["command"] = json.loads(item.pop("command_json") or "[]")
        item["options"] = json.loads(item.pop("options_json") or "{}")
        item["pause_requested"] = bool(item["pause_requested"])
        item["cancel_requested"] = bool(item["cancel_requested"])
        return item

    def get_operation(self, operation_id):
        row = self.connection().execute(
            "SELECT * FROM operations WHERE operation_id = ?", (str(operation_id),)
        ).fetchone()
        return self._decode(row)

    def list_operations(self, include_finished=True):
        where = "" if include_finished else "WHERE status NOT IN ('completed','cancelled')"
        rows = self.connection().execute(
            f"SELECT * FROM operations {where} ORDER BY queue_position, created_at"
        ).fetchall()
        return [self._decode(row) for row in rows]

    def update_operation(self, operation_id, **fields):
        allowed = {
            "status", "queue_position", "pid", "process_created_at", "backup_path",
            "source_path", "destination_path", "options_json", "command_json",
            "current_items", "total_items", "failed_items", "downloaded_bytes",
            "current_folder", "current_page", "sync_mode", "pause_requested",
            "cancel_requested", "heartbeat_at", "started_at", "finished_at", "last_error",
            "rate_limiter_enabled", "rate_limiter_profile", "rate_limiter_wait_seconds",
            "rate_limiter_wait_events", "mime_rate_second", "mime_concurrency",
            "pst_saved_items", "pst_verified_items", "pst_pending_verifications",
            "pst_reconciled_items", "pst_verification_attempts", "pst_audit_failures",
            "pst_verification_mode", "pst_performance_profile", "pst_bottleneck",
            "pst_effective_workers", "pst_effective_queue_limit", "pst_queue_bytes",
            "pst_peak_queue_bytes", "pst_peak_rss_bytes", "pst_prepare_seconds",
            "pst_com_seconds", "pst_queue_wait_seconds", "pst_adaptive_adjustments",
            "pst_eta_seconds", "pst_memory_pressure",
            "pst_resume_total_seconds", "pst_resume_checkpoint_seconds",
            "pst_resume_outlook_seconds", "pst_resume_pending_query_seconds",
            "pst_resume_first_item_seconds", "pst_resume_first_selected_seconds",
            "pst_resume_first_prepared_seconds", "pst_resume_first_committed_seconds",
            "pst_resume_first_source_position", "pst_resume_first_commit_target_seconds",
            "pst_resume_first_commit_target_met", "pst_resume_skipped_before_parse",
            "pst_resume_eligible_items", "pst_resume_failure_reason",
            "pst_capacity_blocked_items",
        }
        assignments, values = [], []
        for key, value in fields.items():
            if key not in allowed:
                continue
            if key in ("pause_requested", "cancel_requested"):
                value = 1 if value else 0
            if key == "options_json" and not isinstance(value, str):
                value = json.dumps(value or {}, ensure_ascii=False)
            if key == "command_json" and not isinstance(value, str):
                value = json.dumps(value or [], ensure_ascii=False)
            assignments.append(f"{key} = ?")
            values.append(value)
        assignments.append("updated_at = ?")
        values.append(utc_now())
        values.append(str(operation_id))
        self.connection().execute(
            f"UPDATE operations SET {', '.join(assignments)} WHERE operation_id = ?",
            values
        )
        return self.get_operation(operation_id)

    def heartbeat(self, operation_id, **progress):
        progress["heartbeat_at"] = utc_now()
        return self.update_operation(operation_id, **progress)

    def request_pause(self, operation_id):
        return self.update_operation(operation_id, pause_requested=True, status="pause_requested")

    def request_cancel(self, operation_id):
        return self.update_operation(operation_id, cancel_requested=True, status="cancel_requested")

    def clear_requests(self, operation_id):
        return self.update_operation(operation_id, pause_requested=False, cancel_requested=False)

    def set_queue_order(self, operation_ids):
        with self.connection() as connection:
            for position, operation_id in enumerate(operation_ids, 1):
                connection.execute(
                    "UPDATE operations SET queue_position = ?, updated_at = ? WHERE operation_id = ?",
                    (position, utc_now(), str(operation_id))
                )
        return self.list_operations()

    def delete_operation(self, operation_id):
        operation = self.get_operation(operation_id)
        if not operation or operation["status"] in (
            "running", "starting", "pausing", "pause_requested", "queued"
        ):
            return False
        self.connection().execute("DELETE FROM operations WHERE operation_id = ?", (operation_id,))
        return True

    def add_event(self, operation_id, event_type, severity="info", friendly_message=None,
                  technical_details=None, payload=None):
        cursor = self.connection().execute(
            """
            INSERT INTO operation_events (
                operation_id, event_type, severity, friendly_message,
                technical_details, payload_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                operation_id, event_type, severity, friendly_message,
                technical_details, json.dumps(payload or {}, ensure_ascii=False), utc_now()
            )
        )
        return self.get_event(cursor.lastrowid)

    def get_event(self, event_id):
        row = self.connection().execute(
            "SELECT * FROM operation_events WHERE event_id = ?", (int(event_id),)
        ).fetchone()
        if not row:
            return None
        item = dict(row)
        item["payload"] = json.loads(item.pop("payload_json") or "{}")
        return item

    def list_events(self, after_id=0, limit=1000):
        rows = self.connection().execute(
            """
            SELECT * FROM operation_events WHERE event_id > ?
            ORDER BY event_id LIMIT ?
            """, (max(0, int(after_id)), max(1, min(int(limit), 5000)))
        ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            item["payload"] = json.loads(item.pop("payload_json") or "{}")
            result.append(item)
        return result

    def set_setting(self, key, value):
        now = utc_now()
        self.connection().execute(
            """
            INSERT INTO application_state (key, value_json, updated_at) VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json, updated_at=excluded.updated_at
            """, (str(key), json.dumps(value, ensure_ascii=False), now)
        )

    def get_setting(self, key, default=None):
        row = self.connection().execute(
            "SELECT value_json FROM application_state WHERE key = ?", (str(key),)
        ).fetchone()
        return json.loads(row["value_json"]) if row else default

    def recover_startup_state(self):
        now = utc_now()
        self.connection().execute(
            """
            UPDATE operations
            SET status='interrupted', pid=NULL, last_error='Coordenador reiniciado durante a operação',
                updated_at=?, finished_at=?
            WHERE status IN ('running','starting','pausing','pause_requested','cancel_requested')
            """, (now, now)
        )
