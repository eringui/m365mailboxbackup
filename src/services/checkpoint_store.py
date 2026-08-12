import json
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path


class CheckpointStore:
    """Transactional item checkpoint backed by SQLite/WAL."""

    SCHEMA_VERSION = 4

    def __init__(self, path, operation_type, source_root, destination_path=None):
        self.path = Path(path).expanduser().resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.operation_type = str(operation_type)
        self.source_root = str(Path(source_root).expanduser().resolve())
        self.destination_path = (
            str(Path(destination_path).expanduser().resolve())
            if destination_path else None
        )
        self._local = threading.local()
        self._initialize()

    def _connect(self):
        connection = sqlite3.connect(
            str(self.path),
            timeout=30,
            isolation_level=None
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA synchronous=NORMAL")
        connection.execute("PRAGMA busy_timeout=30000")
        connection.execute("PRAGMA foreign_keys=ON")
        return connection

    def connection(self):
        connection = getattr(self._local, "connection", None)
        if connection is None:
            connection = self._connect()
            self._local.connection = connection
        return connection

    @contextmanager
    def transaction(self):
        connection = self.connection()
        connection.execute("BEGIN IMMEDIATE")
        try:
            yield connection
            connection.execute("COMMIT")
        except Exception:
            connection.execute("ROLLBACK")
            raise

    def _initialize(self):
        connection = self.connection()
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS operation (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                schema_version INTEGER NOT NULL,
                operation_type TEXT NOT NULL,
                source_root TEXT NOT NULL,
                destination_path TEXT,
                status TEXT NOT NULL,
                total_items INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS items (
                item_key TEXT PRIMARY KEY,
                relative_path TEXT,
                file_size INTEGER,
                modified_ns INTEGER,
                folder_id TEXT,
                output_path TEXT,
                source_key TEXT,
                destination_entry_id TEXT,
                destination_store_id TEXT,
                destination_folder TEXT,
                verification_status TEXT,
                status TEXT NOT NULL,
                attempts INTEGER NOT NULL DEFAULT 0,
                bytes_written INTEGER NOT NULL DEFAULT 0,
                error TEXT,
                started_at TEXT,
                completed_at TEXT,
                updated_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_items_status
            ON items(status);

            CREATE TABLE IF NOT EXISTS folder_sync_state (
                folder_id TEXT PRIMARY KEY,
                folder_path TEXT,
                scope_hash TEXT,
                sync_mode TEXT NOT NULL DEFAULT 'initial',
                status TEXT NOT NULL DEFAULT 'pending',
                next_link TEXT,
                delta_link TEXT,
                page_number INTEGER NOT NULL DEFAULT 0,
                discovered_items INTEGER NOT NULL DEFAULT 0,
                reported_total_items INTEGER NOT NULL DEFAULT 0,
                terminal_reason TEXT,
                limited_scope INTEGER NOT NULL DEFAULT 0,
                last_error TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                completed_at TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_folder_sync_status
            ON folder_sync_state(status);
            """
        )
        # Migração incremental e idempotente para checkpoints criados antes
        # da persistência dos identificadores MAPI do PST.
        item_columns = {row["name"] for row in connection.execute("PRAGMA table_info(items)")}
        for column_name, column_type in (
            ("source_key", "TEXT"),
            ("destination_entry_id", "TEXT"),
            ("destination_store_id", "TEXT"),
            ("destination_folder", "TEXT"),
            ("verification_status", "TEXT"),
        ):
            if column_name not in item_columns:
                connection.execute(f"ALTER TABLE items ADD COLUMN {column_name} {column_type}")
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_items_destination_identity "
            "ON items(destination_store_id, destination_entry_id)"
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_items_source_key ON items(source_key)"
        )
        # Fast-resume metadata. These columns are additive and compatible with
        # checkpoints created by previous versions.
        item_columns = {row["name"] for row in connection.execute("PRAGMA table_info(items)")}
        for column_name, column_type in (
            ("sequence_number", "INTEGER"),
            ("failure_class", "TEXT"),
            ("retryable", "INTEGER NOT NULL DEFAULT 1"),
            ("resume_generation", "INTEGER NOT NULL DEFAULT 0"),
            ("last_attempt_at", "TEXT"),
        ):
            if column_name not in item_columns:
                connection.execute(f"ALTER TABLE items ADD COLUMN {column_name} {column_type}")
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_items_resume "
            "ON items(status, retryable, sequence_number)"
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_items_relative_signature "
            "ON items(relative_path, file_size, modified_ns)"
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_items_verification_resume "
            "ON items(status, verification_status, sequence_number)"
        )
        now = datetime.now().isoformat()
        row = connection.execute(
            "SELECT * FROM operation WHERE id = 1"
        ).fetchone()
        if row is None:
            connection.execute(
                """
                INSERT INTO operation (
                    id, schema_version, operation_type, source_root,
                    destination_path, status, created_at, updated_at
                ) VALUES (1, ?, ?, ?, ?, 'running', ?, ?)
                """,
                (
                    self.SCHEMA_VERSION,
                    self.operation_type,
                    self.source_root,
                    self.destination_path,
                    now,
                    now
                )
            )
        else:
            if int(row["schema_version"] or 0) < self.SCHEMA_VERSION:
                connection.execute(
                    "UPDATE operation SET schema_version = ?, updated_at = ? WHERE id = 1",
                    (self.SCHEMA_VERSION, now)
                )
            if (
                row["operation_type"] != self.operation_type
                or row["source_root"] != self.source_root
                or (row["destination_path"] or None) != self.destination_path
            ):
                raise ValueError(
                    "Checkpoint SQLite pertence a outra origem ou destino."
                )

    def set_operation(self, status=None, total_items=None):
        fields = ["updated_at = ?"]
        values = [datetime.now().isoformat()]
        if status is not None:
            fields.append("status = ?")
            values.append(str(status))
        if total_items is not None:
            fields.append("total_items = ?")
            values.append(max(0, int(total_items)))
        values.append(1)
        self.connection().execute(
            f"UPDATE operation SET {', '.join(fields)} WHERE id = ?",
            values
        )

    def get_operation(self):
        row = self.connection().execute(
            "SELECT * FROM operation WHERE id = 1"
        ).fetchone()
        return dict(row) if row else {}

    def get_item(self, item_key):
        row = self.connection().execute(
            "SELECT * FROM items WHERE item_key = ?",
            (str(item_key),)
        ).fetchone()
        return dict(row) if row else None

    def is_completed(self, item_key, file_size=None, modified_ns=None):
        row = self.connection().execute(
            "SELECT status, file_size, modified_ns FROM items WHERE item_key = ?",
            (str(item_key),)
        ).fetchone()
        if not row or row["status"] != "completed":
            return False
        if file_size is not None and int(row["file_size"] or -1) != int(file_size):
            return False
        if modified_ns is not None and int(row["modified_ns"] or -1) != int(modified_ns):
            return False
        return True

    def mark_processing(
        self,
        item_key,
        relative_path=None,
        file_size=None,
        modified_ns=None,
        folder_id=None,
        output_path=None
    ):
        now = datetime.now().isoformat()
        self.connection().execute(
            """
            INSERT INTO items (
                item_key, relative_path, file_size, modified_ns,
                folder_id, output_path, status, attempts,
                started_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, 'processing', 1, ?, ?)
            ON CONFLICT(item_key) DO UPDATE SET
                relative_path=excluded.relative_path,
                file_size=excluded.file_size,
                modified_ns=excluded.modified_ns,
                folder_id=excluded.folder_id,
                output_path=excluded.output_path,
                status='processing',
                attempts=items.attempts + 1,
                error=NULL,
                started_at=excluded.started_at,
                updated_at=excluded.updated_at
            """,
            (
                str(item_key), relative_path, file_size, modified_ns,
                folder_id, output_path, now, now
            )
        )

    def mark_verification_pending(
        self, item_key, relative_path=None, file_size=None, modified_ns=None,
        output_path=None, source_key=None, destination_entry_id=None,
        destination_store_id=None, destination_folder=None
    ):
        now = datetime.now().isoformat()
        self.connection().execute(
            """
            INSERT INTO items (
                item_key, relative_path, file_size, modified_ns, output_path,
                source_key, destination_entry_id, destination_store_id,
                destination_folder, verification_status, status, attempts, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', 'verification_pending', 1, ?)
            ON CONFLICT(item_key) DO UPDATE SET
                relative_path=COALESCE(excluded.relative_path, items.relative_path),
                file_size=COALESCE(excluded.file_size, items.file_size),
                modified_ns=COALESCE(excluded.modified_ns, items.modified_ns),
                output_path=COALESCE(excluded.output_path, items.output_path),
                source_key=COALESCE(excluded.source_key, items.source_key),
                destination_entry_id=COALESCE(excluded.destination_entry_id, items.destination_entry_id),
                destination_store_id=COALESCE(excluded.destination_store_id, items.destination_store_id),
                destination_folder=COALESCE(excluded.destination_folder, items.destination_folder),
                verification_status='pending', status='verification_pending',
                error=NULL, updated_at=excluded.updated_at
            """,
            (str(item_key), relative_path, file_size, modified_ns, output_path,
             str(source_key or item_key), destination_entry_id, destination_store_id,
             destination_folder, now)
        )

    def verification_pending_items(self, limit=None):
        sql = "SELECT * FROM items WHERE status = 'verification_pending' ORDER BY updated_at"
        params = ()
        if limit is not None:
            sql += " LIMIT ?"
            params = (max(1, int(limit)),)
        return [dict(row) for row in self.connection().execute(sql, params)]

    def verification_counts(self):
        result = {"verified": 0, "reconciled": 0, "pending": 0, "failed": 0}
        for row in self.connection().execute(
            "SELECT COALESCE(verification_status, '') AS verification_status, "
            "status, COUNT(*) AS count FROM items GROUP BY verification_status, status"
        ):
            count = int(row["count"] or 0)
            if row["status"] == "verification_pending":
                result["pending"] += count
            elif row["status"] == "failed":
                result["failed"] += count
            elif row["verification_status"] == "reconciled":
                result["reconciled"] += count
            elif row["status"] == "completed":
                result["verified"] += count
        return result

    def mark_completed(
        self,
        item_key,
        relative_path=None,
        file_size=None,
        modified_ns=None,
        folder_id=None,
        output_path=None,
        bytes_written=0,
        source_key=None,
        destination_entry_id=None,
        destination_store_id=None,
        destination_folder=None,
        verification_status="verified"
    ):
        now = datetime.now().isoformat()
        source_key = str(source_key or item_key)
        self.connection().execute(
            """
            INSERT INTO items (
                item_key, relative_path, file_size, modified_ns,
                folder_id, output_path, source_key, destination_entry_id,
                destination_store_id, destination_folder, verification_status,
                status, attempts, bytes_written, completed_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed', 1, ?, ?, ?)
            ON CONFLICT(item_key) DO UPDATE SET
                relative_path=COALESCE(excluded.relative_path, items.relative_path),
                file_size=COALESCE(excluded.file_size, items.file_size),
                modified_ns=COALESCE(excluded.modified_ns, items.modified_ns),
                folder_id=COALESCE(excluded.folder_id, items.folder_id),
                output_path=COALESCE(excluded.output_path, items.output_path),
                source_key=COALESCE(excluded.source_key, items.source_key),
                destination_entry_id=COALESCE(excluded.destination_entry_id, items.destination_entry_id),
                destination_store_id=COALESCE(excluded.destination_store_id, items.destination_store_id),
                destination_folder=COALESCE(excluded.destination_folder, items.destination_folder),
                verification_status=COALESCE(excluded.verification_status, items.verification_status),
                status='completed', bytes_written=excluded.bytes_written, error=NULL,
                completed_at=excluded.completed_at, updated_at=excluded.updated_at
            """,
            (
                str(item_key), relative_path, file_size, modified_ns, folder_id,
                output_path, source_key, destination_entry_id, destination_store_id,
                destination_folder, verification_status,
                max(0, int(bytes_written or 0)), now, now
            )
        )

    def mark_failed(self, item_key, error, **metadata):
        now = datetime.now().isoformat()
        self.connection().execute(
            """
            INSERT INTO items (
                item_key, relative_path, file_size, modified_ns, folder_id, output_path,
                source_key, destination_entry_id, destination_store_id, destination_folder,
                verification_status, status, attempts, error, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'failed', 1, ?, ?)
            ON CONFLICT(item_key) DO UPDATE SET
                source_key=COALESCE(excluded.source_key, items.source_key),
                destination_entry_id=COALESCE(excluded.destination_entry_id, items.destination_entry_id),
                destination_store_id=COALESCE(excluded.destination_store_id, items.destination_store_id),
                destination_folder=COALESCE(excluded.destination_folder, items.destination_folder),
                verification_status=COALESCE(excluded.verification_status, items.verification_status),
                status='failed', error=excluded.error, updated_at=excluded.updated_at
            """,
            (
                str(item_key), metadata.get("relative_path"), metadata.get("file_size"),
                metadata.get("modified_ns"), metadata.get("folder_id"),
                metadata.get("output_path"), str(metadata.get("source_key") or item_key),
                metadata.get("destination_entry_id"), metadata.get("destination_store_id"),
                metadata.get("destination_folder"), metadata.get("verification_status"),
                str(error), now
            )
        )

    def status_keys(self, status):
        return {
            row["item_key"]
            for row in self.connection().execute(
                "SELECT item_key FROM items WHERE status = ?",
                (str(status),)
            )
        }

    def counts(self):
        counts = {"completed": 0, "failed": 0, "processing": 0, "verification_pending": 0, "bytes_written": 0}
        for row in self.connection().execute(
            "SELECT status, COUNT(*) AS count, "
            "COALESCE(SUM(bytes_written), 0) AS bytes_total "
            "FROM items GROUP BY status"
        ):
            counts[row["status"]] = int(row["count"])
            if row["status"] == "completed":
                counts["bytes_written"] = int(row["bytes_total"] or 0)
        return counts

    def migrate_legacy(self, completed_keys=None, failed_keys=None):
        completed_keys = completed_keys or []
        failed_keys = failed_keys or []
        with self.transaction() as connection:
            now = datetime.now().isoformat()
            connection.executemany(
                """
                INSERT OR IGNORE INTO items (
                    item_key, status, attempts, completed_at, updated_at
                ) VALUES (?, 'completed', 1, ?, ?)
                """,
                [(str(key), now, now) for key in completed_keys]
            )
            connection.executemany(
                """
                INSERT OR IGNORE INTO items (
                    item_key, status, attempts, error, updated_at
                ) VALUES (?, 'failed', 1, 'Migrado do checkpoint legado', ?)
                """,
                [(str(key), now) for key in failed_keys]
            )

    def completed_signatures(self):
        """Return completed file signatures for pre-parse filtering."""
        return {
            (str(row["relative_path"] or ""), int(row["file_size"] or -1), int(row["modified_ns"] or -1))
            for row in self.connection().execute(
                "SELECT relative_path, file_size, modified_ns FROM items WHERE status = 'completed'"
            )
        }

    def set_sequence_number(self, item_key, sequence_number):
        self.connection().execute(
            "UPDATE items SET sequence_number=?, updated_at=? WHERE item_key=?",
            (int(sequence_number), datetime.now().isoformat(), str(item_key))
        )

    def first_resume_candidate(self, include_failed=True):
        rows = self.resume_candidates(limit=1, include_failed=include_failed)
        return rows[0] if rows else None

    def resume_candidates(self, limit=None, include_failed=True):
        statuses = ["processing", "verification_pending"]
        if include_failed:
            statuses.append("failed")
        placeholders = ",".join("?" for _ in statuses)
        sql = (
            f"SELECT * FROM items WHERE status IN ({placeholders}) "
            "AND COALESCE(retryable, 1) = 1 "
            "ORDER BY COALESCE(sequence_number, 2147483647), updated_at"
        )
        params = list(statuses)
        if limit is not None:
            sql += " LIMIT ?"
            params.append(max(1, int(limit)))
        return [dict(row) for row in self.connection().execute(sql, tuple(params))]

    def mark_failure_class(self, item_key, failure_class, retryable=True):
        self.connection().execute(
            "UPDATE items SET failure_class=?, retryable=?, updated_at=? WHERE item_key=?",
            (str(failure_class or "unknown"), 1 if retryable else 0, datetime.now().isoformat(), str(item_key))
        )

    def fast_resume_counts(self):
        result = {
            "completed": 0, "failed": 0, "processing": 0,
            "verification_pending": 0, "retryable": 0, "capacity_blocked": 0
        }
        for row in self.connection().execute(
            "SELECT status, COALESCE(failure_class, '') AS failure_class, "
            "COALESCE(retryable, 1) AS retryable, COUNT(*) AS count "
            "FROM items GROUP BY status, failure_class, retryable"
        ):
            count = int(row["count"] or 0)
            status = str(row["status"] or "")
            result[status] = result.get(status, 0) + count
            if int(row["retryable"] or 0):
                result["retryable"] += count
            if row["failure_class"] == "capacity_blocked":
                result["capacity_blocked"] += count
        return result

    def get_folder_sync_state(self, folder_id, scope_hash=None):
        row = self.connection().execute(
            "SELECT * FROM folder_sync_state WHERE folder_id = ?",
            (str(folder_id),)
        ).fetchone()
        if not row:
            return None
        state = dict(row)
        if scope_hash and state.get("scope_hash") not in (None, str(scope_hash)):
            return None
        return state

    def save_folder_sync_progress(
        self,
        folder_id,
        folder_path=None,
        scope_hash=None,
        sync_mode="initial",
        status="running",
        next_link=None,
        delta_link=None,
        page_number=0,
        discovered_items=0,
        reported_total_items=0,
        terminal_reason=None,
        limited_scope=False,
        last_error=None
    ):
        now = datetime.now().isoformat()
        completed_at = now if status == "complete" else None
        self.connection().execute(
            """
            INSERT INTO folder_sync_state (
                folder_id, folder_path, scope_hash, sync_mode, status,
                next_link, delta_link, page_number, discovered_items,
                reported_total_items, terminal_reason, limited_scope,
                last_error, created_at, updated_at, completed_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(folder_id) DO UPDATE SET
                folder_path=COALESCE(excluded.folder_path, folder_sync_state.folder_path),
                scope_hash=COALESCE(excluded.scope_hash, folder_sync_state.scope_hash),
                sync_mode=excluded.sync_mode,
                status=excluded.status,
                next_link=excluded.next_link,
                delta_link=COALESCE(excluded.delta_link, folder_sync_state.delta_link),
                page_number=excluded.page_number,
                discovered_items=excluded.discovered_items,
                reported_total_items=excluded.reported_total_items,
                terminal_reason=excluded.terminal_reason,
                limited_scope=excluded.limited_scope,
                last_error=excluded.last_error,
                updated_at=excluded.updated_at,
                completed_at=COALESCE(excluded.completed_at, folder_sync_state.completed_at)
            """,
            (
                str(folder_id), folder_path, scope_hash, str(sync_mode),
                str(status), next_link, delta_link, max(0, int(page_number or 0)),
                max(0, int(discovered_items or 0)),
                max(0, int(reported_total_items or 0)),
                terminal_reason, 1 if limited_scope else 0, last_error,
                now, now, completed_at
            )
        )

    def mark_folder_sync_error(self, folder_id, error):
        now = datetime.now().isoformat()
        self.connection().execute(
            """
            UPDATE folder_sync_state
            SET status = 'error', last_error = ?, updated_at = ?
            WHERE folder_id = ?
            """,
            (str(error), now, str(folder_id))
        )

    def reset_folder_sync_state(self, folder_id):
        self.connection().execute(
            "DELETE FROM folder_sync_state WHERE folder_id = ?",
            (str(folder_id),)
        )

    def reset_all_folder_sync_states(self):
        """Force a complete re-enumeration without discarding completed EML items."""
        self.connection().execute("DELETE FROM folder_sync_state")

    def list_folder_sync_states(self):
        return [
            dict(row) for row in self.connection().execute(
                "SELECT * FROM folder_sync_state ORDER BY folder_path, folder_id"
            )
        ]

    def verification_snapshot(self, expected_items=0, expected_folder_ids=None):
        """Return a strict, read-only completion snapshot for the current operation."""
        expected_folder_ids = {str(item) for item in (expected_folder_ids or []) if item}
        counts = self.counts()
        states = self.list_folder_sync_states()
        state_by_id = {str(item.get("folder_id")): item for item in states}
        missing_folders = sorted(expected_folder_ids - set(state_by_id))
        relevant = [state_by_id[item] for item in expected_folder_ids if item in state_by_id]
        incomplete_folders = [
            item for item in relevant
            if item.get("status") != "complete"
            or bool(item.get("next_link"))
            or bool(item.get("last_error"))
        ]
        invalid_files = []
        partial_files = []
        for row in self.connection().execute(
            "SELECT item_key, output_path, bytes_written FROM items WHERE status = 'completed'"
        ):
            output_path = row["output_path"]
            if not output_path:
                invalid_files.append({"item_key": row["item_key"], "reason": "caminho ausente"})
                continue
            path = Path(output_path)
            if path.suffix.casefold() == ".part":
                partial_files.append(str(path))
            if not path.is_file():
                invalid_files.append({"item_key": row["item_key"], "reason": "arquivo ausente", "path": str(path)})
                continue
            physical_size = path.stat().st_size
            recorded_size = int(row["bytes_written"] or 0)
            if physical_size <= 0 or (recorded_size > 0 and physical_size != recorded_size):
                invalid_files.append({
                    "item_key": row["item_key"], "reason": "tamanho inconsistente",
                    "path": str(path), "physical_size": physical_size,
                    "recorded_size": recorded_size
                })
        completed = int(counts.get("completed", 0) or 0)
        failed = int(counts.get("failed", 0) or 0)
        processing = int(counts.get("processing", 0) or 0)
        expected = max(0, int(expected_items or 0))
        pending = max(0, expected - completed)
        verified = (
            expected > 0 and completed >= expected and failed == 0 and processing == 0
            and not missing_folders and not incomplete_folders
            and not invalid_files and not partial_files
        )
        return {
            "completion_verified": verified,
            "expected_items": expected,
            "completed_items": completed,
            "failed_items": failed,
            "processing_items": processing,
            "pending_items": pending,
            "folders_expected": len(expected_folder_ids),
            "folders_complete": len(relevant) - len(incomplete_folders),
            "folders_incomplete": len(incomplete_folders) + len(missing_folders),
            "missing_folder_ids": missing_folders,
            "incomplete_folders": incomplete_folders,
            "invalid_file_count": len(invalid_files),
            "invalid_files": invalid_files[:100],
            "partial_file_count": len(partial_files),
            "partial_files": partial_files[:100]
        }

    def compact(self):
        connection = self.connection()
        connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")

    def close(self):
        connection = getattr(self._local, "connection", None)
        if connection is not None:
            connection.close()
            self._local.connection = None
