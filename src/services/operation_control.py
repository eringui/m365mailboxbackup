"""Cooperative pause/cancel control shared by EML and PST workers."""
import os
import sqlite3
import time
from pathlib import Path


class OperationInterrupted(RuntimeError):
    def __init__(self, reason):
        self.reason = str(reason)
        super().__init__(self.reason)


class OperationControl:
    def __init__(self, poll_seconds=0.20):
        self.operation_id = os.getenv("M365_OPERATION_ID", "").strip()
        self.operation_token = os.getenv("M365_OPERATION_TOKEN", "").strip()
        raw_path = os.getenv("M365_OPERATION_DB", "").strip()
        self.database_path = Path(raw_path).expanduser().resolve() if raw_path else None
        self.poll_seconds = max(0.05, float(poll_seconds))
        self._last_check = 0.0
        self._last_result = None

    @property
    def enabled(self):
        return bool(self.operation_id and self.operation_token and self.database_path)

    def requested(self, force=False):
        if not self.enabled:
            return None
        now = time.monotonic()
        if not force and self._last_result is None and now - self._last_check < self.poll_seconds:
            return None
        if not force and self._last_result is not None:
            return self._last_result
        self._last_check = now
        try:
            connection = sqlite3.connect(
                str(self.database_path), timeout=1.0, isolation_level=None
            )
            try:
                row = connection.execute(
                    "SELECT pause_requested, cancel_requested FROM operations "
                    "WHERE operation_id=? AND operation_token=?",
                    (self.operation_id, self.operation_token),
                ).fetchone()
            finally:
                connection.close()
        except sqlite3.Error:
            return None
        if not row:
            return None
        if bool(row[1]):
            self._last_result = "cancel"
        elif bool(row[0]):
            self._last_result = "pause"
        else:
            self._last_result = None
        return self._last_result

    def checkpoint(self):
        reason = self.requested()
        if reason:
            raise OperationInterrupted(reason)

    def interruptible_sleep(self, seconds, step=0.20):
        deadline = time.monotonic() + max(0.0, float(seconds))
        while time.monotonic() < deadline:
            self.checkpoint()
            time.sleep(min(step, max(0.0, deadline - time.monotonic())))
