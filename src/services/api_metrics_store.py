import csv
import math
import sqlite3
import threading
from datetime import datetime, timedelta
from pathlib import Path


class ApiMetricsStore:
    """Low-overhead Microsoft Graph telemetry stored in SQLite/WAL."""

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
            CREATE TABLE IF NOT EXISTS api_calls (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                mailbox TEXT,
                category TEXT NOT NULL,
                method TEXT NOT NULL,
                status_code INTEGER NOT NULL,
                success INTEGER NOT NULL,
                retry_number INTEGER NOT NULL DEFAULT 0,
                retry_after_seconds REAL NOT NULL DEFAULT 0,
                duration_ms REAL NOT NULL DEFAULT 0,
                bytes_received INTEGER NOT NULL DEFAULT 0,
                friendly_error TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_api_calls_time ON api_calls(created_at);
            CREATE INDEX IF NOT EXISTS idx_api_calls_mailbox ON api_calls(mailbox);
            CREATE INDEX IF NOT EXISTS idx_api_calls_category ON api_calls(category);
            CREATE TABLE IF NOT EXISTS rate_limiter_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                mailbox TEXT,
                wait_seconds REAL NOT NULL DEFAULT 0,
                limiter_scope TEXT NOT NULL,
                profile TEXT,
                effective_rate REAL NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_rate_limiter_time
                ON rate_limiter_events(created_at);
            CREATE INDEX IF NOT EXISTS idx_rate_limiter_mailbox
                ON rate_limiter_events(mailbox);
            """
        )

    def record(self, mailbox, category, method, status_code, success,
               retry_number=0, retry_after_seconds=0, duration_ms=0,
               bytes_received=0, friendly_error=None):
        self.connection().execute(
            """
            INSERT INTO api_calls (
                created_at, mailbox, category, method, status_code, success,
                retry_number, retry_after_seconds, duration_ms,
                bytes_received, friendly_error
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(), mailbox, category, method,
                int(status_code or 0), 1 if success else 0,
                int(retry_number or 0), float(retry_after_seconds or 0),
                float(duration_ms or 0), int(bytes_received or 0), friendly_error
            )
        )

    def record_rate_limiter_wait(self, mailbox, wait_seconds, limiter_scope,
                                 profile=None, effective_rate=0):
        wait_seconds = max(0.0, float(wait_seconds or 0))
        if wait_seconds <= 0:
            return
        self.connection().execute(
            """
            INSERT INTO rate_limiter_events (
                created_at, mailbox, wait_seconds, limiter_scope,
                profile, effective_rate
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(), mailbox, wait_seconds,
                str(limiter_scope), profile, float(effective_rate or 0)
            )
        )

    def _since(self, minutes):
        return (datetime.now() - timedelta(minutes=int(minutes))).isoformat()

    def summary(self, minutes=60):
        since = self._since(minutes)
        row = self.connection().execute(
            """
            SELECT COUNT(*) total, SUM(success) successes,
                   SUM(CASE WHEN success=0 THEN 1 ELSE 0 END) failures,
                   SUM(CASE WHEN retry_number>0 THEN 1 ELSE 0 END) retries,
                   SUM(CASE WHEN status_code=429 THEN 1 ELSE 0 END) throttles,
                   COALESCE(SUM(retry_after_seconds),0) wait_seconds,
                   COALESCE(AVG(duration_ms),0) avg_latency_ms,
                   COALESCE(SUM(bytes_received),0) bytes_received
            FROM api_calls WHERE created_at >= ?
            """,
            (since,)
        ).fetchone()
        result = dict(row)
        limiter_row = self.connection().execute(
            """
            SELECT COALESCE(SUM(wait_seconds),0) wait_seconds,
                   COUNT(*) wait_events
            FROM rate_limiter_events WHERE created_at >= ?
            """,
            (since,)
        ).fetchone()
        result["rate_limiter_wait_seconds"] = float(limiter_row["wait_seconds"] or 0)
        result["rate_limiter_wait_events"] = int(limiter_row["wait_events"] or 0)
        total = int(result.get("total") or 0)
        result["requests_per_second"] = total / max(int(minutes) * 60, 1)
        result["throttle_percent"] = (
            int(result.get("throttles") or 0) / total * 100 if total else 0
        )
        latencies = [
            float(item[0]) for item in self.connection().execute(
                "SELECT duration_ms FROM api_calls WHERE created_at >= ? ORDER BY duration_ms",
                (since,)
            )
        ]
        for percentile in (50, 95, 99):
            if latencies:
                index = min(len(latencies) - 1, math.ceil(len(latencies) * percentile / 100) - 1)
                result[f"p{percentile}_latency_ms"] = latencies[index]
            else:
                result[f"p{percentile}_latency_ms"] = 0
        throttle = result["throttle_percent"]
        avg = float(result.get("avg_latency_ms") or 0)
        failures = int(result.get("failures") or 0)
        if throttle >= 5 or avg >= 3000:
            health = "Limitada"
            recommendation = "Reduza temporariamente o paralelismo e respeite o tempo de espera do Graph."
        elif throttle >= 1 or avg >= 1500 or failures >= 10:
            health = "Atenção"
            recommendation = "Mantenha o paralelismo atual e acompanhe latência e limitações."
        else:
            health = "Normal"
            recommendation = "A API está estável; aumente o paralelismo somente de forma gradual."
        result["health"] = health
        result["recommendation"] = recommendation
        return result

    def grouped(self, field, minutes=60, limit=100):
        if field not in ("category", "mailbox", "status_code"):
            raise ValueError("Agrupamento inválido")
        since = self._since(minutes)
        rows = self.connection().execute(
            f"""
            SELECT COALESCE(CAST({field} AS TEXT),'Sem identificação') name,
                   COUNT(*) requests, SUM(success) successes,
                   SUM(CASE WHEN success=0 THEN 1 ELSE 0 END) failures,
                   SUM(CASE WHEN status_code=429 THEN 1 ELSE 0 END) throttles,
                   COALESCE(SUM(retry_after_seconds),0) wait_seconds,
                   COALESCE(AVG(duration_ms),0) avg_latency_ms,
                   COALESCE(SUM(bytes_received),0) bytes_received
            FROM api_calls WHERE created_at >= ?
            GROUP BY {field} ORDER BY requests DESC LIMIT ?
            """,
            (since, int(limit))
        ).fetchall()
        return [dict(row) for row in rows]

    def timeline(self, minutes=60):
        since = self._since(minutes)
        rows = self.connection().execute(
            """
            SELECT substr(created_at,1,16) minute, COUNT(*) requests,
                   AVG(duration_ms) latency_ms,
                   SUM(CASE WHEN status_code=429 THEN 1 ELSE 0 END) throttles
            FROM api_calls WHERE created_at >= ?
            GROUP BY substr(created_at,1,16) ORDER BY minute
            """,
            (since,)
        ).fetchall()
        return [dict(row) for row in rows]

    def export_csv(self, path, minutes=1440):
        path = Path(path)
        rows = self.grouped("mailbox", minutes=minutes, limit=10000)
        with open(path, "w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=[
                "name", "requests", "successes", "failures", "throttles",
                "wait_seconds", "avg_latency_ms", "bytes_received"
            ])
            writer.writeheader()
            writer.writerows(rows)
        return path
