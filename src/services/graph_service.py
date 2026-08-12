import os
import time
import requests
import msal
import threading
import random
import sqlite3
import math

from requests.adapters import HTTPAdapter
from pathlib import Path
from urllib.parse import quote

try:
    from pyrate_limiter import Duration, Limiter, Rate
    PYRATE_LIMITER_AVAILABLE = True
except ImportError:
    PYRATE_LIMITER_AVAILABLE = False

    class Duration:
        SECOND = 1
        MINUTE = 60

    class Rate:
        def __init__(self, limit, interval):
            self.limit = max(1, int(limit))
            self.interval = max(0.001, float(interval))

    class Limiter:
        """Fallback local para que o backup funcione sem pyrate-limiter instalado."""

        def __init__(self, rate):
            self.rate = rate
            self._lock = threading.Lock()
            self._events = {}

        def try_acquire(self, key, blocking=True, timeout=None):
            started = time.monotonic()
            while True:
                now = time.monotonic()
                with self._lock:
                    events = self._events.setdefault(str(key), [])
                    cutoff = now - self.rate.interval
                    events[:] = [stamp for stamp in events if stamp > cutoff]
                    if len(events) < self.rate.limit:
                        events.append(now)
                        return True
                    wait_for = max(0.01, events[0] + self.rate.interval - now)
                if not blocking:
                    return False
                if timeout is not None:
                    remaining = float(timeout) - (time.monotonic() - started)
                    if remaining <= 0:
                        raise TimeoutError(
                            "Tempo limite aguardando o limitador local do Graph."
                        )
                    wait_for = min(wait_for, remaining)
                time.sleep(min(wait_for, 0.25))
from src.config.settings import GRAPH_URL

try:
    from src.services.api_metrics_store import ApiMetricsStore
except ImportError:
    from api_metrics_store import ApiMetricsStore

from src.config.settings import (
    CLIENT_ID,
    CLIENT_SECRET,
    AUTHORITY,
    GRAPH_SCOPE,
    EML_DOWNLOAD_CHUNK_SIZE,
    MIME_MAX_CONCURRENCY,
    MIME_MIN_INTERVAL_SECONDS,
    THROTTLE_SAFETY_SECONDS,
    THROTTLE_JITTER_MAX_SECONDS,
    THROTTLE_RECOVERY_SECONDS,
    ADAPTIVE_THROTTLING,
    RATE_LIMITER_ENABLED,
    RATE_LIMITER_PROFILE,
    GLOBAL_MIME_RATE_SECOND,
    GLOBAL_MIME_RATE_MINUTE,
    MAILBOX_MIME_RATE_SECOND,
    MAILBOX_MIME_RATE_MINUTE,
    RATE_LIMITER_WAIT_TIMEOUT,
    RATE_LIMITER_DB,
)


class SharedSqliteRateGate:
    """Orçamento global simples entre processos, persistido em SQLite/WAL."""

    def __init__(self, path, per_second, per_minute, timeout=300):
        self.path = Path(path).expanduser().resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.per_second = max(1, int(per_second))
        self.per_minute = max(1, int(per_minute))
        self.timeout = max(1.0, float(timeout))
        with self._connection() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS permits (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at REAL NOT NULL
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_permits_created_at ON permits(created_at)"
            )

    def _connection(self):
        connection = sqlite3.connect(str(self.path), timeout=5, isolation_level=None)
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA synchronous=NORMAL")
        connection.execute("PRAGMA busy_timeout=5000")
        return connection

    def acquire(self):
        started = time.monotonic()
        while True:
            now = time.time()
            connection = self._connection()
            try:
                connection.execute("BEGIN IMMEDIATE")
                connection.execute("DELETE FROM permits WHERE created_at < ?", (now - 60.0,))
                second_count = connection.execute(
                    "SELECT COUNT(*) FROM permits WHERE created_at >= ?", (now - 1.0,)
                ).fetchone()[0]
                minute_count = connection.execute(
                    "SELECT COUNT(*) FROM permits WHERE created_at >= ?", (now - 60.0,)
                ).fetchone()[0]
                if second_count < self.per_second and minute_count < self.per_minute:
                    connection.execute("INSERT INTO permits(created_at) VALUES (?)", (now,))
                    connection.execute("COMMIT")
                    return time.monotonic() - started
                connection.execute("ROLLBACK")
            except sqlite3.OperationalError:
                try:
                    connection.execute("ROLLBACK")
                except sqlite3.Error:
                    pass
            finally:
                connection.close()
            if time.monotonic() - started >= self.timeout:
                raise TimeoutError("Tempo limite aguardando o orçamento global do Graph.")
            time.sleep(0.05)


class PyrateMimeLimiter:
    """Limites preventivos por processo e mailbox usando PyrateLimiter."""

    def __init__(self, global_second, global_minute,
                 mailbox_second, mailbox_minute, timeout=300):
        self.timeout = max(1.0, float(timeout))
        self.global_second_rate = max(1, int(global_second))
        self.mailbox_second_rate = max(1, int(mailbox_second))
        self.limiters = (
            (Limiter(Rate(self.global_second_rate, Duration.SECOND)), "global_second"),
            (Limiter(Rate(max(1, int(global_minute)), Duration.MINUTE)), "global_minute"),
            (Limiter(Rate(self.mailbox_second_rate, Duration.SECOND)), "mailbox_second"),
            (Limiter(Rate(max(1, int(mailbox_minute)), Duration.MINUTE)), "mailbox_minute"),
        )

    def acquire(self, mailbox):
        started = time.monotonic()
        mailbox_key = str(mailbox or "unknown").lower()
        for limiter, scope in self.limiters:
            key = "m365-global" if scope.startswith("global") else mailbox_key
            remaining = max(0.1, self.timeout - (time.monotonic() - started))
            try:
                limiter.try_acquire(key, blocking=True, timeout=remaining)
            except TypeError:
                # Compatibilidade com releases que não expõem timeout nomeado.
                limiter.try_acquire(key, blocking=True)
        return time.monotonic() - started


class GraphService:

    def __init__(self, logger):
        self.graph_url = GRAPH_URL
        self.logger = logger
        self.access_token = None
        self.auth_lock = threading.Lock()
        self.throttle_lock = threading.Lock()
        self.throttle_until = 0.0
        self.mime_rate_lock = threading.Lock()
        self.mime_next_request_at = 0.0
        self.mime_condition = threading.Condition(threading.RLock())
        self.mime_active_requests = 0
        self.mime_configured_concurrency = max(1, int(MIME_MAX_CONCURRENCY))
        self.mime_current_concurrency = self.mime_configured_concurrency
        self.mime_min_interval_seconds = max(0.0, float(MIME_MIN_INTERVAL_SECONDS))
        self.throttle_safety_seconds = max(0.0, float(THROTTLE_SAFETY_SECONDS))
        self.throttle_jitter_max_seconds = max(0.0, float(THROTTLE_JITTER_MAX_SECONDS))
        self.throttle_recovery_seconds = max(30, int(THROTTLE_RECOVERY_SECONDS))
        self.adaptive_throttling = bool(ADAPTIVE_THROTTLING)
        self.last_mime_throttle_at = 0.0
        self.mime_throttle_count = 0
        self.rate_limiter_enabled = bool(RATE_LIMITER_ENABLED)
        self.rate_limiter_profile = str(RATE_LIMITER_PROFILE or "automatic")
        self.rate_limiter = None
        self.shared_rate_gate = None
        self.rate_limiter_wait_seconds = 0.0
        self.rate_limiter_wait_events = 0
        if self.rate_limiter_enabled and not PYRATE_LIMITER_AVAILABLE:
            self.logger.warning(
                "pyrate-limiter não está instalado; usando o limitador local compatível. "
                "O controle global em SQLite permanece ativo."
            )
        if self.rate_limiter_enabled:
            self.rate_limiter = PyrateMimeLimiter(
                GLOBAL_MIME_RATE_SECOND,
                GLOBAL_MIME_RATE_MINUTE,
                MAILBOX_MIME_RATE_SECOND,
                MAILBOX_MIME_RATE_MINUTE,
                max(1, math.ceil(float(RATE_LIMITER_WAIT_TIMEOUT))),
            )
            self.shared_rate_gate = SharedSqliteRateGate(
                RATE_LIMITER_DB,
                GLOBAL_MIME_RATE_SECOND,
                GLOBAL_MIME_RATE_MINUTE,
                max(1, math.ceil(float(RATE_LIMITER_WAIT_TIMEOUT))),
            )
        self.thread_local = threading.local()
        self.session = self._create_session()
        metrics_path = os.getenv("M365_API_METRICS_DB")
        self.metrics_store = ApiMetricsStore(metrics_path) if metrics_path else None
        self.app = msal.ConfidentialClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET
        )

    def _create_session(self):
        session = requests.Session()
        adapter = HTTPAdapter(
            pool_connections=100,
            pool_maxsize=100,
            max_retries=0,
            pool_block=True
        )
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    def _get_thread_session(self):
        session = getattr(self.thread_local, "session", None)
        if session is None:
            session = self._create_session()
            self.thread_local.session = session
        return session

    def _wait_for_shared_throttle(self):
        while True:
            with self.throttle_lock:
                remaining = self.throttle_until - time.monotonic()
            if remaining <= 0:
                return
            time.sleep(min(remaining, 1.0))

    def _set_shared_throttle(self, seconds):
        with self.throttle_lock:
            self.throttle_until = max(
                self.throttle_until,
                time.monotonic() + max(1, seconds)
            )

    def _wait_for_pyrate_limit(self, mailbox_email):
        if not self.rate_limiter_enabled or self.rate_limiter is None:
            return 0.0
        started = time.monotonic()
        local_wait = self.rate_limiter.acquire(mailbox_email)
        shared_wait = self.shared_rate_gate.acquire() if self.shared_rate_gate else 0.0
        waited = max(0.0, time.monotonic() - started)
        self.rate_limiter_wait_seconds += waited
        if waited >= 0.01:
            self.rate_limiter_wait_events += 1
            if self.metrics_store:
                try:
                    self.metrics_store.record_rate_limiter_wait(
                        mailbox_email,
                        waited,
                        "pyrate+sqlite",
                        self.rate_limiter_profile,
                        MAILBOX_MIME_RATE_SECOND,
                    )
                except Exception as error:
                    self.logger.debug(f"Espera do limitador não registrada: {error}")
            if waited >= 1.0:
                self.logger.debug(
                    f"PyrateLimiter aguardou {waited:.2f}s para liberar MIME de {mailbox_email}."
                )
        return max(waited, local_wait, shared_wait)

    def rate_limiter_snapshot(self):
        return {
            "enabled": self.rate_limiter_enabled,
            "profile": self.rate_limiter_profile,
            "wait_seconds": round(self.rate_limiter_wait_seconds, 3),
            "wait_events": self.rate_limiter_wait_events,
            "mailbox_rate_second": int(MAILBOX_MIME_RATE_SECOND),
            "global_rate_second": int(GLOBAL_MIME_RATE_SECOND),
            "mime_concurrency": int(self.mime_current_concurrency),
            "mime_concurrency_max": int(self.mime_configured_concurrency),
        }

    def _wait_for_mime_rate_slot(self):
        if self.mime_min_interval_seconds <= 0:
            return
        with self.mime_rate_lock:
            now = time.monotonic()
            remaining = self.mime_next_request_at - now
            if remaining > 0:
                time.sleep(remaining)
            self.mime_next_request_at = (
                time.monotonic() + self.mime_min_interval_seconds
            )

    def _recover_mime_concurrency_if_stable(self):
        if not self.adaptive_throttling or not self.last_mime_throttle_at:
            return
        now = time.monotonic()
        stable_seconds = now - self.last_mime_throttle_at
        if stable_seconds < self.throttle_recovery_seconds:
            return
        with self.mime_condition:
            if self.mime_current_concurrency < self.mime_configured_concurrency:
                self.mime_current_concurrency += 1
                self.last_mime_throttle_at = now
                self.logger.info(
                    "Controle MIME adaptativo: período estável; "
                    f"concorrência elevada para {self.mime_current_concurrency}/"
                    f"{self.mime_configured_concurrency}."
                )
                self.mime_condition.notify_all()

    def _acquire_mime_slot(self):
        self._recover_mime_concurrency_if_stable()
        with self.mime_condition:
            while self.mime_active_requests >= self.mime_current_concurrency:
                self.mime_condition.wait(timeout=0.5)
                self._recover_mime_concurrency_if_stable()
            self.mime_active_requests += 1

    def _release_mime_slot(self):
        with self.mime_condition:
            self.mime_active_requests = max(0, self.mime_active_requests - 1)
            self.mime_condition.notify_all()

    def _register_mime_throttle(self, retry_after):
        base_wait = max(1.0, float(retry_after or 10))
        jitter = random.uniform(0.0, self.throttle_jitter_max_seconds)
        wait_seconds = base_wait + self.throttle_safety_seconds + jitter
        self._set_shared_throttle(wait_seconds)
        self.last_mime_throttle_at = time.monotonic()
        self.mime_throttle_count += 1
        if self.adaptive_throttling:
            with self.mime_condition:
                previous = self.mime_current_concurrency
                self.mime_current_concurrency = max(
                    1, self.mime_current_concurrency - 1
                )
                if self.mime_current_concurrency != previous:
                    self.mime_condition.notify_all()
                    self.logger.warning(
                        "Controle MIME adaptativo: concorrência reduzida para "
                        f"{self.mime_current_concurrency}/"
                        f"{self.mime_configured_concurrency}."
                    )
        return wait_seconds

    def _metric_context(self, url):
        mailbox = None
        category = "Outras chamadas"
        try:
            path = url.split("graph.microsoft.com", 1)[-1]
            if "/users/" in path:
                mailbox = path.split("/users/", 1)[1].split("/", 1)[0].split("?", 1)[0]
            if "/messages/" in path and path.endswith("/$value"):
                category = "Download de EML"
            elif "/messages" in path:
                category = "Listagem de mensagens"
            elif "/childFolders" in path:
                category = "Listagem de subpastas"
            elif "/mailFolders" in path:
                category = "Listagem de pastas"
            elif "/calendar" in path or "/events" in path:
                category = "Calendário"
            elif "/contacts" in path:
                category = "Contatos"
            elif "/todo" in path or "/tasks" in path:
                category = "Tarefas"
            elif "/users/" in path:
                category = "Consulta de usuário"
        except Exception:
            pass
        return mailbox, category

    def _friendly_http_error(self, status_code):
        return {
            400: "A consulta enviada ao Microsoft 365 não foi aceita.",
            401: "A autenticação precisou ser renovada.",
            403: "O aplicativo não possui permissão para este conteúdo.",
            404: "O item não foi encontrado; ele pode ter sido movido ou excluído.",
            429: "O Microsoft 365 pediu uma redução temporária da velocidade.",
            500: "O serviço Microsoft apresentou uma falha temporária.",
            502: "A comunicação com o serviço Microsoft falhou temporariamente.",
            503: "O serviço Microsoft está temporariamente indisponível.",
            504: "A resposta do serviço Microsoft demorou além do esperado."
        }.get(int(status_code or 0))

    def _record_metric(self, url, response, started_at, retry_count, raw=False):
        if not self.metrics_store:
            return
        try:
            mailbox, category = self._metric_context(url)
            retry_after_value = float(response.headers.get("Retry-After", 0) or 0)
            retry_after = max(0, math.ceil(retry_after_value))
            content_length = int(response.headers.get("Content-Length", 0) or 0)
            if raw and not content_length:
                content_length = len(response.content or b"")
            self.metrics_store.record(
                mailbox=mailbox,
                category=category,
                method="GET",
                status_code=response.status_code,
                success=200 <= response.status_code < 300,
                retry_number=retry_count,
                retry_after_seconds=retry_after,
                duration_ms=(time.perf_counter() - started_at) * 1000,
                bytes_received=content_length,
                friendly_error=self._friendly_http_error(response.status_code)
            )
        except Exception as error:
            self.logger.debug(f"Métrica da API não registrada: {error}")

    def authenticate(self):
        with self.auth_lock:
            self.logger.info("Obtendo token do Microsoft Graph...")

            result = self.app.acquire_token_for_client(
                scopes=[GRAPH_SCOPE]
            )

            if not isinstance(result, dict):
                self.logger.error("Falha ao obter token do Microsoft Graph.")
                self.logger.error(result)
                raise RuntimeError(
                    f"Resposta de autenticação inválida do Microsoft Graph: {result}"
                )

            access_token = result.get("access_token")
            if not isinstance(access_token, str) or not access_token:
                self.logger.error("Falha ao obter token do Microsoft Graph.")
                self.logger.error(result)
                raise RuntimeError(
                    f"Erro ao autenticar no Microsoft Graph: {result}"
                )

            self.access_token = access_token

            self.logger.info("Token obtido com sucesso.")

            return self.access_token

    def get_headers(self):
        if not self.access_token:
            self.authenticate()

        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

    def _normalize_endpoint(self, endpoint):
        if endpoint.startswith("https://"):
            return endpoint

        return f"{GRAPH_URL}{endpoint}"

    def _encode_id(self, value):
        return quote(
            str(value),
            safe=""
        )

    def request_get(self, endpoint, description=None, retry_count=0):
        url = self._normalize_endpoint(endpoint)

        if description:
            self.logger.info(description)

        self.logger.debug(f"GET {url}")

        request_started_at = time.perf_counter()
        response = self.session.get(
            url,
            headers=self.get_headers(),
            timeout=60
        )
        self._record_metric(url, response, request_started_at, retry_count)

        if response.status_code == 401 and retry_count == 0:
            self.logger.warning(
                "Token expirado ou inválido. Renovando token e tentando novamente..."
            )

            self.access_token = None
            self.authenticate()

            return self.request_get(
                endpoint=endpoint,
                description=description,
                retry_count=retry_count + 1
            )

        if response.status_code == 429 and retry_count < 6:
            retry_after = int(
                response.headers.get("Retry-After", "10")
            )

            self.logger.warning(
                f"Throttling detectado. Aguardando {retry_after} segundos..."
            )

            time.sleep(retry_after)

            return self.request_get(
                endpoint=endpoint,
                description=description,
                retry_count=retry_count + 1
            )

        if response.status_code in [500, 502, 503, 504] and retry_count < 3:
            wait_seconds = 5 * (retry_count + 1)

            self.logger.warning(
                f"Erro temporário HTTP {response.status_code}. "
                f"Tentando novamente em {wait_seconds} segundos..."
            )

            time.sleep(wait_seconds)

            return self.request_get(
                endpoint=endpoint,
                description=description,
                retry_count=retry_count + 1
            )

        try:
            data = response.json()
        except Exception:
            self.logger.error("Resposta inválida retornada pelo Graph.")
            self.logger.error(response.text)

            return {
                "success": False,
                "status_code": response.status_code,
                "data": {
                    "raw_response": response.text
                }
            }

        if response.status_code < 200 or response.status_code >= 300:
            self.logger.error(
                f"Erro retornado pelo Graph. HTTP {response.status_code}"
            )

            self.logger.error(data)

            return {
                "success": False,
                "status_code": response.status_code,
                "data": data
            }

        return {
            "success": True,
            "status_code": response.status_code,
            "data": data
        }

    def request_get_raw(self, endpoint, description=None, retry_count=0):
        url = self._normalize_endpoint(endpoint)

        if description:
            self.logger.info(description)

        self.logger.debug(f"GET RAW {url}")

        request_started_at = time.perf_counter()
        response = self.session.get(
            url,
            headers=self.get_headers(),
            timeout=120
        )
        self._record_metric(
            url, response, request_started_at, retry_count, raw=True
        )

        if response.status_code == 401 and retry_count == 0:
            self.logger.warning(
                "Token expirado ou inválido. Renovando token e tentando novamente..."
            )

            self.access_token = None
            self.authenticate()

            return self.request_get_raw(
                endpoint=endpoint,
                description=description,
                retry_count=retry_count + 1
            )

        if response.status_code == 429 and retry_count < 6:
            retry_after = int(
                response.headers.get("Retry-After", "10")
            )

            self.logger.warning(
                f"Throttling detectado. Aguardando {retry_after} segundos..."
            )

            time.sleep(retry_after)

            return self.request_get_raw(
                endpoint=endpoint,
                description=description,
                retry_count=retry_count + 1
            )

        if response.status_code in [500, 502, 503, 504] and retry_count < 3:
            wait_seconds = 5 * (retry_count + 1)

            self.logger.warning(
                f"Erro temporário HTTP {response.status_code}. "
                f"Tentando novamente em {wait_seconds} segundos..."
            )

            time.sleep(wait_seconds)

            return self.request_get_raw(
                endpoint=endpoint,
                description=description,
                retry_count=retry_count + 1
            )

        if response.status_code < 200 or response.status_code >= 300:
            self.logger.error(
                f"Erro ao baixar conteúdo bruto. HTTP {response.status_code}"
            )

            self.logger.error(response.text)

            return {
                "success": False,
                "status_code": response.status_code,
                "content": None,
                "error": response.text
            }

        return {
            "success": True,
            "status_code": response.status_code,
            "content": response.content,
            "error": None
        }

    def get_all_pages(self, endpoint, description=None, max_items=None):
        all_items = []
        next_endpoint = endpoint
        page = 1

        while next_endpoint:
            page_description = description

            if description:
                page_description = f"{description} Página {page}"

            result = self.request_get(
                endpoint=next_endpoint,
                description=page_description
            )

            if not result["success"]:
                return {
                    "success": False,
                    "status_code": result["status_code"],
                    "data": result["data"],
                    "items": all_items
                }

            data = result["data"]

            items = data.get("value", [])

            all_items.extend(items)

            if max_items and len(all_items) >= max_items:
                all_items = all_items[:max_items]
                break

            next_endpoint = data.get("@odata.nextLink")
            page += 1

        return {
            "success": True,
            "status_code": 200,
            "data": {
                "value": all_items
            },
            "items": all_items
        }

    def get_user(self, mailbox_email):
        endpoint = (
            f"/users/{mailbox_email}"
            f"?$select=id,displayName,mail,userPrincipalName,accountEnabled"
        )

        return self.request_get(
            endpoint=endpoint,
            description="Consultando usuário/mailbox..."
        )

    def get_mailbox_summary(self, mailbox_email):
        user_result = self.get_user(mailbox_email)

        if not user_result["success"]:
            return user_result

        user = user_result["data"]

        return {
            "success": True,
            "status_code": 200,
            "data": {
                "id": user.get("id"),
                "displayName": user.get("displayName"),
                "mail": user.get("mail"),
                "userPrincipalName": user.get("userPrincipalName"),
                "accountEnabled": user.get("accountEnabled")
            }
        }

    def get_mail_folders(self, mailbox_email, all_pages=True):
        endpoint = (
            f"/users/{mailbox_email}"
            f"/mailFolders"
            f"?$top=100"
        )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description="Listando todas as pastas da mailbox..."
            )

        return self.request_get(
            endpoint=endpoint,
            description="Listando pastas da mailbox..."
        )

    def get_child_folders(self, mailbox_email, folder_id, all_pages=True):
        folder_id_encoded = self._encode_id(folder_id)

        endpoint = (
            f"/users/{mailbox_email}"
            f"/mailFolders/{folder_id_encoded}/childFolders"
            f"?$top=100"
        )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description=f"Listando subpastas da pasta {folder_id}..."
            )

        return self.request_get(
            endpoint=endpoint,
            description=f"Listando subpastas da pasta {folder_id}..."
        )

    def find_folder_by_name(self, mailbox_email, folder_name):
        folders_result = self.get_mail_folders(
            mailbox_email=mailbox_email,
            all_pages=True
        )

        if not folders_result["success"]:
            return folders_result

        folder_name_normalized = folder_name.lower()

        for folder in folders_result["items"]:
            current_name = folder.get("displayName", "").lower()

            if current_name == folder_name_normalized:
                return {
                    "success": True,
                    "status_code": 200,
                    "data": folder
                }

        return {
            "success": False,
            "status_code": 404,
            "data": {
                "error": f"Pasta não encontrada: {folder_name}"
            }
        }

    def get_inbox(self, mailbox_email):
        endpoint = (
            f"/users/{mailbox_email}"
            f"/mailFolders/inbox"
        )

        return self.request_get(
            endpoint=endpoint,
            description="Localizando Inbox..."
        )

    def get_messages(
        self,
        mailbox_email,
        folder_id=None,
        top=10,
        all_pages=False,
        max_items=None
    ):
        select_fields = (
            "id,"
            "subject,"
            "from,"
            "sender,"
            "toRecipients,"
            "ccRecipients,"
            "receivedDateTime,"
            "sentDateTime,"
            "hasAttachments,"
            "importance,"
            "isRead,"
            "parentFolderId"
        )

        if folder_id:
            folder_id_encoded = self._encode_id(folder_id)

            endpoint = (
                f"/users/{mailbox_email}"
                f"/mailFolders/{folder_id_encoded}"
                f"/messages"
                f"?$top={top}"
                f"&$select={select_fields}"
                f"&$orderby=receivedDateTime desc"
            )
        else:
            endpoint = (
                f"/users/{mailbox_email}"
                f"/messages"
                f"?$top={top}"
                f"&$select={select_fields}"
                f"&$orderby=receivedDateTime desc"
            )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description="Listando mensagens com paginação...",
                max_items=max_items
            )

        return self.request_get(
            endpoint=endpoint,
            description="Listando mensagens..."
        )
    
    def get_message_by_id(self, mailbox_email, message_id):
        endpoint = (
            f"/users/{mailbox_email}"
            f"/messages/{message_id}"
        )

        return self.request_get(
            endpoint=endpoint,
            description=f"Consultando mensagem {message_id}..."
        )

    def get_message_attachments(self, mailbox_email, message_id):
        endpoint = (
            f"/users/{mailbox_email}"
            f"/messages/{message_id}"
            f"/attachments"
        )

        return self.get_all_pages(
            endpoint=endpoint,
            description=f"Listando anexos da mensagem {message_id}..."
        )

    def get_calendar_events(self, mailbox_email, top=10, all_pages=False):
        select_fields = (
            "id,"
            "subject,"
            "start,"
            "end,"
            "organizer,"
            "location,"
            "isCancelled,"
            "createdDateTime,"
            "lastModifiedDateTime"
        )

        endpoint = (
            f"/users/{mailbox_email}"
            f"/events"
            f"?$top={top}"
            f"&$select={select_fields}"
            f"&$orderby=start/dateTime desc"
        )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description="Listando eventos do calendário com paginação..."
            )

        return self.request_get(
            endpoint=endpoint,
            description="Listando eventos do calendário..."
        )

    def get_contacts(self, mailbox_email, top=10, all_pages=False):
        select_fields = (
            "id,"
            "displayName,"
            "givenName,"
            "surname,"
            "emailAddresses,"
            "businessPhones,"
            "mobilePhone,"
            "companyName,"
            "jobTitle"
        )

        endpoint = (
            f"/users/{mailbox_email}"
            f"/contacts"
            f"?$top={top}"
            f"&$select={select_fields}"
        )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description="Listando contatos com paginação..."
            )

        return self.request_get(
            endpoint=endpoint,
            description="Listando contatos..."
        )

    def get_todo_lists(self, mailbox_email, all_pages=True):
        endpoint = (
            f"/users/{mailbox_email}"
            f"/todo/lists"
        )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description="Listando listas de tarefas..."
            )

        return self.request_get(
            endpoint=endpoint,
            description="Listando listas de tarefas..."
        )

    def get_todo_tasks(self, mailbox_email, list_id, all_pages=True):
        endpoint = (
            f"/users/{mailbox_email}"
            f"/todo/lists/{list_id}"
            f"/tasks"
        )

        if all_pages:
            return self.get_all_pages(
                endpoint=endpoint,
                description=f"Listando tarefas da lista {list_id}..."
            )

        return self.request_get(
            endpoint=endpoint,
            description=f"Listando tarefas da lista {list_id}..."
        )

    def list_email_preview(self, mailbox_email, top=10):
        messages_result = self.get_messages(
            mailbox_email=mailbox_email,
            top=top,
            all_pages=False
        )

        if not messages_result["success"]:
            return messages_result

        messages = messages_result["data"].get("value", [])

        preview = []

        for message in messages:
            sender = (
                message.get("from", {})
                .get("emailAddress", {})
                .get("address")
            )

            preview.append(
                {
                    "id": message.get("id"),
                    "subject": message.get("subject"),
                    "from": sender,
                    "receivedDateTime": message.get("receivedDateTime"),
                    "hasAttachments": message.get("hasAttachments"),
                    "parentFolderId": message.get("parentFolderId")
                }
            )

        return {
            "success": True,
            "status_code": 200,
            "data": {
                "value": preview
            },
            "items": preview
        }

    def validate_mailbox_access(self, mailbox_email):
        validation_result = {
            "mailbox": mailbox_email,
            "auth": False,
            "user": False,
            "folders": False,
            "inbox": False,
            "messages": False,
            "calendar": False,
            "contacts": False,
            "tasks": False,
            "errors": []
        }

        try:
            self.authenticate()
            validation_result["auth"] = True
        except Exception as error:
            validation_result["errors"].append(
                f"Erro de autenticação: {error}"
            )

            return validation_result

        user_result = self.get_user(mailbox_email)

        if user_result["success"]:
            validation_result["user"] = True

            user_data = user_result["data"]

            self.logger.info(
                f"Usuário localizado: "
                f"{user_data.get('displayName')} | "
                f"{user_data.get('userPrincipalName')}"
            )
        else:
            validation_result["errors"].append(
                f"Erro ao consultar usuário: {user_result['data']}"
            )

        folders_result = self.get_mail_folders(mailbox_email)

        if folders_result["success"]:
            validation_result["folders"] = True

            folders = folders_result["items"]

            self.logger.info(
                f"Pastas encontradas: {len(folders)}"
            )

            for folder in folders:
                self.logger.info(
                    f"Pasta: {folder.get('displayName')}"
                )
        else:
            validation_result["errors"].append(
                f"Erro ao consultar pastas: {folders_result['data']}"
            )

        inbox_result = self.get_inbox(mailbox_email)

        if inbox_result["success"]:
            validation_result["inbox"] = True

            self.logger.info(
                f"Inbox localizada: {inbox_result['data'].get('displayName')}"
            )
        else:
            validation_result["errors"].append(
                f"Erro ao localizar Inbox: {inbox_result['data']}"
            )

        messages_result = self.get_messages(mailbox_email, top=5)

        if messages_result["success"]:
            validation_result["messages"] = True

            messages = messages_result["data"].get("value", [])

            self.logger.info(
                f"Mensagens retornadas: {len(messages)}"
            )

            for message in messages:
                sender = (
                    message.get("from", {})
                    .get("emailAddress", {})
                    .get("address")
                )

                self.logger.info(
                    "Mensagem | "
                    f"Assunto: {message.get('subject')} | "
                    f"Remetente: {sender} | "
                    f"Recebido: {message.get('receivedDateTime')} | "
                    f"Anexos: {message.get('hasAttachments')}"
                )
        else:
            validation_result["errors"].append(
                f"Erro ao consultar mensagens: {messages_result['data']}"
            )

        calendar_result = self.get_calendar_events(mailbox_email, top=5)

        if calendar_result["success"]:
            validation_result["calendar"] = True

            events = calendar_result["data"].get("value", [])

            self.logger.info(
                f"Eventos retornados: {len(events)}"
            )
        else:
            validation_result["errors"].append(
                f"Erro ao consultar calendário: {calendar_result['data']}"
            )

        contacts_result = self.get_contacts(mailbox_email, top=5)

        if contacts_result["success"]:
            validation_result["contacts"] = True

            contacts = contacts_result["data"].get("value", [])

            self.logger.info(
                f"Contatos retornados: {len(contacts)}"
            )
        else:
            validation_result["errors"].append(
                f"Erro ao consultar contatos: {contacts_result['data']}"
            )

        tasks_result = self.get_todo_lists(mailbox_email)

        if tasks_result["success"]:
            validation_result["tasks"] = True

            task_lists = tasks_result["items"]

            self.logger.info(
                f"Listas de tarefas retornadas: {len(task_lists)}"
            )
        else:
            validation_result["errors"].append(
                f"Erro ao consultar tarefas: {tasks_result['data']}"
            )

        return validation_result

    def download_message_mime_to_file(
        self,
        mailbox_email,
        message_id,
        destination_path,
        retry_count=0
    ):
        """Stream one MIME message atomically while limiting the full transfer.

        A MIME slot now remains occupied until the response body has been written,
        so configured concurrency represents actual downloads instead of only the
        time spent waiting for HTTP headers.
        """
        message_id_encoded = self._encode_id(message_id)
        endpoint = (
            f"/users/{mailbox_email}"
            f"/messages/{message_id_encoded}"
            f"/$value"
        )
        url = self._normalize_endpoint(endpoint)
        destination_path = Path(destination_path)
        temporary_path = destination_path.with_suffix(destination_path.suffix + ".part")
        temporary_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path.unlink(missing_ok=True)
        attempt = max(0, int(retry_count or 0))

        while True:
            self._wait_for_shared_throttle()
            self._wait_for_pyrate_limit(mailbox_email)
            self._wait_for_mime_rate_slot()
            self._acquire_mime_slot()
            response = None
            request_started_at = time.perf_counter()
            try:
                response = self._get_thread_session().get(
                    url,
                    headers=self.get_headers(),
                    timeout=(30, 180),
                    stream=True,
                )
                if response.status_code == 401 and attempt == 0:
                    self.access_token = None
                    self.authenticate()
                    attempt += 1
                    continue
                if response.status_code == 429 and attempt < 6:
                    try:
                        retry_after = int(response.headers.get("Retry-After", "10"))
                    except (TypeError, ValueError):
                        retry_after = 10
                    wait_seconds = self._register_mime_throttle(retry_after)
                    self.logger.warning(
                        "Throttling MIME detectado. "
                        f"Aguardando {wait_seconds:.1f} segundos com margem de segurança..."
                    )
                    attempt += 1
                    continue
                if response.status_code in (500, 502, 503, 504) and attempt < 3:
                    wait_seconds = 5 * (attempt + 1)
                    attempt += 1
                    time.sleep(wait_seconds)
                    continue
                if response.status_code < 200 or response.status_code >= 300:
                    error_text = response.text
                    self._record_metric(
                        url, response, request_started_at, attempt, raw=False
                    )
                    return {
                        "success": False,
                        "status_code": response.status_code,
                        "path": None,
                        "bytes_written": 0,
                        "error": error_text,
                    }

                bytes_written = 0
                with open(temporary_path, "wb") as file:
                    for chunk in response.iter_content(
                        chunk_size=EML_DOWNLOAD_CHUNK_SIZE
                    ):
                        if not chunk:
                            continue
                        file.write(chunk)
                        bytes_written += len(chunk)
                    file.flush()
                    os.fsync(file.fileno())
                self._record_metric(
                    url, response, request_started_at, attempt, raw=False
                )
                if bytes_written <= 0:
                    temporary_path.unlink(missing_ok=True)
                    return {
                        "success": False,
                        "status_code": response.status_code,
                        "path": None,
                        "bytes_written": 0,
                        "error": "O Graph retornou conteúdo MIME vazio.",
                    }
                os.replace(temporary_path, destination_path)
                return {
                    "success": True,
                    "status_code": response.status_code,
                    "path": str(destination_path),
                    "bytes_written": bytes_written,
                    "error": None,
                }
            except Exception as error:
                temporary_path.unlink(missing_ok=True)
                return {
                    "success": False,
                    "status_code": 0,
                    "path": None,
                    "bytes_written": 0,
                    "error": str(error),
                }
            finally:
                if response is not None:
                    response.close()
                self._release_mime_slot()

    def get_message_mime_content(self, mailbox_email, message_id):
        message_id_encoded = self._encode_id(message_id)

        endpoint = (
            f"/users/{mailbox_email}"
            f"/messages/{message_id_encoded}"
            f"/$value"
        )

        return self.request_get_raw(
            endpoint=endpoint,
            description=f"Baixando conteúdo MIME da mensagem {message_id}..."
        )

    def build_mail_folder_tree(self, mailbox_email):
        self.logger.info(
            f"Montando árvore de pastas da mailbox: {mailbox_email}"
        )

        root_result = self.get_mail_folders(
            mailbox_email=mailbox_email,
            all_pages=True
        )

        if not root_result["success"]:
            return {
                "success": False,
                "tree": [],
                "flat": [],
                "errors": [
                    f"Erro ao listar pastas raiz: {root_result['data']}"
                ]
            }

        root_folders = root_result["items"]

        tree = []
        flat = []
        errors = []

        for folder in root_folders:
            node = self._build_folder_node_recursive(
                mailbox_email=mailbox_email,
                folder=folder,
                parent_path=""
            )

            tree.append(node)

            flat.extend(
                self._flatten_folder_node(node)
            )

        return {
            "success": len(errors) == 0,
            "tree": tree,
            "flat": flat,
            "errors": errors
        }

    def _build_folder_node_recursive(
        self,
        mailbox_email,
        folder,
        parent_path=""
    ):
        display_name = folder.get("displayName") or "sem_nome"

        if parent_path:
            folder_path = f"{parent_path}/{display_name}"
        else:
            folder_path = display_name

        node = {
            "id": folder.get("id"),
            "displayName": display_name,
            "parentFolderId": folder.get("parentFolderId"),
            "childFolderCount": folder.get("childFolderCount", 0),
            "totalItemCount": folder.get("totalItemCount", 0),
            "unreadItemCount": folder.get("unreadItemCount", 0),
            "path": folder_path,
            "children": []
        }

        child_count = folder.get("childFolderCount", 0)

        if child_count and child_count > 0:
            children_result = self.get_child_folders(
                mailbox_email=mailbox_email,
                folder_id=folder.get("id"),
                all_pages=True
            )

            if children_result["success"]:
                for child_folder in children_result["items"]:
                    child_node = self._build_folder_node_recursive(
                        mailbox_email=mailbox_email,
                        folder=child_folder,
                        parent_path=folder_path
                    )

                    node["children"].append(child_node)
            else:
                self.logger.error(
                    f"Erro ao consultar subpastas de {display_name}: "
                    f"{children_result['data']}"
                )

        return node

    def _flatten_folder_node(self, node):
        folders = [
            {
                "id": node.get("id"),
                "displayName": node.get("displayName"),
                "parentFolderId": node.get("parentFolderId"),
                "childFolderCount": node.get("childFolderCount"),
                "totalItemCount": node.get("totalItemCount"),
                "unreadItemCount": node.get("unreadItemCount"),
                "path": node.get("path")
            }
        ]

        for child in node.get("children", []):
            folders.extend(
                self._flatten_folder_node(child)
            )

        return folders

    def inspect_mailbox(self, mailbox_email):
        self.logger.info(
            f"Iniciando inspeção da mailbox: {mailbox_email}"
        )

        mailbox_result = self.get_mailbox_summary(mailbox_email)

        if not mailbox_result["success"]:
            return {
                "success": False,
                "errors": [
                    f"Erro ao consultar mailbox: {mailbox_result['data']}"
                ]
            }

        folders_result = self.get_mail_folders(mailbox_email)

        if not folders_result["success"]:
            return {
                "success": False,
                "errors": [
                    f"Erro ao consultar pastas: {folders_result['data']}"
                ]
            }

        inbox_result = self.get_inbox(mailbox_email)

        if not inbox_result["success"]:
            return {
                "success": False,
                "errors": [
                    f"Erro ao localizar Inbox: {inbox_result['data']}"
                ]
            }

        preview_result = self.list_email_preview(
            mailbox_email=mailbox_email,
            top=10
        )

        if not preview_result["success"]:
            return {
                "success": False,
                "errors": [
                    f"Erro ao listar mensagens: {preview_result['data']}"
                ]
            }

        return {
            "success": True,
            "mailbox": mailbox_result["data"],
            "folders": folders_result["items"],
            "inbox": inbox_result["data"],
            "messages_preview": preview_result["items"]
        }

    def iter_message_delta_pages(
        self,
        mailbox_email,
        folder_id,
        resume_link=None,
        top=250,
        max_items=None
    ):
        """Yield delta pages and their opaque continuation tokens.

        resume_link can be an @odata.nextLink (resume interrupted initial/incremental
        round) or an @odata.deltaLink (start a new incremental round).
        """
        select_fields = (
            "id,subject,from,sender,toRecipients,ccRecipients,"
            "receivedDateTime,sentDateTime,hasAttachments,importance,"
            "isRead,parentFolderId,lastModifiedDateTime"
        )
        if resume_link:
            next_endpoint = resume_link
            mode = "resume"
        else:
            folder_id_encoded = self._encode_id(folder_id)
            next_endpoint = (
                f"/users/{mailbox_email}"
                f"/mailFolders/{folder_id_encoded}"
                f"/messages/delta"
                f"?$top={max(1, int(top or 250))}"
                f"&$select={select_fields}"
            )
            mode = "initial"

        delivered = 0
        page_number = 0
        while next_endpoint:
            page_number += 1
            result = self.request_get(endpoint=next_endpoint, description=None)
            if not result["success"]:
                status_code = int(result.get("status_code", 0) or 0)
                yield {
                    "success": False,
                    "error": result.get("data"),
                    "status_code": status_code,
                    "reset_required": status_code in (404, 410),
                    "items": [],
                    "removed_ids": [],
                    "page": page_number,
                    "next_link": next_endpoint,
                    "delta_link": None,
                    "mode": mode
                }
                return

            data = result["data"]
            raw_items = data.get("value", [])
            removed_ids = [
                item.get("id") for item in raw_items
                if item.get("id") and "@removed" in item
            ]
            items = [
                item for item in raw_items
                if item.get("id") and "@removed" not in item
            ]
            if max_items is not None:
                remaining = max(0, int(max_items) - delivered)
                items = items[:remaining]
            delivered += len(items)
            next_link = data.get("@odata.nextLink")
            delta_link = data.get("@odata.deltaLink")
            yield {
                "success": True,
                "error": None,
                "status_code": 200,
                "reset_required": False,
                "items": items,
                "removed_ids": removed_ids,
                "page": page_number,
                "next_link": next_link,
                "delta_link": delta_link,
                "mode": mode
            }
            if max_items is not None and delivered >= int(max_items):
                return
            next_endpoint = next_link

    def iter_messages_by_folder_pages(
        self,
        mailbox_email,
        folder_id,
        top=100,
        max_items=None
    ):
        select_fields = (
            "id,"
            "subject,"
            "from,"
            "sender,"
            "toRecipients,"
            "ccRecipients,"
            "receivedDateTime,"
            "sentDateTime,"
            "hasAttachments,"
            "importance,"
            "isRead,"
            "parentFolderId"
        )

        folder_id_encoded = self._encode_id(folder_id)

        endpoint = (
            f"/users/{mailbox_email}"
            f"/mailFolders/{folder_id_encoded}"
            f"/messages"
            f"?$top={top}"
            f"&$select={select_fields}"
            f"&$orderby=receivedDateTime desc"
        )

        exported_count = 0
        page_number = 0
        next_endpoint = endpoint

        while next_endpoint:
            page_number += 1

            self.logger.info(
                f"Listando mensagens com paginação... Página {page_number}"
            )

            result = self.request_get(
                endpoint=next_endpoint,
                description=None
            )

            if not result["success"]:
                yield {
                    "success": False,
                    "error": result.get("data"),
                    "items": [],
                    "page": page_number
                }
                return

            data = result["data"]
            items = data.get("value", [])

            if max_items is not None:
                remaining = max_items - exported_count

                if remaining <= 0:
                    return

                items = items[:remaining]

            exported_count += len(items)

            next_link = data.get("@odata.nextLink")
            limit_reached = (
                max_items is not None and exported_count >= max_items
            )
            yield {
                "success": True,
                "error": None,
                "items": items,
                "page": page_number,
                "next_link": next_link,
                "terminal": not bool(next_link) or limit_reached,
                "terminal_reason": (
                    "limit_reached" if limit_reached
                    else "end_of_collection" if not next_link
                    else None
                ),
                "limited_scope": bool(max_items is not None),
                "delivered_items": exported_count,
                "page_items": len(items)
            }

            if limit_reached:
                return

            next_endpoint = next_link
