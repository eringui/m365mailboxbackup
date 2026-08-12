import argparse
import asyncio
import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

import psutil
import uvicorn
from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
from pydantic import BaseModel, Field

try:
    from src.services.operation_store import OperationStore
except ImportError:
    from operation_store import OperationStore  # pyright: ignore[reportMissingImports]

HOST = "127.0.0.1"
PORT = 8765


def utc_now():
    return datetime.now(timezone.utc).isoformat()


class OperationCreate(BaseModel):
    operation_type: str = Field(pattern="^(backup|pst)$")
    mailbox: str | None = None
    command: list[str]
    source_path: str | None = None
    destination_path: str | None = None
    backup_path: str | None = None
    options: dict = Field(default_factory=dict)


class QueueOrder(BaseModel):
    operation_ids: list[str]


class ConcurrencySettings(BaseModel):
    backup_workers: int = Field(ge=1, le=20)
    pst_workers: int = Field(ge=1, le=5)


class Coordinator:
    def __init__(self, project_root):
        self.project_root = Path(project_root).resolve()
        self.state_dir = self.project_root / "_gui_state"
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.store = OperationStore(self.state_dir / "operations.sqlite3")
        self.processes = {}
        self.process_lock = threading.RLock()
        self.stop_event = threading.Event()
        self.started_at = utc_now()
        self.websockets = set()
        self.event_loop: asyncio.AbstractEventLoop | None = None
        self._last_progress_event_at = {}
        self._last_progress_event_value = {}
        self.store.recover_startup_state()
        self.scheduler = threading.Thread(target=self._scheduler_loop, daemon=True)
        self.scheduler.start()


    def stop_all_operations(self, timeout=12):
        """Stop every child process because the GUI and engine are one application."""
        self.stop_event.set()
        with self.process_lock:
            processes = list(self.processes.items())
        for operation_id, process in processes:
            if process.poll() is not None:
                continue
            self.store.update_operation(
                operation_id,
                status="pausing",
                pause_requested=True,
                heartbeat_at=utc_now()
            )
            self._event(
                operation_id,
                "application_closing",
                message="A aplicação está sendo encerrada; o progresso será preservado."
            )
            try:
                process.terminate()
            except Exception:
                pass
        deadline = time.monotonic() + max(1, timeout)
        for operation_id, process in processes:
            remaining = max(0.0, deadline - time.monotonic())
            try:
                process.wait(timeout=remaining)
            except Exception:
                try:
                    process.kill()
                except Exception:
                    pass
        with self.process_lock:
            self.processes.clear()

    def watch_parent(self, parent_pid):
        if not parent_pid:
            return
        def monitor():
            while not self.stop_event.wait(1.0):
                if not psutil.pid_exists(parent_pid):
                    self.stop_all_operations()
                    request_server_shutdown()
                    return
        threading.Thread(target=monitor, daemon=True).start()

    def _event(self, operation_id, event_type, severity="info", message=None,
               technical=None, payload=None):
        event = self.store.add_event(
            operation_id, event_type, severity, message, technical, payload
        )
        if self.event_loop and self.event_loop.is_running():
            asyncio.run_coroutine_threadsafe(self._broadcast(event), self.event_loop)
        return event

    async def _broadcast(self, event):
        dead = []
        for websocket in list(self.websockets):
            try:
                await websocket.send_json(event)
            except Exception:
                dead.append(websocket)
        for websocket in dead:
            self.websockets.discard(websocket)

    def create(self, request):
        command = [str(value) for value in request.command]
        if not command:
            raise ValueError("Comando da operação não informado.")
        if request.operation_type == "pst":
            destination = os.path.normcase(os.path.abspath(str(request.destination_path or "")))
            for existing in self.store.list_operations() or []:
                if not isinstance(existing, dict):
                    continue
                if existing.get("operation_type") != "pst":
                    continue
                existing_destination = os.path.normcase(
                    os.path.abspath(str(existing.get("destination_path") or ""))
                )
                if (
                    destination
                    and destination == existing_destination
                    and existing.get("status")
                    not in ("completed", "cancelled", "failed")
                ):
                    raise ValueError(
                        "Já existe uma conversão PST ativa ou pendente para este arquivo de destino."
                    )
        backup_path = request.backup_path
        if request.operation_type == "backup" and not backup_path:
            mailbox = str(request.mailbox or "sem_nome").strip()
            folder_name = mailbox
            for character in '\\/:*?"<>|':
                folder_name = folder_name.replace(character, "_")
            folder_name = " ".join(folder_name.split()).strip(" .")[:120] or "sem_nome"
            backup_path = str(
                Path(request.destination_path or self.project_root).expanduser().resolve()
                / folder_name
            )
        return self.store.create_operation(
            operation_type=request.operation_type,
            command=command,
            mailbox=request.mailbox,
            source_path=request.source_path,
            destination_path=request.destination_path,
            backup_path=backup_path,
            options=request.options,
            status=(
                "pending"
                if request.operation_type == "pst"
                and bool((request.options or {}).get("manual_start"))
                else "queued"
            ),
        )

    @staticmethod
    def _setting_as_int(value, default):
        """Converte uma configuracao numerica, aceitando None e valores persistidos como texto."""
        if value is None:
            return int(default)
        try:
            return int(str(value))
        except (TypeError, ValueError):
            return int(default)

    def start_queued(self):
        raw_operations = self.store.list_operations(include_finished=False) or []
        operations = [item for item in raw_operations if isinstance(item, dict)]
        running = {"backup": 0, "pst": 0}
        for operation in operations:
            if operation["status"] in ("starting", "running", "pausing"):
                running[operation["operation_type"]] += 1
        limits = {
            "backup": self._setting_as_int(
                self.store.get_setting("backup_concurrency", 2), 2
            ),
            "pst": self._setting_as_int(
                self.store.get_setting("pst_concurrency", 1), 1
            ),
        }
        for operation in operations:
            kind = operation["operation_type"]
            if operation["status"] != "queued" or running[kind] >= limits[kind]:
                continue
            self._start_operation(operation)
            running[kind] += 1

    RUNTIME_ENV_PREFIXES = ("M365_",)
    RUNTIME_ENV_EXACT = {"GRAPH_SCOPE", "GRAPH_URL"}
    RUNTIME_ENV_BLOCKED = {
        "M365_OPERATION_ID", "M365_OPERATION_DB", "M365_OPERATION_TOKEN",
        "M365_API_METRICS_DB", "M365_RATE_LIMITER_DB"
    }

    @classmethod
    def _validated_runtime_environment(cls, options):
        snapshot = (options or {}).get("runtime_settings") or {}
        values = snapshot.get("environment") if isinstance(snapshot, dict) else {}
        result = {}
        if not isinstance(values, dict):
            return result
        for key, value in values.items():
            key = str(key).strip()
            allowed = key in cls.RUNTIME_ENV_EXACT or any(key.startswith(prefix) for prefix in cls.RUNTIME_ENV_PREFIXES)
            if not allowed or key in cls.RUNTIME_ENV_BLOCKED or not key:
                continue
            result[key] = str(value)
        return result

    def _start_operation(self, operation):
        operation_id = operation["operation_id"]
        environment = os.environ.copy()
        runtime_environment = self._validated_runtime_environment(operation.get("options") or {})
        environment.update(runtime_environment)
        environment["PYTHONIOENCODING"] = "utf-8"
        environment["PYTHONUTF8"] = "1"
        environment["M365_OPERATION_ID"] = operation_id
        environment["M365_OPERATION_DB"] = str(self.store.path)
        environment["M365_OPERATION_TOKEN"] = operation["operation_token"]
        environment["M365_API_METRICS_DB"] = str(self.state_dir / "api_metrics.sqlite3")
        environment["M365_RATE_LIMITER_DB"] = str(
            self.state_dir / "graph_rate_limits.sqlite3"
        )
        self.store.clear_requests(operation_id)
        self.store.update_operation(
            operation_id, status="starting", started_at=utc_now(), finished_at=None,
            last_error=None
        )
        try:
            process = subprocess.Popen(
                operation["command"], cwd=str(self.project_root),
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", env=environment,
                creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
            )
            with self.process_lock:
                self.processes[operation_id] = process
            self.store.update_operation(
                operation_id, status="running", pid=process.pid,
                process_created_at=datetime.fromtimestamp(
                    psutil.Process(process.pid).create_time(), timezone.utc
                ).isoformat(), heartbeat_at=utc_now()
            )
            operation_before_start = operation
            is_resume = bool(operation_before_start.get("backup_path"))
            self._event(
                operation_id,
                "operation_started",
                message=(
                    "Processo de retomada iniciado; carregando o checkpoint."
                    if is_resume else "Operação iniciada."
                ),
                payload={
                    "pid": process.pid, "resume": is_resume,
                    "runtime_settings_applied": bool(runtime_environment),
                    "runtime_environment": runtime_environment
                }
            )
            threading.Thread(
                target=self._monitor_process, args=(operation_id, process), daemon=True
            ).start()
        except Exception as error:
            self.store.update_operation(
                operation_id, status="failed", last_error=str(error), finished_at=utc_now()
            )
            self._event(operation_id, "operation_failed", "error",
                        "Não foi possível iniciar a operação.", str(error))

    def _monitor_process(self, operation_id, process):
        code = -1
        try:
            stdout = process.stdout
            if stdout is None:
                raise RuntimeError("A saida padrao do processo nao esta disponivel.")
            for line in stdout:
                self._handle_line(operation_id, line.rstrip())
                operation = self.store.get_operation(operation_id)
                if not operation:
                    continue
                if operation["pause_requested"] and process.poll() is None:
                    self.store.update_operation(operation_id, status="pausing")
                    self._event(operation_id, "operation_pausing", message="Pausa solicitada; salvando o ponto atual.")
                    process.terminate()
                elif operation["cancel_requested"] and process.poll() is None:
                    process.terminate()
            code = process.wait()
        except Exception as error:
            code = -1
            self._event(operation_id, "monitor_failed", "error",
                        "Falha ao acompanhar a operação.", str(error))
        finally:
            with self.process_lock:
                self.processes.pop(operation_id, None)
            operation = self.store.get_operation(operation_id)
            if not operation:
                return
            if operation["pause_requested"]:
                status, message = "paused", "Operação pausada com o progresso preservado."
            elif operation["cancel_requested"]:
                status, message = "cancelled", "Operação cancelada."
            elif code == 0:
                current_items = int(operation.get("current_items", 0) or 0)
                total_items = int(operation.get("total_items", 0) or 0)
                if operation.get("operation_type") == "pst":
                    pending_verifications = int(operation.get("pst_pending_verifications", 0) or 0)
                    audit_failures = int(operation.get("pst_audit_failures", 0) or 0)
                    if total_items > 0 and current_items >= total_items and pending_verifications == 0 and audit_failures == 0:
                        status, message = "completed", "Conversão PST concluída e auditada."
                    else:
                        status, message = "incomplete", (
                            "Conversão PST terminou sem importar todos os EML: "
                            f"{current_items}/{total_items or '?'} confirmados."
                        )
                elif total_items > 0 and current_items >= total_items:
                    status, message = "completed", "Operação concluída e validada."
                else:
                    status = "incomplete"
                    message = (
                        "O processo terminou sem comprovar o backup integral: "
                        f"{current_items}/{total_items or '?'} EML confirmados."
                    )
            else:
                current_items = int(operation.get("current_items", 0) or 0)
                total_items = int(operation.get("total_items", 0) or 0)
                if total_items > 0 and current_items < total_items:
                    status = "incomplete"
                    message = (
                        "Backup não concluído; o progresso foi preservado para retomada: "
                        f"{current_items}/{total_items} EML confirmados."
                    )
                else:
                    status, message = "failed", f"Operação encerrada com código {code}."
            self.store.update_operation(
                operation_id, status=status, pid=None, finished_at=utc_now(),
                heartbeat_at=utc_now(), last_error=None if status != "failed" else message
            )
            self._event(operation_id, f"operation_{status}",
                        "error" if status == "failed" else "info", message,
                        payload={"exit_code": code})

    def _handle_line(self, operation_id, line):
        self.store.heartbeat(operation_id)
        lower_line = str(line).lower()
        resume_stages = (
            (("carregando checkpoint", "checkpoint existente"), "Carregando checkpoint"),
            (("retomada rápida", "reutilizados sem novas consultas"), "Reutilizando metadados locais"),
            (("retomada paginada", "continuando diretamente"), "Recuperando página pendente"),
            (("exportando pasta", "enumeração integral"), "Preparando mensagens pendentes"),
            (("[ok] eml salvo",), "Baixando EML novamente"),
            (("preparação mime paralela", "preparacao mime paralela"), "Preparando EML em paralelo"),
            (("primeiro eml de continuação confirmado",), "Primeiro EML confirmado no PST")
        )
        for terms, stage in resume_stages:
            if any(term in lower_line for term in terms):
                self._event(
                    operation_id,
                    "operation_resume_stage",
                    message=stage,
                    technical=line,
                    payload={"stage": stage}
                )
                break
        for marker in ("[PROGRESS] ", "[PST-PROGRESS] "):
            if marker in line:
                try:
                    payload = json.loads(line.split(marker, 1)[1].strip())
                    limiter = payload.get("rate_limiter") or {}
                    verification = payload.get("verification") or {}
                    pipeline = payload.get("pipeline") or {}
                    resume_metrics = payload.get("resume") or {}
                    options = (self.store.get_operation(operation_id) or {}).get("options") or {}
                    self.store.heartbeat(
                        operation_id,
                        current_items=int(payload.get("current", 0) or 0),
                        total_items=int(payload.get("expected", payload.get("total", 0)) or 0),
                        failed_items=int(payload.get("failed", 0) or 0),
                        downloaded_bytes=int(payload.get("downloaded_bytes", 0) or 0),
                        current_folder=payload.get("folder"),
                        current_page=int(payload.get("page", 0) or 0),
                        sync_mode=payload.get("sync_mode"),
                        rate_limiter_enabled=bool(limiter.get("enabled", False)),
                        rate_limiter_profile=limiter.get("profile"),
                        rate_limiter_wait_seconds=float(limiter.get("wait_seconds", 0) or 0),
                        rate_limiter_wait_events=int(limiter.get("wait_events", 0) or 0),
                        mime_rate_second=float(limiter.get("mailbox_rate_second", 0) or 0),
                        mime_concurrency=int(limiter.get("mime_concurrency", 0) or 0),
                        pst_saved_items=int(verification.get("saved", 0) or 0),
                        pst_verified_items=int(verification.get("verified", 0) or 0),
                        pst_pending_verifications=int(verification.get("pending", 0) or 0),
                        pst_reconciled_items=int(verification.get("reconciled", 0) or 0),
                        pst_verification_attempts=int(verification.get("attempts", 0) or 0),
                        pst_audit_failures=int(verification.get("audit_failures", 0) or 0),
                        pst_verification_mode=verification.get("mode"),
                        pst_performance_profile=options.get("performance_profile", "balanced"),
                        pst_bottleneck=pipeline.get("bottleneck"),
                        pst_effective_workers=int(pipeline.get("effective_workers", 0) or 0),
                        pst_effective_queue_limit=int(pipeline.get("effective_queue_limit", 0) or 0),
                        pst_queue_bytes=int(pipeline.get("queue_bytes", 0) or 0),
                        pst_peak_queue_bytes=int(pipeline.get("peak_queue_bytes", 0) or 0),
                        pst_peak_rss_bytes=int(pipeline.get("peak_rss_bytes", 0) or 0),
                        pst_prepare_seconds=float(pipeline.get("prepare_seconds", 0) or 0),
                        pst_com_seconds=float(pipeline.get("com_seconds", 0) or 0),
                        pst_queue_wait_seconds=float(pipeline.get("queue_wait_seconds", 0) or 0),
                        pst_adaptive_adjustments=int(pipeline.get("adaptive_adjustments", 0) or 0),
                        pst_eta_seconds=float(pipeline.get("eta_seconds", 0) or 0),
                        pst_memory_pressure=bool(pipeline.get("memory_pressure", False)),
                        pst_resume_total_seconds=float(resume_metrics.get("total_seconds", 0) or 0),
                        pst_resume_checkpoint_seconds=float(resume_metrics.get("checkpoint_seconds", 0) or 0),
                        pst_resume_outlook_seconds=float(resume_metrics.get("outlook_seconds", 0) or 0),
                        pst_resume_pending_query_seconds=float(resume_metrics.get("pending_query_seconds", 0) or 0),
                        pst_resume_first_item_seconds=float(resume_metrics.get("first_item_seconds", 0) or 0),
                        pst_resume_first_selected_seconds=float(resume_metrics.get("first_selected_seconds", 0) or 0),
                        pst_resume_first_prepared_seconds=float(resume_metrics.get("first_prepared_seconds", 0) or 0),
                        pst_resume_first_committed_seconds=float(resume_metrics.get("first_committed_seconds", 0) or 0),
                        pst_resume_first_source_position=int(resume_metrics.get("first_source_position", 0) or 0),
                        pst_resume_first_commit_target_seconds=float(resume_metrics.get("first_commit_target_seconds", 10) or 10),
                        pst_resume_first_commit_target_met=bool(resume_metrics.get("first_commit_target_met", False)),
                        pst_resume_skipped_before_parse=int(resume_metrics.get("skipped_before_parse", 0) or 0),
                        pst_resume_eligible_items=int(resume_metrics.get("eligible_items", 0) or 0),
                        pst_resume_failure_reason=resume_metrics.get("failure_reason"),
                        pst_capacity_blocked_items=int(resume_metrics.get("capacity_blocked_items", 0) or 0)
                    )
                    now = time.monotonic()
                    current_value = int(payload.get("current", 0) or 0)
                    last_at = float(self._last_progress_event_at.get(operation_id, 0.0))
                    last_value = int(self._last_progress_event_value.get(operation_id, -1))
                    total_value = int(payload.get("expected", payload.get("total", 0)) or 0)
                    important = current_value == total_value and total_value > 0
                    if important or current_value != last_value and now - last_at >= 0.75:
                        self._last_progress_event_at[operation_id] = now
                        self._last_progress_event_value[operation_id] = current_value
                        self._event(operation_id, "operation_progress", payload=payload)
                except Exception:
                    pass
                return
        if "Throttling" in line or "HTTP 429" in line:
            self._event(operation_id, "graph_throttling", "warning",
                        "O Microsoft Graph solicitou uma redução temporária da velocidade.", line)
        elif " | ERROR | " in line or "[ERRO]" in line or "[FALHA]" in line:
            self._event(operation_id, "operation_log_error", "error",
                        "O processo informou um problema.", line)

    def _force_pause_deadline(self, operation_id, process, deadline_seconds=12.0):
        deadline = time.monotonic() + max(1.0, float(deadline_seconds))
        while process.poll() is None and time.monotonic() < deadline:
            time.sleep(0.20)
        if process.poll() is None:
            self._event(
                operation_id,
                "pause_deadline_reached",
                "warning",
                "O processo excedeu o limite de pausa e será interrompido agora."
            )
            try:
                process.kill()
            except Exception:
                pass

    def pause(self, operation_id):
        operation = self.store.get_operation(operation_id)
        if not operation:
            raise KeyError(operation_id)
        if operation["status"] == "queued":
            return self.store.update_operation(
                operation_id,
                status="paused",
                pause_requested=True,
                heartbeat_at=utc_now()
            )
        if operation["status"] not in ("starting", "running", "pause_requested"):
            return operation

        result = self.store.request_pause(operation_id)
        self._event(
            operation_id,
            "operation_pause_requested",
            message="Pausa solicitada; interrompendo o processo em até 15 segundos."
        )

        # Important: do not wait for another stdout line. The previous implementation
        # only terminated inside _monitor_process, which can block indefinitely in
        # `for line in process.stdout` when the backend is silent.
        with self.process_lock:
            process = self.processes.get(operation_id)
        if process is not None and process.poll() is None:
            self.store.update_operation(
                operation_id,
                status="pausing",
                heartbeat_at=utc_now()
            )
            try:
                process.terminate()
            except Exception as error:
                self._event(
                    operation_id,
                    "pause_signal_failed",
                    "warning",
                    "O primeiro sinal de pausa falhou; o limite de segurança continuará ativo.",
                    str(error)
                )
            threading.Thread(
                target=self._force_pause_deadline,
                args=(operation_id, process, 12.0),
                daemon=True
            ).start()
        return self.store.get_operation(operation_id)

    def resume(self, operation_id):
        operation = self.store.get_operation(operation_id)
        if not operation:
            raise KeyError(operation_id)
        if operation["status"] not in (
            "pending", "paused", "interrupted", "failed", "incomplete"
        ):
            return operation
        self.store.clear_requests(operation_id)
        result = self.store.update_operation(
            operation_id,
            status="queued",
            pid=None,
            finished_at=None,
            last_error=None,
            heartbeat_at=utc_now()
        )
        self._event(
            operation_id,
            "operation_queued",
            message="Operação preparada para retomada imediata."
        )
        threading.Thread(target=self.start_queued, daemon=True).start()
        return result

    def cancel(self, operation_id):
        operation = self.store.get_operation(operation_id)
        if not operation:
            raise KeyError(operation_id)
        if operation["status"] == "queued":
            return self.store.update_operation(operation_id, status="cancelled", finished_at=utc_now())
        return self.store.request_cancel(operation_id)

    def _scheduler_loop(self):
        while not self.stop_event.is_set():
            try:
                self.start_queued()
                self._check_processes()
            except Exception:
                pass
            self.stop_event.wait(0.75)

    def _check_processes(self):
        with self.process_lock:
            items = list(self.processes.items())
        for operation_id, process in items:
            if process.poll() is None:
                self.store.heartbeat(operation_id)


PROJECT_ROOT = (Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent)
coordinator = Coordinator(PROJECT_ROOT)
app = FastAPI(title="M365 Mailbox Backup Coordinator", docs_url=None, redoc_url=None)



@app.on_event("startup")
async def startup():
    coordinator.event_loop = asyncio.get_running_loop()


@app.get("/api/v1/health")
def health():
    return {
        "status": "ok", "coordinator_pid": os.getpid(),
        "started_at": coordinator.started_at,
        "active_operations": len(coordinator.processes)
    }


@app.get("/api/v1/operations")
def operations():
    return coordinator.store.list_operations()


@app.post("/api/v1/operations")
def create_operation(request: OperationCreate):
    return coordinator.create(request)


@app.post("/api/v1/operations/{operation_id}/pause")
def pause(operation_id: str):
    try:
        return coordinator.pause(operation_id)
    except KeyError:
        raise HTTPException(404, "Operação não encontrada")


@app.post("/api/v1/operations/{operation_id}/resume")
def resume(operation_id: str):
    try:
        return coordinator.resume(operation_id)
    except KeyError:
        raise HTTPException(404, "Operação não encontrada")


@app.post("/api/v1/operations/{operation_id}/cancel")
def cancel(operation_id: str):
    try:
        return coordinator.cancel(operation_id)
    except KeyError:
        raise HTTPException(404, "Operação não encontrada")


@app.delete("/api/v1/operations/{operation_id}")
def delete(operation_id: str):
    if not coordinator.store.delete_operation(operation_id):
        raise HTTPException(409, "Operação ativa ou inexistente")
    return {"deleted": True}


@app.put("/api/v1/queue/order")
def reorder(request: QueueOrder):
    return coordinator.store.set_queue_order(request.operation_ids)


@app.put("/api/v1/settings/concurrency")
def concurrency(request: ConcurrencySettings):
    coordinator.store.set_setting("backup_concurrency", request.backup_workers)
    coordinator.store.set_setting("pst_concurrency", request.pst_workers)
    return request.model_dump()


@app.get("/api/v1/events")
def events(after_id: int = 0, limit: int = 1000):
    return coordinator.store.list_events(after_id, limit)


@app.websocket("/api/v1/events/ws")
async def event_socket(websocket: WebSocket):
    await websocket.accept()
    coordinator.websockets.add(websocket)
    try:
        await websocket.send_json({"event_type": "connected", "created_at": utc_now()})
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        pass
    finally:
        coordinator.websockets.discard(websocket)


server_instance = None


def request_server_shutdown():
    global server_instance
    if server_instance is not None:
        server_instance.should_exit = True


@app.post("/api/v1/shutdown")
def shutdown_application():
    def shutdown_worker():
        coordinator.stop_all_operations()
        request_server_shutdown()
    threading.Thread(target=shutdown_worker, daemon=True).start()
    return {"status": "stopping"}


@app.on_event("shutdown")
def coordinator_shutdown_event():
    coordinator.stop_all_operations(timeout=5)


def main():
    global server_instance
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default=HOST)
    parser.add_argument("--port", type=int, default=PORT)
    parser.add_argument("--parent-pid", type=int, default=0)
    args = parser.parse_args()
    if args.host not in ("127.0.0.1", "localhost"):
        raise SystemExit("O coordenador aceita somente conexão local.")
    coordinator.watch_parent(args.parent_pid)
    config = uvicorn.Config(
        app, host="127.0.0.1", port=args.port,
        log_level="warning", workers=1
    )
    server_instance = uvicorn.Server(config)
    server_instance.run()


if __name__ == "__main__":
    main()
