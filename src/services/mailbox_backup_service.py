import base64
import csv
import json
import hashlib
import os
import re
import tempfile
import threading
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from datetime import datetime
from email.parser import BytesHeaderParser
from pathlib import Path

from src.config.settings import (
    OUTPUT_DIR,
    EML_DOWNLOAD_WORKERS,
    EML_MAX_PENDING_DOWNLOADS,
    MIME_MAX_CONCURRENCY,
    MIME_RESUME_MAX_CONCURRENCY,
)

try:
    from src.services.checkpoint_store import CheckpointStore
    from src.services.operation_control import OperationControl, OperationInterrupted
except ImportError:
    from checkpoint_store import CheckpointStore
    from operation_control import OperationControl, OperationInterrupted


class MailboxBackupService:
    CHECKPOINT_BATCH_SIZE = 50

    DEFAULT_EXCLUDED_PROFILE_FOLDER_NAMES = {
        "arquivo morto",
        "archive",
        "online archive",
        "in-place archive",
        "recoverable items",
        "deletions",
        "purges",
        "versions",
        "audits",
        "discoveryholds",
        "sync issues",
        "problemas de sincronização",
        "conversation history",
        "histórico de conversas",
        "rss feeds",
        "feeds rss"
    }

    def __init__(self, graph_service, logger):
        self.graph_service = graph_service
        self.logger = logger
        self.json_write_lock = threading.RLock()
        self.eml_executor = ThreadPoolExecutor(
            max_workers=EML_DOWNLOAD_WORKERS,
            thread_name_prefix="m365-eml"
        )
        self._last_progress_log_at = 0.0
        self._last_progress_log_current = -1
        self._last_checkpoint_json_at = 0.0
        self.operation_control = OperationControl()

    def sanitize_name(self, value, max_length=80):
        if not value:
            return "sem_nome"

        value = str(value).strip()

        value = re.sub(
            r'[\\/:*?"<>|]+',
            "_",
            value
        )

        value = re.sub(
            r"\s+",
            " ",
            value
        )

        value = value.strip(" .")

        if not value:
            value = "sem_nome"

        return value[:max_length]

    def should_skip_folder(self, folder_path, excluded_folder_names=None):
        if not folder_path:
            return False

        normalized = str(folder_path).strip().lower()
        # Compara por segmento de caminho (separado por "/"), não por substring
        # crua: nomes genéricos como "Versions"/"Audits" não devem casar com
        # pastas reais que apenas contêm esse texto (ex.: "Auditssociados").
        segments = {
            segment.strip() for segment in normalized.split("/") if segment.strip()
        }

        excluded = set()

        if excluded_folder_names:
            for item in excluded_folder_names:
                if item:
                    excluded.add(str(item).strip().lower())

        return bool(excluded & segments)

    def create_backup_structure(self, mailbox_email):
        """Use a single stable backup root for each mailbox."""
        mailbox_folder_name = self.sanitize_name(mailbox_email, max_length=120)
        backup_dir = OUTPUT_DIR / mailbox_folder_name
        return self.create_backup_structure_from_root(backup_dir)

    def create_backup_structure_from_root(self, backup_dir):
        backup_dir = Path(backup_dir)

        folders = {
            "root": backup_dir,
            "metadata": backup_dir / "metadata",
            "mail": backup_dir / "mail",
            "mail_folders": backup_dir / "mail" / "folders",
            "eml": backup_dir / "mail" / "eml",
            "calendar": backup_dir / "calendar",
            "contacts": backup_dir / "contacts",
            "tasks": backup_dir / "tasks",
            "logs": backup_dir / "logs",
        }

        for folder in folders.values():
            folder.mkdir(
                parents=True,
                exist_ok=True
            )

        return folders

    def save_json(self, path, data):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        with self.json_write_lock:
            temporary_path = None
            try:
                with tempfile.NamedTemporaryFile(
                    mode="w",
                    encoding="utf-8",
                    prefix=f".{path.name}.",
                    suffix=".tmp",
                    dir=str(path.parent),
                    delete=False
                ) as file:
                    temporary_path = Path(file.name)
                    json.dump(data, file, ensure_ascii=False, indent=4)
                    file.flush()
                    os.fsync(file.fileno())

                last_error = None
                for attempt in range(8):
                    try:
                        os.replace(temporary_path, path)
                        temporary_path = None
                        return
                    except PermissionError as error:
                        last_error = error
                        wait_seconds = min(0.05 * (2 ** attempt), 2.0)
                        self.logger.warning(
                            f"Arquivo JSON temporariamente bloqueado: {path.name}. "
                            f"Nova tentativa em {wait_seconds:.2f}s "
                            f"({attempt + 1}/8)."
                        )
                        time.sleep(wait_seconds)
                    except OSError as error:
                        last_error = error
                        wait_seconds = min(0.05 * (2 ** attempt), 2.0)
                        time.sleep(wait_seconds)

                raise last_error or OSError(
                    f"Não foi possível substituir o arquivo JSON: {path}"
                )
            finally:
                if temporary_path is not None:
                    try:
                        temporary_path.unlink(missing_ok=True)
                    except Exception:
                        pass

    def load_json(self, path):
        with open(
            path,
            "r",
            encoding="utf-8"
        ) as file:
            return json.load(file)

    def save_bytes(self, path, content):
        path.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        with open(
            path,
            "wb"
        ) as file:
            file.write(content)

    def build_message_filename(self, index, message):
        received = message.get("receivedDateTime", "")

        received = (
            received
            .replace(":", "-")
            .replace(".", "-")
            .replace("Z", "")
        )

        if not received:
            received = "sem_data"

        subject = self.sanitize_name(
            message.get("subject") or "sem_assunto",
            max_length=60
        )

        message_id_safe = self.sanitize_name(
            message.get("id", "")[-12:],
            max_length=20
        )

        filename = (
            f"{index:06d}_"
            f"{received}_"
            f"{subject}_"
            f"{message_id_safe}.eml"
        )

        return filename

    def build_attachment_filename(self, index, attachment):
        name = attachment.get("name") or "anexo"

        name = self.sanitize_name(
            name,
            max_length=100
        )

        attachment_id_safe = self.sanitize_name(
            attachment.get("id", "")[-10:],
            max_length=20
        )

        return f"{index:04d}_{attachment_id_safe}_{name}"

    def build_folder_local_path(self, base_folder, folder_path):
        parts = str(folder_path or "sem_pasta").split("/")

        safe_parts = []

        for part in parts:
            safe_parts.append(
                self.sanitize_name(
                    part,
                    max_length=60
                )
            )

        return base_folder.joinpath(*safe_parts)

    def create_checkpoint(self, mailbox_email):
        return {
            "project": "M365 Mailbox Backup",
            "mailbox": mailbox_email,
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "status": "running",
            "layout_version": 2,
            "stable_mailbox_root": True,
            "exported_message_ids": [],
            "exported_message_count": 0,
            "expected_message_count": 0,
            "downloaded_bytes": 0,
            "bytes_baseline_complete": True,
            "expected_bytes": 0,
            "discovered_message_bytes": 0,
            "scope_hash": None,
            "progress_updated_at": datetime.now().isoformat(),
            "progress_consistent": True,
            "failed_message_ids": [],
            "processed_folder_ids": [],
            "errors": []
        }

    def load_or_create_checkpoint(self, backup_folders, mailbox_email):
        checkpoint_path = backup_folders["root"] / "checkpoint.json"

        if checkpoint_path.exists():
            self.logger.info(
                f"Carregando checkpoint existente: {checkpoint_path}"
            )

            return self.load_json(checkpoint_path)

        checkpoint = self.create_checkpoint(
            mailbox_email
        )

        self.save_checkpoint(
            backup_folders,
            checkpoint
        )

        return checkpoint

    def save_checkpoint(self, backup_folders, checkpoint):
        checkpoint["updated_at"] = datetime.now().isoformat()

        self.save_json(
            backup_folders["root"] / "checkpoint.json",
            checkpoint
        )

    def sync_checkpoint_sets(self, checkpoint, exported_message_ids,
                             failed_message_ids, processed_folder_ids):
        checkpoint["exported_message_count"] = len(exported_message_ids)
        checkpoint["failed_message_count"] = len(failed_message_ids)
        checkpoint["processed_folder_ids"] = list(processed_folder_ids)
        checkpoint["checkpoint_engine"] = "sqlite-wal"
        checkpoint["checkpoint_sqlite"] = "checkpoint.sqlite3"
        # Legacy arrays are intentionally compacted after automatic migration.
        checkpoint["exported_message_ids"] = []
        checkpoint["failed_message_ids"] = []

    def flush_checkpoint(self, backup_folders, checkpoint,
                         exported_message_ids, failed_message_ids,
                         processed_folder_ids):
        self.sync_checkpoint_sets(
            checkpoint,
            exported_message_ids,
            failed_message_ids,
            processed_folder_ids
        )
        self.save_checkpoint(backup_folders, checkpoint)

    def export_common_data(
        self,
        mailbox_email,
        backup_folders,
        result,
        skip_calendar=False,
        skip_contacts=False,
        skip_tasks=False
    ):
        mailbox_result = self.graph_service.get_mailbox_summary(
            mailbox_email
        )

        if not mailbox_result["success"]:
            result["errors"].append(
                f"Erro ao consultar mailbox: {mailbox_result['data']}"
            )

            return False

        self.save_json(
            backup_folders["metadata"] / "mailbox.json",
            mailbox_result["data"]
        )

        if not skip_calendar:
            calendar_result = self.graph_service.get_calendar_events(
                mailbox_email=mailbox_email,
                top=50,
                all_pages=True
            )

            if not calendar_result["success"]:
                result["errors"].append(
                    f"Erro ao consultar calendário: {calendar_result['data']}"
                )
            else:
                events = calendar_result["items"]

                result["calendar_events"] = len(events)

                self.save_json(
                    backup_folders["calendar"] / "events.json",
                    events
                )
        else:
            self.logger.info("Exportação de calendário ignorada por parâmetro.")

        if not skip_contacts:
            contacts_result = self.graph_service.get_contacts(
                mailbox_email=mailbox_email,
                top=50,
                all_pages=True
            )

            if not contacts_result["success"]:
                result["errors"].append(
                    f"Erro ao consultar contatos: {contacts_result['data']}"
                )
            else:
                contacts = contacts_result["items"]

                result["contacts"] = len(contacts)

                self.save_json(
                    backup_folders["contacts"] / "contacts.json",
                    contacts
                )
        else:
            self.logger.info("Exportação de contatos ignorada por parâmetro.")

        if not skip_tasks:
            tasks_result = self.graph_service.get_todo_lists(
                mailbox_email=mailbox_email,
                all_pages=True
            )

            if not tasks_result["success"]:
                result["errors"].append(
                    f"Erro ao consultar listas de tarefas: {tasks_result['data']}"
                )
            else:
                task_lists = tasks_result["items"]

                result["task_lists"] = len(task_lists)

                self.save_json(
                    backup_folders["tasks"] / "todo_lists.json",
                    task_lists
                )

                all_tasks = []

                for task_list in task_lists:
                    list_id = task_list.get("id")

                    if not list_id:
                        continue

                    task_items_result = self.graph_service.get_todo_tasks(
                        mailbox_email=mailbox_email,
                        list_id=list_id,
                        all_pages=True
                    )

                    if not task_items_result["success"]:
                        result["errors"].append(
                            f"Erro ao consultar tarefas da lista {list_id}: "
                            f"{task_items_result['data']}"
                        )

                        continue

                    for task in task_items_result["items"]:
                        task["_listId"] = list_id
                        task["_listName"] = task_list.get("displayName")
                        all_tasks.append(task)

                result["task_items"] = len(all_tasks)

                self.save_json(
                    backup_folders["tasks"] / "todo_tasks.json",
                    all_tasks
                )
        else:
            self.logger.info("Exportação de tarefas ignorada por parâmetro.")

        return True

    def export_mailbox_local(
        self,
        mailbox_email,
        message_limit=25,
        export_all_messages=False
    ):
        self.logger.info(
            f"Iniciando exportação local simples da mailbox: {mailbox_email}"
        )

        backup_folders = self.create_backup_structure(
            mailbox_email
        )

        result = {
            "success": False,
            "phase": 2,
            "mailbox": mailbox_email,
            "backup_path": str(backup_folders["root"]),
            "folders_count": 0,
            "messages_indexed": 0,
            "messages_exported": 0,
            "messages_failed": 0,
            "calendar_events": 0,
            "contacts": 0,
            "task_lists": 0,
            "task_items": 0,
            "errors": []
        }

        common_ok = self.export_common_data(
            mailbox_email=mailbox_email,
            backup_folders=backup_folders,
            result=result
        )

        if not common_ok:
            return result

        folders_result = self.graph_service.get_mail_folders(
            mailbox_email=mailbox_email,
            all_pages=True
        )

        if not folders_result["success"]:
            result["errors"].append(
                f"Erro ao consultar pastas: {folders_result['data']}"
            )
        else:
            folders = folders_result["items"]

            result["folders_count"] = len(folders)

            self.save_json(
                backup_folders["metadata"] / "folders.json",
                folders
            )

        max_items = None

        if not export_all_messages:
            max_items = message_limit

        messages_result = self.graph_service.get_messages(
            mailbox_email=mailbox_email,
            top=50,
            all_pages=True,
            max_items=max_items
        )

        if not messages_result["success"]:
            result["errors"].append(
                f"Erro ao consultar mensagens: {messages_result['data']}"
            )
        else:
            messages = messages_result["items"]

            result["messages_indexed"] = len(messages)

            self.save_json(
                backup_folders["mail"] / "messages_index.json",
                messages
            )

            for index, message in enumerate(messages, start=1):
                message_id = message.get("id")

                if not message_id:
                    continue

                mime_result = self.graph_service.get_message_mime_content(
                    mailbox_email=mailbox_email,
                    message_id=message_id
                )

                if not mime_result["success"]:
                    result["errors"].append(
                        f"Erro ao baixar mensagem {message_id}: "
                        f"{mime_result['error']}"
                    )
                    result["messages_failed"] += 1

                    continue

                filename = self.build_message_filename(
                    index,
                    message
                )

                eml_path = backup_folders["eml"] / filename

                self.save_bytes(
                    eml_path,
                    mime_result["content"]
                )

                result["messages_exported"] += 1

                self.logger.info(
                    f"Mensagem exportada: {eml_path}"
                )

        self.write_manifest(
            backup_folders=backup_folders,
            result=result
        )

        result["success"] = len(result["errors"]) == 0

        self.logger.info(
            f"Exportação local simples finalizada: {backup_folders['root']}"
        )

        return result

    def export_mailbox_by_folder(
        self,
        mailbox_email,
        message_limit_per_folder=10,
        export_all_messages=False
    ):
        return self.export_mailbox_complete(
            mailbox_email=mailbox_email,
            message_limit_per_folder=message_limit_per_folder,
            export_all_messages=export_all_messages,
            export_attachments=False,
            skip_calendar=False,
            skip_contacts=False,
            skip_tasks=False,
            resume_path=None,
            phase=3,
            excluded_folder_names=None
        )

    def download_eml_batch(self, mailbox_email, prepared_messages):
        """Download MIME files through a continuously refilled bounded window.

        The previous implementation waited for every item in a fixed batch before
        submitting the next batch. One large message could therefore leave several
        workers idle. This window submits a replacement as soon as any future ends.
        """
        results = []
        if not prepared_messages:
            return results
        configured_limit = (
            MIME_RESUME_MAX_CONCURRENCY
            if os.getenv("M365_OPERATION_ID")
            else MIME_MAX_CONCURRENCY
        )
        pending_limit = max(
            1,
            min(
                int(configured_limit or 1),
                int(EML_MAX_PENDING_DOWNLOADS or 1),
                int(EML_DOWNLOAD_WORKERS or 1),
            ),
        )
        iterator = iter(prepared_messages)
        futures = {}

        def submit_next():
            try:
                item = next(iterator)
            except StopIteration:
                return False
            future = self.eml_executor.submit(
                self.graph_service.download_message_mime_to_file,
                mailbox_email,
                item["message_id"],
                item["eml_path"],
            )
            futures[future] = item
            return True

        for _ in range(pending_limit):
            if not submit_next():
                break

        while futures:
            self.operation_control.checkpoint()
            completed, _ = wait(
                tuple(futures), timeout=0.25, return_when=FIRST_COMPLETED
            )
            if not completed:
                continue
            for future in completed:
                item = futures.pop(future)
                try:
                    download_result = future.result()
                except Exception as error:
                    download_result = {
                        "success": False,
                        "error": str(error),
                        "bytes_written": 0,
                        "path": None,
                    }
                results.append((item, download_result))
                submit_next()
        return results

    def existing_eml_index(self, eml_folder):
        """Index completed EML files in one folder without scanning the whole backup."""
        index_by_message_suffix = {}
        highest_sequence = 0
        eml_folder = Path(eml_folder)
        if not eml_folder.exists():
            return index_by_message_suffix, highest_sequence

        for path in eml_folder.glob("*.eml"):
            try:
                if not path.is_file() or path.stat().st_size <= 0:
                    continue
            except OSError:
                continue

            sequence_match = re.match(r"^(\d+)_", path.name)
            if sequence_match:
                highest_sequence = max(highest_sequence, int(sequence_match.group(1)))

            suffix_match = re.search(r"_([^_]+)\.eml$", path.name, re.IGNORECASE)
            if suffix_match:
                index_by_message_suffix.setdefault(suffix_match.group(1), path)

        return index_by_message_suffix, highest_sequence

    def message_file_suffix(self, message_id):
        return self.sanitize_name(str(message_id or "")[-12:], max_length=20)

    def build_backup_scope(
        self,
        mailbox_email,
        export_all_messages,
        message_limit_per_folder,
        export_attachments,
        excluded_folder_names
    ):
        scope = {
            "mailbox": str(mailbox_email).strip().lower(),
            "export_all_messages": bool(export_all_messages),
            "message_limit_per_folder": (
                None if export_all_messages else int(message_limit_per_folder or 0)
            ),
            "export_attachments": bool(export_attachments),
            "excluded_folder_names": sorted(
                str(item).strip().lower()
                for item in (excluded_folder_names or [])
                if str(item).strip()
            )
        }
        serialized = json.dumps(
            scope,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":")
        )
        return scope, hashlib.sha256(serialized.encode("utf-8")).hexdigest()

    def calculate_expected_message_count(
        self,
        folders,
        export_all_messages,
        message_limit_per_folder,
        excluded_folder_names
    ):
        total = 0
        for folder in folders:
            folder_path = folder.get("path") or folder.get("displayName")
            if self.should_skip_folder(folder_path, excluded_folder_names):
                continue
            item_count = max(0, int(folder.get("totalItemCount", 0) or 0))
            if export_all_messages:
                total += item_count
            else:
                total += min(item_count, max(0, int(message_limit_per_folder or 0)))
        return total

    def update_checkpoint_progress(
        self,
        checkpoint,
        exported_message_ids,
        failed_message_ids,
        expected_message_count=None,
        downloaded_bytes=None,
        discovered_message_bytes=None,
        expected_bytes=None
    ):
        exported_count = len(exported_message_ids)
        checkpoint["exported_message_count"] = exported_count
        checkpoint["failed_message_count"] = len(failed_message_ids)
        if expected_message_count is not None:
            requested_expected = int(expected_message_count or 0)
            checkpoint["progress_consistent"] = exported_count <= requested_expected
            checkpoint["expected_message_count"] = max(
                exported_count,
                requested_expected
            )
        if downloaded_bytes is not None:
            checkpoint["downloaded_bytes"] = max(0, int(downloaded_bytes or 0))
        if discovered_message_bytes is not None:
            checkpoint["discovered_message_bytes"] = max(
                0, int(discovered_message_bytes or 0)
            )
        if expected_bytes is not None:
            checkpoint["expected_bytes"] = max(
                int(checkpoint.get("downloaded_bytes", 0) or 0),
                int(expected_bytes or 0)
            )
        checkpoint["progress_updated_at"] = datetime.now().isoformat()

    def log_structured_progress(self, checkpoint, force=False):
        current = int(checkpoint.get("exported_message_count", 0) or 0)
        expected = int(checkpoint.get("expected_message_count", 0) or 0)
        now = time.monotonic()
        if not force:
            if current == self._last_progress_log_current:
                return
            if current != expected and now - self._last_progress_log_at < 0.75:
                return
        self._last_progress_log_at = now
        self._last_progress_log_current = current
        payload = {
            "mailbox": checkpoint.get("mailbox"),
            "scope_hash": checkpoint.get("scope_hash"),
            "expected": int(checkpoint.get("expected_message_count", 0) or 0),
            "current": int(checkpoint.get("exported_message_count", 0) or 0),
            "failed": int(checkpoint.get("failed_message_count", 0) or 0),
            "downloaded_bytes": int(checkpoint.get("downloaded_bytes", 0) or 0),
            "expected_bytes": int(checkpoint.get("expected_bytes", 0) or 0),
            "bytes_baseline_complete": bool(
                checkpoint.get("bytes_baseline_complete", False)
            ),
            "progress_consistent": bool(
                checkpoint.get("progress_consistent", True)
            ),
            "updated_at": checkpoint.get("progress_updated_at")
        }
        limiter_snapshot = getattr(
            self.graph_service, "rate_limiter_snapshot", lambda: {}
        )()
        if limiter_snapshot:
            payload["rate_limiter"] = limiter_snapshot
        self.logger.info(
            "[PROGRESS] " + json.dumps(
                payload,
                ensure_ascii=False,
                separators=(",", ":")
            )
        )

    def export_mailbox_complete(
        self,
        mailbox_email,
        message_limit_per_folder=10,
        export_all_messages=False,
        export_attachments=False,
        skip_calendar=False,
        skip_contacts=False,
        skip_tasks=False,
        resume_path=None,
        phase=4,
        excluded_folder_names=None,
        selected_folder_ids=None,
        selected_folder_paths=None
    ):
        self.logger.info(
            f"Iniciando exportação completa com checkpoint: {mailbox_email}"
        )

        started_at = datetime.now()

        if resume_path:
            backup_folders = self.create_backup_structure_from_root(
                Path(resume_path)
            )

            self.logger.info(
                f"Retomando backup existente em: {backup_folders['root']}"
            )
        else:
            backup_folders = self.create_backup_structure(
                mailbox_email
            )

        checkpoint = self.load_or_create_checkpoint(
            backup_folders=backup_folders,
            mailbox_email=mailbox_email
        )

        checkpoint["status"] = "running"
        self.save_checkpoint(
            backup_folders,
            checkpoint
        )

        checkpoint_store = CheckpointStore(
            path=backup_folders["root"] / "checkpoint.sqlite3",
            operation_type="mailbox_backup",
            source_root=backup_folders["root"],
            destination_path=backup_folders["root"]
        )
        checkpoint_store.migrate_legacy(
            completed_keys=checkpoint.get("exported_message_ids", []),
            failed_keys=checkpoint.get("failed_message_ids", [])
        )
        checkpoint_store.set_operation(status="running")

        exported_message_ids = checkpoint_store.status_keys("completed")
        failed_message_ids = checkpoint_store.status_keys("failed")

        processed_folder_ids = set()
        checkpoint_changes_pending = 0
        enumerated_message_ids = set()
        fully_enumerated_folder_ids = set()

        result = {
            "success": False,
            "phase": phase,
            "mailbox": mailbox_email,
            "backup_path": str(backup_folders["root"]),
            "started_at": started_at.isoformat(),
            "finished_at": None,
            "duration_seconds": 0,
            "folders_count": 0,
            "folders_processed": 0,
            "folders_skipped": 0,
            "messages_indexed": 0,
            "messages_exported": 0,
            "messages_skipped": 0,
            "messages_failed": 0,
            "attachments_exported": 0,
            "calendar_events": 0,
            "contacts": 0,
            "task_lists": 0,
            "task_items": 0,
            "errors": []
        }

        cached_tree_path = backup_folders["metadata"] / "folder_tree.json"
        cached_folders_path = backup_folders["metadata"] / "folders.json"
        resume_metadata_available = bool(
            resume_path
            and cached_tree_path.exists()
            and cached_folders_path.exists()
        )
        if resume_metadata_available:
            common_ok = True
            self.logger.info(
                "Retomada rápida: metadados já salvos serão reutilizados sem "
                "novas consultas iniciais de perfil, calendário, contatos ou tarefas."
            )
        else:
            common_ok = self.export_common_data(
                mailbox_email=mailbox_email,
                backup_folders=backup_folders,
                result=result,
                skip_calendar=skip_calendar,
                skip_contacts=skip_contacts,
                skip_tasks=skip_tasks
            )

        if not common_ok:
            checkpoint["status"] = "failed"
            checkpoint["errors"].extend(result["errors"])

            self.save_checkpoint(
                backup_folders,
                checkpoint
            )

            return result

        if resume_metadata_available:
            try:
                folder_tree = self.load_json(cached_tree_path)
                flat_folders = self.load_json(cached_folders_path)
                if not isinstance(folder_tree, list) or not isinstance(flat_folders, list):
                    raise ValueError("Metadados de pastas em formato inválido.")
                self.logger.info(
                    f"Retomada rápida: {len(flat_folders)} pasta(s) carregada(s) "
                    "do cache local, sem reconstruir a árvore pelo Graph."
                )
            except Exception as error:
                self.logger.warning(
                    f"Cache de pastas indisponível; consultando o Graph: {error}"
                )
                folder_tree_result = self.graph_service.build_mail_folder_tree(
                    mailbox_email=mailbox_email
                )
                if not folder_tree_result["success"]:
                    result["errors"].extend(folder_tree_result.get("errors", []))
                folder_tree = folder_tree_result.get("tree", [])
                flat_folders = folder_tree_result.get("flat", [])
        else:
            folder_tree_result = self.graph_service.build_mail_folder_tree(
                mailbox_email=mailbox_email
            )
            if not folder_tree_result["success"]:
                result["errors"].extend(folder_tree_result.get("errors", []))
            folder_tree = folder_tree_result.get("tree", [])
            flat_folders = folder_tree_result.get("flat", [])
        selected_folder_ids = {
            str(value) for value in (selected_folder_ids or []) if value
        }
        selected_folder_paths_normalized = {
            str(value).strip().lower()
            for value in (selected_folder_paths or []) if value
        }
        if selected_folder_ids:
            flat_folders = [
                folder for folder in flat_folders
                if str(folder.get("id")) in selected_folder_ids
            ]
            self.logger.info(
                f"Escopo personalizado: {len(flat_folders)} pasta(s) selecionada(s)."
            )
        elif selected_folder_paths_normalized:
            # Usado quando um escopo de pastas é aplicado a outras mailboxes de um
            # lote: os IDs de pasta são específicos de cada mailbox, então o escopo
            # viaja como caminhos e precisa ser resolvido aqui.
            flat_folders = [
                folder for folder in flat_folders
                if str(folder.get("path") or "").strip().lower()
                in selected_folder_paths_normalized
            ]
            self.logger.info(
                f"Escopo personalizado (por caminho): {len(flat_folders)} pasta(s) selecionada(s)."
            )

        result["folders_count"] = len(flat_folders)

        backup_scope, scope_hash = self.build_backup_scope(
            mailbox_email=mailbox_email,
            export_all_messages=export_all_messages,
            message_limit_per_folder=message_limit_per_folder,
            export_attachments=export_attachments,
            excluded_folder_names=excluded_folder_names
        )
        expected_message_count = self.calculate_expected_message_count(
            folders=flat_folders,
            export_all_messages=export_all_messages,
            message_limit_per_folder=message_limit_per_folder,
            excluded_folder_names=excluded_folder_names
        )
        backup_scope["selected_folder_ids"] = sorted(selected_folder_ids)
        backup_scope["selected_folder_paths"] = list(selected_folder_paths or [])
        if selected_folder_ids:
            scope_hash = hashlib.sha256(
                json.dumps(backup_scope, sort_keys=True, ensure_ascii=False).encode("utf-8")
            ).hexdigest()
        checkpoint["backup_scope"] = backup_scope
        checkpoint["scope_hash"] = scope_hash
        bytes_baseline_complete = bool(
            checkpoint.get(
                "bytes_baseline_complete",
                len(exported_message_ids) == 0
            )
        )
        checkpoint["bytes_baseline_complete"] = bytes_baseline_complete
        downloaded_bytes = int(checkpoint.get("downloaded_bytes", 0) or 0)
        discovered_message_bytes = 0
        expected_bytes = int(checkpoint.get("expected_bytes", 0) or 0)
        self.update_checkpoint_progress(
            checkpoint=checkpoint,
            exported_message_ids=exported_message_ids,
            failed_message_ids=failed_message_ids,
            expected_message_count=expected_message_count,
            downloaded_bytes=downloaded_bytes,
            discovered_message_bytes=discovered_message_bytes,
            expected_bytes=expected_bytes
        )
        self.save_checkpoint(backup_folders, checkpoint)
        self.log_structured_progress(checkpoint)

        self.save_json(
            backup_folders["metadata"] / "folder_tree.json",
            folder_tree
        )

        self.save_json(
            backup_folders["metadata"] / "folders.json",
            flat_folders
        )

        for folder in flat_folders:
            self.operation_control.checkpoint()
            folder_id = folder.get("id")
            folder_path = folder.get("path") or folder.get("displayName")

            if not folder_id:
                continue

            if self.should_skip_folder(
                folder_path,
                excluded_folder_names=excluded_folder_names
            ):
                self.logger.info(
                    f"Pasta ignorada por filtro: {folder_path}"
                )

                result["folders_skipped"] += 1
                continue

            self.logger.info(
                f"Exportando pasta: {folder_path}"
            )

            local_folder = self.build_folder_local_path(
                backup_folders["mail_folders"],
                folder_path
            )

            eml_folder = local_folder / "eml"
            attachments_folder = local_folder / "attachments"

            try:
                local_folder.mkdir(parents=True, exist_ok=True)
                eml_folder.mkdir(parents=True, exist_ok=True)
                if export_attachments:
                    attachments_folder.mkdir(parents=True, exist_ok=True)
            except OSError as error:
                # Nomes de pasta reservados no Windows (CON, PRN, COM1...) ou
                # caminhos que excedem o MAX_PATH não podem virar diretório;
                # isola a falha nesta pasta em vez de abortar a mailbox inteira.
                error_message = (
                    f"Não foi possível criar a pasta local para '{folder_path}': "
                    f"{error}"
                )
                self.logger.error(error_message)
                result["errors"].append(error_message)
                result["folders_skipped"] += 1
                checkpoint_store.mark_folder_sync_error(folder_id, str(error))
                continue

            max_items = None

            if not export_all_messages:
                max_items = message_limit_per_folder

            result["folders_processed"] += 1

            existing_folder_index = self.load_json(
                local_folder / "messages_index.json"
            ) if (local_folder / "messages_index.json").exists() else {}
            existing_messages = existing_folder_index.get("messages", [])
            if not isinstance(existing_messages, list):
                existing_messages = []
            folder_messages_by_id = {
                message.get("id"): message
                for message in existing_messages
                if isinstance(message, dict) and message.get("id")
            }
            folder_messages_index = list(existing_messages)
            attachments_index = []

            folder_exported_count = 0
            existing_eml_by_suffix, folder_index_counter = self.existing_eml_index(
                eml_folder
            )
            if existing_eml_by_suffix:
                self.logger.info(
                    f"Retomada: {len(existing_eml_by_suffix)} EML existente(s) "
                    f"indexado(s) em {folder_path}; sequência continuará após "
                    f"{folder_index_counter}."
                )
            stop_folder = False

            # Segurança: backups completos não usam deltaLink. Cada execução
            # enumera novamente todas as páginas da pasta e o SQLite impede o
            # download duplicado dos EML já confirmados.
            checkpoint_store.reset_folder_sync_state(folder_id)
            saved_page_number = 0
            saved_discovered_items = 0
            self.logger.info(
                f"Enumeração integral da pasta {folder_path}: "
                "delta desativado; EML já confirmados serão ignorados."
            )
            checkpoint_store.save_folder_sync_progress(
                folder_id=folder_id,
                folder_path=folder_path,
                scope_hash=scope_hash,
                sync_mode="full",
                status="running",
                next_link=None,
                delta_link=None,
                page_number=0,
                discovered_items=0
            )

            for page_result in self.graph_service.iter_messages_by_folder_pages(
                mailbox_email=mailbox_email,
                folder_id=folder_id,
                top=250,
                max_items=max_items
            ):
                self.operation_control.checkpoint()
                if not page_result["success"]:
                    error_message = (
                        f"Erro ao consultar mensagens da pasta {folder_path}: "
                        f"{page_result['error']}"
                    )

                    result["errors"].append(error_message)
                    checkpoint["errors"].append(error_message)
                    if page_result.get("reset_required"):
                        checkpoint_store.reset_folder_sync_state(folder_id)
                        self.logger.warning(
                            f"O ponto incremental da pasta {folder_path} expirou. "
                            "Somente esta pasta será sincronizada novamente na próxima tentativa."
                        )
                    else:
                        checkpoint_store.mark_folder_sync_error(
                            folder_id,
                            page_result.get("error")
                        )

                    self.save_checkpoint(
                        backup_folders,
                        checkpoint
                    )

                    stop_folder = True
                    break

                messages = page_result["items"]
                confirmed_page_number = int(page_result.get("page", 0) or 0)
                confirmed_discovered = saved_discovered_items + len(messages)
                page_next_link = page_result.get("next_link")
                page_terminal = bool(page_result.get("terminal", not page_next_link))
                enumerated_message_ids.update(
                    message.get("id") for message in messages if message.get("id")
                )

                if not messages:
                    checkpoint_store.save_folder_sync_progress(
                        folder_id=folder_id,
                        folder_path=folder_path,
                        scope_hash=scope_hash,
                        sync_mode="full",
                        status="complete" if page_terminal else "running",
                        next_link=page_next_link,
                        delta_link=None,
                        page_number=confirmed_page_number,
                        discovered_items=confirmed_discovered
                    )
                    continue

                result["messages_indexed"] += len(messages)
                for message in messages:
                    if message.get("id"):
                        folder_messages_by_id[message["id"]] = message
                folder_messages_index = list(folder_messages_by_id.values())
                exported_count_for_average = max(1, len(exported_message_ids))
                average_downloaded_size = downloaded_bytes / exported_count_for_average
                estimated_from_average = int(
                    average_downloaded_size * max(
                        expected_message_count,
                        len(exported_message_ids)
                    )
                )
                expected_bytes = max(
                    downloaded_bytes,
                    discovered_message_bytes,
                    estimated_from_average
                )

                self.logger.info(
                    f"Página {page_result['page']} carregada na pasta {folder_path}. "
                    f"Mensagens na página: {len(messages)}"
                )

                pending_messages = [
                    message for message in messages
                    if message.get("id")
                    and message.get("id") not in exported_message_ids
                ]
                result["messages_skipped"] += len(messages) - len(pending_messages)

                # The checkpoint can be behind the filesystem after an abrupt stop.
                # Reconcile each pending Graph ID with an already completed local EML
                # before scheduling any download. Existing files are adopted into the
                # SQLite/JSON checkpoint and never downloaded again.
                truly_pending_messages = []
                adopted_existing = 0
                for message in pending_messages:
                    message_id = message["id"]
                    suffix = self.message_file_suffix(message_id)
                    existing_path = existing_eml_by_suffix.get(suffix)
                    if existing_path is None:
                        truly_pending_messages.append(message)
                        continue
                    try:
                        existing_size = int(existing_path.stat().st_size)
                    except OSError:
                        truly_pending_messages.append(message)
                        continue
                    if existing_size <= 0:
                        truly_pending_messages.append(message)
                        continue

                    exported_message_ids.add(message_id)
                    failed_message_ids.discard(message_id)
                    checkpoint_store.mark_completed(
                        message_id,
                        folder_id=folder_id,
                        output_path=str(existing_path),
                        bytes_written=existing_size
                    )
                    downloaded_bytes += existing_size
                    adopted_existing += 1
                    result["messages_skipped"] += 1
                    self.logger.debug(
                        f"[SKIP-EXISTING] EML já presente; download ignorado: "
                        f"{existing_path.name}"
                    )

                pending_messages = truly_pending_messages
                if adopted_existing:
                    self.update_checkpoint_progress(
                        checkpoint=checkpoint,
                        exported_message_ids=exported_message_ids,
                        failed_message_ids=failed_message_ids,
                        expected_message_count=expected_message_count,
                        downloaded_bytes=downloaded_bytes,
                        discovered_message_bytes=discovered_message_bytes,
                        expected_bytes=max(expected_bytes, downloaded_bytes)
                    )
                    self.flush_checkpoint(
                        backup_folders,
                        checkpoint,
                        exported_message_ids,
                        failed_message_ids,
                        processed_folder_ids
                    )
                    self.log_structured_progress(checkpoint)
                    self.logger.info(
                        f"Retomada da página {confirmed_page_number}: "
                        f"{len(messages) - len(pending_messages)} já confirmados ou existentes; "
                        f"{adopted_existing} incorporados ao checkpoint; "
                        f"{len(pending_messages)} pendentes para download."
                    )

                if not pending_messages:
                    checkpoint_store.save_folder_sync_progress(
                        folder_id=folder_id,
                        folder_path=folder_path,
                        scope_hash=scope_hash,
                        sync_mode="full",
                        status="complete" if page_terminal else "running",
                        next_link=page_next_link,
                        delta_link=None,
                        page_number=confirmed_page_number,
                        discovered_items=confirmed_discovered
                    )
                    continue

                prepared_messages = []
                for message in pending_messages:
                    folder_index_counter += 1
                    message_id = message["id"]
                    filename = self.build_message_filename(
                        folder_index_counter,
                        message
                    )
                    eml_path = eml_folder / filename
                    checkpoint_store.mark_processing(
                        message_id,
                        folder_id=folder_id,
                        output_path=str(eml_path)
                    )
                    prepared_messages.append(
                        {
                            "message": message,
                            "message_id": message_id,
                            "message_index": folder_index_counter,
                            "eml_path": eml_path
                        }
                    )

                for item, mime_result in self.download_eml_batch(
                    mailbox_email,
                    prepared_messages
                ):
                    message = item["message"]
                    message_id = item["message_id"]
                    eml_path = item["eml_path"]

                    if not mime_result["success"]:
                        diagnostic = mime_result.get("diagnostic") or {}
                        error_message = (
                            f"Erro ao baixar mensagem da pasta {folder_path}: "
                            f"{message_id} | estágio={diagnostic.get('stage', 'desconhecido')} | "
                            f"tipo={diagnostic.get('exception_type', 'desconhecido')} | "
                            f"errno={diagnostic.get('errno')} | winerror={diagnostic.get('winerror')} | "
                            f"temporário={diagnostic.get('temporary_length', '?')} caracteres | "
                            f"{mime_result.get('error')}"
                        )
                        self.logger.error(
                            "[ROOT-CAUSE] %s",
                            json.dumps({
                                "mailbox": mailbox_email, "folder_id": folder_id,
                                "folder_path": folder_path, "message_id": message_id,
                                "subject": message.get("subject"),
                                "download": mime_result,
                            }, ensure_ascii=False)
                        )
                        result["errors"].append(error_message)
                        checkpoint["errors"].append(error_message)
                        failed_message_ids.add(message_id)
                        checkpoint_store.mark_failed(
                            message_id,
                            mime_result.get("error"),
                            folder_id=folder_id,
                            output_path=str(eml_path)
                        )
                        result["messages_failed"] += 1
                        checkpoint_changes_pending += 1
                    else:
                        actual_eml_path = Path(mime_result.get("path") or eml_path)
                        folder_exported_count += 1
                        result["messages_exported"] += 1
                        downloaded_bytes += int(
                            mime_result.get("bytes_written", 0) or 0
                        )
                        self.logger.info(
                            f"[OK] EML salvo {folder_exported_count}/{max_items or '?'} "
                            f"na pasta {folder_path}: {actual_eml_path.name} | "
                            f"{mime_result.get('bytes_written', 0)} bytes"
                        )

                        if export_attachments and message.get("hasAttachments"):
                            attachment_count = self.export_message_attachments(
                                mailbox_email=mailbox_email,
                                message=message,
                                message_index=item["message_index"],
                                attachments_folder=attachments_folder,
                                attachments_index=attachments_index
                            )
                            result["attachments_exported"] += attachment_count

                        exported_message_ids.add(message_id)
                        failed_message_ids.discard(message_id)
                        checkpoint_store.mark_completed(
                            message_id,
                            folder_id=folder_id,
                            output_path=str(actual_eml_path),
                            bytes_written=int(
                                mime_result.get("bytes_written", 0) or 0
                            )
                        )
                        existing_eml_by_suffix[self.message_file_suffix(message_id)] = actual_eml_path
                        checkpoint_changes_pending += 1

                    self.update_checkpoint_progress(
                        checkpoint=checkpoint,
                        exported_message_ids=exported_message_ids,
                        failed_message_ids=failed_message_ids,
                        expected_message_count=expected_message_count,
                        downloaded_bytes=downloaded_bytes,
                        discovered_message_bytes=discovered_message_bytes,
                        expected_bytes=max(expected_bytes, downloaded_bytes)
                    )
                    self.log_structured_progress(checkpoint)

                    # SQLite is durable per EML. The larger JSON snapshot is
                    # consolidated to reduce disk contention while preserving resume.
                    now_checkpoint = time.monotonic()
                    if (
                        checkpoint_changes_pending >= 10
                        or now_checkpoint - self._last_checkpoint_json_at >= 2.0
                    ):
                        self.flush_checkpoint(
                            backup_folders,
                            checkpoint,
                            exported_message_ids,
                            failed_message_ids,
                            processed_folder_ids
                        )
                        self._last_checkpoint_json_at = now_checkpoint
                        checkpoint_changes_pending = 0

                # O link só é confirmado depois que toda a página foi processada.
                # Assim, uma interrupção nunca pula mensagens ainda não persistidas.
                checkpoint_store.save_folder_sync_progress(
                    folder_id=folder_id,
                    folder_path=folder_path,
                    scope_hash=scope_hash,
                    sync_mode="full",
                    status="complete" if page_terminal else "running",
                    next_link=page_next_link,
                    delta_link=None,
                    page_number=confirmed_page_number,
                    discovered_items=confirmed_discovered
                )

                if stop_folder:
                    break

            folder_index = {
                "folder": folder,
                "messages_count": len(folder_messages_index),
                "messages": folder_messages_index
            }

            self.save_json(
                local_folder / "messages_index.json",
                folder_index
            )

            if export_attachments:
                self.save_json(
                    local_folder / "attachments_index.json",
                    attachments_index
                )

            if not stop_folder:
                processed_folder_ids.add(folder_id)
                fully_enumerated_folder_ids.add(str(folder_id))
                checkpoint_changes_pending += 1
                self.update_checkpoint_progress(
                    checkpoint=checkpoint,
                    exported_message_ids=exported_message_ids,
                    failed_message_ids=failed_message_ids,
                    expected_message_count=expected_message_count,
                    downloaded_bytes=downloaded_bytes,
                    discovered_message_bytes=discovered_message_bytes,
                    expected_bytes=max(expected_bytes, downloaded_bytes)
                )
                self.flush_checkpoint(
                    backup_folders,
                    checkpoint,
                    exported_message_ids,
                    failed_message_ids,
                    processed_folder_ids
                )
                checkpoint_changes_pending = 0

        sqlite_counts = checkpoint_store.counts()
        exported_message_ids = checkpoint_store.status_keys("completed")
        failed_message_ids = checkpoint_store.status_keys("failed")
        downloaded_bytes = sqlite_counts["bytes_written"]
        self.sync_checkpoint_sets(
            checkpoint,
            exported_message_ids,
            failed_message_ids,
            processed_folder_ids
        )
        self.update_checkpoint_progress(
            checkpoint=checkpoint,
            exported_message_ids=exported_message_ids,
            failed_message_ids=failed_message_ids,
            expected_message_count=expected_message_count,
            downloaded_bytes=downloaded_bytes,
            discovered_message_bytes=discovered_message_bytes,
            expected_bytes=max(expected_bytes, downloaded_bytes)
        )
        self.log_structured_progress(checkpoint)

        # Deduplica e limita o histórico acumulado entre retomadas para que o
        # checkpoint.json não cresça sem limite em backups retomados muitas vezes.
        combined_errors = checkpoint.get("errors", []) + result["errors"]
        seen_errors = set()
        deduplicated_errors = []
        for error_message in combined_errors:
            if error_message not in seen_errors:
                seen_errors.add(error_message)
                deduplicated_errors.append(error_message)
        checkpoint["errors"] = deduplicated_errors[-500:]

        finished_at = datetime.now()

        result["finished_at"] = finished_at.isoformat()
        result["duration_seconds"] = round(
            (finished_at - started_at).total_seconds(),
            2
        )

        expected_folder_ids = {
            str(folder.get("id")) for folder in flat_folders
            if folder.get("id")
            and not self.should_skip_folder(
                folder.get("path") or folder.get("displayName"),
                excluded_folder_names=excluded_folder_names
            )
        }
        sqlite_counts = checkpoint_store.counts()
        all_folders_enumerated = expected_folder_ids.issubset(
            fully_enumerated_folder_ids
        )
        all_enumerated_saved = enumerated_message_ids.issubset(
            exported_message_ids
        )
        reported_expected = max(0, int(expected_message_count or 0))
        enumerated_count = len(enumerated_message_ids)
        completed_count = len(exported_message_ids)
        strict_expected = max(reported_expected, enumerated_count, completed_count)
        enumeration_matches_expected = (
            reported_expected == 0 or enumerated_count >= reported_expected
        )
        checkpoint["reported_expected_message_count"] = reported_expected
        checkpoint["enumerated_message_count"] = enumerated_count
        checkpoint["expected_message_count"] = strict_expected
        checkpoint["exported_message_count"] = completed_count
        checkpoint["enumeration_matches_expected"] = enumeration_matches_expected
        checkpoint["full_scan_verified"] = bool(
            all_folders_enumerated
            and enumeration_matches_expected
            and all_enumerated_saved
            and completed_count >= enumerated_count
            and int(sqlite_counts.get("failed", 0) or 0) == 0
            and int(sqlite_counts.get("processing", 0) or 0) == 0
            and not result["errors"]
        )
        result["success"] = checkpoint["full_scan_verified"]
        result["expected_message_count"] = strict_expected
        result["reported_expected_message_count"] = expected_message_count
        result["enumerated_message_count"] = len(enumerated_message_ids)

        self.log_structured_progress(checkpoint, force=True)
        if result["success"]:
            checkpoint["status"] = "completed"
            self.logger.info(
                "Backup integral validado: "
                f"{len(exported_message_ids)}/{strict_expected} EML confirmados; "
                f"{len(fully_enumerated_folder_ids)}/{len(expected_folder_ids)} "
                "pastas enumeradas."
            )
        else:
            checkpoint["status"] = "incomplete"
            result["errors"].append(
                "A enumeração integral não foi comprovada; a operação permanece "
                "retomável e não será marcada como concluída."
            )
            self.logger.warning(
                "Backup integral não validado: "
                f"EML confirmados={len(exported_message_ids)}, "
                f"enumerados={len(enumerated_message_ids)}, "
                f"pastas={len(fully_enumerated_folder_ids)}/{len(expected_folder_ids)}, "
                f"falhas={sqlite_counts.get('failed', 0)}, "
                f"processando={sqlite_counts.get('processing', 0)}."
            )

        checkpoint_store.set_operation(
            status=checkpoint["status"],
            total_items=int(checkpoint.get("expected_message_count", 0) or 0)
        )
        checkpoint_store.compact()

        self.save_checkpoint(
            backup_folders,
            checkpoint
        )

        self.write_manifest(
            backup_folders=backup_folders,
            result=result
        )

        self.write_audit_report(
            backup_folders=backup_folders,
            result=result
        )

        if result["success"]:
            completed_marker = backup_folders["root"] / "backup_completed.ok"

            with open(completed_marker, "w", encoding="utf-8") as file:
                file.write(datetime.now().isoformat())

        self.logger.info(
            f"Exportação completa finalizada: {backup_folders['root']}"
        )

        return result

    @staticmethod
    def _extract_failed_ids_from_reports(backup_root):
        pattern = re.compile(r"Erro ao baixar mensagem(?: da pasta [^:]+)?:\s*([^|\s]+)")
        found = {}
        candidates = [
            Path(backup_root) / "audit_report.json",
            Path(backup_root) / "manifest.json",
            Path(backup_root) / "checkpoint.json",
        ]
        reports_root = Path(backup_root).parent / "_batch_reports"
        if reports_root.exists():
            candidates.extend(sorted(reports_root.glob("batch_report_*.json"), reverse=True)[:10])
        for report_path in candidates:
            if not report_path.is_file():
                continue
            try:
                data = json.loads(report_path.read_text(encoding="utf-8"))
            except Exception:
                continue
            serialized = json.dumps(data, ensure_ascii=False)
            for match in pattern.finditer(serialized):
                message_id = match.group(1).replace("\\n", "").strip()
                if message_id:
                    found.setdefault(message_id, {"item_key": message_id, "report_source": str(report_path)})
        return found

    @staticmethod
    def _validate_repaired_eml(path):
        path = Path(path)
        if not path.is_file() or path.suffix.lower() != ".eml":
            return False, "O arquivo EML final não existe."
        size = path.stat().st_size
        if size <= 0:
            return False, "O arquivo EML está vazio."
        with path.open("rb") as handle:
            header = BytesHeaderParser().parse(handle)
        if not any(header.get(key) for key in ("Subject", "From", "Date", "Message-ID", "MIME-Version")):
            return False, "Cabeçalho MIME não reconhecido."
        return True, None

    def repair_failed_messages(self, backup_root, mailbox_email=None):
        """Repara somente mensagens ainda falhas, sem enumerar a mailbox inteira."""
        started = datetime.now()
        backup_root = Path(backup_root).expanduser().resolve()
        checkpoint_path = backup_root / "checkpoint.json"
        if not checkpoint_path.is_file():
            raise FileNotFoundError(f"checkpoint.json não encontrado: {checkpoint_path}")
        checkpoint = self.load_json(checkpoint_path)
        mailbox_email = mailbox_email or checkpoint.get("mailbox")
        if not mailbox_email:
            raise ValueError("Mailbox não identificada no checkpoint.")
        backup_folders = self.create_backup_structure_from_root(backup_root)
        store = CheckpointStore(
            path=backup_root / "checkpoint.sqlite3", operation_type="mailbox_backup",
            source_root=backup_root, destination_path=backup_root
        )
        report_candidates = self._extract_failed_ids_from_reports(backup_root)
        failed_rows = {row["item_key"]: row for row in store.failed_items(retryable_only=False)}
        for message_id, fallback in report_candidates.items():
            if store.is_completed(message_id):
                continue
            failed_rows.setdefault(message_id, fallback)
        total = len(failed_rows)
        result = {
            "success": False, "mode": "repair_failures", "mailbox": mailbox_email,
            "backup_path": str(backup_root), "started_at": started.isoformat(),
            "finished_at": None, "duration_seconds": 0, "found": total,
            "attempted": 0, "recovered": 0, "reconciled": 0,
            "still_failed": 0, "items": [], "errors": []
        }
        self.logger.info("[REPAIR-EVENT] %s", json.dumps({
            "stage":"load_failures", "mailbox":mailbox_email, "total":total
        }, ensure_ascii=False))
        if not total:
            result["success"] = True
        metadata_by_id = {}
        for index, (message_id, row) in enumerate(failed_rows.items(), 1):
            self.operation_control.checkpoint()
            item_started = datetime.now().isoformat()
            previous_error = row.get("error")
            original_output = row.get("output_path")
            folder_id = row.get("folder_id")
            if original_output:
                prior_path = Path(original_output)
                try:
                    valid, _ = self._validate_repaired_eml(prior_path)
                except Exception:
                    valid = False
                if valid:
                    size = prior_path.stat().st_size
                    store.mark_completed(message_id, folder_id=folder_id,
                        output_path=str(prior_path), bytes_written=size,
                        verification_status="reconciled")
                    store.record_repair(message_id, True, previous_error=previous_error,
                        folder_id=folder_id, original_output_path=original_output,
                        final_output_path=str(prior_path), bytes_written=size,
                        diagnostic={"stage":"reconciled_existing"}, started_at=item_started)
                    result["reconciled"] += 1
                    result["items"].append({"message_id":message_id,"success":True,
                        "action":"reconciled","path":str(prior_path)})
                    current = result["recovered"] + result["reconciled"]
                    self.logger.info("[REPAIR-PROGRESS] %s", json.dumps({
                        "mailbox":mailbox_email,"current":current,"total":total,
                        "recovered":result["recovered"],"reconciled":result["reconciled"],
                        "failed":index-current,"stage":"reconciled"
                    }, ensure_ascii=False, separators=(",",":")))
                    continue
            result["attempted"] += 1
            meta_result = self.graph_service.get_message_by_id(mailbox_email, message_id)
            message = meta_result.get("data") if meta_result.get("success") else {"id": message_id}
            message = message if isinstance(message, dict) else {"id": message_id}
            message.setdefault("id", message_id)
            if not folder_id:
                folder_id = message.get("parentFolderId")
            if original_output:
                destination = Path(original_output)
                if destination.suffix.lower() == ".part":
                    destination = destination.with_suffix("")
            else:
                folder_path = "Falhas recuperadas"
                destination_dir = self.build_folder_local_path(
                    backup_folders["mail_folders"], folder_path) / "eml"
                destination = destination_dir / self.build_message_filename(index, message)
            store.mark_processing(message_id, folder_id=folder_id, output_path=str(destination))
            download = self.graph_service.download_message_mime_to_file(
                mailbox_email, message_id, destination
            )
            actual_path = Path(download.get("path") or destination)
            valid = False
            validation_error = None
            if download.get("success"):
                try:
                    valid, validation_error = self._validate_repaired_eml(actual_path)
                except Exception as error:
                    validation_error = str(error)
            if download.get("success") and valid:
                size = actual_path.stat().st_size
                store.mark_completed(message_id, folder_id=folder_id,
                    output_path=str(actual_path), bytes_written=size,
                    verification_status="repaired")
                store.record_repair(message_id, True, previous_error=previous_error,
                    folder_id=folder_id, original_output_path=original_output,
                    final_output_path=str(actual_path), bytes_written=size,
                    diagnostic=download.get("diagnostic"), started_at=item_started)
                result["recovered"] += 1
                result["items"].append({"message_id":message_id,"success":True,
                    "action":"downloaded","path":str(actual_path),"bytes_written":size})
            else:
                error = validation_error or download.get("error") or "Falha não identificada."
                diagnostic = download.get("diagnostic") or {"stage":"validate_final"}
                store.mark_failed(message_id, error, folder_id=folder_id, output_path=str(destination))
                store.record_repair(message_id, False, previous_error=previous_error,
                    new_error=error, folder_id=folder_id,
                    original_output_path=original_output, final_output_path=None,
                    diagnostic=diagnostic, started_at=item_started)
                result["still_failed"] += 1
                result["errors"].append(f"{message_id} | {error}")
                result["items"].append({"message_id":message_id,"success":False,
                    "error":error,"diagnostic":diagnostic})
                self.logger.error("[ROOT-CAUSE] %s", json.dumps({
                    "mode":"repair_failures","mailbox":mailbox_email,
                    "message_id":message_id,"download":download,
                    "validation_error":validation_error
                }, ensure_ascii=False))
            current = result["recovered"] + result["reconciled"] + result["still_failed"]
            self.logger.info("[REPAIR-PROGRESS] %s", json.dumps({
                "mailbox":mailbox_email,"current":current,"total":total,
                "recovered":result["recovered"],"reconciled":result["reconciled"],
                "failed":result["still_failed"],"stage":"repairing"
            }, ensure_ascii=False, separators=(",",":")))
        counts = store.counts()
        checkpoint["exported_message_count"] = int(counts.get("completed", 0) or 0)
        checkpoint["failed_message_count"] = int(counts.get("failed", 0) or 0)
        checkpoint["downloaded_bytes"] = int(counts.get("bytes_written", 0) or 0)
        expected = int(checkpoint.get("expected_message_count", 0) or 0)
        completed = checkpoint["exported_message_count"]
        checkpoint["full_scan_verified"] = bool(
            checkpoint.get("enumeration_matches_expected", False)
            and completed >= expected and checkpoint["failed_message_count"] == 0
            and int(counts.get("processing", 0) or 0) == 0
        )
        checkpoint["status"] = "completed" if checkpoint["full_scan_verified"] else "incomplete"
        checkpoint["last_repair_at"] = datetime.now().isoformat()
        result["success"] = result["still_failed"] == 0
        result["active_failures_after"] = checkpoint["failed_message_count"]
        result["completed_accumulated"] = completed
        finished = datetime.now()
        result["finished_at"] = finished.isoformat()
        result["duration_seconds"] = round((finished-started).total_seconds(), 2)
        reports_dir = backup_root / "repair_reports"
        reports_dir.mkdir(parents=True, exist_ok=True)
        stamp = finished.strftime("%Y%m%d_%H%M%S")
        report_path = reports_dir / f"repair_report_{stamp}.json"
        self.save_json(report_path, result)
        self.save_json(reports_dir / "latest_repair_report.json", result)
        self.save_checkpoint(backup_folders, checkpoint)
        summary = {
            "messages_exported_this_run": 0,
            "messages_repaired_this_run": result["recovered"],
            "messages_reconciled_this_run": result["reconciled"],
            "messages_completed_accumulated": completed,
            "messages_failed_active": checkpoint["failed_message_count"],
        }
        for name in ("manifest.json", "audit_report.json"):
            path = backup_root / name
            if path.is_file():
                try:
                    data = self.load_json(path)
                    data["success"] = checkpoint["full_scan_verified"]
                    data["repair_summary"] = summary
                    data["last_repair_report"] = str(report_path)
                    self.save_json(path, data)
                except Exception as error:
                    self.logger.warning("Não foi possível atualizar %s: %s", name, error)
        if checkpoint["full_scan_verified"]:
            (backup_root / "backup_completed.ok").write_text(finished.isoformat(), encoding="utf-8")
        self.logger.info("[REPAIR-SUMMARY] %s", json.dumps(result, ensure_ascii=False, separators=(",",":")))
        return result

    def export_message_attachments(
        self,
        mailbox_email,
        message,
        message_index,
        attachments_folder,
        attachments_index
    ):
        message_id = message.get("id")

        if not message_id:
            return 0

        attachments_result = self.graph_service.get_message_attachments(
            mailbox_email=mailbox_email,
            message_id=message_id
        )

        if not attachments_result["success"]:
            self.logger.error(
                f"Erro ao consultar anexos da mensagem {message_id}: "
                f"{attachments_result['data']}"
            )

            return 0

        attachments = attachments_result["items"]

        exported_count = 0

        message_attachment_folder = (
            attachments_folder
            / f"message_{message_index:06d}"
        )

        message_attachment_folder.mkdir(
            parents=True,
            exist_ok=True
        )

        for attachment_index, attachment in enumerate(attachments, start=1):
            attachment_metadata = dict(attachment)

            attachment_metadata.pop(
                "contentBytes",
                None
            )

            attachment_metadata["_messageId"] = message_id

            attachments_index.append(
                attachment_metadata
            )

            content_bytes = attachment.get("contentBytes")

            if not content_bytes:
                continue

            filename = self.build_attachment_filename(
                attachment_index,
                attachment
            )

            attachment_path = message_attachment_folder / filename

            try:
                decoded_content = base64.b64decode(
                    content_bytes
                )

                self.save_bytes(
                    attachment_path,
                    decoded_content
                )

                exported_count += 1

                self.logger.info(
                    f"Anexo exportado: {attachment_path}"
                )
            except Exception as error:
                self.logger.error(
                    f"Erro ao salvar anexo {filename}: {error}"
                )

        return exported_count

    def write_manifest(self, backup_folders, result):
        manifest = {
            "project": "M365 Mailbox Backup",
            "phase": result.get("phase"),
            "mailbox": result["mailbox"],
            "created_at": datetime.now().isoformat(),
            "started_at": result.get("started_at"),
            "finished_at": result.get("finished_at"),
            "duration_seconds": result.get("duration_seconds", 0),
            "backup_path": result["backup_path"],
            "success": result.get("success", False),
            "summary": {
                "folders_count": result.get("folders_count", 0),
                "folders_processed": result.get("folders_processed", 0),
                "folders_skipped": result.get("folders_skipped", 0),
                "messages_indexed": result.get("messages_indexed", 0),
                "messages_exported": result.get("messages_exported", 0),
                "messages_skipped": result.get("messages_skipped", 0),
                "messages_failed": result.get("messages_failed", 0),
                "attachments_exported": result.get("attachments_exported", 0),
                "calendar_events": result.get("calendar_events", 0),
                "contacts": result.get("contacts", 0),
                "task_lists": result.get("task_lists", 0),
                "task_items": result.get("task_items", 0),
                "errors_count": len(result.get("errors", []))
            },
            "errors": result.get("errors", [])
        }

        self.save_json(
            backup_folders["root"] / "manifest.json",
            manifest
        )

    def load_batch_file(self, batch_path):
        batch_path = Path(batch_path)

        if not batch_path.exists():
            raise FileNotFoundError(
                f"Arquivo CSV não encontrado: {batch_path}"
            )

        mailboxes = []

        with open(
            batch_path,
            "r",
            encoding="utf-8-sig",
            newline=""
        ) as csv_file:
            reader = csv.DictReader(csv_file)

            if "email" not in reader.fieldnames:
                raise ValueError(
                    "O arquivo CSV precisa conter a coluna obrigatória: email"
                )

            for row in reader:
                email = row.get("email", "").strip()

                if not email:
                    continue

                mailboxes.append(email)

        return mailboxes

    def validate_mailbox(self, mailbox_email):
        self.logger.info(
            f"Validando mailbox antes do backup: {mailbox_email}"
        )

        mailbox_result = self.graph_service.get_mailbox_summary(
            mailbox_email
        )

        if not mailbox_result.get("success"):
            return {
                "mailbox": mailbox_email,
                "valid": False,
                "reason": "Mailbox não encontrada ou sem acesso pelo Graph.",
                "details": mailbox_result.get("data")
            }

        mailbox_data = mailbox_result.get("data", {})

        return {
            "mailbox": mailbox_email,
            "valid": True,
            "reason": "Mailbox validada com sucesso.",
            "details": {
                "id": mailbox_data.get("id"),
                "displayName": mailbox_data.get("displayName"),
                "mail": mailbox_data.get("mail"),
                "userPrincipalName": mailbox_data.get("userPrincipalName"),
                "accountEnabled": mailbox_data.get("accountEnabled")
            }
        }

    def precheck_batch(self, mailboxes, batch_path):
        self.logger.info(
            "Iniciando pré-validação das mailboxes do lote."
        )

        started_at = datetime.now()

        precheck_result = {
            "success": False,
            "phase": 7,
            "batch_file": str(batch_path),
            "started_at": started_at.isoformat(),
            "finished_at": None,
            "duration_seconds": 0,
            "total_mailboxes": len(mailboxes),
            "valid_count": 0,
            "invalid_count": 0,
            "valid_mailboxes": [],
            "invalid_mailboxes": [],
            "results": []
        }

        for mailbox_email in mailboxes:
            validation = self.validate_mailbox(
                mailbox_email
            )

            precheck_result["results"].append(
                validation
            )

            if validation["valid"]:
                precheck_result["valid_count"] += 1
                precheck_result["valid_mailboxes"].append(
                    mailbox_email
                )
            else:
                precheck_result["invalid_count"] += 1
                precheck_result["invalid_mailboxes"].append(
                    mailbox_email
                )

        finished_at = datetime.now()

        precheck_result["finished_at"] = finished_at.isoformat()
        precheck_result["duration_seconds"] = round(
            (finished_at - started_at).total_seconds(),
            2
        )

        precheck_result["success"] = precheck_result["invalid_count"] == 0

        self.save_precheck_report(
            precheck_result
        )

        self.logger.info(
            "Pré-validação finalizada. "
            f"Válidas: {precheck_result['valid_count']} | "
            f"Inválidas: {precheck_result['invalid_count']}"
        )

        return precheck_result

    def save_precheck_report(self, precheck_result):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        reports_dir = OUTPUT_DIR / "_precheck_reports"

        reports_dir.mkdir(
            parents=True,
            exist_ok=True
        )

        json_report_path = reports_dir / f"precheck_report_{timestamp}.json"
        csv_report_path = reports_dir / f"precheck_report_{timestamp}.csv"

        precheck_result["report_path"] = str(json_report_path)
        precheck_result["csv_report_path"] = str(csv_report_path)

        self.save_json(
            json_report_path,
            precheck_result
        )

        with open(
            csv_report_path,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as csv_file:
            fieldnames = [
                "mailbox",
                "valid",
                "reason",
                "displayName",
                "mail",
                "userPrincipalName",
                "accountEnabled"
            ]

            writer = csv.DictWriter(
                csv_file,
                fieldnames=fieldnames
            )

            writer.writeheader()

            for item in precheck_result.get("results", []):
                details = item.get("details") or {}

                if not isinstance(details, dict):
                    details = {}

                writer.writerow(
                    {
                        "mailbox": item.get("mailbox"),
                        "valid": item.get("valid"),
                        "reason": item.get("reason"),
                        "displayName": details.get("displayName"),
                        "mail": details.get("mail"),
                        "userPrincipalName": details.get("userPrincipalName"),
                        "accountEnabled": details.get("accountEnabled")
                    }
                )

    def export_batch(
        self,
        batch_path,
        message_limit_per_folder=10,
        export_all_messages=False,
        export_attachments=False,
        skip_calendar=False,
        skip_contacts=False,
        skip_tasks=False,
        resume_path=None,
        excluded_folder_names=None,
        skip_precheck=False,
        selected_folder_ids=None,
        selected_folder_paths=None
    ):
        
        self.logger.info(
            f"Iniciando backup em lote usando CSV: {batch_path}"
        )

        mailboxes = self.load_batch_file(
            batch_path
        )

        batch_started_at = datetime.now()

        if skip_precheck:
            self.logger.info(
                "Pré-validação ignorada por parâmetro. Iniciando backup diretamente."
            )

            precheck_result = {
                "report_path": None,
                "csv_report_path": None,
                "valid_count": len(mailboxes),
                "invalid_count": 0,
                "valid_mailboxes": mailboxes,
                "invalid_mailboxes": []
            }
        else:
            precheck_result = self.precheck_batch(
                mailboxes=mailboxes,
                batch_path=batch_path
            )

        batch_result = {
            "success": False,
            "phase": 5,
            "batch_file": str(batch_path),
            "started_at": batch_started_at.isoformat(),
            "finished_at": None,
            "duration_seconds": 0,
            "total_mailboxes": len(mailboxes),
            "precheck_report_path": precheck_result.get("report_path"),
            "precheck_csv_report_path": precheck_result.get("csv_report_path"),
            "precheck_valid_count": precheck_result.get("valid_count", 0),
            "precheck_invalid_count": precheck_result.get("invalid_count", 0),
            "processed": 0,
            "success_count": 0,
            "failed_count": 0,
            "results": [],
            "errors": []
        }

        if not mailboxes:
            batch_finished_at = datetime.now()

            batch_result["finished_at"] = batch_finished_at.isoformat()
            batch_result["duration_seconds"] = round(
                (batch_finished_at - batch_started_at).total_seconds(),
                2
            )

            batch_result["errors"].append(
                "Nenhuma mailbox encontrada no arquivo CSV."
            )

            self.save_batch_report(
                batch_result
            )

            return batch_result

        valid_mailboxes = precheck_result.get(
            "valid_mailboxes",
            []
        )

        invalid_mailboxes = precheck_result.get(
            "invalid_mailboxes",
            []
        )

        for invalid_mailbox in invalid_mailboxes:
            batch_result["processed"] += 1
            batch_result["failed_count"] += 1

            error_message = f"Mailbox inválida no precheck: {invalid_mailbox}"

            batch_result["errors"].append(
                error_message
            )

            batch_result["results"].append(
                {
                    "mailbox": invalid_mailbox,
                    "success": False,
                    "backup_path": None,
                    "folders_count": 0,
                    "folders_processed": 0,
                    "messages_indexed": 0,
                    "messages_exported": 0,
                    "messages_failed": 0,
                    "attachments_exported": 0,
                    "errors": [
                        error_message
                    ]
                }
            )

        if resume_path and len(valid_mailboxes) != 1:
            error_message = (
                "O parâmetro resume_path só pode ser usado com um CSV contendo "
                "uma única mailbox válida."
            )

            batch_result["errors"].append(error_message)

            batch_finished_at = datetime.now()

            batch_result["finished_at"] = batch_finished_at.isoformat()
            batch_result["duration_seconds"] = round(
                (batch_finished_at - batch_started_at).total_seconds(),
                2
            )

            self.save_batch_report(
                batch_result
            )

            self.logger.error(
                error_message
            )

            return batch_result

        if not valid_mailboxes:
            batch_finished_at = datetime.now()

            batch_result["finished_at"] = batch_finished_at.isoformat()
            batch_result["duration_seconds"] = round(
                (batch_finished_at - batch_started_at).total_seconds(),
                2
            )

            if not batch_result["errors"]:
                batch_result["errors"].append(
                    "Nenhuma mailbox válida encontrada após precheck."
                )

            self.save_batch_report(
                batch_result
            )

            self.logger.info(
                "Backup em lote finalizado sem mailboxes válidas para processar."
            )

            return batch_result

        for position, mailbox_email in enumerate(valid_mailboxes, start=1):
            self.logger.info(
                f"Processando mailbox válida {position}/{len(valid_mailboxes)}: {mailbox_email}"
            )

            try:
                result = self.export_mailbox_complete(
                    mailbox_email=mailbox_email,
                    message_limit_per_folder=message_limit_per_folder,
                    export_all_messages=export_all_messages,
                    export_attachments=export_attachments,
                    skip_calendar=skip_calendar,
                    skip_contacts=skip_contacts,
                    skip_tasks=skip_tasks,
                    resume_path=resume_path,
                    phase=5,
                    excluded_folder_names=excluded_folder_names,
                    selected_folder_ids=selected_folder_ids,
                    selected_folder_paths=selected_folder_paths
                )

                batch_result["processed"] += 1

                item_result = {
                    "mailbox": mailbox_email,
                    "success": result.get("success", False),
                    "backup_path": result.get("backup_path"),
                    "folders_count": result.get("folders_count", 0),
                    "folders_processed": result.get("folders_processed", 0),
                    "messages_indexed": result.get("messages_indexed", 0),
                    "messages_exported": result.get("messages_exported", 0),
                    "messages_failed": result.get("messages_failed", 0),
                    "attachments_exported": result.get("attachments_exported", 0),
                    "errors": result.get("errors", [])
                }

                batch_result["results"].append(
                    item_result
                )

                if result.get("success"):
                    batch_result["success_count"] += 1
                else:
                    batch_result["failed_count"] += 1
                    batch_result["errors"].append(
                        f"Falha na mailbox {mailbox_email}"
                    )

            except OperationInterrupted:
                raise
            except Exception as error:
                self.logger.error(
                    f"Erro inesperado ao processar {mailbox_email}: {error}"
                )

                batch_result["processed"] += 1
                batch_result["failed_count"] += 1

                error_message = (
                    f"Erro inesperado ao processar {mailbox_email}: {error}"
                )

                batch_result["errors"].append(
                    error_message
                )

                batch_result["results"].append(
                    {
                        "mailbox": mailbox_email,
                        "success": False,
                        "backup_path": None,
                        "folders_count": 0,
                        "folders_processed": 0,
                        "messages_indexed": 0,
                        "messages_exported": 0,
                        "messages_failed": 0,
                        "attachments_exported": 0,
                        "errors": [
                            error_message
                        ]
                    }
                )

        batch_finished_at = datetime.now()

        batch_result["finished_at"] = batch_finished_at.isoformat()
        batch_result["duration_seconds"] = round(
            (batch_finished_at - batch_started_at).total_seconds(),
            2
        )

        batch_result["success"] = batch_result["failed_count"] == 0

        self.save_batch_report(
            batch_result
        )

        self.logger.info(
            "Backup em lote finalizado. "
            f"Sucesso: {batch_result['success_count']} | "
            f"Falha: {batch_result['failed_count']}"
        )

        return batch_result

    def save_batch_report(self, batch_result):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        reports_dir = OUTPUT_DIR / "_batch_reports"

        reports_dir.mkdir(
            parents=True,
            exist_ok=True
        )

        json_report_path = reports_dir / f"batch_report_{timestamp}.json"
        csv_report_path = reports_dir / f"batch_report_{timestamp}.csv"

        batch_result["report_path"] = str(json_report_path)
        batch_result["csv_report_path"] = str(csv_report_path)

        self.save_json(
            json_report_path,
            batch_result
        )

        with open(
            csv_report_path,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as csv_file:
            fieldnames = [
                "mailbox",
                "success",
                "backup_path",
                "folders_count",
                "folders_processed",
                "messages_indexed",
                "messages_exported",
                "messages_failed",
                "attachments_exported",
                "errors_count"
            ]

            writer = csv.DictWriter(
                csv_file,
                fieldnames=fieldnames
            )

            writer.writeheader()

            for item in batch_result.get("results", []):
                writer.writerow(
                    {
                        "mailbox": item.get("mailbox"),
                        "success": item.get("success"),
                        "backup_path": item.get("backup_path"),
                        "folders_count": item.get("folders_count", 0),
                        "folders_processed": item.get("folders_processed", 0),
                        "messages_indexed": item.get("messages_indexed", 0),
                        "messages_exported": item.get("messages_exported", 0),
                        "messages_failed": item.get("messages_failed", 0),
                        "attachments_exported": item.get("attachments_exported", 0),
                        "errors_count": len(item.get("errors", []))
                    }
                )

    def write_audit_report(self, backup_folders, result):
        audit_data = {
            "mailbox": result.get("mailbox"),
            "phase": result.get("phase"),
            "success": result.get("success"),
            "backup_path": result.get("backup_path"),
            "started_at": result.get("started_at"),
            "finished_at": result.get("finished_at"),
            "duration_seconds": result.get("duration_seconds", 0),
            "folders_count": result.get("folders_count", 0),
            "folders_processed": result.get("folders_processed", 0),
            "folders_skipped": result.get("folders_skipped", 0),
            "messages_indexed": result.get("messages_indexed", 0),
            "messages_exported": result.get("messages_exported", 0),
            "messages_skipped": result.get("messages_skipped", 0),
            "messages_failed": result.get("messages_failed", 0),
            "attachments_exported": result.get("attachments_exported", 0),
            "calendar_events": result.get("calendar_events", 0),
            "contacts": result.get("contacts", 0),
            "task_lists": result.get("task_lists", 0),
            "task_items": result.get("task_items", 0),
            "errors_count": len(result.get("errors", [])),
            "errors": result.get("errors", [])
        }

        self.save_json(
            backup_folders["root"] / "audit_report.json",
            audit_data
        )

        csv_path = backup_folders["root"] / "audit_report.csv"

        with open(
            csv_path,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as csv_file:
            fieldnames = [
                "mailbox",
                "phase",
                "success",
                "backup_path",
                "started_at",
                "finished_at",
                "duration_seconds",
                "folders_count",
                "folders_processed",
                "folders_skipped",
                "messages_indexed",
                "messages_exported",
                "messages_skipped",
                "messages_failed",
                "attachments_exported",
                "calendar_events",
                "contacts",
                "task_lists",
                "task_items",
                "errors_count"
            ]

            writer = csv.DictWriter(
                csv_file,
                fieldnames=fieldnames
            )

            writer.writeheader()

            writer.writerow(
                {
                    "mailbox": audit_data["mailbox"],
                    "phase": audit_data["phase"],
                    "success": audit_data["success"],
                    "backup_path": audit_data["backup_path"],
                    "started_at": audit_data["started_at"],
                    "finished_at": audit_data["finished_at"],
                    "duration_seconds": audit_data["duration_seconds"],
                    "folders_count": audit_data["folders_count"],
                    "folders_processed": audit_data["folders_processed"],
                    "folders_skipped": audit_data["folders_skipped"],
                    "messages_indexed": audit_data["messages_indexed"],
                    "messages_exported": audit_data["messages_exported"],
                    "messages_skipped": audit_data["messages_skipped"],
                    "messages_failed": audit_data["messages_failed"],
                    "attachments_exported": audit_data["attachments_exported"],
                    "calendar_events": audit_data["calendar_events"],
                    "contacts": audit_data["contacts"],
                    "task_lists": audit_data["task_lists"],
                    "task_items": audit_data["task_items"],
                    "errors_count": audit_data["errors_count"]
                }
            )