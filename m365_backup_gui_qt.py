import csv
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

import msal
import requests
from dotenv import load_dotenv
from PySide6.QtCore import (
    QObject, Qt, QThread, QTimer, QUrl, Signal, QSettings, QItemSelectionModel
)
from PySide6.QtGui import (
    QAction, QColor, QDesktopServices, QFont, QIcon, QKeySequence, QShortcut
)
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QFormLayout, QFrame, QGridLayout, QGroupBox, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QDoubleSpinBox,
    QMenu, QMessageBox, QProgressBar, QPushButton, QScrollArea, QSpinBox,
    QSplitter, QStackedWidget, QTableWidget, QTableWidgetItem, QTextEdit,
    QToolButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget, QTabWidget, QAbstractItemView
)


try:
    from application_runtime import (
        APP_VERSION, AppSettings, CredentialStore, CrashReporter,
        EnvironmentDiagnostics, IntegrityValidator, SingleInstance
    )
except ImportError:
    from .application_runtime import (
        APP_VERSION, AppSettings, CredentialStore, CrashReporter,
        EnvironmentDiagnostics, IntegrityValidator, SingleInstance
    )

from src.services.api_metrics_store import ApiMetricsStore


APP_TITLE = "M365 Mailbox Backup"
STATUS_RUNNING = {"executando", "continuando"}
STATUS_RESUMABLE = {"pausado", "interrompido", "erro", "incompleto"}


def safe_int(value, default=0):
    try:
        return int(value or 0)
    except Exception:
        return default


def format_bytes(value):
    value = float(value or 0)
    units = ("B", "KB", "MB", "GB", "TB")
    index = 0
    while value >= 1024 and index < len(units) - 1:
        value /= 1024
        index += 1
    return f"{int(value)} {units[index]}" if index == 0 else f"{value:.2f} {units[index]}"


def open_local_path(path):
    path = Path(path).expanduser().resolve()
    path.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(path))
    else:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


class OutputWorker(QObject):
    line = Signal(str)
    finished = Signal(int)
    failed = Signal(str)

    def __init__(self, command, cwd, environment):
        super().__init__()
        self.command = command
        self.cwd = cwd
        self.environment = environment
        self.process = None

    def run(self):
        try:
            self.process = subprocess.Popen(
                self.command,
                cwd=str(self.cwd),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=self.environment,
                creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
            )
            stdout = self.process.stdout
            if stdout is None:
                raise RuntimeError("O processo foi iniciado sem canal de saída.")
            for line in stdout:
                self.line.emit(line.rstrip())
            self.finished.emit(self.process.wait())
        except Exception as error:
            self.failed.emit(str(error))
            self.finished.emit(-1)

    def stop(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()


class FolderLoadWorker(QObject):
    loaded = Signal(str, list)
    progress = Signal(str, int, int)
    failed = Signal(str, str)
    completed = Signal()

    def __init__(self, mailbox, env_file, metrics_path):
        super().__init__()
        self.mailbox = str(mailbox)
        self.env_file = str(env_file)
        self.metrics_path = str(metrics_path)

    def run(self):
        try:
            folders = self._fetch_folder_tree_light()
            self.loaded.emit(self.mailbox, folders)
        except Exception as error:
            self.failed.emit(self.mailbox, str(error))
        finally:
            self.completed.emit()

    def _graph_token(self):
        load_dotenv(self.env_file)
        tenant = os.getenv("TENANT_ID")
        client = os.getenv("CLIENT_ID")
        secret = os.getenv("CLIENT_SECRET")
        if not tenant or not client or not secret:
            raise RuntimeError(
                "TENANT_ID, CLIENT_ID ou CLIENT_SECRET ausente no .env."
            )
        application = msal.ConfidentialClientApplication(
            client,
            client_credential=secret,
            authority=f"https://login.microsoftonline.com/{tenant}"
        )
        raw_result = application.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        result = raw_result if isinstance(raw_result, dict) else {}
        if "access_token" not in result:
            raise RuntimeError(
                result.get("error_description") or "Falha na autenticação."
            )
        return result["access_token"]

    def _fetch_folder_tree_light(self):
        token = self._graph_token()
        metrics = ApiMetricsStore(self.metrics_path)
        base = "https://graph.microsoft.com/v1.0"
        folders = []
        seen = set()
        running_total = 0

        def load(url, parent_id=None, parent_path=""):
            nonlocal running_total
            while url:
                started = time.perf_counter()
                response = requests.get(
                    url,
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Accept": "application/json"
                    },
                    timeout=60
                )
                duration = (time.perf_counter() - started) * 1000
                try:
                    metrics.record(
                        self.mailbox,
                        "Listagem de pastas",
                        "GET",
                        response.status_code,
                        response.status_code < 400,
                        duration_ms=int(round(duration)),
                        bytes_received=len(response.content or b""),
                        friendly_error=(
                            None if response.status_code < 400
                            else "As pastas não puderam ser consultadas."
                        )
                    )
                except Exception:
                    pass
                response.raise_for_status()
                data = response.json()
                for folder in data.get("value", []):
                    folder_id = folder.get("id")
                    if not folder_id or folder_id in seen:
                        continue
                    seen.add(folder_id)
                    name = folder.get("displayName") or "Sem nome"
                    path = f"{parent_path}/{name}" if parent_path else name
                    item = dict(folder)
                    item["parent_id"] = parent_id
                    item["path"] = path
                    folders.append(item)
                    running_total += safe_int(folder.get("totalItemCount"))
                    self.progress.emit(
                        self.mailbox,
                        running_total,
                        len(folders)
                    )
                    if safe_int(folder.get("childFolderCount")) > 0:
                        load(
                            base + f"/users/{self.mailbox}/mailFolders/"
                            + quote(str(folder_id), safe="")
                            + "/childFolders?$top=100"
                            + "&$select=id,displayName,totalItemCount,childFolderCount",
                            folder_id,
                            path
                        )
                url = data.get("@odata.nextLink")

        load(
            base + f"/users/{self.mailbox}/mailFolders?$top=100"
            + "&$select=id,displayName,totalItemCount,childFolderCount"
        )
        return folders


class CoordinatorPollWorker(QObject):
    completed = Signal(list, list)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, coordinator_url, after_event_id):
        super().__init__()
        self.coordinator_url = coordinator_url
        self.after_event_id = int(after_event_id or 0)

    def _get_json(self, path, timeout):
        request = Request(self.coordinator_url + path, method="GET")
        with urlopen(request, timeout=timeout) as response:
            content = response.read()
            return json.loads(content.decode("utf-8")) if content else None

    def run(self):
        try:
            operations = self._get_json("/operations", 1.5) or []
            events = self._get_json(
                f"/events?after_id={self.after_event_id}&limit=250", 1.5
            ) or []
            self.completed.emit(operations, events)
        except Exception as error:
            self.failed.emit(str(error))
        finally:
            self.finished.emit()

class StatCard(QFrame):
    def __init__(self, title, value="-"):
        super().__init__()
        self.setObjectName("statCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        title_label = QLabel(title)
        title_label.setObjectName("muted")
        self.value_label = QLabel(value)
        self.value_label.setObjectName("statValue")
        layout.addWidget(title_label)
        layout.addWidget(self.value_label)

    def set_value(self, value):
        self.value_label.setText(str(value))


class FolderScopeDialog(QDialog):
    def __init__(self, mailbox, folders, existing=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Pastas e opções · {mailbox}")
        self.resize(1050, 720)
        self.mailbox = mailbox
        self.folders = folders
        self.existing = existing or {}
        self.result_config = None

        root = QVBoxLayout(self)
        heading = QLabel("Escolha o conteúdo deste backup")
        heading.setObjectName("pageTitle")
        subtitle = QLabel(
            "A consulta abaixo carrega somente a estrutura e os contadores das pastas; "
            "nenhuma mensagem é aberta nesta etapa."
        )
        subtitle.setWordWrap(True)
        subtitle.setObjectName("muted")
        root.addWidget(heading)
        root.addWidget(subtitle)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        root.addWidget(splitter, 1)

        left = QFrame()
        left_layout = QVBoxLayout(left)
        tools = QHBoxLayout()
        select_all = QPushButton("Selecionar tudo")
        clear_all = QPushButton("Desmarcar tudo")
        recommended = QPushButton("Seleção recomendada")
        tools.addWidget(select_all)
        tools.addWidget(clear_all)
        tools.addWidget(recommended)
        tools.addStretch()
        left_layout.addLayout(tools)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Pasta", "Itens informados"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tree.header().resizeSection(1, 120)
        left_layout.addWidget(self.tree)
        splitter.addWidget(left)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right = QWidget()
        options = QVBoxLayout(right)
        content_box = QGroupBox("Conteúdo")
        content_layout = QVBoxLayout(content_box)
        self.all_messages = QCheckBox("Exportar todas as mensagens")
        self.attachments = QCheckBox("Exportar anexos separadamente")
        self.skip_calendar = QCheckBox("Ignorar calendário")
        self.skip_contacts = QCheckBox("Ignorar contatos")
        self.skip_tasks = QCheckBox("Ignorar tarefas")
        self.profile_only = QCheckBox("Somente conteúdo do perfil principal")
        self.skip_precheck = QCheckBox("Iniciar sem pré-análise completa")
        for widget in (
            self.all_messages, self.attachments, self.skip_calendar,
            self.skip_contacts, self.skip_tasks, self.profile_only,
            self.skip_precheck
        ):
            content_layout.addWidget(widget)
        options.addWidget(content_box)

        limits = QGroupBox("Volume")
        limits_form = QFormLayout(limits)
        self.limit = QSpinBox()
        self.limit.setRange(0, 1_000_000)
        self.limit.setSpecialValueText("Sem limite específico")
        limits_form.addRow("Limite por pasta:", self.limit)
        options.addWidget(limits)

        summary_box = QGroupBox("Resumo")
        summary_layout = QVBoxLayout(summary_box)
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        summary_layout.addWidget(self.summary)
        options.addWidget(summary_box)
        options.addStretch()
        right_scroll.setWidget(right)
        splitter.addWidget(right_scroll)
        splitter.setSizes([680, 370])

        actions = QDialogButtonBox()
        self.apply_all_button = actions.addButton(
            "Salvar e aplicar à fila", QDialogButtonBox.ButtonRole.ActionRole
        )
        self.save_button = actions.addButton("Salvar", QDialogButtonBox.ButtonRole.AcceptRole)
        actions.addButton(QDialogButtonBox.StandardButton.Cancel)
        root.addWidget(actions)

        self.item_by_id = {}
        self.folder_by_id = {str(item.get("id")): item for item in folders}
        self._load_options()
        self.tree.itemChanged.connect(self.update_summary)
        self._folder_queue = sorted(
            list(self.folders),
            key=lambda folder: (
                str(folder.get("path") or "").count("/"),
                str(folder.get("path") or folder.get("displayName") or "").lower()
            )
        )
        self._folder_total = len(self._folder_queue)
        self._folder_loaded = 0
        self._tree_loading = True
        self.save_button.setEnabled(False)
        self.apply_all_button.setEnabled(False)
        self.tree.setEnabled(False)
        self.summary.setText(
            f"Preparando a visualização de {self._folder_total} pasta(s)..."
        )
        QTimer.singleShot(0, self._populate_tree_batch)
        for widget in (
            self.all_messages, self.attachments, self.skip_calendar,
            self.skip_contacts, self.skip_tasks, self.profile_only,
            self.skip_precheck
        ):
            widget.toggled.connect(self.update_summary)
        self.limit.valueChanged.connect(self.update_summary)
        select_all.clicked.connect(lambda: self.set_all_checked(True))
        clear_all.clicked.connect(lambda: self.set_all_checked(False))
        recommended.clicked.connect(self.select_recommended)
        self.save_button.clicked.connect(lambda: self.accept_config(False))
        self.apply_all_button.clicked.connect(lambda: self.accept_config(True))
        actions.rejected.connect(self.reject)
        self.update_summary()

    def _populate_tree_batch(self):
        if not self._tree_loading:
            return

        batch_size = 120
        batch = self._folder_queue[:batch_size]
        del self._folder_queue[:batch_size]
        selected = set(self.existing.get("selected_folder_ids") or [])

        self.tree.blockSignals(True)
        self.tree.setUpdatesEnabled(False)
        try:
            for folder in batch:
                folder_id = str(folder.get("id") or "")
                if not folder_id:
                    continue
                parent_id = str(folder.get("parent_id") or "")
                item = QTreeWidgetItem([
                    folder.get("displayName") or "Sem nome",
                    str(folder.get("totalItemCount", 0) or 0)
                ])
                item.setData(0, Qt.ItemDataRole.UserRole, folder_id)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                checked = not selected or folder_id in selected
                item.setCheckState(0, Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

                parent_item = self.item_by_id.get(parent_id)
                if parent_item is not None:
                    parent_item.addChild(item)
                else:
                    # Pastas com pai ausente não bloqueiam a montagem da janela.
                    self.tree.addTopLevelItem(item)

                self.item_by_id[folder_id] = item
                self._folder_loaded += 1
        finally:
            self.tree.setUpdatesEnabled(True)
            self.tree.blockSignals(False)

        if self._folder_queue:
            self.summary.setText(
                f"Preparando pastas: {self._folder_loaded}/{self._folder_total}..."
            )
            QTimer.singleShot(1, self._populate_tree_batch)
            return

        self._tree_loading = False
        self.tree.setEnabled(True)
        self.save_button.setEnabled(True)
        self.apply_all_button.setEnabled(True)
        self.tree.expandToDepth(0)
        self.update_summary()

    def _load_options(self):
        defaults = {
            "all_messages": True,
            "attachments": False,
            "skip_calendar": True,
            "skip_contacts": True,
            "skip_tasks": True,
            "profile_only": False,
            "skip_precheck": True
        }
        for key, default in defaults.items():
            getattr(self, key).setChecked(bool(self.existing.get(key, default)))
        self.limit.setValue(safe_int(self.existing.get("limit"), 0))

    def checked_folder_ids(self):
        return [
            folder_id for folder_id, item in self.item_by_id.items()
            if item.checkState(0) == Qt.CheckState.Checked
        ]

    def set_all_checked(self, checked):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        self.tree.blockSignals(True)
        for item in self.item_by_id.values():
            item.setCheckState(0, state)
        self.tree.blockSignals(False)
        self.update_summary()

    def select_recommended(self):
        excluded = (
            "itens excluídos", "deleted items", "lixo eletrônico", "junk email",
            "rss feeds", "problemas de sincronização", "sync issues",
            "arquivo morto", "archive"
        )
        self.tree.blockSignals(True)
        for folder_id, item in self.item_by_id.items():
            path = (self.folder_by_id[folder_id].get("path") or "").lower()
            item.setCheckState(
                0, Qt.CheckState.Unchecked if any(term in path for term in excluded) else Qt.CheckState.Checked
            )
        self.tree.blockSignals(False)
        self.update_summary()

    def update_summary(self):
        ids = self.checked_folder_ids()
        total = sum(
            safe_int(self.folder_by_id[item].get("totalItemCount"))
            for item in ids if item in self.folder_by_id
        )
        self.summary.setText(
            f"{len(ids)} de {len(self.folders)} pastas selecionadas\n"
            f"{total:,} itens informados pelo Microsoft Graph\n"
            "A quantidade é apenas um contador leve das pastas."
        )

    def accept_config(self, apply_all):
        if getattr(self, "_tree_loading", False):
            QMessageBox.information(
                self,
                "Pastas",
                "A lista de pastas ainda está sendo preparada. Aguarde alguns instantes."
            )
            return
        ids = self.checked_folder_ids()
        if not ids:
            QMessageBox.warning(self, "Pastas", "Selecione pelo menos uma pasta.")
            return
        self.result_config = {
            "selected_folder_ids": ids,
            "selected_folder_paths": sorted(
                self.folder_by_id[item].get("path") for item in ids
                if item in self.folder_by_id
            ),
            "all_messages": self.all_messages.isChecked(),
            "attachments": self.attachments.isChecked(),
            "skip_calendar": self.skip_calendar.isChecked(),
            "skip_contacts": self.skip_contacts.isChecked(),
            "skip_tasks": self.skip_tasks.isChecked(),
            "profile_only": self.profile_only.isChecked(),
            "skip_precheck": self.skip_precheck.isChecked(),
            "limit": self.limit.value() or "",
            "apply_all": apply_all
        }
        self.accept()


class PstCustomizationDialog(QDialog):
    """Configuração completa e persistível de uma conversão PST."""
    def __init__(self, defaults, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Nova conversão PST")
        self.resize(820, 650)
        self.result_config = None
        root = QVBoxLayout(self)
        tabs = QTabWidget()
        root.addWidget(tabs, 1)

        general = QWidget(); form = QFormLayout(general)
        self.source = QLineEdit(str(defaults.get("source", "")))
        source_row = QHBoxLayout(); source_row.addWidget(self.source, 1)
        source_button = QPushButton("Selecionar"); source_button.clicked.connect(self.choose_source)
        source_row.addWidget(source_button); form.addRow("Backup de origem:", source_row)
        self.destination_dir = QLineEdit(str(defaults.get("destination_dir", "")))
        destination_row = QHBoxLayout(); destination_row.addWidget(self.destination_dir, 1)
        destination_button = QPushButton("Selecionar"); destination_button.clicked.connect(self.choose_destination)
        destination_row.addWidget(destination_button); form.addRow("Pasta de destino:", destination_row)
        self.file_name = QLineEdit(str(defaults.get("file_name", "backup.pst")))
        self.display_name = QLineEdit(str(defaults.get("display_name", "M365 Mailbox Backup")))
        form.addRow("Nome do arquivo:", self.file_name)
        form.addRow("Nome exibido no Outlook:", self.display_name)
        self.existing_action = QComboBox()
        for label, value in (("Retomar usando checkpoint", "resume"), ("Criar nome numerado", "number"), ("Substituir PST e auxiliares", "replace"), ("Cancelar se existir", "cancel")):
            self.existing_action.addItem(label, value)
        idx = self.existing_action.findData(defaults.get("existing_action", "resume")); self.existing_action.setCurrentIndex(max(0, idx))
        form.addRow("Se o PST já existir:", self.existing_action)
        self.manual_start = QCheckBox("Adicionar à fila como pendente")
        self.manual_start.setChecked(bool(defaults.get("manual_start", True)))
        form.addRow("", self.manual_start)
        tabs.addTab(general, "Geral")

        content = QWidget(); content_form = QFormLayout(content)
        self.folder_mode = QComboBox(); self.folder_mode.addItem("Preservar árvore original", "preserve"); self.folder_mode.addItem("Reunir em uma pasta única", "single")
        idx = self.folder_mode.findData(defaults.get("folder_mode", "preserve")); self.folder_mode.setCurrentIndex(max(0, idx))
        self.root_folder_name = QLineEdit(str(defaults.get("root_folder_name", "")))
        self.root_folder_name.setPlaceholderText("Opcional, por exemplo: Backup Lucas Duarte")
        self.visible_metadata = QCheckBox("Mostrar metadados originais no corpo da mensagem")
        self.visible_metadata.setChecked(bool(defaults.get("visible_metadata", True)))
        self.import_attachments = QCheckBox("Importar anexos e imagens incorporadas")
        self.import_attachments.setChecked(bool(defaults.get("import_attachments", True)))
        self.image_width = QSpinBox(); self.image_width.setRange(200, 2000); self.image_width.setSuffix(" px"); self.image_width.setValue(safe_int(defaults.get("image_max_width"), 700))
        content_form.addRow("Organização das pastas:", self.folder_mode)
        content_form.addRow("Pasta raiz personalizada:", self.root_folder_name)
        content_form.addRow("", self.visible_metadata); content_form.addRow("", self.import_attachments)
        content_form.addRow("Largura máxima de imagens:", self.image_width)
        tabs.addTab(content, "Conteúdo")

        safety = QWidget(); safety_form = QFormLayout(safety)
        self.import_rate = QDoubleSpinBox(); self.import_rate.setRange(0.5, 100.0); self.import_rate.setDecimals(1); self.import_rate.setSuffix(" EML/s"); self.import_rate.setValue(float(defaults.get("import_rate", 10) or 10))
        self.verification = QComboBox()
        self.verification.addItem("Balanceada, recomendada", "balanced")
        self.verification.addItem("Rápida, auditoria ao final", "quick")
        self.verification.addItem("Completa, confirmação por item", "complete")
        idx = self.verification.findData(defaults.get("verification_level", "balanced")); self.verification.setCurrentIndex(max(0, idx))
        self.verification_batch_size = QSpinBox(); self.verification_batch_size.setRange(1, 500); self.verification_batch_size.setValue(safe_int(defaults.get("verification_batch_size"), 25))
        self.performance_profile = QComboBox()
        for label, value in (("Equilibrado", "balanced"), ("Conservador", "conservative"), ("Desempenho", "performance"), ("Personalizado", "custom")):
            self.performance_profile.addItem(label, value)
        profile_index = self.performance_profile.findData(defaults.get("performance_profile", "balanced")); self.performance_profile.setCurrentIndex(max(0, profile_index))
        self.adaptive_enabled = QCheckBox("Ajustar fila e workers automaticamente"); self.adaptive_enabled.setChecked(bool(defaults.get("adaptive_enabled", True)))
        self.memory_budget_mb = QSpinBox(); self.memory_budget_mb.setRange(128, 8192); self.memory_budget_mb.setSuffix(" MB"); self.memory_budget_mb.setValue(safe_int(defaults.get("memory_budget_mb"), 512))
        self.min_prepare_workers = QSpinBox(); self.min_prepare_workers.setRange(1, 8); self.min_prepare_workers.setValue(safe_int(defaults.get("min_prepare_workers"), 1))
        self.prepare_workers = QSpinBox(); self.prepare_workers.setRange(1, 8); self.prepare_workers.setValue(safe_int(defaults.get("max_prepare_workers", defaults.get("prepare_workers", 4)), 4))
        self.prepare_queue_size = QSpinBox(); self.prepare_queue_size.setRange(2, 100); self.prepare_queue_size.setValue(safe_int(defaults.get("prepare_queue_size"), 12))
        self.large_eml_mb = QSpinBox(); self.large_eml_mb.setRange(1, 500); self.large_eml_mb.setSuffix(" MB"); self.large_eml_mb.setValue(safe_int(defaults.get("large_eml_mb"), 25))
        self.detach_after = QCheckBox("Remover o PST do Outlook ao concluir")
        self.detach_after.setChecked(bool(defaults.get("detach_after", False)))
        self.open_folder_after = QCheckBox("Abrir a pasta de destino ao concluir")
        self.open_folder_after.setChecked(bool(defaults.get("open_folder_after", False)))
        safety_form.addRow("Velocidade:", self.import_rate); safety_form.addRow("Modo de verificação:", self.verification)
        safety_form.addRow("Lote de confirmações:", self.verification_batch_size)
        safety_form.addRow("Perfil de desempenho:", self.performance_profile)
        safety_form.addRow("", self.adaptive_enabled)
        safety_form.addRow("Orçamento de memória:", self.memory_budget_mb)
        safety_form.addRow("Workers mínimos:", self.min_prepare_workers)
        safety_form.addRow("Workers máximos:", self.prepare_workers)
        safety_form.addRow("Fila de preparação:", self.prepare_queue_size)
        safety_form.addRow("EML considerado grande:", self.large_eml_mb)
        safety_form.addRow("", self.detach_after); safety_form.addRow("", self.open_folder_after)
        info = QLabel("A retomada usa checkpoint SQLite, EntryID, StoreID e chave de origem para evitar duplicações.")
        info.setWordWrap(True); info.setObjectName("muted"); safety_form.addRow("Segurança:", info)
        tabs.addTab(safety, "Desempenho e segurança")

        self.preview = QLabel(); self.preview.setWordWrap(True); self.preview.setObjectName("infoBanner")
        root.addWidget(self.preview)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("Adicionar à fila")
        buttons.accepted.connect(self.accept_config); buttons.rejected.connect(self.reject)
        root.addWidget(buttons)
        for widget in (self.source, self.destination_dir, self.file_name, self.display_name):
            widget.textChanged.connect(self.update_preview)
        self.update_preview()

    def choose_source(self):
        path = QFileDialog.getExistingDirectory(self, "Selecionar backup", self.source.text())
        if path:
            self.source.setText(path)
            if self.file_name.text().strip() in ("", "backup.pst"):
                self.file_name.setText(Path(path).name + ".pst")

    def choose_destination(self):
        path = QFileDialog.getExistingDirectory(self, "Pasta de destino", self.destination_dir.text())
        if path: self.destination_dir.setText(path)

    def final_path(self):
        name = self.file_name.text().strip() or "backup.pst"
        if not name.lower().endswith(".pst"): name += ".pst"
        name = re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .")
        return Path(self.destination_dir.text().strip()).expanduser() / name

    def update_preview(self, *args):
        self.preview.setText(f"Arquivo final: {self.final_path()}\nNome no Outlook: {self.display_name.text().strip() or 'M365 Mailbox Backup'}")

    def accept_config(self):
        source = Path(self.source.text().strip()).expanduser()
        destination = Path(self.destination_dir.text().strip()).expanduser()
        if not source.exists():
            QMessageBox.warning(self, "Conversão PST", "Selecione uma origem existente."); return
        if not self.file_name.text().strip() or not self.display_name.text().strip():
            QMessageBox.warning(self, "Conversão PST", "Informe o nome do arquivo e o nome exibido no Outlook."); return
        try: destination.mkdir(parents=True, exist_ok=True)
        except Exception as error:
            QMessageBox.critical(self, "Conversão PST", f"O destino não pôde ser criado: {error}"); return
        self.result_config = {
            "source": str(source.resolve()), "pst_path": str(self.final_path().resolve()),
            "file_name": self.final_path().name, "display_name": self.display_name.text().strip(),
            "existing_action": self.existing_action.currentData(), "manual_start": self.manual_start.isChecked(),
            "folder_mode": self.folder_mode.currentData(), "root_folder_name": self.root_folder_name.text().strip(),
            "visible_metadata": self.visible_metadata.isChecked(), "import_attachments": self.import_attachments.isChecked(),
            "image_max_width": self.image_width.value(), "import_rate": self.import_rate.value(),
            "verification_level": self.verification.currentData(),
            "verification_batch_size": self.verification_batch_size.value(),
            "detach_after": self.detach_after.isChecked(),
            "open_folder_after": self.open_folder_after.isChecked(),
            "prepare_workers": self.prepare_workers.value(),
            "performance_profile": self.performance_profile.currentData(),
            "adaptive_enabled": self.adaptive_enabled.isChecked(),
            "memory_budget_mb": self.memory_budget_mb.value(),
            "min_prepare_workers": self.min_prepare_workers.value(),
            "max_prepare_workers": self.prepare_workers.value(),
            "prepare_queue_size": self.prepare_queue_size.value(),
            "large_eml_mb": self.large_eml_mb.value()
        }
        self.accept()

class DraggableBackupTable(QTableWidget):
    orderChanged = Signal(list)

    def __init__(self, rows, columns, parent=None):
        super().__init__(rows, columns, parent)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDragDropOverwriteMode(False)
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)

    def dropEvent(self, event):
        super().dropEvent(event)
        QTimer.singleShot(0, self._emit_order)

    def _emit_order(self):
        order = []
        for row in range(self.rowCount()):
            item = self.item(row, 0)
            mailbox = item.data(Qt.ItemDataRole.UserRole) if item else None
            if mailbox and mailbox not in order:
                order.append(mailbox)
        if order:
            self.orderChanged.emit(order)

class M365BackupWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1500, 900)
        self.setMinimumSize(1120, 720)

        self.project_root = (
            Path(sys.executable).resolve().parent
            if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parent
        )
        self.temp_dir = self.project_root / "_temp_gui_jobs"
        self.state_dir = self.project_root / "_gui_state"
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.state_dir.mkdir(parents=True, exist_ok=True)
        self.state_path = self.state_dir / "backup_queue_state_qt.json"
        self.metrics_path = self.state_dir / "api_metrics.sqlite3"
        self.metrics = ApiMetricsStore(self.metrics_path)
        self.app_settings = AppSettings(self.project_root)
        self.credential_store = CredentialStore(self.project_root)
        self.crash_reporter = CrashReporter(self.project_root)
        self.app_settings.apply_environment()
        if not self.credential_store.load() and (self.project_root / ".env").exists():
            self.credential_store.import_env(self.project_root / ".env")
        self.credential_store.apply_environment()
        self.coordinator_url = "http://127.0.0.1:8765/api/v1"
        self.coordinator_process = None
        self.coordinator_online = False
        self.last_event_id = 0

        self.backup_jobs = {}
        self.backup_order = []
        self.backup_workers = {}
        self.backup_threads = {}
        self.pst_jobs = {}
        self.pst_workers = {}
        self.pst_threads = {}
        self.scheduler_active = False
        self.pst_counter = 0
        self.folder_defaults = None
        self._start_config_queue = []
        self._start_config_active = False
        self._start_config_started = 0
        self._start_config_skipped = 0
        self._folder_request_context = "manual"
        self._folder_cache = {}
        self._analysis_queue = []
        self._analysis_active_mailbox = None
        self._pending_folder_dialog = None
        self.mailbox_logs = {}
        self.max_mailbox_log_lines = 1200
        self._coordinator_poll_running = False
        self._last_operations_signature = None
        self._backup_render_pending = False
        self._pst_render_pending = False
        self._state_save_pending = False
        self._build_ui()
        self.setup_standard_shortcuts()
        self.load_runtime_settings_into_ui()
        self._apply_style()
        self.load_state()
        self.ensure_coordinator()

        self.coordinator_timer = QTimer(self)
        self.coordinator_timer.timeout.connect(self.refresh_from_coordinator)
        self.coordinator_timer.start(1500)

        self.log_event("Interface PySide6 iniciada.", "Informação")
        if self.crash_reporter.previous_unclean:
            QTimer.singleShot(800, lambda: QMessageBox.warning(
                self, "Recuperação",
                "A execução anterior não foi encerrada normalmente. Verifique os logs e os checkpoints antes de retomar operações."
            ))


    def coordinator_request(self, method, path, payload=None, timeout: float = 3.0):
        headers = {}
        data = None
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        request = Request(
            self.coordinator_url + path,
            data=data,
            headers=headers,
            method=method
        )
        with urlopen(request, timeout=timeout) as response:
            content = response.read()
            return json.loads(content.decode("utf-8")) if content else None

    def ensure_coordinator(self):
        try:
            health = self.coordinator_request("GET", "/health", timeout=0.6) or {}
            self.coordinator_pid = safe_int(health.get("coordinator_pid"))
            self.coordinator_online = True
            return True
        except Exception:
            pass
        command = (
            [str(self.project_root / "m365_backup_coordinator.exe")]
            if getattr(sys, "frozen", False)
            else [sys.executable, str(self.project_root / "m365_backup_coordinator.py")]
        )
        command += ["--parent-pid", str(os.getpid())]
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        try:
            self.coordinator_process = subprocess.Popen(
                command,
                cwd=str(self.project_root),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                creationflags=flags,
                close_fds=(os.name != "nt")
            )
            for _ in range(30):
                time.sleep(0.1)
                # Mantém a UI respondendo durante a espera pelo coordenador
                # (até 3s): sem isso, o Windows chega a marcar a janela como
                # "Não está respondendo" logo na abertura do app ou sempre
                # que o coordenador precisa ser reiniciado no meio da sessão.
                application = QApplication.instance()
                if application is not None:
                    application.processEvents()
                try:
                    health = self.coordinator_request("GET", "/health", timeout=0.4) or {}
                    self.coordinator_pid = safe_int(health.get("coordinator_pid"))
                    self.coordinator_online = True
                    return True
                except Exception:
                    continue
        except Exception as error:
            self.log_event("Não foi possível iniciar o coordenador local.", "Erro", str(error))
        self.coordinator_online = False
        return False

    def refresh_from_coordinator(self):
        if self._coordinator_poll_running:
            return
        self._coordinator_poll_running = True
        thread = QThread(self)
        worker = CoordinatorPollWorker(self.coordinator_url, self.last_event_id)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.completed.connect(self._coordinator_poll_completed, Qt.ConnectionType.QueuedConnection)
        worker.failed.connect(self._coordinator_poll_failed, Qt.ConnectionType.QueuedConnection)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._coordinator_poll_finished)
        self._coordinator_poll_thread = thread
        self._coordinator_poll_worker = worker
        thread.start()

    def _coordinator_poll_finished(self):
        self._coordinator_poll_running = False
        self._coordinator_poll_thread = None
        self._coordinator_poll_worker = None

    def _coordinator_poll_failed(self, error):
        if self.coordinator_online:
            self.coordinator_online = False
            self.log_event(
                "Conexão com o coordenador local perdida; tentando reconectar.",
                "Aviso", error
            )

    def _coordinator_poll_completed(self, operations, events):
        was_offline = not self.coordinator_online
        self.coordinator_online = True
        if was_offline:
            self.log_event("Conexão com o coordenador restabelecida.")
        signature = tuple(
            (
                item.get("operation_id"), item.get("status"),
                safe_int(item.get("current_items")), safe_int(item.get("total_items")),
                safe_int(item.get("failed_items")), safe_int(item.get("downloaded_bytes")),
                item.get("current_folder"), safe_int(item.get("current_page")),
                safe_int(item.get("pst_pending_verifications")),
                safe_int(item.get("pst_verified_items")),
                safe_int(item.get("pst_saved_items")),
                safe_int(item.get("pst_audit_failures"))
            )
            for item in operations
        )
        if signature != self._last_operations_signature:
            self._last_operations_signature = signature
            self.apply_coordinator_snapshot(operations or [])
        for event in events or []:
            self.last_event_id = max(
                self.last_event_id, safe_int(event.get("event_id"))
            )
            if not event.get("friendly_message"):
                continue
            level = {"warning": "Aviso", "error": "Erro"}.get(
                event.get("severity"), "Informação"
            )
            self.log_event(
                event["friendly_message"], level,
                event.get("technical_details")
            )
            operation_id = event.get("operation_id")
            mailbox = next((
                name for name, job in self.backup_jobs.items()
                if job.get("operation_id") == operation_id
            ), None)
            if mailbox:
                details = event.get("technical_details")
                text = event["friendly_message"]
                if details:
                    text += f" | {details}"
                self.append_mailbox_log(mailbox, text, level)
                event_type = event.get("event_type") or ""
                stages = {
                    "operation_queued": "Retomada enfileirada",
                    "operation_started": "Carregando checkpoint",
                    "operation_resume_stage": (event.get("payload") or {}).get("stage"),
                    "operation_progress": "Baixando EML novamente",
                    "operation_pausing": "Salvando ponto de pausa",
                    "operation_paused": "Pausado no checkpoint"
                }
                stage = stages.get(event_type)
                if stage:
                    self.backup_jobs[mailbox]["resume_stage"] = stage
                    self.schedule_backup_render()

    def schedule_backup_render(self, delay=250):
        if self._backup_render_pending:
            return
        self._backup_render_pending = True
        QTimer.singleShot(delay, self._flush_backup_render)

    def _flush_backup_render(self):
        self._backup_render_pending = False
        self.render_backup_table()

    def schedule_pst_render(self, delay=250):
        if self._pst_render_pending:
            return
        self._pst_render_pending = True
        QTimer.singleShot(delay, self._flush_pst_render)

    def _flush_pst_render(self):
        self._pst_render_pending = False
        self.render_pst_table()

    def schedule_state_save(self, delay=1200):
        if self._state_save_pending:
            return
        self._state_save_pending = True
        QTimer.singleShot(delay, self._flush_state_save)

    def _flush_state_save(self):
        self._state_save_pending = False
        self.save_state()

    def update_backup_speed(self, mailbox, new_current, timestamp=None):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return
        now = float(timestamp or time.time())
        new_current = safe_int(new_current)
        previous_current = safe_int(job.get("speed_last_current"), new_current)
        previous_time = float(job.get("speed_last_time") or now)

        if new_current < previous_current:
            job["speed_last_current"] = new_current
            job["speed_last_time"] = now
            job["eml_per_second"] = 0.0
            job["speed_samples"] = 0
            return

        delta_items = new_current - previous_current
        delta_seconds = now - previous_time
        if delta_items > 0 and 0.25 <= delta_seconds <= 180:
            instant_rate = delta_items / delta_seconds
            current_rate = float(job.get("eml_per_second") or 0.0)
            samples = safe_int(job.get("speed_samples"))
            alpha = 0.30 if samples < 4 else 0.18
            smoothed_rate = (
                instant_rate
                if current_rate <= 0
                else (alpha * instant_rate) + ((1.0 - alpha) * current_rate)
            )
            job["eml_per_second"] = max(smoothed_rate, 0.0)
            job["speed_samples"] = samples + 1
            job["speed_last_current"] = new_current
            job["speed_last_time"] = now
        elif delta_seconds > 180:
            job["speed_last_current"] = new_current
            job["speed_last_time"] = now

    def reset_backup_speed_baseline(self, mailbox):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return
        job["speed_last_current"] = safe_int(job.get("current"))
        job["speed_last_time"] = time.time()
        if job.get("status") not in STATUS_RUNNING:
            job["eml_per_second"] = 0.0
            job["speed_samples"] = 0

    def format_duration(self, total_seconds):
        seconds = max(0, safe_int(total_seconds))
        days, seconds = divmod(seconds, 86400)
        hours, seconds = divmod(seconds, 3600)
        minutes, seconds = divmod(seconds, 60)
        if days:
            return f"{days}d {hours:02d}h {minutes:02d}min"
        if hours:
            return f"{hours}h {minutes:02d}min"
        if minutes:
            return f"{minutes}min {seconds:02d}s"
        return f"{seconds}s"

    def backup_eta_text(self, job, total, current):
        status = str(job.get("status") or "")
        remaining = max(safe_int(total) - safe_int(current), 0)
        if status == "concluído" or (total and remaining == 0):
            return "Concluído"
        if status in {"pausado", "interrompido", "erro", "incompleto", "cancelado"}:
            return "Pausado" if status == "pausado" else "Indisponível"
        if status not in STATUS_RUNNING and status not in {"iniciando", "continuando"}:
            return "Aguardando início"
        rate = float(job.get("eml_per_second") or 0.0)
        samples = safe_int(job.get("speed_samples"))
        if not total or rate <= 0 or samples < 2:
            return "Calculando..."
        eta_seconds = remaining / rate
        speed_per_minute = rate * 60.0
        return f"{self.format_duration(eta_seconds)} · {speed_per_minute:.1f} EML/min"

    def apply_coordinator_snapshot(self, operations):
        backup_order = []
        seen_backup = set()
        pst_jobs = {}
        for operation in operations:
            operation_id = operation["operation_id"]
            if operation["operation_type"] == "backup":
                mailbox = operation.get("mailbox") or operation_id
                backup_order.append(mailbox)
                seen_backup.add(mailbox)
                existing = self.backup_jobs.get(mailbox, {})
                reported_current = max(
                    safe_int(operation.get("current_items")),
                    safe_int(existing.get("current")),
                )
                incoming_status = self.display_status(operation.get("status"))
                if (
                    existing.get("resume_path")
                    and incoming_status in {"iniciando", "executando", "continuando"}
                    and existing.get("status") not in {"iniciando", "executando", "continuando"}
                ):
                    existing["resume_baseline"] = safe_int(existing.get("current"))
                    existing["resume_stage"] = "Carregando checkpoint"
                if existing:
                    self.update_backup_speed(mailbox, reported_current)
                    existing = self.backup_jobs.get(mailbox, existing)
                self.backup_jobs[mailbox] = {
                    **existing,
                    "mailbox": mailbox,
                    "operation_id": operation_id,
                    "status": incoming_status,
                    "current": reported_current,
                    "expected": max(
                        safe_int(operation.get("total_items")),
                        safe_int(existing.get("expected")),
                        reported_current,
                    ),
                    "downloaded_bytes": safe_int(operation.get("downloaded_bytes")),
                    "rate_limiter_profile": operation.get("rate_limiter_profile"),
                    "rate_limiter_wait_seconds": float(operation.get("rate_limiter_wait_seconds", 0) or 0),
                    "rate_limiter_wait_events": safe_int(operation.get("rate_limiter_wait_events")),
                    "mime_rate_second": float(operation.get("mime_rate_second", 0) or 0),
                    "mime_concurrency": safe_int(operation.get("mime_concurrency")),
                    "resume_path": operation.get("backup_path") or existing.get("resume_path"),
                    "options": operation.get("options") or existing.get("options") or self.default_job_options(),
                }
            else:
                pst_jobs[operation_id] = {
                    "operation_id": operation_id,
                    "backup": operation.get("source_path") or "",
                    "pst": operation.get("destination_path") or "",
                    "status": self.display_status(operation.get("status")),
                    "current": safe_int(operation.get("current_items")),
                    "expected": safe_int(operation.get("total_items")),
                    "failed": safe_int(operation.get("failed_items")),
                    "options": operation.get("options") or {},
                    "stage": operation.get("current_folder") or "",
                    "verification_mode": operation.get("pst_verification_mode") or (operation.get("options") or {}).get("verification_level", "balanced"),
                    "verification_pending": safe_int(operation.get("pst_pending_verifications")),
                    "verification_verified": safe_int(operation.get("pst_verified_items")),
                    "verification_attempts": safe_int(operation.get("pst_verification_attempts")),
                    "audit_failures": safe_int(operation.get("pst_audit_failures")),
                    "performance_profile": operation.get("pst_performance_profile") or "balanced",
                    "bottleneck": operation.get("pst_bottleneck") or "calculando",
                    "effective_workers": safe_int(operation.get("pst_effective_workers")),
                    "effective_queue_limit": safe_int(operation.get("pst_effective_queue_limit")),
                    "peak_rss_bytes": safe_int(operation.get("pst_peak_rss_bytes")),
                    "eta_seconds": float(operation.get("pst_eta_seconds", 0) or 0),
                    "memory_pressure": bool(operation.get("pst_memory_pressure", False)),
                    "resume_first_committed_seconds": float(operation.get("pst_resume_first_committed_seconds", 0) or 0),
                    "resume_first_source_position": safe_int(operation.get("pst_resume_first_source_position")),
                    "resume_first_commit_target_seconds": float(operation.get("pst_resume_first_commit_target_seconds", 10) or 10),
                    "resume_first_commit_target_met": bool(operation.get("pst_resume_first_commit_target_met", False)),
                }
        local_only = [m for m in self.backup_order if m not in seen_backup and not self.backup_jobs.get(m, {}).get("operation_id")]
        self.backup_order = backup_order + local_only
        self.pst_jobs = pst_jobs
        self.schedule_backup_render()
        self.schedule_pst_render()

    def display_status(self, status):
        return {
            "pending": "pendente", "queued": "pendente", "starting": "iniciando", "running": "executando",
            "pause_requested": "solicitando pausa", "pausing": "pausando",
            "paused": "pausado", "interrupted": "interrompido", "completed": "concluído",
            "incomplete": "incompleto", "failed": "erro",
            "cancel_requested": "cancelando", "cancelled": "cancelado"
        }.get(status or "", status or "")

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(245)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(18, 24, 18, 18)
        brand = QLabel("M365 Mailbox\nBackup")
        brand.setObjectName("brand")
        side.addWidget(brand)
        subtitle = QLabel("Centro de operações")
        subtitle.setObjectName("sidebarMuted")
        side.addWidget(subtitle)
        side.addSpacing(22)

        self.navigation = QListWidget()
        self.navigation.setObjectName("navigation")
        for title in (
            "Backups", "Conversões PST", "Logs", "Configurações"
        ):
            QListWidgetItem(title, self.navigation)
        self.navigation.setCurrentRow(0)
        self.navigation.currentRowChanged.connect(self.change_page)
        side.addWidget(self.navigation, 1)
        version = QLabel("Interface PySide6")
        version.setObjectName("sidebarMuted")
        side.addWidget(version)
        layout.addWidget(sidebar)

        content = QFrame()
        content.setObjectName("content")
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(28, 22, 28, 22)
        self.top_title = QLabel("Backups")
        self.top_title.setObjectName("pageTitle")
        self.top_subtitle = QLabel(
            "Acompanhe a fila, a retomada, o download de EML e os logs por mailbox."
        )
        self.top_subtitle.setObjectName("muted")
        content_layout.addWidget(self.top_title)
        content_layout.addWidget(self.top_subtitle)
        content_layout.addSpacing(12)
        self.stack = QStackedWidget()
        content_layout.addWidget(self.stack, 1)
        layout.addWidget(content, 1)

        self.stack.addWidget(self.build_backup_page())
        self.stack.addWidget(self.build_pst_page())
        self.stack.addWidget(self.build_events_page())
        self.stack.addWidget(self.build_settings_page())

    def setup_standard_shortcuts(self):
        self._shortcuts = []

        def register(sequence, callback):
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.ShortcutContext.ApplicationShortcut)
            shortcut.activated.connect(callback)
            self._shortcuts.append(shortcut)

        register(QKeySequence.StandardKey.SelectAll, self.select_all_in_focused_widget)
        register(QKeySequence.StandardKey.Copy, self.copy_from_focused_widget)
        register(QKeySequence.StandardKey.Find, self.focus_current_filter)
        register("Ctrl+N", self.quick_add_mailbox)
        register("Ctrl+I", self.import_csv)
        register("Ctrl+O", self.add_resume_backup)
        register("Ctrl+Return", self.start_queue)
        register("Ctrl+P", self.pause_current_selection)
        register("Ctrl+R", self.resume_current_selection)
        register("Delete", self.remove_current_selection)
        register("F5", self.refresh_current_page)
        register("Ctrl+,", lambda: self.navigation.setCurrentRow(3))
        register("Ctrl+L", lambda: self.navigation.setCurrentRow(2))
        register("Escape", self.clear_current_filter)
        for index in range(4):
            register(
                f"Ctrl+{index + 1}",
                lambda page=index: self.navigation.setCurrentRow(page)
            )

    def select_all_in_focused_widget(self):
        widget = QApplication.focusWidget()
        if isinstance(widget, (QLineEdit, QTextEdit)):
            widget.selectAll()
        elif isinstance(widget, (QTableWidget, QTreeWidget, QListWidget)):
            widget.setFocus(Qt.FocusReason.ShortcutFocusReason)
            widget.selectAll()

    def copy_from_focused_widget(self):
        widget = QApplication.focusWidget()
        if isinstance(widget, (QLineEdit, QTextEdit)):
            widget.copy()
        elif isinstance(widget, QTableWidget):
            indexes = widget.selectedIndexes()
            if not indexes:
                return
            rows = sorted({index.row() for index in indexes})
            columns = sorted({index.column() for index in indexes})
            content = []
            for row in rows:
                row_values = []
                for column in columns:
                    cell = widget.item(row, column)
                    row_values.append(cell.text() if cell is not None else "")
                content.append("\t".join(row_values))
            QApplication.clipboard().setText("\n".join(content))
        elif isinstance(widget, QTreeWidget):
            content = ["\t".join(
                item.text(column) for column in range(widget.columnCount())
            ) for item in widget.selectedItems()]
            QApplication.clipboard().setText("\n".join(content))
        elif isinstance(widget, QListWidget):
            QApplication.clipboard().setText("\n".join(
                item.text() for item in widget.selectedItems()
            ))

    def focus_current_filter(self):
        target = {
            0: getattr(self, "backup_filter", None),
            1: getattr(self, "pst_filter", None),
            2: getattr(self, "log_search_edit", None)
        }.get(self.stack.currentIndex())
        if target is not None:
            target.setFocus()
            target.selectAll()

    def clear_current_filter(self):
        page = self.stack.currentIndex()
        if page == 0:
            self.clear_backup_filters()
        elif page == 1 and hasattr(self, "pst_filter"):
            self.pst_filter.clear()
            self.pst_status_filter.setCurrentIndex(0)
        elif page == 2 and hasattr(self, "log_search_edit"):
            self.log_search_edit.clear()
            self.log_level_combo.setCurrentIndex(0)

    def pause_current_selection(self):
        if self.stack.currentIndex() == 0:
            self.pause_selected_backup()
        elif self.stack.currentIndex() == 1:
            self.pause_selected_pst()

    def resume_current_selection(self):
        if self.stack.currentIndex() == 0:
            self.resume_selected_backup()
        elif self.stack.currentIndex() == 1:
            self.resume_selected_pst()

    def remove_current_selection(self):
        widget = QApplication.focusWidget()
        if isinstance(widget, (QLineEdit, QTextEdit)):
            return
        if self.stack.currentIndex() == 0:
            self.remove_selected_backup()
        elif self.stack.currentIndex() == 1:
            self.remove_selected_pst()

    def refresh_current_page(self):
        page = self.stack.currentIndex()
        if page == 0:
            self.refresh_from_coordinator()
            self.render_backup_table()
        elif page == 1:
            self.refresh_from_coordinator()
            self.render_pst_table()
        elif page == 2:
            self.reload_selected_log()
        else:
            self.load_runtime_settings_into_ui()

    def build_dashboard_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        cards = QGridLayout()
        self.dashboard_cards = {
            "running": StatCard("Backups em execução", "0"),
            "paused": StatCard("Operações pausadas", "0"),
            "pst": StatCard("Conversões PST", "0"),
            "health": StatCard("Microsoft Graph", "Sem dados")
        }
        for index, card in enumerate(self.dashboard_cards.values()):
            cards.addWidget(card, 0, index)
        root.addLayout(cards)

        quick = QGroupBox("Ações rápidas")
        quick_layout = QHBoxLayout(quick)
        for text, callback in (
            ("Adicionar mailbox", self.quick_add_mailbox),
            ("Importar CSV", self.import_csv),
            ("Continuar backup", self.add_resume_backup),
            ("Nova conversão PST", self.new_pst_job)
        ):
            button = QPushButton(text)
            button.clicked.connect(callback)
            quick_layout.addWidget(button)
        quick_layout.addStretch()
        root.addWidget(quick)

        activity = QGroupBox("Atividade recente")
        activity_layout = QVBoxLayout(activity)
        self.dashboard_activity = QTextEdit()
        self.dashboard_activity.setReadOnly(True)
        self.dashboard_activity.setMaximumHeight(270)
        activity_layout.addWidget(self.dashboard_activity)
        root.addWidget(activity)
        root.addStretch()
        return page

    def build_backup_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        toolbar = QHBoxLayout()
        for text, callback, primary in (
            ("Adicionar", self.quick_add_mailbox, True),
            ("Importar CSV", self.import_csv, False),
            ("Continuar existente", self.add_resume_backup, False),
            ("Configurar pastas", self.configure_selected_backup, False),
            ("Iniciar backups", self.start_queue, True),
            ("Pausar", self.pause_selected_backup, False),
            ("Retomar", self.resume_selected_backup, False),
            ("Corrigir falhas", self.repair_selected_backup, True),
            ("Remover", self.remove_selected_backup, False)
        ):
            button = QPushButton(text)
            if primary:
                button.setObjectName("primaryButton")
            button.clicked.connect(callback)
            toolbar.addWidget(button)
        root.addLayout(toolbar)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Status:"))
        self.backup_status_filter = QComboBox()
        for label, value in (
            ("Todos", ""), ("Pendente", "pendente"),
            ("Iniciando", "iniciando"), ("Executando", "executando"),
            ("Pausado", "pausado"), ("Interrompido", "interrompido"),
            ("Incompleto", "incompleto"), ("Concluído", "concluído"),
            ("Erro", "erro"), ("Cancelado", "cancelado")
        ):
            self.backup_status_filter.addItem(label, value)
        self.backup_status_filter.currentIndexChanged.connect(
            self.filter_backup_table
        )
        filters.addWidget(self.backup_status_filter)

        self.backup_filter = QLineEdit()
        self.backup_filter.setPlaceholderText(
            "Buscar mailbox, status, pasta ou caminho..."
        )
        self.backup_filter.setClearButtonEnabled(True)
        self.backup_filter.textChanged.connect(self.filter_backup_table)
        filters.addWidget(self.backup_filter, 1)

        clear_filters = QPushButton("Limpar filtros")
        clear_filters.clicked.connect(self.clear_backup_filters)
        filters.addWidget(clear_filters)
        self.backup_filter_counter = QLabel("0 exibidos")
        self.backup_filter_counter.setObjectName("muted")
        filters.addWidget(self.backup_filter_counter)
        root.addLayout(filters)

        columns = (
            "Mailbox", "Status", "Etapa atual", "E-mails na caixa", "Análise",
            "Progresso", "Faltam", "Tamanho", "Tempo estimado",
            "Pastas", "Configuração", "Ordem"
        )
        self.backup_table = DraggableBackupTable(0, len(columns), self)
        self.backup_table.setHorizontalHeaderLabels(columns)
        self.backup_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in range(1, len(columns)):
            self.backup_table.horizontalHeader().setSectionResizeMode(
                column, QHeaderView.ResizeMode.ResizeToContents
            )
        self.backup_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.backup_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.backup_table.setAlternatingRowColors(True)
        self.backup_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.backup_table.customContextMenuRequested.connect(self.backup_context_menu)
        self.backup_table.orderChanged.connect(self.apply_dragged_backup_order)
        self.move_up_action = QAction("Subir na fila", self)
        self.move_up_action.setShortcut("Alt+Up")
        self.move_up_action.triggered.connect(self.move_selected_backup_up)
        self.addAction(self.move_up_action)
        self.move_down_action = QAction("Descer na fila", self)
        self.move_down_action.setShortcut("Alt+Down")
        self.move_down_action.triggered.connect(self.move_selected_backup_down)
        self.addAction(self.move_down_action)
        backup_splitter = QSplitter(Qt.Orientation.Vertical)
        backup_splitter.addWidget(self.backup_table)

        mailbox_log_group = QGroupBox("Log da mailbox selecionada")
        mailbox_log_layout = QVBoxLayout(mailbox_log_group)
        mailbox_log_toolbar = QHBoxLayout()
        self.mailbox_log_title = QLabel("Selecione uma mailbox para acompanhar a atividade.")
        self.mailbox_log_title.setObjectName("muted")
        mailbox_log_toolbar.addWidget(self.mailbox_log_title, 1)
        clear_mailbox_log = QPushButton("Limpar visualização")
        clear_mailbox_log.clicked.connect(self.clear_selected_mailbox_log)
        mailbox_log_toolbar.addWidget(clear_mailbox_log)
        mailbox_log_layout.addLayout(mailbox_log_toolbar)
        self.mailbox_log_text = QTextEdit()
        self.mailbox_log_text.setReadOnly(True)
        self.mailbox_log_text.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.mailbox_log_text.setMaximumHeight(230)
        mailbox_log_layout.addWidget(self.mailbox_log_text)
        backup_splitter.addWidget(mailbox_log_group)
        backup_splitter.setSizes([620, 210])
        root.addWidget(backup_splitter, 1)
        self.backup_table.itemSelectionChanged.connect(self.refresh_selected_mailbox_log)
        return page

    def build_pst_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        toolbar = QHBoxLayout()
        for text, callback, primary in (
            ("Nova conversão", self.new_pst_job, True),
            ("Iniciar selecionadas", self.start_selected_pst, True),
            ("Pausar", self.pause_selected_pst, False),
            ("Retomar", self.resume_selected_pst, False),
            ("Remover", self.remove_selected_pst, False),
            ("Abrir destino", lambda: open_local_path(self.pst_output_edit.text()), False)
        ):
            button = QPushButton(text)
            if primary:
                button.setObjectName("primaryButton")
            button.clicked.connect(callback)
            toolbar.addWidget(button)
        toolbar.addStretch()
        root.addLayout(toolbar)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Status:"))
        self.pst_status_filter = QComboBox()
        for label, value in (
            ("Todos", ""), ("Pendente", "pendente"),
            ("Iniciando", "iniciando"), ("Executando", "executando"),
            ("Pausado", "pausado"),
            ("Concluído", "concluído"), ("Erro", "erro")
        ):
            self.pst_status_filter.addItem(label, value)
        self.pst_status_filter.currentIndexChanged.connect(self.render_pst_table)
        filters.addWidget(self.pst_status_filter)
        self.pst_filter = QLineEdit()
        self.pst_filter.setPlaceholderText("Buscar origem, destino, ID ou status...")
        self.pst_filter.setClearButtonEnabled(True)
        self.pst_filter.textChanged.connect(self.render_pst_table)
        filters.addWidget(self.pst_filter, 1)
        root.addLayout(filters)

        columns = ("ID", "Nome", "Backup de origem", "PST de destino", "Nome no Outlook", "Política", "Status", "Progresso", "Verificação", "Falhas", "Faltam")
        self.pst_table = QTableWidget(0, len(columns))
        self.pst_table.setHorizontalHeaderLabels(columns)
        self.pst_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.pst_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        for column in (0, 1, 4, 5, 6, 7, 8, 9, 10):
            self.pst_table.horizontalHeader().setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        self.pst_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.pst_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.pst_table.setAlternatingRowColors(True)
        root.addWidget(self.pst_table, 1)
        return page

    def build_performance_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("Período:"))
        self.metrics_period = QComboBox()
        self.metrics_period.addItem("Últimos 5 minutos", 5)
        self.metrics_period.addItem("Últimos 15 minutos", 15)
        self.metrics_period.addItem("Últimos 60 minutos", 60)
        self.metrics_period.addItem("Últimas 24 horas", 1440)
        self.metrics_period.setCurrentIndex(2)
        self.metrics_period.currentIndexChanged.connect(self.refresh_metrics)
        toolbar.addWidget(self.metrics_period)
        refresh = QPushButton("Atualizar")
        refresh.clicked.connect(self.refresh_metrics)
        export = QPushButton("Exportar relatório")
        export.clicked.connect(self.export_metrics)
        toolbar.addWidget(refresh)
        toolbar.addWidget(export)
        toolbar.addStretch()
        root.addLayout(toolbar)

        cards = QGridLayout()
        self.metric_cards = {
            "health": StatCard("Saúde da API"),
            "requests": StatCard("Requisições"),
            "latency": StatCard("Latência média / P95"),
            "rate": StatCard("Taxa média"),
            "throttle": StatCard("Throttling / espera")
        }
        for index, card in enumerate(self.metric_cards.values()):
            cards.addWidget(card, 0, index)
        root.addLayout(cards)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        category_group = QGroupBox("Consumo por tipo de chamada")
        category_layout = QVBoxLayout(category_group)
        self.category_table = self.metrics_table()
        category_layout.addWidget(self.category_table)
        mailbox_group = QGroupBox("Consumo por mailbox")
        mailbox_layout = QVBoxLayout(mailbox_group)
        self.mailbox_metrics_table = self.metrics_table()
        mailbox_layout.addWidget(self.mailbox_metrics_table)
        splitter.addWidget(category_group)
        splitter.addWidget(mailbox_group)
        root.addWidget(splitter, 1)
        self.recommendation_label = QLabel("Aguardando dados de utilização.")
        self.recommendation_label.setWordWrap(True)
        self.recommendation_label.setObjectName("infoBanner")
        root.addWidget(self.recommendation_label)
        return page

    def metrics_table(self):
        columns = ("Categoria", "Chamadas", "Falhas", "Limitações", "Latência", "Espera")
        table = QTableWidget(0, len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for column in range(1, len(columns)):
            table.horizontalHeader().setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        return table

    def build_events_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("Arquivo:"))
        self.log_file_combo = QComboBox()
        self.log_file_combo.addItem("Log operacional", "m365_mailbox_backup.log")
        self.log_file_combo.addItem("Falhas da aplicação", "application_crash.log")
        self.log_file_combo.currentIndexChanged.connect(self.change_log_file)
        toolbar.addWidget(self.log_file_combo)
        toolbar.addWidget(QLabel("Nível:"))
        self.log_level_combo = QComboBox()
        for label, value in (
            ("Todos", ""), ("DEBUG", "DEBUG"), ("INFO", "INFO"),
            ("WARNING", "WARNING"), ("ERROR", "ERROR"),
            ("CRITICAL", "CRITICAL")
        ):
            self.log_level_combo.addItem(label, value)
        self.log_level_combo.currentIndexChanged.connect(self.rebuild_log_view)
        toolbar.addWidget(self.log_level_combo)
        self.log_search_edit = QLineEdit()
        self.log_search_edit.setPlaceholderText("Pesquisar no log...")
        self.log_search_edit.setClearButtonEnabled(True)
        self.log_search_edit.textChanged.connect(self.rebuild_log_view)
        toolbar.addWidget(self.log_search_edit, 1)
        self.log_pause_button = QPushButton("Pausar")
        self.log_pause_button.setCheckable(True)
        self.log_pause_button.toggled.connect(self.toggle_log_pause)
        toolbar.addWidget(self.log_pause_button)
        reload_button = QPushButton("Recarregar")
        reload_button.clicked.connect(self.reload_selected_log)
        toolbar.addWidget(reload_button)
        clear_button = QPushButton("Limpar tela")
        clear_button.clicked.connect(self.clear_log_view)
        toolbar.addWidget(clear_button)
        open_logs = QPushButton("Abrir pasta")
        open_logs.clicked.connect(
            lambda: open_local_path(self.project_root / "logs")
        )
        toolbar.addWidget(open_logs)
        self.log_auto_scroll = QCheckBox("Rolagem automática")
        self.log_auto_scroll.setChecked(True)
        toolbar.addWidget(self.log_auto_scroll)
        root.addLayout(toolbar)

        status_row = QHBoxLayout()
        self.log_file_status = QLabel("Aguardando leitura do arquivo...")
        self.log_file_status.setObjectName("muted")
        status_row.addWidget(self.log_file_status)
        status_row.addStretch()
        self.log_line_counter = QLabel("0 linhas visíveis")
        self.log_line_counter.setObjectName("muted")
        status_row.addWidget(self.log_line_counter)
        root.addLayout(status_row)

        self.events_text = QTextEdit()
        self.events_text.setReadOnly(True)
        self.events_text.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.events_text.setPlaceholderText(
            "As linhas do arquivo de log aparecerão aqui."
        )
        log_font = QFont("Consolas")
        log_font.setStyleHint(QFont.StyleHint.Monospace)
        self.events_text.setFont(log_font)
        root.addWidget(self.events_text, 1)

        self.log_file_position = 0
        self.log_file_identity = None
        self.log_pending_fragment = ""
        self.log_raw_lines = []
        self.log_view_paused = False
        self.max_log_lines = 5000
        QTimer.singleShot(0, self.initialize_log_viewer)
        return page

    def initialize_log_viewer(self):
        self.logs_directory = self.project_root / "logs"
        self.logs_directory.mkdir(parents=True, exist_ok=True)
        self.log_timer = QTimer(self)
        self.log_timer.timeout.connect(self.read_new_log_lines)
        self.log_timer.start(750)
        self.reload_selected_log()

    def selected_log_path(self):
        filename = self.log_file_combo.currentData() or "m365_mailbox_backup.log"
        return self.project_root / "logs" / filename

    def change_log_file(self, *args):
        self.reload_selected_log()

    def reload_selected_log(self):
        self.log_file_position = 0
        self.log_file_identity = None
        self.log_pending_fragment = ""
        self.log_raw_lines = []
        self.events_text.clear()
        self.load_log_tail()

    def load_log_tail(self):
        path = self.selected_log_path()
        if not path.exists():
            self.log_file_status.setText(f"Arquivo ainda não criado: {path.name}")
            self.update_log_line_counter()
            return
        try:
            file_size = path.stat().st_size
            start_position = max(0, file_size - 2 * 1024 * 1024)
            with path.open("rb") as log_file:
                log_file.seek(start_position)
                if start_position > 0:
                    log_file.readline()
                content = log_file.read()
                self.log_file_position = log_file.tell()
            lines = content.decode("utf-8", errors="replace").splitlines()
            self.log_raw_lines = lines[-self.max_log_lines:]
            self.log_file_identity = self.get_log_file_identity(path)
            self.rebuild_log_view()
            self.log_file_status.setText(
                f"{path.name} · acompanhando em tempo real"
            )
        except Exception as error:
            self.log_file_status.setText(f"Falha ao abrir {path.name}: {error}")

    def get_log_file_identity(self, path):
        try:
            stat = path.stat()
            return (
                getattr(stat, "st_dev", None),
                getattr(stat, "st_ino", None),
                stat.st_ctime_ns
            )
        except Exception:
            return None

    def read_new_log_lines(self):
        if self.log_view_paused:
            return
        path = self.selected_log_path()
        if not path.exists():
            self.log_file_status.setText(f"Arquivo ainda não criado: {path.name}")
            return
        try:
            current_size = path.stat().st_size
            current_identity = self.get_log_file_identity(path)
            replaced = (
                self.log_file_identity is not None
                and current_identity != self.log_file_identity
            )
            truncated = current_size < self.log_file_position
            if replaced or truncated:
                self.log_file_position = 0
                self.log_pending_fragment = ""
                reason = "rotação detectada" if replaced else "arquivo reiniciado"
                self.append_log_lines([f"--- {reason}: {path.name} ---"])
            if current_size == self.log_file_position:
                self.log_file_identity = current_identity
                return
            with path.open("rb") as log_file:
                log_file.seek(self.log_file_position)
                content = log_file.read()
                self.log_file_position = log_file.tell()
            text = self.log_pending_fragment + content.decode(
                "utf-8", errors="replace"
            )
            if text.endswith(("\n", "\r")):
                self.log_pending_fragment = ""
                lines = text.splitlines()
            else:
                parts = text.splitlines()
                self.log_pending_fragment = parts[-1] if parts else text
                lines = parts[:-1] if parts else []
            self.log_file_identity = current_identity
            self.append_log_lines(lines)
            self.log_file_status.setText(
                f"{path.name} · acompanhando em tempo real"
            )
        except PermissionError:
            self.log_file_status.setText(f"{path.name} está temporariamente ocupado.")
        except Exception as error:
            self.log_file_status.setText(f"Erro durante a leitura: {error}")

    def append_log_lines(self, lines):
        if not lines:
            return
        self.log_raw_lines.extend(lines)
        if len(self.log_raw_lines) > self.max_log_lines:
            self.log_raw_lines = self.log_raw_lines[-self.max_log_lines:]
            self.rebuild_log_view()
            return
        visible = [line for line in lines if self.log_line_matches_filters(line)]
        for line in visible:
            self.append_colored_log_line(line)
        if self.log_auto_scroll.isChecked():
            scrollbar = self.events_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        self.update_log_line_counter()

    def append_colored_log_line(self, line):
        escaped = (
            str(line).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        upper = str(line).upper()
        if " | CRITICAL | " in upper:
            color, weight = "#dc2626", "700"
        elif " | ERROR | " in upper or "TRACEBACK" in upper:
            color, weight = "#ef4444", "600"
        elif " | WARNING | " in upper:
            color, weight = "#f59e0b", "600"
        elif " | DEBUG | " in upper:
            color, weight = "#94a3b8", "400"
        elif " | INFO | " in upper:
            color, weight = "#38bdf8", "400"
        else:
            color, weight = "inherit", "400"
        self.events_text.append(
            f'<span style="color:{color}; font-weight:{weight};">{escaped}</span>'
        )

    def log_line_matches_filters(self, line):
        level = self.log_level_combo.currentData() or ""
        search = self.log_search_edit.text().strip().lower()
        upper = str(line).upper()
        if level and f" | {level} | " not in upper:
            return False
        return not search or search in str(line).lower()

    def rebuild_log_view(self, *args):
        self.events_text.setUpdatesEnabled(False)
        try:
            self.events_text.clear()
            for line in self.log_raw_lines:
                if self.log_line_matches_filters(line):
                    self.append_colored_log_line(line)
        finally:
            self.events_text.setUpdatesEnabled(True)
        if self.log_auto_scroll.isChecked():
            scrollbar = self.events_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())
        self.update_log_line_counter()

    def update_log_line_counter(self):
        visible = sum(
            1 for line in self.log_raw_lines
            if self.log_line_matches_filters(line)
        )
        self.log_line_counter.setText(
            f"{visible} visíveis · {len(self.log_raw_lines)} carregadas"
        )

    def toggle_log_pause(self, paused):
        self.log_view_paused = bool(paused)
        self.log_pause_button.setText("Continuar" if paused else "Pausar")
        if paused:
            self.log_file_status.setText(
                "Atualização visual pausada; o arquivo continua sendo gravado."
            )
        else:
            self.read_new_log_lines()

    def clear_log_view(self):
        self.log_raw_lines = []
        self.log_pending_fragment = ""
        self.events_text.clear()
        self.update_log_line_counter()
        self.log_file_status.setText(
            "Visualização limpa; o arquivo original não foi apagado."
        )

    def build_settings_page(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        content = QWidget()
        root = QVBoxLayout(content)
        tabs = QTabWidget()

        general = QWidget(); general_form = QFormLayout(general)
        self.destination_edit = QLineEdit()
        destination_row = QHBoxLayout(); destination_row.addWidget(self.destination_edit)
        choose_destination = QPushButton("Selecionar"); choose_destination.clicked.connect(self.select_destination)
        destination_row.addWidget(choose_destination)
        general_form.addRow("Destino dos backups:", destination_row)
        self.pst_output_edit = QLineEdit(); general_form.addRow("Destino dos PST:", self.pst_output_edit)
        self.parallel_spin = QSpinBox(); self.parallel_spin.setRange(1, 20); general_form.addRow("Backups simultâneos:", self.parallel_spin)
        self.pst_parallel_spin = QSpinBox(); self.pst_parallel_spin.setRange(1, 5); general_form.addRow("Conversões PST simultâneas:", self.pst_parallel_spin)
        self.default_all = QCheckBox("Exportar todas as mensagens"); self.default_all.setChecked(True)
        self.default_skip_calendar = QCheckBox("Ignorar calendário")
        self.default_skip_contacts = QCheckBox("Ignorar contatos")
        self.default_skip_tasks = QCheckBox("Ignorar tarefas")
        self.default_skip_precheck = QCheckBox("Iniciar sem pré-análise completa")
        for widget in (self.default_all, self.default_skip_calendar, self.default_skip_contacts, self.default_skip_tasks, self.default_skip_precheck): general_form.addRow("", widget)
        tabs.addTab(general, "Geral")

        credentials = QWidget(); credentials_form = QFormLayout(credentials)
        current = self.credential_store.load()
        self.tenant_edit = QLineEdit(current.get("tenant_id", ""))
        self.client_edit = QLineEdit(current.get("client_id", ""))
        self.secret_edit = QLineEdit(); self.secret_edit.setEchoMode(QLineEdit.EchoMode.Password); self.secret_edit.setPlaceholderText("Deixe vazio para manter o segredo atual")
        credentials_form.addRow("Tenant ID:", self.tenant_edit); credentials_form.addRow("Client ID:", self.client_edit); credentials_form.addRow("Client Secret:", self.secret_edit)
        test_button = QPushButton("Salvar e testar credenciais"); test_button.clicked.connect(self.save_and_test_credentials)
        credentials_form.addRow("", test_button)
        tabs.addTab(credentials, "Credenciais e Graph")

        performance = QWidget(); performance_form = QFormLayout(performance)
        self.download_workers_spin = QSpinBox(); self.download_workers_spin.setRange(1, 32)
        self.rate_limiter_enabled_check = QCheckBox("Ativar PyrateLimiter e orçamento compartilhado")
        self.rate_limiter_profile_combo = QComboBox()
        self.rate_limiter_profile_combo.addItem("Automático", "automatic")
        self.rate_limiter_profile_combo.addItem("Conservador", "conservative")
        self.rate_limiter_profile_combo.addItem("Desempenho 2x", "performance_2x")
        self.rate_limiter_profile_combo.addItem("Personalizado", "custom")
        self.rate_limiter_profile_combo.currentIndexChanged.connect(
            self.apply_rate_limiter_profile
        )
        self.mailbox_rate_second_spin = QSpinBox(); self.mailbox_rate_second_spin.setRange(1, 50); self.mailbox_rate_second_spin.setSuffix(" MIME/s")
        self.mailbox_rate_minute_spin = QSpinBox(); self.mailbox_rate_minute_spin.setRange(1, 3000); self.mailbox_rate_minute_spin.setSuffix(" MIME/min")
        self.global_rate_second_spin = QSpinBox(); self.global_rate_second_spin.setRange(1, 100); self.global_rate_second_spin.setSuffix(" MIME/s")
        self.global_rate_minute_spin = QSpinBox(); self.global_rate_minute_spin.setRange(1, 6000); self.global_rate_minute_spin.setSuffix(" MIME/min")
        self.mime_concurrency_spin = QSpinBox(); self.mime_concurrency_spin.setRange(1, 16)
        self.resume_mime_concurrency_spin = QSpinBox(); self.resume_mime_concurrency_spin.setRange(1, 8)
        self.page_size_spin = QSpinBox(); self.page_size_spin.setRange(10, 999)
        self.checkpoint_batch_spin = QSpinBox(); self.checkpoint_batch_spin.setRange(1, 1000)
        self.warning_disk_spin = QSpinBox(); self.warning_disk_spin.setRange(1, 10000); self.warning_disk_spin.setSuffix(" GB")
        self.critical_disk_spin = QSpinBox(); self.critical_disk_spin.setRange(1, 10000); self.critical_disk_spin.setSuffix(" GB")
        self.auto_pause_disk = QCheckBox("Pausar automaticamente com pouco espaço")
        performance_form.addRow("Downloads simultâneos:", self.download_workers_spin)
        performance_form.addRow("", self.rate_limiter_enabled_check)
        performance_form.addRow("Perfil do limitador:", self.rate_limiter_profile_combo)
        performance_form.addRow("Taxa por mailbox:", self.mailbox_rate_second_spin)
        performance_form.addRow("Volume por mailbox:", self.mailbox_rate_minute_spin)
        performance_form.addRow("Taxa global:", self.global_rate_second_spin)
        performance_form.addRow("Volume global:", self.global_rate_minute_spin)
        performance_form.addRow("Concorrência MIME:", self.mime_concurrency_spin)
        performance_form.addRow("Concorrência na retomada:", self.resume_mime_concurrency_spin)
        performance_form.addRow("Tamanho da página Graph:", self.page_size_spin)
        performance_form.addRow("Lote de checkpoint:", self.checkpoint_batch_spin)
        performance_form.addRow("Alerta de espaço:", self.warning_disk_spin)
        performance_form.addRow("Limite crítico:", self.critical_disk_spin)
        performance_form.addRow("", self.auto_pause_disk)
        performance_tab_index = tabs.addTab(performance, "Desempenho e armazenamento")
        tabs.setTabVisible(performance_tab_index,False)

        appearance = QWidget(); appearance_form = QFormLayout(appearance)
        self.theme_combo = QComboBox(); self.theme_combo.addItem("Automático", "automatic"); self.theme_combo.addItem("Claro", "light"); self.theme_combo.addItem("Escuro", "dark")
        self.font_spin = QSpinBox(); self.font_spin.setRange(8, 18)
        self.pst_detach = QCheckBox("Remover PST do Outlook ao concluir")
        appearance_form.addRow("Tema:", self.theme_combo); appearance_form.addRow("Tamanho da fonte:", self.font_spin); appearance_form.addRow("", self.pst_detach)
        tabs.addTab(appearance, "Aparência e PST")

        root.addWidget(tabs)
        self.settings_effective_status = QLabel(
            "As alterações salvas serão incorporadas ao ambiente real de cada nova operação."
        )
        self.settings_effective_status.setWordWrap(True)
        self.settings_effective_status.setObjectName("infoBanner")
        root.addWidget(self.settings_effective_status)
        buttons = QHBoxLayout()
        save = QPushButton("Salvar configurações"); save.setObjectName("primaryButton"); save.clicked.connect(self.save_runtime_settings)
        export_profile = QPushButton("Exportar perfil"); export_profile.clicked.connect(self.export_settings_profile)
        import_profile = QPushButton("Importar perfil"); import_profile.clicked.connect(self.import_settings_profile)
        diagnostics = QPushButton("Executar diagnóstico"); diagnostics.clicked.connect(self.show_diagnostics)
        integrity = QPushButton("Validar backup"); integrity.clicked.connect(self.validate_selected_backup)
        for button in (save, export_profile, import_profile, diagnostics, integrity): buttons.addWidget(button)
        buttons.addStretch(); root.addLayout(buttons); root.addStretch()
        scroll.setWidget(content); page_layout.addWidget(scroll)
        return page

    def apply_rate_limiter_profile(self, *args):
        profile = self.rate_limiter_profile_combo.currentData()
        presets = {
            "automatic": (6, 240, 10, 480, 6, 6),
            "conservative": (3, 120, 5, 240, 3, 3),
            "performance_2x": (6, 240, 10, 480, 6, 6),
        }
        values = presets.get(profile)
        if not values:
            return
        widgets = (
            self.mailbox_rate_second_spin,
            self.mailbox_rate_minute_spin,
            self.global_rate_second_spin,
            self.global_rate_minute_spin,
            self.mime_concurrency_spin,
            self.resume_mime_concurrency_spin,
        )
        for widget, value in zip(widgets, values):
            widget.setValue(value)

    def load_runtime_settings_into_ui(self):
        self.destination_edit.setText(str(self.app_settings.resolved_path("backup_root")))
        self.pst_output_edit.setText(str(self.app_settings.resolved_path("pst_root")))
        self.parallel_spin.setValue(safe_int(self.app_settings.get("backup", "parallel", 2), 2))
        self.pst_parallel_spin.setValue(safe_int(self.app_settings.get("pst", "parallel", 2), 2))
        self.download_workers_spin.setValue(safe_int(self.app_settings.get("backup", "download_workers", 24), 24))
        self.rate_limiter_enabled_check.setChecked(bool(self.app_settings.get("backup", "rate_limiter_enabled", True)))
        rate_index = self.rate_limiter_profile_combo.findData(self.app_settings.get("backup", "rate_limiter_profile", "performance_2x")); self.rate_limiter_profile_combo.setCurrentIndex(max(0, rate_index))
        self.mailbox_rate_second_spin.setValue(safe_int(self.app_settings.get("backup", "mailbox_mime_rate_second", 6), 6))
        self.mailbox_rate_minute_spin.setValue(safe_int(self.app_settings.get("backup", "mailbox_mime_rate_minute", 240), 240))
        self.global_rate_second_spin.setValue(safe_int(self.app_settings.get("backup", "global_mime_rate_second", 10), 10))
        self.global_rate_minute_spin.setValue(safe_int(self.app_settings.get("backup", "global_mime_rate_minute", 480), 480))
        self.mime_concurrency_spin.setValue(safe_int(self.app_settings.get("backup", "mime_max_concurrency", 6), 6))
        self.resume_mime_concurrency_spin.setValue(safe_int(self.app_settings.get("backup", "resume_mime_max_concurrency", 6), 6))
        self.page_size_spin.setValue(safe_int(self.app_settings.get("backup", "page_size", 250), 250))
        self.checkpoint_batch_spin.setValue(safe_int(self.app_settings.get("backup", "checkpoint_batch", 50), 50))
        self.warning_disk_spin.setValue(safe_int(self.app_settings.get("storage", "warning_free_gb", 50), 50))
        self.critical_disk_spin.setValue(safe_int(self.app_settings.get("storage", "critical_free_gb", 10), 10))
        self.auto_pause_disk.setChecked(bool(self.app_settings.get("storage", "auto_pause", True)))
        self.font_spin.setValue(safe_int(self.app_settings.get("appearance", "font_size", 10), 10))
        index = self.theme_combo.findData(self.app_settings.get("appearance", "theme", "automatic")); self.theme_combo.setCurrentIndex(max(0, index))
        self.pst_detach.setChecked(bool(self.app_settings.get("pst", "detach_after", False)))
        for widget in (self.default_skip_calendar, self.default_skip_contacts, self.default_skip_tasks, self.default_skip_precheck): widget.setChecked(True)

    def save_runtime_settings(self):
        paths = self.app_settings.data.setdefault("paths", {})
        paths["backup_root"] = self.destination_edit.text().strip(); paths["pst_root"] = self.pst_output_edit.text().strip()
        backup = self.app_settings.data.setdefault("backup", {})
        backup.update({
            "parallel": self.parallel_spin.value(),
            "download_workers": self.download_workers_spin.value(),
            "rate_limiter_enabled": self.rate_limiter_enabled_check.isChecked(),
            "rate_limiter_profile": self.rate_limiter_profile_combo.currentData(),
            "mailbox_mime_rate_second": self.mailbox_rate_second_spin.value(),
            "mailbox_mime_rate_minute": self.mailbox_rate_minute_spin.value(),
            "global_mime_rate_second": self.global_rate_second_spin.value(),
            "global_mime_rate_minute": self.global_rate_minute_spin.value(),
            "mime_max_concurrency": self.mime_concurrency_spin.value(),
            "resume_mime_max_concurrency": self.resume_mime_concurrency_spin.value(),
            "page_size": self.page_size_spin.value(),
            "checkpoint_batch": self.checkpoint_batch_spin.value()
        })
        pst_settings = self.app_settings.data.setdefault("pst", {})
        pst_settings["parallel"] = self.pst_parallel_spin.value()
        pst_settings["detach_after"] = self.pst_detach.isChecked()
        storage = self.app_settings.data.setdefault("storage", {})
        storage.update({"warning_free_gb": self.warning_disk_spin.value(), "critical_free_gb": self.critical_disk_spin.value(), "auto_pause": self.auto_pause_disk.isChecked()})
        appearance = self.app_settings.data.setdefault("appearance", {})
        appearance.update({"theme": self.theme_combo.currentData(), "font_size": self.font_spin.value()})
        errors = self.app_settings.validate_runtime_settings()
        if errors:
            QMessageBox.warning(self, "Configurações inválidas", "\n".join(errors))
            return
        self.app_settings.save()
        applied = self.app_settings.apply_environment()
        self._apply_style()
        coordinator_message = ""
        if self.ensure_coordinator():
            try:
                self.coordinator_request(
                    "PUT", "/settings/concurrency",
                    {"backup_workers": self.parallel_spin.value(),
                     "pst_workers": self.pst_parallel_spin.value()}
                )
                coordinator_message = " Paralelismo atualizado no coordenador."
            except Exception as error:
                coordinator_message = f" O coordenador não aceitou o paralelismo: {error}"
        active = sum(1 for job in list(self.backup_jobs.values()) + list(self.pst_jobs.values())
                     if job.get("status") in {"executando", "iniciando", "continuando", "pausando"})
        message = (
            f"{len(applied)} configurações operacionais foram salvas e serão injetadas "
            "diretamente nos processos de novas operações." + coordinator_message
        )
        if active:
            message += f" {active} operação(ões) em andamento mantêm o snapshot anterior."
        self.settings_effective_status.setText(message)
        QMessageBox.information(self, "Configurações aplicadas", message)

    def save_and_test_credentials(self):
        current = self.credential_store.load(); secret = self.secret_edit.text() or current.get("client_secret", "")
        if not self.tenant_edit.text().strip() or not self.client_edit.text().strip() or not secret:
            QMessageBox.warning(self, "Credenciais", "Preencha Tenant ID, Client ID e Client Secret."); return
        self.credential_store.save(self.tenant_edit.text(), self.client_edit.text(), secret); self.credential_store.apply_environment()
        try:
            app = msal.ConfidentialClientApplication(self.client_edit.text().strip(), client_credential=secret, authority=f"https://login.microsoftonline.com/{self.tenant_edit.text().strip()}")
            raw_result = app.acquire_token_for_client(scopes=[self.app_settings.get("graph", "scope")])
            result = raw_result if isinstance(raw_result, dict) else {}
            if "access_token" not in result: raise RuntimeError(result.get("error_description") or "Falha de autenticação")
            self.secret_edit.clear(); QMessageBox.information(self, "Credenciais", "Credenciais protegidas e autenticação validada.")
        except Exception as error: QMessageBox.critical(self, "Credenciais", str(error))

    def export_settings_profile(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar perfil", str(self.project_root / "m365_backup_profile.json"), "JSON (*.json)")
        if path: self.app_settings.export_profile(path)

    def import_settings_profile(self):
        path, _ = QFileDialog.getOpenFileName(self, "Importar perfil", str(self.project_root), "JSON (*.json)")
        if path: self.app_settings.import_profile(path); self.load_runtime_settings_into_ui(); self._apply_style()

    def show_diagnostics(self):
        results = EnvironmentDiagnostics(
            self.project_root, self.app_settings, self.credential_store
        ).run()
        text = "\n".join(
            f"{'OK' if item['ok'] else 'ERRO'} - {item['name']}: {item['details']}"
            for item in results
        )
        QMessageBox.information(
            self, f"Sobre e diagnóstico - v{APP_VERSION}", text
        )

    def validate_selected_backup(self):
        selected = self.selected_backup_mailboxes()
        if not selected:
            QMessageBox.information(self, "Integridade", "Selecione uma mailbox.")
            return
        path = self.backup_jobs.get(selected[0], {}).get("resume_path")
        if not path:
            QMessageBox.information(
                self, "Integridade", "A operação ainda não possui uma pasta de backup."
            )
            return
        report = IntegrityValidator().validate_backup(path, complete=False)
        message = (
            f"Status: {report['status']}\n"
            f"EML: {report['eml']}\n"
            f"Válidos: {report['valid']}\n"
            f"Inválidos: {len(report['invalid'])}\n"
            f"Temporários: {len(report['partial'])}"
        )
        QMessageBox.information(self, "Integridade", message)

    def _apply_style(self):
        theme = self.app_settings.get("appearance", "theme", "automatic")
        if theme == "automatic":
            theme = "dark" if QApplication.palette().window().color().lightness() < 128 else "light"
        font = safe_int(self.app_settings.get("appearance", "font_size", 10), 10)
        dark = theme == "dark"
        bg = "#0f172a" if dark else "#f5f7fb"; panel = "#172033" if dark else "#ffffff"; text = "#e5e7eb" if dark else "#172033"; muted = "#94a3b8" if dark else "#64748b"; border = "#334155" if dark else "#cbd5e1"; field = "#111827" if dark else "#ffffff"; header = "#1e293b" if dark else "#eef2f7"
        self.setStyleSheet(f"""
            * {{ font-family: 'Segoe UI'; font-size: {font}pt; color: {text}; }}
            QMainWindow, QWidget#content, QDialog {{ background: {bg}; }}
            QFrame#sidebar {{ background: #172033; border: none; }}
            QLabel#brand {{ color: white; font-size: 20pt; font-weight: 700; }}
            QLabel#sidebarMuted, QLabel#muted {{ color: {muted}; }}
            QListWidget#navigation {{ background: transparent; color: #cbd5e1; border: none; outline: 0; }}
            QListWidget#navigation::item {{ padding: 12px 14px; margin: 3px 0; border-radius: 7px; }}
            QListWidget#navigation::item:selected {{ background: #2563eb; color: white; }}
            QFrame#statCard, QGroupBox, QTabWidget::pane {{ background: {panel}; border: 1px solid {border}; border-radius: 9px; }}
            QGroupBox {{ margin-top: 10px; padding-top: 14px; font-weight: 600; }}
            QPushButton {{ background: {field}; border: 1px solid {border}; border-radius: 7px; padding: 8px 13px; }}
            QPushButton:hover {{ background: {header}; }}
            QPushButton#primaryButton {{ background: #2563eb; color: white; border-color: #2563eb; font-weight: 600; }}
            QLineEdit, QComboBox, QSpinBox, QTextEdit {{ background: {field}; border: 1px solid {border}; border-radius: 7px; padding: 7px; }}
            QTableWidget, QTreeWidget {{ background: {field}; border: 1px solid {border}; alternate-background-color: {header}; gridline-color: {border}; }}
            QHeaderView::section {{ background: {header}; border: none; border-bottom: 1px solid {border}; padding: 8px; font-weight: 600; }}
            QProgressBar {{ border: none; background: {header}; border-radius: 5px; text-align: center; min-height: 10px; }}
            QProgressBar::chunk {{ background: #2563eb; border-radius: 5px; }}
            QScrollBar:vertical {{ background: {bg}; width: 12px; margin: 0; }}
            QScrollBar::handle:vertical {{ background: #64748b; min-height: 28px; border-radius: 6px; }}
            QScrollBar:horizontal {{ background: {bg}; height: 12px; }}
            QScrollBar::handle:horizontal {{ background: #64748b; min-width: 28px; border-radius: 6px; }}
        """)
        application = QApplication.instance()
        if isinstance(application, QApplication):
            application.setFont(QFont("Segoe UI", font))


    def change_page(self, index):
        titles = (
            ("Backups", "Acompanhe a fila, a retomada, o download de EML e os logs por mailbox."),
            ("Conversões PST", "Converta backups EML com checkpoint e retomada segura."),
            ("Logs", "Acompanhe os logs operacionais e de falhas em tempo real."),
            ("Configurações", "Preferências gerais da aplicação e dos novos trabalhos.")
        )
        self.stack.setCurrentIndex(index)
        self.top_title.setText(titles[index][0])
        self.top_subtitle.setText(titles[index][1])

    def append_mailbox_log(self, mailbox, message, level="Informação"):
        if not mailbox:
            return
        stamp = datetime.now().strftime("%H:%M:%S")
        line = f"{stamp} · {level} · {message}"
        lines = self.mailbox_logs.setdefault(str(mailbox), [])
        lines.append(line)
        if len(lines) > self.max_mailbox_log_lines:
            del lines[:-self.max_mailbox_log_lines]
        selected = self.selected_backup_mailboxes() if hasattr(self, "backup_table") else []
        if selected and selected[0] == mailbox and hasattr(self, "mailbox_log_text"):
            self.mailbox_log_text.append(line)
            scrollbar = self.mailbox_log_text.verticalScrollBar()
            scrollbar.setValue(scrollbar.maximum())

    def refresh_selected_mailbox_log(self):
        if not hasattr(self, "mailbox_log_text"):
            return
        selected = self.selected_backup_mailboxes()
        if not selected:
            self.mailbox_log_title.setText("Selecione uma mailbox para acompanhar a atividade.")
            self.mailbox_log_text.clear()
            return
        mailbox = selected[0]
        job = self.backup_jobs.get(mailbox, {})
        stage = job.get("resume_stage") or self.backup_stage_text(job)
        self.mailbox_log_title.setText(f"{mailbox} · {stage}")
        self.mailbox_log_text.setPlainText("\n".join(self.mailbox_logs.get(mailbox, [])))
        scrollbar = self.mailbox_log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def clear_selected_mailbox_log(self):
        selected = self.selected_backup_mailboxes()
        if not selected:
            return
        self.mailbox_logs[selected[0]] = []
        self.refresh_selected_mailbox_log()

    def backup_stage_text(self, job):
        status = str(job.get("status") or "")
        if status == "pausado":
            return "Pausado no checkpoint"
        if status in {"solicitando pausa", "pausando"}:
            return "Salvando ponto de pausa"
        if status == "concluído":
            return "Concluído"
        if status in {"erro", "incompleto", "interrompido"}:
            return "Aguardando retomada"
        if status == "iniciando" and job.get("resume_path"):
            return "Carregando checkpoint"
        if status in STATUS_RUNNING and job.get("resume_path"):
            if safe_int(job.get("current")) > safe_int(job.get("resume_baseline")):
                return "Baixando EML novamente"
            return "Retomando · aguardando primeiro EML"
        if status in STATUS_RUNNING:
            return "Baixando EML"
        if status == "pendente":
            return "Aguardando início"
        return status or "Aguardando"

    def log_event(self, message, level="Informação", technical=None):
        stamp = datetime.now().strftime("%H:%M:%S")
        text = f"{stamp} · {level} · {message}"
        # The former dashboard was removed. Keep compatibility if an older
        # customization still creates its activity widget, without requiring it.
        activity = getattr(self, "dashboard_activity", None)
        if activity is not None:
            activity.append(text)
            if technical:
                activity.append(f"    Detalhes: {technical}")

    def quick_add_mailbox(self):
        mailbox, accepted = self.simple_input(
            "Adicionar mailbox", "E-mail do colaborador:"
        )
        if not accepted or not mailbox.strip():
            return
        mailbox = mailbox.strip()
        if mailbox in self.backup_jobs:
            QMessageBox.information(self, "Fila", "Esta mailbox já está na fila.")
            return
        self.add_backup_job(mailbox)
        self.log_event(
            f"{mailbox} adicionada à fila. A configuração será solicitada ao iniciar."
        )

    def simple_input(self, title, label):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(label))
        edit = QLineEdit()
        layout.addWidget(edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        accepted = dialog.exec() == QDialog.DialogCode.Accepted
        return edit.text(), accepted

    def add_backup_job(self, mailbox, resume_path=None, options=None):
        if mailbox in self.backup_jobs:
            return False
        current = 0
        expected = 0
        if resume_path:
            progress = self.read_checkpoint_progress(resume_path)
            current = progress["current"]
            expected = progress["expected"]
            options = options or progress.get("options")
        self.backup_jobs[mailbox] = {
            "mailbox": mailbox,
            "status": "pendente",
            "current": current,
            "expected": expected,
            "downloaded_bytes": 0,
            "resume_path": (
                str(resume_path) if resume_path
                else str(self.stable_backup_path(mailbox))
            ),
            "options": options or self.default_job_options(),
            "worker": None,
            "started_at": None,
            "mailbox_total": expected,
            "analysis_folder_count": 0,
            "analysis_status": "concluída" if expected else "aguardando",
            "analysis_error": "",
            "eml_per_second": 0.0,
            "speed_samples": 0,
            "speed_last_current": current,
            "speed_last_time": time.time()
        }
        self.backup_order.append(mailbox)
        self.render_backup_table()
        self.save_state()
        if not resume_path or expected <= 0:
            self.enqueue_mailbox_analysis(mailbox)
        return True

    def default_job_options(self):
        return {
            "selected_folder_ids": [],
            "selected_folder_paths": [],
            "all_messages": self.default_all.isChecked(),
            "attachments": False,
            "skip_calendar": self.default_skip_calendar.isChecked(),
            "skip_contacts": self.default_skip_contacts.isChecked(),
            "skip_tasks": self.default_skip_tasks.isChecked(),
            "profile_only": False,
            "skip_precheck": self.default_skip_precheck.isChecked(),
            "limit": ""
        }

    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Importar mailboxes", str(self.project_root), "CSV (*.csv)"
        )
        if not path:
            return
        count = 0
        with open(path, "r", encoding="utf-8-sig", newline="") as file:
            for row in csv.DictReader(file):
                mailbox = (row.get("email") or "").strip()
                if mailbox and self.add_backup_job(mailbox):
                    count += 1
        self.log_event(f"{count} mailbox(es) adicionada(s) à fila.")

    def add_resume_backup(self):
        path = QFileDialog.getExistingDirectory(
            self, "Selecionar pasta de backup existente", self.destination_edit.text()
        )
        if not path:
            return
        checkpoint = Path(path) / "checkpoint.json"
        mailbox = Path(path).parent.name
        if checkpoint.exists():
            try:
                mailbox = json.loads(checkpoint.read_text(encoding="utf-8")).get("mailbox") or mailbox
            except Exception:
                pass
        if self.add_backup_job(mailbox, path):
            self.log_event(
                f"Backup existente de {mailbox} adicionado sem recontagem física de arquivos."
            )

    def selected_backup_mailboxes(self):
        result = []
        selection_model = self.backup_table.selectionModel()
        if selection_model is None:
            return result
        for index in selection_model.selectedRows():
            item = self.backup_table.item(index.row(), 0)
            if item is not None:
                result.append(item.data(Qt.ItemDataRole.UserRole))
        return result

    def configure_selected_backup(self):
        selected = self.selected_backup_mailboxes()
        if not selected:
            QMessageBox.information(self, "Configuração", "Selecione uma mailbox.")
            return
        mailbox = selected[0]
        job = self.backup_jobs.get(mailbox, {})
        if job.get("status") in STATUS_RUNNING:
            QMessageBox.information(
                self,
                "Configuração em uso",
                "Pause este backup antes de alterar as pastas ou as opções. "
                "O progresso já salvo será preservado."
            )
            return
        self.configure_mailbox(mailbox)

    def enqueue_mailbox_analysis(self, mailbox, priority=False):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return False
        if mailbox in self._folder_cache:
            job["analysis_status"] = "concluída"
            return True
        if mailbox == self._analysis_active_mailbox:
            return True
        if mailbox in self._analysis_queue:
            if priority:
                self._analysis_queue.remove(mailbox)
                self._analysis_queue.insert(0, mailbox)
            return True
        job["analysis_status"] = "aguardando"
        job["analysis_error"] = ""
        if priority:
            self._analysis_queue.insert(0, mailbox)
        else:
            self._analysis_queue.append(mailbox)
        self.render_backup_table()
        QTimer.singleShot(0, self.start_next_mailbox_analysis)
        return True

    def start_next_mailbox_analysis(self):
        active_thread = getattr(self, "_folder_thread", None)
        if active_thread is not None and active_thread.isRunning():
            return
        while self._analysis_queue:
            mailbox = self._analysis_queue.pop(0)
            job = self.backup_jobs.get(mailbox)
            if not job or mailbox in self._folder_cache:
                continue
            self._analysis_active_mailbox = mailbox
            job["analysis_status"] = "calculando"
            job["analysis_error"] = ""
            self.render_backup_table()
            self.log_event(
                f"Calculando automaticamente a quantidade de e-mails de {mailbox}..."
            )
            thread = QThread()
            worker = FolderLoadWorker(
                mailbox=mailbox,
                env_file=self.project_root / ".env",
                metrics_path=self.metrics_path
            )
            worker.moveToThread(thread)
            thread.started.connect(worker.run)
            worker.progress.connect(
                self._folder_analysis_progress,
                Qt.ConnectionType.QueuedConnection
            )
            worker.loaded.connect(
                self._folder_analysis_loaded,
                Qt.ConnectionType.QueuedConnection
            )
            worker.failed.connect(
                self._folder_analysis_failed,
                Qt.ConnectionType.QueuedConnection
            )
            worker.completed.connect(worker.deleteLater)
            worker.completed.connect(thread.quit)
            thread.finished.connect(self._folder_thread_finished)
            thread.finished.connect(thread.deleteLater)
            self._folder_thread = thread
            self._folder_worker = worker
            thread.start()
            return

    def _folder_analysis_progress(self, mailbox, total_items, folder_count):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return
        job["mailbox_total"] = max(
            safe_int(job.get("mailbox_total")), safe_int(total_items)
        )
        job["analysis_folder_count"] = safe_int(folder_count)
        job["analysis_status"] = "calculando"
        self.render_backup_table()

    def _folder_analysis_loaded(self, mailbox, folders):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return
        total = sum(safe_int(folder.get("totalItemCount")) for folder in folders)
        self._folder_cache[mailbox] = folders
        job["mailbox_total"] = total
        job["analysis_folder_count"] = len(folders)
        job["analysis_status"] = "concluída"
        job["analysis_error"] = ""
        if not safe_int(job.get("expected")):
            job["expected"] = total
        self.render_backup_table()
        self.save_state()
        self.log_event(
            f"Pré-análise concluída para {mailbox}: "
            f"{total:,} e-mails em {len(folders)} pasta(s)."
        )
        pending = self._pending_folder_dialog
        if pending and pending[0] == mailbox:
            self._pending_folder_dialog = None
            context = pending[1]
            QTimer.singleShot(
                0,
                lambda m=mailbox, f=folders, c=context:
                    self._open_cached_folder_dialog(m, f, c)
            )

    def _folder_analysis_failed(self, mailbox, error):
        job = self.backup_jobs.get(mailbox)
        if job:
            job["analysis_status"] = "erro"
            job["analysis_error"] = str(error)
        self.render_backup_table()
        self.save_state()
        self.log_event(
            f"Não foi possível calcular a quantidade de e-mails de {mailbox}.",
            "Erro", str(error)
        )
        pending = self._pending_folder_dialog
        if pending and pending[0] == mailbox:
            self._pending_folder_dialog = None
            if pending[1] == "start":
                self._start_config_skipped += 1
                QTimer.singleShot(0, self.configure_next_backup_for_start)
            else:
                QMessageBox.critical(self, "Pastas", str(error))

    def _folder_thread_finished(self):
        self._folder_worker = None
        self._folder_thread = None
        self._analysis_active_mailbox = None
        QTimer.singleShot(0, self.start_next_mailbox_analysis)

    def configure_mailbox(self, mailbox, context="manual"):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return False
        folders = self._folder_cache.get(mailbox)
        if folders is not None:
            QTimer.singleShot(
                0,
                lambda m=mailbox, f=folders, c=context:
                    self._open_cached_folder_dialog(m, f, c)
            )
            return True
        self._pending_folder_dialog = (mailbox, context)
        self.enqueue_mailbox_analysis(mailbox, priority=True)
        if context == "start":
            self.log_event(
                f"Aguardando a pré-análise já iniciada de {mailbox}; "
                "a consulta não será repetida."
            )
        else:
            self.log_event(
                f"Aguardando a pré-análise de {mailbox} para abrir as pastas..."
            )
        self.start_next_mailbox_analysis()
        return True

    def _open_cached_folder_dialog(self, mailbox, folders, context):
        accepted = self.show_folder_dialog(mailbox, folders)
        if context != "start":
            return
        if accepted:
            if self.submit_backup_to_coordinator(mailbox):
                self._start_config_started += 1
            else:
                self._start_config_skipped += 1
        else:
            self._start_config_skipped += 1
            self.log_event(
                f"Início cancelado para {mailbox}; o backup permaneceu pendente.",
                "Aviso"
            )
        QTimer.singleShot(0, self.configure_next_backup_for_start)

    def show_folder_dialog(self, mailbox, folders):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return False
        dialog = FolderScopeDialog(mailbox, folders, job.get("options"), self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return False
        if not isinstance(dialog.result_config, dict):
            return False
        config = dict(dialog.result_config)
        apply_all = config.pop("apply_all", False)
        targets = list(self.backup_jobs) if apply_all else [mailbox]
        if apply_all:
            by_path = config.get("selected_folder_paths", [])
            for target in targets:
                target_config = dict(config)
                if target != mailbox:
                    target_config["selected_folder_ids"] = []
                    target_config["selected_folder_paths"] = by_path
                self.backup_jobs[target]["options"] = target_config
        else:
            job["options"] = config
        selected_total = sum(
            safe_int(folder.get("totalItemCount"))
            for folder in folders
            if str(folder.get("id")) in set(config.get("selected_folder_ids") or [])
        )
        job["selected_total"] = selected_total
        self.render_backup_table()
        self.save_state()
        self.log_event(
            f"Configuração salva para {len(targets)} trabalho(s); "
            f"{len(config.get('selected_folder_ids', []))} pasta(s) selecionada(s)."
        )
        return True

    def configure_next_backup_for_start(self):
        if not self._start_config_active:
            return
        active_thread = getattr(self, "_folder_thread", None)
        if active_thread is not None and active_thread.isRunning():
            QTimer.singleShot(100, self.configure_next_backup_for_start)
            return
        if not self._start_config_queue:
            self._start_config_active = False
            self._folder_request_context = "manual"
            self.render_backup_table()
            self.save_state()
            self.refresh_from_coordinator()
            self.log_event(
                f"Configuração de início concluída: "
                f"{self._start_config_started} backup(s) iniciado(s) e "
                f"{self._start_config_skipped} mantido(s) pendente(s)."
            )
            return
        mailbox = self._start_config_queue.pop(0)
        job = self.backup_jobs.get(mailbox, {})
        if job.get("operation_id") or job.get("status") != "pendente":
            QTimer.singleShot(0, self.configure_next_backup_for_start)
            return
        if not self.configure_mailbox(mailbox, context="start"):
            self._start_config_skipped += 1
            QTimer.singleShot(0, self.configure_next_backup_for_start)

    def stable_backup_path(self, mailbox):
        folder_name = re.sub(r'[\\/:*?"<>|]+', "_", str(mailbox).strip())
        folder_name = re.sub(r"\s+", " ", folder_name).strip(" .")
        folder_name = (folder_name or "sem_nome")[:120]
        return Path(self.destination_edit.text()).expanduser().resolve() / folder_name

    def submit_backup_to_coordinator(self, mailbox):
        job = self.backup_jobs.get(mailbox, {})
        if not job or job.get("operation_id") or job.get("status") != "pendente":
            return False
        try:
            stable_path = Path(job.get("resume_path") or self.stable_backup_path(mailbox))
            stable_path.mkdir(parents=True, exist_ok=True)
            job["resume_path"] = str(stable_path)
            csv_path = self.temp_dir / f"{self.safe_name(mailbox)}.csv"
            with open(csv_path, "w", encoding="utf-8-sig", newline="") as file:
                writer = csv.DictWriter(file, fieldnames=["email"])
                writer.writeheader()
                writer.writerow({"email": mailbox})
            if job.get("resume_path"):
                job.setdefault("options", {})
                job["options"]["skip_precheck"] = True
                job["options"]["skip_eml_scan_on_resume"] = True
                job["options"]["use_checkpoint_progress"] = True
            options_path = self.temp_dir / f"{self.safe_name(mailbox)}_options.json"
            options_path.write_text(
                json.dumps(job["options"], ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            command = self.backend_command(
                csv_path, options_path, job.get("resume_path")
            )
            operation = self.coordinator_request("POST", "/operations", {
                "operation_type": "backup", "mailbox": mailbox,
                "command": command, "backup_path": job.get("resume_path"),
                "destination_path": self.destination_edit.text(),
                "options": {
                    **job["options"],
                    "runtime_settings": self.app_settings.snapshot_for_operation(
                        "backup", job["options"]
                    )
                }
            })
            if not isinstance(operation, dict) or not operation.get("operation_id"):
                raise RuntimeError(
                    "O coordenador não retornou uma operação de backup válida."
                )
            job["operation_id"] = str(operation["operation_id"])
            job["status"] = "iniciando"
            self.reset_backup_speed_baseline(mailbox)
            self.render_backup_table()
            self.save_state()
            self.log_event(f"Backup iniciado após configuração: {mailbox}.")
            self.append_mailbox_log(mailbox, "Processo iniciado; preparando download de EML.")
            return True
        except Exception as error:
            QMessageBox.critical(
                self, "Iniciar backup",
                f"{mailbox} não pôde ser iniciado.\n\n{error}"
            )
            self.log_event(
                f"Falha ao iniciar {mailbox} após a configuração.",
                "Erro", str(error)
            )
            return False

    def start_queue(self):
        if self._start_config_active:
            QMessageBox.information(
                self, "Iniciar backups",
                "A configuração sequencial de backups já está em andamento."
            )
            return
        if not self.ensure_coordinator():
            QMessageBox.critical(
                self, "Coordenador", "O serviço local não está disponível."
            )
            return
        active_statuses = {
            "iniciando", "executando", "continuando",
            "solicitando pausa", "pausando"
        }
        active_count = sum(
            1 for job in self.backup_jobs.values()
            if job.get("status") in active_statuses
        )
        available_slots = max(self.parallel_spin.value() - active_count, 0)
        if available_slots <= 0:
            QMessageBox.information(
                self, "Iniciar backups",
                "O limite de backups simultâneos já foi atingido."
            )
            return
        selected = self.selected_backup_mailboxes()
        selected_pending = [
            mailbox for mailbox in self.backup_order
            if mailbox in selected
            and not self.backup_jobs.get(mailbox, {}).get("operation_id")
            and self.backup_jobs.get(mailbox, {}).get("status") == "pendente"
        ]
        pending = selected_pending or [
            mailbox for mailbox in self.backup_order
            if not self.backup_jobs.get(mailbox, {}).get("operation_id")
            and self.backup_jobs.get(mailbox, {}).get("status") == "pendente"
        ]
        if not pending:
            QMessageBox.information(
                self, "Iniciar backups", "Não existem backups pendentes para iniciar."
            )
            return
        to_start = pending[:available_slots]
        try:
            self.coordinator_request(
                "PUT", "/settings/concurrency",
                {"backup_workers": self.parallel_spin.value(), "pst_workers": self.pst_parallel_spin.value()}
            )
        except Exception as error:
            QMessageBox.critical(self, "Coordenador", str(error))
            return
        self._start_config_queue = list(to_start)
        self._start_config_active = True
        self._start_config_started = 0
        self._start_config_skipped = 0
        self.log_event(
            f"Iniciando configuração individual de {len(to_start)} backup(s)."
        )
        self.configure_next_backup_for_start()

    def schedule_backups(self):
        # A fila e o paralelismo agora pertencem ao coordenador persistente.
        return

    def start_backup(self, mailbox):
        job = self.backup_jobs[mailbox]
        csv_path = self.temp_dir / f"{self.safe_name(mailbox)}.csv"
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["email"])
            writer.writeheader()
            writer.writerow({"email": mailbox})
        options_path = self.temp_dir / f"{self.safe_name(mailbox)}_options.json"
        options_path.write_text(
            json.dumps(job["options"], ensure_ascii=False, indent=2), encoding="utf-8"
        )
        command = self.backend_command(csv_path, options_path, job.get("resume_path"))
        environment = os.environ.copy()
        environment["PYTHONIOENCODING"] = "utf-8"
        environment["PYTHONUTF8"] = "1"
        environment["M365_BACKUP_OUTPUT_ROOT"] = str(Path(self.destination_edit.text()).resolve())
        environment["M365_API_METRICS_DB"] = str(self.metrics_path)
        environment["M365_BACKUP_ENV_PATH"] = str(self.project_root / ".env")
        job["status"] = "continuando" if job.get("resume_path") else "executando"
        job["started_at"] = time.time()
        self.start_process(
            kind="backup", key=mailbox, command=command, environment=environment,
            line_handler=lambda line: self.handle_backup_line(mailbox, line),
            finish_handler=lambda code: self.backup_finished(mailbox, code)
        )
        self.render_backup_table()
        self.log_event(f"Backup iniciado: {mailbox}.")

    def backend_command(self, csv_path, options_path, resume_path):
        if getattr(sys, "frozen", False):
            command = [str(self.project_root / "run_backend.exe")]
        else:
            command = [sys.executable, "-m", "src.main"]
        command += [
            "--phase", "5", "--batch", str(csv_path),
            "--job-options-file", str(options_path), "--all", "--skip-precheck"
        ]
        if resume_path:
            command += ["--resume-path", str(resume_path)]
        return command

    def start_process(self, kind, key, command, environment, line_handler, finish_handler):
        thread = QThread(self)
        worker = OutputWorker(command, self.project_root, environment)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.line.connect(line_handler)
        worker.finished.connect(finish_handler)
        worker.failed.connect(lambda error: self.log_event(
            f"Falha ao iniciar {key}.", "Erro", error
        ))
        worker.finished.connect(worker.deleteLater)
        worker.finished.connect(thread.quit)
        thread.finished.connect(thread.deleteLater)
        if kind == "backup":
            self.backup_workers[key] = worker
            self.backup_threads[key] = thread
        else:
            self.pst_workers[key] = worker
            self.pst_threads[key] = thread
        thread.start()

    def handle_backup_line(self, mailbox, line):
        repair_marker = "[REPAIR-PROGRESS] "
        if repair_marker in line:
            try:
                payload = json.loads(line.split(repair_marker, 1)[1].strip())
                job = self.backup_jobs[mailbox]
                job["resume_stage"] = (
                    f"Corrigindo falhas · {safe_int(payload.get('current'))}/"
                    f"{safe_int(payload.get('total'))}"
                )
                job["repair_recovered"] = safe_int(payload.get("recovered"))
                job["repair_failed"] = safe_int(payload.get("failed"))
                self.append_mailbox_log(mailbox, job["resume_stage"])
                self.schedule_backup_render()
            except Exception as error:
                self.log_event("Progresso de correção inválido.", "Aviso", str(error))
            return
        marker = "[PROGRESS] "
        if marker in line:
            try:
                payload = json.loads(line.split(marker, 1)[1].strip())
                job = self.backup_jobs[mailbox]
                new_current = safe_int(payload.get("current"))
                self.update_backup_speed(mailbox, new_current)
                job["current"] = new_current
                reported_expected = safe_int(payload.get("expected"))
                if reported_expected:
                    job["expected"] = reported_expected
                elif not safe_int(job.get("expected")):
                    job["expected"] = safe_int(job.get("mailbox_total"))
                job["downloaded_bytes"] = safe_int(payload.get("downloaded_bytes"))
                current, total = job["current"], job["expected"]
                last = safe_int(job.get("last_log"), -1)
                if current == total or current - last >= 50:
                    job["last_log"] = current
                    self.log_event(
                        f"{mailbox}: progresso salvo em {current}/{total or '?'} e-mails."
                    )
                self.schedule_backup_render()
                self.schedule_state_save()
            except Exception as error:
                self.log_event("Progresso recebido em formato inválido.", "Aviso", str(error))
            return
        if " | ERROR | " in line:
            self.log_event(f"{mailbox}: ocorreu um problema durante o backup.", "Erro", line)
        elif "Throttling" in line or "429" in line:
            self.log_event(
                f"{mailbox}: o Microsoft Graph solicitou uma redução temporária da velocidade.",
                "Aviso", line
            )

    def backup_finished(self, mailbox, code):
        job = self.backup_jobs.get(mailbox)
        if not job:
            return
        self.backup_workers.pop(mailbox, None)
        self.backup_threads.pop(mailbox, None)
        if job["status"] == "pausado":
            self.log_event(f"Backup pausado: {mailbox}. O progresso foi preservado.")
        elif code == 0:
            job["status"] = "concluído"
            self.log_event(f"Backup concluído: {mailbox}.")
        else:
            job["status"] = "erro"
            self.log_event(f"Backup finalizado com problema: {mailbox}.", "Erro")
        if not job.get("resume_path"):
            latest = self.find_latest_backup(mailbox)
            if latest:
                job["resume_path"] = str(latest)
        self.render_backup_table()
        self.save_state()

    def repair_selected_backup(self):
        selected = self.selected_backup_mailboxes()
        if not selected:
            QMessageBox.information(self, "Corrigir falhas", "Selecione uma mailbox.")
            return
        if len(selected) > 1:
            QMessageBox.information(self, "Corrigir falhas", "Selecione apenas uma mailbox por vez.")
            return
        mailbox = selected[0]
        job = self.backup_jobs.get(mailbox, {})
        if job.get("status") in STATUS_RUNNING or job.get("status") in {"iniciando", "pausando"}:
            QMessageBox.warning(self, "Corrigir falhas", "Pause o backup antes de iniciar a correção.")
            return
        backup_path = Path(job.get("resume_path") or self.stable_backup_path(mailbox))
        if not (backup_path / "checkpoint.json").is_file():
            QMessageBox.warning(self, "Corrigir falhas", "O checkpoint do backup não foi encontrado.")
            return
        checkpoint = self.read_checkpoint_progress(backup_path)
        answer = QMessageBox.question(
            self, "Corrigir falhas",
            "O sistema tentará novamente somente os EML ainda com falha, sem enumerar "
            "toda a mailbox. Deseja continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        if not self.ensure_coordinator():
            QMessageBox.critical(self, "Coordenador", "O serviço local não está disponível.")
            return
        command = (
            [str(self.project_root / "run_backend.exe")]
            if getattr(sys, "frozen", False)
            else [sys.executable, "-m", "src.main"]
        )
        command += ["--repair-failures", "--resume-path", str(backup_path), "--mailbox", mailbox]
        try:
            operation = self.coordinator_request("POST", "/operations", {
                "operation_type": "backup", "mailbox": mailbox,
                "command": command, "backup_path": str(backup_path),
                "destination_path": self.destination_edit.text(),
                "options": {
                    "operation_mode": "repair_failures",
                    "runtime_settings": self.app_settings.snapshot_for_operation("backup", {})
                }
            })
            job["operation_id"] = operation.get("operation_id")
            job["status"] = "iniciando"
            job["resume_stage"] = "Lendo falhas do checkpoint"
            job["repair_recovered"] = 0
            job["repair_failed"] = safe_int(checkpoint.get("failed", 0))
            self.append_mailbox_log(mailbox, "Correção rápida iniciada; somente falhas serão processadas.")
            self.render_backup_table()
            self.save_state()
        except Exception as error:
            QMessageBox.critical(self, "Corrigir falhas", str(error))

    def pause_selected_backup(self):
        for mailbox in self.selected_backup_mailboxes():
            operation_id = self.backup_jobs.get(mailbox, {}).get("operation_id")
            if operation_id:
                try:
                    self.append_mailbox_log(mailbox, "Solicitação de pausa enviada; salvando checkpoint.")
                    self.coordinator_request("POST", f"/operations/{operation_id}/pause")
                    self.backup_jobs[mailbox]["eml_per_second"] = 0.0
                    self.backup_jobs[mailbox]["speed_samples"] = 0
                except Exception as error:
                    self.log_event(f"Falha ao pausar {mailbox}.", "Erro", str(error))
        self.refresh_from_coordinator()

    def resume_selected_backup(self):
        for mailbox in self.selected_backup_mailboxes():
            job = self.backup_jobs.get(mailbox, {})
            operation_id = job.get("operation_id")
            if operation_id:
                try:
                    self.append_mailbox_log(mailbox, "Retomada solicitada; carregando checkpoint e preparando EML pendentes.")
                    self.coordinator_request("POST", f"/operations/{operation_id}/resume")
                    self.reset_backup_speed_baseline(mailbox)
                except Exception as error:
                    self.log_event(f"Falha ao retomar {mailbox}.", "Erro", str(error))
            elif job.get("status") in STATUS_RESUMABLE:
                job["status"] = "pendente"
        self.render_backup_table()
        self.save_state()

    def apply_dragged_backup_order(self, visible_order):
        if self.backup_filter.text().strip() or self.backup_status_filter.currentData():
            # visible_order só contém as linhas atualmente exibidas na tabela;
            # com um filtro de status ativo, isso é um subconjunto — recalcular
            # backup_order a partir dele empurraria todo o resto da fila
            # (inclusive operações em execução) para o fim, silenciosamente.
            self.render_backup_table()
            QMessageBox.information(
                self, "Ordem da fila",
                "Limpe a busca e o filtro de status antes de arrastar a fila."
            )
            return
        tail = [mailbox for mailbox in self.backup_order if mailbox not in visible_order]
        self.backup_order = list(visible_order) + tail
        for position, mailbox in enumerate(self.backup_order, 1):
            self.backup_jobs.setdefault(mailbox, {})["queue_position"] = position
        operation_ids = [
            self.backup_jobs.get(mailbox, {}).get("operation_id")
            for mailbox in self.backup_order
            if self.backup_jobs.get(mailbox, {}).get("operation_id")
        ]
        if operation_ids:
            try:
                self.coordinator_request("PUT", "/queue/order", {"operation_ids": operation_ids})
            except Exception as error:
                self.log_event("Ordem salva localmente; o coordenador será atualizado depois.", "Aviso", str(error))
        self.save_state()
        self.render_backup_table()
        self.log_event("Ordem da fila alterada por arrastar e soltar.")

    def _reorder_selected_backups(self, destination):
        selected = self.selected_backup_mailboxes()
        if not selected:
            QMessageBox.information(self, "Ordem da fila", "Selecione uma ou mais mailboxes.")
            return
        if self.backup_filter.text().strip() or self.backup_status_filter.currentData():
            QMessageBox.information(
                self,
                "Ordem da fila",
                "Limpe a busca e o filtro de status antes de alterar a ordem "
                "para visualizar a fila completa."
            )
            return

        selected_set = set(selected)
        ordered_selected = [
            mailbox for mailbox in self.backup_order if mailbox in selected_set
        ]
        if not ordered_selected:
            return

        if destination == "up":
            for mailbox in ordered_selected:
                index = self.backup_order.index(mailbox)
                if index > 0 and self.backup_order[index - 1] not in selected_set:
                    self.backup_order[index - 1], self.backup_order[index] = (
                        self.backup_order[index], self.backup_order[index - 1]
                    )
        elif destination == "down":
            for mailbox in reversed(ordered_selected):
                index = self.backup_order.index(mailbox)
                if (
                    index < len(self.backup_order) - 1
                    and self.backup_order[index + 1] not in selected_set
                ):
                    self.backup_order[index + 1], self.backup_order[index] = (
                        self.backup_order[index], self.backup_order[index + 1]
                    )
        elif destination == "top":
            remaining = [
                mailbox for mailbox in self.backup_order if mailbox not in selected_set
            ]
            self.backup_order = ordered_selected + remaining
        elif destination == "bottom":
            remaining = [
                mailbox for mailbox in self.backup_order if mailbox not in selected_set
            ]
            self.backup_order = remaining + ordered_selected
        else:
            return

        self.render_backup_table()
        self._restore_backup_selection(ordered_selected)
        self.save_state()
        operation_ids = [
            self.backup_jobs[mailbox].get("operation_id") for mailbox in self.backup_order
            if self.backup_jobs.get(mailbox, {}).get("operation_id")
        ]
        if operation_ids:
            try:
                self.coordinator_request("PUT", "/queue/order", {"operation_ids": operation_ids})
            except Exception as error:
                self.log_event("A ordem foi salva localmente, mas não chegou ao coordenador.", "Aviso", str(error))
        positions = [str(self.backup_order.index(mailbox) + 1) for mailbox in ordered_selected]
        self.log_event(
            f"Ordem da fila atualizada. Posição(ões): {', '.join(positions)}."
        )
        # Se o agendador estiver ativo, a próxima mailbox será escolhida na nova ordem.
        if self.scheduler_active:
            self.schedule_backups()

    def _restore_backup_selection(self, mailboxes, current_mailbox=None):
        targets = set(mailboxes or [])
        table = self.backup_table
        selection_model = table.selectionModel()
        if selection_model is None:
            return
        table.blockSignals(True)
        try:
            table.clearSelection()
            current_index = None
            for row in range(table.rowCount()):
                item = table.item(row, 0)
                mailbox = item.data(Qt.ItemDataRole.UserRole) if item else None
                if mailbox in targets:
                    row_index = table.model().index(row, 0)
                    selection_model.select(
                        row_index,
                        QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows
                    )
                if mailbox == current_mailbox:
                    current_index = table.model().index(row, 0)
            if current_index is not None:
                selection_model.setCurrentIndex(
                    current_index, QItemSelectionModel.SelectionFlag.NoUpdate
                )
        finally:
            table.blockSignals(False)
        self.refresh_selected_mailbox_log()

    def _restore_pst_selection(self, operation_ids, current_operation_id=None):
        targets = set(operation_ids or [])
        table = self.pst_table
        selection_model = table.selectionModel()
        if selection_model is None:
            return
        table.blockSignals(True)
        try:
            table.clearSelection()
            current_index = None
            for row in range(table.rowCount()):
                item = table.item(row, 0)
                operation_id = item.data(Qt.ItemDataRole.UserRole) if item else None
                if operation_id in targets:
                    row_index = table.model().index(row, 0)
                    selection_model.select(
                        row_index,
                        QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows
                    )
                if operation_id == current_operation_id:
                    current_index = table.model().index(row, 0)
            if current_index is not None:
                selection_model.setCurrentIndex(
                    current_index, QItemSelectionModel.SelectionFlag.NoUpdate
                )
        finally:
            table.blockSignals(False)

    def move_selected_backup_up(self):
        self._reorder_selected_backups("up")

    def move_selected_backup_down(self):
        self._reorder_selected_backups("down")

    def move_selected_backup_top(self):
        self._reorder_selected_backups("top")

    def move_selected_backup_bottom(self):
        self._reorder_selected_backups("bottom")

    def remove_selected_backup(self):
        for mailbox in self.selected_backup_mailboxes():
            if mailbox in self.backup_workers:
                continue
            operation_id = self.backup_jobs.get(mailbox, {}).get("operation_id")
            if operation_id:
                try:
                    self.coordinator_request("DELETE", f"/operations/{operation_id}")
                except Exception:
                    continue
            self.backup_jobs.pop(mailbox, None)
            if mailbox in self.backup_order:
                self.backup_order.remove(mailbox)
        self.render_backup_table()
        self.save_state()

    def backup_context_menu(self, position):
        menu = QMenu(self)
        for text, callback in (
            ("Configurar pastas e opções", self.configure_selected_backup),
            ("Pausar", self.pause_selected_backup),
            ("Retomar", self.resume_selected_backup),
            ("Corrigir somente falhas", self.repair_selected_backup),
            ("Subir na fila", self.move_selected_backup_up),
            ("Descer na fila", self.move_selected_backup_down),
            ("Mover para o topo", self.move_selected_backup_top),
            ("Mover para o final", self.move_selected_backup_bottom),
            ("Abrir pasta do backup", self.open_selected_backup_folder)
        ):
            action = QAction(text, self)
            action.triggered.connect(callback)
            menu.addAction(action)
        menu.exec(self.backup_table.viewport().mapToGlobal(position))

    def open_selected_backup_folder(self):
        selected = self.selected_backup_mailboxes()
        if not selected:
            return
        job = self.backup_jobs[selected[0]]
        path = job.get("resume_path") or Path(self.destination_edit.text()) / selected[0]
        open_local_path(path)

    def render_backup_table(self):
        selected = (
            set(self.selected_backup_mailboxes())
            if self.backup_table.rowCount() else set()
        )
        current_item = self.backup_table.currentItem()
        current_mailbox = None
        if current_item is not None:
            mailbox_item = self.backup_table.item(current_item.row(), 0)
            if mailbox_item is not None:
                current_mailbox = mailbox_item.data(Qt.ItemDataRole.UserRole)
        vertical_position = self.backup_table.verticalScrollBar().value()
        horizontal_position = self.backup_table.horizontalScrollBar().value()
        self.backup_table.setUpdatesEnabled(False)
        self.backup_table.blockSignals(True)
        self.backup_table.setRowCount(0)
        query = (
            self.backup_filter.text().strip().lower()
            if hasattr(self, "backup_filter") else ""
        )
        status_filter = (
            self.backup_status_filter.currentData() or ""
            if hasattr(self, "backup_status_filter") else ""
        )
        visible_count = 0
        for mailbox in self.backup_order:
            job = self.backup_jobs.get(mailbox)
            if not job:
                continue
            status = str(job.get("status") or "").lower()
            options = job.get("options") or {}
            searchable = " ".join((
                mailbox, status, str(job.get("resume_path") or ""),
                " ".join(str(value) for value in options.get(
                    "selected_folder_paths", []
                )), str(job.get("current") or ""),
                str(job.get("expected") or "")
            )).lower()
            if query and query not in searchable:
                continue
            if status_filter and status != status_filter:
                continue
            visible_count += 1
            row = self.backup_table.rowCount()
            self.backup_table.insertRow(row)
            current = safe_int(job.get("current"))
            total = safe_int(job.get("expected")) or safe_int(job.get("mailbox_total"))
            mailbox_total = safe_int(job.get("mailbox_total"))
            analysis_status = str(job.get("analysis_status") or "aguardando")
            if analysis_status == "calculando":
                analysis_text = (
                    f"Calculando · {safe_int(job.get('analysis_folder_count'))} pastas"
                )
            elif analysis_status == "concluída":
                analysis_text = (
                    f"Concluída · {safe_int(job.get('analysis_folder_count'))} pastas"
                )
            elif analysis_status == "erro":
                analysis_text = "Erro na análise"
            else:
                analysis_text = "Aguardando"
            count_text = f"{mailbox_total:,}" if mailbox_total else "Calculando..."
            progress = (
                f"{current}/{total} ({current / total * 100:.1f}%)"
                if total else f"{current} concluídos"
            )
            remaining = max(total - current, 0) if total else "A calcular"
            configured = (
                f"{len(options.get('selected_folder_ids') or [])} pastas"
                if options.get("selected_folder_ids") else "Todas/padrão"
            )
            eta_text = self.backup_eta_text(job, total, current)
            stage_text = job.get("resume_stage") or self.backup_stage_text(job)
            if (
                job.get("resume_path")
                and job.get("status") in STATUS_RUNNING
                and current > safe_int(job.get("resume_baseline"))
            ):
                stage_text = "Baixando EML novamente"
                job["resume_stage"] = stage_text
            values = (
                mailbox, job.get("status", ""), stage_text, count_text, analysis_text,
                progress, str(remaining),
                format_bytes(job.get("downloaded_bytes", 0)), eta_text,
                str(safe_int(job.get("analysis_folder_count"))), configured,
                str(self.backup_order.index(mailbox) + 1)
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, mailbox)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.backup_table.setItem(row, column, item)
        self.backup_table.blockSignals(False)
        self.backup_table.setUpdatesEnabled(True)
        self._restore_backup_selection(selected, current_mailbox)
        self.backup_table.verticalScrollBar().setValue(vertical_position)
        self.backup_table.horizontalScrollBar().setValue(horizontal_position)
        if hasattr(self, "backup_filter_counter"):
            self.backup_filter_counter.setText(
                f"{visible_count} de {len(self.backup_jobs)} exibidos"
            )
        self.refresh_dashboard()

    def filter_backup_table(self, *args):
        self.render_backup_table()

    def clear_backup_filters(self):
        self.backup_filter.clear()
        self.backup_status_filter.setCurrentIndex(0)
        self.render_backup_table()

    def new_pst_job(self):
        pst_settings = self.app_settings.data.get("pst", {})
        defaults = {
            "source": pst_settings.get("last_source_root") or self.destination_edit.text(),
            "destination_dir": self.pst_output_edit.text(),
            "file_name": "backup.pst", "display_name": pst_settings.get("last_display_name", "M365 Mailbox Backup"),
            "existing_action": pst_settings.get("existing_action", "resume"),
            "manual_start": True, "folder_mode": pst_settings.get("folder_mode", "preserve"),
            "root_folder_name": pst_settings.get("root_folder_name", ""),
            "visible_metadata": pst_settings.get("visible_metadata", True),
            "import_attachments": pst_settings.get("import_attachments", True),
            "image_max_width": pst_settings.get("image_max_width", 700),
            "import_rate": pst_settings.get("eml_import_rate_per_second", 10),
            "verification_level": pst_settings.get("verification_level", "balanced"),
            "verification_batch_size": pst_settings.get("verification_batch_size", 25),
            "detach_after": pst_settings.get("detach_after", False),
            "open_folder_after": pst_settings.get("open_folder_after", False),
            "prepare_workers": pst_settings.get("prepare_workers", 3),
            "prepare_queue_size": pst_settings.get("prepare_queue_size", 12),
            "large_eml_mb": pst_settings.get("large_eml_mb", 25),
            "performance_profile": pst_settings.get("performance_profile", "balanced"),
            "adaptive_enabled": pst_settings.get("adaptive_enabled", True),
            "memory_budget_mb": pst_settings.get("memory_budget_mb", 512),
            "min_prepare_workers": pst_settings.get("min_prepare_workers", 1),
            "max_prepare_workers": pst_settings.get("max_prepare_workers", 4),
        }
        dialog = PstCustomizationDialog(defaults, self)
        if dialog.exec() != QDialog.DialogCode.Accepted: return
        config = dialog.result_config
        if not isinstance(config, dict):
            QMessageBox.critical(self, "Conversão PST", "A configuração da conversão não foi gerada.")
            return
        pst_settings.update({
            "last_source_root": config["source"], "last_display_name": config["display_name"],
            "existing_action": config["existing_action"], "folder_mode": config["folder_mode"],
            "root_folder_name": config["root_folder_name"], "visible_metadata": config["visible_metadata"],
            "import_attachments": config["import_attachments"], "image_max_width": config["image_max_width"],
            "eml_import_rate_per_second": config["import_rate"], "verification_level": config["verification_level"],
            "verification_batch_size": config["verification_batch_size"],
            "detach_after": config["detach_after"], "open_folder_after": config["open_folder_after"],
            "prepare_workers": config["prepare_workers"],
            "prepare_queue_size": config["prepare_queue_size"],
            "large_eml_mb": config["large_eml_mb"],
            "performance_profile": config["performance_profile"],
            "adaptive_enabled": config["adaptive_enabled"],
            "memory_budget_mb": config["memory_budget_mb"],
            "min_prepare_workers": config["min_prepare_workers"],
            "max_prepare_workers": config["max_prepare_workers"]
        })
        self.app_settings.save()
        self.enqueue_pst(config["source"], config["pst_path"], config)

    def add_multiple_pst(self):
        root = QFileDialog.getExistingDirectory(
            self, "Selecionar pasta que contém backups", self.destination_edit.text()
        )
        if not root:
            return
        sessions = [path.parent for path in Path(root).rglob("checkpoint.json")]
        for session in sessions:
            target = Path(self.pst_output_edit.text()) / f"{session.parent.name}_{session.name}.pst"
            self.enqueue_pst(session, target)

    def enqueue_pst(self, backup, pst, options=None):
        backup = str(Path(backup).resolve())
        pst = str(Path(pst).resolve())
        options = dict(options or {})
        duplicate = next((
            job_id for job_id, job in self.pst_jobs.items()
            if os.path.normcase(str(Path(job.get("pst") or "").resolve()))
            == os.path.normcase(pst)
            and job.get("status") not in {"concluído", "cancelado", "erro"}
        ), None)
        if duplicate:
            QMessageBox.warning(
                self, "Conversão PST",
                f"Já existe uma conversão ativa ou pendente para este destino:\n{pst}"
            )
            return
        if not self.ensure_coordinator():
            QMessageBox.critical(self, "Coordenador", "O serviço local não está disponível.")
            return
        command = (
            [str(self.project_root / "run_pst.exe")]
            if getattr(sys, "frozen", False)
            else [sys.executable, "-m", "src.services.pst_export_service"]
        )
        command += [
            "--backup-root", str(Path(backup).resolve()),
            "--pst-path", str(Path(pst).resolve()),
            "--pst-display-name", options.get("display_name") or f"Backup {Path(backup).parent.name}",
            "--existing-action", options.get("existing_action", "resume"),
            "--folder-mode", options.get("folder_mode", "preserve"),
            "--image-max-width", str(options.get("image_max_width", 700)),
            "--verification-level", options.get("verification_level", "balanced"),
            "--verification-batch-size", str(options.get("verification_batch_size", 25)),
            "--import-rate", str(options.get("import_rate", 10)),
            "--prepare-workers", str(options.get("prepare_workers", 3)),
            "--prepare-queue-size", str(options.get("prepare_queue_size", 12)),
            "--large-eml-mb", str(options.get("large_eml_mb", 25)),
            "--performance-profile", options.get("performance_profile", "balanced"),
            "--memory-budget-mb", str(options.get("memory_budget_mb", 512)),
            "--min-prepare-workers", str(options.get("min_prepare_workers", 1)),
            "--max-prepare-workers", str(options.get("max_prepare_workers", 4))
        ]
        if not options.get("adaptive_enabled", True):
            command.append("--disable-adaptive")
        if options.get("root_folder_name"):
            command += ["--root-folder-name", options["root_folder_name"]]
        if not options.get("visible_metadata", True): command.append("--hide-visible-metadata")
        if not options.get("import_attachments", True): command.append("--skip-attachments")
        if options.get("detach_after", self.pst_detach.isChecked()):
            command.append("--detach-after")
        try:
            operation = self.coordinator_request("POST", "/operations", {
                "operation_type": "pst", "command": command,
                "source_path": backup,
                "destination_path": pst,
                "options": {
                    **options, "manual_start": options.get("manual_start", True),
                    "runtime_settings": self.app_settings.snapshot_for_operation("pst", options)
                }
            })
            if not isinstance(operation, dict):
                raise RuntimeError("O coordenador não retornou a operação PST criada.")
            self.log_event(
                f"Conversão {operation['operation_id']} adicionada com snapshot operacional da GUI. "
                "Selecione-a e clique em Iniciar selecionadas."
            )
            self.refresh_from_coordinator()
        except Exception as error:
            # Sem isso, uma conversão PST duplicada (ex.: clique duplo antes do
            # próximo poll atualizar self.pst_jobs) virava um HTTP 500 que
            # subia sem tratamento até aqui e desaparecia só no crash log,
            # sem nenhum aviso na tela.
            QMessageBox.critical(
                self, "Conversão PST",
                f"Não foi possível adicionar a conversão à fila.\n\n{error}"
            )
            self.log_event(
                "Falha ao adicionar conversão PST à fila.", "Erro", str(error)
            )

    def start_selected_pst(self):
        selected = self.selected_pst_ids()
        if not selected:
            QMessageBox.information(
                self, "Conversões PST", "Selecione uma ou mais conversões pendentes."
            )
            return
        active_statuses = {"iniciando", "executando", "solicitando pausa", "pausando"}
        active_count = sum(
            1 for job in self.pst_jobs.values()
            if job.get("status") in active_statuses
        )
        available_slots = max(self.pst_parallel_spin.value() - active_count, 0)
        if available_slots <= 0:
            QMessageBox.information(
                self, "Conversões PST",
                "O limite de conversões PST simultâneas já foi atingido."
            )
            return
        pending = [
            operation_id for operation_id in selected
            if self.pst_jobs.get(operation_id, {}).get("status")
            in {"pendente", "pausado", "interrompido", "erro", "incompleto"}
        ]
        if not pending:
            QMessageBox.information(
                self, "Conversões PST", "A seleção não contém conversões iniciáveis."
            )
            return
        to_start = pending[:available_slots]
        try:
            self.coordinator_request(
                "PUT", "/settings/concurrency",
                {
                    "backup_workers": self.parallel_spin.value(),
                    "pst_workers": self.pst_parallel_spin.value()
                }
            )
            started = 0
            for operation_id in to_start:
                self.coordinator_request("POST", f"/operations/{operation_id}/resume")
                started += 1
            self.log_event(
                f"{started} conversão(ões) PST iniciada(s) em processos independentes."
            )
            self.refresh_from_coordinator()
        except Exception as error:
            QMessageBox.critical(self, "Conversões PST", str(error))

    def start_pst(self, job_id):
        job = self.pst_jobs[job_id]
        if getattr(sys, "frozen", False):
            command = [str(self.project_root / "run_pst.exe")]
        else:
            command = [sys.executable, "-m", "src.services.pst_export_service"]
        command += [
            "--backup-root", job["backup"], "--pst-path", job["pst"],
            "--pst-display-name", f"Backup {Path(job['backup']).parent.name}"
        ]
        if self.pst_detach.isChecked():
            command.append("--detach-after")
        environment = os.environ.copy()
        environment["PYTHONIOENCODING"] = "utf-8"
        environment["PYTHONUTF8"] = "1"
        job["status"] = "executando"
        self.start_process(
            "pst", job_id, command, environment,
            lambda line: self.handle_pst_line(job_id, line),
            lambda code: self.pst_finished(job_id, code)
        )
        self.render_pst_table()

    def handle_pst_line(self, job_id, line):
        marker = "[PST-PROGRESS] "
        if marker in line:
            try:
                payload = json.loads(line.split(marker, 1)[1].strip())
                job = self.pst_jobs[job_id]
                job["current"] = safe_int(payload.get("current"))
                job["expected"] = safe_int(payload.get("total")) or job["expected"]
                self.schedule_pst_render()
                self.schedule_state_save()
            except Exception:
                pass
        elif "ERROR" in line or "[FALHA]" in line:
            self.log_event(f"{job_id}: ocorreu um problema na conversão PST.", "Erro", line)

    def pst_finished(self, job_id, code):
        job = self.pst_jobs.get(job_id)
        if not job:
            return
        self.pst_workers.pop(job_id, None)
        self.pst_threads.pop(job_id, None)
        if job["status"] == "pausado":
            self.log_event(f"{job_id} pausado com checkpoint preservado.")
        elif code == 0:
            job["status"] = "concluído"
            job["current"] = job["expected"]
            self.log_event(f"{job_id} concluído.")
        else:
            job["status"] = "erro"
            self.log_event(f"{job_id} finalizado com problema.", "Erro")
        self.render_pst_table()
        self.save_state()

    def selected_pst_ids(self):
        result = []
        selection_model = self.pst_table.selectionModel()
        if selection_model is None:
            return result
        for index in selection_model.selectedRows():
            item = self.pst_table.item(index.row(), 0)
            if item is not None:
                result.append(item.data(Qt.ItemDataRole.UserRole))
        return result

    def pause_selected_pst(self):
        for operation_id in self.selected_pst_ids():
            try:
                self.coordinator_request("POST", f"/operations/{operation_id}/pause")
            except Exception as error:
                self.log_event(f"Falha ao pausar {operation_id}.", "Erro", str(error))
        self.refresh_from_coordinator()

    def resume_selected_pst(self):
        for operation_id in self.selected_pst_ids():
            try:
                self.coordinator_request("POST", f"/operations/{operation_id}/resume")
            except Exception as error:
                self.log_event(f"Falha ao retomar {operation_id}.", "Erro", str(error))
        self.refresh_from_coordinator()

    def remove_selected_pst(self):
        for job_id in self.selected_pst_ids():
            try:
                self.coordinator_request("DELETE", f"/operations/{job_id}")
                self.pst_jobs.pop(job_id, None)
            except Exception:
                pass
        self.render_pst_table()
        self.save_state()

    def render_pst_table(self, *args):
        selected = (
            set(self.selected_pst_ids())
            if self.pst_table.rowCount() else set()
        )
        current_item = self.pst_table.currentItem()
        current_operation_id = None
        if current_item is not None:
            operation_item = self.pst_table.item(current_item.row(), 0)
            if operation_item is not None:
                current_operation_id = operation_item.data(Qt.ItemDataRole.UserRole)
        vertical_position = self.pst_table.verticalScrollBar().value()
        horizontal_position = self.pst_table.horizontalScrollBar().value()
        self.pst_table.setUpdatesEnabled(False)
        self.pst_table.blockSignals(True)
        self.pst_table.setRowCount(0)
        query = (
            self.pst_filter.text().strip().lower()
            if hasattr(self, "pst_filter") else ""
        )
        status_filter = (
            self.pst_status_filter.currentData() or ""
            if hasattr(self, "pst_status_filter") else ""
        )
        for job_id, job in self.pst_jobs.items():
            status = str(job.get("status") or "").lower()
            searchable = " ".join((
                str(job_id), str(job.get("backup") or ""),
                str(job.get("pst") or ""), str((job.get("options") or {}).get("display_name") or ""), status
            )).lower()
            if query and query not in searchable:
                continue
            if status_filter and status != status_filter:
                continue
            row = self.pst_table.rowCount()
            self.pst_table.insertRow(row)
            current = safe_int(job["current"])
            total = safe_int(job["expected"])
            progress = (
                f"{current}/{total} ({current / total * 100:.1f}%)"
                if total else str(current)
            )
            if job.get("stage"):
                progress += f" · {job.get('stage')}"
            options = job.get("options") or {}
            policy_names = {"resume": "Retomar", "number": "Numerar", "replace": "Substituir", "cancel": "Cancelar"}
            values = (
                job_id, options.get("file_name") or Path(job["pst"]).name, job["backup"], job["pst"],
                options.get("display_name") or "M365 Mailbox Backup",
                policy_names.get(options.get("existing_action", "resume"), "Retomar"),
                job["status"], progress,
                (
                    f"{job.get('verification_mode', 'balanced')} · {safe_int(job.get('verification_verified'))} OK · "
                    f"{safe_int(job.get('verification_pending'))} pend. · {job.get('bottleneck', 'calculando')}"
                    + (
                        f" · retomada: EML #{safe_int(job.get('resume_first_source_position'))} confirmado em "
                        f"{float(job.get('resume_first_committed_seconds', 0) or 0):.2f}s"
                        if float(job.get('resume_first_committed_seconds', 0) or 0) > 0 else ""
                    )
                ),
                safe_int(job.get("failed")) + safe_int(job.get("audit_failures")),
                max(total - current, 0) if total else "?"
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, job_id)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.pst_table.setItem(row, column, item)
        self.pst_table.blockSignals(False)
        self.pst_table.setUpdatesEnabled(True)
        self._restore_pst_selection(selected, current_operation_id)
        self.pst_table.verticalScrollBar().setValue(vertical_position)
        self.pst_table.horizontalScrollBar().setValue(horizontal_position)
        self.refresh_dashboard()

    def read_pst_summary(self, pst):
        path = Path(pst).with_suffix(".pst_checkpoint.json")
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def refresh_metrics(self):
        try:
            minutes = self.metrics_period.currentData() or 60
            summary = self.metrics.summary(minutes)
            self.metric_cards["health"].set_value(summary.get("health", "Sem dados"))
            self.metric_cards["requests"].set_value(
                f"{safe_int(summary.get('total'))} · {safe_int(summary.get('failures'))} falhas"
            )
            self.metric_cards["latency"].set_value(
                f"{summary.get('avg_latency_ms', 0):.0f} / {summary.get('p95_latency_ms', 0):.0f} ms"
            )
            self.metric_cards["rate"].set_value(
                f"{summary.get('requests_per_second', 0):.2f} req/s"
            )
            self.metric_cards["throttle"].set_value(
                f"{summary.get('throttle_percent', 0):.2f}% · Graph {summary.get('wait_seconds', 0):.0f}s · "
                f"Limitador {summary.get('rate_limiter_wait_seconds', 0):.0f}s"
            )
            self.recommendation_label.setText(
                f"Capacidade operacional estimada: {summary.get('health', 'Sem dados')}. "
                f"{summary.get('recommendation', '')} "
                "O Microsoft Graph não fornece uma cota restante universal exata."
            )
            self.fill_metrics_table(
                self.category_table, self.metrics.grouped("category", minutes)
            )
            self.fill_metrics_table(
                self.mailbox_metrics_table, self.metrics.grouped("mailbox", minutes)
            )
            dashboard_cards = getattr(self, "dashboard_cards", None)
            if dashboard_cards and "health" in dashboard_cards:
                dashboard_cards["health"].set_value(summary.get("health", "Sem dados"))
        except Exception as error:
            self.recommendation_label.setText(f"Métricas indisponíveis: {error}")

    def fill_metrics_table(self, table, rows):
        table.setRowCount(0)
        for row_data in rows:
            row = table.rowCount()
            table.insertRow(row)
            values = (
                row_data.get("name"), row_data.get("requests", 0),
                row_data.get("failures", 0), row_data.get("throttles", 0),
                f"{row_data.get('avg_latency_ms', 0):.0f} ms",
                f"{row_data.get('wait_seconds', 0):.0f} s"
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, column, item)

    def export_metrics(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exportar relatório", str(self.project_root / "api_performance.csv"),
            "CSV (*.csv)"
        )
        if path:
            self.metrics.export_csv(path, self.metrics_period.currentData() or 60)
            self.log_event(f"Relatório de desempenho exportado: {path}")

    def refresh_dashboard(self):
        # Compatibility no-op. The Visão geral page and its cards were removed.
        dashboard_cards = getattr(self, "dashboard_cards", None)
        if not dashboard_cards:
            return
        running = sum(1 for job in self.backup_jobs.values() if job["status"] in STATUS_RUNNING)
        paused = sum(1 for job in self.backup_jobs.values() if job["status"] == "pausado")
        paused += sum(1 for job in self.pst_jobs.values() if job["status"] == "pausado")
        active_pst = sum(1 for job in self.pst_jobs.values() if job["status"] == "executando")
        dashboard_cards["running"].set_value(running)
        dashboard_cards["paused"].set_value(paused)
        dashboard_cards["pst"].set_value(active_pst)

    def select_destination(self):
        path = QFileDialog.getExistingDirectory(
            self, "Destino dos backups", self.destination_edit.text()
        )
        if path:
            self.destination_edit.setText(path)
            self.save_state()

    def find_latest_backup(self, mailbox):
        root = self.stable_backup_path(mailbox)
        if (root / "checkpoint.json").exists():
            return root
        if not root.exists():
            return None
        candidates = [path.parent for path in root.rglob("checkpoint.json")]
        return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else root

    def read_checkpoint_progress(self, backup_path):
        path = Path(backup_path) / "checkpoint.json"
        try:
            checkpoint = json.loads(path.read_text(encoding="utf-8"))
            scope = checkpoint.get("backup_scope") or {}
            return {
                "current": safe_int(checkpoint.get("exported_message_count")),
                "expected": safe_int(checkpoint.get("expected_message_count")),
                "options": {
                    "selected_folder_ids": scope.get("selected_folder_ids", []),
                    "selected_folder_paths": scope.get("selected_folder_paths", []),
                    "all_messages": scope.get("export_all_messages", True),
                    "attachments": scope.get("export_attachments", False),
                    "skip_calendar": True, "skip_contacts": True,
                    "skip_tasks": True, "profile_only": False,
                    "skip_precheck": True, "limit": scope.get("message_limit_per_folder", "")
                }
            }
        except Exception:
            return {"current": 0, "expected": 0, "options": None}

    def save_state(self):
        data = {
            "backup_order": self.backup_order,
            "backup_jobs": {
                mailbox: {
                    key: value for key, value in job.items()
                    if key not in ("worker", "started_at")
                } for mailbox, job in self.backup_jobs.items()
            },
            "pst_jobs": self.pst_jobs,
            "pst_counter": self.pst_counter,
            "mailbox_logs": self.mailbox_logs,
            "destination": self.destination_edit.text(),
            "pst_output": self.pst_output_edit.text(),
            "parallel": self.parallel_spin.value(),
            "pst_parallel": self.pst_parallel_spin.value()
        }
        temporary = self.state_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, self.state_path)

    def load_state(self):
        try:
            data = json.loads(self.state_path.read_text(encoding="utf-8"))
        except Exception:
            return
        self.backup_order = data.get("backup_order", [])
        self.backup_jobs = data.get("backup_jobs", {})
        self.backup_order = [
            mailbox for mailbox in self.backup_order if mailbox in self.backup_jobs
        ]
        for mailbox in self.backup_jobs:
            if mailbox not in self.backup_order:
                self.backup_order.append(mailbox)
        for mailbox, job in self.backup_jobs.items():
            if job.get("status") in STATUS_RUNNING:
                job["status"] = "pausado"
            job.setdefault("mailbox_total", safe_int(job.get("expected")))
            job.setdefault("analysis_folder_count", 0)
            if job.get("analysis_status") == "calculando":
                job["analysis_status"] = "aguardando"
            job.setdefault(
                "analysis_status",
                "concluída" if safe_int(job.get("mailbox_total")) else "aguardando"
            )
            job.setdefault("analysis_error", "")
            job.setdefault("eml_per_second", 0.0)
            job.setdefault("speed_samples", 0)
            job["speed_last_current"] = safe_int(job.get("current"))
            job["speed_last_time"] = time.time()
        self.pst_jobs = data.get("pst_jobs", {})
        for job in self.pst_jobs.values():
            if job.get("status") == "executando":
                job["status"] = "pausado"
        self.pst_counter = safe_int(data.get("pst_counter"))
        self.mailbox_logs = data.get("mailbox_logs", {}) or {}
        self.destination_edit.setText(data.get("destination") or self.destination_edit.text())
        self.pst_output_edit.setText(data.get("pst_output") or self.pst_output_edit.text())
        self.parallel_spin.setValue(safe_int(data.get("parallel"), 2))
        self.pst_parallel_spin.setValue(safe_int(data.get("pst_parallel"), self.pst_parallel_spin.value()))
        self.render_backup_table()
        self.render_pst_table()
        for mailbox, job in self.backup_jobs.items():
            if (
                not safe_int(job.get("mailbox_total"))
                or job.get("analysis_status") in {"aguardando", "erro"}
            ):
                self.enqueue_mailbox_analysis(mailbox)

    def safe_name(self, value):
        return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value))

    def closeEvent(self, event):
        active = any(
            job.get("status") in {
                "executando", "iniciando", "solicitando pausa", "pausando"
            }
            for job in list(self.backup_jobs.values()) + list(self.pst_jobs.values())
        )
        if active:
            answer = QMessageBox.question(
                self,
                "Encerrar aplicação",
                "Existem operações em andamento. Fechar a interface encerrará toda a "
                "aplicação e pausará os processos no último checkpoint seguro. Deseja continuar?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return

        self.save_state()
        for timer_name in (
            "coordinator_timer", "dashboard_timer", "log_timer"
        ):
            timer = getattr(self, timer_name, None)
            if timer is not None:
                timer.stop()
        try:
            self.coordinator_request("POST", "/shutdown", timeout=2)
        except Exception:
            pass
        process = getattr(self, "coordinator_process", None)
        if process is not None and process.poll() is None:
            try:
                process.wait(timeout=12)
            except Exception:
                try:
                    process.terminate()
                    process.wait(timeout=3)
                except Exception:
                    try:
                        process.kill()
                    except Exception:
                        pass
        self.crash_reporter.clean_shutdown()
        event.accept()


def main():
    instance = SingleInstance()
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setStyle("Fusion")
    if not instance.primary:
        QMessageBox.information(None, APP_TITLE, "O M365 Mailbox Backup já está em execução.")
        return
    window = M365BackupWindow()
    window.show()
    exit_code = app.exec()
    instance.close()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
