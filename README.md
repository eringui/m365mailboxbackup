# M365 Mailbox Backup (GPT 5.6)

A Windows tool that backs up **Microsoft 365** mailboxes via the **Microsoft Graph API**, exports messages to **.eml** with checkpoint-based resume support, and optionally converts the backup to **.pst** using **Outlook Classic** COM automation. It provides a graphical interface (PySide6) with an operation queue and real-time progress, as well as a complete CLI for command-line use or automation/scheduling.

> Application version: `2.0.0`

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Architecture](#architecture)
- [Requirements](#requirements)
- [Microsoft Entra ID App Registration](#microsoft-entra-id-app-registration)
- [Installation](#installation)
- [Configuration (.env)](#configuration-env)
- [Usage — Graphical Interface](#usage--graphical-interface)
- [Interface Screenshots](#interface-screenshots)
- [Usage — Command Line](#usage--command-line)
- [Batch Backup (CSV)](#batch-backup-csv)
- [Resume and Checkpoints](#resume-and-checkpoints)
- [PST Export (Outlook Classic)](#pst-export-outlook-classic)
- [Increasing the .pst Size Limit in Windows (Registry)](#increasing-the-pst-size-limit-in-windows-registry)
- [Rate Limiting and Adaptive Throttling](#rate-limiting-and-adaptive-throttling)
- [Logs, Metrics, and Diagnostics](#logs-metrics-and-diagnostics)
- [Build / Packaging (PyInstaller)](#build--packaging-pyinstaller)
- [Troubleshooting](#troubleshooting)
- [Security](#security)
- [License](#license)

## Overview

The project has two independent stages:

1. **Backup** — authenticates to Microsoft Graph as an application (client credentials, with no interactive user login), traverses folders, messages, calendar, contacts, and tasks from one or more mailboxes, and exports everything to `.eml` files on disk, with a transactional checkpoint to allow resuming.
2. **PST conversion** *(optional)* — reads an existing `.eml` backup and, using COM automation from Outlook Classic installed on the machine, builds a browsable `.pst` file while preserving folder structure, sender/recipient, dates, and attachments.

The graphical interface does not execute these stages directly: on first run, it starts a local HTTP service (**coordinator**, FastAPI at `127.0.0.1:8765`) that manages an operation queue and launches child (CLI) processes for each backup/conversion, persisting state in SQLite. This allows you to close and reopen the GUI without losing the progress of ongoing operations.

## Key Features

- **App-only authentication (MSAL)** — a single application in Entra ID can access any mailbox in the tenant without requiring individual user login.
- **Incremental export** — uses Microsoft Graph *delta queries* (`/messages/delta`) so subsequent runs retrieve only what has changed.
- **Transactional checkpoint (SQLite/WAL)** — each exported item is recorded; interruptions (network failure, cancellation, PC shutdown) can be resumed with `--resume-path` without reprocessing items that have already been saved.
- **Preventive rate limiting + adaptive throttling** — configurable limits per process, per mailbox, and shared across processes (via SQLite), automatically reducing parallelism when Graph responds with `429`.
- **Batch backup via CSV** — one mailbox per row, with pre-validation and results reports in JSON/CSV.
- **PST conversion with verification** — writes in batches, reconciles pending items, and verifies what was actually persisted in the `.pst`.
- **Graphical interface (PySide6)** — operation queues, real-time progress, pause/resume/cancel controls, concurrency management, API telemetry, and environment diagnostics.
- **Credentials protected by DPAPI** on Windows, with `.env` fallback.
- **Graph call telemetry** (latency, throttle rate, success/failure) in SQLite, exportable to CSV.

## Architecture

> Structure inferred from the project's imports (`from src.config...`, `from src.services...`, `python -m src.main`).

```
.
├── .env                        # credentials and local overrides (DO NOT commit)
├── .gitignore
├── requirements.txt
├── application_runtime.py      # AppSettings, CredentialStore (DPAPI), CrashReporter, diagnóstico
├── m365_backup_gui.py           # official GUI entry point
├── m365_backup_gui_qt.py        # graphical interface implementation (PySide6)
├── m365_backup_coordinator.py   # serviço local FastAPI (127.0.0.1:8765) que orquestra operações
├── logs/                        # generated at runtime (rotating)
├── output/
│   ├── backups/                 # default output for .eml backups
│   └── pst/                     # default output for .pst files
├── _gui_state/                  # settings.json, credentials.bin, operations.sqlite3, rate limiter db
└── src/
    ├── main.py                  # CLI — fases 0 a 5
    ├── config/
    │   └── settings.py          # reads .env and default parameters
    ├── services/
    │   ├── graph_service.py         # autenticação MSAL + chamadas ao Microsoft Graph
    │   ├── mailbox_backup_service.py   # orquestra a exportação .eml (fases 2 a 5)
    │   ├── pst_export_service.py       # converts .eml to .pst via Outlook COM
    │   ├── checkpoint_store.py         # checkpoint transacional (SQLite/WAL)
    │   ├── operation_store.py          # coordinator operation queue/state (SQLite/WAL)
    │   └── api_metrics_store.py        # Graph call telemetry (SQLite/WAL)
    └── utils/
        └── logger.py             # rotating logging + dedicated logger per PST operation
```

**Typical flow (GUI):** `m365_backup_gui.py` → starts `m365_backup_coordinator.py` as a subprocess → the coordinator registers the operation in `operations.sqlite3` and launches `python -m src.main --phase 4 ...` (backup) or `python -m src.services.pst_export_service ...` (conversion) as a new process → progress is read from `stdout` (`[PROGRESS]` lines) and sent to the GUI via *polling*/WebSocket.

When packaged with PyInstaller, the same entry points become executables: the GUI, `m365_backup_coordinator.exe`, `run_backend.exe` (equivalent to `src.main`), and `run_pst.exe` (equivalent to `src.services.pst_export_service`).

## Requirements

- **Windows 10/11.** The project depends on `pywin32` and Outlook via COM; it does not work on macOS/Linux.
- **Outlook Classic** (traditional desktop application) installed with a configured profile — required **only** for the `.pst` conversion stage. `.eml` backup works without Outlook installed. The *new Outlook for Windows* is not compatible with the COM automation used here.
- **Python 3.11+** to run from source code (or use the executables generated by PyInstaller).
- An **application (App Registration)** in Microsoft Entra ID with application (app-only) permissions for Microsoft Graph and administrator consent.
- Network access to `login.microsoftonline.com` and `graph.microsoft.com`.

## Microsoft Entra ID App Registration

Authentication uses the *client credentials* flow (app-only, via MSAL) and accesses mailboxes through `/users/{mailbox}/...`. This requires **application permissions** (not delegated permissions):

| Graph permission | Type | Usage in the project |
|---|---|---|
| `User.Read.All` | Application | Resolve the mailbox/user (`/users/{mailbox}`) |
| `Mail.Read` | Application | Read folders and messages (`/mailFolders`, `/messages`, `/messages/delta`) |
| `Calendars.Read` | Application | Export calendar events |
| `Contacts.Read` | Aplicativo | Exportar contacts |
| `Tasks.Read` | Application | Export task lists (To Do) |

Step by step:

1. In the [Azure portal](https://portal.azure.com), go to **Microsoft Entra ID → App registrations → New registration**.
2. Give it a name (e.g. `M365 Mailbox Backup`) and keep the account type as "Accounts in this organizational directory only".
3. Record the **Application (client) ID** and **Directory (tenant) ID** — these will be `CLIENT_ID` and `TENANT_ID`.
4. Under **Certificates & secrets**, create a **New client secret** and copy the value immediately (it will not be displayed again) — this will be `CLIENT_SECRET`.
5. Em **API permissions → Add a permission → Microsoft Graph → Application permissions**, add the permissions from the table above.
6. Click **Grant admin consent** for the tenant.
7. *(Recommended)* Restrict the application's scope to specific mailboxes using an Exchange Online [Application Access Policy](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-configure-application-access-policy) (`New-ApplicationAccessPolicy`), preventing the app from having unrestricted access to the entire tenant.

## Installation

```bash
git clone <url-do-repositório> m365-mailbox-backup
cd m365-mailbox-backup

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

## Configuration (`.env`)

The `_env-example` file included in the repository is empty — use the template below as a reference and save it as `.env` in the project root (or point the `M365_BACKUP_ENV_PATH` variable to another path).

```env
# App Registration credentials (Microsoft Entra ID) — required
TENANT_ID=00000000-0000-0000-0000-000000000000
CLIENT_ID=00000000-0000-0000-0000-000000000000
CLIENT_SECRET=coloque-o-segredo-aqui

# Optional — defaults are already defined in src/config/settings.py
GRAPH_SCOPE=https://graph.microsoft.com/.default
GRAPH_URL=https://graph.microsoft.com/v1.0
M365_BACKUP_OUTPUT_ROOT=./output/backups
```

> In the GUI, credentials can also be stored securely using **DPAPI** (`_gui_state/credentials.bin`) through the settings screen, without requiring a `.env`.

Some advanced variables (all optional, with sensible default values):

| Variável | Padrão | Description |
|---|---|---|
| `M365_EML_DOWNLOAD_WORKERS` | `16` | Parallel `.eml` download threads |
| `M365_MIME_MAX_CONCURRENCY` | `3` | Simultaneous MIME downloads per mailbox |
| `M365_RATE_LIMITER_ENABLED` | `1` | Ativa o limitador preventivo (pyrate-limiter) |
| `M365_GLOBAL_MIME_RATE_MINUTE` | `240` | Global limit of MIME calls per minute |
| `M365_ADAPTIVE_THROTTLING` | `1` | Automatically reduces parallelism after `429` |
| `M365_PST_PREPARE_WORKERS` | `3` | Item preparation threads before COM writing (serial) |
| `M365_PST_CAPACITY_PREFLIGHT` | `1` | Checks disk space before writing the PST |
| `M365_DISK_WARNING_GB` / `M365_DISK_CRITICAL_GB` | `50` / `10` | Free disk space thresholds |

## Usage — Graphical Interface

```bash
python m365_backup_gui.py
```

When opened, the GUI:

1. Ensures a single instance (via a local socket on port `48765`) and starts the local coordinator (`127.0.0.1:8765`) if it is not already running.
2. On first run, requests the Entra ID credentials (or reads `.env`).
3. Allows you to add mailboxes individually or import a CSV for batch backup.
4. Allows you to create a **PST conversion** operation from an already completed backup.
5. Displays an operation queue with real-time progress, pause/resume/cancel, and queue reordering.
6. Allows you to adjust concurrency (how many backups and PST conversions run at the same time).
7. Provides an **API metrics** tab (latency, throttles, requests/sec) and an **"About and diagnostics"** dialog that validates credentials, write destinations, disk space, SQLite, and Outlook Classic availability.

## Interface Screenshots

The images below are stored beside `README.md` and `README_EN.md`. GitHub, VS Code, and other Markdown viewers can therefore display them through relative paths.

### EML backup queue

<p align="center">
<img width="1920" height="1041" alt="aba backup  eml" src="https://github.com/user-attachments/assets/2345cbb5-3f6b-4ba5-859a-b7d0c055d27a" />
</p>

The **Backups** page centralizes the mailbox queue, CSV import, existing-backup resume, folder configuration, pause, resume, and removal actions. The table shows status, current stage, mailbox-reported count, progress, remaining items, transferred volume, completion estimate, selected folders, and queue position. The lower panel displays the log for the selected mailbox.

### Folder selection and backup options

<p align="center">
<img width="1074" height="777" alt="configuração de pasta" src="https://github.com/user-attachments/assets/c22ba9dc-271d-4de0-a95d-b3183d6777f0" />
</p>

The **Folders and options** dialog loads only the folder hierarchy and counters before a backup starts. Operators can select every folder, clear the selection, apply a recommended selection, configure content options, skip calendar, contacts, and tasks, set a per-folder message limit, and apply the configuration to the whole queue.

> The counters shown in this dialog are lightweight values reported by Microsoft Graph. No message content is opened during this configuration stage.

### PST conversion queue

<p align="center">
<img width="1920" height="1039" alt="conversoes pst" src="https://github.com/user-attachments/assets/83ef2fa4-4c74-44da-93ba-f6fc946114df" />
</p>

The **PST Conversions** page supports creating conversions, starting selected jobs, pausing, resuming, removing, and opening the destination. The table shows source, destination, Outlook display name, existing-file policy, progress, verification, failures, and remaining items.

### Settings and credentials

<p align="center">
<img width="1920" height="1039" alt="config imagem" src="https://github.com/user-attachments/assets/7e08dbae-7e4e-4443-ba47-03e7fee2e585" />
</p>

The **Settings** page groups general preferences, Microsoft Entra ID credentials, operational parameters, and appearance. The client secret remains hidden and can be preserved without re-entering it. **Save and test credentials** validates authentication before operations begin.

> Before publishing your own screenshots, redact Tenant IDs, Client IDs, mailbox names, internal paths, and any other environment identifiers.

## Usage — Command Line

The CLI (`src/main.py`) is divided into phases:

| Phase | What it does | Required parameters |
|---|---|---|
| `0` | Validates authentication and mailbox access (quick diagnostic, exports nothing) | `--mailbox` |
| `1` | Inspects the mailbox (user, folders, inbox, message preview) | `--mailbox` |
| `2` | Initial local export (without full folder structure) | `--mailbox` |
| `3` | Export preserving folder structure | `--mailbox` |
| `4` | Complete export with checkpoint (recommended for real-world use) | `--mailbox` |
| `5` | Batch backup via CSV (calls phase 4 for each mailbox) | `--batch` |

Main options (phases 2 to 5):

| Option | Description |
|---|---|
| `--limit N` | Maximum messages per folder (default `25`), ignored if `--all` is used |
| `--all` / `--all-messages` | Exports all messages, with no per-folder limit |
| `--attachments` / `--export-attachments` | Exports attachments separately, in addition to the `.eml` |
| `--skip-calendar` / `--skip-contacts` / `--skip-tasks` | Skips these categories |
| `--resume-path <folder>` | Resumes an existing backup using the saved checkpoint |
| `--profile-only` | Exports only primary mail folders, ignoring archive/system folders (Archive, Recoverable Items, Sync Issues, Conversation History, etc.) |
| `--exclude-folder "Nome"` | Ignores a specified folder; the option can be repeated |
| `--job-options-file file.json` | JSON with `selected_folder_ids` / `selected_folder_paths` to restrict the scope |
| `--skip-precheck` | (phase 5) Skips CSV pre-validation |

Examples:

```bash
# Phase 0 — validates credentials and access to a specific mailbox
python -m src.main --phase 0 --mailbox usuario@empresa.com

# Phase 4 — full backup, with attachments and checkpoint enabled
python -m src.main --phase 4 --mailbox usuario@empresa.com --all --attachments

# Resuming an interrupted backup
python -m src.main --phase 4 --mailbox usuario@empresa.com --all \
    --resume-path "output/backups/usuario@empresa.com"
```

## Batch Backup (CSV)

The CSV file must contain an `email` column (one mailbox per row):

```csv
email
usuario1@empresa.com
usuario2@empresa.com
```

```bash
python -m src.main --phase 5 --batch users.csv --all --attachments
```

Phase 5 generates a pre-validation report and a final report (JSON and CSV) with the result for each mailbox.

## Resume and Checkpoints

Each backup operation maintains a transactional checkpoint in SQLite (WAL) inside the backup folder itself, recording folders, pages, and items already exported. If the process is interrupted (network, power loss, manual cancellation), simply run phase 4 again with `--resume-path` pointing to the same backup folder — items already confirmed in the checkpoint are not downloaded again.

## PST Export (Outlook Classic)

Converts an existing `.eml` backup into a `.pst` file, using COM automation (`Outlook.Application`) to create a **Unicode-format PST** (`namespace.AddStoreEx(path, 2)`), preserving folders, sender/recipient, body, dates, and attachments, as well as additional MAPI properties used for verification.

```bash
python -m src.services.pst_export_service \
    --backup-root "output/backups/usuario@empresa.com" \
    --pst-path "output/pst/usuario@empresa.com.pst" \
    --pst-display-name "Backup - User" \
    --verification-level balanced
```

Main options:

| Option | Padrão | Description |
|---|---|---|
| `--backup-root` | *(required)* | Backup folder or `.zip` containing the `.eml` files |
| `--pst-path` | *(required)* | Output `.pst` path |
| `--pst-display-name` | `M365 Mailbox Backup` | Name displayed in Outlook |
| `--existing-action` | `resume` | `resume` \| `number` \| `replace` \| `cancel`, if the PST already exists |
| `--folder-mode` | `preserve` | `preserve` (original structure) or `single` (single folder) |
| `--skip-attachments` | — | Does not import attachments |
| `--verification-level` | `balanced` | `quick` \| `balanced` \| `complete` |
| `--performance-profile` | `balanced` | `conservative` \| `balanced` \| `performance` \| `custom` |
| `--detach-after` | — | Removes the PST from Outlook after the conversion finishes |

The tool checks free disk space before and during writing (`M365_PST_CAPACITY_PREFLIGHT`) and safely stops the operation if the destination does not have enough space — this is independent of the Outlook limit described below.

## Increasing the .pst Size Limit in Windows (Registry)

### Why This Is Necessary

Because this project always creates the `.pst` in **Unicode format** (not the older ANSI format, which had a 2 GB limit), the actual ceiling you will encounter is not the historical 2 GB limit — it is the **default Outlook limit for Unicode files**, which has been **50 GB** since Outlook 2010 SP1. If a backup generates (or may grow to) a `.pst` larger than this, or if you want more headroom before the "file almost full" warning, you need to increase two values in the Windows Registry on the machine where PST conversion is performed.

### Manual Step-by-Step

1. Close Outlook completely (also check the system tray / Task Manager).
2. Press `Win + R`, type `regedit`, and press Enter (accept the UAC prompt).
3. **Recommended:** back up the current key — right-click `Outlook` (or `PST`, if it already exists) → **Export**.
4. Navigate to:

   ```
   Computador\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\PST
   ```

   `16.0` corresponds to Microsoft 365 / Outlook 2016, 2019, 2021, and 2024. Use `15.0` for Outlook 2013 or `14.0` for Outlook 2010.
5. If the `PST` key does not exist: right-click `Outlook` → **New → Key** → name it `PST`.
6. Inside the `PST` key, create (if they do not already exist) two **DWORD (32-bit)** values: right-click the right-hand pane → **New → DWORD (32-bit) Value** → name it `MaxLargeFileSize`; repeat for `WarnLargeFileSize`.
7. Double-click each value, select the **Decimal** base, and enter the size **in megabytes (MB)**. Using the same values shown in this project's screenshot:

   | Valor | Base | Decimal | Hexadecimal | Tamanho efetivo |
   |---|---|---|---|---|
   | `MaxLargeFileSize` | Decimal | `102400` | `0x00019000` | 100 GB |
   | `WarnLargeFileSize` | Decimal | `92160` | `0x00016000` | 90 GB |

8. Close Registry Editor and **restart Outlook**.

### Alternative Using a `.reg` File

Save the content below as `increase_pst_limit.reg`, double-click it, and confirm the import (adjust `16.0` according to your Outlook version, if necessary):

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\PST]
"MaxLargeFileSize"=dword:00019000
"WarnLargeFileSize"=dword:00016000
```

### Important Notes

- These values affect both `.pst` and `.ost` files (including the Exchange cache, if used by the profile).
- Keep `WarnLargeFileSize` at least ~5% below `MaxLargeFileSize` (in the example above, the difference is 10%), so that the warning appears before the hard limit is reached, without blocking Outlook operations.
- Documented absolute maximum: `MaxLargeFileSize` up to `4294967295` (`0xFFFFFFFF`) and `WarnLargeFileSize` up to `4090445042` (`0xF3CF3CF2`); above these values the setting is ignored.
- If the environment is managed by Group Policy, also check `HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\PST` — when present, this key takes precedence over `Software\Microsoft\...` and can only be changed by an administrator.
- Very large `.pst` files make Outlook slower to open, index, and compact; when possible, prefer splitting the backup by year or mailbox instead of using one huge PST.
- This setting belongs to Outlook/Windows and is **independent** of the tool's own disk-space check (`M365_PST_CAPACITY_PREFLIGHT`): even with the registry adjusted, the export also stops if the destination disk does not have enough free space.
- It must be done on **each Windows machine/user profile** where PST conversion will be performed.

## Rate Limiting and Adaptive Throttling

The project combines two mechanisms:

- **Preventive rate limiter** (`pyrate-limiter`, optional shared SQLite database between processes) — configurable per-second/minute limits, global and per-mailbox (`M365_GLOBAL_MIME_RATE_SECOND`, `M365_MAILBOX_MIME_RATE_SECOND`, etc.), to avoid triggering Graph throttling.
- **Reactive adaptive throttling** — when receiving `429`, respects the `Retry-After` header, performs backoff with jitter, and temporarily reduces parallelism (`M365_THROTTLE_SAFETY_SECONDS`, `M365_THROTTLE_JITTER_MAX_SECONDS`, `M365_THROTTLE_RECOVERY_SECONDS`), gradually returning to normal.

The GUI metrics tab shows API health (Normal / Attention / Limited) with a recommendation for adjusting parallelism, based on recent throttle rates and average latency.

## Logs, Metrics, and Diagnostics

- `logs/m365_mailbox_backup.log` — general rotating log (10 MB × 5 arquivos).
- `logs/pst/<operation-id>.log` — dedicated log for each PST conversion.
- `logs/application_crash.log` — captured by `CrashReporter` in case of an unhandled exception.
- SQLite database of Graph call metrics (latência, throttles, sucesso/falha), exportable to CSV through the GUI.
- The GUI's **"About and diagnostics"** dialog checks credentials, write destinations, disk space, SQLite, and Outlook Classic availability.

## Build / Packaging (PyInstaller)

The `requirements.txt` already includes `pyinstaller` and `pyinstaller-hooks-contrib`. In production, the project is distributed as four independent executables (the GUI and three processes it orchestrates):

| Source entry point | Expected executable |
|---|---|
| `m365_backup_gui.py` | GUI principal |
| `m365_backup_coordinator.py` | `m365_backup_coordinator.exe` |
| `src/main.py` (`python -m src.main`) | `run_backend.exe` |
| `src/services/pst_export_service.py` | `run_pst.exe` |

Illustrative individual build example (adjust `--hidden-import`, icons, and `--add-data` according to your `.spec` configuration, if applicable):

```bash
pyinstaller --onefile --name m365_backup_gui m365_backup_gui.py
pyinstaller --onefile --name m365_backup_coordinator m365_backup_coordinator.py
pyinstaller --onefile --name run_backend src/main.py
pyinstaller --onefile --name run_pst src/services/pst_export_service.py
```

## Troubleshooting

| Symptom | Likely cause | What to check |
|---|---|---|
| Authentication failure in Phase 0 | `TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET` incorrect or admin consent is pending | `.env`, application permissions in Entra ID |
| Many `429` responses / slow operation | Parallelism is too high for the tenant profile | Reduce `M365_MIME_MAX_CONCURRENCY` / concorrência na GUI |
| "Outlook Classic unavailable" no diagnóstico | Outlook is not installed or the *new* Outlook version is being used | Install/open Outlook Classic with a configured profile |
| PST does not grow beyond a certain size | Outlook's default 50 GB limit for Unicode files | See [Increasing the .pst size limit](#increasing-the-pst-size-limit-in-windows-registry) |
| GUI cannot connect to the coordinator | Port `8765` (coordinator) or `48765` (single instance) occupied by another process | Close old app instances/`m365_backup_coordinator.exe` |
| Operation pauses by itself | Insufficient disk space at the destination | `M365_DISK_WARNING_GB` / `M365_DISK_CRITICAL_GB`, free up space |

## Security

- Never commit `.env` (already listed in `.gitignore`) or the App Registration client secret.
- Prefer saving credentials through the GUI, which protects them with **DPAPI** (`_gui_state/credentials.bin`), bound to the Windows user/machine.
- Apply the principle of least privilege to the App Registration: grant only the required permissions and consider an Exchange Online *Application Access Policy* to restrict which mailboxes the app can access.
- `logs/`, `_gui_state/`, `output/backups/*`, and `output/pst/*` are already in `.gitignore` — they may contain personal/sensitive mailbox data.

## License

This project is publicly available for personal, educational, organizational, and non-commercial use.

You are free to:

Use the software for personal, educational, or organizational purposes;
Use the software within commercial organizations for internal business purposes;
Study and modify the source code;
Share the original or modified source code free of charge;
Create derivative works for non-commercial distribution.

## Commercialization

You may not sell, license, rent, lease, sublicense, or otherwise commercially distribute this software or derivative versions of it without explicit written permission from the author.

This includes, but is not limited to:

Selling the software or modified versions of it;
Charging users for access to the software;
Offering the software as a paid SaaS or hosted service;
Redistributing the software as a commercial product;
Licensing or sublicensing the software to third parties for commercial purposes;
Monetizing the software itself or a derivative version of it.

Organizations and companies are explicitly permitted to use this software internally for their own business operations, provided that they do not commercially distribute, sell, license, or monetize the software itself.

For commercial licensing or permission to commercially distribute or monetize the software, please contact the author.

**Copyright © 2026 Erick Paiva Silva. All rights reserved, except for the permissions explicitly granted above.**
