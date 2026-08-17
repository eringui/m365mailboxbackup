# M365 Mailbox Backup

A **Windows** tool that backs up **Microsoft 365** mailboxes using the **Microsoft Graph API**, saves everything as `.eml` files (with automatic resume if the process is interrupted), and optionally converts the backup into an Outlook `.pst` file.

Everything is controlled through a **PySide6** graphical user interface (GUI), with an operation queue, real-time progress, pause/resume support, and environment diagnostics.

> Application version: `3.0.0`

---

## ⚠️ Important Notice

This project was developed with **strong assistance from ChatGPT 5.6** — most of the code, resume logic, graphical interface, and PST conversion functionality was written together with the AI through an iterative process ("vibe coding").

Because of this:

* The project has **not undergone a professional security audit** or an extensive suite of automated tests.
* It may contain **bugs, unexpected behavior, incomplete screens, or inconsistent text** between different parts of the code.
* **Always test in a controlled environment** (for example, a test mailbox) before relying on it to back up important accounts.
* Use it at your own risk. See the [Known Issues and Limitations](#known-issues-and-limitations) section before using it in production.

---

## What the Program Does

The project has two independent stages:

1. **Backup (.eml)** — authenticates to Microsoft Graph as an application (without requiring users to sign in individually), iterates through folders, messages, calendars, contacts, and tasks from one or more mailboxes, and saves everything to `.eml` files on disk. If the process is interrupted (power outage, network failure, manual cancellation, etc.), it can resume from where it stopped.
2. **PST Conversion** *(optional)* — takes an existing `.eml` backup and, using COM automation through **Outlook Classic** installed on the machine, creates a browsable `.pst` file while preserving the folder structure, sender/recipient information, dates, and attachments.

The graphical interface **does not perform this heavy work by itself**. When opened, it starts a local service (the "coordinator", built with FastAPI and running on `127.0.0.1:8765`) that manages the operation queue and launches separate processes for each backup or conversion. This allows you to close and reopen the GUI without losing the progress of operations that are already running.

---

## How the Graphical Interface Works

When you open the program (`python m365_backup_gui.py` or the executable), you will see a sidebar menu with 4 pages:

### 📥 Backups

The main page, where the mailbox download queue is managed.

* **Add** — adds a mailbox by email address.
* **Import CSV** — imports multiple mailboxes at once (a file with an `email` column; see `users-example.csv`).
* **Resume Existing** — points to an existing backup folder and resumes the process from the saved checkpoint without recounting everything from scratch.
* **Configure Folders** — opens a dialog where you can choose which mail folders should be downloaded, with options to export all messages, export attachments separately, skip calendar/contacts/tasks, limit messages per folder, and apply the configuration to the entire queue.
* **Start Backups**, **Pause**, **Resume**, **Remove**.
* **Fix Failures** — instead of reprocessing the entire mailbox, retries **only the emails that failed** during the backup, using the saved checkpoint.
* The table displays status, current stage, mailbox email count, progress, remaining items, downloaded size, estimated time, selected folders, and queue position. Queue order can be rearranged by dragging with the mouse.
* A lower panel displays the real-time log for the mailbox selected in the table.

### 📦 PST Conversions

Where you convert a completed `.eml` backup into an Outlook `.pst` file.

* **New Conversion** opens a dialog with three tabs:

  * **General** — source folder (the backup), destination folder and `.pst` filename, display name inside Outlook, and what to do if the PST already exists (resume from checkpoint, create a numbered name, replace it, or cancel).
  * **Content** — preserve the original folder tree or consolidate everything into a single folder, display original metadata in the message body, import attachments/embedded images, and set the maximum image size.
  * **Performance & Security** — import speed, verification level (fast, balanced, or full), performance profile, automatic worker adjustment, memory budget, and what qualifies as a "large" email.
* The table displays source, destination, Outlook name, conflict policy, progress, verification status, failures, and remaining items.
* **This feature requires Outlook Classic to be installed and configured on the machine** — `.eml` backup works without Outlook.

### 🧾 Logs

A real-time log viewer directly inside the interface, without having to manually open the `logs/` folder.

* Switch between the general operational log and the application failure log.
* Filter by level (DEBUG/INFO/WARNING/ERROR/CRITICAL) and search text.
* Pause automatic scrolling without stopping file logging, clear the view, and open the logs folder.

### ⚙️ Settings

Contains the application's preferences, organized into tabs:

* **General** — backup and PST destination folders, number of simultaneous backups and PST conversions, and default content options (export everything, skip calendar/contacts/tasks).
* **Credentials & Graph** — Tenant ID, Client ID, and Client Secret for the application registered in Microsoft Entra ID. The **Save and Test Credentials** button validates authentication immediately.
* **Appearance & PST** — theme (automatic/light/dark), font size, and whether the PST should be automatically removed from Outlook after completion.
* There are also buttons to **export/import a settings profile** (JSON), **run diagnostics**, and **validate the integrity** of an existing backup.

> There is also a **"Performance & Storage"** tab (download limits, rate/throttling limiter, page size, disk space, etc.) already implemented in the code, but it is **hidden by default** in the current interface version — see the limitations section below.

### Keyboard Shortcuts

`Ctrl+N` add mailbox · `Ctrl+I` import CSV · `Ctrl+O` resume existing backup · `Ctrl+Enter` start queue · `Ctrl+P` / `Ctrl+R` pause/resume selection · `Delete` remove selection · `F5` refresh page · `Ctrl+1..4` switch pages · `Ctrl+,` go to Settings · `Ctrl+L` go to Logs.

---

## Requirements

* **Windows 10/11.** The project depends on `pywin32` and Outlook COM; it does not work on macOS/Linux.
* **Outlook Classic** (the traditional desktop application) installed with a configured profile — required **only** for `.pst` conversion. `.eml` backup works without Outlook installed. The *new Outlook for Windows* is not compatible with the automation used here.
* **Python 3.11+** to run from source code (or use the PyInstaller executables, if available).
* An **application (App Registration)** registered in Microsoft Entra ID, with application permissions (not delegated permissions) in Microsoft Graph and admin consent — for example `User.Read.All`, `Mail.Read`, `Calendars.Read`, `Contacts.Read`, and `Tasks.Read`.
* Network access to `login.microsoftonline.com` and `graph.microsoft.com`.

---

## Quick Installation

```bash
git clone <repository-url> m365-mailbox-backup
cd m365-mailbox-backup

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

Credentials (`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`) can be configured in two ways:

* Create a `.env` file in the project root; **or**
* Enter them directly under **Settings → Credentials & Graph** in the GUI, which stores the secret securely using **Windows DPAPI** (no `.env` file required).

---

## First Use — Step by Step

```bash
python m365_backup_gui.py
```

1. When opening the application for the first time, go to **Settings → Credentials & Graph**, enter the Tenant ID, Client ID, and Client Secret, then click **Save and Test Credentials**.
2. Return to **Backups** and add a mailbox (**Add**, **Import CSV**, or **Resume Existing** if a partial backup already exists).
3. Optionally, use **Configure Folders** to choose what should be downloaded from the mailbox.
4. Click **Start Backups** and monitor the progress in the table and the lower log panel.
5. Once the backup is complete, go to **PST Conversions → New Conversion** to generate a `.pst` file if you need to open the backup in Outlook.

---

## Resume, Checkpoint, and Failure Recovery

Each backup maintains a transactional checkpoint (`checkpoint.json` + SQLite database) inside the backup folder, recording folders, pages, and items that have already been exported.

* If the process is interrupted for any reason, use **Resume Existing** and point it to the same folder — items already confirmed in the checkpoint will not be downloaded again.
* If only specific emails failed (for example, due to a temporary network error), use **Fix Failures** instead of restarting everything. The program retries only the items marked as failed in the checkpoint.

---

## PST Conversion Summary

The conversion uses Outlook Classic COM automation (`Outlook.Application`) to create a **Unicode-format PST**, preserving folders, sender/recipient information, body, dates, and attachments. The program checks available disk space before and during writing and provides three verification levels (fast, balanced, and full) to verify that everything was actually written.

> Outlook has a default **50 GB** limit for Unicode `.pst` files. If the backup grows beyond that, you need to increase `MaxLargeFileSize`/`WarnLargeFileSize` in the Windows Registry — this is an Outlook limitation, not a limitation of the tool itself.

---

## Known Issues and Limitations

* The **"Performance & Storage"** tab (rate limiter, workers, page size, disk alerts) already exists in the Settings code but is **disabled/hidden** in the current interface — internal default values are used until the tab is re-enabled in a future version.
* There is also a **Graph API call metrics** screen (latency, throttling, API health) implemented in the backend, but it is **not connected to the navigation menu** in the current GUI version, so it is not visible to users.
* The application works **only on Windows**; PST conversion depends on Outlook Classic (it does not work with the "new Outlook").
* Since a large portion of the code was generated with AI through an iterative process, there may be **minor inconsistencies** between texts, variable names, or behavior across different screens — check the logs if in doubt.
* Authentication failures usually indicate incorrect `TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET` values or pending administrator consent in Entra ID.
* A large number of `429` errors (Graph request throttling) or slow performance may indicate excessive parallelism for the tenant — reduce concurrency in Settings.
* If the GUI cannot connect to the local coordinator, port `8765` (coordinator) or `48765` (single-instance lock) may already be in use by another instance of the program running in the background.
* If an operation pauses automatically, check the available disk space in the backup destination.
* There is no comprehensive automated test suite; behavioral changes between versions may not be fully documented.

If you encounter an issue not listed here, check the **Logs** tab and run **Settings → Run Diagnostics** before opening an issue.

---

## Security and Privacy

* Never commit the `.env` file (it is already included in `.gitignore`) or the App Registration Client Secret.
* Prefer saving credentials through the GUI, which protects them using **DPAPI**, tied to the Windows user/machine.
* Follow the principle of least privilege when configuring the App Registration: grant only the permissions required and consider restricting application access to specific mailboxes using an *Application Access Policy* in Exchange Online.
* The `logs/`, `_gui_state/`, and `output/` folders may contain personal/sensitive data from real mailboxes — they are already listed in `.gitignore` and should not be shared.

---

## About This Project

This is a personal/internal project developed with the assistance of **ChatGPT 5.6** as a programming partner. It is not an official commercial product, does not include guaranteed support, and may undergo significant behavioral changes between versions. Feel free to review, test, and adapt the code to your environment before using it with important data.

---

## License

This project is publicly available for personal, educational, organizational, and non-commercial use.

You may use, study, modify, and share the original or modified source code free of charge, including within organizations for internal use. **You may not sell, license, sublicense, or commercially distribute this software or derivative versions without explicit written permission from the author.**

**Copyright © 2026 Erick Paiva Silva. All rights reserved, except for the permissions granted above.**
