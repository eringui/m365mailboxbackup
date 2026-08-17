import argparse
import inspect
import json
import sys

from src.config.settings import validate_settings
from src.services.graph_service import GraphService
from src.services.mailbox_backup_service import MailboxBackupService
from src.services.operation_control import OperationInterrupted
from src.utils.logger import setup_logger, setup_report_logger

_report_logger = setup_report_logger()


def _report(message=""):
    _report_logger.info(message)


def call_method_compatible(method, **kwargs):
    """
    Chama métodos do serviço usando apenas argumentos aceitos pela assinatura atual.

    Isso evita erro caso algum arquivo de serviço ainda não tenha recebido
    parâmetros novos como resume_path ou excluded_folder_names.
    """
    signature = inspect.signature(method)
    accepted = {}

    for key, value in kwargs.items():
        if key in signature.parameters:
            accepted[key] = value

    return method(**accepted)

def build_excluded_folder_names(args):
    excluded_folder_names = list(args.exclude_folder or [])

    if args.profile_only:
        excluded_folder_names.extend(
            MailboxBackupService.DEFAULT_EXCLUDED_PROFILE_FOLDER_NAMES
        )

    return excluded_folder_names


def print_phase_0_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 0 — VALIDAÇÃO DO AMBIENTE")
    _report("=" * 70)

    _report(f"Mailbox testada: {result.get('mailbox')}")
    _report("")

    items = {
        "Autenticação Graph": result.get("auth", False),
        "Usuário localizado": result.get("user", False),
        "Pastas localizadas": result.get("folders", False),
        "Inbox localizada": result.get("inbox", False),
        "Mensagens consultadas": result.get("messages", False),
        "Calendário consultado": result.get("calendar", False),
        "Contatos consultados": result.get("contacts", False),
        "Tarefas consultadas": result.get("tasks", False),
    }

    for item, status in items.items():
        if status:
            _report(f"[OK] {item}")
        else:
            _report(f"[FALHA] {item}")

    _report("")

    if result.get("errors"):
        _report("ERROS ENCONTRADOS:")
        _report("-" * 70)

        for error in result.get("errors", []):
            _report(error)

        _report("")

    if all(items.values()):
        _report("[SUCESSO] Fase 0 validada com sucesso.")
        _report("Você pode prosseguir para a Fase 1 quando quiser.")
    else:
        _report("[ATENÇÃO] Fase 0 ainda não foi validada completamente.")
        _report("Corrija os erros acima antes de seguir para a Fase 1.")

    _report("=" * 70)


def print_phase_1_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 1 — GRAPHSERVICE")
    _report("=" * 70)

    if not result.get("success"):
        _report("[FALHA] Inspeção da mailbox não concluída.")

        for error in result.get("errors", []):
            _report(error)

        _report("=" * 70)
        return

    mailbox = result.get("mailbox", {})
    folders = result.get("folders", [])
    inbox = result.get("inbox", {})
    messages = result.get("messages_preview", [])

    _report("[OK] GraphService inicializado")
    _report("[OK] Autenticação Graph funcionando")
    _report("[OK] Consulta de usuário funcionando")
    _report("[OK] Consulta de pastas com paginação funcionando")
    _report("[OK] Localização da Inbox funcionando")
    _report("[OK] Listagem de mensagens funcionando")
    _report("")

    _report("MAILBOX")
    _report("-" * 70)
    _report(f"Nome: {mailbox.get('displayName')}")
    _report(f"Email: {mailbox.get('mail')}")
    _report(f"UPN: {mailbox.get('userPrincipalName')}")
    _report(f"Conta habilitada: {mailbox.get('accountEnabled')}")
    _report("")

    _report("INBOX")
    _report("-" * 70)
    _report(f"Nome: {inbox.get('displayName')}")
    _report(f"ID: {inbox.get('id')}")
    _report(f"Total de itens: {inbox.get('totalItemCount')}")
    _report(f"Itens não lidos: {inbox.get('unreadItemCount')}")
    _report("")

    _report("PASTAS ENCONTRADAS")
    _report("-" * 70)

    for folder in folders:
        _report(
            f"- {folder.get('displayName')} | "
            f"Total: {folder.get('totalItemCount')} | "
            f"Não lidos: {folder.get('unreadItemCount')}"
        )

    _report("")
    _report("PRÉVIA DE EMAILS")
    _report("-" * 70)

    if not messages:
        _report("Nenhuma mensagem retornada.")
    else:
        for message in messages:
            _report("")
            _report(f"Assunto: {message.get('subject')}")
            _report(f"Remetente: {message.get('from')}")
            _report(f"Recebido: {message.get('receivedDateTime')}")
            _report(f"Tem anexos: {message.get('hasAttachments')}")

    _report("")
    _report("[SUCESSO] Fase 1 concluída com sucesso.")
    _report("GraphService está pronto para ser usado nas próximas fases.")
    _report("=" * 70)


def print_phase_2_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 2 — EXPORTAÇÃO LOCAL INICIAL")
    _report("=" * 70)

    _report(f"Mailbox: {result.get('mailbox')}")
    _report(f"Pasta do backup: {result.get('backup_path')}")
    _report("")

    _report("RESUMO")
    _report("-" * 70)
    _report(f"Pastas exportadas/indexadas: {result.get('folders_count', 0)}")
    _report(f"Mensagens indexadas: {result.get('messages_indexed', 0)}")
    _report(f"Mensagens .eml exportadas: {result.get('messages_exported', 0)}")
    _report(f"Eventos de calendário: {result.get('calendar_events', 0)}")
    _report(f"Contatos: {result.get('contacts', 0)}")
    _report(f"Listas de tarefas: {result.get('task_lists', 0)}")
    _report(f"Tarefas: {result.get('task_items', 0)}")
    _report("")

    if result.get("errors"):
        _report("ERROS/AVISOS")
        _report("-" * 70)

        for error in result.get("errors", []):
            _report(error)

        _report("")

    if result.get("success"):
        _report("[SUCESSO] Fase 2 concluída com sucesso.")
        _report("Backup local inicial gerado.")
    else:
        _report("[ATENÇÃO] Fase 2 finalizada com erros.")
        _report("Revise o manifest.json e os logs.")

    _report("=" * 70)


def print_phase_3_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 3 — EXPORTAÇÃO POR ESTRUTURA DE PASTAS")
    _report("=" * 70)

    _report(f"Mailbox: {result.get('mailbox')}")
    _report(f"Pasta do backup: {result.get('backup_path')}")
    _report("")

    _report("RESUMO")
    _report("-" * 70)
    _report(f"Pastas encontradas: {result.get('folders_count', 0)}")
    _report(f"Pastas processadas: {result.get('folders_processed', 0)}")
    _report(f"Mensagens indexadas: {result.get('messages_indexed', 0)}")
    _report(f"Mensagens .eml exportadas: {result.get('messages_exported', 0)}")
    _report(f"Eventos de calendário: {result.get('calendar_events', 0)}")
    _report(f"Contatos: {result.get('contacts', 0)}")
    _report(f"Listas de tarefas: {result.get('task_lists', 0)}")
    _report(f"Tarefas: {result.get('task_items', 0)}")
    _report("")

    if result.get("errors"):
        _report("ERROS/AVISOS")
        _report("-" * 70)

        for error in result.get("errors", []):
            _report(error)

        _report("")

    if result.get("success"):
        _report("[SUCESSO] Fase 3 concluída com sucesso.")
        _report("Backup por estrutura de pastas gerado.")
    else:
        _report("[ATENÇÃO] Fase 3 finalizada com erros.")
        _report("Revise o manifest.json e os logs.")

    _report("=" * 70)


def print_phase_4_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 4 — EXPORTAÇÃO COM CHECKPOINT")
    _report("=" * 70)

    _report(f"Mailbox: {result.get('mailbox')}")
    _report(f"Pasta do backup: {result.get('backup_path')}")
    _report("")

    _report("RESUMO")
    _report("-" * 70)
    _report(f"Duração: {result.get('duration_seconds', 0)} segundos")
    _report(f"Pastas encontradas: {result.get('folders_count', 0)}")
    _report(f"Pastas processadas: {result.get('folders_processed', 0)}")
    _report(f"Pastas ignoradas/puladas: {result.get('folders_skipped', 0)}")
    _report(f"Mensagens indexadas: {result.get('messages_indexed', 0)}")
    _report(f"Mensagens .eml exportadas: {result.get('messages_exported', 0)}")
    _report(f"Mensagens ignoradas pelo checkpoint: {result.get('messages_skipped', 0)}")
    _report(f"Mensagens com falha: {result.get('messages_failed', 0)}")
    _report(f"Anexos exportados separadamente: {result.get('attachments_exported', 0)}")
    _report(f"Eventos de calendário: {result.get('calendar_events', 0)}")
    _report(f"Contatos: {result.get('contacts', 0)}")
    _report(f"Listas de tarefas: {result.get('task_lists', 0)}")
    _report(f"Tarefas: {result.get('task_items', 0)}")
    _report("")

    if result.get("errors"):
        _report("ERROS/AVISOS")
        _report("-" * 70)

        for error in result.get("errors", []):
            _report(error)

        _report("")

    if result.get("success"):
        _report("[SUCESSO] Fase 4 concluída com sucesso.")
        _report("Backup com checkpoint gerado.")
    else:
        _report("[ATENÇÃO] Fase 4 finalizada com erros.")
        _report("Revise manifest.json, checkpoint.json e logs.")

    _report("=" * 70)


def print_phase_5_result(result):
    _report("")
    _report("=" * 70)
    _report("FASE 5 — BACKUP EM LOTE VIA CSV")
    _report("=" * 70)

    _report(f"Arquivo CSV: {result.get('batch_file')}")
    _report(f"Total de mailboxes: {result.get('total_mailboxes', 0)}")
    _report(f"Processadas: {result.get('processed', 0)}")
    _report(f"Sucesso: {result.get('success_count', 0)}")
    _report(f"Falha: {result.get('failed_count', 0)}")
    _report(f"Duração total: {result.get('duration_seconds', 0)} segundos")
    _report("")

    _report("PRÉ-VALIDAÇÃO")
    _report("-" * 70)
    _report(f"Mailboxes válidas: {result.get('precheck_valid_count', 0)}")
    _report(f"Mailboxes inválidas: {result.get('precheck_invalid_count', 0)}")

    if result.get("precheck_report_path"):
        _report(f"Precheck JSON: {result.get('precheck_report_path')}")

    if result.get("precheck_csv_report_path"):
        _report(f"Precheck CSV: {result.get('precheck_csv_report_path')}")

    if result.get("report_path"):
        _report(f"Relatório JSON: {result.get('report_path')}")

    if result.get("csv_report_path"):
        _report(f"Relatório CSV: {result.get('csv_report_path')}")

    _report("")
    _report("RESULTADOS POR MAILBOX")
    _report("-" * 70)

    for item in result.get("results", []):
        status = "OK" if item.get("success") else "FALHA"

        _report("")
        _report(f"[{status}] {item.get('mailbox')}")
        _report(f"Backup: {item.get('backup_path')}")
        _report(f"Pastas: {item.get('folders_processed', 0)}")
        _report(f"Mensagens exportadas: {item.get('messages_exported', 0)}")
        _report(f"Mensagens com falha: {item.get('messages_failed', 0)}")
        _report(f"Anexos exportados: {item.get('attachments_exported', 0)}")

        if item.get("errors"):
            _report("Erros:")

            for error in item.get("errors", []):
                _report(f"- {error}")

    _report("")

    if result.get("success"):
        _report("[SUCESSO] Fase 5 concluída com sucesso.")
        _report("Backup em lote finalizado sem falhas.")
    else:
        _report("[ATENÇÃO] Fase 5 finalizada com uma ou mais falhas.")
        _report("Revise o relatório batch_report e os logs.")

    _report("=" * 70)


def main():
    parser = argparse.ArgumentParser(
        description="M365 Mailbox Backup"
    )

    parser.add_argument(
    "--skip-precheck",
    dest="skip_precheck",
    action="store_true",
    help="Ignora pré-validação do CSV e inicia o backup diretamente."
    )

    parser.add_argument(
        "--phase",
        required=False,
        choices=["0", "1", "2", "3", "4", "5"],
        help="Fase do projeto a executar. Use: --phase 0, --phase 1, --phase 2, --phase 3, --phase 4 ou --phase 5."
    )

    parser.add_argument(
        "--repair-failures", dest="repair_failures", action="store_true",
        help="Tenta novamente somente os EML ainda registrados como falha."
    )
    parser.add_argument(
        "--mailbox",
        required=False,
        help="E-mail da mailbox."
    )

    parser.add_argument(
        "--batch",
        required=False,
        help="Caminho do arquivo CSV para backup em lote. Coluna obrigatória: email."
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=25,
        help="Quantidade máxima de mensagens por pasta quando não estiver usando --all. Padrão: 25."
    )

    parser.add_argument(
        "--all",
        "--all-messages",
        dest="all_messages",
        action="store_true",
        help="Exporta todas as mensagens encontradas, sem aplicar limite por pasta."
    )

    parser.add_argument(
        "--attachments",
        "--export-attachments",
        dest="export_attachments",
        action="store_true",
        help="Exporta anexos separadamente além do arquivo .eml."
    )

    parser.add_argument(
        "--skip-calendar",
        dest="skip_calendar",
        action="store_true",
        help="Ignora exportação de calendário."
    )

    parser.add_argument(
        "--skip-contacts",
        dest="skip_contacts",
        action="store_true",
        help="Ignora exportação de contatos."
    )

    parser.add_argument(
        "--skip-tasks",
        dest="skip_tasks",
        action="store_true",
        help="Ignora exportação de tarefas."
    )

    parser.add_argument(
        "--resume-path",
        dest="resume_path",
        default=None,
        help="Caminho de um backup existente para retomar usando checkpoint.json."
    )

    parser.add_argument(
        "--profile-only",
        dest="profile_only",
        action="store_true",
        help="Exporta somente e-mails da mailbox principal, ignorando Archive/Arquivo Morto e pastas de sistema, se o serviço suportar."
    )

    parser.add_argument(
        "--job-options-file",
        dest="job_options_file",
        default=None,
        help="Arquivo JSON com escopo de pastas e opções específicas do trabalho."
    )

    parser.add_argument(
        "--exclude-folder",
        dest="exclude_folder",
        action="append",
        default=[],
        help="Nome de pasta a ignorar. Pode ser usado várias vezes. Exemplo: --exclude-folder \"Arquivo Morto\""
    )

    args = parser.parse_args()
    logger = setup_logger()

    job_options = {}
    if args.job_options_file:
        try:
            with open(args.job_options_file, "r", encoding="utf-8") as file:
                job_options = json.load(file)
        except Exception as error:
            logger.error(f"Não foi possível carregar as opções do trabalho: {error}")
            sys.exit(1)

    if not args.repair_failures and args.phase in ["0", "1", "2", "3", "4"] and not args.mailbox:
        logger.error("O argumento --mailbox é obrigatório para as fases 0, 1, 2, 3 e 4.")
        sys.exit(1)

    if not args.repair_failures and args.phase == "5" and not args.batch:
        logger.error("O argumento --batch é obrigatório para a fase 5.")
        sys.exit(1)

    logger.info("Iniciando M365 Mailbox Backup")
    logger.info(f"Fase informada: {args.phase}")

    if args.mailbox:
        logger.info(f"Mailbox informada: {args.mailbox}")

    if args.batch:
        logger.info(f"Arquivo batch informado: {args.batch}")

    if args.resume_path:
        logger.info(f"Retomada informada: {args.resume_path}")

    if args.profile_only:
        logger.info("Modo profile-only ativado.")

    if args.exclude_folder:
        logger.info(f"Pastas excluídas manualmente: {args.exclude_folder}")

    try:
        validate_settings()
    except Exception as error:
        logger.error(f"Configuração inválida: {error}")
        sys.exit(1)

    graph_service = GraphService(logger)
    excluded_folder_names = build_excluded_folder_names(args)

    if args.phase == "0":
        result = graph_service.validate_mailbox_access(
            args.mailbox
        )

        print_phase_0_result(result)

        if not all(
            [
                result.get("auth"),
                result.get("user"),
                result.get("folders"),
                result.get("inbox"),
                result.get("messages"),
                result.get("calendar"),
                result.get("contacts"),
                result.get("tasks")
            ]
        ):
            sys.exit(1)

    elif args.phase == "1":
        result = graph_service.inspect_mailbox(
            args.mailbox
        )

        print_phase_1_result(result)

        if not result.get("success"):
            sys.exit(1)

    elif args.phase == "2":
        backup_service = MailboxBackupService(
            graph_service=graph_service,
            logger=logger
        )

        result = call_method_compatible(
            backup_service.export_mailbox_local,
            mailbox_email=args.mailbox,
            message_limit=args.limit,
            export_all_messages=args.all_messages
        )

        print_phase_2_result(result)

        if not result.get("success"):
            sys.exit(1)

    elif args.phase == "3":
        backup_service = MailboxBackupService(
            graph_service=graph_service,
            logger=logger
        )

        result = call_method_compatible(
            backup_service.export_mailbox_by_folder,
            mailbox_email=args.mailbox,
            message_limit_per_folder=args.limit,
            export_all_messages=args.all_messages
        )

        print_phase_3_result(result)

        if not result.get("success"):
            sys.exit(1)

    elif args.phase == "4":
        backup_service = MailboxBackupService(
            graph_service=graph_service,
            logger=logger
        )

        result = call_method_compatible(
            backup_service.export_mailbox_complete,
            mailbox_email=args.mailbox,
            message_limit_per_folder=args.limit,
            export_all_messages=args.all_messages,
            export_attachments=args.export_attachments,
            skip_calendar=args.skip_calendar,
            skip_contacts=args.skip_contacts,
            skip_tasks=args.skip_tasks,
            resume_path=args.resume_path,
            phase=4,
            excluded_folder_names=excluded_folder_names,
            selected_folder_ids=job_options.get("selected_folder_ids"),
            selected_folder_paths=job_options.get("selected_folder_paths")
        )

        print_phase_4_result(result)

        if not result.get("success"):
            sys.exit(1)

    elif args.phase == "5":
        backup_service = MailboxBackupService(
            graph_service=graph_service,
            logger=logger
        )

        result = call_method_compatible(
            backup_service.export_batch,
            batch_path=args.batch,
            message_limit_per_folder=args.limit,
            export_all_messages=args.all_messages,
            export_attachments=args.export_attachments,
            skip_precheck=args.skip_precheck,
            skip_calendar=args.skip_calendar,
            skip_contacts=args.skip_contacts,
            skip_tasks=args.skip_tasks,
            resume_path=args.resume_path,
            excluded_folder_names=excluded_folder_names,
            selected_folder_ids=job_options.get("selected_folder_ids"),
            selected_folder_paths=job_options.get("selected_folder_paths")
        )

        print_phase_5_result(result)

        if not result.get("success"):
            sys.exit(1)


if __name__ == "__main__":
    main()