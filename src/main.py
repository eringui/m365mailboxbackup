import argparse
import inspect
import json
import sys

from src.config.settings import validate_settings
from src.services.graph_service import GraphService
from src.services.mailbox_backup_service import MailboxBackupService
from src.utils.logger import setup_logger


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
            [
                "Arquivo Morto",
                "Archive",
                "Online Archive",
                "In-Place Archive",
                "Recoverable Items",
                "Deletions",
                "Purges",
                "Versions",
                "Audits",
                "DiscoveryHolds",
                "Sync Issues",
                "Problemas de Sincronização",
                "Conversation History",
                "Histórico de Conversas",
                "RSS Feeds",
                "Feeds RSS"
            ]
        )

    return excluded_folder_names


def print_phase_0_result(result):
    print("")
    print("=" * 70)
    print("FASE 0 — VALIDAÇÃO DO AMBIENTE")
    print("=" * 70)

    print(f"Mailbox testada: {result.get('mailbox')}")
    print("")

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
            print(f"[OK] {item}")
        else:
            print(f"[FALHA] {item}")

    print("")

    if result.get("errors"):
        print("ERROS ENCONTRADOS:")
        print("-" * 70)

        for error in result.get("errors", []):
            print(error)

        print("")

    if all(items.values()):
        print("[SUCESSO] Fase 0 validada com sucesso.")
        print("Você pode prosseguir para a Fase 1 quando quiser.")
    else:
        print("[ATENÇÃO] Fase 0 ainda não foi validada completamente.")
        print("Corrija os erros acima antes de seguir para a Fase 1.")

    print("=" * 70)


def print_phase_1_result(result):
    print("")
    print("=" * 70)
    print("FASE 1 — GRAPHSERVICE")
    print("=" * 70)

    if not result.get("success"):
        print("[FALHA] Inspeção da mailbox não concluída.")

        for error in result.get("errors", []):
            print(error)

        print("=" * 70)
        return

    mailbox = result.get("mailbox", {})
    folders = result.get("folders", [])
    inbox = result.get("inbox", {})
    messages = result.get("messages_preview", [])

    print("[OK] GraphService inicializado")
    print("[OK] Autenticação Graph funcionando")
    print("[OK] Consulta de usuário funcionando")
    print("[OK] Consulta de pastas com paginação funcionando")
    print("[OK] Localização da Inbox funcionando")
    print("[OK] Listagem de mensagens funcionando")
    print("")

    print("MAILBOX")
    print("-" * 70)
    print(f"Nome: {mailbox.get('displayName')}")
    print(f"Email: {mailbox.get('mail')}")
    print(f"UPN: {mailbox.get('userPrincipalName')}")
    print(f"Conta habilitada: {mailbox.get('accountEnabled')}")
    print("")

    print("INBOX")
    print("-" * 70)
    print(f"Nome: {inbox.get('displayName')}")
    print(f"ID: {inbox.get('id')}")
    print(f"Total de itens: {inbox.get('totalItemCount')}")
    print(f"Itens não lidos: {inbox.get('unreadItemCount')}")
    print("")

    print("PASTAS ENCONTRADAS")
    print("-" * 70)

    for folder in folders:
        print(
            f"- {folder.get('displayName')} | "
            f"Total: {folder.get('totalItemCount')} | "
            f"Não lidos: {folder.get('unreadItemCount')}"
        )

    print("")
    print("PRÉVIA DE EMAILS")
    print("-" * 70)

    if not messages:
        print("Nenhuma mensagem retornada.")
    else:
        for message in messages:
            print("")
            print(f"Assunto: {message.get('subject')}")
            print(f"Remetente: {message.get('from')}")
            print(f"Recebido: {message.get('receivedDateTime')}")
            print(f"Tem anexos: {message.get('hasAttachments')}")

    print("")
    print("[SUCESSO] Fase 1 concluída com sucesso.")
    print("GraphService está pronto para ser usado nas próximas fases.")
    print("=" * 70)


def print_phase_2_result(result):
    print("")
    print("=" * 70)
    print("FASE 2 — EXPORTAÇÃO LOCAL INICIAL")
    print("=" * 70)

    print(f"Mailbox: {result.get('mailbox')}")
    print(f"Pasta do backup: {result.get('backup_path')}")
    print("")

    print("RESUMO")
    print("-" * 70)
    print(f"Pastas exportadas/indexadas: {result.get('folders_count', 0)}")
    print(f"Mensagens indexadas: {result.get('messages_indexed', 0)}")
    print(f"Mensagens .eml exportadas: {result.get('messages_exported', 0)}")
    print(f"Eventos de calendário: {result.get('calendar_events', 0)}")
    print(f"Contatos: {result.get('contacts', 0)}")
    print(f"Listas de tarefas: {result.get('task_lists', 0)}")
    print(f"Tarefas: {result.get('task_items', 0)}")
    print("")

    if result.get("errors"):
        print("ERROS/AVISOS")
        print("-" * 70)

        for error in result.get("errors", []):
            print(error)

        print("")

    if result.get("success"):
        print("[SUCESSO] Fase 2 concluída com sucesso.")
        print("Backup local inicial gerado.")
    else:
        print("[ATENÇÃO] Fase 2 finalizada com erros.")
        print("Revise o manifest.json e os logs.")

    print("=" * 70)


def print_phase_3_result(result):
    print("")
    print("=" * 70)
    print("FASE 3 — EXPORTAÇÃO POR ESTRUTURA DE PASTAS")
    print("=" * 70)

    print(f"Mailbox: {result.get('mailbox')}")
    print(f"Pasta do backup: {result.get('backup_path')}")
    print("")

    print("RESUMO")
    print("-" * 70)
    print(f"Pastas encontradas: {result.get('folders_count', 0)}")
    print(f"Pastas processadas: {result.get('folders_processed', 0)}")
    print(f"Mensagens indexadas: {result.get('messages_indexed', 0)}")
    print(f"Mensagens .eml exportadas: {result.get('messages_exported', 0)}")
    print(f"Eventos de calendário: {result.get('calendar_events', 0)}")
    print(f"Contatos: {result.get('contacts', 0)}")
    print(f"Listas de tarefas: {result.get('task_lists', 0)}")
    print(f"Tarefas: {result.get('task_items', 0)}")
    print("")

    if result.get("errors"):
        print("ERROS/AVISOS")
        print("-" * 70)

        for error in result.get("errors", []):
            print(error)

        print("")

    if result.get("success"):
        print("[SUCESSO] Fase 3 concluída com sucesso.")
        print("Backup por estrutura de pastas gerado.")
    else:
        print("[ATENÇÃO] Fase 3 finalizada com erros.")
        print("Revise o manifest.json e os logs.")

    print("=" * 70)


def print_phase_4_result(result):
    print("")
    print("=" * 70)
    print("FASE 4 — EXPORTAÇÃO COM CHECKPOINT")
    print("=" * 70)

    print(f"Mailbox: {result.get('mailbox')}")
    print(f"Pasta do backup: {result.get('backup_path')}")
    print("")

    print("RESUMO")
    print("-" * 70)
    print(f"Duração: {result.get('duration_seconds', 0)} segundos")
    print(f"Pastas encontradas: {result.get('folders_count', 0)}")
    print(f"Pastas processadas: {result.get('folders_processed', 0)}")
    print(f"Pastas ignoradas/puladas: {result.get('folders_skipped', 0)}")
    print(f"Mensagens indexadas: {result.get('messages_indexed', 0)}")
    print(f"Mensagens .eml exportadas: {result.get('messages_exported', 0)}")
    print(f"Mensagens ignoradas pelo checkpoint: {result.get('messages_skipped', 0)}")
    print(f"Mensagens com falha: {result.get('messages_failed', 0)}")
    print(f"Anexos exportados separadamente: {result.get('attachments_exported', 0)}")
    print(f"Eventos de calendário: {result.get('calendar_events', 0)}")
    print(f"Contatos: {result.get('contacts', 0)}")
    print(f"Listas de tarefas: {result.get('task_lists', 0)}")
    print(f"Tarefas: {result.get('task_items', 0)}")
    print("")

    if result.get("errors"):
        print("ERROS/AVISOS")
        print("-" * 70)

        for error in result.get("errors", []):
            print(error)

        print("")

    if result.get("success"):
        print("[SUCESSO] Fase 4 concluída com sucesso.")
        print("Backup com checkpoint gerado.")
    else:
        print("[ATENÇÃO] Fase 4 finalizada com erros.")
        print("Revise manifest.json, checkpoint.json e logs.")

    print("=" * 70)


def print_phase_5_result(result):
    print("")
    print("=" * 70)
    print("FASE 5 — BACKUP EM LOTE VIA CSV")
    print("=" * 70)

    print(f"Arquivo CSV: {result.get('batch_file')}")
    print(f"Total de mailboxes: {result.get('total_mailboxes', 0)}")
    print(f"Processadas: {result.get('processed', 0)}")
    print(f"Sucesso: {result.get('success_count', 0)}")
    print(f"Falha: {result.get('failed_count', 0)}")
    print(f"Duração total: {result.get('duration_seconds', 0)} segundos")
    print("")

    print("PRÉ-VALIDAÇÃO")
    print("-" * 70)
    print(f"Mailboxes válidas: {result.get('precheck_valid_count', 0)}")
    print(f"Mailboxes inválidas: {result.get('precheck_invalid_count', 0)}")

    if result.get("precheck_report_path"):
        print(f"Precheck JSON: {result.get('precheck_report_path')}")

    if result.get("precheck_csv_report_path"):
        print(f"Precheck CSV: {result.get('precheck_csv_report_path')}")

    if result.get("report_path"):
        print(f"Relatório JSON: {result.get('report_path')}")

    if result.get("csv_report_path"):
        print(f"Relatório CSV: {result.get('csv_report_path')}")

    print("")
    print("RESULTADOS POR MAILBOX")
    print("-" * 70)

    for item in result.get("results", []):
        status = "OK" if item.get("success") else "FALHA"

        print("")
        print(f"[{status}] {item.get('mailbox')}")
        print(f"Backup: {item.get('backup_path')}")
        print(f"Pastas: {item.get('folders_processed', 0)}")
        print(f"Mensagens exportadas: {item.get('messages_exported', 0)}")
        print(f"Mensagens com falha: {item.get('messages_failed', 0)}")
        print(f"Anexos exportados: {item.get('attachments_exported', 0)}")

        if item.get("errors"):
            print("Erros:")

            for error in item.get("errors", []):
                print(f"- {error}")

    print("")

    if result.get("success"):
        print("[SUCESSO] Fase 5 concluída com sucesso.")
        print("Backup em lote finalizado sem falhas.")
    else:
        print("[ATENÇÃO] Fase 5 finalizada com uma ou mais falhas.")
        print("Revise o relatório batch_report e os logs.")

    print("=" * 70)


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
        required=True,
        choices=["0", "1", "2", "3", "4", "5"],
        help="Fase do projeto a executar. Use: --phase 0, --phase 1, --phase 2, --phase 3, --phase 4 ou --phase 5."
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
    job_options = {}
    if args.job_options_file:
        try:
            with open(args.job_options_file, "r", encoding="utf-8") as file:
                job_options = json.load(file)
        except Exception as error:
            print(f"[ERRO] Não foi possível carregar as opções do trabalho: {error}")
            sys.exit(1)

    if args.phase in ["0", "1", "2", "3", "4"] and not args.mailbox:
        print("[ERRO] O argumento --mailbox é obrigatório para as fases 0, 1, 2, 3 e 4.")
        sys.exit(1)

    if args.phase == "5" and not args.batch:
        print("[ERRO] O argumento --batch é obrigatório para a fase 5.")
        sys.exit(1)

    logger = setup_logger()

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
        logger.error(error)
        print(f"[ERRO] Configuração inválida: {error}")
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