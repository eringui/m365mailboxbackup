# M365 Mailbox Backup (GPT 5.6)

Ferramenta para Windows que faz backup de caixas de correio do **Microsoft 365** via **Microsoft Graph API**, exporta as mensagens para **.eml** com retomada por checkpoint, e opcionalmente converte o backup para **.pst** usando automação COM do **Outlook Classic**. Possui interface gráfica (PySide6) com fila de operações e progresso em tempo real, além de uma CLI completa para uso via linha de comando ou automação/agendamento.

> Versão do aplicativo: `2.0.0`

## Sumário

- [Visão geral](#visão-geral)
- [Principais recursos](#principais-recursos)
- [Arquitetura](#arquitetura)
- [Requisitos](#requisitos)
- [Registro do aplicativo no Microsoft Entra ID](#registro-do-aplicativo-no-microsoft-entra-id)
- [Instalação](#instalação)
- [Configuração (.env)](#configuração-env)
- [Uso — Interface gráfica](#uso--interface-gráfica)
- [Capturas da interface](#capturas-da-interface)
- [Uso — Linha de comando](#uso--linha-de-comando)
- [Backup em lote (CSV)](#backup-em-lote-csv)
- [Retomada e checkpoints](#retomada-e-checkpoints)
- [Exportação para PST (Outlook Classic)](#exportação-para-pst-outlook-classic)
- [Aumentando o limite de tamanho do .pst no Windows (registro)](#aumentando-o-limite-de-tamanho-do-pst-no-windows-registro)
- [Limitação de taxa e throttling adaptativo](#limitação-de-taxa-e-throttling-adaptativo)
- [Logs, métricas e diagnóstico](#logs-métricas-e-diagnóstico)
- [Build / empacotamento (PyInstaller)](#build--empacotamento-pyinstaller)
- [Solução de problemas](#solução-de-problemas)
- [Segurança](#segurança)
- [Licença](#licença)

## Visão geral

O projeto tem dois estágios independentes:

1. **Backup** — autentica no Microsoft Graph como aplicativo (client credentials, sem login interativo por usuário), percorre pastas, mensagens, calendário, contatos e tarefas de uma ou mais mailboxes e exporta tudo para arquivos `.eml` em disco, com checkpoint transacional para permitir retomada.
2. **Conversão para PST** *(opcional)* — lê o backup `.eml` já existente e, usando automação COM do Outlook Classic instalado na máquina, monta um arquivo `.pst` navegável, preservando estrutura de pastas, remetente/destinatário, datas e anexos.

A interface gráfica não executa esses estágios diretamente: ela inicia, na primeira execução, um serviço HTTP local (**coordenador**, FastAPI em `127.0.0.1:8765`) que gerencia uma fila de operações e dispara processos filho (CLI) para cada backup/conversão, persistindo o estado em SQLite. Isso permite fechar e reabrir a GUI sem perder o progresso de operações em andamento.

## Principais recursos

- **Autenticação app-only (MSAL)** — um único aplicativo no Entra ID acessa qualquer mailbox do tenant, sem exigir login individual de cada usuário.
- **Exportação incremental** — usa *delta queries* do Microsoft Graph (`/messages/delta`) para que reexecuções busquem apenas o que mudou.
- **Checkpoint transacional (SQLite/WAL)** — cada item exportado é registrado; interrupções (queda de rede, cancelamento, PC desligado) podem ser retomadas com `--resume-path`, sem reprocessar o que já foi salvo.
- **Rate limiting preventivo + throttling adaptativo** — limites configuráveis por processo, por mailbox e compartilhados entre processos (via SQLite), reduzindo automaticamente o paralelismo quando o Graph responde `429`.
- **Backup em lote via CSV** — uma mailbox por linha, com pré-validação e relatório de resultados em JSON/CSV.
- **Conversão para PST com verificação** — grava em lotes, reconcilia itens pendentes e valida o que foi realmente persistido no `.pst`.
- **Interface gráfica (PySide6)** — fila de operações, progresso em tempo real, pausar/retomar/cancelar, ajuste de concorrência, aba de métricas de API e diagnóstico de ambiente.
- **Credenciais protegidas por DPAPI** no Windows, com fallback para `.env`.
- **Telemetria de chamadas ao Graph** (latência, taxa de throttle, sucesso/falha) em SQLite, exportável em CSV.

## Arquitetura

> Estrutura inferida a partir dos imports do projeto (`from src.config...`, `from src.services...`, `python -m src.main`).

```
.
├── .env                        # credenciais e overrides locais (NÃO versionar)
├── .gitignore
├── requirements.txt
├── application_runtime.py      # AppSettings, CredentialStore (DPAPI), CrashReporter, diagnóstico
├── m365_backup_gui.py           # ponto de entrada oficial da GUI
├── m365_backup_gui_qt.py        # implementação da interface gráfica (PySide6)
├── m365_backup_coordinator.py   # serviço local FastAPI (127.0.0.1:8765) que orquestra operações
├── logs/                        # gerado em runtime (rotativo)
├── output/
│   ├── backups/                 # saída padrão dos backups .eml
│   └── pst/                     # saída padrão dos .pst
├── _gui_state/                  # settings.json, credentials.bin, operations.sqlite3, rate limiter db
└── src/
    ├── main.py                  # CLI — fases 0 a 5
    ├── config/
    │   └── settings.py          # leitura do .env e parâmetros com valores padrão
    ├── services/
    │   ├── graph_service.py         # autenticação MSAL + chamadas ao Microsoft Graph
    │   ├── mailbox_backup_service.py   # orquestra a exportação .eml (fases 2 a 5)
    │   ├── pst_export_service.py       # converte .eml em .pst via COM do Outlook
    │   ├── checkpoint_store.py         # checkpoint transacional (SQLite/WAL)
    │   ├── operation_store.py          # fila/estado das operações do coordenador (SQLite/WAL)
    │   └── api_metrics_store.py        # telemetria de chamadas ao Graph (SQLite/WAL)
    └── utils/
        └── logger.py             # logging rotativo + logger dedicado por operação PST
```

**Fluxo típico (GUI):** `m365_backup_gui.py` → inicia `m365_backup_coordinator.py` como subprocesso → coordenador registra a operação em `operations.sqlite3` e dispara `python -m src.main --phase 4 ...` (backup) ou `python -m src.services.pst_export_service ...` (conversão) como novo processo → progresso é lido do `stdout` (linhas `[PROGRESS]`) e transmitido à GUI por *polling*/WebSocket.

Quando empacotado com PyInstaller, os mesmos pontos de entrada viram executáveis: a GUI, `m365_backup_coordinator.exe`, `run_backend.exe` (equivalente a `src.main`) e `run_pst.exe` (equivalente a `src.services.pst_export_service`).

## Requisitos

- **Windows 10/11.** O projeto depende de `pywin32` e do Outlook via COM; não funciona em macOS/Linux.
- **Outlook Classic** (aplicativo de mesa tradicional) instalado e com um perfil configurado — necessário **apenas** para a etapa de conversão para `.pst`. O backup em `.eml` funciona sem o Outlook instalado. O *novo Outlook para Windows* não é compatível com a automação COM usada aqui.
- **Python 3.11+** para executar a partir do código-fonte (ou use os executáveis gerados pelo PyInstaller).
- Um **aplicativo (App Registration)** no Microsoft Entra ID com permissões de aplicativo (app-only) para o Microsoft Graph e consentimento de administrador.
- Acesso de rede a `login.microsoftonline.com` e `graph.microsoft.com`.

## Registro do aplicativo no Microsoft Entra ID

A autenticação usa o fluxo *client credentials* (app-only, via MSAL) e acessa as mailboxes através de `/users/{mailbox}/...`. Isso exige **permissões de aplicativo** (não delegadas):

| Permissão do Graph | Tipo | Uso no projeto |
|---|---|---|
| `User.Read.All` | Aplicativo | Resolver a mailbox/usuário (`/users/{mailbox}`) |
| `Mail.Read` | Aplicativo | Ler pastas e mensagens (`/mailFolders`, `/messages`, `/messages/delta`) |
| `Calendars.Read` | Aplicativo | Exportar eventos de calendário |
| `Contacts.Read` | Aplicativo | Exportar contatos |
| `Tasks.Read` | Aplicativo | Exportar listas de tarefas (To Do) |

Passo a passo:

1. No [portal do Azure](https://portal.azure.com), acesse **Microsoft Entra ID → Registros de aplicativo → Novo registro**.
2. Dê um nome (ex.: `M365 Mailbox Backup`) e mantenha o tipo de conta como "somente este diretório organizacional".
3. Anote o **ID do aplicativo (cliente)** e o **ID do diretório (locatário)** — serão `CLIENT_ID` e `TENANT_ID`.
4. Em **Certificados e segredos**, crie um **Novo segredo do cliente** e copie o valor imediatamente (ele não é exibido novamente) — será `CLIENT_SECRET`.
5. Em **Permissões de API → Adicionar uma permissão → Microsoft Graph → Permissões de aplicativo**, adicione as permissões da tabela acima.
6. Clique em **Conceder consentimento de administrador** para o locatário.
7. *(Recomendado)* Restrinja o alcance do aplicativo a mailboxes específicas usando uma [Application Access Policy](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-configure-application-access-policy) do Exchange Online (`New-ApplicationAccessPolicy`), evitando que o app tenha acesso irrestrito a todo o tenant.

## Instalação

```bash
git clone <url-do-repositório> m365-mailbox-backup
cd m365-mailbox-backup

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

## Configuração (`.env`)

O arquivo `_env-example` incluído no repositório está vazio — use o modelo abaixo como referência e salve como `.env` na raiz do projeto (ou aponte a variável `M365_BACKUP_ENV_PATH` para outro caminho).

```env
# Credenciais do App Registration (Microsoft Entra ID) — obrigatórias
TENANT_ID=00000000-0000-0000-0000-000000000000
CLIENT_ID=00000000-0000-0000-0000-000000000000
CLIENT_SECRET=coloque-o-segredo-aqui

# Opcionais — já possuem valores padrão em src/config/settings.py
GRAPH_SCOPE=https://graph.microsoft.com/.default
GRAPH_URL=https://graph.microsoft.com/v1.0
M365_BACKUP_OUTPUT_ROOT=./output/backups
```

> Na GUI, as credenciais também podem ser salvas de forma protegida por **DPAPI** (`_gui_state/credentials.bin`) pela tela de configurações, sem precisar de um `.env`.

Algumas variáveis avançadas (todas opcionais, com valores padrão sensatos):

| Variável | Padrão | Descrição |
|---|---|---|
| `M365_EML_DOWNLOAD_WORKERS` | `16` | Threads de download paralelo de `.eml` |
| `M365_MIME_MAX_CONCURRENCY` | `3` | Downloads MIME simultâneos por mailbox |
| `M365_RATE_LIMITER_ENABLED` | `1` | Ativa o limitador preventivo (pyrate-limiter) |
| `M365_GLOBAL_MIME_RATE_MINUTE` | `240` | Limite global de chamadas MIME por minuto |
| `M365_ADAPTIVE_THROTTLING` | `1` | Reduz paralelismo automaticamente após `429` |
| `M365_PST_PREPARE_WORKERS` | `3` | Threads de preparação de itens antes da gravação COM (serial) |
| `M365_PST_CAPACITY_PREFLIGHT` | `1` | Verifica espaço em disco antes de gravar o PST |
| `M365_DISK_WARNING_GB` / `M365_DISK_CRITICAL_GB` | `50` / `10` | Limiares de espaço livre em disco |

## Uso — Interface gráfica

```bash
python m365_backup_gui.py
```

Ao abrir, a GUI:

1. Garante instância única (via socket local na porta `48765`) e inicia o coordenador local (`127.0.0.1:8765`) se ainda não estiver rodando.
2. Na primeira execução, solicita as credenciais do Entra ID (ou lê o `.env`).
3. Permite adicionar mailboxes individualmente ou importar um CSV para backup em lote.
4. Permite criar, a partir de um backup já concluído, uma operação de **conversão para PST**.
5. Exibe fila de operações com progresso em tempo real, pausa/retomada/cancelamento e reordenação da fila.
6. Permite ajustar a concorrência (quantos backups e quantas conversões PST rodam ao mesmo tempo).
7. Traz uma aba de **métricas de API** (latência, throttles, requisições/seg) e um diálogo **"Sobre e diagnóstico"** que valida credenciais, destinos de escrita, espaço em disco, SQLite e disponibilidade do Outlook Classic.

## Capturas da interface

As imagens abaixo ficam na mesma pasta dos arquivos `README.md` e `README_EN.md`. Dessa forma, o GitHub, o VS Code e outros visualizadores compatíveis com Markdown conseguem exibi-las usando caminhos relativos.

### Fila de backups EML

<p align="center">
<img width="1920" height="1041" alt="aba backup  eml" src="https://github.com/user-attachments/assets/2345cbb5-3f6b-4ba5-859a-b7d0c055d27a" />
</p>

A página **Backups** centraliza a fila das mailboxes, importação por CSV, retomada de backups existentes, configuração de pastas, pausa, retomada e remoção. A tabela apresenta status, etapa atual, quantidade informada pela mailbox, progresso, itens restantes, volume transferido, estimativa de conclusão, pastas selecionadas e posição na fila. O painel inferior exibe o log da mailbox selecionada.

### Seleção de pastas e opções do backup

<p align="center">
<img width="1074" height="777" alt="configuração de pasta" src="https://github.com/user-attachments/assets/c22ba9dc-271d-4de0-a95d-b3183d6777f0" />
</p>

A janela **Pastas e opções** carrega somente a estrutura e os contadores das pastas antes do backup. É possível selecionar todas as pastas, limpar a seleção, aplicar uma seleção recomendada, definir opções de conteúdo, ignorar calendário, contatos e tarefas, limitar mensagens por pasta e aplicar a configuração à fila inteira.

> Os contadores exibidos nessa janela são informações leves fornecidas pelo Microsoft Graph. Nenhuma mensagem é aberta durante essa etapa de configuração.

### Fila de conversões PST

<p align="center">
<img width="1920" height="1039" alt="conversoes pst" src="https://github.com/user-attachments/assets/83ef2fa4-4c74-44da-93ba-f6fc946114df" />
</p>

A página **Conversões PST** permite criar novas conversões, iniciar somente os itens selecionados, pausar, retomar, remover e abrir o destino. A tabela mostra origem, destino, nome exibido no Outlook, política para arquivos existentes, progresso, verificação, falhas e itens restantes.

### Configurações e credenciais

<p align="center">
<img width="1920" height="1039" alt="config imagem" src="https://github.com/user-attachments/assets/7e08dbae-7e4e-4443-ba47-03e7fee2e585" />
</p>

A página **Configurações** reúne preferências gerais, credenciais do Microsoft Entra ID, parâmetros operacionais e aparência. O segredo do cliente permanece oculto e pode ser mantido sem redigitação. O botão **Salvar e testar credenciais** valida a autenticação antes de iniciar operações.

> Antes de publicar capturas próprias, oculte Tenant ID, Client ID, nomes de mailboxes, caminhos internos e qualquer outro identificador do ambiente.

## Uso — Linha de comando

A CLI (`src/main.py`) é dividida em fases:

| Fase | O que faz | Parâmetros obrigatórios |
|---|---|---|
| `0` | Valida autenticação e acesso à mailbox (diagnóstico rápido, não exporta nada) | `--mailbox` |
| `1` | Inspeciona a mailbox (usuário, pastas, inbox, prévia de mensagens) | `--mailbox` |
| `2` | Exportação local inicial (sem estrutura de pastas completa) | `--mailbox` |
| `3` | Exportação preservando a estrutura de pastas | `--mailbox` |
| `4` | Exportação completa com checkpoint (recomendada para uso real) | `--mailbox` |
| `5` | Backup em lote via CSV (chama a fase 4 para cada mailbox) | `--batch` |

Principais opções (fases 2 a 5):

| Opção | Descrição |
|---|---|
| `--limit N` | Máximo de mensagens por pasta (padrão `25`), ignorado se `--all` for usado |
| `--all` / `--all-messages` | Exporta todas as mensagens, sem limite por pasta |
| `--attachments` / `--export-attachments` | Exporta anexos separadamente, além do `.eml` |
| `--skip-calendar` / `--skip-contacts` / `--skip-tasks` | Pula essas categorias |
| `--resume-path <pasta>` | Retoma um backup existente usando o checkpoint salvo |
| `--profile-only` | Exporta somente e-mails principais, ignorando pastas de arquivo/sistema (Arquivo Morto, Recoverable Items, Sync Issues, Conversation History, etc.) |
| `--exclude-folder "Nome"` | Ignora uma pasta específica; pode repetir a opção |
| `--job-options-file arquivo.json` | JSON com `selected_folder_ids` / `selected_folder_paths` para restringir o escopo |
| `--skip-precheck` | (fase 5) Pula a pré-validação do CSV |

Exemplos:

```bash
# Fase 0 — valida credenciais e acesso a uma mailbox específica
python -m src.main --phase 0 --mailbox usuario@empresa.com

# Fase 4 — backup completo, com anexos, checkpoint habilitado
python -m src.main --phase 4 --mailbox usuario@empresa.com --all --attachments

# Retomando um backup interrompido
python -m src.main --phase 4 --mailbox usuario@empresa.com --all \
    --resume-path "output/backups/usuario@empresa.com"
```

## Backup em lote (CSV)

O arquivo CSV precisa ter obrigatoriamente uma coluna `email` (uma mailbox por linha):

```csv
email
usuario1@empresa.com
usuario2@empresa.com
```

```bash
python -m src.main --phase 5 --batch users.csv --all --attachments
```

A fase 5 gera um relatório de pré-validação e um relatório final (JSON e CSV) com o resultado por mailbox.

## Retomada e checkpoints

Cada operação de backup mantém um checkpoint transacional em SQLite (WAL) dentro da própria pasta do backup, registrando pasta, página e itens já exportados. Se o processo for interrompido (rede, energia, cancelamento manual), basta executar novamente a fase 4 apontando `--resume-path` para a mesma pasta de backup — os itens já confirmados no checkpoint não são baixados de novo.

## Exportação para PST (Outlook Classic)

Converte um backup `.eml` já existente em um arquivo `.pst`, usando automação COM (`Outlook.Application`) para criar um **PST em formato Unicode** (`namespace.AddStoreEx(caminho, 2)`), preservando pastas, remetente/destinatário, corpo, datas e anexos, além de propriedades MAPI adicionais usadas para verificação.

```bash
python -m src.services.pst_export_service \
    --backup-root "output/backups/usuario@empresa.com" \
    --pst-path "output/pst/usuario@empresa.com.pst" \
    --pst-display-name "Backup - Usuário" \
    --verification-level balanced
```

Principais opções:

| Opção | Padrão | Descrição |
|---|---|---|
| `--backup-root` | *(obrigatório)* | Pasta do backup ou `.zip` contendo os `.eml` |
| `--pst-path` | *(obrigatório)* | Caminho do `.pst` de saída |
| `--pst-display-name` | `M365 Mailbox Backup` | Nome exibido no Outlook |
| `--existing-action` | `resume` | `resume` \| `number` \| `replace` \| `cancel`, se o PST já existir |
| `--folder-mode` | `preserve` | `preserve` (estrutura original) ou `single` (pasta única) |
| `--skip-attachments` | — | Não importa anexos |
| `--verification-level` | `balanced` | `quick` \| `balanced` \| `complete` |
| `--performance-profile` | `balanced` | `conservative` \| `balanced` \| `performance` \| `custom` |
| `--detach-after` | — | Remove o PST do Outlook ao finalizar a conversão |

A ferramenta verifica o espaço livre em disco antes e durante a gravação (`M365_PST_CAPACITY_PREFLIGHT`) e interrompe a operação com segurança se o destino não tiver espaço suficiente — isso é independente do limite do Outlook descrito a seguir.

## Aumentando o limite de tamanho do .pst no Windows (registro)

### Por que isso é necessário

Como este projeto sempre cria o `.pst` em **formato Unicode** (não no antigo formato ANSI, que tinha limite de 2 GB), o teto real que você vai encontrar não é o histórico limite de 2 GB — é o **limite padrão do Outlook para arquivos Unicode**, que desde o Outlook 2010 SP1 é de **50 GB**. Se um backup gerar (ou puder crescer até) um `.pst` maior que isso, ou se você quiser mais margem antes do aviso de "arquivo quase cheio", é preciso aumentar dois valores no Registro do Windows, na máquina onde a conversão para PST é executada.

### Passo a passo manual

1. Feche o Outlook completamente (confira também na bandeja do sistema / Gerenciador de Tarefas).
2. Pressione `Win + R`, digite `regedit` e pressione Enter (aceite o UAC).
3. **Recomendado:** faça um backup da chave atual — clique com o botão direito em `Outlook` (ou em `PST`, se já existir) → **Exportar**.
4. Navegue até:

   ```
   Computador\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\PST
   ```

   `16.0` corresponde ao Microsoft 365 / Outlook 2016, 2019, 2021 e 2024. Use `15.0` para Outlook 2013 ou `14.0` para Outlook 2010.
5. Se a chave `PST` não existir: clique com o botão direito em `Outlook` → **Novo → Chave** → nomeie como `PST`.
6. Dentro da chave `PST`, crie (se ainda não existirem) dois valores **DWORD (32 bits)**: clique com o botão direito no painel à direita → **Novo → Valor DWORD (32 bits)** → nomeie `MaxLargeFileSize`; repita para `WarnLargeFileSize`.
7. Clique duas vezes em cada valor, selecione a base **Decimal** e informe o tamanho **em megabytes (MB)**. Usando os mesmos valores da captura de tela deste projeto:

   | Valor | Base | Decimal | Hexadecimal | Tamanho efetivo |
   |---|---|---|---|---|
   | `MaxLargeFileSize` | Decimal | `102400` | `0x00019000` | 100 GB |
   | `WarnLargeFileSize` | Decimal | `92160` | `0x00016000` | 90 GB |

8. Feche o Editor do Registro e **reinicie o Outlook**.

### Alternativa via arquivo `.reg`

Salve o conteúdo abaixo como `aumentar_limite_pst.reg`, dê duplo clique nele e confirme a importação (ajuste `16.0` conforme a versão do Outlook, se necessário):

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\PST]
"MaxLargeFileSize"=dword:00019000
"WarnLargeFileSize"=dword:00016000
```

### Observações importantes

- Esses valores afetam tanto `.pst` quanto `.ost` (inclusive o cache do Exchange, se usado no perfil).
- Mantenha `WarnLargeFileSize` pelo menos ~5% abaixo de `MaxLargeFileSize` (no exemplo acima a diferença é de 10%), para que o aviso apareça antes do limite rígido ser atingido, sem travar operações do Outlook.
- Valor máximo absoluto documentado: `MaxLargeFileSize` até `4294967295` (`0xFFFFFFFF`) e `WarnLargeFileSize` até `4090445042` (`0xF3CF3CF2`); acima disso a configuração é ignorada.
- Se o ambiente é gerenciado por política de grupo, verifique também `HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\PST` — quando presente, essa chave tem prioridade sobre a de `Software\Microsoft\...` e só pode ser alterada por um administrador.
- Arquivos `.pst` muito grandes deixam o Outlook mais lento para abrir, indexar e compactar; quando possível, prefira dividir o backup por ano ou por mailbox em vez de um único PST enorme.
- Este ajuste é do Outlook/Windows e é **independente** da verificação de espaço em disco da própria ferramenta (`M365_PST_CAPACITY_PREFLIGHT`): mesmo com o registro ajustado, a exportação também para se o disco de destino não tiver espaço livre suficiente.
- Precisa ser feito em **cada máquina/perfil de usuário do Windows** onde a conversão para PST for executada.

## Limitação de taxa e throttling adaptativo

O projeto combina dois mecanismos:

- **Rate limiter preventivo** (`pyrate-limiter`, opcional em SQLite compartilhado entre processos) — limites configuráveis por segundo/minuto, globais e por mailbox (`M365_GLOBAL_MIME_RATE_SECOND`, `M365_MAILBOX_MIME_RATE_SECOND`, etc.), para evitar bater no throttling do Graph.
- **Throttling adaptativo reativo** — ao receber `429`, respeita o cabeçalho `Retry-After`, faz backoff com jitter e reduz temporariamente o paralelismo (`M365_THROTTLE_SAFETY_SECONDS`, `M365_THROTTLE_JITTER_MAX_SECONDS`, `M365_THROTTLE_RECOVERY_SECONDS`), voltando ao normal gradualmente.

A aba de métricas da GUI mostra a saúde da API (Normal / Atenção / Limitada) com recomendação de ajuste de paralelismo, com base em taxa de throttle e latência média recentes.

## Logs, métricas e diagnóstico

- `logs/m365_mailbox_backup.log` — log geral, rotativo (10 MB × 5 arquivos).
- `logs/pst/<id-da-operação>.log` — log dedicado de cada conversão PST.
- `logs/application_crash.log` — capturado pelo `CrashReporter` em caso de exceção não tratada.
- Banco SQLite de métricas de chamadas ao Graph (latência, throttles, sucesso/falha), exportável em CSV pela GUI.
- Diálogo **"Sobre e diagnóstico"** na GUI executa verificações de credenciais, destinos de escrita, espaço em disco, SQLite e disponibilidade do Outlook Classic.

## Build / empacotamento (PyInstaller)

O `requirements.txt` já inclui `pyinstaller` e `pyinstaller-hooks-contrib`. Em produção, o projeto é distribuído como quatro executáveis independentes (a GUI e três processos que ela orquestra):

| Ponto de entrada (fonte) | Executável esperado |
|---|---|
| `m365_backup_gui.py` | GUI principal |
| `m365_backup_coordinator.py` | `m365_backup_coordinator.exe` |
| `src/main.py` (`python -m src.main`) | `run_backend.exe` |
| `src/services/pst_export_service.py` | `run_pst.exe` |

Exemplo ilustrativo de build individual (ajuste `--hidden-import`, ícones e `--add-data` conforme a sua configuração de `.spec`, se houver):

```bash
pyinstaller --onefile --name m365_backup_gui m365_backup_gui.py
pyinstaller --onefile --name m365_backup_coordinator m365_backup_coordinator.py
pyinstaller --onefile --name run_backend src/main.py
pyinstaller --onefile --name run_pst src/services/pst_export_service.py
```

## Solução de problemas

| Sintoma | Causa provável | O que verificar |
|---|---|---|
| Falha de autenticação na Fase 0 | `TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET` incorretos ou consentimento de admin pendente | `.env`, permissões de aplicativo no Entra ID |
| Muitos `429` / operação lenta | Paralelismo alto para o perfil do tenant | Reduza `M365_MIME_MAX_CONCURRENCY` / concorrência na GUI |
| "Outlook Classic indisponível" no diagnóstico | Outlook não instalado ou é a versão *nova* do Outlook | Instale/abra o Outlook Classic com um perfil configurado |
| PST não cresce além de determinado tamanho | Limite padrão do Outlook (50 GB) para arquivos Unicode | Veja [Aumentando o limite de tamanho do .pst](#aumentando-o-limite-de-tamanho-do-pst-no-windows-registro) |
| GUI não conecta ao coordenador | Porta `8765` (coordenador) ou `48765` (instância única) ocupada por outro processo | Feche instâncias antigas do app/`m365_backup_coordinator.exe` |
| Operação pausa sozinha | Pouco espaço em disco no destino | `M365_DISK_WARNING_GB` / `M365_DISK_CRITICAL_GB`, libere espaço |

## Segurança

- Nunca faça commit do `.env` (já listado no `.gitignore`) nem do segredo do cliente do App Registration.
- Prefira salvar as credenciais pela GUI, que as protege com **DPAPI** (`_gui_state/credentials.bin`), vinculado ao usuário/máquina do Windows.
- Aplique o princípio do menor privilégio ao App Registration: conceda apenas as permissões necessárias e considere uma *Application Access Policy* do Exchange Online para restringir quais mailboxes o app pode acessar.
- `logs/`, `_gui_state/`, `output/backups/*` e `output/pst/*` já estão no `.gitignore` — eles podem conter dados pessoais/sensíveis das mailboxes.

## Licença

Este projeto está publicamente disponível para uso pessoal, educacional, organizacional e não comercial.

Você está livre para:

Usar o software para fins pessoais, educacionais ou organizacionais;
Usar o software dentro de organizações comerciais para fins comerciais internos;
Estudar e modificar o código-fonte;
Compartilhar o código-fonte original ou modificado gratuitamente;
Criar trabalhos derivados para distribuição não comercial.
Comercialização

Você não pode vender, licenciar, alugar, arrendar, sublicenciar ou distribuir comercialmente este software ou versões derivadas dele sem a autorização expressa e por escrito do autor.

Isso inclui, mas não se limita a:

Vender o software ou versões modificadas dele;
Cobrar dos usuários pelo acesso ao software;
Oferecer o software como um SaaS pago ou serviço hospedado;
Redistribuir o software como um produto comercial;
Licenciar ou sublicenciar o software a terceiros para fins comerciais;
Monetizar o software em si ou uma versão derivada dele.

Organizações e empresas estão expressamente autorizadas a utilizar este software internamente para suas próprias operações comerciais, desde que não distribuam comercialmente, vendam, licenciem ou monetizem o software em si.

Para obter uma licença comercial ou autorização para distribuir ou monetizar comercialmente o software, entre em contato com o autor.

**Copyright © 2026 Erick Paiva Silva. Todos os direitos reservados, exceto pelas permissões expressamente concedidas acima.**
