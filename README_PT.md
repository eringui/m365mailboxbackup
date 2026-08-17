# M365 Mailbox Backup

Ferramenta para **Windows** que faz backup de caixas de e-mail do **Microsoft 365** usando a **Microsoft Graph API**, salva tudo em arquivos `.eml` (com retomada automática caso o processo seja interrompido) e, opcionalmente, converte esse backup em um arquivo `.pst` do Outlook.

Tudo é controlado por uma interface gráfica (GUI) feita em **PySide6**, com fila de operações, progresso em tempo real, pausa/retomada e diagnóstico do ambiente.

> Versão do aplicativo: `3.0.0`

---

## ⚠️ Aviso importante

Este projeto foi desenvolvido com **forte apoio do ChatGPT 5.6** — a maior parte do código, da lógica de retomada, da interface gráfica e da conversão para PST foi escrita em conjunto com a IA, em um processo iterativo ("vibe coding").

Por causa disso:

- O projeto **não passou por auditoria de segurança profissional** nem por uma bateria extensa de testes automatizados.
- Pode conter **bugs, comportamentos inesperados, telas incompletas ou textos inconsistentes** entre partes diferentes do código.
- **Teste sempre em um ambiente controlado** (uma mailbox de teste, por exemplo) antes de confiar o backup de contas importantes a ele.
- Use por sua conta e risco. Veja a seção [Erros conhecidos e limitações](#erros-conhecidos-e-limitações) antes de usar em produção.

---

## O que o programa faz

O projeto tem duas etapas independentes:

1. **Backup (.eml)** — autentica no Microsoft Graph como aplicativo (sem precisar logar usuário por usuário), percorre pastas, mensagens, calendário, contatos e tarefas de uma ou mais mailboxes, e salva tudo em arquivos `.eml` no disco. Se o processo cair no meio do caminho (queda de energia, rede, cancelamento manual), ele pode ser retomado de onde parou.
2. **Conversão para PST** *(opcional)* — pega um backup `.eml` já existente e, usando automação COM do **Outlook Classic** instalado na máquina, monta um arquivo `.pst` navegável, preservando a estrutura de pastas, remetente/destinatário, datas e anexos.

A interface gráfica **não faz esse trabalho pesado sozinha**: ao abrir, ela liga um serviço local (o "coordenador", feito em FastAPI, rodando em `127.0.0.1:8765`) que controla a fila de operações e dispara processos separados para cada backup ou conversão. Isso permite fechar e reabrir a GUI sem perder o progresso de operações em andamento.

---

## Como a interface gráfica funciona

Ao abrir o programa (`python m365_backup_gui.py` ou o executável), você vê um menu lateral com 4 páginas:

### 📥 Backups

A página principal, onde fica a fila de mailboxes a serem baixadas.

- **Adicionar** — adiciona uma mailbox pelo e-mail.
- **Importar CSV** — importa várias mailboxes de uma vez (arquivo com uma coluna `email`, veja `users-example.csv`).
- **Continuar existente** — aponta para uma pasta de backup já iniciada e retoma o trabalho a partir do checkpoint salvo, sem recontar tudo do zero.
- **Configurar pastas** — abre um diálogo para escolher quais pastas do e-mail serão baixadas, com opções de exportar todas as mensagens, exportar anexos separadamente, ignorar calendário/contatos/tarefas, limitar mensagens por pasta e aplicar essa configuração à fila inteira.
- **Iniciar backups**, **Pausar**, **Retomar**, **Remover**.
- **Corrigir falhas** — em vez de reprocessar a mailbox inteira, tenta novamente **só os e-mails que falharam** durante o backup, usando o checkpoint salvo.
- A tabela mostra status, etapa atual, quantidade de e-mails na caixa, progresso, itens faltando, tamanho já baixado, tempo estimado, pastas selecionadas e posição na fila. A ordem da fila pode ser arrastada com o mouse.
- Um painel inferior mostra o log em tempo real da mailbox selecionada na tabela.

### 📦 Conversões PST

Onde você transforma um backup `.eml` já concluído em um arquivo `.pst` do Outlook.

- **Nova conversão** abre um diálogo com três abas:
  - **Geral** — pasta de origem (o backup), pasta e nome do arquivo `.pst` de destino, nome exibido dentro do Outlook, e o que fazer se o PST já existir (retomar pelo checkpoint, criar um nome numerado, substituir ou cancelar).
  - **Conteúdo** — preservar a árvore de pastas original ou reunir tudo em uma pasta única, exibir metadados originais no corpo da mensagem, importar anexos/imagens embutidas e o tamanho máximo das imagens.
  - **Desempenho e segurança** — velocidade de importação, nível de verificação (rápida, balanceada ou completa), perfil de desempenho, ajuste automático de workers, orçamento de memória e o que é considerado um e-mail "grande".
- A tabela mostra origem, destino, nome no Outlook, política de conflito, progresso, verificação, falhas e itens restantes.
- **Esta função exige o Outlook Classic instalado e configurado na máquina** — o backup em `.eml` funciona sem o Outlook.

### 🧾 Logs

Um visualizador de log em tempo real, direto na interface, sem precisar abrir a pasta `logs/` manualmente.

- Alterna entre o log operacional geral e o log de falhas da aplicação.
- Filtra por nível (DEBUG/INFO/WARNING/ERROR/CRITICAL) e por texto pesquisado.
- Permite pausar a rolagem automática (sem parar a gravação do arquivo), limpar a visualização e abrir a pasta de logs.

### ⚙️ Configurações

Reúne as preferências do aplicativo, divididas em abas:

- **Geral** — pastas de destino dos backups e dos PSTs, quantidade de backups e conversões PST simultâneos, e opções padrão de conteúdo (exportar tudo, ignorar calendário/contatos/tarefas).
- **Credenciais e Graph** — Tenant ID, Client ID e Client Secret do aplicativo cadastrado no Microsoft Entra ID. O botão **Salvar e testar credenciais** já valida a autenticação na hora.
- **Aparência e PST** — tema (automático/claro/escuro), tamanho da fonte e se o PST deve ser removido do Outlook automaticamente ao terminar.
- Também há botões para **exportar/importar um perfil de configurações** (JSON), **executar diagnóstico** do ambiente e **validar a integridade** de um backup já feito.

> Existe também uma aba de **"Desempenho e armazenamento"** (limites de download, limitador de taxa/throttling, tamanho de página, espaço em disco etc.) já implementada no código, mas ela está **oculta por padrão** na versão atual da interface — veja a seção de limitações abaixo.

### Atalhos de teclado

`Ctrl+N` adicionar mailbox · `Ctrl+I` importar CSV · `Ctrl+O` continuar backup existente · `Ctrl+Enter` iniciar fila · `Ctrl+P` / `Ctrl+R` pausar/retomar seleção · `Delete` remover seleção · `F5` atualizar página · `Ctrl+1..4` trocar de página · `Ctrl+,` ir para Configurações · `Ctrl+L` ir para Logs.

---

## Requisitos

- **Windows 10/11.** O projeto depende de `pywin32` e do Outlook via COM; não funciona em macOS/Linux.
- **Outlook Classic** (aplicativo de mesa tradicional) instalado e com um perfil configurado — necessário **só** para a etapa de conversão em `.pst`. O backup em `.eml` funciona sem o Outlook instalado. O *novo Outlook para Windows* não é compatível com a automação usada aqui.
- **Python 3.11+** para rodar a partir do código-fonte (ou use os executáveis gerados pelo PyInstaller, se disponíveis).
- Um **aplicativo (App Registration)** cadastrado no Microsoft Entra ID, com permissões de aplicativo (não delegadas) no Microsoft Graph e consentimento de administrador — por exemplo `User.Read.All`, `Mail.Read`, `Calendars.Read`, `Contacts.Read` e `Tasks.Read`.
- Acesso de rede a `login.microsoftonline.com` e `graph.microsoft.com`.

---

## Instalação rápida

```bash
git clone <url-do-repositório> m365-mailbox-backup
cd m365-mailbox-backup

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

As credenciais (`TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`) podem ser configuradas de duas formas:

- Criando um arquivo `.env` na raiz do projeto; **ou**
- Preenchendo direto na tela **Configurações → Credenciais e Graph** da GUI, que salva o segredo protegido com **DPAPI** do Windows (não precisa de `.env`).

---

## Primeiro uso — passo a passo

```bash
python m365_backup_gui.py
```

1. Ao abrir pela primeira vez, vá em **Configurações → Credenciais e Graph**, preencha Tenant ID, Client ID e Client Secret, e clique em **Salvar e testar credenciais**.
2. Volte para **Backups** e adicione uma mailbox (**Adicionar**, **Importar CSV** ou **Continuar existente**, se já houver um backup parcial).
3. Opcionalmente, use **Configurar pastas** para escolher o que será baixado dessa mailbox.
4. Clique em **Iniciar backups** e acompanhe o progresso na tabela e no log inferior.
5. Quando o backup terminar, vá em **Conversões PST → Nova conversão** para gerar um `.pst`, se precisar abrir o backup no Outlook.

---

## Retomada, checkpoint e correção de falhas

Cada backup mantém um checkpoint transacional (`checkpoint.json` + banco SQLite) dentro da própria pasta do backup, registrando pastas, páginas e itens já exportados.

- Se o processo for interrompido por qualquer motivo, use **Continuar existente** apontando para a mesma pasta — os itens já confirmados no checkpoint não são baixados de novo.
- Se só alguns e-mails específicos falharam (erro de rede pontual, por exemplo), use **Corrigir falhas** em vez de reiniciar tudo: o programa tenta novamente apenas os itens marcados como falha no checkpoint.

---

## Conversão para PST em resumo

A conversão usa automação COM do Outlook Classic (`Outlook.Application`) para criar um **PST em formato Unicode**, preservando pastas, remetente/destinatário, corpo, datas e anexos. O programa verifica o espaço livre em disco antes e durante a escrita, e possui três níveis de verificação (rápida, balanceada e completa) para conferir se tudo foi realmente gravado.

> O Outlook, por padrão, tem um limite de **50 GB** para arquivos `.pst` em formato Unicode. Se o backup crescer além disso, é necessário aumentar `MaxLargeFileSize`/`WarnLargeFileSize` no Registro do Windows — isso é uma limitação do próprio Outlook, não da ferramenta.

---

## Erros conhecidos e limitações

- **Aba "Desempenho e armazenamento"** (limitador de taxa, workers, tamanho de página, alertas de disco) já existe no código das Configurações, mas está **desativada/oculta** na tela atual — os valores padrão internos são usados até que essa aba seja reativada em uma próxima versão.
- Existe também uma tela de **métricas de chamadas ao Graph** (latência, throttling, saúde da API) implementada no back-end, mas ela **não está conectada ao menu de navegação** da versão atual da GUI, então não aparece para o usuário.
- Funciona **somente no Windows**; a conversão PST depende do Outlook Classic clássico (não funciona com o "novo Outlook").
- Como grande parte do código foi gerada com IA em um processo iterativo, pode haver **pequenas inconsistências** entre textos, nomes de variáveis ou comportamentos de telas diferentes — revise o log em caso de dúvida.
- Falha de autenticação normalmente indica `TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET` incorretos ou consentimento de administrador pendente no Entra ID.
- Muitos erros `429` (limite de requisições do Graph) ou lentidão indicam paralelismo alto demais para o tenant — reduza a concorrência nas Configurações.
- Se a GUI não conseguir se conectar ao coordenador local, a porta `8765` (coordenador) ou `48765` (instância única) pode estar ocupada por outra instância do programa ainda aberta em segundo plano.
- Se uma operação pausar sozinha, verifique o espaço livre em disco no destino do backup.
- Não há testes automatizados abrangentes; mudanças de comportamento entre versões podem não estar totalmente documentadas.

Se encontrar um problema que não está listado aqui, confira a aba **Logs** e o diagnóstico em **Configurações → Executar diagnóstico** antes de abrir uma issue.

---

## Segurança e privacidade

- Nunca faça commit do arquivo `.env` (já está no `.gitignore`) nem do Client Secret do App Registration.
- Prefira salvar as credenciais pela GUI, que as protege com **DPAPI**, vinculado ao usuário/máquina do Windows.
- Aplique o princípio do menor privilégio no App Registration: conceda só as permissões necessárias e considere restringir o acesso do aplicativo a mailboxes específicas via *Application Access Policy* do Exchange Online.
- As pastas `logs/`, `_gui_state/` e `output/` podem conter dados pessoais/sensíveis de mailboxes reais — elas já estão listadas no `.gitignore` e não devem ser compartilhadas.

---

## Sobre este projeto

Este é um projeto pessoal/interno, desenvolvido com o auxílio do **ChatGPT 5.6** como parceiro de programação. Ele não é um produto comercial oficial, não tem suporte garantido e pode passar por mudanças bruscas de comportamento entre versões. Sinta-se à vontade para revisar, testar e adaptar o código para a sua realidade antes de usar em ambientes com dados importantes.

---

## Licença

Este projeto está disponível publicamente para uso pessoal, educacional, organizacional e não comercial.

Você pode usar, estudar, modificar e compartilhar o código-fonte original ou modificado gratuitamente, inclusive dentro de organizações, para uso interno. **Não é permitido vender, licenciar, sublicenciar ou distribuir comercialmente** este software ou versões derivadas dele sem permissão explícita por escrito do autor.

**Copyright © 2026 Erick Paiva Silva. Todos os direitos reservados, exceto pelas permissões concedidas acima.**
