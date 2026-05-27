# OTRS-INTEGRATION

Script PowerShell `Menu-OTRS.ps1` para exportar relatórios CCO a partir do Znuny/OTRS, visualizar chamados e sincronizar com o **Hub** (aplicação web em `http://172.16.0.49:3210` ou URL configurável).

## Requisitos

- Windows PowerShell 5.1 ou PowerShell 7+ no posto do operador.

### Acentuação e símbolos nas notas

As respostas HTML do Znuny/OTRS são decodificadas com o **charset** indicado no cabeçalho `Content-Type` ou na meta da página; se UTF-8 não bater com os bytes (caracteres substitutos), o script tenta **Windows-1252** e **ISO-8859-1**, comuns em conteúdo legado em português. O arquivo `Menu-OTRS.ps1` deve ser guardado em **UTF-8** (idealmente com BOM no Windows PowerShell 5.1, conforme o cabeçalho do próprio script).

## Configuração (`config.json`)

| Campo | Descrição |
|--------|------------|
| `BaseURL` | URL do Znuny/OTRS |
| `Username` / `Password` | Credenciais do agente |
| `SearchPath` | Perfil de busca (KPI) para listar chamados ativos |
| `EstadoFile` | Arquivo JSON de cache (estado dos chamados) |
| `OutputPath` | Pasta de saída dos relatórios |
| `HubBaseURL` | URL base do Hub (ex.: `http://172.16.0.49:3210`) para opção de sincronização |
| `HubEncaminharPath` | Rota web relativa após cada envio bem-sucedido (padrão `home`), usada ao abrir o navegador para encaminhar o relatório na interface |
| `HubEmail` | E-mail do login JSON do Hub (`POST /api/login`) — opcional; se preenchido, a opção 7 usa como padrão |
| `HubPassword` | Senha do Hub para a mesma API — **texto claro**; opcional com Enter na sincronização para reutilizar a gravada |
| `HubApiRelatorioPath` | Caminho relativo à URL base para listar/criar relatórios (padrão `api/relatorio`). Se o servidor responder `Cannot POST /api/relatorio`, ajuste para o caminho real (ex.: `api/relatorios`), conforme o Hub expõe na rede (DevTools → Network). |

Copie `config.example.json` para `config.json` e preencha. O arquivo `config.json` está no `.gitignore` para evitar enviar credenciais ao Git. No topo de `Menu-OTRS.ps1` existem `MenuOtrsHubDefaultEmail` e `MenuOtrsHubDefaultPassword` usados quando o JSON não traz essas chaves — pode editar aí em vez do `config.json`.

## Menu principal

1. **Gerar relatório TXT** — Formato WhatsApp/CCO; na mesma execução grava o `Relatorio_CCO_*.json` (resumo ativos/resolvidos). Antes da exportação, pergunta se deseja **todas as notas** ou **apenas as 5 mais recentes** por chamado nos TXT e no JSON de resumo (o cache interno continua com todas as notas).
2. **Gerar relatório JSON** — Mesmo fluxo do item 1: atualiza o cache a partir do OTRS, gera TXT e `Relatorio_CCO_*.json` na mesma execução, com a mesma opção de notas (**todas** vs **5 últimas**) no material exportado.
3. **Visualizar chamados** — Submenu:
   - **OTRS tempo real (4 notas)** — Atualização a cada 60 s; apenas as **quatro notas mais recentes** por chamado (consulta direta ao OTRS).
   - **OTRS tempo real (todas)** — Mesmo fluxo em tempo real, porém com **todas as notas** de cada chamado ativo (mais lento).
   - **Cache local** — Último JSON gerado, sem consultar o OTRS; rolagem livre das notas.
4. **Alterar credenciais** — OTRS.
5. **Configurações** — Inclui URL do Hub.
6. **Salvar credenciais** — Grava `config.json` (senha em texto claro).
7. **Sincronizar com Hub** — Login em `/api/login` (JSON); leitura e envio usam o caminho configurável `HubApiRelatorioPath` (padrão `api/relatorio`, ou seja, `GET`/`POST` em `/api/relatorio` e `PUT` em `/api/relatorio/{numero}`). Fluxo: resumo, ocorrência opcional, HTML de revisão, confirmação e `POST`/`PUT`. Após sucesso, opção de abrir o Hub (`HubEncaminharPath`). Se `HubEmail` / `HubPassword` existirem em `config.json`, o login na opção 7 pode ser só *Enter* na senha.

### Normalização (avisos ao operador)

Nos modos em tempo real e ao recarregar o cache (`R` no visualizador offline), se o **estado** de um chamado passar a um valor considerado resolvido/normalizado (resolvido, fechado, merged, etc.), é exibido um **alerta em tela cheia** listando os chamados afetados.

## Integração Hub / relatório CCO

O payload enviado contém: número do ticket, status, data/hora de abertura, cliente e lista de atualizações (texto/data). O campo **Ocorrência** pode ser informado no terminal durante a sincronização; se preenchido, é incluído no JSON como `ocorrencia` (o Hub precisa aceitar esse campo na API; caso contrário, deixe em branco).

Antes do envio, o script gera um arquivo HTML temporário com o mesmo conteúdo e abre o navegador para **validação visual** pelo operador; o envio só ocorre após confirmação no terminal.

### Erros comuns

- **`Cannot POST /...`** (HTTP 404): o caminho da API no servidor não corresponde ao padrão. No navegador, com o Hub autenticado, abra as ferramentas de rede e veja qual URL o front usa para criar relatório; copie só a parte após o host (ex.: `api/v1/relatorios`) para `HubApiRelatorioPath` no `config.json` ou no menu **Configurações** (5).
- **`JSON primitivo inválido` / lista vazia na leitura**: o `GET` de listagem devolveu corpo que não é JSON (por exemplo `.` ou HTML). Com `HubApiRelatorioPath` correto e sessão válida após login, a resposta deve ser um array ou objeto JSON.

A detecção de mudança para atualização compara: status, cliente, data/hora de abertura e **todas** as entradas de atualização (não só a contagem de notas). Se o payload incluir `ocorrencia`, essa propriedade também entra na comparação com o registo existente no Hub.

O login usa JSON seguro (`ConvertTo-Json`) para evitar problemas com caracteres especiais no e-mail ou na senha.

## Fluxo sugerido no Hub

1. Acessar `http://...:3210/guest` (ou rota de login) e autenticar.
2. Ajustar em **Configurações** o campo `HubEncaminharPath` se a rota de encaminhamento no Hub não for `home`.
3. No Windows, executar o menu e a opção **7** após gerar o cache (opções 1 ou 2): revisar o HTML, confirmar o envio e, se desejar, abrir o Hub para concluir o encaminhamento na UI.
