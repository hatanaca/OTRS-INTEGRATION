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
| `HubEncaminharPath` | Rota aberta no navegador após envio (padrão `api/relatorio` = página **Gerador de Relatório CCO**). Use `home` se preferir o painel inicial. |
| `HubEmail` | E-mail do login JSON do Hub (`POST /api/login`) — opcional; se preenchido, a opção 7 usa como padrão |
| `HubPassword` | Senha do Hub para a mesma API — **texto claro**; opcional com Enter na sincronização para reutilizar a gravada |
| `HubApiRelatorioPath` | Prefixo da API (padrão `api/relatorio`). A lista é lida com **GET** `…/tickets` (corpo JSON com `tickets[]`); se falhar, tenta o prefixo sozinho (hubs legados). |
| `HubPostTicketPaths` | (Opcional) Caminhos POST para criar ticket, separados por `;` ou `,`. Vazio = ordem automática: **`api/relatorio/tickets`** primeiro, depois `…/ticket`, prefixo raiz e outras rotas. |
| `HubPutTicketPaths` | (Opcional) Caminhos PUT; use `{numero}` para o número OTRS. Vazio = tenta `…/tickets/{numero}`, `…/{numero}`, `…/ticket/{numero}`, etc. |
| `HubFormSelectors` | (Opcional) Objeto JSON: por campo (`number`, `status`, `openingDate`, `openingHour`, `client`, `occurrence`) uma lista de selectores CSS a tentar **antes** dos padroes, ao usar a pagina «Preencher Hub» (script na consola). |
| `HubWebDriverEnabled` | Se `true` (padrão no script), na opção **7** o Menu-OTRS **escreve directamente** no Gerador CCO via WebDriver (JavaScript na página). Requer Selenium (Gallery, `tools\Selenium\` ou `HubSeleniumModulePath`). |
| `HubWebDriverAutoFill` | Se `true` (padrão no script), inicia o WebDriver e preenche **sem** pergunta `s/N`. Defina `false` no `config.json` para voltar a confirmar no terminal. |
| `HubBrowserDirectWrite` | Se `true` (padrão), o WebDriver injecta o mesmo JavaScript da consola no DOM (`HubBrowserDirectWrite=false` usa SendKeys). |
| `HubWebDriverDebugAddress` | (Opcional) Ligar ao Chrome/Edge **já aberto** (ex.: `127.0.0.1:9222`). Inicie o browser com `--remote-debugging-port=9222`; o Menu-OTRS não fecha essa janela. |
| `HubWebDriverBrowser` | `Chrome` ou `Edge` (padrão `Chrome`). Usado com `HubWebDriverEnabled`. |
| `HubSeleniumModulePath` | (Opcional) Caminho para `Selenium.psd1` ou para a pasta do módulo Selenium **copiada** (sem `Install-Module`). Vazio = usa `tools\Selenium\` junto ao `Menu-OTRS.ps1` ou o módulo na Gallery. |
| `FilterCustomerVisibleNotesOnly` | Se `true`, aplica o filtro de sessão OTRS **`CustomerVisibilityFilter=1`** (quando disponível no Znuny) e valida cada linha pela classe **`VisibleForCustomer`** / **`NotVisibleForCustomer`**. Padrão no exemplo: `false`. |
| `FilterNotesByReportDynamicField` | Se `true` (padrão no exemplo), exporta **somente** notas com o campo dinâmico de artigo **`DynamicField_Enviarpararelatorio=Sim`** (checkbox **Enviar para relatório** ao criar a nota em `AgentTicketNote`). Equivalente ao script shell/curl que envia `DynamicField_Enviarpararelatorio=Sim` no POST. O cache `estado_chamados.json` é apagado quando qualquer filtro de notas está ativo. |
| `ArticleDynamicFieldReportName` | Nome do campo dinâmico **sem** o prefixo `DynamicField_` (padrão `Enviarpararelatorio`). |
| `ArticleDynamicFieldReportValue` | Valor exigido para incluir a nota (padrão `Sim`). |

Copie `config.example.json` para `config.json` e preencha **Username**, **Password** (OTRS) e **HubPassword**. O arquivo `config.json` está no `.gitignore` (não vai para o Git). Na primeira execução, se `config.json` não existir, o `Menu-OTRS.ps1` cria uma cópia a partir de `config.example.json`. Também pode usar `scripts/Setup-Config.ps1`.

**Ambiente Microset (padrão no exemplo):**

| Campo | Valor |
|--------|--------|
| `BaseURL` | `http://172.16.0.12/znuny` |
| `SearchPath` | `index.pl?Action=AgentKPISearch;Subaction=Search;TakeLastSearch=1;Profile=94_8` |
| `HubBaseURL` | `http://172.16.0.49:3210` |
| `FilterCustomerVisibleNotesOnly` | `false` |
| `FilterNotesByReportDynamicField` | `true` |
| `ArticleDynamicFieldReportName` | `Enviarpararelatorio` |
| `ArticleDynamicFieldReportValue` | `Sim` |

Após preencher credenciais, use a opção **9** (Críticos relatório) para o mesmo critério do script interceptador, **8** (visíveis ao cliente) ou **1**/**2** conforme `config.json`. Apague `estado_chamados.json` se tiver exportado antes com filtros diferentes.

## Menu principal

1. **Gerar relatório TXT** — Formato WhatsApp/CCO; na mesma execução grava o `Relatorio_CCO_*.json` (resumo ativos/resolvidos). Por padrão (exemplo), **só entram notas com Enviar para relatório = Sim** (`DynamicField_Enviarpararelatorio`). Antes da exportação, pergunta se deseja **todas as notas** ou **apenas as 5 mais recentes** por chamado nos TXT e no JSON de resumo.
2. **Gerar relatório JSON** — Mesmo fluxo do item 1: atualiza o cache a partir do OTRS, gera TXT e `Relatorio_CCO_*.json` na mesma execução, com a mesma opção de notas (**todas** vs **5 últimas**) no material exportado.
3. **Visualizar chamados** — Submenu:
   - **OTRS tempo real (4 notas)** — Atualização a cada 60 s; **quatro notas mais recentes** por chamado (respeita o filtro de visibilidade ao cliente se estiver ativo).
   - **OTRS tempo real (todas)** — Mesmo fluxo com **todas as notas** por chamado (mais lento).
   - **Cache local** — Último JSON gerado, sem consultar o OTRS.
   - **Críticos visíveis (4)** — Perfil KPI + **somente** notas com “Ficar visível para o Cliente”; 4 notas recentes.
4. **Alterar credenciais** — OTRS.
5. **Configurações** — Inclui URL do Hub.
6. **Salvar credenciais** — Grava `config.json` (senha em texto claro).
7. **Sincronizar com Hub** — Login em `/api/login`; **GET** `…/api/relatorio/tickets`; **POST**/`PUT`. HTML de pré-visualização; com **`HubWebDriverEnabled`** o script **escreve directamente** no Gerador (WebDriver + JavaScript), numa janela nova ou no browser aberto (`HubWebDriverDebugAddress`). Após envio API, pode abrir o Hub (`HubEncaminharPath`).
8. **Críticos visíveis** — Perfil KPI + **apenas** notas com “Ficar visível para o Cliente” (`IsVisibleForCustomer`), independentemente do `config.json`.
9. **Críticos relatório** — Perfil KPI + **apenas** notas com `DynamicField_Enviarpararelatorio=Sim` (mesmo critério do POST `AgentTicketNote` do script shell/curl interceptador). Visibilidade ao cliente **desligada** nesta opção.

**Diagnóstico:** na exportação com `-DiagMode`, o script grava `diag_main_*.html`, `diag_filtrado_*.html` e `diag_artigos_*.txt` (colunas `visivel_cliente` e `enviar_relatorio` por `ArticleID`) na pasta de saída.

### Onde o OTRS marca notas para o relatório

| Tela | URL / `Action` | O que o Menu-OTRS usa |
|------|----------------|------------------------|
| **Adicionar nota** (popup) | `AgentTicketNote` — campo `DynamicField_Enviarpararelatorio=Sim` (**Enviar para relatório**) | **Sim** — ao enviar a nota (como no script curl: `Subaction=Store`, `DynamicField_Enviarpararelatorio=Sim`). |
| **Adicionar nota** (popup) | `AgentTicketNote` — checkbox `IsVisibleForCustomer` | Opcional — filtro separado (`FilterCustomerVisibleNotesOnly`). |
| **Detalhe do chamado** | `AgentTicketZoom` — widget do artigo (`ArticleMetaFields`) | **Sim** — `<label>Enviar para relatório</label><p class="Value">Sim</p>` dentro do bloco `Article{ArticleID}`. |

### Onde o OTRS grava “visível para o cliente”

| Tela | URL / `Action` | O que o Menu-OTRS usa |
|------|----------------|------------------------|
| **Adicionar nota** (popup) | `AgentTicketNote` — checkbox `IsVisibleForCustomer` (“Ficar visível para o Cliente”) | **Não** — só define a visibilidade **ao enviar** a nota nova. |
| **Detalhe do chamado** | `AgentTicketZoom` — tabela de artigos (`ArticleTable`) | **Sim** — cada linha `<tr id="Row1" …>` traz a classe **`VisibleForCustomer`** ou **`NotVisibleForCustomer`** conforme o valor já salvo no banco. |

`BaseURL` deve incluir o prefixo do Znuny, por exemplo `http://servidor/znuny` (como no seu HTML: `/znuny/index.pl`).

Para validar um chamado (ex. `TicketID=2840100`), abra o **chamado completo** no agente, use opção **9**, **8** ou export com `-DiagMode`, e confira `diag_artigos_*.txt`.

### Normalização (avisos ao operador)

Nos modos em tempo real e ao recarregar o cache (`R` no visualizador offline), se o **estado** de um chamado passar a um valor considerado resolvido/normalizado (resolvido, fechado, merged, etc.), é exibido um **alerta em tela cheia** listando os chamados afetados.

## Integração Hub / relatório CCO

O payload segue o front **`relatorioCco.js`**: campos com atributo **`data-field`** (`number`, `status`, `openingDate`, `openingHour`, `client`, `occurrence`), `updates` com **`updateDate`**, **`updateHour`**, **`text`**, e **PUT** em `/api/relatorio/tickets/{id}`. A UI do Hub limita a **4** linhas de atualização por ticket (`maxUpdates: 4`); o Menu-OTRS pode enviar mais notas na API, mas o formulário só mostra quatro.

Antes do envio, o script gera um arquivo HTML temporário com o mesmo conteúdo e abre o navegador para **validação visual** pelo operador; o envio só ocorre após confirmação no terminal.

### Preencher o formulário HTML do Gerador (`/api/relatorio`)

Um ficheiro HTML local **não pode** escrever noutra aba por política de origem do browser. O Menu-OTRS oferece três caminhos:

1. **Escrita directa (recomendado)** — `HubWebDriverEnabled`: `true`, Selenium instalado, `HubBrowserDirectWrite`: `true` (padrão). O WebDriver injecta JavaScript na página do Hub (mesma lógica do `relatorioCco.js`: `[data-field]`, `.update-entry`). Com **`HubWebDriverDebugAddress`** (ex.: `127.0.0.1:9222`) liga-se ao Chrome/Edge que **já tem aberto** (`chrome.exe --remote-debugging-port=9222`); sem esse campo abre uma janela WebDriver nova.

2. **Consola (sem Selenium)** — página «Preencher Hub»: copiar script, F12 → Consola no separador do Gerador, colar e Enter.

3. **Só API** — confirmação no terminal; recarregue o Gerador para ver os tickets.

Defina **`HubFormSelectors`** se o DOM do Hub for diferente. Com **`HubWebDriverAutoFill`**: `true` o preenchimento directo corre após a pré-visualização, sem pergunta `s/N`.

Exemplo de `HubFormSelectors` (inspecione o DOM do Hub com F12 → inspetor):

```json
"HubFormSelectors": {
  "number": ["[data-field=\"number\"]"],
  "occurrence": ["textarea[data-field=\"occurrence\"]"]
}
```

A página **`/api/relatorio`** do Hub é o gerador integrado; o **browser** usa `GET /api/relatorio/sync` para polling leve e `GET`/`POST` em **`/api/relatorio/tickets`** para lista e criação (conforme HAR do `relatorioCco.js`). O Menu-OTRS espelha o **POST**/**GET** de tickets nessa rota quando o prefixo é `api/relatorio`.

### Erros comuns

- **`Cannot POST /api/relatorio`**: essa rota costuma ser só **GET HTML** (página do gerador). O **POST** real do Hub analisado é **`POST /api/relatorio/tickets`**. O script usa essa rota em primeiro lugar; se a sua instância for diferente, defina `HubPostTicketPaths` com o URL visto no DevTools (Network). Se todas as tentativas falharem, o script grava JSON + HTML de ajuda e copia o payload.
- **`JSON primitivo inválido` / lista vazia na leitura**: o `GET` de listagem devolveu corpo que não é JSON (por exemplo `.` ou HTML). Com `HubApiRelatorioPath` correto e sessão válida após login, a resposta deve ser um array ou objeto JSON.

A detecção de mudança para atualização compara: status, cliente, data/hora de abertura e **todas** as entradas de atualização (`text`, `updateDate`/`updateHour` no envio; no Hub lido aceita também `date`/`hour` legados). Se o payload incluir `ocorrencia`, essa propriedade também entra na comparação com o registo existente no Hub.

O login usa JSON seguro (`ConvertTo-Json`) para evitar problemas com caracteres especiais no e-mail ou na senha.

## Fluxo sugerido no Hub

1. Autenticar no Hub (`/guest` ou rota de login da sua instalação).
2. Na mesma máquina, com o **Gerador** (`/api/relatorio`) aberto, use a segunda página HTML gerada pela opção **7** para **copiar o script** e colá-lo na **Consola** (F12), preenchendo o formulário; ou confie só na **API** e recarregue o Gerador para ver os tickets.
3. Use **Gerar Relatório**, **Copiar**, **WhatsApp** ou **E-mail** na própria página quando quiser concluir o envio ao CCO/cliente.
