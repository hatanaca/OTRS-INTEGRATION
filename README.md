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
| `HubApiRelatorioPath` | Caminho relativo à URL base para **GET** da lista (padrão `api/relatorio`). O **POST** de novo ticket pode estar noutra rota — o script tenta varias automaticamente ou use `HubPostTicketPaths`. |
| `HubPostTicketPaths` | (Opcional) Caminhos POST para criar ticket, separados por `;` ou `,`. Vazio = tentativas automaticas (`api/relatorio/ticket`, `api/tickets`, etc.). |
| `HubPutTicketPaths` | (Opcional) Caminhos PUT para atualizar; use `{numero}` no caminho. Vazio = tentativas automaticas. |

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
7. **Sincronizar com Hub** — Login em `/api/login`; `GET`/`POST` em `/api/relatorio` (configurável via `HubApiRelatorioPath`). O script envia os dados dos tickets; a **página** `…/api/relatorio` no navegador é o *Gerador CCO* (campos, gravação incremental). **Gerar relatório**, **WhatsApp** e **E-mail** são ações na própria interface, não feitas pelo Menu-OTRS. Após o `POST`/`PUT`, pode abrir o Hub (`HubEncaminharPath`, por defeito `api/relatorio`).

### Normalização (avisos ao operador)

Nos modos em tempo real e ao recarregar o cache (`R` no visualizador offline), se o **estado** de um chamado passar a um valor considerado resolvido/normalizado (resolvido, fechado, merged, etc.), é exibido um **alerta em tela cheia** listando os chamados afetados.

## Integração Hub / relatório CCO

O payload inclui `number`, `status`, `openingDate`, `openingHour`, `client`, `updates` e, se informado no terminal, **`ocorrencia`** e **`occurrence`** (este último alinha ao campo `data-field="occurrence"` do formulário HTML do Hub).

Antes do envio, o script gera um arquivo HTML temporário com o mesmo conteúdo e abre o navegador para **validação visual** pelo operador; o envio só ocorre após confirmação no terminal.

A página **`/api/relatorio`** do Hub é o gerador integrado: os tickets são **guardados** à medida que edita; não há um passo separado de “encaminhar” só pela API — o relatório final e envios (WhatsApp/e-mail) usam os botões da interface.

### Erros comuns

- **`Cannot POST /api/relatorio`**: em muitos Hubs essa URL só serve a **página** (GET HTML); o **POST** do ticket está noutro caminho. O script já tenta várias rotas em sequência. Se ainda falhar, veja no DevTools (Network) o URL exato ao gravar um ticket e preencha `HubPostTicketPaths` no `config.json` ou no menu **5**. Se **todas** as tentativas falharem, o script grava `Hub_ticket_<n>_*.json`, copia o JSON para a área de transferência, abre o Gerador CCO e uma página HTML de ajuda na pasta de saída.
- **`JSON primitivo inválido` / lista vazia na leitura**: o `GET` de listagem devolveu corpo que não é JSON (por exemplo `.` ou HTML). Com `HubApiRelatorioPath` correto e sessão válida após login, a resposta deve ser um array ou objeto JSON.

A detecção de mudança para atualização compara: status, cliente, data/hora de abertura e **todas** as entradas de atualização (não só a contagem de notas). Se o payload incluir `ocorrencia`, essa propriedade também entra na comparação com o registo existente no Hub.

O login usa JSON seguro (`ConvertTo-Json`) para evitar problemas com caracteres especiais no e-mail ou na senha.

## Fluxo sugerido no Hub

1. Autenticar no Hub (`/guest` ou rota de login da sua instalação).
2. A página **`/api/relatorio`** (Gerador de Relatório CCO) sincroniza com a API: o Menu-OTRS (opção 7) envia/atualiza tickets; no browser os campos são preenchidos e **gravados** conforme o fluxo do `relatorioCco.js`.
3. Use **Gerar Relatório**, **Copiar**, **WhatsApp** ou **E-mail** na própria página quando quiser concluir o envio ao CCO/cliente.
