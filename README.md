# OTRS-INTEGRATION

Script PowerShell `Menu-OTRS.ps1` para exportar relatórios CCO a partir do Znuny/OTRS, visualizar chamados e sincronizar com o **Hub** (aplicação web em `http://172.16.0.49:3210` ou URL configurável).

## Requisitos

- Windows PowerShell 5.1 ou PowerShell 7+ no posto do operador.

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

## Menu principal

1. **Gerar relatório TXT** — Formato WhatsApp/CCO.
2. **Gerar relatório JSON** — Cache completo para outras ferramentas.
3. **Visualizar chamados** — Submenu:
   - **OTRS tempo real (4 notas)** — Atualização a cada 60 s; apenas as **quatro notas mais recentes** por chamado (consulta direta ao OTRS).
   - **OTRS tempo real (todas)** — Mesmo fluxo em tempo real, porém com **todas as notas** de cada chamado ativo (mais lento).
   - **Cache local** — Último JSON gerado, sem consultar o OTRS; rolagem livre das notas.
4. **Alterar credenciais** — OTRS.
5. **Configurações** — Inclui URL do Hub.
6. **Salvar credenciais** — Grava `config.json` (senha em texto claro).
7. **Sincronizar com Hub** — Login em `/api/login` (JSON); leitura de `/api/relatorio`; para cada inclusão ou alteração: resumo no terminal, **Ocorrência** opcional (texto multilinha), **pré-visualização em HTML** no navegador (espelho do payload da API), confirmação do operador e então `POST` ou `PUT` em `/api/relatorio`. Após sucesso, pergunta se deseja abrir o Hub (`HubEncaminharPath`, por padrão `/home`) para o passo de **encaminhar** na interface web.

### Normalização (avisos ao operador)

Nos modos em tempo real e ao recarregar o cache (`R` no visualizador offline), se o **estado** de um chamado passar a um valor considerado resolvido/normalizado (resolvido, fechado, merged, etc.), é exibido um **alerta em tela cheia** listando os chamados afetados.

## Integração Hub / relatório CCO

O payload enviado contém: número do ticket, status, data/hora de abertura, cliente e lista de atualizações (texto/data). O campo **Ocorrência** pode ser informado no terminal durante a sincronização; se preenchido, é incluído no JSON como `ocorrencia` (o Hub precisa aceitar esse campo na API; caso contrário, deixe em branco).

Antes do envio, o script gera um arquivo HTML temporário com o mesmo conteúdo e abre o navegador para **validação visual** pelo operador; o envio só ocorre após confirmação no terminal.

A detecção de mudança para atualização compara: status, cliente, data/hora de abertura e **todas** as entradas de atualização (não só a contagem de notas). Se o payload incluir `ocorrencia`, essa propriedade também entra na comparação com o registo existente no Hub.

O login usa JSON seguro (`ConvertTo-Json`) para evitar problemas com caracteres especiais no e-mail ou na senha.

## Fluxo sugerido no Hub

1. Acessar `http://...:3210/guest` (ou rota de login) e autenticar.
2. Ajustar em **Configurações** o campo `HubEncaminharPath` se a rota de encaminhamento no Hub não for `home`.
3. No Windows, executar o menu e a opção **7** após gerar o cache (opções 1 ou 2): revisar o HTML, confirmar o envio e, se desejar, abrir o Hub para concluir o encaminhamento na UI.
