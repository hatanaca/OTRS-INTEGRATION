# Pré-validação do relatório CCO (Hub)

Este diretório contém um **exemplo reutilizável** para o Hub: antes de mostrar o formulário do Gerador de Relatório CCO, a página confirma que existe **sessão autenticada** (por exemplo com `GET /api/relatorio`, o mesmo endpoint já usado pelo script PowerShell `Menu-OTRS.ps1` para ler dados).

## Ficheiros

| Ficheiro | Função |
|----------|--------|
| `js/relatorioCcoAuth.js` | Exporta `ensureCcoAccess()`: faz o pedido, mantém `#appContainer` oculto até `200`, e mostra `#preAuthGate` em caso de falha. |
| `js/sessionReminder.js` | Stub vazio para não dar 404 em testes; no Hub, substitua pelo ficheiro real. |
| `fragments-pre-auth.html` | Fragmento HTML (painel `#preAuthGate` e atributos `data-*` no `<body>`) para colar na vossa rota. |

## Integração no Hub

1. Copie `js/relatorioCcoAuth.js` para a pasta estática do Hub (ex.: `public/js/`).
2. No HTML da rota do relatório CCO, inclua o bloco **`#preAuthGate`** (ver `fragments-pre-auth.html`) **antes** de `#appContainer` e mantenha `#appContainer` com a classe `hidden` até `ensureCcoAccess()` concluir com sucesso.
3. No **início** do vosso `relatorioCco.js` (já carregado como `type="module"`), adicione:

```js
import { ensureCcoAccess } from "/js/relatorioCcoAuth.js";

await ensureCcoAccess();
// … resto do código existente …
```

4. Ajuste o URL de verificação se o Hub expuser um endpoint mais leve do que listar todo o relatório, por exemplo no `<body>`:

```html
<body data-auth-check-url="/api/session" data-login-url="/guest" data-home-url="/home">
```

Se `data-auth-check-url` não for definido, usa-se `/api/relatorio`.

## Comportamento

- **200**: o painel de validação oculta-se e `#appContainer` é exibido; o utilizador pode preencher campos.
- **401 / 403**: mensagem de sessão expirada ou sem permissão, com ligações para login e início.
- **Outros erros ou rede**: mensagem de erro e botão **Tentar novamente**.

## Nota de desempenho

`GET /api/relatorio` pode ser pesado se devolver muitos tickets. Para produção, recomenda-se um endpoint dedicado (por exemplo `GET /api/session` ou `HEAD /api/relatorio`) que apenas valide a sessão.
