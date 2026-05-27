# Automatizar o preenchimento do Gerador CCO (`/api/relatorio`)

O Menu-OTRS **não consegue**, sozinho e só com HTTP a partir do PowerShell, simular cliques e teclas **dentro** do separador do Chrome/Edge já aberto pelo utilizador. Para “automação nativa” do formulário existem três famílias de solução.

## 1. Alterar o Hub (recomendado para o produto final)

No **mesmo origin** que serve `/api/relatorio`, implemente por exemplo:

- Leitura de um **JSON na query** ou hash controlado (`?import=base64…`) ao carregar a página, ou  
- `POST /api/relatorio/draft` que grava rascunho e o `relatorioCco.js` chama no `mounted`, ou  
- `localStorage` com chave acordada, preenchida por uma página estática **no mesmo host** que redireciona para `/api/relatorio`.

Assim o operador abre um link ou recarrega e o formulário vem preenchido **sem** WebDriver nem consola.

## 2. WebDriver (Selenium / Playwright) chamado a partir do PowerShell

O script (Menu-OTRS ou um `.ps1` à parte) **inicia um browser controlado por API**, navega para o Hub, faz login se necessário e preenche `input`/`textarea` com selectores reais do DOM.

- **Vantagem:** comportamento idêntico ao do utilizador; funciona com React se esperar pelos elementos e dispara `input`/`change` como no snippet da consola.  
- **Desvantagem:** instalar módulo ou drivers, manter versão do browser alinhada ao driver, e **ajustar selectores** ao HTML real do Hub.

Fluxo típico:

1. Instalar dependências (uma vez por máquina).  
2. O script exporta o payload (JSON) — pode reutilizar o mesmo objeto que o Menu-OTRS envia à API.  
3. WebDriver abre o URL do Gerador, espera o formulário, `SendKeys` / `ExecuteScript` por campo.

Ver exemplo em **`scripts/Exemplo-HubRelatorio-Selenium.ps1`** (API clássica `Start-SeChrome` / `Find-SeElement`). Para o pacote **Selenium 4.x** do Gallery (`Start-SeDriver`, `Get-SeElement`, `Invoke-SeKeys`), siga os exemplos em [adamdriscoll/selenium-powershell](https://github.com/adamdriscoll/selenium-powershell).

## 3. Extensão do browser ou aplicação de desktop (UI Automation)

Menos habitual para formulários web complexos: extensão que recebe JSON (mensagem nativa ou ficheiro) e preenche o DOM na página activa; ou UI Automation no HWND — **frágil** para SPAs.

---

## Integrar com o Menu-OTRS

Sem tornar o Selenium obrigatório para todos:

- **Opção A:** Manter o fluxo actual (API + página “Preencher Hub” com script na consola) e usar o exemplo em `scripts/` só nas máquinas que precisem de automação total.  
- **Opção B:** No `Menu-OTRS.ps1`, após `Build-HubTicket`, gravar o JSON num ficheiro temporário e chamar `& .\scripts\Exemplo-HubRelatorio-Selenium.ps1 -PayloadPath …` se existir um ficheiro “gancho” ou parâmetro `-UseWebDriver` (pode adicionar ao `config.json` uma chave `HubWebDriverScriptPath` apontando para o seu `.ps1` personalizado).

A equipa do Hub continua a ser o sítio certo para uma solução **sem** WebDriver, se o objectivo for zero instalação extra nos postos.
