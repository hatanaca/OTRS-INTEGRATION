# Automatizar o preenchimento do Gerador CCO (`/api/relatorio`)

## Opção mais compacta (sem permissões de instalação)

1. **Só o browser:** na opção **7**, use a página **«Preencher Hub»** e cole o script na **Consola (F12)** no Gerador — **zero** módulos PowerShell.
2. **Selenium sem `Install-Module`:** copie o módulo para `tools\Selenium\` ao lado do `Menu-OTRS.ps1` ou defina **`HubSeleniumModulePath`** no `config.json` (pasta vinda de `Save-Module` noutro PC). Instruções: **`tools/Selenium/README.md`**.
3. **Só ficheiros de utilizador:** noutro posto, `Save-Module -Name Selenium -Path "$env:USERPROFILE\Documents\WindowsPowerShell\Modules"` (muitas vezes **não** precisa de admin; depende da política).

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

No `config.json` defina **`HubWebDriverEnabled`** como `true` e **`HubWebDriverBrowser`** como `Chrome` ou `Edge`. O Selenium pode vir da **Gallery** (`Install-Module Selenium -Scope CurrentUser`), de uma pasta **`tools\Selenium\`** ao lado do script, ou de **`HubSeleniumModulePath`** — ver **`tools/Selenium/README.md`** se não puder instalar módulos.

Com **`HubWebDriverAutoFill`** a `true`, na sincronização (opção **7**) o script abre o WebDriver e preenche o Gerador **sem** perguntar `s/N` após os HTML; caso contrário o operador confirma no terminal.

Sem **`HubWebDriverEnabled`** no `config.json`, o Menu-OTRS não oferece Selenium na opção **7** (apenas HTML + API). Para experimentar WebDriver fora do menu, use **`scripts/Exemplo-HubRelatorio-Selenium.ps1`** com um payload JSON.
