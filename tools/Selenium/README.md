# Selenium portátil (sem `Install-Module` / sem administrador)

O `Menu-OTRS.ps1` procura o módulo **nesta ordem**:

1. Caminho em **`HubSeleniumModulePath`** no `config.json` (ficheiro `Selenium.psd1` ou pasta que o contenha).
2. Pasta **`tools\Selenium\`** ao lado do `Menu-OTRS.ps1` (deve existir `tools\Selenium\Selenium.psd1`).
3. Módulo **Selenium** instalado na máquina (`Install-Module` / Gallery).

## Copiar o módulo a partir de outro PC (recomendado em redes bloqueadas)

Num computador onde **possa** instalar:

```powershell
Save-Module -Name Selenium -Path 'C:\Temp\SeleniumZip' -Force
```

Copie a pasta resultante (contém `Selenium.psd1`, `Selenium.psm1`, `assemblies\`, etc.) para o posto restrito:

- **Opção A:** Coloque tudo em `tools\Selenium\` junto ao `Menu-OTRS.ps1` (crie `tools\Selenium` se não existir).
- **Opção B:** Guarde onde a política permitir (ex.: `D:\Libs\Selenium\`) e no `config.json` use `"HubSeleniumModulePath": "D:\\Libs\\Selenium"`.

Reabra o PowerShell e teste:

```powershell
Import-Module -LiteralPath '.\tools\Selenium\Selenium.psd1' -Force
Get-Command Start-SeChrome
```

## Opção ainda mais compacta (zero Selenium)

Use só a página **«Preencher Hub»** gerada na opção **7**: copia um script para colar na **Consola (F12)** no Gerador — não precisa de módulos nem de permissões extra.

## `chromedriver`

O pacote do módulo Selenium no Gallery inclui drivers; se faltar, siga o aviso na consola ou a documentação do [selenium-powershell](https://github.com/adamdriscoll/selenium-powershell).
