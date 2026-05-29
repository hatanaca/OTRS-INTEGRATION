#Requires -Version 5.1
<#
.SYNOPSIS
    Exemplo: preencher o Gerador CCO no browser com Selenium (WebDriver).

.DESCRIPTION
    NAO e chamado pelo Menu-OTRS por defeito (use HubWebDriverEnabled=true
    no config.json para o fluxo integrado na opcao 7 - ver docs/automacao-formulario-hub.md).
    Selenium sem Gallery: copie o modulo para tools\Selenium\ ou use HubSeleniumModulePath.
    Use este ficheiro isoladamente quando quiser automacao real no formulario.

    Pre-requisitos:
      Install-Module Selenium -Scope CurrentUser
      Chrome instalado (Start-SeChrome). Para Edge use Start-SeEdge.

    Este exemplo usa a API classica do modulo (Start-SeChrome, Enter-SeUrl,
    Find-SeElement, Send-SeKeys). Se tiver apenas Selenium 4.x no Gallery,
    veja docs/automacao-formulario-hub.md e o README do modulo.

.PARAMETER HubBaseUrl
    Ex.: http://172.16.0.49:3210

.PARAMETER PayloadPath
    JSON: number, status, openingDate, openingHour, client, occurrence, updates...
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$HubBaseUrl,

    [Parameter(Mandatory = $true)]
    [string]$PayloadPath
)

if (-not (Get-Module -ListAvailable -Name Selenium)) {
    Write-Error "Instale: Install-Module Selenium -Scope CurrentUser"
    exit 1
}

Import-Module Selenium -ErrorAction Stop

if (-not (Test-Path -LiteralPath $PayloadPath)) {
    Write-Error "Ficheiro nao encontrado: $PayloadPath"
    exit 1
}

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$startUrl = $HubBaseUrl.TrimEnd('/') + '/api/relatorio'

if (-not (Get-Command Start-SeChrome -ErrorAction SilentlyContinue)) {
    Write-Error "Este exemplo espera Start-SeChrome (modulo Selenium classico). Veja docs/automacao-formulario-hub.md para v4."
    exit 1
}

$Driver = Start-SeChrome
Enter-SeUrl -Url $startUrl -Driver $Driver

function Wait-FindElement {
    param($By, [string]$Value, [int]$TimeoutSec = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            switch ($By) {
                'Name' { $e = Find-SeElement -Driver $Driver -Wait -Timeout 2 -Name $Value -ErrorAction SilentlyContinue }
                'Id'   { $e = Find-SeElement -Driver $Driver -Wait -Timeout 2 -Id $Value -ErrorAction SilentlyContinue }
                'Css'  { $e = Find-SeElement -Driver $Driver -Wait -Timeout 2 -Css $Value -ErrorAction SilentlyContinue }
            }
            if ($e) { return $e }
        } catch { }
        Start-Sleep -Milliseconds 400
    }
    return $null
}

function Set-Field {
    param($Element, [string]$Text)
    if ($null -eq $Element) { return }
    Send-SeKeys -Element $Element -Keys ([OpenQA.Selenium.Keys]::Control + 'a')
    Send-SeKeys -Element $Element -Keys $Text
}

try {
    $n = Wait-FindElement 'Name' 'number'
    Set-Field $n ([string]$payload.number)

    $s = Wait-FindElement 'Name' 'status'
    if ($s) { Set-Field $s ([string]$payload.status) }

    $od = Wait-FindElement 'Name' 'openingDate'
    if ($od) { Set-Field $od ([string]$payload.openingDate) }

    $oh = Wait-FindElement 'Name' 'openingHour'
    if ($oh) { Set-Field $oh ([string]$payload.openingHour) }

    $cl = Wait-FindElement 'Name' 'client'
    if ($cl) { Set-Field $cl ([string]$payload.client) }

    if ($payload.occurrence) {
        $oc = Wait-FindElement 'Name' 'occurrence'
        if ($oc) { Set-Field $oc ([string]$payload.occurrence) }
    }

    Write-Host "Campos principais preenchidos. Revise e grave no Hub." -ForegroundColor Green
}
finally {
    # Stop-SeDriver -Target $Driver
}
