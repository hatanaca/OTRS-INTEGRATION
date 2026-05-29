# Cria config.json a partir de config.example.json (nao sobrescreve se ja existir).
param(
    [string]$ConfigPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config.json'),
    [string]$ExamplePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config.example.json'),
    [switch]$Force
)

$root = Split-Path $PSScriptRoot -Parent
if (-not (Test-Path $ExamplePath)) {
    $ExamplePath = Join-Path $root 'config.example.json'
}
if (-not (Test-Path $ConfigPath)) {
    $ConfigPath = Join-Path $root 'config.json'
}

if ((Test-Path $ConfigPath) -and -not $Force) {
    Write-Host "Ja existe: $ConfigPath (use -Force para sobrescrever)." -ForegroundColor Yellow
    exit 0
}

if (-not (Test-Path $ExamplePath)) {
    Write-Error "Modelo nao encontrado: $ExamplePath"
    exit 1
}

Copy-Item -LiteralPath $ExamplePath -Destination $ConfigPath -Force
Write-Host "Criado: $ConfigPath" -ForegroundColor Green
Write-Host "Edite Username, Password e HubPassword antes de executar Menu-OTRS.ps1." -ForegroundColor Cyan
