# =============================================================================
# Menu-OTRS.ps1 - Interface TUI para relatorios CCO (Znuny/OTRS)
# Versao unificada - Export-CcoReport incorporado
# Compativel com Windows PowerShell 5.1 e PowerShell 7+
# =============================================================================

[CmdletBinding()]
param(
    [string]$ConfigFile = "config.json",
    [string]$ScriptPath = $null
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$script:ExportLoaded = $false

# =============================================================================
# Helpers de UI
# =============================================================================

function Clear-Screen { [Console]::Clear() }

function Write-Centered {
    param([string]$Text, [ConsoleColor]$Color = 'White')
    $w   = [Math]::Min([Console]::WindowWidth, 80)
    $pad = [math]::Max(0, ($w - $Text.Length) / 2)
    Write-Host ((' ' * [int]$pad) + $Text) -ForegroundColor $Color
}

function Write-Divider {
    param([ConsoleColor]$Color = 'DarkGray', [int]$W = 72)
    Write-Host ('=' * $W) -ForegroundColor $Color
}

function Write-ThinDiv {
    param([ConsoleColor]$Color = 'DarkGray', [int]$W = 72)
    Write-Host ('-' * $W) -ForegroundColor $Color
}

function Write-Banner {
    Clear-Screen
    Write-Host ""
    Write-Divider 'DarkCyan'
    Write-Centered "  ___  _____ ____  ____     __  _   _  ___   ____  " 'Cyan'
    Write-Centered " / _ \|_   _|  _ \/ ___|   / / | \ | |/ _ \ / ___| " 'Cyan'
    Write-Centered "| | | | | | | |_) \___ \  / /  |  \| | | | | |     " 'Cyan'
    Write-Centered "| |_| | | | |  _ < ___) |/ /   | |\  | |_| | |___  " 'Cyan'
    Write-Centered " \___/  |_| |_| \_\____//_/    |_| \_|\___/ \____| " 'Cyan'
    Write-Host ""
    Write-Centered "   Exportador de Relatorios CCO - Znuny/OTRS NOC   " 'White'
    Write-Divider 'DarkCyan'
    Write-Host ""
}

function Write-StatusBar {
    param([string]$User, [string]$HostDisplay, [string]$CfgStatus)
    Write-ThinDiv 'DarkGray'
    Write-Host ("  Usuario: " + $User + "   Servidor: " + $HostDisplay + "   Config: " + $CfgStatus) -ForegroundColor DarkGray
    Write-ThinDiv 'DarkGray'
    Write-Host ""
}

function Write-MenuOpt {
    param([string]$Key, [string]$Label, [ConsoleColor]$Color = 'White', [string]$Detail = '')
    Write-Host ("  [" + $Key + "]") -ForegroundColor Cyan -NoNewline
    Write-Host ("  " + $Label) -ForegroundColor $Color -NoNewline
    if ($Detail) { Write-Host ("  - " + $Detail) -ForegroundColor DarkGray }
    else         { Write-Host "" }
}

function Read-PasswordSecure {
    Write-Host "  Senha: " -ForegroundColor Yellow -NoNewline
    $ss  = Read-Host -AsSecureString
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
    try   { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
}

function Read-Field {
    param([string]$Prompt, [string]$Default = '')
    if ($Default) { $hint = " [" + $Default + "]" } else { $hint = "" }
    Write-Host ("  " + $Prompt + $hint + ": ") -ForegroundColor Yellow -NoNewline
    $val = Read-Host
    if ($val) { return $val } else { return $Default }
}

function Write-OK   { param([string]$M) Write-Host ("  [OK]   " + $M) -ForegroundColor Green  }
function Write-Warn { param([string]$M) Write-Host ("  [!]    " + $M) -ForegroundColor Yellow }
function Write-Err  { param([string]$M) Write-Host ("  [ERRO] " + $M) -ForegroundColor Red   }
function Write-Info { param([string]$M) Write-Host ("  >>     " + $M) -ForegroundColor Cyan  }

function Pause-Screen {
    Write-Host ""
    Write-Host "  Pressione qualquer tecla para continuar..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# =============================================================================
# Config
# =============================================================================

function Load-Config {
    param([string]$Path)
    $cfg = @{
        BaseURL    = 'http://172.16.0.12/znuny'
        Username   = ''
        Password   = ''
        EstadoFile = 'estado_chamados.json'
        OutputPath = $PWD.Path
        SearchPath = 'index.pl?Action=AgentKPISearch;Subaction=Search;TakeLastSearch=1;Profile=94_8'
    }
    if (Test-Path $Path) {
        try {
            $j = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($j.BaseURL)    { $cfg.BaseURL    = $j.BaseURL    }
            if ($j.Username)   { $cfg.Username   = $j.Username   }
            if ($j.Password)   { $cfg.Password   = $j.Password   }
            if ($j.EstadoFile) { $cfg.EstadoFile = $j.EstadoFile }
            if ($j.OutputPath) { $cfg.OutputPath = $j.OutputPath }
            if ($j.SearchPath) { $cfg.SearchPath = $j.SearchPath }
        } catch { }
    }
    return $cfg
}

function Save-Config {
    param([hashtable]$Cfg, [string]$Path)
    [ordered]@{
        BaseURL    = $Cfg.BaseURL
        Username   = $Cfg.Username
        Password   = $Cfg.Password
        EstadoFile = $Cfg.EstadoFile
        OutputPath = $Cfg.OutputPath
        SearchPath = $Cfg.SearchPath
    } | ConvertTo-Json | Out-File $Path -Encoding UTF8
}

function Get-CfgStatus {
    param([string]$Path)
    if (Test-Path $Path) { return "[Salvo] " + $Path } else { return "Nao salvo" }
}

# =============================================================================
# Verificacao da funcao Export-CcoReport (agora inline)
# =============================================================================

function Ensure-ExportScript {
    # A funcao Export-CcoReport ja esta definida inline neste mesmo script.
    # Nao e mais necessario carregar um arquivo externo.
    if (-not (Get-Command Export-CcoReport -ErrorAction SilentlyContinue)) {
        throw "Export-CcoReport ainda nao esta definida. Verifique se o codigo foi incorporado corretamente."
    }
    $script:ExportLoaded = $true
}

function Resolve-ExportScript {
    param([string]$Hint)
    return "embedded"
}

# =============================================================================
# Modulo/script CCO Export v11.9 - OTRS / Znuny
# (Cliente = ID do Cliente, Unidade = campo customizado)
# =============================================================================
# Compativel com Windows PowerShell 5.1 e PowerShell 7+
# =============================================================================
# Salve este arquivo como UTF-8 com BOM para preservar caracteres acentuados
# =============================================================================

# Forcar TLS 1.2 e ignorar certificados
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# =============================================================================
# Configuracoes centralizadas
# =============================================================================
$now = Get-Date
$standardHour = [math]::Floor([int]$now.Hour / 2) * 2
$standardHourStr = $standardHour.ToString('00')
$script:CcoConfig = @{
    TituloRelatorio  = "Chamados Cr" + [char]237 + "ticos - " + $now.ToString('dd/MM/yyyy') + " " + $standardHourStr + "h"
    TituloResolvidos = "Chamados Normalizados - " + $now.ToString('dd/MM/yyyy') + " " + $standardHourStr + "h"
    LabelHorario     = "Hor" + [char]225 + "rio de abertura"
    LabelOcorrencia  = "Ocorr" + [char]234 + "ncia"
    Separador        = "*--------------------------------------------------------------*"
    EstadosResolvidos = @('resolvido','fechado','removido','encerrado','resolved','closed','merged','agrupado')
    AutoNotePatterns = @(
        'Sistema de Monitoramento',
        'Host:.*IP:.*Incidente:',
        'Prioridade:\s*(Disaster|High|Average|Information)',
        'Agradecemos seu contato',
        'solicita.{1,10}ser.{1,5}tratada atrav',
        'sob os cuidados da nossa equipe',
        'Para mais informa.{1,10}contate',
        'Para proteger sua privacidade',
        'conte.do remoto foi desabilitado',
        'Carregar conte.do bloqueado',
        'Chamado enviada para campo',
        'atingiu.{1,20}(HORAS|HORA|DUAS)',
        'Para seu conhecimento.*O Ticket',
        'Notifica.{1,5}o de Eventos',
        'Por favor, verificar o Incidente conforme dados',
        'Raz.o Social:.*CNPJ:.*Dificuldade',
        'Alerta - \[Ticket#',
        'Direcionamento de Ticket',
        'ticket.*foi direcionado',
        'COMUNICADO.*chave de e.mail',
        'plantao.*fora do hor',
        'Problema com Status OK',
        'Problema persiste',
        'Gr.ficos personalizados',
        'Estado do link Zabbix',
        'Status:\s*(OK|PROBLEM|DISASTER)'
    )
}

# =============================================================================
# Classes
# =============================================================================
class TicketData {
    [string]$Numero
    [string]$Estado
    [string]$Criado
    [string]$Cliente
    [string]$Unidade
    [System.Collections.Generic.List[object]]$Articles = [System.Collections.Generic.List[object]]::new()

    TicketData([string]$numero, [string]$estado, [string]$criado, [string]$cliente, [string]$unidade, $articles) {
        $this.Numero  = $numero
        $this.Estado  = $estado
        $this.Criado  = $criado
        $this.Cliente = $cliente
        $this.Unidade = $unidade
        if ($articles) { $this.Articles = $articles }
    }
}

class OtrsClient {
    [string]$BaseURL
    [object]$Session
    [int]$RetryMax

    OtrsClient([string]$baseUrl, [int]$retryMax = 3) {
        $this.BaseURL  = $baseUrl.TrimEnd('/')
        $this.RetryMax = $retryMax
    }

    [void] Login([string]$user, [string]$pass) {
        $loginUrl = "$($this.BaseURL)/index.pl"
        $body = @{ Action='Login'; RequestedURL=''; Lang='pt_BR'; TimeOffset='-180'; User=$user; Password=$pass }
        try {
            $tempSession = $null
            $response = Invoke-WebRequest -Uri $loginUrl -Method POST -UseBasicParsing -Body $body -SessionVariable 'tempSession' -ErrorAction Stop
            $this.Session = $tempSession
            if ($response.Content -notmatch 'Action=Logout|/logout|Sair') {
                throw "Login falhou - resposta nao contem indicador de sessao ativa."
            }
            Write-Verbose "Login bem-sucedido em $($this.BaseURL)"
        } catch {
            throw "Erro no login: $_"
        }
    }

    [void] Logout() {
        if ($this.Session) {
            try { $this.Session.Close() } catch {}
            $this.Session = $null
        }
    }

    [string[]] GetActiveTicketIDs([string]$searchPath) {
        $uri = "$($this.BaseURL)/$searchPath"
        $response = $this.InvokeWithRetry($uri)
        $matches = [regex]::Matches($response.Content, 'TicketID=(\d+)')
        return $matches | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
    }

    [string] GetTicketHtml([string]$ticketID) {
        $uri = "$($this.BaseURL)/index.pl?Action=AgentTicketZoom;TicketID=$ticketID"
        $response = $this.InvokeWithRetry($uri)
        return $this.GetResponseText($response)
    }

    # Busca o widget CustomerInformation via AJAX
    [string] GetCustomerWidgetHtml([string]$ticketID, [string]$challengeToken) {
        $uri  = "$($this.BaseURL)/index.pl"
        $body = "Action=AgentTicketZoom;Subaction=LoadWidget;TicketID=$ticketID;ElementID=Async_0200-CustomerInformation;ChallengeToken=$challengeToken"
        $headers = @{
            'X-Requested-With' = 'XMLHttpRequest'
            'Content-Type'     = 'application/x-www-form-urlencoded; charset=UTF-8'
        }
        $attempt = 0
        while ($attempt -le $this.RetryMax) {
            try {
                $response = Invoke-WebRequest -Uri $uri -Method POST -Body $body `
                    -Headers $headers -WebSession $this.Session -UseBasicParsing -ErrorAction Stop
                return $this.GetResponseText($response)
            } catch {
                $attempt++
                if ($attempt -gt $this.RetryMax) { throw }
                Start-Sleep -Milliseconds (200 * $attempt)
                Write-Verbose "Retry $attempt/$($this.RetryMax) para widget CustomerInformation (ticket $ticketID)"
            }
        }
        return ''
    }

    [string] GetArticleContent([string]$ticketID, [string]$articleID) {
        $uri = "$($this.BaseURL)/index.pl?Action=AgentTicketArticleContent;Subaction=HTMLView;TicketID=$ticketID;ArticleID=$articleID;FileID=;"
        $response = $this.InvokeWithRetry($uri)
        return $this.GetResponseText($response)
    }

    hidden [object] InvokeWithRetry([string]$uri) {
        return $this.InvokeWithRetry($uri, 'GET', $null)
    }

    hidden [object] InvokeWithRetry([string]$uri, [string]$method) {
        return $this.InvokeWithRetry($uri, $method, $null)
    }

    hidden [object] InvokeWithRetry([string]$uri, [string]$method, $body) {
        $attempt = 0
        while ($attempt -le $this.RetryMax) {
            try {
                $params = @{
                    Uri             = $uri
                    WebSession      = $this.Session
                    UseBasicParsing = $true
                    ErrorAction     = 'Stop'
                    Method          = $method
                }
                if ($body) { $params.Body = $body }
                return Invoke-WebRequest @params
            } catch {
                $attempt++
                if ($attempt -gt $this.RetryMax) { throw }
                Start-Sleep -Milliseconds (200 * $attempt)
                Write-Verbose "Retry $attempt/$($this.RetryMax) para $uri"
            }
        }
        throw "Loop de retry excedido (nunca alcancado)"
    }

    hidden [string] GetResponseText($response) {
        if (-not $response -or -not $response.RawContentStream) { return '' }
        return [System.Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray())
    }
}

class TicketCache {
    [string]$FilePath
    [hashtable]$Data = @{}

    TicketCache([string]$path) {
        $this.FilePath = $path
        $this.Load()
    }

    [void] Load() {
        if (-not (Test-Path $this.FilePath)) { return }
        try {
            $json = Get-Content $this.FilePath -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($json)) { return }
            $obj = $json | ConvertFrom-Json
            $this.Data = @{}
            foreach ($prop in $obj.PSObject.Properties) {
                $v = $prop.Value
                $notas = @()
                if ($v.Notas) {
                    foreach ($n in $v.Notas) {
                        $notas += [PSCustomObject]@{ Date = $n.Date; Text = $n.Text }
                    }
                }
                $this.Data[$prop.Name] = @{
                    Estado   = $v.Estado
                    Cliente  = $v.Cliente
                    Unidade  = $v.Unidade
                    Criado   = $v.Criado
                    Numero   = $v.Numero
                    Notas    = $notas
                }
            }
        } catch {
            Write-Warning "Falha ao carregar cache: $_"
        }
    }

    [void] Save() {
        $obj = [ordered]@{}
        foreach ($key in $this.Data.Keys) {
            $obj[$key] = $this.Data[$key]
        }
        $obj | ConvertTo-Json -Depth 6 | Out-File -FilePath $this.FilePath -Encoding UTF8
    }

    [void] Update([string]$id, [TicketData]$ticket) {
        $this.Data[$id] = @{
            Estado   = $ticket.Estado
            Cliente  = $ticket.Cliente
            Unidade  = $ticket.Unidade
            Criado   = $ticket.Criado
            Numero   = $ticket.Numero
            Notas    = $ticket.Articles
        }
    }

    [bool] IsCachedAndResolved([string]$id) {
        if (-not $this.Data.ContainsKey($id)) { return $false }
        $estado = $this.Data[$id].Estado.ToLower().Trim()
        return $estado -in $script:CcoConfig.EstadosResolvidos
    }

    [TicketData] GetCachedTicket([string]$id) {
        $c = $this.Data[$id]
        if ($c.Notas) { $notas = $c.Notas } else { $notas = @() }
        return [TicketData]::new($c.Numero, $c.Estado, $c.Criado, $c.Cliente, $c.Unidade, $notas)
    }
}

# =============================================================================
# Funcoes de extracao
# =============================================================================
function Remove-Html {
    param([string]$Text)
    if (-not $Text) { return '' }
    try {
        $Text = $Text -replace '(?s)<style[^>]*>.*?</style>', ''
        $Text = $Text -replace '(?s)<script[^>]*>.*?</script>', ''
        $Text = $Text -replace '<[^>]+>', ''
        $Text = [System.Net.WebUtility]::HtmlDecode($Text)
    } catch {
        Write-Warning "Erro ao decodificar HTML: $_"
    }
    return ($Text -replace '\s+', ' ').Trim()
}

function Get-CustomerNameFromHtml {
    param([string]$HTML)

    $pattern1 = @'
(?s)<label[^>]*>\s*Nome\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
    if (($m = [regex]::Match($HTML, $pattern1)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val) {
            $val = $val -replace '\s*-\s*\d+\s*$', ''
            return $val.Trim()
        }
    }
    $pattern1b = @'
(?s)<label[^>]*>\s*Nome\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>
'@
    if (($m = [regex]::Match($HTML, $pattern1b)).Success) {
        $val = (Remove-Html $m.Groups[1].Value).Trim()
        if ($val) {
            $val = $val -replace '\s*-\s*\d+\s*$', ''
            return $val.Trim()
        }
    }

    $pattern2 = @'
(?s)<label[^>]*>\s*ID do Cliente\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
    if (($m = [regex]::Match($HTML, $pattern2)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val) { return $val }
    }
    if (($m = [regex]::Match($HTML, '(?s)<label[^>]*>\s*ID do Cliente\s*:\s*</label>\s*(?:<[^>]+>\s*)*<a[^>]*>([^<]+)</a>')).Success) {
        return $m.Groups[1].Value.Trim()
    }

    $pattern3 = @'
(?s)<label[^>]*>\s*Cliente\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
    if (($m = [regex]::Match($HTML, $pattern3)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val) { return $val }
    }
    if (($m = [regex]::Match($HTML, '(?s)<label[^>]*>\s*Cliente\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>')).Success) {
        $val = (Remove-Html $m.Groups[1].Value).Trim()
        if ($val) { return $val }
    }
    return 'N/D'
}

function Get-CustomerUnitFromHtml {
    param(
        [string]$HTML,
        [switch]$FallbackOnly
    )

    if (-not $FallbackOnly) {

        $patternCod = @'
(?s)<label[^>]*>\s*CodigoCliente\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
        if (($m = [regex]::Match($HTML, $patternCod)).Success) {
            $val = $m.Groups[1].Value.Trim()
            if ($val) { return $val }
        }
        if (($m = [regex]::Match($HTML, '(?s)<label[^>]*>\s*CodigoCliente\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>')).Success) {
            $val = (Remove-Html $m.Groups[1].Value).Trim()
            if ($val) { return $val }
        }

        # Usuario: (login)
        $patternUser = @'
(?s)<label[^>]*>\s*Usu.rio\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
        if (($m = [regex]::Match($HTML, $patternUser)).Success) {
            $val = $m.Groups[1].Value.Trim()
            if ($val -and $val -notmatch '@') { return $val }
        }
        $patternUser2 = @'
(?s)<label[^>]*>\s*Usu.rio\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>
'@
        if (($m = [regex]::Match($HTML, $patternUser2)).Success) {
            $val = (Remove-Html $m.Groups[1].Value).Trim()
            if ($val -and $val -notmatch '@') { return $val }
        }

        # Outros labels
        $unitLabels = @(
            'Unidade',
            'Loja',
            'C\s*o\s*d\s*i\s*g\s*o\s*Cliente',
            'Custom\s?[Uu]ser',
            'Customer\s?[Uu]ser'
        )
        foreach ($label in $unitLabels) {
            $pattern = "(?s)<label[^>]*>\s*$label\s*:\s*</label>\s*<p[^>]+title=`"([^`"]+)`""
            if (($m = [regex]::Match($HTML, $pattern)).Success) {
                $val = $m.Groups[1].Value.Trim()
                if ($val) { return $val }
            }
            $pattern2 = "(?s)<label[^>]*>\s*$label\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>"
            if (($m = [regex]::Match($HTML, $pattern2)).Success) {
                $val = (Remove-Html $m.Groups[1].Value).Trim()
                if ($val) { return $val }
            }
        }

        # Fallback: ultimo segmento numerico de "Nome:"
        $patternNome = @'
(?s)<label[^>]*>\s*Nome\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
        if (($m = [regex]::Match($HTML, $patternNome)).Success) {
            $nome = $m.Groups[1].Value.Trim()
            if ($nome -match '-\s*(\d{3,})\s*$') { return $matches[1] }
            if ($nome -match '-\s*([^\s-][^-]*)\s*$') { return $matches[1].Trim() }
        }
    }
    return 'N/D'
}

function Test-AutoNote {
    param([string]$Text)
    if ($Text.Length -lt 5) { return $true }
    foreach ($pattern in $script:CcoConfig.AutoNotePatterns) {
        if ($Text -match $pattern) { return $true }
    }
    return $false
}

function Get-TicketDataFromHtml {
    param(
        [string]$HTML,
        [string]$ticketID,
        [OtrsClient]$client,
        [int]$maxArticles,
        [int]$fetchLimit,
        [int]$sleepMs
    )

    if (($m = [regex]::Match($HTML, 'Ticket#([0-9]+)')).Success) { $numero = $m.Groups[1].Value } else { $numero = 'N/D' }
    if (($m = [regex]::Match($HTML, '(?s)<span[^>]+class="[^"]*pill[^"]*"[^>]+title="([^"]+)"')).Success) { $estado = $m.Groups[1].Value }
    elseif (($m = [regex]::Match($HTML, '(?s)label[^>]*>\s*Estado:\s*[^<]*</label>\s*<span[^>]+title="([^"]+)"')).Success) { $estado = $m.Groups[1].Value }
    else { $estado = 'N/D' }
    if (($m = [regex]::Match($HTML, '(?s)label[^>]*>\s*Criado:\s*</label>\s*<p[^>]+title="([^"]+)"')).Success) { $criado = $m.Groups[1].Value } else { $criado = 'N/D' }

    $token = ''
    $tokenPatterns = @(
        'name="ChallengeToken"\s+value="([A-Za-z0-9]+)"',
        'value="([A-Za-z0-9]+)"\s+name="ChallengeToken"',
        '"ChallengeToken",\s*"([A-Za-z0-9]+)"',
        'ChallengeToken=([A-Za-z0-9]{{16,}})',
        'data-challenge-token="([A-Za-z0-9]+)"',
        'ChallengeToken:\s*"([A-Za-z0-9]+)"'
    )
    foreach ($tp in $tokenPatterns) {
        if (($m = [regex]::Match($HTML, $tp)).Success) {
            $token = $m.Groups[1].Value
            Write-Verbose "ChallengeToken encontrado no HTML principal (ticket $ticketID): $token"
            break
        }
    }
    if (-not $token) {
        Write-Verbose "ChallengeToken NAO encontrado no HTML principal do ticket $ticketID"
    }

    $widgetHtml = ''
    if ($token -and $client) {
        try {
            $widgetHtml = $client.GetCustomerWidgetHtml($ticketID, $token)
            if ($widgetHtml) {
                Write-Verbose "Widget CustomerInformation obtido: $($widgetHtml.Length) chars (ticket $ticketID)"
            } else {
                Write-Verbose "Widget CustomerInformation retornou VAZIO para ticket $ticketID"
            }
        } catch {
            Write-Verbose "Widget CustomerInformation indisponivel para ticket $ticketID : $_"
        }
    } else {
        Write-Verbose "ChallengeToken ausente ou cliente invalido - widget nao buscado (ticket $ticketID)"
    }

    $cliente = 'N/D'
    $unidade = 'N/D'
    foreach ($src in @($widgetHtml, $HTML) | Where-Object { $_ }) {
        if ($cliente -eq 'N/D') { $cliente = Get-CustomerNameFromHtml $src }
        if ($unidade -eq 'N/D') { $unidade = Get-CustomerUnitFromHtml $src }
        if ($cliente -ne 'N/D' -and $unidade -ne 'N/D') { break }
    }
    Write-Verbose "Resultado extracao: Cliente=$cliente, Unidade=$unidade (ticket $ticketID)"

    # Ajuste: se unidade == cliente, algo saiu errado - tenta campo alternativo
    if ($unidade -ne 'N/D' -and $unidade -ceq $cliente) {
        Write-Verbose "Unidade igual ao Cliente, tentando campo alternativo..."
        $searchSrc = $widgetHtml + $HTML

        $patternAlt1 = @'
(?s)<label[^>]*>\s*Usu.rio\s*:\s*</label>\s*<p[^>]+title="([^"]+)"
'@
        if (($m = [regex]::Match($searchSrc, $patternAlt1)).Success) {
            $alt = $m.Groups[1].Value.Trim()
            if ($alt -and $alt -notmatch '@') { $unidade = $alt }
        }
        else {
            $patternAlt2 = '(?s)<label[^>]*>\s*CodigoCliente\s*:\s*</label>\s*<p[^>]+title="([^"]+)"'
            if (($m = [regex]::Match($searchSrc, $patternAlt2)).Success) {
                $alt = $m.Groups[1].Value.Trim()
                if ($alt) { $unidade = $alt }
            }
        }
        Write-Verbose "Apos ajuste: Cliente=$cliente, Unidade=$unidade"
    }

    $articles = [System.Collections.Generic.List[object]]::new()
    $rowPattern = @'
(?s)<tr[^>]*id="Row\d+"[^>]*>(.*?)</tr>
'@
    $rowMatches = [regex]::Matches($HTML, $rowPattern)

    foreach ($row in $rowMatches) {
        $rowHtml = $row.Groups[1].Value
        if (($m = [regex]::Match($rowHtml, '<input[^>]+class="ArticleID"[^>]+value="(\d+)"')).Success) { $aid = $m.Groups[1].Value } else { $aid = $null }
        if (($m = [regex]::Match($rowHtml, '<td class="Created">[^<]*<div title="([^"]+)"')).Success) { $adate = $m.Groups[1].Value } else { $adate = 'N/D' }
        if (-not $aid) { continue }
        try {
            $raw = $client.GetArticleContent($ticketID, $aid)
            $raw = $raw -replace '(?s)<style[^>]*>.*?</style>', ''
            $raw = $raw -replace '(?s)<script[^>]*>.*?</script>', ''
            if (($m = [regex]::Match($raw, '(?s)<body[^>]*>(.*)</body>')).Success) { $body = $m.Groups[1].Value.Trim() } else { $body = '' }
            if (-not $body) { continue }
            $text = Remove-Html $body

            if (($am = [regex]::Match($text, 'Analista designado[:\s]*@([\w\s]+)')).Success) {
                $text = "Em tratativas com N2 @$($am.Groups[1].Value.Trim())"
            }
            $text = $text -replace '\s*\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s*-\s*\S+@\S+\s+escreveu:.*$', ''
            $text = $text -replace '\s*Em\s+\d{2}/\d{2}/\d{4}.*?escreveu:.*$', ''
            $text = ($text -replace '\s+', ' ').Trim()

            if ($text.Length -ge 5 -and -not (Test-AutoNote $text)) {
                $articles.Add([PSCustomObject]@{ Date = $adate; Text = $text })
            }
            Start-Sleep -Milliseconds $sleepMs
        } catch { Write-Verbose "Erro ao obter artigo $aid para ticket $ticketID" }
    }
    $articles.Reverse()
    return [TicketData]::new($numero, $estado, $criado, $cliente, $unidade, $articles)
}

# =============================================================================
# Formatacao e funcao principal
# =============================================================================
function ConvertTo-TicketBlock {
    param([TicketData]$Ticket)
    $c = $script:CcoConfig
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("*Ticket#$($Ticket.Numero) - $($Ticket.Estado)*")
    $lines.Add("*$($c.LabelHorario): $($Ticket.Criado)*")
    $lines.Add("*Cliente: $($Ticket.Cliente)*")
    $lines.Add("*Campo do cliente: $($Ticket.Unidade)*")
    $lines.Add("*$($c.LabelOcorrencia): *")
    $lines.Add("")
    foreach ($n in $Ticket.Articles) {
        $lines.Add("*$($n.Date) status:* $($n.Text)")
    }
    $lines.Add("")
    $lines.Add($c.Separador)
    $lines.Add("")
    return $lines
}

function New-CcoFileContent {
    param(
        [TicketData[]]$Tickets,
        [string]$Titulo,
        [string]$EmptyMessage = "Nenhum chamado ativo no momento."
    )
    $content = [System.Collections.Generic.List[string]]::new()
    $c = $script:CcoConfig
    $content.Add($c.Separador)
    $content.Add("                    *$Titulo*")
    $content.Add($c.Separador)
    $content.Add("")

    $count = 0
    foreach ($t in $Tickets) {
        foreach ($linha in (ConvertTo-TicketBlock $t)) {
            $content.Add([string]$linha)
        }
        $count++
    }
    if ($count -eq 0) { $content.Add($EmptyMessage) }
    return $content
}

function Export-CcoReport {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$BaseURL,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][string]$Password,
        [string]$SearchPath = "index.pl?Action=AgentKPISearch;Subaction=Search;TakeLastSearch=1;Profile=94_8",
        [string]$EstadoFile = "estado_chamados.json",
        [string]$OutputDir = $PWD,
        [int]$MaxArticles = 9999,
        [int]$FetchLimit = 9999,
        [int]$RetryMax = 3,
        [int]$SleepArticleMs = 50,
        [int]$SleepTicketMs = 100,
        [switch]$AbrirRelatorio,
        [switch]$DiagMode
    )

    $now = Get-Date
    $stdHour = [math]::Floor([int]$now.Hour / 2) * 2
    $dateStr = $now.ToString('dd-MM-yyyy')
    $hourStr = $stdHour.ToString('00') + 'h'

    $ativoPath     = Join-Path $OutputDir "Chamados_Criticos_${dateStr}_${hourStr}.txt"
    $resolvidoPath = Join-Path $OutputDir "Chamados_Criticos_Resolvidos_${dateStr}_${hourStr}.txt"

    Write-Verbose "BaseURL: $BaseURL"
    Write-Verbose "Perfil: $SearchPath"
    Write-Verbose "Cache: $EstadoFile"

    $client = [OtrsClient]::new($BaseURL, $RetryMax)
    $client.Login($Username, $Password)

    if ($DiagMode) {
        $diagIds = $client.GetActiveTicketIDs($SearchPath)
        if (-not $diagIds) { Write-Warning "Nenhum ticket encontrado para diagnostico."; $client.Logout(); return }
        $diagId  = $diagIds[0]
        Write-Host "`n[DIAG] Inspecionando ticket $diagId ..." -ForegroundColor Cyan

        $diagMain = $client.GetTicketHtml($diagId)
        $diagMainPath = Join-Path $OutputDir "diag_main_$diagId.html"
        $diagMain | Out-File $diagMainPath -Encoding UTF8
        Write-Host "[DIAG] HTML principal -> $diagMainPath ($($diagMain.Length) chars)" -ForegroundColor Yellow

        $diagToken = ''
        foreach ($tp in @(
            'name="ChallengeToken"\s+value="([A-Za-z0-9]+)"',
            'value="([A-Za-z0-9]+)"\s+name="ChallengeToken"',
            '"ChallengeToken",\s*"([A-Za-z0-9]+)"',
            'ChallengeToken=([A-Za-z0-9]{{16,}})',
            'data-challenge-token="([A-Za-z0-9]+)"'
        )) {
            if (($m = [regex]::Match($diagMain, $tp)).Success) { $diagToken = $m.Groups[1].Value; break }
        }

        if ($diagToken) {
            $diagWidget = $client.GetCustomerWidgetHtml($diagId, $diagToken)
            $diagWidgetPath = Join-Path $OutputDir "diag_widget_$diagId.html"
            $diagWidget | Out-File $diagWidgetPath -Encoding UTF8
            Write-Host "[DIAG] Widget CustomerInformation -> $diagWidgetPath ($($diagWidget.Length) chars)" -ForegroundColor Yellow
            Write-Host "[DIAG] ChallengeToken: $diagToken" -ForegroundColor Gray
        } else {
            Write-Warning "[DIAG] ChallengeToken nao encontrado - widget nao buscado. Verifique diag_main_$diagId.html."
        }

        Write-Host "`n[DIAG] Abra os arquivos acima e procure pelas labels 'Cliente' e 'Unidade' para confirmar a estrutura HTML.`n" -ForegroundColor Cyan
        $client.Logout()
        return
    }

    $activeIds = $client.GetActiveTicketIDs($SearchPath)
    $cache = [TicketCache]::new($EstadoFile)

    Write-Verbose "Ativos na busca: $($activeIds.Count)"
    Write-Verbose "Em cache: $($cache.Data.Count)"

    $allIds = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $activeIds) { $allIds.Add($id) }
    foreach ($id in $cache.Data.Keys) {
        if ($id -notin $allIds) { $allIds.Add($id) }
    }
    $allIds = $allIds | Sort-Object { [int]$_ }

    Write-Verbose "Total a processar: $($allIds.Count)"

    $ativos    = [System.Collections.Generic.List[TicketData]]::new()
    $resolvidos = [System.Collections.Generic.List[TicketData]]::new()
    $cacheHits = 0
    $httpCount = 0

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($tid in $allIds) {
        Write-Progress -Activity "Processando tickets" -Status "Ticket $tid" -PercentComplete (($httpCount+$cacheHits)/$allIds.Count*100)

        if ($cache.IsCachedAndResolved($tid) -and $tid -notin $activeIds) {
            Write-Verbose "TID $tid - CACHE"
            $ticket = $cache.GetCachedTicket($tid)
            $resolvidos.Add($ticket)
            $cacheHits++
            continue
        }

        Write-Verbose "TID $tid - HTTP"
        try {
            $html = $client.GetTicketHtml($tid)
            $ticket = Get-TicketDataFromHtml -HTML $html -ticketID $tid -client $client `
                -maxArticles $MaxArticles -fetchLimit $FetchLimit -sleepMs $SleepArticleMs
            if (-not $ticket) {
                Write-Warning "Nao foi possivel extrair dados do ticket $tid"
                continue
            }
            $cache.Update($tid, $ticket)

            if ($ticket.Estado.ToLower().Trim() -in $script:CcoConfig.EstadosResolvidos) {
                $resolvidos.Add($ticket)
            } else {
                $ativos.Add($ticket)
            }
            $httpCount++
            Start-Sleep -Milliseconds $SleepTicketMs
        } catch {
            Write-Warning "Falha ao processar ticket $tid : $_"
        }
    }
    $stopwatch.Stop()
    Write-Progress -Activity "Processando tickets" -Completed

    $cache.Save()

    if ($PSCmdlet.ShouldProcess("Gerar relatorios", "Criar arquivos CCO e Resolvidos")) {
        $ativoContent     = New-CcoFileContent -Tickets $ativos    -Titulo $script:CcoConfig.TituloRelatorio -EmptyMessage "Nenhum chamado critico ativo no momento."
        $resolvidoContent = New-CcoFileContent -Tickets $resolvidos -Titulo $script:CcoConfig.TituloResolvidos -EmptyMessage "Nenhum chamado normalizado."

        $ativoContent | Out-File $ativoPath -Encoding UTF8
        if ($resolvidos.Count -gt 0) {
            $resolvidoContent | Out-File $resolvidoPath -Encoding UTF8
            Write-Verbose "Resolvidos salvos em $resolvidoPath"
        }
        Write-Host "Relatorio Critico: $ativoPath" -ForegroundColor Cyan
    }

    Write-Host "Processados: $($allIds.Count) | Cache: $cacheHits | HTTP: $httpCount | Tempo: $($stopwatch.Elapsed.ToString('mm\:ss'))" -ForegroundColor Green

    if ($AbrirRelatorio) {
        Start-Process notepad.exe $ativoPath
        if ($resolvidos.Count -gt 0) {
            Start-Process notepad.exe $resolvidoPath
        }
    }

    $client.Logout()
}

# =============================================================================
# Login
# =============================================================================

function Show-LoginScreen {
    param([hashtable]$Cfg)
    Write-Banner
    Write-Centered "-- AUTENTICACAO --" 'White'
    Write-Host ""
    Write-Info "Enter = manter o valor entre [ ]"
    Write-Host ""
    $Cfg.BaseURL  = Read-Field "URL do servidor" $Cfg.BaseURL
    $Cfg.Username = Read-Field "Usuario"         $Cfg.Username
    if ($Cfg.Password) {
        Write-Host "  Senha [****]: " -ForegroundColor Yellow -NoNewline
        $inp = Read-Host
        if ($inp) { $Cfg.Password = $inp }
    } else {
        $Cfg.Password = Read-PasswordSecure
    }
    Write-Host ""
    return $Cfg
}

# =============================================================================
# Gerar TXT
# =============================================================================

function Invoke-GerarTxt {
    param([hashtable]$Cfg, [string]$ExportScript)
    Write-Banner
    Write-Centered "-- GERANDO RELATORIO TXT --" 'White'
    Write-Host ""
    Write-Info ("Servidor: " + $Cfg.BaseURL)
    Write-Info ("Usuario : " + $Cfg.Username)
    Write-Info ("Saida   : " + $Cfg.OutputPath)
    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Host ""
    try {
        Ensure-ExportScript $ExportScript
    } catch {
        Write-Err ("Falha ao carregar script: " + $_); Pause-Screen; return
    }
    $params = @{
        BaseURL        = $Cfg.BaseURL
        Username       = $Cfg.Username
        Password       = $Cfg.Password
        SearchPath     = $Cfg.SearchPath
        EstadoFile     = $Cfg.EstadoFile
        OutputDir      = $Cfg.OutputPath
        AbrirRelatorio = $false
        DiagMode       = $false
    }
    try {
        Export-CcoReport @params
        Write-Host ""
        Write-OK "Relatorio TXT gerado com sucesso."
    } catch {
        Write-Host ""
        Write-Err ("Erro: " + $_)
    }
    Pause-Screen
}

# =============================================================================
# Gerar JSON
# =============================================================================

function Invoke-GerarJson {
    param([hashtable]$Cfg, [string]$ExportScript)
    Write-Banner
    Write-Centered "-- GERANDO RELATORIO JSON --" 'White'
    Write-Host ""
    Write-Info ("Servidor: " + $Cfg.BaseURL)
    Write-Info ("Usuario : " + $Cfg.Username)
    Write-Info ("Saida   : " + $Cfg.OutputPath)
    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Host ""
    try {
        Ensure-ExportScript $ExportScript
    } catch {
        Write-Err ("Falha ao carregar script: " + $_); Pause-Screen; return
    }
    $params = @{
        BaseURL        = $Cfg.BaseURL
        Username       = $Cfg.Username
        Password       = $Cfg.Password
        SearchPath     = $Cfg.SearchPath
        EstadoFile     = $Cfg.EstadoFile
        OutputDir      = $Cfg.OutputPath
        AbrirRelatorio = $false
        DiagMode       = $false
    }
    try {
        Export-CcoReport @params
    } catch {
        Write-Err ("Erro na exportacao: " + $_); Pause-Screen; return
    }

    $cachePath = $Cfg.EstadoFile
    if (-not [System.IO.Path]::IsPathRooted($cachePath)) {
        $cachePath = Join-Path $Cfg.OutputPath $cachePath
    }
    if (-not (Test-Path $cachePath)) {
        Write-Warn "Cache nao encontrado, nenhum JSON exportado."
        Pause-Screen; return
    }

    $now      = Get-Date
    $jsonName = "Relatorio_CCO_" + $now.ToString('yyyy-MM-dd_HH-mm') + ".json"
    $jsonPath = Join-Path $Cfg.OutputPath $jsonName

    $rawCache = Get-Content $cachePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $ativos     = [System.Collections.Generic.List[object]]::new()
    $resolvidos = [System.Collections.Generic.List[object]]::new()
    $estadosResolvidos = @('resolvido','fechado','removido','encerrado','resolved','closed','merged','agrupado')

    foreach ($prop in $rawCache.PSObject.Properties) {
        $t = $prop.Value
        $obj = [ordered]@{
            ID      = $prop.Name
            Numero  = $t.Numero
            Estado  = $t.Estado
            Criado  = $t.Criado
            Cliente = $t.Cliente
            Unidade = $t.Unidade
            Notas   = @()
        }
        if ($t.Notas) {
            $notaList = [System.Collections.Generic.List[object]]::new()
            foreach ($n in $t.Notas) {
                $notaList.Add([ordered]@{ Data = $n.Date; Texto = $n.Text })
            }
            $obj.Notas = $notaList
        }
        if ($t.Estado -and ($t.Estado.ToLower().Trim() -in $estadosResolvidos)) {
            $resolvidos.Add($obj)
        } else {
            $ativos.Add($obj)
        }
    }

    $report = [ordered]@{
        Gerado      = $now.ToString('dd/MM/yyyy HH:mm:ss')
        TotalAtivos = $ativos.Count
        TotalResolvidos = $resolvidos.Count
        Ativos      = $ativos
        Resolvidos  = $resolvidos
    }
    $report | ConvertTo-Json -Depth 8 | Out-File $jsonPath -Encoding UTF8

    Write-Host ""
    Write-OK ("JSON salvo em: " + $jsonPath)
    Write-Info ("Ativos: " + $ativos.Count + "   Resolvidos: " + $resolvidos.Count)
    Pause-Screen
}

# =============================================================================
# VISUALIZADOR DE CHAMADOS
# =============================================================================


# =============================================================================
# Tracking de normalizacao entre refreshes
# =============================================================================
$script:EstadosAnteriores = @{}

function Update-Normalizados {
    param($Tickets)
    $resolvidos = @('resolvido','fechado','removido','encerrado','resolved','closed','merged','agrupado')
    $norm = [System.Collections.Generic.List[object]]::new()
    foreach ($t in $Tickets) {
        $num = $t.Numero
        if (-not $num) { continue }
        if ($t.Estado) { $est = $t.Estado.ToLower().Trim() } else { $est = '' }
        if ($script:EstadosAnteriores.ContainsKey($num)) {
            if ($script:EstadosAnteriores[$num] -ne $est -and ($resolvidos -contains $est)) {
                $norm.Add($t)
            }
        }
        $script:EstadosAnteriores[$num] = $est
    }
    return @($norm)
}

function Show-NormalizacaoAlert {
    param($Norm)
    if ($Norm.Count -eq 0) { return }
    $W = [Math]::Max(50, [Console]::WindowWidth - 1)
    [Console]::Clear()
    Write-Host ('*' * $W) -ForegroundColor Green
    Write-Centered " CHAMADOS NORMALIZADOS " 'Green'
    Write-Host ""
    foreach ($t in $Norm) {
        Write-Host ("  [OK] #" + $t.Numero + "  " + $t.Cliente) -ForegroundColor Green
        Write-Host ("       Novo estado: " + $t.Estado) -ForegroundColor DarkGray
        Write-Host ""
    }
    Write-Host ('*' * $W) -ForegroundColor Green
    Write-Host ""
    Write-Host "  Pressione qualquer tecla para continuar..." -ForegroundColor White
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Get-LiveTickets {
    param([hashtable]$Cfg)
    $client = [OtrsClient]::new($Cfg.BaseURL, 45)
    try {
        $client.Login($Cfg.Username, $Cfg.Password)
        $ids = $client.GetActiveTicketIDs($Cfg.SearchPath)
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($id in $ids) {
            try {
                $html   = $client.GetTicketHtml($id)
                $ticket = Get-TicketDataFromHtml -HTML $html -ticketID $id -client $client `
                    -maxArticles 9999 -fetchLimit 9999 -sleepMs 50
                if ($null -ne $ticket -and $ticket.Articles.Count -gt 4) {
                    $short = [System.Collections.Generic.List[object]]::new()
                    for ($i = 0; $i -lt 4; $i++) { $short.Add($ticket.Articles[$i]) }
                    $ticket = [TicketData]::new($ticket.Numero, $ticket.Estado, $ticket.Criado, $ticket.Cliente, $ticket.Unidade, $short)
                }
                $list.Add($ticket)
            } catch { }
        }
        return @($list)
    } finally {
        $client.Logout()
    }
}

# =============================================================================
# Visualizador de chamados
# =============================================================================

function Format-NoteDate {
    param([string]$D)
    if ($D -match '(\d{2}/\d{2})/\d{4} (\d{2}:\d{2})') { return ($Matches[1] + " " + $Matches[2]) }
    if ($D.Length -gt 11) { return $D.Substring(0, 11) }
    return $D
}

function Show-TicketCard {
    param(
        $Ticket,
        [int]$TicketIdx,
        [int]$TicketTotal,
        [int]$NoteOffset,
        [int]$MaxNotes   = 0,   # 0 = dinamico por altura da janela
        [int]$RefreshSec = -1   # -1 = sem timer; >= 0 mostra countdown
    )

    $W = [Math]::Max(50, [Console]::WindowWidth - 1)
    $H = [Console]::WindowHeight

    # Linhas fixas: header(3) + blank + 5 info + blank + 3 notas-hdr + footer(2) + blank = 16
    # Se tem timer, +1 linha no footer
    if ($RefreshSec -ge 0) { $FIXED = 19 } else { $FIXED = 18 }
    if ($MaxNotes -gt 0) {
        $visibleNotes = $MaxNotes
    } else {
        $visibleNotes = [Math]::Max(1, $H - $FIXED)
    }

    $resolvidos = @('resolvido','fechado','removido','encerrado','resolved','closed','merged','agrupado')
    if ($Ticket.Estado) { $isRes = $resolvidos -contains $Ticket.Estado.ToLower().Trim() } else { $isRes = $false }

    $arts      = @($Ticket.Articles)
    $totalNotes = $arts.Count
    $maxOffset = [Math]::Max(0, $totalNotes - $visibleNotes)
    if ($NoteOffset -gt $maxOffset) { $NoteOffset = $maxOffset }
    if ($NoteOffset -lt 0) { $NoteOffset = 0 }

    [Console]::Clear()

    # Cabecalho
    Write-Host ('=' * $W) -ForegroundColor DarkCyan
    $counter = "Chamado " + ($TicketIdx + 1) + "/" + $TicketTotal
    if ($RefreshSec -ge 0) {
        $nav = "[<][>] Ticket  [^][v] Notas  [R] Refresh(" + $RefreshSec + "s)  [Q] Sair"
    } else {
        $nav = "[<][>] Ticket  [^][v] Notas  [R] Reload  [Q] Sair"
    }
    $hdr = " CCO Viewer | " + $counter + " | " + $nav
    if ($hdr.Length -gt $W) { $hdr = $hdr.Substring(0, $W) } else { $hdr = $hdr.PadRight($W) }
    Write-Host $hdr -ForegroundColor Cyan
    Write-Host ('=' * $W) -ForegroundColor DarkCyan
    Write-Host ""

    # Info do ticket
    if ($isRes) { $tag = "[RESOLVIDO]" ; $tagColor = "DarkGray" }
    else        { $tag = "[ATIVO]"     ; $tagColor = "Green"    }

    $numTxt = "  Ticket #" + $Ticket.Numero
    $padLen = $W - $numTxt.Length - $tag.Length - 1
    if ($padLen -lt 0) { $padLen = 0 }
    Write-Host ($numTxt + (' ' * $padLen)) -NoNewline -ForegroundColor White
    Write-Host $tag -ForegroundColor $tagColor

    $eLine = "  Estado  : " + $Ticket.Estado
    if ($eLine.Length -gt $W) { $eLine = $eLine.Substring(0, $W) }
    if ($isRes) { $eColor = "DarkGray" } else { $eColor = "Yellow" }
    Write-Host $eLine -ForegroundColor $eColor

    $aLine = "  Abertura: " + $Ticket.Criado
    if ($aLine.Length -gt $W) { $aLine = $aLine.Substring(0, $W) }
    Write-Host $aLine -ForegroundColor Gray

    $cLine = "  Cliente : " + $Ticket.Cliente
    if ($cLine.Length -gt ($W - 3)) { $cLine = $cLine.Substring(0, $W - 3) + "..." }
    Write-Host $cLine -ForegroundColor White

    $uLine = "  Unidade : " + $Ticket.Unidade
    if ($uLine.Length -gt ($W - 3)) { $uLine = $uLine.Substring(0, $W - 3) + "..." }
    Write-Host $uLine -ForegroundColor White
    Write-Host ""

    # Secao de notas
    Write-Host ('-' * $W) -ForegroundColor DarkGray
    $endIdx = $NoteOffset + $visibleNotes - 1
    if ($endIdx -ge $totalNotes) { $endIdx = $totalNotes - 1 }

    if ($totalNotes -eq 0) {
        $notesHdr = "  Notas: nenhuma registrada"
    } else {
        $from = $NoteOffset + 1
        if ($endIdx -ge 0) { $to = $endIdx + 1 } else { $to = 0 }
        $notesHdr = "  Notas " + $from + "-" + $to + " de " + $totalNotes
        if ($MaxNotes -gt 0) { $notesHdr += "  (modo tempo real: " + $MaxNotes + " mais recentes)" }
        elseif ($totalNotes -gt $visibleNotes) { $notesHdr += "  [^][v] para rolar" }
    }
    if ($notesHdr.Length -gt $W) { $notesHdr = $notesHdr.Substring(0, $W) }
    Write-Host $notesHdr -ForegroundColor Cyan
    Write-Host ('-' * $W) -ForegroundColor DarkGray

    $prefixLen = 22
    $textW     = [Math]::Max(10, $W - $prefixLen - 1)
    $rendered  = 0

    if ($totalNotes -gt 0 -and $endIdx -ge 0) {
        for ($ni = $NoteOffset; $ni -le $endIdx; $ni++) {
            $nota    = $arts[$ni]
            $dateFmt = Format-NoteDate $nota.Date
            $numPart = ($ni + 1).ToString().PadLeft(2)
            $prefix  = "  [" + $numPart + "] " + $dateFmt + "  "
            $txt = ($nota.Text -replace '\s+', ' ').Trim()
            if ($txt.Length -gt $textW) { $txt = $txt.Substring(0, $textW - 3) + "..." }
            Write-Host $prefix -ForegroundColor DarkGray -NoNewline
            Write-Host $txt    -ForegroundColor White
            $rendered++
        }
    }

    $blank = $visibleNotes - $rendered
    for ($b = 0; $b -lt $blank; $b++) { Write-Host "" }

    Write-Host ('=' * $W) -ForegroundColor DarkGray
    $foot = "  [<-] Anterior   [->] Proximo   [^] Nota Acima   [v] Nota Abaixo   [Q] Menu"
    if ($foot.Length -gt $W) { $foot = $foot.Substring(0, $W) }
    Write-Host $foot -ForegroundColor DarkGray
}

function Load-TicketsFromCache {
    param([string]$CachePath, [string]$ExportScript)
    Ensure-ExportScript $ExportScript
    if (-not (Test-Path $CachePath)) { return @() }
    $cache = [TicketCache]::new($CachePath)
    $list  = [System.Collections.Generic.List[object]]::new()
    foreach ($id in ($cache.Data.Keys | Sort-Object { [int]$_ })) {
        $list.Add($cache.GetCachedTicket($id))
    }
    return @($list)
}

# Modo completo: leitura do cache, todas as notas, navegacao
function Show-VisualizadorCompleto {
    param([hashtable]$Cfg, [string]$ExportScript)

    $cachePath = $Cfg.EstadoFile
    if (-not [System.IO.Path]::IsPathRooted($cachePath)) {
        $cachePath = Join-Path $Cfg.OutputPath $cachePath
    }

    if (-not (Test-Path $cachePath)) {
        Write-Banner
        Write-Warn ("Cache nao encontrado: " + $cachePath)
        Write-Info "Gere o relatorio primeiro (opcao 1 ou 2)."
        Write-Host ""
        Write-Host "  Buscar do OTRS agora? [s/N]: " -ForegroundColor Yellow -NoNewline
        $resp = Read-Host
        if ($resp -notmatch '^[Ss]') { return }
        Invoke-GerarTxt $Cfg $ExportScript
        if (-not (Test-Path $cachePath)) { return }
    }

    try {
        $tickets = Load-TicketsFromCache $cachePath $ExportScript
    } catch {
        Write-Err ("Falha ao carregar cache: " + $_); Pause-Screen; return
    }

    if ($tickets.Count -eq 0) {
        Write-Banner; Write-Warn "Nenhum chamado no cache."; Pause-Screen; return
    }

    $null = Update-Normalizados $tickets

    $idx = 0 ; $noteOffset = 0

    while ($true) {
        $arts      = @($tickets[$idx].Articles)
        $maxOffset = [Math]::Max(0, $arts.Count - [Math]::Max(1, [Console]::WindowHeight - 18))

        Show-TicketCard $tickets[$idx] $idx $tickets.Count $noteOffset

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $vk  = [int]$key.VirtualKeyCode
        $ch  = [string]$key.Character

        if     ($vk -eq 39 -or $vk -eq 34) { if ($idx -lt ($tickets.Count-1)) { $idx++ } else { $idx=0 } ; $noteOffset=0 }
        elseif ($vk -eq 37 -or $vk -eq 33) { if ($idx -gt 0) { $idx-- } else { $idx=$tickets.Count-1 }   ; $noteOffset=0 }
        elseif ($vk -eq 40) { if ($noteOffset -lt $maxOffset) { $noteOffset++ } }
        elseif ($vk -eq 38) { if ($noteOffset -gt 0) { $noteOffset-- } }
        elseif ($ch -eq 'r' -or $ch -eq 'R') {
            try {
                $tickets = Load-TicketsFromCache $cachePath $ExportScript
                $norm = Update-Normalizados $tickets
                if ($idx -ge $tickets.Count) { $idx = 0 }
                if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
            } catch { }
            $noteOffset = 0
        }
        elseif ($ch -eq 'q' -or $ch -eq 'Q' -or $vk -eq 27) { [Console]::Clear(); return }
    }
}

# Modo tempo real: busca do OTRS a cada 60s, 4 ultimas notas
function Show-VisualizadorRealTime {
    param([hashtable]$Cfg, [string]$ExportScript)

    $REFRESH_SEC = 60

    Write-Banner
    Write-Centered "-- BUSCANDO DO OTRS (TEMPO REAL) --" 'White'
    Write-Host ""
    Write-Info ("Servidor: " + $Cfg.BaseURL)
    Write-Info ("Usuario : " + $Cfg.Username)
    Write-Info "Conectando e buscando chamados ativos..."
    Write-Host ""

    try {
        Ensure-ExportScript $ExportScript
        $tickets = Get-LiveTickets $Cfg
    } catch {
        Write-Err ("Falha ao buscar: " + $_); Pause-Screen; return
    }

    if ($tickets.Count -eq 0) {
        Write-Warn "Nenhum chamado ativo encontrado."; Pause-Screen; return
    }

    $null = Update-Normalizados $tickets

    $idx         = 0
    $lastRefresh = [DateTime]::Now

    while ($true) {
        $secsLeft = $REFRESH_SEC - [int]([DateTime]::Now - $lastRefresh).TotalSeconds
        if ($secsLeft -lt 0) { $secsLeft = 0 }

        Show-TicketCard $tickets[$idx] $idx $tickets.Count 0 4 $secsLeft

        # Espera ate 500ms por tecla
        $waited = 0 ; $keyFound = $false
        while ($waited -lt 500) {
            if ($Host.UI.RawUI.KeyAvailable) { $keyFound = $true; break }
            Start-Sleep -Milliseconds 100
            $waited += 100
        }

        # Verifica se chegou hora do refresh automatico
        $elapsed = ([DateTime]::Now - $lastRefresh).TotalSeconds
        if (-not $keyFound -and $elapsed -ge $REFRESH_SEC) {
            try {
                $newT = Get-LiveTickets $Cfg
                $norm = Update-Normalizados $newT
                $tickets = $newT
                if ($idx -ge $tickets.Count) { $idx = 0 }
                $lastRefresh = [DateTime]::Now
                if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
            } catch { }
            continue
        }

        if (-not $keyFound) { continue }

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $vk  = [int]$key.VirtualKeyCode
        $ch  = [string]$key.Character

        if     ($vk -eq 39 -or $vk -eq 34) { if ($idx -lt ($tickets.Count-1)) { $idx++ } else { $idx=0 } }
        elseif ($vk -eq 37 -or $vk -eq 33) { if ($idx -gt 0) { $idx-- } else { $idx=$tickets.Count-1 } }
        elseif ($ch -eq 'r' -or $ch -eq 'R') {
            try {
                $newT = Get-LiveTickets $Cfg
                $norm = Update-Normalizados $newT
                $tickets = $newT
                if ($idx -ge $tickets.Count) { $idx = 0 }
                $lastRefresh = [DateTime]::Now
                if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
            } catch { }
        }
        elseif ($ch -eq 'q' -or $ch -eq 'Q' -or $vk -eq 27) { [Console]::Clear(); return }
    }
}

# Submenu do Visualizador
function Show-VisualizadorMenu {
    param([hashtable]$Cfg, [string]$ExportScript)

    Write-Banner
    Write-Centered "-- VISUALIZAR CHAMADOS --" 'White'
    Write-Host ""
    Write-MenuOpt '1' 'Tempo Real   ' 'White'    'Busca do OTRS a cada 60s, ultimas 4 notas por chamado'
    Write-MenuOpt '2' 'Completo     ' 'White'    'Leitura do cache, todas as notas, scroll livre'
    Write-MenuOpt '0' 'Voltar       ' 'DarkGray' ''
    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Host ""
    Write-Host "  Opcao: " -ForegroundColor Cyan -NoNewline
    $ch = (Read-Host).Trim()
    switch ($ch) {
        '1' { Show-VisualizadorRealTime   $Cfg $ExportScript }
        '2' { Show-VisualizadorCompleto   $Cfg $ExportScript }
    }
}


# Nome antigo usado em copias do menu; evita erro se o switch ainda chamar Show-Visualizador.
function Show-Visualizador {
    param([hashtable]$Cfg, [string]$ExportScript)
    Show-VisualizadorMenu $Cfg $ExportScript
}


# =============================================================================
# Integracao com Hub (http://172.16.0.49:3210)
# =============================================================================
$script:HubSession = $null

function Hub-Login {
    param([string]$HubUrl, [string]$Email, [string]$Pass)
    $body = '{"email":"' + $Email + '","password":"' + $Pass + '"}'
    $resp = Invoke-WebRequest -Uri ($HubUrl + "/api/login") `
        -Method POST -ContentType "application/json" -Body $body `
        -SessionVariable 'WebSess' -UseBasicParsing
    $script:HubSession = $WebSess
    return $resp.Content | ConvertFrom-Json
}

function Hub-Get {
    param([string]$HubUrl, [string]$Path)
    $resp = Invoke-WebRequest -Uri ($HubUrl + $Path) `
        -Method GET -WebSession $script:HubSession -UseBasicParsing
    return $resp.Content | ConvertFrom-Json
}

function Hub-Post {
    param([string]$HubUrl, [string]$Path, $Body)
    $json = $Body | ConvertTo-Json -Depth 8
    $resp = Invoke-WebRequest -Uri ($HubUrl + $Path) `
        -Method POST -ContentType "application/json" -Body $json `
        -WebSession $script:HubSession -UseBasicParsing
    return $resp.Content | ConvertFrom-Json
}

function Hub-Put {
    param([string]$HubUrl, [string]$Path, $Body)
    $json = $Body | ConvertTo-Json -Depth 8
    $resp = Invoke-WebRequest -Uri ($HubUrl + $Path) `
        -Method PUT -ContentType "application/json" -Body $json `
        -WebSession $script:HubSession -UseBasicParsing
    return $resp.Content | ConvertFrom-Json
}

function Build-HubTicket {
    param($Ticket, [int]$MaxUpdates = 10)

    $openDate = "" ; $openHour = ""
    if ($Ticket.Criado -match '(\d{2})/(\d{2})/(\d{4}) (\d{2}):(\d{2})') {
        $openDate = $Matches[3] + "-" + $Matches[2] + "-" + $Matches[1]
        $openHour = $Matches[4] + ":" + $Matches[5]
    }

    $arts    = @($Ticket.Articles)
    $take    = [Math]::Min($MaxUpdates, $arts.Count)
    $updates = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $take; $i++) {
        $n = $arts[$i]
        $uDate = "" ; $uHour = ""
        if ($n.Date -match '(\d{2})/(\d{2})/(\d{4}) (\d{2}):(\d{2})') {
            $uDate = $Matches[3] + "-" + $Matches[2] + "-" + $Matches[1]
            $uHour = $Matches[4] + ":" + $Matches[5]
        }
        $updates.Add([ordered]@{ date=$uDate; hour=$uHour; text=$n.Text })
    }

    return [ordered]@{
        number      = $Ticket.Numero
        status      = $Ticket.Estado
        openingDate = $openDate
        openingHour = $openHour
        client      = $Ticket.Cliente
        updates     = @($updates)
    }
}

function Show-HubTicketPreview {
    param($Payload)
    Write-Host "  Ticket   : #" -ForegroundColor DarkGray -NoNewline
    Write-Host $Payload.number -ForegroundColor Yellow
    Write-Host "  Cliente  : " -ForegroundColor DarkGray -NoNewline
    Write-Host $Payload.client -ForegroundColor White
    Write-Host "  Estado   : " -ForegroundColor DarkGray -NoNewline
    Write-Host $Payload.status -ForegroundColor Yellow
    Write-Host "  Abertura : " -ForegroundColor DarkGray -NoNewline
    Write-Host ($Payload.openingDate + " " + $Payload.openingHour) -ForegroundColor White
    Write-Host "  Atualizac: " -ForegroundColor DarkGray -NoNewline
    Write-Host ($Payload.updates.Count.ToString() + " notas") -ForegroundColor White
}

function Invoke-SyncHub {
    param([hashtable]$Cfg, [string]$ExportScript)

    Write-Banner
    Write-Centered "-- SINCRONIZAR COM HUB --" 'White'
    Write-Host ""

    # Credenciais do Hub
    $hubUrl   = Read-Field "URL do Hub" "http://172.16.0.49:3210"
    $hubEmail = Read-Field "Email Hub"  ""
    Write-Host "  Senha Hub: " -ForegroundColor Yellow -NoNewline
    $hubPass = Read-Host
    Write-Host ""

    # Login
    Write-Info "Conectando ao Hub..."
    try {
        $loginResp = Hub-Login $hubUrl $hubEmail $hubPass
        Write-OK "Login realizado no Hub."
    } catch {
        Write-Err ("Falha no login: " + $_); Pause-Screen; return
    }

    # Tickets existentes no Hub
    Write-Info "Buscando tickets no Hub..."
    $existingMap = @{}
    try {
        $hubTickets = Hub-Get $hubUrl "/api/relatorio"
        if ($hubTickets -is [array]) {
            foreach ($ht in $hubTickets) {
                if ($ht.number) { $existingMap[$ht.number.ToString()] = $ht }
            }
        }
        Write-OK ($existingMap.Count.ToString() + " tickets encontrados no Hub.")
    } catch {
        Write-Warn ("Nao foi possivel buscar tickets existentes: " + $_)
    }

    # Tickets do OTRS (cache)
    Write-Info "Carregando cache OTRS..."
    $cachePath = $Cfg.EstadoFile
    if (-not [System.IO.Path]::IsPathRooted($cachePath)) {
        $cachePath = Join-Path $Cfg.OutputPath $cachePath
    }
    if (-not (Test-Path $cachePath)) {
        Write-Warn "Cache nao encontrado. Execute Gerar Relatorio primeiro."
        Pause-Screen; return
    }
    try {
        $otrsTickets = Load-TicketsFromCache $cachePath $ExportScript
        Write-OK ($otrsTickets.Count.ToString() + " tickets no cache OTRS.")
    } catch {
        Write-Err ("Falha ao carregar cache: " + $_); Pause-Screen; return
    }

    Write-Host ""
    Write-ThinDiv 'DarkGray'

    $nNew = 0 ; $nUpd = 0 ; $nSkip = 0

    foreach ($ticket in $otrsTickets) {
        $num = $ticket.Numero
        if (-not $num) { continue }

        $payload = Build-HubTicket $ticket

        if ($existingMap.ContainsKey($num)) {
            # Ticket ja existe - verifica se ha novas notas
            $hub = $existingMap[$num]
            if ($hub.updates) { $hubNoteCount = $hub.updates.Count } else { $hubNoteCount = 0 }
            $newNoteCount = $payload.updates.Count

            if ($newNoteCount -gt $hubNoteCount) {
                Write-Host ""
                Write-Host ("  [UPD] #" + $num + " - " + ($newNoteCount - $hubNoteCount) + " novas notas") -ForegroundColor Cyan
                Show-HubTicketPreview $payload
                Write-Host "  Atualizar? [S/n]: " -ForegroundColor Yellow -NoNewline
                $r = Read-Host
                if ($r -eq '' -or $r -match '^[Ss]') {
                    try {
                        Hub-Put $hubUrl ("/api/relatorio/" + $num) $payload | Out-Null
                        Write-OK ("Atualizado: #" + $num)
                        $nUpd++
                    } catch {
                        Write-Warn ("Falha: " + $_)
                    }
                } else {
                    Write-Info ("Pulado: #" + $num) ; $nSkip++
                }
            } else {
                Write-Host ("  [=] #" + $num + " sem alteracoes") -ForegroundColor DarkGray
                $nSkip++
            }
        } else {
            # Ticket novo - pede confirmacao
            Write-Host ""
            Write-ThinDiv 'Yellow'
            Write-Host "  NOVO TICKET:" -ForegroundColor Yellow
            Show-HubTicketPreview $payload
            Write-Host "  Adicionar ao Hub? [S/n]: " -ForegroundColor Yellow -NoNewline
            $r = Read-Host
            if ($r -eq '' -or $r -match '^[Ss]') {
                try {
                    Hub-Post $hubUrl "/api/relatorio" $payload | Out-Null
                    Write-OK ("Adicionado: #" + $num)
                    $nNew++
                } catch {
                    Write-Warn ("Falha: " + $_)
                }
            } else {
                Write-Info ("Pulado: #" + $num) ; $nSkip++
            }
        }
    }

    Write-Host ""
    Write-Divider 'DarkGray'
    Write-OK ($nNew.ToString() + " adicionados   " + $nUpd.ToString() + " atualizados   " + $nSkip.ToString() + " sem alteracao")
    Pause-Screen
}

# =============================================================================
# Configuracoes centralizadas
# =============================================================================
$now = Get-Date
$standardHour = [math]::Floor([int]$now.Hour / 2) * 2
$standardHourStr = $standardHour.ToString('00')
$script:CcoConfig = @{
    TituloRelatorio  = "Chamados Cr" + [char]237 + "ticos - " + $now.ToString('dd/MM/yyyy') + " " + $standardHourStr + "h"
    TituloResolvidos = "Chamados Normalizados - " + $now.ToString('dd/MM/yyyy') + " " + $standardHourStr + "h"
    LabelHorario     = "Hor" + [char]225 + "rio de abertura"
    LabelOcorrencia  = "Ocorr" + [char]234 + "ncia"
    Separador        = "*--------------------------------------------------------------*"
    EstadosResolvidos = @('resolvido','fechado','removido','encerrado','resolved','closed','merged','agrupado')
    AutoNotePatterns = @(
        'Sistema de Monitoramento',
        'Host:.*IP:.*Incidente:',
        'Prioridade:\s*(Disaster|High|Average|Information)',
        'Agradecemos seu contato',
        'solicita.{1,10}ser.{1,5}tratada atrav',
        'sob os cuidados da nossa equipe',
        'Para mais informa.{1,10}contate',
        'Para proteger sua privacidade',
        'conte.do remoto foi desabilitado',
        'Carregar conte.do bloqueado',
        'Chamado enviada para campo',
        'atingiu.{1,20}(HORAS|HORA|DUAS)',
        'Para seu conhecimento.*O Ticket',
        'Notifica.{1,5}o de Eventos',
        'Por favor, verificar o Incidente conforme dados',
        'Raz.o Social:.*CNPJ:.*Dificuldade',
        'Alerta - \[Ticket#',
        'Direcionamento de Ticket',
        'ticket.*foi direcionado',
        'COMUNICADO.*chave de e.mail',
        'plantao.*fora do hor',
        'Problema com Status OK',
        'Problema persiste',
        'Gr.ficos personalizados',
        'Estado do link Zabbix',
        'Status:\s*(OK|PROBLEM|DISASTER)'
    )
}

# =============================================================================
# Configuracoes
# =============================================================================

function Show-Configuracoes {
    param([hashtable]$Cfg, [string]$ConfigFilePath)
    Write-Banner
    Write-Centered "-- CONFIGURACOES --" 'White'
    Write-Host ""
    Write-Info "Enter = manter valor atual."
    Write-Host ""
    $Cfg.SearchPath = Read-Field "Perfil de busca (SearchPath)" $Cfg.SearchPath
    $Cfg.EstadoFile = Read-Field "Arquivo de cache (JSON)"      $Cfg.EstadoFile
    $Cfg.OutputPath = Read-Field "Pasta de saida"               $Cfg.OutputPath
    Write-Host ""
    Write-Host ("  Salvar em " + $ConfigFilePath + "? [S/n]: ") -ForegroundColor Yellow -NoNewline
    $resp = Read-Host
    if ($resp -eq '' -or $resp -match '^[Ss]') {
        Save-Config $Cfg $ConfigFilePath
        Write-OK ("Salvo em " + $ConfigFilePath)
    } else {
        Write-Warn "Nao salvo (apenas em memoria)."
    }
    Pause-Screen
    return $Cfg
}

function Show-SalvarCredenciais {
    param([hashtable]$Cfg, [string]$ConfigFilePath)
    Write-Banner
    Write-Centered "-- SALVAR CREDENCIAIS --" 'White'
    Write-Host ""
    Write-Warn ("Senha armazenada em TEXTO CLARO em: " + $ConfigFilePath)
    Write-Host ""
    Write-Host "  Confirma? [s/N]: " -ForegroundColor Yellow -NoNewline
    $resp = Read-Host
    if ($resp -match '^[Ss]') {
        Save-Config $Cfg $ConfigFilePath
        Write-OK ("Credenciais salvas em " + $ConfigFilePath)
    } else {
        Write-Info "Cancelado."
    }
    Pause-Screen
}

# =============================================================================
# Menu principal
# =============================================================================

function Show-MainMenu {
    param([hashtable]$Cfg, [string]$ExportScript, [string]$ConfigFilePath)

    while ($true) {
        $cfgStatus   = Get-CfgStatus $ConfigFilePath
        $hostDisplay = try { ([uri]$Cfg.BaseURL).Host } catch { $Cfg.BaseURL }

        Write-Banner
        Write-StatusBar $Cfg.Username $hostDisplay $cfgStatus
        Write-Centered "-- MENU PRINCIPAL --" 'White'
        Write-Host ""

        Write-MenuOpt '1' 'Gerar Relatorio TXT'     'White'    'Cria arquivo .txt no formato CCO'
        Write-MenuOpt '2' 'Gerar Relatorio JSON'    'White'    'Cria arquivo .json com todos os dados'
        Write-MenuOpt '3' 'Visualizar Chamados'     'Yellow'   'Navega pelos chamados no terminal (<- ->)'
        Write-Host ""
        Write-MenuOpt '4' 'Alterar Credenciais'     'Cyan'     'Usuario, senha e URL'
        Write-MenuOpt '5' 'Configuracoes'           'Cyan'     'Perfil de busca, cache, pasta de saida'
        Write-MenuOpt '6' 'Salvar Credenciais'      'DarkGray' ''
        Write-Host ""
        Write-MenuOpt '0' 'Sair'                    'DarkGray' ''
        Write-Host ""
        Write-ThinDiv 'DarkGray'
        Write-Host ""
        Write-Host "  Opcao: " -ForegroundColor Cyan -NoNewline
        $choice = Read-Host

        switch ($choice.Trim()) {
            '1' { Invoke-GerarTxt  $Cfg $ExportScript }
            '2' { Invoke-GerarJson $Cfg $ExportScript }
            '3' { Show-VisualizadorMenu $Cfg $ExportScript }
            '4' { $Cfg = Show-LoginScreen $Cfg }
            '5' { $Cfg = Show-Configuracoes $Cfg $ConfigFilePath }
            '6' { Show-SalvarCredenciais $Cfg $ConfigFilePath }
            '0' {
                Write-Banner
                Write-Centered "Ate logo!" 'DarkCyan'
                Write-Host ""
                return
            }
            default {
                Write-Warn "Opcao invalida."
                Start-Sleep -Milliseconds 600
            }
        }
    }
}

# =============================================================================
# Ponto de entrada
# =============================================================================

$cfg             = Load-Config $ConfigFile
$exportScriptPath = Resolve-ExportScript $ScriptPath

if (-not $cfg.Username -or -not $cfg.Password) {
    $cfg = Show-LoginScreen $cfg
} else {
    Write-Banner
    Write-Centered "-- BEM-VINDO --" 'White'
    Write-Host ""
    Write-OK  ("Credenciais carregadas de: " + $ConfigFile)
    Write-Info ("Usuario : " + $cfg.Username)
    Write-Info ("Servidor: " + $cfg.BaseURL)
    Write-Host ""
    Write-Host "  Usar estas credenciais? [S/n]: " -ForegroundColor Yellow -NoNewline
    $r = Read-Host
    if ($r -match '^[Nn]') { $cfg = Show-LoginScreen $cfg }
}

# Export-CcoReport embutido

Show-MainMenu $cfg $exportScriptPath $ConfigFile