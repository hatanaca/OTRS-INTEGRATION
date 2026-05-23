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
        BaseURL         = 'http://172.16.0.12/znuny'
        Username        = ''
        Password        = ''
        EstadoFile      = 'estado_chamados.json'
        OutputPath      = $PWD.Path
        SearchPath      = 'index.pl?Action=AgentKPISearch;Subaction=Search;TakeLastSearch=1;Profile=94_8'
        HubBaseURL      = 'http://172.16.0.49:3210'
        # Credenciais Hub - preenchimento automatico na sincronizacao (texto claro: nao commitar em repo publico).
        HubEmail        = 'thiago.ratanaka@microset.net.br'
        HubPassword     = 'SenhaN2M7@'
        SleepArticleMs  = 5
        SleepTicketMs   = 15
    }
    if (Test-Path $Path) {
        try {
            $j = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($j.BaseURL)     { $cfg.BaseURL     = $j.BaseURL     }
            if ($j.Username)    { $cfg.Username    = $j.Username    }
            if ($j.Password)    { $cfg.Password    = $j.Password    }
            if ($j.EstadoFile)  { $cfg.EstadoFile  = $j.EstadoFile  }
            if ($j.OutputPath)  { $cfg.OutputPath  = $j.OutputPath  }
            if ($j.SearchPath)  { $cfg.SearchPath  = $j.SearchPath  }
            if ($j.HubBaseURL)  { $cfg.HubBaseURL  = $j.HubBaseURL  }
            if ($j.HubEmail)    { $cfg.HubEmail    = $j.HubEmail    }
            if ($j.HubPassword) { $cfg.HubPassword = $j.HubPassword }
            if ($null -ne $j.SleepArticleMs) { $cfg.SleepArticleMs = [int]$j.SleepArticleMs }
            if ($null -ne $j.SleepTicketMs)  { $cfg.SleepTicketMs  = [int]$j.SleepTicketMs  }
        } catch { }
    }
    return $cfg
}

function Save-Config {
    param([hashtable]$Cfg, [string]$Path)
    [ordered]@{
        BaseURL         = $Cfg.BaseURL
        Username        = $Cfg.Username
        Password        = $Cfg.Password
        EstadoFile      = $Cfg.EstadoFile
        OutputPath      = $Cfg.OutputPath
        SearchPath      = $Cfg.SearchPath
        HubBaseURL      = $Cfg.HubBaseURL
        HubEmail        = $(if ($null -ne $Cfg.HubEmail) { $Cfg.HubEmail } else { '' })
        HubPassword     = $(if ($null -ne $Cfg.HubPassword) { $Cfg.HubPassword } else { '' })
        SleepArticleMs  = $(if ($null -ne $Cfg.SleepArticleMs) { [int]$Cfg.SleepArticleMs } else { 5 })
        SleepTicketMs   = $(if ($null -ne $Cfg.SleepTicketMs)  { [int]$Cfg.SleepTicketMs  } else { 15 })
    } | ConvertTo-Json -Compress | Out-File $Path -Encoding UTF8
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
    [int]$TimeoutSec

    OtrsClient([string]$baseUrl, [int]$retryMax = 3) {
        $this.BaseURL    = $baseUrl.TrimEnd('/')
        $this.RetryMax   = $retryMax
        $this.TimeoutSec = 120
    }

    [void] Login([string]$user, [string]$pass) {
        $loginUrl = "$($this.BaseURL)/index.pl"
        $body = @{ Action='Login'; RequestedURL=''; Lang='pt_BR'; TimeOffset='-180'; User=$user; Password=$pass }
        try {
            $tempSession = $null
            $response = Invoke-WebRequest -Uri $loginUrl -Method POST -UseBasicParsing -Body $body `
                -SessionVariable 'tempSession' -TimeoutSec $this.TimeoutSec -ErrorAction Stop
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
        Ensure-ZnunyParserRegexes
        $uri = "$($this.BaseURL)/$searchPath"
        $response = $this.InvokeWithRetry($uri)
        $content = $this.GetResponseText($response)
        $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($mm in $script:_RxTicketIdSearch.Matches($content)) {
            [void]$set.Add($mm.Groups[1].Value)
        }
        $arr = New-Object string[] $set.Count
        $set.CopyTo($arr)
        return $arr
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
                    -Headers $headers -WebSession $this.Session -UseBasicParsing -TimeoutSec $this.TimeoutSec -ErrorAction Stop
                return $this.GetResponseText($response)
            } catch {
                $attempt++
                if ($attempt -gt $this.RetryMax) { throw }
                $waitMs = [Math]::Min(1200, 50 * [int][Math]::Pow(2, $attempt - 1))
                Start-Sleep -Milliseconds $waitMs
                Write-Verbose "Retry $attempt/$($this.RetryMax) para widget CustomerInformation (ticket $ticketID)"
            }
        }
        return ''
    }

    [string] GetArticleRaw([string]$ticketID, [string]$articleID, [string]$subaction) {
        $uri = "$($this.BaseURL)/index.pl?Action=AgentTicketArticleContent;Subaction=$subaction;TicketID=$ticketID;ArticleID=$articleID;FileID=;"
        $response = $this.InvokeWithRetry($uri)
        return $this.GetResponseText($response)
    }

    [string] GetArticleContent([string]$ticketID, [string]$articleID) {
        return $this.GetArticleRaw($ticketID, $articleID, 'HTMLView')
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
                    TimeoutSec      = $this.TimeoutSec
                }
                if ($body) { $params.Body = $body }
                return Invoke-WebRequest @params
            } catch {
                $attempt++
                if ($attempt -gt $this.RetryMax) { throw }
                $waitMs = [Math]::Min(1200, 50 * [int][Math]::Pow(2, $attempt - 1))
                Start-Sleep -Milliseconds $waitMs
                Write-Verbose "Retry $attempt/$($this.RetryMax) para $uri"
            }
        }
        throw "Loop de retry excedido (nunca alcancado)"
    }

    hidden [string] GetResponseText($response) {
        if (-not $response) { return '' }
        # Preferir Content (string ja materializada) - mais rapido que ler RawContentStream inteiro.
        $c = $response.Content
        if ($null -ne $c -and $c.Length -gt 0) { return [string]$c }
        if (-not $response.RawContentStream) { return '' }
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
                    $nl = [System.Collections.Generic.List[object]]::new()
                    foreach ($n in $v.Notas) {
                        $nl.Add([PSCustomObject]@{ Date = $n.Date; Text = $n.Text })
                    }
                    $notas = $nl.ToArray()
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
        $obj | ConvertTo-Json -Depth 6 -Compress | Out-File -FilePath $this.FilePath -Encoding UTF8
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

function Ensure-ZnunyParserRegexes {
    if ($script:_ZnunyRxReady) { return }
    $ro = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $roIc = $ro -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $script:_RxTrArticleRows = [regex]::new('(?s)<tr[^>]*id="Row\d+"[^>]*>(.*?)</tr>', $ro)
    $script:_RxArtIdInRow = [regex]::new('<input[^>]+class="ArticleID"[^>]+value="(\d+)"', $roIc)
    $script:_RxArtDateInRow = [regex]::new('<td class="Created">[^<]*<div title="([^"]+)"', $ro)
    $script:_RxArtDateSortKey = [regex]::new('^(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2})', $ro)
    $script:_RxArtBody = [regex]::new('(?s)<body[^>]*>(.*)</body>', $roIc)
    $script:_RxStripStyle = [regex]::new('(?s)<style[^>]*>.*?</style>', $ro)
    $script:_RxStripScript = [regex]::new('(?s)<script[^>]*>.*?</script>', $ro)
    $script:_RxStripTags = [regex]::new('<[^>]+>', $ro)
    $script:_RxStripWs = [regex]::new('\s+', $ro)
    $script:_RxTicketIdSearch = [regex]::new('TicketID=(\d+)', $roIc)
    $script:_RxTicketNumeroHdr = [regex]::new('Ticket#([0-9]+)', $roIc)
    $script:_RxEstadoPill = [regex]::new('(?s)<span[^>]+class="[^"]*pill[^"]*"[^>]+title="([^"]+)"', $roIc)
    $script:_RxEstadoLabel = [regex]::new('(?s)label[^>]*>\s*Estado:\s*[^<]*</label>\s*<span[^>]+title="([^"]+)"', $roIc)
    $script:_RxCriadoTitle = [regex]::new('(?s)label[^>]*>\s*Criado:\s*</label>\s*<p[^>]+title="([^"]+)"', $roIc)
    $script:_RxLoginProbe = [regex]::new('Action=Login|name\s*=\s*"Password"|Login de Agente', $roIc)
    $script:_RxN2Assign = [regex]::new('Analista designado[:\s]*@([\w\s]+)', $roIc)
    $script:_RxFootEscreveu = [regex]::new('\s*\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s*-\s*\S+@\S+\s+escreveu:.*$', $ro)
    $script:_RxFootEmBlock = [regex]::new('\s*Em\s+\d{2}/\d{2}/\d{4}.*?escreveu:.*$', $ro)
    $script:_RxUnidadeAltUsuario = [regex]::new('(?s)<label[^>]*>\s*Usu.rio\s*:\s*</label>\s*<p[^>]+title="([^"]+)"', $roIc)
    $tokList = [System.Collections.Generic.List[regex]]::new()
    foreach ($tp in @(
            'name="ChallengeToken"\s+value="([A-Za-z0-9]+)"',
            'value="([A-Za-z0-9]+)"\s+name="ChallengeToken"',
            '"ChallengeToken",\s*"([A-Za-z0-9]+)"',
            'ChallengeToken=([A-Za-z0-9]{16,})',
            'data-challenge-token="([A-Za-z0-9]+)"',
            'ChallengeToken:\s*"([A-Za-z0-9]+)"'
        )) {
        try { $tokList.Add([regex]::new($tp, $roIc)) } catch { }
    }
    $script:_RxChallengeTokens = $tokList

    # Cliente / unidade (HTML principal e widget) - evita [regex]::Match repetido por ticket
    $script:_RxCustNomeTitle = [regex]::new('(?s)<label[^>]*>\s*Nome\s*:\s*</label>\s*<p[^>]+title="([^"]+)"', $roIc)
    $script:_RxCustNomeP = [regex]::new('(?s)<label[^>]*>\s*Nome\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>', $roIc)
    $script:_RxCustNomeSuffixTail = [regex]::new('\s*-\s*\d+\s*$', $ro)
    $script:_RxCustIdCliTitle = [regex]::new('(?s)<label[^>]*>\s*ID do Cliente\s*:\s*</label>\s*<p[^>]+title="([^"]+)"', $roIc)
    $script:_RxCustIdCliA = [regex]::new('(?s)<label[^>]*>\s*ID do Cliente\s*:\s*</label>\s*(?:<[^>]+>\s*)*<a[^>]*>([^<]+)</a>', $roIc)
    $script:_RxCustClienteTitle = [regex]::new('(?s)<label[^>]*>\s*Cliente\s*:\s*</label>\s*<p[^>]+title="([^"]+)"', $roIc)
    $script:_RxCustClienteP = [regex]::new('(?s)<label[^>]*>\s*Cliente\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>', $roIc)
    $script:_RxUnitUsuarioP = [regex]::new('(?s)<label[^>]*>\s*Usu.rio\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>', $roIc)
    $script:_RxHtmlTagProbe = [regex]::new('<[a-zA-Z!/]', $ro)
    $dynT = [System.Collections.Generic.List[regex]]::new()
    $dynP = [System.Collections.Generic.List[regex]]::new()
    foreach ($lb in @(
            'Unidade',
            'Loja',
            'Filial',
            'Campo\s*do\s*[Cc]liente',
            'Custom\s?[Uu]ser',
            'Customer\s?[Uu]ser'
        )) {
        $dynT.Add([regex]::new("(?s)<label[^>]*>\s*$lb\s*:\s*</label>\s*<p[^>]+title=`"([^`"]+)`"", $roIc))
        $dynP.Add([regex]::new("(?s)<label[^>]*>\s*$lb\s*:\s*</label>\s*(?:<[^>]+>\s*)*<p[^>]*>(.*?)</p>", $roIc))
    }
    $script:_RxUnitDynamicTitles = $dynT
    $script:_RxUnitDynamicPs = $dynP
    $script:_ZnunyRxReady = $true
}

function Get-ArticleBodyChunk {
    param([string]$Raw)
    if (-not $Raw) { return '' }
    Ensure-ZnunyParserRegexes
    if (($m = $script:_RxArtBody.Match($Raw)).Success) { return $m.Groups[1].Value.Trim() }
    if (-not $script:_RxHtmlTagProbe.IsMatch($Raw)) { return $Raw.Trim() }
    return ''
}

function Test-ArticleRawLooksLikeLoginPage {
    param([string]$Raw)
    if (-not $Raw) { return $true }
    Ensure-ZnunyParserRegexes
    return $script:_RxLoginProbe.IsMatch($Raw)
}

function Remove-Html {
    param([string]$Text)
    if (-not $Text) { return '' }
    Ensure-ZnunyParserRegexes
    try {
        $Text = $script:_RxStripStyle.Replace($Text, '')
        $Text = $script:_RxStripScript.Replace($Text, '')
        $Text = $script:_RxStripTags.Replace($Text, '')
        $Text = [System.Net.WebUtility]::HtmlDecode($Text)
    } catch {
        Write-Warning "Erro ao decodificar HTML: $_"
    }
    return $script:_RxStripWs.Replace($Text, ' ').Trim()
}

function Get-CustomerNameFromHtml {
    param([string]$HTML)
    Ensure-ZnunyParserRegexes

    if (($m = $script:_RxCustNomeTitle.Match($HTML)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val) {
            return $script:_RxCustNomeSuffixTail.Replace($val, '').Trim()
        }
    }
    if (($m = $script:_RxCustNomeP.Match($HTML)).Success) {
        $val = (Remove-Html $m.Groups[1].Value).Trim()
        if ($val) {
            return $script:_RxCustNomeSuffixTail.Replace($val, '').Trim()
        }
    }

    if (($m = $script:_RxCustIdCliTitle.Match($HTML)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val) { return $val }
    }
    if (($m = $script:_RxCustIdCliA.Match($HTML)).Success) {
        return $m.Groups[1].Value.Trim()
    }

    if (($m = $script:_RxCustClienteTitle.Match($HTML)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val) { return $val }
    }
    if (($m = $script:_RxCustClienteP.Match($HTML)).Success) {
        $val = (Remove-Html $m.Groups[1].Value).Trim()
        if ($val) { return $val }
    }
    return 'N/D'
}

# Texto util para "Campo do cliente" / unidade: evita codigo numerico puro (CodigoCliente) sem descricao.
function Test-UnidadeTextoUtil {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $t = $Text.Trim()
    if ($t.Length -lt 2) { return $false }
    if ($t -match '^\d+$') { return $false }
    if ($t -match '[\p{L}]') { return $true }
    if ($t -match '(?i)(loja|filial|unidade|matriz|agencia|predio|sala|andar)') { return $true }
    return $false
}

# Ex.: "ROYAL FIC - UNIDADE 50 - PORTO NACIONAL - 27163" -> "ROYAL FIC - UNIDADE 50 - PORTO NACIONAL" (remove codigo final)
function Get-UnidadeFromNomeClienteString {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($seg in ($Text.Trim() -split '\s*-\s*')) {
        $s = $seg.Trim()
        if ($s) { $parts.Add($s) }
    }
    while ($parts.Count -gt 1) {
        $last = $parts[$parts.Count - 1]
        if ($last -match '^\d+$') { $parts.RemoveAt($parts.Count - 1) }
        else { break }
    }
    if ($parts.Count -eq 0) { return '' }
    $out = [string]::Join(' - ', $parts).Trim()
    if (Test-UnidadeTextoUtil $out) { return $out }
    return ''
}

function Get-UnidadeFromNomeClienteHtml {
    param(
        [string]$HTML,
        [string]$ExcludeEqualTo = $null
    )
    Ensure-ZnunyParserRegexes
    $raw = ''
    if (($m = $script:_RxCustNomeTitle.Match($HTML)).Success) {
        $raw = $m.Groups[1].Value.Trim()
    }
    elseif (($m = $script:_RxCustNomeP.Match($HTML)).Success) {
        $raw = (Remove-Html $m.Groups[1].Value).Trim()
    }
    if (-not $raw) { return '' }
    $u = Get-UnidadeFromNomeClienteString $raw
    if (-not $u) { return '' }
    if ($ExcludeEqualTo -and ($u -eq $ExcludeEqualTo.Trim())) { return '' }
    return $u
}

function Try-MatchUnidadeDescriptive {
    param(
        [string]$HTML,
        [string]$ExcludeEqualTo = $null
    )
    if ([string]::IsNullOrWhiteSpace($HTML)) { return '' }
    Ensure-ZnunyParserRegexes
    for ($ui = 0; $ui -lt $script:_RxUnitDynamicTitles.Count; $ui++) {
        if (($m = $script:_RxUnitDynamicTitles[$ui].Match($HTML)).Success) {
            $val = $m.Groups[1].Value.Trim()
            if (-not (Test-UnidadeTextoUtil $val)) { continue }
            if ($ExcludeEqualTo -and ($val -eq $ExcludeEqualTo.Trim())) { continue }
            return $val
        }
        if (($m = $script:_RxUnitDynamicPs[$ui].Match($HTML)).Success) {
            $val = (Remove-Html $m.Groups[1].Value).Trim()
            if (-not (Test-UnidadeTextoUtil $val)) { continue }
            if ($ExcludeEqualTo -and ($val -eq $ExcludeEqualTo.Trim())) { continue }
            return $val
        }
    }
    $fromNome = Get-UnidadeFromNomeClienteHtml $HTML $ExcludeEqualTo
    if ($fromNome) { return $fromNome }
    if (($m = $script:_RxUnidadeAltUsuario.Match($HTML)).Success) {
        $val = $m.Groups[1].Value.Trim()
        if ($val -and $val -notmatch '@' -and (Test-UnidadeTextoUtil $val)) {
            if (-not $ExcludeEqualTo -or ($val -ne $ExcludeEqualTo.Trim())) { return $val }
        }
    }
    if (($m = $script:_RxUnitUsuarioP.Match($HTML)).Success) {
        $val = (Remove-Html $m.Groups[1].Value).Trim()
        if ($val -and $val -notmatch '@' -and (Test-UnidadeTextoUtil $val)) {
            if (-not $ExcludeEqualTo -or ($val -ne $ExcludeEqualTo.Trim())) { return $val }
        }
    }
    return ''
}

function Get-CustomerUnitFromHtml {
    param(
        [string]$HTML,
        [switch]$FallbackOnly
    )

    if ($FallbackOnly) { return 'N/D' }
    $u = Try-MatchUnidadeDescriptive $HTML
    if ($u) { return $u }
    return 'N/D'
}

function Test-AutoNote {
    param([string]$Text)
    if ($Text.Length -lt 5) { return $true }
    if ($null -eq $script:_AutoNoteRxList) {
        $lst = [System.Collections.Generic.List[System.Text.RegularExpressions.Regex]]::new()
        $opt = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
            [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
        foreach ($p in $script:CcoConfig.AutoNotePatterns) {
            try {
                $lst.Add([System.Text.RegularExpressions.Regex]::new($p, $opt))
            } catch {
                Write-Verbose "Pattern AutoNote invalido (ignorado): $p"
            }
        }
        $script:_AutoNoteRxList = $lst
    }
    foreach ($rx in $script:_AutoNoteRxList) {
        if ($rx.IsMatch($Text)) { return $true }
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

    Ensure-ZnunyParserRegexes
    if (($m = $script:_RxTicketNumeroHdr.Match($HTML)).Success) { $numero = $m.Groups[1].Value } else { $numero = 'N/D' }
    if (($m = $script:_RxEstadoPill.Match($HTML)).Success) { $estado = $m.Groups[1].Value }
    elseif (($m = $script:_RxEstadoLabel.Match($HTML)).Success) { $estado = $m.Groups[1].Value }
    else { $estado = 'N/D' }
    if (($m = $script:_RxCriadoTitle.Match($HTML)).Success) { $criado = $m.Groups[1].Value } else { $criado = 'N/D' }

    # Cliente/Unidade a partir do HTML principal (sem HTTP extra).
    $cliente = Get-CustomerNameFromHtml $HTML
    $unidade = Get-CustomerUnitFromHtml $HTML

    $widgetHtml = ''
    if (($cliente -eq 'N/D' -or $unidade -eq 'N/D') -and $client) {
        $token = ''
        foreach ($rxTok in $script:_RxChallengeTokens) {
            if (($m = $rxTok.Match($HTML)).Success) {
                $token = $m.Groups[1].Value
                Write-Verbose "ChallengeToken encontrado (ticket $ticketID): $token"
                break
            }
        }
        if (-not $token) {
            Write-Verbose "ChallengeToken NAO encontrado; widget nao buscado (ticket $ticketID)"
        } else {
            try {
                $widgetHtml = $client.GetCustomerWidgetHtml($ticketID, $token)
                if ($widgetHtml) {
                    Write-Verbose "Widget CustomerInformation obtido: $($widgetHtml.Length) chars (ticket $ticketID)"
                    if ($cliente -eq 'N/D') { $cliente = Get-CustomerNameFromHtml $widgetHtml }
                    if ($unidade -eq 'N/D') { $unidade = Get-CustomerUnitFromHtml $widgetHtml }
                } else {
                    Write-Verbose "Widget CustomerInformation retornou VAZIO para ticket $ticketID"
                }
            } catch {
                Write-Verbose "Widget CustomerInformation indisponivel para ticket $ticketID : $_"
            }
        }
    } else {
        Write-Verbose "Cliente e unidade ja no HTML principal; pulando widget AJAX (ticket $ticketID)"
    }

    Write-Verbose "Resultado extracao: Cliente=$cliente, Unidade=$unidade (ticket $ticketID)"

    # Ajuste: se unidade == cliente, algo saiu errado - tenta outro campo descritivo (ex.: Loja/Filial no widget)
    if ($unidade -ne 'N/D' -and $unidade -eq $cliente) {
        Write-Verbose "Unidade igual ao Cliente, tentando campo alternativo..."
        $searchSrc = $widgetHtml + $HTML
        $alt = Try-MatchUnidadeDescriptive $searchSrc -ExcludeEqualTo $cliente
        if ($alt) { $unidade = $alt }
        Write-Verbose "Apos ajuste: Cliente=$cliente, Unidade=$unidade"
    }

    $articles = [System.Collections.Generic.List[object]]::new()
    $rowMatches = $script:_RxTrArticleRows.Matches($HTML)

    # Linhas de artigo (sem baixar corpo ainda). Em modo limitado (ex.: 4 notas no live),
    # ordena por data decrescente para buscar primeiro as mais recentes e parar cedo.
    $rowList = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $rowMatches) {
        $rowHtml = $row.Groups[1].Value
        $mAid = $script:_RxArtIdInRow.Match($rowHtml)
        if (-not $mAid.Success) { continue }
        $aid = $mAid.Groups[1].Value
        $mDt = $script:_RxArtDateInRow.Match($rowHtml)
        if ($mDt.Success) { $adate = $mDt.Groups[1].Value } else { $adate = 'N/D' }
        $sortKey = [datetime]::MinValue
        if (($mSk = $script:_RxArtDateSortKey.Match($adate)).Success) {
            try {
                $sortKey = Get-Date -Year ([int]$mSk.Groups[3].Value) -Month ([int]$mSk.Groups[2].Value) -Day ([int]$mSk.Groups[1].Value) `
                    -Hour ([int]$mSk.Groups[4].Value) -Minute ([int]$mSk.Groups[5].Value) -Second 0
            } catch {}
        }
        $rowList.Add([ordered]@{ Row = $rowHtml; Aid = $aid; Adate = $adate; SortKey = $sortKey })
    }

    $unlimitedArticles = ($maxArticles -ge 9000)
    $unlimitedFetch = ($fetchLimit -ge 9000)
    $fetchCount = 0

    if (-not $unlimitedArticles -and $rowList.Count -gt 1) {
        $rowList.Sort([System.Comparison[object]]{
            param($a, $b)
            return ([datetime]$b.SortKey).CompareTo([datetime]$a.SortKey)
        })
    }

    $articlePlainPreferred = $null
    foreach ($item in $rowList) {
        if (-not $unlimitedArticles -and $articles.Count -ge $maxArticles) { break }
        if (-not $unlimitedFetch -and $fetchCount -ge $fetchLimit) { break }

        $aid   = $item.Aid
        $adate = $item.Adate
        try {
            if ($null -eq $articlePlainPreferred) {
                $httpThis = 0
                $rawTry = $null
                try {
                    $rawTry = $client.GetArticleRaw($ticketID, $aid, 'Plain')
                    $httpThis++
                } catch { }
                $probeChunk = Get-ArticleBodyChunk $rawTry
                if ($probeChunk.Length -ge 5 -and -not (Test-ArticleRawLooksLikeLoginPage $rawTry)) {
                    $articlePlainPreferred = $true
                    $raw = $rawTry
                } else {
                    $articlePlainPreferred = $false
                    $raw = $client.GetArticleRaw($ticketID, $aid, 'HTMLView')
                    $httpThis++
                }
                $fetchCount += $httpThis
            } elseif ($articlePlainPreferred) {
                $raw = $client.GetArticleRaw($ticketID, $aid, 'Plain')
                $fetchCount++
            } else {
                $raw = $client.GetArticleRaw($ticketID, $aid, 'HTMLView')
                $fetchCount++
            }

            $body = Get-ArticleBodyChunk $raw
            if (-not $body) { continue }
            $text = Remove-Html $body

            if (($am = $script:_RxN2Assign.Match($text)).Success) {
                $text = "Em tratativas com N2 @$($am.Groups[1].Value.Trim())"
            }
            $text = $script:_RxFootEscreveu.Replace($text, '')
            $text = $script:_RxFootEmBlock.Replace($text, '')
            $text = $script:_RxStripWs.Replace($text, ' ').Trim()

            if ($text.Length -ge 5 -and -not (Test-AutoNote $text)) {
                $articles.Add([PSCustomObject]@{ Date = $adate; Text = $text })
            }
            if ($sleepMs -gt 0) { Start-Sleep -Milliseconds $sleepMs }
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
        [int]$SleepArticleMs = 5,
        [int]$SleepTicketMs = 15,
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
        Ensure-ZnunyParserRegexes
        foreach ($rxTok in $script:_RxChallengeTokens) {
            if (($m = $rxTok.Match($diagMain)).Success) { $diagToken = $m.Groups[1].Value; break }
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
    $activeSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($id in $activeIds) { [void]$activeSet.Add($id) }

    $cache = [TicketCache]::new($EstadoFile)

    Write-Verbose "Ativos na busca: $($activeIds.Count)"
    Write-Verbose "Em cache: $($cache.Data.Count)"

    $allIds = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $activeIds) { $allIds.Add($id) }
    foreach ($id in $cache.Data.Keys) {
        if (-not $activeSet.Contains($id)) { $allIds.Add($id) }
    }
    if ($allIds.Count -gt 1) {
        $allIds.Sort([System.Comparison[string]]{
            param($a, $b)
            return ([int]$a).CompareTo([int]$b)
        })
    }

    Write-Verbose "Total a processar: $($allIds.Count)"

    $ativos    = [System.Collections.Generic.List[TicketData]]::new()
    $resolvidos = [System.Collections.Generic.List[TicketData]]::new()
    $cacheHits = 0
    $httpCount = 0

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $progN = 0
    $progStep = if ($allIds.Count -gt 120) { 15 } elseif ($allIds.Count -gt 60) { 10 } elseif ($allIds.Count -gt 25) { 7 } else { 5 }
    foreach ($tid in $allIds) {
        $progN++
        if (($progN % $progStep) -eq 0 -or $progN -eq 1) {
            $pct = [int]([Math]::Min(100, $progN * 100 / [Math]::Max(1, $allIds.Count)))
            Write-Progress -Activity "Processando tickets" -Status "Ticket $tid ($progN/$($allIds.Count))" -PercentComplete $pct
        }

        if ($cache.IsCachedAndResolved($tid) -and -not $activeSet.Contains($tid)) {
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
            if ($SleepTicketMs -gt 0) { Start-Sleep -Milliseconds $SleepTicketMs }
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
    $sa = 5; $st = 15
    if ($null -ne $Cfg.SleepArticleMs) { $sa = [int]$Cfg.SleepArticleMs }
    if ($null -ne $Cfg.SleepTicketMs)  { $st = [int]$Cfg.SleepTicketMs }
    $params = @{
        BaseURL        = $Cfg.BaseURL
        Username       = $Cfg.Username
        Password       = $Cfg.Password
        SearchPath     = $Cfg.SearchPath
        EstadoFile     = $Cfg.EstadoFile
        OutputDir      = $Cfg.OutputPath
        SleepArticleMs = $sa
        SleepTicketMs  = $st
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
    $sa = 5; $st = 15
    if ($null -ne $Cfg.SleepArticleMs) { $sa = [int]$Cfg.SleepArticleMs }
    if ($null -ne $Cfg.SleepTicketMs)  { $st = [int]$Cfg.SleepTicketMs }
    $params = @{
        BaseURL        = $Cfg.BaseURL
        Username       = $Cfg.Username
        Password       = $Cfg.Password
        SearchPath     = $Cfg.SearchPath
        EstadoFile     = $Cfg.EstadoFile
        OutputDir      = $Cfg.OutputPath
        SleepArticleMs = $sa
        SleepTicketMs  = $st
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
    param(
        [hashtable]$Cfg,
        # Se > 0, mantem apenas as N notas mais recentes (apos ordenacao cronologica no TicketData).
        # Se 0, retorna todas as notas obtidas do OTRS.
        [int]$RecentArticlesOnly = 4,
        # Quando informado, reutiliza a sessao ja autenticada (evita login/logout a cada refresh).
        [OtrsClient]$SessionClient = $null
    )
    $ownClient = $false
    $client = $SessionClient
    if ($null -eq $client) {
        $client = [OtrsClient]::new($Cfg.BaseURL, 8)
        $client.Login($Cfg.Username, $Cfg.Password)
        $ownClient = $true
    }
    try {
        $ids = $client.GetActiveTicketIDs($Cfg.SearchPath)
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($id in $ids) {
            try {
                $html   = $client.GetTicketHtml($id)
                $maxArt = 9999
                $fetchLim = 9999
                if ($RecentArticlesOnly -gt 0) {
                    $maxArt = $RecentArticlesOnly
                    $fetchLim = [Math]::Max(60, $RecentArticlesOnly * 25)
                }
                $ticket = Get-TicketDataFromHtml -HTML $html -ticketID $id -client $client `
                    -maxArticles $maxArt -fetchLimit $fetchLim -sleepMs 0
                if ($null -eq $ticket) { continue }
                if ($RecentArticlesOnly -gt 0 -and $ticket.Articles.Count -gt $RecentArticlesOnly) {
                    $short = [System.Collections.Generic.List[object]]::new()
                    $start = $ticket.Articles.Count - $RecentArticlesOnly
                    for ($i = $start; $i -lt $ticket.Articles.Count; $i++) {
                        $short.Add($ticket.Articles[$i])
                    }
                    $ticket = [TicketData]::new($ticket.Numero, $ticket.Estado, $ticket.Criado, $ticket.Cliente, $ticket.Unidade, $short)
                }
                $list.Add($ticket)
            } catch { }
        }
        return @($list)
    } finally {
        if ($ownClient) { $client.Logout() }
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
    Ensure-ZnunyParserRegexes

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
            $txt = $script:_RxStripWs.Replace($nota.Text, ' ').Trim()
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
    $kc = $cache.Data.Keys.Count
    $keyArr = New-Object string[] $kc
    $ki = 0
    foreach ($k in $cache.Data.Keys) {
        $keyArr[$ki++] = [string]$k
    }
    if ($keyArr.Length -gt 1) {
        [Array]::Sort($keyArr, [System.Comparison[string]]{
            param($a, $b)
            return ([int]$a).CompareTo([int]$b)
        })
    }
    foreach ($id in $keyArr) {
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

# Modo tempo real: busca do OTRS a cada 60s.
# -TodasNotas: todas as notas de cada chamado (MaxNotes 0 = rolagem por altura).
# Caso contrario: apenas as 4 notas mais recentes por chamado.
function Show-VisualizadorRealTime {
    param([hashtable]$Cfg, [string]$ExportScript, [switch]$TodasNotas)

    $REFRESH_SEC = 60
    $recentParam = if ($TodasNotas) { 0 } else { 4 }
    $cardMaxNotes = if ($TodasNotas) { 0 } else { 4 }

    Write-Banner
    if ($TodasNotas) {
        Write-Centered "-- OTRS TEMPO REAL (TODAS AS NOTAS) --" 'White'
    } else {
        Write-Centered "-- OTRS TEMPO REAL (4 ULTIMAS NOTAS) --" 'White'
    }
    Write-Host ""
    Write-Info ("Servidor: " + $Cfg.BaseURL)
    Write-Info ("Usuario : " + $Cfg.Username)
    if ($TodasNotas) {
        Write-Warn "Modo completo: cada atualizacao baixa todas as notas (pode ser lento)."
    }
    Write-Info "Sessao unica no OTRS: um login ao entrar e logout ao sair com [Q] (menos avisos de excesso de logins)."
    Write-Info "Conectando e buscando chamados ativos..."
    Write-Host ""

    $otrsClient = [OtrsClient]::new($Cfg.BaseURL, 8)
    try {
        Ensure-ExportScript $ExportScript
        $otrsClient.Login($Cfg.Username, $Cfg.Password)
        $tickets = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam -SessionClient $otrsClient
    } catch {
        Write-Err ("Falha ao buscar: " + $_)
        try { $otrsClient.Logout() } catch {}
        Pause-Screen; return
    }

    if ($tickets.Count -eq 0) {
        Write-Warn "Nenhum chamado ativo encontrado."
        try { $otrsClient.Logout() } catch {}
        Pause-Screen; return
    }

    $null = Update-Normalizados $tickets

    $idx         = 0
    $lastRefresh = [DateTime]::Now

    try {
        while ($true) {
            $secsLeft = $REFRESH_SEC - [int]([DateTime]::Now - $lastRefresh).TotalSeconds
            if ($secsLeft -lt 0) { $secsLeft = 0 }

            Show-TicketCard $tickets[$idx] $idx $tickets.Count 0 $cardMaxNotes $secsLeft

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
                    $newT = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam -SessionClient $otrsClient
                    $norm = Update-Normalizados $newT
                    $tickets = $newT
                    if ($idx -ge $tickets.Count) { $idx = 0 }
                    $lastRefresh = [DateTime]::Now
                    if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
                } catch {
                    try {
                        Write-Warn ("Atualizacao automatica falhou; tentando novo login: " + $_)
                        $otrsClient.Login($Cfg.Username, $Cfg.Password)
                        $newT = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam -SessionClient $otrsClient
                        $norm = Update-Normalizados $newT
                        $tickets = $newT
                        if ($idx -ge $tickets.Count) { $idx = 0 }
                        $lastRefresh = [DateTime]::Now
                        if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
                    } catch {
                        Write-Warn ("Ainda falhou: " + $_)
                    }
                }
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
                    $newT = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam -SessionClient $otrsClient
                    $norm = Update-Normalizados $newT
                    $tickets = $newT
                    if ($idx -ge $tickets.Count) { $idx = 0 }
                    $lastRefresh = [DateTime]::Now
                    if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
                } catch {
                    try {
                        Write-Warn ("Refresh manual falhou; tentando novo login: " + $_)
                        $otrsClient.Login($Cfg.Username, $Cfg.Password)
                        $newT = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam -SessionClient $otrsClient
                        $norm = Update-Normalizados $newT
                        $tickets = $newT
                        if ($idx -ge $tickets.Count) { $idx = 0 }
                        $lastRefresh = [DateTime]::Now
                        if ($norm.Count -gt 0) { Show-NormalizacaoAlert $norm }
                    } catch {
                        Write-Warn ("Ainda falhou: " + $_)
                    }
                }
            }
            elseif ($ch -eq 'q' -or $ch -eq 'Q' -or $vk -eq 27) { [Console]::Clear(); return }
        }
    } finally {
        try { $otrsClient.Logout() } catch {}
    }
}

# Submenu do Visualizador
function Show-VisualizadorMenu {
    param([hashtable]$Cfg, [string]$ExportScript)

    Write-Banner
    Write-Centered "-- VISUALIZAR CHAMADOS --" 'White'
    Write-Host ""
    Write-MenuOpt '1' 'OTRS tempo real (4 notas) ' 'White'    'Sessao unica; atualiza a cada 60s; 4 notas mais recentes por chamado'
    Write-MenuOpt '2' 'OTRS tempo real (todas)   ' 'White'    'Sessao unica; atualiza a cada 60s; todas as notas (mais lento)'
    Write-MenuOpt '3' 'Cache local (offline)     ' 'Yellow'   'Ultimo relatorio salvo em JSON; todas as notas, sem consultar OTRS'
    Write-MenuOpt '0' 'Voltar                    ' 'DarkGray' ''
    Write-Host ""
    Write-Info "Em qualquer modo OTRS: alerta em tela cheia se um chamado passar a estado normalizado."
    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Host ""
    Write-Host "  Opcao: " -ForegroundColor Cyan -NoNewline
    $ch = (Read-Host).Trim()
    switch ($ch) {
        '1' { Show-VisualizadorRealTime   $Cfg $ExportScript }
        '2' { Show-VisualizadorRealTime   $Cfg $ExportScript -TodasNotas }
        '3' { Show-VisualizadorCompleto   $Cfg $ExportScript }
    }
}


# =============================================================================
# Integracao com Hub (http://172.16.0.49:3210) - API externa: validacao e timeouts
# =============================================================================
$script:HubSession = $null
$script:HubRequestTimeoutSec = 90

function Hub-InvokeWebRequest {
    param(
        [string]$Uri,
        [string]$Method = 'GET',
        [object]$WebSession = $null,
        [string]$ContentType = $null,
        [string]$Body = $null,
        [int]$TimeoutSec = 0,
        [string]$SessionVariable = $null
    )
    if ($TimeoutSec -le 0) { $TimeoutSec = $script:HubRequestTimeoutSec }
    $p = @{
        Uri             = $Uri
        Method          = $Method
        UseBasicParsing = $true
        TimeoutSec      = $TimeoutSec
        ErrorAction     = 'Stop'
    }
    if ($SessionVariable) {
        $p.SessionVariable = $SessionVariable
    } elseif ($WebSession) {
        $p.WebSession = $WebSession
    }
    if ($ContentType) { $p.ContentType = $ContentType }
    if ($null -ne $Body) { $p.Body = $Body }
    $resp = Invoke-WebRequest @p
    if ($SessionVariable) {
        $script:HubSession = Get-Variable -Name $SessionVariable -Scope Local -ValueOnly -ErrorAction Stop
        Remove-Variable -Name $SessionVariable -Scope Local -ErrorAction SilentlyContinue
    }
    return $resp
}

function Normalize-HubUpdatesFromAny {
    param($updates)
    $list = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $updates) { return $list }
    foreach ($u in @($updates)) {
        if ($null -eq $u) { continue }
        $d = ''
        if ($null -ne $u.date) { $d = [string]$u.date }
        elseif ($null -ne $u.updateDate) { $d = [string]$u.updateDate }
        $h = ''
        if ($null -ne $u.hour) { $h = [string]$u.hour }
        elseif ($null -ne $u.updateHour) { $h = [string]$u.updateHour }
        $tx = if ($null -ne $u.text) { [string]$u.text } else { '' }
        $list.Add([ordered]@{ date = $d.Trim(); hour = $h.Trim(); text = $tx })
    }
    return $list
}

function Normalize-HubTicketObject {
    param($ht)
    if ($null -eq $ht) { return $null }
    $num = $ht.number
    if ($null -eq $num) { return $null }
    $nu = ([string]$num).Trim()
    if ($nu.Length -eq 0) { return $null }
    return [ordered]@{
        number      = $nu
        status      = if ($null -ne $ht.status) { [string]$ht.status } else { '' }
        openingDate = if ($null -ne $ht.openingDate) { [string]$ht.openingDate } else { '' }
        openingHour = if ($null -ne $ht.openingHour) { [string]$ht.openingHour } else { '' }
        client      = if ($null -ne $ht.client) { [string]$ht.client } else { '' }
        updates     = @((Normalize-HubUpdatesFromAny $ht.updates))
    }
}

function Parse-HubRelatorioResponseToTickets {
    param($obj, [System.Collections.Generic.List[string]]$warnOut)
    $out = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $obj) {
        [void]$warnOut.Add('Resposta JSON nula do Hub.')
        return $out
    }
    if ($obj -is [System.Array]) {
        foreach ($x in $obj) {
            $n = Normalize-HubTicketObject $x
            if ($n) { $out.Add($n) }
            elseif ($null -ne $x) { [void]$warnOut.Add('Item na lista ignorado: sem campo number valido.') }
        }
        return $out
    }
    $hasSuccess = $obj.PSObject.Properties.Match('success')
    if ($hasSuccess.Count -gt 0 -and $obj.success -eq $false) {
        [void]$warnOut.Add('Hub retornou success=false; lista de tickets nao sera usada.')
        return $out
    }
    if ($null -ne $obj.tickets) {
        foreach ($x in @($obj.tickets)) {
            $n = Normalize-HubTicketObject $x
            if ($n) { $out.Add($n) }
            else { [void]$warnOut.Add('Ticket ignorado no array tickets: number invalido ou ausente.') }
        }
        return $out
    }
    $single = Normalize-HubTicketObject $obj
    if ($single) { $out.Add($single) }
    return $out
}

function Get-HubRelatorioTicketsList {
    param([string]$HubUrl)
    $warn = [System.Collections.Generic.List[string]]::new()
    $base = $HubUrl.TrimEnd('/')
    $tries = @(
        @{ Path = '/api/relatorio/tickets'; Label = 'GET /api/relatorio/tickets' },
        @{ Path = '/api/relatorio'; Label = 'GET /api/relatorio (legado)' }
    )
    foreach ($tr in $tries) {
        try {
            $uri = $base + $tr.Path
            $resp = Hub-InvokeWebRequest -Uri $uri -Method GET -WebSession $script:HubSession
            $raw = $resp.Content
            if ([string]::IsNullOrWhiteSpace($raw)) {
                [void]$warn.Add($tr.Label + ': corpo vazio.')
                continue
            }
            $obj = $null
            try {
                $obj = $raw | ConvertFrom-Json -ErrorAction Stop
            } catch {
                [void]$warn.Add($tr.Label + ': JSON invalido - ' + $_.Exception.Message)
                continue
            }
            $list = Parse-HubRelatorioResponseToTickets $obj $warn
            $syncV = $null
            $syncAt = $null
            if ($obj -and $obj -isnot [System.Array]) {
                if ($obj.PSObject.Properties.Match('syncVersion').Count -gt 0) { $syncV = $obj.syncVersion }
                if ($obj.PSObject.Properties.Match('syncUpdatedAt').Count -gt 0) { $syncAt = [string]$obj.syncUpdatedAt }
            }
            return @{
                ok            = $true
                tickets       = @($list)
                syncVersion   = $syncV
                syncUpdatedAt = $syncAt
                source        = $tr.Label
                warnings      = $warn
            }
        } catch {
            [void]$warn.Add($tr.Label + ' falhou: ' + $_.Exception.Message)
        }
    }
    return @{ ok = $false; tickets = @(); syncVersion = $null; syncUpdatedAt = $null; source = ''; warnings = $warn }
}

function Hub-VerifySession {
    param([string]$HubUrl)
    $base = $HubUrl.TrimEnd('/')
    try {
        $resp = Hub-InvokeWebRequest -Uri ($base + '/api/status/me') -Method GET -WebSession $script:HubSession
        $j = $resp.Content | ConvertFrom-Json -ErrorAction Stop
        if ($j -and $j.email) {
            Write-OK ('Sessao Hub valida: ' + [string]$j.email)
            return $true
        }
        Write-Warn 'GET /api/status/me nao retornou email; sessao pode estar incompleta.'
        return $false
    } catch {
        Write-Warn ('Validacao de sessao ignorada (/api/status/me): ' + $_.Exception.Message)
        return $false
    }
}

function Hub-Login {
    param([string]$HubUrl, [string]$Email, [string]$Pass)
    $base = $HubUrl.TrimEnd('/')
    $loginBody = (@{ email = $Email; password = $Pass } | ConvertTo-Json -Compress)
    $sv = 'HubSess_' + ([Guid]::NewGuid().ToString('N').Substring(0, 12))
    $resp = Hub-InvokeWebRequest -Uri ($base + '/api/login') -Method POST `
        -ContentType 'application/json; charset=utf-8' -Body $loginBody `
        -SessionVariable $sv
    $raw = $resp.Content
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw | ConvertFrom-Json -ErrorAction Stop
}

function Hub-Get {
    param([string]$HubUrl, [string]$Path)
    $base = $HubUrl.TrimEnd('/')
    $resp = Hub-InvokeWebRequest -Uri ($base + $Path) -Method GET -WebSession $script:HubSession
    $raw = $resp.Content
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw | ConvertFrom-Json -ErrorAction Stop
}

function Hub-Post {
    param([string]$HubUrl, [string]$Path, $Body)
    $base = $HubUrl.TrimEnd('/')
    $json = $Body | ConvertTo-Json -Depth 8
    $resp = Hub-InvokeWebRequest -Uri ($base + $Path) -Method POST `
        -ContentType 'application/json; charset=utf-8' -Body $json -WebSession $script:HubSession
    $raw = $resp.Content
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw | ConvertFrom-Json -ErrorAction Stop
}

function Hub-Put {
    param([string]$HubUrl, [string]$Path, $Body)
    $base = $HubUrl.TrimEnd('/')
    $json = $Body | ConvertTo-Json -Depth 8
    $resp = Hub-InvokeWebRequest -Uri ($base + $Path) -Method PUT `
        -ContentType 'application/json; charset=utf-8' -Body $json -WebSession $script:HubSession
    $raw = $resp.Content
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw | ConvertFrom-Json -ErrorAction Stop
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
        # Contrato observado na API GET do Hub (updateDate / updateHour / text)
        $updates.Add([ordered]@{ updateDate = $uDate; updateHour = $uHour; text = $n.Text })
    }

    # Campos alinhados ao formulario do Hub (relatorio CCO). O campo "ocorrencia" nao e enviado
    # para o operador preencher manualmente no navegador.
    return [ordered]@{
        number      = $Ticket.Numero
        status      = $Ticket.Estado
        openingDate = $openDate
        openingHour = $openHour
        client      = $Ticket.Cliente
        updates     = @($updates)
    }
}

function Test-HubRelatorioChanged {
    param($hubTicket, $payload)
    if (-not $hubTicket) { return $true }
    $pn = Normalize-HubTicketObject $payload
    if (-not $pn) { return $true }
    if (($hubTicket.status -or '') -ne ($pn.status -or '')) { return $true }
    if (($hubTicket.openingDate -or '') -ne ($pn.openingDate -or '')) { return $true }
    if (($hubTicket.openingHour -or '') -ne ($pn.openingHour -or '')) { return $true }
    if (($hubTicket.client -or '') -ne ($pn.client -or '')) { return $true }
    $hu = @($hubTicket.updates)
    $pu = @($pn.updates)
    if ($hu.Count -ne $pu.Count) { return $true }
    for ($i = 0; $i -lt $hu.Count; $i++) {
        if (($hu[$i].text -or '') -ne ($pu[$i].text -or '')) { return $true }
        if (($hu[$i].date -or '') -ne ($pu[$i].date -or '')) { return $true }
        if (($hu[$i].hour -or '') -ne ($pu[$i].hour -or '')) { return $true }
    }
    return $false
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
    Write-Info "O campo Ocorrencia do Hub nao e alterado por esta sincronizacao (preenchimento manual)."
    Write-Host ""

    # Credenciais do Hub
    $hubDefault = if ($Cfg.HubBaseURL) { $Cfg.HubBaseURL } else { 'http://172.16.0.49:3210' }
    $hubUrl   = Read-Field "URL do Hub" $hubDefault
    $defEmail = if ($Cfg.HubEmail) { [string]$Cfg.HubEmail } else { '' }
    $hubEmail = Read-Field "Email Hub" $defEmail
    Write-Host "  Senha Hub [Enter = usar senha salva no script/config]: " -ForegroundColor Yellow -NoNewline
    $hubPass = Read-Host
    if ([string]::IsNullOrWhiteSpace($hubPass)) {
        $hubPass = if ($Cfg.HubPassword) { [string]$Cfg.HubPassword } else { '' }
    }
    $hubUrl = $hubUrl.TrimEnd('/')
    Write-Host ""

    if ([string]::IsNullOrWhiteSpace($hubEmail) -or [string]::IsNullOrWhiteSpace($hubPass)) {
        Write-Err "Email ou senha Hub vazios. Ajuste em Configuracoes (opcao 5) ou nos padroes em Load-Config."
        Pause-Screen
        return
    }

    # Login
    Write-Info "Conectando ao Hub..."
    try {
        $null = Hub-Login $hubUrl $hubEmail $hubPass
        Write-OK "Login realizado no Hub."
    } catch {
        Write-Err ("Falha no login: " + $_); Pause-Screen; return
    }

    $null = Hub-VerifySession $hubUrl

    # Tickets existentes no Hub (GET /api/relatorio/tickets com fallback e validacao)
    Write-Info "Buscando tickets no Hub..."
    $existingMap = @{}
    $hubListResult = Get-HubRelatorioTicketsList $hubUrl
    foreach ($w in @($hubListResult.warnings)) {
        if ($w) { Write-Warn $w }
    }
    if (-not $hubListResult.ok) {
        Write-Warn "Nao foi possivel obter lista de tickets do Hub; sincronizacao tratara todos como novos (POST)."
    } else {
        Write-Info ("Fonte: " + $hubListResult.source)
        if ($null -ne $hubListResult.syncVersion) {
            Write-Info ("syncVersion: " + [string]$hubListResult.syncVersion + "  syncUpdatedAt: " + [string]$hubListResult.syncUpdatedAt)
        }
        foreach ($t in @($hubListResult.tickets)) {
            if ($t -and $t.number) { $existingMap[[string]$t.number] = $t }
        }
        Write-OK ($existingMap.Count.ToString() + " tickets encontrados no Hub (normalizados).")
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
            # Ticket ja existe - compara estado, cliente, abertura e todas as atualizacoes
            $hub = $existingMap[$num]
            if (Test-HubRelatorioChanged $hub $payload) {
                Write-Host ""
                Write-Host ("  [UPD] #" + $num + " - diferenca em relacao ao Hub (estado, cliente, datas ou notas)") -ForegroundColor Cyan
                Show-HubTicketPreview $payload
                Write-Host "  Atualizar registro no Hub? [S/n]: " -ForegroundColor Yellow -NoNewline
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
$script:_AutoNoteRxList = $null

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
    if (-not $Cfg.HubBaseURL) { $Cfg.HubBaseURL = 'http://172.16.0.49:3210' }
    $Cfg.HubBaseURL = Read-Field "URL base do Hub (relatorio CCO)" $Cfg.HubBaseURL
    if ($null -eq $Cfg.HubEmail) { $Cfg.HubEmail = '' }
    $Cfg.HubEmail = Read-Field "Email Hub (login relatorio)" $Cfg.HubEmail
    Write-Host "  Senha Hub [Enter = manter atual]: " -ForegroundColor Yellow -NoNewline
    $hubPwNew = Read-Host
    if ($hubPwNew) { $Cfg.HubPassword = $hubPwNew }
    elseif ($null -eq $Cfg.HubPassword) { $Cfg.HubPassword = '' }
    if ($null -eq $Cfg.SleepArticleMs) { $Cfg.SleepArticleMs = 5 }
    if ($null -eq $Cfg.SleepTicketMs)  { $Cfg.SleepTicketMs  = 15 }
    $Cfg.SleepArticleMs = [int](Read-Field "Pausa entre notas ao exportar (ms, 0=sem)" $Cfg.SleepArticleMs.ToString())
    $Cfg.SleepTicketMs  = [int](Read-Field "Pausa entre tickets ao exportar (ms, 0=sem)" $Cfg.SleepTicketMs.ToString())
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
# Diagnostico de rede (ping / tracert) - didatico; destino validado
# =============================================================================

function Test-IsWindowsHost {
    if ($PSVersionTable.PSVersion.Major -ge 6 -and $null -ne $PSVersionTable.Platform) {
        return $PSVersionTable.Platform -eq 'Win32NT'
    }
    return $env:OS -eq 'Windows_NT'
}

function Test-DestinoRedeSeguro {
    param([string]$Texto)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return $false }
    $t = $Texto.Trim()
    if ($t.Length -gt 253) { return $false }
    if ($t -match '[\s;|&`$<>''"\\]') { return $false }
    if ($t -notmatch '^[a-zA-Z0-9.\-:\[\]%_]+$') { return $false }
    return $true
}

function Read-DestinoRede {
    param([string]$TituloAjuda)
    Write-Host ""
    Write-Info $TituloAjuda
    Write-Host "  Exemplos validos:  google.com   8.8.8.8   znuny.suaempresa.local" -ForegroundColor DarkGray
    Write-Host "  (Use apenas letras, numeros, pontos, hifens e dois-pontos em IPv6.)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Destino: " -ForegroundColor Yellow -NoNewline
    $d = Read-Host
    return $(if ($null -eq $d) { '' } else { $d.Trim() })
}

function Show-PingDidatico {
    Write-Banner
    Write-Centered "-- TESTE DE PING --" 'White'
    Write-Host ""
    Write-Host "  O que e o ping?" -ForegroundColor Cyan
    Write-Host "  E um teste simples: seu computador envia alguns pacotes pequenos ate o"
    Write-Host "  destino e espera resposta. Se houver resposta, em geral o caminho ate la"
    Write-Host "  esta funcionando. O tempo em milissegundos (ms) indica atraso (latencia)."
    Write-Host ""
    Write-Host "  O que NAO e o ping?" -ForegroundColor DarkGray
    Write-Host "  Nao prova que um site ou servico (porta) esta no ar - so testa alcance"
    Write-Host "  basico na rede. Alguns firewalls bloqueiam ping; falha nem sempre e problema seu."
    Write-Host ""

    $dest = Read-DestinoRede "Digite o host ou IP que deseja testar (somente o destino)."
    if (-not (Test-DestinoRedeSeguro $dest)) {
        Write-Err "Destino invalido ou vazio. Use apenas nome ou IP, sem espacos ou caracteres especiais."
        Pause-Screen
        return
    }

    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Info ("Executando ping (4 tentativas) para: " + $dest)
    Write-ThinDiv 'DarkGray'
    Write-Host ""

    try {
        if (Test-IsWindowsHost) {
            $pingExe = Join-Path $env:SystemRoot 'System32\ping.exe'
            if (-not (Test-Path -LiteralPath $pingExe)) { throw "ping.exe nao encontrado." }
            $proc = Start-Process -FilePath $pingExe -ArgumentList @('-n', '4', $dest) -NoNewWindow -Wait -PassThru
            if ($proc.ExitCode -ne 0) {
                Write-Warn ("Ping finalizou com codigo " + $proc.ExitCode.ToString() + " (nem sempre indica falha total; leia as linhas acima).")
            }
        } else {
            $pingCmd = Get-Command ping -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
            $pingPath = $null
            if ($pingCmd) {
                if ($pingCmd.PSObject.Properties['Path'] -and $pingCmd.Path) { $pingPath = [string]$pingCmd.Path }
                elseif ($pingCmd.PSObject.Properties['Source'] -and $pingCmd.Source) { $pingPath = [string]$pingCmd.Source }
            }
            if (-not $pingPath) { throw "Comando ping nao encontrado neste sistema." }
            $proc = Start-Process -FilePath $pingPath -ArgumentList @('-c', '4', $dest) -NoNewWindow -Wait -PassThru
            if ($proc.ExitCode -ne 0) {
                Write-Warn ("Ping finalizou com codigo " + $proc.ExitCode.ToString() + ".")
            }
        }
    } catch {
        Write-Err $_.Exception.Message
    }

    Write-Host ""
    Write-Info "Dica: tempos estaveis e baixos (ex.: abaixo de 50 ms na mesma cidade) costumam ser bons sinais."
    Pause-Screen
}

function Show-TracertDidatico {
    Write-Banner
    Write-Centered "-- TRACERT (CAMINHO / ROTA) --" 'White'
    Write-Host ""
    Write-Host "  O que e o tracert?" -ForegroundColor Cyan
    Write-Host "  (No Linux costuma chamar-se traceroute - aqui usamos o comando do sistema.)"
    Write-Host "  Ele mostra cada ""salto"" (roteador) entre o seu PC e o destino, em ordem."
    Write-Host "  Assim da para ver em qual trecho a rota demora mais ou para de responder."
    Write-Host ""
    Write-Host "  Como ler rapidamente:" -ForegroundColor Yellow
    Write-Host "  - Cada linha e um passo na rede; tres tempos sao tres medicoes para aquele salto."
    Write-Host "  - ""*"" ou ""Request timed out"" em um salto nem sempre e problema - alguns roteadores nao respondem a ping."
    Write-Host "  - Se o destino final nunca aparece, pode haver bloqueio ou caminho interrompido."
    Write-Host ""

    $dest = Read-DestinoRede "Digite o host ou IP para rastrear a rota (somente o destino)."
    if (-not (Test-DestinoRedeSeguro $dest)) {
        Write-Err "Destino invalido ou vazio. Use apenas nome ou IP, sem espacos ou caracteres especiais."
        Pause-Screen
        return
    }

    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Info ("Rastreando rota (ate 30 saltos) para: " + $dest)
    Write-ThinDiv 'DarkGray'
    Write-Host ""

    try {
        if (Test-IsWindowsHost) {
            $exe = Join-Path $env:SystemRoot 'System32\tracert.exe'
            if (-not (Test-Path -LiteralPath $exe)) { throw "tracert.exe nao encontrado." }
            $proc = Start-Process -FilePath $exe -ArgumentList @('-h', '30', '-w', '2000', $dest) -NoNewWindow -Wait -PassThru
            if ($proc.ExitCode -ne 0) {
                Write-Warn ("Tracert finalizou com codigo " + $proc.ExitCode.ToString() + ".")
            }
        } else {
            $tr = Get-Command traceroute -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
            $trPath = $null
            if ($tr) {
                if ($tr.PSObject.Properties['Path'] -and $tr.Path) { $trPath = [string]$tr.Path }
                elseif ($tr.PSObject.Properties['Source'] -and $tr.Source) { $trPath = [string]$tr.Source }
            }
            if ($trPath) {
                $proc = Start-Process -FilePath $trPath -ArgumentList @('-m', '30', $dest) -NoNewWindow -Wait -PassThru
            } else {
                $tp = Get-Command tracepath -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
                $tpPath = $null
                if ($tp) {
                    if ($tp.PSObject.Properties['Path'] -and $tp.Path) { $tpPath = [string]$tp.Path }
                    elseif ($tp.PSObject.Properties['Source'] -and $tp.Source) { $tpPath = [string]$tp.Source }
                }
                if ($tpPath) {
                    $proc = Start-Process -FilePath $tpPath -ArgumentList @($dest) -NoNewWindow -Wait -PassThru
                } else {
                    throw "Nao encontrei traceroute nem tracepath. Instale um deles ou use o menu no Windows."
                }
            }
        }
    } catch {
        Write-Err $_.Exception.Message
    }

    Write-Host ""
    Write-Info "Dica: compare com o ping - se o ping funciona mas o tracert trava no meio, ainda pode haver rota assimetrica ou filtros."
    Pause-Screen
}

function Test-PortaTcpValida {
    param([string]$Texto)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return $false }
    $t = $Texto.Trim()
    if ($t -notmatch '^\d+$') { return $false }
    $n = 0
    if (-not [int]::TryParse($t, [ref]$n)) { return $false }
    return ($n -ge 1 -and $n -le 65535)
}

function Show-NslookupDidatico {
    Write-Banner
    Write-Centered "-- CONSULTA DNS (NSLOOKUP / HOST) --" 'White'
    Write-Host ""
    Write-Host "  O que e uma consulta DNS?" -ForegroundColor Cyan
    Write-Host "  O DNS traduz nomes (ex.: znuny.empresa.com) em enderecos IP. Se esta etapa"
    Write-Host "  falha, navegadores e aplicacoes podem nao achar o servidor mesmo com rede ok."
    Write-Host ""
    Write-Host "  O que ver no resultado:" -ForegroundColor Yellow
    Write-Host "  - Em geral aparece um ou mais enderecos IPv4 (A) ou IPv6 (AAAA)."
    Write-Host "  - Se aparecer ""NXDOMAIN"" ou ""not found"", o nome nao existe naquele servidor DNS."
    Write-Host "  - Se o IP parece errado, pode ser cache DNS ou entrada antiga no servidor."
    Write-Host ""

    $dest = Read-DestinoRede "Digite o nome de host ou IP para consultar no DNS (somente o destino)."
    if (-not (Test-DestinoRedeSeguro $dest)) {
        Write-Err "Destino invalido ou vazio."
        Pause-Screen
        return
    }

    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Info ("Consultando DNS para: " + $dest)
    Write-ThinDiv 'DarkGray'
    Write-Host ""

    try {
        if (Test-IsWindowsHost) {
            $exe = Join-Path $env:SystemRoot 'System32\nslookup.exe'
            if (-not (Test-Path -LiteralPath $exe)) { throw "nslookup.exe nao encontrado." }
            $null = Start-Process -FilePath $exe -ArgumentList @($dest) -NoNewWindow -Wait -PassThru
        } else {
            $hostAppCmd = Get-Command host -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
            $path = $null
            if ($hostAppCmd) {
                if ($hostAppCmd.PSObject.Properties['Path'] -and $hostAppCmd.Path) { $path = [string]$hostAppCmd.Path }
                elseif ($hostAppCmd.PSObject.Properties['Source'] -and $hostAppCmd.Source) { $path = [string]$hostAppCmd.Source }
            }
            if ($path) {
                $null = Start-Process -FilePath $path -ArgumentList @('-W', '3', $dest) -NoNewWindow -Wait -PassThru
            } else {
                $digCmd = Get-Command dig -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
                $digPath = $null
                if ($digCmd) {
                    if ($digCmd.PSObject.Properties['Path'] -and $digCmd.Path) { $digPath = [string]$digCmd.Path }
                    elseif ($digCmd.PSObject.Properties['Source'] -and $digCmd.Source) { $digPath = [string]$digCmd.Source }
                }
                if ($digPath) {
                    $null = Start-Process -FilePath $digPath -ArgumentList @('+time=2', '+tries=1', $dest) -NoNewWindow -Wait -PassThru
                } else {
                    throw "Nao encontrei host nem dig. No Linux instale bind-utils ou use o menu no Windows."
                }
            }
        }
    } catch {
        Write-Err $_.Exception.Message
    }

    Write-Host ""
    Write-Info "Dica: se o ping ao IP funciona mas o nome nao resolve, foque no DNS (servidor interno, hosts, VPN)."
    Pause-Screen
}

function Show-TestNetConnectionDidatico {
    Write-Banner
    Write-Centered "-- TESTE DE PORTA TCP --" 'White'
    Write-Host ""
    if (-not (Test-IsWindowsHost)) {
        Write-Host "  O Test-NetConnection e um cmdlet do Windows PowerShell." -ForegroundColor Yellow
        Write-Host "  Em Linux/macOS, equivalente aproximado no terminal:  nc -vz <host> <porta>" -ForegroundColor DarkGray
        Write-Host "  ou:  curl -v telnet://host:porta" -ForegroundColor DarkGray
        Pause-Screen
        return
    }
    if (-not (Get-Command Test-NetConnection -ErrorAction SilentlyContinue)) {
        Write-Err "Test-NetConnection nao disponivel nesta sessao (modulo NetTCPIP)."
        Pause-Screen
        return
    }

    Write-Host "  O que este teste faz?" -ForegroundColor Cyan
    Write-Host "  Tenta abrir uma conexao TCP ate a porta informada no destino - como um cliente"
    Write-Host "  checando se alguem ""escuta"" naquela porta. E diferente do ping (ICMP)."
    Write-Host ""
    Write-Host "  Interpretacao rapida:" -ForegroundColor Yellow
    Write-Host "  - TcpTestSucceeded = True: a porta aceitou conexao (servico provavelmente ativo ou firewall permitiu)."
    Write-Host "  - False: bloqueio de firewall, servico parado, ou host inalcancavel."
    Write-Host "  - Ping pode falhar e a porta 443 ainda responder (ICMP bloqueado e HTTPS liberado)."
    Write-Host ""

    $dest = Read-DestinoRede "Digite o host ou IP do servidor (somente o destino)."
    if (-not (Test-DestinoRedeSeguro $dest)) {
        Write-Err "Destino invalido ou vazio."
        Pause-Screen
        return
    }

    Write-Host ""
    Write-Host "  Porta TCP (1-65535). Exemplos: 80 HTTP, 443 HTTPS, 22 SSH, 5985 WinRM: " -ForegroundColor Yellow -NoNewline
    $portStr = Read-Host
    if (-not (Test-PortaTcpValida $portStr)) {
        Write-Err "Porta invalida. Use apenas numeros de 1 a 65535."
        Pause-Screen
        return
    }
    $portNum = [int]$portStr.Trim()

    Write-Host ""
    Write-ThinDiv 'DarkGray'
    Write-Info ("Test-NetConnection -ComputerName " + $dest + " -Port " + $portNum.ToString())
    Write-ThinDiv 'DarkGray'
    Write-Host ""

    try {
        $tncParams = @{
            ComputerName   = $dest
            Port             = $portNum
            WarningAction    = 'Continue'
            ErrorAction      = 'Stop'
        }
        if ((Get-Command Test-NetConnection).Parameters.Keys -contains 'InformationLevel') {
            $tncParams.InformationLevel = 'Detailed'
        }
        $r = Test-NetConnection @tncParams
        $r | Format-List ComputerName, RemoteAddress, RemotePort, TcpTestSucceeded, PingSucceeded, InterfaceAlias
    } catch {
        Write-Err $_.Exception.Message
    }

    Write-Host ""
    Write-Info "Dica: para o Znuny/HTTP em geral teste a porta 80 ou 443 conforme a URL que voce usa no navegador."
    Pause-Screen
}

function Show-MenuDiagnosticoRede {
    while ($true) {
        Write-Banner
        Write-Centered "-- DIAGNOSTICO DE REDE --" 'White'
        Write-Host ""
        Write-Host "  Escolha uma ferramenta. Em todas voce informa apenas host/IP (e na opcao 4 tambem a porta)." -ForegroundColor DarkGray
        Write-Host ""
        Write-MenuOpt '1' 'Ping'                    'Green'    'ICMP: alcance e latencia (4 tentativas)'
        Write-MenuOpt '2' 'Tracert / rota'        'Green'    'Saltos ate o destino (ate 30 hops)'
        Write-MenuOpt '3' 'Consulta DNS'          'Cyan'     'nslookup (Windows) ou host/dig (Linux)'
        Write-MenuOpt '4' 'Teste de porta TCP'    'Cyan'     'Windows: Test-NetConnection (host + porta)'
        Write-Host ""
        Write-MenuOpt '0' 'Voltar'                  'DarkGray' ''
        Write-Host ""
        Write-ThinDiv 'DarkGray'
        Write-Host ""
        Write-Host "  Opcao: " -ForegroundColor Cyan -NoNewline
        $ch = (Read-Host).Trim()
        switch ($ch) {
            '1' { Show-PingDidatico }
            '2' { Show-TracertDidatico }
            '3' { Show-NslookupDidatico }
            '4' { Show-TestNetConnectionDidatico }
            '0' { return }
            default {
                Write-Warn "Opcao invalida."
                Start-Sleep -Milliseconds 500
            }
        }
    }
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
        Write-MenuOpt '3' 'Visualizar Chamados'     'Yellow'   'OTRS em tempo real ou cache local; alerta de normalizacao'
        Write-Host ""
        Write-MenuOpt '4' 'Alterar Credenciais'     'Cyan'     'Usuario, senha e URL'
        Write-MenuOpt '5' 'Configuracoes'           'Cyan'     'Busca, cache, Hub, pausas HTTP (performance)'
        Write-MenuOpt '6' 'Salvar Credenciais'      'DarkGray' ''
        Write-MenuOpt '7' 'Sincronizar com Hub'   'Magenta'  'Login /api/login e envio para /api/relatorio (sem ocorrencia)'
        Write-Host ""
        Write-MenuOpt '8' 'Diagnostico de rede'   'Green'    'Submenu: ping, tracert, DNS e teste de porta TCP (didatico)'
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
            '7' { Invoke-SyncHub $Cfg $ExportScript }
            '8' { Show-MenuDiagnosticoRede }
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