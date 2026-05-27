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

# -----------------------------------------------------------------------------
# Hub: padroes se config.json nao definir HubEmail / HubPassword.
# Edite abaixo; valores em config.json (menu 5) prevalecem quando a chave existir.
# ATENCAO: senha em texto claro. Nao publique este arquivo em repositorio publico.
# -----------------------------------------------------------------------------
$script:MenuOtrsHubDefaultEmail     = 'thiago.ratanaka@microset.net.br'
$script:MenuOtrsHubDefaultPassword  = 'SenhaN2M7@'

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
        BaseURL            = 'http://172.16.0.12/znuny'
        Username           = ''
        Password           = ''
        EstadoFile         = 'estado_chamados.json'
        OutputPath         = $PWD.Path
        SearchPath         = 'index.pl?Action=AgentKPISearch;Subaction=Search;TakeLastSearch=1;Profile=94_8'
        HubBaseURL           = 'http://172.16.0.49:3210'
        HubEncaminharPath    = 'api/relatorio'
        HubEmail             = ([string]$script:MenuOtrsHubDefaultEmail).Trim()
        HubPassword          = [string]$script:MenuOtrsHubDefaultPassword
        HubApiRelatorioPath  = 'api/relatorio'
        HubPostTicketPaths   = ''
        HubPutTicketPaths    = ''
        HubFormSelectors     = $null
        HubWebDriverEnabled  = $false
        HubWebDriverBrowser  = 'Chrome'
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
            if ($j.HubBaseURL)           { $cfg.HubBaseURL           = $j.HubBaseURL           }
            if ($j.HubEncaminharPath)    { $cfg.HubEncaminharPath    = $j.HubEncaminharPath    }
            if ($j.HubApiRelatorioPath)  { $cfg.HubApiRelatorioPath  = $j.HubApiRelatorioPath  }
            $propNames = @($j.PSObject.Properties | ForEach-Object { $_.Name })
            if ($propNames -contains 'HubPostTicketPaths') { $cfg.HubPostTicketPaths = [string]$j.HubPostTicketPaths }
            if ($propNames -contains 'HubPutTicketPaths')  { $cfg.HubPutTicketPaths  = [string]$j.HubPutTicketPaths  }
            if ($propNames -contains 'HubEmail')    { $cfg.HubEmail    = [string]$j.HubEmail }
            if ($propNames -contains 'HubPassword') { $cfg.HubPassword = [string]$j.HubPassword }
            if ($propNames -contains 'HubFormSelectors' -and $null -ne $j.HubFormSelectors) {
                $cfg.HubFormSelectors = $j.HubFormSelectors
            }
            if ($propNames -contains 'HubWebDriverEnabled') {
                $v = $j.HubWebDriverEnabled
                if ($v -is [bool]) { $cfg.HubWebDriverEnabled = $v }
                elseif ($v -is [string]) { $cfg.HubWebDriverEnabled = $v.Trim() -match '^(1|true|s|S|sim|SIM|yes|YES)$' }
                else { $cfg.HubWebDriverEnabled = [bool]$v }
            }
            if ($propNames -contains 'HubWebDriverBrowser') {
                $cfg.HubWebDriverBrowser = [string]$j.HubWebDriverBrowser
            }
        } catch { }
    }
    return $cfg
}

function Save-Config {
    param([hashtable]$Cfg, [string]$Path)
    [ordered]@{
        BaseURL     = $Cfg.BaseURL
        Username    = $Cfg.Username
        Password    = $Cfg.Password
        EstadoFile  = $Cfg.EstadoFile
        OutputPath  = $Cfg.OutputPath
        SearchPath  = $Cfg.SearchPath
        HubBaseURL          = $Cfg.HubBaseURL
        HubEncaminharPath   = $Cfg.HubEncaminharPath
        HubEmail            = $Cfg.HubEmail
        HubPassword         = $Cfg.HubPassword
        HubApiRelatorioPath = $Cfg.HubApiRelatorioPath
        HubPostTicketPaths  = $Cfg.HubPostTicketPaths
        HubPutTicketPaths   = $Cfg.HubPutTicketPaths
        HubFormSelectors    = $Cfg.HubFormSelectors
        HubWebDriverEnabled = $Cfg.HubWebDriverEnabled
        HubWebDriverBrowser = $Cfg.HubWebDriverBrowser
    } | ConvertTo-Json -Depth 12 | Out-File $Path -Encoding UTF8
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
        if (-not $response) { return '' }
        $bytes = $null
        if ($response.RawContentStream) {
            try { $bytes = $response.RawContentStream.ToArray() } catch { $bytes = $null }
        }
        if (-not $bytes -or $bytes.Length -eq 0) {
            if ($null -ne $response.Content) { return [string]$response.Content }
            return ''
        }
        return [OtrsClient]::DecodeResponseBody($bytes, $response)
    }

    static [string] DecodeResponseBody([byte[]]$Bytes, $Response) {
        $headerCharset = ''
        $contentType = ''
        if ($Response.BaseResponse -and $Response.BaseResponse.ContentType) {
            $contentType = [string]$Response.BaseResponse.ContentType
        }
        elseif ($Response.Headers -and $Response.Headers['Content-Type']) {
            $contentType = [string]$Response.Headers['Content-Type']
        }
        if ($contentType -match 'charset\s*=\s*([^;\s]+)') {
            $headerCharset = $Matches[1].Trim().Trim('"').Trim("'")
        }

        $metaCharset = ''
        if ($Bytes.Length -gt 10) {
            $probeLen = [Math]::Min($Bytes.Length, 16384)
            $probeEnc = [System.Text.Encoding]::ASCII
            try {
                $head = $probeEnc.GetString($Bytes, 0, $probeLen)
                if ($head -match '(?i)charset\s*=\s*["'']?([^"''\s>]+)') { $metaCharset = $Matches[1].Trim() }
            } catch { }
        }

        $candidates = @($headerCharset, $metaCharset) | Where-Object { $_ }
        foreach ($name in $candidates) {
            try {
                $enc = [System.Text.Encoding]::GetEncoding($name)
                return $enc.GetString($Bytes)
            } catch { }
        }

        # UTF-8 BOM
        if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
            return [System.Text.Encoding]::UTF8.GetString($Bytes, 3, $Bytes.Length - 3)
        }

        $utf8 = [System.Text.Encoding]::UTF8
        $utf8Lenient = New-Object System.Text.UTF8Encoding $false, $false
        $asUtf8 = $utf8Lenient.GetString($Bytes)
        if ($asUtf8.IndexOf([char]0xFFFD) -lt 0) {
            return $asUtf8
        }

        foreach ($encName in @('windows-1252', 'iso-8859-1')) {
            try {
                return [System.Text.Encoding]::GetEncoding($encName).GetString($Bytes)
            } catch { }
        }
        return $utf8.GetString($Bytes)
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
function Copy-TicketWithArticleLimit {
    param(
        [TicketData]$Ticket,
        [int]$MaxNotes
    )
    if ($null -eq $Ticket) { return $null }
    if ($MaxNotes -le 0) { return $Ticket }
    $arts = @($Ticket.Articles)
    if ($arts.Count -le $MaxNotes) { return $Ticket }
    $short = [System.Collections.Generic.List[object]]::new()
    $start = $arts.Count - $MaxNotes
    for ($i = $start; $i -lt $arts.Count; $i++) {
        $short.Add($arts[$i])
    }
    return [TicketData]::new($Ticket.Numero, $Ticket.Estado, $Ticket.Criado, $Ticket.Cliente, $Ticket.Unidade, $short)
}

function Copy-TicketsWithArticleLimit {
    param(
        $Tickets,
        [int]$MaxNotes
    )
    if ($MaxNotes -le 0) { return $Tickets }
    $out = [System.Collections.Generic.List[TicketData]]::new()
    foreach ($t in $Tickets) {
        $out.Add((Copy-TicketWithArticleLimit $t $MaxNotes))
    }
    return @($out)
}

function Read-NotasExportOption {
    Write-Host ""
    Write-Host "  Notas por chamado no arquivo de saida:" -ForegroundColor Yellow
    Write-Host "  [1] Todas as notas"
    Write-Host "  [2] Apenas as 5 mais recentes"
    Write-Host "  Opcao [1]: " -ForegroundColor Cyan -NoNewline
    $o = (Read-Host).Trim()
    if ($o -eq '2') {
        Write-Info "Exportacao com as 5 notas mais recentes por chamado (o cache JSON interno permanece com todas as notas)."
        return 5
    }
    Write-Info "Exportacao com todas as notas obtidas de cada chamado."
    return 0
}

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
        [int]$MaxNotesExport = 0,
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
    Write-Verbose "MaxNotesExport (TXT): $(if ($MaxNotesExport -gt 0) { $MaxNotesExport } else { 'todas' })"

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
        $ativosParaTxt    = Copy-TicketsWithArticleLimit $ativos    $MaxNotesExport
        $resolvParaTxt    = Copy-TicketsWithArticleLimit $resolvidos $MaxNotesExport
        $ativoContent     = New-CcoFileContent -Tickets $ativosParaTxt -Titulo $script:CcoConfig.TituloRelatorio -EmptyMessage "Nenhum chamado critico ativo no momento."
        $resolvidoContent = New-CcoFileContent -Tickets $resolvParaTxt -Titulo $script:CcoConfig.TituloResolvidos -EmptyMessage "Nenhum chamado normalizado."

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
# JSON resumo do relatorio (Ativos / Resolvidos) a partir do cache EstadoFile
# =============================================================================
function Export-RelatorioCcoJsonFile {
    param(
        [hashtable]$Cfg,
        [int]$MaxNotesExport = 0
    )

    $cachePath = $Cfg.EstadoFile
    if (-not [System.IO.Path]::IsPathRooted($cachePath)) {
        $cachePath = Join-Path $Cfg.OutputPath $cachePath
    }
    if (-not (Test-Path $cachePath)) { return $null }

    $now      = Get-Date
    $jsonName = "Relatorio_CCO_" + $now.ToString('yyyy-MM-dd_HH-mm') + ".json"
    $jsonPath = Join-Path $Cfg.OutputPath $jsonName

    $rawCache = Get-Content $cachePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $ativos     = [System.Collections.Generic.List[object]]::new()
    $resolvidos = [System.Collections.Generic.List[object]]::new()
    $estadosResolvidos = $script:CcoConfig.EstadosResolvidos

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
            if ($MaxNotesExport -gt 0 -and $notaList.Count -gt $MaxNotesExport) {
                $from = $notaList.Count - $MaxNotesExport
                $trimmed = [System.Collections.Generic.List[object]]::new()
                for ($i = $from; $i -lt $notaList.Count; $i++) {
                    $trimmed.Add($notaList[$i])
                }
                $notaList = $trimmed
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
        Gerado            = $now.ToString('dd/MM/yyyy HH:mm:ss')
        NotasPorChamado   = if ($MaxNotesExport -gt 0) { "ultimas_$MaxNotesExport" } else { "todas" }
        TotalAtivos       = $ativos.Count
        TotalResolvidos   = $resolvidos.Count
        Ativos            = $ativos
        Resolvidos        = $resolvidos
    }
    $report | ConvertTo-Json -Depth 8 | Out-File $jsonPath -Encoding UTF8
    return [pscustomobject]@{
        Path              = $jsonPath
        TotalAtivos       = $ativos.Count
        TotalResolvidos   = $resolvidos.Count
    }
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
    $maxNotesExport = Read-NotasExportOption
    try {
        Ensure-ExportScript $ExportScript
    } catch {
        Write-Err ("Falha ao carregar script: " + $_); Pause-Screen; return
    }
    $params = @{
        BaseURL         = $Cfg.BaseURL
        Username        = $Cfg.Username
        Password        = $Cfg.Password
        SearchPath      = $Cfg.SearchPath
        EstadoFile      = $Cfg.EstadoFile
        OutputDir       = $Cfg.OutputPath
        AbrirRelatorio  = $false
        DiagMode        = $false
        MaxNotesExport  = $maxNotesExport
    }
    try {
        Export-CcoReport @params
        Write-Host ""
        Write-OK "Relatorio TXT gerado com sucesso."
        $jsonMeta = Export-RelatorioCcoJsonFile $Cfg -MaxNotesExport $maxNotesExport
        if ($jsonMeta) {
            Write-OK ("Resumo JSON salvo em: " + $jsonMeta.Path)
            Write-Info ("Ativos: " + $jsonMeta.TotalAtivos + "   Resolvidos: " + $jsonMeta.TotalResolvidos)
            Write-Info "Atualizacao TXT e JSON de resumo concluida na mesma execucao."
        } else {
            Write-Warn "Cache nao encontrado apos export; JSON de resumo nao foi gerado."
        }
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
    $maxNotesExport = Read-NotasExportOption
    try {
        Ensure-ExportScript $ExportScript
    } catch {
        Write-Err ("Falha ao carregar script: " + $_); Pause-Screen; return
    }
    $params = @{
        BaseURL         = $Cfg.BaseURL
        Username        = $Cfg.Username
        Password        = $Cfg.Password
        SearchPath      = $Cfg.SearchPath
        EstadoFile      = $Cfg.EstadoFile
        OutputDir       = $Cfg.OutputPath
        AbrirRelatorio  = $false
        DiagMode        = $false
        MaxNotesExport  = $maxNotesExport
    }
    try {
        Export-CcoReport @params
    } catch {
        Write-Err ("Erro na exportacao: " + $_); Pause-Screen; return
    }

    $jsonMeta = Export-RelatorioCcoJsonFile $Cfg -MaxNotesExport $maxNotesExport
    if (-not $jsonMeta) {
        Write-Warn "Cache nao encontrado, nenhum JSON de resumo exportado."
        Pause-Screen; return
    }

    Write-Host ""
    Write-OK ("JSON salvo em: " + $jsonMeta.Path)
    Write-Info ("TXT do CCO atualizado na mesma execucao. Ativos: " + $jsonMeta.TotalAtivos + "   Resolvidos: " + $jsonMeta.TotalResolvidos)
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
        [int]$RecentArticlesOnly = 4
    )
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
    Write-Info "Conectando e buscando chamados ativos..."
    Write-Host ""

    try {
        Ensure-ExportScript $ExportScript
        $tickets = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam
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
                $newT = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam
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
                $newT = Get-LiveTickets $Cfg -RecentArticlesOnly $recentParam
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
    Write-MenuOpt '1' 'OTRS tempo real (4 notas) ' 'White'    'Atualiza a cada 60s; apenas as 4 notas mais recentes por chamado'
    Write-MenuOpt '2' 'OTRS tempo real (todas)   ' 'White'    'Atualiza a cada 60s; todas as notas de cada chamado ativo (mais lento)'
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


# Compat: copias antigas chamavam Show-Visualizador no menu (funcao inexistente).
function Show-Visualizador {
    param([hashtable]$Cfg, [string]$ExportScript)
    Show-VisualizadorMenu $Cfg $ExportScript
}


# =============================================================================
# Integracao com Hub (http://172.16.0.49:3210)
# =============================================================================
$script:HubSession = $null

function Get-HubRelatorioPathPrefix {
    param([hashtable]$Cfg)
    $p = 'api/relatorio'
    if ($null -ne $Cfg.HubApiRelatorioPath -and $Cfg.HubApiRelatorioPath.ToString().Trim()) {
        $p = $Cfg.HubApiRelatorioPath.ToString().Trim().TrimStart('/')
    }
    if (-not $p.StartsWith('/')) { $p = '/' + $p }
    return $p
}

function Get-HubHttpErrorDetail {
    param($ErrRecord)
    try {
        if ($ErrRecord.ErrorDetails -and $ErrRecord.ErrorDetails.Message) {
            $ed = ($ErrRecord.ErrorDetails.Message -as [string]).Trim()
            if ($ed.Length -gt 0) { return $ed }
        }
    } catch { }
    $msg = $ErrRecord.Exception.Message
    try {
        $resp = $ErrRecord.Exception.Response
        if ($resp) {
            $code = [int]$resp.StatusCode
            $stream = $resp.GetResponseStream()
            if ($stream) {
                $rdr = New-Object System.IO.StreamReader($stream)
                $body = $rdr.ReadToEnd()
                if ($body) {
                    $short = ($body -replace '[\r\n]+', ' ')
                    if ($short.Length -gt 280) { $short = $short.Substring(0, 277) + '...' }
                    return ("HTTP " + $code + ": " + $short)
                }
                return ("HTTP " + $code + ": " + $msg)
            }
        }
    } catch { }
    return $msg
}

function Hub-Login {
    param([string]$HubUrl, [string]$Email, [string]$Pass)
    $base = $HubUrl.TrimEnd('/')
    $loginBody = (@{ email = $Email; password = $Pass } | ConvertTo-Json -Compress)
    $resp = Invoke-WebRequest -Uri ($base + "/api/login") `
        -Method POST -ContentType "application/json; charset=utf-8" -Body $loginBody `
        -SessionVariable 'WebSess' -UseBasicParsing
    $script:HubSession = $WebSess
    return $resp.Content | ConvertFrom-Json
}

function Hub-Get {
    param([string]$HubUrl, [string]$Path)
    $base = $HubUrl.TrimEnd('/')
    if (-not $Path.StartsWith('/')) { $Path = '/' + $Path }
    $resp = Invoke-WebRequest -Uri ($base + $Path) `
        -Method GET -WebSession $script:HubSession -UseBasicParsing `
        -Headers @{ Accept = 'application/json' }
    $raw = ($resp.Content -as [string]).Trim()
    if ($raw.Length -eq 0) { return $null }
    if ($raw.Length -lt 2) { return $null }
    $c0 = $raw[0]
    if ($c0 -ne '[' -and $c0 -ne '{') {
        throw ("Resposta nao-JSON em GET " + $Path + ": " + ($raw.Substring(0, [Math]::Min(200, $raw.Length))))
    }
    return $raw | ConvertFrom-Json
}

function Get-HubTicketsArrayFromResponse {
    param($Json)
    if ($null -eq $Json) { return @() }
    if ($Json -is [System.Array]) { return ,@($Json) }
    $names = @($Json.PSObject.Properties | ForEach-Object { $_.Name })
    if ($names -contains 'tickets') {
        $t = $Json.tickets
        if ($null -eq $t) { return @() }
        return @($t)
    }
    if ($names -contains 'number' -or $names -contains 'id') {
        return ,@($Json)
    }
    return @()
}

function Get-HubRelatorioTicketsList {
    param([string]$HubUrl, [string]$ApiRelRoot)
    $root = $ApiRelRoot.TrimEnd('/')
    if (-not $root.StartsWith('/')) { $root = '/' + $root.TrimStart('/') }
    $tryPaths = @( ($root + '/tickets'), $root )
    $lastEx = $null
    foreach ($p in $tryPaths) {
        try {
            $j = Hub-Get $HubUrl $p
            return ,@(Get-HubTicketsArrayFromResponse $j)
        } catch {
            $lastEx = $_
        }
    }
    if ($lastEx) { throw $lastEx }
    return @()
}

function Hub-Post {
    param([string]$HubUrl, [string]$Path, $Body)
    $base = $HubUrl.TrimEnd('/')
    if (-not $Path.StartsWith('/')) { $Path = '/' + $Path }
    $uri = $base + $Path
    $json = $Body | ConvertTo-Json -Depth 8
    try {
        $resp = Invoke-WebRequest -Uri $uri `
            -Method POST -ContentType "application/json; charset=utf-8" -Body $json `
            -WebSession $script:HubSession -UseBasicParsing `
            -Headers @{ Accept = 'application/json' } -ErrorAction Stop
        $out = ($resp.Content -as [string]).Trim()
        if ($out.Length -eq 0) { return $null }
        return $out | ConvertFrom-Json
    } catch {
        throw ((Get-HubHttpErrorDetail $_) + " (POST " + $uri + ")")
    }
}

function Hub-Put {
    param([string]$HubUrl, [string]$Path, $Body)
    $base = $HubUrl.TrimEnd('/')
    if (-not $Path.StartsWith('/')) { $Path = '/' + $Path }
    $uri = $base + $Path
    $json = $Body | ConvertTo-Json -Depth 8
    try {
        $resp = Invoke-WebRequest -Uri $uri `
            -Method PUT -ContentType "application/json; charset=utf-8" -Body $json `
            -WebSession $script:HubSession -UseBasicParsing `
            -Headers @{ Accept = 'application/json' } -ErrorAction Stop
        $out = ($resp.Content -as [string]).Trim()
        if ($out.Length -eq 0) { return $null }
        return $out | ConvertFrom-Json
    } catch {
        throw ((Get-HubHttpErrorDetail $_) + " (PUT " + $uri + ")")
    }
}

function Get-HubPostTicketCandidatePaths {
    param([string]$ApiRelRoot, [hashtable]$Cfg)
    $root = $ApiRelRoot.TrimEnd('/')
    $list = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    function Add-HubPath([string]$p) {
        if (-not $p) { return }
        if (-not $p.StartsWith('/')) { $p = '/' + $p.TrimStart('/') }
        if ($seen.ContainsKey($p)) { return }
        [void]$seen.Add($p, $true)
        $list.Add($p)
    }
    $custom = ($Cfg.HubPostTicketPaths -as [string])
    if ($custom -and $custom.Trim()) {
        foreach ($part in ($custom -split '[,;\|]')) {
            Add-HubPath $part.Trim()
        }
        return ,$list.ToArray()
    }
    Add-HubPath ($root + '/tickets')
    Add-HubPath ($root + '/ticket')
    Add-HubPath $ApiRelRoot
    Add-HubPath '/api/ticket'
    Add-HubPath '/api/tickets'
    Add-HubPath '/api/relatorio/ticket'
    Add-HubPath '/api/cco/ticket'
    return ,$list.ToArray()
}

function Get-HubPutTicketCandidatePaths {
    param([string]$ApiRelRoot, [string]$Num, [hashtable]$Cfg)
    $root = $ApiRelRoot.TrimEnd('/')
    $list = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    function Add-HubPath([string]$p) {
        if (-not $p) { return }
        $p = $p.Replace('{numero}', $Num).Replace('{n}', $Num)
        if (-not $p.StartsWith('/')) { $p = '/' + $p.TrimStart('/') }
        if ($seen.ContainsKey($p)) { return }
        [void]$seen.Add($p, $true)
        $list.Add($p)
    }
    $custom = ($Cfg.HubPutTicketPaths -as [string])
    if ($custom -and $custom.Trim()) {
        foreach ($part in ($custom -split '[,;\|]')) {
            Add-HubPath $part.Trim()
        }
        return ,$list.ToArray()
    }
    Add-HubPath ($root + '/tickets/' + $Num)
    Add-HubPath ($root + '/' + $Num)
    Add-HubPath ($root + '/ticket/' + $Num)
    Add-HubPath ('/api/tickets/' + $Num)
    Add-HubPath ('/api/ticket/' + $Num)
    return ,$list.ToArray()
}

function Hub-PostWithPathFallback {
    param([string]$HubUrl, [string[]]$Paths, $Body)
    $lastEx = $null
    foreach ($p in $Paths) {
        try {
            $r = Hub-Post $HubUrl $p $Body
            Write-Info ("Hub: POST aceito em " + $p)
            return $r
        } catch {
            $lastEx = $_
        }
    }
    throw $lastEx
}

function Hub-PutWithPathFallback {
    param([string]$HubUrl, [string[]]$Paths, $Body)
    $lastEx = $null
    foreach ($p in $Paths) {
        try {
            $r = Hub-Put $HubUrl $p $Body
            Write-Info ("Hub: PUT aceito em " + $p)
            return $r
        } catch {
            $lastEx = $_
        }
    }
    throw $lastEx
}

function Export-HubTicketAssistantHtml {
    param([string]$JsonText, [string]$HubPageUrl, [string]$TicketNum, [string]$OutPath)
    $b64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($JsonText))
    $escUrl = [System.Net.WebUtility]::HtmlEncode($HubPageUrl)
    $html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<title>Hub ticket #$TicketNum — JSON</title>
<style>
body{font-family:Segoe UI,sans-serif;background:#1e293b;color:#e2e8f0;padding:1.5rem;max-width:56rem;margin:0 auto;}
textarea{width:100%;min-height:14rem;font-family:Consolas,monospace;font-size:12px;background:#0f172a;color:#e2e8f0;border:1px solid #475569;border-radius:6px;padding:10px;}
button,a.btn{margin:8px 8px 0 0;padding:10px 16px;border-radius:6px;border:none;cursor:pointer;font-weight:600;text-decoration:none;display:inline-block;}
button{background:#2563eb;color:#fff;}
a.btn{background:#059669;color:#fff;}
p{color:#94a3b8;font-size:14px;}
</style>
</head>
<body>
<h1>Ticket #$TicketNum</h1>
<p>O POST automatico para a API do Hub falhou. Use o JSON abaixo: copie e siga o fluxo do Gerador CCO no navegador (ou peça ao desenvolvedor a rota POST correta).</p>
<p><a class="btn" href="$escUrl" target="_blank" rel="noopener">Abrir Gerador CCO</a></p>
<textarea id="jx" readonly></textarea>
<p>
<button type="button" id="cp">Copiar JSON</button>
</p>
<script>
(function(){
  var b64 = '$b64';
  var bin = atob(b64);
  var bytes = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) { bytes[i] = bin.charCodeAt(i); }
  var txt = new TextDecoder('utf-8').decode(bytes);
  document.getElementById('jx').value = txt;
  document.getElementById('cp').onclick = function() {
    var el = document.getElementById('jx');
    el.select();
    el.setSelectionRange(0, 99999999);
    try { navigator.clipboard.writeText(el.value); alert('Copiado.'); }
    catch (e) { document.execCommand('copy'); alert('Tente Ctrl+C.'); }
  };
})();
</script>
</body>
</html>
"@
    $html | Out-File -LiteralPath $OutPath -Encoding utf8
}

function Invoke-HubManualPayloadAssist {
    param([string]$HubUrl, [string]$RelPagePath, $Payload, [string]$TicketNum, [hashtable]$Cfg)
    $json = $Payload | ConvertTo-Json -Depth 8
    $outDir = $Cfg.OutputPath
    if (-not [System.IO.Path]::IsPathRooted($outDir)) { $outDir = Join-Path $PWD.Path $outDir }
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMddHHmmss'
    $jsonPath = Join-Path $outDir ('Hub_ticket_' + $TicketNum + '_' + $stamp + '.json')
    $htmlPath = Join-Path $outDir ('Hub_ticket_' + $TicketNum + '_' + $stamp + '_ajuda.html')
    $json | Out-File -LiteralPath $jsonPath -Encoding utf8
    $rel = if ($RelPagePath) { $RelPagePath.Trim().TrimStart('/') } else { 'api/relatorio' }
    $hubPage = $HubUrl.TrimEnd('/') + '/' + $rel
    Export-HubTicketAssistantHtml -JsonText $json -HubPageUrl $hubPage -TicketNum $TicketNum -OutPath $htmlPath
    $clipOk = $false
    try {
        $json | Set-Clipboard
        $clipOk = $true
    } catch { }
    try { Start-Process -FilePath $hubPage } catch { }
    try { Start-Process -FilePath $htmlPath } catch { }
    Write-Host ""
    Write-Warn "Nenhuma rota POST automatica foi aceita pelo servidor (ou sessao sem permissao)."
    Write-Info "JSON gravado:"
    Write-Host ("  " + $jsonPath) -ForegroundColor Yellow
    Write-Info "Pagina de ajuda (copiar JSON / link Hub):"
    Write-Host ("  " + $htmlPath) -ForegroundColor Yellow
    if ($clipOk) {
        Write-Info "O JSON tambem foi copiado para a area de transferencia (Ctrl+V)."
    }
    Write-Info "Abra o Gerador CCO, adicione o ticket e preencha com estes dados se o Hub nao tiver importacao automatica."
}

function Build-HubTicket {
    param($Ticket, [int]$MaxUpdates = 10, [string]$Ocorrencia = '')

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
        $updates.Add([ordered]@{ updateDate=$uDate; updateHour=$uHour; text=$n.Text })
    }

    # Campos alinhados ao relatorioCco.js / API Hub: updates com updateDate, updateHour, text.
    # Ocorrencia e opcional (terminal na sincronizacao); inclui ocorrencia + occurrence no JSON se preenchida.
    $h = [ordered]@{
        number      = $Ticket.Numero
        status      = $Ticket.Estado
        openingDate = $openDate
        openingHour = $openHour
        client      = $Ticket.Cliente
        updates     = @($updates)
    }
    $occTrim = ($Ocorrencia -as [string]).Trim()
    if ($occTrim.Length -gt 0) {
        $h.ocorrencia  = $occTrim
        $h.occurrence  = $occTrim
    }
    return $h
}

function Test-HubRelatorioChanged {
    param($hubTicket, $payload)
    if (-not $hubTicket) { return $true }
    if ($payload.Keys -contains 'occurrence') {
        if (($hubTicket.occurrence -or '') -ne ($payload.occurrence -or '')) { return $true }
    }
    if ($payload.Keys -contains 'ocorrencia') {
        if (($hubTicket.ocorrencia -or '') -ne ($payload.ocorrencia -or '')) { return $true }
    }
    if (($hubTicket.status -or '') -ne ($payload.status -or '')) { return $true }
    if (($hubTicket.openingDate -or '') -ne ($payload.openingDate -or '')) { return $true }
    if (($hubTicket.openingHour -or '') -ne ($payload.openingHour -or '')) { return $true }
    if (($hubTicket.client -or '') -ne ($payload.client -or '')) { return $true }
    $hu = @($hubTicket.updates)
    $pu = @($payload.updates)
    if ($hu.Count -ne $pu.Count) { return $true }
    for ($i = 0; $i -lt $hu.Count; $i++) {
        $hd = if ($hu[$i].updateDate) { [string]$hu[$i].updateDate } elseif ($hu[$i].date) { [string]$hu[$i].date } else { '' }
        $hh = if ($hu[$i].updateHour) { [string]$hu[$i].updateHour } elseif ($hu[$i].hour) { [string]$hu[$i].hour } else { '' }
        $pd = if ($pu[$i].updateDate) { [string]$pu[$i].updateDate } elseif ($pu[$i].date) { [string]$pu[$i].date } else { '' }
        $ph = if ($pu[$i].updateHour) { [string]$pu[$i].updateHour } elseif ($pu[$i].hour) { [string]$pu[$i].hour } else { '' }
        if (($hu[$i].text -or '') -ne ($pu[$i].text -or '')) { return $true }
        if ($hd -ne $pd) { return $true }
        if ($hh -ne $ph) { return $true }
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
    if ($Payload.ocorrencia -or $Payload.occurrence) {
        $oc = if ($Payload.ocorrencia) { [string]$Payload.ocorrencia } else { [string]$Payload.occurrence }
        if ($oc.Length -gt 160) { $oc = $oc.Substring(0, 157) + "..." }
        Write-Host "  Ocorren. : " -ForegroundColor DarkGray -NoNewline
        Write-Host $oc -ForegroundColor White
    }
}

function Read-OcorrenciaMultiline {
    Write-Host ""
    Write-Host "  Ocorrencia (opcional, como no formulario do Hub)." -ForegroundColor DarkGray
    Write-Host "  Linha vazia encerra o texto; apenas Enter = sem ocorrencia." -ForegroundColor DarkGray
    Write-Host ""
    $lines = New-Object System.Collections.Generic.List[string]
    while ($true) {
        Write-Host "  > " -ForegroundColor DarkGray -NoNewline
        $line = Read-Host
        if ($line -eq '' -and $lines.Count -eq 0) { return '' }
        if ($line -eq '') { return ($lines -join "`n") }
        $lines.Add($line)
    }
}

function Get-HubEncaminharUri {
    param([string]$HubUrl, [string]$RelativePath)
    $p = if ($RelativePath) { $RelativePath.Trim().TrimStart('/') } else { 'api/relatorio' }
    if (-not $p) { $p = 'api/relatorio' }
    return ($HubUrl.TrimEnd('/') + '/' + $p)
}

function Export-HubRelatorioFormHtml {
    param($Payload, [string]$HubBaseUrl, [string]$OutPath)

    function HtmlEsc([string]$t) {
        if ($null -eq $t) { return '' }
        return [System.Net.WebUtility]::HtmlEncode([string]$t)
    }

    $num   = HtmlEsc ([string]$Payload.number)
    $cli   = HtmlEsc ([string]$Payload.client)
    $st    = HtmlEsc ([string]$Payload.status)
    $od    = HtmlEsc ([string]$Payload.openingDate)
    $oh    = HtmlEsc ([string]$Payload.openingHour)
    $hubL  = HtmlEsc ($HubBaseUrl.TrimEnd('/'))

    $sbRows = New-Object System.Text.StringBuilder
    $ix = 0
    foreach ($u in @($Payload.updates)) {
        $ix++
        $udRaw = if ($u.updateDate) { $u.updateDate } else { $u.date }
        $uhRaw = if ($u.updateHour) { $u.updateHour } else { $u.hour }
        $ud = HtmlEsc ([string]$udRaw)
        $uh = HtmlEsc ([string]$uhRaw)
        $ut = HtmlEsc ([string]$u.text)
        $ut = ($ut -replace "`r`n", '<br/>') -replace "`n", '<br/>'
        [void]$sbRows.Append('<tr><td>').Append($ix).Append('</td><td>').Append($ud).Append(' ')
        [void]$sbRows.Append($uh).Append('</td><td class="note">').Append($ut).Append('</td></tr>')
    }

    if ($Payload.ocorrencia -or $Payload.occurrence) {
        $occRaw = if ($Payload.ocorrencia) { [string]$Payload.ocorrencia } else { [string]$Payload.occurrence }
        $occBody = (HtmlEsc $occRaw) -replace "`r`n", '<br/>' -replace "`n", '<br/>'
    } else {
        $occBody = '<span class="muted">(vazio neste envio)</span>'
    }

    $html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Relatorio CCO #$num</title>
<style>
body{font-family:Segoe UI,Tahoma,sans-serif;margin:1.25rem;background:#f4f6f8;color:#222;}
h1{font-size:1.15rem;margin:0 0 .5rem;}
.box{background:#fff;border:1px solid #ccd2d8;border-radius:6px;padding:1rem;margin:.75rem 0;max-width:52rem;}
table{border-collapse:collapse;width:100%;font-size:.88rem;}
th,td{border:1px solid #e2e6ea;padding:.4rem .5rem;vertical-align:top;}
th{background:#eef2f6;text-align:left;}
td.note{white-space:pre-wrap;}
.muted{color:#666;}
.hint{font-size:.85rem;color:#444;margin:.5rem 0 1rem;}
kbd{background:#eee;padding:.1rem .35rem;border-radius:3px;}
a{color:#0b5;}
</style>
</head>
<body>
<h1>Pre-visualizacao — Relatorio CCO (Hub)</h1>
<p class="hint">Revise os campos abaixo (espelho do que sera enviado pela API). Depois volte ao <kbd>Menu-OTRS</kbd> e confirme o envio. Hub: <a href="$hubL">$hubL</a></p>
<div class="box">
<table>
<tr><th>Ticket</th><td>#$num</td></tr>
<tr><th>Cliente</th><td>$cli</td></tr>
<tr><th>Estado</th><td>$st</td></tr>
<tr><th>Abertura</th><td>$od $oh</td></tr>
</table>
</div>
<div class="box">
<h2 style="font-size:1rem;margin-top:0;">Ocorrencia</h2>
<div class="note">$occBody</div>
</div>
<div class="box">
<h2 style="font-size:1rem;margin-top:0;">Atualizacoes</h2>
<table><thead><tr><th>#</th><th>Data</th><th>Nota</th></tr></thead>
<tbody>
$($sbRows.ToString())
</tbody></table>
</div>
</body>
</html>
"@
    $dir = [System.IO.Path]::GetDirectoryName($OutPath)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $html | Out-File -LiteralPath $OutPath -Encoding utf8
}

function Export-HubRelatorioAutofillHelperHtml {
    param($Payload, [string]$HubPageUrl, [string]$OutPath, [hashtable]$Cfg)

    $payloadJson = ($Payload | ConvertTo-Json -Depth 12 -Compress)
    $b64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadJson))
    $selLiteral = '{}'
    if ($Cfg -and $null -ne $Cfg.HubFormSelectors) {
        $selLiteral = ($Cfg.HubFormSelectors | ConvertTo-Json -Depth 12 -Compress)
    }

    $jsCore = @'
(function(){
  "use strict";
  var payload = JSON.parse(new TextDecoder("utf-8").decode(Uint8Array.from(atob("B64TOKEN"), function(c){ return c.charCodeAt(0); })));
  var selO = SELTOKEN;
  function selList(key) {
    var o = selO[key];
    var def = (DEFAULT_SEL[key] || []).slice();
    if (Array.isArray(o) && o.length) return o.concat(def);
    if (typeof o === "string" && o.trim()) return [o.trim()].concat(def);
    return def;
  }
  var DEFAULT_SEL = {
    number: ["input[name=\"number\"]","#number","input#number","[data-field=\"number\"]"],
    status: ["input[name=\"status\"]","select[name=\"status\"]","#status","[data-field=\"status\"]"],
    openingDate: ["input[name=\"openingDate\"]","#openingDate","[data-field=\"openingDate\"]"],
    openingHour: ["input[name=\"openingHour\"]","#openingHour","[data-field=\"openingHour\"]"],
    client: ["textarea[name=\"client\"]","input[name=\"client\"]","#client","[data-field=\"client\"]"],
    occurrence: ["textarea[name=\"occurrence\"]","textarea[name=\"ocorrencia\"]","#occurrence","#ocorrencia","[data-field=\"occurrence\"]"]
  };
  function setNativeValue(el, value) {
    if (!el || value === undefined || value === null) return false;
    var v = String(value);
    if (el.tagName === "SELECT") {
      el.value = v;
      el.dispatchEvent(new Event("change", { bubbles: true }));
      return true;
    }
    var proto = el.tagName === "TEXTAREA" ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
    var desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) { desc.set.call(el, v); }
    else { el.value = v; }
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }
  function tryFill(key, value) {
    var arr = selList(key);
    for (var i = 0; i < arr.length; i++) {
      try {
        var el = document.querySelector(arr[i]);
        if (el) {
          setNativeValue(el, value);
          console.log("[Menu-OTRS] Campo \"" + key + "\" via " + arr[i]);
          return true;
        }
      } catch (e) { console.warn(e); }
    }
    console.warn("[Menu-OTRS] Campo nao encontrado: " + key);
    return false;
  }
  tryFill("number", payload.number || "");
  tryFill("status", payload.status || "");
  tryFill("openingDate", payload.openingDate || "");
  tryFill("openingHour", payload.openingHour || "");
  tryFill("client", payload.client || "");
  var occ = (payload.occurrence || payload.ocorrencia || "");
  if (occ) tryFill("occurrence", occ);
  var updates = payload.updates || [];
  var tables = document.querySelectorAll("table");
  var best = null;
  for (var ti = 0; ti < tables.length; ti++) {
    var trs = tables[ti].querySelectorAll("tbody tr");
    if (trs.length >= updates.length && updates.length > 0) { best = trs; break; }
  }
  if (!best && tables.length) {
    var last = tables[tables.length - 1];
    best = last.querySelectorAll("tbody tr");
  }
  if (best && best.length && updates.length) {
    for (var r = 0; r < updates.length && r < best.length; r++) {
      var row = best[r];
      var fields = row.querySelectorAll("input, textarea");
      var u = updates[r];
      var ud = u.updateDate || u.date || "";
      var uh = u.updateHour || u.hour || "";
      var tx = u.text || "";
      if (fields.length >= 3) {
        setNativeValue(fields[0], ud);
        setNativeValue(fields[1], uh);
        setNativeValue(fields[2], tx);
        console.log("[Menu-OTRS] Linha atualizacao " + (r + 1));
      } else if (fields.length === 2) {
        setNativeValue(fields[0], ud);
        setNativeValue(fields[1], tx);
      } else if (fields.length === 1) {
        setNativeValue(fields[0], tx);
      }
    }
  } else {
    console.warn("[Menu-OTRS] Tabela de atualizacoes nao casou; verifique HubFormSelectors no config.json.");
  }
  console.log("[Menu-OTRS] Preenchimento concluido. Revise os campos e grave no Hub.");
})();
'@
    $jsCore = $jsCore.Replace('B64TOKEN', $b64).Replace('SELTOKEN', $selLiteral)

    $hubEsc = [System.Net.WebUtility]::HtmlEncode($HubPageUrl)
    $snippetEsc = [System.Net.WebUtility]::HtmlEncode($jsCore)
    $html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Preencher Hub — #$($Payload.number)</title>
<style>
body{font-family:Segoe UI,Tahoma,sans-serif;margin:1.25rem;background:#0f172a;color:#e2e8f0;max-width:48rem;}
h1{font-size:1.1rem;}
p,li{color:#94a3b8;font-size:.95rem;line-height:1.45;}
a.btn,a.btn:visited{display:inline-block;margin:.5rem .5rem 0 0;padding:.55rem 1rem;background:#2563eb;color:#fff;border-radius:6px;text-decoration:none;font-weight:600;}
textarea{width:100%;min-height:16rem;font-family:Consolas,monospace;font-size:11px;background:#020617;color:#e2e8f0;border:1px solid #334155;border-radius:6px;padding:10px;}
button{margin-top:.75rem;padding:.55rem 1rem;border-radius:6px;border:none;background:#059669;color:#fff;font-weight:600;cursor:pointer;}
.box{border:1px solid #334155;border-radius:8px;padding:1rem;margin:1rem 0;background:#1e293b;}
.warn{color:#fbbf24;}
</style>
</head>
<body>
<h1>Preencher o formulario do Gerador CCO no navegador</h1>
<p>A pagina <strong>$hubEsc</strong> e o formulario do Hub; por seguranca do navegador o Menu-OTRS <span class="warn">nao pode</span> escrever diretamente nessa aba a partir de um ficheiro local. Este fluxo copia um script para colar na <strong>Consola (F12)</strong> <em>ja com o separador do Hub activo</em>.</p>
<ol>
<li>Abra o Hub e entre na rota do Gerador (login se preciso).</li>
<li>Deixe essa aba em primeiro plano e pressione <strong>F12</strong> &gt; separador <strong>Consola</strong>.</li>
<li>Aqui em baixo: <strong>Copiar script</strong>, volte ao Hub, <strong>Cole</strong> na consola e <strong>Enter</strong>.</li>
<li>Confira os campos e use o botao de gravar / fluxo normal do Hub.</li>
</ol>
<p class="warn">Se algum campo nao mudar (React), ajuste os selectores em <code>HubFormSelectors</code> no <code>config.json</code> (menu 5) — ver README.</p>
<p><a class="btn" href="$hubEsc" target="_blank" rel="noopener">Abrir Gerador CCO no Hub</a></p>
<div class="box">
<label for="sn"><strong>Script para colar na consola do Hub</strong></label>
<textarea id="sn" readonly>$snippetEsc</textarea>
<p><button type="button" id="cp">Copiar script</button></p>
</div>
<script>
(function(){
  var t = document.getElementById("sn");
  document.getElementById("cp").onclick = function() {
    t.select();
    t.setSelectionRange(0, 99999999);
    try { navigator.clipboard.writeText(t.value); alert("Copiado. Cole na consola da aba do Hub."); }
    catch (e) { try { document.execCommand("copy"); alert("Copiado (modo legado)."); } catch (e2) { alert("Use Ctrl+C no texto."); } }
  };
})();
</script>
</body>
</html>
"@
    $dir = [System.IO.Path]::GetDirectoryName($OutPath)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $html | Out-File -LiteralPath $OutPath -Encoding utf8
}

function Show-HubRelatorioFormHtml {
    param($Payload, [string]$HubUrl, [hashtable]$Cfg)
    $stamp = Get-Date -Format 'yyyyMMddHHmmss'
    $num = [string]$Payload.number
    $fn = 'HubRelatorio_' + $num + '_' + $stamp + '.html'
    $path = Join-Path ([System.IO.Path]::GetTempPath()) $fn
    Export-HubRelatorioFormHtml -Payload $Payload -HubBaseUrl $HubUrl -OutPath $path
    Start-Process -FilePath $path
    if (-not $Cfg) { $Cfg = @{} }
    $relEnc = if ($Cfg.HubEncaminharPath -and $Cfg.HubEncaminharPath.ToString().Trim()) {
        $Cfg.HubEncaminharPath.ToString().Trim()
    } else { 'api/relatorio' }
    $hubPage = Get-HubEncaminharUri $HubUrl $relEnc
    $fn2 = 'HubRelatorio_PreencherHub_' + $num + '_' + $stamp + '.html'
    $path2 = Join-Path ([System.IO.Path]::GetTempPath()) $fn2
    Export-HubRelatorioAutofillHelperHtml -Payload $Payload -HubPageUrl $hubPage -OutPath $path2 -Cfg $Cfg
    Start-Process -FilePath $path2
}

function Test-HubSeleniumModuleAvailable {
    return [bool](Get-Module -ListAvailable -Name Selenium)
}

function Invoke-HubRelatorioSeleniumFill {
    param([string]$HubPageUrl, $Payload, [string]$Browser = 'Chrome')

    Import-Module Selenium -ErrorAction Stop

    $Driver = $null
    try {
        $br = ($Browser -as [string]).Trim()
        if ($br -match '^[Ee]dge') {
            if (Get-Command Start-SeEdge -ErrorAction SilentlyContinue) {
                $Driver = Start-SeEdge
            }
        }
        if (-not $Driver -and (Get-Command Start-SeChrome -ErrorAction SilentlyContinue)) {
            $Driver = Start-SeChrome
        }
        if (-not $Driver) {
            throw "Nao foi possivel iniciar Chrome nem Edge (Start-SeChrome / Start-SeEdge). Verifique o modulo Selenium."
        }

        Enter-SeUrl -Url $HubPageUrl -Driver $Driver

        function Get-SeFieldByName {
            param([string]$FieldName, [int]$TimeoutSec = 30)
            $deadline = (Get-Date).AddSeconds($TimeoutSec)
            while ((Get-Date) -lt $deadline) {
                try {
                    $el = Find-SeElement -Driver $Driver -Wait -Timeout 2 -Name $FieldName -ErrorAction SilentlyContinue
                    if ($el) { return $el }
                } catch { }
                Start-Sleep -Milliseconds 400
            }
            return $null
        }

        function Set-SeFieldValue {
            param($Element, [string]$Text)
            if ($null -eq $Element) { return }
            Send-SeKeys -Element $Element -Keys ([OpenQA.Selenium.Keys]::Control + 'a')
            Send-SeKeys -Element $Element -Keys $Text
        }

        $elNum = Get-SeFieldByName 'number'
        Set-SeFieldValue $elNum ([string]$Payload.number)

        $elSt = Get-SeFieldByName 'status'
        if ($elSt) { Set-SeFieldValue $elSt ([string]$Payload.status) }

        $elOd = Get-SeFieldByName 'openingDate'
        if ($elOd) { Set-SeFieldValue $elOd ([string]$Payload.openingDate) }

        $elOh = Get-SeFieldByName 'openingHour'
        if ($elOh) { Set-SeFieldValue $elOh ([string]$Payload.openingHour) }

        $elCl = Get-SeFieldByName 'client'
        if ($elCl) { Set-SeFieldValue $elCl ([string]$Payload.client) }

        if ($Payload.occurrence -or $Payload.ocorrencia) {
            $occ = if ($Payload.occurrence) { [string]$Payload.occurrence } else { [string]$Payload.ocorrencia }
            $elOc = Get-SeFieldByName 'occurrence'
            if (-not $elOc) { $elOc = Get-SeFieldByName 'ocorrencia' }
            if ($elOc) { Set-SeFieldValue $elOc $occ }
        }

        $updates = @($Payload.updates)
        if ($updates.Count -gt 0) {
            try {
                $rows = $Driver.FindElements([OpenQA.Selenium.By]::CssSelector('table tbody tr'))
            } catch {
                $rows = @()
            }
            if ($rows -and $rows.Count -gt 0) {
                $max = [Math]::Min($updates.Count, $rows.Count)
                for ($ri = 0; $ri -lt $max; $ri++) {
                    $row = $rows[$ri]
                    try {
                        $fields = $row.FindElements([OpenQA.Selenium.By]::CssSelector('input, textarea'))
                    } catch { $fields = @() }
                    if (-not $fields -or $fields.Count -eq 0) { continue }
                    $u = $updates[$ri]
                    $ud = ''
                    $uh = ''
                    $tx = ''
                    if ($null -ne $u.updateDate) { $ud = [string]$u.updateDate } elseif ($null -ne $u.date) { $ud = [string]$u.date }
                    if ($null -ne $u.updateHour) { $uh = [string]$u.updateHour } elseif ($null -ne $u.hour) { $uh = [string]$u.hour }
                    if ($null -ne $u.text) { $tx = [string]$u.text }
                    if ($fields.Count -ge 3) {
                        Set-SeFieldValue $fields[0] $ud
                        Set-SeFieldValue $fields[1] $uh
                        Set-SeFieldValue $fields[2] $tx
                    } elseif ($fields.Count -eq 2) {
                        Set-SeFieldValue $fields[0] $ud
                        Set-SeFieldValue $fields[1] $tx
                    } elseif ($fields.Count -eq 1) {
                        Set-SeFieldValue $fields[0] $tx
                    }
                }
            } else {
                Write-Info "Selenium: tabela de atualizacoes nao encontrada (CSS table tbody tr); campos principais preenchidos."
            }
        }

        Write-OK "Selenium: formulario preenchido na janela do WebDriver. Faca login no Hub se necessario e grave na UI."
    } catch {
        Write-Warn ("Selenium: " + $_)
    } finally {
        if ($null -ne $Driver) {
            Write-Host "  Encerrar o WebDriver (fecha o browser automatizado)? [S/n]: " -ForegroundColor Yellow -NoNewline
            $q = Read-Host
            if ($q -eq '' -or $q -match '^[Ss]') {
                try {
                    if (Get-Command Stop-SeDriver -ErrorAction SilentlyContinue) {
                        Stop-SeDriver -Target $Driver
                    } else {
                        $Driver.Quit()
                    }
                } catch {
                    Write-Warn ("Nao foi possivel encerrar o WebDriver: " + $_)
                }
            }
        }
    }
}

function Invoke-HubMaybeSeleniumFormFill {
    param([hashtable]$Cfg, [string]$HubUrl, [string]$EncPathRel, $Payload)
    if (-not $Cfg.HubWebDriverEnabled) { return }
    if (-not (Test-HubSeleniumModuleAvailable)) {
        Write-Warn "HubWebDriverEnabled=true mas o modulo Selenium nao esta instalado (Install-Module Selenium -Scope CurrentUser)."
        return
    }
    Write-Host ""
    Write-Host "  Preencher o Gerador CCO via Selenium (nova janela de browser)? [s/N]: " -ForegroundColor Yellow -NoNewline
    $a = Read-Host
    if ($a -notmatch '^[Ss]') { return }
    $page = Get-HubEncaminharUri $HubUrl $EncPathRel
    $br = if ($Cfg.HubWebDriverBrowser -and $Cfg.HubWebDriverBrowser.ToString().Trim()) {
        $Cfg.HubWebDriverBrowser.ToString().Trim()
    } else { 'Chrome' }
    Invoke-HubRelatorioSeleniumFill -HubPageUrl $page -Payload $Payload -Browser $br
}

function Invoke-HubPerguntaAbrirEncaminhar {
    param([string]$HubUrl, [string]$EncaminharPath)
    Write-Host "  Abrir Hub no navegador (pagina Gerador CCO / encaminhar)? [S/n]: " -ForegroundColor Yellow -NoNewline
    $r2 = Read-Host
    if ($r2 -eq '' -or $r2 -match '^[Ss]') {
        try {
            $u = Get-HubEncaminharUri $HubUrl $EncaminharPath
            Start-Process -FilePath $u
        } catch {
            Write-Warn ("Nao foi possivel abrir o navegador: " + $_)
        }
    }
}

function Invoke-SyncHub {
    param([hashtable]$Cfg, [string]$ExportScript)

    Write-Banner
    Write-Centered "-- SINCRONIZAR COM HUB --" 'White'
    Write-Host ""
    Write-Info "Fluxo: resumo no terminal, ocorrencia opcional, HTML de revisao + guia consola; opcionalmente Selenium (WebDriver) para preencher o Gerador; depois confirmacao e API."
    Write-Info "Com HubWebDriverEnabled=true e modulo Selenium instalado, pode abrir-se uma janela de browser automatizada no Gerador CCO."
    Write-Info "A sincronizacao por API (GET/POST/PUT em .../tickets) grava no servidor; depois pode abrir o Hub na rota configurada (padrao api/relatorio)."
    Write-Host ""

    # Credenciais do Hub (opcional: HubEmail / HubPassword em config.json)
    $hubDefault = if ($Cfg.HubBaseURL) { $Cfg.HubBaseURL } else { 'http://172.16.0.49:3210' }
    $hubUrl   = Read-Field "URL do Hub" $hubDefault
    $defEmail = if ($Cfg.HubEmail) { [string]$Cfg.HubEmail } else { '' }
    $hubEmail = Read-Field "Email Hub" $defEmail
    if (-not ($hubEmail -as [string]).Trim()) {
        Write-Err "Email Hub obrigatorio. Configure em Configuracoes (opcao 5) ou informe aqui."
        Pause-Screen; return
    }
    $storedPass = if ($Cfg.HubPassword) { [string]$Cfg.HubPassword } else { '' }
    if ($storedPass.Length -gt 0) {
        Write-Host "  Senha Hub [Enter = usar senha salva no config.json]: " -ForegroundColor Yellow -NoNewline
    } else {
        Write-Host "  Senha Hub: " -ForegroundColor Yellow -NoNewline
    }
    $hubPassIn = Read-Host
    if ($hubPassIn) {
        $hubPass = $hubPassIn
    } elseif ($storedPass.Length -gt 0) {
        $hubPass = $storedPass
    } else {
        Write-Err "Senha Hub obrigatoria. Salve HubPassword em config.json (opcao 5 ou 6) ou digite aqui."
        Pause-Screen; return
    }
    $hubUrl = $hubUrl.TrimEnd('/')
    $encPathRel = if ($Cfg.HubEncaminharPath -and $Cfg.HubEncaminharPath.ToString().Trim()) {
        $Cfg.HubEncaminharPath.ToString().Trim()
    } else { 'api/relatorio' }
    $apiRelRoot = Get-HubRelatorioPathPrefix $Cfg
    Write-Host ""
    Write-Info ("API lista tickets (GET):  " + $hubUrl + $apiRelRoot + "/tickets")
    Write-Info ("API novo ticket (POST):   tenta " + $apiRelRoot.TrimEnd('/') + "/tickets primeiro, depois rotas legadas")
    Write-Info ("API atualizacao (PUT):    " + $hubUrl + $apiRelRoot + "/tickets/{numeroTicket} (e fallbacks)")
    if (-not $storedPass) {
        Write-Info "Dica: defina HubEmail e HubPassword em config.json (menu 5) para login automatico."
    }
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
        $hubTickets = Get-HubRelatorioTicketsList $hubUrl $apiRelRoot
        foreach ($ht in $hubTickets) {
            if ($ht.number) { $existingMap[$ht.number.ToString()] = $ht }
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

        $basePayload = Build-HubTicket $ticket

        if ($existingMap.ContainsKey($num)) {
            $hub = $existingMap[$num]
            if (Test-HubRelatorioChanged $hub $basePayload) {
                Write-Host ""
                Write-Host ("  [UPD] #" + $num + " - diferenca em relacao ao Hub (dados OTRS ou ocorrencia no payload)") -ForegroundColor Cyan
                Show-HubTicketPreview $basePayload
                $occ = Read-OcorrenciaMultiline
                $payload = Build-HubTicket $ticket -Ocorrencia $occ
                Show-HubTicketPreview $payload
                Write-Info "Abrindo formulario HTML de revisao no navegador..."
                try { Show-HubRelatorioFormHtml $payload $hubUrl $Cfg } catch { Write-Warn ("Falha ao abrir HTML: " + $_) }
                Invoke-HubMaybeSeleniumFormFill $Cfg $hubUrl $encPathRel $payload
                Write-Host ""
                Write-Info "Valide os dados no navegador; em seguida confirme o envio abaixo."
                $null = Read-Host "  Pressione Enter para continuar"
                Write-Host "  Atualizar registro no Hub? [S/n]: " -ForegroundColor Yellow -NoNewline
                $r = Read-Host
                if ($r -eq '' -or $r -match '^[Ss]') {
                    try {
                        $putPaths = @(Get-HubPutTicketCandidatePaths $apiRelRoot $num $Cfg)
                        Hub-PutWithPathFallback $hubUrl $putPaths $payload | Out-Null
                        Write-OK ("Atualizado: #" + $num)
                        $nUpd++
                        Invoke-HubPerguntaAbrirEncaminhar $hubUrl $encPathRel
                    } catch {
                        Write-Warn ("Falha: " + (Get-HubHttpErrorDetail $_))
                        Invoke-HubManualPayloadAssist $hubUrl $encPathRel $payload $num $Cfg
                    }
                } else {
                    Write-Info ("Pulado: #" + $num) ; $nSkip++
                }
            } else {
                Write-Host ("  [=] #" + $num + " sem alteracoes") -ForegroundColor DarkGray
                $nSkip++
            }
        } else {
            Write-Host ""
            Write-ThinDiv 'Yellow'
            Write-Host "  NOVO TICKET:" -ForegroundColor Yellow
            Show-HubTicketPreview $basePayload
            $occ = Read-OcorrenciaMultiline
            $payload = Build-HubTicket $ticket -Ocorrencia $occ
            Show-HubTicketPreview $payload
            Write-Info "Abrindo formulario HTML de revisao no navegador..."
            try { Show-HubRelatorioFormHtml $payload $hubUrl $Cfg } catch { Write-Warn ("Falha ao abrir HTML: " + $_) }
            Invoke-HubMaybeSeleniumFormFill $Cfg $hubUrl $encPathRel $payload
            Write-Host ""
            Write-Info "Valide os dados no navegador; em seguida confirme o envio abaixo."
            $null = Read-Host "  Pressione Enter para continuar"
            Write-Host "  Adicionar ao Hub? [S/n]: " -ForegroundColor Yellow -NoNewline
            $r = Read-Host
            if ($r -eq '' -or $r -match '^[Ss]') {
                try {
                    $postPaths = @(Get-HubPostTicketCandidatePaths $apiRelRoot $Cfg)
                    Hub-PostWithPathFallback $hubUrl $postPaths $payload | Out-Null
                    Write-OK ("Adicionado: #" + $num)
                    $nNew++
                    Invoke-HubPerguntaAbrirEncaminhar $hubUrl $encPathRel
                } catch {
                    Write-Warn ("Falha: " + (Get-HubHttpErrorDetail $_))
                    Invoke-HubManualPayloadAssist $hubUrl $encPathRel $payload $num $Cfg
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
    if (-not $Cfg.HubEncaminharPath) { $Cfg.HubEncaminharPath = 'api/relatorio' }
    if (-not $Cfg.HubApiRelatorioPath) { $Cfg.HubApiRelatorioPath = 'api/relatorio' }
    $Cfg.HubBaseURL = Read-Field "URL base do Hub (relatorio CCO)" $Cfg.HubBaseURL
    $Cfg.HubEncaminharPath = Read-Field "Hub: rota apos envio (ex.: api/relatorio = Gerador CCO, ou home)" $Cfg.HubEncaminharPath
    $Cfg.HubApiRelatorioPath = Read-Field "Hub: prefixo API relatorio (ex.: api/relatorio; GET lista em .../tickets)" $Cfg.HubApiRelatorioPath
    $Cfg.HubEmail = Read-Field "Hub: email (login API /api/login)" ([string]$Cfg.HubEmail)
    $hpHint = if ($Cfg.HubPassword -and $Cfg.HubPassword.ToString().Length -gt 0) { "[senha gravada; Enter mantem]" } else { "[sem senha salva]" }
    Write-Host ("  Hub: senha API " + $hpHint + ": ") -ForegroundColor Yellow -NoNewline
    $hpNew = Read-Host
    if ($hpNew) { $Cfg.HubPassword = $hpNew }
    if (-not $Cfg.HubPostTicketPaths) { $Cfg.HubPostTicketPaths = '' }
    if (-not $Cfg.HubPutTicketPaths) { $Cfg.HubPutTicketPaths = '' }
    Write-Info "Avancado: se o POST/PUT automatico falhar, defina rotas (separadas por ;). Deixe vazio para tentativas automaticas."
    $Cfg.HubPostTicketPaths = Read-Field "Hub POST (vazio=auto; principal: api/relatorio/tickets; legado: api/relatorio/ticket)" ([string]$Cfg.HubPostTicketPaths)
    $Cfg.HubPutTicketPaths  = Read-Field "Hub PUT (vazio=auto; {numero} = n. OTRS; ex.: api/relatorio/tickets/{numero})" ([string]$Cfg.HubPutTicketPaths)
    $wdDef = if ($Cfg.HubWebDriverEnabled) { 's' } else { 'n' }
    $wdIn = Read-Field "Hub: oferecer preenchimento via Selenium WebDriver na sync (s/N)" $wdDef
    $Cfg.HubWebDriverEnabled = ($wdIn -match '^[Ss]')
    $Cfg.HubWebDriverBrowser = Read-Field "Hub WebDriver: browser (Chrome ou Edge)" ([string]$Cfg.HubWebDriverBrowser)
    if (-not ($Cfg.HubWebDriverBrowser -as [string]).Trim()) { $Cfg.HubWebDriverBrowser = 'Chrome' }
    Write-Info "Selenium: Install-Module Selenium -Scope CurrentUser (API classica Start-SeChrome). Ver docs/automacao-formulario-hub.md."
    Write-Info "Opcional: HubFormSelectors no config.json (JSON) afinar selectores ao colar o script na consola do Hub — ver README."
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

        Write-MenuOpt '1' 'Gerar Relatorio TXT'     'White'    'TXT + Relatorio_CCO_*.json; pergunta todas as notas ou 5 no export'
        Write-MenuOpt '2' 'Gerar Relatorio JSON'    'White'    'Mesmo export (TXT+cache+JSON); pergunta todas as notas ou 5 no export'
        Write-MenuOpt '3' 'Visualizar Chamados'     'Yellow'   'OTRS em tempo real ou cache local; alerta de normalizacao'
        Write-Host ""
        Write-MenuOpt '4' 'Alterar Credenciais'     'Cyan'     'Usuario, senha e URL'
        Write-MenuOpt '5' 'Configuracoes'           'Cyan'     'Busca, cache, Hub (URL, API, email/senha API)'
        Write-MenuOpt '6' 'Salvar Credenciais'      'DarkGray' ''
        Write-MenuOpt '7' 'Sincronizar com Hub'   'Magenta'  'API + HTML; Selenium opcional; JSON+ajuda se API falhar'
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