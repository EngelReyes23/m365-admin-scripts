#Requires -Version 7.4
<#
.SYNOPSIS
    Microsoft 365 Recycle Bin and Hold Cleaner

.DESCRIPTION
    Herramienta interactiva para analizar y limpiar la papelera de reciclaje
    de SharePoint Online y OneDrive for Business.

    Incluye:
      - TUI moderna con menús numéricos
      - SharePoint Online y OneDrive for Business
      - Búsqueda dinámica de sitios desde el tenant
      - URL directa como alternativa
      - Gestión de App Registration de Microsoft Entra ID
      - Validación real de Client ID / tenant / acceso PnP
      - Conexiones PnP explícitas
      - Modo simulación por defecto
      - Primer y segundo nivel de papelera
      - Preservation Hold Library con detección dinámica
      - Confirmación reforzada para operaciones destructivas
      - Throttling, Retry-After, backoff exponencial y jitter
      - Reportes CSV de auditoría
      - Configuración persistente

.REQUIREMENTS
    PowerShell 7.4+
    PnP.PowerShell 3.2+

.NOTES
    Get-PnPRecycleBinItem requiere que la identidad conectada sea
    Site Collection Administrator del sitio de destino.
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'


# =============================================================================
# APLICACION
# =============================================================================

$script:AppName           = 'Microsoft 365 Recycle Bin and Hold Cleaner'
$script:AppVersion        = '2.1.0'
$script:MinimumPnPVersion = [version]'3.2.0'

if ($env:LOCALAPPDATA) {
    $script:AppRoot = Join-Path $env:LOCALAPPDATA 'M365RecycleBinHoldCleaner'
}
else {
    $script:AppRoot = Join-Path $HOME '.M365RecycleBinHoldCleaner'
}

$script:ConfigPath  = Join-Path $script:AppRoot 'config.json'
$script:ReportsPath = Join-Path $script:AppRoot 'Reportes'

$script:Settings       = $null
$script:Target         = $null
$script:PnPConnection  = $null
$script:Ansi           = $false
$script:ThrottleEvents = [System.Collections.Generic.List[object]]::new()

# =============================================================================
# LANGUAGE
# =============================================================================

$script:Language = 'es'
$script:UiTranslations = @(
    @{ From = 'Inicio'; To = 'Home' }; @{ From = 'Finalizado'; To = 'Finished' }; @{ From = 'Error inesperado'; To = 'Unexpected error' }
    @{ From = 'Detalles técnicos:'; To = 'Technical details:' }; @{ From = 'Configuración'; To = 'Settings' }; @{ From = 'Salir'; To = 'Exit' }
    @{ From = 'Cancelar'; To = 'Cancel' }; @{ From = 'Volver'; To = 'Back' }; @{ From = 'Diagnóstico'; To = 'Diagnostics' }
    @{ From = 'Autenticación'; To = 'Authentication' }; @{ From = 'Selecciona una opción'; To = 'Select an option' }
    @{ From = 'Número'; To = 'Number' }; @{ From = 'Bibliotecas'; To = 'Libraries' }; @{ From = 'Biblioteca'; To = 'Library' }
    @{ From = 'Sitio'; To = 'Site' }; @{ From = 'Modo SIMULACIÓN'; To = 'SIMULATION mode' }; @{ From = 'SIMULACIÓN'; To = 'SIMULATION' }
    @{ From = 'REAL'; To = 'LIVE' }; @{ From = 'Enter para continuar'; To = 'Press Enter to continue' }
    @{ From = 'Vaciar primer nivel'; To = 'Empty first stage' }; @{ From = 'Vaciar segundo nivel'; To = 'Empty second stage' }
    @{ From = 'Vaciar ambos niveles'; To = 'Empty both stages' }; @{ From = 'Procesar Preservation Hold Library'; To = 'Process Preservation Hold Library' }
    @{ From = 'Confirmación'; To = 'Confirmation' }; @{ From = 'Sesión finalizada.'; To = 'Session finished.' }
)
function Get-LocalizedText {
    param([AllowNull()][object]$Text)
    if ($null -eq $Text) { return '' }; $result = [string]$Text
    if ($script:Language -eq 'en') { foreach ($translation in $script:UiTranslations) { $result = $result.Replace($translation.From, $translation.To) } }
    return $result
}
function Write-Host {
    [CmdletBinding()] param([Parameter(Position=0,ValueFromRemainingArguments=$true)][object[]]$Object,[ConsoleColor]$ForegroundColor,[ConsoleColor]$BackgroundColor,[switch]$NoNewline)
    $text = if ($null -eq $Object) { '' } else { ($Object | ForEach-Object { [string]$_ }) -join ' ' }
    $parameters = @{ Object = (Get-LocalizedText $text) }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters.ForegroundColor = $ForegroundColor }; if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $parameters.BackgroundColor = $BackgroundColor }; if ($NoNewline) { $parameters.NoNewline = $true }
    Microsoft.PowerShell.Utility\Write-Host @parameters
}
function Read-Host {
    param([Parameter(Position=0)][string]$Prompt,[switch]$AsSecureString)
    $localizedPrompt = Get-LocalizedText $Prompt
    if ($AsSecureString) { return Microsoft.PowerShell.Utility\Read-Host -Prompt $localizedPrompt -AsSecureString }
    return Microsoft.PowerShell.Utility\Read-Host -Prompt $localizedPrompt
}
function Initialize-AppLanguage {
    while ($true) {
        Microsoft.PowerShell.Utility\Write-Host ''; Microsoft.PowerShell.Utility\Write-Host 'Select language / Seleccione idioma:'; Microsoft.PowerShell.Utility\Write-Host '1. English'; Microsoft.PowerShell.Utility\Write-Host '2. Español'
        $choice = (Microsoft.PowerShell.Utility\Read-Host 'Choice / Opción').Trim()
        if ($choice -eq '1') { $script:Language = 'en'; return }; if ($choice -eq '2') { $script:Language = 'es'; return }
        Microsoft.PowerShell.Utility\Write-Host 'Please choose 1 or 2 / Elija 1 o 2.'
    }
}


# =============================================================================
# TUI MODERNA
# =============================================================================

function Initialize-AppTerminal {
    $script:Ansi = $false

    try {
        if ($Host.UI -and $Host.UI.SupportsVirtualTerminal) {
            $script:Ansi = $true
        }
        elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -or $env:TERM) {
            $script:Ansi = $true
        }
    }
    catch {
        $script:Ansi = $false
    }
}

function Get-AppAnsi {
    param(
        [ValidateSet('Reset','Bold','Dim','Cyan','Green','Yellow','Red','Magenta')]
        [string]$Name
    )

    if (-not $script:Ansi) {
        return ''
    }

    $esc = [char]27

    switch ($Name) {
        'Reset'   { "$esc[0m" }
        'Bold'    { "$esc[1m" }
        'Dim'     { "$esc[2m" }
        'Cyan'    { "$esc[36m" }
        'Green'   { "$esc[92m" }
        'Yellow'  { "$esc[93m" }
        'Red'     { "$esc[91m" }
        'Magenta' { "$esc[95m" }
    }
}

function Write-AppStyled {
    param(
        [Parameter(Mandatory)]
        [string]$Text,

        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal',

        [switch]$NoNewline
    )

    $prefix   = ''
    $fallback = 'Gray'

    switch ($Style) {
        'Muted'   { $prefix = Get-AppAnsi Dim; $fallback = 'DarkGray' }
        'Primary' { $prefix = (Get-AppAnsi Bold) + (Get-AppAnsi Cyan); $fallback = 'Cyan' }
        'Success' { $prefix = Get-AppAnsi Green; $fallback = 'Green' }
        'Warning' { $prefix = Get-AppAnsi Yellow; $fallback = 'Yellow' }
        'Danger'  { $prefix = Get-AppAnsi Red; $fallback = 'Red' }
        'Accent'  { $prefix = Get-AppAnsi Magenta; $fallback = 'Magenta' }
    }

    if ($script:Ansi) {
        $suffix = Get-AppAnsi Reset

        if ($NoNewline) {
            Write-Host "$prefix$Text$suffix" -NoNewline
        }
        else {
            Write-Host "$prefix$Text$suffix"
        }
    }
    else {
        if ($NoNewline) {
            Write-Host $Text -ForegroundColor $fallback -NoNewline
        }
        else {
            Write-Host $Text -ForegroundColor $fallback
        }
    }
}

function Clear-AppScreen {
    Clear-Host
}

function Show-AppHeader {
    param(
        [string]$Section = 'Inicio'
    )

    $width = 72

    try {
        $width = [Math]::Min(
            76,
            [Math]::Max(30, $Host.UI.RawUI.WindowSize.Width - 2)
        )
    }
    catch {
        $width = 72
    }

    Clear-AppScreen

    Write-AppStyled -Text "$script:AppName  v$script:AppVersion" -Style Primary

    if (-not [string]::IsNullOrWhiteSpace($Section) -and $Section -ne 'Inicio') {
        Write-AppStyled -Text $Section -Style Muted
    }

    Write-AppStyled -Text ('━' * $width) -Style Muted
}

function Write-AppOk {
    param([Parameter(Mandatory)][string]$Message)
    Write-AppStyled -Text "✓ $Message" -Style Success
}

function Write-AppInfo {
    param([Parameter(Mandatory)][string]$Message)
    Write-AppStyled -Text "● $Message" -Style Primary
}

function Write-AppWarning {
    param([Parameter(Mandatory)][string]$Message)
    Write-AppStyled -Text "! $Message" -Style Warning
}

function Write-AppError {
    param([Parameter(Mandatory)][string]$Message)
    Write-AppStyled -Text "× $Message" -Style Danger
}

function Write-AppMuted {
    param([Parameter(Mandatory)][string]$Message)
    Write-AppStyled -Text $Message -Style Muted
}

function Write-AppSection {
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ''
    Write-AppStyled -Text $Title -Style Primary
}

function Write-AppField {
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        $Value = '—'
    }

    Write-AppStyled -Text ('{0,-16}' -f $Name) -Style Muted -NoNewline
    Write-Host $Value
}

function Wait-App {
    param([string]$Message = 'Enter para continuar')

    Write-Host ''
    [void](Read-Host $Message)
}

function Show-AppErrorScreen {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message
    )

    Show-AppHeader -Section $Title
    Write-Host ''
    Write-AppError -Message $Message
    Wait-App
}


# =============================================================================
# ENTRADAS / VALIDACIONES
# =============================================================================

function Test-AppGuid {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $guid = [Guid]::Empty
    return [Guid]::TryParse($Value, [ref]$guid)
}

function Test-AppTenant {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return $Value.Trim() -match '^[A-Za-z0-9-]+\.onmicrosoft\.com$'
}

function Test-AppM365Url {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $uri = $null

    if (-not [Uri]::TryCreate($Value.Trim(), [UriKind]::Absolute, [ref]$uri)) {
        return $false
    }

    if ($uri.Scheme -ne 'https') {
        return $false
    }

    return $uri.Host -match '(?i)\.sharepoint\.com$'
}

function Read-AppText {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = '',
        [switch]$AllowEmpty,
        [scriptblock]$Validator,
        [string]$ValidationMessage = 'Valor inválido.'
    )

    while ($true) {
        $caption = if (-not [string]::IsNullOrWhiteSpace($Default)) {
            "$Prompt [$Default]"
        }
        else {
            $Prompt
        }

        $value = Read-Host $caption

        if ([string]::IsNullOrWhiteSpace($value) -and -not [string]::IsNullOrWhiteSpace($Default)) {
            $value = $Default
        }

        if ([string]::IsNullOrWhiteSpace($value)) {
            if ($AllowEmpty) {
                return ''
            }

            Write-AppWarning -Message 'El valor no puede estar vacío.'
            continue
        }

        $value = $value.Trim()

        if ($Validator -and -not (& $Validator $value)) {
            Write-AppWarning -Message $ValidationMessage
            continue
        }

        return $value
    }
}

function Read-AppYesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$DefaultYes = $false
    )

    $hint = if ($DefaultYes) { 'S/n' } else { 's/N' }

    while ($true) {
        $answer = (Read-Host "$Prompt [$hint]").Trim()

        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $DefaultYes
        }

        if ($answer -match '^(s|si|sí|y|yes)$') {
            return $true
        }

        if ($answer -match '^(n|no)$') {
            return $false
        }

        Write-AppWarning -Message 'Responde S o N.'
    }
}

function Read-AppInteger {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][int]$Minimum,
        [Parameter(Mandatory)][int]$Maximum,
        [Parameter(Mandatory)][int]$DefaultValue
    )

    while ($true) {
        $raw = (Read-Host "$Prompt [$DefaultValue]").Trim()

        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $DefaultValue
        }

        $number = 0

        if (
            [int]::TryParse($raw, [ref]$number) -and
            $number -ge $Minimum -and
            $number -le $Maximum
        ) {
            return $number
        }

        Write-AppWarning -Message "Introduce un número entre $Minimum y $Maximum."
    }
}

function Read-AppMenuChoice {
    param(
        [Parameter(Mandatory)][array]$Items,
        [string]$Prompt = 'Selecciona una opción',
        [switch]$AllowBack
    )

    $normalized = @(
        foreach ($item in $Items) {
            $key = ''
            $label = ''
            $description = ''
            $value = ''

            if ($item -is [System.Collections.IDictionary]) {
                if ($item.Contains('Key'))         { $key = [string]$item['Key'] }
                if ($item.Contains('Label'))       { $label = [string]$item['Label'] }
                if ($item.Contains('Description')) { $description = [string]$item['Description'] }
                if ($item.Contains('Value'))       { $value = [string]$item['Value'] }
            }
            else {
                $keyProperty = $item.PSObject.Properties['Key']
                $labelProperty = $item.PSObject.Properties['Label']
                $descriptionProperty = $item.PSObject.Properties['Description']
                $valueProperty = $item.PSObject.Properties['Value']

                if ($null -ne $keyProperty)         { $key = [string]$keyProperty.Value }
                if ($null -ne $labelProperty)       { $label = [string]$labelProperty.Value }
                if ($null -ne $descriptionProperty) { $description = [string]$descriptionProperty.Value }
                if ($null -ne $valueProperty)       { $value = [string]$valueProperty.Value }
            }

            if ([string]::IsNullOrWhiteSpace($key)) {
                throw 'Cada elemento del menú debe tener Key.'
            }

            if ([string]::IsNullOrWhiteSpace($label)) {
                throw "El elemento '$key' debe tener Label."
            }

            [pscustomobject]@{
                Key         = $key
                Label       = $label
                Description = $description
                Value       = $value
            }
        }
    )

    foreach ($item in $normalized) {
        Write-AppStyled -Text ('{0,3}' -f $item.Key) -Style Accent -NoNewline
        Write-Host "  $($item.Label)"

        if (-not [string]::IsNullOrWhiteSpace($item.Description)) {
            Write-AppMuted -Message "     $($item.Description)"
        }
    }

    $valid = @($normalized | ForEach-Object { $_.Key })

    if ($AllowBack -and -not ($valid -contains '0')) {
        Write-AppStyled -Text '  0' -Style Accent -NoNewline
        Write-Host '  Volver'
        $valid += '0'
    }

    while ($true) {
        Write-Host ''
        $choice = (Read-Host $Prompt).Trim()

        if ($valid -contains $choice) {
            if ($choice -eq '0' -and -not ($normalized.Key -contains '0')) {
                return [pscustomobject]@{ Key = '0'; Label = 'Volver'; Description = ''; Value = 'Back' }
            }

            return $normalized | Where-Object { $_.Key -eq $choice } | Select-Object -First 1
        }

        Write-AppError -Message 'Opción inválida.'
    }
}

function Select-AppSingleByNumber {
    param(
        [Parameter(Mandatory)][array]$Items,
        [Parameter(Mandatory)][scriptblock]$Label,
        [string]$Title = 'Seleccionar',
        [int]$MaxDisplay = 50
    )

    $list = @($Items)

    if ($list.Count -eq 0) {
        return $null
    }

    if ($list.Count -gt $MaxDisplay) {
        Write-AppWarning -Message "Hay $($list.Count) resultados. Se muestran los primeros $MaxDisplay."
        $list = @($list | Select-Object -First $MaxDisplay)
    }

    Write-AppSection -Title $Title

    for ($i = 0; $i -lt $list.Count; $i++) {
        Write-AppStyled -Text ('{0,3}' -f ($i + 1)) -Style Accent -NoNewline
        Write-Host ('  ' + (& $Label $list[$i]))
    }

    Write-AppStyled -Text '  0' -Style Accent -NoNewline
    Write-Host '  Cancelar'

    while ($true) {
        Write-Host ''
        $raw = Read-Host 'Número'
        $number = 0

        if ([int]::TryParse($raw, [ref]$number)) {
            if ($number -eq 0) {
                return $null
            }

            if ($number -ge 1 -and $number -le $list.Count) {
                return $list[$number - 1]
            }
        }

        Write-AppError -Message 'Selección inválida.'
    }
}


# =============================================================================
# CONFIGURACION
# =============================================================================

function Initialize-AppFolders {
    if (-not (Test-Path -LiteralPath $script:AppRoot)) {
        [void](New-Item -Path $script:AppRoot -ItemType Directory -Force)
    }

    if (-not (Test-Path -LiteralPath $script:ReportsPath)) {
        [void](New-Item -Path $script:ReportsPath -ItemType Directory -Force)
    }
}

function Get-DefaultAppSettings {
    $clientId = ''

    if ($env:ENTRAID_APP_ID) {
        $clientId = $env:ENTRAID_APP_ID
    }
    elseif ($env:ENTRAID_CLIENT_ID) {
        $clientId = $env:ENTRAID_CLIENT_ID
    }

    return [pscustomobject]@{
        Tenant              = ''
        ClientId            = $clientId
        AppRegistrationName = 'Microsoft 365 Recycle Bin and Hold Cleaner'
        PersistLogin        = $true
        Simulation          = $true
        DefaultLimit        = 1000
        MaximumLimit        = 10000
        RequestDelayMs      = 150
        MaxRetries          = 8
        RetryBaseSeconds    = 5
        RetryMaxSeconds     = 120
        AdminUpn            = ''
    }
}

function Merge-AppSettings {
    param([object]$Loaded)

    $defaults = Get-DefaultAppSettings

    if ($null -eq $Loaded) {
        return $defaults
    }

    foreach ($property in $defaults.PSObject.Properties) {
        if ($null -eq $Loaded.PSObject.Properties[$property.Name]) {
            Add-Member -InputObject $Loaded -NotePropertyName $property.Name -NotePropertyValue $property.Value
        }
    }

    return $Loaded
}

function Get-AppSettings {
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        return Get-DefaultAppSettings
    }

    try {
        $loaded = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        return Merge-AppSettings -Loaded $loaded
    }
    catch {
        return Get-DefaultAppSettings
    }
}

function Save-AppSettings {
    Initialize-AppFolders

    $script:Settings |
        ConvertTo-Json -Depth 8 |
        Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}


# =============================================================================
# PNP.POWERSHELL
# =============================================================================

function Test-AppPowerShellVersion {
    if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
        Write-AppError -Message (
            "PowerShell $($PSVersionTable.PSVersion) detectado. " +
            'Se requiere PowerShell 7.4 o superior.'
        )
        return $false
    }

    return $true
}

function Get-AppPnPInstalledVersion {
    $module = Get-Module -ListAvailable -Name PnP.PowerShell |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -ne $module) {
        return $module.Version
    }

    return $null
}

function Confirm-AppPnP {
    if (-not (Test-AppPowerShellVersion)) {
        return $false
    }

    $version = Get-AppPnPInstalledVersion

    if ($null -ne $version -and [version]$version -ge $script:MinimumPnPVersion) {
        try {
            Import-Module PnP.PowerShell -MinimumVersion $script:MinimumPnPVersion -ErrorAction Stop
            return $true
        }
        catch {
            Show-AppErrorScreen -Title 'PnP.PowerShell' -Message $_.Exception.Message
            return $false
        }
    }

    Show-AppHeader -Section 'PnP.PowerShell'

    if ($null -eq $version) {
        Write-AppWarning -Message 'PnP.PowerShell no está instalado.'
    }
    else {
        Write-AppWarning -Message "PnP.PowerShell $version es anterior a $script:MinimumPnPVersion."
    }

    Write-Host ''

    if (-not (Read-AppYesNo -Prompt '¿Instalar/actualizar ahora?' -DefaultYes $true)) {
        return $false
    }

    try {
        Install-Module PnP.PowerShell `
            -Scope CurrentUser `
            -Force `
            -AllowClobber `
            -MinimumVersion $script:MinimumPnPVersion `
            -ErrorAction Stop

        Import-Module PnP.PowerShell `
            -MinimumVersion $script:MinimumPnPVersion `
            -Force `
            -ErrorAction Stop

        Write-AppOk -Message 'PnP.PowerShell está listo.'
        Wait-App
        return $true
    }
    catch {
        Show-AppErrorScreen -Title 'PnP.PowerShell' -Message $_.Exception.Message
        return $false
    }
}

function Get-AppTenantPrefix {
    if (-not (Test-AppTenant -Value $script:Settings.Tenant)) {
        return $null
    }

    return ($script:Settings.Tenant -split '\.')[0]
}

function Get-AppAdminUrl {
    $prefix = Get-AppTenantPrefix

    if ([string]::IsNullOrWhiteSpace($prefix)) {
        return $null
    }

    return "https://$prefix-admin.sharepoint.com"
}

function Connect-AppUrl {
    param(
        [Parameter(Mandatory)][string]$Url,
        [switch]$SkipEnsureAuth
    )

    if (-not (Confirm-AppPnP)) {
        throw 'PnP.PowerShell no está disponible.'
    }

    if (-not $SkipEnsureAuth) {
        if (-not (Ensure-AppAuthentication)) {
            throw 'No existe una autenticación válida configurada.'
        }
    }

    $params = @{
        Url              = $Url
        ClientId         = $script:Settings.ClientId
        Tenant           = $script:Settings.Tenant
        Interactive      = $true
        ReturnConnection = $true
        ErrorAction      = 'Stop'
    }

    if ([bool]$script:Settings.PersistLogin) {
        $params.PersistLogin = $true
    }

    return Connect-PnPOnline @params
}

function Disconnect-AppM365 {
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
    }

    $script:PnPConnection = $null
}


# =============================================================================
# THROTTLING / RETRIES
# =============================================================================

function Test-AppThrottleException {
    param([Parameter(Mandatory)][System.Exception]$Exception)

    $text = $Exception.ToString()

    return $text -match (
        '(?i)' +
        '\b429\b|' +
        'Too Many Requests|' +
        'throttl|' +
        '\b503\b|' +
        'Server Too Busy|' +
        'Server is busy|' +
        'Service Unavailable|' +
        'temporarily unavailable|' +
        'timeout|' +
        'timed out'
    )
}

function Get-AppRetryAfterSeconds {
    param([Parameter(Mandatory)][System.Exception]$Exception)

    $current = $Exception

    while ($null -ne $current) {
        try {
            $responseProperty = $current.PSObject.Properties['Response']

            if ($null -ne $responseProperty -and $null -ne $responseProperty.Value) {
                $response = $responseProperty.Value
                $headersProperty = $response.PSObject.Properties['Headers']

                if ($null -ne $headersProperty -and $null -ne $headersProperty.Value) {
                    $headers = $headersProperty.Value
                    $retryAfter = $headers['Retry-After']

                    if ($null -ne $retryAfter) {
                        $seconds = 0

                        if ([int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -gt 0) {
                            return $seconds
                        }
                    }
                }
            }
        }
        catch {
        }

        $current = $current.InnerException
    }

    return $null
}

function Get-AppRetryDelay {
    param(
        [Parameter(Mandatory)][int]$Attempt,
        [Parameter(Mandatory)][System.Exception]$Exception
    )

    $retryAfter = Get-AppRetryAfterSeconds -Exception $Exception

    if ($null -ne $retryAfter) {
        return [Math]::Min([int]$retryAfter, [int]$script:Settings.RetryMaxSeconds)
    }

    $base = [int]$script:Settings.RetryBaseSeconds
    $max  = [int]$script:Settings.RetryMaxSeconds

    $exponential = $base * [Math]::Pow(2, $Attempt - 1)
    $jitter = Get-Random -Minimum 0 -Maximum 4

    return [Math]::Min([int][Math]::Ceiling($exponential + $jitter), $max)
}

function Invoke-AppWithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [Parameter(Mandatory)][string]$OperationName,
        [switch]$NoPacing
    )

    $maxRetries = [int]$script:Settings.MaxRetries

    for ($attempt = 1; $attempt -le ($maxRetries + 1); $attempt++) {
        if (-not $NoPacing) {
            $delayMs = [int]$script:Settings.RequestDelayMs

            if ($delayMs -gt 0) {
                Start-Sleep -Milliseconds $delayMs
            }
        }

        try {
            return & $Operation
        }
        catch {
            $exception = $_.Exception

            if (-not (Test-AppThrottleException -Exception $exception)) {
                throw
            }

            if ($attempt -gt $maxRetries) {
                throw
            }

            $waitSeconds = Get-AppRetryDelay -Attempt $attempt -Exception $exception

            [void]$script:ThrottleEvents.Add(
                [pscustomobject]@{
                    FechaHora       = Get-Date
                    Operacion       = $OperationName
                    Intento         = $attempt
                    EsperaSegundos  = $waitSeconds
                    Mensaje         = $exception.Message
                }
            )

            Write-AppWarning -Message 'SharePoint/OneDrive aplicó throttling.'
            Write-AppMuted -Message "${OperationName}: reintento $attempt/$maxRetries en ${waitSeconds}s."

            Start-Sleep -Seconds $waitSeconds
        }
    }
}


# =============================================================================
# AUTENTICACION / APP REGISTRATION
# =============================================================================

function Ensure-AppTenant {
    if (Test-AppTenant -Value $script:Settings.Tenant) {
        return $true
    }

    Show-AppHeader -Section 'Tenant'

    $script:Settings.Tenant = Read-AppText `
        -Prompt 'Dominio inicial del tenant' `
        -Validator { param($v) Test-AppTenant -Value $v } `
        -ValidationMessage 'Usa el dominio inicial, por ejemplo empresa.onmicrosoft.com.'

    Save-AppSettings
    return $true
}

function Test-AppConfiguredApplication {
    param(
        [switch]$InteractiveResult
    )

    if (-not (Test-AppTenant -Value $script:Settings.Tenant)) {
        if ($InteractiveResult) {
            Write-AppError -Message 'Tenant no configurado o inválido.'
        }
        return $false
    }

    if (-not (Test-AppGuid -Value $script:Settings.ClientId)) {
        if ($InteractiveResult) {
            Write-AppError -Message 'Client ID no configurado o inválido.'
        }
        return $false
    }

    try {
        $adminUrl = Get-AppAdminUrl

        if ([string]::IsNullOrWhiteSpace($adminUrl)) {
            throw 'No fue posible determinar la URL del centro de administración.'
        }

        if ($InteractiveResult) {
            Write-AppInfo -Message 'Validando autenticación contra el tenant...'
        }

        $connection = Connect-AppUrl -Url $adminUrl -SkipEnsureAuth

        $null = Get-PnPWeb -Connection $connection -ErrorAction Stop

        # Además valida que esta identidad pueda enumerar sitios,
        # necesario para la búsqueda dinámica del toolkit.
        $null = Get-PnPTenantSite `
            -Connection $connection `
            -Detailed `
            -ErrorAction Stop |
                Select-Object -First 1

        if ($InteractiveResult) {
            Write-AppOk -Message 'Aplicación, tenant y acceso PnP validados correctamente.'
        }

        return $true
    }
    catch {
        if ($InteractiveResult) {
            Write-AppError -Message 'La validación falló.'
            Write-AppMuted -Message $_.Exception.Message
        }
        return $false
    }
}

function Set-AppExistingApplication {
    Show-AppHeader -Section 'Usar aplicación existente'

    if (-not (Ensure-AppTenant)) {
        return
    }

    $oldClientId = [string]$script:Settings.ClientId

    $clientId = Read-AppText `
        -Prompt 'Client ID de la aplicación existente' `
        -Validator { param($v) Test-AppGuid -Value $v } `
        -ValidationMessage 'Debe ser un GUID válido.'

    $script:Settings.ClientId = $clientId

    Write-Host ''

    if (Read-AppYesNo -Prompt '¿Validar antes de guardar?' -DefaultYes $true) {
        if (-not (Test-AppConfiguredApplication -InteractiveResult)) {
            $script:Settings.ClientId = $oldClientId
            Write-AppWarning -Message 'Se conservó el Client ID anterior.'
            Wait-App
            return
        }
    }

    Save-AppSettings

    Write-AppOk -Message 'Aplicación configurada.'
    Write-AppMuted -Message $clientId
    Wait-App
}

function New-AppEntraRegistration {
    Show-AppHeader -Section 'Registrar nueva aplicación Entra'

    if (-not (Ensure-AppTenant)) {
        return
    }

    Write-AppMuted -Message 'La aplicación solicitará el permiso delegado de SharePoint:'
    Write-AppStyled -Text 'AllSites.FullControl' -Style Warning
    Write-Host ''

    if (-not (Read-AppYesNo -Prompt '¿Continuar con el registro?' -DefaultYes $false)) {
        return
    }

    $name = Read-AppText `
        -Prompt 'Nombre de la aplicación' `
        -Default $script:Settings.AppRegistrationName

    try {
        $result = Register-PnPEntraIDAppForInteractiveLogin `
            -ApplicationName $name `
            -Tenant $script:Settings.Tenant `
            -SharePointDelegatePermissions 'AllSites.FullControl' `
            -ErrorAction Stop

        $clientId = ''

        foreach ($propertyName in @('ClientId','AppId','ApplicationId','Id')) {
            $property = $result.PSObject.Properties[$propertyName]

            if ($null -ne $property -and (Test-AppGuid -Value ([string]$property.Value))) {
                $clientId = [string]$property.Value
                break
            }
        }

        if (-not (Test-AppGuid -Value $clientId)) {
            $clientId = Read-AppText `
                -Prompt 'Client ID creado' `
                -Validator { param($v) Test-AppGuid -Value $v } `
                -ValidationMessage 'Debe ser un GUID válido.'
        }

        $script:Settings.ClientId = $clientId
        $script:Settings.AppRegistrationName = $name
        Save-AppSettings

        Write-AppOk -Message 'Nueva aplicación configurada.'
        Write-AppMuted -Message $clientId
        Write-Host ''

        if (Read-AppYesNo -Prompt '¿Probarla ahora?' -DefaultYes $true) {
            [void](Test-AppConfiguredApplication -InteractiveResult)
        }

        Wait-App
    }
    catch {
        Show-AppErrorScreen -Title 'Registrar aplicación' -Message $_.Exception.Message
    }
}

function Remove-AppLocalClientId {
    Show-AppHeader -Section 'Quitar Client ID local'

    Write-AppWarning -Message 'Esto no elimina la aplicación de Microsoft Entra ID.'
    Write-AppMuted -Message 'Solo elimina el Client ID guardado por esta herramienta.'
    Write-Host ''

    if (Read-AppYesNo -Prompt '¿Continuar?' -DefaultYes $false) {
        $script:Settings.ClientId = ''
        Save-AppSettings
        Write-AppOk -Message 'Client ID eliminado de la configuración local.'
    }

    Wait-App
}

function Clear-AppPersistedLogin {
    Show-AppHeader -Section 'Sesión persistente'

    try {
        Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction Stop
        Write-AppOk -Message 'Sesión persistente eliminada.'
    }
    catch {
        Write-AppWarning -Message $_.Exception.Message
    }

    Wait-App
}

function Show-AppAuthenticationMenu {
    while ($true) {
        Show-AppHeader -Section 'Aplicación Entra / autenticación PnP'

        $configured = (
            (Test-AppTenant -Value $script:Settings.Tenant) -and
            (Test-AppGuid -Value $script:Settings.ClientId)
        )

        Write-AppField -Name 'Tenant' -Value $script:Settings.Tenant
        Write-AppField -Name 'App guardada' -Value $script:Settings.AppRegistrationName
        Write-AppField -Name 'Client ID' -Value $script:Settings.ClientId
        Write-AppField -Name 'Estado config' -Value $(if ($configured) { 'CONFIGURADA' } else { 'NO CONFIGURADA' })
        Write-Host ''

        if ($configured) {
            $choice = Read-AppMenuChoice -Items @(
                @{ Key='1'; Value='Validate'; Label='Validar aplicación configurada'; Description='Comprueba autenticación y acceso a búsqueda de sitios.' }
                @{ Key='2'; Value='Existing'; Label='Usar otra aplicación existente'; Description='Introduce y opcionalmente valida otro Client ID.' }
                @{ Key='3'; Value='Register'; Label='Registrar una nueva aplicación Entra'; Description='Crea otra app PnP y la configura en esta herramienta.' }
                @{ Key='4'; Value='Remove'; Label='Quitar Client ID de la configuración local'; Description='No elimina la aplicación en Entra.' }
                @{ Key='5'; Value='Clear'; Label='Eliminar sesión persistente' }
                @{ Key='0'; Value='Back'; Label='Volver' }
            )
        }
        else {
            $choice = Read-AppMenuChoice -Items @(
                @{ Key='1'; Value='Existing'; Label='Usar una aplicación existente'; Description='Introduce y opcionalmente valida un Client ID ya registrado.' }
                @{ Key='2'; Value='Register'; Label='Registrar una nueva aplicación Entra'; Description='Crea una nueva app PnP para esta herramienta.' }
                @{ Key='3'; Value='Clear'; Label='Eliminar sesión persistente' }
                @{ Key='0'; Value='Back'; Label='Volver' }
            )
        }

        switch ($choice.Value) {
            'Validate' {
                Show-AppHeader -Section 'Validar aplicación'
                [void](Test-AppConfiguredApplication -InteractiveResult)
                Wait-App
            }

            'Existing' {
                Set-AppExistingApplication
            }

            'Register' {
                New-AppEntraRegistration
            }

            'Remove' {
                Remove-AppLocalClientId
            }

            'Clear' {
                Clear-AppPersistedLogin
            }

            'Back' {
                return
            }
        }
    }
}

function Ensure-AppAuthentication {
    if (-not (Confirm-AppPnP)) {
        return $false
    }

    if (-not (Test-AppTenant -Value $script:Settings.Tenant)) {
        [void](Ensure-AppTenant)
    }

    if (Test-AppGuid -Value $script:Settings.ClientId) {
        return $true
    }

    Show-AppHeader -Section 'Autenticación requerida'
    Write-AppWarning -Message 'No hay una aplicación Entra configurada.'
    Write-Host ''

    $choice = Read-AppMenuChoice -Items @(
        @{ Key='1'; Value='Existing'; Label='Usar una aplicación existente' }
        @{ Key='2'; Value='Register'; Label='Registrar una nueva aplicación Entra' }
        @{ Key='0'; Value='Back'; Label='Cancelar' }
    )

    switch ($choice.Value) {
        'Existing' { Set-AppExistingApplication }
        'Register' { New-AppEntraRegistration }
        'Back' { return $false }
    }

    return (
        (Test-AppTenant -Value $script:Settings.Tenant) -and
        (Test-AppGuid -Value $script:Settings.ClientId)
    )
}


# =============================================================================
# BUSQUEDA DE SITIOS
# =============================================================================

function Get-AppTenantSites {
    param(
        [ValidateSet('SharePoint','OneDrive')]
        [string]$Kind
    )

    if (-not (Ensure-AppAuthentication)) {
        return @()
    }

    $adminUrl = Get-AppAdminUrl

    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        throw 'No fue posible determinar la URL del centro de administración.'
    }

    Write-AppInfo -Message "Conectando al centro de administración: $adminUrl"

    $adminConnection = Connect-AppUrl -Url $adminUrl

    Write-AppInfo -Message 'Consultando sitios del tenant...'

    $sites = @(
        Get-PnPTenantSite `
            -IncludeOneDriveSites `
            -Detailed `
            -Connection $adminConnection `
            -ErrorAction Stop
    )

    if ($Kind -eq 'OneDrive') {
        return @(
            $sites |
                Where-Object {
                    [string]$_.Url -match '(?i)-my\.sharepoint\.com/personal/'
                } |
                Sort-Object Owner, Url
        )
    }

    return @(
        $sites |
            Where-Object {
                [string]$_.Url -notmatch '(?i)-my\.sharepoint\.com/personal/' -and
                [string]$_.Url -notmatch '(?i)-admin\.sharepoint\.com/?$'
            } |
            Sort-Object Title, Url
    )
}

function Find-AppSiteInteractively {
    param(
        [ValidateSet('SharePoint','OneDrive')]
        [string]$Kind
    )

    try {
        $sites = @(Get-AppTenantSites -Kind $Kind)
    }
    catch {
        Write-AppError -Message 'No se pudieron enumerar los sitios.'
        Write-AppMuted -Message $_.Exception.Message
        Wait-App
        return $null
    }

    if ($sites.Count -eq 0) {
        Write-AppWarning -Message 'No se encontraron sitios accesibles.'
        Wait-App
        return $null
    }

    while ($true) {
        Show-AppHeader -Section $(if ($Kind -eq 'OneDrive') { 'Buscar OneDrive' } else { 'Buscar SharePoint' })

        Write-AppField -Name 'Disponibles' -Value $sites.Count
        Write-Host ''

        $filter = Read-AppText `
            -Prompt 'Filtrar por nombre, URL o propietario (vacío = todos)' `
            -AllowEmpty

        $filtered = $sites

        if (-not [string]::IsNullOrWhiteSpace($filter)) {
            $needle = [regex]::Escape($filter)

            $filtered = @(
                $sites |
                    Where-Object {
                        ([string]$_.Title -match "(?i)$needle") -or
                        ([string]$_.Url -match "(?i)$needle") -or
                        ([string]$_.Owner -match "(?i)$needle")
                    }
            )
        }

        if ($filtered.Count -eq 0) {
            Write-AppWarning -Message 'No hay coincidencias.'
            Start-Sleep -Milliseconds 800
            continue
        }

        if ($filtered.Count -gt 50) {
            Write-AppWarning -Message "Hay $($filtered.Count) coincidencias. Refina el filtro a 50 o menos."
            Start-Sleep -Milliseconds 900
            continue
        }

        return Select-AppSingleByNumber `
            -Items $filtered `
            -Title $(if ($Kind -eq 'OneDrive') { 'OneDrive' } else { 'Sitios SharePoint' }) `
            -Label {
                param($site)

                $titleProperty = $site.PSObject.Properties['Title']
                $ownerProperty = $site.PSObject.Properties['Owner']
                $urlProperty = $site.PSObject.Properties['Url']

                $title = if ($null -ne $titleProperty) { [string]$titleProperty.Value } else { '' }
                $owner = if ($null -ne $ownerProperty) { [string]$ownerProperty.Value } else { '' }
                $url = if ($null -ne $urlProperty) { [string]$urlProperty.Value } else { '' }

                if ([string]::IsNullOrWhiteSpace($title)) {
                    $title = if (-not [string]::IsNullOrWhiteSpace($owner)) { $owner } else { '(sin título)' }
                }

                $ownerText = if (-not [string]::IsNullOrWhiteSpace($owner)) { " · $owner" } else { '' }

                "$title$ownerText`n     $url"
            }
    }
}

function New-AppDirectTarget {
    Show-AppHeader -Section 'URL directa'

    $url = Read-AppText `
        -Prompt 'URL completa de SharePoint / OneDrive' `
        -Validator { param($v) Test-AppM365Url -Value $v } `
        -ValidationMessage 'Introduce una URL HTTPS válida de *.sharepoint.com.'

    $kind = if ($url -match '(?i)-my\.sharepoint\.com/personal/') {
        'OneDrive'
    }
    else {
        'SharePoint'
    }

    return [pscustomobject]@{
        Type  = $kind
        Url   = $url.TrimEnd('/')
        Title = 'Destino directo'
        Owner = ''
    }
}

function Test-AppRecycleBinAccess {
    param(
        [Parameter(Mandatory)]$Connection
    )

    try {
        $null = Get-PnPRecycleBinItem `
            -FirstStage `
            -RowLimit 1 `
            -Connection $Connection `
            -ErrorAction Stop

        return [pscustomobject]@{
            Success = $true
            Message = ''
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

function Set-AppTarget {
    Show-AppHeader -Section 'Seleccionar destino'

    if (-not (Ensure-AppAuthentication)) {
        return
    }

    $choice = Read-AppMenuChoice -AllowBack -Items @(
        @{ Key='1'; Value='SharePoint'; Label='SharePoint Online'; Description='Buscar sitios del tenant.' }
        @{ Key='2'; Value='OneDrive'; Label='OneDrive for Business'; Description='Buscar OneDrive reales del tenant.' }
        @{ Key='3'; Value='Direct'; Label='URL directa'; Description='Detecta automáticamente SharePoint u OneDrive.' }
    )

    if ($choice.Value -eq 'Back') {
        return
    }

    $site = $null

    switch ($choice.Value) {
        'SharePoint' {
            $site = Find-AppSiteInteractively -Kind SharePoint
        }

        'OneDrive' {
            $site = Find-AppSiteInteractively -Kind OneDrive
        }

        'Direct' {
            $site = New-AppDirectTarget
        }
    }

    if ($null -eq $site) {
        return
    }

    $urlProperty = $site.PSObject.Properties['Url']

    if ($null -eq $urlProperty -or -not (Test-AppM365Url -Value ([string]$urlProperty.Value))) {
        Show-AppErrorScreen -Title 'Destino' -Message 'El sitio seleccionado no contiene una URL válida.'
        return
    }

    $url = [string]$urlProperty.Value

    $typeProperty = $site.PSObject.Properties['Type']
    $type = if ($null -ne $typeProperty -and -not [string]::IsNullOrWhiteSpace([string]$typeProperty.Value)) {
        [string]$typeProperty.Value
    }
    elseif ($url -match '(?i)-my\.sharepoint\.com/personal/') {
        'OneDrive'
    }
    else {
        'SharePoint'
    }

    Show-AppHeader -Section 'Validando destino'
    Write-AppInfo -Message $url

    try {
        $connection = Connect-AppUrl -Url $url

        $web = Get-PnPWeb `
            -Connection $connection `
            -Includes Title,Url `
            -ErrorAction Stop

        Write-AppOk -Message 'Conexión establecida.'
        Write-AppMuted -Message $web.Title

        Write-AppInfo -Message 'Validando acceso a la papelera...'

        $access = Test-AppRecycleBinAccess -Connection $connection

        if (-not $access.Success) {
            Write-AppError -Message 'La cuenta conectada no puede leer la papelera de este sitio.'
            Write-AppMuted -Message 'Get-PnPRecycleBinItem requiere Site Collection Administrator.'
            Write-AppMuted -Message $access.Message
            Wait-App
            return
        }

        Write-AppOk -Message 'Acceso a la papelera validado.'

        $owner = ''
        $ownerProperty = $site.PSObject.Properties['Owner']
        if ($null -ne $ownerProperty) {
            $owner = [string]$ownerProperty.Value
        }

        $script:PnPConnection = $connection
        $script:Target = [pscustomobject]@{
            Type    = $type
            Url     = $url.TrimEnd('/')
            Title   = [string]$web.Title
            Owner   = $owner
        }

        Start-Sleep -Milliseconds 700
    }
    catch {
        Show-AppErrorScreen -Title 'Conexión al destino' -Message $_.Exception.Message
    }
}

function Confirm-AppTarget {
    if ($null -eq $script:Target) {
        Write-AppWarning -Message 'Primero selecciona un sitio SharePoint o OneDrive.'
        Wait-App
        return $false
    }

    if ($null -eq $script:PnPConnection) {
        try {
            $script:PnPConnection = Connect-AppUrl -Url $script:Target.Url
        }
        catch {
            Show-AppErrorScreen -Title 'Conexión' -Message $_.Exception.Message
            return $false
        }
    }

    return $true
}


# =============================================================================
# PAPELERA
# =============================================================================

function Get-AppRecycleBinItems {
    param(
        [ValidateSet('FirstStage','SecondStage')]
        [string]$Stage,

        [Parameter(Mandatory)][int]$Limit
    )

    $operation = {
        if ($Stage -eq 'SecondStage') {
            return @(
                Get-PnPRecycleBinItem `
                    -SecondStage `
                    -RowLimit $Limit `
                    -Connection $script:PnPConnection `
                    -ErrorAction Stop
            )
        }

        return @(
            Get-PnPRecycleBinItem `
                -FirstStage `
                -RowLimit $Limit `
                -Connection $script:PnPConnection `
                -ErrorAction Stop
        )
    }

    return @(
        Invoke-AppWithRetry `
            -Operation $operation `
            -OperationName "Leer $Stage"
    )
}

function Get-AppRecycleItemName {
    param([Parameter(Mandatory)]$Item)

    foreach ($propertyName in @('LeafName','DirName','Title')) {
        $property = $Item.PSObject.Properties[$propertyName]

        if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return [string]$Item.Id
}

function Confirm-AppDestructiveOperation {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Description,
        [string]$Phrase = 'ELIMINAR'
    )

    if ([bool]$script:Settings.Simulation) {
        return $true
    }

    Show-AppHeader -Section $Title
    Write-AppError -Message 'MODO REAL'
    Write-AppMuted -Message $Description
    Write-Host ''
    Write-AppStyled -Text "Escribe $Phrase para continuar." -Style Danger
    Write-Host ''

    $confirmation = (Read-Host 'Confirmación').Trim()

    return ($confirmation -ceq $Phrase)
}

function New-AppAuditLog {
    # PowerShell enumera las colecciones devueltas por una función.
    # Una Generic.List vacía produciría cero objetos y el llamador recibiría $null.
    # La coma unaria fuerza a devolver la lista como un único objeto.
    return ,([System.Collections.Generic.List[object]]::new())
}

function Export-AppAuditLog {
    param(
        [Parameter(Mandatory)]$Log,
        [Parameter(Mandatory)][string]$Prefix
    )

    Initialize-AppFolders

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $path = Join-Path $script:ReportsPath "${Prefix}-${stamp}.csv"

    if ($null -eq $Log -or $Log.Count -eq 0) {
        Set-Content -LiteralPath $path -Value '' -Encoding UTF8
        return $path
    }

    $rows = if ($Log.PSObject.Methods['ToArray']) {
        $Log.ToArray()
    }
    else {
        @($Log | ForEach-Object { $_ })
    }

    $rows |
        Export-Csv `
            -LiteralPath $path `
            -NoTypeInformation `
            -Encoding UTF8 `
            -Force

    return $path
}

function Move-AppRecycleBinItemToSecondStage {
    param(
        [Parameter(Mandatory)]$Item
    )

    # Get-PnPRecycleBinItem devuelve normalmente un objeto CSOM
    # Microsoft.SharePoint.Client.RecycleBinItem. Si dispone directamente del
    # método MoveToSecondStage(), lo usamos. Si no, resolvemos el objeto por ID
    # desde el contexto CSOM actual.
    if ($null -ne $Item.PSObject.Methods['MoveToSecondStage']) {
        $Item.MoveToSecondStage()

        $contextProperty = $Item.PSObject.Properties['Context']

        if ($null -ne $contextProperty -and $null -ne $contextProperty.Value) {
            $contextProperty.Value.ExecuteQuery()
            return
        }
    }

    $ctx = Get-PnPContext -Connection $script:PnPConnection

    if ($null -eq $ctx) {
        throw 'No se pudo obtener el contexto CSOM de la conexión PnP.'
    }

    $id = [Guid]::Empty

    if (-not [Guid]::TryParse([string]$Item.Id, [ref]$id)) {
        throw "El elemento de papelera no contiene un Id GUID válido: $($Item.Id)"
    }

    $recycleItem = $ctx.Web.RecycleBin.GetById($id)
    $recycleItem.MoveToSecondStage()
    $ctx.ExecuteQuery()
}

function Invoke-AppRecycleBinCleanup {
    param(
        [ValidateSet('FirstStage','SecondStage')]
        [string]$Stage,

        [Parameter(Mandatory)][int]$Limit
    )

    if (-not (Confirm-AppTarget)) {
        return $null
    }

    $isFirstStage = ($Stage -eq 'FirstStage')

    $stageLabel = if ($isFirstStage) {
        'Primer nivel'
    }
    else {
        'Segundo nivel'
    }

    $operationLabel = if ($isFirstStage) {
        'Mover al segundo nivel'
    }
    else {
        'Eliminar permanentemente'
    }

    Show-AppHeader -Section "Papelera · $stageLabel"

    Write-AppField -Name 'Sitio' -Value $script:Target.Url
    Write-AppField -Name 'Nivel' -Value $stageLabel
    Write-AppField -Name 'Acción' -Value $operationLabel
    Write-AppField -Name 'Límite' -Value $Limit
    Write-AppField -Name 'Modo' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' })
    Write-Host ''

    if ($isFirstStage) {
        Write-AppInfo -Message 'Los elementos del primer nivel se moverán al segundo nivel; no se eliminarán permanentemente.'
    }
    else {
        Write-AppWarning -Message 'Los elementos del segundo nivel se eliminarán permanentemente.'
    }

    Write-Host ''
    Write-AppInfo -Message 'Leyendo elementos...'

    try {
        $items = @(Get-AppRecycleBinItems -Stage $Stage -Limit $Limit)
    }
    catch {
        Show-AppErrorScreen -Title "Papelera · $stageLabel" -Message $_.Exception.Message
        return $null
    }

    if ($items.Count -eq 0) {
        Write-AppOk -Message 'No hay elementos en este nivel dentro del límite solicitado.'
        Wait-App

        return [pscustomobject]@{
            Stage     = $Stage
            Found     = 0
            Processed = 0
            Moved     = 0
            Removed   = 0
            Simulated = 0
            Errors    = 0
            Report    = ''
        }
    }

    Write-AppOk -Message "$($items.Count) elemento(s) encontrados."

    $description = if ($isFirstStage) {
        "Se moverán hasta $($items.Count) elemento(s) del primer nivel al segundo nivel. No es una eliminación permanente."
    }
    else {
        "Se eliminarán permanentemente hasta $($items.Count) elemento(s) del segundo nivel."
    }

    $phrase = if ($isFirstStage) {
        'MOVER'
    }
    else {
        'ELIMINAR'
    }

    if (-not (Confirm-AppDestructiveOperation `
        -Title "Confirmar · $stageLabel" `
        -Description $description `
        -Phrase $phrase)) {

        Write-AppWarning -Message 'Operación cancelada.'
        Wait-App
        return $null
    }

    Show-AppHeader -Section "Procesando · $stageLabel"

    $log = [System.Collections.Generic.List[object]]::new()
    $moved = 0
    $removed = 0
    $simulated = 0
    $errors = 0

    for ($i = 0; $i -lt $items.Count; $i++) {
        $item = $items[$i]
        $name = Get-AppRecycleItemName -Item $item
        $status = 'OK'
        $message = ''

        $action = if ($script:Settings.Simulation) {
            if ($isFirstStage) { 'Simulado: mover a segundo nivel' } else { 'Simulado: eliminar permanentemente' }
        }
        else {
            if ($isFirstStage) { 'Movido a segundo nivel' } else { 'Eliminado permanentemente' }
        }

        $percent = [int]((($i + 1) / [Math]::Max(1, $items.Count)) * 100)

        Write-Progress `
            -Id 1 `
            -Activity "Procesando $stageLabel" `
            -Status "$($i + 1) de $($items.Count) · $name" `
            -PercentComplete $percent

        try {
            if ($script:Settings.Simulation) {
                $simulated++
            }
            elseif ($isFirstStage) {
                Invoke-AppWithRetry `
                    -Operation {
                        Move-AppRecycleBinItemToSecondStage -Item $item
                    } `
                    -OperationName "Mover a segundo nivel · $name"

                $moved++
            }
            else {
                Invoke-AppWithRetry `
                    -Operation {
                        Clear-PnPRecycleBinItem `
                            -Identity $item `
                            -Force `
                            -Connection $script:PnPConnection `
                            -ErrorAction Stop
                    } `
                    -OperationName "Eliminar permanentemente · $name"

                $removed++
            }
        }
        catch {
            $status = 'ERROR'
            $message = $_.Exception.Message
            $errors++
        }

        [void]$log.Add(
            [pscustomobject]@{
                FechaHora = Get-Date
                Sitio     = $script:Target.Url
                TipoSitio = $script:Target.Type
                Etapa     = $Stage
                Id        = [string]$item.Id
                Nombre    = $name
                Accion    = $action
                Estado    = $status
                Mensaje   = $message
            }
        )
    }

    Write-Progress -Id 1 -Activity "Procesando $stageLabel" -Completed

    $report = Export-AppAuditLog -Log $log -Prefix "Papelera-$Stage"

    Show-AppHeader -Section "Resultado · $stageLabel"

    Write-AppField -Name 'Encontrados' -Value $items.Count
    Write-AppField -Name 'Procesados' -Value $log.Count

    if ($isFirstStage) {
        Write-AppField -Name 'Movidos' -Value $moved
    }
    else {
        Write-AppField -Name 'Eliminados' -Value $removed
    }

    Write-AppField -Name 'Simulados' -Value $simulated
    Write-AppField -Name 'Errores' -Value $errors
    Write-AppField -Name 'Reporte' -Value $report

    Write-Host ''

    if ($errors -gt 0) {
        Write-AppWarning -Message 'La operación terminó con errores. Revisa el reporte.'
    }
    elseif ($script:Settings.Simulation) {
        if ($isFirstStage) {
            Write-AppOk -Message 'Simulación completada. No se movió ningún elemento.'
        }
        else {
            Write-AppOk -Message 'Simulación completada. No se eliminó ningún elemento.'
        }
    }
    elseif ($isFirstStage) {
        Write-AppOk -Message 'Los elementos fueron movidos al segundo nivel de la papelera.'
    }
    else {
        Write-AppOk -Message 'Los elementos del segundo nivel fueron eliminados permanentemente.'
    }

    Wait-App

    return [pscustomobject]@{
        Stage     = $Stage
        Found     = $items.Count
        Processed = $log.Count
        Moved     = $moved
        Removed   = $removed
        Simulated = $simulated
        Errors    = $errors
        Report    = $report
    }
}

# =============================================================================
# PRESERVATION HOLD LIBRARY
# =============================================================================

function Get-AppPreservationHoldLibrary {
    if (-not (Confirm-AppTarget)) {
        return $null
    }

    $operation = {
        @(
            Get-PnPList `
                -Connection $script:PnPConnection `
                -Includes Title,Id,RootFolder,BaseType,BaseTemplate,Hidden,ItemCount `
                -ErrorAction Stop
        )
    }

    $lists = @(
        Invoke-AppWithRetry `
            -Operation $operation `
            -OperationName 'Detectar Preservation Hold Library'
    )

    foreach ($list in $lists) {
        if ([string]$list.BaseType -ne 'DocumentLibrary') {
            continue
        }

        $root = $list.RootFolder

        if ($null -eq $root) {
            try {
                $root = Get-PnPProperty `
                    -ClientObject $list `
                    -Property RootFolder `
                    -Connection $script:PnPConnection
            }
            catch {
                continue
            }
        }

        $rootUrl = if ($null -ne $root) { [string]$root.ServerRelativeUrl } else { '' }
        $titleNormalized = ([string]$list.Title -replace '[\s_-]', '')
        $rootLeaf = if (-not [string]::IsNullOrWhiteSpace($rootUrl)) {
            ($rootUrl.TrimEnd('/') -split '/')[-1]
        }
        else {
            ''
        }

        if (
            $titleNormalized -match '(?i)^PreservationHoldLibrary$' -or
            $rootLeaf -match '(?i)^PreservationHoldLibrary$'
        ) {
            return $list
        }
    }

    return $null
}

function Get-AppPreservationItems {
    param(
        [Parameter(Mandatory)]$Library,
        [Parameter(Mandatory)][int]$Limit
    )

    $operation = {
        @(
            Get-PnPListItem `
                -List $Library.Id `
                -PageSize 200 `
                -Connection $script:PnPConnection `
                -ErrorAction Stop |
                    Select-Object -First $Limit
        )
    }

    return @(
        Invoke-AppWithRetry `
            -Operation $operation `
            -OperationName 'Leer Preservation Hold Library'
    )
}

function Get-AppListItemName {
    param([Parameter(Mandatory)]$Item)

    if ($null -ne $Item.FieldValues) {
        foreach ($fieldName in @('FileLeafRef','Title','FileRef')) {
            if ($Item.FieldValues.ContainsKey($fieldName)) {
                $value = [string]$Item.FieldValues[$fieldName]

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    return $value
                }
            }
        }
    }

    return "Item $($Item.Id)"
}

function Invoke-AppPreservationHoldCleanup {
    param(
        [Parameter(Mandatory)][int]$Limit
    )

    if (-not (Confirm-AppTarget)) {
        return
    }

    Show-AppHeader -Section 'Preservation Hold Library'

    Write-AppField -Name 'Sitio' -Value $script:Target.Url
    Write-AppField -Name 'Límite' -Value $Limit
    Write-AppField -Name 'Modo' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' })
    Write-Host ''

    Write-AppInfo -Message 'Buscando Preservation Hold Library...'

    try {
        $library = Get-AppPreservationHoldLibrary
    }
    catch {
        Show-AppErrorScreen -Title 'Preservation Hold Library' -Message $_.Exception.Message
        return
    }

    if ($null -eq $library) {
        Write-AppWarning -Message 'No se encontró Preservation Hold Library en este sitio.'
        Write-AppMuted -Message 'Puede no existir, estar vacía/no provisionada o no ser accesible con la identidad actual.'
        Wait-App
        return
    }

    Write-AppOk -Message "Biblioteca detectada: $($library.Title)"
    Write-AppMuted -Message ([string]$library.RootFolder.ServerRelativeUrl)
    Write-Host ''

    try {
        $items = @(Get-AppPreservationItems -Library $library -Limit $Limit)
    }
    catch {
        Show-AppErrorScreen -Title 'Preservation Hold Library' -Message $_.Exception.Message
        return
    }

    if ($items.Count -eq 0) {
        Write-AppOk -Message 'No hay elementos para procesar dentro del límite solicitado.'
        Wait-App
        return
    }

    Write-AppInfo -Message "$($items.Count) elemento(s) encontrados."

    if (-not $script:Settings.Simulation) {
        Write-AppWarning -Message 'Eliminar contenido de Preservation Hold Library puede verse bloqueado por políticas de retención.'
        Write-AppMuted -Message 'La herramienta no deshabilita ni modifica políticas de cumplimiento o retención.'
    }

    if (-not (Confirm-AppDestructiveOperation `
        -Title 'Confirmar Preservation Hold Library' `
        -Description "Se intentará eliminar permanentemente hasta $($items.Count) elemento(s) de la biblioteca detectada." `
        -Phrase 'ELIMINAR PHL')) {

        Write-AppWarning -Message 'Operación cancelada.'
        Wait-App
        return
    }

    Show-AppHeader -Section 'Procesando Preservation Hold Library'

    $log = [System.Collections.Generic.List[object]]::new()
    $removed = 0
    $simulated = 0
    $errors = 0

    for ($i = 0; $i -lt $items.Count; $i++) {
        $item = $items[$i]
        $name = Get-AppListItemName -Item $item
        $status = 'OK'
        $message = ''
        $action = if ($script:Settings.Simulation) { 'Simulado' } else { 'Eliminado' }

        $percent = [int]((($i + 1) / [Math]::Max(1, $items.Count)) * 100)

        Write-Progress `
            -Id 2 `
            -Activity 'Procesando Preservation Hold Library' `
            -Status "$($i + 1) de $($items.Count) · $name" `
            -PercentComplete $percent

        try {
            if ($script:Settings.Simulation) {
                $simulated++
            }
            else {
                Invoke-AppWithRetry `
                    -Operation {
                        Remove-PnPListItem `
                            -List $library.Id `
                            -Identity $item.Id `
                            -Force `
                            -Connection $script:PnPConnection `
                            -ErrorAction Stop
                    } `
                    -OperationName "Eliminar PHL · $name"

                $removed++
            }
        }
        catch {
            $status = 'ERROR'
            $message = $_.Exception.Message
            $errors++
        }

        [void]$log.Add(
            [pscustomobject]@{
                FechaHora = Get-Date
                Sitio     = $script:Target.Url
                TipoSitio = $script:Target.Type
                Biblioteca= [string]$library.Title
                ItemId    = $item.Id
                Nombre    = $name
                Accion    = $action
                Estado    = $status
                Mensaje   = $message
            }
        )
    }

    Write-Progress -Id 2 -Activity 'Procesando Preservation Hold Library' -Completed

    $report = Export-AppAuditLog -Log $log -Prefix 'Preservation-Hold-Library'

    Show-AppHeader -Section 'Resultado · Preservation Hold Library'

    Write-AppField -Name 'Encontrados' -Value $items.Count
    Write-AppField -Name 'Procesados' -Value $log.Count
    Write-AppField -Name 'Eliminados' -Value $removed
    Write-AppField -Name 'Simulados' -Value $simulated
    Write-AppField -Name 'Errores' -Value $errors
    Write-AppField -Name 'Reporte' -Value $report
    Write-Host ''

    if ($errors -gt 0) {
        Write-AppWarning -Message 'La operación terminó con errores. Revisa el reporte.'
    }
    elseif ($script:Settings.Simulation) {
        Write-AppOk -Message 'Simulación completada. No se modificó la biblioteca.'
    }
    else {
        Write-AppOk -Message 'Procesamiento completado.'
    }

    Wait-App
}


# =============================================================================
# FLUJOS DE LIMPIEZA
# =============================================================================

function Read-AppCleanupLimit {
    return Read-AppInteger `
        -Prompt 'Cantidad máxima de elementos a procesar' `
        -Minimum 1 `
        -Maximum ([int]$script:Settings.MaximumLimit) `
        -DefaultValue ([int]$script:Settings.DefaultLimit)
}

function Start-AppStageCleanup {
    param(
        [ValidateSet('FirstStage','SecondStage','Both','Preservation')]
        [string]$Mode
    )

    if (-not (Confirm-AppTarget)) {
        return
    }

    Show-AppHeader -Section 'Límite de procesamiento'

    Write-AppMuted -Message "Máximo configurado por ejecución: $($script:Settings.MaximumLimit)"
    Write-AppMuted -Message 'Primer nivel: mover al segundo. Segundo nivel: eliminación permanente.'
    Write-AppMuted -Message 'En "ambos niveles", el límite se aplica por separado a cada nivel.'
    Write-Host ''

    $limit = Read-AppCleanupLimit

    switch ($Mode) {
        'FirstStage' {
            [void](Invoke-AppRecycleBinCleanup -Stage FirstStage -Limit $limit)
        }

        'SecondStage' {
            [void](Invoke-AppRecycleBinCleanup -Stage SecondStage -Limit $limit)
        }

        'Both' {
            [void](Invoke-AppRecycleBinCleanup -Stage FirstStage -Limit $limit)
            [void](Invoke-AppRecycleBinCleanup -Stage SecondStage -Limit $limit)
        }

        'Preservation' {
            Invoke-AppPreservationHoldCleanup -Limit $limit
        }
    }
}


# =============================================================================
# CONFIGURACION / DIAGNOSTICO
# =============================================================================

function Open-AppReports {
    Initialize-AppFolders

    try {
        if ($IsWindows) {
            Start-Process explorer.exe -ArgumentList $script:ReportsPath
        }
        else {
            Write-AppInfo -Message $script:ReportsPath
            Wait-App
        }
    }
    catch {
        Show-AppErrorScreen -Title 'Reportes' -Message "No fue posible abrir $script:ReportsPath"
    }
}

function Show-AppThrottleSettings {
    Show-AppHeader -Section 'Throttling y reintentos'

    Write-AppMuted -Message 'Valores conservadores para trabajo secuencial.'
    Write-Host ''

    $script:Settings.RequestDelayMs = Read-AppInteger `
        -Prompt 'Pausa entre solicitudes (ms)' `
        -Minimum 0 `
        -Maximum 10000 `
        -DefaultValue ([int]$script:Settings.RequestDelayMs)

    $script:Settings.MaxRetries = Read-AppInteger `
        -Prompt 'Máximo de reintentos' `
        -Minimum 1 `
        -Maximum 20 `
        -DefaultValue ([int]$script:Settings.MaxRetries)

    $script:Settings.RetryBaseSeconds = Read-AppInteger `
        -Prompt 'Espera base de retry (segundos)' `
        -Minimum 1 `
        -Maximum 60 `
        -DefaultValue ([int]$script:Settings.RetryBaseSeconds)

    $script:Settings.RetryMaxSeconds = Read-AppInteger `
        -Prompt 'Espera máxima de retry (segundos)' `
        -Minimum 5 `
        -Maximum 600 `
        -DefaultValue ([int]$script:Settings.RetryMaxSeconds)

    Save-AppSettings

    Write-Host ''
    Write-AppOk -Message 'Configuración de throttling guardada.'
    Wait-App
}

function Show-AppSettingsMenu {
    while ($true) {
        Show-AppHeader -Section 'Configuración'

        Write-AppField -Name 'Tenant' -Value $script:Settings.Tenant
        Write-AppField -Name 'Client ID' -Value $script:Settings.ClientId
        Write-AppField -Name 'Persist login' -Value $(if ($script:Settings.PersistLogin) { 'Sí' } else { 'No' })
        Write-AppField -Name 'Modo inicial' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' })
        Write-AppField -Name 'Límite defecto' -Value $script:Settings.DefaultLimit
        Write-AppField -Name 'Límite máximo' -Value $script:Settings.MaximumLimit
        Write-AppField -Name 'Reportes' -Value $script:ReportsPath
        Write-Host ''

        $choice = Read-AppMenuChoice -Items @(
            @{ Key='1'; Value='Auth'; Label='Aplicación Entra / autenticación' }
            @{ Key='2'; Value='Persist'; Label='Alternar persistencia de login' }
            @{ Key='3'; Value='DefaultLimit'; Label='Cambiar límite predeterminado' }
            @{ Key='4'; Value='MaxLimit'; Label='Cambiar límite máximo permitido'; Description='Rango permitido por la herramienta: 1 a 10000.' }
            @{ Key='5'; Value='Throttle'; Label='Throttling y reintentos' }
            @{ Key='6'; Value='Reports'; Label='Abrir reportes' }
            @{ Key='7'; Value='Reset'; Label='Restablecer configuración local' }
            @{ Key='0'; Value='Back'; Label='Volver' }
        )

        switch ($choice.Value) {
            'Auth' {
                Show-AppAuthenticationMenu
            }

            'Persist' {
                $script:Settings.PersistLogin = -not [bool]$script:Settings.PersistLogin
                Save-AppSettings
            }

            'DefaultLimit' {
                Show-AppHeader -Section 'Límite predeterminado'
                $script:Settings.DefaultLimit = Read-AppInteger `
                    -Prompt 'Límite predeterminado' `
                    -Minimum 1 `
                    -Maximum ([int]$script:Settings.MaximumLimit) `
                    -DefaultValue ([int]$script:Settings.DefaultLimit)
                Save-AppSettings
            }

            'MaxLimit' {
                Show-AppHeader -Section 'Límite máximo'
                $newMaximum = Read-AppInteger `
                    -Prompt 'Límite máximo' `
                    -Minimum 1 `
                    -Maximum 10000 `
                    -DefaultValue ([int]$script:Settings.MaximumLimit)

                $script:Settings.MaximumLimit = $newMaximum

                if ([int]$script:Settings.DefaultLimit -gt $newMaximum) {
                    $script:Settings.DefaultLimit = $newMaximum
                }

                Save-AppSettings
            }

            'Throttle' {
                Show-AppThrottleSettings
            }

            'Reports' {
                Open-AppReports
            }

            'Reset' {
                Show-AppHeader -Section 'Restablecer configuración'
                Write-AppWarning -Message 'Se eliminarán las preferencias locales de esta herramienta.'
                Write-Host ''

                if (Read-AppYesNo -Prompt '¿Continuar?' -DefaultYes $false) {
                    if (Test-Path -LiteralPath $script:ConfigPath) {
                        Remove-Item -LiteralPath $script:ConfigPath -Force
                    }

                    $script:Settings = Get-DefaultAppSettings
                    $script:Target = $null
                    $script:PnPConnection = $null

                    Write-AppOk -Message 'Configuración restablecida.'
                    Wait-App
                }
            }

            'Back' {
                return
            }
        }
    }
}

function Show-AppDiagnostics {
    Show-AppHeader -Section 'Diagnóstico'

    $pnpVersion = Get-AppPnPInstalledVersion

    Write-AppField -Name 'PowerShell' -Value $PSVersionTable.PSVersion
    Write-AppField -Name 'PnP.PowerShell' -Value $(if ($null -ne $pnpVersion) { $pnpVersion } else { 'No instalado' })
    Write-AppField -Name 'Tenant' -Value $script:Settings.Tenant
    Write-AppField -Name 'Client ID' -Value $script:Settings.ClientId
    Write-AppField -Name 'Persist login' -Value $(if ($script:Settings.PersistLogin) { 'Sí' } else { 'No' })
    Write-AppField -Name 'Config' -Value $script:ConfigPath
    Write-AppField -Name 'Reportes' -Value $script:ReportsPath
    Write-AppField -Name 'Throttling' -Value $script:ThrottleEvents.Count
    Write-Host ''

    $choice = Read-AppMenuChoice -Items @(
        @{ Key='1'; Value='Validate'; Label='Validar aplicación configurada' }
        @{ Key='2'; Value='PnP'; Label='Instalar / actualizar PnP.PowerShell' }
        @{ Key='3'; Value='Target'; Label='Probar nuevamente el sitio actual' }
        @{ Key='0'; Value='Back'; Label='Volver' }
    )

    switch ($choice.Value) {
        'Validate' {
            Show-AppHeader -Section 'Validar aplicación'
            [void](Test-AppConfiguredApplication -InteractiveResult)
            Wait-App
        }

        'PnP' {
            [void](Confirm-AppPnP)
        }

        'Target' {
            if ($null -eq $script:Target) {
                Write-AppWarning -Message 'No hay un sitio seleccionado.'
                Wait-App
                return
            }

            try {
                $script:PnPConnection = Connect-AppUrl -Url $script:Target.Url
                $access = Test-AppRecycleBinAccess -Connection $script:PnPConnection

                if ($access.Success) {
                    Write-AppOk -Message 'Conexión y acceso a papelera correctos.'
                }
                else {
                    Write-AppError -Message 'La conexión funciona, pero no hay acceso suficiente a la papelera.'
                    Write-AppMuted -Message $access.Message
                }
            }
            catch {
                Write-AppError -Message $_.Exception.Message
            }

            Wait-App
        }
    }
}


# =============================================================================
# PANTALLA PRINCIPAL
# =============================================================================

function Show-AppContext {
    Write-AppSection -Title 'Contexto'

    Write-AppField -Name 'Tenant' -Value $script:Settings.Tenant

    if ($null -ne $script:Target) {
        Write-AppField -Name 'Tipo' -Value $script:Target.Type
        Write-AppField -Name 'Sitio' -Value $script:Target.Title
        Write-AppField -Name 'URL' -Value $script:Target.Url
    }
    else {
        Write-AppField -Name 'Destino' -Value 'No seleccionado'
    }

    Write-AppField -Name 'Modo' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' })
}

function Show-AppMainMenu {
    while ($true) {
        Show-AppHeader -Section 'Inicio'
        Show-AppContext
        Write-Host ''

        $choice = Read-AppMenuChoice -Items @(
            @{ Key='1'; Value='Target'; Label='Seleccionar / cambiar sitio'; Description='Buscar SharePoint u OneDrive en el tenant o usar una URL directa.' }
            @{ Key='2'; Value='First'; Label='Vaciar primer nivel'; Description='Mueve los elementos al segundo nivel; no los elimina permanentemente.' }
            @{ Key='3'; Value='Second'; Label='Vaciar segundo nivel'; Description='Elimina permanentemente elementos del segundo nivel.' }
            @{ Key='4'; Value='Both'; Label='Vaciar ambos niveles'; Description='Primero mueve el primer nivel al segundo y después elimina del segundo; el límite se aplica por nivel.' }
            @{ Key='5'; Value='PHL'; Label='Procesar Preservation Hold Library'; Description='Detecta dinámicamente la biblioteca; no modifica políticas de retención.' }
            @{ Key='6'; Value='Mode'; Label=$(if ($script:Settings.Simulation) { 'Cambiar a modo REAL' } else { 'Cambiar a modo SIMULACIÓN' }); Description='Simulación es el modo seguro.' }
            @{ Key='7'; Value='Auth'; Label='Aplicación Entra / autenticación' }
            @{ Key='8'; Value='Settings'; Label='Configuración' }
            @{ Key='9'; Value='Diagnostics'; Label='Diagnóstico' }
            @{ Key='0'; Value='Exit'; Label='Salir' }
        )

        switch ($choice.Value) {
            'Target' {
                Set-AppTarget
            }

            'First' {
                Start-AppStageCleanup -Mode FirstStage
            }

            'Second' {
                Start-AppStageCleanup -Mode SecondStage
            }

            'Both' {
                Start-AppStageCleanup -Mode Both
            }

            'PHL' {
                Start-AppStageCleanup -Mode Preservation
            }

            'Mode' {
                $script:Settings.Simulation = -not [bool]$script:Settings.Simulation
                Save-AppSettings
            }

            'Auth' {
                Show-AppAuthenticationMenu
            }

            'Settings' {
                Show-AppSettingsMenu
            }

            'Diagnostics' {
                Show-AppDiagnostics
            }

            'Exit' {
                return
            }
        }
    }
}


# =============================================================================
# INICIO
# =============================================================================

try {
    Initialize-AppLanguage
    Initialize-AppTerminal
    Initialize-AppFolders

    $script:Settings = Get-AppSettings

    try {
        $Host.UI.RawUI.WindowTitle = "$script:AppName $script:AppVersion"
    }
    catch {
    }

    if (-not (Test-AppPowerShellVersion)) {
        Write-Host ''
        Write-AppWarning -Message 'Ejecuta la herramienta desde PowerShell 7.4 o superior usando pwsh.'
        Wait-App
        exit 1
    }

    Show-AppMainMenu
}
catch {
    Show-AppHeader -Section 'Error inesperado'
    Write-AppError -Message $_.Exception.Message

    if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
        Write-Host ''
        Write-AppMuted -Message 'Detalles técnicos:'
        Write-AppMuted -Message $_.ScriptStackTrace
    }

    Wait-App
}
finally {
    Disconnect-AppM365
}

Show-AppHeader -Section 'Finalizado'
Write-AppMuted -Message 'Sesión finalizada.'
