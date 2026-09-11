#Requires -Version 7.4

<#
.SYNOPSIS
    Microsoft 365 Version History Cleaner

.DESCRIPTION
    Herramienta interactiva para analizar y limpiar versiones históricas
    en SharePoint Online y OneDrive for Business.

    Funciones:
      - Menús numéricos
      - Selector visual de bibliotecas con flechas + ESPACIO
      - SharePoint Online
      - OneDrive for Business
      - Bibliotecas estándar (101), OneDrive/MySite (700) y demás DocumentLibrary
      - Enumeración recursiva de archivos
      - Vista previa obligatoria
      - Resumen por archivo
      - Estimación de almacenamiento recuperable cuando está disponible
      - Revalidación de versiones inmediatamente antes de eliminar
      - Protección frente a cambios ocurridos después del análisis
      - Throttling 429/503
      - Retry-After cuando está disponible
      - Backoff exponencial + jitter
      - Pausa configurable entre peticiones
      - Sin operaciones paralelas
      - Papelera de reciclaje
      - Reportes CSV
      - Registro de auditoría
      - Configuración persistente
      - Autenticación moderna PnP / Microsoft Entra ID

.CONFIGURACION
    %LOCALAPPDATA%\M365VersionHistoryCleaner\config.json

.REPORTES
    %LOCALAPPDATA%\M365VersionHistoryCleaner\Reportes
#>

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"


# ============================================================
# APLICACION
# ============================================================

$script:AppName = "Microsoft 365 Version History Cleaner"
$script:AppVersion = "8.0"

if ($env:LOCALAPPDATA) {
    $script:AppRoot = Join-Path $env:LOCALAPPDATA "M365VersionHistoryCleaner"
}
else {
    $script:AppRoot = Join-Path $HOME ".M365VersionHistoryCleaner"
}

$script:ConfigPath = Join-Path $script:AppRoot "config.json"
$script:ReportsPath = Join-Path $script:AppRoot "Reportes"

$script:PnPConnection = $null
$script:ThrottleEvents = [System.Collections.Generic.List[object]]::new()


# ============================================================
# INTERFAZ
# ============================================================

# ============================================================
# INTERFAZ MODERNA
# ============================================================

$script:Ansi = $false

function Initialize-AppTerminal {
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

    if (-not $script:Ansi) { return '' }

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
        [Parameter(Mandatory = $true)][string]$Text,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal',
        [switch]$NoNewline
    )

    $prefix = ''
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
        if ($NoNewline) { Write-Host "$prefix$Text$suffix" -NoNewline }
        else { Write-Host "$prefix$Text$suffix" }
    }
    else {
        if ($NoNewline) { Write-Host $Text -ForegroundColor $fallback -NoNewline }
        else { Write-Host $Text -ForegroundColor $fallback }
    }
}

function Clear-AppScreen { Clear-Host }

function Show-AppHeader {
    param([Parameter(Mandatory = $true)][string]$Section)

    $width = 72
    try {
        $width = [Math]::Min(76, [Math]::Max(30, $Host.UI.RawUI.WindowSize.Width - 2))
    }
    catch { $width = 72 }

    Write-AppStyled -Text "$script:AppName  v$script:AppVersion" -Style Primary
    if ($Section -ne 'Inicio') { Write-AppStyled -Text $Section -Style Muted }
    Write-AppStyled -Text ('━' * $width) -Style Muted
}

function Write-AppOk {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-AppStyled -Text "✓ $Message" -Style Success
}
function Write-AppInfo {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-AppStyled -Text "● $Message" -Style Primary
}
function Write-AppWarning {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-AppStyled -Text "! $Message" -Style Warning
}
function Write-AppError {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-AppStyled -Text "× $Message" -Style Danger
}
function Write-AppMuted {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-AppStyled -Text $Message -Style Muted
}
function Wait-App {
    param([string]$Message = 'Enter para continuar')
    Write-Host ''
    [void](Read-Host $Message)
}
function Show-AppErrorScreen {
    param([Parameter(Mandatory = $true)][string]$Title,[Parameter(Mandatory = $true)][string]$Message)
    Clear-AppScreen
    Show-AppHeader -Section $Title
    Write-Host ''
    Write-AppError -Message $Message
    Wait-App
}

# ============================================================
# FORMATO
# ============================================================

function Format-AppBytes {
    param([Nullable[long]]$Bytes)
    if ($null -eq $Bytes -or $Bytes -lt 0) { return 'No disponible' }
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N2} KB' -f ($Bytes / 1KB)) }
    return "$Bytes bytes"
}

# ============================================================
# ENTRADAS Y VALIDACIONES
# ============================================================

function Read-AppYesNo {
    param([Parameter(Mandatory = $true)][string]$Prompt,[bool]$DefaultYes = $false)
    $hint = if ($DefaultYes) { 'S/n' } else { 's/N' }
    while ($true) {
        $answer = (Read-Host "$Prompt [$hint]").Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) { return $DefaultYes }
        if ($answer -match '^(s|si|sí|y|yes)$') { return $true }
        if ($answer -match '^(n|no)$') { return $false }
        Write-AppWarning -Message 'Responde S o N.'
    }
}

function Read-AppInteger {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][int]$Minimum,
        [Parameter(Mandatory = $true)][int]$Maximum,
        [Parameter(Mandatory = $true)][int]$DefaultValue
    )
    while ($true) {
        $raw = (Read-Host "$Prompt [$DefaultValue]").Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { return $DefaultValue }
        $number = 0
        if ([int]::TryParse($raw,[ref]$number) -and $number -ge $Minimum -and $number -le $Maximum) { return $number }
        Write-AppWarning -Message "Introduce un número entre $Minimum y $Maximum."
    }
}

function Test-AppGuid {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $guid = [Guid]::Empty
    return [Guid]::TryParse($Value,[ref]$guid)
}

function Test-AppTenant {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return $Value.Trim() -match '^[A-Za-z0-9-]+\.onmicrosoft\.com$'
}

function Read-AppTenant {
    param([string]$DefaultValue = '')
    while ($true) {
        $prompt = if ($DefaultValue) { "Tenant [$DefaultValue]" } else { 'Tenant (ej. contoso.onmicrosoft.com)' }
        $value = (Read-Host $prompt).Trim()
        if ([string]::IsNullOrWhiteSpace($value)) { $value = $DefaultValue }
        if (Test-AppTenant -Value $value) { return $value }
        Write-AppWarning -Message 'Formato esperado: nombre.onmicrosoft.com'
    }
}

function Read-AppM365Url {
    param([Parameter(Mandatory = $true)][string]$Prompt,[string]$DefaultUrl = '')
    while ($true) {
        $caption = if ($DefaultUrl) { "$Prompt [$DefaultUrl]" } else { $Prompt }
        $url = (Read-Host $caption).Trim()
        if ([string]::IsNullOrWhiteSpace($url)) { $url = $DefaultUrl }
        $uri = $null
        if ([Uri]::TryCreate($url,[UriKind]::Absolute,[ref]$uri) -and $uri.Scheme -eq 'https') {
            return $url.TrimEnd('/')
        }
        Write-AppWarning -Message 'URL HTTPS no válida.'
    }
}

# ============================================================
# MENU NUMERICO MODERNO
# ============================================================

function Show-NumberMenu {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][array]$Items,
        [string[]]$Description = @()
    )

    while ($true) {
        Clear-AppScreen
        Show-AppHeader -Section $Title

        if ($Description.Count -gt 0) {
            Write-Host ''
            foreach ($line in $Description) { Write-AppMuted -Message $line }
        }
        Write-Host ''

        $normalItems = @($Items | Where-Object { [string]$_.Value -notin @('Back','Exit') })
        $zeroItems = @($Items | Where-Object { [string]$_.Value -in @('Back','Exit') })
        $map = @{}

        for ($i=0; $i -lt $normalItems.Count; $i++) {
            $key = [string]($i + 1)
            $item = $normalItems[$i]
            $map[$key] = $item
            Write-AppStyled -Text ('{0,3}' -f $key) -Style Accent -NoNewline
            Write-Host "  $($item.Label)"
            $hintProp = $item.PSObject.Properties['Hint']
            if ($null -ne $hintProp -and -not [string]::IsNullOrWhiteSpace([string]$hintProp.Value)) {
                Write-AppMuted -Message "     $($hintProp.Value)"
            }
        }

        if ($zeroItems.Count -gt 0) {
            $item = $zeroItems[0]
            $map['0'] = $item
            Write-AppStyled -Text '  0' -Style Accent -NoNewline
            Write-Host "  $($item.Label)"
            $hintProp = $item.PSObject.Properties['Hint']
            if ($null -ne $hintProp -and -not [string]::IsNullOrWhiteSpace([string]$hintProp.Value)) {
                Write-AppMuted -Message "     $($hintProp.Value)"
            }
        }

        Write-Host ''
        $raw = (Read-Host 'Selecciona una opción').Trim()
        if ($map.ContainsKey($raw)) { return $map[$raw] }
        Write-AppError -Message 'Opción inválida.'
        Start-Sleep -Milliseconds 650
    }
}

function Select-AppSingleByNumber {
    param(
        [Parameter(Mandatory = $true)][array]$Items,
        [Parameter(Mandatory = $true)][scriptblock]$Label,
        [string]$Title = 'Seleccionar',
        [int]$MaxDisplay = 50
    )

    $list = @($Items)
    if ($list.Count -eq 0) { return $null }
    if ($list.Count -gt $MaxDisplay) {
        Write-AppWarning -Message "Hay $($list.Count) resultados; refina el filtro."
        return $null
    }

    Write-Host ''
    Write-AppStyled -Text $Title -Style Primary
    Write-Host ''
    for ($i=0; $i -lt $list.Count; $i++) {
        Write-AppStyled -Text ('{0,3}' -f ($i + 1)) -Style Accent -NoNewline
        Write-Host ('  ' + (& $Label $list[$i]))
    }
    Write-AppStyled -Text '  0' -Style Accent -NoNewline
    Write-Host '  Cancelar'

    while ($true) {
        Write-Host ''
        $raw = Read-Host 'Número'
        $number = 0
        if ([int]::TryParse($raw,[ref]$number)) {
            if ($number -eq 0) { return $null }
            if ($number -ge 1 -and $number -le $list.Count) { return $list[$number - 1] }
        }
        Write-AppError -Message 'Selección inválida.'
    }
}

# ============================================================
# SELECTOR VISUAL DE BIBLIOTECAS
# ============================================================

function Show-LibrarySelector {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Items)

    if ($Items.Count -eq 0) { return @() }

    $selected = @{}

    # Todas las bibliotecas comienzan desmarcadas.
    # El usuario selecciona explícitamente las que desea procesar.
    foreach ($item in $Items) {
        $selected[$item.Key] = $false
    }

    $index = 0

    while ($true) {
        Clear-AppScreen
        Show-AppHeader -Section 'Selección de bibliotecas'
        Write-Host ''
        Write-AppMuted -Message '↑/↓ mover   Espacio marcar   A todas   N ninguna   Enter aceptar   Esc cancelar'
        Write-Host ''

        for ($i=0; $i -lt $Items.Count; $i++) {
            $item = $Items[$i]
            $pointer = if ($i -eq $index) { '›' } else { ' ' }
            $mark = if ($selected[$item.Key]) { '●' } else { '○' }
            $line = "$pointer $mark  $($item.Label)"

            if ($i -eq $index) { Write-AppStyled -Text $line -Style Primary }
            elseif ($selected[$item.Key]) { Write-AppStyled -Text $line -Style Success }
            else { Write-Host $line }

            $hintProp = $item.PSObject.Properties['Hint']
            if ($null -ne $hintProp -and -not [string]::IsNullOrWhiteSpace([string]$hintProp.Value)) {
                Write-AppMuted -Message "      $($hintProp.Value)"
            }
        }

        $count = @($Items | Where-Object { $selected[$_.Key] }).Count
        Write-Host ''
        Write-AppStyled -Text "Seleccionadas: $count/$($Items.Count)" -Style Accent

        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        switch ($key.VirtualKeyCode) {
            38 { $index--; if ($index -lt 0) { $index = $Items.Count - 1 } }
            40 { $index++; if ($index -ge $Items.Count) { $index = 0 } }
            32 { $k = $Items[$index].Key; $selected[$k] = -not $selected[$k] }
            65 { foreach ($item in $Items) { $selected[$item.Key] = $true } }
            78 { foreach ($item in $Items) { $selected[$item.Key] = $false } }
            27 { return @() }
            13 {
                $result = foreach ($item in $Items) { if ($selected[$item.Key]) { $item.Object } }
                if (@($result).Count -gt 0) { return @($result) }
                Write-AppWarning -Message 'Marca al menos una biblioteca.'
                Start-Sleep -Milliseconds 650
            }
        }
    }
}

# ============================================================
# DIRECTORIOS
# ============================================================

function Initialize-AppFolders {
    if (-not (Test-Path -Path $script:AppRoot)) {
        New-Item -Path $script:AppRoot -ItemType Directory -Force | Out-Null
    }

    if (-not (Test-Path -Path $script:ReportsPath)) {
        New-Item -Path $script:ReportsPath -ItemType Directory -Force | Out-Null
    }
}


# ============================================================
# CONFIGURACION
# ============================================================

function Get-DefaultAppConfig {
    return [PSCustomObject]@{
        ClientId = ""
        Tenant = ""
        AppRegistrationName = "Microsoft 365 Version History Cleaner"
        LastSharePointUrl = ""
        LastOneDriveUrl = ""
        VersionsToKeep = 10
        CountMode = "Historical"
        PersistLogin = $false

        RequestDelayMs = 150
        MaxRetries = 8
        RetryBaseSeconds = 5
        RetryMaxSeconds = 120
    }
}

function Get-AppConfig {
    $default = Get-DefaultAppConfig

    if (-not (Test-Path -Path $script:ConfigPath)) {
        return $default
    }

    try {
        $saved = Get-Content -Path $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

        foreach ($property in $default.PSObject.Properties) {
            $name = $property.Name

            if ($null -eq $saved.PSObject.Properties[$name]) {
                $saved | Add-Member -NotePropertyName $name -NotePropertyValue $property.Value
            }
        }

        return $saved
    }
    catch {
        return $default
    }
}

function Save-AppConfig {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    Initialize-AppFolders

    $Config | ConvertTo-Json -Depth 5 | Set-Content -Path $script:ConfigPath -Encoding UTF8
}


# ============================================================
# THROTTLING / RETRIES
# ============================================================

function Test-AppThrottleException {
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception
    )

    $text = $Exception.ToString()

    $patterns = @(
        "429"
        "Too Many Requests"
        "throttl"
        "503"
        "Server Too Busy"
        "Server is busy"
        "Service Unavailable"
        "temporarily unavailable"
    )

    foreach ($pattern in $patterns) {
        if ($text -match [regex]::Escape($pattern)) {
            return $true
        }
    }

    return $false
}

function Get-AppRetryAfterSeconds {
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception
    )

    $current = $Exception

    while ($null -ne $current) {
        try {
            if ($null -ne $current.PSObject.Properties["Response"]) {
                $response = $current.Response

                if ($null -ne $response) {
                    if ($null -ne $response.PSObject.Properties["Headers"]) {
                        $headers = $response.Headers

                        if ($null -ne $headers) {
                            try {
                                $retryAfter = $headers["Retry-After"]

                                if ($null -ne $retryAfter) {
                                    $seconds = 0

                                    if ([int]::TryParse([string]$retryAfter, [ref]$seconds)) {
                                        if ($seconds -gt 0) {
                                            return $seconds
                                        }
                                    }
                                }
                            }
                            catch {
                            }
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
        [Parameter(Mandatory = $true)]
        [int]$Attempt,

        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception
    )

    $config = Get-AppConfig

    $retryAfter = Get-AppRetryAfterSeconds -Exception $Exception

    if ($null -ne $retryAfter) {
        return [Math]::Min(
            [int]$retryAfter,
            [int]$config.RetryMaxSeconds
        )
    }

    $base = [int]$config.RetryBaseSeconds
    $max = [int]$config.RetryMaxSeconds

    $exponential = $base * [Math]::Pow(2, $Attempt - 1)
    $jitter = Get-Random -Minimum 0 -Maximum 4

    $delay = [int][Math]::Ceiling($exponential + $jitter)

    return [Math]::Min($delay, $max)
}

function Invoke-AppWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$OperationName,

        [switch]$NoPacing
    )

    $config = Get-AppConfig
    $maxRetries = [int]$config.MaxRetries

    for ($attempt = 1; $attempt -le ($maxRetries + 1); $attempt++) {
        if (-not $NoPacing) {
            $delayMs = [int]$config.RequestDelayMs

            if ($delayMs -gt 0) {
                Start-Sleep -Milliseconds $delayMs
            }
        }

        try {
            return & $Operation
        }
        catch {
            $exception = $_.Exception
            $isThrottle = Test-AppThrottleException -Exception $exception

            if (-not $isThrottle) {
                throw
            }

            if ($attempt -gt $maxRetries) {
                throw
            }

            $waitSeconds = Get-AppRetryDelay -Attempt $attempt -Exception $exception

            $script:ThrottleEvents.Add(
                [PSCustomObject]@{
                    FechaHora = Get-Date
                    Operacion = $OperationName
                    Intento = $attempt
                    EsperaSegundos = $waitSeconds
                    Mensaje = $exception.Message
                }
            )

            Write-Host ""
            Write-AppWarning -Message "SharePoint/OneDrive aplicó throttling."
            Write-AppMuted -Message "Operación: $OperationName"
            Write-AppMuted -Message "Reintento $attempt de $maxRetries en $waitSeconds segundo(s)."

            Start-Sleep -Seconds $waitSeconds
        }
    }
}


# ============================================================
# PNP.POWERSHELL
# ============================================================

function Get-PnPInstalledVersion {
    $module = Get-Module -ListAvailable -Name "PnP.PowerShell" |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -ne $module) {
        return $module.Version
    }

    return $null
}

function Install-AppPnP {
    Clear-AppScreen
    Show-AppHeader -Section "PnP.PowerShell"

    $version = Get-PnPInstalledVersion

    if ($null -ne $version) {
        Write-AppOk -Message "PnP.PowerShell $version está instalado."
        Write-Host ""

        if (-not (Read-AppYesNo -Prompt "¿Desea actualizarlo?" -DefaultYes $false)) {
            return $true
        }
    }
    else {
        Write-AppWarning -Message "PnP.PowerShell no está instalado."
        Write-Host ""

        if (-not (Read-AppYesNo -Prompt "¿Desea instalarlo?" -DefaultYes $true)) {
            return $false
        }
    }

    try {
        $params = @{
            Name = "PnP.PowerShell"
            Scope = "CurrentUser"
            Force = $true
            AllowClobber = $true
            ErrorAction = "Stop"
        }

        Install-Module @params
        Import-Module "PnP.PowerShell" -Force -ErrorAction Stop

        Write-AppOk -Message "PnP.PowerShell está listo."
        Wait-App

        return $true
    }
    catch {
        Show-AppErrorScreen -Title "PnP.PowerShell" -Message $_.Exception.Message
        return $false
    }
}

function Confirm-AppPnP {
    if ($null -eq (Get-PnPInstalledVersion)) {
        return Install-AppPnP
    }

    try {
        Import-Module "PnP.PowerShell" -ErrorAction Stop
        return $true
    }
    catch {
        Show-AppErrorScreen -Title "PnP.PowerShell" -Message $_.Exception.Message
        return $false
    }
}


# ============================================================
# AUTENTICACION
# ============================================================

# ============================================================
# AUTENTICACION / APP REGISTRATION
# ============================================================

function Get-AppAdminUrl {
    param([Parameter(Mandatory = $true)][string]$Tenant)
    if (-not (Test-AppTenant -Value $Tenant)) { throw 'Tenant no válido.' }
    $prefix = ($Tenant -split '\.')[0]
    return "https://$prefix-admin.sharepoint.com"
}

function New-AppPnPConnection {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$Tenant,
        [bool]$PersistLogin = $false
    )

    $params = @{
        Url = $Url
        ClientId = $ClientId
        Tenant = $Tenant
        Interactive = $true
        ReturnConnection = $true
        ErrorAction = 'Stop'
    }
    if ($PersistLogin) { $params.PersistLogin = $true }
    return Connect-PnPOnline @params
}

function Test-AppRegistration {
    param(
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$Tenant,
        [switch]$Quiet
    )

    if (-not (Test-AppGuid -Value $ClientId)) {
        if (-not $Quiet) { Write-AppError -Message 'El Client ID no es un GUID válido.' }
        return $false
    }
    if (-not (Test-AppTenant -Value $Tenant)) {
        if (-not $Quiet) { Write-AppError -Message 'El tenant no tiene formato nombre.onmicrosoft.com.' }
        return $false
    }

    try {
        $adminUrl = Get-AppAdminUrl -Tenant $Tenant
        if (-not $Quiet) { Write-AppInfo -Message "Validando contra $adminUrl" }
        $conn = New-AppPnPConnection -Url $adminUrl -ClientId $ClientId -Tenant $Tenant -PersistLogin:$false
        $null = Get-PnPWeb -Connection $conn -Includes Title,Url -ErrorAction Stop
        if (-not $Quiet) { Write-AppOk -Message 'La aplicación pudo autenticarse correctamente.' }
        return $true
    }
    catch {
        if (-not $Quiet) {
            Write-AppError -Message 'No fue posible validar la aplicación.'
            Write-AppMuted -Message $_.Exception.Message
        }
        return $false
    }
}

function Set-AppExistingClientId {
    $config = Get-AppConfig
    Clear-AppScreen
    Show-AppHeader -Section 'Usar aplicación existente'
    Write-Host ''

    $tenant = Read-AppTenant -DefaultValue ([string]$config.Tenant)

    while ($true) {
        $clientId = (Read-Host 'Client ID de la aplicación existente').Trim()
        if (Test-AppGuid -Value $clientId) { break }
        Write-AppWarning -Message 'Client ID no válido.'
    }

    Write-Host ''
    $validate = Read-AppYesNo -Prompt '¿Validar la aplicación antes de guardarla?' -DefaultYes $true
    if ($validate -and -not (Test-AppRegistration -ClientId $clientId -Tenant $tenant)) {
        Write-Host ''
        Write-AppWarning -Message 'La configuración anterior se conservará.'
        return $false
    }

    $config.ClientId = $clientId
    $config.Tenant = $tenant
    $config.PersistLogin = Read-AppYesNo -Prompt '¿Mantener sesión autenticada?' -DefaultYes ([bool]$config.PersistLogin)
    Save-AppConfig -Config $config
    Write-AppOk -Message 'Aplicación configurada.'
    return $true
}

function New-AppEntraRegistration {
    $config = Get-AppConfig
    Clear-AppScreen
    Show-AppHeader -Section 'Registrar nueva aplicación Entra'
    Write-Host ''
    Write-AppMuted -Message 'Permiso delegado solicitado:'
    Write-AppStyled -Text 'AllSites.FullControl' -Style Warning
    Write-Host ''

    $tenant = Read-AppTenant -DefaultValue ([string]$config.Tenant)
    $defaultName = if ([string]::IsNullOrWhiteSpace([string]$config.AppRegistrationName)) { $script:AppName } else { [string]$config.AppRegistrationName }
    $name = (Read-Host "Nombre de la aplicación [$defaultName]").Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $defaultName }

    Write-Host ''
    if (-not (Read-AppYesNo -Prompt '¿Registrar esta aplicación?' -DefaultYes $false)) { return $false }

    try {
        $result = Register-PnPEntraIDAppForInteractiveLogin `
            -ApplicationName $name `
            -Tenant $tenant `
            -SharePointDelegatePermissions @('AllSites.FullControl') `
            -ErrorAction Stop

        $clientId = ''
        foreach ($propertyName in @('ClientId','AppId','ApplicationId','Id')) {
            $prop = $result.PSObject.Properties[$propertyName]
            if ($null -ne $prop -and (Test-AppGuid -Value ([string]$prop.Value))) {
                $clientId = [string]$prop.Value
                break
            }
        }

        if (-not (Test-AppGuid -Value $clientId)) {
            while ($true) {
                $clientId = (Read-Host 'Client ID creado').Trim()
                if (Test-AppGuid -Value $clientId) { break }
                Write-AppWarning -Message 'Client ID no válido.'
            }
        }

        $config.ClientId = $clientId
        $config.Tenant = $tenant
        $config.AppRegistrationName = $name
        Save-AppConfig -Config $config

        Write-Host ''
        Write-AppOk -Message "Aplicación registrada: $name"
        Write-AppMuted -Message $clientId

        if (Read-AppYesNo -Prompt '¿Validar la nueva aplicación ahora?' -DefaultYes $true) {
            [void](Test-AppRegistration -ClientId $clientId -Tenant $tenant)
        }
        return $true
    }
    catch {
        Write-AppError -Message "No se pudo registrar la aplicación: $($_.Exception.Message)"
        return $false
    }
}

function Remove-AppPersistedLogin {
    try {
        Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction Stop
        Write-AppOk -Message 'Sesión persistente eliminada.'
    }
    catch { Write-AppWarning -Message $_.Exception.Message }
}

function Show-AppAuthenticationMenu {
    while ($true) {
        $config = Get-AppConfig
        $configured = (Test-AppGuid -Value ([string]$config.ClientId)) -and (Test-AppTenant -Value ([string]$config.Tenant))

        $items = @()
        if ($configured) {
            $items += [PSCustomObject]@{ Label='Validar aplicación configurada'; Value='Validate'; Hint='Prueba autenticación real contra este tenant' }
        }
        $items += [PSCustomObject]@{ Label=$(if ($configured) {'Usar otra aplicación existente'} else {'Usar una aplicación existente'}); Value='Existing'; Hint='Introduce un Client ID ya registrado' }
        $items += [PSCustomObject]@{ Label='Registrar una nueva aplicación Entra'; Value='Register'; Hint='Crea una app PnP nueva' }
        if ($configured) {
            $items += [PSCustomObject]@{ Label='Quitar Client ID de la configuración local'; Value='RemoveLocal'; Hint='No elimina la app de Entra' }
        }
        $items += [PSCustomObject]@{ Label='Eliminar sesión persistente'; Value='Clear'; Hint='Limpia el token persistido de PnP' }
        $items += [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='' }

        $desc = @(
            "Tenant: $(if ($config.Tenant) {$config.Tenant} else {'—'})",
            "Client ID: $(if ($config.ClientId) {$config.ClientId} else {'—'})",
            "Estado: $(if ($configured) {'CONFIGURADA'} else {'NO CONFIGURADA'})"
        )
        $choice = Show-NumberMenu -Title 'Aplicación Entra / autenticación PnP' -Items $items -Description $desc

        switch ($choice.Value) {
            'Validate' { [void](Test-AppRegistration -ClientId ([string]$config.ClientId) -Tenant ([string]$config.Tenant)); Wait-App }
            'Existing' { [void](Set-AppExistingClientId); Wait-App }
            'Register' { [void](New-AppEntraRegistration); Wait-App }
            'RemoveLocal' {
                if (Read-AppYesNo -Prompt '¿Quitar el Client ID guardado?' -DefaultYes $false) {
                    $config.ClientId = ''
                    Save-AppConfig -Config $config
                    Write-AppOk -Message 'Client ID eliminado de la configuración local.'
                }
                Wait-App
            }
            'Clear' { Remove-AppPersistedLogin; Wait-App }
            'Back' { return }
        }
    }
}

function Confirm-AppAuthenticationConfigured {
    $config = Get-AppConfig
    if ((Test-AppGuid -Value ([string]$config.ClientId)) -and (Test-AppTenant -Value ([string]$config.Tenant))) { return $true }

    Clear-AppScreen
    Show-AppHeader -Section 'Autenticación requerida'
    Write-Host ''
    Write-AppWarning -Message 'No hay una aplicación Entra válida configurada.'
    $items = @(
        [PSCustomObject]@{ Label='Usar una aplicación existente'; Value='Existing'; Hint='Client ID ya registrado' },
        [PSCustomObject]@{ Label='Registrar una nueva aplicación Entra'; Value='Register'; Hint='Crear mediante PnP.PowerShell' },
        [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='' }
    )
    $choice = Show-NumberMenu -Title 'Autenticación requerida' -Items $items
    switch ($choice.Value) {
        'Existing' { return [bool](Set-AppExistingClientId) }
        'Register' { return [bool](New-AppEntraRegistration) }
        default { return $false }
    }
}

# ============================================================
# DESCUBRIMIENTO Y SELECCION DE DESTINO
# ============================================================

function Get-AppTenantSites {
    param([ValidateSet('SharePoint','OneDrive')][string]$Kind)

    $config = Get-AppConfig
    if (-not (Test-AppGuid -Value ([string]$config.ClientId))) { throw 'No existe un Client ID configurado.' }
    if (-not (Test-AppTenant -Value ([string]$config.Tenant))) { throw 'No existe un tenant válido configurado.' }

    $adminUrl = Get-AppAdminUrl -Tenant ([string]$config.Tenant)
    Write-AppInfo -Message "Conectando al centro de administración: $adminUrl"
    $conn = New-AppPnPConnection -Url $adminUrl -ClientId ([string]$config.ClientId) -Tenant ([string]$config.Tenant) -PersistLogin:([bool]$config.PersistLogin)

    Write-AppInfo -Message 'Consultando sitios del tenant...'
    $sites = @(Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Connection $conn -ErrorAction Stop)

    if ($Kind -eq 'OneDrive') {
        return @($sites | Where-Object { [string]$_.Url -match '-my\.sharepoint\.com/personal/' } | Sort-Object Owner,Url)
    }

    return @($sites | Where-Object {
        [string]$_.Url -notmatch '-my\.sharepoint\.com/personal/' -and
        [string]$_.Url -notmatch '-admin\.sharepoint\.com/?$'
    } | Sort-Object Title,Url)
}

function Find-AppSiteInteractively {
    param([ValidateSet('SharePoint','OneDrive')][string]$Kind)

    try { $sites = @(Get-AppTenantSites -Kind $Kind) }
    catch {
        Show-AppErrorScreen -Title 'Buscar sitios' -Message $_.Exception.Message
        return $null
    }

    if ($sites.Count -eq 0) {
        Write-AppWarning -Message 'No se encontraron sitios accesibles.'
        Wait-App
        return $null
    }

    while ($true) {
        Clear-AppScreen
        Show-AppHeader -Section $(if ($Kind -eq 'OneDrive') {'Buscar OneDrive'} else {'Buscar SharePoint'})
        Write-Host ''
        Write-AppMuted -Message "Sitios disponibles: $($sites.Count)"
        Write-AppMuted -Message 'Filtra por nombre, URL o propietario. Deja vacío para mostrar todos si son 50 o menos.'
        Write-Host ''
        $filter = (Read-Host 'Filtro').Trim()
        $filtered = $sites

        if ($filter) {
            $needle = [regex]::Escape($filter)
            $filtered = @($sites | Where-Object {
                ([string]$_.Title -match "(?i)$needle") -or
                ([string]$_.Url -match "(?i)$needle") -or
                ([string]$_.Owner -match "(?i)$needle")
            })
        }

        if ($filtered.Count -eq 0) {
            Write-AppWarning -Message 'Sin coincidencias.'
            Start-Sleep -Milliseconds 700
            continue
        }
        if ($filtered.Count -gt 50) {
            Write-AppWarning -Message "Hay $($filtered.Count) coincidencias. Refina el filtro."
            Start-Sleep -Milliseconds 900
            continue
        }

        $selected = Select-AppSingleByNumber -Items $filtered -Title $(if ($Kind -eq 'OneDrive') {'OneDrive'} else {'Sitios SharePoint'}) -Label {
            param($s)
            $title = if (-not [string]::IsNullOrWhiteSpace([string]$s.Title)) { [string]$s.Title } elseif (-not [string]::IsNullOrWhiteSpace([string]$s.Owner)) { [string]$s.Owner } else { '(sin título)' }
            $owner = if (-not [string]::IsNullOrWhiteSpace([string]$s.Owner)) { " · $($s.Owner)" } else { '' }
            "$title$owner`n     $($s.Url)"
        }
        return $selected
    }
}

function Select-AppSharePointTarget {
    $config = Get-AppConfig
    $items = @(
        [PSCustomObject]@{ Label='Buscar sitio en el tenant'; Value='Search'; Hint='Busca por nombre, URL o propietario' },
        [PSCustomObject]@{ Label='Introducir URL directamente'; Value='Direct'; Hint='Para un sitio ya conocido' }
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$config.LastSharePointUrl)) {
        $items += [PSCustomObject]@{ Label='Usar último sitio'; Value='Last'; Hint=[string]$config.LastSharePointUrl }
    }
    $items += [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='' }

    $choice = Show-NumberMenu -Title 'SharePoint Online' -Items $items
    switch ($choice.Value) {
        'Search' {
            $site = Find-AppSiteInteractively -Kind SharePoint
            if ($null -eq $site) { return $null }
            return [PSCustomObject]@{ Type='SharePoint'; Url=[string]$site.Url; Title=[string]$site.Title; Owner=[string]$site.Owner }
        }
        'Direct' {
            $url = Read-AppM365Url -Prompt 'URL del sitio'
            return [PSCustomObject]@{ Type='SharePoint'; Url=$url; Title=''; Owner='' }
        }
        'Last' { return [PSCustomObject]@{ Type='SharePoint'; Url=[string]$config.LastSharePointUrl; Title=''; Owner='' } }
        default { return $null }
    }
}

function Select-AppOneDriveTarget {
    $config = Get-AppConfig
    $items = @(
        [PSCustomObject]@{ Label='Buscar OneDrive en el tenant'; Value='Search'; Hint='Busca por propietario o URL real' },
        [PSCustomObject]@{ Label='Introducir URL directamente'; Value='Direct'; Hint='Para un OneDrive ya conocido' }
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$config.LastOneDriveUrl)) {
        $items += [PSCustomObject]@{ Label='Usar último OneDrive'; Value='Last'; Hint=[string]$config.LastOneDriveUrl }
    }
    $items += [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='' }

    $choice = Show-NumberMenu -Title 'OneDrive for Business' -Items $items
    switch ($choice.Value) {
        'Search' {
            $site = Find-AppSiteInteractively -Kind OneDrive
            if ($null -eq $site) { return $null }
            return [PSCustomObject]@{ Type='OneDrive'; Url=[string]$site.Url; Title=[string]$site.Title; Owner=[string]$site.Owner }
        }
        'Direct' {
            $url = Read-AppM365Url -Prompt 'URL de OneDrive'
            return [PSCustomObject]@{ Type='OneDrive'; Url=$url; Title=''; Owner='' }
        }
        'Last' { return [PSCustomObject]@{ Type='OneDrive'; Url=[string]$config.LastOneDriveUrl; Title=''; Owner='' } }
        default { return $null }
    }
}

function Select-AppTarget {
    $items = @(
        [PSCustomObject]@{ Label='SharePoint Online'; Value='SharePoint'; Hint='Buscar en el tenant o usar URL directa' },
        [PSCustomObject]@{ Label='OneDrive for Business'; Value='OneDrive'; Hint='Buscar OneDrive reales del tenant' },
        [PSCustomObject]@{ Label='URL directa'; Value='Direct'; Hint='Detecta automáticamente SPO u OneDrive por la URL' },
        [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='' }
    )

    $choice = Show-NumberMenu -Title 'Nueva limpieza' -Items $items -Description @('Selecciona el origen que quieres analizar.')
    switch ($choice.Value) {
        'SharePoint' { return Select-AppSharePointTarget }
        'OneDrive' { return Select-AppOneDriveTarget }
        'Direct' {
            Clear-AppScreen
            Show-AppHeader -Section 'URL directa'
            Write-Host ''
            $url = Read-AppM365Url -Prompt 'URL del sitio / OneDrive'
            $type = if ($url -match '-my\.sharepoint\.com/personal/') { 'OneDrive' } else { 'SharePoint' }
            return [PSCustomObject]@{ Type=$type; Url=$url; Title=''; Owner='' }
        }
        default { return $null }
    }
}

# ============================================================
# CONEXION
# ============================================================

function Connect-AppM365 {
    param([Parameter(Mandatory = $true)][object]$Target)

    $config = Get-AppConfig
    if (-not (Test-AppGuid -Value ([string]$config.ClientId))) { throw 'No existe un Client ID configurado.' }
    if (-not (Test-AppTenant -Value ([string]$config.Tenant))) { throw 'No existe un tenant válido configurado.' }

    $script:PnPConnection = New-AppPnPConnection `
        -Url ([string]$Target.Url) `
        -ClientId ([string]$config.ClientId) `
        -Tenant ([string]$config.Tenant) `
        -PersistLogin:([bool]$config.PersistLogin)

    if ($null -eq $script:PnPConnection) { throw 'PnP no devolvió una conexión válida.' }

    return Get-PnPWeb -Connection $script:PnPConnection -Includes Title,Url -ErrorAction Stop
}

function Disconnect-AppM365 {
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue }
    catch { }
    $script:PnPConnection = $null
}

function Get-AppDocumentLibraries {
    $operation = {
        $params = @{
            Connection = $script:PnPConnection
            Includes = @(
                "RootFolder"
                "BaseType"
                "BaseTemplate"
                "Hidden"
                "EnableVersioning"
                "ItemCount"
            )
            ErrorAction = "Stop"
        }

        @(Get-PnPList @params)
    }

    $lists = Invoke-AppWithRetry -Operation $operation -OperationName "Obtener bibliotecas"

    return @(
        $lists |
            Where-Object {
                $_.BaseType.ToString() -eq "DocumentLibrary" -and
                -not $_.Hidden
            } |
            Sort-Object Title
    )
}

function Select-AppLibraries {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Libraries
    )

    $items = @()

    for ($i = 0; $i -lt $Libraries.Count; $i++) {
        $library = $Libraries[$i]

        if ($library.BaseTemplate -eq 700) {
            $template = "OneDrive / MySite (700)"
        }
        elseif ($library.BaseTemplate -eq 101) {
            $template = "Biblioteca estándar (101)"
        }
        else {
            $template = "Plantilla $($library.BaseTemplate)"
        }

        if ($library.EnableVersioning) {
            $versioning = "Versionado activo"
        }
        else {
            $versioning = "Versionado desactivado"
        }

        $items += [PSCustomObject]@{
            Key = [string]$i
            Label = $library.Title
            Hint = "$template | $versioning | Items: $($library.ItemCount)"
            Object = $library
        }
    }

    return @(Show-LibrarySelector -Items $items)
}


# ============================================================
# ENUMERACION RECURSIVA
# ============================================================

function Get-AppLibraryFiles {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Library
    )

    $caml = @"
<View Scope='RecursiveAll'>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='FSObjType' />
                <Value Type='Integer'>0</Value>
            </Eq>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='FileRef' />
        <FieldRef Name='FileLeafRef' />
        <FieldRef Name='FSObjType' />
        <FieldRef Name='File_x0020_Size' />
    </ViewFields>
    <RowLimit Paged='TRUE'>500</RowLimit>
</View>
"@

    $operation = {
        $params = @{
            List = $Library
            Query = $caml
            PageSize = 500
            Connection = $script:PnPConnection
            ErrorAction = "Stop"
        }

        @(Get-PnPListItem @params)
    }

    $items = Invoke-AppWithRetry -Operation $operation -OperationName "Enumerar archivos de $($Library.Title)"

    $files = @()

    foreach ($item in $items) {
        if ($item.FieldValues.ContainsKey("FileRef")) {
            $fileRef = [string]$item.FieldValues["FileRef"]

            if (-not [string]::IsNullOrWhiteSpace($fileRef)) {
                $files += $item
            }
        }
    }

    return $files
}


# ============================================================
# VERSIONES
# ============================================================

function Get-AppFileVersions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileUrl,

        [string]$OperationName = "Obtener versiones"
    )

    $operation = {
        $params = @{
            Url = $FileUrl
            Connection = $script:PnPConnection
            ErrorAction = "Stop"
        }

        @(Get-PnPFileVersion @params | Sort-Object Created -Descending)
    }

    return @(
        Invoke-AppWithRetry -Operation $operation -OperationName $OperationName
    )
}

function Get-AppVersionSize {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Version
    )

    foreach ($propertyName in @("Size", "Length", "FileSize")) {
        if ($null -ne $Version.PSObject.Properties[$propertyName]) {
            $value = $Version.$propertyName

            if ($null -ne $value) {
                $size = 0L

                if ([long]::TryParse([string]$value, [ref]$size)) {
                    if ($size -ge 0) {
                        return $size
                    }
                }
            }
        }
    }

    return $null
}


# ============================================================
# RETENCION
# ============================================================

function Select-AppRetention {
    $config = Get-AppConfig

    $items = @(
        [PSCustomObject]@{
            Label = "Conservar N históricas + versión actual"
            Value = "Historical"
            Hint = "Ejemplo: 10 = actual + 10 históricas"
        },
        [PSCustomObject]@{
            Label = "Conservar N versiones totales"
            Value = "Total"
            Hint = "Ejemplo: 10 = actual + 9 históricas"
        },
        [PSCustomObject]@{
            Label = "Volver"
            Value = "Back"
            Hint = ""
        }
    )

    $choice = Show-NumberMenu -Title "Retención" -Items $items

    if ($choice.Value -eq "Back") {
        return $null
    }

    Clear-AppScreen
    Show-AppHeader -Section "Retención"

    $number = Read-AppInteger -Prompt "Número de versiones a conservar" -Minimum 1 -Maximum 50000 -DefaultValue ([int]$config.VersionsToKeep)

    if ($choice.Value -eq "Historical") {
        $historicalToKeep = $number

        Write-AppOk -Message "Versión actual: conservar"
        Write-AppOk -Message "$number versiones históricas: conservar"
    }
    else {
        $historicalToKeep = [Math]::Max(0, $number - 1)

        Write-AppOk -Message "$number versiones totales"
        Write-AppMuted -Message "Actual + $historicalToKeep histórica(s)"
    }

    Write-Host ""
    Write-AppWarning -Message "Las versiones históricas más antiguas serán elegibles para limpieza."
    Write-Host ""

    if (-not (Read-AppYesNo -Prompt "¿Continuar?" -DefaultYes $true)) {
        return $null
    }

    $config.VersionsToKeep = $number
    $config.CountMode = $choice.Value

    Save-AppConfig -Config $config

    return [PSCustomObject]@{
        Requested = $number
        HistoricalToKeep = $historicalToKeep
        Mode = $choice.Value
    }
}


# ============================================================
# ANALISIS
# ============================================================

function New-AppAnalysisResult {
    return [PSCustomObject]@{
        Start = Get-Date
        End = $null

        Libraries = 0
        Files = 0
        FilesWithHistory = 0
        FilesAffected = 0

        VersionsFound = 0
        VersionsKept = 0
        VersionsEligible = 0

        EstimatedBytes = 0L
        VersionsWithKnownSize = 0
        VersionsWithoutKnownSize = 0

        Errors = 0

        Details = [System.Collections.Generic.List[object]]::new()
        ByLibrary = [System.Collections.Generic.List[object]]::new()
        ByFile = [System.Collections.Generic.List[object]]::new()
    }
}

function Invoke-AppVersionAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Libraries,

        [Parameter(Mandatory = $true)]
        [int]$HistoricalToKeep
    )

    $result = New-AppAnalysisResult

    for ($libraryIndex = 0; $libraryIndex -lt $Libraries.Count; $libraryIndex++) {
        $library = $Libraries[$libraryIndex]
        $result.Libraries++

        Clear-AppScreen
        Show-AppHeader -Section "Analizando"

        Write-Host $library.Title -ForegroundColor Cyan
        Write-AppMuted -Message "Biblioteca $($libraryIndex + 1) de $($Libraries.Count)"
        Write-Host ""

        $libraryFiles = 0
        $libraryVersions = 0
        $libraryAffected = 0
        $libraryEligible = 0
        $libraryErrors = 0
        $libraryBytes = 0L

        try {
            $files = @(Get-AppLibraryFiles -Library $library)
            Write-AppOk -Message "$($files.Count) archivo(s) encontrado(s)."
        }
        catch {
            $result.Errors++
            $libraryErrors++

            $result.ByLibrary.Add(
                [PSCustomObject]@{
                    Biblioteca = $library.Title
                    Plantilla = $library.BaseTemplate
                    Archivos = 0
                    Versiones = 0
                    ArchivosAfectados = 0
                    VersionesEliminar = 0
                    EspacioEstimadoBytes = 0
                    Errores = 1
                }
            )

            continue
        }

        for ($fileIndex = 0; $fileIndex -lt $files.Count; $fileIndex++) {
            $item = $files[$fileIndex]
            $fileUrl = [string]$item.FieldValues["FileRef"]

            $result.Files++
            $libraryFiles++

            $percent = [int]((($fileIndex + 1) / [Math]::Max(1, $files.Count)) * 100)

            Write-Progress -Id 1 -Activity "Analizando $($library.Title)" -Status "$($fileIndex + 1) de $($files.Count) - $fileUrl" -PercentComplete $percent

            try {
                $versions = @(Get-AppFileVersions -FileUrl $fileUrl -OperationName "Versiones de $fileUrl")

                if ($versions.Count -eq 0) {
                    continue
                }

                $result.FilesWithHistory++
                $result.VersionsFound += $versions.Count
                $libraryVersions += $versions.Count

                $keepCount = [Math]::Min($HistoricalToKeep, $versions.Count)
                $result.VersionsKept += $keepCount

                $toRemove = @(
                    $versions | Select-Object -Skip $HistoricalToKeep
                )

                if ($toRemove.Count -eq 0) {
                    continue
                }

                $result.FilesAffected++
                $libraryAffected++

                $fileBytes = 0L
                $fileKnownSizes = 0

                foreach ($version in $toRemove) {
                    $result.VersionsEligible++
                    $libraryEligible++

                    $size = Get-AppVersionSize -Version $version

                    if ($null -ne $size) {
                        $result.EstimatedBytes += $size
                        $libraryBytes += $size
                        $fileBytes += $size
                        $result.VersionsWithKnownSize++
                        $fileKnownSizes++
                    }
                    else {
                        $result.VersionsWithoutKnownSize++
                    }

                    $result.Details.Add(
                        [PSCustomObject]@{
                            Biblioteca = $library.Title
                            Archivo = $fileUrl
                            VersionId = $version.Id
                            Version = $version.VersionLabel
                            Creada = $version.Created
                            TamanoBytes = $size
                            Accion = "ELIMINAR"
                            Estado = "Vista previa"
                            Mensaje = ""
                        }
                    )
                }

                $result.ByFile.Add(
                    [PSCustomObject]@{
                        Biblioteca = $library.Title
                        Archivo = $fileUrl
                        VersionesHistoricas = $versions.Count
                        HistoricasConservar = $keepCount
                        VersionesEliminar = $toRemove.Count
                        EspacioEstimadoBytes = $fileBytes
                        TamanosDisponibles = $fileKnownSizes
                    }
                )
            }
            catch {
                $result.Errors++
                $libraryErrors++

                $result.Details.Add(
                    [PSCustomObject]@{
                        Biblioteca = $library.Title
                        Archivo = $fileUrl
                        VersionId = ""
                        Version = ""
                        Creada = ""
                        TamanoBytes = $null
                        Accion = "ERROR"
                        Estado = "Error de análisis"
                        Mensaje = $_.Exception.Message
                    }
                )
            }
        }

        Write-Progress -Id 1 -Activity "Analizando" -Completed

        $result.ByLibrary.Add(
            [PSCustomObject]@{
                Biblioteca = $library.Title
                Plantilla = $library.BaseTemplate
                Archivos = $libraryFiles
                Versiones = $libraryVersions
                ArchivosAfectados = $libraryAffected
                VersionesEliminar = $libraryEligible
                EspacioEstimadoBytes = $libraryBytes
                Errores = $libraryErrors
            }
        )
    }

    $result.End = Get-Date
    return $result
}


# ============================================================
# REPORTES DE ANALISIS
# ============================================================

function Export-AppAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Analysis
    )

    Initialize-AppFolders

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"

    $detailsPath = Join-Path $script:ReportsPath "Analisis-Versiones-$stamp.csv"
    $librariesPath = Join-Path $script:ReportsPath "Resumen-Bibliotecas-$stamp.csv"
    $filesPath = Join-Path $script:ReportsPath "Resumen-Archivos-$stamp.csv"

    $Analysis.Details | Export-Csv -Path $detailsPath -NoTypeInformation -Encoding UTF8
    $Analysis.ByLibrary | Export-Csv -Path $librariesPath -NoTypeInformation -Encoding UTF8
    $Analysis.ByFile | Export-Csv -Path $filesPath -NoTypeInformation -Encoding UTF8

    return [PSCustomObject]@{
        Details = $detailsPath
        Libraries = $librariesPath
        Files = $filesPath
    }
}


# ============================================================
# RESUMEN
# ============================================================

function Show-AppAnalysisSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Analysis,

        [Parameter(Mandatory = $true)]
        [object]$Target
    )

    Clear-AppScreen
    Show-AppHeader -Section "Resultado del análisis"

    Write-AppMuted -Message $Target.Type
    Write-Host $Target.Url -ForegroundColor Cyan
    Write-Host ""

    $duration = $Analysis.End - $Analysis.Start

    Write-Host ("Bibliotecas analizadas".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.Libraries)

    Write-Host ("Archivos analizados".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.Files)

    Write-Host ("Con historial".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.FilesWithHistory)

    Write-Host ("Archivos afectados".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.FilesAffected)

    Write-Host ""

    Write-Host ("Versiones históricas".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.VersionsFound)

    Write-Host ("Históricas a conservar".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.VersionsKept) -ForegroundColor Green

    Write-Host ("Históricas a eliminar".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.VersionsEligible) -ForegroundColor Yellow

    Write-Host ""

    if ($Analysis.VersionsWithKnownSize -gt 0) {
        $formattedSize = Format-AppBytes -Bytes $Analysis.EstimatedBytes

        Write-Host ("Espacio recuperable estimado".PadRight(38)) -NoNewline
        Write-Host ("{0,12}" -f $formattedSize) -ForegroundColor Cyan

        if ($Analysis.VersionsWithoutKnownSize -gt 0) {
            Write-AppMuted -Message "La estimación no incluye $($Analysis.VersionsWithoutKnownSize) versión(es) sin tamaño disponible."
        }
    }
    else {
        Write-Host ("Espacio recuperable estimado".PadRight(38)) -NoNewline
        Write-Host ("{0,12}" -f "N/D")
    }

    Write-Host ""

    Write-Host ("Errores".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $Analysis.Errors)

    Write-Host ("Eventos de throttling".PadRight(38)) -NoNewline
    Write-Host ("{0,12:N0}" -f $script:ThrottleEvents.Count)

    Write-Host ("Duración".PadRight(38)) -NoNewline
    Write-Host ("{0,12}" -f $duration.ToString("hh\:mm\:ss"))

    Write-Host ""
    Write-AppOk -Message "No se ha modificado ningún archivo."
}


# ============================================================
# VISTA PREVIA POR ARCHIVO
# ============================================================

function Show-AppFileSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Analysis
    )

    Clear-AppScreen
    Show-AppHeader -Section "Archivos afectados"

    $rows = @(
        $Analysis.ByFile |
            Sort-Object VersionesEliminar -Descending |
            Select-Object -First 50 |
            ForEach-Object {
                [PSCustomObject]@{
                    Biblioteca = $_.Biblioteca
                    Archivo = $_.Archivo
                    Historicas = $_.VersionesHistoricas
                    Conserva = $_.HistoricasConservar
                    Elimina = $_.VersionesEliminar
                    Espacio = Format-AppBytes -Bytes $_.EspacioEstimadoBytes
                }
            }
    )

    $rows | Format-Table Biblioteca, Archivo, Historicas, Conserva, Elimina, Espacio -AutoSize -Wrap

    if ($Analysis.ByFile.Count -gt 50) {
        Write-AppMuted -Message "Se muestran los 50 archivos con más versiones a eliminar."
    }

    Wait-App
}

function Show-AppLibrarySummary {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Analysis
    )

    Clear-AppScreen
    Show-AppHeader -Section "Resumen por biblioteca"

    $rows = @(
        $Analysis.ByLibrary |
            ForEach-Object {
                [PSCustomObject]@{
                    Biblioteca = $_.Biblioteca
                    Plantilla = $_.Plantilla
                    Archivos = $_.Archivos
                    Versiones = $_.Versiones
                    Afectados = $_.ArchivosAfectados
                    Eliminar = $_.VersionesEliminar
                    Espacio = Format-AppBytes -Bytes $_.EspacioEstimadoBytes
                    Errores = $_.Errores
                }
            }
    )

    $rows | Format-Table Biblioteca, Plantilla, Archivos, Versiones, Afectados, Eliminar, Espacio, Errores -AutoSize

    Wait-App
}


# ============================================================
# REVALIDACION
# ============================================================

function Test-AppVersionStillEligible {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileUrl,

        [Parameter(Mandatory = $true)]
        [string]$VersionId,

        [Parameter(Mandatory = $true)]
        [int]$HistoricalToKeep
    )

    $liveVersions = @(
        Get-AppFileVersions -FileUrl $FileUrl -OperationName "Revalidar $FileUrl"
    )

    $eligibleNow = @(
        $liveVersions | Select-Object -Skip $HistoricalToKeep
    )

    foreach ($version in $eligibleNow) {
        if ([string]$version.Id -eq [string]$VersionId) {
            return [PSCustomObject]@{
                Eligible = $true
                Exists = $true
                Version = $version
                Reason = "Elegible"
            }
        }
    }

    foreach ($version in $liveVersions) {
        if ([string]$version.Id -eq [string]$VersionId) {
            return [PSCustomObject]@{
                Eligible = $false
                Exists = $true
                Version = $version
                Reason = "Ya no cumple la política de eliminación"
            }
        }
    }

    return [PSCustomObject]@{
        Eligible = $false
        Exists = $false
        Version = $null
        Reason = "La versión ya no existe"
    }
}


# ============================================================
# CONFIRMACION
# ============================================================

function Confirm-AppCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Analysis
    )

    Clear-AppScreen
    Show-AppHeader -Section "Confirmar limpieza"

    Write-AppWarning -Message "Se revalidará cada versión antes de eliminarla."
    Write-AppMuted -Message "Si el historial cambió después del análisis, la versión se omitirá si deja de ser elegible."

    Write-Host ""
    Write-Host "Archivos afectados : $($Analysis.FilesAffected)"
    Write-Host "Versiones previstas: " -NoNewline
    Write-Host $Analysis.VersionsEligible -ForegroundColor Yellow

    if ($Analysis.VersionsWithKnownSize -gt 0) {
        Write-Host "Espacio estimado   : " -NoNewline
        Write-Host (Format-AppBytes -Bytes $Analysis.EstimatedBytes) -ForegroundColor Cyan
    }

    Write-Host ""
    Write-AppMuted -Message "Las versiones se enviarán a la Papelera de reciclaje."
    Write-Host ""

    Write-Host "Escriba ELIMINAR para continuar." -ForegroundColor Red
    Write-Host ""

    $confirmation = (Read-Host "Confirmación").Trim()

    return ($confirmation -ceq "ELIMINAR")
}


# ============================================================
# ELIMINACION SEGURA
# ============================================================

function Remove-AppVersionSafely {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [int]$HistoricalToKeep
    )

    $validation = Test-AppVersionStillEligible -FileUrl $Entry.Archivo -VersionId ([string]$Entry.VersionId) -HistoricalToKeep $HistoricalToKeep

    if (-not $validation.Eligible) {
        return [PSCustomObject]@{
            Status = "Omitida"
            Message = $validation.Reason
        }
    }

    $config = Get-AppConfig
    $maxRetries = [int]$config.MaxRetries

    for ($attempt = 1; $attempt -le ($maxRetries + 1); $attempt++) {
        $delayMs = [int]$config.RequestDelayMs

        if ($delayMs -gt 0) {
            Start-Sleep -Milliseconds $delayMs
        }

        try {
            $params = @{
                Url = $Entry.Archivo
                Identity = $Entry.VersionId
                Recycle = $true
                Force = $true
                Connection = $script:PnPConnection
                ErrorAction = "Stop"
            }

            Remove-PnPFileVersion @params

            return [PSCustomObject]@{
                Status = "Eliminada"
                Message = ""
            }
        }
        catch {
            $exception = $_.Exception
            $isThrottle = Test-AppThrottleException -Exception $exception

            if (-not $isThrottle) {
                throw
            }

            $verification = Test-AppVersionStillEligible -FileUrl $Entry.Archivo -VersionId ([string]$Entry.VersionId) -HistoricalToKeep $HistoricalToKeep

            if (-not $verification.Exists) {
                return [PSCustomObject]@{
                    Status = "Eliminada"
                    Message = "La eliminación parece haberse completado antes de recibir la respuesta."
                }
            }

            if (-not $verification.Eligible) {
                return [PSCustomObject]@{
                    Status = "Omitida"
                    Message = "El historial cambió durante la operación."
                }
            }

            if ($attempt -gt $maxRetries) {
                throw
            }

            $waitSeconds = Get-AppRetryDelay -Attempt $attempt -Exception $exception

            $script:ThrottleEvents.Add(
                [PSCustomObject]@{
                    FechaHora = Get-Date
                    Operacion = "Eliminar versión $($Entry.Version) de $($Entry.Archivo)"
                    Intento = $attempt
                    EsperaSegundos = $waitSeconds
                    Mensaje = $exception.Message
                }
            )

            Write-Host ""
            Write-AppWarning -Message "Throttling detectado durante eliminación."
            Write-AppMuted -Message "Esperando $waitSeconds segundo(s) antes del siguiente intento."

            Start-Sleep -Seconds $waitSeconds
        }
    }
}


# ============================================================
# LIMPIEZA
# ============================================================

function Invoke-AppCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Analysis,

        [Parameter(Mandatory = $true)]
        [int]$HistoricalToKeep
    )

    $entries = @(
        $Analysis.Details |
            Where-Object {
                $_.Accion -eq "ELIMINAR"
            }
    )

    $log = [System.Collections.Generic.List[object]]::new()

    $removed = 0
    $skipped = 0
    $errors = 0

    $start = Get-Date

    for ($i = 0; $i -lt $entries.Count; $i++) {
        $entry = $entries[$i]

        $percent = [int]((($i + 1) / [Math]::Max(1, $entries.Count)) * 100)

        Write-Progress -Id 2 -Activity "Revalidando y eliminando versiones" -Status "$($i + 1) de $($entries.Count) - $($entry.Archivo)" -PercentComplete $percent

        try {
            $result = Remove-AppVersionSafely -Entry $entry -HistoricalToKeep $HistoricalToKeep

            if ($result.Status -eq "Eliminada") {
                $removed++
            }
            elseif ($result.Status -eq "Omitida") {
                $skipped++
            }

            $log.Add(
                [PSCustomObject]@{
                    FechaHora = Get-Date
                    Biblioteca = $entry.Biblioteca
                    Archivo = $entry.Archivo
                    VersionId = $entry.VersionId
                    Version = $entry.Version
                    Creada = $entry.Creada
                    TamanoBytes = $entry.TamanoBytes
                    Estado = $result.Status
                    Mensaje = $result.Message
                }
            )
        }
        catch {
            $errors++

            $log.Add(
                [PSCustomObject]@{
                    FechaHora = Get-Date
                    Biblioteca = $entry.Biblioteca
                    Archivo = $entry.Archivo
                    VersionId = $entry.VersionId
                    Version = $entry.Version
                    Creada = $entry.Creada
                    TamanoBytes = $entry.TamanoBytes
                    Estado = "Error"
                    Mensaje = $_.Exception.Message
                }
            )
        }
    }

    Write-Progress -Id 2 -Activity "Revalidando y eliminando versiones" -Completed

    return [PSCustomObject]@{
        Start = $start
        End = Get-Date
        Attempted = $entries.Count
        Removed = $removed
        Skipped = $skipped
        Errors = $errors
        Log = $log
    }
}


# ============================================================
# REPORTES DE EJECUCION
# ============================================================

function Export-AppCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result
    )

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"

    $executionPath = Join-Path $script:ReportsPath "Ejecucion-Limpieza-$stamp.csv"
    $throttlePath = Join-Path $script:ReportsPath "Throttling-$stamp.csv"

    $Result.Log | Export-Csv -Path $executionPath -NoTypeInformation -Encoding UTF8

    if ($script:ThrottleEvents.Count -gt 0) {
        $script:ThrottleEvents | Export-Csv -Path $throttlePath -NoTypeInformation -Encoding UTF8
    }

    return [PSCustomObject]@{
        Execution = $executionPath
        Throttling = $throttlePath
    }
}


# ============================================================
# CONFIGURACION DE THROTTLING
# ============================================================

function Show-AppThrottleSettings {
    $config = Get-AppConfig

    Clear-AppScreen
    Show-AppHeader -Section "Throttling"

    Write-AppMuted -Message "Valores conservadores recomendados para trabajo secuencial."
    Write-Host ""

    $delay = Read-AppInteger -Prompt "Pausa entre solicitudes (ms)" -Minimum 0 -Maximum 10000 -DefaultValue ([int]$config.RequestDelayMs)
    $retries = Read-AppInteger -Prompt "Máximo de reintentos" -Minimum 1 -Maximum 20 -DefaultValue ([int]$config.MaxRetries)
    $base = Read-AppInteger -Prompt "Espera base de retry (segundos)" -Minimum 1 -Maximum 60 -DefaultValue ([int]$config.RetryBaseSeconds)
    $max = Read-AppInteger -Prompt "Espera máxima de retry (segundos)" -Minimum 5 -Maximum 600 -DefaultValue ([int]$config.RetryMaxSeconds)

    $config.RequestDelayMs = $delay
    $config.MaxRetries = $retries
    $config.RetryBaseSeconds = $base
    $config.RetryMaxSeconds = $max

    Save-AppConfig -Config $config

    Write-Host ""
    Write-AppOk -Message "Configuración de throttling guardada."
    Wait-App
}


# ============================================================
# REPORTES
# ============================================================

function Open-AppReports {
    Initialize-AppFolders

    try {
        Start-Process -FilePath "explorer.exe" -ArgumentList $script:ReportsPath
    }
    catch {
        Show-AppErrorScreen -Title "Reportes" -Message "No fue posible abrir $script:ReportsPath"
    }
}


# ============================================================
# ASISTENTE
# ============================================================

function Start-AppCleanupWizard {
    $script:ThrottleEvents.Clear()

    if (-not (Confirm-AppPnP)) {
        return
    }

    if (-not (Confirm-AppAuthenticationConfigured)) {
        return
    }

    $config = Get-AppConfig

    $target = Select-AppTarget

    if ($null -eq $target) {
        return
    }

    try {
        Clear-AppScreen
        Show-AppHeader -Section "Conectando"

        Write-AppInfo -Message $target.Url

        $web = Connect-AppM365 -Target $target

        Write-Host ""
        Write-AppOk -Message "Conexión establecida."

        if (-not [string]::IsNullOrWhiteSpace([string]$web.Title)) {
            Write-AppMuted -Message $web.Title
        }

        $config = Get-AppConfig

        if ($target.Type -eq "OneDrive") {
            $config.LastOneDriveUrl = $target.Url
        }
        else {
            $config.LastSharePointUrl = $target.Url
        }

        Save-AppConfig -Config $config

        $libraries = @(Get-AppDocumentLibraries)

        if ($libraries.Count -eq 0) {
            throw "No se encontraron bibliotecas documentales visibles."
        }

        $selectedLibraries = @(Select-AppLibraries -Libraries $libraries)

        if ($selectedLibraries.Count -eq 0) {
            return
        }

        $retention = Select-AppRetention

        if ($null -eq $retention) {
            return
        }

        $analysis = Invoke-AppVersionAnalysis -Libraries $selectedLibraries -HistoricalToKeep $retention.HistoricalToKeep
        $reports = Export-AppAnalysis -Analysis $analysis

        while ($true) {
            Show-AppAnalysisSummary -Analysis $analysis -Target $target

            Write-Host ""
            Write-AppMuted -Message "Reporte detallado:"
            Write-AppMuted -Message $reports.Details
            Write-Host ""

            $items = @()

            if ($analysis.VersionsEligible -gt 0) {
                $items += [PSCustomObject]@{
                    Label = "Ejecutar limpieza"
                    Value = "Cleanup"
                    Hint = "Revalida cada versión antes de eliminar"
                }
            }

            $items += [PSCustomObject]@{
                Label = "Ver archivos afectados"
                Value = "Files"
                Hint = "Hasta 50 archivos ordenados por impacto"
            }

            $items += [PSCustomObject]@{
                Label = "Resumen por biblioteca"
                Value = "Libraries"
                Hint = ""
            }

            $items += [PSCustomObject]@{
                Label = "Abrir reportes"
                Value = "Reports"
                Hint = ""
            }

            $items += [PSCustomObject]@{
                Label = "Volver"
                Value = "Back"
                Hint = ""
            }

            Wait-App -Message "Presione ENTER para continuar"

            $choice = Show-NumberMenu -Title "Resultado del análisis" -Items $items

            switch ($choice.Value) {
                "Files" {
                    Show-AppFileSummary -Analysis $analysis
                }

                "Libraries" {
                    Show-AppLibrarySummary -Analysis $analysis
                }

                "Reports" {
                    Open-AppReports
                }

                "Back" {
                    return
                }

                "Cleanup" {
                    if (-not (Confirm-AppCleanup -Analysis $analysis)) {
                        continue
                    }

                    Clear-AppScreen
                    Show-AppHeader -Section "Limpieza"

                    Write-AppWarning -Message "No cierre esta ventana."
                    Write-AppMuted -Message "Cada versión será revalidada contra el historial actual."
                    Write-Host ""

                    $cleanup = Invoke-AppCleanup -Analysis $analysis -HistoricalToKeep $retention.HistoricalToKeep
                    $executionReports = Export-AppCleanup -Result $cleanup

                    Clear-AppScreen
                    Show-AppHeader -Section "Limpieza completada"

                    $duration = $cleanup.End - $cleanup.Start

                    Write-Host ("Versiones previstas".PadRight(34)) -NoNewline
                    Write-Host ("{0,10:N0}" -f $cleanup.Attempted)

                    Write-Host ("Eliminadas".PadRight(34)) -NoNewline
                    Write-Host ("{0,10:N0}" -f $cleanup.Removed) -ForegroundColor Green

                    Write-Host ("Omitidas por revalidación".PadRight(34)) -NoNewline
                    Write-Host ("{0,10:N0}" -f $cleanup.Skipped) -ForegroundColor Yellow

                    Write-Host ("Errores".PadRight(34)) -NoNewline
                    Write-Host ("{0,10:N0}" -f $cleanup.Errors)

                    Write-Host ("Eventos de throttling".PadRight(34)) -NoNewline
                    Write-Host ("{0,10:N0}" -f $script:ThrottleEvents.Count)

                    Write-Host ("Duración".PadRight(34)) -NoNewline
                    Write-Host ("{0,10}" -f $duration.ToString("hh\:mm\:ss"))

                    Write-Host ""
                    Write-AppMuted -Message "Auditoría:"
                    Write-Host $executionReports.Execution -ForegroundColor Cyan

                    if ($script:ThrottleEvents.Count -gt 0) {
                        Write-Host ""
                        Write-AppMuted -Message "Registro de throttling:"
                        Write-Host $executionReports.Throttling -ForegroundColor Cyan
                    }

                    Wait-App
                    return
                }
            }
        }
    }
    catch {
        Show-AppErrorScreen -Title "Proceso interrumpido" -Message $_.Exception.Message
    }
    finally {
        Disconnect-AppM365
    }
}


# ============================================================
# CONFIGURACION
# ============================================================

function Show-AppSettings {
    while ($true) {
        $config = Get-AppConfig

        if ((Test-AppGuid -Value ([string]$config.ClientId)) -and (Test-AppTenant -Value ([string]$config.Tenant))) {
            $authStatus = "Configurada"
        }
        else {
            $authStatus = "No configurada"
        }

        if ($config.PersistLogin) {
            $persistStatus = "Activada"
        }
        else {
            $persistStatus = "Desactivada"
        }

        $throttleHint = "$($config.RequestDelayMs) ms | $($config.MaxRetries) retries | máximo $($config.RetryMaxSeconds) s"

        $items = @(
            [PSCustomObject]@{
                Label = "Autenticación"
                Value = "Auth"
                Hint = $authStatus
            },
            [PSCustomObject]@{
                Label = "Sesión persistente"
                Value = "Persist"
                Hint = $persistStatus
            },
            [PSCustomObject]@{
                Label = "Throttling y reintentos"
                Value = "Throttle"
                Hint = $throttleHint
            },
            [PSCustomObject]@{
                Label = "Abrir reportes"
                Value = "Reports"
                Hint = ""
            },
            [PSCustomObject]@{
                Label = "Restablecer configuración"
                Value = "Reset"
                Hint = ""
            },
            [PSCustomObject]@{
                Label = "Volver"
                Value = "Back"
                Hint = ""
            }
        )

        $choice = Show-NumberMenu -Title "Configuración" -Items $items

        switch ($choice.Value) {
            "Auth" {
                Show-AppAuthenticationMenu
            }

            "Persist" {
                $config.PersistLogin = -not [bool]$config.PersistLogin
                Save-AppConfig -Config $config
            }

            "Throttle" {
                Show-AppThrottleSettings
            }

            "Reports" {
                Open-AppReports
            }

            "Reset" {
                Clear-AppScreen
                Show-AppHeader -Section "Restablecer configuración"

                Write-AppWarning -Message "Se eliminarán las preferencias locales."
                Write-Host ""

                if (Read-AppYesNo -Prompt "¿Continuar?" -DefaultYes $false) {
                    if (Test-Path -Path $script:ConfigPath) {
                        Remove-Item -Path $script:ConfigPath -Force
                    }

                    Write-AppOk -Message "Configuración restablecida."
                    Wait-App
                }
            }

            "Back" {
                return
            }
        }
    }
}


# ============================================================
# MENU PRINCIPAL
# ============================================================

function Show-AppMainMenu {
    while ($true) {
        $config = Get-AppConfig
        $pnpVersion = Get-PnPInstalledVersion

        if ((Test-AppGuid -Value ([string]$config.ClientId)) -and (Test-AppTenant -Value ([string]$config.Tenant))) {
            $authStatus = "Configurada"
        }
        else {
            $authStatus = "Requiere configuración"
        }

        if ($null -ne $pnpVersion) {
            $pnpStatus = "Versión $pnpVersion"
        }
        else {
            $pnpStatus = "No instalado"
        }

        $items = @(
            [PSCustomObject]@{
                Label = "Nueva limpieza"
                Value = "Cleanup"
                Hint = "SharePoint Online o OneDrive for Business"
            },
            [PSCustomObject]@{
                Label = "Autenticación"
                Value = "Auth"
                Hint = $authStatus
            },
            [PSCustomObject]@{
                Label = "PnP.PowerShell"
                Value = "PnP"
                Hint = $pnpStatus
            },
            [PSCustomObject]@{
                Label = "Configuración"
                Value = "Settings"
                Hint = ""
            },
            [PSCustomObject]@{
                Label = "Salir"
                Value = "Exit"
                Hint = ""
            }
        )

        $choice = Show-NumberMenu -Title "Inicio" -Items $items -Description @(
            "Limpieza segura del historial de versiones",
            "SharePoint Online y OneDrive for Business"
        )

        switch ($choice.Value) {
            "Cleanup" {
                Start-AppCleanupWizard
            }

            "Auth" {
                if (Confirm-AppPnP) {
                    Show-AppAuthenticationMenu
                }
            }

            "PnP" {
                [void](Install-AppPnP)
            }

            "Settings" {
                Show-AppSettings
            }

            "Exit" {
                return
            }
        }
    }
}


# ============================================================
# INICIO
# ============================================================

try {
    try {
        $Host.UI.RawUI.WindowTitle = "$script:AppName $script:AppVersion"
    }
    catch {
    }

    Initialize-AppTerminal
    Initialize-AppFolders
    Show-AppMainMenu
}
catch {
    Clear-AppScreen
    Show-AppHeader -Section "Error inesperado"

    Write-AppError -Message $_.Exception.Message

    if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
        Write-Host ""
        Write-AppMuted -Message "Detalles técnicos:"
        Write-AppMuted -Message $_.ScriptStackTrace
    }

    Wait-App
}
finally {
    Disconnect-AppM365
}

Clear-AppScreen
Show-AppHeader -Section "Finalizado"
Write-AppMuted -Message "Sesión finalizada."
Wait-App -Message "Presione ENTER para cerrar"
