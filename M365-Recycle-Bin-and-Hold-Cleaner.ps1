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
$script:AppVersion        = '2.2.0'
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
$script:AdminConnection = $null
$script:AuthRecoveryUsed = $false
$script:TemporarySiteAdmin = $null
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
# BEGIN M365 TOOLKIT CORE
# Canonical embedded source; repository maintenance scripts synchronize this block.
# No initialization, external modules or tenant operations at import time.
function Get-LocalizedString {
    param([Parameter(Mandatory)][string]$Key, [object[]]$Values = @())
    $language = if ($script:Language -eq 'en') { 'en' } else { 'es' }
    if (-not $script:ToolStrings[$language].ContainsKey($Key)) { throw "Unknown tool localization key: $Key" }
    return $script:ToolStrings[$language][$Key] -f $Values
}

function Get-ToolkitString {
    param([Parameter(Mandatory)][string]$Key, [object[]]$Arguments = @())
    $catalog = @{
        en = @{
            Invalid = 'Invalid selection.'; Select = 'Select an option'; Back = 'Back'; Cancel = 'Cancel'
            PendingAdmin = 'Remove the pending temporary administrator before changing the target.'
            Continue = 'Press Enter to continue'; YesNo = 'Answer Y or N.'
            ConfirmPhrase = 'Type {0} to continue (Enter to cancel)'; Repair = 'REPAIR'; Delete = 'DELETE'; Move = 'MOVE'; DeletePHL = 'DELETE PHL'
            Runtime = 'PowerShell 7.4+ (Core) is required. Run this tool using pwsh.'
            Missing = 'PnP.PowerShell {0}+ is required. Installation uses CurrentUser; administrator elevation is not required.'
            Install = 'Install the compatible PnP.PowerShell version now?'
            Update = 'Check/install a PnP.PowerShell update for CurrentUser now?'
            Register = 'Register a new Entra application now? Administrator consent may be required.'
            Ready = 'PnP.PowerShell ready: {0}'; ModuleError = 'PnP.PowerShell installation/import failed: {0}'
            ModuleAbsent = 'The compatible module was not found after installation.'
        }
        es = @{
            Invalid = 'Selección no válida.'; Select = 'Selecciona una opción'; Back = 'Volver'; Cancel = 'Cancelar'
            PendingAdmin = 'Retira el administrador temporal pendiente antes de cambiar de destino.'
            Continue = 'Presiona Enter para continuar'; YesNo = 'Responde S o N.'
            ConfirmPhrase = 'Escribe {0} para continuar (Enter para cancelar)'; Repair = 'REPARAR'; Delete = 'ELIMINAR'; Move = 'MOVER'; DeletePHL = 'ELIMINAR PHL'
            Runtime = 'Se requiere PowerShell 7.4+ (Core). Ejecuta esta herramienta con pwsh.'
            Missing = 'Se requiere PnP.PowerShell {0}+. Se instalará para CurrentUser; no requiere elevación de administrador.'
            Install = '¿Instalar ahora la versión compatible de PnP.PowerShell?'
            Update = '¿Buscar/instalar ahora una actualización de PnP.PowerShell para CurrentUser?'
            Register = '¿Registrar ahora una aplicación Entra nueva? Puede requerir consentimiento administrativo.'
            Ready = 'PnP.PowerShell listo: {0}'; ModuleError = 'Falló la instalación/importación de PnP.PowerShell: {0}'
            ModuleAbsent = 'No se encontró el módulo compatible después de instalarlo.'
        }
    }
    $language = if ($script:Language -eq 'en') { 'en' } else { 'es' }
    if (-not $catalog[$language].ContainsKey($Key)) { throw "Unknown localization key: $Key" }
    return $catalog[$language][$Key] -f $Arguments
}

function Write-ToolkitText {
    param([AllowNull()][AllowEmptyString()][string]$Text, [string]$Style = 'Normal', [switch]$NoNewline)
    $colors = @{ Normal='Gray'; Muted='DarkGray'; Primary='Cyan'; Info='Cyan'; Success='Green'; Warning='Yellow'; Danger='Red'; Accent='Magenta' }
    $color = if ($colors.ContainsKey($Style)) { $colors[$Style] } else { 'Gray' }
    # Qualified cmdlet avoids the legacy translation wrappers changing paths/data.
    Microsoft.PowerShell.Utility\Write-Host -Object $Text -ForegroundColor $color -NoNewline:$NoNewline
}

function Get-ToolkitWidth {
    $width = 72
    try { if ($Host.UI.RawUI.WindowSize.Width -gt 2) { $width = [Math]::Min(84, $Host.UI.RawUI.WindowSize.Width - 2) } } catch { }
    return [Math]::Max(1, $width)
}

function Write-ToolkitHeader {
    param([string]$Name, [string]$Version, [string]$Section)
    try { if (-not [Console]::IsOutputRedirected) { Clear-Host } } catch { }
    $width = Get-ToolkitWidth
    Write-ToolkitText ('=' * $width) Accent
    Write-ToolkitText "$Name  v$Version" Primary
    if ($Section) { Write-ToolkitText "[$Section]" Muted }
    Write-ToolkitText ('-' * $width) Muted
    Write-ToolkitText ''
}

function Write-ToolkitField {
    param([string]$Name, [AllowNull()]$Value, [string]$Style = 'Normal')
    $display = if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { '-' } else { [string]$Value }
    if ($Name.Length -gt 24) {
        Write-ToolkitText ("  {0}" -f $Name) Muted
        Write-ToolkitText (' ' * 27) Muted -NoNewline
    }
    else { Write-ToolkitText ('  {0,-24} ' -f $Name) Muted -NoNewline }
    Write-ToolkitText $display $Style
}

function Write-ToolkitStatus {
    param([string]$Level, [string]$Message)
    $style = switch ($Level) { 'Ok' {'Success'}; 'Info' {'Info'}; 'Warn' {'Warning'}; default {'Danger'} }
    $label = switch ($Level) { 'Ok' {'OK'}; 'Info' {'INFO'}; 'Warn' {'WARNING'}; default {'ERROR'} }
    Write-ToolkitText "  [$label] $Message" $style
}

function Read-ToolkitYesNo {
    param([string]$Prompt, [bool]$Default = $false)
    $yes = if ($script:Language -eq 'en') { 'Y' } else { 'S' }
    $hint = if ($Default) { "$yes/n" } else { "$($yes.ToLowerInvariant())/N" }
    while ($true) {
        $answer = Read-ToolkitInput "  $Prompt [$hint]"
        if (-not $answer) { return $Default }
        if ($answer -match '^(s|si|sí|y|yes)$') { return $true }
        if ($answer -match '^(n|no)$') { return $false }
        Write-ToolkitStatus Error (Get-ToolkitString YesNo)
    }
}

function Initialize-ToolkitLanguage {
    param([string]$Default = 'en', [string]$Name, [string]$Version)
    $script:Language = if ($Default -in @('en','es')) { $Default } else { 'en' }
    Write-ToolkitHeader $Name $Version 'Language / Idioma'
    Write-ToolkitText '  [1] English' Primary
    Write-ToolkitText ''
    Write-ToolkitText '  [2] Español' Primary
    Write-ToolkitText ''
    $defaultKey = if ($script:Language -eq 'en') { '1' } else { '2' }
    while ($true) {
        $choice = Read-ToolkitInput "  Language / Idioma [$defaultKey]"
        if (-not $choice) { return }
        if ($choice -eq '1') { $script:Language = 'en'; return }
        if ($choice -eq '2') { $script:Language = 'es'; return }
        Write-ToolkitStatus Error (Get-ToolkitString Invalid)
    }
}

function Test-ToolkitClientId {
    param([AllowNull()][string]$ClientId)
    $parsed = [guid]::Empty
    return [guid]::TryParseExact($ClientId, 'D', [ref]$parsed) -and $parsed -ne [guid]::Empty
}

function Read-ToolkitInput {
    param([string]$Prompt)
    return (Microsoft.PowerShell.Utility\Read-Host $Prompt).Trim()
}

function Get-ToolkitRegistrationClientId {
    param([AllowNull()]$Result)
    foreach ($item in @($Result)) {
        if ($null -eq $item) { continue }
        if (($item -is [string] -or $item -is [guid]) -and (Test-ToolkitClientId ([string]$item))) { return [string]$item }
        foreach ($name in @('ClientId','AppId','ApplicationId')) {
            $value = if ($item -is [Collections.IDictionary]) { $item[$name] } elseif ($item.PSObject.Properties[$name]) { $item.$name } else { $null }
            if (Test-ToolkitClientId ([string]$value)) { return [string]$value }
        }
    }
    return ''
}

function Confirm-ToolkitAction {
    param([ValidateSet('Repair','Delete','Move','DeletePHL')][string]$Action)
    $phrase = Get-ToolkitString $Action
    $answer = Read-ToolkitInput ('  ' + (Get-ToolkitString ConfirmPhrase @($phrase)))
    return $answer -ceq $phrase
}

function Read-ToolkitLegacyMenu {
    param([array]$Items, [ValidateSet('Key','Object','Indexed','Legacy')][string]$Mode,
        [string]$Title, [string[]]$Description = @(), [scriptblock]$RenderBody,
        [string]$Prompt, [switch]$AllowBack, [string]$ZeroLabel = 'Volver')
    if ($Title) {
        $name = if (Get-Variable AppName -Scope Script -ErrorAction SilentlyContinue) { $script:AppName } else { 'OneDrive / SharePoint Path Analyzer' }
        $version = if (Get-Variable AppVersion -Scope Script -ErrorAction SilentlyContinue) { $script:AppVersion } else { '2.0' }
        Write-ToolkitHeader $name $version (Get-LocalizedText $Title)
    }
    foreach ($line in $Description) { Write-ToolkitText (Get-LocalizedText $line) Muted }
    if ($Description.Count) { Write-ToolkitText '' }
    if ($RenderBody) { & $RenderBody; Write-ToolkitText '' }
    $map = @{}
    $index = 0
    foreach ($item in $Items) {
        $value = if ($item -is [Collections.IDictionary]) { $item['Value'] } elseif ($item.PSObject.Properties['Value']) { $item.Value } else { '' }
        if ($Mode -in @('Key','Object')) { $key = [string]$item.Key }
        elseif ($Mode -eq 'Indexed' -and $value -in @('Back','Exit')) { $key = '0' }
        else { $index++; $key = [string]$index }
        if ($key -notmatch '^\d+$' -or $map.ContainsKey($key)) { throw (Get-ToolkitString Invalid) }
        $hintName = if ($Mode -in @('Key','Object')) { 'Description' } else { 'Hint' }
        $hint = if ($item -is [Collections.IDictionary]) { [string]$item[$hintName] } elseif ($item.PSObject.Properties[$hintName]) { [string]$item.$hintName } else { '' }
        $map[$key] = if ($Mode -eq 'Object') { [pscustomobject]@{Key=$key;Label=[string]$item.Label;Description=$hint;Value=[string]$value} } else { $item }
        Write-ToolkitText ('  [{0}] {1}' -f $key,(Get-LocalizedText ([string]$item.Label))) Primary
        if ($hint) { Write-ToolkitText ('      {0}' -f (Get-LocalizedText $hint)) Muted }
        Write-ToolkitText ''
    }
    if (($AllowBack -or $Mode -eq 'Legacy') -and -not $map.ContainsKey('0')) {
        $map['0'] = if ($Mode -eq 'Object') { [pscustomobject]@{Key='0';Label=$ZeroLabel;Description='';Value='Back'} } else { $null }
        Write-ToolkitText ('  [0] {0}' -f (Get-LocalizedText $ZeroLabel)) Muted
        Write-ToolkitText ''
    }
    $caption = if ($Prompt) { Get-LocalizedText $Prompt } else { Get-ToolkitString Select }
    while ($true) {
        $choice = Read-ToolkitInput ('  ' + $caption)
        if ($map.ContainsKey($choice)) {
            if ($Mode -eq 'Key') { return $choice }
            return $map[$choice]
        }
        Write-ToolkitStatus Error (Get-ToolkitString Invalid)
    }
}

function Get-ToolkitSiteEndpoint {
    param([AllowNull()][string]$Url)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.UserInfo -or -not $uri.IsDefaultPort) { return 'Invalid' }
    if ($uri.DnsSafeHost -notmatch '^[a-z0-9][a-z0-9-]*\.sharepoint\.(com|us|cn)$') { return 'Invalid' }
    if ($uri.DnsSafeHost -match '-admin\.sharepoint\.') { return 'Admin' }
    if ($uri.DnsSafeHost -match '-my\.sharepoint\.') {
        if ($uri.AbsolutePath -match '^/personal/[^/]+(?:/|$)') { return 'OneDrive' }
        return 'OneDriveSystem'
    }
    return 'SharePoint'
}

function Get-ToolkitCompatiblePnP {
    param([version]$MinimumVersion = '3.2.0')
    Get-Module -Name PnP.PowerShell -ListAvailable -ErrorAction Stop |
        Where-Object { $_.Version -ge $MinimumVersion -and ($null -eq $_.PowerShellVersion -or $_.PowerShellVersion -le $PSVersionTable.PSVersion) } |
        Sort-Object Version -Descending | Select-Object -First 1
}

function Initialize-ToolkitPnP {
    param([version]$MinimumVersion = '3.2.0', [switch]$Update)
    if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion -lt [version]'7.4') {
        Write-ToolkitStatus Error (Get-ToolkitString Runtime)
        return $false
    }
    try {
        $module = Get-ToolkitCompatiblePnP $MinimumVersion
        if ($null -eq $module -or $Update) {
            Write-ToolkitStatus Warn (Get-ToolkitString Missing @($MinimumVersion))
            $prompt = if ($Update) { Get-ToolkitString Update } else { Get-ToolkitString Install }
            if (-not (Read-ToolkitYesNo $prompt)) { return $false }
            # Pin the suite's tested minimum instead of silently upgrading the runtime requirement.
            if ($Update) {
                Install-Module -Name PnP.PowerShell -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }
            else {
                Install-Module -Name PnP.PowerShell -RequiredVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }
            $module = Get-ToolkitCompatiblePnP $MinimumVersion
            if ($null -eq $module) { throw (Get-ToolkitString ModuleAbsent) }
        }
        Import-Module -Name $module.Path -ErrorAction Stop
        $loaded = Get-Module -Name PnP.PowerShell | Where-Object { $_.Version -eq $module.Version } | Select-Object -First 1
        if ($null -eq $loaded) { throw (Get-ToolkitString ModuleAbsent) }
        Write-ToolkitStatus Ok (Get-ToolkitString Ready @($loaded.Version))
        return $true
    }
    catch {
        Write-ToolkitStatus Error (Get-ToolkitString ModuleError @($_.Exception.Message))
        return $false
    }
}

# END M365 TOOLKIT CORE
function Get-LocalizedText {
    param([AllowNull()][object]$Text)
    if ($null -eq $Text) { return '' }; $result = [string]$Text
    if ($script:Language -eq 'en') {
        $phrases = @(
            @{ From = 'Selecciona una opción'; To = 'Select an option' }; @{ From = 'Seleccionar destino'; To = 'Select target' }
            @{ From = 'Idioma'; To = 'Language' }; @{ From = 'Cambiar idioma'; To = 'Change language' }
            @{ From = 'No seleccionado'; To = 'Not selected' }; @{ From = 'No configurado'; To = 'Not configured' }; @{ From = 'No instalado'; To = 'Not installed' }
            @{ From = 'Destino'; To = 'Target' }; @{ From = 'CONFIGURADA'; To = 'CONFIGURED' }; @{ From = 'NO CONFIGURADA'; To = 'NOT CONFIGURED' }
            @{ From = 'Activado'; To = 'Enabled' }; @{ From = 'Desactivado'; To = 'Disabled' }; @{ From = 'Sí'; To = 'Yes' }; @{ From = 'SÍ'; To = 'YES' }
            @{ From = 'Primero selecciona un sitio SharePoint o OneDrive.'; To = 'First select a SharePoint or OneDrive site.' }
            @{ From = 'Mueve los elementos al segundo nivel; no los elimina permanentemente.'; To = 'Moves items to the second stage; does not permanently delete them.' }
            @{ From = 'Elimina permanentemente elementos del segundo nivel.'; To = 'Permanently deletes items from the second stage.' }
            @{ From = 'Primero mueve el primer nivel al segundo y después elimina del segundo; el límite se aplica por nivel.'; To = 'First moves the first stage to the second, then deletes from the second; the limit applies per stage.' }
            @{ From = 'Detecta dinámicamente la biblioteca; no modifica políticas de retención.'; To = 'Dynamically detects the library; does not modify retention policies.' }
            @{ From = 'Simulación es el modo seguro.'; To = 'Simulation is the safe mode.' }; @{ From = 'Cambiar a modo REAL'; To = 'Switch to LIVE mode' }
            @{ From = 'Cambiar a modo SIMULACIÓN'; To = 'Switch to SIMULATION mode' }; @{ From = 'Operación cancelada.'; To = 'Operation canceled.' }
            @{ From = 'Confirmación'; To = 'Confirmation' }; @{ From = 'ELIMINAR'; To = 'DELETE' }; @{ From = 'ELIMINAR PHL'; To = 'DELETE PHL' }
            @{ From = 'Los elementos del primer nivel se moverán al segundo nivel; no se eliminarán permanentemente.'; To = 'First-stage items will be moved to the second stage; they will not be permanently deleted.' }
            @{ From = 'Los elementos del segundo nivel se eliminarán permanentemente.'; To = 'Second-stage items will be permanently deleted.' }
            @{ From = 'No hay elementos en este nivel dentro del límite solicitado.'; To = 'There are no items in this stage within the requested limit.' }
            @{ From = 'Valor inválido.'; To = 'Invalid value.' }; @{ From = 'El valor no puede estar vacío.'; To = 'Value cannot be empty.' }
            @{ From = 'Cada elemento del menú debe tener Key.'; To = 'Every menu item must have a Key.' }; @{ From = 'Selección inválida.'; To = 'Invalid selection.' }
            @{ From = 'No hay coincidencias.'; To = 'No matches.' }; @{ From = 'Filtrar por nombre, URL o propietario (vacío = todos)'; To = 'Filter by name, URL, or owner (empty = all)' }
            @{ From = 'No se pudieron enumerar los sitios.'; To = 'The sites could not be enumerated.' }; @{ From = 'No se encontraron sitios accesibles.'; To = 'No accessible sites were found.' }
            @{ From = 'Primero selecciona un sitio SharePoint o OneDrive.'; To = 'First select a SharePoint or OneDrive site.' }
            @{ From = 'La cuenta conectada no puede leer la papelera de este sitio.'; To = 'The connected account cannot read this site recycle bin.' }
            @{ From = 'Acceso a la papelera validado.'; To = 'Recycle-bin access validated.' }; @{ From = 'Conexión establecida.'; To = 'Connection established.' }
            @{ From = 'Cantidad máxima de elementos a procesar'; To = 'Maximum number of items to process' }
            @{ From = 'Límite de procesamiento'; To = 'Processing limit' }; @{ From = 'Primer nivel: mover al segundo. Segundo nivel: eliminación permanente.'; To = 'First stage: move to the second. Second stage: permanent deletion.' }
            @{ From = 'En "ambos niveles", el límite se aplica por separado a cada nivel.'; To = 'For "both stages", the limit applies separately to each stage.' }
            @{ From = 'La operación terminó con errores. Revisa el reporte.'; To = 'The operation finished with errors. Review the report.' }
            @{ From = 'Procesamiento completado.'; To = 'Processing completed.' }; @{ From = 'No se encontró Preservation Hold Library en este sitio.'; To = 'Preservation Hold Library was not found on this site.' }
            @{ From = 'No hay elementos para procesar dentro del límite solicitado.'; To = 'There are no items to process within the requested limit.' }
            @{ From = 'Configuración de throttling guardada.'; To = 'Throttling settings saved.' }; @{ From = 'Restablecer configuración local'; To = 'Reset local settings' }
            @{ From = 'La operación terminó con errores. Revisa el reporte.'; To = 'The operation finished with errors. Review the report.' }
            @{ From = 'Simulación completada. No se movió ningún elemento.'; To = 'Simulation completed. No items were moved.' }
            @{ From = 'Simulación completada. No se eliminó ningún elemento.'; To = 'Simulation completed. No items were deleted.' }
            @{ From = 'Los elementos fueron movidos al segundo nivel de la papelera.'; To = 'Items were moved to the second recycle-bin stage.' }
            @{ From = 'Los elementos del segundo nivel fueron eliminados permanentemente.'; To = 'Second-stage items were permanently deleted.' }
            @{ From = 'No se encontró Preservation Hold Library en este sitio.'; To = 'Preservation Hold Library was not found on this site.' }
            @{ From = 'Puede no existir, estar vacía/no provisionada o no ser accesible con la identidad actual.'; To = 'It may not exist, be empty/not provisioned, or be inaccessible to the current identity.' }
            @{ From = 'Eliminar contenido de Preservation Hold Library puede verse bloqueado por políticas de retención.'; To = 'Deleting Preservation Hold Library content may be blocked by retention policies.' }
            @{ From = 'La herramienta no deshabilita ni modifica políticas de cumplimiento o retención.'; To = 'The tool does not disable or modify compliance or retention policies.' }
            @{ From = 'Aplicación Entra / autenticación'; To = 'Entra application / authentication' }
            @{ From = 'Aplicación Entra / autenticación PnP'; To = 'Entra application / PnP authentication' }
            @{ From = 'Comprueba autenticación y acceso a búsqueda de sitios.'; To = 'Checks authentication and access to site search.' }
            @{ From = 'Usar otra aplicación existente'; To = 'Use another existing application' }
            @{ From = 'Introduce y opcionalmente valida otro Client ID.'; To = 'Enter and optionally validate another Client ID.' }
            @{ From = 'Registrar una nueva aplicación Entra'; To = 'Register a new Entra application' }
            @{ From = 'Crea otra app PnP y la configura en esta herramienta.'; To = 'Creates another PnP app and configures it in this tool.' }
            @{ From = 'Quitar Client ID de la configuración local'; To = 'Remove Client ID from local configuration' }
            @{ From = 'No elimina la aplicación en Entra.'; To = 'Does not delete the application in Entra.' }
            @{ From = 'Eliminar sesión persistente'; To = 'Clear persisted session' }
            @{ From = 'Introduce y opcionalmente valida un Client ID ya registrado.'; To = 'Enter and optionally validate an already registered Client ID.' }
            @{ From = 'Crea una nueva app PnP para esta herramienta.'; To = 'Creates a new PnP app for this tool.' }
            @{ From = 'Usar una aplicación existente'; To = 'Use an existing application' }
            @{ From = 'Instalar / actualizar PnP.PowerShell'; To = 'Install / update PnP.PowerShell' }
            @{ From = 'Probar nuevamente el sitio actual'; To = 'Test the current site again' }
            @{ From = 'Alternar persistencia de login'; To = 'Toggle login persistence' }
            @{ From = 'Cambiar límite predeterminado'; To = 'Change default limit' }
            @{ From = 'Cambiar límite máximo permitido'; To = 'Change maximum allowed limit' }
            @{ From = 'Rango permitido por la herramienta: 1 a 10000.'; To = 'Allowed range for this tool: 1 to 10,000.' }
            @{ From = 'Abrir reportes'; To = 'Open reports' }
            @{ From = 'Seleccionar / cambiar sitio'; To = 'Select / change site' }
            @{ From = 'Buscar SharePoint u OneDrive en el tenant o usar una URL directa.'; To = 'Search for SharePoint or OneDrive in the tenant, or use a direct URL.' }
            @{ From = 'Buscar sitios del tenant.'; To = 'Search for sites in the tenant.' }
            @{ From = 'Buscar OneDrive reales del tenant.'; To = 'Search for real OneDrive sites in the tenant.' }
            @{ From = 'Detecta automáticamente SharePoint u OneDrive.'; To = 'Automatically detects SharePoint or OneDrive.' }
            @{ From = 'Validar aplicación configurada'; To = 'Validate configured application' }
            @{ From = 'Validar aplicación'; To = 'Validate application' }
            @{ From = 'Autenticación requerida'; To = 'Authentication required' }
            @{ From = 'Configuración'; To = 'Settings' }
            @{ From = 'Diagnóstico'; To = 'Diagnostics' }
            @{ From = 'Alternar persistencia de login'; To = 'Toggle login persistence' }
            @{ From = 'Introduce un número entre $Minimum y $Maximum.'; To = 'Enter a number between $Minimum and $Maximum.' }
            @{ From = 'Opción inválida.'; To = 'Invalid option.' }
            @{ From = 'PnP.PowerShell no está instalado.'; To = 'PnP.PowerShell is not installed.' }
            @{ From = 'PnP.PowerShell $version es anterior a $script:MinimumPnPVersion.'; To = 'PnP.PowerShell $version is older than $script:MinimumPnPVersion.' }
            @{ From = 'PnP.PowerShell está listo.'; To = 'PnP.PowerShell is ready.' }
            @{ From = 'PnP.PowerShell no está disponible.'; To = 'PnP.PowerShell is not available.' }
            @{ From = ' está instalado.'; To = ' is installed.' }; @{ From = ' es anterior a '; To = ' is older than ' }; @{ From = ' está listo.'; To = ' is ready.' }; @{ From = ' no está disponible.'; To = ' is not available.' }
            @{ From = 'No existe una autenticación válida configurada.'; To = 'No valid authentication is configured.' }
            @{ From = 'Tenant no configurado o inválido.'; To = 'Tenant is not configured or is invalid.' }
            @{ From = 'Client ID no configurado o inválido.'; To = 'Client ID is not configured or is invalid.' }
            @{ From = 'No fue posible determinar la URL del centro de administración.'; To = 'Could not determine the admin center URL.' }
            @{ From = 'Validando autenticación contra el tenant...'; To = 'Validating authentication against the tenant...' }
            @{ From = 'La validación falló.'; To = 'Validation failed.' }
            @{ From = 'Usar aplicación existente'; To = 'Use existing application' }
            @{ From = 'Client ID de la aplicación existente'; To = 'Client ID of the existing application' }
            @{ From = 'Se conservó el Client ID anterior.'; To = 'The previous Client ID was preserved.' }
            @{ From = 'Registrar nueva aplicación Entra'; To = 'Register new Entra application' }
            @{ From = 'La aplicación solicitará el permiso delegado de SharePoint:'; To = 'The application will request delegated SharePoint permission:' }
            @{ From = 'Nombre de la aplicación'; To = 'Application name' }
            @{ From = 'Nueva aplicación configurada.'; To = 'New application configured.' }
            @{ From = 'Registrar aplicación'; To = 'Register application' }
            @{ From = 'Esto no elimina la aplicación de Microsoft Entra ID.'; To = 'This does not delete the application from Microsoft Entra ID.' }
            @{ From = 'Client ID eliminado de la configuración local.'; To = 'Client ID removed from local configuration.' }
            @{ From = 'Sesión persistente'; To = 'Persisted session' }
            @{ From = 'Sesión persistente eliminada.'; To = 'Persisted session removed.' }
            @{ From = 'No hay una aplicación Entra configurada.'; To = 'No Entra application is configured.' }
            @{ From = 'Conectando al centro de administración: $adminUrl'; To = 'Connecting to the admin center: $adminUrl' }
            @{ From = 'Conectando al centro de administración: '; To = 'Connecting to the admin center: ' }
            @{ From = '(sin título)'; To = '(untitled)' }
            @{ From = 'Introduce una URL HTTPS válida de *.sharepoint.com.'; To = 'Enter a valid HTTPS URL from *.sharepoint.com.' }
            @{ From = 'El sitio seleccionado no contiene una URL válida.'; To = 'The selected site does not contain a valid URL.' }
            @{ From = 'Destino seleccionado: '; To = 'Target selected: ' }
            @{ From = 'Validando destino'; To = 'Validating target' }
            @{ From = 'Validación del destino'; To = 'Target validation' }
            @{ From = 'No se obtuvo una conexión PnP válida.'; To = 'A valid PnP connection was not obtained.' }
            @{ From = 'Access denied: '; To = 'Access denied: ' }
            @{ From = 'Se moverán hasta $($items.Count) elemento(s) del primer nivel al segundo nivel. No es una eliminación permanente.'; To = 'Up to $($items.Count) item(s) will be moved from the first stage to the second stage. This is not a permanent deletion.' }
            @{ From = 'Se eliminarán permanentemente hasta $($items.Count) elemento(s) del segundo nivel.'; To = 'Up to $($items.Count) item(s) will be permanently deleted from the second stage.' }
            @{ From = 'Se intentará eliminar permanentemente hasta $($items.Count) elemento(s) de la biblioteca detectada.'; To = 'Up to $($items.Count) item(s) will be permanently deleted from the detected library.' }
            @{ From = 'Simulation completed. No se modificó la biblioteca.'; To = 'Simulation completed. The library was not modified.' }
            @{ From = 'Máximo configurado por ejecución: '; To = 'Maximum configured per run: ' }
            @{ From = 'Espera máxima de retry (segundos)'; To = 'Maximum retry wait (seconds)' }
            @{ From = 'Modo inicial'; To = 'Initial mode' }
            @{ From = 'Se eliminarán las preferencias locales de esta herramienta.'; To = 'Local preferences for this tool will be deleted.' }
            @{ From = 'La conexión funciona, pero no hay acceso suficiente a la papelera.'; To = 'The connection works, but there is not enough access to the recycle bin.' }
            @{ From = 'SharePoint/OneDrive aplicó throttling.'; To = 'SharePoint/OneDrive applied throttling.' }
            @{ From = 'No se pudo obtener el contexto CSOM de la conexión PnP.'; To = 'Could not obtain the CSOM context from the PnP connection.' }
            @{ From = 'Acción'; To = 'Action' }; @{ From = 'Modo'; To = 'Mode' }
            @{ From = 'MODO REAL'; To = 'LIVE MODE' }
            @{ From = 'Simulación completada. No se modificó la biblioteca.'; To = 'Simulation completed. The library was not modified.' }
            @{ From = 'Simulation completed. No se modificó la biblioteca.'; To = 'Simulation completed. The library was not modified.' }
            @{ From = 'Restablecer configuración'; To = 'Reset settings' }
            @{ From = 'La cuenta conectada no es Site Collection Administrator para este sitio.'; To = 'The connected account is not a Site Collection Administrator for this site.' }
            @{ From = 'Get-PnPRecycleBinItem requiere Site Collection Administrator.'; To = 'Get-PnPRecycleBinItem requires Site Collection Administrator access.' }
            @{ From = '¿Deseas agregarte temporalmente como Site Collection Administrator?'; To = 'Do you want to add yourself temporarily as Site Collection Administrator?' }
            @{ From = 'UPN de la cuenta que se agregará como Site Collection Administrator'; To = 'UPN of the account to add as Site Collection Administrator' }
            @{ From = 'No se pudo detectar el UPN de la cuenta conectada.'; To = 'Could not detect the UPN of the connected account.' }
            @{ From = 'Se agregó acceso temporal como Site Collection Administrator.'; To = 'Temporary Site Collection Administrator access was added.' }
            @{ From = 'La cuenta conectada no tiene permisos para agregar Site Collection Administrators desde el centro de administración.'; To = 'The connected account does not have permission to add Site Collection Administrators from the admin center.' }
            @{ From = 'El acceso temporal se mantendrá hasta que lo retires desde el menú principal.'; To = 'Temporary access will remain until you remove it from the main menu.' }
            @{ From = 'Retirar Site Collection Admin temporal'; To = 'Remove temporary Site Collection Admin' }
            @{ From = 'Hay un acceso temporal de Site Collection Admin activo para este sitio.'; To = 'Temporary Site Collection Admin access is active for this site.' }
            @{ From = '¿Retirar ahora el acceso temporal?'; To = 'Remove temporary access now?' }
            @{ From = '¿Retirar el acceso temporal ahora?'; To = 'Remove temporary access now?' }
            @{ From = 'Acceso temporal retirado correctamente.'; To = 'Temporary access removed successfully.' }
            @{ From = 'No se retiró el acceso temporal.'; To = 'Temporary access was not removed.' }
            @{ From = 'El acceso temporal sigue activo.'; To = 'Temporary access is still active.' }
            @{ From = 'La operación terminó. Puedes retirar ahora el acceso temporal.'; To = 'The operation is finished. You can remove temporary access now.' }
            @{ From = 'Debes retirar el acceso temporal antes de cambiar de sitio.'; To = 'You must remove temporary access before changing sites.' }
            @{ From = '¿Retirar el acceso temporal antes de salir?'; To = 'Remove temporary access before exiting?' }
            @{ From = 'No se puede continuar con otro sitio mientras el acceso temporal siga activo.'; To = 'You cannot continue with another site while temporary access remains active.' }
            @{ From = '¿Agregar '; To = 'Add ' }; @{ From = ' como Site Collection Administrator?'; To = ' as Site Collection Administrator?' }
            @{ From = 'Se agregará '; To = 'Will add ' }; @{ From = ' al sitio como Site Collection Administrator.'; To = ' to the site as Site Collection Administrator.' }
            @{ From = 'UPN detectado: '; To = 'Detected UPN: ' }
            @{ From = 'Esperando propagación del acceso administrativo...'; To = 'Waiting for administrative access propagation...' }
            @{ From = 'Reintentando validación de acceso...'; To = 'Retrying access validation...' }
            @{ From = 'No se pudo confirmar el acceso administrativo después de agregar el usuario.'; To = 'Administrative access could not be confirmed after adding the user.' }
            @{ From = 'PnP no devolvió un token de SharePoint reutilizable.'; To = 'PnP did not return a reusable SharePoint token.' }
            @{ From = 'Sesión de autenticación'; To = 'Authentication session' }
            @{ From = 'Reportes'; To = 'Reports' }
            @{ From = 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'; To = 'The in-memory session is reused while the tool remains open.' }
            @{ From = 'Sesión persistente: si la activas, PnP puede reutilizar el inicio de sesión cuando abras el script nuevamente.'; To = 'Persisted session: when enabled, PnP can reuse the sign-in when you open the script again.' }
            @{ From = 'Si la desactivas, solo se usa la sesión actual y podrás iniciar sesión de nuevo en la siguiente ejecución.'; To = 'When disabled, the login is used only for the current execution and you can sign in again on the next run.' }
            @{ From = '¿Guardar la sesión para futuras ejecuciones?'; To = 'Save the session for future runs?' }
            @{ From = 'Persist login controla si PnP puede reutilizar el login al abrir el script nuevamente.'; To = 'Persist login controls whether PnP can reuse the login when you open the script again.' }
            @{ From = 'Cambiar tenant o aplicación conectada'; To = 'Change tenant or connected application' }
            @{ From = 'Cambiar tenant'; To = 'Change tenant' }
            @{ From = 'Actualiza el tenant conectado y libera la sesión anterior.'; To = 'Updates the connected tenant and releases the previous session.' }
            @{ From = 'Cambia el idioma de la interfaz.'; To = 'Changes the interface language.' }
            @{ From = 'Controla si PnP puede reutilizar el login cuando abras el script nuevamente.'; To = 'Controls whether PnP can reuse the login when you open the script again.' }
            @{ From = 'Cantidad inicial de elementos a procesar por operación.'; To = 'Initial number of items to process per operation.' }
            @{ From = 'Ajusta pausas, backoff y cantidad de reintentos.'; To = 'Adjusts pacing, backoff, and retry count.' }
            @{ From = 'Abre la carpeta donde se guardan los CSV de auditoría.'; To = 'Opens the folder where audit CSV files are saved.' }
            @{ From = 'Borra valores guardados; conserva idioma y reportes.'; To = 'Clears saved values; preserves language and reports.' }
            @{ From = 'Limpia el destino y la conexión de sitio, pero conserva tenant, aplicación y reportes.'; To = 'Clears the target and site connection, but preserves the tenant, application, and reports.' }
            @{ From = 'Abre la configuración y libera la conexión anterior cuando cambies estos valores.'; To = 'Opens settings and releases the previous connection when you change these values.' }
            @{ From = 'Cierra la herramienta.'; To = 'Closes the tool.' }
            @{ From = 'Regresa al menú anterior.'; To = 'Returns to the previous menu.' }
            @{ From = 'Sin descripción adicional.'; To = 'No additional details.' }
            @{ From = 'Gestión del contexto'; To = 'Context management' }
            @{ From = 'Limpiar contexto de trabajo'; To = 'Clear working context' }
            @{ From = 'Limpia el destino y la conexión de sitio. Conserva tenant, aplicación, preferencias y reportes.'; To = 'Clears the target and site connection. Preserves the tenant, application, preferences, and reports.' }
            @{ From = 'La sesión de autenticación base se conserva; cambiar tenant o aplicación la renovará.'; To = 'The base authentication session is preserved; changing the tenant or application will renew it.' }
            @{ From = 'No se puede limpiar el contexto mientras exista un acceso temporal activo.'; To = 'The context cannot be cleared while temporary access is active.' }
            @{ From = 'Retíralo primero desde el menú principal.'; To = 'Remove it first from the main menu.' }
            @{ From = 'Retira primero el acceso temporal antes de cambiar tenant o aplicación.'; To = 'Remove temporary access before changing the tenant or application.' }
            @{ From = 'Conexión de sitio'; To = 'Site connection' }
            @{ From = 'VACÍO'; To = 'EMPTY' }
            @{ From = 'App guardada'; To = 'Saved app' }
            @{ From = 'Cambia el idioma de la interfaz.'; To = 'Changes the interface language.' }
            @{ From = 'Introduce y valida un Client ID ya registrado.'; To = 'Enter and validate an already registered Client ID.' }
            @{ From = 'Crea una nueva aplicación PnP para esta herramienta.'; To = 'Creates a new PnP application for this tool.' }
            @{ From = 'Sesión en memoria'; To = 'In-memory session' }
            @{ From = 'ACTIVA'; To = 'ACTIVE' }
            @{ From = 'INACTIVA'; To = 'INACTIVE' }
            @{ From = 'VACÍA'; To = 'EMPTY' }
            @{ From = 'Contexto activo · limpia el destino o cambia tenant/aplicación conectada.'; To = 'Active context · clear the target or change the connected tenant/application.' }
            @{ From = 'No hay contexto activo · limpia selecciones o cambia tenant/aplicación conectada.'; To = 'No active context · clear selections or change the connected tenant/application.' }
            @{ From = 'Limpia el sitio seleccionado y la conexión de sitio. Conserva tenant, aplicación, preferencias y reportes.'; To = 'Clears the selected site and site connection. Preserves the tenant, application, preferences, and reports.' }
            @{ From = 'CONFIGURADA · valida, cambia o registra otra aplicación.'; To = 'CONFIGURED · validate, change, or register another application.' }
            @{ From = 'NO CONFIGURADA · usa una aplicación existente o registra una nueva.'; To = 'NOT CONFIGURED · use an existing application or register a new one.' }
            @{ From = 'Tenant, límites, throttling, persistencia, idioma y reportes.'; To = 'Tenant, limits, throttling, persistence, language, and reports.' }
            @{ From = 'Valida PowerShell, PnP, autenticación y acceso al sitio actual.'; To = 'Validates PowerShell, PnP, authentication, and access to the current site.' }
            @{ From = 'SIMULACIÓN es el modo seguro.'; To = 'SIMULATION is the safe mode.' }
            @{ From = 'REAL elimina o modifica elementos según la operación elegida.'; To = 'LIVE removes or modifies items according to the selected operation.' }
            @{ From = 'La sesión autenticada no pudo reutilizarse; se solicitará autenticación nuevamente una sola vez.'; To = 'The authenticated session could not be reused; authentication will be requested once.' }
            @{ From = 'Gestionar contexto de trabajo'; To = 'Manage working context' }
            @{ From = 'Abrir carpeta de reportes'; To = 'Open reports folder' }
            @{ From = 'Admin temporal'; To = 'Temporary admin' }
            @{ From = 'Estado config'; To = 'Configuration status' }
            @{ From = 'Contexto'; To = 'Context' }
            @{ From = 'Tipo'; To = 'Type' }
            @{ From = 'Validar aplicación configurada'; To = 'Validate configured application' }
            @{ From = 'Cierra la sesión actual y elimina el login persistido por PnP.PowerShell.'; To = 'Closes the current session and removes the login persisted by PnP.PowerShell.' }
            @{ From = 'Consultando sitios del tenant...'; To = 'Querying sites in the tenant...' }
            @{ From = 'Disponibles'; To = 'Available' }
            @{ From = 'Hay '; To = 'There are ' }
            @{ From = ' coincidencias. Refina el filtro a 50 o menos.'; To = ' matches. Refine the filter to 50 or fewer.' }
            @{ From = 'URL completa de SharePoint / OneDrive'; To = 'Full SharePoint / OneDrive URL' }
            @{ From = 'Papelera · '; To = 'Recycle bin · ' }
            @{ From = 'Primer nivel'; To = 'First stage' }
            @{ From = 'Segundo nivel'; To = 'Second stage' }
            @{ From = 'Mover al segundo nivel'; To = 'Move to second stage' }
            @{ From = 'Eliminar permanentemente'; To = 'Permanently delete' }
            @{ From = 'Leyendo elementos...'; To = 'Reading items...' }
            @{ From = 'elemento(s) encontrados.'; To = 'item(s) found.' }
            @{ From = 'Escribe '; To = 'Type ' }
            @{ From = 'Simulado: mover a segundo nivel'; To = 'Simulated: move to second stage' }
            @{ From = 'Movido a segundo nivel'; To = 'Moved to second stage' }
            @{ From = 'Simulado: eliminar permanentemente'; To = 'Simulated: permanently delete' }
            @{ From = 'Eliminado permanentemente'; To = 'Permanently deleted' }
            @{ From = 'Procesando Preservation Hold Library'; To = 'Processing Preservation Hold Library' }
            @{ From = 'Buscando Preservation Hold Library...'; To = 'Searching for Preservation Hold Library...' }
            @{ From = 'Biblioteca detectada: '; To = 'Library detected: ' }
            @{ From = 'Confirmar Preservation Hold Library'; To = 'Confirm Preservation Hold Library' }
            @{ From = 'Valores conservadores para trabajo secuencial.'; To = 'Conservative values for sequential work.' }
            @{ From = 'Pausa entre solicitudes (ms)'; To = 'Delay between requests (ms)' }
            @{ From = 'Máximo de reintentos'; To = 'Maximum retries' }
            @{ From = 'Espera base de retry (segundos)'; To = 'Base retry wait (seconds)' }
            @{ From = 'Restablecer configuración'; To = 'Reset settings' }
            @{ From = 'Responde S o N.'; To = 'Answer Y or N.' }
            @{ From = 'Introduce un número entre '; To = 'Enter a number between ' }
            @{ From = 'Usa el dominio inicial, por ejemplo empresa.onmicrosoft.com.'; To = 'Use the initial domain, for example company.onmicrosoft.com.' }
            @{ From = 'Debe ser un GUID válido.'; To = 'Must be a valid GUID.' }
            @{ From = 'URL directa'; To = 'Direct URL' }
            @{ From = 'Confirmar · '; To = 'Confirm · ' }
            @{ From = 'Nivel'; To = 'Stage' }
            @{ From = 'Movidos'; To = 'Moved' }
            @{ From = 'Eliminados'; To = 'Deleted' }
            @{ From = 'Simulados'; To = 'Simulated' }
            @{ From = '¿Instalar/actualizar ahora?'; To = 'Install/update now?' }
            @{ From = 'Dominio inicial del tenant'; To = 'Initial tenant domain' }
            @{ From = '¿Validar antes de guardar?'; To = 'Validate before saving?' }
            @{ From = '¿Continuar con el registro?'; To = 'Continue with registration?' }
            @{ From = 'Client ID creado'; To = 'Created Client ID' }
            @{ From = '¿Probarla ahora?'; To = 'Test it now?' }
            @{ From = '¿Limpiar el contexto de trabajo?'; To = 'Clear the working context?' }
            @{ From = '¿Retirar el acceso temporal antes de cambiar de sitio?'; To = 'Remove temporary access before changing sites?' }
            @{ From = 'Retira primero el acceso temporal antes de cambiar tenant o aplicación.'; To = 'Remove temporary access before changing the tenant or application.' }
            @{ From = 'Aplicación, tenant y acceso PnP validados correctamente.'; To = 'Application, tenant, and PnP access validated successfully.' }
            @{ From = 'Aplicación configurada.'; To = 'Application configured.' }
            @{ From = 'Solo elimina el Client ID guardado por esta herramienta.'; To = 'Only removes the Client ID saved by this tool.' }
            @{ From = 'Límite de procesamiento'; To = 'Processing limit' }
            @{ From = 'Primer nivel: mover al segundo. Segundo nivel: eliminación permanente.'; To = 'First stage: move to the second. Second stage: permanent deletion.' }
            @{ From = 'En "ambos niveles", el límite se aplica por separado a cada nivel.'; To = 'For "both stages", the limit applies separately to each stage.' }
            @{ From = 'Se moverán hasta '; To = 'Up to ' }
            @{ From = ' elemento(s) del primer nivel al segundo nivel. No es una eliminación permanente.'; To = ' item(s) will be moved from the first stage to the second stage. This is not a permanent deletion.' }
            @{ From = 'Se eliminarán permanentemente hasta '; To = 'Up to ' }
            @{ From = ' elemento(s) del segundo nivel.'; To = ' item(s) will be permanently deleted from the second stage.' }
            @{ From = 'Se intentará eliminar permanentemente hasta '; To = 'Up to ' }
            @{ From = ' elemento(s) de la biblioteca detectada.'; To = ' item(s) from the detected library.' }
            @{ From = ' detectado. '; To = ' detected. ' }
            @{ From = 'Se requiere PowerShell '; To = 'PowerShell ' }
            @{ From = ' o superior.'; To = ' or later.' }
            @{ From = 'Límite defecto'; To = 'Default limit' }
            @{ From = ' de '; To = ' of ' }
            @{ From = 'reintento'; To = 'retry' }
        )
        foreach ($translation in $phrases) { $result = $result.Replace($translation.From, $translation.To) }
        foreach ($translation in $script:UiTranslations) { $result = $result.Replace($translation.From, $translation.To) }
        $common = @(
            @{ From = 'Selecciona'; To = 'Select' }; @{ From = 'Seleccione'; To = 'Select' }; @{ From = 'Introducir'; To = 'Enter' }; @{ From = 'Introduzca'; To = 'Enter' }
            @{ From = 'Buscar'; To = 'Search' }; @{ From = 'Usar'; To = 'Use' }; @{ From = 'Cambiar'; To = 'Change' }; @{ From = 'Guardar'; To = 'Save' }
            @{ From = 'Validar'; To = 'Validate' }; @{ From = 'Registrar'; To = 'Register' }; @{ From = 'Procesar'; To = 'Process' }; @{ From = 'Procesando'; To = 'Processing' }
            @{ From = 'Eliminar'; To = 'Remove' }; @{ From = 'Eliminación'; To = 'Removal' }; @{ From = 'Vaciar'; To = 'Empty' }; @{ From = 'mover'; To = 'move' }
            @{ From = 'Aplicación'; To = 'Application' }; @{ From = 'Conexión'; To = 'Connection' }; @{ From = 'establecida'; To = 'established' }
            @{ From = 'Sitios'; To = 'Sites' }; @{ From = 'Sitio'; To = 'Site' }; @{ From = 'Bibliotecas'; To = 'Libraries' }; @{ From = 'Biblioteca'; To = 'Library' }
            @{ From = 'Elementos'; To = 'Items' }; @{ From = 'elementos'; To = 'items' }; @{ From = 'Encontrados'; To = 'Found' }; @{ From = 'Procesados'; To = 'Processed' }
            @{ From = 'Estado'; To = 'Status' }; @{ From = 'Resultado'; To = 'Result' }; @{ From = 'Reporte'; To = 'Report' }; @{ From = 'Límite'; To = 'Limit' }
            @{ From = 'Máximo'; To = 'Maximum' }; @{ From = 'predeterminado'; To = 'default' }; @{ From = 'permitido'; To = 'allowed' }
            @{ From = 'No hay'; To = 'There are no' }; @{ From = 'No se encontró'; To = 'No ... was found' }; @{ From = 'No fue posible'; To = 'It was not possible' }
            @{ From = 'Operación cancelada'; To = 'Operation canceled' }; @{ From = 'Simulación completada'; To = 'Simulation completed' }; @{ From = 'Errores'; To = 'Errors' }
        )
        foreach ($translation in $common) { $result = $result.Replace($translation.From, $translation.To) }
        $words = @(
            @{ From = 'al'; To = 'when' }; @{ From = 'alguna'; To = 'some' }; @{ From = 'actual'; To = 'current' }; @{ From = 'administrador'; To = 'administrator' }
            @{ From = 'archivo'; To = 'file' }; @{ From = 'archivos'; To = 'files' }; @{ From = 'analizar'; To = 'analyze' }; @{ From = 'buscar'; To = 'search' }
            @{ From = 'cambiar'; To = 'change' }; @{ From = 'con'; To = 'with' }; @{ From = 'contenido'; To = 'content' }; @{ From = 'después'; To = 'after' }
            @{ From = 'dentro'; To = 'within' }; @{ From = 'eliminar'; To = 'remove' }; @{ From = 'eliminación'; To = 'removal' }; @{ From = 'en'; To = 'in' }
            @{ From = 'estado'; To = 'status' }; @{ From = 'elemento'; To = 'item' }; @{ From = 'elementos'; To = 'items' }; @{ From = 'encontrados'; To = 'found' }
            @{ From = 'introducir'; To = 'enter' }; @{ From = 'límite'; To = 'limit' }; @{ From = 'máximo'; To = 'maximum' }; @{ From = 'nivel'; To = 'stage' }
            @{ From = 'nueva'; To = 'new' }; @{ From = 'opción'; To = 'option' }; @{ From = 'para'; To = 'for' }; @{ From = 'papelera'; To = 'recycle bin' }
            @{ From = 'procesar'; To = 'process' }; @{ From = 'procesando'; To = 'processing' }; @{ From = 'reporte'; To = 'report' }; @{ From = 'reportes'; To = 'reports' }
            @{ From = 'resultado'; To = 'result' }; @{ From = 'selecciona'; To = 'select' }; @{ From = 'seleccionar'; To = 'select' }; @{ From = 'seleccionado'; To = 'selected' }
            @{ From = 'sitio'; To = 'site' }; @{ From = 'sitios'; To = 'sites' }; @{ From = 'sin'; To = 'without' }; @{ From = 'total'; To = 'total' }
            @{ From = 'vaciar'; To = 'empty' }; @{ From = 'válido'; To = 'valid' }; @{ From = 'y'; To = 'and' }; @{ From = 'u'; To = 'or' }
        )
        foreach ($translation in $words) {
            $pattern = '(?<![\p{L}])' + [regex]::Escape($translation.From) + '(?![\p{L}])'
            $result = [regex]::Replace($result, $pattern, $translation.To)
        }
    }
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
function Write-Progress {
    [CmdletBinding()] param([int]$Id = 0,[string]$Activity,[string]$Status,[int]$PercentComplete,[int]$SecondsRemaining,[switch]$Completed)
    $parameters = @{ Id = $Id }
    if ($PSBoundParameters.ContainsKey('Activity')) { $parameters.Activity = Get-LocalizedText $Activity }
    if ($PSBoundParameters.ContainsKey('Status')) { $parameters.Status = Get-LocalizedText $Status }
    if ($PSBoundParameters.ContainsKey('PercentComplete')) { $parameters.PercentComplete = $PercentComplete }
    if ($PSBoundParameters.ContainsKey('SecondsRemaining')) { $parameters.SecondsRemaining = $SecondsRemaining }
    if ($Completed) { $parameters.Completed = $true }
    Microsoft.PowerShell.Utility\Write-Progress @parameters
}
function Initialize-AppLanguage {
    param([string]$Default = 'en')
    Initialize-ToolkitLanguage -Default $Default -Name $script:AppName -Version $script:AppVersion
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
    Write-ToolkitText (Get-LocalizedText $Text) $Style -NoNewline:$NoNewline
}

function Clear-AppScreen {
    try { if (-not [Console]::IsOutputRedirected) { Clear-Host } } catch { }
}

function Show-AppHeader {
    param(
        [string]$Section = 'Inicio'
    )
    Write-ToolkitHeader $script:AppName $script:AppVersion (Get-LocalizedText $Section)
}

function Write-AppOk {
    param([Parameter(Mandatory)][string]$Message)
    Write-ToolkitStatus Ok (Get-LocalizedText $Message)
}

function Write-AppInfo {
    param([Parameter(Mandatory)][string]$Message)
    Write-ToolkitStatus Info (Get-LocalizedText $Message)
}

function Write-AppWarning {
    param([Parameter(Mandatory)][string]$Message)
    Write-ToolkitStatus Warn (Get-LocalizedText $Message)
}

function Write-AppError {
    param([Parameter(Mandatory)][string]$Message)
    Write-ToolkitStatus Error (Get-LocalizedText $Message)
}

function Write-AppMuted {
    param([Parameter(Mandatory)][string]$Message)
    Write-AppStyled -Text $Message -Style Muted
}

function Write-AppSection {
    param([Parameter(Mandatory)][string]$Title)
    Write-ToolkitText ''
    Write-ToolkitText (Get-LocalizedText $Title) Primary
}

function Write-AppField {
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$Value,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal'
    )
    Write-ToolkitField (Get-LocalizedText $Name) $Value $Style
}

function Wait-App {
    param([string]$Message = 'Enter para continuar')
    [void](Read-ToolkitInput ('  ' + (Get-ToolkitString Continue)))
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
    return Test-ToolkitClientId $Value
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

    return (Get-ToolkitSiteEndpoint $Value) -in @('SharePoint','OneDrive')
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
    return Read-ToolkitYesNo (Get-LocalizedText $Prompt) $DefaultYes
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
    return Read-ToolkitLegacyMenu -Items $Items -Mode Object -Prompt $Prompt -AllowBack:$AllowBack
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
        Write-AppStyled -Text ('  {0,2}  ' -f ($i + 1)) -Style Primary -NoNewline
        Write-Host (& $Label $list[$i])
        Write-Host ''
    }

    Write-AppStyled -Text '  0  ' -Style Primary -NoNewline
    Write-Host 'Cancelar'
    Write-AppMuted -Message '     Regresa al menú anterior.'

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
        PersistLogin        = $false
        PersistLoginConfigured = $false
        Language            = 'es'
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

function Close-AppConnection {
    param([AllowNull()]$Connection)

    if ($null -eq $Connection) {
        return
    }

    try {
        $contextProperty = $Connection.PSObject.Properties['Context']

        if ($null -ne $contextProperty -and $null -ne $contextProperty.Value) {
            $contextProperty.Value.Dispose()
        }
    }
    catch {
    }
}

function Release-AppConnections {
    param([switch]$PreserveRecoveryState)

    $adminConnection = $script:AdminConnection
    $siteConnection = $script:PnPConnection

    if ($null -ne $adminConnection) {
        Close-AppConnection -Connection $adminConnection
    }

    if ($null -ne $siteConnection -and
        -not [object]::ReferenceEquals($siteConnection, $adminConnection)) {
        Close-AppConnection -Connection $siteConnection
    }

    $script:AdminConnection = $null
    $script:PnPConnection = $null

    if (-not $PreserveRecoveryState) {
        $script:AuthRecoveryUsed = $false
    }
}

function Release-AppSiteConnection {
    if ($null -ne $script:PnPConnection -and
        -not [object]::ReferenceEquals($script:PnPConnection, $script:AdminConnection)) {
        Close-AppConnection -Connection $script:PnPConnection
    }

    $script:PnPConnection = $null
}

function Test-AppConnectionChangeAllowed {
    if ($null -eq $script:TemporarySiteAdmin) {
        return $true
    }

    Write-AppWarning -Message 'Retira primero el acceso temporal antes de cambiar tenant o aplicación.'
    Wait-App
    return $false
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
    return Initialize-ToolkitPnP -MinimumVersion '3.2.0'
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
        [switch]$SkipEnsureAuth,
        [AllowNull()]$ReuseConnection,
        [switch]$ForceAuthentication
    )

    if (-not (Confirm-AppPnP)) {
        throw 'PnP.PowerShell no está disponible.'
    }

    if ($ForceAuthentication) {
        Release-AppConnections -PreserveRecoveryState

        try {
            Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    if (-not $SkipEnsureAuth) {
        if (-not (Ensure-AppAuthentication)) {
            throw 'No existe una autenticación válida configurada.'
        }
    }

    if ($null -eq $ReuseConnection -and $null -ne $script:AdminConnection) {
        $ReuseConnection = $script:AdminConnection
    }

    if ($null -ne $ReuseConnection) {
        try {
            $accessToken = Get-PnPAccessToken `
                -ResourceTypeName SharePoint `
                -Connection $ReuseConnection `
                -ErrorAction Stop

            if ([string]::IsNullOrWhiteSpace([string]$accessToken)) {
                throw 'PnP no devolvió un token de SharePoint reutilizable.'
            }

            return Connect-PnPOnline `
                -Url $Url `
                -AccessToken ([string]$accessToken) `
                -ReturnConnection `
                -ErrorAction Stop
        }
        catch {
            if ($ForceAuthentication -or
                $script:AuthRecoveryUsed -or
                -not (Test-AppAuthenticationRecoveryError -Exception $_.Exception)) {
                throw
            }

            $script:AuthRecoveryUsed = $true
            Write-AppWarning -Message 'La sesión autenticada no pudo reutilizarse; se solicitará autenticación nuevamente una sola vez.'

            return Connect-AppUrl `
                -Url $Url `
                -SkipEnsureAuth:$SkipEnsureAuth `
                -ForceAuthentication
        }
    }

    $params = @{
        Url              = $Url
        ClientId         = $script:Settings.ClientId
        Tenant           = $script:Settings.Tenant
        Interactive      = $true
        ReturnConnection = $true
        ValidateConnection = $true
        ErrorAction      = 'Stop'
    }

    Ensure-AppPersistLoginPreference

    if ($ForceAuthentication) {
        $params.ForceAuthentication = $true
    }

    if ([bool]$script:Settings.PersistLogin) {
        $params.PersistLogin = $true
    }

    $connection = Connect-PnPOnline @params

    if ($null -eq $script:AdminConnection) {
        $script:AdminConnection = $connection
    }

    return $connection
}

function Disconnect-AppM365 {
    Release-AppConnections

    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
    }

    $script:PnPConnection = $null
}

function Ensure-AppPersistLoginPreference {
    if ([bool]$script:Settings.PersistLoginConfigured) {
        return
    }

    Show-AppHeader -Section 'Sesión de autenticación'
    Write-AppMuted -Message 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'
    Write-AppMuted -Message 'Sesión persistente: si la activas, PnP puede reutilizar el inicio de sesión cuando abras el script nuevamente.'
    Write-AppMuted -Message 'Si la desactivas, solo se usa la sesión actual y podrás iniciar sesión de nuevo en la siguiente ejecución.'
    Write-Host ''

    $script:Settings.PersistLogin = Read-AppYesNo `
        -Prompt '¿Guardar la sesión para futuras ejecuciones?' `
        -DefaultYes ([bool]$script:Settings.PersistLogin)
    $script:Settings.PersistLoginConfigured = $true
    Save-AppSettings
}

function Test-AppAuthenticationRecoveryError {
    param([Parameter(Mandatory)][System.Exception]$Exception)

    $text = $Exception.ToString()
    return $text -match '(?i)AADSTS|access token|token|jwt|authentication|interactive|login|no connection|current connection'
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

    $tenant = Read-AppText `
        -Prompt 'Dominio inicial del tenant' `
        -Validator { param($v) Test-AppTenant -Value $v } `
        -ValidationMessage 'Usa el dominio inicial, por ejemplo empresa.onmicrosoft.com.'

    $tenant = $tenant.Trim().ToLowerInvariant()

    if ($tenant -ne [string]$script:Settings.Tenant) {
        Release-AppConnections
    }

    $script:Settings.Tenant = $tenant

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

    if (-not (Test-AppConnectionChangeAllowed)) {
        return
    }

    if (-not (Ensure-AppTenant)) {
        return
    }

    $oldClientId = [string]$script:Settings.ClientId

    $clientId = Read-AppText `
        -Prompt 'Client ID de la aplicación existente' `
        -Validator { param($v) Test-AppGuid -Value $v } `
        -ValidationMessage 'Debe ser un GUID válido.'

    if ($clientId -ne $oldClientId) {
        Release-AppConnections
        $script:Target = $null
    }

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

    if (-not (Test-AppConnectionChangeAllowed)) {
        return
    }

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

        $clientId = Get-ToolkitRegistrationClientId -Result $result

        if (-not (Test-AppGuid -Value $clientId)) {
            $clientId = Read-AppText `
                -Prompt 'Client ID creado' `
                -Validator { param($v) Test-AppGuid -Value $v } `
                -ValidationMessage 'Debe ser un GUID válido.'
        }

        if ($clientId -ne [string]$script:Settings.ClientId) {
            Release-AppConnections
            $script:Target = $null
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

    if (-not (Test-AppConnectionChangeAllowed)) {
        return
    }

    Write-AppWarning -Message 'Esto no elimina la aplicación de Microsoft Entra ID.'
    Write-AppMuted -Message 'Solo elimina el Client ID guardado por esta herramienta.'
    Write-Host ''

    if (Read-AppYesNo -Prompt '¿Continuar?' -DefaultYes $false) {
        Release-AppConnections
        $script:Settings.ClientId = ''
        $script:Target = $null
        Save-AppSettings
        Write-AppOk -Message 'Client ID eliminado de la configuración local.'
    }

    Wait-App
}

function Clear-AppPersistedLogin {
    Show-AppHeader -Section 'Sesión persistente'

    if (-not (Test-AppConnectionChangeAllowed)) {
        return
    }

    try {
        Release-AppConnections
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

        Write-AppField -Name 'Tenant' -Value $(if ($configured) { $script:Settings.Tenant } else { 'No configurado' }) -Style $(if ($configured) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'App guardada' -Value $script:Settings.AppRegistrationName -Style Muted
        Write-AppField -Name 'Client ID' -Value $(if (Test-AppGuid -Value $script:Settings.ClientId) { $script:Settings.ClientId } else { 'No configurado' }) -Style $(if (Test-AppGuid -Value $script:Settings.ClientId) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'Estado config' -Value $(if ($configured) { 'CONFIGURADA' } else { 'NO CONFIGURADA' }) -Style $(if ($configured) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
        Write-AppField -Name 'Persist login' -Value $(if ($script:Settings.PersistLogin) { 'Activado' } else { 'Desactivado' }) -Style $(if ($script:Settings.PersistLogin) { 'Success' } else { 'Muted' })
        Write-AppMuted -Message 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'
        Write-AppMuted -Message 'Persist login controla si PnP puede reutilizar el login al abrir el script nuevamente.'
        Write-Host ''

        if ($configured) {
            $choice = Read-AppMenuChoice -Items @(
                @{ Key='1'; Value='Validate'; Label='Validar aplicación configurada'; Description='Comprueba autenticación y acceso a búsqueda de sitios.' }
                @{ Key='2'; Value='Existing'; Label='Usar otra aplicación existente'; Description='Introduce y opcionalmente valida otro Client ID.' }
                @{ Key='3'; Value='Register'; Label='Registrar una nueva aplicación Entra'; Description='Crea otra app PnP y la configura en esta herramienta.' }
                @{ Key='4'; Value='Remove'; Label='Quitar Client ID de la configuración local'; Description='No elimina la aplicación en Entra.' }
                @{ Key='5'; Value='Clear'; Label='Eliminar sesión persistente'; Description='Cierra la sesión actual y elimina el login persistido por PnP.PowerShell.' }
                @{ Key='0'; Value='Back'; Label='Volver'; Description='Regresa al menú anterior.' }
            )
        }
        else {
            $choice = Read-AppMenuChoice -Items @(
                @{ Key='1'; Value='Existing'; Label='Usar una aplicación existente'; Description='Introduce y opcionalmente valida un Client ID ya registrado.' }
                @{ Key='2'; Value='Register'; Label='Registrar una nueva aplicación Entra'; Description='Crea una nueva app PnP para esta herramienta.' }
                @{ Key='3'; Value='Clear'; Label='Eliminar sesión persistente'; Description='Cierra la sesión actual y elimina el login persistido por PnP.PowerShell.' }
                @{ Key='0'; Value='Back'; Label='Volver'; Description='Regresa al menú anterior.' }
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
        @{ Key='1'; Value='Existing'; Label='Usar una aplicación existente'; Description='Introduce y valida un Client ID ya registrado.' }
        @{ Key='2'; Value='Register'; Label='Registrar una nueva aplicación Entra'; Description='Crea una nueva aplicación PnP para esta herramienta.' }
        @{ Key='0'; Value='Back'; Label='Cancelar'; Description='Regresa al menú anterior.' }
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
                    (Get-ToolkitSiteEndpoint ([string]$_.Url)) -eq 'OneDrive'
                } |
                Sort-Object Owner, Url
        )
    }

    return @(
        $sites |
            Where-Object {
                (Get-ToolkitSiteEndpoint ([string]$_.Url)) -eq 'SharePoint'
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

function Test-AppUserUpn {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return $Value.Trim() -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Get-AppConnectedUserUpn {
    param(
        [Parameter(Mandatory)]$Connection
    )

    foreach ($propertyName in @('PSCredential', 'Credential')) {
        $credentialProperty = $Connection.PSObject.Properties[$propertyName]

        if ($null -eq $credentialProperty -or $null -eq $credentialProperty.Value) {
            continue
        }

        $userNameProperty = $credentialProperty.Value.PSObject.Properties['UserName']

        if ($null -ne $userNameProperty -and
            -not [string]::IsNullOrWhiteSpace([string]$userNameProperty.Value) -and
            [string]$userNameProperty.Value -match '(?<Upn>[^\s|]+@[^\s|]+)') {

            return $Matches['Upn']
        }
    }

    $userProperty = $Connection.PSObject.Properties['User']

    if ($null -ne $userProperty -and $null -ne $userProperty.Value) {
        foreach ($propertyName in @('Email', 'UserPrincipalName', 'LoginName')) {
            $identityProperty = $userProperty.Value.PSObject.Properties[$propertyName]

            if ($null -ne $identityProperty -and
                [string]$identityProperty.Value -match '(?<Upn>[^\s|]+@[^\s|]+)') {

                return $Matches['Upn']
            }
        }
    }

    return ''
}

function Add-AppTemporarySiteCollectionAdmin {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [AllowNull()]$SiteConnection
    )

    $detectedUpn = if ($null -ne $SiteConnection) {
        Get-AppConnectedUserUpn -Connection $SiteConnection
    }
    else {
        ''
    }
    $defaultUpn = if (Test-AppUserUpn -Value $detectedUpn) {
        $detectedUpn
    }
    elseif (Test-AppUserUpn -Value ([string]$script:Settings.AdminUpn)) {
        [string]$script:Settings.AdminUpn
    }
    else {
        ''
    }

    if (-not [string]::IsNullOrWhiteSpace($defaultUpn)) {
        Write-AppMuted -Message "UPN detectado: $defaultUpn"
    }
    else {
        Write-AppWarning -Message 'No se pudo detectar el UPN de la cuenta conectada.'
    }

    $upn = Read-AppText `
        -Prompt 'UPN de la cuenta que se agregará como Site Collection Administrator' `
        -Default $defaultUpn `
        -Validator { param($value) Test-AppUserUpn -Value $value } `
        -ValidationMessage 'Introduce un UPN válido.'

    if (-not (Read-AppYesNo -Prompt "¿Agregar $upn como Site Collection Administrator?" -DefaultYes $true)) {
        return $null
    }

    $adminUrl = Get-AppAdminUrl

    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        throw 'No fue posible determinar la URL del centro de administración.'
    }

    Write-AppInfo -Message "Conectando al centro de administración: $adminUrl"
    $adminConnection = Connect-AppUrl `
        -Url $adminUrl `
        -SkipEnsureAuth `
        -ReuseConnection $SiteConnection

    # Si la conexión inicial al sitio no llegó a crearse, la conexión al
    # Admin Center se convierte en la conexión base para la validación. Así
    # tampoco se abre una autenticación interactiva en cada reintento.
    $authenticationConnection = if ($null -ne $SiteConnection) {
        $SiteConnection
    }
    else {
        $adminConnection
    }

    Write-AppInfo -Message "Se agregará $upn al sitio como Site Collection Administrator."

    Set-PnPTenantSite `
        -Identity $SiteUrl `
        -Owners @($upn) `
        -Connection $adminConnection `
        -ErrorAction Stop

    # Track the grant immediately. If propagation takes longer than expected,
    # the user can still remove the temporary administrator from the main menu.
    $script:TemporarySiteAdmin = [pscustomobject]@{
        Upn        = $upn
        SiteUrl    = $SiteUrl
        Connection = $SiteConnection
    }

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        Write-AppInfo -Message 'Esperando propagación del acceso administrativo...'
        Start-Sleep -Seconds 3

        try {
            $refreshedConnection = Connect-AppUrl `
                -Url $SiteUrl `
                -SkipEnsureAuth `
                -ReuseConnection $authenticationConnection
            $access = Test-AppRecycleBinAccess -Connection $refreshedConnection

            if ($access.Success) {
                Write-AppOk -Message 'Se agregó acceso temporal como Site Collection Administrator.'
                Write-AppMuted -Message 'El acceso temporal se mantendrá hasta que lo retires desde el menú principal.'

                return [pscustomobject]@{
                    Upn        = $upn
                    SiteUrl    = $SiteUrl
                    Connection = $refreshedConnection
                }
            }
        }
        catch {
            $access = [pscustomobject]@{
                Success = $false
                Message = $_.Exception.Message
            }
        }

        if ($attempt -lt 6) {
            Write-AppMuted -Message 'Reintentando validación de acceso...'
        }
    }

    throw 'No se pudo confirmar el acceso administrativo después de agregar el usuario.'
}

function Request-AppTemporarySiteCollectionAdmin {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [AllowNull()]$SiteConnection,
        [Parameter(Mandatory)][string]$ErrorMessage
    )

    if ($ErrorMessage -notmatch '(?i)unauthori[sz]ed|access denied|forbidden|\b401\b|\b403\b|attempted to perform') {
        return $null
    }

    if ($null -ne $script:TemporarySiteAdmin) {
        if (([string]$script:TemporarySiteAdmin.SiteUrl).TrimEnd('/') -eq $SiteUrl.TrimEnd('/')) {
            return $script:TemporarySiteAdmin
        }

        return $null
    }

    if (-not (Read-AppYesNo -Prompt '¿Deseas agregarte temporalmente como Site Collection Administrator?' -DefaultYes $true)) {
        return $null
    }

    try {
        return Add-AppTemporarySiteCollectionAdmin `
            -SiteUrl $SiteUrl `
            -SiteConnection $SiteConnection
    }
    catch {
        Show-AppErrorScreen `
            -Title 'Agregar Site Collection Administrator' `
            -Message "La cuenta conectada no tiene permisos para agregar Site Collection Administrators desde el centro de administración. $($_.Exception.Message)"
        return $null
    }
}

function Remove-AppTemporarySiteCollectionAdmin {
    param(
        [switch]$Ask
    )

    if ($null -eq $script:TemporarySiteAdmin) {
        return $true
    }

    $temporaryAdmin = $script:TemporarySiteAdmin

    if ($Ask -and -not (Read-AppYesNo -Prompt '¿Retirar el acceso temporal ahora?' -DefaultYes $true)) {
        Write-AppWarning -Message 'No se retiró el acceso temporal.'
        return $false
    }

    try {
        $connection = Connect-AppUrl `
            -Url $temporaryAdmin.SiteUrl `
            -SkipEnsureAuth `
            -ReuseConnection $temporaryAdmin.Connection

        Remove-PnPSiteCollectionAdmin `
            -Owners @($temporaryAdmin.Upn) `
            -Connection $connection `
            -ErrorAction Stop

        $script:TemporarySiteAdmin = $null
        $script:PnPConnection = $null
        $script:Target = $null
        Write-AppOk -Message 'Acceso temporal retirado correctamente.'
        return $true
    }
    catch {
        Write-AppError -Message 'No se pudo retirar el acceso temporal.'
        Write-AppMuted -Message $_.Exception.Message
        return $false
    }
}

function Set-AppTarget {
    if ($null -ne $script:TemporarySiteAdmin) {
        Write-AppWarning -Message 'Hay un acceso temporal de Site Collection Admin activo para este sitio.'

        if (-not (Read-AppYesNo -Prompt '¿Retirar el acceso temporal antes de cambiar de sitio?' -DefaultYes $true)) {
            Write-AppWarning -Message 'No se puede continuar con otro sitio mientras el acceso temporal siga activo.'
            Wait-App
            return
        }

        if (-not (Remove-AppTemporarySiteCollectionAdmin)) {
            Wait-App
            return
        }
    }

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

    Release-AppSiteConnection
    $connection = $null
    $temporaryAdmin = $null
    $loadTarget = {
        if ($null -eq $connection) {
            throw 'No se obtuvo una conexión PnP válida.'
        }

        $web = Get-PnPWeb `
            -Connection $connection `
            -Includes Title,Url `
            -ErrorAction Stop

        Write-AppOk -Message 'Conexión establecida.'
        Write-AppMuted -Message $web.Title

        Write-AppInfo -Message 'Validando acceso a la papelera...'

        $access = Test-AppRecycleBinAccess -Connection $connection

        if (-not $access.Success) {
            Write-AppError -Message 'La cuenta conectada no es Site Collection Administrator para este sitio.'
            Write-AppMuted -Message 'Get-PnPRecycleBinItem requiere Site Collection Administrator.'
            Write-AppMuted -Message $access.Message
            # Mark this failure explicitly as an authorization failure. PnP can
            # return different localized messages for OneDrive recycle-bin access.
            throw "Access denied: $($access.Message)"
        }

        Write-AppOk -Message 'Acceso a la papelera validado.'

        return [pscustomobject]@{
            Web        = $web
            Connection = $connection
        }
    }

    try {
        $connection = Connect-AppUrl -Url $url
        $targetData = & $loadTarget
    }
    catch {
        $message = $_.Exception.Message
        $temporaryAdmin = Request-AppTemporarySiteCollectionAdmin `
            -SiteUrl $url `
            -SiteConnection $connection `
            -ErrorMessage $message

        if ($null -eq $temporaryAdmin) {
            Show-AppErrorScreen -Title 'Conexión al destino' -Message $message
            return
        }

        try {
            $connection = $temporaryAdmin.Connection
            $targetData = & $loadTarget
        }
        catch {
            Show-AppErrorScreen -Title 'Validación del destino' -Message $_.Exception.Message
            return
        }
    }

    try {
        $web = $targetData.Web
        $connection = $targetData.Connection

        $owner = ''
        $ownerProperty = $site.PSObject.Properties['Owner']
        if ($null -ne $ownerProperty) {
            $owner = [string]$ownerProperty.Value
        }

        $script:PnPConnection = $connection

        if ($null -ne $temporaryAdmin) {
            $script:TemporarySiteAdmin = $temporaryAdmin
        }

        $script:Target = [pscustomobject]@{
            Type    = $type
            Url     = $url.TrimEnd('/')
            Title   = [string]$web.Title
            Owner   = $owner
        }

        Write-AppOk -Message "Destino seleccionado: $($script:Target.Title)"
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

    if ((Get-ToolkitSiteEndpoint ([string]$script:Target.Url)) -notin @('SharePoint','OneDrive')) {
        Write-AppError -Message (Get-ToolkitString Invalid)
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
    $action = switch ($Phrase) {
        'MOVER' { 'Move' }
        'ELIMINAR' { 'Delete' }
        'ELIMINAR PHL' { 'DeletePHL' }
        default { throw 'Unknown destructive confirmation action.' }
    }
    return Confirm-ToolkitAction -Action $action
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

    Write-AppField -Name 'Sitio' -Value $script:Target.Url -Style Primary
    Write-AppField -Name 'Nivel' -Value $stageLabel -Style Primary
    Write-AppField -Name 'Acción' -Value $operationLabel -Style $(if ($isFirstStage) { 'Warning' } else { 'Danger' })
    Write-AppField -Name 'Límite' -Value $Limit -Style Primary
    Write-AppField -Name 'Modo' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' }) -Style $(if ($script:Settings.Simulation) { 'Warning' } else { 'Danger' })
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

    Write-AppField -Name 'Encontrados' -Value $items.Count -Style Primary
    Write-AppField -Name 'Procesados' -Value $log.Count -Style Primary

    if ($isFirstStage) {
        Write-AppField -Name 'Movidos' -Value $moved -Style Success
    }
    else {
        Write-AppField -Name 'Eliminados' -Value $removed -Style Danger
    }

    Write-AppField -Name 'Simulados' -Value $simulated -Style Warning
    Write-AppField -Name 'Errores' -Value $errors -Style $(if ($errors -gt 0) { 'Danger' } else { 'Success' })
    Write-AppField -Name 'Reporte' -Value $report -Style Muted

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

    Write-AppField -Name 'Sitio' -Value $script:Target.Url -Style Primary
    Write-AppField -Name 'Límite' -Value $Limit -Style Primary
    Write-AppField -Name 'Modo' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' }) -Style $(if ($script:Settings.Simulation) { 'Warning' } else { 'Danger' })
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

    Write-AppField -Name 'Encontrados' -Value $items.Count -Style Primary
    Write-AppField -Name 'Procesados' -Value $log.Count -Style Primary
    Write-AppField -Name 'Eliminados' -Value $removed -Style Danger
    Write-AppField -Name 'Simulados' -Value $simulated -Style Warning
    Write-AppField -Name 'Errores' -Value $errors -Style $(if ($errors -gt 0) { 'Danger' } else { 'Success' })
    Write-AppField -Name 'Reporte' -Value $report -Style Muted
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
            $firstStageResult = Invoke-AppRecycleBinCleanup -Stage FirstStage -Limit $limit
            if ($null -eq $firstStageResult) { return }
            [void](Invoke-AppRecycleBinCleanup -Stage SecondStage -Limit $limit)
        }

        'Preservation' {
            Invoke-AppPreservationHoldCleanup -Limit $limit
        }
    }

    if ($null -ne $script:TemporarySiteAdmin) {
        Write-AppInfo -Message 'La operación terminó. Puedes retirar ahora el acceso temporal.'
        [void](Remove-AppTemporarySiteCollectionAdmin -Ask)

        if ($null -ne $script:TemporarySiteAdmin) {
            Write-AppWarning -Message 'El acceso temporal sigue activo.'
        }

        Wait-App
    }
}


# =============================================================================
# CONFIGURACION / DIAGNOSTICO
# =============================================================================

function Clear-AppWorkingContext {
    Show-AppHeader -Section 'Limpiar contexto de trabajo'

    Write-AppField -Name 'Destino' -Value $(if ($null -ne $script:Target) { $script:Target.Title } else { 'No seleccionado' }) -Style $(if ($null -ne $script:Target) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Conexión de sitio' -Value $(if ($null -ne $script:PnPConnection) { 'ACTIVA' } else { 'VACÍA' }) -Style $(if ($null -ne $script:PnPConnection) { 'Success' } else { 'Muted' })
    Write-Host ''
    Write-AppMuted -Message 'Limpia el sitio seleccionado y la conexión de sitio. Conserva tenant, aplicación, preferencias y reportes.'
    Write-AppMuted -Message 'La sesión de autenticación base se conserva; cambiar tenant o aplicación la renovará.'

    if ($null -ne $script:TemporarySiteAdmin) {
        Write-AppWarning -Message 'No se puede limpiar el contexto mientras exista un acceso temporal activo.'
        Write-AppMuted -Message 'Retíralo primero desde el menú principal.'
        Wait-App
        return
    }

    if (-not (Read-AppYesNo -Prompt '¿Limpiar el contexto de trabajo?' -DefaultYes $false)) {
        return
    }

    Release-AppSiteConnection
    $script:Target = $null

    Write-AppOk -Message 'Contexto de trabajo limpiado.'
    Wait-App
}

function Show-AppContextMenu {
    while ($true) {
        Show-AppHeader -Section 'Gestión del contexto'

        $contextActive = ($null -ne $script:Target) -or ($null -ne $script:PnPConnection)
        $tenantConfigured = Test-AppTenant -Value $script:Settings.Tenant
        $appConfigured = $tenantConfigured -and (Test-AppGuid -Value $script:Settings.ClientId)
        $targetLabel = if ($null -ne $script:Target) { $script:Target.Title } else { 'No seleccionado' }
        $appLabel = if (-not [string]::IsNullOrWhiteSpace([string]$script:Settings.AppRegistrationName)) {
            $script:Settings.AppRegistrationName
        }
        elseif ($appConfigured) {
            $script:Settings.ClientId
        }
        else {
            'No configurado'
        }

        Write-AppField -Name 'Contexto' -Value $(if ($contextActive) { 'ACTIVO' } else { 'VACÍO' }) -Style $(if ($contextActive) { 'Success' } else { 'Muted' })
        Write-AppField -Name 'Destino' -Value $targetLabel -Style $(if ($null -ne $script:Target) { 'Success' } else { 'Muted' })
        Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $script:Settings.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'Aplicación' -Value $appLabel -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
        Write-Host ''

        $choice = Read-AppMenuChoice -Items @(
            @{ Key='1'; Value='Clear'; Label='Limpiar contexto de trabajo'; Description='Limpia el destino y la conexión de sitio, pero conserva tenant, aplicación y reportes.' }
            @{ Key='2'; Value='Connection'; Label='Cambiar tenant o aplicación conectada'; Description='Abre la configuración y libera la conexión anterior cuando cambies estos valores.' }
            @{ Key='0'; Value='Back'; Label='Volver'; Description='Regresa al menú anterior.' }
        )

        switch ($choice.Value) {
            'Clear' { Clear-AppWorkingContext }
            'Connection' { Show-AppSettingsMenu }
            'Back' { return }
        }
    }
}

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

        $tenantConfigured = Test-AppTenant -Value $script:Settings.Tenant
        $appConfigured = $tenantConfigured -and (Test-AppGuid -Value $script:Settings.ClientId)

        Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $script:Settings.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'Client ID' -Value $(if ($appConfigured) { $script:Settings.ClientId } else { 'No configurado' }) -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
        Write-AppField -Name 'Idioma' -Value $(if ($script:Language -eq 'en') { 'English' } else { 'Español' }) -Style Primary
        Write-AppField -Name 'Persist login' -Value $(if ($script:Settings.PersistLogin) { 'Activado' } else { 'Desactivado' }) -Style $(if ($script:Settings.PersistLogin) { 'Success' } else { 'Muted' })
        Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
        Write-AppField -Name 'Modo inicial' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' }) -Style $(if ($script:Settings.Simulation) { 'Warning' } else { 'Danger' })
        Write-AppField -Name 'Límite defecto' -Value $script:Settings.DefaultLimit -Style Primary
        Write-AppField -Name 'Límite máximo' -Value $script:Settings.MaximumLimit -Style Primary
        Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
        Write-AppMuted -Message 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'
        Write-AppMuted -Message 'Persist login controla si PnP puede reutilizar el login al abrir el script nuevamente.'
        Write-Host ''

        $choice = Read-AppMenuChoice -Items @(
            @{ Key='1'; Value='Tenant'; Label='Cambiar tenant'; Description='Actualiza el tenant conectado y libera la sesión anterior.' }
            @{ Key='2'; Value='Auth'; Label='Aplicación Entra / autenticación PnP'; Description='Valida, cambia o registra la aplicación usada por PnP.PowerShell.' }
            @{ Key='3'; Value='Language'; Label='Cambiar idioma'; Description='Cambia el idioma de la interfaz.' }
            @{ Key='4'; Value='Persist'; Label='Alternar persistencia de login'; Description='Controla si PnP puede reutilizar el login cuando abras el script nuevamente.' }
            @{ Key='5'; Value='DefaultLimit'; Label='Cambiar límite predeterminado'; Description='Cantidad inicial de elementos a procesar por operación.' }
            @{ Key='6'; Value='MaxLimit'; Label='Cambiar límite máximo permitido'; Description='Rango permitido por la herramienta: 1 a 10000.' }
            @{ Key='7'; Value='Throttle'; Label='Throttling y reintentos'; Description='Ajusta pausas, backoff y cantidad de reintentos.' }
            @{ Key='8'; Value='Reports'; Label='Abrir carpeta de reportes'; Description='Abre la carpeta donde se guardan los CSV de auditoría.' }
            @{ Key='9'; Value='Reset'; Label='Restablecer configuración local'; Description='Borra valores guardados; conserva idioma y reportes.' }
            @{ Key='0'; Value='Back'; Label='Volver'; Description='Regresa al menú anterior.' }
        )

        switch ($choice.Value) {
            'Tenant' {
                Show-AppHeader -Section 'Tenant'

                if (-not (Test-AppConnectionChangeAllowed)) {
                    continue
                }

                $tenant = Read-AppText `
                    -Prompt 'Dominio inicial del tenant' `
                    -Default $script:Settings.Tenant `
                    -Validator { param($value) Test-AppTenant -Value $value } `
                    -ValidationMessage 'Usa el dominio inicial, por ejemplo empresa.onmicrosoft.com.'

                $tenant = $tenant.Trim().ToLowerInvariant()
                if ($tenant -ne [string]$script:Settings.Tenant) {
                    Release-AppConnections
                    $script:Target = $null
                }

                $script:Settings.Tenant = $tenant
                Save-AppSettings
            }

            'Auth' {
                Show-AppAuthenticationMenu
            }

            'Language' {
                Initialize-AppLanguage
                $script:Settings.Language = $script:Language
                Save-AppSettings
            }

            'Persist' {
                $script:Settings.PersistLogin = -not [bool]$script:Settings.PersistLogin
                $script:Settings.PersistLoginConfigured = $true
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
                    $canReset = $true

                    if ($null -ne $script:TemporarySiteAdmin) {
                        $canReset = Remove-AppTemporarySiteCollectionAdmin -Ask
                    }

                    if ($canReset) {
                        $language = $script:Language
                        Release-AppConnections

                        if (Test-Path -LiteralPath $script:ConfigPath) {
                            Remove-Item -LiteralPath $script:ConfigPath -Force
                        }

                        $script:Settings = Get-DefaultAppSettings
                        $script:Settings.Language = $language
                        $script:Target = $null
                        $script:Language = $language

                        Write-AppOk -Message 'Configuración restablecida.'
                        Wait-App
                    }
                    else {
                        Write-AppWarning -Message 'El acceso temporal sigue activo.'
                        Wait-App
                    }
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

    $tenantConfigured = Test-AppTenant -Value $script:Settings.Tenant
    $appConfigured = $tenantConfigured -and (Test-AppGuid -Value $script:Settings.ClientId)

    Write-AppField -Name 'PowerShell' -Value $PSVersionTable.PSVersion -Style $(if ($PSVersionTable.PSVersion -ge [version]'7.4') { 'Success' } else { 'Danger' })
    Write-AppField -Name 'PnP.PowerShell' -Value $(if ($null -ne $pnpVersion) { $pnpVersion } else { 'No instalado' }) -Style $(if ($null -ne $pnpVersion -and $pnpVersion -ge $script:MinimumPnPVersion) { 'Success' } else { 'Warning' })
    Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $script:Settings.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })
    Write-AppField -Name 'Client ID' -Value $(if ($appConfigured) { $script:Settings.ClientId } else { 'No configurado' }) -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
    Write-AppField -Name 'Persist login' -Value $(if ($script:Settings.PersistLogin) { 'Activado' } else { 'Desactivado' }) -Style $(if ($script:Settings.PersistLogin) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Config' -Value $script:ConfigPath -Style Muted
    Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
    Write-AppField -Name 'Throttling' -Value $script:ThrottleEvents.Count -Style $(if ($script:ThrottleEvents.Count -eq 0) { 'Success' } else { 'Warning' })
    Write-Host ''

    $choice = Read-AppMenuChoice -Items @(
        @{ Key='1'; Value='Validate'; Label='Validar aplicación configurada'; Description='Comprueba tenant, Client ID, autenticación y acceso a búsqueda de sitios.' }
        @{ Key='2'; Value='PnP'; Label='Instalar / actualizar PnP.PowerShell'; Description='Verifica el requisito local y ofrece instalar o actualizar el módulo.' }
        @{ Key='3'; Value='Target'; Label='Probar nuevamente el sitio actual'; Description='Valida conexión y permisos de la papelera para el destino seleccionado.' }
        @{ Key='0'; Value='Back'; Label='Volver'; Description='Regresa al menú anterior.' }
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

    $tenantConfigured = Test-AppTenant -Value $script:Settings.Tenant
    $appConfigured = $tenantConfigured -and (Test-AppGuid -Value $script:Settings.ClientId)

    Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $script:Settings.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })

    if ($null -ne $script:Target) {
        Write-AppField -Name 'Tipo' -Value $script:Target.Type -Style Muted
        Write-AppField -Name 'Sitio' -Value $script:Target.Title -Style Primary
        Write-AppField -Name 'URL' -Value $script:Target.Url -Style Muted
    }
    else {
        Write-AppField -Name 'Destino' -Value 'No seleccionado' -Style Muted
    }

    Write-AppField -Name 'Autenticación' -Value $(if ($appConfigured) { 'CONFIGURADA' } else { 'NO CONFIGURADA' }) -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
    Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Admin temporal' -Value $(if ($null -ne $script:TemporarySiteAdmin) { 'ACTIVO' } else { 'No' }) -Style $(if ($null -ne $script:TemporarySiteAdmin) { 'Warning' } else { 'Success' })
    Write-AppField -Name 'Modo' -Value $(if ($script:Settings.Simulation) { 'SIMULACIÓN' } else { 'REAL' }) -Style $(if ($script:Settings.Simulation) { 'Warning' } else { 'Danger' })
    Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
}

function Show-AppMainMenu {
    while ($true) {
        Show-AppHeader -Section 'Inicio'
        Show-AppContext
        Write-Host ''

        $appConfigured = (Test-AppTenant -Value $script:Settings.Tenant) -and (Test-AppGuid -Value $script:Settings.ClientId)
        $contextActive = ($null -ne $script:Target) -or ($null -ne $script:PnPConnection)

        $choiceItems = @(
            @{ Key='1'; Value='Target'; Label='Seleccionar / cambiar sitio'; Description='Buscar SharePoint u OneDrive en el tenant o usar una URL directa.' }
            @{ Key='2'; Value='First'; Label='Vaciar primer nivel'; Description='Mueve los elementos al segundo nivel; no los elimina permanentemente.' }
            @{ Key='3'; Value='Second'; Label='Vaciar segundo nivel'; Description='Elimina permanentemente elementos del segundo nivel.' }
            @{ Key='4'; Value='Both'; Label='Vaciar ambos niveles'; Description='Primero mueve el primer nivel al segundo y después elimina del segundo; el límite se aplica por nivel.' }
            @{ Key='5'; Value='PHL'; Label='Procesar Preservation Hold Library'; Description='Detecta dinámicamente la biblioteca; no modifica políticas de retención.' }
            @{ Key='6'; Value='Mode'; Label=$(if ($script:Settings.Simulation) { 'Cambiar a modo REAL' } else { 'Cambiar a modo SIMULACIÓN' }); Description=$(if ($script:Settings.Simulation) { 'SIMULACIÓN es el modo seguro.' } else { 'REAL elimina o modifica elementos según la operación elegida.' }) }
            @{ Key='7'; Value='Context'; Label='Gestionar contexto de trabajo'; Description=$(if ($contextActive) { 'Contexto activo · limpia el destino o cambia tenant/aplicación conectada.' } else { 'No hay contexto activo · limpia selecciones o cambia tenant/aplicación conectada.' }) }
            @{ Key='8'; Value='Auth'; Label='Aplicación Entra / autenticación PnP'; Description=$(if ($appConfigured) { 'CONFIGURADA · valida, cambia o registra otra aplicación.' } else { 'NO CONFIGURADA · usa una aplicación existente o registra una nueva.' }) }
            @{ Key='9'; Value='Settings'; Label='Configuración'; Description='Tenant, límites, throttling, persistencia, idioma y reportes.' }
            @{ Key='10'; Value='Diagnostics'; Label='Diagnóstico'; Description='Valida PowerShell, PnP, autenticación y acceso al sitio actual.' }
        )

        if ($null -ne $script:TemporarySiteAdmin) {
            $choiceItems += @{ Key='11'; Value='RemoveTemporaryAdmin'; Label='Retirar Site Collection Admin temporal'; Description="Retira $($script:TemporarySiteAdmin.Upn) de $($script:TemporarySiteAdmin.SiteUrl)." }
        }

        $choiceItems += @{ Key='0'; Value='Exit'; Label='Salir'; Description='Cierra la herramienta.' }
        $choice = Read-AppMenuChoice -Items $choiceItems

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

            'Context' {
                Show-AppContextMenu
            }

            'Exit' {
                if ($null -ne $script:TemporarySiteAdmin) {
                    if (-not (Remove-AppTemporarySiteCollectionAdmin -Ask)) {
                        Write-AppWarning -Message 'El acceso temporal sigue activo.'
                        Wait-App
                        continue
                    }
                }

                return
            }

            'RemoveTemporaryAdmin' {
                [void](Remove-AppTemporarySiteCollectionAdmin -Ask)
                Wait-App
            }
        }
    }
}


# =============================================================================
# INICIO
# =============================================================================

try {
    Initialize-AppTerminal
    Initialize-AppLanguage
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

    if ($DebugPreference -ne 'SilentlyContinue' -and -not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
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
