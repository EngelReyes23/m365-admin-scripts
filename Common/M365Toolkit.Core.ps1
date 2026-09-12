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
