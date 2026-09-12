# Canonical adapters for the three read-only tenant reporters. Embedded, not imported.
function Get-PropertyValue {
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$Name, [AllowNull()]$Default = $null)
    if ($Object -is [Collections.IDictionary]) {
        if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] }
        return $Default
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Read-MenuChoice {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Items, [switch]$AllowBack, [scriptblock]$RenderBody)
    $normalized = @(
        foreach ($item in $Items) {
            $key = [string](Get-PropertyValue $item Key '')
            $label = [string](Get-PropertyValue $item Label '')
            if ($key -notmatch '^\d+$' -or -not $label) { throw (Get-Text 'Cada opción requiere Key numérico y Label.' 'Every option requires a numeric Key and Label.') }
            [pscustomobject]@{ Key=$key; Label=$label; Description=[string](Get-PropertyValue $item Description '') }
        }
    )
    if ($RenderBody) { & $RenderBody; Write-ToolkitText '' }
    foreach ($item in $normalized) {
        Write-ToolkitText ('  [{0}] {1}' -f $item.Key,$item.Label) Primary
        if ($item.Description) { Write-ToolkitText ('      {0}' -f $item.Description) Muted }
        Write-ToolkitText ''
    }
    $valid = @($normalized | ForEach-Object { $_.Key })
    if ($AllowBack -and -not ($valid -contains '0')) {
        Write-ToolkitText ('  [0] {0}' -f (Get-ToolkitString Back)) Muted
        Write-ToolkitText ''
    }
    if ($AllowBack) { $valid += '0' }
    while ($true) {
        $choice = Read-ToolkitInput ('  ' + (Get-ToolkitString Select))
        if ($valid -contains $choice) { return $choice }
        Write-Status Error (Get-ToolkitString Invalid)
    }
}

function Test-IsOneDriveUrl {
    param([AllowNull()][string]$Url)
    return (Get-ToolkitSiteEndpoint $Url) -eq 'OneDrive'
}

function Test-IsSystemSite {
    param([Parameter(Mandatory)]$Site)
    $url = [string](Get-PropertyValue $Site Url '')
    if ((Get-ToolkitSiteEndpoint $url) -in @('Invalid','Admin','OneDriveSystem')) { return $true }
    $template = [string](Get-PropertyValue $Site Template '')
    if ($template -match '^(REDIRECTSITE|SPSMSITEHOST|APPCATALOG)(#\d+)?$') { return $true }
    # Exact paths only: search-project/appcatalogue and locked customer sites remain legitimate.
    $path = ([uri]$url).AbsolutePath.TrimEnd('/')
    return $path -in @('/search','/sites/appcatalog','/sites/contenttypehub')
}

function Get-AdminUrl {
    $configured = [string](Get-PropertyValue $script:Settings AdminUrl '')
    $configuredTenant = [string](Get-PropertyValue $script:Settings AdminUrlTenant '')
    if ($configuredTenant -eq $script:Settings.Tenant -and (Get-ToolkitSiteEndpoint $configured) -eq 'Admin') { return $configured.TrimEnd('/') }
    $prefix = Get-TenantPrefix
    if (-not $prefix) { return '' }
    return "https://$prefix-admin.sharepoint.com"
}

function Ensure-TenantConfigured {
    if (-not (Test-TenantFormat $script:Settings.Tenant)) {
        $tenant = Read-TextValue -Prompt (Get-Text 'Tenant inicial (Enter para cancelar)' 'Initial tenant (Enter to cancel)') -AllowEmpty -Validator { param($v) Test-TenantFormat $v } -ValidationMessage (Get-Text 'Usa el dominio inicial de Entra terminado en .onmicrosoft.com.' 'Use the initial Entra domain ending in .onmicrosoft.com.')
        if (-not $tenant) { return $false }
        $script:Settings.Tenant = $tenant.ToLowerInvariant()
    }
    if ([string](Get-PropertyValue $script:Settings AdminUrlTenant '') -ne $script:Settings.Tenant) {
        $url = Read-TextValue -Prompt (Get-Text 'URL del centro de administración SharePoint (confirma si el dominio fue renombrado)' 'SharePoint admin center URL (check if the domain was renamed)') -Default (Get-AdminUrl) -Validator { param($v) (Get-ToolkitSiteEndpoint $v) -eq 'Admin' -and ([uri]$v).AbsolutePath -eq '/' -and -not ([uri]$v).Query -and -not ([uri]$v).Fragment } -ValidationMessage (Get-Text 'Introduce la URL HTTPS del centro de administración SharePoint.' 'Enter the HTTPS SharePoint admin center URL.')
        $script:Settings | Add-Member -NotePropertyName AdminUrl -NotePropertyValue $url.TrimEnd('/') -Force
        $script:Settings | Add-Member -NotePropertyName AdminUrlTenant -NotePropertyValue $script:Settings.Tenant -Force
        Save-Settings
    }
    return $true
}

function Ensure-ClientId {
    if (Test-ClientIdFormat $script:Settings.ClientId) { return $true }
    Write-Section (Get-Text 'Aplicación PnP' 'PnP Application')
    $choice = Read-MenuChoice -Items @(
        [pscustomobject]@{Key='1';Label=(Get-Text 'Usar una aplicación existente' 'Use an existing application')},
        [pscustomobject]@{Key='2';Label=(Get-Text 'Registrar una nueva aplicación' 'Register a new application')},
        [pscustomobject]@{Key='0';Label=(Get-ToolkitString Cancel)}
    )
    switch ($choice) {
        '1' { return Set-ExistingPnPClientId }
        '2' { return Register-NewPnPApp }
        default { return $false }
    }
}

function Test-PnPAppRegistration {
    param([Parameter(Mandatory)][string]$ClientId, [switch]$Quiet)
    if (-not (Test-ClientIdFormat $ClientId)) {
        if (-not $Quiet) { Write-Status Error (Get-Text 'El Client ID no es un GUID válido.' 'The Client ID is not a valid GUID.') }
        return $false
    }
    $connection = $null
    try {
        if (-not $Quiet) { Write-Status Info (Get-Text 'Validando autenticación y acceso al inventario del tenant...' 'Validating authentication and access to the tenant inventory...') }
        $connection = Connect-M365AdminWithClientId -ClientId $ClientId
        if ($null -eq $connection) { return $false }
        $null = Get-PnPTenant -Connection $connection -ErrorAction Stop
        $null = Get-PnPTenantSite -Connection $connection -ErrorAction Stop | Select-Object -First 1
        if (-not $Quiet) { Write-Status Ok (Get-Text 'Aplicación y acceso validados.' 'Application and access validated.') }
        return $true
    }
    catch {
        if (-not $Quiet) { Write-Status Error (Get-Text "Falló la validación de la aplicación: $($_.Exception.Message)" "Application validation failed: $($_.Exception.Message)") }
        return $false
    }
    finally {
        if ($null -ne $connection) {
            try { Disconnect-PnPOnline -Connection $connection -ErrorAction Stop }
            catch { if (-not $Quiet) { Write-Status Warn (Get-Text "No se pudo cerrar la conexión de prueba: $($_.Exception.Message)" "Could not close the test connection: $($_.Exception.Message)") } }
        }
    }
}

function Set-ExistingPnPClientId {
    if (-not (Ensure-TenantConfigured)) { return $false }
    $current = if (Test-ClientIdFormat $script:Settings.ClientId) { $script:Settings.ClientId } else { '' }
    $clientId = Read-TextValue -Prompt (Get-Text 'Client ID existente (Enter sin valor para cancelar)' 'Existing Client ID (Enter with no value to cancel)') -Default $current -AllowEmpty -Validator { param($v) Test-ClientIdFormat $v } -ValidationMessage (Get-Text 'Debe ser un GUID válido distinto de cero.' 'A valid nonzero GUID is required.')
    if (-not $clientId) { return $false }
    if (-not (Test-PnPAppRegistration -ClientId $clientId)) { return $false }
    $script:Settings.ClientId = $clientId
    Save-Settings
    Write-Status Ok (Get-Text 'Client ID validado y guardado.' 'Client ID validated and saved.')
    return $true
}

function Register-NewPnPApp {
    if (-not (Ensure-PnPModule -OfferInstall)) { return $false }
    if (-not (Ensure-TenantConfigured)) { return $false }
    Write-Section (Get-Text 'Registrar nueva aplicación Entra' 'Register new Entra application')
    $appName = Read-TextValue -Prompt (Get-Text 'Nombre de la aplicación' 'Application name') -Default $script:AppName
    Write-Styled (Get-Text 'Creará una aplicación Entra. Solicitará AllSites.FullControl delegado de SharePoint y User.Read delegado de Graph. Requiere autorización para registrar aplicaciones y consentimiento administrativo. No cambia sitios; no crea ni guarda secretos.' 'Creates an Entra application. Requests delegated SharePoint AllSites.FullControl and delegated Graph User.Read. Requires authorization to register applications and administrator consent. Does not change sites; creates or saves no secrets.') Muted
    if (-not (Read-YesNo (Get-Text '¿Registrar esta aplicación?' 'Register this application?'))) { return $false }
    try {
        $result = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName $appName -Tenant $script:Settings.Tenant -SharePointDelegatePermissions 'AllSites.FullControl' -GraphDelegatePermissions 'User.Read' -ErrorAction Stop
        $candidate = ''
        foreach ($item in @($result)) {
            if ($null -eq $item) { continue }
            if (($item -is [string] -or $item -is [guid]) -and (Test-ClientIdFormat ([string]$item))) { $candidate = [string]$item; break }
            # Id alone may be the directory object ID, NOT the Application (client) ID.
            foreach ($property in @('ClientId','AppId','ApplicationId')) {
                $value = [string](Get-PropertyValue $item $property '')
                if (Test-ClientIdFormat $value) { $candidate = $value; break }
            }
            if ($candidate) { break }
        }
        if (-not $candidate) {
            $candidate = Read-TextValue -Prompt (Get-Text 'Application (client) ID creado (Enter para cancelar)' 'Created Application (client) ID (Enter to cancel)') -AllowEmpty -Validator { param($v) Test-ClientIdFormat $v } -ValidationMessage (Get-Text 'Debe ser un GUID válido.' 'A valid GUID is required.')
        }
        if (-not $candidate) { return $false }
        Write-Field 'Client ID' $candidate
        if (-not (Test-PnPAppRegistration -ClientId $candidate)) {
            Write-Status Warn (Get-Text 'La aplicación puede haberse creado, pero no se cambió el Client ID guardado. Revisa consentimiento/permisos y vuelve a validar usando el ID mostrado.' 'The application may have been created, but the saved Client ID was not changed. Check consent/permissions and validate again using the displayed ID.')
            return $false
        }
        $script:Settings.ClientId = $candidate
        $script:Settings.AppRegistrationName = $appName
        Save-Settings
        Write-Status Ok (Get-Text 'Aplicación validada y guardada.' 'Application validated and saved.')
        return $true
    }
    catch { Write-Status Error (Get-Text "Falló el registro o validación: $($_.Exception.Message)" "Registration or validation failed: $($_.Exception.Message)"); return $false }
}
