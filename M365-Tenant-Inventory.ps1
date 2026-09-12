#requires -Version 7.4
<#
.SYNOPSIS
    M365 Tenant Inventory

.DESCRIPTION
    Inventario de sitios SharePoint Online y OneDrive for Business mediante
    Get-PnPTenantSite. La herramienta es de solo lectura: no modifica sitios,
    permisos, archivos, papelera ni configuraciones del tenant.

    Incluye:
      - TUI bilingüe español/inglés.
      - Configuración persistente y contexto de trabajo.
      - Validación de PowerShell, PnP.PowerShell, tenant, Client ID y reportes.
      - Inventario detallado con filtros de SharePoint, OneDrive y sitios de sistema.
      - Reintentos para throttling y errores transitorios.
      - Reporte CSV y resumen de ejecución.

.REQUIREMENTS
    PowerShell 7.4+
    PnP.PowerShell 3.2+
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$script:AppName = 'M365 Tenant Inventory'
$script:AppVersion = '1.0.0'
$script:MinimumPnPVersion = [version]'3.2.0'
$script:DefaultPageSize = 2000
$script:DefaultMaxRetries = 5
$script:DefaultRetryDelay = 3
$script:Language = 'es'
$script:Settings = $null
$script:LastInventory = @()
$script:LastReport = ''
$script:AdminConnection = $null
$script:Ansi = $false

$toolHome = if (-not [string]::IsNullOrWhiteSpace($env:M365_PNP_TOOLKIT_HOME)) {
    $env:M365_PNP_TOOLKIT_HOME
}
else {
    Join-Path $HOME '.m365-tenant-inventory'
}

$script:ConfigRoot = $toolHome
$script:LegacyConfigPath = Join-Path $script:ConfigRoot 'settings.json'
$script:ConfigPath = if ([string]::IsNullOrWhiteSpace($env:M365_PNP_TOOLKIT_HOME)) {
    $script:LegacyConfigPath
}
else {
    Join-Path $script:ConfigRoot (($script:AppName -replace ' ', '-') + '.settings.json')
}
$script:DefaultReportFolder = Join-Path $script:ConfigRoot 'reports'

# BEGIN TOOL LOCALIZATION
$script:ToolStrings = @{
    en = @{
        'T0189AF383D0E36EB' = 'Every option requires a numeric Key and Label.'
        'T033C612F2B730951' = 'Folder where CSV files are saved.'
        'T034A7E198D0F964E' = 'Does not delete the Entra application.'
        'T0541F65548771B8B' = 'Installs the module for the current user.'
        'T05F4A03C76C33C0D' = 'Include SharePoint'
        'T06114834B99A977B' = 'Save and validate a GUID Client ID.'
        'T072048A32EDDCD0D' = 'Include OneDrive'
        'T075FE1325BE5680B' = 'Initial domain, for example company.onmicrosoft.com.'
        'T0AA6F91071C0E09A' = 'Could not read settings.json: {0}'
        'T0B87C78C47271034' = 'Inventory error'
        'T0BCC1B2949E0D07C' = 'Not configured'
        'T17F8A7EE7D2836C7' = 'Maximum retries'
        'T18909694363DF916' = 'Included'
        'T1A7181893CBD2CBD' = 'Application validated and saved.'
        'T1B28D4A4F8DB5557' = 'Register new Entra application'
        'T1C617A82B8E4FA08' = 'Change report folder'
        'T1C98B1CF431D40C2' = 'Could not obtain a valid PnP connection.'
        'T1D1B97E641421A50' = 'Register a new application'
        'T21684FB55A9C0C7F' = 'Entra application / Client ID'
        'T23A28531ADC058A5' = 'Could not close the test connection: {0}'
        'T24233FFEAFA5A31B' = 'Last report'
        'T252ABE94A18F569C' = 'Opens generated CSV files.'
        'T286BA6B6A8316B4B' = 'Start the inventory?'
        'T2AE3BC2415EDD9E9' = 'Settings'
        'T2CF03D19A06FCB0A' = 'Closes the tool.'
        'T316F477DF9D03327' = 'No'
        'T333A93E8B7B8141E' = '{0} record(s) have no owner in the returned data.'
        'T3398EA4E730E335A' = 'SharePoint sites'
        'T34DAFAC7B2AEA842' = 'Current context'
        'T34EC709E7C49F79C' = 'Working context cleared.'
        'T38A305F6D93FCE98' = 'Value cannot be empty.'
        'T40AF8B61CFB012E6' = 'Toggle login persistence'
        'T45CC585A13EBF2B6' = 'Press Enter to continue'
        'T46FD0615D558D339' = 'Lists SharePoint and OneDrive sites and creates a CSV.'
        'T4AD0822C94B5C26F' = 'Configure or validate the application used by PnP.'
        'T4E41BE171E54180F' = 'Use the initial domain: name.onmicrosoft.com.'
        'T4F5665CECA963F1B' = 'Could not determine the admin center URL.'
        'T51E3D41527BA8048' = 'OneDrive'
        'T52E93CC5662EBAFE' = 'Site query'
        'T534828BA9DE92BC0' = 'Exclude system sites'
        'T54B061A7A1874670' = 'Exit'
        'T563959F6AD9ABA7C' = 'Remove local Client ID'
        'T5ACFBF3FDB36E3FC' = 'Run tenant inventory'
        'T5C4361A89124A502' = 'Records in memory'
        'T5EF13EC70DB15A81' = 'ACTIVE'
        'T5F1FB498D5AC98B9' = 'Yes'
        'T5FADEDA3BED70263' = 'Application validation failed: {0}'
        'T600AA05D1DE02A5A' = 'Unavailable'
        'T60ADA32FD4A834C8' = '{0} site(s) use 80% or more of the known quota.'
        'T6107CBE19582BE77' = 'Writable reports'
        'T619CE4E753307E04' = 'Configured'
        'T659E54FF1E1C6B20' = 'Inventory completed in read-only mode.'
        'T6BF640A36D04A9D7' = 'Environment'
        'T6BFD73A8B70AC9CD' = 'Entra application / PnP authentication'
        'T6D9277282E99B0F1' = 'Total inventoried'
        'T6D986E2E69A4F27C' = 'Summary'
        'T6DF39E753D15E204' = 'Tests authentication and admin center access.'
        'T6DFA1CCB7C205119' = 'Report'
        'T6F71913E9F915127' = 'Existing Client ID (Enter with no value to cancel)'
        'T70C8492F6D720D66' = 'No sites were found with the current filters.'
        'T719405B573B93C68' = 'Reports'
        'T7225AFDB6B9A872E' = 'Validate configured application'
        'T744B21B35A8DC099' = 'Could not open the reports folder.'
        'T790FBA20E4A85B79' = 'The report folder is not writable: {0}'
        'T7C8980091972774B' = 'The application may have been created, but the saved Client ID was not changed. Check consent/permissions and validate again using the displayed ID.'
        'T7D931A4CAEB7EDA4' = 'SharePoint'
        'T7DB06AAD1C19F5E7' = 'Remove the saved Client ID?'
        'T7DE227CE7EF59180' = 'Opens the CSV location.'
        'T80892645BA2B4B9F' = 'Excluded'
        'T80FF714B30CFF53F' = 'Home'
        'T876978C64A8020C3' = 'A valid GUID is required.'
        'T8A86B2B2DEF9119B' = '{0} retries · initial wait {1}s.'
        'T8C0C562527CD36DC' = 'The Client ID is not a valid GUID.'
        'T8EF3881D450B7B3A' = 'SharePoint admin center URL (check if the domain was renamed)'
        'T93C24282273F635C' = 'Status'
        'T93E42BF8804F2A67' = 'Current: {1}.'
        'T940842D8DF8C5890' = 'Open reports folder'
        'T9BCDC5AEC758501E' = 'Reports folder opened.'
        'T9E04B92C3940BFCF' = 'Change tenant'
        'TA34EB2DA62F0B793' = 'System sites'
        'TA3DEEDEE87C3E82D' = 'Configure retries'
        'TA453DE1959978E60' = '{0}: retry {2}/{1} in {3}s.'
        'TA8A2E919238F2284' = 'INACTIVE'
        'TABC142D4249FBEF1' = 'Registration or validation failed: {0}'
        'TADCF6F0FB851F13A' = 'Inventory result'
        'TAE9E2DADFCE21C09' = 'Register this application?'
        'TAF67DD963947C57D' = 'Findings and conclusion'
        'TB01DF30BD2D65CCA' = 'Admin session'
        'TB1ED45769A1B1C85' = 'PnP Application'
        'TB66DAB38B644329B' = 'Excluded'
        'TB7E83AF80BAA9767' = 'Use an existing application'
        'TBA0150205AD30BE3' = 'Register a new PnP application'
        'TBBF03A4A7AC3A41E' = 'Diagnostics'
        'TBC0C62F4657BE857' = 'A valid nonzero GUID is required.'
        'TBC64A680FC28821E' = 'Application and access validated.'
        'TBED8661029594CA8' = 'Tenant, application, filters, retries, and reports.'
        'TC00AC5CED944A045' = 'Created Application (client) ID (Enter to cancel)'
        'TC3EC5F1EFDC1CD2C' = 'Unexpected error'
        'TC5ED5CC66EA25BD5' = 'Reports folder'
        'TCAD7A53B85CACF66' = 'Language'
        'TCAE54F3552C0E84E' = 'Context'
        'TCC739D07A0F3DE6C' = 'Known quota (GB)'
        'TCE0DDB0921C76A21' = 'Aggregate usage'
        'TCF3F99F70AA7359C' = 'The reports folder is unavailable.'
        'TD0452D8BA331528E' = 'Clears results from this run only.'
        'TD2E63ADBD1B3CD73' = 'Last inventory'
        'TD3CD73DA633DE7F5' = 'Enter a number between {0} and {1}.'
        'TD47590581B3251EE' = 'Install / update PnP.PowerShell'
        'TD9624157F8433648' = 'No sites were found above 80% of known quota.'
        'TD966A6A2C7F0E31B' = 'Application name'
        'TDA312E9CFDF6B65B' = 'Enter the HTTPS SharePoint admin center URL.'
        'TDFE01DEF28079DC9' = 'Included'
        'TE0955ABC442A2E77' = 'PowerShell {0} detected; 7.4+ is required.'
        'TE18D2B1EC139BC91' = 'Querying tenant sites...'
        'TE1CAA9DC71DAAF93' = 'Creates and saves an Entra app for this toolkit.'
        'TE3D1042D5CB934D8' = '{0} locations inventoried: {1} SharePoint and {2} OneDrive.'
        'TE68D0307D45F4DD6' = 'Quota is available for {1} of {0} records; review the CSV for missing values.'
        'TE81D97B25436053E' = 'Not installed'
        'TE8C25CA0BEF17D40' = 'Spanish / English.'
        'TE8C49C15FBB66948' = 'Use the initial Entra domain ending in .onmicrosoft.com.'
        'TEBD48D58A9F26D21' = 'Validating authentication and access to the tenant inventory...'
        'TEC085F1061A9B82C' = 'Known usage (GB)'
        'TEC625AA49CFFCBE9' = 'Creates an Entra application. Requests delegated SharePoint AllSites.FullControl and delegated Graph User.Read. Requires authorization to register applications and administrator consent. Does not change sites; creates or saves no secrets.'
        'TF0356D45BEBDA171' = 'Run inventory'
        'TF0AEE53D891CBAAB' = 'Client ID removed from local settings.'
        'TF33260139CCAFC6D' = 'Clear context'
        'TF43296E41AEECCBA' = 'Client ID validated and saved.'
        'TF66DE509F1AD5A86' = 'Largest consumers:'
        'TF7B621FA5B0491DC' = 'This operation only reads tenant information and creates a CSV.'
        'TF8F9BE4BC25501AF' = 'Saved app'
        'TF915851F215548DC' = 'Tenant'
        'TFB5F3E7091BB532E' = 'Change language'
        'TFC874B1E571463DF' = 'Initial tenant (Enter to cancel)'
        'TFE86491BB28F90EB' = 'Initial delay in seconds'
        'TFF17FC308F5462DB' = 'Validates PowerShell, PnP, configuration, and reports folder.'
    }
    es = @{
        'T0189AF383D0E36EB' = 'Cada opción requiere Key numérico y Label.'
        'T033C612F2B730951' = 'Ruta donde se guardan los CSV.'
        'T034A7E198D0F964E' = 'No elimina la aplicación de Entra.'
        'T0541F65548771B8B' = 'Instala el módulo para el usuario actual.'
        'T05F4A03C76C33C0D' = 'Incluir SharePoint'
        'T06114834B99A977B' = 'Guarda y valida un Client ID GUID.'
        'T072048A32EDDCD0D' = 'Incluir OneDrive'
        'T075FE1325BE5680B' = 'Dominio inicial, por ejemplo empresa.onmicrosoft.com.'
        'T0AA6F91071C0E09A' = 'No se pudo leer settings.json: {0}'
        'T0B87C78C47271034' = 'Error de inventario'
        'T0BCC1B2949E0D07C' = 'No configurado'
        'T17F8A7EE7D2836C7' = 'Máximo de reintentos'
        'T18909694363DF916' = 'Incluido'
        'T1A7181893CBD2CBD' = 'Aplicación validada y guardada.'
        'T1B28D4A4F8DB5557' = 'Registrar nueva aplicación Entra'
        'T1C617A82B8E4FA08' = 'Cambiar carpeta de reportes'
        'T1C98B1CF431D40C2' = 'No se pudo obtener una conexión PnP válida.'
        'T1D1B97E641421A50' = 'Registrar una nueva aplicación'
        'T21684FB55A9C0C7F' = 'Aplicación Entra / Client ID'
        'T23A28531ADC058A5' = 'No se pudo cerrar la conexión de prueba: {0}'
        'T24233FFEAFA5A31B' = 'Último reporte'
        'T252ABE94A18F569C' = 'Abre los CSV generados.'
        'T286BA6B6A8316B4B' = '¿Iniciar el inventario?'
        'T2AE3BC2415EDD9E9' = 'Configuración'
        'T2CF03D19A06FCB0A' = 'Cierra la herramienta.'
        'T316F477DF9D03327' = 'No'
        'T333A93E8B7B8141E' = '{0} registro(s) no incluyen propietario en los datos devueltos.'
        'T3398EA4E730E335A' = 'Sitios SharePoint'
        'T34DAFAC7B2AEA842' = 'Contexto actual'
        'T34EC709E7C49F79C' = 'Contexto de trabajo limpiado.'
        'T38A305F6D93FCE98' = 'El valor no puede estar vacío.'
        'T40AF8B61CFB012E6' = 'Alternar persistencia de login'
        'T45CC585A13EBF2B6' = 'Presiona Enter para continuar'
        'T46FD0615D558D339' = 'Lista sitios SharePoint y OneDrive y genera un CSV.'
        'T4AD0822C94B5C26F' = 'Configura o valida la aplicación usada por PnP.'
        'T4E41BE171E54180F' = 'Usa el dominio inicial: nombre.onmicrosoft.com.'
        'T4F5665CECA963F1B' = 'No se pudo determinar la URL del centro de administración.'
        'T51E3D41527BA8048' = 'OneDrive'
        'T52E93CC5662EBAFE' = 'Consulta de sitios'
        'T534828BA9DE92BC0' = 'Excluir sitios de sistema'
        'T54B061A7A1874670' = 'Salir'
        'T563959F6AD9ABA7C' = 'Quitar Client ID local'
        'T5ACFBF3FDB36E3FC' = 'Ejecutar inventario del tenant'
        'T5C4361A89124A502' = 'Registros en memoria'
        'T5EF13EC70DB15A81' = 'ACTIVA'
        'T5F1FB498D5AC98B9' = 'Sí'
        'T5FADEDA3BED70263' = 'Falló la validación de la aplicación: {0}'
        'T600AA05D1DE02A5A' = 'No disponible'
        'T60ADA32FD4A834C8' = '{0} sitio(s) usan 80% o más de la cuota conocida.'
        'T6107CBE19582BE77' = 'Reportes escribibles'
        'T619CE4E753307E04' = 'Configurado'
        'T659E54FF1E1C6B20' = 'Inventario completado en modo solo lectura.'
        'T6BF640A36D04A9D7' = 'Entorno'
        'T6BFD73A8B70AC9CD' = 'Aplicación Entra / autenticación PnP'
        'T6D9277282E99B0F1' = 'Total inventariado'
        'T6D986E2E69A4F27C' = 'Resumen'
        'T6DF39E753D15E204' = 'Prueba autenticación y acceso al centro de administración.'
        'T6DFA1CCB7C205119' = 'Reporte'
        'T6F71913E9F915127' = 'Client ID existente (Enter sin valor para cancelar)'
        'T70C8492F6D720D66' = 'No se encontraron sitios con los filtros actuales.'
        'T719405B573B93C68' = 'Reportes'
        'T7225AFDB6B9A872E' = 'Validar aplicación configurada'
        'T744B21B35A8DC099' = 'No se pudo abrir la carpeta de reportes.'
        'T790FBA20E4A85B79' = 'La carpeta de reportes no es escribible: {0}'
        'T7C8980091972774B' = 'La aplicación puede haberse creado, pero no se cambió el Client ID guardado. Revisa consentimiento/permisos y vuelve a validar usando el ID mostrado.'
        'T7D931A4CAEB7EDA4' = 'SharePoint'
        'T7DB06AAD1C19F5E7' = '¿Quitar el Client ID guardado?'
        'T7DE227CE7EF59180' = 'Abre la ubicación de los CSV.'
        'T80892645BA2B4B9F' = 'Excluido'
        'T80FF714B30CFF53F' = 'Inicio'
        'T876978C64A8020C3' = 'Debe ser un GUID válido.'
        'T8A86B2B2DEF9119B' = '{0} reintentos · espera inicial {1}s.'
        'T8C0C562527CD36DC' = 'El Client ID no es un GUID válido.'
        'T8EF3881D450B7B3A' = 'URL del centro de administración SharePoint (confirma si el dominio fue renombrado)'
        'T93C24282273F635C' = 'Estado'
        'T93E42BF8804F2A67' = 'Actual: {0}.'
        'T940842D8DF8C5890' = 'Abrir carpeta de reportes'
        'T9BCDC5AEC758501E' = 'Carpeta de reportes abierta.'
        'T9E04B92C3940BFCF' = 'Cambiar tenant'
        'TA34EB2DA62F0B793' = 'Sitios sistema'
        'TA3DEEDEE87C3E82D' = 'Configurar reintentos'
        'TA453DE1959978E60' = '{0}: reintento {2}/{1} en {3}s.'
        'TA8A2E919238F2284' = 'INACTIVA'
        'TABC142D4249FBEF1' = 'Falló el registro o validación: {0}'
        'TADCF6F0FB851F13A' = 'Resultado del inventario'
        'TAE9E2DADFCE21C09' = '¿Registrar esta aplicación?'
        'TAF67DD963947C57D' = 'Hallazgos y conclusión'
        'TB01DF30BD2D65CCA' = 'Sesión admin'
        'TB1ED45769A1B1C85' = 'Aplicación PnP'
        'TB66DAB38B644329B' = 'Excluidos'
        'TB7E83AF80BAA9767' = 'Usar una aplicación existente'
        'TBA0150205AD30BE3' = 'Registrar una nueva aplicación PnP'
        'TBBF03A4A7AC3A41E' = 'Diagnóstico'
        'TBC0C62F4657BE857' = 'Debe ser un GUID válido distinto de cero.'
        'TBC64A680FC28821E' = 'Aplicación y acceso validados.'
        'TBED8661029594CA8' = 'Tenant, aplicación, filtros, reintentos y reportes.'
        'TC00AC5CED944A045' = 'Application (client) ID creado (Enter para cancelar)'
        'TC3EC5F1EFDC1CD2C' = 'Error inesperado'
        'TC5ED5CC66EA25BD5' = 'Carpeta de reportes'
        'TCAD7A53B85CACF66' = 'Idioma'
        'TCAE54F3552C0E84E' = 'Contexto'
        'TCC739D07A0F3DE6C' = 'Cuota conocida (GB)'
        'TCE0DDB0921C76A21' = 'Uso agregado'
        'TCF3F99F70AA7359C' = 'La carpeta de reportes no está disponible.'
        'TD0452D8BA331528E' = 'Elimina resultados solo de la memoria de esta ejecución.'
        'TD2E63ADBD1B3CD73' = 'Último inventario'
        'TD3CD73DA633DE7F5' = 'Introduce un número entre {0} y {1}.'
        'TD47590581B3251EE' = 'Instalar / actualizar PnP.PowerShell'
        'TD9624157F8433648' = 'No se detectaron sitios por encima del 80% de la cuota conocida.'
        'TD966A6A2C7F0E31B' = 'Nombre de la aplicación'
        'TDA312E9CFDF6B65B' = 'Introduce la URL HTTPS del centro de administración SharePoint.'
        'TDFE01DEF28079DC9' = 'Incluidos'
        'TE0955ABC442A2E77' = 'PowerShell {0} detectado; se requiere 7.4+.'
        'TE18D2B1EC139BC91' = 'Consultando sitios del tenant...'
        'TE1CAA9DC71DAAF93' = 'Crea y guarda una app Entra para este toolkit.'
        'TE3D1042D5CB934D8' = 'Se inventariaron {0} ubicaciones: {1} SharePoint y {2} OneDrive.'
        'TE68D0307D45F4DD6' = 'La cuota está disponible para {1} de {0} registros; revisa el CSV para los faltantes.'
        'TE81D97B25436053E' = 'No instalado'
        'TE8C25CA0BEF17D40' = 'Español / English.'
        'TE8C49C15FBB66948' = 'Usa el dominio inicial de Entra terminado en .onmicrosoft.com.'
        'TEBD48D58A9F26D21' = 'Validando autenticación y acceso al inventario del tenant...'
        'TEC085F1061A9B82C' = 'Uso conocido (GB)'
        'TEC625AA49CFFCBE9' = 'Creará una aplicación Entra. Solicitará AllSites.FullControl delegado de SharePoint y User.Read delegado de Graph. Requiere autorización para registrar aplicaciones y consentimiento administrativo. No cambia sitios; no crea ni guarda secretos.'
        'TF0356D45BEBDA171' = 'Ejecutar inventario'
        'TF0AEE53D891CBAAB' = 'Client ID eliminado de la configuración local.'
        'TF33260139CCAFC6D' = 'Limpiar contexto'
        'TF43296E41AEECCBA' = 'Client ID validado y guardado.'
        'TF66DE509F1AD5A86' = 'Mayores consumos:'
        'TF7B621FA5B0491DC' = 'Esta operación solo consulta información del tenant y genera un CSV.'
        'TF8F9BE4BC25501AF' = 'App guardada'
        'TF915851F215548DC' = 'Tenant'
        'TFB5F3E7091BB532E' = 'Cambiar idioma'
        'TFC874B1E571463DF' = 'Tenant inicial (Enter para cancelar)'
        'TFE86491BB28F90EB' = 'Espera inicial en segundos'
        'TFF17FC308F5462DB' = 'Valida PowerShell, PnP, configuración y carpeta de reportes.'
    }
}
# END TOOL LOCALIZATION
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
function Get-Text {
    param(
        [Parameter(Mandatory)][string]$Spanish,
        [Parameter(Mandatory)][string]$English
    )

    if ($script:Language -eq 'en') { return $English }
    return $Spanish
}

function Initialize-AppLanguage {
    param([string]$Default = 'en')
    Initialize-ToolkitLanguage -Default $Default -Name $script:AppName -Version $script:AppVersion
}

function Initialize-Terminal {
    try {
        $script:Ansi = -not [Console]::IsOutputRedirected
    }
    catch {
        $script:Ansi = $false
    }
}

function Clear-AppScreen {
    try { if (-not [Console]::IsOutputRedirected) { Clear-Host } } catch { }
}

function Get-AnsiCode {
    param([ValidateSet('Primary','Success','Info','Warning','Danger','Muted','Accent')][string]$Style)

    switch ($Style) {
        'Primary' { return "`e[96m" }
        'Success' { return "`e[92m" }
        'Info'    { return "`e[94m" }
        'Warning' { return "`e[93m" }
        'Danger'  { return "`e[91m" }
        'Accent'  { return "`e[95m" }
        default   { return "`e[90m" }
    }
}

function Write-Styled {
    param(
        [AllowEmptyString()][string]$Text,
        [ValidateSet('Primary','Success','Info','Warning','Danger','Muted','Accent')][string]$Style = 'Primary',
        [switch]$NoNewline
    )
    Write-ToolkitText $Text $Style -NoNewline:$NoNewline
}

function Write-AppHeader {
    param([Parameter(Mandatory)][string]$Section)
    Write-ToolkitHeader $script:AppName $script:AppVersion $Section
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    Write-ToolkitText ''
    Write-ToolkitText $Title Primary
}

function Write-Field {
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$Value,
        [ValidateSet('Primary','Success','Info','Warning','Danger','Muted','Accent')][string]$Style = 'Primary'
    )
    Write-ToolkitField $Name $Value $Style
}

function Write-Status {
    param(
        [Parameter(Mandatory)][ValidateSet('Ok','Info','Warn','Error')][string]$Kind,
        [Parameter(Mandatory)][string]$Message
    )
    Write-ToolkitStatus $Kind $Message
}

function Pause-Tui {
    [void](Read-ToolkitInput ('  ' + (Get-ToolkitString Continue)))
}

function Read-MenuChoice {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Items, [switch]$AllowBack, [scriptblock]$RenderBody)
    $normalized = @(
        foreach ($item in $Items) {
            $key = [string](Get-PropertyValue $item Key '')
            $label = [string](Get-PropertyValue $item Label '')
            if ($key -notmatch '^\d+$' -or -not $label) { throw (Get-LocalizedString -Key 'T0189AF383D0E36EB') }
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

function Read-TextValue {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = '',
        [switch]$AllowEmpty,
        [scriptblock]$Validator,
        [string]$ValidationMessage = 'Valor inválido.'
    )

    while ($true) {
        $caption = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
        $value = (Read-Host $caption).Trim()
        if ([string]::IsNullOrWhiteSpace($value) -and $Default) { $value = $Default }
        if ([string]::IsNullOrWhiteSpace($value)) {
            if ($AllowEmpty) { return '' }
            Write-Status Error (Get-LocalizedString -Key 'T38A305F6D93FCE98')
            continue
        }
        if ($Validator -and -not (& $Validator $value)) {
            Write-Status Error $ValidationMessage
            continue
        }
        return $value
    }
}

function Read-YesNo {
    param([Parameter(Mandatory)][string]$Prompt, [bool]$Default = $false)
    return Read-ToolkitYesNo $Prompt $Default
}

function Read-Integer {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [int]$Default,
        [int]$Minimum = [int]::MinValue,
        [int]$Maximum = [int]::MaxValue
    )

    while ($true) {
        $raw = Read-Host "$Prompt [$Default]"
        if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
        $number = 0
        if ([int]::TryParse($raw, [ref]$number) -and $number -ge $Minimum -and $number -le $Maximum) {
            return $number
        }
        Write-Status Error (Get-LocalizedString -Key 'TD3CD73DA633DE7F5' -Values @($Minimum, $Maximum))
    }
}

function New-DefaultSettings {
    $clientId = ''
    if ($env:ENTRAID_APP_ID) { $clientId = $env:ENTRAID_APP_ID }
    elseif ($env:ENTRAID_CLIENT_ID) { $clientId = $env:ENTRAID_CLIENT_ID }

    [pscustomobject]@{
        Tenant = ''
        ClientId = $clientId
        AppRegistrationName = $script:AppName
        ReportFolder = $script:DefaultReportFolder
        PersistLogin = $true
        Language = 'es'
        IncludeSharePoint = $true
        IncludeOneDrive = $true
        ExcludeSystemSites = $true
        PageSize = $script:DefaultPageSize
        MaxRetries = $script:DefaultMaxRetries
        RetryDelay = $script:DefaultRetryDelay
    }
}

function Merge-Settings {
    param($Loaded)

    $defaults = New-DefaultSettings
    if ($null -eq $Loaded) { return $defaults }

    foreach ($property in $defaults.PSObject.Properties.Name) {
        if ($null -eq $Loaded.PSObject.Properties[$property]) {
            Add-Member -InputObject $Loaded -NotePropertyName $property -NotePropertyValue $defaults.$property
        }
    }
    return $Loaded
}

function Load-Settings {
    $readPath = if (Test-Path -LiteralPath $script:ConfigPath) { $script:ConfigPath } else { $script:LegacyConfigPath }

    if (-not (Test-Path -LiteralPath $readPath)) { return New-DefaultSettings }
    try {
        $loaded = Get-Content -LiteralPath $readPath -Raw -Encoding UTF8 | ConvertFrom-Json
        return Merge-Settings $loaded
    }
    catch {
        Write-Status Warn (Get-LocalizedString -Key 'T0AA6F91071C0E09A' -Values @($($_.Exception.Message)))
        return New-DefaultSettings
    }

}

function Save-Settings {
    if (-not (Test-Path -LiteralPath $script:ConfigRoot)) {
        [void](New-Item -ItemType Directory -Path $script:ConfigRoot -Force)
    }
    $script:Settings | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}

function Ensure-ReportFolder {
    if ([string]::IsNullOrWhiteSpace($script:Settings.ReportFolder)) {
        $script:Settings.ReportFolder = $script:DefaultReportFolder
    }
    if (-not (Test-Path -LiteralPath $script:Settings.ReportFolder)) {
        [void](New-Item -ItemType Directory -Path $script:Settings.ReportFolder -Force)
    }

    $testFile = Join-Path $script:Settings.ReportFolder ".write-test-$([guid]::NewGuid().ToString('N')).tmp"
    try {
        Set-Content -LiteralPath $testFile -Value 'test' -Encoding UTF8
        Remove-Item -LiteralPath $testFile -Force
        return $true
    }
    catch {
        Write-Status Error (Get-LocalizedString -Key 'T790FBA20E4A85B79' -Values @($($_.Exception.Message)))
        return $false
    }
}

function Open-ReportsFolder {
    try {
        if (-not (Ensure-ReportFolder)) { Pause-Tui; return }
        $folder = [System.IO.Path]::GetFullPath($script:Settings.ReportFolder)
        if ($IsWindows) {
            Start-Process -FilePath 'explorer.exe' -ArgumentList @($folder) | Out-Null
        }
        elseif (Get-Command xdg-open -ErrorAction SilentlyContinue) {
            Start-Process -FilePath 'xdg-open' -ArgumentList @($folder) | Out-Null
        }
        else { Invoke-Item -LiteralPath $folder }
        Write-Status Ok (Get-LocalizedString -Key 'T9BCDC5AEC758501E')
    }
    catch {
        Write-Status Error (Get-LocalizedString -Key 'T744B21B35A8DC099')
    }
    Pause-Tui
}

function Test-PowerShellVersion {
    if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
        Write-Status Error (Get-LocalizedString -Key 'TE0955ABC442A2E77' -Values @($($PSVersionTable.PSVersion)))
        return $false
    }
    return $true
}

function Get-InstalledPnPModule {
    Get-Module -Name PnP.PowerShell -ListAvailable |
        Sort-Object Version -Descending |
        Select-Object -First 1
}

function Install-PnPModule {
    return Initialize-ToolkitPnP -MinimumVersion '3.2.0' -Update
}

function Ensure-PnPModule {
    param([switch]$OfferInstall)
    return Initialize-ToolkitPnP -MinimumVersion '3.2.0'
}

function Test-TenantFormat {
    param([AllowNull()][string]$Tenant)
    return (-not [string]::IsNullOrWhiteSpace($Tenant)) -and ($Tenant.Trim() -match '^[A-Za-z0-9-]+\.onmicrosoft\.com$')
}

function Test-ClientIdFormat {
    param([AllowNull()][string]$ClientId)
    return Test-ToolkitClientId $ClientId
}

function Get-TenantPrefix {
    if (-not (Test-TenantFormat $script:Settings.Tenant)) { return '' }
    return ($script:Settings.Tenant -replace '(?i)\.onmicrosoft\.com$', '')
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
        $tenant = Read-TextValue -Prompt (Get-LocalizedString -Key 'TFC874B1E571463DF') -AllowEmpty -Validator { param($v) Test-TenantFormat $v } -ValidationMessage (Get-LocalizedString -Key 'TE8C49C15FBB66948')
        if (-not $tenant) { return $false }
        $script:Settings.Tenant = $tenant.ToLowerInvariant()
    }
    if ([string](Get-PropertyValue $script:Settings AdminUrlTenant '') -ne $script:Settings.Tenant) {
        $url = Read-TextValue -Prompt (Get-LocalizedString -Key 'T8EF3881D450B7B3A') -Default (Get-AdminUrl) -Validator { param($v) (Get-ToolkitSiteEndpoint $v) -eq 'Admin' -and ([uri]$v).AbsolutePath -eq '/' -and -not ([uri]$v).Query -and -not ([uri]$v).Fragment } -ValidationMessage (Get-LocalizedString -Key 'TDA312E9CFDF6B65B')
        $script:Settings | Add-Member -NotePropertyName AdminUrl -NotePropertyValue $url.TrimEnd('/') -Force
        $script:Settings | Add-Member -NotePropertyName AdminUrlTenant -NotePropertyValue $script:Settings.Tenant -Force
        Save-Settings
    }
    return $true
}

function Ensure-ClientId {
    if (Test-ClientIdFormat $script:Settings.ClientId) { return $true }
    Write-Section (Get-LocalizedString -Key 'TB1ED45769A1B1C85')
    $choice = Read-MenuChoice -Items @(
        [pscustomobject]@{Key='1';Label=(Get-LocalizedString -Key 'TB7E83AF80BAA9767')},
        [pscustomobject]@{Key='2';Label=(Get-LocalizedString -Key 'T1D1B97E641421A50')},
        [pscustomobject]@{Key='0';Label=(Get-ToolkitString Cancel)}
    )
    switch ($choice) {
        '1' { return Set-ExistingPnPClientId }
        '2' { return Register-NewPnPApp }
        default { return $false }
    }
}

function Connect-M365AdminWithClientId {
    param([Parameter(Mandatory)][string]$ClientId)

    if (-not (Ensure-PnPModule)) { return $null }
    if (-not (Ensure-TenantConfigured)) { return $null }
    if (-not (Test-ClientIdFormat $ClientId)) { throw (Get-ToolkitString Invalid) }

    $adminUrl = Get-AdminUrl
    if ([string]::IsNullOrWhiteSpace($adminUrl)) { throw (Get-LocalizedString -Key 'T4F5665CECA963F1B') }

    $params = @{
        Url = $adminUrl
        ClientId = $ClientId
        Tenant = [string]$script:Settings.Tenant
        Interactive = $true
        ReturnConnection = $true
        ErrorAction = 'Stop'
    }
    if ($script:Settings.PersistLogin) { $params.PersistLogin = $true }
    return Connect-PnPOnline @params
}

function Connect-M365Admin {
    if (-not (Ensure-PnPModule)) { return $null }
    if (-not (Ensure-TenantConfigured)) { return $null }
    if (-not (Ensure-ClientId)) { return $null }

    $adminUrl = Get-AdminUrl
    if ([string]::IsNullOrWhiteSpace($adminUrl)) { throw (Get-LocalizedString -Key 'T4F5665CECA963F1B') }

    $params = @{
        Url = $adminUrl
        ClientId = [string]$script:Settings.ClientId
        Tenant = [string]$script:Settings.Tenant
        Interactive = $true
        ReturnConnection = $true
        ErrorAction = 'Stop'
    }
    if ($script:Settings.PersistLogin) { $params.PersistLogin = $true }
    return Connect-PnPOnline @params
}

function Test-PnPAppRegistration {
    param([Parameter(Mandatory)][string]$ClientId, [switch]$Quiet)
    if (-not (Test-ClientIdFormat $ClientId)) {
        if (-not $Quiet) { Write-Status Error (Get-LocalizedString -Key 'T8C0C562527CD36DC') }
        return $false
    }
    $connection = $null
    try {
        if (-not $Quiet) { Write-Status Info (Get-LocalizedString -Key 'TEBD48D58A9F26D21') }
        $connection = Connect-M365AdminWithClientId -ClientId $ClientId
        if ($null -eq $connection) { return $false }
        $null = Get-PnPTenant -Connection $connection -ErrorAction Stop
        $null = Get-PnPTenantSite -Connection $connection -ErrorAction Stop | Select-Object -First 1
        if (-not $Quiet) { Write-Status Ok (Get-LocalizedString -Key 'TBC64A680FC28821E') }
        return $true
    }
    catch {
        if (-not $Quiet) { Write-Status Error (Get-LocalizedString -Key 'T5FADEDA3BED70263' -Values @($($_.Exception.Message))) }
        return $false
    }
    finally {
        if ($null -ne $connection) {
            try { Disconnect-PnPOnline -Connection $connection -ErrorAction Stop }
            catch { if (-not $Quiet) { Write-Status Warn (Get-LocalizedString -Key 'T23A28531ADC058A5' -Values @($($_.Exception.Message))) } }
        }
    }
}

function Set-ExistingPnPClientId {
    if (-not (Ensure-TenantConfigured)) { return $false }
    $current = if (Test-ClientIdFormat $script:Settings.ClientId) { $script:Settings.ClientId } else { '' }
    $clientId = Read-TextValue -Prompt (Get-LocalizedString -Key 'T6F71913E9F915127') -Default $current -AllowEmpty -Validator { param($v) Test-ClientIdFormat $v } -ValidationMessage (Get-LocalizedString -Key 'TBC0C62F4657BE857')
    if (-not $clientId) { return $false }
    if (-not (Test-PnPAppRegistration -ClientId $clientId)) { return $false }
    $script:Settings.ClientId = $clientId
    Save-Settings
    Write-Status Ok (Get-LocalizedString -Key 'TF43296E41AEECCBA')
    return $true
}

function Register-NewPnPApp {
    if (-not (Ensure-PnPModule -OfferInstall)) { return $false }
    if (-not (Ensure-TenantConfigured)) { return $false }
    Write-Section (Get-LocalizedString -Key 'T1B28D4A4F8DB5557')
    $appName = Read-TextValue -Prompt (Get-LocalizedString -Key 'TD966A6A2C7F0E31B') -Default $script:AppName
    Write-Styled (Get-LocalizedString -Key 'TEC625AA49CFFCBE9') Muted
    if (-not (Read-YesNo (Get-LocalizedString -Key 'TAE9E2DADFCE21C09'))) { return $false }
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
            $candidate = Read-TextValue -Prompt (Get-LocalizedString -Key 'TC00AC5CED944A045') -AllowEmpty -Validator { param($v) Test-ClientIdFormat $v } -ValidationMessage (Get-LocalizedString -Key 'T876978C64A8020C3')
        }
        if (-not $candidate) { return $false }
        Write-Field 'Client ID' $candidate
        if (-not (Test-PnPAppRegistration -ClientId $candidate)) {
            Write-Status Warn (Get-LocalizedString -Key 'T7C8980091972774B')
            return $false
        }
        $script:Settings.ClientId = $candidate
        $script:Settings.AppRegistrationName = $appName
        Save-Settings
        Write-Status Ok (Get-LocalizedString -Key 'T1A7181893CBD2CBD')
        return $true
    }
    catch { Write-Status Error (Get-LocalizedString -Key 'TABC142D4249FBEF1' -Values @($($_.Exception.Message))); return $false }
}

function Disconnect-Admin {
    if ($null -ne $script:AdminConnection) {
        try { Disconnect-PnPOnline -Connection $script:AdminConnection -ErrorAction SilentlyContinue } catch { }
        $script:AdminConnection = $null
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$OperationName = 'operación'
    )

    $attempt = 0
    $delay = [Math]::Max(1, [int]$script:Settings.RetryDelay)
    $maxRetries = [Math]::Max(1, [int]$script:Settings.MaxRetries)

    while ($true) {
        try { return & $ScriptBlock }
        catch {
            $attempt++
            $message = $_.Exception.Message
            $retryable = $message -match '(?i)\b429\b|throttl|too many requests|\b503\b|service unavailable|temporarily unavailable|timeout|timed out'
            if (-not $retryable -or $attempt -ge $maxRetries) { throw }
            $sleep = [Math]::Min(120, $delay + (Get-Random -Minimum 0 -Maximum 3))
            Write-Status Warn (Get-LocalizedString -Key 'TA453DE1959978E60' -Values @(${OperationName}, $maxRetries, $attempt, ${sleep}))
            Start-Sleep -Seconds $sleep
            $delay = [Math]::Min(120, $delay * 2)
        }
    }
}

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

function Get-SiteKind {
    param([Parameter(Mandatory)]$Site)
    if (Test-IsOneDriveUrl ([string](Get-PropertyValue -Object $Site -Name 'Url' -Default ''))) {
        return 'OneDrive'
    }
    return 'SharePoint'
}

function Convert-MbToGb {
    param([AllowNull()]$Megabytes)
    $value = 0.0
    if ($null -eq $Megabytes -or -not [double]::TryParse(([string]$Megabytes), [ref]$value)) { return $null }
    return [math]::Round(($value / 1024), 2)
}

function Format-SiteDate {
    param([AllowNull()]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try { return ([datetime]$Value).ToString('o') } catch { return [string]$Value }
}

function Convert-SiteToInventoryRecord {
    param([Parameter(Mandatory)]$Site)

    $storageCurrentMb = Get-PropertyValue -Object $Site -Name 'StorageUsageCurrent' -Default $null
    $storageQuotaMb = Get-PropertyValue -Object $Site -Name 'StorageQuota' -Default $null
    $storagePercent = $null
    if ($null -ne $storageCurrentMb -and $null -ne $storageQuotaMb -and [double]$storageQuotaMb -gt 0) {
        $storagePercent = [math]::Round(([double]$storageCurrentMb / [double]$storageQuotaMb) * 100, 2)
    }

    $url = [string](Get-PropertyValue -Object $Site -Name 'Url' -Default '')
    $kind = Get-SiteKind -Site $Site
    [pscustomobject]@{
        SiteType = $kind
        Title = [string](Get-PropertyValue -Object $Site -Name 'Title' -Default '')
        Url = $url
        Owner = [string](Get-PropertyValue -Object $Site -Name 'Owner' -Default '')
        Template = [string](Get-PropertyValue -Object $Site -Name 'Template' -Default '')
        Status = [string](Get-PropertyValue -Object $Site -Name 'Status' -Default '')
        LockState = [string](Get-PropertyValue -Object $Site -Name 'LockState' -Default '')
        LastContentModified = Format-SiteDate (Get-PropertyValue -Object $Site -Name 'LastContentModifiedDate' -Default $null)
        StorageUsageMB = if ($null -ne $storageCurrentMb) { [math]::Round([double]$storageCurrentMb, 2) } else { $null }
        StorageUsageGB = Convert-MbToGb $storageCurrentMb
        StorageQuotaMB = if ($null -ne $storageQuotaMb) { [math]::Round([double]$storageQuotaMb, 2) } else { $null }
        StorageQuotaGB = Convert-MbToGb $storageQuotaMb
        StorageUsagePercent = $storagePercent
        IsSystemSite = [bool](Test-IsSystemSite -Site $Site)
        InventoryTimestamp = (Get-Date).ToString('o')
    }
}

function Get-TenantInventory {
    param([Parameter(Mandatory)]$Connection)

    $sites = @(Invoke-WithRetry -OperationName (Get-LocalizedString -Key 'T52E93CC5662EBAFE') -ScriptBlock {
        Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Connection $Connection -ErrorAction Stop
    })

    $records = foreach ($site in $sites) {
        $kind = Get-SiteKind -Site $site
        if ($kind -eq 'SharePoint' -and -not [bool]$script:Settings.IncludeSharePoint) { continue }
        if ($kind -eq 'OneDrive' -and -not [bool]$script:Settings.IncludeOneDrive) { continue }
        if ([bool]$script:Settings.ExcludeSystemSites -and (Test-IsSystemSite -Site $site)) { continue }
        Convert-SiteToInventoryRecord -Site $site
    }

    return @($records | Sort-Object SiteType, Title, Url)
}

function New-ReportPath {
    param([string]$Prefix)
    if (-not (Ensure-ReportFolder)) { throw (Get-LocalizedString -Key 'TCF3F99F70AA7359C') }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    return Join-Path $script:Settings.ReportFolder "$Prefix-$stamp.csv"
}

function Export-InventoryReport {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Records)

    $path = New-ReportPath -Prefix 'TenantInventory'
    $columns = @(
        'SiteType','Title','Url','Owner','Template','Status','LockState',
        'LastContentModified','StorageUsageMB','StorageUsageGB','StorageQuotaMB',
        'StorageQuotaGB','StorageUsagePercent','IsSystemSite','InventoryTimestamp'
    )

    if (@($Records).Count -gt 0) {
        $Records | Select-Object $columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Force
    }
    else {
        ($columns -join ',') | Set-Content -LiteralPath $path -Encoding UTF8
    }
    return $path
}

function Show-InventorySummary {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Records,
        [Parameter(Mandatory)][string]$ReportPath
    )

    Write-AppHeader (Get-LocalizedString -Key 'TADCF6F0FB851F13A')
    $sharePoint = @($Records | Where-Object SiteType -eq 'SharePoint')
    $oneDrive = @($Records | Where-Object SiteType -eq 'OneDrive')
    $withQuota = @($Records | Where-Object { $null -ne $_.StorageUsagePercent })
    $withoutOwner = @($Records | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.Owner) })
    $nearQuota = @($Records | Where-Object { $null -ne $_.StorageUsagePercent -and $_.StorageUsagePercent -ge 80 })
    $topStorage = @($Records | Where-Object { $null -ne $_.StorageUsageGB } | Sort-Object StorageUsageGB -Descending | Select-Object -First 5)
    $totalUsage = ($withQuota | Measure-Object -Property StorageUsageGB -Sum).Sum
    $totalQuota = ($withQuota | Measure-Object -Property StorageQuotaGB -Sum).Sum
    $percent = if ($totalQuota -gt 0) { [math]::Round(($totalUsage / $totalQuota) * 100, 2) } else { $null }

    Write-Section (Get-LocalizedString -Key 'T6D986E2E69A4F27C')
    Write-Field (Get-LocalizedString -Key 'T3398EA4E730E335A') $sharePoint.Count Primary
    Write-Field (Get-LocalizedString -Key 'T51E3D41527BA8048') $oneDrive.Count Primary
    Write-Field (Get-LocalizedString -Key 'T6D9277282E99B0F1') @($Records).Count Success
    Write-Field (Get-LocalizedString -Key 'TEC085F1061A9B82C') ([math]::Round([double]$totalUsage, 2)) Info
    Write-Field (Get-LocalizedString -Key 'TCC739D07A0F3DE6C') ([math]::Round([double]$totalQuota, 2)) Info
    Write-Field (Get-LocalizedString -Key 'TCE0DDB0921C76A21') $(if ($null -ne $percent) { "$percent%" } else { Get-LocalizedString -Key 'T600AA05D1DE02A5A' }) $(if ($percent -ge 80) { 'Warning' } else { 'Success' })

    Write-Section (Get-LocalizedString -Key 'TAF67DD963947C57D')
    if (@($Records).Count -eq 0) {
        Write-Status Warn (Get-LocalizedString -Key 'T70C8492F6D720D66')
    }
    else {
        Write-Status Info (Get-LocalizedString -Key 'TE3D1042D5CB934D8' -Values @($(@($Records).Count), $($sharePoint.Count), $($oneDrive.Count)))
        if ($nearQuota.Count -gt 0) {
            Write-Status Warn (Get-LocalizedString -Key 'T60ADA32FD4A834C8' -Values @($($nearQuota.Count)))
        }
        else {
            Write-Status Ok (Get-LocalizedString -Key 'TD9624157F8433648')
        }
        if ($withoutOwner.Count -gt 0) {
            Write-Status Warn (Get-LocalizedString -Key 'T333A93E8B7B8141E' -Values @($($withoutOwner.Count)))
        }
        if ($withQuota.Count -lt @($Records).Count) {
            Write-Status Info (Get-LocalizedString -Key 'TE68D0307D45F4DD6' -Values @($(@($Records).Count), $($withQuota.Count)))
        }
        if ($topStorage.Count -gt 0) {
            Write-Styled (Get-LocalizedString -Key 'TF66DE509F1AD5A86') Muted
            foreach ($site in $topStorage) {
                $label = if ([string]::IsNullOrWhiteSpace([string]$site.Title)) { [string]$site.Url } else { [string]$site.Title }
                Write-Styled ("  · {0} — {1:N2} GB" -f $label, $site.StorageUsageGB) Primary
            }
        }
    }

    Write-Section (Get-LocalizedString -Key 'T6DFA1CCB7C205119')
    Write-Field 'CSV' $ReportPath Muted
    Write-Status Ok (Get-LocalizedString -Key 'T659E54FF1E1C6B20')
    Pause-Tui
}

function Invoke-TenantInventory {
    if (-not (Ensure-PnPModule)) { Pause-Tui; return }
    if (-not (Ensure-TenantConfigured)) { Pause-Tui; return }
    if (-not (Ensure-ClientId)) { Pause-Tui; return }
    if (-not (Ensure-ReportFolder)) { Pause-Tui; return }

    Write-AppHeader (Get-LocalizedString -Key 'TF0356D45BEBDA171')
    Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant Primary
    Write-Field (Get-LocalizedString -Key 'T7D931A4CAEB7EDA4') $(if ($script:Settings.IncludeSharePoint) { Get-LocalizedString -Key 'T18909694363DF916' } else { Get-LocalizedString -Key 'T80892645BA2B4B9F' })
    Write-Field 'OneDrive' $(if ($script:Settings.IncludeOneDrive) { Get-LocalizedString -Key 'T18909694363DF916' } else { Get-LocalizedString -Key 'T80892645BA2B4B9F' })
    Write-Field (Get-LocalizedString -Key 'TA34EB2DA62F0B793') $(if ($script:Settings.ExcludeSystemSites) { Get-LocalizedString -Key 'TB66DAB38B644329B' } else { Get-LocalizedString -Key 'TDFE01DEF28079DC9' })
    Write-Status Info (Get-LocalizedString -Key 'TF7B621FA5B0491DC')

    if (-not (Read-YesNo -Prompt (Get-LocalizedString -Key 'T286BA6B6A8316B4B') -Default $true)) { return }

    try {
        $script:AdminConnection = Connect-M365Admin
        if ($null -eq $script:AdminConnection) { throw (Get-LocalizedString -Key 'T1C98B1CF431D40C2') }
        Write-Status Info (Get-LocalizedString -Key 'TE18D2B1EC139BC91')
        $records = @(Get-TenantInventory -Connection $script:AdminConnection)
        $script:LastInventory = $records
        $script:LastReport = Export-InventoryReport -Records $records
        Save-Settings
        Show-InventorySummary -Records $records -ReportPath $script:LastReport
    }
    catch {
        Write-AppHeader (Get-LocalizedString -Key 'T0B87C78C47271034')
        Write-Status Error $_.Exception.Message
        Pause-Tui
    }
    finally {
        Disconnect-Admin
    }
}

function Show-AppRegistrationMenu {
    while ($true) {
        Write-AppHeader (Get-LocalizedString -Key 'T6BFD73A8B70AC9CD')
        Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant
        Write-Field (Get-LocalizedString -Key 'TF8F9BE4BC25501AF') $script:Settings.AppRegistrationName
        Write-Field 'Client ID' $script:Settings.ClientId
        Write-Field (Get-LocalizedString -Key 'T93C24282273F635C') $(if (Test-ClientIdFormat $script:Settings.ClientId) { Get-LocalizedString -Key 'T619CE4E753307E04' } else { Get-LocalizedString -Key 'T0BCC1B2949E0D07C' }) $(if (Test-ClientIdFormat $script:Settings.ClientId) { 'Success' } else { 'Warning' })
        Write-Host ''

        $choice = Read-MenuChoice -AllowBack -Items @(
            @{ Key='1'; Label=(Get-LocalizedString -Key 'TD47590581B3251EE'); Description=(Get-LocalizedString -Key 'T0541F65548771B8B') },
            @{ Key='2'; Label=(Get-LocalizedString -Key 'TB7E83AF80BAA9767'); Description=(Get-LocalizedString -Key 'T06114834B99A977B') },
            @{ Key='3'; Label=(Get-LocalizedString -Key 'TBA0150205AD30BE3'); Description=(Get-LocalizedString -Key 'TE1CAA9DC71DAAF93') },
            @{ Key='4'; Label=(Get-LocalizedString -Key 'T7225AFDB6B9A872E'); Description=(Get-LocalizedString -Key 'T6DF39E753D15E204') },
            @{ Key='5'; Label=(Get-LocalizedString -Key 'T563959F6AD9ABA7C'); Description=(Get-LocalizedString -Key 'T034A7E198D0F964E') }
        )

        switch ($choice) {
            '0' { return }
            '1' { [void](Install-PnPModule); Pause-Tui }
            '2' { [void](Set-ExistingPnPClientId); Pause-Tui }
            '3' { [void](Register-NewPnPApp); Pause-Tui }
            '4' {
                [void](Test-PnPAppRegistration -ClientId ([string]$script:Settings.ClientId))
                Pause-Tui
            }
            '5' {
                if (Read-YesNo -Prompt (Get-LocalizedString -Key 'T7DB06AAD1C19F5E7') -Default $false) {
                    $script:Settings.ClientId = ''
                    Save-Settings
                    Write-Status Ok (Get-LocalizedString -Key 'TF0AEE53D891CBAAB')
                }
                Pause-Tui
            }
        }
    }
}

function Show-SettingsMenu {
    while ($true) {
        Write-AppHeader (Get-LocalizedString -Key 'T2AE3BC2415EDD9E9')
        Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant
        Write-Field 'Client ID' $script:Settings.ClientId
        Write-Field (Get-LocalizedString -Key 'T719405B573B93C68') $script:Settings.ReportFolder Muted
        Write-Field (Get-LocalizedString -Key 'TCAD7A53B85CACF66') $(if ($script:Language -eq 'en') { 'English' } else { 'Español' })
        Write-Field (Get-LocalizedString -Key 'T7D931A4CAEB7EDA4') $(if ($script:Settings.IncludeSharePoint) { Get-LocalizedString -Key 'T18909694363DF916' } else { Get-LocalizedString -Key 'T80892645BA2B4B9F' })
        Write-Field 'OneDrive' $(if ($script:Settings.IncludeOneDrive) { Get-LocalizedString -Key 'T18909694363DF916' } else { Get-LocalizedString -Key 'T80892645BA2B4B9F' })
        Write-Field (Get-LocalizedString -Key 'TA34EB2DA62F0B793') $(if ($script:Settings.ExcludeSystemSites) { Get-LocalizedString -Key 'TB66DAB38B644329B' } else { Get-LocalizedString -Key 'TDFE01DEF28079DC9' })
        Write-Host ''

        $choice = Read-MenuChoice -AllowBack -Items @(
            @{ Key='1'; Label=(Get-LocalizedString -Key 'T9E04B92C3940BFCF'); Description=(Get-LocalizedString -Key 'T075FE1325BE5680B') },
            @{ Key='2'; Label=(Get-LocalizedString -Key 'T21684FB55A9C0C7F'); Description=(Get-LocalizedString -Key 'T4AD0822C94B5C26F') },
            @{ Key='3'; Label=(Get-LocalizedString -Key 'T1C617A82B8E4FA08'); Description=(Get-LocalizedString -Key 'T033C612F2B730951') },
            @{ Key='4'; Label=(Get-LocalizedString -Key 'T05F4A03C76C33C0D'); Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.IncludeSharePoint) { 'Sí' } else { 'No' }), $(if ($script:Settings.IncludeSharePoint) { 'Yes' } else { 'No' }))) },
            @{ Key='5'; Label=(Get-LocalizedString -Key 'T072048A32EDDCD0D'); Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.IncludeOneDrive) { 'Sí' } else { 'No' }), $(if ($script:Settings.IncludeOneDrive) { 'Yes' } else { 'No' }))) },
            @{ Key='6'; Label=(Get-LocalizedString -Key 'T534828BA9DE92BC0'); Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.ExcludeSystemSites) { 'Sí' } else { 'No' }), $(if ($script:Settings.ExcludeSystemSites) { 'Yes' } else { 'No' }))) },
            @{ Key='7'; Label=(Get-LocalizedString -Key 'T40AF8B61CFB012E6'); Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.PersistLogin) { 'Sí' } else { 'No' }), $(if ($script:Settings.PersistLogin) { 'Yes' } else { 'No' }))) },
            @{ Key='8'; Label=(Get-LocalizedString -Key 'TA3DEEDEE87C3E82D'); Description=(Get-LocalizedString -Key 'T8A86B2B2DEF9119B' -Values @($($script:Settings.MaxRetries), $($script:Settings.RetryDelay))) },
            @{ Key='9'; Label=(Get-LocalizedString -Key 'TFB5F3E7091BB532E'); Description=(Get-LocalizedString -Key 'TE8C25CA0BEF17D40') },
            @{ Key='10'; Label=(Get-LocalizedString -Key 'T940842D8DF8C5890'); Description=(Get-LocalizedString -Key 'T7DE227CE7EF59180') }
        )

        switch ($choice) {
            '0' { Save-Settings; return }
            '1' {
                $script:Settings.Tenant = Read-TextValue -Prompt (Get-LocalizedString -Key 'TF915851F215548DC') -Default ([string]$script:Settings.Tenant) `
                    -Validator { param($v) Test-TenantFormat $v } `
                    -ValidationMessage (Get-LocalizedString -Key 'T4E41BE171E54180F')
            }
            '2' { Show-AppRegistrationMenu }
            '3' { $script:Settings.ReportFolder = Read-TextValue -Prompt (Get-LocalizedString -Key 'TC5ED5CC66EA25BD5') -Default ([string]$script:Settings.ReportFolder) }
            '4' { $script:Settings.IncludeSharePoint = -not [bool]$script:Settings.IncludeSharePoint }
            '5' { $script:Settings.IncludeOneDrive = -not [bool]$script:Settings.IncludeOneDrive }
            '6' { $script:Settings.ExcludeSystemSites = -not [bool]$script:Settings.ExcludeSystemSites }
            '7' { $script:Settings.PersistLogin = -not [bool]$script:Settings.PersistLogin }
            '8' {
                $script:Settings.MaxRetries = Read-Integer -Prompt (Get-LocalizedString -Key 'T17F8A7EE7D2836C7') -Default ([int]$script:Settings.MaxRetries) -Minimum 1 -Maximum 10
                $script:Settings.RetryDelay = Read-Integer -Prompt (Get-LocalizedString -Key 'TFE86491BB28F90EB') -Default ([int]$script:Settings.RetryDelay) -Minimum 1 -Maximum 60
            }
            '9' {
                $script:Language = if ($script:Language -eq 'en') { 'es' } else { 'en' }
                $script:Settings.Language = $script:Language
            }
            '10' { Open-ReportsFolder }
        }
        Save-Settings
    }
}

function Show-Diagnostics {
    Write-AppHeader (Get-LocalizedString -Key 'TBBF03A4A7AC3A41E')
    $pnp = Get-InstalledPnPModule
    $tenantOk = Test-TenantFormat $script:Settings.Tenant
    $clientOk = Test-ClientIdFormat $script:Settings.ClientId
    $folderOk = Test-Path -LiteralPath $script:Settings.ReportFolder

    Write-Section (Get-LocalizedString -Key 'T6BF640A36D04A9D7')
    Write-Field 'PowerShell' $PSVersionTable.PSVersion $(if ($PSVersionTable.PSVersion -ge [version]'7.4') { 'Success' } else { 'Danger' })
    Write-Field 'PnP.PowerShell' $(if ($pnp) { $pnp.Version } else { Get-LocalizedString -Key 'TE81D97B25436053E' }) $(if ($pnp -and [version]$pnp.Version -ge $script:MinimumPnPVersion) { 'Success' } else { 'Danger' })
    Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $(if ($tenantOk) { $script:Settings.Tenant } else { Get-LocalizedString -Key 'T0BCC1B2949E0D07C' }) $(if ($tenantOk) { 'Success' } else { 'Warning' })
    Write-Field 'Client ID' $(if ($clientOk) { $script:Settings.ClientId } else { Get-LocalizedString -Key 'T0BCC1B2949E0D07C' }) $(if ($clientOk) { 'Success' } else { 'Warning' })
    Write-Field (Get-LocalizedString -Key 'T6107CBE19582BE77') $(if ($folderOk) { Get-LocalizedString -Key 'T5F1FB498D5AC98B9' } else { Get-LocalizedString -Key 'T316F477DF9D03327' }) $(if ($folderOk) { 'Success' } else { 'Warning' })

    Write-Section (Get-LocalizedString -Key 'TCAE54F3552C0E84E')
    Write-Field (Get-LocalizedString -Key 'TD2E63ADBD1B3CD73') @($script:LastInventory).Count
    Write-Field (Get-LocalizedString -Key 'T24233FFEAFA5A31B') $script:LastReport Muted
    Write-Field (Get-LocalizedString -Key 'TB01DF30BD2D65CCA') $(if ($script:AdminConnection) { Get-LocalizedString -Key 'T5EF13EC70DB15A81' } else { Get-LocalizedString -Key 'TA8A2E919238F2284' }) $(if ($script:AdminConnection) { 'Success' } else { 'Muted' })
    Pause-Tui
}

function Show-Context {
    Write-Section (Get-LocalizedString -Key 'T34DAFAC7B2AEA842')
    Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant
    Write-Field (Get-LocalizedString -Key 'T5C4361A89124A502') @($script:LastInventory).Count
    Write-Field (Get-LocalizedString -Key 'T24233FFEAFA5A31B') $script:LastReport Muted
}

function Clear-WorkingContext {
    $script:LastInventory = @()
    $script:LastReport = ''
    Write-Status Ok (Get-LocalizedString -Key 'T34EC709E7C49F79C')
}

function Show-MainMenu {
    while ($true) {
        Write-AppHeader (Get-LocalizedString -Key 'T80FF714B30CFF53F')
        Show-Context
        Write-Host ''

        $choice = Read-MenuChoice -Items @(
            @{ Key='1'; Label=(Get-LocalizedString -Key 'T5ACFBF3FDB36E3FC'); Description=(Get-LocalizedString -Key 'T46FD0615D558D339') },
            @{ Key='2'; Label=(Get-LocalizedString -Key 'T2AE3BC2415EDD9E9'); Description=(Get-LocalizedString -Key 'TBED8661029594CA8') },
            @{ Key='3'; Label=(Get-LocalizedString -Key 'TBBF03A4A7AC3A41E'); Description=(Get-LocalizedString -Key 'TFF17FC308F5462DB') },
            @{ Key='4'; Label=(Get-LocalizedString -Key 'T940842D8DF8C5890'); Description=(Get-LocalizedString -Key 'T252ABE94A18F569C') },
            @{ Key='5'; Label=(Get-LocalizedString -Key 'TF33260139CCAFC6D'); Description=(Get-LocalizedString -Key 'TD0452D8BA331528E') },
            @{ Key='0'; Label=(Get-LocalizedString -Key 'T54B061A7A1874670'); Description=(Get-LocalizedString -Key 'T2CF03D19A06FCB0A') }
        )

        switch ($choice) {
            '1' { Invoke-TenantInventory }
            '2' { Show-SettingsMenu }
            '3' { Show-Diagnostics }
            '4' { Open-ReportsFolder }
            '5' { Clear-WorkingContext; Pause-Tui }
            '0' { return }
        }
    }
}

try {
    Initialize-Terminal
    $script:Settings = Load-Settings
    Initialize-AppLanguage -Default ([string]$script:Settings.Language)
    $script:Settings.Language = $script:Language
    Save-Settings
    if (-not (Test-Path -LiteralPath $script:ConfigRoot)) {
        [void](New-Item -ItemType Directory -Path $script:ConfigRoot -Force)
    }
    Show-MainMenu
}
catch {
    Write-AppHeader (Get-LocalizedString -Key 'TC3EC5F1EFDC1CD2C')
    Write-Status Error $_.Exception.Message
    if ($DebugPreference -ne 'SilentlyContinue' -and -not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
        Write-Styled $_.ScriptStackTrace Muted
    }
    Pause-Tui
}
finally {
    Disconnect-Admin
}
