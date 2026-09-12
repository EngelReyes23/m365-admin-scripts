#requires -Version 7.4
<#
.SYNOPSIS
    M365 Site Storage Analyzer

.DESCRIPTION
    Analiza el uso y la cuota de sitios SharePoint Online y OneDrive for
    Business mediante Get-PnPTenantSite.

    La herramienta es de solo lectura. No elimina, mueve, archiva ni cambia
    permisos, propietarios o configuraciones del tenant.

    Incluye:
      - Resumen de hallazgos y conclusiones en pantalla.
      - CSV detallado para análisis posterior.
      - Umbrales configurables de advertencia y crítico.
      - TUI bilingüe, configuración persistente y diagnóstico.
      - Instalación/actualización de PnP.PowerShell.
      - Registro de una nueva aplicación PnP o uso de una existente.

.REQUIREMENTS
    PowerShell 7.4+
    PnP.PowerShell 3.2+
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$script:AppName = 'M365 Site Storage Analyzer'
$script:AppVersion = '1.0.0'
$script:MinimumPnPVersion = [version]'3.2.0'
$script:DefaultWarningPercent = 80
$script:DefaultCriticalPercent = 90
$script:DefaultMaxRetries = 5
$script:DefaultRetryDelay = 3
$script:Language = 'es'
$script:Settings = $null
$script:LastResults = @()
$script:LastReport = ''
$script:AdminConnection = $null
$script:Ansi = $false

$toolHome = if (-not [string]::IsNullOrWhiteSpace($env:M365_PNP_TOOLKIT_HOME)) {
    $env:M365_PNP_TOOLKIT_HOME
}
else {
    Join-Path $HOME '.m365-site-storage-analyzer'
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
        'T004FC77772100F64' = 'Clear results'
        'T0189AF383D0E36EB' = 'Every option requires a numeric Key and Label.'
        'T033C612F2B730951' = 'Folder where CSV files are saved.'
        'T034A7E198D0F964E' = 'Does not delete the Entra application.'
        'T0541F65548771B8B' = 'Installs the module for the current user.'
        'T05F4A03C76C33C0D' = 'Include SharePoint'
        'T05F5593316258D52' = '{1} site(s) exceed the {0}% warning threshold.'
        'T06114834B99A977B' = 'Save and validate a GUID Client ID.'
        'T072048A32EDDCD0D' = 'Include OneDrive'
        'T07C01E91E4F6D5DF' = 'Critical percentage'
        'T0886B5308A4B2E4B' = 'Known total usage (GB)'
        'T09297025FD78F4BC' = 'Largest consumers:'
        'T0AA6F91071C0E09A' = 'Could not read settings.json: {0}'
        'T0BCC1B2949E0D07C' = 'Not configured'
        'T17F8A7EE7D2836C7' = 'Maximum retries'
        'T18909694363DF916' = 'Included'
        'T1A7181893CBD2CBD' = 'Application validated and saved.'
        'T1B28D4A4F8DB5557' = 'Register new Entra application'
        'T1C617A82B8E4FA08' = 'Change report folder'
        'T1C98B1CF431D40C2' = 'Could not obtain a valid PnP connection.'
        'T1D1B97E641421A50' = 'Register a new application'
        'T21684FB55A9C0C7F' = 'Entra application / Client ID'
        'T22CB68F04A384C84' = 'Known total quota (GB)'
        'T23A28531ADC058A5' = 'Could not close the test connection: {0}'
        'T24233FFEAFA5A31B' = 'Last report'
        'T252ABE94A18F569C' = 'Opens generated CSV files.'
        'T290D52D21C62B0F0' = 'No critical sites were detected.'
        'T2AE3BC2415EDD9E9' = 'Settings'
        'T2CF03D19A06FCB0A' = 'Closes the tool.'
        'T2FDC3440E29935C3' = 'Warnings'
        'T316F477DF9D03327' = 'No'
        'T3398EA4E730E335A' = 'SharePoint sites'
        'T34DAFAC7B2AEA842' = 'Current context'
        'T366C34ED9B2D342F' = 'Only quota and usage are read; no site is modified.'
        'T38A305F6D93FCE98' = 'Value cannot be empty.'
        'T3F9DD0020AB38424' = 'Observed capacity is within configured thresholds.'
        'T45CC585A13EBF2B6' = 'Press Enter to continue'
        'T4F5665CECA963F1B' = 'Could not determine the admin center URL.'
        'T51E3D41527BA8048' = 'OneDrive'
        'T52E93CC5662EBAFE' = 'Site query'
        'T534828BA9DE92BC0' = 'Exclude system sites'
        'T54B061A7A1874670' = 'Exit'
        'T563959F6AD9ABA7C' = 'Remove local Client ID'
        'T57DDBCEBB8F1CBFA' = 'Analysis error'
        'T5EB08507E50EAE31' = 'Total analyzed'
        'T5F1FB498D5AC98B9' = 'Yes'
        'T5FADEDA3BED70263' = 'Application validation failed: {0}'
        'T600AA05D1DE02A5A' = 'Unavailable'
        'T6107CBE19582BE77' = 'Writable reports'
        'T614B22FE495A5A24' = 'Analysis completed in read-only mode.'
        'T615A36FC1B3ECB93' = '{0}: retry {1}/{3} in {2}s.'
        'T61ED27B50DEFC3AA' = 'Review priority sites before deciding on archiving or quota increases.'
        'T6377FF387E6EB3CA' = 'Use name.onmicrosoft.com.'
        'T67DB58904A00E641' = 'No sites exceeded the warning threshold.'
        'T6BF640A36D04A9D7' = 'Environment'
        'T6BFD73A8B70AC9CD' = 'Entra application / PnP authentication'
        'T6D986E2E69A4F27C' = 'Summary'
        'T6DF39E753D15E204' = 'Tests authentication and admin center access.'
        'T6F71913E9F915127' = 'Existing Client ID (Enter with no value to cancel)'
        'T70C8492F6D720D66' = 'No sites were found with the current filters.'
        'T719405B573B93C68' = 'Reports'
        'T7225AFDB6B9A872E' = 'Validate configured application'
        'T744B21B35A8DC099' = 'Could not open the reports folder.'
        'T790FBA20E4A85B79' = 'The report folder is not writable: {0}'
        'T7A2F1DB72073F287' = 'Install PnP, use an existing app, or register a new one.'
        'T7C8980091972774B' = 'The application may have been created, but the saved Client ID was not changed. Check consent/permissions and validate again using the displayed ID.'
        'T7D931A4CAEB7EDA4' = 'SharePoint'
        'T7DB06AAD1C19F5E7' = 'Remove the saved Client ID?'
        'T801A9EAD33FF2FE9' = 'Client ID removed.'
        'T80892645BA2B4B9F' = 'Excluded'
        'T80FF714B30CFF53F' = 'Home'
        'T817C87DEC30A5146' = 'Quota percentage for critical.'
        'T863E730B3F1DA130' = 'Critical'
        'T876978C64A8020C3' = 'A valid GUID is required.'
        'T8A997D127A26115F' = 'Results cleared.'
        'T8C0C562527CD36DC' = 'The Client ID is not a valid GUID.'
        'T8D4B140831155297' = 'Thresholds, tenant, authentication, filters, and reports.'
        'T8E30C29F9536FF73' = 'Change critical threshold'
        'T8EF3881D450B7B3A' = 'SharePoint admin center URL (check if the domain was renamed)'
        'T93E42BF8804F2A67' = 'Current: {1}.'
        'T940842D8DF8C5890' = 'Open reports folder'
        'T967A50F325D4A209' = 'Validates PowerShell, PnP, configuration, and reports.'
        'T9BCDC5AEC758501E' = 'Reports folder opened.'
        'T9E04B92C3940BFCF' = 'Change tenant'
        'T9E64B04D5EF3A9A2' = 'Analyze storage'
        'TA3569AD84A45AB2E' = '{0} site(s) are CRITICAL; they require priority review.'
        'TA3DEEDEE87C3E82D' = 'Configure retries'
        'TA4C5912BF7F8925A' = 'Creates and saves an Entra app for this analyzer.'
        'TAB17A0FE41D11328' = 'Start the analysis?'
        'TABC142D4249FBEF1' = 'Registration or validation failed: {0}'
        'TAE9E2DADFCE21C09' = 'Register this application?'
        'TAF67DD963947C57D' = 'Findings and conclusion'
        'TB141FEDEBD796396' = 'Results in memory'
        'TB1ED45769A1B1C85' = 'PnP Application'
        'TB68BB14DE138359F' = 'Clears results from this run.'
        'TB70528D786095E81' = 'Warning percentage'
        'TB7E83AF80BAA9767' = 'Use an existing application'
        'TBA0150205AD30BE3' = 'Register a new PnP application'
        'TBA1D61F98085B693' = 'Querying site usage and quota...'
        'TBB551C671A636DCE' = 'Storage analysis result'
        'TBBF03A4A7AC3A41E' = 'Diagnostics'
        'TBC0C62F4657BE857' = 'A valid nonzero GUID is required.'
        'TBC64A680FC28821E' = 'Application and access validated.'
        'TBD7FFB21E40C6526' = 'Quota percentage for warning.'
        'TC00AC5CED944A045' = 'Created Application (client) ID (Enter to cancel)'
        'TC3EC5F1EFDC1CD2C' = 'Unexpected error'
        'TC5ED5CC66EA25BD5' = 'Reports folder'
        'TCAE54F3552C0E84E' = 'Context'
        'TCD2FCB6D2EFE7AF5' = 'Critical'
        'TCE0DDB0921C76A21' = 'Aggregate usage'
        'TCF3F99F70AA7359C' = 'The reports folder is unavailable.'
        'TD15B34710D8330FD' = '{0} site(s) have no usable quota; the CSV needs detailed review.'
        'TD1E0FCA6A88559C6' = 'Warning'
        'TD25AF807973E77F6' = 'Compares usage and quota and shows priorities on screen.'
        'TD3CD73DA633DE7F5' = 'Enter a number between {0} and {1}.'
        'TD47590581B3251EE' = 'Install / update PnP.PowerShell'
        'TD966A6A2C7F0E31B' = 'Application name'
        'TDA312E9CFDF6B65B' = 'Enter the HTTPS SharePoint admin center URL.'
        'TE0955ABC442A2E77' = 'PowerShell {0} detected; 7.4+ is required.'
        'TE2C4D8495A4CE0C4' = 'Unknown quota'
        'TE81D97B25436053E' = 'Not installed'
        'TE8C49C15FBB66948' = 'Use the initial Entra domain ending in .onmicrosoft.com.'
        'TEAF6645EB92E0C61' = 'Detailed report'
        'TEBD48D58A9F26D21' = 'Validating authentication and access to the tenant inventory...'
        'TEC625AA49CFFCBE9' = 'Creates an Entra application. Requests delegated SharePoint AllSites.FullControl and delegated Graph User.Read. Requires authorization to register applications and administrator consent. Does not change sites; creates or saves no secrets.'
        'TF43296E41AEECCBA' = 'Client ID validated and saved.'
        'TF8F9BE4BC25501AF' = 'Saved app'
        'TF915851F215548DC' = 'Tenant'
        'TFB5F3E7091BB532E' = 'Change language'
        'TFC7B0F49E99778CE' = 'Change warning threshold'
        'TFC874B1E571463DF' = 'Initial tenant (Enter to cancel)'
        'TFE86491BB28F90EB' = 'Initial delay in seconds'
    }
    es = @{
        'T004FC77772100F64' = 'Limpiar resultados'
        'T0189AF383D0E36EB' = 'Cada opción requiere Key numérico y Label.'
        'T033C612F2B730951' = 'Ruta donde se guardan los CSV.'
        'T034A7E198D0F964E' = 'No elimina la aplicación de Entra.'
        'T0541F65548771B8B' = 'Instala el módulo para el usuario actual.'
        'T05F4A03C76C33C0D' = 'Incluir SharePoint'
        'T05F5593316258D52' = '{1} sitio(s) superan el umbral de advertencia de {0}%.'
        'T06114834B99A977B' = 'Guarda y valida un Client ID GUID.'
        'T072048A32EDDCD0D' = 'Incluir OneDrive'
        'T07C01E91E4F6D5DF' = 'Porcentaje crítico'
        'T0886B5308A4B2E4B' = 'Uso total conocido (GB)'
        'T09297025FD78F4BC' = 'Mayores consumidores:'
        'T0AA6F91071C0E09A' = 'No se pudo leer settings.json: {0}'
        'T0BCC1B2949E0D07C' = 'No configurado'
        'T17F8A7EE7D2836C7' = 'Máximo de reintentos'
        'T18909694363DF916' = 'Incluido'
        'T1A7181893CBD2CBD' = 'Aplicación validada y guardada.'
        'T1B28D4A4F8DB5557' = 'Registrar nueva aplicación Entra'
        'T1C617A82B8E4FA08' = 'Cambiar carpeta de reportes'
        'T1C98B1CF431D40C2' = 'No se pudo obtener una conexión PnP válida.'
        'T1D1B97E641421A50' = 'Registrar una nueva aplicación'
        'T21684FB55A9C0C7F' = 'Aplicación Entra / Client ID'
        'T22CB68F04A384C84' = 'Cuota total conocida (GB)'
        'T23A28531ADC058A5' = 'No se pudo cerrar la conexión de prueba: {0}'
        'T24233FFEAFA5A31B' = 'Último reporte'
        'T252ABE94A18F569C' = 'Abre los CSV generados.'
        'T290D52D21C62B0F0' = 'No se detectaron sitios en nivel crítico.'
        'T2AE3BC2415EDD9E9' = 'Configuración'
        'T2CF03D19A06FCB0A' = 'Cierra la herramienta.'
        'T2FDC3440E29935C3' = 'Advertencias'
        'T316F477DF9D03327' = 'No'
        'T3398EA4E730E335A' = 'Sitios SharePoint'
        'T34DAFAC7B2AEA842' = 'Contexto actual'
        'T366C34ED9B2D342F' = 'Solo se consultan cuotas y uso; no se modifica ningún sitio.'
        'T38A305F6D93FCE98' = 'El valor no puede estar vacío.'
        'T3F9DD0020AB38424' = 'La capacidad observada está dentro de los umbrales configurados.'
        'T45CC585A13EBF2B6' = 'Presiona Enter para continuar'
        'T4F5665CECA963F1B' = 'No se pudo determinar la URL del centro de administración.'
        'T51E3D41527BA8048' = 'OneDrive'
        'T52E93CC5662EBAFE' = 'Consulta de sitios'
        'T534828BA9DE92BC0' = 'Excluir sitios de sistema'
        'T54B061A7A1874670' = 'Salir'
        'T563959F6AD9ABA7C' = 'Quitar Client ID local'
        'T57DDBCEBB8F1CBFA' = 'Error de análisis'
        'T5EB08507E50EAE31' = 'Total analizado'
        'T5F1FB498D5AC98B9' = 'Sí'
        'T5FADEDA3BED70263' = 'Falló la validación de la aplicación: {0}'
        'T600AA05D1DE02A5A' = 'No disponible'
        'T6107CBE19582BE77' = 'Reportes escribibles'
        'T614B22FE495A5A24' = 'Análisis completado en modo solo lectura.'
        'T615A36FC1B3ECB93' = '{0}: reintento {1}/{3} en {2}s.'
        'T61ED27B50DEFC3AA' = 'Revisa los sitios prioritarios antes de tomar decisiones de archivado o ampliación de cuota.'
        'T6377FF387E6EB3CA' = 'Usa nombre.onmicrosoft.com.'
        'T67DB58904A00E641' = 'No se detectaron sitios por encima del umbral de advertencia.'
        'T6BF640A36D04A9D7' = 'Entorno'
        'T6BFD73A8B70AC9CD' = 'Aplicación Entra / autenticación PnP'
        'T6D986E2E69A4F27C' = 'Resumen'
        'T6DF39E753D15E204' = 'Prueba autenticación y acceso al centro de administración.'
        'T6F71913E9F915127' = 'Client ID existente (Enter sin valor para cancelar)'
        'T70C8492F6D720D66' = 'No se encontraron sitios con los filtros actuales.'
        'T719405B573B93C68' = 'Reportes'
        'T7225AFDB6B9A872E' = 'Validar aplicación configurada'
        'T744B21B35A8DC099' = 'No se pudo abrir la carpeta de reportes.'
        'T790FBA20E4A85B79' = 'La carpeta de reportes no es escribible: {0}'
        'T7A2F1DB72073F287' = 'Instala PnP, usa una app existente o registra una nueva.'
        'T7C8980091972774B' = 'La aplicación puede haberse creado, pero no se cambió el Client ID guardado. Revisa consentimiento/permisos y vuelve a validar usando el ID mostrado.'
        'T7D931A4CAEB7EDA4' = 'SharePoint'
        'T7DB06AAD1C19F5E7' = '¿Quitar el Client ID guardado?'
        'T801A9EAD33FF2FE9' = 'Client ID eliminado.'
        'T80892645BA2B4B9F' = 'Excluido'
        'T80FF714B30CFF53F' = 'Inicio'
        'T817C87DEC30A5146' = 'Porcentaje de cuota para nivel crítico.'
        'T863E730B3F1DA130' = 'Críticos'
        'T876978C64A8020C3' = 'Debe ser un GUID válido.'
        'T8A997D127A26115F' = 'Resultados limpiados.'
        'T8C0C562527CD36DC' = 'El Client ID no es un GUID válido.'
        'T8D4B140831155297' = 'Umbrales, tenant, autenticación, filtros y reportes.'
        'T8E30C29F9536FF73' = 'Cambiar umbral crítico'
        'T8EF3881D450B7B3A' = 'URL del centro de administración SharePoint (confirma si el dominio fue renombrado)'
        'T93E42BF8804F2A67' = 'Actual: {0}.'
        'T940842D8DF8C5890' = 'Abrir carpeta de reportes'
        'T967A50F325D4A209' = 'Valida PowerShell, PnP, configuración y reportes.'
        'T9BCDC5AEC758501E' = 'Carpeta de reportes abierta.'
        'T9E04B92C3940BFCF' = 'Cambiar tenant'
        'T9E64B04D5EF3A9A2' = 'Analizar almacenamiento'
        'TA3569AD84A45AB2E' = '{0} sitio(s) están en nivel CRÍTICO; requieren revisión prioritaria.'
        'TA3DEEDEE87C3E82D' = 'Configurar reintentos'
        'TA4C5912BF7F8925A' = 'Crea y guarda una app Entra para este analyzer.'
        'TAB17A0FE41D11328' = '¿Iniciar el análisis?'
        'TABC142D4249FBEF1' = 'Falló el registro o validación: {0}'
        'TAE9E2DADFCE21C09' = '¿Registrar esta aplicación?'
        'TAF67DD963947C57D' = 'Hallazgos y conclusión'
        'TB141FEDEBD796396' = 'Resultados en memoria'
        'TB1ED45769A1B1C85' = 'Aplicación PnP'
        'TB68BB14DE138359F' = 'Limpia resultados de esta ejecución.'
        'TB70528D786095E81' = 'Porcentaje de advertencia'
        'TB7E83AF80BAA9767' = 'Usar una aplicación existente'
        'TBA0150205AD30BE3' = 'Registrar una nueva aplicación PnP'
        'TBA1D61F98085B693' = 'Consultando uso y cuota de sitios...'
        'TBB551C671A636DCE' = 'Resultado del análisis de almacenamiento'
        'TBBF03A4A7AC3A41E' = 'Diagnóstico'
        'TBC0C62F4657BE857' = 'Debe ser un GUID válido distinto de cero.'
        'TBC64A680FC28821E' = 'Aplicación y acceso validados.'
        'TBD7FFB21E40C6526' = 'Porcentaje de cuota para advertencia.'
        'TC00AC5CED944A045' = 'Application (client) ID creado (Enter para cancelar)'
        'TC3EC5F1EFDC1CD2C' = 'Error inesperado'
        'TC5ED5CC66EA25BD5' = 'Carpeta de reportes'
        'TCAE54F3552C0E84E' = 'Contexto'
        'TCD2FCB6D2EFE7AF5' = 'Crítico'
        'TCE0DDB0921C76A21' = 'Uso agregado'
        'TCF3F99F70AA7359C' = 'La carpeta de reportes no está disponible.'
        'TD15B34710D8330FD' = '{0} sitio(s) no tienen cuota utilizable; el CSV requiere revisión detallada.'
        'TD1E0FCA6A88559C6' = 'Advertencia'
        'TD25AF807973E77F6' = 'Compara uso y cuota y muestra prioridades en pantalla.'
        'TD3CD73DA633DE7F5' = 'Introduce un número entre {0} y {1}.'
        'TD47590581B3251EE' = 'Instalar / actualizar PnP.PowerShell'
        'TD966A6A2C7F0E31B' = 'Nombre de la aplicación'
        'TDA312E9CFDF6B65B' = 'Introduce la URL HTTPS del centro de administración SharePoint.'
        'TE0955ABC442A2E77' = 'PowerShell {0} detectado; se requiere 7.4+.'
        'TE2C4D8495A4CE0C4' = 'Sin cuota conocida'
        'TE81D97B25436053E' = 'No instalado'
        'TE8C49C15FBB66948' = 'Usa el dominio inicial de Entra terminado en .onmicrosoft.com.'
        'TEAF6645EB92E0C61' = 'Reporte detallado'
        'TEBD48D58A9F26D21' = 'Validando autenticación y acceso al inventario del tenant...'
        'TEC625AA49CFFCBE9' = 'Creará una aplicación Entra. Solicitará AllSites.FullControl delegado de SharePoint y User.Read delegado de Graph. Requiere autorización para registrar aplicaciones y consentimiento administrativo. No cambia sitios; no crea ni guarda secretos.'
        'TF43296E41AEECCBA' = 'Client ID validado y guardado.'
        'TF8F9BE4BC25501AF' = 'App guardada'
        'TF915851F215548DC' = 'Tenant'
        'TFB5F3E7091BB532E' = 'Cambiar idioma'
        'TFC7B0F49E99778CE' = 'Cambiar umbral de advertencia'
        'TFC874B1E571463DF' = 'Tenant inicial (Enter para cancelar)'
        'TFE86491BB28F90EB' = 'Espera inicial en segundos'
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
    param([Parameter(Mandatory)][string]$Spanish, [Parameter(Mandatory)][string]$English)
    if ($script:Language -eq 'en') { return $English }
    return $Spanish
}

function Clear-AppScreen {
    try { if (-not [Console]::IsOutputRedirected) { Clear-Host } } catch { }
}

function Initialize-AppLanguage {
    param([string]$Default = 'en')
    Initialize-ToolkitLanguage -Default $Default -Name $script:AppName -Version $script:AppVersion
}

function Initialize-Terminal {
    try { $script:Ansi = -not [Console]::IsOutputRedirected } catch { $script:Ansi = $false }
}

function Get-AnsiCode {
    param([ValidateSet('Primary','Success','Info','Warning','Danger','Muted','Accent')][string]$Style)
    switch ($Style) {
        'Primary' { return "`e[96m" }; 'Success' { return "`e[92m" }; 'Info' { return "`e[94m" }
        'Warning' { return "`e[93m" }; 'Danger' { return "`e[91m" }; 'Accent' { return "`e[95m" }
        default { return "`e[90m" }
    }
}

function Write-Styled {
    param([AllowEmptyString()][string]$Text, [ValidateSet('Primary','Success','Info','Warning','Danger','Muted','Accent')][string]$Style='Primary', [switch]$NoNewline)
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
    param([Parameter(Mandatory)][string]$Name, [AllowNull()][object]$Value, [ValidateSet('Primary','Success','Info','Warning','Danger','Muted','Accent')][string]$Style='Primary')
    Write-ToolkitField $Name $Value $Style
}

function Write-Status {
    param([Parameter(Mandatory)][ValidateSet('Ok','Info','Warn','Error')][string]$Kind, [Parameter(Mandatory)][string]$Message)
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
    param([Parameter(Mandatory)][string]$Prompt, [string]$Default='', [switch]$AllowEmpty, [scriptblock]$Validator, [string]$ValidationMessage='Valor inválido.')
    while ($true) {
        $caption=if ($Default) { "$Prompt [$Default]" } else { $Prompt }
        $value=(Read-Host $caption).Trim()
        if ([string]::IsNullOrWhiteSpace($value) -and $Default) { $value=$Default }
        if ([string]::IsNullOrWhiteSpace($value)) { if ($AllowEmpty) { return '' }; Write-Status Error (Get-LocalizedString -Key 'T38A305F6D93FCE98'); continue }
        if ($Validator -and -not (& $Validator $value)) { Write-Status Error $ValidationMessage; continue }
        return $value
    }
}

function Read-YesNo {
    param([Parameter(Mandatory)][string]$Prompt, [bool]$Default=$false)
    return Read-ToolkitYesNo $Prompt $Default
}

function Read-Integer {
    param([Parameter(Mandatory)][string]$Prompt, [int]$Default, [int]$Minimum=1, [int]$Maximum=100)
    while ($true) {
        $raw=Read-Host "$Prompt [$Default]"
        if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
        $number=0
        if ([int]::TryParse($raw,[ref]$number) -and $number -ge $Minimum -and $number -le $Maximum) { return $number }
        Write-Status Error (Get-LocalizedString -Key 'TD3CD73DA633DE7F5' -Values @($Minimum, $Maximum))
    }
}

function New-DefaultSettings {
    $clientId=if ($env:ENTRAID_APP_ID) {$env:ENTRAID_APP_ID} elseif ($env:ENTRAID_CLIENT_ID) {$env:ENTRAID_CLIENT_ID} else {''}
    [pscustomobject]@{
        Tenant=''; ClientId=$clientId; AppRegistrationName=$script:AppName; ReportFolder=$script:DefaultReportFolder
        PersistLogin=$true; Language='es'; IncludeSharePoint=$true; IncludeOneDrive=$true; ExcludeSystemSites=$true
        WarningPercent=$script:DefaultWarningPercent; CriticalPercent=$script:DefaultCriticalPercent
        MaxRetries=$script:DefaultMaxRetries; RetryDelay=$script:DefaultRetryDelay
    }
}

function Merge-Settings {
    param($Loaded)
    $defaults=New-DefaultSettings
    if ($null -eq $Loaded) { return $defaults }
    foreach ($property in $defaults.PSObject.Properties.Name) {
        if ($null -eq $Loaded.PSObject.Properties[$property]) { Add-Member -InputObject $Loaded -NotePropertyName $property -NotePropertyValue $defaults.$property }
    }
    return $Loaded
}

function Load-Settings {
    $readPath = if (Test-Path -LiteralPath $script:ConfigPath) { $script:ConfigPath } else { $script:LegacyConfigPath }

    if (-not (Test-Path -LiteralPath $readPath)) { return New-DefaultSettings }
    try { return Merge-Settings (Get-Content -LiteralPath $readPath -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch { Write-Status Warn (Get-LocalizedString -Key 'T0AA6F91071C0E09A' -Values @($($_.Exception.Message))); return New-DefaultSettings }

}

function Save-Settings {
    if (-not (Test-Path -LiteralPath $script:ConfigRoot)) { [void](New-Item -ItemType Directory -Path $script:ConfigRoot -Force) }
    $script:Settings | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}

function Ensure-ReportFolder {
    if ([string]::IsNullOrWhiteSpace($script:Settings.ReportFolder)) { $script:Settings.ReportFolder=$script:DefaultReportFolder }
    if (-not (Test-Path -LiteralPath $script:Settings.ReportFolder)) { [void](New-Item -ItemType Directory -Path $script:Settings.ReportFolder -Force) }
    $testFile=Join-Path $script:Settings.ReportFolder ".write-test-$([guid]::NewGuid().ToString('N')).tmp"
    try { Set-Content -LiteralPath $testFile -Value 'test' -Encoding UTF8; Remove-Item -LiteralPath $testFile -Force; return $true }
    catch { Write-Status Error (Get-LocalizedString -Key 'T790FBA20E4A85B79' -Values @($($_.Exception.Message))); return $false }
}

function Open-ReportsFolder {
    try {
        if (-not (Ensure-ReportFolder)) { Pause-Tui; return }
        $folder=[System.IO.Path]::GetFullPath($script:Settings.ReportFolder)
        if ($IsWindows) { Start-Process -FilePath 'explorer.exe' -ArgumentList @($folder) | Out-Null }
        elseif (Get-Command xdg-open -ErrorAction SilentlyContinue) { Start-Process -FilePath 'xdg-open' -ArgumentList @($folder) | Out-Null }
        else { Invoke-Item -LiteralPath $folder }
        Write-Status Ok (Get-LocalizedString -Key 'T9BCDC5AEC758501E')
    }
    catch { Write-Status Error (Get-LocalizedString -Key 'T744B21B35A8DC099') }
    Pause-Tui
}

function Test-PowerShellVersion {
    if ($PSVersionTable.PSVersion -lt [version]'7.4.0') { Write-Status Error (Get-LocalizedString -Key 'TE0955ABC442A2E77' -Values @($($PSVersionTable.PSVersion))); return $false }
    return $true
}

function Get-InstalledPnPModule { Get-Module -Name PnP.PowerShell -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1 }

function Install-PnPModule {
    return Initialize-ToolkitPnP -MinimumVersion '3.2.0' -Update
}

function Ensure-PnPModule {
    param([switch]$OfferInstall)
    return Initialize-ToolkitPnP -MinimumVersion '3.2.0'
}

function Test-TenantFormat { param([AllowNull()][string]$Tenant); return (-not [string]::IsNullOrWhiteSpace($Tenant)) -and ($Tenant.Trim() -match '^[A-Za-z0-9-]+\.onmicrosoft\.com$') }
function Test-ClientIdFormat {
    param([AllowNull()][string]$ClientId)
    return Test-ToolkitClientId $ClientId
}
function Get-TenantPrefix { if (-not (Test-TenantFormat $script:Settings.Tenant)) { return '' }; return ($script:Settings.Tenant -replace '(?i)\.onmicrosoft\.com$','') }
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
    if (-not (Ensure-PnPModule)) { return $null }; if (-not (Ensure-TenantConfigured)) { return $null }; if (-not (Test-ClientIdFormat $ClientId)) { throw (Get-ToolkitString Invalid) }
    $adminUrl=Get-AdminUrl; if (-not $adminUrl) { throw (Get-LocalizedString -Key 'T4F5665CECA963F1B') }
    $params=@{Url=$adminUrl; ClientId=$ClientId; Tenant=[string]$script:Settings.Tenant; Interactive=$true; ReturnConnection=$true; ErrorAction='Stop'}
    if ($script:Settings.PersistLogin) { $params.PersistLogin=$true }
    return Connect-PnPOnline @params
}

function Connect-M365Admin {
    if (-not (Ensure-PnPModule)) { return $null }; if (-not (Ensure-TenantConfigured)) { return $null }; if (-not (Ensure-ClientId)) { return $null }
    $adminUrl=Get-AdminUrl; if (-not $adminUrl) { throw (Get-LocalizedString -Key 'T4F5665CECA963F1B') }
    $params=@{Url=$adminUrl; ClientId=[string]$script:Settings.ClientId; Tenant=[string]$script:Settings.Tenant; Interactive=$true; ReturnConnection=$true; ErrorAction='Stop'}
    if ($script:Settings.PersistLogin) { $params.PersistLogin=$true }
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
    if ($null -ne $script:AdminConnection) { try { Disconnect-PnPOnline -Connection $script:AdminConnection -ErrorAction SilentlyContinue } catch {}; $script:AdminConnection=$null }
}

function Invoke-WithRetry {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[string]$OperationName='operación')
    $attempt=0; $delay=[math]::Max(1,[int]$script:Settings.RetryDelay); $max=[math]::Max(1,[int]$script:Settings.MaxRetries)
    while ($true) {
        try { return & $ScriptBlock }
        catch {
            $attempt++; $message=$_.Exception.Message
            if ($message -notmatch '(?i)\b429\b|throttl|too many requests|\b503\b|service unavailable|temporarily unavailable|timeout|timed out' -or $attempt -ge $max) { throw }
            $sleep=[math]::Min(120,$delay+(Get-Random -Minimum 0 -Maximum 3))
            Write-Status Warn (Get-LocalizedString -Key 'T615A36FC1B3ECB93' -Values @(${OperationName}, $attempt, ${sleep}, $max))
            Start-Sleep -Seconds $sleep; $delay=[math]::Min(120,$delay*2)
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
function Get-SiteKind { param([Parameter(Mandatory)]$Site); if (Test-IsOneDriveUrl ([string](Get-PropertyValue $Site 'Url' ''))) { return 'OneDrive' }; return 'SharePoint' }
function Convert-MbToGb { param([AllowNull()]$Megabytes); $number=0.0; if ($null -eq $Megabytes -or -not [double]::TryParse(([string]$Megabytes),[ref]$number)) { return $null }; return [math]::Round($number/1024,2) }

function Convert-SiteToStorageResult {
    param([Parameter(Mandatory)]$Site)
    $usage=Get-PropertyValue $Site 'StorageUsageCurrent' $null; $quota=Get-PropertyValue $Site 'StorageQuota' $null; $percent=$null
    if ($null -ne $usage -and $null -ne $quota -and [double]$quota -gt 0) { $percent=[math]::Round(([double]$usage/[double]$quota)*100,2) }
    $level=if ($null -eq $percent) {'UNKNOWN'} elseif ($percent -ge [double]$script:Settings.CriticalPercent) {'CRITICAL'} elseif ($percent -ge [double]$script:Settings.WarningPercent) {'WARNING'} else {'NORMAL'}
    [pscustomobject]@{
        StorageStatus=$level; SiteType=(Get-SiteKind $Site); Title=[string](Get-PropertyValue $Site 'Title' ''); Url=[string](Get-PropertyValue $Site 'Url' '')
        Owner=[string](Get-PropertyValue $Site 'Owner' ''); Template=[string](Get-PropertyValue $Site 'Template' ''); Status=[string](Get-PropertyValue $Site 'Status' '')
        StorageUsageMB=if ($null -ne $usage) {[math]::Round([double]$usage,2)} else {$null}; StorageUsageGB=Convert-MbToGb $usage
        StorageQuotaMB=if ($null -ne $quota) {[math]::Round([double]$quota,2)} else {$null}; StorageQuotaGB=Convert-MbToGb $quota
        StorageUsagePercent=$percent; WarningPercent=[int]$script:Settings.WarningPercent; CriticalPercent=[int]$script:Settings.CriticalPercent
        IsSystemSite=[bool](Test-IsSystemSite $Site); AnalyzedAt=(Get-Date).ToString('o')
    }
}

function Get-StorageResults {
    param([Parameter(Mandatory)]$Connection)
    $sites=@(Invoke-WithRetry -OperationName (Get-LocalizedString -Key 'T52E93CC5662EBAFE') -ScriptBlock { Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Connection $Connection -ErrorAction Stop })
    $results=foreach ($site in $sites) {
        $kind=Get-SiteKind $site
        if ($kind -eq 'SharePoint' -and -not [bool]$script:Settings.IncludeSharePoint) { continue }
        if ($kind -eq 'OneDrive' -and -not [bool]$script:Settings.IncludeOneDrive) { continue }
        if ([bool]$script:Settings.ExcludeSystemSites -and (Test-IsSystemSite $site)) { continue }
        Convert-SiteToStorageResult -Site $site
    }
    return @($results | Sort-Object `
        @{Expression={switch($_.StorageStatus){'CRITICAL'{0};'WARNING'{1};'UNKNOWN'{2};default{3}}}}, `
        @{Expression='StorageUsageGB';Descending=$true}, `
        SiteType, Title)
}

function New-ReportPath { param([string]$Prefix); if (-not (Ensure-ReportFolder)) { throw (Get-LocalizedString -Key 'TCF3F99F70AA7359C') }; return Join-Path $script:Settings.ReportFolder "$Prefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv" }

function Export-StorageReport {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Records)
    $path=New-ReportPath 'SiteStorage'
    $columns=@('StorageStatus','SiteType','Title','Url','Owner','Template','Status','StorageUsageMB','StorageUsageGB','StorageQuotaMB','StorageQuotaGB','StorageUsagePercent','WarningPercent','CriticalPercent','IsSystemSite','AnalyzedAt')
    if (@($Records).Count -gt 0) { $Records | Select-Object $columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Force } else { ($columns -join ',') | Set-Content -LiteralPath $path -Encoding UTF8 }
    return $path
}

function Show-StorageSummary {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Records,[Parameter(Mandatory)][string]$ReportPath)
    Write-AppHeader (Get-LocalizedString -Key 'TBB551C671A636DCE')
    $sp=@($Records | Where-Object SiteType -eq 'SharePoint'); $od=@($Records | Where-Object SiteType -eq 'OneDrive'); $known=@($Records | Where-Object {$null -ne $_.StorageUsagePercent})
    $warning=@($Records | Where-Object StorageStatus -eq 'WARNING'); $critical=@($Records | Where-Object StorageStatus -eq 'CRITICAL'); $unknown=@($Records | Where-Object StorageStatus -eq 'UNKNOWN')
    $top=@($Records | Where-Object {$null -ne $_.StorageUsageGB} | Sort-Object StorageUsageGB -Descending | Select-Object -First 5)
    $totalUsage=($known | Measure-Object -Property StorageUsageGB -Sum).Sum; $totalQuota=($known | Measure-Object -Property StorageQuotaGB -Sum).Sum
    $aggregate=if ($totalQuota -gt 0) {[math]::Round(($totalUsage/$totalQuota)*100,2)} else {$null}
    Write-Section (Get-LocalizedString -Key 'T6D986E2E69A4F27C')
    Write-Field (Get-LocalizedString -Key 'T3398EA4E730E335A') $sp.Count Primary
    Write-Field (Get-LocalizedString -Key 'T51E3D41527BA8048') $od.Count Primary
    Write-Field (Get-LocalizedString -Key 'T5EB08507E50EAE31') @($Records).Count Info
    Write-Field (Get-LocalizedString -Key 'T0886B5308A4B2E4B') ([math]::Round([double]$totalUsage,2)) Info
    Write-Field (Get-LocalizedString -Key 'T22CB68F04A384C84') ([math]::Round([double]$totalQuota,2)) Info
    Write-Field (Get-LocalizedString -Key 'TCE0DDB0921C76A21') $(if ($null -ne $aggregate) {"$aggregate%"} else {Get-LocalizedString -Key 'T600AA05D1DE02A5A'}) $(if ($aggregate -ge $script:Settings.WarningPercent) {'Warning'} else {'Success'})
    Write-Field (Get-LocalizedString -Key 'T2FDC3440E29935C3') $warning.Count $(if ($warning.Count -gt 0) {'Warning'} else {'Success'})
    Write-Field (Get-LocalizedString -Key 'T863E730B3F1DA130') $critical.Count $(if ($critical.Count -gt 0) {'Danger'} else {'Success'})
    Write-Field (Get-LocalizedString -Key 'TE2C4D8495A4CE0C4') $unknown.Count $(if ($unknown.Count -gt 0) {'Warning'} else {'Success'})

    Write-Section (Get-LocalizedString -Key 'TAF67DD963947C57D')
    if (@($Records).Count -eq 0) {
        Write-Status Warn (Get-LocalizedString -Key 'T70C8492F6D720D66')
    }
    else {
        if ($critical.Count -gt 0) { Write-Status Error (Get-LocalizedString -Key 'TA3569AD84A45AB2E' -Values @($($critical.Count))) }
        else { Write-Status Ok (Get-LocalizedString -Key 'T290D52D21C62B0F0') }
        if ($warning.Count -gt 0) { Write-Status Warn (Get-LocalizedString -Key 'T05F5593316258D52' -Values @($($script:Settings.WarningPercent), $($warning.Count))) }
        else { Write-Status Ok (Get-LocalizedString -Key 'T67DB58904A00E641') }
        if ($unknown.Count -gt 0) { Write-Status Warn (Get-LocalizedString -Key 'TD15B34710D8330FD' -Values @($($unknown.Count))) }
        if ($top.Count -gt 0) {
            Write-Styled (Get-LocalizedString -Key 'T09297025FD78F4BC') Muted
            foreach ($site in $top) {
                $label=if ([string]::IsNullOrWhiteSpace([string]$site.Title)) {[string]$site.Url} else {[string]$site.Title}
                $style=switch($site.StorageStatus) {'CRITICAL' {'Danger'};'WARNING' {'Warning'};default {'Primary'}}
                Write-Styled ("  · {0} — {1:N2} GB ({2})" -f $label,$site.StorageUsageGB,$site.StorageStatus) $style
            }
        }
        if ($critical.Count -eq 0 -and $warning.Count -eq 0) { Write-Status Ok (Get-LocalizedString -Key 'T3F9DD0020AB38424') }
        else { Write-Status Info (Get-LocalizedString -Key 'T61ED27B50DEFC3AA') }
    }
    Write-Section (Get-LocalizedString -Key 'TEAF6645EB92E0C61')
    Write-Field 'CSV' $ReportPath Muted
    Write-Status Ok (Get-LocalizedString -Key 'T614B22FE495A5A24')
    Pause-Tui
}

function Invoke-StorageAnalysis {
    if (-not (Ensure-PnPModule)) { Pause-Tui; return }; if (-not (Ensure-TenantConfigured)) { Pause-Tui; return }; if (-not (Ensure-ClientId)) { Pause-Tui; return }; if (-not (Ensure-ReportFolder)) { Pause-Tui; return }
    Write-AppHeader (Get-LocalizedString -Key 'T9E64B04D5EF3A9A2')
    Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant Primary
    Write-Field (Get-LocalizedString -Key 'TD1E0FCA6A88559C6') "$($script:Settings.WarningPercent)%"
    Write-Field (Get-LocalizedString -Key 'TCD2FCB6D2EFE7AF5') "$($script:Settings.CriticalPercent)%"
    Write-Field (Get-LocalizedString -Key 'T7D931A4CAEB7EDA4') $(if ($script:Settings.IncludeSharePoint) {Get-LocalizedString -Key 'T18909694363DF916'} else {Get-LocalizedString -Key 'T80892645BA2B4B9F'})
    Write-Field 'OneDrive' $(if ($script:Settings.IncludeOneDrive) {Get-LocalizedString -Key 'T18909694363DF916'} else {Get-LocalizedString -Key 'T80892645BA2B4B9F'})
    Write-Status Info (Get-LocalizedString -Key 'T366C34ED9B2D342F')
    if (-not (Read-YesNo (Get-LocalizedString -Key 'TAB17A0FE41D11328') $true)) { return }
    try {
        $script:AdminConnection=Connect-M365Admin; if ($null -eq $script:AdminConnection) { throw (Get-LocalizedString -Key 'T1C98B1CF431D40C2') }
        Write-Status Info (Get-LocalizedString -Key 'TBA1D61F98085B693')
        $script:LastResults=@(Get-StorageResults -Connection $script:AdminConnection); $script:LastReport=Export-StorageReport -Records $script:LastResults; Show-StorageSummary -Records $script:LastResults -ReportPath $script:LastReport
    }
    catch { Write-AppHeader (Get-LocalizedString -Key 'T57DDBCEBB8F1CBFA'); Write-Status Error $_.Exception.Message; Pause-Tui }
    finally { Disconnect-Admin }
}

function Show-AppRegistrationMenu {
    while ($true) {
        Write-AppHeader (Get-LocalizedString -Key 'T6BFD73A8B70AC9CD')
        Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant; Write-Field (Get-LocalizedString -Key 'TF8F9BE4BC25501AF') $script:Settings.AppRegistrationName; Write-Field 'Client ID' $script:Settings.ClientId
        $choice=Read-MenuChoice -AllowBack -Items @(
            @{Key='1';Label=(Get-LocalizedString -Key 'TD47590581B3251EE');Description=(Get-LocalizedString -Key 'T0541F65548771B8B')},
            @{Key='2';Label=(Get-LocalizedString -Key 'TB7E83AF80BAA9767');Description=(Get-LocalizedString -Key 'T06114834B99A977B')},
            @{Key='3';Label=(Get-LocalizedString -Key 'TBA0150205AD30BE3');Description=(Get-LocalizedString -Key 'TA4C5912BF7F8925A')},
            @{Key='4';Label=(Get-LocalizedString -Key 'T7225AFDB6B9A872E');Description=(Get-LocalizedString -Key 'T6DF39E753D15E204')},
            @{Key='5';Label=(Get-LocalizedString -Key 'T563959F6AD9ABA7C');Description=(Get-LocalizedString -Key 'T034A7E198D0F964E')}
        )
        switch ($choice) {
            '0' {return}
            '1' {[void](Install-PnPModule);Pause-Tui}
            '2' {[void](Set-ExistingPnPClientId);Pause-Tui}
            '3' {[void](Register-NewPnPApp);Pause-Tui}
            '4' {[void](Test-PnPAppRegistration -ClientId ([string]$script:Settings.ClientId));Pause-Tui}
            '5' {if (Read-YesNo (Get-LocalizedString -Key 'T7DB06AAD1C19F5E7') $false) {$script:Settings.ClientId='';Save-Settings;Write-Status Ok (Get-LocalizedString -Key 'T801A9EAD33FF2FE9')};Pause-Tui}
        }
    }
}

function Show-SettingsMenu {
    while ($true) {
        Write-AppHeader (Get-LocalizedString -Key 'T2AE3BC2415EDD9E9')
        Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant; Write-Field 'Client ID' $script:Settings.ClientId; Write-Field (Get-LocalizedString -Key 'TD1E0FCA6A88559C6') "$($script:Settings.WarningPercent)%"; Write-Field (Get-LocalizedString -Key 'TCD2FCB6D2EFE7AF5') "$($script:Settings.CriticalPercent)%"; Write-Field (Get-LocalizedString -Key 'T719405B573B93C68') $script:Settings.ReportFolder Muted
        $choice=Read-MenuChoice -AllowBack -Items @(
            @{Key='1';Label=(Get-LocalizedString -Key 'T9E04B92C3940BFCF');Description='nombre.onmicrosoft.com'},
            @{Key='2';Label=(Get-LocalizedString -Key 'T21684FB55A9C0C7F');Description=(Get-LocalizedString -Key 'T7A2F1DB72073F287')},
            @{Key='3';Label=(Get-LocalizedString -Key 'TFC7B0F49E99778CE');Description=(Get-LocalizedString -Key 'TBD7FFB21E40C6526')},
            @{Key='4';Label=(Get-LocalizedString -Key 'T8E30C29F9536FF73');Description=(Get-LocalizedString -Key 'T817C87DEC30A5146')},
            @{Key='5';Label=(Get-LocalizedString -Key 'T1C617A82B8E4FA08');Description=(Get-LocalizedString -Key 'T033C612F2B730951')},
            @{Key='6';Label=(Get-LocalizedString -Key 'T05F4A03C76C33C0D');Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.IncludeSharePoint) {'Sí'} else {'No'}), $(if ($script:Settings.IncludeSharePoint) {'Yes'} else {'No'})))},
            @{Key='7';Label=(Get-LocalizedString -Key 'T072048A32EDDCD0D');Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.IncludeOneDrive) {'Sí'} else {'No'}), $(if ($script:Settings.IncludeOneDrive) {'Yes'} else {'No'})))},
            @{Key='8';Label=(Get-LocalizedString -Key 'T534828BA9DE92BC0');Description=(Get-LocalizedString -Key 'T93E42BF8804F2A67' -Values @($(if ($script:Settings.ExcludeSystemSites) {'Sí'} else {'No'}), $(if ($script:Settings.ExcludeSystemSites) {'Yes'} else {'No'})))},
            @{Key='9';Label=(Get-LocalizedString -Key 'TA3DEEDEE87C3E82D');Description="$($script:Settings.MaxRetries) · $($script:Settings.RetryDelay)s"},
            @{Key='10';Label=(Get-LocalizedString -Key 'TFB5F3E7091BB532E');Description='Español / English'},
            @{Key='11';Label=(Get-LocalizedString -Key 'T940842D8DF8C5890');Description=(Get-LocalizedString -Key 'T252ABE94A18F569C')}
        )
        switch ($choice) {
            '0' {Save-Settings;return}
            '1' {$script:Settings.Tenant=Read-TextValue (Get-LocalizedString -Key 'TF915851F215548DC') ([string]$script:Settings.Tenant) -Validator {param($v) Test-TenantFormat $v} -ValidationMessage (Get-LocalizedString -Key 'T6377FF387E6EB3CA')}
            '2' {Show-AppRegistrationMenu}
            '3' {$script:Settings.WarningPercent=Read-Integer (Get-LocalizedString -Key 'TB70528D786095E81') ([int]$script:Settings.WarningPercent) 1 99; if ($script:Settings.WarningPercent -ge $script:Settings.CriticalPercent) {$script:Settings.CriticalPercent=[math]::Min(100,$script:Settings.WarningPercent+10)}}
            '4' {$script:Settings.CriticalPercent=Read-Integer (Get-LocalizedString -Key 'T07C01E91E4F6D5DF') ([int]$script:Settings.CriticalPercent) 2 100; if ($script:Settings.CriticalPercent -le $script:Settings.WarningPercent) {$script:Settings.WarningPercent=[math]::Max(1,$script:Settings.CriticalPercent-10)}}
            '5' {$script:Settings.ReportFolder=Read-TextValue (Get-LocalizedString -Key 'TC5ED5CC66EA25BD5') ([string]$script:Settings.ReportFolder)}
            '6' {$script:Settings.IncludeSharePoint=-not [bool]$script:Settings.IncludeSharePoint}
            '7' {$script:Settings.IncludeOneDrive=-not [bool]$script:Settings.IncludeOneDrive}
            '8' {$script:Settings.ExcludeSystemSites=-not [bool]$script:Settings.ExcludeSystemSites}
            '9' {$script:Settings.MaxRetries=Read-Integer (Get-LocalizedString -Key 'T17F8A7EE7D2836C7') ([int]$script:Settings.MaxRetries) 1 10; $script:Settings.RetryDelay=Read-Integer (Get-LocalizedString -Key 'TFE86491BB28F90EB') ([int]$script:Settings.RetryDelay) 1 60}
            '10' {$script:Language=if ($script:Language -eq 'en') {'es'} else {'en'}; $script:Settings.Language=$script:Language}
            '11' {Open-ReportsFolder}
        }
        Save-Settings
    }
}

function Show-Diagnostics {
    Write-AppHeader (Get-LocalizedString -Key 'TBBF03A4A7AC3A41E'); $pnp=Get-InstalledPnPModule; $tenantOk=Test-TenantFormat $script:Settings.Tenant; $clientOk=Test-ClientIdFormat $script:Settings.ClientId; $folderOk=Test-Path -LiteralPath $script:Settings.ReportFolder
    Write-Section (Get-LocalizedString -Key 'T6BF640A36D04A9D7')
    Write-Field 'PowerShell' $PSVersionTable.PSVersion $(if ($PSVersionTable.PSVersion -ge [version]'7.4') {'Success'} else {'Danger'})
    Write-Field 'PnP.PowerShell' $(if ($pnp) {$pnp.Version} else {Get-LocalizedString -Key 'TE81D97B25436053E'}) $(if ($pnp -and [version]$pnp.Version -ge $script:MinimumPnPVersion) {'Success'} else {'Danger'})
    Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $(if ($tenantOk) {$script:Settings.Tenant} else {Get-LocalizedString -Key 'T0BCC1B2949E0D07C'}) $(if ($tenantOk) {'Success'} else {'Warning'})
    Write-Field 'Client ID' $(if ($clientOk) {$script:Settings.ClientId} else {Get-LocalizedString -Key 'T0BCC1B2949E0D07C'}) $(if ($clientOk) {'Success'} else {'Warning'})
    Write-Field (Get-LocalizedString -Key 'T6107CBE19582BE77') $(if ($folderOk) {Get-LocalizedString -Key 'T5F1FB498D5AC98B9'} else {Get-LocalizedString -Key 'T316F477DF9D03327'}) $(if ($folderOk) {'Success'} else {'Warning'})
    Write-Section (Get-LocalizedString -Key 'TCAE54F3552C0E84E'); Write-Field (Get-LocalizedString -Key 'TB141FEDEBD796396') @($script:LastResults).Count; Write-Field (Get-LocalizedString -Key 'T24233FFEAFA5A31B') $script:LastReport Muted
    Pause-Tui
}

function Show-MainMenu {
    while ($true) {
        Write-AppHeader (Get-LocalizedString -Key 'T80FF714B30CFF53F'); Write-Section (Get-LocalizedString -Key 'T34DAFAC7B2AEA842'); Write-Field (Get-LocalizedString -Key 'TF915851F215548DC') $script:Settings.Tenant; Write-Field (Get-LocalizedString -Key 'TD1E0FCA6A88559C6') "$($script:Settings.WarningPercent)%"; Write-Field (Get-LocalizedString -Key 'TCD2FCB6D2EFE7AF5') "$($script:Settings.CriticalPercent)%"; Write-Field (Get-LocalizedString -Key 'T24233FFEAFA5A31B') $script:LastReport Muted
        $choice=Read-MenuChoice -Items @(
            @{Key='1';Label=(Get-LocalizedString -Key 'T9E64B04D5EF3A9A2');Description=(Get-LocalizedString -Key 'TD25AF807973E77F6')},
            @{Key='2';Label=(Get-LocalizedString -Key 'T2AE3BC2415EDD9E9');Description=(Get-LocalizedString -Key 'T8D4B140831155297')},
            @{Key='3';Label=(Get-LocalizedString -Key 'TBBF03A4A7AC3A41E');Description=(Get-LocalizedString -Key 'T967A50F325D4A209')},
            @{Key='4';Label=(Get-LocalizedString -Key 'T940842D8DF8C5890');Description=(Get-LocalizedString -Key 'T252ABE94A18F569C')},
            @{Key='5';Label=(Get-LocalizedString -Key 'T004FC77772100F64');Description=(Get-LocalizedString -Key 'TB68BB14DE138359F')},
            @{Key='0';Label=(Get-LocalizedString -Key 'T54B061A7A1874670');Description=(Get-LocalizedString -Key 'T2CF03D19A06FCB0A')}
        )
        switch ($choice) {
            '1' {Invoke-StorageAnalysis}
            '2' {Show-SettingsMenu}
            '3' {Show-Diagnostics}
            '4' {Open-ReportsFolder}
            '5' {$script:LastResults=@();$script:LastReport='';Write-Status Ok (Get-LocalizedString -Key 'T8A997D127A26115F');Pause-Tui}
            '0' {return}
        }
    }
}

try {
    Initialize-Terminal
    $script:Settings=Load-Settings
    Initialize-AppLanguage -Default ([string]$script:Settings.Language)
    $script:Settings.Language=$script:Language
    Save-Settings
    if (-not (Test-Path -LiteralPath $script:ConfigRoot)) {[void](New-Item -ItemType Directory -Path $script:ConfigRoot -Force)}
    Show-MainMenu
}
catch { Write-AppHeader (Get-LocalizedString -Key 'TC3EC5F1EFDC1CD2C'); Write-Status Error $_.Exception.Message; Pause-Tui }
finally { Disconnect-Admin }
