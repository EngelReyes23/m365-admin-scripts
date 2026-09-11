#requires -Version 7.4
<#
.SYNOPSIS
    Microsoft 365 Legacy UserId Cleaner

.DESCRIPTION
    Herramienta interactiva para detectar y reparar entradas legacy de usuarios
    recreados en las UserInfoList de SharePoint Online y OneDrive for Business.

    Flujo recomendado:
      1. Validar autenticación y usuario afectado.
      2. Enumerar los sitios del tenant.
      3. Registrar el estado previo de acceso administrativo cuando sea posible.
      4. Agregar temporalmente al operador como Site Collection Administrator
         en los sitios que lo requieran.
      5. Desconectar del Admin Center.
      6. Volver a conectar sitio por sitio con reintentos de propagación.
      7. Comparar el SID actual del usuario con UserId.NameId almacenado.
      8. En modo REAL, eliminar únicamente entradas con ID mismatch confirmado.
      9. Retirar únicamente los Site Collection Admin temporales añadidos por
         esta herramienta, preservando los accesos protegidos/preexistentes.
     10. Guardar auditoría y un estado recuperable si el proceso se interrumpe.

    PROTECCIONES:
      - Modo SIMULACIÓN por defecto: no elimina usuarios de UserInfoList.
      - La elevación temporal de Site Collection Admin puede ocurrir también
        durante simulación, porque es necesaria para inspeccionar sitios sin acceso.
      - El OneDrive propio del USUARIO AFECTADO nunca se modifica.
      - Los sitios SPO protegidos se conservan como Site Collection Admin.
      - Por defecto, OneDrive ajenos con acceso administrativo PREEXISTENTE
        no se alteran. Solo se retiran elevaciones creadas por esta ejecución.
      - Existe una recuperación de grants temporales si la ejecución se interrumpe.

.REQUIREMENTS
    PowerShell 7.4+
    PnP.PowerShell 3.2+
    Rol SharePoint Administrator o Global Administrator para las operaciones tenant.
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# =============================================================================
# APLICACIÓN
# =============================================================================

$script:AppName = 'Microsoft 365 Legacy UserId Cleaner'
$script:AppVersion = '2.2.0'
$script:MinimumPnPVersion = [version]'3.2.0'

$script:DefaultMaxRetries = 5
$script:DefaultRetryDelaySeconds = 3
$script:DefaultPropagationTimeoutSeconds = 120
$script:DefaultPropagationInitialDelaySeconds = 3

$toolHome = if (-not [string]::IsNullOrWhiteSpace($env:M365_LEGACY_USERID_CLEANER_HOME)) {
    $env:M365_LEGACY_USERID_CLEANER_HOME
}
else {
    Join-Path $HOME '.m365-legacy-userid-cleaner'
}

$script:ConfigRoot = $toolHome
$script:ConfigPath = Join-Path $script:ConfigRoot 'settings.json'
$script:StatePath = Join-Path $script:ConfigRoot 'active-admin-grants.json'
$script:ReportsFolder = Join-Path $script:ConfigRoot 'reports'

$script:Settings = $null
$script:Ansi = $false
$script:AdminConnection = $null
$script:AffectedUser = $null
$script:LastInventory = @()
$script:LastResults = @()

# =============================================================================
# LANGUAGE
# =============================================================================

$script:Language = 'es'
$script:UiTranslations = @(
    @{ From = 'Contexto'; To = 'Context' }; @{ From = 'Inicio'; To = 'Home' }; @{ From = 'Finalizado'; To = 'Finished' }
    @{ From = 'Error inesperado'; To = 'Unexpected error' }; @{ From = 'Detalles técnicos:'; To = 'Technical details:' }
    @{ From = 'Configuración'; To = 'Settings' }; @{ From = 'Salir'; To = 'Exit' }; @{ From = 'Volver'; To = 'Back' }
    @{ From = 'Cancelar'; To = 'Cancel' }; @{ From = 'Diagnóstico'; To = 'Diagnostics' }; @{ From = 'Autenticación'; To = 'Authentication' }
    @{ From = 'Selecciona una opción'; To = 'Select an option' }; @{ From = 'Seleccionar'; To = 'Select' }; @{ From = 'Sitio'; To = 'Site' }
    @{ From = 'Bibliotecas'; To = 'Libraries' }; @{ From = 'Biblioteca'; To = 'Library' }; @{ From = 'Usuario afectado'; To = 'Affected user' }
    @{ From = 'SIMULACIÓN'; To = 'SIMULATION' }; @{ From = 'REAL'; To = 'LIVE' }; @{ From = 'Activar modo REAL'; To = 'Enable LIVE mode' }
    @{ From = 'Escribe REPARAR para continuar'; To = 'Type REPAIR to continue' }; @{ From = 'Procesando'; To = 'Processing' }
    @{ From = 'Preparando acceso'; To = 'Preparing access' }; @{ From = 'Restaurando Site Collection Admin'; To = 'Restoring Site Collection Admin' }
    @{ From = 'UserInfoList'; To = 'UserInfoList' }; @{ From = 'Revisión'; To = 'Review' }; @{ From = 'Error'; To = 'Error' }
)
function Get-LocalizedText {
    param([AllowNull()][object]$Text)
    if ($null -eq $Text) { return '' }; $result = [string]$Text
    if ($script:Language -eq 'en') {
        $phrases = @(
            @{ From = 'Selecciona una opción'; To = 'Select an option' }; @{ From = 'Seleccionar / validar usuario afectado'; To = 'Select / validate affected user' }
            @{ From = 'Idioma'; To = 'Language' }; @{ From = 'Cambiar idioma'; To = 'Change language' }
            @{ From = 'No seleccionado'; To = 'Not selected' }; @{ From = 'No configurado'; To = 'Not configured' }; @{ From = 'No instalado'; To = 'Not installed' }
            @{ From = 'Destino'; To = 'Target' }; @{ From = 'CONFIGURADA'; To = 'CONFIGURED' }; @{ From = 'NO CONFIGURADA'; To = 'NOT CONFIGURED' }
            @{ From = 'Activado'; To = 'Enabled' }; @{ From = 'Desactivado'; To = 'Disabled' }; @{ From = 'Sí'; To = 'Yes' }; @{ From = 'SÍ'; To = 'YES' }
            @{ From = 'Ejecutar reparación tenant'; To = 'Run tenant repair' }; @{ From = 'Cambiar a modo REAL'; To = 'Switch to LIVE mode' }
            @{ From = 'Cambiar a SIMULACIÓN'; To = 'Switch to SIMULATION' }; @{ From = 'Aplicación Entra / autenticación PnP'; To = 'Entra application / PnP authentication' }
            @{ From = 'Sitios SPO protegidos'; To = 'Protected SPO sites' }; @{ From = 'Recuperar / retirar grants temporales'; To = 'Recover / remove temporary grants' }
            @{ From = 'Escribe REPARAR para continuar'; To = 'Type REPAIR to continue' }; @{ From = '¿Iniciar el análisis?'; To = 'Start the analysis?' }
            @{ From = 'MODO REAL: se eliminarán únicamente mismatches con SID actual válido y comparación inequívoca.'; To = 'LIVE MODE: only mismatches with a valid current SID and unambiguous comparison will be removed.' }
            @{ From = 'SIMULACIÓN no elimina UserInfoList; los grants administrativos sí son reales.'; To = 'SIMULATION does not remove UserInfoList entries; administrative grants are still real.' }
            @{ From = 'Selecciona primero el usuario afectado.'; To = 'Select the affected user first.' }
            @{ From = 'Identidad actual confirmada.'; To = 'Current identity confirmed.' }; @{ From = 'Validando la identidad ACTUAL del usuario...'; To = 'Validating the user CURRENT identity...' }
            @{ From = 'No hay una sesión pendiente de limpieza.'; To = 'There is no pending cleanup session.' }
            @{ From = 'Recuperación finalizada.'; To = 'Recovery finished.' }; @{ From = 'Operación cancelada.'; To = 'Operation canceled.' }
            @{ From = 'No se encontraron sitios.'; To = 'No sites were found.' }; @{ From = 'No hay sitios protegidos configurados.'; To = 'No protected sites are configured.' }
            @{ From = 'Sitio agregado a la lista protegida.'; To = 'Site added to the protected list.' }; @{ From = 'Sitio retirado de la lista protegida.'; To = 'Site removed from the protected list.' }
            @{ From = 'No se recomienda procesar hosts raíz ni colecciones de sistema.'; To = 'Processing root hosts or system collections is not recommended.' }
            @{ From = 'Esta espera se aplica UNA sola vez después de completar todos los grants.'; To = 'This wait is applied ONCE after all grants are completed.' }
            @{ From = 'Los sitios que todavía devuelvan 401/403 se reintentan individualmente.'; To = 'Sites that still return 401/403 are retried individually.' }
            @{ From = 'Validando la identidad ACTUAL del usuario...'; To = 'Validating the user CURRENT identity...' }
            @{ From = 'El SID actual está vacío. Se bloquea la ejecución para impedir falsos positivos.'; To = 'The current SID is empty. Execution is blocked to prevent false positives.' }
            @{ From = 'Opción no válida.'; To = 'Invalid option.' }; @{ From = 'Opción inválida.'; To = 'Invalid option.' }; @{ From = 'Selección inválida.'; To = 'Invalid selection.' }
            @{ From = 'Valor inválido.'; To = 'Invalid value.' }; @{ From = 'El valor no puede estar vacío.'; To = 'Value cannot be empty.' }
            @{ From = 'Prueba token, Admin Center, Get-PnPTenant y enumeración de sitios.'; To = 'Tests the token, Admin Center, Get-PnPTenant, and site enumeration.' }
            @{ From = 'La configuración actual solo cambia después de validar correctamente.'; To = 'The current configuration changes only after successful validation.' }
            @{ From = 'Crea y valida una nueva app de PnP con AllSites.FullControl delegado.'; To = 'Creates and validates a new PnP app with delegated AllSites.FullControl.' }
            @{ From = 'Limpia la autenticación persistida por PnP.PowerShell.'; To = 'Clears persisted PnP.PowerShell authentication.' }
            @{ From = 'Se preserva Site Collection Admin después de las ejecuciones.'; To = 'Site Collection Admin is preserved after runs.' }
            @{ From = 'Buscar dinámicamente un sitio SPO del tenant.'; To = 'Dynamically search for an SPO site in the tenant.' }
            @{ From = 'Dejará de preservarse por esta política.'; To = 'It will no longer be preserved by this policy.' }
            @{ From = 'Estos sitios conservan tu Site Collection Admin al terminar.'; To = 'These sites retain your Site Collection Admin when finished.' }
            @{ From = 'Dominio inicial'; To = 'Initial domain' }; @{ From = 'UPN del administrador'; To = 'Administrator UPN' }; @{ From = 'Alcance predeterminado'; To = 'Default scope' }
            @{ From = 'Timeout de propagación'; To = 'Propagation timeout' }; @{ From = 'Exclusiones de sitios de sistema'; To = 'System-site exclusions' }
            @{ From = 'Recomendado: mantener activado.'; To = 'Recommended: keep enabled.' }; @{ From = 'Muestra roots, templates y fragmentos de URL excluidos.'; To = 'Shows excluded roots, templates, and URL fragments.' }
            @{ From = 'Solo sitios SPO.'; To = 'SPO sites only.' }; @{ From = 'Recorre ambos tipos.'; To = 'Processes both types.' }
            @{ From = 'Excluye automáticamente el OneDrive propio del usuario afectado.'; To = "Automatically excludes the affected user's own OneDrive." }
            @{ From = 'Esta opción intenta retirar únicamente los grants marcados como temporales en el estado guardado.'; To = 'This option attempts to remove only grants marked as temporary in the saved state.' }
            @{ From = 'No se recomienda procesar hosts raíz ni colecciones de sistema.'; To = 'Processing root hosts or system collections is not recommended.' }
            @{ From = 'Validando la identidad ACTUAL del usuario...'; To = "Validating the user's CURRENT identity..." }
            @{ From = 'Existe un estado de grants administrativos pendiente de recuperación.'; To = 'Recoverable administrative-grant state exists.' }
            @{ From = 'Aplicación Entra / autenticación PnP'; To = 'Entra application / PnP authentication' }
            @{ From = 'Validar aplicación configurada'; To = 'Validate configured application' }
            @{ From = 'Usar otra aplicación existente'; To = 'Use another existing application' }
            @{ From = 'La configuración actual solo cambia después de validar correctamente.'; To = 'The current configuration changes only after successful validation.' }
            @{ From = 'Registrar una nueva aplicación Entra'; To = 'Register a new Entra application' }
            @{ From = 'Quitar Client ID de la configuración local'; To = 'Remove Client ID from local configuration' }
            @{ From = 'Eliminar sesión persistente'; To = 'Clear persisted session' }
            @{ From = 'Autenticación requerida'; To = 'Authentication required' }
            @{ From = 'Validar un Client ID ya registrado antes de guardarlo.'; To = 'Validate an already registered Client ID before saving it.' }
            @{ From = 'Crear una app PnP nueva y validarla.'; To = 'Create and validate a new PnP app.' }
            @{ From = 'No hay una App Registration válida configurada.'; To = 'No valid App Registration is configured.' }
            @{ From = 'Selecciona cómo quieres configurar la autenticación.'; To = 'Select how you want to configure authentication.' }
            @{ From = 'Prueba token, Admin Center, Get-PnPTenant y enumeración de sitios.'; To = 'Tests the token, Admin Center, Get-PnPTenant, and site enumeration.' }
            @{ From = 'No elimina la aplicación en Entra ID.'; To = 'Does not delete the application in Entra ID.' }
            @{ From = 'Quitar configuración de aplicación'; To = 'Remove application configuration' }
            @{ From = 'Esto NO elimina la App Registration en Entra ID.'; To = 'This does NOT delete the App Registration in Entra ID.' }
            @{ From = '¿Quitar el Client ID de esta herramienta?'; To = 'Remove the Client ID from this tool?' }
            @{ From = 'Seleccionar sitio protegido'; To = 'Select protected site' }
            @{ From = 'Quitar sitio SPO protegido'; To = 'Remove protected SPO site' }
            @{ From = 'Agregar sitio protegido'; To = 'Add protected site' }
            @{ From = 'Quitar sitio protegido'; To = 'Remove protected site' }
            @{ From = 'Buscar dinámicamente un sitio SPO del tenant.'; To = 'Dynamically search for an SPO site in the tenant.' }
            @{ From = 'Dejará de preservarse por esta política.'; To = 'It will no longer be preserved by this policy.' }
            @{ From = 'Configurados: '; To = 'Configured: ' }
            @{ From = 'Estos sitios conservan tu Site Collection Admin al terminar.'; To = 'These sites retain your Site Collection Admin when finished.' }
            @{ From = 'Fast path / propagación'; To = 'Fast path / propagation' }
            @{ From = 'Validación de App Registration'; To = 'App Registration validation' }
            @{ From = 'Sesión persistente'; To = 'Persisted session' }
            @{ From = 'Solo SharePoint'; To = 'SharePoint only' }
            @{ From = 'Solo OneDrive'; To = 'OneDrive only' }
            @{ From = 'SharePoint + OneDrive'; To = 'SharePoint + OneDrive' }
            @{ From = 'sitio(s)'; To = 'site(s)' }
            @{ From = 'segundos'; To = 'seconds' }
            @{ From = 'intento(s)'; To = 'attempt(s)' }
            @{ From = 'Estas reglas se evalúan antes de conceder Site Collection Admin.'; To = 'These rules are evaluated before granting Site Collection Admin.' }
            @{ From = 'Recomendado: mantener activado.'; To = 'Recommended: keep enabled.' }
            @{ From = 'Muestra roots, templates y fragmentos de URL excluidos.'; To = 'Shows excluded roots, templates, and URL fragments.' }
            @{ From = 'Desactivar exclusiones de sistema'; To = 'Disable system exclusions' }
            @{ From = 'Activar exclusiones de sistema'; To = 'Enable system exclusions' }
            @{ From = 'Ver reglas activas'; To = 'View active rules' }
            @{ From = 'SharePoint Online'; To = 'SharePoint Online' }
            @{ From = 'OneDrive for Business'; To = 'OneDrive for Business' }
            @{ From = 'SPO no protegido se considera acceso temporal.'; To = 'Unprotected SPO is treated as temporary access.' }
            @{ From = 'Own OneDrive del afectado se omite; ODB de terceros se limpia al final.'; To = "The affected user's own OneDrive is skipped; third-party ODB is cleaned at the end." }
            @{ From = 'Ejecuta ambos alcances.'; To = 'Runs both scopes.' }
            @{ From = 'Alcance de esta ejecución'; To = 'Scope for this run' }
            @{ From = 'Predeterminado: '; To = 'Default: ' }
            @{ From = 'Seleccionar / validar usuario afectado'; To = 'Select / validate affected user' }
            @{ From = 'Resuelve AccountName, SID actual y OneDrive propio.'; To = "Resolves AccountName, current SID, and the user's own OneDrive." }
            @{ From = 'Ejecutar reparación tenant'; To = 'Run tenant repair' }
            @{ From = 'Grant masivo → espera global → procesar directo → restaurar admins.'; To = 'Bulk grant → global wait → direct processing → restore admins.' }
            @{ From = 'Cambiar a SIMULACIÓN'; To = 'Switch to SIMULATION' }
            @{ From = 'SIMULACIÓN no elimina UserInfoList; los grants temporales sí pueden aplicarse.'; To = 'SIMULATION does not remove UserInfoList entries; temporary grants may still be applied.' }
            @{ From = 'REAL elimina únicamente mismatches inequívocos.'; To = 'LIVE removes only unambiguous mismatches.' }
            @{ From = 'Configurada · validar, cambiar o registrar otra app.'; To = 'Configured · validate, change, or register another app.' }
            @{ From = 'No configurada · usar existente o registrar nueva.'; To = 'Not configured · use an existing app or register a new one.' }
            @{ From = 'Recuperar / retirar grants temporales'; To = 'Recover / remove temporary grants' }
            @{ From = 'Hay un estado pendiente de restauración.'; To = 'A pending restoration state exists.' }
            @{ From = 'No hay grants pendientes registrados.'; To = 'No pending grants are registered.' }
            @{ From = 'Tenant, admin, alcance, exclusiones de sistema, fast path y autenticación.'; To = 'Tenant, admin, scope, system exclusions, fast path, and authentication.' }
            @{ From = 'Valida PowerShell, PnP, App Registration y acceso tenant.'; To = 'Validates PowerShell, PnP, App Registration, and tenant access.' }
            @{ From = 'Connect-PnPOnline no devolvió una conexión.'; To = 'Connect-PnPOnline did not return a connection.' }
            @{ From = 'Get-PnPTenant no devolvió información del tenant.'; To = 'Get-PnPTenant did not return tenant information.' }
            @{ From = 'Validando aplicación · intento '; To = 'Validating application · attempt ' }
            @{ From = 'La validación aún no pasó. Reintentando en '; To = 'Validation has not passed yet. Retrying in ' }
            @{ From = 'No fue posible ejecutar la validación.'; To = 'The validation could not be executed.' }
            @{ From = 'La aplicación no pasó la validación.'; To = 'The application did not pass validation.' }
            @{ From = 'Usar aplicación existente'; To = 'Use existing application' }
            @{ From = 'La configuración actual no se reemplazará hasta que la nueva aplicación pase la validación.'; To = 'The current configuration will not be replaced until the new application passes validation.' }
            @{ From = 'No se modificó la configuración guardada.'; To = 'The saved configuration was not modified.' }
            @{ From = '¿Mantener la sesión autenticada?'; To = 'Keep the authenticated session?' }
            @{ From = 'Registrar nueva aplicación Entra'; To = 'Register new Entra application' }
            @{ From = 'La aplicación se crea para login interactivo de PnP.PowerShell.'; To = 'The application is created for interactive PnP.PowerShell sign-in.' }
            @{ From = 'La app anterior seguirá activa hasta que la nueva pase una prueba real.'; To = 'The previous app remains active until the new one passes a live test.' }
            @{ From = '¿Registrar una nueva aplicación?'; To = 'Register a new application?' }
            @{ From = 'Nombre de la aplicación'; To = 'Application name' }
            @{ From = 'PnP no devolvió el Client ID en un formato reconocible.'; To = 'PnP did not return the Client ID in a recognizable format.' }
            @{ From = 'No fue posible obtener un Client ID válido para la aplicación creada.'; To = 'Could not obtain a valid Client ID for the created application.' }
            @{ From = 'Nueva aplicación validada'; To = 'New application validated' }
            @{ From = 'La app fue creada en Entra, pero NO reemplazó la configuración activa.'; To = 'The app was created in Entra, but did NOT replace the active configuration.' }
            @{ From = 'Puedes volver a validarla más tarde desde este mismo menú.'; To = 'You can validate it again later from this same menu.' }
            @{ From = 'La nueva aplicación quedó configurada como activa.'; To = 'The new application is now configured as active.' }
            @{ From = 'La configuración anterior se conservó.'; To = 'The previous configuration was preserved.' }
            @{ From = 'La configuración no contiene tenant y Client ID válidos.'; To = 'The configuration does not contain a valid tenant and Client ID.' }
            @{ From = 'Usar una aplicación existente'; To = 'Use an existing application' }
            @{ From = 'Sitio raíz del tenant'; To = 'Tenant root site' }
            @{ From = 'My Site Host raíz'; To = 'Root My Site Host' }
            @{ From = 'SharePoint no devolvió un perfil para '; To = 'SharePoint did not return a profile for ' }
            @{ From = 'El usuario fue localizado, pero PnP no devolvió un SID utilizable. Shape recibido: '; To = 'The user was located, but PnP did not return a usable SID. Received shape: ' }
            @{ From = 'Múltiples entradas legacy coincidentes'; To = 'Multiple matching legacy entries' }
            @{ From = 'Múltiples coincidencias; ninguna mismatch inequívoca'; To = 'Multiple matches; no unambiguous mismatch' }
            @{ From = 'No hay una entrada única y segura para eliminar.'; To = 'There is no single safe entry to remove.' }
            @{ From = 'Fast path: no se abre una conexión previa a cada sitio.'; To = 'Fast path: no connection is opened in advance for each site.' }
            @{ From = 'El ID antiguo todavía resolvió inmediatamente después de Remove-PnPUser.'; To = 'The old ID still resolved immediately after Remove-PnPUser.' }
            @{ From = 'Espera global de propagación: '; To = 'Global propagation wait: ' }
            @{ From = 'Timeout esperando propagación administrativa: '; To = 'Timeout while waiting for administrative propagation: ' }
            @{ From = 'Acceso aún no propagado · intento '; To = 'Access not propagated yet · attempt ' }
            @{ From = 'Hay $($matches.Count) coincidencias. Refina la búsqueda.'; To = '$($matches.Count) matches found. Refine the search.' }
            @{ From = 'Propagación'; To = 'Propagation' }
            @{ From = 'Reglas de exclusión'; To = 'Exclusion rules' }
            @{ From = 'Hosts raíz'; To = 'Root hosts' }
            @{ From = 'Intentos de validación'; To = 'Validation attempts' }
            @{ From = 'Confirmar ejecución'; To = 'Confirm run' }
            @{ From = 'Hosts raíz y sitios de sistema: OMITIDOS antes de cualquier grant.'; To = 'Root hosts and system sites: SKIPPED before any grant.' }
            @{ From = 'Validación previa'; To = 'Pre-validation' }
            @{ From = 'Restauración incompleta'; To = 'Incomplete restoration' }
            @{ From = 'Ejecutando prueba real de autenticación...'; To = 'Running live authentication test...' }
            @{ From = '$(@($script:Settings.ProtectedSpoSites).Count) sitio(s) conservarán tu Site Collection Admin.'; To = '$(@($script:Settings.ProtectedSpoSites).Count) site(s) will retain your Site Collection Admin.' }
            @{ From = 'Activar modo REAL'; To = 'Enable LIVE mode' }
            @{ From = 'El modo LIVE puede eliminar entradas legacy de UserInfoList.'; To = 'LIVE mode can remove legacy UserInfoList entries.' }
            @{ From = 'El modo REAL puede eliminar entradas legacy de UserInfoList.'; To = 'LIVE mode can remove legacy UserInfoList entries.' }
            @{ From = 'La confirmación REPARAR seguirá siendo obligatoria al ejecutar.'; To = 'REPAIR confirmation will still be required when running.' }
            @{ From = 'Modo'; To = 'Mode' }
            @{ From = 'Operación'; To = 'Operation' }
            @{ From = 'Detección ODB'; To = 'ODB detection' }
            @{ From = ' conservarán tu Site Collection Admin.'; To = ' will retain your Site Collection Admin.' }
        )
        foreach ($translation in $phrases) { $result = $result.Replace($translation.From, $translation.To) }
        foreach ($translation in $script:UiTranslations) { $result = $result.Replace($translation.From, $translation.To) }
        $common = @(
            @{ From = 'Selecciona'; To = 'Select' }; @{ From = 'Seleccione'; To = 'Select' }; @{ From = 'Introducir'; To = 'Enter' }; @{ From = 'Introduzca'; To = 'Enter' }
            @{ From = 'Buscar'; To = 'Search' }; @{ From = 'Usar'; To = 'Use' }; @{ From = 'Cambiar'; To = 'Change' }; @{ From = 'Guardar'; To = 'Save' }
            @{ From = 'Validar'; To = 'Validate' }; @{ From = 'Registrar'; To = 'Register' }; @{ From = 'Obtener'; To = 'Get' }; @{ From = 'Restablecer'; To = 'Reset' }
            @{ From = 'Analizar'; To = 'Analyze' }; @{ From = 'Procesar'; To = 'Process' }; @{ From = 'Procesando'; To = 'Processing' }; @{ From = 'Preparando'; To = 'Preparing' }
            @{ From = 'Restaurando'; To = 'Restoring' }; @{ From = 'Aplicación'; To = 'Application' }; @{ From = 'Autenticación'; To = 'Authentication' }
            @{ From = 'Sitios'; To = 'Sites' }; @{ From = 'Sitio'; To = 'Site' }; @{ From = 'Usuario'; To = 'User' }; @{ From = 'Usuarios'; To = 'Users' }
            @{ From = 'afectado'; To = 'affected' }; @{ From = 'actual'; To = 'current' }; @{ From = 'Revisión'; To = 'Review' }; @{ From = 'Error'; To = 'Error' }
            @{ From = 'Estado'; To = 'Status' }; @{ From = 'Resultado'; To = 'Result' }; @{ From = 'Reporte'; To = 'Report' }; @{ From = 'Auditoría'; To = 'Audit' }
            @{ From = 'SIMULACIÓN'; To = 'SIMULATION' }; @{ From = 'REAL'; To = 'LIVE' }; @{ From = 'eliminar'; To = 'remove' }; @{ From = 'Eliminar'; To = 'Remove' }
            @{ From = 'Protección'; To = 'Protection' }; @{ From = 'Protegido'; To = 'Protected' }; @{ From = 'Recuperación'; To = 'Recovery' }
            @{ From = 'No hay'; To = 'There are no' }; @{ From = 'No se encontró'; To = 'No ... was found' }; @{ From = 'No fue posible'; To = 'It was not possible' }
            @{ From = 'Operación cancelada'; To = 'Operation canceled' }; @{ From = 'Límite'; To = 'Limit' }; @{ From = 'Máximo'; To = 'Maximum' }
        )
        foreach ($translation in $common) { $result = $result.Replace($translation.From, $translation.To) }
        $words = @(
            @{ From = 'al'; To = 'when' }; @{ From = 'alguna'; To = 'some' }; @{ From = 'actual'; To = 'current' }; @{ From = 'administrador'; To = 'administrator' }
            @{ From = 'afectado'; To = 'affected' }; @{ From = 'analizar'; To = 'analyze' }; @{ From = 'análisis'; To = 'analysis' }; @{ From = 'agregar'; To = 'add' }
            @{ From = 'buscar'; To = 'search' }; @{ From = 'cambiar'; To = 'change' }; @{ From = 'con'; To = 'with' }; @{ From = 'configuración'; To = 'configuration' }
            @{ From = 'confirmada'; To = 'confirmed' }; @{ From = 'después'; To = 'after' }; @{ From = 'dentro'; To = 'within' }; @{ From = 'eliminar'; To = 'remove' }
            @{ From = 'en'; To = 'in' }; @{ From = 'estado'; To = 'status' }; @{ From = 'exclusiones'; To = 'exclusions' }; @{ From = 'identidad'; To = 'identity' }
            @{ From = 'introducir'; To = 'enter' }; @{ From = 'límite'; To = 'limit' }; @{ From = 'máximo'; To = 'maximum' }; @{ From = 'nueva'; To = 'new' }
            @{ From = 'opción'; To = 'option' }; @{ From = 'para'; To = 'for' }; @{ From = 'procesar'; To = 'process' }; @{ From = 'procesando'; To = 'processing' }
            @{ From = 'protegido'; To = 'protected' }; @{ From = 'protegidos'; To = 'protected' }; @{ From = 'recuperación'; To = 'recovery' }; @{ From = 'reporte'; To = 'report' }
            @{ From = 'reportes'; To = 'reports' }; @{ From = 'resultado'; To = 'result' }; @{ From = 'selecciona'; To = 'select' }; @{ From = 'seleccionar'; To = 'select' }
            @{ From = 'seleccionado'; To = 'selected' }; @{ From = 'seleccionados'; To = 'selected' }; @{ From = 'sitio'; To = 'site' }; @{ From = 'sitios'; To = 'sites' }
            @{ From = 'sistema'; To = 'system' }; @{ From = 'sin'; To = 'without' }; @{ From = 'usuario'; To = 'user' }; @{ From = 'usuarios'; To = 'users' }
            @{ From = 'válido'; To = 'valid' }; @{ From = 'válida'; To = 'valid' }; @{ From = 'y'; To = 'and' }; @{ From = 'u'; To = 'or' }
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
    while ($true) {
        Microsoft.PowerShell.Utility\Write-Host ''; Microsoft.PowerShell.Utility\Write-Host 'Select language / Seleccione idioma:'; Microsoft.PowerShell.Utility\Write-Host '1. English'; Microsoft.PowerShell.Utility\Write-Host '2. Español'
        $choice = (Microsoft.PowerShell.Utility\Read-Host 'Choice / Opción').Trim()
        if ($choice -eq '1') { $script:Language = 'en'; return }; if ($choice -eq '2') { $script:Language = 'es'; return }
        Microsoft.PowerShell.Utility\Write-Host 'Please choose 1 or 2 / Elija 1 o 2.'
    }
}

# =============================================================================
# TUI
# =============================================================================

function Initialize-Terminal {
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

function Get-Ansi {
    param(
        [ValidateSet('Reset','Bold','Dim','Cyan','Blue','Green','Yellow','Red','Magenta','Gray')]
        [string]$Name
    )

    if (-not $script:Ansi) { return '' }

    $esc = [char]27
    switch ($Name) {
        'Reset'   { "$esc[0m" }
        'Bold'    { "$esc[1m" }
        'Dim'     { "$esc[2m" }
        'Cyan'    { "$esc[36m" }
        'Blue'    { "$esc[94m" }
        'Green'   { "$esc[92m" }
        'Yellow'  { "$esc[93m" }
        'Red'     { "$esc[91m" }
        'Magenta' { "$esc[95m" }
        'Gray'    { "$esc[90m" }
    }
}

function Write-Styled {
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal',
        [switch]$NoNewline
    )

    $prefix = ''
    $suffix = ''
    $fallback = 'Gray'

    switch ($Style) {
        'Normal'  { $fallback = 'Gray' }
        'Muted'   { $prefix = Get-Ansi Dim; $fallback = 'DarkGray' }
        'Primary' { $prefix = (Get-Ansi Bold) + (Get-Ansi Cyan); $fallback = 'Cyan' }
        'Success' { $prefix = Get-Ansi Green; $fallback = 'Green' }
        'Warning' { $prefix = Get-Ansi Yellow; $fallback = 'Yellow' }
        'Danger'  { $prefix = Get-Ansi Red; $fallback = 'Red' }
        'Accent'  { $prefix = Get-Ansi Magenta; $fallback = 'Magenta' }
    }

    if ($script:Ansi) {
        $suffix = Get-Ansi Reset
        if ($NoNewline) { Write-Host "$prefix$Text$suffix" -NoNewline }
        else { Write-Host "$prefix$Text$suffix" }
    }
    else {
        if ($NoNewline) { Write-Host $Text -ForegroundColor $fallback -NoNewline }
        else { Write-Host $Text -ForegroundColor $fallback }
    }
}

function Write-AppHeader {
    param([string]$Context = '')

    $width = 72
    try {
        $width = [Math]::Min(84, [Math]::Max(36, $Host.UI.RawUI.WindowSize.Width - 2))
    }
    catch { $width = 72 }

    Clear-Host
    Write-Styled "$script:AppName  v$script:AppVersion" Primary
    if (-not [string]::IsNullOrWhiteSpace($Context)) {
        Write-Styled $Context Muted
    }
    Write-Styled ('━' * $width) Muted
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    Write-Host ''
    Write-Styled $Title Primary
}

function Write-Status {
    param(
        [ValidateSet('Info','Ok','Warn','Error')][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )

    switch ($Level) {
        'Info'  { Write-Styled "● $Message" Primary }
        'Ok'    { Write-Styled "✓ $Message" Success }
        'Warn'  { Write-Styled "! $Message" Warning }
        'Error' { Write-Styled "× $Message" Danger }
    }
}

function Write-Field {
    param(
        [string]$Name,
        [AllowNull()]$Value,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal'
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        $Value = '—'
    }

    Write-Styled ('{0,-18}' -f $Name) Muted -NoNewline
    Write-Styled ([string]$Value) $Style
}

function Pause-Tui {
    Write-Host ''
    [void](Read-Host 'Enter para continuar')
}

function Read-MenuChoice {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Items,
        [string[]]$Description = @(),
        [string]$ZeroLabel = 'Volver'
    )

    while ($true) {
        Write-AppHeader $Title

        foreach ($line in $Description) {
            Write-Styled $line Muted
        }

        if ($Description.Count -gt 0) { Write-Host '' }

        for ($i = 0; $i -lt $Items.Count; $i++) {
            $item = $Items[$i]
            $label = [string]$item.Label
            $hint = ''
            if ($null -ne $item.PSObject.Properties['Hint']) {
                $hint = [string]$item.Hint
            }

            Write-Styled ('  {0,2}  ' -f ($i + 1)) Primary -NoNewline
            Write-Host $label

            if (-not [string]::IsNullOrWhiteSpace($hint)) {
                Write-Styled "      $hint" Muted
            }

            Write-Host ''
        }

        Write-Styled '   0  ' Muted -NoNewline
        Write-Host $ZeroLabel
        Write-Host ''

        $raw = (Read-Host 'Selecciona una opción').Trim()
        $n = 0

        if ([int]::TryParse($raw, [ref]$n)) {
            if ($n -eq 0) { return $null }
            if ($n -ge 1 -and $n -le $Items.Count) {
                return $Items[$n - 1]
            }
        }

        Write-Host ''
        Write-Status Warn 'Opción no válida.'
        Start-Sleep -Milliseconds 700
    }
}

function Read-TextValue {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = '',
        [switch]$Required
    )

    while ($true) {
        $raw = if ([string]::IsNullOrWhiteSpace($Default)) {
            Read-Host $Prompt
        }
        else {
            Read-Host "$Prompt [$Default]"
        }

        if ($null -eq $raw) { $raw = '' }
        $raw = $raw.Trim()

        if ([string]::IsNullOrWhiteSpace($raw)) {
            $raw = $Default
        }

        if (-not $Required -or -not [string]::IsNullOrWhiteSpace($raw)) {
            return $raw
        }

        Write-Status Warn 'Este valor es obligatorio.'
    }
}

function Read-YesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$DefaultYes = $false
    )

    while ($true) {
        $suffix = if ($DefaultYes) { '[S/n]' } else { '[s/N]' }
        $answer = (Read-Host "$Prompt $suffix").Trim()

        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $DefaultYes
        }

        if ($answer -match '^(s|si|sí|y|yes)$') { return $true }
        if ($answer -match '^(n|no)$') { return $false }

        Write-Status Warn 'Responde S o N.'
    }
}

function Read-Integer {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [int]$Minimum,
        [int]$Maximum,
        [int]$Default
    )

    while ($true) {
        $raw = (Read-Host "$Prompt [$Default]").Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return $Default
        }

        $n = 0
        if ([int]::TryParse($raw, [ref]$n) -and $n -ge $Minimum -and $n -le $Maximum) {
            return $n
        }

        Write-Status Warn "Introduce un valor entre $Minimum y $Maximum."
    }
}

# =============================================================================
# CONFIGURACIÓN
# =============================================================================

function New-DefaultSettings {
    [pscustomobject]@{
        Tenant                         = ''
        ClientId                       = ''
        AppRegistrationName            = 'Microsoft 365 Legacy UserId Cleaner'
        PersistLogin                    = $false
        Language                        = 'es'
        AdminUpn                        = ''
        AffectedUserUpn                 = ''
        DryRun                          = $true
        Scope                           = 'Both'
        PropagationTimeoutSeconds       = $script:DefaultPropagationTimeoutSeconds
        PropagationInitialDelaySeconds  = $script:DefaultPropagationInitialDelaySeconds
        GlobalGrantSettleSeconds         = 5
        FastProcessingMode               = $true

        # Exclusiones conservadoras de sitios de sistema. Estas reglas se
        # aplican antes de solicitar Site Collection Admin.
        SkipSystemSites                  = $true
        ExcludedSystemTemplates          = @(
            'APPCATALOG#0',
            'SPSMSITEHOST#0',
            'SRCHCEN#0',
            'SRCHCENTERLITE#0',
            'EDISC#0'
        )
        ExcludedSystemUrlFragments       = @(
            '/sites/appcatalog',
            '/sites/contenttypehub'
        )

        MaxRetries                      = $script:DefaultMaxRetries
        RetryDelaySeconds               = $script:DefaultRetryDelaySeconds
        AppValidationRetries             = 3
        AppValidationDelaySeconds        = 4
        ProtectedSpoSites               = @()
        RemovePreExistingThirdPartyOdbAdmin = $false # Compatibilidad v1.x; v2 aplica política fija para ODB de terceros.
    }
}

function Merge-Settings {
    param([Parameter(Mandatory)]$Saved)

    $default = New-DefaultSettings
    foreach ($prop in $default.PSObject.Properties) {
        if ($null -eq $Saved.PSObject.Properties[$prop.Name]) {
            $Saved | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value
        }
    }

    if ($null -eq $Saved.ProtectedSpoSites) {
        $Saved.ProtectedSpoSites = @()
    }

    return $Saved
}

function Ensure-AppFolders {
    foreach ($path in @($script:ConfigRoot, $script:ReportsFolder)) {
        if (-not (Test-Path -LiteralPath $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Load-Settings {
    Ensure-AppFolders

    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        return New-DefaultSettings
    }

    try {
        $saved = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        return Merge-Settings -Saved $saved
    }
    catch {
        return New-DefaultSettings
    }
}

function Save-Settings {
    Ensure-AppFolders
    $script:Settings | ConvertTo-Json -Depth 10 |
        Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}

function Save-ActiveGrantState {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Records)

    Ensure-AppFolders

    $state = [pscustomobject]@{
        CreatedAt = (Get-Date).ToString('o')
        AdminUpn = $script:Settings.AdminUpn
        Records = @($Records)
    }

    $state | ConvertTo-Json -Depth 10 |
        Set-Content -LiteralPath $script:StatePath -Encoding UTF8
}

function Load-ActiveGrantState {
    if (-not (Test-Path -LiteralPath $script:StatePath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $script:StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Clear-ActiveGrantState {
    if (Test-Path -LiteralPath $script:StatePath) {
        Remove-Item -LiteralPath $script:StatePath -Force
    }
}

# =============================================================================
# PNP.POWERSHELL / AUTENTICACIÓN
# =============================================================================

function Get-InstalledPnPModule {
    Get-Module -ListAvailable -Name PnP.PowerShell |
        Sort-Object Version -Descending |
        Select-Object -First 1
}

function Ensure-PnPModule {
    $module = Get-InstalledPnPModule

    if ($null -eq $module) {
        Write-AppHeader 'PnP.PowerShell'
        Write-Status Warn 'PnP.PowerShell no está instalado.'

        if (-not (Read-YesNo '¿Instalar PnP.PowerShell para CurrentUser?' $true)) {
            return $false
        }

        try {
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            Write-Status Error $_.Exception.Message
            Pause-Tui
            return $false
        }

        $module = Get-InstalledPnPModule
    }

    if ($null -eq $module -or $module.Version -lt $script:MinimumPnPVersion) {
        Write-AppHeader 'PnP.PowerShell'
        Write-Status Error "Se requiere PnP.PowerShell $($script:MinimumPnPVersion) o superior."
        if ($null -ne $module) {
            Write-Field 'Instalada' $module.Version Warning
        }
        Pause-Tui
        return $false
    }

    try {
        Import-Module PnP.PowerShell -MinimumVersion $script:MinimumPnPVersion -ErrorAction Stop
        return $true
    }
    catch {
        Write-AppHeader 'PnP.PowerShell'
        Write-Status Error $_.Exception.Message
        Pause-Tui
        return $false
    }
}

function Test-ClientIdFormat {
    param([string]$ClientId)

    if ([string]::IsNullOrWhiteSpace($ClientId)) { return $false }

    $guid = [guid]::Empty
    return [guid]::TryParse($ClientId.Trim(), [ref]$guid)
}

function Test-TenantFormat {
    param([string]$Tenant)

    if ([string]::IsNullOrWhiteSpace($Tenant)) { return $false }

    return (
        $Tenant.Trim() -match
        '^[a-zA-Z0-9][a-zA-Z0-9.-]*\.onmicrosoft\.com$'
    )
}

function Ensure-TenantConfigured {
    if (Test-TenantFormat $script:Settings.Tenant) {
        return $true
    }

    Write-AppHeader 'Tenant'
    Write-Status Warn 'No hay un tenant válido configurado.'
    Write-Host ''

    $tenant = Read-TextValue 'Tenant (ej. contoso.onmicrosoft.com)' `
        -Default $script:Settings.Tenant -Required

    if (-not (Test-TenantFormat $tenant)) {
        Write-Status Error 'El tenant debe tener formato *.onmicrosoft.com.'
        Pause-Tui
        return $false
    }

    $script:Settings.Tenant = $tenant.Trim().ToLowerInvariant()
    Save-Settings
    return $true
}

function Get-TenantPrefix {
    if (-not (Test-TenantFormat $script:Settings.Tenant)) { return '' }
    return ($script:Settings.Tenant -split '\.')[0]
}

function Get-AdminUrl {
    $prefix = Get-TenantPrefix
    if ([string]::IsNullOrWhiteSpace($prefix)) { return '' }
    return "https://$prefix-admin.sharepoint.com"
}

function Connect-M365SiteWithClientId {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$Tenant = $script:Settings.Tenant,
        [switch]$NoPersist
    )

    if (-not (Test-ClientIdFormat $ClientId)) {
        throw 'Client ID no válido.'
    }

    if (-not (Test-TenantFormat $Tenant)) {
        throw 'Tenant no válido.'
    }

    $params = @{
        Url                = $Url
        ClientId           = $ClientId
        Tenant             = $Tenant
        Interactive        = $true
        ReturnConnection   = $true
        ValidateConnection = $true
        ErrorAction        = 'Stop'
    }

    if (-not $NoPersist -and [bool]$script:Settings.PersistLogin) {
        $params.PersistLogin = $true
    }

    return Connect-PnPOnline @params
}

function Close-PnPConnection {
    param([AllowNull()]$Connection)

    if ($null -eq $Connection) { return }

    try {
        if ($null -ne $Connection.PSObject.Properties['Context'] -and
            $null -ne $Connection.Context) {
            $Connection.Context.Dispose()
        }
    }
    catch {
    }
}

function Test-PnPAppRegistration {
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Tenant
    )

    if (-not (Test-ClientIdFormat $ClientId)) {
        return [pscustomobject]@{
            Success      = $false
            ClientId     = $ClientId
            Tenant       = $Tenant
            AdminUrl     = ''
            TenantAccess = $false
            Message      = 'El Client ID no es un GUID válido.'
        }
    }

    if (-not (Test-TenantFormat $Tenant)) {
        return [pscustomobject]@{
            Success      = $false
            ClientId     = $ClientId
            Tenant       = $Tenant
            AdminUrl     = ''
            TenantAccess = $false
            Message      = 'El tenant no tiene formato *.onmicrosoft.com.'
        }
    }

    $prefix = ($Tenant -split '\.')[0]
    $adminUrl = "https://$prefix-admin.sharepoint.com"
    $conn = $null

    try {
        $conn = Connect-M365SiteWithClientId `
            -Url $adminUrl `
            -ClientId $ClientId `
            -Tenant $Tenant `
            -NoPersist

        if ($null -eq $conn) {
            throw 'Connect-PnPOnline no devolvió una conexión.'
        }

        # La prueba no se limita a obtener un token. Se valida que el operador
        # pueda usar la aplicación contra las APIs tenant de SharePoint.
        $tenantInfo = Get-PnPTenant -Connection $conn -ErrorAction Stop
        if ($null -eq $tenantInfo) {
            throw 'Get-PnPTenant no devolvió información del tenant.'
        }

        # Segunda comprobación: enumeración tenant, que es necesaria para el
        # flujo real de esta herramienta.
        $probe = @(
            Get-PnPTenantSite `
                -IncludeOneDriveSites `
                -Connection $conn `
                -ErrorAction Stop |
                Select-Object -First 1
        )

        return [pscustomobject]@{
            Success      = $true
            ClientId     = $ClientId
            Tenant       = $Tenant
            AdminUrl     = $adminUrl
            TenantAccess = $true
            Message      = 'Autenticación y acceso tenant validados correctamente.'
        }
    }
    catch {
        return [pscustomobject]@{
            Success      = $false
            ClientId     = $ClientId
            Tenant       = $Tenant
            AdminUrl     = $adminUrl
            TenantAccess = $false
            Message      = $_.Exception.Message
        }
    }
    finally {
        Close-PnPConnection $conn
    }
}

function Test-PnPAppRegistrationWithRetry {
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Tenant,
        [int]$Retries = $script:Settings.AppValidationRetries,
        [int]$DelaySeconds = $script:Settings.AppValidationDelaySeconds,
        [switch]$Quiet
    )

    $attempts = [Math]::Max(1, $Retries)
    $last = $null

    for ($attempt = 1; $attempt -le $attempts; $attempt++) {
        if (-not $Quiet) {
            Write-Status Info "Validando aplicación · intento $attempt/$attempts..."
        }

        $last = Test-PnPAppRegistration -ClientId $ClientId -Tenant $Tenant

        if ($last.Success) {
            $last | Add-Member -NotePropertyName Attempts -NotePropertyValue $attempt -Force
            return $last
        }

        if ($attempt -lt $attempts) {
            if (-not $Quiet) {
                Write-Styled "La validación aún no pasó. Reintentando en $DelaySeconds s..." Muted
            }
            Start-Sleep -Seconds ([Math]::Max(1, $DelaySeconds))
        }
    }

    if ($null -eq $last) {
        $last = [pscustomobject]@{
            Success      = $false
            ClientId     = $ClientId
            Tenant       = $Tenant
            AdminUrl     = ''
            TenantAccess = $false
            Message      = 'No fue posible ejecutar la validación.'
        }
    }

    $last | Add-Member -NotePropertyName Attempts -NotePropertyValue $attempts -Force
    return $last
}

function Show-AppValidationResult {
    param(
        [Parameter(Mandatory)]$Result,
        [string]$SuccessTitle = 'Aplicación válida'
    )

    Write-Host ''

    if ($Result.Success) {
        Write-Status Ok $SuccessTitle
        Write-Field 'Tenant' $Result.Tenant
        Write-Field 'Client ID' $Result.ClientId Primary
        Write-Field 'Admin Center' $Result.AdminUrl Muted
        Write-Field 'Acceso tenant' 'VALIDADO' Success
        Write-Field 'Intentos' $Result.Attempts
    }
    else {
        Write-Status Error 'La aplicación no pasó la validación.'
        Write-Field 'Tenant' $Result.Tenant
        Write-Field 'Client ID' $Result.ClientId Warning
        Write-Field 'Admin Center' $Result.AdminUrl Muted
        Write-Field 'Acceso tenant' 'NO VALIDADO' Danger
        Write-Host ''
        Write-Styled $Result.Message Muted
    }
}

function Set-ExistingPnPClientId {
    Write-AppHeader 'Usar aplicación existente'
    Write-Styled 'La configuración actual no se reemplazará hasta que la nueva aplicación pase la validación.' Muted
    Write-Host ''

    $tenant = Read-TextValue 'Tenant' `
        -Default $script:Settings.Tenant -Required

    if (-not (Test-TenantFormat $tenant)) {
        Write-Status Error 'Tenant no válido. Debe tener formato *.onmicrosoft.com.'
        Pause-Tui
        return
    }

    $tenant = $tenant.Trim().ToLowerInvariant()

    $clientId = Read-TextValue 'Client ID' `
        -Default $script:Settings.ClientId -Required

    if (-not (Test-ClientIdFormat $clientId)) {
        Write-Status Error 'Client ID no válido.'
        Pause-Tui
        return
    }

    $clientId = $clientId.Trim()

    Write-Host ''
    $result = Test-PnPAppRegistrationWithRetry `
        -ClientId $clientId `
        -Tenant $tenant

    Show-AppValidationResult -Result $result -SuccessTitle 'Aplicación existente validada'

    if (-not $result.Success) {
        Write-Host ''
        Write-Status Warn 'No se modificó la configuración guardada.'
        Pause-Tui
        return
    }

    $persist = Read-YesNo '¿Mantener la sesión autenticada?' `
        ([bool]$script:Settings.PersistLogin)

    # Commit únicamente después de una validación satisfactoria.
    $script:Settings.Tenant = $tenant
    $script:Settings.ClientId = $clientId
    $script:Settings.PersistLogin = $persist
    Save-Settings

    Write-Host ''
    Write-Status Ok 'Aplicación guardada como configuración activa.'
    Pause-Tui
}

function Get-RegisteredClientIdFromResult {
    param([AllowNull()]$Result)

    if ($null -eq $Result) { return '' }

    foreach ($propertyName in @(
        'AppId',
        'ApplicationId',
        'ClientId',
        'Id'
    )) {
        if ($null -ne $Result.PSObject.Properties[$propertyName]) {
            $candidate = [string]$Result.$propertyName
            if (Test-ClientIdFormat $candidate) {
                return $candidate.Trim()
            }
        }
    }

    return ''
}

function Register-NewPnPApp {
    if (-not (Ensure-TenantConfigured)) { return }

    Write-AppHeader 'Registrar nueva aplicación Entra'
    Write-Styled 'La aplicación se crea para login interactivo de PnP.PowerShell.' Muted
    Write-Host ''

    Write-Field 'Tenant' $script:Settings.Tenant Primary
    Write-Field 'Permiso SharePoint' 'AllSites.FullControl (Delegated)' Warning
    Write-Field 'Permisos Graph' 'No solicitados' Success

    Write-Host ''
    Write-Status Warn 'El registro puede requerir consentimiento de administrador.'
    Write-Styled 'La app anterior seguirá activa hasta que la nueva pase una prueba real.' Muted
    Write-Host ''

    if (-not (Read-YesNo '¿Registrar una nueva aplicación?' $false)) {
        return
    }

    $name = Read-TextValue 'Nombre de la aplicación' `
        -Default $script:Settings.AppRegistrationName `
        -Required

    $oldTenant = $script:Settings.Tenant
    $oldClientId = $script:Settings.ClientId
    $oldName = $script:Settings.AppRegistrationName

    try {
        $params = @{
            ApplicationName               = $name
            Tenant                        = $script:Settings.Tenant
            SharePointDelegatePermissions = @('AllSites.FullControl')
            ErrorAction                   = 'Stop'
        }

        Write-Host ''
        Write-Status Info 'Iniciando registro y consentimiento...'

        $registration = Register-PnPEntraIDAppForInteractiveLogin @params
        $clientId = Get-RegisteredClientIdFromResult -Result $registration

        if (-not (Test-ClientIdFormat $clientId)) {
            Write-Host ''
            Write-Status Warn 'PnP no devolvió el Client ID en un formato reconocible.'
            $clientId = Read-TextValue 'Introduce el Client ID creado' -Required
        }

        if (-not (Test-ClientIdFormat $clientId)) {
            throw 'No fue posible obtener un Client ID válido para la aplicación creada.'
        }

        Write-Host ''
        Write-Status Ok 'Registro completado.'
        Write-Field 'Aplicación' $name
        Write-Field 'Client ID' $clientId Primary
        Write-Host ''

        $validation = Test-PnPAppRegistrationWithRetry `
            -ClientId $clientId `
            -Tenant $script:Settings.Tenant

        Show-AppValidationResult `
            -Result $validation `
            -SuccessTitle 'Nueva aplicación validada'

        if (-not $validation.Success) {
            # No reemplazamos la configuración previa.
            $script:Settings.Tenant = $oldTenant
            $script:Settings.ClientId = $oldClientId
            $script:Settings.AppRegistrationName = $oldName

            Write-Host ''
            Write-Status Warn 'La app fue creada en Entra, pero NO reemplazó la configuración activa.'
            Write-Styled 'Puedes volver a validarla más tarde desde este mismo menú.' Muted
            Pause-Tui
            return
        }

        $script:Settings.ClientId = $clientId
        $script:Settings.AppRegistrationName = $name
        $script:Settings.PersistLogin = Read-YesNo `
            '¿Mantener la sesión autenticada?' `
            ([bool]$script:Settings.PersistLogin)

        Save-Settings

        Write-Host ''
        Write-Status Ok 'La nueva aplicación quedó configurada como activa.'
    }
    catch {
        $script:Settings.Tenant = $oldTenant
        $script:Settings.ClientId = $oldClientId
        $script:Settings.AppRegistrationName = $oldName

        Write-Host ''
        Write-Status Error $_.Exception.Message
        Write-Styled 'La configuración anterior se conservó.' Muted
    }

    Pause-Tui
}

function Get-AppAuthState {
    $configured = (
        (Test-TenantFormat $script:Settings.Tenant) -and
        (Test-ClientIdFormat $script:Settings.ClientId)
    )

    if (-not $configured) {
        return [pscustomobject]@{
            Configured = $false
            Label = 'NO CONFIGURADA'
            Style = 'Warning'
        }
    }

    return [pscustomobject]@{
        Configured = $true
        Label = 'CONFIGURADA'
        Style = 'Success'
    }
}

function Show-AppRegistrationMenu {
    while ($true) {
        $state = Get-AppAuthState
        $items = @()

        # La numeración es siempre dinámica. Al ocultar una opción, las demás
        # se renumeran automáticamente y nunca quedan huecos.
        if ($state.Configured) {
            $items += [pscustomobject]@{
                Label = 'Validar aplicación configurada'
                Value = 'Validate'
                Hint = 'Prueba token, Admin Center, Get-PnPTenant y enumeración de sitios.'
            }
        }

        $items += [pscustomobject]@{
            Label = if ($state.Configured) {
                'Usar otra aplicación existente'
            }
            else {
                'Usar una aplicación existente'
            }
            Value = 'Existing'
            Hint = 'La configuración actual solo cambia después de validar correctamente.'
        }

        $items += [pscustomobject]@{
            Label = 'Registrar una nueva aplicación Entra'
            Value = 'Register'
            Hint = 'Crea y valida una nueva app de PnP con AllSites.FullControl delegado.'
        }

        if ($state.Configured) {
            $items += [pscustomobject]@{
                Label = 'Quitar Client ID de la configuración local'
                Value = 'RemoveLocal'
                Hint = 'No elimina la aplicación en Entra ID.'
            }
        }

        $items += [pscustomobject]@{
            Label = 'Eliminar sesión persistente'
            Value = 'ClearLogin'
            Hint = 'Limpia la autenticación persistida por PnP.PowerShell.'
        }

        $choice = Read-MenuChoice `
            -Title 'Aplicación Entra / autenticación PnP' `
            -Items $items `
            -Description @(
                "Estado: $($state.Label)",
                "Tenant: $(if (Test-TenantFormat $script:Settings.Tenant) { $script:Settings.Tenant } else { 'No configurado' })",
                "Aplicación: $(if ([string]::IsNullOrWhiteSpace($script:Settings.AppRegistrationName)) { '—' } else { $script:Settings.AppRegistrationName })",
                "Client ID: $(if (Test-ClientIdFormat $script:Settings.ClientId) { $script:Settings.ClientId } else { 'No configurado' })"
            )

        if ($null -eq $choice) { return }

        switch ($choice.Value) {
            'Validate' {
                Write-AppHeader 'Validar aplicación configurada'

                if (-not (Test-TenantFormat $script:Settings.Tenant) -or
                    -not (Test-ClientIdFormat $script:Settings.ClientId)) {
                    Write-Status Error 'La configuración no contiene tenant y Client ID válidos.'
                    Pause-Tui
                    continue
                }

                $result = Test-PnPAppRegistrationWithRetry `
                    -ClientId $script:Settings.ClientId `
                    -Tenant $script:Settings.Tenant

                Show-AppValidationResult `
                    -Result $result `
                    -SuccessTitle 'Aplicación configurada validada'

                Pause-Tui
            }

            'Existing' {
                Set-ExistingPnPClientId
            }

            'Register' {
                Register-NewPnPApp
            }

            'RemoveLocal' {
                Write-AppHeader 'Quitar configuración de aplicación'
                Write-Status Warn 'Esto NO elimina la App Registration en Entra ID.'
                Write-Field 'Client ID' $script:Settings.ClientId Warning
                Write-Host ''

                if (Read-YesNo '¿Quitar el Client ID de esta herramienta?' $false) {
                    $script:Settings.ClientId = ''
                    Save-Settings
                    Write-Status Ok 'Client ID local eliminado.'
                }

                Pause-Tui
            }

            'ClearLogin' {
                Write-AppHeader 'Sesión persistente'
                try {
                    Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction Stop
                    Write-Status Ok 'Sesión persistente eliminada.'
                }
                catch {
                    Write-Status Warn $_.Exception.Message
                }
                Pause-Tui
            }
        }
    }
}

function Ensure-ClientId {
    if ((Test-TenantFormat $script:Settings.Tenant) -and
        (Test-ClientIdFormat $script:Settings.ClientId)) {
        return $true
    }

    while ($true) {
        $items = @(
            [pscustomobject]@{
                Label = 'Usar una aplicación existente'
                Value = 'Existing'
                Hint = 'Validar un Client ID ya registrado antes de guardarlo.'
            },
            [pscustomobject]@{
                Label = 'Registrar una nueva aplicación Entra'
                Value = 'Register'
                Hint = 'Crear una app PnP nueva y validarla.'
            }
        )

        $choice = Read-MenuChoice `
            -Title 'Autenticación requerida' `
            -Items $items `
            -Description @(
                'No hay una App Registration válida configurada.',
                'Selecciona cómo quieres configurar la autenticación.'
            )

        if ($null -eq $choice) { return $false }

        switch ($choice.Value) {
            'Existing' { Set-ExistingPnPClientId }
            'Register' { Register-NewPnPApp }
        }

        if ((Test-TenantFormat $script:Settings.Tenant) -and
            (Test-ClientIdFormat $script:Settings.ClientId)) {
            return $true
        }
    }
}

# =============================================================================
# RETRIES / CONEXIONES
# =============================================================================

function Test-TransientException {
    param([Parameter(Mandatory)][System.Exception]$Exception)

    $text = $Exception.ToString()
    return (
        $text -match '429' -or
        $text -match '503' -or
        $text -match 'throttl' -or
        $text -match 'temporar' -or
        $text -match 'timed out' -or
        $text -match 'timeout'
    )
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Operation,
        [Parameter(Mandatory)][string]$OperationName,
        [int]$MaxRetries = $script:Settings.MaxRetries,
        [int]$BaseDelaySeconds = $script:Settings.RetryDelaySeconds
    )

    for ($attempt = 1; $attempt -le ($MaxRetries + 1); $attempt++) {
        try {
            return & $Operation
        }
        catch {
            if ($attempt -gt $MaxRetries -or -not (Test-TransientException $_.Exception)) {
                throw
            }

            $delay = [Math]::Min(60, [Math]::Max(1, $BaseDelaySeconds) * [Math]::Pow(2, $attempt - 1))
            $delay = [int][Math]::Ceiling($delay + (Get-Random -Minimum 0 -Maximum 3))

            Write-Status Warn "${OperationName}: reintento $attempt/$MaxRetries en $delay s."
            Start-Sleep -Seconds $delay
        }
    }
}

function Connect-AdminCenter {
    if (-not (Ensure-ClientId)) {
        throw 'No hay Client ID configurado.'
    }

    if (-not (Ensure-TenantConfigured)) {
        throw 'No hay tenant configurado.'
    }

    $url = Get-AdminUrl
    $script:AdminConnection = Connect-M365SiteWithClientId -Url $url -ClientId $script:Settings.ClientId
    return $script:AdminConnection
}

function Disconnect-Safe {
    param([AllowNull()]$Connection)
    Close-PnPConnection $Connection
}

function Connect-Site {
    param([Parameter(Mandatory)][string]$Url)

    return Connect-M365SiteWithClientId -Url $Url -ClientId $script:Settings.ClientId
}

# =============================================================================
# INVENTARIO DE SITIOS
# =============================================================================

function Get-TenantRootUrls {
    if (-not (Test-TenantFormat $script:Settings.Tenant)) {
        return [pscustomobject]@{
            SharePointRoot = ''
            MySiteHostRoot = ''
            AdminRoot      = ''
        }
    }

    $prefix = Get-TenantPrefix

    return [pscustomobject]@{
        SharePointRoot = "https://$prefix.sharepoint.com"
        MySiteHostRoot = "https://$prefix-my.sharepoint.com"
        AdminRoot      = "https://$prefix-admin.sharepoint.com"
    }
}

function Test-IsSystemSite {
    param(
        [Parameter(Mandatory)]$Site
    )

    $url = ([string]$Site.Url).TrimEnd('/')
    $template = [string]$Site.Template
    $roots = Get-TenantRootUrls

    # Admin Center
    if (-not [string]::IsNullOrWhiteSpace($roots.AdminRoot) -and
        $url.Equals($roots.AdminRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{
            IsSystem = $true
            Reason = 'SharePoint Admin Center'
        }
    }

    # SharePoint tenant root. El usuario pidió no procesar hosts raíz.
    if (-not [string]::IsNullOrWhiteSpace($roots.SharePointRoot) -and
        $url.Equals($roots.SharePointRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{
            IsSystem = $true
            Reason = 'Sitio raíz del tenant'
        }
    }

    # My Site Host raíz: tenant-my.sharepoint.com. Este NO es un OneDrive
    # personal; los ODB reales están bajo /personal/.
    if (-not [string]::IsNullOrWhiteSpace($roots.MySiteHostRoot) -and
        $url.Equals($roots.MySiteHostRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{
            IsSystem = $true
            Reason = 'My Site Host raíz'
        }
    }

    # Templates de sistema explícitos y conservadores.
    foreach ($excludedTemplate in @($script:Settings.ExcludedSystemTemplates)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$excludedTemplate) -and
            $template.Equals(
                [string]$excludedTemplate,
                [System.StringComparison]::OrdinalIgnoreCase
            )) {
            return [pscustomobject]@{
                IsSystem = $true
                Reason = "Template de sistema: $template"
            }
        }
    }

    # Rutas de sistema conocidas. Se compara contra fragmentos configurables.
    foreach ($fragment in @($script:Settings.ExcludedSystemUrlFragments)) {
        if ([string]::IsNullOrWhiteSpace([string]$fragment)) { continue }

        if ($url.IndexOf(
            [string]$fragment,
            [System.StringComparison]::OrdinalIgnoreCase
        ) -ge 0) {
            return [pscustomobject]@{
                IsSystem = $true
                Reason = "Ruta de sistema: $fragment"
            }
        }
    }

    # Redirects tampoco se procesan.
    if ($template -like '*Redirect*') {
        return [pscustomobject]@{
            IsSystem = $true
            Reason = "Template redirect: $template"
        }
    }

    return [pscustomobject]@{
        IsSystem = $false
        Reason = ''
    }
}

function Get-TenantSites {
    param(
        [ValidateSet('SharePoint','OneDrive','Both')]
        [string]$Scope = 'Both'
    )

    if ($null -eq $script:AdminConnection) {
        $null = Connect-AdminCenter
    }

    $all = @(
        Invoke-WithRetry -Operation {
            Get-PnPTenantSite `
                -IncludeOneDriveSites `
                -Detailed `
                -Connection $script:AdminConnection `
                -ErrorAction Stop
        } -OperationName 'Enumerar sitios'
    )

    $filtered = @(
        $all | Where-Object {
            $url = [string]$_.Url
            if ([string]::IsNullOrWhiteSpace($url)) { return $false }

            $isOdb = $url -match '-my\.sharepoint\.com/personal/'

            switch ($Scope) {
                'SharePoint' { return -not $isOdb }
                'OneDrive'   { return $isOdb }
                'Both'       { return $true }
            }
        } | Sort-Object Url
    )

    return $filtered
}

function Get-SiteKind {
    param([string]$Url)
    if ($Url -match '-my\.sharepoint\.com/personal/') { return 'OneDrive' }
    return 'SharePoint'
}

function Test-SameUpn {
    param([string]$A, [string]$B)
    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return $false }
    return $A.Trim().Equals($B.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-IsAffectedUsersOwnOneDrive {
    param(
        [Parameter(Mandatory)]$Site,
        [Parameter(Mandatory)]$Affected
    )

    if ((Get-SiteKind $Site.Url) -ne 'OneDrive') { return $false }

    if (Test-SameUpn ([string]$Site.Owner) ([string]$Affected.Upn)) {
        return $true
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Affected.PersonalUrl)) {
        if ([string]$Site.Url -eq ([string]$Affected.PersonalUrl).TrimEnd('/')) {
            return $true
        }
    }

    return $false
}

function Test-IsAdminOwnOneDrive {
    param([Parameter(Mandatory)]$Site)

    if ((Get-SiteKind $Site.Url) -ne 'OneDrive') { return $false }
    return (Test-SameUpn ([string]$Site.Owner) ([string]$script:Settings.AdminUpn))
}

function Test-IsProtectedSpoSite {
    param([Parameter(Mandatory)][string]$Url)

    foreach ($protected in @($script:Settings.ProtectedSpoSites)) {
        if ([string]$protected -eq $Url) { return $true }
    }

    return $false
}

# =============================================================================
# USUARIO AFECTADO / SID ACTUAL
# =============================================================================

function Get-NormalizedSharePointSid {
    param([AllowNull()][string]$Sid)

    if ([string]::IsNullOrWhiteSpace($Sid)) { return '' }

    $value = $Sid.Trim()
    $value = $value -replace '^i:0h\.f\|membership\|', ''
    $value = $value -replace '@live\.com$', ''
    $value = $value.Trim('|')
    return $value
}

function Get-SafeObjectProperty {
    param(
        [AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return $null }

    # IDictionary / Hashtable
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            if ([string]$key -eq $Name) {
                return $Object[$key]
            }
        }
    }

    # Propiedad normal de PSObject/objeto .NET
    try {
        $property = $Object.PSObject.Properties[$Name]
        if ($null -ne $property) {
            return $property.Value
        }
    }
    catch {
    }

    return $null
}

function Get-ProfilePropertyValue {
    param(
        [AllowNull()]$ProfileResult,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $ProfileResult) { return $null }

    # Get-PnPUserProfileProperty puede producir una colección cuando -Account
    # es String[] o variar el shape de salida entre versiones. Normalizamos.
    $objects = @($ProfileResult)

    foreach ($obj in $objects) {
        if ($null -eq $obj) { continue }

        # 1) Propiedad directa: .SID / .PersonalUrl / .AccountName
        $direct = Get-SafeObjectProperty -Object $obj -Name $Name
        if ($null -ne $direct -and
            -not [string]::IsNullOrWhiteSpace([string]$direct)) {
            return $direct
        }

        # 2) Shape clásico: .UserProfileProperties["SID"]
        $bag = Get-SafeObjectProperty -Object $obj -Name 'UserProfileProperties'
        if ($null -ne $bag) {
            if ($bag -is [System.Collections.IDictionary]) {
                foreach ($key in $bag.Keys) {
                    if ([string]$key -eq $Name) {
                        $value = $bag[$key]
                        if ($null -ne $value -and
                            -not [string]::IsNullOrWhiteSpace([string]$value)) {
                            return $value
                        }
                    }
                }
            }
            else {
                $value = Get-SafeObjectProperty -Object $bag -Name $Name
                if ($null -ne $value -and
                    -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    return $value
                }
            }
        }

        # 3) Algunos shapes representan propiedades como Key/Value,
        #    Name/Value o PropertyName/Value.
        foreach ($nameField in @('Key','Name','PropertyName')) {
            $entryName = Get-SafeObjectProperty -Object $obj -Name $nameField

            if ($null -ne $entryName -and
                ([string]$entryName).Equals(
                    $Name,
                    [System.StringComparison]::OrdinalIgnoreCase
                )) {

                foreach ($valueField in @('Value','PropertyValue')) {
                    $entryValue = Get-SafeObjectProperty -Object $obj -Name $valueField
                    if ($null -ne $entryValue -and
                        -not [string]::IsNullOrWhiteSpace([string]$entryValue)) {
                        return $entryValue
                    }
                }
            }
        }
    }

    return $null
}

function Get-AffectedUserOneDriveLocation {
    param(
        [Parameter(Mandatory)][string]$Upn,
        [Parameter(Mandatory)]$Connection
    )

    $result = [pscustomobject]@{
        Found = $false
        Url = ''
        SiteId = ''
        Location = ''
        Method = ''
        Message = ''
    }

    # Método preferido en PnP.PowerShell moderno.
    $cmd = Get-Command Get-PnPUserOneDriveLocation -ErrorAction SilentlyContinue

    if ($null -ne $cmd) {
        try {
            $odb = Invoke-WithRetry -Operation {
                Get-PnPUserOneDriveLocation `
                    -UserPrincipalName $Upn `
                    -Connection $Connection `
                    -ErrorAction Stop
            } -OperationName "Localizar OneDrive de $Upn"

            if ($null -ne $odb) {
                $url = [string](Get-SafeObjectProperty -Object $odb -Name 'MySiteUrl')

                if (-not [string]::IsNullOrWhiteSpace($url)) {
                    $result.Found = $true
                    $result.Url = $url.TrimEnd('/')
                    $result.SiteId = [string](Get-SafeObjectProperty -Object $odb -Name 'SiteId')
                    $result.Location = [string](Get-SafeObjectProperty -Object $odb -Name 'Location')
                    $result.Method = 'Get-PnPUserOneDriveLocation'
                    return $result
                }
            }
        }
        catch {
            $result.Message = $_.Exception.Message
        }
    }

    return $result
}

function Resolve-AffectedUser {
    param([Parameter(Mandatory)][string]$Upn)

    if ($null -eq $script:AdminConnection) {
        $null = Connect-AdminCenter
    }

    # Get-PnPUserProfileProperty has changed output shape across PnP/CSOM
    # combinations. We request the complete profile and normalize it instead
    # of assuming .UserProfileProperties always exists.
    $profile = Invoke-WithRetry -Operation {
        Get-PnPUserProfileProperty `
            -Account $Upn `
            -Connection $script:AdminConnection `
            -ErrorAction Stop
    } -OperationName "Resolver perfil de $Upn"

    if ($null -eq $profile) {
        throw "SharePoint no devolvió un perfil para '$Upn'."
    }

    $sidRaw = Get-ProfilePropertyValue -ProfileResult $profile -Name 'SID'
    $sid = Get-NormalizedSharePointSid ([string]$sidRaw)

    if ([string]::IsNullOrWhiteSpace($sid)) {
        $shape = @(
            @($profile) |
                ForEach-Object {
                    if ($null -ne $_) {
                        $_.PSObject.Properties.Name -join ', '
                    }
                }
        ) -join ' | '

        throw "El usuario fue localizado, pero PnP no devolvió un SID utilizable. Shape recibido: $shape"
    }

    $accountName = [string](Get-ProfilePropertyValue -ProfileResult $profile -Name 'AccountName')
    if ([string]::IsNullOrWhiteSpace($accountName)) {
        $accountName = "i:0#.f|membership|$Upn"
    }

    $personalUrl = [string](Get-ProfilePropertyValue -ProfileResult $profile -Name 'PersonalUrl')
    $oneDriveSource = ''

    if (-not [string]::IsNullOrWhiteSpace($personalUrl)) {
        $personalUrl = $personalUrl.TrimEnd('/')
        $oneDriveSource = 'User Profile'
    }

    # El cmdlet dedicado es la fuente preferida para confirmar el OneDrive.
    $odb = Get-AffectedUserOneDriveLocation `
        -Upn $Upn `
        -Connection $script:AdminConnection

    if ($odb.Found) {
        $personalUrl = $odb.Url
        $oneDriveSource = $odb.Method
    }

    return [pscustomobject]@{
        Upn                   = $Upn
        AccountName           = $accountName
        CurrentSid            = $sid
        PersonalUrl           = $personalUrl
        OneDriveFound         = (-not [string]::IsNullOrWhiteSpace($personalUrl))
        OneDriveSource        = $oneDriveSource
        OneDriveSiteId        = $odb.SiteId
        OneDriveLocation      = $odb.Location
        OneDriveLookupMessage = $odb.Message
    }
}


function Select-AffectedUser {
    if (-not (Ensure-PnPModule)) { return }
    if (-not (Ensure-ClientId)) { return }

    try {
        Write-AppHeader 'Usuario afectado'
        $upn = Read-TextValue 'UPN del usuario recreado' -Default $script:Settings.AffectedUserUpn -Required

        if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            Write-Status Error 'UPN no válido.'
            Pause-Tui
            return
        }

        Write-Host ''
        Write-Status Info 'Consultando el perfil de SharePoint...'

        $null = Connect-AdminCenter
        $resolved = Resolve-AffectedUser -Upn $upn
        $script:AffectedUser = $resolved

        $script:Settings.AffectedUserUpn = $upn
        Save-Settings

        Write-Status Ok 'Identidad actual validada.'
        Write-Field 'UPN' $resolved.Upn Primary
        Write-Field 'AccountName' $resolved.AccountName
        Write-Field 'SID actual' $resolved.CurrentSid Success

        if ($resolved.OneDriveFound) {
            Write-Field 'OneDrive' 'ENCONTRADO' Success
            Write-Field 'OneDrive propio' $resolved.PersonalUrl Primary
            Write-Field 'Detección ODB' $resolved.OneDriveSource Muted

            if (-not [string]::IsNullOrWhiteSpace([string]$resolved.OneDriveLocation)) {
                Write-Field 'Geo ODB' $resolved.OneDriveLocation Muted
            }
        }
        else {
            Write-Field 'OneDrive' 'NO CONFIRMADO' Warning

            if (-not [string]::IsNullOrWhiteSpace([string]$resolved.OneDriveLookupMessage)) {
                Write-Styled $resolved.OneDriveLookupMessage Muted
            }
        }
    }
    catch {
        Write-Status Error $_.Exception.Message
    }
    finally {
        Disconnect-Safe $script:AdminConnection
        $script:AdminConnection = $null
    }

    Pause-Tui
}

# =============================================================================
# OPERADOR / POLÍTICA DE ADMIN
# =============================================================================

function Ensure-AdminUpn {
    if (-not [string]::IsNullOrWhiteSpace($script:Settings.AdminUpn)) {
        return $true
    }

    Write-AppHeader 'Cuenta administradora'
    $upn = Read-TextValue 'UPN del administrador que ejecuta la herramienta' -Required

    if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        Write-Status Error 'UPN no válido.'
        Pause-Tui
        return $false
    }

    $script:Settings.AdminUpn = $upn
    Save-Settings
    return $true
}

# =============================================================================
# USERINFOLIST / DETECCIÓN ID MISMATCH
# =============================================================================

function Find-AffectedUserInSite {
    param(
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)]$Affected
    )

    $users = @()

    # Primero imitamos el patrón de la muestra oficial: resolver por AccountName.
    try {
        $direct = Get-PnPUser -Identity $Affected.AccountName -Connection $Connection -ErrorAction Stop
        if ($null -ne $direct) { $users += $direct }
    }
    catch {}

    # Si no fue suficiente, revisamos usuarios del sitio por UPN/email.
    if ($users.Count -eq 0) {
        try {
            $users = @(
                Get-PnPUser -Connection $Connection -ErrorAction Stop |
                    Where-Object {
                        (Test-SameUpn ([string]$_.Email) $Affected.Upn) -or
                        (Test-SameUpn ([string]$_.LoginName) $Affected.AccountName) -or
                        ([string]$_.LoginName).EndsWith("|$($Affected.Upn)", [System.StringComparison]::OrdinalIgnoreCase)
                    }
            )
        }
        catch {
            throw
        }
    }

    if ($users.Count -eq 0) {
        return [pscustomobject]@{
            Found = $false
            Ambiguous = $false
            User = $null
            StoredSid = ''
            IsMismatch = $false
            Reason = 'Usuario no encontrado'
        }
    }

    $evaluated = @(
        foreach ($u in $users) {
            $storedSid = ''
            try {
                if ($null -ne $u.UserId -and $null -ne $u.UserId.NameId) {
                    $storedSid = Get-NormalizedSharePointSid ([string]$u.UserId.NameId)
                }
            }
            catch {}

            [pscustomobject]@{
                User = $u
                StoredSid = $storedSid
                IsMismatch = (
                    -not [string]::IsNullOrWhiteSpace($storedSid) -and
                    -not $storedSid.Equals($Affected.CurrentSid, [System.StringComparison]::OrdinalIgnoreCase)
                )
            }
        }
    )

    $mismatches = @($evaluated | Where-Object IsMismatch)

    if ($mismatches.Count -eq 1) {
        return [pscustomobject]@{
            Found = $true
            Ambiguous = $false
            User = $mismatches[0].User
            StoredSid = $mismatches[0].StoredSid
            IsMismatch = $true
            Reason = 'ID mismatch confirmado'
        }
    }

    if ($mismatches.Count -gt 1) {
        return [pscustomobject]@{
            Found = $true
            Ambiguous = $true
            User = $null
            StoredSid = ($mismatches.StoredSid -join '; ')
            IsMismatch = $true
            Reason = 'Múltiples entradas legacy coincidentes'
        }
    }

    # Si hay una única entrada con SID vacío, no es seguro decidir.
    if ($evaluated.Count -eq 1 -and [string]::IsNullOrWhiteSpace($evaluated[0].StoredSid)) {
        return [pscustomobject]@{
            Found = $true
            Ambiguous = $true
            User = $evaluated[0].User
            StoredSid = ''
            IsMismatch = $false
            Reason = 'La entrada no expone UserId.NameId'
        }
    }

    return [pscustomobject]@{
        Found = $true
        Ambiguous = ($evaluated.Count -gt 1)
        User = if ($evaluated.Count -eq 1) { $evaluated[0].User } else { $null }
        StoredSid = ($evaluated.StoredSid -join '; ')
        IsMismatch = $false
        Reason = if ($evaluated.Count -gt 1) { 'Múltiples coincidencias; ninguna mismatch inequívoca' } else { 'Identidad correcta' }
    }
}

function Remove-LegacyUserFromSite {
    param(
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)]$Match
    )

    $user = $Match.User
    if ($null -eq $user) {
        throw 'No hay una entrada única y segura para eliminar.'
    }

    if ([bool]$user.IsSiteAdmin) {
        Remove-PnPSiteCollectionAdmin -Owners @([string]$user.LoginName) -Connection $Connection -ErrorAction Stop
    }

    Remove-PnPUser -Identity ([int]$user.Id) -Force -Connection $Connection -ErrorAction Stop
}

# =============================================================================
# REPORTES
# =============================================================================

function New-ReportPath {
    param(
        [Parameter(Mandatory)][string]$Prefix,
        [string]$Extension = 'csv'
    )

    Ensure-AppFolders
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    return Join-Path $script:ReportsFolder "$Prefix-$stamp.$Extension"
}

function Export-RecordsCsv {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Records,
        [Parameter(Mandatory)][string]$Path
    )

    $rows = @($Records)
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        '' | Set-Content -LiteralPath $Path -Encoding UTF8
    }
}

function Show-RunSummary {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Results,
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Inventory,
        [Parameter(Mandatory)][string]$ReportPath
    )

    Write-AppHeader 'Resultado'

    $processed = @($Results).Count
    $found = @($Results | Where-Object UserFound).Count
    $mismatch = @($Results | Where-Object Mismatch).Count
    $removed = @($Results | Where-Object Action -eq 'Eliminado de UserInfoList').Count
    $simulated = @($Results | Where-Object Action -eq 'SIMULACIÓN: eliminar').Count
    $review = @($Results | Where-Object Status -eq 'Revisión').Count
    $skippedOwn = @($Inventory | Where-Object SkipReason -eq 'OneDrive propio del usuario afectado').Count
    $skippedSystem = @($Inventory | Where-Object IsSystemSite).Count
    $errors = @($Results | Where-Object Status -eq 'Error').Count

    $granted = @($Inventory | Where-Object AdminGrantRequested).Count
    $cleaned = @($Inventory | Where-Object AdminCleanupStatus -eq 'Retirado').Count
    $protected = @($Inventory | Where-Object ProtectedSpo).Count
    $cleanupErrors = @($Inventory | Where-Object AdminCleanupStatus -eq 'Error').Count

    $attemptsTotal = 0
    foreach ($row in @($Results)) {
        $attemptsTotal += [int]$row.Attempts
    }

    Write-Section 'Identidad'
    Write-Field 'Usuario' $script:AffectedUser.Upn Primary
    Write-Field 'SID actual' $script:AffectedUser.CurrentSid
    Write-Field 'Modo' `
        $(if ($script:Settings.DryRun) { 'SIMULACIÓN' } else { 'REAL' }) `
        $(if ($script:Settings.DryRun) { 'Warning' } else { 'Danger' })

    Write-Section 'UserInfoList'
    Write-Field 'Sitios procesados' $processed
    Write-Field 'Conexiones/intentos' $attemptsTotal
    Write-Field 'Usuario encontrado' $found
    Write-Field 'Mismatch' $mismatch $(if ($mismatch -gt 0) { 'Warning' } else { 'Success' })
    Write-Field 'Eliminados' $removed $(if ($removed -gt 0) { 'Danger' } else { 'Normal' })
    Write-Field 'Simulados' $simulated
    Write-Field 'Revisión manual' $review $(if ($review -gt 0) { 'Warning' } else { 'Normal' })
    Write-Field 'Own OneDrive omitido' $skippedOwn
    Write-Field 'Sitios sistema omitidos' $skippedSystem
    Write-Field 'Errores' $errors $(if ($errors -gt 0) { 'Danger' } else { 'Success' })

    Write-Section 'Site Collection Admin'
    Write-Field 'Grants solicitados' $granted
    Write-Field 'Retirados' $cleaned
    Write-Field 'SPO protegidos' $protected
    Write-Field 'Errores cleanup' $cleanupErrors $(if ($cleanupErrors -gt 0) { 'Danger' } else { 'Success' })

    Write-Host ''
    Write-Styled 'Reporte de UserInfoList:' Muted
    Write-Styled $ReportPath Primary
}

# =============================================================================
# PIPELINE RÁPIDO: GRANTS → ESPERA GLOBAL → PROCESAR → RESTAURAR
# =============================================================================

function New-SiteRecord {
    param([Parameter(Mandatory)]$Site)

    $kind = Get-SiteKind $Site.Url
    $protected = ($kind -eq 'SharePoint' -and (Test-IsProtectedSpoSite $Site.Url))

    [pscustomobject]@{
        Url                 = [string]$Site.Url
        Title               = [string]$Site.Title
        Owner               = [string]$Site.Owner
        Template            = [string]$Site.Template
        Kind                = $kind

        Skip                = $false
        SkipReason          = ''
        IsSystemSite        = $false
        SystemSiteReason    = ''

        ProtectedSpo        = $protected
        AdminOwnOneDrive    = (Test-IsAdminOwnOneDrive $Site)

        # v2.1 prioriza velocidad. La política de preservación de SPO se basa
        # en la lista explícita de sitios protegidos, no en una conexión previa
        # a cada sitio para descubrir el baseline.
        BaselineKnown       = $false
        PreExistingAdmin    = $false
        BaselineMessage     = 'No consultado en FastProcessingMode'

        AdminGrantRequested = $false
        AdminAddedByTool    = $false
        AdminGrantStatus    = ''
        RemoveAdminAfterRun = $false
        KeepAdminAfterRun   = $false
        KeepAdminReason     = ''

        PropagationStatus   = ''
        PropagationMessage  = ''
        ProcessingStatus    = ''
        ProcessingMessage   = ''

        AdminCleanupStatus  = ''
        AdminCleanupMessage = ''
    }
}

function Request-SiteAdminGrant {
    param(
        [Parameter(Mandatory)]$Site,
        [Parameter(Mandatory)]$Record
    )

    Invoke-WithRetry -Operation {
        Set-PnPTenantSite `
            -Identity $Site.Url `
            -Owners @($script:Settings.AdminUpn) `
            -Connection $script:AdminConnection `
            -ErrorAction Stop
    } -OperationName "Agregar Site Collection Admin en $($Site.Url)" | Out-Null

    $Record.AdminGrantRequested = $true
    $Record.AdminAddedByTool = $true
    $Record.AdminGrantStatus = 'Solicitado'
}

function Invoke-PrepareAdminAccess {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('SharePoint','OneDrive','Both')]
        [string]$Scope
    )

    Write-AppHeader 'Fase 1/3 · Preparar acceso administrativo'
    Write-Styled 'Fast path: no se abre una conexión previa a cada sitio.' Muted
    Write-Styled 'Se completa primero toda la ronda de Site Collection Admin desde Admin Center.' Muted
    Write-Host ''

    Write-Status Info 'Enumerando SharePoint y OneDrive...'
    $sites = @(Get-TenantSites -Scope $Scope)
    Write-Status Ok "$($sites.Count) sitio(s) candidato(s)."
    Write-Host ''

    $inventoryList = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $sites.Count; $i++) {
        $site = $sites[$i]
        $record = New-SiteRecord -Site $site

        $pct = [int]((($i + 1) / [Math]::Max(1, $sites.Count)) * 100)
        Write-Progress `
            -Id 10 `
            -Activity 'Fase 1/3 · Preparando acceso' `
            -Status "$($i + 1)/$($sites.Count) $($site.Url)" `
            -PercentComplete $pct

        # Excluir hosts y colecciones de sistema ANTES de solicitar Site Admin.
        if ([bool]$script:Settings.SkipSystemSites) {
            $systemCheck = Test-IsSystemSite -Site $site

            if ($systemCheck.IsSystem) {
                $record.Skip = $true
                $record.SkipReason = 'Sitio de sistema'
                $record.IsSystemSite = $true
                $record.SystemSiteReason = [string]$systemCheck.Reason
                $record.KeepAdminAfterRun = $true
                $record.KeepAdminReason = 'No se modifica un sitio de sistema'

                [void]$inventoryList.Add($record)
                Save-ActiveGrantState -Records $inventoryList.ToArray()
                continue
            }
        }

        # OneDrive propio del usuario afectado: nunca tocar.
        if (Test-IsAffectedUsersOwnOneDrive -Site $site -Affected $script:AffectedUser) {
            $record.Skip = $true
            $record.SkipReason = 'OneDrive propio del usuario afectado'
            $record.KeepAdminAfterRun = $true
            $record.KeepAdminReason = 'Protección absoluta'
            [void]$inventoryList.Add($record)
            Save-ActiveGrantState -Records $inventoryList.ToArray()
            continue
        }

        try {
            if ($record.Kind -eq 'OneDrive' -and $record.AdminOwnOneDrive) {
                # Tu propio OneDrive se procesa sin cambiar tu rol.
                $record.AdminGrantStatus = 'No requerido'
                $record.KeepAdminAfterRun = $true
                $record.KeepAdminReason = 'OneDrive propio del administrador'
            }
            else {
                # SPO protegido: grant si hace falta, pero conservar al final.
                # SPO no protegido: grant temporal y retirar al final.
                # ODB de terceros: grant temporal y retirar siempre al final.
                Request-SiteAdminGrant -Site $site -Record $record

                if ($record.Kind -eq 'SharePoint' -and $record.ProtectedSpo) {
                    $record.KeepAdminAfterRun = $true
                    $record.KeepAdminReason = 'Sitio SPO protegido'
                }
                else {
                    $record.RemoveAdminAfterRun = $true
                }
            }
        }
        catch {
            $record.AdminGrantStatus = 'Error'
            $record.ProcessingStatus = 'Error'
            $record.ProcessingMessage = $_.Exception.Message
        }

        [void]$inventoryList.Add($record)
        Save-ActiveGrantState -Records $inventoryList.ToArray()
    }

    Write-Progress -Id 10 -Activity 'Fase 1/3 · Preparando acceso' -Completed

    $inventory = @($inventoryList.ToArray())
    Save-ActiveGrantState -Records $inventory
    return $inventory
}

function Test-AccessOrPropagationError {
    param([Parameter(Mandatory)][System.Exception]$Exception)

    $text = $Exception.ToString()

    return (
        $text -match '(?i)\b401\b' -or
        $text -match '(?i)\b403\b' -or
        $text -match '(?i)forbidden' -or
        $text -match '(?i)access denied' -or
        $text -match '(?i)unauthori[sz]ed' -or
        $text -match '(?i)attempted to perform an unauthorized operation'
    )
}

function Invoke-SiteProcessingAttempt {
    param(
        [Parameter(Mandatory)]$Record,
        [Parameter(Mandatory)]$Result
    )

    $conn = $null

    try {
        $conn = Connect-Site -Url $Record.Url

        # No hacemos una consulta separada de IsSiteAdmin. La propia operación
        # de UserInfoList es la prueba de acceso real y ahorra una llamada por sitio.
        $match = Find-AffectedUserInSite `
            -Connection $conn `
            -Affected $script:AffectedUser

        $Result.UserFound = [bool]$match.Found
        $Result.Mismatch = [bool]$match.IsMismatch
        $Result.Ambiguous = [bool]$match.Ambiguous
        $Result.StoredSid = [string]$match.StoredSid
        $Result.Message = [string]$match.Reason

        if ($null -ne $match.User) {
            $Result.SharePointUserId = [string]$match.User.Id
            $Result.LoginName = [string]$match.User.LoginName
            $Result.WasAffectedUserSiteAdmin = [bool]$match.User.IsSiteAdmin
        }

        if (-not $match.Found) {
            $Result.Action = 'No encontrado'
        }
        elseif ($match.Ambiguous) {
            $Result.Status = 'Revisión'
            $Result.Action = 'Omitido por ambigüedad'
        }
        elseif (-not $match.IsMismatch) {
            $Result.Action = 'Identidad correcta'
        }
        elseif ($script:Settings.DryRun) {
            $Result.Action = 'SIMULACIÓN: eliminar'
        }
        else {
            Remove-LegacyUserFromSite `
                -Connection $conn `
                -Match $match

            $Result.Action = 'Eliminado de UserInfoList'

            try {
                $verify = Get-PnPUser `
                    -Identity ([int]$match.User.Id) `
                    -Connection $conn `
                    -ErrorAction Stop

                if ($null -ne $verify) {
                    $Result.Status = 'Revisión'
                    $Result.Message = 'El ID antiguo todavía resolvió inmediatamente después de Remove-PnPUser.'
                }
            }
            catch {
                $Result.Message = 'ID mismatch eliminado y verificado.'
            }
        }

        return [pscustomobject]@{
            Success = $true
            AccessIssue = $false
            Message = ''
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false
            AccessIssue = (Test-AccessOrPropagationError $_.Exception)
            Message = $_.Exception.Message
        }
    }
    finally {
        Disconnect-Safe $conn
    }
}

function New-RepairResult {
    param([Parameter(Mandatory)]$Record)

    [pscustomobject]@{
        FechaHora                = Get-Date
        SiteUrl                  = $Record.Url
        SiteTitle                = $Record.Title
        Tipo                     = $Record.Kind
        Owner                    = $Record.Owner
        Usuario                  = $script:AffectedUser.Upn
        CurrentSid               = $script:AffectedUser.CurrentSid
        StoredSid                = ''
        UserFound                = $false
        Mismatch                 = $false
        Ambiguous                = $false
        SharePointUserId         = ''
        LoginName                = ''
        WasAffectedUserSiteAdmin = $false
        Attempts                 = 0
        Action                   = 'Ninguna'
        Status                   = 'OK'
        Message                  = ''
    }
}

function Invoke-ProcessAndRepairSites {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Inventory
    )

    Write-AppHeader 'Fase 2/3 · Analizar y reparar'
    Write-Styled 'Fast path: conectar → revisar UserInfoList → reparar → cerrar.' Muted
    Write-Styled 'Solo se espera/reintenta cuando un sitio devuelve 401/403/Access denied.' Muted
    Write-Host ''

    Write-Field 'Usuario' $script:AffectedUser.Upn Primary
    Write-Field 'SID actual' $script:AffectedUser.CurrentSid
    Write-Field 'Modo' `
        $(if ($script:Settings.DryRun) { 'SIMULACIÓN' } else { 'REAL' }) `
        $(if ($script:Settings.DryRun) { 'Warning' } else { 'Danger' })

    $settle = [Math]::Max(0, [int]$script:Settings.GlobalGrantSettleSeconds)

    if ($settle -gt 0) {
        Write-Host ''
        Write-Status Info "Espera global de propagación: $settle s."
        Start-Sleep -Seconds $settle
    }

    Write-Host ''

    $results = [System.Collections.Generic.List[object]]::new()
    $rows = @($Inventory | Where-Object { -not $_.Skip })

    for ($i = 0; $i -lt $rows.Count; $i++) {
        $record = $rows[$i]
        $result = New-RepairResult -Record $record

        $pct = [int]((($i + 1) / [Math]::Max(1, $rows.Count)) * 100)
        Write-Progress `
            -Id 11 `
            -Activity 'Fase 2/3 · Procesando UserInfoList' `
            -Status "$($i + 1)/$($rows.Count) $($record.Url)" `
            -PercentComplete $pct

        if ($record.AdminGrantStatus -eq 'Error') {
            $result.Status = 'Error'
            $result.Action = 'No procesado'
            $result.Message = "No fue posible preparar Site Collection Admin: $($record.ProcessingMessage)"
            [void]$results.Add($result)
            continue
        }

        $start = Get-Date
        $delay = [Math]::Max(1, [int]$script:Settings.PropagationInitialDelaySeconds)
        $timeout = [Math]::Max(15, [int]$script:Settings.PropagationTimeoutSeconds)
        $attempt = 0

        while ($true) {
            $attempt++
            $result.Attempts = $attempt

            $attemptResult = Invoke-SiteProcessingAttempt `
                -Record $record `
                -Result $result

            if ($attemptResult.Success) {
                $record.PropagationStatus = 'OK'
                $record.PropagationMessage = "Acceso operativo en intento $attempt"
                $record.ProcessingStatus = $result.Status
                $record.ProcessingMessage = $result.Message
                break
            }

            if (-not $attemptResult.AccessIssue) {
                $result.Status = 'Error'
                $result.Action = 'Error'
                $result.Message = $attemptResult.Message
                $record.ProcessingStatus = 'Error'
                $record.ProcessingMessage = $attemptResult.Message
                break
            }

            $elapsed = [int]((Get-Date) - $start).TotalSeconds

            if ($elapsed -ge $timeout) {
                $result.Status = 'Error'
                $result.Action = 'No procesado'
                $result.Message = "Timeout esperando propagación administrativa: $($attemptResult.Message)"
                $record.PropagationStatus = 'Error'
                $record.PropagationMessage = $result.Message
                $record.ProcessingStatus = 'Error'
                $record.ProcessingMessage = $result.Message
                break
            }

            Write-Status Warn (
                "Acceso aún no propagado · intento $attempt · " +
                "reintento en $delay s · $($record.Url)"
            )

            Start-Sleep -Seconds $delay
            $delay = [Math]::Min(15, $delay + 2)
        }

        [void]$results.Add($result)
        Save-ActiveGrantState -Records $Inventory
    }

    Write-Progress -Id 11 -Activity 'Fase 2/3 · Procesando UserInfoList' -Completed
    return @($results.ToArray())
}

function Test-ShouldRemoveAdminAfterRun {
    param([Parameter(Mandatory)]$Record)

    if ($Record.Skip) { return $false }
    if ($Record.AdminOwnOneDrive) { return $false }
    if ($Record.ProtectedSpo) { return $false }

    return [bool]$Record.RemoveAdminAfterRun
}

function Remove-AdminFromSite {
    param([Parameter(Mandatory)]$Record)

    $conn = $null

    try {
        $conn = Connect-Site -Url $Record.Url

        Remove-PnPSiteCollectionAdmin `
            -Owners @($script:Settings.AdminUpn) `
            -Connection $conn `
            -ErrorAction Stop

        $Record.AdminCleanupStatus = 'Retirado'
        $Record.AdminCleanupMessage = ''
        return $true
    }
    catch {
        $Record.AdminCleanupStatus = 'Error'
        $Record.AdminCleanupMessage = $_.Exception.Message
        return $false
    }
    finally {
        Disconnect-Safe $conn
    }
}

function Invoke-RestoreAdminAccess {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Inventory
    )

    Write-AppHeader 'Fase 3/3 · Restaurar Site Collection Admin'
    Write-Styled 'ODB de terceros y SPO no protegidos: retirar operador.' Muted
    Write-Styled 'Tu OneDrive y SPO protegidos: conservar.' Muted
    Write-Host ''

    $rows = @($Inventory | Where-Object { -not $_.Skip })

    for ($i = 0; $i -lt $rows.Count; $i++) {
        $record = $rows[$i]

        $pct = [int]((($i + 1) / [Math]::Max(1, $rows.Count)) * 100)
        Write-Progress `
            -Id 12 `
            -Activity 'Fase 3/3 · Restaurando Site Collection Admin' `
            -Status "$($i + 1)/$($rows.Count) $($record.Url)" `
            -PercentComplete $pct

        if (Test-ShouldRemoveAdminAfterRun -Record $record) {
            [void](Remove-AdminFromSite -Record $record)
        }
        else {
            if ($record.AdminOwnOneDrive) {
                $record.AdminCleanupStatus = 'Conservado · OneDrive propio del admin'
            }
            elseif ($record.ProtectedSpo) {
                $record.AdminCleanupStatus = 'Conservado · SPO protegido'
            }
            elseif ($record.Skip) {
                $record.AdminCleanupStatus = 'Omitido'
            }
            else {
                $record.AdminCleanupStatus = 'No requerido'
            }
        }

        Save-ActiveGrantState -Records $Inventory
    }

    Write-Progress -Id 12 -Activity 'Fase 3/3 · Restaurando Site Collection Admin' -Completed

    $remaining = @(
        $Inventory | Where-Object {
            (Test-ShouldRemoveAdminAfterRun -Record $_) -and
            $_.AdminCleanupStatus -ne 'Retirado'
        }
    )

    if ($remaining.Count -eq 0) {
        Clear-ActiveGrantState
    }
    else {
        Save-ActiveGrantState -Records $Inventory
    }
}

# Compatibilidad con estados de versiones anteriores.
function Invoke-CleanupAdminAccess {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Inventory
    )

    foreach ($record in $Inventory) {
        if ($null -eq $record.PSObject.Properties['RemoveAdminAfterRun']) {
            $remove = $false

            if ($record.Kind -eq 'OneDrive') {
                $remove = (-not [bool]$record.AdminOwnOneDrive) -and
                          (-not [bool]$record.Skip) -and
                          [bool]$record.AdminAddedByTool
            }
            elseif ($record.Kind -eq 'SharePoint') {
                $remove = [bool]$record.AdminAddedByTool -and
                          (-not [bool]$record.ProtectedSpo)
            }

            $record | Add-Member `
                -NotePropertyName RemoveAdminAfterRun `
                -NotePropertyValue $remove
        }
    }

    Invoke-RestoreAdminAccess -Inventory $Inventory
}

# =============================================================================
# RECUPERACIÓN DE ELEVACIONES
# =============================================================================

function Invoke-RecoverAdminGrants {
    $state = Load-ActiveGrantState

    Write-AppHeader 'Recuperar elevaciones temporales'

    if ($null -eq $state -or $null -eq $state.Records) {
        Write-Status Ok 'No hay una sesión pendiente de limpieza.'
        Pause-Tui
        return
    }

    Write-Field 'Creada' $state.CreatedAt
    Write-Field 'Administrador' $state.AdminUpn
    Write-Field 'Registros' @($state.Records).Count
    Write-Host ''
    Write-Status Warn 'Esta opción intenta retirar únicamente los grants marcados como temporales en el estado guardado.'
    Write-Host ''

    if (-not (Read-YesNo '¿Continuar con la recuperación?' $false)) {
        return
    }

    $oldAdmin = $script:Settings.AdminUpn
    if (-not [string]::IsNullOrWhiteSpace([string]$state.AdminUpn)) {
        $script:Settings.AdminUpn = [string]$state.AdminUpn
    }

    try {
        Invoke-CleanupAdminAccess -Inventory @($state.Records)
        Write-Status Ok 'Recuperación finalizada.'
    }
    catch {
        Write-Status Error $_.Exception.Message
    }
    finally {
        $script:Settings.AdminUpn = $oldAdmin
    }

    Pause-Tui
}

# =============================================================================
# SITIOS SPO PROTEGIDOS
# =============================================================================

function Find-SitesByText {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Sites,
        [Parameter(Mandatory)][string]$Query
    )

    $q = $Query.Trim()
    return @(
        $Sites | Where-Object {
            ([string]$_.Title -like "*$q*") -or
            ([string]$_.Url -like "*$q*") -or
            ([string]$_.Owner -like "*$q*")
        } | Sort-Object Title, Url
    )
}

function Add-ProtectedSpoSite {
    try {
        $null = Connect-AdminCenter
        $sites = @(Get-TenantSites -Scope SharePoint)

        Write-AppHeader 'Agregar sitio SPO protegido'
        $query = Read-TextValue 'Buscar por nombre, URL o propietario' -Required
        $matches = @(Find-SitesByText -Sites $sites -Query $query)

        if ($matches.Count -eq 0) {
            Write-Status Warn 'No se encontraron sitios.'
            Pause-Tui
            return
        }

        if ($matches.Count -gt 50) {
            Write-Status Warn "Hay $($matches.Count) coincidencias. Refina la búsqueda."
            Pause-Tui
            return
        }

        $items = @(
            foreach ($site in $matches) {
                [pscustomobject]@{
                    Label = if ([string]::IsNullOrWhiteSpace([string]$site.Title)) { [string]$site.Url } else { [string]$site.Title }
                    Value = [string]$site.Url
                    Hint = "$($site.Url) | Owner: $($site.Owner)"
                }
            }
        )

        $choice = Read-MenuChoice -Title 'Seleccionar sitio protegido' -Items $items
        if ($null -eq $choice) { return }

        if (@($script:Settings.ProtectedSpoSites) -notcontains $choice.Value) {
            $script:Settings.ProtectedSpoSites = @($script:Settings.ProtectedSpoSites) + @($choice.Value)
            Save-Settings
        }

        Write-Status Ok 'Sitio agregado a la lista protegida.'
    }
    catch {
        Write-Status Error $_.Exception.Message
    }
    finally {
        Disconnect-Safe $script:AdminConnection
        $script:AdminConnection = $null
    }

    Pause-Tui
}

function Remove-ProtectedSpoSite {
    $sites = @($script:Settings.ProtectedSpoSites)

    if ($sites.Count -eq 0) {
        Write-AppHeader 'Sitios SPO protegidos'
        Write-Status Info 'No hay sitios protegidos configurados.'
        Pause-Tui
        return
    }

    $items = @(
        foreach ($url in $sites) {
            [pscustomobject]@{
                Label = $url
                Value = $url
                Hint = 'Se preserva Site Collection Admin después de las ejecuciones.'
            }
        }
    )

    $choice = Read-MenuChoice -Title 'Quitar sitio SPO protegido' -Items $items
    if ($null -eq $choice) { return }

    $script:Settings.ProtectedSpoSites = @(
        $sites | Where-Object { $_ -ne $choice.Value }
    )
    Save-Settings

    Write-AppHeader 'Sitios SPO protegidos'
    Write-Status Ok 'Sitio retirado de la lista protegida.'
    Pause-Tui
}

function Show-ProtectedSpoMenu {
    while ($true) {
        $items = @(
            [pscustomobject]@{
                Label = 'Agregar sitio protegido'
                Value = 'Add'
                Hint = 'Buscar dinámicamente un sitio SPO del tenant.'
            },
            [pscustomobject]@{
                Label = 'Quitar sitio protegido'
                Value = 'Remove'
                Hint = 'Dejará de preservarse por esta política.'
            }
        )

        $choice = Read-MenuChoice -Title 'Sitios SPO protegidos' -Items $items -Description @(
            "Configurados: $(@($script:Settings.ProtectedSpoSites).Count)",
            'Estos sitios conservan tu Site Collection Admin al terminar.'
        )

        if ($null -eq $choice) { return }

        switch ($choice.Value) {
            'Add' { Add-ProtectedSpoSite }
            'Remove' { Remove-ProtectedSpoSite }
        }
    }
}

# =============================================================================
# CONFIGURACIÓN
# =============================================================================

function Show-SettingsMenu {
    while ($true) {
        $scopeLabel = switch ($script:Settings.Scope) {
            'SharePoint' { 'Solo SharePoint' }
            'OneDrive' { 'Solo OneDrive' }
            default { 'SharePoint + OneDrive' }
        }

        $items = @(
            [pscustomobject]@{
                Label = 'Tenant'
                Value = 'Tenant'
                Hint = $script:Settings.Tenant
            },
            [pscustomobject]@{
                Label = 'UPN del administrador'
                Value = 'Admin'
                Hint = $script:Settings.AdminUpn
            },
            [pscustomobject]@{
                Label = 'Alcance predeterminado'
                Value = 'Scope'
                Hint = $scopeLabel
            },
            [pscustomobject]@{
                Label = 'Sitios SPO protegidos'
                Value = 'Protected'
                Hint = "$(@($script:Settings.ProtectedSpoSites).Count) sitio(s)"
            },
            [pscustomobject]@{
                Label = 'Timeout de propagación'
                Value = 'Propagation'
                Hint = "$($script:Settings.PropagationTimeoutSeconds) segundos"
            },
            [pscustomobject]@{
                Label = 'Exclusiones de sitios de sistema'
                Value = 'SystemSites'
                Hint = if ($script:Settings.SkipSystemSites) { 'ACTIVADAS · root tenant, My Site Host, App Catalog, Search/eDiscovery...' } else { 'DESACTIVADAS' }
            },
            [pscustomobject]@{
                Label = 'Fast path / propagación'
                Value = 'FastPath'
                Hint = "Espera global $($script:Settings.GlobalGrantSettleSeconds) s | retry individual solo en 401/403"
            },
            [pscustomobject]@{
                Label = 'Validación de App Registration'
                Value = 'AppValidation'
                Hint = "$($script:Settings.AppValidationRetries) intento(s) | espera $($script:Settings.AppValidationDelaySeconds) s"
            },
            [pscustomobject]@{
                Label = 'Sesión persistente'
                Value = 'Persist'
                Hint = if ($script:Settings.PersistLogin) { 'Activada' } else { 'Desactivada' }
            }
            [pscustomobject]@{
                Label = 'Cambiar idioma'
                Value = 'Language'
                Hint = if ($script:Language -eq 'en') { 'English' } else { 'Español' }
            }
        )

        $choice = Read-MenuChoice -Title 'Configuración' -Items $items
        if ($null -eq $choice) { return }

        switch ($choice.Value) {
            'Language' {
                Initialize-AppLanguage
                $script:Settings.Language = $script:Language
                Save-Settings
            }

            'Tenant' {
                Write-AppHeader 'Tenant'
                $value = Read-TextValue 'Tenant' -Default $script:Settings.Tenant -Required
                if ($value -match '^[a-zA-Z0-9][a-zA-Z0-9.-]*\.onmicrosoft\.com$') {
                    $script:Settings.Tenant = $value
                    Save-Settings
                }
                else {
                    Write-Status Error 'Tenant no válido.'
                    Pause-Tui
                }
            }

            'Admin' {
                Write-AppHeader 'Administrador'
                $value = Read-TextValue 'UPN administrador' -Default $script:Settings.AdminUpn -Required
                if ($value -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                    $script:Settings.AdminUpn = $value
                    Save-Settings
                }
                else {
                    Write-Status Error 'UPN no válido.'
                    Pause-Tui
                }
            }

            'Scope' {
                $scopeItems = @(
                    [pscustomobject]@{ Label='SharePoint Online'; Value='SharePoint'; Hint='Solo sitios SPO.' },
                    [pscustomobject]@{ Label='OneDrive for Business'; Value='OneDrive'; Hint='Excluye automáticamente el OneDrive propio del usuario afectado.' },
                    [pscustomobject]@{ Label='SharePoint + OneDrive'; Value='Both'; Hint='Recorre ambos tipos.' }
                )

                $s = Read-MenuChoice -Title 'Alcance predeterminado' -Items $scopeItems
                if ($null -ne $s) {
                    $script:Settings.Scope = $s.Value
                    Save-Settings
                }
            }

            'Protected' { Show-ProtectedSpoMenu }

            'Propagation' {
                Write-AppHeader 'Propagación'
                $script:Settings.PropagationTimeoutSeconds = Read-Integer -Prompt 'Timeout total por sitio (segundos)' -Minimum 15 -Maximum 900 -Default ([int]$script:Settings.PropagationTimeoutSeconds)
                $script:Settings.PropagationInitialDelaySeconds = Read-Integer -Prompt 'Espera inicial entre pruebas (segundos)' -Minimum 1 -Maximum 30 -Default ([int]$script:Settings.PropagationInitialDelaySeconds)
                Save-Settings
            }

            'SystemSites' {
                while ($true) {
                    $sysItems = @(
                        [pscustomobject]@{
                            Label = if ($script:Settings.SkipSystemSites) { 'Desactivar exclusiones de sistema' } else { 'Activar exclusiones de sistema' }
                            Value = 'Toggle'
                            Hint = 'Recomendado: mantener activado.'
                        },
                        [pscustomobject]@{
                            Label = 'Ver reglas activas'
                            Value = 'View'
                            Hint = 'Muestra roots, templates y fragmentos de URL excluidos.'
                        }
                    )

                    $sysChoice = Read-MenuChoice `
                        -Title 'Exclusiones de sitios de sistema' `
                        -Items $sysItems `
                        -Description @(
                            "Estado: $(if ($script:Settings.SkipSystemSites) { 'ACTIVADAS' } else { 'DESACTIVADAS' })",
                            'Estas reglas se evalúan antes de conceder Site Collection Admin.'
                        )

                    if ($null -eq $sysChoice) { break }

                    if ($sysChoice.Value -eq 'Toggle') {
                        if ($script:Settings.SkipSystemSites) {
                            Write-AppHeader 'Desactivar exclusiones'
                            Write-Status Warn 'No se recomienda procesar hosts raíz ni colecciones de sistema.'
                            Write-Host ''
                            if (Read-YesNo '¿Desactivar de todas formas?' $false) {
                                $script:Settings.SkipSystemSites = $false
                                Save-Settings
                            }
                        }
                        else {
                            $script:Settings.SkipSystemSites = $true
                            Save-Settings
                        }
                    }
                    elseif ($sysChoice.Value -eq 'View') {
                        Write-AppHeader 'Reglas de exclusión'
                        $roots = Get-TenantRootUrls

                        Write-Section 'Hosts raíz'
                        Write-Field 'SharePoint root' $roots.SharePointRoot Warning
                        Write-Field 'My Site Host' $roots.MySiteHostRoot Warning
                        Write-Field 'Admin Center' $roots.AdminRoot Warning

                        Write-Section 'Templates'
                        foreach ($tpl in @($script:Settings.ExcludedSystemTemplates)) {
                            Write-Styled "• $tpl" Muted
                        }

                        Write-Section 'Rutas'
                        foreach ($fragment in @($script:Settings.ExcludedSystemUrlFragments)) {
                            Write-Styled "• $fragment" Muted
                        }

                        Pause-Tui
                    }
                }
            }

            'FastPath' {
                Write-AppHeader 'Fast path / propagación'
                Write-Styled 'Esta espera se aplica UNA sola vez después de completar todos los grants.' Muted
                Write-Styled 'Los sitios que todavía devuelvan 401/403 se reintentan individualmente.' Muted
                Write-Host ''

                $script:Settings.GlobalGrantSettleSeconds = Read-Integer `
                    -Prompt 'Espera global después de grants (segundos)' `
                    -Minimum 0 `
                    -Maximum 60 `
                    -Default ([int]$script:Settings.GlobalGrantSettleSeconds)

                $script:Settings.PropagationTimeoutSeconds = Read-Integer `
                    -Prompt 'Timeout máximo solo para sitios con 401/403' `
                    -Minimum 15 `
                    -Maximum 600 `
                    -Default ([int]$script:Settings.PropagationTimeoutSeconds)

                Save-Settings
            }

            'AppValidation' {
                Write-AppHeader 'Validación de App Registration'
                $script:Settings.AppValidationRetries = Read-Integer `
                    -Prompt 'Intentos de validación' `
                    -Minimum 1 `
                    -Maximum 10 `
                    -Default ([int]$script:Settings.AppValidationRetries)

                $script:Settings.AppValidationDelaySeconds = Read-Integer `
                    -Prompt 'Espera entre intentos (segundos)' `
                    -Minimum 1 `
                    -Maximum 30 `
                    -Default ([int]$script:Settings.AppValidationDelaySeconds)

                Save-Settings
            }

            'Persist' {
                $script:Settings.PersistLogin = -not [bool]$script:Settings.PersistLogin
                Save-Settings
            }
        }
    }
}

# =============================================================================
# FLUJO COMPLETO
# =============================================================================

function Confirm-FullRun {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('SharePoint','OneDrive','Both')]
        [string]$Scope
    )

    Write-AppHeader 'Confirmar ejecución'

    Write-Field 'Usuario' $script:AffectedUser.Upn Primary
    Write-Field 'Administrador' $script:Settings.AdminUpn
    Write-Field 'Alcance' $Scope
    Write-Field 'SPO protegidos' @($script:Settings.ProtectedSpoSites).Count
    Write-Field 'Excluir sistema' $(if ($script:Settings.SkipSystemSites) { 'SÍ' } else { 'NO' }) $(if ($script:Settings.SkipSystemSites) { 'Success' } else { 'Warning' })
    Write-Field 'Fast path' 'ACTIVADO' Success
    Write-Field 'Espera global' "$($script:Settings.GlobalGrantSettleSeconds) s"
    Write-Field 'Modo' `
        $(if ($script:Settings.DryRun) { 'SIMULACIÓN' } else { 'REAL' }) `
        $(if ($script:Settings.DryRun) { 'Warning' } else { 'Danger' })

    Write-Host ''
    Write-Status Info 'Flujo'
    Write-Styled '1. Grant de Site Collection Admin a todos los targets.' Muted
    Write-Styled '2. Cerrar Admin Center y esperar unos segundos globalmente.' Muted
    Write-Styled '3. Conectar/procesar cada sitio directamente; retry solo en 401/403.' Muted
    Write-Styled '4. Retirar Site Admin temporal.' Muted

    Write-Host ''
    Write-Status Warn 'Hosts raíz y sitios de sistema: OMITIDOS antes de cualquier grant.'
    Write-Styled 'Incluye tenant.sharepoint.com, tenant-my.sharepoint.com, Admin Center y templates/rutas de sistema configuradas.' Muted
    Write-Status Warn 'OneDrive propio del usuario afectado: OMITIDO completamente.'
    Write-Styled 'OneDrive de terceros: el operador se retira siempre al final.' Muted
    Write-Styled 'SPO no protegido: el operador se retira al final.' Muted
    Write-Styled 'SPO protegido: el operador permanece Site Collection Admin.' Muted

    Write-Host ''

    if (-not $script:Settings.DryRun) {
        Write-Status Warn 'MODO REAL: se eliminarán únicamente mismatches con SID actual válido y comparación inequívoca.'
        Write-Host ''
        $confirmation = (Read-Host 'Escribe REPARAR para continuar').Trim()
        return ($confirmation -ceq 'REPARAR')
    }

    Write-Styled 'SIMULACIÓN no elimina UserInfoList; los grants administrativos sí son reales.' Muted
    Write-Host ''
    return (Read-YesNo '¿Iniciar el análisis?' $false)
}

function Start-FullRepair {
    if (-not (Ensure-PnPModule)) { return }
    if (-not (Ensure-TenantConfigured)) { return }
    if (-not (Ensure-ClientId)) { return }
    if (-not (Ensure-AdminUpn)) { return }

    try {
        Write-AppHeader 'Validación previa'
        Write-Status Info 'Validando la identidad ACTUAL del usuario...'

        if ([string]::IsNullOrWhiteSpace($script:Settings.AffectedUserUpn)) {
            Write-Status Warn 'Selecciona primero el usuario afectado.'
            Pause-Tui
            return
        }

        $null = Connect-AdminCenter
        $script:AffectedUser = Resolve-AffectedUser `
            -Upn $script:Settings.AffectedUserUpn

        if ([string]::IsNullOrWhiteSpace($script:AffectedUser.CurrentSid)) {
            throw 'El SID actual está vacío. Se bloquea la ejecución para impedir falsos positivos.'
        }

        Write-Status Ok 'Identidad actual confirmada.'
        Write-Field 'UPN' $script:AffectedUser.Upn Primary
        Write-Field 'SID actual' $script:AffectedUser.CurrentSid Success
        Write-Field 'OneDrive propio' $script:AffectedUser.PersonalUrl Muted
    }
    catch {
        Write-Status Error $_.Exception.Message
        Pause-Tui
        return
    }
    finally {
        Disconnect-Safe $script:AdminConnection
        $script:AdminConnection = $null
    }

    $scopeItems = @(
        [pscustomobject]@{
            Label='SharePoint Online'
            Value='SharePoint'
            Hint='SPO no protegido se considera acceso temporal.'
        },
        [pscustomobject]@{
            Label='OneDrive for Business'
            Value='OneDrive'
            Hint='Own OneDrive del afectado se omite; ODB de terceros se limpia al final.'
        },
        [pscustomobject]@{
            Label='SharePoint + OneDrive'
            Value='Both'
            Hint='Ejecuta ambos alcances.'
        }
    )

    $scopeChoice = Read-MenuChoice `
        -Title 'Alcance de esta ejecución' `
        -Items $scopeItems `
        -Description @("Predeterminado: $($script:Settings.Scope)")

    if ($null -eq $scopeChoice) { return }

    $scope = $scopeChoice.Value
    $script:Settings.Scope = $scope
    Save-Settings

    if (-not (Confirm-FullRun -Scope $scope)) {
        return
    }

    $inventory = @()
    $results = @()
    $reportPath = ''
    $inventoryPath = ''
    $fatal = $null

    try {
        # FASE 1: grants masivos / rápidos desde Admin Center.
        $null = Connect-AdminCenter
        $inventory = @(Invoke-PrepareAdminAccess -Scope $scope)
        $script:LastInventory = $inventory

        Disconnect-Safe $script:AdminConnection
        $script:AdminConnection = $null

        # FASE 2: el propio intento de leer UserInfoList valida el acceso.
        $results = @(Invoke-ProcessAndRepairSites -Inventory $inventory)
        $script:LastResults = $results
    }
    catch {
        $fatal = $_
    }
    finally {
        Disconnect-Safe $script:AdminConnection
        $script:AdminConnection = $null

        # FASE 3: intentar siempre la restauración.
        if (@($inventory).Count -gt 0) {
            try {
                Invoke-RestoreAdminAccess -Inventory $inventory
            }
            catch {
                Write-AppHeader 'Restauración incompleta'
                Write-Status Error $_.Exception.Message
                Write-Styled "Estado recuperable: $script:StatePath" Muted
                Pause-Tui
            }
        }
    }

    try {
        $reportPath = New-ReportPath -Prefix 'UserInfoList-IDMismatch'
        $inventoryPath = New-ReportPath -Prefix 'AdminAccess'

        Export-RecordsCsv -Records $results -Path $reportPath
        Export-RecordsCsv -Records $inventory -Path $inventoryPath
    }
    catch {
        if ($null -eq $fatal) { $fatal = $_ }
    }

    if ($null -ne $fatal) {
        Write-AppHeader 'Proceso interrumpido'
        Write-Status Error $fatal.Exception.Message

        if (Test-Path -LiteralPath $script:StatePath) {
            Write-Status Warn 'Existe estado recuperable de Site Collection Admin.'
            Write-Styled $script:StatePath Muted
        }

        Pause-Tui
        return
    }

    Show-RunSummary `
        -Results $results `
        -Inventory $inventory `
        -ReportPath $reportPath

    Write-Host ''
    Write-Styled 'Auditoría de Site Collection Admin:' Muted
    Write-Styled $inventoryPath Primary
    Pause-Tui
}

# =============================================================================
# DIAGNÓSTICO
# =============================================================================

function Run-Diagnostics {
    Write-AppHeader 'Diagnóstico'

    $module = Get-InstalledPnPModule
    $auth = Get-AppAuthState
    $statePending = Test-Path -LiteralPath $script:StatePath

    Write-Section 'Entorno'
    Write-Field 'PowerShell' $PSVersionTable.PSVersion $(if ($PSVersionTable.PSVersion -ge [version]'7.4') { 'Success' } else { 'Danger' })
    Write-Field 'PnP.PowerShell' $(if ($null -eq $module) { 'No instalado' } else { $module.Version }) $(if ($null -ne $module -and $module.Version -ge $script:MinimumPnPVersion) { 'Success' } else { 'Danger' })

    Write-Section 'Autenticación'
    Write-Field 'Estado' $auth.Label $auth.Style
    Write-Field 'Tenant' $script:Settings.Tenant
    Write-Field 'Aplicación' $script:Settings.AppRegistrationName
    Write-Field 'Client ID' $(if (Test-ClientIdFormat $script:Settings.ClientId) { $script:Settings.ClientId } else { 'No configurado' })
    Write-Field 'PersistLogin' $(if ($script:Settings.PersistLogin) { 'Activado' } else { 'Desactivado' })

    Write-Section 'Operación'
    Write-Field 'Administrador' $script:Settings.AdminUpn
    Write-Field 'Usuario afectado' $script:Settings.AffectedUserUpn
    Write-Field 'Estado pendiente' $(if ($statePending) { 'SÍ' } else { 'No' }) $(if ($statePending) { 'Warning' } else { 'Success' })

    if ($null -ne $module -and
        $module.Version -ge $script:MinimumPnPVersion -and
        $auth.Configured) {

        Write-Host ''
        Write-Status Info 'Ejecutando prueba real de autenticación...'

        $test = Test-PnPAppRegistrationWithRetry `
            -ClientId $script:Settings.ClientId `
            -Tenant $script:Settings.Tenant

        Show-AppValidationResult `
            -Result $test `
            -SuccessTitle 'Autenticación operativa'
    }

    Pause-Tui
}

# =============================================================================
# MENÚ PRINCIPAL
# =============================================================================

function Get-AffectedUserStatusText {
    if ($null -ne $script:AffectedUser) {
        return "$($script:AffectedUser.Upn) · SID validado"
    }

    if (-not [string]::IsNullOrWhiteSpace($script:Settings.AffectedUserUpn)) {
        return "$($script:Settings.AffectedUserUpn) · pendiente de validar"
    }

    return 'No seleccionado'
}

function Write-MainDashboard {
    $auth = Get-AppAuthState
    $mode = if ($script:Settings.DryRun) { 'SIMULACIÓN' } else { 'REAL' }
    $modeStyle = if ($script:Settings.DryRun) { 'Warning' } else { 'Danger' }
    $pending = Test-Path -LiteralPath $script:StatePath

    Write-Section 'Estado'
    Write-Field 'Tenant' $(if (Test-TenantFormat $script:Settings.Tenant) { $script:Settings.Tenant } else { 'No configurado' })
    Write-Field 'Autenticación' $auth.Label $auth.Style
    Write-Field 'Administrador' $(if ([string]::IsNullOrWhiteSpace($script:Settings.AdminUpn)) { 'No configurado' } else { $script:Settings.AdminUpn })
    Write-Field 'Usuario' (Get-AffectedUserStatusText)
    Write-Field 'Modo' $mode $modeStyle

    if ($pending) {
        Write-Field 'Recuperación' 'GRANTS PENDIENTES' Warning
    }
    else {
        Write-Field 'Recuperación' 'Limpia' Success
    }

    Write-Host ''
}

function Show-MainMenu {
    while ($true) {
        $mode = if ($script:Settings.DryRun) { 'SIMULACIÓN' } else { 'REAL' }
        $auth = Get-AppAuthState
        $pending = Test-Path -LiteralPath $script:StatePath

        Write-AppHeader 'Inicio'
        Write-MainDashboard

        $items = @(
            [pscustomobject]@{
                Label = 'Seleccionar / validar usuario afectado'
                Value = 'User'
                Hint = 'Resuelve AccountName, SID actual y OneDrive propio.'
            },
            [pscustomobject]@{
                Label = 'Ejecutar reparación tenant'
                Value = 'Run'
                Hint = 'Grant masivo → espera global → procesar directo → restaurar admins.'
            },
            [pscustomobject]@{
                Label = if ($script:Settings.DryRun) { 'Cambiar a modo REAL' } else { 'Cambiar a SIMULACIÓN' }
                Value = 'Mode'
                Hint = if ($script:Settings.DryRun) {
                    'SIMULACIÓN no elimina UserInfoList; los grants temporales sí pueden aplicarse.'
                }
                else {
                    'REAL elimina únicamente mismatches inequívocos.'
                }
            },
            [pscustomobject]@{
                Label = 'Aplicación Entra / autenticación PnP'
                Value = 'Auth'
                Hint = if ($auth.Configured) {
                    'Configurada · validar, cambiar o registrar otra app.'
                }
                else {
                    'No configurada · usar existente o registrar nueva.'
                }
            },
            [pscustomobject]@{
                Label = 'Sitios SPO protegidos'
                Value = 'Protected'
                Hint = "$(@($script:Settings.ProtectedSpoSites).Count) sitio(s) conservarán tu Site Collection Admin."
            },
            [pscustomobject]@{
                Label = 'Recuperar / retirar grants temporales'
                Value = 'Recover'
                Hint = if ($pending) {
                    'Hay un estado pendiente de restauración.'
                }
                else {
                    'No hay grants pendientes registrados.'
                }
            },
            [pscustomobject]@{
                Label = 'Configuración'
                Value = 'Settings'
                Hint = 'Tenant, admin, alcance, exclusiones de sistema, fast path y autenticación.'
            },
            [pscustomobject]@{
                Label = 'Diagnóstico'
                Value = 'Diagnostics'
                Hint = 'Valida PowerShell, PnP, App Registration y acceso tenant.'
            }
        )

        # Dibujamos el menú aquí, conservando la TUI numérica y el 0 para salir.
        for ($i = 0; $i -lt $items.Count; $i++) {
            $item = $items[$i]
            Write-Styled ('  {0,2}  ' -f ($i + 1)) Primary -NoNewline
            Write-Host $item.Label
            if (-not [string]::IsNullOrWhiteSpace([string]$item.Hint)) {
                Write-Styled "      $($item.Hint)" Muted
            }
            Write-Host ''
        }

        Write-Styled '   0  ' Muted -NoNewline
        Write-Host 'Salir'
        Write-Host ''

        $raw = (Read-Host 'Selecciona una opción').Trim()
        $selected = 0

        if (-not [int]::TryParse($raw, [ref]$selected) -or
            $selected -lt 0 -or
            $selected -gt $items.Count) {

            Write-Host ''
            Write-Status Warn 'Opción no válida.'
            Start-Sleep -Milliseconds 700
            continue
        }

        if ($selected -eq 0) { return }

        $choice = $items[$selected - 1]

        switch ($choice.Value) {
            'User' {
                Select-AffectedUser
            }

            'Run' {
                Start-FullRepair
            }

            'Mode' {
                if ($script:Settings.DryRun) {
                    Write-AppHeader 'Activar modo REAL'
                    Write-Status Warn 'El modo REAL puede eliminar entradas legacy de UserInfoList.'
                    Write-Styled 'La confirmación REPARAR seguirá siendo obligatoria al ejecutar.' Muted
                    Write-Host ''

                    if (Read-YesNo '¿Cambiar a modo REAL?' $false) {
                        $script:Settings.DryRun = $false
                        Save-Settings
                    }
                }
                else {
                    $script:Settings.DryRun = $true
                    Save-Settings
                }
            }

            'Auth' {
                if (Ensure-PnPModule) {
                    Show-AppRegistrationMenu
                }
            }

            'Protected' {
                if ((Ensure-PnPModule) -and
                    (Ensure-ClientId) -and
                    (Ensure-TenantConfigured)) {

                    Show-ProtectedSpoMenu
                }
            }

            'Recover' {
                if ((Ensure-PnPModule) -and
                    (Ensure-ClientId) -and
                    (Ensure-TenantConfigured)) {

                    Invoke-RecoverAdminGrants
                }
            }

            'Settings' {
                Show-SettingsMenu
            }

            'Diagnostics' {
                Run-Diagnostics
            }
        }
    }
}

# =============================================================================
# INICIO
# =============================================================================

try {
    Initialize-AppLanguage
    Initialize-Terminal
    Ensure-AppFolders
    $script:Settings = Load-Settings

    try {
        $Host.UI.RawUI.WindowTitle = "$script:AppName $script:AppVersion"
    }
    catch {}

    Show-MainMenu
}
catch {
    Write-AppHeader 'Error inesperado'
    Write-Status Error $_.Exception.Message

    if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
        Write-Host ''
        Write-Styled 'Detalles técnicos:' Muted
        Write-Styled $_.ScriptStackTrace Muted
    }

    if (Test-Path -LiteralPath $script:StatePath) {
        Write-Host ''
        Write-Status Warn 'Existe un estado de grants administrativos pendiente de recuperación.'
        Write-Styled $script:StatePath Muted
    }

    Pause-Tui
}
finally {
    Disconnect-Safe $script:AdminConnection
    $script:AdminConnection = $null
}

Clear-Host
