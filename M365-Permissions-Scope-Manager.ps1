#requires -Version 7.4
<#
.SYNOPSIS
    M365 Permissions Scope Manager
.DESCRIPTION
    TUI interactiva para descubrir sitios SharePoint / OneDrive, seleccionar bibliotecas,
    contar Unique Permission Scopes y restablecer herencia.

    Diseño:
      - Sin URLs de sitios, nombres de bibliotecas o rutas de reportes hardcodeadas.
      - Menús por número.
      - Selector múltiple de bibliotecas con ↑/↓, Espacio y Enter.
      - Configuración persistente mínima; el target se descubre en tiempo de ejecución.
      - Conexiones PnP explícitas (-ReturnConnection) para evitar depender de estado global.
      - Dry-run activado por defecto.

.REQUIREMENTS
    PowerShell 7.4+
    PnP.PowerShell 3.2+
#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# Runtime / paths
# -----------------------------------------------------------------------------

$script:AppName = 'M365 Permissions Scope Manager'
$script:AppVersion = '6.4.1'
$script:MinimumPnPVersion = [version]'3.2.0'
$script:DefaultPageSize = 2000
$script:DefaultMaxRetries = 5
$script:DefaultRetryDelay = 3
$script:RecommendedUniqueScopes = 5000
$script:MaximumUniqueScopes = 50000

$toolHome = if (-not [string]::IsNullOrWhiteSpace($env:M365_PNP_TOOLKIT_HOME)) {
    $env:M365_PNP_TOOLKIT_HOME
}
else {
    Join-Path $HOME '.m365-permissions-scope-manager'
}

$script:ConfigRoot = $toolHome
$script:ConfigPath = Join-Path $script:ConfigRoot 'settings.json'
$script:DefaultReportFolder = Join-Path $script:ConfigRoot 'reports'

$script:Settings = $null
$script:Target = $null
$script:TemporarySiteAdmin = $null
$script:Ansi = $false

# -----------------------------------------------------------------------------
# LANGUAGE
# -----------------------------------------------------------------------------

$script:Language = 'es'
$script:UiTranslations = @(
    @{ From = 'Contexto'; To = 'Context' }; @{ From = 'Inicio'; To = 'Home' }; @{ From = 'Finalizado'; To = 'Finished' }
    @{ From = 'Error inesperado'; To = 'Unexpected error' }; @{ From = 'Detalles técnicos:'; To = 'Technical details:' }
    @{ From = 'Configuración'; To = 'Settings' }; @{ From = 'Salir'; To = 'Exit' }; @{ From = 'Volver'; To = 'Back' }
    @{ From = 'Cancelar'; To = 'Cancel' }; @{ From = 'Diagnóstico'; To = 'Diagnostics' }; @{ From = 'autenticación'; To = 'authentication' }
    @{ From = 'Seleccionar / cambiar sitio'; To = 'Select / change site' }; @{ From = 'Cambiar bibliotecas del sitio actual'; To = 'Change libraries for the current site' }
    @{ From = 'Analizar Unique Permission Scopes'; To = 'Analyze Unique Permission Scopes' }; @{ From = 'Restablecer herencia'; To = 'Reset inheritance' }
    @{ From = 'Gestionar Site Collection Admin'; To = 'Manage Site Collection Admin' }; @{ From = 'Seleccionar'; To = 'Select' }
    @{ From = 'Sitio'; To = 'Site' }; @{ From = 'Bibliotecas'; To = 'Libraries' }; @{ From = 'Biblioteca'; To = 'Library' }
    @{ From = 'SIMULACIÓN'; To = 'SIMULATION' }; @{ From = 'REAL'; To = 'LIVE' }; @{ From = 'Modo'; To = 'Mode' }
    @{ From = 'Enter para continuar'; To = 'Press Enter to continue' }; @{ From = 'Escribe RESET para continuar'; To = 'Type RESET to continue' }
    @{ From = 'Procesando'; To = 'Processing' }; @{ From = 'Leyendo elementos'; To = 'Reading items' }
    @{ From = 'Correctos'; To = 'Successful' }; @{ From = 'Errores'; To = 'Errors' }; @{ From = 'Reporte'; To = 'Report' }
)
function Get-LocalizedText {
    param([AllowNull()][object]$Text)
    if ($null -eq $Text) { return '' }; $result = [string]$Text
    if ($script:Language -eq 'en') {
        $phrases = @(
            @{ From = 'Selecciona una opción'; To = 'Select an option' }
            @{ From = 'Idioma'; To = 'Language' }; @{ From = 'Cambiar idioma'; To = 'Change language' }
            @{ From = 'No seleccionado'; To = 'Not selected' }; @{ From = 'No configurado'; To = 'Not configured' }; @{ From = 'No instalado'; To = 'Not installed' }
            @{ From = 'Destino'; To = 'Target' }; @{ From = 'Actuales'; To = 'Current' }; @{ From = 'Seleccionadas'; To = 'Selected' }
            @{ From = 'CONFIGURADA'; To = 'CONFIGURED' }; @{ From = 'NO CONFIGURADA'; To = 'NOT CONFIGURED' }; @{ From = 'Activado'; To = 'Enabled' }; @{ From = 'Desactivado'; To = 'Disabled' }
            @{ From = 'Sí'; To = 'Yes' }; @{ From = 'SÍ'; To = 'YES' }
            @{ From = 'SharePoint u OneDrive; al elegirlo se seleccionan sus bibliotecas.'; To = 'SharePoint or OneDrive; its libraries are selected when chosen.' }
            @{ From = 'Disponible después de seleccionar un sitio.'; To = 'Available after selecting a site.' }
            @{ From = 'Solo lectura · muestra recomendado 5,000 y máximo 50,000 por biblioteca · genera CSV.'; To = 'Read-only · shows the recommended 5,000 and maximum 50,000 per library · generates a CSV.' }
            @{ From = 'Respeta Dry-run y procesa las bibliotecas seleccionadas.'; To = 'Respects dry-run mode and processes the selected libraries.' }
            @{ From = 'El modo simulación es el valor seguro.'; To = 'Simulation mode is the safe option.' }
            @{ From = 'Cambiar a modo REAL'; To = 'Switch to LIVE mode' }; @{ From = 'Cambiar a modo SIMULACIÓN'; To = 'Switch to SIMULATION mode' }
            @{ From = 'El modo REAL puede'; To = 'LIVE mode can' }; @{ From = 'Modo REAL:'; To = 'LIVE mode:' }; @{ From = 'Modo SIMULACIÓN'; To = 'SIMULATION mode' }
            @{ From = 'Escribe RESET para continuar'; To = 'Type RESET to continue' }; @{ From = 'Operación cancelada.'; To = 'Operation canceled.' }
            @{ From = 'Primero selecciona un sitio SharePoint o OneDrive.'; To = 'First select a SharePoint or OneDrive site.' }
            @{ From = 'Primero selecciona un destino y al menos una biblioteca.'; To = 'First select a target and at least one library.' }
            @{ From = 'No se encontraron bibliotecas de documentos'; To = 'No document libraries were found' }
            @{ From = 'No se pudieron actualizar las bibliotecas:'; To = 'The libraries could not be updated:' }
            @{ From = 'La selección de bibliotecas fue actualizada sin cambiar de sitio.'; To = 'The library selection was updated without changing the site.' }
            @{ From = 'La operación terminó con errores; revisa el CSV.'; To = 'The operation finished with errors; review the CSV.' }
            @{ From = 'Simulación completada sin modificar permisos.'; To = 'Simulation completed without modifying permissions.' }
            @{ From = 'Restablecimiento completado.'; To = 'Reset completed.' }
            @{ From = 'Dry-run activo: se detectarán los cambios, pero no se modificará ningún permiso.'; To = 'Dry-run is active: changes will be detected, but no permissions will be modified.' }
            @{ From = 'Modo REAL: se eliminarán asignaciones de permisos únicas y los objetos volverán a heredar.'; To = 'LIVE mode: unique permission assignments will be removed and objects will inherit again.' }
            @{ From = 'Seleccionar destino'; To = 'Select target' }; @{ From = 'Cambiar bibliotecas'; To = 'Change libraries' }
            @{ From = 'Registrar una nueva aplicación Entra'; To = 'Register a new Entra application' }
            @{ From = 'Usar una aplicación existente'; To = 'Use an existing application' }
            @{ From = 'Validar aplicación configurada'; To = 'Validate configured application' }
            @{ From = 'Gestionar aplicación Entra / Client ID'; To = 'Manage Entra application / Client ID' }
            @{ From = 'Probar conexión al centro de administración'; To = 'Test connection to the admin center' }
            @{ From = 'Limpiar login persistido de PnP'; To = 'Clear persisted PnP login' }
            @{ From = 'LÍMITE MÁXIMO'; To = 'MAXIMUM LIMIT' }; @{ From = 'SOBRE RECOMENDADO'; To = 'ABOVE RECOMMENDED' }; @{ From = 'DENTRO DE RECOMENDADO'; To = 'WITHIN RECOMMENDED' }
            @{ From = 'Cada elemento de menú debe contener una propiedad Key.'; To = 'Every menu item must contain a Key property.' }
            @{ From = 'Opción inválida.'; To = 'Invalid option.' }; @{ From = 'Valor inválido.'; To = 'Invalid value.' }; @{ From = 'El valor no puede estar vacío.'; To = 'Value cannot be empty.' }
            @{ From = 'Selección inválida.'; To = 'Invalid selection.' }; @{ From = 'Números separados por coma (ej. 1,3,4)'; To = 'Comma-separated numbers (e.g. 1,3,4)' }
            @{ From = 'Selección de bibliotecas'; To = 'Library selection' }; @{ From = 'Marca al menos una biblioteca.'; To = 'Select at least one library.' }
            @{ From = 'UPN inválido.'; To = 'Invalid UPN.' }; @{ From = 'No se pudo determinar la URL del centro de administración.'; To = 'Could not determine the admin center URL.' }
            @{ From = 'No se pudo determinar una URL para inicializar la sesión PnP.'; To = 'Could not determine a URL to initialize the PnP session.' }
            @{ From = 'La aplicación existe y puede autenticarse correctamente.'; To = 'The application exists and can authenticate successfully.' }
            @{ From = 'La aplicación configurada no pudo validarse.'; To = 'The configured application could not be validated.' }
            @{ From = 'El Client ID no se guardó porque la validación falló.'; To = 'The Client ID was not saved because validation failed.' }
            @{ From = 'Se solicitará AllSites.FullControl delegado para las operaciones de este toolkit.'; To = 'Delegated AllSites.FullControl will be requested for this tool.' }
            @{ From = 'PnP no devolvió el Client ID de forma utilizable.'; To = 'PnP did not return a usable Client ID.' }
            @{ From = 'La opción seleccionada no tiene una acción asociada.'; To = 'The selected option has no associated action.' }
            @{ From = 'Filtrar por nombre, URL o propietario (vacío = mostrar todo)'; To = 'Filter by name, URL, or owner (empty = show all)' }
            @{ From = 'Útil para un sitio ya conocido.'; To = 'Useful for a known site.' }
            @{ From = 'Acceso insuficiente. Auto-grant está activado.'; To = 'Insufficient access. Auto-grant is enabled.' }
            @{ From = 'No se pudo identificar la biblioteca principal automáticamente; se muestran todas las bibliotecas detectadas.'; To = 'The main library could not be identified automatically; all detected libraries are shown.' }
            @{ From = 'Actualizando la lista de bibliotecas del sitio actual...'; To = 'Updating the library list for the current site...' }
            @{ From = 'La raíz de la biblioteca principal de OneDrive se omite del total accionable.'; To = 'The root of the main OneDrive library is excluded from the actionable total.' }
            @{ From = 'Raíz de la biblioteca principal de OneDrive omitida.'; To = 'Root of the main OneDrive library omitted.' }
            @{ From = 'ERROR DE ANÁLISIS'; To = 'ANALYSIS ERROR' }; @{ From = 'Resumen por biblioteca'; To = 'Summary by library' }
            @{ From = '{0:N0} por biblioteca'; To = '{0:N0} per library' }
            @{ From = 'Los límites se aplican individualmente a cada lista/biblioteca; el total observado no se compara contra 5,000 o 50,000.'; To = 'Limits apply individually to each list/library; the observed total is not compared against 5,000 or 50,000.' }
            @{ From = 'Todas las bibliotecas analizadas están dentro del límite recomendado.'; To = 'All analyzed libraries are within the recommended limit.' }
            @{ From = 'Cambiar solo las bibliotecas del sitio actual; conserva el sitio seleccionado.'; To = 'Change only the libraries for the current site; keep the selected site.' }
            @{ From = 'Dominio inicial, por ejemplo empresa.onmicrosoft.com'; To = 'Initial domain, for example company.onmicrosoft.com' }
            @{ From = 'Cambiar nombre de app Entra'; To = 'Change Entra app name' }; @{ From = 'Cambiar carpeta de reportes'; To = 'Change report folder' }
            @{ From = 'Alternar persistencia de login'; To = 'Toggle login persistence' }; @{ From = 'Alternar dry-run'; To = 'Toggle dry-run' }
            @{ From = 'Alternar auto-grant de Site Collection Admin'; To = 'Toggle Site Collection Admin auto-grant' }; @{ From = 'Cambiar Admin UPN'; To = 'Change Admin UPN' }
            @{ From = 'Comprueba que el Client ID pueda autenticarse en este tenant.'; To = 'Checks that the Client ID can authenticate in this tenant.' }
            @{ From = 'Introduce y opcionalmente valida un Client ID ya registrado.'; To = 'Enter and optionally validate an already registered Client ID.' }
            @{ From = 'Crea una app PnP nueva y la guarda como configuración activa.'; To = 'Creates a new PnP app and saves it as the active configuration.' }
            @{ From = 'Quitar Client ID de la configuración local'; To = 'Remove Client ID from local configuration' }
            @{ From = 'No elimina la aplicación de Entra; solo deja de usarla en este toolkit.'; To = 'Does not delete the Entra application; it only stops using it in this toolkit.' }
            @{ From = 'Buscar sitios del tenant o introducir una URL.'; To = 'Search for sites in the tenant or enter a URL.' }
            @{ From = 'Buscar OneDrive reales desde el tenant; no se construyen URLs manualmente.'; To = 'Search for real OneDrive sites from the tenant; URLs are not constructed manually.' }
            @{ From = 'Buscar en el tenant'; To = 'Search in the tenant' }; @{ From = 'Introducir URL'; To = 'Enter URL' }
            @{ From = 'Otorgar Site Collection Admin'; To = 'Grant Site Collection Admin' }; @{ From = 'Remover Site Collection Admin'; To = 'Remove Site Collection Admin' }
            @{ From = 'Retirar acceso temporal de Site Collection Admin'; To = 'Remove temporary Site Collection Admin access' }
            @{ From = 'Quita la elevación temporal usada para esta tarea.'; To = 'Removes the temporary elevation used for this task.' }
            @{ From = 'La tarea terminó. Puedes retirar ahora el acceso temporal.'; To = 'The task is finished. You can remove temporary access now.' }
            @{ From = 'El acceso temporal sigue activo.'; To = 'Temporary access is still active.' }
            @{ From = '¿Retirar el acceso temporal ahora?'; To = 'Remove temporary access now?' }
            @{ From = 'No se retiró el acceso temporal.'; To = 'Temporary access was not removed.' }
            @{ From = 'Acceso temporal retirado correctamente.'; To = 'Temporary access removed successfully.' }
            @{ From = 'No se pudo retirar el acceso temporal.'; To = 'Temporary access could not be removed.' }
            @{ From = 'No se pudo obtener una conexión PnP válida.'; To = 'A valid PnP connection could not be obtained.' }
            @{ From = '¿Deseas agregarte temporalmente como Site Collection Admin para continuar?'; To = 'Do you want to add yourself temporarily as Site Collection Admin to continue?' }
            @{ From = 'Auto-grant está activado; se intentará otorgar acceso temporal.'; To = 'Auto-grant is enabled; temporary access will be granted.' }
            @{ From = 'La operación requiere permisos administrativos para este sitio.'; To = 'The operation requires administrative permissions for this site.' }
            @{ From = 'Instalar / actualizar PnP.PowerShell'; To = 'Install / update PnP.PowerShell' }
            @{ From = 'Validar, usar una existente o registrar una nueva.'; To = 'Validate, use an existing one, or register a new one.' }
            @{ From = 'Usa la aplicación actualmente configurada.'; To = 'Uses the currently configured application.' }
            @{ From = 'Cambiar bibliotecas del sitio actual'; To = 'Change libraries for the current site' }
            @{ From = 'Change libraries del sitio actual'; To = 'Change libraries for the current site' }
            @{ From = 'Analizar Unique Permission Scopes'; To = 'Analyze unique permission scopes' }
            @{ From = 'El elemento de menú'; To = 'The menu item' }
            @{ From = 'Hay $($list.Count) resultados. Se muestran los primeros $MaxDisplay; usa un filtro más específico si el elemento no aparece.'; To = '$($list.Count) results found. Showing the first $MaxDisplay; use a more specific filter if the item is not listed.' }
            @{ From = 'PnP.PowerShell no está instalado.'; To = 'PnP.PowerShell is not installed.' }
            @{ From = 'PnP.PowerShell $($pnp.Version) es menor que $script:MinimumPnPVersion.'; To = 'PnP.PowerShell $($pnp.Version) is older than $script:MinimumPnPVersion.' }
            @{ From = 'PnP.PowerShell listo: $((Get-Module PnP.PowerShell).Version)'; To = 'PnP.PowerShell ready: $((Get-Module PnP.PowerShell).Version)' }
            @{ From = ' es menor que '; To = ' is older than ' }; @{ From = ' listo: '; To = ' ready: ' }
            @{ From = 'Client ID inválido: '; To = 'Invalid Client ID: ' }
            @{ From = 'no tiene consentimiento suficiente o el inicio de sesión fue cancelado.'; To = 'does not have sufficient consent or sign-in was canceled.' }
            @{ From = 'Si tu tenant requiere consentimiento administrativo, concédelo antes de usar '; To = 'If your tenant requires administrative consent, grant it before using ' }
            @{ From = 'Autenticación'; To = 'Authentication' }
            @{ From = 'Conectando al centro de administración: $adminUrl'; To = 'Connecting to the admin center: $adminUrl' }
            @{ From = 'El filtro devuelve $($filtered.Count) sitios. Refina la búsqueda para evitar una lista enorme.'; To = 'The filter returns $($filtered.Count) sites. Refine the search to avoid an oversized list.' }
            @{ From = '(sin título)'; To = '(untitled)' }
            @{ From = 'No se pudieron leer las bibliotecas: '; To = 'The libraries could not be read: ' }
            @{ From = 'REST página $page'; To = 'REST page $page' }
            @{ From = 'Página $page · $($items.Count) elementos'; To = 'Page $page · $($items.Count) items' }
            @{ From = '$librariesAtMaximum biblioteca(s) alcanzaron o superaron el límite máximo de $($script:MaximumUniqueScopes) scopes.'; To = '$librariesAtMaximum library/libraries reached or exceeded the maximum limit of $($script:MaximumUniqueScopes) scopes.' }
            @{ From = '$librariesOverRecommended biblioteca(s) superan el límite recomendado de $($script:RecommendedUniqueScopes) scopes.'; To = '$librariesOverRecommended library/libraries exceed the recommended limit of $($script:RecommendedUniqueScopes) scopes.' }
            @{ From = 'LIVE mode: se eliminarán asignaciones de permisos únicas y los objetos volverán a heredar.'; To = 'LIVE mode: unique permission assignments will be removed and objects will inherit again.' }
            @{ From = 'Reset raíz $($library.Title)'; To = 'Reset root $($library.Title)' }
        )
        foreach ($translation in $phrases) { $result = $result.Replace($translation.From, $translation.To) }
        foreach ($translation in $script:UiTranslations) { $result = $result.Replace($translation.From, $translation.To) }
        $common = @(
            @{ From = 'Selecciona'; To = 'Select' }; @{ From = 'Seleccione'; To = 'Select' }; @{ From = 'Introducir'; To = 'Enter' }; @{ From = 'Introduzca'; To = 'Enter' }
            @{ From = 'Buscar'; To = 'Search' }; @{ From = 'Usar'; To = 'Use' }; @{ From = 'Cambiar'; To = 'Change' }; @{ From = 'Guardar'; To = 'Save' }
            @{ From = 'Validar'; To = 'Validate' }; @{ From = 'Registrar'; To = 'Register' }; @{ From = 'Obtener'; To = 'Get' }; @{ From = 'Restablecer'; To = 'Reset' }
            @{ From = 'Analizar'; To = 'Analyze' }; @{ From = 'Análisis'; To = 'Analysis' }; @{ From = 'Procesando'; To = 'Processing' }; @{ From = 'Procesados'; To = 'Processed' }
            @{ From = 'Aplicación'; To = 'Application' }; @{ From = 'Conexión'; To = 'Connection' }; @{ From = 'establecida'; To = 'established' }
            @{ From = 'Sitios'; To = 'Sites' }; @{ From = 'Sitio'; To = 'Site' }; @{ From = 'Bibliotecas'; To = 'Libraries' }; @{ From = 'Biblioteca'; To = 'Library' }
            @{ From = 'Permisos'; To = 'Permissions' }; @{ From = 'herencia'; To = 'inheritance' }; @{ From = 'Herencia'; To = 'Inheritance' }
            @{ From = 'Estado'; To = 'Status' }; @{ From = 'Resultado'; To = 'Result' }; @{ From = 'Resumen'; To = 'Summary' }; @{ From = 'Reporte'; To = 'Report' }
            @{ From = 'Error'; To = 'Error' }; @{ From = 'Errores'; To = 'Errors' }; @{ From = 'Correctos'; To = 'Successful' }; @{ From = 'Límite'; To = 'Limit' }
            @{ From = 'Máximo'; To = 'Maximum' }; @{ From = 'Recomendado'; To = 'Recommended' }; @{ From = 'Operación cancelada'; To = 'Operation canceled' }
            @{ From = 'No hay'; To = 'There are no' }; @{ From = 'No se encontraron'; To = 'No ... were found' }; @{ From = 'No se pudo'; To = 'Could not' }
        )
        foreach ($translation in $common) { $result = $result.Replace($translation.From, $translation.To) }
        $words = @(
            @{ From = 'al'; To = 'when' }; @{ From = 'alguna'; To = 'some' }; @{ From = 'actual'; To = 'current' }; @{ From = 'administrador'; To = 'administrator' }
            @{ From = 'aplicación'; To = 'application' }; @{ From = 'aplicaciones'; To = 'applications' }; @{ From = 'archivo'; To = 'file' }; @{ From = 'archivos'; To = 'files' }
            @{ From = 'analizar'; To = 'analyze' }; @{ From = 'análisis'; To = 'analysis' }; @{ From = 'buscar'; To = 'search' }; @{ From = 'cambiar'; To = 'change' }
            @{ From = 'con'; To = 'with' }; @{ From = 'conexión'; To = 'connection' }; @{ From = 'configuración'; To = 'configuration' }; @{ From = 'contenido'; To = 'content' }
            @{ From = 'después'; To = 'after' }; @{ From = 'dentro'; To = 'within' }; @{ From = 'eliminar'; To = 'remove' }; @{ From = 'en'; To = 'in' }
            @{ From = 'estado'; To = 'status' }; @{ From = 'elemento'; To = 'item' }; @{ From = 'elementos'; To = 'items' }; @{ From = 'encontrados'; To = 'found' }
            @{ From = 'introducir'; To = 'enter' }; @{ From = 'límite'; To = 'limit' }; @{ From = 'lista'; To = 'list' }; @{ From = 'máximo'; To = 'maximum' }
            @{ From = 'mínimo'; To = 'minimum' }; @{ From = 'nombre'; To = 'name' }; @{ From = 'nueva'; To = 'new' }; @{ From = 'opción'; To = 'option' }
            @{ From = 'para'; To = 'for' }; @{ From = 'permiso'; To = 'permission' }; @{ From = 'permisos'; To = 'permissions' }; @{ From = 'procesando'; To = 'processing' }
            @{ From = 'reporte'; To = 'report' }; @{ From = 'reportes'; To = 'reports' }; @{ From = 'resultado'; To = 'result' }; @{ From = 'seleccionar'; To = 'select' }
            @{ From = 'selecciona'; To = 'select' }; @{ From = 'seleccionadas'; To = 'selected' }; @{ From = 'seleccionados'; To = 'selected' }; @{ From = 'selección'; To = 'selection' }
            @{ From = 'sitio'; To = 'site' }; @{ From = 'sitios'; To = 'sites' }; @{ From = 'sin'; To = 'without' }; @{ From = 'total'; To = 'total' }
            @{ From = 'únicos'; To = 'unique' }; @{ From = 'válido'; To = 'valid' }; @{ From = 'válida'; To = 'valid' }; @{ From = 'y'; To = 'and' }; @{ From = 'u'; To = 'or' }
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

# -----------------------------------------------------------------------------
# TUI
# -----------------------------------------------------------------------------

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
    param([ValidateSet('Reset','Bold','Dim','Cyan','Blue','Green','Yellow','Red','Magenta','Gray')][string]$Name)

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
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')][string]$Style = 'Normal',
        [switch]$NoNewline
    )

    $prefix = ''
    $suffix = ''
    $fallback = 'Gray'

    switch ($Style) {
        'Normal'  { $prefix = ''; $fallback = 'Gray' }
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
        $width = [Math]::Min(76, [Math]::Max(28, $Host.UI.RawUI.WindowSize.Width - 2))
    }
    catch { $width = 72 }

    Clear-Host
    Write-Styled "$script:AppName  v$script:AppVersion" Primary
    if ($Context) { Write-Styled $Context Muted }
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
    param([string]$Name, [AllowNull()]$Value, [string]$State = '')

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { $Value = '—' }
    $left = ('{0,-14}' -f $Name)
    Write-Styled $left Muted -NoNewline
    Write-Host "$Value$State"
}

function Pause-Tui {
    Write-Host ''
    [void](Read-Host 'Enter para continuar')
}

function Get-ScopeStatus {
    param(
        [Parameter(Mandatory)]
        [int]$Count
    )

    if ($Count -ge $script:MaximumUniqueScopes) {
        return [pscustomobject]@{
            Label = 'LÍMITE MÁXIMO'
            Style = 'Danger'
        }
    }

    if ($Count -gt $script:RecommendedUniqueScopes) {
        return [pscustomobject]@{
            Label = 'SOBRE RECOMENDADO'
            Style = 'Warning'
        }
    }

    return [pscustomobject]@{
        Label = 'DENTRO DE RECOMENDADO'
        Style = 'Success'
    }
}

function Write-ScopeMeter {
    param(
        [Parameter(Mandatory)]
        [int]$Count,

        [int]$Width = 32
    )

    $recommended = [Math]::Max(1, [int]$script:RecommendedUniqueScopes)
    $maximum = [Math]::Max($recommended, [int]$script:MaximumUniqueScopes)

    $recommendedPct = [Math]::Round(($Count / $recommended) * 100, 1)
    $maximumPct = [Math]::Round(($Count / $maximum) * 100, 1)

    $fillRatio = [Math]::Min(1.0, ($Count / $recommended))
    $filled = [int][Math]::Round($fillRatio * $Width)
    $empty = [Math]::Max(0, $Width - $filled)

    $bar = ('█' * $filled) + ('░' * $empty)
    $state = Get-ScopeStatus -Count $Count

    Write-Styled $bar $state.Style
    Write-Field 'Detectados' $Count
    Write-Field 'Recomendado' ("{0:N0}" -f $recommended)
    Write-Field 'Máximo' ("{0:N0}" -f $maximum)
    Write-Field 'Uso recomendado' ("$recommendedPct%")
    Write-Field 'Uso máximo' ("$maximumPct%")
    Write-Styled ("Estado         {0}" -f $state.Label) $state.Style
}

function Read-MenuChoice {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Items,
        [string]$Prompt = 'Selecciona una opción',
        [switch]$AllowBack
    )

    $normalizedItems = @(
        foreach ($item in $Items) {
            $key = ''
            $label = ''
            $description = ''

            if ($item -is [System.Collections.IDictionary]) {
                if ($item.Contains('Key'))         { $key = [string]$item['Key'] }
                if ($item.Contains('Label'))       { $label = [string]$item['Label'] }
                if ($item.Contains('Description')) { $description = [string]$item['Description'] }
            }
            else {
                $keyProperty = $item.PSObject.Properties['Key']
                $labelProperty = $item.PSObject.Properties['Label']
                $descriptionProperty = $item.PSObject.Properties['Description']

                if ($null -ne $keyProperty)         { $key = [string]$keyProperty.Value }
                if ($null -ne $labelProperty)       { $label = [string]$labelProperty.Value }
                if ($null -ne $descriptionProperty) { $description = [string]$descriptionProperty.Value }
            }

            if ([string]::IsNullOrWhiteSpace($key)) {
                throw 'Cada elemento de menú debe contener una propiedad Key.'
            }

            if ([string]::IsNullOrWhiteSpace($label)) {
                throw "El elemento de menú '$key' debe contener una propiedad Label."
            }

            [pscustomobject]@{
                Key = $key
                Label = $label
                Description = $description
            }
        }
    )

    foreach ($item in $normalizedItems) {
        Write-Styled ("{0,3}" -f $item.Key) Accent -NoNewline
        Write-Host "  $($item.Label)"

        if (-not [string]::IsNullOrWhiteSpace($item.Description)) {
            Write-Styled ("     $($item.Description)") Muted
        }
    }

    $valid = @($normalizedItems | ForEach-Object { [string]$_.Key })

    if ($AllowBack -and -not ($valid -contains '0')) {
        Write-Styled '  0' Accent -NoNewline
        Write-Host '  Volver'
        $valid += '0'
    }

    while ($true) {
        Write-Host ''
        $choice = (Read-Host $Prompt).Trim()

        if ($valid -contains $choice) {
            return $choice
        }

        Write-Status Error 'Opción inválida.'
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
        $value = Read-Host $caption

        if ([string]::IsNullOrWhiteSpace($value) -and $Default) { $value = $Default }
        if ([string]::IsNullOrWhiteSpace($value)) {
            if ($AllowEmpty) { return '' }
            Write-Status Error 'El valor no puede estar vacío.'
            continue
        }

        $value = $value.Trim()
        if ($Validator -and -not (& $Validator $value)) {
            Write-Status Error $ValidationMessage
            continue
        }

        return $value
    }
}

function Read-YesNo {
    param([Parameter(Mandatory)][string]$Prompt, [bool]$Default = $false)

    $hint = if ($Default) { 'S/n' } else { 's/N' }
    while ($true) {
        $answer = (Read-Host "$Prompt [$hint]").Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        if ($answer -match '^(s|si|sí|y|yes)$') { return $true }
        if ($answer -match '^(n|no)$') { return $false }
        Write-Status Error 'Responde S o N.'
    }
}

function Select-SingleByNumber {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Items,
        [Parameter(Mandatory)][scriptblock]$Label,
        [string]$Title = 'Seleccionar',
        [int]$MaxDisplay = 50
    )

    $list = @($Items)
    if ($list.Count -eq 0) { return $null }

    if ($list.Count -gt $MaxDisplay) {
        Write-Status Warn "Hay $($list.Count) resultados. Se muestran los primeros $MaxDisplay; usa un filtro más específico si el elemento no aparece."
        $list = @($list | Select-Object -First $MaxDisplay)
    }

    Write-Section $Title
    for ($i = 0; $i -lt $list.Count; $i++) {
        Write-Styled ("{0,3}" -f ($i + 1)) Accent -NoNewline
        Write-Host ('  ' + (& $Label $list[$i]))
    }
    Write-Styled '  0' Accent -NoNewline
    Write-Host '  Cancelar'

    while ($true) {
        Write-Host ''
        $raw = Read-Host 'Número'
        $number = 0
        if ([int]::TryParse($raw, [ref]$number)) {
            if ($number -eq 0) { return $null }
            if ($number -ge 1 -and $number -le $list.Count) { return $list[$number - 1] }
        }
        Write-Status Error 'Selección inválida.'
    }
}

function Get-LibraryFlagsText {
    param([Parameter(Mandatory)]$Library)

    $flags = [System.Collections.Generic.List[string]]::new()

    if ($Library.IsDefaultDocumentLibrary) { $flags.Add('PRINCIPAL') }
    if ($Library.IsSystemList)              { $flags.Add('SISTEMA') }
    if ($Library.Hidden)                    { $flags.Add('OCULTA') }

    if ($flags.Count -eq 0) { return '' }
    return '  [' + ($flags -join ' · ') + ']'
}

function Select-LibrariesInteractive {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Libraries,

        [AllowEmptyCollection()]
        [string[]]$PreselectedIds = @()
    )

    $libs = @($Libraries)
    if ($libs.Count -eq 0) { return @() }

    $rawUiAvailable = $false
    try {
        $null = $Host.UI.RawUI.WindowSize
        $rawUiAvailable = $true
    }
    catch { $rawUiAvailable = $false }

    if (-not $rawUiAvailable -or [Console]::IsInputRedirected) {
        Write-Section 'Bibliotecas'
        for ($i = 0; $i -lt $libs.Count; $i++) {
            $flags = Get-LibraryFlagsText -Library $libs[$i]
            Write-Styled ("{0,3}" -f ($i + 1)) Accent -NoNewline
            Write-Host "  $($libs[$i].Title)$flags"
            Write-Styled "     $($libs[$i].ItemCount) elementos · $($libs[$i].RootFolderUrl)" Muted
        }
        Write-Host ''
        $raw = Read-Host 'Números separados por coma (ej. 1,3,4)'
        $indexes = @($raw -split ',' | ForEach-Object {
            $n = 0
            if ([int]::TryParse($_.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $libs.Count) { $n - 1 }
        } | Select-Object -Unique)
        return @($indexes | ForEach-Object { $libs[$_] })
    }

    $selected = [System.Collections.Generic.HashSet[int]]::new()
    $cursor = 0
    $hasPreselection = $false

    # Al volver a seleccionar bibliotecas conservamos visualmente la selección
    # actual. Los IDs son GUID de SharePoint, por lo que no dependemos del nombre.
    if ($PreselectedIds.Count -gt 0) {
        for ($i = 0; $i -lt $libs.Count; $i++) {
            if ($PreselectedIds -contains [string]$libs[$i].Id) {
                [void]$selected.Add($i)
                if (-not $hasPreselection) {
                    $cursor = $i
                    $hasPreselection = $true
                }
            }
        }
    }

    # En una selección inicial, o si ninguna selección anterior sigue existiendo,
    # marcamos la biblioteca principal como opción conveniente por defecto.
    if (-not $hasPreselection) {
        for ($i = 0; $i -lt $libs.Count; $i++) {
            if ($libs[$i].IsDefaultDocumentLibrary) {
                [void]$selected.Add($i)
                $cursor = $i
                break
            }
        }
    }

    while ($true) {
        Write-AppHeader 'Selección de bibliotecas'
        Write-Styled '↑/↓ mover   Espacio marcar   A todas   N ninguna   Enter aceptar   Esc cancelar' Muted
        Write-Host ''

        for ($i = 0; $i -lt $libs.Count; $i++) {
            $pointer = if ($i -eq $cursor) { '›' } else { ' ' }
            $mark = if ($selected.Contains($i)) { '●' } else { '○' }
            $flags = Get-LibraryFlagsText -Library $libs[$i]
            $line = "$pointer $mark  $($libs[$i].Title)$flags"

            if ($i -eq $cursor) { Write-Styled $line Primary }
            elseif ($selected.Contains($i)) { Write-Styled $line Success }
            else { Write-Host $line }

            Write-Styled ("      {0} elementos · {1}" -f $libs[$i].ItemCount, $libs[$i].RootFolderUrl) Muted
        }

        Write-Host ''
        Write-Styled ("Seleccionadas: {0}/{1}" -f $selected.Count, $libs.Count) Accent

        $key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        switch ($key.VirtualKeyCode) {
            38 { if ($cursor -gt 0) { $cursor-- } }
            40 { if ($cursor -lt ($libs.Count - 1)) { $cursor++ } }
            32 {
                if ($selected.Contains($cursor)) { [void]$selected.Remove($cursor) }
                else { [void]$selected.Add($cursor) }
            }
            65 {
                $selected.Clear()
                for ($i = 0; $i -lt $libs.Count; $i++) { [void]$selected.Add($i) }
            }
            78 { $selected.Clear() }
            13 {
                if ($selected.Count -eq 0) {
                    Write-Status Warn 'Marca al menos una biblioteca.'
                    Start-Sleep -Milliseconds 700
                    continue
                }
                $result = foreach ($index in ($selected | Sort-Object)) { $libs[$index] }
                return @($result)
            }
            27 { return @() }
        }
    }
}

# -----------------------------------------------------------------------------
# Settings
# -----------------------------------------------------------------------------

function New-DefaultSettings {
    $clientId = ''
    if ($env:ENTRAID_APP_ID) { $clientId = $env:ENTRAID_APP_ID }
    elseif ($env:ENTRAID_CLIENT_ID) { $clientId = $env:ENTRAID_CLIENT_ID }

    [pscustomobject]@{
        Tenant          = ''
        ClientId        = $clientId
        AppRegistrationName = 'M365 Permissions Scope Manager'
        ReportFolder    = $script:DefaultReportFolder
        PersistLogin    = $true
        Language        = 'es'
        DryRun          = $true
        AutoGrantAdmin  = $false
        AdminUpn        = ''
        PageSize        = $script:DefaultPageSize
        MaxRetries      = $script:DefaultMaxRetries
        RetryDelay      = $script:DefaultRetryDelay
    }
}

function Merge-Settings {
    param($Loaded)

    $defaults = New-DefaultSettings
    if (-not $Loaded) { return $defaults }

    foreach ($p in $defaults.PSObject.Properties.Name) {
        if (-not $Loaded.PSObject.Properties[$p]) {
            Add-Member -InputObject $Loaded -NotePropertyName $p -NotePropertyValue $defaults.$p
        }
    }

    return $Loaded
}

function Load-Settings {
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) { return New-DefaultSettings }

    try {
        $loaded = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        return Merge-Settings $loaded
    }
    catch {
        Write-Status Warn "No se pudo leer settings.json: $($_.Exception.Message)"
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
}

function Show-SettingsMenu {
    while ($true) {
        Write-AppHeader 'Configuración'
        Write-Field 'Tenant' $script:Settings.Tenant
        Write-Field 'Client ID' $script:Settings.ClientId
        Write-Field 'App' $script:Settings.AppRegistrationName
        Write-Field 'Reportes' $script:Settings.ReportFolder
        Write-Field 'Idioma' $(if ($script:Language -eq 'en') { 'English' } else { 'Español' })
        Write-Field 'Persist login' $(if ($script:Settings.PersistLogin) { 'Sí' } else { 'No' })
        Write-Field 'Dry-run' $(if ($script:Settings.DryRun) { 'Sí' } else { 'No' })
        Write-Field 'Auto admin' $(if ($script:Settings.AutoGrantAdmin) { 'Sí' } else { 'No' })
        Write-Field 'Admin UPN' $script:Settings.AdminUpn
        Write-Host ''

        $choice = Read-MenuChoice -AllowBack -Items @(
            @{ Key='1'; Label='Cambiar tenant'; Description='Dominio inicial, por ejemplo empresa.onmicrosoft.com' },
            @{ Key='2'; Label='Cambiar Client ID' },
            @{ Key='3'; Label='Cambiar nombre de app Entra' },
            @{ Key='4'; Label='Cambiar carpeta de reportes' },
            @{ Key='5'; Label='Alternar persistencia de login' },
            @{ Key='6'; Label='Alternar dry-run' },
            @{ Key='7'; Label='Alternar auto-grant de Site Collection Admin' },
            @{ Key='8'; Label='Cambiar Admin UPN' },
            @{ Key='9'; Label='Cambiar idioma' }
        )

        switch ($choice) {
            '0' { Save-Settings; return }
            '1' {
                $script:Settings.Tenant = Read-TextValue -Prompt 'Tenant' -Default $script:Settings.Tenant `
                    -Validator { param($v) $v -match '^[A-Za-z0-9-]+\.onmicrosoft\.com$' } `
                    -ValidationMessage 'Usa el dominio inicial: nombre.onmicrosoft.com'
                $script:Target = $null
            }
            '2' {
                $script:Settings.ClientId = Read-TextValue -Prompt 'Client ID' -Default $script:Settings.ClientId `
                    -Validator { param($v) $v -match '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$' } `
                    -ValidationMessage 'Debe ser un GUID válido.'
            }
            '3' { $script:Settings.AppRegistrationName = Read-TextValue -Prompt 'Nombre de app' -Default $script:Settings.AppRegistrationName }
            '4' { $script:Settings.ReportFolder = Read-TextValue -Prompt 'Carpeta de reportes' -Default $script:Settings.ReportFolder }
            '5' { $script:Settings.PersistLogin = -not $script:Settings.PersistLogin }
            '6' { $script:Settings.DryRun = -not $script:Settings.DryRun }
            '7' { $script:Settings.AutoGrantAdmin = -not $script:Settings.AutoGrantAdmin }
            '8' {
                $script:Settings.AdminUpn = Read-TextValue -Prompt 'Admin UPN' -Default $script:Settings.AdminUpn -AllowEmpty `
                    -Validator { param($v) [string]::IsNullOrWhiteSpace($v) -or $v -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' } `
                    -ValidationMessage 'UPN inválido.'
            }
            '9' {
                Initialize-AppLanguage
                $script:Settings.Language = $script:Language
            }
        }
        Save-Settings
    }
}

# -----------------------------------------------------------------------------
# Environment / PnP
# -----------------------------------------------------------------------------

function Test-PowerShellVersion {
    if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
        Write-Status Error "PowerShell $($PSVersionTable.PSVersion) detectado. PnP.PowerShell 3.x requiere PowerShell 7.4+."
        return $false
    }
    return $true
}

function Get-InstalledPnPModule {
    Get-Module -Name PnP.PowerShell -ListAvailable |
        Sort-Object Version -Descending |
        Select-Object -First 1
}

function Ensure-PnPModule {
    param([switch]$OfferInstall)

    if (-not (Test-PowerShellVersion)) { return $false }

    $pnp = Get-InstalledPnPModule
    if ($pnp -and [version]$pnp.Version -ge $script:MinimumPnPVersion) {
        Import-Module PnP.PowerShell -MinimumVersion $script:MinimumPnPVersion -ErrorAction Stop
        return $true
    }

    if (-not $OfferInstall) {
        if (-not $pnp) { Write-Status Error 'PnP.PowerShell no está instalado.' }
        else { Write-Status Error "PnP.PowerShell $($pnp.Version) es menor que $script:MinimumPnPVersion." }
        return $false
    }

    Write-Status Warn 'PnP.PowerShell debe instalarse o actualizarse.'
    if (-not (Read-YesNo '¿Instalar/actualizar ahora?' $true)) { return $false }

    try {
        if ($pnp) {
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -MinimumVersion $script:MinimumPnPVersion
        }
        else {
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -MinimumVersion $script:MinimumPnPVersion
        }
        Import-Module PnP.PowerShell -MinimumVersion $script:MinimumPnPVersion -Force
        Write-Status Ok "PnP.PowerShell listo: $((Get-Module PnP.PowerShell).Version)"
        return $true
    }
    catch {
        Write-Status Error "No se pudo instalar/importar PnP.PowerShell: $($_.Exception.Message)"
        return $false
    }
}

function Ensure-TenantConfigured {
    if (-not [string]::IsNullOrWhiteSpace($script:Settings.Tenant)) { return $true }

    Write-Section 'Tenant'
    $script:Settings.Tenant = Read-TextValue -Prompt 'Dominio inicial del tenant (ej. empresa.onmicrosoft.com)' `
        -Validator { param($v) $v -match '^[A-Za-z0-9-]+\.onmicrosoft\.com$' } `
        -ValidationMessage 'Formato esperado: nombre.onmicrosoft.com'
    Save-Settings
    return $true
}

function Get-TenantPrefix {
    if (-not (Ensure-TenantConfigured)) { return $null }
    return ($script:Settings.Tenant -split '\.')[0]
}

function Get-AdminUrl {
    $prefix = Get-TenantPrefix
    if (-not $prefix) { return $null }
    return "https://$prefix-admin.sharepoint.com"
}

function Test-ClientIdFormat {
    param(
        [AllowNull()]
        [string]$ClientId
    )

    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        return $false
    }

    return $ClientId -match (
        '^[0-9A-Fa-f]{8}-' +
        '[0-9A-Fa-f]{4}-' +
        '[0-9A-Fa-f]{4}-' +
        '[0-9A-Fa-f]{4}-' +
        '[0-9A-Fa-f]{12}$'
    )
}

function Connect-M365SiteWithClientId {
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$ClientId
    )

    if (-not (Ensure-PnPModule -OfferInstall)) {
        return $null
    }

    if (-not (Ensure-TenantConfigured)) {
        return $null
    }

    if (-not (Test-ClientIdFormat -ClientId $ClientId)) {
        throw "Client ID inválido: '$ClientId'."
    }

    $params = @{
        Url              = $Url
        ClientId         = $ClientId
        Tenant           = $script:Settings.Tenant
        Interactive      = $true
        ReturnConnection = $true
        ErrorAction      = 'Stop'
    }

    if ($script:Settings.PersistLogin) {
        $params.PersistLogin = $true
    }

    return Connect-PnPOnline @params
}

function Test-PnPAppRegistration {
    param(
        [Parameter(Mandatory)]
        [string]$ClientId,

        [switch]$Quiet
    )

    if (-not (Test-ClientIdFormat -ClientId $ClientId)) {
        if (-not $Quiet) {
            Write-Status Error 'El Client ID no tiene formato GUID válido.'
        }
        return $false
    }

    try {
        $testUrl = Get-AdminUrl

        if ([string]::IsNullOrWhiteSpace($testUrl)) {
            throw 'No se pudo determinar la URL del centro de administración.'
        }

        if (-not $Quiet) {
            Write-Status Info "Validando app $ClientId contra $testUrl ..."
            Write-Styled 'Puede aparecer una ventana de autenticación de Microsoft.' Muted
        }

        $conn = Connect-M365SiteWithClientId `
            -Url $testUrl `
            -ClientId $ClientId

        if ($null -eq $conn) {
            throw 'No se obtuvo una conexión PnP.'
        }

        $web = Get-PnPWeb `
            -Connection $conn `
            -Includes Title,Url `
            -ErrorAction Stop

        if (-not $Quiet) {
            Write-Status Ok 'La aplicación existe y puede autenticarse correctamente.'
            Write-Field 'Client ID' $ClientId
            Write-Field 'Destino' $web.Url
        }

        return $true
    }
    catch {
        if (-not $Quiet) {
            Write-Status Error 'La aplicación configurada no pudo validarse.'
            Write-Styled $_.Exception.Message Danger
            Write-Host ''
            Write-Styled (
                'Esto puede indicar que la app fue eliminada, pertenece a otro tenant, ' +
                'no tiene consentimiento suficiente o el inicio de sesión fue cancelado.'
            ) Muted
        }

        return $false
    }
}

function Set-ExistingPnPClientId {
    Write-Section 'Usar aplicación existente'

    $currentDefault = if (Test-ClientIdFormat -ClientId $script:Settings.ClientId) {
        $script:Settings.ClientId
    }
    else {
        ''
    }

    $clientId = Read-TextValue `
        -Prompt 'Client ID de la aplicación existente' `
        -Default $currentDefault `
        -Validator {
            param($v)
            Test-ClientIdFormat -ClientId $v
        } `
        -ValidationMessage 'Debe ser un GUID válido.'

    $validate = Read-YesNo `
        -Prompt '¿Validar esta aplicación ahora?' `
        -Default $true

    if ($validate) {
        if (-not (Test-PnPAppRegistration -ClientId $clientId)) {
            Write-Status Warn 'El Client ID no se guardó porque la validación falló.'
            return $false
        }
    }

    $script:Settings.ClientId = $clientId
    Save-Settings

    Write-Status Ok "Client ID guardado: $clientId"
    return $true
}

function Register-NewPnPApp {
    if (-not (Ensure-PnPModule -OfferInstall)) {
        return $false
    }

    if (-not (Ensure-TenantConfigured)) {
        return $false
    }

    Write-Section 'Registrar nueva aplicación Entra'

    $defaultName = $script:Settings.AppRegistrationName

    if ([string]::IsNullOrWhiteSpace($defaultName)) {
        $defaultName = 'M365 Permissions Scope Manager'
    }

    $appName = Read-TextValue `
        -Prompt 'Nombre de la nueva aplicación' `
        -Default $defaultName

    Write-Status Info 'Iniciando registro de la aplicación Entra...'
    Write-Styled (
        'Se solicitará AllSites.FullControl delegado para las operaciones de este toolkit.'
    ) Muted

    try {
        $result = Register-PnPEntraIDAppForInteractiveLogin `
            -ApplicationName $appName `
            -Tenant $script:Settings.Tenant `
            -SharePointDelegatePermissions 'AllSites.FullControl' `
            -ErrorAction Stop

        $candidate = $null

        if ($null -ne $result) {
            foreach ($resultItem in ($result | ForEach-Object { $_ })) {
                foreach ($name in @('ClientId','AppId','ApplicationId','Id')) {
                    $property = $resultItem.PSObject.Properties[$name]

                    if ($null -ne $property -and $property.Value) {
                        $possible = [string]$property.Value

                        if (Test-ClientIdFormat -ClientId $possible) {
                            $candidate = $possible
                            break
                        }
                    }
                }

                if (Test-ClientIdFormat -ClientId $candidate) {
                    break
                }
            }
        }

        if (-not (Test-ClientIdFormat -ClientId $candidate)) {
            Write-Status Warn 'PnP no devolvió el Client ID de forma utilizable.'

            $candidate = Read-TextValue `
                -Prompt 'Pega el Application (client) ID creado' `
                -Validator {
                    param($v)
                    Test-ClientIdFormat -ClientId $v
                } `
                -ValidationMessage 'Debe ser un GUID válido.'
        }

        $script:Settings.ClientId = $candidate
        $script:Settings.AppRegistrationName = $appName
        Save-Settings

        Write-Status Ok "Nueva aplicación registrada: $appName"
        Write-Field 'Client ID' $candidate
        Write-Host ''
        Write-Styled (
            'Si tu tenant requiere consentimiento administrativo, concédelo antes de usar ' +
            'operaciones que necesiten esos permisos.'
        ) Muted

        if (Read-YesNo -Prompt '¿Validar la nueva aplicación ahora?' -Default $true) {
            [void](Test-PnPAppRegistration -ClientId $candidate)
        }

        return $true
    }
    catch {
        Write-Status Error "No se pudo registrar la app: $($_.Exception.Message)"
        return $false
    }
}

function Show-AppRegistrationMenu {
    while ($true) {
        Write-AppHeader 'Aplicación Entra / autenticación PnP'

        $configured = Test-ClientIdFormat -ClientId $script:Settings.ClientId

        Write-Field 'Tenant' $script:Settings.Tenant
        Write-Field 'App guardada' $script:Settings.AppRegistrationName
        Write-Field 'Client ID' $script:Settings.ClientId
        Write-Field 'Estado config' $(if ($configured) { 'CONFIGURADA' } else { 'NO CONFIGURADA' })

        Write-Host ''

        # Construimos el menú con numeración continua y asociamos cada número
        # a una acción lógica. Así nunca aparecen opciones 2/3 sin existir la 1.
        $items = @()
        $actionMap = @{}
        $nextKey = 1

        if ($configured) {
            $key = [string]$nextKey
            $items += @{
                Key         = $key
                Label       = 'Validar aplicación configurada'
                Description = 'Comprueba que el Client ID pueda autenticarse en este tenant.'
            }
            $actionMap[$key] = 'Validate'
            $nextKey++
        }

        $key = [string]$nextKey
        $items += @{
            Key         = $key
            Label       = 'Usar una aplicación existente'
            Description = 'Introduce y opcionalmente valida un Client ID ya registrado.'
        }
        $actionMap[$key] = 'UseExisting'
        $nextKey++

        $key = [string]$nextKey
        $items += @{
            Key         = $key
            Label       = 'Registrar una nueva aplicación Entra'
            Description = 'Crea una app PnP nueva y la guarda como configuración activa.'
        }
        $actionMap[$key] = 'RegisterNew'
        $nextKey++

        if ($configured) {
            $key = [string]$nextKey
            $items += @{
                Key         = $key
                Label       = 'Quitar Client ID de la configuración local'
                Description = 'No elimina la aplicación de Entra; solo deja de usarla en este toolkit.'
            }
            $actionMap[$key] = 'RemoveLocal'
            $nextKey++
        }

        $choice = Read-MenuChoice `
            -AllowBack `
            -Items $items

        if ($choice -eq '0') {
            return
        }

        if (-not $actionMap.ContainsKey($choice)) {
            Write-Status Error 'La opción seleccionada no tiene una acción asociada.'
            Pause-Tui
            continue
        }

        switch ($actionMap[$choice]) {
            'Validate' {
                [void](Test-PnPAppRegistration -ClientId $script:Settings.ClientId)
                Pause-Tui
            }

            'UseExisting' {
                [void](Set-ExistingPnPClientId)
                Pause-Tui
            }

            'RegisterNew' {
                [void](Register-NewPnPApp)
                Pause-Tui
            }

            'RemoveLocal' {
                if (
                    Read-YesNo `
                        -Prompt '¿Quitar el Client ID guardado de este toolkit?' `
                        -Default $false
                ) {
                    $script:Settings.ClientId = ''
                    Save-Settings
                    Write-Status Ok 'Client ID eliminado de la configuración local.'
                }
                Pause-Tui
            }
        }
    }
}

function Ensure-ClientId {
    if (Test-ClientIdFormat -ClientId $script:Settings.ClientId) {
        return $true
    }

    if (-not (Ensure-PnPModule -OfferInstall)) {
        return $false
    }

    if (-not (Ensure-TenantConfigured)) {
        return $false
    }

    Write-AppHeader 'Autenticación'
    Write-Status Warn 'No hay una aplicación Entra válida configurada.'
    Write-Styled (
        'Puedes usar una aplicación existente o registrar una nueva desde este mismo toolkit.'
    ) Muted

    $choice = Read-MenuChoice `
        -AllowBack `
        -Items @(
            @{
                Key         = '1'
                Label       = 'Usar una aplicación existente'
                Description = 'Introduce un Client ID ya registrado.'
            },
            @{
                Key         = '2'
                Label       = 'Registrar una nueva aplicación Entra'
                Description = 'Crea una aplicación nueva para PnP.PowerShell.'
            }
        )

    switch ($choice) {
        '0' {
            return $false
        }

        '1' {
            return [bool](Set-ExistingPnPClientId)
        }

        '2' {
            return [bool](Register-NewPnPApp)
        }
    }

    return $false
}

function Connect-M365Site {
    param([Parameter(Mandatory)][string]$Url)

    if (-not (Ensure-PnPModule -OfferInstall)) { return $null }
    if (-not (Ensure-TenantConfigured)) { return $null }
    if (-not (Ensure-ClientId)) { return $null }

    $params = @{
        Url        = $Url
        ClientId   = $script:Settings.ClientId
        Tenant     = $script:Settings.Tenant
        ReturnConnection = $true
        ErrorAction = 'Stop'
    }
    if ($script:Settings.PersistLogin) { $params.PersistLogin = $true }

    return Connect-PnPOnline @params
}

function Test-UnauthorizedError {
    param([string]$Message)
    return $Message -match '(?i)unauthori[sz]ed|access denied|forbidden|\b401\b|\b403\b|attempted to perform'
}

function Try-EnableTemporarySiteAdmin {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$ErrorMessage
    )

    if (-not (Test-UnauthorizedError -Message $ErrorMessage)) {
        return $false
    }

    if ($null -ne $script:TemporarySiteAdmin) {
        return ([string]$script:TemporarySiteAdmin.SiteUrl).TrimEnd('/') -eq $SiteUrl.TrimEnd('/')
    }

    if ($script:Settings.AutoGrantAdmin) {
        Write-Status Warn 'Auto-grant está activado; se intentará otorgar acceso temporal.'
    }
    elseif (-not (Read-YesNo -Prompt '¿Deseas agregarte temporalmente como Site Collection Admin para continuar?' -Default $true)) {
        return $false
    }

    if (-not (Grant-SiteAdminAccess -SiteUrl $SiteUrl -TrackTemporary)) {
        return $false
    }

    Write-Status Info 'Esperando propagación del acceso administrativo...'
    Start-Sleep -Seconds 5
    return $true
}

function Connect-M365SiteWithRecovery {
    param(
        [Parameter(Mandatory)][string]$Url,
        [switch]$SkipAutoGrant
    )

    try {
        $connection = Connect-M365Site -Url $Url
        if ($null -eq $connection) {
            throw 'No se pudo obtener una conexión PnP válida.'
        }

        return $connection
    }
    catch {
        if ($SkipAutoGrant -or -not (Try-EnableTemporarySiteAdmin -SiteUrl $Url -ErrorMessage $_.Exception.Message)) {
            throw
        }

        $connection = Connect-M365Site -Url $Url
        if ($null -eq $connection) {
            throw 'No se pudo obtener una conexión PnP válida después de otorgar acceso.'
        }

        return $connection
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
            $msg = $_.Exception.Message
            $retryable = $msg -match '(?i)\b429\b|throttl|too many requests|\b503\b|service unavailable|temporarily unavailable|timeout|timed out'
            if (-not $retryable -or $attempt -ge $maxRetries) { throw }

            $jitter = Get-Random -Minimum 0 -Maximum 3
            $sleep = [Math]::Min(120, $delay + $jitter)
            Write-Status Warn "${OperationName}: reintento $attempt/$maxRetries en ${sleep}s."
            Start-Sleep -Seconds $sleep
            $delay = [Math]::Min(120, $delay * 2)
        }
    }
}

# -----------------------------------------------------------------------------
# Site discovery
# -----------------------------------------------------------------------------

function Get-TenantSites {
    param([ValidateSet('SharePoint','OneDrive')][string]$Kind)

    $adminUrl = Get-AdminUrl
    if (-not $adminUrl) { return @() }

    Write-Status Info "Conectando al centro de administración: $adminUrl"
    $adminConnection = Connect-M365Site -Url $adminUrl
    if (-not $adminConnection) { return @() }

    Write-Status Info 'Consultando sitios del tenant...'
    $sites = @(Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Connection $adminConnection -ErrorAction Stop)

    if ($Kind -eq 'OneDrive') {
        return @($sites | Where-Object { $_.Url -match '-my\.sharepoint\.com/personal/' } | Sort-Object Owner, Url)
    }

    return @($sites | Where-Object {
        $_.Url -notmatch '-my\.sharepoint\.com/personal/' -and
        $_.Url -notmatch '-admin\.sharepoint\.com/?$'
    } | Sort-Object Title, Url)
}

function Find-SiteInteractively {
    param([ValidateSet('SharePoint','OneDrive')][string]$Kind)

    try {
        $sites = @(Get-TenantSites -Kind $Kind)
    }
    catch {
        Write-Status Error "No se pudieron enumerar sitios: $($_.Exception.Message)"
        return $null
    }

    if ($sites.Count -eq 0) {
        Write-Status Warn 'No se encontraron sitios accesibles.'
        return $null
    }

    while ($true) {
        Write-Host ''
        $filter = Read-TextValue -Prompt 'Filtrar por nombre, URL o propietario (vacío = mostrar todo)' -AllowEmpty
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
            Write-Status Warn 'Sin coincidencias. Prueba otro filtro.'
            continue
        }

        if ($filtered.Count -gt 50) {
            Write-Status Warn "El filtro devuelve $($filtered.Count) sitios. Refina la búsqueda para evitar una lista enorme."
            continue
        }

        return Select-SingleByNumber -Items $filtered -Title $(if ($Kind -eq 'OneDrive') { 'OneDrive' } else { 'Sitios SharePoint' }) -Label {
            param($s)
            $label = if ($s.Title) { $s.Title } elseif ($s.Owner) { $s.Owner } else { '(sin título)' }
            $owner = if ($s.Owner) { " · $($s.Owner)" } else { '' }
            "$label$owner`n     $($s.Url)"
        }
    }
}

function Get-DefaultDocumentLibraryInfo {
    param([Parameter(Mandatory)]$Connection)

    try {
        $ctx = Get-PnPContext -Connection $Connection
        if ($null -eq $ctx) { return $null }

        $defaultLibrary = $ctx.Web.DefaultDocumentLibrary()
        if ($null -eq $defaultLibrary) { return $null }

        $ctx.Load($defaultLibrary)
        $ctx.Load($defaultLibrary.RootFolder)
        $ctx.ExecuteQuery()

        return [pscustomobject]@{
            Id            = [string]$defaultLibrary.Id
            Title         = [string]$defaultLibrary.Title
            RootFolderUrl = [string]$defaultLibrary.RootFolder.ServerRelativeUrl
        }
    }
    catch {
        return $null
    }
}

function Resolve-Libraries {
    param(
        [Parameter(Mandatory)]$Connection,
        [ValidateSet('SharePoint','OneDrive')][string]$TargetKind = 'SharePoint'
    )

    $defaultLibraryInfo = $null
    if ($TargetKind -eq 'OneDrive') {
        $defaultLibraryInfo = Get-DefaultDocumentLibraryInfo -Connection $Connection
    }

    $lists = @(Get-PnPList -Connection $Connection `
        -Includes RootFolder,HasUniqueRoleAssignments,ItemCount,BaseTemplate,BaseType,Hidden,IsSystemList `
        -ErrorAction Stop |
        Where-Object { [string]$_.BaseType -eq 'DocumentLibrary' })

    $result = foreach ($list in $lists) {
        $root = Get-PnPProperty -ClientObject $list -Property RootFolder -Connection $Connection
        $unique = Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments -Connection $Connection
        $listId = [string]$list.Id
        $isDefault = $false

        if ($null -ne $defaultLibraryInfo -and -not [string]::IsNullOrWhiteSpace($defaultLibraryInfo.Id)) {
            $isDefault = ($listId -eq [string]$defaultLibraryInfo.Id)
        }

        [pscustomobject]@{
            Id                       = $listId
            Title                    = [string]$list.Title
            RootFolderUrl            = [string]$root.ServerRelativeUrl
            ItemCount                = [int]$list.ItemCount
            Hidden                   = [bool]$list.Hidden
            IsSystemList             = [bool]$list.IsSystemList
            BaseTemplate             = [int]$list.BaseTemplate
            BaseType                 = [string]$list.BaseType
            HasUniqueRoleAssignments = [bool]$unique
            IsDefaultDocumentLibrary = $isDefault
        }
    }

    return @($result | Sort-Object `
        @{ Expression = {
            if ($_.IsDefaultDocumentLibrary) { 0 }
            elseif (-not $_.Hidden -and -not $_.IsSystemList) { 1 }
            elseif ($_.IsSystemList) { 2 }
            else { 3 }
        }}, Title)
}

function Select-Target {
    Write-AppHeader 'Seleccionar destino'

    if (-not (Ensure-PnPModule -OfferInstall)) { Pause-Tui; return }
    if (-not (Ensure-TenantConfigured)) { Pause-Tui; return }
    if (-not (Ensure-ClientId)) { Pause-Tui; return }

    $choice = Read-MenuChoice -AllowBack -Items @(
        @{ Key='1'; Label='SharePoint Online'; Description='Buscar sitios del tenant o introducir una URL.' },
        @{ Key='2'; Label='OneDrive for Business'; Description='Buscar OneDrive reales desde el tenant; no se construyen URLs manualmente.' },
        @{ Key='3'; Label='URL directa'; Description='Útil para un sitio ya conocido.' }
    )
    if ($choice -eq '0') { return }

    $site = $null
    $kind = 'SharePoint'

    switch ($choice) {
        '1' {
            $mode = Read-MenuChoice -AllowBack -Items @(
                @{ Key='1'; Label='Buscar en el tenant' },
                @{ Key='2'; Label='Introducir URL' }
            )
            if ($mode -eq '0') { return }
            if ($mode -eq '1') { $site = Find-SiteInteractively -Kind SharePoint }
            else {
                $url = Read-TextValue -Prompt 'URL completa del sitio' `
                    -Validator { param($v) [uri]::IsWellFormedUriString($v, [UriKind]::Absolute) -and $v -match '^https://' } `
                    -ValidationMessage 'Introduce una URL HTTPS válida.'
                $site = [pscustomobject]@{ Url=$url; Title='Sitio SharePoint'; Owner='' }
            }
        }
        '2' {
            $kind = 'OneDrive'
            $site = Find-SiteInteractively -Kind OneDrive
        }
        '3' {
            $url = Read-TextValue -Prompt 'URL completa del sitio / OneDrive' `
                -Validator { param($v) [uri]::IsWellFormedUriString($v, [UriKind]::Absolute) -and $v -match '^https://' } `
                -ValidationMessage 'Introduce una URL HTTPS válida.'
            if ($url -match '-my\.sharepoint\.com/personal/') { $kind = 'OneDrive' }
            $site = [pscustomobject]@{ Url=$url; Title='Destino directo'; Owner='' }
        }
    }

    if (-not $site) { return }

    $siteUrl = [string]$site.Url
    if ([string]::IsNullOrWhiteSpace($siteUrl)) {
        Write-Status Error 'El sitio seleccionado no tiene una URL válida.'
        Pause-Tui
        return
    }

    Write-Status Info "Conectando a $siteUrl"
    try {
        $connection = Connect-M365SiteWithRecovery -Url $siteUrl
    }
    catch {
        Write-Status Error "No se pudo conectar: $($_.Exception.Message)"
        Pause-Tui
        return
    }

    try {
        $web = Get-PnPWeb -Connection $connection -Includes Title,Url -ErrorAction Stop
        $libraries = @(Resolve-Libraries -Connection $connection -TargetKind $kind)
    }
    catch {
        $message = $_.Exception.Message

        try {
            if (-not (Try-EnableTemporarySiteAdmin -SiteUrl $siteUrl -ErrorMessage $message)) {
                throw $message
            }

            $connection = Connect-M365SiteWithRecovery -Url $siteUrl -SkipAutoGrant
            $web = Get-PnPWeb -Connection $connection -Includes Title,Url -ErrorAction Stop
            $libraries = @(Resolve-Libraries -Connection $connection -TargetKind $kind)
        }
        catch {
            Write-Status Error "No se pudieron leer las bibliotecas: $($_.Exception.Message)"
            Pause-Tui
            return
        }
    }

    if ($libraries.Count -eq 0) {
        Write-Status Warn 'No se encontraron bibliotecas de documentos en este web.'
        Pause-Tui
        return
    }

    if ($kind -eq 'OneDrive') {
        $defaultLib = $libraries | Where-Object { $_.IsDefaultDocumentLibrary } | Select-Object -First 1
        if ($null -ne $defaultLib) {
            Write-Status Ok "Biblioteca principal detectada: $($defaultLib.Title) · $($defaultLib.RootFolderUrl)"
        }
        else {
            Write-Status Warn 'No se pudo identificar la biblioteca principal automáticamente; se muestran todas las bibliotecas detectadas.'
            Start-Sleep -Milliseconds 900
        }
    }

    try {
        $selected = @(Select-LibrariesInteractive -Libraries $libraries)
    }
    catch {
        Write-Status Error "No se pudieron seleccionar las bibliotecas: $($_.Exception.Message)"
        Pause-Tui
        return
    }

    if ($selected.Count -eq 0) {
        Write-Status Warn 'No se seleccionó ninguna biblioteca; el destino no fue guardado.'
        Pause-Tui
        return
    }

    $script:Target = [pscustomobject]@{
        Kind = $kind
        SiteUrl = $siteUrl.TrimEnd('/')
        SiteTitle = [string]$web.Title
        Owner = [string]$site.Owner
        Libraries = $selected
    }

    $script:PnPConnection = $connection

    Write-Status Ok "Destino seleccionado: $($script:Target.SiteTitle)"
    Start-Sleep -Milliseconds 700
}

function Reselect-TargetLibraries {
    Write-AppHeader 'Cambiar bibliotecas'

    if ($null -eq $script:Target -or [string]::IsNullOrWhiteSpace([string]$script:Target.SiteUrl)) {
        Write-Status Warn 'Primero selecciona un sitio SharePoint o OneDrive.'
        Pause-Tui
        return
    }

    Write-Field 'Tipo' $script:Target.Kind
    Write-Field 'Sitio' $script:Target.SiteTitle
    Write-Field 'URL' $script:Target.SiteUrl
    Write-Field 'Actuales' (($script:Target.Libraries.Title) -join ', ')
    Write-Host ''
    Write-Status Info 'Actualizando la lista de bibliotecas del sitio actual...'

    try {
        $connection = Connect-M365Site -Url $script:Target.SiteUrl
        if ($null -eq $connection) {
            throw 'No se obtuvo una conexión PnP válida.'
        }

        $libraries = @(
            Resolve-Libraries `
                -Connection $connection `
                -TargetKind $script:Target.Kind
        )
    }
    catch {
        Write-Status Error (
            'No se pudieron actualizar las bibliotecas: ' +
            "$($_.Exception.Message)"
        )
        Pause-Tui
        return
    }

    if ($libraries.Count -eq 0) {
        Write-Status Warn 'No se encontraron bibliotecas de documentos en el sitio actual.'
        Pause-Tui
        return
    }

    $currentIds = @(
        $script:Target.Libraries |
            ForEach-Object { [string]$_.Id }
    )

    $selected = @(
        Select-LibrariesInteractive `
            -Libraries $libraries `
            -PreselectedIds $currentIds
    )

    # Esc cancela sin destruir la selección anterior.
    if ($selected.Count -eq 0) {
        return
    }

    $script:Target.Libraries = $selected

    Write-AppHeader 'Bibliotecas actualizadas'
    Write-Field 'Sitio' $script:Target.SiteTitle
    Write-Field 'Seleccionadas' (($selected.Title) -join ', ')
    Write-Status Ok 'La selección de bibliotecas fue actualizada sin cambiar de sitio.'
    Start-Sleep -Milliseconds 900
}


function Test-TargetSelected {
    if (-not $script:Target -or -not $script:Target.SiteUrl -or @($script:Target.Libraries).Count -eq 0) {
        Write-Status Warn 'Primero selecciona un destino y al menos una biblioteca.'
        return $false
    }
    return $true
}

# -----------------------------------------------------------------------------
# Permissions / REST
# -----------------------------------------------------------------------------

function Convert-ToPnPRestUrl {
    param([string]$Url)

    if ($Url -match '^https?://') {
        $uri = [uri]$Url
        return $uri.PathAndQuery
    }
    return $Url
}

function Get-LibraryItemsViaRest {
    param(
        [Parameter(Mandatory)][string]$ListId,
        [Parameter(Mandatory)]$Connection
    )

    $pageSize = [Math]::Max(100, [int]$script:Settings.PageSize)
    $select = 'Id,FileRef,FileLeafRef,FileSystemObjectType,HasUniqueRoleAssignments'
    $nextUrl = "/_api/web/lists(guid'$ListId')/items?`$select=$select&`$orderby=Id&`$top=$pageSize"
    $items = [System.Collections.Generic.List[object]]::new()
    $page = 0

    while (-not [string]::IsNullOrWhiteSpace($nextUrl)) {
        $page++
        $requestUrl = Convert-ToPnPRestUrl $nextUrl
        $resp = Invoke-WithRetry -OperationName "REST página $page" -ScriptBlock {
            Invoke-PnPSPRestMethod -Url $requestUrl -Method Get -Connection $Connection -ErrorAction Stop
        }

        $valueProperty = $resp.PSObject.Properties['value']
        if ($null -ne $valueProperty -and $null -ne $valueProperty.Value) {
            foreach ($row in $valueProperty.Value) { [void]$items.Add($row) }
        }

        $nextUrl = $null
        foreach ($prop in @('@odata.nextLink','odata.nextLink','__next')) {
            $property = $resp.PSObject.Properties[$prop]
            if ($null -ne $property -and $property.Value) {
                $nextUrl = [string]$property.Value
                break
            }
        }

        Write-Progress -Activity 'Leyendo elementos' -Status "Página $page · $($items.Count) elementos"
    }

    Write-Progress -Activity 'Leyendo elementos' -Completed
    return $items
}

function Reset-ItemInheritanceViaRest {
    param([string]$ListId, [int]$ItemId, $Connection)
    $url = "/_api/web/lists(guid'$ListId')/items($ItemId)/resetroleinheritance"
    Invoke-PnPSPRestMethod -Url $url -Method Post -Connection $Connection -ErrorAction Stop | Out-Null
}

function Reset-LibraryInheritanceViaRest {
    param([string]$ListId, $Connection)
    $url = "/_api/web/lists(guid'$ListId')/resetroleinheritance"
    Invoke-PnPSPRestMethod -Url $url -Method Post -Connection $Connection -ErrorAction Stop | Out-Null
}

function New-ReportPath {
    param([string]$Prefix)
    Ensure-ReportFolder
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    return Join-Path $script:Settings.ReportFolder "${Prefix}_${stamp}.csv"
}

function Export-ObjectListToCsv {
    param(
        [Parameter(Mandatory)]$List,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        if ($null -eq $List) {
            Set-Content -LiteralPath $Path -Value '' -Encoding UTF8
            return
        }

        $rows = if ($List.PSObject.Methods['ToArray']) {
            $List.ToArray()
        }
        else {
            @($List | ForEach-Object { $_ })
        }

        if ($null -eq $rows -or $rows.Count -eq 0) {
            Set-Content -LiteralPath $Path -Value '' -Encoding UTF8
            return
        }

        $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Force
    }
    catch {
        throw "No se pudo exportar el CSV '$Path': $($_.Exception.Message)"
    }
}

function Run-CountUniqueScopes {
    param([switch]$SkipAutoGrant)

    Write-AppHeader 'Análisis de permisos únicos'
    if (-not (Test-TargetSelected)) { Pause-Tui; return }

    try {
        $connection = Connect-M365SiteWithRecovery `
            -Url $script:Target.SiteUrl `
            -SkipAutoGrant:$SkipAutoGrant
    }
    catch {
        Write-Status Error "Conexión fallida: $($_.Exception.Message)"
        Pause-Tui
        return
    }

    $report = [System.Collections.Generic.List[object]]::new()
    $summary = [System.Collections.Generic.List[object]]::new()

    foreach ($library in $script:Target.Libraries) {
        Write-Section $library.Title
        $flags = Get-LibraryFlagsText -Library $library
        if (-not [string]::IsNullOrWhiteSpace($flags)) { Write-Styled $flags.Trim() Muted }
        Write-Styled $library.RootFolderUrl Muted

        try {
            $list = Get-PnPList -Identity $library.Id -Connection $connection -Includes HasUniqueRoleAssignments,RootFolder -ErrorAction Stop
            $rootUnique = [bool](Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments -Connection $connection)
            $skipRoot = ($script:Target.Kind -eq 'OneDrive' -and $library.IsDefaultDocumentLibrary)
            $rootCounted = ($rootUnique -and -not $skipRoot)

            if ($rootCounted) {
                [void]$report.Add([pscustomobject]@{
                    Site      = $script:Target.SiteUrl
                    Library   = $library.Title
                    ScopeType = 'Library'
                    ItemId    = $null
                    Path      = $library.RootFolderUrl
                })
            }

            $items = @(Get-LibraryItemsViaRest -ListId $library.Id -Connection $connection | ForEach-Object { $_ })
            $uniqueItems = @($items | Where-Object { $_.HasUniqueRoleAssignments -eq $true })

            foreach ($item in $uniqueItems) {
                [void]$report.Add([pscustomobject]@{
                    Site      = $script:Target.SiteUrl
                    Library   = $library.Title
                    ScopeType = if ([int]$item.FileSystemObjectType -eq 1) { 'Folder' } else { 'File' }
                    ItemId    = [int]$item.Id
                    Path      = [string]$item.FileRef
                })
            }

            $total = $uniqueItems.Count
            if ($rootCounted) { $total++ }

            $scopeState = Get-ScopeStatus -Count $total

            [void]$summary.Add([pscustomobject]@{
                Library      = $library.Title
                Principal    = $library.IsDefaultDocumentLibrary
                Items        = $items.Count
                UniqueScopes = $total
                Recommended  = $script:RecommendedUniqueScopes
                Maximum      = $script:MaximumUniqueScopes
                Status       = $scopeState.Label
                RootUnique   = $rootUnique
                RootIncluded = $rootCounted
            })

            Write-Status Ok "$($items.Count) elementos · $total scopes únicos"

            if ($skipRoot -and $rootUnique) {
                Write-Styled 'La raíz de la biblioteca principal de OneDrive se omite del total accionable.' Muted
            }
        }
        catch {
            $message = $_.Exception.Message

            if (-not $SkipAutoGrant -and
                (Try-EnableTemporarySiteAdmin -SiteUrl $script:Target.SiteUrl -ErrorMessage $message)) {
                [void](Run-CountUniqueScopes -SkipAutoGrant)
                return
            }

            Write-Status Error "$($library.Title): $message"
            [void]$summary.Add([pscustomobject]@{
                Library      = $library.Title
                Principal    = $library.IsDefaultDocumentLibrary
                Items        = 0
                UniqueScopes = 0
                Recommended  = $script:RecommendedUniqueScopes
                Maximum      = $script:MaximumUniqueScopes
                Status       = 'ERROR DE ANÁLISIS'
                RootUnique   = $null
                RootIncluded = $false
            })
        }
    }

    $path = New-ReportPath -Prefix 'UniquePermissionScopes'
    try {
        Export-ObjectListToCsv -List $report -Path $path
    }
    catch {
        Write-Status Error $_.Exception.Message
        Pause-Tui
        return
    }

    Write-Section 'Resumen por biblioteca'

    if ($summary.Count -gt 0) {
        $summaryText = $summary.ToArray() |
            Format-Table Library,Items,UniqueScopes,Recommended,Maximum,Status -AutoSize |
            Out-String
        Write-Host $summaryText
    }

    $grandTotal = 0
    $librariesOverRecommended = 0
    $librariesAtMaximum = 0

    foreach ($row in $summary) {
        $grandTotal += [int]$row.UniqueScopes

        if ([int]$row.UniqueScopes -ge $script:MaximumUniqueScopes) {
            $librariesAtMaximum++
        }
        elseif ([int]$row.UniqueScopes -gt $script:RecommendedUniqueScopes) {
            $librariesOverRecommended++
        }
    }

    # Los límites de SharePoint se evalúan por lista/biblioteca, no sumando
    # bibliotecas independientes. El total siguiente es solo informativo.
    if ($summary.Count -eq 1) {
        Write-Section 'Capacidad de scopes'
        Write-ScopeMeter -Count ([int]$summary[0].UniqueScopes)
    }
    else {
        Write-Section 'Límites'
        Write-Field 'Recomendado' ("{0:N0} por biblioteca" -f $script:RecommendedUniqueScopes)
        Write-Field 'Máximo' ("{0:N0} por biblioteca" -f $script:MaximumUniqueScopes)
        Write-Field 'Total observado' $grandTotal
        Write-Styled 'Los límites se aplican individualmente a cada lista/biblioteca; el total observado no se compara contra 5,000 o 50,000.' Muted
    }

    Write-Section 'Estado'

    if ($librariesAtMaximum -gt 0) {
        Write-Status Error "$librariesAtMaximum biblioteca(s) alcanzaron o superaron el límite máximo de $($script:MaximumUniqueScopes) scopes."
    }

    if ($librariesOverRecommended -gt 0) {
        Write-Status Warn "$librariesOverRecommended biblioteca(s) superan el límite recomendado de $($script:RecommendedUniqueScopes) scopes."
    }

    if ($librariesAtMaximum -eq 0 -and $librariesOverRecommended -eq 0) {
        Write-Status Ok 'Todas las bibliotecas analizadas están dentro del límite recomendado.'
    }

    if ($null -ne $script:TemporarySiteAdmin) {
        Write-Status Info 'La tarea terminó. Puedes retirar ahora el acceso temporal.'
        [void](Remove-TemporarySiteAdmin -Ask)
        if ($null -ne $script:TemporarySiteAdmin) {
            Write-Status Warn 'El acceso temporal sigue activo.'
        }
    }

    Write-Field 'Reporte' $path
    Pause-Tui
}

function Invoke-ResetItemBatch {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Items,
        [Parameter(Mandatory)][string]$ListId,
        [Parameter(Mandatory)][string]$LibraryTitle,
        [Parameter(Mandatory)]$Connection,
        [Parameter(Mandatory)]$Log
    )

    $rows = @($Items)
    $count = $rows.Count

    # Una colección vacía es un estado válido: por ejemplo, una biblioteca puede
    # tener archivos con permisos únicos pero ninguna carpeta, o no tener ningún
    # elemento con permisos únicos. No debe tratarse como error.
    if ($count -eq 0) {
        return
    }

    for ($i = 0; $i -lt $count; $i++) {
        $item = $rows[$i]
        $type = if ([int]$item.FileSystemObjectType -eq 1) { 'Folder' } else { 'File' }
        $state = 'OK'
        $action = if ($script:Settings.DryRun) { 'Simulated' } else { 'Reset' }
        $errorText = ''

        try {
            if (-not $script:Settings.DryRun) {
                Invoke-WithRetry -OperationName "Reset $type ID $($item.Id)" -ScriptBlock {
                    Reset-ItemInheritanceViaRest -ListId $ListId -ItemId ([int]$item.Id) -Connection $Connection
                }
            }
        }
        catch {
            $state = 'ERROR'
            $errorText = $_.Exception.Message
        }

        $Log.Add([pscustomobject]@{
            Site = $script:Target.SiteUrl
            Library = $LibraryTitle
            Type = $type
            ItemId = [int]$item.Id
            Path = [string]$item.FileRef
            Action = $action
            Status = $state
            Error = $errorText
        })

        $pct = if ($count -gt 0) { (($i + 1) / $count) * 100 } else { 100 }
        Write-Progress -Activity "Procesando $LibraryTitle" -Status "$($i + 1)/$count" -PercentComplete $pct
    }
    Write-Progress -Activity "Procesando $LibraryTitle" -Completed
}

function Run-ResetInheritance {
    param([switch]$SkipAutoGrant)

    Write-AppHeader 'Restablecer herencia'
    if (-not (Test-TargetSelected)) { Pause-Tui; return }

    Write-Field 'Sitio' $script:Target.SiteUrl
    Write-Field 'Bibliotecas' (($script:Target.Libraries.Title) -join ', ')
    Write-Field 'Modo' $(if ($script:Settings.DryRun) { 'SIMULACIÓN' } else { 'REAL' })

    if ($script:Settings.DryRun) {
        Write-Status Warn 'Dry-run activo: se detectarán los cambios, pero no se modificará ningún permiso.'
    }
    else {
        Write-Status Error 'Modo REAL: se eliminarán asignaciones de permisos únicas y los objetos volverán a heredar.'
        $confirmation = Read-Host 'Escribe RESET para continuar'
        if ($confirmation -cne 'RESET') {
            Write-Status Warn 'Operación cancelada.'
            Pause-Tui
            return
        }
    }

    try {
        $connection = Connect-M365SiteWithRecovery `
            -Url $script:Target.SiteUrl `
            -SkipAutoGrant:$SkipAutoGrant
    }
    catch {
        Write-Status Error "Conexión fallida: $($_.Exception.Message)"
        Pause-Tui
        return
    }

    $log = [System.Collections.Generic.List[object]]::new()

    foreach ($library in $script:Target.Libraries) {
        Write-Section $library.Title
        try {
            $items = @(Get-LibraryItemsViaRest -ListId $library.Id -Connection $connection | ForEach-Object { $_ })
            $unique = @($items | Where-Object { $_.HasUniqueRoleAssignments -eq $true })
            $folders = @($unique | Where-Object { [int]$_.FileSystemObjectType -eq 1 } | Sort-Object { ([string]$_.FileRef -split '/').Count })
            $files = @($unique | Where-Object { [int]$_.FileSystemObjectType -ne 1 })

            Write-Status Info "$($unique.Count) elementos con permisos únicos."

            if ($folders.Count -gt 0) {
                Invoke-ResetItemBatch -Items $folders -ListId $library.Id -LibraryTitle $library.Title -Connection $connection -Log $log
            }

            if ($files.Count -gt 0) {
                Invoke-ResetItemBatch -Items $files -ListId $library.Id -LibraryTitle $library.Title -Connection $connection -Log $log
            }

            $skipRoot = ($script:Target.Kind -eq 'OneDrive' -and $library.IsDefaultDocumentLibrary)
            if ($skipRoot) {
                Write-Styled 'Raíz de la biblioteca principal de OneDrive omitida.' Muted
            }
            else {
                $list = Get-PnPList -Identity $library.Id -Connection $connection -Includes HasUniqueRoleAssignments -ErrorAction Stop
                $rootUnique = [bool](Get-PnPProperty -ClientObject $list -Property HasUniqueRoleAssignments -Connection $connection)
                if ($rootUnique) {
                    $state = 'OK'
                    $err = ''
                    try {
                        if (-not $script:Settings.DryRun) {
                            Invoke-WithRetry -OperationName "Reset raíz $($library.Title)" -ScriptBlock {
                                Reset-LibraryInheritanceViaRest -ListId $library.Id -Connection $connection
                            }
                        }
                    }
                    catch { $state = 'ERROR'; $err = $_.Exception.Message }

                    [void]$log.Add([pscustomobject]@{
                        Site    = $script:Target.SiteUrl
                        Library = $library.Title
                        Type    = 'Library'
                        ItemId  = $null
                        Path    = $library.RootFolderUrl
                        Action  = if ($script:Settings.DryRun) { 'Simulated' } else { 'Reset' }
                        Status  = $state
                        Error   = $err
                    })
                }
            }
        }
        catch {
            $message = $_.Exception.Message

            if (-not $SkipAutoGrant -and
                (Try-EnableTemporarySiteAdmin -SiteUrl $script:Target.SiteUrl -ErrorMessage $message)) {
                [void](Run-ResetInheritance -SkipAutoGrant)
                return
            }

            Write-Status Error "$($library.Title): $message"
            [void]$log.Add([pscustomobject]@{
                Site    = $script:Target.SiteUrl
                Library = $library.Title
                Type    = 'Library'
                ItemId  = $null
                Path    = $library.RootFolderUrl
                Action  = if ($script:Settings.DryRun) { 'Simulated' } else { 'Reset' }
                Status  = 'ERROR'
                Error   = $_.Exception.Message
            })
        }
    }

    $path = New-ReportPath -Prefix 'ResetInheritance'
    try {
        Export-ObjectListToCsv -List $log -Path $path
    }
    catch {
        Write-Status Error $_.Exception.Message
        Pause-Tui
        return
    }

    $ok = 0
    $errors = 0
    foreach ($row in $log) {
        if ($row.Status -eq 'OK') { $ok++ }
        elseif ($row.Status -eq 'ERROR') { $errors++ }
    }

    Write-Section 'Resultado'
    Write-Field 'Procesados' $log.Count
    Write-Field 'Correctos' $ok
    Write-Field 'Errores' $errors
    Write-Field 'Reporte' $path

    if ($errors -gt 0) { Write-Status Warn 'La operación terminó con errores; revisa el CSV.' }
    elseif ($script:Settings.DryRun) { Write-Status Ok 'Simulación completada sin modificar permisos.' }
    else { Write-Status Ok 'Restablecimiento completado.' }

    if ($null -ne $script:TemporarySiteAdmin) {
        Write-Status Info 'La tarea terminó. Puedes retirar ahora el acceso temporal.'
        [void](Remove-TemporarySiteAdmin -Ask)
        if ($null -ne $script:TemporarySiteAdmin) {
            Write-Status Warn 'El acceso temporal sigue activo.'
        }
    }

    Pause-Tui
}

# -----------------------------------------------------------------------------
# Site Collection Admin
# -----------------------------------------------------------------------------

function Get-AdminUpn {
    if (-not [string]::IsNullOrWhiteSpace($script:Settings.AdminUpn)) { return $script:Settings.AdminUpn }

    $value = Read-TextValue -Prompt 'UPN del administrador' `
        -Validator { param($v) $v -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' } `
        -ValidationMessage 'UPN inválido.'

    if (Read-YesNo '¿Guardar este UPN para futuras ejecuciones?' $true) {
        $script:Settings.AdminUpn = $value
        Save-Settings
    }
    return $value
}

function Grant-SiteAdminAccess {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [switch]$TrackTemporary
    )

    try {
        $adminUpn = Get-AdminUpn
        $adminUrl = Get-AdminUrl
        $conn = Connect-M365Site -Url $adminUrl
        Set-PnPTenantSite -Identity $SiteUrl -Owners $adminUpn -Connection $conn -ErrorAction Stop

        if ($TrackTemporary) {
            # Track the grant immediately so it can still be removed if the
            # following site reconnection or propagation check fails.
            $script:TemporarySiteAdmin = [pscustomobject]@{
                Upn     = $adminUpn
                SiteUrl = $SiteUrl
            }
        }

        Write-Status Ok "$adminUpn agregado como Site Collection Admin."
        return $true
    }
    catch {
        Write-Status Error "No se pudo otorgar acceso: $($_.Exception.Message)"
        return $false
    }
}

function Remove-TemporarySiteAdmin {
    param([switch]$Ask)

    if ($null -eq $script:TemporarySiteAdmin) {
        return $true
    }

    $temporaryAdmin = $script:TemporarySiteAdmin

    if ($Ask -and -not (Read-YesNo -Prompt '¿Retirar el acceso temporal ahora?' -Default $true)) {
        Write-Status Warn 'No se retiró el acceso temporal.'
        return $false
    }

    try {
        $connection = Connect-M365Site -Url $temporaryAdmin.SiteUrl

        Remove-PnPSiteCollectionAdmin `
            -Owners $temporaryAdmin.Upn `
            -Connection $connection `
            -ErrorAction Stop

        $script:TemporarySiteAdmin = $null
        Write-Status Ok 'Acceso temporal retirado correctamente.'
        return $true
    }
    catch {
        Write-Status Error 'No se pudo retirar el acceso temporal.'
        Write-Styled $_.Exception.Message Danger
        return $false
    }
}

function Revoke-SiteAdminAccess {
    param([Parameter(Mandatory)][string]$SiteUrl)

    try {
        $adminUpn = Get-AdminUpn
        $conn = Connect-M365Site -Url $SiteUrl
        Remove-PnPSiteCollectionAdmin -Owners $adminUpn -Connection $conn -ErrorAction Stop
        Write-Status Ok "$adminUpn removido de Site Collection Admin."
        return $true
    }
    catch {
        Write-Status Error "No se pudo remover acceso: $($_.Exception.Message)"
        return $false
    }
}

function Show-AdminMenu {
    Write-AppHeader 'Site Collection Admin'
    if (-not (Test-TargetSelected)) { Pause-Tui; return }

    Write-Field 'Sitio' $script:Target.SiteUrl
    Write-Host ''
    $choice = Read-MenuChoice -AllowBack -Items @(
        @{ Key='1'; Label='Otorgar Site Collection Admin' },
        @{ Key='2'; Label='Remover Site Collection Admin' }
    )

    switch ($choice) {
        '1' { [void](Grant-SiteAdminAccess -SiteUrl $script:Target.SiteUrl); Pause-Tui }
        '2' { [void](Revoke-SiteAdminAccess -SiteUrl $script:Target.SiteUrl); Pause-Tui }
    }
}

# -----------------------------------------------------------------------------
# Diagnostics / setup
# -----------------------------------------------------------------------------

function Run-SetupDiagnostics {
    Write-AppHeader 'Diagnóstico y autenticación'

    Write-Field 'PowerShell' $PSVersionTable.PSVersion

    $pnp = Get-InstalledPnPModule

    Write-Field 'PnP.PowerShell' $(if ($pnp) { $pnp.Version } else { 'No instalado' })
    Write-Field 'Tenant' $script:Settings.Tenant
    Write-Field 'Client ID' $script:Settings.ClientId
    Write-Field 'App' $script:Settings.AppRegistrationName
    Write-Field 'Persist login' $(if ($script:Settings.PersistLogin) { 'Sí' } else { 'No' })
    Write-Host ''

    $choice = Read-MenuChoice -AllowBack -Items @(
        @{
            Key   = '1'
            Label = 'Instalar / actualizar PnP.PowerShell'
        },
        @{
            Key         = '2'
            Label       = 'Gestionar aplicación Entra / Client ID'
            Description = 'Validar, usar una existente o registrar una nueva.'
        },
        @{
            Key         = '3'
            Label       = 'Probar conexión al centro de administración'
            Description = 'Usa la aplicación actualmente configurada.'
        },
        @{
            Key   = '4'
            Label = 'Limpiar login persistido de PnP'
        }
    )

    switch ($choice) {
        '0' {
            return
        }

        '1' {
            [void](Ensure-PnPModule -OfferInstall)
            Pause-Tui
        }

        '2' {
            Show-AppRegistrationMenu
        }

        '3' {
            try {
                if (-not (Ensure-ClientId)) {
                    Write-Status Warn 'No hay una aplicación configurada para realizar la prueba.'
                    Pause-Tui
                    return
                }

                $conn = Connect-M365Site -Url (Get-AdminUrl)
                $web = Get-PnPWeb -Connection $conn -Includes Title,Url -ErrorAction Stop

                Write-Status Ok "Conexión correcta: $($web.Title)"
                Write-Field 'URL' $web.Url
                Write-Field 'Client ID' $script:Settings.ClientId
            }
            catch {
                Write-Status Error $_.Exception.Message
            }

            Pause-Tui
        }

        '4' {
            try {
                $url = if ($script:Target) {
                    $script:Target.SiteUrl
                }
                else {
                    Get-AdminUrl
                }

                if ([string]::IsNullOrWhiteSpace($url)) {
                    throw 'No se pudo determinar una URL para inicializar la sesión PnP.'
                }

                if (-not (Ensure-ClientId)) {
                    Write-Status Warn 'No hay una aplicación configurada.'
                    Pause-Tui
                    return
                }

                [void](Connect-M365Site -Url $url)

                Disconnect-PnPOnline `
                    -ClearPersistedLogin `
                    -ErrorAction Stop

                Write-Status Ok 'Login persistido de PnP eliminado.'
            }
            catch {
                Write-Status Error $_.Exception.Message
            }

            Pause-Tui
        }
    }
}

# -----------------------------------------------------------------------------
# Main screen
# -----------------------------------------------------------------------------

function Show-CurrentTarget {
    Write-Section 'Contexto'
    Write-Field 'Tenant' $script:Settings.Tenant

    if ($script:Target) {
        $selectedLibraries = @($script:Target.Libraries)
        $libraryNames = ($selectedLibraries.Title -join ', ')

        Write-Field 'Tipo' $script:Target.Kind
        Write-Field 'Sitio' $script:Target.SiteTitle
        Write-Field 'URL' $script:Target.SiteUrl
        Write-Field 'Libraries' ("{0} seleccionada(s)" -f $selectedLibraries.Count)
        Write-Styled ("               " + $libraryNames) Muted
    }
    else {
        Write-Field 'Destino' 'No seleccionado'
    }

    if ($script:Settings.DryRun) {
        Write-Styled 'Modo           ● SIMULACIÓN' Success
    }
    else {
        Write-Styled 'Modo           ● REAL' Danger
    }
}

function Show-MainMenu {
    while ($true) {
        Write-AppHeader
        Show-CurrentTarget
        Write-Host ''

        $changeLibrariesDescription = if ($script:Target) {
            'Cambiar solo las bibliotecas del sitio actual; conserva el sitio seleccionado.'
        }
        else {
            'Disponible después de seleccionar un sitio.'
        }

        $menuItems = @(
            @{ Key='1'; Label='Seleccionar / cambiar sitio'; Description='SharePoint u OneDrive; al elegirlo se seleccionan sus bibliotecas.' },
            @{ Key='2'; Label='Cambiar bibliotecas del sitio actual'; Description=$changeLibrariesDescription },
            @{ Key='3'; Label='Analizar Unique Permission Scopes'; Description='Solo lectura · muestra recomendado 5,000 y máximo 50,000 por biblioteca · genera CSV.' },
            @{ Key='4'; Label='Restablecer herencia'; Description='Respeta Dry-run y procesa las bibliotecas seleccionadas.' },
            @{ Key='5'; Label=$(if ($script:Settings.DryRun) { 'Cambiar a modo REAL' } else { 'Cambiar a modo SIMULACIÓN' }); Description='El modo simulación es el valor seguro.' },
            @{ Key='6'; Label='Gestionar Site Collection Admin' },
            @{ Key='7'; Label='Configuración' },
            @{ Key='8'; Label='Diagnóstico / autenticación' }
        )

        if ($null -ne $script:TemporarySiteAdmin) {
            $menuItems += @{ Key='9'; Label='Retirar acceso temporal de Site Collection Admin'; Description='Quita la elevación temporal usada para esta tarea.' }
        }

        $menuItems += @{ Key='0'; Label='Salir' }
        $choice = Read-MenuChoice -Items $menuItems

        switch ($choice) {
            '1' { Select-Target }
            '2' { Reselect-TargetLibraries }
            '3' { Run-CountUniqueScopes }
            '4' { Run-ResetInheritance }
            '5' {
                $script:Settings.DryRun = -not $script:Settings.DryRun
                Save-Settings
            }
            '6' { Show-AdminMenu }
            '7' { Show-SettingsMenu }
            '8' { Run-SetupDiagnostics }
            '9' {
                [void](Remove-TemporarySiteAdmin -Ask)
                Pause-Tui
            }
            '0' {
                if ($null -ne $script:TemporarySiteAdmin) {
                    if (-not (Remove-TemporarySiteAdmin -Ask)) {
                        Write-Status Warn 'El acceso temporal sigue activo.'
                        Pause-Tui
                        continue
                    }
                }

                return
            }
        }
    }
}

# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------

Initialize-AppLanguage
Initialize-Terminal

if (-not (Test-Path -LiteralPath $script:ConfigRoot)) {
    [void](New-Item -ItemType Directory -Path $script:ConfigRoot -Force)
}

$script:Settings = Load-Settings
Ensure-ReportFolder
Save-Settings

if (-not (Test-PowerShellVersion)) {
    Write-Host ''
    Write-Styled 'Ejecuta este toolkit desde PowerShell 7.4 o superior (pwsh).' Warning
    exit 1
}

Show-MainMenu
