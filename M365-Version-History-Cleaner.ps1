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

$script:MinimumPnPVersion = [version]'3.2.0'
$script:AdminConnection = $null
$script:PnPConnection = $null
$script:AuthRecoveryUsed = $false
$script:Target = $null
$script:LastAnalysis = $null
$script:LastAnalysisReports = $null
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
    @{ From = 'REAL'; To = 'LIVE' }; @{ From = 'Presione ENTER para cerrar'; To = 'Press ENTER to close' }
    @{ From = 'Enter para continuar'; To = 'Press Enter to continue' }; @{ From = 'No disponible'; To = 'Not available' }
    @{ From = 'Analizando'; To = 'Analyzing' }; @{ From = 'Revalidando y eliminando versiones'; To = 'Revalidating and removing versions' }
    @{ From = 'Versiones históricas'; To = 'Historical versions' }; @{ From = 'Archivos afectados'; To = 'Affected files' }
    @{ From = 'Errores'; To = 'Errors' }; @{ From = 'Duración'; To = 'Duration' }; @{ From = 'Eliminadas'; To = 'Removed' }
    @{ From = 'Omitidas por revalidación'; To = 'Skipped during revalidation' }; @{ From = 'Eventos de throttling'; To = 'Throttling events' }
    @{ From = 'Escriba ELIMINAR para continuar.'; To = 'Type DELETE to continue.' }; @{ From = 'Confirmación'; To = 'Confirmation' }
    @{ From = 'Sesión finalizada.'; To = 'Session finished.' }
    @{ From = 'Limpiar contexto de trabajo'; To = 'Clear working context' }; @{ From = 'Gestión del contexto'; To = 'Working context' }
    @{ From = 'Gestionar contexto de trabajo'; To = 'Manage working context' }; @{ From = 'Cambiar tenant o aplicación conectada'; To = 'Change tenant or connected application' }
    @{ From = 'Abrir carpeta de reportes'; To = 'Open reports folder' }; @{ From = 'Contexto'; To = 'Context' }
    @{ From = 'Sesión en memoria'; To = 'In-memory session' }; @{ From = 'ACTIVA'; To = 'ACTIVE' }; @{ From = 'INACTIVA'; To = 'INACTIVE' }
    @{ From = 'VACÍA'; To = 'EMPTY' }; @{ From = 'Tipo'; To = 'Type' }; @{ From = 'URL'; To = 'URL' }
    @{ From = 'Aplicación configurada'; To = 'Configured application' }; @{ From = 'Aplicación'; To = 'Application' }
    @{ From = 'Estado'; To = 'Status' }; @{ From = 'Operativa'; To = 'Ready' }; @{ From = 'Requiere configuración'; To = 'Configuration required' }
    @{ From = 'Limpiar destino y conexión'; To = 'Clear target and connection' }; @{ From = 'Cambiar tenant o aplicación'; To = 'Change tenant or application' }
    @{ From = 'No hay contexto activo'; To = 'No active context' }; @{ From = 'Contexto activo'; To = 'Active context' }
    @{ From = 'No hay un sitio seleccionado.'; To = 'No site is selected.' }; @{ From = 'El contexto de trabajo fue limpiado.'; To = 'Working context cleared.' }
    @{ From = 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'; To = 'The in-memory session is reused while the tool remains open.' }
    @{ From = 'Sesión persistente: si la activas, PnP puede reutilizar el inicio de sesión cuando abras el script nuevamente.'; To = 'Persisted session: when enabled, PnP can reuse the sign-in when you open the script again.' }
    @{ From = 'Si la desactivas, solo se usa la sesión actual y podrás iniciar sesión de nuevo en la siguiente ejecución.'; To = 'When disabled, only the current session is used and you can sign in again on the next run.' }
    @{ From = '¿Guardar la sesión para futuras ejecuciones?'; To = 'Save the session for future runs?' }
    @{ From = 'La sesión autenticada se conserva en memoria mientras el script está abierto.'; To = 'The authenticated session stays in memory while the script is open.' }
    @{ From = 'La persistencia controla si PnP reutiliza el inicio de sesión al abrir el script nuevamente.'; To = 'Persistence controls whether PnP reuses the sign-in when the script is opened again.' }
    @{ From = 'No se puede cambiar tenant o aplicación mientras existe una operación activa.'; To = 'Tenant or application cannot be changed while an operation is active.' }
    @{ From = 'Tenant actualizado.'; To = 'Tenant updated.' }; @{ From = 'Persistencia de login'; To = 'Login persistence' }
    @{ From = 'Alternar persistencia de login'; To = 'Toggle login persistence' }; @{ From = 'Restablecer configuración local'; To = 'Reset local settings' }
    @{ From = 'Borra valores guardados; conserva idioma y reportes.'; To = 'Deletes saved values; preserves language and reports.' }
    @{ From = 'El idioma se conserva.'; To = 'The language is preserved.' }; @{ From = 'El módulo instalado no cumple el mínimo requerido.'; To = 'The installed module does not meet the required minimum.' }
    @{ From = 'Instalar / actualizar PnP.PowerShell'; To = 'Install / update PnP.PowerShell' }; @{ From = 'Verifica el requisito local y ofrece instalar o actualizar el módulo.'; To = 'Checks the local requirement and offers to install or update the module.' }
    @{ From = 'Comprueba tenant, Client ID, autenticación y acceso al centro de administración.'; To = 'Checks tenant, Client ID, authentication, and admin center access.' }
    @{ From = 'Valida la conexión y los permisos del destino seleccionado.'; To = 'Validates the connection and permissions for the selected target.' }
    @{ From = 'Nueva limpieza'; To = 'New cleanup' }; @{ From = 'Limpieza segura del historial de versiones'; To = 'Safe version-history cleanup' }
    @{ From = 'Configuración de la aplicación, idioma, sesión, throttling y reportes.'; To = 'Application, language, session, throttling, and report settings.' }
    @{ From = 'Selecciona una acción para continuar.'; To = 'Select an action to continue.' }; @{ From = 'Regresa al menú anterior.'; To = 'Return to the previous menu.' }
    @{ From = 'Cierra la herramienta.'; To = 'Closes the tool.' }; @{ From = 'No hay una aplicación Entra válida configurada.'; To = 'No valid Entra application is configured.' }
    @{ From = 'Configuración restablecida.'; To = 'Settings reset.' }; @{ From = 'No se pudo abrir la carpeta de reportes.'; To = 'The reports folder could not be opened.' }
    @{ From = 'Configuración guardada.'; To = 'Settings saved.' }; @{ From = 'Prueba una conexión real contra el destino actual.'; To = 'Tests a real connection against the current target.' }
    @{ From = 'Prueba autenticación real contra este tenant'; To = 'Test real authentication against this tenant' }
    @{ From = 'Introduce un Client ID ya registrado'; To = 'Enter an already registered Client ID' }
    @{ From = 'Crea una app PnP nueva'; To = 'Create a new PnP app' }
    @{ From = 'No elimina la app de Entra'; To = 'Does not delete the Entra app' }
    @{ From = 'Limpia el token persistido de PnP'; To = 'Clears the persisted PnP token' }
    @{ From = 'Client ID ya registrado'; To = 'Already registered Client ID' }
    @{ From = 'Crear mediante PnP.PowerShell'; To = 'Create using PnP.PowerShell' }
    @{ From = 'Busca por nombre, URL o propietario'; To = 'Search by name, URL, or owner' }
    @{ From = 'Busca por propietario o URL real'; To = 'Search by owner or real URL' }
    @{ From = 'Buscar en el tenant o usar URL directa'; To = 'Search in the tenant or use a direct URL' }
    @{ From = 'Buscar OneDrive reales del tenant'; To = 'Search for real OneDrive sites in the tenant' }
    @{ From = 'Detecta automáticamente SPO u OneDrive por la URL'; To = 'Automatically detects SPO or OneDrive from the URL' }
    @{ From = 'Cambia la aplicación'; To = 'Changes the application' }
    @{ From = 'Versionado activo'; To = 'Versioning enabled' }
    @{ From = 'Versionado desactivado'; To = 'Versioning disabled' }
    @{ From = 'OneDrive / MySite (700)'; To = 'OneDrive / MySite (700)' }
    @{ From = 'Plantilla '; To = 'Template ' }
    @{ From = 'Items: '; To = 'Items: ' }
    @{ From = 'Número de versiones a conservar'; To = 'Number of versions to keep' }
    @{ From = 'Pausa entre solicitudes (ms)'; To = 'Delay between requests (ms)' }
    @{ From = 'Espera base de retry (segundos)'; To = 'Base retry wait (seconds)' }
    @{ From = 'No se encontraron sitios accesibles.'; To = 'No accessible sites were found.' }
    @{ From = 'Sin coincidencias.'; To = 'No matches.' }
    @{ From = 'Hay '; To = 'There are ' }
    @{ From = ' resultados; refina el filtro.'; To = ' results; refine the filter.' }
    @{ From = ' resultados. Se muestran los primeros '; To = ' results. Showing the first ' }
    @{ From = 'Seleccionadas: '; To = 'Selected: ' }
    @{ From = 'Sitios disponibles: '; To = 'Available sites: ' }
    @{ From = 'Filtra por nombre, URL o propietario. Deja vacío para mostrar todos si son 50 o menos.'; To = 'Filter by name, URL, or owner. Leave empty to show all when there are 50 or fewer.' }
    @{ From = 'Reintento '; To = 'Retry ' }
    @{ From = ' segundo(s).'; To = ' second(s).' }
    @{ From = 'segundo(s)'; To = 'second(s)' }
    @{ From = 'Activada'; To = 'Enabled' }
    @{ From = 'Desactivada'; To = 'Disabled' }
    @{ From = 'DISPONIBLE'; To = 'AVAILABLE' }
    @{ From = 'NO DISPONIBLE'; To = 'NOT AVAILABLE' }
    @{ From = 'VACÍA'; To = 'EMPTY' }
    @{ From = 'ACTIVO'; To = 'ACTIVE' }
    @{ From = 'VACÍO'; To = 'EMPTY' }
    @{ From = 'No hay un sitio seleccionado.'; To = 'No site is selected.' }
    @{ From = 'Elija 1 o 2.'; To = 'Please choose 1 or 2.' }
    @{ From = 'Responde S o N.'; To = 'Answer Y or N.' }
    @{ From = 'Actualiza el tenant conectado y libera la sesión anterior.'; To = 'Updates the connected tenant and releases the previous session.' }
    @{ From = 'Valida, cambia o registra la aplicación usada por PnP.PowerShell.'; To = 'Validates, changes, or registers the application used by PnP.PowerShell.' }
    @{ From = 'Controla si PnP puede reutilizar el login al abrir el script nuevamente.'; To = 'Controls whether PnP can reuse the sign-in when the script is opened again.' }
    @{ From = 'Tenant, idioma, persistencia de login, throttling y reportes.'; To = 'Tenant, language, login persistence, throttling, and reports.' }
    @{ From = 'Analiza y limpia el historial de versiones de SharePoint o OneDrive.'; To = 'Analyzes and cleans version history from SharePoint or OneDrive.' }
    @{ From = 'ACTIVO · limpia el destino, cambia tenant/aplicación o abre reportes.'; To = 'ACTIVE · clears the target, changes tenant/application, or opens reports.' }
    @{ From = 'VACÍO · limpia selecciones o cambia tenant/aplicación.'; To = 'EMPTY · clears selections or changes tenant/application.' }
    @{ From = 'Limpia destino, conexión y resultados; conserva tenant, aplicación y reportes.'; To = 'Clears the target, connection, and results; preserves tenant, application, and reports.' }
    @{ From = 'Abre configuración y libera la sesión anterior cuando cambien estos valores.'; To = 'Opens settings and releases the previous session when these values change.' }
    @{ From = 'Muestra los resultados agrupados por biblioteca.'; To = 'Shows results grouped by library.' }
    @{ From = 'Abre la carpeta con los CSV generados.'; To = 'Opens the folder containing the generated CSV files.' }
    @{ From = 'Revalida cada versión antes de eliminar'; To = 'Revalidates each version before removal' }
    @{ From = 'Hasta 50 archivos ordenados por impacto'; To = 'Up to 50 files sorted by impact' }
    @{ From = 'No cierre esta ventana.'; To = 'Do not close this window.' }
    @{ From = 'Elimina valores guardados; conserva idioma y reportes.'; To = 'Deletes saved values; preserves language and reports.' }
    @{ From = '¿Desea actualizarlo?'; To = 'Update it?' }
    @{ From = '¿Instalar/actualizar ahora?'; To = 'Install/update now?' }
    @{ From = '¿Quitar el Client ID guardado?'; To = 'Remove the saved Client ID?' }
    @{ From = '¿Registrar esta aplicación?'; To = 'Register this application?' }
    @{ From = '¿Validar la nueva aplicación ahora?'; To = 'Validate the new application now?' }
    @{ From = 'El tenant no tiene formato nombre.onmicrosoft.com.'; To = 'The tenant does not match the name.onmicrosoft.com format.' }
    @{ From = 'El Client ID no es un GUID válido.'; To = 'The Client ID is not a valid GUID.' }
    @{ From = 'Client ID no válido.'; To = 'Invalid Client ID.' }
    @{ From = 'No se pudo registrar la aplicación: '; To = 'The application could not be registered: ' }
    @{ From = 'Aplicación registrada: '; To = 'Application registered: ' }
    @{ From = 'Conexión establecida.'; To = 'Connection established.' }
    @{ From = 'No se encontraron bibliotecas documentales visibles.'; To = 'No visible document libraries were found.' }
    @{ From = 'No se ha modificado ningún archivo.'; To = 'No files were modified.' }
    @{ From = 'Auditoría:'; To = 'Audit:' }
    @{ From = 'Registro de throttling:'; To = 'Throttling log:' }
    @{ From = 'Proceso interrumpido'; To = 'Process interrupted' }
    @{ From = 'Limpieza completada'; To = 'Cleanup completed' }
    @{ From = 'Conexión y permisos del sitio correctos.'; To = 'Site connection and permissions are correct.' }
    @{ From = 'No se puede abrir la carpeta de reportes.'; To = 'The reports folder could not be opened.' }
    @{ From = 'Se muestran hasta 50 archivos ordenados por cantidad de versiones elegibles.'; To = 'Up to 50 files are shown, ordered by eligible version count.' }
    @{ From = 'Resultados agrupados por biblioteca.'; To = 'Results grouped by library.' }
    @{ From = 'La sesión autenticada no pudo reutilizarse; se solicitará autenticación nuevamente una sola vez.'; To = 'The authenticated session could not be reused; sign-in will be requested again only once.' }
    @{ From = 'La configuración guardada no se modificó.'; To = 'The saved configuration was not changed.' }
    @{ From = 'Permiso delegado solicitado:'; To = 'Delegated permission requested:' }
    @{ From = 'Consultando sitios del tenant...'; To = 'Querying tenant sites...' }
    @{ From = ' archivo(s) encontrado(s).'; To = ' file(s) found.' }
    @{ From = 'Throttling detectado durante eliminación.'; To = 'Throttling detected during removal.' }
    @{ From = 'Esperando '; To = 'Waiting ' }
    @{ From = ' segundo(s) antes del siguiente intento.'; To = ' second(s) before the next attempt.' }
    @{ From = 'Valores conservadores recomendados para trabajo secuencial.'; To = 'Conservative values recommended for sequential work.' }
    @{ From = 'Reporte detallado:'; To = 'Detailed report:' }
    @{ From = 'Limpia el destino, la conexión de sitio y los resultados del análisis. Conserva tenant, aplicación, preferencias y reportes.'; To = 'Clears the target, site connection, and analysis results. Preserves tenant, application, preferences, and reports.' }
    @{ From = 'La sesión de autenticación base se conserva; cambiar tenant o aplicación la renovará.'; To = 'The base authentication session is preserved; changing the tenant or application renews it.' }
    @{ From = 'Se eliminarán las preferencias locales.'; To = 'Local preferences will be deleted.' }
    @{ From = 'Ejecuta la herramienta desde PowerShell 7.4 o superior usando pwsh.'; To = 'Run the tool from PowerShell 7.4 or later using pwsh.' }
    @{ From = '¿Continuar?'; To = 'Continue?' }
    @{ From = '¿Limpiar el contexto de trabajo?'; To = 'Clear the working context?' }
    @{ From = 'URL del sitio'; To = 'Site URL' }; @{ From = 'URL de OneDrive'; To = 'OneDrive URL' }
    @{ From = 'URL del sitio / OneDrive'; To = 'Site / OneDrive URL' }; @{ From = 'Filtro'; To = 'Filter' }
    @{ From = 'Client ID creado'; To = 'Created Client ID' }; @{ From = 'Nombre de la aplicación ['; To = 'Application name [' }
    @{ From = 'Se muestran los 50 archivos con más versiones a eliminar.'; To = 'The 50 files with the most versions to remove are shown.' }
    @{ From = 'Las versiones históricas más antiguas serán elegibles para limpieza.'; To = 'The oldest historical versions will be eligible for cleanup.' }
    @{ From = 'Versión actual: conservar'; To = 'Current version: keep' }
    @{ From = 'versiones históricas: conservar'; To = 'historical versions: keep' }
    @{ From = 'Actual + '; To = 'Current + ' }
    @{ From = ' histórica(s)'; To = ' historical version(s)' }
    @{ From = ' versiones totales'; To = ' total versions' }
)
function Get-LocalizedText {
    param([AllowNull()][object]$Text)
    if ($null -eq $Text) { return '' }; $result = [string]$Text
    if ($script:Language -eq 'en') {
        $phrases = @(
            @{ From = 'Selecciona una opción'; To = 'Select an option' }; @{ From = 'Nueva limpieza'; To = 'New cleanup' }
            @{ From = 'Idioma'; To = 'Language' }; @{ From = 'Cambiar idioma'; To = 'Change language' }
            @{ From = 'No seleccionado'; To = 'Not selected' }; @{ From = 'No configurado'; To = 'Not configured' }; @{ From = 'No instalado'; To = 'Not installed' }
            @{ From = 'Destino'; To = 'Target' }; @{ From = 'CONFIGURADA'; To = 'CONFIGURED' }; @{ From = 'NO CONFIGURADA'; To = 'NOT CONFIGURED' }
            @{ From = 'Activado'; To = 'Enabled' }; @{ From = 'Desactivado'; To = 'Disabled' }; @{ From = 'Sí'; To = 'Yes' }; @{ From = 'SÍ'; To = 'YES' }
            @{ From = 'Limpieza segura del historial de versiones'; To = 'Safe version-history cleanup' }
            @{ From = 'SharePoint Online o OneDrive for Business'; To = 'SharePoint Online or OneDrive for Business' }
            @{ From = 'Buscar sitio en el tenant'; To = 'Search for a site in the tenant' }; @{ From = 'Introducir URL directamente'; To = 'Enter URL directly' }
            @{ From = 'Usar último sitio'; To = 'Use last site' }; @{ From = 'Usar último OneDrive'; To = 'Use last OneDrive' }
            @{ From = 'Para un sitio ya conocido'; To = 'For a known site' }; @{ From = 'Para un OneDrive ya conocido'; To = 'For a known OneDrive' }
            @{ From = 'Detecta automáticamente SPO u OneDrive por la URL'; To = 'Automatically detects SPO or OneDrive from the URL' }
            @{ From = 'Selecciona el origen que quieres analizar.'; To = 'Select the source you want to analyze.' }
            @{ From = 'Conservar N históricas + versión actual'; To = 'Keep N historical versions + current version' }
            @{ From = 'Conservar N versiones totales'; To = 'Keep N total versions' }
            @{ From = 'Versión actual: conservar'; To = 'Current version: keep' }; @{ From = 'versiones históricas: conservar'; To = 'historical versions: keep' }
            @{ From = 'Las versiones históricas más antiguas serán elegibles para limpieza.'; To = 'The oldest historical versions will be eligible for cleanup.' }
            @{ From = 'Se revalidará cada versión antes de eliminarla.'; To = 'Each version will be revalidated before removal.' }
            @{ From = 'Si el historial cambió después del análisis, la versión se omitirá si deja de ser elegible.'; To = 'If the history changed after analysis, the version will be skipped if it is no longer eligible.' }
            @{ From = 'Las versiones se enviarán a la Papelera de reciclaje.'; To = 'Versions will be sent to the recycle bin.' }
            @{ From = 'Escriba ELIMINAR para continuar.'; To = 'Type DELETE to continue.' }; @{ From = 'Confirmación'; To = 'Confirmation' }
            @{ From = 'No se ha modificado ningún archivo.'; To = 'No files were modified.' }
            @{ From = 'La estimación no incluye'; To = 'The estimate does not include' }; @{ From = 'versión(es) sin tamaño disponible.'; To = 'version(s) without an available size.' }
            @{ From = 'Cada versión será revalidada contra el historial actual.'; To = 'Each version will be revalidated against the current history.' }
            @{ From = 'Configuración de throttling guardada.'; To = 'Throttling settings saved.' }
            @{ From = 'Opción inválida.'; To = 'Invalid option.' }; @{ From = 'Selección inválida.'; To = 'Invalid selection.' }; @{ From = 'Selección de bibliotecas'; To = 'Library selection' }
            @{ From = 'Marca al menos una biblioteca.'; To = 'Select at least one library.' }; @{ From = 'URL HTTPS no válida.'; To = 'Invalid HTTPS URL.' }
            @{ From = 'Formato esperado: nombre.onmicrosoft.com'; To = 'Expected format: name.onmicrosoft.com' }; @{ From = 'Tenant (ej. contoso.onmicrosoft.com)'; To = 'Tenant (e.g. contoso.onmicrosoft.com)' }
            @{ From = 'No existe un tenant válido configurado.'; To = 'No valid tenant is configured.' }; @{ From = 'PnP no devolvió una conexión válida.'; To = 'PnP did not return a valid connection.' }
            @{ From = 'No se encontraron bibliotecas documentales visibles.'; To = 'No visible document libraries were found.' }; @{ From = 'Obtener bibliotecas'; To = 'Get libraries' }
            @{ From = 'Biblioteca estándar (101)'; To = 'Standard library (101)' }; @{ From = 'Con historial'; To = 'With history' }
            @{ From = 'Se muestran los 50 archivos con más versiones a eliminar.'; To = 'The 50 files with the most versions to remove are shown.' }
            @{ From = 'Ya no cumple la política de eliminación'; To = 'No longer meets the removal policy' }; @{ From = 'La versión ya no existe'; To = 'The version no longer exists' }
            @{ From = 'El historial cambió durante la operación.'; To = 'The history changed during the operation.' }
            @{ From = 'Máximo de reintentos'; To = 'Maximum retries' }; @{ From = 'Espera máxima de retry (segundos)'; To = 'Maximum retry wait (seconds)' }
            @{ From = 'Revalida cada versión antes de eliminar'; To = 'Revalidate each version before removal' }; @{ From = 'Resumen por biblioteca'; To = 'Summary by library' }
            @{ From = 'Restablecer configuración'; To = 'Reset settings' }; @{ From = 'Requiere configuración'; To = 'Configuration required' }
            @{ From = 'Se eliminarán las preferencias locales.'; To = 'Local preferences will be deleted.' }; @{ From = 'No se encontraron bibliotecas documentales visibles.'; To = 'No visible document libraries were found.' }
            @{ From = 'Cada versión será revalidada contra el historial actual.'; To = 'Each version will be revalidated against the current history.' }
            @{ From = 'Aplicación Entra / autenticación PnP'; To = 'Entra application / PnP authentication' }
            @{ From = 'Validar aplicación configurada'; To = 'Validate configured application' }
            @{ From = 'Prueba autenticación real contra este tenant'; To = 'Test real authentication against this tenant' }
            @{ From = 'Usar otra aplicación existente'; To = 'Use another existing application' }
            @{ From = 'Usar una aplicación existente'; To = 'Use an existing application' }
            @{ From = 'Introduce un Client ID ya registrado'; To = 'Enter an already registered Client ID' }
            @{ From = 'Registrar una nueva aplicación Entra'; To = 'Register a new Entra application' }
            @{ From = 'Crea una app PnP nueva'; To = 'Create a new PnP app' }
            @{ From = 'Quitar Client ID de la configuración local'; To = 'Remove Client ID from local configuration' }
            @{ From = 'No elimina la app de Entra'; To = 'Does not delete the Entra app' }
            @{ From = 'Eliminar sesión persistente'; To = 'Clear persisted session' }
            @{ From = 'Limpia el token persistido de PnP'; To = 'Clears the persisted PnP token' }
            @{ From = 'Autenticación requerida'; To = 'Authentication required' }
            @{ From = 'Crear mediante PnP.PowerShell'; To = 'Create using PnP.PowerShell' }
            @{ From = 'Buscar sitio en el tenant'; To = 'Search for a site in the tenant' }
            @{ From = 'Busca por nombre, URL o propietario'; To = 'Search by name, URL, or owner' }
            @{ From = 'Buscar OneDrive en el tenant'; To = 'Search for OneDrive in the tenant' }
            @{ From = 'Busca por propietario o URL real'; To = 'Search by owner or real URL' }
            @{ From = 'Buscar en el tenant o usar URL directa'; To = 'Search in the tenant or use a direct URL' }
            @{ From = 'Buscar OneDrive reales del tenant'; To = 'Search for real OneDrive sites in the tenant' }
            @{ From = 'Retención'; To = 'Retention' }
            @{ From = 'Ejemplo: 10 = actual + 10 históricas'; To = 'Example: 10 = current + 10 historical' }
            @{ From = 'Ejemplo: 10 = actual + 9 históricas'; To = 'Example: 10 = current + 9 historical' }
            @{ From = 'Resultado del análisis'; To = 'Analysis result' }
            @{ From = 'Ejecutar limpieza'; To = 'Run cleanup' }
            @{ From = 'Ver archivos afectados'; To = 'View affected files' }
            @{ From = 'Hasta 50 archivos ordenados por impacto'; To = 'Up to 50 files sorted by impact' }
            @{ From = 'Abrir reportes'; To = 'Open reports' }
            @{ From = 'Sesión persistente'; To = 'Persisted session' }
            @{ From = 'Throttling y reintentos'; To = 'Throttling and retries' }
            @{ From = 'Inicio'; To = 'Home' }
            @{ From = 'Limpieza de historial de versiones'; To = 'Version history cleanup' }
            @{ From = 'Selecciona un sitio y una biblioteca para comenzar.'; To = 'Select a site and library to begin.' }
            @{ From = 'Seleccionar sitio y biblioteca'; To = 'Select site and library' }
            @{ From = 'Selecciona un sitio SharePoint o OneDrive.'; To = 'Select a SharePoint or OneDrive site.' }
            @{ From = 'Cambiar destino'; To = 'Change target' }
            @{ From = 'Limpiar historial de versiones'; To = 'Clean version history' }
            @{ From = 'Configuración'; To = 'Settings' }
            @{ From = 'SharePoint Online o OneDrive for Business'; To = 'SharePoint Online or OneDrive for Business' }
            @{ From = 'Introduce un número entre $Minimum y $Maximum.'; To = 'Enter a number between $Minimum and $Maximum.' }
            @{ From = 'SharePoint/OneDrive aplicó throttling.'; To = 'SharePoint/OneDrive applied throttling.' }
            @{ From = 'Operación: $OperationName'; To = 'Operation: $OperationName' }
            @{ From = 'Operación: '; To = 'Operation: ' }
            @{ From = 'PnP.PowerShell $version está instalado.'; To = 'PnP.PowerShell $version is installed.' }
            @{ From = 'PnP.PowerShell no está instalado.'; To = 'PnP.PowerShell is not installed.' }
            @{ From = 'PnP.PowerShell está listo.'; To = 'PnP.PowerShell is ready.' }
            @{ From = ' está instalado.'; To = ' is installed.' }; @{ From = ' no está instalado.'; To = ' is not installed.' }; @{ From = ' está listo.'; To = ' is ready.' }
            @{ From = ' es menor que '; To = ' is older than ' }
            @{ From = 'La aplicación pudo autenticarse correctamente.'; To = 'The application authenticated successfully.' }
            @{ From = 'No fue posible validar la aplicación.'; To = 'The application could not be validated.' }
            @{ From = 'Usar aplicación existente'; To = 'Use existing application' }
            @{ From = 'Client ID de la aplicación existente'; To = 'Client ID of the existing application' }
            @{ From = '¿Validar la aplicación antes de guardarla?'; To = 'Validate the application before saving it?' }
            @{ From = 'La configuración anterior se conservará.'; To = 'The previous configuration will be preserved.' }
            @{ From = '¿Mantener sesión autenticada?'; To = 'Keep the authenticated session?' }
            @{ From = 'Registrar nueva aplicación Entra'; To = 'Register new Entra application' }
            @{ From = 'Nombre de la aplicación [$defaultName]'; To = 'Application name [$defaultName]' }
            @{ From = '¿Registrar esta aplicación?'; To = 'Register this application?' }
            @{ From = '¿Validar la nueva aplicación ahora?'; To = 'Validate the new application now?' }
            @{ From = 'No se pudo registrar la aplicación: '; To = 'The application could not be registered: ' }
            @{ From = 'Client ID eliminado de la configuración local.'; To = 'Client ID removed from local configuration.' }
            @{ From = 'No hay una aplicación Entra válida configurada.'; To = 'No valid Entra application is configured.' }
            @{ From = 'Conectando al centro de administración: $adminUrl'; To = 'Connecting to the admin center: $adminUrl' }
            @{ From = 'Conectando al centro de administración: '; To = 'Connecting to the admin center: ' }
            @{ From = 'Filtra por nombre, URL o propietario. Deja vacío para mostrar todos si son 50 o menos.'; To = 'Filter by name, URL, or owner. Leave empty to show all when there are 50 or fewer.' }
            @{ From = '(sin título)'; To = '(untitled)' }
            @{ From = 'Actual + $historicalToKeep histórica(s)'; To = 'Current + $historicalToKeep historical version(s)' }
            @{ From = 'La versión ya no existe'; To = 'The version no longer exists' }
            @{ From = 'Eliminar versión '; To = 'Remove version ' }
            @{ From = ' de '; To = ' of ' }
            @{ From = 'Auditoría:'; To = 'Audit:' }
            @{ From = 'Versión $pnpVersion'; To = 'Version $pnpVersion' }
        )
        foreach ($translation in $phrases) { $result = $result.Replace($translation.From, $translation.To) }
        foreach ($translation in $script:UiTranslations) { $result = $result.Replace($translation.From, $translation.To) }
        $common = @(
            @{ From = 'Selecciona'; To = 'Select' }; @{ From = 'Seleccione'; To = 'Select' }; @{ From = 'Introducir'; To = 'Enter' }; @{ From = 'Introduzca'; To = 'Enter' }
            @{ From = 'Buscar'; To = 'Search' }; @{ From = 'Usar'; To = 'Use' }; @{ From = 'Cambiar'; To = 'Change' }; @{ From = 'Guardar'; To = 'Save' }
            @{ From = 'Validar'; To = 'Validate' }; @{ From = 'Registrar'; To = 'Register' }; @{ From = 'Obtener'; To = 'Get' }; @{ From = 'Limpiar'; To = 'Clean' }
            @{ From = 'Restablecer'; To = 'Reset' }; @{ From = 'Eliminar'; To = 'Remove' }; @{ From = 'Eliminación'; To = 'Removal' }; @{ From = 'Analizar'; To = 'Analyze' }
            @{ From = 'Análisis'; To = 'Analysis' }; @{ From = 'Procesando'; To = 'Processing' }; @{ From = 'Procesadas'; To = 'Processed' }
            @{ From = 'Aplicación'; To = 'Application' }; @{ From = 'Conexión'; To = 'Connection' }; @{ From = 'establecida'; To = 'established' }
            @{ From = 'Sitios'; To = 'Sites' }; @{ From = 'Sitio'; To = 'Site' }; @{ From = 'Bibliotecas'; To = 'Libraries' }; @{ From = 'Biblioteca'; To = 'Library' }
            @{ From = 'Archivos'; To = 'Files' }; @{ From = 'Archivo'; To = 'File' }; @{ From = 'Versiones'; To = 'Versions' }; @{ From = 'históricas'; To = 'historical' }
            @{ From = 'Conservar'; To = 'Keep' }; @{ From = 'conservar'; To = 'keep' }; @{ From = 'eliminar'; To = 'remove' }; @{ From = 'Históricas'; To = 'Historical' }
            @{ From = 'Estado'; To = 'Status' }; @{ From = 'Resultado'; To = 'Result' }; @{ From = 'Resumen'; To = 'Summary' }; @{ From = 'Reporte'; To = 'Report' }
            @{ From = 'Errores'; To = 'Errors' }; @{ From = 'Correctos'; To = 'Successful' }; @{ From = 'Omitidas'; To = 'Skipped' }; @{ From = 'Omitidos'; To = 'Skipped' }
            @{ From = 'No hay'; To = 'There are no' }; @{ From = 'No existe'; To = 'There is no' }; @{ From = 'No se encontraron'; To = 'No ... were found' }
            @{ From = 'No fue posible'; To = 'It was not possible' }; @{ From = 'no válido'; To = 'invalid' }; @{ From = 'válido'; To = 'valid' }
            @{ From = 'Operación cancelada'; To = 'Operation canceled' }; @{ From = 'Duración'; To = 'Duration' }; @{ From = 'Límite'; To = 'Limit' }
        )
        foreach ($translation in $common) { $result = $result.Replace($translation.From, $translation.To) }
        $words = @(
            @{ From = 'al'; To = 'when' }; @{ From = 'alguna'; To = 'some' }; @{ From = 'alguna(s)'; To = 'some' }; @{ From = 'actual'; To = 'current' }
            @{ From = 'administrador'; To = 'administrator' }; @{ From = 'archivo'; To = 'file' }; @{ From = 'archivos'; To = 'files' }; @{ From = 'afectados'; To = 'affected' }
            @{ From = 'analizar'; To = 'analyze' }; @{ From = 'análisis'; To = 'analysis' }; @{ From = 'buscar'; To = 'search' }; @{ From = 'cambiar'; To = 'change' }
            @{ From = 'con'; To = 'with' }; @{ From = 'conservar'; To = 'keep' }; @{ From = 'contenido'; To = 'content' }; @{ From = 'creada'; To = 'created' }
            @{ From = 'después'; To = 'after' }; @{ From = 'dentro'; To = 'within' }; @{ From = 'eliminar'; To = 'remove' }; @{ From = 'eliminación'; To = 'removal' }
            @{ From = 'en'; To = 'in' }; @{ From = 'entre'; To = 'between' }; @{ From = 'es'; To = 'is' }; @{ From = 'espera'; To = 'wait' }
            @{ From = 'estado'; To = 'status' }; @{ From = 'historial'; To = 'history' }; @{ From = 'históricas'; To = 'historical' }; @{ From = 'históricas'; To = 'historical' }
            @{ From = 'introducir'; To = 'enter' }; @{ From = 'límite'; To = 'limit' }; @{ From = 'máximo'; To = 'maximum' }; @{ From = 'menú'; To = 'menu' }
            @{ From = 'nueva'; To = 'new' }; @{ From = 'nuevo'; To = 'new' }; @{ From = 'opción'; To = 'option' }; @{ From = 'para'; To = 'for' }
            @{ From = 'papelera'; To = 'recycle bin' }; @{ From = 'procesando'; To = 'processing' }; @{ From = 'procesar'; To = 'process' }; @{ From = 'reporte'; To = 'report' }
            @{ From = 'reportes'; To = 'reports' }; @{ From = 'resultado'; To = 'result' }; @{ From = 'selecciona'; To = 'select' }; @{ From = 'seleccionar'; To = 'select' }
            @{ From = 'seleccionadas'; To = 'selected' }; @{ From = 'selección'; To = 'selection' }; @{ From = 'sitio'; To = 'site' }; @{ From = 'sitios'; To = 'sites' }
            @{ From = 'sin'; To = 'without' }; @{ From = 'sobre'; To = 'over' }; @{ From = 'total'; To = 'total' }; @{ From = 'versión'; To = 'version' }
            @{ From = 'versiones'; To = 'versions' }; @{ From = 'usuario'; To = 'user' }; @{ From = 'válido'; To = 'valid' }; @{ From = 'y'; To = 'and' }; @{ From = 'u'; To = 'or' }
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
        Clear-AppScreen
        Show-AppHeader -Section 'Idioma / Language'
        Write-Host ''
        Write-AppStyled -Text 'Select language / Seleccione idioma:' -Style Primary
        Write-Host ''
        Write-AppStyled -Text '  1  English' -Style Primary
        Write-Host ''
        Write-AppStyled -Text '  2  Español' -Style Primary
        Write-Host ''
        $choice = (Microsoft.PowerShell.Utility\Read-Host 'Choice / Opción').Trim()
        if ($choice -eq '1') { $script:Language = 'en'; return }; if ($choice -eq '2') { $script:Language = 'es'; return }
        Write-AppWarning -Message 'Elija 1 o 2.'
    }
}


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

    $Text = Get-LocalizedText $Text
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
        $width = [Math]::Min(84, [Math]::Max(36, $Host.UI.RawUI.WindowSize.Width - 2))
    }
    catch { $width = 72 }

    Clear-AppScreen
    Write-AppStyled -Text ('═' * $width) -Style Accent
    Write-AppStyled -Text "$script:AppName  v$script:AppVersion" -Style Primary
    if (-not [string]::IsNullOrWhiteSpace($Section)) { Write-AppStyled -Text "[$Section]" -Style Muted }
    Write-AppStyled -Text ('─' * $width) -Style Muted
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
function Write-AppSection {
    param([Parameter(Mandatory = $true)][string]$Title)
    Write-Host ''
    Write-AppStyled -Text $Title -Style Primary
}
function Write-AppField {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal'
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        $Value = '—'
    }

    Write-AppStyled -Text ('{0,-18}' -f $Name) -Style Muted -NoNewline
    Write-AppStyled -Text ([string]$Value) -Style $Style
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

    $yesToken = if ($script:Language -eq 'en') { 'Y' } else { 'S' }
    $noToken = 'N'
    $hint = if ($DefaultYes) { "$yesToken/n" } else { "$($yesToken.ToLowerInvariant())/$noToken" }

    while ($true) {
        $localizedPrompt = Get-LocalizedText $Prompt
        $answer = (Microsoft.PowerShell.Utility\Read-Host -Prompt "$localizedPrompt [$hint]").Trim()
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
        $numberMessage = if ($script:Language -eq 'en') {
            "Enter a number between $Minimum and $Maximum."
        }
        else {
            "Introduce un número entre $Minimum y $Maximum."
        }
        Write-AppWarning -Message $numberMessage
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
        if ([Uri]::TryCreate($url,[UriKind]::Absolute,[ref]$uri) -and
            $uri.Scheme -eq 'https' -and
            $uri.Host -match '(?i)\.sharepoint\.com$') {
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
        [string[]]$Description = @(),
        [scriptblock]$RenderBody
    )

    while ($true) {
        Clear-AppScreen
        Show-AppHeader -Section $Title

        if ($Description.Count -gt 0) {
            Write-Host ''
            foreach ($line in $Description) { Write-AppMuted -Message $line }
        }

        if ($null -ne $RenderBody) {
            & $RenderBody
        }

        Write-Host ''

        $normalItems = @($Items | Where-Object { [string]$_.Value -notin @('Back','Exit') })
        $zeroItems = @($Items | Where-Object { [string]$_.Value -in @('Back','Exit') })
        $map = @{}

        for ($i=0; $i -lt $normalItems.Count; $i++) {
            $key = [string]($i + 1)
            $item = $normalItems[$i]
            $map[$key] = $item
            Write-AppStyled -Text ('  {0,2}  ' -f $key) -Style Primary -NoNewline
            Write-AppStyled -Text ([string]$item.Label) -Style Normal
            $hintProp = $item.PSObject.Properties['Hint']
            $hint = if ($null -ne $hintProp -and -not [string]::IsNullOrWhiteSpace([string]$hintProp.Value)) { [string]$hintProp.Value } else { 'Sin descripción adicional.' }
            Write-AppMuted -Message "      $hint"
            Write-Host ''
        }

        if ($zeroItems.Count -gt 0) {
            $item = $zeroItems[0]
            $map['0'] = $item
            Write-AppStyled -Text '   0  ' -Style Primary -NoNewline
            Write-AppStyled -Text ([string]$item.Label) -Style Normal
            $hintProp = $item.PSObject.Properties['Hint']
            $hint = if ($null -ne $hintProp -and -not [string]::IsNullOrWhiteSpace([string]$hintProp.Value)) { [string]$hintProp.Value } else { 'Regresa al menú anterior.' }
            Write-AppMuted -Message "      $hint"
            Write-Host ''
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
        Write-AppStyled -Text ('  {0,2}  ' -f ($i + 1)) -Style Primary -NoNewline
        Write-AppStyled -Text ((& $Label $list[$i])) -Style Normal
        Write-Host ''
    }
    Write-AppStyled -Text '   0  ' -Style Primary -NoNewline
    Write-AppStyled -Text 'Cancelar' -Style Normal
    Write-Host ''

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
        $libraryInstructions = if ($script:Language -eq 'en') {
            '↑/↓ move   Space select   A select all   N select none   Enter accept   Esc cancel'
        }
        else {
            '↑/↓ mover   Espacio marcar   A todas   N ninguna   Enter aceptar   Esc cancelar'
        }
        Write-AppMuted -Message $libraryInstructions
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
        PersistLoginConfigured = $false
        Language = "es"

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

function Install-AppPnP {
    Show-AppHeader -Section "PnP.PowerShell"

    $version = Get-PnPInstalledVersion

    if ($null -ne $version -and $version -ge $script:MinimumPnPVersion) {
        Write-AppOk -Message "PnP.PowerShell $version está instalado."
        Write-Host ""

        if (-not (Read-AppYesNo -Prompt "¿Desea actualizarlo?" -DefaultYes $false)) {
            return $true
        }
    }
    elseif ($null -ne $version) {
        $olderVersionMessage = if ($script:Language -eq 'en') {
            "PnP.PowerShell $version is older than $script:MinimumPnPVersion."
        }
        else {
            "PnP.PowerShell $version es anterior a $script:MinimumPnPVersion."
        }
        Write-AppWarning -Message $olderVersionMessage
        Write-Host ""

        if (-not (Read-AppYesNo -Prompt "¿Instalar/actualizar ahora?" -DefaultYes $true)) {
            return $false
        }
    }
    else {
        Write-AppWarning -Message "PnP.PowerShell no está instalado."
        Write-Host ""

        if (-not (Read-AppYesNo -Prompt "¿Instalar/actualizar ahora?" -DefaultYes $true)) {
            return $false
        }
    }

    try {
        $params = @{
            Name = "PnP.PowerShell"
            Scope = "CurrentUser"
            Force = $true
            AllowClobber = $true
            MinimumVersion = $script:MinimumPnPVersion
            ErrorAction = "Stop"
        }

        Install-Module @params
        Import-Module "PnP.PowerShell" -MinimumVersion $script:MinimumPnPVersion -Force -ErrorAction Stop

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
    if (-not (Test-AppPowerShellVersion)) {
        return $false
    }

    $version = Get-PnPInstalledVersion
    if ($null -ne $version -and $version -ge $script:MinimumPnPVersion) {
        try {
            Import-Module "PnP.PowerShell" -MinimumVersion $script:MinimumPnPVersion -ErrorAction Stop
            return $true
        }
        catch {
            Show-AppErrorScreen -Title "PnP.PowerShell" -Message $_.Exception.Message
            return $false
        }
    }

    return Install-AppPnP
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
        [AllowNull()]$ReuseConnection,
        [switch]$ForceAuthentication
    )

    if ($ForceAuthentication) {
        Release-AppConnections -PreserveRecoveryState
        try {
            Disconnect-PnPOnline -ClearPersistedLogin -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    if ($null -eq $ReuseConnection -and $null -ne $script:AdminConnection) {
        $ReuseConnection = $script:AdminConnection
    }

    if ($null -ne $ReuseConnection) {
        try {
            $accessToken = Get-PnPAccessToken -ResourceTypeName SharePoint -Connection $ReuseConnection -ErrorAction Stop

            if ([string]::IsNullOrWhiteSpace([string]$accessToken)) {
                throw 'PnP no devolvió un token de SharePoint reutilizable.'
            }

            return Connect-PnPOnline -Url $Url -AccessToken ([string]$accessToken) -ReturnConnection -ErrorAction Stop
        }
        catch {
            if ($ForceAuthentication -or
                $script:AuthRecoveryUsed -or
                -not (Test-AppAuthenticationRecoveryError -Exception $_.Exception)) {
                throw
            }

            $script:AuthRecoveryUsed = $true
            Write-AppWarning -Message 'La sesión autenticada no pudo reutilizarse; se solicitará autenticación nuevamente una sola vez.'

            $reconnectParams = @{
                Url = $Url
                ClientId = $ClientId
                Tenant = $Tenant
                ForceAuthentication = $true
            }
            return New-AppPnPConnection @reconnectParams
        }
    }

    Ensure-AppPersistLoginPreference

    $params = @{
        Url = $Url
        ClientId = $ClientId
        Tenant = $Tenant
        Interactive = $true
        ReturnConnection = $true
        ValidateConnection = $true
        ErrorAction = 'Stop'
    }

    if ($ForceAuthentication) {
        $params.ForceAuthentication = $true
    }

    $config = Get-AppConfig
    if ([bool]$config.PersistLogin) {
        $params.PersistLogin = $true
    }

    $connection = Connect-PnPOnline @params

    if ($null -eq $script:AdminConnection) {
        $script:AdminConnection = $connection
    }

    return $connection
}

function Test-AppAuthenticationRecoveryError {
    param([Parameter(Mandatory = $true)][System.Exception]$Exception)

    $text = $Exception.ToString()
    return $text -match '(?i)(access denied|unauthorized|forbidden|invalid|expired|token|interactive|authentication|login|consent|AADSTS)'
}

function Ensure-AppPersistLoginPreference {
    $config = Get-AppConfig

    if ([bool]$config.PersistLoginConfigured) {
        return
    }

    Show-AppHeader -Section 'Sesión de autenticación'
    Write-AppMuted -Message 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'
    Write-AppMuted -Message 'Sesión persistente: si la activas, PnP puede reutilizar el inicio de sesión cuando abras el script nuevamente.'
    Write-AppMuted -Message 'Si la desactivas, solo se usa la sesión actual y podrás iniciar sesión de nuevo en la siguiente ejecución.'
    Write-Host ''

    $persistPrompt = @{
        Prompt = '¿Guardar la sesión para futuras ejecuciones?'
        DefaultYes = [bool]$config.PersistLogin
    }
    $config.PersistLogin = Read-AppYesNo @persistPrompt
    $config.PersistLoginConfigured = $true
    Save-AppConfig -Config $config
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
        $conn = New-AppPnPConnection -Url $adminUrl -ClientId $ClientId -Tenant $Tenant
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

    $configurationChanged = $tenant -ne [string]$config.Tenant -or $clientId -ne [string]$config.ClientId
    if ($configurationChanged) {
        Release-AppConnections
        $script:Target = $null
        $script:LastAnalysis = $null
        $script:LastAnalysisReports = $null
    }

    Write-Host ''
    $validate = Read-AppYesNo -Prompt '¿Validar la aplicación antes de guardarla?' -DefaultYes $true
    if ($validate -and -not (Test-AppRegistration -ClientId $clientId -Tenant $tenant)) {
        Write-Host ''
        Write-AppWarning -Message 'La configuración guardada no se modificó.'
        return $false
    }

    $config.ClientId = $clientId
    $config.Tenant = $tenant
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

        $configurationChanged = $tenant -ne [string]$config.Tenant -or $clientId -ne [string]$config.ClientId
        if ($configurationChanged) {
            Release-AppConnections
            $script:Target = $null
            $script:LastAnalysis = $null
            $script:LastAnalysisReports = $null
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
        $items += [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='Regresa al menú anterior.' }

        $choice = Show-NumberMenu -Title 'Aplicación Entra / autenticación PnP' -Items $items -RenderBody {
            Write-AppField -Name 'Tenant' -Value $(if ($config.Tenant) { $config.Tenant } else { 'No configurado' }) -Style $(if ($config.Tenant) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Client ID' -Value $(if ($config.ClientId) { $config.ClientId } else { 'No configurado' }) -Style $(if ($configured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Estado' -Value $(if ($configured) { 'CONFIGURADA' } else { 'NO CONFIGURADA' }) -Style $(if ($configured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
        }

        switch ($choice.Value) {
            'Validate' { [void](Test-AppRegistration -ClientId ([string]$config.ClientId) -Tenant ([string]$config.Tenant)); Wait-App }
            'Existing' { [void](Set-AppExistingClientId); Wait-App }
            'Register' { [void](New-AppEntraRegistration); Wait-App }
            'RemoveLocal' {
                if (Read-AppYesNo -Prompt '¿Quitar el Client ID guardado?' -DefaultYes $false) {
                    Release-AppConnections
                    $script:Target = $null
                    $script:LastAnalysis = $null
                    $script:LastAnalysisReports = $null
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
        [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='Regresa al menú anterior.' }
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
    $conn = New-AppPnPConnection -Url $adminUrl -ClientId ([string]$config.ClientId) -Tenant ([string]$config.Tenant)

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
    $items += [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='Regresa al menú anterior.' }

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
    $items += [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='Regresa al menú anterior.' }

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
        [PSCustomObject]@{ Label='Volver'; Value='Back'; Hint='Regresa al menú anterior.' }
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

    Release-AppSiteConnection
    $script:PnPConnection = New-AppPnPConnection `
        -Url ([string]$Target.Url) `
        -ClientId ([string]$config.ClientId) `
        -Tenant ([string]$config.Tenant)

    if ($null -eq $script:PnPConnection) { throw 'PnP no devolvió una conexión válida.' }

    return Get-PnPWeb -Connection $script:PnPConnection -Includes Title,Url -ErrorAction Stop
}

function Disconnect-AppM365 {
    Release-AppConnections
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
            Hint = "Regresa al menú anterior."
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

        Write-AppField -Name 'Biblioteca' -Value $library.Title -Style Primary
        $libraryProgress = if ($script:Language -eq 'en') {
            "Library $($libraryIndex + 1) of $($Libraries.Count)"
        }
        else {
            "Biblioteca $($libraryIndex + 1) de $($Libraries.Count)"
        }
        Write-AppMuted -Message $libraryProgress
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

            $fileProgress = if ($script:Language -eq 'en') {
                "$($fileIndex + 1) of $($files.Count) - $fileUrl"
            }
            else {
                "$($fileIndex + 1) de $($files.Count) - $fileUrl"
            }
            Write-Progress -Id 1 -Activity "Analizando $($library.Title)" -Status $fileProgress -PercentComplete $percent

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

    $duration = $Analysis.End - $Analysis.Start

    Write-AppField -Name 'Tipo' -Value $Target.Type -Style Muted
    Write-AppField -Name 'Destino' -Value $Target.Url -Style Primary
    Write-Host ''

    Write-AppField -Name 'Bibliotecas analizadas' -Value ('{0:N0}' -f $Analysis.Libraries) -Style Primary
    Write-AppField -Name 'Archivos analizados' -Value ('{0:N0}' -f $Analysis.Files) -Style Primary
    Write-AppField -Name 'Con historial' -Value ('{0:N0}' -f $Analysis.FilesWithHistory) -Style Primary
    Write-AppField -Name 'Archivos afectados' -Value ('{0:N0}' -f $Analysis.FilesAffected) -Style $(if ($Analysis.FilesAffected -gt 0) { 'Warning' } else { 'Success' })
    Write-Host ''

    Write-AppField -Name 'Versiones históricas' -Value ('{0:N0}' -f $Analysis.VersionsFound) -Style Primary
    Write-AppField -Name 'Históricas a conservar' -Value ('{0:N0}' -f $Analysis.VersionsKept) -Style Success
    Write-AppField -Name 'Históricas a eliminar' -Value ('{0:N0}' -f $Analysis.VersionsEligible) -Style $(if ($Analysis.VersionsEligible -gt 0) { 'Warning' } else { 'Success' })
    Write-Host ''

    if ($Analysis.VersionsWithKnownSize -gt 0) {
        $formattedSize = Format-AppBytes -Bytes $Analysis.EstimatedBytes
        Write-AppField -Name 'Espacio recuperable estimado' -Value $formattedSize -Style Primary

        if ($Analysis.VersionsWithoutKnownSize -gt 0) {
            Write-AppMuted -Message "La estimación no incluye $($Analysis.VersionsWithoutKnownSize) versión(es) sin tamaño disponible."
        }
    }
    else {
        Write-AppField -Name 'Espacio recuperable estimado' -Value 'No disponible' -Style Muted
    }

    Write-Host ""
    Write-AppField -Name 'Errores' -Value ('{0:N0}' -f $Analysis.Errors) -Style $(if ($Analysis.Errors -gt 0) { 'Danger' } else { 'Success' })
    Write-AppField -Name 'Eventos de throttling' -Value ('{0:N0}' -f $script:ThrottleEvents.Count) -Style $(if ($script:ThrottleEvents.Count -gt 0) { 'Warning' } else { 'Success' })
    Write-AppField -Name 'Duración' -Value $duration.ToString("hh\:mm\:ss") -Style Muted

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
    Write-AppMuted -Message "Se muestran hasta 50 archivos ordenados por cantidad de versiones elegibles."
    Write-Host ''

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

    if ($script:Language -eq 'en') {
        $rows | Format-Table `
            @{Label='Library'; Expression={$_.Biblioteca}},
            @{Label='File'; Expression={$_.Archivo}},
            @{Label='Historical'; Expression={$_.Historicas}},
            @{Label='Keep'; Expression={$_.Conserva}},
            @{Label='Remove'; Expression={$_.Elimina}},
            @{Label='Space'; Expression={$_.Espacio}} -AutoSize -Wrap
    }
    else {
        $rows | Format-Table Biblioteca, Archivo, Historicas, Conserva, Elimina, Espacio -AutoSize -Wrap
    }

    if ($Analysis.ByFile.Count -gt 50 -and $script:Language -eq 'es') {
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
    Write-AppMuted -Message "Resultados agrupados por biblioteca."
    Write-Host ''

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

    if ($script:Language -eq 'en') {
        $rows | Format-Table `
            @{Label='Library'; Expression={$_.Biblioteca}},
            @{Label='Template'; Expression={$_.Plantilla}},
            @{Label='Files'; Expression={$_.Archivos}},
            @{Label='Versions'; Expression={$_.Versiones}},
            @{Label='Affected'; Expression={$_.Afectados}},
            @{Label='Remove'; Expression={$_.Eliminar}},
            @{Label='Space'; Expression={$_.Espacio}},
            @{Label='Errors'; Expression={$_.Errores}} -AutoSize
    }
    else {
        $rows | Format-Table Biblioteca, Plantilla, Archivos, Versiones, Afectados, Eliminar, Espacio, Errores -AutoSize
    }

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
    Write-AppField -Name 'Archivos afectados' -Value ('{0:N0}' -f $Analysis.FilesAffected) -Style Primary
    Write-AppField -Name 'Versiones previstas' -Value ('{0:N0}' -f $Analysis.VersionsEligible) -Style Warning

    if ($Analysis.VersionsWithKnownSize -gt 0) {
        Write-AppField -Name 'Espacio estimado' -Value (Format-AppBytes -Bytes $Analysis.EstimatedBytes) -Style Primary
    }

    Write-Host ""
    Write-AppMuted -Message "Las versiones se enviarán a la Papelera de reciclaje."
    Write-Host ""

    $confirmationToken = if ($script:Language -eq 'en') { 'DELETE' } else { 'ELIMINAR' }
    $confirmationMessage = if ($script:Language -eq 'en') { 'Type DELETE to continue.' } else { 'Escriba ELIMINAR para continuar.' }
    Write-AppStyled -Text $confirmationMessage -Style Danger
    Write-Host ""

    $confirmation = (Microsoft.PowerShell.Utility\Read-Host -Prompt (Get-LocalizedText 'Confirmación')).Trim()

    return ($confirmation -ceq $confirmationToken)
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

        $cleanupProgress = if ($script:Language -eq 'en') {
            "$($i + 1) of $($entries.Count) - $($entry.Archivo)"
        }
        else {
            "$($i + 1) de $($entries.Count) - $($entry.Archivo)"
        }
        Write-Progress -Id 2 -Activity "Revalidando y eliminando versiones" -Status $cleanupProgress -PercentComplete $percent

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
        if ($IsWindows) {
            Start-Process -FilePath "explorer.exe" -ArgumentList $script:ReportsPath
        }
        else {
            Write-AppInfo -Message $script:ReportsPath
            Wait-App
        }
    }
    catch {
        Show-AppErrorScreen -Title "Reportes" -Message "No se pudo abrir la carpeta de reportes."
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

        $resolvedTitle = if (-not [string]::IsNullOrWhiteSpace([string]$web.Title)) {
            [string]$web.Title
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$target.Title)) {
            [string]$target.Title
        }
        else {
            [string]$target.Url
        }

        $script:Target = [PSCustomObject]@{
            Type = [string]$target.Type
            Url = [string]$target.Url
            Title = $resolvedTitle
            Owner = [string]$target.Owner
        }
        $target = $script:Target

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
        $script:LastAnalysis = $analysis
        $script:LastAnalysisReports = $reports

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
                Hint = "Muestra los resultados agrupados por biblioteca."
            }

            $items += [PSCustomObject]@{
                Label = "Abrir reportes"
                Value = "Reports"
                Hint = "Abre la carpeta con los CSV generados."
            }

            $items += [PSCustomObject]@{
                Label = "Volver"
                Value = "Back"
                Hint = "Regresa al menú anterior."
            }

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

                    Write-AppField -Name 'Versiones previstas' -Value ('{0:N0}' -f $cleanup.Attempted) -Style Primary
                    Write-AppField -Name 'Eliminadas' -Value ('{0:N0}' -f $cleanup.Removed) -Style Success
                    Write-AppField -Name 'Omitidas por revalidación' -Value ('{0:N0}' -f $cleanup.Skipped) -Style Warning
                    Write-AppField -Name 'Errores' -Value ('{0:N0}' -f $cleanup.Errors) -Style $(if ($cleanup.Errors -gt 0) { 'Danger' } else { 'Success' })
                    Write-AppField -Name 'Eventos de throttling' -Value ('{0:N0}' -f $script:ThrottleEvents.Count) -Style $(if ($script:ThrottleEvents.Count -gt 0) { 'Warning' } else { 'Success' })
                    Write-AppField -Name 'Duración' -Value $duration.ToString("hh\:mm\:ss") -Style Muted

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
}


# ============================================================
# CONFIGURACION
# ============================================================

function Show-AppContext {
    $config = Get-AppConfig
    $tenantConfigured = Test-AppTenant -Value ([string]$config.Tenant)
    $appConfigured = $tenantConfigured -and (Test-AppGuid -Value ([string]$config.ClientId))
    $contextActive = ($null -ne $script:Target) -or ($null -ne $script:PnPConnection) -or ($null -ne $script:LastAnalysis)

    Write-AppSection -Title 'Contexto'
    Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $config.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })

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
    Write-AppField -Name 'Análisis' -Value $(if ($null -ne $script:LastAnalysis) { 'DISPONIBLE' } else { 'NO DISPONIBLE' }) -Style $(if ($null -ne $script:LastAnalysis) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Contexto' -Value $(if ($contextActive) { 'ACTIVO' } else { 'VACÍO' }) -Style $(if ($contextActive) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
}

function Clear-AppWorkingContext {
    Show-AppHeader -Section 'Limpiar contexto de trabajo'

    Write-AppField -Name 'Destino' -Value $(if ($null -ne $script:Target) { $script:Target.Title } else { 'No seleccionado' }) -Style $(if ($null -ne $script:Target) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Conexión de sitio' -Value $(if ($null -ne $script:PnPConnection) { 'ACTIVA' } else { 'VACÍA' }) -Style $(if ($null -ne $script:PnPConnection) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Análisis' -Value $(if ($null -ne $script:LastAnalysis) { 'DISPONIBLE' } else { 'NO DISPONIBLE' }) -Style $(if ($null -ne $script:LastAnalysis) { 'Success' } else { 'Muted' })
    Write-Host ''
    Write-AppMuted -Message 'Limpia el destino, la conexión de sitio y los resultados del análisis. Conserva tenant, aplicación, preferencias y reportes.'
    Write-AppMuted -Message 'La sesión de autenticación base se conserva; cambiar tenant o aplicación la renovará.'

    if (-not (Read-AppYesNo -Prompt '¿Limpiar el contexto de trabajo?' -DefaultYes $false)) {
        return
    }

    Release-AppSiteConnection
    $script:Target = $null
    $script:LastAnalysis = $null
    $script:LastAnalysisReports = $null

    Write-AppOk -Message 'El contexto de trabajo fue limpiado.'
    Wait-App
}

function Show-AppContextMenu {
    while ($true) {
        $config = Get-AppConfig
        $contextActive = ($null -ne $script:Target) -or ($null -ne $script:PnPConnection) -or ($null -ne $script:LastAnalysis)

        $items = @(
            [PSCustomObject]@{ Label = 'Limpiar contexto de trabajo'; Value = 'Clear'; Hint = 'Limpia destino, conexión y resultados; conserva tenant, aplicación y reportes.' }
            [PSCustomObject]@{ Label = 'Cambiar tenant o aplicación conectada'; Value = 'Connection'; Hint = 'Abre configuración y libera la sesión anterior cuando cambien estos valores.' }
            [PSCustomObject]@{ Label = 'Abrir carpeta de reportes'; Value = 'Reports'; Hint = $script:ReportsPath }
            [PSCustomObject]@{ Label = 'Volver'; Value = 'Back'; Hint = 'Regresa al menú anterior.' }
        )

        $choice = Show-NumberMenu -Title 'Gestión del contexto' -Items $items -Description @(
            "Estado: $(if ($contextActive) { 'ACTIVO' } else { 'VACÍO' })",
            "Tenant: $(if ($config.Tenant) { $config.Tenant } else { 'No configurado' })",
            "Destino: $(if ($null -ne $script:Target) { $script:Target.Title } else { 'No seleccionado' })"
        )

        switch ($choice.Value) {
            'Clear' { Clear-AppWorkingContext }
            'Connection' { Show-AppSettings }
            'Reports' { Open-AppReports }
            'Back' { return }
        }
    }
}

function Show-AppDiagnostics {
    while ($true) {
        $config = Get-AppConfig
        $pnpVersion = Get-PnPInstalledVersion
        $tenantConfigured = Test-AppTenant -Value ([string]$config.Tenant)
        $appConfigured = $tenantConfigured -and (Test-AppGuid -Value ([string]$config.ClientId))

        $items = @(
            [PSCustomObject]@{ Label = 'Validar aplicación configurada'; Value = 'Validate'; Hint = 'Comprueba tenant, Client ID, autenticación y acceso al centro de administración.' }
            [PSCustomObject]@{ Label = 'Instalar / actualizar PnP.PowerShell'; Value = 'PnP'; Hint = 'Verifica el requisito local y ofrece instalar o actualizar el módulo.' }
            [PSCustomObject]@{ Label = 'Probar nuevamente el sitio actual'; Value = 'Target'; Hint = 'Valida la conexión y los permisos del destino seleccionado.' }
            [PSCustomObject]@{ Label = 'Volver'; Value = 'Back'; Hint = 'Regresa al menú anterior.' }
        )

        $choice = Show-NumberMenu -Title 'Diagnóstico' -Items $items -RenderBody {
            Write-AppField -Name 'PowerShell' -Value $PSVersionTable.PSVersion -Style $(if ($PSVersionTable.PSVersion -ge [version]'7.4') { 'Success' } else { 'Danger' })
            Write-AppField -Name 'PnP.PowerShell' -Value $(if ($null -ne $pnpVersion) { $pnpVersion } else { 'No instalado' }) -Style $(if ($null -ne $pnpVersion -and $pnpVersion -ge $script:MinimumPnPVersion) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $config.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Client ID' -Value $(if ($appConfigured) { $config.ClientId } else { 'No configurado' }) -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
            Write-AppField -Name 'Destino' -Value $(if ($null -ne $script:Target) { $script:Target.Title } else { 'No seleccionado' }) -Style $(if ($null -ne $script:Target) { 'Success' } else { 'Muted' })
            Write-AppField -Name 'Configuración' -Value $script:ConfigPath -Style Muted
            Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
        }

        switch ($choice.Value) {
            'Validate' {
                if (-not $appConfigured) {
                    Write-AppWarning -Message 'No hay una aplicación Entra válida configurada.'
                }
                else {
                    [void](Test-AppRegistration -ClientId ([string]$config.ClientId) -Tenant ([string]$config.Tenant))
                }
                Wait-App
            }
            'PnP' { [void](Install-AppPnP) }
            'Target' {
                if ($null -eq $script:Target) {
                    Write-AppWarning -Message 'No hay un sitio seleccionado.'
                }
                else {
                    try {
                        $web = Connect-AppM365 -Target $script:Target
                        Write-AppOk -Message 'Conexión y permisos del sitio correctos.'
                        Write-AppField -Name 'Sitio' -Value $web.Title -Style Primary
                    }
                    catch {
                        Write-AppError -Message $_.Exception.Message
                    }
                }
                Wait-App
            }
            'Back' { return }
        }
    }
}

function Show-AppSettings {
    while ($true) {
        $config = Get-AppConfig
        $tenantConfigured = Test-AppTenant -Value ([string]$config.Tenant)
        $appConfigured = $tenantConfigured -and (Test-AppGuid -Value ([string]$config.ClientId))
        $persistStatus = if ($config.PersistLogin) { 'Activada' } else { 'Desactivada' }
        $throttleHint = if ($script:Language -eq 'en') {
            "$($config.RequestDelayMs) ms | $($config.MaxRetries) retries | maximum $($config.RetryMaxSeconds) s"
        }
        else {
            "$($config.RequestDelayMs) ms | $($config.MaxRetries) reintentos | máximo $($config.RetryMaxSeconds) s"
        }

        $items = @(
            [PSCustomObject]@{
                Label = 'Cambiar tenant'
                Value = 'Tenant'
                Hint = 'Actualiza el tenant conectado y libera la sesión anterior.'
            },
            [PSCustomObject]@{
                Label = 'Aplicación Entra / autenticación PnP'
                Value = 'Auth'
                Hint = 'Valida, cambia o registra la aplicación usada por PnP.PowerShell.'
            },
            [PSCustomObject]@{
                Label = 'Cambiar idioma'
                Value = 'Language'
                Hint = if ($script:Language -eq 'en') { 'English' } else { 'Español' }
            },
            [PSCustomObject]@{
                Label = 'Alternar persistencia de login'
                Value = 'Persist'
                Hint = 'Controla si PnP puede reutilizar el login al abrir el script nuevamente.'
            },
            [PSCustomObject]@{
                Label = 'Throttling y reintentos'
                Value = 'Throttle'
                Hint = $throttleHint
            },
            [PSCustomObject]@{
                Label = 'Abrir carpeta de reportes'
                Value = 'Reports'
                Hint = $script:ReportsPath
            },
            [PSCustomObject]@{
                Label = 'Restablecer configuración local'
                Value = 'Reset'
                Hint = 'Borra valores guardados; conserva idioma y reportes.'
            },
            [PSCustomObject]@{
                Label = 'Volver'
                Value = 'Back'
                Hint = 'Regresa al menú anterior.'
            }
        )

        $choice = Show-NumberMenu -Title 'Configuración' -Items $items -RenderBody {
            Write-AppField -Name 'Tenant' -Value $(if ($tenantConfigured) { $config.Tenant } else { 'No configurado' }) -Style $(if ($tenantConfigured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Client ID' -Value $(if ($appConfigured) { $config.ClientId } else { 'No configurado' }) -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Aplicación' -Value $(if ($config.AppRegistrationName) { $config.AppRegistrationName } else { 'No configurada' }) -Style $(if ($appConfigured) { 'Success' } else { 'Warning' })
            Write-AppField -Name 'Idioma' -Value $(if ($script:Language -eq 'en') { 'English' } else { 'Español' }) -Style Primary
            Write-AppField -Name 'Persist login' -Value $persistStatus -Style $(if ($config.PersistLogin) { 'Success' } else { 'Muted' })
            Write-AppField -Name 'Sesión en memoria' -Value $(if ($null -ne $script:AdminConnection) { 'ACTIVA' } else { 'INACTIVA' }) -Style $(if ($null -ne $script:AdminConnection) { 'Success' } else { 'Muted' })
            Write-AppField -Name 'Reportes' -Value $script:ReportsPath -Style Muted
            Write-AppMuted -Message 'La sesión en memoria se reutiliza mientras la herramienta permanezca abierta.'
            Write-AppMuted -Message 'La persistencia controla si PnP reutiliza el inicio de sesión al abrir el script nuevamente.'
        }

        switch ($choice.Value) {
            'Tenant' {
                Clear-AppScreen
                Show-AppHeader -Section 'Tenant'
                $tenant = Read-AppTenant -DefaultValue ([string]$config.Tenant)

                if ($tenant -ne [string]$config.Tenant) {
                    Release-AppConnections
                    $script:Target = $null
                    $script:LastAnalysis = $null
                    $script:LastAnalysisReports = $null
                }

                $config.Tenant = $tenant.ToLowerInvariant()
                Save-AppConfig -Config $config
                Write-AppOk -Message 'Tenant actualizado.'
                Wait-App
            }

            "Auth" {
                Show-AppAuthenticationMenu
            }

            "Language" {
                Initialize-AppLanguage
                $config.Language = $script:Language
                Save-AppConfig -Config $config
            }

            "Persist" {
                $config.PersistLogin = -not [bool]$config.PersistLogin
                $config.PersistLoginConfigured = $true
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
                    $language = $script:Language
                    Release-AppConnections
                    $script:Target = $null
                    $script:LastAnalysis = $null
                    $script:LastAnalysisReports = $null

                    if (Test-Path -Path $script:ConfigPath) {
                        Remove-Item -Path $script:ConfigPath -Force
                    }

                    $newConfig = Get-AppConfig
                    $newConfig.Language = $language
                    Save-AppConfig -Config $newConfig
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
        $appConfigured = (Test-AppGuid -Value ([string]$config.ClientId)) -and (Test-AppTenant -Value ([string]$config.Tenant))
        $pnpStatus = if ($null -ne $pnpVersion) {
            if ($script:Language -eq 'en') { "Version $pnpVersion" } else { "Versión $pnpVersion" }
        }
        else {
            'No instalado'
        }
        $pnpHint = if ($script:Language -eq 'en') {
            "$pnpStatus · minimum required $script:MinimumPnPVersion"
        }
        else {
            "$pnpStatus · mínimo requerido $script:MinimumPnPVersion"
        }
        $contextStatus = if (($null -ne $script:Target) -or ($null -ne $script:PnPConnection) -or ($null -ne $script:LastAnalysis)) { 'ACTIVO' } else { 'VACÍO' }
        $authHint = if ($appConfigured) {
            if ($script:Language -eq 'en') { 'CONFIGURED · validate, change, or register another application.' } else { 'CONFIGURADA · valida, cambia o registra otra aplicación.' }
        }
        else {
            if ($script:Language -eq 'en') { 'NOT CONFIGURED · use an existing application or register a new one.' } else { 'NO CONFIGURADA · usa una aplicación existente o registra una nueva.' }
        }
        $contextHint = if ($script:Language -eq 'en') {
            "$contextStatus · clears the target, changes tenant/application, or opens reports."
        }
        else {
            "$contextStatus · limpia el destino, cambia tenant/aplicación o abre reportes."
        }

        $items = @(
            [PSCustomObject]@{
                Label = 'Nueva limpieza'
                Value = 'Cleanup'
                Hint = 'Analiza y limpia el historial de versiones de SharePoint o OneDrive.'
            },
            [PSCustomObject]@{
                Label = 'Aplicación Entra / autenticación PnP'
                Value = 'Auth'
                Hint = $authHint
            },
            [PSCustomObject]@{
                Label = 'PnP.PowerShell'
                Value = 'PnP'
                Hint = $pnpHint
            },
            [PSCustomObject]@{
                Label = 'Gestionar contexto de trabajo'
                Value = 'Context'
                Hint = $contextHint
            },
            [PSCustomObject]@{
                Label = 'Configuración'
                Value = 'Settings'
                Hint = 'Tenant, idioma, persistencia de login, throttling y reportes.'
            },
            [PSCustomObject]@{
                Label = 'Diagnóstico'
                Value = 'Diagnostics'
                Hint = 'Valida PowerShell, PnP, autenticación y acceso al destino.'
            },
            [PSCustomObject]@{
                Label = 'Salir'
                Value = 'Exit'
                Hint = 'Cierra la herramienta.'
            }
        )

        $choice = Show-NumberMenu -Title 'Inicio' -Items $items -Description @(
            'Limpieza segura del historial de versiones',
            'SharePoint Online y OneDrive for Business'
        ) -RenderBody {
            Show-AppContext
        }

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

            "Context" {
                Show-AppContextMenu
            }

            "Settings" {
                Show-AppSettings
            }

            "Diagnostics" {
                Show-AppDiagnostics
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
    Initialize-AppTerminal
    Initialize-AppLanguage
    try {
        $Host.UI.RawUI.WindowTitle = "$script:AppName $script:AppVersion"
    }
    catch {
    }

    Initialize-AppFolders

    if (-not (Test-AppPowerShellVersion)) {
        Write-Host ''
        Write-AppWarning -Message 'Ejecuta la herramienta desde PowerShell 7.4 o superior usando pwsh.'
        Wait-App
        exit 1
    }

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
