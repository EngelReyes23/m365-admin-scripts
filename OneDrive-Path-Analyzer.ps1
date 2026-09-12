<#
.SYNOPSIS
    OneDrive / SharePoint Path Analyzer - Windows

.DESCRIPTION
    Herramienta de troubleshooting para detectar rutas de usuario
    cercanas o superiores a los límites relevantes de
    OneDrive / SharePoint / Windows / Office.

    Compatible con:
      - Windows PowerShell 5.1
      - PowerShell 7+

    INCLUYE:
      - Archivos normales
      - Carpetas normales
      - Archivos ocultos normales del usuario

    EXCLUYE:
      - Elementos con atributo System
      - Metadatos internos conocidos
      - Archivos temporales de Office
      - OneDriveCloudTemp
      - OneDriveTemp

    FUNCIONES:
      - Detecta OneDrive Personal
      - Detecta OneDrive Business
      - Detecta bibliotecas SharePoint sincronizadas
      - Reconoce ubicaciones/shortcuts anidados
      - Evita doble escaneo
      - Escanea contenido de usuario
      - Muestra progreso
      - Exporta únicamente AVISO / CRITICO
      - No genera CSV si no existen hallazgos
#>

# ============================================================
# IDIOMA
# ============================================================

#requires -Version 5.1
$script:Language = 'es'
$script:UiTranslations = @(
    @{ From = 'Diagnóstico de rutas largas para troubleshooting de sincronización.'; To = 'Long-path diagnostics for synchronization troubleshooting.' }
    @{ From = 'Buscando ubicaciones sincronizadas...'; To = 'Searching for synchronized locations...' }
    @{ From = 'No se detectaron ubicaciones automáticamente.'; To = 'No locations were detected automatically.' }
    @{ From = 'Puede introducir una ruta manual.'; To = 'You can enter a path manually.' }
    @{ From = 'Ubicaciones detectadas'; To = 'Detected locations' }
    @{ From = 'Analizar todas'; To = 'Analyze all' }
    @{ From = 'Introducir ruta manual'; To = 'Enter a path manually' }
    @{ From = 'Ruta a analizar'; To = 'Path to analyze' }
    @{ From = 'No existen ubicaciones detectadas.'; To = 'No detected locations exist.' }
    @{ From = 'La ruta indicada no existe.'; To = 'The specified path does not exist.' }
    @{ From = 'Esta ubicación corresponde a contenido auxiliar y no será analizada.'; To = 'This location contains auxiliary content and will not be analyzed.' }
    @{ From = 'Use M para introducir una ruta manual.'; To = 'Use M to enter a path manually.' }
    @{ From = 'Selección no válida'; To = 'Invalid selection' }
    @{ From = 'Número fuera de rango'; To = 'Number out of range' }
    @{ From = 'Umbrales activos:'; To = 'Active thresholds:' }
    @{ From = 'Ruta de usuario más larga:'; To = 'Longest user path:' }
    @{ From = 'No se encontraron rutas cercanas o superiores a los límites.'; To = 'No paths near or over the limits were found.' }
    @{ From = 'Se encontraron rutas cercanas a los límites.'; To = 'Paths near the limits were found.' }
    @{ From = 'Se encontraron rutas que deberían corregirse.'; To = 'Paths that should be corrected were found.' }
    @{ From = 'No se generó CSV porque no existen rutas con AVISO o CRITICO.'; To = 'No CSV was generated because no WARNING or CRITICAL paths exist.' }
    @{ From = 'Análisis completado.'; To = 'Analysis completed.' }
    @{ From = 'Algunas carpetas no pudieron enumerarse.'; To = 'Some folders could not be enumerated.' }
    @{ From = 'El resultado podría no cubrir el 100 % del contenido.'; To = 'The result may not cover 100% of the content.' }
    @{ From = 'CSV generado:'; To = 'CSV generated:' }
    @{ From = 'ADVERTENCIA:'; To = 'WARNING:' }
    @{ From = 'RESULTADO:'; To = 'RESULT:' }
    @{ From = 'REQUIERE REVISION'; To = 'NEEDS REVIEW' }
    @{ From = 'CRITICO'; To = 'CRITICAL' }
    @{ From = 'AVISO'; To = 'WARNING' }
    @{ From = 'caracteres'; To = 'characters' }
    @{ From = 'bytes UTF-8'; To = 'UTF-8 bytes' }
    @{ From = 'Inicio'; To = 'Home' }
    @{ From = 'Finalizado'; To = 'Finished' }
    @{ From = 'Error inesperado'; To = 'Unexpected error' }
    @{ From = 'Detalles técnicos:'; To = 'Technical details:' }
    @{ From = 'Selecciona una opción'; To = 'Select an option' }
    @{ From = 'Selección'; To = 'Selection' }
    @{ From = 'Número'; To = 'Number' }
    @{ From = 'Cancelar'; To = 'Cancel' }
    @{ From = 'Volver'; To = 'Back' }
    @{ From = 'Configuración'; To = 'Settings' }
    @{ From = 'Salir'; To = 'Exit' }
    @{ From = 'Diagnóstico'; To = 'Diagnostics' }
    @{ From = 'Autenticación'; To = 'Authentication' }
    @{ From = 'Bibliotecas'; To = 'Libraries' }
    @{ From = 'Biblioteca'; To = 'Library' }
    @{ From = 'Sitio'; To = 'Site' }
    @{ From = 'SIMULACIÓN'; To = 'SIMULATION' }
    @{ From = 'REAL'; To = 'LIVE' }
    @{ From = 'Enter para continuar'; To = 'Press Enter to continue' }
    @{ From = 'Presione ENTER para cerrar'; To = 'Press ENTER to close' }
    @{ From = 'Shortcut / ubicación anidada'; To = 'Shortcut / nested location' }
    @{ From = 'Ubicaciones seleccionadas'; To = 'Selected locations' }
    @{ From = 'Raíz principal'; To = 'Main root' }
    @{ From = 'Raices físicas a recorrer'; To = 'Physical roots to scan' }
    @{ From = 'SELECCIONAR UBICACIONES'; To = 'SELECT LOCATIONS' }
    @{ From = 'LISTO PARA ANALIZAR'; To = 'READY TO ANALYZE' }
    @{ From = 'ANALIZANDO'; To = 'ANALYZING' }
    @{ From = 'RUTAS QUE REQUIEREN ATENCION'; To = 'PATHS REQUIRING ATTENTION' }
    @{ From = 'Analizando rutas de OneDrive / SharePoint'; To = 'Analyzing OneDrive / SharePoint paths' }
    @{ From = 'Windows : aviso desde 240 | crítico desde 256'; To = 'Windows : warning at 240 | critical at 256' }
    @{ From = 'Cloud   : aviso desde 360 | crítico sobre 400'; To = 'Cloud   : warning at 360 | critical over 400' }
    @{ From = 'Nombre  : aviso desde 240 | crítico sobre 255'; To = 'Name    : warning at 240 | critical over 255' }
    @{ From = 'Rutas de usuario analizadas'; To = 'Analyzed user paths' }
    @{ From = 'Rutas en AVISO'; To = 'WARNING paths' }
    @{ From = 'Rutas CRITICAS'; To = 'CRITICAL paths' }
    @{ From = 'Errores de lectura'; To = 'Read errors' }
    @{ From = 'Ruta de usuario más larga:'; To = 'Longest user path:' }
    @{ From = 'RESULTADO: CRITICO'; To = 'RESULT: CRITICAL' }
    @{ From = 'RESULTADO: REQUIERE REVISION'; To = 'RESULT: NEEDS REVIEW' }
    @{ From = 'RESULTADO: OK'; To = 'RESULT: OK' }
    @{ From = 'No se generó CSV porque no existen rutas con AVISO o CRITICO.'; To = 'No CSV was generated because no WARNING or CRITICAL paths exist.' }
    @{ From = 'Gestionar selección'; To = 'Manage selection' }
    @{ From = 'Analizar ubicaciones seleccionadas'; To = 'Analyze selected locations' }
    @{ From = 'Configurar umbrales'; To = 'Configure thresholds' }
    @{ From = 'Abrir carpeta de reportes'; To = 'Open reports folder' }
    @{ From = 'Limpiar selección'; To = 'Clear selection' }
    @{ From = 'Cerrar'; To = 'Exit' }
    @{ From = 'Estado'; To = 'Status' }
    @{ From = 'Selección actual'; To = 'Current selection' }
    @{ From = 'No seleccionadas'; To = 'None selected' }
    @{ From = 'Detectadas'; To = 'Detected' }
    @{ From = 'Último resultado'; To = 'Last result' }
    @{ From = 'Sin análisis'; To = 'No analysis' }
    @{ From = 'No generado'; To = 'Not generated' }
    @{ From = 'Selección de ubicaciones'; To = 'Location selection' }
    @{ From = 'Selecciona números separados por coma, A para todas, M para manual o 0 para volver.'; To = 'Enter comma-separated numbers, A for all, M for manual, or 0 to go back.' }
    @{ From = 'Regresa al menú anterior.'; To = 'Return to the previous menu.' }
    @{ From = 'No hay ubicaciones seleccionadas.'; To = 'No locations are selected.' }
    @{ From = 'Selecciona ubicaciones antes de iniciar el análisis.'; To = 'Select locations before starting the analysis.' }
    @{ From = 'La selección fue limpiada.'; To = 'The selection was cleared.' }
    @{ From = '¿Limpiar la selección actual?'; To = 'Clear the current selection?' }
    @{ From = '¿Iniciar el análisis?'; To = 'Start the analysis?' }
    @{ From = 'Umbrales activos'; To = 'Active thresholds' }
    @{ From = 'Windows / Office'; To = 'Windows / Office' }
    @{ From = 'Cloud / SharePoint'; To = 'Cloud / SharePoint' }
    @{ From = 'OneDrive Sync'; To = 'OneDrive Sync' }
    @{ From = 'Nombre individual'; To = 'Individual name' }
    @{ From = 'Aviso desde'; To = 'Warning at' }
    @{ From = 'Crítico desde'; To = 'Critical at' }
    @{ From = 'Crítico sobre'; To = 'Critical over' }
    @{ From = 'Valores configurables para detectar rutas cercanas o superiores a los límites.'; To = 'Configurable values used to detect paths near or over the limits.' }
    @{ From = 'Umbrales guardados en memoria.'; To = 'Thresholds saved in memory.' }
    @{ From = 'Diagnóstico del sistema y del analizador.'; To = 'System and analyzer diagnostics.' }
    @{ From = 'Sistema operativo'; To = 'Operating system' }
    @{ From = 'Versión PowerShell'; To = 'PowerShell version' }
    @{ From = 'Ruta de reportes'; To = 'Report path' }
    @{ From = 'Carpeta de reportes'; To = 'Reports folder' }
    @{ From = 'No se pudo abrir la carpeta de reportes.'; To = 'The reports folder could not be opened.' }
    @{ From = 'Análisis'; To = 'Analysis' }
    @{ From = 'Hallazgos'; To = 'Findings' }
    @{ From = 'El análisis no ha comenzado.'; To = 'Analysis has not started.' }
    @{ From = 'La ruta manual fue agregada.'; To = 'The manual path was added.' }
    @{ From = 'La ruta fue agregada.'; To = 'The path was added.' }
    @{ From = 'Ruta manual'; To = 'Manual path' }
    @{ From = 'Ruta local Windows/Office >= '; To = 'Local Windows/Office path >= ' }
    @{ From = 'Ruta local cerca del limite Windows'; To = 'Local path near the Windows limit' }
    @{ From = 'Ruta cloud > '; To = 'Cloud path > ' }
    @{ From = 'Ruta cloud cerca de '; To = 'Cloud path near ' }
    @{ From = 'OneDrive Sync > '; To = 'OneDrive Sync > ' }
    @{ From = 'OneDrive Sync cerca de '; To = 'OneDrive Sync near ' }
    @{ From = 'Nombre individual > '; To = 'Individual name > ' }
    @{ From = 'Nombre individual cerca de '; To = 'Individual name near ' }
    @{ From = 'Seleccionar ubicaciones'; To = 'Select locations' }
    @{ From = 'Elige ubicaciones detectadas o introduce una ruta manual.'; To = 'Choose detected locations or enter a manual path.' }
    @{ From = 'Escanea el contenido local y reporta rutas AVISO o CRÍTICO.'; To = 'Scans local content and reports WARNING or CRITICAL paths.' }
    @{ From = 'Quita la selección y el resultado en memoria; no borra archivos ni CSV.'; To = 'Clears the in-memory selection and result; does not delete files or CSVs.' }
    @{ From = 'Ajusta los umbrales de Windows, Cloud, Sync y nombres.'; To = 'Adjusts the Windows, Cloud, Sync, and name thresholds.' }
    @{ From = 'Muestra información del sistema y actualiza detecciones.'; To = 'Shows system information and refreshes detections.' }
    @{ From = 'Abre el escritorio o la carpeta del último CSV.'; To = 'Opens the desktop or the folder containing the last CSV.' }
    @{ From = 'Todavía no existe un resultado.'; To = 'No result exists yet.' }
    @{ From = 'Disponible para consultar.'; To = 'Available to view.' }
    @{ From = 'Analiza contenido local de OneDrive y SharePoint sincronizado.'; To = 'Analyzes synchronized OneDrive and SharePoint content locally.' }
    @{ From = 'Analiza contenido local y reporta rutas largas.'; To = 'Analyzes local content and reports long paths.' }
    @{ From = 'Ajusta los umbrales usados para clasificar las rutas.'; To = 'Adjusts the thresholds used to classify paths.' }
    @{ From = 'Los cambios se mantienen durante esta ejecución.'; To = 'Changes remain in effect for this run.' }
    @{ From = 'Configura los umbrales de longitud de ruta local.'; To = 'Configures local path-length thresholds.' }
    @{ From = 'Configura los umbrales de la ruta relativa en la nube.'; To = 'Configures cloud-relative path thresholds.' }
    @{ From = 'Configura los umbrales de sincronización local.'; To = 'Configures local synchronization thresholds.' }
    @{ From = 'Configura el límite del nombre de archivo o carpeta.'; To = 'Configures the file or folder name limit.' }
    @{ From = 'Cambiar idioma'; To = 'Change language' }
    @{ From = 'Restablecer umbrales predeterminados'; To = 'Reset default thresholds' }
    @{ From = 'Restaura los valores recomendados para Windows, Cloud, Sync y nombres.'; To = 'Restores recommended values for Windows, Cloud, Sync, and names.' }
    @{ From = 'Restablecer umbrales'; To = 'Reset thresholds' }
    @{ From = 'Se restaurarán los umbrales recomendados.'; To = 'Recommended thresholds will be restored.' }
    @{ From = '¿Restablecer los umbrales?'; To = 'Reset thresholds?' }
    @{ From = 'Umbrales restablecidos.'; To = 'Thresholds reset.' }
    @{ From = 'Elija 1 o 2.'; To = 'Choose 1 or 2.' }
    @{ From = 'Responde S o N.'; To = 'Answer Y or N.' }
    @{ From = 'Opción inválida.'; To = 'Invalid option.' }
    @{ From = 'Sin descripción adicional.'; To = 'No additional description.' }
    @{ From = 'Limpiar selección / contexto'; To = 'Clear selection / working context' }
    @{ From = 'Limpiar contexto de trabajo'; To = 'Clear working context' }
    @{ From = 'Ver último resultado'; To = 'View last result' }
    @{ From = 'Cierra la herramienta.'; To = 'Closes the tool.' }
    @{ From = 'Limpia las ubicaciones seleccionadas y el resultado del análisis. No elimina archivos ni reportes existentes.'; To = 'Clears selected locations and the analysis result. Does not delete existing files or reports.' }
    @{ From = 'La selección no contiene ubicaciones válidas.'; To = 'The selection does not contain valid locations.' }
    @{ From = 'Actualizar ubicaciones detectadas'; To = 'Refresh detected locations' }
    @{ From = 'Resultado'; To = 'Result' }
    @{ From = 'Listo para analizar'; To = 'Ready to analyze' }
    @{ From = 'Analizando'; To = 'Analyzing' }
    @{ From = 'Contexto'; To = 'Context' }
    @{ From = 'Reportes'; To = 'Reports' }
    @{ From = 'Ubicaciones'; To = 'Locations' }
    @{ From = 'Último análisis'; To = 'Last analysis' }
    @{ From = 'Análisis completado'; To = 'Analysis completed' }
    @{ From = 'Avisos'; To = 'Warnings' }
    @{ From = 'Criticas'; To = 'Critical' }
    @{ From = 'Archivo'; To = 'File' }
    @{ From = 'Carpeta'; To = 'Folder' }
    @{ From = 'Seleccionar por número'; To = 'Select by number' }
    @{ From = 'Elige una o varias ubicaciones separadas por coma.'; To = 'Choose one or more comma-separated locations.' }
    @{ From = 'Incluye todas las ubicaciones detectadas.'; To = 'Includes all detected locations.' }
    @{ From = 'Analiza una carpeta local que no fue detectada automáticamente.'; To = 'Analyzes a local folder that was not detected automatically.' }
    @{ From = 'Selecciona las ubicaciones locales que quieres analizar.'; To = 'Select the local locations to analyze.' }
    @{ From = 'Vuelve a consultar las ubicaciones sincronizadas del equipo.'; To = 'Re-queries synchronized locations on this computer.' }
    @{ From = 'Abre el escritorio o la carpeta del último CSV generado.'; To = 'Opens the desktop or the folder containing the last generated CSV.' }
    @{ From = 'El analizador es local y no requiere tenant, aplicación ni autenticación.'; To = 'The analyzer is local and does not require a tenant, application, or authentication.' }
    @{ From = 'Rutas analizadas'; To = 'Paths analyzed' }
    @{ From = 'Rutas CRÍTICAS'; To = 'CRITICAL paths' }
    @{ From = 'Ruta más larga'; To = 'Longest path' }
    @{ From = 'Ubicación'; To = 'Location' }
    @{ From = 'Modo'; To = 'Mode' }
    @{ From = 'LECTURA LOCAL'; To = 'LOCAL READ-ONLY' }
    @{ From = 'DISPONIBLE'; To = 'AVAILABLE' }
    @{ From = 'NO DISPONIBLE'; To = 'NOT AVAILABLE' }
    @{ From = 'Reporte'; To = 'Report' }
    @{ From = 'Rutas que requieren atención'; To = 'Paths requiring attention' }
    @{ From = 'Aviso desde caracteres'; To = 'Warning at characters' }
    @{ From = 'Crítico desde caracteres'; To = 'Critical at characters' }
    @{ From = 'Crítico sobre caracteres'; To = 'Critical over characters' }
    @{ From = 'No modifica archivos; solo analiza longitudes y exporta hallazgos.'; To = 'Does not modify files; only analyzes lengths and exports findings.' }
    @{ From = 'Reporte generado'; To = 'Report generated' }
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
    if ($null -eq $Text) { return '' }
    $result = [string]$Text
    if ($script:Language -eq 'en') {
        foreach ($translation in $script:UiTranslations) { $result = $result.Replace($translation.From, $translation.To) }
        $common = @(
            @{ From = 'Raices físicas a recorrer'; To = 'Physical roots to scan' }; @{ From = 'LISTO PARA ANALIZAR'; To = 'READY TO ANALYZE' }
            @{ From = 'SELECCIONAR UBICACIONES'; To = 'SELECT LOCATIONS' }; @{ From = 'RUTAS QUE REQUIEREN ATENCION'; To = 'PATHS REQUIRING ATTENTION' }
            @{ From = 'Selecciona'; To = 'Select' }; @{ From = 'Seleccione'; To = 'Select' }; @{ From = 'Introducir'; To = 'Enter' }
            @{ From = 'Introduzca'; To = 'Enter' }; @{ From = 'Ubicación'; To = 'Location' }; @{ From = 'ubicación'; To = 'location' }
            @{ From = 'Ruta'; To = 'Path' }; @{ From = 'rutas'; To = 'paths' }; @{ From = 'Rutas'; To = 'Paths' }
            @{ From = 'Estado'; To = 'Status' }; @{ From = 'Motivo'; To = 'Reason' }; @{ From = 'Tipo'; To = 'Type' }
            @{ From = 'Nombre'; To = 'Name' }; @{ From = 'Longitud'; To = 'Length' }; @{ From = 'analizadas'; To = 'analyzed' }
            @{ From = 'analizar'; To = 'analyze' }; @{ From = 'Analizar'; To = 'Analyze' }; @{ From = 'automáticamente'; To = 'automatically' }
            @{ From = 'desde'; To = 'from' }; @{ From = 'sobre'; To = 'over' }; @{ From = 'cerca de'; To = 'near' }
            @{ From = 'individual'; To = 'individual' }; @{ From = 'Contenido'; To = 'Content' }; @{ From = 'auxiliar'; To = 'auxiliary' }
            @{ From = 'No se'; To = 'No ' }; @{ From = 'se encontraron'; To = 'were found' }; @{ From = 'se generó'; To = 'was generated' }
            @{ From = 'podría no cubrir'; To = 'may not cover' }; @{ From = 'La ruta indicada'; To = 'The specified path' }
            @{ From = 'no existe'; To = 'does not exist' }; @{ From = 'no será analizada'; To = 'will not be analyzed' }
            @{ From = 'selección'; To = 'selection' }; @{ From = 'Número'; To = 'Number' }; @{ From = 'Ejemplo'; To = 'Example' }
            @{ From = 'Varias'; To = 'Multiple' }; @{ From = 'Críticas'; To = 'Critical' }; @{ From = 'crítico'; To = 'critical' }
        )
        foreach ($translation in $common) { $result = $result.Replace($translation.From, $translation.To) }
        $words = @(
            @{ From = 'al'; To = 'when' }; @{ From = 'alguna'; To = 'some' }; @{ From = 'algunas'; To = 'some' }; @{ From = 'actual'; To = 'current' }
            @{ From = 'administrador'; To = 'administrator' }; @{ From = 'administración'; To = 'administration' }; @{ From = 'ahora'; To = 'now' }
            @{ From = 'archivo'; To = 'file' }; @{ From = 'archivos'; To = 'files' }; @{ From = 'abrir'; To = 'open' }; @{ From = 'acceso'; To = 'access' }
            @{ From = 'afectado'; To = 'affected' }; @{ From = 'analizadas'; To = 'analyzed' }; @{ From = 'analizar'; To = 'analyze' }; @{ From = 'atención'; To = 'attention' }
            @{ From = 'automáticamente'; To = 'automatically' }; @{ From = 'buscar'; To = 'search' }; @{ From = 'carpeta'; To = 'folder' }; @{ From = 'carpetas'; To = 'folders' }
            @{ From = 'cambiar'; To = 'change' }; @{ From = 'cantidad'; To = 'amount' }; @{ From = 'con'; To = 'with' }; @{ From = 'contenido'; To = 'content' }
            @{ From = 'correcto'; To = 'correct' }; @{ From = 'crítico'; To = 'critical' }; @{ From = 'datos'; To = 'data' }; @{ From = 'después'; To = 'after' }
            @{ From = 'dentro'; To = 'within' }; @{ From = 'desde'; To = 'from' }; @{ From = 'elemento'; To = 'item' }; @{ From = 'elementos'; To = 'items' }
            @{ From = 'en'; To = 'in' }; @{ From = 'entre'; To = 'between' }; @{ From = 'es'; To = 'is' }; @{ From = 'estado'; To = 'status' }
            @{ From = 'existe'; To = 'exists' }; @{ From = 'externo'; To = 'external' }; @{ From = 'filtro'; To = 'filter' }; @{ From = 'generado'; To = 'generated' }
            @{ From = 'guardar'; To = 'save' }; @{ From = 'hasta'; To = 'up to' }; @{ From = 'individual'; To = 'individual' }; @{ From = 'inicial'; To = 'initial' }
            @{ From = 'introducir'; To = 'enter' }; @{ From = 'límite'; To = 'limit' }; @{ From = 'límites'; To = 'limits' }; @{ From = 'lista'; To = 'list' }
            @{ From = 'máximo'; To = 'maximum' }; @{ From = 'máximo'; To = 'maximum' }; @{ From = 'mínimo'; To = 'minimum' }; @{ From = 'nombre'; To = 'name' }
            @{ From = 'nueva'; To = 'new' }; @{ From = 'nuevo'; To = 'new' }; @{ From = 'opción'; To = 'option' }; @{ From = 'opciones'; To = 'options' }
            @{ From = 'para'; To = 'for' }; @{ From = 'permiso'; To = 'permission' }; @{ From = 'permisos'; To = 'permissions' }; @{ From = 'primero'; To = 'first' }
            @{ From = 'puede'; To = 'can' }; @{ From = 'reporte'; To = 'report' }; @{ From = 'reportes'; To = 'reports' }; @{ From = 'resultado'; To = 'result' }
            @{ From = 'revisar'; To = 'review' }; @{ From = 'seleccionado'; To = 'selected' }; @{ From = 'seleccionados'; To = 'selected' }; @{ From = 'seleccionar'; To = 'select' }
            @{ From = 'selección'; To = 'selection' }; @{ From = 'selecciones'; To = 'selections' }; @{ From = 'sin'; To = 'without' }; @{ From = 'sobre'; To = 'over' }
            @{ From = 'suficiente'; To = 'sufficient' }; @{ From = 'tipo'; To = 'type' }; @{ From = 'total'; To = 'total' }; @{ From = 'usuario'; To = 'user' }
            @{ From = 'usuarios'; To = 'users' }; @{ From = 'válido'; To = 'valid' }; @{ From = 'válida'; To = 'valid' }; @{ From = 'valor'; To = 'value' }
            @{ From = 'versiones'; To = 'versions' }; @{ From = 'versión'; To = 'version' }; @{ From = 'volver'; To = 'back' }; @{ From = 'y'; To = 'and' }; @{ From = 'u'; To = 'or' }
        )
        foreach ($translation in $words) {
            $pattern = '(?<![\p{L}])' + [regex]::Escape($translation.From) + '(?![\p{L}])'
            $result = [regex]::Replace($result, $pattern, $translation.To)
        }
    }
    return $result
}

function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)][object[]]$Object,
        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor,
        [switch]$NoNewline
    )
    $text = if ($null -eq $Object) { '' } else { ($Object | ForEach-Object { [string]$_ }) -join ' ' }
    $parameters = @{ Object = (Get-LocalizedText $text) }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $parameters.ForegroundColor = $ForegroundColor }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $parameters.BackgroundColor = $BackgroundColor }
    if ($NoNewline) { $parameters.NoNewline = $true }
    Microsoft.PowerShell.Utility\Write-Host @parameters
}

function Read-Host {
    param([Parameter(Position = 0)][string]$Prompt, [switch]$AsSecureString)
    $localizedPrompt = Get-LocalizedText $Prompt
    if ($AsSecureString) { return Microsoft.PowerShell.Utility\Read-Host -Prompt $localizedPrompt -AsSecureString }
    return Microsoft.PowerShell.Utility\Read-Host -Prompt $localizedPrompt
}

function Write-Progress {
    [CmdletBinding()] param([int]$Id = 0,[string]$Activity,[string]$Status,[string]$CurrentOperation,[int]$PercentComplete,[int]$SecondsRemaining,[switch]$Completed)
    $parameters = @{ Id = $Id }
    if ($PSBoundParameters.ContainsKey('Activity')) { $parameters.Activity = Get-LocalizedText $Activity }
    if ($PSBoundParameters.ContainsKey('Status')) { $parameters.Status = Get-LocalizedText $Status }
    if ($PSBoundParameters.ContainsKey('CurrentOperation')) { $parameters.CurrentOperation = Get-LocalizedText $CurrentOperation }
    if ($PSBoundParameters.ContainsKey('PercentComplete')) { $parameters.PercentComplete = $PercentComplete }
    if ($PSBoundParameters.ContainsKey('SecondsRemaining')) { $parameters.SecondsRemaining = $SecondsRemaining }
    if ($Completed) { $parameters.Completed = $true }
    Microsoft.PowerShell.Utility\Write-Progress @parameters
}

function Initialize-AppLanguage {
    param([string]$Default = 'en')
    Initialize-ToolkitLanguage -Default $Default -Name 'OneDrive / SharePoint Path Analyzer' -Version '2.0'
}

# ============================================================
# CONFIGURACIÓN
# ============================================================

# Windows / Office
$script:WARN_WINDOWS      = 240
$script:CRITICAL_WINDOWS  = 256

# OneDrive / SharePoint Cloud
$script:WARN_CLOUD        = 360
$script:LIMIT_CLOUD       = 400

# OneDrive Sync
$script:WARN_SYNC         = 480
$script:LIMIT_SYNC        = 520

# Nombre individual
$script:WARN_NAME         = 240
$script:LIMIT_NAME        = 255


# ============================================================
# ELEMENTOS AUXILIARES
# ============================================================

$ExcludedExactNames = @(
    "OneDriveCloudTemp",
    "OneDriveTemp",
    "desktop.ini",
    "thumbs.db",
    "ehthumbs.db",
    ".DS_Store"
)


# ============================================================
# INTERFAZ
# ============================================================

function Write-Banner {
    Clear-AppScreen
    Show-AppHeader -Section 'Inicio'
    Write-AppMuted -Message 'Diagnóstico de rutas largas para troubleshooting de sincronización.'
}


function Write-Section {
    param([string]$Title)
    Write-ToolkitText ''
    Write-ToolkitText (Get-LocalizedText $Title) Primary
}


function Write-StatusLine {
    param(
        [string]$Label,
        [AllowNull()]$Value,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal'
    )

    Write-AppField -Name $Label -Value $Value -Style $Style
}

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

    if (-not $script:Ansi) { return '' }

    $esc = [char]27
    switch ($Name) {
        'Reset' { "$esc[0m" }
        'Bold' { "$esc[1m" }
        'Dim' { "$esc[2m" }
        'Cyan' { "$esc[36m" }
        'Green' { "$esc[92m" }
        'Yellow' { "$esc[93m" }
        'Red' { "$esc[91m" }
        'Magenta' { "$esc[95m" }
    }
}

function Clear-AppScreen {
    try { if (-not [Console]::IsOutputRedirected) { Clear-Host } } catch { }
}

function Write-AppStyled {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal',
        [switch]$NoNewline
    )
    Write-ToolkitText (Get-LocalizedText $Text) $Style -NoNewline:$NoNewline
}

function Show-AppHeader {
    param([string]$Section = 'Inicio')
    Write-ToolkitHeader 'OneDrive / SharePoint Path Analyzer' '2.0' (Get-LocalizedText $Section)
}

function Write-AppSection {
    param([Parameter(Mandatory = $true)][string]$Title)
    Write-ToolkitText ''
    Write-ToolkitText (Get-LocalizedText $Title) Primary
}

function Write-AppField {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value,
        [ValidateSet('Normal','Muted','Primary','Success','Warning','Danger','Accent')]
        [string]$Style = 'Normal'
    )
    Write-ToolkitField (Get-LocalizedText $Name) $Value $Style
}

function Write-AppOk {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-ToolkitStatus Ok (Get-LocalizedText $Message)
}

function Write-AppInfo {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-ToolkitStatus Info (Get-LocalizedText $Message)
}

function Write-AppWarning {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-ToolkitStatus Warn (Get-LocalizedText $Message)
}

function Write-AppError {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-ToolkitStatus Error (Get-LocalizedText $Message)
}

function Write-AppMuted {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-AppStyled -Text $Message -Style Muted
}

function Wait-App {
    param([string]$Message = 'Enter para continuar')
    [void](Read-ToolkitInput ('  ' + (Get-ToolkitString Continue)))
}

function Show-AppErrorScreen {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message
    )

    Show-AppHeader -Section $Title
    Write-Host ''
    Write-AppError -Message $Message
    Wait-App
}

function Read-AppYesNo {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [bool]$DefaultYes = $false
    )
    return Read-ToolkitYesNo (Get-LocalizedText $Prompt) $DefaultYes
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
        if ([int]::TryParse($raw, [ref]$number) -and $number -ge $Minimum -and $number -le $Maximum) {
            return $number
        }

        $message = if ($script:Language -eq 'en') {
            "Enter a number between $Minimum and $Maximum."
        }
        else {
            "Introduce un número entre $Minimum y $Maximum."
        }
        Write-AppWarning -Message $message
    }
}

function Show-NumberMenu {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][array]$Items,
        [string[]]$Description = @(),
        [scriptblock]$RenderBody
    )
    return Read-ToolkitLegacyMenu -Items $Items -Mode Indexed -Title $Title -Description $Description -RenderBody $RenderBody
}


# ============================================================
# UTILIDADES DE RUTA
# ============================================================

function Normalize-LocalPath {

    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    try {

        $resolved =
            (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path

    }
    catch {

        return ""
    }

    if ($resolved.Length -gt 3) {
        $resolved = $resolved.TrimEnd('\')
    }

    return $resolved
}


function Test-PathWithin {

    param(
        [string]$ChildPath,
        [string]$ParentPath
    )

    if (
        [string]::IsNullOrWhiteSpace($ChildPath) -or
        [string]::IsNullOrWhiteSpace($ParentPath)
    ) {
        return $false
    }

    if (
        $ChildPath.Equals(
            $ParentPath,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        return $true
    }

    $prefix =
        $ParentPath.TrimEnd('\') + "\"

    return $ChildPath.StartsWith(
        $prefix,
        [System.StringComparison]::OrdinalIgnoreCase
    )
}


# ============================================================
# FILTROS DE CONTENIDO
# ============================================================

function Test-IsExcludedSyncPath {

    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $true
    }

    $segments =
        $Path -split '[\\/]'

    foreach ($segment in $segments) {

        if (
            $segment -ieq "OneDriveCloudTemp" -or
            $segment -ieq "OneDriveTemp"
        ) {
            return $true
        }
    }

    return $false
}


function Test-IsAuxiliaryName {

    param([string]$Name)

    foreach ($excluded in $ExcludedExactNames) {

        if ($Name -ieq $excluded) {
            return $true
        }
    }

    # Temporales/locks de Office
    if ($Name -like '~$*') {
        return $true
    }

    # AppleDouble
    if ($Name -like '._*') {
        return $true
    }

    return $false
}


function Test-IsUserItem {

    param(
        [System.IO.FileSystemInfo]$Item
    )

    if ($null -eq $Item) {
        return $false
    }

    # Atributo System = excluir
    # Hidden sin System = incluir
    try {

        if (
            ($Item.Attributes -band
                [System.IO.FileAttributes]::System) -ne 0
        ) {
            return $false
        }

    }
    catch {
    }

    if (Test-IsAuxiliaryName $Item.Name) {
        return $false
    }

    return $true
}


# ============================================================
# DETECCIÓN DE UBICACIONES SYNC
# ============================================================

function Get-SyncLocations {

    $candidates = @()
    $seen = @{}

    $registryRoot =
        "HKCU:\Software\SyncEngines\Providers\OneDrive"


    # --------------------------------------------------------
    # REGISTRO SYNC ENGINE
    # --------------------------------------------------------

    if (Test-Path -LiteralPath $registryRoot) {

        $entries = @(
            Get-ChildItem `
                -Path $registryRoot `
                -Recurse `
                -ErrorAction SilentlyContinue
        )

        foreach ($entry in $entries) {

            try {

                $properties =
                    Get-ItemProperty `
                        -LiteralPath $entry.PSPath `
                        -ErrorAction Stop

                $mountPoint =
                    [string]$properties.MountPoint

                if ([string]::IsNullOrWhiteSpace($mountPoint)) {
                    continue
                }

                if (
                    -not (
                        Test-Path `
                            -LiteralPath $mountPoint `
                            -PathType Container
                    )
                ) {
                    continue
                }

                $resolved =
                    Normalize-LocalPath $mountPoint

                if ([string]::IsNullOrWhiteSpace($resolved)) {
                    continue
                }

                if (Test-IsExcludedSyncPath $resolved) {
                    continue
                }

                $key =
                    $resolved.ToLowerInvariant()

                if ($seen.ContainsKey($key)) {
                    continue
                }


                # URL
                $url = ""

                if ($null -ne $properties.UrlNamespace) {
                    $url = [string]$properties.UrlNamespace
                }


                # Nombre
                $displayName = ""

                if ($null -ne $properties.DisplayName) {
                    $displayName = [string]$properties.DisplayName
                }

                if ([string]::IsNullOrWhiteSpace($displayName)) {
                    $displayName = Split-Path -Leaf $resolved
                }


                # LibraryType
                $libraryType = ""

                if ($null -ne $properties.LibraryType) {
                    $libraryType = [string]$properties.LibraryType
                }


                # Tipo
                $type = "OneDrive"

                if (
                    $url -match "sharepoint\.com/sites/" -or
                    $url -match "sharepoint\.com/teams/" -or
                    $libraryType -match "team"
                ) {

                    $type = "SharePoint"

                }
                elseif ($url -match "\-my\.sharepoint\.com") {

                    $type = "OneDrive Business"

                }
                elseif ($entry.PSChildName -match "Personal") {

                    $type = "OneDrive Personal"

                }
                elseif ($resolved -match "\\OneDrive\s*-\s*") {

                    $type = "OneDrive Business"
                }


                $candidates += [PSCustomObject]@{
                    Tipo   = $type
                    Nombre = $displayName
                    Ruta   = $resolved
                    URL    = $url
                    Fuente = "SyncEngine"
                }

                $seen[$key] = $true

            }
            catch {
            }
        }
    }


    # --------------------------------------------------------
    # VARIABLES DE ENTORNO COMO FALLBACK
    # --------------------------------------------------------

    $environmentPaths = @(
        $env:OneDrive,
        $env:OneDriveConsumer,
        $env:OneDriveCommercial
    )


    foreach ($path in $environmentPaths) {

        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        if (
            -not (
                Test-Path `
                    -LiteralPath $path `
                    -PathType Container
            )
        ) {
            continue
        }

        $resolved =
            Normalize-LocalPath $path

        if ([string]::IsNullOrWhiteSpace($resolved)) {
            continue
        }

        if (Test-IsExcludedSyncPath $resolved) {
            continue
        }

        $key =
            $resolved.ToLowerInvariant()

        if ($seen.ContainsKey($key)) {
            continue
        }


        $type = "OneDrive"

        if ($resolved -match "\\OneDrive\s*-\s*") {

            $type = "OneDrive Business"

        }
        elseif ((Split-Path -Leaf $resolved) -ieq "OneDrive") {

            $type = "OneDrive Personal"
        }


        $candidates += [PSCustomObject]@{
            Tipo   = $type
            Nombre = (Split-Path -Leaf $resolved)
            Ruta   = $resolved
            URL    = ""
            Fuente = "Environment"
        }

        $seen[$key] = $true
    }


    # --------------------------------------------------------
    # IDENTIFICAR UBICACIONES ANIDADAS
    # --------------------------------------------------------

    $final = @()

    foreach ($candidate in $candidates) {

        $ancestors = @(
            $candidates |
                Where-Object {

                    $_.Ruta -ne $candidate.Ruta -and

                    (
                        Test-PathWithin `
                            -ChildPath $candidate.Ruta `
                            -ParentPath $_.Ruta
                    )
                }
        )


        $isNested = $false
        $physicalRoot = $candidate.Ruta


        if ($ancestors.Count -gt 0) {

            $isNested = $true

            $physicalRoot = (
                $ancestors |
                    Sort-Object { $_.Ruta.Length } |
                    Select-Object -First 1
            ).Ruta
        }


        if ($isNested) {

            $role = "Shortcut / ubicación anidada"

        }
        else {

            $role = "Raíz principal"
        }


        $final += [PSCustomObject]@{
            Tipo       = $candidate.Tipo
            Nombre     = $candidate.Nombre
            Ruta       = $candidate.Ruta
            URL        = $candidate.URL
            Fuente     = $candidate.Fuente

            Rol        = $role
            EsAnidada  = $isNested
            RaizFisica = $physicalRoot
        }
    }


    return @(
        $final |
            Sort-Object `
                @{ Expression = { $_.RaizFisica } },
                @{ Expression = { $_.Ruta.Length } },
                Ruta
    )
}


# ============================================================
# OBTENER RAÍCES MÍNIMAS A RECORRER
# ============================================================

function Get-MinimalScanRoots {

    param(
        [object[]]$Locations
    )

    $ordered = @(
        $Locations |
            Sort-Object { $_.Ruta.Length }
    )

    $roots = @()


    foreach ($location in $ordered) {

        $covered = $false

        foreach ($root in $roots) {

            if (
                Test-PathWithin `
                    -ChildPath $location.Ruta `
                    -ParentPath $root.Ruta
            ) {

                $covered = $true
                break
            }
        }


        if (-not $covered) {

            $roots += [PSCustomObject]@{
                Ruta = $location.Ruta
            }
        }
    }


    return @($roots)
}


# ============================================================
# UBICACIONES LÓGICAS PARA EL ESCANEO
# ============================================================

function Get-AnalysisLocations {

    param(
        [object[]]$SelectedLocations,
        [object[]]$AllLocations,
        [object[]]$ScanRoots
    )

    $result = @()
    $seen = @{}


    foreach ($selected in $SelectedLocations) {

        $key = $selected.Ruta.ToLowerInvariant()

        if (-not $seen.ContainsKey($key)) {

            $result += $selected
            $seen[$key] = $true
        }
    }


    foreach ($root in $ScanRoots) {

        foreach ($location in $AllLocations) {

            if (
                Test-PathWithin `
                    -ChildPath $location.Ruta `
                    -ParentPath $root.Ruta
            ) {

                $key =
                    $location.Ruta.ToLowerInvariant()

                if (-not $seen.ContainsKey($key)) {

                    $result += $location
                    $seen[$key] = $true
                }
            }
        }
    }


    return @($result)
}


function Get-LogicalLocationForPath {

    param(
        [string]$FullPath,
        [object[]]$Locations
    )

    $matches = @(
        $Locations |
            Where-Object {

                Test-PathWithin `
                    -ChildPath $FullPath `
                    -ParentPath $_.Ruta
            }
    )


    if ($matches.Count -eq 0) {
        return $null
    }


    return (
        $matches |
            Sort-Object { $_.Ruta.Length } -Descending |
            Select-Object -First 1
    )
}


# ============================================================
# DIAGNÓSTICO DE LONGITUD
# ============================================================

function Get-PathDiagnostic {

    param(
        [System.IO.FileSystemInfo]$Item,
        [object]$LogicalLocation
    )


    $fullPath =
        $Item.FullName

    $name =
        $Item.Name

    $logicalRoot =
        $LogicalLocation.Ruta


    # --------------------------------------------------------
    # RUTA CLOUD / RELATIVA
    # --------------------------------------------------------

    $prefix =
        $logicalRoot.TrimEnd('\') + "\"


    if (
        $fullPath.StartsWith(
            $prefix,
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {

        $cloudPath =
            $fullPath.Substring($prefix.Length)

    }
    else {

        $cloudPath =
            $fullPath
    }


    # --------------------------------------------------------
    # MEDICIONES
    # --------------------------------------------------------

    $localChars =
        $fullPath.Length

    $cloudChars =
        $cloudPath.Length

    $nameChars =
        $name.Length


    # --------------------------------------------------------
    # ESTADO
    # --------------------------------------------------------

    $criticalReasons =
        New-Object System.Collections.Generic.List[string]

    $warningReasons =
        New-Object System.Collections.Generic.List[string]


    # WINDOWS
    if ($localChars -ge $CRITICAL_WINDOWS) {

        $criticalReasons.Add(
            "Ruta local Windows/Office >= $CRITICAL_WINDOWS"
        )

    }
    elseif ($localChars -ge $WARN_WINDOWS) {

        $warningReasons.Add(
            "Ruta local cerca del limite Windows"
        )
    }


    # CLOUD
    if ($cloudChars -gt $LIMIT_CLOUD) {

        $criticalReasons.Add(
            "Ruta cloud > $LIMIT_CLOUD"
        )

    }
    elseif ($cloudChars -ge $WARN_CLOUD) {

        $warningReasons.Add(
            "Ruta cloud cerca de $LIMIT_CLOUD"
        )
    }


    # ONEDRIVE SYNC
    if ($localChars -gt $LIMIT_SYNC) {

        $criticalReasons.Add(
            "OneDrive Sync > $LIMIT_SYNC"
        )

    }
    elseif ($localChars -ge $WARN_SYNC) {

        $warningReasons.Add(
            "OneDrive Sync cerca de $LIMIT_SYNC"
        )
    }


    # NOMBRE
    if ($nameChars -gt $LIMIT_NAME) {

        $criticalReasons.Add(
            "Nombre individual > $LIMIT_NAME"
        )

    }
    elseif ($nameChars -ge $WARN_NAME) {

        $warningReasons.Add(
            "Nombre individual cerca de $LIMIT_NAME"
        )
    }


    if ($criticalReasons.Count -gt 0) {

        $state = "CRITICO"

        $reasons =
            @($criticalReasons + $warningReasons) -join "; "

    }
    elseif ($warningReasons.Count -gt 0) {

        $state = "AVISO"

        $reasons =
            $warningReasons -join "; "

    }
    else {

        $state = "OK"
        $reasons = ""
    }


    if ($Item.PSIsContainer) {

        $itemType = "Carpeta"

    }
    else {

        $itemType = "Archivo"
    }


    return [PSCustomObject]@{
        Estado           = $state
        Motivo           = $reasons

        TipoSync         = $LogicalLocation.Tipo
        Tipo             = $itemType

        CaracteresLocal  = $localChars
        CaracteresCloud  = $cloudChars
        CaracteresNombre = $nameChars

        Nombre           = $name
        RutaCompleta     = $fullPath
    }
}


# ============================================================
# PROCESAR ELEMENTO
# ============================================================

function Process-UserItem {

    param(
        [System.IO.FileSystemInfo]$Item,
        [object[]]$AnalysisLocations
    )


    $logicalLocation =
        Get-LogicalLocationForPath `
            -FullPath $Item.FullName `
            -Locations $AnalysisLocations


    if ($null -eq $logicalLocation) {
        return
    }


    $script:TotalScanned++


    if (
        $Item.FullName.Length -gt
        $script:LongestLength
    ) {

        $script:LongestLength =
            $Item.FullName.Length

        $script:LongestPath =
            $Item.FullName
    }


    $diagnostic =
        Get-PathDiagnostic `
            -Item $Item `
            -LogicalLocation $logicalLocation


    switch ($diagnostic.Estado) {

        "CRITICO" {

            $script:CriticalCount++
            $script:Results.Add($diagnostic)
        }

        "AVISO" {

            $script:WarningCount++
            $script:Results.Add($diagnostic)
        }
    }


    if (($script:TotalScanned % 250) -eq 0) {

        Write-Progress `
            -Activity "Analizando rutas de OneDrive / SharePoint" `
            -Status "$($script:TotalScanned) rutas analizadas | Avisos: $($script:WarningCount) | Criticas: $($script:CriticalCount)" `
            -CurrentOperation $Item.FullName
    }
}


# ============================================================
# RECORRIDO RECURSIVO
# ============================================================

function Scan-DirectoryRecursive {

    param(
        [string]$DirectoryPath,
        [object[]]$AnalysisLocations
    )


    try {

        $children = @(
            Get-ChildItem `
                -LiteralPath $DirectoryPath `
                -Force `
                -ErrorAction Stop
        )

    }
    catch {

        $script:EnumerationErrors++
        return
    }


    foreach ($item in $children) {

        if (-not (Test-IsUserItem $item)) {
            continue
        }


        Process-UserItem `
            -Item $item `
            -AnalysisLocations $AnalysisLocations


        if ($item.PSIsContainer) {

            Scan-DirectoryRecursive `
                -DirectoryPath $item.FullName `
                -AnalysisLocations $AnalysisLocations
        }
    }
}

function Write-AppThresholdSummary {
    $windowsValue = if ($script:Language -eq 'en') {
        "warning >= $script:WARN_WINDOWS | critical >= $script:CRITICAL_WINDOWS"
    }
    else {
        "aviso >= $script:WARN_WINDOWS | crítico >= $script:CRITICAL_WINDOWS"
    }
    $cloudValue = if ($script:Language -eq 'en') {
        "warning >= $script:WARN_CLOUD | critical > $script:LIMIT_CLOUD"
    }
    else {
        "aviso >= $script:WARN_CLOUD | crítico > $script:LIMIT_CLOUD"
    }
    $syncValue = if ($script:Language -eq 'en') {
        "warning >= $script:WARN_SYNC | critical > $script:LIMIT_SYNC"
    }
    else {
        "aviso >= $script:WARN_SYNC | crítico > $script:LIMIT_SYNC"
    }
    $nameValue = if ($script:Language -eq 'en') {
        "warning >= $script:WARN_NAME | critical > $script:LIMIT_NAME"
    }
    else {
        "aviso >= $script:WARN_NAME | crítico > $script:LIMIT_NAME"
    }

    Write-AppSection -Title 'Umbrales activos'
    Write-AppField -Name 'Windows / Office' -Value $windowsValue -Style Primary
    Write-AppField -Name 'Cloud / SharePoint' -Value $cloudValue -Style Primary
    Write-AppField -Name 'OneDrive Sync' -Value $syncValue -Style Primary
    Write-AppField -Name 'Nombre individual' -Value $nameValue -Style Primary
}

function Write-AppDetectedLocations {
    param([AllowEmptyCollection()][object[]]$Locations)

    Write-AppSection -Title 'Ubicaciones detectadas'
    if ($Locations.Count -eq 0) {
        Write-AppWarning -Message 'No se detectaron ubicaciones automáticamente.'
        Write-AppMuted -Message 'Puede introducir una ruta manual.'
        return
    }

    Write-AppField -Name 'Detectadas' -Value $Locations.Count -Style Success
    Write-Host ''

    for ($i = 0; $i -lt $Locations.Count; $i++) {
        $location = $Locations[$i]
        Write-AppStyled -Text ('  {0,2}  ' -f ($i + 1)) -Style Primary -NoNewline
        Write-AppStyled -Text ([string]$location.Tipo) -Style Normal
        Write-AppField -Name 'Nombre' -Value $location.Nombre -Style Primary
        Write-AppField -Name 'Ruta' -Value $location.Ruta -Style Muted

        if ($location.EsAnidada) {
            Write-AppWarning -Message 'Shortcut / ubicación anidada'
        }
        Write-Host ''
    }
}

function New-AppManualLocation {
    $path = (Read-Host 'Ruta a analizar').Trim().Trim('"')

    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        Write-AppError -Message 'La ruta indicada no existe.'
        return $null
    }

    $resolved = Normalize-LocalPath $path
    if ([string]::IsNullOrWhiteSpace($resolved)) {
        Write-AppError -Message 'La ruta indicada no existe.'
        return $null
    }

    if (Test-IsExcludedSyncPath $resolved) {
        Write-AppError -Message 'Esta ubicación corresponde a contenido auxiliar y no será analizada.'
        return $null
    }

    return [PSCustomObject]@{
        Tipo = 'Manual'
        Nombre = Split-Path -Leaf $resolved
        Ruta = $resolved
        URL = ''
        Fuente = 'Manual'
        Rol = 'Ruta manual'
        EsAnidada = $false
        RaizFisica = $resolved
    }
}

function Read-AppLocationNumbers {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Locations)

    $prompt = if ($script:Language -eq 'en') {
        'Enter comma-separated location numbers'
    }
    else {
        'Introducir números de ubicaciones separados por coma'
    }

    $raw = (Read-Host $prompt).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    $selected = @()
    foreach ($token in ($raw -split ',')) {
        $number = 0
        $value = $token.Trim()

        if (-not [int]::TryParse($value, [ref]$number) -or
            $number -lt 1 -or
            $number -gt $Locations.Count) {
            $message = if ($script:Language -eq 'en') {
                "Invalid selection: $value"
            }
            else {
                "Selección no válida: $value"
            }
            Write-AppError -Message $message
            return $null
        }

        $selected += $Locations[$number - 1]
    }

    return @($selected | Sort-Object Ruta -Unique)
}

function Select-AppLocations {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Locations)

    while ($true) {
        $items = @()
        if ($Locations.Count -gt 0) {
            $items += [PSCustomObject]@{
                Label = 'Seleccionar por número'
                Value = 'Numbers'
                Hint = 'Elige una o varias ubicaciones separadas por coma.'
            }
            $items += [PSCustomObject]@{
                Label = 'Analizar todas'
                Value = 'All'
                Hint = 'Incluye todas las ubicaciones detectadas.'
            }
        }
        $items += [PSCustomObject]@{
            Label = 'Introducir ruta manual'
            Value = 'Manual'
            Hint = 'Analiza una carpeta local que no fue detectada automáticamente.'
        }
        $items += [PSCustomObject]@{
            Label = 'Volver'
            Value = 'Back'
            Hint = 'Regresa al menú anterior.'
        }

        $choice = Show-NumberMenu -Title 'Selección de ubicaciones' -Items $items -Description @(
            'Selecciona las ubicaciones locales que quieres analizar.'
        ) -RenderBody {
            Write-AppDetectedLocations -Locations $Locations
        }

        switch ($choice.Value) {
            'Numbers' {
                $selected = Read-AppLocationNumbers -Locations $Locations
                if ($null -ne $selected -and $selected.Count -gt 0) {
                    return @($selected)
                }
                Write-AppWarning -Message 'La selección no contiene ubicaciones válidas.'
                Start-Sleep -Milliseconds 700
            }
            'All' {
                return @($Locations)
            }
            'Manual' {
                Show-AppHeader -Section 'Ruta manual'
                $manual = New-AppManualLocation
                if ($null -ne $manual) {
                    Write-AppOk -Message 'La ruta fue agregada.'
                    Wait-App
                    return @($manual)
                }
                Start-Sleep -Milliseconds 700
            }
            'Back' {
                return $null
            }
        }
    }
}

function Clear-AppWorkingContext {
    Show-AppHeader -Section 'Limpiar contexto de trabajo'

    Write-AppField -Name 'Ubicaciones' -Value $script:SelectedLocations.Count -Style $(if ($script:SelectedLocations.Count -gt 0) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Último resultado' -Value $(if ($null -ne $script:LastScan) { 'DISPONIBLE' } else { 'NO DISPONIBLE' }) -Style $(if ($null -ne $script:LastScan) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Reporte' -Value $(if ($script:LastReportPath) { $script:LastReportPath } else { 'No generado' }) -Style Muted
    Write-Host ''
    Write-AppMuted -Message 'Limpia las ubicaciones seleccionadas y el resultado del análisis. No elimina archivos ni reportes existentes.'
    Write-Host ''

    if (-not (Read-AppYesNo -Prompt '¿Limpiar la selección actual?' -DefaultYes $false)) {
        return
    }

    $script:SelectedLocations = @()
    $script:LastScan = $null
    $script:LastReportPath = ''
    Write-AppOk -Message 'La selección fue limpiada.'
    Wait-App
}

function Show-AppSettings {
    while ($true) {
        $items = @(
            [PSCustomObject]@{
                Label = 'Windows / Office'
                Value = 'Windows'
                Hint = 'Configura los umbrales de longitud de ruta local.'
            },
            [PSCustomObject]@{
                Label = 'Cloud / SharePoint'
                Value = 'Cloud'
                Hint = 'Configura los umbrales de la ruta relativa en la nube.'
            },
            [PSCustomObject]@{
                Label = 'OneDrive Sync'
                Value = 'Sync'
                Hint = 'Configura los umbrales de sincronización local.'
            },
            [PSCustomObject]@{
                Label = 'Nombre individual'
                Value = 'Name'
                Hint = 'Configura el límite del nombre de archivo o carpeta.'
            },
            [PSCustomObject]@{
                Label = 'Cambiar idioma'
                Value = 'Language'
                Hint = if ($script:Language -eq 'en') { 'English' } else { 'Español' }
            },
            [PSCustomObject]@{
                Label = 'Restablecer umbrales predeterminados'
                Value = 'Reset'
                Hint = 'Restaura los valores recomendados para Windows, Cloud, Sync y nombres.'
            },
            [PSCustomObject]@{
                Label = 'Volver'
                Value = 'Back'
                Hint = 'Regresa al menú anterior.'
            }
        )

        $choice = Show-NumberMenu -Title 'Configuración' -Items $items -Description @(
            'Ajusta los umbrales usados para clasificar las rutas.'
        ) -RenderBody {
            Write-AppThresholdSummary
            Write-AppMuted -Message 'Los cambios se mantienen durante esta ejecución.'
        }

        switch ($choice.Value) {
            'Windows' {
                Show-AppHeader -Section 'Windows / Office'
                $script:WARN_WINDOWS = Read-AppInteger -Prompt 'Aviso desde caracteres' -Minimum 1 -Maximum 10000 -DefaultValue $script:WARN_WINDOWS
                $script:CRITICAL_WINDOWS = Read-AppInteger -Prompt 'Crítico desde caracteres' -Minimum $script:WARN_WINDOWS -Maximum 10000 -DefaultValue ([Math]::Max($script:WARN_WINDOWS, $script:CRITICAL_WINDOWS))
                Write-AppOk -Message 'Umbrales guardados en memoria.'
                Wait-App
            }
            'Cloud' {
                Show-AppHeader -Section 'Cloud / SharePoint'
                $script:WARN_CLOUD = Read-AppInteger -Prompt 'Aviso desde caracteres' -Minimum 1 -Maximum 10000 -DefaultValue $script:WARN_CLOUD
                $script:LIMIT_CLOUD = Read-AppInteger -Prompt 'Crítico sobre caracteres' -Minimum $script:WARN_CLOUD -Maximum 10000 -DefaultValue ([Math]::Max($script:WARN_CLOUD, $script:LIMIT_CLOUD))
                Write-AppOk -Message 'Umbrales guardados en memoria.'
                Wait-App
            }
            'Sync' {
                Show-AppHeader -Section 'OneDrive Sync'
                $script:WARN_SYNC = Read-AppInteger -Prompt 'Aviso desde caracteres' -Minimum 1 -Maximum 10000 -DefaultValue $script:WARN_SYNC
                $script:LIMIT_SYNC = Read-AppInteger -Prompt 'Crítico sobre caracteres' -Minimum $script:WARN_SYNC -Maximum 10000 -DefaultValue ([Math]::Max($script:WARN_SYNC, $script:LIMIT_SYNC))
                Write-AppOk -Message 'Umbrales guardados en memoria.'
                Wait-App
            }
            'Name' {
                Show-AppHeader -Section 'Nombre individual'
                $script:WARN_NAME = Read-AppInteger -Prompt 'Aviso desde caracteres' -Minimum 1 -Maximum 10000 -DefaultValue $script:WARN_NAME
                $script:LIMIT_NAME = Read-AppInteger -Prompt 'Crítico sobre caracteres' -Minimum $script:WARN_NAME -Maximum 10000 -DefaultValue ([Math]::Max($script:WARN_NAME, $script:LIMIT_NAME))
                Write-AppOk -Message 'Umbrales guardados en memoria.'
                Wait-App
            }
            'Language' {
                Initialize-AppLanguage
            }
            'Reset' {
                Show-AppHeader -Section 'Restablecer umbrales'
                Write-AppWarning -Message 'Se restaurarán los umbrales recomendados.'
                if (Read-AppYesNo -Prompt '¿Restablecer los umbrales?' -DefaultYes $false) {
                    $script:WARN_WINDOWS = 240
                    $script:CRITICAL_WINDOWS = 256
                    $script:WARN_CLOUD = 360
                    $script:LIMIT_CLOUD = 400
                    $script:WARN_SYNC = 480
                    $script:LIMIT_SYNC = 520
                    $script:WARN_NAME = 240
                    $script:LIMIT_NAME = 255
                    Write-AppOk -Message 'Umbrales restablecidos.'
                    Wait-App
                }
            }
            'Back' {
                return
            }
        }
    }
}

function Open-AppReports {
    $folder = [Environment]::GetFolderPath('Desktop')
    if ([string]::IsNullOrWhiteSpace($folder)) {
        $folder = (Get-Location).Path
    }

    if ($script:LastReportPath -and (Test-Path -LiteralPath $script:LastReportPath)) {
        $folder = Split-Path -Parent $script:LastReportPath
    }

    try {
        Start-Process -FilePath 'explorer.exe' -ArgumentList $folder
        Write-AppInfo -Message $folder
    }
    catch {
        Show-AppErrorScreen -Title 'Reportes' -Message 'No se pudo abrir la carpeta de reportes.'
    }
}

function Show-AppDiagnostics {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Locations)

    $items = @(
        [PSCustomObject]@{
            Label = 'Actualizar ubicaciones detectadas'
            Value = 'Refresh'
            Hint = 'Vuelve a consultar las ubicaciones sincronizadas del equipo.'
        },
        [PSCustomObject]@{
            Label = 'Abrir carpeta de reportes'
            Value = 'Reports'
            Hint = 'Abre el escritorio o la carpeta del último CSV generado.'
        },
        [PSCustomObject]@{
            Label = 'Volver'
            Value = 'Back'
            Hint = 'Regresa al menú anterior.'
        }
    )

    $choice = Show-NumberMenu -Title 'Diagnóstico' -Items $items -Description @(
        'Diagnóstico del sistema y del analizador.'
    ) -RenderBody {
        Write-AppField -Name 'Sistema operativo' -Value $env:OS -Style Muted
        Write-AppField -Name 'PowerShell' -Value $PSVersionTable.PSVersion -Style Success
        Write-AppField -Name 'Ubicaciones detectadas' -Value $Locations.Count -Style Primary
        Write-AppField -Name 'Ubicaciones seleccionadas' -Value $script:SelectedLocations.Count -Style Primary
        Write-AppField -Name 'Último análisis' -Value $(if ($null -ne $script:LastScan) { 'DISPONIBLE' } else { 'NO DISPONIBLE' }) -Style $(if ($null -ne $script:LastScan) { 'Success' } else { 'Muted' })
        Write-AppField -Name 'Reporte' -Value $(if ($script:LastReportPath) { $script:LastReportPath } else { 'No generado' }) -Style Muted
        Write-AppMuted -Message 'El analizador es local y no requiere tenant, aplicación ni autenticación.'
    }

    switch ($choice.Value) {
        'Refresh' { return }
        'Reports' { Open-AppReports }
        'Back' { return }
    }
}

function Show-AppContext {
    param([Parameter(Mandatory = $true)][int]$DetectedCount)

    Write-AppSection -Title 'Contexto'
    Write-AppField -Name 'Ubicaciones detectadas' -Value $DetectedCount -Style Primary
    Write-AppField -Name 'Ubicaciones seleccionadas' -Value $script:SelectedLocations.Count -Style $(if ($script:SelectedLocations.Count -gt 0) { 'Success' } else { 'Warning' })
    Write-AppField -Name 'Análisis' -Value $(if ($null -ne $script:LastScan) { 'DISPONIBLE' } else { 'NO DISPONIBLE' }) -Style $(if ($null -ne $script:LastScan) { 'Success' } else { 'Muted' })
    Write-AppField -Name 'Reporte' -Value $(if ($script:LastReportPath) { $script:LastReportPath } else { 'No generado' }) -Style Muted
    Write-AppField -Name 'Modo' -Value 'LECTURA LOCAL' -Style Primary
    Write-Host ''
    Write-AppMuted -Message 'No modifica archivos; solo analiza longitudes y exporta hallazgos.'
}

function Export-AppScanReport {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Results)

    if ($Results.Count -eq 0) {
        return ''
    }

    $desktop = [Environment]::GetFolderPath('Desktop')
    if ([string]::IsNullOrWhiteSpace($desktop)) {
        $desktop = (Get-Location).Path
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outputFile = Join-Path -Path $desktop -ChildPath "OneDrive_Path_Issues_$timestamp.csv"

    if ($script:Language -eq 'en') {
        $Results |
            Select-Object `
                @{Name='Status';Expression={Get-LocalizedText $_.Estado}},
                @{Name='Reason';Expression={Get-LocalizedText $_.Motivo}},
                @{Name='SyncType';Expression={Get-LocalizedText $_.TipoSync}},
                @{Name='ItemType';Expression={Get-LocalizedText $_.Tipo}},
                @{Name='LocalPathLength';Expression={$_.CaracteresLocal}},
                @{Name='CloudPathLength';Expression={$_.CaracteresCloud}},
                @{Name='NameLength';Expression={$_.CaracteresNombre}},
                @{Name='Name';Expression={$_.Nombre}},
                @{Name='FullPath';Expression={$_.RutaCompleta}} |
            Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding UTF8 -UseCulture
    }
    else {
        $Results |
            Select-Object Estado,Motivo,
                @{Name='TipoSincronizacion';Expression={$_.TipoSync}},
                @{Name='TipoElemento';Expression={$_.Tipo}},
                @{Name='LongitudRutaLocal';Expression={$_.CaracteresLocal}},
                @{Name='LongitudRutaCloud';Expression={$_.CaracteresCloud}},
                @{Name='LongitudNombre';Expression={$_.CaracteresNombre}},
                Nombre,RutaCompleta |
            Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding UTF8 -UseCulture
    }

    return $outputFile
}

function Show-AppScanResult {
    param([Parameter(Mandatory = $true)][object]$Scan)

    Show-AppHeader -Section 'Resultado'
    Write-AppField -Name 'Rutas analizadas' -Value $Scan.TotalScanned -Style Primary
    Write-AppField -Name 'Rutas en AVISO' -Value $Scan.WarningCount -Style $(if ($Scan.WarningCount -gt 0) { 'Warning' } else { 'Success' })
    Write-AppField -Name 'Rutas CRÍTICAS' -Value $Scan.CriticalCount -Style $(if ($Scan.CriticalCount -gt 0) { 'Danger' } else { 'Success' })
    Write-AppField -Name 'Errores de lectura' -Value $Scan.EnumerationErrors -Style $(if ($Scan.EnumerationErrors -gt 0) { 'Warning' } else { 'Success' })

    if ($Scan.LongestLength -gt 0) {
        Write-AppField -Name 'Ruta más larga' -Value "$($Scan.LongestLength) caracteres" -Style Primary
        Write-AppField -Name 'Ubicación' -Value $Scan.LongestPath -Style Muted
    }

    Write-Host ''
    if ($Scan.CriticalCount -gt 0) {
        Write-AppError -Message 'Se encontraron rutas que deberían corregirse.'
    }
    elseif ($Scan.WarningCount -gt 0) {
        Write-AppWarning -Message 'Se encontraron rutas cercanas a los límites.'
    }
    else {
        Write-AppOk -Message 'No se encontraron rutas cercanas o superiores a los límites.'
    }

    if ($Scan.Results.Count -gt 0) {
        Write-AppSection -Title 'Rutas que requieren atención'
        $rows = @($Scan.Results | Select-Object -First 20 Estado,CaracteresLocal,CaracteresCloud,CaracteresNombre,Motivo,RutaCompleta)
        if ($script:Language -eq 'en') {
            $rows | Format-Table `
                @{Label='Status';Expression={Get-LocalizedText $_.Estado}},
                @{Label='Local';Expression={$_.CaracteresLocal}},
                @{Label='Cloud';Expression={$_.CaracteresCloud}},
                @{Label='Name';Expression={$_.CaracteresNombre}},
                @{Label='Reason';Expression={Get-LocalizedText $_.Motivo}},
                @{Label='Full path';Expression={$_.RutaCompleta}} -Wrap -AutoSize
        }
        else {
            $rows | Format-Table -Wrap -AutoSize
        }
    }

    if ($Scan.ReportPath) {
        Write-Host ''
        Write-AppInfo -Message 'CSV generado:'
        Write-AppStyled -Text $Scan.ReportPath -Style Primary
    }
    else {
        Write-AppMuted -Message 'No se generó CSV porque no existen rutas con AVISO o CRITICO.'
    }

    if ($Scan.EnumerationErrors -gt 0) {
        Write-AppWarning -Message 'Algunas carpetas no pudieron enumerarse.'
        Write-AppMuted -Message 'El resultado podría no cubrir el 100 % del contenido.'
    }

    Wait-App
}

function Invoke-AppScan {
    param(
        [Parameter(Mandatory = $true)][object[]]$SelectedLocations,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$AllLocations
    )

    $scanRoots = @(Get-MinimalScanRoots $SelectedLocations)
    $analysisLocations = @(Get-AnalysisLocations -SelectedLocations $SelectedLocations -AllLocations $AllLocations -ScanRoots $scanRoots)

    Show-AppHeader -Section 'Listo para analizar'
    Write-AppField -Name 'Ubicaciones seleccionadas' -Value $SelectedLocations.Count -Style Primary
    Write-AppField -Name 'Raíces físicas a recorrer' -Value $scanRoots.Count -Style Primary
    Write-AppThresholdSummary
    Write-Host ''

    if (-not (Read-AppYesNo -Prompt '¿Iniciar el análisis?' -DefaultYes $true)) {
        return $null
    }

    $script:TotalScanned = 0
    $script:WarningCount = 0
    $script:CriticalCount = 0
    $script:EnumerationErrors = 0
    $script:LongestLength = 0
    $script:LongestPath = ''
    $script:Results = New-Object System.Collections.Generic.List[object]

    Show-AppHeader -Section 'Analizando'
    for ($i = 0; $i -lt $scanRoots.Count; $i++) {
        Write-AppInfo -Message "[$($i + 1)/$($scanRoots.Count)] $($scanRoots[$i].Ruta)"
        Scan-DirectoryRecursive -DirectoryPath $scanRoots[$i].Ruta -AnalysisLocations $analysisLocations
    }

    Write-Progress -Activity 'Analizando rutas de OneDrive / SharePoint' -Completed

    $sortedResults = @(
        $script:Results |
            Sort-Object `
                @{Expression={if ($_.Estado -eq 'CRITICO') { 0 } else { 1 }}},
                @{Expression='CaracteresLocal';Descending=$true}
    )
    $reportPath = Export-AppScanReport -Results $sortedResults

    $scan = [PSCustomObject]@{
        TotalScanned = $script:TotalScanned
        WarningCount = $script:WarningCount
        CriticalCount = $script:CriticalCount
        EnumerationErrors = $script:EnumerationErrors
        LongestLength = $script:LongestLength
        LongestPath = $script:LongestPath
        Results = @($sortedResults)
        ReportPath = $reportPath
    }

    $script:LastScan = $scan
    $script:LastReportPath = $reportPath
    Show-AppScanResult -Scan $scan
    return $scan
}

function Show-AppMainMenu {
    while ($true) {
        $locations = @(Get-SyncLocations)
        $lastResultHint = if ($null -ne $script:LastScan) { 'Disponible para consultar.' } else { 'Todavía no existe un resultado.' }

        $items = @(
            [PSCustomObject]@{
                Label = 'Seleccionar ubicaciones'
                Value = 'Select'
                Hint = 'Elige ubicaciones detectadas o introduce una ruta manual.'
            },
            [PSCustomObject]@{
                Label = 'Analizar ubicaciones seleccionadas'
                Value = 'Analyze'
                Hint = 'Escanea el contenido local y reporta rutas AVISO o CRÍTICO.'
            },
            [PSCustomObject]@{
                Label = 'Limpiar selección / contexto'
                Value = 'Clear'
                Hint = 'Quita la selección y el resultado en memoria; no borra archivos ni CSV.'
            },
            [PSCustomObject]@{
                Label = 'Ver último resultado'
                Value = 'Result'
                Hint = $lastResultHint
            },
            [PSCustomObject]@{
                Label = 'Configuración'
                Value = 'Settings'
                Hint = 'Ajusta los umbrales de Windows, Cloud, Sync y nombres.'
            },
            [PSCustomObject]@{
                Label = 'Diagnóstico'
                Value = 'Diagnostics'
                Hint = 'Muestra información del sistema y actualiza detecciones.'
            },
            [PSCustomObject]@{
                Label = 'Abrir carpeta de reportes'
                Value = 'Reports'
                Hint = 'Abre el escritorio o la carpeta del último CSV.'
            },
            [PSCustomObject]@{
                Label = 'Salir'
                Value = 'Exit'
                Hint = 'Cierra la herramienta.'
            }
        )

        $choice = Show-NumberMenu -Title 'Inicio' -Items $items -Description @(
            'Diagnóstico de rutas largas para troubleshooting de sincronización.',
            'Analiza contenido local de OneDrive y SharePoint sincronizado.'
        ) -RenderBody {
            Show-AppContext -DetectedCount $locations.Count
        }

        switch ($choice.Value) {
            'Select' {
                $selection = Select-AppLocations -Locations $locations
                if ($null -ne $selection) {
                    $script:SelectedLocations = @($selection)
                }
            }
            'Analyze' {
                if ($script:SelectedLocations.Count -eq 0) {
                    Write-AppWarning -Message 'Selecciona ubicaciones antes de iniciar el análisis.'
                    Wait-App
                    continue
                }
                [void](Invoke-AppScan -SelectedLocations $script:SelectedLocations -AllLocations $locations)
            }
            'Clear' {
                Clear-AppWorkingContext
            }
            'Result' {
                if ($null -eq $script:LastScan) {
                    Write-AppWarning -Message 'El análisis no ha comenzado.'
                    Wait-App
                }
                else {
                    Show-AppScanResult -Scan $script:LastScan
                }
            }
            'Settings' {
                Show-AppSettings
            }
            'Diagnostics' {
                Show-AppDiagnostics -Locations $locations
            }
            'Reports' {
                Open-AppReports
            }
            'Exit' {
                return
            }
        }
    }
}

if ($false) {
# ============================================================
# INICIO
# ============================================================

# ============================================================
# INICIO
# ============================================================

Write-Banner

Write-Host " Buscando ubicaciones sincronizadas..." -ForegroundColor Yellow
Write-Host ""

$syncLocations =
    @(Get-SyncLocations)

$detectedCount =
    $syncLocations.Count


# ============================================================
# MOSTRAR UBICACIONES
# ============================================================

if ($detectedCount -eq 0) {

    Write-Host " No se detectaron ubicaciones automáticamente." -ForegroundColor Yellow
    Write-Host " Puede introducir una ruta manual." -ForegroundColor Gray
    Write-Host ""

}
else {

    Write-StatusLine `
        "Ubicaciones detectadas" `
        $detectedCount `
        Green

    Write-Host ""


    for ($i = 0; $i -lt $detectedCount; $i++) {

        $location =
            $syncLocations[$i]

        Write-Host " [$($i + 1)] " -NoNewline -ForegroundColor Cyan
        Write-Host "$($location.Tipo)" -ForegroundColor White

        Write-Host "     $($location.Nombre)"
        Write-Host "     $($location.Ruta)" -ForegroundColor DarkGray

        if ($location.EsAnidada) {

            Write-Host "     Shortcut / ubicación anidada" -ForegroundColor DarkYellow
        }

        Write-Host ""
    }
}


# ============================================================
# MENÚ
# ============================================================

Write-Section "SELECCIONAR UBICACIONES"

Write-Host "  Número       Ejemplo: 1"
Write-Host "  Varias       Ejemplo: 1,2,4"
Write-Host "  A            Analizar todas"
Write-Host "  M            Introducir ruta manual"
Write-Host ""

$selection =
    Read-Host " Selección"

$selectedLocations = @()


# ============================================================
# PROCESAR SELECCIÓN
# ============================================================

if ($selection -match "^[Aa]$") {

    if ($detectedCount -eq 0) {

        Write-Host ""
        Write-Host " No existen ubicaciones detectadas." -ForegroundColor Red
        exit 1
    }

    $selectedLocations =
        $syncLocations

}
elseif ($selection -match "^[Mm]$") {

    Write-Host ""

    $manualPath =
        Read-Host " Ruta a analizar"

    $manualPath =
        $manualPath.Trim().Trim('"')


    if (
        -not (
            Test-Path `
                -LiteralPath $manualPath `
                -PathType Container
        )
    ) {

        Write-Host ""
        Write-Host " La ruta indicada no existe." -ForegroundColor Red
        exit 1
    }


    $resolvedManual =
        Normalize-LocalPath $manualPath


    if (Test-IsExcludedSyncPath $resolvedManual) {

        Write-Host ""
        Write-Host " Esta ubicación corresponde a contenido auxiliar y no será analizada." -ForegroundColor Red
        exit 1
    }


    $selectedLocations = @(
        [PSCustomObject]@{
            Tipo       = "Manual"
            Nombre     = (Split-Path -Leaf $resolvedManual)
            Ruta       = $resolvedManual
            URL        = ""
            Fuente     = "Manual"
            Rol        = "Ruta manual"
            EsAnidada  = $false
            RaizFisica = $resolvedManual
        }
    )

}
else {

    if ($detectedCount -eq 0) {

        Write-Host ""
        Write-Host " Use M para introducir una ruta manual." -ForegroundColor Red
        exit 1
    }


    foreach ($token in ($selection -split ",")) {

        $token =
            $token.Trim()

        $index = 0


        if (
            -not [int]::TryParse(
                $token,
                [ref]$index
            )
        ) {

            Write-Host ""
            Write-Host " Selección no válida: $token" -ForegroundColor Red
            exit 1
        }


        if (
            $index -lt 1 -or
            $index -gt $detectedCount
        ) {

            Write-Host ""
            Write-Host " Número fuera de rango: $index" -ForegroundColor Red
            exit 1
        }


        $selectedLocations +=
            $syncLocations[$index - 1]
    }


    $selectedLocations = @(
        $selectedLocations |
            Sort-Object Ruta -Unique
    )
}


# ============================================================
# PREPARAR ESCANEO
# ============================================================

$scanRoots =
    @(Get-MinimalScanRoots $selectedLocations)


$analysisLocations =
    @(
        Get-AnalysisLocations `
            -SelectedLocations $selectedLocations `
            -AllLocations $syncLocations `
            -ScanRoots $scanRoots
    )


Write-Section "LISTO PARA ANALIZAR"

Write-StatusLine `
    "Ubicaciones seleccionadas" `
    $selectedLocations.Count `
    Cyan

Write-StatusLine `
    "Raices físicas a recorrer" `
    $scanRoots.Count `
    Cyan


Write-Host ""
Write-Host " Umbrales activos:" -ForegroundColor White
Write-Host "   Windows : aviso desde 240 | crítico desde 256"
Write-Host "   Cloud   : aviso desde 360 | crítico sobre 400"
Write-Host "   Nombre  : aviso desde 240 | crítico sobre 255"
Write-Host ""


# ============================================================
# CONTADORES
# ============================================================

$TotalScanned      = 0
$WarningCount      = 0
$CriticalCount     = 0
$EnumerationErrors = 0

$LongestLength     = 0
$LongestPath       = ""

$Results =
    New-Object System.Collections.Generic.List[object]


# ============================================================
# ESCANEO
# ============================================================

Write-Section "ANALIZANDO"

$rootNumber = 0

foreach ($root in $scanRoots) {

    $rootNumber++

    Write-Host ""
    Write-Host " [$rootNumber/$($scanRoots.Count)] " -NoNewline -ForegroundColor Cyan
    Write-Host $root.Ruta


    Scan-DirectoryRecursive `
        -DirectoryPath $root.Ruta `
        -AnalysisLocations $analysisLocations
}


Write-Progress `
    -Activity "Analizando rutas de OneDrive / SharePoint" `
    -Completed


# ============================================================
# ORDENAR HALLAZGOS
# ============================================================

$sortedResults = @(
    $Results |
        Sort-Object `
            @{
                Expression = {
                    if ($_.Estado -eq "CRITICO") {
                        0
                    }
                    else {
                        1
                    }
                }
            },
            @{
                Expression = "CaracteresLocal"
                Descending = $true
            }
)


$riskCount =
    $WarningCount + $CriticalCount


# ============================================================
# RESULTADO
# ============================================================

Write-Section "RESULTADO"

Write-StatusLine `
    "Rutas de usuario analizadas" `
    $TotalScanned `
    White

Write-StatusLine `
    "Rutas en AVISO" `
    $WarningCount `
    Yellow

Write-StatusLine `
    "Rutas CRITICAS" `
    $CriticalCount `
    Red


if ($EnumerationErrors -gt 0) {

    Write-StatusLine `
        "Errores de lectura" `
        $EnumerationErrors `
        Yellow
}


if ($LongestLength -gt 0) {

    Write-Host ""
    Write-Host " Ruta de usuario más larga:" -ForegroundColor Cyan
    Write-Host "   $LongestLength caracteres"
    Write-Host "   $LongestPath" -ForegroundColor Gray
}


# ============================================================
# ESTADO GENERAL
# ============================================================

Write-Host ""

if ($CriticalCount -gt 0) {

    Write-Host " RESULTADO: CRITICO" -ForegroundColor Red
    Write-Host " Se encontraron rutas que deberían corregirse." -ForegroundColor Red

}
elseif ($WarningCount -gt 0) {

    Write-Host " RESULTADO: REQUIERE REVISION" -ForegroundColor Yellow
    Write-Host " Se encontraron rutas cercanas a los límites." -ForegroundColor Yellow

}
else {

    Write-Host " RESULTADO: OK" -ForegroundColor Green
    Write-Host " No se encontraron rutas cercanas o superiores a los límites." -ForegroundColor Green
}


# ============================================================
# MOSTRAR HALLAZGOS
# ============================================================

if ($riskCount -gt 0) {

    Write-Section "RUTAS QUE REQUIEREN ATENCION"

    $sortedResults |
        Select-Object `
            -First 20 `
            Estado,
            CaracteresLocal,
            CaracteresCloud,
            CaracteresNombre,
            Motivo,
            RutaCompleta |
        Format-Table -Wrap -AutoSize
}


# ============================================================
# CSV
# ============================================================

if ($riskCount -gt 0) {

    $desktop =
        [Environment]::GetFolderPath("Desktop")

    if ([string]::IsNullOrWhiteSpace($desktop)) {
        $desktop = (Get-Location).Path
    }


    $timestamp =
        Get-Date -Format "yyyyMMdd_HHmmss"


    $outputFile =
        Join-Path `
            -Path $desktop `
            -ChildPath "OneDrive_Path_Issues_$timestamp.csv"


    # --------------------------------------------------------
    # Encabezados orientados al cliente
    # --------------------------------------------------------

    $sortedResults |
        Select-Object `
            Estado,
            Motivo,
            @{Name="TipoSincronizacion"; Expression={$_.TipoSync}},
            @{Name="TipoElemento"; Expression={$_.Tipo}},
            @{Name="LongitudRutaLocal"; Expression={$_.CaracteresLocal}},
            @{Name="LongitudRutaCloud"; Expression={$_.CaracteresCloud}},
            @{Name="LongitudNombre"; Expression={$_.CaracteresNombre}},
            Nombre,
            RutaCompleta |
        Export-Csv `
            -LiteralPath $outputFile `
            -NoTypeInformation `
            -Encoding UTF8 `
            -UseCulture


    Write-Host ""
    Write-Host " CSV generado:" -ForegroundColor Cyan
    Write-Host " $outputFile" -ForegroundColor Green

}
else {

    Write-Host ""
    Write-Host " No se generó CSV porque no existen rutas con AVISO o CRITICO." -ForegroundColor DarkGray
}


# ============================================================
# ADVERTENCIA DE COBERTURA
# ============================================================

if ($EnumerationErrors -gt 0) {

    Write-Host ""
    Write-Host " ADVERTENCIA:" -ForegroundColor Yellow
    Write-Host " Algunas carpetas no pudieron enumerarse."
    Write-Host " El resultado podría no cubrir el 100 % del contenido."
}


Write-Host ""
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host " Análisis completado." -ForegroundColor White
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host ""
}

try {
    Initialize-AppTerminal
    Initialize-AppLanguage

    try {
        $Host.UI.RawUI.WindowTitle = 'OneDrive / SharePoint Path Analyzer'
    }
    catch {
    }

    $script:SelectedLocations = @()
    $script:LastScan = $null
    $script:LastReportPath = ''
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

Show-AppHeader -Section 'Finalizado'
Write-AppMuted -Message 'Análisis finalizado.'
