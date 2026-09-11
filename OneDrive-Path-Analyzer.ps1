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

Clear-Host


# ============================================================
# CONFIGURACIÓN
# ============================================================

# Windows / Office
$WARN_WINDOWS      = 240
$CRITICAL_WINDOWS  = 256

# OneDrive / SharePoint Cloud
$WARN_CLOUD        = 360
$LIMIT_CLOUD       = 400

# OneDrive Sync
$WARN_SYNC         = 480
$LIMIT_SYNC        = 520

# Nombre individual
$WARN_NAME         = 240
$LIMIT_NAME        = 255


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

    Clear-Host

    Write-Host ""
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host "       ONEDRIVE / SHAREPOINT PATH ANALYZER - WINDOWS" -ForegroundColor White
    Write-Host "====================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " Diagnóstico de rutas largas para troubleshooting de sincronización."
    Write-Host ""
}


function Write-Section {

    param([string]$Title)

    Write-Host ""
    Write-Host "--------------------------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------------------------" -ForegroundColor DarkCyan
}


function Write-StatusLine {

    param(
        [string]$Label,
        [string]$Value,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    Write-Host (" {0,-30}: " -f $Label) -NoNewline
    Write-Host $Value -ForegroundColor $Color
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