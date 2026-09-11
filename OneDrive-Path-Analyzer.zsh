#!/bin/zsh

#
# OneDrive / SharePoint Path Analyzer - macOS
#
# Herramienta de troubleshooting para localizar rutas
# de usuario cercanas o superiores a los límites relevantes.
#
# Incluye:
#   - Archivos normales
#   - Carpetas normales
#   - Archivos ocultos normales del usuario
#
# Excluye:
#   - Metadatos conocidos de macOS
#   - Temporales de Office
#   - OneDriveCloudTemp
#   - OneDriveTemp
#
# No genera CSV si no existen hallazgos.
#


# ============================================================
# CONFIGURACIÓN
# ============================================================

WARN_CLOUD=360
LIMIT_CLOUD=400

WARN_SYNC=480
LIMIT_SYNC=520

WARN_NAME=240
LIMIT_NAME=255

WARN_FINDER=900
LIMIT_FINDER=1024

WARN_OFFICE_BYTES=900
LIMIT_OFFICE_BYTES=1024


# ============================================================
# COLORES
# ============================================================

if [[ -t 1 ]]; then

    C_RESET=$'\033[0m'
    C_CYAN=$'\033[36m'
    C_WHITE=$'\033[97m'
    C_GREEN=$'\033[32m'
    C_YELLOW=$'\033[33m'
    C_RED=$'\033[31m'
    C_GRAY=$'\033[90m'

else

    C_RESET=""
    C_CYAN=""
    C_WHITE=""
    C_GREEN=""
    C_YELLOW=""
    C_RED=""
    C_GRAY=""

fi


# ============================================================
# INTERFAZ
# ============================================================

banner() {

    clear

    echo ""
    echo "${C_CYAN}====================================================================${C_RESET}"
    echo "${C_WHITE}        ONEDRIVE / SHAREPOINT PATH ANALYZER - macOS${C_RESET}"
    echo "${C_CYAN}====================================================================${C_RESET}"
    echo ""
    echo " Diagnóstico de rutas largas para troubleshooting de sincronización."
    echo ""

}


section() {

    echo ""
    echo "${C_CYAN}--------------------------------------------------------------------${C_RESET}"
    echo "${C_CYAN} $1${C_RESET}"
    echo "${C_CYAN}--------------------------------------------------------------------${C_RESET}"

}


status_line() {

    local label="$1"
    local value="$2"
    local color="$3"

    printf " %-30s: %s%s%s\n" \
        "$label" \
        "$color" \
        "$value" \
        "$C_RESET"

}


# ============================================================
# UTILIDADES
# ============================================================

canonical_path() {

    local target="$1"

    if [[ -d "$target" ]]; then

        (
            cd "$target" 2>/dev/null &&
            pwd -P
        )

    fi

}


path_is_within() {

    local child="$1"
    local parent="$2"

    if [[ "$child" == "$parent" ]]; then
        return 0
    fi

    if [[ "$child" == "$parent/"* ]]; then
        return 0
    fi

    return 1

}


byte_length() {

    local value="$1"
    local LC_ALL=C

    print -r -- "${#value}"

}


csv_escape() {

    local value="$1"

    value="${value//\"/\"\"}"

    printf '"%s"' "$value"

}


# ============================================================
# FILTRO DE CONTENIDO AUXILIAR
# ============================================================

is_auxiliary_name() {

    local name="$1"

    case "$name" in

        .DS_Store)
            return 0
            ;;

        ._*)
            return 0
            ;;

        .Spotlight-V100)
            return 0
            ;;

        .Trashes)
            return 0
            ;;

        .fseventsd)
            return 0
            ;;

        .DocumentRevisions-V100)
            return 0
            ;;

        .TemporaryItems)
            return 0
            ;;

        .VolumeIcon.icns)
            return 0
            ;;

        OneDriveCloudTemp)
            return 0
            ;;

        OneDriveTemp)
            return 0
            ;;

        '~$'*)
            return 0
            ;;

    esac

    return 1

}


is_auxiliary_path() {

    local path="$1"

    local -a components
    components=("${(@s:/:)path}")

    local component

    for component in "${components[@]}"; do

        if is_auxiliary_name "$component"; then
            return 0
        fi

    done

    return 1

}


# ============================================================
# PROBLEMAS
# ============================================================

append_critical() {

    local reason="$1"

    if [[ -z "$criticalReasons" ]]; then

        criticalReasons="$reason"

    else

        criticalReasons="$criticalReasons; $reason"

    fi

}


append_warning() {

    local reason="$1"

    if [[ -z "$warningReasons" ]]; then

        warningReasons="$reason"

    else

        warningReasons="$warningReasons; $reason"

    fi

}


# ============================================================
# UBICACIONES SYNC
# ============================================================

typeset -a syncPaths
typeset -a syncTypes
typeset -a syncNames

syncPaths=()
syncTypes=()
syncNames=()


add_sync_location() {

    local path="$1"
    local type="$2"

    local resolved

    resolved="$(canonical_path "$path")"

    if [[ -z "$resolved" ]]; then
        return
    fi


    if is_auxiliary_path "$resolved"; then
        return
    fi


    local existing

    for existing in "${syncPaths[@]}"; do

        if [[ "$existing" == "$resolved" ]]; then
            return
        fi

    done


    syncPaths+=("$resolved")
    syncTypes+=("$type")
    syncNames+=("${resolved:t}")

}


# ============================================================
# DETECCIÓN MODERNA - FILE PROVIDER
# ============================================================

cloudStorage="$HOME/Library/CloudStorage"

if [[ -d "$cloudStorage" ]]; then

    for path in "$cloudStorage"/OneDrive*(N/); do

        if is_auxiliary_path "$path"; then
            continue
        fi


        baseName="${path:t}"
        syncType="OneDrive"


        if [[ "$baseName" == *"Personal"* ]]; then

            syncType="OneDrive Personal"

        elif [[ "$baseName" == *"Shared Libraries"* ]]; then

            syncType="SharePoint"

        elif [[ "$baseName" == OneDrive-* ]]; then

            syncType="OneDrive Business / Organization"

        fi


        add_sync_location \
            "$path" \
            "$syncType"

    done

fi


# ============================================================
# FALLBACK LEGACY
# ============================================================

for path in "$HOME"/OneDrive*(N/); do

    if is_auxiliary_path "$path"; then
        continue
    fi

    add_sync_location \
        "$path" \
        "OneDrive Legacy"

done


# ============================================================
# INICIO
# ============================================================

banner

echo "${C_YELLOW} Buscando ubicaciones sincronizadas...${C_RESET}"
echo ""

detectedCount=${#syncPaths[@]}


if (( detectedCount == 0 )); then

    echo "${C_YELLOW} No se detectaron ubicaciones automáticamente.${C_RESET}"
    echo " Puede introducir una ruta manual."
    echo ""

else

    status_line \
        "Ubicaciones detectadas" \
        "$detectedCount" \
        "$C_GREEN"

    echo ""


    for (( i=1; i<=detectedCount; i++ )); do

        echo " ${C_CYAN}[$i]${C_RESET} ${C_WHITE}${syncTypes[$i]}${C_RESET}"
        echo "     ${syncNames[$i]}"
        echo "     ${C_GRAY}${syncPaths[$i]}${C_RESET}"
        echo ""

    done

fi


# ============================================================
# MENÚ
# ============================================================

section "SELECCIONAR UBICACIONES"

echo "  Número       Ejemplo: 1"
echo "  Varias       Ejemplo: 1,2,4"
echo "  A            Analizar todas"
echo "  M            Introducir ruta manual"
echo ""

printf " Selección: "
IFS= read -r selection


# ============================================================
# SELECCIÓN
# ============================================================

typeset -a selectedPaths
typeset -a selectedTypes

selectedPaths=()
selectedTypes=()


if [[ "$selection" == [Aa] ]]; then

    if (( detectedCount == 0 )); then

        echo ""
        echo "${C_RED} No existen ubicaciones detectadas.${C_RESET}"
        exit 1

    fi


    selectedPaths=("${syncPaths[@]}")
    selectedTypes=("${syncTypes[@]}")


elif [[ "$selection" == [Mm] ]]; then

    echo ""

    printf " Ruta a analizar: "
    IFS= read -r manualPath


    manualPath="${manualPath#\"}"
    manualPath="${manualPath%\"}"
    manualPath="${manualPath#\'}"
    manualPath="${manualPath%\'}"


    if [[ "$manualPath" == "~" ]]; then

        manualPath="$HOME"

    elif [[ "$manualPath" == "~/"* ]]; then

        manualPath="$HOME/${manualPath#\~/}"

    fi


    if [[ ! -d "$manualPath" ]]; then

        echo ""
        echo "${C_RED} La ruta indicada no existe.${C_RESET}"
        exit 1

    fi


    resolved="$(canonical_path "$manualPath")"


    if is_auxiliary_path "$resolved"; then

        echo ""
        echo "${C_RED} Esta ubicación corresponde a contenido auxiliar.${C_RESET}"
        exit 1

    fi


    selectedPaths=("$resolved")
    selectedTypes=("Manual")


else

    if (( detectedCount == 0 )); then

        echo ""
        echo "${C_RED} Utilice M para introducir una ruta manual.${C_RESET}"
        exit 1

    fi


    parts=("${(@s:,:)selection}")


    for part in "${parts[@]}"; do

        part="${part//[[:space:]]/}"


        if [[ "$part" != <-> ]]; then

            echo ""
            echo "${C_RED} Selección no válida: $part${C_RESET}"
            exit 1

        fi


        index=$part


        if (( index < 1 || index > detectedCount )); then

            echo ""
            echo "${C_RED} Número fuera de rango: $index${C_RESET}"
            exit 1

        fi


        candidate="${syncPaths[$index]}"
        alreadySelected=0


        for existing in "${selectedPaths[@]}"; do

            if [[ "$existing" == "$candidate" ]]; then

                alreadySelected=1
                break

            fi

        done


        if (( alreadySelected == 0 )); then

            selectedPaths+=("${syncPaths[$index]}")
            selectedTypes+=("${syncTypes[$index]}")

        fi

    done

fi


# ============================================================
# ELIMINAR RAÍCES ANIDADAS DEL RECORRIDO
# ============================================================

typeset -a scanRoots
typeset -a scanRootTypes

scanRoots=()
scanRootTypes=()


for (( i=1; i<=${#selectedPaths[@]}; i++ )); do

    candidate="${selectedPaths[$i]}"
    candidateType="${selectedTypes[$i]}"

    covered=0


    for existing in "${scanRoots[@]}"; do

        if path_is_within "$candidate" "$existing"; then

            covered=1
            break

        fi

    done


    if (( covered == 0 )); then

        typeset -a newRoots
        typeset -a newTypes

        newRoots=()
        newTypes=()


        for (( j=1; j<=${#scanRoots[@]}; j++ )); do

            existing="${scanRoots[$j]}"

            if path_is_within "$existing" "$candidate"; then
                continue
            fi

            newRoots+=("$existing")
            newTypes+=("${scanRootTypes[$j]}")

        done


        scanRoots=("${newRoots[@]}")
        scanRootTypes=("${newTypes[@]}")

        scanRoots+=("$candidate")
        scanRootTypes+=("$candidateType")

    fi

done


# ============================================================
# PREPARAR INTERFAZ
# ============================================================

section "LISTO PARA ANALIZAR"

status_line \
    "Ubicaciones seleccionadas" \
    "${#selectedPaths[@]}" \
    "$C_CYAN"

status_line \
    "Raices físicas a recorrer" \
    "${#scanRoots[@]}" \
    "$C_CYAN"


echo ""
echo " Umbrales activos:"
echo ""
echo "   Cloud          : aviso desde 360 | crítico sobre 400"
echo "   OneDrive Sync  : aviso desde 480 | crítico sobre 520"
echo "   Finder         : aviso desde 900 | crítico sobre 1024 caracteres"
echo "   Office macOS   : aviso desde 900 | crítico sobre 1024 bytes UTF-8"
echo "   Nombre         : aviso desde 240 | crítico sobre 255"
echo ""


# ============================================================
# CSV TEMPORAL
# ============================================================

timestamp="$(date '+%Y%m%d_%H%M%S')"

desktop="$HOME/Desktop"

if [[ ! -d "$desktop" ]]; then
    desktop="$PWD"
fi


outputFile="$desktop/OneDrive_Path_Issues_macOS_${timestamp}.csv"

temporaryFile="$(mktemp -t OneDrivePathAnalyzer)"


# Encabezados orientados al cliente
echo '"Estado","Motivo","TipoSincronizacion","TipoElemento","LongitudRutaLocal","BytesUTF8RutaLocal","LongitudRutaCloud","LongitudNombre","Nombre","RutaCompleta"' \
    > "$temporaryFile"


# ============================================================
# CONTADORES
# ============================================================

totalScanned=0

warningCount=0
criticalCount=0

longestChars=0
longestBytes=0
longestPath=""

findingCount=0


# ============================================================
# ESCANEO
# ============================================================

section "ANALIZANDO"


for (( rootIndex=1; rootIndex<=${#scanRoots[@]}; rootIndex++ )); do

    rootPath="${scanRoots[$rootIndex]}"
    rootType="${scanRootTypes[$rootIndex]}"

    echo ""
    echo " ${C_CYAN}[$rootIndex/${#scanRoots[@]}]${C_RESET} $rootPath"


    while IFS= read -r -d '' fullPath; do

        # ----------------------------------------------------
        # OMITIR AUXILIARES
        # ----------------------------------------------------

        if is_auxiliary_path "$fullPath"; then
            continue
        fi


        (( totalScanned++ ))


        # ----------------------------------------------------
        # ENCONTRAR RAÍZ LÓGICA MÁS ESPECÍFICA
        # ----------------------------------------------------

        logicalRoot="$rootPath"
        logicalType="$rootType"

        longestLogicalLength=${#logicalRoot}


        for (( i=1; i<=detectedCount; i++ )); do

            candidateRoot="${syncPaths[$i]}"

            if path_is_within "$fullPath" "$candidateRoot"; then

                candidateLength=${#candidateRoot}

                if (( candidateLength > longestLogicalLength )); then

                    logicalRoot="$candidateRoot"
                    logicalType="${syncTypes[$i]}"
                    longestLogicalLength=$candidateLength

                fi

            fi

        done


        # ----------------------------------------------------
        # RUTA CLOUD
        # ----------------------------------------------------

        if [[ "$fullPath" == "$logicalRoot" ]]; then

            cloudPath=""

        else

            cloudPath="${fullPath#$logicalRoot/}"

        fi


        # ----------------------------------------------------
        # MEDICIONES
        # ----------------------------------------------------

        charLength=${#fullPath}
        byteLength=$(byte_length "$fullPath")

        cloudLength=${#cloudPath}

        fileName="${fullPath:t}"
        nameLength=${#fileName}


        # ----------------------------------------------------
        # TIPO
        # ----------------------------------------------------

        if [[ -d "$fullPath" ]]; then

            itemType="Carpeta"

        else

            itemType="Archivo"

        fi


        # ----------------------------------------------------
        # MÁS LARGA
        # ----------------------------------------------------

        if (( charLength > longestChars )); then

            longestChars=$charLength
            longestBytes=$byteLength
            longestPath="$fullPath"

        fi


        # ----------------------------------------------------
        # DIAGNÓSTICO
        # ----------------------------------------------------

        criticalReasons=""
        warningReasons=""


        # Cloud

        if (( cloudLength > LIMIT_CLOUD )); then

            append_critical \
                "Ruta cloud > $LIMIT_CLOUD"

        elif (( cloudLength >= WARN_CLOUD )); then

            append_warning \
                "Ruta cloud cerca de $LIMIT_CLOUD"

        fi


        # OneDrive Sync

        if (( charLength > LIMIT_SYNC )); then

            append_critical \
                "OneDrive Sync > $LIMIT_SYNC"

        elif (( charLength >= WARN_SYNC )); then

            append_warning \
                "OneDrive Sync cerca de $LIMIT_SYNC"

        fi


        # Finder

        if (( charLength > LIMIT_FINDER )); then

            append_critical \
                "Finder > $LIMIT_FINDER caracteres"

        elif (( charLength >= WARN_FINDER )); then

            append_warning \
                "Finder cerca de $LIMIT_FINDER caracteres"

        fi


        # Office macOS

        if (( byteLength > LIMIT_OFFICE_BYTES )); then

            append_critical \
                "Office macOS > $LIMIT_OFFICE_BYTES bytes UTF-8"

        elif (( byteLength >= WARN_OFFICE_BYTES )); then

            append_warning \
                "Office macOS cerca de $LIMIT_OFFICE_BYTES bytes UTF-8"

        fi


        # Nombre

        if (( nameLength > LIMIT_NAME )); then

            append_critical \
                "Nombre individual > $LIMIT_NAME"

        elif (( nameLength >= WARN_NAME )); then

            append_warning \
                "Nombre individual cerca de $LIMIT_NAME"

        fi


        # ----------------------------------------------------
        # ESTADO FINAL
        # ----------------------------------------------------

        state="OK"
        reason=""


        if [[ -n "$criticalReasons" ]]; then

            state="CRITICO"
            reason="$criticalReasons"

            if [[ -n "$warningReasons" ]]; then
                reason="$reason; $warningReasons"
            fi

            (( criticalCount++ ))
            (( findingCount++ ))

        elif [[ -n "$warningReasons" ]]; then

            state="AVISO"
            reason="$warningReasons"

            (( warningCount++ ))
            (( findingCount++ ))

        fi


        # ----------------------------------------------------
        # EXPORTAR SOLO HALLAZGOS
        # ----------------------------------------------------

        if [[ "$state" != "OK" ]]; then

            {
                csv_escape "$state"
                printf ","

                csv_escape "$reason"
                printf ","

                csv_escape "$logicalType"
                printf ","

                csv_escape "$itemType"
                printf ","

                printf "%d,%d,%d,%d," \
                    "$charLength" \
                    "$byteLength" \
                    "$cloudLength" \
                    "$nameLength"

                csv_escape "$fileName"
                printf ","

                csv_escape "$fullPath"

                printf "\n"

            } >> "$temporaryFile"

        fi


        # ----------------------------------------------------
        # PROGRESO
        # ----------------------------------------------------

        if (( totalScanned % 250 == 0 )); then

            printf "\r Analizadas: %-10d | Avisos: %-6d | Criticas: %-6d" \
                "$totalScanned" \
                "$warningCount" \
                "$criticalCount"

        fi


    done < <(
        find "$rootPath" \
            -mindepth 1 \
            -print0 \
            2>/dev/null
    )

done


printf "\r                                                                    \r"


# ============================================================
# RESULTADO
# ============================================================

section "RESULTADO"

status_line \
    "Rutas de usuario analizadas" \
    "$totalScanned" \
    "$C_WHITE"

status_line \
    "Rutas en AVISO" \
    "$warningCount" \
    "$C_YELLOW"

status_line \
    "Rutas CRITICAS" \
    "$criticalCount" \
    "$C_RED"


if (( longestChars > 0 )); then

    echo ""
    echo "${C_CYAN} Ruta de usuario más larga:${C_RESET}"
    echo "   $longestChars caracteres"
    echo "   $longestBytes bytes UTF-8"
    echo "   ${C_GRAY}$longestPath${C_RESET}"

fi


echo ""


# ============================================================
# ESTADO GENERAL
# ============================================================

if (( criticalCount > 0 )); then

    echo "${C_RED} RESULTADO: CRITICO${C_RESET}"
    echo "${C_RED} Se encontraron rutas que deberían corregirse.${C_RESET}"

elif (( warningCount > 0 )); then

    echo "${C_YELLOW} RESULTADO: REQUIERE REVISION${C_RESET}"
    echo "${C_YELLOW} Se encontraron rutas cercanas a los límites.${C_RESET}"

else

    echo "${C_GREEN} RESULTADO: OK${C_RESET}"
    echo "${C_GREEN} No se encontraron rutas cercanas o superiores a los límites.${C_RESET}"

fi


# ============================================================
# CSV
# ============================================================

if (( findingCount > 0 )); then

    mv "$temporaryFile" "$outputFile"

    echo ""
    echo "${C_CYAN} CSV generado:${C_RESET}"
    echo "${C_GREEN} $outputFile${C_RESET}"

else

    rm -f "$temporaryFile"

    echo ""
    echo "${C_GRAY} No se generó CSV porque no existen rutas con AVISO o CRITICO.${C_RESET}"

fi


echo ""
echo "${C_CYAN}====================================================================${C_RESET}"
echo "${C_WHITE} Análisis completado.${C_RESET}"
echo "${C_CYAN}====================================================================${C_RESET}"
echo ""