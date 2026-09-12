#requires -Version 7.4
[CmdletBinding(SupportsShouldProcess)]
param([Parameter(Mandatory)][string[]]$ScriptName)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = [IO.File]::ReadAllText((Join-Path $root 'Common/M365Toolkit.Core.ps1'))
foreach ($name in $ScriptName) {
    if ($name -notin @('M365-Inactive-Site-Finder.ps1','M365-Site-Storage-Analyzer.ps1','M365-Tenant-Inventory.ps1','M365-Permissions-Scope-Manager.ps1','M365-Legacy-UserId-Cleaner.ps1','M365-Recycle-Bin-and-Hold-Cleaner.ps1','M365-Version-History-Cleaner.ps1','OneDrive-Path-Analyzer.ps1')) { throw "Not a suite script: $name" }
    $path = Join-Path $root $name
    $source = [IO.File]::ReadAllText($path)
    $source = $source.Replace('if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace))', 'if ($DebugPreference -ne ''SilentlyContinue'' -and -not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace))')
    if ($name -in @('M365-Inactive-Site-Finder.ps1','M365-Site-Storage-Analyzer.ps1','M365-Tenant-Inventory.ps1','M365-Permissions-Scope-Manager.ps1')) {
        $configInitialization = @'
$script:LegacyConfigPath = Join-Path $script:ConfigRoot 'settings.json'
$script:ConfigPath = if ([string]::IsNullOrWhiteSpace($env:M365_PNP_TOOLKIT_HOME)) {
    $script:LegacyConfigPath
}
else {
    Join-Path $script:ConfigRoot (($script:AppName -replace ' ', '-') + '.settings.json')
}
'@
        $source = $source.Replace('$script:ConfigPath = Join-Path $script:ConfigRoot ''settings.json''', $configInitialization)
    }
    if ($name -in @('M365-Inactive-Site-Finder.ps1','M365-Site-Storage-Analyzer.ps1','M365-Tenant-Inventory.ps1')) {
        $errorPairs = @{
            'No se pudo determinar la URL del centro de administración.' = 'Could not determine the admin center URL.'
            'La carpeta de reportes no está disponible.' = 'The reports folder is unavailable.'
            'No se pudo obtener una conexión PnP válida.' = 'Could not obtain a valid PnP connection.'
        }
        foreach ($message in $errorPairs.Keys) {
            $source = $source.Replace("throw '$message'", "throw (Get-Text '$message' '$($errorPairs[$message])')")
        }
        $source = $source.Replace('throw "Client ID inválido: ''$ClientId''."', 'throw (Get-ToolkitString Invalid)')
    }
    $source = [regex]::Replace($source, '(?s)# BEGIN M365 TOOLKIT CORE.*?# END M365 TOOLKIT CORE\r?\n', '')
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    $functions = @($ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true))
    $legacy = $name -notin @('M365-Inactive-Site-Finder.ps1','M365-Site-Storage-Analyzer.ps1','M365-Tenant-Inventory.ps1')
    $appName = if ($name -eq 'OneDrive-Path-Analyzer.ps1') { "'OneDrive / SharePoint Path Analyzer'" } else { '$script:AppName' }
    $appVersion = if ($name -eq 'OneDrive-Path-Analyzer.ps1') { "'2.0'" } else { '$script:AppVersion' }
    $edits = [Collections.Generic.List[object]]::new()
    $reporterFunctions = @{}
    if (-not $legacy) {
        $reporterAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Common/M365Toolkit.Reporter.ps1'),[ref]$null,[ref]$errors)
        if ($errors.Count) { throw ($errors | Out-String) }
        foreach ($fn in $reporterAst.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)) { $reporterFunctions[$fn.Name] = $fn.Extent.Text }
    }
    foreach ($fn in $functions) {
        if ($reporterFunctions.ContainsKey($fn.Name)) {
            $edits.Add(@{ Start=$fn.Extent.StartOffset; Length=$fn.Extent.EndOffset-$fn.Extent.StartOffset; Text=$reporterFunctions[$fn.Name] })
            continue
        }
        $body = $null
        $paramText = if ($fn.Body.ParamBlock) { $fn.Body.ParamBlock.Extent.Text } else { '' }
        switch ($fn.Name) {
            'Clear-AppScreen' { $body = "$paramText`n    try { if (-not [Console]::IsOutputRedirected) { Clear-Host } } catch { }" }
            'Get-RegisteredClientIdFromResult' { $body = "$paramText`n    return Get-ToolkitRegistrationClientId -Result `$Result" }
            { $_ -in 'Wait-App','Pause-Tui' } { $body = "$paramText`n    [void](Read-ToolkitInput ('  ' + (Get-ToolkitString Continue)))" }
            'Load-Settings' {
                if ($source.Contains('$script:LegacyConfigPath') -and -not $fn.Extent.Text.Contains('$readPath')) {
                    $originalBody = $fn.Body.Extent.Text.Substring(1, $fn.Body.Extent.Text.Length - 2).Replace('$script:ConfigPath','$readPath')
                    $body = '$readPath = if (Test-Path -LiteralPath $script:ConfigPath) { $script:ConfigPath } else { $script:LegacyConfigPath }' + "`n" + $originalBody
                }
            }
            'Read-MenuChoice' {
                if ($name -eq 'M365-Legacy-UserId-Cleaner.ps1') {
                    $body = "$paramText`n    return Read-ToolkitLegacyMenu -Items `$Items -Mode Legacy -Title `$Title -Description `$Description -ZeroLabel `$ZeroLabel -RenderBody `$RenderBody"
                }
                elseif ($legacy) { $body = "$paramText`n    return Read-ToolkitLegacyMenu -Items `$Items -Mode Key -Prompt `$Prompt -AllowBack:`$AllowBack -RenderBody `$RenderBody" }
            }
            'Read-AppMenuChoice' { $body = "$paramText`n    return Read-ToolkitLegacyMenu -Items `$Items -Mode Object -Prompt `$Prompt -AllowBack:`$AllowBack" }
            'Show-NumberMenu' { $body = "$paramText`n    return Read-ToolkitLegacyMenu -Items `$Items -Mode Indexed -Title `$Title -Description `$Description -RenderBody `$RenderBody" }
            { $_ -in 'Write-AppHeader','Show-AppHeader' } {
                $section = if ($fn.Body.ParamBlock.Parameters.Name.VariablePath.UserPath -contains 'Context') { '$Context' } else { '$Section' }
                if ($legacy) { $section = "(Get-LocalizedText $section)" }
                $body = "$paramText`n    Write-ToolkitHeader $appName $appVersion $section"
            }
            { $_ -in 'Write-Field','Write-AppField' } {
                $label = if ($legacy) { '(Get-LocalizedText $Name)' } else { '$Name' }
                $body = "$paramText`n    Write-ToolkitField $label `$Value `$Style"
            }
            { $_ -in 'Write-Section','Write-AppSection' } {
                # Old unused Path Write-Section uses a positional Text parameter.
                if ($fn.Body.ParamBlock.Parameters.Name.VariablePath.UserPath -contains 'Title') {
                    $label = if ($legacy) { '(Get-LocalizedText $Title)' } else { '$Title' }
                    $body = "$paramText`n    Write-ToolkitText ''`n    Write-ToolkitText $label Primary"
                }
            }
            'Write-Status' {
                $level = if ($fn.Body.ParamBlock.Parameters.Name.VariablePath.UserPath -contains 'Kind') { '$Kind' } else { '$Level' }
                $message = if ($legacy) { '(Get-LocalizedText $Message)' } else { '$Message' }
                $body = "$paramText`n    Write-ToolkitStatus $level $message"
            }
            { $_ -in 'Write-AppInfo','Write-AppOk','Write-AppWarning','Write-AppError' } {
                $level = @{ 'Write-AppInfo'='Info'; 'Write-AppOk'='Ok'; 'Write-AppWarning'='Warn'; 'Write-AppError'='Error' }[$fn.Name]
                $body = "$paramText`n    Write-ToolkitStatus $level (Get-LocalizedText `$Message)"
            }
            { $_ -in 'Read-YesNo','Read-AppYesNo' } {
                $default = if ($fn.Name -eq 'Read-AppYesNo') { '$DefaultYes' } else { '$Default' }
                $prompt = if ($legacy) { '(Get-LocalizedText $Prompt)' } else { '$Prompt' }
                $body = "$paramText`n    return Read-ToolkitYesNo $prompt $default"
            }
            'Initialize-AppLanguage' {
                $body = "param([string]`$Default = 'en')`n    Initialize-ToolkitLanguage -Default `$Default -Name $appName -Version $appVersion"
            }
            { $_ -in 'Ensure-PnPModule','Confirm-AppPnP','Install-PnPModule','Install-AppPnP' } {
                $update = if ($fn.Name -like 'Install-*') { ' -Update' } else { '' }
                $body = "$paramText`n    return Initialize-ToolkitPnP -MinimumVersion '3.2.0'$update"
            }
            'Test-ClientIdFormat' { $body = "$paramText`n    return Test-ToolkitClientId `$ClientId" }
            'Test-AppGuid' { $body = "$paramText`n    return Test-ToolkitClientId `$Value" }
            { $_ -in 'Write-Styled','Write-AppStyled' } {
                # Legacy callers still rely on phrase translation; no double Write-Host translation.
                $text = if ($legacy) { '(Get-LocalizedText $Text)' } else { '$Text' }
                $body = "$paramText`n    Write-ToolkitText $text `$Style -NoNewline:`$NoNewline"
            }
        }
        if ($null -ne $body) { $edits.Add(@{ Start=$fn.Extent.StartOffset; Length=$fn.Extent.EndOffset-$fn.Extent.StartOffset; Text="function $($fn.Name) {`n    $($body.Trim())`n}" }) }
    }
    foreach ($edit in ($edits | Sort-Object Start -Descending)) { $source = $source.Remove($edit.Start,$edit.Length).Insert($edit.Start,$edit.Text) }
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
    $first = $ast.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true)
    $source = $source.Insert($first.Extent.StartOffset,"# BEGIN M365 TOOLKIT CORE`n$core`n# END M365 TOOLKIT CORE`n")
    $null = [System.Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    if ($PSCmdlet.ShouldProcess($path,'Synchronize embedded canonical helpers')) {
        [IO.File]::WriteAllText($path,$source,[Text.UTF8Encoding]::new($true))
        Write-Output "PASS parser: $name"
        if (-not $legacy) { & (Join-Path $PSScriptRoot 'Convert-ReporterLocalization.ps1') -ScriptName $name }
    }
}
