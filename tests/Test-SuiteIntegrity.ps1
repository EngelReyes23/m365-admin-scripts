#requires -Version 7.4
[CmdletBinding()]
param([string]$BaselineDirectory)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = [IO.File]::ReadAllText((Join-Path $root 'Common/M365Toolkit.Core.ps1')).Trim()
$protected = @(
    'Convert-SiteToResult','Convert-SiteToStorageResult','Convert-SiteToInventoryRecord','Convert-MbToGb',
    'Reset-ItemInheritanceViaRest','Reset-LibraryInheritanceViaRest','Get-AppFileVersions','Get-AppVersionSize',
    'Invoke-AppVersionAnalysis','Remove-AppVersionSafely','Find-AffectedUserInSite','Remove-LegacyUserFromSite',
    'Get-PathDiagnostic','Scan-DirectoryRecursive','Get-MinimalScanRoots','Get-AnalysisLocations',
    'Get-AppRecycleBinItems','Move-AppRecycleBinItemToSecondStage','Get-AppPreservationHoldLibrary'
)
foreach ($file in Get-ChildItem -LiteralPath $root -File -Filter '*.ps1') {
    $source = [IO.File]::ReadAllText($file.FullName)
    $embedded = [regex]::Match($source,'(?s)# BEGIN M365 TOOLKIT CORE\r?\n(.*?)\r?\n# END M365 TOOLKIT CORE')
    if (-not $embedded.Success -or $embedded.Groups[1].Value.Trim() -cne $core) { throw "Core drift: $($file.Name)" }
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    $functions = @($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst]},$true))
    $map = @{}
    foreach ($fn in $functions) {
        if ($map.ContainsKey($fn.Name)) { throw "Duplicate function: $($file.Name) / $($fn.Name)" }
        $map[$fn.Name] = $fn.Extent.Text
    }
    $checks = 0
    if ($BaselineDirectory) {
        $baseline = [Management.Automation.Language.Parser]::ParseFile((Join-Path $BaselineDirectory $file.Name),[ref]$null,[ref]$errors)
        if ($errors.Count) { throw ($errors | Out-String) }
        foreach ($fn in $baseline.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)) {
            if (-not $map.ContainsKey($fn.Name)) { throw "Original function removed: $($file.Name)/$($fn.Name)" }
            if ($fn.Name -in $protected) {
                if ($map[$fn.Name] -cne $fn.Extent.Text) { throw "Business function changed: $($file.Name)/$($fn.Name)" }
                $checks++
            }
        }
        $beforeExports = @($baseline.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Export-Csv'},$true) | ForEach-Object {$_.Extent.Text})
        $afterExports = @($ast.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Export-Csv'},$true) | ForEach-Object {$_.Extent.Text})
        if (($beforeExports -join "`n") -cne ($afterExports -join "`n")) { throw "CSV command changed: $($file.Name)" }
    }
    Write-Output "PASS integrity: $($file.Name) (identical core, no duplicate/missing functions, $checks unchanged business functions, CSV commands)"
}
