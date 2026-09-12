#requires -Version 7.4
# Mechanical migration of existing complete translation pairs, preserving interpolation expressions.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$ScriptName)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if ($ScriptName -notin @('M365-Inactive-Site-Finder.ps1','M365-Site-Storage-Analyzer.ps1','M365-Tenant-Inventory.ps1')) { throw 'Only reporter translation pairs are supported.' }
$path = Join-Path $root $ScriptName
$source = [IO.File]::ReadAllText($path)
$source = [regex]::Replace($source,'(?s)# BEGIN TOOL LOCALIZATION.*?# END TOOL LOCALIZATION\r?\n','')
$catalogPath = Join-Path $root ('Common/Locales/' + [IO.Path]::GetFileNameWithoutExtension($ScriptName) + '.json')
$catalog = @{en=@{};es=@{}}
if (Test-Path -LiteralPath $catalogPath) { $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json -AsHashtable }
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$changes = [Collections.Generic.List[object]]::new()
foreach ($command in $ast.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-Text'},$true)) {
    if ($command.CommandElements.Count -ne 3) { throw "Unexpected translation signature: $($command.Extent.Text)" }
    $values = [Collections.Generic.List[string]]::new()
    $templates = @()
    foreach ($argument in $command.CommandElements | Select-Object -Skip 1) {
        if ($argument -isnot [Management.Automation.Language.StringConstantExpressionAst] -and $argument -isnot [Management.Automation.Language.ExpandableStringExpressionAst]) { throw "Non-string translation pair: $($command.Extent.Text)" }
        $template = $argument.Value.Replace('{','{{').Replace('}','}}')
        if ($argument -is [Management.Automation.Language.ExpandableStringExpressionAst]) {
            foreach ($nested in ($argument.NestedExpressions | Sort-Object {$_.Extent.Text.Length} -Descending)) {
                $expression = $nested.Extent.Text
                if (-not $values.Contains($expression)) { $values.Add($expression) }
                $template = $template.Replace($expression.Replace('{','{{').Replace('}','}}'), ('{' + $values.IndexOf($expression) + '}'))
            }
        }
        $templates += $template
    }
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes(($templates -join "`n"))))
    for ($side = 0; $side -lt 2; $side++) {
        $roundTrip = $templates[$side]
        for ($index = 0; $index -lt $values.Count; $index++) {
            $roundTrip = $roundTrip.Replace(('{' + $index + '}'),$values[$index].Replace('{','{{').Replace('}','}}'))
        }
        $roundTrip = $roundTrip.Replace('{{','{').Replace('}}','}')
        if ($roundTrip -cne $command.CommandElements[$side + 1].Value) { throw "Interpolation migration failed: $($command.Extent.Text)" }
    }
    $key = 'T' + $hash.Substring(0,16)
    $catalog.es[$key] = $templates[0]
    $catalog.en[$key] = $templates[1]
    $call = "Get-LocalizedString -Key '$key'"
    if ($values.Count) { $call += ' -Values @(' + ($values -join ', ') + ')' }
    $changes.Add(@{Start=$command.Extent.StartOffset;Length=$command.Extent.EndOffset-$command.Extent.StartOffset;Text=$call})
}
foreach ($change in $changes | Sort-Object Start -Descending) { $source = $source.Remove($change.Start,$change.Length).Insert($change.Start,$change.Text) }
$lines = [Collections.Generic.List[string]]::new()
$lines.Add('# BEGIN TOOL LOCALIZATION')
$lines.Add('$script:ToolStrings = @{')
foreach ($language in @('en','es')) {
    $lines.Add("    $language = @{")
    foreach ($key in $catalog[$language].Keys | Sort-Object) {
        $text = $catalog[$language][$key].Replace("'","''")
        $lines.Add("        '$key' = '$text'")
    }
    $lines.Add('    }')
}
$lines.Add('}')
$lines.Add('# END TOOL LOCALIZATION')
$ast = [Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
$first = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst]},$true)
# Insert before the core marker, not inside it, so a core sync does not remove translations.
$offset = $source.IndexOf('# BEGIN M365 TOOLKIT CORE')
if ($offset -lt 0) { $offset = $first.Extent.StartOffset }
$source = $source.Insert($offset,($lines -join "`n") + "`n")
$null = [Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
[void][IO.Directory]::CreateDirectory((Split-Path $catalogPath -Parent))
[IO.File]::WriteAllText($catalogPath,($catalog | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText($path,$source,[Text.UTF8Encoding]::new($true))
Write-Output "PASS localized/parser: $ScriptName ($($catalog.en.Count) paired keys)"
