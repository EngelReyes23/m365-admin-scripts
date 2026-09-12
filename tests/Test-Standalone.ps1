#requires -Version 7.4
[CmdletBinding()]
param([string[]]$ScriptName = @(
    'M365-Inactive-Site-Finder.ps1',
    'M365-Legacy-UserId-Cleaner.ps1',
    'M365-Permissions-Scope-Manager.ps1',
    'M365-Recycle-Bin-and-Hold-Cleaner.ps1',
    'M365-Site-Storage-Analyzer.ps1',
    'M365-Tenant-Inventory.ps1',
    'M365-Version-History-Cleaner.ps1',
    'OneDrive-Path-Analyzer.ps1'
))
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$pwsh = (Get-Command pwsh -ErrorAction Stop).Source
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('m365-standalone-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-StandaloneStatic {
    param([string]$Path)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    $source = [IO.File]::ReadAllText($Path)
    if ($source -match '(?im)(?:^|[\s''"])(?:Common|tools|tests)[\\/]') { throw "Repository path reference: $Path" }
    if ($source -match '(?i)\$PSScriptRoot|\.psm1|\.psd1') { throw "Local module/resource reference: $Path" }
    $dotSources = @($ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.CommandAst] -and
        $node.InvocationOperator -eq [Management.Automation.Language.TokenKind]::Dot
    },$true))
    if ($dotSources.Count) { throw "Dot-sourcing found: $Path" }
    foreach ($command in $ast.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst]},$true)) {
        $name = $command.GetCommandName()
        if ($name -in @('Import-Clixml','Import-Csv','Invoke-Expression')) { throw "External resource/code loader $name found: $Path" }
        if ($name -eq 'Import-Module') {
            $text = $command.Extent.Text
            if ($text -notmatch '(?i)PnP\.PowerShell|\$module\.Path') { throw "Non-PnP module import found: $text" }
        }
    }
    return $ast
}

function Invoke-ExitSmoke {
    param([string]$Path,[string]$ProfileRoot)
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $pwsh
    $start.ArgumentList.Add('-NoProfile')
    $start.ArgumentList.Add('-File')
    $start.ArgumentList.Add($Path)
    $start.WorkingDirectory = Split-Path $Path -Parent
    $start.UseShellExecute = $false
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    $start.Environment['M365_PNP_TOOLKIT_HOME'] = Join-Path $ProfileRoot 'toolkit'
    $start.Environment['M365_LEGACY_USERID_CLEANER_HOME'] = Join-Path $ProfileRoot 'legacy'
    $start.Environment['LOCALAPPDATA'] = Join-Path $ProfileRoot 'localappdata'
    # If an unexpected network path is reached, it must fail locally instead of contacting M365.
    $start.Environment['HTTP_PROXY'] = 'http://127.0.0.1:1'
    $start.Environment['HTTPS_PROXY'] = 'http://127.0.0.1:1'
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    [void]$process.Start()
    $process.StandardInput.WriteLine('1')
    $process.StandardInput.WriteLine('0')
    $process.StandardInput.WriteLine('')
    $process.StandardInput.WriteLine('')
    $process.StandardInput.Close()
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(20000)) {
        $process.Kill($true)
        throw "Startup smoke timed out: $Path"
    }
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    if ($process.ExitCode -ne 0) { throw "Startup smoke failed ($($process.ExitCode)): $Path`n$stderr`n$stdout" }
    if ($stdout -match '(?i)Connect-PnP|Register-PnP|installing PnP|instalando PnP|authentication window|ventana de autenticación') {
        throw "Unexpected Microsoft 365/dependency flow during exit smoke: $Path"
    }
}

try {
    foreach ($name in $ScriptName) {
        $source = Join-Path $root $name
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing suite script: $name" }
        $isolated = Join-Path $testRoot ([IO.Path]::GetFileNameWithoutExtension($name))
        $profile = Join-Path $testRoot ('profile-' + [IO.Path]::GetFileNameWithoutExtension($name))
        [void][IO.Directory]::CreateDirectory($isolated)
        [void][IO.Directory]::CreateDirectory($profile)
        $copy = Join-Path $isolated $name
        Copy-Item -LiteralPath $source -Destination $copy
        if (@(Get-ChildItem -LiteralPath $isolated -Force).Count -ne 1) { throw "Isolation setup contains extra files: $isolated" }
        $ast = Assert-StandaloneStatic -Path $copy
        & (Join-Path $PSScriptRoot 'Test-Toolkit.ps1') -ScriptName $name -ScriptDirectory $isolated
        Invoke-ExitSmoke -Path $copy -ProfileRoot $profile
        if (@(Get-ChildItem -LiteralPath $isolated -Force).Count -ne 1) { throw "Script created a repository-local runtime dependency: $name" }
        Write-Output "PASS standalone: $name (one-file parser, offline helpers, direct startup/exit)"
    }
}
finally {
    # Test artifacts live only under the exact GUID-scoped temp directory.
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
