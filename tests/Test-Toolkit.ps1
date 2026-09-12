#requires -Version 7.4
# Offline regression tests: only function ASTs are loaded, never script entrypoints.
[CmdletBinding()]
param(
    [string[]]$ScriptName = @('M365-Inactive-Site-Finder.ps1'),
    [string]$ScriptDirectory
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if ([string]::IsNullOrWhiteSpace($ScriptDirectory)) { $ScriptDirectory = $root }
function Assert-Equal($Actual, $Expected, $Label) {
    if ($Actual -cne $Expected) { throw "$Label : expected <$Expected>, got <$Actual>" }
}
foreach ($name in $ScriptName) {
    & {
        Set-StrictMode -Version 3.0
        $path = Join-Path $ScriptDirectory $name
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$parseErrors)
        if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
        $functions = @($ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]},$true))
        foreach ($fn in $functions | Where-Object Name -Match 'Toolkit|^Get-LocalizedString$|^Get-Text$|^Test-Is(SystemSite|OneDriveUrl)$|^Get-PropertyValue$|^Test-ClientIdFormat$|^Test-PnPAppRegistration$|^Register-NewPnPApp$|^Set-ExistingPnPClientId$|^Ensure-ClientId$') {
            . ([scriptblock]::Create($fn.Extent.Text))
        }
        $toolCatalog = $ast.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$script:ToolStrings'},$true)
        if ($toolCatalog) {
            . ([scriptblock]::Create($toolCatalog.Extent.Text))
            Assert-Equal (($script:ToolStrings.en.Keys | Sort-Object) -join ',') (($script:ToolStrings.es.Keys | Sort-Object) -join ',') 'Tool localization keys'
            foreach ($language in @('en','es')) {
                $script:Language = $language
                foreach ($key in $script:ToolStrings[$language].Keys) { $null = Get-LocalizedString -Key $key -Values @('A','B','C','D','E','F','G','H','I','J') }
            }
        }
        $script:Language = 'en'
        Assert-Equal (Get-ToolkitString Back) 'Back' 'English'
        $script:Language = 'es'
        Assert-Equal (Get-ToolkitString Back) 'Volver' 'Spanish'
        $stringFn = $functions | Where-Object Name -EQ Get-ToolkitString
        $catalogAssignment = $stringFn.Body.Find({param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$catalog'},$true)
        $catalog = & ([scriptblock]::Create($catalogAssignment.Right.Extent.Text))
        Assert-Equal (($catalog.en.Keys | Sort-Object) -join ',') (($catalog.es.Keys | Sort-Object) -join ',') 'Localization keys'
        foreach ($key in $catalog.en.Keys) {
            Assert-Equal (([regex]::Matches($catalog.en[$key],'\{\d+\}') | ForEach-Object Value | Sort-Object) -join ',') (([regex]::Matches($catalog.es[$key],'\{\d+\}') | ForEach-Object Value | Sort-Object) -join ',') "Format placeholders $key"
        }
        Assert-Equal (Test-ToolkitClientId '00000000-0000-0000-0000-000000000000') $false 'Empty GUID'
        Assert-Equal (Test-ToolkitClientId 'not-a-guid') $false 'Invalid GUID'
        Assert-Equal (Test-ToolkitClientId '9717f640-f90a-4072-82dd-d24c16570d60') $true 'Client GUID'
        foreach ($case in @(
            @('https://example.sharepoint.com/sites/admin-project','SharePoint'),
            @('https://example-admin.sharepoint.com/','Admin'),
            @('https://example-my.sharepoint.com/','OneDriveSystem'),
            @('https://example-my.sharepoint.com/personal/user_example_com/','OneDrive'),
            @('https://example-my.sharepoint.com/personal/','OneDriveSystem'),
            @('https://example.sharepoint.com.attacker.invalid/sites/a','Invalid'),
            @('http://example.sharepoint.com/sites/a','Invalid'),
            @('https://user@example.sharepoint.com/sites/a','Invalid')
        )) { Assert-Equal (Get-ToolkitSiteEndpoint $case[0]) $case[1] "Endpoint $($case[0])" }
        function Get-LocalizedText { param($Text) return $Text }
        function Write-ToolkitText { param($Text,$Style,[switch]$NoNewline) $script:Captured.Add([string]$Text) }
        $script:Captured = [Collections.Generic.List[string]]::new()
        Write-ToolkitField -Name 'Path' -Value 'C:\sitios\Eliminar archivos\Informe.csv'
        Assert-Equal $script:Captured[1] 'C:\sitios\Eliminar archivos\Informe.csv' 'Data is not translated'
        function Read-ToolkitInput { param($Prompt) return $script:Inputs.Dequeue() }
        $script:Inputs = [Collections.Generic.Queue[string]]::new()
        function Clear-Host { }
        $script:Inputs.Enqueue('1')
        Initialize-ToolkitLanguage -Default es -Name 'Test' -Version '1.0'
        Assert-Equal $script:Language 'en' 'Language 1 English'
        $script:Inputs.Enqueue('2')
        Initialize-ToolkitLanguage -Default en -Name 'Test' -Version '1.0'
        Assert-Equal $script:Language 'es' 'Language 2 Spanish'
        $script:Inputs.Enqueue('')
        Initialize-ToolkitLanguage -Default es -Name 'Test' -Version '1.0'
        Assert-Equal $script:Language 'es' 'Saved language default'
        $script:Inputs.Enqueue('')
        Assert-Equal (Read-ToolkitYesNo 'Test') $false 'Confirmation defaults to no'
        foreach ($language in @('en','es')) {
            $script:Language = $language
            foreach ($action in @('Repair','Delete','Move','DeletePHL')) {
                $script:Inputs.Enqueue((Get-ToolkitString $action))
                Assert-Equal (Confirm-ToolkitAction $action) $true "Confirmation $language/$action"
                $script:Inputs.Enqueue('')
                Assert-Equal (Confirm-ToolkitAction $action) $false "Cancel $language/$action"
                $script:Inputs.Enqueue('wrong')
                Assert-Equal (Confirm-ToolkitAction $action) $false "Wrong phrase $language/$action"
            }
        }
        $sampleId = '9717f640-f90a-4072-82dd-d24c16570d60'
        Assert-Equal (Get-ToolkitRegistrationClientId $null) '' 'Null registration result'
        Assert-Equal (Get-ToolkitRegistrationClientId ([pscustomobject]@{Id=$sampleId})) '' 'Directory object ID rejected'
        Assert-Equal (Get-ToolkitRegistrationClientId @($null,@{AppId=$sampleId})) $sampleId 'Array/dictionary registration result'
        Assert-Equal (Get-ToolkitRegistrationClientId $sampleId) $sampleId 'String registration result'
        $script:Inputs.Enqueue('invalid'); $script:Inputs.Enqueue('2')
        $items = @([pscustomobject]@{Label='First';Value='First'},[pscustomobject]@{Label='Second';Value='Second'},[pscustomobject]@{Label='Exit';Value='Exit'})
        Assert-Equal (Read-ToolkitLegacyMenu -Mode Indexed -Items $items).Value 'Second' 'Indexed menu preserves Value'
        $script:Inputs.Enqueue('0')
        Assert-Equal (Read-ToolkitLegacyMenu -Mode Indexed -Items $items).Value 'Exit' 'Indexed zero returns original Exit'
        $script:Inputs.Enqueue('0')
        Assert-Equal (Read-ToolkitLegacyMenu -Mode Legacy -Items $items) $null 'Legacy zero returns null'
        $script:Inputs.Enqueue('1')
        Assert-Equal (Read-ToolkitLegacyMenu -Mode Key -Items @(@{Key='1';Label='One'})) '1' 'Key menu dictionary'
        $script:Inputs.Enqueue('0')
        Assert-Equal (Read-ToolkitLegacyMenu -Mode Object -Items @(@{Key='1';Label='One'}) -AllowBack).Value 'Back' 'Object menu synthetic Back'
        function Write-ToolkitStatus { param($Level,$Message) }
        function Get-ToolkitCompatiblePnP { param($MinimumVersion) return $script:MockModule }
        function Read-ToolkitYesNo { param($Prompt,$Default) return $script:ApproveInstall }
        function Install-Module { param($Name,$RequiredVersion,$MinimumVersion,$Scope,[switch]$Force,[switch]$AllowClobber,$ErrorAction) $script:InstallCount++; if ($script:InstallFails) { throw 'mock installation failure' }; $script:MockModule = [pscustomobject]@{Version=[version]'3.2.0';Path='mock.psd1'} }
        function Import-Module { param($Name,$ErrorAction) if ($script:ImportFails) { throw 'mock import failure' } }
        function Get-Module { param($Name) return $script:MockModule }
        $script:InstallFails=$false; $script:ImportFails=$false; $script:InstallCount=0
        $script:MockModule=$null; $script:ApproveInstall=$false
        Assert-Equal (Initialize-ToolkitPnP) $false 'Missing module / cancel'
        Assert-Equal $script:InstallCount 0 'No silent install'
        $script:ApproveInstall=$true
        Assert-Equal (Initialize-ToolkitPnP) $true 'Missing module / install / import'
        Assert-Equal $script:InstallCount 1 'Single installation'
        Assert-Equal (Initialize-ToolkitPnP) $true 'Installed module'
        Assert-Equal $script:InstallCount 1 'Existing module not installed again'
        $script:ImportFails=$true
        Assert-Equal (Initialize-ToolkitPnP) $false 'Import failure is not success'
        $script:ImportFails=$false; $script:MockModule=$null; $script:InstallFails=$true
        Assert-Equal (Initialize-ToolkitPnP) $false 'Installation failure'
        if ($name -in @('M365-Inactive-Site-Finder.ps1','M365-Site-Storage-Analyzer.ps1','M365-Tenant-Inventory.ps1')) {
            . ([scriptblock]::Create(($functions | Where-Object Name -EQ Read-MenuChoice).Extent.Text))
            $script:Inputs.Enqueue('1')
            Assert-Equal (Read-MenuChoice -Items @(@{Key='1';Label='One'})) '1' 'Reporter menu dictionary without optional description'
            $script:Inputs.Enqueue('0')
            Assert-Equal (Read-MenuChoice -Items @([pscustomobject]@{Key='1';Label='One'}) -AllowBack) '0' 'Reporter menu cancellation'
            $script:Inputs.Enqueue('0')
            Assert-Equal (Read-MenuChoice -Items @() -AllowBack) '0' 'Empty reporter menu cancellation'
            foreach ($url in @('https://example.sharepoint.com/sites/search-project','https://example.sharepoint.com/sites/appcatalogue','https://example.sharepoint.com/')) {
                Assert-Equal (Test-IsSystemSite ([pscustomobject]@{Url=$url;Template='STS#3';Status='Locked'})) $false 'Legitimate locked/customer site'
            }
            Assert-Equal (Test-IsSystemSite ([pscustomobject]@{Url='https://example-my.sharepoint.com/';Template=''})) $true 'OneDrive technical root excluded'
            Assert-Equal (Test-IsOneDriveUrl 'https://example-my.sharepoint.com/personal/user/') $true 'Personal OneDrive'
            function Connect-M365AdminWithClientId { param($ClientId) return [pscustomobject]@{Mock=$true} }
            function Get-PnPTenant { param($Connection,$ErrorAction) if ($script:ValidationFails) { throw 'mock access denied' } }
            function Get-PnPTenantSite { param($Connection,$ErrorAction) return [pscustomobject]@{Url='https://example.sharepoint.com'} }
            function Disconnect-PnPOnline { param($Connection,$ErrorAction) $script:DisconnectCount++ }
            function Ensure-PnPModule { param([switch]$OfferInstall) return $true }
            function Ensure-TenantConfigured { return $true }
            function Write-Status { param($Level,$Message) }
            function Write-Styled { param($Text,$Style) }
            function Write-Section { param($Title) }
            function Write-Field { param($Name,$Value) }
            function Save-Settings { $script:SaveCount++ }
            function Read-YesNo { param($Prompt,$Default) return $script:ApproveRegistration }
            function Read-TextValue { param($Prompt,$Default,[switch]$AllowEmpty,$Validator,$ValidationMessage) return $script:TextInputs.Dequeue() }
            function Register-PnPEntraIDAppForInteractiveLogin { param($ApplicationName,$Tenant,$SharePointDelegatePermissions,$GraphDelegatePermissions,$ErrorAction) $script:RegistrationCount++; return $script:RegistrationResult }
            $script:TextInputs = [Collections.Generic.Queue[string]]::new()
            $script:AppName='Test reporter'; $script:SaveCount=0; $script:DisconnectCount=0
            $script:ValidationFails=$true; $script:RegistrationCount=0; $script:ApproveRegistration=$true
            $old='9717f640-f90a-4072-82dd-d24c16570d60'; $new='6d6026e2-b0aa-48a2-a445-78754de77dd4'
            $script:Settings=[pscustomobject]@{Tenant='example.onmicrosoft.com';ClientId=$old;AppRegistrationName='Previous'}
            Assert-Equal (Test-PnPAppRegistration $new -Quiet) $false 'App validation denies access'
            Assert-Equal $script:DisconnectCount 1 'Failed app validation cleans up'
            $script:TextInputs.Enqueue($new)
            Assert-Equal (Set-ExistingPnPClientId) $false 'Existing application failed validation'
            Assert-Equal $script:Settings.ClientId $old 'Old configuration preserved'
            $script:ValidationFails=$false
            $script:TextInputs.Enqueue($new)
            Assert-Equal (Set-ExistingPnPClientId) $true 'Existing application valid'
            Assert-Equal $script:Settings.ClientId $new 'Valid client saved'
            $script:Settings.ClientId=$old
            $script:RegistrationResult=[pscustomobject]@{AppId=$new}
            $script:ApproveRegistration=$false; $script:TextInputs.Enqueue('New app')
            Assert-Equal (Register-NewPnPApp) $false 'Registration cancelled'
            Assert-Equal $script:RegistrationCount 0 'Cancelled registration has no external write'
            $script:ApproveRegistration=$true; $script:TextInputs.Enqueue('New app')
            Assert-Equal (Register-NewPnPApp) $true 'New application valid'
            Assert-Equal $script:Settings.ClientId $new 'New application saved after validation'
            $script:Settings.ClientId=$old; $script:RegistrationResult=[pscustomobject]@{Id=$old}
            $script:TextInputs.Enqueue('New app'); $script:TextInputs.Enqueue($new)
            Assert-Equal (Register-NewPnPApp) $true 'Object ID requires explicit client ID'
            Assert-Equal $script:Settings.ClientId $new 'Object ID not mistaken for client ID'
        }
        Write-Output "PASS offline: $name (parser, locales, GUID, endpoints, dependency branches)"
    }
}
