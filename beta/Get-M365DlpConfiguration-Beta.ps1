<#
.SYNOPSIS
    BETA - includes optional Microsoft Graph collection (licensing, DLP alerts, Entra role holders, SharePoint sharing).
    The main, supported version is Get-M365DlpConfiguration.ps1 in the repository root.
    Microsoft Purview Data Loss Prevention (DLP) discovery tool for Microsoft 365 tenants - menu driven.

.DESCRIPTION
    Read-only. Only Get-* and Export-* cmdlets are used against the tenant; nothing is changed.

    Menu options:
      1. Install / update the required PowerShell modules
      2. Connect to tenant (sign in, or switch account/tenant, and show what the account can access)
      3. Test credentials and tenant access (checks every cmdlet the reports need)
      4. Report on individual DLP items (choose one or more areas)
      5. Complete tenant DLP report (every area plus an assessment summary)
      6. Settings (output folder, Activity Explorer, Exchange Online)
      7. Disconnect / switch tenant

    Areas available as individual reports:
       1  Tenant information
       2  DLP policies and workload coverage (Exchange, SharePoint, OneDrive, Teams, Endpoint, Copilot)
       3  DLP rules - conditions, actions, policy tips, block with override
       4  Classifiers and confidence levels used in rules (SITs, EDM, trainable classifiers, labels)
       5  Sensitive information types - built-in and custom, rule packages, keyword dictionaries
       6  Exact Data Match (EDM)
       7  Sensitivity labels, label policies and auto-labelling
       8  Endpoint DLP - USB / removable media, printing, browser and cloud uploads
       9  DLP alerts and alert policies
      10  Operational ownership - Purview role groups and members
      11  Activity Explorer (activity data for the last N days)
      12  Exchange Online mail flow rules and audit configuration

    Output for each report run goes to a new time-stamped folder:
      CSV + JSON per dataset, DLP_Configuration.xlsx (if ImportExcel is installed), Collection.log,
      and for the complete report 00_Assessment_Summary.csv/.html and DLP_Assessment_Report.docx
      (a plain-English Word report explaining each area, the findings and Microsoft Learn links).

.PARAMETER Frameworks
    Frameworks referenced in the Word report: any of NCSC, NISTCSF, NIST80053, ISO27001 (or NIST, ISO, All, None).
    Default: all four.

.PARAMETER IncludeGraph
    Also collect Microsoft Graph data (licensing, DLP alerts, Entra role holders, SharePoint sharing settings)
    in a separate PowerShell process. Needs Microsoft.Graph.Authentication and admin consent for read-only scopes:
    LicenseAssignment.Read.All, SecurityAlert.Read.All, RoleManagement.Read.Directory, SharePointTenantSettings.Read.All.

.PARAMETER NoWordReport
    Do not create DLP_Assessment_Report.docx at the end of the complete report.

.PARAMETER DisableWAM
    Sign in without the Windows Web Account Manager broker (passes -DisableWAM where the module supports it).
    Use if sign-in fails, for example with "You cannot call a method on a null-valued expression".

.PARAMETER Mode
    Menu (default) | InstallModules | TestAccess | Full | Area

.PARAMETER Area
    With -Mode Area: one or more area numbers (1-12) or names, e.g. -Area 2,3,8

.EXAMPLE
    .\Get-M365DlpConfiguration-Beta.ps1
    Starts the interactive menu.

.EXAMPLE
    .\Get-M365DlpConfiguration-Beta.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com -IncludeActivityExplorer
    Runs the complete report without the menu.

.EXAMPLE
    .\Get-M365DlpConfiguration-Beta.ps1 -Mode Area -Area 2,3,8 -UserPrincipalName admin@contoso.onmicrosoft.com
    Reports on policies, rules and Endpoint DLP only.

.EXAMPLE
    .\Get-M365DlpConfiguration-Beta.ps1 -Mode TestAccess -UserPrincipalName admin@contoso.onmicrosoft.com

.NOTES
    Requires: ExchangeOnlineManagement 3.10.1 or later. Optional: ImportExcel; Microsoft.Graph.Authentication 2.41.0+ (for -IncludeGraph).
    Works in Windows PowerShell 5.1 and PowerShell 7+.
#>

[CmdletBinding()]
param(
    [ValidateSet('Menu', 'InstallModules', 'TestAccess', 'Full', 'Area')]
    [string]$Mode = 'Menu',
    [string[]]$Area,
    [string]$OutputRoot = [Environment]::GetFolderPath('MyDocuments'),
    [string]$UserPrincipalName,
    [switch]$IncludeActivityExplorer,
    [ValidateRange(1, 30)][int]$ActivityDays = 7,
    [ValidateRange(1, 200)][int]$ActivityMaxPages = 20,
    [switch]$IncludeExchange,
    [switch]$IncludeGraph,
    [switch]$DisableWAM,
    [switch]$NoWordReport,
    [string[]]$Frameworks = @('NCSC', 'NISTCSF', 'NIST80053', 'ISO27001'),
    [switch]$NoDisconnect
)

$ErrorActionPreference = 'Stop'

#region ======================= Settings and state ======================================

$Script:Settings = [ordered]@{
    OutputRoot        = $OutputRoot
    UserPrincipalName = $UserPrincipalName
    IncludeActivity   = [bool]$IncludeActivityExplorer
    ActivityDays      = $ActivityDays
    ActivityMaxPages  = $ActivityMaxPages
    IncludeExchange   = [bool]$IncludeExchange
    IncludeGraph      = [bool]$IncludeGraph
    DisableWAM        = [bool]$DisableWAM
    WordReport        = -not [bool]$NoWordReport
    Frameworks        = $Frameworks
}
$Script:RequiredModules = @(
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; MinVersion = [version]'3.10.1'; Required = $true;  Purpose = 'Security & Compliance and Exchange Online cmdlets' }
    [pscustomobject]@{ Name = 'ImportExcel';              MinVersion = [version]'7.0.0'; Required = $false; Purpose = 'Excel workbook output (optional)' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; MinVersion = [version]'2.41.0'; Required = $false; Purpose = 'Microsoft Graph data: licensing, DLP alerts, Entra roles, sharing (optional)' }
)
$Script:Areas = [ordered]@{
    '1'  = 'Tenant information'
    '2'  = 'DLP policies and workload coverage'
    '3'  = 'DLP rules - conditions, actions, policy tips and overrides'
    '4'  = 'Classifiers and confidence levels used in rules'
    '5'  = 'Sensitive information types (built-in and custom)'
    '6'  = 'Exact Data Match (EDM)'
    '7'  = 'Sensitivity labels and auto-labelling'
    '8'  = 'Endpoint DLP - USB, printing, browser and cloud uploads'
    '9'  = 'DLP alerts and alert policies'
    '10' = 'Operational ownership - role groups and members'
    '11' = 'Activity Explorer'
    '12' = 'Exchange Online mail flow rules and audit'
    '13' = 'Microsoft Graph: licensing, DLP alerts, Entra roles, sharing settings'
}

$Script:LogFile   = $null
$Script:OutputDir = $null
$Script:Datasets  = [ordered]@{}
$Script:Errors    = [System.Collections.Generic.List[object]]::new()
$Script:Cache     = @{}
$Script:ExcludeProps = @('PSComputerName', 'RunspaceId', 'PSShowComputerName', 'ObjectState', 'IsValid', 'OrganizationId', 'ExchangeVersion', 'ObjectClass')

#endregion

#region ======================= General helpers =========================================

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    if ($Script:LogFile) { Add-Content -Path $Script:LogFile -Value $line }
    $colour = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'OK' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $colour
}

function Read-YesNo {
    param([string]$Prompt, [bool]$Default = $false)
    $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
    $a = Read-Host "$Prompt $suffix"
    if ([string]::IsNullOrWhiteSpace($a)) { return $Default }
    return ($a.Trim() -match '^(y|yes)$')
}

function Wait-Key { [void](Read-Host 'Press Enter to return to the menu') }

function Start-Run {
    <# Creates a fresh output folder and resets per-run state (data cache, datasets, errors). #>
    param([string]$Label)
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $Script:OutputDir = Join-Path $Script:Settings.OutputRoot "M365-DLP-$Label-$ts"
    New-Item -Path $Script:OutputDir -ItemType Directory -Force | Out-Null
    $Script:LogFile  = Join-Path $Script:OutputDir 'Collection.log'
    $Script:Datasets = [ordered]@{}
    $Script:Errors   = [System.Collections.Generic.List[object]]::new()
    $Script:Cache    = @{}
    Write-Log "Output folder: $Script:OutputDir"
}

function Invoke-Step {
    <# Runs a collection step. Skips cleanly when a cmdlet is unavailable; logs failures and carries on. #>
    param([string]$Name, [scriptblock]$Action, [string[]]$Requires)
    Write-Log "Collecting: $Name"
    $missing = @($Requires | Where-Object { $_ -and -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        $msg = "Cmdlet(s) not available in this session: $($missing -join ', '). The signed-in account is missing the required Purview role, or the tenant is not licensed for this feature."
        Write-Log "  $Name skipped - $msg" 'WARN'
        $Script:Errors.Add([pscustomobject]@{ Step = $Name; Outcome = 'Skipped - cmdlet not available'; Error = $msg })
        return , @()
    }
    try {
        $result = @(& $Action)
        Write-Log ("  {0}: {1} item(s)" -f $Name, $result.Count) 'OK'
        return , $result
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        $msg = "$($_.Exception.Message.Split([Environment]::NewLine)[0]) Likely cause: missing Purview role or licence."
        Write-Log "  $Name skipped - $msg" 'WARN'
        $Script:Errors.Add([pscustomobject]@{ Step = $Name; Outcome = 'Skipped - cmdlet not available'; Error = $msg })
        return , @()
    }
    catch {
        $msg = $_.Exception.Message
        Write-Log "  $Name failed: $msg" 'ERROR'
        $Script:Errors.Add([pscustomobject]@{ Step = $Name; Outcome = 'Failed'; Error = $msg })
        return , @()
    }
}

function Get-Cached {
    <# Collects a dataset once per run and re-uses it (several reports share the same source data). Emits items unrolled. #>
    param([string]$Key, [scriptblock]$Builder)
    if (-not $Script:Cache.ContainsKey($Key)) {
        $r = & $Builder
        $Script:Cache[$Key] = @($r | Where-Object { $null -ne $_ })
    }
    $Script:Cache[$Key]
}

function Get-StepGap {
    param([string[]]$Steps)
    $hit = @($Script:Errors | Where-Object { $Steps -contains $_.Step })
    if ($hit.Count -eq 0) { return '' }
    return ' NOT COLLECTED: ' + (($hit | ForEach-Object { "$($_.Step) ($($_.Outcome))" }) -join '; ') + ' - check roles/licensing and re-run.'
}

function ConvertTo-CellValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { $s = $Value }
    elseif ($Value -is [ValueType]) { $s = [string]$Value }
    elseif ($Value -is [System.Collections.IDictionary]) { $s = ($Value | ConvertTo-Json -Depth 8 -Compress) }
    elseif ($Value -is [System.Collections.IEnumerable]) {
        $parts = foreach ($i in $Value) {
            if ($null -eq $i) { continue }
            if ($i -is [string] -or $i -is [ValueType]) { [string]$i } else { $i | ConvertTo-Json -Depth 8 -Compress }
        }
        $s = ($parts -join ' | ')
    }
    else {
        $s = [string]$Value
        if ($s -match '^System\.|^Microsoft\.') { $s = ($Value | ConvertTo-Json -Depth 6 -Compress) }
    }
    if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) + ' ...[truncated]' }
    return $s
}

function ConvertTo-FlatObject {
    param([object[]]$InputObject)
    foreach ($o in $InputObject) {
        if ($null -eq $o) { continue }
        $h = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) {
            if ($Script:ExcludeProps -contains $p.Name) { continue }
            $h[$p.Name] = ConvertTo-CellValue $p.Value
        }
        [pscustomobject]$h
    }
}

function Export-Dataset {
    param([string]$Name, [object[]]$Data, [switch]$AlreadyFlat)
    if ($Script:Datasets.Contains($Name)) { return }   # already written in this run
    $Data = @($Data | Where-Object { $null -ne $_ })
    $jsonPath = Join-Path $Script:OutputDir "$Name.json"
    $csvPath  = Join-Path $Script:OutputDir "$Name.csv"
    if ($Data.Count -eq 0) {
        '[]' | Set-Content -Path $jsonPath -Encoding UTF8
        Write-Log "  $Name : no data" 'WARN'
        $Script:Datasets[$Name] = @()
        return
    }
    try { $Data | ConvertTo-Json -Depth 12 | Set-Content -Path $jsonPath -Encoding UTF8 }
    catch { Write-Log "  JSON export for $Name failed: $($_.Exception.Message)" 'WARN' }
    $flat = if ($AlreadyFlat) { $Data } else { @(ConvertTo-FlatObject $Data) }
    $flat | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    $Script:Datasets[$Name] = $flat
}

function Get-Val {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [System.Collections.IDictionary]) {
        foreach ($k in $Obj.Keys) { if ([string]$k -ieq $Name) { return $Obj[$k] } }
        return $null
    }
    $p = $Obj.PSObject.Properties | Where-Object { $_.Name -ieq $Name } | Select-Object -First 1
    if ($p) { return $p.Value }
    return $null
}

function Get-Members {
    param($Obj)
    if ($null -eq $Obj -or $Obj -is [string] -or $Obj -is [ValueType]) { return $null }
    if ($Obj -is [System.Collections.IDictionary]) {
        return @(foreach ($k in $Obj.Keys) { [pscustomobject]@{ Name = [string]$k; Value = $Obj[$k] } })
    }
    if ($Obj -is [System.Collections.IEnumerable]) { return $null }
    return @(foreach ($p in $Obj.PSObject.Properties) {
            if ($p.MemberType -in 'NoteProperty', 'Property') { [pscustomobject]@{ Name = $p.Name; Value = $p.Value } }
        })
}

function Find-RuleClassifiers {
    <# Walks rule conditions / AdvancedRule JSON and emits each SIT, EDM, trainable classifier and label referenced. #>
    param($Node, [string]$Policy, [string]$Rule, [string]$Source, [string]$GroupName = '', [int]$Depth = 0)
    if ($null -eq $Node -or $Depth -gt 15) { return }
    if ($Node -is [string]) {
        $t = $Node.Trim()
        if ($t.StartsWith('{') -or $t.StartsWith('[')) { try { $Node = $t | ConvertFrom-Json } catch { return } }
        else { return }
    }
    if ($Node -is [System.Collections.IEnumerable] -and $Node -isnot [System.Collections.IDictionary]) {
        foreach ($i in $Node) { Find-RuleClassifiers $i $Policy $Rule $Source $GroupName ($Depth + 1) }
        return
    }
    $members = Get-Members $Node
    if (-not $members) { return }
    $thisGroup = Get-Val $Node 'name'
    if (-not $thisGroup) { $thisGroup = $GroupName }
    foreach ($m in $members) {
        if ($m.Name -ieq 'sensitivetypes') {
            foreach ($st in @($m.Value)) {
                if ($null -eq $st) { continue }
                [pscustomobject]@{
                    Policy = $Policy; Rule = $Rule; Source = $Source; Group = $GroupName
                    ReferenceType   = 'Classifier'
                    Name            = Get-Val $st 'name'
                    Id              = Get-Val $st 'id'
                    ClassifierType  = $(if (Get-Val $st 'classifiertype') { Get-Val $st 'classifiertype' } else { 'Content (SIT)' })
                    ConfidenceLevel = Get-Val $st 'confidencelevel'
                    MinConfidence   = Get-Val $st 'minconfidence'
                    MaxConfidence   = Get-Val $st 'maxconfidence'
                    MinCount        = Get-Val $st 'mincount'
                    MaxCount        = Get-Val $st 'maxcount'
                }
            }
        }
        elseif ($m.Name -ieq 'labels') {
            foreach ($lb in @($m.Value)) {
                if ($null -eq $lb) { continue }
                [pscustomobject]@{
                    Policy = $Policy; Rule = $Rule; Source = $Source; Group = $GroupName
                    ReferenceType = 'SensitivityLabel'
                    Name = Get-Val $lb 'name'; Id = Get-Val $lb 'id'; ClassifierType = Get-Val $lb 'type'
                    ConfidenceLevel = $null; MinConfidence = $null; MaxConfidence = $null; MinCount = $null; MaxCount = $null
                }
            }
        }
        else { Find-RuleClassifiers $m.Value $Policy $Rule $Source ([string]$thisGroup) ($Depth + 1) }
    }
}

function Count-Where { param($Data, [scriptblock]$Filter) @($Data | Where-Object $Filter).Count }

#endregion

#region ======================= Modules and connection ==================================

function Get-ModuleStatus {
    foreach ($m in $Script:RequiredModules) {
        $installed = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
        $loaded    = Get-Module -Name $m.Name | Select-Object -First 1
        [pscustomobject]@{
            Module    = $m.Name
            Required  = $(if ($m.Required) { 'Required' } else { 'Optional' })
            Minimum   = $m.MinVersion.ToString()
            Installed = $(if ($installed) { $installed.Version.ToString() } else { 'Not installed' })
            Loaded    = $(if ($loaded) { $loaded.Version.ToString() } else { '' })
            Status    = $(if (-not $installed) { 'MISSING' } elseif ($installed.Version -lt $m.MinVersion) { 'OUTDATED' } else { 'OK' })
            Purpose   = $m.Purpose
        }
    }
}

function Show-ModuleStatus {
    $status = @(Get-ModuleStatus)
    Write-Host ''
    Write-Host ('PowerShell {0} ({1})   Execution policy (effective): {2}' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, (Get-ExecutionPolicy)) -ForegroundColor Cyan
    foreach ($s in $status) {
        $c = switch ($s.Status) { 'OK' { 'Green' } 'OUTDATED' { 'Yellow' } default { if ($s.Required -eq 'Required') { 'Red' } else { 'Yellow' } } }
        Write-Host ('  {0,-26} {1,-9} min {2,-7} installed {3,-14} {4}' -f $s.Module, $s.Required, $s.Minimum, $s.Installed, $s.Status) -ForegroundColor $c
    }
    return , $status
}

function Install-RequiredModules {
    Write-Host ''
    Write-Host '=== Install / update PowerShell modules ===' -ForegroundColor Cyan
    $status = Show-ModuleStatus
    $todo = @($status | Where-Object { $_.Status -ne 'OK' })
    if ($todo.Count -eq 0 -and @(Get-Module -ListAvailable -Name ExchangeOnlineManagement).Count -gt 1) {
        $todo = @($status | Where-Object Module -eq 'ExchangeOnlineManagement')
    }
    if ($todo.Count -eq 0) {
        Write-Host ''
        Write-Host 'All modules are installed and up to date.' -ForegroundColor Green
        if (-not (Read-YesNo 'Check the PowerShell Gallery for newer versions anyway?' $false)) { return }
        $todo = $status
    }
    Write-Host ''
    Write-Host 'Modules are installed for the current user only (-Scope CurrentUser); administrator rights are not needed.' -ForegroundColor Gray
    if (-not (Read-YesNo ('Install / update: ' + (($todo | ForEach-Object Module) -join ', ') + '?') $true)) { return }

    try {
        if ($PSVersionTable.PSEdition -ne 'Core') {
            # Windows PowerShell 5.1: TLS 1.2 and the NuGet provider are needed to reach the PowerShell Gallery
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $nuget = Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' }
            if (-not $nuget) {
                Write-Host 'Installing NuGet package provider...' -ForegroundColor Gray
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
            }
        }
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
            Write-Host 'Note: PSGallery is untrusted on this machine; installing with -Force for these modules only.' -ForegroundColor Gray
        }
    }
    catch { Write-Host "Gallery preparation warning: $($_.Exception.Message)" -ForegroundColor Yellow }

    $eomVersions = @(Get-Module -ListAvailable -Name ExchangeOnlineManagement)
    $clean = $false
    if ($eomVersions.Count -gt 1) {
        Write-Host ''
        Write-Host "$($eomVersions.Count) versions of ExchangeOnlineManagement are installed ($((($eomVersions | ForEach-Object Version) -join ', '))). Multiple versions commonly cause sign-in errors." -ForegroundColor Yellow
        $clean = Read-YesNo 'Do a clean reinstall (remove every version, then install the latest)?' $true
    }
    elseif ($eomVersions.Count -eq 1) {
        $clean = Read-YesNo 'Do a clean reinstall of ExchangeOnlineManagement (only needed if you have connection errors)?' $false
    }
    if ($clean) {
        if (Get-Module ExchangeOnlineManagement) {
            Write-Host 'The module is loaded in this window. It will be removed from the session first; restart PowerShell afterwards.' -ForegroundColor Gray
            Remove-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
        }
        try {
            Uninstall-Module -Name ExchangeOnlineManagement -AllVersions -Force -ErrorAction Stop
            Write-Host '  Removed all versions installed through PowerShellGet.' -ForegroundColor Green
        }
        catch { Write-Host "  Uninstall-Module reported: $($_.Exception.Message)" -ForegroundColor Yellow }
        $left = @(Get-Module -ListAvailable -Name ExchangeOnlineManagement)
        if ($left.Count) {
            Write-Host '  These copies could not be removed automatically (installed for all users or outside PowerShellGet). Delete these folders manually, ideally from an elevated PowerShell:' -ForegroundColor Yellow
            $left | ForEach-Object { Write-Host "    $(Split-Path $_.ModuleBase -Parent)" -ForegroundColor Yellow }
        }
        if (-not ($todo | Where-Object Module -eq 'ExchangeOnlineManagement')) {
            $todo = @($todo) + @($status | Where-Object Module -eq 'ExchangeOnlineManagement')
        }
    }

    foreach ($t in $todo) {
        $def = $Script:RequiredModules | Where-Object Name -eq $t.Module
        try {
            Write-Host "Installing $($t.Module) (minimum $($def.MinVersion))..." -ForegroundColor Gray
            Install-Module -Name $t.Module -Scope CurrentUser -MinimumVersion $def.MinVersion -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Host "  $($t.Module) installed." -ForegroundColor Green
        }
        catch { Write-Host "  $($t.Module) failed: $($_.Exception.Message)" -ForegroundColor Red }
    }
    Write-Host ''
    [void](Show-ModuleStatus)
    if ($clean) { Write-Host 'Clean reinstall complete. Close this PowerShell window and start the script in a new one before connecting.' -ForegroundColor Yellow }
    $eom = Get-Module ExchangeOnlineManagement
    $eomBest = Get-Module -ListAvailable ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if ($eom -and $eomBest -and $eom.Version -lt $eomBest.Version) {
        Write-Host ''
        Write-Host "An older ExchangeOnlineManagement ($($eom.Version)) is loaded in this window. Close PowerShell and start the script again to use $($eomBest.Version)." -ForegroundColor Yellow
    }
}

function Assert-Modules {
    $eom = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $eom) { Write-Host 'ExchangeOnlineManagement is not installed. Use menu option 1 (or -Mode InstallModules) first.' -ForegroundColor Red; return $false }
    if ($eom.Version -lt [version]'3.10.1') { Write-Host "ExchangeOnlineManagement $($eom.Version) is too old (3.10.1+ required; Microsoft requires 3.10.1 or newer for Exchange Online PowerShell authentication). Use menu option 1 to update." -ForegroundColor Red; return $false }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    return $true
}

function Get-IppsConnection {
    if (-not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) { return $null }
    Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.ConnectionUri -match 'compliance|protection\.outlook\.com/powershell' -and $_.IsEopSession -and $_.State -eq 'Connected' } |
        Select-Object -First 1
}

function Get-ConnectArgs {
    param([string]$Command, [switch]$ForceDisableWAM)
    $params = (Get-Command $Command -ErrorAction SilentlyContinue).Parameters
    $a = @{}
    if ($params -and $params.ContainsKey('ShowBanner')) { $a['ShowBanner'] = $false }
    if ($Script:Settings.UserPrincipalName) { $a['UserPrincipalName'] = $Script:Settings.UserPrincipalName }
    if (($Script:Settings.DisableWAM -or $ForceDisableWAM) -and $params -and $params.ContainsKey('DisableWAM')) { $a['DisableWAM'] = $true }
    return $a
}

function Invoke-ExoConnect {
    <#
      Calls Connect-IPPSSession / Connect-ExchangeOnline. If the first attempt fails and the module supports
      -DisableWAM, retries once without the Windows Web Account Manager sign-in broker, which avoids a common
      class of sign-in failures (including "You cannot call a method on a null-valued expression").
    #>
    param([string]$Command)
    $a = Get-ConnectArgs -Command $Command
    try {
        & $Command @a -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        $Script:LastConnectError = $_
        Write-Log "$Command failed: $($_.Exception.Message)" 'ERROR'
        if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) { Write-Log ("  At: " + ($_.InvocationInfo.PositionMessage -replace '\s+', ' ')) 'INFO' }
        if ($_.ScriptStackTrace) { Write-Log ("  Stack: " + ((@($_.ScriptStackTrace -split "`n" | Select-Object -First 3)) -join ' <- ')) 'INFO' }
        $params = (Get-Command $Command -ErrorAction SilentlyContinue).Parameters
        if (-not $a.ContainsKey('DisableWAM') -and $params -and $params.ContainsKey('DisableWAM')) {
            Write-Log "Retrying $Command with -DisableWAM (browser-based sign-in instead of the Windows account broker)..." 'WARN'
            $a2 = Get-ConnectArgs -Command $Command -ForceDisableWAM
            try {
                & $Command @a2 -ErrorAction Stop | Out-Null
                Write-Log 'Connected with -DisableWAM. This setting will be used for the rest of the session (Settings > 7).' 'OK'
                $Script:Settings.DisableWAM = $true
                return $true
            }
            catch {
                $Script:LastConnectError = $_
                Write-Log "$Command -DisableWAM also failed: $($_.Exception.Message)" 'ERROR'
            }
        }
        return $false
    }
}

function Show-ConnectionDiagnostics {
    <# Explains the most common causes of connection failures on this workstation. #>
    param([string]$Command)
    Write-Host ''
    Write-Host '--- Connection diagnostics ---' -ForegroundColor Cyan
    $issues = 0
    $all = @(Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending)
    $loaded = Get-Module -Name ExchangeOnlineManagement
    Write-Host ("  ExchangeOnlineManagement installed: {0}" -f (($all | ForEach-Object { "$($_.Version) ($($_.ModuleBase))" }) -join '; '))
    Write-Host ("  ExchangeOnlineManagement loaded   : {0}" -f $(if ($loaded) { ($loaded | ForEach-Object { $_.Version }) -join ', ' } else { 'none' }))
    if ($all.Count -gt 1) {
        $issues++
        Write-Host '  ! More than one version is installed. This is the most common cause of "null-valued expression" errors.' -ForegroundColor Yellow
        Write-Host '    Fix: menu option 1 > clean reinstall (removes every version and installs the latest), then restart PowerShell.' -ForegroundColor Yellow
    }
    if (@($loaded).Count -gt 0 -and $all.Count -gt 0 -and (@($loaded)[0].Version -lt $all[0].Version)) {
        $issues++
        Write-Host "  ! An older version ($(@($loaded)[0].Version)) is loaded in this window. Close PowerShell and start a new window." -ForegroundColor Yellow
    }
    $conflicts = @(Get-Module | Where-Object { $_.Name -match '^(Microsoft\.Graph|Az\.Accounts|AzureAD|MSOnline|MicrosoftTeams|PnP\.PowerShell|Microsoft\.Online\.SharePoint)' } | Select-Object -ExpandProperty Name -Unique)
    if ($conflicts.Count) {
        $issues++
        Write-Host "  ! Other Microsoft 365 modules are loaded in this window: $($conflicts -join ', ')." -ForegroundColor Yellow
        Write-Host '    They can load a different version of the sign-in library (MSAL). Start a fresh PowerShell window and run only this script.' -ForegroundColor Yellow
    }
    if ($Host.Name -match 'ISE') {
        $issues++
        Write-Host '  ! Running in PowerShell ISE, which is not supported for interactive sign-in with this module. Use Windows PowerShell or PowerShell 7 in a console window.' -ForegroundColor Yellow
    }
    $params = (Get-Command $Command -ErrorAction SilentlyContinue).Parameters
    if ($params -and -not $params.ContainsKey('DisableWAM')) {
        Write-Host '  i This module version does not support -DisableWAM. Updating to the latest ExchangeOnlineManagement (menu option 1) adds it.' -ForegroundColor Gray
    }
    if ($issues -eq 0) {
        Write-Host '  No local workstation issues detected. Other causes: Conditional Access blocking the sign-in, the sign-in window closed or cancelled,' -ForegroundColor Gray
        Write-Host '  a proxy blocking login.microsoftonline.com, or a temporary service issue. Try the command on its own to see the raw error:' -ForegroundColor Gray
        Write-Host "    $Command -UserPrincipalName <admin UPN> -DisableWAM" -ForegroundColor Gray
    }
    Write-Host '  Full details are in Collection.log (when a report run is active).' -ForegroundColor Gray
}

function Connect-Tenant {
    <# Connects to Security & Compliance PowerShell, re-using a live connection unless -Force is used. #>
    param([switch]$Force, [switch]$NoPrompt)
    if (-not (Assert-Modules)) { return $false }
    $existing = Get-IppsConnection
    if ($existing -and -not $Force) {
        Write-Log "Using existing connection: $($existing.UserPrincipalName) ($($existing.TenantID))"
        return $true
    }
    if ($existing -and $Force) {
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
    if (-not $Script:Settings.UserPrincipalName -and $Mode -eq 'Menu' -and -not $NoPrompt) {
        $u = Read-Host 'Admin account to sign in with (Enter to choose in the sign-in window)'
        if ($u) { $Script:Settings.UserPrincipalName = $u.Trim() }
    }
    Write-Log 'Connecting to Security & Compliance PowerShell (Connect-IPPSSession)...'
    $ok = Invoke-ExoConnect -Command 'Connect-IPPSSession'
    if (-not $ok) { Show-ConnectionDiagnostics -Command 'Connect-IPPSSession'; return $false }
    if (-not (Get-Command Get-DlpCompliancePolicy -ErrorAction SilentlyContinue)) {
        Write-Log 'Connected, but no DLP cmdlets were loaded. The account has no Purview DLP role. Run menu option 3 (Test credentials) for details.' 'WARN'
    }
    $c = Get-IppsConnection
    if ($c) { Write-Log "Connected as $($c.UserPrincipalName) to tenant $($c.TenantID)" 'OK' }
    return $true
}

function Invoke-ConnectMenu {
    <# Menu option 2: connect (or reconnect with a different account) before running anything else. #>
    Write-Host ''
    Write-Host '=== Connect to tenant ===' -ForegroundColor Cyan
    $existing = Get-IppsConnection
    $force = $false
    if ($existing) {
        Write-Host "Already connected as $($existing.UserPrincipalName) to tenant $($existing.TenantID)." -ForegroundColor Green
        if (-not (Read-YesNo 'Disconnect and connect with a different account or tenant?' $false)) { return }
        $force = $true
        $Script:Settings.UserPrincipalName = ''
    }
    $prompt = if ($Script:Settings.UserPrincipalName) { "Admin account [$($Script:Settings.UserPrincipalName)] (Enter to keep, '-' to choose in the sign-in window)" } else { 'Admin account, e.g. admin@contoso.onmicrosoft.com (Enter to choose in the sign-in window)' }
    $u = Read-Host $prompt
    if ($u -eq '-') { $Script:Settings.UserPrincipalName = '' }
    elseif ($u) { $Script:Settings.UserPrincipalName = $u.Trim() }

    if (-not (Connect-Tenant -Force:$force -NoPrompt)) { return }

    $c = Get-IppsConnection
    $dlpCmdlets = @('Get-DlpCompliancePolicy', 'Get-DlpComplianceRule', 'Get-DlpSensitiveInformationType', 'Get-DlpEdmSchema', 'Get-Label', 'Get-LabelPolicy',
        'Get-AutoSensitivityLabelPolicy', 'Get-PolicyConfig', 'Get-ProtectionAlert', 'Get-RoleGroup', 'Export-ActivityExplorerData')
    $loaded = @($dlpCmdlets | Where-Object { Get-Command $_ -ErrorAction SilentlyContinue })
    Write-Host ''
    Write-Host 'Connection details' -ForegroundColor Cyan
    Write-Host ("  Account      : {0}" -f $(if ($c) { $c.UserPrincipalName } else { 'unknown' }))
    Write-Host ("  Tenant ID    : {0}" -f $(if ($c) { $c.TenantID } else { 'unknown' }))
    Write-Host ("  Organisation : {0}" -f $(if ($c) { $c.Organization } else { '' }))
    Write-Host ("  Endpoint     : {0}" -f $(if ($c) { $c.ConnectionUri } else { '' }))
    $col = if ($loaded.Count -eq $dlpCmdlets.Count) { 'Green' } elseif ($loaded.Count -gt 0) { 'Yellow' } else { 'Red' }
    Write-Host ("  DLP cmdlets  : {0} of {1} key cmdlets available to this account" -f $loaded.Count, $dlpCmdlets.Count) -ForegroundColor $col
    if ($loaded.Count -lt $dlpCmdlets.Count) {
        Write-Host ("  Not available: {0}" -f (($dlpCmdlets | Where-Object { $loaded -notcontains $_ }) -join ', ')) -ForegroundColor Yellow
        Write-Host '  Use option 3 (Test credentials and tenant access) for a full check.' -ForegroundColor Yellow
    }
}

function Connect-Exo {
    Write-Log 'Connecting to Exchange Online (Connect-ExchangeOnline)...'
    if (-not (Invoke-ExoConnect -Command 'Connect-ExchangeOnline')) {
        Show-ConnectionDiagnostics -Command 'Connect-ExchangeOnline'
        throw "Connect-ExchangeOnline failed: $($Script:LastConnectError.Exception.Message)"
    }
}

function Reset-AfterExo {
    <# Exchange Online and Security & Compliance share some cmdlet names (e.g. Get-RoleGroup). Disconnect so the
       next action reconnects cleanly to Security & Compliance. #>
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
}

function Disconnect-Tenant {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null; Write-Host 'Disconnected.' -ForegroundColor Green } catch { }
}

#endregion

#region ======================= Source data (collected once per run) ====================

function Get-Policies          { Get-Cached 'Policies'      { Invoke-Step 'DLP compliance policies' -Requires Get-DlpCompliancePolicy -Action { Get-DlpCompliancePolicy } } }
function Get-Rules             { Get-Cached 'Rules'         { Invoke-Step 'DLP compliance rules' -Requires Get-DlpComplianceRule -Action { Get-DlpComplianceRule } } }
function Get-SITs              { Get-Cached 'SITs'          { Invoke-Step 'Sensitive information types' -Requires Get-DlpSensitiveInformationType -Action { Get-DlpSensitiveInformationType } } }
function Get-RulePackages      { Get-Cached 'RulePackages'  { Invoke-Step 'SIT rule packages' -Requires Get-DlpSensitiveInformationTypeRulePackage -Action { Get-DlpSensitiveInformationTypeRulePackage } } }
function Get-Dictionaries      { Get-Cached 'Dictionaries'  { Invoke-Step 'Keyword dictionaries' -Requires Get-DlpKeywordDictionary -Action { Get-DlpKeywordDictionary } } }
function Get-EdmSchemas        { Get-Cached 'EdmSchemas'    { Invoke-Step 'Exact Data Match schemas' -Requires Get-DlpEdmSchema -Action { Get-DlpEdmSchema } } }
function Get-Labels            { Get-Cached 'Labels'        { Invoke-Step 'Sensitivity labels' -Requires Get-Label -Action { Get-Label } } }
function Get-LabelPolicies     { Get-Cached 'LabelPolicies' { Invoke-Step 'Sensitivity label policies' -Requires Get-LabelPolicy -Action { Get-LabelPolicy } } }
function Get-AutoLabelPolicies { Get-Cached 'AutoLabelPol'  { Invoke-Step 'Auto-labelling policies' -Requires Get-AutoSensitivityLabelPolicy -Action { Get-AutoSensitivityLabelPolicy } } }
function Get-AutoLabelRules    { Get-Cached 'AutoLabelRule' { Invoke-Step 'Auto-labelling rules' -Requires Get-AutoSensitivityLabelRule -Action { Get-AutoSensitivityLabelRule } } }
function Get-PolicyConfigData  { Get-Cached 'PolicyConfig'  { Invoke-Step 'Endpoint DLP / tenant policy configuration' -Requires Get-PolicyConfig -Action { Get-PolicyConfig } } }
function Get-AlertPolicies     { Get-Cached 'AlertPolicies' { Invoke-Step 'Alert policies' -Requires Get-ProtectionAlert -Action { Get-ProtectionAlert } } }
function Get-RoleGroups        { Get-Cached 'RoleGroups'    { Invoke-Step 'Purview role groups' -Requires Get-RoleGroup -Action { Get-RoleGroup } } }

function Get-RoleGroupMembers {
    Get-Cached 'RoleGroupMembers' {
        $groups = @(Get-RoleGroups)
        Invoke-Step 'Purview role group members' -Requires Get-RoleGroupMember -Action {
            foreach ($rg in $groups) {
                $members = @()
                try { $members = @(Get-RoleGroupMember -Identity ([string](Get-Val $rg 'Identity')) -ErrorAction Stop) } catch { }
                if ($members.Count -eq 0) { [pscustomobject]@{ RoleGroup = $rg.Name; Roles = ConvertTo-CellValue (Get-Val $rg 'Roles'); Member = ''; MemberType = '' } }
                foreach ($m in $members) {
                    [pscustomobject]@{
                        RoleGroup  = $rg.Name
                        Roles      = ConvertTo-CellValue (Get-Val $rg 'Roles')
                        Member     = $(if (Get-Val $m 'PrimarySmtpAddress') { Get-Val $m 'PrimarySmtpAddress' } else { Get-Val $m 'Name' })
                        MemberType = ConvertTo-CellValue (Get-Val $m 'RecipientType')
                    }
                }
            }
        }
    }
}

$Script:WorkloadPatterns = [ordered]@{
    'Exchange' = 'Exchange'; 'SharePoint' = 'SharePoint'; 'OneDrive' = 'OneDrive'; 'Teams' = 'Teams'; 'Endpoint' = 'Endpoint'
    'Copilot' = 'Copilot|Applications'; 'OnPremises' = 'OnPremises'; 'PowerBI' = 'PowerBI'; 'CloudApps' = 'ThirdParty|CloudApp|MicrosoftDefenderForCloudApps'
}
$Script:RuleActionProps = @(
    'Disabled', 'Mode', 'Priority', 'ReportSeverityLevel', 'BlockAccess', 'BlockAccessScope', 'AccessScope',
    'NotifyUser', 'NotifyUserType', 'NotifyPolicyTipCustomText', 'NotifyPolicyTipDisplayOption', 'NotifyEmailCustomSubject', 'NotifyEmailCustomText',
    'NotifyAllowOverride', 'NotifyOverrideRequirements', 'GenerateAlert', 'AlertProperties', 'GenerateIncidentReport', 'IncidentReportContent',
    'EndpointDlpRestrictions', 'RestrictBrowserAccess', 'EnforcementPlanes', 'EncryptRMSTemplate', 'ApplyHtmlMessageBody', 'Quarantine',
    'RedirectMessageTo', 'ModifySubject', 'StopPolicyProcessing', 'Comment', 'WhenChangedUTC'
)

function Get-PolicySummary {
    Get-Cached 'PolicySummary' {
        foreach ($p in @(Get-Policies)) {
            $h = [ordered]@{
                Name = $p.Name; Mode = $p.Mode; Enabled = Get-Val $p 'Enabled'
                Workload = ConvertTo-CellValue (Get-Val $p 'Workload'); Priority = Get-Val $p 'Priority'; Type = Get-Val $p 'Type'
                DistributionStatus = ConvertTo-CellValue (Get-Val $p 'DistributionStatus'); Comment = Get-Val $p 'Comment'
                CreatedBy = Get-Val $p 'CreatedBy'; WhenCreatedUTC = Get-Val $p 'WhenCreatedUTC'
                LastModifiedBy = Get-Val $p 'LastModifiedBy'; WhenChangedUTC = Get-Val $p 'WhenChangedUTC'
            }
            foreach ($prop in $p.PSObject.Properties) {
                if ($prop.Name -match 'Location|Scope|Simulation|AdminUnit|Enforcement') { if (-not $h.Contains($prop.Name)) { $h[$prop.Name] = ConvertTo-CellValue $prop.Value } }
            }
            [pscustomobject]$h
        }
    }
}

function Get-Coverage {
    Get-Cached 'Coverage' {
        foreach ($p in @(Get-Policies)) {
            $text = ((ConvertTo-CellValue (Get-Val $p 'Workload')) + ' ' + (ConvertTo-CellValue (Get-Val $p 'Locations')))
            $h = [ordered]@{ Policy = $p.Name; Mode = $p.Mode }
            foreach ($k in $Script:WorkloadPatterns.Keys) { $h[$k] = [bool]($text -match $Script:WorkloadPatterns[$k]) }
            [pscustomobject]$h
        }
    }
}

function Get-RuleSummary {
    Get-Cached 'RuleSummary' {
        foreach ($r in @(Get-Rules)) {
            $h = [ordered]@{ Policy = Get-Val $r 'ParentPolicyName'; Rule = $r.Name }
            foreach ($a in $Script:RuleActionProps) { $h[$a] = ConvertTo-CellValue (Get-Val $r $a) }
            $conds = foreach ($prop in $r.PSObject.Properties) {
                if ($Script:RuleActionProps -contains $prop.Name) { continue }
                if ($prop.Name -match '^(ContentContains|ExceptIf|Sent|From|Recipient|Document|Content|Subject|Message|Access|Has|Header|Anyof|Processing|With|Unscannable|Evaluate|Advanced)') {
                    $v = ConvertTo-CellValue $prop.Value
                    if ($v -and $v -ne 'False' -and $v -ne '0') { "$($prop.Name)=$v" }
                }
            }
            $h['ConditionsAndExceptions'] = ConvertTo-CellValue (@($conds) -join ' ;; ')
            $h['UsesAdvancedRule'] = [bool](Get-Val $r 'AdvancedRule')
            $block = ConvertTo-CellValue (Get-Val $r 'BlockAccess')
            $override = ConvertTo-CellValue (Get-Val $r 'NotifyAllowOverride')
            $h['ResponseInterpretation'] = if ($block -eq 'True' -and $override) { 'Block with override' }
            elseif ($block -eq 'True') { 'Block' }
            elseif (Get-Val $r 'NotifyUser') { 'Notify / policy tip' }
            elseif (ConvertTo-CellValue (Get-Val $r 'EndpointDlpRestrictions')) { 'Endpoint restrictions (see EndpointDlpRestrictions)' }
            else { 'Audit only' }
            [pscustomobject]$h
        }
    }
}

function Get-RuleClassifiers {
    Get-Cached 'RuleClassifiers' {
        foreach ($r in @(Get-Rules)) {
            $pol = [string](Get-Val $r 'ParentPolicyName')
            Find-RuleClassifiers (Get-Val $r 'ContentContainsSensitiveInformation') $pol $r.Name 'Condition'
            Find-RuleClassifiers (Get-Val $r 'ExceptIfContentContainsSensitiveInformation') $pol $r.Name 'Exception'
            Find-RuleClassifiers (Get-Val $r 'AdvancedRule') $pol $r.Name 'AdvancedRule'
        }
    }
}

function Get-EndpointRuleActions {
    Get-Cached 'EndpointRuleActions' {
        foreach ($r in @(Get-Rules)) {
            foreach ($er in @(Get-Val $r 'EndpointDlpRestrictions')) {
                if ($null -eq $er) { continue }
                [pscustomobject]@{
                    Policy = Get-Val $r 'ParentPolicyName'; Rule = $r.Name
                    Setting = Get-Val $er 'Setting'; Value = ConvertTo-CellValue (Get-Val $er 'Value'); Raw = ConvertTo-CellValue $er
                }
            }
        }
    }
}

function Get-SITView {
    Get-Cached 'SITView' {
        $refs = @(Get-RuleClassifiers)
        foreach ($s in @(Get-SITs)) {
            $pub = [string](Get-Val $s 'Publisher'); $id = Get-Val $s 'Id'
            [pscustomobject]@{
                Name = $s.Name; Id = $id; Publisher = $pub
                IsCustom = [bool]($pub -and $pub -notmatch '^Microsoft Corporation$')
                Type = Get-Val $s 'Type'; Classifier = ConvertTo-CellValue (Get-Val $s 'ClassifierType')
                RulePackId = ConvertTo-CellValue (Get-Val $s 'RulePackId'); RecommendedConfidence = Get-Val $s 'RecommendedConfidence'
                State = Get-Val $s 'State'; Description = Get-Val $s 'Description'
                UsedInDlpRules = [bool](@($refs | Where-Object { $_.Id -eq $id -or $_.Name -eq $s.Name }).Count)
            }
        }
    }
}

function Get-LabelView {
    Get-Cached 'LabelView' {
        $refs = @(Get-RuleClassifiers | Where-Object ReferenceType -eq 'SensitivityLabel')
        foreach ($l in @(Get-Labels)) {
            $g = ConvertTo-CellValue (Get-Val $l 'Guid')
            [pscustomobject]@{
                DisplayName = Get-Val $l 'DisplayName'; Name = $l.Name; Guid = $g; ParentId = ConvertTo-CellValue (Get-Val $l 'ParentId')
                Priority = Get-Val $l 'Priority'; ContentType = ConvertTo-CellValue (Get-Val $l 'ContentType'); Disabled = Get-Val $l 'Disabled'
                Tooltip = Get-Val $l 'Tooltip'; EncryptionEnabled = [bool]((ConvertTo-CellValue (Get-Val $l 'LabelActions')) -match 'encrypt')
                UsedInDlpRules = [bool](@($refs | Where-Object { $_.Id -eq $g -or $_.Name -eq $l.Name }).Count)
                WhenChangedUTC = Get-Val $l 'WhenChangedUTC'
            }
        }
    }
}

function Get-EndpointSettings {
    Get-Cached 'EndpointSettings' {
        foreach ($pc in @(Get-PolicyConfigData)) {
            foreach ($prop in $pc.PSObject.Properties) {
                if ($prop.Name -notmatch 'Endpoint|Dlp|Printer|RemovableMedia|NetworkShare|AppGroup|Browser|Domain|Bluetooth|Vpn') { continue }
                $items = @($prop.Value)
                if ($items.Count -eq 0 -or ($items.Count -eq 1 -and $null -eq $items[0])) { [pscustomobject]@{ Area = $prop.Name; Setting = ''; Value = '' }; continue }
                foreach ($i in $items) {
                    $setting = Get-Val $i 'Setting'
                    if ($setting) { [pscustomobject]@{ Area = $prop.Name; Setting = $setting; Value = ConvertTo-CellValue (Get-Val $i 'Value') } }
                    else { [pscustomobject]@{ Area = $prop.Name; Setting = ''; Value = ConvertTo-CellValue $i } }
                }
            }
        }
    }
}

function Get-DlpAlertPolicies {
    Get-Cached 'DlpAlertPolicies' {
        Get-AlertPolicies | Where-Object {
            (ConvertTo-CellValue (Get-Val $_ 'Category')) -match 'DataLoss|DLP' -or $_.Name -match 'DLP|Data loss' -or (ConvertTo-CellValue (Get-Val $_ 'Operation')) -match 'Dlp'
        }
    }
}

function Get-DlpRoleMembers {
    Get-Cached 'DlpRoleMembers' {
        Get-RoleGroupMembers | Where-Object { $_.Member -and ($_.Roles -match 'DLP|Data Loss|Compliance Admin|Information Protection|Alert|Case Management|Activity Explorer|Data Classification' -or $_.RoleGroup -match 'Compliance|Information Protection|Security|Organization Management') }
    }
}

function Get-ActivityRecords {
    Get-Cached 'Activity' {
        $days = $Script:Settings.ActivityDays; $maxPages = $Script:Settings.ActivityMaxPages
        Invoke-Step "Activity Explorer ($days days)" -Requires Export-ActivityExplorerData -Action {
            $end = (Get-Date).ToUniversalTime(); $start = $end.AddDays(-$days)
            $all = [System.Collections.Generic.List[object]]::new()
            $page = Export-ActivityExplorerData -StartTime $start -EndTime $end -OutputFormat Json -PageSize 5000
            $n = 1
            while ($true) {
                if ($page.ResultData) { foreach ($rec in @($page.ResultData | ConvertFrom-Json)) { $all.Add($rec) } }
                if ($page.LastPage -or -not $page.WaterMark -or $n -ge $maxPages) { break }
                $page = Export-ActivityExplorerData -StartTime $start -EndTime $end -OutputFormat Json -PageSize 5000 -PageCookie $page.WaterMark
                $n++
            }
            if ($n -ge $maxPages -and -not $page.LastPage) { Write-Log "  Activity Explorer page cap ($maxPages) reached; data is partial." 'WARN' }
            $all
        }
    }
}

#endregion

#region ======================= Area reports ============================================

function Export-AreaTenant {
    $c = Get-IppsConnection
    $info = [pscustomobject]@{
        CollectedAtUtc = (Get-Date).ToUniversalTime().ToString('u')
        CollectedBy    = $(if ($c) { $c.UserPrincipalName } else { $Script:Settings.UserPrincipalName })
        TenantId       = $(if ($c) { $c.TenantID } else { '' })
        Organization   = $(if ($c) { $c.Organization } else { '' })
        ConnectionUri  = $(if ($c) { $c.ConnectionUri } else { '' })
        ModuleVersion  = (Get-Module ExchangeOnlineManagement | Select-Object -First 1).Version.ToString()
        PowerShell     = $PSVersionTable.PSVersion.ToString()
    }
    $Script:Cache['TenantInfo'] = @($info)
    Export-Dataset '01_Tenant_Info' @($info)
}

function Export-AreaPolicies {
    Export-Dataset '10_DLP_Policies_Full' @(Get-Policies)
    Export-Dataset '11_DLP_Policies_Summary' @(Get-PolicySummary) -AlreadyFlat
    Export-Dataset '12_DLP_Workload_Coverage' @(Get-Coverage) -AlreadyFlat
}

function Export-AreaRules {
    Export-Dataset '13_DLP_Rules_Full' @(Get-Rules)
    Export-Dataset '14_DLP_Rules_Conditions_Actions' @(Get-RuleSummary) -AlreadyFlat
    Export-Dataset '16_Endpoint_DLP_Rule_Restrictions' @(Get-EndpointRuleActions) -AlreadyFlat
}

function Export-AreaClassifiers {
    $rc = @(Get-RuleClassifiers)
    Export-Dataset '15_DLP_Rule_Classifiers_and_Confidence' $rc -AlreadyFlat
    Export-Dataset '26_Trainable_Classifiers_In_DLP_Rules' @($rc | Where-Object { $_.ReferenceType -eq 'Classifier' -and $_.ClassifierType -match 'MLModel|Trainable|Classifier' }) -AlreadyFlat
    $conf = @($rc | Where-Object ReferenceType -eq 'Classifier' |
        Group-Object { if ($_.ConfidenceLevel) { [string]$_.ConfidenceLevel } elseif ($_.MinConfidence) { "MinConfidence $($_.MinConfidence)" } else { 'Not specified' } } |
        ForEach-Object { [pscustomobject]@{ Confidence = $_.Name; References = $_.Count; Rules = (@($_.Group.Rule | Sort-Object -Unique) -join ' | ') } })
    $Script:Cache['ConfidenceSummary'] = $conf
    Export-Dataset '27_Confidence_Level_Summary' $conf -AlreadyFlat
}

function Export-AreaSITs {
    $view = @(Get-SITView)
    Export-Dataset '20_Sensitive_Information_Types_All' $view -AlreadyFlat
    Export-Dataset '21_Sensitive_Information_Types_Custom' @($view | Where-Object IsCustom) -AlreadyFlat
    $rps = @(Get-RulePackages)
    Export-Dataset '22_SIT_Rule_Packages' @(foreach ($rp in $rps) {
            [pscustomobject]@{
                Name = $rp.Name; Identity = ConvertTo-CellValue (Get-Val $rp 'Identity'); Publisher = ConvertTo-CellValue (Get-Val $rp 'Publisher')
                Version = ConvertTo-CellValue (Get-Val $rp 'Version'); RuleCount = @((Get-Val $rp 'Rules')).Count; WhenChangedUTC = Get-Val $rp 'WhenChangedUTC'
            }
        }) -AlreadyFlat
    foreach ($rp in $rps) {
        $pub = ConvertTo-CellValue (Get-Val $rp 'Publisher'); $xml = Get-Val $rp 'SerializedClassificationRuleCollection'
        if ($xml -and $pub -notmatch 'Microsoft') {
            $safe = ($rp.Name -replace '[^\w\-]', '_')
            try {
                $bytes = if ($xml -is [byte[]]) { $xml } else { [byte[]]$xml }
                [IO.File]::WriteAllBytes((Join-Path $Script:OutputDir "22_RulePackage_$safe.xml"), $bytes)
                Write-Log "  Saved custom rule package XML: $safe" 'OK'
            }
            catch { Write-Log "  Could not save rule package XML for $($rp.Name): $($_.Exception.Message)" 'WARN' }
        }
    }
    Export-Dataset '23_Keyword_Dictionaries' @(Get-Dictionaries | Select-Object Name, Identity, Description, @{n = 'KeywordCount'; e = { @(Get-Val $_ 'KeywordDictionary' | ForEach-Object { $_ -split ',' }).Count } }, WhenChangedUTC)
}

function Export-AreaEDM {
    Export-Dataset '24_EDM_Schemas' @(Get-EdmSchemas)
    Export-Dataset '25_EDM_Sensitive_Information_Types' @(Get-SITView | Where-Object { $_.Classifier -match 'ExactMatch|EDM' -or $_.Type -match 'Exact|EDM' }) -AlreadyFlat
    Export-Dataset '15_DLP_Rule_Classifiers_and_Confidence' @(Get-RuleClassifiers) -AlreadyFlat
}

function Export-AreaLabels {
    Export-Dataset '30_Sensitivity_Labels' @(Get-LabelView) -AlreadyFlat
    Export-Dataset '31_Sensitivity_Label_Policies' @(Get-LabelPolicies)
    Export-Dataset '32_Auto_Labelling_Policies' @(Get-AutoLabelPolicies)
    Export-Dataset '33_Auto_Labelling_Rules' @(Get-AutoLabelRules)
    Export-Dataset '34_Labels_Used_As_DLP_Conditions' @(Get-RuleClassifiers | Where-Object ReferenceType -eq 'SensitivityLabel') -AlreadyFlat
}

function Export-AreaEndpoint {
    Export-Dataset '40_Endpoint_DLP_PolicyConfig_Full' @(Get-PolicyConfigData)
    Export-Dataset '41_Endpoint_DLP_Settings' @(Get-EndpointSettings) -AlreadyFlat
    Export-Dataset '16_Endpoint_DLP_Rule_Restrictions' @(Get-EndpointRuleActions) -AlreadyFlat
    Export-Dataset '12_DLP_Workload_Coverage' @(Get-Coverage) -AlreadyFlat
}

function Export-AreaAlerts {
    $sel = 'Name', 'Category', 'Severity', 'Disabled', 'Operation', 'NotifyUser', 'AggregationType', 'Threshold', 'TimeWindow', 'IsSystemRule', 'Comment', 'CreatedBy', 'LastModifiedTime'
    Export-Dataset '50_Alert_Policies_All' @(Get-AlertPolicies | Select-Object $sel)
    Export-Dataset '51_Alert_Policies_DLP' @(Get-DlpAlertPolicies | Select-Object $sel)
    Export-Dataset '52_DLP_Rule_Alerting' @(Get-RuleSummary | Where-Object { $_.GenerateAlert -or $_.GenerateIncidentReport } |
        Select-Object Policy, Rule, ReportSeverityLevel, GenerateAlert, AlertProperties, GenerateIncidentReport, IncidentReportContent) -AlreadyFlat
}

function Export-AreaOwnership {
    Export-Dataset '53_Purview_Role_Groups' @(Get-RoleGroups | Select-Object Name, Description, Roles, WhenChangedUTC)
    Export-Dataset '54_Purview_Role_Group_Members' @(Get-RoleGroupMembers) -AlreadyFlat
    Export-Dataset '55_DLP_Relevant_Role_Members' @(Get-DlpRoleMembers) -AlreadyFlat
}

function Export-AreaActivity {
    $recs = @(Get-ActivityRecords)
    Export-Dataset '60_Activity_Explorer_Records' $recs
    Export-Dataset '61_Activity_Explorer_Summary' @($recs | Group-Object { '{0}|{1}' -f (Get-Val $_ 'Activity'), (Get-Val $_ 'Workload') } | ForEach-Object {
            $parts = $_.Name -split '\|', 2
            [pscustomobject]@{ Activity = $parts[0]; Workload = $parts[1]; Count = $_.Count }
        } | Sort-Object Count -Descending) -AlreadyFlat
}

function Export-AreaExchange {
    try { Connect-Exo }
    catch {
        Write-Log "Exchange Online connection failed: $($_.Exception.Message)" 'ERROR'
        $Script:Errors.Add([pscustomobject]@{ Step = 'Mail flow (transport) rules'; Outcome = 'Failed'; Error = $_.Exception.Message })
        return
    }
    $tr = Invoke-Step 'Mail flow (transport) rules' -Requires Get-TransportRule -Action { Get-TransportRule }
    $Script:Cache['TransportRules'] = @($tr)
    Export-Dataset '70_Exchange_Transport_Rules' @(foreach ($t in $tr) {
            [pscustomobject]@{
                Name = $t.Name; State = Get-Val $t 'State'; Mode = Get-Val $t 'Mode'; Priority = Get-Val $t 'Priority'; DlpPolicy = Get-Val $t 'DlpPolicy'
                UsesSensitiveInfo = [bool](Get-Val $t 'MessageContainsDataClassifications'); SentToScope = Get-Val $t 'SentToScope'; Description = Get-Val $t 'Description'
            }
        }) -AlreadyFlat
    $audit = Invoke-Step 'Unified audit log configuration' -Requires Get-AdminAuditLogConfig -Action { Get-AdminAuditLogConfig | Select-Object UnifiedAuditLogIngestionEnabled, AdminAuditLogEnabled }
    $Script:Cache['Audit'] = @($audit)
    Export-Dataset '71_Audit_Configuration' $audit
    Reset-AfterExo
}

$Script:AreaFunctions = @{
    '1' = 'Export-AreaTenant'; '2' = 'Export-AreaPolicies'; '3' = 'Export-AreaRules'; '4' = 'Export-AreaClassifiers'
    '5' = 'Export-AreaSITs'; '6' = 'Export-AreaEDM'; '7' = 'Export-AreaLabels'; '8' = 'Export-AreaEndpoint'
    '9' = 'Export-AreaAlerts'; '10' = 'Export-AreaOwnership'; '11' = 'Export-AreaActivity'; '12' = 'Export-AreaExchange'; '13' = 'Export-AreaGraph'
}

#endregion

#region ======================= Microsoft Graph (optional, separate process) ==========

$Script:GraphScopes = @('LicenseAssignment.Read.All', 'SecurityAlert.Read.All', 'RoleManagement.Read.Directory', 'SharePointTenantSettings.Read.All')

# Collector script, run in a separate PowerShell process so the Microsoft Graph and Exchange Online
# modules never share a session (their sign-in libraries can clash).
$Script:GraphCollectorCode = @'
param([string]$OutFile, [string]$TenantId, [string]$Scopes, [switch]$TestOnly, [int]$AlertDays = 90, [int]$MaxAlertPages = 10)
$ErrorActionPreference = 'Stop'
$result = [ordered]@{ Context = $null; Skus = @(); Alerts = @(); Roles = @(); Spo = $null; Errors = @(); Tests = @() }
function Add-Err([string]$Step, [string]$Msg) { $script:result.Errors += [pscustomobject]@{ Step = $Step; Outcome = 'Failed'; Error = $Msg } }
function Get-Pages([string]$Uri, [int]$Max = 50) {
    $items = @(); $next = $Uri; $n = 0
    while ($next -and $n -lt $Max) {
        $r = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($r.value) { $items += @($r.value) }
        $next = $r.'@odata.nextLink'; $n++
    }
    , $items
}
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $c = @{ Scopes = ($Scopes -split ','); NoWelcome = $true; ContextScope = 'Process' }
    if ($TenantId) { $c['TenantId'] = $TenantId }
    Write-Host 'Signing in to Microsoft Graph (use the same admin account)...' -ForegroundColor Gray
    Connect-MgGraph @c | Out-Null
    $ctx = Get-MgContext
    $result.Context = [pscustomobject]@{ Account = $ctx.Account; TenantId = $ctx.TenantId; Scopes = (@($ctx.Scopes) -join ', ') }
}
catch {
    Add-Err 'Graph: Connect' $_.Exception.Message
    $result | ConvertTo-Json -Depth 8 | Set-Content -Path $OutFile -Encoding UTF8
    exit 2
}

$tests = [ordered]@{
    'Graph: Licensing'          = @{ Scope = 'LicenseAssignment.Read.All'; Uri = 'https://graph.microsoft.com/v1.0/subscribedSkus' }
    'Graph: DLP alerts'         = @{ Scope = 'SecurityAlert.Read.All'; Uri = 'https://graph.microsoft.com/v1.0/security/alerts_v2?$filter=serviceSource eq ''dataLossPrevention''&$top=1' }
    'Graph: Entra role holders' = @{ Scope = 'RoleManagement.Read.Directory'; Uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$top=1' }
    'Graph: SharePoint sharing' = @{ Scope = 'SharePointTenantSettings.Read.All'; Uri = 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' }
}
if ($TestOnly) {
    foreach ($k in $tests.Keys) {
        try { [void](Invoke-MgGraphRequest -Method GET -Uri $tests[$k].Uri -OutputType PSObject); $o = 'PASS'; $d = 'Returned successfully.' }
        catch { $o = 'FAIL'; $d = ($_.Exception.Message -split "`n")[0] }
        $result.Tests += [pscustomobject]@{ Area = '13 ' + $k; Cmdlet = "GET $($tests[$k].Uri -replace '\?.*$', '') ($($tests[$k].Scope))"; Result = $o; Detail = $d }
    }
}
else {
    Write-Host 'Collecting: subscriptions and service plans' -ForegroundColor Gray
    try {
        $skus = Get-Pages 'https://graph.microsoft.com/v1.0/subscribedSkus'
        $result.Skus = @(foreach ($s in $skus) {
                [pscustomobject]@{
                    SkuPartNumber = $s.skuPartNumber; SkuId = $s.skuId; CapabilityStatus = $s.capabilityStatus
                    ConsumedUnits = $s.consumedUnits; EnabledUnits = $s.prepaidUnits.enabled
                    ServicePlans = ((@($s.servicePlans | ForEach-Object { "$($_.servicePlanName)=$($_.provisioningStatus)" })) -join ';')
                }
            })
    }
    catch { Add-Err 'Graph: Licensing' $_.Exception.Message }

    Write-Host 'Collecting: DLP alerts' -ForegroundColor Gray
    try {
        $since = (Get-Date).ToUniversalTime().AddDays(-$AlertDays)
        $alerts = Get-Pages 'https://graph.microsoft.com/v1.0/security/alerts_v2?$filter=serviceSource eq ''dataLossPrevention''&$top=200' $MaxAlertPages
        $result.Alerts = @(foreach ($a in $alerts) {
                $created = $null; try { $created = [datetime]$a.createdDateTime } catch { }
                if ($created -and $created -lt $since) { continue }
                [pscustomobject]@{
                    Created = $a.createdDateTime; Title = $a.title; Severity = $a.severity; Status = $a.status
                    Classification = $a.classification; Determination = $a.determination; AssignedTo = $a.assignedTo
                    LastUpdated = $a.lastUpdateDateTime; Id = $a.id
                }
            })
    }
    catch { Add-Err 'Graph: DLP alerts' $_.Exception.Message }

    Write-Host 'Collecting: Microsoft Entra role holders' -ForegroundColor Gray
    try {
        $defs = @{}; foreach ($d in (Get-Pages 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName')) { $defs[$d.id] = $d.displayName }
        $wanted = 'Global Administrator', 'Global Reader', 'Compliance Administrator', 'Compliance Data Administrator', 'Security Administrator', 'Security Operator', 'Security Reader', 'Exchange Administrator', 'SharePoint Administrator', 'Teams Administrator', 'Privileged Role Administrator'
        $assign = Get-Pages 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=principal'
        $result.Roles = @(foreach ($r in $assign) {
                $name = $defs[$r.roleDefinitionId]
                if ($wanted -notcontains $name) { continue }
                $p = $r.principal
                [pscustomobject]@{
                    Role = $name
                    Principal = $(if ($p.userPrincipalName) { $p.userPrincipalName } else { $p.displayName })
                    PrincipalType = ([string]$p.'@odata.type') -replace '#microsoft.graph.', ''
                    Scope = $r.directoryScopeId
                }
            })
    }
    catch { Add-Err 'Graph: Entra role holders' $_.Exception.Message }

    Write-Host 'Collecting: SharePoint and OneDrive sharing settings' -ForegroundColor Gray
    try { $result.Spo = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/admin/sharepoint/settings' -OutputType PSObject }
    catch { Add-Err 'Graph: SharePoint sharing' $_.Exception.Message }
}
try { Disconnect-MgGraph | Out-Null } catch { }
$result | ConvertTo-Json -Depth 8 | Set-Content -Path $OutFile -Encoding UTF8
exit 0
'@

function Get-ChildPowerShell {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        foreach ($n in 'pwsh.exe', 'pwsh') { $p = Join-Path $PSHOME $n; if (Test-Path $p) { return $p } }
        return 'pwsh'
    }
    $p = Join-Path $PSHOME 'powershell.exe'
    if (Test-Path $p) { return $p }
    return 'powershell.exe'
}

function Invoke-GraphCollector {
    <# Runs the collector in a separate PowerShell process and returns its JSON result as an object. #>
    param([switch]$TestOnly)
    $mod = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $mod) {
        $msg = 'Microsoft.Graph.Authentication module is not installed. Use menu option 1 to install it.'
        Write-Log "Microsoft Graph skipped - $msg" 'WARN'
        $Script:Errors.Add([pscustomobject]@{ Step = 'Graph: Connect'; Outcome = 'Skipped - module not installed'; Error = $msg })
        return $null
    }
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("DlpGraph-" + [guid]::NewGuid().ToString('N'))
    $script = "$tmp.ps1"; $out = "$tmp.json"
    Set-Content -Path $script -Value $Script:GraphCollectorCode -Encoding UTF8
    $c = Get-IppsConnection
    $tenant = if ($c) { [string]$c.TenantID } else { '' }
    $exe = Get-ChildPowerShell
    Write-Log "Starting Microsoft Graph collection in a separate PowerShell process ($([IO.Path]::GetFileName($exe)))..."
    Write-Log "Read-only permissions requested: $($Script:GraphScopes -join ', '). An administrator may need to approve them the first time." 'INFO'
    $args2 = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script, '-OutFile', $out, '-Scopes', ($Script:GraphScopes -join ','), '-AlertDays', '90')
    if ($tenant) { $args2 += @('-TenantId', $tenant) }
    if ($TestOnly) { $args2 += '-TestOnly' }
    try { & $exe @args2 } catch { Write-Log "Graph collector could not start: $($_.Exception.Message)" 'ERROR' }
    $res = $null
    if (Test-Path $out) {
        try { $res = Get-Content -Path $out -Raw | ConvertFrom-Json } catch { Write-Log "Could not read Graph results: $($_.Exception.Message)" 'ERROR' }
    }
    Remove-Item $script, $out -Force -ErrorAction SilentlyContinue
    if (-not $res) {
        $Script:Errors.Add([pscustomobject]@{ Step = 'Graph: Connect'; Outcome = 'Failed'; Error = 'The Graph collector did not return results. See the console output above.' })
        return $null
    }
    foreach ($e in @($res.Errors)) { if ($e) { $Script:Errors.Add($e); Write-Log "  $($e.Step) failed: $($e.Error)" 'ERROR' } }
    if ($res.Context) { Write-Log "Microsoft Graph: signed in as $($res.Context.Account)" 'OK' }
    return $res
}

# Licence indicators: features that usually need premium licensing, matched against SKU and service plan names.
$Script:LicenceFeatures = @(
    @{ Feature = 'DLP for Exchange, SharePoint and OneDrive'; Area = 'Coverage'; Plans = 'EXCHANGE_S_ENTERPRISE|SHAREPOINTENTERPRISE|MIP_S_CLP1'; Skus = 'ENTERPRISEPACK|ENTERPRISEPREMIUM|SPE_E3|SPE_E5|SPB' }
    @{ Feature = 'DLP for Teams chat and channel messages'; Area = 'Coverage'; Plans = 'COMMUNICATIONS_DLP'; Skus = 'ENTERPRISEPREMIUM|SPE_E5|INFORMATION_PROTECTION_COMPLIANCE|M365_E5_SUITE_COMPONENTS' }
    @{ Feature = 'Endpoint DLP (devices)'; Area = 'Endpoint'; Plans = 'MICROSOFTENDPOINTDLP'; Skus = 'SPE_E5|INFORMATION_PROTECTION_COMPLIANCE|M365_E5_SUITE_COMPONENTS' }
    @{ Feature = 'Auto-labelling and premium information protection'; Area = 'Labels'; Plans = 'MIP_S_CLP2'; Skus = 'ENTERPRISEPREMIUM|SPE_E5|INFORMATION_PROTECTION_COMPLIANCE|M365_E5_SUITE_COMPONENTS' }
    @{ Feature = 'Trainable classifiers and Exact Data Match'; Area = 'EDM'; Plans = 'ML_CLASSIFICATION|MIP_S_CLP2'; Skus = 'ENTERPRISEPREMIUM|SPE_E5|INFORMATION_PROTECTION_COMPLIANCE|M365_E5_SUITE_COMPONENTS' }
    @{ Feature = 'Microsoft 365 Copilot'; Area = 'Copilot'; Plans = 'M365_COPILOT'; Skus = 'Copilot' }
)

function Get-LicenceIndicators {
    param([object[]]$Skus)
    $active = @($Skus | Where-Object { $_.CapabilityStatus -match 'Enabled|Warning' -and [int]('0' + $_.EnabledUnits) -gt 0 })
    foreach ($f in $Script:LicenceFeatures) {
        $hits = @()
        foreach ($s in $active) {
            if ([string]$s.SkuPartNumber -match "(?i)$($f.Skus)") { $hits += "$($s.SkuPartNumber) ($($s.ConsumedUnits)/$($s.EnabledUnits))"; continue }
            $plans = @(([string]$s.ServicePlans) -split ';' | Where-Object { $_ -match "(?i)^($($f.Plans))=Success" })
            if ($plans.Count) { $hits += "$($s.SkuPartNumber) [$(($plans | ForEach-Object { ($_ -split '=')[0] }) -join ', ')] ($($s.ConsumedUnits)/$($s.EnabledUnits))" }
        }
        [pscustomobject]@{ Feature = $f.Feature; Area = $f.Area; Indicator = $(if ($hits.Count) { 'Found' } else { 'Not found' }); Evidence = ($hits -join '; ') }
    }
}

function Export-AreaGraph {
    $g = Invoke-GraphCollector
    if (-not $g) { return }
    $Script:Cache['Graph'] = $g
    $skus = @($g.Skus | Where-Object { $_ })
    Export-Dataset '80_Graph_Subscriptions' $skus -AlreadyFlat
    $ind = @(Get-LicenceIndicators -Skus $skus)
    $Script:Cache['LicenceIndicators'] = $ind
    Export-Dataset '81_Graph_DLP_Licence_Indicators' $ind -AlreadyFlat
    $alerts = @($g.Alerts | Where-Object { $_ })
    $Script:Cache['GraphAlerts'] = $alerts
    Export-Dataset '82_Graph_DLP_Alerts' $alerts -AlreadyFlat
    Export-Dataset '83_Graph_DLP_Alert_Summary' @($alerts | Group-Object Severity, Status, Classification | ForEach-Object {
            $p = $_.Name -split ', '
            [pscustomobject]@{ Severity = $p[0]; Status = $p[1]; Classification = $(if ($p.Count -gt 2) { $p[2] } else { '' }); Count = $_.Count }
        }) -AlreadyFlat
    $roles = @($g.Roles | Where-Object { $_ })
    $Script:Cache['EntraRoles'] = $roles
    Export-Dataset '84_Graph_Entra_Role_Holders' $roles -AlreadyFlat
    if ($g.Spo) {
        $spoRows = @(foreach ($p in $g.Spo.PSObject.Properties) { if ($p.Name -notmatch '^@odata') { [pscustomobject]@{ Setting = $p.Name; Value = ConvertTo-CellValue $p.Value } } })
        $Script:Cache['SpoSettings'] = $g.Spo
        Export-Dataset '85_Graph_SharePoint_Settings' $spoRows -AlreadyFlat
    }
}

#endregion

#region ======================= Word report (no Word or extra modules required) ==========

# Microsoft Learn links used in the report. Kept in one place so they are easy to update.
$Script:LearnLinks = @{
    DlpOverview      = @{ Title = 'Learn about data loss prevention'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp' }
    DlpDeploy        = @{ Title = 'Create and deploy data loss prevention policies'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-create-deploy-policy' }
    DlpReference     = @{ Title = 'Data loss prevention policy reference'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-policy-reference' }
    DlpExchange      = @{ Title = 'DLP Exchange conditions and actions reference'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-exchange-conditions-and-actions' }
    DlpSimulation    = @{ Title = 'Learn about DLP simulation mode'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-simulation-mode-learn' }
    DlpSimStart      = @{ Title = 'Get started with DLP simulation mode'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-simulation-mode-get-started' }
    PolicyTips       = @{ Title = 'Policy tips in Outlook and Office apps'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-ol365-win32-policy-tips' }
    SitOverview      = @{ Title = 'Learn about sensitive information types'; Url = 'https://learn.microsoft.com/en-us/purview/sensitive-information-type-learn-about' }
    SitDefinitions   = @{ Title = 'Sensitive information type entity definitions'; Url = 'https://learn.microsoft.com/en-us/purview/sit-sensitive-information-type-entity-definitions' }
    SitCustom        = @{ Title = 'Create custom sensitive information types'; Url = 'https://learn.microsoft.com/en-us/purview/sit-create-a-custom-sensitive-information-type' }
    KeywordDict      = @{ Title = 'Create a keyword dictionary'; Url = 'https://learn.microsoft.com/en-us/purview/sit-create-a-keyword-dictionary' }
    EdmLearn         = @{ Title = 'Learn about exact data match based sensitive information types'; Url = 'https://learn.microsoft.com/en-us/purview/sit-learn-about-exact-data-match-based-sits' }
    EdmStart         = @{ Title = 'Get started with exact data match based sensitive information types'; Url = 'https://learn.microsoft.com/en-us/purview/sit-get-started-exact-data-match-based-sits-overview' }
    Classifiers      = @{ Title = 'Learn about trainable classifiers'; Url = 'https://learn.microsoft.com/en-us/purview/classifier-learn-about' }
    Labels           = @{ Title = 'Learn about sensitivity labels'; Url = 'https://learn.microsoft.com/en-us/purview/sensitivity-labels' }
    AutoLabel        = @{ Title = 'Apply a sensitivity label to content automatically'; Url = 'https://learn.microsoft.com/en-us/purview/apply-sensitivity-label-automatically' }
    LabelCondition   = @{ Title = 'Use sensitivity labels as conditions in DLP policies'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-sensitivity-label-as-condition' }
    EndpointLearn    = @{ Title = 'Learn about Endpoint data loss prevention'; Url = 'https://learn.microsoft.com/en-us/purview/endpoint-dlp-learn-about' }
    EndpointSettings = @{ Title = 'Configure endpoint DLP settings'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-configure-endpoint-settings' }
    EndpointUsing    = @{ Title = 'Use Endpoint data loss prevention'; Url = 'https://learn.microsoft.com/en-us/purview/endpoint-dlp-using' }
    Copilot          = @{ Title = 'Learn about using DLP to protect interactions with Microsoft 365 Copilot'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-microsoft365-copilot-location-learn-about' }
    Alerts           = @{ Title = 'Get started with DLP alerts'; Url = 'https://learn.microsoft.com/en-us/purview/dlp-alerts-dashboard-get-started' }
    ActivityExplorer = @{ Title = 'Get started with Activity explorer'; Url = 'https://learn.microsoft.com/en-us/purview/data-classification-activity-explorer' }
    Permissions      = @{ Title = 'Permissions in the Microsoft Purview portal'; Url = 'https://learn.microsoft.com/en-us/purview/purview-permissions' }
    MailFlow         = @{ Title = 'Mail flow rules (transport rules) in Exchange Online'; Url = 'https://learn.microsoft.com/en-us/exchange/security-and-compliance/mail-flow-rules/mail-flow-rules' }
}

$Script:DocWidth = 9864   # A4 portrait content width in twips (1.8 cm margins)

function New-WordDoc {
    $Script:Doc = @{
        Body  = [System.Text.StringBuilder]::new()
        Rels  = [System.Collections.Generic.List[string]]::new()
        RelId = 100
    }
    $Script:ReportObs = [System.Collections.Generic.List[object]]::new()
}

function ConvertTo-XmlText {
    param($Text)
    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    $s = $s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return [System.Security.SecurityElement]::Escape($s)
}

function New-DocRuns {
    <# Builds runs from text. **text** is bold. #>
    param([string]$Text, [int]$Size = 0, [string]$Color = '', [switch]$Bold, [switch]$Italic)
    $sb = [System.Text.StringBuilder]::new()
    $parts = [regex]::Split([string]$Text, '(\*\*[^*]+\*\*)')
    foreach ($p in $parts) {
        if ([string]::IsNullOrEmpty($p)) { continue }
        $b = $Bold
        if ($p.StartsWith('**') -and $p.EndsWith('**') -and $p.Length -gt 4) { $p = $p.Substring(2, $p.Length - 4); $b = $true }
        $rpr = ''
        if ($b) { $rpr += '<w:b/>' }
        if ($Italic) { $rpr += '<w:i/>' }
        if ($Color) { $rpr += "<w:color w:val=`"$Color`"/>" }
        if ($Size -gt 0) { $rpr += "<w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/>" }
        if ($rpr) { $rpr = "<w:rPr>$rpr</w:rPr>" }
        [void]$sb.Append("<w:r>$rpr<w:t xml:space=`"preserve`">$(ConvertTo-XmlText $p)</w:t></w:r>")
    }
    return $sb.ToString()
}

function Add-DocXml { param([string]$Xml) [void]$Script:Doc.Body.Append($Xml) }

function Add-DocHeading {
    param([string]$Text, [int]$Level = 1)
    Add-DocXml "<w:p><w:pPr><w:pStyle w:val=`"Heading$Level`"/></w:pPr>$(New-DocRuns $Text)</w:p>"
}

function Add-DocPara {
    param([string]$Text, [int]$Size = 0, [string]$Color = '', [switch]$Bold, [switch]$Italic, [string]$Align = '', [int]$After = -1, [int]$Before = -1)
    $ppr = ''
    if ($Align) { $ppr += "<w:jc w:val=`"$Align`"/>" }
    if ($After -ge 0 -or $Before -ge 0) {
        $sp = '<w:spacing'
        if ($Before -ge 0) { $sp += " w:before=`"$Before`"" }
        if ($After -ge 0) { $sp += " w:after=`"$After`"" }
        $ppr = "$sp/>" + $ppr
    }
    if ($ppr) { $ppr = "<w:pPr>$ppr</w:pPr>" }
    Add-DocXml "<w:p>$ppr$(New-DocRuns $Text -Size $Size -Color $Color -Bold:$Bold -Italic:$Italic)</w:p>"
}

function Add-DocBullet {
    param([string]$Text)
    Add-DocXml "<w:p><w:pPr><w:pStyle w:val=`"ListParagraph`"/><w:numPr><w:ilvl w:val=`"0`"/><w:numId w:val=`"1`"/></w:numPr><w:spacing w:after=`"60`"/></w:pPr>$(New-DocRuns $Text)</w:p>"
}

function Add-DocLink {
    param([string]$Title, [string]$Url, [switch]$AsBullet)
    $id = "rIdH$($Script:Doc.RelId)"; $Script:Doc.RelId++
    $Script:Doc.Rels.Add("<Relationship Id=`"$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`" Target=`"$(ConvertTo-XmlText $Url)`" TargetMode=`"External`"/>")
    $ppr = if ($AsBullet) { '<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:after="60"/></w:pPr>' } else { '' }
    Add-DocXml "<w:p>$ppr<w:hyperlink r:id=`"$id`" w:history=`"1`"><w:r><w:rPr><w:rStyle w:val=`"Hyperlink`"/></w:rPr><w:t xml:space=`"preserve`">$(ConvertTo-XmlText $Title)</w:t></w:r></w:hyperlink></w:p>"
}

function Add-DocPageBreak { Add-DocXml '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' }

function Add-DocTable {
    <#
      Headers: string[]; Rows: array of string[] (or objects + -Properties); Widths: relative column weights.
    #>
    param([string[]]$Headers, [object[]]$Rows, [double[]]$Widths, [int]$FontSize = 17, [int]$MaxRows = 0, [string]$Note = '')
    $Rows = @($Rows)
    # A single row passed through the pipeline arrives flattened (an array of strings): wrap it back into one row.
    if ($Rows.Count -gt 0 -and -not ($Rows | Where-Object { $_ -is [System.Collections.IList] })) { $Rows = , $Rows }
    $total = $Rows.Count
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) { $Rows = $Rows[0..($MaxRows - 1)] }
    if (-not $Widths -or $Widths.Count -ne $Headers.Count) { $Widths = @(1) * $Headers.Count }
    $sum = ($Widths | Measure-Object -Sum).Sum
    $tw = @($Widths | ForEach-Object { [int]([math]::Floor($Script:DocWidth * $_ / $sum)) })
    $border = '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFC8D6"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFC8D6"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFC8D6"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFC8D6"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFC8D6"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFC8D6"/>'
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<w:tbl><w:tblPr><w:tblW w:w=`"$Script:DocWidth`" w:type=`"dxa`"/><w:tblBorders>$border</w:tblBorders><w:tblLayout w:type=`"fixed`"/><w:tblCellMar><w:top w:w=`"50`" w:type=`"dxa`"/><w:left w:w=`"90`" w:type=`"dxa`"/><w:bottom w:w=`"50`" w:type=`"dxa`"/><w:right w:w=`"90`" w:type=`"dxa`"/></w:tblCellMar></w:tblPr><w:tblGrid>")
    foreach ($w in $tw) { [void]$sb.Append("<w:gridCol w:w=`"$w`"/>") }
    [void]$sb.Append('</w:tblGrid>')
    # header
    [void]$sb.Append('<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>')
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        [void]$sb.Append("<w:tc><w:tcPr><w:tcW w:w=`"$($tw[$i])`" w:type=`"dxa`"/><w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"0F4C81`"/></w:tcPr><w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`"/></w:pPr>$(New-DocRuns $Headers[$i] -Size $FontSize -Color 'FFFFFF' -Bold)</w:p></w:tc>")
    }
    [void]$sb.Append('</w:tr>')
    $r = 0
    foreach ($row in $Rows) {
        $cells = @($row)
        $fill = if ($r % 2 -eq 1) { '<w:shd w:val="clear" w:color="auto" w:fill="EEF3F9"/>' } else { '' }
        [void]$sb.Append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $v = if ($i -lt $cells.Count) { [string]$cells[$i] } else { '' }
            if ($v.Length -gt 600) { $v = $v.Substring(0, 600) + ' ...' }
            $cellFill = $fill; $bold = $false
            if ($v -match '^\[\[([GARN])\]\]') {
                $rag = @{ G = 'C6EFCE'; A = 'FFEB9C'; R = 'FFC7CE'; N = 'E7E6E6' }[$Matches[1]]
                $cellFill = "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$rag`"/>"; $v = $v.Substring(5); $bold = $true
            }
            [void]$sb.Append("<w:tc><w:tcPr><w:tcW w:w=`"$($tw[$i])`" w:type=`"dxa`"/>$cellFill</w:tcPr><w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`"/></w:pPr>$(New-DocRuns $v -Size $FontSize -Bold:$bold)</w:p></w:tc>")
        }
        [void]$sb.Append('</w:tr>')
        $r++
    }
    [void]$sb.Append('</w:tbl>')
    Add-DocXml $sb.ToString()
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) {
        Add-DocPara "Showing the first $MaxRows of $total rows. $Note" -Size 16 -Color '666666' -Italic -After 160
    }
    elseif ($Note) { Add-DocPara $Note -Size 16 -Color '666666' -Italic -After 160 }
    else { Add-DocPara '' -After 80 }
}

function Add-DocCallout {
    param([string]$Label, [string]$Text, [string]$Fill = 'EEF3F9', [string]$Edge = '0F4C81')
    $w = $Script:DocWidth
    $tb = "<w:top w:val=`"nil`"/><w:left w:val=`"single`" w:sz=`"24`" w:space=`"0`" w:color=`"$Edge`"/><w:bottom w:val=`"nil`"/><w:right w:val=`"nil`"/>"
    Add-DocXml ("<w:tbl><w:tblPr><w:tblW w:w=`"$w`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/></w:tblPr><w:tblGrid><w:gridCol w:w=`"$w`"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=`"$w`" w:type=`"dxa`"/>" +
        "<w:tcBorders>$tb</w:tcBorders><w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$Fill`"/><w:tcMar><w:top w:w=`"100`" w:type=`"dxa`"/><w:left w:w=`"160`" w:type=`"dxa`"/><w:bottom w:w=`"100`" w:type=`"dxa`"/><w:right w:w=`"160`" w:type=`"dxa`"/></w:tcMar></w:tcPr>" +
        "<w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`"/></w:pPr>$(New-DocRuns ($Label + '  ') -Bold -Size 19)$(New-DocRuns $Text -Size 19)</w:p></w:tc></w:tr></w:tbl>")
    Add-DocPara '' -After 80
}

function Add-DocMetrics {
    <# Grid of headline numbers, four per row. Items: @{ Value; Label; Color } #>
    param([object[]]$Items)
    $cols = 4
    $cw = [int]([math]::Floor($Script:DocWidth / $cols))
    $sb = [System.Text.StringBuilder]::new()
    $bd = '<w:top w:val="single" w:sz="18" w:space="0" w:color="FFFFFF"/><w:left w:val="single" w:sz="18" w:space="0" w:color="FFFFFF"/><w:bottom w:val="single" w:sz="18" w:space="0" w:color="FFFFFF"/><w:right w:val="single" w:sz="18" w:space="0" w:color="FFFFFF"/>'
    [void]$sb.Append("<w:tbl><w:tblPr><w:tblW w:w=`"$($cw * $cols)`" w:type=`"dxa`"/><w:tblLayout w:type=`"fixed`"/></w:tblPr><w:tblGrid>")
    for ($i = 0; $i -lt $cols; $i++) { [void]$sb.Append("<w:gridCol w:w=`"$cw`"/>") }
    [void]$sb.Append('</w:tblGrid>')
    for ($i = 0; $i -lt $Items.Count; $i += $cols) {
        [void]$sb.Append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        for ($j = 0; $j -lt $cols; $j++) {
            $it = if (($i + $j) -lt $Items.Count) { $Items[$i + $j] } else { $null }
            $color = if ($it -and $it.Color) { $it.Color } else { '0F4C81' }
            $fill = if ($it) { 'EEF3F9' } else { 'FFFFFF' }
            [void]$sb.Append("<w:tc><w:tcPr><w:tcW w:w=`"$cw`" w:type=`"dxa`"/><w:tcBorders>$bd</w:tcBorders><w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$fill`"/><w:tcMar><w:top w:w=`"120`" w:type=`"dxa`"/><w:bottom w:w=`"120`" w:type=`"dxa`"/></w:tcMar></w:tcPr>")
            if ($it) {
                [void]$sb.Append("<w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`"/><w:jc w:val=`"center`"/></w:pPr>$(New-DocRuns ([string]$it.Value) -Size 40 -Color $color -Bold)</w:p>")
                [void]$sb.Append("<w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`"/><w:jc w:val=`"center`"/></w:pPr>$(New-DocRuns ([string]$it.Label) -Size 16 -Color '444444')</w:p>")
            }
            else { [void]$sb.Append('<w:p/>') }
            [void]$sb.Append('</w:tc>')
        }
        [void]$sb.Append('</w:tr>')
    }
    [void]$sb.Append('</w:tbl>')
    Add-DocXml $sb.ToString()
    Add-DocPara '' -After 80
}

function Add-DocToc {
    Add-DocXml '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>'
    Add-DocXml '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-2" \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/><w:color w:val="777777"/></w:rPr><w:t>Right-click here and choose Update Field to build the table of contents.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
}

function Add-Obs {
    <# Records an observation for the executive summary and returns its text. Severity: High, Medium, Low, Info #>
    param([string]$Area, [string]$Severity, [string]$Text)
    $Script:ReportObs.Add([pscustomobject]@{ Area = $Area; Severity = $Severity; Text = $Text })
    return $Text
}

function Add-SectionIntro {
    param([string]$What, [string]$Why)
    Add-DocHeading 'What is it?' 3
    foreach ($p in ($What -split '\|\|')) { Add-DocPara $p.Trim() }
    Add-DocHeading 'Why it matters' 3
    foreach ($p in ($Why -split '\|\|')) { Add-DocPara $p.Trim() }
    Add-DocHeading 'What we found' 3
}

function Add-SectionGap {
    param([string[]]$Steps)
    $g = Get-StepGap $Steps
    if ($g) { Add-DocCallout 'Not fully collected:' ($g.Trim() -replace '^NOT COLLECTED: ', '') 'FFF4E5' 'E08A00' }
}

function Add-SectionEnd {
    param([string[]]$Considerations, [string[]]$LinkKeys, [string]$Area)
    Add-DocHeading 'Things to consider' 3
    $c = @($Considerations | Where-Object { $_ })
    if ($c.Count -eq 0) { Add-DocPara 'No specific issues were identified from the configuration collected for this area.' }
    else { foreach ($x in $c) { Add-DocBullet $x } }
    if ($Area) { Add-FrameworkTable -Area $Area }
    Add-DocHeading 'Learn more on Microsoft Learn' 3
    foreach ($k in $LinkKeys) { $l = $Script:LearnLinks[$k]; if ($l) { Add-DocLink -Title $l.Title -Url $l.Url -AsBullet } }
}

function Get-Plural { param([int]$n, [string]$one, [string]$many) if ($n -eq 1) { "$n $one" } else { "$n $many" } }

function Save-WordDoc {
    param([string]$Path, [string]$Title, [string]$Author)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $bullet = [string][char]0x2022

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'
    $rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'
    $docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' + ($Script:Doc.Rels -join '') + '</Relationships>'
    $font = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + $W + '">' +
        '<w:docDefaults><w:rPrDefault><w:rPr>' + $font + '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-GB"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="200"/></w:pPr><w:rPr><w:b/><w:color w:val="0F4C81"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:pageBreakBefore/><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="0F4C81"/></w:pBdr><w:spacing w:before="0" w:after="200"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="0F4C81"/><w:sz w:val="34"/><w:szCs w:val="34"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="300" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="0F4C81"/><w:sz w:val="27"/><w:szCs w:val="27"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3B57"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:pageBreakBefore/><w:spacing w:before="0" w:after="200"/></w:pPr><w:rPr><w:b/><w:color w:val="0F4C81"/><w:sz w:val="34"/><w:szCs w:val="34"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9854"/></w:tabs><w:spacing w:before="80" w:after="40"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9854"/></w:tabs><w:spacing w:before="0" w:after="20"/><w:ind w:left="240"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="540"/></w:pPr></w:style>' +
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
        '</w:styles>'
    $numbering = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="' + $W + '"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="' + $bullet + '"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="540" w:hanging="270"/></w:pPr><w:rPr>' + $font + '</w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    $settings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="' + $W + '"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:updateFields w:val="true"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
    $footer = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="' + $W + '" xmlns:r="' + $R + '"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">' + (ConvertTo-XmlText $Title) + '  |  Page </w:t></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
    $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + (ConvertTo-XmlText $Title) + '</dc:title><dc:creator>' + (ConvertTo-XmlText $Author) + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:modified></cp:coreProperties>'
    $sect = '<w:sectPr><w:footerReference w:type="default" r:id="rId4"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1021" w:bottom="1134" w:left="1021" w:header="567" w:footer="567" w:gutter="0"/><w:titlePg/></w:sectPr>'
    $document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + $W + '" xmlns:r="' + $R + '"><w:body>' + $Script:Doc.Body.ToString() + $sect + '</w:body></w:document>'

    $parts = [ordered]@{
        '[Content_Types].xml'          = $contentTypes
        '_rels/.rels'                  = $rootRels
        'docProps/core.xml'            = $core
        'word/document.xml'            = $document
        'word/_rels/document.xml.rels' = $docRels
        'word/styles.xml'              = $styles
        'word/numbering.xml'           = $numbering
        'word/settings.xml'            = $settings
        'word/footer1.xml'             = $footer
    }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $enc = New-Object System.Text.UTF8Encoding($false)
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $parts.Keys) {
                $entry = $zip.CreateEntry($name)
                $sw = New-Object System.IO.StreamWriter($entry.Open(), $enc)
                try { $sw.Write($parts[$name]) } finally { $sw.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

# ---------- Framework references (indicative mappings; confirm against the framework documents) ----------
$Script:FrameworkInfo = [ordered]@{
    NCSC      = @{ Name = 'NCSC Cyber Assessment Framework (CAF) v4.0 and 10 Steps'; Short = 'NCSC'; Url = 'https://www.ncsc.gov.uk/collection/cyber-assessment-framework'; About = 'UK National Cyber Security Centre framework of outcomes for organisations that support essential services. Principle B3 (Data security) is the main principle for DLP. The NCSC 10 Steps to Cyber Security (Data security, Logging and monitoring) give lighter-weight guidance for other organisations.' }
    NISTCSF   = @{ Name = 'NIST Cybersecurity Framework (CSF) 2.0'; Short = 'NIST CSF 2.0'; Url = 'https://www.nist.gov/cyberframework'; About = 'US NIST outcome-based framework. The Data Security category (PR.DS) covers data at rest (PR.DS-01), in transit (PR.DS-02) and in use (PR.DS-10).' }
    NIST80053 = @{ Name = 'NIST SP 800-53 Rev. 5 security controls'; Short = 'NIST 800-53'; Url = 'https://csrc.nist.gov/pubs/sp/800/53/r5/upd1/final'; About = 'Catalogue of security controls. AC-4 (Information flow enforcement) and SC-7(10) (Prevent exfiltration) are the controls most directly met by DLP. NIST SP 1800-28 (Data Confidentiality) gives a practical reference architecture.' }
    ISO27001  = @{ Name = 'ISO/IEC 27001:2022 Annex A'; Short = 'ISO 27001'; Url = 'https://www.iso.org/standard/27001'; About = 'International standard for information security management. Control 8.12 (Data leakage prevention) was added in the 2022 revision and is the core DLP control.' }
}

# Area -> framework references. Expects = what the framework is looking for, in plain English.
$Script:FrameworkMap = @{
    Policies    = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3 Data security; 10 Steps: Data security'; Expects = 'Data is protected from unauthorised access, modification or deletion, using controls proportionate to its importance.' }
        @{ Fw = 'NISTCSF'; Ref = 'GV.PO-01, PR.DS'; Expects = 'A policy for managing data risk exists and data security outcomes are implemented.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-4, SC-7(10)'; Expects = 'Approved rules control where information can flow, and exfiltration is prevented.' }
        @{ Fw = 'ISO27001'; Ref = '8.12, 5.1'; Expects = 'Data leakage prevention measures are applied and supported by an information security policy.' }
    )
    Coverage    = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.b Data in transit; B3.c Stored data'; Expects = 'Data is protected when sent (email, chat, sharing) and where it is stored (sites, OneDrive).' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.DS-01, PR.DS-02'; Expects = 'Confidentiality of data at rest and in transit is protected.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-4, SC-7(10)'; Expects = 'Information flow controls apply on every route data can leave the organisation.' }
        @{ Fw = 'ISO27001'; Ref = '8.12, 5.14'; Expects = 'DLP applies to the systems and channels that process sensitive data, including information transfer.' }
    )
    Rules       = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.b, B3.c'; Expects = 'Controls actively prevent inappropriate sharing, not only record it.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.DS-01, PR.DS-02'; Expects = 'Data protection is enforced in line with risk.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-4, SC-7(10)'; Expects = 'Unauthorised flows are blocked, with approved exceptions.' }
        @{ Fw = 'ISO27001'; Ref = '8.12, 5.14'; Expects = 'Leakage is detected and prevented, with defined actions.' }
    )
    UX          = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B6 Staff awareness'; Expects = 'Staff understand how to handle data securely and are supported to do so.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.AT-01'; Expects = 'People are made aware of security risks relevant to their work.' }
        @{ Fw = 'NIST80053'; Ref = 'AT-2'; Expects = 'Users receive security awareness at the point it is needed.' }
        @{ Fw = 'ISO27001'; Ref = '6.3, 8.12'; Expects = 'Awareness of information security responsibilities; DLP measures explained to users.' }
    )
    SIT         = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a Understanding data'; Expects = 'You know what sensitive data you hold and can recognise it.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.AM-07'; Expects = 'Sensitive data types are identified and inventoried.' }
        @{ Fw = 'NIST80053'; Ref = 'RA-2, AC-16'; Expects = 'Information is categorised and its sensitivity can be identified.' }
        @{ Fw = 'ISO27001'; Ref = '5.12, 8.12'; Expects = 'Information is classified, and DLP identifies sensitive information.' }
    )
    CustomSIT   = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a Understanding data'; Expects = 'Organisation-specific important data is understood and identifiable.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.AM-07'; Expects = 'Designated data types, including organisation-specific ones, are inventoried.' }
        @{ Fw = 'NIST80053'; Ref = 'RA-2, AC-16'; Expects = 'Categorisation covers the organisation''s own information types.' }
        @{ Fw = 'ISO27001'; Ref = '5.12, 8.12'; Expects = 'Classification reflects the organisation''s own information.' }
    )
    EDM         = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a Understanding data'; Expects = 'The most important datasets are known and precisely identified.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.AM-07, PR.DS-01'; Expects = 'Key datasets are inventoried and protected.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-16, SC-7(10)'; Expects = 'Accurate identification supports effective exfiltration controls.' }
        @{ Fw = 'ISO27001'; Ref = '8.12'; Expects = 'Detection is accurate enough to act on.' }
    )
    Classifiers = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a Understanding data'; Expects = 'Sensitive document types are recognised, not only identifiers.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.AM-07'; Expects = 'Data types are identified, including unstructured content.' }
        @{ Fw = 'NIST80053'; Ref = 'RA-2, AC-16'; Expects = 'Content is categorised by its nature.' }
        @{ Fw = 'ISO27001'; Ref = '5.12, 8.12'; Expects = 'Classification covers documents as well as data items.' }
    )
    Confidence  = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a, B3.b'; Expects = 'Controls are tuned so they protect data without excessive disruption.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.DS-01, PR.DS-02'; Expects = 'Protection is effective in practice.' }
        @{ Fw = 'NIST80053'; Ref = 'SC-7(10)'; Expects = 'Exfiltration controls are tested and effective.' }
        @{ Fw = 'ISO27001'; Ref = '8.12'; Expects = 'DLP is operating effectively.' }
    )
    Labels      = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a, B3.c Stored data'; Expects = 'Sensitivity of data is understood and stored data is protected accordingly.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.AM-07, PR.DS-01'; Expects = 'Data is classified and protected at rest.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-16, MP-3'; Expects = 'Security attributes (labels) are applied and content is marked.' }
        @{ Fw = 'ISO27001'; Ref = '5.12, 5.13, 8.12'; Expects = 'Information is classified and labelled, and labels drive handling.' }
    )
    Endpoint    = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.d Mobile data; 10 Steps: Data security'; Expects = 'Data on devices and removable media is protected.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.DS-10, PR.DS-01'; Expects = 'Data in use on devices is protected.' }
        @{ Fw = 'NIST80053'; Ref = 'MP-7, SC-7(10)'; Expects = 'Use of removable media is controlled and exfiltration from endpoints is prevented.' }
        @{ Fw = 'ISO27001'; Ref = '7.10, 8.1, 8.12'; Expects = 'Storage media and user endpoint devices are controlled; DLP applies to endpoints.' }
    )
    Copilot     = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.a, B3.c'; Expects = 'Sensitive data is protected wherever it is processed, including by AI services.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.DS-10'; Expects = 'Data in use is protected.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-4'; Expects = 'Information flows into AI services follow approved rules.' }
        @{ Fw = 'ISO27001'; Ref = '8.12, 5.23'; Expects = 'DLP extends to cloud services, including AI.' }
    )
    Simulation  = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B1 Policies, processes and procedures'; Expects = 'Protection is delivered through defined, managed processes.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.IM-01'; Expects = 'Evaluations are used to identify improvements.' }
        @{ Fw = 'NIST80053'; Ref = 'CM-4'; Expects = 'The impact of changes is analysed before they are made.' }
        @{ Fw = 'ISO27001'; Ref = '8.32'; Expects = 'Changes are controlled and tested.' }
    )
    Alerts      = @(
        @{ Fw = 'NCSC'; Ref = 'CAF C1 Security monitoring; 10 Steps: Logging and monitoring'; Expects = 'Security events are monitored and acted on.' }
        @{ Fw = 'NISTCSF'; Ref = 'DE.CM-03, DE.AE'; Expects = 'Activity is monitored and adverse events are analysed.' }
        @{ Fw = 'NIST80053'; Ref = 'SI-4, AU-13'; Expects = 'Systems are monitored, including for information disclosure.' }
        @{ Fw = 'ISO27001'; Ref = '8.16, 8.12'; Expects = 'Activities are monitored and leakage events are acted on.' }
    )
    Activity    = @(
        @{ Fw = 'NCSC'; Ref = 'CAF C1 Security monitoring'; Expects = 'Monitoring data is collected and reviewed.' }
        @{ Fw = 'NISTCSF'; Ref = 'DE.CM-03'; Expects = 'Personnel activity is monitored to find adverse events.' }
        @{ Fw = 'NIST80053'; Ref = 'AU-13, SI-4'; Expects = 'Information disclosure is monitored.' }
        @{ Fw = 'ISO27001'; Ref = '8.15, 8.16'; Expects = 'Activity is logged and monitored.' }
    )
    Exchange    = @(
        @{ Fw = 'NCSC'; Ref = 'CAF C1 Security monitoring; B3.b'; Expects = 'Audit data is available and email controls are consistent.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.PS-04, DE.CM-03'; Expects = 'Log records are generated and available for monitoring.' }
        @{ Fw = 'NIST80053'; Ref = 'AU-2, AU-12, AC-4'; Expects = 'Auditable events are logged; email flow controls are managed.' }
        @{ Fw = 'ISO27001'; Ref = '8.15, 5.14'; Expects = 'Logging is enabled and information transfer is controlled.' }
    )
    Ownership   = @(
        @{ Fw = 'NCSC'; Ref = 'CAF A1 Governance; B2 Identity and access control'; Expects = 'Clear accountability, and privileged access is limited and managed.' }
        @{ Fw = 'NISTCSF'; Ref = 'GV.RR-02, PR.AA-05'; Expects = 'Roles and responsibilities are defined; access follows least privilege.' }
        @{ Fw = 'NIST80053'; Ref = 'AC-5, AC-6'; Expects = 'Duties are separated and privileges are minimised.' }
        @{ Fw = 'ISO27001'; Ref = '5.2, 5.3, 8.2'; Expects = 'Roles are assigned, duties segregated and privileged access controlled.' }
    )
}
$Script:AreaNames = [ordered]@{
    Policies = 'DLP policies'; Coverage = 'Where DLP is applied'; Rules = 'Rules and actions'; UX = 'Policy tips and overrides'
    SIT = 'Sensitive information types'; CustomSIT = 'Custom SITs and dictionaries'; EDM = 'Exact Data Match'; Classifiers = 'Trainable classifiers'
    Confidence = 'Confidence levels'; Labels = 'Sensitivity labels'; Endpoint = 'Endpoint DLP (devices)'; Copilot = 'Microsoft 365 Copilot'
    Simulation = 'Deployment and simulation'; Alerts = 'Alerts'; Activity = 'Activity Explorer'; Exchange = 'Exchange mail flow and audit'; Ownership = 'Operational ownership'
}
$Script:StatusText = @{ G = 'Supported'; A = 'Partial'; R = 'Gap'; N = 'Not assessed' }

function Get-SelectedFrameworks {
    $sel = @($Script:Settings.Frameworks | ForEach-Object { ($_ -split ',') } | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ })
    if ($sel -contains 'NIST') { $sel += 'NISTCSF', 'NIST80053' }
    if ($sel -contains 'ISO') { $sel += 'ISO27001' }
    if ($sel -contains 'ALL' -or $sel.Count -eq 0) { $sel = @($Script:FrameworkInfo.Keys) }
    @($Script:FrameworkInfo.Keys | Where-Object { $sel -contains $_.ToUpper() })
}

function Add-FrameworkTable {
    param([string]$Area)
    $fws = @(Get-SelectedFrameworks)
    if ($fws.Count -eq 0 -or -not $Script:FrameworkMap.ContainsKey($Area)) { return }
    $st = if ($Script:AreaStatus.ContainsKey($Area)) { $Script:AreaStatus[$Area] } else { 'N' }
    Add-DocHeading 'Framework alignment' 3
    $rows = @(foreach ($m in $Script:FrameworkMap[$Area]) {
            if ($fws -notcontains $m.Fw) { continue }
            , @($Script:FrameworkInfo[$m.Fw].Short, $m.Ref, $m.Expects, "[[$st]]$($Script:StatusText[$st])")
        })
    Add-DocTable -Headers @('Framework', 'Reference', 'What the framework expects', 'This tenant') -Widths @(1.3, 1.9, 4, 1.1) -FontSize 16 -Rows $rows
}

function Add-DocCode {
    <# Monospace, shaded block for PowerShell examples. #>
    param([string]$Code)
    $lines = @($Code.Trim("`r", "`n") -split "`r?`n")
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $after = if ($i -eq $lines.Count - 1) { 160 } else { 0 }
        $t = if ($lines[$i] -eq '') { ' ' } else { $lines[$i] }
        Add-DocXml "<w:p><w:pPr><w:pBdr><w:left w:val=`"single`" w:sz=`"18`" w:space=`"6`" w:color=`"0F4C81`"/></w:pBdr><w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"F2F4F7`"/><w:spacing w:before=`"0`" w:after=`"$after`" w:line=`"240`" w:lineRule=`"auto`"/><w:ind w:left=`"180`"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=`"Consolas`" w:hAnsi=`"Consolas`" w:cs=`"Consolas`"/><w:color w:val=`"1E1E1E`"/><w:sz w:val=`"16`"/><w:szCs w:val=`"16`"/></w:rPr><w:t xml:space=`"preserve`">$(ConvertTo-XmlText $t)</w:t></w:r></w:p>"
    }
}

function Add-DocSteps {
    param([string[]]$Steps)
    $n = 1
    foreach ($s in $Steps) {
        Add-DocXml "<w:p><w:pPr><w:spacing w:after=`"60`"/><w:ind w:left=`"540`" w:hanging=`"360`"/></w:pPr>$(New-DocRuns "**$n.**")<w:r><w:tab/></w:r>$(New-DocRuns $s)</w:p>"
        $n++
    }
}

function Add-Guidance {
    <# One block of implementation guidance. #>
    param([string]$Title, [string]$When, [string]$Area, [string[]]$Portal, [string]$PowerShell, [string[]]$Validate, [string[]]$LinkKeys)
    Add-DocHeading $Title 2
    Add-DocCallout 'Why this applies:' $When 'FFF4E5' 'E08A00'
    if ($Area -and $Script:FrameworkMap.ContainsKey($Area)) {
        $fws = @(Get-SelectedFrameworks)
        $refs = @($Script:FrameworkMap[$Area] | Where-Object { $fws -contains $_.Fw } | ForEach-Object { "$($Script:FrameworkInfo[$_.Fw].Short) $($_.Ref)" })
        if ($refs.Count) { Add-DocPara ("**Supports:** " + ($refs -join '  |  ')) -Size 18 -Color '1F3B57' }
    }
    Add-DocHeading 'Steps in the Microsoft Purview portal' 3
    Add-DocSteps $Portal
    if ($PowerShell) {
        Add-DocHeading 'PowerShell example' 3
        Add-DocPara 'Run in Security & Compliance PowerShell (Connect-IPPSSession). Replace the example names and values, and create new policies in simulation mode first.' -Size 18 -Color '555555' -Italic
        Add-DocCode $PowerShell
    }
    if ($Validate) {
        Add-DocHeading 'Check it worked' 3
        foreach ($v in $Validate) { Add-DocBullet $v }
    }
    if ($LinkKeys) {
        Add-DocHeading 'Microsoft Learn' 3
        foreach ($k in $LinkKeys) { $l = $Script:LearnLinks[$k]; if ($l) { Add-DocLink -Title $l.Title -Url $l.Url -AsBullet } }
    }
}

function Add-FrameworkMatrix {
    $fws = @(Get-SelectedFrameworks)
    if ($fws.Count -eq 0) { return }
    $headers = @('Area') + @($fws | ForEach-Object { $Script:FrameworkInfo[$_].Short }) + @('This tenant')
    $widths = @(2.2) + @($fws | ForEach-Object { 1.9 }) + @(1.2)
    $rows = @(foreach ($a in $Script:AreaNames.Keys) {
            $st = if ($Script:AreaStatus.ContainsKey($a)) { $Script:AreaStatus[$a] } else { 'N' }
            $cells = @($Script:AreaNames[$a])
            foreach ($f in $fws) { $cells += ((@($Script:FrameworkMap[$a] | Where-Object { $_.Fw -eq $f } | ForEach-Object { $_.Ref })) -join '; ') }
            $cells += "[[$st]]$($Script:StatusText[$st])"
            , $cells
        })
    Add-DocTable -Headers $headers -Widths $widths -FontSize 15 -Rows $rows
}

function Add-ImplementationGuidance {
    param($Ctx)
    Add-DocHeading '22. Applying the configuration' 1
    Add-DocPara 'This section gives practical guidance for the gaps found in this assessment. Only guidance relevant to this tenant is included. Each item shows why it applies, which framework references it supports, the steps in the Microsoft Purview portal, an example PowerShell command, and how to check the change worked.'
    Add-DocCallout 'Before you start:' 'Agree each change with the policy owner, test new or changed policies in simulation mode first, pilot with a small group, and record changes through your normal change process. Menu names in the Microsoft Purview portal change from time to time; if a path differs, follow the linked Microsoft Learn article.' 'EEF3F9' '0F4C81'

    Add-DocHeading 'Recommended order of work' 2
    Add-DocSteps @(
        'Agree the sensitive information types and scenarios that matter most (customer, financial, employee and organisation-specific data).'
        'Make sure sensitivity labels are defined and published, so people and policies can use them.'
        'Create or extend DLP policies to cover email, SharePoint, OneDrive and Teams, starting in simulation mode.'
        'Extend protection to devices (Endpoint DLP) and Microsoft 365 Copilot.'
        'Review simulation results, tune confidence and instance counts, and agree policy tip wording.'
        'Move policies to enforcement in waves, with a pilot group first.'
        'Set up alerting, review processes and ownership, and re-run this assessment to confirm the changes.'
    )
    $n = 0
    if ($Ctx.Gaps.Count -gt 0 -or $Ctx.Policies -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Create or extend DLP policies to all core locations" -Area 'Coverage' `
            -When $(if ($Ctx.Policies -eq 0) { 'No DLP policies exist in the tenant.' } else { "No enforced policy covers: $($Ctx.Gaps -join ', ')." }) `
            -Portal @(
            'Go to the Microsoft Purview portal (purview.microsoft.com) > Solutions > Data loss prevention > Policies.'
            'To extend an existing policy: select it, choose Edit policy, and on the Locations page turn on the missing locations (Exchange email, SharePoint sites, OneDrive accounts, Teams chat and channel messages).'
            'To create a new policy: select Create policy, choose a template that matches the data (for example a UK or financial template) or Custom, and select the locations.'
            'Define rules: conditions (sensitive information types or labels, and whether content is shared outside the organisation) and actions (policy tip, block, block with override).'
            'On the Policy mode page select Run the policy in simulation mode, with policy tips if you want users to see them during the pilot.'
            'Review and submit.'
        ) -PowerShell @'
# Extend an existing policy to Teams (repeat for other missing locations)
Set-DlpCompliancePolicy -Identity "<Existing policy name>" -AddTeamsLocation All

# Or create a new baseline policy in simulation mode across the core locations
New-DlpCompliancePolicy -Name "Baseline - Sensitive data" `
    -ExchangeLocation All -SharePointLocation All -OneDriveLocation All -TeamsLocation All `
    -Mode TestWithNotifications

New-DlpComplianceRule -Name "Baseline - Credit cards shared externally" -Policy "Baseline - Sensitive data" `
    -ContentContainsSensitiveInformation @{Name = "Credit Card Number"; minCount = "5"} `
    -AccessScope NotInOrganization -BlockAccess $true `
    -NotifyUser Owner -NotifyAllowOverride WithJustification `
    -GenerateIncidentReport SiteAdmin -ReportSeverityLevel High
'@ -Validate @(
            'The policy shows the expected locations and mode under Data loss prevention > Policies.'
            'Get-DlpCompliancePolicy -Identity "<name>" | Format-List Name, Mode, Workload, DistributionStatus shows DistributionStatus Success.'
            'Simulation results appear on the policy''s simulation dashboard within a few hours.'
        ) -LinkKeys @('DlpDeploy', 'DlpReference')
    }
    if ($Ctx.Test -gt 0) {
        $n++
        Add-Guidance -Title "22.$n Move policies from simulation to enforcement" -Area 'Simulation' `
            -When "$($Ctx.Test) policies are in simulation or test mode." `
            -Portal @(
            'Open each policy in Data loss prevention > Policies and review its simulation results: matches, top sensitive information types, locations and users.'
            'Investigate a sample of matches to estimate false positives; adjust conditions, confidence or instance counts if needed.'
            'Agree the go-live date and pilot group with the policy owner and communicate it to users.'
            'Edit the policy, and on the Policy mode page select Turn the policy on immediately.'
            'Monitor alerts and user feedback closely for the first two weeks.'
        ) -PowerShell @'
# List policies in simulation
Get-DlpCompliancePolicy | Where-Object Mode -like 'Test*' | Select-Object Name, Mode, Workload

# Turn on enforcement for one policy
Set-DlpCompliancePolicy -Identity "<Policy name>" -Mode Enable
'@ -Validate @('Get-DlpCompliancePolicy shows Mode Enable for the policy.', 'Rule matches now apply their actions (check Activity Explorer for DLPRuleMatch events).') -LinkKeys @('DlpSimulation', 'DlpSimStart')
    }
    if ($Ctx.AuditOnly -gt 0 -or ($Ctx.Rules -gt 0 -and $Ctx.Tips -eq 0)) {
        $n++
        Add-Guidance -Title "22.$n Add policy tips and proportionate actions to rules" -Area 'UX' `
            -When $(if ($Ctx.AuditOnly -gt 0) { "$($Ctx.AuditOnly) rules only audit, so users are not told and nothing is prevented." } else { 'No rules notify users.' }) `
            -Portal @(
            'Open the policy in Data loss prevention > Policies, choose Edit policy and go to Advanced DLP rules.'
            'Edit the rule. Under User notifications, turn on notifications and policy tips, and add short custom text: what is wrong, what to do, who to contact.'
            'Under Actions, choose the response that fits the risk: notify only for low volumes, block with override (justification required) for medium risk, block for high-risk external sharing.'
            'Under User overrides, allow overrides with a business justification where appropriate.'
            'Save and keep the policy in simulation until the wording and actions are agreed.'
        ) -PowerShell @'
Set-DlpComplianceRule -Identity "<Rule name>" `
    -NotifyUser Owner `
    -NotifyPolicyTipCustomText "This item contains sensitive information. Remove it or share only with approved recipients. Questions: servicedesk@contoso.com" `
    -BlockAccess $true -NotifyAllowOverride WithJustification
'@ -Validate @('Send a test item containing sample data to an external test address and confirm the policy tip appears.', 'Override justifications appear in Activity Explorer.') -LinkKeys @('PolicyTips', 'DlpReference')
    }
    if ($Ctx.Labels -eq 0 -or $Ctx.LabelPolicies -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Define and publish sensitivity labels" -Area 'Labels' `
            -When $(if ($Ctx.Labels -eq 0) { 'No sensitivity labels exist.' } else { 'Labels exist but no label policy publishes them to users.' }) `
            -Portal @(
            'Agree a simple taxonomy with the business, typically four or five labels such as Public, General, Confidential and Highly Confidential, with sublabels only where needed.'
            'Go to Microsoft Purview portal > Solutions > Information Protection > Sensitivity labels and create each label, including any encryption and content marking.'
            'Go to Publishing policies (label policies) and publish the labels to a pilot group, with a default label and justification for downgrades if required.'
            'Expand to all users once the pilot is complete.'
        ) -PowerShell @'
# Publish existing labels to a pilot group
New-LabelPolicy -Name "Pilot label policy" `
    -Labels "Public","General","Confidential","Highly Confidential" `
    -ExchangeLocation "pilot-group@contoso.com"
'@ -Validate @('Pilot users see the labels in Office apps within 24 hours.', 'Get-LabelPolicy lists the policy with the expected labels.') -LinkKeys @('Labels', 'AutoLabel')
    }
    if ($Ctx.Labels -gt 0 -and $Ctx.LabelRefs -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Use sensitivity labels as DLP conditions" -Area 'Labels' `
            -When 'Sensitivity labels exist but no DLP rule uses them.' `
            -Portal @(
            'Open (or create) a DLP policy that covers Exchange, SharePoint, OneDrive and Teams.'
            'Create a rule with the condition Content contains > Sensitivity labels, and select the highest labels (for example Highly Confidential).'
            'Add the condition Content is shared from Microsoft 365 > with people outside my organization.'
            'Set the action to restrict access or block, with a policy tip explaining why.'
            'Run in simulation, review results, then enforce.'
        ) -PowerShell @'
# Find the label GUID
Get-Label | Format-Table DisplayName, Guid

$labelCondition = @{
    operator = "And"
    groups   = @(@{
        operator = "Or"
        name     = "Default"
        labels   = @(@{ name = "<Label GUID>"; type = "Sensitivity" })
    })
}
New-DlpComplianceRule -Name "Block external sharing - Highly Confidential" -Policy "<Policy name>" `
    -ContentContainsSensitiveInformation $labelCondition `
    -AccessScope NotInOrganization -BlockAccess $true -NotifyUser Owner
'@ -Validate @('A Highly Confidential test document cannot be shared with an external test account.', '15_DLP_Rule_Classifiers_and_Confidence shows the label as a condition when this assessment is re-run.') -LinkKeys @('LabelCondition')
    }
    if ($Ctx.EndpointCov -eq 0 -or $Ctx.Usb -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Onboard devices and restrict USB, printing and uploads" -Area 'Endpoint' `
            -When $(if ($Ctx.EndpointCov -eq 0) { 'No enforced DLP policy covers devices.' } else { 'Device policies exist but none restrict copying to USB or removable storage.' }) `
            -Portal @(
            'Confirm prerequisites: supported Windows 10/11 or macOS devices and the required licences.'
            'Onboard devices: Microsoft Purview portal > Settings > Device onboarding (devices already onboarded to Microsoft Defender for Endpoint can be turned on there).'
            'Configure tenant-wide settings: Settings > Data loss prevention > Endpoint DLP settings: service domains (allowed or blocked upload sites), unallowed browsers, printer groups, removable storage device groups and file path exclusions.'
            'Create or edit a DLP policy with the Devices location. In the rule, add the action Audit or restrict activities on devices and set Copy to removable USB device, Print, and Upload to a restricted cloud service domain to Audit, Warn (block with override) or Block.'
            'Run in simulation with a pilot group, then enforce.'
        ) -PowerShell @'
New-DlpCompliancePolicy -Name "Devices - Sensitive data" `
    -EndpointDlpLocation All -Mode TestWithNotifications

New-DlpComplianceRule -Name "Devices - USB and print" -Policy "Devices - Sensitive data" `
    -ContentContainsSensitiveInformation @{Name = "Credit Card Number"} `
    -EndpointDlpRestrictions @(
        @{Setting = "RemovableMedia"; Value = "Block"},
        @{Setting = "Print";          Value = "Warn"}
    )
# Values can be Audit, Warn or Block. Configure cloud upload restrictions and device groups
# in Endpoint DLP settings (see Microsoft Learn).
'@ -Validate @('Onboarded devices appear in the device list with a healthy policy sync status.', 'Copying a test file containing sample data to USB on a pilot device is blocked or warned.', 'FileCopiedToRemovableMedia events appear in Activity Explorer.') -LinkKeys @('EndpointLearn', 'EndpointSettings', 'EndpointUsing')
    }
    if ($Ctx.CopilotCov -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Protect Microsoft 365 Copilot interactions" -Area 'Copilot' `
            -When 'No DLP policy covers Microsoft 365 Copilot and Copilot Chat.' `
            -Portal @(
            'Go to Data loss prevention > Policies > Create policy, and choose the Custom template (the Copilot location is only available there).'
            'On the Locations page, turn on only Microsoft 365 Copilot and Copilot Chat (selecting it disables other locations for that policy).'
            'Create a rule with the condition Content contains > Sensitivity labels and select the labels Copilot must not use (for example Highly Confidential).'
            'Optionally add a rule that checks prompts for sensitive information types.'
            'Run in simulation, review, then enforce.'
        ) -PowerShell @'
# Adapted from the Microsoft Learn New-DlpCompliancePolicy example for Microsoft 365 Copilot
$loc = '[{"Workload":"Applications","Location":"470f2276-e011-4e9d-a6ec-20768be3a4b0","Inclusions":[{"Type":"Tenant","Identity":"All"}]}]'
New-DlpCompliancePolicy -Name "Copilot - Exclude Highly Confidential" -Locations $loc -EnforcementPlanes @("CopilotExperiences")

$advRule = @{
    Version   = "1.0"
    Condition = @{
        Operator      = "And"
        SubConditions = @(@{
            ConditionName = "ContentContainsSensitiveInformation"
            Value = @(@{ groups = @(@{ Operator = "Or"; name = "Default"
                labels = @(@{ name = "<Label GUID>"; type = "Sensitivity" }) }) })
        })
    }
} | ConvertTo-Json -Depth 100
New-DlpComplianceRule -Name "Copilot - Exclude labelled content" -Policy "Copilot - Exclude Highly Confidential" `
    -AdvancedRule $advRule -RestrictAccess @(@{setting = "ExcludeContentProcessing"; value = "Block"})
'@ -Validate @('Copilot does not summarise or reference a Highly Confidential test document for a pilot user.', 'The policy shows the Microsoft 365 Copilot and Copilot Chat location.') -LinkKeys @('Copilot')
    }
    if ($Ctx.AlertRules -eq 0 -or $Ctx.SiteOnly -gt 0) {
        $n++
        Add-Guidance -Title "22.$n Configure alerts and a review process" -Area 'Alerts' `
            -When $(if ($Ctx.AlertRules -eq 0) { 'No DLP rules raise alerts.' } else { "$($Ctx.SiteOnly) rules send alerts only to SiteAdmin." }) `
            -Portal @(
            'Agree who triages DLP alerts (a named team or shared mailbox), target response times and escalation routes.'
            'Edit each high-severity rule: under Incident reports, set the severity, turn on alerts to admins, and add the team mailbox as a recipient. Consider aggregated alerts (volume thresholds) for noisy rules.'
            'Review alerts in Data loss prevention > Alerts, or in Microsoft Defender if you manage incidents there.'
            'Document the investigation steps and who can approve exceptions.'
        ) -PowerShell @'
Set-DlpComplianceRule -Identity "<Rule name>" `
    -ReportSeverityLevel High `
    -GenerateAlert "dlp-alerts@contoso.com" `
    -GenerateIncidentReport "dlp-alerts@contoso.com"
'@ -Validate @('A test match produces an alert in the DLP alerts dashboard and an email to the team mailbox.') -LinkKeys @('Alerts')
    }
    if ($Ctx.LowConf -gt 0) {
        $n++
        Add-Guidance -Title "22.$n Tune confidence levels and instance counts" -Area 'Confidence' `
            -When "$($Ctx.LowConf) detections use low confidence, which commonly causes false positives." `
            -Portal @(
            'Use simulation results and Activity Explorer to find the rules and sensitive information types with the most false positives.'
            'Edit the rule and, for each sensitive information type, raise the confidence level (Medium or High) and/or the minimum instance count.'
            'Use different rules for low and high volumes: notify for low counts, block for high counts.'
            'Re-run simulation and compare match volumes.'
        ) -PowerShell @'
Set-DlpComplianceRule -Identity "<Rule name>" `
    -ContentContainsSensitiveInformation @{Name = "<Sensitive information type>"; minCount = "5"; minConfidence = "85"}
'@ -Validate @('Match volumes drop in simulation without losing genuine incidents.') -LinkKeys @('SitOverview', 'SitDefinitions')
    }
    if ($Ctx.Custom -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Create custom sensitive information types for organisation-specific data" -Area 'CustomSIT' `
            -When 'No custom sensitive information types exist, so internal identifiers (customer, account or staff numbers) cannot be detected.' `
            -Portal @(
            'List the organisation-specific identifiers that need protecting, with examples and their exact format.'
            'Go to Microsoft Purview portal > Data classification > Classifiers > Sensitive info types > Create sensitive info type.'
            'Define the primary element (a regular expression or keyword list), supporting keywords near the match, and confidence levels.'
            'Use Test to check the pattern against sample files, then add the new type to a DLP rule in simulation mode.'
        ) -PowerShell '' -Validate @('The custom type matches test documents and appears in 21_Sensitive_Information_Types_Custom when this assessment is re-run.') -LinkKeys @('SitCustom', 'KeywordDict')
    }
    if ($Ctx.Edm -eq 0) {
        $n++
        Add-Guidance -Title "22.$n Consider Exact Data Match for key structured datasets" -Area 'EDM' `
            -When 'Exact Data Match is not in use.' `
            -Portal @(
            'Identify one well-defined dataset (for example the customer or employee list), its owner and how often it changes.'
            'Go to Data classification > Classifiers > EDM classifiers and create an EDM classifier: define the schema (columns and which are searchable) and the primary elements.'
            'Hash and upload the data using the EDM upload agent from a secure workstation, and schedule refreshes.'
            'Use the EDM sensitive information type in a DLP rule in simulation mode.'
        ) -PowerShell '' -Validate @('Test documents containing real records match; documents with look-alike numbers do not.') -LinkKeys @('EdmLearn', 'EdmStart')
    }
    if ($Ctx.AuditOff) {
        $n++
        Add-Guidance -Title "22.$n Turn on the unified audit log" -Area 'Exchange' `
            -When 'The unified audit log is disabled. DLP reporting and Activity Explorer depend on it.' `
            -Portal @('Go to Microsoft Purview portal > Solutions > Audit.', 'If prompted, select Start recording user and admin activity.', 'Allow up to 60 minutes for recording to begin.') `
            -PowerShell @'
# Exchange Online PowerShell (Connect-ExchangeOnline)
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
'@ -Validate @('Get-AdminAuditLogConfig shows UnifiedAuditLogIngestionEnabled True.') -LinkKeys @('ActivityExplorer')
    }
    if ($n -eq 0) { Add-DocPara 'No specific gaps were identified that need configuration guidance. Review the recommended order of work above and the Things to consider in each section.' }
}

function Add-FrameworkAppendix {
    $fws = @(Get-SelectedFrameworks)
    if ($fws.Count -eq 0) { return }
    Add-DocHeading 'Appendix D: Framework references' 1
    Add-DocPara 'This report maps each area of DLP configuration to recognised security frameworks, so findings can be read alongside the organisation''s compliance obligations.'
    Add-DocCallout 'How to read the mapping:' 'The references show which framework outcomes or controls each area of DLP configuration contributes to. They are indicative mappings, not a compliance assessment: framework outcomes also depend on policies, processes and people that this tool cannot see. Confirm references against the framework documents before using them in formal assurance work.' 'FFF4E5' 'E08A00'
    foreach ($f in $fws) {
        $i = $Script:FrameworkInfo[$f]
        Add-DocHeading $i.Name 2
        Add-DocPara $i.About
        Add-DocLink -Title $i.Url -Url $i.Url
    }
    if ($fws -contains 'NIST80053') { Add-DocLink -Title 'NIST SP 1800-28: Data Confidentiality - Identifying and Protecting Assets Against Data Breaches' -Url 'https://csrc.nist.gov/pubs/sp/1800/28/final' }
    Add-DocHeading 'Status values' 2
    Add-DocTable -Headers @('Status', 'Meaning') -Widths @(1.3, 6) -Rows @(
        , @('[[G]]Supported', 'The configuration provides evidence towards the referenced outcomes.')
        , @('[[A]]Partial', 'Some configuration exists, but there are gaps or it is not enforced.')
        , @('[[R]]Gap', 'No configuration was found that supports the referenced outcomes.')
        , @('[[N]]Not assessed', 'The data was not collected in this run, or the area cannot be judged from configuration alone.')
    )
}

function New-DlpWordReport {
    <# Builds DLP_Assessment_Report.docx from the data collected by the complete report. #>
    Write-Log 'Building Word assessment report...'
    New-WordDoc
    $Script:AreaStatus = @{}
    $ti = @($Script:Cache['TenantInfo'])
    $org = if ($ti.Count -and $ti[0].Organization) { $ti[0].Organization } else { 'Microsoft 365 tenant' }
    $tid = if ($ti.Count) { $ti[0].TenantId } else { '' }
    $by = if ($ti.Count) { $ti[0].CollectedBy } else { '' }
    $when = Get-Date -Format 'dd MMMM yyyy HH:mm'

    # ---------- data ----------
    $policies = @(Get-Policies); $polSum = @(Get-PolicySummary); $cov = @(Get-Coverage)
    $rules = @(Get-Rules); $ruleSum = @(Get-RuleSummary); $rc = @(Get-RuleClassifiers)
    $sitView = @(Get-SITView); $custom = @($sitView | Where-Object IsCustom)
    $dicts = @(Get-Dictionaries); $edm = @(Get-EdmSchemas)
    $edmSits = @($sitView | Where-Object { $_.Classifier -match 'ExactMatch|EDM' -or $_.Type -match 'Exact|EDM' })
    $labels = @(Get-LabelView); $labelPol = @(Get-LabelPolicies); $autoPol = @(Get-AutoLabelPolicies)
    $epSet = @(Get-EndpointSettings); $epRules = @(Get-EndpointRuleActions)
    $alertPol = @(Get-DlpAlertPolicies); $roleMembers = @(Get-DlpRoleMembers)
    $summaryRows = @($Script:Cache['SummaryRows'])
    $classRefs = @($rc | Where-Object ReferenceType -eq 'Classifier')
    $labelRefs = @($rc | Where-Object ReferenceType -eq 'SensitivityLabel')

    $pEnf = @($policies | Where-Object { $_.Mode -eq 'Enable' }).Count
    $pTest = @($policies | Where-Object { $_.Mode -like 'Test*' }).Count
    $pOff = @($policies | Where-Object { $_.Mode -eq 'Disable' }).Count
    $respCount = @{}
    foreach ($k in 'Block', 'Block with override', 'Notify / policy tip', 'Audit only') { $respCount[$k] = @($ruleSum | Where-Object { $_.ResponseInterpretation -eq $k }).Count }
    $respCount['Endpoint'] = @($ruleSum | Where-Object { $_.ResponseInterpretation -like 'Endpoint*' }).Count

    # ---------- cover ----------
    Add-DocPara '' -After 2400
    Add-DocXml "<w:p><w:pPr><w:pStyle w:val=`"Title`"/></w:pPr>$(New-DocRuns 'Microsoft Purview Data Loss Prevention Assessment')</w:p>"
    Add-DocPara $org -Size 32 -Color '444444' -After 600
    Add-DocTable -Headers @('Item', 'Detail') -Widths @(1, 3) -FontSize 19 -Rows @(
        , @('Organisation', $org)
        , @('Tenant ID', $tid)
        , @('Data collected', $when)
        , @('Collected by', $by)
        , @('Method', 'Get-M365DlpConfiguration-Beta.ps1 (BETA, read-only discovery using Security & Compliance PowerShell and, if selected, Microsoft Graph)')
    )
    Add-DocCallout 'Confidential:' 'This report describes how sensitive information is protected in this tenant. Store and share it in line with the organisation''s data handling policy.' 'FFF4E5' 'E08A00'
    Add-DocToc

    # ---------- 1 About ----------
    Add-DocHeading '1. About this report' 1
    Add-DocPara 'This report describes the Microsoft Purview Data Loss Prevention (DLP) configuration found in the tenant at the time of collection. It explains, in plain language, what each part of DLP does, what is currently configured, and points worth discussing.'
    Add-DocPara 'The information was gathered with a read-only script. Nothing in the tenant was changed. The report is a point-in-time snapshot; configuration may have changed since it was produced.'
    Add-DocHeading 'How each section is organised' 2
    Add-DocBullet '**What is it?** A short, non-technical explanation of the feature.'
    Add-DocBullet '**Why it matters** How the feature helps protect sensitive information.'
    Add-DocBullet '**What we found** The configuration in this tenant, with counts and tables.'
    Add-DocBullet '**Things to consider** Observations generated from the data, to discuss and validate. These are prompts for discussion, not mandatory changes.'
    Add-DocBullet '**Framework alignment** The NCSC, NIST and ISO 27001 references the area supports, what each expects, and the status in this tenant.'
    Add-DocBullet '**Learn more** Links to the relevant Microsoft Learn documentation.'
    Add-DocPara 'Section 20 checks licensing for DLP features (when Microsoft Graph data is included), and section 22 gives step-by-step guidance for applying the configuration needed to close the gaps found.'
    Add-DocPara 'Large tables are shortened in this report. The full data for every area is in the CSV, JSON and Excel files produced alongside this document (see Appendix C).'

    # ---------- 2 Executive summary ----------
    Add-DocHeading '2. Executive summary' 1
    Add-DocPara 'Headline figures from the collected configuration:'
    $covCount = @{}; foreach ($k in $Script:WorkloadPatterns.Keys) { $covCount[$k] = @($cov | Where-Object { $_.$k -eq $true -and $_.Mode -eq 'Enable' }).Count }
    $workloadsEnforced = @('Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Endpoint', 'Copilot' | Where-Object { $covCount[$_] -gt 0 }).Count
    Add-DocMetrics -Items @(
        @{ Value = $policies.Count; Label = 'DLP policies' }
        @{ Value = $pEnf; Label = 'Policies enforced'; Color = '1E7B34' }
        @{ Value = $pTest; Label = 'In simulation / test'; Color = 'B26B00' }
        @{ Value = $pOff; Label = 'Disabled'; Color = '8A8A8A' }
        @{ Value = $rules.Count; Label = 'DLP rules' }
        @{ Value = "$workloadsEnforced of 6"; Label = 'Core locations with an enforced policy' }
        @{ Value = $custom.Count; Label = 'Custom sensitive info types' }
        @{ Value = $labels.Count; Label = 'Sensitivity labels' }
    )
    $execObsAnchor = $Script:Doc.Body.Length   # observations are inserted here after all sections are built

    # ---------- 3 Policies ----------
    Add-DocHeading '3. DLP policies' 1
    Add-SectionIntro -What 'A DLP policy is the container for an organisation''s data loss rules. Each policy defines **where** it applies (for example email, SharePoint, Teams or devices), **what** it looks for (such as credit card numbers or documents labelled Confidential) and **what happens** when a match is found (such as warning the user or blocking the action).||A policy can be **enforced** (actions are applied), in **simulation / test** mode (matches are reported but nothing is blocked), or **disabled**.' -Why 'Policies are where business and regulatory requirements become technical controls. A clear, well-structured set of policies makes protection predictable for users and easy to maintain for administrators.'
    Add-SectionGap 'DLP compliance policies', 'DLP compliance rules'
    if ($policies.Count -eq 0) { Add-DocPara '**No DLP policies were found in this tenant.**' }
    else {
        Add-DocPara ("The tenant has **$(Get-Plural $policies.Count 'DLP policy' 'DLP policies')**: $pEnf enforced, $pTest in simulation or test mode and $pOff disabled, containing $(Get-Plural $rules.Count 'rule' 'rules') in total.")
        $modeText = @{ 'Enable' = 'Enforced'; 'TestWithNotifications' = 'Simulation (with policy tips)'; 'TestWithoutNotifications' = 'Simulation (silent)'; 'Disable' = 'Disabled' }
        $rows = foreach ($p in ($polSum | Sort-Object { [int]('0' + $_.Priority) })) {
            $m = if ($modeText.ContainsKey([string]$p.Mode)) { $modeText[[string]$p.Mode] } else { [string]$p.Mode }
            $changed = ''
            try { if ($p.WhenChangedUTC) { $changed = ([datetime]$p.WhenChangedUTC).ToString('dd MMM yyyy') } } catch { $changed = [string]$p.WhenChangedUTC }
            , @([string]$p.Name, $m, ([string]$p.Workload -replace 'OneDriveForBusiness', 'OneDrive' -replace ' \| ', ', '), $(@($ruleSum | Where-Object { $_.Policy -eq $p.Name }).Count), $changed)
        }
        Add-DocTable -Headers @('Policy', 'Mode', 'Locations', 'Rules', 'Last changed') -Widths @(3, 2, 3, 0.8, 1.4) -Rows @($rows) -MaxRows 60 -Note 'Full detail: 11_DLP_Policies_Summary.'
    }
    $cons = @()
    if ($policies.Count -eq 0) { $cons += Add-Obs 'DLP policies' 'High' 'No DLP policies are configured, so sensitive information is not currently protected by DLP. Start by agreeing the information types and scenarios that matter most, then deploy policies in simulation mode.' }
    if ($pTest -gt 0) { $cons += Add-Obs 'DLP policies' 'Medium' "$(Get-Plural $pTest 'policy is' 'policies are') in simulation or test mode. Review the simulation results and decide when each should move to enforcement, so protection is not left in test indefinitely." }
    if ($pOff -gt 0) { $cons += Add-Obs 'DLP policies' 'Low' "$(Get-Plural $pOff 'policy is' 'policies are') disabled. Confirm whether they are still needed; removing obsolete policies keeps the configuration easier to understand." }
    $stale = @($polSum | Where-Object { try { $_.WhenChangedUTC -and ([datetime]$_.WhenChangedUTC) -lt (Get-Date).AddMonths(-18) } catch { $false } })
    if ($stale.Count -gt 0) { $cons += Add-Obs 'DLP policies' 'Low' "$(Get-Plural $stale.Count 'policy has' 'policies have') not been changed for over 18 months. Check they still reflect current business requirements and data types." }
    $Script:AreaStatus['Policies'] = $(if ($pEnf -gt 0) { 'G' } elseif ($policies.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('DlpOverview', 'DlpDeploy', 'DlpReference') -Area 'Policies'

    # ---------- 4 Coverage ----------
    Add-DocHeading '4. Where DLP is applied' 1
    Add-SectionIntro -What 'DLP policies are applied to **locations**: the services and places where people work with information. Each location protects a different route by which information could leave the organisation.' -Why 'Sensitive information moves between email, files, chat, devices and now AI assistants. A location with no enforced policy is a route where sensitive data can be shared without any check.'
    $locDesc = [ordered]@{
        'Exchange'   = 'Email and attachments sent from Exchange Online mailboxes.'
        'SharePoint' = 'Documents stored and shared from SharePoint Online sites.'
        'OneDrive'   = 'Documents in users'' OneDrive accounts, including sharing links.'
        'Teams'      = 'Teams chat and channel messages.'
        'Endpoint'   = 'Activity on onboarded Windows and macOS devices: copying to USB, printing, uploading, and more.'
        'Copilot'    = 'Microsoft 365 Copilot and Copilot Chat prompts and responses.'
        'OnPremises' = 'On-premises file shares and SharePoint Server, via the information protection scanner.'
        'PowerBI'    = 'Microsoft Fabric and Power BI workspaces.'
        'CloudApps'  = 'Non-Microsoft cloud apps through Defender for Cloud Apps.'
    }
    $rows = foreach ($k in $locDesc.Keys) {
        $e = @($cov | Where-Object { $_.$k -eq $true -and $_.Mode -eq 'Enable' }).Count
        $t = @($cov | Where-Object { $_.$k -eq $true -and $_.Mode -like 'Test*' }).Count
        $status = if ($e -gt 0) { 'Protected' } elseif ($t -gt 0) { 'Simulation only' } else { 'No policy' }
        , @($k, $locDesc[$k], [string]$e, [string]$t, $status)
    }
    Add-DocTable -Headers @('Location', 'What it covers', 'Enforced', 'Simulation', 'Status') -Widths @(1.3, 4, 1, 1, 1.4) -Rows @($rows) -Note 'Policies can cover several locations, so counts can add up to more than the number of policies. Detail: 12_DLP_Workload_Coverage.'
    $cons = @()
    $core = 'Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Endpoint'
    $gaps = @($core | Where-Object { $covCount[$_] -eq 0 })
    if ($gaps.Count -gt 0 -and $policies.Count -gt 0) { $cons += Add-Obs 'Coverage' 'High' "No enforced policy covers: $($gaps -join ', '). Confirm whether these locations are in scope and, if so, extend existing policies or create new ones." }
    if ($covCount['Teams'] -eq 0 -and $covCount['Exchange'] -gt 0) { $cons += 'Email is protected but Teams is not. Users can often share the same information in a chat message, so consider applying equivalent rules to Teams.' }
    if ($covCount['Endpoint'] -eq 0) { $cons += 'Without Endpoint DLP, protection stops once a file is downloaded to a device. See section 13.' }
    if ($Script:Cache.ContainsKey('SpoSettings')) {
        $spo = $Script:Cache['SpoSettings']
        $capText = @{
            'disabled'                        = 'External sharing is turned off'
            'existingExternalUserSharingOnly' = 'Sharing only with guests already in the directory'
            'externalUserSharingOnly'         = 'Sharing with new and existing guests (sign-in required)'
            'externalUserAndGuestSharing'     = 'Anyone links allowed (no sign-in required)'
        }
        $cap = [string]$spo.sharingCapability
        Add-DocHeading 'SharePoint and OneDrive sharing settings (Microsoft Graph)' 3
        Add-DocPara 'How widely people can share files sets the level of risk that DLP must manage.'
        $rowsSpo = @(
            , @('Organisation-level external sharing', $(if ($capText.ContainsKey($cap)) { $capText[$cap] } else { $cap }))
            , @('Domain restriction mode', [string]$spo.sharingDomainRestrictionMode)
            , @('Allowed domains', ((@($spo.sharingAllowedDomainList)) -join ', '))
            , @('Blocked domains', ((@($spo.sharingBlockedDomainList)) -join ', '))
            , @('Guests can reshare', [string]$spo.isResharingByExternalUsersEnabled)
        )
        Add-DocTable -Headers @('Setting', 'Value') -Widths @(3, 5) -Rows $rowsSpo -Note 'Graph exposes a subset of SharePoint settings; site-level sharing can be more restrictive. Detail: 85_Graph_SharePoint_Settings.'
        if ($cap -eq 'externalUserAndGuestSharing' -and ($covCount['SharePoint'] -eq 0 -or $covCount['OneDrive'] -eq 0)) {
            $cons += Add-Obs 'Coverage' 'High' 'Anyone links are allowed for SharePoint and OneDrive, but no enforced DLP policy covers both locations. Sensitive files can be shared with anyone who has the link.'
        }
        elseif ($cap -eq 'externalUserAndGuestSharing') { $cons += 'Anyone links are allowed. Confirm DLP rules block or warn on sharing sensitive content through anyone links, and consider restricting them for sensitive sites.' }
    }
    $Script:AreaStatus['Coverage'] = $(if ($gaps.Count -eq 0 -and $policies.Count -gt 0) { 'G' } elseif ($gaps.Count -lt 5) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('DlpOverview', 'DlpReference') -Area 'Coverage'

    # ---------- 5 Rules ----------
    Add-DocHeading '5. DLP rules: conditions and actions' 1
    Add-SectionIntro -What 'Each policy contains one or more **rules**. A rule has **conditions** (what to look for, for example "content contains 5 or more UK bank account numbers" or "shared with people outside the organisation"), optional **exceptions**, and **actions** (what to do: audit, notify the user, block, or block but allow an override).' -Why 'Rules decide the user experience and the level of protection. Rules that only audit provide visibility but no protection; rules that block without explanation can disrupt legitimate work. A proportionate mix, matched to the risk, is the goal.'
    Add-SectionGap 'DLP compliance rules'
    if ($rules.Count -gt 0) {
        Add-DocPara "The $(Get-Plural $rules.Count 'rule responds' 'rules respond') as follows:"
        Add-DocTable -Headers @('Response', 'Rules', 'What it means') -Widths @(2, 0.8, 5) -Rows @(
            , @('Block', [string]$respCount['Block'], 'The action is stopped and the user cannot continue.')
            , @('Block with override', [string]$respCount['Block with override'], 'The action is stopped, but the user can continue with a business justification.')
            , @('Notify / policy tip', [string]$respCount['Notify / policy tip'], 'The user is warned and the event is recorded, but the action is allowed.')
            , @('Endpoint restrictions', [string]$respCount['Endpoint'], 'Device activities (USB, print, upload and so on) are audited, warned or blocked; see section 13.')
            , @('Audit only', [string]$respCount['Audit only'], 'The event is recorded for reporting. The user sees nothing and the action is allowed.')
        )
        $rows = foreach ($r in $ruleSum) {
            $al = if ($r.GenerateAlert -or $r.GenerateIncidentReport) { 'Yes' } else { 'No' }
            $resp = ([string]$r.ResponseInterpretation) -replace ' \(see EndpointDlpRestrictions\)', ''
            , @([string]$r.Policy, [string]$r.Rule, $resp, [string]$r.ReportSeverityLevel, $al, $(if ($r.Disabled -eq 'True') { 'Disabled' } else { 'Active' }))
        }
        Add-DocTable -Headers @('Policy', 'Rule', 'Response', 'Severity', 'Alert', 'State') -Widths @(2.4, 2.6, 1.8, 1, 0.8, 1) -Rows @($rows) -MaxRows 60 -Note 'Full conditions, exceptions and actions: 14_DLP_Rules_Conditions_Actions.'
    }
    $cons = @()
    if ($respCount['Audit only'] -gt 0) { $cons += Add-Obs 'Rules' 'Medium' "$(Get-Plural $respCount['Audit only'] 'rule only audits' 'rules only audit'). That is right while tuning, but these rules do not protect information. Plan to add policy tips or blocking once false positives are understood." }
    $highNoAlert = @($ruleSum | Where-Object { $_.ReportSeverityLevel -eq 'High' -and -not ($_.GenerateAlert -or $_.GenerateIncidentReport) })
    if ($highNoAlert.Count -gt 0) { $cons += Add-Obs 'Rules' 'Medium' "$(Get-Plural $highNoAlert.Count 'high-severity rule does' 'high-severity rules do') not raise an alert or incident report, so matches may go unnoticed." }
    $blockNoTip = @($ruleSum | Where-Object { $_.ResponseInterpretation -like 'Block*' -and -not $_.NotifyUser })
    if ($blockNoTip.Count -gt 0) { $cons += "$(Get-Plural $blockNoTip.Count 'blocking rule does' 'blocking rules do') not notify the user. Without a policy tip, users may not understand why their action failed and may contact the service desk." }
    $disabledRules = @($ruleSum | Where-Object { $_.Disabled -eq 'True' })
    if ($disabledRules.Count -gt 0) { $cons += "$(Get-Plural $disabledRules.Count 'rule is' 'rules are') disabled inside otherwise active policies. Confirm this is intentional." }
    $Script:AreaStatus['Rules'] = $(if ($rules.Count -eq 0) { 'R' } elseif ($respCount['Audit only'] -gt 0) { 'A' } else { 'G' })
    Add-SectionEnd $cons @('DlpReference', 'DlpExchange', 'DlpDeploy') -Area 'Rules'

    # ---------- 6 User experience ----------
    Add-DocHeading '6. User experience: policy tips and overrides' 1
    Add-SectionIntro -What 'A **policy tip** is a message shown to users in Outlook, Office apps, SharePoint, OneDrive or Teams when what they are doing may break a data protection rule. It explains the issue and what to do.||**Block with override** stops an action but lets the user continue if they have a valid business reason, usually by entering a justification that is recorded for review.' -Why 'Good policy tips educate users and prevent mistakes without stopping legitimate work. Overrides allow controlled exceptions for genuine business processes, while keeping an audit trail of who overrode a policy and why.'
    $tips = @($ruleSum | Where-Object { $_.NotifyUser })
    $customTip = @($ruleSum | Where-Object { $_.NotifyPolicyTipCustomText })
    $ovr = @($ruleSum | Where-Object { $_.NotifyAllowOverride })
    Add-DocBullet "**$($tips.Count)** of $($rules.Count) rules notify users."
    Add-DocBullet "**$($customTip.Count)** rules use custom policy tip text; the rest use Microsoft's default wording."
    Add-DocBullet "**$($ovr.Count)** rules allow users to override with a justification."
    if ($customTip.Count -gt 0) {
        $rows = foreach ($r in $customTip) { , @([string]$r.Rule, [string]$r.NotifyPolicyTipCustomText) }
        Add-DocTable -Headers @('Rule', 'Custom policy tip text') -Widths @(2, 5) -Rows @($rows) -MaxRows 20
    }
    if ($ovr.Count -gt 0) {
        $rows = foreach ($r in $ovr) { , @([string]$r.Policy, [string]$r.Rule, [string]$r.NotifyAllowOverride) }
        Add-DocTable -Headers @('Policy', 'Rule', 'Override requirement') -Widths @(2.5, 2.5, 2.5) -Rows @($rows) -MaxRows 30
    }
    $cons = @()
    if ($tips.Count -gt 0 -and $customTip.Count -eq 0) { $cons += 'All policy tips use the default wording. Short, organisation-specific text (what is wrong, what to do, who to contact) is usually better understood.' }
    if ($ovr.Count -gt 0) { $cons += Add-Obs 'User experience' 'Low' "$(Get-Plural $ovr.Count 'rule allows' 'rules allow') overrides. Agree who reviews override justifications and how often, otherwise overrides can become a silent bypass." }
    if ($tips.Count -eq 0 -and $rules.Count -gt 0) { $cons += Add-Obs 'User experience' 'Medium' 'No rules notify users. Users are not being told when they handle sensitive information inappropriately, which removes an important education opportunity.' }
    $Script:AreaStatus['UX'] = $(if ($rules.Count -eq 0 -or $tips.Count -eq 0) { 'R' } elseif ($customTip.Count -eq 0) { 'A' } else { 'G' })
    Add-SectionEnd $cons @('PolicyTips', 'DlpReference') -Area 'UX'

    # ---------- 7 SITs ----------
    Add-DocHeading '7. Sensitive information types' 1
    Add-SectionIntro -What 'A **sensitive information type (SIT)** is a detector that recognises a specific kind of sensitive data, such as a credit card number, a UK National Insurance number, a passport number or a bank account number. Microsoft provides several hundred built-in types. Each uses patterns, keywords and checksums to decide whether something genuinely is sensitive.' -Why 'DLP can only protect information it can recognise. Choosing the right SITs, the ones that match the organisation''s real data, is the foundation of accurate DLP with few false alarms.'
    Add-SectionGap 'Sensitive information types'
    $usedNames = @($classRefs | Where-Object { $_.ClassifierType -eq 'Content (SIT)' } | Select-Object -ExpandProperty Name -Unique)
    Add-DocPara "$(Get-Plural $sitView.Count 'sensitive information type is' 'sensitive information types are') available in the tenant ($($sitView.Count - $custom.Count) built-in, $($custom.Count) custom). DLP rules reference **$(Get-Plural $usedNames.Count 'distinct type' 'distinct types')**."
    if ($usedNames.Count -gt 0) {
        $rows = foreach ($g in ($classRefs | Group-Object Name | Sort-Object Count -Descending)) {
            , @([string]$g.Name, [string]$g.Count, ((@($g.Group.Policy | Sort-Object -Unique)) -join ', '))
        }
        Add-DocTable -Headers @('Sensitive information type / classifier', 'Rule references', 'Used in policies') -Widths @(3, 1, 4) -Rows @($rows) -MaxRows 40 -Note 'Detail: 15_DLP_Rule_Classifiers_and_Confidence.'
    }
    $cons = @()
    if ($usedNames.Count -eq 0 -and $rules.Count -gt 0) { $cons += 'No sensitive information types are referenced in DLP rules; rules may rely only on labels or other conditions.' }
    $cons += 'Compare the types in use with the organisation''s data inventory. Common gaps include national identifiers, bank account and payment details, internal customer or account references and staff identifiers.'
    $Script:AreaStatus['SIT'] = $(if ($usedNames.Count -gt 0) { 'G' } else { 'A' })
    Add-SectionEnd $cons @('SitOverview', 'SitDefinitions') -Area 'SIT'

    # ---------- 8 Custom SITs ----------
    Add-DocHeading '8. Custom sensitive information types and keyword dictionaries' 1
    Add-SectionIntro -What 'A **custom SIT** detects information in an organisation-specific format, such as an internal customer number, a project code or an employee ID. A **keyword dictionary** is a large list of words or phrases (for example code names or product names) that can be used as, or within, a custom SIT.' -Why 'Built-in detectors cannot recognise identifiers unique to the organisation. Custom SITs close that gap. Unused or poorly tuned custom SITs add complexity and can cause false positives.'
    Add-SectionGap 'Sensitive information types', 'SIT rule packages', 'Keyword dictionaries'
    if ($custom.Count -eq 0) { Add-DocPara 'No custom sensitive information types were found.' }
    else {
        $rows = foreach ($s in $custom) { , @([string]$s.Name, [string]$s.Publisher, [string]$s.RecommendedConfidence, $(if ($s.UsedInDlpRules) { 'Yes' } else { 'No' })) }
        Add-DocTable -Headers @('Custom SIT', 'Publisher', 'Recommended confidence', 'Used in DLP') -Widths @(3.5, 2, 1.5, 1) -Rows @($rows) -MaxRows 40 -Note 'Definitions of custom rule packages are saved as 22_RulePackage_*.xml.'
    }
    Add-DocPara "$(Get-Plural $dicts.Count 'keyword dictionary is' 'keyword dictionaries are') configured."
    $cons = @()
    $unused = @($custom | Where-Object { -not $_.UsedInDlpRules })
    if ($unused.Count -gt 0) { $cons += Add-Obs 'Custom SITs' 'Low' "$(Get-Plural $unused.Count 'custom SIT is' 'custom SITs are') not used by any DLP rule: either a policy is missing or the SIT is obsolete." }
    if ($custom.Count -eq 0) { $cons += 'Identify organisation-specific identifiers (customer, account, policy or staff numbers) that need protecting and consider custom SITs for them.' }
    $Script:AreaStatus['CustomSIT'] = $(if ($custom.Count -gt 0 -and $unused.Count -eq 0) { 'G' } else { 'A' })
    Add-SectionEnd $cons @('SitCustom', 'KeywordDict') -Area 'CustomSIT'

    # ---------- 9 EDM ----------
    Add-DocHeading '9. Exact Data Match (EDM)' 1
    Add-SectionIntro -What '**Exact Data Match** detects real records from an approved dataset, for example the actual customer list or employee database, rather than anything that merely looks like sensitive data. The data is uploaded as secure hashes and matched against content.' -Why 'Pattern-based detection can mistake random numbers for sensitive data. EDM matches only genuine records, which greatly reduces false positives for the organisation''s most important structured data.'
    Add-SectionGap 'Exact Data Match schemas'
    $edmRefs = @($classRefs | Where-Object { $_.ClassifierType -match 'ExactMatch|EDM' })
    Add-DocPara "$(Get-Plural $edm.Count 'EDM schema is' 'EDM schemas are') configured, with $(Get-Plural $edmSits.Count 'EDM-based SIT' 'EDM-based SITs'). EDM classifiers are referenced $(Get-Plural $edmRefs.Count 'time' 'times') in DLP rules."
    if ($edmSits.Count -gt 0) { Add-DocTable -Headers @('EDM sensitive information type', 'Used in DLP') -Widths @(5, 1) -Rows @(foreach ($s in $edmSits) { , @([string]$s.Name, $(if ($s.UsedInDlpRules) { 'Yes' } else { 'No' })) }) }
    $cons = @()
    if ($edm.Count -eq 0) { $cons += 'EDM is not in use. Consider it for well-defined structured datasets such as customer or employee records, where accuracy matters most. Agree data ownership and refresh frequency first.' }
    $Script:AreaStatus['EDM'] = $(if ($edm.Count -gt 0) { 'G' } else { 'A' })
    Add-SectionEnd $cons @('EdmLearn', 'EdmStart') -Area 'EDM'

    # ---------- 10 Trainable classifiers ----------
    Add-DocHeading '10. Trainable classifiers' 1
    Add-SectionIntro -What 'A **trainable classifier** uses machine learning to recognise the overall type of a document, such as a financial statement, contract, CV or source code, rather than looking for specific numbers or words. Microsoft provides ready-made classifiers and organisations can train their own.' -Why 'Some sensitive documents contain no easily detectable numbers. Classifiers help protect content by its nature, which complements sensitive information types.'
    $tc = @($classRefs | Where-Object { $_.ClassifierType -match 'MLModel|Trainable' })
    if ($tc.Count -eq 0) { Add-DocPara 'No trainable classifiers are referenced in DLP rules.' }
    else { Add-DocTable -Headers @('Classifier', 'Policy', 'Rule') -Widths @(3, 2.5, 2.5) -Rows @(foreach ($t in $tc) { , @([string]$t.Name, [string]$t.Policy, [string]$t.Rule) }) -MaxRows 30 }
    Add-DocCallout 'Note:' 'There is no supported PowerShell command to list every classifier available in the tenant. Review the full list in the Microsoft Purview portal under Data classification > Classifiers.'
    $cons = @('Test relevant classifiers against representative documents before using them to block, as accuracy varies by content type.')
    $Script:AreaStatus['Classifiers'] = $(if ($tc.Count -gt 0) { 'G' } else { 'A' })
    Add-SectionEnd $cons @('Classifiers') -Area 'Classifiers'

    # ---------- 11 Confidence ----------
    Add-DocHeading '11. Confidence levels and instance counts' 1
    Add-SectionIntro -What 'Each detection has a **confidence level** (High, Medium or Low) showing how certain Purview is that a match is genuine. High confidence needs more supporting evidence, such as nearby keywords or a valid checksum. Rules also set an **instance count**: how many matches must appear before the rule triggers (for example 1 credit card number versus 10 or more).' -Why 'These two settings have the biggest effect on false positives. Low confidence or very low counts can flag ordinary content; very high thresholds can miss genuine leaks. Typical practice is to use higher counts and confidence for blocking rules and lower values for notify-only rules.'
    $conf = @($Script:Cache['ConfidenceSummary'])
    if ($conf.Count -gt 0) { Add-DocTable -Headers @('Confidence', 'References', 'Rules') -Widths @(1.5, 1, 5) -Rows @(foreach ($c in $conf) { , @([string]$c.Confidence, [string]$c.References, [string]$c.Rules) }) -MaxRows 20 }
    $lowConf = @($classRefs | Where-Object { $_.ConfidenceLevel -eq 'Low' -or ($_.MinConfidence -and [int]('0' + $_.MinConfidence) -lt 75 -and [int]('0' + $_.MinConfidence) -gt 0) })
    $single = @($classRefs | Where-Object { [string]$_.MinCount -eq '1' })
    Add-DocPara "$(Get-Plural $lowConf.Count 'reference uses' 'references use') low confidence, and $(Get-Plural $single.Count 'reference triggers' 'references trigger') on a single instance."
    $cons = @()
    if ($lowConf.Count -gt 0) { $cons += Add-Obs 'Confidence' 'Medium' "$(Get-Plural $lowConf.Count 'detection uses' 'detections use') low confidence. Check these against false-positive reports, especially in blocking rules." }
    $blockRules = @($ruleSum | Where-Object { $_.ResponseInterpretation -like 'Block*' } | Select-Object -ExpandProperty Rule)
    $blockSingle = @($single | Where-Object { $blockRules -contains $_.Rule })
    if ($blockSingle.Count -gt 0) { $cons += "$(Get-Plural $blockSingle.Count 'blocking detection triggers' 'blocking detections trigger') on a single instance. Confirm this is intended; a single match is more likely to be a false positive." }
    $Script:AreaStatus['Confidence'] = $(if ($classRefs.Count -eq 0) { 'N' } elseif ($lowConf.Count -gt 0) { 'A' } else { 'G' })
    Add-SectionEnd $cons @('SitOverview', 'SitDefinitions') -Area 'Confidence'

    # ---------- 12 Labels ----------
    Add-DocHeading '12. Sensitivity labels and DLP' 1
    Add-SectionIntro -What 'A **sensitivity label** (for example Public, General, Confidential, Highly Confidential) describes how sensitive a document or email is. Labels can add headers and watermarks and can encrypt content. **Label policies** publish labels to users; **auto-labelling** applies labels automatically when content matches conditions.||Labels and DLP work together: the label says **what** something is, and DLP controls **what people can do** with it. DLP rules can use labels as conditions, for example "block external sharing of anything labelled Highly Confidential".' -Why 'Labels give a consistent, user-visible classification that travels with the content. Using them in DLP gives simple, explainable rules and complements pattern-based detection.'
    Add-SectionGap 'Sensitivity labels', 'Sensitivity label policies', 'Auto-labelling policies', 'Auto-labelling rules'
    if ($labels.Count -eq 0) { Add-DocPara 'No sensitivity labels were found.' }
    else {
        $rows = foreach ($l in ($labels | Sort-Object { [int]('0' + $_.Priority) })) {
            $nm = if ($l.DisplayName) { [string]$l.DisplayName } else { [string]$l.Name }
            if ($l.ParentId) { $nm = '   - ' + $nm }
            , @($nm, [string]$l.ContentType, $(if ($l.EncryptionEnabled) { 'Yes' } else { 'No' }), $(if ($l.UsedInDlpRules) { 'Yes' } else { 'No' }))
        }
        Add-DocTable -Headers @('Label (in priority order)', 'Applies to', 'Encrypts', 'Used in DLP') -Widths @(3, 3, 1, 1) -Rows @($rows) -MaxRows 50 -Note 'Sublabels are indented. Detail: 30_Sensitivity_Labels.'
    }
    Add-DocPara "$(Get-Plural $labelPol.Count 'label policy publishes' 'label policies publish') labels to users. $(Get-Plural $autoPol.Count 'auto-labelling policy is' 'auto-labelling policies are') configured. Labels are used as conditions in DLP rules $(Get-Plural $labelRefs.Count 'time' 'times')."
    $cons = @()
    if ($labels.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'High' 'No sensitivity labels exist. Labels are a foundation for information protection; consider a simple taxonomy (four or five labels) before relying on label-based DLP.' }
    elseif ($labelRefs.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'Medium' 'Sensitivity labels exist but are not used in any DLP rule. Using them as conditions (for example blocking external sharing of Highly Confidential content) is a simple, explainable control.' }
    if ($labels.Count -gt 0 -and $labelPol.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'High' 'Labels exist but no label policy publishes them, so users cannot apply them.' }
    if ($labels.Count -gt 0 -and $autoPol.Count -eq 0 -and -not (Get-StepGap 'Auto-labelling policies')) { $cons += 'No auto-labelling policies are configured. Relying entirely on users to label content usually leads to inconsistent labelling.' }
    $Script:AreaStatus['Labels'] = $(if ($labelRefs.Count -gt 0 -and $labelPol.Count -gt 0) { 'G' } elseif ($labels.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('Labels', 'AutoLabel', 'LabelCondition') -Area 'Labels'

    # ---------- 13 Endpoint ----------
    Add-DocHeading '13. Endpoint DLP (devices)' 1
    Add-SectionIntro -What '**Endpoint DLP** extends protection to Windows and macOS devices that are onboarded to Microsoft Purview. It can audit, warn or block activities on sensitive files, including:||**USB and removable storage**: copying files to USB drives and other removable media. **Printing**: sending sensitive files to printers. **Browser and cloud uploads**: uploading files to websites and unapproved cloud services, or using unallowed browsers. Also copying to network shares, the clipboard, Bluetooth and unallowed apps.||Tenant-wide **Endpoint DLP settings** define approved and blocked service domains, browsers, printer groups, removable media groups and file path exclusions that all device policies use.' -Why 'Cloud protection stops once a user downloads a file. Endpoint DLP keeps protection in place on the device, covering the most common ways data leaves managed systems: USB drives, printouts and personal cloud storage.'
    Add-SectionGap 'Endpoint DLP / tenant policy configuration'
    Add-DocPara "$(Get-Plural $covCount['Endpoint'] 'enforced policy covers' 'enforced policies cover') devices. $(Get-Plural $epRules.Count 'device restriction is' 'device restrictions are') defined in rules."
    $usb = @($epRules | Where-Object { $_.Setting -match 'RemovableMedia|Usb' })
    $prn = @($epRules | Where-Object { $_.Setting -match 'Print' })
    $upl = @($epRules | Where-Object { $_.Setting -match 'Upload|Browser|Cloud|Egress' })
    $oth = @($epRules | Where-Object { $usb -notcontains $_ -and $prn -notcontains $_ -and $upl -notcontains $_ })
    Add-DocTable -Headers @('Activity', 'Restrictions in rules', 'Actions used') -Widths @(2.5, 1.2, 4) -Rows @(
        , @('USB / removable storage', [string]$usb.Count, ((@($usb.Value | Sort-Object -Unique)) -join ', '))
        , @('Printing', [string]$prn.Count, ((@($prn.Value | Sort-Object -Unique)) -join ', '))
        , @('Browser and cloud uploads', [string]$upl.Count, ((@($upl.Value | Sort-Object -Unique)) -join ', '))
        , @('Other (clipboard, network share, Bluetooth, apps)', [string]$oth.Count, ((@($oth.Value | Sort-Object -Unique)) -join ', '))
    ) -Note 'Detail: 16_Endpoint_DLP_Rule_Restrictions.'
    $setRows = @($epSet | Where-Object { $_.Value } | ForEach-Object { , @([string]$_.Area, [string]$_.Setting, [string]$_.Value) })
    if ($setRows.Count -gt 0) {
        Add-DocHeading 'Tenant-wide Endpoint DLP settings' 3
        Add-DocTable -Headers @('Area', 'Setting', 'Value') -Widths @(2.5, 2.5, 4) -Rows $setRows -MaxRows 40 -Note 'Detail: 41_Endpoint_DLP_Settings.'
    }
    Add-DocCallout 'Note:' 'Device onboarding status is not available through PowerShell. Check onboarded devices in the Purview portal (Settings > Device onboarding) or in Microsoft Defender.'
    $cons = @()
    if ($covCount['Endpoint'] -eq 0) { $cons += Add-Obs 'Endpoint DLP' 'High' 'No enforced policy covers devices. Files downloaded from Microsoft 365 can be copied to USB, printed or uploaded elsewhere without any DLP check.' }
    else {
        if ($usb.Count -eq 0) { $cons += Add-Obs 'Endpoint DLP' 'Medium' 'Device policies exist but none restrict copying to USB or removable storage.' }
        if ($prn.Count -eq 0) { $cons += 'No rule restricts printing of sensitive files. Identify business processes that genuinely require printing before adding controls.' }
        if ($upl.Count -eq 0) { $cons += 'No rule restricts browser or cloud uploads. Define approved service domains and browsers, then consider warning or blocking uploads elsewhere.' }
    }
    $Script:AreaStatus['Endpoint'] = $(if ($covCount['Endpoint'] -gt 0 -and $usb.Count -gt 0) { 'G' } elseif ($covCount['Endpoint'] -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('EndpointLearn', 'EndpointSettings', 'EndpointUsing') -Area 'Endpoint'

    # ---------- 14 Copilot ----------
    Add-DocHeading '14. DLP for Microsoft 365 Copilot' 1
    Add-SectionIntro -What 'Purview DLP can apply to **Microsoft 365 Copilot and Copilot Chat**. Policies for this location can stop Copilot using content with specific sensitivity labels in its responses, and can restrict prompts containing sensitive information.' -Why 'AI assistants create new ways for people to find and reuse sensitive information. Applying DLP to Copilot keeps the organisation''s most sensitive content out of AI-generated answers where it should not appear.'
    Add-DocPara "$(Get-Plural $covCount['Copilot'] 'enforced policy covers' 'enforced policies cover') Microsoft 365 Copilot."
    $cp = @($cov | Where-Object { $_.Copilot -eq $true })
    if ($cp.Count -gt 0) { Add-DocTable -Headers @('Policy', 'Mode') -Widths @(5, 2) -Rows @(foreach ($p in $cp) { , @([string]$p.Policy, [string]$p.Mode) }) }
    $cons = @()
    if ($cp.Count -eq 0) { $cons += Add-Obs 'Copilot' 'Medium' 'No DLP policy covers Microsoft 365 Copilot. If Copilot is deployed or planned, consider a policy that excludes highly confidential labelled content from Copilot responses.' }
    $cons += 'Copilot location detection uses pattern matching on policy locations. If a Copilot policy is missing here, confirm in the Purview portal.'
    $Script:AreaStatus['Copilot'] = $(if ($covCount['Copilot'] -gt 0) { 'G' } elseif ($cp.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('Copilot') -Area 'Copilot'

    # ---------- 15 Deployment ----------
    Add-DocHeading '15. Deployment state and simulation' 1
    Add-SectionIntro -What '**Simulation mode** runs a policy as if it were enforced, but without blocking anyone. Matches are shown in a separate dashboard so the policy can be tuned. Best practice is to move policies through discovery, design, simulation, tuning, a pilot group and then progressive enforcement.' -Why 'Enforcing an untested policy can disrupt business processes and generate many false positives. Simulation and phased rollout reduce disruption and build confidence with stakeholders.'
    Add-DocTable -Headers @('State', 'Policies') -Widths @(4, 1) -Rows @(
        , @('Enforced', [string]$pEnf)
        , @('Simulation with policy tips', [string]@($policies | Where-Object { $_.Mode -eq 'TestWithNotifications' }).Count)
        , @('Simulation without policy tips', [string]@($policies | Where-Object { $_.Mode -eq 'TestWithoutNotifications' }).Count)
        , @('Disabled', [string]$pOff)
    )
    $cons = @('Agree policy waves, pilot groups, acceptance criteria (for example an acceptable false-positive rate) and enforcement dates for each policy.')
    $Script:AreaStatus['Simulation'] = $(if ($pEnf -gt 0) { 'G' } elseif ($policies.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('DlpSimulation', 'DlpSimStart', 'DlpDeploy') -Area 'Simulation'

    # ---------- 16 Alerts ----------
    Add-DocHeading '16. Alerts and monitoring' 1
    Add-SectionIntro -What 'DLP can raise an **alert** when an important or high-risk event happens, and send **incident reports** by email. Alerts appear in the DLP alerts dashboard and in Microsoft Defender, where they can be investigated and resolved.' -Why 'Alerts only add value when someone owns, reviews and responds to them. Too many alerts cause fatigue; too few mean real incidents are missed.'
    Add-SectionGap 'Alert policies'
    $ruleAlert = @($ruleSum | Where-Object { $_.GenerateAlert -or $_.GenerateIncidentReport })
    Add-DocPara "$(Get-Plural $ruleAlert.Count 'rule raises' 'rules raise') alerts or incident reports. $(Get-Plural $alertPol.Count 'DLP-related alert policy exists' 'DLP-related alert policies exist')."
    if ($ruleAlert.Count -gt 0) {
        Add-DocTable -Headers @('Policy', 'Rule', 'Severity', 'Alert / report recipients') -Widths @(2.5, 2.5, 1, 3) -Rows @(foreach ($r in $ruleAlert) { , @([string]$r.Policy, [string]$r.Rule, [string]$r.ReportSeverityLevel, ((@([string]$r.GenerateAlert, [string]$r.GenerateIncidentReport) | Where-Object { $_ } | Select-Object -Unique) -join ' | ')) }) -MaxRows 40
    }
    $cons = @()
    if ($ruleAlert.Count -eq 0 -and $rules.Count -gt 0) { $cons += Add-Obs 'Alerts' 'High' 'No DLP rules raise alerts, so nobody is notified of policy matches.' }
    $siteOnly = @($ruleAlert | Where-Object { ([string]$_.GenerateAlert + [string]$_.GenerateIncidentReport) -match '^(SiteAdmin\s*\|?\s*)+$' })
    if ($siteOnly.Count -gt 0) { $cons += Add-Obs 'Alerts' 'Medium' "$(Get-Plural $siteOnly.Count 'rule sends' 'rules send') alerts only to SiteAdmin. Route alerts to a named team or shared mailbox with a defined triage process." }
    if ($Script:Cache.ContainsKey('GraphAlerts')) {
        $ga = @($Script:Cache['GraphAlerts'])
        Add-DocHeading 'DLP alerts raised in the last 90 days (Microsoft Graph)' 3
        if ($ga.Count -eq 0) { Add-DocPara 'No DLP alerts were raised in the last 90 days.' }
        else {
            $open = @($ga | Where-Object { $_.Status -match 'new|inProgress' })
            $fp = @($ga | Where-Object { $_.Classification -eq 'falsePositive' })
            $classified = @($ga | Where-Object { $_.Classification -and $_.Classification -ne 'unknown' })
            Add-DocMetrics -Items @(
                @{ Value = $ga.Count; Label = 'DLP alerts (90 days)' }
                @{ Value = @($ga | Where-Object { $_.Severity -eq 'high' }).Count; Label = 'High severity'; Color = 'B42318' }
                @{ Value = $open.Count; Label = 'Still open'; Color = 'B26B00' }
                @{ Value = $fp.Count; Label = 'Marked false positive'; Color = '8A8A8A' }
            )
            $bySev = @($ga | Group-Object Severity | Sort-Object Count -Descending | ForEach-Object { , @([string]$_.Name, [string]$_.Count, [string]@($_.Group | Where-Object { $_.Status -match 'new|inProgress' }).Count) })
            Add-DocTable -Headers @('Severity', 'Alerts', 'Still open') -Widths @(3, 1.5, 1.5) -Rows $bySev
            $topTitles = @($ga | Group-Object Title | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { , @([string]$_.Name, [string]$_.Count) })
            Add-DocTable -Headers @('Most frequent alerts', 'Count') -Widths @(6, 1) -Rows $topTitles -Note 'Detail: 82_Graph_DLP_Alerts and 83_Graph_DLP_Alert_Summary.'
            if ($ga.Count -ge 10 -and ($open.Count / $ga.Count) -gt 0.5) { $cons += Add-Obs 'Alerts' 'Medium' "$($open.Count) of $($ga.Count) DLP alerts from the last 90 days are still open. Alerts are being raised but not consistently triaged." }
            if ($classified.Count -ge 10 -and ($fp.Count / $classified.Count) -gt 0.3) { $cons += Add-Obs 'Alerts' 'Medium' "$([math]::Round(100 * $fp.Count / $classified.Count))% of classified DLP alerts were false positives. Tune the rules behind the most frequent alerts (confidence levels, instance counts, exceptions)." }
            if ($classified.Count -eq 0) { $cons += 'No DLP alerts have been classified (true or false positive). Classifying alerts gives the evidence needed to tune rules.' }
        }
    }
    $cons += 'Define alert severity, recipients, investigation steps and escalation routes, and who reviews override justifications.'
    $Script:AreaStatus['Alerts'] = $(if ($ruleAlert.Count -gt 0 -and $siteOnly.Count -eq 0) { 'G' } elseif ($ruleAlert.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('Alerts') -Area 'Alerts'

    # ---------- 17 Activity Explorer ----------
    Add-DocHeading '17. Activity Explorer' 1
    Add-SectionIntro -What '**Activity Explorer** shows how sensitive and labelled information is being used: DLP rule matches, labels applied or changed, files copied to USB, printed, uploaded and more. It is the main source of evidence for policy tuning and investigations.' -Why 'Real activity data shows whether policies are working, where false positives occur, and which workloads and users generate the most risk.'
    $act = @(if ($Script:Cache.ContainsKey('Activity')) { $Script:Cache['Activity'] } else { @() })
    if (-not $Script:Cache.ContainsKey('Activity')) { Add-DocPara 'Activity Explorer data was not included in this run. Re-run the complete report with Activity Explorer enabled to add this section.' }
    else {
        Add-SectionGap "Activity Explorer ($($Script:Settings.ActivityDays) days)"
        Add-DocPara "$($act.Count) activity records were collected for the last $($Script:Settings.ActivityDays) days."
        $grp = @($act | Group-Object { '{0}|{1}' -f (Get-Val $_ 'Activity'), (Get-Val $_ 'Workload') } | Sort-Object Count -Descending)
        if ($grp.Count) { Add-DocTable -Headers @('Activity', 'Workload', 'Count') -Widths @(4, 2.5, 1) -Rows @(foreach ($g in $grp) { $p = $g.Name -split '\|', 2; , @($p[0], $p[1], [string]$g.Count) }) -MaxRows 25 -Note 'Detail: 60_Activity_Explorer_Records (contains user and file names).' }
    }
    $Script:AreaStatus['Activity'] = $(if ($Script:Cache.ContainsKey('Activity')) { 'G' } else { 'N' })
    Add-SectionEnd @('Agree a regular review of Activity Explorer (for example weekly during rollout, then monthly) and the operational reports stakeholders need.') @('ActivityExplorer') -Area 'Activity'

    # ---------- 18 Exchange ----------
    $cons = @()
    Add-DocHeading '18. Exchange Online mail flow and auditing' 1
    Add-SectionIntro -What '**Mail flow rules** (transport rules) are an older, Exchange-only way of acting on email. Some organisations still use them to detect sensitive information. The **unified audit log** records user and admin activity across Microsoft 365 and is required for DLP reporting and Activity Explorer.' -Why 'Legacy mail flow rules that inspect sensitive information can overlap with, or contradict, Purview DLP policies. Auditing must be on for DLP evidence to exist.'
    if (-not $Script:Cache.ContainsKey('TransportRules')) { Add-DocPara 'Exchange Online data was not included in this run. Re-run the complete report with Exchange Online enabled to add this section.' }
    else {
        $tr = @($Script:Cache['TransportRules'])
        $trSit = @($tr | Where-Object { Get-Val $_ 'MessageContainsDataClassifications' })
        Add-DocPara "$(Get-Plural $tr.Count 'mail flow rule is' 'mail flow rules are') configured; $(Get-Plural $trSit.Count 'inspects' 'inspect') sensitive information types."
        $aud = @($Script:Cache['Audit'])
        if ($aud.Count) { Add-DocPara ("Unified audit log ingestion: **{0}**." -f $(if ($aud[0].UnifiedAuditLogIngestionEnabled) { 'Enabled' } else { 'Disabled' })) }
        $cons = @()
        if ($trSit.Count -gt 0) { $cons += Add-Obs 'Exchange' 'Medium' "$(Get-Plural $trSit.Count 'mail flow rule inspects' 'mail flow rules inspect') sensitive information. Consider migrating this logic to Purview DLP so all data protection is managed in one place." }
        if ($aud.Count -and -not $aud[0].UnifiedAuditLogIngestionEnabled) { $cons += Add-Obs 'Auditing' 'High' 'The unified audit log is disabled. DLP reporting, Activity Explorer and investigations depend on it.' }
    }
    $Script:AreaStatus['Exchange'] = $(if (-not $Script:Cache.ContainsKey('TransportRules')) { 'N' } elseif ($Script:Cache.ContainsKey('Audit') -and @($Script:Cache['Audit']).Count -and -not @($Script:Cache['Audit'])[0].UnifiedAuditLogIngestionEnabled) { 'R' } elseif (@(@($Script:Cache['TransportRules']) | Where-Object { Get-Val $_ 'MessageContainsDataClassifications' }).Count -gt 0) { 'A' } else { 'G' })
    Add-SectionEnd $cons @('MailFlow', 'DlpExchange') -Area 'Exchange'
    $cons = @()

    # ---------- 19 Ownership ----------
    Add-DocHeading '19. Operational ownership' 1
    Add-SectionIntro -What 'Purview uses **role groups** to control who can manage DLP policies, view alerts, investigate activity and change classification settings.' -Why 'DLP needs named owners for policies, alerts, investigations, exceptions and changes. Too many administrators increases risk; too few creates a single point of failure.'
    Add-SectionGap 'Purview role groups', 'Purview role group members'
    $byGroup = @($roleMembers | Group-Object RoleGroup | Sort-Object Name)
    if ($byGroup.Count -gt 0) {
        Add-DocTable -Headers @('Role group', 'Members') -Widths @(2.5, 5) -Rows @(foreach ($g in $byGroup) { , @([string]$g.Name, ((@($g.Group.Member | Sort-Object -Unique)) -join ', ')) }) -MaxRows 40 -Note 'Direct members only; access can also come from Microsoft Entra roles. Detail: 55_DLP_Relevant_Role_Members.'
    }
    $cons = @('Create a responsibility matrix covering security, compliance, IT, legal, HR and business owners: who owns each policy, who triages alerts, who approves exceptions and changes.')
    if ($Script:Cache.ContainsKey('EntraRoles')) {
        $er = @($Script:Cache['EntraRoles'])
        Add-DocHeading 'Microsoft Entra roles with access to DLP (Microsoft Graph)' 3
        Add-DocPara 'These directory roles also give access to Microsoft Purview, in addition to the role groups above. Only active assignments are shown; roles that are eligible through Privileged Identity Management are not included.'
        $byRole = @($er | Group-Object Role | Sort-Object Name | ForEach-Object { , @([string]$_.Name, [string]$_.Count, ((@($_.Group.Principal | Sort-Object -Unique | Select-Object -First 12)) -join ', ')) })
        if ($byRole.Count) { Add-DocTable -Headers @('Entra role', 'Assignments', 'Assigned to (first 12)') -Widths @(2.4, 1, 5) -Rows $byRole -Note 'Detail: 84_Graph_Entra_Role_Holders.' }
        $ga2 = @($er | Where-Object { $_.Role -eq 'Global Administrator' } | Select-Object -ExpandProperty Principal -Unique)
        if ($ga2.Count -gt 4) { $cons += Add-Obs 'Ownership' 'Medium' "$($ga2.Count) principals hold the Global Administrator role, which includes full access to DLP. Keep this role to a small number of accounts and use Compliance Administrator or narrower Purview roles for day-to-day DLP work." }
        if (@($er | Where-Object { $_.Role -match '^Compliance' }).Count -eq 0) { $cons += 'Nobody holds the Compliance Administrator or Compliance Data Administrator role. Confirm who administers DLP day to day and that they do not rely on Global Administrator.' }
    }
    $uniq = @($roleMembers | Select-Object -ExpandProperty Member -Unique)
    if ($uniq.Count -gt 15) { $cons += Add-Obs 'Ownership' 'Low' "$($uniq.Count) people or groups hold DLP-relevant roles. Review whether all still need them." }
    $Script:AreaStatus['Ownership'] = $(if ($roleMembers.Count -gt 0) { 'G' } else { 'N' })
    Add-SectionEnd $cons @('Permissions') -Area 'Ownership'

    # ---------- 20 Licensing ----------
    Add-DocHeading '20. Licensing for DLP features' 1
    Add-SectionIntro -What 'Several DLP capabilities need premium licences, typically Microsoft 365 E5, E5 Compliance or E5 Information Protection and Governance, or specific add-ons. Examples include DLP for Teams, Endpoint DLP, auto-labelling, trainable classifiers, Exact Data Match and Copilot.' -Why 'Licensing decides what can be deployed. A feature that is licensed but not configured is a quick win; a feature that is configured without the right licence is a compliance risk for the organisation.'
    if (-not $Script:Cache.ContainsKey('LicenceIndicators')) {
        Add-DocPara 'Licensing data was not collected in this run. Re-run the complete report with Microsoft Graph enabled (menu Settings > 10, or -IncludeGraph) to add this section.'
        Add-SectionEnd @() @('DlpOverview')
    }
    else {
        Add-SectionGap 'Graph: Licensing'
        $skusAll = @($Script:Cache['Graph'].Skus | Where-Object { $_ -and $_.CapabilityStatus -match 'Enabled|Warning' -and [int]('0' + $_.EnabledUnits) -gt 0 })
        Add-DocPara "$(Get-Plural $skusAll.Count 'active subscription was' 'active subscriptions were') found."
        Add-DocTable -Headers @('Subscription (SKU)', 'Status', 'Assigned', 'Available') -Widths @(4, 1.5, 1, 1) -Rows @($skusAll | Sort-Object SkuPartNumber | ForEach-Object { , @([string]$_.SkuPartNumber, [string]$_.CapabilityStatus, [string]$_.ConsumedUnits, [string]$_.EnabledUnits) }) -MaxRows 40 -Note 'Product names for each SKU are listed in Microsoft''s "Product names and service plan identifiers for licensing". Detail: 80_Graph_Subscriptions.'
        $configured = @{
            'Coverage' = ($covCount['Exchange'] -gt 0 -or $covCount['SharePoint'] -gt 0 -or $covCount['OneDrive'] -gt 0)
            'Endpoint' = ($covCount['Endpoint'] -gt 0)
            'Labels'   = ($autoPol.Count -gt 0)
            'EDM'      = ($edm.Count -gt 0 -or @($classRefs | Where-Object { $_.ClassifierType -match 'MLModel|Trainable' }).Count -gt 0)
            'Copilot'  = ($covCount['Copilot'] -gt 0)
        }
        $teamsConfigured = ($covCount['Teams'] -gt 0)
        $quick = @(); $risk = @()
        $rowsLic = foreach ($i in @($Script:Cache['LicenceIndicators'])) {
            $conf = if ($i.Feature -like '*Teams*') { $teamsConfigured } else { [bool]$configured[$i.Area] }
            $assess = if ($i.Indicator -eq 'Found' -and $conf) { '[[G]]Licensed and in use' }
            elseif ($i.Indicator -eq 'Found') { $quick += $i.Feature; '[[A]]Licensed, not configured' }
            elseif ($conf) { $risk += $i.Feature; '[[R]]In use, licence not detected' }
            else { '[[N]]Not licensed or not detected' }
            , @([string]$i.Feature, $(if ($i.Indicator -eq 'Found') { 'Yes' } else { 'No' }), $(if ($conf) { 'Yes' } else { 'No' }), $assess)
        }
        Add-DocTable -Headers @('Capability', 'Licence indicator', 'Configured', 'Assessment') -Widths @(3.2, 1.3, 1.2, 2.4) -Rows @($rowsLic) -Note 'Indicators are matched on subscription and service plan names (81_Graph_DLP_Licence_Indicators shows the evidence). Confirm entitlement against Microsoft licensing guidance before making commercial decisions.'
        $cons = @()
        if ($quick.Count) { $cons += Add-Obs 'Licensing' 'Medium' "Licensed but not configured: $($quick -join '; '). These are quick wins that need no additional spend." }
        if ($risk.Count) { $cons += Add-Obs 'Licensing' 'Medium' "In use but no matching licence was detected: $($risk -join '; '). Confirm the organisation is correctly licensed for these features." }
        $cons += 'Licence indicators show whether the tenant holds a qualifying licence, not whether every user who is protected by a policy is licensed. Check user-level licensing requirements for each feature.'
        Add-SectionEnd $cons @('DlpOverview')
    }

    # ---------- 21 Next steps ----------
    Add-DocHeading '21. Recommended next steps' 1
    Add-DocPara 'The table below lists suggested activities for each DLP topic, with the automated finding from this assessment. Use it as the starting point for the engagement plan; the same content, with Status, Owner and Target date columns, is in 00_Assessment_Summary.csv.'
    if ($summaryRows.Count) { Add-DocTable -Headers @('Topic', 'Finding', 'Suggested activity') -Widths @(2, 3.5, 3) -FontSize 16 -Rows @(foreach ($s in $summaryRows) { , @([string]$s.Topic, [string]$s.Finding, [string]$s.SuggestedActivity) }) }

    # ---------- 22 Applying the configuration ----------
    $auditOff = [bool]($Script:Cache.ContainsKey('Audit') -and @($Script:Cache['Audit']).Count -and -not @($Script:Cache['Audit'])[0].UnifiedAuditLogIngestionEnabled)
    $ctx = @{
        Gaps = @($gaps); Policies = $policies.Count; Test = $pTest; AuditOnly = $respCount['Audit only']; Rules = $rules.Count; Tips = $tips.Count
        Labels = $labels.Count; LabelPolicies = $labelPol.Count; LabelRefs = $labelRefs.Count; EndpointCov = $covCount['Endpoint']; Usb = $usb.Count
        CopilotCov = $covCount['Copilot']; AlertRules = $ruleAlert.Count; SiteOnly = $siteOnly.Count; LowConf = $lowConf.Count; Custom = $custom.Count
        Edm = $edm.Count; AuditOff = $auditOff
    }
    Add-ImplementationGuidance -Ctx $ctx

    # ---------- Appendices ----------
    Add-DocHeading 'Appendix A: Collection details' 1
    Add-DocTable -Headers @('Item', 'Value') -Widths @(2, 5) -Rows @(
        , @('Organisation', $org)
        , @('Tenant ID', $tid)
        , @('Collected by', $by)
        , @('Collected', $when)
        , @('ExchangeOnlineManagement', $(if ($ti.Count) { [string]$ti[0].ModuleVersion } else { '' }))
        , @('PowerShell', $(if ($ti.Count) { [string]$ti[0].PowerShell } else { '' }))
        , @('Activity Explorer included', [string]$Script:Cache.ContainsKey('Activity'))
        , @('Exchange Online included', [string]$Script:Cache.ContainsKey('TransportRules'))
        , @('Microsoft Graph included', [string]$Script:Cache.ContainsKey('Graph'))
    )
    if ($Script:Errors.Count -gt 0) {
        Add-DocHeading 'Areas not fully collected' 2
        Add-DocPara 'These steps were skipped or failed, usually because the account used did not have the required role or the tenant is not licensed for the feature. Related findings are incomplete.'
        Add-DocTable -Headers @('Step', 'Outcome', 'Reason') -Widths @(2, 1.6, 4.5) -Rows @(foreach ($e in $Script:Errors) { , @([string]$e.Step, [string]$e.Outcome, [string]$e.Error) })
    }

    Add-DocHeading 'Appendix B: Glossary' 1
    Add-DocTable -Headers @('Term', 'Meaning') -Widths @(2, 6) -Rows @(
        , @('DLP', 'Data loss prevention: detecting and controlling the sharing or movement of sensitive information.')
        , @('Policy', 'A container of DLP rules applied to one or more locations.')
        , @('Rule', 'Conditions, exceptions and actions inside a policy.')
        , @('Location', 'Where a policy applies: Exchange, SharePoint, OneDrive, Teams, devices, Copilot and others.')
        , @('SIT', 'Sensitive information type: a detector for a kind of sensitive data, such as credit card numbers.')
        , @('EDM', 'Exact Data Match: detection of real records from an approved, hashed dataset.')
        , @('Trainable classifier', 'A machine-learning model that recognises a type of document.')
        , @('Confidence level', 'How certain Purview is that a match is genuine (High, Medium, Low).')
        , @('Instance count', 'How many matches must be present before a rule triggers.')
        , @('Sensitivity label', 'A classification applied to content, which can add markings and encryption.')
        , @('Auto-labelling', 'Automatically applying a sensitivity label when content matches conditions.')
        , @('Policy tip', 'A message shown to users when an action may break a DLP rule.')
        , @('Override', 'Allowing a user to continue a blocked action with a business justification.')
        , @('Simulation mode', 'Running a policy without enforcement to see what it would match.')
        , @('Endpoint DLP', 'DLP for activities on onboarded Windows and macOS devices.')
        , @('Activity Explorer', 'A Purview view of activity on sensitive and labelled content.')
    )

    Add-DocHeading 'Appendix C: Evidence files' 1
    Add-DocPara "The full data behind this report is in the same folder: $Script:OutputDir"
    Add-DocTable -Headers @('File', 'Contents') -Widths @(3.5, 4.5) -FontSize 16 -Rows @(
        , @('00_Assessment_Summary.csv / .html', 'Topic-by-topic findings and tracker columns')
        , @('DLP_Configuration.xlsx', 'Every dataset as a worksheet (if ImportExcel is installed)')
        , @('11_DLP_Policies_Summary', 'Policies, modes and locations')
        , @('12_DLP_Workload_Coverage', 'Policy-by-location matrix')
        , @('14_DLP_Rules_Conditions_Actions', 'Rules with conditions, exceptions and actions')
        , @('15_DLP_Rule_Classifiers_and_Confidence', 'Classifiers and labels used in each rule')
        , @('20 / 21_Sensitive_Information_Types', 'All and custom sensitive information types')
        , @('30-34_Sensitivity_Labels...', 'Labels, label policies, auto-labelling')
        , @('16 / 41_Endpoint_DLP...', 'Device restrictions and Endpoint DLP settings')
        , @('50-52_Alert...', 'Alert policies and rule alerting')
        , @('53-55_Purview_Role...', 'Role groups and members')
        , @('80-85_Graph_...', 'Subscriptions, licence indicators, DLP alerts, Entra role holders, SharePoint settings (Microsoft Graph)')
        , @('99_Collection_Errors', 'Steps skipped or failed')
    )

    Add-FrameworkAppendix

    # ---------- insert executive observations ----------
    $obsXml = [System.Text.StringBuilder]::new()
    $saved = $Script:Doc.Body; $Script:Doc.Body = $obsXml
    Add-DocHeading 'Key observations' 2
    $order = @{ 'High' = 0; 'Medium' = 1; 'Low' = 2; 'Info' = 3 }
    $obs = @($Script:ReportObs | Sort-Object { $order[$_.Severity] })
    if ($obs.Count -eq 0) { Add-DocPara 'No significant issues were identified automatically. Review each section with the stakeholders to confirm the configuration meets requirements.' }
    else {
        Add-DocTable -Headers @('Priority', 'Area', 'Observation') -Widths @(1, 1.6, 6) -Rows @(foreach ($o in $obs) { , @($o.Severity, $o.Area, $o.Text) })
        Add-DocPara 'Priorities are a starting point generated from the configuration; agree final priorities with the organisation based on its risks and requirements.' -Size 17 -Color '666666' -Italic
    }
    if ($Script:Errors.Count -gt 0) {
        Add-DocCallout 'Collection gaps:' "$(Get-Plural $Script:Errors.Count 'area was' 'areas were') not fully collected (see Appendix A). Findings for those areas may be incomplete." 'FFF4E5' 'E08A00'
    }
    Add-DocHeading 'Status by area' 2
    $statusRows = @(
        , @('DLP policies', $(if ($pEnf -gt 0) { 'Configured' } elseif ($policies.Count -gt 0) { 'Partial (no enforced policies)' } else { 'Not configured' }))
        , @('Email (Exchange)', $(if ($covCount['Exchange']) { 'Protected' } else { 'Not protected' }))
        , @('SharePoint / OneDrive', $(if ($covCount['SharePoint'] -and $covCount['OneDrive']) { 'Protected' } elseif ($covCount['SharePoint'] -or $covCount['OneDrive']) { 'Partial' } else { 'Not protected' }))
        , @('Teams', $(if ($covCount['Teams']) { 'Protected' } else { 'Not protected' }))
        , @('Devices (Endpoint DLP)', $(if ($covCount['Endpoint']) { 'Protected' } else { 'Not protected' }))
        , @('Microsoft 365 Copilot', $(if ($covCount['Copilot']) { 'Protected' } else { 'Not protected' }))
        , @('Custom detection (custom SITs / EDM)', $(if ($custom.Count -or $edm.Count) { 'In use' } else { 'Not in use' }))
        , @('Sensitivity labels in DLP', $(if ($labelRefs.Count) { 'In use' } elseif ($labels.Count) { 'Labels exist, not used in DLP' } else { 'No labels' }))
        , @('Alerting', $(if (@($ruleSum | Where-Object { $_.GenerateAlert -or $_.GenerateIncidentReport }).Count) { 'Configured' } else { 'Not configured' }))
    )
    Add-DocTable -Headers @('Area', 'Status') -Widths @(3, 3) -Rows $statusRows
    if (@(Get-SelectedFrameworks).Count -gt 0) {
        Add-DocHeading 'Framework alignment at a glance' 2
        Add-DocPara 'How each area of the DLP configuration supports recognised security frameworks. Each area section repeats its row with a plain-English description of what the framework expects; Appendix D explains the frameworks and status values.'
        Add-FrameworkMatrix
        Add-DocPara 'Indicative mapping, not a compliance assessment. Section 22 gives guidance for closing the gaps.' -Size 16 -Color '666666' -Italic
    }
    $Script:Doc.Body = $saved
    [void]$Script:Doc.Body.Insert($execObsAnchor, $obsXml.ToString())

    $path = Join-Path $Script:OutputDir 'DLP_Assessment_Report.docx'
    Save-WordDoc -Path $path -Title "DLP Assessment - $org" -Author $(if ($by) { $by } else { 'Get-M365DlpConfiguration' })
    Write-Log "Word report: $path" 'OK'
    return $path
}

#endregion

#region ======================= Summary, workbook, run wrappers ========================

function Export-AssessmentSummary {
    $policies = @(Get-Policies); $rules = @(Get-Rules); $ruleSum = @(Get-RuleSummary); $cov = @(Get-Coverage)
    $rc = @(Get-RuleClassifiers); $sitView = @(Get-SITView); $custom = @($sitView | Where-Object IsCustom)
    $edmSits = @($sitView | Where-Object { $_.Classifier -match 'ExactMatch|EDM' -or $_.Type -match 'Exact|EDM' })
    $ep = @(Get-EndpointRuleActions); $epSet = @(Get-EndpointSettings | Where-Object { $_.Value })
    $roleMembers = @(Get-DlpRoleMembers)
    $conf = if ($Script:Cache.ContainsKey('ConfidenceSummary')) { $Script:Cache['ConfidenceSummary'] } else { @() }

    $pTotal = $policies.Count
    $pEnforce = Count-Where $policies { $_.Mode -eq 'Enable' }
    $pTest = Count-Where $policies { $_.Mode -like 'Test*' }
    $pOff = Count-Where $policies { $_.Mode -eq 'Disable' -or (Get-Val $_ 'Enabled') -eq $false }
    $c = @{}; foreach ($k in $Script:WorkloadPatterns.Keys) { $c[$k] = @($cov | Where-Object { $_.$k -eq $true }).Count }
    $tips = Count-Where $ruleSum { $_.NotifyUser }
    $override = Count-Where $ruleSum { $_.ResponseInterpretation -eq 'Block with override' }
    $blocks = Count-Where $ruleSum { $_.ResponseInterpretation -eq 'Block' }
    $auditOnly = Count-Where $ruleSum { $_.ResponseInterpretation -eq 'Audit only' }
    $alerting = Count-Where $ruleSum { $_.GenerateAlert -or $_.GenerateIncidentReport }
    $actDays = $Script:Settings.ActivityDays

    function New-Row { param($Category, $Topic, $Finding, $Evidence, $Activity, $Learn)
        [pscustomobject][ordered]@{ Category = $Category; Topic = $Topic; Finding = $Finding; EvidenceFile = $Evidence; SuggestedActivity = $Activity
            Status = 'Not Started'; Owner = ''; TargetDate = ''; Notes = ''; LearnMore = "https://learn.microsoft.com/en-us/purview/$Learn" } }

    $rows = @(
        New-Row 'Fundamentals' 'Microsoft Purview Data Loss Prevention' ("$pTotal DLP policies: $pEnforce enforced, $pTest in test/simulation, $pOff disabled. $($rules.Count) rules." + (Get-StepGap 'DLP compliance policies', 'DLP compliance rules')) '11_DLP_Policies_Summary' 'Confirm which types of information are considered sensitive.' 'dlp-learn-about-dlp'
        New-Row 'Classification' 'Sensitive Information Types' ("$($sitView.Count) SITs available; $(@($rc | Where-Object ReferenceType -eq 'Classifier' | Select-Object -ExpandProperty Name -Unique).Count) distinct classifiers referenced in DLP rules." + (Get-StepGap 'Sensitive information types')) '20_Sensitive_Information_Types_All' 'Review built-in SITs against data requirements.' 'sensitive-information-type-learn-about'
        New-Row 'Classification' 'Custom Sensitive Information Types' ("$($custom.Count) custom SITs; $(Count-Where $custom { $_.UsedInDlpRules }) used in DLP rules. $(@(Get-Dictionaries).Count) keyword dictionaries." + (Get-StepGap 'Sensitive information types', 'SIT rule packages', 'Keyword dictionaries')) '21_Sensitive_Information_Types_Custom' 'Identify organisation-specific identifiers that require custom detection.' 'sit-create-a-custom-sensitive-information-type'
        New-Row 'Classification' 'Exact Data Match' ("$(@(Get-EdmSchemas).Count) EDM schemas; $($edmSits.Count) EDM SITs; $(Count-Where $rc { $_.ClassifierType -match 'ExactMatch|EDM' }) EDM references in DLP rules." + (Get-StepGap 'Exact Data Match schemas')) '24_EDM_Schemas' 'Identify suitable structured datasets and confirm data ownership.' 'sit-learn-about-exact-data-match-based-sits'
        New-Row 'Classification' 'Trainable Classifiers' "$(@($rc | Where-Object { $_.ClassifierType -match 'MLModel|Trainable' } | Select-Object -ExpandProperty Name -Unique).Count) trainable classifiers referenced in DLP rules. Full list: Purview portal > Data classification > Classifiers." '26_Trainable_Classifiers_In_DLP_Rules' 'Review available classifiers and test them against representative documents.' 'classifier-learn-about'
        New-Row 'Classification' 'Confidence Levels' ((@($conf | ForEach-Object { "$($_.Confidence): $($_.References)" }) -join '; ')) '27_Confidence_Level_Summary' 'Agree acceptable detection thresholds and review false positives during testing.' 'sensitive-information-type-learn-about'
        New-Row 'Classification' 'Sensitivity Labels and DLP' ("$(@(Get-Labels).Count) labels, $(@(Get-LabelPolicies).Count) label policies, $(@(Get-AutoLabelPolicies).Count) auto-labelling policies; $(@($rc | Where-Object ReferenceType -eq 'SensitivityLabel' | Select-Object -ExpandProperty Name -Unique).Count) labels used as DLP conditions." + (Get-StepGap 'Sensitivity labels', 'Sensitivity label policies', 'Auto-labelling policies', 'Auto-labelling rules')) '30_Sensitivity_Labels / 34_Labels_Used_As_DLP_Conditions' 'Review current sensitivity label design and adoption.' 'dlp-sensitivity-label-as-condition'
        New-Row 'Policy design' 'DLP Policies' "$pTotal policies / $($rules.Count) rules. $(Count-Where $ruleSum { $_.UsesAdvancedRule }) rules use the advanced rule builder." '11_DLP_Policies_Summary' 'Create a requirements-to-policy mapping for each agreed use case.' 'dlp-learn-about-dlp'
        New-Row 'Policy design' 'Conditions and Actions' "Rule responses - Block: $blocks; Block with override: $override; Notify: $tips; Audit only: $auditOnly." '14_DLP_Rules_Conditions_Actions' 'Define the condition and required response for each use case.' 'dlp-policy-reference'
        New-Row 'Email' 'Exchange Online DLP' ("$($c['Exchange']) policies scoped to Exchange." + $(if ($Script:Cache.ContainsKey('TransportRules')) { " $(@($Script:Cache['TransportRules']).Count) mail flow rules collected." } else { '' }) + (Get-StepGap 'Mail flow (transport) rules')) '12_DLP_Workload_Coverage' 'Define external email, recipient and attachment scenarios that require control.' 'dlp-learn-about-dlp'
        New-Row 'Documents' 'SharePoint Online DLP' "$($c['SharePoint']) policies scoped to SharePoint." '12_DLP_Workload_Coverage' 'Identify important SharePoint sites and existing external-sharing arrangements.' 'dlp-learn-about-dlp'
        New-Row 'Documents' 'OneDrive DLP' "$($c['OneDrive']) policies scoped to OneDrive." '12_DLP_Workload_Coverage' 'Review OneDrive sharing behaviour and high-risk user scenarios.' 'dlp-learn-about-dlp'
        New-Row 'Collaboration' 'Microsoft Teams DLP' "$($c['Teams']) policies scoped to Teams chat and channel messages." '12_DLP_Workload_Coverage' 'Identify sensitive data that users might share through chats and channels.' 'dlp-learn-about-dlp'
        New-Row 'Devices' 'Endpoint DLP' ("$($c['Endpoint']) policies scoped to devices; $($epSet.Count) populated Endpoint DLP settings. Check device onboarding in the portal / Defender." + (Get-StepGap 'Endpoint DLP / tenant policy configuration')) '41_Endpoint_DLP_Settings' 'Confirm device platforms, management arrangements and onboarding prerequisites.' 'endpoint-dlp-learn-about'
        New-Row 'Devices' 'USB and Removable Storage' "$(Count-Where $ep { $_.Setting -match 'RemovableMedia|Usb' }) rule restrictions for removable media." '16_Endpoint_DLP_Rule_Restrictions' 'Agree whether activities should be audited, warned, justified or blocked.' 'endpoint-dlp-learn-about'
        New-Row 'Devices' 'Printing' "$(Count-Where $ep { $_.Setting -match 'Print' }) rule restrictions for printing." '16_Endpoint_DLP_Rule_Restrictions' 'Identify sensitive data types and business processes where printing is required.' 'endpoint-dlp-learn-about'
        New-Row 'Devices' 'Browser and Cloud Uploads' "$(Count-Where $ep { $_.Setting -match 'Upload|Browser|Cloud|Paste|Clipboard' }) rule restrictions for browser / cloud upload / clipboard." '16_Endpoint_DLP_Rule_Restrictions / 41_Endpoint_DLP_Settings' 'Define approved service domains, browsers and upload destinations.' 'endpoint-dlp-learn-about'
        New-Row 'User experience' 'Policy Tips' "$tips rules notify users; $(Count-Where $ruleSum { $_.NotifyPolicyTipCustomText }) use custom tip text." '14_DLP_Rules_Conditions_Actions' 'Agree simple, user-friendly policy-tip wording.' 'dlp-ol365-win32-policy-tips'
        New-Row 'User experience' 'Block with Override' "$override rules block with override." '14_DLP_Rules_Conditions_Actions' 'Define which policies permit overrides and who reviews the justifications.' 'dlp-learn-about-dlp'
        New-Row 'Deployment' 'DLP Simulation' "$pTest policies in simulation/test mode." '11_DLP_Policies_Summary' 'Use simulation, review results, pilot with a limited group, then enforce progressively.' 'dlp-create-deploy-policy'
        New-Row 'Monitoring' 'Activity Explorer' ($(if ($Script:Cache.ContainsKey('Activity')) { "$(@($Script:Cache['Activity']).Count) activity records in the last $actDays days." } else { 'Not collected (enable Activity Explorer in Settings or use -IncludeActivityExplorer).' }) + (Get-StepGap "Activity Explorer ($actDays days)")) '61_Activity_Explorer_Summary' 'Define a regular review process and required operational reports.' 'data-classification-activity-explorer'
        New-Row 'Monitoring' 'DLP Alerts' ("$alerting rules generate alerts or incident reports; $(@(Get-DlpAlertPolicies).Count) DLP-related alert policies." + (Get-StepGap 'Alert policies')) '51_Alert_Policies_DLP / 52_DLP_Rule_Alerting' 'Define alert severity, recipients, investigation steps and escalation routes.' 'dlp-alerts-dashboard-get-started'
        New-Row 'AI and Copilot' 'DLP and Microsoft 365 Copilot' "$($c['Copilot']) policies scoped to Microsoft 365 Copilot / Copilot Chat." '12_DLP_Workload_Coverage' 'Confirm whether Copilot and generative-AI data-loss scenarios are in scope.' 'dlp-microsoft365-copilot-location-learn-about'
        New-Row 'Governance' 'Operational Ownership' ("$(@($roleMembers | Select-Object -ExpandProperty Member -Unique).Count) people/groups hold DLP-relevant Purview roles across $(@($roleMembers | Select-Object -ExpandProperty RoleGroup -Unique).Count) role groups." + (Get-StepGap 'Purview role groups', 'Purview role group members')) '55_DLP_Relevant_Role_Members' 'Create a responsibility matrix covering security, compliance, IT, legal, HR and business owners.' 'purview-permissions'
        New-Row 'Delivery approach' 'Phased DLP Deployment' "Current state - enforced: $pEnforce, test: $pTest, disabled: $pOff." '11_DLP_Policies_Summary' 'Agree policy waves, pilot groups, acceptance criteria and enforcement dates.' 'dlp-create-deploy-policy'
    )
    if ($Script:Errors.Count -gt 0) {
        $rows += New-Row 'Collection' 'Collection gaps' ('Steps skipped or failed: ' + (($Script:Errors | ForEach-Object { "$($_.Step) ($($_.Outcome))" }) -join '; ')) '99_Collection_Errors' 'Re-run with an account that has the missing roles (use menu option 3 to check).' 'purview-permissions'
    }
    # Put the summary first in the workbook
    $rest = $Script:Datasets; $Script:Datasets = [ordered]@{}
    $Script:Cache['SummaryRows'] = $rows
    Export-Dataset '00_Assessment_Summary' $rows -AlreadyFlat
    foreach ($k in $rest.Keys) { $Script:Datasets[$k] = $rest[$k] }

    $ti = @($Script:Cache['TenantInfo'])
    $tenantLabel = if ($ti.Count) { "$($ti[0].Organization) $($ti[0].TenantId)" } else { '' }
    $css = '<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222}h1{font-size:22px}table{border-collapse:collapse;width:100%;font-size:13px}th{background:#0f4c81;color:#fff;text-align:left;padding:6px}td{border-bottom:1px solid #ddd;padding:6px;vertical-align:top}tr:nth-child(even){background:#f6f8fa}</style>'
    $rows | Select-Object Category, Topic, Finding, EvidenceFile, SuggestedActivity |
        ConvertTo-Html -Head $css -Title 'M365 DLP Discovery' -PreContent "<h1>Microsoft Purview DLP discovery</h1><p>Tenant: $tenantLabel<br>Collected: $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>" |
        Set-Content -Path (Join-Path $Script:OutputDir '00_Assessment_Summary.html') -Encoding UTF8
}

function Complete-Run {
    Export-Dataset '99_Collection_Errors' @($Script:Errors) -AlreadyFlat
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel
        $xlsx = Join-Path $Script:OutputDir 'DLP_Configuration.xlsx'
        Write-Log 'Building Excel workbook...'
        $used = @{}
        foreach ($name in $Script:Datasets.Keys) {
            $data = @($Script:Datasets[$name]); if ($data.Count -eq 0) { continue }
            $sheet = ($name -replace '_', ' '); if ($sheet.Length -gt 31) { $sheet = $sheet.Substring(0, 31).Trim() }
            while ($used.ContainsKey($sheet)) { $sheet = $sheet.Substring(0, [Math]::Min(28, $sheet.Length)) + (Get-Random -Maximum 99) }
            $used[$sheet] = $true
            try { $data | Export-Excel -Path $xlsx -WorksheetName $sheet -TableName ('T' + ($name -replace '[^A-Za-z0-9]', '')) -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow -MaxAutoSizeRows 200 }
            catch { Write-Log "  Excel sheet $sheet failed: $($_.Exception.Message)" 'WARN' }
        }
        Write-Log "Workbook: $xlsx" 'OK'
    }
    else { Write-Log 'ImportExcel not installed - CSV/JSON only (menu option 1 installs it).' 'WARN' }
    $lvl = if ($Script:Errors.Count) { 'WARN' } else { 'OK' }
    Write-Log ('Finished. {0} dataset(s) written. {1} step(s) skipped or failed.' -f $Script:Datasets.Count, $Script:Errors.Count) $lvl
    Write-Host "Output: $Script:OutputDir" -ForegroundColor Cyan
    $Script:LogFile = $null
}

function Invoke-AreaReport {
    param([string[]]$AreaKeys)
    if (-not (Connect-Tenant)) { return }
    Start-Run 'Areas'
    Export-AreaTenant
    $keys = @($AreaKeys | Where-Object { $_ -ne '1' } | Sort-Object { [int]$_ })
    # Exchange last: it changes the loaded cmdlets
    $keys = @($keys | Where-Object { $_ -ne '12' }) + @($keys | Where-Object { $_ -eq '12' })
    foreach ($k in $keys) {
        Write-Host ''
        Write-Host "--- $($Script:Areas[$k]) ---" -ForegroundColor Cyan
        & $Script:AreaFunctions[$k]
    }
    Complete-Run
}

function Invoke-FullReport {
    if (-not (Connect-Tenant)) { return }
    Start-Run 'Full'
    $keys = @('1', '2', '3', '4', '5', '6', '7', '8', '9', '10')
    if ($Script:Settings.IncludeActivity) { $keys += '11' }
    foreach ($k in $keys) {
        Write-Host ''
        Write-Host "--- $($Script:Areas[$k]) ---" -ForegroundColor Cyan
        & $Script:AreaFunctions[$k]
    }
    if ($Script:Settings.IncludeGraph) { Write-Host ''; Write-Host "--- $($Script:Areas['13']) ---" -ForegroundColor Cyan; Export-AreaGraph }
    if ($Script:Settings.IncludeExchange) { Write-Host ''; Write-Host "--- $($Script:Areas['12']) ---" -ForegroundColor Cyan; Export-AreaExchange }
    Export-AssessmentSummary
    $wordPath = $null
    if ($Script:Settings.WordReport) {
        try { $wordPath = New-DlpWordReport }
        catch {
            Write-Log "Word report failed: $($_.Exception.Message)" 'ERROR'
            $Script:Errors.Add([pscustomobject]@{ Step = 'Word assessment report'; Outcome = 'Failed'; Error = $_.Exception.Message })
        }
    }
    Complete-Run
    $html = Join-Path $Script:OutputDir '00_Assessment_Summary.html'
    Write-Host "Summary: $html" -ForegroundColor Cyan
    if ($wordPath) { Write-Host "Word report: $wordPath" -ForegroundColor Cyan }
}

#endregion

#region ======================= Access test ============================================

function Test-TenantAccess {
    Write-Host ''
    Write-Host '=== Test credentials and tenant access ===' -ForegroundColor Cyan
    [void](Show-ModuleStatus)
    if (-not (Connect-Tenant)) { return }
    Start-Run 'AccessTest'

    $conn = Get-IppsConnection
    $upn = if ($conn) { $conn.UserPrincipalName } else { $Script:Settings.UserPrincipalName }
    Write-Log "Signed in as: $upn   Tenant: $(if ($conn) { $conn.TenantID } else { 'unknown' })"

    $end = (Get-Date).ToUniversalTime()
    $checks = @(
        @{ Area = '2,3 Policies and rules';         Cmdlet = 'Get-DlpCompliancePolicy';                  Test = { Get-DlpCompliancePolicy | Select-Object -First 1 } }
        @{ Area = '2,3 Policies and rules';         Cmdlet = 'Get-DlpComplianceRule';                    Test = { Get-DlpComplianceRule | Select-Object -First 1 } }
        @{ Area = '5 Sensitive information types';  Cmdlet = 'Get-DlpSensitiveInformationType';          Test = { Get-DlpSensitiveInformationType | Select-Object -First 1 } }
        @{ Area = '5 Sensitive information types';  Cmdlet = 'Get-DlpSensitiveInformationTypeRulePackage'; Test = { Get-DlpSensitiveInformationTypeRulePackage | Select-Object -First 1 } }
        @{ Area = '5 Sensitive information types';  Cmdlet = 'Get-DlpKeywordDictionary';                 Test = { Get-DlpKeywordDictionary | Select-Object -First 1 } }
        @{ Area = '6 Exact Data Match';             Cmdlet = 'Get-DlpEdmSchema';                         Test = { Get-DlpEdmSchema | Select-Object -First 1 } }
        @{ Area = '7 Sensitivity labels';           Cmdlet = 'Get-Label';                                Test = { Get-Label | Select-Object -First 1 } }
        @{ Area = '7 Sensitivity labels';           Cmdlet = 'Get-LabelPolicy';                          Test = { Get-LabelPolicy | Select-Object -First 1 } }
        @{ Area = '7 Auto-labelling';               Cmdlet = 'Get-AutoSensitivityLabelPolicy';           Test = { Get-AutoSensitivityLabelPolicy | Select-Object -First 1 } }
        @{ Area = '7 Auto-labelling';               Cmdlet = 'Get-AutoSensitivityLabelRule';             Test = { Get-AutoSensitivityLabelRule | Select-Object -First 1 } }
        @{ Area = '8 Endpoint DLP';                 Cmdlet = 'Get-PolicyConfig';                         Test = { Get-PolicyConfig | Select-Object -First 1 } }
        @{ Area = '9 Alerts';                       Cmdlet = 'Get-ProtectionAlert';                      Test = { Get-ProtectionAlert | Select-Object -First 1 } }
        @{ Area = '10 Ownership';                   Cmdlet = 'Get-RoleGroup';                            Test = { Get-RoleGroup | Select-Object -First 1 } }
        @{ Area = '10 Ownership';                   Cmdlet = 'Get-RoleGroupMember';                      Test = { $g = Get-RoleGroup | Select-Object -First 1; if ($g) { Get-RoleGroupMember -Identity ([string]$g.Identity) | Select-Object -First 1 } } }
        @{ Area = '11 Activity Explorer';           Cmdlet = 'Export-ActivityExplorerData';              Test = { Export-ActivityExplorerData -StartTime $end.AddHours(-1) -EndTime $end -OutputFormat Json -PageSize 1 } }
    )

    $results = foreach ($c in $checks) {
        $available = [bool](Get-Command $c.Cmdlet -ErrorAction SilentlyContinue)
        $outcome = 'MISSING'; $detail = 'Cmdlet not loaded for this account: missing Purview role, or feature not licensed.'
        if ($available) {
            try { [void](& $c.Test); $outcome = 'PASS'; $detail = 'Cmdlet available and returned successfully.' }
            catch { $outcome = 'FAIL'; $detail = $_.Exception.Message.Split([Environment]::NewLine)[0] }
        }
        [pscustomobject]@{ Area = $c.Area; Cmdlet = $c.Cmdlet; Result = $outcome; Detail = $detail }
    }

    if ($Script:Settings.IncludeGraph -or ($Mode -eq 'Menu' -and (Read-YesNo 'Also test Microsoft Graph access (licensing, DLP alerts, Entra roles, sharing settings)?' $false))) {
        $gt = Invoke-GraphCollector -TestOnly
        if ($gt) { $results = @($results) + @($gt.Tests | Where-Object { $_ }) }
        else { $results = @($results) + [pscustomobject]@{ Area = '13 Microsoft Graph'; Cmdlet = 'Connect-MgGraph'; Result = 'FAIL'; Detail = 'Could not connect to Microsoft Graph (module missing, sign-in cancelled or consent not granted).' } }
    }
    if ($Script:Settings.IncludeExchange -or ($Mode -eq 'Menu' -and (Read-YesNo 'Also test Exchange Online access (mail flow rules, audit config)?' $false))) {
        try {
            Connect-Exo
            foreach ($x in @(@{ C = 'Get-TransportRule'; T = { Get-TransportRule | Select-Object -First 1 } }, @{ C = 'Get-AdminAuditLogConfig'; T = { Get-AdminAuditLogConfig } })) {
                $o = 'MISSING'; $d = 'Cmdlet not loaded for this account.'
                if (Get-Command $x.C -ErrorAction SilentlyContinue) { try { [void](& $x.T); $o = 'PASS'; $d = 'Cmdlet available and returned successfully.' } catch { $o = 'FAIL'; $d = $_.Exception.Message.Split([Environment]::NewLine)[0] } }
                $results = @($results) + [pscustomobject]@{ Area = '12 Exchange Online'; Cmdlet = $x.C; Result = $o; Detail = $d }
            }
        }
        catch { $results = @($results) + [pscustomobject]@{ Area = '12 Exchange Online'; Cmdlet = 'Connect-ExchangeOnline'; Result = 'FAIL'; Detail = $_.Exception.Message } }
        Reset-AfterExo
    }

    Write-Host ''
    foreach ($r in $results) {
        $col = switch ($r.Result) { 'PASS' { 'Green' } 'FAIL' { 'Red' } default { 'Yellow' } }
        Write-Host ('  {0,-7} {1,-30} {2,-44} {3}' -f $r.Result, $r.Area, $r.Cmdlet, $(if ($r.Result -ne 'PASS') { $r.Detail } else { '' })) -ForegroundColor $col
    }

    # Role groups the signed-in account is a direct member of
    $membership = @()
    if ($upn -and (Get-Command Get-RoleGroupMember -ErrorAction SilentlyContinue)) {
        Write-Log 'Checking Purview role group membership for the signed-in account (direct membership only)...'
        foreach ($g in @(Get-RoleGroup -ErrorAction SilentlyContinue)) {
            try {
                $m = @(Get-RoleGroupMember -Identity ([string]$g.Identity) -ErrorAction Stop | Where-Object {
                        (Get-Val $_ 'PrimarySmtpAddress') -eq $upn -or (Get-Val $_ 'WindowsLiveID') -eq $upn -or (Get-Val $_ 'Name') -eq $upn })
                if ($m.Count) { $membership += [pscustomobject]@{ RoleGroup = $g.Name; Roles = ConvertTo-CellValue (Get-Val $g 'Roles') } }
            }
            catch { }
        }
        Write-Host ''
        if ($membership.Count) {
            Write-Host 'Purview role groups for this account:' -ForegroundColor Cyan
            $membership | ForEach-Object { Write-Host "  - $($_.RoleGroup)" -ForegroundColor Gray }
        }
        else { Write-Host 'No direct Purview role group membership found (access may come from an Entra ID role such as Global Administrator or Compliance Administrator, or via a group).' -ForegroundColor Yellow }
    }

    $pass = Count-Where $results { $_.Result -eq 'PASS' }; $total = @($results).Count
    Write-Host ''
    if ($pass -eq $total) { Write-Host "All $total checks passed. This account can run the complete tenant DLP report." -ForegroundColor Green }
    else {
        Write-Host "$pass of $total checks passed. Reports will skip the areas marked MISSING or FAIL." -ForegroundColor Yellow
        Write-Host 'Typical fix: add the account to the Compliance Administrator role group (or the read-only roles listed in the user guide), wait up to an hour, and test again.' -ForegroundColor Yellow
    }
    Export-Dataset 'Access_Test_Results' @($results) -AlreadyFlat
    Export-Dataset 'Access_Test_RoleGroups' $membership -AlreadyFlat
    Write-Host "Results saved to: $Script:OutputDir" -ForegroundColor Cyan
    $Script:LogFile = $null
}

#endregion

#region ======================= Menus ==================================================

function Show-Header {
    Clear-Host
    $c = Get-IppsConnection
    $status = if ($c) { "Connected: $($c.UserPrincipalName)  |  Tenant: $($c.TenantID)" } else { 'Not connected' }
    Write-Host '================================================================================' -ForegroundColor DarkCyan
    Write-Host '   Microsoft Purview DLP Discovery  (read-only)   *** BETA ***' -ForegroundColor Cyan
    Write-Host '   Includes optional Microsoft Graph data. Main version: Get-M365DlpConfiguration.ps1' -ForegroundColor DarkYellow
    Write-Host '================================================================================' -ForegroundColor DarkCyan
    Write-Host "   $status" -ForegroundColor $(if ($c) { 'Green' } else { 'Yellow' })
    Write-Host ("   Output: {0}   Activity Explorer: {1}   Exchange: {2}" -f $Script:Settings.OutputRoot,
        $(if ($Script:Settings.IncludeActivity) { "On ($($Script:Settings.ActivityDays) days)" } else { 'Off' }),
        $(if ($Script:Settings.IncludeExchange) { 'On' } else { 'Off' })) -ForegroundColor Gray
    Write-Host '--------------------------------------------------------------------------------' -ForegroundColor DarkCyan
}

function Show-AreaMenu {
    Show-Header
    Write-Host '  Report on individual DLP items' -ForegroundColor Cyan
    Write-Host ''
    foreach ($k in $Script:Areas.Keys) { Write-Host ('   {0,2}. {1}' -f $k, $Script:Areas[$k]) }
    Write-Host ''
    Write-Host '   Enter one or more numbers separated by commas (e.g. 2,3,8), or B to go back.' -ForegroundColor Gray
    $sel = Read-Host '  Selection'
    if ($sel -match '^\s*[bB]\s*$' -or [string]::IsNullOrWhiteSpace($sel)) { return }
    $keys = @($sel -split '[,\s]+' | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Where-Object { $Script:Areas.Contains($_) } | Select-Object -Unique)
    $bad = @($sel -split '[,\s]+' | Where-Object { $_ -and -not $Script:Areas.Contains($_.Trim()) })
    if ($bad.Count) { Write-Host "Ignored invalid selection(s): $($bad -join ', ')" -ForegroundColor Yellow }
    if ($keys.Count -eq 0) { Write-Host 'Nothing selected.' -ForegroundColor Yellow; Wait-Key; return }
    Invoke-AreaReport -AreaKeys $keys
    Wait-Key
}

function Show-SettingsMenu {
    while ($true) {
        Show-Header
        Write-Host '  Settings' -ForegroundColor Cyan
        Write-Host ''
        Write-Host "   1. Output folder                      : $($Script:Settings.OutputRoot)"
        Write-Host "   2. Admin account (UPN)                : $(if ($Script:Settings.UserPrincipalName) { $Script:Settings.UserPrincipalName } else { '(choose at sign-in)' })"
        Write-Host "   3. Include Activity Explorer in full  : $($Script:Settings.IncludeActivity)"
        Write-Host "   4. Activity Explorer days (1-30)      : $($Script:Settings.ActivityDays)"
        Write-Host "   5. Activity Explorer max pages        : $($Script:Settings.ActivityMaxPages)"
        Write-Host "   6. Include Exchange Online in full    : $($Script:Settings.IncludeExchange)"
        Write-Host "  10. Include Microsoft Graph in full    : $($Script:Settings.IncludeGraph)  (licensing, DLP alerts, Entra roles, sharing)"
        Write-Host "   7. Sign in with -DisableWAM           : $($Script:Settings.DisableWAM)  (use if sign-in fails or the sign-in window does not appear)"
        Write-Host "   8. Word report with complete report   : $($Script:Settings.WordReport)  (DLP_Assessment_Report.docx)"
        Write-Host "   9. Frameworks in Word report           : $(($Script:Settings.Frameworks) -join ', ')  (NCSC, NISTCSF, NIST80053, ISO27001, All, None)"
        Write-Host '   B. Back'
        $s = Read-Host '  Choose'
        switch ($s.Trim().ToUpper()) {
            '1' { $p = Read-Host '  New output folder'; if ($p) { if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }; $Script:Settings.OutputRoot = (Resolve-Path $p).Path } }
            '2' { $u = Read-Host '  Admin UPN (blank to clear)'; $Script:Settings.UserPrincipalName = $u.Trim() }
            '3' { $Script:Settings.IncludeActivity = -not $Script:Settings.IncludeActivity }
            '4' { $d = Read-Host '  Days'; if ($d -as [int] -and [int]$d -ge 1 -and [int]$d -le 30) { $Script:Settings.ActivityDays = [int]$d } }
            '5' { $d = Read-Host '  Max pages'; if ($d -as [int] -and [int]$d -ge 1 -and [int]$d -le 200) { $Script:Settings.ActivityMaxPages = [int]$d } }
            '6' { $Script:Settings.IncludeExchange = -not $Script:Settings.IncludeExchange }
            '7' { $Script:Settings.DisableWAM = -not $Script:Settings.DisableWAM }
            '8' { $Script:Settings.WordReport = -not $Script:Settings.WordReport }
            '10' { $Script:Settings.IncludeGraph = -not $Script:Settings.IncludeGraph }
            '9' { $f = Read-Host '  Frameworks (comma separated)'; if ($f) { $Script:Settings.Frameworks = @($f -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } }
            'B' { return }
        }
    }
}

function Show-MainMenu {
    while ($true) {
        Show-Header
        Write-Host '   1. Install / update required PowerShell modules'
        Write-Host '   2. Connect to tenant'
        Write-Host '   3. Test credentials and tenant access'
        Write-Host '   4. Report on individual DLP items'
        Write-Host '   5. Complete tenant DLP report'
        Write-Host '   6. Settings'
        Write-Host '   7. Disconnect / switch tenant'
        Write-Host '   Q. Quit'
        Write-Host ''
        $choice = Read-Host '  Choose an option'
        switch ($choice.Trim().ToUpper()) {
            '1' { Install-RequiredModules; Wait-Key }
            '2' { Invoke-ConnectMenu; Wait-Key }
            '3' { Test-TenantAccess; Wait-Key }
            '4' { Show-AreaMenu }
            '5' {
                Show-Header
                Write-Host '  Complete tenant DLP report' -ForegroundColor Cyan
                $Script:Settings.IncludeActivity = Read-YesNo "  Include Activity Explorer data ($($Script:Settings.ActivityDays) days)?" $Script:Settings.IncludeActivity
                $Script:Settings.IncludeExchange = Read-YesNo '  Include Exchange Online mail flow rules and audit config?' $Script:Settings.IncludeExchange
                $Script:Settings.IncludeGraph = Read-YesNo '  Include Microsoft Graph (licensing, DLP alerts, Entra roles, sharing settings)?' $Script:Settings.IncludeGraph
                Invoke-FullReport; Wait-Key
            }
            '6' { Show-SettingsMenu }
            '7' { Disconnect-Tenant; $Script:Settings.UserPrincipalName = ''; Wait-Key }
            'Q' {
                if (-not $NoDisconnect) { Disconnect-Tenant }
                return
            }
        }
    }
}

#endregion

#region ======================= Entry point ============================================

switch ($Mode) {
    'Menu'           { Show-MainMenu }
    'InstallModules' { Install-RequiredModules }
    'TestAccess'     { Test-TenantAccess; if (-not $NoDisconnect) { Disconnect-Tenant } }
    'Full'           { Invoke-FullReport; if (-not $NoDisconnect) { Disconnect-Tenant } }
    'Area' {
        if (-not $Area) { throw 'Use -Area with one or more area numbers (1-12), e.g. -Area 2,3,8' }
        $keys = foreach ($a in $Area) {
            foreach ($part in ($a -split ',')) {
                $p = $part.Trim()
                if ($Script:Areas.Contains($p)) { $p }
                else {
                    $match = $Script:Areas.Keys | Where-Object { $Script:Areas[$_] -like "*$p*" } | Select-Object -First 1
                    if ($match) { $match } else { Write-Warning "Unknown area '$p' ignored." }
                }
            }
        }
        Invoke-AreaReport -AreaKeys @($keys | Select-Object -Unique)
        if (-not $NoDisconnect) { Disconnect-Tenant }
    }
}

#endregion
