<#
.SYNOPSIS
    Microsoft Purview Data Loss Prevention (DLP) discovery tool for Microsoft 365 tenants - menu driven.

.DESCRIPTION
    Read-only. Only Get-* and Export-* cmdlets are used against the tenant; nothing is changed.

    Menu options:
      1. Install / update the required PowerShell modules
      2. Connect to tenant (sign in, or switch account/tenant, and show what the account can access)
      3. Test credentials and tenant access (checks every cmdlet the reports need)
      4. Report on individual DLP items (choose one or more areas)
      5. Complete tenant DLP report (every area plus an assessment summary)
      6. Settings (output folder, Activity Explorer, Exchange Online, Content Explorer, DLP rule activity,
         customer name, assessor, logo, comparison with a previous report)
      7. Disconnect / switch tenant
      8. DLP alerts dashboard (DLP alerts and rule matches for the last N days, as an easy-to-read HTML page)

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
      13  Content Explorer - how many stored items contain each sensitive information type / label
      14  DLP rule activity and alerts from the unified audit log - matches, overrides, false positives, alerts,
          plus DLP_Alerts_Dashboard.html
      15  Data lifecycle and records management - retention policies and rules, retention labels (publish and
          auto-apply), records and regulatory records, disposition review, event-based retention, adaptive scopes
      16  Risk and compliance - insider risk management policies, communication compliance policies and rules,
          information barrier policies and segments
      17  Audit and eDiscovery - audit log retention policies, unified audit log status, eDiscovery cases and holds

    The complete report (option 5) includes areas 15 to 17 by default. Turn them off in Settings or with -Scope.

    Output for each report run goes to a new time-stamped folder:
      CSV + JSON per dataset, DLP_Configuration.xlsx (if ImportExcel is installed), Collection.log,
      and for the complete report 00_Assessment_Summary.csv/.html and DLP_Assessment_Report.docx
      (a plain-English Word report explaining each area, the findings and Microsoft Learn links),
      plus 80_Action_Plan (prioritised actions with effort and licence) and 81-84 (maturity scores,
      hygiene checks, changes since the previous report, coverage map).

    No Microsoft Graph: everything uses Security & Compliance PowerShell and, optionally, Exchange Online PowerShell.

    READ-ONLY: the only tenant commands the script runs are Get-*, Export-ActivityExplorerData,
    Export-ContentExplorerData and Search-UnifiedAuditLog (all of which only read data), plus Connect-/Disconnect-.
    Set-/New- DLP, label, retention and audit commands appear ONLY as example text inside @' '@ blocks in
    Add-ImplementationGuidance, which prints them into section 24 of the Word report for the customer to review.
    They are never executed. Commands that write files (New-Item, Set-Content, Remove-Item on the report file)
    act only on the output folder; Install-/Uninstall-Module run only from menu option 1 after you confirm.

.PARAMETER Frameworks
    Frameworks referenced in the Word report: any of NCSC, NISTCSF, NIST80053, ISO27001 (or NIST, ISO, All, None).
    Default: all four.

.PARAMETER NoWordReport
    Do not create DLP_Assessment_Report.docx at the end of the complete report.

.PARAMETER DisableWAM
    Sign in without the Windows Web Account Manager broker (passes -DisableWAM where the module supports it).
    Use if sign-in fails, for example with "You cannot call a method on a null-valued expression".

.PARAMETER IncludeContentExplorer
    Complete report: count stored items containing each sensitive information type (those used in rules plus common
    built-in types) and each sensitivity label, per location, with Export-ContentExplorerData.
    Needs the Content Explorer List Viewer role. ContentExplorerMaxTags limits the number of types queried (default 25).

.PARAMETER IncludeDlpAudit
    Complete report: read DLP rule matches, overrides and false positives from the unified audit log (Search-UnifiedAuditLog,
    record types ComplianceDLPExchange, ComplianceDLPSharePoint, DLPEndpoint) for the last AuditDays days (default 30).
    Signs in to Exchange Online; needs View-Only Audit Logs or Audit Logs. AuditMaxRecords caps records per type.

.PARAMETER MaskNamesInDashboard
    Replace people's names with User 1, User 2 ... and hide file names and subjects in DLP_Alerts_Dashboard.html.

.PARAMETER CustomerName
    Name shown on the report cover ("Prepared for"). Default: the tenant organisation name.

.PARAMETER AssessorName
    Shown on the report cover as "Prepared by".

.PARAMETER LogoPath
    PNG or JPEG logo shown on the report cover.

.PARAMETER CompareWith
    A previous complete-report folder, or Latest for the newest earlier report for the same tenant. Adds
    "Changes since the last assessment" and maturity changes to the report. In the menu you are asked when one exists.

.PARAMETER Mode
    Menu (default) | InstallModules | TestAccess | Full | Area | AlertDashboard

.PARAMETER Area
    With -Mode Area: one or more area numbers (1-17) or names, e.g. -Area 2,3,8,15

.PARAMETER Scope
    What the complete report covers. DLP is always included. Add Lifecycle (retention and records),
    RiskCompliance (insider risk, communication compliance, information barriers) and Audit (audit log retention,
    eDiscovery cases and holds). Default: All. Example: -Scope DLP, Lifecycle

.PARAMETER MaxEDiscoveryCases
    Maximum number of eDiscovery cases whose holds are read (1-500, default 100). Each case needs one extra call.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1
    Starts the interactive menu.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com -IncludeActivityExplorer
    Runs the complete report without the menu.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode Full -IncludeExchange -IncludeContentExplorer -IncludeDlpAudit -CustomerName 'Fabrikam Ltd' -AssessorName 'Jane Smith' -LogoPath .\logo.png -CompareWith Latest
    Complete report with evidence of where sensitive data is stored and how often rules fire, branded for the customer,
    compared with the previous report for the same tenant.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode Area -Area 2,3,8 -UserPrincipalName admin@contoso.onmicrosoft.com
    Reports on policies, rules and Endpoint DLP only.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode Full -Scope DLP -UserPrincipalName admin@contoso.onmicrosoft.com
    Complete report covering DLP only (no retention, insider risk, communication compliance, audit or eDiscovery).

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode TestAccess -UserPrincipalName admin@contoso.onmicrosoft.com

.NOTES
    Requires: ExchangeOnlineManagement 3.2.0 or later. Optional: ImportExcel.
    Works in Windows PowerShell 5.1 and PowerShell 7+.
#>

[CmdletBinding()]
param(
    [ValidateSet('Menu', 'InstallModules', 'TestAccess', 'Full', 'Area', 'AlertDashboard')]
    [string]$Mode = 'Menu',
    [string[]]$Area,
    [string]$OutputRoot = [Environment]::GetFolderPath('MyDocuments'),
    [string]$UserPrincipalName,
    [switch]$IncludeActivityExplorer,
    [ValidateRange(1, 30)][int]$ActivityDays = 7,
    [ValidateRange(1, 200)][int]$ActivityMaxPages = 20,
    [switch]$IncludeExchange,
    [switch]$DisableWAM,
    [switch]$NoWordReport,
    [string[]]$Frameworks = @('NCSC', 'NISTCSF', 'NIST80053', 'ISO27001'),
    [switch]$IncludeContentExplorer,
    [ValidateRange(1, 60)][int]$ContentExplorerMaxTags = 25,
    [switch]$IncludeDlpAudit,
    [ValidateRange(1, 90)][int]$AuditDays = 30,
    [ValidateRange(100, 50000)][int]$AuditMaxRecords = 20000,
    [string]$CustomerName,
    [string]$AssessorName,
    [string]$LogoPath,
    [string]$CompareWith,
    [switch]$MaskNamesInDashboard,
    [ValidateSet('All', 'DLP', 'Lifecycle', 'RiskCompliance', 'Audit')]
    [string[]]$Scope = @('All'),
    [ValidateRange(1, 500)][int]$MaxEDiscoveryCases = 100,
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
    DisableWAM        = [bool]$DisableWAM
    WordReport        = -not [bool]$NoWordReport
    Frameworks        = $Frameworks
    IncludeContentExplorer = [bool]$IncludeContentExplorer
    ContentExplorerMaxTags = $ContentExplorerMaxTags
    IncludeDlpAudit   = [bool]$IncludeDlpAudit
    AuditDays         = $AuditDays
    AuditMaxRecords   = $AuditMaxRecords
    CustomerName      = $CustomerName
    AssessorName      = $AssessorName
    LogoPath          = $LogoPath
    CompareWith       = $CompareWith
    AlertDashboardShowNames = -not [bool]$MaskNamesInDashboard
    IncludeLifecycle      = [bool]($Scope -contains 'All' -or $Scope -contains 'Lifecycle')
    IncludeRiskCompliance = [bool]($Scope -contains 'All' -or $Scope -contains 'RiskCompliance')
    IncludeAuditDiscovery = [bool]($Scope -contains 'All' -or $Scope -contains 'Audit')
    MaxEDiscoveryCases    = $MaxEDiscoveryCases
}
$Script:RequiredModules = @(
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; MinVersion = [version]'3.2.0'; Required = $true;  Purpose = 'Security & Compliance and Exchange Online cmdlets' }
    [pscustomobject]@{ Name = 'ImportExcel';              MinVersion = [version]'7.0.0'; Required = $false; Purpose = 'Excel workbook output (optional)' }
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
    '13' = 'Content Explorer - where sensitive data is stored'
    '14' = 'DLP rule activity and alerts from the audit log (with alerts dashboard)'
    '15' = 'Data lifecycle and records - retention policies, retention labels, records, adaptive scopes'
    '16' = 'Risk and compliance - insider risk, communication compliance, information barriers'
    '17' = 'Audit and eDiscovery - audit log retention, audit status, eDiscovery cases and holds'
}

$Script:LogFile   = $null
$Script:OutputDir = $null
$Script:Datasets  = [ordered]@{}
$Script:Errors    = [System.Collections.Generic.List[object]]::new()
$Script:Cache     = @{}
$Script:AskCompare    = $false
$Script:CompareFolder = ''
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
    # Current tenants return each SIT as a flat entry (name, id, mincount, confidencelevel...) rather than inside a 'sensitivetypes' group.
    $flatName = Get-Val $Node 'name'
    if ($flatName -and -not (Get-Val $Node 'sensitivetypes') -and -not (Get-Val $Node 'labels') -and ((Get-Val $Node 'mincount') -or (Get-Val $Node 'confidencelevel') -or (Get-Val $Node 'minconfidence') -or (Get-Val $Node 'classifiertype'))) {
        $ct = [string](Get-Val $Node 'classifiertype')
        [pscustomobject]@{
            Policy = $Policy; Rule = $Rule; Source = $Source; Group = $GroupName
            ReferenceType   = 'Classifier'
            Name            = $flatName
            Id              = Get-Val $Node 'id'
            ClassifierType  = $(if (-not $ct -or $ct -eq 'Content') { 'Content (SIT)' } else { $ct })
            ConfidenceLevel = Get-Val $Node 'confidencelevel'
            MinConfidence   = Get-Val $Node 'minconfidence'
            MaxConfidence   = Get-Val $Node 'maxconfidence'
            MinCount        = Get-Val $Node 'mincount'
            MaxCount        = Get-Val $Node 'maxcount'
        }
        return
    }
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
    if ($eom.Version -lt [version]'3.2.0') { Write-Host "ExchangeOnlineManagement $($eom.Version) is too old (3.2.0+ required). Use menu option 1 to update." -ForegroundColor Red; return $false }
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
            $found = @(
                Find-RuleClassifiers (Get-Val $r 'ContentContainsSensitiveInformation') $pol $r.Name 'Condition'
                Find-RuleClassifiers (Get-Val $r 'ExceptIfContentContainsSensitiveInformation') $pol $r.Name 'Exception'
            )
            # AdvancedRule repeats the same conditions as JSON; only use it when the rule has no simple SIT properties.
            if ($found.Count -eq 0) { $found = @(Find-RuleClassifiers (Get-Val $r 'AdvancedRule') $pol $r.Name 'AdvancedRule') }
            $found
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

# ---------- Sensitivity label settings (encryption, content marking, scopes, sites and groups) ----------

$Script:LabelScopeNames = [ordered]@{ File = 'Files'; Email = 'Emails'; Site = 'Sites'; UnifiedGroup = 'Groups and Teams'; Teamwork = 'Teams meetings'; SchematizedData = 'Data assets' }

function ConvertFrom-LabelActions {
    <# Get-Label returns LabelActions as JSON strings. Returns one object per action: Type, SubType and a settings hashtable (lower-case keys). #>
    param($Label)
    foreach ($a in @(Get-Val $Label 'LabelActions')) {
        if ($null -eq $a) { continue }
        $o = $null
        if ($a -is [string]) { try { $o = $a | ConvertFrom-Json } catch { continue } } else { $o = $a }
        $h = @{}
        foreach ($s in @(Get-Val $o 'Settings')) {
            $k = [string](Get-Val $s 'Key'); if ($k) { $h[$k.ToLower()] = [string](Get-Val $s 'Value') }
        }
        [pscustomobject]@{ Type = ([string](Get-Val $o 'Type')).ToLower(); SubType = ([string](Get-Val $o 'SubType')).ToLower(); Settings = $h }
    }
}

function Get-LabelSettingValue {
    <# Reads "[key, value]" entries from a label's Settings property. #>
    param($Label, [string]$Key)
    foreach ($s in @(Get-Val $Label 'Settings')) {
        $m = [regex]::Match([string]$s, '^\[\s*([^,]+?)\s*,\s*(.*?)\s*\]$')
        if ($m.Success -and $m.Groups[1].Value -ieq $Key) { return $m.Groups[2].Value }
    }
    return $null
}

function Get-LabelSettingsView {
    Get-Cached 'LabelSettingsView' {
        $byGuid = @{}; foreach ($l in @(Get-Labels)) { $byGuid[[string](Get-Val $l 'Guid')] = $l }
        foreach ($l in @(Get-Labels)) {
            $acts = @(ConvertFrom-LabelActions $l)
            # Scope
            $ct = [string](ConvertTo-CellValue (Get-Val $l 'ContentType'))
            $scope = @(foreach ($k in $Script:LabelScopeNames.Keys) { if ($ct -match "\b$k\b") { $Script:LabelScopeNames[$k] } })
            # Encryption
            $enc = 'None'; $encType = ''
            $e = $acts | Where-Object { $_.Type -eq 'encrypt' } | Select-Object -First 1
            if ($e -and $e.Settings['disabled'] -ne 'true') {
                $pt = [string]$e.Settings['protectiontype']; $encType = $pt
                $parts = @()
                switch -regex ($pt) {
                    '^template$' {
                        $rd = $e.Settings['rightsdefinitions']; $n = 0; $any = $false
                        if ($rd) { try { $rdo = @($rd | ConvertFrom-Json); $n = $rdo.Count; $any = [bool]@($rdo | Where-Object { [string]$_.Identity -match 'AuthenticatedUsers' }).Count } catch { } }
                        $parts += $(if ($n) { "Admin-assigned permissions ($(Get-Plural $n 'user, group or domain' 'users, groups or domains'))" } else { 'Admin-assigned permissions' })
                        if ($any) { $parts += 'includes any authenticated user' }
                    }
                    '^userdefined$' {
                        $ud = @()
                        if ($e.Settings['donotforward'] -eq 'true') { $ud += 'Do Not Forward in Outlook' }
                        if ($e.Settings['encryptonly'] -eq 'true') { $ud += 'Encrypt-Only in Outlook' }
                        if ($e.Settings['promptuser'] -eq 'true') { $ud += 'users choose permissions in Office apps' }
                        $parts += 'Users assign permissions' + $(if ($ud.Count) { ' (' + ($ud -join ', ') + ')' } else { '' })
                    }
                    '^removeprotection$' { $parts += 'Removes encryption' }
                    default { $parts += $(if ($pt) { "Encrypts ($pt)" } else { 'Encrypts' }) }
                }
                if ($e.Settings['doublekeyencryptionurl']) { $parts += 'Double Key Encryption' }
                $exp = [string]$e.Settings['contentexpiredondateindaysornever']
                if ($exp -and $exp -ne 'Never') { $parts += "access expires after $exp days" } elseif ($exp -eq 'Never') { $parts += 'access never expires' }
                $off = [string]$e.Settings['offlineaccessdays']
                if ($off -eq '-1') { $parts += 'offline access always' } elseif ($off -eq '0') { $parts += 'no offline access' } elseif ($off) { $parts += "offline access $off days" }
                $enc = $parts -join '; '
            }
            elseif (-not $acts.Count -and (Get-Val $l 'EncryptionEnabled') -eq $true) { $enc = 'Encrypts'; $encType = [string](Get-Val $l 'EncryptionProtectionType') }
            # Content marking
            $mark = @()
            foreach ($m in @($acts | Where-Object { $_.Type -match '^applycontentmarking|^applywatermarking' })) {
                if ($m.Settings['disabled'] -eq 'true') { continue }
                $txt = [string]$m.Settings['text']; if ($txt.Length -gt 60) { $txt = $txt.Substring(0, 60) + '...' }
                $kind = if ($m.Type -match 'watermark') { 'Watermark' } elseif ($m.SubType -match 'footer' -or $m.Type -match 'footer') { 'Footer' } else { 'Header' }
                $mark += $(if ($txt) { "$($kind): '$txt'" } else { $kind })
            }
            # Sites, groups and Teams
            $site = @()
            foreach ($g in @($acts | Where-Object { $_.Type -match '^protect(group|site|teams|domain)' })) {
                foreach ($k in $g.Settings.Keys) {
                    $v = [string]$g.Settings[$k]
                    switch ($k) {
                        'privacy' { $site += "Privacy: $v" }
                        'allowaccesstoguestusers' { $site += $(if ($v -eq 'true') { 'Guests allowed' } else { 'Guests not allowed' }) }
                        'allowemailfromguestusers' { if ($v -eq 'false') { $site += 'Guests cannot email the group' } }
                        'externalsharingcontroltype' {
                            $site += 'External sharing: ' + $(switch ($v) { 'Disabled' { 'only people in the organisation' } 'ExistingExternalUserSharingOnly' { 'existing guests only' } 'ExternalUserSharingOnly' { 'new and existing guests' } 'ExternalUserAndGuestSharing' { 'anyone' } default { $v } })
                        }
                        'allowfullaccess' { if ($v -eq 'true') { $site += 'Unmanaged devices: full access' } }
                        'allowlimitedaccess' { if ($v -eq 'true') { $site += 'Unmanaged devices: web-only access' } }
                        'blockaccess' { if ($v -eq 'true') { $site += 'Unmanaged devices: blocked' } }
                        'defaultsharinglinktype' { if ($v) { $site += "Default sharing link: $v" } }
                    }
                }
            }
            $cond = ConvertTo-CellValue (Get-Val $l 'Conditions')
            $autoApp = if ($cond -and $cond -ne '{}' -and $cond -ne '[]') { 'Configured' } else { '' }
            $parent = [string](ConvertTo-CellValue (Get-Val $l 'ParentId'))
            $parentName = if ($parent -and $byGuid.ContainsKey($parent)) { $p = $byGuid[$parent]; if (Get-Val $p 'DisplayName') { [string](Get-Val $p 'DisplayName') } else { [string]$p.Name } } else { $parent }
            [pscustomobject][ordered]@{
                Label = $(if (Get-Val $l 'DisplayName') { [string](Get-Val $l 'DisplayName') } else { [string]$l.Name }); Name = [string]$l.Name; Parent = $parentName
                Priority = Get-Val $l 'Priority'; AppliesTo = ($scope -join ', '); Encryption = $enc; EncryptionType = $encType
                ContentMarking = ($mark -join '; '); SitesGroupsTeams = ($site -join '; '); AutoLabelInApps = $autoApp
                Colour = Get-LabelSettingValue $l 'color'; Disabled = Get-Val $l 'Disabled'; Guid = [string](Get-Val $l 'Guid')
            }
        }
    }
}

# ---------- Data lifecycle and records management (Security & Compliance PowerShell) ----------

function Get-RetentionPolicies    { Get-Cached 'RetentionPolicies'  { Invoke-Step 'Retention policies' -Requires Get-RetentionCompliancePolicy -Action { Get-RetentionCompliancePolicy } } }
function Get-RetentionRules       { Get-Cached 'RetentionRules'     { Invoke-Step 'Retention rules' -Requires Get-RetentionComplianceRule -Action { Get-RetentionComplianceRule } } }
function Get-RetentionLabels      { Get-Cached 'RetentionLabels'    { Invoke-Step 'Retention labels' -Requires Get-ComplianceTag -Action { Get-ComplianceTag } } }
function Get-AdaptiveScopes       { Get-Cached 'AdaptiveScopes'     { Invoke-Step 'Adaptive scopes' -Requires Get-AdaptiveScope -Action { Get-AdaptiveScope } } }
function Get-AppRetentionPolicies { Get-Cached 'AppRetention'       { Invoke-Step 'Retention policies for other apps' -Requires Get-AppRetentionCompliancePolicy -Action { Get-AppRetentionCompliancePolicy } } }

$Script:RetentionLocationNames = [ordered]@{
    ExchangeLocation = 'Exchange mailboxes'; SharePointLocation = 'SharePoint sites'; OneDriveLocation = 'OneDrive accounts'; ModernGroupLocation = 'Microsoft 365 group mailboxes and sites'
    TeamsChatLocation = 'Teams chats and Copilot interactions'; TeamsChannelLocation = 'Teams channel messages'; PublicFolderLocation = 'Exchange public folders'; SkypeLocation = 'Skype for Business'
}
$Script:CoreRetentionLocations = [ordered]@{ ExchangeLocation = 'Exchange mailboxes'; SharePointLocation = 'SharePoint sites'; OneDriveLocation = 'OneDrive accounts'; TeamsChatLocation = 'Teams chats'; TeamsChannelLocation = 'Teams channel messages' }

function Get-DurationText {
    <# Days (or Unlimited) as plain English. #>
    param([string]$Days)
    if (-not $Days) { return '' }
    if ($Days -match 'Unlimited') { return 'forever' }
    $d = 0; if (-not [int]::TryParse(($Days -replace '[^\d]', ''), [ref]$d)) { return $Days }
    if ($d -ge 365 -and $d % 365 -eq 0) { return (Get-Plural ($d / 365) 'year' 'years') }
    if ($d -ge 30 -and $d % 30 -eq 0 -and $d -lt 365) { return (Get-Plural ($d / 30) 'month' 'months') }
    return (Get-Plural $d 'day' 'days')
}

function Get-RetentionText {
    param([string]$Action, [string]$Duration, [string]$From)
    $from = switch -regex ($From) { 'Creat' { 'when created' } 'Modif' { 'when last modified' } 'Tagged|Label' { 'when labelled' } 'Event' { 'an event' } default { '' } }
    $dur = Get-DurationText $Duration
    $fromText = if ($from) { " from $from" } else { '' }
    switch -regex ($Action) {
        '^KeepAndDelete$' { return "Keep for $dur$fromText, then delete" }
        '^Keep$' { return $(if ($dur -eq 'forever') { 'Keep forever' } else { "Keep for $dur$fromText" }) }
        '^Delete$' { return "Delete after $dur$fromText" }
        default { return $(if ($dur) { "$dur$fromText" } else { 'Classification only (no retention)' }) }
    }
}

function Get-TagNames {
    <# Label names from PublishComplianceTag / ApplyComplianceTag values. #>
    param($Value, [hashtable]$Map)
    $t = [string](ConvertTo-CellValue $Value)
    if (-not $t) { return @() }
    @($t -split '\s*[,|]\s*' | Where-Object { $_ } | ForEach-Object { if ($Map.ContainsKey($_)) { $Map[$_] } elseif ($_ -match '^[0-9a-f\-]{36}$') { $null } else { $_ } } | Where-Object { $_ } | Select-Object -Unique)
}

function Get-RetentionLabelMap {
    $map = @{}
    foreach ($t in @(Get-RetentionLabels)) { foreach ($k in @([string](Get-Val $t 'Guid'), [string](Get-Val $t 'ImmutableId'), [string](Get-Val $t 'Identity'), [string]$t.Name)) { if ($k -and -not $map.ContainsKey($k)) { $map[$k] = [string]$t.Name } } }
    return $map
}

function Get-RetentionPolicyView {
    Get-Cached 'RetentionPolicyView' {
        $rules = @(Get-RetentionRules); $map = Get-RetentionLabelMap
        foreach ($p in @(Get-RetentionPolicies)) {
            $ids = @([string](Get-Val $p 'Guid'), [string](Get-Val $p 'Identity'), [string](Get-Val $p 'Id'), [string]$p.Name) | Where-Object { $_ }
            $pr = @($rules | Where-Object { $ids -contains [string](Get-Val $_ 'Policy') -or $ids -contains [string](Get-Val $_ 'ParentPolicyName') })
            $publish = @(foreach ($r in $pr) { Get-TagNames (Get-Val $r 'PublishComplianceTag') $map }) | Select-Object -Unique
            $apply = @(foreach ($r in $pr) { Get-TagNames (Get-Val $r 'ApplyComplianceTag') $map }) | Select-Object -Unique
            $type = if (@($publish).Count) { 'Publishes retention labels' } elseif (@($apply).Count) { 'Auto-applies a retention label' } else { 'Retention policy' }
            $setting = switch ($type) {
                'Publishes retention labels' { 'Labels: ' + (@($publish) -join ', ') }
                'Auto-applies a retention label' {
                    $what = @(foreach ($r in $pr) {
                            foreach ($c in @(Find-RuleClassifiers (Get-Val $r 'ContentContainsSensitiveInformation') ([string]$p.Name) ([string]$r.Name) 'Condition')) { [string]$c.Name }
                            if (Get-Val $r 'ContentMatchQuery') { 'a keyword query' }
                        }) | Where-Object { $_ } | Select-Object -Unique
                    'Label: ' + (@($apply) -join ', ') + $(if (@($what).Count) { '; looks for ' + (@($what) -join ', ') } else { '' })
                }
                default {
                    $r = $pr | Where-Object { Get-Val $_ 'RetentionDuration' } | Select-Object -First 1
                    if ($r) { Get-RetentionText ([string](Get-Val $r 'RetentionComplianceAction')) ([string](Get-Val $r 'RetentionDuration')) ([string](Get-Val $r 'ExpirationDateOption')) } else { 'See 102_Retention_Rules_Full' }
                }
            }
            $locs = @(); $locKeys = @(); $locAll = @()
            foreach ($k in $Script:RetentionLocationNames.Keys) {
                $t = Get-ItemText (Get-Val $p $k)
                if ($t) {
                    $locKeys += $k
                    if ($t -eq 'All') { $locAll += $k; $locs += "$($Script:RetentionLocationNames[$k]) (all)" } else { $locs += "$($Script:RetentionLocationNames[$k]) (selected)" }
                }
            }
            $scopes = Get-ItemText (Get-Val $p 'AdaptiveScopeLocation')
            if ($scopes) { $locs += "Adaptive scopes: $scopes" }
            $enabled = Get-Val $p 'Enabled'
            [pscustomobject][ordered]@{
                Name = [string]$p.Name; Type = $type; Enabled = $enabled; Mode = [string](Get-Val $p 'Mode')
                Locations = ($locs -join '; '); LocationKeys = ($locKeys -join ','); LocationAllKeys = ($locAll -join ','); AdaptiveScopes = $scopes; Setting = $setting
                PreservationLock = [bool](Get-Val $p 'RestrictiveRetention'); DistributionStatus = [string](Get-Val $p 'DistributionStatus')
                Labels = ((@($publish) + @($apply)) -join ' | '); RuleCount = $pr.Count; WhenChangedUTC = Get-Val $p 'WhenChangedUTC'
            }
        }
    }
}

function Get-RetentionLabelView {
    Get-Cached 'RetentionLabelView' {
        $pv = @(Get-RetentionPolicyView)
        $published = @($pv | Where-Object Type -eq 'Publishes retention labels' | ForEach-Object { $_.Labels -split ' \| ' })
        $applied = @($pv | Where-Object Type -eq 'Auto-applies a retention label' | ForEach-Object { $_.Labels -split ' \| ' })
        foreach ($t in @(Get-RetentionLabels)) {
            $n = [string]$t.Name
            $how = @(); if ($published -contains $n) { $how += 'Published' }; if ($applied -contains $n) { $how += 'Auto-applied' }
            $ev = [string](ConvertTo-CellValue (Get-Val $t 'EventType'))
            $rev = [string](ConvertTo-CellValue (Get-Val $t 'ReviewerEmail'))
            $kind = if ((Get-Val $t 'Regulatory') -eq $true) { 'Regulatory record' } elseif ((Get-Val $t 'IsRecordLabel') -eq $true) { 'Record' } else { 'Retention label' }
            [pscustomobject][ordered]@{
                Name = $n; Kind = $kind
                Retention = Get-RetentionText ([string](Get-Val $t 'RetentionAction')) ([string](Get-Val $t 'RetentionDuration')) ([string](Get-Val $t 'RetentionType'))
                EventBased = [bool]($ev -or [string](Get-Val $t 'RetentionType') -match 'Event'); EventType = $ev
                DispositionReview = [bool]$rev; ReviewerCount = $(if ($rev) { @($rev -split '\s*[|,;]\s*' | Where-Object { $_ }).Count } else { 0 })
                HowApplied = $(if ($how.Count) { $how -join ', ' } else { 'Not in any policy' })
                FilePlan = [bool](ConvertTo-CellValue (Get-Val $t 'FilePlanMetadata')); Notes = [string](Get-Val $t 'Notes')
                RetentionAction = [string](Get-Val $t 'RetentionAction'); RetentionDuration = [string](Get-Val $t 'RetentionDuration'); RetentionType = [string](Get-Val $t 'RetentionType')
                WhenChangedUTC = Get-Val $t 'WhenChangedUTC'
            }
        }
    }
}

function Get-AdaptiveScopeView {
    Get-Cached 'AdaptiveScopeView' {
        $pv = @(Get-RetentionPolicyView)
        foreach ($s in @(Get-AdaptiveScopes)) {
            $q = [string](Get-Val $s 'RawQuery'); if (-not $q) { $q = [string](ConvertTo-CellValue (Get-Val $s 'FilterConditions')) }
            [pscustomobject][ordered]@{
                Name = [string]$s.Name; Type = [string](Get-Val $s 'LocationType'); Query = $q
                UsedBy = ((@($pv | Where-Object { $_.AdaptiveScopes -and ($_.AdaptiveScopes -split '\s*[,;|]\s*') -contains [string]$s.Name } | ForEach-Object Name)) -join ' | ')
                WhenChangedUTC = Get-Val $s 'WhenChangedUTC'
            }
        }
    }
}

# ---------- Risk and compliance: insider risk, communication compliance, information barriers ----------

function Get-InsiderRiskPoliciesRaw { Get-Cached 'InsiderRiskRaw'  { Invoke-Step 'Insider risk policies' -Requires Get-InsiderRiskPolicy -Action { Get-InsiderRiskPolicy } } }
function Get-CommCompliancePolicies { Get-Cached 'CommCompPol'     { Invoke-Step 'Communication compliance policies' -Requires Get-SupervisoryReviewPolicyV2 -Action { Get-SupervisoryReviewPolicyV2 } } }
function Get-CommComplianceRules    { Get-Cached 'CommCompRule'    { Invoke-Step 'Communication compliance rules' -Requires Get-SupervisoryReviewRule -Action { Get-SupervisoryReviewRule } } }
function Get-IbPolicies             { Get-Cached 'IbPolicies'      { Invoke-Step 'Information barrier policies' -Requires Get-InformationBarrierPolicy -Action { Get-InformationBarrierPolicy } } }
function Get-OrgSegments            { Get-Cached 'OrgSegments'     { Invoke-Step 'Organisation segments' -Requires Get-OrganizationSegment -Action { Get-OrganizationSegment } } }
function Get-IbStatus               { Get-Cached 'IbStatus'        { Invoke-Step 'Information barrier application status' -Requires Get-InformationBarrierPoliciesApplicationStatus -Action { Get-InformationBarrierPoliciesApplicationStatus } } }

function Get-InsiderRiskTemplateName {
    param([string]$Scenario)
    switch -regex ($Scenario) {
        'Theft|Departing' { 'Data theft by departing users' }
        'PriorityUser|HighValue' { 'Data leaks by priority users' }
        'RiskyUser|Disgruntled' { 'Data leaks by risky users' }
        'Leak' { 'Data leaks' }
        'Security' { 'Security policy violations' }
        'Health|Patient' { 'Patient data misuse' }
        'Browser' { 'Risky browser usage' }
        'Agent' { 'Risky agents' }
        'RiskyAI|AIUsage|GenAI' { 'Risky AI usage' }
        'Forensic' { 'Forensic evidence' }
        default { ConvertFrom-CamelCase $Scenario }
    }
}

function Get-InsiderRiskView {
    Get-Cached 'InsiderRiskView' {
        foreach ($p in @(Get-InsiderRiskPoliciesRaw)) {
            $sc = [string](Get-Val $p 'InsiderRiskScenario')
            if ($sc -eq 'TenantSetting' -or [string]$p.Name -like 'IRM_Tenant_Setting*') { continue }   # tenant settings object, not a policy
            $en = Get-Val $p 'Enabled'; $mode = [string](Get-Val $p 'Mode')
            $state = if ($null -ne $en) { $(if ($en) { 'On' } else { 'Off' }) } elseif ($mode) { $mode } else { 'Unknown' }
            [pscustomobject][ordered]@{ Name = [string]$p.Name; Template = Get-InsiderRiskTemplateName $sc; Scenario = $sc; State = $state; WhenChangedUTC = Get-Val $p 'WhenChangedUTC' }
        }
    }
}

function Get-CommComplianceView {
    Get-Cached 'CommCompView' {
        $rules = @(Get-CommComplianceRules)
        foreach ($p in @(Get-CommCompliancePolicies)) {
            $ids = @([string](Get-Val $p 'Guid'), [string](Get-Val $p 'Identity'), [string]$p.Name) | Where-Object { $_ }
            $pr = @($rules | Where-Object { $ids -contains [string](Get-Val $_ 'Policy') })
            $rev = @(Get-Val $p 'Reviewers' | Where-Object { $_ })
            $en = Get-Val $p 'Enabled'
            $sample = @($pr | ForEach-Object { Get-Val $_ 'SamplingRate' } | Where-Object { $null -ne $_ } | Select-Object -Unique)
            $cond = @($pr | ForEach-Object { [string](ConvertTo-CellValue (Get-Val $_ 'Condition')) } | Where-Object { $_ })
            [pscustomobject][ordered]@{
                Name = [string]$p.Name; State = $(if ($null -eq $en) { 'Unknown' } elseif ($en) { 'On' } else { 'Off' }); Reviewers = $rev.Count
                ReviewPercent = ($sample -join ', '); RuleCount = $pr.Count
                DetectsSensitiveInfo = [bool]@($cond | Where-Object { $_ -match 'SensitiveInformation|DataClassification' }).Count
                UsesClassifiers = [bool]@($cond | Where-Object { $_ -match 'Classifier|MLModel' }).Count
                WhenChangedUTC = Get-Val $p 'WhenChangedUTC'
            }
        }
    }
}

function Get-IbPolicyView {
    Get-Cached 'IbPolicyView' {
        foreach ($p in @(Get-IbPolicies)) {
            $allowed = Get-ItemText (Get-Val $p 'SegmentsAllowed'); $blocked = Get-ItemText (Get-Val $p 'SegmentsBlocked')
            [pscustomobject][ordered]@{
                Name = [string]$p.Name; Segment = [string](Get-ItemText (Get-Val $p 'AssignedSegment')); State = [string](Get-Val $p 'State')
                Rule = $(if ($blocked) { "Blocked from: $blocked" } elseif ($allowed) { "Only allowed with: $allowed" } else { '' })
                WhenChangedUTC = Get-Val $p 'WhenChangedUTC'
            }
        }
    }
}

# ---------- Audit and eDiscovery ----------

function Get-AuditRetentionPolicies { Get-Cached 'AuditRetention' { Invoke-Step 'Audit log retention policies' -Requires Get-UnifiedAuditLogRetentionPolicy -Action { Get-UnifiedAuditLogRetentionPolicy } } }

function Get-AuditRetentionView {
    Get-Cached 'AuditRetentionView' {
        foreach ($p in @(Get-AuditRetentionPolicies | Sort-Object { [int]('0' + (Get-Val $_ 'Priority')) })) {
            $rt = Get-ItemText (Get-Val $p 'RecordTypes'); $ops = Get-ItemText (Get-Val $p 'Operations'); $users = @(Get-Val $p 'UserIds' | Where-Object { $_ })
            $dur = [string](Get-Val $p 'RetentionDuration')
            $durText = switch -regex ($dur) {
                '^SevenDays$' { '7 days' } '^OneMonth$' { '1 month' } '^ThreeMonths$' { '3 months' } '^SixMonths$' { '6 months' } '^NineMonths$' { '9 months' }
                '^(TwelveMonths|OneYear)$' { '1 year' } '^ThreeYears$' { '3 years' } '^FiveYears$' { '5 years' } '^SevenYears$' { '7 years' } '^TenYears$' { '10 years' } default { $dur }
            }
            $what = @(); if ($rt) { $what += "Record types: $rt" }; if ($ops) { $what += "Activities: $ops" }; if ($users.Count) { $what += "$(Get-Plural $users.Count 'user' 'users')" }
            [pscustomobject][ordered]@{
                Name = [string]$p.Name; Priority = Get-Val $p 'Priority'; AppliesTo = $(if ($what.Count) { $what -join '; ' } else { 'All activity' })
                Retention = $durText; RetentionDuration = $dur; RecordTypes = $rt; Operations = $ops; UserCount = $users.Count
                CoversDlp = [bool]($rt -match 'DLP')
                Enabled = Get-Val $p 'Enabled'; WhenChangedUTC = Get-Val $p 'WhenChangedUTC'
            }
        }
    }
}

function Get-EDiscoveryCases {
    Get-Cached 'EDiscoveryCases' {
        Invoke-Step 'eDiscovery cases' -Requires Get-ComplianceCase -Action {
            $seen = @{}
            foreach ($c in @(Get-ComplianceCase -ErrorAction Stop)) { $k = [string](Get-Val $c 'Identity'); if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $c } }
            try { foreach ($c in @(Get-ComplianceCase -CaseType AdvancedEdiscovery -ErrorAction Stop)) { $k = [string](Get-Val $c 'Identity'); if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $c } } }
            catch { Write-Log "  eDiscovery (Premium) cases not read: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'WARN' }
        }
    }
}

function Get-EDiscoveryCaseView {
    Get-Cached 'EDiscoveryCaseView' {
        foreach ($c in @(Get-EDiscoveryCases)) {
            $t = [string](Get-Val $c 'CaseType')
            [pscustomobject][ordered]@{
                Name = [string]$c.Name; Type = $(if ($t -match 'Advanced|Premium') { 'Premium' } elseif ($t) { 'Standard' } else { 'Unknown' }); Status = [string](Get-Val $c 'Status')
                Created = Get-Val $c 'CreatedDateTime'; Closed = Get-Val $c 'ClosedDateTime'; Identity = [string](Get-Val $c 'Identity')
            }
        }
    }
}

function Get-CaseHolds {
    Get-Cached 'CaseHolds' {
        $active = @(Get-EDiscoveryCaseView | Where-Object { $_.Status -ne 'Closed' })
        if ($active.Count -eq 0) { return }
        $max = $Script:Settings.MaxEDiscoveryCases
        if ($active.Count -gt $max) { Write-Log "  $($active.Count) active cases; holds are read for the first $max (MaxEDiscoveryCases)." 'WARN'; $active = @($active | Select-Object -First $max) }
        Invoke-Step 'eDiscovery case holds' -Requires Get-CaseHoldPolicy -Action {
            foreach ($c in $active) {
                try {
                    foreach ($h in @(Get-CaseHoldPolicy -Case $c.Identity -ErrorAction Stop)) {
                        $query = $false
                        if (Get-Command Get-CaseHoldRule -ErrorAction SilentlyContinue) {
                            try { $query = [bool]@(Get-CaseHoldRule -Policy ([string](Get-Val $h 'Guid')) -ErrorAction Stop | Where-Object { Get-Val $_ 'ContentMatchQuery' }).Count } catch { }
                        }
                        [pscustomobject][ordered]@{
                            Case = $c.Name; CaseType = $c.Type; Hold = [string]$h.Name; Enabled = Get-Val $h 'Enabled'
                            Mailboxes = @(Get-Val $h 'ExchangeLocation' | Where-Object { $_ }).Count; Sites = @(Get-Val $h 'SharePointLocation' | Where-Object { $_ }).Count
                            PublicFolders = [bool]@(Get-Val $h 'PublicFolderLocation' | Where-Object { $_ }).Count; QueryBased = $query
                            DistributionStatus = [string](Get-Val $h 'DistributionStatus'); WhenCreatedUTC = Get-Val $h 'WhenCreatedUTC'; WhenChangedUTC = Get-Val $h 'WhenChangedUTC'
                        }
                    }
                }
                catch { Write-Log "  Holds for one case could not be read: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'WARN' }
            }
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
    Export-Dataset '35_Sensitivity_Label_Settings' @(Get-LabelSettingsView) -AlreadyFlat
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
    param([switch]$WithDlpAudit)
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
    Export-AuditConfig
    if ($WithDlpAudit) { Export-AreaDlpAudit -AlreadyConnected }
    Reset-AfterExo
}

function Export-AreaLifecycle {
    Export-Dataset '100_Retention_Policies' @(Get-RetentionPolicyView) -AlreadyFlat
    Export-Dataset '101_Retention_Policies_Full' @(Get-RetentionPolicies)
    Export-Dataset '102_Retention_Rules_Full' @(Get-RetentionRules)
    Export-Dataset '103_Retention_Labels' @(Get-RetentionLabelView) -AlreadyFlat
    Export-Dataset '104_Retention_Labels_Full' @(Get-RetentionLabels)
    Export-Dataset '105_Adaptive_Scopes' @(Get-AdaptiveScopeView) -AlreadyFlat
    Export-Dataset '106_Retention_Policies_Other_Apps' @(Get-AppRetentionPolicies)
}

function Export-AreaRiskCompliance {
    Export-Dataset '200_Insider_Risk_Policies' @(Get-InsiderRiskView) -AlreadyFlat
    Export-Dataset '201_Insider_Risk_Policies_Full' @(Get-InsiderRiskPoliciesRaw)
    Export-Dataset '210_Communication_Compliance_Policies' @(Get-CommComplianceView) -AlreadyFlat
    Export-Dataset '211_Communication_Compliance_Policies_Full' @(Get-CommCompliancePolicies)
    Export-Dataset '212_Communication_Compliance_Rules_Full' @(Get-CommComplianceRules)
    Export-Dataset '220_Information_Barrier_Policies' @(Get-IbPolicyView) -AlreadyFlat
    Export-Dataset '221_Organisation_Segments' @(Get-OrgSegments | Select-Object Name, UserGroupFilter, WhenChangedUTC)
    if (@(Get-IbPolicies).Count) { Export-Dataset '222_Information_Barrier_Status' @(Get-IbStatus) }
}

function Export-AreaAuditDiscovery {
    <# Security & Compliance part. The unified audit log status needs Exchange Online and is read in the Exchange Online phase (Export-AuditConfig). #>
    Export-Dataset '300_Audit_Retention_Policies' @(Get-AuditRetentionView) -AlreadyFlat
    Export-Dataset '310_eDiscovery_Cases' @(Get-EDiscoveryCaseView) -AlreadyFlat
    Export-Dataset '311_eDiscovery_Case_Holds' @(Get-CaseHolds) -AlreadyFlat
}

function Export-AuditConfig {
    <# Expects Exchange Online to be connected. In Security & Compliance PowerShell UnifiedAuditLogIngestionEnabled always reads False, so Exchange Online is used. #>
    if ($Script:Cache.ContainsKey('Audit')) { return }
    $audit = Invoke-Step 'Unified audit log configuration' -Requires Get-AdminAuditLogConfig -Action { Get-AdminAuditLogConfig | Select-Object UnifiedAuditLogIngestionEnabled, AdminAuditLogEnabled }
    $Script:Cache['Audit'] = @($audit)
    Export-Dataset '71_Audit_Configuration' $audit
}

function Invoke-ExoPhase {
    <# Exchange Online work runs last: connecting to Exchange Online changes the loaded cmdlets. One sign-in covers every Exchange Online area selected. #>
    param([bool]$Exchange, [bool]$DlpAudit, [bool]$AuditConfig)
    if ($Exchange) {
        Write-Host ''; Write-Host "--- $($Script:Areas['12']) ---" -ForegroundColor Cyan
        Export-AreaExchange -WithDlpAudit:$DlpAudit
        return
    }
    if (-not ($DlpAudit -or $AuditConfig)) { return }
    Write-Host ''
    Write-Host $(if ($DlpAudit) { "--- $($Script:Areas['14']) ---" } else { '--- Unified audit log status (Exchange Online) ---' }) -ForegroundColor Cyan
    try { Connect-Exo }
    catch {
        Write-Log "Exchange Online connection failed: $($_.Exception.Message)" 'ERROR'
        if ($DlpAudit) { $Script:Errors.Add([pscustomobject]@{ Step = "DLP rule activity from the audit log ($($Script:Settings.AuditDays) days)"; Outcome = 'Failed'; Error = $_.Exception.Message }) }
        if ($AuditConfig) { $Script:Errors.Add([pscustomobject]@{ Step = 'Unified audit log configuration'; Outcome = 'Failed'; Error = $_.Exception.Message }) }
        return
    }
    if ($DlpAudit) { Export-AreaDlpAudit -AlreadyConnected }
    if ($AuditConfig) { Export-AuditConfig }
    Reset-AfterExo
}

# ---------- Content Explorer and audit-log evidence (Security & Compliance / Exchange Online PowerShell; no Graph) ----------

# Common built-in types also counted by Content Explorer, so the report can show sensitive data that no DLP rule covers.
# Only names that exist in the tenant are queried.
$Script:CommonSits = @(
    'Credit Card Number', 'International Banking Account Number (IBAN)', 'EU Debit Card Number', 'SWIFT Code',
    'U.K. National Insurance Number (NINO)', 'U.K. National Health Service Number', 'U.S. / U.K. Passport Number',
    'U.K. Unique Taxpayer Reference Number', 'U.S. Social Security Number (SSN)', 'U.S. Bank Account Number',
    'General Password', 'Azure Storage Account Key', 'Azure SQL connection string'
)
$Script:CeWorkloads = [ordered]@{ Exchange = 'EXO'; SharePoint = 'SPO'; OneDrive = 'ODB'; Teams = 'Teams' }

function Get-ContentExplorerCounts {
    Get-Cached 'ContentExplorer' {
        Invoke-Step 'Content Explorer item counts' -Requires Export-ContentExplorerData -Action {
            $known = @{}; foreach ($s in @(Get-SITs)) { $known[[string]$s.Name] = $true }
            $inRules = @(Get-RuleClassifiers | Where-Object { $_.ReferenceType -eq 'Classifier' -and $_.ClassifierType -eq 'Content (SIT)' } | ForEach-Object { [string]$_.Name } | Select-Object -Unique)
            $names = @(@($inRules) + @($Script:CommonSits | Where-Object { $known.ContainsKey($_) }) | Select-Object -Unique | Select-Object -First $Script:Settings.ContentExplorerMaxTags)
            $tags = @(foreach ($n in $names) { [pscustomobject]@{ TagType = 'SensitiveInformationType'; TagName = $n; UsedInDlpRules = ($inRules -contains $n) } })
            foreach ($l in @(Get-LabelView)) {
                $ln = if ($l.DisplayName) { [string]$l.DisplayName } else { [string]$l.Name }
                $tags += [pscustomobject]@{ TagType = 'Sensitivity'; TagName = $ln; UsedInDlpRules = [bool]$l.UsedInDlpRules }
            }
            $ok = 0; $fail = 0; $i = 0
            foreach ($t in $tags) {
                $i++
                Write-Progress -Activity 'Content Explorer' -Status $t.TagName -PercentComplete ([int](100 * $i / [Math]::Max(1, $tags.Count)))
                foreach ($w in $Script:CeWorkloads.Keys) {
                    $count = $null; $err = ''
                    try {
                        $res = @(Export-ContentExplorerData -TagType $t.TagType -TagName $t.TagName -Workload $Script:CeWorkloads[$w] -PageSize 1 -ErrorAction Stop)
                        $hdr = $res | Where-Object { $_ -and $_.PSObject.Properties['TotalCount'] } | Select-Object -First 1
                        if ($hdr) { $count = [int64]("0" + [string]$hdr.TotalCount) }
                        $ok++
                    }
                    catch {
                        $err = $_.Exception.Message.Split([Environment]::NewLine)[0]; $fail++
                        if ($ok -eq 0 -and $fail -ge 4) { Write-Progress -Activity 'Content Explorer' -Completed; throw "Content Explorer queries failed ($err). The account needs the Content Explorer List Viewer role." }
                    }
                    [pscustomobject]@{ TagType = $t.TagType; TagName = $t.TagName; Workload = $w; ItemCount = $count; UsedInDlpRules = $t.UsedInDlpRules; Error = $err }
                }
            }
            Write-Progress -Activity 'Content Explorer' -Completed
        }
    }
}

function Get-LocationKey {
    <# Picks a readable location (site or mailbox) from a Content Explorer record, whatever the property is called. #>
    param($Record)
    foreach ($p in 'SiteUrl', 'Site', 'UserPrincipalName', 'Upn', 'Mailbox', 'Location', 'FileUrl', 'Url', 'Path', 'FilePath') {
        $v = [string](Get-Val $Record $p)
        if (-not $v) { continue }
        $m = [regex]::Match($v, '^(https://[^/]+/(sites|teams|personal)/[^/?#]+)', 'IgnoreCase')
        if ($m.Success) { return $m.Groups[1].Value }
        return $v
    }
    return ''
}

function Get-ContentExplorerLocations {
    <# For the most common sensitive information types, samples records to show where the data is concentrated. #>
    Get-Cached 'ContentExplorerLocations' {
        $counts = @(Get-ContentExplorerCounts | Where-Object { $_.TagType -eq 'SensitiveInformationType' -and $_.ItemCount -gt 0 -and $_.Workload -ne 'Teams' })
        if ($counts.Count -eq 0) { return }
        Invoke-Step 'Content Explorer top locations' -Requires Export-ContentExplorerData -Action {
            foreach ($c in ($counts | Sort-Object ItemCount -Descending | Select-Object -First 8)) {
                try {
                    $res = @(Export-ContentExplorerData -TagType 'SensitiveInformationType' -TagName $c.TagName -Workload $Script:CeWorkloads[$c.Workload] -PageSize 500 -ErrorAction Stop)
                    $recs = @($res | Where-Object { $_ -and -not $_.PSObject.Properties['TotalCount'] })
                    $grp = @($recs | ForEach-Object { Get-LocationKey $_ } | Where-Object { $_ } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5)
                    foreach ($g in $grp) {
                        [pscustomobject]@{ TagName = $c.TagName; Workload = $c.Workload; Location = $g.Name; ItemsInSample = $g.Count; SampleSize = $recs.Count; TotalItems = $c.ItemCount }
                    }
                }
                catch { Write-Log "  Top locations for $($c.TagName) ($($c.Workload)) not available: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'WARN' }
            }
        }
    }
}

function ConvertTo-DlpAuditRows {
    <# Turns one unified audit log record (DLP record types) into one row per matched policy rule. #>
    param($Record, [string]$RecordType)
    $d = $null
    try { $d = [string]$Record.AuditData | ConvertFrom-Json } catch { return }
    if (-not $d) { return }
    $wl = [string](Get-Val $d 'Workload')
    $workload = switch -Regex ($wl) {
        'Teams' { 'Teams'; break } 'OneDrive' { 'OneDrive'; break } 'SharePoint' { 'SharePoint'; break } 'Exchange' { 'Exchange'; break } 'Endpoint' { 'Devices'; break }
        default { switch ($RecordType) { 'ComplianceDLPExchange' { 'Exchange' } 'DLPEndpoint' { 'Devices' } default { 'SharePoint' } } }
    }
    $ex = Get-Val $d 'ExceptionInfo'
    $reason = [string](Get-Val $ex 'Reason'); $just = [string](Get-Val $ex 'Justification')
    $fp = [string](Get-Val $ex 'FalsePositive')
    $em = Get-Val $d 'ExchangeMetaData'; $sm = Get-Val $d 'SharePointMetaData'
    $item = [string](Get-Val $em 'Subject')
    if (-not $item) { $item = [string](Get-Val $sm 'FileName') }
    if (-not $item) { $item = [string](Get-Val $sm 'FilePathUrl') }
    if (-not $item) { $item = [string](Get-Val $d 'ObjectId') }
    $rcpt = @(@(Get-Val $em 'To') + @(Get-Val $em 'CC') + @(Get-Val $em 'BCC') | Where-Object { $_ } | ForEach-Object { ([string]$_ -split '@')[-1].Trim().ToLower() } | Where-Object { $_ -and $_ -notmatch '\s' } | Select-Object -Unique)
    $common = @{
        Item = $(if ($item.Length -gt 250) { $item.Substring(0, 250) } else { $item }); RecipientDomains = ($rcpt -join ', ')
        Date = [string](Get-Val $d 'CreationTime'); Operation = [string](Get-Val $d 'Operation'); Workload = $workload; User = [string](Get-Val $d 'UserId')
        Override = [bool]($reason -match 'Override' -or $just); FalsePositive = [bool]($fp -eq 'True' -or $reason -match 'FalsePositive'); Justification = $just
    }
    $emitted = $false
    foreach ($pd in @(Get-Val $d 'PolicyDetails')) {
        if ($null -eq $pd) { continue }
        foreach ($r in @(Get-Val $pd 'Rules')) {
            if ($null -eq $r) { continue }
            $cm = Get-Val $r 'ConditionsMatched'
            $sits = @(foreach ($si in @(Get-Val $cm 'SensitiveInformation')) { [string](Get-Val $si 'SensitiveInformationTypeName') }) | Where-Object { $_ } | Select-Object -Unique
            $row = [ordered]@{ Date = $common.Date; Operation = $common.Operation; Workload = $common.Workload; User = $common.User
                Policy = [string](Get-Val $pd 'PolicyName'); Rule = [string](Get-Val $r 'RuleName'); Severity = [string](Get-Val $r 'Severity')
                Actions = (@(Get-Val $r 'Actions') -join ', '); SensitiveTypes = (@($sits) -join ', ')
                Override = $common.Override; FalsePositive = $common.FalsePositive; Justification = $common.Justification; Item = $common.Item; RecipientDomains = $common.RecipientDomains }
            [pscustomobject]$row; $emitted = $true
        }
    }
    foreach ($pm in @(Get-Val $d 'PolicyMatchInfo')) {
        if ($null -eq $pm) { continue }
        [pscustomobject][ordered]@{ Date = $common.Date; Operation = $common.Operation; Workload = $common.Workload; User = $common.User
            Policy = [string](Get-Val $pm 'PolicyName'); Rule = [string](Get-Val $pm 'RuleName'); Severity = [string](Get-Val $d 'Severity')
            Actions = [string](Get-Val $d 'EnforcementMode'); SensitiveTypes = ''
            Override = $common.Override; FalsePositive = $common.FalsePositive; Justification = $common.Justification; Item = $common.Item; RecipientDomains = $common.RecipientDomains }
        $emitted = $true
    }
    if (-not $emitted -and $common.Operation) {
        [pscustomobject][ordered]@{ Date = $common.Date; Operation = $common.Operation; Workload = $common.Workload; User = $common.User
            Policy = ''; Rule = ''; Severity = ''; Actions = ''; SensitiveTypes = ''
            Override = $common.Override; FalsePositive = $common.FalsePositive; Justification = $common.Justification; Item = $common.Item; RecipientDomains = $common.RecipientDomains }
    }
}

function Get-DlpAuditRecords {
    Get-Cached 'DlpAudit' {
        $days = $Script:Settings.AuditDays
        Invoke-Step "DLP rule activity from the audit log ($days days)" -Requires Search-UnifiedAuditLog -Action {
            $end = (Get-Date).ToUniversalTime(); $start = $end.AddDays(-$days)
            foreach ($rt in 'ComplianceDLPExchange', 'ComplianceDLPSharePoint', 'DLPEndpoint') {
                $sid = [guid]::NewGuid().ToString(); $n = 0
                try {
                    while ($true) {
                        $page = @(Search-UnifiedAuditLog -StartDate $start -EndDate $end -RecordType $rt -SessionId $sid -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
                        if ($page.Count -eq 0) { break }
                        foreach ($p in $page) { ConvertTo-DlpAuditRows -Record $p -RecordType $rt }
                        $n += $page.Count
                        $last = $page[-1]
                        $idx = [int]("0" + [string](Get-Val $last 'ResultIndex')); $total = [int]("0" + [string](Get-Val $last 'ResultCount'))
                        if ($n -ge $Script:Settings.AuditMaxRecords -or ($total -gt 0 -and $idx -ge $total) -or $page.Count -lt 5000) { break }
                    }
                    Write-Log "  $($rt): $n audit record(s)"
                    if ($n -ge $Script:Settings.AuditMaxRecords) { Write-Log "  $rt record cap ($($Script:Settings.AuditMaxRecords)) reached; counts are partial." 'WARN' }
                }
                catch { Write-Log "  $rt search failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'WARN' }
            }
        }
    }
}

function Get-DlpRuleActivity {
    <# Summarises audit rows per policy rule: matches, overrides, false positives, distinct users. #>
    Get-Cached 'DlpRuleActivity' {
        $rows = @(Get-DlpAuditRecords | Where-Object { $_.Policy })
        foreach ($g in ($rows | Group-Object { "$($_.Policy)|$($_.Rule)" })) {
            $pr = $g.Name -split '\|', 2
            $hits = @($g.Group | Where-Object { $_.Operation -notmatch 'Undo|DlpInfo' })
            $last = ''
            try { $last = (@($g.Group | ForEach-Object { [datetime]$_.Date }) | Sort-Object -Descending | Select-Object -First 1).ToString('dd MMM yyyy') } catch { }
            [pscustomobject][ordered]@{
                Policy = $pr[0]; Rule = $pr[1]; Matches = $hits.Count
                Overrides = @($g.Group | Where-Object Override).Count; FalsePositives = @($g.Group | Where-Object FalsePositive).Count
                Undone = @($g.Group | Where-Object { $_.Operation -match 'Undo' }).Count
                DistinctUsers = @($g.Group | Select-Object -ExpandProperty User -Unique | Where-Object { $_ }).Count
                Workloads = (@($g.Group | Select-Object -ExpandProperty Workload -Unique) -join ', '); LastMatch = $last
            }
        }
    }
}

function Export-AreaContentExplorer {
    Export-Dataset '62_Content_Explorer_Counts' @(Get-ContentExplorerCounts) -AlreadyFlat
    Export-Dataset '63_Content_Explorer_Top_Locations' @(Get-ContentExplorerLocations) -AlreadyFlat
}

function Export-AreaDlpAudit {
    param([switch]$AlreadyConnected)
    if (-not $AlreadyConnected) {
        try { Connect-Exo }
        catch {
            Write-Log "Exchange Online connection failed: $($_.Exception.Message)" 'ERROR'
            $Script:Errors.Add([pscustomobject]@{ Step = "DLP rule activity from the audit log ($($Script:Settings.AuditDays) days)"; Outcome = 'Failed'; Error = $_.Exception.Message })
            return
        }
    }
    Export-Dataset '64_DLP_Audit_Rule_Matches' @(Get-DlpAuditRecords) -AlreadyFlat
    Export-Dataset '65_DLP_Rule_Activity_Summary' @(Get-DlpRuleActivity) -AlreadyFlat
    Export-AreaDlpAlerts
    if (-not $AlreadyConnected) { Reset-AfterExo }
}

# ---------- DLP alerts from the unified audit log (SecurityComplianceAlerts) and the alerts dashboard ----------

function Get-AcceptedDomainList {
    Get-Cached 'AcceptedDomains' {
        if (Get-Command Get-AcceptedDomain -ErrorAction SilentlyContinue) {
            try { @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { ([string]$_.DomainName).ToLower() }) } catch { Write-Log "  Accepted domains not available: $($_.Exception.Message)" 'WARN' }
        }
    }
}

function Get-DataValue {
    <# Reads the first matching key from an alert's Data JSON (key names vary between alert types). #>
    param($Data, [string[]]$Keys)
    if ($null -eq $Data) { return '' }
    foreach ($k in $Keys) { $v = Get-Val $Data $k; if ($v) { return (ConvertTo-CellValue $v) } }
    return ''
}

function Get-DlpAlertRecords {
    Get-Cached 'DlpAlertEvents' {
        $days = $Script:Settings.AuditDays
        Invoke-Step "DLP alerts from the audit log ($days days)" -Requires Search-UnifiedAuditLog -Action {
            $end = (Get-Date).ToUniversalTime(); $start = $end.AddDays(-$days)
            $sid = [guid]::NewGuid().ToString(); $n = 0
            while ($true) {
                $page = @(Search-UnifiedAuditLog -StartDate $start -EndDate $end -RecordType SecurityComplianceAlerts -SessionId $sid -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
                if ($page.Count -eq 0) { break }
                foreach ($rec in $page) {
                    $d = $null; try { $d = [string]$rec.AuditData | ConvertFrom-Json } catch { continue }
                    $cat = [string](Get-Val $d 'Category'); $name = [string](Get-Val $d 'Name')
                    if ($cat -notmatch 'DataLossPrevention|DLP' -and $name -notmatch '^DLP|data loss') { continue }
                    $data = $null; $raw = [string](Get-Val $d 'Data')
                    if ($raw) { try { $data = $raw | ConvertFrom-Json } catch { } }
                    $etype = [string](Get-Val $d 'EntityType'); $eid = [string](Get-Val $d 'AlertEntityId')
                    $user = if ($etype -match 'User|Mailbox|Account' -and $eid) { $eid } else { Get-DataValue $data @('f3u', 'upn', 'UserId', 'eid', 'au', 'ui') }
                    [pscustomobject][ordered]@{
                        Time = [string](Get-Val $d 'CreationTime'); Operation = [string](Get-Val $d 'Operation'); AlertId = [string](Get-Val $d 'AlertId')
                        Name = $name; Severity = [string](Get-Val $d 'Severity'); Status = [string](Get-Val $d 'Status'); Category = $cat
                        PolicyId = [string](Get-Val $d 'PolicyId'); EntityType = $etype; User = $user
                        Workload = Get-DataValue $data @('wl', 'Workload'); Comments = [string](Get-Val $d 'Comments')
                        Data = $(if ($raw.Length -gt 2000) { $raw.Substring(0, 2000) } else { $raw })
                    }
                }
                $n += $page.Count
                $last = $page[-1]
                $idx = [int]("0" + [string](Get-Val $last 'ResultIndex')); $total = [int]("0" + [string](Get-Val $last 'ResultCount'))
                if ($n -ge $Script:Settings.AuditMaxRecords -or ($total -gt 0 -and $idx -ge $total) -or $page.Count -lt 5000) { break }
            }
        }
    }
}

function Get-AlertRuleName {
    <# DLP alert policies created from rules are named "DLP-<rule name>" (sometimes cut short with an ID suffix). #>
    param([string]$AlertName, [string[]]$RuleNames)
    $n = ($AlertName -replace '^DLP-', '' -replace '-[0-9a-f]{8}$', '').Trim()
    $hit = $RuleNames | Where-Object { $_ -eq $n } | Select-Object -First 1
    if (-not $hit -and $n.Length -ge 20) { $hit = $RuleNames | Where-Object { $_.StartsWith($n) -or $n.StartsWith($_) } | Select-Object -First 1 }
    if ($hit) { return $hit } else { return $n }
}

function Get-DlpAlertSummary {
    <# One row per alert, linked to the DLP rule match events behind it (same rule and user, up to an hour before). #>
    Get-Cached 'DlpAlerts' {
        $events = @(Get-DlpAlertRecords)
        if ($events.Count -eq 0) { return }
        $ruleNames = @(); if ($Script:Cache.ContainsKey('Rules')) { $ruleNames = @($Script:Cache['Rules'] | ForEach-Object { [string]$_.Name }) }
        $ruleToPolicy = @{}; if ($Script:Cache.ContainsKey('Rules')) { foreach ($r in $Script:Cache['Rules']) { $ruleToPolicy[[string]$r.Name] = [string](Get-Val $r 'ParentPolicyName') } }
        # index rule-match events by rule and by rule|user, with parsed times, so linking stays fast in large tenants
        $byRuleUser = @{}; $byRule = @{}
        foreach ($m in @(if ($Script:Cache.ContainsKey('DlpAudit')) { $Script:Cache['DlpAudit'] | Where-Object { $_.Rule } })) {
            $t = ConvertTo-DateSafe $m.Date; if (-not $t) { continue }
            $e = [pscustomobject]@{ T = $t; Ev = $m }
            $k1 = "$($m.Rule)|$($m.User)".ToLower(); $k2 = ([string]$m.Rule).ToLower()
            if (-not $byRuleUser.ContainsKey($k1)) { $byRuleUser[$k1] = [System.Collections.Generic.List[object]]::new() }; $byRuleUser[$k1].Add($e)
            if (-not $byRule.ContainsKey($k2)) { $byRule[$k2] = [System.Collections.Generic.List[object]]::new() }; $byRule[$k2].Add($e)
        }
        $accepted = @(Get-AcceptedDomainList)
        foreach ($g in ($events | Where-Object AlertId | Group-Object AlertId)) {
            $ev = @($g.Group | Sort-Object { ConvertTo-DateSafe $_.Time })
            $first = $ev | Where-Object { $_.Operation -match 'Triggered' } | Select-Object -First 1
            if (-not $first) { $first = $ev[0] }
            $lastStatus = ($ev | Where-Object Status | Select-Object -Last 1).Status
            $raised = ConvertTo-DateSafe $first.Time
            $rule = Get-AlertRuleName ([string]$first.Name) $ruleNames
            $users = @($ev | ForEach-Object User | Where-Object { $_ } | Select-Object -Unique)
            $cand = @(if ($users.Count) { foreach ($u in $users) { $k = "$rule|$u".ToLower(); if ($byRuleUser.ContainsKey($k)) { $byRuleUser[$k] } } } else { $k = $rule.ToLower(); if ($byRule.ContainsKey($k)) { $byRule[$k] } })
            $rel = @(if ($raised) { $cand | Where-Object { $_.T -le $raised.AddMinutes(5) -and $_.T -ge $raised.AddMinutes(-60) } | ForEach-Object Ev })
            $wl = @(@($ev | ForEach-Object Workload) + @($rel | ForEach-Object Workload) | Where-Object { $_ } | Select-Object -Unique)
            $sits = @($rel | ForEach-Object { $_.SensitiveTypes -split ', ' } | Where-Object { $_ } | Select-Object -Unique)
            $items = @($rel | ForEach-Object Item | Where-Object { $_ } | Select-Object -Unique)
            $doms = @($rel | ForEach-Object { $_.RecipientDomains -split ', ' } | Where-Object { $_ } | Select-Object -Unique)
            $ext = @($doms | Where-Object { $accepted.Count -eq 0 -or $accepted -notcontains $_.ToLower() })
            [pscustomobject][ordered]@{
                AlertId = $g.Name; Raised = $(if ($raised) { $raised.ToString('yyyy-MM-dd HH:mm') } else { [string]$first.Time })
                Alert = [string]$first.Name; Rule = $rule; Policy = $(if ($ruleToPolicy.ContainsKey($rule)) { $ruleToPolicy[$rule] } else { '' })
                Severity = [string]$first.Severity; Status = $(if ($lastStatus) { [string]$lastStatus } else { 'Unknown' })
                Users = ($users -join ', '); Workloads = ($wl -join ', '); SensitiveTypes = ($sits -join ', ')
                Items = ((@($items | Select-Object -First 5)) -join ' | '); ExternalDomains = ($ext -join ', '); LinkedMatches = $rel.Count
            }
        }
    }
}

function ConvertTo-HtmlText { param($Text) [System.Net.WebUtility]::HtmlEncode([string]$Text) }

function New-HtmlBars {
    <# Horizontal bar list. Items: objects with Label, Value, optional Note. #>
    param([object[]]$Items, [string]$Color = '#0f4c81', [int]$Max = 10, [string]$Empty = 'No data')
    $Items = @($Items | Sort-Object Value -Descending | Select-Object -First $Max)
    if ($Items.Count -eq 0) { return "<p class='empty'>$(ConvertTo-HtmlText $Empty)</p>" }
    $top = [double](($Items | Measure-Object Value -Maximum).Maximum); if ($top -le 0) { $top = 1 }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($i in $Items) {
        $pct = [Math]::Max(2, [int](100 * $i.Value / $top))
        $note = if ($i.Note) { "<span class='note'>$(ConvertTo-HtmlText $i.Note)</span>" } else { '' }
        [void]$sb.Append("<div class='bar'><div class='lbl' title='$(ConvertTo-HtmlText $i.Label)'>$(ConvertTo-HtmlText $i.Label)$note</div><div class='trk'><div class='fill' style='width:$pct%;background:$Color'></div></div><div class='val'>$('{0:N0}' -f $i.Value)</div></div>")
    }
    return $sb.ToString()
}

function New-HtmlDailyChart {
    <# Stacked columns per day by severity, as inline SVG. #>
    param([object[]]$Alerts, [datetime]$Start, [datetime]$End)
    $sevColor = [ordered]@{ High = '#c0392b'; Medium = '#e08a00'; Low = '#2e86c1'; Informational = '#8a8a8a' }
    $days = [int][Math]::Ceiling(($End.Date - $Start.Date).TotalDays) + 1
    $byDay = @{}
    foreach ($a in $Alerts) {
        $t = ConvertTo-DateSafe $a.Raised; if (-not $t) { continue }
        $k = $t.ToString('yyyy-MM-dd'); $s = if ($sevColor.Contains([string]$a.Severity)) { [string]$a.Severity } else { 'Informational' }
        if (-not $byDay.ContainsKey($k)) { $byDay[$k] = @{} }
        $byDay[$k][$s] = 1 + [int]$byDay[$k][$s]
    }
    $maxDay = 1; foreach ($v in $byDay.Values) { $sum = 0; foreach ($x in $v.Values) { $sum += $x }; if ($sum -gt $maxDay) { $maxDay = $sum } }
    $w = 900; $h = 180; $pad = 28; $bw = [Math]::Max(2, [Math]::Floor(($w - $pad) / [Math]::Max(1, $days)) - 2)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $w $($h + 30)' class='chart' role='img' aria-label='Alerts per day'>")
    [void]$sb.Append("<line x1='$pad' y1='$h' x2='$w' y2='$h' stroke='#bbb'/><text x='0' y='12' class='ax'>$maxDay</text><text x='0' y='$h' class='ax'>0</text>")
    for ($i = 0; $i -lt $days; $i++) {
        $day = $Start.Date.AddDays($i); $k = $day.ToString('yyyy-MM-dd'); $x = $pad + $i * ($bw + 2); $y = $h
        if ($byDay.ContainsKey($k)) {
            foreach ($s in $sevColor.Keys) {
                $c = [int]$byDay[$k][$s]; if ($c -le 0) { continue }
                $bh = [Math]::Max(1, [int](($h - 10) * $c / $maxDay)); $y -= $bh
                [void]$sb.Append("<rect x='$x' y='$y' width='$bw' height='$bh' fill='$($sevColor[$s])'><title>$k : $c $s</title></rect>")
            }
        }
        if ($i -eq 0 -or $days -le 14 -or ($days -le 45 -and $day.DayOfWeek -eq 'Monday') -or ($days -gt 45 -and $day.Day -eq 1)) {
            [void]$sb.Append("<text x='$x' y='$($h + 16)' class='ax'>$($day.ToString('dd MMM'))</text>")
        }
    }
    [void]$sb.Append('</svg>')
    [void]$sb.Append("<div class='legend'>" + (($sevColor.Keys | ForEach-Object { "<span><i style='background:$($sevColor[$_])'></i>$_</span>" }) -join '') + '</div>')
    return $sb.ToString()
}

function New-DlpAlertDashboard {
    <# Writes DLP_Alerts_Dashboard.html: a single self-contained page (no internet access needed to view it). #>
    $alerts = @(Get-DlpAlertSummary)
    $matches2 = @(if ($Script:Cache.ContainsKey('DlpAudit')) { $Script:Cache['DlpAudit'] | Where-Object { $_.Rule -and $_.Operation -notmatch 'Undo|DlpInfo' } })
    $days = $Script:Settings.AuditDays
    $end = (Get-Date); $start = $end.AddDays(-$days)
    $ti = @($Script:Cache['TenantInfo'])
    $org = if ($Script:Settings.CustomerName) { $Script:Settings.CustomerName } elseif ($ti.Count -and $ti[0].Organization) { $ti[0].Organization } else { 'Microsoft 365 tenant' }

    # optional masking of people
    $mask = -not $Script:Settings.AlertDashboardShowNames
    $map = @{}
    $who = { param($u) if (-not $mask -or -not $u) { return $u }; if (-not $map.ContainsKey($u)) { $map[$u] = "User $($map.Count + 1)" }; return $map[$u] }

    $total = $alerts.Count
    $sev = @{}; foreach ($s in 'High', 'Medium', 'Low', 'Informational') { $sev[$s] = @($alerts | Where-Object Severity -eq $s).Count }
    $open = @($alerts | Where-Object { $_.Status -match 'Active|New|InProgress|Investigating|Unknown' }).Count
    $userList = @($alerts | ForEach-Object { $_.Users -split ', ' } | Where-Object { $_ })
    $distinctUsers = @($userList | Select-Object -Unique).Count
    $ruleGroups = @($alerts | Group-Object Rule | Sort-Object Count -Descending)
    $extAlerts = @($alerts | Where-Object ExternalDomains).Count

    # ---- plain-English summary
    $ins = [System.Collections.Generic.List[string]]::new()
    if ($total -eq 0) {
        $ins.Add("No DLP alerts were raised in the last $days days.")
        if ($matches2.Count) { $ins.Add("DLP rules did match <b>$('{0:N0}' -f $matches2.Count)</b> times in the same period, so the rules are working but are not set to raise alerts (or the alert threshold was not reached). See the rule matches below.") }
        else { $ins.Add('No DLP rule matches were found either. Check the unified audit log is on and that the account can read DLP events.') }
    }
    else {
        $ins.Add("<b>$total</b> DLP alerts were raised in the last $days days: <b>$($sev['High'])</b> high, <b>$($sev['Medium'])</b> medium and <b>$($sev['Low'])</b> low severity.")
        $tr = $ruleGroups[0]; $pctR = [int](100 * $tr.Count / $total)
        $pol = ($alerts | Where-Object Rule -eq $tr.Name | Select-Object -First 1).Policy
        $ins.Add("Most alerts ($pctR%) come from the rule <b>$(ConvertTo-HtmlText $tr.Name)</b>$(if ($pol) { " in the policy <b>$(ConvertTo-HtmlText $pol)</b>" }).")
        $sitAll = @($alerts | ForEach-Object { $_.SensitiveTypes -split ', ' } | Where-Object { $_ })
        if ($sitAll.Count) { $ts = $sitAll | Group-Object | Sort-Object Count -Descending | Select-Object -First 1; $ins.Add("The information most often found was <b>$(ConvertTo-HtmlText $ts.Name)</b> (in $($ts.Count) alerts).") }
        $wlAll = @($alerts | ForEach-Object { $_.Workloads -split ', ' } | Where-Object { $_ })
        if ($wlAll.Count) { $tw = $wlAll | Group-Object | Sort-Object Count -Descending | Select-Object -First 1; $ins.Add("Most alerts happened in <b>$(ConvertTo-HtmlText $tw.Name)</b>.") }
        if ($extAlerts) { $ins.Add("<b>$extAlerts</b> alerts involved recipients outside the organisation$(if ($ext1 = (@($alerts | ForEach-Object { $_.ExternalDomains -split ', ' } | Where-Object { $_ } | Group-Object | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object Name) -join ', ')) { ", most often <b>$(ConvertTo-HtmlText $ext1)</b>" }).") }
        if ($distinctUsers -gt 0) {
            $topU = @($userList | Group-Object | Sort-Object Count -Descending | Select-Object -First 3)
            $share = [int](100 * (($topU | Measure-Object Count -Sum).Sum) / [Math]::Max(1, $userList.Count))
            if ($distinctUsers -ge 4 -and $share -ge 60) { $ins.Add("Three people account for <b>$share%</b> of alerts. That usually points to a business process (for example a team that routinely sends this data) rather than widespread risk: confirm with them and consider an exception or a safer way to share.") }
            else { $ins.Add("Alerts involve <b>$distinctUsers</b> different people, so the behaviour is spread across the organisation; user awareness and policy tips will help.") }
        }
        if ($open) { $ins.Add("<b>$open</b> alerts are still open (not resolved). Agree who triages DLP alerts and how quickly.") }
        $peak = @($alerts | Where-Object { ([string]$_.Raised).Length -ge 10 } | Group-Object { ([string]$_.Raised).Substring(0, 10) } | Sort-Object Count -Descending | Select-Object -First 1)
        if ($peak.Count -and $peak[0].Count -ge [Math]::Max(5, 3 * $total / [Math]::Max(1, $days))) { $ins.Add("Alerts peaked on <b>$($peak[0].Name)</b> ($($peak[0].Count) alerts); check what happened that day, such as a bulk upload or a mailing.") }
        $unlinked = @($alerts | Where-Object { $_.LinkedMatches -eq 0 }).Count
        if ($unlinked -eq $total) { $ins.Add("The rule match events behind the alerts could not be read, so what was found and where it was going are not shown. The account may need the <b>View-Only Audit Logs</b> role, or matches may be older than the search period.") }
    }

    # ---- datasets for the panels
    $byRule = @($ruleGroups | ForEach-Object { [pscustomobject]@{ Label = $_.Name; Value = $_.Count; Note = (($alerts | Where-Object Rule -eq $_.Name | Select-Object -First 1).Policy) } })
    $bySit = @($alerts | ForEach-Object { $_.SensitiveTypes -split ', ' } | Where-Object { $_ } | Group-Object | ForEach-Object { [pscustomobject]@{ Label = $_.Name; Value = $_.Count } })
    $sitTitle = 'What information was found (alerts)'
    if ($bySit.Count -eq 0 -and $matches2.Count) { $bySit = @($matches2 | ForEach-Object { $_.SensitiveTypes -split ', ' } | Where-Object { $_ } | Group-Object | ForEach-Object { [pscustomobject]@{ Label = $_.Name; Value = $_.Count } }); $sitTitle = 'What information was found (all rule matches)' }
    $byWl = @($alerts | ForEach-Object { $_.Workloads -split ', ' } | Where-Object { $_ } | Group-Object | ForEach-Object { [pscustomobject]@{ Label = $_.Name; Value = $_.Count } })
    $byUser = @($userList | Group-Object | ForEach-Object { [pscustomobject]@{ Label = (& $who $_.Name); Value = $_.Count } })
    $byDom = @($alerts | ForEach-Object { $_.ExternalDomains -split ', ' } | Where-Object { $_ } | Group-Object | ForEach-Object { [pscustomobject]@{ Label = $_.Name; Value = $_.Count } })
    $byStatus = @($alerts | Group-Object Status | ForEach-Object { [pscustomobject]@{ Label = $_.Name; Value = $_.Count } })
    $matchByRule = @($matches2 | Group-Object Rule | ForEach-Object {
            $rn = $_.Name; $ac = @($alerts | Where-Object Rule -eq $rn).Count
            [pscustomobject]@{ Label = $rn; Value = $_.Count; Note = $(if ($ac) { "$ac alerts" } else { 'no alerts' }) } })

    $css = @'
*{box-sizing:border-box}body{margin:0;font-family:"Segoe UI",Arial,sans-serif;background:#f3f5f8;color:#1f2a37}
header{background:#0f4c81;color:#fff;padding:22px 32px}header h1{margin:0;font-size:24px;font-weight:600}header p{margin:6px 0 0;opacity:.85;font-size:14px}
main{max-width:1180px;margin:0 auto;padding:22px 24px 40px}
.card{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:18px 20px;margin-bottom:18px}
.card h2{font-size:16px;margin:0 0 12px;color:#0f4c81}
.summary ul{margin:0;padding-left:20px;line-height:1.6}.summary li{margin-bottom:4px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(122px,1fr));gap:12px;margin-bottom:18px}
.kpi{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:14px 16px;border-top:4px solid #0f4c81}
.kpi b{display:block;font-size:28px}.kpi span{font-size:13px;color:#555}
.kpi.high{border-color:#c0392b}.kpi.med{border-color:#e08a00}.kpi.low{border-color:#2e86c1}.kpi.open{border-color:#7d3c98}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}@media(max-width:860px){.grid{grid-template-columns:1fr}}
.bar{display:grid;grid-template-columns:minmax(140px,52%) 1fr 56px;gap:10px;align-items:center;margin:6px 0;font-size:13px}
.lbl{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.note{display:block;color:#777;font-size:11px}
.trk{background:#eef1f5;border-radius:4px;height:14px}.fill{height:14px;border-radius:4px}.val{text-align:right;font-weight:600}
.chart{width:100%;height:auto}.ax{font-size:10px;fill:#666}.legend{font-size:12px;margin-top:6px}.legend span{margin-right:14px}.legend i{display:inline-block;width:10px;height:10px;margin-right:5px;border-radius:2px}
.empty{color:#777;font-style:italic}
table{width:100%;border-collapse:collapse;font-size:12.5px}th{background:#0f4c81;color:#fff;text-align:left;padding:7px;position:sticky;top:0}td{border-bottom:1px solid #e3e7ec;padding:6px 7px;vertical-align:top}
tr:nth-child(even) td{background:#f8fafc}.sev{padding:2px 8px;border-radius:10px;color:#fff;font-size:11px;font-weight:600}
.sev.High{background:#c0392b}.sev.Medium{background:#e08a00}.sev.Low{background:#2e86c1}.sev.Informational{background:#8a8a8a}
.tools{display:flex;gap:10px;margin-bottom:10px;flex-wrap:wrap}.tools input,.tools select{padding:7px 9px;border:1px solid #c9d1db;border-radius:6px;font-size:13px}.tools input{flex:1;min-width:220px}
.scroll{max-height:560px;overflow:auto}footer{font-size:12px;color:#666;max-width:1180px;margin:0 auto;padding:0 24px 30px}
'@
    $js = @'
function filt(){var q=document.getElementById('q').value.toLowerCase(),s=document.getElementById('s').value,n=0;
document.querySelectorAll('#t tbody tr').forEach(function(r){var ok=(!q||r.textContent.toLowerCase().indexOf(q)>=0)&&(!s||r.getAttribute('data-sev')===s);r.style.display=ok?'':'none';if(ok)n++;});
document.getElementById('n').textContent=n+' shown';}
'@
    $h = [System.Text.StringBuilder]::new()
    [void]$h.Append("<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>DLP alerts - $(ConvertTo-HtmlText $org)</title><style>$css</style></head><body>")
    [void]$h.Append("<header><h1>DLP alerts dashboard</h1><p>$(ConvertTo-HtmlText $org) &middot; last $days days ($($start.ToString('dd MMM yyyy')) to $($end.ToString('dd MMM yyyy'))) &middot; generated $($end.ToString('dd MMM yyyy HH:mm'))</p></header><main>")
    [void]$h.Append("<div class='card summary'><h2>What the alerts are telling you</h2><ul>" + (($ins | ForEach-Object { "<li>$_</li>" }) -join '') + '</ul></div>')
    [void]$h.Append("<div class='kpis'><div class='kpi'><b>$total</b><span>DLP alerts</span></div><div class='kpi high'><b>$($sev['High'])</b><span>High severity</span></div><div class='kpi med'><b>$($sev['Medium'])</b><span>Medium severity</span></div><div class='kpi low'><b>$($sev['Low'])</b><span>Low severity</span></div><div class='kpi open'><b>$open</b><span>Still open</span></div><div class='kpi'><b>$distinctUsers</b><span>People involved</span></div><div class='kpi'><b>$extAlerts</b><span>Alerts with external recipients</span></div><div class='kpi'><b>$('{0:N0}' -f $matches2.Count)</b><span>DLP rule matches (all)</span></div></div>")
    [void]$h.Append("<div class='card'><h2>Alerts per day</h2>$(New-HtmlDailyChart -Alerts $alerts -Start $start -End $end)</div>")
    [void]$h.Append("<div class='grid'><div class='card'><h2>Which rules raise the alerts</h2>$(New-HtmlBars $byRule '#0f4c81' 10 'No alerts')</div><div class='card'><h2>$sitTitle</h2>$(New-HtmlBars $bySit '#1e7b34' 10 'Not available: the rule match details could not be linked')</div></div>")
    [void]$h.Append("<div class='grid'><div class='card'><h2>Where it happened</h2>$(New-HtmlBars $byWl '#2e86c1' 8 'Not available')</div><div class='card'><h2>Who triggered the alerts</h2>$(New-HtmlBars $byUser '#7d3c98' 10 'Not available')</div></div>")
    [void]$h.Append("<div class='grid'><div class='card'><h2>Where it was going (external domains)</h2>$(New-HtmlBars $byDom '#c0392b' 10 'No external recipients found in the linked events')</div><div class='card'><h2>Alert status</h2>$(New-HtmlBars $byStatus '#5d6d7e' 6 'No alerts')</div></div>")
    [void]$h.Append("<div class='card'><h2>All DLP rule matches by rule (with and without alerts)</h2>$(New-HtmlBars $matchByRule '#8e9aaf' 15 'No rule matches were read (DLP rule activity not collected)')</div>")
    [void]$h.Append("<div class='card'><h2>Alerts</h2><div class='tools'><input id='q' placeholder='Search alerts, people, information types, files...' oninput='filt()'><select id='s' onchange='filt()'><option value=''>All severities</option><option>High</option><option>Medium</option><option>Low</option><option>Informational</option></select><span id='n' style='align-self:center;font-size:12px;color:#666'></span></div><div class='scroll'><table id='t'><thead><tr><th>Raised</th><th>Severity</th><th>Rule</th><th>Who</th><th>Where</th><th>What was found</th><th>Item</th><th>External</th><th>Status</th></tr></thead><tbody>")
    foreach ($a in ($alerts | Sort-Object Raised -Descending | Select-Object -First 2000)) {
        $people = (@($a.Users -split ', ' | Where-Object { $_ } | ForEach-Object { & $who $_ }) -join ', ')
        $item = if ($mask) { $(if ($a.Items) { '(hidden)' } else { '' }) } else { $a.Items }
        [void]$h.Append("<tr data-sev='$(ConvertTo-HtmlText $a.Severity)'><td>$(ConvertTo-HtmlText $a.Raised)</td><td><span class='sev $(ConvertTo-HtmlText $a.Severity)'>$(ConvertTo-HtmlText $a.Severity)</span></td><td>$(ConvertTo-HtmlText $a.Rule)</td><td>$(ConvertTo-HtmlText $people)</td><td>$(ConvertTo-HtmlText $a.Workloads)</td><td>$(ConvertTo-HtmlText $a.SensitiveTypes)</td><td>$(ConvertTo-HtmlText $item)</td><td>$(ConvertTo-HtmlText $a.ExternalDomains)</td><td>$(ConvertTo-HtmlText $a.Status)</td></tr>")
    }
    [void]$h.Append('</tbody></table></div></div></main>')
    [void]$h.Append("<footer>Source: the unified audit log (DLP alert records and DLP rule match records) read with Exchange Online PowerShell; no Microsoft Graph. 'What was found', 'Item' and 'External' come from the rule match events recorded up to an hour before each alert for the same rule and person. $(if ($mask) { 'People and items are masked (Settings).' } else { 'This page contains user names and file or email subjects: store and share it in line with the organisation''s data handling policy.' }) Detail: 66_DLP_Alert_Events, 67_DLP_Alerts, 64_DLP_Audit_Rule_Matches.</footer>")
    [void]$h.Append("<script>$js filt();</script></body></html>")
    $path = Join-Path $Script:OutputDir 'DLP_Alerts_Dashboard.html'
    [IO.File]::WriteAllText($path, $h.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    $Script:Cache['AlertDashboardPath'] = @($path)
    Write-Log "Alerts dashboard: $path" 'OK'
    return $path
}

function Export-AreaDlpAlerts {
    <# Expects Exchange Online to be connected. Rule-match events (Get-DlpAuditRecords) are read first so alerts can be linked to them. #>
    [void](Get-AcceptedDomainList)
    [void](Get-DlpAuditRecords)
    Export-Dataset '66_DLP_Alert_Events' @(Get-DlpAlertRecords) -AlreadyFlat
    Export-Dataset '67_DLP_Alerts' @(Get-DlpAlertSummary) -AlreadyFlat
    try { [void](New-DlpAlertDashboard) }
    catch {
        Write-Log "Alerts dashboard failed: $($_.Exception.Message)" 'ERROR'
        $Script:Errors.Add([pscustomobject]@{ Step = 'DLP alerts dashboard'; Outcome = 'Failed'; Error = $_.Exception.Message })
    }
}

function Invoke-AlertDashboard {
    <# Menu option 8 / -Mode AlertDashboard: DLP alerts and rule matches for the last N days, as a dashboard. #>
    if (-not (Connect-Tenant)) { return }
    Start-Run 'Alerts'
    Export-AreaTenant
    # Security & Compliance data first (rule and policy names), then Exchange Online for the audit log
    [void](Get-Rules); [void](Get-RuleClassifiers)
    try { Connect-Exo }
    catch {
        Write-Log "Exchange Online connection failed: $($_.Exception.Message)" 'ERROR'
        $Script:Errors.Add([pscustomobject]@{ Step = 'DLP alerts dashboard'; Outcome = 'Failed'; Error = $_.Exception.Message })
        Complete-Run; return
    }
    Export-Dataset '64_DLP_Audit_Rule_Matches' @(Get-DlpAuditRecords) -AlreadyFlat
    Export-Dataset '65_DLP_Rule_Activity_Summary' @(Get-DlpRuleActivity) -AlreadyFlat
    Export-AreaDlpAlerts
    Reset-AfterExo
    Complete-Run
    $p = @($Script:Cache['AlertDashboardPath'])
    if ($p.Count -and (Test-Path $p[0])) {
        Write-Host "Dashboard: $($p[0])" -ForegroundColor Cyan
        if ($Mode -eq 'Menu') { try { Invoke-Item $p[0] } catch { } }
    }
}

$Script:AreaFunctions = @{
    '1' = 'Export-AreaTenant'; '2' = 'Export-AreaPolicies'; '3' = 'Export-AreaRules'; '4' = 'Export-AreaClassifiers'
    '5' = 'Export-AreaSITs'; '6' = 'Export-AreaEDM'; '7' = 'Export-AreaLabels'; '8' = 'Export-AreaEndpoint'
    '9' = 'Export-AreaAlerts'; '10' = 'Export-AreaOwnership'; '11' = 'Export-AreaActivity'; '12' = 'Export-AreaExchange'
    '13' = 'Export-AreaContentExplorer'; '14' = 'Export-AreaDlpAudit'; '15' = 'Export-AreaLifecycle'; '16' = 'Export-AreaRiskCompliance'
    '17' = 'Export-AreaAuditDiscovery'
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
    ContentExplorer  = @{ Title = 'Learn about content explorer'; Url = 'https://learn.microsoft.com/en-us/purview/data-classification-content-explorer' }
    AuditSearch      = @{ Title = 'Search the audit log'; Url = 'https://learn.microsoft.com/en-us/purview/audit-search' }
    LabelEncryption  = @{ Title = 'Restrict access to content by using sensitivity labels to apply encryption'; Url = 'https://learn.microsoft.com/en-us/purview/encryption-sensitivity-labels' }
    LabelContainers  = @{ Title = 'Use sensitivity labels to protect Teams, Microsoft 365 groups and SharePoint sites'; Url = 'https://learn.microsoft.com/en-us/purview/sensitivity-labels-teams-groups-sites' }
    Retention        = @{ Title = 'Learn about retention policies and retention labels'; Url = 'https://learn.microsoft.com/en-us/purview/retention' }
    RecordsMgmt      = @{ Title = 'Learn about records management'; Url = 'https://learn.microsoft.com/en-us/purview/records-management' }
    AdaptiveScopes   = @{ Title = 'Adaptive or static policy scopes for retention'; Url = 'https://learn.microsoft.com/en-us/purview/purview-adaptive-scopes' }
    Disposition      = @{ Title = 'Disposition of content'; Url = 'https://learn.microsoft.com/en-us/purview/disposition' }
    InsiderRisk      = @{ Title = 'Learn about insider risk management'; Url = 'https://learn.microsoft.com/en-us/purview/insider-risk-management' }
    CommCompliance   = @{ Title = 'Learn about communication compliance'; Url = 'https://learn.microsoft.com/en-us/purview/communication-compliance' }
    InfoBarriers     = @{ Title = 'Learn about information barriers'; Url = 'https://learn.microsoft.com/en-us/purview/information-barriers' }
    AuditSolutions   = @{ Title = 'Learn about auditing solutions in Microsoft Purview'; Url = 'https://learn.microsoft.com/en-us/purview/audit-solutions-overview' }
    AuditRetention   = @{ Title = 'Manage audit log retention policies'; Url = 'https://learn.microsoft.com/en-us/purview/audit-log-retention-policies' }
    EDiscovery       = @{ Title = 'Learn about eDiscovery'; Url = 'https://learn.microsoft.com/en-us/purview/edisc' }
    Licensing        = @{ Title = 'Microsoft 365 guidance for security & compliance (licensing)'; Url = 'https://learn.microsoft.com/en-us/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/microsoft-365-security-compliance-licensing-guidance' }
}

$Script:DocWidth = 9864   # A4 portrait content width in twips (1.8 cm margins)

function New-WordDoc {
    $Script:Doc = @{
        Body  = [System.Text.StringBuilder]::new()
        Rels  = [System.Collections.Generic.List[string]]::new()
        RelId = 100
        Media = [System.Collections.Generic.List[object]]::new()
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

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'
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
    $document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + $W + '" xmlns:r="' + $R + '" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' + $Script:Doc.Body.ToString() + $sect + '</w:body></w:document>'

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
            foreach ($m in @($Script:Doc.Media)) {
                $es = $zip.CreateEntry("word/media/$($m.Name)").Open()
                try { $es.Write($m.Bytes, 0, $m.Bytes.Length) } finally { $es.Dispose() }
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
    Lifecycle   = @(
        @{ Fw = 'NCSC'; Ref = 'CAF B3.c Stored data; A1 Governance'; Expects = 'Stored data is kept only as long as needed, and important records are protected from change.' }
        @{ Fw = 'NISTCSF'; Ref = 'ID.AM-08, GV.OC-03'; Expects = 'Data is managed through its life cycle in line with legal and regulatory requirements.' }
        @{ Fw = 'NIST80053'; Ref = 'SI-12, MP-6'; Expects = 'Information is retained as required and disposed of when no longer needed.' }
        @{ Fw = 'ISO27001'; Ref = '5.33, 8.10, 5.31'; Expects = 'Records are protected, information is deleted when no longer required, and legal requirements are met.' }
    )
    RiskCompliance = @(
        @{ Fw = 'NCSC'; Ref = 'CAF C1 Security monitoring; C2 Proactive security event discovery'; Expects = 'Risky behaviour that could lead to data loss is detected through monitoring and analysis.' }
        @{ Fw = 'NISTCSF'; Ref = 'DE.CM-03, DE.AE-02'; Expects = 'Personnel activity is monitored and potentially adverse events are analysed.' }
        @{ Fw = 'NIST80053'; Ref = 'PM-12, SI-4, AC-4'; Expects = 'An insider threat programme exists and information flows between groups are controlled.' }
        @{ Fw = 'ISO27001'; Ref = '8.16, 5.3'; Expects = 'Activity is monitored for anomalies and conflicting duties are segregated.' }
    )
    AuditDiscovery = @(
        @{ Fw = 'NCSC'; Ref = 'CAF C1.a Monitoring coverage; C1.b Securing logs'; Expects = 'Logs are collected and kept long enough to support investigation.' }
        @{ Fw = 'NISTCSF'; Ref = 'PR.PS-04, RS.AN-03'; Expects = 'Log records are generated and available, and incidents can be analysed to establish what happened.' }
        @{ Fw = 'NIST80053'; Ref = 'AU-11, AU-12'; Expects = 'Audit records are generated and retained to support after-the-fact investigations.' }
        @{ Fw = 'ISO27001'; Ref = '8.15, 5.28'; Expects = 'Activity is logged and evidence can be identified, collected and preserved.' }
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
    Lifecycle = 'Retention and records'; RiskCompliance = 'Insider risk and communication compliance'; AuditDiscovery = 'Audit and eDiscovery'
}
$Script:ReportSkipAreas = @()
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
            if (@($Script:ReportSkipAreas) -contains $a) { continue }
            $st = if ($Script:AreaStatus.ContainsKey($a)) { $Script:AreaStatus[$a] } else { 'N' }
            $cells = @($Script:AreaNames[$a])
            foreach ($f in $fws) { $cells += ((@($Script:FrameworkMap[$a] | Where-Object { $_.Fw -eq $f } | ForEach-Object { $_.Ref })) -join '; ') }
            $cells += "[[$st]]$($Script:StatusText[$st])"
            , $cells
        })
    Add-DocTable -Headers $headers -Widths $widths -FontSize 15 -Rows $rows
}

function Add-ImplementationGuidance {
    <#
        Writes the "Applying the configuration" section of the Word report (section 21 to 24, depending on scope). The Set-/New- commands in the @' '@ blocks below are EXAMPLE TEXT
        for the customer: they are passed to Add-DocCode as strings and printed into the report. Nothing in this
        function runs a command against the tenant.
    #>
    param($Ctx, [int]$Section = 21)
    Add-DocHeading "$Section. Applying the configuration" 1
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
        Add-Guidance -Title "$Section.$n Create or extend DLP policies to all core locations" -Area 'Coverage' `
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
        Add-Guidance -Title "$Section.$n Move policies from simulation to enforcement" -Area 'Simulation' `
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
        Add-Guidance -Title "$Section.$n Add policy tips and proportionate actions to rules" -Area 'UX' `
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
        Add-Guidance -Title "$Section.$n Define and publish sensitivity labels" -Area 'Labels' `
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
        Add-Guidance -Title "$Section.$n Use sensitivity labels as DLP conditions" -Area 'Labels' `
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
        Add-Guidance -Title "$Section.$n Onboard devices and restrict USB, printing and uploads" -Area 'Endpoint' `
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
        Add-Guidance -Title "$Section.$n Protect Microsoft 365 Copilot interactions" -Area 'Copilot' `
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
        Add-Guidance -Title "$Section.$n Configure alerts and a review process" -Area 'Alerts' `
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
        Add-Guidance -Title "$Section.$n Tune confidence levels and instance counts" -Area 'Confidence' `
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
        Add-Guidance -Title "$Section.$n Create custom sensitive information types for organisation-specific data" -Area 'CustomSIT' `
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
        Add-Guidance -Title "$Section.$n Consider Exact Data Match for key structured datasets" -Area 'EDM' `
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
        Add-Guidance -Title "$Section.$n Turn on the unified audit log" -Area 'Exchange' `
            -When 'The unified audit log is disabled. DLP reporting and Activity Explorer depend on it.' `
            -Portal @('Go to Microsoft Purview portal > Solutions > Audit.', 'If prompted, select Start recording user and admin activity.', 'Allow up to 60 minutes for recording to begin.') `
            -PowerShell @'
# Exchange Online PowerShell (Connect-ExchangeOnline)
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
'@ -Validate @('Get-AdminAuditLogConfig shows UnifiedAuditLogIngestionEnabled True.') -LinkKeys @('ActivityExplorer')
    }
    if ($Ctx.LifecycleOn -and @($Ctx.RetentionGaps).Count) {
        $n++
        Add-Guidance -Title "$Section.$n Apply retention policies to the core locations" -Area 'Lifecycle' `
            -When "No enabled retention policy covers: $(@($Ctx.RetentionGaps) -join ', ')." `
            -Portal @(
            'Agree the retention periods with legal, records and business owners, for example in a short retention schedule. The periods below are examples only.'
            'Go to the Microsoft Purview portal > Solutions > Data Lifecycle Management > Policies > Retention policies and select New retention policy.'
            'Choose a static scope (all mailboxes, sites or accounts) or an adaptive scope, and select the locations. Teams chats and channel messages need a policy of their own.'
            'Choose whether to keep content, delete it, or keep it and then delete it, and the period.'
            'Review and submit. The policy can take up to seven days to apply everywhere.'
        ) -PowerShell @'
# Email, SharePoint and OneDrive: keep for 7 years from creation, then delete (example period)
New-RetentionCompliancePolicy -Name "Org retention - 7 years" -ExchangeLocation All -SharePointLocation All -OneDriveLocation All
New-RetentionComplianceRule -Name "Org retention - 7 years rule" -Policy "Org retention - 7 years" `
    -RetentionDuration 2555 -RetentionComplianceAction KeepAndDelete -ExpirationDateOption CreationAgeInDays

# Teams chats and channel messages need their own policy (example period)
New-RetentionCompliancePolicy -Name "Teams retention - 2 years" -TeamsChatLocation All -TeamsChannelLocation All
New-RetentionComplianceRule -Name "Teams retention - 2 years rule" -Policy "Teams retention - 2 years" `
    -RetentionDuration 730 -RetentionComplianceAction KeepAndDelete
'@ -Validate @('Get-RetentionCompliancePolicy | Format-Table Name, Enabled, DistributionStatus shows Success for the new policies.', 'Re-run this assessment: the retention coverage table shows the locations as covered.') -LinkKeys @('Retention', 'AdaptiveScopes')
    }
    if ($Ctx.AuditRetNoDlp) {
        $n++
        Add-Guidance -Title "$Section.$n Keep DLP audit records for longer" -Area 'AuditDiscovery' `
            -When 'No audit log retention policy keeps DLP activity beyond the default retention.' `
            -Portal @(
            'Confirm the organisation has Audit (Premium) (Microsoft 365 E5, E5 eDiscovery & Audit, or Microsoft Purview Suite).'
            'Go to the Microsoft Purview portal > Solutions > Audit > Policies and select Create audit retention policy.'
            'Name the policy, choose the DLP record types (ComplianceDLPExchange, ComplianceDLPSharePoint and DLPEndpoint), choose a duration and a priority, and save.'
        ) -PowerShell @'
New-UnifiedAuditLogRetentionPolicy -Name "DLP activity - 1 year" `
    -RecordTypes ComplianceDLPExchange, ComplianceDLPSharePoint, DLPEndpoint `
    -RetentionDuration TwelveMonths -Priority 100
'@ -Validate @('Get-UnifiedAuditLogRetentionPolicy lists the policy with the DLP record types.') -LinkKeys @('AuditRetention', 'AuditSolutions')
    }
    if ($Ctx.IrmNone) {
        $n++
        Add-Guidance -Title "$Section.$n Start insider risk management with a Data leaks policy" -Area 'RiskCompliance' `
            -When 'Insider risk management is available but no policies exist.' `
            -Portal @(
            'Agree the purpose, who can see results and how cases are handled with HR, legal and privacy teams. Keep the default pseudonymisation of user names.'
            'Go to the Microsoft Purview portal > Solutions > Insider Risk Management > Settings and review the policy indicators and analytics settings.'
            'Go to Policies > Create policy and choose the Data leaks template.'
            'Choose the users in scope and, as the triggering event, select a DLP policy whose rules raise high-severity alerts.'
            'Choose the indicators to score, review and submit. Check alerts weekly during the first month and adjust thresholds.'
        ) -PowerShell '' -Validate @('The policy shows as Healthy on the Insider Risk Management policies page.', 'High-severity DLP alerts for in-scope users start the policy within 24 hours.') -LinkKeys @('InsiderRisk')
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

# ---------- Readable configuration detail (plain-English descriptions of collected settings) ----------

function ConvertFrom-CamelCase {
    param([string]$Text)
    if (-not $Text) { return '' }
    $t = $Text -creplace '([a-z0-9])([A-Z])', '$1 $2' -creplace '([A-Z]+)([A-Z][a-z])', '$1 $2'
    return ($t.Substring(0, 1).ToUpper() + $t.Substring(1).ToLower()) -replace '\bdlp\b', 'DLP' -replace '\bid\b', 'ID' -replace '\bip\b', 'IP' -replace '\burl\b', 'URL'
}

function Get-ItemText {
    <# Turns a location / recipient / group value into readable text. #>
    param($Value)
    $out = foreach ($v in @($Value)) {
        if ($null -eq $v) { continue }
        if ($v -is [string] -or $v -is [ValueType]) { [string]$v; continue }
        $n = Get-Val $v 'DisplayName'; if (-not $n) { $n = Get-Val $v 'Name' }; if (-not $n) { $n = [string]$v }
        $n
    }
    $list = @($out | Where-Object { $_ -and $_ -ne '{}' })
    if ($list.Count -eq 0) { return '' }
    if ($list.Count -gt 15) { return (($list | Select-Object -First 15) -join ', ') + " and $($list.Count - 15) more" }
    return ($list -join ', ')
}

$Script:LocationNames = [ordered]@{
    ExchangeLocation = 'Exchange email'; SharePointLocation = 'SharePoint sites'; OneDriveLocation = 'OneDrive accounts'; TeamsLocation = 'Teams chat and channel messages'
    EndpointDlpLocation = 'Devices'; OnPremisesScannerDlpLocation = 'On-premises repositories'; ThirdPartyAppDlpLocation = 'Non-Microsoft cloud apps'; PowerBIDlpLocation = 'Power BI workspaces'
    ExchangeSenderMemberOf = 'Exchange: senders in groups'; OneDriveSharedByMemberOf = 'OneDrive: accounts of group members'; TeamsLocationException = 'Teams: excluded'
    ModernGroupLocation = 'Microsoft 365 groups'
}

function Get-PolicyLocationRows {
    param($Policy, [string]$Workload)
    $inc = @(); $exc = @()
    foreach ($p in $Policy.PSObject.Properties) {
        if ($p.Name -notmatch '(Location|MemberOf|SharedBy)(Exception)?$' -or $p.Name -match '^Locations$') { continue }
        $txt = Get-ItemText $p.Value
        if (-not $txt) { continue }
        $base = $p.Name -replace 'Exception$', ''
        $label = if ($Script:LocationNames.Contains($base)) { $Script:LocationNames[$base] } else { ConvertFrom-CamelCase $base }
        if ($p.Name -match 'Exception$') { $exc += "$label - $txt" } else { $inc += $(if ($txt -eq 'All') { "$label (all)" } else { "$label - $txt" }) }
    }
    if ($Workload -match 'Applications|Copilot') { $inc += 'Microsoft 365 Copilot and Copilot Chat' }
    return @{ Included = $inc; Excluded = $exc }
}

$Script:EndpointActionNames = @{
    RemovableMedia = 'Copy to removable USB storage'; Print = 'Print'; CloudEgress = 'Upload to cloud service domains'; CopyPaste = 'Copy to clipboard'
    NetworkShare = 'Copy to a network share'; UnallowedApps = 'Access by restricted apps'; UnallowedBluetoothApp = 'Transfer over Bluetooth'
    UnallowedBrowser = 'Access by unallowed browsers'; RemoteDesktopServices = 'Copy over Remote Desktop'; RdpCopyPaste = 'Copy over Remote Desktop'
    PasteToBrowser = 'Paste to browser'; CopyToRemovableMedia = 'Copy to removable USB storage'
}

function Get-ClassifierText {
    param($Row, [hashtable]$LabelNames)
    if ($Row.ReferenceType -eq 'SensitivityLabel') {
        $n = [string]$Row.Name; if ($LabelNames.ContainsKey($n)) { $n = $LabelNames[$n] } elseif ($Row.Id -and $LabelNames.ContainsKey([string]$Row.Id)) { $n = $LabelNames[[string]$Row.Id] }
        return "content labelled **$n**"
    }
    $bits = @()
    if ($Row.ClassifierType -match 'MLModel|Trainable') { $bits += 'trainable classifier' }
    elseif ($Row.ClassifierType -match 'ExactMatch|EDM') { $bits += 'Exact Data Match' }
    if ($Row.ConfidenceLevel) { $bits += "$(([string]$Row.ConfidenceLevel).ToLower()) confidence" }
    elseif ($Row.MinConfidence) { $bits += "confidence $($Row.MinConfidence)%+" }
    $min = [string]$Row.MinCount; $max = [string]$Row.MaxCount
    if ($min -and $min -ne '1') { $bits += $(if ($max -and $max -ne '-1') { "$min to $max instances" } else { "$min or more" }) }
    elseif ($min -eq '1' -and $max -and $max -ne '-1') { $bits += "up to $max instances" }
    $t = "**$($Row.Name)**"
    if ($bits.Count) { $t += " ($($bits -join ', '))" }
    return $t
}

$Script:RuleMetaProps = '^(Name|Policy|ParentPolicyName|Guid|Identity|Id|ImmutableId|Priority|Mode|Disabled|Comment|Workload|Created.*|LastModified.*|When.*|DistributionStatus|ReadOnly|IsValid|ObjectVersion|ExchangeVersion|ObjectState|OrganizationId|RunspaceId|PSComputerName|PSShowComputerName|ContentContainsSensitiveInformation|ExceptIfContentContainsSensitiveInformation|AdvancedRule|RuleErrorAction|ReportSeverityLevel)$'
$Script:RuleCondProps = '^(ContentContains|ExceptIf|Sent|From|Recipient|Document|Content|Subject|Message|Access|Has|Header|Anyof|Processing|With|Unscannable|Evaluate|Sender)'

function Get-RuleDescription {
    <# Returns @{ When = string[]; Except = string[]; Then = string[] } describing a DLP rule in plain English. #>
    param($Rule, [object[]]$ClassRows, [hashtable]$LabelNames)
    $when = @(); $except = @(); $then = @()
    $cond = @($ClassRows | Where-Object { $_.Source -ne 'Exception' })
    $exc = @($ClassRows | Where-Object { $_.Source -eq 'Exception' })
    if ($cond.Count) { $when += 'Content contains any of: ' + ((@($cond | ForEach-Object { Get-ClassifierText $_ $LabelNames } | Select-Object -Unique)) -join '; ') }
    if ($exc.Count) { $except += 'Content contains any of: ' + ((@($exc | ForEach-Object { Get-ClassifierText $_ $LabelNames } | Select-Object -Unique)) -join '; ') }
    foreach ($p in $Rule.PSObject.Properties) {
        if ($p.Name -match $Script:RuleMetaProps -or $p.Name -notmatch $Script:RuleCondProps -or ($Script:RuleActionProps -contains $p.Name -and $p.Name -ne 'AccessScope')) { continue }
        $v = ConvertTo-CellValue $p.Value
        if (-not $v -or $v -eq 'False' -or $v -eq '0' -or $v -eq '{}' -or $v -eq '[]') { continue }
        if ($p.Name -eq 'AccessScope') {
            $when += $(switch ($v) { 'NotInOrganization' { 'Content is shared with people **outside the organisation**' } 'InOrganization' { 'Content is shared **inside the organisation**' } default { "Access scope: $v" } })
            continue
        }
        $isExc = $p.Name -like 'ExceptIf*'
        $label = ConvertFrom-CamelCase ($p.Name -replace '^ExceptIf', '')
        $text = if ($v -eq 'True') { $label } else { "$($label): $(($v -replace ' \| ', ', '))" }
        if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + ' ...' }
        if ($isExc) { $except += $text } else { $when += $text }
    }
    if ($when.Count -eq 0) { $when += 'Any content in the locations the policy covers' }

    $block = ConvertTo-CellValue (Get-Val $Rule 'BlockAccess')
    if ($block -eq 'True') {
        $scope = [string](Get-Val $Rule 'BlockAccessScope')
        $then += $(switch ($scope) { 'PerUser' { '**Block** access for people outside the organisation' } 'PerAnonymousUser' { '**Block** access through "anyone" links' } default { '**Block** the action (or access for everyone except the owner)' } })
    }
    $ovr = ConvertTo-CellValue (Get-Val $Rule 'NotifyAllowOverride')
    if ($ovr) {
        $ovText = @(($ovr -split ' \| ') | ForEach-Object { switch ($_) { 'WithJustification' { 'with a business justification' } 'WithoutJustification' { 'without a justification' } 'FalsePositive' { 'by reporting a false positive' } 'WithAcknowledgement' { 'after acknowledging a warning' } default { $_ } } })
        $then += "Users can **override** $($ovText -join ' or ')"
    }
    $notify = ConvertTo-CellValue (Get-Val $Rule 'NotifyUser')
    if ($notify) {
        $then += "**Notify** $(Get-PersonText ($notify -split ' \| ')) with a policy tip and/or email"
    }
    $tip = [string](Get-Val $Rule 'NotifyPolicyTipCustomText')
    if ($tip) { $then += "Policy tip text: ""$tip""" }
    foreach ($er in @(Get-Val $Rule 'EndpointDlpRestrictions')) {
        if ($null -eq $er) { continue }
        $set = [string](Get-Val $er 'Setting'); $val = [string](Get-Val $er 'Value')
        if (-not $set) { continue }
        $nm = if ($Script:EndpointActionNames.ContainsKey($set)) { $Script:EndpointActionNames[$set] } else { ConvertFrom-CamelCase $set }
        $act = switch ($val) { 'Block' { 'block' } 'Warn' { 'warn (user can override)' } 'Audit' { 'audit only' } default { $val.ToLower() } }
        $then += "Devices - $($nm): **$act**"
    }
    $enc = [string](Get-Val $Rule 'EncryptRMSTemplate'); if ($enc) { $then += "Encrypt the message with template ""$enc""" }
    if (ConvertTo-CellValue (Get-Val $Rule 'Quarantine')) { $then += 'Quarantine the message' }
    $redir = Get-ItemText (Get-Val $Rule 'RedirectMessageTo'); if ($redir) { $then += "Redirect the message to $redir" }
    $subj = [string](Get-Val $Rule 'ModifySubject'); if ($subj) { $then += 'Modify the message subject' }
    $alert = ConvertTo-CellValue (Get-Val $Rule 'GenerateAlert')
    if ($alert) {
        $alertTo = @(($alert -split ' \| ') | Where-Object { $_ -and $_ -notmatch '^(true|false)$' })
        $then += $(if ($alertTo.Count) { "Raise an **alert** and email $(Get-PersonText $alertTo)" } else { 'Raise an **alert** (see section 16 for who is notified)' })
    }
    $inc = ConvertTo-CellValue (Get-Val $Rule 'GenerateIncidentReport'); if ($inc) { $then += "Email an incident report to $(Get-PersonText ($inc -split ' \| '))" }
    if ((ConvertTo-CellValue (Get-Val $Rule 'StopPolicyProcessing')) -eq 'True') { $then += 'Stop processing further rules' }
    if (-not ($then | Where-Object { $_ -notmatch '^(Raise an|Email an incident|Policy tip text|Stop processing)' })) { $then = @('**Audit only**: record the event (users are not told and the action is allowed)') + $then }
    return @{ When = $when; Except = $except; Then = $then }
}

function Add-PolicyDetailAppendix {
    param([object[]]$Policies, [object[]]$Rules, [object[]]$ClassRows, [hashtable]$LabelNames)
    Add-DocHeading 'Appendix E: DLP policies and rules in detail' 1
    Add-DocPara 'This appendix describes every DLP policy and rule collected from the tenant in plain English, so the configuration can be reviewed without opening the spreadsheets. Policies are listed in priority order; rules are listed in the order they are evaluated. Where a rule groups conditions with AND / OR logic, the exact logic is in 14_DLP_Rules_Conditions_Actions.'
    if ($Policies.Count -eq 0) { Add-DocPara 'No DLP policies were found.'; return }
    $modeText = @{ 'Enable' = 'Enforced (actions are applied)'; 'TestWithNotifications' = 'Simulation, with policy tips shown to users'; 'TestWithoutNotifications' = 'Simulation, silent (nothing shown to users)'; 'Disable' = 'Disabled'; 'PendingDeletion' = 'Being deleted' }
    foreach ($p in ($Policies | Sort-Object { [int]('0' + (Get-Val $_ 'Priority')) })) {
        Add-DocHeading ([string]$p.Name) 2
        $loc = Get-PolicyLocationRows -Policy $p -Workload (ConvertTo-CellValue (Get-Val $p 'Workload'))
        $pRules = @($Rules | Where-Object { [string](Get-Val $_ 'ParentPolicyName') -eq [string]$p.Name } | Sort-Object { [int]('0' + (Get-Val $_ 'Priority')) })
        $changed = ''
        try { if (Get-Val $p 'WhenChangedUTC') { $changed = ([datetime](Get-Val $p 'WhenChangedUTC')).ToString('dd MMM yyyy') } } catch { }
        $by = [string](Get-Val $p 'LastModifiedBy')
        $mode = [string]$p.Mode
        $card = @(
            , @('Mode', $(if ($modeText.ContainsKey($mode)) { $modeText[$mode] } else { $mode }))
            , @('Applies to', $(if ($loc.Included.Count) { $loc.Included -join '; ' } else { 'No locations found' }))
        )
        if ($loc.Excluded.Count) { $card += , @('Excludes', ($loc.Excluded -join '; ')) }
        $desc = [string](Get-Val $p 'Comment'); if ($desc) { $card += , @('Description', $desc) }
        $card += , @('Rules', [string]$pRules.Count)
        $card += , @('Priority', [string](Get-Val $p 'Priority'))
        if ($changed -or $by) { $card += , @('Last changed', ((@($changed, $(if ($by) { "by $by" })) | Where-Object { $_ }) -join ' ')) }
        Add-DocTable -Headers @('Policy setting', 'Value') -Widths @(1.6, 6) -FontSize 17 -Rows $card
        if ($pRules.Count -eq 0) { Add-DocPara 'This policy has no rules.' -Italic; continue }
        $n = 1
        foreach ($r in $pRules) {
            $state = if ((ConvertTo-CellValue (Get-Val $r 'Disabled')) -eq 'True') { 'Disabled' } else { 'Active' }
            $sev = [string](Get-Val $r 'ReportSeverityLevel')
            $meta = @($state, $(if ($sev -and $sev -ne 'None') { "severity $($sev.ToLower())" })) | Where-Object { $_ }
            Add-DocPara "**Rule $($n): $($r.Name)**  ($($meta -join ', '))" -Size 21 -Color '1F3B57' -After 40 -Before 120
            $d = Get-RuleDescription -Rule $r -ClassRows @($ClassRows | Where-Object { $_.Policy -eq [string]$p.Name -and $_.Rule -eq [string]$r.Name }) -LabelNames $LabelNames
            Add-DocBullet ('**When:** ' + ($d.When -join '; **and** '))
            if ($d.Except.Count) { Add-DocBullet ('**Except when:** ' + ($d.Except -join '; **or** ')) }
            Add-DocBullet ('**Then:** ' + ($d.Then -join '; '))
            $n++
        }
    }
}

$Script:EndpointSettingNames = @{
    CloudAppMode = 'Service domain mode (allow or block list)'; CloudAppRestrictions = 'Service domains'; UnallowedBrowser = 'Unallowed browsers'
    UnallowedApp = 'Restricted apps'; UnallowedCloudSyncApp = 'Unallowed cloud sync apps'; UnallowedBluetoothApp = 'Unallowed Bluetooth apps'
    PathExclusion = 'Excluded file paths (Windows)'; MacPathExclusion = 'Excluded file paths (macOS)'; AdvancedClassificationEnabled = 'Advanced classification scanning'
    BandwidthLimitEnabled = 'Bandwidth limit for advanced classification'; DailyBandwidthLimitInMB = 'Daily bandwidth limit (MB)'; EvidenceStoreSettings = 'Evidence collection for file activities'
    IncludePredefinedUnallowedBluetoothApps = 'Include predefined unallowed Bluetooth apps'; AlwaysAuditEndpointActivity = 'Always audit device activity'
    VPNSettings = 'VPN settings'; NetworkPrinter = 'Network printers'; BusinessJustificationList = 'Override justification options'
    DefaultPathExclusionsEnabled = 'Default file path exclusions in use'
}
$Script:EndpointAreaNames = @{
    EndpointDlpGlobalSettings = 'Global settings'; EndpointDlpBrowserRestrictions = 'Browser restrictions'; DlpPrinterGroups = 'Printer groups'
    DlpRemovableMediaGroups = 'Removable storage groups'; DlpNetworkShareGroups = 'Network share groups'; DlpAppGroups = 'App groups'
    DlpExtensionGroups = 'File extension groups'; ExtendTeamsDlpToSpoOdbConsent = 'Teams DLP extended to SharePoint and OneDrive files'
}

function Get-EndpointSettingRows {
    param([object[]]$Settings)
    foreach ($s in $Settings) {
        if (-not $s.Value -or ([string]$s.Value).Trim() -in '{}', '[]') { continue }
        $areaKey = ([string]$s.Area) -replace 'Psws$', ''
        $area = if ($Script:EndpointAreaNames.ContainsKey($areaKey)) { $Script:EndpointAreaNames[$areaKey] } else { ConvertFrom-CamelCase $areaKey }
        $name = [string]$s.Setting; $val = [string]$s.Value
        if (-not $name -and $val.Trim().StartsWith('{')) {
            try {
                $o = $val | ConvertFrom-Json
                $g = Get-Val $o 'groupName'; if (-not $g) { $g = Get-Val $o 'name' }
                $name = if ($g) { [string]$g } else { 'Group' }
                $parts = foreach ($pp in $o.PSObject.Properties) { if ($pp.Name -notmatch '^(groupName|name)$') { "$(ConvertFrom-CamelCase $pp.Name): $(ConvertTo-CellValue $pp.Value)" } }
                $val = (@($parts) -join '; ')
            }
            catch { }
        }
        elseif ($Script:EndpointSettingNames.ContainsKey($name)) { $name = $Script:EndpointSettingNames[$name] }
        elseif ($name) { $name = ConvertFrom-CamelCase $name }
        if ($val.Length -gt 400) { $val = $val.Substring(0, 400) + ' ...' }
        , @($area, $name, ($val -replace ' \| ', ', '))
    }
}

function Get-LabelPolicyRows {
    param([object[]]$LabelPolicies, [hashtable]$LabelNames)
    foreach ($lp in $LabelPolicies) {
        $labs = @(foreach ($l in @(Get-Val $lp 'Labels')) { $k = [string](Get-ItemText $l); if ($LabelNames.ContainsKey($k)) { $LabelNames[$k] } else { $k } }) | Where-Object { $_ }
        $scope = @(foreach ($p in $lp.PSObject.Properties) {
                if ($p.Name -notmatch 'Location$') { continue }
                $t = Get-ItemText $p.Value
                if ($t) {
                    $where = switch ($p.Name -replace 'Location$', '') { 'Exchange' { 'Users and groups' } 'ModernGroup' { 'Microsoft 365 groups' } default { ConvertFrom-CamelCase $_ } }
                    "$($where): $(if ($t -eq 'All') { 'all' } else { $t })"
                }
            })
        $settings = @(foreach ($s in @(Get-Val $lp 'Settings')) {
                $m = [regex]::Match([string]$s, '^\[\s*([^,]+?)\s*,\s*(.*?)\s*\]$')
                if (-not $m.Success) { continue }
                $k = $m.Groups[1].Value.ToLower(); $v = $m.Groups[2].Value
                switch ($k) {
                    'mandatory' { if ($v -eq 'true') { 'Labelling is mandatory' } }
                    'requiredowngradejustification' { if ($v -eq 'true') { 'Justification required to remove or lower a label' } }
                    'defaultlabelid' { "Default label: $(if ($LabelNames.ContainsKey($v)) { $LabelNames[$v] } else { $v })" }
                    'outlookdefaultlabel' { if ($v -and $v -ne 'None') { "Default email label: $(if ($LabelNames.ContainsKey($v)) { $LabelNames[$v] } else { $v })" } }
                    'disablemandatoryinoutlook' { if ($v -eq 'true') { 'Mandatory labelling not applied in Outlook' } }
                    'powerbimandatory' { if ($v -eq 'true') { 'Mandatory labelling for Power BI' } }
                    'siteandgroupmandatory' { if ($v -eq 'true') { 'Label required for sites and groups' } }
                    'teamworkmandatory' { if ($v -eq 'true') { 'Label required for Teams meetings' } }
                    'siteandgroupdefaultlabelid' { if ($v -and $v -ne 'None') { "Default site and group label: $(if ($LabelNames.ContainsKey($v)) { $LabelNames[$v] } else { $v })" } }
                }
            })
        , @([string]$lp.Name, (@($labs) -join ', '), $(if ($scope.Count) { $scope -join '; ' } else { 'All users' }), $(if ($settings.Count) { $settings -join '; ' } else { 'Default settings' }))
    }
}

function Get-PersonText {
    <# Turns recipient keywords (Owner, SiteAdmin, TenantAdmins...) into plain English. #>
    param([string[]]$Values)
    $list = @(foreach ($v in $Values) {
            if (-not $v) { continue }
            switch ($v) { 'Owner' { 'the content owner' } 'LastModifier' { 'the person who last modified it' } 'SiteAdmin' { 'the site admin' } 'TenantAdmins' { 'tenant admins' } default { $v } }
        })
    if ($list.Count -le 1) { return ($list -join '') }
    return (($list[0..($list.Count - 2)] -join ', ') + ' and ' + $list[-1])
}

function Get-AutoLabelRows {
    param([object[]]$AutoPolicies, [object[]]$AutoRules, [hashtable]$LabelNames)
    $modeText = @{ 'Enable' = 'Applying labels'; 'TestWithNotifications' = 'Simulation (with notifications)'; 'TestWithoutNotifications' = 'Simulation'; 'Disable' = 'Off' }
    foreach ($ap in $AutoPolicies) {
        $lab = [string](Get-ItemText (Get-Val $ap 'ApplySensitivityLabel')); if ($LabelNames.ContainsKey($lab)) { $lab = $LabelNames[$lab] }
        $mode = [string]$ap.Mode; if ($modeText.ContainsKey($mode)) { $mode = $modeText[$mode] }
        $loc = Get-PolicyLocationRows -Policy $ap -Workload ''
        $rows = @($AutoRules | Where-Object { [string](Get-Val $_ 'ParentPolicyName') -eq [string]$ap.Name })
        $sits = @(foreach ($r in $rows) { Find-RuleClassifiers (Get-Val $r 'ContentContainsSensitiveInformation') ([string]$ap.Name) ([string]$r.Name) 'Condition' }) |
            ForEach-Object { if ($_.ReferenceType -eq 'SensitivityLabel') { "label $($_.Name)" } else { [string]$_.Name } } | Where-Object { $_ } | Select-Object -Unique
        , @([string]$ap.Name, $lab, $mode, $(if ($loc.Included.Count) { $loc.Included -join '; ' } else { '' }), $(if (@($sits).Count) { @($sits) -join ', ' } else { "$($rows.Count) rule(s); see 33_Auto_Labelling_Rules" }))
    }
}

function Get-DictionaryRows {
    param([object[]]$Dictionaries)
    foreach ($d in $Dictionaries) {
        $kw = [string](Get-Val $d 'KeywordDictionary')
        $count = if ($kw) { @($kw -split '[,\r\n]+' | Where-Object { $_.Trim() }).Count } else { 0 }
        , @([string]$d.Name, [string](Get-Val $d 'Description'), $(if ($count) { [string]$count } else { 'Not shown' }))
    }
}

function Get-EdmSchemaRows {
    param([object[]]$Schemas)
    foreach ($s in $Schemas) {
        $xmlText = [string](Get-Val $s 'EdmSchemaXML'); if (-not $xmlText) { $xmlText = [string](Get-Val $s 'SchemaXml') }
        $fields = @([regex]::Matches($xmlText, '<Field\b[^>]*?name="([^"]+)"[^>]*?>') | ForEach-Object {
                $n = $_.Groups[1].Value; if ($_.Value -match 'searchable="true"') { "**$n** (searchable)" } else { $n } })
        $name = [string](Get-Val $s 'Name'); if (-not $name) { $m = [regex]::Match($xmlText, '<DataStore\b[^>]*?name="([^"]+)"'); if ($m.Success) { $name = $m.Groups[1].Value } }
        , @($name, $(if ($fields.Count) { $fields -join ', ' } else { 'Field list not available' }))
    }
}

function Get-AlertPolicyRows {
    param([object[]]$AlertPolicies)
    foreach ($a in $AlertPolicies) {
        $state = if ((ConvertTo-CellValue (Get-Val $a 'Disabled')) -eq 'True') { 'Off' } else { 'On' }
        $who = Get-ItemText (Get-Val $a 'NotifyUser'); if (-not $who) { $who = 'No email recipients' }
        $agg = [string](Get-Val $a 'AggregationType'); $thr = [string](Get-Val $a 'Threshold'); $win = [string](Get-Val $a 'TimeWindow')
        $trigger = switch ($agg) { 'None' { 'Every matching event' } 'SimpleAggregation' { $(if ($thr) { "$thr or more events in $win minutes" } else { 'Grouped events' }) } 'AnomalyAggregation' { 'Unusual volume of events' } default { $(if ($agg) { $agg } else { 'Every matching event' }) } }
        , @([string]$a.Name, [string](Get-Val $a 'Severity'), $state, $trigger, $who)
    }
}

function Get-TransportRuleRows {
    param([object[]]$TransportRules)
    foreach ($t in ($TransportRules | Sort-Object { [int]('0' + (Get-Val $_ 'Priority')) })) {
        $desc = ([string](Get-Val $t 'Description')) -replace '\s*[\r\n]+\s*', ' ' -replace '\s{2,}', ' '
        if ($desc.Length -gt 350) { $desc = $desc.Substring(0, 350) + ' ...' }
        $sit = if (Get-Val $t 'MessageContainsDataClassifications') { 'Yes' } else { 'No' }
        $mode = [string](Get-Val $t 'Mode'); $mode = switch ($mode) { 'Enforce' { 'Enforce' } 'Audit' { 'Test, no policy tips' } 'AuditAndNotify' { 'Test with policy tips' } default { $mode } }
        , @([string]$t.Name, "$([string](Get-Val $t 'State')), $mode", $sit, $desc)
    }
}

function Get-LabelNameMap {
    param([object[]]$Labels)
    $map = @{}
    foreach ($l in $Labels) {
        $dn = if ($l.DisplayName) { [string]$l.DisplayName } else { [string]$l.Name }
        foreach ($k in @([string]$l.Guid, [string]$l.Name)) { if ($k -and -not $map.ContainsKey($k)) { $map[$k] = $dn } }
    }
    return $map
}

# ---------- Assessment analysis: licensing, coverage map, scenarios, hygiene checks, maturity, comparison ----------

$Script:LicenceText = @{
    Core     = 'Included in Microsoft 365 E3 / Business Premium'
    Advanced = 'Microsoft 365 E5, E5 Information Protection & Governance, or Microsoft Purview Suite (formerly E5 Compliance)'
    None     = 'No additional licence'
}
$Script:LicenceText['Teams'] = 'Teams chat and channel DLP: ' + $Script:LicenceText['Advanced']
$Script:LicenceText['Copilot'] = $Script:LicenceText['Advanced'] + ', plus Microsoft 365 Copilot for the users in scope'
$Script:LicenceText['Labels'] = 'Manual and default labelling: E3 / Business Premium. Automatic (service-side) labelling: ' + $Script:LicenceText['Advanced']
$Script:LicenceText['Alerts'] = 'Single-event alerts: E3 / Business Premium. Volume-based (aggregated) alerts: ' + $Script:LicenceText['Advanced']
$Script:LicenceText['Lifecycle'] = 'Retention policies and manually applied retention labels: E3 / Business Premium. Auto-applied labels, adaptive scopes, records, disposition review and event-based retention: ' + $Script:LicenceText['Advanced']
$Script:LicenceText['Risk'] = 'Microsoft 365 E5, E5 Insider Risk Management, or Microsoft Purview Suite (formerly E5 Compliance)'
$Script:LicenceText['AuditPrem'] = 'Audit (Standard) and eDiscovery holds: E3 / Business Premium. Audit (Premium) with audit log retention policies, and eDiscovery (Premium): Microsoft 365 E5, E5 eDiscovery & Audit, or Microsoft Purview Suite. Ten-year audit retention needs the 10-Year Audit Log Retention add-on'
$Script:TopicLicence = @{
    'Microsoft Purview Data Loss Prevention' = 'Core'; 'Sensitive Information Types' = 'Core'; 'Custom Sensitive Information Types' = 'Core'; 'Exact Data Match' = 'Advanced'
    'Trainable Classifiers' = 'Advanced'; 'Confidence Levels' = 'Core'; 'Sensitivity Labels and DLP' = 'Labels'; 'DLP Policies' = 'Core'; 'Conditions and Actions' = 'Core'
    'Exchange Online DLP' = 'Core'; 'SharePoint Online DLP' = 'Core'; 'OneDrive DLP' = 'Core'; 'Microsoft Teams DLP' = 'Teams'; 'Endpoint DLP' = 'Advanced'
    'USB and Removable Storage' = 'Advanced'; 'Printing' = 'Advanced'; 'Browser and Cloud Uploads' = 'Advanced'; 'Policy Tips' = 'Core'; 'Block with Override' = 'Core'
    'DLP Simulation' = 'Core'; 'Activity Explorer' = 'Advanced'; 'DLP Alerts' = 'Alerts'; 'DLP and Microsoft 365 Copilot' = 'Copilot'; 'Operational Ownership' = 'None'
    'Phased DLP Deployment' = 'Core'; 'Collection gaps' = 'None'; 'Content Explorer' = 'Advanced'; 'DLP Rule Activity' = 'Core'
    'Sensitivity Label Settings' = 'Labels'; 'Retention Policies' = 'Lifecycle'; 'Retention Labels and Records' = 'Lifecycle'; 'Insider Risk Management' = 'Risk'
    'Communication Compliance' = 'Risk'; 'Information Barriers' = 'Risk'; 'Audit Log Retention' = 'AuditPrem'; 'eDiscovery Cases and Holds' = 'AuditPrem'
}

function Get-LicenceKey {
    <# Works out the licence an action needs from its area and wording. #>
    param([string]$Area, [string]$Text)
    $t = "$Area $Text"
    if ($Area -eq 'Data lifecycle') { return 'Lifecycle' }
    if ($Area -match '^(Insider risk|Communication compliance|Information barriers)$') { return 'Risk' }
    if ($Area -match '^(Audit|eDiscovery)$') { return 'AuditPrem' }
    if ($t -match 'Copilot') { return 'Copilot' }
    if ($t -match 'Endpoint|device|USB|removable|print|browser|cloud upload') { return 'Advanced' }
    if ($t -match 'Teams') { return 'Teams' }
    if ($t -match 'Exact Data Match|\bEDM\b|trainable|Activity Explorer|Content Explorer') { return 'Advanced' }
    if ($t -match 'auto-label|automatic label|auto label') { return 'Labels' }
    if ($t -match 'alert') { return 'Alerts' }
    if ($Area -match 'Ownership|Collection') { return 'None' }
    return 'Core'
}

function Get-EffortText {
    param([string]$Area, [string]$Text)
    $t = "$Area $Text"
    if ($t -match 'Endpoint|device|onboard|Exact Data Match|\bEDM\b|trainable|Copilot|No DLP policies|retention schedule|No insider risk') { return 'Large' }
    if ($Area -match '^(eDiscovery|Audit)$') { return 'Small' }
    if ($t -match 'polic|extend|Teams|label|simulation|enforce|overlap|tailor') { return 'Medium' }
    return 'Small'
}

function Get-DataCategory {
    param([string]$Name, [string]$RefType)
    if ($RefType -eq 'SensitivityLabel') { return 'Labelled content' }
    if ($Name -match 'password|\bkey\b|secret|token|connection string|credential|certificate|shared access') { return 'Credentials' }
    if ($Name -match 'health|medical|NHS|drug enforcement|patient|ICD-|disease|medic') { return 'Health' }
    if ($Name -match 'credit card|debit card|bank|IBAN|SWIFT|routing|account number|financial|tax|\bVAT\b|payment|sort code') { return 'Financial' }
    if ($Name -match 'passport|national insurance|social security|identity|ID card|identification|driver|licen|full name|address|birth|citizen|residen|national ID|personal') { return 'Personal and ID' }
    return 'Other'
}
$Script:DataCategories = @('Financial', 'Health', 'Personal and ID', 'Credentials', 'Labelled content', 'Other')
$Script:MapLocations = [ordered]@{ Exchange = 'Email (Exchange)'; SharePoint = 'SharePoint'; OneDrive = 'OneDrive'; Teams = 'Teams'; Endpoint = 'Devices'; Copilot = 'Copilot' }

function Get-RuleLevel {
    <# Strength of a rule's response: 4 block, 3 override/warn, 2 notify, 1 audit. Endpoint uses the device restrictions. #>
    param($RuleSum, [object[]]$EndpointActions, [switch]$ForDevices)
    if ($ForDevices) {
        $vals = @($EndpointActions | Where-Object { $_.Policy -eq $RuleSum.Policy -and $_.Rule -eq $RuleSum.Rule } | ForEach-Object { [string]$_.Value })
        if ($vals -match '^Block') { return 4 }
        if ($vals -match 'Warn|Override') { return 3 }
        if ($vals.Count) { return 1 }
    }
    switch ($RuleSum.ResponseInterpretation) {
        'Block' { return 4 } 'Block with override' { return 3 } 'Notify / policy tip' { return 2 } default { return 1 }
    }
}

function Get-CoverageMap {
    <# Rows: location x data category, with the strongest response that applies. #>
    param([object[]]$Coverage, [object[]]$RuleSum, [object[]]$ClassRows, [object[]]$EndpointActions)
    $cell = @{}
    foreach ($r in $RuleSum) {
        if ((ConvertTo-CellValue $r.Disabled) -eq 'True') { continue }
        $pc = $Coverage | Where-Object { $_.Policy -eq $r.Policy } | Select-Object -First 1
        if (-not $pc -or $pc.Mode -eq 'Disable') { continue }
        $sim = [string]$pc.Mode -like 'Test*'
        $cats = @($ClassRows | Where-Object { $_.Policy -eq $r.Policy -and $_.Rule -eq $r.Rule -and $_.Source -ne 'Exception' } | ForEach-Object { Get-DataCategory ([string]$_.Name) $_.ReferenceType } | Select-Object -Unique)
        if ($cats.Count -eq 0) { $cats = @('Other') }
        foreach ($loc in $Script:MapLocations.Keys) {
            if (-not $pc.$loc) { continue }
            $lvl = if ($sim) { 0.5 } else { Get-RuleLevel -RuleSum $r -EndpointActions $EndpointActions -ForDevices:($loc -eq 'Endpoint') }
            foreach ($c in $cats) { $k = "$loc|$c"; if (-not $cell.ContainsKey($k) -or $cell[$k] -lt $lvl) { $cell[$k] = $lvl } }
        }
    }
    foreach ($loc in $Script:MapLocations.Keys) {
        $row = [ordered]@{ Location = $Script:MapLocations[$loc] }
        foreach ($c in $Script:DataCategories) {
            $k = "$loc|$c"
            $row[$c] = if (-not $cell.ContainsKey($k)) { '' } else {
                switch ($cell[$k]) { 4 { 'Block' } 3 { $(if ($loc -eq 'Endpoint') { 'Warn' } else { 'Override' }) } 2 { 'Notify' } 1 { 'Audit only' } default { 'Simulation' } }
            }
        }
        [pscustomobject]$row
    }
}

function Get-MapCellMarkup {
    param([string]$Value)
    switch ($Value) { 'Block' { '[[G]]Block' } 'Override' { '[[G]]Override' } 'Warn' { '[[G]]Warn' } 'Notify' { '[[A]]Notify' } 'Audit only' { '[[A]]Audit only' } 'Simulation' { '[[A]]Simulation' } default { '[[N]]-' } }
}

function Get-ScenarioText {
    <# One plain-English "what actually happens" paragraph for a rule. #>
    param($Rule, $RuleSum, $Policy, [object[]]$ClassRows, [hashtable]$LabelNames, [object[]]$EndpointActions)
    $locs = @()
    $wl = [string](ConvertTo-CellValue (Get-Val $Policy 'Workload'))
    if ($wl -match 'Exchange') { $locs += 'emails' }
    if ($wl -match 'SharePoint|OneDrive') { $locs += 'shares a file' }
    if ($wl -match 'Teams') { $locs += 'posts a Teams message' }
    $device = $wl -match 'Endpoint'
    if ($device -and $locs.Count -eq 0) { $locs += 'copies, prints or uploads a file on a device' }
    if ($locs.Count -eq 0) { $locs += 'works with content' }
    $verb = if ($locs.Count -gt 1) { (($locs[0..($locs.Count - 2)]) -join ', ') + ' or ' + $locs[-1] } else { $locs[0] }
    $cond = @($ClassRows | Where-Object { $_.Source -ne 'Exception' })
    $what = if ($cond.Count -eq 0) { 'sensitive content' } else {
        $first = $cond[0]
        $nm = [string]$first.Name
        if ($first.ReferenceType -eq 'SensitivityLabel') { $n2 = if ($LabelNames.ContainsKey($nm)) { $LabelNames[$nm] } else { $nm }; "content labelled **$n2**" }
        else {
            $min = [int]('0' + [string]$first.MinCount)
            $t = if ($min -gt 1) { "**$min or more** items of **$nm**" } else { "a **$nm**" }
            if ($cond.Count -gt 1) { $t += " (or any of the other $($cond.Count - 1) listed types)" }
            $t
        }
    }
    $scope = [string](Get-Val $Rule 'AccessScope')
    $who = if ($scope -eq 'NotInOrganization') { ' with someone **outside the organisation**' } elseif ($scope -eq 'InOrganization') { ' inside the organisation' } else { '' }
    $mode = [string]$Policy.Mode
    $outcome = @()
    if ($mode -eq 'TestWithoutNotifications') { $outcome += 'nothing is blocked and the user sees nothing, because the policy is in silent simulation. The match is only recorded for review' }
    elseif ($mode -eq 'TestWithNotifications') { $outcome += 'the user may see a policy tip, but nothing is blocked because the policy is in simulation' }
    elseif ($mode -eq 'Disable') { $outcome += 'nothing happens, because the policy is disabled' }
    else {
        switch ($RuleSum.ResponseInterpretation) {
            'Block' { $outcome += 'the action is **blocked**' }
            'Block with override' { $outcome += 'the action is **blocked unless the user overrides it** with a business justification' }
            'Notify / policy tip' { $outcome += 'the action **goes ahead**; the user is warned with a policy tip and/or email' }
            default { if (-not $device) { $outcome += 'the action **goes ahead** and the user is not told; the event is only recorded' } }
        }
        if ($device) {
            $ea = @($EndpointActions | Where-Object { $_.Policy -eq $RuleSum.Policy -and $_.Rule -eq $RuleSum.Rule })
            foreach ($e in $ea) {
                $nm = if ($Script:EndpointActionNames.ContainsKey([string]$e.Setting)) { $Script:EndpointActionNames[[string]$e.Setting] } else { ConvertFrom-CamelCase ([string]$e.Setting) }
                $outcome += "on a device, **$($nm.ToLower())** is $(switch -Regex ([string]$e.Value) { '^Block' { 'blocked' } 'Warn|Override' { 'warned (the user can continue)' } default { 'audited only' } })"
            }
        }
        $alert = ConvertTo-CellValue (Get-Val $Rule 'GenerateAlert')
        if ($alert) { $outcome += 'an **alert** is raised for the security team' }
        $inc = ConvertTo-CellValue (Get-Val $Rule 'GenerateIncidentReport')
        if ($inc) { $outcome += "an incident report is emailed to $(Get-PersonText ($inc -split ' \| '))" }
    }
    if ($outcome.Count -eq 0) { $outcome = @('the action goes ahead and the event is recorded') }
    $o = if ($outcome.Count -gt 1) { (($outcome[0..($outcome.Count - 2)]) -join '; ') + '; and ' + $outcome[-1] } else { $outcome[0] }
    return "If someone $verb containing $what$who, $o."
}

function Get-HygieneChecks {
    param([object[]]$Policies, [object[]]$PolSum, [object[]]$Rules, [object[]]$RuleSum, [object[]]$ClassRows, [object[]]$Coverage,
        [object[]]$SitView, [object[]]$Labels, [object[]]$LabelPolicies, [object[]]$AlertPolicies, [object[]]$CeCounts, [object[]]$Activity, [bool]$ActivityCollected)
    $out = [System.Collections.Generic.List[object]]::new()
    function Add-Check { param($Name, $Status, $Detail, $Advice) $out.Add([pscustomobject][ordered]@{ Check = $Name; Status = $Status; Detail = $Detail; Advice = $Advice }) }

    # 1 Unchanged templates
    $tpl = @($PolSum | Where-Object { [string]$_.Comment -match '^Helps detect the presence of information commonly considered' })
    $tplRules = @($RuleSum | Where-Object { $_.Rule -match '^(Low|High) volume of content detected' })
    if ($Policies.Count -eq 0) { Add-Check 'Policies tailored to the organisation' 'N' 'No DLP policies.' '' }
    elseif ($tpl.Count -or $tplRules.Count) { Add-Check 'Policies tailored to the organisation' 'A' "$(Get-Plural $tpl.Count 'policy still uses' 'policies still use') Microsoft's template description and $(Get-Plural $tplRules.Count 'rule keeps' 'rules keep') the template rule name ($(@($tpl | ForEach-Object Name) -join ', '))." 'Review each template against the organisation''s own data, thresholds and responses, then rename and describe it so its purpose is clear.' }
    else { Add-Check 'Policies tailored to the organisation' 'G' 'No policies appear to be unchanged Microsoft templates.' '' }

    # 2 High-volume / high-severity rules that do not block
    $weak = @($RuleSum | Where-Object { ($_.Rule -match 'High volume' -or [string]$_.ReportSeverityLevel -eq 'High') -and $_.ResponseInterpretation -notmatch '^Block' })
    if ($RuleSum.Count -eq 0) { Add-Check 'High-risk rules take action' 'N' 'No rules.' '' }
    elseif ($weak.Count) { Add-Check 'High-risk rules take action' 'R' "$(Get-Plural $weak.Count 'high-volume or high-severity rule does' 'high-volume or high-severity rules do') not block or require an override: $(@($weak | ForEach-Object Rule) -join '; ')." 'For the highest-risk cases (large volumes, external sharing), use block with override at least, after a period in simulation.' }
    else { Add-Check 'High-risk rules take action' 'G' 'High-volume and high-severity rules block or require an override.' '' }

    # 3 Alert recipients
    $adminOnly = @($AlertPolicies | Where-Object { $u = @(Get-Val $_ 'NotifyUser' | ForEach-Object { [string]$_ }); $u.Count -gt 0 -and -not ($u | Where-Object { $_ -notmatch '^(TenantAdmins|SiteAdmin)$' }) })
    $noRecipients = @($AlertPolicies | Where-Object { -not (Get-Val $_ 'NotifyUser') })
    if ($AlertPolicies.Count -eq 0) { Add-Check 'Alerts reach a named team' 'N' 'No DLP alert policies.' '' }
    elseif ($adminOnly.Count -or $noRecipients.Count) { Add-Check 'Alerts reach a named team' 'A' "$(Get-Plural ($adminOnly.Count + $noRecipients.Count) 'alert policy emails' 'alert policies email') only tenant admins or nobody." 'Send DLP alerts to a monitored shared mailbox or security team with a documented triage process.' }
    else { Add-Check 'Alerts reach a named team' 'G' 'Alert policies email named recipients.' '' }

    # 4 Overlapping rules (same detection, overlapping locations, different policies)
    $sig = @{}
    foreach ($r in $RuleSum) {
        $names = @($ClassRows | Where-Object { $_.Policy -eq $r.Policy -and $_.Rule -eq $r.Rule -and $_.Source -ne 'Exception' } | ForEach-Object { "$($_.Name):$($_.MinCount)" } | Sort-Object -Unique)
        if ($names.Count -eq 0) { continue }
        $k = $names -join ','
        if (-not $sig.ContainsKey($k)) { $sig[$k] = [System.Collections.Generic.List[object]]::new() }
        $sig[$k].Add($r)
    }
    $dups = @($sig.Values | Where-Object { @($_ | Select-Object -ExpandProperty Policy -Unique).Count -gt 1 })
    if ($dups.Count) { Add-Check 'No overlapping rules' 'A' "$(Get-Plural $dups.Count 'set of rules looks' 'sets of rules look') for the same information in different policies: $(@($dups | ForEach-Object { (@($_ | ForEach-Object { "$($_.Policy) / $($_.Rule)" })) -join ' and ' }) -join '; ')." 'Check the overlap is intended; duplicate rules generate duplicate alerts and conflicting user experiences.' }
    elseif ($RuleSum.Count) { Add-Check 'No overlapping rules' 'G' 'No rules with identical detection were found in different policies.' '' }

    # 5 Stale configuration
    $old = @($PolSum | Where-Object { try { $_.WhenChangedUTC -and ([datetime]$_.WhenChangedUTC) -lt (Get-Date).AddMonths(-12) } catch { $false } })
    if ($Policies.Count -and $old.Count) { Add-Check 'Configuration reviewed in the last 12 months' 'A' "$(Get-Plural $old.Count 'policy has' 'policies have') not changed for over 12 months: $(@($old | ForEach-Object Name) -join ', ')." 'Review policies at least annually, and whenever the organisation''s data or regulations change.' }
    elseif ($Policies.Count) { Add-Check 'Configuration reviewed in the last 12 months' 'G' 'All policies have been changed within the last 12 months.' '' }

    # 6 Simulation left running
    $simOld = @($PolSum | Where-Object { [string]$_.Mode -like 'Test*' -and $(try { ([datetime]$_.WhenChangedUTC) -lt (Get-Date).AddDays(-30) } catch { $false }) })
    if ($simOld.Count) { Add-Check 'Simulation has an end date' 'A' "$(Get-Plural $simOld.Count 'policy has' 'policies have') been in simulation for over 30 days: $(@($simOld | ForEach-Object Name) -join ', ')." 'Review simulation results and set a date to enforce or retire each policy.' }
    elseif ($Policies.Count) { Add-Check 'Simulation has an end date' 'G' 'No policies have been left in simulation for more than 30 days.' '' }

    # 7 Disabled rules
    $dis = @($RuleSum | Where-Object { (ConvertTo-CellValue $_.Disabled) -eq 'True' })
    if ($dis.Count) { Add-Check 'No disabled rules left behind' 'A' "$(Get-Plural $dis.Count 'rule is' 'rules are') disabled: $(@($dis | ForEach-Object Rule) -join '; ')." 'Remove disabled rules that are no longer needed so the configuration stays easy to understand.' }
    elseif ($RuleSum.Count) { Add-Check 'No disabled rules left behind' 'G' 'No disabled rules.' '' }

    # 8 Default policy tips
    $tips = @($RuleSum | Where-Object { $_.NotifyUser }); $custom = @($tips | Where-Object { $_.NotifyPolicyTipCustomText })
    if ($tips.Count -and $custom.Count -eq 0) { Add-Check 'Policy tips use organisation wording' 'A' 'All policy tips use Microsoft''s default text.' 'Write short tips that say what is wrong, what to do and who to contact.' }
    elseif ($tips.Count) { Add-Check 'Policy tips use organisation wording' 'G' "$($custom.Count) of $($tips.Count) notifying rules use custom tip text." '' }

    # 9 Unused custom SITs
    $unusedSit = @($SitView | Where-Object { $_.IsCustom -and -not $_.UsedInDlpRules })
    if ($unusedSit.Count) { Add-Check 'Custom detectors are used' 'A' "$(Get-Plural $unusedSit.Count 'custom sensitive information type is' 'custom sensitive information types are') not used by any rule: $(@($unusedSit | Select-Object -First 10 | ForEach-Object Name) -join ', ')." 'Use them in a policy or remove them.' }
    elseif (@($SitView | Where-Object IsCustom).Count) { Add-Check 'Custom detectors are used' 'G' 'All custom sensitive information types are used in rules.' '' }

    # 10 Labels not published
    $published = @{}
    foreach ($lp in $LabelPolicies) { foreach ($l in @(Get-Val $lp 'Labels') + @(Get-Val $lp 'ScopedLabels')) { $published[[string](Get-ItemText $l)] = $true } }
    $unpub = @($Labels | Where-Object { -not ($published.ContainsKey([string]$_.Name) -or $published.ContainsKey([string]$_.Guid) -or $published.ContainsKey([string]$_.DisplayName)) })
    if ($Labels.Count -and $unpub.Count) { Add-Check 'Labels are published' 'A' "$(Get-Plural $unpub.Count 'label is' 'labels are') not in any label policy: $(@($unpub | ForEach-Object { if ($_.DisplayName) { $_.DisplayName } else { $_.Name } }) -join ', ')." 'Publish the labels users need, or remove unused ones.' }
    elseif ($Labels.Count) { Add-Check 'Labels are published' 'G' 'All labels are published in at least one label policy.' '' }

    # 11 Sensitive data stored but not protected (Content Explorer)
    if ($CeCounts.Count) {
        $unprot = @($CeCounts | Where-Object { $_.TagType -eq 'SensitiveInformationType' -and -not $_.UsedInDlpRules -and $_.ItemCount -gt 0 } | Group-Object TagName | ForEach-Object { "$($_.Name) ($('{0:N0}' -f ($_.Group | Measure-Object ItemCount -Sum).Sum) items)" })
        if ($unprot.Count) { Add-Check 'Stored sensitive data is covered by a rule' 'R' "Content Explorer found data that no DLP rule looks for: $($unprot -join '; ')." 'Decide whether this information needs protecting and, if so, add it to a policy.' }
        else { Add-Check 'Stored sensitive data is covered by a rule' 'G' 'Every sensitive information type found by Content Explorer is used in a DLP rule.' '' }
    }
    else { Add-Check 'Stored sensitive data is covered by a rule' 'N' 'Content Explorer counts were not collected.' 'Run the complete report with Content Explorer enabled (needs the Content Explorer List Viewer role).' }

    # 12 Rules that never fire / heavy overrides
    if ($ActivityCollected) {
        $active = @{}; foreach ($a in $Activity) { $active["$($a.Policy)|$($a.Rule)"] = $a }
        $silent = @($RuleSum | Where-Object { -not $active.ContainsKey("$($_.Policy)|$($_.Rule)") })
        $heavy = @($Activity | Where-Object { $_.Matches -ge 10 -and ($_.Overrides + $_.FalsePositives) / [double]$_.Matches -ge 0.3 })
        if ($heavy.Count) { Add-Check 'Rules are not routinely overridden' 'R' "$(Get-Plural $heavy.Count 'rule is' 'rules are') overridden or reported as false positives in 30% or more of matches: $(@($heavy | ForEach-Object Rule) -join '; ')." 'Review the justifications; tune thresholds or exceptions so users only see the rule when it matters.' }
        else { Add-Check 'Rules are not routinely overridden' 'G' 'No rule has a high override or false-positive rate.' '' }
        if ($silent.Count) { Add-Check 'Rules are triggering' 'A' "$(Get-Plural $silent.Count 'rule had' 'rules had') no matches in the audit period: $(@($silent | Select-Object -First 10 | ForEach-Object Rule) -join '; ')." 'Confirm the rule still detects what it should (test with sample data) or retire it.' }
        else { Add-Check 'Rules are triggering' 'G' 'Every rule matched at least once in the audit period.' '' }
    }
    else { Add-Check 'Rules are triggering' 'N' 'DLP rule activity was not collected.' 'Run the complete report with DLP rule activity enabled (needs Exchange Online and an audit log reader role).' }
    $out
}

function Get-MaturityScores {
    param([object[]]$Policies, [object[]]$RuleSum, [object[]]$ClassRows, [hashtable]$CovCount, [object[]]$SitView, [object[]]$Edm, [object[]]$Labels,
        [object[]]$LabelPolicies, [object[]]$AutoPolicies, [object[]]$AlertPolicies, [object[]]$RoleMembers, [object[]]$Hygiene, [bool]$AuditOn, [bool]$EvidenceCollected)
    $rows = [System.Collections.Generic.List[object]]::new()
    function Add-Score { param($Dim, [int]$Score, $Evidence, $Next) $rows.Add([pscustomobject][ordered]@{ Dimension = $Dim; Score = [Math]::Max(0, [Math]::Min(5, $Score)); Evidence = $Evidence; NextStep = $Next }) }

    $core = @('Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Endpoint', 'Copilot' | Where-Object { $CovCount[$_] -gt 0 })
    $s = [int][Math]::Round(5 * $core.Count / 6)
    Add-Score 'Coverage' $s "$($core.Count) of 6 core locations have an enforced policy ($($core -join ', '))." $(if ($s -lt 5) { "Extend protection to: $(@('Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Endpoint', 'Copilot' | Where-Object { $core -notcontains $_ }) -join ', ')." } else { 'Maintain coverage as new services are adopted.' })

    $active = @($RuleSum | Where-Object { (ConvertTo-CellValue $_.Disabled) -ne 'True' })
    $strong = @($active | Where-Object { $_.ResponseInterpretation -match '^Block' -or $_.ResponseInterpretation -like 'Endpoint*' })
    $notify = @($active | Where-Object { $_.ResponseInterpretation -eq 'Notify / policy tip' })
    $custTip = @($active | Where-Object { $_.NotifyPolicyTipCustomText })
    $s = if ($active.Count -eq 0) { 0 } elseif ($strong.Count -eq 0 -and $notify.Count -eq 0) { 1 } elseif ($strong.Count -eq 0) { 2 } elseif ($strong.Count / [double]$active.Count -lt 0.5) { 3 } elseif ($custTip.Count -eq 0) { 4 } else { 5 }
    Add-Score 'Enforcement' $s "$($strong.Count) of $($active.Count) active rules block, require an override or restrict devices; $($notify.Count) only notify." $(switch ($s) { 0 { 'Create rules for the agreed use cases.' } 1 { 'Turn on user notifications so people learn the policy.' } 2 { 'Move the highest-risk rules to block with override.' } 3 { 'Extend blocking to more high-risk rules.' } 4 { 'Add custom policy tip wording.' } default { 'Keep tuning with real activity data.' } })

    $sitRefs = @($ClassRows | Where-Object { $_.ReferenceType -eq 'Classifier' }); $labRefs = @($ClassRows | Where-Object { $_.ReferenceType -eq 'SensitivityLabel' })
    $tiers = @($sitRefs | Where-Object { [int]('0' + [string]$_.MinCount) -gt 1 }).Count -gt 0
    $customUsed = @($SitView | Where-Object { $_.IsCustom -and $_.UsedInDlpRules }).Count
    $advanced = @($sitRefs | Where-Object { $_.ClassifierType -match 'MLModel|Trainable|ExactMatch|EDM' }).Count
    $s = if ($sitRefs.Count -eq 0 -and $labRefs.Count -eq 0) { 0 } else { 1 + [int]$tiers + [int]($labRefs.Count -gt 0) + [int]($customUsed -gt 0) + [int]($advanced -gt 0 -or $Edm.Count -gt 0) }
    Add-Score 'Detection' $s "$(@($sitRefs | Select-Object -ExpandProperty Name -Unique).Count) sensitive information types in rules; volume tiers: $(if ($tiers) { 'yes' } else { 'no' }); labels as conditions: $($labRefs.Count); custom types in use: $customUsed; EDM or trainable classifiers: $advanced." $(if ($s -lt 5) { 'Add organisation-specific detection: custom types, labels as conditions, EDM or trainable classifiers.' } else { 'Review detection accuracy with simulation and false-positive data.' })

    $mand = @($LabelPolicies | Where-Object { (@(Get-Val $_ 'Settings') -join ' ') -match '\[mandatory, true\]|\[defaultlabelid, [^\]]+\]' })
    $s = if ($Labels.Count -eq 0) { 0 } elseif ($LabelPolicies.Count -eq 0) { 1 } else { 2 + [int]($mand.Count -gt 0) + [int]($AutoPolicies.Count -gt 0) + [int]($labRefs.Count -gt 0) }
    Add-Score 'Labelling integration' $s "$($Labels.Count) labels, $($LabelPolicies.Count) label policies ($($mand.Count) with mandatory or default labelling), $($AutoPolicies.Count) auto-labelling policies, labels used in DLP: $($labRefs.Count)." $(if ($labRefs.Count -eq 0 -and $Labels.Count) { 'Use labels as DLP conditions so labelled content is protected wherever it goes.' } elseif ($AutoPolicies.Count -eq 0 -and $Labels.Count) { 'Add auto-labelling for the most sensitive information.' } else { 'Review label adoption and accuracy.' })

    $alertRules = @($RuleSum | Where-Object { $_.GenerateAlert -or $_.GenerateIncidentReport })
    $named = @($AlertPolicies | Where-Object { @(Get-Val $_ 'NotifyUser' | ForEach-Object { [string]$_ } | Where-Object { $_ -notmatch '^(TenantAdmins|SiteAdmin)$' }).Count -gt 0 })
    $s = [int]($alertRules.Count -gt 0) + [int]($named.Count -gt 0) + [int]$AuditOn + [int]$EvidenceCollected + [int](@($RoleMembers).Count -gt 0)
    Add-Score 'Operations and monitoring' $s "Rules raising alerts: $($alertRules.Count); alert policies to a named team: $($named.Count); unified audit log on: $(if ($AuditOn) { 'yes' } else { 'no / not checked' }); rule activity reviewed in this assessment: $(if ($EvidenceCollected) { 'yes' } else { 'no' }); people with DLP roles: $(@($RoleMembers | Select-Object -ExpandProperty Member -Unique).Count)." 'Route alerts to a named team, agree a review cadence and record who owns each policy.'

    $enf = @($Policies | Where-Object { $_.Mode -eq 'Enable' }).Count
    $issues = @($Hygiene | Where-Object { $_.Status -in 'R', 'A' }).Count
    $s = if ($Policies.Count -eq 0) { 0 } elseif ($enf -eq 0) { 2 } elseif ($issues -ge 4) { 3 } elseif ($issues -ge 1) { 4 } else { 5 }
    Add-Score 'Lifecycle and hygiene' $s "$enf of $($Policies.Count) policies enforced; $issues hygiene checks need attention." 'Fix the hygiene items in section 15 and review the configuration at least annually.'
    $rows
}

function Get-MaturityLevel {
    param([double]$Score)
    if ($Score -lt 1) { 'Initial' } elseif ($Score -lt 2) { 'Developing' } elseif ($Score -lt 3) { 'Defined' } elseif ($Score -lt 4) { 'Managed' } else { 'Optimised' }
}

function Get-ScoreBar {
    param([int]$Score)
    $full = [string][char]0x25A0; $empty = [string][char]0x25A1
    ($full * $Score) + ($empty * (5 - $Score))
}

function ConvertTo-DateSafe {
    <# Reads dates written by either PowerShell 5.1 (/Date(ms)/) or PowerShell 7 (ISO 8601). Returns $null if unreadable. #>
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $m = [regex]::Match([string]$Value, 'Date\((-?\d+)')
    if ($m.Success) { return [DateTimeOffset]::FromUnixTimeMilliseconds([int64]$m.Groups[1].Value).UtcDateTime }
    try { return ([datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)).ToUniversalTime() } catch { return $null }
}

function Find-PreviousRun {
    <# Latest earlier complete-report folder for the same tenant, or $null. #>
    param([string]$TenantId)
    $dirs = @(Get-ChildItem -Path $Script:Settings.OutputRoot -Directory -Filter 'M365-DLP-Full-*' -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $Script:OutputDir } | Sort-Object Name -Descending)
    foreach ($d in $dirs) {
        $ti = Join-Path $d.FullName '01_Tenant_Info.json'
        if (-not (Test-Path (Join-Path $d.FullName '10_DLP_Policies_Full.json'))) { continue }
        if (-not $TenantId) { return $d.FullName }
        try { if ((Get-Content $ti -Raw | ConvertFrom-Json).TenantId -eq $TenantId) { return $d.FullName } } catch { }
    }
    return $null
}

function Read-RunFile {
    param([string]$Folder, [string]$Name)
    $p = Join-Path $Folder "$Name.json"
    if (-not (Test-Path $p)) { return @() }
    try { return @(Get-Content $p -Raw -Encoding UTF8 | ConvertFrom-Json | ForEach-Object { $_ }) } catch { return @() }
}

function Get-ChangesSinceLastRun {
    param([string]$Folder, [object[]]$Policies, [object[]]$Rules, [object[]]$Labels, [object[]]$LabelPolicies, [object[]]$AlertPolicies)
    $out = [System.Collections.Generic.List[object]]::new()
    function Add-Change { param($Type, $Name, $Change) $out.Add([pscustomobject][ordered]@{ Item = $Type; Name = $Name; Change = $Change }) }
    function Compare-Set {
        param([string]$Type, [object[]]$Old, [object[]]$New, [scriptblock]$Key, [scriptblock]$Describe)
        $o = @{}; foreach ($x in $Old) { $k = [string](& $Key $x); if ($k) { $o[$k] = $x } }
        $n = @{}; foreach ($x in $New) { $k = [string](& $Key $x); if ($k) { $n[$k] = $x } }
        foreach ($k in $n.Keys) {
            if (-not $o.ContainsKey($k)) { Add-Change $Type $k 'Added' }
            else { $d = & $Describe $o[$k] $n[$k]; if ($d) { Add-Change $Type $k $d } }
        }
        foreach ($k in $o.Keys) { if (-not $n.ContainsKey($k)) { Add-Change $Type $k 'Removed' } }
    }
    $modeText = @{ 'Enable' = 'Enforced'; 'TestWithNotifications' = 'Simulation with tips'; 'TestWithoutNotifications' = 'Silent simulation'; 'Disable' = 'Disabled' }
    $chg = { param($a, $b)
        $ca = ConvertTo-DateSafe (Get-Val $a 'WhenChangedUTC'); $cb = ConvertTo-DateSafe (Get-Val $b 'WhenChangedUTC')
        if ($ca -and $cb -and $cb -gt $ca.AddMinutes(1)) { return "Modified $($cb.ToString('dd MMM yyyy'))" }
        return '' }
    Compare-Set 'DLP policy' (Read-RunFile $Folder '10_DLP_Policies_Full') $Policies { param($x) $x.Name } {
        param($a, $b)
        $ma = [string]$a.Mode; $mb = [string]$b.Mode
        if ($ma -ne $mb) { return "Mode: $(if ($modeText[$ma]) { $modeText[$ma] } else { $ma }) -> $(if ($modeText[$mb]) { $modeText[$mb] } else { $mb })" }
        & $chg $a $b }
    Compare-Set 'DLP rule' (Read-RunFile $Folder '13_DLP_Rules_Full') $Rules { param($x) "$(Get-Val $x 'ParentPolicyName') / $($x.Name)" } {
        param($a, $b)
        $da = [string](Get-Val $a 'Disabled'); $db = [string](Get-Val $b 'Disabled')
        if ($da -ne $db) { return $(if ($db -eq 'True') { 'Disabled' } else { 'Enabled' }) }
        & $chg $a $b }
    Compare-Set 'Sensitivity label' (Read-RunFile $Folder '30_Sensitivity_Labels') $Labels { param($x) if ($x.DisplayName) { $x.DisplayName } else { $x.Name } } $chg
    Compare-Set 'Label policy' (Read-RunFile $Folder '31_Sensitivity_Label_Policies') $LabelPolicies { param($x) $x.Name } $chg
    Compare-Set 'Alert policy' (Read-RunFile $Folder '51_Alert_Policies_DLP') $AlertPolicies { param($x) $x.Name } { param($a, $b) $da = [string](Get-Val $a 'Disabled'); $db = [string](Get-Val $b 'Disabled'); if ($da -ne $db) { if ($db -eq 'True') { 'Turned off' } else { 'Turned on' } } }
    $out
}

function Add-DocImage {
    <# Adds a PNG or JPEG image as its own paragraph. WidthCm keeps the aspect ratio. #>
    param([string]$Path, [double]$WidthCm = 5, [string]$Align = 'left')
    if (-not $Path -or -not (Test-Path $Path)) { return }
    $bytes = [IO.File]::ReadAllBytes((Resolve-Path $Path).Path)
    $ext = [IO.Path]::GetExtension($Path).ToLower().TrimStart('.')
    if ($ext -eq 'jpg') { $ext = 'jpeg' }
    if ($ext -notin 'png', 'jpeg') { Write-Log "Logo must be a PNG or JPEG file: $Path" 'WARN'; return }
    $w = 0; $h = 0
    if ($ext -eq 'png' -and $bytes.Length -gt 24) {
        $w = ([int]$bytes[16] -shl 24) -bor ([int]$bytes[17] -shl 16) -bor ([int]$bytes[18] -shl 8) -bor [int]$bytes[19]
        $h = ([int]$bytes[20] -shl 24) -bor ([int]$bytes[21] -shl 16) -bor ([int]$bytes[22] -shl 8) -bor [int]$bytes[23]
    }
    else {
        try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop; $img = [System.Drawing.Image]::FromFile((Resolve-Path $Path).Path); $w = $img.Width; $h = $img.Height; $img.Dispose() } catch { }
    }
    if ($w -le 0 -or $h -le 0) { $w = 3; $h = 1 }
    $cx = [int64]($WidthCm * 360000); $cy = [int64]($cx * $h / $w)
    $n = $Script:Doc.Media.Count + 1
    $name = "image$n.$ext"
    $id = "rIdImg$n"
    $Script:Doc.Media.Add(@{ Name = $name; Bytes = $bytes })
    $Script:Doc.Rels.Add("<Relationship Id=`"$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$name`"/>")
    $picId = 500 + $n
    Add-DocXml ("<w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"200`"/><w:jc w:val=`"$Align`"/></w:pPr><w:r><w:drawing><wp:inline distT=`"0`" distB=`"0`" distL=`"0`" distR=`"0`"><wp:extent cx=`"$cx`" cy=`"$cy`"/><wp:docPr id=`"$picId`" name=`"Logo $n`"/>" +
        "<a:graphic xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`"><a:graphicData uri=`"http://schemas.openxmlformats.org/drawingml/2006/picture`"><pic:pic xmlns:pic=`"http://schemas.openxmlformats.org/drawingml/2006/picture`">" +
        "<pic:nvPicPr><pic:cNvPr id=`"$picId`" name=`"$name`"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=`"$id`"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=`"0`" y=`"0`"/><a:ext cx=`"$cx`" cy=`"$cy`"/></a:xfrm><a:prstGeom prst=`"rect`"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>")
}

function Export-ActionPlan {
    <# Prioritised action plan from the report observations, with effort and licence. Written as 80_Action_Plan. #>
    $order = @{ 'High' = 0; 'Medium' = 1; 'Low' = 2; 'Info' = 3 }
    $n = 0
    $rows = foreach ($o in @($Script:ReportObs | Sort-Object { $order[$_.Severity] })) {
        $n++
        [pscustomobject][ordered]@{
            Id = $n; Priority = $o.Severity; Area = $o.Area; Action = ($o.Text -replace '\*\*', '')
            Effort = Get-EffortText $o.Area $o.Text; Licence = $Script:LicenceText[(Get-LicenceKey $o.Area $o.Text)]
            Owner = ''; TargetDate = ''; Status = 'Not started'; Notes = ''
        }
    }
    $Script:Cache['ActionPlan'] = @($rows)
    Export-Dataset '80_Action_Plan' @($rows) -AlreadyFlat
}

function Add-ReportPartError {
    <# Keeps the rest of the report when one generated part fails. #>
    param([string]$Part, $ErrorRecord)
    $msg = $ErrorRecord.Exception.Message
    Write-Log "  Report part '$Part' could not be built: $msg" 'WARN'
    $Script:Errors.Add([pscustomobject]@{ Step = "Word report: $Part"; Outcome = 'Failed'; Error = $msg })
    Add-DocCallout 'Not available:' "This part of the report could not be generated ($msg). The rest of the report is unaffected." 'FFF4E5' 'E08A00'
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
    $labelNames = Get-LabelNameMap $labels
    $autoRules = @(Get-AutoLabelRules)
    $ceOn = $Script:Cache.ContainsKey('ContentExplorer'); $ceCounts = @(if ($ceOn) { $Script:Cache['ContentExplorer'] })
    $ceLocs = @(if ($Script:Cache.ContainsKey('ContentExplorerLocations')) { $Script:Cache['ContentExplorerLocations'] })
    $actOn = $Script:Cache.ContainsKey('DlpAudit'); $ruleAct = @(if ($actOn) { Get-DlpRuleActivity })
    $epAll = @(Get-EndpointRuleActions)
    $labelSet = @(Get-LabelSettingsView)
    $customer = if ($Script:Settings.CustomerName) { $Script:Settings.CustomerName } else { $null }

    # Sections 20 onwards are numbered by what this run includes
    $lifeOn = [bool]$Script:Settings.IncludeLifecycle; $riskOn = [bool]$Script:Settings.IncludeRiskCompliance; $audOn = [bool]$Script:Settings.IncludeAuditDiscovery
    $secNo = 20; $nLife = 0; $nRisk = 0; $nAudit = 0
    if ($lifeOn) { $nLife = $secNo; $secNo++ }
    if ($riskOn) { $nRisk = $secNo; $secNo++ }
    if ($audOn) { $nAudit = $secNo; $secNo++ }
    $nNext = $secNo; $nApply = $secNo + 1
    $Script:ReportSkipAreas = @(@{ Lifecycle = $lifeOn; RiskCompliance = $riskOn; AuditDiscovery = $audOn }.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
    $extraText = @(if ($lifeOn) { "data lifecycle and records management (section $nLife)" }; if ($riskOn) { "insider risk, communication compliance and information barriers (section $nRisk)" }; if ($audOn) { "audit and eDiscovery (section $nAudit)" })

    $pEnf = @($policies | Where-Object { $_.Mode -eq 'Enable' }).Count
    $pTest = @($policies | Where-Object { $_.Mode -like 'Test*' }).Count
    $pOff = @($policies | Where-Object { $_.Mode -eq 'Disable' }).Count
    $respCount = @{}
    foreach ($k in 'Block', 'Block with override', 'Notify / policy tip', 'Audit only') { $respCount[$k] = @($ruleSum | Where-Object { $_.ResponseInterpretation -eq $k }).Count }
    $respCount['Endpoint'] = @($ruleSum | Where-Object { $_.ResponseInterpretation -like 'Endpoint*' }).Count

    # ---------- cover ----------
    if ($Script:Settings.LogoPath -and (Test-Path $Script:Settings.LogoPath)) { Add-DocPara '' -After 1200; Add-DocImage -Path $Script:Settings.LogoPath -WidthCm 5 }
    else { Add-DocPara '' -After 2400 }
    Add-DocXml "<w:p><w:pPr><w:pStyle w:val=`"Title`"/></w:pPr>$(New-DocRuns 'Microsoft Purview Data Loss Prevention Assessment')</w:p>"
    Add-DocPara $(if ($customer) { $customer } else { $org }) -Size 32 -Color '444444' -After 600
    $coverRows = @()
    if ($customer) { $coverRows += , @('Prepared for', $customer) }
    if ($Script:Settings.AssessorName) { $coverRows += , @('Prepared by', $Script:Settings.AssessorName) }
    $coverRows += , @('Organisation', $org)
    $coverRows += , @('Tenant ID', $tid)
    $coverRows += , @('Data collected', $when)
    $coverRows += , @('Collected by', $by)
    $coverRows += , @('Method', 'Get-M365DlpConfiguration.ps1 (read-only discovery using Security & Compliance and Exchange Online PowerShell)')
    Add-DocTable -Headers @('Item', 'Detail') -Widths @(1, 3) -FontSize 19 -Rows $coverRows
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
    if ($extraText.Count) { Add-DocPara ("As well as DLP, this assessment covers related Microsoft Purview areas that affect how information is protected: " + $(if ($extraText.Count -gt 1) { ($extraText[0..($extraText.Count - 2)] -join ', ') + ' and ' + $extraText[-1] } else { $extraText[0] }) + '.') }
    Add-DocPara "Section $nApply gives step-by-step guidance for applying the configuration needed to close the gaps found."
    Add-DocPara "The executive summary scores the DLP programme on six dimensions (0 to 5). Section 4 includes a **coverage map** of data types against locations, section 6 shows **what actually happens** when someone shares sensitive data, section 15 lists **hygiene checks**, and section $nNext is a **prioritised action plan** with effort and licence notes. Where collected, section 7 shows **where sensitive data is stored** (Content Explorer) and section 17 shows **how often each rule fires** (audit log)."
    Add-DocPara 'Where an area has configuration worth reading in full, such as label policies, alert policies or Endpoint DLP settings, it is described in a **Configuration details** part of that section. Every DLP policy and rule is written out in plain English in **Appendix E**, so the report can be reviewed without opening the spreadsheets.'
    Add-DocPara 'Very long lists are shortened. The complete raw data for every area is still in the CSV, JSON and Excel files produced alongside this document (see Appendix C).'

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
    try {
        $map = @(Get-CoverageMap -Coverage $cov -RuleSum $ruleSum -ClassRows $rc -EndpointActions $epAll)
        Export-Dataset '84_Coverage_Map' $map -AlreadyFlat
        Add-DocHeading 'Coverage map: what is protected, where' 3
        Add-DocPara 'Each cell shows the strongest response that applies when that type of information is found in that location. Green means the action is stopped or needs an override; amber means it is allowed but the user is told, or only recorded; grey means no rule covers it.'
        Add-DocTable -Headers (@('Location') + $Script:DataCategories) -Widths @(1.5, 1.1, 1.1, 1.2, 1.1, 1.2, 1) -FontSize 16 -Rows @(foreach ($m in $map) { , (@($m.Location) + @(foreach ($c in $Script:DataCategories) { Get-MapCellMarkup $m.$c })) }) -Note 'Categories are worked out from the names of the sensitive information types in each rule. Simulation = the policy is not enforced yet.'
        $cons = @()
        $unprotCats = @($Script:DataCategories | Where-Object { $c = $_; $c -notin 'Other', 'Labelled content' -and -not ($map | Where-Object { $_.$c -in 'Block', 'Override', 'Warn' }) })
        if ($policies.Count -gt 0 -and $unprotCats.Count -gt 0) { $cons += "No location stops or challenges sharing of: **$($unprotCats -join ', ')**. Confirm whether these types of information exist in the organisation (section 7) and need protecting." }
    }
    catch { Add-ReportPartError 'Coverage map' $_ }
    $core = 'Exchange', 'SharePoint', 'OneDrive', 'Teams', 'Endpoint'
    $gaps = @($core | Where-Object { $covCount[$_] -eq 0 })
    if ($gaps.Count -gt 0 -and $policies.Count -gt 0) { $cons += Add-Obs 'Coverage' 'High' "No enforced policy covers: $($gaps -join ', '). Confirm whether these locations are in scope and, if so, extend existing policies or create new ones." }
    if ($covCount['Teams'] -eq 0 -and $covCount['Exchange'] -gt 0) { $cons += 'Email is protected but Teams is not. Users can often share the same information in a chat message, so consider applying equivalent rules to Teams.' }
    if ($covCount['Endpoint'] -eq 0) { $cons += 'Without Endpoint DLP, protection stops once a file is downloaded to a device. See section 13.' }
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
        Add-DocTable -Headers @('Policy', 'Rule', 'Response', 'Severity', 'Alert', 'State') -Widths @(2.4, 2.6, 1.8, 1, 0.8, 1) -Rows @($rows) -MaxRows 60 -Note 'Appendix E describes what each rule looks for, its exceptions and what it does, in plain English.'
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
    try {
        $scen = @()
        foreach ($p in ($policies | Where-Object { $_.Mode -ne 'Disable' } | Sort-Object { [int]('0' + (Get-Val $_ 'Priority')) })) {
            foreach ($r in @($rules | Where-Object { [string](Get-Val $_ 'ParentPolicyName') -eq [string]$p.Name -and (ConvertTo-CellValue (Get-Val $_ 'Disabled')) -ne 'True' } | Sort-Object { [int]('0' + (Get-Val $_ 'Priority')) })) {
                $rs = $ruleSum | Where-Object { $_.Policy -eq [string]$p.Name -and $_.Rule -eq [string]$r.Name } | Select-Object -First 1
                if (-not $rs) { continue }
                $scen += [pscustomobject]@{ Policy = [string]$p.Name; Rule = [string]$r.Name; Text = (Get-ScenarioText -Rule $r -RuleSum $rs -Policy $p -ClassRows @($rc | Where-Object { $_.Policy -eq [string]$p.Name -and $_.Rule -eq [string]$r.Name }) -LabelNames $labelNames -EndpointActions $epAll) }
            }
        }
        if ($scen.Count -gt 0) {
            Add-DocHeading 'What actually happens: example scenarios' 3
            Add-DocPara 'Each example below is worked out from a rule''s conditions and actions, to show the effect on a user in plain terms.'
            foreach ($sc in ($scen | Select-Object -First 16)) {
                Add-DocPara "**$($sc.Policy) - $($sc.Rule)**" -Size 19 -Color '1F3B57' -After 20
                Add-DocPara $sc.Text -After 140
            }
            if ($scen.Count -gt 16) { Add-DocPara "Showing 16 of $($scen.Count) rules; Appendix E describes every rule." -Size 16 -Color '666666' -Italic }
            $nothingStops = @($ruleSum | Where-Object { $_.ResponseInterpretation -match '^Block' }).Count -eq 0 -and $pEnf -gt 0
            if ($nothingStops) { $cons += Add-Obs 'User experience' 'High' 'In every scenario the action **goes ahead**: no enforced rule blocks or asks for an override, so DLP currently records and warns but does not prevent data leaving the organisation.' }
        }
    }
    catch { Add-ReportPartError 'Example scenarios' $_ }
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
    try {
        Add-DocHeading 'Where this information is stored (Content Explorer)' 3
        if (-not $ceOn) { Add-DocPara 'Content Explorer counts were not collected in this run. Enable Content Explorer in the complete report (it needs the **Content Explorer List Viewer** role) to show how many emails, files and messages contain each type of sensitive information.' }
        else {
            Add-SectionGap 'Content Explorer item counts'
            Add-DocPara 'Content Explorer counts the items that contain each sensitive information type or label, by location. These are items already stored in Microsoft 365, whether or not a DLP rule looks for them.'
            $ceRows = foreach ($g in ($ceCounts | Where-Object TagType -eq 'SensitiveInformationType' | Group-Object TagName)) {
                $tot = [int64](($g.Group | Measure-Object ItemCount -Sum).Sum)
                $cells = @(foreach ($w in $Script:CeWorkloads.Keys) { $x = $g.Group | Where-Object Workload -eq $w | Select-Object -First 1; if ($x -and $null -ne $x.ItemCount) { '{0:N0}' -f $x.ItemCount } else { 'n/a' } })
                $used = [bool]($g.Group[0].UsedInDlpRules)
                [pscustomobject]@{ Total = $tot; Row = (@([string]$g.Name) + $cells + @(('{0:N0}' -f $tot), $(if ($used) { '[[G]]Yes' } elseif ($tot -gt 0) { '[[R]]No' } else { '[[N]]No' }))) }
            }
            $ceRows = @($ceRows | Sort-Object Total -Descending)
            if ($ceRows.Count) { Add-DocTable -Headers @('Sensitive information type', 'Email', 'SharePoint', 'OneDrive', 'Teams', 'Total', 'In a DLP rule') -Widths @(3, 1, 1.1, 1.1, 0.9, 1, 1.1) -FontSize 16 -Rows @(foreach ($x in $ceRows) { , $x.Row }) -Note 'Detail: 62_Content_Explorer_Counts. n/a = the count could not be retrieved for that location.' }
            $labRows = foreach ($g in ($ceCounts | Where-Object TagType -eq 'Sensitivity' | Group-Object TagName)) {
                $tot = [int64](($g.Group | Measure-Object ItemCount -Sum).Sum)
                , (@([string]$g.Name) + @(foreach ($w in $Script:CeWorkloads.Keys) { $x = $g.Group | Where-Object Workload -eq $w | Select-Object -First 1; if ($x -and $null -ne $x.ItemCount) { '{0:N0}' -f $x.ItemCount } else { 'n/a' } }) + @(('{0:N0}' -f $tot)))
            }
            if (@($labRows).Count) { Add-DocPara 'Items carrying each sensitivity label:'; Add-DocTable -Headers @('Sensitivity label', 'Email', 'SharePoint', 'OneDrive', 'Teams', 'Total') -Widths @(3, 1, 1.1, 1.1, 0.9, 1) -FontSize 16 -Rows @($labRows) }
            if ($ceLocs.Count) {
                Add-DocPara 'Where the largest concentrations are (sample of up to 500 items per type and location):'
                Add-DocTable -Headers @('Sensitive information type', 'Location', 'Site or mailbox', 'Items in sample') -Widths @(2.4, 1.1, 4, 1) -FontSize 16 -Rows @(foreach ($l in $ceLocs) { , @([string]$l.TagName, [string]$l.Workload, [string]$l.Location, "$($l.ItemsInSample) of $($l.SampleSize)") }) -MaxRows 30 -Note 'Detail: 63_Content_Explorer_Top_Locations.'
            }
            $gapsCe = @($ceCounts | Where-Object { $_.TagType -eq 'SensitiveInformationType' -and -not $_.UsedInDlpRules -and $_.ItemCount -gt 0 } | Group-Object TagName)
            if ($gapsCe.Count) { $cons += Add-Obs 'Data discovery' 'High' "Sensitive information is stored in Microsoft 365 that no DLP rule looks for: $(@($gapsCe | ForEach-Object { '{0} ({1:N0} items)' -f $_.Name, ($_.Group | Measure-Object ItemCount -Sum).Sum }) -join '; '). Decide whether it needs protecting." }
            $big = @($ceCounts | Where-Object { $_.TagType -eq 'SensitiveInformationType' -and $_.UsedInDlpRules -and $_.ItemCount -gt 0 } | Group-Object TagName | Where-Object { ($_.Group | Measure-Object ItemCount -Sum).Sum -ge 1000 })
            if ($big.Count) { $cons += "Large volumes of protected information are stored ($(@($big | ForEach-Object Name) -join ', ')). Check the rules for these types are strong enough for the amount of data involved." }
        }
    }
    catch { Add-ReportPartError 'Content Explorer' $_ }
    Add-SectionEnd $cons @('SitOverview', 'SitDefinitions', 'ContentExplorer') -Area 'SIT'

    # ---------- 8 Custom SITs ----------
    Add-DocHeading '8. Custom sensitive information types and keyword dictionaries' 1
    Add-SectionIntro -What 'A **custom SIT** detects information in an organisation-specific format, such as an internal customer number, a project code or an employee ID. A **keyword dictionary** is a large list of words or phrases (for example code names or product names) that can be used as, or within, a custom SIT.' -Why 'Built-in detectors cannot recognise identifiers unique to the organisation. Custom SITs close that gap. Unused or poorly tuned custom SITs add complexity and can cause false positives.'
    Add-SectionGap 'Sensitive information types', 'SIT rule packages', 'Keyword dictionaries'
    if ($custom.Count -eq 0) { Add-DocPara 'No custom sensitive information types were found.' }
    else {
        $rows = foreach ($s in $custom) { , @([string]$s.Name, [string]$s.Publisher, [string]$s.RecommendedConfidence, $(if ($s.UsedInDlpRules) { 'Yes' } else { 'No' })) }
        Add-DocTable -Headers @('Custom SIT', 'Publisher', 'Recommended confidence', 'Used in DLP') -Widths @(3.5, 2, 1.5, 1) -Rows @($rows) -MaxRows 40 -Note 'Definitions of custom rule packages are saved as 22_RulePackage_*.xml.'
        $described = @($custom | Where-Object { $_.Description })
        if ($described.Count -gt 0) {
            Add-DocHeading 'Configuration details: what each custom SIT detects' 3
            Add-DocTable -Headers @('Custom SIT', 'Description (as entered by the author)') -Widths @(2.5, 5.5) -Rows @(foreach ($s in $described) { , @([string]$s.Name, [string]$s.Description) }) -MaxRows 40
        }
    }
    Add-DocPara "$(Get-Plural $dicts.Count 'keyword dictionary is' 'keyword dictionaries are') configured."
    if ($dicts.Count -gt 0) {
        Add-DocHeading 'Configuration details: keyword dictionaries' 3
        Add-DocTable -Headers @('Dictionary', 'Description', 'Keywords') -Widths @(2.5, 4.5, 1) -Rows @(Get-DictionaryRows $dicts) -MaxRows 30 -Note 'The keyword lists themselves are not reproduced here; review them in the Purview portal.'
    }
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
    if ($edm.Count -gt 0) {
        Add-DocHeading 'Configuration details: EDM schemas' 3
        Add-DocPara 'Each schema describes the columns of the uploaded dataset. **Searchable** columns are the ones Purview looks for in content; the others are used to confirm a match.'
        Add-DocTable -Headers @('Schema', 'Columns') -Widths @(2, 6) -Rows @(Get-EdmSchemaRows $edm) -MaxRows 20
    }
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
    if ($labelPol.Count -gt 0) {
        Add-DocHeading 'Configuration details: label policies' 3
        Add-DocPara 'Which labels each policy makes available, who receives them, and the main settings that change the user experience.'
        Add-DocTable -Headers @('Label policy', 'Labels published', 'Published to', 'Key settings') -Widths @(1.8, 2.6, 1.8, 2.6) -FontSize 16 -Rows @(Get-LabelPolicyRows $labelPol $labelNames) -MaxRows 25 -Note 'Advanced settings are in 31_Sensitivity_Label_Policies.'
    }
    if ($autoPol.Count -gt 0) {
        Add-DocHeading 'Configuration details: auto-labelling policies' 3
        Add-DocPara 'Which label each policy applies automatically, where, and what content it looks for.'
        Add-DocTable -Headers @('Auto-labelling policy', 'Applies label', 'Status', 'Where', 'Looks for') -Widths @(1.8, 1.3, 1.3, 2, 2.4) -FontSize 16 -Rows @(Get-AutoLabelRows $autoPol $autoRules $labelNames) -MaxRows 25
    }
    $encLabels = @($labelSet | Where-Object { $_.Encryption -ne 'None' -and $_.Encryption -notmatch '^Removes' })
    $markLabels = @($labelSet | Where-Object ContentMarking)
    $containerLabels = @($labelSet | Where-Object { $_.AppliesTo -match 'Sites|Groups' })
    if ($labelSet.Count) {
        try {
            Add-DocHeading 'Configuration details: label settings' 3
            Add-DocPara "What each label does when it is applied. $(Get-Plural $encLabels.Count 'label encrypts' 'labels encrypt') content, $(Get-Plural $markLabels.Count 'label adds' 'labels add') a visible header, footer or watermark, and $(Get-Plural $containerLabels.Count 'label applies' 'labels apply') to sites, groups and Teams."
            $rows = foreach ($l in ($labelSet | Sort-Object { [int]('0' + $_.Priority) })) {
                $nm = [string]$l.Label; if ($l.Parent) { $nm = '   - ' + $nm }
                if ((ConvertTo-CellValue $l.Disabled) -eq 'True') { $nm += ' (turned off)' }
                $site = [string]$l.SitesGroupsTeams; if (-not $site) { $site = $(if ($l.AppliesTo -match 'Sites|Groups') { 'Default settings' } else { 'Not applicable' }) }
                , @($nm, [string]$l.AppliesTo, [string]$l.Encryption, $(if ($l.ContentMarking) { [string]$l.ContentMarking } else { 'None' }), $site)
            }
            Add-DocTable -Headers @('Label', 'Applies to', 'Encryption', 'Content marking', 'Sites, groups and Teams') -Widths @(1.6, 1.3, 2.4, 1.8, 2.1) -FontSize 15 -Rows @($rows) -MaxRows 50 -Note 'Names of people and groups given permissions are not shown. Detail: 35_Sensitivity_Label_Settings.'
        }
        catch { Add-ReportPartError 'Label settings' $_ }
    }
    $cons = @()
    if ($labels.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'High' 'No sensitivity labels exist. Labels are a foundation for information protection; consider a simple taxonomy (four or five labels) before relying on label-based DLP.' }
    elseif ($labelRefs.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'Medium' 'Sensitivity labels exist but are not used in any DLP rule. Using them as conditions (for example blocking external sharing of Highly Confidential content) is a simple, explainable control.' }
    if ($labels.Count -gt 0 -and $labelPol.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'High' 'Labels exist but no label policy publishes them, so users cannot apply them.' }
    if ($labels.Count -gt 0 -and $autoPol.Count -eq 0 -and -not (Get-StepGap 'Auto-labelling policies')) { $cons += 'No auto-labelling policies are configured. Relying entirely on users to label content usually leads to inconsistent labelling.' }
    if ($labelSet.Count) {
        $hcNoEnc = @($labelSet | Where-Object { $_.Label -match 'Highly|Secret|Restricted' -and $_.Encryption -eq 'None' -and $_.AppliesTo -match 'Files|Emails' -and (ConvertTo-CellValue $_.Disabled) -ne 'True' })
        if ($hcNoEnc.Count) { $cons += Add-Obs 'Sensitivity labels' 'Medium' "$(Get-Plural $hcNoEnc.Count 'label for highly sensitive content does' 'labels for highly sensitive content do') not encrypt: $(($hcNoEnc.Label | Select-Object -First 5) -join ', '). Encryption keeps protection with a file when it is downloaded or forwarded; consider adding it where the people who need the content can still open it." }
        if ($containerLabels.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'Low' 'No label applies to sites, groups and Teams. Container labels set privacy, guest access and external sharing for Teams and SharePoint sites, so sensitive workspaces are protected as well as individual files.' }
        $anyDefault = @($labelPol | Where-Object { (@(Get-Val $_ 'Settings') -join ' ') -match '\[\s*(mandatory\s*,\s*true|defaultlabelid\s*,\s*[0-9a-f])' })
        if ($labelPol.Count -and $anyDefault.Count -eq 0) { $cons += Add-Obs 'Sensitivity labels' 'Low' 'No label policy sets a default label or requires users to label content. A default label (for example General) gives new documents and emails a baseline classification that DLP can use.' }
    }
    $Script:AreaStatus['Labels'] = $(if ($labelRefs.Count -gt 0 -and $labelPol.Count -gt 0) { 'G' } elseif ($labels.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('Labels', 'LabelEncryption', 'LabelContainers', 'AutoLabel', 'LabelCondition') -Area 'Labels'

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
    $epDetail = @($epRules | Group-Object { "$($_.Policy)|$($_.Rule)" } | ForEach-Object {
            $pr = $_.Name -split '\|', 2
            $acts = @($_.Group | ForEach-Object {
                    $nm = if ($Script:EndpointActionNames.ContainsKey([string]$_.Setting)) { $Script:EndpointActionNames[[string]$_.Setting] } else { ConvertFrom-CamelCase ([string]$_.Setting) }
                    "$($nm): **$(([string]$_.Value).ToLower())**" })
            , @($pr[0], $pr[1], ($acts -join '; '))
        })
    if ($epDetail.Count -gt 0) {
        Add-DocHeading 'Configuration details: device restrictions by rule' 3
        Add-DocTable -Headers @('Policy', 'Rule', 'What happens on the device') -Widths @(2, 2, 4.5) -FontSize 16 -Rows $epDetail -MaxRows 30
    }
    $setRows = @(Get-EndpointSettingRows $epSet)
    if ($setRows.Count -gt 0) {
        Add-DocHeading 'Configuration details: tenant-wide Endpoint DLP settings' 3
        Add-DocPara 'These settings apply to every device policy: which browsers, apps, websites, printers and USB devices are allowed or restricted, and which file paths are excluded from monitoring.'
        Add-DocTable -Headers @('Area', 'Setting', 'Value') -Widths @(1.8, 2.4, 4.3) -FontSize 16 -Rows $setRows -MaxRows 40 -Note 'Detail: 41_Endpoint_DLP_Settings.'
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
    $hyg = @()
    try {
        $hyg = @(Get-HygieneChecks -Policies $policies -PolSum $polSum -Rules $rules -RuleSum $ruleSum -ClassRows $rc -Coverage $cov -SitView $sitView -Labels $labels -LabelPolicies $labelPol -AlertPolicies $alertPol -CeCounts $ceCounts -Activity $ruleAct -ActivityCollected $actOn)
        Export-Dataset '82_Hygiene_Checks' $hyg -AlreadyFlat
        Add-DocHeading 'Configuration hygiene checks' 3
        Add-DocPara 'Automatic checks for common configuration problems. Green = no issue found, amber = worth reviewing, red = likely gap, grey = not checked.'
        $statusWord = @{ G = 'OK'; A = 'Review'; R = 'Issue'; N = 'Not checked' }
        Add-DocTable -Headers @('Check', 'Result', 'Detail') -Widths @(2.2, 0.9, 5.4) -FontSize 16 -Rows @(foreach ($h in $hyg) { , @([string]$h.Check, "[[$($h.Status)]]$($statusWord[[string]$h.Status])", $(if ($h.Advice) { "$($h.Detail) **Suggested:** $($h.Advice)" } else { [string]$h.Detail })) })
        foreach ($h in @($hyg | Where-Object { $_.Status -eq 'R' -and $_.Check -notin 'Stored sensitive data is covered by a rule', 'Rules are not routinely overridden' })) { [void](Add-Obs 'Hygiene' 'Medium' "$($h.Check): $($h.Detail) $($h.Advice)") }
    }
    catch { Add-ReportPartError 'Hygiene checks' $_ }
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
    if ($alertPol.Count -gt 0) {
        Add-DocHeading 'Configuration details: DLP alert policies' 3
        Add-DocPara 'Alert policies decide when a DLP alert is raised and who is emailed about it.'
        Add-DocTable -Headers @('Alert policy', 'Severity', 'State', 'Raised when', 'Email recipients') -Widths @(2.6, 0.9, 0.7, 1.8, 2.5) -FontSize 16 -Rows @(Get-AlertPolicyRows $alertPol) -MaxRows 30
    }
    $cons = @()
    try {
        if ($Script:Cache.ContainsKey('DlpAlertEvents')) {
            $al = @(if ($Script:Cache.ContainsKey('DlpAlerts')) { $Script:Cache['DlpAlerts'] })
            Add-DocHeading "DLP alerts raised (audit log, last $($Script:Settings.AuditDays) days)" 3
            if ($al.Count -eq 0) { Add-DocPara 'No DLP alerts were raised in the period.' }
            else {
                $openA = @($al | Where-Object { $_.Status -match 'Active|New|InProgress|Investigating|Unknown' }).Count
                Add-DocPara "**$($al.Count)** DLP alerts were raised: $(@($al | Where-Object Severity -eq 'High').Count) high, $(@($al | Where-Object Severity -eq 'Medium').Count) medium and $(@($al | Where-Object Severity -eq 'Low').Count) low severity; **$openA** still open. **DLP_Alerts_Dashboard.html** in the output folder shows what the alerts are about: the rules, information types, locations, people and external domains involved."
                Add-DocTable -Headers @('Rule raising alerts', 'Alerts', 'Most common information found', 'External recipients') -Widths @(3, 0.8, 3, 1.5) -FontSize 16 -Rows @(foreach ($g in ($al | Group-Object Rule | Sort-Object Count -Descending | Select-Object -First 10)) {
                        $ts = @($g.Group | ForEach-Object { $_.SensitiveTypes -split ', ' } | Where-Object { $_ } | Group-Object | Sort-Object Count -Descending | Select-Object -First 2 | ForEach-Object Name)
                        , @([string]$g.Name, [string]$g.Count, $(if ($ts.Count) { $ts -join ', ' } else { 'Not linked' }), [string]@($g.Group | Where-Object ExternalDomains).Count)
                    }) -Note 'Detail: 67_DLP_Alerts (one row per alert) and 66_DLP_Alert_Events.'
                if ($openA -gt 0 -and $openA -ge $al.Count / 2) { $cons += Add-Obs 'Alerts' 'Medium' "$openA of $($al.Count) DLP alerts in the period are still open. Agree who triages DLP alerts, how quickly, and how they are closed." }
            }
        }
    }
    catch { Add-ReportPartError 'DLP alerts raised' $_ }
    if ($ruleAlert.Count -eq 0 -and $rules.Count -gt 0) { $cons += Add-Obs 'Alerts' 'High' 'No DLP rules raise alerts, so nobody is notified of policy matches.' }
    $siteOnly = @($ruleAlert | Where-Object { ([string]$_.GenerateAlert + [string]$_.GenerateIncidentReport) -match '^(SiteAdmin\s*\|?\s*)+$' })
    if ($siteOnly.Count -gt 0) { $cons += Add-Obs 'Alerts' 'Medium' "$(Get-Plural $siteOnly.Count 'rule sends' 'rules send') alerts only to SiteAdmin. Route alerts to a named team or shared mailbox with a defined triage process." }
    $cons += 'Define alert severity, recipients, investigation steps and escalation routes, and who reviews override justifications.'
    $Script:AreaStatus['Alerts'] = $(if ($ruleAlert.Count -gt 0 -and $siteOnly.Count -eq 0) { 'G' } elseif ($ruleAlert.Count -gt 0) { 'A' } else { 'R' })
    Add-SectionEnd $cons @('Alerts') -Area 'Alerts'

    # ---------- 17 Activity Explorer ----------
    Add-DocHeading '17. Activity Explorer and DLP rule activity' 1
    Add-SectionIntro -What '**Activity Explorer** shows how sensitive and labelled information is being used: DLP rule matches, labels applied or changed, files copied to USB, printed, uploaded and more. It is the main source of evidence for policy tuning and investigations.' -Why 'Real activity data shows whether policies are working, where false positives occur, and which workloads and users generate the most risk.'
    $act = @(if ($Script:Cache.ContainsKey('Activity')) { $Script:Cache['Activity'] } else { @() })
    if (-not $Script:Cache.ContainsKey('Activity')) { Add-DocPara 'Activity Explorer data was not included in this run. Re-run the complete report with Activity Explorer enabled to add this section.' }
    else {
        Add-SectionGap "Activity Explorer ($($Script:Settings.ActivityDays) days)"
        Add-DocPara "$($act.Count) activity records were collected for the last $($Script:Settings.ActivityDays) days."
        $grp = @($act | Group-Object { '{0}|{1}' -f (Get-Val $_ 'Activity'), (Get-Val $_ 'Workload') } | Sort-Object Count -Descending)
        if ($grp.Count) { Add-DocTable -Headers @('Activity', 'Workload', 'Count') -Widths @(4, 2.5, 1) -Rows @(foreach ($g in $grp) { $p = $g.Name -split '\|', 2; , @($p[0], $p[1], [string]$g.Count) }) -MaxRows 25 -Note 'Detail: 60_Activity_Explorer_Records (contains user and file names).' }
    }
    $actCons = @()
    try {
        Add-DocHeading "DLP rule activity (audit log, last $($Script:Settings.AuditDays) days)" 3
        $actCons = @()
        if (-not $actOn) { Add-DocPara 'DLP rule activity was not collected in this run. Enable it in the complete report to show how often each rule matches, how often users override it and how often matches are reported as false positives. It reads the unified audit log through Exchange Online PowerShell (role: **View-Only Audit Logs** or **Audit Logs**).' }
        else {
            Add-SectionGap "DLP rule activity from the audit log ($($Script:Settings.AuditDays) days)"
            $tot = ($ruleAct | Measure-Object Matches -Sum).Sum; $ov = ($ruleAct | Measure-Object Overrides -Sum).Sum; $fpx = ($ruleAct | Measure-Object FalsePositives -Sum).Sum
            Add-DocPara "The audit log recorded **$('{0:N0}' -f [int64]$tot)** DLP rule matches in the period, with **$ov** overrides and **$fpx** false-positive reports."
            $rowsAct = foreach ($r in $ruleSum) {
                $a = $ruleAct | Where-Object { $_.Policy -eq $r.Policy -and $_.Rule -eq $r.Rule } | Select-Object -First 1
                if ($a) { , @([string]$r.Policy, [string]$r.Rule, ('{0:N0}' -f $a.Matches), [string]$a.Overrides, [string]$a.FalsePositives, [string]$a.DistinctUsers, [string]$a.Workloads, [string]$a.LastMatch) }
                else { , @([string]$r.Policy, [string]$r.Rule, '[[A]]0', '0', '0', '0', '', '') }
            }
            Add-DocTable -Headers @('Policy', 'Rule', 'Matches', 'Overrides', 'False positives', 'Users', 'Where', 'Last match') -Widths @(1.8, 2.2, 0.8, 0.8, 0.8, 0.7, 1.3, 1.1) -FontSize 15 -Rows @($rowsAct) -MaxRows 40 -Note 'User names are not shown here. Detail: 65_DLP_Rule_Activity_Summary; individual events (with users) in 64_DLP_Audit_Rule_Matches.'
            $silentN = @($rowsAct | Where-Object { $_[2] -eq '[[A]]0' }).Count
            if ($silentN) { $actCons += "$(Get-Plural $silentN 'rule had' 'rules had') no matches in the period. Test them with sample data to confirm they still detect what they should." }
            $heavyAct = @($ruleAct | Where-Object { $_.Matches -ge 10 -and ($_.Overrides + $_.FalsePositives) / [double]$_.Matches -ge 0.3 })
            if ($heavyAct.Count) { $actCons += Add-Obs 'Rule activity' 'Medium' "$(Get-Plural $heavyAct.Count 'rule is' 'rules are') overridden or reported as a false positive in 30% or more of matches ($(@($heavyAct | ForEach-Object Rule) -join '; ')). Review the justifications and tune the rule." }
            $other = @(Get-DlpAuditRecords | Where-Object { $_.Policy -and -not ($ruleSum | Where-Object Policy -eq $_.Policy) } | Select-Object -ExpandProperty Policy -Unique)
            if ($other.Count) { $actCons += "Matches were also recorded for policies that no longer exist or were not collected: $($other -join ', ')." }
        }
    }
    catch { Add-ReportPartError 'DLP rule activity' $_ }
    $Script:AreaStatus['Activity'] = $(if ($Script:Cache.ContainsKey('Activity') -or $actOn) { 'G' } else { 'N' })
    Add-SectionEnd (@($actCons) + @('Agree a regular review of Activity Explorer and rule activity (for example weekly during rollout, then monthly) and the operational reports stakeholders need.')) @('ActivityExplorer', 'AuditSearch') -Area 'Activity'

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
        if ($tr.Count -gt 0) {
            Add-DocHeading 'Configuration details: mail flow rules' 3
            Add-DocPara 'Rules are listed in priority order. The description is the summary Exchange Online generates for each rule.'
            Add-DocTable -Headers @('Mail flow rule', 'State', 'Inspects SITs', 'What it does') -Widths @(2, 1.3, 0.9, 4.3) -FontSize 16 -Rows @(Get-TransportRuleRows $tr) -MaxRows 40 -Note 'Detail: 70_Exchange_Transport_Rules.'
        }
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
    $uniq = @($roleMembers | Select-Object -ExpandProperty Member -Unique)
    if ($uniq.Count -gt 15) { $cons += Add-Obs 'Ownership' 'Low' "$($uniq.Count) people or groups hold DLP-relevant roles. Review whether all still need them." }
    $Script:AreaStatus['Ownership'] = $(if ($roleMembers.Count -gt 0) { 'G' } else { 'N' })
    Add-SectionEnd $cons @('Permissions') -Area 'Ownership'

    # ---------- Data lifecycle and records management ----------
    $retGaps = @()
    if ($lifeOn) {
        Add-DocHeading "$nLife. Data lifecycle and records management" 1
        Add-SectionIntro -What '**Retention policies** keep or delete content across whole locations, such as every mailbox or every SharePoint site. **Retention labels** do the same for individual items: they are published for people to apply, or applied automatically when content matches conditions.||**Records management** builds on retention labels: a label can declare an item a **record** (it cannot be edited or deleted while it is retained) or a **regulatory record** (stricter, and the label cannot be removed). **Disposition review** asks named people to confirm deletion at the end of the retention period, and **event-based retention** starts the clock from an event such as an employee leaving or a contract ending. **Adaptive scopes** target policies using attributes such as department or site type, so they stay up to date as people move.' -Why 'Keeping content for the right length of time meets legal and regulatory obligations, and deleting it when it is no longer needed reduces the information that could leak, that DLP has to protect and that Microsoft 365 Copilot can surface. Records management proves that important records have not been changed.'
        Add-SectionGap 'Retention policies', 'Retention rules', 'Retention labels', 'Adaptive scopes'
        $retPol = @(Get-RetentionPolicyView); $retLab = @(Get-RetentionLabelView); $scopesV = @(Get-AdaptiveScopeView)
        $plain = @($retPol | Where-Object Type -eq 'Retention policy'); $pubPol = @($retPol | Where-Object Type -eq 'Publishes retention labels'); $autoRet = @($retPol | Where-Object Type -eq 'Auto-applies a retention label')
        $recCount = @($retLab | Where-Object Kind -eq 'Record').Count; $regCount = @($retLab | Where-Object Kind -eq 'Regulatory record').Count
        $retPolGap = [bool](Get-StepGap 'Retention policies')
        Add-DocPara ("The tenant has **$(Get-Plural $plain.Count 'retention policy' 'retention policies')**, $(Get-Plural $pubPol.Count 'policy that publishes' 'policies that publish') retention labels and $(Get-Plural $autoRet.Count 'policy that auto-applies' 'policies that auto-apply') a retention label. " +
            "$(Get-Plural $retLab.Count 'retention label is' 'retention labels are') defined, of which $recCount declare records and $regCount regulatory records. $(Get-Plural $scopesV.Count 'adaptive scope is' 'adaptive scopes are') configured.")
        if (-not $retPolGap) {
            $covRows = foreach ($k in $Script:CoreRetentionLocations.Keys) {
                $on = @($plain | Where-Object { $_.Enabled -ne $false })
                $all = @($on | Where-Object { ($_.LocationAllKeys -split ',') -contains $k })
                $some = @($on | Where-Object { ($_.LocationKeys -split ',') -contains $k })
                if ($all.Count -eq 0 -and $some.Count -eq 0) { $retGaps += $Script:CoreRetentionLocations[$k] }
                , @($Script:CoreRetentionLocations[$k], [string]$some.Count, $(if ($all.Count) { '[[G]]All covered' } elseif ($some.Count) { '[[A]]Selected only' } else { '[[R]]No retention policy' }))
            }
            Add-DocHeading 'Retention coverage by location' 3
            Add-DocTable -Headers @('Location', 'Enabled retention policies', 'Coverage') -Widths @(3, 1.6, 2) -Rows @($covRows) -Note 'Retention labels are not counted here because they apply only to the items they are applied to. Adaptive policies count as selected only.'
        }
        if ($retPol.Count) {
            Add-DocHeading 'Configuration details: retention policies' 3
            $rows = foreach ($r in ($retPol | Sort-Object Type, Name)) {
                $st = @(if ($r.Enabled -eq $false) { 'Off' } else { 'On' })
                if ($r.DistributionStatus -match 'Error|Fail') { $st += 'distribution error' }
                if ($r.PreservationLock) { $st += 'Preservation Lock' }
                , @([string]$r.Name, [string]$r.Type, ($st -join ', '), [string]$r.Locations, [string]$r.Setting)
            }
            Add-DocTable -Headers @('Policy', 'Type', 'Status', 'Locations', 'Retention or labels') -Widths @(1.8, 1.4, 1.1, 2.3, 2.2) -FontSize 16 -Rows @($rows) -MaxRows 40 -Note 'Detail: 100_Retention_Policies and 102_Retention_Rules_Full.'
        }
        if ($retLab.Count) {
            Add-DocHeading 'Configuration details: retention labels' 3
            $rows = foreach ($t in ($retLab | Sort-Object Name)) {
                $ret = [string]$t.Retention; if ($t.EventType) { $ret += " (event: $($t.EventType))" }
                , @([string]$t.Name, [string]$t.Kind, $ret, $(if ($t.DispositionReview) { "Yes ($(Get-Plural $t.ReviewerCount 'reviewer' 'reviewers'))" } else { 'No' }), [string]$t.HowApplied)
            }
            Add-DocTable -Headers @('Retention label', 'Type', 'Retention', 'Disposition review', 'How applied') -Widths @(2, 1.3, 2.6, 1.3, 1.6) -FontSize 16 -Rows @($rows) -MaxRows 50 -Note 'Reviewer names are not shown. Detail: 103_Retention_Labels.'
        }
        if ($scopesV.Count) {
            Add-DocHeading 'Configuration details: adaptive scopes' 3
            Add-DocTable -Headers @('Adaptive scope', 'Type', 'Query', 'Used by') -Widths @(1.8, 0.9, 3.6, 2.2) -FontSize 16 -Rows @(foreach ($a in $scopesV) {
                    $q = [string]$a.Query; if ($q.Length -gt 180) { $q = $q.Substring(0, 180) + ' ...' }
                    , @([string]$a.Name, [string]$a.Type, $q, $(if ($a.UsedBy) { ($a.UsedBy -replace ' \| ', ', ') } else { 'Not used' }))
                }) -MaxRows 30 -Note 'Detail: 105_Adaptive_Scopes.'
        }
        $cons = @()
        if (-not $retPolGap) {
            if ($retPol.Count -eq 0 -and $retLab.Count -eq 0 -and -not (Get-StepGap 'Retention labels')) { $cons += Add-Obs 'Data lifecycle' 'High' 'No retention policies or retention labels exist. Content is kept or deleted only by users. Agree a retention schedule with legal and records owners and apply retention policies to the core locations.' }
            elseif ($retGaps.Count) { $cons += Add-Obs 'Data lifecycle' 'Medium' "No enabled retention policy covers: $($retGaps -join ', '). Content there is kept or deleted only by users. Agree a retention period and apply a retention policy (Teams chats and channels need their own policy)." }
            $err = @($retPol | Where-Object { $_.DistributionStatus -match 'Error|Fail' })
            if ($err.Count) { $cons += Add-Obs 'Data lifecycle' 'Medium' "$(Get-Plural $err.Count 'retention policy shows' 'retention policies show') a distribution error, so may not be applied everywhere: $(($err.Name | Select-Object -First 5) -join ', '). Retry distribution from the policy page." }
            $off = @($retPol | Where-Object { $_.Enabled -eq $false })
            if ($off.Count) { $cons += Add-Obs 'Data lifecycle' 'Low' "$(Get-Plural $off.Count 'retention policy is' 'retention policies are') turned off: $(($off.Name | Select-Object -First 5) -join ', '). Remove them or turn them on if still needed." }
            $lock = @($retPol | Where-Object PreservationLock)
            if ($lock.Count) { $cons += Add-Obs 'Data lifecycle' 'Info' "$(Get-Plural $lock.Count 'policy uses' 'policies use') Preservation Lock, so cannot be turned off or made less strict. Confirm this is intended." }
        }
        if ($retLab.Count -and -not (Get-StepGap 'Retention labels')) {
            $unusedRet = @($retLab | Where-Object HowApplied -eq 'Not in any policy')
            if ($unusedRet.Count) { $cons += Add-Obs 'Data lifecycle' 'Low' "$(Get-Plural $unusedRet.Count 'retention label is' 'retention labels are') not published or auto-applied, so cannot be used: $(($unusedRet.Name | Select-Object -First 6) -join ', ')." }
            $recNoRev = @($retLab | Where-Object { $_.Kind -ne 'Retention label' -and $_.RetentionAction -match 'Delete' -and -not $_.DispositionReview })
            if ($recNoRev.Count) { $cons += Add-Obs 'Data lifecycle' 'Low' "$(Get-Plural $recNoRev.Count 'record label deletes' 'record labels delete') content at the end of the retention period without a disposition review. Consider a disposition review so a person confirms that records can be deleted." }
            if (@($autoRet).Count -eq 0) { $cons += 'No retention label is applied automatically. Auto-applying labels to content that matches sensitive information types, keywords or trainable classifiers gives consistent retention without relying on users.' }
        }
        $cons += 'Retention that keeps content takes priority over deletion, and eDiscovery holds keep content regardless of retention settings. Check that deletion policies are not silently overridden by long-running holds.'
        $Script:AreaStatus['Lifecycle'] = $(if ($retPolGap -and (Get-StepGap 'Retention labels')) { 'N' } elseif ($retPol.Count -eq 0 -and $retLab.Count -eq 0) { 'R' } elseif ($retGaps.Count -eq 0 -and -not $retPolGap) { 'G' } else { 'A' })
        Add-SectionEnd $cons @('Retention', 'RecordsMgmt', 'AdaptiveScopes', 'Disposition') -Area 'Lifecycle'
        $cons = @()
    }

    # ---------- Insider risk, communication compliance and information barriers ----------
    if ($riskOn) {
        Add-DocHeading "$nRisk. Insider risk, communication compliance and information barriers" 1
        Add-SectionIntro -What '**Insider risk management** looks for patterns of risky activity, such as downloading large amounts of data before leaving the organisation, by combining signals from DLP alerts, the audit log, devices and HR systems. Users are pseudonymised by default.||**Communication compliance** detects inappropriate, regulated or sensitive content in email, Teams and Microsoft 365 Copilot interactions and sends it to named reviewers.||**Information barriers** stop groups of people (segments) from communicating or collaborating, for example to keep an advisory team apart from a trading team.' -Why 'DLP judges each action on its own. Insider risk management shows the pattern behind several actions and can use high-severity DLP alerts as the trigger for a Data leaks policy, so the two work best together. Communication compliance meets supervision obligations in regulated sectors and supports conduct policies, and information barriers enforce ethical walls.'
        Add-SectionGap 'Insider risk policies', 'Communication compliance policies', 'Communication compliance rules', 'Information barrier policies', 'Organisation segments'
        $irm = @(Get-InsiderRiskView); $cc = @(Get-CommComplianceView); $ib = @(Get-IbPolicyView); $seg = @(Get-OrgSegments)
        $irmGap = [bool](Get-StepGap 'Insider risk policies'); $ccGap = [bool](Get-StepGap 'Communication compliance policies'); $ibGap = [bool](Get-StepGap 'Information barrier policies')
        $hiAlert = @($ruleSum | Where-Object { ($_.GenerateAlert -or $_.GenerateIncidentReport) -and [string]$_.ReportSeverityLevel -eq 'High' })
        Add-DocHeading 'Insider risk management' 3
        if ($irmGap) { Add-DocPara 'Insider risk policies could not be read with the account used (see the note above).' }
        elseif ($irm.Count -eq 0) { Add-DocPara 'No insider risk policies were found.' }
        else {
            Add-DocPara "$(Get-Plural $irm.Count 'insider risk policy is' 'insider risk policies are') configured."
            Add-DocTable -Headers @('Policy', 'Template', 'State') -Widths @(3, 3, 1.2) -Rows @(foreach ($p in $irm) { , @([string]$p.Name, [string]$p.Template, [string]$p.State) }) -MaxRows 30 -Note 'Detail: 200_Insider_Risk_Policies.'
        }
        $leak = @($irm | Where-Object { $_.Template -match 'Data leaks' })
        if ($leak.Count) { Add-DocPara "Data leaks policies can use high-severity DLP alerts as their triggering event. $(Get-Plural $hiAlert.Count 'DLP rule raises' 'DLP rules raise') high-severity alerts." }
        Add-DocHeading 'Communication compliance' 3
        if ($ccGap) { Add-DocPara 'Communication compliance policies could not be read with the account used (see the note above).' }
        elseif ($cc.Count -eq 0) { Add-DocPara 'No communication compliance policies were found.' }
        else {
            Add-DocTable -Headers @('Policy', 'State', 'Reviewers', 'Share reviewed (%)', 'Detects sensitive info') -Widths @(3, 0.9, 1, 1.4, 1.5) -Rows @(foreach ($p in $cc) {
                    , @([string]$p.Name, [string]$p.State, [string]$p.Reviewers, $(if ($p.ReviewPercent) { [string]$p.ReviewPercent } else { '' }), $(if ($p.DetectsSensitiveInfo) { 'Yes' } else { 'No' }))
                }) -MaxRows 30 -Note 'Reviewer names are not shown. Detail: 210_Communication_Compliance_Policies.'
        }
        Add-DocHeading 'Information barriers' 3
        if ($ibGap) { Add-DocPara 'Information barrier policies could not be read with the account used (see the note above).' }
        elseif ($ib.Count -eq 0) { Add-DocPara "No information barrier policies were found. $(Get-Plural $seg.Count 'organisation segment is' 'organisation segments are') defined." }
        else {
            Add-DocPara "$(Get-Plural $ib.Count 'information barrier policy is' 'information barrier policies are') defined across $(Get-Plural $seg.Count 'segment' 'segments')."
            Add-DocTable -Headers @('Policy', 'Segment', 'Rule', 'State') -Widths @(2.2, 1.6, 3, 1) -FontSize 16 -Rows @(foreach ($p in $ib) { , @([string]$p.Name, [string]$p.Segment, [string]$p.Rule, [string]$p.State) }) -MaxRows 30 -Note 'Detail: 220_Information_Barrier_Policies and 221_Organisation_Segments.'
            $ibs = @(Get-IbStatus | Select-Object -First 1)
            if ($ibs.Count) { Add-DocPara "Last policy application: $([string](Get-Val $ibs[0] 'Status'))$(if (Get-Val $ibs[0] 'EndTime') { ', finished ' + [string](Get-Val $ibs[0] 'EndTime') })." }
        }
        $cons = @()
        if (-not $irmGap -and $irm.Count -eq 0) { $cons += Add-Obs 'Insider risk' 'Medium' 'No insider risk management policies exist. A Data leaks policy that uses high-severity DLP alerts as its trigger is a practical first step; agree it with HR, legal and privacy teams and keep users pseudonymised.' }
        if ($leak.Count -and $hiAlert.Count -eq 0) { $cons += Add-Obs 'Insider risk' 'Medium' 'A Data leaks insider risk policy exists but no DLP rule raises high-severity alerts, so the policy may never be triggered by DLP. Set high severity on the DLP rules that matter most, or choose a different triggering event.' }
        $irmOff = @($irm | Where-Object State -eq 'Off')
        if ($irmOff.Count) { $cons += Add-Obs 'Insider risk' 'Low' "$(Get-Plural $irmOff.Count 'insider risk policy is' 'insider risk policies are') turned off: $(($irmOff.Name | Select-Object -First 5) -join ', ')." }
        if (-not $ccGap -and $cc.Count -eq 0) { $cons += Add-Obs 'Communication compliance' 'Info' 'No communication compliance policies exist. Consider them if the organisation must supervise communications (for example regulated financial services) or wants to detect harassment, threats or sensitive information in messages and Copilot interactions.' }
        $ccNoRev = @($cc | Where-Object { $_.Reviewers -eq 0 })
        if ($ccNoRev.Count) { $cons += Add-Obs 'Communication compliance' 'Medium' "$(Get-Plural $ccNoRev.Count 'communication compliance policy has' 'communication compliance policies have') no reviewers, so matches are not reviewed: $(($ccNoRev.Name | Select-Object -First 5) -join ', ')." }
        $ibInactive = @($ib | Where-Object { $_.State -and $_.State -ne 'Active' })
        if ($ibInactive.Count) { $cons += Add-Obs 'Information barriers' 'Low' "$(Get-Plural $ibInactive.Count 'information barrier policy is' 'information barrier policies are') not active. Policies only take effect once they are set to active and applied." }
        $cons += 'Information barriers are only needed where groups must not communicate with each other; not having them is not a gap on its own.'
        $cons += 'Insider risk and communication compliance process personal data about employees. Involve HR, legal and privacy teams, document the purpose, and limit who can see the results.'
        $Script:AreaStatus['RiskCompliance'] = $(if ($irmGap -and $ccGap -and $ibGap) { 'N' } elseif (@($irm | Where-Object State -ne 'Off').Count) { $(if ($leak.Count -and $hiAlert.Count -eq 0) { 'A' } else { 'G' }) } elseif ($irm.Count -or $cc.Count -or $ib.Count) { 'A' } else { 'R' })
        Add-SectionEnd $cons @('InsiderRisk', 'CommCompliance', 'InfoBarriers') -Area 'RiskCompliance'
        $cons = @()
    }

    # ---------- Audit and eDiscovery ----------
    $dlpAuditRet = $true; $auditRetGap = $true
    if ($audOn) {
        Add-DocHeading "$nAudit. Audit and eDiscovery" 1
        Add-SectionIntro -What 'The **unified audit log** records user and admin activity across Microsoft 365, including every DLP rule match. **Audit (Standard)** keeps records for 180 days. **Audit (Premium)** keeps Exchange, SharePoint, OneDrive and Microsoft Entra records for one year and adds **audit log retention policies**, which keep chosen activities for up to ten years.||**eDiscovery cases** are used to find and preserve content for legal matters and investigations. A **hold** in a case keeps content in mailboxes and sites, even if users delete it or a retention policy would delete it.' -Why 'Investigating a DLP incident depends on the audit trail being switched on and kept for long enough. eDiscovery holds preserve evidence, but holds that are no longer needed keep content longer than the retention schedule intends.'
        Add-SectionGap 'Audit log retention policies', 'Unified audit log configuration', 'eDiscovery cases', 'eDiscovery case holds'
        $audRet = @(Get-AuditRetentionView); $cases = @(Get-EDiscoveryCaseView); $holds = @(Get-CaseHolds)
        $auditRetGap = [bool](Get-StepGap 'Audit log retention policies'); $casesGap = [bool](Get-StepGap 'eDiscovery cases')
        Add-DocHeading 'Audit' 3
        $aud = @($Script:Cache['Audit'])
        if ($aud.Count) { Add-DocPara ("Unified audit log recording: **{0}**." -f $(if ($aud[0].UnifiedAuditLogIngestionEnabled) { 'On' } else { 'Off' })) }
        else { Add-DocPara 'The unified audit log status was not read in this run (it is read over Exchange Online).' }
        $dlpAuditRet = [bool]@($audRet | Where-Object CoversDlp).Count
        if (-not $auditRetGap) {
            if ($audRet.Count) {
                Add-DocPara "$(Get-Plural $audRet.Count 'audit log retention policy is' 'audit log retention policies are') configured$(if ($dlpAuditRet) { ', including DLP activity' } else { '; none covers DLP activity' })."
                Add-DocTable -Headers @('Policy', 'Priority', 'Applies to', 'Keeps for') -Widths @(2.2, 0.8, 4, 1.2) -FontSize 16 -Rows @(foreach ($p in $audRet) { , @([string]$p.Name, [string]$p.Priority, [string]$p.AppliesTo, [string]$p.Retention) }) -MaxRows 30 -Note 'User names are not shown. Detail: 300_Audit_Retention_Policies.'
            }
            else { Add-DocPara 'No audit log retention policies exist, so the default retention applies: 180 days with Audit (Standard), or one year for Exchange, SharePoint, OneDrive and Microsoft Entra activity with Audit (Premium).' }
        }
        Add-DocHeading 'eDiscovery cases and holds' 3
        $activeCases = @($cases | Where-Object Status -ne 'Closed'); $closedCases = @($cases | Where-Object Status -eq 'Closed')
        if ($casesGap) { Add-DocPara 'eDiscovery cases could not be read with the account used (see the note above).' }
        elseif ($cases.Count -eq 0) { Add-DocPara 'No eDiscovery cases are visible to the account used. Only cases the account is a member of are listed, unless it is an eDiscovery Administrator.' }
        else {
            Add-DocPara "$(Get-Plural $cases.Count 'eDiscovery case is' 'eDiscovery cases are') visible to the account used: $($activeCases.Count) active and $($closedCases.Count) closed. Case names are not shown in this report; they are in 310_eDiscovery_Cases."
            $types = @($cases | Select-Object -ExpandProperty Type -Unique | Sort-Object)
            Add-DocTable -Headers @('Case type', 'Active', 'Closed', 'Holds on active cases', 'Mailboxes on hold', 'Sites on hold') -Widths @(1.6, 0.9, 0.9, 1.5, 1.4, 1.2) -Rows @(foreach ($t in $types) {
                    $h = @($holds | Where-Object { $_.CaseType -eq $t -and $_.Enabled -ne $false })
                    , @($t, [string]@($activeCases | Where-Object Type -eq $t).Count, [string]@($closedCases | Where-Object Type -eq $t).Count, [string]$h.Count, [string](($h | Measure-Object Mailboxes -Sum).Sum + 0), [string](($h | Measure-Object Sites -Sum).Sum + 0))
                })
            $qb = @($holds | Where-Object QueryBased)
            if ($holds.Count) { Add-DocPara "$(Get-Plural $qb.Count 'hold is' 'holds are') query-based (only matching content is kept); the rest keep all content in the locations." }
        }
        $cons = @()
        if ($aud.Count -and -not $aud[0].UnifiedAuditLogIngestionEnabled -and -not $Script:Cache.ContainsKey('TransportRules')) { $cons += Add-Obs 'Auditing' 'High' 'The unified audit log is disabled. DLP reporting, Activity Explorer and investigations depend on it.' }
        if (-not $auditRetGap -and -not $dlpAuditRet) { $cons += Add-Obs 'Audit' 'Low' 'No audit log retention policy keeps DLP activity beyond the default retention. Investigations and trend reporting often need DLP rule matches for longer; consider a policy for the DLP record types (needs Audit (Premium)).' }
        $oldCases = @($activeCases | Where-Object { $d = ConvertTo-DateSafe $_.Created; $d -and $d -lt (Get-Date).ToUniversalTime().AddYears(-2) })
        if ($oldCases.Count) { $cons += Add-Obs 'eDiscovery' 'Low' "$(Get-Plural $oldCases.Count 'active eDiscovery case was' 'active eDiscovery cases were') opened more than two years ago. Review with legal whether the cases and their holds are still needed; holds keep content regardless of retention policies." }
        $cons += 'Limit eDiscovery Administrator and eDiscovery Manager membership to the people who need it: they can search and export content from any mailbox and site.'
        $Script:AreaStatus['AuditDiscovery'] = $(if ($aud.Count -and -not $aud[0].UnifiedAuditLogIngestionEnabled) { 'R' } elseif ($auditRetGap -and $casesGap -and -not $aud.Count) { 'N' } elseif ($aud.Count -and $dlpAuditRet) { 'G' } else { 'A' })
        Add-SectionEnd $cons @('AuditSolutions', 'AuditRetention', 'EDiscovery') -Area 'AuditDiscovery'
        $cons = @()
    }

    # ---------- Next steps ----------
    Add-DocHeading "$nNext. Recommended next steps" 1
    $licShort = @{}
    try {
        Add-DocHeading 'Prioritised action plan' 2
        Export-ActionPlan
        $plan = @($Script:Cache['ActionPlan'])
        $licShort = @{}
        foreach ($k in $Script:LicenceText.Keys) { $licShort[$Script:LicenceText[$k]] = $(switch ($k) { 'Core' { 'E3 / Business Premium' } 'None' { 'None' } 'Labels' { 'E3; auto-labelling E5*' } 'Alerts' { 'E3; volume alerts E5*' } 'Teams' { 'E5*' } 'Copilot' { 'E5* + Copilot' } 'Lifecycle' { 'E3; advanced E5*' } 'Risk' { 'E5 IRM*' } 'AuditPrem' { 'E3; premium E5 eD&A*' } default { 'E5*' } }) }
        if ($plan.Count -eq 0) { Add-DocPara 'No actions were generated automatically. Use the topic table below to plan the engagement.' }
        else {
            Add-DocPara 'Actions generated from the observations in this report, highest priority first. **Effort** is a rough guide: Small = days, Medium = one to three weeks, Large = one to three months including testing and communication. **Licence** shows what the organisation needs for the action. The same list, with Owner, Target date and Status columns, is in **80_Action_Plan** (CSV and Excel) for tracking.'
            $prioCell = @{ High = '[[R]]High'; Medium = '[[A]]Medium'; Low = '[[N]]Low'; Info = '[[N]]Info' }
            Add-DocTable -Headers @('#', 'Priority', 'Action', 'Effort', 'Licence') -Widths @(0.4, 0.9, 5.2, 0.8, 1.3) -FontSize 16 -Rows @(foreach ($a in $plan) { , @([string]$a.Id, $prioCell[[string]$a.Priority], [string]$a.Action, [string]$a.Effort, $(if ($licShort.ContainsKey([string]$a.Licence)) { $licShort[[string]$a.Licence] } else { [string]$a.Licence })) }) -MaxRows 40 -Note 'E5* = Microsoft 365 E5, E5 Information Protection & Governance, or Microsoft Purview Suite (formerly E5 Compliance). E5 IRM* = Microsoft 365 E5, E5 Insider Risk Management, or Microsoft Purview Suite. E5 eD&A* = Microsoft 365 E5, E5 eDiscovery & Audit, or Microsoft Purview Suite.'
        }
    }
    catch { Add-ReportPartError 'Action plan' $_ }
    Add-DocHeading 'Licensing notes' 2
    Add-DocPara 'Core DLP for Exchange, SharePoint and OneDrive, and manual sensitivity labelling, are included in Microsoft 365 E3 and Business Premium. DLP for Teams chat, Endpoint DLP, DLP for Microsoft 365 Copilot, automatic labelling, Exact Data Match, trainable classifiers, Activity Explorer and Content Explorer need Microsoft 365 E5, the E5 Information Protection & Governance add-on, or the Microsoft Purview Suite (formerly E5 Compliance). Licensing changes over time: confirm against Microsoft''s current guidance and the organisation''s agreement before planning purchases.'
    if ($lifeOn -or $riskOn -or $audOn) {
        $licLines = @()
        if ($lifeOn) { $licLines += '**Data lifecycle and records management:** ' + $Script:LicenceText['Lifecycle'] + '.' }
        if ($riskOn) { $licLines += '**Insider risk management, communication compliance and information barriers:** ' + $Script:LicenceText['Risk'] + '.' }
        if ($audOn) { $licLines += '**Audit and eDiscovery:** ' + $Script:LicenceText['AuditPrem'] + '.' }
        foreach ($x in $licLines) { Add-DocBullet $x }
    }
    $l = $Script:LearnLinks['Licensing']; Add-DocLink -Title $l.Title -Url $l.Url -AsBullet
    Add-DocHeading 'Activities by topic' 2
    Add-DocPara 'Suggested activities for each DLP topic, with the automated finding from this assessment and the licence it needs. The same content, with Status, Owner and Target date columns, is in 00_Assessment_Summary.csv.'
    if ($summaryRows.Count) { Add-DocTable -Headers @('Topic', 'Finding', 'Suggested activity', 'Licence') -Widths @(1.7, 3.2, 2.6, 1.3) -FontSize 15 -Rows @(foreach ($s in $summaryRows) { , @([string]$s.Topic, [string]$s.Finding, [string]$s.SuggestedActivity, $(if ($licShort.ContainsKey([string]$s.LicenceNeeded)) { $licShort[[string]$s.LicenceNeeded] } else { [string]$s.LicenceNeeded })) }) }

    # ---------- 21 Applying the configuration ----------
    $auditOff = [bool]($Script:Cache.ContainsKey('Audit') -and @($Script:Cache['Audit']).Count -and -not @($Script:Cache['Audit'])[0].UnifiedAuditLogIngestionEnabled)
    $ctx = @{
        Gaps = @($gaps); Policies = $policies.Count; Test = $pTest; AuditOnly = $respCount['Audit only']; Rules = $rules.Count; Tips = $tips.Count
        Labels = $labels.Count; LabelPolicies = $labelPol.Count; LabelRefs = $labelRefs.Count; EndpointCov = $covCount['Endpoint']; Usb = $usb.Count
        CopilotCov = $covCount['Copilot']; AlertRules = $ruleAlert.Count; SiteOnly = $siteOnly.Count; LowConf = $lowConf.Count; Custom = $custom.Count
        Edm = $edm.Count; AuditOff = $auditOff
        RetentionGaps = @($retGaps); LifecycleOn = $lifeOn; IrmNone = [bool]($riskOn -and -not (Get-StepGap 'Insider risk policies') -and @(Get-InsiderRiskView).Count -eq 0)
        AuditRetNoDlp = [bool]($audOn -and -not $auditRetGap -and -not $dlpAuditRet)
    }
    Add-ImplementationGuidance -Ctx $ctx -Section $nApply

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
        , @('Retention and records included', [string]$lifeOn)
        , @('Risk and compliance included', [string]$riskOn)
        , @('Audit and eDiscovery included', [string]$audOn)
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
        , @('Retention policy', 'Keeps or deletes content across a whole location, such as all mailboxes or all sites.')
        , @('Retention label', 'Keeps or deletes an individual item; can also declare the item a record.')
        , @('Record', 'An item that cannot be edited or deleted while its retention label applies. A regulatory record is stricter still.')
        , @('Disposition review', 'People confirm that content can be deleted at the end of its retention period.')
        , @('Adaptive scope', 'A query (for example department or site type) that decides which users, groups or sites a policy applies to.')
        , @('Insider risk management', 'Finds patterns of risky activity by combining signals such as DLP alerts, downloads and HR events.')
        , @('Communication compliance', 'Detects inappropriate, regulated or sensitive content in messages for review.')
        , @('Information barrier', 'Stops defined groups of people (segments) from communicating or collaborating.')
        , @('Audit log retention policy', 'Keeps chosen audit records for longer than the default (Audit Premium).')
        , @('eDiscovery hold', 'Preserves content in mailboxes and sites for a legal matter or investigation.')
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
        , @('62 / 63_Content_Explorer...', 'Items containing each sensitive information type and label, and top locations')
        , @('64 / 65_DLP_Audit... / Rule_Activity...', 'DLP rule matches, overrides and false positives from the audit log (64 contains user names)')
        , @('66 / 67_DLP_Alert... and DLP_Alerts_Dashboard.html', 'DLP alerts from the audit log, linked to the rule matches behind them, and the alerts dashboard')
        , @('35_Sensitivity_Label_Settings', 'Encryption, content marking and site and group settings for each label')
        , @('100-106_Retention...', 'Retention policies and rules, retention labels, adaptive scopes (when included)')
        , @('200-222_...', 'Insider risk, communication compliance and information barrier policies and segments (when included)')
        , @('300 / 310 / 311...', 'Audit log retention policies, eDiscovery cases and holds (when included; 310 contains case names)')
        , @('80_Action_Plan', 'Prioritised action plan with effort, licence, owner and target date columns')
        , @('81-84', 'Maturity scores, hygiene checks, changes since the last report and the coverage map')
        , @('99_Collection_Errors', 'Steps skipped or failed')
    )

    Add-FrameworkAppendix

    Add-PolicyDetailAppendix -Policies $policies -Rules $rules -ClassRows $rc -LabelNames $labelNames

    # ---------- insert executive observations ----------
    $obsXml = [System.Text.StringBuilder]::new()
    $saved = $Script:Doc.Body; $Script:Doc.Body = $obsXml
    try {
        $auditOn = $false
        if ($Script:Cache.ContainsKey('Audit') -and @($Script:Cache['Audit']).Count) { $auditOn = [bool]@($Script:Cache['Audit'])[0].UnifiedAuditLogIngestionEnabled }
        $mat = @(Get-MaturityScores -Policies $policies -RuleSum $ruleSum -ClassRows $rc -CovCount $covCount -SitView $sitView -Edm $edm -Labels $labels -LabelPolicies $labelPol -AutoPolicies $autoPol -AlertPolicies $alertPol -RoleMembers $roleMembers -Hygiene $hyg -AuditOn $auditOn -EvidenceCollected ($actOn -or $ceOn))
        $overall = [Math]::Round((($mat | Measure-Object Score -Average).Average), 1)
        Export-Dataset '81_Maturity_Score' (@($mat) + @([pscustomobject][ordered]@{ Dimension = 'Overall'; Score = $overall; Evidence = (Get-MaturityLevel $overall); NextStep = '' })) -AlreadyFlat
        Add-DocHeading 'DLP maturity' 2
        Add-DocPara "Overall score: **$overall out of 5 ($(Get-MaturityLevel $overall))**. Each dimension is scored from 0 (not in place) to 5 (optimised) using the collected configuration. Use the scores to track progress between assessments rather than as an absolute grade."
        $prevMat = @(); if ($Script:CompareFolder) { $prevMat = @(Read-RunFile $Script:CompareFolder '81_Maturity_Score') }
        $matRows = foreach ($m in $mat) {
            $cell = if ($m.Score -ge 4) { '[[G]]' } elseif ($m.Score -ge 2) { '[[A]]' } else { '[[R]]' }
            $pm = $prevMat | Where-Object Dimension -eq $m.Dimension | Select-Object -First 1
            $delta = if ($pm) { $d = [int]$m.Score - [int]$pm.Score; if ($d -gt 0) { " (+$d)" } elseif ($d -lt 0) { " ($d)" } else { ' (=)' } } else { '' }
            , @([string]$m.Dimension, "$cell$($m.Score) / 5$delta", [string]$m.Evidence, [string]$m.NextStep)
        }
        Add-DocTable -Headers @('Dimension', 'Score', 'Based on', 'To improve') -Widths @(1.5, 1.3, 3.4, 2.6) -FontSize 16 -Rows @($matRows) -Note $(if ($prevMat.Count) { 'Brackets show the change since the previous assessment.' } else { '' })
        if ($Script:CompareFolder) {
            $prevInfo = @(Read-RunFile $Script:CompareFolder '01_Tenant_Info')
            $prevWhen = if ($prevInfo.Count -and $prevInfo[0].CollectedAtUtc) { try { ([datetime]$prevInfo[0].CollectedAtUtc).ToString('dd MMMM yyyy') } catch { [string]$prevInfo[0].CollectedAtUtc } } else { Split-Path $Script:CompareFolder -Leaf }
            Add-DocHeading "Changes since the last assessment ($prevWhen)" 2
            if ($prevInfo.Count -and $tid -and $prevInfo[0].TenantId -and $prevInfo[0].TenantId -ne $tid) { Add-DocCallout 'Different tenant:' 'The previous report is for a different tenant, so no comparison was made.' 'FFF4E5' 'E08A00' }
            else {
                $changes = @(Get-ChangesSinceLastRun -Folder $Script:CompareFolder -Policies $policies -Rules $rules -Labels $labels -LabelPolicies $labelPol -AlertPolicies $alertPol)
                Export-Dataset '83_Changes_Since_Last_Run' $changes -AlreadyFlat
                if ($changes.Count -eq 0) { Add-DocPara 'No changes were found to DLP policies, rules, sensitivity labels, label policies or DLP alert policies.' }
                else {
                    $pm0 = $prevMat | Where-Object Dimension -eq 'Overall' | Select-Object -First 1
                    if ($pm0) { Add-DocPara "Overall maturity has moved from **$($pm0.Score)** to **$overall**." }
                    Add-DocTable -Headers @('Item', 'Name', 'Change') -Widths @(1.5, 4, 3) -FontSize 16 -Rows @(foreach ($c in $changes) { , @([string]$c.Item, [string]$c.Name, $(switch -Regex ([string]$c.Change) { '^Added' { '[[G]]Added' } '^Removed' { '[[R]]Removed' } default { [string]$c.Change } })) }) -MaxRows 40 -Note "Compared with: $(Split-Path $Script:CompareFolder -Leaf). Detail: 83_Changes_Since_Last_Run."
                }
            }
        }
    }
    catch { Add-ReportPartError 'Maturity and changes' $_ }
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
    $stWord = @{ G = 'In place'; A = 'Partly in place'; R = 'Not in place'; N = 'Not assessed' }
    if ($lifeOn) { $statusRows += , @('Retention and records', $stWord[$(if ($Script:AreaStatus.ContainsKey('Lifecycle')) { $Script:AreaStatus['Lifecycle'] } else { 'N' })]) }
    if ($riskOn) { $statusRows += , @('Insider risk and communication compliance', $stWord[$(if ($Script:AreaStatus.ContainsKey('RiskCompliance')) { $Script:AreaStatus['RiskCompliance'] } else { 'N' })]) }
    if ($audOn) { $statusRows += , @('Audit and eDiscovery', $stWord[$(if ($Script:AreaStatus.ContainsKey('AuditDiscovery')) { $Script:AreaStatus['AuditDiscovery'] } else { 'N' })]) }
    Add-DocTable -Headers @('Area', 'Status') -Widths @(3, 3) -Rows $statusRows
    if (@(Get-SelectedFrameworks).Count -gt 0) {
        Add-DocHeading 'Framework alignment at a glance' 2
        Add-DocPara 'How each area of the DLP configuration supports recognised security frameworks. Each area section repeats its row with a plain-English description of what the framework expects; Appendix D explains the frameworks and status values.'
        Add-FrameworkMatrix
        Add-DocPara "Indicative mapping, not a compliance assessment. Section $nApply gives guidance for closing the gaps." -Size 16 -Color '666666' -Italic
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
        $lk = if ($Script:TopicLicence.ContainsKey($Topic)) { $Script:TopicLicence[$Topic] } else { 'Core' }
        [pscustomobject][ordered]@{ Category = $Category; Topic = $Topic; Finding = $Finding; EvidenceFile = $Evidence; SuggestedActivity = $Activity
            LicenceNeeded = $Script:LicenceText[$lk]; Status = 'Not Started'; Owner = ''; TargetDate = ''; Notes = ''; LearnMore = "https://learn.microsoft.com/en-us/purview/$Learn" } }

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
        New-Row 'Monitoring' 'Content Explorer' ($(if ($Script:Cache.ContainsKey('ContentExplorer')) { $ce = @($Script:Cache['ContentExplorer'] | Where-Object { $_.TagType -eq 'SensitiveInformationType' -and $_.ItemCount -gt 0 }); "$(@($ce | Select-Object -ExpandProperty TagName -Unique).Count) sensitive information types found in stored content; $(@($ce | Where-Object { -not $_.UsedInDlpRules } | Select-Object -ExpandProperty TagName -Unique).Count) of them are not used in any DLP rule." } else { 'Not collected (enable Content Explorer in Settings or use -IncludeContentExplorer).' }) + (Get-StepGap 'Content Explorer item counts')) '62_Content_Explorer_Counts' 'Confirm where sensitive information is stored and whether it is protected.' 'data-classification-content-explorer'
        New-Row 'Monitoring' 'DLP Rule Activity' ($(if ($Script:Cache.ContainsKey('DlpAudit')) { $ra = @(Get-DlpRuleActivity); "$(($ra | Measure-Object Matches -Sum).Sum) rule matches, $(($ra | Measure-Object Overrides -Sum).Sum) overrides and $(($ra | Measure-Object FalsePositives -Sum).Sum) false-positive reports in the last $($Script:Settings.AuditDays) days." } else { 'Not collected (enable DLP rule activity in Settings or use -IncludeDlpAudit).' }) + (Get-StepGap "DLP rule activity from the audit log ($($Script:Settings.AuditDays) days)")) '65_DLP_Rule_Activity_Summary' 'Review rule matches, overrides and false positives regularly and tune rules.' 'audit-search'
        New-Row 'Monitoring' 'DLP Alerts' ("$alerting rules generate alerts or incident reports; $(@(Get-DlpAlertPolicies).Count) DLP-related alert policies." + (Get-StepGap 'Alert policies')) '51_Alert_Policies_DLP / 52_DLP_Rule_Alerting' 'Define alert severity, recipients, investigation steps and escalation routes.' 'dlp-alerts-dashboard-get-started'
        New-Row 'AI and Copilot' 'DLP and Microsoft 365 Copilot' "$($c['Copilot']) policies scoped to Microsoft 365 Copilot / Copilot Chat." '12_DLP_Workload_Coverage' 'Confirm whether Copilot and generative-AI data-loss scenarios are in scope.' 'dlp-microsoft365-copilot-location-learn-about'
        New-Row 'Governance' 'Operational Ownership' ("$(@($roleMembers | Select-Object -ExpandProperty Member -Unique).Count) people/groups hold DLP-relevant Purview roles across $(@($roleMembers | Select-Object -ExpandProperty RoleGroup -Unique).Count) role groups." + (Get-StepGap 'Purview role groups', 'Purview role group members')) '55_DLP_Relevant_Role_Members' 'Create a responsibility matrix covering security, compliance, IT, legal, HR and business owners.' 'purview-permissions'
        New-Row 'Delivery approach' 'Phased DLP Deployment' "Current state - enforced: $pEnforce, test: $pTest, disabled: $pOff." '11_DLP_Policies_Summary' 'Agree policy waves, pilot groups, acceptance criteria and enforcement dates.' 'dlp-create-deploy-policy'
    )
    $ls = @(Get-LabelSettingsView)
    $rows += New-Row 'Classification' 'Sensitivity Label Settings' ("$(Count-Where $ls { $_.Encryption -ne 'None' -and $_.Encryption -notmatch '^Removes' }) labels encrypt, $(Count-Where $ls { $_.ContentMarking }) add content marking, $(Count-Where $ls { $_.AppliesTo -match 'Sites|Groups' }) apply to sites, groups and Teams." + (Get-StepGap 'Sensitivity labels')) '35_Sensitivity_Label_Settings' 'Confirm that encryption, marking and container settings match the handling rules for each classification.' 'encryption-sensitivity-labels'
    if ($Script:Settings.IncludeLifecycle) {
        $rp = @(Get-RetentionPolicyView); $rl = @(Get-RetentionLabelView)
        $rows += New-Row 'Data lifecycle' 'Retention Policies' ("$(Count-Where $rp { $_.Type -eq 'Retention policy' }) retention policies; $(Count-Where $rp { $_.Type -ne 'Retention policy' }) label policies (publish or auto-apply); $(@(Get-AdaptiveScopeView).Count) adaptive scopes." + (Get-StepGap 'Retention policies', 'Retention rules', 'Adaptive scopes')) '100_Retention_Policies' 'Agree a retention schedule and apply retention policies to email, SharePoint, OneDrive and Teams.' 'retention'
        $rows += New-Row 'Data lifecycle' 'Retention Labels and Records' ("$($rl.Count) retention labels; $(Count-Where $rl { $_.Kind -ne 'Retention label' }) declare records; $(Count-Where $rl { $_.DispositionReview }) use disposition review; $(Count-Where $rl { $_.HowApplied -eq 'Not in any policy' }) not in any policy." + (Get-StepGap 'Retention labels')) '103_Retention_Labels' 'Identify records that must be protected from change and the people who approve their deletion.' 'records-management'
    }
    if ($Script:Settings.IncludeRiskCompliance) {
        $rows += New-Row 'Risk and compliance' 'Insider Risk Management' ("$(@(Get-InsiderRiskView).Count) insider risk policies." + (Get-StepGap 'Insider risk policies')) '200_Insider_Risk_Policies' 'Agree insider risk scenarios with HR, legal and privacy, starting with a Data leaks policy triggered by DLP alerts.' 'insider-risk-management'
        $rows += New-Row 'Risk and compliance' 'Communication Compliance' ("$(@(Get-CommComplianceView).Count) communication compliance policies." + (Get-StepGap 'Communication compliance policies')) '210_Communication_Compliance_Policies' 'Confirm whether communications must be supervised and who reviews them.' 'communication-compliance'
        $rows += New-Row 'Risk and compliance' 'Information Barriers' ("$(@(Get-IbPolicyView).Count) information barrier policies; $(@(Get-OrgSegments).Count) segments." + (Get-StepGap 'Information barrier policies')) '220_Information_Barrier_Policies' 'Confirm whether any groups must be kept apart (ethical walls).' 'information-barriers'
    }
    if ($Script:Settings.IncludeAuditDiscovery) {
        $ar = @(Get-AuditRetentionView); $ec = @(Get-EDiscoveryCaseView)
        $audTxt = if ($Script:Cache.ContainsKey('Audit') -and @($Script:Cache['Audit']).Count) { " Unified audit log: $(if (@($Script:Cache['Audit'])[0].UnifiedAuditLogIngestionEnabled) { 'on' } else { 'OFF' })." } else { '' }
        $rows += New-Row 'Audit and eDiscovery' 'Audit Log Retention' ("$($ar.Count) audit log retention policies; $(Count-Where $ar { $_.CoversDlp }) cover DLP activity.$audTxt" + (Get-StepGap 'Audit log retention policies', 'Unified audit log configuration')) '300_Audit_Retention_Policies / 71_Audit_Configuration' 'Agree how long DLP and other audit records must be kept for investigations.' 'audit-log-retention-policies'
        $rows += New-Row 'Audit and eDiscovery' 'eDiscovery Cases and Holds' ("$($ec.Count) eDiscovery cases visible ($(Count-Where $ec { $_.Status -ne 'Closed' }) active); $(@(Get-CaseHolds).Count) holds on active cases." + (Get-StepGap 'eDiscovery cases', 'eDiscovery case holds')) '310_eDiscovery_Cases / 311_eDiscovery_Case_Holds' 'Review active cases and holds with legal; close those no longer needed.' 'edisc'
    }
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
    # Exchange Online areas last: connecting to Exchange changes the loaded cmdlets
    $exo = @($keys | Where-Object { $_ -in '12', '14' })
    if ($exo -contains '14') { [void](Get-Rules) }   # rule names are needed to link alerts to rules
    $keys = @($keys | Where-Object { $_ -notin '12', '14' })
    foreach ($k in $keys) {
        Write-Host ''
        Write-Host "--- $($Script:Areas[$k]) ---" -ForegroundColor Cyan
        & $Script:AreaFunctions[$k]
    }
    Invoke-ExoPhase -Exchange ($exo -contains '12') -DlpAudit ($exo -contains '14') -AuditConfig ($keys -contains '17')
    Complete-Run
}

function Invoke-FullReport {
    if (-not (Connect-Tenant)) { return }
    Start-Run 'Full'
    $Script:CompareFolder = ''
    $c = Get-IppsConnection
    $tidNow = if ($c) { [string]$c.TenantID } else { '' }
    if ($Script:Settings.CompareWith -eq 'Latest') { $Script:CompareFolder = Find-PreviousRun -TenantId $tidNow }
    elseif ($Script:Settings.CompareWith) { $Script:CompareFolder = $Script:Settings.CompareWith }
    elseif ($Script:AskCompare -and $Script:Settings.WordReport) {
        $prev = Find-PreviousRun -TenantId $tidNow
        if ($prev -and (Read-YesNo "  Show changes since the previous report ($(Split-Path $prev -Leaf))?" $true)) { $Script:CompareFolder = $prev }
    }
    if ($Script:CompareFolder) { Write-Log "Comparing with previous report: $Script:CompareFolder" }
    $keys = @('1', '2', '3', '4', '5', '6', '7', '8', '9', '10')
    if ($Script:Settings.IncludeActivity) { $keys += '11' }
    if ($Script:Settings.IncludeContentExplorer) { $keys += '13' }
    if ($Script:Settings.IncludeLifecycle) { $keys += '15' }
    if ($Script:Settings.IncludeRiskCompliance) { $keys += '16' }
    if ($Script:Settings.IncludeAuditDiscovery) { $keys += '17' }
    foreach ($k in $keys) {
        Write-Host ''
        Write-Host "--- $($Script:Areas[$k]) ---" -ForegroundColor Cyan
        & $Script:AreaFunctions[$k]
    }
    Invoke-ExoPhase -Exchange $Script:Settings.IncludeExchange -DlpAudit $Script:Settings.IncludeDlpAudit -AuditConfig $Script:Settings.IncludeAuditDiscovery
    Export-AssessmentSummary
    $wordPath = $null
    if ($Script:Settings.WordReport) {
        try { $wordPath = New-DlpWordReport }
        catch {
            Write-Log "Word report failed: $($_.Exception.Message)" 'ERROR'
            $Script:Errors.Add([pscustomobject]@{ Step = 'Word assessment report'; Outcome = 'Failed'; Error = $_.Exception.Message })
        }
    }
    if ($Script:ReportObs) { try { Export-ActionPlan } catch { Write-Log "Action plan export failed: $($_.Exception.Message)" 'WARN' } }
    Complete-Run
    $html = Join-Path $Script:OutputDir '00_Assessment_Summary.html'
    Write-Host "Summary: $html" -ForegroundColor Cyan
    if ($wordPath) { Write-Host "Word report: $wordPath" -ForegroundColor Cyan }
    if ($Script:Cache.ContainsKey('AlertDashboardPath')) { Write-Host "Alerts dashboard: $(@($Script:Cache['AlertDashboardPath'])[0])" -ForegroundColor Cyan }
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
        @{ Area = '13 Content Explorer';           Cmdlet = 'Export-ContentExplorerData';               Test = { Export-ContentExplorerData -TagType SensitiveInformationType -TagName 'Credit Card Number' -Workload SPO -PageSize 1 | Out-Null } }
        @{ Area = '15 Data lifecycle and records';  Cmdlet = 'Get-RetentionCompliancePolicy';            Test = { Get-RetentionCompliancePolicy | Select-Object -First 1 } }
        @{ Area = '15 Data lifecycle and records';  Cmdlet = 'Get-RetentionComplianceRule';              Test = { Get-RetentionComplianceRule | Select-Object -First 1 } }
        @{ Area = '15 Data lifecycle and records';  Cmdlet = 'Get-ComplianceTag';                        Test = { Get-ComplianceTag | Select-Object -First 1 } }
        @{ Area = '15 Data lifecycle and records';  Cmdlet = 'Get-AdaptiveScope';                        Test = { Get-AdaptiveScope | Select-Object -First 1 } }
        @{ Area = '16 Insider risk';                Cmdlet = 'Get-InsiderRiskPolicy';                    Test = { Get-InsiderRiskPolicy | Select-Object -First 1 } }
        @{ Area = '16 Communication compliance';    Cmdlet = 'Get-SupervisoryReviewPolicyV2';            Test = { Get-SupervisoryReviewPolicyV2 | Select-Object -First 1 } }
        @{ Area = '16 Communication compliance';    Cmdlet = 'Get-SupervisoryReviewRule';                Test = { Get-SupervisoryReviewRule | Select-Object -First 1 } }
        @{ Area = '16 Information barriers';        Cmdlet = 'Get-InformationBarrierPolicy';             Test = { Get-InformationBarrierPolicy | Select-Object -First 1 } }
        @{ Area = '16 Information barriers';        Cmdlet = 'Get-OrganizationSegment';                  Test = { Get-OrganizationSegment | Select-Object -First 1 } }
        @{ Area = '17 Audit';                       Cmdlet = 'Get-UnifiedAuditLogRetentionPolicy';       Test = { Get-UnifiedAuditLogRetentionPolicy | Select-Object -First 1 } }
        @{ Area = '17 eDiscovery';                  Cmdlet = 'Get-ComplianceCase';                       Test = { Get-ComplianceCase | Select-Object -First 1 } }
        @{ Area = '17 eDiscovery';                  Cmdlet = 'Get-CaseHoldPolicy';                       Test = { $c = Get-ComplianceCase | Select-Object -First 1; if ($c) { Get-CaseHoldPolicy -Case ([string]$c.Identity) | Select-Object -First 1 } } }
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

    if ($Script:Settings.IncludeExchange -or ($Mode -eq 'Menu' -and (Read-YesNo 'Also test Exchange Online access (mail flow rules, audit log status, DLP rule activity)?' $false))) {
        try {
            Connect-Exo
            foreach ($x in @(@{ C = 'Get-TransportRule'; T = { Get-TransportRule | Select-Object -First 1 } }, @{ C = 'Get-AdminAuditLogConfig'; T = { Get-AdminAuditLogConfig } }, @{ C = 'Search-UnifiedAuditLog'; T = { Search-UnifiedAuditLog -StartDate $end.AddDays(-1) -EndDate $end -RecordType ComplianceDLPExchange -ResultSize 1 | Out-Null } })) {
                $o = 'MISSING'; $d = 'Cmdlet not loaded for this account.'
                if (Get-Command $x.C -ErrorAction SilentlyContinue) { try { [void](& $x.T); $o = 'PASS'; $d = 'Cmdlet available and returned successfully.' } catch { $o = 'FAIL'; $d = $_.Exception.Message.Split([Environment]::NewLine)[0] } }
                $results = @($results) + [pscustomobject]@{ Area = '12, 14, 17 Exchange Online'; Cmdlet = $x.C; Result = $o; Detail = $d }
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
        if (@($results | Where-Object { $_.Result -ne 'PASS' -and $_.Area -match '^1[5-7] ' }).Count) {
            Write-Host 'Areas 15 to 17 (retention, insider risk, communication compliance, information barriers, audit, eDiscovery) each have their own Purview roles. See Permissions in the user guide. Missing areas are simply skipped and marked in the report.' -ForegroundColor Yellow
        }
    }
    if (@($results | Where-Object { $_.Cmdlet -eq 'Get-ComplianceCase' -and $_.Result -eq 'PASS' }).Count) {
        Write-Host 'eDiscovery: only cases this account is a member of are visible unless it is an eDiscovery Administrator.' -ForegroundColor Gray
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
    Write-Host '   Microsoft Purview DLP Discovery  (read-only)' -ForegroundColor Cyan
    Write-Host '================================================================================' -ForegroundColor DarkCyan
    Write-Host "   $status" -ForegroundColor $(if ($c) { 'Green' } else { 'Yellow' })
    Write-Host ("   Output: {0}   Activity Explorer: {1}   Exchange: {2}" -f $Script:Settings.OutputRoot,
        $(if ($Script:Settings.IncludeActivity) { "On ($($Script:Settings.ActivityDays) days)" } else { 'Off' }),
        $(if ($Script:Settings.IncludeExchange) { 'On' } else { 'Off' })) -ForegroundColor Gray
    Write-Host ("   Content Explorer: {0}   DLP rule activity: {1}   Customer: {2}" -f $(if ($Script:Settings.IncludeContentExplorer) { 'On' } else { 'Off' }),
        $(if ($Script:Settings.IncludeDlpAudit) { "On ($($Script:Settings.AuditDays) days)" } else { 'Off' }),
        $(if ($Script:Settings.CustomerName) { $Script:Settings.CustomerName } else { '(tenant name)' })) -ForegroundColor Gray
    Write-Host ("   Retention and records: {0}   Risk and compliance: {1}   Audit and eDiscovery: {2}" -f $(if ($Script:Settings.IncludeLifecycle) { 'On' } else { 'Off' }),
        $(if ($Script:Settings.IncludeRiskCompliance) { 'On' } else { 'Off' }), $(if ($Script:Settings.IncludeAuditDiscovery) { 'On' } else { 'Off' })) -ForegroundColor Gray
    Write-Host '--------------------------------------------------------------------------------' -ForegroundColor DarkCyan
}

function Show-AreaMenu {
    Show-Header
    Write-Host '  Report on individual items' -ForegroundColor Cyan
    Write-Host ''
    foreach ($k in $Script:Areas.Keys) {
        if ($k -eq '15') { Write-Host ''; Write-Host '   Related Microsoft Purview areas' -ForegroundColor Cyan }
        Write-Host ('   {0,2}. {1}' -f $k, $Script:Areas[$k])
    }
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
        Write-Host "   7. Sign in with -DisableWAM           : $($Script:Settings.DisableWAM)  (use if sign-in fails or the sign-in window does not appear)"
        Write-Host "   8. Word report with complete report   : $($Script:Settings.WordReport)  (DLP_Assessment_Report.docx)"
        Write-Host "   9. Frameworks in Word report           : $(($Script:Settings.Frameworks) -join ', ')  (NCSC, NISTCSF, NIST80053, ISO27001, All, None)"
        Write-Host "  10. Include Content Explorer in full   : $($Script:Settings.IncludeContentExplorer)  (where sensitive data is stored; needs Content Explorer List Viewer)"
        Write-Host "  11. Content Explorer max SITs           : $($Script:Settings.ContentExplorerMaxTags)"
        Write-Host "  12. Include DLP rule activity in full  : $($Script:Settings.IncludeDlpAudit)  (rule matches, alerts and alerts dashboard from the audit log; needs Exchange Online and View-Only Audit Logs)"
        Write-Host "  13. Rule activity / alerts days (1-90)  : $($Script:Settings.AuditDays)"
        Write-Host "  14. Customer name on the report         : $(if ($Script:Settings.CustomerName) { $Script:Settings.CustomerName } else { '(tenant organisation name)' })"
        Write-Host "  15. Assessor name on the report         : $(if ($Script:Settings.AssessorName) { $Script:Settings.AssessorName } else { '(not shown)' })"
        Write-Host "  16. Logo on the report cover (PNG/JPEG) : $(if ($Script:Settings.LogoPath) { $Script:Settings.LogoPath } else { '(none)' })"
        Write-Host "  17. Compare with previous report folder : $(if ($Script:Settings.CompareWith) { $Script:Settings.CompareWith } else { '(ask when running option 5)' })"
        Write-Host "  18. Names in the alerts dashboard       : $(if ($Script:Settings.AlertDashboardShowNames) { 'Shown' } else { 'Masked (User 1, User 2 ...)' })"
        Write-Host "  19. Include retention and records       : $($Script:Settings.IncludeLifecycle)  (retention policies and labels, records, disposition, adaptive scopes)"
        Write-Host "  20. Include risk and compliance         : $($Script:Settings.IncludeRiskCompliance)  (insider risk, communication compliance, information barriers)"
        Write-Host "  21. Include audit and eDiscovery        : $($Script:Settings.IncludeAuditDiscovery)  (audit log retention and status, eDiscovery cases and holds)"
        Write-Host "  22. Max eDiscovery cases for holds      : $($Script:Settings.MaxEDiscoveryCases)"
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
            '9' { $f = Read-Host '  Frameworks (comma separated)'; if ($f) { $Script:Settings.Frameworks = @($f -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } }
            '10' { $Script:Settings.IncludeContentExplorer = -not $Script:Settings.IncludeContentExplorer }
            '11' { $d = Read-Host '  Max sensitive information types (1-60)'; if ($d -as [int] -and [int]$d -ge 1 -and [int]$d -le 60) { $Script:Settings.ContentExplorerMaxTags = [int]$d } }
            '12' { $Script:Settings.IncludeDlpAudit = -not $Script:Settings.IncludeDlpAudit }
            '13' { $d = Read-Host '  Days'; if ($d -as [int] -and [int]$d -ge 1 -and [int]$d -le 90) { $Script:Settings.AuditDays = [int]$d } }
            '14' { $Script:Settings.CustomerName = (Read-Host '  Customer name (blank for the tenant name)').Trim() }
            '15' { $Script:Settings.AssessorName = (Read-Host '  Assessor name / company (blank to hide)').Trim() }
            '16' { $p = (Read-Host '  Path to a PNG or JPEG logo (blank for none)').Trim('" '); if (-not $p -or (Test-Path $p)) { $Script:Settings.LogoPath = $p } else { Write-Host '  File not found.' -ForegroundColor Yellow; Start-Sleep 2 } }
            '18' { $Script:Settings.AlertDashboardShowNames = -not $Script:Settings.AlertDashboardShowNames }
            '17' { $p = (Read-Host '  Previous report folder (blank to ask each time)').Trim('" '); if (-not $p -or (Test-Path $p)) { $Script:Settings.CompareWith = $p } else { Write-Host '  Folder not found.' -ForegroundColor Yellow; Start-Sleep 2 } }
            '19' { $Script:Settings.IncludeLifecycle = -not $Script:Settings.IncludeLifecycle }
            '20' { $Script:Settings.IncludeRiskCompliance = -not $Script:Settings.IncludeRiskCompliance }
            '21' { $Script:Settings.IncludeAuditDiscovery = -not $Script:Settings.IncludeAuditDiscovery }
            '22' { $d = Read-Host '  Max cases (1-500)'; if ($d -as [int] -and [int]$d -ge 1 -and [int]$d -le 500) { $Script:Settings.MaxEDiscoveryCases = [int]$d } }
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
        Write-Host '   4. Report on individual items'
        Write-Host '   5. Complete tenant report (DLP plus retention, risk and compliance, audit and eDiscovery)'
        Write-Host '   6. Settings'
        Write-Host '   7. Disconnect / switch tenant'
        Write-Host '   8. DLP alerts dashboard (what the DLP alerts are about)'
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
                Write-Host '  Complete tenant report' -ForegroundColor Cyan
                $Script:Settings.IncludeActivity = Read-YesNo "  Include Activity Explorer data ($($Script:Settings.ActivityDays) days)?" $Script:Settings.IncludeActivity
                $Script:Settings.IncludeExchange = Read-YesNo '  Include Exchange Online mail flow rules and audit config?' $Script:Settings.IncludeExchange
                $Script:Settings.IncludeContentExplorer = Read-YesNo '  Include Content Explorer (where sensitive data is stored; can take several minutes)?' $Script:Settings.IncludeContentExplorer
                $Script:Settings.IncludeDlpAudit = Read-YesNo "  Include DLP rule activity and alerts from the audit log, with the alerts dashboard ($($Script:Settings.AuditDays) days; signs in to Exchange Online)?" $Script:Settings.IncludeDlpAudit
                $Script:Settings.IncludeLifecycle = Read-YesNo '  Include data lifecycle and records management (retention policies and labels, records, adaptive scopes)?' $Script:Settings.IncludeLifecycle
                $Script:Settings.IncludeRiskCompliance = Read-YesNo '  Include risk and compliance (insider risk, communication compliance, information barriers)?' $Script:Settings.IncludeRiskCompliance
                $Script:Settings.IncludeAuditDiscovery = Read-YesNo '  Include audit and eDiscovery (audit log retention and status, eDiscovery cases and holds; also signs in to Exchange Online for the audit log status)?' $Script:Settings.IncludeAuditDiscovery
                $Script:AskCompare = $true
                Invoke-FullReport; Wait-Key; $Script:AskCompare = $false
            }
            '6' { Show-SettingsMenu }
            '7' { Disconnect-Tenant; $Script:Settings.UserPrincipalName = ''; Wait-Key }
            '8' {
                Show-Header
                Write-Host '  DLP alerts dashboard' -ForegroundColor Cyan
                Write-Host '  Reads DLP alerts and rule matches from the unified audit log (signs in to Exchange Online) and builds DLP_Alerts_Dashboard.html.' -ForegroundColor Gray
                $d = Read-Host "  How many days to include (1-90) [$($Script:Settings.AuditDays)]"
                if ($d -as [int] -and [int]$d -ge 1 -and [int]$d -le 90) { $Script:Settings.AuditDays = [int]$d }
                Invoke-AlertDashboard; Wait-Key
            }
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
    'AlertDashboard' { Invoke-AlertDashboard; if (-not $NoDisconnect) { Disconnect-Tenant } }
    'Area' {
        if (-not $Area) { throw 'Use -Area with one or more area numbers (1-17), e.g. -Area 2,3,8,15' }
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
