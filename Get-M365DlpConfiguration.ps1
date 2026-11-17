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
      and for the complete report 00_Assessment_Summary.csv/.html.

.PARAMETER DisableWAM
    Sign in without the Windows Web Account Manager broker (passes -DisableWAM where the module supports it).
    Use if sign-in fails, for example with "You cannot call a method on a null-valued expression".

.PARAMETER Mode
    Menu (default) | InstallModules | TestAccess | Full | Area

.PARAMETER Area
    With -Mode Area: one or more area numbers (1-12) or names, e.g. -Area 2,3,8

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1
    Starts the interactive menu.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com -IncludeActivityExplorer
    Runs the complete report without the menu.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode Area -Area 2,3,8 -UserPrincipalName admin@contoso.onmicrosoft.com
    Reports on policies, rules and Endpoint DLP only.

.EXAMPLE
    .\Get-M365DlpConfiguration.ps1 -Mode TestAccess -UserPrincipalName admin@contoso.onmicrosoft.com

.NOTES
    Requires: ExchangeOnlineManagement 3.2.0 or later. Optional: ImportExcel.
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
    [switch]$DisableWAM,
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
    Export-Dataset '71_Audit_Configuration' $audit
    Reset-AfterExo
}

$Script:AreaFunctions = @{
    '1' = 'Export-AreaTenant'; '2' = 'Export-AreaPolicies'; '3' = 'Export-AreaRules'; '4' = 'Export-AreaClassifiers'
    '5' = 'Export-AreaSITs'; '6' = 'Export-AreaEDM'; '7' = 'Export-AreaLabels'; '8' = 'Export-AreaEndpoint'
    '9' = 'Export-AreaAlerts'; '10' = 'Export-AreaOwnership'; '11' = 'Export-AreaActivity'; '12' = 'Export-AreaExchange'
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
    if ($Script:Settings.IncludeExchange) { Write-Host ''; Write-Host "--- $($Script:Areas['12']) ---" -ForegroundColor Cyan; Export-AreaExchange }
    Export-AssessmentSummary
    Complete-Run
    $html = Join-Path $Script:OutputDir '00_Assessment_Summary.html'
    Write-Host "Summary: $html" -ForegroundColor Cyan
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
    Write-Host '   Microsoft Purview DLP Discovery  (read-only)' -ForegroundColor Cyan
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
        Write-Host "   7. Sign in with -DisableWAM           : $($Script:Settings.DisableWAM)  (use if sign-in fails or the sign-in window does not appear)"
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
