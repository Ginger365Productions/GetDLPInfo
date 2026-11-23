# Get-M365DlpConfiguration

A read-only PowerShell tool for discovering and documenting **Microsoft Purview Data Loss Prevention (DLP)** configuration in any Microsoft 365 tenant. It also covers the related Purview areas that affect how information is protected: **data lifecycle and records management**, **insider risk, communication compliance and information barriers**, and **audit and eDiscovery**.

It is built for consultants and administrators who need a consistent evidence pack at the start of a DLP engagement. Run it the same way in every tenant, get the same set of outputs, and compare tenants or runs over time.

> **Read-only:** the script only reads from the tenant. It never creates, changes, enables or deletes any policy, rule, label or setting. The `Set-`/`New-` commands you will see in the code are example text for the Word report, not commands it runs; see [Is it really read-only?](#is-it-really-read-only).

> **Beta available:** a beta version with optional **Microsoft Graph** data (licensing for DLP features, DLP alerts raised, Microsoft Entra role holders and SharePoint sharing settings) is in the [`beta`](beta/README.md) folder. The script in this root folder is the main, supported version.

> **New to the tool?** Start with the [Quick start guide](QUICKSTART.md): what the tool does, how to install the modules, connect to a tenant and run the full report. A Word version with screenshots is in [`M365-DLP-Discovery-Quick-Start-Guide.docx`](M365-DLP-Discovery-Quick-Start-Guide.docx).

---

## Contents

- [Is it really read-only?](#is-it-really-read-only)
- [What it collects](#what-it-collects)
- [Prerequisites](#prerequisites)
- [Quick start](#quick-start)
- [The menu](#the-menu)
- [Running without the menu](#running-without-the-menu)
- [Report areas](#report-areas)
- [Output](#output)
- [Word assessment report](#word-assessment-report)
- [DLP alerts dashboard](#dlp-alerts-dashboard)
- [Permissions](#permissions)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [Documentation](#documentation)

---

## What it collects

| Category | What is reported |
|---|---|
| DLP policies | Mode (enforced, simulation, disabled), priority, workloads, locations and exceptions, change history |
| Workload coverage | Which policies cover Exchange, SharePoint, OneDrive, Teams, Endpoint, Microsoft 365 Copilot, on-premises, Power BI and cloud apps |
| Rules, conditions and actions | Every rule with its conditions, exceptions, block / block with override / notify / audit response, policy tips, override settings, alerts and incident reports |
| Classification | Built-in and custom sensitive information types, custom rule package XML, keyword dictionaries, Exact Data Match schemas, trainable classifiers used in rules |
| Confidence levels | Confidence level and instance counts for every classifier referenced by every rule |
| Sensitivity labels | Labels, label policies, auto-labelling policies and rules, and labels used as DLP conditions. **Label settings:** encryption (who gets access, user-defined permissions, offline access, expiry), content marking (header, footer, watermark), what each label applies to (files, emails, sites, groups and Teams, meetings) and site and group settings (privacy, guest access, external sharing, unmanaged devices); default and mandatory labelling from label policies |
| Endpoint DLP | Global settings (browsers, service domains, printer, removable media and network share groups) and per-rule device restrictions (USB, printing, browser and cloud upload, clipboard) |
| Alerts | DLP alert policies, and rules that raise alerts or incident reports |
| Ownership | Purview role groups and members holding DLP-relevant roles |
| Activity Explorer *(optional)* | DLP rule matches and other activity for the last 1–30 days |
| Exchange Online *(optional)* | Mail flow rules (including legacy Exchange DLP) and unified audit log status |
| Content Explorer *(optional)* | How many stored emails, files and Teams messages contain each sensitive information type and label, by location, and where the largest concentrations are |
| DLP rule activity and alerts *(optional)* | DLP rule matches, overrides and false positives, and the DLP alerts raised, from the unified audit log (default 30 days), plus an **alerts dashboard** |
| Data lifecycle and records *(on by default)* | Retention policies and their rules, retention label policies (published or auto-applied), retention labels (records, regulatory records, disposition review, event-based retention), adaptive scopes, and retention coverage of email, SharePoint, OneDrive and Teams |
| Risk and compliance *(on by default)* | Insider risk management policies and their templates, communication compliance policies and rules, information barrier policies and segments |
| Audit and eDiscovery *(on by default)* | Audit log retention policies, unified audit log status, eDiscovery cases (Standard and Premium) and their holds |

**No Microsoft Graph.** Everything uses Security & Compliance PowerShell and, optionally, Exchange Online PowerShell.

The three related Purview areas are included in the complete report by default and can each be turned off (**Settings 19–21**, the questions in **option 5**, or `-Scope`). Each is also available on its own in **option 4** (areas 15, 16 and 17).

A complete run also produces a **Word assessment report** (`DLP_Assessment_Report.docx`): a plain-English document that explains what each DLP feature is and why it matters, shows what was found in the tenant, lists things to consider, and links to the relevant Microsoft Learn pages. It is built directly by the script; Microsoft Word and extra modules are not required.

The report also scores the DLP programme (**maturity**), maps data types against locations (**coverage map**), explains **what actually happens** when someone shares sensitive data, runs **hygiene checks**, and ends with a **prioritised action plan** showing effort and the licence each action needs. It can show **changes since the previous report** and carry the customer's name and logo.

It also produces an **assessment summary**: one row per DLP topic with an automatically calculated finding, the file that holds the evidence, a suggested next activity, and blank Status / Owner / TargetDate / Notes columns ready to use as an engagement tracker.

---

## Is it really read-only?

Yes. If you look through the script you will see `Set-` and `New-` commands, so here is exactly what they are.

**Commands that read from the tenant (the only tenant commands the script runs)**

| Command | What it does |
|---|---|
| `Get-*` (for example `Get-DlpCompliancePolicy`, `Get-Label`, `Get-TransportRule`) | Read configuration |
| `Export-ActivityExplorerData`, `Export-ContentExplorerData` | Despite the name, these only **read** data from Purview and return it to the script |
| `Search-UnifiedAuditLog` | Reads DLP rule matches and alerts from the audit log |
| `Connect-IPPSSession`, `Connect-ExchangeOnline`, `Disconnect-ExchangeOnline` | Sign in and out; they change nothing in the tenant |

**Tenant-changing commands that appear in the code but are never run**

`Set-DlpCompliancePolicy`, `New-DlpCompliancePolicy`, `Set-DlpComplianceRule`, `New-DlpComplianceRule`, `New-LabelPolicy`, `Set-AdminAuditLogConfig`, `New-RetentionCompliancePolicy`, `New-RetentionComplianceRule` and `New-UnifiedAuditLogRetentionPolicy` appear only inside `Add-ImplementationGuidance`, inside `@' ... '@` blocks. PowerShell treats these blocks as plain text. The script prints them into the **"Applying the configuration"** section of the Word report (the last numbered section) as examples for the customer to review and run themselves, if they choose to.

**Commands that change your own PC (not the tenant)**

| Command | What it does | When |
|---|---|---|
| `New-Item` | Creates the output folder | Every run, or when you change the output folder in Settings |
| `Set-Content`, `Add-Content`, `Export-Csv`, `Export-Excel` | Write the output files and `Collection.log` | Every run |
| `Remove-Item` | Deletes `DLP_Assessment_Report.docx` in that run's output folder before writing it again | When the Word report is built |
| `Install-Module`, `Install-PackageProvider` | Install ExchangeOnlineManagement and ImportExcel for your user account only | Only when you choose option 1 and answer Yes |
| `Uninstall-Module`, `Remove-Module` | Remove old copies of ExchangeOnlineManagement | Only when you choose the clean reinstall in option 1 |
| `Import-Module`, `Add-Type` | Load modules and .NET libraries into the PowerShell window | When needed |

**The script's own functions.** Most `New-`, `Add-`, `Export-` and `Invoke-` names in the code are the script's own functions, for example `New-DlpWordReport`, `Add-DocTable` and `Export-Dataset`. They build the report and output files.

**Check it yourself.** Run the script with an account that has only read-only roles (see [Permissions](#permissions)). If anything tried to make a change, the tenant would refuse it. To list every command the script runs that could change a tenant, search the code outside the example blocks:

```powershell
Select-String -Path .\Get-M365DlpConfiguration.ps1 -Pattern '^\s*(Set|New|Remove|Enable|Disable|Start|Stop)-(Dlp|Label|AutoSensitivity|AdminAuditLog|TransportRule|Protection|Retention|ComplianceTag|UnifiedAuditLog|InsiderRisk|SupervisoryReview|InformationBarrier|OrganizationSegment|ComplianceCase|CaseHold|AdaptiveScope)' |
    Where-Object { $_.Line -notmatch '^\s*#' }
```

Every result is inside an example block in `Add-ImplementationGuidance`.

---

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | Windows PowerShell 5.1 or PowerShell 7.2+ (7 recommended) |
| **ExchangeOnlineManagement** module | **Required**, version **3.2.0 or later** |
| **ImportExcel** module | Optional, for the multi-sheet Excel workbook (Excel itself is not needed) |
| Network | HTTPS to Microsoft 365 and `login.microsoftonline.com` |
| Account | An admin account in the target tenant with read access to Microsoft Purview (see [Permissions](#permissions)) |

### Install the modules

The easiest way is **menu option 1**, or:

```powershell
.\Get-M365DlpConfiguration.ps1 -Mode InstallModules
```

To install manually:

```powershell
# Windows PowerShell 5.1 only
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force

# All versions
Install-Module ExchangeOnlineManagement -Scope CurrentUser -MinimumVersion 3.2.0
Install-Module ImportExcel -Scope CurrentUser
```

Check what is installed:

```powershell
Get-Module -ListAvailable ExchangeOnlineManagement, ImportExcel | Select-Object Name, Version, ModuleBase
```

> Keep only **one** version of ExchangeOnlineManagement installed. Multiple side-by-side versions are a common cause of sign-in errors (see [Troubleshooting](#troubleshooting)).

---

## Quick start

```powershell
# 1. Get the script
git clone <this repository URL>
cd GetDLPInfo

# 2. Allow it to run in this window only
Unblock-File .\Get-M365DlpConfiguration.ps1
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 3. Start the menu
.\Get-M365DlpConfiguration.ps1
```

Then work through the menu in order: **1** (modules) → **2** (connect) → **3** (test access) → **5** (complete report).

---

## The menu

```
================================================================================
   Microsoft Purview DLP Discovery  (read-only)
================================================================================
   Connected: admin@contoso.onmicrosoft.com  |  Tenant: 00000000-0000-...
   Output: C:\Users\you\Documents   Activity Explorer: Off   Exchange: Off
   Content Explorer: Off   DLP rule activity: Off   Customer: (tenant name)
   Retention and records: On   Risk and compliance: On   Audit and eDiscovery: On
--------------------------------------------------------------------------------
   1. Install / update required PowerShell modules
   2. Connect to tenant
   3. Test credentials and tenant access
   4. Report on individual items
   5. Complete tenant report (DLP plus retention, risk and compliance, audit and eDiscovery)
   6. Settings
   7. Disconnect / switch tenant
   8. DLP alerts dashboard (what the DLP alerts are about)
   Q. Quit
```

| Option | Description |
|---|---|
| **1. Install / update modules** | Shows installed versions, installs or updates ExchangeOnlineManagement and ImportExcel for the current user (no admin rights needed), and offers a **clean reinstall** that removes every version first when more than one is found. |
| **2. Connect to tenant** | Signs in to Security & Compliance PowerShell, or switches to a different account or tenant. Shows the account, tenant ID and organisation, and how many key DLP cmdlets the account can use. |
| **3. Test credentials and tenant access** | Runs a minimal read-only call against every cmdlet the reports use and shows **PASS**, **MISSING** (role or licence missing) or **FAIL** per area. Lists the account's Purview role groups and saves the results to CSV. Optionally tests Exchange Online. |
| **4. Report on individual items** | Choose one or more of the 17 [report areas](#report-areas), for example `2,3,8` or `15,16,17`. |
| **5. Complete tenant report** | Every area plus the assessment summary (CSV and HTML), the Word assessment report, the action plan and the Excel workbook. Asks whether to include Activity Explorer, Exchange Online, Content Explorer, DLP rule activity, **data lifecycle and records**, **risk and compliance** and **audit and eDiscovery** (the last three default to Yes), and, when an earlier report for the same tenant exists, whether to show the changes since then. |
| **6. Settings** | Output folder, admin account, Activity Explorer days and page cap, Exchange Online, sign-in with `-DisableWAM`, the Word report and its frameworks, Content Explorer (on/off, number of types), DLP rule activity (on/off, days), **customer name, assessor name and logo** for the report cover, a fixed folder to compare with, names in the alerts dashboard, and **items 19–22**: include retention and records, include risk and compliance, include audit and eDiscovery, and the maximum number of eDiscovery cases whose holds are read. |
| **7. Disconnect / switch tenant** | Signs out so you can move to the next customer tenant without restarting. |
| **8. DLP alerts dashboard** | Asks how many days to include (1–90), reads the DLP alerts and rule matches from the unified audit log, and opens **DLP_Alerts_Dashboard.html**: an easy-to-read page explaining what the alerts are about. See [DLP alerts dashboard](#dlp-alerts-dashboard). |

Once connected, the connection is re-used by every option. Options 3–5 connect automatically if you have not used option 2.

---

## Running without the menu

Use `-Mode` for repeat runs, scripting or several tenants.

```powershell
# Install or update modules
.\Get-M365DlpConfiguration.ps1 -Mode InstallModules

# Test access
.\Get-M365DlpConfiguration.ps1 -Mode TestAccess -UserPrincipalName admin@contoso.onmicrosoft.com

# Complete report
.\Get-M365DlpConfiguration.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com `
    -OutputRoot "D:\Engagements\Contoso"

# Complete report including Activity Explorer (30 days) and Exchange Online
.\Get-M365DlpConfiguration.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com `
    -IncludeActivityExplorer -ActivityDays 30 -ActivityMaxPages 50 -IncludeExchange

# Complete report with evidence, customer branding and comparison with the previous run
.\Get-M365DlpConfiguration.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com `
    -IncludeExchange -IncludeContentExplorer -IncludeDlpAudit -AuditDays 30 `
    -CustomerName 'Contoso Ltd' -AssessorName 'Jane Smith, Fabrikam Consulting' -LogoPath .\contoso-logo.png `
    -CompareWith Latest

# DLP alerts dashboard only (last 30 days), with people's names masked
.\Get-M365DlpConfiguration.ps1 -Mode AlertDashboard -AuditDays 30 -MaskNamesInDashboard

# Selected areas only (policies, rules, Endpoint DLP)
.\Get-M365DlpConfiguration.ps1 -Mode Area -Area 2,3,8 -UserPrincipalName admin@contoso.onmicrosoft.com

# Retention, risk and compliance, and audit and eDiscovery only
.\Get-M365DlpConfiguration.ps1 -Mode Area -Area 15,16,17 -UserPrincipalName admin@contoso.onmicrosoft.com

# Complete report covering DLP only, or DLP plus retention
.\Get-M365DlpConfiguration.ps1 -Mode Full -Scope DLP -UserPrincipalName admin@contoso.onmicrosoft.com
.\Get-M365DlpConfiguration.ps1 -Mode Full -Scope DLP, Lifecycle -UserPrincipalName admin@contoso.onmicrosoft.com

# Areas by name
.\Get-M365DlpConfiguration.ps1 -Mode Area -Area 'Sensitivity labels','Endpoint'
```

Several tenants in sequence:

```powershell
$tenants = @(
    @{ Upn = 'admin@contoso.onmicrosoft.com';  Folder = 'D:\Engagements\Contoso'  }
    @{ Upn = 'admin@fabrikam.onmicrosoft.com'; Folder = 'D:\Engagements\Fabrikam' }
)
foreach ($t in $tenants) {
    .\Get-M365DlpConfiguration.ps1 -Mode TestAccess -UserPrincipalName $t.Upn -OutputRoot $t.Folder
    .\Get-M365DlpConfiguration.ps1 -Mode Full       -UserPrincipalName $t.Upn -OutputRoot $t.Folder
}
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-Mode` | `Menu` | `Menu`, `InstallModules`, `TestAccess`, `Full`, `Area` or `AlertDashboard` |
| `-Area` | | With `-Mode Area`: area numbers 1–17 or part of an area name |
| `-Scope` | `All` | What the complete report covers. DLP is always included; add `Lifecycle` (retention and records), `RiskCompliance` (insider risk, communication compliance, information barriers) and `Audit` (audit log retention, eDiscovery). `All` includes everything |
| `-MaxEDiscoveryCases` | `100` | Maximum number of active eDiscovery cases whose holds are read (1–500); each case needs one extra call |
| `-OutputRoot` | Documents | Folder in which the time-stamped output folder is created |
| `-UserPrincipalName` | | Admin account to sign in with |
| `-IncludeActivityExplorer` | Off | Include Activity Explorer in the complete report |
| `-ActivityDays` | `7` | Days of Activity Explorer data (1–30) |
| `-ActivityMaxPages` | `20` | Cap on 5,000-record Activity Explorer pages (1–200) |
| `-IncludeExchange` | Off | Include Exchange Online in the complete report and access test |
| `-DisableWAM` | Off | Sign in through the browser instead of the Windows account broker |
| `-Frameworks` | All four | Frameworks referenced in the Word report: `NCSC`, `NISTCSF`, `NIST80053`, `ISO27001`, `NIST`, `ISO`, `All` or `None` |
| `-NoWordReport` | Off | Do not create the Word assessment report with the complete report |
| `-IncludeContentExplorer` | Off | Count stored items containing each sensitive information type and label, by location (needs Content Explorer List Viewer) |
| `-ContentExplorerMaxTags` | `25` | Maximum number of sensitive information types queried in Content Explorer (1–60) |
| `-IncludeDlpAudit` | Off | Read DLP rule matches, overrides and false positives from the unified audit log (signs in to Exchange Online) |
| `-AuditDays` | `30` | Days of DLP rule activity and alerts (1–90) |
| `-MaskNamesInDashboard` | Off | Show people as User 1, User 2 … and hide file names and subjects in the alerts dashboard |
| `-AuditMaxRecords` | `20000` | Cap on audit records read per record type |
| `-CustomerName` | Tenant name | Shown on the report cover as "Prepared for" |
| `-AssessorName` | | Shown on the report cover as "Prepared by" |
| `-LogoPath` | | PNG or JPEG logo for the report cover |
| `-CompareWith` | | A previous complete-report folder, or `Latest` for the newest earlier report for the same tenant |
| `-NoDisconnect` | Off | Leave the session connected when the script ends |

Full help: `Get-Help .\Get-M365DlpConfiguration.ps1 -Detailed`

---

## Report areas

| # | Area | Output files |
|---|---|---|
| 1 | Tenant information | `01` |
| 2 | DLP policies and workload coverage | `10`, `11`, `12` |
| 3 | DLP rules: conditions, actions, policy tips and overrides | `13`, `14`, `16` |
| 4 | Classifiers and confidence levels used in rules | `15`, `26`, `27` |
| 5 | Sensitive information types (built-in and custom) | `20`, `21`, `22` (+ XML), `23` |
| 6 | Exact Data Match (EDM) | `24`, `25`, `15` |
| 7 | Sensitivity labels and auto-labelling | `30`–`34` |
| 8 | Endpoint DLP: USB, printing, browser and cloud uploads | `40`, `41`, `16`, `12` |
| 9 | DLP alerts and alert policies | `50`, `51`, `52` |
| 10 | Operational ownership: role groups and members | `53`, `54`, `55` |
| 11 | Activity Explorer | `60`, `61` |
| 12 | Exchange Online mail flow rules and audit | `70`, `71` |
| 13 | Content Explorer: where sensitive data is stored | `62`, `63` |
| 14 | DLP rule activity and alerts from the audit log, with the alerts dashboard | `64`–`67`, `DLP_Alerts_Dashboard.html` |
| 15 | Data lifecycle and records: retention policies, retention labels, records, adaptive scopes | `100`–`106` |
| 16 | Risk and compliance: insider risk, communication compliance, information barriers | `200`, `201`, `210`–`212`, `220`–`222` |
| 17 | Audit and eDiscovery: audit log retention, audit log status, eDiscovery cases and holds | `300`, `310`, `311`, `71` (signs in to Exchange Online for the audit log status) |

Area 7 also writes `35_Sensitivity_Label_Settings` (encryption, content marking, scope and site and group settings for each label). Files for areas 15–17 are numbered in the 100s, 200s and 300s so the DLP file numbers stay the same as earlier versions.

---

## Output

Each run writes to a new folder under the output root:

| Folder | Created by |
|---|---|
| `M365-DLP-Full-yyyyMMdd-HHmmss` | Complete report (option 5 / `-Mode Full`) |
| `M365-DLP-Areas-yyyyMMdd-HHmmss` | Individual areas (option 4 / `-Mode Area`) |
| `M365-DLP-AccessTest-yyyyMMdd-HHmmss` | Access test (option 3 / `-Mode TestAccess`) |

Inside each folder:

| File | Description |
|---|---|
| `DLP_Assessment_Report.docx` | Complete report only. Plain-English Word report: executive summary, key observations, an explanation and findings for every area, Microsoft Learn links, glossary |
| `00_Assessment_Summary.html` / `.csv` | Complete report only. One row per topic with findings, the licence needed and tracker columns |
| `DLP_Alerts_Dashboard.html` | When DLP rule activity is included, or from menu option 8. Self-contained page explaining what the DLP alerts are about |
| `66_DLP_Alert_Events`, `67_DLP_Alerts` | Alert records from the audit log, and one row per alert linked to the rule matches behind it |
| `80_Action_Plan.csv` | Complete report only. Prioritised actions with effort, licence, owner, target date and status columns |
| `81_Maturity_Score`, `82_Hygiene_Checks`, `83_Changes_Since_Last_Run`, `84_Coverage_Map` | Complete report only. The scores, checks, changes and coverage map shown in the Word report |
| `DLP_Configuration.xlsx` | Every dataset as a filterable worksheet (requires ImportExcel) |
| `NN_Name.csv` | Flattened data for each dataset |
| `NN_Name.json` | Full-fidelity data, including nested rule conditions |
| `22_RulePackage_*.xml` | Custom sensitive information type definitions |
| `99_Collection_Errors.csv` | Steps that were **skipped** (cmdlet not available) or **failed**, with the reason |
| `Collection.log` | Time-stamped log of every step |

Areas that could not be collected are marked **NOT COLLECTED** in the summary, so a permissions gap is never mistaken for "nothing configured".

Useful queries on the output:

```powershell
# Rules that only audit
Import-Csv .\14_DLP_Rules_Conditions_Actions.csv | Where-Object ResponseInterpretation -eq 'Audit only'

# Policies still in simulation
Import-Csv .\11_DLP_Policies_Summary.csv | Where-Object Mode -like 'Test*'

# Custom SITs not used by any rule
Import-Csv .\21_Sensitive_Information_Types_Custom.csv | Where-Object UsedInDlpRules -eq 'False'
```

> Activity Explorer output contains user names, file names and paths. Store and share it according to the customer's data-handling requirements.

---

## Word assessment report

Every complete report (menu option 5, or `-Mode Full`) also creates **`DLP_Assessment_Report.docx`** in the output folder. It is a plain-English report for stakeholders who are not DLP specialists. It explains what each part of Microsoft Purview DLP is and why it matters, shows what is configured in the tenant, highlights things to consider, and links to the relevant Microsoft Learn pages.

- The script builds the document itself. **Microsoft Word and extra PowerShell modules are not required** on the machine that runs it.
- It is created from the same data, in the same run, as the CSV, JSON and Excel outputs, so everything is consistent.
- It is **on by default**. Turn it off with **Settings > 8** or the `-NoWordReport` parameter.
- It is only produced by the complete report, because it needs data from every area. Individual area reports (option 4) do not create it.
- If report generation fails, the rest of the output is still written, and the error is logged in `Collection.log` and `99_Collection_Errors.csv`.

### Opening the report

1. Open `DLP_Assessment_Report.docx` from the `M365-DLP-Full-yyyyMMdd-HHmmss` folder.
2. If Word opens it in Protected View, select **Enable Editing**.
3. When Word asks whether to update fields, select **Yes**. This builds the table of contents and page numbers.
   If you are not prompted, right-click the Contents page and choose **Update Field > Update entire table**.
4. Save. The report is a normal Word document: brand it, add commentary, or export it to PDF (**File > Export > Create PDF**).

### What the report contains

| Section | Contents |
|---|---|
| Cover and contents | Organisation, tenant ID, collection date, account used, method, confidentiality note |
| 1. About this report | Purpose, how the data was collected (read-only), how sections are organised |
| 2. Executive summary | Headline figures, **DLP maturity** scores (0–5 for six dimensions, with changes since the last report), **changes since the last assessment**, **key observations** ranked High / Medium / Low, collection gaps, **status by area**, **framework alignment at a glance** |
| 3. DLP policies | Each policy with mode, locations, rule count and last change date |
| 4. Where DLP is applied | Exchange, SharePoint, OneDrive, Teams, devices, Copilot and other locations: covered or not; **coverage map** of data types (financial, health, personal and ID, credentials, labelled content) against locations, showing the strongest response in each cell |
| 5. DLP rules | How rules respond (block, block with override, notify, endpoint restrictions, audit only) and a table of rules (each rule is written out in full in Appendix E) |
| 6. User experience | Policy tips, custom tip text and overrides; **what actually happens** scenarios, one per rule, in plain English |
| 7. Sensitive information types | Detectors in use and the policies that use them; **where the information is stored** (Content Explorer counts by location, types found but not protected, top sites) |
| 8. Custom SITs and keyword dictionaries | Organisation-specific detectors, what each detects, whether they are used, and each dictionary with its keyword count |
| 9. Exact Data Match | EDM schemas with their columns (searchable columns marked) and EDM-based detectors |
| 10. Trainable classifiers | Machine-learning classifiers referenced by rules |
| 11. Confidence levels and instance counts | Detection thresholds, low-confidence and single-instance detections |
| 12. Sensitivity labels and DLP | Label taxonomy, encryption, labels used in DLP; each label policy (labels published, who to, mandatory/default label settings) and each auto-labelling policy (label applied, status, where, what it looks for); **label settings** for every label: what it applies to, encryption, content marking, and sites, groups and Teams settings |
| 13. Endpoint DLP | USB and removable storage, printing, browser and cloud uploads; device restrictions by rule; tenant-wide device settings with friendly names |
| 14. DLP for Microsoft 365 Copilot | Policies protecting Copilot and Copilot Chat |
| 15. Deployment state and simulation | Enforced, simulation and disabled policies; **configuration hygiene checks** (unchanged templates, high-risk rules that do not block, alert recipients, overlapping rules, stale or long-running simulation policies, default policy tips, unused detectors and labels, unprotected stored data, rules that never fire or are routinely overridden) |
| 16. Alerts and monitoring | Rules that raise alerts or incident reports, and who receives them; each DLP alert policy with severity, trigger and recipients |
| 17. Activity Explorer and DLP rule activity | Activity summary, and matches, overrides, false positives and distinct users per rule from the audit log (user names are kept out of the report) |
| 18. Exchange mail flow and auditing | Each mail flow rule with its state and Exchange's own description, and unified audit log status (when Exchange Online is included) |
| 19. Operational ownership | Role groups and members with DLP-relevant roles |
| 20. Data lifecycle and records management *(when included)* | Retention coverage of email, SharePoint, OneDrive, Teams chats and channels; each retention policy (type, status, locations, retention); each retention label (record type, retention, disposition review, how applied); adaptive scopes |
| 21. Insider risk, communication compliance and information barriers *(when included)* | Insider risk policies and templates, and whether DLP raises the high-severity alerts a Data leaks policy needs; communication compliance policies (reviewers, share reviewed); information barrier policies, segments and application status |
| 22. Audit and eDiscovery *(when included)* | Unified audit log status, audit log retention policies (and whether DLP activity is kept longer), eDiscovery cases and holds by type (case names are kept out of the report) |
| 23. Recommended next steps | **Prioritised action plan** (priority, action, effort, licence), licensing notes, and topic-by-topic findings with the licence each needs |
| 24. Applying the configuration | Order of work and step-by-step guidance for each gap found |
| Appendices | A Collection details and gaps, B Glossary, C Evidence files, D Framework references, **E DLP policies and rules in detail** |

Sections 20 to 22 appear only for the areas included in the run, and the sections after them are numbered to follow on. With every area included the next steps are section 23 and the guidance section 24; with `-Scope DLP` they are sections 20 and 21 as in earlier versions.

### How each area section is laid out

Sections 3 to 22 all use the same layout:

| Part | What it gives the reader |
|---|---|
| **What is it?** | A short, non-technical explanation, for example what a sensitivity label or a sensitive information type is |
| **Why it matters** | How the feature helps protect information |
| **What we found** | The tenant's configuration, with counts and tables. Areas that could not be fully read are flagged in an orange **Not fully collected** box |
| **Things to consider** | Observations generated from the data, plus good-practice prompts |
| **Framework alignment** | NCSC, NIST and ISO 27001 references, what each expects, and the tenant status |
| **Learn more** | Clickable links to the relevant Microsoft Learn pages |

Long tables are shortened in the report (typically to the first 40 or 60 rows) with a pointer to the CSV file that holds the full data.

### Configuration details and Appendix E (no need to open the spreadsheets)

The report includes the collected configuration itself, written so it can be read and reviewed without opening the CSV or Excel files. It is translated rather than dumped: GUIDs become label names, setting names become plain descriptions, and modes become phrases such as "Simulation, silent (nothing shown to users)".

- **Configuration details** parts inside sections 8, 9, 12, 13, 16 and 18 cover custom SITs and dictionaries, EDM schemas, label and auto-labelling policies, device restrictions and Endpoint DLP settings, alert policies and mail flow rules.
- **Appendix E: DLP policies and rules in detail** has one entry per policy, in priority order: a short card (mode, locations included and excluded, description, priority, last changed) followed by every rule as three plain-English lines:

  | Line | Example |
  |---|---|
  | **When** | Content contains any of: **Credit Card Number** (high confidence, 10 or more); and content is shared with people **outside the organisation** |
  | **Except when** | Recipient domain is: contoso-auditors.co.uk |
  | **Then** | **Block** access for people outside the organisation; users can **override** with a business justification; send an **alert** to dlp-alerts@contoso.com |

  Rules with no blocking, notification or device action are shown as **Audit only**. Where a rule groups conditions with AND / OR logic, the exact logic is in `14_DLP_Rules_Conditions_Actions`.

The raw files are still produced for filtering, evidence and anything the report shortens.

### Observations and priorities

Observations come from simple, transparent checks on the collected configuration. For example:

| Priority | Example checks |
|---|---|
| **High** | No DLP policies; a core location (Exchange, SharePoint, OneDrive, Teams, devices) with no enforced policy; no sensitivity labels, or labels not published; no rules raising alerts; unified audit log disabled; no retention policies or retention labels at all |
| **Medium** | Labels for highly sensitive content that do not encrypt; a core location with no retention policy; retention policy distribution errors; no insider risk policies, or a Data leaks policy with no high-severity DLP alerts to trigger it; communication compliance policies with no reviewers; policies left in simulation; rules that only audit; high-severity rules without alerts; low-confidence detections; labels not used in DLP; device policies without USB restrictions; no Copilot policy; alerts sent only to SiteAdmin; legacy mail flow rules inspecting sensitive information |
| **Low** | Disabled or stale policies; rules allowing overrides; unused custom SITs; a large number of people holding DLP roles; no label for sites, groups and Teams; no default or mandatory label; retention labels in no policy; record labels deleting without disposition review; no audit retention for DLP activity; eDiscovery cases open for more than two years |

> Observations are prompts for discussion, generated from configuration alone. They do not know the organisation's risk appetite, licensing or plans. Review and agree priorities with the customer before presenting them as recommendations.

### Framework alignment (NCSC, NIST, ISO 27001)

The report links every area of DLP configuration to recognised security frameworks, so findings can be read against the organisation's compliance obligations:

| Framework | Main references used |
|---|---|
| **NCSC** Cyber Assessment Framework v4.0 and 10 Steps | B3 Data security (B3.a Understanding data, B3.b Data in transit, B3.c Stored data, B3.d Mobile data), B6 Staff awareness, C1 Security monitoring, A1 Governance |
| **NIST CSF 2.0** | PR.DS-01 data at rest, PR.DS-02 data in transit, PR.DS-10 data in use, ID.AM-07 data inventory, DE.CM-03 monitoring |
| **NIST SP 800-53 Rev. 5** | AC-4 Information flow enforcement, SC-7(10) Prevent exfiltration, MP-7 Media use, AC-16 Security attributes, SI-4, AU-13 |
| **ISO/IEC 27001:2022** | 8.12 Data leakage prevention, 5.12 Classification, 5.13 Labelling, 5.14 Information transfer, 7.10 Storage media, 8.16 Monitoring |

The related Purview areas add: retention and records (CAF B3.c; CSF ID.AM-08; 800-53 SI-12, MP-6; ISO 5.33, 8.10), insider risk and communication compliance (CAF C1, C2; CSF DE.CM-03, DE.AE-02; 800-53 PM-12, SI-4; ISO 8.16, 5.3) and audit and eDiscovery (CAF C1; CSF PR.PS-04, RS.AN-03; 800-53 AU-11, AU-12; ISO 8.15, 5.28).

References appear in three places:

- **Executive summary: Framework alignment at a glance.** One row per area, one column per framework, and a colour-coded status for the tenant: Supported (green), Partial (amber), Gap (red), Not assessed (grey).
- **Each area section: Framework alignment.** The framework, the reference, **what the framework expects in plain English**, and the tenant status.
- **Appendix D: Framework references.** What each framework is, links to the source, and the status definitions.

Choose which frameworks appear with **Settings > 9** or `-Frameworks` (`NCSC`, `NISTCSF`, `NIST80053`, `ISO27001`, or `NIST`, `ISO`, `All`, `None`). The default is all four.

```powershell
.\Get-M365DlpConfiguration.ps1 -Mode Full -UserPrincipalName admin@contoso.onmicrosoft.com -Frameworks NCSC,ISO27001
```

> The mapping is indicative, not a compliance assessment. Framework outcomes also depend on policies, processes and people the script cannot see. Confirm references against the framework documents before using them in formal assurance work.

### Applying the configuration (last numbered section)

The "Applying the configuration" section (section 24 when every area is included) turns findings into practical guidance. It starts with a recommended order of work, then includes a guidance block **only for the gaps found in that tenant**:

| Guidance | Included when |
|---|---|
| Create or extend DLP policies to all core locations | No policies, or a core location has no enforced policy |
| Move policies from simulation to enforcement | Policies are in simulation or test mode |
| Add policy tips and proportionate actions | Rules only audit, or no rule notifies users |
| Define and publish sensitivity labels | No labels, or labels not published |
| Use sensitivity labels as DLP conditions | Labels exist but no rule uses them |
| Onboard devices and restrict USB, printing and uploads | No enforced device policy, or no USB restriction |
| Protect Microsoft 365 Copilot interactions | No Copilot DLP policy |
| Configure alerts and a review process | No alerting rules, or alerts only to SiteAdmin |
| Tune confidence levels and instance counts | Low-confidence detections |
| Create custom sensitive information types | No custom SITs |
| Consider Exact Data Match | EDM not in use |
| Turn on the unified audit log | Exchange data collected and auditing disabled |
| Apply retention policies to the core locations | Retention and records included, and email, SharePoint, OneDrive or Teams has no retention policy |
| Keep DLP audit records for longer | Audit and eDiscovery included, and no audit log retention policy covers DLP activity |
| Start insider risk management with a Data leaks policy | Risk and compliance included, insider risk available, and no policies exist |

Each block shows **why it applies**, the **framework references** it supports, **numbered steps in the Microsoft Purview portal**, a **PowerShell example** to adapt, **how to check it worked**, and **Microsoft Learn** links.

> The PowerShell examples are starting points. Replace the example names, sensitive information types, label GUIDs and recipients, create policies in simulation mode first, and follow the customer's change process.

### Microsoft Learn links

Each section links to the relevant Microsoft Learn documentation: data loss prevention, policy reference, simulation mode, policy tips, sensitive information types and entity definitions, custom SITs, keyword dictionaries, Exact Data Match, trainable classifiers, sensitivity labels, label encryption, labels for Teams, groups and sites, auto-labelling, labels as DLP conditions, Endpoint DLP and its settings, DLP for Microsoft 365 Copilot, DLP alerts, Activity Explorer, Purview permissions, Exchange mail flow rules, retention, records management, adaptive scopes, disposition, insider risk management, communication compliance, information barriers, auditing solutions, audit log retention policies and eDiscovery.

All links are kept in one table (`$Script:LearnLinks`) near the start of the Word report region of the script, so a page that Microsoft moves is a one-line fix.

### Customising

| To change | Where |
|---|---|
| Wording, branding, logo, extra commentary | Edit the generated document in Word |
| Microsoft Learn links | `$Script:LearnLinks` in the script |
| Section explanations and observation logic | `New-DlpWordReport` in the script (one commented block per section) |
| Framework references | `$Script:FrameworkMap` and `$Script:FrameworkInfo` in the script |
| Configuration guidance | `Add-ImplementationGuidance` in the script |
| How many rows a table shows | The `-MaxRows` value on each `Add-DocTable` call |

### Sharing

- The report describes how sensitive information is protected. Treat it as confidential to the customer.
- Appendix A lists the admin account used and any collection gaps. Remove it if sharing beyond the project team.
- If Activity Explorer was included, the underlying CSV contains user and file names. Do not share it more widely than the report.

Section 6 of the [user guide](M365-DLP-Discovery-Script-User-Guide.docx) has sample screens of every part of the report.

---

## DLP alerts dashboard

`DLP_Alerts_Dashboard.html` answers "what are our DLP alerts actually telling us?". It is one self-contained page: open it in any browser; it needs no internet access and no add-ins.

| Part | What it shows |
|---|---|
| What the alerts are telling you | Plain-English findings: how many alerts and how severe, which rule causes most of them, the information most often found, where it happened, external domains involved, whether a few people account for most alerts (usually a business process rather than widespread risk), how many are still open, and any peak day |
| Headline tiles | Alerts by severity, still open, people involved, alerts with external recipients, and all rule matches |
| Alerts per day | Daily chart, coloured by severity |
| Breakdowns | Rules raising alerts (with their policy), information found, where it happened, who triggered the alerts, external domains, alert status |
| Rule matches with and without alerts | Shows rules that match often but never alert, and rules that alert on most matches |
| Alerts table | Every alert with time, severity, rule, person, location, information found, file or email subject, external domains and status; search box and severity filter |

**Where the data comes from:** the unified audit log, read with `Search-UnifiedAuditLog` over Exchange Online PowerShell (no Microsoft Graph). DLP alerts are the `SecurityComplianceAlerts` records in the DataLossPrevention category. Each alert is linked to the DLP rule-match events (`ComplianceDLPExchange`, `ComplianceDLPSharePoint`, `DLPEndpoint`) recorded for the same rule and person in the hour before it, which is where the information types, file or subject and recipient domains come from. External domains are those not in the tenant's accepted domains.

**How to create it:** menu option 8 (asks for the number of days), `-Mode AlertDashboard`, or the complete report with DLP rule activity included (it is then also summarised in section 16 of the Word report).

**Privacy:** the dashboard shows names, file names and email subjects. Turn on **Settings > 18** or use `-MaskNamesInDashboard` to show people as User 1, User 2 … and hide items.

---

## Permissions

The simplest option is the **Compliance Administrator** role group. For least privilege, use read-only roles:

| Area | Read-only role(s) |
|---|---|
| DLP policies, rules, Endpoint DLP settings, SITs, EDM | View-Only DLP Compliance Management |
| Sensitivity labels and auto-labelling | Information Protection Reader / Sensitivity Label Reader |
| Alert policies | View-Only Manage Alerts |
| Role groups and members | View-Only Role Management / View-Only Recipients |
| Activity Explorer | Activity Explorer Content Viewer and Data Classification List Viewer (e.g. Information Protection Analysts) |
| Exchange Online | Exchange View-Only Organization Management |
| Content Explorer | Content Explorer List Viewer |
| DLP rule activity, alerts and the alerts dashboard (audit log) | View-Only Audit Logs (Exchange Online) or Audit Reader (Purview) |
| Retention policies, retention labels, records, adaptive scopes (area 15) | View-Only Retention Management and View-Only Record Management |
| Insider risk policies (area 16) | Insider Risk Management Admins role group (there is no view-only role for policies) |
| Communication compliance policies and rules (area 16) | Communication Compliance Admins role group (there is no view-only role for policies) |
| Information barrier policies and segments (area 16) | Information Barrier Compliance Management (included in Compliance Administrator) |
| Audit log retention policies (area 17) | Audit Reader or Audit Manager |
| Unified audit log status (areas 12 and 17) | View-Only Audit Logs or Exchange View-Only Organization Management (Exchange Online) |
| eDiscovery cases and holds (area 17) | eDiscovery Manager. Only cases the account is a member of are listed unless it is also an **eDiscovery Administrator** |

Insider risk, communication compliance and information barriers have no view-only role for their policies, so the account needs that feature's own role group. The script still only runs `Get-` commands against them. If you would rather not grant those roles, turn the area off (Settings 20, or `-Scope DLP, Lifecycle, Audit`): the report then leaves the section out instead of showing a gap.

Role changes can take up to an hour to apply. Use **option 3** to confirm what the account can read.

Some features also need licences. DLP for Exchange, SharePoint and OneDrive and manual labelling are in Microsoft 365 E3 and Business Premium. Teams chat DLP, Endpoint DLP, Copilot DLP, auto-labelling, EDM, trainable classifiers, Activity Explorer and Content Explorer need Microsoft 365 E5, E5 Information Protection & Governance, or the Microsoft Purview Suite (formerly E5 Compliance). Retention policies, manual retention labels, Audit (Standard) and eDiscovery holds are in E3. Auto-applied retention labels, adaptive scopes and records management need E5, E5 Information Protection & Governance or Purview Suite; insider risk management, communication compliance and information barriers need E5, E5 Insider Risk Management or Purview Suite; Audit (Premium) with audit log retention policies and eDiscovery (Premium) need E5, E5 eDiscovery & Audit or Purview Suite. The report shows the licence each recommended action needs; it reports configuration, not entitlement, so confirm against [Microsoft's licensing guidance](https://learn.microsoft.com/en-us/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/microsoft-365-security-compliance-licensing-guidance).

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `running scripts is disabled on this system` | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |
| Script will not run after download | `Unblock-File .\Get-M365DlpConfiguration.ps1` |
| `You cannot call a method on a null-valued expression` when connecting | Usually multiple ExchangeOnlineManagement versions, or a sign-in broker issue. Use option 1 → clean reinstall, close PowerShell and open a new window. The script automatically retries with `-DisableWAM` where supported and prints connection diagnostics. Also avoid loading Microsoft Graph, Az, Teams or PnP modules in the same window, and do not use PowerShell ISE. |
| A step shows `WARN ... skipped - Cmdlet(s) not available`, e.g. `Get-AutoSensitivityLabelPolicy` | The account lacks the role for that area, or the tenant is not licensed for the feature. Run option 3, add the role, wait up to an hour and re-run. |
| Option 3 shows FAIL for Activity Explorer | Add the account to Information Protection Analysts (or Investigators). |
| Menu opens when you expected the script to run straight through | Add `-Mode Full`, `-Mode Area` or `-Mode TestAccess`. |
| Word report contents page says "Right-click here and choose Update Field" | Select **Enable Editing** if in Protected View, then right-click the contents and choose **Update Field > Update entire table**. |
| `Word report failed` in the log and no `DLP_Assessment_Report.docx` | The rest of the output is complete. Check `Collection.log`, then re-run option 5. Turn the report off with Settings > 8 if needed. |
| Word report says Activity Explorer or Exchange data was not included | Re-run option 5 and answer Yes to those prompts. |
| No Excel workbook | Install ImportExcel (option 1), or use the CSV files. |
| Content Explorer step fails or shows `n/a` | Add the account to **Content Explorer List Viewer**, wait up to an hour and re-run. Option 3 tests it. |
| Alerts dashboard shows alerts but "What was found" is empty | The rule-match events behind the alerts could not be read or are older than the period. Check View-Only Audit Logs and increase the number of days. |
| DLP rule activity shows no matches | Check the unified audit log is on (section 18), the account has View-Only Audit Logs, and DLP events exist in the period. Audit events can take up to 24 hours to appear. |
| A report section says "Not available" | That part could not be built; the rest of the report is complete. The reason is in `Collection.log` and `99_Collection_Errors`. |
| "Not connected" after an Exchange Online area | Expected. The script disconnects Exchange Online to avoid cmdlet name clashes and reconnects to Security & Compliance automatically. |

---

## Limitations

- **Trainable classifiers:** there is no supported cmdlet that lists all classifiers. Only classifiers referenced in DLP rules are reported. Check the full list in the Purview portal under *Data classification > Classifiers*.
- **Endpoint DLP devices:** device onboarding status is not available through Security & Compliance PowerShell. Check the Purview portal (*Settings > Device onboarding*) or Microsoft Defender.
- **Copilot locations:** Microsoft continues to change how Copilot locations are stored, so detection uses pattern matching. If a Copilot policy shows as not covered, check `11_DLP_Policies_Summary`.
- **Activity Explorer:** large tenants can exceed the page cap. A warning is logged when data is partial.
- **Content Explorer:** counts the sensitive information types used in rules plus a set of common built-in types (up to `-ContentExplorerMaxTags`), not every type. Top locations come from a sample of up to 500 items.
- **DLP rule activity:** limited by audit log retention (180 days with Audit Standard) and `-AuditMaxRecords`. Exchange Online transport-rule DLP is not included.
- **Maturity scores, coverage map and scenarios** are generated from the configuration. They are a starting point for discussion, not a formal compliance assessment.
- **Cell size:** CSV and Excel values are truncated at 32,000 characters. Use the JSON file for the full value.
- **eDiscovery:** only cases visible to the account are reported, and holds are read for up to `-MaxEDiscoveryCases` active cases. Case names are in `310_eDiscovery_Cases` but are kept out of the Word report.
- **Insider risk and communication compliance:** the cmdlets return policy settings, not alerts, cases or user risk scores, so the report describes what is configured, not what has been found.
- **Retention coverage:** a location counts as covered when an enabled retention policy includes it. Adaptive policies count as "selected only", and retention labels are not counted because they apply only to labelled items.
- **Maturity scores** cover DLP only; the related Purview areas have their own status in the framework tables and status by area.

---

## Documentation

A full user guide, including screen examples, sample output for every dataset, sample pages of the Word assessment report and a complete dataset reference, is in [`M365-DLP-Discovery-Script-User-Guide.docx`](M365-DLP-Discovery-Script-User-Guide.docx).
