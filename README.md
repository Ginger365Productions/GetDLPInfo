# Get-M365DlpConfiguration

A read-only PowerShell tool for discovering and documenting **Microsoft Purview Data Loss Prevention (DLP)** configuration in any Microsoft 365 tenant.

It is built for consultants and administrators who need a consistent evidence pack at the start of a DLP engagement. Run it the same way in every tenant, get the same set of outputs, and compare tenants or runs over time.

> **Read-only:** the script only calls `Get-*` and `Export-*` cmdlets. It never creates, changes, enables or deletes anything in the tenant. The only files it writes are its output files on your workstation.

---

## Contents

- [What it collects](#what-it-collects)
- [Prerequisites](#prerequisites)
- [Quick start](#quick-start)
- [The menu](#the-menu)
- [Running without the menu](#running-without-the-menu)
- [Report areas](#report-areas)
- [Output](#output)
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
| Sensitivity labels | Labels, label policies, auto-labelling policies and rules, and labels used as DLP conditions |
| Endpoint DLP | Global settings (browsers, service domains, printer, removable media and network share groups) and per-rule device restrictions (USB, printing, browser and cloud upload, clipboard) |
| Alerts | DLP alert policies, and rules that raise alerts or incident reports |
| Ownership | Purview role groups and members holding DLP-relevant roles |
| Activity Explorer *(optional)* | DLP rule matches and other activity for the last 1–30 days |
| Exchange Online *(optional)* | Mail flow rules (including legacy Exchange DLP) and unified audit log status |

A complete run also produces an **assessment summary**: one row per DLP topic with an automatically calculated finding, the file that holds the evidence, a suggested next activity, and blank Status / Owner / TargetDate / Notes columns ready to use as an engagement tracker.

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
--------------------------------------------------------------------------------
   1. Install / update required PowerShell modules
   2. Connect to tenant
   3. Test credentials and tenant access
   4. Report on individual DLP items
   5. Complete tenant DLP report
   6. Settings
   7. Disconnect / switch tenant
   Q. Quit
```

| Option | Description |
|---|---|
| **1. Install / update modules** | Shows installed versions, installs or updates ExchangeOnlineManagement and ImportExcel for the current user (no admin rights needed), and offers a **clean reinstall** that removes every version first when more than one is found. |
| **2. Connect to tenant** | Signs in to Security & Compliance PowerShell, or switches to a different account or tenant. Shows the account, tenant ID and organisation, and how many key DLP cmdlets the account can use. |
| **3. Test credentials and tenant access** | Runs a minimal read-only call against every cmdlet the reports use and shows **PASS**, **MISSING** (role or licence missing) or **FAIL** per area. Lists the account's Purview role groups and saves the results to CSV. Optionally tests Exchange Online. |
| **4. Report on individual DLP items** | Choose one or more of the 12 [report areas](#report-areas), for example `2,3,8`. |
| **5. Complete tenant DLP report** | Every area plus the assessment summary (CSV and HTML) and Excel workbook. Asks whether to include Activity Explorer and Exchange Online. |
| **6. Settings** | Output folder, admin account, Activity Explorer days and page cap, Exchange Online, and sign-in with `-DisableWAM`. |
| **7. Disconnect / switch tenant** | Signs out so you can move to the next customer tenant without restarting. |

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

# Selected areas only (policies, rules, Endpoint DLP)
.\Get-M365DlpConfiguration.ps1 -Mode Area -Area 2,3,8 -UserPrincipalName admin@contoso.onmicrosoft.com

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
| `-Mode` | `Menu` | `Menu`, `InstallModules`, `TestAccess`, `Full` or `Area` |
| `-Area` | | With `-Mode Area`: area numbers 1–12 or part of an area name |
| `-OutputRoot` | Documents | Folder in which the time-stamped output folder is created |
| `-UserPrincipalName` | | Admin account to sign in with |
| `-IncludeActivityExplorer` | Off | Include Activity Explorer in the complete report |
| `-ActivityDays` | `7` | Days of Activity Explorer data (1–30) |
| `-ActivityMaxPages` | `20` | Cap on 5,000-record Activity Explorer pages (1–200) |
| `-IncludeExchange` | Off | Include Exchange Online in the complete report and access test |
| `-DisableWAM` | Off | Sign in through the browser instead of the Windows account broker |
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
| `00_Assessment_Summary.html` / `.csv` | Complete report only. One row per topic with findings and tracker columns |
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

Role changes can take up to an hour to apply. Use **option 3** to confirm what the account can read.

Some features also need licences. For example, auto-labelling, Endpoint DLP, Teams DLP, EDM and Copilot DLP generally need Microsoft 365 E5, E5 Compliance or the relevant add-on. The script reports configuration, not entitlement.

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
| No Excel workbook | Install ImportExcel (option 1), or use the CSV files. |
| "Not connected" after an Exchange Online area | Expected. The script disconnects Exchange Online to avoid cmdlet name clashes and reconnects to Security & Compliance automatically. |

---

## Limitations

- **Trainable classifiers:** there is no supported cmdlet that lists all classifiers. Only classifiers referenced in DLP rules are reported. Check the full list in the Purview portal under *Data classification > Classifiers*.
- **Endpoint DLP devices:** device onboarding status is not available through Security & Compliance PowerShell. Check the Purview portal (*Settings > Device onboarding*) or Microsoft Defender.
- **Copilot locations:** Microsoft continues to change how Copilot locations are stored, so detection uses pattern matching. If a Copilot policy shows as not covered, check `11_DLP_Policies_Summary`.
- **Activity Explorer:** large tenants can exceed the page cap. A warning is logged when data is partial.
- **Cell size:** CSV and Excel values are truncated at 32,000 characters. Use the JSON file for the full value.

---

## Documentation

A full user guide, including screen examples, sample output for every dataset and a complete dataset reference, is in [`M365-DLP-Discovery-Script-User-Guide.docx`](M365-DLP-Discovery-Script-User-Guide.docx).
