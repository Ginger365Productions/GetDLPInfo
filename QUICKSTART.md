# Quick start: M365 DLP discovery script

This guide gets you from download to a finished DLP audit report in about 15 minutes. For everything else, see the [README](README.md) and the full user guide (`M365-DLP-Discovery-Script-User-Guide.docx`).

---

## What the tool does

`Get-M365DlpConfiguration.ps1` connects to a Microsoft 365 tenant and collects the **Microsoft Purview Data Loss Prevention (DLP)** configuration: policies, rules, sensitive information types, sensitivity labels, Endpoint DLP, alerts and who manages them. It then produces:

- **`DLP_Assessment_Report.docx`**: a plain-English Word report. It covers what is configured, how mature the setup is, what happens when someone shares sensitive data, which problems were found, and a prioritised action plan.
- **`DLP_Alerts_Dashboard.html`** *(optional)*: an easy-to-read page showing what the DLP alerts are about.
- **`DLP_Configuration.xlsx`** plus CSV and JSON files: all the raw data.

It is **read-only**: it never changes anything in the tenant, and it does not use Microsoft Graph.

> **Why are there `Set-` and `New-` commands in the code?** They are example commands that the script writes into section 21 of the Word report, for the customer to review. The script never runs them. The only tenant commands it runs are `Get-*`, `Export-ActivityExplorerData`, `Export-ContentExplorerData` and `Search-UnifiedAuditLog`, which only read data. The only things it changes are files in its output folder on your PC, plus PowerShell modules if you choose to install them in option 1. Full details: [Is it really read-only?](README.md#is-it-really-read-only)

---

## Before you start

| You need | Details |
|---|---|
| A Windows PC | Windows PowerShell 5.1 or PowerShell 7. Do not use PowerShell ISE. |
| An admin account in the tenant | The simplest option is the **Compliance Administrator** role. For read-only roles, see the README. |
| Internet access | To install the PowerShell modules and sign in. |

---

## Step 1: Download and open the script

1. Download the repository from GitHub (**Code > Download ZIP**) and extract it, or clone it.
2. Open **PowerShell** (not as administrator) in the extracted folder.
3. Allow the script to run in this window and unblock the download:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
Unblock-File .\Get-M365DlpConfiguration.ps1
```

4. Start the menu:

```powershell
.\Get-M365DlpConfiguration.ps1
```

```
================================================================================
   Microsoft Purview DLP Discovery  (read-only)
================================================================================
   Not connected
--------------------------------------------------------------------------------
   1. Install / update required PowerShell modules
   2. Connect to tenant
   3. Test credentials and tenant access
   4. Report on individual DLP items
   5. Complete tenant DLP report
   6. Settings
   7. Disconnect / switch tenant
   8. DLP alerts dashboard (what the DLP alerts are about)
   Q. Quit
```

---

## Step 2: Check the modules (option 1)

Press **1**. The script shows which modules are installed and offers to install or update them for your account. You don't need admin rights.

| Module | Needed for |
|---|---|
| **ExchangeOnlineManagement** (3.2.0 or later) | Required: connects to the tenant |
| **ImportExcel** | Optional: creates the Excel workbook |

If more than one version of ExchangeOnlineManagement is found, choose the **clean reinstall**. Then **close PowerShell, open a new window** and start the script again.

---

## Step 3: Connect to the tenant (option 2)

1. Press **2**.
2. Enter the admin account, for example `admin@contoso.onmicrosoft.com`.
3. Sign in and complete MFA in the window that opens.
4. Check the screen shows the right **account** and **tenant**.

The menu header now shows `Connected: <account> | Tenant: <tenant ID>`.

**Optional:** press **3** to test access. Each area shows **PASS**, **MISSING** (role or licence missing) or **FAIL**. Fix any gaps that matter before running the report.

---

## Step 4: Create the full audit report (option 5)

Press **5**. You are asked a few questions; press **Enter** to accept the default shown in brackets.

| Question | Suggested answer |
|---|---|
| Include Activity Explorer data? | **Y** if you want recent activity (needs Activity Explorer roles) |
| Include Exchange Online mail flow rules and audit config? | **Y** |
| Include Content Explorer (where sensitive data is stored)? | **Y** for a fuller picture (needs the Content Explorer List Viewer role; adds a few minutes) |
| Include DLP rule activity and alerts, with the alerts dashboard? | **Y** (needs View-Only Audit Logs; signs in to Exchange Online, so you may be asked to sign in again) |
| Show changes since the previous report? | Only asked when an earlier report for the same tenant exists. **Y** to include a "what has changed" section |

The script works through each area and shows progress. Green = done, yellow = skipped or no data, red = failed. A typical run takes 5 to 20 minutes, depending on the size of the tenant and the options chosen.

When it finishes, it shows the output folder, for example:

```
Output: C:\Users\you\Documents\M365-DLP-Full-20261120-104512
Word report: ...\DLP_Assessment_Report.docx
Alerts dashboard: ...\DLP_Alerts_Dashboard.html
```

---

## Step 5: Open the results

1. Open **`DLP_Assessment_Report.docx`**. If Word asks to update fields, choose **Yes** so the contents page fills in. If it opens in Protected View, select **Enable Editing** first.
2. Start with **section 2 (Executive summary)** for the maturity score and key observations, then **section 20** for the action plan.
3. Open **`DLP_Alerts_Dashboard.html`** in a browser to see what the DLP alerts are about.
4. Check **`99_Collection_Errors.csv`**. Anything listed there was not collected, usually because a role is missing.

> The output contains details of the customer's configuration and, in some files, user names. Store and share it in line with the customer's data handling requirements.

When you have finished, press **7** to disconnect (or **Q** to quit).

---

## Optional: brand the report

Before pressing 5, open **Settings (6)** and set:

- **14:** customer name
- **15:** assessor name (you or your company)
- **16:** a PNG or JPEG logo

These appear on the report cover.

---

## If something goes wrong

| Problem | Fix |
|---|---|
| `running scripts is disabled on this system` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in the same window. |
| `You cannot call a method on a null-valued expression` when connecting | Option 1 > clean reinstall, then close PowerShell and open a new window. In Settings, turn on **7 (Sign in with -DisableWAM)**. |
| A step says `skipped - Cmdlet(s) not available` | The account is missing a role, or the tenant isn't licensed for that feature. Use option 3 to see which. Role changes can take up to an hour. |
| The report doesn't include new features after an update | Close PowerShell completely and start the script in a new window. |

Still stuck? Send `Collection.log` and `99_Collection_Errors.csv` from the output folder to the script owner.
