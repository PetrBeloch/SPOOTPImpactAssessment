# Get-SPOOTPImpactAssessment

Assesses the impact of **SharePoint One-Time Passcode (SPO OTP) retirement** ([MC1243549](https://mc.merill.net/message/MC1243549)) on your Microsoft 365 tenant.

Starting **July 2026**, external users who access SharePoint/OneDrive via SPO OTP and **do not have** a Microsoft Entra B2B guest account will receive **Access Denied**. This script identifies those users before the cutoff so you can take action proactively.

## What It Does

1. Retrieves all **Entra ID B2B guest accounts** (these users are safe — not affected)
2. Enumerates **external users** across all SharePoint site collections
3. Cross-references the two lists to find external users **without** a B2B guest account
4. Filters out **inactive/stale accounts** based on a configurable threshold
5. Maps each affected user to the **specific sites** they have access to
6. Enriches activity data from **Unified Audit Log** (if available)
7. Exports a **CSV report** with impact classification (HIGH / LOW)

## Operating Modes

| Mode | How | Pros | Cons |
|------|-----|------|------|
| **SPO + Graph** (default) | SPO Management Shell + Graph REST API | Most detailed data, includes `WhenCreated`, per-site external user lists | Requires SPO module auth to work |
| **Graph-only** (`-GraphOnly`) | Microsoft Graph API only | No SPO module dependency, works everywhere | Less detailed per-site data |

If SPO module authentication fails, the script **automatically falls back** to Graph-only mode.

## Requirements

### Modules

| Module | Required | Purpose |
|--------|----------|---------|
| `Microsoft.Graph.Authentication` | **Yes** | All Entra ID / Graph API queries |
| `Microsoft.Online.SharePoint.PowerShell` | No (recommended) | SPO site/external user enumeration |
| `ExchangeOnlineManagement` | No (recommended) | Unified Audit Log enrichment for accurate last-activity dates |

> **Note:** `Microsoft.Graph.Users` and `Microsoft.Graph.Reports` are intentionally **not used**. The SPO module bundles its own version of Graph assemblies, causing version conflicts. All Graph queries use `Invoke-MgGraphRequest` (REST) instead.

### Permissions

| Service | Role / Scope |
|---------|-------------|
| SharePoint Online | SharePoint Administrator (for SPO mode) |
| Microsoft Graph | `User.Read.All`, `AuditLog.Read.All`, `Sites.Read.All` |
| Exchange Online | Access to Unified Audit Log (optional, for activity enrichment) |

### PowerShell Version

- **PowerShell 5.1** — fully compatible
- **PowerShell 7.x** — fully compatible

## Installation

```powershell
# Install required modules (if not already present)
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser  # optional
Install-Module ExchangeOnlineManagement -Scope CurrentUser                # optional, for audit log

# Clone or download the script
git clone https://github.com/yourrepo/Get-SPOOTPImpactAssessment.git
```

## Usage

### Basic (auto-detects best mode)

```powershell
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso'
```

### Graph-only mode (skip SPO module entirely)

```powershell
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso' -GraphOnly
```

### Full scan with OneDrive sites and 90-day inactivity threshold

```powershell
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso' -InactiveDaysThreshold 90 -IncludeOneDriveSites
```

### Pre-connect SPO manually (recommended if SPO auth fails)

```powershell
Connect-SPOService -Url 'https://contoso-admin.sharepoint.com'
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso'
```

### Flexible tenant name input

The `-TenantName` parameter accepts any of these formats:

```
contoso
contoso.sharepoint.com
https://contoso.sharepoint.com
https://contoso-admin.sharepoint.com
contoso.sharepoint.us          # sovereign clouds
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-TenantName` | String | *(required)* | SharePoint tenant name or URL |
| `-InactiveDaysThreshold` | Int | `180` | Days since last activity to flag as inactive (30–730) |
| `-IncludeInactiveUsers` | Bool | `$true` | Include inactive users in report (flagged as LOW impact) |
| `-OutputPath` | String | `SPO_OTP_Impact_<timestamp>.csv` | CSV output file path |
| `-IncludeOneDriveSites` | Switch | `$false` | Include OneDrive for Business sites in scan |
| `-GraphOnly` | Switch | `$false` | Use only Graph API (no SPO module required) |

## Output

### CSV Report Columns

| Column | Description |
|--------|-------------|
| `Email` | External user's email address |
| `DisplayName` | Display name |
| `InvitedAs` | Original invitation email |
| `AcceptedAs` | Email used to accept the invitation |
| `WhenCreated` | Account creation date (SPO mode only) |
| `IsActive` | Whether the user is considered active |
| `LastActivityDate` | Last known activity date |
| `ActivitySource` | Where the activity date came from (`Unified Audit Log`, `SPO WhenCreated`, etc.) |
| `HasB2BAccount` | Always `False` (affected users by definition don't have one) |
| `SiteCount` | Number of sites the user has access to |
| `SiteUrls` | Semicolon-separated list of site URLs |
| `SiteTitles` | Semicolon-separated list of site titles |
| `Impact` | `HIGH - Active user will lose access` or `LOW - Inactive account` |
| `DataSource` | `SPO Module` or `Graph API` |

### Console Summary

```
==========================================
         ASSESSMENT SUMMARY
==========================================
Mode:                                   SPO + Graph
Total external users found:             142
Already have B2B guest account (safe):  98
Affected by SPO OTP retirement:         44
  - Active (HIGH impact):               31
  - Inactive (LOW impact):              13
Sites with affected users:              17
==========================================
```

### Logging

The script writes a persistent log file alongside the CSV output (`SPO_OTP_Impact_<timestamp>.log`). The log captures the full session including connection attempts, per-site scan results, progress checkpoints every 50 sites, and any errors. Useful for debugging long-running scans or sharing results with colleagues.

```
SPO_OTP_Impact_20260313_112804.csv   # Report
SPO_OTP_Impact_20260313_112804.log   # Full session log
```

## Remediation Options

For users identified as **HIGH impact**, you have several options before the July 2026 deadline:

1. **Bulk invite via Entra admin center** — Create B2B guest accounts manually or via `New-MgInvitation`
2. **Re-share content** — Have an internal user share or re-share at least one file/folder/site with the external user. This automatically creates a B2B guest account.
3. **Enable Email OTP in Entra External ID** — Ensure email one-time passcode is **not disabled** in Entra External ID settings as a fallback authentication method. See [Email OTP for B2B guests](https://learn.microsoft.com/entra/external-id/one-time-passcode).

## Known Issues

### SPO Module Assembly Conflicts

The `Microsoft.Online.SharePoint.PowerShell` module bundles its own version of `Microsoft.Graph.Authentication` assemblies. If you have `Microsoft.Graph.Users` or other Graph submodules loaded in the same session, you may see:

```
Could not load file or assembly 'Microsoft.Graph.Authentication, Version=2.x.x.x'
```

**This script avoids this by design** — it uses only `Invoke-MgGraphRequest` for all Graph calls and never imports `Microsoft.Graph.Users`.

### SPO Module Auth Failures (400 Bad Request)

Common with older SPO module versions or Conditional Access policies blocking legacy auth. Workarounds:

- Update the module: `Update-Module Microsoft.Online.SharePoint.PowerShell -Force`
- Use `-GraphOnly` switch
- Pre-connect manually: `Connect-SPOService -Url 'https://tenant-admin.sharepoint.com'`

## Timeline Reference (MC1243549)

| Date | Event |
|------|-------|
| **May 2026** | New external sharing starts using Entra B2B. Existing SPO OTP users keep access. |
| **July 2026** | SPO OTP retirement begins. Users without B2B guest accounts get **Access Denied**. |
| **August 31, 2026** | Retirement expected to complete. |

## License

MIT

## Author

**Petr Beloch**
