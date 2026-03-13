# SPO OTP Retirement Impact Assessment v3.1

Toolset for assessing the impact of **SharePoint One-Time Passcode (SPO OTP) retirement** ([MC1243549](https://mc.merill.net/message/MC1243549)) on your Microsoft 365 tenant.

Starting **July 2026**, external users who access SharePoint/OneDrive via SPO OTP and **do not have** a Microsoft Entra B2B guest account will receive **Access Denied**.

## Background: SPO OTP vs Entra B2B Email OTP

These are **two different mechanisms** — understanding the difference is key:

| | SPO OTP (retiring) | Entra B2B Email OTP |
|---|---|---|
| Where it lives | SharePoint layer | Microsoft Entra ID |
| Creates guest account in Entra? | **No** — user exists only in SharePoint external user store | **Yes** — creates B2B guest account |
| Visible in Graph API? | **No** | Yes |
| What happens after retirement | User gets **Access Denied** | Continues working normally |

After SPO OTP retires, all external authentication routes through Entra B2B. If a guest has no Microsoft Account, no federated identity, and **Entra Email OTP is disabled**, they will be forced to create a Microsoft Account to access shared content.

## Contents

| File | Purpose |
|------|---------|
| `Test-SPOOTPDiagnostic.ps1` | **Step 1** — Quick diagnostic (1–2 min). Checks tenant config and estimates impact. |
| `Get-SPOOTPImpactAssessment.ps1` | **Step 2** — Full assessment with detailed CSV report. |

## Requirements

### Modules

| Module | Required | Purpose |
|--------|----------|---------|
| `Microsoft.Graph.Authentication` | **Yes** | Entra ID queries, Email OTP policy check |
| `Microsoft.Online.SharePoint.PowerShell` | Recommended | SPO tenant settings, external user enumeration |
| `ExchangeOnlineManagement` | Optional | Unified Audit Log enrichment for last-activity dates |

> **Note:** `Microsoft.Graph.Users` and other Graph submodules are intentionally **not used**. The SPO module bundles its own version of Graph assemblies, causing version conflicts. All Graph queries use `Invoke-MgGraphRequest` (REST) instead.

### Permissions

| Service | Required Role / Scope |
|---------|----------------------|
| Microsoft Graph | `Policy.Read.All`, `User.Read.All`, `IdentityProvider.Read.All`, `AuditLog.Read.All`, `Sites.Read.All` |
| SharePoint Online | SharePoint Administrator or Global Administrator |
| Exchange Online | Access to Unified Audit Log (optional) |

### PowerShell Version

Compatible with **PowerShell 5.1** and **PowerShell 7.x**.

### Installation

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser   # optional

git clone https://github.com/yourrepo/SPO-OTP-Impact-Assessment.git
```

---

## Step 1: Run the Diagnostic

`Test-SPOOTPDiagnostic.ps1` runs in **1–2 minutes** and answers the key questions:

1. **Is Entra B2B Email OTP enabled?** — If not, guests will be forced to create Microsoft Accounts after retirement.
2. **Is `EnableAzureADB2BIntegration` already on?** — If yes, your tenant already routes external sharing through Entra B2B and the impact is minimal.
3. **What is the tenant's `SharingCapability`?** — Determines the scope of external sharing.
4. **How many SPO external users exist without a B2B guest account?** — The actual number of at-risk users.
5. **Which external identity providers are configured?** — Federation, MSA, Email OTP.

### Usage

```powershell
# Connect to SPO first (recommended for full results)
Connect-SPOService -Url 'https://contoso-admin.sharepoint.com'

# Run diagnostic
.\Test-SPOOTPDiagnostic.ps1 -TenantName 'contoso'
```

If SPO connection is not possible, the diagnostic still checks Entra policies (checks 1 and 2) but cannot count SPO external users.

### Interpreting Results

**If `EnableAzureADB2BIntegration = True`:**
Your tenant already uses Entra B2B for external sharing. Impact is minimal — only users who shared content before B2B integration was enabled and never re-authenticated could be affected. In most cases, no further action is needed.

**If `EnableAzureADB2BIntegration = False` and SPO external users > 0:**
You have at-risk users. Proceed to Step 2 for a detailed report, or enable B2B integration now:

```powershell
Set-SPOTenant -EnableAzureADB2BIntegration $true
```

**If Email OTP is disabled:**
Enable it immediately in Entra admin center → External Identities → All identity providers → Email one-time passcode → **Yes**. Without this, guests without MSA/federation will be unable to authenticate after retirement.

---

## Step 2: Full Assessment

`Get-SPOOTPImpactAssessment.ps1` produces a detailed CSV report of all affected users. Run this only if the diagnostic (Step 1) indicates there are at-risk users.

### Two Operating Modes

| | SPO + Graph (default) | Graph-only (`-GraphOnly`) |
|---|---|---|
| **How it works** | Enumerates external users per SharePoint site via SPO Management Shell, cross-references with Entra B2B guest accounts | Analyzes Entra guest accounts by their identity type (`identities` array) |
| **Data quality** | Full: per-site mapping, `WhenCreated`, site URLs/titles | Partial: identifies at-risk identity types, no per-site mapping |
| **What it finds** | External users in SharePoint who have **no Entra B2B guest account** (true SPO OTP users) | Entra guests with `EmailOTP` identity type (already have B2B account but use email OTP auth) |
| **Requires** | `Microsoft.Online.SharePoint.PowerShell` + SharePoint Admin role | Only `Microsoft.Graph.Authentication` |
| **Estimated runtime** | **30 min – 2+ hours** depending on site count (progress logged every 50 sites) | **1–5 minutes** (queries Entra ID only, no site scanning) |
| **Best for** | Production assessment with full per-site detail | Quick overview when SPO module auth is unavailable, or for very large tenants where per-site scanning is impractical |

> **Important:** Graph-only mode cannot detect true SPO-only OTP users (those who exist only in SharePoint's external user store and have no Entra representation at all). For a complete picture, use SPO mode.

If the SPO module fails to authenticate, the script **automatically falls back** to Graph-only mode and logs a warning.

### Usage

**SPO + Graph mode** (recommended — connect SPO first):

```powershell
Connect-SPOService -Url 'https://contoso-admin.sharepoint.com'
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso'
```

**SPO + Graph with OneDrive sites and 90-day threshold:**

```powershell
Connect-SPOService -Url 'https://contoso-admin.sharepoint.com'
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso' -InactiveDaysThreshold 90 -IncludeOneDriveSites
```

> Including OneDrive sites significantly increases runtime (can add thousands of sites to scan).

**Graph-only mode** (no SPO module needed):

```powershell
.\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso' -GraphOnly
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-TenantName` | String | *(required)* | Tenant name or URL (`contoso`, `contoso.sharepoint.com`, `https://contoso-admin.sharepoint.com`) |
| `-InactiveDaysThreshold` | Int | `180` | Days since last activity to flag as inactive (30–730) |
| `-IncludeInactiveUsers` | Bool | `$true` | Include inactive users in report (flagged as LOW impact) |
| `-OutputPath` | String | `SPO_OTP_Impact_<timestamp>.csv` | CSV report output path |
| `-IncludeOneDriveSites` | Switch | `$false` | Include OneDrive for Business sites (increases runtime) |
| `-GraphOnly` | Switch | `$false` | Use only Graph API, skip SPO module |

### Output

The script generates two files:

```
SPO_OTP_Impact_20260313_112804.csv   # Detailed report
SPO_OTP_Impact_20260313_112804.log   # Full session log with diagnostics
```

**CSV columns:**

| Column | Description |
|--------|-------------|
| `Email` | External user's email address |
| `DisplayName` | Display name |
| `InvitedAs` / `AcceptedAs` | Invitation and acceptance email |
| `WhenCreated` | Account creation date |
| `IsActive` | Whether the user is considered active |
| `LastActivityDate` | Last known activity date |
| `ActivitySource` | Source: `Entra SignInActivity`, `Unified Audit Log`, `SPO WhenCreated`, `M365 SharePoint Usage Report` |
| `IdentityType` | `EmailOTP` (at risk), `Federated` / `MicrosoftAccount` (safe), `Unknown` — Graph-only mode |
| `ExternalUserState` | `Accepted`, `PendingAcceptance` |
| `AccountEnabled` | Whether the Entra account is enabled |
| `SiteCount` | Number of sites with access (SPO mode only) |
| `SiteUrls` / `SiteTitles` | Sites the user can access (SPO mode only) |
| `Impact` | `HIGH - Active user will lose access` or `LOW - Inactive account` |
| `DataSource` | `SPO Module` or `Graph API` |

---

## Remediation

For users identified as **HIGH impact**:

1. **Enable B2B integration** (if not already):
   ```powershell
   Set-SPOTenant -EnableAzureADB2BIntegration $true
   ```
2. **Ensure Email OTP is enabled** in Entra admin center → External Identities → All identity providers → Email one-time passcode → **Yes**. See [Email OTP for B2B guests](https://learn.microsoft.com/entra/external-id/one-time-passcode).
3. **Bulk invite** affected users via Entra admin center or `New-MgInvitation` to create B2B guest accounts proactively.
4. **Re-share content** — have an internal user share or re-share at least one file/folder/site with the external user. This automatically creates a B2B guest account and restores access to all previously shared content.
5. **Communicate** to internal users that some external collaborators may see Access Denied starting July 2026, and that re-sharing resolves it.

## Known Issues

### SPO Module Assembly Conflicts

The `Microsoft.Online.SharePoint.PowerShell` module bundles its own `Microsoft.Graph.Authentication` assemblies. Loading `Microsoft.Graph.Users` or other Graph submodules in the same session causes:

```
Could not load file or assembly 'Microsoft.Graph.Authentication, Version=2.x.x.x'
```

**Both scripts avoid this by design** — all Graph calls use `Invoke-MgGraphRequest` (REST only).

### SPO Module Auth Failures (400 Bad Request)

Common with older SPO module versions or Conditional Access policies. Workarounds:

- Pre-connect manually before running scripts: `Connect-SPOService -Url 'https://tenant-admin.sharepoint.com'`
- Update the module: `Update-Module Microsoft.Online.SharePoint.PowerShell -Force`
- Use `-GraphOnly` switch as fallback

### `Get-SPOExternalUser` Returns 0 with Parameters

Some SPO module versions return 0 results when `-Position` and `-PageSize` parameters are used, but return data without parameters. The diagnostic script handles this by trying multiple retrieval methods automatically.

## Timeline (MC1243549)

| Date | Event |
|------|-------|
| **May 2026** | New external sharing starts using Entra B2B. Existing SPO OTP users keep access. |
| **July 2026** | SPO OTP retirement begins. Users without B2B guest accounts get **Access Denied**. |
| **August 31, 2026** | Retirement expected to complete. |

## Changelog

### v3.1 (2026-03-13)

**Test-SPOOTPDiagnostic.ps1:**
- Added `IdentityProvider.Read.All` scope (fixes `Forbidden` on identity providers endpoint)
- Fixed verbose Graph context dump on reconnect
- Added bare `Get-SPOExternalUser` call (Method 1) — workaround for SPO module paging bug where `-Position`/`-PageSize` returns 0 results
- Added `SharingCapability` analysis with per-value explanation
- 3-method fallback for external user retrieval: bare → paginated → site sampling

**Get-SPOOTPImpactAssessment.ps1:**
- Added persistent file logging (`.log` alongside `.csv` output)
- Fixed `Write-Log` crash on empty string (`[Mandatory]` + empty validation in PS 5.1)
- Added `ExchangeOnlineManagement` module auto-loading and `Connect-ExchangeOnline` for Unified Audit Log enrichment (`Search-UnifiedAuditLog` requires active EXO session)
- Added per-site diagnostic counters (external/B2B/affected per site) in scan log
- Added progress summary every 50 sites
- Graph-only mode: replaced broken `sites/{id}/permissions` API scanning with Entra identity classification (`identities` array analysis)
- Graph-only mode: added M365 SharePoint usage report enrichment for activity data
- Added `IdentityType`, `ExternalUserState`, `AccountEnabled` columns to CSV output

**README.md:**
- Complete rewrite with diagnostic-first workflow (Step 1 → Step 2)
- Added SPO OTP vs Entra B2B Email OTP background explanation
- Added SPO vs Graph-only mode comparison table with runtime estimates
- Added `SharingCapability` and `EnableAzureADB2BIntegration` interpretation guide
- Documented `Get-SPOExternalUser` paging bug workaround

### v3.0 (2026-03-13)

- Complete rewrite of `Get-SPOOTPImpactAssessment.ps1` for PowerShell 5.1 compatibility (removed `??`, ternary operators)
- Replaced `Get-MgUser` with `Invoke-MgGraphRequest` (REST) to avoid assembly conflicts with SPO module's bundled Graph assemblies
- Added `-GraphOnly` switch with automatic fallback when SPO authentication fails
- Multi-strategy SPO connection (standard → `-Browser` → auto-fallback to Graph-only)
- Flexible `-TenantName` parsing (URL, admin URL, bare name, sovereign clouds)
- Removed `#Requires -Modules` to prevent assembly load conflicts at startup
- Initial release of `Test-SPOOTPDiagnostic.ps1`

## License

MIT

## Author

**Petr Beloch**
