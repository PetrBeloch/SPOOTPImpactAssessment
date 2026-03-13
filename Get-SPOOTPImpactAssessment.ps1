#Requires -Version 5.1

<#
.SYNOPSIS
    Assesses impact of SharePoint One-Time Passcode (SPO OTP) retirement (MC1243549).

.DESCRIPTION
    Identifies external users who access SharePoint/OneDrive via SPO OTP authentication
    and do NOT have a Microsoft Entra B2B guest account. These users will lose access
    after July 2026 when SPO OTP is retired.

    Two operating modes:
    - Default: Uses SPO Management Shell + Graph REST API (most detailed)
    - GraphOnly: Uses only Microsoft Graph (no SPO module needed, works everywhere)

    The script:
    - Enumerates external users across SharePoint site collections
    - Cross-references with Entra ID B2B guest accounts
    - Filters out stale/inactive accounts based on configurable thresholds
    - Maps each affected user to sites they access
    - Exports a detailed CSV report and console summary

.PARAMETER TenantName
    SharePoint Online tenant name. Accepts flexible formats:
    'contoso', 'contoso.sharepoint.com', 'https://contoso.sharepoint.com',
    'https://contoso-admin.sharepoint.com'

.PARAMETER InactiveDaysThreshold
    Days since last activity to flag account as inactive. Default: 180.

.PARAMETER IncludeInactiveUsers
    Include inactive users in report (flagged as LOW impact). Default: true.

.PARAMETER OutputPath
    CSV report output path. Default: script directory with timestamp.

.PARAMETER IncludeOneDriveSites
    Include personal OneDrive sites in scan. Increases runtime. Default: false.

.PARAMETER GraphOnly
    Use only Microsoft Graph API (no SPO Management Shell required).
    Less detailed per-site external user data but avoids SPO module auth issues.

.EXAMPLE
    .\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso'

.EXAMPLE
    .\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso' -GraphOnly

.EXAMPLE
    .\Get-SPOOTPImpactAssessment.ps1 -TenantName 'contoso' -InactiveDaysThreshold 90 -IncludeOneDriveSites

.NOTES
    Author:         Petr Beloch
    Date:           2026-03-13
    Reference:      MC1243549 - Retirement of SharePoint OTP
    Modules:        Microsoft.Graph.Authentication (required)
                    Microsoft.Online.SharePoint.PowerShell (optional, used when -GraphOnly is not set)
    Permissions:    SharePoint Administrator (for SPO mode), Global Reader or equivalent
    Graph Scopes:   User.Read.All, AuditLog.Read.All, Sites.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [Alias('Tenant', 'TenantUrl')]
    [string]$TenantName,

    [ValidateRange(30, 730)]
    [int]$InactiveDaysThreshold = 180,

    [bool]$IncludeInactiveUsers = $true,

    [string]$OutputPath,

    [switch]$IncludeOneDriveSites,

    [switch]$GraphOnly
)

#region Configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Continue'

# Normalize TenantName - accept any common format
$cleanTenant = $TenantName.Trim().TrimEnd('/')
$cleanTenant = $cleanTenant -replace '^https?://', ''

if ($cleanTenant -match '^([a-zA-Z0-9\-]+?)(?:-admin)?\.sharepoint\.') {
    $cleanTenant = $Matches[1]
}
elseif ($cleanTenant -match '^[a-zA-Z0-9\-]+$') {
    # Already just the tenant name
}
else {
    throw "Cannot parse tenant name from '$TenantName'. Use the tenant prefix (e.g., 'contoso') or full URL."
}

# Detect sovereign cloud suffix
$spoSuffix = 'sharepoint.com'
if ($TenantName -match 'sharepoint\.(us|de|cn)') {
    $spoSuffix = "sharepoint.$($Matches[1])"
}

$spoAdminUrl = "https://$cleanTenant-admin.$spoSuffix"
$spoTenantUrl = "https://$cleanTenant.$spoSuffix"
$graphScopes = @('User.Read.All', 'AuditLog.Read.All', 'Sites.Read.All')
$spoExternalUserPageSize = 50
$activityCutoffDate = (Get-Date).AddDays(-$InactiveDaysThreshold)

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputPath = Join-Path -Path $scriptDir -ChildPath "SPO_OTP_Impact_$timestamp.csv"
}
#endregion

#region Helper Functions
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    switch ($Level) {
        'Info'    { Write-Host "[$ts] [Info] $Message" -ForegroundColor Cyan }
        'Warning' { Write-Host "[$ts] [Warning] $Message" -ForegroundColor Yellow }
        'Error'   { Write-Host "[$ts] [Error] $Message" -ForegroundColor Red }
        'Success' { Write-Host "[$ts] [Success] $Message" -ForegroundColor Green }
    }
}

function Coalesce {
    <# PS 5.1 compatible null-coalescing. Returns first non-null/non-empty value. #>
    [CmdletBinding()]
    param([Parameter(ValueFromRemainingArguments)][object[]]$Values)
    foreach ($v in $Values) {
        if ($null -ne $v -and [string]$v -ne '') { return $v }
    }
    return $null
}

function Initialize-Modules {
    [CmdletBinding()]
    param([switch]$SkipSPO)

    if (-not $SkipSPO) {
        # Load SPO module first (it bundles its own Graph assemblies)
        $spoMod = Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue
        if (-not $spoMod) {
            $spoInstalled = Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -ListAvailable -ErrorAction SilentlyContinue |
                Sort-Object Version -Descending | Select-Object -First 1
            if (-not $spoInstalled) {
                Write-Log 'Microsoft.Online.SharePoint.PowerShell not installed. Switching to Graph-only mode.' -Level Warning
                $script:ForceGraphOnly = $true
                return
            }
            Import-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -ErrorAction Stop
            Write-Log "Loaded Microsoft.Online.SharePoint.PowerShell v$($spoInstalled.Version)." -Level Success
        }
        else {
            Write-Log "Microsoft.Online.SharePoint.PowerShell v$($spoMod.Version) already loaded."
        }
    }

    # Load Graph Authentication
    $graphMod = Get-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue
    if (-not $graphMod) {
        try {
            Import-Module -Name 'Microsoft.Graph.Authentication' -Force -ErrorAction Stop
            $graphMod = Get-Module -Name 'Microsoft.Graph.Authentication'
            Write-Log "Loaded Microsoft.Graph.Authentication v$($graphMod.Version)." -Level Success
        }
        catch {
            Write-Log "Cannot load Microsoft.Graph.Authentication: $($_.Exception.Message)" -Level Error
            throw 'Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
        }
    }
    else {
        Write-Log "Microsoft.Graph.Authentication v$($graphMod.Version) already loaded."
    }

    if (-not (Get-Command -Name 'Invoke-MgGraphRequest' -ErrorAction SilentlyContinue)) {
        throw 'Invoke-MgGraphRequest not available. Graph module may not be loaded correctly.'
    }
}

function Connect-Services {
    [CmdletBinding()]
    param([switch]$SkipSPO)

    $skipSpoConnect = $SkipSPO -or $script:ForceGraphOnly

    # --- SharePoint Online ---
    if (-not $skipSpoConnect) {
        Write-Log "Connecting to SharePoint Online: $spoAdminUrl"

        $spoMod = Get-Module -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue
        if ($spoMod) { Write-Log "  SPO module version: $($spoMod.Version)" }

        $spoConnected = $false

        # Check existing connection
        try {
            $null = Get-SPOSite -Identity $spoTenantUrl -ErrorAction Stop
            Write-Log 'SPO: Already connected.' -Level Success
            $spoConnected = $true
        }
        catch { }

        if (-not $spoConnected) {
            try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch { }

            # Strategy 1: Standard
            try {
                Connect-SPOService -Url $spoAdminUrl -ErrorAction Stop
                Write-Log 'SPO: Connected.' -Level Success
                $spoConnected = $true
            }
            catch {
                Write-Log "SPO strategy 1 (standard) failed: $($_.Exception.Message)" -Level Warning
            }

            # Strategy 2: Browser auth (SPO module 16.0.24322+)
            if (-not $spoConnected) {
                $cmdParams = (Get-Command Connect-SPOService -ErrorAction SilentlyContinue).Parameters
                if ($cmdParams -and $cmdParams.ContainsKey('Browser')) {
                    try {
                        Connect-SPOService -Url $spoAdminUrl -Browser -ErrorAction Stop
                        Write-Log 'SPO: Connected via browser.' -Level Success
                        $spoConnected = $true
                    }
                    catch {
                        Write-Log "SPO strategy 2 (browser) failed: $($_.Exception.Message)" -Level Warning
                    }
                }
            }

            if (-not $spoConnected) {
                Write-Log '' -Level Warning
                Write-Log '====== SPO CONNECTION FAILED ======' -Level Warning
                Write-Log 'Auto-switching to Graph-only mode. Results may be less detailed.' -Level Warning
                Write-Log '' -Level Warning
                Write-Log 'To use full SPO mode next time, connect manually BEFORE running:' -Level Warning
                Write-Log "  Connect-SPOService -Url '$spoAdminUrl'" -Level Warning
                Write-Log '' -Level Warning
                Write-Log 'Other fixes to try:' -Level Warning
                Write-Log '  Update-Module Microsoft.Online.SharePoint.PowerShell -Force' -Level Warning
                Write-Log '  Install-Module PnP.PowerShell -Scope CurrentUser' -Level Warning
                Write-Log '===================================' -Level Warning
                $script:ForceGraphOnly = $true
            }
        }
    }

    # --- Microsoft Graph ---
    Write-Log 'Connecting to Microsoft Graph...'
    try {
        $ctx = Get-MgContext
        if ($null -eq $ctx) { throw 'No session' }

        $missing = $graphScopes | Where-Object { $_ -notin $ctx.Scopes }
        if ($missing) {
            Write-Log "Missing scopes: $($missing -join ', '). Reconnecting..." -Level Warning
            throw 'Insufficient scopes'
        }
        Write-Log "Graph: Connected as $($ctx.Account)." -Level Success
    }
    catch {
        try {
            Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
            $ctx = Get-MgContext
            Write-Log "Graph: Connected as $($ctx.Account)." -Level Success
        }
        catch {
            Write-Log "Graph connection failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
}

function Get-EntraGuestAccounts {
    <# Retrieves all B2B guest accounts from Entra ID via Graph REST API. #>
    [CmdletBinding()]
    param()

    Write-Log 'Retrieving Entra ID B2B guest accounts...'
    $guests = @{}
    $count = 0

    try {
        $select = 'id,displayName,mail,otherMails,userPrincipalName,createdDateTime,signInActivity,accountEnabled,externalUserState'
        $uri = "v1.0/users?`$filter=userType eq 'Guest'&`$select=$select&`$top=999&`$count=true"

        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' } -ErrorAction Stop

            foreach ($g in $response.value) {
                $count++
                $emails = @()
                if ($g.mail) { $emails += $g.mail.ToLower() }
                if ($g.otherMails) {
                    foreach ($m in $g.otherMails) { $emails += $m.ToLower() }
                }

                # Extract email from UPN: user_domain.com#EXT#@tenant.onmicrosoft.com
                if ($g.userPrincipalName -match '^(.+)#EXT#@') {
                    $upnEmail = ($Matches[1] -replace '_', '@')
                    if (($upnEmail -split '@').Count -eq 2) {
                        $emails += $upnEmail.ToLower()
                    }
                }

                # Parse last sign-in
                $lastSignIn = $null
                if ($g.signInActivity) {
                    $candidates = @()
                    if ($g.signInActivity.lastSignInDateTime) {
                        try { $candidates += [datetime]$g.signInActivity.lastSignInDateTime } catch { }
                    }
                    if ($g.signInActivity.lastNonInteractiveSignInDateTime) {
                        try { $candidates += [datetime]$g.signInActivity.lastNonInteractiveSignInDateTime } catch { }
                    }
                    if ($candidates.Count -gt 0) {
                        $lastSignIn = ($candidates | Sort-Object -Descending | Select-Object -First 1)
                    }
                }

                $info = [PSCustomObject]@{
                    Id                = $g.id
                    DisplayName       = $g.displayName
                    Mail              = $g.mail
                    UPN               = $g.userPrincipalName
                    AccountEnabled    = $g.accountEnabled
                    ExternalUserState = $g.externalUserState
                    CreatedDateTime   = $g.createdDateTime
                    LastSignIn        = $lastSignIn
                }

                foreach ($e in ($emails | Select-Object -Unique)) {
                    $guests[$e] = $info
                }
            }

            $uri = $response.'@odata.nextLink'
        } while ($uri)

        Write-Log "Found $count Entra B2B guest accounts." -Level Success
    }
    catch {
        Write-Log "Error retrieving guest accounts: $($_.Exception.Message)" -Level Error
        throw
    }

    return $guests
}

function Get-SitesSPO {
    <# Gets sites via SPO Management Shell. #>
    [CmdletBinding()]
    param([switch]$IncludeOneDrive)

    Write-Log 'Retrieving sites via SPO module...'
    $sites = [System.Collections.Generic.List[object]]::new()

    $spoSites = Get-SPOSite -Limit All -IncludePersonalSite:$false |
        Where-Object {
            $_.Status -eq 'Active' -and
            $_.Template -notlike 'SRCHCEN*' -and
            $_.Template -notlike 'SPSMSITEHOST*' -and
            $_.Template -notlike 'APPCATALOG*' -and
            $_.Template -notlike 'POINTPUBLISHINGHUB*' -and
            $_.Template -notlike 'RedirectSite*'
        }

    foreach ($s in $spoSites) {
        $sites.Add([PSCustomObject]@{ Url = $s.Url; Title = $s.Title; SiteId = $null })
    }
    Write-Log "Found $($spoSites.Count) standard sites."

    if ($IncludeOneDrive) {
        Write-Log 'Retrieving OneDrive sites...'
        $odSites = Get-SPOSite -Limit All -IncludePersonalSite $true -Filter "Url -like '-my.sharepoint.com/personal'" |
            Where-Object { $_.Status -eq 'Active' }
        foreach ($s in $odSites) {
            $sites.Add([PSCustomObject]@{ Url = $s.Url; Title = $s.Title; SiteId = $null })
        }
        Write-Log "Found $($odSites.Count) OneDrive sites."
    }

    return $sites
}

function Get-SitesGraph {
    <# Gets sites via Microsoft Graph REST API. #>
    [CmdletBinding()]
    param([switch]$IncludeOneDrive)

    Write-Log 'Retrieving sites via Graph API...'
    $sites = [System.Collections.Generic.List[object]]::new()

    # Try getAllSites first, fall back to search
    $uri = "v1.0/sites/getAllSites?`$select=id,webUrl,displayName,isPersonalSite&`$top=999"
    $useFallback = $false

    do {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        }
        catch {
            if (-not $useFallback) {
                Write-Log 'getAllSites not available, falling back to search...' -Level Warning
                $uri = "v1.0/sites?search=*&`$top=999"
                $useFallback = $true
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            }
            else { throw }
        }

        foreach ($s in $response.value) {
            $isPersonal = $false
            if ($s.isPersonalSite -eq $true) { $isPersonal = $true }
            if ($s.webUrl -like '*-my.sharepoint.*') { $isPersonal = $true }

            if ($isPersonal -and -not $IncludeOneDrive) { continue }

            $sites.Add([PSCustomObject]@{
                Url    = $s.webUrl
                Title  = $s.displayName
                SiteId = $s.id
            })
        }

        $uri = $response.'@odata.nextLink'
    } while ($uri)

    Write-Log "Found $($sites.Count) sites via Graph." -Level Success
    return $sites
}

function Get-ExternalUsersSPO {
    <# Gets external users for a site via SPO module. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SiteUrl)

    $users = [System.Collections.Generic.List[object]]::new()
    $position = 0
    $hasMore = $true

    while ($hasMore) {
        try {
            $batch = Get-SPOExternalUser -SiteUrl $SiteUrl -Position $position -PageSize $spoExternalUserPageSize -ErrorAction Stop

            if ($batch -and $batch.ExternalUsers -and $batch.ExternalUsers.Count -gt 0) {
                $users.AddRange($batch.ExternalUsers)
                $position += $batch.ExternalUsers.Count
                if ($batch.ExternalUsers.Count -lt $spoExternalUserPageSize) { $hasMore = $false }
            }
            else { $hasMore = $false }
        }
        catch {
            if ($_.Exception.Message -like '*Access denied*' -or $_.Exception.Message -like '*not found*') {
                Write-Log "Skipping site (access): $SiteUrl" -Level Warning
            }
            else {
                Write-Log "Error at $SiteUrl`: $($_.Exception.Message)" -Level Warning
            }
            $hasMore = $false
        }
    }

    return $users
}

function Get-ExternalUsersGraph {
    <# Gets external/guest users for a site via Graph site permissions API. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [string]$SiteId
    )

    $users = [System.Collections.Generic.List[object]]::new()

    try {
        # Resolve site ID if not provided
        if (-not $SiteId) {
            $parsedUri = [System.Uri]$SiteUrl
            $hostName = $parsedUri.Host
            $sitePath = $parsedUri.AbsolutePath.TrimEnd('/')
            if ([string]::IsNullOrEmpty($sitePath) -or $sitePath -eq '/') {
                $lookupUri = "v1.0/sites/${hostName}:/"
            }
            else {
                $lookupUri = "v1.0/sites/${hostName}:${sitePath}"
            }

            try {
                $siteInfo = Invoke-MgGraphRequest -Method GET -Uri $lookupUri -ErrorAction Stop
                $SiteId = $siteInfo.id
            }
            catch {
                return $users
            }
        }

        # Get site permissions
        $permUri = "v1.0/sites/$SiteId/permissions?`$top=999"
        do {
            try {
                $permResponse = Invoke-MgGraphRequest -Method GET -Uri $permUri -ErrorAction Stop
            }
            catch {
                break
            }

            foreach ($perm in $permResponse.value) {
                if ($perm.grantedToIdentitiesV2) {
                    foreach ($identity in $perm.grantedToIdentitiesV2) {
                        $siteUser = $identity.siteUser
                        if (-not $siteUser) { continue }

                        $loginName = $siteUser.loginName
                        if (-not $loginName) { continue }
                        if ($loginName -notlike '*#ext#*' -and $loginName -notlike '*urn:spo:guest*') { continue }

                        $email = ''
                        if ($siteUser.email) {
                            $email = $siteUser.email
                        }
                        elseif ($loginName -match 'i:0#\.f\|membership\|(.+)') {
                            $email = $Matches[1]
                        }

                        $displayName = Coalesce $siteUser.displayName $email
                        $users.Add([PSCustomObject]@{
                            DisplayName = $displayName
                            Email       = $email
                            InvitedAs   = $email
                            AcceptedAs  = $email
                            WhenCreated = $null
                        })
                    }
                }
            }

            $permUri = $permResponse.'@odata.nextLink'
        } while ($permUri)
    }
    catch {
        Write-Log "Graph permissions error for $SiteUrl`: $($_.Exception.Message)" -Level Warning
    }

    return $users
}
#endregion

#region Main Execution
$script:ForceGraphOnly = $false

try {
    Write-Log '=========================================='
    Write-Log 'SPO OTP Retirement Impact Assessment'
    Write-Log 'Reference: MC1243549'
    Write-Log '=========================================='
    Write-Log "Tenant input:          $TenantName"
    Write-Log "Resolved admin URL:    $spoAdminUrl"
    Write-Log "Resolved tenant URL:   $spoTenantUrl"
    Write-Log "Inactivity threshold:  $InactiveDaysThreshold days (before $($activityCutoffDate.ToString('yyyy-MM-dd')))"
    Write-Log "Include inactive:      $IncludeInactiveUsers"
    Write-Log "Include OneDrive:      $IncludeOneDriveSites"
    Write-Log "Graph-only mode:       $GraphOnly"
    Write-Log "Output:                $OutputPath"
    Write-Log '=========================================='

    # Step 1: Load modules
    Write-Log '--- Step 1: Loading modules ---'
    Initialize-Modules -SkipSPO:$GraphOnly

    # Step 2: Connect to services
    Write-Log '--- Step 2: Connecting to services ---'
    Connect-Services -SkipSPO:$GraphOnly

    $effectiveGraphOnly = $GraphOnly.IsPresent -or $script:ForceGraphOnly
    if ($effectiveGraphOnly) {
        Write-Log 'Operating in GRAPH-ONLY mode.' -Level Warning
        Write-Log 'Per-site external user data may be less detailed than SPO mode.' -Level Warning
    }

    # Step 3: Get all Entra B2B guest accounts (safe users)
    Write-Log '--- Step 3: Entra B2B guest accounts ---'
    $entraGuests = Get-EntraGuestAccounts

    # Step 4: Get all sites
    Write-Log '--- Step 4: Enumerating sites ---'
    if ($effectiveGraphOnly) {
        $sites = Get-SitesGraph -IncludeOneDrive:$IncludeOneDriveSites
    }
    else {
        $sites = Get-SitesSPO -IncludeOneDrive:$IncludeOneDriveSites
    }

    # Step 5: Enumerate external users per site
    Write-Log '--- Step 5: Scanning sites for external users ---'
    $affectedUsers = @{}
    $totalExternal = 0
    $alreadyB2B = 0
    $siteIndex = 0

    foreach ($site in $sites) {
        $siteIndex++
        $pct = [math]::Round(($siteIndex / $sites.Count) * 100, 1)
        Write-Progress -Activity 'Scanning sites for external users' -Status "$siteIndex / $($sites.Count) - $($site.Url)" -PercentComplete $pct

        if ($effectiveGraphOnly) {
            $siteExternalUsers = Get-ExternalUsersGraph -SiteUrl $site.Url -SiteId $site.SiteId
        }
        else {
            $siteExternalUsers = Get-ExternalUsersSPO -SiteUrl $site.Url
        }

        foreach ($ext in $siteExternalUsers) {
            $totalExternal++
            $email = Coalesce $ext.Email $ext.AcceptedAs $ext.InvitedAs
            if (-not $email) {
                Write-Log "External user with no email on $($site.Url) - $($ext.DisplayName)" -Level Warning
                continue
            }
            $email = $email.ToLower().Trim()

            # Check if user has a B2B guest account
            $hasB2B = $entraGuests.ContainsKey($email)
            if (-not $hasB2B -and $ext.AcceptedAs) {
                $hasB2B = $entraGuests.ContainsKey($ext.AcceptedAs.ToLower().Trim())
            }
            if (-not $hasB2B -and $ext.InvitedAs) {
                $hasB2B = $entraGuests.ContainsKey($ext.InvitedAs.ToLower().Trim())
            }

            if ($hasB2B) {
                $alreadyB2B++
                continue
            }

            # Affected user - accumulate sites
            if ($affectedUsers.ContainsKey($email)) {
                $existingSites = $affectedUsers[$email].Sites
                if ($existingSites -notcontains $site.Url) {
                    $affectedUsers[$email].Sites += $site.Url
                    $affectedUsers[$email].SiteTitles += $site.Title
                }
            }
            else {
                # Determine activity status
                $isActive = $true
                $lastActivity = $null
                $activitySource = 'No data (assumed active)'

                if ($ext.WhenCreated) {
                    try {
                        $lastActivity = [datetime]$ext.WhenCreated
                        $activitySource = 'SPO WhenCreated'
                        $isActive = $lastActivity -ge $activityCutoffDate
                    }
                    catch { }
                }

                $affectedUsers[$email] = @{
                    Email            = $email
                    DisplayName      = $ext.DisplayName
                    InvitedAs        = $ext.InvitedAs
                    AcceptedAs       = $ext.AcceptedAs
                    WhenCreated      = $ext.WhenCreated
                    IsActive         = $isActive
                    LastActivityDate = $lastActivity
                    ActivitySource   = $activitySource
                    Sites            = [System.Collections.Generic.List[string]]@($site.Url)
                    SiteTitles       = [System.Collections.Generic.List[string]]@($site.Title)
                }
            }
        }
    }
    Write-Progress -Activity 'Scanning sites for external users' -Completed

    # Step 6: Enrich with Unified Audit Log (best-effort)
    Write-Log '--- Step 6: Enriching activity data ---'
    $enriched = 0

    if (Get-Command -Name 'Search-UnifiedAuditLog' -ErrorAction SilentlyContinue) {
        foreach ($email in @($affectedUsers.Keys)) {
            $u = $affectedUsers[$email]
            try {
                $auditParams = @{
                    StartDate  = $activityCutoffDate
                    EndDate    = Get-Date
                    UserIds    = $email
                    Operations = @('FileAccessed', 'FileDownloaded', 'FileModified', 'PageViewed')
                    ResultSize = 1
                }
                $audit = Search-UnifiedAuditLog @auditParams -ErrorAction SilentlyContinue

                if ($audit -and $audit.Count -gt 0) {
                    $lastDate = ($audit | Sort-Object CreationDate -Descending | Select-Object -First 1).CreationDate
                    $u.LastActivityDate = $lastDate
                    $u.ActivitySource = 'Unified Audit Log'
                    $u.IsActive = $lastDate -ge $activityCutoffDate
                    $enriched++
                }
            }
            catch { }
        }
        if ($enriched -gt 0) { Write-Log "Enriched $enriched users from audit logs." -Level Success }
    }
    else {
        Write-Log 'Search-UnifiedAuditLog not available. Skipping enrichment.' -Level Warning
    }

    # Step 7: Build report
    Write-Log '--- Step 7: Building report ---'
    $reportRows = @()
    foreach ($email in $affectedUsers.Keys) {
        $u = $affectedUsers[$email]

        if (-not $IncludeInactiveUsers -and -not $u.IsActive) { continue }

        $lastDateStr = 'N/A'
        if ($u.LastActivityDate) {
            $lastDateStr = ([datetime]$u.LastActivityDate).ToString('yyyy-MM-dd')
        }

        $impactText = 'LOW - Inactive account'
        if ($u.IsActive) { $impactText = 'HIGH - Active user will lose access' }

        $modeText = 'SPO Module'
        if ($effectiveGraphOnly) { $modeText = 'Graph API' }

        $reportRows += [PSCustomObject]@{
            Email            = $u.Email
            DisplayName      = $u.DisplayName
            InvitedAs        = $u.InvitedAs
            AcceptedAs       = $u.AcceptedAs
            WhenCreated      = $u.WhenCreated
            IsActive         = $u.IsActive
            LastActivityDate = $lastDateStr
            ActivitySource   = $u.ActivitySource
            HasB2BAccount    = $false
            SiteCount        = $u.Sites.Count
            SiteUrls         = ($u.Sites | Select-Object -Unique) -join '; '
            SiteTitles       = ($u.SiteTitles | Select-Object -Unique) -join '; '
            Impact           = $impactText
            DataSource       = $modeText
        }
    }

    $reportRows = $reportRows | Sort-Object @{Expression = 'IsActive'; Descending = $true}, @{Expression = 'SiteCount'; Descending = $true}

    # Step 8: Export
    if ($reportRows.Count -gt 0) {
        $reportRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Log "Report exported to: $OutputPath" -Level Success
    }
    else {
        Write-Log 'No affected users found. No CSV generated.' -Level Success
    }

    # Step 9: Summary
    $activeAffected = @($reportRows | Where-Object { $_.IsActive -eq $true })
    $inactiveAffected = @($reportRows | Where-Object { $_.IsActive -eq $false })
    $uniqueSites = @($reportRows | ForEach-Object { $_.SiteUrls -split '; ' } | Where-Object { $_ } | Select-Object -Unique)

    Write-Log ''
    Write-Log '=========================================='
    Write-Log '         ASSESSMENT SUMMARY'
    Write-Log '=========================================='
    $modeLabel = 'SPO + Graph'
    if ($effectiveGraphOnly) { $modeLabel = 'Graph-only' }
    Write-Log "Mode:                                   $modeLabel"
    Write-Log "Total external users found:             $totalExternal"
    Write-Log "Already have B2B guest account (safe):  $alreadyB2B"
    Write-Log "Affected by SPO OTP retirement:         $($affectedUsers.Count)"
    Write-Log "  - Active (HIGH impact):               $($activeAffected.Count)"
    Write-Log "  - Inactive (LOW impact):              $($inactiveAffected.Count)"
    Write-Log "Sites with affected users:              $($uniqueSites.Count)"
    Write-Log '=========================================='

    if ($activeAffected.Count -gt 0) {
        Write-Log ''
        Write-Log 'RECOMMENDATION: Create B2B guest accounts for active affected users before July 2026.' -Level Warning
        Write-Log 'Options:' -Level Warning
        Write-Log '  1. Bulk invite via Entra admin center or New-MgInvitation' -Level Warning
        Write-Log '  2. Have internal users re-share content (auto-creates guest account)' -Level Warning
        Write-Log '  3. Enable Email OTP in Entra External ID settings as fallback' -Level Warning
    }
    else {
        Write-Log 'No active external users affected. No immediate action required.' -Level Success
    }

    Write-Output $reportRows
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    throw
}
finally {
    Write-Progress -Activity 'Scanning sites for external users' -Completed
}
#endregion
