#Requires -Version 5.1

<#
.SYNOPSIS
    Diagnostic for SPO OTP retirement (MC1243549).
    Checks tenant configuration and identifies affected external users.

.PARAMETER TenantName
    SharePoint tenant name (e.g., 'contoso'). Required for SPO checks.

.PARAMETER SampleSiteCount
    Number of sites to sample for external users. Default: 20.

.NOTES
    Author:  Petr Beloch
    Modules: Microsoft.Graph.Authentication (required)
             Microsoft.Online.SharePoint.PowerShell (required for full check)
    Scopes:  Policy.Read.All, User.Read.All, IdentityProvider.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantName,

    [int]$SampleSiteCount = 20
)

$ErrorActionPreference = 'Stop'

# Normalize tenant name
$clean = $TenantName.Trim().TrimEnd('/') -replace '^https?://', ''
if ($clean -match '^([a-zA-Z0-9\-]+?)(?:-admin)?\.sharepoint\.') { $clean = $Matches[1] }
$spoAdminUrl = "https://$clean-admin.sharepoint.com"
$spoTenantUrl = "https://$clean.sharepoint.com"

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host '  SPO OTP RETIREMENT (MC1243549) - DIAGNOSTIC' -ForegroundColor Cyan
Write-Host "  Tenant: $clean" -ForegroundColor Cyan
Write-Host "================================================================`n" -ForegroundColor Cyan

# =====================================================================
# CHECK 1: Graph connection + Entra B2B Email OTP policy
# =====================================================================
Write-Host '[1/5] Entra B2B Email OTP policy' -ForegroundColor Yellow
Write-Host '      (Determines what happens AFTER SPO OTP retirement)' -ForegroundColor Gray

$requiredScopes = @('Policy.Read.All', 'User.Read.All', 'IdentityProvider.Read.All')
try {
    $ctx = Get-MgContext
    if ($null -eq $ctx) { throw 'No session' }
    $missing = $requiredScopes | Where-Object { $_ -notin $ctx.Scopes }
    if ($missing) {
        Write-Host "      Missing scopes: $($missing -join ', '). Reconnecting..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $null = Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        $ctx = Get-MgContext
    }
    Write-Host "      Connected as: $($ctx.Account)" -ForegroundColor Green
}
catch {
    $null = Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    $ctx = Get-MgContext
    Write-Host "      Connected as: $($ctx.Account)" -ForegroundColor Green
}

try {
    $emailOtp = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/Email' -ErrorAction Stop

    $state = $emailOtp.state
    $extIdOtp = $emailOtp.allowExternalIdToUseEmailOtp

    Write-Host ''
    Write-Host "      State:                         $state"
    Write-Host "      allowExternalIdToUseEmailOtp:   $extIdOtp"
    Write-Host ''

    if ($state -eq 'enabled' -and $extIdOtp -ne 'disabled') {
        Write-Host '      RESULT: Email OTP is ENABLED.' -ForegroundColor Green
        Write-Host '      After retirement, guests without MSA/federation will use Entra Email OTP.' -ForegroundColor Green
    }
    elseif ($state -eq 'disabled' -or $extIdOtp -eq 'disabled') {
        Write-Host '      WARNING: Email OTP is DISABLED!' -ForegroundColor Red
        Write-Host '      After retirement, guests will be forced to create Microsoft Account.' -ForegroundColor Red
        Write-Host '      Enable: Entra admin > External Identities > All identity providers > Email OTP > Yes' -ForegroundColor Yellow
    }
}
catch {
    Write-Host "      Error: $($_.Exception.Message)" -ForegroundColor Red
}

# =====================================================================
# CHECK 2: External Identity Providers
# =====================================================================
Write-Host "`n[2/5] External Identity Providers" -ForegroundColor Yellow
Write-Host '      (Federated IdPs that guests can use instead of OTP)' -ForegroundColor Gray

try {
    $idProviders = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/identity/identityProviders' -ErrorAction Stop

    if ($idProviders.value -and $idProviders.value.Count -gt 0) {
        foreach ($idp in $idProviders.value) {
            $name = if ($idp.displayName) { $idp.displayName } else { $idp.identityProviderType }
            Write-Host "      - $name ($($idp.identityProviderType))" -ForegroundColor White
        }
    }
    else {
        Write-Host '      None configured (only MSA and Email OTP available for guests).' -ForegroundColor Yellow
    }
}
catch {
    Write-Host "      Error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host '      Check manually: Entra admin > External Identities > All identity providers' -ForegroundColor Gray
}

# =====================================================================
# CHECK 3: SPO B2B Integration setting
# =====================================================================
Write-Host "`n[3/5] SharePoint B2B Integration setting" -ForegroundColor Yellow
Write-Host '      (If already enabled, external sharing already uses Entra B2B)' -ForegroundColor Gray

$spoConnected = $false
try {
    $null = Get-SPOSite -Identity $spoTenantUrl -ErrorAction Stop
    $spoConnected = $true
    Write-Host "      SPO connected to: $spoTenantUrl" -ForegroundColor Green
}
catch {
    Write-Host '      SPO not connected. Attempting connection...' -ForegroundColor Yellow
    try {
        Connect-SPOService -Url $spoAdminUrl -ErrorAction Stop
        $spoConnected = $true
        Write-Host "      SPO connected." -ForegroundColor Green
    }
    catch {
        Write-Host "      Cannot connect: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Connect manually: Connect-SPOService -Url '$spoAdminUrl'" -ForegroundColor Yellow
    }
}

$b2bAlreadyEnabled = $false
if ($spoConnected) {
    try {
        $spoTenant = Get-SPOTenant -ErrorAction Stop

        $b2bIntegration = $spoTenant.EnableAzureADB2BIntegration
        $sharingCapability = $spoTenant.SharingCapability

        Write-Host ''
        Write-Host "      EnableAzureADB2BIntegration:  $b2bIntegration"
        Write-Host "      SharingCapability:            $sharingCapability"
        Write-Host ''

        if ($b2bIntegration -eq $true) {
            $b2bAlreadyEnabled = $true
            Write-Host '      RESULT: B2B integration is ALREADY ENABLED.' -ForegroundColor Green
            Write-Host '      All new external sharing already creates Entra B2B guest accounts.' -ForegroundColor Green
            Write-Host '      Impact: Only users who shared BEFORE this was enabled and never' -ForegroundColor Green
            Write-Host '      re-authenticated through B2B could be affected.' -ForegroundColor Green
        }
        else {
            Write-Host '      RESULT: B2B integration is NOT enabled.' -ForegroundColor Yellow
            Write-Host '      SPO currently uses its own OTP mechanism for some external users.' -ForegroundColor Yellow
            Write-Host '      After May 2026 this will be forced to B2B automatically.' -ForegroundColor Yellow
        }

        # SharingCapability analysis
        Write-Host ''
        switch ($sharingCapability) {
            'Disabled' {
                Write-Host '      SharingCapability: External sharing is DISABLED.' -ForegroundColor Green
                Write-Host '      No external users can access content. Minimal SPO OTP risk.' -ForegroundColor Green
            }
            'ExistingExternalUserSharingOnly' {
                Write-Host '      SharingCapability: Only EXISTING external users allowed.' -ForegroundColor Yellow
                Write-Host '      No new external sharing invitations. Existing SPO OTP users' -ForegroundColor Yellow
                Write-Host '      may still be affected if they lack B2B guest accounts.' -ForegroundColor Yellow
            }
            'ExternalUserSharingOnly' {
                Write-Host '      SharingCapability: Sharing with authenticated external users.' -ForegroundColor Yellow
            }
            'ExternalUserAndGuestSharing' {
                Write-Host '      SharingCapability: Anyone (including anonymous links).' -ForegroundColor Red
                Write-Host '      Broadest sharing - highest potential SPO OTP impact.' -ForegroundColor Red
            }
            default {
                Write-Host "      SharingCapability: $sharingCapability" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "      Error reading SPO tenant settings: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# =====================================================================
# CHECK 4: Count SPO external users and cross-reference with Entra B2B
# =====================================================================
Write-Host "`n[4/5] SPO external users (tenant-wide)" -ForegroundColor Yellow
Write-Host '      (Users in SharePoint external user store)' -ForegroundColor Gray

if (-not $spoConnected) {
    Write-Host '      Skipped: SPO not connected.' -ForegroundColor Red
}
else {
    $allSpoExternal = @()

    # Method 1: Bare cmdlet (no parameters) - works in some SPO module versions where -Position/-PageSize breaks
    Write-Host '      Method 1: Get-SPOExternalUser (bare, no params)...' -ForegroundColor Gray
    try {
        $bareResult = @(Get-SPOExternalUser -ErrorAction Stop)
        if ($bareResult.Count -gt 0) {
            $allSpoExternal = $bareResult
            Write-Host "      Bare result: $($bareResult.Count) external users" -ForegroundColor Green
        }
        else {
            Write-Host '      Bare result: 0 external users'
        }
    }
    catch {
        Write-Host "      Bare call failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Method 2: Paginated (handles tenants with many external users)
    if ($allSpoExternal.Count -eq 0) {
        Write-Host '      Method 2: Get-SPOExternalUser -Position -PageSize (paginated)...' -ForegroundColor Gray
        $position = 0
        $hasMore = $true

        while ($hasMore) {
            try {
                $batch = Get-SPOExternalUser -Position $position -PageSize 50 -ErrorAction Stop
                if ($batch -and $batch.ExternalUsers -and $batch.ExternalUsers.Count -gt 0) {
                    $allSpoExternal += $batch.ExternalUsers
                    $position += $batch.ExternalUsers.Count
                    if ($batch.ExternalUsers.Count -lt 50) { $hasMore = $false }
                }
                else { $hasMore = $false }
            }
            catch {
                Write-Host "      Error: $($_.Exception.Message)" -ForegroundColor Yellow
                $hasMore = $false
            }
        }
        Write-Host "      Paginated result: $($allSpoExternal.Count) external users"
    }

    # Method 3: If still 0, sample individual sites
    if ($allSpoExternal.Count -eq 0) {
        Write-Host ''
        Write-Host "      Method 3: Sampling $SampleSiteCount sites with -SiteUrl..." -ForegroundColor Gray

        $sampleSites = Get-SPOSite -Limit $SampleSiteCount -IncludePersonalSite:$false |
            Where-Object {
                $_.Status -eq 'Active' -and
                $_.Template -notlike 'SRCHCEN*' -and
                $_.Template -notlike 'SPSMSITEHOST*' -and
                $_.Template -notlike 'APPCATALOG*' -and
                $_.Template -notlike 'POINTPUBLISHINGHUB*' -and
                $_.SharingCapability -ne 'Disabled'
            } | Select-Object -First $SampleSiteCount

        $sampledExternal = @()
        foreach ($site in $sampleSites) {
            try {
                $batch = Get-SPOExternalUser -SiteUrl $site.Url -Position 0 -PageSize 50 -ErrorAction Stop
                if ($batch -and $batch.ExternalUsers -and $batch.ExternalUsers.Count -gt 0) {
                    Write-Host "        $($site.Url) -> $($batch.ExternalUsers.Count) external users" -ForegroundColor White
                    $sampledExternal += $batch.ExternalUsers
                }
            }
            catch { }
        }

        Write-Host ''
        Write-Host "      Sampled result: $($sampledExternal.Count) external users across $($sampleSites.Count) sites"
        $allSpoExternal = $sampledExternal
    }

    # Cross-reference with Entra B2B
    if ($allSpoExternal.Count -gt 0) {
        Write-Host ''
        Write-Host '      Cross-referencing with Entra B2B guest accounts...' -ForegroundColor Gray

        # Build Entra guest lookup
        $guestLookup = @{}
        $uri = "v1.0/users?`$filter=userType eq 'Guest'&`$select=mail,otherMails,userPrincipalName&`$top=999&`$count=true"
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' } -ErrorAction Stop
            foreach ($g in $resp.value) {
                if ($g.mail) { $guestLookup[$g.mail.ToLower()] = $true }
                if ($g.otherMails) {
                    foreach ($m in $g.otherMails) { $guestLookup[$m.ToLower()] = $true }
                }
                if ($g.userPrincipalName -match '^(.+)#EXT#@') {
                    $extracted = ($Matches[1] -replace '_', '@')
                    if (($extracted -split '@').Count -eq 2) { $guestLookup[$extracted.ToLower()] = $true }
                }
            }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)

        Write-Host "      Entra guest lookup: $($guestLookup.Count) unique emails indexed"

        $withB2B = 0
        $withoutB2B = 0
        $noEmail = 0
        $sampleWithout = @()

        foreach ($ext in $allSpoExternal) {
            $email = if ($ext.Email) { $ext.Email }
                     elseif ($ext.AcceptedAs) { $ext.AcceptedAs }
                     elseif ($ext.InvitedAs) { $ext.InvitedAs }
                     else { $null }

            if (-not $email) { $noEmail++; continue }
            $emailLower = $email.ToLower().Trim()

            $found = $guestLookup.ContainsKey($emailLower)
            if (-not $found -and $ext.AcceptedAs) {
                $found = $guestLookup.ContainsKey($ext.AcceptedAs.ToLower().Trim())
            }
            if (-not $found -and $ext.InvitedAs) {
                $found = $guestLookup.ContainsKey($ext.InvitedAs.ToLower().Trim())
            }

            if ($found) {
                $withB2B++
            }
            else {
                $withoutB2B++
                if ($sampleWithout.Count -lt 15) {
                    $sampleWithout += [PSCustomObject]@{
                        Email       = $email
                        DisplayName = $ext.DisplayName
                        AcceptedAs  = $ext.AcceptedAs
                        WhenCreated = $ext.WhenCreated
                    }
                }
            }
        }

        Write-Host ''
        Write-Host '      ===========================================' -ForegroundColor Cyan
        Write-Host '      SPO EXTERNAL USER ANALYSIS' -ForegroundColor Cyan
        Write-Host '      ===========================================' -ForegroundColor Cyan
        Write-Host "      Total SPO external users found:        $($allSpoExternal.Count)" -ForegroundColor White
        Write-Host "      Have Entra B2B account (SAFE):         $withB2B" -ForegroundColor Green
        Write-Host "      NO B2B account (AT RISK - SPO OTP):    $withoutB2B" -ForegroundColor $(if ($withoutB2B -gt 0) { 'Red' } else { 'Green' })
        Write-Host "      No email (skipped):                    $noEmail" -ForegroundColor $(if ($noEmail -gt 0) { 'Yellow' } else { 'Gray' })
        Write-Host '      ===========================================' -ForegroundColor Cyan

        if ($sampleWithout.Count -gt 0) {
            Write-Host ''
            Write-Host '      Sample at-risk users:' -ForegroundColor Yellow
            $sampleWithout | Format-Table Email, DisplayName, AcceptedAs, WhenCreated -AutoSize | Out-String | ForEach-Object { Write-Host "      $_" -ForegroundColor Gray }
        }
    }
    else {
        Write-Host ''
        Write-Host '      No SPO external users found.' -ForegroundColor Yellow
        if ($b2bAlreadyEnabled) {
            Write-Host '      This is expected: B2B integration is already enabled.' -ForegroundColor Green
            Write-Host '      All external sharing goes through Entra B2B.' -ForegroundColor Green
        }
        else {
            Write-Host '      This may indicate no external sharing has occurred,' -ForegroundColor Yellow
            Write-Host '      or the cmdlet requires different permissions.' -ForegroundColor Yellow
        }
    }
}

# =====================================================================
# CHECK 5: Summary & recommendations
# =====================================================================
Write-Host "`n[5/5] SUMMARY & RECOMMENDATIONS" -ForegroundColor Yellow
Write-Host '      ===========================================' -ForegroundColor Cyan

if ($b2bAlreadyEnabled) {
    Write-Host '      EnableAzureADB2BIntegration is already TRUE.' -ForegroundColor Green
    Write-Host '      Your tenant has been using Entra B2B for external sharing.' -ForegroundColor Green
    Write-Host ''
    Write-Host '      Residual risk:' -ForegroundColor Yellow
    Write-Host '      - Users who shared files BEFORE B2B integration was enabled' -ForegroundColor Yellow
    Write-Host '        and authenticated via old SPO OTP may still exist.' -ForegroundColor Yellow
    Write-Host '      - After July 2026, these users get Access Denied on old links.' -ForegroundColor Yellow
    Write-Host '      - An internal user can re-share to auto-create B2B guest account.' -ForegroundColor Yellow
}
else {
    Write-Host '      EnableAzureADB2BIntegration is FALSE.' -ForegroundColor Red
    Write-Host '      Starting May 2026, Microsoft will force B2B for new sharing.' -ForegroundColor Yellow
    Write-Host '      Starting July 2026, existing SPO OTP users lose access.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '      Recommended actions:' -ForegroundColor Yellow
    Write-Host '        1. Run full assessment with SPO mode to find affected users' -ForegroundColor White
    Write-Host '        2. Enable B2B integration now:' -ForegroundColor White
    Write-Host "           Set-SPOTenant -EnableAzureADB2BIntegration `$true" -ForegroundColor Gray
    Write-Host '        3. Ensure Email OTP is enabled in Entra (already done)' -ForegroundColor White
    Write-Host '        4. Communicate to users about potential access interruption' -ForegroundColor White
}

Write-Host ''
Write-Host '      ===========================================' -ForegroundColor Cyan
Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host '  DIAGNOSTIC COMPLETE' -ForegroundColor Cyan
Write-Host "================================================================`n" -ForegroundColor Cyan
