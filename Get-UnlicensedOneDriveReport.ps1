<#
.SYNOPSIS
    Identifies all unlicensed OneDrive accounts across the tenant (all geo locations)
    using Microsoft Graph API only — no SPO PowerShell module, no per-geo tokens,
    and no manual admin center navigation required.

.DESCRIPTION
    Microsoft 365 multi-geo is handled transparently by the Graph API: a single
    access token scoped to the home tenant automatically routes /users/{id}/drive
    requests to the correct regional datacenter (NAM, APC, CAN, DEU, GBR, IND, JPN, and etc).

    The script identifies two populations of unlicensed OneDrive accounts:

    POPULATION 1 — Active Entra ID users without an enabled OneDrive/SharePoint plan
      These users still have an active Entra account but their license no longer
      includes an enabled OneDrive or SharePoint Online service plan.
      Graph endpoint : GET /users?$select=id,userPrincipalName,...,assignedPlans
      Unlicensed date: GET /auditLogs/directoryAudits (bulk, optional)

    POPULATION 2 — Soft-deleted users (within Entra ID 30-day recycle bin)
      These users were deleted from Entra ID. Their OneDrives still exist.
      Graph endpoint : GET /directory/deletedItems/microsoft.graph.user
      Unlicensed date: deletedDateTime from the deleted user object.

    POPULATION 3 — Currently archived OneDrive sites (Sites API)
      Personal OneDrive sites that Microsoft has already archived — typically because
      the owner was deleted from Entra ID more than 30 days ago (purged from the
      recycle bin). The Entra user object no longer exists so Phases 1 and 2 cannot
      find these accounts. They are discovered by enumerating all SharePoint sites.
      Graph endpoint : GET /beta/sites/getAllSites?$filter=isPersonalSite eq true&$select=...siteCollection
      Strategy       : Pass 1 — bulk beta call; archivalDetails inline = no per-site call needed.
                       Pass 2 — per-site fallback (beta GET + HTTP 423) for sites with null archivalDetails.
      Unlicensed date: Not available — occurred before the Entra purge (>30 days ago)
      Requires       : Sites.Read.All (Application)
      Toggle via     : $GetCurrentlyArchived = $true / $false

    LIMITATION: Users deleted >30 days ago whose OneDrive has already been purged
    (not just archived) are permanently gone and cannot be discovered via Graph.

    Timeline per Microsoft docs (enforcement began Jan 27, 2025):
      Day 60 → read-only mode
      Day 93 → archived (or deletion begins if billing not enabled)
    https://learn.microsoft.com/en-us/sharepoint/unlicensed-onedrive-accounts

.PARAMETER None
    All configuration is set in the CONFIGURATION SECTION below.

.NOTES
    File Name   : Get-UnlicensedOneDriveReport.ps1
    Author      : Mike Lee
    Date Created: 4/28/26

    Required Microsoft Graph App Permissions (Application type):
      User.Read.All           — Enumerate users and inspect assignedPlans/licenses
      Directory.Read.All      — Read soft-deleted users from Entra recycle bin
      Files.Read.All          — Read OneDrive drive metadata for any user
      AuditLog.Read.All       — [OPTIONAL] directoryAudits for license-change dates
                                 Set $includeLicenseRemovalDates = $false to skip.
      Sites.Read.All          — [OPTIONAL] GET /sites/getAllSites for currently archived OneDrive sites
                                 Set $GetCurrentlyArchived = $false to skip.

    A SINGLE app registration in the HOME TENANT covers all geo locations.
    No per-geo tokens required — Graph handles multi-geo routing automatically.

.OUTPUTS
    CSV: UnlicensedOneDrive_<timestamp>.csv

.EXAMPLE
    PS> .\Get-UnlicensedOneDriveReport.ps1

.LINK
    https://learn.microsoft.com/en-us/sharepoint/unlicensed-onedrive-accounts
    https://learn.microsoft.com/en-us/graph/api/user-list
    https://learn.microsoft.com/en-us/graph/api/drive-get
    https://learn.microsoft.com/en-us/graph/api/directoryaudit-list
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

# ---- Debug output ----
$debug = $false

# ---- Tenant & App Registration ----
# A SINGLE registration in the home (NAM) tenant covers all geo locations.
# Graph routes /users/{id}/drive transparently to APC, CAN, DEU, GBR, IND, JPN.
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57'

# ---- Authentication type: 'Certificate' or 'ClientSecret' ----
$AuthType = 'Certificate'

# Certificate thumbprint (used when $AuthType = 'Certificate')
$Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9'

# Certificate store: 'LocalMachine' or 'CurrentUser'
$CertStore = 'LocalMachine'

# Client Secret (used when $AuthType = 'ClientSecret')
$clientSecret = ''

# ---- Report output ----
$OutputFolder = $env:TEMP

# ---- Archival timeline thresholds (days) — per Microsoft documentation ----
$ReadOnlyThresholdDays = 60
$ArchiveThresholdDays = 93

# ---- Audit log: license removal date discovery for ACTIVE unlicensed users ----
# When $true: queries directoryAudits (bulk) to find when each active user's
# OneDrive/SharePoint license was removed. Requires AuditLog.Read.All.
# When $false: active users will show UnlicensedDate as 'Unknown'.
$includeLicenseRemovalDates = $true

# How far back to search for license-change audit events (max 180 days).
$AuditLogLookbackDays = 180

# ---- Currently archived OneDrive sites (Sites API) ----
# When $true: queries GET /sites/getAllSites to find personal OneDrive sites that
# Microsoft has already archived. These are accounts whose Entra user was deleted
# >30 days ago (purged from recycle bin) and whose OneDrive has since been archived.
# Requires Sites.Read.All (Application) on the app registration.
# When $false: only reports on active/soft-deleted populations (no Sites.Read.All needed).
$GetCurrentlyArchived = $true

# ---- Request throttling ----
$MaxRetries = 15
$InitialBackoffSec = 3
$RequestTimeoutSec = 300

# ---- Delay between individual drive queries (seconds). 0 = no delay. ----
$delayBetweenRequests = 0

##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
$date = Get-Date -Format 'yyyyMMddHHmmss'
$today = (Get-Date).Date
$outputLog = Join-Path $OutputFolder "UnlicensedOneDrive_$date.csv"

$global:token = $null
$global:tokenExpiry = $null
#endregion Initialization

#region Constants — OneDrive & SharePoint Online Service Plan IDs
# A user is licensed for OneDrive when at least one of these plan IDs appears
# in their assignedPlans with capabilityStatus = 'Enabled'.
# Source: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
# Additional IDs from: https://github.com/michevnew/PowerShell/blob/master/Report_Unlicensed_OneDrives.ps1
$OneDrivePlanIds = @(
    'b4ac11a0-32ff-4e78-982d-e039fa803dec'  # ONEDRIVELITE_IW         — Office for the web with OneDrive (Basic Collaboration)
    'f7e5b77d-f293-410a-bae8-f941f19fe680'  # ONEDRIVECLIPCHAMP        — OneDrive included with Clipchamp Premium
    '13696edf-5a08-49f6-8134-03083ed8ba30'  # ONEDRIVESTANDARD         — OneDrive for Business Plan 1 (M365 Apps, E1)
    '4495894f-534f-41ca-9d3b-0ebf1220a423'  # ONEDRIVE_BASIC variant   — (unlisted in MS docs; retained from community reference)
    'afcafa6a-d966-4462-918c-ec0b4e0fe642'  # ONEDRIVEENTERPRISE        — OneDrive for Business Plan 2 (standalone)
    'da792a53-cbc0-4184-a10d-e544dd34b3c1'  # ONEDRIVE_BASIC            — OneDrive for Business Basic (Visio plans)
    '98709c2e-96b5-4244-95f5-a0ebe139fb8a'  # ONEDRIVE_BASIC_GOV        — OneDrive for Business Basic for Government
)
$SharePointPlanIds = @(
    'e95bec33-7c88-4a70-8e19-b10bd9d0c014'  # SHAREPOINTWAC             — Office for the web (E1/E3/E5 and most M365 plans)
    '5dbe027f-2339-4123-9542-606e4d348a72'  # SHAREPOINTENTERPRISE      — SharePoint Online Plan 2 (E3/E5, Project, Dynamics)
    '902b47e5-dcb2-4fdc-858b-c63a90a2bdb9'  # SHAREPOINTDESKLESS        — SharePoint deskless (Teams Free, F-tier)
    '63038b2c-28d0-45f6-bc36-33062963b498'  # SHAREPOINTENTERPRISE_EDU  — SharePoint Plan 2 for Education
    '6b5b6a67-fc72-4a1f-a2b5-beecf05de761'  # SHAREPOINTENTERPRISE_MIDMARKET — SharePoint Plan 2 mid-market
    'c7699d2e-19aa-44de-8edf-1736da088ca1'  # SHAREPOINTSTANDARD        — SharePoint Online Plan 1 (standalone, Project P1)
    '0a4983bb-d3e5-4a09-95d8-b2d0127b3df5'  # SHAREPOINTSTANDARD_EDU   — SharePoint Plan 1 for Education
)
# HashSet for O(1) lookups inside the high-frequency user-enumeration loop
$AllOneDrivePlanIds = [System.Collections.Generic.HashSet[string]]::new()
foreach ($id in ($OneDrivePlanIds + $SharePointPlanIds)) { $AllOneDrivePlanIds.Add($id) | Out-Null }
#endregion Constants

#region Helper Functions

function Invoke-GraphRequestWithThrottleHandling {
    <#
    .SYNOPSIS
        Wraps Invoke-RestMethod with Retry-After / exponential-backoff throttle handling
        for Microsoft Graph API calls (429, 502, 503, 504, timeouts).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]   $Uri,
        [Parameter(Mandatory)] [string]   $Method,
        [Parameter()]          [hashtable] $Headers = @{},
        [Parameter()]          [string]    $Body = $null,
        [Parameter()]          [string]    $ContentType = 'application/json',
        [Parameter()]          [int]      $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]      $InitialBackoffSeconds = $script:InitialBackoffSec,
        [Parameter()]          [int]      $TimeoutSeconds = $script:RequestTimeoutSec
    )

    $retryCount = 0
    $backoffSec = $InitialBackoffSeconds
    $result = $null

    if ($debug) { Write-Host "  Graph -> $Method $Uri" -ForegroundColor DarkGray }

    while ($retryCount -le $MaxRetries) {
        try {
            $invokeParams = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ContentType = $ContentType
                TimeoutSec  = $TimeoutSeconds
                ErrorAction = 'Stop'
            }
            if ($Body) { $invokeParams['Body'] = $Body }

            $result = Invoke-RestMethod @invokeParams
            return $result
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $isRetryable = $statusCode -in @(429, 502, 503, 504) -or
            $_.Exception -is [System.Net.WebException] -and (
                $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout -or
                $_.Exception.Status -eq [System.Net.WebExceptionStatus]::ConnectionClosed
            )

            if (-not $isRetryable) { throw $_ }

            if ($retryCount -ge $MaxRetries) {
                Write-Host "    Max retries reached for: $Uri" -ForegroundColor Red
                throw $_
            }

            $waitSec = $backoffSec
            if ($statusCode -eq 429) {
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $waitSec = [int]$ra }
                }
                catch {}
            }

            $retryCount++
            Write-Host "    Throttled ($statusCode). Waiting ${waitSec}s (attempt $retryCount/$MaxRetries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $waitSec
            $backoffSec = [Math]::Min($backoffSec * 2, 300)
        }
    }
}

function ConvertTo-UPNFromSiteUrl {
    <#
    .SYNOPSIS
        Reconstructs a best-effort UPN from a SharePoint personal site URL.
        SharePoint encodes UPNs by lowercasing, replacing @ with _ and . with _.
        Example: John.Doe@contoso.com -> john_doe_contoso_com

        The tenant name is extracted from the hostname to locate the split point
        between username and domain in the encoded string.

        LIMITATION: Usernames containing . or _ are ambiguous after encoding
        (both map to _). The reconstructed UPN may differ from the original.
    #>
    param ([Parameter(Mandatory)] [string]$SiteUrl)

    # Pattern: https://<tenant>-my.sharepoint.com/personal/<encodedUPN>
    if ($SiteUrl -notmatch 'https://([^-]+)-my\.sharepoint\.com/personal/(.+)$') {
        return ''
    }

    $tenantName = $matches[1].ToLower()
    $encodedPart = $matches[2].ToLower().TrimEnd('/')

    # The domain portion begins at _<tenantName>_ in the encoded string.
    # Everything before that underscore-delimited boundary is the username.
    $domainSearch = "_$($tenantName)_"
    $domainIdx = $encodedPart.IndexOf($domainSearch, [System.StringComparison]::OrdinalIgnoreCase)

    if ($domainIdx -gt 0) {
        $userName = $encodedPart.Substring(0, $domainIdx)
        $domainEncoded = $encodedPart.Substring($domainIdx + 1)   # skip the leading _
        $domain = $domainEncoded.Replace('_', '.')
        return "$userName@$domain"
    }

    # Fallback: return the raw encoded form (caller can use DisplayName instead)
    return $encodedPart
}

#endregion Helper Functions

#region Authentication Functions

function AcquireToken {
    <#
    .SYNOPSIS
        Acquires a Microsoft Graph access token (scope: graph.microsoft.com/.default).
        One token covers all Graph endpoints across all geo datacenters.
    #>
    Write-Host "Authenticating to Microsoft Graph ($AuthType)..." -ForegroundColor Cyan

    $scope = 'https://graph.microsoft.com/.default'
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    if ($AuthType -eq 'ClientSecret') {
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = $scope
        }
        try {
            $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            $global:token = $resp.access_token
            $expiresIn = if ($resp.expires_in) { $resp.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  Connected via Client Secret. Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Authentication failed (ClientSecret): $($_.Exception.Message)" -ForegroundColor Red
            Exit
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        try {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
        }
        catch {
            Write-Host "  Certificate $Thumbprint not found in $CertStore\My store." -ForegroundColor Red
            Exit
        }

        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        $nbf = $now.ToUnixTimeSeconds()

        $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_') } | ConvertTo-Json -Compress
        $payload = @{ aud = $tokenUri; exp = $exp; iss = $clientId; jti = [System.Guid]::NewGuid().ToString(); nbf = $nbf; sub = $clientId } | ConvertTo-Json -Compress

        $hB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $pB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $toSign = "$hB64.$pB64"
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        if (-not $rsa) {
            Write-Host "  Unable to access RSA private key for certificate $Thumbprint." -ForegroundColor Red
            Exit
        }
        $sig = $rsa.SignData(
            [System.Text.Encoding]::UTF8.GetBytes($toSign),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $jwt = "$toSign.$([Convert]::ToBase64String($sig).TrimEnd('=').Replace('+', '-').Replace('/', '_'))"

        $body = @{
            client_id             = $clientId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwt
            scope                 = $scope
            grant_type            = 'client_credentials'
        }

        try {
            $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            $global:token = $resp.access_token
            $expiresIn = if ($resp.expires_in) { $resp.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  Connected via Certificate. Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Authentication failed (Certificate): $($_.Exception.Message)" -ForegroundColor Red
            Exit
        }
    }
    else {
        Write-Host "  Invalid AuthType '$AuthType'. Use 'Certificate' or 'ClientSecret'." -ForegroundColor Red
        Exit
    }
}

function Test-ValidToken {
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Write-Host 'Token expired or expiring soon — refreshing...' -ForegroundColor Yellow
        AcquireToken
    }
}

#endregion Authentication Functions

#region Data Collection Functions

function Get-ActiveUnlicensedOneDriveUsers {
    <#
    .SYNOPSIS
        Pages through ALL active Entra ID users, checks each user's assignedPlans,
        and returns those without an enabled OneDrive or SharePoint service plan.
        Only users with an existing OneDrive site appear in the final report —
        users with no drive are filtered out in Phase 4.
    #>
    Write-Host "`nPhase 1: Enumerating active users and checking OneDrive license plans..." -ForegroundColor Cyan

    $unlicensedUsers = [System.Collections.Generic.List[object]]::new()
    $totalScanned = 0

    $nextUri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,accountEnabled,assignedLicenses,assignedPlans&$top=999'

    do {
        Test-ValidToken
        $headers = @{ Authorization = "Bearer $global:token" }
        $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers

        foreach ($user in $response.value) {
            $totalScanned++

            $hasActivePlan = $false
            foreach ($plan in $user.assignedPlans) {
                if ($script:AllOneDrivePlanIds.Contains($plan.servicePlanId) -and
                    $plan.capabilityStatus -eq 'Enabled') {
                    $hasActivePlan = $true
                    break
                }
            }

            if (-not $hasActivePlan) {
                $unlicensedUsers.Add([PSCustomObject]@{
                        UserId            = $user.id
                        UserPrincipalName = $user.userPrincipalName
                        DisplayName       = $user.displayName
                        AccountEnabled    = $user.accountEnabled
                        HasAnyLicense     = ($user.assignedLicenses -and $user.assignedLicenses.Count -gt 0)
                        UserSource        = 'Active'
                        UnlicensedDate    = $null
                        UnlicensedDueTo   = 'License removed by admin'
                        DriveInfo         = $null
                    })
            }
        }

        $nextUri = $response.'@odata.nextLink'
        Write-Host "  Scanned $totalScanned users... $($unlicensedUsers.Count) without active OneDrive plan." -ForegroundColor Gray
    } while ($nextUri)

    Write-Host "  Active users scanned: $totalScanned | Unlicensed for OneDrive: $($unlicensedUsers.Count)" -ForegroundColor Green
    return $unlicensedUsers
}

function Get-SoftDeletedUsers {
    <#
    .SYNOPSIS
        Returns users in the Entra ID soft-delete recycle bin (deleted within 30 days).
        Their OneDrives still exist and are subject to the Day-60/Day-93 archival timeline.
        deletedDateTime is used as the unlicensed date.
        NOTE: Users deleted >30 days ago are permanently purged — not included in this report.
    #>
    Write-Host "`nPhase 2: Enumerating soft-deleted users (Entra ID 30-day recycle bin)..." -ForegroundColor Cyan

    $deletedUsers = [System.Collections.Generic.List[object]]::new()
    $nextUri = 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$select=id,userPrincipalName,displayName,deletedDateTime&$top=999'

    do {
        Test-ValidToken
        $headers = @{ Authorization = "Bearer $global:token" }
        $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers

        foreach ($user in $response.value) {
            $deletedDate = $null
            if ($user.deletedDateTime) {
                try { $deletedDate = [datetime]::Parse($user.deletedDateTime) } catch {}
            }

            $deletedUsers.Add([PSCustomObject]@{
                    UserId            = $user.id
                    UserPrincipalName = $user.userPrincipalName
                    DisplayName       = $user.displayName
                    AccountEnabled    = $false
                    HasAnyLicense     = $false
                    UserSource        = 'SoftDeleted'
                    UnlicensedDate    = $deletedDate
                    UnlicensedDueTo   = 'Owner deleted from Entra ID'
                    DriveInfo         = $null
                })
        }

        $nextUri = $response.'@odata.nextLink'
        Write-Host "  Soft-deleted users found so far: $($deletedUsers.Count)..." -ForegroundColor Gray
    } while ($nextUri)

    Write-Host "  Soft-deleted users: $($deletedUsers.Count)" -ForegroundColor Green
    return $deletedUsers
}

function Get-LicenseChangeDates {
    <#
    .SYNOPSIS
        Single bulk query of directoryAudits for 'Change user license' and
        'Remove user from licensed group' events. Returns userId -> most-recent-event-date
        lookup table. Requires AuditLog.Read.All.
    #>
    param (
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$TargetUserIds
    )

    Write-Host "`nPhase 3: Querying audit logs for license removal dates (bulk query)..." -ForegroundColor Cyan
    Write-Host "  Lookback: $AuditLogLookbackDays days | Requires AuditLog.Read.All" -ForegroundColor Gray

    $lookupTable = [System.Collections.Generic.Dictionary[string, datetime]]::new()
    $cutoffDate = (Get-Date).AddDays(-$AuditLogLookbackDays)

    # directoryAudits does not support 'or' on activityDisplayName in a single $filter.
    # The activityDateTime ge filter also causes 400 in some tenants, so date filtering
    # is done in PowerShell after retrieving results.
    $activityNames = @('Change user license', 'Remove user from licensed group')
    $eventCount = 0
    $queryFailed = $false

    foreach ($activityName in $activityNames) {
        $encodedFilter = [Uri]::EscapeDataString("activityDisplayName eq '$activityName'")
        $nextUri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=$encodedFilter&`$select=activityDateTime,targetResources&`$top=500"

        do {
            Test-ValidToken
            $headers = @{ Authorization = "Bearer $global:token" }

            try {
                $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers
            }
            catch {
                Write-Host "  Warning: Audit log query failed for '$activityName'. Verify AuditLog.Read.All is granted." -ForegroundColor Yellow
                Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
                $queryFailed = $true
                break
            }

            foreach ($auditEvent in $response.value) {
                $eventCount++
                $eventDate = $null
                try { $eventDate = [datetime]::Parse($auditEvent.activityDateTime) } catch { continue }
                if ($eventDate -lt $cutoffDate) { continue }

                foreach ($target in $auditEvent.targetResources) {
                    if (-not $target.id) { continue }
                    if (-not $TargetUserIds.Contains($target.id)) { continue }
                    if (-not $lookupTable.ContainsKey($target.id) -or $eventDate -gt $lookupTable[$target.id]) {
                        $lookupTable[$target.id] = $eventDate
                    }
                    break
                }
            }

            Write-Host "  [$activityName] Audit events processed: $eventCount | Dates found: $($lookupTable.Count)..." -ForegroundColor Gray
            $nextUri = $response.'@odata.nextLink'
        } while ($nextUri)

        if ($queryFailed) { break }
    }

    $missing = $TargetUserIds.Count - $lookupTable.Count
    Write-Host "  Audit scan complete. License removal dates found for $($lookupTable.Count) / $($TargetUserIds.Count) users." -ForegroundColor Green
    if ($missing -gt 0) {
        Write-Host ("  {0} users have no audit event within {1} days. UnlicensedDate will show 'Unknown'." -f $missing, $AuditLogLookbackDays) -ForegroundColor Yellow
    }
    return $lookupTable
}

function Get-UserDriveInfo {
    <#
    .SYNOPSIS
        Queries GET /users/{id}/drive for a single user. Graph routes this call to
        the correct geo datacenter automatically. Returns Found=$false for 404.
    #>
    param (
        [Parameter(Mandatory)] [string]$UserId,
        [Parameter(Mandatory)] [string]$UserPrincipalName
    )

    Test-ValidToken
    $headers = @{ Authorization = "Bearer $global:token" }
    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive"

    try {
        $drive = Invoke-GraphRequestWithThrottleHandling -Uri $uri -Method GET -Headers $headers

        $storageUsedGB = if ($drive.quota -and $null -ne $drive.quota.used) { [Math]::Round($drive.quota.used / 1GB, 3) } else { 0 }
        $storageTotalGB = if ($drive.quota -and $null -ne $drive.quota.total) { [Math]::Round($drive.quota.total / 1GB, 3) } else { 0 }

        if ($debug) { Write-Host "    [OK] $UserPrincipalName -> $($drive.webUrl)" -ForegroundColor DarkGreen }

        return [PSCustomObject]@{
            Found             = $true
            DriveId           = $drive.id
            DriveWebUrl       = $drive.webUrl
            DriveType         = $drive.driveType
            StorageUsedGB     = $storageUsedGB
            StorageTotalGB    = $storageTotalGB
            DriveLastModified = $drive.lastModifiedDateTime
            Note              = ''
        }
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }

        $note = switch ($statusCode) {
            404 { 'No OneDrive found (404) — never provisioned or already purged' }
            403 { 'Access denied (403) — check Files.Read.All permission' }
            $null { "Network error: $($_.Exception.Message)" }
            default { "HTTP $statusCode : $($_.Exception.Message)" }
        }

        if ($debug) { Write-Host "    [--] $UserPrincipalName : $note" -ForegroundColor DarkYellow }

        return [PSCustomObject]@{
            Found             = $false
            DriveId           = ''
            DriveWebUrl       = ''
            DriveType         = ''
            StorageUsedGB     = ''
            StorageTotalGB    = ''
            DriveLastModified = ''
            Note              = $note
        }
    }
}

function Get-SiteDriveInfo {
    <#
    .SYNOPSIS
        Queries GET /sites/{siteId}/drive for an archived OneDrive site.
        Used for Phase 2b (archived sites discovered via getAllSites) where no
        Entra user object exists, so GET /users/{id}/drive cannot be used.
        Returns Found=$true even on error since the site is known to exist.
    #>
    param (
        [Parameter(Mandatory)] [string]$SiteId,
        [Parameter(Mandatory)] [string]$SiteUrl
    )

    Test-ValidToken
    $headers = @{ Authorization = "Bearer $global:token" }
    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drive"

    try {
        $drive = Invoke-GraphRequestWithThrottleHandling -Uri $uri -Method GET -Headers $headers

        $storageUsedGB = if ($drive.quota -and $null -ne $drive.quota.used) { [Math]::Round($drive.quota.used / 1GB, 3) } else { 0 }
        $storageTotalGB = if ($drive.quota -and $null -ne $drive.quota.total) { [Math]::Round($drive.quota.total / 1GB, 3) } else { 0 }

        return [PSCustomObject]@{
            Found             = $true
            DriveId           = $drive.id
            DriveWebUrl       = if ($drive.webUrl) { $drive.webUrl } else { $SiteUrl }
            DriveType         = $drive.driveType
            StorageUsedGB     = $storageUsedGB
            StorageTotalGB    = $storageTotalGB
            DriveLastModified = $drive.lastModifiedDateTime
            Note              = ''
        }
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }

        $note = switch ($statusCode) {
            404 { 'Drive details unavailable — site may be fully archived or purged' }
            403 { 'Access denied (403) — check Files.Read.All / Sites.Read.All' }
            $null { "Network error: $($_.Exception.Message)" }
            default { "HTTP $statusCode : $($_.Exception.Message)" }
        }

        # Return Found=$true — we know the site exists from getAllSites even if drive query failed
        return [PSCustomObject]@{
            Found             = $true
            DriveId           = ''
            DriveWebUrl       = $SiteUrl
            DriveType         = ''
            StorageUsedGB     = ''
            StorageTotalGB    = ''
            DriveLastModified = ''
            Note              = $note
        }
    }
}

function Get-ArchivedOneDriveSites {
    <#
    .SYNOPSIS
        Queries GET /beta/sites/getAllSites to find personal OneDrive sites that Microsoft
        has already archived. These sites belong to users whose Entra account was
        deleted more than 30 days ago and whose OneDrive has entered archival.

        Requires Sites.Read.All (Application) on the app registration.

        Two-pass strategy to minimise per-site API calls:
          Pass 1 — Bulk: beta getAllSites with siteCollection in $select.
                   If archivalDetails.archiveStatus is returned inline, the site is
                   classified immediately — no per-site call required.
          Pass 2 — Per-site fallback: only for sites where archivalDetails was null
                   in the bulk response. Individual GET /beta/sites/{id}?$select=siteCollection
                   is used; HTTP 423 Locked is also treated as an archived signal.

        Sites with archiveStatus 'reactivating' or 'unknownFutureValue' are skipped.
    #>
    Write-Host "`nPhase 2b: Querying personal OneDrive sites for archived accounts..." -ForegroundColor Cyan
    Write-Host "  Requires Sites.Read.All permission on the app registration." -ForegroundColor Gray
    Write-Host "  Pass 1: bulk beta getAllSites (archivalDetails inline where available)." -ForegroundColor Gray

    $archivedSites = [System.Collections.Generic.List[object]]::new()
    $sitesNeedingCheck = [System.Collections.Generic.List[object]]::new()
    $totalScanned = 0

    # Pass 1: Bulk enumeration via the beta endpoint with siteCollection in $select.
    # On the beta endpoint, archivalDetails.archiveStatus is returned inline for archived
    # personal sites when siteCollection is explicitly selected. Sites where it comes back
    # populated are classified here with no further API call. Sites where it is null are
    # queued for the per-site fallback in Pass 2.
    $filterParam = [Uri]::EscapeDataString('isPersonalSite eq true')
    $nextUri = "https://graph.microsoft.com/beta/sites/getAllSites?`$filter=$filterParam&`$select=id,displayName,webUrl,isPersonalSite,siteCollection&`$top=200"

    do {
        Test-ValidToken
        $headers = @{ Authorization = "Bearer $global:token" }

        try {
            $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers
        }
        catch {
            Write-Host "  Warning: getAllSites query failed. Verify Sites.Read.All is granted." -ForegroundColor Yellow
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
            return $archivedSites
        }

        foreach ($site in $response.value) {
            $totalScanned++
            $archStatus = $site.siteCollection.archivalDetails.archiveStatus

            if ($null -ne $archStatus) {
                # archivalDetails returned inline — classify without a per-site call.
                if ($archStatus -in @('reactivating', 'unknownFutureValue')) { continue }

                $upn = ConvertTo-UPNFromSiteUrl -SiteUrl $site.webUrl
                $driveInfo = Get-SiteDriveInfo -SiteId $site.id -SiteUrl $site.webUrl

                $archNote = "archiveStatus: $archStatus"
                $driveInfo.Note = if ($driveInfo.Note) { "$archNote | $($driveInfo.Note)" } else { $archNote }

                if ($debug) { Write-Host "  [ARCHIVED-BULK] $archStatus — $($site.webUrl)" -ForegroundColor DarkGreen }

                $archivedSites.Add([PSCustomObject]@{
                        UserId            = ''       # No Entra user object — user purged from recycle bin
                        UserPrincipalName = $upn
                        DisplayName       = $site.displayName
                        AccountEnabled    = $false
                        HasAnyLicense     = $false
                        UserSource        = 'Archived'
                        UnlicensedDate    = $null    # Date unavailable — predates Entra purge (>30 days ago)
                        UnlicensedDueTo   = 'OneDrive archived by Microsoft'
                        ArchiveStatus     = $archStatus
                        DriveInfo         = $driveInfo
                    })
            }
            else {
                # archivalDetails was null in bulk response — queue for per-site fallback.
                $sitesNeedingCheck.Add($site)
            }
        }

        Write-Host "  Scanned: $totalScanned | Archived (bulk): $($archivedSites.Count) | Pending per-site check: $($sitesNeedingCheck.Count)..." -ForegroundColor Gray
        $nextUri = $response.'@odata.nextLink'
    } while ($nextUri)

    Write-Host "  Pass 1 complete. Archived (bulk): $($archivedSites.Count) | Sites needing per-site check: $($sitesNeedingCheck.Count)" -ForegroundColor Gray

    # Pass 2: Per-site fallback for sites where bulk response did not include archivalDetails.
    # GET /beta/sites/{id}?$select=id,siteCollection satisfies the "Requires $select" constraint
    # at the individual resource level and returns archivalDetails for archived sites.
    # HTTP 423 Locked is also treated as an archived signal — Graph refuses metadata
    # requests for archived sites and returns 423 instead of a response body.
    if ($sitesNeedingCheck.Count -gt 0) {
        Write-Host "  Pass 2: Per-site archival check for $($sitesNeedingCheck.Count) sites..." -ForegroundColor Gray
        $checked = 0
        foreach ($site in $sitesNeedingCheck) {
            $checked++
            if ($checked % 25 -eq 0 -or $checked -eq $sitesNeedingCheck.Count) {
                Write-Host "  Per-site check: $checked / $($sitesNeedingCheck.Count) | Archived found: $($archivedSites.Count)..." -ForegroundColor Gray
            }

            Test-ValidToken
            $headers = @{ Authorization = "Bearer $global:token" }
            $siteUri = "https://graph.microsoft.com/beta/sites/$($site.id)?`$select=id,siteCollection"

            try {
                $siteDetail = Invoke-GraphRequestWithThrottleHandling -Uri $siteUri -Method GET -Headers $headers
            }
            catch {
                $statusCode = $null
                if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }

                # HTTP 423 Locked = site is archived. Graph refuses the metadata request
                # for an archived site and returns 423 rather than a response body.
                if ($statusCode -eq 423) {
                    $upn = ConvertTo-UPNFromSiteUrl -SiteUrl $site.webUrl
                    if ($debug) { Write-Host "  [ARCHIVED-423] 423 Locked — $($site.webUrl)" -ForegroundColor DarkGreen }

                    $archivedSites.Add([PSCustomObject]@{
                            UserId            = ''
                            UserPrincipalName = $upn
                            DisplayName       = $site.displayName
                            AccountEnabled    = $false
                            HasAnyLicense     = $false
                            UserSource        = 'Archived'
                            UnlicensedDate    = $null
                            UnlicensedDueTo   = 'OneDrive archived by Microsoft'
                            ArchiveStatus     = 'archived'
                            DriveInfo         = [PSCustomObject]@{
                                Found             = $true
                                DriveId           = ''
                                DriveWebUrl       = $site.webUrl
                                DriveType         = 'business'
                                StorageUsedGB     = ''
                                StorageTotalGB    = ''
                                DriveLastModified = ''
                                Note              = 'Site is archived (HTTP 423 Locked) — storage details unavailable while archived'
                            }
                        })
                }
                else {
                    if ($debug) { Write-Host "  Warning: Could not get site details for $($site.webUrl): $($_.Exception.Message)" -ForegroundColor DarkYellow }
                }
                continue
            }

            # Dump raw JSON for the first successfully-returned site when debug is on
            if ($debug -and $checked -eq 1) {
                Write-Host "  [DEBUG] First per-site fallback response:" -ForegroundColor DarkGray
                Write-Host ($siteDetail | ConvertTo-Json -Depth 6) -ForegroundColor DarkGray
            }

            $archStatus = $siteDetail.siteCollection.archivalDetails.archiveStatus
            if ($null -eq $archStatus) { continue }
            if ($archStatus -in @('reactivating', 'unknownFutureValue')) { continue }

            $upn = ConvertTo-UPNFromSiteUrl -SiteUrl $site.webUrl
            $driveInfo = Get-SiteDriveInfo -SiteId $site.id -SiteUrl $site.webUrl

            $archNote = "archiveStatus: $archStatus"
            $driveInfo.Note = if ($driveInfo.Note) { "$archNote | $($driveInfo.Note)" } else { $archNote }

            $archivedSites.Add([PSCustomObject]@{
                    UserId            = ''       # No Entra user object — user purged from recycle bin
                    UserPrincipalName = $upn
                    DisplayName       = $site.displayName
                    AccountEnabled    = $false
                    HasAnyLicense     = $false
                    UserSource        = 'Archived'
                    UnlicensedDate    = $null    # Date unavailable — predates Entra purge (>30 days ago)
                    UnlicensedDueTo   = 'OneDrive archived by Microsoft'
                    ArchiveStatus     = $archStatus
                    DriveInfo         = $driveInfo
                })
        }
    }

    Write-Host "  Sites enumeration complete. Archived personal OneDrives: $($archivedSites.Count)" -ForegroundColor Green
    return $archivedSites
}

#endregion Data Collection Functions

#region Enrichment Functions

function Add-MilestoneCalculations {
    <#
    .SYNOPSIS
        Enriches a list of unlicensed OneDrive account objects with Day-60 / Day-93
        milestone dates, days-remaining counters, and a traffic-light urgency label.
    #>
    param (
        [Parameter(Mandatory)] [object[]]$Accounts
    )

    $enriched = [System.Collections.Generic.List[object]]::new()

    foreach ($acct in $Accounts) {
        $unlicensedDate = $acct.UnlicensedDate
        $readOnlyDate = $null
        $archiveDate = $null
        $daysSinceUnlicensed = $null
        $daysUntilReadOnly = $null
        $daysUntilArchive = $null
        $urgencyStatus = 'Unknown - No Unlicensed Date'

        if ($unlicensedDate) {
            $readOnlyDate = $unlicensedDate.AddDays($script:ReadOnlyThresholdDays)
            $archiveDate = $unlicensedDate.AddDays($script:ArchiveThresholdDays)
            $daysSinceUnlicensed = ($script:today - $unlicensedDate.Date).Days
            $daysUntilReadOnly = ($readOnlyDate.Date - $script:today).Days
            $daysUntilArchive = ($archiveDate.Date - $script:today).Days

            $urgencyStatus = switch ($true) {
                ($daysUntilArchive -lt 0) { 'ARCHIVED - Past Day 93' }
                ($daysUntilArchive -eq 0) { 'CRITICAL - Archives TODAY' }
                ($daysUntilArchive -le 7) { 'CRITICAL - Archives within 7 days' }
                ($daysUntilReadOnly -lt 0 -and $daysUntilArchive -gt 7) { 'WARNING - Read-Only, Archive pending' }
                ($daysUntilReadOnly -eq 0) { 'WARNING - Goes Read-Only TODAY' }
                ($daysUntilReadOnly -le 7) { 'WARNING - Read-Only within 7 days' }
                ($daysUntilReadOnly -le 30) { 'MONITOR - Read-Only within 30 days' }
                default { 'OK - More than 30 days remaining' }
            }
        }

        # For Archived population (Phase 2b / Sites API): no UnlicensedDate is available
        # since the Entra user was purged >30 days ago. Set urgency from archiveStatus.
        if (-not $unlicensedDate -and $acct.UserSource -eq 'Archived') {
            $urgencyStatus = switch ($acct.ArchiveStatus) {
                'fullyArchived' { 'ARCHIVED - Fully Archived' }
                'recentlyArchived' { 'ARCHIVED - Recently Archived' }
                default { 'ARCHIVED - Currently Archived' }
            }
        }

        $driveInfo = $acct.DriveInfo

        $enriched.Add([PSCustomObject]@{
                UserSource          = $acct.UserSource
                DisplayName         = $acct.DisplayName
                UserPrincipalName   = $acct.UserPrincipalName
                AccountEnabled      = $acct.AccountEnabled
                UnlicensedDueTo     = $acct.UnlicensedDueTo
                UnlicensedDate      = if ($unlicensedDate) { $unlicensedDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                DaysSinceUnlicensed = $daysSinceUnlicensed
                ReadOnlyDate        = if ($readOnlyDate) { $readOnlyDate.ToString('yyyy-MM-dd') } else { '' }
                ArchiveDate         = if ($archiveDate) { $archiveDate.ToString('yyyy-MM-dd') } else { '' }
                DaysUntilReadOnly   = $daysUntilReadOnly
                DaysUntilArchive    = $daysUntilArchive
                UrgencyStatus       = $urgencyStatus
                StorageUsedGB       = $driveInfo.StorageUsedGB
                StorageTotalGB      = $driveInfo.StorageTotalGB
                DriveUrl            = $driveInfo.DriveWebUrl
                DriveLastModified   = $driveInfo.DriveLastModified
                DriveType           = $driveInfo.DriveType
                Notes               = $driveInfo.Note
            })
    }

    return $enriched
}

#endregion Enrichment Functions

#region Output Functions

function Write-ConsoleSummary {
    param ([object[]]$Records)

    Write-Host "`n======================================================" -ForegroundColor Cyan
    Write-Host "  UNLICENSED ONEDRIVE REPORT — SUMMARY" -ForegroundColor Cyan
    Write-Host ("  Run date : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
    Write-Host "  Tenant   : $tenantId" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ("  Total accounts in report: {0}" -f $Records.Count) -ForegroundColor White

    Write-Host "`n  By Source:" -ForegroundColor White
    $Records | Group-Object UserSource | Sort-Object Name | ForEach-Object {
        Write-Host ("    {0,-15} {1,5} accounts" -f $_.Name, $_.Count) -ForegroundColor Gray
    }

    Write-Host "`n  By Urgency (most critical first):" -ForegroundColor White
    $Records | Group-Object UrgencyStatus | Sort-Object @{
        Expression = {
            switch -Wildcard ($_.Name) {
                'CRITICAL*' { 1 } 'ARCHIVED*' { 2 } 'WARNING*' { 3 }
                'MONITOR*' { 4 } 'OK*' { 5 } default { 6 }
            }
        }
    } | ForEach-Object {
        $color = switch -Wildcard ($_.Name) {
            'CRITICAL*' { 'Red' }
            'ARCHIVED*' { 'DarkRed' }
            'WARNING*' { 'Yellow' }
            'MONITOR*' { 'Magenta' }
            'OK*' { 'Green' }
            default { 'Gray' }
        }
        Write-Host ("    {0,-45} {1,5} accounts" -f $_.Name, $_.Count) -ForegroundColor $color
    }
}

#endregion Output Functions

#region Main Execution

Write-Host '======================================================' -ForegroundColor Cyan
Write-Host '  Unlicensed OneDrive Report — Microsoft Graph API' -ForegroundColor Cyan
Write-Host '  Multi-geo: all geos covered by single Graph token' -ForegroundColor Cyan
Write-Host ("  Run date : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host ("Timeline  : Day {0} = Read-Only  |  Day {1} = Archived/Deleted" -f $ReadOnlyThresholdDays, $ArchiveThresholdDays) -ForegroundColor White
Write-Host ''

# Step 1: Authenticate — single Graph token, covers all geo datacenters
AcquireToken

# Step 2 (Phase 1): Active users without an enabled OneDrive license plan
$activeUnlicensed = Get-ActiveUnlicensedOneDriveUsers

# Step 3 (Phase 2): Soft-deleted users (Entra 30-day recycle bin)
$softDeletedUsers = Get-SoftDeletedUsers

# Step 4: Merge both populations, de-duplicate on UserId
$activeUserIds = [System.Collections.Generic.HashSet[string]]($activeUnlicensed | Select-Object -ExpandProperty UserId)
$deletedFiltered = $softDeletedUsers | Where-Object { -not $activeUserIds.Contains($_.UserId) }

$allCandidates = [System.Collections.Generic.List[object]]::new()
foreach ($u in $activeUnlicensed) { $allCandidates.Add($u) }
foreach ($u in $deletedFiltered) { $allCandidates.Add($u) }

Write-Host ''
Write-Host "Total candidates to check for OneDrive: $($allCandidates.Count)" -ForegroundColor White
Write-Host "  Active unlicensed : $($activeUnlicensed.Count)" -ForegroundColor Gray
Write-Host "  Soft-deleted      : $($($deletedFiltered).Count)" -ForegroundColor Gray

if ($allCandidates.Count -eq 0 -and -not $GetCurrentlyArchived) {
    Write-Host "`nNo unlicensed candidates found. Exiting." -ForegroundColor Green
    Exit
}

# Step 4b (Phase 2b): Archived OneDrive sites discovered via GET /sites/getAllSites.
# These are personal OneDrive sites archived by Microsoft whose Entra user was deleted
# >30 days ago (purged from the recycle bin). Drive info is gathered inside the function
# via GET /sites/{siteId}/drive, so Phase 4 does not process these.
$archivedSites = [System.Collections.Generic.List[object]]::new()
if ($GetCurrentlyArchived) {
    $rawArchivedSites = Get-ArchivedOneDriveSites

    # Deduplicate: if a UPN from the Sites API already exists in $allCandidates (e.g., a user
    # soft-deleted within the 30-day window also shows up in getAllSites), keep the Entra record.
    $existingUpns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($c in $allCandidates) {
        if ($c.UserPrincipalName) { $existingUpns.Add($c.UserPrincipalName) | Out-Null }
    }
    foreach ($s in $rawArchivedSites) {
        if ($s.UserPrincipalName -and $existingUpns.Contains($s.UserPrincipalName)) {
            if ($debug) { Write-Host "  Dedup: $($s.UserPrincipalName) already in Entra candidates — skipping archived site entry." -ForegroundColor DarkGray }
            continue
        }
        $archivedSites.Add($s)
    }
    if ($rawArchivedSites.Count -gt 0) {
        Write-Host "  Archived sites added (after dedup): $($archivedSites.Count) of $($rawArchivedSites.Count)" -ForegroundColor Gray
    }
}
else {
    Write-Host "`nPhase 2b: Skipped (`$GetCurrentlyArchived = `$false)." -ForegroundColor Gray
}

if ($allCandidates.Count -eq 0 -and $archivedSites.Count -eq 0) {
    Write-Host "`nNo unlicensed or archived OneDrive accounts found. Exiting." -ForegroundColor Green
    Exit
}

# Step 5 (Phase 3): Bulk audit log query for license-removal dates (active users only)
if ($includeLicenseRemovalDates -and $activeUnlicensed.Count -gt 0) {
    $licenseChangeDates = Get-LicenseChangeDates -TargetUserIds $activeUserIds
    foreach ($user in $activeUnlicensed) {
        if ($licenseChangeDates.ContainsKey($user.UserId)) {
            $user.UnlicensedDate = $licenseChangeDates[$user.UserId]
        }
    }
}
else {
    $reason = if (-not $includeLicenseRemovalDates) { '($includeLicenseRemovalDates = $false)' } else { '(no active unlicensed users)' }
    Write-Host "`nPhase 3: Skipped $reason — active users will show UnlicensedDate = 'Unknown'." -ForegroundColor Gray
}

# Step 6 (Phase 4): Query OneDrive for each candidate
# Graph routes each /users/{id}/drive call to the correct geo — no per-geo iteration needed.
Write-Host "`nPhase 4: Querying OneDrive drive info for $($allCandidates.Count) candidates..." -ForegroundColor Cyan
Write-Host "  (Graph routes each call to the correct geo datacenter automatically)" -ForegroundColor Gray

$confirmedUnlicensed = [System.Collections.Generic.List[object]]::new()
$total = $allCandidates.Count
$current = 0
$driveFound = 0
$driveNotFound = 0
$driveErrors = 0

foreach ($user in $allCandidates) {
    $current++
    $pct = [Math]::Round(($current / $total) * 100)
    Write-Progress -Activity 'Querying OneDrive' `
        -Status   "$current / $total ($pct%) — $($user.UserPrincipalName)" `
        -PercentComplete $pct

    $driveInfo = Get-UserDriveInfo -UserId $user.UserId -UserPrincipalName $user.UserPrincipalName
    $user.DriveInfo = $driveInfo

    if ($driveInfo.Found) {
        $driveFound++
        $confirmedUnlicensed.Add($user)
    }
    elseif ($driveInfo.Note -match '404') {
        $driveNotFound++
        # Pure 404 = never provisioned or already purged; skip.
    }
    else {
        $driveErrors++
        # 403/timeouts — include in report for admin review.
        $confirmedUnlicensed.Add($user)
    }

    if ($delayBetweenRequests -gt 0) { Start-Sleep -Seconds $delayBetweenRequests }
}

Write-Progress -Activity 'Querying OneDrive' -Completed
Write-Host "  Drive queries complete." -ForegroundColor Green
Write-Host "  OneDrive found      : $driveFound" -ForegroundColor Green
Write-Host "  No OneDrive (404)   : $driveNotFound  (skipped — never provisioned or purged)" -ForegroundColor Gray
Write-Host "  Drive errors        : $driveErrors  (403/timeouts — included in report for admin review)" -ForegroundColor Yellow

# Merge archived sites (Phase 2b) into confirmed list.
# Their DriveInfo was already populated by Get-SiteDriveInfo inside Get-ArchivedOneDriveSites.
if ($archivedSites.Count -gt 0) {
    foreach ($s in $archivedSites) { $confirmedUnlicensed.Add($s) }
    Write-Host "  Archived sites merged (Phase 2b): $($archivedSites.Count)" -ForegroundColor Green
}

# Step 7 (Phase 5): Enrich with Day-$ReadOnlyThresholdDays / Day-$ArchiveThresholdDays milestones
Write-Host "`nPhase 5: Calculating Day-$ReadOnlyThresholdDays / Day-$ArchiveThresholdDays milestones..." -ForegroundColor Cyan
$enriched = if ($confirmedUnlicensed.Count -gt 0) {
    Add-MilestoneCalculations -Accounts $confirmedUnlicensed
}
else { @() }

# Step 8: Sort by urgency (most critical first), then days until archive
$sorted = $enriched | Sort-Object @(
    @{
        Expression = {
            switch -Wildcard ($_.UrgencyStatus) {
                'CRITICAL*' { 1 } 'ARCHIVED*' { 2 } 'WARNING*' { 3 }
                'MONITOR*' { 4 } 'OK*' { 5 } default { 6 }
            }
        }
    },
    @{ Expression = 'DaysUntilArchive'; Ascending = $true }
)

# Step 9: Export report
if ($sorted.Count -gt 0) {
    # Write with UTF-8 BOM so Excel opens the file without character garbling
    [System.IO.File]::WriteAllLines($outputLog, ($sorted | ConvertTo-Csv -NoTypeInformation), [System.Text.Encoding]::UTF8)
    Write-ConsoleSummary -Records $sorted

    Write-Host "`n======================================================" -ForegroundColor Cyan
    Write-Host "  Report written: $outputLog" -ForegroundColor Green
    Write-Host "  Total records : $($sorted.Count)" -ForegroundColor Green
    Write-Host '======================================================' -ForegroundColor Cyan
}
else {
    Write-Host "`nNo unlicensed OneDrive accounts with active drives found." -ForegroundColor Green
}

#endregion Main Execution
