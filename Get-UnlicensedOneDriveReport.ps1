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
    Author      : Mike Lee | Mariel Williams
    Date Created: 4/28/26
    Date Updated: 4/30/26 added cost estimation and email notification features
    Date Updated: 5/1/26: 
    - Fixed performance issue in Get-LicenseChangeDates by doing a single bulk query instead of per-user queries. 
    - Added methods to clear memory and dispose of HTTP responses properly.
    Date Updated: 6/11/26: 
    - Added support for multi-geo scenarios
    - Added download functionality for existing OneDrive Archive reports.
    - Merged the download functionality with the existing report generation.
    Date: 6/12/26
    - Added throttling for Audit Queries
    

    Required Microsoft Graph App Permissions (Application type):
      User.Read.All           — Enumerate users and inspect assignedPlans/licenses
      Directory.Read.All      — Read soft-deleted users from Entra recycle bin
      Files.Read.All          — Read OneDrive drive metadata for any user
      AuditLog.Read.All       — [OPTIONAL] directoryAudits for license-change dates
                                 Set $includeLicenseRemovalDates = $false to skip.
      Sites.Read.All          — [OPTIONAL] GET /sites/getAllSites for currently archived OneDrive sites
                                 Set $GetCurrentlyArchived = $false to skip.
      Mail.Send               — [OPTIONAL] Send alert emails via Graph API (POST /users/{sender}/sendMail)
                                 Set $SendEmailNotifications = $false to skip.
                                 The $EmailFrom mailbox must be a licensed Exchange Online mailbox.
    
    Required SharePoint Permissions (Application type):
    Sites.FullControl.All - Required to Download OneDrive Reports from the Admin API

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

# ---- Post-archive deletion timeline (MC1381110) ----
# If PAYG is not enabled, archived OneDrive accounts are subject to deletion
# after this many days from unlicensed date (even if retention hold exists).
$ArchiveDeletionThresholdDays = 365

# PAYG status for archived unlicensed OneDrive accounts.
# IMPORTANT: There is no reliable API/cmdlet readback for this tenant toggle,
# so this script uses a manual value.
$PayGEnabledForUnlicensedOneDrive = $false

# ---- Audit log: license removal date discovery for ACTIVE unlicensed users ----
# When $true: queries directoryAudits (bulk) to find when each active user's
# OneDrive/SharePoint license was removed. Requires AuditLog.Read.All.
# When $false: active users will show UnlicensedDate as 'Unknown'.
$includeLicenseRemovalDates = $true

# How far back to search for license-change audit events (max 180 days).
$AuditLogLookbackDays = 180

# ---- Currently archived OneDrive accounts ----
# Master switch. When $true, include currently archived OneDrive accounts.
# When $false, only reports active/soft-deleted populations.
$GetCurrentlyArchived = $true

# Primary source for archived accounts:
#   'SPODownload' (recommended): SharePoint Admin ExportToCSV download first.
#   'GraphSites'               : Graph Sites API enumeration first.
$ArchivedCollectionMode = 'SPODownload'

# SharePoint Admin URLs used by the SPO export downloader (one per geo location).
# Format: https://<tenant>-admin.sharepoint.com (no trailing slash)
$SPOAdminUrls = @(
    'https://m365cpi13246019-admin.sharepoint.com'
)

# Poll settings for SPO ExportToCSV download.
$SPOExportPollIntervalSec = 5
$SPOExportMaxWaitSec = 120

# When $true, all downloaded SPO unlicensed reports (across admin URLs/geos)
# are merged into a single CSV in $OutputFolder.
$MergeDownloadedSPOReports = $true

# When $true, rows from downloaded SPO reports are used to backfill/add records
# in the final report (useful when audit lookup does not return dates or users).
$IncludeDownloadedRowsInMainReport = $true

# ---- Request throttling ----
$MaxRetries = 15
$InitialBackoffSec = 3
$RequestTimeoutSec = 300

# ---- Delay between individual drive queries (seconds). 0 = no delay. ----
$delayBetweenRequests = 0

# ---- Cost estimation rates — per Microsoft unlicensed OneDrive pricing ----
# https://learn.microsoft.com/en-us/sharepoint/unlicensed-onedrive-accounts
$ArchivalStorageCostPerGBMonth = 0.05   # USD per GB per month — ongoing storage fee for archived OneDrives (past Day 93)
$ReactivationCostPerGB = 0.60   # USD per GB one-time fee to reactivate an archived OneDrive

# ---- Email Notifications ----
# Set $SendEmailNotifications = $true to send alert emails to admins after the report runs.
# Three separate emails are sent when accounts fall within the configured day thresholds:
#   (1) Approaching Read-Only  — when DaysUntilReadOnly  <= $DaysToNotifyBeforeReadOnly
#   (2) Approaching Archive    — when DaysUntilArchive   <= $DaysToNotifyBeforeArchive
#   (3) Approaching Deletion   — when DaysUntilDeletion  <= $DaysToNotifyBeforeDeletion
$SendEmailNotifications = $false

# Recipients — individual addresses or mail-enabled group/distribution-list addresses.
$EmailTo = @(
    'admin@contoso.onmicrosoft.com'
    'Test-Email-Security-Group@contoso.onmicrosoft.com'
)

# Sender address — must be a licensed Exchange Online mailbox in the tenant.
# The app registration must have Mail.Send (Application) permission granted in Entra ID.
# Email is sent via Graph API (POST /users/{EmailFrom}/sendMail) — no SMTP relay needed.
$EmailFrom = 'admin@contoso.onmicrosoft.com'

# Notification windows — an alert email is sent when a site's days-until-event falls
# at or below this value. Set to 0 to only notify on the day of the event itself.
$DaysToNotifyBeforeReadOnly = 14   # Notify admins this many days before a site goes read-only
$DaysToNotifyBeforeArchive = 14   # Notify admins this many days before a site is archived
$DaysToNotifyBeforeDeletion = 30   # Notify admins this many days before an archived site reaches deletion risk window

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

# SPO admin token cache, keyed by admin URL (different auth audience per geo).
$global:spoTokenByAdminUrl = @{}
$global:spoDownloadedReportFiles = [System.Collections.Generic.List[string]]::new()
$global:spoDownloadedReportRows = [System.Collections.Generic.List[object]]::new()
$global:spoMergedDownloadReportPath = ''
$global:tenantPayGStatus = $null

# Required for HTML-encoding display names and UPNs in alert email bodies
Add-Type -AssemblyName System.Web
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
                Verbose     = $false
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

function Get-ObjectPropertyValue {
    <#
    .SYNOPSIS
        Returns the first available property value from a candidate list.
        Supports both exact and case-insensitive property-name lookups.
    #>
    param (
        [Parameter(Mandatory)] [object]$InputObject,
        [Parameter(Mandatory)] [string[]]$CandidateNames
    )

    foreach ($name in $CandidateNames) {
        $normalizedCandidate = "$name".Trim().Trim([char]0xFEFF)

        $prop = $InputObject.PSObject.Properties | Where-Object {
            $_.Name -and $_.Name.ToString().Trim().Trim([char]0xFEFF) -ieq $normalizedCandidate
        } | Select-Object -First 1

        if ($prop) {
            $value = $prop.Value
            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                return $value
            }
        }
    }

    return $null
}

function Get-TenantPayGStatus {
    <#
    .SYNOPSIS
        Returns PAYG status for archived OneDrive reactivation from manual config.
        Returns:
          @{ IsEnabled = [bool]; DetectionMode = 'Manual'; Message = '...' }
    #>
    if ($global:tenantPayGStatus) { return $global:tenantPayGStatus }

    $global:tenantPayGStatus = [PSCustomObject]@{
        IsEnabled     = [bool]$PayGEnabledForUnlicensedOneDrive
        DetectionMode = 'Manual'
        Message       = if ($PayGEnabledForUnlicensedOneDrive) { 'PAYG enabled (manual config)' } else { 'PAYG not enabled (manual config)' }
    }
    return $global:tenantPayGStatus
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
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
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
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
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

function Get-SPOTokenForAdminUrl {
    <#
    .SYNOPSIS
        Acquires or refreshes an SPO admin token for a specific admin URL.
        Scope is <adminUrl>/.default, which is distinct per geo admin host.
    #>
    param (
        [Parameter(Mandatory)] [string]$AdminUrl
    )

    $cached = $global:spoTokenByAdminUrl[$AdminUrl]
    if ($cached -and $cached.expiry -and (Get-Date) -lt $cached.expiry) {
        return $cached.access_token
    }

    $scope = "$AdminUrl/.default"
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    if ($AuthType -eq 'ClientSecret') {
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = $scope
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        $nbf = $now.ToUnixTimeSeconds()

        $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_') } | ConvertTo-Json -Compress
        $payload = @{ aud = $tokenUri; exp = $exp; iss = $clientId; jti = [System.Guid]::NewGuid().ToString(); nbf = $nbf; sub = $clientId } | ConvertTo-Json -Compress

        $hB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $pB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $toSign = "$hB64.$pB64"
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        if (-not $rsa) { throw "Unable to access RSA private key for certificate $Thumbprint." }
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
    }
    else {
        throw "Invalid AuthType '$AuthType'. Use 'Certificate' or 'ClientSecret'."
    }

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
        -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false

    $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
    $expiry = (Get-Date).AddSeconds($expiresIn - 300)
    $global:spoTokenByAdminUrl[$AdminUrl] = @{ access_token = $resp.access_token; expiry = $expiry }

    Write-Host "  SPO token acquired for $AdminUrl. Valid until: $expiry" -ForegroundColor Green
    return $resp.access_token
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
    $cutoffDateUtc = $cutoffDate.ToUniversalTime().ToString('o')

    # directoryAudits does not support 'or' on activityDisplayName in a single $filter.
    # Prefer a server-side activityDateTime filter to cut request volume, then fall back
    # to client-side cutoff logic if the tenant rejects the combined filter/orderby.
    $activityNames = @('Change user license', 'Remove user from licensed group')
    $eventCount = 0
    $queryFailed = $false

    foreach ($activityName in $activityNames) {
        $filterWithDate = [Uri]::EscapeDataString("activityDisplayName eq '$activityName' and activityDateTime ge $cutoffDateUtc")
        $filterWithoutDate = [Uri]::EscapeDataString("activityDisplayName eq '$activityName'")
        $queryModes = @(
            @{
                Name            = 'server-side date filter'
                Uri             = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=$filterWithDate&`$select=activityDateTime,targetResources&`$orderby=activityDateTime desc&`$top=500"
                AllowsEarlyStop = $false
            },
            @{
                Name            = 'fallback scan'
                Uri             = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=$filterWithoutDate&`$select=activityDateTime,targetResources&`$orderby=activityDateTime desc&`$top=500"
                AllowsEarlyStop = $true
            }
        )
        $completedActivity = $false

        foreach ($queryMode in $queryModes) {
            $nextUri = $queryMode.Uri
            $shouldTryNextMode = $false

            do {
                Test-ValidToken
                $headers = @{ Authorization = "Bearer $global:token" }

                try {
                    $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextUri -Method GET -Headers $headers
                }
                catch {
                    $statusCode = $null
                    if ($_.Exception.Response) {
                        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                    }

                    if ($statusCode -eq 400 -and $queryMode.Name -eq 'server-side date filter') {
                        Write-Host "  '$activityName' rejected the date-bounded filter. Retrying with fallback scan..." -ForegroundColor Yellow
                        $shouldTryNextMode = $true
                        break
                    }

                    Write-Host "  Warning: Audit log query failed for '$activityName'. Verify AuditLog.Read.All is granted." -ForegroundColor Yellow
                    Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
                    $queryFailed = $true
                    break
                }

                $pageNewestEventDate = $null

                foreach ($auditEvent in $response.value) {
                    $eventCount++
                    $eventDate = $null
                    try { $eventDate = [datetime]::Parse($auditEvent.activityDateTime) } catch { continue }
                    if (-not $pageNewestEventDate -or $eventDate -gt $pageNewestEventDate) {
                        $pageNewestEventDate = $eventDate
                    }
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

                Write-Host "  [$activityName][$($queryMode.Name)] Audit events processed: $eventCount | Dates found: $($lookupTable.Count)..." -ForegroundColor Gray
                $nextUri = $response.'@odata.nextLink'

                if ($queryMode.AllowsEarlyStop -and $pageNewestEventDate -and $pageNewestEventDate -lt $cutoffDate) {
                    Write-Host "  [$activityName] Remaining audit pages are older than the lookback window. Stopping fallback scan early." -ForegroundColor Gray
                    $nextUri = $null
                }
            } while ($nextUri)

            if ($queryFailed) { break }
            if ($shouldTryNextMode) { continue }

            $completedActivity = $true
            break
        }

        if ($queryFailed) { break }
        if (-not $completedActivity) { break }
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
                # Store only the three fields Pass 2 uses — avoids holding the full
                # paged response objects in memory for tenants with many personal sites.
                $sitesNeedingCheck.Add([PSCustomObject]@{
                        id          = $site.id
                        webUrl      = $site.webUrl
                        displayName = $site.displayName
                    })
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
        $siteCount = $sitesNeedingCheck.Count
        $checked = 0
        for ($i = 0; $i -lt $siteCount; $i++) {
            $site = $sitesNeedingCheck[$i]
            $sitesNeedingCheck[$i] = $null   # release reference so GC can reclaim after this iteration
            if ($i -gt 0 -and $i % 500 -eq 0) { [System.GC]::Collect() }  # periodic GC hint for large tenants
            $checked++
            if ($checked % 25 -eq 0 -or $checked -eq $siteCount) {
                Write-Host "  Per-site check: $checked / $siteCount | Archived found: $($archivedSites.Count)..." -ForegroundColor Gray
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

function Download-UnlicensedOneDriveCsvFromSPO {
    <#
    .SYNOPSIS
        Downloads the same unlicensed OneDrive CSV report exposed by
        SharePoint Admin Center "Download report" for a given admin URL.
    #>
    param (
        [Parameter(Mandatory)] [string]$AdminUrl,
        [Parameter()]          [string]$OutputPath = $OutputFolder
    )

    $spoToken = Get-SPOTokenForAdminUrl -AdminUrl $AdminUrl

    $contextHeaders = @{
        Authorization = "Bearer $spoToken"
        Accept        = 'application/json;odata=verbose'
    }
    $contextInfo = Invoke-GraphRequestWithThrottleHandling -Uri "$AdminUrl/_api/contextinfo" -Method POST -Headers $contextHeaders -ContentType 'application/json;odata=verbose'
    $digest = $contextInfo.d.GetContextWebInformation.FormDigestValue

    $viewXml = '<View><Query><Where><And><And>' +
    '<And><And>' +
    '<IsNotNull><FieldRef Name="UnlicensedOdbReason"/></IsNotNull>' +
    '<Neq><FieldRef Name="UnlicensedOdbReason"/><Value Type=''Integer''>0</Value></Neq>' +
    '</And>' +
    '<IsNotNull><FieldRef Name="UnlicensedOdbCleanupBlockReason"/></IsNotNull>' +
    '</And>' +
    '<And>' +
    '<Eq><FieldRef Name="TemplateId"/><Value Type=''Integer''>21</Value></Eq>' +
    '<IsNull><FieldRef Name="TimeDeleted"/></IsNull>' +
    '</And>' +
    '</And>' +
    '<And>' +
    '<Neq><FieldRef Name=''TemplateName''/><Value Type=''Text''>TEAMCHANNEL#0</Value></Neq>' +
    '<Neq><FieldRef Name=''TemplateName''/><Value Type=''Text''>TEAMCHANNEL#1</Value></Neq>' +
    '</And></And></Where></Query>' +
    '<ViewFields>' +
    '<FieldRef Name="Title"/><FieldRef Name="SiteOwnerName"/><FieldRef Name="StorageUsed"/>' +
    '<FieldRef Name="UnlicensedOdbReason"/><FieldRef Name="UnlicensedOdbStartDate"/>' +
    '<FieldRef Name="UnlicensedOdbCleanupBlockReason"/><FieldRef Name="SiteOwnerEmail"/>' +
    '<FieldRef Name="UnlicensedOdbToBeDeletedOn"/><FieldRef Name="ArchiveStatus"/>' +
    '<FieldRef Name="UnlicensedOdbProvisionedForUPN"/><FieldRef Name="SiteUrl"/>' +
    '</ViewFields></View>'

    $columnsInfo = @(
        @{ columnName = 'TITLE'; viewFieldName = 'Title' }
        @{ columnName = 'PRIMARY_ADMIN'; viewFieldName = 'SiteOwnerName' }
        @{ columnName = 'STORAGE_USED'; viewFieldName = 'StorageUsed' }
        @{ columnName = 'UNLICENSED_REASON'; viewFieldName = 'UnlicensedOdbReason' }
        @{ columnName = 'UNLICENSED_ON'; viewFieldName = 'UnlicensedOdbStartDate' }
        @{ columnName = 'DELETION_BLOCK_REASON'; viewFieldName = 'UnlicensedOdbCleanupBlockReason' }
        @{ columnName = 'SITE_OWNER_EMAIL'; viewFieldName = 'SiteOwnerEmail' }
        @{ columnName = 'DELETION_SCHEDULED_ON'; viewFieldName = 'UnlicensedOdbToBeDeletedOn' }
        @{ columnName = 'ARCHIVE_STATUS'; viewFieldName = 'ArchiveStatus' }
        @{ columnName = 'ACCOUNT_PROVISIONED_FOR'; viewFieldName = 'UnlicensedOdbProvisionedForUPN' }
        @{ columnName = 'URL'; viewFieldName = 'SiteUrl' }
    )

    $exportBody = [ordered]@{
        viewXml     = $viewXml
        columnsInfo = $columnsInfo
        listName    = 'DO_NOT_DELETE_SPLIST_TENANTADMIN_ALL_SITES_AGGREGATED_SITECOLLECTIONS'
    } | ConvertTo-Json -Depth 5 -Compress

    $postHeaders = @{
        Authorization     = "Bearer $spoToken"
        Accept            = 'application/json;odata.metadata=minimal'
        'X-RequestDigest' = $digest
        'odata-version'   = '4.0'
    }

    $exportResp = Invoke-GraphRequestWithThrottleHandling `
        -Uri "$AdminUrl/_api/SPO.Tenant/ExportToCSV" `
        -Method POST `
        -Headers $postHeaders `
        -Body $exportBody `
        -ContentType 'application/json;charset=utf-8'

    $relPath = $null
    if ($exportResp.d -and $exportResp.d.ExportToCSV) {
        $relPath = $exportResp.d.ExportToCSV
    }
    elseif ($exportResp.value) {
        $relPath = $exportResp.value
    }
    if (-not $relPath) {
        throw 'ExportToCSV did not return a file path.'
    }

    $relPath = $relPath.TrimStart('/')
    $csvUrl = "$AdminUrl/$relPath"
    $downloadHeaders = @{ Authorization = "Bearer $spoToken" }

    $elapsed = 0
    $ready = $false
    while ($elapsed -lt $SPOExportMaxWaitSec) {
        try {
            Invoke-GraphRequestWithThrottleHandling -Uri $csvUrl -Method HEAD -Headers $downloadHeaders | Out-Null
            $ready = $true
            break
        }
        catch {
            $sc = $null
            if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
            if ($sc -eq 404) {
                Start-Sleep -Seconds $SPOExportPollIntervalSec
                $elapsed += $SPOExportPollIntervalSec
            }
            else {
                throw $_
            }
        }
    }

    if (-not $ready) {
        throw "Report file was not available after ${SPOExportMaxWaitSec}s: $csvUrl"
    }

    if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
    $fileName = Split-Path $relPath -Leaf
    $tenantLabel = if ($AdminUrl -match 'https://([^-]+)-admin\.sharepoint\.com') { $matches[1] } else { $AdminUrl -replace 'https?://' }
    $localFile = Join-Path $OutputPath "UnlicensedOneDrive_${tenantLabel}_$fileName"

    Invoke-WebRequest -Uri $csvUrl -Headers $downloadHeaders -OutFile $localFile -UseBasicParsing -ErrorAction Stop
    Write-Host "  SPO report downloaded: $localFile" -ForegroundColor Green
    return $localFile
}

function Convert-SPOCsvToReportAccounts {
    <#
    .SYNOPSIS
        Converts downloaded SPO unlicensed OneDrive CSV rows into the script's
        common account object shape (Active/SoftDeleted/Archived).
    #>
    param (
        [Parameter(Mandatory)] [string]$CsvPath,
        [Parameter(Mandatory)] [string]$AdminUrl
    )

    $accounts = [System.Collections.Generic.List[object]]::new()
    $rows = Import-Csv -Path $CsvPath

    foreach ($row in $rows) {
        $archiveStatusRaw = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('ARCHIVE_STATUS', 'Archive status', 'ArchiveStatus')
        $archiveStatus = if ($archiveStatusRaw) { "$archiveStatusRaw".Trim() } else { '' }

        $upn = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('ACCOUNT_PROVISIONED_FOR', 'Account provisioned for (UPN)', 'UnlicensedOdbProvisionedForUPN', 'Owner email', 'SITE_OWNER_EMAIL')
        $displayName = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('Display name', 'TITLE', 'Title', 'Username')
        $url = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('URL', 'SiteUrl')
        $storageUsedRaw = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('STORAGE_USED', 'Storage used (GB)', 'StorageUsed')
        $unlicensedReason = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('UNLICENSED_REASON', 'Unlicensed due to', 'UnlicensedOdbReason')
        $unlicensedOnRaw = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('UNLICENSED_ON', 'Unlicensed on', 'UnlicensedOdbStartDate')
        $deletionBlockedByRaw = Get-ObjectPropertyValue -InputObject $row -CandidateNames @('DELETION_BLOCK_REASON', 'Deletion blocked by', 'UnlicensedOdbCleanupBlockReason')
        $deletionBlockedBy = if ($deletionBlockedByRaw) { "$deletionBlockedByRaw".Trim() } else { '' }

        if (-not $upn -and $url) {
            $upn = ConvertTo-UPNFromSiteUrl -SiteUrl "$url"
        }
        if (-not $displayName) {
            if ($upn -and $upn.Contains('@')) {
                $displayName = $upn.Split('@')[0]
            }
            elseif ($url) {
                $displayName = ConvertTo-UPNFromSiteUrl -SiteUrl "$url"
            }
            else {
                $displayName = 'Unknown User'
            }
        }

        # If neither identity nor URL exists, ignore this row.
        if (-not $upn -and -not $url) { continue }

        $storageUsedGB = ''
        if ($null -ne $storageUsedRaw -and "$storageUsedRaw".Trim() -ne '') {
            try { $storageUsedGB = [Math]::Round([double]("$storageUsedRaw"), 3) } catch { $storageUsedGB = "$storageUsedRaw" }
        }

        $unlicensedDate = $null
        if ($unlicensedOnRaw) {
            try { $unlicensedDate = [datetime]::Parse("$unlicensedOnRaw") } catch {}
        }

        $userSource = 'Active'
        $accountEnabled = $true
        if ($archiveStatus -and $archiveStatus -notin @('None', 'none', 'reactivating', 'unknownFutureValue')) {
            $userSource = 'Archived'
            $accountEnabled = $false
        }
        elseif ($unlicensedReason -and "$unlicensedReason" -match 'Owner deleted from Entra ID') {
            $userSource = 'SoftDeleted'
            $accountEnabled = $false
        }

        $noteParts = [System.Collections.Generic.List[string]]::new()
        if ($archiveStatus) { $noteParts.Add("archiveStatus: $archiveStatus") | Out-Null }
        if ($unlicensedReason) { $noteParts.Add("Reason: $unlicensedReason") | Out-Null }
        if ($deletionBlockedBy) { $noteParts.Add("DeletionBlockedBy: $deletionBlockedBy") | Out-Null }
        $noteParts.Add('Source=SPO ExportToCSV') | Out-Null
        $noteParts.Add("AdminUrl=$AdminUrl") | Out-Null

        $accounts.Add([PSCustomObject]@{
                UserId            = ''
                UserPrincipalName = if ($upn) { "$upn" } else { '' }
                DisplayName       = if ($displayName) { "$displayName" } else { '' }
                AccountEnabled    = $accountEnabled
                HasAnyLicense     = $false
                UserSource        = $userSource
                UnlicensedDate    = $unlicensedDate
                UnlicensedDueTo   = if ($unlicensedReason) { "$unlicensedReason" } else { 'Unknown from SPO export' }
                DeletionBlockedBy = $deletionBlockedBy
                ArchiveStatus     = "$archiveStatus"
                DriveInfo         = [PSCustomObject]@{
                    Found             = $true
                    DriveId           = ''
                    DriveWebUrl       = if ($url) { "$url" } else { '' }
                    StorageUsedGB     = $storageUsedGB
                    StorageTotalGB    = ''
                    DriveLastModified = ''
                    Note              = ($noteParts -join ' | ')
                }
            })
    }

    return $accounts
}

function Merge-SPODownloadedReports {
    <#
    .SYNOPSIS
        Merges all downloaded SPO report rows from configured admin URLs into
        a single CSV file for a concise tenant-wide multi-geo view.
    #>
    param (
        [Parameter(Mandatory)] [object[]]$MergedRows
    )

    if ($MergedRows.Count -eq 0) { return '' }

    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $mergedPath = Join-Path $OutputFolder "UnlicensedOneDrive_SPO_Merged_$date.csv"
    $MergedRows | Export-Csv -Path $mergedPath -NoTypeInformation -Encoding UTF8
    Write-Host "  SPO merged download report written: $mergedPath" -ForegroundColor Green
    return $mergedPath
}

function Get-ArchivedOneDriveSitesFromSPOExport {
    <#
    .SYNOPSIS
        Fast-path collector for archived OneDrive accounts by downloading the
        SharePoint Admin "Unlicensed OneDrive" CSV per configured admin URL.
    #>
    $fromCsv = [System.Collections.Generic.List[object]]::new()
    $global:spoDownloadedReportFiles.Clear()
    $global:spoDownloadedReportRows.Clear()
    $global:spoMergedDownloadReportPath = ''

    $mergedRawRows = [System.Collections.Generic.List[object]]::new()

    if (-not $SPOAdminUrls -or $SPOAdminUrls.Count -eq 0) {
        Write-Host '  SPO fast path skipped: no $SPOAdminUrls configured.' -ForegroundColor Gray
        return $fromCsv
    }

    Write-Host "`nPhase 2b: Downloading archived candidates from SPO admin export..." -ForegroundColor Cyan
    Write-Host "  Admin URLs configured: $($SPOAdminUrls.Count)" -ForegroundColor Gray

    foreach ($adminUrl in $SPOAdminUrls) {
        try {
            Write-Host "  SPO export: $adminUrl" -ForegroundColor Gray
            $csvPath = Download-UnlicensedOneDriveCsvFromSPO -AdminUrl $adminUrl -OutputPath $OutputFolder
            $global:spoDownloadedReportFiles.Add($csvPath) | Out-Null

            $rawRows = Import-Csv -Path $csvPath
            foreach ($r in $rawRows) {
                $r | Add-Member -NotePropertyName 'SourceAdminUrl' -NotePropertyValue $adminUrl -Force
                $mergedRawRows.Add($r) | Out-Null
            }

            $parsedAll = Convert-SPOCsvToReportAccounts -CsvPath $csvPath -AdminUrl $adminUrl
            foreach ($item in $parsedAll) { $global:spoDownloadedReportRows.Add($item) | Out-Null }

            $parsedArchived = $parsedAll | Where-Object { $_.UserSource -eq 'Archived' }
            foreach ($item in $parsedArchived) { $fromCsv.Add($item) }

            Write-Host "    CSV rows: $($parsedAll.Count) | Archived rows: $($parsedArchived.Count)" -ForegroundColor Gray
        }
        catch {
            Write-Host "    SPO export failed for $adminUrl : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($MergeDownloadedSPOReports -and $mergedRawRows.Count -gt 0) {
        $global:spoMergedDownloadReportPath = Merge-SPODownloadedReports -MergedRows $mergedRawRows
    }

    Write-Host "  SPO export path complete. Archived candidates: $($fromCsv.Count)" -ForegroundColor Green
    return $fromCsv
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
        $rawDaysUntilReadOnly = $null
        $rawDaysUntilArchive = $null
        $daysUntilDeletion = ''
        $urgencyStatus = 'Unknown - No Unlicensed Date'

        if ($unlicensedDate) {
            $readOnlyDate = $unlicensedDate.AddDays($script:ReadOnlyThresholdDays)
            $archiveDate = $unlicensedDate.AddDays($script:ArchiveThresholdDays)
            $daysSinceUnlicensed = ($script:today - $unlicensedDate.Date).Days
            $rawDaysUntilReadOnly = ($readOnlyDate.Date - $script:today).Days
            $rawDaysUntilArchive = ($archiveDate.Date - $script:today).Days

            # Clamp day counters for report readability while preserving raw values
            # for urgency classification.
            if ($rawDaysUntilArchive -lt 0) {
                $daysUntilReadOnly = 'Already Archived'
                $daysUntilArchive = 'Already Archived'
            }
            else {
                $daysUntilReadOnly = [Math]::Max($rawDaysUntilReadOnly, 0)
                $daysUntilArchive = [Math]::Max($rawDaysUntilArchive, 0)
            }

            if ($rawDaysUntilArchive -lt 0) {
                $urgencyStatus = 'ARCHIVED - Past Day 93'
            }
            elseif ($rawDaysUntilArchive -eq 0) {
                $urgencyStatus = 'CRITICAL - Archives TODAY'
            }
            elseif ($rawDaysUntilArchive -le 7) {
                $urgencyStatus = 'CRITICAL - Archives within 7 days'
            }
            elseif ($rawDaysUntilReadOnly -lt 0 -and $rawDaysUntilArchive -gt 7) {
                $urgencyStatus = 'WARNING - Read-Only, Archive pending'
            }
            elseif ($rawDaysUntilReadOnly -eq 0) {
                $urgencyStatus = 'WARNING - Goes Read-Only TODAY'
            }
            elseif ($rawDaysUntilReadOnly -le 7) {
                $urgencyStatus = 'WARNING - Read-Only within 7 days'
            }
            elseif ($rawDaysUntilReadOnly -le 30) {
                $urgencyStatus = 'MONITOR - Read-Only within 30 days'
            }
            else {
                $urgencyStatus = 'OK - More than 30 days remaining'
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
            $daysUntilReadOnly = 'Already Archived'
            $daysUntilArchive = 'Already Archived'
        }

        # Post-archive deletion risk (MC1381110): if PAYG is not enabled,
        # deletion occurs after ArchiveDeletionThresholdDays from unlicensed date.
        if ($acct.UserSource -eq 'Archived') {
            $paygEnabled = $false
            $paygKnown = $false
            if ($global:tenantPayGStatus -and $null -ne $global:tenantPayGStatus.IsEnabled) {
                $paygEnabled = [bool]$global:tenantPayGStatus.IsEnabled
                $paygKnown = $true
            }

            if ($paygKnown) {
            }

            if ($paygEnabled) {
                $daysUntilDeletion = 'NA - PAYG Enabled'
            }
            else {
                if ($unlicensedDate) {
                    $deletionDate = $unlicensedDate.AddDays($script:ArchiveDeletionThresholdDays)
                    $rawDaysUntilDeletion = ($deletionDate.Date - $script:today).Days
                    if ($rawDaysUntilDeletion -lt 0) {
                        $daysUntilDeletion = 'Deletion Overdue'
                    }
                    else {
                        $daysUntilDeletion = $rawDaysUntilDeletion
                    }
                    $urgencyStatus = 'HIGH RISK - Archived, PAYG not enabled'
                }
                else {
                    $daysUntilDeletion = 'Unknown - No Unlicensed Date'
                    $urgencyStatus = 'HIGH RISK - Archived, PAYG unknown date'
                }
            }
        }
        else {
            $daysUntilDeletion = 'NA - Not Archived'
        }

        $driveInfo = $acct.DriveInfo

        # Cost estimation — projected costs for all accounts with known storage.
        # Archived accounts show what they are actively costing; pre-archive accounts
        # show what they will cost if unlicensed status continues to Day 93.
        $projMonthlyStorageCost = $null
        $projReactivationCost = $null

        $storageGB = $null
        if ($null -ne $driveInfo.StorageUsedGB -and $driveInfo.StorageUsedGB -ne '') {
            try { $storageGB = [double]$driveInfo.StorageUsedGB } catch {}
        }

        if ($null -ne $storageGB) {
            $projMonthlyStorageCost = [Math]::Round($storageGB * $script:ArchivalStorageCostPerGBMonth, 4)
            $projReactivationCost = [Math]::Round($storageGB * $script:ReactivationCostPerGB, 2)
        }

        $enriched.Add([PSCustomObject]@{
                UserSource             = $acct.UserSource
                DisplayName            = $acct.DisplayName
                UserPrincipalName      = $acct.UserPrincipalName
                AccountEnabled         = $acct.AccountEnabled
                UnlicensedDueTo        = $acct.UnlicensedDueTo
                UnlicensedDate         = if ($unlicensedDate) { $unlicensedDate.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                DaysSinceUnlicensed    = $daysSinceUnlicensed
                ReadOnlyDate           = if ($readOnlyDate) { $readOnlyDate.ToString('yyyy-MM-dd') } else { '' }
                ArchiveDate            = if ($archiveDate) { $archiveDate.ToString('yyyy-MM-dd') } else { '' }
                DaysUntilReadOnly      = $daysUntilReadOnly
                DaysUntilArchive       = $daysUntilArchive
                DeletionBlockedBy      = if ($acct.PSObject.Properties['DeletionBlockedBy'] -and $acct.DeletionBlockedBy) { $acct.DeletionBlockedBy } else { '' }
                DaysUntilDeletion      = $daysUntilDeletion
                UrgencyStatus          = $urgencyStatus
                StorageUsedGB          = $driveInfo.StorageUsedGB
                StorageTotalGB         = $driveInfo.StorageTotalGB
                ProjMonthlyStorageCost = $projMonthlyStorageCost
                ProjReactivationCost   = $projReactivationCost
                DriveUrl               = $driveInfo.DriveWebUrl
                DriveLastModified      = $driveInfo.DriveLastModified
                Notes                  = $driveInfo.Note
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

    $projectedWithCost = $Records | Where-Object { $null -ne $_.ProjMonthlyStorageCost }
    if ($projectedWithCost) {
        $totalProjMonthly = ($projectedWithCost | Measure-Object -Property ProjMonthlyStorageCost -Sum).Sum
        $totalProjReactivation = ($projectedWithCost | Measure-Object -Property ProjReactivationCost   -Sum).Sum
        Write-Host "`n  Projected Costs (@ Microsoft unlicensed OneDrive pricing):" -ForegroundColor White
        Write-Host "  Note: Archived accounts reflect active ongoing costs. Pre-archive accounts show projected costs if they reach Day 93." -ForegroundColor Gray
        Write-Host ("    Monthly storage cost  (@`$0.05/GB/month) : `${0:N4} USD/month" -f $totalProjMonthly)      -ForegroundColor Yellow
        Write-Host ("    Reactivation cost     (@`$0.60/GB)       : `${0:N2} USD one-time" -f $totalProjReactivation) -ForegroundColor Yellow
        Write-Host ("    Accounts with storage data: {0} of {1} total" -f $projectedWithCost.Count, $Records.Count) -ForegroundColor Gray
    }
    else {
        Write-Host "`n  Projected Costs: No accounts with storage data available for cost estimation." -ForegroundColor Gray
    }
}

function Remove-IntermediateOutputFiles {
    <#
    .SYNOPSIS
        Deletes intermediate downloaded/merged SPO CSV files created during this run,
        keeping only the main output report file.
    #>
    param (
        [Parameter(Mandatory)] [string]$MainReportPath
    )

    $filesToDelete = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($f in $global:spoDownloadedReportFiles) {
        if ($f) { $filesToDelete.Add([string]$f) | Out-Null }
    }

    if ($global:spoMergedDownloadReportPath) {
        $filesToDelete.Add([string]$global:spoMergedDownloadReportPath) | Out-Null
    }

    # Never delete the primary report.
    if ($MainReportPath) {
        $filesToDelete.Remove([string]$MainReportPath) | Out-Null
    }

    $deletedCount = 0
    foreach ($filePath in $filesToDelete) {
        try {
            if (Test-Path -LiteralPath $filePath) {
                Remove-Item -LiteralPath $filePath -Force -ErrorAction Stop
                $deletedCount++
            }
        }
        catch {
            Write-Host "  Cleanup warning: could not delete $filePath : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Host "  Intermediate files removed: $deletedCount" -ForegroundColor Gray
}

#endregion Output Functions

#region Email Functions

function Send-OneDriveAlertEmail {
    <#
    .SYNOPSIS
        Sends an HTML alert email to the admin list ($EmailTo) summarising OneDrive
        accounts that are within the configured notification window before going
        read-only or being archived.

        Called once per notification type after the report is generated.
        Uses the Microsoft Graph API (POST /users/{EmailFrom}/sendMail) with the
        existing bearer token — no SMTP relay required.
        Requires $SendEmailNotifications = $true and Mail.Send (Application) on the app registration.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [ValidateSet('ReadOnly', 'Archive', 'Deletion')] [string]   $NotificationType,
        [Parameter(Mandatory)] [object[]]                            $AffectedAccounts,
        [Parameter(Mandatory)] [int]                                 $ThresholdDays
    )

    if ($AffectedAccounts.Count -eq 0) { return }

    $thresholdLabel = switch ($NotificationType) {
        'ReadOnly' { 'Read-Only' }
        'Archive' { 'Archive' }
        'Deletion' { 'Deletion' }
    }
    $thresholdDay = switch ($NotificationType) {
        'ReadOnly' { $script:ReadOnlyThresholdDays }
        'Archive' { $script:ArchiveThresholdDays }
        'Deletion' { $script:ArchiveDeletionThresholdDays }
    }

    # --- Subject line ---
    $subject = "[OneDrive Alert] $($AffectedAccounts.Count) account(s) approaching $thresholdLabel within $ThresholdDays day(s) — Tenant: $script:tenantId"

    # --- Build HTML table rows, sorted by days remaining (most urgent first) ---
    $sortedAccounts = $AffectedAccounts | Sort-Object {
        switch ($NotificationType) {
            'ReadOnly' { $_.DaysUntilReadOnly }
            'Archive' { $_.DaysUntilArchive }
            'Deletion' { $_.DaysUntilDeletion }
        }
    }

    $tableRows = foreach ($acct in $sortedAccounts) {
        $daysRemaining = switch ($NotificationType) {
            'ReadOnly' { $acct.DaysUntilReadOnly }
            'Archive' { $acct.DaysUntilArchive }
            'Deletion' { $acct.DaysUntilDeletion }
        }
        $targetDate = switch ($NotificationType) {
            'ReadOnly' { $acct.ReadOnlyDate }
            'Archive' { $acct.ArchiveDate }
            'Deletion' {
                $deletionDateText = ''
                if ($acct.UnlicensedDate) {
                    try {
                        $parsedUnlicensedDate = [datetime]::Parse($acct.UnlicensedDate)
                        $deletionDateText = $parsedUnlicensedDate.AddDays($script:ArchiveDeletionThresholdDays).ToString('yyyy-MM-dd')
                    }
                    catch {}
                }
                $deletionDateText
            }
        }
        $storageText = if ($acct.StorageUsedGB -ne '') { "$($acct.StorageUsedGB) GB" } else { 'N/A' }

        # Row colour: red shading ≤ 3 days, amber ≤ 7 days, white otherwise
        $rowColor = if ($daysRemaining -le 3) { '#fde8e8' } elseif ($daysRemaining -le 7) { '#fff3cd' } else { '#ffffff' }

        $safeName = [System.Web.HttpUtility]::HtmlEncode($acct.DisplayName)
        $safeUpn = [System.Web.HttpUtility]::HtmlEncode($acct.UserPrincipalName)
        $safeStatus = [System.Web.HttpUtility]::HtmlEncode($acct.UrgencyStatus)

        "<tr style='background-color:$rowColor;'>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;'>$safeName</td>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;'>$safeUpn</td>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;'>$($acct.UserSource)</td>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;text-align:right;'>$storageText</td>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;'>$targetDate</td>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;text-align:center;'>$daysRemaining</td>
          <td style='padding:5px 10px;border:1px solid #d0d0d0;'>$safeStatus</td>
        </tr>"
    }

    $headerColor = switch ($NotificationType) {
        'ReadOnly' { '#1a5276' }
        'Archive' { '#7b241c' }
        'Deletion' { '#922b21' }
    }
    $alertHeading = switch ($NotificationType) {
        'ReadOnly' { "OneDrive Read-Only Alert — $($AffectedAccounts.Count) account(s) go read-only within $ThresholdDays day(s)" }
        'Archive' { "OneDrive Archive Alert — $($AffectedAccounts.Count) account(s) will be archived within $ThresholdDays day(s)" }
        'Deletion' { "OneDrive Deletion Risk Alert — $($AffectedAccounts.Count) account(s) reach deletion risk window within $ThresholdDays day(s)" }
    }

    $body = @"
<!DOCTYPE html>
<html>
<body style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222;margin:20px;">
  <h2 style="color:$headerColor;margin-bottom:4px;">$alertHeading</h2>
  <p style="margin-top:0;color:#555;font-size:13px;">
    Tenant: <strong>$script:tenantId</strong> &nbsp;|&nbsp;
    Report date: <strong>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</strong> &nbsp;|&nbsp;
    Threshold: Day <strong>$thresholdDay</strong>
  </p>
  <p>
    The accounts below have not been relicensed and are within
    <strong>$ThresholdDays day(s)</strong> of going <strong>$thresholdLabel</strong>.
    Please review and take action (relicense, transfer data, or delete).
  </p>
  <table style="border-collapse:collapse;width:100%;font-size:13px;">
    <thead>
      <tr style="background-color:$headerColor;color:#fff;">
        <th style="padding:6px 10px;border:1px solid #999;text-align:left;">Display Name</th>
        <th style="padding:6px 10px;border:1px solid #999;text-align:left;">UPN</th>
        <th style="padding:6px 10px;border:1px solid #999;text-align:left;">Source</th>
        <th style="padding:6px 10px;border:1px solid #999;text-align:right;">Storage Used</th>
        <th style="padding:6px 10px;border:1px solid #999;text-align:left;">$thresholdLabel Date</th>
        <th style="padding:6px 10px;border:1px solid #999;text-align:center;">Days Left</th>
        <th style="padding:6px 10px;border:1px solid #999;text-align:left;">Urgency</th>
      </tr>
    </thead>
    <tbody>
      $($tableRows -join "`n      ")
    </tbody>
  </table>
  <br/>
  <p style="font-size:12px;color:#888;">
    Full report saved to: $script:outputLog<br/>
    Generated by Get-UnlicensedOneDriveReport.ps1
  </p>
</body>
</html>
"@

    # Build the Graph sendMail payload.
    # toRecipients is constructed from the $EmailTo array — each address becomes
    # a separate emailAddress object so both individual mailboxes and mail-enabled
    # groups are handled correctly.
    $toRecipients = @($script:EmailTo | ForEach-Object {
            @{ emailAddress = @{ address = $_ } }
        })

    $graphMailBody = @{
        message         = @{
            subject      = $subject
            body         = @{ contentType = 'HTML'; content = $body }
            toRecipients = $toRecipients
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 6 -Compress

    # The sender mailbox must match $EmailFrom. With app-only auth the call is
    # POST /users/{sender}/sendMail — /me is not valid for client-credentials tokens.
    $sendUri = "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($script:EmailFrom))/sendMail"

    Test-ValidToken
    $headers = @{ Authorization = "Bearer $global:token" }

    try {
        Invoke-GraphRequestWithThrottleHandling -Uri $sendUri -Method POST -Headers $headers `
            -Body $graphMailBody -ContentType 'application/json'
        Write-Host "  [$thresholdLabel alert] Email sent to: $($script:EmailTo -join ', ')" -ForegroundColor Green
    }
    catch {
        Write-Host "  [$thresholdLabel alert] Email send failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Verify: Mail.Send (Application) is granted for app '$script:clientId' in Entra ID." -ForegroundColor Yellow
        Write-Host "  Sender mailbox '$script:EmailFrom' must be a licensed Exchange Online mailbox." -ForegroundColor Yellow
    }
}

#endregion Email Functions

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
    $rawArchivedSites = [System.Collections.Generic.List[object]]::new()

    switch ($ArchivedCollectionMode) {
        'SPODownload' {
            $rawArchivedSites = Get-ArchivedOneDriveSitesFromSPOExport

            if ($rawArchivedSites.Count -eq 0) {
                Write-Host '  Phase 2b: SPO export returned no archived rows.' -ForegroundColor Gray
            }
        }
        'GraphSites' {
            $rawArchivedSites = Get-ArchivedOneDriveSites
        }
        default {
            Write-Host "  Invalid ArchivedCollectionMode '$ArchivedCollectionMode'. Using 'SPODownload'." -ForegroundColor Yellow
            $rawArchivedSites = Get-ArchivedOneDriveSitesFromSPOExport
        }
    }

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

# Graph JSON batching: POST /$batch with up to 20 requests reduces N serial HTTP
# round-trips to ceil(N/20), a ~20x improvement for large tenants.
# If the batch POST itself fails, the catch block falls back to per-user sequential calls.
$batchSize = 20

for ($batchStart = 0; $batchStart -lt $total; $batchStart += $batchSize) {
    $batchEnd = [Math]::Min($batchStart + $batchSize - 1, $total - 1)
    $batchSlice = $allCandidates[$batchStart..$batchEnd]

    # Build id → user lookup and the requests array for this batch.
    $batchMap = @{}
    $batchReqs = [System.Collections.Generic.List[object]]::new()
    for ($j = 0; $j -lt $batchSlice.Count; $j++) {
        $batchMap["$j"] = $batchSlice[$j]
        $batchReqs.Add(@{ id = "$j"; method = 'GET'; url = "/users/$($batchSlice[$j].UserId)/drive" })
    }

    $batchBody = @{ requests = $batchReqs } | ConvertTo-Json -Depth 3 -Compress

    Test-ValidToken
    $headers = @{ Authorization = "Bearer $global:token" }

    try {
        $batchResp = Invoke-GraphRequestWithThrottleHandling `
            -Uri     'https://graph.microsoft.com/v1.0/$batch' `
            -Method  POST `
            -Headers $headers `
            -Body    $batchBody

        foreach ($resp in $batchResp.responses) {
            $user = $batchMap[$resp.id]
            $current++
            $pct = [Math]::Round(($current / $total) * 100)
            Write-Progress -Activity 'Querying OneDrive' `
                -Status      "$current / $total ($pct%) — $($user.UserPrincipalName)" `
                -PercentComplete $pct

            switch ($resp.status) {
                200 {
                    $drive = $resp.body
                    $usedGB = if ($drive.quota -and $null -ne $drive.quota.used) { [Math]::Round($drive.quota.used / 1GB, 3) } else { 0 }
                    $totalGB = if ($drive.quota -and $null -ne $drive.quota.total) { [Math]::Round($drive.quota.total / 1GB, 3) } else { 0 }
                    if ($debug) { Write-Host "    [OK] $($user.UserPrincipalName) -> $($drive.webUrl)" -ForegroundColor DarkGreen }
                    $user.DriveInfo = [PSCustomObject]@{
                        Found             = $true
                        DriveId           = $drive.id
                        DriveWebUrl       = $drive.webUrl
                        StorageUsedGB     = $usedGB
                        StorageTotalGB    = $totalGB
                        DriveLastModified = $drive.lastModifiedDateTime
                        Note              = ''
                    }
                    $driveFound++
                    $confirmedUnlicensed.Add($user)
                }
                404 {
                    if ($debug) { Write-Host "    [--] $($user.UserPrincipalName) : No OneDrive (404)" -ForegroundColor DarkYellow }
                    $driveNotFound++
                    # Never provisioned or already purged; skip.
                }
                default {
                    $note = switch ($resp.status) {
                        403 { 'Access denied (403) — check Files.Read.All permission' }
                        429 { 'Throttled in batch (429) — re-run or increase $delayBetweenRequests' }
                        default { "HTTP $($resp.status)" }
                    }
                    if ($debug) { Write-Host "    [--] $($user.UserPrincipalName) : $note" -ForegroundColor DarkYellow }
                    $user.DriveInfo = [PSCustomObject]@{
                        Found             = $false
                        DriveId           = ''
                        DriveWebUrl       = ''
                        StorageUsedGB     = ''
                        StorageTotalGB    = ''
                        DriveLastModified = ''
                        Note              = $note
                    }
                    $driveErrors++
                    # 403/timeouts — included in report for admin review.
                    $confirmedUnlicensed.Add($user)
                }
            }
        }
    }
    catch {
        # Batch POST itself failed — fall back to per-user sequential calls for this slice.
        Write-Host "  Batch request failed — falling back to sequential for this slice: $($_.Exception.Message)" -ForegroundColor Yellow
        foreach ($user in $batchSlice) {
            $current++
            $pct = [Math]::Round(($current / $total) * 100)
            Write-Progress -Activity 'Querying OneDrive' `
                -Status      "$current / $total ($pct%) — $($user.UserPrincipalName)" `
                -PercentComplete $pct

            $driveInfo = Get-UserDriveInfo -UserId $user.UserId -UserPrincipalName $user.UserPrincipalName
            $user.DriveInfo = $driveInfo
            if ($driveInfo.Found) {
                $driveFound++
                $confirmedUnlicensed.Add($user)
            }
            elseif ($driveInfo.Note -match '404') {
                $driveNotFound++
            }
            else {
                $driveErrors++
                $confirmedUnlicensed.Add($user)
            }
        }
    }

    if ($delayBetweenRequests -gt 0) { Start-Sleep -Seconds $delayBetweenRequests }
}

Write-Progress -Activity 'Querying OneDrive' -Completed
Write-Host "  Drive queries complete." -ForegroundColor Green
Write-Host "  OneDrive found      : $driveFound" -ForegroundColor Green
Write-Host "  No OneDrive (404)   : $driveNotFound  (skipped — never provisioned or purged)" -ForegroundColor Gray
Write-Host "  Drive errors        : $driveErrors  (403/timeouts — included in report for admin review)" -ForegroundColor Yellow

# Merge archived sites (Phase 2b) into confirmed list.
# DriveInfo is already populated by the collection path (SPO export or Graph).
if ($archivedSites.Count -gt 0) {
    foreach ($s in $archivedSites) { $confirmedUnlicensed.Add($s) }
    Write-Host "  Archived sites merged (Phase 2b): $($archivedSites.Count)" -ForegroundColor Green
}

# Step 6b: Backfill/add accounts from downloaded SPO reports.
# This helps include records that may not be found via audit lookup or candidate enumeration.
if ($IncludeDownloadedRowsInMainReport -and $global:spoDownloadedReportRows.Count -gt 0) {
    $addedFromDownload = 0
    $updatedFromDownload = 0

    foreach ($dl in $global:spoDownloadedReportRows) {
        $match = $null

        if ($dl.UserPrincipalName) {
            $match = $confirmedUnlicensed | Where-Object {
                $_.UserPrincipalName -and $_.UserPrincipalName.Equals($dl.UserPrincipalName, [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1
        }

        if (-not $match -and $dl.DriveInfo -and $dl.DriveInfo.DriveWebUrl) {
            $match = $confirmedUnlicensed | Where-Object {
                $_.DriveInfo -and $_.DriveInfo.DriveWebUrl -and $_.DriveInfo.DriveWebUrl.Equals($dl.DriveInfo.DriveWebUrl, [System.StringComparison]::OrdinalIgnoreCase)
            } | Select-Object -First 1
        }

        if ($match) {
            $changed = $false
            if (-not $match.DisplayName -and $dl.DisplayName) { $match.DisplayName = $dl.DisplayName; $changed = $true }
            if (-not $match.UserPrincipalName -and $dl.UserPrincipalName) { $match.UserPrincipalName = $dl.UserPrincipalName; $changed = $true }
            if (-not $match.UnlicensedDate -and $dl.UnlicensedDate) { $match.UnlicensedDate = $dl.UnlicensedDate; $changed = $true }
            if ((-not $match.UnlicensedDueTo -or $match.UnlicensedDueTo -eq 'Unknown from SPO export') -and $dl.UnlicensedDueTo) { $match.UnlicensedDueTo = $dl.UnlicensedDueTo; $changed = $true }
            if ($dl.DeletionBlockedBy) {
                if ($match.PSObject.Properties['DeletionBlockedBy']) {
                    if (-not $match.DeletionBlockedBy) {
                        $match.DeletionBlockedBy = $dl.DeletionBlockedBy
                        $changed = $true
                    }
                }
                else {
                    $match | Add-Member -NotePropertyName 'DeletionBlockedBy' -NotePropertyValue $dl.DeletionBlockedBy -Force
                    $changed = $true
                }
            }
            if ($match.DriveInfo -and (-not $match.DriveInfo.DriveWebUrl) -and $dl.DriveInfo -and $dl.DriveInfo.DriveWebUrl) {
                $match.DriveInfo.DriveWebUrl = $dl.DriveInfo.DriveWebUrl
                $changed = $true
            }
            if ($match.DriveInfo -and (($match.DriveInfo.StorageUsedGB -eq '' -or $null -eq $match.DriveInfo.StorageUsedGB) -and $dl.DriveInfo -and $dl.DriveInfo.StorageUsedGB -ne '')) {
                $match.DriveInfo.StorageUsedGB = $dl.DriveInfo.StorageUsedGB
                $changed = $true
            }
            if ($changed) { $updatedFromDownload++ }
            continue
        }

        $confirmedUnlicensed.Add($dl)
        $addedFromDownload++
    }

    Write-Host "  SPO download backfill: added $addedFromDownload | updated $updatedFromDownload" -ForegroundColor Green
    if ($global:spoMergedDownloadReportPath) {
        Write-Host "  SPO merged source file: $($global:spoMergedDownloadReportPath)" -ForegroundColor Gray
    }
}

# Determine tenant PAYG status for deletion-risk timeline calculation.
$global:tenantPayGStatus = Get-TenantPayGStatus
Write-Host "  PAYG status: $($global:tenantPayGStatus.Message)" -ForegroundColor Gray
if ($global:tenantPayGStatus -and $global:tenantPayGStatus.DetectionMode) {
    Write-Host "  PAYG detection mode: $($global:tenantPayGStatus.DetectionMode)" -ForegroundColor Gray
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

    # Keep only the main report output from this run.
    Remove-IntermediateOutputFiles -MainReportPath $outputLog
}
else {
    Write-Host "`nNo unlicensed OneDrive accounts with active drives found." -ForegroundColor Green
}

# Step 10: Email notifications
if ($SendEmailNotifications) {
    Write-Host "`nStep 10: Sending email notifications..." -ForegroundColor Cyan

    if ($sorted.Count -eq 0) {
        Write-Host "  No accounts in report — skipping email." -ForegroundColor Gray
    }
    else {
        # Read-Only alert — accounts whose read-only date is within the configured window.
        # Excludes accounts already past read-only (DaysUntilReadOnly < 0) since those
        # are covered by the separate Archive alert.
        $readOnlyAlert = @($sorted | Where-Object {
                $null -ne $_.DaysUntilReadOnly -and
                $_.DaysUntilReadOnly -ge 0 -and
                $_.DaysUntilReadOnly -le $DaysToNotifyBeforeReadOnly
            })

        # Archive alert — accounts whose archive date is within the configured window.
        $archiveAlert = @($sorted | Where-Object {
                $null -ne $_.DaysUntilArchive -and
                $_.DaysUntilArchive -ge 0 -and
                $_.DaysUntilArchive -le $DaysToNotifyBeforeArchive
            })

        # Deletion risk alert — archived accounts approaching deletion-risk threshold.
        # Only numeric day counters are included; text values such as
        # 'NA - PAYG Enabled' or 'Unknown - No Unlicensed Date' are excluded.
        $deletionAlert = @($sorted | Where-Object {
                $daysUntilDeletion = $_.DaysUntilDeletion -as [int]
                $null -ne $daysUntilDeletion -and
                $daysUntilDeletion -ge 0 -and
                $daysUntilDeletion -le $DaysToNotifyBeforeDeletion
            })

        if ($readOnlyAlert.Count -gt 0) {
            Write-Host "  Read-Only alert: $($readOnlyAlert.Count) account(s) within $DaysToNotifyBeforeReadOnly day(s) of read-only." -ForegroundColor Yellow
            Send-OneDriveAlertEmail -NotificationType ReadOnly -AffectedAccounts $readOnlyAlert -ThresholdDays $DaysToNotifyBeforeReadOnly
        }
        else {
            Write-Host "  Read-Only alert: No accounts within $DaysToNotifyBeforeReadOnly day(s) of read-only — email skipped." -ForegroundColor Gray
        }

        if ($archiveAlert.Count -gt 0) {
            Write-Host "  Archive alert  : $($archiveAlert.Count) account(s) within $DaysToNotifyBeforeArchive day(s) of archive." -ForegroundColor Yellow
            Send-OneDriveAlertEmail -NotificationType Archive -AffectedAccounts $archiveAlert -ThresholdDays $DaysToNotifyBeforeArchive
        }
        else {
            Write-Host "  Archive alert  : No accounts within $DaysToNotifyBeforeArchive day(s) of archive — email skipped." -ForegroundColor Gray
        }

        if ($deletionAlert.Count -gt 0) {
            Write-Host "  Deletion alert : $($deletionAlert.Count) account(s) within $DaysToNotifyBeforeDeletion day(s) of deletion risk." -ForegroundColor Yellow
            Send-OneDriveAlertEmail -NotificationType Deletion -AffectedAccounts $deletionAlert -ThresholdDays $DaysToNotifyBeforeDeletion
        }
        else {
            Write-Host "  Deletion alert : No accounts within $DaysToNotifyBeforeDeletion day(s) of deletion risk — email skipped." -ForegroundColor Gray
        }
    }
}
else {
    Write-Host "`nStep 10: Email notifications skipped (`$SendEmailNotifications = `$false)." -ForegroundColor Gray
}

#endregion Main Execution
