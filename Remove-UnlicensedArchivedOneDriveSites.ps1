<#
.SYNOPSIS
    Safely removes unlicensed archived OneDrive sites identified in the SPO admin report,
    unlocking locked sites first when needed and only deleting sites that are actually deletable.

.DESCRIPTION
    Reads one or more SharePoint Admin Center unlicensed OneDrive CSV exports, filters for
    deletable archived OneDrive candidates, checks each site's LockState, unlocks locked sites
    when configured, and then deletes the site if the lock state is clear and the row is safe to remove.

    This workflow is intentionally defensive:
      1. Read the CSV report.
      2. Skip rows blocked by retention/hold or missing URL data.
      3. Query the current LockState with Get-PnPTenantSite.
      4. If the site is ReadOnly, NoAccess, or NoAdditions, unlock it first.
      5. Confirm the site is now Unlock.
      6. Delete the site only when $PerformDeletes is true.

    For multi-geo tenants, add one entry per satellite geo to $SPOAdminUrls.
    Each geo requires its own SPO-scoped token (separate OAuth audience).

    When $OneDriveSiteList is populated, process only those URLs instead of reading a CSV.
    The administrator is responsible for selecting eligible sites in this mode: report-based
    archived/unlicensed and deletion-block filtering is not available. SharePoint still
    enforces retention/hold restrictions, and $PerformDeletes still controls deletion.

    CSV mode requires URL, archive-status, unlicensed-reason, and deletion-block columns.
    Only Archived, RecentlyArchived, and FullyArchived statuses qualify. Unlicensed reasons
    must be 'Owner deleted from Entra ID' or 'License removed by admin'; other values are
    skipped for review. Only blank, 0, None, Not blocked, or No blockers block values qualify.
    Report downloading is a retained helper, not an automatic step in normal execution.
    Export a current report before running CSV mode. Site processing failures, including
    retention blocks, are audited and summarized with a warning rather than a final throw.
    Run-level failures still cause a terminating error after the final summary.
    CSV URLs are normalized and deduplicated; conflicting duplicate eligibility values
    are rejected. Audit files have unique run IDs and are created before processing.
    The audit is an append-only journal: InProgress rows record mutation checkpoints;
    the last row for each site records its final outcome. Summary counts use final outcomes.
    An audit-write failure terminates processing immediately, before any further mutations.
    InProgress attempt flags record intent, not proof that SharePoint received the operation.

.PARAMETER (inline configuration — edit the #region Configuration section)
    $tenantId               Azure AD tenant ID of the home tenant.
    $clientId               App registration client ID.
    $AuthType               'Certificate' (default) or 'ClientSecret'.
    $Thumbprint             Certificate thumbprint (Certificate auth only).
    $CertStore              Certificate store: 'LocalMachine' or 'CurrentUser'.
    $clientSecret           Client secret value (ClientSecret auth only).
    $SPOAdminUrls           Array of SharePoint Admin URLs, one per geo location.
    $OneDriveSiteList       Optional array of OneDrive URLs; overrides -CsvPath and CSV discovery.
    $OutputFolder           Local path to review downloaded CSV files. Defaults to $env:TEMP.
    $PerformDeletes         Default is $false (dry run). Set to $true to actually delete sites.
    $UnlockLockedSitesBeforeDelete  When $true, unlock locked sites before trying to delete them.
    $MaxRetries             Maximum retry attempts for report REST calls (not PnP operations).
    $InitialBackoffSec      Initial back-off delay in seconds before the first retry.
    $RequestTimeoutSec      HTTP request timeout in seconds.

.NOTES
    Requirements
    ------------
    - PowerShell 7.2 or later.
    - PnP.PowerShell 2.x module.
    - Azure AD app registration with the following API permission granted and admin-consented:
        SharePoint > Application > Sites.FullControl.All
    - For Certificate auth: the certificate private key must be accessible in the
      specified certificate store on the machine running this script.

    Output
    ------
    A CSV audit log is written to $OutputFolder showing each site, its original lock state,
    unlock/delete checkpoints, the Outcome (InProgress, Deleted, DryRun, Skipped, or Failed),
    and the delete result. Missing URLs and per-site failures are included in the audit.
    If a run is interrupted, a site's last InProgress row means its result is unconfirmed;
    check its current SharePoint state before retrying.

    Created by: Mike Lee
Date: 9/8/26

.EXAMPLE
    # Safe dry run: list what would be deleted, but do not delete anything.
    .\Remove-UnlicensedArchivedOneDriveSites.ps1

.EXAMPLE
    # Actual delete mode: set $PerformDeletes = $true in the Configuration section,
    # then run the script normally.
    .\Remove-UnlicensedArchivedOneDriveSites.ps1
#>

param(
    [string]$CsvPath
)

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

# ---- Debug output ----
#region Debug Settings
$debug = $false
#endregion Debug Settings

#region Tenant and App Authentication Settings
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
#endregion Tenant and App Authentication Settings

#region Site Scope and Report Settings
# ---- SharePoint Admin URLs (multi-geo: add one entry per geo location) ----
# Format: https://<tenant>-admin.sharepoint.com  (no trailing slashes)
$SPOAdminUrls = @(
    'https://m365cpi13246019-admin.sharepoint.com'
    # 'https://contoso-EUR-admin.sharepoint.com'
    # 'https://contoso-APC-admin.sharepoint.com'
)

# ---- Report output ----
$OutputFolder = $env:TEMP

# ---- Explicit OneDrive sites (optional; leave empty to use CSV reports) ----
# Only these sites are processed when populated. Include the matching geo in $SPOAdminUrls.
# Report eligibility filtering is bypassed; SharePoint still enforces retention/hold restrictions.
$OneDriveSiteList = @(
    #'https://m365cpi13246019-my.sharepoint.com/personal/robink_m365cpi13246019_onmicrosoft_com'
)
#endregion Site Scope and Report Settings

#region Deletion Safety Settings
# ---- Delete workflow ----
# Safe default: review first and only delete when explicitly enabled.
$PerformDeletes = $false
$UnlockLockedSitesBeforeDelete = $true
#endregion Deletion Safety Settings

#region REST Retry and Report Polling Settings
# ---- Request throttling ----
$MaxRetries = 15
$InitialBackoffSec = 3
$RequestTimeoutSec = 300

# ---- SPO ExportToCSV poll settings ----
$SPOExportPollIntervalSec = 5
$SPOExportMaxWaitSec = 120
#endregion REST Retry and Report Polling Settings

##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
# SharePoint admin token used by PnP.PowerShell and the report download function.
$global:spoToken = $null
$global:spoTokenExpiry = $null
$script:spoTokenCache = @{}
$script:pnpConnectionCache = @{}
#endregion Initialization

#region Authentication Functions

#region OAuth Token Acquisition
function Get-OAuthClientCredentialToken {
    <#
    .SYNOPSIS
        Internal helper — acquires a client-credentials OAuth 2.0 token for any scope.
        Returns a hashtable: @{ access_token = '...'; expiry = [datetime] }
        Throws on failure (callers set the appropriate global variable and handle errors).
    #>
    param(
        [Parameter(Mandatory)] [string] $Scope,
        [Parameter(Mandatory)] [string] $DisplayName   # used only in Write-Host messages
    )

    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    if ($AuthType -eq 'ClientSecret') {
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = $Scope
        }
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
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
        $jwt = "$toSign.$([Convert]::ToBase64String($sig).TrimEnd('=').Replace('+','-').Replace('/','_'))"

        $body = @{
            client_id             = $clientId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwt
            scope                 = $Scope
            grant_type            = 'client_credentials'
        }
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
    }
    else {
        throw "Invalid AuthType '$AuthType'. Use 'Certificate' or 'ClientSecret'."
    }

    $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
    $expiry = (Get-Date).AddSeconds($expiresIn - 300)
    Write-Host "  $DisplayName token acquired ($AuthType). Valid until: $expiry" -ForegroundColor Green
    return @{ access_token = $resp.access_token; expiry = $expiry }
}

#endregion OAuth Token Acquisition

#region SharePoint Token Cache and Refresh
function AcquireSPOToken {
    <#
    .SYNOPSIS
        Acquires a SharePoint Online admin token (scope: <SPOAdminUrl>/.default).
        Used by PnP.PowerShell and the report download function.
    #>
    param([Parameter(Mandatory)] [string] $AdminUrl)
    $cacheKey = $AdminUrl.TrimEnd('/').ToLowerInvariant()
    Write-Host "Authenticating to SharePoint Online Admin ($AuthType)..." -ForegroundColor Cyan
    try {
        $result = Get-OAuthClientCredentialToken -Scope "$AdminUrl/.default" -DisplayName 'SPO Admin'
        $global:spoToken = $result.access_token
        $global:spoTokenExpiry = $result.expiry
        $script:spoTokenCache[$cacheKey] = @{ Token = $result.access_token; Expiry = $result.expiry }
    }
    catch {
        Write-Host "  SPO Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Test-ValidSPOToken {
    param([Parameter(Mandatory)] [string] $AdminUrl)
    $cacheKey = $AdminUrl.TrimEnd('/').ToLowerInvariant()
    $cachedToken = $script:spoTokenCache[$cacheKey]

    if ($cachedToken -and $cachedToken.Token -and $cachedToken.Expiry) {
        $now = Get-Date
        if ($now -lt $cachedToken.Expiry.AddMinutes(-5)) {
            $global:spoToken = $cachedToken.Token
            $global:spoTokenExpiry = $cachedToken.Expiry
            return
        }
    }

    Write-Host 'SPO token expired or expiring soon — refreshing...' -ForegroundColor Yellow
    AcquireSPOToken -AdminUrl $AdminUrl
}

#endregion SharePoint Token Cache and Refresh

#region PnP Admin Connection Cache
function Get-PnPAdminConnection {
    param([Parameter(Mandatory)] [string] $AdminUrl)

    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            throw 'PnP.PowerShell is required for site lock and deletion operations. Install it with Install-Module PnP.PowerShell.'
        }
        Import-Module PnP.PowerShell -ErrorAction Stop
    }

    Test-ValidSPOToken -AdminUrl $AdminUrl
    $cacheKey = $AdminUrl.TrimEnd('/').ToLowerInvariant()
    $cachedConnection = $script:pnpConnectionCache[$cacheKey]
    if ($cachedConnection -and $cachedConnection.Token -eq $global:spoToken) {
        return $cachedConnection.Connection
    }

    $connection = Connect-PnPOnline `
        -Url $AdminUrl `
        -AccessToken $global:spoToken `
        -ReturnConnection `
        -ErrorAction Stop
    $script:pnpConnectionCache[$cacheKey] = @{
        Token      = $global:spoToken
        Connection = $connection
    }
    return $connection
}

#endregion PnP Admin Connection Cache

#endregion Authentication Functions

#region Report REST Support

#region REST Throttling and Retry Handling
function Invoke-SPORequestWithThrottleHandling {
    <#
    .SYNOPSIS
        Wraps Invoke-RestMethod with Retry-After / exponential-backoff throttle handling
        for SharePoint Online REST API calls (429, 502, 503, 504, timeouts).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]    $Uri,
        [Parameter(Mandatory)] [string]    $Method,
        [Parameter()]          [hashtable] $Headers = @{},
        [Parameter()]          [string]    $Body = $null,
        [Parameter()]          [string]    $ContentType = 'application/json;odata=verbose',
        [Parameter()]          [string]    $OutFile = $null,
        [Parameter()]          [int]       $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]       $InitialBackoffSeconds = $script:InitialBackoffSec,
        [Parameter()]          [int]       $TimeoutSeconds = $script:RequestTimeoutSec
    )

    $retryCount = 0
    $backoffSec = $InitialBackoffSeconds

    if ($debug) { Write-Host "  SPO -> $Method $Uri" -ForegroundColor DarkGray }

    while ($true) {
        try {
            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ContentType = $ContentType
                TimeoutSec  = $TimeoutSeconds
                ErrorAction = 'Stop'
                Verbose     = $false
            }
            if ($Body) { $params['Body'] = $Body }
            if ($OutFile) { $params['OutFile'] = $OutFile }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }

            $isRetryable = $statusCode -in @(429, 502, 503, 504) -or
            ($_.Exception -is [System.Net.WebException] -and (
                $_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout -or
                $_.Exception.Status -eq [System.Net.WebExceptionStatus]::ConnectionClosed))

            if (-not $isRetryable) { throw $_ }
            if ($retryCount -ge $MaxRetries) {
                Write-Host "    Max retries reached for: $Uri" -ForegroundColor Red
                throw $_
            }

            $waitSec = $backoffSec
            if ($statusCode -eq 429) {
                try { $ra = $_.Exception.Response.Headers['Retry-After']; if ($ra) { $waitSec = [int]$ra } } catch {}
            }
            $retryCount++
            Write-Host "    SPO throttled ($statusCode). Waiting ${waitSec}s (attempt $retryCount/$MaxRetries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $waitSec
            $backoffSec = [Math]::Min($backoffSec * 2, 300)
        }
    }
}

#endregion REST Throttling and Retry Handling

#region Report Request Form Digest
function Get-SPOFormDigest {
    <#
    .SYNOPSIS
        Retrieves the FormDigestValue required for POST/PUT/DELETE calls to the classic
        SharePoint REST API (_api). App-only OAuth tokens still require this for write ops.
    #>
    param([Parameter(Mandatory)] [string] $AdminUrl)

    $headers = @{
        Authorization = "Bearer $global:spoToken"
        Accept        = 'application/json;odata=verbose'
    }
    $resp = Invoke-SPORequestWithThrottleHandling `
        -Uri     "$AdminUrl/_api/contextinfo" `
        -Method  'POST' `
        -Headers $headers
    return $resp.d.GetContextWebInformation.FormDigestValue
}

#endregion Report Request Form Digest

#endregion Report REST Support

#region PnP Site Operations

#region Read Site Lock State
function Get-SiteLockState {
    <#
    .SYNOPSIS
        Gets the current LockState for a specific OneDrive site with PnP.PowerShell.
    #>
    param(
        [Parameter(Mandatory)] [string] $AdminUrl,
        [Parameter(Mandatory)] [string] $SiteUrl
    )

    $connection = Get-PnPAdminConnection -AdminUrl $AdminUrl
    $site = Get-PnPTenantSite -Identity $SiteUrl -Detailed -Connection $connection -ErrorAction Stop
    if (-not $site) {
        throw "Get-PnPTenantSite returned no site for '$SiteUrl'."
    }
    return [string]$site.LockState
}

#endregion Read Site Lock State

#region Update Site Lock State
function Set-SiteLockState {
    <#
    .SYNOPSIS
        Sets a site's lock state with PnP.PowerShell.
    #>
    param(
        [Parameter(Mandatory)] [string] $AdminUrl,
        [Parameter(Mandatory)] [string] $SiteUrl,
        [Parameter(Mandatory)] [ValidateSet('Unlock','ReadOnly','NoAccess')] [string] $LockState
    )

    $connection = Get-PnPAdminConnection -AdminUrl $AdminUrl
    Set-PnPTenantSite -Identity $SiteUrl -LockState $LockState -Wait -Connection $connection -ErrorAction Stop | Out-Null
}

#endregion Update Site Lock State

#region Move Site to Deleted Sites
function Delete-ArchivedUnlicensedSite {
    <#
    .SYNOPSIS
        Moves the site to Deleted Sites with PnP.PowerShell (recoverable deletion).
    #>
    param(
        [Parameter(Mandatory)] [string] $AdminUrl,
        [Parameter(Mandatory)] [string] $SiteUrl
    )

    $connection = Get-PnPAdminConnection -AdminUrl $AdminUrl
    Remove-PnPTenantSite -Url $SiteUrl -Force -Connection $connection -ErrorAction Stop
}

#endregion Move Site to Deleted Sites

#endregion PnP Site Operations

#region Candidate Processing and Audit Results

#region Shared Report Fields and Geo Routing
function Get-OneDriveReportField {
    param(
        [Parameter(Mandatory)] [object] $Row,
        [Parameter(Mandatory)] [string[]] $Names
    )

    foreach ($name in $Names) {
        if ($Row.PSObject.Properties[$name]) { return [string]$Row.$name }
    }
    return ''
}

function Test-OneDriveReportSchema {
    param([Parameter(Mandatory)] [object] $Row)

    foreach ($aliases in @(
        @('URL', 'SiteUrl'),
        @('ARCHIVE_STATUS', 'ArchiveStatus', 'Archive status'),
        @('UNLICENSED_REASON', 'Unlicensed due to', 'UnlicensedDueTo'),
        @('DELETION_BLOCK_REASON', 'Deletion blocked by', 'DeletionBlockedBy')
    )) {
        if (@($aliases | Where-Object { $Row.PSObject.Properties[$_] }).Count -eq 0) {
            return $false
        }
    }
    return $true
}

function Get-OneDriveAdminUrl {
    param(
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $SiteUrl,
        [Parameter(Mandatory)] [string[]] $AdminUrls
    )

    $siteUri = $null
    if (-not [uri]::TryCreate($SiteUrl, [UriKind]::Absolute, [ref]$siteUri) -or
        $siteUri.Scheme -ne 'https' -or -not $siteUri.IsDefaultPort -or
        $siteUri.UserInfo -or $siteUri.Query -or $siteUri.Fragment -or
        $siteUri.AbsolutePath -notmatch '^/personal/[^/]+/?$') {
        throw "Invalid OneDrive site URL: '$SiteUrl'. Supply an HTTPS /personal/<site> URL."
    }
    $matchingAdmins = @($AdminUrls | Where-Object {
        ([uri]$_).Host -replace '-admin\.', '-my.' -eq $siteUri.Host
    } | Select-Object -Unique)
    if ($matchingAdmins.Count -ne 1) {
        throw "OneDrive URL '$SiteUrl' must match exactly one configured SPOAdminUrls host."
    }
    return $matchingAdmins[0]
}
#endregion Shared Report Fields and Geo Routing

#region Load CSV Candidates
function Write-OneDriveAuditCheckpoint {
    param(
        [Parameter(Mandatory)] [object] $Audit,
        [Parameter(Mandatory)] [string] $AuditPath
    )

    try {
        $Audit | Export-Csv -LiteralPath $AuditPath -Append -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        $script:auditWriteFailed = $true
        throw "Audit write failed; stopping further site operations. Path '$AuditPath': $($_.Exception.Message)"
    }
}

function Process-UnlicensedOneDriveCsv {
    <#
    .SYNOPSIS
        Reads a report CSV and processes each site for safe deletion.
    #>
    param(
        [Parameter(Mandatory)] [string] $CsvPath,
        [Parameter(Mandatory)] [string[]] $AdminUrls,
        [Parameter()] [bool] $AllowDelete = $false,
        [Parameter()] [string] $AuditPath
    )

    $rows = @(Import-Csv -Path $CsvPath -ErrorAction Stop)
    if ($rows.Count -eq 0) { throw "Report '$CsvPath' contains no data rows." }
    if (-not (Test-OneDriveReportSchema -Row $rows[0])) {
        throw "Report '$CsvPath' is missing required URL, archive-status, unlicensed-reason, or deletion-block columns."
    }
    $normalizedRows = foreach ($row in $rows) {
        $url = (Get-OneDriveReportField -Row $row -Names @('URL', 'SiteUrl')).Trim()
        $uri = $null
        if ([uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -and
            $uri.AbsolutePath -match '^/personal/[^/]+/?$') {
            $url = $uri.AbsoluteUri.TrimEnd('/')
        }
        [PSCustomObject]@{
            URL = $url
            ARCHIVE_STATUS = (Get-OneDriveReportField -Row $row -Names @('ARCHIVE_STATUS', 'ArchiveStatus', 'Archive status')).Trim()
            UNLICENSED_REASON = (Get-OneDriveReportField -Row $row -Names @('UNLICENSED_REASON', 'Unlicensed due to', 'UnlicensedDueTo')).Trim()
            DELETION_BLOCK_REASON = (Get-OneDriveReportField -Row $row -Names @('DELETION_BLOCK_REASON', 'Deletion blocked by', 'DeletionBlockedBy')).Trim()
            DuplicateConflict = $false
        }
    }
    $uniqueRows = @(
        foreach ($group in ($normalizedRows | Group-Object URL)) {
            $candidate = $group.Group[0]
            foreach ($field in @('ARCHIVE_STATUS', 'UNLICENSED_REASON', 'DELETION_BLOCK_REASON')) {
                if (@($group.Group.$field | Sort-Object -Unique).Count -gt 1) {
                    $candidate.DuplicateConflict = $true
                }
            }
            $candidate
        }
    )
    Write-Host "  Report contains $($rows.Count) row(s), $($uniqueRows.Count) unique URL(s)." -ForegroundColor Cyan
    Process-OneDriveSites -Rows $uniqueRows -AdminUrls $AdminUrls -Source $CsvPath -AllowDelete $AllowDelete -AuditPath $AuditPath
}

#endregion Load CSV Candidates

#region Shared CSV and Explicit-List Workflow
function Process-OneDriveSites {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]] $Rows,
        [Parameter(Mandatory)] [string[]] $AdminUrls,
        [Parameter(Mandatory)] [string] $Source,
        [Parameter()] [bool] $AllowDelete = $false,
        [Parameter()] [switch] $ExplicitSiteList,
        [Parameter()] [string] $AuditPath
    )

    if ($AllowDelete -and -not $AuditPath) {
        throw 'An initialized audit path is required before allowing site mutations.'
    }

    foreach ($row in $Rows) {
        $audit = [PSCustomObject]@{
            SiteUrl = ''
            AdminUrl = ''
            LockStateBefore = ''
            LockStateAfter = ''
            UnlockAttempted = $false
            DeleteAttempted = $false
            DeleteSucceeded = $false
            Outcome = 'Skipped'
            Reason = ''
            CSV = $Source
        }
        try {
            #region Read Candidate Fields
            if (-not $ExplicitSiteList -and -not (Test-OneDriveReportSchema -Row $row)) {
                throw 'Row is missing required report columns.'
            }
            $siteUrl = (Get-OneDriveReportField -Row $row -Names @('URL', 'SiteUrl')).Trim()
            $audit.SiteUrl = $siteUrl
            if ($row.PSObject.Properties['DuplicateConflict'] -and $row.DuplicateConflict) {
                throw "Conflicting eligibility values in duplicate CSV rows for '$siteUrl'. No operation attempted."
            }
            $adminUrl = Get-OneDriveAdminUrl -SiteUrl $siteUrl -AdminUrls $AdminUrls
            $audit.AdminUrl = $adminUrl
            $archiveStatus = (Get-OneDriveReportField -Row $row -Names @('ARCHIVE_STATUS', 'ArchiveStatus', 'Archive status')).Trim()
            $deleteBlockedBy = (Get-OneDriveReportField -Row $row -Names @('DELETION_BLOCK_REASON', 'Deletion blocked by', 'DeletionBlockedBy')).Trim()
            $unlicensedReason = (Get-OneDriveReportField -Row $row -Names @('UNLICENSED_REASON', 'Unlicensed due to', 'UnlicensedDueTo')).Trim()
            #endregion Read Candidate Fields

            #region Report Eligibility and Retention Filtering
            $isArchiveCandidate = $archiveStatus -in @('Archived', 'RecentlyArchived', 'FullyArchived')
            if (-not $ExplicitSiteList -and -not $isArchiveCandidate) {
                $audit.Reason = "Skipped: archive status '$archiveStatus' is not eligible"
                Write-Host "    $($audit.Reason): $siteUrl" -ForegroundColor Yellow
                continue
            }
            if (-not $ExplicitSiteList -and $unlicensedReason -notin @('Owner deleted from Entra ID', 'License removed by admin')) {
                $audit.Reason = "Skipped: unlicensed reason '$unlicensedReason' is not recognized as eligible"
                Write-Host "    $($audit.Reason): $siteUrl" -ForegroundColor Yellow
                continue
            }
            if (-not $ExplicitSiteList -and $deleteBlockedBy -notin @('', '0', 'None', 'Not blocked', 'No blockers')) {
                $audit.Reason = "Skipped: deletion blocked by '$deleteBlockedBy'"
                Write-Host "    $($audit.Reason): $siteUrl" -ForegroundColor Yellow
                continue
            }

            #endregion Report Eligibility and Retention Filtering

            #region Check and Unlock the Site
            Write-Host "    Candidate found: $siteUrl" -ForegroundColor Cyan
            $lockStateBefore = Get-SiteLockState -AdminUrl $AdminUrl -SiteUrl $siteUrl
            $audit.LockStateBefore = $lockStateBefore
            $audit.LockStateAfter = $lockStateBefore
            if ($lockStateBefore -notin @('Unlock', 'ReadOnly', 'NoAccess', 'NoAdditions')) {
                throw "Could not confirm a valid lock state for '$siteUrl': '$lockStateBefore'."
            }
            Write-Host "      LockState before: $lockStateBefore" -ForegroundColor DarkCyan
            if (-not $AllowDelete) {
                if ($lockStateBefore -eq 'Unlock') {
                    $audit.Outcome = 'DryRun'
                    $audit.Reason = 'Dry run: would attempt deletion; SharePoint enforces retention/hold restrictions'
                }
                elseif ($UnlockLockedSitesBeforeDelete) {
                    $audit.Outcome = 'DryRun'
                    $audit.Reason = 'Dry run: would unlock, verify Unlock, then attempt deletion; SharePoint enforces retention/hold restrictions'
                }
                else {
                    $audit.Reason = "Skipped: lock state is '$lockStateBefore' and unlocking is disabled"
                }
                Write-Host "      $($audit.Reason)" -ForegroundColor Yellow
                continue
            }
            if ($UnlockLockedSitesBeforeDelete -and $lockStateBefore -in @('ReadOnly', 'NoAccess', 'NoAdditions')) {
                $audit.UnlockAttempted = $true
                if ($AuditPath) {
                    $audit.Outcome = 'InProgress'
                    $audit.Reason = 'About to attempt unlock; completion not yet confirmed'
                    Write-OneDriveAuditCheckpoint -Audit $audit -AuditPath $AuditPath
                }
                Write-Host "      Unlocking site before delete: $siteUrl (current lock: $lockStateBefore)" -ForegroundColor Yellow
                Set-SiteLockState -AdminUrl $AdminUrl -SiteUrl $siteUrl -LockState 'Unlock' | Out-Null
                Start-Sleep -Seconds 5
            }

            $lockStateAfter = Get-SiteLockState -AdminUrl $AdminUrl -SiteUrl $siteUrl
            $audit.LockStateAfter = $lockStateAfter
            if ($lockStateAfter -notin @('Unlock', 'ReadOnly', 'NoAccess', 'NoAdditions')) {
                throw "Could not confirm a valid lock state after unlock for '$siteUrl': '$lockStateAfter'."
            }
            Write-Host "      LockState after: $lockStateAfter" -ForegroundColor DarkCyan
            #endregion Check and Unlock the Site

            #region Delete or Record Dry-Run Outcome
            if ($lockStateAfter -ne 'Unlock') {
                if ($audit.UnlockAttempted) { throw "Lock state remains '$lockStateAfter' after unlocking '$siteUrl'." }
                $audit.Reason = "Skipped: lock state remains '$lockStateAfter'"
                Write-Host "      Site not deletable because its lock state is '$lockStateAfter'." -ForegroundColor Yellow
            }
            else {
                $audit.DeleteAttempted = $true
                if ($AuditPath) {
                    $audit.Outcome = 'InProgress'
                    $audit.Reason = 'About to attempt deletion; completion not yet confirmed'
                    Write-OneDriveAuditCheckpoint -Audit $audit -AuditPath $AuditPath
                }
                Delete-ArchivedUnlicensedSite -AdminUrl $AdminUrl -SiteUrl $siteUrl | Out-Null
                $audit.DeleteSucceeded = $true
                $audit.Outcome = 'Deleted'
                $audit.Reason = 'Moved to Deleted Sites'
                Write-Host "      Moved site to Deleted Sites: $siteUrl" -ForegroundColor Red
            }

            #endregion Delete or Record Dry-Run Outcome

        }
        #region Record Site Audit Result
        catch {
            if ($script:auditWriteFailed) { throw }
            $audit.Outcome = 'Failed'
            $audit.Reason = "Processing failed: $($_.Exception.Message)"
            Write-Host "    ERROR processing '$($audit.SiteUrl)': $($_.Exception.Message)" -ForegroundColor Red
        }
        finally {
            if (-not $script:auditWriteFailed) {
                if ($AuditPath) { Write-OneDriveAuditCheckpoint -Audit $audit -AuditPath $AuditPath }
                $audit
            }
        }
        #endregion Record Site Audit Result
    }

}

#endregion Shared CSV and Explicit-List Workflow

#endregion Candidate Processing and Audit Results

#region OneDrive Report Export and Download

function Get-UnlicensedOneDriveReport {
    <#
    .SYNOPSIS
        Downloads the "Unlicensed OneDrive accounts" CSV report from the SharePoint
        Admin Center using the same /_api/SPO.Tenant/ExportToCSV endpoint that the
        admin UI calls when you click "Download report".

    .DESCRIPTION
        Flow (mirrors the Fiddler trace):
          1. POST  /_api/SPO.Tenant/ExportToCSV   — triggers report generation
          2. Parse the response for the server-relative file path
          3. GET   /<library>/<filename>.csv        — downloads the generated file

        The generated file lands in the DO_NOT_DELETE_DOCLIB_ACTIVE_SITES_REPORT
        document library on the admin site, named Sites_<timestamp>.csv.

    .PARAMETER AdminUrl
        SharePoint Admin URL, e.g. https://contoso-admin.sharepoint.com

    .PARAMETER OutputPath
        Local folder to save the downloaded CSV. Defaults to $OutputFolder.

    .OUTPUTS
        Full path of the saved CSV file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $AdminUrl,
        [Parameter()]          [string] $OutputPath = $OutputFolder
    )

    # --- Ensure valid SPO token ---
    Test-ValidSPOToken -AdminUrl $AdminUrl

    # --- Get form digest for the POST ---
    Write-Host 'Getting SPO form digest...' -ForegroundColor Cyan
    $digest = Get-SPOFormDigest -AdminUrl $AdminUrl

    # --- POST to ExportToCSV to trigger report generation ---
    Write-Host 'Requesting Unlicensed OneDrive accounts report export...' -ForegroundColor Cyan

    $postHeaders = @{
        Authorization     = "Bearer $global:spoToken"
        Accept            = 'application/json;odata.metadata=minimal'
        'X-RequestDigest' = $digest
        'odata-version'   = '4.0'
    }

    # Full body extracted from HAR — the endpoint requires all three parameters:
    #   viewXml    : CAML query filtering unlicensed OneDrive accounts only
    #   columnsInfo: column-name-to-field mappings for the CSV header row
    #   listName   : the internal tenant-admin aggregated sites list
    # Without listName the server throws ArgumentNullException (Parameter name: s).
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

    $exportResp = Invoke-SPORequestWithThrottleHandling `
        -Uri         "$AdminUrl/_api/SPO.Tenant/ExportToCSV" `
        -Method      'POST' `
        -Headers     $postHeaders `
        -Body        $exportBody `
        -ContentType 'application/json;charset=utf-8'

    # --- Parse the server-relative path returned by ExportToCSV ---
    # OData 4.0 minimal: { "@odata.context": "...", "value": "DO_NOT_DELETE_.../Sites_<ts>.csv" }
    # OData verbose fallback: { d: { ExportToCSV: '...' } }
    $relPath = $null
    if ($exportResp.d -and $exportResp.d.ExportToCSV) {
        $relPath = $exportResp.d.ExportToCSV
    }
    elseif ($exportResp.value) {
        $relPath = $exportResp.value
    }

    if (-not $relPath) {
        Write-Host "  Unexpected ExportToCSV response. Raw:" -ForegroundColor Red
        Write-Host ($exportResp | ConvertTo-Json -Depth 5) -ForegroundColor Gray
        throw 'ExportToCSV did not return a file path.'
    }

    # Strip leading slash if present so Join-Path / string concat works cleanly
    $relPath = $relPath.TrimStart('/')
    $csvUrl = "$AdminUrl/$relPath"
    Write-Host "  Report file path: $relPath" -ForegroundColor Gray

    # --- Poll until the file is ready (the export may take a few seconds) ---
    $downloadHeaders = @{ Authorization = "Bearer $global:spoToken" }
    $elapsed = 0
    $ready = $false

    Write-Host "  Waiting for file to be ready..." -ForegroundColor Cyan
    while ($elapsed -lt $SPOExportMaxWaitSec) {
        try {
            # HEAD request to check existence without downloading the full file
            Invoke-SPORequestWithThrottleHandling `
                -Uri     $csvUrl `
                -Method  'HEAD' `
                -Headers $downloadHeaders | Out-Null
            $ready = $true
            break
        }
        catch {
            $sc = $null
            if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
            if ($sc -eq 404) {
                Write-Host "    File not ready yet — waiting ${SPOExportPollIntervalSec}s..." -ForegroundColor Yellow
                Start-Sleep -Seconds $SPOExportPollIntervalSec
                $elapsed += $SPOExportPollIntervalSec
            }
            else { throw $_ }
        }
    }

    if (-not $ready) {
        throw "Report file was not available after ${SPOExportMaxWaitSec}s: $csvUrl"
    }

    # --- Download the CSV ---
    if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

    $fileName = Split-Path $relPath -Leaf
    # Include the tenant hostname so files from different geo locations don't collide.
    # e.g. UnlicensedOneDrive_m365cpi13246019_Sites_20260505164854854.csv
    $tenantLabel = if ($AdminUrl -match 'https://([^-]+)-admin\.sharepoint\.com') { $matches[1] } else { $AdminUrl -replace 'https?://' }
    $localFile = Join-Path $OutputPath "UnlicensedOneDrive_${tenantLabel}_$fileName"

    Write-Host "  Downloading CSV to: $localFile" -ForegroundColor Cyan
    Invoke-SPORequestWithThrottleHandling `
        -Uri     $csvUrl `
        -Method  'GET' `
        -Headers $downloadHeaders `
        -OutFile $localFile

    Write-Host "  Report saved: $localFile" -ForegroundColor Green
    return $localFile
}

#endregion OneDrive Report Export and Download

#region Main Execution

#region Startup and Execution Mode
Write-Host '===  Remove Unlicensed Archived OneDrive Sites  ===' -ForegroundColor Magenta
Write-Host "  Processing $($SPOAdminUrls.Count) admin URL(s)..." -ForegroundColor Cyan

if (-not $PerformDeletes) {
    Write-Host '  Dry-run mode enabled. Set $PerformDeletes = $true in the Configuration section to delete sites.' -ForegroundColor Yellow
}

#endregion Startup and Execution Mode

#region Select Explicit Sites or CSV Input
$csvFiles = @()
$explicitSites = @()
if ($OneDriveSiteList.Count -gt 0) {
    $explicitSites = @(
        foreach ($siteUrl in $OneDriveSiteList) {
            $adminUrl = Get-OneDriveAdminUrl -SiteUrl $siteUrl -AdminUrls $SPOAdminUrls
            $siteUri = [uri]$siteUrl
            [PSCustomObject]@{
                URL = $siteUri.AbsoluteUri.TrimEnd('/')
                AdminUrl = $adminUrl
            }
        }
    ) | Sort-Object URL -Unique
    $explicitSites = @($explicitSites)
    Write-Host "  Explicit site list: $($explicitSites.Count) unique site(s). CSV input/discovery is bypassed." -ForegroundColor Yellow
    Write-Host '  Report eligibility filtering is bypassed. SharePoint still enforces retention/hold restrictions.' -ForegroundColor Yellow
}
elseif ($CsvPath) {
    $csvFiles = @($CsvPath)
}
else {
    $candidates = @(Get-ChildItem -Path $OutputFolder -Filter 'UnlicensedOneDrive*.csv' -File -ErrorAction Stop |
        Where-Object Name -NotLike 'UnlicensedOneDrive_DeleteAudit_*' |
        Sort-Object LastWriteTime -Descending)
    foreach ($candidateFile in $candidates) {
        try {
            $sampleRow = Import-Csv -LiteralPath $candidateFile.FullName -ErrorAction Stop | Select-Object -First 1
            if ($sampleRow -and (Test-OneDriveReportSchema -Row $sampleRow)) {
                $csvFiles = @($candidateFile.FullName)
                break
            }
            Write-Warning "Skipping empty or incompatible report '$($candidateFile.FullName)'."
        }
        catch {
            Write-Warning "Skipping unreadable report '$($candidateFile.FullName)': $($_.Exception.Message)"
        }
    }
}

if ($explicitSites.Count -eq 0 -and $csvFiles.Count -eq 0) {
    Write-Host "  No compatible Unlicensed OneDrive report found in $OutputFolder. Download a report first or pass -CsvPath <report.csv>." -ForegroundColor Red
    throw 'No input sites or compatible report available.'
}

#endregion Select Explicit Sites or CSV Input

#region Initialize Durable Audit Journal
$script:auditWriteFailed = $false
$runId = [guid]::NewGuid().ToString('N')
$auditPath = Join-Path $OutputFolder "UnlicensedOneDrive_DeleteAudit_$(Get-Date -Format 'yyyyMMddHHmmss')_$runId.csv"
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -ErrorAction Stop | Out-Null
}
$auditTemplate = [PSCustomObject]@{
    SiteUrl = ''
    AdminUrl = ''
    LockStateBefore = ''
    LockStateAfter = ''
    UnlockAttempted = $false
    DeleteAttempted = $false
    DeleteSucceeded = $false
    Outcome = ''
    Reason = ''
    CSV = ''
}
$auditHeader = ($auditTemplate | ConvertTo-Csv -NoTypeInformation)[0]
$auditHeader | Out-File -LiteralPath $auditPath -NoClobber -Encoding utf8 -ErrorAction Stop
Write-Host "  Audit journal initialized: $auditPath" -ForegroundColor Cyan
#endregion Initialize Durable Audit Journal

#region Process Explicit OneDrive Site List
$allAuditRows = [System.Collections.Generic.List[object]]::new()
$runFailures = 0
if ($explicitSites.Count -gt 0) {
    foreach ($adminUrl in ($explicitSites.AdminUrl | Select-Object -Unique)) {
        Write-Host "`n--- $adminUrl ---" -ForegroundColor Magenta
        foreach ($site in ($explicitSites | Where-Object AdminUrl -eq $adminUrl)) {
            try {
                Process-OneDriveSites -Rows @($site) -AdminUrls @($adminUrl) `
                    -Source 'Configuration: OneDriveSiteList' -AllowDelete $PerformDeletes -ExplicitSiteList -AuditPath $auditPath |
                    ForEach-Object { $allAuditRows.Add($_) }
            }
            catch {
                if ($script:auditWriteFailed) { throw }
                $message = $_.Exception.Message
                $runFailures++
                Write-Host "  ERROR processing $($site.URL): $message" -ForegroundColor Red
                $failureAudit = [PSCustomObject]@{
                    SiteUrl = $site.URL
                    AdminUrl = $adminUrl
                    LockStateBefore = ''
                    LockStateAfter = ''
                    UnlockAttempted = $false
                    DeleteAttempted = $false
                    DeleteSucceeded = $false
                    Outcome = 'Failed'
                    Reason = "Processing failed: $message"
                    CSV = 'Configuration: OneDriveSiteList'
                }
                Write-OneDriveAuditCheckpoint -Audit $failureAudit -AuditPath $auditPath
                $allAuditRows.Add($failureAudit)
            }
        }
    }
}
#endregion Process Explicit OneDrive Site List

#region Process CSV Reports
foreach ($csvFile in $csvFiles) {
    Write-Host "`nUsing report: $csvFile" -ForegroundColor Cyan
    Write-Host '  Looking for archived, unlicensed OneDrive sites that are deletable...' -ForegroundColor Cyan
    try {
        Process-UnlicensedOneDriveCsv -CsvPath $csvFile -AdminUrls $SPOAdminUrls -AllowDelete $PerformDeletes -AuditPath $auditPath |
            ForEach-Object { $allAuditRows.Add($_) }
    }
    catch {
        if ($script:auditWriteFailed) { throw }
        $runFailures++
        Write-Host "  ERROR processing report '$csvFile': $($_.Exception.Message)" -ForegroundColor Red
    }
}

#endregion Process CSV Reports

#region Export Audit Log and Display Summary
Write-Host "`nAudit journal saved: $auditPath" -ForegroundColor Green

$failedSites = @($allAuditRows | Where-Object Outcome -eq 'Failed').Count
$deletedSites = @($allAuditRows | Where-Object Outcome -eq 'Deleted').Count
$dryRunSites = @($allAuditRows | Where-Object Outcome -eq 'DryRun').Count
$skippedSites = @($allAuditRows | Where-Object Outcome -eq 'Skipped').Count
$summaryColor = if ($failedSites -gt 0 -or $runFailures -gt 0) { 'Red' } else { 'Green' }
Write-Host "`nComplete. $($allAuditRows.Count) row(s) reviewed: $deletedSites deleted, $dryRunSites dry-run, $skippedSites skipped, $failedSites failed; $runFailures run-level error(s)." -ForegroundColor $summaryColor
if (-not $PerformDeletes) {
    Write-Host '  Set $PerformDeletes = $true in the Configuration section to perform the unlock/delete actions shown in the audit log.' -ForegroundColor Yellow
}
if ($runFailures -gt 0) {
    throw "OneDrive processing completed with $failedSites failed site(s) and $runFailures run-level error(s). Review the audit log and errors above."
}
if ($failedSites -gt 0) {
    Write-Warning "OneDrive processing completed with $failedSites failed site(s). These failures were handled and audited. Review '$auditPath' and the site errors above before retrying."
}

#endregion Export Audit Log and Display Summary

#endregion Main Execution
