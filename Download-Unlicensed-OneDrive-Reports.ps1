<#
.SYNOPSIS
    Downloads the Unlicensed OneDrive Accounts report from the SharePoint Admin Center
    programmatically for one or more geo locations (multi-geo support).

.DESCRIPTION
    Replicates the "Download report" button in the SharePoint Admin Center under
    OneDrive accounts > Unlicensed OneDrive accounts, using the same REST API
    endpoint the admin UI calls:  POST /_api/SPO.Tenant/ExportToCSV

    Flow per geo location:
      1. Acquire a SharePoint Online OAuth 2.0 bearer token (app-only, client credentials).
      2. Obtain a form digest (X-RequestDigest) from /_api/contextinfo.
      3. POST to /_api/SPO.Tenant/ExportToCSV with the CAML view and column mappings
         that filter for unlicensed OneDrive accounts.
      4. Poll the returned server-relative file path until the CSV is ready.
      5. Download and save the CSV locally with the tenant hostname embedded in the
         filename, e.g. UnlicensedOneDrive_contoso_Sites_20260505164854854.csv

    For multi-geo tenants, add one entry per satellite geo to $SPOAdminUrls.
    Each geo requires its own SPO-scoped token (separate OAuth audience).

.PARAMETER (inline configuration — edit the #region Configuration section)
    $tenantId               Azure AD tenant ID of the home tenant.
    $clientId               App registration client ID.
    $AuthType               'Certificate' (default) or 'ClientSecret'.
    $Thumbprint             Certificate thumbprint (Certificate auth only).
    $CertStore              Certificate store: 'LocalMachine' or 'CurrentUser'.
    $clientSecret           Client secret value (ClientSecret auth only).
    $SPOAdminUrls           Array of SharePoint Admin URLs, one per geo location.
    $OutputFolder           Local path to save downloaded CSV files. Defaults to $env:TEMP.
    $MaxRetries             Maximum retry attempts on throttled/transient errors.
    $InitialBackoffSec      Initial back-off delay in seconds before the first retry.
    $RequestTimeoutSec      HTTP request timeout in seconds.
    $SPOExportPollIntervalSec  Seconds to wait between file-readiness polls.
    $SPOExportMaxWaitSec    Maximum seconds to wait for the export file to become available.

.NOTES
    Requirements
    ------------
    - PowerShell 5.1 or later (no external modules required).
    - Azure AD app registration with the following API permission granted and admin-consented:
        SharePoint > Application > Sites.FullControl.All
    - For Certificate auth: the certificate private key must be accessible in the
      specified certificate store on the machine running this script.

    Output
    ------
    One CSV file per geo location, saved to $OutputFolder.
    Filename format: UnlicensedOneDrive_<tenantLabel>_Sites_<timestamp>.csv
    Columns: Display name, Username, Storage used (GB), Unlicensed due to,
             Unlicensed on, Deletion blocked by, Owner email,
             Deletion scheduled on, Archive status,
             Account provisioned for (UPN), URL

             Created by: Mike Lee
Date: 5/5/26

.EXAMPLE
    # Single-geo — run as-is after filling in tenantId, clientId, Thumbprint.
    .\Download-Unlicensed-OneDrive-Reports.ps1

.EXAMPLE
    # Multi-geo — add satellite admin URLs to $SPOAdminUrls in the config section:
    $SPOAdminUrls = @(
        'https://contoso-admin.sharepoint.com'
        'https://contoso-EUR-admin.sharepoint.com'
        'https://contoso-APC-admin.sharepoint.com'
    )
    .\Download-Unlicensed-OneDrive-Reports.ps1
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

# ---- SharePoint Admin URLs (multi-geo: add one entry per geo location) ----
# Format: https://<tenant>-admin.sharepoint.com  (no trailing slashes)
$SPOAdminUrls = @(
    'https://m365cpi13246019-admin.sharepoint.com'
    # 'https://contoso-EUR-admin.sharepoint.com'
    # 'https://contoso-APC-admin.sharepoint.com'
)

# ---- Report output ----
$OutputFolder = $env:TEMP

# ---- Request throttling ----
$MaxRetries = 15
$InitialBackoffSec = 3
$RequestTimeoutSec = 300

# ---- SPO ExportToCSV poll settings ----
$SPOExportPollIntervalSec = 5
$SPOExportMaxWaitSec = 120

##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
# SPO admin token (separate audience from Graph)
$global:spoToken = $null
$global:spoTokenExpiry = $null
#endregion Initialization

#region Helper Functions
# (No helpers required for this script — Graph API is not used.)
#endregion Helper Functions

#region Authentication Functions

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

function AcquireSPOToken {
    <#
    .SYNOPSIS
        Acquires a SharePoint Online admin token (scope: <SPOAdminUrl>/.default).
        Required for the SPO REST API (/_api/SPO.Tenant/*).
        The app registration needs Sites.FullControl.All (or Sites.Selected) in SharePoint.
    #>
    param([Parameter(Mandatory)] [string] $AdminUrl)
    Write-Host "Authenticating to SharePoint Online Admin ($AuthType)..." -ForegroundColor Cyan
    try {
        $result = Get-OAuthClientCredentialToken -Scope "$AdminUrl/.default" -DisplayName 'SPO Admin'
        $global:spoToken = $result.access_token
        $global:spoTokenExpiry = $result.expiry
    }
    catch {
        Write-Host "  SPO Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Test-ValidSPOToken {
    param([Parameter(Mandatory)] [string] $AdminUrl)
    if ($null -eq $global:spoTokenExpiry -or (Get-Date) -gt $global:spoTokenExpiry) {
        Write-Host 'SPO token expired or expiring soon — refreshing...' -ForegroundColor Yellow
        AcquireSPOToken -AdminUrl $AdminUrl
    }
}

#endregion Authentication Functions

#region SPO Admin Report Functions

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

#endregion SPO Admin Report Functions

#region Main Execution

Write-Host '===  OneDrive Report Download  ===' -ForegroundColor Magenta
Write-Host "  Processing $($SPOAdminUrls.Count) admin URL(s)..." -ForegroundColor Cyan

$reportFiles = [System.Collections.Generic.List[string]]::new()

foreach ($adminUrl in $SPOAdminUrls) {
    Write-Host "`n--- $adminUrl ---" -ForegroundColor Magenta
    try {
        # Token is acquired on-demand inside Get-UnlicensedOneDriveReport via Test-ValidSPOToken;
        # pre-acquire here so auth errors surface before any work is attempted.
        AcquireSPOToken -AdminUrl $adminUrl
        $reportFile = Get-UnlicensedOneDriveReport -AdminUrl $adminUrl -OutputPath $OutputFolder
        $reportFiles.Add($reportFile)
    }
    catch {
        Write-Host "  ERROR processing $adminUrl : $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ''
Write-Host "Done. $($reportFiles.Count) report(s) saved:" -ForegroundColor Green
foreach ($f in $reportFiles) { Write-Host "  $f" -ForegroundColor Green }

#endregion Main Execution
