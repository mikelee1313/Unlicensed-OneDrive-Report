<#
.SYNOPSIS
    Offboarding handoff workflow for OneDrive sites.

.DESCRIPTION
    For each user in an input CSV, this script:
    1) Resolves the user's manager from Entra ID manager field.
    2) Resolves the user's OneDrive URL.
    3) Grants the manager Site Collection Admin (SCA) rights.
    4) Sets the OneDrive site LockState to ReadOnly.
    5) Locates sharing links by enumerating SharingLinks groups and matching link IDs.
    6) Emails the manager with ownership details and sharing-link inventory.

    The notification includes a note that links should be considered expired
    90 days after account deletion.

.INPUT FILE FORMAT (CSV)
    Required column:
      - UserPrincipalName

    Optional columns:
      - DeletedDate   (yyyy-MM-dd or any parseable date)
      - OneDriveUrl   (if known; otherwise script resolves via Graph)

.NOTES
    Required modules:
      - PnP.PowerShell
      - Microsoft Graph API access via app registration

    Required app permissions:
      Microsoft Graph (Application):
        - User.Read.All         — user lookup and manager resolution
        - Directory.Read.All    — query soft-deleted users (Entra recycle bin)
        - Files.Read.All        — resolve OneDrive drive URL via /users/{id}/drive
        - AuditLog.Read.All     — license-removal date lookup via directoryAudits
        - Mail.Send             — send manager notification email via Graph sendMail

      SharePoint (Application):
        - Sites.FullControl.All — grant Site Collection Admin, set LockState,
                                  read site groups, members, and sharing links
    
    Author: Mike Lee
    Date: 5/28/2026

.DISCLAIMER
    The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
#>

#region Configuration
$tenantName = 'm365cpi13246019'
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId = 'abc64618-283f-47ba-a185-50d935d51d57'
$thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9'

# Mailbox used to send notifications with Graph app-only sendMail.
$emailFrom = 'admin@M365CPI13246019.onmicrosoft.com'

# Expiration policy for offboarded user's existing sharing links.
$linkExpiryDaysAfterDeletion = 90

# Set to $true to lock the OneDrive site ReadOnly after granting SCA.
# Set to $false to grant SCA only, leaving the site accessible.
$lockSite = $false

# Manager escalation level for offboarding workflow.
# 1 = user's direct manager, 2 = manager's manager, 3 = manager's manager's manager
$managerLevel = 3

# Whether to grant SCA and send notification to all manager levels (true)
# or only the first level manager (false).
$givePermsToAllManagers = $false

# How far back to search audit logs for license-change events when no DeletedDate
# is supplied and the user is not in the Entra soft-delete bin.
# Requires AuditLog.Read.All on the app registration. Set to 0 to skip audit lookup.
$auditLogLookbackDays = 180

# Input / output
$inputCsvPath = 'C:\Temp\OneDriveOffboardingInput.csv'
$outputFolder = $env:TEMP
$runStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outputCsvPath = Join-Path $outputFolder "OneDrive_Manager_Handoff_$runStamp.csv"

# Basic logging
$debug = $false
#endregion Configuration

#region Globals
$global:graphToken = $null
$global:graphTokenExpiry = $null
#endregion Globals

#region Helper Functions

#region Logging
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    if ($Level -eq 'DEBUG' -and -not $debug) {
        return
    }

    $color = switch ($Level) {
        'INFO' { 'Cyan' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        'DEBUG' { 'DarkGray' }
    }

    Write-Host "[$Level] $Message" -ForegroundColor $color
}
#endregion Logging

#region Graph Authentication
function Get-GraphToken {
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $scope = 'https://graph.microsoft.com/.default'

    try {
        $cert = Get-Item -Path "Cert:\LocalMachine\My\$thumbprint" -ErrorAction Stop
    }
    catch {
        throw "Certificate with thumbprint '$thumbprint' was not found in LocalMachine\\My."
    }

    $now = [System.DateTimeOffset]::UtcNow
    $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
    $nbf = $now.ToUnixTimeSeconds()

    $x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress
    $payload = @{
        aud = $tokenUri
        exp = $exp
        iss = $clientId
        jti = [Guid]::NewGuid().ToString()
        nbf = $nbf
        sub = $clientId
    } | ConvertTo-Json -Compress

    $headerB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $payloadB64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $toSign = "$headerB64.$payloadB64"

    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if (-not $rsa) {
        throw "Unable to access private key for certificate '$thumbprint'."
    }

    $signature = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($toSign),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $sigB64 = [Convert]::ToBase64String($signature).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $clientAssertion = "$toSign.$sigB64"

    $body = @{
        client_id             = $clientId
        scope                 = $scope
        grant_type            = 'client_credentials'
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        client_assertion      = $clientAssertion
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        $global:graphToken = $resp.access_token
        $global:graphTokenExpiry = (Get-Date).AddSeconds([int]$resp.expires_in - 300)
    }
    catch {
        $rawDetail = ''
        try {
            if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $rawDetail = $reader.ReadToEnd()
            }
        }
        catch {}

        if (-not [string]::IsNullOrWhiteSpace($rawDetail)) {
            throw "Graph token acquisition failed: $rawDetail"
        }

        throw "Graph token acquisition failed: $($_.Exception.Message)"
    }
}

function Ensure-GraphToken {
    if (-not $global:graphToken -or -not $global:graphTokenExpiry -or (Get-Date) -ge $global:graphTokenExpiry) {
        Write-Log -Message 'Acquiring Microsoft Graph token...' -Level INFO
        Get-GraphToken
    }
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [Parameter()] [object]$Body = $null
    )

    Ensure-GraphToken
    $headers = @{ Authorization = "Bearer $global:graphToken" }

    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = $headers
        ErrorAction = 'Stop'
    }

    if ($null -ne $Body) {
        $params.ContentType = 'application/json'
        $params.Body = ($Body | ConvertTo-Json -Depth 8 -Compress)
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        $rawDetail = ''
        try {
            if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $rawDetail = $reader.ReadToEnd()
            }
        }
        catch {}

        if (-not [string]::IsNullOrWhiteSpace($rawDetail)) {
            throw "Graph call failed [$Method] $Uri :: $rawDetail"
        }

        throw "Graph call failed [$Method] $Uri :: $($_.Exception.Message)"
    }
}
#endregion Graph Authentication

#region User and Manager Resolution
function Get-UserAndManager {
    param([Parameter(Mandatory)] [string]$UserPrincipalName)

    $escapedUpn = $UserPrincipalName.Replace("'", "''")
    $filter = [Uri]::EscapeDataString("userPrincipalName eq '$escapedUpn'")
    $userUri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,displayName,userPrincipalName,mail&`$top=1"
    $userResp = Invoke-Graph -Method GET -Uri $userUri

    if (-not $userResp.value -or $userResp.value.Count -eq 0) {
        throw "User not found in Graph for UPN: $UserPrincipalName"
    }

    $user = $userResp.value[0]

    # Fetch the requested number of manager levels in the hierarchy.
    $managers = [System.Collections.Generic.List[object]]::new()
    $currentUserId = $user.id

    for ($level = 1; $level -le $managerLevel; $level++) {
        $managerUri = "https://graph.microsoft.com/v1.0/users/$currentUserId/manager?`$select=id,displayName,userPrincipalName,mail"
        $currentManager = $null

        try {
            $currentManager = Invoke-Graph -Method GET -Uri $managerUri
        }
        catch {
            if ($level -eq 1) {
                throw "No manager found for $UserPrincipalName in the manager field."
            }
            else {
                Write-Log -Message "No manager found at level $level for user $UserPrincipalName." -Level WARN
                break
            }
        }

        if ($currentManager) {
            $managers.Add($currentManager)
            $currentUserId = $currentManager.id
        }
        else {
            if ($level -eq 1) {
                throw "No manager found for $UserPrincipalName in the manager field."
            }
            break
        }
    }

    return [PSCustomObject]@{
        User     = $user
        Managers = $managers
    }
}

function Get-UserDeletedDate {
    # Returns the deletedDateTime for a soft-deleted user from the Entra recycle bin,
    # or $null if the user is still active or has already been purged past 30 days.
    param([Parameter(Mandatory)] [string]$UserPrincipalName)

    $escapedUpn = $UserPrincipalName.Replace("'", "''")
    $filter = [Uri]::EscapeDataString("userPrincipalName eq '$escapedUpn'")
    $uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?`$filter=$filter&`$select=deletedDateTime&`$top=1"

    try {
        $resp = Invoke-Graph -Method GET -Uri $uri
        if ($resp.value -and $resp.value.Count -gt 0 -and $resp.value[0].deletedDateTime) {
            return [datetime]::Parse($resp.value[0].deletedDateTime)
        }
    }
    catch {
        Write-Log -Message "Could not query soft-deleted items for '$UserPrincipalName': $($_.Exception.Message)" -Level DEBUG
    }

    return $null
}

function Get-UserUnlicensedDateFromAudit {
    # Returns the most recent date a license was removed for the given user by searching
    # directoryAudits for 'Change user license' and 'Remove user from licensed group' events.
    # This covers active users whose account still exists but whose OneDrive license was revoked.
    # Requires AuditLog.Read.All on the app registration.
    # Returns $null if no matching event is found within $auditLogLookbackDays days.
    param(
        [Parameter(Mandatory)] [string]$UserId
    )

    if ($auditLogLookbackDays -le 0) {
        Write-Log -Message 'Audit log lookup skipped ($auditLogLookbackDays = 0).' -Level DEBUG
        return $null
    }

    $cutoffDate = (Get-Date).AddDays(-$auditLogLookbackDays)
    $activityNames = @('Change user license', 'Remove user from licensed group')
    $latestDate = $null

    foreach ($activityName in $activityNames) {
        $encodedFilter = [Uri]::EscapeDataString("activityDisplayName eq '$activityName'")
        $nextUri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=$encodedFilter&`$select=activityDateTime,targetResources&`$top=500"

        do {
            try {
                $resp = Invoke-Graph -Method GET -Uri $nextUri
            }
            catch {
                Write-Log -Message "Audit log query failed for '$activityName': $($_.Exception.Message). Verify AuditLog.Read.All is granted." -Level WARN
                break
            }

            $stopPaging = $false
            foreach ($auditEvent in $resp.value) {
                $eventDate = $null
                try { $eventDate = [datetime]::Parse($auditEvent.activityDateTime) } catch { continue }

                # Audit logs are returned newest-first; stop paging once past the lookback window.
                if ($eventDate -lt $cutoffDate) {
                    $stopPaging = $true
                    break
                }

                foreach ($target in $auditEvent.targetResources) {
                    if ($target.id -eq $UserId) {
                        if (-not $latestDate -or $eventDate -gt $latestDate) {
                            $latestDate = $eventDate
                        }
                        break
                    }
                }
            }

            $nextUri = if ($stopPaging) { $null } else { $resp | Select-Object -ExpandProperty '@odata.nextLink' -ErrorAction SilentlyContinue }
        } while ($nextUri)
    }

    return $latestDate
}
#endregion User and Manager Resolution

#region OneDrive Site Management
function Get-OneDriveUrl {
    param(
        [Parameter(Mandatory)] [string]$UserId,
        [string]$InputOneDriveUrl
    )

    # Always resolve via Graph GET /users/{id}/drive, which returns sharePointIds.siteUrl —
    # the canonical personal-site root URL with no path suffix to strip.
    # A CSV-supplied URL is only used as a last-resort override when Graph returns nothing.
    $driveUri = "https://graph.microsoft.com/v1.0/users/$UserId/drive?`$select=id,webUrl,sharePointIds"
    $drive = Invoke-Graph -Method GET -Uri $driveUri

    Write-Log -Message "Drive response - webUrl: '$($drive.webUrl)'  siteUrl: '$($drive.sharePointIds.siteUrl)'" -Level DEBUG

    # sharePointIds.siteUrl is the clean personal-site root (no /Documents suffix).
    if ($drive.sharePointIds -and -not [string]::IsNullOrWhiteSpace($drive.sharePointIds.siteUrl)) {
        return $drive.sharePointIds.siteUrl.TrimEnd('/')
    }

    # Fallback 1: CSV-provided URL.
    if (-not [string]::IsNullOrWhiteSpace($InputOneDriveUrl)) {
        Write-Log -Message "sharePointIds.siteUrl not returned — using CSV-supplied URL: $InputOneDriveUrl" -Level WARN
        return $InputOneDriveUrl.Trim().TrimEnd('/')
    }

    # Fallback 2: drive.webUrl (may include /Documents — still usable for Connect-PnPOnline).
    if (-not [string]::IsNullOrWhiteSpace($drive.webUrl)) {
        Write-Log -Message "sharePointIds.siteUrl not returned — using drive.webUrl: $($drive.webUrl)" -Level WARN
        return $drive.webUrl.TrimEnd('/')
    }

    throw "Unable to resolve OneDrive site URL for user ID '$UserId'."
}

function Set-ManagerAccessAndReadOnly {
    param(
        [Parameter(Mandatory)] [string]$OneDriveUrl,
        [Parameter(Mandatory)] [string[]]$ManagerUpns
    )

    # Both operations run from the tenant admin context (app-only).
    # Set-PnPTenantSite -Owners grants SCA without the E_ACCESSDENIED that
    # Add-PnPSiteCollectionAdmin produces when connected directly to a personal site.
    $adminUrl = "https://$tenantName-admin.sharepoint.com"
    Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId -WarningAction SilentlyContinue

    foreach ($managerUpn in $ManagerUpns) {
        try {
            Write-Log -Message "Granting SCA to '$managerUpn' on '$OneDriveUrl'..." -Level INFO
            Set-PnPTenantSite -Identity $OneDriveUrl -Owners $managerUpn -ErrorAction Stop | Out-Null
        }
        catch {
            throw "Failed to grant SCA on site '$OneDriveUrl' for manager '$managerUpn': $($_.Exception.Message)"
        }
    }

    if ($lockSite) {
        try {
            Write-Log -Message "Setting LockState=ReadOnly on '$OneDriveUrl'..." -Level INFO
            Set-PnPTenantSite -Identity $OneDriveUrl -LockState ReadOnly -ErrorAction Stop | Out-Null
        }
        catch {
            throw "Failed to set LockState=ReadOnly for site '$OneDriveUrl': $($_.Exception.Message)"
        }
    }
    else {
        Write-Log -Message "Skipping site lock (\$lockSite = \$false)." -Level INFO
    }

    try {
        $site = Get-PnPTenantSite -Identity $OneDriveUrl -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace([string]$site.LockState)) {
            $rawLockState = [string]$site.LockState
            $displayLockState = switch ($rawLockState) {
                'Unlock' { 'Unlocked' }
                default { $rawLockState }
            }

            return $displayLockState
        }

        return 'Unknown'
    }
    catch {
        Write-Log -Message "Unable to query current lock state for '$OneDriveUrl': $($_.Exception.Message)" -Level WARN
        return 'Unknown'
    }
}

function Get-OneDriveSharingLinks {
    param(
        [Parameter(Mandatory)] [string]$OneDriveUrl
    )

    $sharingLinks = [System.Collections.Generic.List[object]]::new()

    try {
        Connect-PnPOnline -Url $OneDriveUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId -WarningAction SilentlyContinue

        $groups = Get-PnPGroup | Where-Object { $_.Title -like 'SharingLinks*' }
        foreach ($group in $groups) {
        $groupName = $group.Title
        $documentId = $null

        if ($groupName -match 'SharingLinks\.([0-9a-fA-F\-]{36})\.') {
            $documentId = $matches[1]
        }

        $members = @()
        try {
            $members = Get-PnPGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
        }
        catch {
            $members = @()
        }

        $memberText = if (@($members).Count -gt 0) {
            ($members | ForEach-Object {
                if ($_.Email) { "$($_.Title) <$($_.Email)>" } else { $_.Title }
            }) -join '; '
        }
        else {
            'No members'
        }

        $linkUrl = ''
        $linkExpiration = ''
        $linkId = ''

        # Extract the linkId baked into the group name as the last GUID segment.
        # Group name format: SharingLinks.{docId}.{type}.{linkId}
        $linkIdFromGroup = $null
        if ($groupName -match '\.([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$') {
            $linkIdFromGroup = $matches[1]
        }

        if ($documentId) {
            try {
                # Resolve the file's server-relative URL so Get-PnPFileSharingLink can
                # locate it reliably. GUID-only identity is inconsistent across PnP versions.
                $fileServerUrl = $null

                try {
                    $fileInfo = Invoke-PnPSPRestMethod -Method Get `
                        -Url "/_api/web/GetFileById('$documentId')?`$select=ServerRelativeUrl" `
                        -ErrorAction Stop
                    if ($fileInfo -and $fileInfo.ServerRelativeUrl) {
                        $fileServerUrl = $fileInfo.ServerRelativeUrl
                    }
                }
                catch { }

                if (-not $fileServerUrl) {
                    try {
                        $itemInfo = Invoke-PnPSPRestMethod -Method Get `
                            -Url "/_api/web/GetListItemByUniqueId('$documentId')?`$select=FileRef" `
                            -ErrorAction Stop
                        if ($itemInfo -and $itemInfo.FileRef) {
                            $fileServerUrl = $itemInfo.FileRef
                        }
                    }
                    catch { }
                }

                $identityArg = if ($fileServerUrl) { $fileServerUrl } else { $documentId }
                Write-Log -Message "Fetching sharing links for '$identityArg'" -Level DEBUG
                $linkCandidates = Get-PnPFileSharingLink -Identity $identityArg -ErrorAction SilentlyContinue

                $match = $linkCandidates | Where-Object {
                    $_.Id -and ($groupName -like "*$($_.Id)*" -or $_.Id -eq $linkIdFromGroup)
                } | Select-Object -First 1

                if ($match) {
                    $linkId = $match.Id
                    if ($match.link -and ($match.link | Select-Object -ExpandProperty WebUrl -ErrorAction SilentlyContinue)) {
                        $linkUrl = $match.link.WebUrl
                    }
                    if ($match.link -and ($match.link | Select-Object -ExpandProperty ExpirationDateTime -ErrorAction SilentlyContinue)) {
                        try {
                            $linkExpiration = ([datetime]::Parse($match.link.ExpirationDateTime)).ToString('yyyy-MM-dd HH:mm:ss')
                        }
                        catch {
                            $linkExpiration = [string]($match.link | Select-Object -ExpandProperty ExpirationDateTime -ErrorAction SilentlyContinue)
                        }
                    }
                    elseif ($match | Select-Object -ExpandProperty ExpirationDateTime -ErrorAction SilentlyContinue) {
                        try {
                            $linkExpiration = ([datetime]::Parse(($match | Select-Object -ExpandProperty ExpirationDateTime -ErrorAction SilentlyContinue))).ToString('yyyy-MM-dd HH:mm:ss')
                        }
                        catch {
                            $linkExpiration = [string]($match | Select-Object -ExpandProperty ExpirationDateTime -ErrorAction SilentlyContinue)
                        }
                    }
                }
                else {
                    Write-Log -Message "No matching sharing link found for group '$groupName'" -Level DEBUG
                }
            }
            catch {
                Write-Log -Message "Unable to resolve sharing-link metadata for group '$groupName' on $OneDriveUrl : $($_.Exception.Message)" -Level WARN
            }
        }

        $sharingLinks.Add([PSCustomObject]@{
                OneDriveUrl           = $OneDriveUrl
                SharingGroupName      = $groupName
                DocumentId            = if ($documentId) { $documentId } else { '' }
                SharingLinkId         = $linkId
                SharingLinkUrl        = $linkUrl
                SharingLinkExpiration = $linkExpiration
                SharingLinkMembers    = $memberText
            })
        }
    }
    catch {
        Write-Log -Message "Error retrieving sharing links for $OneDriveUrl : $($_.Exception.Message)" -Level WARN
    }

    return $sharingLinks
}
#endregion OneDrive Site Management

#region Email Notification
function Send-ManagerNotification {
    param(
        [Parameter(Mandatory)] [pscustomobject]$User,
        [Parameter(Mandatory)] [pscustomobject]$Manager,
        [Parameter(Mandatory)] [string]$OneDriveUrl,
        [Parameter(Mandatory)] [datetime]$DeletedDate,
        [Parameter(Mandatory)] [string]$SiteLockState,
        [Parameter()] [object[]]$SharingLinks = @()
    )

    $expiryDate = $DeletedDate.AddDays($linkExpiryDaysAfterDeletion)
    $daysRemaining = [Math]::Max(0, ($expiryDate.Date - (Get-Date).Date).Days)

    $managerAddress = if ($Manager.mail) { $Manager.mail } else { $Manager.userPrincipalName }

    $rows = if (@($SharingLinks).Count -gt 0) {
        $SharingLinks | ForEach-Object {
            $safeGroup = [System.Web.HttpUtility]::HtmlEncode($_.SharingGroupName)
            $safeMembers = [System.Web.HttpUtility]::HtmlEncode($_.SharingLinkMembers)
            $safeLink = [System.Web.HttpUtility]::HtmlEncode($_.SharingLinkUrl)
            $safeExp = [System.Web.HttpUtility]::HtmlEncode($_.SharingLinkExpiration)
            "<tr><td>$safeGroup</td><td>$safeMembers</td><td>$safeLink</td><td>$safeExp</td></tr>"
        }
    }
    else {
        @('<tr><td colspan="4">No sharing links found.</td></tr>')
    }

    $safeManagerName = [System.Web.HttpUtility]::HtmlEncode($Manager.displayName)
    $safeUserName = [System.Web.HttpUtility]::HtmlEncode($User.displayName)
    $safeUserUpn = [System.Web.HttpUtility]::HtmlEncode($User.userPrincipalName)
    $safeOneDrive = [System.Web.HttpUtility]::HtmlEncode($OneDriveUrl)
    $safeSiteLockState = [System.Web.HttpUtility]::HtmlEncode($SiteLockState)

    $bodyHtml = @"
<html>
<body style='font-family:Segoe UI, Arial, sans-serif;'>
  <p>Hello $safeManagerName,</p>
  <p>You have been granted Site Collection Administrator access for:</p>
  <ul>
    <li>User: $safeUserName ($safeUserUpn)</li>
    <li>OneDrive: $safeOneDrive</li>
        <li>Site Lock State: $safeSiteLockState</li>
  </ul>

  <p><strong>Sharing links currently found:</strong></p>
  <table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;'>
    <thead>
      <tr>
        <th>Sharing Group</th>
        <th>Members</th>
        <th>Sharing Link URL</th>
        <th>Link Expiration (if set)</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>

  <p style='margin-top:16px;'>
    Note: Link lifecycle for this offboarded account should be treated as expiring in
    <strong>$daysRemaining</strong> day(s), on <strong>$($expiryDate.ToString('yyyy-MM-dd'))</strong>.
  </p>
</body>
</html>
"@

    $mailPayload = @{
        message         = @{
            subject      = "OneDrive ownership handoff: $($User.userPrincipalName)"
            body         = @{
                contentType = 'HTML'
                content     = $bodyHtml
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $managerAddress } }
            )
        }
        saveToSentItems = $false
    }

    $sendUri = "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($emailFrom))/sendMail"
    Invoke-Graph -Method POST -Uri $sendUri -Body $mailPayload | Out-Null
}
#endregion Email Notification
#endregion Helper Functions

#region Main

#region Initialization
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' OneDrive Offboarding Handoff Workflow' -ForegroundColor Cyan
Write-Host " Run date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan

if (-not (Test-Path -Path $inputCsvPath)) {
    throw "Input CSV not found: $inputCsvPath"
}

Add-Type -AssemblyName System.Web

Write-Log -Message 'Connecting to SharePoint tenant admin via PnP...' -Level INFO
$adminUrl = "https://$tenantName-admin.sharepoint.com"
Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId -WarningAction SilentlyContinue

$rows = Import-Csv -Path $inputCsvPath
if (-not $rows -or @($rows).Count -eq 0) {
    throw "Input CSV has no rows: $inputCsvPath"
}
#endregion Initialization

$results = [System.Collections.Generic.List[object]]::new()

#region User Processing Loop
foreach ($row in $rows) {
    $upn = [string]$row.UserPrincipalName
    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-Log -Message 'Skipping a row because UserPrincipalName is empty.' -Level WARN
        continue
    }

    Write-Log -Message "Processing $upn" -Level INFO

    try {
        $identity = Get-UserAndManager -UserPrincipalName $upn
        $user = $identity.User
        $managers = $identity.Managers

        # Determine which managers should receive permissions and notifications
        $targetManagers = if ($givePermsToAllManagers) { $managers } else { @($managers[0]) }

        $oneDriveUrl = Get-OneDriveUrl -UserId $user.id -InputOneDriveUrl ([string]($row | Select-Object -ExpandProperty OneDriveUrl -ErrorAction SilentlyContinue))
        $deletedDate = $null

        if (-not [string]::IsNullOrWhiteSpace([string]($row | Select-Object -ExpandProperty DeletedDate -ErrorAction SilentlyContinue))) {
            try {
                $deletedDate = [datetime]::Parse([string]($row | Select-Object -ExpandProperty DeletedDate -ErrorAction SilentlyContinue))
            }
            catch {
                $csvDeletedDate = $row | Select-Object -ExpandProperty DeletedDate -ErrorAction SilentlyContinue
                Write-Log -Message "DeletedDate '$($csvDeletedDate)' is invalid for $upn. Using current date." -Level WARN
            }
        }

        if (-not $deletedDate) {
            # Try the Entra soft-delete recycle bin for deletedDateTime (deleted users).
            Write-Log -Message "No DeletedDate in CSV — checking Entra soft-deleted items..." -Level INFO
            $deletedDate = Get-UserDeletedDate -UserPrincipalName $upn
            if ($deletedDate) {
                Write-Log -Message "Found deletion date from Entra soft-delete: $($deletedDate.ToString('yyyy-MM-dd'))" -Level INFO
            }
        }

        if (-not $deletedDate) {
            # User may still be active but unlicensed — check audit logs for license removal date.
            Write-Log -Message "Checking audit logs for license removal date (last $auditLogLookbackDays days)..." -Level INFO
            $deletedDate = Get-UserUnlicensedDateFromAudit -UserId $user.id
            if ($deletedDate) {
                Write-Log -Message "Found license removal date from audit log: $($deletedDate.ToString('yyyy-MM-dd'))" -Level INFO
            }
            else {
                Write-Log -Message "No unlicensed/deletion date found in CSV, Entra, or audit logs — using today as baseline for 90-day window." -Level WARN
                $deletedDate = Get-Date
            }
        }

        $siteLockState = Set-ManagerAccessAndReadOnly -OneDriveUrl $oneDriveUrl -ManagerUpns @($targetManagers | ForEach-Object { $_.userPrincipalName })
        $sharingLinks = @(Get-OneDriveSharingLinks -OneDriveUrl $oneDriveUrl)

        foreach ($targetManager in $targetManagers) {
            Send-ManagerNotification -User $user -Manager $targetManager -OneDriveUrl $oneDriveUrl -DeletedDate $deletedDate -SiteLockState $siteLockState -SharingLinks $sharingLinks
        }

        # Build manager columns for output CSV - include all levels up to $managerLevel
        $managerOutput = @{}
        for ($i = 1; $i -le $managerLevel; $i++) {
            if ($i -le @($managers).Count) {
                $managerOutput["Manager$i"] = $managers[$i - 1].userPrincipalName
            }
            else {
                $managerOutput["Manager$i"] = ''
            }
        }

        $resultObject = [PSCustomObject]@{
                UserPrincipalName = $user.userPrincipalName
                OneDriveUrl       = $oneDriveUrl
                SiteLockState     = $siteLockState
                SharingLinkCount  = @($sharingLinks).Count
                NotificationSent  = $true
                Notes             = ''
            }

        # Add manager columns in order
        for ($i = 1; $i -le $managerLevel; $i++) {
            $resultObject | Add-Member -NotePropertyName "Manager$i" -NotePropertyValue $managerOutput["Manager$i"]
        }

        $results.Add($resultObject)

        Write-Log -Message "Completed $upn" -Level INFO
    }
    catch {
        $err = $_.Exception.Message
        Write-Log -Message "Failed for $upn : $err" -Level ERROR

        $errorResultObject = [PSCustomObject]@{
                UserPrincipalName = $upn
                OneDriveUrl       = ''
                SiteLockState     = ''
                SharingLinkCount  = 0
                NotificationSent  = $false
                Notes             = $err
            }

        # Add empty manager columns for consistency
        for ($i = 1; $i -le $managerLevel; $i++) {
            $errorResultObject | Add-Member -NotePropertyName "Manager$i" -NotePropertyValue ''
        }

        $results.Add($errorResultObject)
    }
}
#endregion User Processing Loop

#region Output
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8
Write-Host ''
Write-Host "Results written to: $outputCsvPath" -ForegroundColor Green
Write-Host "Processed users   : $($results.Count)" -ForegroundColor Green
#endregion Output
#endregion Main
