# Get-UnlicensedOneDriveReport.ps1

Automates OneDrive offboarding handoff for unlicensed or deleted users by granting manager access, optionally locking the site, inventorying sharing links, and sending a manager notification email.

## Overview

This script processes users from an input CSV and performs the following workflow for each user:

1. Resolve the target user in Microsoft Graph.
2. Resolve the manager from the Entra manager attribute.
3. Resolve the OneDrive personal site URL.
4. Grant the manager Site Collection Admin (SCA) access.
5. Optionally set the OneDrive site lock state to ReadOnly.
6. Enumerate SharingLinks groups and map them to sharing-link metadata.
7. Send an HTML notification email to the manager with site and sharing-link details.
8. Export run results to CSV.

## Key Capabilities

- App-only authentication to Microsoft Graph using certificate-based client credentials.
- App-only SharePoint admin operations with PnP.PowerShell.
- Fallback logic to calculate offboarding baseline date:
  - CSV DeletedDate
  - Entra soft-delete deletedDateTime
  - Directory audit log license-removal event date
  - Current date fallback
- Sharing link inventory built from SharePoint SharingLinks groups plus link lookup.
- Manager notification email sent through Graph sendMail.
- Run summary export for success/failure tracking.

## Requirements

### PowerShell and Modules

- PowerShell 7.x recommended.
- PnP.PowerShell module installed.

Install module:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### App Registration Permissions

Grant admin consent for the following Application permissions.

Microsoft Graph:

| Permission | Purpose |
|---|---|
| User.Read.All | Resolve user details |
| Directory.Read.All | Query deleted users and manager references |
| Files.Read.All | Resolve OneDrive drive metadata and site URL |
| AuditLog.Read.All | Read directory audit logs for license-removal events |
| Mail.Send | Send manager notification email |

SharePoint:

| Permission | Purpose |
|---|---|
| Sites.FullControl.All | Grant SCA, set lock state, and inspect sharing metadata |

### Certificate Requirements

- Certificate private key must be available on the host running the script.
- Current script expects certificate in LocalMachine\My by thumbprint.
- App registration must trust the certificate public key.

## Input CSV

### Required Column

- UserPrincipalName

### Optional Columns

- DeletedDate (any parseable date, recommended yyyy-MM-dd)
- OneDriveUrl (used only as fallback if Graph does not return siteUrl)

### Example

```csv
UserPrincipalName,DeletedDate,OneDriveUrl
user1@contoso.com,2026-05-01,
user2@contoso.com,,
user3@contoso.com,,https://contoso-my.sharepoint.com/personal/user3_contoso_com
```

## Configuration

Edit values in the Configuration region at the top of the script.

| Variable | Description |
|---|---|
| $tenantName | Tenant short name used for SharePoint admin URL |
| $tenantId | Entra tenant ID |
| $clientId | App registration (client) ID |
| $thumbprint | Certificate thumbprint used for Graph and PnP auth |
| $emailFrom | Mailbox identity used with Graph sendMail |
| $linkExpiryDaysAfterDeletion | Expiry guidance window shown in email |
| $lockSite | If true, sets LockState to ReadOnly after SCA grant |
| $auditLogLookbackDays | Days to search audit logs for license-removal events |
| $inputCsvPath | Path to source CSV |
| $outputFolder | Destination folder for results CSV |
| $debug | Enables DEBUG logging output |

## Execution

Run from PowerShell:

```powershell
.\Get-UnlicensedOneDriveReport.ps1
```

The script writes progress logs to console and exports a timestamped CSV result file:

```text
OneDrive_Manager_Handoff_yyyyMMdd_HHmmss.csv
```

## Processing Logic

For each user row:

1. Validate UserPrincipalName.
2. Resolve user and manager in Graph.
3. Resolve OneDrive URL from Graph /users/{id}/drive.
4. Determine offboarding baseline date in this order:
   - CSV DeletedDate
   - Entra soft-delete deletedDateTime
   - Audit logs (Change user license or Remove user from licensed group)
   - Today
5. Grant manager SCA on OneDrive site.
6. If enabled, set site LockState to ReadOnly.
7. Enumerate sharing links via SharingLinks groups and Get-PnPFileSharingLink.
8. Send manager notification with site details and sharing-link table.
9. Write success/failure to output CSV.

## Output CSV

| Column | Description |
|---|---|
| UserPrincipalName | Processed user UPN |
| Manager | Manager UPN |
| OneDriveUrl | Resolved OneDrive site URL |
| SiteLockState | Current lock state (for example, Unlocked or ReadOnly) |
| SharingLinkCount | Number of sharing-link group records found |
| NotificationSent | True when email call succeeded |
| Notes | Error message when row fails |

## Email Notification Content

Each manager email includes:

- Offboarded user name and UPN.
- OneDrive URL.
- Current site lock state.
- Sharing link inventory table:
  - Sharing group name
  - Group members
  - Sharing link URL
  - Link expiration (if present)
- Calculated days remaining based on baseline date + link expiry window.

## Troubleshooting

### Certificate not found

Error resembles:

```text
Certificate with thumbprint '<thumbprint>' was not found in LocalMachine\My.
```

Actions:

- Verify certificate is installed in LocalMachine\My.
- Verify thumbprint in script exactly matches certificate thumbprint.
- Ensure account running script can read certificate private key.

### Graph token acquisition failed

Actions:

- Confirm app registration IDs and tenant ID.
- Verify certificate credential is configured on app registration.
- Confirm required Graph Application permissions are granted and consented.

### No manager found

Cause:

- Manager attribute is not set on the user object.

Action:

- Populate manager in Entra for the user, then rerun.

### Audit log lookup returns nothing

Actions:

- Confirm AuditLog.Read.All is granted.
- Increase $auditLogLookbackDays.
- Validate that relevant license-change events exist for the user.

### SCA grant or lock state update fails

Actions:

- Confirm SharePoint Sites.FullControl.All Application permission is granted.
- Verify personal site exists and URL resolves correctly.
- Verify admin URL format: https://<tenant>-admin.sharepoint.com.

### sendMail fails

Actions:

- Confirm Mail.Send Application permission with admin consent.
- Ensure $emailFrom is a valid mailbox identity.
- Verify Exchange Online licensing and mailbox availability for sender account.

## Security and Operational Guidance

- Use dedicated app registrations and certificates for automation.
- Rotate certificates on a defined schedule.
- Limit script host access and protect private keys.
- Store input CSV in restricted locations.
- Review and archive output CSV because it may contain user and sharing data.
- Pilot in a test tenant or with a small user batch before full runs.

## Known Limitations

- Sharing link discovery depends on SharePoint SharingLinks group patterns and available metadata.
- Audit logs may not include historical events beyond retention or lookback window.
- If a deleted user has been permanently purged, soft-delete date cannot be retrieved from directory deletedItems.

## Disclaimer

This script is provided as-is without warranty of any kind. Validate behavior in a non-production environment before broad deployment.

## License

Use your organization standard license for publication (for example, MIT).
