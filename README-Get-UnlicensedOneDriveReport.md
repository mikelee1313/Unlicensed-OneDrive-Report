# Get-UnlicensedOneDriveReport.ps1

PowerShell script to identify and report unlicensed OneDrive accounts across a Microsoft 365 tenant (including multi-geo), with archive-risk, cost, and notification workflows.

The script uses Microsoft Graph app permissions and SharePoint Admin ExportToCSV download to enrich archived account state.

## Summary

This report is designed for Microsoft unlicensed OneDrive enforcement tracking, including:

- Day 60 read-only
- Day 93 archived
- Day 275 undiscoverable eDiscovery risk (when PAYG is not enabled)
- Day 365 deletion-risk timeline with MC1381110 enforcement floor handling
- Reactivated-but-still-unlicensed accounts and their 30-day rearchive window

## Account Populations

The script discovers and merges:

1. Active Entra users without enabled OneDrive/SharePoint plan.
2. Soft-deleted Entra users in recycle bin (up to 30 days).
3. Archived/unlicensed OneDrive rows from SharePoint Admin downloaded report.

## Timeline and Risk Model

| Phase | Day | Behavior |
|---|---:|---|
| Read-only | 60 | Site becomes read-only |
| Archive | 93 | Site is archived |
| Undiscoverable risk | 275 | Archived site can become undiscoverable via eDiscovery when PAYG is not enabled |
| Deletion risk | 365 | Risk tracking starts from unlicensed+365, with MC1381110 enforcement floor applied |

### MC1381110 Enforcement Floor

Deletion risk is calculated with an effective deletion date:

`max(UnlicensedDate + ArchiveDeletionThresholdDays, DeletionEnforcementStartDate)`

Current script default:

- `$ArchiveDeletionThresholdDays = 365`
- `$DeletionEnforcementStartDate = 2027-07-01`

This prevents pre-enforcement deletion risk dates from appearing earlier than the policy floor.

### Source of Truth for Archive State

When SharePoint downloaded report data is available, its archive state is treated as current truth.

- If timeline math says "should be archived" but downloaded report archive state is not archived (`None`, `reactivating`, `unknownFutureValue`), the account is flagged as reactivated.
- Reactivated and still-unlicensed accounts are tracked with:
  - `RearchiveDate`
  - `DaysUntilRearchive`
  - `ArchiveStatus = Reactivated`

## Key Capabilities

- Multi-geo support with single app registration/token.
- Bulk Graph user and drive discovery.
- Optional audit-log enrichment for active-user unlicensed dates.
- SharePoint Admin ExportToCSV ingestion and backfill into final dataset.
- Archive-state reconciliation using downloaded report fields.
- Cost projection:
  - monthly storage (`$0.05/GB/month`)
  - reactivation (`$0.60/GB` one-time)
- HTML alert emails for approaching milestones.
- Retry/throttle handling for Graph/SPO requests.

## Authentication

Certificate authentication is used.

Ensure certificate exists in one of:

- `Cert:\LocalMachine\My\<thumbprint>`
- `Cert:\CurrentUser\My\<thumbprint>`

## Required Permissions

Graph application permissions:

- `User.Read.All`
- `Directory.Read.All`
- `Files.Read.All`
- `AuditLog.Read.All` (if audit enrichment enabled)
- `Sites.Read.All` (optional for Graph sites path)
- `Mail.Send` (optional for notifications)

SharePoint application permission:

- `Sites.FullControl.All` (required for ExportToCSV download path)

## Important Configuration

Update in script `CONFIGURATION SECTION`:

- Tenant/auth
  - `$tenantId`
  - `$clientId`
  - `$Thumbprint`
  - `$CertStore`
- Timeline/risk
  - `$ReadOnlyThresholdDays`
  - `$ArchiveThresholdDays`
  - `$UndiscoverableThresholdDays`
  - `$RearchiveThresholdDays`
  - `$ArchiveDeletionThresholdDays`
  - `$DeletionEnforcementStartDate`
  - `$PayGEnabledForUnlicensedOneDrive`
- Data collection
  - `$GetCurrentlyArchived`
  - `$ArchivedCollectionMode`
  - `$SPOAdminUrls`
  - `$includeLicenseRemovalDates`
  - `$IncludeDownloadedRowsInMainReport`
- Notifications
  - `$SendEmailNotifications`
  - `$EmailFrom`
  - `$EmailTo`
  - `$DaysToNotifyBeforeReadOnly`
  - `$DaysToNotifyBeforeArchive`
  - `$DaysToNotifyBeforeRearchive`
  - `$DaysToNotifyBeforeDeletion`

## Notifications

When enabled, script can send 4 alert types:

1. Read-only alert
2. Archive alert
3. Rearchive alert (reactivated + still unlicensed)
4. Deletion-risk alert

Each alert uses its own threshold window and date/remaining-day columns.

## Output CSV Schema

Core columns include:

- `UserSource`
- `DisplayName`
- `UserPrincipalName`
- `AccountEnabled`
- `UnlicensedDueTo`
- `UnlicensedDate`
- `DaysSinceUnlicensed`
- `ReadOnlyDate`
- `ArchiveDate`
- `ArchiveStatus`
- `UndiscoverableDate`
- `RearchiveDate`
- `DaysUntilReadOnly`
- `DaysUntilArchive`
- `DaysUntilUndiscoverable`
- `DaysUntilRearchive`
- `DeletionBlockedBy`
- `DaysUntilDeletion`
- `UrgencyStatus`
- `StorageUsedGB`
- `StorageTotalGB`
- `ProjMonthlyStorageCost`
- `ProjReactivationCost`
- `DriveUrl`
- `DriveLastModified`
- `Notes`

## How to Run

```powershell
.\Get-UnlicensedOneDriveReport.ps1
```

Output file:

- `UnlicensedOneDrive_<timestamp>.csv` in `$OutputFolder` (default `%TEMP%`).

## Processing Flow

1. Acquire Graph token.
2. Enumerate active unlicensed users.
3. Enumerate soft-deleted users.
4. Download archived rows from SharePoint Admin export.
5. Optionally enrich active users from audit logs.
6. Resolve drive metadata.
7. Backfill/merge downloaded report rows into final candidate list.
8. Calculate milestones, risks, archive-status reconciliation, and costs.
9. Export final CSV.
10. Optionally send milestone/risk alert emails.

## Troubleshooting

### Archive state looks inconsistent

- Check downloaded SPO report `ARCHIVE_STATUS` values.
- If report says non-archived but day math is past 93, account is intentionally flagged as `Reactivated`.

### Deletion risk appears later than day 365

- This is expected when `DeletionEnforcementStartDate` is later than `UnlicensedDate + 365`.

### Missing unlicensed dates for active users

- Confirm `AuditLog.Read.All` and increase `$AuditLogLookbackDays` (max 180 in current implementation).

### Email failures

- Confirm `Mail.Send` app permission + admin consent.
- Confirm `$EmailFrom` is a licensed Exchange mailbox.

## Known Limitations

- PAYG status is manual (`$PayGEnabledForUnlicensedOneDrive`).
- Very old purged users may be unrecoverable from APIs.
- UPN reconstruction from OneDrive URL is best-effort.
- `SPODownload` path does not automatically fallback to GraphSites when no rows are returned.

## Disclaimer

Validate in a non-production tenant first. Use report outputs as decision support, and confirm policy/compliance actions with tenant admins.

## Authors

- Mike Lee
- Mariel Williams
